"""
Python collection runner.
Orchestrates all section collectors in parallel, assembles a SchemaVersion 2 snapshot,
and saves it to <export_path>/Support/<tenant>-Snap.json.
"""

from __future__ import annotations

import logging
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from rich.console import Console

_con = Console()

from ..snapshot import create as _snap_create, save as _snap_save
from ..utils.paths import resolve_snapshot_output_context
from .auth import GraphAuthProvider
from .graph_client import GraphClient
from .collectors import tenant as col_tenant
from .collectors import identity as col_identity
from .collectors import exchange as col_exchange
from .collectors import exchange_exo as col_exchange_exo
from .collectors import collaboration as col_collab
from .collectors import security as col_security
from .collectors import governance as col_governance

log = logging.getLogger(__name__)


def collect(
    auth: GraphAuthProvider,
    export_path: str | Path,
    tenant_name: str | None = None,
) -> Path:
    """
    Run all collectors, assemble a snapshot, save to disk, and return the snapshot path.
    Collectors run in parallel sections (max 4 workers).
    """
    export_path = Path(export_path)
    client      = GraphClient(auth)

    t_start = time.monotonic()
    log.info("Python collection starting...")

    # Tenant section must run first — provides tenant name and domain list
    _con.print("  [dim]Connecting to Microsoft Graph...[/dim]")
    t0 = time.monotonic()
    tenant_data = col_tenant.collect(client)

    # Resolve tenant name for output path
    if not tenant_name:
        tenant_name = (
            tenant_data.get("TenantInfo", {}).get("DisplayName")
            or "Tenant"
        )
    tenant_name = _sanitize_name(tenant_name)
    _con.print(f"  [green]v[/green] Tenant       [dim]{tenant_name}[/dim]  [dim]({time.monotonic()-t0:.1f}s)[/dim]")

    section_data: dict[str, Any] = {"Tenant": tenant_data}
    errors: list[str] = []
    warnings: list[str] = []

    # Report anonymization check — must happen before parallel so the warning is visible early
    if tenant_data.get("ReportSettings", {}).get("displayConcealedNames"):
        _con.print()
        _con.print("  [yellow bold]WARNING:[/yellow bold] [yellow]Report anonymization is ON for this tenant.[/yellow]")
        _con.print("  [dim]  Usage reports (Teams, email, SharePoint, OneDrive) will contain GUIDs[/dim]")
        _con.print("  [dim]  instead of user names. User-level report analysis will not be readable.[/dim]")
        _con.print("  [dim]  Disable in: M365 admin center > Settings > Org settings > Reports.[/dim]")
        warnings.append("Report anonymization (isDisplayConcealedNames) is enabled — usage reports contain GUIDs, not user names")

    _con.print()
    _con.print("  [dim]Collecting 5 sections in parallel...[/dim]")

    # All remaining sections run concurrently
    _SECTIONS = {
        "Identity":      col_identity.collect,
        "Exchange":      col_exchange.collect,
        "Collaboration": col_collab.collect,
        "Security":      col_security.collect,
        "Governance":    col_governance.collect,
    }

    with ThreadPoolExecutor(max_workers=4) as pool:
        submit_times: dict = {}
        futures: dict = {}
        for section, fn in _SECTIONS.items():
            f = pool.submit(fn, client)
            futures[f] = section
            submit_times[f] = time.monotonic()

        for future in as_completed(futures):
            section = futures[future]
            elapsed_section = time.monotonic() - submit_times[future]
            try:
                section_data[section] = future.result()
                log.info("Section %s complete", section)
                _con.print(f"  [green]v[/green] {section:<14} [dim]({elapsed_section:.1f}s)[/dim]")
            except Exception as exc:
                log.error("Section %s failed: %s", section, exc, exc_info=True)
                errors.append(f"{section}: {exc}")
                section_data[section] = {}
                _con.print(f"  [red]x[/red] {section:<14} [red]failed[/red]  [dim]({elapsed_section:.1f}s)[/dim]")

    # DNS-based email auth records (no Graph needed — uses domain list from Tenant)
    _con.print("  [dim]Resolving email authentication DNS records...[/dim]")
    domain_names = list((tenant_data.get("Domains") or {}).keys())
    try:
        dns_records = col_governance.collect_domain_auth(domain_names)
        section_data["Governance"]["DomainAuthenticationRecords"] = dns_records
        _con.print(f"  [green]v[/green] DNS auth      [dim]{len(dns_records)} domain(s)[/dim]")
    except Exception as exc:
        log.warning("DNS auth collection failed: %s", exc)
        section_data["Governance"]["DomainAuthenticationRecords"] = {}
        _con.print("  [yellow]![/yellow] DNS auth      [dim]skipped[/dim]")

    # EXO PS subprocess (optional: ExchangeOnlineManagement module + Exchange.ManageAsApp)
    # Collects AllMailboxes, MailFlowRules, MailFlowConnectors, SpamFilteringConfig,
    # RemoteDomains, PublicFolderDetails, InactiveMailboxDetails, ForwardingPolicySummary.
    _con.print("  [dim]Exchange Online PS enrichment (optional -- EXO module required)...[/dim]")
    org_domain = next(
        (d for d in (tenant_data.get("Domains") or {}) if d.lower().endswith(".onmicrosoft.com")),
        (tenant_data.get("TenantInfo") or {}).get("DefaultDomain", ""),
    )
    t_exa = time.monotonic()
    try:
        exa_data = col_exchange_exo.collect(auth, org_domain)
        if exa_data:
            for k, v in exa_data.items():
                section_data["Exchange"][k] = v
            all_mbx   = exa_data.get("AllMailboxes") or {}
            mbx_vals  = list(all_mbx.values()) if isinstance(all_mbx, dict) else (all_mbx or [])
            fwd_count = sum(
                1 for m in mbx_vals if isinstance(m, dict) and
                (m.get("ForwardingSmtpAddress") or m.get("ForwardingAddress"))
            )
            _con.print(
                f"  [green]v[/green] EXO PS        "
                f"[dim]{len(exa_data)} dataset(s), {len(mbx_vals)} mailbox(es), "
                f"{fwd_count} with forwarding[/dim]  "
                f"[dim]({time.monotonic()-t_exa:.1f}s)[/dim]"
            )
        else:
            _con.print("  [dim]-[/dim] EXO PS        [dim]skipped[/dim]")
    except Exception as exc:
        log.warning("EXO PS subprocess failed: %s", exc)
        _con.print("  [yellow]![/yellow] EXO PS        [dim]failed[/dim]")

    # Per-user Exchange enrichment (optional: MailboxSettings.Read)
    # Runs after parallel phase so Identity user IDs are available
    _con.print("  [dim]Per-user mailbox and inbox rule scan (MailboxSettings.Read)...[/dim]")
    user_ids_list = list((section_data.get("Identity", {}).get("Users", {}) or {}).keys())
    accepted_domains: set[str] = {d.lower() for d in (tenant_data.get("Domains") or {}).keys()}
    if user_ids_list:
        try:
            mbx_types = col_exchange.collect_mailbox_types(client, user_ids_list)
            if mbx_types:
                section_data["Exchange"]["MailboxPurposes"]    = mbx_types.get("MailboxPurposes", {})
                section_data["Exchange"]["SharedMailboxes"]    = mbx_types.get("SharedMailboxes", [])
                section_data["Exchange"]["EquipmentMailboxes"] = mbx_types.get("EquipmentMailboxes", [])
                shared_count = len(mbx_types.get("SharedMailboxes", []))
                _con.print(f"  [green]v[/green] Mailbox types [dim]{shared_count} shared mailbox(es) found[/dim]")
            else:
                _con.print("  [dim]-[/dim] Mailbox types [dim]skipped (no MailboxSettings.Read)[/dim]")
        except Exception as exc:
            log.warning("Mailbox type detection failed: %s", exc)
            _con.print("  [yellow]![/yellow] Mailbox types [dim]failed[/dim]")

        try:
            fwd_rules = col_exchange.collect_inbox_forwarding_rules(
                client, user_ids_list, accepted_domains
            )
            if fwd_rules is not None:
                section_data["Exchange"]["InboxRulesExternalForwarding"] = fwd_rules
                ext_count = sum(1 for r in fwd_rules if r.get("ExternalAddresses"))
                _con.print(f"  [green]v[/green] Inbox rules   [dim]{ext_count} rule(s) with external forwarding[/dim]")
            else:
                _con.print("  [dim]-[/dim] Inbox rules   [dim]skipped (no MailboxSettings.Read)[/dim]")
        except Exception as exc:
            log.warning("Inbox rule scan failed: %s", exc)
            _con.print("  [yellow]![/yellow] Inbox rules   [dim]failed[/dim]")
    else:
        _con.print("  [dim]-[/dim] Per-user Exchange [dim]no users found — skipped[/dim]")

    elapsed = time.monotonic() - t_start
    _con.print()
    _con.print(f"  [dim]Collection complete --[/dim] {elapsed:.1f}s")
    _con.print()
    log.info("Collection complete in %.1fs. Errors: %d", elapsed, len(errors))

    # Build derived section
    derived = _build_derived(section_data)

    # Assemble snapshot
    generated_at = datetime.now(timezone.utc).isoformat()
    metadata = {
        "GeneratedAt":        generated_at,
        "TenantDisplayName":  tenant_name,
        "CollectionMethod":   "Python/GraphAPI",
        "CollectionDuration": f"{elapsed:.1f}s",
        "OutputProfileLabel": "SolutionsEngineer",
    }
    collection_plan = {s: {"Status": "Collected" if not any(s in e for e in errors) else "Failed"}
                       for s in section_data}

    diagnostics = {
        "WarningCount":  len(warnings),
        "ErrorCount":    len(errors),
        "WarningSummary": warnings,
        "ErrorSummary":   errors,
        "SourceCoverage": {},
        "CollectorStats": {"DurationSeconds": round(elapsed, 1)},
    }

    snapshot = _snap_create(
        metadata=metadata,
        collection_plan=collection_plan,
        data=section_data,
        derived=derived,
        diagnostics=diagnostics,
    )

    # Resolve output path using the same logic as the reporting pipeline
    ctx       = resolve_snapshot_output_context(snapshot, export_path)
    stem      = ctx["FileStem"]
    support   = export_path / "Support"
    support.mkdir(parents=True, exist_ok=True)
    snap_path = support / f"{stem}-Snap.json"

    _snap_save(snapshot, snap_path)
    _con.print(f"  [dim]Snapshot :[/dim] {snap_path}")
    log.info("Snapshot saved: %s", snap_path)
    return snap_path


# ---------------------------------------------------------------------------
# Derived section assembly
# ---------------------------------------------------------------------------

def _build_derived(section_data: dict) -> dict:
    """Build the Derived section from raw collected data."""
    derived: dict[str, Any] = {}

    identity = section_data.get("Identity", {})

    # MfaRegistrationSummary — used directly by reporting
    mfa_enrollment = identity.get("MfaEnrollmentSummary", {})
    mfa_details    = identity.get("MfaRegistrationDetails", [])
    if mfa_enrollment:
        method_posture = identity.get("MfaMethodPostureSummary", {})
        derived["MfaRegistrationSummary"] = {
            **mfa_enrollment,
            **method_posture,
            "WeakMethodCounts":             {},
            "StrongMethodCounts":           {},
            "PhishingResistantMethodCounts": {},
            "DefaultMethodCounts":          {},
        }

    # AuthenticationConfigSummary
    ca_summary = identity.get("ConditionalAccessPolicySummary", {}).get("Summary", {})
    auth_config = identity.get("AuthenticationConfig", {})
    derived["AuthenticationConfigSummary"] = {
        "Summary": {
            "MFAEnabled":                 ca_summary.get("HasRiskBasedCoverage") or bool(ca_summary.get("EnabledPolicies")),
            "MFAConditionalAccessPolicies": ca_summary.get("PoliciesBlockingLegacyAuth", 0),
            "SelfServicePasswordReset":   "Unknown",
            "DefaultUserCanCreateApps":   "Unknown",
            "GuestUserRoleLabel":         "Unknown",
        }
    }

    # TenantInfoSummary
    tenant = section_data.get("Tenant", {})
    tenant_info = tenant.get("TenantInfo", {})
    derived["TenantInfoSummary"] = {
        "TenantId":     tenant_info.get("TenantId", ""),
        "DisplayName":  tenant_info.get("DisplayName", ""),
        "DefaultDomain": tenant_info.get("DefaultDomain", ""),
        "IsHybrid":     tenant.get("HybridConfiguration", {}).get("IsHybrid", False),
    }

    # FederationSummary
    fed_config = tenant.get("FederationConfiguration", {})
    derived["FederationSummary"] = {
        "FederatedDomainCount": len(fed_config),
        "FederatedDomains":     list(fed_config.keys()),
    }

    # UnmanagedObjects — built from collaboration cleanup candidates
    collab = section_data.get("Collaboration", {})
    derived["UnmanagedObjects"] = collab.get("TeamsGroupsCleanupCandidates", {})

    # OwnershipGovernanceSummary
    cleanup = collab.get("TeamsGroupsCleanupCandidates", {})
    unified = collab.get("UnifiedGroups", {})
    ownerless = [v for v in cleanup.values() if "Ownerless" in str(v.get("RiskSignal", ""))]
    derived["OwnershipGovernanceSummary"] = {
        "Summary": {
            "TotalObjectsReviewed":   len(unified),
            "OwnerlessObjectCount":   len(ownerless),
            "CleanupCandidateCount":  len(cleanup),
        }
    }

    # SpamFilteringSummary (stub — EXO-only)
    derived["SpamFilteringSummary"] = {}

    # MigrationReadiness (stub)
    derived["MigrationReadiness"] = {}

    # EmployeeExperienceInsightsSummary (stub)
    derived["EmployeeExperienceInsightsSummary"] = {}

    # OneDriveOwnerMismatches (stub)
    derived["OneDriveOwnerMismatches"] = []

    return derived


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sanitize_name(name: str) -> str:
    """Strip filesystem-unsafe characters from a tenant display name."""
    unsafe = r'\/:*?"<>|'
    for ch in unsafe:
        name = name.replace(ch, "")
    return name.strip() or "Tenant"
