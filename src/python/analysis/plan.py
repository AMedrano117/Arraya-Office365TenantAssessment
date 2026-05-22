"""
Plan generation — Python equivalent of New-M365TenantImprovementPlan.ps1.

Takes a snapshot, enriches BestPracticeFindings with Phase, Workstream, and
narrative fields, then generates hybrid findings from Summary datasets.
Builds a Plan.json that powers the EngPack.md and Word reports.
"""

from __future__ import annotations

import json
import logging
import re
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from ..snapshot import get_metadata
from ..utils.converters import ensure_list

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Severity / Phase mappings (mirrors PS Convert-ToPriorityBand)
# ---------------------------------------------------------------------------

_SEVERITY_DISPLAY: dict[str, str] = {
    "Risk": "High",
    "Warning": "Medium",
    "Info": "Low",
}

_SEV_ORDER: dict[str, int] = {"High": 0, "Medium": 1, "Low": 2}
_PHASE_ORDER: dict[str, int] = {"Immediate": 0, "Near Term": 1, "Planned": 2, "Monitor": 3}

# (snap_severity, priority_int) -> RoadmapPhase
_PHASE_MAP: dict[tuple[str, int], str] = {
    ("Risk",    1): "Immediate",
    ("Risk",    2): "Near Term",
    ("Risk",    3): "Planned",
    ("Warning", 1): "Near Term",
    ("Warning", 2): "Planned",
    ("Warning", 3): "Planned",
    ("Info",    1): "Monitor",
    ("Info",    2): "Monitor",
    ("Info",    3): "Monitor",
}

# WorkstreamSummary severity label (aggregated highest)
_DISPLAY_SEV_LABEL: dict[str, str] = {
    "High": "Critical",
    "Medium": "Medium",
    "Low": "Info",
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate(
    snapshot: dict,
    support_dir: Path,
    tenant_name: str = "",
) -> dict:
    """Enrich findings, generate hybrid findings, build Plan.json, and return the plan dict."""
    support_dir = Path(support_dir)
    support_dir.mkdir(parents=True, exist_ok=True)

    if not tenant_name:
        meta = get_metadata(snapshot)
        tenant_obj = meta.get("Tenant") or {}
        tenant_name = (
            (tenant_obj.get("DisplayName") if isinstance(tenant_obj, dict) else None)
            or meta.get("TenantDisplayName")
            or meta.get("TenantDomain")
            or "Tenant"
        )

    raw = _extract_raw_findings(snapshot)
    findings = [_enrich(f) for f in raw]
    bpf_categories = {f.get("Category", "") for f in findings}
    findings += _generate_hybrid_findings(snapshot, bpf_categories)
    findings.sort(key=lambda f: (
        _SEV_ORDER.get(f.get("Severity", "Low"), 99),
        _PHASE_ORDER.get(f.get("RoadmapPhase", "Monitor"), 99),
    ))

    ws_summaries = _build_workstream_summaries(findings)
    consultative = _build_consultative_summaries(findings)
    roadmap_actions = _build_roadmap_actions(findings)

    plan: dict[str, Any] = {
        "GeneratedAt": datetime.now(timezone.utc).isoformat(),
        "TenantName": tenant_name,
        "FindingCount": len(findings),
        "WorkstreamSummaries": ws_summaries,
        "RoadmapActions": roadmap_actions,
        "Findings": findings,
        "ConsultativeSummaries": consultative,
    }

    plan_path = support_dir / f"{tenant_name}-Plan.json"
    plan_path.write_text(json.dumps(plan, indent=2, default=str), encoding="utf-8")
    log.info("Plan.json written: %s (%d findings)", plan_path, len(findings))
    return plan


# ---------------------------------------------------------------------------
# BestPracticeFindings extraction + enrichment
# ---------------------------------------------------------------------------

def _extract_raw_findings(snapshot: dict) -> list[dict]:
    derived = snapshot.get("Derived") or {}
    src = derived.get("BestPracticeFindings") or derived.get("Findings") or {}
    if isinstance(src, dict):
        return [{"_key": k, **v} for k, v in src.items() if isinstance(v, dict)]
    return [f for f in ensure_list(src) if isinstance(f, dict)]


def _enrich(f: dict) -> dict:
    key       = f.get("_key", "")
    area      = str(f.get("Area", ""))
    category  = str(f.get("Category", ""))
    severity  = str(f.get("Severity", "Info"))
    priority  = int(f.get("Priority", 2) or 2)
    message   = str(f.get("Message") or f.get("Description") or "")
    rec       = str(f.get("RecommendedAction") or f.get("Remediation") or "")
    worksheet = str(f.get("RelatedWorksheet") or "")
    section   = str(f.get("RelatedSection") or "")
    src_type  = str(f.get("SourceType") or "")

    display_sev = _SEVERITY_DISPLAY.get(severity, severity)
    phase       = _PHASE_MAP.get((severity, priority), "Monitor")
    workstream  = _get_workstream(area, category)

    evidence_parts = []
    if worksheet:
        evidence_parts.append(f"Worksheet: {worksheet}")
    if section:
        evidence_parts.append(f"Section: {section}")
    if src_type:
        evidence_parts.append(f"Source: {src_type}")
    evidence_location = "; ".join(evidence_parts)

    return {
        "RuleId":               key,
        "Area":                 area,
        "Category":             category,
        "Severity":             display_sev,
        "RoadmapPhase":         phase,
        "Workstream":           workstream,
        "Finding":              message,
        "Recommendation":       rec,
        "WhyFlagged":           f"Flagged because the assessment observed: {message}" if message else "",
        "TechnicalRemediation": rec,
        "CurrentEvidence":      message,
        "EvidenceLocation":     evidence_location,
    }


# ---------------------------------------------------------------------------
# Hybrid findings — generated from Summary datasets
# ---------------------------------------------------------------------------

def _generate_hybrid_findings(snapshot: dict, bpf_categories: set[str] | None = None) -> list[dict]:
    """Generate findings from Summary and raw datasets not covered by BestPracticeFindings."""
    bpf_categories = bpf_categories or set()
    data    = snapshot.get("Data") or {}
    derived = snapshot.get("Derived") or {}
    diag    = snapshot.get("Diagnostics") or {}
    findings: list[dict] = []

    def _rows(section: str, table: str) -> list[dict]:
        tbl = (data.get(section) or {}).get(table) or {}
        if isinstance(tbl, list):
            return [r for r in tbl if isinstance(r, dict)]
        if isinstance(tbl, dict):
            return [v for v in tbl.values() if isinstance(v, dict)]
        return []

    def _first(section: str, table: str) -> dict:
        rows = _rows(section, table)
        return rows[0] if rows else {}

    def _derived_first(key: str) -> dict:
        val = derived.get(key) or {}
        if isinstance(val, list):
            return val[0] if val else {}
        if isinstance(val, dict):
            vals = list(val.values())
            first = vals[0] if vals else {}
            return first if isinstance(first, dict) else {}
        return {}

    def _hf(
        rule_id: str,
        area: str,
        category: str,
        severity: str,
        phase: str,
        workstream: str,
        finding: str,
        remediation: str,
        worksheet: str,
        section: str = "",
        source: str = "",
    ) -> dict:
        ev_parts = [f"Worksheet: {worksheet}"]
        if section:
            ev_parts.append(f"Section: {section}")
        if source:
            ev_parts.append(f"Source: {source}")
        return {
            "RuleId":               rule_id,
            "Area":                 area,
            "Category":             category,
            "Severity":             severity,
            "RoadmapPhase":         phase,
            "Workstream":           workstream,
            "Finding":              finding,
            "Recommendation":       remediation,
            "WhyFlagged":           f"Flagged because the assessment observed: {finding}",
            "TechnicalRemediation": remediation,
            "CurrentEvidence":      finding,
            "EvidenceLocation":     "; ".join(ev_parts),
        }

    # -----------------------------------------------------------------------
    # COLLABORATION — Ownership & Stewardship
    # -----------------------------------------------------------------------
    og = _derived_first("OwnershipGovernanceSummary")
    unmanaged = int(og.get("UnmanagedObjectCount", 0) or 0)
    missing_owners = int(og.get("MissingOwnerCount", 0) or 0)
    owner_health = int(og.get("OwnerHealthRiskCount", 0) or 0)
    od_mismatch = int(og.get("OneDriveOwnerMismatchCount", 0) or 0)

    if unmanaged > 0:
        findings.append(_hf(
            "COL-001",
            "Ownership & Stewardship", "Unmanaged Objects",
            "High", "Near Term", "Collaboration",
            f"{unmanaged} unmanaged collaboration object(s) identified (missing owners, stale owners, or unknown state).",
            "Assign an active accountable owner or documented steward to each flagged asset, confirm the business "
            "purpose, and retire spaces that no longer have a sponsor.",
            "OwnershipGovernanceSummary", "Ownership Governance", "Derived/OwnershipGovernance",
        ))

    if owner_health > 0 and "Owner Health" not in bpf_categories:
        findings.append(_hf(
            "COL-002",
            "Ownership & Stewardship", "Owner Health",
            "Medium", "Planned", "Collaboration",
            f"{owner_health} collaboration object(s) are owned by disabled or stale owner accounts.",
            "Reassign ownership from disabled or stale accounts to active custodians and formalize backup ownership coverage.",
            "UnmanagedObjects", "ownership-governance", "Derived/OwnershipGovernance",
        ))

    if od_mismatch > 0 and "OneDrive Ownership Mismatch" not in bpf_categories:
        findings.append(_hf(
            "COL-003",
            "Ownership & Stewardship", "OneDrive Ownership Mismatch",
            "Medium", "Planned", "Collaboration",
            f"{od_mismatch} OneDrive site(s) have a current owner different from the URL-derived default owner.",
            "Review OneDrive sites where the current owner differs from the URL-derived default user and confirm documented stewardship.",
            "UnmanagedObjects", "ownership-governance", "Derived/OwnershipGovernance",
        ))

    # -----------------------------------------------------------------------
    # COLLABORATION — Teams Governance
    # -----------------------------------------------------------------------
    teams = _rows("Collaboration", "AllTeams")
    ownerless_teams = [t for t in teams if _iget(t, "OwnerCount", 1) == 0 and not t.get("IsArchived")]
    if ownerless_teams and "Teams Ownership" not in bpf_categories:
        findings.append(_hf(
            "TM-001",
            "Teams Collaboration", "Teams Ownership",
            "High", "Near Term", "Collaboration",
            f"{len(ownerless_teams)} team(s) do not have any owners.",
            "Assign at least one active owner to every flagged Team, confirm the business purpose, "
            "and archive Teams that no longer need to remain active.",
            "AllTeams", "Teams Governance", "Hybrid/Teams",
        ))

    today = datetime.now(timezone.utc).date()
    dormant_teams = []
    for t in teams:
        lad = t.get("LastActivityDate")
        if lad:
            try:
                d = datetime.strptime(str(lad)[:10], "%Y-%m-%d").date()
                if (today - d).days > 90 and not t.get("IsArchived"):
                    dormant_teams.append(t)
            except ValueError:
                pass
    if dormant_teams:
        findings.append(_hf(
            "TM-006",
            "Teams Collaboration", "Teams Activity",
            "Low", "Monitor", "Collaboration",
            f"{len(dormant_teams)} team(s) with activity older than 90 days detected.",
            "Review Teams, Microsoft 365 groups, SharePoint, and OneDrive locations for accountable ownership, "
            "lifecycle state, guest exposure, and storage growth.",
            "AllTeams", "Teams Governance", "Hybrid/Teams",
        ))

    # TeamsGroupsCleanupCandidates
    tgcc = _rows("Collaboration", "TeamsGroupsCleanupCandidates")
    ownerless_c = sum(1 for t in tgcc if _iget(t, "OwnerCount", 1) == 0)
    no_members_c = sum(1 for t in tgcc if _iget(t, "MemberCount", 1) == 0)
    dormant_c = sum(1 for t in tgcc if _is_stale(t.get("LastActivityDate"), 90))
    if tgcc:
        findings.append(_hf(
            "TM-009",
            "Teams Collaboration", "Teams Cleanup",
            "Medium", "Planned", "Collaboration",
            f"{len(tgcc)} cleanup candidate(s) identified; "
            f"ownerless={ownerless_c}; no members={no_members_c}; dormant={dormant_c}.",
            "Review Teams and Microsoft 365 groups for accountable ownership, lifecycle state, and guest exposure. "
            "Use ownership assignment, archive or expiration decisions, and documented exceptions.",
            "TeamsGroupsCleanupCandidates", "Teams Governance", "Summary/TeamsGroupsCleanup",
        ))

    # ExternalExposureFindings — collaboration angle (guest-enabled ownerless groups)
    eef = _rows("Tenant", "ExternalExposureFindings")
    collab_eef = [r for r in eef if "Group" in str(r.get("AssetType", "")) or "Team" in str(r.get("Workload", ""))]
    if collab_eef:
        findings.append(_hf(
            "TM-008",
            "Teams Collaboration", "External Exposure",
            "Medium", "Planned", "Collaboration",
            f"{len(collab_eef)} externally relevant Teams or groups show guest-heavy, dormant, or ownerless patterns.",
            "Assign an active accountable owner or documented steward to each flagged asset, confirm the business "
            "purpose, and retire spaces that no longer have a sponsor.",
            "ExternalExposureFindings", "External Exposure Review", "Summary/ExternalExposure",
        ))

    # SharePoint stale sites
    sp_sites = _rows("Collaboration", "SharePoint")
    stale_sp = [s for s in sp_sites if _is_stale(s.get("LastContentModifiedDate"), 180)]
    od_sites = _rows("Collaboration", "OneDrive")
    stale_od = [s for s in od_sites if _is_stale(s.get("LastContentModifiedDate"), 180)]
    if stale_sp or stale_od:
        findings.append(_hf(
            "COL-004",
            "SharePoint & OneDrive", "Inactive Sites",
            "Low", "Monitor", "Collaboration",
            f"{len(stale_sp)} stale SharePoint site(s); {len(stale_od)} stale OneDrive site(s) "
            f"(no activity in 180+ days).",
            "Review Teams, Microsoft 365 groups, SharePoint, and OneDrive locations for lifecycle state "
            "and storage growth. Use archive or expiration decisions for inactive locations.",
            "SharePoint", "Storage and Activity", "Hybrid/Collaboration",
        ))

    # -----------------------------------------------------------------------
    # ENDPOINT — Device Management
    # -----------------------------------------------------------------------
    # DeviceManagementSummary is a flat scalar dict — _first() doesn't work; access directly
    dms = (data.get("Identity") or {}).get("DeviceManagementSummary") or {}
    total_devices = int(dms.get("TotalDevices", 0) or 0)
    managed = int(dms.get("ManagedDevices", 0) or 0)
    unmanaged_d = int(dms.get("UnmanagedDevices", 0) or 0)
    non_compliant = int(dms.get("NonCompliantDevices", 0) or 0)
    unsupported_os = int(dms.get("UnsupportedOsDevices", 0) or 0)

    if non_compliant > 0 and total_devices > 0 and "Device Compliance" not in bpf_categories:
        findings.append(_hf(
            "DEV-002",
            "Devices", "Device Compliance",
            "Medium", "Planned", "Endpoint",
            f"{non_compliant} non-compliant device(s) identified in the assessment snapshot.",
            "Review each non-compliant device with endpoint operations, validate whether it should be remediated, "
            "excluded, or removed from access, and resolve undocumented exceptions.",
            "DeviceDetails", "Devices", "Hybrid/Devices",
        ))

    if total_devices > 0 and managed / total_devices < 0.5:
        pct = round(managed / total_devices * 100, 1)
        findings.append(_hf(
            "DEV-004",
            "Devices", "Intune Enrollment",
            "Low", "Monitor", "Endpoint",
            f"{pct}% of discovered devices show Intune management ({managed}/{total_devices}).",
            "Review endpoint-management scope, close enrollment gaps for devices expected to access protected "
            "resources, and document approved unmanaged exceptions.",
            "DeviceDetails", "Devices", "Hybrid/Devices",
        ))

    if total_devices > 0 and unmanaged_d > 0:
        pct = round(unmanaged_d / total_devices * 100, 1)
        findings.append(_hf(
            "DEV-005",
            "Devices", "Unmanaged Devices",
            "Medium", "Planned", "Endpoint",
            f"{pct}% unmanaged ({unmanaged_d}/{total_devices}); device-management summary shows "
            f"a meaningful unmanaged-device population.",
            "Confirm the intended managed-device scope, review the unmanaged-device population, "
            "and close enrollment gaps for devices that should access protected resources.",
            "DeviceManagementSummary", "Devices", "Summary/DeviceManagement",
        ))

    if unsupported_os > 0:
        findings.append(_hf(
            "DEV-006",
            "Devices", "Unsupported OS",
            "Medium", "Planned", "Endpoint",
            f"{unsupported_os} unsupported or older operating-system version device(s) detected.",
            "Review endpoint-management scope, stale devices, and unsupported operating systems. "
            "Document approved exceptions before enforcing stricter access controls.",
            "DeviceManagementSummary", "Devices", "Summary/DeviceManagement",
        ))

    # -----------------------------------------------------------------------
    # IDENTITY — Inactive Users / Admins / Guests
    # -----------------------------------------------------------------------
    all_users = (data.get("Identity") or {}).get("Users") or {}
    if isinstance(all_users, dict):
        all_user_list = list(all_users.values())
    else:
        all_user_list = all_users if isinstance(all_users, list) else []

    _now_date = datetime.now(timezone.utc)

    def _days_since(u: dict, threshold: int) -> bool:
        sia  = u.get("signInActivity") or {}
        last = sia.get("lastSignInDateTime") or sia.get("lastNonInteractiveSignInDateTime")
        if not last:
            return True  # no sign-in data counts as inactive
        try:
            dt = datetime.fromisoformat(last.replace("Z", "+00:00"))
            return (_now_date - dt).days >= threshold
        except (ValueError, AttributeError):
            return True

    enabled_members = [u for u in all_user_list
                       if u.get("accountEnabled") and (u.get("userType") or "").lower() == "member"]
    enabled_guests  = [u for u in all_user_list
                       if u.get("accountEnabled") and (u.get("userType") or "").lower() == "guest"]

    inactive_members_180 = [u for u in enabled_members if _days_since(u, 180)]
    inactive_members_90  = [u for u in enabled_members if _days_since(u, 90)]
    inactive_guests_90   = [u for u in enabled_guests  if _days_since(u, 90)]

    if inactive_members_180 and "Inactive Users" not in bpf_categories:
        pct = round(len(inactive_members_180) / len(enabled_members) * 100) if enabled_members else 0
        findings.append(_hf(
            "ID-010",
            "Identity & Admins", "Inactive Users",
            "High", "Immediate", "Identity",
            f"{len(inactive_members_180)} enabled member account(s) ({pct}%) have no recorded sign-in within "
            "the last 180 days. These accounts retain active credentials and access rights "
            "despite showing no evidence of use.",
            "Review enabled member accounts with no recent sign-in activity. Disable or delete accounts "
            "that are no longer in use, remove license assignments from inactive accounts, and document "
            "any legitimate service accounts that authenticate non-interactively.",
            "Users", "Identity Hygiene", "Hybrid/Users",
        ))
    elif inactive_members_90 and "Inactive Users" not in bpf_categories:
        pct = round(len(inactive_members_90) / len(enabled_members) * 100) if enabled_members else 0
        findings.append(_hf(
            "ID-010",
            "Identity & Admins", "Inactive Users",
            "Medium", "Near Term", "Identity",
            f"{len(inactive_members_90)} enabled member account(s) ({pct}%) have no recorded sign-in "
            "within the last 90 days.",
            "Review enabled member accounts with no recent sign-in activity and disable or delete "
            "accounts that are no longer in use.",
            "Users", "Identity Hygiene", "Hybrid/Users",
        ))

    # Inactive admins — cross-reference Admins list with Users signInActivity
    admins_raw_list = (data.get("Identity") or {}).get("Admins") or []
    if isinstance(admins_raw_list, dict):
        admins_raw_list = list(admins_raw_list.values())
    priv_admin_list = [a for a in admins_raw_list if a.get("accountEnabled", True)]
    inactive_admins = []
    for admin in priv_admin_list:
        upn = (admin.get("userPrincipalName") or "").lower()
        user_obj = all_users.get(upn) if isinstance(all_users, dict) else {}
        user_obj = user_obj or {}
        if _days_since(user_obj, 180):
            inactive_admins.append(admin)
    if inactive_admins and "Inactive Admin Accounts" not in bpf_categories:
        findings.append(_hf(
            "ADMIN-003",
            "Identity & Admins", "Inactive Admin Accounts",
            "High", "Immediate", "Identity",
            f"{len(inactive_admins)} privileged account(s) have no recorded sign-in within 180 days. "
            "Stale privileged accounts with active credentials represent the highest-risk dormant exposure in the tenant.",
            "Review privileged accounts with no recent sign-in. Remove or downscope role assignments from "
            "accounts that are no longer in use. Validate break-glass account procedures separately from "
            "routine privileged account hygiene.",
            "Admins", "Privileged Access", "Hybrid/Admins",
        ))

    if inactive_guests_90 and "Inactive Guest Users" not in bpf_categories:
        pct = round(len(inactive_guests_90) / len(enabled_guests) * 100) if enabled_guests else 0
        findings.append(_hf(
            "ID-002",
            "Identity & Admins", "Inactive Guest Users",
            "Medium", "Planned", "Identity",
            f"{len(inactive_guests_90)} enabled guest account(s) ({pct}% of all guests) have no recorded "
            "sign-in within the last 90 days. Guest accounts retain access to SharePoint sites, Teams, "
            "and shared resources regardless of activity state.",
            "Review inactive guest accounts with the business sponsor for each collaboration relationship. "
            "Disable or remove guests whose collaboration need has ended and consider a guest access review "
            "cadence in Entra Identity Governance.",
            "Users", "Guest Access", "Hybrid/Users",
        ))

    # -----------------------------------------------------------------------
    # GOVERNANCE — Licensing at capacity
    # -----------------------------------------------------------------------
    lic_skus = (data.get("Identity") or {}).get("LicenseSKUs") or {}
    if isinstance(lic_skus, dict):
        at_capacity_skus = [
            (name, sku)
            for name, sku in lic_skus.items()
            if (sku.get("prepaidUnits") or {}).get("enabled", 0) > 0
            and sku.get("consumedUnits", 0) >= (sku.get("prepaidUnits") or {}).get("enabled", 0)
        ]
    else:
        at_capacity_skus = []

    if at_capacity_skus and "At Capacity" not in bpf_categories:
        sku_names = ", ".join(f"'{n}'" for n, _ in at_capacity_skus[:6])
        findings.append(_hf(
            "LIC-001",
            "Licensing", "At Capacity",
            "Medium", "Near Term", "Governance",
            f"{len(at_capacity_skus)} license SKU(s) are fully allocated with no remaining seats: {sku_names}. "
            "At-capacity SKUs block new assignments silently and can cause onboarding failures.",
            "Review at-capacity SKUs against upcoming headcount changes. Reclaim licenses from inactive "
            "or disabled accounts before purchasing additional seats. Prioritize reclamation from SKUs "
            "identified in the License Optimization Candidates table.",
            "LicenseSKUs", "Licensing", "Data/LicenseSKUs",
        ))

    # -----------------------------------------------------------------------
    # MESSAGING — Unverified Domains
    # -----------------------------------------------------------------------
    tenant_domains = (data.get("Tenant") or {}).get("Domains") or {}
    if isinstance(tenant_domains, dict):
        unverified_domains = [
            name for name, d in tenant_domains.items()
            if not d.get("isVerified") and not d.get("isInitial")
        ]
    else:
        unverified_domains = []

    if unverified_domains and "Domain Verification" not in bpf_categories:
        names = ", ".join(f"'{n}'" for n in unverified_domains[:5])
        findings.append(_hf(
            "DOM-001",
            "Domains", "Domain Verification",
            "High", "Near Term", "Messaging",
            f"{len(unverified_domains)} accepted domain(s) are not verified: {names}. "
            "Unverified domains cannot be used for mail flow and may indicate abandoned or misconfigured domain relationships.",
            "Verify or remove unverified accepted domains. Unverified domains cannot send or receive mail and "
            "create ambiguity in mail routing. Remove domains that are no longer needed and complete "
            "verification for any domain that should remain active.",
            "Domains", "DNS Configuration", "Data/Tenant/Domains",
        ))

    # -----------------------------------------------------------------------
    # GOVERNANCE — Licensing (optimization candidates)
    # -----------------------------------------------------------------------
    loc = _rows("Identity", "LicenseOptimizationCandidates")
    inactive_lic = [r for r in loc if "inactive" in str(r.get("Issue", "")).lower()
                    or "disabled" in str(r.get("Issue", "")).lower()]
    duplicate_lic = [r for r in loc if "same sku" in str(r.get("Issue", "")).lower()]
    duplicate_suite = [r for r in loc if "suite" in str(r.get("SkuFamily", "")).lower()
                       or "duplicate suite" in str(r.get("Issue", "")).lower()]

    if inactive_lic:
        findings.append(_hf(
            "LIC-002",
            "Licensing", "Inactive Licensed Users",
            "Medium", "Near Term", "Governance",
            f"{len(inactive_lic)} inactive or disabled user(s) still appear to hold paid licenses.",
            "Review paid license assignments for capacity-constrained SKUs, reclaim licenses from inactive "
            "or ineligible accounts, and review duplicate direct-plus-group assignments.",
            "LicenseOptimizationCandidates", "Licensing", "Summary/LicenseOptimization",
        ))

    gls = _rows("Identity", "GroupLicensingSummary")
    group_lic_in_use = [g for g in gls if "group-based licensing in use" in str(g.get("RiskSignal", "")).lower()]
    ownerless_lic_groups = [g for g in gls if _iget(g, "OwnerCount", 1) == 0]

    if ownerless_lic_groups:
        findings.append(_hf(
            "LIC-004",
            "Licensing", "Ownerless License Groups",
            "Medium", "Near Term", "Governance",
            f"{len(ownerless_lic_groups)} license-managing group(s) without owners detected.",
            "Confirm that group-based licensing groups have accountable owners and a documented "
            "assignment-error review cadence.",
            "GroupLicensingSummary", "Licensing", "Summary/GroupLicensing",
        ))

    if group_lic_in_use:
        findings.append(_hf(
            "LIC-003",
            "Licensing", "Group-Based Licensing",
            "Low", "Planned", "Governance",
            f"{len(group_lic_in_use)} group-based licensing group(s) in use and should have documented ownership.",
            "Review paid license assignments for capacity-constrained SKUs, confirm licensing groups have "
            "accountable owners and a documented assignment-error review cadence.",
            "GroupLicensingSummary", "Licensing", "Summary/GroupLicensing",
        ))

    if duplicate_lic:
        findings.append(_hf(
            "LIC-005",
            "Licensing", "Duplicate Assignments",
            "Medium", "Near Term", "Governance",
            f"{len(duplicate_lic)} user(s) with both direct and group-based assignment for the same SKU detected.",
            "Remove duplicate direct assignments after confirming the group-based assignment is intentional and healthy.",
            "LicenseOptimizationCandidates", "Licensing", "Summary/LicenseOptimization",
        ))

    if duplicate_suite:
        findings.append(_hf(
            "LIC-006",
            "Licensing", "Duplicate Suite Assignments",
            "Low", "Planned", "Governance",
            f"{len(duplicate_suite)} likely duplicate Microsoft 365 or Office 365 suite assignment(s) detected.",
            "Review paid license assignments and align suite assignments to user role without unnecessary overlap.",
            "LicenseOptimizationCandidates", "Licensing", "Summary/LicenseOptimization",
        ))

    # Diagnostics
    warn_count = int(diag.get("WarningCount", 0) or 0)
    err_count = int(diag.get("ErrorCount", 0) or 0)
    findings.append(_hf(
        "DIAG-001",
        "Collection Diagnostics", "Diagnostics",
        "Low", "Monitor", "Governance",
        f"Snapshot contains collection diagnostics that should be reviewed before using "
        f"the assessment as a remediation baseline. Warnings={warn_count}; Errors={err_count}",
        "Review collector warnings and errors, then re-run collection if any important workload data was incomplete.",
        "Diagnostics", "Collection Diagnostics", "Summary/Diagnostics",
    ))

    # -----------------------------------------------------------------------
    # IDENTITY — Conditional Access
    # -----------------------------------------------------------------------
    ca_summary = _first("Identity", "ConditionalAccessPolicySummary")
    report_only = int(ca_summary.get("ReportOnlyPolicies", 0) or 0)
    with_exclusions = int(ca_summary.get("PoliciesWithExclusions", 0) or 0)

    if report_only > 0:
        findings.append(_hf(
            "CA-002",
            "Conditional Access & MFA", "Report-Only Policies",
            "Medium", "Planned", "Identity",
            f"{report_only} Conditional Access policy/policies remain in report-only mode.",
            "Validate impact in report-only or pilot scope, protect privileged users, guests, MFA registration, "
            "and legacy-authentication scenarios, then move approved policies into enforcement.",
            "ConditionalAccessPolicies", "Conditional Access", "Hybrid/ConditionalAccessPolicies",
        ))

    if with_exclusions > 0:
        findings.append(_hf(
            "CA-007",
            "Conditional Access & MFA", "Policies With Exclusions",
            "Medium", "Planned", "Identity",
            f"{with_exclusions} Conditional Access policy/policies appear to use exclusions.",
            "Review Conditional Access exclusions as part of staged deployment: minimize and document all "
            "exclusions and avoid using Security Defaults and Conditional Access as overlapping baseline models.",
            "ConditionalAccessPolicies", "Conditional Access", "Hybrid/ConditionalAccessPolicies",
        ))

    # CA optimization overlap candidates
    cao = _rows("Identity", "ConditionalAccessOptimization")
    overlap = [r for r in cao if "overlap" in str(r.get("Signal", "")).lower()
               or "duplicate" in str(r.get("Signal", "")).lower()]
    if not overlap:
        # Fall back to rows with Review status and policy count > 1
        overlap = [r for r in cao if r.get("Status") == "Review" and int(r.get("PolicyCount", 0) or 0) > 1]
    if overlap:
        overlap_count = sum(int(r.get("PolicyCount", 1) or 1) for r in overlap)
        findings.append(_hf(
            "CA-019",
            "Conditional Access & MFA", "Policy Overlap",
            "Low", "Monitor", "Identity",
            f"{overlap_count} Conditional Access policy/policies share a similar target, "
            f"condition, grant, and session-control signature.",
            "Review the Conditional Access baseline and consolidate or intentionally retain overlapping policies "
            "with documented rationale.",
            "ConditionalAccessOptimization", "Conditional Access", "Summary/ConditionalAccessOptimization",
        ))

    # -----------------------------------------------------------------------
    # IDENTITY — Global Administrator standing access
    # -----------------------------------------------------------------------
    priv_rem = (data.get("Identity") or {}).get("PrivilegedAccessRemediationSummary") or []
    stale_row = next((r for r in priv_rem if "Stale privileged" in str(r.get("Signal", ""))), {})
    stale_ga  = int(stale_row.get("Count", 0) or 0)
    # Count enabled GAs directly from Admins list for consistency with exec summary
    admins_all = (data.get("Identity") or {}).get("Admins") or []
    admins_all = admins_all if isinstance(admins_all, list) else list(admins_all.values())
    ga_count   = sum(1 for a in admins_all if a.get("RoleName") == "Global Administrator" and a.get("accountEnabled", True) is not False)

    if ga_count > 4 and "Global Administrator" not in bpf_categories:
        findings.append(_hf(
            "ADMIN-001",
            "Identity & Admins", "Global Administrator Count",
            "High", "Immediate", "Identity",
            f"{ga_count} standing Global Administrator accounts detected. "
            "Microsoft recommends 2-4 for most organizations, limited to emergency break-glass scenarios.",
            "Reduce standing Global Administrator accounts. Use Privileged Identity Management (PIM) for "
            "just-in-time elevation, assign least-privilege roles for routine administrative tasks, "
            "and restrict permanent Global Admin to 2-4 break-glass accounts with documented procedures.",
            "Admins", "Privileged Access", "Hybrid/Admins",
        ))

    # Compute stale enabled GAs directly from admins list (signInActivity may be missing on older snapshots)
    _now_plan = datetime.now(timezone.utc)
    def _ga_is_stale(a: dict) -> bool:
        sia = a.get("signInActivity") or {}
        last = sia.get("lastSignInDateTime") or sia.get("lastNonInteractiveSignInDateTime")
        if not last:
            return True  # no sign-in data → count as stale
        try:
            return (_now_plan - datetime.fromisoformat(last.replace("Z", "+00:00"))).days >= 180
        except (ValueError, AttributeError):
            return True
    enabled_gas = [a for a in admins_all if a.get("RoleName") == "Global Administrator" and a.get("accountEnabled", True) is not False]
    stale_enabled_ga = sum(1 for a in enabled_gas if _ga_is_stale(a))

    if stale_enabled_ga > 0 and stale_enabled_ga < ga_count and "Stale Privileged" not in bpf_categories:
        findings.append(_hf(
            "ADMIN-002",
            "Identity & Admins", "Stale Global Administrators",
            "High", "Near Term", "Identity",
            f"{stale_enabled_ga} of {ga_count} enabled Global Administrator account(s) show no recent interactive sign-in. "
            "Stale privileged accounts that are not actively monitored represent persistent credential risk.",
            "Review stale Global Administrator accounts. Remove unnecessary role assignments, "
            "validate emergency-access account procedures, and confirm break-glass credentials are tested.",
            "Admins", "Privileged Access", "Hybrid/Admins",
        ))

    # -----------------------------------------------------------------------
    # IDENTITY — MFA enforcement gap
    # -----------------------------------------------------------------------
    gap_users = (data.get("Identity") or {}).get("MfaEnforcementGapUsers") or []
    gap_count = len(gap_users) if isinstance(gap_users, list) else 0
    if gap_count > 0 and "MFA Enforcement Gap" not in bpf_categories:
        findings.append(_hf(
            "MFA-001",
            "Conditional Access & MFA", "MFA Enforcement Gap",
            "High", "Immediate", "Identity",
            f"{gap_count} member account(s) have no MFA registration and can authenticate with only a password.",
            "Create or update a Conditional Access policy to require MFA for all users. "
            "Run report-only mode to validate scope, then enforce. "
            "Address unregistered accounts with a targeted self-service MFA registration campaign.",
            "MfaEnforcementGapUsers", "MFA Enrollment", "Summary/MfaEnforcementGap",
        ))

    # -----------------------------------------------------------------------
    # SECURITY — Secure Score
    # -----------------------------------------------------------------------
    score_data = (data.get("Security") or {}).get("SecuritySecureScore") or {}
    current_sc = float(score_data.get("currentScore") or 0)
    max_sc     = float(score_data.get("maxScore") or 0)
    if current_sc and max_sc and "Secure Score" not in bpf_categories:
        pct_sc   = round(current_sc / max_sc * 100)
        comp_sc  = score_data.get("averageComparativeScores") or []
        avg_sc   = next((float(s.get("averageScore", 0)) for s in comp_sc if s.get("basis") == "AllTenants"), None)
        sev_sc   = "High" if pct_sc < 30 else "Medium"
        phase_sc = "Immediate" if pct_sc < 30 else "Planned"
        avg_str  = f" Cross-tenant average is {avg_sc:.0f} points ({current_sc - avg_sc:+.0f} vs. this tenant)." if avg_sc else ""
        findings.append(_hf(
            "SEC-001",
            "Security", "Secure Score",
            sev_sc, phase_sc, "Security",
            f"Microsoft Secure Score is {current_sc:.0f} of {max_sc:.0f} points ({pct_sc}%).{avg_str} "
            "The top unimplemented controls and their point values are documented in the Security workstream section.",
            "Address the highest-impact unimplemented Secure Score controls. "
            "Prioritize controls that overlap with the Identity and MFA remediation items already in this plan -- "
            "closing those gaps will improve Secure Score as a side effect.",
            "SecuritySecureScore", "Secure Score", "Data/Security",
        ))

    # MFA weak methods
    mfa_posture = _rows("Identity", "MfaMethodPostureSummary")
    weak_row = next((r for r in mfa_posture if "weak" in str(r.get("Signal", "")).lower()), None)
    if weak_row and int(weak_row.get("UserCount", 0) or 0) > 0:
        findings.append(_hf(
            "MFA-002",
            "Conditional Access & MFA", "Weak MFA Methods",
            "Medium", "Planned", "Identity",
            f"MFA enrollment still relies on weaker methods for part of the tenant. "
            f"{weak_row.get('CurrentState', '')}",
            "Reduce SMS, voice, and email OTP reliance; encourage Microsoft Authenticator, passwordless, "
            "FIDO2/passkeys, or other phishing-resistant methods for privileged and sensitive access.",
            "MfaEnrollmentSummary", "MFA Enrollment", "Heuristic",
        ))

    # Guest lifecycle (GuestSignInSummary is a flat scalar dict — access directly)
    guest_summary = (data.get("Identity") or {}).get("GuestSignInSummary") or {}
    total_guests = int(guest_summary.get("GuestCount", 0) or 0)
    inactive_guests_90 = int(guest_summary.get("InactiveGuests90Days", 0) or 0)

    if inactive_guests_90 > 0:
        findings.append(_hf(
            "ID-002",
            "Identity & Admins", "Guest Lifecycle",
            "Medium", "Planned", "Identity",
            f"Guest lifecycle hygiene needs review. "
            f"{total_guests} guests; {inactive_guests_90} inactive/stale (90+ days).",
            "Review inactive guests with the business sponsor, confirm whether each guest still needs access, "
            "and remove or disable accounts that no longer support an approved collaboration scenario.",
            "Users", "Guests", "Hybrid/Users",
        ))

        findings.append(_hf(
            "ID-005",
            "Identity & Admins", "Inactive Guests",
            "Medium", "Near Term", "Identity",
            f"Inactive guest accounts identified by the guest sign-in governance summary. "
            f"{inactive_guests_90} inactive guest account(s) over 90 days.",
            "Review inactive guests with the business sponsor, confirm whether each guest still needs access, "
            "and remove or disable accounts that no longer support an approved collaboration scenario.",
            "GuestSignInSummary", "Guest Access", "Summary/GuestSignIn",
        ))

    # Privileged access stale accounts (PrivilegedAccessSummary is a mixed dict — access directly)
    priv_summary = (data.get("Identity") or {}).get("PrivilegedAccessSummary") or {}
    stale_priv = int(priv_summary.get("StalePrivilegedAccounts90Days", 0) or 0)

    if stale_priv > 0:
        findings.append(_hf(
            "ID-003",
            "Identity & Admins", "Stale Privileged Accounts",
            "Medium", "Near Term", "Identity",
            f"Enabled privileged accounts appear stale based on sign-in recency. "
            f"{stale_priv} stale privileged account(s).",
            "Validate whether stale privileged accounts are still needed and remove or downgrade "
            "unused standing admin roles.",
            "Admins", "Privileged Access", "Hybrid/Admins",
        ))

        findings.append(_hf(
            "ID-006",
            "Identity & Admins", "Stale Privileged Summary",
            "Medium", "Near Term", "Identity",
            f"Privileged identities with stale sign-in activity identified by the "
            f"privileged-access governance summary. {stale_priv} stale privileged account(s) over 90 days.",
            "Review stale privileged accounts, remove unused role assignments, and validate "
            "emergency access documentation.",
            "PrivilegedAccessSummary", "Privileged Access", "Summary/PrivilegedAccess",
        ))

    # External identity / cross-tenant trust posture
    ext_eef = [r for r in eef
               if "external identity" in str(r.get("Workload", "")).lower()
               or "cross-tenant" in str(r.get("Workload", "")).lower()
               or "cross-tenant" in str(r.get("ExposureCategory", "")).lower()]
    if ext_eef:
        obs = "; ".join(r.get("ObservedSetting", "") for r in ext_eef if r.get("ObservedSetting"))
        findings.append(_hf(
            "ID-008",
            "Identity & Admins", "External Trust Posture",
            "Medium", "Planned", "Identity",
            f"External invitation or cross-tenant trust posture should be reviewed. "
            f"{len(ext_eef)} external identity / trust review row(s); {obs}",
            "Review cross-tenant access settings, B2B collaboration policies, and external identity trust "
            "configuration. Confirm that inbound MFA trust, cross-tenant partner trust, and guest invitation "
            "scope are explicitly configured to match the approved external-access baseline.",
            "ExternalExposureFindings", "External Exposure Review", "Summary/ExternalExposure",
        ))

    # -----------------------------------------------------------------------
    # IDENTITY — Guest invite policy
    # -----------------------------------------------------------------------
    gac = (data.get("Identity") or {}).get("GuestAccessConfiguration") or {}
    invite_from = gac.get("allowInvitesFrom", "")
    if invite_from == "everyone" and "Guest Invite Policy" not in bpf_categories:
        findings.append(_hf(
            "GUEST-001",
            "Identity & Admins", "Guest Invite Policy",
            "Medium", "Near Term", "Identity",
            "Guest invitations are permitted from 'everyone', meaning any user in the tenant can invite "
            "external guests without admin approval. This creates unmanaged external access surface.",
            "Restrict guest invitation permissions to admins and designated guest inviters. "
            "Set allowInvitesFrom to 'adminsAndGuestInviters' in the authorization policy and "
            "designate a controlled set of approved inviters.",
            "GuestAccessConfiguration", "Guest Access", "Identity/GuestAccess",
        ))

    # -----------------------------------------------------------------------
    # IDENTITY — AD Connect / Hybrid sync staleness
    # -----------------------------------------------------------------------
    adc = (data.get("Tenant") or {}).get("AdConnectConfiguration") or {}
    adc_summary = adc.get("Summary") or {}
    sync_enabled  = adc_summary.get("OnPremisesSyncEnabled", False)
    last_sync_raw = adc_summary.get("LastSyncDateTime") or ""
    if sync_enabled and last_sync_raw:
        try:
            sync_dt   = datetime.fromisoformat(last_sync_raw.replace("Z", "+00:00"))
            sync_days = (datetime.now(timezone.utc) - sync_dt).days
            if sync_days > 3 and "AD Connect Sync" not in bpf_categories:
                sev_sync   = "High" if sync_days > 30 else "Medium"
                phase_sync = "Immediate" if sync_days > 30 else "Near Term"
                findings.append(_hf(
                    "HYBRID-001",
                    "Identity & Admins", "AD Connect Sync Staleness",
                    sev_sync, phase_sync, "Identity",
                    f"Azure AD Connect last sync was {sync_days} days ago ({last_sync_raw[:10]}). "
                    "A stale sync means on-premises user changes (disable, password reset, role removal) "
                    "are not reflected in Entra ID, leaving access controls out of date.",
                    "Investigate the AD Connect sync health. Check the synchronization service manager "
                    "on the AD Connect server for errors, validate that the sync account has required "
                    "permissions, and restore the sync cycle. Consider upgrading to Entra Cloud Sync "
                    "if the on-premises AD Connect server is out of date.",
                    "AdConnectConfiguration", "Hybrid Sync", "Tenant/HybridConfig",
                ))
        except (ValueError, AttributeError):
            pass

    # -----------------------------------------------------------------------
    # MESSAGING — Exchange
    # -----------------------------------------------------------------------
    mailboxes = _rows("Exchange", "AllMailboxes")
    fwd_mailboxes = [
        m for m in mailboxes
        if m.get("ForwardingSmtpAddress") or m.get("ForwardingAddress")
    ]
    fps = _first("Exchange", "ForwardingPolicySummary")
    remote_fwd = int(fps.get("RemoteDomainsAllowingAutoForwarding", 0) or 0)
    policies_allow = int(fps.get("PoliciesExplicitlyAllowingAutoForwarding", 0) or 0)

    if fwd_mailboxes or policies_allow > 0:
        fwd_detail = (
            f"{len(fwd_mailboxes)} mailbox(es) with forwarding configured; "
            f"{policies_allow} hosted outbound policy/policies explicitly allow auto-forwarding; "
            f"{remote_fwd} remote domain(s) have AutoForwardEnabled; "
            f"Policy modes: {fps.get('PolicyAutoForwardingModes', '')}"
        )
        findings.append(_hf(
            "EX-001",
            "Exchange", "Mailbox Forwarding",
            "High", "Near Term", "Messaging",
            f"Mailbox forwarding is enabled for one or more mailboxes. {fwd_detail}",
            "Review Exchange mail-flow dependencies across accepted domains, SPF, DKIM, DMARC, connectors, "
            "remote domains, SMTP relay paths, mailbox forwarding, inbox-rule forwarding, and public folders. "
            "Remove unsupported paths, tighten relay and auto-forwarding exceptions, and document approved "
            "mail-routing dependencies.",
            "AllMailboxes", "Mailbox Forwarding", "Hybrid/Exchange",
        ))

    # Public folders
    pf = _rows("Exchange", "PublicFolderDetails")
    if pf:
        findings.append(_hf(
            "EX-004",
            "Exchange", "Public Folders",
            "Medium", "Planned", "Messaging",
            f"Public folders are still present in the tenant. {len(pf)} public folder object(s).",
            "Review Exchange mail-flow dependencies. Remove unsupported paths and document approved "
            "mail-routing dependencies.",
            "PublicFolderDetails", "Public Folders", "Hybrid/Exchange",
        ))

    # Shared mailbox governance
    smg = _first("Exchange", "SharedMailboxGovernanceSummary")
    smb_total = int(smg.get("SharedMailboxCount", 0) or 0)
    smb_no_owner = int(smg.get("SharedMailboxesWithoutOwnerSignal", 0) or 0)
    smb_oversized = int(smg.get("OversizedSharedMailboxes", 0) or 0)

    if smb_no_owner > 0 or smb_oversized > 0:
        findings.append(_hf(
            "EX-005",
            "Exchange", "Shared Mailbox Governance",
            "Low", "Monitor", "Messaging",
            f"Shared mailbox governance requires review. "
            f"Ownerless={smb_no_owner}; oversized={smb_oversized} (of {smb_total} total).",
            "Review shared mailbox ownership and oversized growth. Remove unsupported mail-routing paths "
            "and document approved dependencies.",
            "AllMailboxes", "Shared Mailboxes", "Hybrid/Exchange",
        ))

        findings.append(_hf(
            "EX-007",
            "Exchange", "Shared Mailbox Summary",
            "Low", "Monitor", "Messaging",
            f"Shared mailbox governance summary shows ownership or growth gaps. "
            f"Oversized={smb_oversized}; lacking owner signal={smb_no_owner}.",
            "Review shared mailbox ownership and oversized growth. Remove unsupported mail-routing paths "
            "and document approved dependencies.",
            "SharedMailboxGovernanceSummary", "Shared Mailboxes", "Summary/SharedMailboxGovernance",
        ))

    # -----------------------------------------------------------------------
    # SECURITY — Consent governance
    # -----------------------------------------------------------------------
    auth_cfg = _first("Identity", "AuthenticationConfig")
    pgp = auth_cfg.get("PermissionGrantPoliciesAssigned") or []
    if isinstance(pgp, str):
        pgp = [p.strip() for p in pgp.split(";") if p.strip()]
    # Flag if broad user-consent policies are present
    broad_consent = [p for p in pgp if "user-default-allow" in p.lower() or "user-default-recommended" in p.lower()]
    if broad_consent:
        findings.append(_hf(
            "SEC-003",
            "Security", "Consent Governance",
            "Low", "Planned", "Security",
            f"Permission grant policy settings may allow broader user consent than desired. "
            f"{'; '.join(pgp)}",
            "Review enterprise applications for accountable owners, business purpose, granted delegated and "
            "application permissions. Remove unused applications, reduce broad consent where possible, and "
            "avoid long-lived client secrets for production integrations.",
            "AuthenticationConfig", "Consent Governance", "Hybrid/AuthenticationConfig",
        ))

    return findings


def _is_stale(date_str: object, threshold_days: int) -> bool:
    """Return True if date_str is older than threshold_days from today."""
    if not date_str:
        return False
    try:
        d = datetime.strptime(str(date_str)[:10], "%Y-%m-%d").date()
        return (datetime.now(timezone.utc).date() - d).days > threshold_days
    except ValueError:
        return False


def _iget(obj: dict, key: str, default: int) -> int:
    """Get integer field from dict, using default only when value is None (not when 0)."""
    val = obj.get(key)
    if val is None:
        return default
    try:
        return int(val)
    except (TypeError, ValueError):
        return default


# ---------------------------------------------------------------------------
# Area -> Workstream (mirrors PS Get-OwnerTeam regex logic)
# ---------------------------------------------------------------------------

def _get_workstream(area: str, category: str = "") -> str:
    lookup = f"{area} {category}".lower()
    if re.search(
        r"\bidentity\b|conditional access|\bmfa\b|admin|privileged|entra"
        r"|authentication|\bguest\b|password",
        lookup,
    ):
        return "Identity"
    if re.search(
        r"exchange|\bmail\b|\bsmtp\b|public folder|connector|spam"
        r"|inactive mailbox|forwarding|dmarc|anti.?spoof|\bdomains?\b",
        lookup,
    ):
        return "Messaging"
    if re.search(
        r"sharepoint|onedrive|\bteams\b|\bgroups?\b|ownership"
        r"|collaboration|stewardship|\bsites?\b|external exposure",
        lookup,
    ):
        return "Collaboration"
    if re.search(r"\bdevices?\b|endpoint|intune|\bcompliance\b|\bretention\b", lookup):
        return "Endpoint"
    if re.search(r"secure score|\bsecurity\b|defender|zero trust", lookup):
        return "Security"
    return "Governance"


# ---------------------------------------------------------------------------
# RoadmapActions — one consolidated action bundle per workstream
# ---------------------------------------------------------------------------

_ACTION_TEMPLATES: dict[str, dict] = {
    "Collaboration": {
        "ActionTitle":        "Establish accountable ownership for collaboration spaces",
        "Theme":              "Ownership and collaboration lifecycle",
        "ExecutionPattern":   "Collaboration lifecycle governance",
        "DependencyTier":     "Moderate",
        "PrimaryOwner":       "Collaboration service owner / workspace sponsor",
        "FirstValidationStep":
            "Confirm which collaboration spaces need an owner, steward, transfer, or retirement decision first.",
        "SuccessCheck":
            "Each in-scope collaboration space has an accountable owner or a documented lifecycle decision.",
        "RelatedSection":     "Ownership Governance",
        "BusinessValue":
            "Addressing this helps improve ownership, governance, and lifecycle control for collaboration spaces and business content.",
        "WhyItMatters":
            "This matters because the assessment found collaboration assets without a clear accountable owner, "
            "which creates gaps in stewardship, lifecycle handling, and access governance.",
        "RecommendedNextStep":
            "Assign an active accountable owner or documented steward to each flagged asset, confirm the "
            "business purpose, and retire spaces that no longer have a sponsor.",
    },
    "Identity": {
        "ActionTitle":        "Reduce privileged access and strengthen identity controls",
        "Theme":              "Identity and privileged access governance",
        "ExecutionPattern":   "Privileged access and identity hygiene",
        "DependencyTier":     "Moderate",
        "PrimaryOwner":       "Identity and access administration",
        "FirstValidationStep":
            "Confirm the target identity protection baseline and decide which privileged, guest, and core "
            "user populations should be brought under it first.",
        "SuccessCheck":
            "Privileged and external access follow one approved protection model, with only documented exceptions remaining.",
        "RelatedSection":     "identity-admins",
        "BusinessValue":
            "Addressing this helps reduce identity compromise risk, strengthens access control, and improves the tenant security baseline.",
        "WhyItMatters":
            "This matters because the tenant has more standing Global Administrator assignments than the "
            "recommended operating threshold, increasing privileged access exposure.",
        "RecommendedNextStep":
            "Review each standing Global Administrator assignment using a least-privilege model, remove "
            "the role entirely from inactive or stale administrators, move infrequent admins to lower-privilege "
            "roles such as Global Reader where possible, and retain only the minimum approved permanent admins "
            "plus documented break-glass emergency access accounts.",
    },
    "Endpoint": {
        "ActionTitle":        "Improve device compliance and managed endpoint coverage",
        "Theme":              "Endpoint compliance and device control",
        "ExecutionPattern":   "Endpoint remediation",
        "DependencyTier":     "High",
        "PrimaryOwner":       "Endpoint engineering / device administration",
        "FirstValidationStep":
            "Confirm which devices are expected to retain access and which endpoint exceptions are still justified.",
        "SuccessCheck":
            "Protected access is limited to approved device states, with documented exceptions only.",
        "RelatedSection":     "devices",
        "BusinessValue":
            "Addressing this helps improve policy enforcement, reduces unmanaged access risk, and strengthens endpoint visibility.",
        "WhyItMatters":
            "This matters because the device posture output shows compliance results below the expected baseline, "
            "leaving managed access policies less effective.",
        "RecommendedNextStep":
            "Review why the compliance baseline is being missed, remediate the highest-volume failure "
            "conditions, and tighten exception handling for devices that should not remain non-compliant.",
    },
    "Messaging": {
        "ActionTitle":        "Review external forwarding and mail flow exposure",
        "Theme":              "Mail flow risk and anti-spoofing posture",
        "ExecutionPattern":   "Mail flow and forwarding review",
        "DependencyTier":     "Moderate",
        "PrimaryOwner":       "Messaging and email administration",
        "FirstValidationStep":
            "Confirm the approved domain trust and mail-authentication baseline before cleanup begins.",
        "SuccessCheck":
            "Required domains and mail-authentication controls align to the approved baseline.",
        "RelatedSection":     "domains",
        "BusinessValue":
            "Addressing this helps reduce data-loss and mail-flow risk while improving operational control over messaging.",
        "WhyItMatters":
            "This matters because the tenant mail-flow configuration shows forwarding exposure or domain "
            "hygiene gaps that could lead to data leakage or delivery failures.",
        "RecommendedNextStep":
            "Review Exchange mail-flow dependencies across accepted domains, SPF, DKIM, DMARC, connectors, "
            "remote domains, SMTP relay paths, mailbox forwarding, inbox-rule forwarding, and public folders. "
            "Remove unsupported paths, tighten relay and auto-forwarding exceptions, and document approved "
            "mail-routing dependencies.",
    },
    "Security": {
        "ActionTitle":        "Strengthen baseline security and access protections",
        "Theme":              "Security baseline and zero-trust readiness",
        "ExecutionPattern":   "Security baseline enforcement",
        "DependencyTier":     "High",
        "PrimaryOwner":       "Security operations / control owner",
        "FirstValidationStep":
            "Confirm the target protection baseline and sequence the first rollout wave around the highest-value control gaps.",
        "SuccessCheck":
            "The agreed protection baseline is active for the intended population and no longer relies on broad temporary exceptions.",
        "RelatedSection":     "secure-score",
        "BusinessValue":
            "Addressing this helps reduce identity compromise risk, strengthens access control, and improves the tenant security baseline.",
        "WhyItMatters":
            "This matters because the tenant currently shows a Microsoft Secure Score below the target range "
            "for a mature tenant baseline.",
        "RecommendedNextStep":
            "Use Microsoft Secure Score as a prioritization signal for high-value security improvements, then "
            "validate each recommended action against tenant risk, licensing, user impact, and operational "
            "ownership before implementation.",
    },
    "Governance": {
        "ActionTitle":        "Reconcile license capacity and tenant governance gaps",
        "Theme":              "Licensing, capacity, and tenant governance",
        "ExecutionPattern":   "Governance and licensing reconciliation",
        "DependencyTier":     "Moderate",
        "PrimaryOwner":       "Tenant governance and platform ownership",
        "FirstValidationStep":
            "Confirm the target licensing, governance, and external-access baseline before cleanup work begins.",
        "SuccessCheck":
            "Capacity, ownership, and governance decisions are aligned to the approved operating baseline.",
        "RelatedSection":     "Licensing",
        "BusinessValue":
            "Addressing this helps improve cost control, avoids licensing blockers, and makes future growth easier to plan.",
        "WhyItMatters":
            "This matters because the tenant shows inactive licensed users and group-based licensing "
            "governance gaps that create cost and accountability risk.",
        "RecommendedNextStep":
            "Review paid license assignments for capacity-constrained SKUs, reclaim licenses from inactive "
            "or ineligible accounts, review duplicate direct-plus-group assignments, and confirm that "
            "group-based licensing groups have accountable owners and a documented assignment-error review cadence.",
    },
}

_EFFORT_HOURS: dict[str, str] = {
    "Quick":   "12-22 hours",
    "Standard":"18-32 hours",
    "Complex": "28-48 hours",
}


def _build_roadmap_actions(findings: list[dict]) -> list[dict]:
    """Generate one consolidated RoadmapAction per workstream from findings."""
    by_ws: dict[str, list[dict]] = defaultdict(list)
    for f in findings:
        ws = f.get("Workstream", "Governance")
        by_ws[ws].append(f)

    actions: list[dict] = []
    for ws, ws_findings in sorted(by_ws.items()):
        tmpl = _ACTION_TEMPLATES.get(ws, _ACTION_TEMPLATES["Governance"])

        high_findings = [f for f in ws_findings if f.get("Severity") == "High"]
        top_f = high_findings[0] if high_findings else ws_findings[0]

        areas = sorted({f.get("Area", "") for f in ws_findings if f.get("Area")})
        areas_str = ", ".join(areas[:4])
        n = len(ws_findings)

        # EffortTier by finding count
        if n <= 3:
            effort = "Quick"
        elif n <= 10:
            effort = "Standard"
        else:
            effort = "Complex"

        # Phase = earliest (most urgent) phase in the group
        best_phase = min(
            ws_findings,
            key=lambda f: _PHASE_ORDER.get(f.get("RoadmapPhase", "Monitor"), 99),
        ).get("RoadmapPhase", "Monitor")

        highest_sev = min(
            ws_findings,
            key=lambda f: _SEV_ORDER.get(f.get("Severity", "Low"), 99),
        ).get("Severity", "Low")

        quick_win = effort == "Quick" or best_phase == "Immediate"

        example_text = str(top_f.get("CurrentEvidence") or top_f.get("Finding") or "")
        if len(example_text) > 120:
            example_text = example_text[:117] + "..."

        actions.append({
            "PriorityBand":         best_phase,
            "Priority":             best_phase,
            "RoadmapPhase":         best_phase,
            "ExecutionPattern":     tmpl["ExecutionPattern"],
            "EffortTier":           effort,
            "DependencyTier":       tmpl["DependencyTier"],
            "QuickWinEligible":     quick_win,
            "EstimatedPsHours":     _EFFORT_HOURS[effort],
            "ActionTitle":          tmpl["ActionTitle"],
            "Theme":                tmpl["Theme"],
            "Workstream":           ws,
            "HighestSeverity":      highest_sev,
            "FindingCount":         n,
            "WhatThisAddresses":    f"This work item addresses {n} related findings across {areas_str}.",
            "StandoutReason":       (
                f"This pattern stood out during review because the current source shows "
                f"{example_text} across {areas_str}, which is a visible concentration rather than "
                f"a one-off exception."
            ),
            "ExampleText":          example_text,
            "WhyItMatters":         tmpl["WhyItMatters"],
            "RecommendedNextStep":  tmpl["RecommendedNextStep"],
            "BusinessValue":        tmpl["BusinessValue"],
            "PrimaryOwner":         tmpl["PrimaryOwner"],
            "FirstValidationStep":  tmpl["FirstValidationStep"],
            "SuccessCheck":         tmpl["SuccessCheck"],
            "RelatedSection":       tmpl["RelatedSection"],
        })

    # Sort by phase then workstream
    actions.sort(key=lambda a: (
        _PHASE_ORDER.get(a["RoadmapPhase"], 99),
        a["Workstream"],
    ))
    return actions


# ---------------------------------------------------------------------------
# WorkstreamSummaries
# ---------------------------------------------------------------------------

def _build_workstream_summaries(findings: list[dict]) -> list[dict]:
    groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for f in findings:
        groups[(f.get("Workstream", "Governance"), f.get("Area", ""))].append(f)

    summaries: list[dict] = []
    for (workstream, area), group in sorted(groups.items()):
        top_sev = min(group, key=lambda x: _SEV_ORDER.get(x.get("Severity", "Low"), 99))
        sev_label = _DISPLAY_SEV_LABEL.get(top_sev.get("Severity", "Low"), "Info")

        cat_counts: dict[str, int] = {}
        for f in group:
            cat = f.get("Category", "")
            if cat:
                cat_counts[cat] = cat_counts.get(cat, 0) + 1
        top_signals = "; ".join(
            f"{c}: {n}" for c, n in sorted(cat_counts.items(), key=lambda x: -x[1])[:5]
        )

        summaries.append({
            "Workstream":    workstream,
            "Area":          area,
            "Severity":      sev_label,
            "OpenFindings":  len(group),
            "CriticalCount": sum(1 for f in group if f.get("Severity") == "High"),
            "WarningCount":  sum(1 for f in group if f.get("Severity") == "Medium"),
            "InfoCount":     sum(1 for f in group if f.get("Severity") == "Low"),
            "TopSignals":    top_signals,
        })

    summaries.sort(key=lambda s: (
        _SEV_ORDER.get({"Critical": "High", "Medium": "Medium", "Info": "Low"}.get(s["Severity"], "Low"), 99),
        s["Workstream"],
    ))
    return summaries


# ---------------------------------------------------------------------------
# ConsultativeSummaries
# ---------------------------------------------------------------------------

def _build_consultative_summaries(findings: list[dict]) -> dict[str, Any]:
    by_ws: dict[str, list[dict]] = defaultdict(list)
    for f in findings:
        by_ws[f.get("Workstream", "Governance")].append(f)

    summaries: dict[str, Any] = {}
    for ws, ws_findings in sorted(by_ws.items()):
        high   = [f for f in ws_findings if f.get("Severity") == "High"]
        medium = [f for f in ws_findings if f.get("Severity") == "Medium"]
        low    = [f for f in ws_findings if f.get("Severity") == "Low"]
        areas  = sorted({f.get("Area", "") for f in ws_findings if f.get("Area")})

        snapshot_rows = []
        for f in ws_findings[:8]:
            cat      = f.get("Category", "")
            evidence = f.get("CurrentEvidence", "")
            if cat and evidence:
                short_ev = evidence[:120] + "..." if len(evidence) > 120 else evidence
                snapshot_rows.append({"Signal": cat, "State": short_ev})

        summaries[f"{ws}ConsultativeSummary"] = {
            "Title":        f"{ws} Assessment",
            "Narrative":    _build_ws_narrative(ws, ws_findings, high, medium, low, areas),
            "FindingCount": len(ws_findings),
            "HighCount":    len(high),
            "MediumCount":  len(medium),
            "Areas":        areas,
            "SnapshotRows": snapshot_rows,
        }

    return summaries


def _build_ws_narrative(
    ws: str,
    all_findings: list[dict],
    high: list[dict],
    medium: list[dict],
    low: list[dict],
    areas: list[str],
) -> str:
    total     = len(all_findings)
    areas_str = ", ".join(areas[:3])

    def _find_rule(prefix: str) -> dict | None:
        return next(
            (f for f in all_findings if str(f.get("RuleId", "")).startswith(prefix)),
            None,
        )

    def _ev(f: dict | None, max_len: int = 160) -> str:
        if not f:
            return ""
        s = str(f.get("CurrentEvidence") or f.get("Finding") or "").rstrip(". ")
        return (s[:max_len] + "...") if len(s) > max_len else s

    parts: list[str] = []

    if ws == "Identity":
        admin_f  = _find_rule("ADMIN-001") or _find_rule("IDENTITYADMINS")
        stale_f  = _find_rule("ID-003") or _find_rule("ID-006")
        guest_f  = _find_rule("ID-002") or _find_rule("ID-005")
        ca_count = sum(
            1 for f in medium
            if "Conditional" in str(f.get("Area", "")) or str(f.get("RuleId", "")).startswith("CA")
        )
        if high:
            parts.append(
                f"The identity and access review identified {len(high)} high-priority gap(s) "
                f"across {areas_str}."
            )
        ev = _ev(admin_f)
        if ev:
            parts.append(f"Privileged access is the top concern: {ev}.")
        ev = _ev(stale_f)
        if ev:
            parts.append(f"Stale privileged account activity was also detected: {ev}.")
        ev = _ev(guest_f)
        if ev:
            parts.append(f"Guest and external identity hygiene requires review: {ev}.")
        if ca_count:
            parts.append(
                f"{ca_count} Conditional Access gap(s) are scheduled for planned remediation, "
                f"including policy staging and exclusion review."
            )
        if not parts:
            parts.append(f"The identity review covers {total} finding(s) across {areas_str}.")

    elif ws == "Collaboration":
        col_f     = _find_rule("COL-001") or _find_rule("OWNERSHIPSTEWARDSHIP")
        tm_f      = _find_rule("TM-001")
        cleanup_f = _find_rule("TM-009")
        if high:
            parts.append(
                f"The collaboration review identified {len(high)} high-priority ownership and "
                f"lifecycle governance gap(s) across Teams, SharePoint, and OneDrive."
            )
        ev = _ev(col_f)
        if ev:
            parts.append(f"Collaboration ownership gaps are the primary concern: {ev}.")
        ev = _ev(tm_f)
        if ev:
            parts.append(f"Teams governance gaps were also identified: {ev}.")
        ev = _ev(cleanup_f)
        if ev:
            parts.append(f"Teams and group lifecycle cleanup is recommended: {ev}.")
        if medium or low:
            parts.append(
                f"An additional {len(medium) + len(low)} medium and informational finding(s) "
                f"cover sharing policy, stale content, and external exposure review."
            )
        if not parts:
            parts.append(f"The collaboration review covers {total} finding(s) across {areas_str}.")

    elif ws == "Messaging":
        fwd_f = _find_rule("EX-001")
        pf_f  = _find_rule("EX-004")
        smb_f = _find_rule("EX-005") or _find_rule("EX-007")
        if high:
            parts.append(
                f"The messaging review identified {len(high)} high-priority mail-flow "
                f"risk(s) that require remediation or documented approval."
            )
        ev = _ev(fwd_f)
        if ev:
            parts.append(f"Mail forwarding exposure is the primary concern: {ev}.")
        ev = _ev(pf_f)
        if ev:
            parts.append(f"Public folder presence requires migration or retirement planning: {ev}.")
        ev = _ev(smb_f)
        if ev:
            parts.append(f"Shared mailbox governance should also be reviewed: {ev}.")
        if not parts:
            parts.append(f"The messaging review covers {total} finding(s) across {areas_str}.")

    elif ws == "Endpoint":
        comp_f     = _find_rule("DEVICES") or _find_rule("DEV-002")
        unmanaged_f = _find_rule("DEV-005")
        os_f       = _find_rule("DEV-006")
        if high:
            parts.append(
                f"The endpoint review identified {len(high)} high-priority device posture gap(s) "
                f"that affect secure access readiness."
            )
        ev = _ev(comp_f)
        if ev:
            parts.append(f"Device compliance is the primary concern: {ev}.")
        ev = _ev(unmanaged_f)
        if ev:
            parts.append(f"Unmanaged device exposure was also identified: {ev}.")
        ev = _ev(os_f)
        if ev:
            parts.append(f"Unsupported operating systems require attention: {ev}.")
        if not parts:
            parts.append(f"The endpoint review covers {total} finding(s) across {areas_str}.")

    elif ws == "Security":
        sec_f = _find_rule("SEC-001") or next(
            (f for f in all_findings if "Secure Score" in str(f.get("Area", ""))), None
        )
        consent_f = _find_rule("SEC-003")
        if high or medium:
            parts.append(
                f"The security baseline review identified {len(high) + len(medium)} finding(s) "
                f"across {areas_str}."
            )
        ev = _ev(sec_f)
        if ev:
            parts.append(f"Secure Score posture is the primary signal: {ev}.")
        ev = _ev(consent_f)
        if ev:
            parts.append(f"Application consent governance also requires review: {ev}.")
        if not parts:
            parts.append(f"The security review covers {total} finding(s) across {areas_str}.")

    elif ws == "Governance":
        lic_f         = _find_rule("LIC-002")
        ownerless_f   = _find_rule("LIC-004")
        diag_f        = _find_rule("DIAG-001")
        if high or medium:
            parts.append(
                f"The governance review identified {len(high) + len(medium)} finding(s) requiring "
                f"licensing and tenant hygiene attention across {areas_str}."
            )
        ev = _ev(lic_f)
        if ev:
            parts.append(f"Inactive licensed user cleanup is the top priority: {ev}.")
        ev = _ev(ownerless_f)
        if ev:
            parts.append(f"License-managing group ownership gaps were also found: {ev}.")
        ev = _ev(diag_f)
        if ev:
            parts.append(
                f"Collection diagnostics should be reviewed before treating findings as the "
                f"remediation baseline: {ev}."
            )
        if not parts:
            parts.append(f"The governance review covers {total} finding(s) across {areas_str}.")

    else:
        if high:
            parts.append(
                f"{ws} findings include {len(high)} high-priority item(s) requiring "
                f"immediate or near-term attention across {areas_str}."
            )
        if medium:
            parts.append(f"{len(medium)} medium-priority finding(s) are scheduled for planned remediation.")
        if low:
            parts.append(f"{len(low)} informational finding(s) are flagged for monitoring.")
        if not parts:
            parts.append(f"{ws} findings reviewed. {total} total finding(s) across {areas_str}.")

    return " ".join(parts)
