"""
Excel workbook generation — Python equivalent of Export-HashTableToExcel.
Uses openpyxl instead of the ImportExcel PowerShell module.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from ..utils.converters import to_export_friendly_record, to_export_friendly_value

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Worksheet ordering mirrors Export-HashTableToExcel $defaultDesiredOrder
# ---------------------------------------------------------------------------

_DEFAULT_ORDER = [
    # Tenant Overview
    "TenantInfo", "LicenseSKUs", "GroupLicensingSummary", "LicenseOptimizationCandidates", "AdConnectConfiguration",
    # Identity
    "Users", "UserFullDetails", "Admins", "EntraIDGroups", "Domains",
    "AuthenticationMethods", "AuthenticationSSOApplications", "EnterpriseApplications", "AuthenticationConfig",
    "MfaEnrollmentSummary", "MfaMethodPostureSummary", "MfaEnforcementSummary", "MfaEnforcementGapUsers",
    "MfaEnforcementScopeReview", "ConditionalAccessPolicySummary", "ConditionalAccessOptimization",
    "ConditionalAccessPolicies", "SecurityDefaultsPolicy", "GuestSignInSummary", "GuestAccessConfiguration",
    "ExternalIdentityRestrictions", "PrivilegedAccessSummary", "PrivilegedAccessRemediationSummary",
    # Exchange
    "HybridConfiguration",
    "AllRecipients", "AllMailboxes", "PrimaryMailboxStats", "MailboxUsageDetails", "MailboxFullDetails",
    "MailboxCalendarDelegatePermissions", "NonUserMailboxes", "SharedMailboxes", "EquipmentMailboxes",
    "RoomMailboxes", "ArchiveMailboxes", "ArchiveMailboxStats",
    "LitigationHoldMailboxes", "InactiveMailboxes", "InactiveMailboxDetails", "AllExchangeGroups",
    "PublicFolderDetails", "PublicFolderPerms", "MailFlowRules", "MailFlowConnectors", "RemoteDomains",
    "EmailActivityTopSenders", "EmailActivityTopReceivers", "SMTPRelayConfig", "SMTPRelayServiceAccounts",
    "SpamFilteringConfig", "SharedMailboxGovernanceSummary", "ForwardingPolicySummary",
    "InboxRuleForwardingSummary", "InboxRulesExternalForwarding",
    # Collaboration
    "UnifiedGroups", "AllTeams", "TeamsGroupsCleanupCandidates", "TeamsVoiceSummary", "TeamsVoice",
    "SharePoint",
    "SharePointSharingSummary", "SharePointSiteUsage", "SharePointActivityUserDetail",
    "OneDrive", "OneDriveUsageDetails", "OneDriveActivityUserDetail",
    "TeamsActivityTopUsers", "Office365GroupsActivityTopGroups",
    "EmployeeExperienceInsightsSummary", "CollaborationActivitySummary",
    # Endpoint
    "DeviceDetails", "IntuneDevices", "DeviceManagementSummary",
    # Governance
    "SecuritySecureScore", "SecureScoreActions", "SMTPRelaySummary", "RetentionPolicies", "DlpPolicies",
    "PasswordLifecycleSummary", "ExternalSharingSummary", "ExternalSharingSiteOverrides",
    "ExternalExposureFindings", "UnmanagedObjects", "OneDriveOwnerMismatches",
    # Assessment Outputs
    "BestPractices", "BestPracticeFindings", "MigrationReadiness",
]

_DEFAULT_EXCLUDED = {
    "OwnershipGovernanceSummary", "TenantInfoSummary", "AuthenticationConfigSummary",
    "SpamFilteringSummary", "FederationSummary", "MfaRegistrationSummary", "EmailActivitySummary",
    "PrimaryMailboxStatsCollectionSummary", "UnifiedGroupMailboxStatsCollectionSummary",
    # Lookup-index variants of AllMailboxes — same records keyed by alternate identifiers
    "AllMailboxes-MailIdentity", "AllMailboxes-UserPrincipalName", "AllMailboxes-PrimarySmtpAddress",
    # Per-team enrichment raw dicts — consumed and merged into AllTeams by pipeline; not separate sheets
    "TeamOwners", "TeamMemberCounts", "PrivateChannels",
    # Raw mailbox purpose index — data surfaced through SharedMailboxes/NonUserMailboxes instead
    "MailboxPurposes",
    # HybridSyncDetails is a nested object that doesn't flatten well; key signals are in HybridConfiguration
    "HybridSyncDetails",
    # ReportSettings is a single-record tenant config; not useful as a standalone sheet
    "ReportSettings",
}

_WORKSHEET_ALIASES = {
    "Office365GroupsActivityTopGroups": "O365GroupsActivityTop",
    "EmployeeExperienceInsightsSummary": "EmployeeExpInsights",
    "MailboxCalendarDelegatePermissions": "MailboxCalendarDelegatePerms",
    "PrivilegedAccessRemediationSummary": "PrivilegedAccessRemediation",
}

_OPTIONAL_EMPTY = {
    "AuthenticationSSOApplications", "EnterpriseApplications", "SMTPRelaySummary", "TeamsVoiceSummary",
    "UnmanagedObjects", "OneDriveOwnerMismatches", "TeamsActivityTopUsers",
    "Office365GroupsActivityTopGroups", "EmployeeExperienceInsightsSummary",
    "InboxRulesExternalForwarding", "ExternalSharingSiteOverrides", "MailboxCalendarDelegatePermissions",
    "MfaEnforcementGapUsers", "MfaEnforcementScopeReview", "ConditionalAccessOptimization",
    "MfaMethodPostureSummary", "PrivilegedAccessRemediationSummary", "TeamsGroupsCleanupCandidates",
    "GroupLicensingSummary", "LicenseOptimizationCandidates",
    "MailboxUsageDetails", "SharedMailboxes", "EquipmentMailboxes", "RoomMailboxes",
    "NonUserMailboxes", "IntuneDevices",
    "SharePointSiteUsage", "OneDriveUsageDetails",
    "SharePointActivityUserDetail", "OneDriveActivityUserDetail", "TeamsVoice",
}

_HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
_HEADER_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
_BODY_FONT = Font(name="Calibri", size=10)

# --- Container worksheet shaping -------------------------------------------
# Some worksheets arrive as a Summary/Configuration record plus sibling arrays.
# Rendering those directly joins each array into a single cell, or drops it, so
# they are flattened to Section / Item / Value / Notes with one row per record.

_CONTAINER_SECTIONS = {
    "AdConnectConfiguration": "Directory Sync",
    "TeamsVoice": "Voice Summary",
    "AuthenticationConfig": "Authentication",
}

_LABEL_CANDIDATES = (
    "DisplayName", "Name", "Title", "ServiceName", "PolicyName", "ErrorType", "Category",
    "Identity", "UserPrincipalName", "Mail", "TelephoneNumber", "PhoneNumber", "Domain",
    "AppId", "Id",
)

_MFA_CATEGORIES = (
    ("Registered", "MethodCounts", "Users with this authentication method registered."),
    ("Weak", "WeakMethodCounts", "SMS, voice, and email OTP methods are phishable."),
    ("Strong", "StrongMethodCounts", "Authenticator app and comparable strong methods."),
    ("PhishingResistant", "PhishingResistantMethodCounts", "FIDO2, passkeys, and certificate-based methods."),
    ("Default", "DefaultMethodCounts", "Method currently set as the user default."),
)

_RE_LOWER_UPPER = re.compile(r"(?<=[a-z0-9])([A-Z])")
_RE_ACRONYM = re.compile(r"(?<=[A-Z])([A-Z][a-z])")


def export(
    data: dict,
    output_path: str | Path,
    policy: str = "Default",
) -> None:
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    ordered = _build_sheet_order(data, policy)

    wb = Workbook()
    wb.remove(wb.active)

    for logical_name in ordered:
        ws_name = _WORKSHEET_ALIASES.get(logical_name, logical_name)[:31]
        table_value = data.get(logical_name)

        # None means explicitly not collected (e.g. EXO-only datasets) — skip silently
        if table_value is None:
            continue

        rows = _get_rows(logical_name, table_value, data)

        if not rows:
            if logical_name not in _OPTIONAL_EMPTY:
                log.warning("No data for worksheet '%s'", logical_name)
            if logical_name == "OneDriveOwnerMismatches":
                rows = [{"DisplayName": "N/A", "SiteUrl": "N/A", "CurrentOwner": "N/A",
                         "ExpectedDefaultOwner": "N/A", "Notes": "No owner mismatches detected."}]
            else:
                continue

        ws = wb.create_sheet(title=ws_name)
        _write_rows(ws, rows)
        log.debug("Wrote sheet '%s' (%d rows)", ws_name, len(rows))

    wb.save(output_path)
    log.info("Workbook saved: %s", output_path)


# ---------------------------------------------------------------------------
# Internals
# ---------------------------------------------------------------------------

def _build_sheet_order(data: dict, policy: str) -> list[str]:
    excluded = set(_DEFAULT_EXCLUDED)
    ordered: list[str] = []
    seen: set[str] = set()

    for name in _DEFAULT_ORDER:
        if name not in excluded and name in data:
            ordered.append(name)
            seen.add(name)

    for name in sorted(data.keys()):
        if name not in seen and name not in excluded:
            ordered.append(name)

    return ordered


def _label(text: Any) -> str:
    """PascalCase to spaced words, preserving acronyms.

    SyncServices -> Sync Services, SSOApplications -> SSO Applications.
    """
    if text is None or str(text) == "":
        return ""
    spaced = _RE_LOWER_UPPER.sub(r" \1", str(text))
    spaced = _RE_ACRONYM.sub(r" \1", spaced)
    return spaced.strip()


def _display(value: Any, default: str = "") -> str:
    text = to_export_friendly_value(value)
    if text is None or str(text).strip() == "":
        return default
    return str(text).strip()


def _config_row(section: str, item: str, value: Any, notes: Any = None) -> dict:
    return {
        "Section": section,
        "Item": item,
        "Value": to_export_friendly_value(value),
        "Notes": to_export_friendly_value(notes),
    }


def _collection_rows(section: str, collection: Any) -> list[dict]:
    """One row per element, so arrays never collapse into a single cell."""
    rows: list[dict] = []
    for item in collection or []:
        if item is None:
            continue
        index = len(rows) + 1

        if not isinstance(item, (dict, list, tuple)):
            rows.append(_config_row(section, _display(item, f"Entry {index}"), "Configured"))
            continue

        fields = to_export_friendly_record(item) if isinstance(item, dict) else {}
        label = ""
        for candidate in _LABEL_CANDIDATES:
            if candidate in fields:
                label = _display(fields[candidate])
                if label:
                    break
        if not label:
            label = f"Entry {index}"

        # Remaining fields go to Notes so no collected detail is dropped.
        detail = [
            f"{_label(name)}: {_display(field_value)}"
            for name, field_value in fields.items()
            if _display(field_value) and _display(field_value) != label
        ]
        rows.append(_config_row(section, label, "Detected", "; ".join(detail)))

    if not rows:
        rows.append(_config_row(section, "None detected", 0,
                                f"No {section} records were collected in this run."))
    return rows


def _record_rows(section: str, record: Any) -> list[dict]:
    """One row per field. List-valued fields expand rather than joining into a cell."""
    rows: list[dict] = []
    fields = record if isinstance(record, dict) else to_export_friendly_record(record)
    for name, value in fields.items():
        if isinstance(value, (list, tuple)):
            rows.extend(_collection_rows(_label(name), value))
        else:
            rows.append(_config_row(section, _label(name), value))
    return rows


def _is_container(value: Any) -> bool:
    """True only when flattening actually rescues data that would otherwise collapse.

    That means a list somewhere: either a sibling array (AdConnectConfiguration,
    TeamsVoice) or one nested inside a Summary/Configuration record
    (AuthenticationConfig from a PowerShell snapshot).

    Deliberately excluded so they keep one-record-per-value rendering: uniform
    lookup tables (Users, SharePoint) and single-record summaries holding only
    scalars (CollaborationActivitySummary, SharePointSharingSummary).
    """
    if not isinstance(value, dict) or not value:
        return False

    if any(isinstance(v, (list, tuple)) for v in value.values()):
        return True

    for key in ("Summary", "Configuration"):
        nested = value.get(key)
        if isinstance(nested, dict) and any(isinstance(v, (list, tuple)) for v in nested.values()):
            return True

    return False


def _container_rows(logical_name: str, value: dict) -> list[dict]:
    default_section = _CONTAINER_SECTIONS.get(logical_name, "Summary")
    rows: list[dict] = []

    # Summary/Configuration first, then the rest alphabetically, so the sheet is
    # reproducible and reads top-down regardless of dict insertion order.
    primary = [k for k in value if k in ("Summary", "Configuration")]
    secondary = sorted(k for k in value if k not in ("Summary", "Configuration"))

    for key in primary + secondary:
        item = value[key]
        if item is None:
            continue
        if isinstance(item, (list, tuple)):
            rows.extend(_collection_rows(_label(key), item))
        elif key in ("Summary", "Configuration"):
            rows.extend(_record_rows(default_section, item))
        elif isinstance(item, dict):
            rows.extend(_record_rows(_label(key), item))
        else:
            rows.append(_config_row("General", _label(key), item))

    return rows


def _mfa_method_rows(value: Any, data: dict) -> list[dict] | None:
    """One row per authentication method.

    Handles both collectors: the Python one carries the raw count maps on
    MfaEnrollmentSummary itself, while a PowerShell snapshot pre-joins them into
    *Breakdown strings and keeps the raw maps on MfaRegistrationSummary.
    """
    summary = value
    if isinstance(summary, dict) and isinstance(summary.get("Summary"), dict):
        summary = summary["Summary"]
    if not isinstance(summary, dict):
        return None

    sources: list[dict] = [summary]
    registration = data.get("MfaRegistrationSummary")
    if isinstance(registration, dict):
        if isinstance(registration.get("Summary"), dict):
            registration = registration["Summary"]
        sources.append(registration)

    def _count_map(name: str) -> dict | None:
        for source in sources:
            candidate = source.get(name)
            if isinstance(candidate, dict) and candidate:
                return candidate
        return None

    total_users = summary.get("TotalUsers")
    try:
        total_users = float(total_users)
    except (TypeError, ValueError):
        total_users = None

    rows: list[dict] = []
    for total_name in ("TotalUsers", "RegisteredUsers", "NotRegisteredUsers", "RegistrationPercent",
                       "UsersWithWeakMethodsOnly", "UsersWithWeakDefaultMethod",
                       "UsersWithStrongMethods", "UsersWithPhishingResistantMethods"):
        if summary.get(total_name) is None:
            continue
        rows.append({
            "Category": "Totals",
            "Method": _label(total_name),
            "UserCount": summary[total_name],
            "PercentOfUsers": None,
            "Notes": "Tenant-wide MFA enrollment total.",
        })

    for category, map_name, note in _MFA_CATEGORIES:
        counts = _count_map(map_name)
        if not counts:
            continue
        for method_name, count in counts.items():
            try:
                percent = round(float(count) / total_users * 100, 1) if total_users else None
            except (TypeError, ValueError):
                percent = None
            rows.append({
                "Category": category,
                # Graph enum values (microsoftAuthenticator, fido2) stay verbatim so they
                # match what operators see in Entra.
                "Method": str(method_name),
                "UserCount": count,
                "PercentOfUsers": percent,
                "Notes": note,
            })

    return rows or None


def _get_rows(logical_name: str, value: Any, data: dict | None = None) -> list[dict]:
    data = data or {}

    if value is None:
        return []

    if logical_name == "MfaEnrollmentSummary":
        method_rows = _mfa_method_rows(value, data)
        if method_rows:
            return method_rows

    if isinstance(value, list):
        return [to_export_friendly_record(r) for r in value if r is not None]

    if isinstance(value, dict):
        vals = list(value.values())
        if not vals:
            return []
        # Containers (AdConnectConfiguration, TeamsVoice, AuthenticationConfig from a
        # PowerShell snapshot) flatten to Section/Item/Value/Notes, one row per record.
        if _is_container(value):
            container_rows = _container_rows(logical_name, value)
            if container_rows:
                return container_rows
        # All-scalar flat dict (PasswordLifecycleSummary, etc.) → single row
        if all(not isinstance(v, (dict, list)) for v in vals):
            return [to_export_friendly_record(value)]
        # All-dict lookup table (AllMailboxes, Users, ConditionalAccessPolicies, etc.) → one row per value
        if all(isinstance(v, dict) for v in vals):
            return [to_export_friendly_record(v) for v in vals]
        # All-list dict (PublicFolderPerms) → flatten sub-lists into individual rows
        if all(isinstance(v, list) for v in vals):
            rows: list[dict] = []
            for sub in vals:
                rows.extend(to_export_friendly_record(r) for r in sub if isinstance(r, dict))
            return rows
        # Mixed container — render as single flat row
        return [to_export_friendly_record(value)]

    return [{"Value": to_export_friendly_value(value)}]


def _write_rows(ws: Any, rows: list[dict]) -> None:
    if not rows:
        return

    # Union of every row's keys, not just the first row's: a later record that adds a
    # column would otherwise be silently dropped by record.get(header) below.
    headers: list[str] = []
    seen: set[str] = set()
    for record in rows:
        for key in record:
            if key not in seen:
                seen.add(key)
                headers.append(key)

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

    for row_idx, record in enumerate(rows, 2):
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=record.get(header))
            cell.font = _BODY_FONT

    for col_idx, header in enumerate(headers, 1):
        col_letter = get_column_letter(col_idx)
        max_len = max((len(str(r.get(header, "") or "")) for r in rows), default=0)
        ws.column_dimensions[col_letter].width = min(max(len(header), max_len) + 2, 60)
