"""
Excel workbook generation — Python equivalent of Export-HashTableToExcel.
Uses openpyxl instead of the ImportExcel PowerShell module.
"""

from __future__ import annotations

import logging
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
    "UnifiedGroups", "AllTeams", "TeamsGroupsCleanupCandidates", "TeamsVoiceSummary", "SharePoint",
    "SharePointSharingSummary", "SharePointSiteUsage", "OneDrive", "OneDriveUsageDetails",
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
}

_HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
_HEADER_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
_BODY_FONT = Font(name="Calibri", size=10)


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

        rows = _get_rows(logical_name, table_value)

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


def _get_rows(logical_name: str, value: Any) -> list[dict]:
    if value is None:
        return []
    if isinstance(value, list):
        return [to_export_friendly_record(r) for r in value if r is not None]
    if isinstance(value, dict):
        vals = list(value.values())
        if not vals:
            return []
        # All-scalar flat dict (MfaEnrollmentSummary, PasswordLifecycleSummary, etc.) → single row
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
        # Container with a Summary key (AdConnectConfiguration) → Summary row + scalar siblings
        summary = value.get("Summary")
        if isinstance(summary, dict):
            record = dict(to_export_friendly_record(summary))
            for k, v in value.items():
                if k != "Summary" and not isinstance(v, (dict, list)):
                    record[str(k)] = to_export_friendly_value(v)
            return [record]
        # Mixed container — render as single flat row
        return [to_export_friendly_record(value)]
    return [{"Value": to_export_friendly_value(value)}]


def _write_rows(ws: Any, rows: list[dict]) -> None:
    if not rows:
        return

    headers = list(rows[0].keys())
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
