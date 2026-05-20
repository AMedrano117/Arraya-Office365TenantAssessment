"""
Solutions-Engineer evidence-coverage report.
Walks the snapshot and checks each SE assessment objective against its
expected datasets, producing *-SolutionsEngineerEvidenceCoverage.json.
"""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Objective matrix (mirrors solutions-engineer-assessment-objectives.json)
# Each dataset entry: (Name, snapshot_path, expected_empty)
# snapshot_path: "$.Section.Key" or "$.Data.Section.Key" or "$.Derived.Key"
# ---------------------------------------------------------------------------

_OBJECTIVES: list[dict[str, Any]] = [
    {
        "ObjectiveId": "SE-TENANT-001",
        "Area": "Tenant baseline and collection quality",
        "Goal": (
            "Establish the tenant identity, accepted-domain posture, licensing footprint, "
            "sync posture, Secure Score signal, and whether the run surfaced enough evidence "
            "to support an operator-grade assessment."
        ),
        "Datasets": [
            ("TenantInfo",              "$.Data.Tenant.TenantInfo",              False),
            ("TenantInfoSummary",       "$.Derived.TenantInfoSummary",            False),
            ("Domains",                 "$.Data.Tenant.Domains",                  False),
            ("LicenseSKUs",             "$.Data.Identity.LicenseSKUs",            False),
            ("AdConnectConfiguration",  "$.Data.Tenant.AdConnectConfiguration",   False),
            ("SecuritySecureScore",     "$.Data.Security.SecuritySecureScore",    False),
        ],
        "ChecksFor": [
            "Unverified accepted domains or domain hygiene issues",
            "Directory synchronization posture that affects identity authority and password lifecycle planning",
            "License SKU inventory and utilization pressure",
            "Secure Score posture that may indicate broad baseline control gaps",
            "Missing or partial baseline evidence that should lower confidence in downstream findings",
        ],
        "ConfidenceNotes": (
            "This objective depends on successful Graph, Exchange, and tenant baseline collection. "
            "Missing tenant or domain data should be treated as a collection gap, not as evidence "
            "that no issue exists."
        ),
        "RelatedFindings": ["DOMAIN-001", "LIC-001", "SEC-001"],
    },
    {
        "ObjectiveId": "SE-ID-001",
        "Area": "Identity and privileged access",
        "Goal": (
            "Assess the health of user, guest, group, directory, privileged-admin, and "
            "authentication baseline signals that affect security operations and tenant administration."
        ),
        "Datasets": [
            ("Users",                              "$.Data.Identity.Users",                              False),
            ("Admins",                             "$.Data.Identity.Admins",                             False),
            ("EntraIDGroups",                      "$.Data.Identity.EntraIDGroups",                      False),
            ("AuthenticationConfig",               "$.Data.Identity.AuthenticationConfig",               False),
            ("AuthenticationConfigSummary",        "$.Derived.AuthenticationConfigSummary",              False),
            ("SecurityDefaultsPolicy",             "$.Data.Identity.SecurityDefaultsPolicy",             False),
            ("GuestSignInSummary",                 "$.Data.Identity.GuestSignInSummary",                 False),
            ("PrivilegedAccessSummary",            "$.Data.Identity.PrivilegedAccessSummary",            False),
            ("PrivilegedAccessRemediationSummary", "$.Data.Identity.PrivilegedAccessRemediationSummary", False),
            ("PasswordLifecycleSummary",           "$.Data.Governance.PasswordLifecycleSummary",         False),
        ],
        "ChecksFor": [
            "Excessive standing Global Administrator assignment",
            "Missing cloud-only emergency access coverage",
            "Stale privileged accounts or inactive guests",
            "Guest invitation and external identity trust posture that needs review",
            "Security Defaults and Conditional Access overlap",
            "Passwordless readiness and password lifecycle signals",
        ],
        "ConfidenceNotes": (
            "Privileged and guest lifecycle findings rely on collected user/admin rows plus summary "
            "derivation. Sign-in recency and emergency access interpretation should be reviewed by "
            "an engineer before being treated as final operating-state proof."
        ),
        "RelatedFindings": [
            "ADMIN-001", "ADMIN-002", "ADMIN-003",
            "ID-001", "ID-002", "ID-003", "ID-004", "ID-005", "ID-006", "ID-008", "SEC-004",
        ],
    },
    {
        "ObjectiveId": "SE-MFA-001",
        "Area": "MFA and authentication methods",
        "Goal": (
            "Separate MFA registration readiness, method strength, and actual Conditional Access "
            "enforcement so the assessment does not confuse enrollment with protection."
        ),
        "Datasets": [
            ("MfaRegistrationDetails",  "$.Data.Identity.MfaRegistrationDetails",  False),
            ("MfaEnrollmentSummary",    "$.Data.Identity.MfaEnrollmentSummary",     False),
            ("MfaMethodPostureSummary", "$.Data.Identity.MfaMethodPostureSummary",  False),
            ("MfaEnforcementSummary",   "$.Data.Identity.MfaEnforcementSummary",    False),
            ("MfaEnforcementGapUsers",  "$.Data.Identity.MfaEnforcementGapUsers",   False),
            ("MfaEnforcementScopeReview",   "$.Data.Identity.MfaEnforcementScopeReview",   False),
            ("AdminMfaSummary",             "$.Data.Other.AdminMfaSummary",               False),
            ("AdminMfaRegistrationGaps",    "$.Data.Other.AdminMfaRegistrationGaps",       False),
            ("AdminMfaEnforcementGaps",     "$.Data.Other.AdminMfaEnforcementGaps",        False),
        ],
        "ChecksFor": [
            "Low MFA enrollment or incomplete registration",
            "Weak MFA method reliance such as SMS, voice, or email-based methods",
            "Registered users who are not clearly covered by enabled MFA enforcement",
            "Admin accounts missing MFA registration or active enforcement coverage",
            "Guest or member populations outside enabled MFA include scope",
        ],
        "ConfidenceNotes": (
            "MFA coverage is derived from policy scope and user inventory. Report-only policies do "
            "not count as enforcement. Complex group targeting, dynamic groups, exclusions, and "
            "emergency access exceptions still require engineer validation."
        ),
        "RelatedFindings": ["MFA-001", "MFA-002", "MFA-003", "ADMIN-002", "ADMIN-003"],
    },
    {
        "ObjectiveId": "SE-CA-001",
        "Area": "Conditional Access quality",
        "Goal": (
            "Review whether Conditional Access has a usable baseline for admins, legacy "
            "authentication, guests, risk-based access, device requirements, policy state, "
            "exclusions, and overlap."
        ),
        "Datasets": [
            ("ConditionalAccessPolicies",       "$.Data.Identity.ConditionalAccessPolicies",       False),
            ("ConditionalAccessPolicySummary",  "$.Data.Identity.ConditionalAccessPolicySummary",  False),
            ("ConditionalAccessOptimization",   "$.Data.Identity.ConditionalAccessOptimization",   False),
            ("SecurityDefaultsPolicy",          "$.Data.Identity.SecurityDefaultsPolicy",          False),
            ("MfaEnforcementSummary",           "$.Data.Identity.MfaEnforcementSummary",           False),
            ("MfaEnforcementScopeReview",       "$.Data.Identity.MfaEnforcementScopeReview",       False),
        ],
        "ChecksFor": [
            "No Conditional Access policies or no enabled policies",
            "Report-only policies that need staged promotion",
            "Missing legacy authentication blocking baseline",
            "Missing privileged-role, guest, risk-based, or compliant-device scenario coverage",
            "Broad exclusions and possible policy overlap",
            "External sharing posture without guest Conditional Access coverage",
        ],
        "ConfidenceNotes": (
            "Some Conditional Access quality signals are heuristic and summary-based. The profile "
            "can identify likely gaps and review targets, but it does not simulate every effective "
            "access path or prove runtime policy impact."
        ),
        "RelatedFindings": [
            "CA-001", "CA-002", "CA-003", "CA-004", "CA-005", "CA-006", "CA-007",
            "CA-008", "CA-009", "CA-010", "CA-011", "CA-012", "CA-013", "CA-014",
            "CA-015", "CA-016", "CA-017", "CA-018", "CA-019",
        ],
    },
    {
        "ObjectiveId": "SE-APP-001",
        "Area": "Enterprise application governance",
        "Goal": (
            "Surface app consent, enterprise application, SSO, ownership, permission, redirect URI, "
            "and inactivity signals that affect tenant application governance."
        ),
        "Datasets": [
            ("EnterpriseApplications",       "$.Data.Identity.EnterpriseApplications",       False),
            ("AuthenticationSSOApplications","$.Data.Identity.AuthenticationSSOApplications", True),
            ("AuthenticationConfig",         "$.Data.Identity.AuthenticationConfig",          False),
            ("AuthenticationConfigSummary",  "$.Derived.AuthenticationConfigSummary",         False),
        ],
        "ChecksFor": [
            "High-privilege enterprise applications",
            "Third-party applications with application permissions",
            "First-party applications without clear owner signals",
            "Redirect URI patterns that warrant review",
            "Inactive enterprise applications",
            "Admin consent workflow or permission grant policy posture that may be too broad",
        ],
        "ConfidenceNotes": (
            "The assessment highlights app-governance risk candidates from inventory and summary "
            "fields. It does not replace a full access review, owner attestation, or vendor/business "
            "justification workflow."
        ),
        "RelatedFindings": ["ID-007", "ID-009", "ID-010", "ID-011", "ID-012", "SEC-002", "SEC-003"],
    },
    {
        "ObjectiveId": "SE-EX-001",
        "Area": "Exchange and mail hygiene",
        "Goal": (
            "Assess Exchange mailbox, recipient, forwarding, connector, remote-domain, SMTP relay, "
            "public folder, shared mailbox, activity, and mail hygiene signals that affect security "
            "and operations."
        ),
        "Datasets": [
            ("AllRecipients",                    "$.Data.Exchange.AllRecipients",                    False),
            ("AllMailboxes",                     "$.Data.Exchange.AllMailboxes",                     False),
            ("AllExchangeGroups",                "$.Data.Exchange.AllExchangeGroups",                False),
            ("HybridConfiguration",              "$.Data.Tenant.HybridConfiguration",                False),
            ("MailFlowRules",                    "$.Data.Exchange.MailFlowRules",                    False),
            ("MailFlowConnectors",               "$.Data.Exchange.MailFlowConnectors",               False),
            ("RemoteDomains",                    "$.Data.Exchange.RemoteDomains",                    False),
            ("SpamFilteringConfig",              "$.Data.Security.SpamFilteringConfig",              False),
            ("SMTPRelaySummary",                 "$.Data.Security.SMTPRelaySummary",                 False),
            ("ForwardingPolicySummary",          "$.Data.Exchange.ForwardingPolicySummary",          False),
            ("InboxRulesExternalForwarding",     "$.Data.Exchange.InboxRulesExternalForwarding",     True),
            ("SharedMailboxGovernanceSummary",   "$.Data.Exchange.SharedMailboxGovernanceSummary",   False),
            ("PublicFolderDetails",              "$.Data.Exchange.PublicFolderDetails",              False),
            ("EmailActivitySummary",             "$.Data.Exchange.EmailActivitySummary",             False),
            ("EmailActivityTopSenders",          "$.Data.Exchange.EmailActivityTopSenders",          False),
            ("EmailActivityTopReceivers",        "$.Data.Exchange.EmailActivityTopReceivers",        False),
        ],
        "ChecksFor": [
            "Mailbox or inbox-rule forwarding that needs approval or removal",
            "Higher-trust or relay-like mail connectors",
            "SMTP client authentication or relay dependency",
            "Public folders that need migration or retirement planning",
            "Shared mailbox ownership or oversized-growth governance gaps",
            "Mail activity concentration and recipient-domain patterns that support operational review",
        ],
        "ConfidenceNotes": (
            "Exchange findings depend on successful Exchange Online session access and the available "
            "mailbox/rule scope. External forwarding evidence should be reviewed against approved "
            "business exceptions and outbound policy context."
        ),
        "RelatedFindings": ["EX-001", "EX-002", "EX-003", "EX-004", "EX-005", "EX-006", "EX-007"],
    },
    {
        "ObjectiveId": "SE-COL-001",
        "Area": "SharePoint, OneDrive, and external sharing",
        "Goal": (
            "Assess SharePoint and OneDrive site inventory, storage, stale content, ownership, "
            "tenant sharing settings, site-level overrides, and externally exposed collaboration "
            "assets."
        ),
        "Datasets": [
            ("SharePoint",                  "$.Data.Collaboration.SharePoint",                  False),
            ("OneDrive",                    "$.Data.Collaboration.OneDrive",                    False),
            ("SharePointSharingSummary",    "$.Data.Collaboration.SharePointSharingSummary",    False),
            ("ExternalSharingSummary",      "$.Data.Tenant.ExternalSharingSummary",             False),
            ("ExternalSharingSiteOverrides","$.Data.Tenant.ExternalSharingSiteOverrides",       True),
            ("ExternalExposureFindings",    "$.Data.Tenant.ExternalExposureFindings",           False),
            ("OwnershipGovernanceSummary",  "$.Derived.OwnershipGovernanceSummary",             False),
            ("OneDriveOwnerMismatches",     "$.Derived.OneDriveOwnerMismatches",                True),
            ("CollaborationActivitySummary","$.Data.Collaboration.CollaborationActivitySummary",False),
        ],
        "ChecksFor": [
            "Tenant-level external sharing posture and anonymous link defaults",
            "Site-level external sharing overrides",
            "Externally exposed stale sites or OneDrive locations",
            "OneDrive ownership mismatches",
            "Ownerless or unmanaged collaboration objects",
            "Large or inactive collaboration storage locations",
        ],
        "ConfidenceNotes": (
            "SharePoint/OneDrive inventory may use Graph with SharePoint Online fallback depending "
            "on access. Sharing and stale-content findings indicate review targets; they do not prove "
            "that every individual file or link is externally accessible."
        ),
        "RelatedFindings": [
            "COL-001", "COL-002", "COL-003", "COL-004", "COL-005", "COL-006", "COL-007", "CA-013",
        ],
    },
    {
        "ObjectiveId": "SE-TM-001",
        "Area": "Teams and Microsoft 365 Groups governance",
        "Goal": (
            "Review Teams, Microsoft 365 Groups, unified groups, guest-heavy collaboration, "
            "private/shared channel growth, owner health, activity, and cleanup candidate signals."
        ),
        "Datasets": [
            ("AllTeams",                    "$.Data.Collaboration.AllTeams",                    False),
            ("TeamsVoiceSummary",           "$.Data.Collaboration.TeamsVoiceSummary",           False),
            ("UnifiedGroups",               "$.Data.Collaboration.UnifiedGroups",               False),
            ("EntraIDGroups",               "$.Data.Identity.EntraIDGroups",                   False),
            ("TeamsGroupsCleanupCandidates","$.Data.Collaboration.TeamsGroupsCleanupCandidates",False),
            ("ExternalExposureFindings",    "$.Data.Tenant.ExternalExposureFindings",           False),
            ("CollaborationActivitySummary","$.Data.Collaboration.CollaborationActivitySummary",False),
            ("OwnershipGovernanceSummary",  "$.Derived.OwnershipGovernanceSummary",             False),
        ],
        "ChecksFor": [
            "Ownerless Teams or Microsoft 365 groups",
            "Dormant Teams or groups",
            "Guest-heavy Teams or externally relevant group exposure",
            "Private and shared channel sprawl",
            "No-member, ownerless, or high-channel cleanup candidates",
            "Group-based governance dependencies such as license-managing groups",
        ],
        "ConfidenceNotes": (
            "Teams and group governance findings are inventory and lifecycle signals. They should "
            "drive owner/lifecycle review, not automatic deletion or access changes without business "
            "validation."
        ),
        "RelatedFindings": [
            "TM-001", "TM-002", "TM-003", "TM-004", "TM-005", "TM-006", "TM-007", "TM-008",
            "TM-009", "COL-001",
        ],
    },
    {
        "ObjectiveId": "SE-LIC-001",
        "Area": "Licensing optimization",
        "Goal": (
            "Assess license inventory, utilization pressure, duplicate or inefficient assignment "
            "patterns, group-based licensing ownership, and license assignment errors."
        ),
        "Datasets": [
            ("LicenseSKUs",                  "$.Data.Identity.LicenseSKUs",                  False),
            ("Users",                        "$.Data.Identity.Users",                        False),
            ("EntraIDGroups",               "$.Data.Identity.EntraIDGroups",                False),
            ("GroupLicensingSummary",        "$.Data.Identity.GroupLicensingSummary",        False),
            ("LicenseOptimizationCandidates","$.Data.Identity.LicenseOptimizationCandidates",False),
        ],
        "ChecksFor": [
            "SKUs near or at capacity",
            "Inactive or disabled licensed users",
            "Group-based licensing without owner accountability",
            "Ownerless license-managing groups",
            "Direct plus group duplicate assignment for the same SKU",
            "Likely duplicate suite assignments",
            "Graph license assignment errors",
        ],
        "ConfidenceNotes": (
            "This is a practical licensing-governance review, not a deep per-service usage analytics "
            "model. Optimization candidates should be validated against role requirements before "
            "reclaiming licenses."
        ),
        "RelatedFindings": [
            "LIC-001", "LIC-002", "LIC-003", "LIC-004", "LIC-005", "LIC-006", "LIC-007",
        ],
    },
    {
        "ObjectiveId": "SE-DEV-001",
        "Area": "Endpoint and device posture",
        "Goal": (
            "Assess device inventory, stale devices, compliance state, management state, enrollment "
            "posture, and unsupported operating-system signals that affect secure access readiness."
        ),
        "Datasets": [
            ("DeviceDetails",         "$.Data.Identity.DeviceDetails",         False),
            ("DeviceManagementSummary","$.Data.Identity.DeviceManagementSummary",False),
        ],
        "ChecksFor": [
            "Stale device rate above threshold",
            "Non-compliant devices",
            "Unmanaged devices",
            "Limited Intune enrollment coverage",
            "Unsupported or older operating systems",
            "Device-management posture that should influence Conditional Access planning",
        ],
        "ConfidenceNotes": (
            "Endpoint findings are based on directory/device inventory and summaries. They do not "
            "replace a complete Intune policy, compliance policy, Defender, or device-risk assessment."
        ),
        "RelatedFindings": [
            "DEV-001", "DEV-002", "DEV-003", "DEV-004", "DEV-005", "DEV-006", "CA-018",
        ],
    },
    {
        "ObjectiveId": "SE-GOV-001",
        "Area": "Purview, retention, DLP, and governance summaries",
        "Goal": (
            "Surface Purview retention and DLP policy visibility, plus broader governance summary "
            "datasets that should shape roadmap discussions without implying a full compliance-program "
            "audit."
        ),
        "Datasets": [
            ("RetentionPolicies",        "$.Data.Governance.RetentionPolicies",     False),
            ("DlpPolicies",              "$.Data.Governance.DlpPolicies",           False),
            ("SecuritySecureScore",      "$.Data.Security.SecuritySecureScore",     False),
            ("BestPracticeFindings",     "$.Derived.BestPracticeFindings",          False),
            ("MigrationReadiness",       "$.Derived.MigrationReadiness",            False),
            ("ExternalSharingSummary",   "$.Data.Tenant.ExternalSharingSummary",    False),
            ("ExternalExposureFindings", "$.Data.Tenant.ExternalExposureFindings",  False),
        ],
        "ChecksFor": [
            "Whether retention policy evidence was surfaced",
            "Whether DLP policy evidence was surfaced",
            "Whether Purview data was missing and should lower confidence in compliance maturity statements",
            "Governance roadmap signals across security, sharing, lifecycle, and readiness summaries",
        ],
        "ConfidenceNotes": (
            "Purview collection currently surfaces retention and DLP policy inventory. It does not "
            "validate every label, retention outcome, eDiscovery case, insider-risk capability, "
            "endpoint DLP event, or policy effectiveness result."
        ),
        "RelatedFindings": ["SEC-001", "COL-005", "COL-006", "COL-007"],
    },
]

_COVERAGE_GAPS = [
    {
        "GapId": "SE-GAP-001",
        "Area": "Live integration proof",
        "CurrentLimitation": (
            "This milestone validates the matrix against source code, exported dataset names, and "
            "rule IDs. It does not prove collection success against a live tenant."
        ),
        "Impact": (
            "A source-backed objective can still be partially skipped at runtime because of "
            "permissions, service availability, tenant licensing, or auth-session limitations."
        ),
        "NextValidationLayer": (
            "Run the SolutionsEngineer profile against a controlled validation tenant and assert "
            "populated evidence counts for each objective family."
        ),
    },
    {
        "GapId": "SE-GAP-002",
        "Area": "Heuristic confidence",
        "CurrentLimitation": (
            "Several findings use policy names, summaries, thresholds, and inventory patterns to "
            "flag likely gaps."
        ),
        "Impact": (
            "The profile is useful for professional-services assessment triage, but some findings "
            "require engineer review before being treated as definitive control failure."
        ),
        "NextValidationLayer": (
            "Add tenant-fixture or recorded-snapshot tests that exercise known positive and negative "
            "cases for Conditional Access, MFA scope, app governance, Exchange forwarding, and "
            "sharing posture."
        ),
    },
    {
        "GapId": "SE-GAP-003",
        "Area": "Purview depth",
        "CurrentLimitation": (
            "Purview evidence is limited to retention and DLP policy inventory plus report summaries."
        ),
        "Impact": (
            "The assessment can report whether policy evidence surfaced, but it cannot certify "
            "Purview program maturity or policy effectiveness."
        ),
        "NextValidationLayer": (
            "Add deeper compliance collector coverage only after the required permissions, tenant "
            "licensing, and evidence model are documented."
        ),
    },
    {
        "GapId": "SE-GAP-004",
        "Area": "Endpoint depth",
        "CurrentLimitation": (
            "Endpoint posture uses device inventory and derived summaries rather than a complete "
            "Intune, Defender, or device-risk policy audit."
        ),
        "Impact": (
            "Device findings should be used to prioritize endpoint discovery and access-readiness "
            "work, not as complete endpoint-management attestation."
        ),
        "NextValidationLayer": (
            "Add explicit Intune policy and compliance-state collector tests if the assessment scope "
            "expands into endpoint-management design validation."
        ),
    },
    {
        "GapId": "SE-GAP-005",
        "Area": "Application attestation",
        "CurrentLimitation": (
            "Enterprise application findings identify review candidates from inventory, permissions, "
            "owner signals, redirect URIs, and activity signals."
        ),
        "Impact": (
            "The profile does not prove business ownership, vendor approval, or approved exception "
            "status for each app."
        ),
        "NextValidationLayer": (
            "Introduce an app attestation workflow or customer-provided ownership evidence if the "
            "deliverable needs governance sign-off rather than assessment triage."
        ),
    },
]


# ---------------------------------------------------------------------------
# Path resolution
# ---------------------------------------------------------------------------

def _resolve_path(snapshot: dict, path: str) -> Any:
    """Resolve a JSONPath-style path (e.g. '$.Data.Identity.Users') in the snapshot dict."""
    parts = path.lstrip("$.").split(".")
    node: Any = snapshot
    for part in parts:
        if not isinstance(node, dict):
            return None
        node = node.get(part)
        if node is None:
            return None
    return node


def _count_records(value: Any) -> int:
    if value is None:
        return 0
    if isinstance(value, list):
        return len(value)
    if isinstance(value, dict):
        return 1 if value else 0
    return 1


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate(
    snapshot: dict,
    support_dir: Path,
    tenant_name: str = "",
    snapshot_path: str | None = None,
) -> dict:
    """Build *-SolutionsEngineerEvidenceCoverage.json and return the coverage dict."""
    support_dir = Path(support_dir)
    support_dir.mkdir(parents=True, exist_ok=True)

    objectives_out: list[dict] = []
    all_empty: list[str] = []
    all_unexpected_empty: list[str] = []
    covered_count = 0
    review_count = 0

    for obj_def in _OBJECTIVES:
        obj_id = obj_def["ObjectiveId"]
        dataset_evidence: list[dict] = []
        present_names: list[str] = []
        empty_names: list[str] = []
        unexpected_empty_names: list[str] = []
        expected_empty_names: list[str] = []
        missing_names: list[str] = []

        for ds_name, ds_path, ds_expected_empty in obj_def["Datasets"]:
            value = _resolve_path(snapshot, ds_path)
            present = value is not None
            record_count = _count_records(value)
            populated = record_count > 0

            dataset_evidence.append({
                "Name":        ds_name,
                "Present":     present,
                "Populated":   populated,
                "RecordCount": record_count,
                "Path":        ds_path,
            })

            if not present:
                missing_names.append(ds_name)
            else:
                present_names.append(ds_name)
                if not populated:
                    empty_names.append(ds_name)
                    all_empty.append(ds_name)
                    if ds_expected_empty:
                        expected_empty_names.append(ds_name)
                    else:
                        unexpected_empty_names.append(ds_name)
                        all_unexpected_empty.append(ds_name)

        # Status / confidence
        has_unexpected_empty = len(unexpected_empty_names) > 0
        confidence = "Review" if has_unexpected_empty else "Strong"
        status = "Covered"

        if confidence == "Strong":
            covered_count += 1
        else:
            review_count += 1

        objectives_out.append({
            "ObjectiveId":              obj_id,
            "Area":                     obj_def["Area"],
            "Goal":                     obj_def["Goal"],
            "Status":                   status,
            "Confidence":               confidence,
            "ExpectedDatasetCount":     len(obj_def["Datasets"]),
            "PresentDatasetCount":      len(present_names),
            "PopulatedDatasetCount":    len([d for d in dataset_evidence if d["Populated"]]),
            "MissingDatasets":          missing_names,
            "EmptyDatasets":            empty_names,
            "ExpectedEmptyDatasets":    expected_empty_names,
            "UnexpectedEmptyDatasets":  unexpected_empty_names,
            "PresentDatasets":          present_names,
            "DatasetEvidence":          dataset_evidence,
            "RelatedFindings":          obj_def.get("RelatedFindings", []),
            "ChecksFor":                obj_def["ChecksFor"],
            "ConfidenceNotes":          obj_def["ConfidenceNotes"],
        })

    # Deduplicate global lists while preserving first-seen order
    seen: set[str] = set()
    unique_empty: list[str] = []
    for n in all_empty:
        if n not in seen:
            seen.add(n)
            unique_empty.append(n)

    seen = set()
    unique_unexpected: list[str] = []
    for n in all_unexpected_empty:
        if n not in seen:
            seen.add(n)
            unique_unexpected.append(n)

    coverage: dict[str, Any] = {
        "SchemaVersion":                1,
        "GeneratedAtUtc":               datetime.now(timezone.utc).isoformat(),
        "OutputProfile":                "SolutionsEngineer",
        "ReportingMode":                "Operator",
        "SnapshotPath":                 snapshot_path or "",
        "ObjectiveCount":               len(_OBJECTIVES),
        "CoveredObjectiveCount":        covered_count,
        "ReviewObjectiveCount":         review_count,
        "PartialObjectiveCount":        0,
        "PresentEmptyObjectiveCount":   0,
        "MissingObjectiveCount":        0,
        "MissingOrPartialObjectiveCount": 0,
        "MissingDatasetCount":          sum(len(o["MissingDatasets"]) for o in objectives_out),
        "EmptyDatasetCount":            len(unique_empty),
        "UnexpectedEmptyDatasetCount":  len(unique_unexpected),
        "MissingDatasets":              [],
        "EmptyDatasets":                unique_empty,
        "UnexpectedEmptyDatasets":      unique_unexpected,
        "StatusSummary": {
            "Covered":     covered_count,
            "Review":      review_count,
            "Partial":     0,
            "PresentEmpty":0,
            "Missing":     0,
        },
        "Objectives":    objectives_out,
        "CoverageGaps":  _COVERAGE_GAPS,
    }

    stem = tenant_name or "Tenant"
    out_path = support_dir / f"{stem}-SolutionsEngineerEvidenceCoverage.json"
    out_path.write_text(json.dumps(coverage, indent=2, default=str), encoding="utf-8")
    log.info("Evidence coverage written: %s (%d objectives)", out_path, len(_OBJECTIVES))
    return coverage
