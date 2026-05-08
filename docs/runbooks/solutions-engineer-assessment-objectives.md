# Solutions Engineer Assessment Objectives

This runbook explains what the `SolutionsEngineer` output profile is intended to check and how those checks map back to collected evidence. The machine-readable contract lives at:

`src/config/baseline/solutions-engineer-assessment-objectives.json`

The profile is assessment based. It is meant to support professional-services conversations about security gaps, operational risk, governance maturity, and areas for improvement. This milestone does not change generated report output. It documents the current evidence-backed behavior and adds tests so the repo can answer: what does this assessment check, why does it matter, where is the evidence, and how do we know the evidence path exists?

## Profile Contract

- Output profile: `SolutionsEngineer`
- Reporting mode: `Operator`
- Current output behavior: unchanged
- Primary evidence outputs: workbook, JSON snapshot, improvement plan, customer report, roadmap, engineer pack
- Primary interpretation rule: if evidence is missing or partial, call it a coverage or confidence gap instead of forcing the data to fit the expected story

## Objective Summary

| Objective | Area | What it checks | Evidence examples | Finding examples |
| --- | --- | --- | --- | --- |
| `SE-TENANT-001` | Tenant baseline and collection quality | Tenant identity, domains, licensing footprint, sync posture, Secure Score, and collection completeness | `TenantInfo`, `Domains`, `LicenseSKUs`, `AdConnectConfiguration`, `SecuritySecureScore` | `DOMAIN-001`, `LIC-001`, `SEC-001` |
| `SE-ID-001` | Identity and privileged access | Users, admins, groups, guests, security defaults, privileged access, password lifecycle, emergency access | `Users`, `Admins`, `EntraIDGroups`, `GuestSignInSummary`, `PrivilegedAccessSummary` | `ADMIN-001`, `ADMIN-002`, `ID-001`, `ID-005`, `SEC-004` |
| `SE-MFA-001` | MFA and authentication methods | MFA registration, method quality, enforcement coverage, admin MFA gaps, policy scope gaps | `MfaEnrollmentSummary`, `MfaRegistrationDetails`, `MfaEnforcementGapUsers`, `AdminMfaSummary` | `MFA-001`, `MFA-002`, `MFA-003`, `ADMIN-003` |
| `SE-CA-001` | Conditional Access quality | Baseline coverage for admins, legacy auth, guests, risk, devices, policy state, exclusions, overlap | `ConditionalAccessPolicies`, `ConditionalAccessPolicySummary`, `ConditionalAccessOptimization` | `CA-001`, `CA-003`, `CA-012`, `CA-014`, `CA-019` |
| `SE-APP-001` | Enterprise application governance | Consent posture, app permissions, app ownership, redirect URIs, inactive apps, SSO inventory | `EnterpriseApplications`, `AuthenticationSSOApplications`, `AuthenticationConfig` | `ID-007`, `ID-009`, `ID-011`, `SEC-002`, `SEC-003` |
| `SE-EX-001` | Exchange and mail hygiene | Mailboxes, forwarding, connectors, remote domains, SMTP relay, public folders, shared mailbox governance | `AllMailboxes`, `MailFlowConnectors`, `InboxRulesExternalForwarding`, `SMTPRelaySummary` | `EX-001`, `EX-002`, `EX-003`, `EX-006`, `EX-007` |
| `SE-COL-001` | SharePoint, OneDrive, and external sharing | Site inventory, OneDrive inventory, storage, stale content, tenant sharing, overrides, external exposure | `SharePoint`, `OneDrive`, `SharePointSharingSummary`, `ExternalExposureFindings` | `COL-001`, `COL-005`, `COL-006`, `COL-007`, `CA-013` |
| `SE-TM-001` | Teams and Microsoft 365 Groups governance | Teams, unified groups, owners, guests, dormant spaces, channel growth, cleanup candidates | `AllTeams`, `UnifiedGroups`, `TeamsGroupsCleanupCandidates`, `CollaborationActivitySummary` | `TM-001`, `TM-005`, `TM-007`, `TM-009` |
| `SE-LIC-001` | Licensing optimization | SKU utilization, dormant licensed users, group-based licensing, duplicate assignments, assignment errors | `LicenseSKUs`, `GroupLicensingSummary`, `LicenseOptimizationCandidates` | `LIC-001`, `LIC-003`, `LIC-005`, `LIC-007` |
| `SE-DEV-001` | Endpoint and device posture | Device inventory, stale devices, compliance, management, enrollment, unsupported operating systems | `DeviceDetails`, `DeviceManagementSummary` | `DEV-001`, `DEV-003`, `DEV-005`, `DEV-006` |
| `SE-GOV-001` | Purview, retention, DLP, and governance summaries | Retention policy visibility, DLP policy visibility, Secure Score, sharing, exposure, readiness summaries | `RetentionPolicies`, `DlpPolicies`, `SecuritySecureScore`, `ExternalSharingSummary` | `SEC-001`, `COL-005`, `COL-007` |

## Examples Of Gaps This Profile Looks For

- Identity and admin risk: too many Global Administrators, stale privileged accounts, missing emergency access accounts, inactive guests, weak passwordless readiness.
- MFA risk: users not registered for MFA, weak registered methods, admins not registered or not covered by active enforcement, report-only enforcement policies.
- Conditional Access quality: no enabled baseline, missing legacy-auth block, missing privileged-role or guest baseline, broad exclusions, overlapping policies.
- App governance: high-privilege enterprise apps, third-party apps with application permissions, missing owners, risky redirect URI patterns, inactive apps.
- Exchange hygiene: mailbox forwarding, inbox-rule forwarding, relay-like connectors, SMTP client auth or relay dependencies, public folders, shared mailbox ownership issues.
- Collaboration governance: external sharing posture, anonymous link defaults, externally exposed stale content, OneDrive owner mismatches, ownerless or dormant collaboration objects.
- Teams and groups governance: ownerless Teams or groups, dormant spaces, guest-heavy Teams, private/shared channel sprawl, cleanup candidates.
- Licensing optimization: high utilization SKUs, inactive licensed users, ownerless license groups, duplicate assignment patterns, license assignment errors.
- Endpoint posture: stale, unmanaged, non-compliant, or unsupported devices that may weaken secure access readiness.
- Purview and governance: whether retention and DLP policy evidence was collected and whether missing Purview evidence should lower confidence in compliance maturity claims.

## Confidence And Coverage Notes

The profile is designed to be a professional-services assessment, not a proof engine. It collects tenant evidence and derives useful findings, but some conclusions are intentionally framed as review targets.

- Conditional Access findings can be heuristic. The profile reviews policy state, names, scopes, summaries, exclusions, and derived optimization rows, but it does not simulate every effective access path.
- MFA enforcement is kept separate from MFA registration. Registration means readiness; enforcement means a policy or baseline is actively requiring MFA for the intended population.
- Exchange forwarding and relay findings identify conditions that need approval or removal. They do not know the customer's approved exception list.
- SharePoint and OneDrive external exposure findings identify tenant settings, overrides, stale content, and ownership drift. They do not prove every file-level sharing link.
- Licensing findings are practical governance signals, not full per-service usage optimization.
- Purview evidence currently covers retention and DLP policy inventory. It does not prove label deployment, eDiscovery maturity, insider risk, endpoint DLP activity, or policy effectiveness.
- Endpoint findings use device inventory and derived summaries. They are not a complete Intune, Defender, or compliance-policy audit.

## Coverage Gaps To Keep Visible

- Live integration proof: unit tests validate the matrix against source code, datasets, and rule IDs, but they do not prove collection against a live tenant. A validation tenant run is the next layer.
- Heuristic confidence: several rules use thresholds, names, summaries, and inventory patterns. Engineers should validate the most consequential findings with customer context.
- Purview depth: retention and DLP inventory is useful, but deeper compliance evidence requires explicit scope expansion.
- Endpoint depth: device summaries support posture triage, not full endpoint-management attestation.
- Application attestation: the assessment identifies app-review candidates; it does not prove owner approval or business justification.

## Maintainer Notes

When adding or changing a Solutions Engineer objective:

1. Add or update the row in `solutions-engineer-assessment-objectives.json`.
2. Link the objective to at least one real dataset key collected, exported, or consumed by the reporting/improvement pipeline.
3. Use existing rule IDs when there is a known finding. If no rule exists, leave `RelatedFindings` empty and explain the limitation in `ConfidenceNotes`.
4. Keep aspirational goals out of the matrix unless they are explicitly marked as coverage gaps.
5. Run the Solutions Engineer objective tests before treating the documentation as current.
