# T2T Migration Assessment Backlog

Review date: 2026-11-10
Related documents: [coverage](T2T-Migration-Assessment-Coverage.md), [synthetic scenario spec](T2T-Synthetic-Scenario-Spec.md), [discovery questionnaire](T2T-Migration-Discovery-Questionnaire.md), [review summary](T2T-Review-Summary.md)

Status: the P0 batch is implemented on branch `feature/t2t-p0-correctness`. P0-01 to P0-05 and P0-07 are implemented in code with regression tests. P0-06 is handled as a documentation and profile-contract correction plus an "Evidence coverage" disclosure in the checklist; the profile expansion itself is P1-01. Details, semantic changes, and test coverage are in [T2T-P0-Implementation-Status.md](T2T-P0-Implementation-Status.md). P1 to P3 items are a reviewed backlog and are not implemented.

## Design principles for every item

- **Milestone-neutral.** The assessment does not define "Day 1", fixed phases, or a required cutover. Engagements differ: phased, coexistence, wave-based, single cutover, identity-first, messaging-first, collaboration-later, endpoint work during or after, or no single business date. Milestones, waves, dates, and acceptance criteria are supplied by the engagement through decisions, questionnaires, and wave plans. Use neutral terms: migration readiness, target-state readiness, migration dependency, sequencing dependency, migration blocker, migration risk, coexistence requirement, customer decision required, migration wave, post-migration requirement, and cutover-specific requirement only when the engagement defines a cutover.
- **Evidence semantics.** Unknown is not Ready. Not collected is not "not configured". A collection failure is not absence. Partial evidence is not a confirmed state. Not applicable is not "not detected". Insufficient permission is not absence. Unsupported automatic collection does not mean the customer does not use a feature. Use the shared helpers `Get-ArrayaDatasetEvidenceState`, `Get-ArrayaDirectorySyncEvidence`, `Get-ArrayaExchangeHybridEvidence`, `Get-ArrayaMailExchangerClassification`, and `Get-ArrayaMailboxHoldEvidence` instead of re-deriving these states.
- **Two distinct migration profiles.** `Presales` is for scoping, complexity, quotation inputs, and effort estimation. `TenantToTenantMigration` is for the evidence needed to design and execute a migration. They share collectors, normalized evidence, and interpretation logic, and may differ in depth, runtime, report detail, and artifacts. They are not to be merged or renamed.

## Priority definitions

| Priority | Meaning |
| --- | --- |
| P0 | Required to avoid materially incorrect migration planning. A current output is wrong or unsupported. |
| P1 | Required for migration readiness and risk identification. |
| P2 | Valuable for detailed migration design and estimation. |
| P3 | Optional usability, reporting, or future enhancement. |

## Conventions that every item must follow

- Reuse the existing patterns: collector steps through `New-ArrayaCollectorStep`; evidence states such as `Collected`, `Partial`, `Unavailable`, `NotCollected`; register rows keyed by `DecisionKey`; workbook order lists in `Export-HashTableToExcel.ps1`; AST-extracted function tests as in `tests/unit/MigrationScopeReconciliation.Tests.ps1`.
- Read-only Microsoft Graph and Exchange Online calls only.
- Never write secret values, key material, token contents, or certificate bytes to the snapshot, logs, or reports. Credential rows carry type, usage, and dates only.
- ASCII-only string literals in PowerShell and Markdown.
- New Graph permissions must be added in three places together: `Get-AssessmentGraphDelegatedScopes`, the preflight probe table in `Test-AssessmentPermissionPreflight`, and `RUN.md` plus `docs/runbooks/app-registration-setup.md`. Mark every new permission optional so existing app registrations keep working, and emit `Needs Data` with a permission note when it is absent.
- A missing permission, a failed call, and a genuinely empty result are three different states and must stay distinguishable in the snapshot.

## P0: change set (implemented; original specification kept for reference)

All seven items are narrow, have a reproduction (R ids in the coverage document), need no new permission, and add no runtime. They change values in existing columns and add rows that use existing columns. They do not change the snapshot schema or worksheet schemas. P0-06 also changes two documentation files.

The text below is the specification written during the review. Where the implementation differs, [T2T-P0-Implementation-Status.md](T2T-P0-Implementation-Status.md) is authoritative. The differences are: P0-02 adds no evidence-state field and instead treats the organization read as successful when `TenantInfo` carries a tenant id or display name; P0-06 item (c), the merged-profile warning, was not implemented and moved to P1-18; P0-07 keeps the item name "Mailboxes on hold" and never scores it Ready.

### P0-01 Complexity flags report AD Connect and Exchange hybrid as configured for every tenant

- Evidence: R1, R2. `Update-MigrationComplexityFlags` (`src/scripts/migrated/legacy/Get-FullTenantReportDetails.ps1` lines 12716 to 12728) tests whether `AdConnectConfiguration` and `HybridConfiguration` contain values. Both collectors always populate their container.
- Change type: analysis.
- Change: evaluate `AdConnectConfiguration.Summary.OnPremisesSyncEnabled` and `HybridConfiguration.ExchangeHybrid.IsHybridConfigured` (with `EvidenceCount` and `MigrationEndpointCount` as secondary signals). Three outcomes: true adds the existing `Review` row; false or a collected null adds no row; an absent summary adds a `Needs Data` row using the pattern already used for "Applications with a custom-domain application URI".
- Permissions, on-premises access, runtime, throttling: none.
- Compatibility: `MigrationQuoteReadiness` loses two false `Condition` rows for cloud-only tenants. Decision keys `FLAG:ad-connect-directory-synchronization` and `FLAG:exchange-hybrid` keep their names. Replayed snapshots are recomputed.
- Privacy and security: none.
- Tests: fail-first cases in `tests/unit/MigrationScopeReconciliation.Tests.ps1` for cloud-only, synchronized, hybrid, non-hybrid, and not-collected stores (scenarios S01, S03, N01, N02).
- Acceptance criteria: a cloud-only fixture produces neither flag; a synchronized fixture produces the AD Connect flag; an uncollected fixture produces `Needs Data`; all existing tests pass.

### P0-02 Unknown directory-synchronization and sign-in state is reported as cloud-only

- Evidence: R4, R8. `Invoke-HTMLHelperFunctions.ps1` line 3946 coerces a missing value to false and line 4056 prints "Cloud-only identity model" with status `Ready`. `Export-TenantToTenantQuestionnaireMarkdown.ps1` line 354 prints "No active on-premises directory synchronization detected." and line 353 prints "Entra ID cloud authentication appears to be the primary sign-on provider." when nothing was collected.
- Change type: analysis, report.
- Change: make the state tri-valued. True: "On-prem sync enabled", `Review`. False, or null with a successful organization read: "No directory synchronization reported by Microsoft Graph", `Info`, with a note that Microsoft Graph cannot show on-premises directories that are not synchronized. Unknown: "Directory synchronization state not collected", `Needs Data`. Record whether the organization read succeeded in `AdConnectConfiguration.Summary` as an evidence-state field. Apply the same three outcomes to the two questionnaire answers and to line 357 (SSO applications).
- Permissions, runtime: none.
- Compatibility: adds one evidence-state property to `AdConnectConfiguration.Summary`; older snapshots without it are treated as unknown only when `OnPremisesSyncEnabled` is also absent.
- Tests: `tests/unit/Reporting.Tests.ps1` or `Common.Tests.ps1` with mocked organization payloads (S02, N01); questionnaire render test.
- Acceptance criteria: no output contains "Cloud-only identity model" unless the organization read succeeded; the phrase "appears to be the primary sign-on provider" is produced only when the authentication dataset was collected.

### P0-03 "Ready" is derived from absent data and uncollected items vanish from the cutover checklist

- Evidence: R3, R5, R6. `CutoverPrepSummary` (legacy script lines 14409 to 14422), `Add-TenantToTenantChecklistRow` (line 14457), `MigrationReadiness` rows (`Invoke-HTMLHelperFunctions.ps1` lines 4054 to 4070), and the HTML fallback `New-MigrationReadinessChecklistRows` (line 7974).
- Change type: analysis, report.
- Change: add one helper that returns the evidence state of a dataset from three inputs already in the store: the matching `CollectorPlanResults` step status, whether the dataset key exists, and the collector's own evidence-state field where one exists. Use it for domains, mailboxes, mail flow connectors, remote domains, public folders, hybrid configuration, and Teams. A count of zero yields `Ready` only when the state is collected; otherwise `Needs Data` with the reason (skipped by profile, failed, permission). Keep hiding `Ready` and `Info` rows, never hide `Needs Data`.
- Permissions, runtime: none.
- Compatibility: more `Needs Data` rows in degraded runs. No column change. `tests/unit/Common.Tests.ps1` fixtures that omit `CollectorPlanResults` need an explicit collected state or must expect `Needs Data`.
- Tests: fail-first cases for an empty store, a store with collected-empty datasets, and a store with skipped steps (S01, N03, N04); render-only check of the cutover HTML.
- Acceptance criteria: R3, R5, and R6 fixtures produce no `Ready` status; a fixture with successful empty collections still produces `Ready`.

### P0-04 Multi-Geo is read from properties that do not exist

- Evidence: R7. `Get-TenantOverviewInfo` (legacy script lines 17321 to 17336) reads `IsMultipleGeolocationEnabled` and `MultiGeoConfiguration`. The Graph organization resource exposes `isMultipleDataLocationsForServicesEnabled`. The questionnaire (line 342) turns the resulting null into "No multi-geo configuration detected".
- Change type: collection, report.
- Change: read `isMultipleDataLocationsForServicesEnabled` with both casings, keep null as unknown, and change the questionnaire answer to three outcomes. Leave `MultiGeoAllowed` and `MultiGeoCentral` empty unless a supported source is added later.
- Permissions: none (`Organization.Read.All` is already required). Runtime: none.
- Compatibility: `TenantInfo.MultiGeoEnabled` becomes populated. No column change.
- Tests: mocked organization payloads with true, false, null, and missing property (S17, N05).
- Acceptance criteria: the true payload yields "Yes"; the null payload yields an unknown answer that does not contain the word "No".

### P0-05 MX classification treats DNS failure and mx.microsoft records as third-party routing

- Evidence: R9. Legacy script line 10749 computes `Office365MailExchanger` as a wildcard match on the joined MX names. A failed lookup and an `mx.microsoft` target both return false, which sets `ThirdPartySpamFilterReview = Review` and increments "domain(s) with non-Microsoft 365 MX".
- Change type: collection, analysis.
- Change: move the classification into a small function that returns true, false, or null. Null when the lookup returned nothing. True when every MX target ends in `mail.protection.outlook.com`, `mx.microsoft`, or the other documented Exchange Online suffixes. Record a per-domain DNS evidence state. Count only false as non-Microsoft 365 MX, and report null as `Needs Data`.
- Permissions: none. Runtime: none. Note: the lookups use `Resolve-DnsName` against a fixed public resolver; outbound DNS is often blocked on customer networks, which is the main trigger of the null case.
- Compatibility: `Office365MailExchanger` gains a null value; consumers that compare with `-eq $false` already behave correctly.
- Tests: unit tests for the new function (S01, N06); the existing source-pattern tests that reference the MX strings stay valid.
- Acceptance criteria: the three R9 inputs return null, true, true.

### P0-06 The TenantToTenantMigration profile promises more than it collects and never says what it skipped

- Evidence: code reading. `Resolve-AssessmentProfileCollectionPlan` (legacy script lines 652 to 673) disables eleven collectors and four derived-table builders. `Export-HashTableToExcel.ps1` (lines 1137 to 1151) excludes the matching sheets. `README.md` line 203 and `RUN.md` line 212 describe the profile as deep migration readiness with scope "All". Merged selections reset the collection and workbook policies to `Default` (runner lines 717 to 720), so `Prepare-AssessmentExportData` (line 11239) skips the tenant-to-tenant shaping.
- Change type: report, documentation. The collection change itself is P1-01.
- Change: (a) add one `Needs Data` row per skipped dataset to `MigrationReadinessChecklist`, driven by `CollectorPlanResults` rows with status `Skipped`, with the note "Not collected in this profile" and the action "Run with the profile that collects this dataset"; (b) correct the two documentation lines so the profile is described as an Exchange and collaboration cutover pack until P1-01 lands; (c) log a warning when `TenantToTenantMigration` is merged with another profile, stating that the cutover shaping is not applied.
- Permissions, runtime: none.
- Compatibility: checklist gains rows; no column change.
- Tests: checklist test with skipped collector results (S01, N04); `DocumentationLinks.Tests.ps1` continues to pass.
- Acceptance criteria: a T2T-only run lists enterprise applications, Entra groups, Microsoft 365 Groups, devices, Conditional Access, MFA registration, administrators, and Purview policies as not collected.

### P0-07 "Mailboxes on hold" counts litigation hold only

- Evidence: R10. Legacy script lines 14408 and 14413, and the same test in `Update-MigrationComplexityFlags` line 12756 (which is correctly labeled "litigation hold").
- Change type: analysis.
- Change: compute the hold signal from every field already collected: `LitigationHoldEnabled`, `DelayHoldApplied`, a non-empty `InPlaceHolds`, and report `RetentionHoldEnabled` as a separate retention-processing signal. State in the row value which hold types were evaluated and that compliance tag holds and organization-wide policies are not evaluated (P1-14). When `InPlaceHolds` was not collected (non-All modes), say so.
- Permissions, runtime: none.
- Compatibility: value text changes; the item name is a lookup key in the HTML order list (`Invoke-HTMLHelperFunctions.ps1` line 8009) and should not be renamed.
- Tests: R10 fixture as a fail-first test (S20, N07).
- Acceptance criteria: the R10 fixture yields `Review` with a count of two and names the evaluated hold types.

## P1: migration readiness and risk identification

### P1-01 Give the tenant-to-tenant profile the evidence it needs

- Files: `Resolve-AssessmentProfileCollectionPlan` and `Prepare-AssessmentExportData` in the legacy script; `Export-HashTableToExcel.ps1` (`tenantToTenantCutoverDesiredOrder`, exclusions); `Get-ArrayaAssessmentOutputProfilePolicy.ps1`; `Arraya.M365.AssessmentRunner.psm1` (merge rules); `README.md`, `RUN.md`.
- Change type: collection, report.
- Change: in the T2T plan enable `CollectAuthenticationConfiguration`, `CollectEntraGroups`, `CollectUnifiedGroups`, `CollectExchangeGroups` (already on), `CollectDevices`, `CollectConditionalAccessPolicies`, `CollectAdmins`, and `CollectGovernanceCompliancePolicies`. Add the `Migration*` registers, `EnterpriseApplications`, `AuthenticationSSOApplications`, `EntraIDGroups`, `AllExchangeGroups`, `UnifiedGroups`, `GroupWorkloadReconciliation`, `DeviceDetails`, `MailFlowRules`, `LicenseSKUs`, and `GroupLicensingSummary` to the T2T workbook order. Let a merged selection that includes `TenantToTenantMigration` keep the cutover shaping and the delegate collection depth: `Get-ArrayaCollectionDepthPolicy` (legacy script line 3114) enables mailbox and calendar delegate enumeration only for the `TenantToTenantCutover` scope policy, and that same branch disables Microsoft 365 Group mailbox statistics and full SharePoint detail, which the registers need.
- Permissions: none new.
- Runtime: the authentication step adds sign-in log queries and per-application calls; Entra group counts are batched. Expect the T2T run to approach the SolutionsEngineer run time. Keep MFA registration detail optional.
- Throttling and pagination: existing batching (`Invoke-ArrayaGraphCollectionBatch`) and the 500-user page cap with `signInActivity` already apply.
- Compatibility: larger T2T workbook; sheet names unchanged. `tests/unit/Common.Tests.ps1` assertions on the T2T sheet list need updating.
- Privacy: more identity data in the T2T workbook; keep the existing curated user columns.
- Tests: collection-plan unit tests; workbook sheet-order test; E2E synthetic export.
- Acceptance criteria: one T2T run produces the registers and the application, group, and device inventories.
- Direction: the profile should gain migration-relevant evidence: Entra groups, Microsoft 365 Groups, enterprise applications and SSO, application assignments, hybrid identity detail, the Conditional Access and MFA dependencies that affect migrated identities, a device and endpoint summary, fuller SharePoint detail, B2B and cross-tenant evidence, destination licensing evidence, and tenant rename evidence. It may take longer to run than `Presales`. It should not become a full security assessment: full Secure Score analysis, broad security posture findings, unrelated Purview maturity analysis, Teams Voice with no migration dependency, and unrelated usage analytics stay conditional on migration relevance.
- The profile contract test in `tests/unit/T2TP0Correctness.Tests.ps1` pins the current collector flags and the `RUN.md` profile table. Update the code, both documents, and that test together.

### P1-02 Enterprise application SSO evidence

- Files: `Get-AuthenticationConfiguration` in the legacy script (service principal query at line 16412; application query at line 16115; record builder `Add-AssessmentEnterpriseApplicationRecord`); `Export-HashTableToExcel.ps1` column lists for `EnterpriseApplications`.
- Change type: collection, normalization.
- Evidence to add, per service principal whose `preferredSingleSignOnMode` is `saml`, or which has a provisioning job:

| Evidence | Source | New field |
| --- | --- | --- |
| Entity ID candidates | `servicePrincipalNames` on the service principal; `identifierUris` on the application | `SamlEntityIds` |
| Sign-on, logout, relay state | `loginUrl`, `logoutUrl`, `samlSingleSignOnSettings.relayState` | `SignOnUrl`, `LogoutUrl`, `RelayState` |
| Signing certificate | `keyCredentials` (type, usage, `endDateTime`, `customKeyIdentifier` only) and `preferredTokenSigningKeyThumbprint` | See P1-04 |
| Notification addresses | `notificationEmailAddresses` | `NotificationEmailCount` and the addresses in the engineer workbook only |
| Owners | `GET /servicePrincipals/{id}/owners` | `ServicePrincipalOwners`, `ServicePrincipalOwnerCount` |
| Assignments | `GET /servicePrincipals/{id}/appRoleAssignedTo` | `AssignedUserCount`, `AssignedGroupCount`, `AssignedGroupIds` |
| Claims | `GET /servicePrincipals/{id}/claimsMappingPolicies`; `GET /beta/servicePrincipals/{id}/claimsPolicy` | `NameIdSource`, `NameIdFormat`, `NameIdTransformation`, `AdditionalClaimCount`, `GroupClaimsMode`, `ClaimsEvidenceState` |
| Group and optional claims | `groupMembershipClaims`, `optionalClaims` on the application | `GroupMembershipClaims`, `OptionalClaimNames` |
| Provisioning | `GET /servicePrincipals/{id}/synchronization/jobs` and each job schema | `ProvisioningJobCount`, `ProvisioningStatus`, `ProvisioningMatchAttributes` (attribute mappings with `matchingPriority` greater than zero, in order), `ProvisioningTargetUserNameSource`, `ProvisioningExternalIdSource`, `ProvisioningEvidenceState` |
| Conditional Access references | Existing `ConditionalAccessPolicies` | `ConditionalAccessPolicyCount` |
| Application proxy | `onPremisesPublishing` on the application (beta) | `ApplicationProxyPublished` |

- Derived risk fields: `NameIdUsesUpnOrMail` (true when the NameID source is `user.userprincipalname`, `user.mail`, or a transformation of them), `MatchAttributeUsesUpnOrMail`, `NoOwner`, `SourceDomainReferenced`. When no claims policy is readable, set `ClaimsEvidenceState = Unavailable` and do not assume the default NameID.
- Permissions: `Application.Read.All` and `Policy.Read.All` (both already required) cover service principal properties, owners, assignments, and claims mapping policies. `Synchronization.Read.All` is new and optional. The beta claims policy and application proxy calls must be verified against current documentation before use; treat failure as `Unavailable`.
- Runtime: up to five extra calls per SAML or provisioned application. Batch with `$batch` (20 requests per batch). Tenants with several hundred enterprise applications add minutes, not hours, because only SAML and provisioned applications are expanded.
- Throttling and pagination: page `appRoleAssignedTo` (large assignments); respect `Retry-After`; cap stored group ids and record the total.
- Compatibility: additive fields; no rename.
- Privacy and security: never store `keyCredentials.key`, `passwordCredentials.secretText`, or provisioning secrets (`synchronization/secrets` must not be called). Store assignment counts and group ids, not user lists.
- Tests: mocked Graph payloads for scenarios S10 to S16; privacy test asserting that no key material property name appears in the snapshot.
- Acceptance criteria: every expected finding in S10 to S16 is produced from mocked responses, and a missing `Synchronization.Read.All` yields `ProvisioningEvidenceState = Unavailable` instead of "no provisioning".

### P1-03 Application identity mapping register and worksheet

- Files: `Update-MigrationPlanningRegisters` (new register `MigrationApplicationIdentityMapping`); `Convert-ArrayaLegacyTenantStatsToSnapshot.ps1` (add to derived keys); `Export-HashTableToExcel.ps1` (column list, editable-schema list, T2T and Presales order); `Merge-M365MigrationScopeDecisions.ps1`; the template `docs/templates/T2T-Application-Identity-Mapping-Template.csv`.
- Change type: analysis, report.
- Change: one row per SSO application, `DecisionKey = APP:{appId}`. Pre-fill only what was collected (application, protocol, current tenant, NameID source and format, SCIM match attribute, evidence source). Leave every owner, vendor, criticality, JIT, rename, bulk-map, dual-IdP, test environment, cutover, and rollback field empty with status `Needs Input`. Add a `RiskCases` column computed from P1-02 fields using the nine risk cases in the scenario spec. Support a transitional design field for an application-specific NameID that keeps the source value while the Entra UPN changes; the value is entered by the engineer and is never proposed automatically.
- Permissions, runtime: none beyond P1-02.
- Compatibility: additive register.
- Tests: S10 to S16 expected rows; merge round trip; render-only workbook test.
- Acceptance criteria: the worksheet contains the 22 required columns in the template order, and no customer or vendor field is pre-filled.

### P1-04 SAML signing certificate expiry

- Files: `Get-AssessmentEnterpriseApplicationCredentialMetadata` and the service principal query in the legacy script.
- Change type: collection, analysis.
- Change: read service principal `keyCredentials` and `preferredTokenSigningKeyThumbprint`; report `SamlSigningCertificateExpiry`, `SamlSigningCertificateState` (Active, Expiring, Expired, None, Unavailable), and the count of inactive certificates. Use a configurable window with a default that covers a migration program (for example 120 days) in addition to the existing 30-day window.
- Permissions: none new. Runtime: none (same query, more properties).
- Privacy: dates, usage, and thumbprint only. Thumbprints are certificate identifiers; keep them out of customer-facing documents.
- Tests: S15.
- Acceptance criteria: a SAML application with a certificate ending inside the window is reported as expiring even when its application object has no credentials.

### P1-05 Tenant role and source-to-destination comparison

- Files: `Start-M365TenantAssessment.ps1` and the runner (new `-TenantRole Source|Destination` recorded in snapshot `Metadata`); a new script under `src/scripts/reporting` modeled on `Compare-M365TenantAssessmentSnapshots.ps1`; `Get-AllUserDetails` select list.
- Change type: collection (small), analysis, report.
- Change: add `identities`, `otherMails`, `onPremisesDomainName`, `onPremisesUserPrincipalName`, and `onPremisesSecurityIdentifier` to the user selection. Add an offline comparison that takes a source snapshot, a destination snapshot, and an optional mapping rule (new UPN domain, new SMTP domain, whether existing destination users keep their UPN). Output registers: `MigrationIdentityCollisions` (proposed UPN, primary SMTP, proxy address, mail nickname, and display name against destination users, groups, contacts, and recipients), `MigrationGuestToMemberMapping` (source guests matched to destination members by mail, other mail, and identity issuer; destination guests that are source members), and `MigrationDomainPresence` (each source and proposed domain against the destination domain list). Replace `ConflictStatus = "Not evaluated against destination"` with the result when a destination snapshot is supplied.
- Permissions: none new for the comparison. Optional `findTenantInformationByDomainName` lookup for a proposed domain uses `CrossTenantInformation.ReadBasic.All`, which is already documented.
- Runtime: offline; hash joins over users and recipients.
- Compatibility: additive; single-snapshot runs unchanged.
- Privacy: two tenants' identity data in one output; label the workbook accordingly and keep it in the engineer pack.
- Tests: S01, S05, S06, S07 with a paired synthetic fixture.
- Acceptance criteria: seeded duplicates are each reported once with both object identifiers, and a run without a destination snapshot keeps the existing placeholder.
- Design notes: build the comparison on reusable normalized objects and explicit matching keys for users, UPNs, SMTP and proxy addresses, mailboxes, groups, guests, Microsoft 365 Groups, Teams, SharePoint, enterprise applications, domains, and licensing. Candidate comparison states: `Matched`, `SourceOnly`, `DestinationOnly`, `Conflict`, `NeedsMapping`, `NeedsRecreation`, `NeedsDecision`, `NotComparable`. These states are a proposal only and are not in any production schema.

### P1-06 Hybrid identity evidence and explicit unknowns

- Files: `Get-AdConnectSyncDetails`; `Update-TenantToTenantMigrationExportData` (AD Connect worksheet rows).
- Change type: collection, analysis.
- Change:
  - Read the remaining `onPremisesSynchronization` feature flags (soft-match and hard-match blocking, cloud password policy, group writeback) and `configuration` values (accidental deletion threshold, synchronization interval).
  - Infer sync server candidates from directory synchronization service accounts: users whose UPN begins with `Sync_` and members of the Directory Synchronization Accounts role. Report as `Inferred`, with last sign-in, and state that stale accounts from retired servers can remain. Cloud Sync is suggested by an `ADToAADSyncServiceAccount` account; confirm the detection rule against current documentation.
  - Read staged rollout policies (`/policies/featureRolloutPolicies`) and per-domain federation configuration.
  - Always emit rows for "Sync server identity", "Active or staging mode", "Sync client version", "OU and domain filtering", and "Custom synchronization rules" with `EvidenceClass = Not collectible` and a pointer to the questionnaire, instead of omitting the section.
  - Remove the `Get-AzureADConnectHealth*` path or guard it with an evidence state; update the three source-pattern assertions in `tests/unit/TenantAssessmentCollector.Tests.ps1` accordingly.
  - Add an interpretation row for synchronized recipients: mail attributes of synchronized users are mastered on-premises, so a new primary SMTP for existing destination users is an AD change.
- Permissions: none new (`OnPremDirectorySynchronization.Read.All`, `Policy.Read.All`, `RoleManagement.Read.Directory`, `User.Read.All`).
- Runtime: three to five calls.
- Tests: S01, S02, S03 with mocked payloads.
- Acceptance criteria: a synchronized tenant with no inferable server produces an explicit "Entra Connect server unknown" evidence gap.

### P1-07 Licensing readiness

- Files: `Update-MigrationPlanningRegisters`; `Get-ArrayaLicenseAnalytics.ps1` (reuse); new decision keys.
- Change type: analysis, report.
- Change: add decision `LIC-01` "Destination SKU and add-ons approved" as a blocking migration dependency, split from ID-03. Add evidence rows, each labeled as source-tenant evidence: in-scope licensed member count; count against the 300-seat limit of Microsoft 365 Business Premium (a consideration, not a verdict, with a re-verify note); mailboxes above 50 GB, archive-enabled mailboxes, and shared mailboxes above 50 GB or with archives (license-profile populations); whether assigned SKUs contain a Teams service plan; groups used for licensing. Rename the "Target licensing readiness" row to "Source license utilization".
- Permissions: none new. Runtime: none.
- Compatibility: one renamed row label in `MigrationReadiness`.
- Tests: S04 and a negative case where `LIC-01` is confirmed.
- Acceptance criteria: with no license decision, the output contains a `Blocker` of type "Target-state decision" and never names a recommended SKU.

### P1-08 Mail flow, gateway, and migration-endpoint evidence

- Files: `Get-MailFlowRulesandConnectors.ps1`, `Get-ThirdPartySpamFilteringConfig.ps1`, `Get-SMTPRelayConfiguration.ps1`, `Get-ExchangeHybridConfiguration.ps1` in `src/modules/Arraya.M365.Exchange/Public`; `Update-TenantToTenantMigrationExportData`.
- Change type: collection, analysis.
- Change: collect with read-only cmdlets and report with evidence states: `Get-OrganizationConfig` (EWS enabled, EWS application access policy and lists); `Get-CASMailbox` EWS state counts; `Get-JournalRule`; `Get-DkimSigningConfig`; `Get-HostedConnectionFilterPolicy` allow lists; inbound connector enhanced-filtering settings; `Get-ReportSubmissionPolicy`. Cross-reference transport rules and connectors that mention accepted domains. Use the EWS evidence as `DiscoveredValue` for BT-09 while keeping the decision itself customer-confirmed.
- Permissions: existing Exchange read roles. Runtime: `Get-CASMailbox` for all mailboxes is one paged call; keep it to counts.
- Privacy: IP allow lists and gateway host names are sensitive; engineer pack only.
- Tests: mocked cmdlet outputs; S01 gateway case.
- Acceptance criteria: BT-09 shows collected EWS state for the source and remains `Needs Input` until confirmed.

### P1-09 Calendar and meeting evidence and semantics

- Files: new collector function in `Arraya.M365.Exchange` or `Arraya.M365.Graph`; `Update-TenantToTenantMigrationExportData`; cutover pack `PackSummary`.
- Change type: report first, then optional collection.
- Change, part A (no collection): add fixed semantics rows to the checklist and pack summary: calendar items are copied content; organizer identity, updates, cancellations, and responses for existing series are not preserved by a copy; Teams meeting links created in the source tenant continue to point at it; critical recurring meetings should be recreated by their organizers in the destination. Mark each as "Requirement" or "Risk" and state that third-party tool behavior must be confirmed with the vendor.
- Change, part B (opt-in switch, off by default): per in-scope mailbox, count series masters where the user is organizer, how many are online meetings, and how many have an end date beyond a cutover date, when the engagement supplies one. Counts only; no subject, body, attendee addresses, or locations.
- Permissions: part B needs application read access to calendars. `Calendars.ReadBasic.All` is the least-privileged candidate and must be verified; treat as optional.
- Runtime: one filtered query per mailbox; significant on large tenants. Throttling: mailbox-scoped limits apply; use batching and back-off.
- Privacy: calendar metadata is sensitive even as counts; opt-in, documented, engineer pack only.
- Tests: S09 with mocked responses; render-only test for part A.
- Acceptance criteria: part A rows appear in every T2T run; part B is absent unless the switch is set.

### P1-10 Group disposition register

- Files: `Update-MigrationPlanningRegisters` (new register `MigrationGroupDisposition`); workbook order and editable-schema lists.
- Change type: analysis, report.
- Change: one row per group across `EntraIDGroups`, `AllExchangeGroups`, `UnifiedGroups`, and `AllTeams`, joined on object id and primary SMTP. Columns: class (security, mail-enabled security, distribution, dynamic distribution, Microsoft 365 Group, Team, dynamic Entra, role-assignable), source of authority (cloud or synchronized), owners and owner count, primary SMTP, hidden membership, moderation, delivery restrictions, member count, and usage references (license assignment, Conditional Access include or exclude, application assignment from P1-02, role-assignable). Disposition is an entry column restricted to: Migrate, Recreate, Merge, Retain temporarily, Replace with another control, Retire, Customer decision required. Default is "Customer decision required". `DecisionKey = GROUP:{id}`.
- Permissions: none new. Runtime: none (join of collected data).
- Compatibility: additive; existing `MigrationObjectDisposition` group rows stay.
- Tests: S06, S08.
- Acceptance criteria: a group referenced by Conditional Access and by an application assignment shows both references and no default disposition other than "Customer decision required".

### P1-11 B2B guest replacement and cross-tenant access

- Files: `Export-ArrayaTenantToTenantCutoverPack.ps1`; `Get-FederationAndCrossTenantConfiguration`; `Update-MigrationPlanningRegisters`.
- Change type: collection, analysis, report.
- Change: leave `Direction` empty in `CollaborationAccessTemplate` and add it to the decision prompts; add counts of excluded guests and disabled users to `PackSummary`. Add a guest register (guest UPN, mail, external state, creation type, issuer, last sign-in, group and Team counts where known). Collect B2B collaboration inbound and outbound settings, identity synchronization (cross-tenant sync) per partner, and automatic redemption; fix the `DefaultInboundAccess` label; record a partner-lookup evidence state. Add fixed statements that removing a guest does not migrate 1:1 chat history, group chat history, meeting chat history, meeting ownership, or historical identity references, and that replacement access must be added and validated before guests are removed.
- Permissions: none new (`Policy.Read.All`).
- Compatibility: `Direction` changes from a constant to empty in the cutover pack CSV. Downstream helper scripts that expect the constant must be checked.
- Tests: S07; `Common.Tests.ps1` cutover pack assertions.
- Acceptance criteria: no cutover pack column contains a pre-filled design decision.

### P1-12 Readiness output model: item type, workstream, engagement-supplied milestones, governance keys

- Files: `Update-MigrationPlanningRegisters`, `Update-MigrationComplexityFlags`, `Update-TenantToTenantMigrationExportData`; workbook column lists; presales questionnaire exporter; `Merge-M365MigrationScopeDecisions.ps1`.
- Change type: analysis, report.
- Change: add the attributes defined in the coverage document (`ItemType`, `Milestone`, `Workstream`, `TenantRole`, `EvidenceClass`, `EvidenceSource`, `DependencyType`, `BlockedBy`, `OwnerState`, `DateState`). Add decision keys: `HYB-01` Entra Connect server confirmed, `HYB-02` AD administration server and writable domain controller confirmed, `HYB-03` destination OU and sync scope confirmed, `LIC-01` (P1-07), `ACC-01` administrative access granted in both tenants and on-premises, `TOOL-01` migration-tool licensing procured, `DNS-01` DNS authority and registrar access confirmed, `APP-01` application owners identified, `APP-02` SAML and SCIM matching behavior confirmed per application, `EP-01` endpoint transition approach confirmed, `COM-01` user communication plan approved, `RB-01` rollback tested, `SRC-01` source-access coexistence period and method confirmed, `REN-01` SharePoint rename and fallback domain decision. Each defaults to `Needs Input` with empty owner and date.
- Permissions, runtime: none.
- Compatibility: additive columns on register worksheets; `Status` values unchanged.
- Tests: S04, S18 to S21; negative cases for defaults.
- Acceptance criteria: every seeded governance gap in the scenario spec appears as a separate row with the expected item type, and no owner, date, or milestone is ever generated.

### P1-13 Tool-neutral mailbox routing guidance

- Files: `Update-TenantToTenantMigrationExportData` (checklist text), `Invoke-HTMLHelperFunctions.ps1` line 4061, cutover pack notes.
- Change type: report.
- Change: split the guidance into "applies to any method" (preserve `LegacyExchangeDN` as an X500 address on the target, preserve existing X500 and X400 proxies, re-apply delegates) and "Microsoft native cross-tenant mailbox migration only" (MailUser stamping of `ExchangeGuid` and `ArchiveGuid`, source onmicrosoft.com routing address). Show the second group only when the confirmed tool decision is the native method, otherwise mark it "Not applicable to the selected tool" or "Tool not confirmed".
- Tests: render-only.
- Acceptance criteria: no unconditional instruction to stamp `ExchangeGuid` appears when the tool decision is a third-party copy tool.

### P1-14 Hold coverage beyond collected fields

- Files: `Get-AllExchangeMailboxDetails.ps1`; `Get-PurviewCompliancePolicies` in the legacy script.
- Change type: collection, analysis.
- Change: add `ComplianceTagHoldApplied` and `DelayReleaseHoldApplied` to the mailbox property lists; read organization-level holds from `Get-OrganizationConfig`; classify `InPlaceHolds` entries by prefix; optionally list eDiscovery cases and case hold policies through the existing Purview session. Report counts per hold type and never a legal conclusion.
- Permissions: existing Exchange read; Purview case read roles for the optional part.
- Tests: S20 extended.
- Acceptance criteria: each hold type is counted separately with its evidence state.

### P1-15 Domain evidence and evidence classes

- Files: `Get-AllOffice365Domains`; `Update-MigrationPlanningRegisters` domain section.
- Change type: collection, analysis.
- Change: populate `IsInitial`, `IsRoot`, `IsAdminManaged`, `SupportedServices`, and `State` from the domain resource. Call `GET /domains/{id}/domainNameReferences` and store counts by object type. Read `federationConfiguration` for federated domains. Count onmicrosoft.com domains against the limit of five. Add four evidence-class columns per domain: verified in tenant (collected), DNS records observed (collected with resolver state), DNS administration owner confirmed (customer), production mail flow validated (customer or test evidence). Never derive the last two.
- Permissions: none new. Runtime: one call per domain.
- Throttling: `domainNameReferences` can be large; request counts with paging and cap stored samples.
- Tests: S01, S17, N06.
- Acceptance criteria: a verified domain with no customer confirmation shows "DNS owner: not confirmed" and "Mail flow: not validated".

### P1-16 SharePoint tenant rename readiness

- Files: `Get-AssessmentSharePointAdminUrl` (persist hostname); `Update-TenantToTenantMigrationExportData`; decision `REN-01`.
- Change type: normalization, analysis, report.
- Change: store the SharePoint hostname in `TenantInfo`. Add a `TenantRenameReadiness` table: total SharePoint plus OneDrive site count against the documented standard limit; Multi-Geo flag (P0-04); whether the hostname differs from the initial domain; onmicrosoft.com domain count; cloud environment; items that are not discoverable (previous rename count, SharePoint Embedded containers, public site, alternate URLs). Add owner-validation rows for Power Automate, Power BI, Forms, SPFx, custom applications, backup products, and hard-coded URLs, each as "Customer question" with no assumed outcome, and a fixed statement that redirects are temporary and do not cover every workload. Keep limit values in a baseline JSON with a last-reviewed date, as `migration-tooling-options.json` does.
- Permissions: none new.
- Tests: S17.
- Acceptance criteria: the table never states "eligible"; it reports evidence, limits with their review date, and open questions.

### P1-17 Endpoint evidence honesty

- Files: `Get-AllDevicesReport`; `Update-DeviceManagementSummary`; questionnaire exporter.
- Change type: analysis, report.
- Change: replace `MDMSolution = "Not Managed"` with "No management reported in Entra"; add a fixed note on device tables that Entra lists only registered, joined, and hybrid-joined devices and cannot show devices joined only to on-premises Active Directory; add `EP-01` and a statement that the absence of Intune enrollment does not establish that devices can move without a reset or profile work.
- Permissions: none.
- Tests: S03, S19.
- Acceptance criteria: a source tenant with zero hybrid-joined devices and no AD evidence reports domain-joined device count as "Not collected", not zero.

### P1-18 Multi-profile composition

Status: criteria composition is implemented (see [T2T-Composition-Status.md](T2T-Composition-Status.md)); combined runs build one deterministic register and show missing delegate evidence as `NotRequested`. Collection-level composition below is still open.

- Problem: requesting several profiles runs once at the highest depth, but the collection scope and workbook policy reset to `Default`. When `TenantToTenantMigration` is combined with another profile, the tenant-to-tenant shaping is skipped, mailbox and calendar delegate permissions are not collected, the planning registers are not built, and the cutover pack is still written from unshaped data. A combined run can yield less migration evidence than the same profiles run separately. The limitation is documented in `RUN.md`; no code was changed.
- Change surface: profile merge in `src/modules/Arraya.M365.AssessmentRunner/Arraya.M365.AssessmentRunner.psm1` (the block that sets `WorkbookExportPolicy`, `CollectionScopePolicy`, and `GenerateMigrationPack` for merged selections); `Resolve-AssessmentProfileCollectionPlan` (the `IsMergedOutputProfileSelection` branches); `Get-ArrayaCollectionDepthPolicy` (scope-specific depth flags, including delegate collection); `Prepare-AssessmentExportData` (the `isTenantToTenantCutover` and `isPresalesScope` tests); `Invoke-M365TenantAssessmentExportPipeline.ps1` (one workbook policy per run); `Export-HashTableToExcel.ps1` (policy-specific sheet lists).
- Target invariant: requesting several profiles never yields less evidence than any one of them alone. Likely design, to be reviewed before implementation: collectors are the safe union of the requested profiles; depth is the highest required; outputs are the union where compatible; each requested deliverable gets the shaping it needs.
- Interim option: log a warning and add a checklist row when a migration profile is merged with another profile.
- Tests: plan and depth-policy tests for each pair of profiles; a test that the union never drops a collector or depth flag.

### P1-19 Evidence state for collectors that catch their own errors

Status: implemented for public folders, mail flow rules and connectors, domains, remote domains, and Purview policies through `Set-ArrayaCollectorEvidenceState`. Mailbox enumeration remains open.

- Problem: several collectors log an error and return, so the collector plan records the step as Completed and leaves an empty container. `Get-ArrayaDatasetEvidenceState` cannot tell that apart from a genuinely empty result. Affected today: `Get-AllPublicFolderDetails.ps1`, `Get-MailFlowRulesandConnectors.ps1`, the Exchange portion of `Get-AllOffice365Domains`, and mailbox enumeration in `Get-AllExchangeMailboxDetails.ps1` (errors suppressed).
- Change: have each collector record an evidence state the way `Get-ExchangeHybridConfiguration.ps1` and `Get-SMTPRelayConfiguration.ps1` already do, or return a failed step result, and read it in `Get-ArrayaDatasetEvidenceState`. Decide where the state is stored so it does not add worksheets.
- Until then: a collected-empty result for public folders, connectors, and transport rules is trusted when the step reports Completed. Domains and remote domains are protected by the rule that they cannot be empty.
- Tests: mocked cmdlet failures per collector.

## P2: detailed design and estimation

| ID | Change | Files or modules | Type | New permissions or access | Runtime and throttling | Compatibility | Privacy and security | Tests and acceptance |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| P2-01 | Optional endpoint and AD discovery with an import contract (join state on the device, BitLocker protector and escrow location, local administrators, mapped drives, printers, VPN, certificates, Wi-Fi profiles, Windows Hello state, services and scheduled tasks using domain or user credentials, work or school account registrations, profile SID). Tenant side: BitLocker recovery key metadata per device. | New module outside the tenant collector; import function in `Arraya.M365.Common`; snapshot `Data.Endpoint` | Collection, normalization | On-premises or endpoint management tooling; `BitLockerKey.ReadBasic.All` (verify) for metadata | Endpoint side is asynchronous; tenant metadata is one paged call | Additive; import is optional | Never collect recovery key values; device names and users are personal data | Import schema test; S03, S19. Accept when an import produces device counts by join state with evidence class "Imported" |
| P2-02 | Optional file-server discovery module with an import contract (server, share, source path, size, file and folder counts, age buckets, invalid characters, path length, large files, blocked types, locked files, NTFS groups and nested groups, home drives, departmental shares, stale data, owner, destination mapping, migration worker placement). Prefer importing the migration tool's own source analysis where one exists. | New module; import function; register `MigrationFileShareDisposition` | Collection, normalization, report | On-premises read on shares; local administrator for share enumeration | Long-running on-premises scan; out of band | Additive | File names and ACLs are sensitive; keep aggregates in customer documents | Import schema test; S18. Accept when a share marked critical with no plan for its agreed milestone raises a migration risk |
| P2-03 | Link the AD workflow and Entra Connect server export to the snapshot: export AD report tables to a documented JSON, import as `Data.OnPremises`; accept an optional sync-server export (connectors, filtering, rules, staging mode, version). | `Get-ActiveDirectoryReport.ps1`; `Arraya.M365.Common` import | Normalization | On-premises: domain user with RSAT; local administrator on the sync server | Out of band | Additive | Directory data; engineer pack only | Import tests; S02, S03 |
| P2-04 | SharePoint permission and item-level evidence by import (sharing links, external users, long paths, large and checked-out files, workflows). Do not rebuild a scanner. | Import function; `SharePoint` enrichment columns | Normalization | None in the tenant collector | None | Additive | Aggregates only | Import schema test |
| P2-05 | Teams membership and chat readiness: owner and member identities per Team and channel; per-user private chat, meeting, and call counts from activity reports; neutral method comparison rows (rehydrated chats, group-chat conversion, reference exports) as vendor questions. | `Get-TeamsDetails`; reports collector | Collection, report | `ChannelMember.Read.All` (new, optional); `TeamMember.Read.All` (optional today) | One call per channel; batch; large tenants add notable time | Additive | Membership lists in engineer pack only; activity reports may be anonymized by tenant setting, which must be reported as an evidence state | S08; mocked responses |
| P2-06 | Additional Exchange evidence: non-calendar folder permissions (opt-in), mailbox audit settings, `Get-CalendarProcessing` for resource mailboxes, room lists, mobile device partnerships (counts), IRM and message-encryption configuration, full inbox-rule counts, Recoverable Items subfolder sizes. | `Arraya.M365.Exchange` | Collection | Existing Exchange read | Per-mailbox calls; opt-in for folder permissions | Additive | Rule and folder names can be sensitive | Mocked cmdlet tests |
| P2-07 | SMTP AUTH actual usage from sign-in logs (client app "Authenticated SMTP") and relay volume. | `Get-SMTPRelayConfiguration.ps1` | Collection, analysis | None new (`AuditLog.Read.All`) | Filtered sign-in query; retention-limited | Additive | None beyond existing sign-in data | Mocked responses |
| P2-08 | Pre-stage split and alternative sequencing: optional mailbox age evidence; optional wave templates for other sequences (for example identity and mail flow before history backfill), chosen by the engagement; engagement-supplied milestone names on wave rows. | `Update-MigrationPlanningRegisters` | Analysis, report | None | Age evidence is per-mailbox and costly; opt-in | New wave template selected by decision DC-02 | None | S20; render-only |
| P2-09 | Sensitivity labels and label policies; resolve site label ids to names. | Purview collector | Collection | Existing Purview session or Graph label read (verify) | Small | Additive | None | Mocked responses |
| P2-10 | Security detail: named location definitions, PIM eligibility, per-application failed sign-in summary. | Legacy collector | Collection | None new | Sign-in summary is the costly part; cap by window | Additive | IP ranges are sensitive | Mocked responses |
| P2-11 | User-experience readiness register with the fixed classification from the questionnaire document. | `Update-MigrationPlanningRegisters` | Report | None | None | Additive | None | Render-only |
| P2-12 | Tenant-to-tenant objectives baseline and evidence coverage validator, modeled on `solutions-engineer-assessment-objectives.json` and `Test-SolutionsEngineerAssessmentEvidence.ps1`. | `src/config/baseline`; `src/scripts/reporting` | Test, report | None | None | Additive | None | Validator tests |
| P2-13 | AD workflow defects noticed in passing: a stale `$user` reference in the group loop (line 1298), `ManufacturerInstallDate` assigned a script block (line 1340), and tables keyed by name. Needs its own review. | `Get-ActiveDirectoryReport.ps1` | Collection | None | None | Behavior change in AD export | None | New tests |

## P3: optional

| ID | Change | Files or modules | Notes |
| --- | --- | --- | --- |
| P3-01 | Data validation lists for entry columns in register worksheets. | `Export-HashTableToExcel.ps1` | Usability only |
| P3-02 | Questionnaire round trip for the new decision keys. | Presales questionnaire exporter; `Merge-M365MigrationScopeDecisions.ps1` | After P1-12 |
| P3-03 | Profile documentation and ASCII cleanup of existing Markdown. | `README.md`, `RUN.md` | `RUN.md` contains non-ASCII punctuation |
| P3-04 | Register diff between two runs of the same tenant. | `Compare-M365TenantAssessmentSnapshots.ps1` | Tracks decision progress |
| P3-05 | Rename misleading field `DefaultInboundAccess` (holds the MFA trust value). | `Get-FederationAndCrossTenantConfiguration` | Needs a compatibility alias |
| P3-06 | Target mapping rule helper (apply a UPN and SMTP domain rule to draft target columns, always marked "Draft - confirmation required"). | `Update-MigrationPlanningRegisters` | After P1-05 |

## New permissions summary

| Permission | Type | Needed by | Required or optional |
| --- | --- | --- | --- |
| `Synchronization.Read.All` | Graph application or delegated | P1-02 provisioning evidence | Optional; verified on Microsoft Learn 2026-11-10 |
| `Calendars.ReadBasic.All` | Graph application | P1-09 part B | Optional, opt-in; verify before use |
| `ChannelMember.Read.All` | Graph application | P2-05 | Optional |
| `BitLockerKey.ReadBasic.All` | Graph | P2-01 | Optional; verify application support before use |
| Purview case read roles | Purview | P1-14 optional part, P2-09 | Optional |

P0 items and P1-01, P1-03 to P1-08, P1-10 to P1-13, and P1-15 to P1-17 need no new permission.

## Recommended sequence

1. Done: P0-01 to P0-07 as one change set, each with its fail-first test.
2. P1-01 with P1-18 and P1-19 (profile scope, profile composition, collector evidence states), then P1-12 (output model), because later items report into them.
3. P1-02, P1-04, P1-03 (application evidence, certificate expiry, mapping worksheet).
4. P1-05 and P1-11 (two-tenant comparison and guest replacement).
5. P1-06, P1-15, P1-16, P1-07 (hybrid identity, domains, rename, licensing).
6. P1-08, P1-13, P1-14, P1-09 part A, P1-10, P1-17.
7. P2 items by project demand; P2-12 early if several engineers will extend the rules.
