# T2T Synthetic Scenario Spec

Review date: 2026-11-10
Related documents: [coverage](T2T-Migration-Assessment-Coverage.md), [backlog](T2T-Migration-Assessment-Backlog.md)

This document defines the synthetic fixtures, expected findings, expected report output, negative cases, and regression tests that the backlog items must satisfy. Expected results are declared here before any implementation.

## 1. Rules

- Synthetic data only. No customer export, tenant export, or real identifier may be used as a fixture or as a model for one.
- Reserved names only: `contoso.example` (source), `fabrikam.example` (destination), `newbrand.example` (new UPN and SMTP domain), `source-domain.example` (old identifier in application examples), `contoso.onmicrosoft.com` and `fabrikam.onmicrosoft.com` as placeholder initial domains, GUIDs of the form `00000000-0000-0000-0000-0000000000xx`.
- Tests call production functions. Follow the AST-extraction pattern in `tests/unit/MigrationScopeReconciliation.Tests.ps1` for functions in `Get-FullTenantReportDetails.ps1`, dot-source `Invoke-HTMLHelperFunctions.ps1` and the questionnaire exporters as `tests/unit/Reporting.Tests.ps1` and `Common.Tests.ps1` do, and import modules by manifest. Do not re-implement analysis logic inside a test.
- Collectors are tested with mocked Microsoft Graph and Exchange responses by stubbing the transport helpers (`Get-ArrayaGraphResource`, `Get-AssessmentTenantOrganization`, `Invoke-ArrayaGraphCollectionBatch`, Exchange cmdlets) and never by a live call.
- No secrets, tokens, tenant ids, or certificate material in fixtures, logs, or expected output. Certificate fixtures carry dates and usage only.
- ASCII only.

## 2. Validation ladder

Run in this order and stop at the first level that can prove the behavior.

1. Write the expected behavior and acceptance criteria (this document and the backlog item).
2. Run the existing suites: `./tools/run-pester.ps1` and the two documented Python commands in `docs/runbooks/assessment-review-agent.md` and `docs/runbooks/customer-deliverables-mvp.md`.
3. Add deterministic fixtures with the expected results declared in sections 4 to 6.
4. Use focused object diagnostics where a result is surprising (print the single row under test, not the whole store).
5. For report changes, use render-only generation: build the store in memory and call the exporter (`Export-HashTableToExcel`, `Export-ArrayaTenantToTenantCutoverPack`, `New-TenantMigrationCutoverHtmlReport`, questionnaire exporters) or `-Action Report` against a synthetic snapshot.
6. Convert each confirmed defect into a fail-first regression test (section 7) before changing the code.
7. Reuse production collection and analysis code in every test.
8. Run the small end-to-end synthetic reference dataset (section 8).
9. Use the approved internal test tenant, read-only, only when a specific Microsoft behavior cannot be established from documentation or mocks. Candidates are listed in section 9. The repository does not currently define a test-tenant procedure, so one must be written and approved first.
10. Do not run a production-scale assessment to validate presentation or report wording.

## 3. Fixture design

### 3.1 Builders

Add one fixture file, for example `tests/unit/T2TMigrationFixtures.ps1`, with small composable builders in the style of `New-TestStore`:

| Builder | Returns |
| --- | --- |
| `New-T2TSourceStore` | A cloud-only source tenant store (`contoso.example`) with collected-state markers for every dataset |
| `New-T2TDestinationStore` | A hybrid destination tenant store (`fabrikam.example`) |
| `Add-T2TCollectorState` | Sets `CollectorPlanResults` rows (Completed, Skipped, Failed) for named datasets |
| `New-T2TUser`, `New-T2TGuest`, `New-T2TMailbox`, `New-T2TGroup`, `New-T2TTeam`, `New-T2TSite`, `New-T2TDevice`, `New-T2TApplication` | Single rows with only the fields the production code reads |
| `New-T2TGraphPayload` | Mocked Graph JSON objects for organization, domains, service principals, claims policies, synchronization jobs and schemas |
| `New-T2TDecision` | A `MigrationScopeDecisions` row |

Every builder must take explicit parameters for the values a test asserts on. No hidden defaults that could make a test pass by accident.

### 3.2 Baseline population

| Object | Source store | Destination store |
| --- | --- | --- |
| Members | 12 enabled, 2 disabled, 1 service account, 1 shared mailbox owner account | 20 enabled synchronized users, 2 cloud-only administrators |
| Guests | 3 (one whose mail matches a destination member) | 2 (one whose mail matches a source member) |
| Mailboxes | 12 user, 2 shared, 1 room, 1 inactive with archive | 20 user |
| Groups | 2 cloud security, 1 dynamic, 1 role-assignable, 1 licensing, 2 distribution, 1 mail-enabled security, 3 Microsoft 365 Groups (2 Teams) | 2 synchronized security, 1 cloud security |
| Sites | 5 SharePoint, 12 OneDrive | 3 SharePoint, 20 OneDrive |
| Devices | 6 Entra registered, 0 hybrid joined | 10 hybrid joined |
| Applications | 6 SAML or SCIM applications as defined in S10 to S16 | 1 |
| Domains | `contoso.example` verified; `contoso.onmicrosoft.com` | `fabrikam.example` verified; `fabrikam.onmicrosoft.com` |

## 4. Required scenarios and expected results

"Today" states the current behavior where it was reproduced (R ids from the coverage document). "Expected" is the acceptance target.

| ID | Scenario | Fixture essentials | Expected findings | Expected report output | Backlog |
| --- | --- | --- | --- | --- | --- |
| S01 | Destination hybrid identity with Entra Connect | Destination store: `OnPremisesSyncEnabled = true`, last sync time set, password hash sync true, synchronized users and groups; role `Destination` | "Directory synchronization: On-prem sync enabled" as Confirmed current state; requirement that new accounts are created in AD; interpretation that mail attributes of synchronized users change on-premises; no "Exchange hybrid" flag unless hybrid evidence is true | AD Connect worksheet lists feature flags and the five not-collectible rows; checklist row with phase Day 1 and workstream Identity | P0-01, P1-05, P1-06 |
| S02 | Unknown Entra Connect server | S01 plus no inferable sync service account; organization read succeeded | Evidence gap "Entra Connect server unknown" with `EvidenceClass = Not collectible`; decision `HYB-01` Needs Input, owner not assigned, date not confirmed | Row present in checklist and decisions. Today: section silently omitted | P1-06, P1-12 |
| S03 | Source cloud-only users plus source on-premises AD devices | Source store: `OnPremisesSyncEnabled = null` with successful organization read; zero hybrid-joined devices; questionnaire answer or AD import states 40 domain-joined devices | "No directory synchronization reported by Microsoft Graph" as Info, never "Cloud-only identity model, Ready"; no AD Connect flag (today: flag raised, R1); domain-joined device count "Not collected" unless imported, then 40 with class Imported or Customer-confirmed | Device tables carry the Entra-only note; `EP-01` Needs Input | P0-01, P0-02, P1-17, P2-01 |
| S04 | Unresolved target license SKU | No `LIC-01` decision; 15 in-scope licensed members; one 60 GB mailbox; two archives | Blocker of type Target-state decision, phase Day 1, "Destination SKU and add-ons approved"; evidence rows for population counts labeled as source evidence; no SKU named or recommended | Present in readiness and decisions; "Target licensing readiness" label no longer used | P1-07 |
| S05 | Duplicate SMTP and UPN values | Paired stores with mapping rule to `newbrand.example`: one proposed UPN equals an existing destination UPN; one proposed primary SMTP equals a destination proxy address; one mail nickname clash; one display-name clash | Four collision rows, each with both object ids, collision type, and severity (UPN and SMTP as Blocker, nickname as Review, display name as Info) | `MigrationIdentityCollisions` worksheet; `ConflictStatus` updated on the matching mapping rows | P1-05 |
| S06 | Mixed cloud-only and synchronized groups | Destination groups as in the baseline; source groups cloud-only | Group register shows source of authority per group; synchronized groups marked "created in AD"; default disposition "Customer decision required" | `MigrationGroupDisposition` worksheet | P1-10 |
| S07 | Guest-to-member replacement mapping | Source guest with mail `pat@fabrikam.example`; destination member with the same UPN; one unmatched guest | One mapping row with match basis "mail equals destination UPN"; one "No destination match" row; fixed statements about what guest removal does not migrate; requirement "add and validate replacement access before removing guests" | Guest register; cutover pack `Direction` empty (today: `MemberToGuest`) | P1-05, P1-11 |
| S08 | Microsoft 365 Group and Teams ownership | One Team with two owners, one of them a guest; one Team with zero owners; one group whose only owner is disabled | Ownership gaps listed per group; owner identities available in the engineer workbook when membership permissions are granted, otherwise counts with `Needs Data` for identities | Group register usage and owner columns | P1-10, P2-05 |
| S09 | Recurring meetings organized by source users | Part A: none. Part B (opt-in): mocked responses giving 3 series for one organizer, 2 online, 2 ending after Day 1 | Part A rows always present: copied items are not operational continuity; organizer identity; recreate critical series; meeting links tied to source. Part B: counts only, no subjects | Checklist and pack summary semantics rows; Part B table absent without the switch | P1-09 |
| S10 | SAML NameID based on source UPN | Service principal with SAML mode; claims evidence giving NameID source `user.userprincipalname`, format emailAddress | `NameIdUsesUpnOrMail = true`; risk case 1; worksheet row with NameID pre-filled and vendor fields empty | Application identity mapping worksheet | P1-02, P1-03 |
| S11 | SAML NameID based on an extension attribute | NameID source `user.extensionattribute1` | `NameIdUsesUpnOrMail = false`; no risk case 1; note to confirm the attribute is populated for new destination accounts | Worksheet row | P1-02, P1-03 |
| S12 | SCIM matching on userName | Synchronization schema: `userName` mapped from `userPrincipalName`, matching priority 1 | `MatchAttributeUsesUpnOrMail = true`; risk case 5; vendor question on behavior when userName changes | Worksheet row | P1-02, P1-03 |
| S13 | SCIM matching on externalId | Schema: `externalId` mapped from `objectId` or `employeeId`, matching priority 1; userName still mapped | `MatchAttributeUsesUpnOrMail = false`; note that userName will still be updated and the application reaction must be confirmed; note that an object id match does not survive a new destination object | Worksheet row | P1-02, P1-03 |
| S14 | JIT duplicate-account risk | S10 application; worksheet answer `JIT enabled = Yes` merged by `DecisionKey` | Risk cases 3 and 4; item type Risk, phase Day 1; "Customer/vendor confirmation required" lists the open vendor questions | Worksheet and readiness rows | P1-03 |
| S15 | Expiring SAML signing certificate | Service principal key credential with usage Sign ending in 45 days; application object without credentials | `SamlSigningCertificateState = Expiring` with the date (today: `HasCredentialsExpiringSoon = false`) | Application worksheet and readiness risk row | P1-04 |
| S16 | Application with no owner | Gallery SAML application; zero service principal owners; no notification address | `NoOwner = true`; decision `APP-01` Needs Input; "Owner not assigned" | Worksheet row with empty owner fields; readiness row | P1-02, P1-12 |
| S17 | SharePoint tenant rename with integration dependencies | Source store with 17 sites; organization payload variants for Multi-Geo true, false, null; hostname equal to and different from the initial domain; questionnaire answers naming two flows and one backup product | Rename readiness table with site count against the documented limit and its review date; Multi-Geo value from the payload (today: always "No", R7); owner-validation rows for each integration class; statement that redirects are temporary; never the word "eligible" | `TenantRenameReadiness`; decision `REN-01` | P0-04, P1-16 |
| S18 | Critical file shares not migrated by Day 1 | Imported file-share rows: two marked critical, one without a Day 1 plan | Risk, phase Day 1, naming the share and its owner field (empty if not provided) | File-share register | P2-02 |
| S19 | Endpoint migration deferred to Phase 2 | Decision `EP-01` confirmed with phase Phase 2 and a customer-entered date; device rows as in S03 | Endpoint rows carry phase Phase 2; a Day 1 requirement remains for "existing devices keep working with the new Microsoft 365 identity"; statement that absence of Intune does not establish no reset or profile work | Readiness rows by phase | P1-12, P1-17 |
| S20 | Old mailbox retained temporarily for historical access | Decision `SRC-01` confirmed with a customer-entered period; two mailboxes with hold signals as in R10 | Confirmed target-state decision for source access; hold rows `Review` with count 2 and evaluated hold types (today: Ready, 0); requirement to keep licensing and holds on the source until release | Checklist and decisions | P0-07, P1-12, P1-14 |
| S21 | Out-of-scope integration impact requiring owner acknowledgment | Decision marking a Power BI report as out of scope with no acknowledgment recorded | Item with phase Out of scope and dependency "Customer decision"; evidence gap "owner acknowledgment not recorded"; it must not disappear from the report | Readiness and decisions | P1-12, P1-16 |

## 5. Expected report output rules

These apply to every scenario and should be asserted by render-only tests.

1. A row whose evidence was not collected is rendered as `Needs Data` with the reason. It is never rendered as `Ready` and never omitted.
2. Owner, due date, target mapping, and customer-confirmed columns are empty unless the fixture supplied them.
3. No output proposes a SKU, a tool, a cutover date, or a group disposition other than "Customer decision required".
4. Tool-specific guidance appears only when the tool decision is confirmed, and names the tool it applies to.
5. Every row built from two tenants names the tenant role of each value.
6. Customer-facing outputs contain no thumbprints, key ids, IP allow lists, or user lists from application assignments.
7. Worksheet names and existing column names are unchanged; new columns are appended.

## 6. Negative cases

| ID | Case | Must not happen | Must happen |
| --- | --- | --- | --- |
| N01 | Organization read fails | "Cloud-only identity model"; "No active on-premises directory synchronization detected" | `Needs Data` with the failure reason |
| N02 | Hybrid collector ran and found nothing | "Exchange hybrid, Configured" | No flag; hybrid worksheet shows "Not configured" |
| N03 | Dataset collected and genuinely empty (zero public folders, zero connectors) | `Needs Data` | `Ready`, hidden from the checklist as today |
| N04 | Collector skipped by profile | Silent omission | `Needs Data`, "Not collected in this profile" |
| N05 | Multi-Geo property null or absent | "No multi-geo configuration detected" | "Not reported by Microsoft Graph; confirm with the customer" |
| N06 | DNS resolver unreachable; MX at `mx.microsoft` | "non-Microsoft 365 MX"; third-party filter review | Null with DNS evidence state for the first; true for the second |
| N07 | `InPlaceHolds` not collected (non-All mode) | "0 mailbox(es) on hold" | Hold row states which hold types were evaluated |
| N08 | `Synchronization.Read.All` not granted | "No provisioning configured" | `ProvisioningEvidenceState = Unavailable` with the permission name |
| N09 | Claims policy unreadable | Assuming the default NameID | `ClaimsEvidenceState = Unavailable`; worksheet NameID empty |
| N10 | SAML application with NameID on a stable attribute | Risk case 1 | No NameID risk |
| N11 | No destination snapshot supplied | Collision rows or "no collisions found" | `ConflictStatus = "Not evaluated against destination"` |
| N12 | Section-level approval only (for example ID-02 approved) | Object rows confirmed | Object rows stay unconfirmed, as existing tests already require |
| N13 | Decision with status approved but empty owner or date | Generated owner or date | "Owner not assigned"; "Date not confirmed" |
| N14 | Application credentials fixture | Any `key`, `secretText`, or token property in the snapshot or workbook | Dates, usage, and type only |
| N15 | Guest whose mail matches no destination object | A proposed target identity | "No destination match"; customer decision required |
| N16 | Tenant with exactly the documented site limit, and one more | "Eligible" or "Not eligible" verdicts | Count, limit, review date, and the open questions |
| N17 | Merged profile selection including `TenantToTenantMigration` | Silent loss of cutover shaping | Shaping applied, or an explicit warning and checklist row |
| N18 | Tool decision not confirmed | Unconditional `ExchangeGuid` stamping instruction | Guidance grouped by method with "Tool not confirmed" |

## 7. Regression tests for confirmed defects

Each test must fail on the current code and pass after the fix. Suggested homes follow existing suites.

| Defect | Fixture | Assertion | Suggested test file |
| --- | --- | --- | --- |
| R1 | Store with `AdConnectConfiguration` container whose `Summary.OnPremisesSyncEnabled` is null, empty `SyncServices`, `ErrorCount` 0 | `MigrationComplexityFlags` has no item "AD Connect directory synchronization"; `MigrationQuoteReadiness` has no key `FLAG:ad-connect-directory-synchronization` | `MigrationScopeReconciliation.Tests.ps1` |
| R2 | Store with `HybridConfiguration.ExchangeHybrid.IsHybridConfigured = false` | No item "Exchange hybrid" | Same |
| R3 | Store with users only and no Exchange keys | No `CutoverPrepSummary` row has status `Ready`; `MigrationReadinessChecklist` contains `Needs Data` rows for connectors, public folders, holds, forwarding, and remote domains | Same |
| R4 | `Get-AdConnectSyncDetails` with a throwing organization stub, then `Update-AssessmentReportTables` | "Directory synchronization" status is `Needs Data` | `Reporting.Tests.ps1` |
| R5 | Empty `Domains` | "Verified custom domains" and "Mail routing" are `Needs Data` | Same |
| R6 | No hybrid or public folder keys | Both rows `Needs Data` | Same |
| R7 | `Get-TenantOverviewInfo` with an organization stub carrying `isMultipleDataLocationsForServicesEnabled = true`, authentication type set to a non-interactive mode | `TenantInfo.MultiGeoEnabled` is true; questionnaire line contains "Yes" | `TenantAssessmentCollector.Tests.ps1`; `Common.Tests.ps1` |
| R8 | Questionnaire export with only `TenantInfo` | No line contains "No active on-premises", "appears to be the primary", or a zero count for an uncollected dataset | `Common.Tests.ps1` |
| R9 | MX classification function with null, an `mx.microsoft` name, and a legacy name | Null, true, true | New unit test beside the extracted function |
| R10 | Two mailboxes with `InPlaceHolds`, `DelayHoldApplied`, `RetentionHoldEnabled` and no litigation hold | "Mailboxes on hold" is `Review` with value 2 and names the evaluated hold types | `MigrationScopeReconciliation.Tests.ps1` |

The reproduction used during this review loaded the same functions and fixtures as above and produced the "today" results recorded in the coverage document. It was a diagnostic and is not part of the repository.

## 8. End-to-end synthetic reference dataset

- One source snapshot and one destination snapshot in schema v2, built by the fixture builders and saved under the test drive, never under a tracked path.
- Run `-Action Report` with the `TenantToTenantMigration` profile against the source snapshot, then the two-snapshot comparison.
- Assert the artifact set (workbook, cutover pack, HTML, JSON, manifest), the sheet order, the row counts of each register against the baseline population in section 3.2, and the presence of every expected finding in section 4.
- Keep the dataset small enough to run in the unit suite (tens of objects).

## 9. Candidates for read-only test-tenant confirmation

Use only if documentation and mocks leave the behavior uncertain, and only after a test-tenant procedure is approved.

| Question | Why a mock may not be enough |
| --- | --- |
| Shape of the beta custom claims policy response for a portal-configured SAML application | Preview API; property names can change |
| Whether portal-configured claims appear as a claims mapping policy, a custom claims policy, or neither | Determines whether NameID is collectible at all |
| Where gallery SAML applications store the Entity ID (`servicePrincipalNames` or `identifierUris`) | Varies by application template |
| Synchronization schema shape and `matchingPriority` for a gallery SCIM application | Schema is large and template-specific |
| Whether `isMultipleDataLocationsForServicesEnabled` is returned without `$select` | Default property sets differ by resource |
| Naming pattern of directory synchronization service accounts for Entra Connect Sync and Cloud Sync | Inference rule in P1-06 depends on it |
| `domainNameReferences` paging and object types | Volume and type names |

## 10. Test results recorded for this review

| Suite | Command | Result on 2026-11-10 |
| --- | --- | --- |
| Pester (all containers) | `./tools/run-pester.ps1 -Path ./tests -SkipCodeCoverage` with results written outside the repository | 404 total, 403 passed, 0 failed, 1 skipped; 26 containers, 0 failed |
| Assessment review (Python) | `.venv\Scripts\python.exe -B -m unittest discover -s tests/assessment_review -t .` | 54 run, OK, 4 skipped |
| Customer deliverables (Python) | `.venv\Scripts\python.exe -B -m unittest discover -s tests -p test_customer_deliverables.py` | 19 run, OK |
| Defect reproduction (diagnostic) | Production functions against synthetic fixtures | R1 to R10 reproduced |

No tenant was contacted.
