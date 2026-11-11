# T2T P0 Implementation Status

Date: 2026-11-11
Branch: `feature/t2t-p0-correctness`, created from `main` at `66f48a8`
Related documents: [backlog](T2T-Migration-Assessment-Backlog.md), [coverage](T2T-Migration-Assessment-Coverage.md), [synthetic scenario spec](T2T-Synthetic-Scenario-Spec.md), [review summary](T2T-Review-Summary.md)

This records what the P0 correctness batch changed, how each change is tested, and what was deliberately left for later. It covers P0 only. No P1 capability was added.

## Scope

| Item | Status | Kind of change |
| --- | --- | --- |
| P0-01 AD Connect and Exchange hybrid reported as configured without evidence | Implemented | Analysis |
| P0-02 Unknown synchronization state reported as cloud-only | Implemented | Analysis, report |
| P0-03 Items without evidence scored Ready and hidden | Implemented, with one documented limit | Analysis, report |
| P0-04 Multi-Geo read from non-existent Graph properties | Implemented | Collection, report |
| P0-05 DNS failure and `mx.microsoft` classified as third-party routing | Implemented | Collection, analysis |
| P0-06 Profile contract | Implemented narrowly: documentation, contract tests, evidence-coverage rows. Profile expansion is not implemented (P1-01) | Documentation, test, report |
| P0-07 "Mailboxes on hold" reflected litigation hold only | Implemented | Analysis |

Not implemented on this branch, by design: the `TenantToTenantMigration` profile expansion, multi-profile composition, the source-versus-destination comparison, SAML and SCIM evidence, the licensing model, endpoint evidence, and any milestone or phase model.

## Shared interpretation logic

Five functions were added to `Arraya.M365.Common` so the evidence rules live in one place:

| Function | Returns |
| --- | --- |
| `Get-ArrayaDatasetEvidenceState` | `Collected`, `NotCollected`, `Failed`, or `Unknown` for a dataset, from its rows, the `CollectorPlanResults` step status, and presence in the store |
| `Get-ArrayaDirectorySyncEvidence` | `Synchronized`, `NotSynchronized`, or `Unknown` |
| `Get-ArrayaExchangeHybridEvidence` | `Configured`, `Indicators`, `NotConfigured`, or `Unknown` |
| `Get-ArrayaMailExchangerClassification` | `Microsoft365`, `ThirdParty`, `Mixed`, `NoMxRecord`, or `Unknown`, plus a true, false, or null `IsMicrosoft365` |
| `Get-ArrayaMailboxHoldEvidence` | Counts per mailbox-level hold indicator, and which indicators were and were not available |

Existing status values are reused: `Ready`, `Review`, `Blocker`, `Needs Data`, `Info`. No new status was introduced.

## Semantic corrections

### P0-01 Hybrid and directory synchronization

- Before: `Update-MigrationComplexityFlags` raised "AD Connect directory synchronization: Configured" and "Exchange hybrid: Configured" whenever the collector had written its container, which it always does.
- After: the flags come from the evidence inside the container.
  - Synchronization confirmed: `Review`, `Configured`.
  - Organization read succeeded and reports no synchronization: no flag.
  - State cannot be established: `Needs Data`, `Not established`.
  - Exchange hybrid confirmed: `Review`, `Configured`. Partial signals without confirmed hybrid: `Review`, `Indicators detected`. Explicit false with complete probes: no flag. Probes incomplete or collector not run: `Needs Data`.
- Also corrected: the per-domain `HybridRoutingReview` never saw the collector's verdict because it read the wrong key. Confirmed hybrid and hybrid indicators now mark domains for review, and an unknown hybrid state yields `Needs Data` instead of "No immediate signal".
- Cloud Sync is not distinguished from Entra Connect Sync by the collected evidence. The output says so. Detection is backlog item P1-06.

### P0-02 Unknown synchronization state

- Before: an empty flag, including one left by a failed organization read, produced "Cloud-only identity model, Ready" and the questionnaire answer "No active on-premises directory synchronization detected."
- After: cloud-only is reported only when the organization read succeeded and reports no synchronization, and no collected user object is synchronized. The read counts as successful when `TenantInfo`, which comes from the same organization call, carries a tenant id or display name. Otherwise the row is `Needs Data` with the value "Directory synchronization state not established" and the reason.
- No snapshot field was added and no new permission is needed.
- The questionnaire answers for sign-in provider, directory synchronization, and enterprise SSO now say "Not collected in this run" when the evidence is absent. The tenant overview section of the technical HTML shows "Not reported" for an empty flag instead of "No".

### P0-03 No evidence must not score Ready

- Before: a count of zero scored `Ready` whether or not anything was collected, and `Ready` rows are hidden from the cutover checklist, so uncollected items disappeared.
- After, in `CutoverPrepSummary`, `MigrationReadinessChecklist`, and `MigrationReadiness`: zero scores `Ready` only when the dataset's evidence state is `Collected`. Otherwise the row is `Needs Data`, stays visible, and carries the reason. This applies to domains, mail routing, mailbox inventory, forwarding, delegates, mail flow connectors, remote domains, public folders, hybrid indicators, inactive mailboxes, oversized mailboxes and archives, Teams shared channels, and SharePoint and OneDrive sizing.
- Domains and remote domains are treated as unknown when empty, because every tenant has at least one of each.
- A defect found while tracing: `@($null).Count` reported "1 confirmed relay connector" when no relay configuration was collected. Corrected.
- New "Evidence coverage" rows list migration-relevant collector steps that were skipped by the profile or failed, with status `Needs Data`.
- The questionnaire no longer prints zero counts for mailboxes, recipients, public folders, transport rules, SharePoint, OneDrive, devices, or guests that were not collected.
- Limit: a collector that catches its own errors still reports its step as Completed, so an internal failure that leaves an empty container is trusted as an empty result for public folders, connectors, and transport rules. Backlog item P1-19.

### P0-04 Multi-Geo

- Before: `Get-TenantOverviewInfo` read `IsMultipleGeolocationEnabled` and `MultiGeoConfiguration`, which are not properties of the Graph organization resource, so the value was always empty and the questionnaire said none was detected.
- After: it reads `isMultipleDataLocationsForServicesEnabled` and keeps true, false, and not set apart. The questionnaire answers "Yes" for true, "not Multi-Geo enabled" for false, "Not reported by Microsoft Graph ... confirm with the customer" when the flag is not set, and "Not collected" when the tenant overview is missing.
- `TenantInfo` keeps the same properties. `MultiGeoAllowed` and `MultiGeoCentral` stay empty because Graph does not expose them on this resource.
- To confirm in a test tenant when one is available: that the property is returned by the organization read without `$select`.

### P0-05 DNS and MX classification

- Before: one wildcard comparison on the joined MX names. A failed lookup and an `mx.microsoft` record both produced `Office365MailExchanger = False`.
- After: `Office365MailExchanger` is true when every MX host is an Exchange Online endpoint, false only when a well-formed non-Microsoft MX host is published (including mixed records), and empty when the lookup returned nothing, returned only malformed values, or the domain has no MX record. Hosts are matched by suffix on a label boundary, so a look-alike host is not accepted.
- Downstream: unknown MX yields `Needs Data` for "Mail routing" and for the per-domain `ThirdPartySpamFilterReview`. The domain table in the technical HTML shows "Unknown" instead of "No".
- The lookups still use `Resolve-DnsName` against a fixed public resolver. Outbound DNS is often blocked on customer networks, which now surfaces as unknown instead of a false third-party finding.

### P0-06 Profile contract

- `README.md` and `RUN.md` now describe `TenantToTenantMigration` as Exchange and collaboration cutover preparation, state what each migration profile collects, correct the `Presales` depth from Minimum to Operator, and document the limits of combining profiles.
- A contract test compares the `RUN.md` profile table with `Get-ArrayaAssessmentOutputProfilePolicy` and pins the current collector and depth flags, so the documents and the code cannot drift apart silently.
- The readiness checklist discloses skipped collectors (see P0-03).
- No collector was enabled, no profile was renamed or merged, and profile composition was not changed.

### P0-07 Mailbox holds

- Before: "Mailboxes on hold" counted `LitigationHoldEnabled` only and scored `Ready` at zero.
- After: it counts mailboxes with any mailbox-level indicator the collector already gathers: litigation hold, delay hold, and non-exclusion `InPlaceHolds` entries. `RetentionHoldEnabled` is reported separately because it suspends retention processing and is not a preservation hold. The value lists the breakdown; the notes list which fields were evaluated and which were not collected.
- The row is never `Ready`. Organization-wide retention policies, compliance tag holds, and eDiscovery case detail are not collected, so zero indicators does not establish that no hold applies. Zero indicators scores `Needs Data`.
- `InPlaceHolds` is collected only at the deepest mailbox detail level; when it is missing the notes say so.
- No Purview or eDiscovery collection was added. That is backlog item P1-14.

## Regression tests

All tests are in `tests/unit/T2TP0Correctness.Tests.ps1`. They load the production functions from the shipped scripts, use synthetic fixtures with reserved `.example` domains, and replace Microsoft Graph, Exchange, and DNS calls with local stubs. No test contacts a tenant or performs a DNS query.

| Area | Cases |
| --- | --- |
| Dataset evidence state | Rows present; present and empty; absent; step skipped; step failed; empty where empty is not possible; rows present although a step was skipped |
| Directory synchronization | Confirmed synchronization; confirmed no synchronization; explicit false; unknown; collector not run; collector failed; synchronized users as affirmative evidence |
| Exchange hybrid | Confirmed; explicit not configured; partial indicators; probes unanswered; collector not run |
| Complexity flags | Cloud-only without hybrid; both confirmed; neither established; hybrid probes incomplete |
| Readiness rows | Organization read failed through the production collector; affirmative cloud-only; synchronized; no domains; complete Microsoft-routed domains; unknown MX and third-party MX; nothing collected; collected and empty |
| Cutover summary and checklist | No Exchange evidence; collected and empty; collector failed with an empty container; skipped collectors listed |
| Multi-Geo | Enabled; explicitly not enabled; property absent; tenant overview not collected |
| Questionnaire | Nothing collected; evidence present |
| MX classification | Exchange Online Protection; `mx.microsoft`; trailing dot and mixed case; third-party; mixed; answered without MX; lookup failed; malformed; look-alike suffix; production domain collector with mocked DNS |
| Mailbox holds | All indicators; fields not collected; non-litigation holds in summary and checklist; litigation hold; no indicator; incomplete fields |
| Render-only | Cutover workbook sheet schema and statuses; cutover HTML checklist |
| Profile contract | Plan resolves for all six profiles; default profiles keep identity, application, group, and device collection; current tenant-to-tenant flags pinned; delegate collection scope; `RUN.md` depth matches code; documents no longer overstate the profile; composition limit documented |

Fail-first record: before any production change, 52 of the first 64 tests failed. The 12 that passed were positive cases that already held. Two render-only tests were added afterwards.

## Test results

| Suite | Before (at `66f48a8`) | After |
| --- | --- | --- |
| Pester, all containers | 404 total, 403 passed, 0 failed, 1 skipped | 470 total, 469 passed, 0 failed, 1 skipped |
| Assessment review (Python) | 54 run, OK, 4 skipped | 54 run, OK, 4 skipped |
| Customer deliverables (Python) | 19 run, OK | 19 run, OK |

The only existing test file that changed is `tests/unit/Common.Tests.ps1`, where the expected module export list gained the five new function names. No existing assertion was weakened.

No tenant was contacted. No customer data was used.

## Compatibility

| Surface | Result |
| --- | --- |
| Snapshot schema | Unchanged. No dataset or property was added or removed |
| Workbook schemas and worksheet names | Unchanged. The exporter, its column lists, and its sheet order were not edited; the existing workbook schema tests pass |
| Profile names and command-line parameters | Unchanged |
| `DecisionKey` behavior | Unchanged. Keys `FLAG:ad-connect-directory-synchronization` and `FLAG:exchange-hybrid` keep their names |
| Cutover pack CSV schemas | Unchanged. The cutover pack exporter was not edited |
| Module exports | Five functions added to `Arraya.M365.Common` |

Values that changed inside existing columns:

- `Domains.Office365MailExchanger` can now be empty (unknown) as well as true or false.
- `Domains.ThirdPartySpamFilterReview` and `Domains.HybridRoutingReview` can now be `Needs Data`.
- `Status` in `CutoverPrepSummary`, `MigrationReadinessChecklist`, `MigrationReadiness`, and `MigrationComplexityFlags` returns `Needs Data` in cases that returned `Ready`, `Info`, or `Review` before.
- `MigrationReadinessChecklist` gains "Evidence coverage" rows and keeps `Needs Data` rows that used to be hidden.
- `MigrationQuoteReadiness` gains a blocking `Needs Data` condition when synchronization or hybrid state cannot be established, and loses the two false conditions on cloud-only tenants.
- `TenantInfo.MultiGeoEnabled` is populated when Graph reports a value.
- Replaying an older snapshot recomputes these rows with the new rules.

## Deferred

| Topic | Where it is tracked |
| --- | --- |
| Expanding `TenantToTenantMigration` to collect application, group, device, and identity evidence | P1-01 |
| Multi-profile composition | P1-18 |
| Collectors that catch their own errors | P1-19 |
| Source-versus-destination comparison | P1-05 |
| Entra Connect versus Cloud Sync, server identity, staging mode | P1-06 |
| Compliance tag holds, organization-wide retention policies, eDiscovery cases | P1-14 |
| Domain evidence classes and removal dependencies | P1-15 |
| `New-MigrationReadinessChecklistRows` in the HTML helper, a fallback used only when a snapshot has no checklist, still scores from counts | With P1-12 |
| "Target licensing readiness" row reports source utilization | P1-07 |
| Cutover pack `Direction = MemberToGuest` preset | P1-11 |
| Mailbox routing guidance that assumes native cross-tenant migration | P1-13 |

## Live validation

None was run and none is required to accept this batch. One item is worth confirming in the approved internal test tenant, read-only, when a test-tenant procedure exists: that the organization read returns `isMultipleDataLocationsForServicesEnabled` without `$select`. If it does not, Multi-Geo stays "Not reported", which is safe but uninformative.
