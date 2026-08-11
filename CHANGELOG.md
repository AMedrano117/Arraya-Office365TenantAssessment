# Changelog

All notable changes to this repository are tracked here. The detailed commit ledger at the bottom preserves the historical commit subjects that predate this changelog.

## Unreleased

### Collector Runtime And Performance
- Added a shared collector runtime context, run-scoped cache helpers, and collector step result tracking for clearer live collection diagnostics.
- Replaced the hardcoded live collector sequence with a plan-driven six-section flow: Tenant Overview, Identity, Exchange, Collaboration, Endpoint, and Governance.
- Added cache-first collection paths for repeated Graph, Exchange, and report CSV calls, including mailbox stats, unified groups, report lookups, service principal/resource lookups, and tenant metadata.
- Routed report and high-volume Graph calls through shared collection helpers to avoid duplicate requests within a run.
- Cached collector report lookups and reused populated runtime caches during snapshot creation/export.
- Improved Exchange and unified group performance by preferring Graph/report data first and falling back to EXO only for unresolved records.
- Added collector diagnostics for cache/request behavior, produced row counts, mailbox stats source breakdowns, and optional/fallback workload states.

### Modular Run Modes
- Added `Invoke-M365TenantConnectionPreflight` for connection and permission validation only.
- Added `Invoke-M365TenantDataCollection -UseExistingConnections` for data-only collection from already established sessions.
- Added export-only routing through `Invoke-M365TenantAssessmentExport -AssessmentJsonPath`.
- Preserved existing defaults for `M365`, `M365Collect`, `Invoke-M365TenantAssessment`, and `Invoke-M365TenantDataCollection`.
- Added internal targeted collector switches for development and live validation:
  - `-CollectorSection`
  - `-CollectorStep`
- Recalculated targeted-run progress totals from the filtered collector plan and emitted only matching section headers.

### Authentication And Preflight
- Added staged workload connection/preflight behavior for Graph, Exchange Online, Purview compliance, SharePoint admin, and Teams PowerShell.
- Improved interactive-auth console output with clearer tenant/session summaries and less noisy progress text.
- Added existing-session tenant confirmation and session reuse validation for Graph, Exchange, Purview, SharePoint, and Teams.
- Added support for declining existing-session reuse and disconnecting current sessions.
- Fixed delegated SharePoint Graph preflight and clarified Graph/SPO fallback behavior.
- Added clearer Graph preflight guidance for missing delegated scopes and cached-token reuse.
- Added app-auth/certificate-aware Graph permission reuse validation so app-only permissions such as `Directory.Read.All` can satisfy equivalent live probes.
- Verified certificate-based auth for separate preflight, targeted collect, existing-session collect, and export-only flows.
- Documented certificate auth behavior and app-based workload limitations in operator docs.
- Added a cross-workload tenant guard that stops the run when Microsoft Graph and Exchange Online are connected to different tenants, which previously produced a workbook mixing two tenants' data without warning.
- Interactive runs now obtain a tenant id before collection, prompting when `-TenantId` is not supplied and failing fast on non-interactive hosts, so the cached `Connect-MgGraph` session is always validated against the intended tenant.

### Reporting, Export, And Deliverables
- Added DOCX customer reporting, remediation roadmap output, and engineer action pack improvements.
- Expanded Solutions Engineer evidence coverage and surfaced coverage artifacts in operator outputs.
- Improved customer report readability, application insights, MFA/storage visuals, and roadmap timing.
- Expanded external exposure, guest access, ownership governance, licensing governance, and full-assessment governance reporting.
- Normalized output layout into `Deliverables` and `Support` paths, including manifests, snapshots, evidence coverage, debugging logs, and optional legacy artifacts.
- Improved workbook/export memory behavior and reused cached mailbox/group data during export.
- Added snapshot replay/export improvements and hardened JSON snapshot compatibility.
- Fixed worksheets that reflected .NET collections into columns, which surfaced array members such as `Length`, `Rank`, and `SyncRoot` as headers instead of data.
- Flattened container worksheets (`AdConnectConfiguration`, `TeamsVoice`, `AuthenticationConfig`) to `Section`/`Item`/`Value`/`Notes` with one row per record, so SSO applications, calling policies, phone numbers, and sync services are no longer collapsed into a single cell or dropped.
- Expanded `MfaEnrollmentSummary` to one row per authentication method with user counts and percentages, replacing the pre-joined breakdown string.
- Registered `TeamsVoice` in the worksheet ordering instead of letting it fall through to the append-remaining branch.
- Padded ragged records to the union of their keys, since `Export-Excel` derives columns from the first record only and silently dropped columns a later record added.
- Suppressed cells that contained only a .NET type name, which filled roughly 50 user columns with `Microsoft.Graph.PowerShell.Models.*` and no data.
- Dropped worksheet columns that are empty for every row, logging the removals per sheet. Skipped for the tenant-to-tenant explicit column contract and the flattened configuration sheets.
- Added a curated, ordered column set for `Users` and `UserFullDetails`, replacing roughly 145 reflected columns that buried `DisplayName` behind alphabetical navigation properties.
- Preserved leading `+` on phone numbers in flattened configuration sheets by opting those columns out of Excel number coercion.

### Tenant-To-Tenant And Migration Outputs
- Added and refined tenant-to-tenant migration readiness outputs, cutover checklist data, mailbox migration columns, and reduced-scope T2T reporting.
- Added migration-focused workbook/export behavior and clearer not-collected guidance for reduced profiles.

### Governance, Identity, And Security Collection
- Expanded MFA registration/enforcement reporting, admin MFA assessments, Conditional Access role expansion, and guest MFA reporting.
- Improved authentication/SSO, enterprise application inventory, application sign-in enrichment, redirect-risk analysis, and app ownership reporting.
- Added compliance retention/DLP insights and hardened Purview compliance collection.
- Improved Secure Score, device, SharePoint/OneDrive, external sharing, guest access, ownership, forwarding, inbox rule, and public folder review outputs.
- Fixed user collection requesting no `$select` at the `geek` and `all` detail levels, which returned only Microsoft Graph's 11 default properties and left account state, licensing, hybrid sync, and sign-in columns empty. The most detailed profiles were producing the least data.
- Added `-PageSize 500` to user collection, which Microsoft requires whenever `signInActivity` is selected.
- Made `signInActivity` degrade gracefully: a missing Entra ID P1/P2 licence or `AuditLog.Read.All` now costs the sign-in columns instead of the whole user inventory.
- Surfaced the licensed/unlicensed split on the console, including an explicit warning when no licensed users resolve.
- Restructured SharePoint/OneDrive site collection into an ordered cascade (Graph SDK, then Graph `getAllSites`, then the SharePoint Online module), where each stage reports success and any failure falls through rather than only 401/403.
- Made the SharePoint Online fallback establish its own session, resolving the admin URL from Graph and prompting when unavailable, instead of requiring a session to already exist.
- Collected all four SharePoint and OneDrive usage and activity reports at D180 through `Get-GraphAPIActivityReport`, replacing two reports pulled at D7 through direct URIs.
- Added `SharePointActivityUserDetail` and `OneDriveActivityUserDetail` worksheets and rolled their aggregates into `CollaborationActivitySummary`.
- Joined the usage reports on the SharePoint Online fallback path, which previously downloaded them and then discarded them, and enriched site rows with file counts, page views, activity state, geo location, sensitivity label, and migration notes.
- Backfilled site owners for group-connected sites from the Microsoft 365 group inventory, recording provenance in `OwnerSource`.
- Clarified that Graph `getAllSites` is application-permission only, so the 403 on interactive runs is expected and logged as informational rather than a warning.

### Test And Maintenance
- Added focused routing tests for preflight-only, collect-only, export-only, existing-connection mode, and targeted collector filters.
- Added source-level tests for collector plan filtering, permission preflight behavior, app-auth fallback behavior, and output compatibility.
- Expanded unit coverage across runner, common, reporting, improvement, evidence coverage, and collector modules.
- Updated `.gitignore` for local validation artifacts.
- Refactored module import behavior and suppressed noisy module import warnings.

## Historical Commit Ledger

### 2026-05
- `a1213c4` - Fix certificate Graph permission reuse validation
- `00c3fa1` - Add targeted collector run filters
- `f91df4d` - Add non-live run mode routing tests
- `10a1c19` - Add existing-connection collection mode
- `522eb53` - Clarify interactive auth console output
- `8ebd01a` - Fix delegated SharePoint Graph preflight
- `d0cc04d` - Quiet interactive auth preflight output
- `227b801` - Improve Graph preflight auth guidance
- `c90a344` - Cache collector report lookups
- `3a53b59` - Route report Graph calls through shared collection helper
- `718ba8a` - Polish customer assessment deliverables
- `53d9641` - Surface Solutions Engineer evidence coverage in operator outputs
- `bf5a334` - Add Solutions Engineer evidence coverage matrix
- `25eda54` - Improve licensing governance reporting
- `74be0bd` - Expand full assessment governance reporting
- `874dfb9` - Improve roadmap readability
- `ab481b6` - Refine roadmap timing and customer document versioning
- `01180f7` - Add remediation roadmap companion output
- `71f05a9` - Add MFA and storage visuals to customer report
- `08e5841` - Refine customer report readability
- `60bed4a` - Enhance customer report application insights

### 2026-04
- `8573cfe` - Ignore local test run artifacts
- `983d8b9` - Refine tenant-to-tenant migration outputs
- `52c60ed` - Improve MFA enforcement reporting and CA role expansion
- `3f929d0` - Refactor module import statements to suppress warnings
- `01a496b` - Add M365 preflight mode and reorder assessment flow
- `94b0e98` - Fix staged preflight workload selection
- `e1cd11d` - Stage workload preflight and modernize console output
- `b516ce5` - Enhance MFA reporting and admin assessments
- `ce28bba` - Refine guest MFA reporting and governance wording
- `2614a2e` - Enhance preflight checks for Purview compliance cmdlets and add detailed error guidance for tenant-specific access issues
- `1519078` - Trim default Graph permission guidance and scope requests
- `8433813` - Speed up preflight checks and quiet Exchange reuse
- `33957cc` - Fix preflight Graph wrapper warning suppression
- `fc287f3` - Improve preflight diagnostics and app-auth startup behavior
- `566f5fb` - Document and harden Purview compliance collection
- `d85a3d8` - Fix same-session module reuse during assessment startup
- `7ccb4bf` - Add permission preflight bypass and document operator usage
- `6d78449` - Improve authentication section and app inventory reporting
- `5b76f87` - Fix MFA coverage scope analysis and preflight progress
- `c21ec14` - Improve consultative reporting and assessment coverage
- `0df4862` - Improve customer reporting and external exposure assessment
- `9c51eb5` - Adopt DOCX customer reporting and expand exposure data
- `1ba01a1` - Simplify report layout and remediation wording

### 2026-03
- `18c0bad` - Fix operator mode compatibility and SPO cert fallback
- `d8a9f7c` - Rebalance assessment depth and workbook defaults
- `ad12d1c` - Quiet inbox rule warnings and refresh HTML branding
- `83f282f` - Quiet Graph progress and tighten engineer pack
- `b7b3cbc` - Simplify operator outputs and evidence paths
- `d3024f0` - Consolidate remediation outputs and refresh wording
- `efc44af` - Simplify default assessment remediation workflow
- `3fd39ac` - Refine improve findings output
- `f743436` - Simplify improve launcher flow
- `8ac98bd` - Refactor snapshot workflows and expand improvement planning
- `aeacef0` - Add operator run guide and permissions details
- `61be13c` - Harden Exchange auth fallback
- `c74b288` - Add new functions for Office365 management and migration tasks
- `e49e2fa` - Improve auth flow and report generation
- `84bb707` - Organize assessment debug artifacts into dedicated folders
- `e4b8149` - Refactor assessment modules and normalize report output
- `c5d12f4` - Add compliance retention insights and modular activity analysis helpers
- `875b659` - Refactor assessment snapshot/export pipeline and fix HTML/export gaps
- `361c634` - Unify multi-profile runs and add mailbox stats source diagnostics
- `d8f85b0` - Use Graph-first unified group mailbox pre-cache with EXO fallback
- `01a2dd5` - Optimize combined Exchange group metadata reuse and SPO usage coverage
- `54ea859` - Optimize mailbox stats reuse and harden tenant stats state handling
- `f15eb06` - Optimize combined Exchange mailbox payload and public folder permission collection
- `6c5dae9` - Optimize combined user and mailbox consolidation path
- `6192754` - Avoid redundant unified-group mailbox stat pre-cache in combined runs
- `0522a25` - Optimize profile-gated post-processing and export snapshot creation
- `8620a83` - Refine assessment profiles, ownership governance, and export cleanup
- `813409f` - Refactor report HTML around KPI, reality gap, and roadmap
- `f2036a4` - Remove Methodology section from Tenant Assessment HTML Report
- `45f4735` - Expand friendly license mappings for Windows 365 and Windows 10/11 SKUs
- `61fb4e6` - Improve export memory usage and reuse cached mailbox/group data
- `b4d0a96` - Optimize tenant collection path and reduce Graph/EXO overhead
- `34b7277` - Fix SKU friendly-name resolution order and validate license output
- `466bd55` - Optimize cache-first collectors and source license names from Entra reference
- `63cc541` - Optimize unified group mailbox stats retrieval for minimum mode
- `bee22af` - Fix SharePoint size rendering and simplify Secure Score KPI output
- `ea5e485` - Refactor assessment collectors and add shared runtime helpers
- `dc1aa63` - Enhance progress reporting and error handling in Graph request functions
- `a5a33a4` - Enhance error handling and progress reporting in assessment scripts
- `00e22b6` - Fix Graph auth strict-mode regressions in assessment flow
- `7ae01f4` - Refactor module imports to ensure required commands are available and improve error handling
- `6de2cfc` - Consolidate assessment Graph helpers and centralize error export
- `b407932` - Update README without local details
- `20ea0b5` - Fix cert auth and simplify assessment prompts
- `6502384` - Update `.gitignore` and README
- `6a25934` - Clean up artifacts
- `a4f3a41` - Refactor code structure for improved readability and maintainability
- `b442d7f` - Refine assessment outputs and auth options
- `144717b` - Add lean/standard/full output profiles, clean up minimum-mode workbook tabs, and generate a dedicated assessment-only HTML report

### 2026-02
- `460d00f` - Add new functions for managing domain migrations and user updates
- `0e9aa0c` - Initial commit for Microsoft 365 Tenant Automation project
- `5e10d19` - Added README.md

## 0.1.0
- Initial Azure DevOps repo scaffold
