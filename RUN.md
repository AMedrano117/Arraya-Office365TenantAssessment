# Run Modern Workplace Tenant Assessment

Use this guide when you are ready to execute the assessment and review the results. For repository background, see [README.md](README.md).

## Prerequisites

- PowerShell 7 or later
- A local clone of this repository
- Access to the Microsoft 365 tenant you want to assess

Install the required PowerShell modules:

```powershell
.\tools\install-dependencies.ps1 -Scope Runtime
```

## Authentication Options

The assessment supports three authentication modes:

| Mode | Best for | Notes |
| ---- | -------- | ----- |
| `Interactive` | Guided consultant runs | Default. Sign in when prompted. Gives the broadest workload coverage. |
| `Certificate` | Unattended or repeatable runs | Recommended for production use. Supports Graph, Exchange, and Purview app auth. |
| `ClientSecret` | Compatibility cases | Graph app auth works; Exchange falls back to delegated sign-in. Not recommended as a permanent solution. |

If you are setting up app-based auth for the first time, complete these runbooks before running:

- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)

### Useful auth flags

- `-SkipAuth` — reuse workload sessions already connected in the current PowerShell session
- `-SkipPermissionPreflight` — bypass the startup access check and let collection run on a best-effort basis

### Authentication order in Interactive mode

The assessment connects workloads in this order: `Exchange Online → Purview → Microsoft Graph`. If a previous Graph session has already been established in the same shell, close that session and start fresh before retrying.

### Tenant validation

Microsoft Graph reuses whatever `Connect-MgGraph` session is already cached in the shell, and that session can belong to a different tenant than the one you are assessing. Two checks guard against a mixed-tenant report:

- Interactive runs prompt for the tenant when `-TenantId` is not supplied, then validate the Graph context against it. Pass `-TenantId '<tenant-guid>'` to skip the prompt. Non-interactive hosts cannot be prompted, so they require the flag.
- After both workloads connect, Graph and Exchange Online tenant ids are compared. A mismatch stops the run before any collection, because every Graph-derived dataset would otherwise belong to a different tenant than the Exchange data.

If a mismatch is reported, run `Disconnect-MgGraph`, start a fresh PowerShell session, and reconnect to the intended tenant.

## Required Access and Permissions

### Operator roles (Interactive mode)

For full coverage, the operator account should have:

- `Exchange Administrator`
- `SharePoint Administrator`
- `Teams Administrator`
- `Security Reader`
- `Reports Reader`
- `Directory Readers` or `Global Reader`

### Microsoft Graph application permissions (app-based auth)

Grant these core permissions and consent them on the app registration:

- `Application.Read.All`
- `AuditLog.Read.All`
- `Channel.ReadBasic.All`
- `CrossTenantInformation.ReadBasic.All`
- `Device.Read.All`
- `Domain.Read.All`
- `Group.Read.All`
- `GroupMember.Read.All`
- `OnPremDirectorySynchronization.Read.All`
- `Organization.Read.All`
- `Policy.Read.All`
- `Reports.Read.All`
- `ReportSettings.Read.All`
- `RoleManagement.Read.All`
- `SecurityEvents.Read.All`
- `SharePointTenantSettings.Read.All`
- `Sites.Read.All`
- `Team.ReadBasic.All`
- `User.Read.All`

Optional — add these only when you need deeper Teams membership detail:

- `TeamMember.Read.All`

Optional extended enrichment (add only when explicitly needed):

- `DeviceManagementApps.Read.All`
- `DeviceManagementConfiguration.Read.All`
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementRBAC.Read.All`
- `DeviceManagementScripts.Read.All`
- `DeviceManagementServiceConfig.Read.All`
- `DirectoryRecommendations.Read.All`
- `IdentityRiskEvent.Read.All`
- `IdentityRiskyUser.Read.All`
- `LicenseAssignment.Read.All`
- `ServiceMessage.Read.All`
- `User.Export.All`
- `UserAuthenticationMethod.Read.All`

### Access outside Microsoft Graph

- **Exchange Online**: requires Exchange app-only access configured on the app registration for certificate auth
- **Purview**: uses `Connect-IPPSSession` from `ExchangeOnlineManagement`; supports delegated and certificate auth; client-secret auth is not supported
- **SharePoint**: certificate and client-secret runs use Microsoft Graph for SharePoint and OneDrive collection; SPO admin PowerShell is only used in interactive runs. Site inventory comes from Graph `getAllSites`, which Microsoft supports with **application permissions only**; delegated access is not supported at any scope. An interactive run therefore always sees a 403 from that endpoint and falls back to the SharePoint Online PowerShell module by design; this is expected and no Graph scope changes it. The SharePoint and OneDrive usage and activity reports are collected from the Graph reporting endpoints either way, so they populate even when site inventory falls back
- **Teams**: Teams PowerShell is delegated-only; member and guest count enrichment requires `TeamMember.Read.All`

## Common Commands

Run the interactive menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Run the standard full assessment:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full
```

Run with certificate auth:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

Check permissions before a long collection run:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Preflight `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

Collect data only (no report yet):

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Collect `
  -OutputProfile SolutionsEngineer `
  -ExportPath 'C:\Assessment-Outputs'
```

Generate reports from an existing snapshot:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Report
```

Generate the improvement plan from an existing snapshot:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Improve
```

Compare two saved snapshots:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Compare
```

Run without the improvement plan step:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -SkipImprove
```

Run a full assessment and write outputs to a specific folder:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -ExportPath 'C:\Assessment-Outputs'
```

Run with a specific output profile:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -OutputProfile SolutionsEngineer,ExecutiveLevel
```

### Prompts for some actions

- `Report` — asks for the saved JSON snapshot path
- `Improve` — asks for the snapshot or manifest JSON path and optional output folder
- `Compare` — asks for baseline and current snapshot paths

## Output Profiles

| Profile | Depth | Collection scope | Best for |
| ------- | ----- | ---------------- | -------- |
| `SolutionsEngineer` | Operator | Default: every collector | Default. Standard consultant depth. |
| `ExecutiveLevel` | Minimum | Default, without recipients, Exchange groups, mail flow, public folders, SMTP relay, and Microsoft 365 Groups | Leadership summaries. |
| `Presales` | Operator | Presales: migration scoping | Migration scoping, sizing, complexity, and quote readiness. |
| `TenantToTenantMigration` | All | TenantToTenantCutover: Exchange and collaboration cutover | Mailbox cutover preparation and the cutover pack. Not yet a complete migration assessment. |
| `Geek` | Geek | Default: every collector | Detailed engineer troubleshooting. |
| `Machine` | Automation | Default: every collector | JSON-first automation and replay. |

Depth is the `ReportingMode` returned by `Get-ArrayaAssessmentOutputProfilePolicy`. Collection scope decides which collectors run. A unit test compares this table with the code.

### What the two migration profiles collect

`Presales` and `TenantToTenantMigration` are separate profiles with different purposes. Neither is a superset of the other today.

| Evidence | `Presales` | `TenantToTenantMigration` |
| -------- | ---------- | ------------------------- |
| Mailboxes, sizes, archives, inactive mailboxes | Yes | Yes |
| Full Access, Send As, and calendar folder permissions | No | Yes (only this profile collects them) |
| Domains, connectors, remote domains, SMTP relay, public folders, hybrid signals | Yes | Yes |
| Teams, SharePoint, and OneDrive inventories | Yes, with full SharePoint detail and Microsoft 365 Group mailbox statistics | Yes, without full SharePoint detail or Microsoft 365 Group mailbox statistics |
| Enterprise applications and SSO | Yes | No |
| Microsoft 365 Groups inventory | Yes | No |
| Devices | Yes | No |
| Entra groups | Yes | No |
| Administrators, Conditional Access, MFA registration, Purview retention and DLP | No | No |
| Migration planning registers (decisions, identity mapping, domain dependencies, object disposition, effort, waves) | Workbook and scoping brief | JSON snapshot only |
| Cutover pack, cutover checklist, delegate worksheets | No | Yes |

When a collector is skipped or fails, the tenant-to-tenant readiness checklist lists it under "Evidence coverage" with status `Needs Data`. Evidence that was not collected is never reported as absent or ready.

`TenantToTenantMigration` does not yet represent all the evidence needed to plan a tenant-to-tenant migration. The gaps and the planned expansion are recorded in [T2T-Migration-Assessment-Coverage.md](docs/t2t-migration-review/T2T-Migration-Assessment-Coverage.md) and [T2T-Migration-Assessment-Backlog.md](docs/t2t-migration-review/T2T-Migration-Assessment-Backlog.md). Until that work lands, run `TenantToTenantMigration` for cutover preparation and a second, separate run of `Presales` or `SolutionsEngineer` for application, group, and device evidence.

### Combining profiles

Pass multiple profiles as a comma-separated list. The assessment then runs once at the highest required depth, with these current limitations:

- The collection scope and the workbook layout reset to `Default`. Profile-specific scopes (`Presales`, `TenantToTenantCutover`) are not applied.
- When `TenantToTenantMigration` is combined with another profile, the tenant-to-tenant shaping does not run, mailbox and calendar delegate permissions are not collected, and the migration planning registers are not built. The cutover pack is still written, from unshaped data.
- A combined run can therefore produce less migration evidence than the same profiles run separately.

Run the migration profiles on their own until profile composition is redesigned. The intended rule, tracked in the backlog, is that requesting several profiles never yields less evidence than any one of them alone.

## Default Output Location

If you do not pass `-ExportPath`, outputs are written to:

```text
%LOCALAPPDATA%\Arraya\M365TenantAssessment\Outputs
```

## What Gets Generated

A standard `Full` run writes files into two folders.

`Deliverables` — files to share or review first:

- `*-Microsoft 365 Tenant Best Practices Assessment-<date>.docx`
- `*-Microsoft 365 Remediation Roadmap-<date>.docx`
- `*-EngPack.md`
- `*.xlsx` when the selected profile enables workbook output

`Support` — machine-readable and troubleshooting files:

- `*-AssessmentSnapshot.json`
- `*-SolutionsEngineerEvidenceCoverage.json`
- `*-Plan.json`
- `*.manifest.json`
- `*-Snips.ps1`
- `Debugging\*`

To include older HTML, PDF, and questionnaire artifacts, add `-IncludeLegacyAssessmentArtifacts`.

### Workbook worksheets

The workbook only exports columns that carry data. Columns empty for every row in the tenant are dropped, and the removals are listed in the run log, so the same worksheet can be narrower for one tenant than another.

- `Users` carries a fixed, assessment-focused column set in reading order: identity, account state, org placement, licensing, hybrid sync, then sign-in activity. `UserFullDetails` adds mailbox size and item counts, archive totals, and the OneDrive URL and storage. Other Microsoft Graph attributes are intentionally not exported.
- `SharePoint` and `OneDrive` merge the Graph usage reports into the site inventory. Alongside storage they carry `FileCount`, `ActiveFileCount`, `PageViewCount`, `LastActivityDate`, `ActivityState` (Active, Low, Stale, Dormant), `GeoLocation`, `SensitivityLabelId`, `OwnerSource`, and a `MigrationNotes` column flagging dormant content, unresolved owners, sensitivity labels, multi-geo placement, locked or archived sites, quota pressure, and Teams channel sites that cannot migrate independently.
- `SharePointActivityUserDetail` and `OneDriveActivityUserDetail` hold the per-user Graph activity reports. They come from the reporting endpoints, so they populate even when site inventory falls back to the SharePoint Online module.
- `AdConnectConfiguration`, `TeamsVoice`, and `AuthenticationConfig` are flattened to `Section`, `Item`, `Value`, `Notes` with one row per underlying record, so SSO applications, calling policies, and sync services are readable instead of collapsed into a single cell.
- `MfaEnrollmentSummary` reports one row per authentication method with a user count and percentage, rather than a single joined breakdown string.

## How To Read the Results

Start with the artifact that best matches your audience:

| Artifact | Audience |
| -------- | -------- |
| `Deliverables\*-Best Practices Assessment-*.docx` | Customer |
| `Deliverables\*-Remediation Roadmap-*.docx` | Customer or leadership |
| `Deliverables\*-EngPack.md` | Engineer or consultant |
| `Support\*-Plan.json` | Automation, filtering, downstream processing |

Recommended review order:

1. Open the customer assessment DOCX in `Deliverables` first.
2. Review the roadmap DOCX and engineer pack for implementation planning.
3. Keep `Support\*-Plan.json` as your baseline for comparison or downstream use.

## Validating on a New Machine

If you need to validate the workflow on another workstation or jump box:

1. Install PowerShell 7 or later.
2. Clone or copy the repo.
3. Install required modules: `.\tools\install-dependencies.ps1 -Scope Runtime`
4. If using certificate auth, import the `.pfx` and verify it has a private key.
5. Start with the lowest-risk path:
   - Run `Improve` against an existing known-good snapshot first.
   - Then run `Collect`.
   - Then run `Improve` against the new snapshot.

This separates environment setup issues from tenant authentication issues.

## Troubleshooting

**Missing module**: rerun `.\tools\install-dependencies.ps1 -Scope Runtime`

**Unexpected auth prompt**: check whether `-SkipAuth` should be used to reuse an existing session.

**Exchange or Purview auth fails after Graph**: close the PowerShell window and start a fresh session. The assessment authenticates `Exchange Online → Purview → Microsoft Graph` by design, but if Graph already ran in the shell, Exchange-family auth in that same session may fail.

**SharePoint fails before opening a sign-in window**: verify the admin URL in the console. A renamed SharePoint hostname can differ from the original `onmicrosoft.com` tenant domain. The assessment discovers the hostname using Graph `GET /v1.0/sites/root` and only falls back to the initial domain if discovery fails. Connect to the actual admin URL with `Connect-SPOService -Url 'https://<actual-sharepoint-name>-admin.sharepoint.com' -UseSystemBrowser $true`, using a SharePoint Administrator account. Graph sign-in is separate from SPO sign-in; Graph `getAllSites` requires application permissions in both v1.0 and beta, so switching to beta does not enable delegated tenant-wide inventory.

**Run stops at the permission gate**: add `-SkipPermissionPreflight` to let collection continue on a best-effort basis.

**PDF output missing**: PDF generation requires a locally installed Chromium-based browser (Google Chrome or Microsoft Edge). The assessment still completes successfully without it.

**App-only auth fails**: review [App Registration Setup](docs/runbooks/app-registration-setup.md) and [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md).

## Additional References

- [Repository Quick Start](README.md#quick-start)
- [Customer Execution Checklist](docs/runbooks/customer-execution-checklist.md)
- [Improvement Plan Rule Taxonomy](docs/runbooks/improvement-plan-rule-taxonomy.md)
