# Run Modern Workplace Tenant Assessment

Use this guide when you are ready to execute the assessment and review the results. For repository background and module layout, start with [README.md](README.md).

## Prerequisites

- PowerShell 7 or later installed.
- A local clone of this repository.
- Access to the Microsoft 365 tenant you want to assess.
- The required PowerShell modules installed for the services you plan to query.

Install the user-facing module prerequisites:

```powershell
.\tools\install-microsoft-modules.ps1
```

If you also want the test and lint tooling used by maintainers:

```powershell
.\tools\install-microsoft-modules.ps1 -IncludeDevTools
```

## Authentication Options

The assessment supports three authentication modes:

- `Interactive`: default sign-in flow for guided, delegated access.
- `Certificate`: app-based authentication using a certificate.
- `ClientSecret`: app-based authentication using a client secret.

Coverage is not identical across those modes:

- `Interactive`: best choice for the fullest workload coverage.
- `Certificate`: supports Microsoft Graph, Exchange Online app auth, and Purview certificate auth; SharePoint admin PowerShell is skipped by design and falls back to Microsoft Graph collection.
- `ClientSecret`: Microsoft Graph supports app auth, Exchange falls back to delegated sign-in in this workflow, SharePoint admin cmdlets are skipped, and Teams PowerShell is skipped.

If you are using app-based authentication, complete the setup guidance first:

- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)

The assessment now uses an assessment-owned, staged login flow instead of a generic "connect everything" bootstrap. Graph and Exchange are the baseline live-collection workloads, Purview is only connected when the active run needs governance policy collection, and SharePoint/Teams PowerShell are attempted only when the active run can use them.

If you already connected the required workloads for the active run in the same PowerShell session, you can reuse those sessions with `-SkipAuth`.
If the active profile includes Purview retention or DLP collection, `-SkipAuth` requires a usable existing compliance PowerShell session, not only the presence of Purview cmdlet names in scope.
The assessment now performs staged workload preflight checks during connection startup: Graph checks after Graph connects, Exchange checks after Exchange connects, and Purview checks after the compliance session connects.
Use `-SkipPermissionPreflight` if you want to bypass those staged startup checks and let the run continue until a later collector hits missing access.
The launcher also reuses the repo modules already loaded from this repository in the same PowerShell session, so repeat runs should not keep re-importing the assessment modules.
If you want the assessment to continue without the startup permission gate, you can add `-SkipPermissionPreflight`. This skips the initial required-access validation and allows collection to continue on a best-effort basis, so missing permissions may still show up later as workload-specific warnings or failures.
The permission preflight now shows the specific check being evaluated, then prints a short summary of how many checks succeeded and how many remain. It uses the current Graph token claims to fast-pass the clear-cut permissions and keeps live endpoint probes for the more ambiguous checks, so startup validation is usually faster while still catching real workload-specific access gaps. Non-blocking checks such as `OnPremDirectorySynchronization.Read.All` are surfaced as structured preflight warnings so the operator can see the remaining gap without relying on a raw Graph 403 line.
Assessment progress output now uses plain-language governance step names instead of older internal `Tier B` terminology.

Purview retention and DLP policy collection is a separate compliance PowerShell surface in this workflow. It uses `Connect-IPPSSession` and the compliance cmdlets `Get-RetentionCompliancePolicy` and `Get-DlpCompliancePolicy` rather than the main Graph collector path.
In `Interactive` mode, the assessment now intentionally retries the Purview sign-in flow with device code if the first interactive attempt does not complete cleanly, so the operator can still complete the compliance login in sessions where browser-based auth is blocked or unstable.
If that connection fails during permission preflight, the run now reports the auth path used, the tenant organization value when applicable, and a next-step message so the operator can tell whether the problem is missing module availability, unsupported auth mode, certificate/app access, or missing compliance cmdlets after connect.
The same certificate and app registration can also behave differently across tenants: one tenant may expose the Purview retention/DLP cmdlets to the app session while another rejects that feature surface. When the compliance endpoint accepts the certificate sign-in but returns "No cmdlet assigned to the user have this feature enabled," treat it as a tenant-specific Purview/compliance access or licensing gap rather than a generic Graph auth failure.

## Required Access And API Permissions

The current app registration and consent baseline is documented here from the active collector set used by this repo. Treat this as the operator source of truth for app-based auth.

### Recommended Operator Access

For a reliable full `M365` run in delegated interactive mode, the operator account should be able to connect to:

- Microsoft Graph
- Exchange Online PowerShell
- SharePoint Online Admin PowerShell when delegated SharePoint admin connectivity is desired for the active run
- Microsoft Teams PowerShell

The practical role set for full coverage is:

- `Exchange Administrator`
- `SharePoint Administrator`
- `Teams Administrator`
- `Security Reader`
- `Reports Reader`
- `Directory Readers` or `Global Reader`

If you are setting up app-based authentication, you will also need enough Entra permission to create and consent the app registration, typically:

- `Application Administrator`
- `Cloud Application Administrator`
- or `Global Administrator`

### Microsoft Graph Delegated Scopes

When you run the main assessment in `Interactive` mode, the connector now requests the core delegated scope set used by the standard assessment:

- `Organization.Read.All`
- `User.Read.All`
- `AuditLog.Read.All`
- `Group.Read.All`
- `GroupMember.Read.All`
- `RoleManagement.Read.Directory`
- `Domain.Read.All`
- `Device.Read.All`
- `Reports.Read.All`
- `ReportSettings.Read.All`
- `Policy.Read.All`
- `CrossTenantInformation.ReadBasic.All`
- `SecurityEvents.Read.All`
- `Application.Read.All`
- `Sites.Read.All`
- `SharePointTenantSettings.Read.All`
- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`
- `OnPremDirectorySynchronization.Read.All`

These scopes support the repo's current Graph-based collection for tenant, identity, reporting, security, collaboration, and SharePoint data.
They are meant to cover the standard assessment path, not every possible legacy helper or optional enrichment path in the repo.

### App-Based Permission Baseline

For `Certificate` or `ClientSecret` mode, the app registration should be granted this core Microsoft Graph application-permission baseline for the standard assessment:

- `Application.Read.All`
- `AuditLog.Read.All`
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
- `Channel.ReadBasic.All`
- `Team.ReadBasic.All`
- `User.Read.All`

Optional extended enrichment permissions can be added when you explicitly want broader coverage beyond the standard assessment path:

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

Delegated baseline:

- `User.Read`

For app-based operation beyond Graph:

- Exchange Online requires Exchange app-only access for the app registration when using `Certificate` mode.
- SharePoint Online certificate and client-secret runs skip `Connect-SPOService` and rely on Microsoft Graph collection instead of SPO admin cmdlets.
- Purview retention and DLP policy collection uses `Connect-IPPSSession` from `ExchangeOnlineManagement`.
- Purview compliance PowerShell supports delegated auth and certificate auth in this workflow.
- Delegated Purview auth now retries with device code in `Interactive` mode if the initial interactive prompt fails.
- Purview compliance PowerShell does not support client-secret auth in this workflow.
- Teams PowerShell in this workflow is still delegated-only, so app-based runs may have reduced Teams detail.
- Teams member and guest counts are optional enrichment in app-based Graph runs. If `TeamMember.Read.All` or `TeamMember.ReadWrite.All` is not granted, the assessment continues with team and channel inventory only.

### Workload Notes

- Exchange Online connectivity is required for the full `M365` assessment path. If Exchange auth fails, the run stops.
- SharePoint Online admin connectivity is optional and only attempted for delegated runs where the active assessment path can use SPO cmdlets. Certificate and client-secret runs skip the SharePoint module import and use Microsoft Graph for SharePoint and OneDrive collection.
- Teams PowerShell connectivity is optional and only attempted for delegated runs where the active assessment path can use it; Graph remains the primary inventory source.
- Purview retention/DLP collection depends on compliance PowerShell cmdlets being available after `Connect-IPPSSession`. If that connection cannot be established, governance/compliance policy sections will be skipped or can stop the run during permission preflight.
- Teams connectivity is non-blocking in app-based auth modes, but some Teams-specific enrichment may be reduced or skipped.
- Team and channel inventory remain part of the standard app-based path, but member-count and guest-count enrichment require the broader Teams member read scope set.

## Default Output Location

If you do not provide `-ExportPath`, the launcher writes outputs under:

```text
%LOCALAPPDATA%\Arraya\M365TenantAssessment\Outputs
```

On a typical Windows workstation this resolves to a path similar to:

```text
C:\Users\<your-user>\AppData\Local\Arraya\M365TenantAssessment\Outputs
```

## Run The Launcher

The main user entrypoint is:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Running the script with no parameters opens a menu for the supported workflows:

- `M365`
- `M365Preflight`
- `M365Collect`
- `M365Export`
- `AD`
- `Improve`
- `Compare`

The current launcher menu order is:

1. `M365`
2. `M365Preflight` for Microsoft 365 connection and permission preflight only
3. `M365Collect`
4. `M365Export`
5. `AD`
6. `Improve`
7. `Compare`

For a workflow-by-workflow breakdown, including how the actions relate to each other and where maintainers may want to reduce overlap, see [docs/architecture/workflow-overview.md](docs/architecture/workflow-overview.md).

## Phase Pipeline Entry Points

The redesign branch also includes explicit phase-runner entry points for checkpoint-backed debugging and resume:

- `src/scripts/assessments/tenant-wide/Invoke-M365TenantPipeline.ps1`
- `src/scripts/assessments/tenant-wide/Resume-M365TenantPipeline.ps1`

Use these when you want to:

- run only through a named phase
- rerun one phase against a saved checkpoint chain
- resume a stopped assessment from `Support\Pipeline`

The redesign architecture and branch strategy are documented in [docs/architecture/assessment-pipeline-redesign.md](docs/architecture/assessment-pipeline-redesign.md).

For operator guidance on reading `Improve` findings and understanding rule IDs, see [docs/runbooks/improvement-plan-rule-taxonomy.md](docs/runbooks/improvement-plan-rule-taxonomy.md).

## Common Commands

Start the menu-driven launcher:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Run a full Microsoft 365 assessment with the default interactive sign-in flow:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365
```

Run only the Microsoft 365 connection and permission preflight, then exit without collection or export:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Preflight `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

Run a full assessment and write outputs to a specific folder:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -ExportPath 'C:\Assessment-Outputs'
```

Run a full assessment with certificate-based app authentication:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

Run a full assessment with certificate-based auth and skip the permission preflight gate:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -SkipPermissionPreflight
```

Run a full assessment with client-secret authentication:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode ClientSecret `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -ClientSecret '<client-secret>'
```

Collect tenant data only and save a JSON snapshot for later export:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Collect `
  -OutputProfile SolutionsEngineer
```

Collect data only and let the run continue even if some startup permission checks would normally stop it:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Collect `
  -OutputProfile SolutionsEngineer `
  -SkipPermissionPreflight
```

The following actions prompt you for additional inputs at runtime:

- `M365Export`: asks for the saved JSON snapshot path.
- `Improve`: asks for the tenant assessment JSON path or manifest path and optional output folder.
- `Compare`: asks for baseline and current JSON snapshot paths.

`M365Preflight` uses the same auth inputs as the main M365 workflow, including `-SkipAuth`, `-SkipPermissionPreflight`, `-AuthMode`, `-TenantId`, `-ClientId`, `-CertificateThumbprint`, and `-ClientSecret`.

Generate artifacts from an existing JSON snapshot:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365Export
```

Compare two saved tenant snapshots:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Compare
```

Build an improvement plan from an existing tenant JSON file:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Improve
```

Run a full assessment and automatically chain `Improve` from the artifacts that same run produced:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365
```

`M365` now includes the `Improve` step by default and uses a consolidated remediation-first output model. The workflow preserves the JSON snapshot needed for replay and post-processing, so `-SkipJsonReport` is ignored unless you explicitly opt back into the legacy assessment artifact family with `-IncludeLegacyAssessmentArtifacts`.

If you need the older assessment-only behavior, use:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -SkipImprove
```

End-to-end example to collect a snapshot and then build an improvement plan from it:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Collect `
  -OutputProfile SolutionsEngineer `
  -ExportPath 'C:\Assessment-Outputs'

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Improve `
  -LiveRefresh
```

`Improve` is still available as a separate workflow for post-processing an existing snapshot, but the standard `M365` run now includes it automatically. `M365Collect` remains the advanced snapshot-only workflow, and `-RunImprove` is still available there when you want to chain post-processing from a collection run.

For a fully non-interactive improvement-plan run, call the reporting wrapper directly:

```powershell
.\src\scripts\reporting\Invoke-M365TenantImprovementPlan.ps1 `
  -AssessmentJsonPath 'C:\Assessment-Outputs\<tenant-snapshot>.json' `
  -OutputFolder 'C:\Assessment-Outputs\Improve' `
  -LiveRefresh
```

Use `-LiveRefresh` when you want a "snapshot plus live refresh" run. It is a friendlier alias for the existing `-UseGraphFallback` switch and tells `Improve` to use the saved snapshot first, then fill supported gaps from Microsoft Graph when needed.

Add `-IncludeLegacyArtifacts` only if you still want the older `*-ImprovementPlan.csv` and `*-ImprovementPlan.md` planning artifacts in addition to the default set.
Add `-IncludeLegacyAssessmentArtifacts` when you also want the older assessment HTML / best-practices HTML / questionnaire / PDF family.

The `Improve` workflow now produces this simplified default output set:

- top level: `*-CustomerAssessmentReport.docx`, `*-EngineerActionPack.md`
- support folder: `*-AssessmentSnapshot.json`, `*-ImprovementPlan.json`, `*.manifest.json`, `*-RemediationSnippets.ps1`

If you still need the older technical CSV and Markdown artifacts, generate them explicitly with `-IncludeLegacyArtifacts` when calling the reporting wrapper directly.

Use the taxonomy guide to understand whether a finding came from the derived assessment layer or from a built-in remediation rule:

- [docs/runbooks/improvement-plan-rule-taxonomy.md](docs/runbooks/improvement-plan-rule-taxonomy.md)

## Output Profiles

Choose the output profile based on the audience and collection/reporting depth you need:

- `Presales`: `Minimum` scope for fast presales posture reviews.
- `ExecutiveLevel`: `Minimum` scope for leadership-ready summaries.
- `SolutionsEngineer`: `Operator` scope for the balanced remediation-first assessment run.
- `Machine`: `Automation` scope for JSON-first collection, replay, and downstream processing.
- `Geek`: `Geek` scope for deep engineer troubleshooting.
- `TenantToTenantMigration`: `All` scope for the deepest migration-oriented collection.

`Operator` is the standard balanced assessment depth and does not run the full combined user/mailbox projection.
`All` is the migration-oriented deep mode and includes the combined user/mailbox projection.

You can pass more than one profile in a comma-separated list:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -OutputProfile SolutionsEngineer,ExecutiveLevel
```

## What Gets Generated

Default `M365` now creates these top-level operator deliverables:

- `*-CustomerAssessmentReport.docx`
- `*-EngineerActionPack.md`
- `*.xlsx` for profiles whose policy enables workbook output, including `SolutionsEngineer` and `TenantToTenantMigration`

It also creates these support artifacts under `Support\...`:

- `Support\*-ImprovementPlan.json`
- `Support\*.manifest.json`
- `Support\*-RemediationSnippets.ps1`

The run still preserves the reusable assessment snapshot JSON internally because `Improve`, `M365Export`, and comparison/replay depend on it.

If you opt into `-IncludeLegacyAssessmentArtifacts`, the assessment can also create:

- `*-TenantSnapshot.html`
- `*-BestPracticesSnapshot.html`
- `*-TenantToTenantQuestionnaire.md`
- `*.pdf`

Operational logs and exported error details are written to a `Debugging` subfolder alongside the main outputs.

## How To Read The Results

Start with the artifact that best matches your audience:

- `CustomerAssessmentReport.docx`: primary customer-facing deliverable.
- `EngineerActionPack.md`: primary operator-facing deliverable.
- `Support\ImprovementPlan.json`: best for export reuse, filtering, comparisons, and automation.
- `Support\RemediationSnippets.ps1`: support helper commands.

Recommended review flow:

1. Open `CustomerAssessmentReport.docx` first.
2. Review `EngineerActionPack.md` for implementation planning.
3. Keep `Support\ImprovementPlan.json` as your baseline for later comparison or downstream processing.
4. Use legacy workbook / assessment HTML outputs only when you explicitly generated them for a deeper technical review.

When reviewing `Improve` outputs:

1. Start with `*-CustomerAssessmentReport.docx` for stakeholder-facing messaging.
2. Use `*-EngineerActionPack.md` for operator execution planning.
3. Use `Support\ImprovementPlan.json` for filtering, automation, or downstream transformations.
4. Use `RelatedWorksheet` and `Source` to trace each finding back to its evidence and rule origin.

## Running On Another Machine

If you want to test the workflow on another workstation or jump box, use this checklist:

1. Install PowerShell 7 or later.
2. Clone or copy this repo locally.
3. Install required modules:

```powershell
.\tools\install-microsoft-modules.ps1
```

4. If you want to run tests, install/update Pester in your user scope:

```powershell
Install-Module Pester -Scope CurrentUser -Force
```

5. If you use certificate auth, import the matching `.pfx` on that machine and verify the cert has a private key.
6. Start with a low-risk validation:
   - first run `Improve` against an existing known-good snapshot JSON
   - then run `M365Collect`
   - then run `Improve` against that fresh snapshot

Recommended order for another-machine validation:

```powershell
.\src\scripts\reporting\Invoke-M365TenantImprovementPlan.ps1 `
  -AssessmentJsonPath 'C:\KnownGood\assessment.json' `
  -OutputFolder 'C:\KnownGood\Improve'
```

Then:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Collect `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -OutputProfile Machine `
  -ExportPath 'C:\Assessment-Outputs'
```

Then:

```powershell
.\src\scripts\reporting\Invoke-M365TenantImprovementPlan.ps1 `
  -AssessmentJsonPath 'C:\Assessment-Outputs\<tenant-snapshot>.json' `
  -OutputFolder 'C:\Assessment-Outputs\Improve'
```

## Troubleshooting

If a required module is missing, rerun:

```powershell
.\tools\install-microsoft-modules.ps1
```

If you are prompted for authentication unexpectedly, check whether you intended to use the default interactive mode or whether `-SkipAuth` should be used to reuse an existing session.
If the run stops before collection because of the startup access gate and you want best-effort behavior instead, rerun with `-SkipPermissionPreflight`.

If PDF output is missing, the assessment can still complete successfully. PDF generation depends on a locally installed Chromium-based browser such as Google Chrome or Microsoft Edge.

If app-only authentication fails, review the setup runbooks:

- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)

## Additional Guidance

- [Tenant Assessment Quick Start](docs/runbooks/tenant-assessment-quick-start.md)
- [Customer Execution Checklist](docs/runbooks/customer-execution-checklist.md)
- [Onboarding Engineer](docs/runbooks/onboarding-engineer.md)
