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
- `Certificate`: supports Microsoft Graph, Exchange Online app auth, and SharePoint Online certificate auth.
- `ClientSecret`: Microsoft Graph supports app auth, but Exchange falls back to delegated auth, SharePoint Online admin cmdlets are not used with client secret auth, and Teams PowerShell is skipped.

If you are using app-based authentication, complete the setup guidance first:

- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)

If you already connected to Microsoft Graph and Exchange Online in the same PowerShell session, you can reuse those sessions with `-SkipAuth`.

## Required Access And API Permissions

This repo does not yet maintain a separate, minimal permission manifest. The guidance below is derived from the current launcher and connector code paths used by the assessment.

### Recommended Operator Access

For a reliable full `M365` run in delegated interactive mode, the operator account should be able to connect to:

- Microsoft Graph
- Exchange Online PowerShell
- SharePoint Online Admin PowerShell
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

When you run the main assessment in `Interactive` mode, the connector requests these Microsoft Graph delegated scopes:

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
- `Files.Read.All`
- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`

These scopes support the repo's current Graph-based collection for tenant, identity, reporting, security, collaboration, and SharePoint data.

### App-Based Permission Expectations

For `Certificate` or `ClientSecret` mode, the app registration should be granted the Microsoft Graph application permissions needed to cover the same read-focused collection areas where application permissions exist, especially:

- `Organization.Read.All`
- `User.Read.All`
- `AuditLog.Read.All`
- `Group.Read.All`
- `GroupMember.Read.All`
- `RoleManagement.Read.Directory`
- `Domain.Read.All`
- `Device.Read.All`
- `Reports.Read.All`
- `Policy.Read.All`
- `SecurityEvents.Read.All`
- `Application.Read.All`
- `Sites.Read.All`
- `Files.Read.All`
- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`

For app-based operation beyond Graph:

- Exchange Online requires Exchange app-only access for the app registration when using `Certificate` mode.
- SharePoint Online certificate mode requires certificate-based app access for `Connect-SPOService`.
- Teams PowerShell in this workflow is still delegated-only, so app-based runs may have reduced Teams detail.

### Workload Notes

- Exchange Online connectivity is required for the full `M365` assessment path. If Exchange auth fails, the run stops.
- SharePoint Online admin connectivity is helpful for fuller coverage, but some SharePoint collection can still proceed through Microsoft Graph.
- Teams connectivity is non-blocking in app-based auth modes, but some Teams-specific enrichment may be reduced or skipped.

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
- `M365Collect`
- `M365Export`
- `AD`
- `Improve`
- `Compare`

For a workflow-by-workflow breakdown, including how the actions relate to each other and where maintainers may want to reduce overlap, see [docs/architecture/workflow-overview.md](docs/architecture/workflow-overview.md).

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

The following actions prompt you for additional inputs at runtime:

- `M365Export`: asks for the saved JSON snapshot path.
- `Improve`: asks for the tenant assessment JSON path or manifest path and optional output folder.
- `Compare`: asks for baseline and current JSON snapshot paths.

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
  -Action M365 `
  -RunImprove
```

If the selected output profile does not normally emit JSON, `-RunImprove` automatically adds the `Machine` profile for that run so a reusable snapshot is preserved. If you pass `-SkipJsonReport`, the launcher ignores it when `-RunImprove` is present.

End-to-end example to collect a snapshot and then build an improvement plan from it:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Collect `
  -OutputProfile SolutionsEngineer `
  -ExportPath 'C:\Assessment-Outputs'

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Improve `
  -UseGraphFallback
```

`Improve` is a separate workflow. A standard `M365` run does not automatically create an improvement plan, and the default `SolutionsEngineer` profile does not emit JSON unless you use `M365Collect` or choose a JSON-enabled profile such as `Geek` or `Machine`.

As a temporary operator shortcut, `-RunImprove` removes most of that manual coordination by preserving JSON automatically and invoking `Improve` against the latest manifest from the just-finished run.

For a fully non-interactive improvement-plan run, call the reporting wrapper directly:

```powershell
.\src\scripts\reporting\Invoke-M365TenantImprovementPlan.ps1 `
  -AssessmentJsonPath 'C:\Assessment-Outputs\<tenant-snapshot>.json' `
  -OutputFolder 'C:\Assessment-Outputs\Improve' `
  -UseGraphFallback
```

The `Improve` workflow now produces both operator and customer-oriented outputs by default:

- `*-ImprovementPlan.json`
- `*-ImprovementPlan.csv`
- `*-ImprovementPlan.md`
- `*-CustomerRemediationReport.md`
- `*-EngineerActionPack.md`
- `*-RemediationSnippets.ps1`

Use the taxonomy guide to understand whether a finding came from the derived assessment layer or from a built-in remediation rule:

- [docs/runbooks/improvement-plan-rule-taxonomy.md](docs/runbooks/improvement-plan-rule-taxonomy.md)

## Output Profiles

Choose the output profile based on the audience and artifact set you need:

- `Presales`: minimum-scope collection with technical HTML and questionnaire output.
- `SolutionsEngineer`: default profile with workbook and technical HTML.
- `ExecutiveLevel`: minimum-scope collection with best-practices analysis HTML.
- `TenantToTenantMigration`: workbook, technical HTML, and questionnaire.
- `Geek`: expanded output set including workbook, HTML, questionnaire, and JSON.
- `Machine`: JSON-focused output for downstream processing.

You can pass more than one profile in a comma-separated list:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -OutputProfile SolutionsEngineer,ExecutiveLevel
```

## What Gets Generated

Depending on the chosen action and output profile, the assessment can create:

- `*.xlsx`: the main assessment workbook.
- `*.html`: the full technical tenant report.
- `*-BestPracticesAnalysis.html`: the assessment-focused summary HTML.
- `*-TenantToTenantQuestionnaire.md`: the migration questionnaire populated from tenant data.
- `*.json`: the reusable assessment snapshot.
- `*.pdf`: a PDF rendered from the full HTML report when a supported browser is available.
- `*.manifest.json`: an index of generated artifacts.

Operational logs and exported error details are written to a `Debugging` subfolder alongside the main outputs.

## How To Read The Results

Start with the artifact that best matches your audience:

- `Workbook (.xlsx)`: primary technical deliverable. Begin with `BestPractices`, `BestPracticeFindings`, `MigrationReadiness`, and `SecureScoreActions`.
- `BestPracticesAnalysis.html`: best for leadership or quick review of posture and recommended actions.
- `Full HTML report`: best for browser-based technical walkthroughs.
- `TenantToTenantQuestionnaire.md`: best for migration discovery follow-up and stakeholder interviews.
- `JSON snapshot`: best for export reuse, comparisons, and automation.

Recommended review flow:

1. Open the workbook or best-practices HTML first.
2. Identify the highest-impact findings and blockers.
3. Use the questionnaire markdown to fill any discovery gaps with the customer.
4. Keep the JSON snapshot as your baseline for later comparison runs.

When reviewing `Improve` outputs:

1. Start with `*-CustomerRemediationReport.md` for stakeholder-facing messaging.
2. Use `*-EngineerActionPack.md` for operator execution planning.
3. Use `ImprovementPlan.json` or `ImprovementPlan.csv` for filtering, backlog import, or automation.
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

If PDF output is missing, the assessment can still complete successfully. PDF generation depends on a locally installed Chromium-based browser such as Google Chrome or Microsoft Edge.

If app-only authentication fails, review the setup runbooks:

- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)

## Additional Guidance

- [Tenant Assessment Quick Start](docs/runbooks/tenant-assessment-quick-start.md)
- [Customer Execution Checklist](docs/runbooks/customer-execution-checklist.md)
- [Onboarding Engineer](docs/runbooks/onboarding-engineer.md)
