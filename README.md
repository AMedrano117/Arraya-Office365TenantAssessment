# Modern Workplace Tenant Assessment

PowerShell automation for Microsoft 365 tenant assessments, reporting, and improvement planning.

For operator-focused setup and execution guidance, use [RUN.md](RUN.md).

## User Quick Start
Run the launcher and pick an action from the menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Users do not need to manually import modules. The launcher imports `Arraya.M365.AssessmentRunner`, which then loads `Office365Custom` only when needed and only if it is not already imported.

## Auth Modes
- `Delegated`: interactive sign-in for full module compatibility.
- `Certificate`: noninteractive app auth for Graph and Exchange, with Graph-based SharePoint fallback in PowerShell 7.
- `Client secret`: noninteractive app auth for Graph and Exchange app-only; Teams PowerShell remains limited in app-secret mode.

If you do not pass `-AuthMode`, the assessment defaults to delegated interactive sign-in. For backward compatibility, supplying `-CertificateThumbprint` still switches the run to certificate auth, and supplying `-ClientSecret` still switches the run to client-secret auth.
If you already connected to Microsoft Graph and Exchange Online in the current session, you can run with `-SkipAuth` to reuse those sessions and bypass the repo's authentication bootstrap.

Examples:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode Interactive

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -SkipAuth

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode ClientSecret `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -ClientSecret '<client-secret>'

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -ExportPath 'C:\Assessment-Outputs' `
  -OutputProfile SolutionsEngineer,ExecutiveLevel
```

## Collection Vs Export
The M365 workflow now supports explicit separation between data collection and artifact export:

- `M365Collect`: runs discovery/collection only and writes a JSON snapshot.
- `M365Export`: loads a previously collected JSON snapshot and generates artifacts without re-collecting tenant data.
- `M365`: legacy combined path (collect + export in one run) remains available.
- each run also writes a `*.manifest.json` artifact index next to the primary export filename.
- JSON snapshots now use a versioned V2 contract with explicit sections: `Metadata`, `CollectionPlan`, `Data`, `Derived`, and `Diagnostics`.
- Import is backward compatible with older V1 snapshots through an in-memory adapter.

Examples:

```powershell
# Collect data only (JSON snapshot)
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Collect `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -OutputProfile SolutionsEngineer

# Export artifacts from an existing JSON snapshot
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Export `
  -OutputProfile ExecutiveLevel
```

## Assessment Outputs
For a Microsoft 365 tenant assessment run, the primary deliverables are:

- `*.xlsx`: the main assessment workbook. This is the primary human-readable assessment artifact.
- `*-BestPracticesAnalysis.html`: the assessment-only HTML focused on `BestPractices`, `BestPracticeFindings`, `MigrationReadiness`, and top `SecureScoreActions`.
- `*.html`: the full styled tenant report for browser-based review.
- `*-TenantToTenantQuestionnaire.md`: the migration questionnaire filled from discovered tenant data.
- `*.json`: the machine-readable snapshot used for reuse, comparison, and downstream reporting when enabled.
- `*.pdf`: the fixed-layout PDF generated from the full HTML report through headless Chrome/Edge when enabled and available.

Operational logs and exported error/debug bundles are written into a `Debugging` subfolder inside each assessment output folder so primary deliverables stay easier to scan.

The workbook is the file that contains the assessment summary and findings. The most assessment-oriented worksheets are:

- `BestPractices`
- `BestPracticeFindings`
- `MigrationReadiness`
- `SecureScoreActions`

The questionnaire Markdown is intended as a migration intake companion, not a replacement for the workbook.

## Output Profiles
The assessment run supports six need-based output profiles:

- `Presales`: scope `Minimum`; outputs technical HTML + questionnaire
- `SolutionsEngineer`: scope `Combined`; outputs workbook + technical HTML
- `ExecutiveLevel`: scope `Minimum`; outputs best practices analysis HTML
- `TenantToTenantMigration`: scope `Combined`; outputs workbook + technical HTML + questionnaire
- `Geek`: scope `Geek`; outputs workbook + technical HTML + best practices HTML + questionnaire + JSON
- `Machine`: scope `Geek`; outputs JSON only

`SolutionsEngineer` is the default profile.
You can run multiple profiles in one command by passing a comma-separated list (for example `SolutionsEngineer,ExecutiveLevel`).
When multiple profiles are supplied, the assessment runs once using the highest required reporting scope and produces the union of requested artifacts.

Reporting scope is not prompted interactively. Scope is automatically derived from the selected output profile.

## Runtime Notes
- In `Minimum` scope profiles, the collector depth policy trims high-cardinality enrichment to reduce runtime and memory pressure.
- Examples: Entra group deep membership/license expansion and detailed SSO app inventory are reduced in `Minimum`.
- The output contract is preserved: workbook tabs and report artifacts still generate with compatible values.
- Run logs now include collector duration and memory summaries (`[CollectorMetrics]`) plus inventory row counts (`[CollectorInventory]`) for hotspot review.

## HTML And PDF
- The best practices analysis HTML is generated independently of the full HTML report.
- The full HTML report is generated only when the selected output profile enables technical HTML output.
- PDF is disabled by default for all profiles in this model and remains skippable with `-SkipPdfReport`.
- PDF rendering uses a locally installed Chromium-based browser, preferring Google Chrome and falling back to Microsoft Edge.

If a supported browser is unavailable, the workbook, questionnaire, and HTML outputs still complete.

## Primary Entry Scripts
- `src/scripts/operations/Start-M365TenantAssessment.ps1`
- `src/scripts/assessments/tenant-wide/Invoke-M365FullTenantAssessment.ps1`
- `src/scripts/assessments/tenant-wide/Invoke-M365TenantDataCollection.ps1`
- `src/scripts/assessments/identity/Invoke-ActiveDirectoryTenantAssessment.ps1`
- `src/scripts/assessments/tenant-wide/Invoke-M365GraphActivityReport.ps1`
- `src/scripts/reporting/Invoke-M365TenantAssessmentExport.ps1`
- `src/scripts/reporting/Invoke-M365TenantImprovementPlan.ps1`
- `src/scripts/reporting/Invoke-M365TenantAssessmentComparison.ps1`

## Repo Layout
- `docs`: user-facing templates and questionnaires. The tenant-to-tenant questionnaire template lives under `docs/templates`.
- `src/modules/Arraya.M365.Common`: shared helpers and Office365Custom local import function.
- `src/modules/Arraya.M365.AssessmentRunner`: user-facing commands that execute assessment/report scripts.
- `src/scripts/migrated/legacy`: migrated legacy scripts kept for compatibility, including the core M365 assessment engine, HTML/PDF helper, and questionnaire exporter.
- `src/vendor/Office365Custom/1.2.0`: vendored module used as shared function source.

## Development
1. Install PowerShell 7+.
2. Run `tools/bootstrap-dev.ps1`.
3. Validate with `tools/invoke-scriptanalyzer.ps1` and `tools/run-pester.ps1`.
