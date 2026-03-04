# Modern Workplace Tenant Assessment

PowerShell automation for Microsoft 365 tenant assessments, reporting, and improvement planning.

## User Quick Start
Run the launcher and pick an action from the menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Users do not need to manually import modules. The launcher imports `Arraya.M365.AssessmentRunner`, which then loads `Office365Custom` only when needed and only if it is not already imported.

## Assessment Outputs
For a Microsoft 365 tenant assessment run, the primary deliverables are:

- `*.xlsx`: the main assessment workbook. This is the primary human-readable assessment artifact.
- `*-Assessment.html`: the assessment-only HTML focused on `BestPractices`, `BestPracticeFindings`, `MigrationReadiness`, and top `SecureScoreActions`.
- `*.html`: the full styled tenant report for browser-based review.
- `*-TenantToTenantQuestionnaire.md`: the migration questionnaire filled from discovered tenant data.
- `*.json`: the machine-readable snapshot used for reuse, comparison, and downstream reporting when enabled.
- `*.pdf`: the fixed-layout PDF generated from the full HTML report through headless Chrome/Edge when enabled and available.

The workbook is the file that contains the assessment summary and findings. The most assessment-oriented worksheets are:

- `BestPractices`
- `BestPracticeFindings`
- `MigrationReadiness`
- `SecureScoreActions`

The questionnaire Markdown is intended as a migration intake companion, not a replacement for the workbook.

## Output Profiles
The assessment run supports three output profiles:

- `Lean`: workbook, assessment HTML, and questionnaire only
- `Standard`: workbook, assessment HTML, full HTML, and questionnaire
- `Full`: workbook, assessment HTML, full HTML, questionnaire, JSON, and PDF

`Standard` is the default. Explicit skip switches still override the profile.

## HTML And PDF
- The assessment-only HTML is generated independently of the full HTML report.
- The full HTML report is skipped when `-SkipHtmlReport` is used or the `Lean` profile is selected.
- PDF is generated from the full HTML report unless `-SkipPdfReport` is used or the selected profile disables it.
- PDF rendering uses a locally installed Chromium-based browser, preferring Google Chrome and falling back to Microsoft Edge.

If a supported browser is unavailable, the workbook, questionnaire, and HTML outputs still complete.

## Primary Entry Scripts
- `src/scripts/operations/Start-M365TenantAssessment.ps1`
- `src/scripts/assessments/tenant-wide/Invoke-M365FullTenantAssessment.ps1`
- `src/scripts/assessments/identity/Invoke-ActiveDirectoryTenantAssessment.ps1`
- `src/scripts/assessments/tenant-wide/Invoke-M365GraphActivityReport.ps1`
- `src/scripts/reporting/Invoke-M365TenantImprovementPlan.ps1`
- `src/scripts/reporting/Invoke-M365TenantAssessmentComparison.ps1`

## Repo Layout
- `docs`: user-facing templates and questionnaires. The tenant-to-tenant questionnaire template lives under `docs/templates`.
- `artifacts`: generated assessment outputs from validation and test runs.
- `src/modules/Arraya.M365.Common`: shared helpers and Office365Custom local import function.
- `src/modules/Arraya.M365.AssessmentRunner`: user-facing commands that execute assessment/report scripts.
- `src/scripts/migrated/legacy`: migrated legacy scripts kept for compatibility, including the core M365 assessment engine, HTML/PDF helper, and questionnaire exporter.
- `src/vendor/Office365Custom/1.2.0`: vendored module used as shared function source.

## Development
1. Install PowerShell 7+.
2. Run `tools/bootstrap-dev.ps1`.
3. Validate with `tools/invoke-scriptanalyzer.ps1` and `tools/run-pester.ps1`.
