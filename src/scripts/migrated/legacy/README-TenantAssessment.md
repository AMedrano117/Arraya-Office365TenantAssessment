# Tenant Assessment Quick Start

## Easiest way for users
Run one script and choose from the menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Users do not need to import any module manually. The launcher imports the local runner module automatically.
The runner module then loads the local `Office365Custom` module (latest version folder) only if it is not already imported.

## Auth Modes
- `Delegated`: interactive sign-in for full module compatibility.
- `Certificate`: noninteractive app auth for Graph and Exchange, with Graph-based SharePoint fallback in PowerShell 7.
- `Client secret`: noninteractive app auth for Graph and Exchange app-only; Teams PowerShell remains limited in app-secret mode.

Examples:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -ClientSecret '<client-secret>'
```

## Direct actions (optional)
```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action AD
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Graph
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Improve
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Compare
```

## Current Output Set
A Microsoft 365 assessment run now produces:

- `*.xlsx`: the main assessment workbook
- `*-BestPracticesAnalysis.html`: the assessment-only HTML focused on best practices, detailed findings, migration readiness, and Secure Score actions
- `*.html`: the browser-friendly full tenant assessment report
- `*-TenantToTenantQuestionnaire.md`: the tenant-to-tenant migration questionnaire populated from discovered data
- `*.json`: the reusable assessment snapshot when enabled
- `*.pdf`: the fixed-layout PDF rendered from the full HTML report when enabled and Chrome or Edge is available

The workbook is the primary assessment deliverable. The most important worksheets for assessment review are:

- `BestPractices`
- `BestPracticeFindings`
- `MigrationReadiness`
- `SecureScoreActions`

## Current Script Locations
Runner commands are in:

`src\modules\Arraya.M365.AssessmentRunner\`

Shared helper functions (such as `Write-ProgressHelper`) are sourced from:

`src\vendor\Office365Custom\<version>\`

The core M365 assessment engine is:

`src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1`

The tenant-to-tenant questionnaire exporter is:

`src\scripts\migrated\legacy\Export-TenantToTenantQuestionnaireMarkdown.ps1`

The HTML and PDF helper is:

`src\scripts\migrated\legacy\New-TenantHtmlReport.ps1`

The questionnaire template is:

`docs\templates\Microsoft 365 Tenant to Tenant Questionnaire.md`

## Output Profiles
The assessment script supports six need-based output profiles:

- `Presales`: scope `Minimum`; outputs technical HTML + questionnaire
- `SolutionsEngineer`: scope `Combined`; outputs workbook + technical HTML
- `ExecutiveLevel`: scope `Minimum`; outputs best practices analysis HTML
- `TenantToTenantMigration`: scope `Combined`; outputs workbook + technical HTML + questionnaire
- `Geek`: scope `Geek`; outputs workbook + technical HTML + best practices HTML + questionnaire + JSON
- `Machine`: scope `Geek`; outputs JSON only

`SolutionsEngineer` is the default profile.

## HTML And PDF
The assessment script now attempts to generate:

- `*-BestPracticesAnalysis.html` independently of the full HTML output
- `*.html` when the selected profile enables technical HTML output and `-SkipHtmlReport` is not provided
- `*.pdf` from the full HTML only when explicitly enabled in profile policy (disabled by default in current model)

PDF rendering prefers Google Chrome and falls back to Microsoft Edge when available.

## Local Output Location
Validation and test outputs should not be written into this repo.

- Preferred local output root: `%LOCALAPPDATA%\Arraya\M365TenantAssessment\Outputs`
- Example on this machine: `C:\Users\amedrano\AppData\Local\Arraya\M365TenantAssessment\Outputs`
- Future ad hoc test runs should target that local folder instead of `artifacts/`

Removing tracked artifacts from the repo prevents future growth, but it does not shrink existing git history by itself. If you want the repository size reduced retroactively, the next step is a history rewrite with `git filter-repo` or BFG.

Phase 3 alignment status:

- Local Office365Custom import bootstrap is centralized in:
  - `.\Import-Office365CustomLocal.ps1`
- The following entry points now dot-source the shared loader and call `Import-Office365CustomLocal`:
  - `Get-FullTenantReportDetails.ps1`
  - `Get-ActiveDirectoryReport.ps1`
  - `Get-GraphAPIActivityReport.ps1`
  - `Arraya.TenantAssessment\Arraya.TenantAssessment.psm1`
- Import behavior checks if `Office365Custom` is already loaded and does not re-import unnecessarily.

Phase 2 alignment status:

- `Get-FullTenantReportDetails.ps1`
  - `Capture-ErrorHelper` and `Write-Log` now use `Office365Custom` implementations with compatibility shims for legacy no-ErrorRecord error logging.
  - `Get-ExportPath` now uses `Office365Custom\Get-ExportPath` for interactive use and keeps a non-interactive shim for automation (`UserInputPath`).
- `Get-ActiveDirectoryReport.ps1`
  - `Capture-ErrorHelper` now uses `Office365Custom\Capture-ErrorHelper` directly.
  - `Write-Log` now uses `Office365Custom\Write-Log` for non-error log levels and keeps an AD-specific compatibility fallback for `ERROR` entries (legacy path naming and no ErrorRecord flows).
  - `Get-ExportPath` now uses `Office365Custom\Get-ExportPath` and preserves the script's two-value return contract.
