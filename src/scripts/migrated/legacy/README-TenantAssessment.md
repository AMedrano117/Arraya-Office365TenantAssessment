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
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365 -SkipImprove
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action AD
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Improve
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Compare
```

`M365` now includes the `Improve` step by default. Use `-SkipImprove` only when you explicitly want the older assessment-only behavior.
If you run `Improve` separately, use `-LiveRefresh` when you want snapshot-plus-live-refresh behavior; it is a friendlier alias for `-UseGraphFallback`.

## Current Output Set
A default Microsoft 365 assessment run now produces:

- `*-CustomerRemediationReport.html`
- `*-EngineerActionPack.md`
- `*-ImprovementPlan.json`
- `*-RemediationSnippets.ps1`
- `*.manifest.json`
- `Debugging\...`

The reusable assessment snapshot JSON is still preserved because `Improve`, `M365Export`, and replay/comparison workflows depend on it.

If you explicitly pass `-IncludeLegacyAssessmentArtifacts`, the older assessment artifact family is also generated:

- `*.xlsx`
- `*-BestPracticesAnalysis.html`
- `*.html`
- `*-TenantToTenantQuestionnaire.md`
- `*.pdf`

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

- `Presales`: scope `Minimum`
- `SolutionsEngineer`: scope `Combined`
- `ExecutiveLevel`: scope `Minimum`
- `TenantToTenantMigration`: scope `Combined`
- `Geek`: scope `Geek`
- `Machine`: scope `Geek`; JSON-focused automation profile

`SolutionsEngineer` is the default profile.

## HTML And PDF
The remediation workflow now generates `*-CustomerRemediationReport.html` by default.

The legacy assessment HTML/PDF family is generated only when `-IncludeLegacyAssessmentArtifacts` is used:

- `*-BestPracticesAnalysis.html`
- `*.html`
- `*.pdf`

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
