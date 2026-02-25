# Tenant Assessment Quick Start

## Easiest way for users
Run one script and choose from the menu:

```powershell
.\Start-TenantAssessment.ps1
```

Users do not need to import any module manually. The launcher imports the local runner module automatically.
The runner module then loads the local `Office365Custom` module (latest version folder) only if it is not already imported.

## Direct actions (optional)
```powershell
.\Start-TenantAssessment.ps1 -Action M365
.\Start-TenantAssessment.ps1 -Action AD
.\Start-TenantAssessment.ps1 -Action Graph
.\Start-TenantAssessment.ps1 -Action Improve
.\Start-TenantAssessment.ps1 -Action Compare
```

## Module location
Runner commands are in:

`.\Arraya.TenantAssessment\Arraya.TenantAssessment.psm1`

Shared helper functions (such as `Write-ProgressHelper`) are sourced from:

`.\Office365Custom\<version>\`

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
