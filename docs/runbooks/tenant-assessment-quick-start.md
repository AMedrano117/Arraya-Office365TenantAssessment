# Tenant Assessment Quick Start

## Menu launcher (recommended)
Run one script and choose from the menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

## Direct launch options
```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action AD
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Graph
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Improve
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Compare
```

## Script entry points
- `src/scripts/assessments/tenant-wide/Invoke-M365FullTenantAssessment.ps1`
- `src/scripts/assessments/identity/Invoke-ActiveDirectoryTenantAssessment.ps1`
- `src/scripts/assessments/tenant-wide/Invoke-M365GraphActivityReport.ps1`
- `src/scripts/reporting/Invoke-M365TenantImprovementPlan.ps1`
- `src/scripts/reporting/Invoke-M365TenantAssessmentComparison.ps1`

## Module import behavior
- `src/modules/Arraya.M365.AssessmentRunner` is imported by launcher/wrapper scripts.
- `src/modules/Arraya.M365.Common/Public/Import-ArrayaOffice365CustomLocal.ps1` handles local `Office365Custom` import.
- `Office365Custom` is imported only if not already loaded.
- Required commands are validated after import.

## Legacy compatibility
Migrated legacy scripts are stored in `src/scripts/migrated/legacy`.
The legacy `Import-Office365CustomLocal.ps1` now delegates to the shared Common module loader.
