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
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Improve
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Compare
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365 -RunImprove
```

## Script entry points
- `src/scripts/assessments/tenant-wide/Invoke-M365FullTenantAssessment.ps1`
- `src/scripts/assessments/identity/Invoke-ActiveDirectoryTenantAssessment.ps1`
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

## Operator references
- [../../RUN.md](../../RUN.md)
- [improvement-plan-rule-taxonomy.md](improvement-plan-rule-taxonomy.md)

## Improve workflow notes
- `Improve` is a separate post-processing workflow that consumes an assessment JSON snapshot.
- The launcher now supports `-RunImprove` on `M365` and `M365Collect` as a temporary operator shortcut.
- When `-RunImprove` is used and the chosen profile would not normally emit JSON, the launcher appends `Machine` for that run so the snapshot is preserved.
- `Improve` can now take either the snapshot JSON path or the `*.manifest.json` path from the same run.
- The improvement plan combines derived assessment findings with built-in remediation heuristics.
- Rule IDs such as `ID-007`, `CA-012`, `DEV-006`, and `EX-007` are internal repo rule identifiers, not Microsoft-native control IDs.
- Use the taxonomy guide when you need to explain where a finding came from or how to trace it back to the supporting worksheet.

## Another machine validation
When validating on another machine, start with the safest path first:

1. Run `Improve` against an existing known-good snapshot.
2. Run `M365Collect`.
3. Run `Improve` against the newly collected snapshot.

This separates environment/setup issues from tenant-authentication and collector issues.
