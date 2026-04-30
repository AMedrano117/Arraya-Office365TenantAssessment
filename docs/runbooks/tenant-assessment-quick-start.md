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
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365 -SkipImprove
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
The standalone `Invoke-EntraAppReport.ps1` file in that folder is kept only as an unsupported reference script and now requires an explicit override switch to run.

## Operator references
- [../../RUN.md](../../RUN.md)
- [improvement-plan-rule-taxonomy.md](improvement-plan-rule-taxonomy.md)

## Improve workflow notes
- `M365` is now the standard full assessment path and includes `Improve` by default.
- Use `-SkipImprove` only when you explicitly want the older assessment-only behavior.
- `Improve` is still available as a separate post-processing workflow that consumes an assessment JSON snapshot.
- The launcher still supports `-RunImprove` on `M365Collect` when you want to chain post-processing from a snapshot-only run.
- The default `M365` flow now centers on operator-facing top-level deliverables like `CustomerAssessmentReport.docx`, `EngineerActionPack.md`, and the workbook when the selected profile enables it.
- `AssessmentSnapshot.json`, `ImprovementPlan.json`, `*.manifest.json`, and `RemediationSnippets.ps1` are written under `Support`.
- `SolutionsEngineer` and `TenantToTenantMigration` include workbook output by default.
- Use `-IncludeLegacyAssessmentArtifacts` only when you explicitly want the older assessment HTML / questionnaire / PDF family.
- Use `-IncludeLegacyArtifacts` only when you explicitly want the older `Improve` CSV/Markdown planning artifacts.
- `Improve` can now take either the snapshot JSON path or the `*.manifest.json` path from the same run.
- Use `-LiveRefresh` when you want snapshot-plus-live-refresh behavior during `Improve`; it is a friendlier alias for `-UseGraphFallback`.
- The improvement plan combines derived assessment findings with built-in remediation heuristics.
- Rule IDs such as `ID-007`, `CA-012`, `DEV-006`, and `EX-007` are internal repo rule identifiers, not Microsoft-native control IDs.
- Use the taxonomy guide when you need to explain where a finding came from or how to trace it back to the supporting worksheet.

## Output profiles and detail levels
- `Presales`: `Minimum`
- `ExecutiveLevel`: `Minimum`
- `SolutionsEngineer`: `Operator`
- `Machine`: `Automation`
- `Geek`: `Geek`
- `TenantToTenantMigration`: `All`

`Combined` is still accepted as a compatibility alias for `Operator`.

## Another machine validation
When validating on another machine, start with the safest path first:

1. Run `Improve` against an existing known-good snapshot.
2. Run `M365Collect`.
3. Run `Improve` against the newly collected snapshot.

This separates environment/setup issues from tenant-authentication and collector issues.
