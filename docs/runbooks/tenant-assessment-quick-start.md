# Tenant Assessment Quick Start

## 1. Install prerequisites

```powershell
.\tools\install-microsoft-modules.ps1
```

## 2. Run the assessment

Launch with the interactive menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Or pass `-Action` directly to skip the menu:

```powershell
# Standard full run
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full

# Full run with a specific export path
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -ExportPath 'C:\Assessment-Outputs'

# Full run without the improvement plan step
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -SkipImprove

# Active Directory assessment
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action AD
```

## 3. Find your outputs

Outputs go to `%LOCALAPPDATA%\Arraya\M365TenantAssessment\Outputs` by default, or to the folder you passed with `-ExportPath`.

| What you want | Where to look |
| ------------- | ------------- |
| Customer report | `Deliverables\*-Best Practices Assessment-*.docx` |
| Remediation roadmap | `Deliverables\*-Remediation Roadmap-*.docx` |
| Engineer action pack | `Deliverables\*-EngPack.md` |
| Snapshot for replay | `Support\*-AssessmentSnapshot.json` |

## Actions reference

| Action | What it does |
| ------ | ------------ |
| `Full` | Full assessment and improvement plan (standard run) |
| `Preflight` | Check connections and permissions only, no data collection |
| `Collect` | Collect data to a JSON snapshot, no report yet |
| `Report` | Generate reports from an existing snapshot |
| `Improve` | Build the improvement plan from an existing snapshot |
| `Compare` | Compare two snapshots side by side |
| `AD` | Active Directory assessment |

## If this is a new machine or new app registration

Run a preflight check first to validate connections and permissions before a long collection:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Preflight `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

## Output profiles

`SolutionsEngineer` is the default and covers most assessment runs. To use a different profile:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -OutputProfile SolutionsEngineer,ExecutiveLevel
```

Available profiles: `SolutionsEngineer`, `ExecutiveLevel`, `Presales`, `TenantToTenantMigration`, `Geek`, `Machine`

## Next steps

- Full execution guide: [RUN.md](../../RUN.md)
- App-based auth setup: [App Registration Setup](app-registration-setup.md)
- Understanding findings: [Improvement Plan Rule Taxonomy](improvement-plan-rule-taxonomy.md)
