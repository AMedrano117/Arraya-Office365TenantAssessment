# Modern Workplace Tenant Assessment

PowerShell automation for Microsoft 365 tenant assessments, reporting, and improvement planning.

## Quick Start

Run the launcher and pick an action from the interactive menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

No module setup required. The launcher imports everything it needs automatically.

For detailed setup and execution guidance see [RUN.md](RUN.md) and the [Tenant Assessment Quick Start](docs/runbooks/tenant-assessment-quick-start.md).

## Actions

| Action | What it does |
| ------ | ------------ |
| `Full` | Full assessment + improvement plan (standard run) |
| `Preflight` | Test connections and permissions only, no collection |
| `Collect` | Collect data to a JSON snapshot, no report yet |
| `Report` | Generate reports from an existing JSON snapshot |
| `Improve` | Build improvement plan documents from a snapshot |
| `Compare` | Compare two snapshots side by side |
| `AD` | Active Directory assessment |

Pass `-Action` directly to skip the menu:

```powershell
# Standard full run
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full

# Full run, skip the improvement plan step
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -SkipImprove

# Preflight only — verify access before a long collection
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Preflight `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'

# Collect data to a snapshot (no report)
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Collect `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -OutputProfile SolutionsEngineer

# Generate reports from an existing snapshot
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Report `
  -OutputProfile ExecutiveLevel
```

## Auth Modes

| Mode | When to use |
| ---- | ----------- |
| Interactive (default) | Day-to-day operator runs with delegated sign-in |
| Certificate | Unattended/production runs; recommended for app-only auth |
| Client secret | Compatibility path; Exchange falls back to delegated sign-in |

If you do not pass `-AuthMode`, the assessment defaults to interactive sign-in.

Certificate auth example:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -ExportPath 'C:\Assessment-Outputs' `
  -OutputProfile SolutionsEngineer,ExecutiveLevel
```

Client secret example (use `-ClientSecretSecure` to avoid plain-string secrets in shell history):

```powershell
$clientSecret = Read-Host 'Client secret' -AsSecureString
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -AuthMode ClientSecret `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -ClientSecretSecure $clientSecret
```

Useful flags:

- `-SkipAuth` — reuse already-connected workload sessions in the current shell
- `-SkipPermissionPreflight` — skip startup access validation and let collection run on a best-effort basis

For app registration requirements see [App Registration Setup](docs/runbooks/app-registration-setup.md) and [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md).

## Output Profiles

| Profile | Scope | Use case |
| ------- | ----- | -------- |
| `SolutionsEngineer` (default) | Operator | Standard consultant depth |
| `ExecutiveLevel` | Minimum | Leadership and presales |
| `Presales` | Minimum | Lighter presales story |
| `TenantToTenantMigration` | All | Migration readiness and cutover |
| `Geek` | Geek | Deep engineer troubleshooting |
| `Machine` | Automation | JSON-focused automation and replay |

Pass multiple profiles as a comma-separated list:

```powershell
-OutputProfile SolutionsEngineer,ExecutiveLevel
```

## What Gets Generated

A standard `Full` run produces:

**Customer-facing deliverables** (in the run output folder):

- `*-Microsoft 365 Tenant Best Practices Assessment-<date>.docx` — primary customer report
- `*-Microsoft 365 Remediation Roadmap-<date>.docx` — companion executive roadmap

**Engineer-facing deliverable**:

- `*-EngPack.md` — prioritized remediation action pack

**Support artifacts** (under `Support\`):

- `*-AssessmentSnapshot.json` — tenant data snapshot (used by `Improve`, `Report`, and `Compare`)
- `*-ImprovementPlan.json` — machine-readable remediation payload
- `*.manifest.json` — artifact index for the run

Legacy HTML, PDF, and workbook outputs are also available; see [RUN.md](RUN.md) for details.

## Operator References

- [RUN.md](RUN.md) — full execution guide
- [Tenant Assessment Quick Start](docs/runbooks/tenant-assessment-quick-start.md)
- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)
- [Improvement Plan Rule Taxonomy](docs/runbooks/improvement-plan-rule-taxonomy.md)

## Development

1. Install PowerShell 7+.
2. Run `tools/bootstrap-dev.ps1`.
3. Validate with `tools/invoke-scriptanalyzer.ps1` and `tools/run-pester.ps1`.
