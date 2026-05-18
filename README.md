# Modern Workplace Tenant Assessment

Microsoft 365 tenant assessment tooling for consultants and engineers who need to collect tenant data, build customer-ready reports, and turn the findings into a practical remediation plan.

The main entry point is one PowerShell launcher. Run it without parameters when you want the guided menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Most day-to-day runs start with the standard Microsoft 365 assessment:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full
```

That standard run collects the tenant snapshot, generates customer deliverables, and builds the remediation outputs used by the engineer.

## When To Use This

Use this repo when you need to:

- Assess Microsoft 365 tenant configuration, security posture, collaboration settings, and governance signals.
- Produce customer-facing assessment and remediation documents.
- Save a reusable JSON snapshot for replay, comparison, or later reporting.
- Run preflight checks before a longer collection.
- Compare two assessment snapshots over time.

For the full operator guide, see [RUN.md](RUN.md).

## Quick Start

1. Install PowerShell 7 or later.
2. Install the Microsoft modules used by the assessment:

```powershell
.\tools\install-microsoft-modules.ps1
```

3. Start with a preflight check if this is a new tenant, new workstation, or new app registration:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Preflight
```

4. Run the assessment:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full
```

If you prefer the guided flow, run the launcher with no `-Action` and choose from the menu.

## Common Actions

| Action | Use it when you want to |
| --- | --- |
| `Full` | Run the standard Microsoft 365 assessment and remediation workflow. |
| `Preflight` | Check authentication, connection, and permission readiness without collecting tenant data. |
| `Collect` | Collect the tenant snapshot only, so reporting can happen later. |
| `Report` | Generate reports from an existing snapshot. |
| `Improve` | Build remediation outputs from an existing snapshot or manifest. |
| `Compare` | Compare two saved tenant snapshots. |
| `AD` | Run the Active Directory assessment workflow. |

The older assessment-only behavior is still available:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -SkipImprove
```

## Authentication

The launcher supports three auth modes:

| Mode | Best for | Notes |
| --- | --- | --- |
| `Interactive` | Guided consultant runs | Default mode. Sign in when prompted. |
| `Certificate` | Repeatable app-based runs | Recommended for unattended or production-style execution. |
| `ClientSecret` | Compatibility cases | Graph app auth works, but some workloads fall back or are skipped. Prefer certificate auth when possible. |

Certificate example:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

Client secret example:

```powershell
$clientSecret = Read-Host 'Client secret' -AsSecureString

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -AuthMode ClientSecret `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -ClientSecretSecure $clientSecret
```

Setup references:

- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)

## Output Profiles

Profiles tune how much data is collected and which deliverables are created.

| Profile | Best for |
| --- | --- |
| `SolutionsEngineer` | Standard consultant run and the default choice for most assessments. |
| `ExecutiveLevel` | Leadership-friendly summary depth. |
| `Presales` | Lighter discovery and presales posture review. |
| `TenantToTenantMigration` | Deep migration readiness and cutover planning. |
| `Geek` | Detailed engineer troubleshooting. |
| `Machine` | JSON-first automation, replay, and downstream processing. |

You can request more than one profile in a single run:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -OutputProfile SolutionsEngineer,ExecutiveLevel
```

## What Gets Created

A standard `Full` run writes the main working files into two folders.

`Deliverables` contains the files you are most likely to share or review first:

- `*-Microsoft 365 Tenant Best Practices Assessment-<date>.docx`
- `*-Microsoft 365 Remediation Roadmap-<date>.docx`
- `*-EngPack.md`
- `*.xlsx` when the selected profile enables workbook output

`Support` contains the machine-readable and troubleshooting artifacts:

- `*-AssessmentSnapshot.json`
- `*-SolutionsEngineerEvidenceCoverage.json`
- `*-Plan.json`
- `*.manifest.json`
- `*-Snips.ps1`
- `Debugging\*`

By default, outputs are written under:

```text
%LOCALAPPDATA%\Arraya\M365TenantAssessment\Outputs
```

Pass `-ExportPath` when you want to choose the output folder:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -ExportPath 'C:\Assessment-Outputs'
```

## Helpful Runbooks

- [RUN.md](RUN.md)
- [Customer Execution Checklist](docs/runbooks/customer-execution-checklist.md)
- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)
- [Improvement Plan Rule Taxonomy](docs/runbooks/improvement-plan-rule-taxonomy.md)

## Development

Install the maintainer tooling:

```powershell
.\tools\install-microsoft-modules.ps1 -IncludeDevTools
```

Run the usual validation checks:

```powershell
.\tools\invoke-scriptanalyzer.ps1
.\tools\run-pester.ps1
```
