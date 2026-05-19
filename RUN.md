# Run Modern Workplace Tenant Assessment

Use this guide when you are ready to execute the assessment and review the results. For repository background, see [README.md](README.md).

## Prerequisites

- PowerShell 7 or later
- Python 3.11 or later (required for the Python CLI — see [Python CLI](#python-cli) below)
- A local clone of this repository
- Access to the Microsoft 365 tenant you want to assess

Install the required PowerShell modules:

```powershell
.\tools\install-microsoft-modules.ps1
```

Install Python dependencies (one-time):

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Authentication Options

The assessment supports three authentication modes:

| Mode | Best for | Notes |
| ---- | -------- | ----- |
| `Interactive` | Guided consultant runs | Default. Sign in when prompted. Gives the broadest workload coverage. |
| `Certificate` | Unattended or repeatable runs | Recommended for production use. Supports Graph, Exchange, and Purview app auth. |
| `ClientSecret` | Compatibility cases | Graph app auth works; Exchange falls back to delegated sign-in. Not recommended as a permanent solution. |

If you are setting up app-based auth for the first time, complete these runbooks before running:

- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)

### Useful auth flags

- `-SkipAuth` — reuse workload sessions already connected in the current PowerShell session
- `-SkipPermissionPreflight` — bypass the startup access check and let collection run on a best-effort basis

### Authentication order in Interactive mode

The assessment connects workloads in this order: `Exchange Online → Purview → Microsoft Graph`. If a previous Graph session has already been established in the same shell, close that session and start fresh before retrying.

## Required Access and Permissions

### Operator roles (Interactive mode)

For full coverage, the operator account should have:

- `Exchange Administrator`
- `SharePoint Administrator`
- `Teams Administrator`
- `Security Reader`
- `Reports Reader`
- `Directory Readers` or `Global Reader`

### Microsoft Graph application permissions (app-based auth)

Grant these core permissions and consent them on the app registration:

- `Application.Read.All`
- `AuditLog.Read.All`
- `Channel.ReadBasic.All`
- `CrossTenantInformation.ReadBasic.All`
- `Device.Read.All`
- `Domain.Read.All`
- `Group.Read.All`
- `GroupMember.Read.All`
- `OnPremDirectorySynchronization.Read.All`
- `Organization.Read.All`
- `Policy.Read.All`
- `Reports.Read.All`
- `ReportSettings.Read.All`
- `RoleManagement.Read.All`
- `SecurityEvents.Read.All`
- `SharePointTenantSettings.Read.All`
- `Sites.Read.All`
- `Team.ReadBasic.All`
- `User.Read.All`

Optional — add these only when you need deeper Teams membership detail:

- `TeamMember.Read.All`

Optional extended enrichment (add only when explicitly needed):

- `DeviceManagementApps.Read.All`
- `DeviceManagementConfiguration.Read.All`
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementRBAC.Read.All`
- `DeviceManagementScripts.Read.All`
- `DeviceManagementServiceConfig.Read.All`
- `DirectoryRecommendations.Read.All`
- `IdentityRiskEvent.Read.All`
- `IdentityRiskyUser.Read.All`
- `LicenseAssignment.Read.All`
- `ServiceMessage.Read.All`
- `User.Export.All`
- `UserAuthenticationMethod.Read.All`

### Access outside Microsoft Graph

- **Exchange Online**: requires Exchange app-only access configured on the app registration for certificate auth
- **Purview**: uses `Connect-IPPSSession` from `ExchangeOnlineManagement`; supports delegated and certificate auth; client-secret auth is not supported
- **SharePoint**: certificate and client-secret runs use Microsoft Graph for SharePoint and OneDrive collection; SPO admin PowerShell is only used in interactive runs
- **Teams**: Teams PowerShell is delegated-only; member and guest count enrichment requires `TeamMember.Read.All`

## Common Commands

Run the interactive menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Run the standard full assessment:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full
```

Run with certificate auth:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

Check permissions before a long collection run:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Preflight `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

Collect data only (no report yet):

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Collect `
  -OutputProfile SolutionsEngineer `
  -ExportPath 'C:\Assessment-Outputs'
```

Generate reports from an existing snapshot:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Report
```

Generate the improvement plan from an existing snapshot:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Improve
```

Compare two saved snapshots:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Compare
```

Run without the improvement plan step:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -SkipImprove
```

Run a full assessment and write outputs to a specific folder:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -ExportPath 'C:\Assessment-Outputs'
```

Run with a specific output profile:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -OutputProfile SolutionsEngineer,ExecutiveLevel
```

### Prompts for some actions

- `Report` — asks for the saved JSON snapshot path
- `Improve` — asks for the snapshot or manifest JSON path and optional output folder
- `Compare` — asks for baseline and current snapshot paths

## Output Profiles

| Profile | Scope | Best for |
| ------- | ----- | -------- |
| `SolutionsEngineer` | Operator | Default. Standard consultant depth. |
| `ExecutiveLevel` | Minimum | Leadership summaries. |
| `Presales` | Minimum | Lighter presales discovery. |
| `TenantToTenantMigration` | All | Deep migration readiness. |
| `Geek` | Geek | Detailed engineer troubleshooting. |
| `Machine` | Automation | JSON-first automation and replay. |

Pass multiple profiles as a comma-separated list. When multiple profiles are supplied, the assessment runs once at the highest required scope.

## Default Output Location

If you do not pass `-ExportPath`, outputs are written to:

```text
%LOCALAPPDATA%\Arraya\M365TenantAssessment\Outputs
```

## What Gets Generated

A standard `Full` run writes files into two folders.

`Deliverables` — files to share or review first:

- `*-Microsoft 365 Tenant Best Practices Assessment-<date>.docx`
- `*-Microsoft 365 Remediation Roadmap-<date>.docx`
- `*-EngPack.md`
- `*.xlsx` when the selected profile enables workbook output

`Support` — machine-readable and troubleshooting files:

- `*-AssessmentSnapshot.json`
- `*-SolutionsEngineerEvidenceCoverage.json`
- `*-Plan.json`
- `*.manifest.json`
- `*-Snips.ps1`
- `Debugging\*`

To include older HTML, PDF, and questionnaire artifacts, add `-IncludeLegacyAssessmentArtifacts`.

## How To Read the Results

Start with the artifact that best matches your audience:

| Artifact | Audience |
| -------- | -------- |
| `Deliverables\*-Best Practices Assessment-*.docx` | Customer |
| `Deliverables\*-Remediation Roadmap-*.docx` | Customer or leadership |
| `Deliverables\*-EngPack.md` | Engineer or consultant |
| `Support\*-Plan.json` | Automation, filtering, downstream processing |

Recommended review order:

1. Open the customer assessment DOCX in `Deliverables` first.
2. Review the roadmap DOCX and engineer pack for implementation planning.
3. Keep `Support\*-Plan.json` as your baseline for comparison or downstream use.

## Validating on a New Machine

If you need to validate the workflow on another workstation or jump box:

1. Install PowerShell 7 or later.
2. Clone or copy the repo.
3. Install required modules: `.\tools\install-microsoft-modules.ps1`
4. If using certificate auth, import the `.pfx` and verify it has a private key.
5. Start with the lowest-risk path:
   - Run `Improve` against an existing known-good snapshot first.
   - Then run `Collect`.
   - Then run `Improve` against the new snapshot.

This separates environment setup issues from tenant authentication issues.

## Troubleshooting

**Missing module**: rerun `.\tools\install-microsoft-modules.ps1`

**Unexpected auth prompt**: check whether `-SkipAuth` should be used to reuse an existing session.

**Exchange or Purview auth fails after Graph**: close the PowerShell window and start a fresh session. The assessment authenticates `Exchange Online → Purview → Microsoft Graph` by design, but if Graph already ran in the shell, Exchange-family auth in that same session may fail.

**Run stops at the permission gate**: add `-SkipPermissionPreflight` to let collection continue on a best-effort basis.

**PDF output missing**: PDF generation requires a locally installed Chromium-based browser (Google Chrome or Microsoft Edge). The assessment still completes successfully without it.

**App-only auth fails**: review [App Registration Setup](docs/runbooks/app-registration-setup.md) and [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md).

## Python CLI

This branch ships a Python-based CLI (`cli.py`) that replaces the PowerShell launcher for everything except data collection. The split is intentional:

| Layer | Runtime | Why |
| --- | --- | --- |
| **Data collection** | PowerShell (unchanged) | Requires Exchange Online, Graph SDK, Purview, Teams PS modules |
| **Report generation** | Python | openpyxl, Jinja2, python-docx — no PS module dependencies |
| **Orchestration / comparison / improvement plan** | Python | Pure data processing, no M365 connection needed |

### Setup (one-time)

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Python CLI equivalents

| PowerShell | Python CLI | Notes |
| --- | --- | --- |
| `Start-M365TenantAssessment.ps1` (menu) | `python cli.py` | Same interactive menu |
| `-Action Full` | `python cli.py full` | Collect via PS, then report via Python |
| `-Action Preflight` | `python cli.py preflight` | Calls PS |
| `-Action Collect` | `python cli.py collect` | Calls PS |
| `-Action Report` | `python cli.py report <snapshot>` | Pure Python — no PS needed |
| `-Action Improve` | `python cli.py improve <snapshot>` | Pure Python |
| `-Action Compare` | `python cli.py compare <baseline> <current>` | Pure Python |
| `-Action AD` | `python cli.py ad` | Calls PS |

### Common Python CLI commands

Run the interactive menu:

```powershell
python cli.py
```

Collect with certificate auth:

```powershell
python cli.py collect `
  --auth-mode Certificate `
  --tenant-id '<tenant-guid>' `
  --client-id '<app-id>' `
  --cert-thumbprint '<thumbprint>'
```

Generate all reports from an existing snapshot (no M365 connection required):

```powershell
python cli.py report 'output\Contoso SE\Support\Contoso-Snap.json'
```

Generate reports with a specific output folder:

```powershell
python cli.py report 'Support\Contoso-Snap.json' --export-path 'C:\Assessment-Outputs'
```

Build the improvement plan only:

```powershell
python cli.py improve 'Support\Contoso-Snap.json' --output-dir 'C:\Assessment-Outputs\ImprovementPlan'
```

Compare two snapshots:

```powershell
python cli.py compare 'Support\Contoso-Snap-Jan.json' 'Support\Contoso-Snap-May.json'
```

Run a full assessment with cert auth:

```powershell
python cli.py full `
  --auth-mode Certificate `
  --tenant-id '<tenant-guid>' `
  --client-id '<app-id>' `
  --cert-thumbprint '<thumbprint>' `
  --profile SolutionsEngineer
```

### Python CLI output

The Python pipeline writes to a `Deliverables/` subfolder under the output path:

| File | Description |
| --- | --- |
| `<Tenant>-Tenant Details.xlsx` | Excel workbook (openpyxl) |
| `<Tenant>-Report.html` | HTML report (Jinja2) |
| `<Tenant>-Assessment.docx` | Word document (python-docx) |
| `ImprovementPlan/improvement-plan.json` | Structured finding rows |
| `ImprovementPlan/improvement-plan.md` | Markdown remediation plan |

### Getting help

```powershell
python cli.py --help
python cli.py report --help
```

## Additional References

- [Tenant Assessment Quick Start](docs/runbooks/tenant-assessment-quick-start.md)
- [Customer Execution Checklist](docs/runbooks/customer-execution-checklist.md)
- [Improvement Plan Rule Taxonomy](docs/runbooks/improvement-plan-rule-taxonomy.md)
