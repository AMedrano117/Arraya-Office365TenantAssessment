# Engineer Onboarding

Use this guide when setting up this repo on a new workstation for the first time.

## 1. Install PowerShell 7

Download and install PowerShell 7 or later from the [PowerShell GitHub releases](https://github.com/PowerShell/PowerShell/releases).

## 2. Clone the repository

```powershell
git clone <repo-url>
cd <repo-folder>
```

## 3. Install required modules

```powershell
.\tools\install-microsoft-modules.ps1
```

To also install test and lint tooling:

```powershell
.\tools\install-microsoft-modules.ps1 -IncludeDevTools
```

## 4. Validate the setup

Run the script analyzer and test suite:

```powershell
.\tools\invoke-scriptanalyzer.ps1
.\tools\run-pester.ps1
```

All Pester tests should pass before working on any changes.

## 5. Validate the launcher

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

The menu should open without errors.

## Next steps

- [Tenant Assessment Quick Start](tenant-assessment-quick-start.md) — how to run your first assessment
- [App Registration Setup](app-registration-setup.md) — set up app-based auth
- [Certificate Auth Setup](certificate-auth-setup.md) — configure certificate auth
- [RUN.md](../../RUN.md) — full execution reference
