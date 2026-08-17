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

Module versions come from `dependencies.psd1` at the repository root, so every machine
resolves the same versions.

```powershell
.\tools\install-dependencies.ps1 -Scope Runtime
```

To also install app-registration and test/lint tooling:

```powershell
.\tools\install-dependencies.ps1 -Scope All
```

To check an existing machine without installing anything:

```powershell
.\tools\install-dependencies.ps1 -Scope All -Validate
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

When you pick an action, an interactive run asks for the tenant being assessed unless you started the launcher with `-TenantId`. That value is used to validate the cached Microsoft Graph session, so supply the tenant you actually intend to assess:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -TenantId '<tenant-guid>'
```

If the run stops reporting that Graph and Exchange Online are on different tenants, run `Disconnect-MgGraph`, open a fresh PowerShell session, and reconnect to the intended tenant.

## Next steps

- [Repository Quick Start](../../README.md#quick-start) — how to run your first assessment
- [App Registration Setup](app-registration-setup.md) — set up app-based auth
- [Certificate Auth Setup](certificate-auth-setup.md) — configure certificate auth
- [RUN.md](../../RUN.md) — full execution reference
