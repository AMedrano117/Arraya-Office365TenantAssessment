# Modern Workplace Tenant Assessment

PowerShell tooling that collects read-only configuration data from a Microsoft 365 tenant and turns it into customer-ready reports, a remediation roadmap, and a reusable JSON snapshot.

This page walks an operator through a run from start to finish:

1. [Before you start](#1-before-you-start)
2. [Install dependencies](#2-install-dependencies)
3. [Choose how to sign in](#3-choose-how-to-sign-in)
4. [Set up delegated sign-in (run preflight)](#4-set-up-delegated-sign-in-run-preflight)
5. [Set up app-based sign-in (app registration)](#5-set-up-app-based-sign-in-app-registration)
6. [Sign-in examples](#6-sign-in-examples)
7. [Run the assessment](#7-run-the-assessment)
8. [Check the run was complete](#8-check-the-run-was-complete)
9. [Find the outputs](#9-find-the-outputs)
10. [After the engagement](#10-after-the-engagement)
11. [Other things you can do](#11-other-things-you-can-do)
12. [Troubleshooting](#12-troubleshooting)

Every command below runs from the repository root in **PowerShell 7** (`pwsh`), not Windows PowerShell 5.1.

## 1. Before you start

You need:

- Windows with PowerShell 7 or later. App registration setup creates a certificate with `New-SelfSignedCertificate`, which is Windows-only.
- A local clone of this repository.
- The tenant ID (GUID) of the tenant you are assessing.
- An account in that tenant with the roles listed in step 4 (delegated) or step 5 (app-based).
- Optional: Microsoft Edge or Google Chrome installed, for the PDF outputs. Without one, the run still completes but writes no PDFs.

Open a **new PowerShell 7 window for each tenant**. Sign-in sessions from an earlier run can belong to a different tenant, or break the Exchange sign-in. The assessment stops if it detects a tenant mismatch.

If you copied the repository as a downloaded ZIP instead of cloning it, Windows may block the scripts. Unblock them once from the repository root:

```powershell
Get-ChildItem -Recurse -File | Unblock-File
```

The assessment only reads tenant data. The one step that changes the tenant is creating the app registration in step 5, and that step is optional.

## 2. Install dependencies

Install the PowerShell modules the assessment needs. Versions come from [dependencies.psd1](dependencies.psd1).

```powershell
.\tools\install-dependencies.ps1 -Scope Runtime
```

If you will also create the app registration (step 5), include the onboarding modules:

```powershell
.\tools\install-dependencies.ps1 -Scope Runtime, Onboarding
```

Check what is installed without changing anything:

```powershell
.\tools\install-dependencies.ps1 -Scope Runtime -Validate
```

The installer leaves a module alone if an installed version already meets the requirement. To upgrade to the newest allowed versions, add `-Force`.

## 3. Choose how to sign in

| Mode | What it is | Use it when |
| --- | --- | --- |
| `Interactive` (default) | Delegated. You sign in as yourself and the assessment uses your roles. | A consultant is running the assessment hands-on. Nothing needs to be created in the tenant. |
| `Certificate` | App-based. An app registration authenticates with a certificate on your machine. | You want repeatable runs with fixed, consented permissions, including Exchange and Purview. Teams PowerShell supports only delegated sign-in, so Teams policy detail is best-effort in this mode. |
| `ClientSecret` | App-based with a secret. | Compatibility only. Exchange falls back to delegated sign-in and Purview is not supported. Prefer `Certificate`. |

Delegated: go to step 4. App-based: go to step 5.

## 4. Set up delegated sign-in (run preflight)

No setup in the tenant is needed. Your account needs these roles for full coverage:

- `Global Reader` (or `Directory Readers`)
- `Exchange Administrator`
- `SharePoint Administrator`
- `Teams Administrator`
- `Security Reader`
- `Reports Reader`

If your tenant uses PIM, activate the roles before you start.

Then run a preflight. It signs you in to each workload and checks your access, without collecting any tenant data:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Preflight -TenantId '<tenant-guid>'
```

Fix anything the preflight reports as missing, then continue to [step 7](#7-run-the-assessment).

## 5. Set up app-based sign-in (app registration)

`onboarding\New-AssessmentAppRegistration.ps1` creates everything app-based sign-in needs, in the tenant you are assessing:

- a single-tenant app registration (default name `Arraya M365 Tenant Assessment`)
- a self-signed certificate, uploaded to the app
- the service principal
- admin consent for the read-only Microsoft Graph permissions and Exchange `Exchange.ManageAsApp`
- the directory role for the access model you choose

You sign in interactively. Your account needs `Global Administrator`, or `Application Administrator` plus `Privileged Role Administrator`, because the script grants admin consent and assigns a directory role.

Preview the changes first. `-WhatIf` changes nothing in the tenant:

```powershell
.\onboarding\New-AssessmentAppRegistration.ps1 -TenantId '<tenant-guid>' -AccessModel GlobalReader -WhatIf
```

Then create it:

```powershell
.\onboarding\New-AssessmentAppRegistration.ps1 -TenantId '<tenant-guid>' -AccessModel GlobalReader
```

Choose the access model:

| `-AccessModel` | What it does | Use it when |
| --- | --- | --- |
| `GlobalReader` | Assigns the read-only `Global Reader` role. | You want the complete certificate-based run, including Exchange and Purview retention and DLP, with no further setup. |
| `LeastPrivilege` (default) | Assigns no directory role. | You will grant read-only Exchange and Purview access yourself through workload role groups. Until you do, those workloads cannot be collected. See [App Registration Setup](docs/runbooks/app-registration-setup.md). |

The script prints a result object. Record these values for the sign-in examples in step 6:

- `AppId`: your `-ClientId`
- `CertificateThumbprint`: your `-CertificateThumbprint`
- `TenantId`
- `CertificateNotAfter`: when the certificate expires (2 years by default)

> **Certificate warning.** By default the script creates a **brand-new self-signed certificate on every run** and adds it to the app.
>
> - The certificate, including its private key, is saved only in **your Windows user's certificate store** (`Cert:\CurrentUser\My`) on **this machine**. Certificate sign-in works only from this machine and Windows user, unless you export the certificate and import it elsewhere.
> - Running the script again for the same app adds another certificate to it. To update permissions without creating another certificate, re-run with `-SkipCertificate`.
> - To use your organization's own certificate, run with `-SkipCertificate` and upload the public certificate yourself. See [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md).
> - Never commit `.pfx` files or private keys to this repository.

Other useful options:

- `-PermissionSet Core | Standard | Extended`. `Standard` is the default. `Extended` adds optional Intune, risk, and authentication-method permissions.
- `-CertificateValidityYears 1-5` sets how long the certificate is valid (default 2).
- `-DisableWebAccountManager` uses the system browser instead of the Windows sign-in broker, for terminals that fail with "A window handle must be configured".
- `-UseDeviceAuthentication` signs in with a device code.

Re-running the script is safe. It reuses an app with the same display name, adds only missing permissions, and skips consent that is already granted.

Optionally, confirm Exchange app-only access works before a full run:

```powershell
.\onboarding\Test-ExchangeAppOnlyConnection.ps1 `
  -AppId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -CustomerOrganization '<tenant>.onmicrosoft.com'
```

Then run a preflight with the certificate (see step 6) before the full assessment.

## 6. Sign-in examples

Replace the placeholders with your own values. `<tenant-guid>` is the directory (tenant) ID, `<app-id>` is the app's client ID, and `<cert-thumbprint>` is the 40-character thumbprint from step 5.

### Interactive (delegated)

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Preflight -TenantId '<tenant-guid>'
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -TenantId '<tenant-guid>'
```

A browser or sign-in prompt opens for each workload. Without `-TenantId` you are asked for the tenant at the start.

### Certificate (app-based)

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Preflight `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

Change `-Action Preflight` to `-Action Full` for the real run. To confirm the certificate is on this machine:

```powershell
Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My | Where-Object Thumbprint -eq '<cert-thumbprint>'
```

### Client secret (compatibility only)

```powershell
$clientSecret = Read-Host 'Client secret' -AsSecureString

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -AuthMode ClientSecret `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -ClientSecretSecure $clientSecret
```

Read the secret with `Read-Host -AsSecureString` as shown. Never type it into a command line or save it in a script.

### Reuse sessions you already connected

If you connected the workloads yourself, pass `-SkipAuth`. Connect Exchange Online **before** Microsoft Graph, or the two modules can conflict:

```powershell
Connect-ExchangeOnline -AppId '<app-id>' -Organization '<tenant>.onmicrosoft.com' -CertificateThumbprint '<cert-thumbprint>'
Connect-MgGraph -TenantId '<tenant-guid>' -ClientId '<app-id>' -CertificateThumbprint '<cert-thumbprint>'

.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -SkipAuth
```

## 7. Run the assessment

The standard run collects the tenant data, writes the customer deliverables, and builds the remediation outputs. Add the sign-in parameters from step 6:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -TenantId '<tenant-guid>'
```

Prefer a menu? Run the launcher with no parameters:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

A typical run takes 15 to 25 minutes. Exchange mailbox enumeration is usually the longest step. The console shows each section as it runs: Tenant Overview, Identity, Exchange, Collaboration, Endpoint, and Governance.

What to expect during the run:

- **Sign-in order.** In an interactive run you sign in to Exchange Online, then Purview, then Microsoft Graph, and possibly SharePoint and Teams. Answer each prompt; the run waits for you.
- **An expected SharePoint 403.** In an interactive run, the Graph site inventory call always returns 403, because Microsoft allows it only with app permissions. The assessment then switches to the SharePoint Online module automatically. This is not a failure.
- **The permission gate.** If the startup access check finds a blocking gap, the run stops before collecting. Fix the access and run again. To collect whatever your access allows, add `-SkipPermissionPreflight`; anything that can't be collected shows as missing in the outputs.
- **Keep the window open**, and keep the machine awake until the run finishes.

### Choose an output profile

`-OutputProfile` decides how much is collected and which deliverables are written. The default is `SolutionsEngineer`.

| Profile | Use it for |
| --- | --- |
| `SolutionsEngineer` | The standard consultant assessment. |
| `ExecutiveLevel` | A leadership-level summary. |
| `Presales` | Migration scoping: sizing, complexity, and quote readiness. |
| `TenantToTenantMigration` | Exchange and collaboration cutover preparation and the cutover pack. Not yet a complete migration assessment. |
| `Geek` | Detailed engineer troubleshooting. |
| `Machine` | JSON-first output for automation and replay. |

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -TenantId '<tenant-guid>' -OutputProfile Presales
```

Run `Presales` and `TenantToTenantMigration` on their own. Combining either one with another profile collects less migration evidence. [RUN.md](RUN.md#output-profiles) explains what each profile collects. The [tenant-to-tenant review](docs/t2t-migration-review/T2T-Review-Summary.md) lists what the migration profile does not cover yet.

### Choose where outputs go

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action Full -TenantId '<tenant-guid>' -ExportPath 'C:\Assessment-Outputs'
```

## 8. Check the run was complete

Before you share anything, check that the run collected what you expected:

1. Scroll the console for steps reported as `Failed` or `Skipped`. A skipped step is usually the output profile at work. A failed step usually means missing access or a service error.
2. Open the run log in `Support\Debugging`. The console prints its location at the end of the run.
3. In migration workbooks, missing evidence is labelled `Needs Data`, never shown as absent or ready. The `MigrationCriteriaRegister` worksheet lists each gap with the action that closes it.

If a workload failed because of access, fix the access and run again. A partial run is still usable, but tell the customer which areas are incomplete.

## 9. Find the outputs

Without `-ExportPath`, outputs are written under:

```text
%LOCALAPPDATA%\Arraya\M365TenantAssessment\Outputs
```

`Deliverables` holds the files you review and share:

- `*-Best Practices Assessment-*.docx`: the customer-facing assessment report
- `*-Remediation Roadmap-*.docx`: the phased action plan (0-30, 31-60, 61-90 days)
- `*-EngPack.md`: the engineer action pack with findings, evidence, and PowerShell snippets
- `*.xlsx`: the workbook, when the profile produces one

`Support` holds the machine-readable files:

- `*-AssessmentSnapshot.json`: the full tenant snapshot. Keep it. You can regenerate reports from it without reconnecting to the tenant.
- `*-Plan.json`: structured findings
- `*-Snips.ps1`: remediation snippets
- `Debugging\*`: logs for troubleshooting a run

Outputs contain customer tenant data. Store them in approved customer locations, never in this repository.

## 10. After the engagement

- **Keep** `Support\*-AssessmentSnapshot.json` in the approved customer location, for later reports or comparisons.
- **Delete** the local copies of the outputs if the engagement agreement requires it.
- **App-based runs:** once the customer no longer needs repeat runs, remove the access you created:
  - Delete the app registration (default name `Arraya M365 Tenant Assessment`) in the customer's Entra admin center, under **App registrations**. This also removes its consented permissions and role assignment.
  - Remove the certificate from the machine that ran the assessment:

    ```powershell
    Remove-Item -Path 'Cert:\CurrentUser\My\<cert-thumbprint>'
    ```

  Agree with the customer who removes the app and when. The customer may want to keep it for a follow-up assessment.

## 11. Other things you can do

| Action | What it does |
| --- | --- |
| `-Action Collect` | Collect the snapshot only, and report later. |
| `-Action Report` | Rebuild reports from a saved snapshot. Prompts for the JSON path. |
| `-Action Improve` | Rebuild the remediation outputs from a saved snapshot. |
| `-Action Compare` | Compare two saved snapshots, such as before and after remediation. |
| `-Action AD` | Run the on-premises Active Directory assessment. |
| `-SkipImprove` | Run `Full` without building the remediation outputs. |

## 12. Troubleshooting

| Symptom | What to do |
| --- | --- |
| The run stops because Graph and Exchange are connected to different tenants | Run `Disconnect-MgGraph`, open a new PowerShell 7 window, and run again with `-TenantId`. |
| Sign-in fails with "A window handle must be configured" | Run from a normal PowerShell 7 window. For the app registration script, add `-DisableWebAccountManager`. |
| Certificate sign-in fails | Confirm the thumbprint is in `Cert:\CurrentUser\My` on this machine and matches a certificate on the app registration. |
| Exchange or Purview data is missing in a certificate run | Use `-AccessModel GlobalReader`, or finish the `LeastPrivilege` role group setup. Client-secret runs never collect Purview. |
| A module is missing or too old | Run `.\tools\install-dependencies.ps1 -Scope Runtime -Validate`, then install with `-Force`. |
| Exchange or Purview sign-in fails after Graph has already connected in the same window | Close the window and run again in a new PowerShell 7 window. |
| SharePoint fails before a sign-in window opens | The tenant's SharePoint name may differ from its `onmicrosoft.com` name. See [RUN.md](RUN.md#troubleshooting). |
| Scripts are blocked or "not digitally signed" | Unblock them, as shown in [step 1](#1-before-you-start). |
| No PDF files | Install Microsoft Edge or Google Chrome. The other outputs are unaffected. |

More detail: [RUN.md](RUN.md#troubleshooting).

## More documentation

- [RUN.md](RUN.md): the full operator reference, including permissions, profiles, every output, and how to read the results
- [Customer Execution Checklist](docs/runbooks/customer-execution-checklist.md)
- [App Registration Setup](docs/runbooks/app-registration-setup.md): the full permission list and the manual portal steps
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md): using your own certificate, rotation, and moving it to another machine
- [Improvement Plan Rule Taxonomy](docs/runbooks/improvement-plan-rule-taxonomy.md): how to read and explain each finding

## For maintainers

```powershell
.\tools\install-dependencies.ps1 -Scope All
.\tools\invoke-scriptanalyzer.ps1
.\tools\run-pester.ps1
```

The assessment is moving from the legacy orchestration script (`src/scripts/migrated/legacy/Get-FullTenantReportDetails.ps1`) toward modules under `src/modules/`. New collection logic goes into module functions.
