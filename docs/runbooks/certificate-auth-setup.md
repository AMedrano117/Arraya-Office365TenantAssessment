# Certificate Auth Setup

Use this runbook when you want to run the assessment with `-AuthMode Certificate`.

## Create or obtain the certificate
Use an approved internal PKI or generate a certificate suitable for app authentication. The certificate must include a private key and be exportable if you need to move it to another execution host.

Minimum practical requirements:
- private key present
- usable for client authentication
- stored in the current user or local machine certificate store on the execution host

## Upload the public certificate to the app registration
1. Open the Entra app registration created for the assessment.
2. Go to `Certificates & secrets`.
3. Upload the public certificate.
4. Record the thumbprint of the installed certificate on the execution host.

## Install the certificate on the execution machine
Import the certificate with private key on the workstation, jump box, or automation host that will run the assessment.

Verify it locally:

```powershell
Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My |
  Where-Object Thumbprint -eq '<cert-thumbprint>'
```

## Example run
```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

## Certificate auth coverage in this workflow
Certificate auth in this repo is not limited to Graph.

- Microsoft Graph uses `TenantId + ClientId + CertificateThumbprint`
- Exchange Online uses `AppId + Organization + CertificateThumbprint`
- Purview retention and DLP policy collection uses compliance PowerShell through `Connect-IPPSSession` with `AppId + Organization + CertificateThumbprint`
- SharePoint and OneDrive collection use Microsoft Graph in certificate mode; the workflow does not import the SharePoint Online module or call `Connect-SPOService` for certificate auth

Notes:
- for Purview collection, `Organization` is the tenant initial domain, not the GUID tenant ID
- `ExchangeOnlineManagement` must be installed because it provides both Exchange Online connection support and `Connect-IPPSSession`
- client-secret auth is not supported for the Purview compliance PowerShell portion of this workflow

## Rotation guidance
- track certificate expiry before it becomes operationally urgent
- upload the replacement public cert before removing the old one
- validate a non-production or known-good run after rotation
- remove expired certificates from the app registration and the execution hosts after the replacement is verified

## Security guidance
- do not commit PFX files or private keys to the repo
- avoid sharing certificates between operators unless that is part of an approved operational model
- prefer user- or host-scoped certificate installation over loose files on disk

## Using `-SkipAuth` to reuse an existing session

If you have already connected all required workloads in the current PowerShell session, you can pass `-SkipAuth` to bypass the assessment's authentication bootstrap and reuse those sessions.

**Required connection order before using `-SkipAuth`:**

1. `ExchangeOnlineManagement` (`Connect-ExchangeOnline`) — must be loaded and connected **before** the Microsoft Graph SDK. Loading ExchangeOnlineManagement after the Graph SDK can cause MSAL assembly version conflicts that silently break Exchange collection in PowerShell 7.
2. Microsoft Graph SDK (`Connect-MgGraph`) — connect with the required scopes for the active profile.
3. SharePoint Online and Teams — connect these if your active output profile includes SharePoint or Teams collection.

Example:

```powershell
# 1. Exchange first
Connect-ExchangeOnline -AppId '<app-id>' -Organization '<tenant>.onmicrosoft.com' -CertificateThumbprint '<thumbprint>'

# 2. Graph SDK second
Connect-MgGraph -TenantId '<tenant-guid>' -ClientId '<app-id>' -CertificateThumbprint '<thumbprint>'

# 3. Then run the assessment with SkipAuth
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365 -SkipAuth
```

`-SkipAuth` validates only the workloads the active output profile actually needs. Missing a required workload connection surfaces as a collector failure, not an auth setup error.
