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

## Rotation guidance
- track certificate expiry before it becomes operationally urgent
- upload the replacement public cert before removing the old one
- validate a non-production or known-good run after rotation
- remove expired certificates from the app registration and the execution hosts after the replacement is verified

## Security guidance
- do not commit PFX files or private keys to the repo
- avoid sharing certificates between operators unless that is part of an approved operational model
- prefer user- or host-scoped certificate installation over loose files on disk
