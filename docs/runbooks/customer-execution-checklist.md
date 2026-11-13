# Customer Execution Checklist

Use this checklist before and during a customer assessment engagement.

## Before the engagement

- [ ] Confirm the tenant name and primary domain with the customer contact
- [ ] Confirm the customer contact who will sign in to create the app registration (Global Administrator, or Application Administrator plus Privileged Role Administrator), or who will provide a delegated account for an interactive run
- [ ] Confirm the assessment scope (full M365, AD only, migration readiness, or a subset)
- [ ] Confirm the output profile (`SolutionsEngineer` for most engagements; `TenantToTenantMigration` for migration work)
- [ ] Confirm the output folder or delivery path for deliverables

## App registration and auth

Skip this section for an interactive (delegated) run. Run the preflight instead (see [Day of the run](#day-of-the-run)).

The assessment app is a **single-tenant app registration in the customer's tenant**. Each customer gets their own app; there is no shared Arraya app.

- [ ] Access model agreed with the customer: `GlobalReader` (complete run, read-only) or `LeastPrivilege` (customer grants workload role groups)
- [ ] Changes previewed with `-WhatIf`:

```powershell
.\onboarding\New-AssessmentAppRegistration.ps1 -TenantId '<tenant-guid>' -AccessModel GlobalReader -WhatIf
```

- [ ] App registration created in the customer tenant, signed in as the customer admin:

```powershell
.\onboarding\New-AssessmentAppRegistration.ps1 -TenantId '<tenant-guid>' -AccessModel GlobalReader
```

- [ ] `AppId`, `CertificateThumbprint`, `TenantId`, and `CertificateNotAfter` recorded from the script output
- [ ] Run from the machine and Windows user that will run the assessment. The script creates a new self-signed certificate in `Cert:\CurrentUser\My` on every run; re-run with `-SkipCertificate` to update permissions without creating another certificate. See [Certificate Auth Setup](certificate-auth-setup.md)
- [ ] For `LeastPrivilege` only: Exchange and Purview read-only role groups granted to the app's service principal. See [App Registration Setup](app-registration-setup.md)
- [ ] Optional: Exchange app-only access confirmed with `.\onboarding\Test-ExchangeAppOnlyConnection.ps1`

## Day of the run

- [ ] Run a preflight check before the full collection to validate connections and permissions:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Preflight `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'
```

- [ ] Preflight shows no blocking access gaps
- [ ] Export path confirmed and has write access
- [ ] Run the full assessment:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action Full `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -ExportPath 'C:\Assessment-Outputs\<customer-name>'
```

## After the run

- [ ] Confirm `Deliverables\` folder was created and contains the DOCX files and engineer pack
- [ ] Open the customer assessment DOCX and verify the tenant name and date are correct
- [ ] Review the engineer pack for any finding anomalies before presenting
- [ ] Copy `Support\*-AssessmentSnapshot.json` to a safe location for replay or comparison later
- [ ] Remove any customer data from the execution machine after delivery if required by the engagement agreement
- [ ] App-based runs: agree with the customer whether to keep the app for follow-up runs. If not, delete the app registration in the customer tenant and remove the certificate from the execution machine (see [After the engagement](../../README.md#10-after-the-engagement))

## References

- [Operator walkthrough](../../README.md)
- [RUN.md](../../RUN.md)
- [App Registration Setup](app-registration-setup.md)
- [Certificate Auth Setup](certificate-auth-setup.md)
