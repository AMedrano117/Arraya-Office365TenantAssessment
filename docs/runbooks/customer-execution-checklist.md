# Customer Execution Checklist

Use this checklist before and during a customer assessment engagement.

## Before the engagement

- [ ] Confirm the tenant name and primary domain with the customer contact
- [ ] Confirm the customer contact who will grant consent or provide credentials
- [ ] Confirm the assessment scope (full M365, AD only, migration readiness, or a subset)
- [ ] Confirm the output profile (`SolutionsEngineer` for most engagements; `TenantToTenantMigration` for migration work)
- [ ] Confirm the output folder or delivery path for deliverables

## App registration and auth

- [ ] App registration created in your Arraya Entra tenant (not the customer tenant)
- [ ] All required Graph permissions granted and admin-consented — see [App Registration Setup](app-registration-setup.md)
- [ ] Certificate generated, uploaded to the app registration, and installed on the execution machine — see [Certificate Auth Setup](certificate-auth-setup.md)
- [ ] Cert thumbprint, client ID, and tenant ID recorded for the run command
- [ ] Exchange app-only access configured on the app registration (required for Exchange collection in certificate mode)

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

## References

- [Tenant Assessment Quick Start](tenant-assessment-quick-start.md)
- [RUN.md](../../RUN.md)
- [App Registration Setup](app-registration-setup.md)
- [Certificate Auth Setup](certificate-auth-setup.md)
