# App Registration Setup

Use this runbook when you want to run the assessment with app-based auth. Certificate auth is the recommended production path. `ClientSecret` remains available as a compatibility mode for Graph-focused automation, but it is not equivalent to the certificate path in this workflow.

## Automated setup

[New-AssessmentAppRegistration.ps1](../../onboarding/New-AssessmentAppRegistration.ps1) creates the app registration, certificate, service principal, and Graph/Exchange application consent. Its `-AccessModel` parameter selects one of two mutually exclusive workload authorization models.

```powershell
.\onboarding\New-AssessmentAppRegistration.ps1 -TenantId '<tenant-guid>' -PermissionSet Standard -AccessModel LeastPrivilege
```

For the complete certificate-authenticated assessment, including Purview retention and DLP inventory, use the read-only Global Reader option:

```powershell
.\onboarding\New-AssessmentAppRegistration.ps1 -TenantId '<tenant-guid>' -PermissionSet Standard -AccessModel GlobalReader
```

- `LeastPrivilege` is the default. It assigns no Entra directory role; Exchange and Purview must use workload-scoped read-only RBAC role groups.
- `GlobalReader` assigns the read-only Global Reader directory role and removes a legacy Exchange Administrator assignment from the same service principal. It is broader but simpler and supports both Exchange Online and Security & Compliance PowerShell.

- `-PermissionSet` accepts `Core` (the 19 baseline permissions), `Standard` (adds `TeamMember.Read.All`, the default), or `Extended` (adds the optional enrichment permissions listed below).
- `-WhatIf` previews the changes without writing to the tenant.
- `-SkipExchange` provisions a Graph-only app. `-SkipCertificate` skips certificate generation.
- `-DisableWebAccountManager` forces the system browser when Web Account Manager sign-in fails with a window handle error, which happens in some embedded terminals and background hosts.
- Re-running is safe. An existing app of the same display name is reused, only missing permissions are added, and existing consent is left alone.

The signed-in admin needs Global Administrator, or Application Administrator plus Privileged Role Administrator, because the script both grants admin consent and assigns a directory role. Sign-in is interactive, so run it from a terminal you can respond in.

The manual steps below remain accurate if you would rather click through the portal.

## Create the app registration

1. Go to Entra admin center.
2. Open `App registrations`.
3. Create a new single-tenant app registration for the assessment.
4. Record:
   - Application (client) ID
   - Directory (tenant) ID

## Grant Microsoft Graph application permissions

Add these core Microsoft Graph application permissions and grant admin consent for the standard assessment:

- `Application.Read.All`
- `AuditLog.Read.All`
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
- `Channel.ReadBasic.All`
- `Team.ReadBasic.All`
- `User.Read.All`

Optional for deeper Teams membership enrichment:

- `TeamMember.Read.All`
  or
- `TeamMember.ReadWrite.All`

Keep delegated `User.Read` available for interactive sign-in scenarios, but the app-based baseline above is the primary collector requirement.

If you want optional extended enrichment beyond the standard assessment path, add these only when needed:

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

## Grant workload access outside Graph

### Exchange Online (required for certificate auth)

Exchange Online app-only access must be configured on the app registration. Without it, Exchange collection in certificate mode will fail.

### Purview compliance (required for certificate auth)

Purview retention and DLP collection uses `Connect-IPPSSession` from `ExchangeOnlineManagement`. In certificate mode this requires:

1. The app's service principal exists in the **target tenant** (created when the customer admin grants consent to the app).
2. Either workload-scoped read-only Purview RBAC is configured (`LeastPrivilege`) or **Global Reader** is assigned (`GlobalReader`). Do not stack Global Reader with Exchange Administrator.

This is a per-engagement step — it must be done in every customer tenant, not just in the Arraya tenant where the app is registered.

To assign the role in the customer tenant:

1. Sign in to the customer's Entra admin center.
2. Go to **Enterprise applications** and find the Arraya assessment app.
3. For the `GlobalReader` model, assign **Global Reader** to the service principal. Do not also assign **Exchange Administrator**.

`Exchange Administrator` alone does not expose Security & Compliance PowerShell cmdlets to an app-only session. Without a supported Purview role, `Connect-IPPSSession` can accept the certificate while exposing no retention or DLP cmdlets.

`Global Reader` is broad but read-only. For `LeastPrivilege`, create workload custom role groups containing only the Exchange inventory and Purview retention/DLP view roles required by the collector, then add the service principal to those groups.

> **Note:** Client secret auth is not supported for Purview compliance PowerShell. Certificate auth is the only supported app-based path for Purview collection.

### SharePoint and Teams

- SharePoint and OneDrive standard collection use Microsoft Graph in the app-based workflow.
- Teams PowerShell is delegated-only; treat Teams detail as best-effort when running non-interactively.
- Without `TeamMember.Read.All` or `TeamMember.ReadWrite.All`, the assessment still collects team and channel inventory but skips member-count and guest-count enrichment.

## Recommended operator/admin roles for setup

- `Application Administrator`
- `Cloud Application Administrator`
- or `Global Administrator`

## Keep secrets and certificates out of the repo

- Do not store client secrets, PFX files, or exported private keys in this repository.
- Store certificate material in approved secure storage and import it on the execution machine only.
- If you must use `ClientSecret`, pass it through approved secret storage or an environment variable rather than embedding it directly in scripts or shell history.

## Next step

After the app registration exists, continue with [certificate-auth-setup.md](certificate-auth-setup.md) if you plan to use certificate auth.
