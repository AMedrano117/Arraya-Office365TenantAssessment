# App Registration Setup

Use this runbook when you want to run the assessment with app-based auth. Certificate auth is the recommended production path. `ClientSecret` remains available as a compatibility mode for Graph-focused automation, but it is not equivalent to the certificate path in this workflow.

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
2. The **Exchange Administrator** role is assigned to that service principal in the target tenant.

This is a per-engagement step — it must be done in every customer tenant, not just in the Arraya tenant where the app is registered.

To assign the role in the customer tenant:

1. Sign in to the customer's Entra admin center.
2. Go to **Enterprise applications** and find the Arraya assessment app.
3. Under **Roles and administrators**, assign **Exchange Administrator** to the service principal.

Without this, `Connect-IPPSSession` will accept the certificate but the compliance cmdlets will not be exposed to the app session. Purview collection is skipped and retention/DLP data will be absent from the report.

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
