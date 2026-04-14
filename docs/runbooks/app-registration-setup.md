# App Registration Setup

Use this runbook when you want to run the assessment with `-AuthMode Certificate` or `-AuthMode ClientSecret`.

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
- Exchange Online app-only access must be configured for certificate-based Exchange collection.
- SharePoint and OneDrive standard collection use Microsoft Graph in the current app-based workflow.
- Teams PowerShell remains reduced in app-based auth modes, so treat Teams detail as best-effort when running non-interactively.

## Recommended operator/admin roles for setup
- `Application Administrator`
- `Cloud Application Administrator`
- or `Global Administrator`

## Keep secrets and certificates out of the repo
- Do not store client secrets, PFX files, or exported private keys in this repository.
- Store certificate material in approved secure storage and import it on the execution machine only.

## Next step
After the app registration exists, continue with [certificate-auth-setup.md](certificate-auth-setup.md) if you plan to use certificate auth.
