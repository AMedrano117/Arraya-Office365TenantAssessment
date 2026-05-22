# App Registration Permissions

Required Microsoft Graph **Application** permissions for the Tenant Assessment tool.
`Directory.Read.All` has been decomposed into five least-privilege alternatives.

## Core Permissions (always required)

| Permission | Replaces | Used For |
| --- | --- | --- |
| `User.Read.All` | — | List users, user properties |
| `AuditLog.Read.All` | — | `signInActivity` on users (Entra P1/P2 required on tenant) |
| `Organization.Read.All` | `Directory.Read.All` | `GET /organization`, `GET /subscribedSkus` |
| `Domain.Read.All` | `Directory.Read.All` | `GET /domains`, `GET /domains/{id}/federationConfiguration` |
| `Group.Read.All` | `Directory.Read.All` | `GET /groups` (M365, security, and distribution groups) |
| `Device.Read.All` | `Directory.Read.All` | `GET /devices` |
| `RoleManagement.Read.Directory` | `Directory.Read.All` | `GET /directoryRoles`, role member enumeration |
| `Policy.Read.All` | — | CA policies, named locations, cross-tenant access, auth methods policy, security defaults, authorization policy, auth strength policies |
| `Application.Read.All` | — | `GET /servicePrincipals` (enterprise apps) |
| `Reports.Read.All` | — | All `GET /reports/*` usage and activity endpoints, MFA registration details |
| `Team.ReadBasic.All` | — | `GET /teams` (Teams inventory) |
| `TeamMember.Read.All` | — | Team owners, member count, guest members per team |
| `Channel.ReadBasic.All` | — | `GET /teams/{id}/channels` (private channel detection) |
| `SharePointTenantSettings.Read.All` | — | `GET /admin/sharepoint/settings` (v1.0) |
| `SecurityEvents.Read.All` | — | Secure Score and Secure Score control profiles |

## Optional Permissions (graceful degradation if absent)

| Permission | Used For | Requirement |
| --- | --- | --- |
| `IdentityRiskyUser.Read.All` | `GET /identityProtection/riskyUsers` — risky user list | Entra ID P2 license on target tenant; collector returns empty list on 403 |
| `MailboxSettings.Read` | Shared/room/equipment mailbox detection (`userPurpose`), inbox rule external forwarding scan | Exchange Online mailbox required per user; adds per-user API calls |
| `Place.Read.All` | `GET /places/microsoft.graph.room` — room mailbox inventory | Exchange Online; cleaner alternative to userPurpose for rooms only |
| `TeamsAppInstallation.Read.All` | `GET /teams/{id}/installedApps` — detect sideloaded/third-party apps | Adds per-team API calls |

## What Directory.Read.All covered vs. least-privilege replacements

`Directory.Read.All` is a single broad permission that grants read access to the entire directory. The five narrower permissions above cover exactly the same endpoints this tool calls, with no excess:

| Endpoint | Old (broad) | New (specific) |
| --- | --- | --- |
| `GET /organization` | `Directory.Read.All` | `Organization.Read.All` |
| `GET /subscribedSkus` | `Directory.Read.All` | `Organization.Read.All` |
| `GET /domains` | `Directory.Read.All` | `Domain.Read.All` |
| `GET /domains/{id}/federationConfiguration` | `Directory.Read.All` | `Domain.Read.All` |
| `GET /groups` (all types) | `Directory.Read.All` | `Group.Read.All` |
| `GET /devices` | `Directory.Read.All` | `Device.Read.All` |
| `GET /directoryRoles` | `Directory.Read.All` | `RoleManagement.Read.Directory` |
| `GET /directoryRoles/{id}/members` | `Directory.Read.All` | `RoleManagement.Read.Directory` |

## Endpoints still requiring EXO PowerShell (no Graph equivalent)

These signals cannot be collected via Graph API regardless of permissions:

| Signal | EXO Cmdlet | Reason |
| --- | --- | --- |
| Mailbox SMTP forwarding | `Get-Mailbox \| Select ForwardingSmtpAddress` | Not in Graph mailboxSettings; Exchange-backend only |
| Mail flow / transport rules | `Get-TransportRule` | No Graph REST endpoint exists |
| Anti-spam policies | `Get-HostedContentFilterPolicy` | Only via Exchange.ManageAsApp DSC framework, not REST GET |
| Litigation hold status | `Get-Mailbox \| Select LitigationHoldEnabled` | Not in Graph mailboxSettings |
| Inactive mailboxes | `Get-Mailbox -InactiveMailboxOnly` | Not surfaced in Graph at all |
| Per-site external sharing level | `Get-SPOSite \| Select SharingCapability` | SPO Admin API only; Graph site permissions endpoint returns app grants only |

## Notes

- `AuditLog.Read.All` requires Entra ID P1 or P2 on the target tenant for `signInActivity` to be populated. Without it, sign-in timestamps will be null.
- `IdentityRiskyUser.Read.All` requires **both** the permission grant **and** an Entra ID P2 license. The collector returns an empty list on 403 without failing.
- `MailboxSettings.Read` triggers per-user calls. For large tenants this is batched (20 per request). It is optional — the assessment runs cleanly without it, but shared mailbox and inbox rule data will not be collected.
- `Group.Read.All` covers M365 Unified groups, security groups, and traditional mail-enabled distribution lists (using advanced query filter with `ConsistencyLevel: eventual`).

## Auth Modes

| Mode | Description | When to Use |
| --- | --- | --- |
| `interactive` | Device code flow (browser prompt) | Ad-hoc runs, demos |
| `certificate` | Certificate-based client credentials | Automated / production runs |
| `client-secret` | Client secret credentials | Dev/test environments |
