# Office365Custom Module

## Overview
The **Office365Custom** module centralises scripts for connecting to Microsoft 365 services, running Microsoft Graph queries, and handling data import/export tasks used by the repository. Version `1.2.1` is the current module release in this folder and follows a Public/Private layout so helper routines remain internal while only the supported entry points are exported.

## Version
- Module version: `1.2.1`
- This release reflects the updated manifest metadata and current exported command set for the 1.2.1 module folder.

## Exported Commands
- `Capture-ErrorHelper` – Build a structured error object and optionally track discovery errors for later review.
- `Connect-Office365` – Connect to Microsoft Graph first, then establish Exchange Online, SharePoint Online, and Teams connectivity for the tenant.
- `Convert-MgObjectToUri` – Compose Microsoft Graph request URIs from object metadata, filters, and select clauses.
- `Export-MigratingDomainRecipients` - Exports recipients with email addresses matching a specified migrating domain to Excel for analysis.
- `Get-AllOneDriveURLs` – Retrieve all OneDrive for Business URLs across the tenant, including duplicate detection.
- `Get-ExportPath` – Prompt for an export location and return a fully qualified path with sensible defaults.
- `Get-GraphAPIActivityReport` – Download Graph reporting CSVs for services such as SharePoint, Teams, and Exchange.
- `Get-GraphData` – Execute Graph REST queries with paging, retry handling, and token refresh support.
- `Handle-ErrorHelper` – Produce a friendly error message payload for host output.
- `Load-HashTableFromJson` – Read configuration data persisted as JSON back into a cleaned hashtable.
- `Manage-DomainAlias` - Handles adding and removing domain aliases from recipient objects during migrations.
- `Manage-UserBatches` - Processes user operations in configurable batches to control migration velocity.
- `Save-HashTableToJson` – Persist a hashtable to disk while removing empty keys or values.
- `Select-FromMultipleO365Matches` - Helps resolve ambiguous recipient matches by presenting options for manual selection.
- `Select-ImportedHeader` – Display available column headers from imported data and prompt for a selection.
- `Set-PrimarySMTPAddress` - Updates primary SMTP addresses for recipients during domain migrations.
- `Start-T2TDomainCutoverMigration` - Orchestrates domain cutover migrations between tenants by managing UPN updates, SMTP address changes, and domain alias operations.
- `Update-UserUPN` - Changes user principal names (UPNs) using Microsoft Graph during tenant migrations.
- `Validate-ImportedData` – Accept an array, CSV path, or Excel path and return the parsed objects.
- `Write-Log` – Emit structured log entries to the console or a file location.
- `Write-MigrationLog` - Records migration actions and changes to a log file for tracking and auditing purposes.
- `Write-ProgressHelper` – Display a progress bar with elapsed and estimated time remaining information.

## Prerequisites
- Microsoft.Graph PowerShell SDK (or Microsoft.Graph.* modules on Windows PowerShell).
- ExchangeOnlineManagement module.
- MicrosoftTeams module.
- Microsoft.Online.SharePoint.PowerShell module.
- ImportExcel module for Excel import/export features.

Helper functions under `Private/` handle retry logic, module imports, and service-specific connectivity and are intentionally not exported.
