# Office365Custom Module

## Overview
The **Office365Custom** module centralises scripts for connecting to Microsoft 365 services, running Microsoft Graph queries, and handling data import/export tasks used by the repository. The module now follows a Public/Private layout so helper routines remain internal while only the supported entry points are exported.

## Exported Commands
- `Capture-ErrorHelper` – Build a structured error object and optionally track discovery errors for later review.
- `Connect-MicrosoftGraph` – Guide the operator through connecting with the Microsoft Graph SDK or REST workflow.
- `Connect-MicrosoftGraphAPI` – Establish an application-based connection to Microsoft Graph using client credentials.
- `Connect-Office365Services` – Interactive wrapper for connecting to Exchange Online, SharePoint Online, Microsoft Graph, and other workloads.
- `Convert-MgObjectToUri` – Compose Microsoft Graph request URIs from object metadata, filters, and select clauses.
- `Get-AllOneDriveURLs` – Retrieve all OneDrive for Business URLs across the tenant, including duplicate detection.
- `Get-ExportPath` – Prompt for an export location and return a fully qualified path with sensible defaults.
- `Get-GraphAPIActivityReport` – Download Graph reporting CSVs for services such as SharePoint, Teams, and Exchange.
- `Get-GraphData` – Execute Graph REST queries with paging, retry handling, and token refresh support.
- `Handle-ErrorHelper` – Produce a friendly error message payload for host output.
- `Load-HashTableFromJson` – Read configuration data persisted as JSON back into a cleaned hashtable.
- `Save-HashTableToJson` – Persist a hashtable to disk while removing empty keys or values.
- `Select-ImportedHeader` – Display available column headers from imported data and prompt for a selection.
- `Validate-ImportedData` – Accept an array, CSV path, or Excel path and return the parsed objects.
- `Write-Log` – Emit structured log entries to the console or a file location.
- `Write-ProgressHelper` – Display a progress bar with elapsed and estimated time remaining information.
- `Start-T2TDomainCutoverMigration` - Orchestrates domain cutover migrations between tenants by managing UPN updates, SMTP address changes, and domain alias operations.
- `Write-MigrationLog` - Records migration actions and changes to a log file for tracking and auditing purposes.
- `Export-MigratingDomainRecipients` - Exports recipients with email addresses matching a specified migrating domain to Excel for analysis.
- `Set-PrimarySMTPAddress` - Updates primary SMTP addresses for recipients during domain migrations.
- `Update-UserUPN` - Changes user principal names (UPNs) using Microsoft Graph during tenant migrations.
- `Select-FromMultipleO365Matches` - Helps resolve ambiguous recipient matches by presenting options for manual selection.
- `Manage-DomainAlias` - Handles adding and removing domain aliases from recipient objects during migrations.
- `Manage-UserBatches` - Processes user operations in configurable batches to control migration velocity.

## Prerequisites
- Microsoft.Graph PowerShell SDK (or Microsoft.Graph.* modules on Windows PowerShell).
- ExchangeOnlineManagement module.
- MicrosoftTeams module.
- Microsoft.Online.SharePoint.PowerShell module.
- ImportExcel module for Excel import/export features.

Helper functions under `Private/` handle retry logic, module imports, and service-specific connectivity and are intentionally not exported.
