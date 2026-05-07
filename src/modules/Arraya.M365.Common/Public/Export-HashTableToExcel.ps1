## Export Hash Table to Excel
function Export-HashTableToExcel {
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True)] 
        [Hashtable]$hashtable,
        [Parameter(Mandatory=$True)] 
        [string]$ExportDetails,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Default', 'TenantToTenantCutover')]
        [string]$WorkbookExportPolicy = 'Default'
        #[Parameter(Mandatory=$false)] [switch]$tenant
    )

    # Ensure ExportDetails has .xlsx in the path name
    if ($ExportDetails -notmatch '\.xlsx$') {
        $ExportDetails += ".xlsx"
    }

    function Test-FileLocked {
        param([string]$Path)
        if (-not (Test-Path $Path)) { return $false }
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $stream.Close()
            return $false
        } catch {
            return $true
        }
    }

    # Central worksheet alias map to avoid Excel auto-truncation warnings.
    $worksheetAliases = @{
        'Office365GroupsActivityTopGroups' = 'O365GroupsActivityTop'
        'EmployeeExperienceInsightsSummary' = 'EmployeeExpInsights'
        'MailboxCalendarDelegatePermissions' = 'MailboxCalendarDelegatePerms'
        'PrivilegedAccessRemediationSummary' = 'PrivilegedAccessRemediation'
    }

    $worksheetCleanupAliases = @{
        'Office365GroupsActivityTopGroups' = @('Office365GroupsActivityTopGroups', 'Office365GroupsActivityTopGroup', 'O365GroupsActivityTop')
        'EmployeeExperienceInsightsSummary' = @('EmployeeExperienceInsightsSummary', 'EmployeeExperienceInsightsSumma', 'EmployeeExpInsights')
        'MailboxCalendarDelegatePermissions' = @('MailboxCalendarDelegatePermissions', 'MailboxCalendarDelegatePermissi', 'MailboxCalendarDelegatePerms')
        'PrivilegedAccessRemediationSummary' = @('PrivilegedAccessRemediationSummary', 'PrivilegedAccessRemediationSumm', 'PrivilegedAccessRemediation')
    }

    function Resolve-WorksheetExportName {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$LogicalName
        )

        if ($worksheetAliases.ContainsKey($LogicalName)) {
            return [string]$worksheetAliases[$LogicalName]
        }
        return $LogicalName
    }

    function Get-WorksheetCleanupCandidates {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$LogicalName
        )

        if ($worksheetCleanupAliases.ContainsKey($LogicalName)) {
            return @($worksheetCleanupAliases[$LogicalName] | Select-Object -Unique)
        }

        return @($LogicalName)
    }

    $optionalEmptySheets = @(
        'AuthenticationSSOApplications',
        'EnterpriseApplications',
        'SMTPRelaySummary',
        'TeamsVoiceSummary',
        'UnmanagedObjects',
        'OneDriveOwnerMismatches',
        'TeamsActivityTopUsers',
        'Office365GroupsActivityTopGroups',
        'EmployeeExperienceInsightsSummary',
        'InboxRulesExternalForwarding',
        'ExternalSharingSiteOverrides',
        'MailboxCalendarDelegatePermissions',
        'MfaEnforcementGapUsers',
        'MfaEnforcementScopeReview',
        'ConditionalAccessOptimization',
        'MfaMethodPostureSummary',
        'PrivilegedAccessRemediationSummary',
        'TeamsGroupsCleanupCandidates',
        'GroupLicensingSummary',
        'LicenseOptimizationCandidates'
    )

    $defaultExcludedWorksheets = @(
        'OwnershipGovernanceSummary',
        'TenantInfoSummary',
        'AuthenticationConfigSummary',
        'SpamFilteringSummary',
        'FederationSummary',
        'MfaRegistrationSummary',
        'EmailActivitySummary',
        'PrimaryMailboxStatsCollectionSummary',
        'UnifiedGroupMailboxStatsCollectionSummary',
        'PrimaryMailboxStatsCollectionSu',
        'UnifiedGroupMailboxStatsCollect'
    )

    function Wait-ForExportFileAvailability {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,
            [Parameter(Mandatory = $false)]
            [int]$MaxAttempts = 3
        )

        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            if (-not (Test-FileLocked -Path $Path)) {
                return
            }

            if ($attempt -lt $MaxAttempts) {
                Write-Log -Type WARNING -Message "Excel file is locked (attempt $attempt/$MaxAttempts). Close the file and retrying in 5 seconds..." -ExportFileLocation $ExportDetails
                Start-Sleep -Seconds 5
                continue
            }

            throw "Excel file '$Path' is locked and could not be opened for export."
        }
    }

    function Get-WorksheetExportSourceInfo {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            $TableValue
        )

        $sourceEnumerable = $null
        $sourceCount = 0
        $singleValue = $null

        if ($TableValue -is [hashtable] -or $TableValue -is [System.Collections.Specialized.OrderedDictionary]) {
            $sourceCount = $TableValue.Count
            $sourceEnumerable = $TableValue.Values
        }
        elseif ($TableValue -is [System.Collections.IEnumerable] -and -not ($TableValue -is [string])) {
            if ($TableValue.PSObject.Properties['Count']) {
                try { $sourceCount = [int]$TableValue.Count } catch { $sourceCount = -1 }
            }
            else {
                $sourceCount = -1
            }
            $sourceEnumerable = $TableValue
        }
        else {
            if ($null -ne $TableValue) {
                $singleValue = $TableValue
                $sourceCount = 1
            }
        }

        if ($sourceCount -lt 0 -and $null -ne $sourceEnumerable) {
            $sourceEnumerable = @($sourceEnumerable)
            $sourceCount = $sourceEnumerable.Count
        }

        return [pscustomobject]@{
            Count  = $sourceCount
            Source = if ($null -ne $singleValue) { @($singleValue) } else { $sourceEnumerable }
        }
    }

    function New-ExplicitWorksheetPlaceholderRow {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string[]]$Columns
        )

        $row = [ordered]@{}
        foreach ($column in $Columns) {
            $row[$column] = $null
        }

        return [pscustomobject]$row
    }

    function Get-TenantToTenantWorksheetColumns {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$LogicalName
        )

        $columnMap = @{
            'TenantInfo'                      = @('DisplayName', 'TenantId', 'InitialDomain', 'DefaultDomain', 'Country', 'CountryLetterCode', 'PreferredDataLocation', 'MultiGeoEnabled', 'MultiGeoAllowed', 'MultiGeoCentral', 'SelfServicePurchase', 'SelfServicePurchaseNotes', 'AzureResourceUsage', 'AzureResourceUsageNotes')
            'MigrationExecutiveSummary'       = @('Section', 'Metric', 'Value', 'Notes')
            'RecipientDomainSummary'          = @('Domain', 'RecipientCount', 'UserMailboxCount', 'SharedMailboxCount', 'GroupRecipientCount', 'HiddenFromAddressListsCount', 'Notes')
            'MailboxMigrationSummary'         = @('RecipientTypeDetails', 'MailboxCount', 'ActiveMailboxCount', 'InactiveMailboxCount', 'ArchiveEnabledCount', 'ForwardingCount', 'MailboxesWithDelegateDependencies', 'TotalDataToMigrateGB')
            'BitTitanLicenseSummary'          = @('Section', 'Metric', 'Value', 'Notes')
            'DelegateSummary'                 = @('PermissionType', 'AffectedMailboxCount', 'AssignmentCount', 'CollectionStates', 'Notes')
            'CollaborationSummary'            = @('Workload', 'TotalCount', 'TotalStorageGB', 'UnknownStorageCount', 'LargestObjectName', 'LargestObjectSizeGB', 'Notes')
            'CutoverPrepSummary'              = @('Category', 'Item', 'Status', 'Value', 'Notes')
            'AdConnectConfiguration'          = @('Section', 'Item', 'Value', 'Notes')
            'HybridConfiguration'             = @('Section', 'Item', 'Value', 'Notes')
            'Domains'                         = @('Name', 'IsDefault', 'IsInitial', 'Verified', 'AuthenticationType', 'SupportedServices', 'MXRecords', 'Office365MailExchanger', 'ThirdPartySpamFilterReview', 'HybridRoutingReview', 'RecipientCount', 'Notes')
            'Users'                           = @('DisplayName', 'UserPrincipalName', 'Mail', 'AccountEnabled', 'OnPremisesSyncEnabled', 'AssignedLicensesFriendly', 'HasMailbox', 'MailboxPrimarySmtpAddress', 'RecipientTypeDetails', 'TargetUPN', 'TargetPrimarySmtpAddress', 'Wave', 'MigrationState', 'CutoverDate', 'Notes')
            'AllRecipients'                   = @('DisplayName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'UserPrincipalName', 'WindowsEmailAddress', 'PrimaryDomain', 'EmailAddresses', 'OnMicrosoftAlias', 'OnMicrosoftAliases', 'HiddenFromAddressListsEnabled', 'Notes')
            'AllMailboxes'                    = @('DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'IsInactiveMailbox', 'OnPremisesSyncEnabled', 'ExchangeGuid', 'ArchiveGuid', 'OnMicrosoftAlias', 'OnMicrosoftAliases', 'LegacyExchangeDn', 'LegacyExchangeDnX500', 'X500Addresses', 'X400Addresses', 'MailboxSizeGB', 'DeletedItemsGB', 'ArchiveStatus', 'ArchiveSizeGB', 'ArchiveDeletedItemsGB', 'TotalDataToMigrateGB', 'BitTitanLicenseType', 'BitTitanLicenseCount', 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward', 'LitigationHoldEnabled', 'RetentionPolicy', 'FullAccessDelegateCount', 'SendAsDelegateCount', 'GrantSendOnBehalfToCount', 'CalendarDelegateCount', 'Wave', 'MigrationState', 'CutoverDate', 'Notes')
            'MailboxDelegateAssignments'      = @('MailboxDisplayName', 'MailboxPrimarySmtpAddress', 'MailboxUserPrincipalName', 'RecipientTypeDetails', 'PermissionType', 'Delegate', 'DelegateCountSource', 'CollectionState', 'Wave', 'Notes')
            'MailboxCalendarDelegatePermissions' = @('MailboxDisplayName', 'MailboxPrimarySmtpAddress', 'MailboxUserPrincipalName', 'RecipientTypeDetails', 'CalendarName', 'CalendarPath', 'PermissionTarget', 'PermissionTargetDisplayName', 'PermissionTargetType', 'AccessRights', 'SharingPermissionFlags', 'Wave', 'Notes')
            'InactiveMailboxDetails'          = @('DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'IsInactiveMailbox', 'OnPremisesSyncEnabled', 'ExchangeGuid', 'ArchiveGuid', 'OnMicrosoftAlias', 'OnMicrosoftAliases', 'LegacyExchangeDn', 'LegacyExchangeDnX500', 'X500Addresses', 'X400Addresses', 'MailboxSizeGB', 'DeletedItemsGB', 'ArchiveStatus', 'ArchiveSizeGB', 'ArchiveDeletedItemsGB', 'TotalDataToMigrateGB', 'BitTitanLicenseType', 'BitTitanLicenseCount', 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward', 'LitigationHoldEnabled', 'RetentionPolicy', 'FullAccessDelegateCount', 'SendAsDelegateCount', 'GrantSendOnBehalfToCount', 'CalendarDelegateCount', 'Wave', 'MigrationState', 'CutoverDate', 'Notes')
            'PublicFolderDetails'             = @('Name', 'Identity', 'Path', 'MailEnabled', 'PrimarySmtpAddress', 'ItemCount', 'FolderCount', 'EntryId', 'Notes')
            'MailFlowConnectors'              = @('Name', 'Enabled', 'ConnectorType', 'ConnectorSource', 'SenderDomains', 'RecipientDomains', 'SmartHosts', 'TlsSettings', 'Comment', 'Notes')
            'RemoteDomains'                   = @('Name', 'DomainName', 'AutoForwardEnabled', 'AllowedOOFType', 'TNEFEnabled', 'TrustedMailOutboundEnabled', 'Notes')
            'SMTPRelayServiceAccounts'        = @('DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'IsDirSynced', 'AssignedLicensesFriendly', 'Notes')
            'AllTeams'                        = @('DisplayName', 'Visibility', 'IsArchived', 'SharePointSiteUrl', 'SiteSize-GB', 'TotalChannels', 'SharedChannelCount', 'SharedChannels', 'OwnerCount', 'MemberCount', 'GuestCount', 'LastActivityDate', 'Notes')
            'SharePoint'                      = @('Title', 'Url', 'Template', 'Owner', 'StorageUsedGB', 'StorageQuota', 'LastContentModifiedDate', 'LockState', 'ArchiveStatus', 'SharingCapability', 'IsTeamsConnected', 'Notes')
            'OneDrive'                        = @('Title', 'Url', 'Template', 'Owner', 'StorageUsedGB', 'StorageQuota', 'LastContentModifiedDate', 'LockState', 'ArchiveStatus', 'SharingCapability', 'IsTeamsConnected', 'Notes')
        }

        if ($columnMap.ContainsKey($LogicalName)) {
            return @($columnMap[$LogicalName])
        }

        return @()
    }

    function Get-TenantToTenantWorksheetSourceName {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$LogicalName,
            [Parameter(Mandatory = $true)]
            [hashtable]$ExportTables
        )

        switch ($LogicalName) {
            'TenantInfo' {
                if (
                    $ExportTables.ContainsKey('TenantInfoSummary') -and
                    (Get-WorksheetExportSourceInfo -TableValue $ExportTables['TenantInfoSummary']).Count -gt 0
                ) {
                    return 'TenantInfoSummary'
                }
            }
            'AdConnectConfiguration' {
                return 'TenantToTenantAdConnectConfiguration'
            }
            'HybridConfiguration' {
                return 'TenantToTenantHybridConfiguration'
            }
            'AllMailboxes' {
                if (
                    $ExportTables.ContainsKey('MailboxFullDetails') -and
                    (Get-WorksheetExportSourceInfo -TableValue $ExportTables['MailboxFullDetails']).Count -gt 0
                ) {
                    return 'MailboxFullDetails'
                }
            }
        }

        return $LogicalName
    }

    function Get-TenantToTenantColumnWidth {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$ColumnName
        )

        if ($ColumnName -match '^(MailboxSizeGB|DeletedItemsGB|ArchiveSizeGB|ArchiveDeletedItemsGB|TotalDataToMigrateGB|BitTitanLicenseCount|RecipientCount|UserMailboxCount|SharedMailboxCount|GroupRecipientCount|HiddenFromAddressListsCount|MailboxCount|ActiveMailboxCount|InactiveMailboxCount|ArchiveEnabledCount|ForwardingCount|MailboxesWithDelegateDependencies|AffectedMailboxCount|AssignmentCount|TotalCount|TotalStorageGB|UnknownStorageCount|LargestObjectSizeGB|OwnerCount|MemberCount|GuestCount|TotalChannels|SharedChannelCount|Count|Status|AccountEnabled|HasMailbox|IsInactiveMailbox|OnPremisesSyncEnabled|IsArchived|Verified|IsDefault|IsInitial|Office365MailExchanger|LitigationHoldEnabled|DeliverToMailboxAndForward)$') {
            return 14
        }

        if ($ColumnName -match '^(DisplayName|UserPrincipalName|PrimarySmtpAddress|WindowsEmailAddress|MailboxPrimarySmtpAddress|MailboxUserPrincipalName|Mail|Owner|Title|Url|SharePointSiteUrl|Source|Target|TargetUPN|TargetPrimarySmtpAddress|TargetMail|SourceUPN|SourceEmail|SourceOneDriveUrl|TargetOneDriveUrl|LargestObjectName|ForwardingAddress|ForwardingSmtpAddress)$') {
            return 26
        }

        if ($ColumnName -match '^(EmailAddresses|OnMicrosoftAliases|LegacyExchangeDn|LegacyExchangeDnX500|X500Addresses|X400Addresses|Delegate|FullAccessDelegates|SendAsDelegates|GrantSendOnBehalfTo|SmartHosts|SenderDomains|RecipientDomains|Comment|Notes|CollectionStates|SupportedServices|TlsSettings|MXRecords|SharedChannels)$') {
            return 34
        }

        return 20
    }

    function Set-TenantToTenantWorksheetColumnWidths {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            $ExcelPackage,
            [Parameter(Mandatory = $true)]
            [string]$WorksheetName,
            [Parameter(Mandatory = $true)]
            [string[]]$Columns
        )

        $worksheet = $ExcelPackage.Workbook.Worksheets[$WorksheetName]
        if ($null -eq $worksheet) {
            return
        }

        for ($columnIndex = 1; $columnIndex -le $Columns.Count; $columnIndex++) {
            $worksheet.Column($columnIndex).Width = Get-TenantToTenantColumnWidth -ColumnName $Columns[$columnIndex - 1]
            $worksheet.Column($columnIndex).BestFit = $false
        }
    }
    
    # === Sheet ordering ===
    $defaultDesiredOrder = @(
        # Tenant Overview
        "TenantInfo", "LicenseSKUs", "GroupLicensingSummary", "LicenseOptimizationCandidates", "AdConnectConfiguration",

        # Identity
        "Users", "UserFullDetails", "Admins", "EntraIDGroups", "Domains",
        "AuthenticationMethods", "AuthenticationSSOApplications", "EnterpriseApplications", "AuthenticationConfig",
        "MfaEnrollmentSummary", "MfaMethodPostureSummary", "MfaEnforcementSummary", "MfaEnforcementGapUsers", "MfaEnforcementScopeReview",
        "ConditionalAccessPolicySummary", "ConditionalAccessOptimization", "ConditionalAccessPolicies", "SecurityDefaultsPolicy", "GuestSignInSummary",
        "GuestAccessConfiguration", "ExternalIdentityRestrictions", "PrivilegedAccessSummary", "PrivilegedAccessRemediationSummary",

        # Exchange
        "HybridConfiguration",
        "AllRecipients", "AllMailboxes", "PrimaryMailboxStats", "MailboxFullDetails", "MailboxCalendarDelegatePermissions", "NonUserMailboxes",
        "ArchiveMailboxes", "ArchiveMailboxStats", "LitigationHoldMailboxes", "InactiveMailboxes", "InactiveMailboxDetails",
        "AllExchangeGroups", "PublicFolderDetails", "PublicFolderPerms",
        "MailFlowRules", "MailFlowConnectors", "RemoteDomains", "EmailActivityTopSenders", "EmailActivityTopReceivers",
        "SMTPRelayConfig", "SMTPRelayServiceAccounts", "SpamFilteringConfig", "SharedMailboxGovernanceSummary",
        "ForwardingPolicySummary", "InboxRuleForwardingSummary", "InboxRulesExternalForwarding",

        # Collaboration
        "UnifiedGroups", "AllTeams", "TeamsGroupsCleanupCandidates", "TeamsVoiceSummary", "SharePoint", "OneDrive",
        "SharePointSharingSummary", "TeamsActivityTopUsers", "Office365GroupsActivityTopGroups",
        "EmployeeExperienceInsightsSummary", "CollaborationActivitySummary",

        # Endpoint
        "DeviceDetails", "DeviceManagementSummary",

        # Governance
        "SecuritySecureScore", "SecureScoreActions", "SMTPRelaySummary", "RetentionPolicies", "DlpPolicies",
        "PasswordLifecycleSummary", "ExternalSharingSummary", "ExternalSharingSiteOverrides", "ExternalExposureFindings",
        "UnmanagedObjects", "OneDriveOwnerMismatches",

        # Assessment Outputs
        "BestPractices", "BestPracticeFindings", "MigrationReadiness"
    )

    $tenantToTenantCutoverDesiredOrder = @(
        'AllMailboxes',
        'TenantInfo',
        'MigrationExecutiveSummary',
        'RecipientDomainSummary',
        'MailboxMigrationSummary',
        'BitTitanLicenseSummary',
        'DelegateSummary',
        'CollaborationSummary',
        'CutoverPrepSummary',
        'Domains',
        'AdConnectConfiguration',
        'HybridConfiguration',
        'Users',
        'AllRecipients',
        'MailboxDelegateAssignments',
        'MailboxCalendarDelegatePermissions',
        'InactiveMailboxDetails',
        'PublicFolderDetails',
        'MailFlowConnectors',
        'RemoteDomains',
        'SMTPRelayServiceAccounts',
        'AllTeams',
        'SharePoint',
        'OneDrive'
    )
    $tenantToTenantCutoverExcludedWorksheets = @(
        'Admins', 'AuthenticationConfig', 'AuthenticationSSOApplications', 'MfaEnrollmentSummary', 'MfaEnforcementSummary',
        'ConditionalAccessPolicySummary', 'ConditionalAccessPolicies', 'SecurityDefaultsPolicy', 'GuestAccessConfiguration',
        'ExternalIdentityRestrictions', 'MailboxFullDetails', 'ArchiveMailboxes', 'ArchiveMailboxStats', 'UnifiedGroups',
        'SpamFilteringConfig', 'SMTPRelayConfig', 'SharedMailboxGovernanceSummary', 'ForwardingPolicySummary', 'BestPractices',
        'BestPracticeFindings', 'UserFullDetails', 'EntraIDGroups', 'EnterpriseApplications', 'SecuritySecureScore',
        'SecureScoreActions', 'EmailActivityTopSenders', 'EmailActivityTopReceivers', 'TeamsVoiceSummary',
        'CollaborationActivitySummary', 'TeamsActivityTopUsers', 'Office365GroupsActivityTopGroups',
        'EmployeeExperienceInsightsSummary', 'DeviceDetails', 'DeviceManagementSummary', 'LitigationHoldMailboxes',
        'NonUserMailboxes', 'InactiveMailboxes', 'PublicFolderPerms', 'RetentionPolicies', 'DlpPolicies',
        'UnmanagedObjects', 'OneDriveOwnerMismatches', 'ExternalSharingSummary', 'ExternalSharingSiteOverrides',
        'SharePointSharingSummary', 'AllExchangeGroups', 'MigrationReadiness', 'ConditionalAccessOptimization',
        'MfaMethodPostureSummary', 'PrivilegedAccessRemediationSummary', 'TeamsGroupsCleanupCandidates',
        'GroupLicensingSummary', 'LicenseOptimizationCandidates'
    )

    $desiredOrder = $defaultDesiredOrder
    $excludedWorksheets = $defaultExcludedWorksheets
    $appendRemainingWorksheets = $true
    switch ($WorkbookExportPolicy) {
        'TenantToTenantCutover' {
            $desiredOrder = $tenantToTenantCutoverDesiredOrder
            $excludedWorksheets = @($defaultExcludedWorksheets + $tenantToTenantCutoverExcludedWorksheets | Select-Object -Unique)
            $appendRemainingWorksheets = $false
        }
    }

    $orderedTables = @()
    foreach ($name in $desiredOrder) {
        if (($excludedWorksheets -notcontains $name) -and $hashtable.ContainsKey($name)) {
            $orderedTables += $name
        }
    }
    if ($appendRemainingWorksheets) {
        $orderedTables += ($hashtable.Keys | Where-Object { ($orderedTables -notcontains $_) -and ($excludedWorksheets -notcontains $_) } | Sort-Object)
    }

    foreach ($excludedSheet in $excludedWorksheets) {
        if ($hashtable.ContainsKey($excludedSheet)) {
            Write-Log -Type INFO -Message "Skipping worksheet '$excludedSheet' by export policy '$WorkbookExportPolicy'." -ExportFileLocation $ExportDetails
        }
    }

    $totalCount = ($orderedTables | Measure-Object).Count
    $excelPackage = $null
    $autoSizeRowLimit = 1000

    try {
        Wait-ForExportFileAvailability -Path $ExportDetails
        if (Test-Path -Path $ExportDetails) {
            Remove-Item -LiteralPath $ExportDetails -Force -ErrorAction Stop
            Write-Log -Type INFO -Message "Removed existing workbook prior to export so the workbook can be rebuilt in a single optimized pass." -ExportFileLocation $ExportDetails
        }

        $excelPackage = Open-ExcelPackage -Path $ExportDetails -Create

        foreach ($table in $orderedTables) {
            try {
                $worksheetName = Resolve-WorksheetExportName -LogicalName $table
                Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash To Excel"
                Write-Log -Type DEBUG -Message ("Exporting '{0}' Hash Table to '{1}' as worksheet '{2}'" -f $table, $ExportDetails, $worksheetName) -ExportFileLocation $ExportDetails

                $tableValue = if ($WorkbookExportPolicy -eq 'TenantToTenantCutover') {
                    $sourceTableName = Get-TenantToTenantWorksheetSourceName -LogicalName $table -ExportTables $hashtable
                    if ($hashtable.ContainsKey($sourceTableName)) { $hashtable[$sourceTableName] } else { $null }
                }
                else {
                    $hashtable[$table]
                }
                $sourceInfo = Get-WorksheetExportSourceInfo -TableValue $tableValue
                $sourceCount = [int]$sourceInfo.Count
                $exportSource = $sourceInfo.Source
                $explicitColumns = if ($WorkbookExportPolicy -eq 'TenantToTenantCutover') { @(Get-TenantToTenantWorksheetColumns -LogicalName $table) } else { @() }

                if ($sourceCount -gt 0) {
                    $autoSizeSheet = ($sourceCount -le $autoSizeRowLimit -and $WorkbookExportPolicy -ne 'TenantToTenantCutover')
                    if (-not $autoSizeSheet) {
                        Write-Log -Type INFO -Message "Skipping AutoSize for worksheet '$worksheetName' due to row count ($sourceCount) to reduce export runtime/memory pressure." -ExportFileLocation $ExportDetails
                    }

                    $excelSplat = @{
                        ExcelPackage = $excelPackage
                        WorksheetName = $worksheetName
                        ClearSheet    = $true
                        BoldTopRow    = $true
                    }
                    if ($autoSizeSheet) {
                        $excelSplat.AutoSize = $true
                    }

                    $rowsToExport = if ($explicitColumns.Count -gt 0) {
                        @($exportSource | Select-Object $explicitColumns)
                    }
                    else {
                        @($exportSource)
                    }

                    $rowsToExport |
                        ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                        Export-Excel @excelSplat -PassThru |
                        ForEach-Object { $excelPackage = $_ }

                    if ($WorkbookExportPolicy -eq 'TenantToTenantCutover' -and $explicitColumns.Count -gt 0) {
                        Set-TenantToTenantWorksheetColumnWidths -ExcelPackage $excelPackage -WorksheetName $worksheetName -Columns $explicitColumns
                    }
                    continue
                }

                if ($WorkbookExportPolicy -eq 'TenantToTenantCutover' -and $explicitColumns.Count -gt 0) {
                    @(New-ExplicitWorksheetPlaceholderRow -Columns $explicitColumns) |
                        ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                        Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheetName -ClearSheet -BoldTopRow -PassThru |
                        ForEach-Object { $excelPackage = $_ }
                    Set-TenantToTenantWorksheetColumnWidths -ExcelPackage $excelPackage -WorksheetName $worksheetName -Columns $explicitColumns
                    Write-Log -Type INFO -Message "No data found for '$table'; exported an empty-schema worksheet for the tenant-to-tenant workbook contract." -ExportFileLocation $ExportDetails
                    continue
                }

                if ($table -eq 'OneDriveOwnerMismatches') {
                    $placeholderRows = @(
                        [PSCustomObject]@{
                            DisplayName          = 'N/A'
                            SiteUrl              = 'N/A'
                            CurrentOwner         = 'N/A'
                            ExpectedDefaultOwner = 'N/A'
                            Notes                = 'No owner mismatches detected in this run.'
                        }
                    )

                    $placeholderRows |
                        ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                        Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheetName -ClearSheet -BoldTopRow -AutoSize -PassThru |
                        ForEach-Object { $excelPackage = $_ }
                    Write-Log -Type INFO -Message "No owner mismatch rows found; exported placeholder row to worksheet '$worksheetName'." -ExportFileLocation $ExportDetails
                    continue
                }

                $logType = if ($optionalEmptySheets -contains $table) { 'INFO' } else { 'WARNING' }
                Write-Log -Type $logType -Message "No data found for $table to export to Excel" -ExportFileLocation $ExportDetails
            }
            catch {
                Write-Log -Type Error -Message "An error occurred in Exporting Hash To Excel for $($table) to '$($ExportDetails)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
        }
    }
    finally {
        Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash To Excel" -Completed
        if ($null -ne $excelPackage) {
            try {
                Close-ExcelPackage -ExcelPackage $excelPackage
            }
            catch {
                try {
                    $excelPackage.Save()
                    $excelPackage.Dispose()
                }
                catch {
                    Write-Log -Type WARNING -Message "Unable to close Excel package cleanly: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                }
            }
        }
    }

    Write-Host "The report has been exported to: $($ExportDetails)" -ForegroundColor Green
}
