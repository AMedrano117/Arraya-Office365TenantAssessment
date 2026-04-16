## Export Hash Table to Excel
function Export-HashTableToExcel {
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True)] 
        [Hashtable]$hashtable,
        [Parameter(Mandatory=$True)] 
        [string]$ExportDetails
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
    }

    $worksheetCleanupAliases = @{
        'Office365GroupsActivityTopGroups' = @('Office365GroupsActivityTopGroups', 'Office365GroupsActivityTopGroup', 'O365GroupsActivityTop')
        'EmployeeExperienceInsightsSummary' = @('EmployeeExperienceInsightsSummary', 'EmployeeExperienceInsightsSumma', 'EmployeeExpInsights')
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
        'MfaEnforcementGapUsers',
        'MfaEnforcementScopeReview'
    )

    $excludedWorksheets = @(
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
    
    # === Sheet ordering ===
    $desiredOrder = @(
        # Tenant Overview
        "TenantInfo", "LicenseSKUs", "AdConnectConfiguration",

        # Identity
        "Users", "UserFullDetails", "Admins", "EntraIDGroups", "Domains",
        "AuthenticationMethods", "AuthenticationSSOApplications", "EnterpriseApplications", "AuthenticationConfig",
        "MfaEnrollmentSummary", "MfaEnforcementSummary", "MfaEnforcementGapUsers", "MfaEnforcementScopeReview",
        "ConditionalAccessPolicySummary", "ConditionalAccessPolicies", "SecurityDefaultsPolicy", "GuestSignInSummary",
        "GuestAccessConfiguration", "ExternalIdentityRestrictions", "PrivilegedAccessSummary",

        # Exchange
        "HybridConfiguration",
        "AllRecipients", "AllMailboxes", "PrimaryMailboxStats", "MailboxFullDetails", "NonUserMailboxes",
        "ArchiveMailboxes", "ArchiveMailboxStats", "LitigationHoldMailboxes", "InactiveMailboxes", "InactiveMailboxDetails",
        "AllExchangeGroups", "PublicFolderDetails", "PublicFolderPerms",
        "MailFlowRules", "MailFlowConnectors", "RemoteDomains", "EmailActivityTopSenders", "EmailActivityTopReceivers",
        "SMTPRelayConfig", "SMTPRelayServiceAccounts", "SpamFilteringConfig", "SharedMailboxGovernanceSummary",
        "ForwardingPolicySummary", "InboxRuleForwardingSummary", "InboxRulesExternalForwarding",

        # Collaboration
        "UnifiedGroups", "AllTeams", "TeamsVoiceSummary", "SharePoint", "OneDrive",
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

    $orderedTables = @()
    foreach ($name in $desiredOrder) {
        if (($excludedWorksheets -notcontains $name) -and $hashtable.ContainsKey($name)) {
            $orderedTables += $name
        }
    }
    $orderedTables += ($hashtable.Keys | Where-Object { ($orderedTables -notcontains $_) -and ($excludedWorksheets -notcontains $_) } | Sort-Object)

    foreach ($excludedSheet in $excludedWorksheets) {
        if ($hashtable.ContainsKey($excludedSheet)) {
            Write-Log -Type INFO -Message "Skipping worksheet '$excludedSheet' by export policy." -ExportFileLocation $ExportDetails
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

                $sourceInfo = Get-WorksheetExportSourceInfo -TableValue $hashtable[$table]
                $sourceCount = [int]$sourceInfo.Count
                $exportSource = $sourceInfo.Source

                if ($sourceCount -gt 0) {
                    $autoSizeSheet = ($sourceCount -le $autoSizeRowLimit)
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

                    $exportSource |
                        ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                        Export-Excel @excelSplat -PassThru |
                        ForEach-Object { $excelPackage = $_ }
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
