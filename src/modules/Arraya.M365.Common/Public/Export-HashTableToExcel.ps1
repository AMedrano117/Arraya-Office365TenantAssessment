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

    if (Test-Path -Path $ExportDetails) {
        try {
            $existingSheetNames = @()
            if (Get-Command -Name Get-ExcelSheetInfo -ErrorAction SilentlyContinue) {
                $existingSheetNames = @(Get-ExcelSheetInfo -Path $ExportDetails | Select-Object -ExpandProperty Name)
            }

            $sheetsToRemove = @($excludedWorksheets | Where-Object { $existingSheetNames -contains $_ })
            if ($sheetsToRemove.Count -gt 0) {
                Remove-Worksheet -Path $ExportDetails -WorksheetName $sheetsToRemove
                Write-Log -Type INFO -Message "Removed excluded worksheet(s) from existing workbook: $([string]::Join(', ', $sheetsToRemove))" -ExportFileLocation $ExportDetails
            }
        }
        catch {
            Write-Log -Type WARNING -Message "Unable to remove excluded worksheets from existing workbook before export: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }
    }
    
    # === Sheet ordering ===
    $desiredOrder = @(
        # Assessment Outputs
        "BestPractices", "BestPracticeFindings", "MigrationReadiness", "SecureScoreActions", "UnmanagedObjects", "OneDriveOwnerMismatches",

        # Licensing & Tenant Info
        "LicenseSKUs", "Domains", "AuthenticationMethods", "AuthenticationSSOApplications", "EnterpriseApplications", "AuthenticationConfig", "MfaEnrollmentSummary", "MfaEnforcementSummary", "MfaEnforcementGapUsers", "MfaEnforcementScopeReview", "ConditionalAccessPolicySummary", "SecurityDefaultsPolicy", "GuestSignInSummary", "GuestAccessConfiguration", "ExternalIdentityRestrictions", "PrivilegedAccessSummary", "Admins",

        # Users
        "Users", "UserFullDetails", "DeviceDetails",

        # Mailboxes
        "AllMailboxes", "PrimaryMailboxStats", "MailboxFullDetails", "SharedMailboxGovernanceSummary", "ForwardingPolicySummary", "InboxRuleForwardingSummary", "InboxRulesExternalForwarding", "ArchiveMailboxes", "ArchiveMailboxStats", "LitigationHoldMailboxes", "InactiveMailboxes", "InactiveMailboxDetails", "EmailActivityTopSenders", "EmailActivityTopReceivers", "TeamsActivityTopUsers", "Office365GroupsActivityTopGroups", "EmployeeExperienceInsightsSummary", "CollaborationActivitySummary", "NonUserMailboxes", "AllRecipients",

        # Groups
        "AllExchangeGroups", "UnifiedGroups", "AllTeams", "EntraIDGroups",

        # Public Folders
        "PublicFolderDetails", "PublicFolderPerms",

        # Mail Flow
        "MailFlowRules", "MailFlowConnectors", "RemoteDomains", "SMTPRelayConfig", "SMTPRelayServiceAccounts",

        # Security & Compliance
        "SecuritySecureScore", "ConditionalAccessPolicies", "SMTPRelaySummary", "TeamsVoiceSummary", "SpamFilteringConfig", "RetentionPolicies", "DlpPolicies", "PasswordLifecycleSummary",

        # Cloud Services
        "OneDrive",
        "SharePoint", "SharePointSharingSummary", "ExternalSharingSummary", "ExternalSharingSiteOverrides", "ExternalExposureFindings",

        # Hybrid / Infra
        "HybridConfiguration", "TenantInfo", "DeviceManagementSummary"
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
    foreach ($table in $orderedTables) {
        try {
            $worksheetName = Resolve-WorksheetExportName -LogicalName $table
            Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash To Excel"
            Write-Log -Type DEBUG -Message ("Exporting '{0}' Hash Table to '{1}' as worksheet '{2}'" -f $table, $ExportDetails, $worksheetName) -ExportFileLocation $ExportDetails

            $tableValue = $hashtable[$table]
            $sourceEnumerable = $null
            $sourceCount = 0
            $singleValue = $null
            if ($tableValue -is [hashtable] -or $tableValue -is [System.Collections.Specialized.OrderedDictionary]) {
                $sourceCount = $tableValue.Count
                $sourceEnumerable = $tableValue.Values
            } elseif ($tableValue -is [System.Collections.IEnumerable] -and -not ($tableValue -is [string])) {
                if ($tableValue.PSObject.Properties['Count']) {
                    try { $sourceCount = [int]$tableValue.Count } catch { $sourceCount = -1 }
                } else {
                    $sourceCount = -1
                }
                $sourceEnumerable = $tableValue
            } else {
                if ($null -ne $tableValue) {
                    $singleValue = $tableValue
                    $sourceCount = 1
                }
            }

            if ($sourceCount -lt 0 -and $null -ne $sourceEnumerable) {
                $sourceEnumerable = @($sourceEnumerable)
                $sourceCount = $sourceEnumerable.Count
            }

            if ($sourceCount -gt 0) {
                $attempt = 0
                $maxAttempts = 3
                $saved = $false
                $exportSource = if ($null -ne $singleValue) { @($singleValue) } else { $sourceEnumerable }
                $autoSizeSheet = ($sourceCount -le 5000)
                if (-not $autoSizeSheet) {
                    Write-Log -Type INFO -Message "Skipping AutoSize for worksheet '$worksheetName' due to row count ($sourceCount) to reduce export runtime/memory pressure." -ExportFileLocation $ExportDetails
                }
                while (-not $saved -and $attempt -lt $maxAttempts) {
                    $attempt++
                    if (Test-FileLocked -Path $ExportDetails) {
                        Write-Log -Type WARNING -Message "Excel file is locked (attempt $attempt/$maxAttempts). Close the file and retrying in 5 seconds..." -ExportFileLocation $ExportDetails
                        Start-Sleep -Seconds 5
                        continue
                    }
                    try {
                        $excelSplat = @{
                            Path          = $ExportDetails
                            WorksheetName = $worksheetName
                            ClearSheet    = $true
                            BoldTopRow    = $true
                        }
                        if ($autoSizeSheet) {
                            $excelSplat.AutoSize = $true
                        }

                        $exportSource |
                            ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                            Export-Excel @excelSplat
                        $saved = $true
                    } catch {
                        if ($attempt -lt $maxAttempts) {
                            Write-Log -Type WARNING -Message "Save failed for '$table' (attempt $attempt/$maxAttempts): $($_.Exception.Message). Retrying in 5 seconds..." -ExportFileLocation $ExportDetails
                            Start-Sleep -Seconds 5
                        } else {
                            throw
                        }
                    }
                }
            } else {
                if ($table -eq 'OneDriveOwnerMismatches') {
                    try {
                        $placeholderRows = @(
                            [PSCustomObject]@{
                                DisplayName          = 'N/A'
                                SiteUrl              = 'N/A'
                                CurrentOwner         = 'N/A'
                                ExpectedDefaultOwner = 'N/A'
                                Notes                = 'No owner mismatches detected in this run.'
                            }
                        )
                        $excelSplat = @{
                            Path          = $ExportDetails
                            WorksheetName = $worksheetName
                            ClearSheet    = $true
                            BoldTopRow    = $true
                            AutoSize      = $true
                        }
                        $placeholderRows |
                            ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                            Export-Excel @excelSplat
                        Write-Log -Type INFO -Message "No owner mismatch rows found; exported placeholder row to worksheet '$worksheetName'." -ExportFileLocation $ExportDetails
                        continue
                    }
                    catch {
                        Write-Log -Type WARNING -Message "Unable to export placeholder row for '$table': $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                    }
                }

                $logType = if ($optionalEmptySheets -contains $table) { 'INFO' } else { 'WARNING' }
                Write-Log -Type $logType -Message "No data found for $table to export to Excel" -ExportFileLocation $ExportDetails

                if (Test-Path -Path $ExportDetails) {
                    try {
                        $existingSheetNames = @()
                        if (Get-Command -Name Get-ExcelSheetInfo -ErrorAction SilentlyContinue) {
                            $existingSheetNames = @(Get-ExcelSheetInfo -Path $ExportDetails | Select-Object -ExpandProperty Name)
                        }

                        $staleCandidates = Get-WorksheetCleanupCandidates -LogicalName $table
                        $staleWorksheets = @($staleCandidates | Where-Object { $existingSheetNames -contains $_ } | Select-Object -Unique)
                        if ($staleWorksheets.Count -gt 0) {
                            Remove-Worksheet -Path $ExportDetails -WorksheetName $staleWorksheets
                            Write-Log -Type INFO -Message "Removed stale worksheet(s) '$($staleWorksheets -join ', ')' because '$table' produced no rows." -ExportFileLocation $ExportDetails
                        }
                    }
                    catch {
                        Write-Log -Type WARNING -Message "Unable to remove stale worksheet(s) for '$table': $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                    }
                }
            }
        }
        catch {
            Write-Log -Type Error -Message "An error occurred in Exporting Hash To Excel for $($table) to '$($ExportDetails)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }
    }
    Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash To Excel" -Completed
    Write-Host "The report has been exported to: $($ExportDetails)" -ForegroundColor Green
}
