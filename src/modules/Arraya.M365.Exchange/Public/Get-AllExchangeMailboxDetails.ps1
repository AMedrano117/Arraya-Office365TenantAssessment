#Gather all Exchange mailboxes and mailbox statistics
function Get-AllExchangeMailboxDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory = $false)]
        $Context
    )

    function Get-MailboxUsageReportPrincipalValue {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [object]$Row
        )

        foreach ($fieldName in @(
            'User Principal Name',
            'User Principal Name ',
            'User Principal Name (UPN)',
            'Owner Principal Name',
            'Account'
        )) {
            $rawValue = [string](Get-ArrayaObjectValue -Object $Row -Names @($fieldName))
            if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
                return $rawValue.Trim()
            }
        }

        return $null
    }

    function Format-MailboxUsageReportDebugSample {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [object]$Row
        )

        $propertyNames = @()
        if ($Row -and $Row.PSObject) {
            $propertyNames = @($Row.PSObject.Properties.Name | Select-Object -First 20)
        }

        $knownFieldPairs = @()
        foreach ($fieldName in @(
            'User Principal Name',
            'User Principal Name ',
            'User Principal Name (UPN)',
            'Owner Principal Name',
            'Account',
            'Display Name',
            'Last Activity Date',
            'Is Deleted',
            'Deleted Date',
            'Storage Used (Byte)',
            'Item Count'
        )) {
            $rawValue = [string](Get-ArrayaObjectValue -Object $Row -Names @($fieldName))
            if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
                $knownFieldPairs += ("{0}='{1}'" -f $fieldName, ($rawValue -replace "'", "''"))
            }
        }

        $dynamicPairs = @()
        if ($Row -and $Row.PSObject) {
            $dynamicPairs = @(
                $Row.PSObject.Properties |
                    Where-Object { $_.Name -match '(?i)(principal|upn|account|mail|owner|display)' } |
                    Select-Object -First 10 |
                    ForEach-Object {
                        $value = [string]$_.Value
                        if (-not [string]::IsNullOrWhiteSpace($value)) {
                            "{0}='{1}'" -f $_.Name, ($value -replace "'", "''")
                        }
                    }
            )
        }

        return "Type={0}; Properties=[{1}]; Known=[{2}]; Dynamic=[{3}]" -f `
            $Row.GetType().FullName, `
            ($propertyNames -join ', '), `
            ($knownFieldPairs -join '; '), `
            ($dynamicPairs -join '; ')
    }

    function Format-MailboxLookupDebugSample {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [object]$Mailbox
        )

        $lookupKey = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$Mailbox.UserPrincipalName)) {
            $lookupKey = ([string]$Mailbox.UserPrincipalName).Trim().ToLowerInvariant()
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$Mailbox.PrimarySmtpAddress)) {
            $lookupKey = ([string]$Mailbox.PrimarySmtpAddress).Trim().ToLowerInvariant()
        }

        return "DisplayName='{0}'; RecipientType='{1}'; UPN='{2}'; PrimarySMTP='{3}'; Alias='{4}'; ExchangeGuid='{5}'; Guid='{6}'; LookupKey='{7}'; WhenMailboxCreated='{8}'; WhenCreated='{9}'; IsInactive='{10}'; WasInactive='{11}'; WhenSoftDeleted='{12}'; UsageLocation='{13}'; AccountDisabled='{14}'; IsDirSynced='{15}'; HiddenFromAddressLists='{16}'" -f `
            (($Mailbox.DisplayName -as [string]) -replace "'", "''"), `
            (($Mailbox.RecipientTypeDetails -as [string]) -replace "'", "''"), `
            (($Mailbox.UserPrincipalName -as [string]) -replace "'", "''"), `
            (($Mailbox.PrimarySmtpAddress -as [string]) -replace "'", "''"), `
            (($Mailbox.Alias -as [string]) -replace "'", "''"), `
            (($Mailbox.ExchangeGuid -as [string]) -replace "'", "''"), `
            (($Mailbox.Guid -as [string]) -replace "'", "''"), `
            (($lookupKey -as [string]) -replace "'", "''"), `
            (($Mailbox.WhenMailboxCreated -as [string]) -replace "'", "''"), `
            (($Mailbox.WhenCreated -as [string]) -replace "'", "''"), `
            (($Mailbox.IsInactiveMailbox -as [string]) -replace "'", "''"), `
            (($Mailbox.WasInactiveMailbox -as [string]) -replace "'", "''"), `
            (($Mailbox.WhenSoftDeleted -as [string]) -replace "'", "''"), `
            (($Mailbox.UsageLocation -as [string]) -replace "'", "''"), `
            (($Mailbox.AccountDisabled -as [string]) -replace "'", "''"), `
            (($Mailbox.IsDirSynced -as [string]) -replace "'", "''"), `
            (($Mailbox.HiddenFromAddressListsEnabled -as [string]) -replace "'", "''")
    }

    function Test-DictionaryContainsKey {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [object]$Dictionary,
            [Parameter(Mandatory = $true)]
            [string]$Key
        )

        if ($Dictionary.PSObject.Methods['ContainsKey']) {
            return [bool]$Dictionary.ContainsKey($Key)
        }

        if ($Dictionary -is [System.Collections.IDictionary]) {
            return [bool]$Dictionary.Contains($Key)
        }

        return $false
    }

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context -DetailLevel $detailLevel
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $collectionDepthPolicy = $Context.Policies['CollectionDepth']
    $initialStart = Get-ArrayaExchangeCollectorStartTime -Context $Context
    if (-not $Context.Runtime.Contains('MailboxUsageGraphLookup') -or -not ($Context.Runtime['MailboxUsageGraphLookup'] -is [System.Collections.IDictionary])) {
        $Context.Runtime['MailboxUsageGraphLookup'] = @{}
    }
    if (-not $Context.Runtime.Contains('UnifiedGroupsInventoryCache')) {
        $Context.Runtime['UnifiedGroupsInventoryCache'] = @()
    }
    if (-not $Context.Runtime.Contains('Office365GroupsActivityMailboxLookup')) {
        $Context.Runtime['Office365GroupsActivityMailboxLookup'] = $null
    }
    $mailboxUsageGraphLookup = $Context.Runtime['MailboxUsageGraphLookup']
    $unifiedGroupsInventoryCache = @($Context.Runtime['UnifiedGroupsInventoryCache'])

    $start = Get-Date
    $mailboxInventoryProgressId = 30
    Write-Host "Getting all mailboxes and inactive mailboxes with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] START: Getting all mailboxes with $($detailLevel) details" -ExportFileLocation $exportDetails
    try {
        # Gather Mailboxes - Include InActive Mailboxes
        switch ($detailLevel) {
            "minimum" {
                $Properties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress"
                    "Identity", "Guid", "ExchangeGuid", "ArchiveStatus", "ArchiveState", "ArchiveGuid", "ArchiveName"
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted"
                    "LitigationHoldEnabled", "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy"
                    "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias", "EmailAddresses"
                )

                $DesiredProperties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress",
                    "Identity", "Guid", "ExchangeGuid", "ArchiveStatus", "ArchiveState", "ArchiveGuid",
                    @{Name="ArchiveName"; Expression={$_.ArchiveName -join ","}},
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted",
                    "LitigationHoldEnabled", "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy",
                    "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias",
                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}
                )

                $exoMailboxes = Invoke-QuietCommand -ScriptBlock {
                    Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -Properties $Properties -IncludeInactiveMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Select-Object $DesiredProperties
                }
            }
            "combined" {
                # Combined mode optimization: request a slimmer property set to reduce EXO payload and local projection overhead.
                $Properties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Office", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress"
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted"
                    "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias", "EmailAddresses"
                    "Identity", "WhenCreated", "Guid", "DeliverToMailboxAndForward", "ForwardingAddress"
                    "ForwardingSmtpAddress", "LitigationHoldEnabled", "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy"
                    "ExchangeGuid", "ArchiveStatus", "ArchiveState", "ArchiveGuid", "ArchiveName", "AutoExpandingArchiveEnabled"
                )

                $DesiredProperties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Office", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress",
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted",
                    "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias",
                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}},
                    "Identity", "WhenCreated", "Guid", "DeliverToMailboxAndForward", "ForwardingAddress", "ForwardingSmtpAddress",
                    "LitigationHoldEnabled", "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy",
                    "ExchangeGuid", "ArchiveStatus", "ArchiveState", "ArchiveGuid",
                    @{Name="ArchiveName"; Expression={$_.ArchiveName -join ","}}, "AutoExpandingArchiveEnabled"
                )

                $exoMailboxes = Invoke-QuietCommand -ScriptBlock {
                    Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -Properties $Properties -IncludeInactiveMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Select-Object $DesiredProperties
                }
            }
            "all" {
                $Properties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Office", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress"
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted"
                    "InPlaceHolds", "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias"
                    "EmailAddresses", "GrantSendOnBehalfTo", "AcceptMessagesOnlyFrom", "AcceptMessagesOnlyFromDLMembers", "AcceptMessagesOnlyFromSendersOrMembers"
                    "RejectMessagesFrom", "RejectMessagesFromDLMembers", "RejectMessagesFromSendersOrMembers", "RequireSenderAuthenticationEnabled", "WindowsEmailAddress"
                    "DistinguishedName", "Identity", "WhenChanged", "WhenCreated", "ExchangeObjectId"
                    "Guid", "DeliverToMailboxAndForward", "ForwardingAddress", "ForwardingSmtpAddress", "LitigationHoldEnabled"
                    "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy", "ExchangeGuid", "IsResource"
                    "IsShared", "ResourceType", "RoomMailboxAccountEnabled", "WindowsLiveID", "MicrosoftOnlineServicesID"
                    "EffectivePublicFolderMailbox", "MailboxPlan", "ArchiveStatus", "ArchiveState", "ArchiveName"
                    "ArchiveGuid", "AutoExpandingArchiveEnabled", "DisabledArchiveGuid", "PersistedCapabilities"
                )

                $DesiredProperties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Office", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress",
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted",
                    @{Name="InPlaceHolds"; Expression={$_.InPlaceHolds -join ","}},"AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias",
                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, 
                    @{Name="GrantSendOnBehalfTo"; Expression={$_.GrantSendOnBehalfTo -join ","}}, 
                    @{Name="AcceptMessagesOnlyFrom"; Expression={$_.AcceptMessagesOnlyFrom -join ","}}, 
                    @{Name="AcceptMessagesOnlyFromDLMembers"; Expression={$_.AcceptMessagesOnlyFromDLMembers -join ","}}, 
                    @{Name="AcceptMessagesOnlyFromSendersOrMembers"; Expression={$_.AcceptMessagesOnlyFromSendersOrMembers -join ","}}, 
                    @{Name="RejectMessagesFrom"; Expression={$_.RejectMessagesFrom -join ","}}, 
                    @{Name="RejectMessagesFromDLMembers"; Expression={$_.RejectMessagesFromDLMembers -join ","}}, 
                    @{Name="RejectMessagesFromSendersOrMembers"; Expression={$_.RejectMessagesFromSendersOrMembers -join ","}}, 
                    "RequireSenderAuthenticationEnabled", "WindowsEmailAddress",
                    "DistinguishedName", "Identity", "WhenChanged", "WhenCreated", "ExchangeObjectId",
                    "Guid", "DeliverToMailboxAndForward", "ForwardingAddress", "ForwardingSmtpAddress", "LitigationHoldEnabled",
                    "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy", "ExchangeGuid", "IsResource",
                    "IsShared", "ResourceType", "RoomMailboxAccountEnabled", "WindowsLiveID", "MicrosoftOnlineServicesID", "EffectivePublicFolderMailbox", "MailboxPlan", 
                    "ArchiveStatus", "ArchiveState", @{Name="ArchiveName"; Expression={$_.ArchiveName -join ","}}, "ArchiveGuid", "AutoExpandingArchiveEnabled", "DisabledArchiveGuid",
                    @{Name="PersistedCapabilities"; Expression={$_.PersistedCapabilities -join ","}}
                )

                $exoMailboxes = Invoke-QuietCommand -ScriptBlock {
                    Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -Properties $Properties -IncludeInactiveMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Select-Object $DesiredProperties
                }
            }
            geek {
                $exoMailboxes = Invoke-QuietCommand -ScriptBlock {
                    Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -IncludeInactiveMailbox -PropertySets All -ResultSize Unlimited -ErrorAction SilentlyContinue
                }
            }
        }
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering all mailboxes (Get-EXOMailbox) including Inactive Mailboxes" -ExportFileLocation $exportDetails
        
        $tenantStatsHash["AllMailboxes"] = @{}
        $tenantStatsHash["AllMailboxes-MailIdentity"] = @{}
        $tenantStatsHash["AllMailboxes-UserPrincipalName"] = @{}
        $tenantStatsHash["AllMailboxes-PrimarySmtpAddress"] = @{}
        $tenantStatsHash["NonUserMailboxes"] = @{}
        $tenantStatsHash["ArchiveMailboxes"] = @{}
        $tenantStatsHash["InactiveMailboxes"] = @{}
        $tenantStatsHash["LitigationHoldMailboxes"] = @{}

        # Insert individual mailboxes into the hashtable
        $totalCount = $exoMailboxes.Count
        foreach ($mailbox in $exoMailboxes) {
            $key = @(
                [string]$mailbox.ExternalDirectoryObjectId
                if ($mailbox.ExchangeGuid) { [string]$mailbox.ExchangeGuid }
                if ($mailbox.Guid) { [string]$mailbox.Guid }
                [string]$mailbox.UserPrincipalName
                [string]$mailbox.PrimarySmtpAddress
                [string]$mailbox.Identity
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
            if ([string]::IsNullOrWhiteSpace($key)) {
                $key = "mailbox:$([guid]::NewGuid().Guid)"
            }
            Write-ProgressHelper -Total $totalCount -Id $mailboxInventoryProgressId -Activity "Gathering All Exchange Mailbox Details" -Operation "Gathering Mailbox Details for $($key)"
            
            # Set the key based on the mailbox type
            #$MailboxTypeKey = $mailbox.RecipientTypeDetails.tostring() # Use RecipientTypeDetails as the key
            $tenantStatsHash["AllMailboxes"][$key] = $mailbox
            $tenantStatsHash["AllMailboxes-MailIdentity"][$mailbox.Identity] = $mailbox
            if ($mailbox.UserPrincipalName) {
                $tenantStatsHash["AllMailboxes-UserPrincipalName"][[string]$mailbox.UserPrincipalName] = $mailbox
            }
            if ($mailbox.PrimarySmtpAddress) {
                $tenantStatsHash["AllMailboxes-PrimarySmtpAddress"][[string]$mailbox.PrimarySmtpAddress] = $mailbox
            }

            if ($mailbox.RecipientTypeDetails -ne "UserMailbox" -and $mailbox.RecipientTypeDetails -ne "GroupMailbox") {
                $tenantStatsHash["NonUserMailboxes"][$key] = $mailbox
            }
            if ($mailbox.ArchiveStatus -eq "Active") {
                $tenantStatsHash["ArchiveMailboxes"][$key] = $mailbox
            }
            if ($mailbox.IsInactiveMailbox -eq $true) {
                $tenantStatsHash["InactiveMailboxes"][$key] = $mailbox
            }
            if ($mailbox.LitigationHoldEnabled -eq $true) {
                $tenantStatsHash["LitigationHoldMailboxes"][$key] = $mailbox
            }
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Mailbox Details. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    } finally {
        Write-ProgressHelper -Total ([Math]::Max($totalCount, 1)) -Id $mailboxInventoryProgressId -Activity "Gathering All Exchange Mailbox Details" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] COMPLETED: Gathering All Mailbox Details in $($CompletedTime)" -ExportFileLocation $exportDetails
    }

    #Mailbox Statistics to Hash Table
    ###########################################################################################################################################
    $primaryStatsProgressId = 31
    $archiveStatsProgressId = 32

    ## Primary Mailbox Stats
    try {
        $start = Get-Date
        Write-Progress -Id $primaryStatsProgressId -Activity "Gathering All Primary Mailbox Statistics" -Status (((Get-Date) - $initialStart).ToString('hh\:mm\:ss'))
        Write-Host "  Getting primary mailbox stats..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering all primary mailbox statistics for collected mailboxes (including inactive where available)." -ExportFileLocation $exportDetails

        $tenantStatsHash["PrimaryMailboxStats"] = @{}
        $Context.Runtime['MailboxUsageGraphLookup'] = @{}
        $mailboxUsageGraphLookup = $Context.Runtime['MailboxUsageGraphLookup']
        $activeMailboxes = $tenantStatsHash['AllMailboxes'].Values | Where-Object { $_.IsInactiveMailbox -ne $true }
        $graphEligibleMailboxes = @(
            $activeMailboxes | Where-Object {
                $_.PSObject.Properties['RecipientTypeDetails'] -and
                @('UserMailbox', 'SharedMailbox') -contains ([string]$_.RecipientTypeDetails)
            }
        )
        $graphLikelyEligibleMailboxCount = $graphEligibleMailboxes.Count
        $activeMailboxTypeBreakdown = @(
            $activeMailboxes |
                Group-Object -Property RecipientTypeDetails |
                Sort-Object -Property Count -Descending |
                ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }
        ) -join '; '
        $activeMailboxUnmatchedAfterGraphCount = @($activeMailboxes).Count
        $graphEligibleMailboxUnmatchedCount = $graphLikelyEligibleMailboxCount
        $graphUnsupportedActiveMailboxCount = [Math]::Max(($activeMailboxes.Count - $graphLikelyEligibleMailboxCount), 0)
        $graphStatsCount = 0
        $graphReportRowCount = 0
        $graphMatchedMailboxCount = 0
        $graphUsablePrincipalRowCount = 0
        $graphRowsWithoutPrincipalCount = 0
        $graphRowsPotentiallyObscuredPrincipalCount = 0
        $graphRowsDuplicatePrincipalCount = 0
        $activeMailboxWithoutLookupKeyCount = 0
        $activeMailboxLookupMissCount = 0
        $activeMailboxMatchedByGraphCount = 0
        $graphRowsWithoutPrincipalSamples = New-Object System.Collections.Generic.List[string]
        $graphUsablePrincipalSamples = New-Object System.Collections.Generic.List[string]
        $activeMailboxWithoutLookupKeySamples = New-Object System.Collections.Generic.List[string]
        $activeMailboxLookupMissSamples = New-Object System.Collections.Generic.List[string]
        $exoFallbackRequestedCount = 0
        $exoFallbackReturnedCount = 0
        $graphPhaseSeconds = 0
        $exoFallbackPhaseSeconds = 0
        $unifiedPreCachePhaseSeconds = 0
        $graphCoverageWarning = $null
        $graphIdentifiableNamesInReports = $null
        $shouldPreCacheUnifiedGroupStats = Test-ShouldCollectUnifiedGroupMailboxStats -Context $Context -DetailLevel $detailLevel
        $estimatedGroupMailboxCount = @(
            $tenantStatsHash['AllMailboxes'].Values | Where-Object {
                $_.PSObject.Properties['RecipientTypeDetails'] -and [string]$_.RecipientTypeDetails -eq 'GroupMailbox'
            }
        ).Count
        if ($shouldPreCacheUnifiedGroupStats) {
            #Write-Host ("    Step 3/3 Unified group pre-cache: scheduled | estimated group mailboxes in inventory={0}" -f $estimatedGroupMailboxCount) -ForegroundColor DarkGray
        }

        # Try Graph mailbox usage report (fast) for active mailboxes
        $graphPhaseStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $adminReportSettings = Office365Custom\Get-GraphData -Uri 'https://graph.microsoft.com/v1.0/admin/reportSettings' -Activity 'Admin report settings' -PageSize 50
            if ($adminReportSettings -and (Get-ArrayaObjectValue -Object $adminReportSettings -Names @('displayConcealedNames')) -ne $null) {
                $displayConcealedNames = Get-ArrayaObjectValue -Object $adminReportSettings -Names @('displayConcealedNames')
                try { $graphIdentifiableNamesInReports = -not [bool]$displayConcealedNames } catch { $graphIdentifiableNamesInReports = $null }
            }

            $graphReportData = @(Office365Custom\Get-GraphAPIActivityReport -ServiceName 'MailboxUsage' -PeriodDuration 'D180')
            $graphReportRowCount = @($graphReportData).Count
            if ($graphReportData -and $graphReportData.Count -gt 0) {
                foreach ($item in $graphReportData) {
                    $principalValue = Get-MailboxUsageReportPrincipalValue -Row $item
                    if ([string]::IsNullOrWhiteSpace($principalValue)) {
                        $graphRowsWithoutPrincipalCount++
                        if ($graphRowsWithoutPrincipalSamples.Count -lt 5) {
                            [void]$graphRowsWithoutPrincipalSamples.Add((Format-MailboxUsageReportDebugSample -Row $item))
                        }
                        continue
                    }

                    $principalKey = $principalValue.Trim().ToLowerInvariant()
                    if (-not ($principalValue -match '@')) {
                        $graphRowsPotentiallyObscuredPrincipalCount++
                    }
                    if (Test-DictionaryContainsKey -Dictionary $mailboxUsageGraphLookup -Key $principalKey) {
                        $graphRowsDuplicatePrincipalCount++
                        continue
                    }

                    $mailboxUsageGraphLookup[$principalKey] = $item
                    $graphUsablePrincipalRowCount++
                    if ($graphUsablePrincipalSamples.Count -lt 5) {
                        $graphUsablePrincipalSamples.Add(("PrincipalKey='{0}'; {1}" -f ($principalKey -replace "'", "''"), (Format-MailboxUsageReportDebugSample -Row $item))) | Out-Null
                    }
                }

                foreach ($mailbox in $graphEligibleMailboxes) {
                    $mailboxLookupKey = $null
                    if (-not [string]::IsNullOrWhiteSpace([string]$mailbox.UserPrincipalName)) {
                        $mailboxLookupKey = ([string]$mailbox.UserPrincipalName).ToLowerInvariant()
                    }
                    elseif (-not [string]::IsNullOrWhiteSpace([string]$mailbox.PrimarySmtpAddress)) {
                        $mailboxLookupKey = ([string]$mailbox.PrimarySmtpAddress).ToLowerInvariant()
                    }

                    if (-not $mailboxLookupKey) {
                        $activeMailboxWithoutLookupKeyCount++
                        if ($activeMailboxWithoutLookupKeySamples.Count -lt 5) {
                            [void]$activeMailboxWithoutLookupKeySamples.Add((Format-MailboxLookupDebugSample -Mailbox $mailbox))
                        }
                        continue
                    }
                    if (-not (Test-DictionaryContainsKey -Dictionary $mailboxUsageGraphLookup -Key $mailboxLookupKey)) {
                        $activeMailboxLookupMissCount++
                        if ($activeMailboxLookupMissSamples.Count -lt 5) {
                            [void]$activeMailboxLookupMissSamples.Add((Format-MailboxLookupDebugSample -Mailbox $mailbox))
                        }
                        continue
                    }

                    $graphData = $mailboxUsageGraphLookup[$mailboxLookupKey]
                    $storageBytes = [int64](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $graphData -Names @('Storage Used (Byte)', 'Storage Used (Bytes)', 'Storage Used')) -AsInt64)
                    $deletedBytes = [int64](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $graphData -Names @('Deleted Item Size (Byte)', 'Deleted Item Size (Bytes)', 'Deleted Item Size')) -AsInt64)
                    $itemCount = Get-ArrayaObjectValue -Object $graphData -Names @('Item Count')
                    if ([string]::IsNullOrWhiteSpace([string]$itemCount)) { $itemCount = '0' }
                    $guidKey = if ($mailbox.ExchangeGuid) { Convert-ToMailboxGuidKey -GuidValue $mailbox.ExchangeGuid } elseif ($mailbox.Guid) { Convert-ToMailboxGuidKey -GuidValue $mailbox.Guid } else { $null }
                    if (-not $guidKey) { continue }
                    $stats = [PSCustomObject]@{
                        DisplayName = $mailbox.DisplayName
                        TotalItemSize = "$([math]::Round($storageBytes / 1GB, 4)) GB ($storageBytes bytes)"
                        TotalItemSizeBytes = [int64]$storageBytes
                        ItemCount = $itemCount
                        TotalDeletedItemSize = "$([math]::Round($deletedBytes / 1GB, 4)) GB ($deletedBytes bytes)"
                        TotalDeletedItemSizeBytes = [int64]$deletedBytes
                        MailboxType = $mailbox.RecipientTypeDetails
                        MailboxGuid = if ($mailbox.ExchangeGuid) { $mailbox.ExchangeGuid } else { $mailbox.Guid }
                    }
                    $tenantStatsHash["PrimaryMailboxStats"][$guidKey] = $stats
                    $activeMailboxMatchedByGraphCount++
                }
            }
            $graphStatsCount = $tenantStatsHash["PrimaryMailboxStats"].Count
            $graphMatchedMailboxCount = $graphStatsCount
            $activeMailboxUnmatchedAfterGraphCount = [Math]::Max(($activeMailboxes.Count - $activeMailboxMatchedByGraphCount), 0)
            $graphEligibleMailboxUnmatchedCount = [Math]::Max(($graphEligibleMailboxes.Count - $activeMailboxMatchedByGraphCount), 0)
            if (
                $graphReportRowCount -gt 0 -and
                $graphMatchedMailboxCount -eq 0 -and
                $graphRowsPotentiallyObscuredPrincipalCount -gt 0
            ) {
                $graphCoverageWarning = "Graph mailbox report appears to contain obfuscated principals (non-UPN rows=$graphRowsPotentiallyObscuredPrincipalCount), which can force EXO fallback."
            }
            elseif (
                $graphReportRowCount -gt 0 -and
                $graphLikelyEligibleMailboxCount -gt 0 -and
                $graphMatchedMailboxCount -eq 0 -and
                $graphIdentifiableNamesInReports -eq $true
            ) {
                $graphCoverageWarning = "Graph mailbox report contains identifiable names but did not match any eligible Graph mailboxes. Falling back to EXO for all eligible Graph mailboxes."
            }
        } catch {
            Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] Graph mailbox usage report failed: $($_.Exception.Message)" -ExportFileLocation $exportDetails
        }
        finally {
            if ($graphPhaseStopwatch -and $graphPhaseStopwatch.IsRunning) {
                $graphPhaseStopwatch.Stop()
            }
            if ($graphPhaseStopwatch) {
                $graphPhaseSeconds = [math]::Round($graphPhaseStopwatch.Elapsed.TotalSeconds, 2)
            }
        }
        Write-Host ("    Step 1/3 Graph mailbox usage: {0}s | rows={1}, usable={2}, matched={3}/{4} eligible, unresolvedEligible={5}, unsupportedActive={6}" -f $graphPhaseSeconds, $graphReportRowCount, $graphUsablePrincipalRowCount, $graphMatchedMailboxCount, $graphLikelyEligibleMailboxCount, $graphEligibleMailboxUnmatchedCount, $graphUnsupportedActiveMailboxCount) -ForegroundColor DarkGray
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 1/3 (Graph mailbox usage) completed in $graphPhaseSeconds sec; reportRows=$graphReportRowCount; usablePrincipalRows=$graphUsablePrincipalRowCount; missingPrincipalRows=$graphRowsWithoutPrincipalCount; nonUpnRows=$graphRowsPotentiallyObscuredPrincipalCount; duplicatePrincipalRows=$graphRowsDuplicatePrincipalCount; activeMailboxCount=$($activeMailboxes.Count); eligibleGraphMailboxes=$graphLikelyEligibleMailboxCount; eligibleMatched=$activeMailboxMatchedByGraphCount; eligibleWithoutLookupKey=$activeMailboxWithoutLookupKeyCount; eligibleLookupMiss=$activeMailboxLookupMissCount; eligibleUnmatched=$graphEligibleMailboxUnmatchedCount; unsupportedActiveMailboxTypes=$graphUnsupportedActiveMailboxCount; populated=$graphMatchedMailboxCount; unresolvedActive=$activeMailboxUnmatchedAfterGraphCount; identifiableNamesInReports=$graphIdentifiableNamesInReports." -ExportFileLocation $exportDetails
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Graph mailbox usage endpoint scope diagnostics: likelyEligibleGraphMailboxes=$graphLikelyEligibleMailboxCount; activeMailboxCount=$($activeMailboxes.Count); activeMailboxTypeBreakdown='$activeMailboxTypeBreakdown'. This Graph report can cover user/shared mailbox activity, but it does not fully represent resource/group/system mailbox inventory." -ExportFileLocation $exportDetails
        if ($graphReportRowCount -gt 0 -and $graphUsablePrincipalRowCount -eq 0) {
            Write-Host "    Graph join diagnostics: no usable principals resolved from mailbox usage rows; see run log for sample row fields." -ForegroundColor Yellow
            Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] Graph mailbox usage returned $graphReportRowCount row(s) but zero usable principal values were extracted. AdminReportSettings.identifiableNamesInReports=$graphIdentifiableNamesInReports. SampleRows=$($graphRowsWithoutPrincipalSamples -join ' || ')." -ExportFileLocation $exportDetails
        }
        elseif ($graphUsablePrincipalRowCount -gt 0 -and $graphMatchedMailboxCount -eq 0 -and $graphLikelyEligibleMailboxCount -gt 0) {
            Write-Host "    Graph join diagnostics: principal rows were present but none matched eligible Graph mailboxes; see run log for sample keys." -ForegroundColor Yellow
            Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] Graph mailbox usage produced usable principal rows but none matched eligible mailboxes. SampleGraphPrincipalRows=$($graphUsablePrincipalSamples -join ' || '); SampleEligibleMailboxLookupMisses=$($activeMailboxLookupMissSamples -join ' || ')." -ExportFileLocation $exportDetails
        }
        elseif ($activeMailboxLookupMissCount -gt 0) {
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Graph mailbox usage matched most eligible Graph mailboxes, but $activeMailboxLookupMissCount eligible mailbox(es) were still absent from the Graph lookup. EXO fallback will populate statistics for the remaining mailboxes. SampleEligibleMailboxLookupMisses=$($activeMailboxLookupMissSamples -join ' || '); SampleGraphPrincipalRows=$($graphUsablePrincipalSamples -join ' || ')." -ExportFileLocation $exportDetails
        }
        if ($activeMailboxWithoutLookupKeyCount -gt 0) {
            Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] Some eligible mailboxes did not expose a Graph lookup key (UPN/PrimarySMTP). SampleEligibleMailboxesWithoutLookupKey=$($activeMailboxWithoutLookupKeySamples -join ' || ')." -ExportFileLocation $exportDetails
        }
        if (-not [string]::IsNullOrWhiteSpace($graphCoverageWarning)) {
            Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] $graphCoverageWarning" -ExportFileLocation $exportDetails
        }

        Write-Progress -Id $primaryStatsProgressId -Activity "Gathering All Primary Mailbox Statistics" -Completed

        # Fallback to EXO stats for unresolved mailboxes only (Graph-first, profile-aware filtering)
        $mailboxesNeedingStatsList = New-Object System.Collections.Generic.List[object]
        $profileSkippedForFallbackCount = 0
        $deferredUnifiedGroupMailboxCount = 0
        $exoFallbackActiveCandidateCount = 0
        $exoFallbackInactiveCandidateCount = 0
        $minimumExcludedRecipientTypesForExoStats = @(
            'AuditLogMailbox',
            'AuxAuditLogMailbox',
            'ArbitrationMailbox',
            'DiscoveryMailbox',
            'MonitoringMailbox',
            'PublicFolderMailbox',
            'SupervisoryReviewPolicyMailbox'
        )
        $isMinimumDepth = ($collectionDepthPolicy -and $collectionDepthPolicy.IsMinimum)
        foreach ($mailbox in $tenantStatsHash['AllMailboxes'].Values) {
            if (Test-MailboxStatCached -MailboxRecord $mailbox -StatsHash $tenantStatsHash["PrimaryMailboxStats"]) {
                continue
            }

            if (
                $shouldPreCacheUnifiedGroupStats -and
                $mailbox.PSObject.Properties['RecipientTypeDetails'] -and
                [string]$mailbox.RecipientTypeDetails -eq 'GroupMailbox'
            ) {
                $deferredUnifiedGroupMailboxCount++
                continue
            }

            if (
                $isMinimumDepth -and
                $mailbox.PSObject.Properties['RecipientTypeDetails'] -and
                $minimumExcludedRecipientTypesForExoStats -contains ([string]$mailbox.RecipientTypeDetails)
            ) {
                $profileSkippedForFallbackCount++
                continue
            }

            [void]$mailboxesNeedingStatsList.Add($mailbox)
            if ($mailbox.PSObject.Properties['IsInactiveMailbox'] -and $mailbox.IsInactiveMailbox -eq $true) {
                $exoFallbackInactiveCandidateCount++
            }
            else {
                $exoFallbackActiveCandidateCount++
            }
        }
        $mailboxesNeedingStats = @($mailboxesNeedingStatsList.ToArray())
        $exoFilledCount = 0
        if ($profileSkippedForFallbackCount -gt 0) {
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Minimum mode optimization skipped EXO mailbox stats fallback for $profileSkippedForFallbackCount system mailbox(es)." -ExportFileLocation $exportDetails
        }
        if ($deferredUnifiedGroupMailboxCount -gt 0) {
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Deferred EXO mailbox stats fallback for $deferredUnifiedGroupMailboxCount unified group mailbox(es) to unified-group pre-cache stage." -ExportFileLocation $exportDetails
        }
        if ($mailboxesNeedingStats.Count -gt 0) {
            $exoFallbackStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $exoFallbackRequestedCount = $mailboxesNeedingStats.Count
            Write-Host ("    Step 2/3 EXO fallback: starting | unresolved={0} (active={1}, inactive={2}, deferredGroups={3}, profileSkipped={4})" -f $mailboxesNeedingStats.Count, $exoFallbackActiveCandidateCount, $exoFallbackInactiveCandidateCount, $deferredUnifiedGroupMailboxCount, $profileSkippedForFallbackCount) -ForegroundColor DarkGray
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Graph report covered $($tenantStatsHash['PrimaryMailboxStats'].Count) mailboxes; fetching EXO stats for $($mailboxesNeedingStats.Count) missing/inactive." -ExportFileLocation $exportDetails

            $allRemainingMBXStatsResult = Get-ExoMailboxStatisticsSafe -MailboxObjects $mailboxesNeedingStats -Context $Context -ProgressActivity "Gathering All Primary Mailbox Statistics" -ProgressId $primaryStatsProgressId
            $allRemainingMBXStats = @($allRemainingMBXStatsResult.Results)
            $exoFallbackReturnedCount = $allRemainingMBXStats.Count
            $allRemainingMBXStats | ForEach-Object {
                $key = Convert-ToMailboxGuidKey -GuidValue $_.MailboxGuid
                if (-not $key) { continue }
                if (-not $tenantStatsHash["PrimaryMailboxStats"].ContainsKey($key)) {
                    $tenantStatsHash["PrimaryMailboxStats"][$key] = $_
                    $exoFilledCount++
                }
            }
            Write-ExoStatisticsFailureSummary -OperationName 'Get-AllExchangeMailboxDetails primary mailbox statistics' -Context $Context -Failures $allRemainingMBXStatsResult.Failures
            $exoFallbackStopwatch.Stop()
            $exoFallbackPhaseSeconds = [math]::Round($exoFallbackStopwatch.Elapsed.TotalSeconds, 2)
            Write-Host ("    Step 2/3 EXO fallback: {0}s | requested={1}, returned={2}, populated={3}" -f $exoFallbackPhaseSeconds, $exoFallbackRequestedCount, $exoFallbackReturnedCount, $exoFilledCount) -ForegroundColor DarkGray
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 2/3 (EXO fallback) completed in $exoFallbackPhaseSeconds sec; requested=$exoFallbackRequestedCount; returned=$exoFallbackReturnedCount; populated=$exoFilledCount." -ExportFileLocation $exportDetails
        }
        else {
            Write-Host ("    Step 2/3 EXO fallback: skipped | no unresolved mailboxes (deferredGroups={0}, profileSkipped={1})" -f $deferredUnifiedGroupMailboxCount, $profileSkippedForFallbackCount) -ForegroundColor DarkGray
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 2/3 (EXO fallback) skipped; no unresolved mailboxes. deferredUnifiedGroups=$deferredUnifiedGroupMailboxCount; profileSkipped=$profileSkippedForFallbackCount." -ExportFileLocation $exportDetails
        }

        $finalStatsCount = $tenantStatsHash["PrimaryMailboxStats"].Count
        if ($exoFilledCount -eq 0 -and $finalStatsCount -gt $graphStatsCount) {
            $exoFilledCount = $finalStatsCount - $graphStatsCount
        }
        Write-Verbose "exoStatsFilledCount: $exoFilledCount"
        Write-Verbose "graphStatsCount: $graphStatsCount"
        Write-Verbose "finalStatsCount: $finalStatsCount"

        if ($exoFilledCount -lt 0) { $exoFilledCount = 0 }
        $totalMailboxes = $tenantStatsHash['AllMailboxes'].Values.Count
        $missingMailboxStatsCount = [Math]::Max(($totalMailboxes - $finalStatsCount), 0)
        $tenantStatsHash["PrimaryMailboxStatsCollectionSummary"] = [PSCustomObject]@{
            GraphReportRows        = [int]$graphReportRowCount
            GraphUsablePrincipalRows = [int]$graphUsablePrincipalRowCount
            GraphRowsWithoutPrincipal = [int]$graphRowsWithoutPrincipalCount
            GraphRowsPotentiallyObscuredPrincipal = [int]$graphRowsPotentiallyObscuredPrincipalCount
            GraphRowsDuplicatePrincipal = [int]$graphRowsDuplicatePrincipalCount
            GraphPopulated         = [int]$graphMatchedMailboxCount
            ActiveMailboxes        = [int]$activeMailboxes.Count
            GraphEligibleMailboxes = [int]$graphLikelyEligibleMailboxCount
            ActiveMailboxMatchedByGraph = [int]$activeMailboxMatchedByGraphCount
            ActiveMailboxesWithoutLookupKey = [int]$activeMailboxWithoutLookupKeyCount
            ActiveMailboxLookupMiss = [int]$activeMailboxLookupMissCount
            GraphEligibleMailboxUnmatched = [int]$graphEligibleMailboxUnmatchedCount
            ActiveMailboxUnmatchedAfterGraph = [int]$activeMailboxUnmatchedAfterGraphCount
            GraphUnsupportedActiveMailboxTypes = [int]$graphUnsupportedActiveMailboxCount
            GraphIdentifiableNamesInReports = $graphIdentifiableNamesInReports
            ExoFallbackRequested   = [int]$exoFallbackRequestedCount
            ExoFallbackRequestedActive = [int]$exoFallbackActiveCandidateCount
            ExoFallbackRequestedInactive = [int]$exoFallbackInactiveCandidateCount
            ExoFallbackReturned    = [int]$exoFallbackReturnedCount
            ExoFallbackPopulated   = [int]$exoFilledCount
            DeferredUnifiedGroupFallback = [int]$deferredUnifiedGroupMailboxCount
            ProfileSkippedFallback = [int]$profileSkippedForFallbackCount
            GraphPhaseSeconds      = [double]$graphPhaseSeconds
            ExoFallbackPhaseSeconds = [double]$exoFallbackPhaseSeconds
            Populated              = [int]$finalStatsCount
            Missing                = [int]$missingMailboxStatsCount
            TotalMailboxes         = [int]$totalMailboxes
            GraphCoverageWarning   = $graphCoverageWarning
        }
        Write-Host ("  Primary mailbox stats source breakdown: Graph={0}, EXO fallback={1}, Missing={2}, ActiveUnmatchedAfterGraph={3}" -f $graphMatchedMailboxCount, $exoFilledCount, $missingMailboxStatsCount, $activeMailboxUnmatchedAfterGraphCount) -ForegroundColor DarkGray
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Mailbox stats source breakdown: GraphReportRows=$graphReportRowCount; GraphUsablePrincipalRows=$graphUsablePrincipalRowCount; GraphRowsWithoutPrincipal=$graphRowsWithoutPrincipalCount; GraphRowsPotentiallyObscuredPrincipal=$graphRowsPotentiallyObscuredPrincipalCount; GraphRowsDuplicatePrincipal=$graphRowsDuplicatePrincipalCount; GraphPopulated=$graphMatchedMailboxCount; ActiveMailboxCount=$($activeMailboxes.Count); ActiveMailboxMatchedByGraph=$activeMailboxMatchedByGraphCount; ActiveMailboxWithoutLookupKey=$activeMailboxWithoutLookupKeyCount; ActiveMailboxLookupMiss=$activeMailboxLookupMissCount; ActiveMailboxUnmatchedAfterGraph=$activeMailboxUnmatchedAfterGraphCount; EXOFallbackRequested=$exoFallbackRequestedCount; EXOFallbackRequestedActive=$exoFallbackActiveCandidateCount; EXOFallbackRequestedInactive=$exoFallbackInactiveCandidateCount; EXOFallbackReturned=$exoFallbackReturnedCount; EXOFallbackPopulated=$exoFilledCount; DeferredUnifiedGroupFallback=$deferredUnifiedGroupMailboxCount; ProfileSkippedFallback=$profileSkippedForFallbackCount; Populated=$finalStatsCount; Missing=$missingMailboxStatsCount; TotalMailboxes=$totalMailboxes." -ExportFileLocation $exportDetails

        # Pre-cache unified group mailbox stats so later unified-group collection can reuse this data.
        if ($shouldPreCacheUnifiedGroupStats) {
            $unifiedPreCacheStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                Write-Host "    Step 3/3 Unified group pre-cache: starting..." -ForegroundColor DarkGray
                Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Pre-caching unified group mailbox stats into PrimaryMailboxStats." -ExportFileLocation $exportDetails
                $groupActivityLookup = Get-Office365GroupsActivityMailboxLookup -Context $Context
                $groupActivityReportRows = if ($groupActivityLookup -and $groupActivityLookup.PSObject.Properties['Rows']) { [int]$groupActivityLookup.Rows } else { 0 }
                $unifiedGroupsForStats = @()
                $groupMailboxCandidateSource = 'AllMailboxes(GroupMailbox)'
                $groupMailboxCandidates = @($tenantStatsHash['AllMailboxes'].Values | Where-Object {
                    $_.PSObject.Properties['RecipientTypeDetails'] -and [string]$_.RecipientTypeDetails -eq 'GroupMailbox'
                })
                if ($groupMailboxCandidates.Count -eq 0) {
                    $groupMailboxCandidateSource = 'UnifiedGroupsInventory'
                    $unifiedGroupsForStats = if ($unifiedGroupsInventoryCache) { @($unifiedGroupsInventoryCache) } else { @() }
                    if ($unifiedGroupsForStats.Count -eq 0) {
                        switch ($detailLevel) {
                            {$_ -in "minimum", "combined", "all"} {
                                $desiredUnifiedGroupProperties = @(
                                    "PrimarySmtpAddress", "DisplayName", "AccessType", "RecipientTypeDetails",
                                    "ExternalDirectoryObjectId",
                                    "ExchangeGuid", @{Name="ManagedByDetails"; Expression={$_.ManagedByDetails -join ','}}, "Notes",
                                    "SharePointSiteUrl", "ContentMailboxName", "GroupMemberCount",
                                    "AllowAddGuests", "WhenSoftDeleted", "HiddenFromExchangeClientsEnabled",
                                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, @{Name="ModeratedBy"; Expression={$_.ModeratedBy -join ','}}, "FolderPath",
                                    @{Name="Description"; Expression={$_.Description -join ','}}, "WhenCreated"
                                )
                                $unifiedGroupsForStats = @(
                                    Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue } |
                                        Select-Object $desiredUnifiedGroupProperties
                                )
                            }
                            default {
                                $unifiedGroupsForStats = @(
                                    Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue }
                                )
                            }
                        }
                    }
                    $groupMailboxCandidates = @($unifiedGroupsForStats)
                }
                $groupMailboxCandidateCount = $groupMailboxCandidates.Count
                $groupsMissingStats = New-Object System.Collections.Generic.List[object]
                $cachedUnifiedGroupStats = 0
                $graphActivityUnifiedGroupStats = 0
                $groupsWithoutGuidForPreCacheCount = 0

                foreach ($groupMailbox in $groupMailboxCandidates) {
                    $groupGuidKey = $null
                    if ($groupMailbox.PSObject.Properties['ExchangeGuid'] -and $groupMailbox.ExchangeGuid) {
                        $groupGuidKey = Convert-ToMailboxGuidKey -GuidValue $groupMailbox.ExchangeGuid
                    }
                    elseif ($groupMailbox.PSObject.Properties['Guid'] -and $groupMailbox.Guid) {
                        $groupGuidKey = Convert-ToMailboxGuidKey -GuidValue $groupMailbox.Guid
                    }

                    if ([string]::IsNullOrWhiteSpace($groupGuidKey)) {
                        $groupsWithoutGuidForPreCacheCount++
                        continue
                    }

                    if ($tenantStatsHash["PrimaryMailboxStats"].ContainsKey($groupGuidKey)) {
                        $cachedUnifiedGroupStats++
                        continue
                    }

                    $groupActivityRow = Get-GroupMailboxActivityRowForUnifiedGroup -GroupRecord $groupMailbox -ActivityLookup $groupActivityLookup
                    if ($groupActivityRow) {
                        $groupActivityStat = New-GroupMailboxStatFromActivityRow -GroupRecord $groupMailbox -ActivityRow $groupActivityRow
                        if ($groupActivityStat) {
                            $tenantStatsHash["PrimaryMailboxStats"][$groupGuidKey] = $groupActivityStat
                            $graphActivityUnifiedGroupStats++
                            continue
                        }
                    }

                    [void]$groupsMissingStats.Add($groupMailbox)
                }

                $fetchedUnifiedGroupStats = 0
                if ($groupsMissingStats.Count -gt 0) {
                    $unifiedGroupStatsResult = Get-ExoMailboxStatisticsSafe -MailboxObjects $groupsMissingStats.ToArray() -Context $Context -ProgressActivity "Pre-caching Unified Group Mailbox Statistics" -ProgressId $primaryStatsProgressId
                    foreach ($groupStat in $unifiedGroupStatsResult.Results) {
                        $key = Convert-ToMailboxGuidKey -GuidValue $groupStat.MailboxGuid
                        if (-not $key) { continue }
                        $tenantStatsHash["PrimaryMailboxStats"][$key] = $groupStat
                    }
                    $fetchedUnifiedGroupStats = @($unifiedGroupStatsResult.Results).Count
                    Write-ExoStatisticsFailureSummary -OperationName 'Get-AllExchangeMailboxDetails unified group mailbox statistics pre-cache' -Context $Context -Failures $unifiedGroupStatsResult.Failures
                }

                # Pre-cache unified group inventory for later collectors.
                if (-not $unifiedGroupsForStats) {
                    $unifiedGroupsForStats = @()
                }
                if ($unifiedGroupsForStats.Count -eq 0) {
                    $unifiedGroupsForStats = if ($unifiedGroupsInventoryCache) { @($unifiedGroupsInventoryCache) } else { @() }
                }
                if ($unifiedGroupsForStats.Count -eq 0) {
                    switch ($detailLevel) {
                        {$_ -in "minimum", "combined", "all"} {
                            $desiredUnifiedGroupProperties = @(
                                "PrimarySmtpAddress", "DisplayName", "AccessType", "RecipientTypeDetails",
                                "ExternalDirectoryObjectId",
                                "ExchangeGuid", @{Name="ManagedByDetails"; Expression={$_.ManagedByDetails -join ','}}, "Notes",
                                "SharePointSiteUrl", "ContentMailboxName", "GroupMemberCount",
                                "AllowAddGuests", "WhenSoftDeleted", "HiddenFromExchangeClientsEnabled",
                                @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, @{Name="ModeratedBy"; Expression={$_.ModeratedBy -join ','}}, "FolderPath",
                                @{Name="Description"; Expression={$_.Description -join ','}}, "WhenCreated"
                            )
                            $unifiedGroupsForStats = @(
                                Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue } |
                                    Select-Object $desiredUnifiedGroupProperties
                            )
                        }
                        default {
                            $unifiedGroupsForStats = @(
                                Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue }
                            )
                        }
                    }
                }
                $Context.Runtime['UnifiedGroupsInventoryCache'] = @($unifiedGroupsForStats)
                $unifiedGroupsInventoryCache = @($Context.Runtime['UnifiedGroupsInventoryCache'])
                $unifiedGroupInventoryCount = $unifiedGroupsForStats.Count

                Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Unified group stats pre-cache summary: source=$groupMailboxCandidateSource; groupMailboxCandidates=$groupMailboxCandidateCount; reused=$cachedUnifiedGroupStats; graphActivityPopulated=$graphActivityUnifiedGroupStats; fetched=$fetchedUnifiedGroupStats; unresolved=$($groupsMissingStats.Count); withoutGuid=$groupsWithoutGuidForPreCacheCount; groupActivityReportRows=$groupActivityReportRows; unifiedGroupInventory=$unifiedGroupInventoryCount." -ExportFileLocation $exportDetails
                if ($groupMailboxCandidateCount -gt 0) {
                    if ($unifiedPreCacheStopwatch -and $unifiedPreCacheStopwatch.IsRunning) {
                        $unifiedPreCacheStopwatch.Stop()
                    }
                    if ($unifiedPreCacheStopwatch) {
                        $unifiedPreCachePhaseSeconds = [math]::Round($unifiedPreCacheStopwatch.Elapsed.TotalSeconds, 2)
                    }
                    Write-Host ("    Step 3/3 Unified group pre-cache: {0}s | source={1}, candidates={2}, reused={3}, graph={4}, fetched={5}, unresolved={6}" -f $unifiedPreCachePhaseSeconds, $groupMailboxCandidateSource, $groupMailboxCandidateCount, $cachedUnifiedGroupStats, $graphActivityUnifiedGroupStats, $fetchedUnifiedGroupStats, $groupsMissingStats.Count) -ForegroundColor DarkGray
                    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 3/3 (unified group pre-cache) completed in $unifiedPreCachePhaseSeconds sec; source=$groupMailboxCandidateSource; groupMailboxCandidates=$groupMailboxCandidateCount; reused=$cachedUnifiedGroupStats; graphActivityPopulated=$graphActivityUnifiedGroupStats; fetched=$fetchedUnifiedGroupStats; unresolved=$($groupsMissingStats.Count); withoutGuid=$groupsWithoutGuidForPreCacheCount; groupActivityReportRows=$groupActivityReportRows; unifiedGroupInventory=$unifiedGroupInventoryCount." -ExportFileLocation $exportDetails
                }
                else {
                    if ($unifiedPreCacheStopwatch -and $unifiedPreCacheStopwatch.IsRunning) {
                        $unifiedPreCacheStopwatch.Stop()
                    }
                    if ($unifiedPreCacheStopwatch) {
                        $unifiedPreCachePhaseSeconds = [math]::Round($unifiedPreCacheStopwatch.Elapsed.TotalSeconds, 2)
                    }
                    Write-Host ("    Step 3/3 Unified group pre-cache: {0}s | no group mailboxes found" -f $unifiedPreCachePhaseSeconds) -ForegroundColor DarkGray
                    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 3/3 (unified group pre-cache) completed in $unifiedPreCachePhaseSeconds sec; no group mailboxes found." -ExportFileLocation $exportDetails
                }
            }
            catch {
                if ($unifiedPreCacheStopwatch -and $unifiedPreCacheStopwatch.IsRunning) {
                    $unifiedPreCacheStopwatch.Stop()
                }
                if ($unifiedPreCacheStopwatch) {
                    $unifiedPreCachePhaseSeconds = [math]::Round($unifiedPreCacheStopwatch.Elapsed.TotalSeconds, 2)
                }
                Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] Unified group mailbox statistics pre-cache failed: $($_.Exception.Message)" -ExportFileLocation $exportDetails
                Write-Host ("    Step 3/3 Unified group pre-cache: failed after {0}s" -f $unifiedPreCachePhaseSeconds) -ForegroundColor Yellow
            }
            finally {
                Write-Progress -Id $primaryStatsProgressId -Activity "Pre-caching Unified Group Mailbox Statistics" -Completed
            }
        }
        else {
            Write-Host "    Step 3/3 Unified group pre-cache: skipped by profile/depth" -ForegroundColor DarkGray
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Unified group mailbox statistics pre-cache is disabled for this profile/depth." -ExportFileLocation $exportDetails
        }
        if ($tenantStatsHash.ContainsKey("PrimaryMailboxStatsCollectionSummary") -and $tenantStatsHash["PrimaryMailboxStatsCollectionSummary"]) {
            $tenantStatsHash["PrimaryMailboxStatsCollectionSummary"] | Add-Member -NotePropertyName 'UnifiedPreCachePhaseSeconds' -NotePropertyValue ([double]$unifiedPreCachePhaseSeconds) -Force
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Mailbox Satistics and adding to Hash Table. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-Progress -Id $primaryStatsProgressId -Activity "Gathering All Primary Mailbox Statistics" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] COMPLETED: Gathering All Primary Mailbox Statistics in $($CompletedTime)" -ExportFileLocation $exportDetails
    }
    
    ## Archive Mailbox Stats to Hash Table
    if ($tenantStatsHash['AllMailboxes'].Values | Where-Object {$_.ArchiveStatus -ne "None"}) {
        try {
            $start = Get-Date
            Write-Host "  Getting archive mailbox stats..." -ForegroundColor Cyan -nonewline
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Archive mailbox size/item metrics are not exposed in Graph mailbox usage reports; using EXO statistics for archive mailboxes." -ExportFileLocation $exportDetails
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering All Archive Mailbox Statistics. Including Group and Inactive Mailboxes" -ExportFileLocation $exportDetails
            Write-Progress -Id $archiveStatsProgressId -Activity "Gathering All Archive Mailbox Statistics" -Status (((Get-Date) - $initialStart).ToString('hh\:mm\:ss'))
            $archiveMailboxCandidates = @($tenantStatsHash['AllMailboxes'].Values | Where-Object {$_.ArchiveStatus -ne "None"})
            $archiveMailboxStatsResult = Get-ExoMailboxStatisticsSafe -MailboxObjects $archiveMailboxCandidates -Context $Context -Archive -ProgressActivity "Gathering All Archive Mailbox Statistics" -ProgressId $archiveStatsProgressId
            $archiveMailboxStats = @($archiveMailboxStatsResult.Results)
            if ($archiveMailboxStats) {
                $tenantStatsHash["ArchiveMailboxStats"] = @{}
                
                #Add to Tenant Stats Hash
                $archiveMailboxStats | ForEach-Object {
                    # Using MailboxGUID from the Mailbox Statistics as the key; matches against the ExchangeGUID from the Mailbox
                    $key = Convert-ToMailboxGuidKey -GuidValue $_.MailboxGuid
                    if (-not $key) { continue }
                    $value = $_
                    $tenantStatsHash["ArchiveMailboxStats"][$key] = $value
                }
            }
            Write-ExoStatisticsFailureSummary -OperationName 'Get-AllExchangeMailboxDetails archive mailbox statistics' -Context $Context -Failures $archiveMailboxStatsResult.Failures
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Adding Archive Mailbox Statistics to Tenant Stats Hash." -ExportFileLocation $exportDetails
        }
        catch {
            if ($_.Exception.Message -like "You cannot call a method on a null-valued expression") {
                Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Archive Mailbox Satistics and adding to Hash Table. There are no Archive mailboxes found. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
            }
            else {
                Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Archive Mailbox Satistics and adding to Hash Table. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
            }
        }
        finally { 
            Write-Progress -Id $archiveStatsProgressId -Activity "Gathering All Archive Mailbox Statistics" -Completed
            $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
            Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails]COMPLETED: Gathering All Archive Mailbox Statistics in $($CompletedTime)" -ExportFileLocation $exportDetails
        }
    } else {
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] No Archive Mailboxes found." -ExportFileLocation $exportDetails
        Write-Host "  No Archive Mailboxes found." -ForegroundColor Yellow
    }
}
