# Exchange Group Details
function Get-ExchangeGroupDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'operator', 'combined', 'automation', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory = $false)]
        $Context
    )

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context -DetailLevel $detailLevel
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $exchangeGroupProgressId = 36
    $exchangeGroupProgressTotal = 1
    $depthPolicy = $Context.Policies['CollectionDepth']
    $isCombinedDepth = [bool]($depthPolicy.PSObject.Properties['IsCombined'] -and $depthPolicy.IsCombined)
    $collectExchangeGroupMetadataDetails = ($depthPolicy.IsAll -or $depthPolicy.IsGeek)
    $collectExchangeGroupMemberExpansion = ($depthPolicy.IsAll -or $depthPolicy.IsGeek)
    # Member rows are kept only where membership was actually expanded, so a group that was
    # not expanded is reported as not collected rather than as empty.
    $groupMemberRows = New-Object System.Collections.Generic.List[object]
    try {
        $start = Get-Date
        $tenantStatsHash["AllExchangeGroups"] = @{}
        $cachedUnifiedGroupsByAddress = @{}
        if (
            $tenantStatsHash.ContainsKey('UnifiedGroups') -and
            $tenantStatsHash['UnifiedGroups'] -is [System.Collections.IDictionary]
        ) {
            foreach ($entry in $tenantStatsHash['UnifiedGroups'].GetEnumerator()) {
                $smtp = [string]$entry.Key
                if ([string]::IsNullOrWhiteSpace($smtp)) {
                    continue
                }
                $cachedUnifiedGroupsByAddress[$smtp.ToLowerInvariant()] = $entry.Value
            }
        }
        if (
            ([System.Collections.IDictionary]$Context.Runtime).Contains('UnifiedGroupsInventoryCache') -and
            $Context.Runtime['UnifiedGroupsInventoryCache']
        ) {
            foreach ($group in @($Context.Runtime['UnifiedGroupsInventoryCache'])) {
                if (-not $group -or -not $group.PSObject.Properties['PrimarySmtpAddress']) {
                    continue
                }
                $smtp = [string]$group.PrimarySmtpAddress
                if ([string]::IsNullOrWhiteSpace($smtp)) {
                    continue
                }
                $lookupKey = $smtp.ToLowerInvariant()
                if (-not $cachedUnifiedGroupsByAddress.ContainsKey($lookupKey)) {
                    $cachedUnifiedGroupsByAddress[$lookupKey] = $group
                }
            }
        }
        $metadataReusedFromRecipientCount = 0
        $metadataReusedFromUnifiedCacheCount = 0
        $metadataFetchedFromExoCount = 0
        
        Write-ArrayaExchangeCollectorBanner -Message '[Get-ExchangeGroupDetails] START: Gathering all Exchange Online Groups' -ExportFileLocation $exportDetails
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] START: Gathering all Exchange Online Groups with $($detailLevel) details" -ExportFileLocation $exportDetails
        if (-not $collectExchangeGroupMetadataDetails) {
            if ($depthPolicy.IsMinimum) {
                Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Minimum mode optimization active. Reusing recipient inventory and skipping EXO group metadata/member expansion." -ExportFileLocation $exportDetails
            }
            elseif ($isCombinedDepth) {
                Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Combined mode optimization active. Reusing recipient inventory and unified-group cache for metadata; skipping per-group EXO metadata/member expansion." -ExportFileLocation $exportDetails
            }
        }
        elseif (-not $collectExchangeGroupMemberExpansion) {
            Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Combined mode optimization active. Gathering group metadata while skipping deep member expansion." -ExportFileLocation $exportDetails
        }

        # gather All Exchange Online Groups
        $allMailGroups = $tenantStatsHash['AllRecipients'].Values | Where-Object { $_.RecipientTypeDetails -like "*group" } | Sort-Object DisplayName
        Write-ArrayaExchangeCollectorSubstep -Message 'Exchange groups: inventory from Exchange recipient data'
        if ($collectExchangeGroupMetadataDetails -or $collectExchangeGroupMemberExpansion) {
            Write-ArrayaExchangeCollectorSubstep -Message 'Exchange groups: metadata and membership enrichment'
        }
        
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Gathering all Exchange Online Groups Details" -ExportFileLocation $exportDetails
        $totalCount = $allMailGroups.count
        $exchangeGroupProgressTotal = [Math]::Max($totalCount, 1)
        foreach ($object in $allMailGroups) {
            try {
                $identity = $object.identity.tostring()
                $PrimarySMTPAddress = if ($object.PrimarySMTPAddress) { $object.PrimarySMTPAddress.ToString() } else { $identity }
                $primarySmtpLookupKey = if ([string]::IsNullOrWhiteSpace($PrimarySMTPAddress)) { $null } else { $PrimarySMTPAddress.ToLowerInvariant() }
                Write-ProgressHelper -Total $exchangeGroupProgressTotal -Id $exchangeGroupProgressId -Activity "Gathering All Exchange Online Group Details" -Operation "Gathering Group Details for $($PrimarySMTPAddress)"
                Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] Gathering '{0}' '{1}' Group Details" -f $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $exportDetails

                # Clear details
                $attributesToClear = @('groupDetails', 'groupOwners','groupMembers')
                foreach ($attribute in $attributesToClear) {
                    Set-Variable -Name $attribute -Value @()
                }
                $groupMembersCount = 0
                $cachedUnifiedGroup = $null
                $usedUnifiedGroupCache = $false
                $membersExpanded = $false

                if (
                    $primarySmtpLookupKey -and
                    $cachedUnifiedGroupsByAddress.ContainsKey($primarySmtpLookupKey)
                ) {
                    $cachedUnifiedGroup = $cachedUnifiedGroupsByAddress[$primarySmtpLookupKey]
                }

                # Conditional logic for different recipient types
                switch ($object.RecipientTypeDetails) {
                    "DynamicDistributionGroup" {
                        if ($collectExchangeGroupMetadataDetails) {
                            $groupDetails = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails details for $PrimarySMTPAddress" -DefaultValue $null -ExportFileLocation $exportDetails -ScriptBlock {
                                Get-DynamicDistributionGroup $identity -ErrorAction Stop
                            }
                            if ($groupDetails) {
                                $metadataFetchedFromExoCount++
                            }
                        }
                        else {
                            $groupDetails = $object
                            $metadataReusedFromRecipientCount++
                        }

                        if ($collectExchangeGroupMemberExpansion) {
                            # Wrapped so an empty group is distinguishable from a failed lookup;
                            # a bare empty array unrolls to $null on return.
                            $memberResult = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails members for $PrimarySMTPAddress" -DefaultValue $null -ExportFileLocation $exportDetails -ScriptBlock {
                                [pscustomobject]@{ Members = @(Get-DynamicDistributionGroupMember $identity -ErrorAction Stop -ResultSize unlimited -WarningAction SilentlyContinue) }
                            }
                            $membersExpanded = ($null -ne $memberResult)
                            $groupMembers = if ($membersExpanded) { @($memberResult.Members | Where-Object { $null -ne $_ }) } else { @() }
                        }
                        else {
                            $groupMembers = @()
                        }
                    }
                    {$_ -in 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', "MailNonUniversalGroup"} {
                        if ($collectExchangeGroupMetadataDetails) {
                            $groupDetails = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails details for $PrimarySMTPAddress" -DefaultValue $null -ExportFileLocation $exportDetails -ScriptBlock {
                                Get-DistributionGroup $identity -ErrorAction Stop
                            }
                            if ($groupDetails) {
                                $metadataFetchedFromExoCount++
                            }
                        }
                        else {
                            $groupDetails = $object
                            $metadataReusedFromRecipientCount++
                        }

                        if ($collectExchangeGroupMemberExpansion) {
                            $memberResult = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails members for $PrimarySMTPAddress" -DefaultValue $null -ExportFileLocation $exportDetails -ScriptBlock {
                                [pscustomobject]@{ Members = @(Get-DistributionGroupMember $identity -ResultSize unlimited -ErrorAction Stop) }
                            }
                            $membersExpanded = ($null -ne $memberResult)
                            $groupMembers = if ($membersExpanded) { @($memberResult.Members | Where-Object { $null -ne $_ }) } else { @() }
                        }
                        else {
                            $groupMembers = @()
                        }
                    }
                    "GroupMailbox" {
                        if ($cachedUnifiedGroup) {
                            $groupDetails = $cachedUnifiedGroup
                            $usedUnifiedGroupCache = $true
                            $metadataReusedFromUnifiedCacheCount++
                            if ($cachedUnifiedGroup.PSObject.Properties['GroupMemberCount'] -and $cachedUnifiedGroup.GroupMemberCount -ne $null -and $cachedUnifiedGroup.GroupMemberCount -ne '') {
                                try { $groupMembersCount = [int]$cachedUnifiedGroup.GroupMemberCount } catch { $groupMembersCount = 0 }
                            }
                            $groupMembers = @()
                        }

                        if ($collectExchangeGroupMetadataDetails -and -not $usedUnifiedGroupCache) {
                            $groupDetails = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails details for $PrimarySMTPAddress" -DefaultValue $null -ExportFileLocation $exportDetails -ScriptBlock {
                                Get-UnifiedGroup $identity -ErrorAction Stop
                            }
                            if ($groupDetails) {
                                $metadataFetchedFromExoCount++
                            }
                        }
                        elseif (-not $groupDetails) {
                            $groupDetails = $object
                            $metadataReusedFromRecipientCount++
                        }

                        if ($collectExchangeGroupMemberExpansion -and -not $usedUnifiedGroupCache) {
                            $memberResult = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails members for $PrimarySMTPAddress" -DefaultValue $null -ExportFileLocation $exportDetails -ScriptBlock {
                                [pscustomobject]@{ Members = @(Get-UnifiedGroupLinks -Identity $identity -LinkType Member -ResultSize unlimited -ErrorAction Stop) }
                            }
                            $membersExpanded = ($null -ne $memberResult)
                            $groupMembers = if ($membersExpanded) { @($memberResult.Members | Where-Object { $null -ne $_ }) } else { @() }
                        }
                        else {
                            $groupMembers = @()
                        }
                    }
                }

                if (-not $groupDetails) {
                    $groupDetails = $object
                    $metadataReusedFromRecipientCount++
                }

                # Owners. The recipient inventory joins ManagedBy into one comma-separated
                # string, so a string must be split before counting; counting it directly
                # reported one owner for any group with owners.
                $managedByValue = $object.ManagedBy
                if ($usedUnifiedGroupCache -and $cachedUnifiedGroup.PSObject.Properties['ManagedByDetails'] -and -not [string]::IsNullOrWhiteSpace([string]$cachedUnifiedGroup.ManagedByDetails)) {
                    $managedByValue = $cachedUnifiedGroup.ManagedByDetails
                }
                $groupOwners = @(
                    foreach ($ownerEntry in @($managedByValue)) {
                        if ($null -eq $ownerEntry) { continue }
                        foreach ($ownerName in @(([string]$ownerEntry) -split ',')) {
                            $trimmedOwner = $ownerName.Trim()
                            if (-not [string]::IsNullOrWhiteSpace($trimmedOwner)) { $trimmedOwner }
                        }
                    }
                )
                if ($groupOwners.Count -ge 1) {
                    Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] '{0}' Owners Found for '{1}' '{2}'" -f $groupOwners.Count, $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $exportDetails
                }

                $memberInventoryStatus = if ($membersExpanded) {
                    'Collected'
                }
                elseif ($groupMembersCount -gt 0) {
                    'CountOnly'
                }
                elseif ($collectExchangeGroupMemberExpansion -and -not $usedUnifiedGroupCache) {
                    'LookupFailed'
                }
                else {
                    'NotCollected'
                }

                if ($membersExpanded) {
                    $membershipSource = if ([string]$object.RecipientTypeDetails -eq 'DynamicDistributionGroup') { 'Computed from recipient filter' } else { 'Assigned' }
                    foreach ($member in @($groupMembers)) {
                        $groupMemberRows.Add([pscustomobject]@{
                            GroupDisplayName           = $object.DisplayName
                            GroupPrimarySmtpAddress    = [string]$object.PrimarySMTPAddress
                            GroupRecipientTypeDetails  = [string]$object.RecipientTypeDetails
                            MemberDisplayName          = [string](Get-ArrayaObjectValue -Object $member -Names @('DisplayName', 'Name'))
                            MemberPrimarySmtpAddress   = [string](Get-ArrayaObjectValue -Object $member -Names @('PrimarySmtpAddress', 'WindowsLiveID', 'ExternalEmailAddress'))
                            MemberRecipientTypeDetails = [string](Get-ArrayaObjectValue -Object $member -Names @('RecipientTypeDetails', 'RecipientType'))
                            MembershipSource           = $membershipSource
                        }) | Out-Null
                    }
                }
                #Check Group Members Size and Get Group Addresses
                if ($groupMembers.count -ge 1) {
                    Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] '{0}' Members Found for '{1}' '{2}'" -f $groupMembers.count, $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $exportDetails
                }
                if ($groupMembersCount -le 0) {
                    $groupMembersCount = ($groupMembers | Measure-Object).Count
                }
                Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] Create Group Output Details for '{0}' '{1}'" -f $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $exportDetails

                #Output Group Details
                $currentobject = [PSCustomObject]@{
                    DisplayName                              = $object.DisplayName
                    Identity                                 = $identity
                    ExternalDirectoryObjectId                = $object.ExternalDirectoryObjectId
                    Alias                                    = $object.alias
                    Notes                                    = $object.Notes
                    IsDirSynced                              = $groupDetails.IsDirSynced
                    HiddenFromAddressListsEnabled            = $object.HiddenFromAddressListsEnabled
                    PrimarySMTPAddress                       = $object.PrimarySMTPAddress
                    RecipientTypeDetails                     = $object.RecipientTypeDetails
                    ResourceProvisioningOptions              = ($groupDetails.ResourceProvisioningOptions -join ",")
                    IsMailboxConfigured                      = $groupDetails.IsMailboxConfigured
                    EmailAddresses                           = $object.EmailAddresses
                    Owners                                   = ($groupOwners -join '; ')
                    OwnersCount                              = $groupOwners.Count
                    MembersCount                             = $groupMembersCount
                    MemberInventoryStatus                    = $memberInventoryStatus
                    RecipientFilter                          = $groupDetails.RecipientFilter
                    MemberJoinRestriction                    = $groupDetails.MemberJoinRestriction
                    MemberDepartRestriction                  = $groupDetails.MemberDepartRestriction
                    GrantSendOnBehalfTo                      = ($groupDetails.GrantSendOnBehalfTo -join ",")
                    HiddenGroupMembershipEnabled             = ($groupDetails.HiddenGroupMembershipEnabled -join ",")
                    ModerationEnabled                        = $groupDetails.ModerationEnabled
                    ModeratedBy                              = ($groupDetails.ModeratedBy -join ",")
                    RequireSenderAuthenticationEnabled       = $groupDetails.RequireSenderAuthenticationEnabled
                    AcceptMessagesOnlyFrom                   = ($groupDetails.AcceptMessagesOnlyFrom -join ",")
                    AcceptMessagesOnlyFromDLMembers          = ($groupDetails.AcceptMessagesOnlyFromDLMembers -join ",")
                    AcceptMessagesOnlyFromSendersOrMembers   = ($groupDetails.AcceptMessagesOnlyFromSendersOrMembers -join ",")
                    RejectMessagesFrom                       = ($groupDetails.RejectMessagesFrom -join ",")
                    RejectMessagesFromDLMembers              = ($groupDetails.RejectMessagesFromDLMembers -join ",")
                    RejectMessagesFromSendersOrMembers       = ($groupDetails.RejectMessagesFromSendersOrMembers -join ",")
                    AccessType                               = $groupDetails.AccessType
                    AllowAddGuests                           = $groupDetails.AllowAddGuests
                    SharePointSiteUrl                        = $groupDetails.SharePointSiteUrl
                }

                Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] Add '{0}' '{1}' Group Details to Tenant Stats Hash" -f $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $exportDetails
                $tenantStatsHash["AllExchangeGroups"][$object.identity] = $currentobject
            }
            catch {
                Write-Log -Type ERROR -Message "[Get-ExchangeGroupDetails] An error occurred in running Get-ExchangeGroupDetails function. Exception: $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
            }
        }
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Group metadata source breakdown: RecipientCache=$metadataReusedFromRecipientCount; UnifiedCache=$metadataReusedFromUnifiedCacheCount; EXOMetadataCalls=$metadataFetchedFromExoCount; TotalGroups=$totalCount." -ExportFileLocation $exportDetails
        Write-Host ("  Group metadata source breakdown: Recipient cache={0}, Unified cache={1}, EXO metadata calls={2}" -f $metadataReusedFromRecipientCount, $metadataReusedFromUnifiedCacheCount, $metadataFetchedFromExoCount) -ForegroundColor DarkGray
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-ExchangeGroupDetails] An error occurred in running Get-ExchangeGroupDetails function. Exception: $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        $tenantStatsHash['ExchangeGroupMembers'] = @($groupMemberRows.ToArray())
        Write-ProgressHelper -Total $exchangeGroupProgressTotal -Id $exchangeGroupProgressId -Activity "Gathering All Exchange Online Group Details" -Completed
    }
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-ArrayaExchangeCollectorCompletionBanner -Message "[Get-ExchangeGroupDetails] COMPLETED: Gathering all Exchange Online Groups in $($CompletedTime)" -ExportFileLocation $exportDetails
    Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] COMPLETED: Gathering all Exchange Online Groups in $($CompletedTime)" -ExportFileLocation $exportDetails
}
