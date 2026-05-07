function Get-EntraIDGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Provide the level of detail')]
        [ValidateSet('minimum', 'operator', 'combined', 'automation', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory = $false, HelpMessage = 'Provide the Graph Authentication Type')]
        [ValidateSet('SDK', 'REST')]
        [string[]]$GraphAuthType = @('REST'),
        [Parameter(Mandatory = $false)]
        $Context
    )

    $groupFetchProgressId = 61
    $groupLoopProgressId = 62
    $groupDetailProgressId = 63
    $Context = Resolve-ArrayaGraphCollectorContext -Context $Context -DetailLevel $detailLevel
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $depthPolicy = $Context.Policies['CollectionDepth']
    $collectDeepGroupDetails = ($depthPolicy.CollectEntraGroupDeepDetails -eq $true)
    $collectGroupLicenseChecks = ($depthPolicy.CollectEntraGroupLicenseChecks -eq $true)
    $collectGroupMemberCounts = ($depthPolicy.CollectEntraGroupMemberCounts -eq $true)
    $collectGroupOwnerCounts = ($depthPolicy.CollectEntraGroupOwnerCounts -eq $true)
    $groupMemberCountLookup = @{}
    $groupOwnerCountLookup = @{}
    $licenseSkuLookupById = @{}
    $tenantStatsHash['EntraIDGroups'] = @{}
    $groupSelectProperties = @(
        'id',
        'displayName',
        'description',
        'visibility',
        'createdDateTime',
        'groupTypes',
        'mailEnabled',
        'securityEnabled',
        'onPremisesSyncEnabled',
        'onPremisesLastSyncDateTime',
        'isAssignableToRole',
        'mail',
        'membershipRule',
        'assignedLicenses'
    )

    if ($tenantStatsHash -and $tenantStatsHash.ContainsKey('LicenseSKUs') -and $tenantStatsHash['LicenseSKUs']) {
        $licenseSkuRows = if ($tenantStatsHash['LicenseSKUs'] -is [System.Collections.IDictionary]) {
            @($tenantStatsHash['LicenseSKUs'].Values)
        }
        else {
            @($tenantStatsHash['LicenseSKUs'])
        }
        foreach ($licenseSkuRow in @($licenseSkuRows)) {
            if ($null -eq $licenseSkuRow) {
                continue
            }

            $skuId = [string](Get-ArrayaObjectValue -Object $licenseSkuRow -Names @('SkuId', 'skuId', 'Id'))
            if ([string]::IsNullOrWhiteSpace($skuId)) {
                continue
            }

            $licenseSkuLookupById[$skuId] = [pscustomobject]@{
                SkuId         = $skuId
                SkuPartNumber = [string](Get-ArrayaObjectValue -Object $licenseSkuRow -Names @('SkuPartNumber', 'skuPartNumber'))
                FriendlyName  = [string](Get-ArrayaObjectValue -Object $licenseSkuRow -Names @('FriendlyName', 'Name', 'SkuPartNumber', 'skuPartNumber'))
            }
        }
    }

    function Get-EntraGroupCountLookups {
        param(
            [Parameter(Mandatory = $true)]
            [array]$Groups
        )

        $lookups = @{
            Members = @{}
            Owners  = @{}
        }

        if (-not $Groups -or $Groups.Count -eq 0) {
            return $lookups
        }
        if (-not $collectGroupMemberCounts -and -not $collectGroupOwnerCounts) {
            return $lookups
        }

        $requests = New-Object System.Collections.Generic.List[object]
        foreach ($group in $Groups) {
            $groupId = [string](Get-ArrayaObjectValue -Object $group -Names @('id'))
            if ([string]::IsNullOrWhiteSpace($groupId)) {
                continue
            }

            $isDynamicDistributionGroup = ($group.groupTypes -contains 'DynamicMembership') -and ($group.mailEnabled -eq $true) -and ($group.securityEnabled -eq $false) -and (-not ($group.groupTypes -contains 'Unified'))
            if ($isDynamicDistributionGroup) {
                continue
            }

            if ($collectGroupMemberCounts) {
                $requests.Add([PSCustomObject]@{
                    id      = "m:$groupId"
                    method  = 'GET'
                    url     = "/groups/$groupId/members/`$count"
                    headers = @{ ConsistencyLevel = 'eventual' }
                }) | Out-Null
            }

            if ($collectGroupOwnerCounts) {
                $requests.Add([PSCustomObject]@{
                    id      = "o:$groupId"
                    method  = 'GET'
                    url     = "/groups/$groupId/owners/`$count"
                    headers = @{ ConsistencyLevel = 'eventual' }
                }) | Out-Null
            }
        }

        if ($requests.Count -eq 0) {
            return $lookups
        }

        Write-Host '    > Entra groups: member and owner count prefetch' -ForegroundColor DarkCyan
        Write-Log -Type INFO -Message "[Get-EntraIDGroups] Prefetching group member/owner counts via Graph batch for $($Groups.Count) groups ($($requests.Count) count request(s))." -ExportFileLocation $exportDetails
        $responses = Invoke-ArrayaGraphBatchRequest -Requests $requests.ToArray() -GraphAuthType $GraphAuthType -Activity 'Group member/owner count prefetch' -ExportFileLocation $exportDetails
        foreach ($response in @($responses.Values)) {
            if (-not $response -or -not $response.id) {
                continue
            }

            $responseId = [string]$response.id
            $statusCode = 0
            try {
                $statusCode = [int]$response.status
            }
            catch {
                $statusCode = 0
            }
            if ($statusCode -lt 200 -or $statusCode -ge 300) {
                continue
            }

            $countValue = Convert-ArrayaGraphBatchCountValue -Body $response.body
            if ($null -eq $countValue) {
                continue
            }

            if ($responseId.StartsWith('m:')) {
                $lookups.Members[$responseId.Substring(2)] = [int]$countValue
            }
            elseif ($responseId.StartsWith('o:')) {
                $lookups.Owners[$responseId.Substring(2)] = [int]$countValue
            }
        }

        Write-Log -Type INFO -Message "[Get-EntraIDGroups] Group count prefetch complete. MemberCounts=$($lookups.Members.Count) OwnerCounts=$($lookups.Owners.Count)" -ExportFileLocation $exportDetails
        return $lookups
    }

    function Get-EntraGroupDetails {
        param(
            [Parameter(Mandatory = $true)]
            $Group
        )

        $groupId = [string](Get-ArrayaObjectValue -Object $Group -Names @('id'))
        if ([string]::IsNullOrWhiteSpace($groupId)) {
            return $null
        }

        $groupDetails = $Group
        if ($collectDeepGroupDetails) {
            $groupDetails = Office365Custom\Get-GraphData -PageSize 999 -Uri "https://graph.microsoft.com/v1.0/groups/$groupId" -Id $groupDetailProgressId -Activity 'Gathering Group Details'
            if ($groupDetails -is [array]) {
                $groupDetails = @($groupDetails | Select-Object -First 1)
            }
        }

        if ($null -eq $groupDetails) {
            return $null
        }

        $classification = Get-ArrayaEntraGroupClassification -GroupDetails $groupDetails
        $isDynamicDistributionGroup = ($groupDetails.groupTypes -contains 'DynamicMembership') -and ($groupDetails.mailEnabled -eq $true) -and ($groupDetails.securityEnabled -eq $false) -and (-not ($groupDetails.groupTypes -contains 'Unified'))
        $assignedLicenses = @((Get-ArrayaObjectValue -Object $groupDetails -Names @('assignedLicenses')))
        $assignedLicenseSkuIds = New-Object 'System.Collections.Generic.List[string]'
        $assignedLicenseSkuPartNumbers = New-Object 'System.Collections.Generic.List[string]'
        $assignedLicenseFriendlyNames = New-Object 'System.Collections.Generic.List[string]'

        if ($collectGroupLicenseChecks) {
            foreach ($assignedLicense in @($assignedLicenses)) {
                if ($null -eq $assignedLicense) {
                    continue
                }

                $skuId = [string](Get-ArrayaObjectValue -Object $assignedLicense -Names @('skuId', 'SkuId'))
                if ([string]::IsNullOrWhiteSpace($skuId)) {
                    continue
                }

                $assignedLicenseSkuIds.Add($skuId) | Out-Null
                if ($licenseSkuLookupById.ContainsKey($skuId)) {
                    $skuLookup = $licenseSkuLookupById[$skuId]
                    if (-not [string]::IsNullOrWhiteSpace([string]$skuLookup.SkuPartNumber)) {
                        $assignedLicenseSkuPartNumbers.Add([string]$skuLookup.SkuPartNumber) | Out-Null
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$skuLookup.FriendlyName)) {
                        $assignedLicenseFriendlyNames.Add([string]$skuLookup.FriendlyName) | Out-Null
                    }
                }
                else {
                    $assignedLicenseSkuPartNumbers.Add($skuId) | Out-Null
                    $assignedLicenseFriendlyNames.Add($skuId) | Out-Null
                }
            }
        }

        $isManagingLicenses = if ($collectGroupLicenseChecks) {
            $assignedLicenses.Count -gt 0
        }
        else {
            'NotCollected (minimum mode)'
        }

        $memberCount = if ($isDynamicDistributionGroup) {
            'Skipped (Dynamic Distribution Group)'
        }
        elseif (-not $collectGroupMemberCounts) {
            'NotCollected (minimum mode)'
        }
        elseif ($groupMemberCountLookup.ContainsKey($groupId)) {
            [int]$groupMemberCountLookup[$groupId]
        }
        else {
            $memberResult = Office365Custom\Get-GraphData -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members/`$count" -Id $groupDetailProgressId -Activity 'Counting Members'
            if ($memberResult -is [array]) { [int]($memberResult | Select-Object -First 1) } else { [int]$memberResult }
        }

        $ownerCount = if ($isDynamicDistributionGroup) {
            'Skipped (Dynamic Distribution Group)'
        }
        elseif (-not $collectGroupOwnerCounts) {
            'NotCollected (minimum mode)'
        }
        elseif ($groupOwnerCountLookup.ContainsKey($groupId)) {
            [int]$groupOwnerCountLookup[$groupId]
        }
        else {
            $ownerResult = Office365Custom\Get-GraphData -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/owners/`$count" -Id $groupDetailProgressId -Activity 'Counting Owners'
            if ($ownerResult -is [array]) { [int]($ownerResult | Select-Object -First 1) } else { [int]$ownerResult }
        }

        return [PSCustomObject]@{
            ID                         = $groupDetails.id
            DisplayName                = $groupDetails.displayName
            Description                = $groupDetails.description
            Source                     = $classification.Source
            Visibility                 = $groupDetails.visibility
            CreatedDateTime            = $groupDetails.createdDateTime
            GroupType                  = $classification.GroupType
            MembershipType             = $classification.MembershipType
            MembershipRule             = $classification.MembershipRule
            HasNestedMembers           = $false
            OnPremisesSyncEnabled      = $groupDetails.onPremisesSyncEnabled
            OnPremisesLastSyncDateTime = $groupDetails.onPremisesLastSyncDateTime
            IsManagingLicenses         = $isManagingLicenses
            AssignedLicenseCount       = if ($collectGroupLicenseChecks) { $assignedLicenseSkuIds.Count } else { 'NotCollected (minimum mode)' }
            AssignedLicenseSkuIds      = if ($collectGroupLicenseChecks) { ($assignedLicenseSkuIds.ToArray() -join ',') } else { 'NotCollected (minimum mode)' }
            AssignedLicenseSkuPartNumbers = if ($collectGroupLicenseChecks) { ($assignedLicenseSkuPartNumbers.ToArray() | Select-Object -Unique) -join ',' } else { 'NotCollected (minimum mode)' }
            AssignedLicenseFriendlyNames = if ($collectGroupLicenseChecks) { ($assignedLicenseFriendlyNames.ToArray() | Select-Object -Unique) -join ',' } else { 'NotCollected (minimum mode)' }
            IsAssignableToRole         = $groupDetails.isAssignableToRole
            Mail                       = $groupDetails.mail
            MailEnabled                = $groupDetails.mailEnabled
            SecurityEnabled            = $groupDetails.securityEnabled
            MemberCount                = $memberCount
            OwnerCount                 = $ownerCount
        }
    }

    try {
        $groupsEndpoint = "https://graph.microsoft.com/v1.0/groups?`$select=$($groupSelectProperties -join ',')"
        Write-Host '    > Entra groups: inventory retrieval' -ForegroundColor DarkCyan
        Write-Log -Type INFO -Message 'Fetching initial Entra Groups' -ExportFileLocation $exportDetails
        $groups = @(Office365Custom\Get-GraphData -PageSize 999 -Uri $groupsEndpoint -Id $groupFetchProgressId -Activity 'Gathering Group Details')
        if ($collectDeepGroupDetails -and ($collectGroupMemberCounts -or $collectGroupOwnerCounts) -and $groups.Count -gt 0) {
            $countLookups = Get-EntraGroupCountLookups -Groups $groups
            $groupMemberCountLookup = $countLookups.Members
            $groupOwnerCountLookup = $countLookups.Owners
        }

        $totalGroups = $groups.Count
        Write-Host '    > Entra groups: per-group detail enrichment' -ForegroundColor DarkCyan
        foreach ($group in $groups) {
            Write-ProgressHelper -Total ([Math]::Max($totalGroups, 1)) -Id $groupLoopProgressId -Activity 'Getting Group Details' -Operation "Processing Group: $($group.displayName)"
            Write-Log -Type INFO -Message "Checking group $($group.displayName)" -ExportFileLocation $exportDetails
            $groupDetails = Invoke-ArrayaCollectionStepSafe -OperationName "Get-EntraIDGroups details for $($group.displayName)" -DefaultValue $null -ExportFileLocation $exportDetails -ScriptBlock {
                Get-EntraGroupDetails -Group $group
            }

            if ($groupDetails) {
                $tenantStatsHash['EntraIDGroups'][$groupDetails.ID] = $groupDetails
            }
            else {
                Write-Log -Type WARNING -Message "[Get-EntraIDGroups] Could not retrieve details for group $($group.displayName). Continuing." -ExportFileLocation $exportDetails
            }
        }
    }
    catch {
        Write-Log -Type ERROR -Message "Error fetching initial Entra Groups: $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
        return @()
    }
    finally {
        Write-Progress -Activity 'Getting Group Details' -Completed -Id $groupFetchProgressId
        Write-ProgressHelper -Total 1 -Id $groupLoopProgressId -Activity 'Getting Group Details' -Completed
    }

    return $tenantStatsHash['EntraIDGroups']
}
