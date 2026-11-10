# License inventory, SKU naming, group-based licensing, and license optimization analysis.
# Shared by the HTML assessment report and the tenant improvement plan so both
# deliverables describe the same tenant from one implementation.

function Get-ArrayaLicenseSkuLookup {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object[]]$Licenses = @())

    $lookup = @{}
    foreach ($license in @($Licenses)) {
        if ($null -eq $license) { continue }
        $skuId = [string](Get-ArrayaObjectValue -Object $license -Names @('SkuId', 'skuId', 'Id'))
        $skuPartNumber = [string](Get-ArrayaObjectValue -Object $license -Names @('SkuPartNumber', 'skuPartNumber', 'AccountSkuId'))
        $friendlyName = [string](Get-ArrayaObjectValue -Object $license -Names @('FriendlyName', 'SkuFriendlyName', 'SkuDisplayName', 'ProductName', 'DisplayName', 'Name', 'SkuPartNumber', 'skuPartNumber'))
        $skuRecord = [pscustomobject]@{
            SkuId         = $skuId
            SkuPartNumber = $skuPartNumber
            FriendlyName  = if ([string]::IsNullOrWhiteSpace($friendlyName)) { $skuPartNumber } else { $friendlyName }
        }
        foreach ($key in @($skuId, $skuPartNumber, $friendlyName)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$key)) {
                $lookup[[string]$key] = $skuRecord
            }
        }
    }

    return $lookup
}

function Resolve-ArrayaLicenseSkuName {
    [CmdletBinding()]
    param(
        [AllowNull()]$SkuValue,
        [Parameter(Mandatory = $false)][hashtable]$SkuLookup = @{}
    )

    $skuText = ([string]$SkuValue).Trim()
    if ([string]::IsNullOrWhiteSpace($skuText)) { return $null }
    if ($SkuLookup.ContainsKey($skuText)) {
        $skuRecord = $SkuLookup[$skuText]
        if (-not [string]::IsNullOrWhiteSpace([string]$skuRecord.FriendlyName)) { return [string]$skuRecord.FriendlyName }
        if (-not [string]::IsNullOrWhiteSpace([string]$skuRecord.SkuPartNumber)) { return [string]$skuRecord.SkuPartNumber }
    }

    return $skuText
}

function Get-ArrayaLicenseSuiteFamily {
    [CmdletBinding()]
    param([AllowNull()][string]$SkuName)

    if ([string]::IsNullOrWhiteSpace($SkuName)) { return $null }
    $normalized = $SkuName.ToUpperInvariant()
    if ($normalized -match 'SPE_E5|MICROSOFT 365 E5|M365 E5|ENTERPRISEPREMIUM|OFFICE 365 E5|O365 E5|SPE_E3|MICROSOFT 365 E3|M365 E3|ENTERPRISEPACK|OFFICE 365 E3|O365 E3|SPE_F1|MICROSOFT 365 F3|M365 F3|DESKLESSPACK|OFFICE 365 F3|O365 F3') {
        return 'Microsoft 365 / Office 365 suite'
    }
    if ($normalized -match 'EMS|ENTERPRISE MOBILITY') {
        return 'Enterprise Mobility + Security'
    }

    return $null
}

function Get-ArrayaUserLicenseAssignmentRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$User,
        [Parameter(Mandatory = $false)][hashtable]$SkuLookup = @{}
    )

    $stateValue = Get-ArrayaObjectValue -Object $User -Names @('LicenseAssignmentStates', 'licenseAssignmentStates')
    $stateRows = @()
    if ($stateValue) {
        if (
            $stateValue -is [System.Collections.IDictionary] -and
            ([System.Collections.IDictionary]$stateValue).Contains('SkuId')
        ) {
            $stateRows = @($stateValue)
        }
        else {
            $stateRows = @(Convert-ArrayaObjectToArray $stateValue)
        }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($stateRow in @($stateRows)) {
        if ($null -eq $stateRow -or $stateRow -is [string]) { continue }
        $skuId = [string](Get-ArrayaObjectValue -Object $stateRow -Names @('SkuId', 'skuId'))
        $skuName = [string](Get-ArrayaObjectValue -Object $stateRow -Names @('SkuName', 'SkuPartNumber', 'FriendlyName'))
        if ([string]::IsNullOrWhiteSpace($skuName)) {
            $skuName = Resolve-ArrayaLicenseSkuName -SkuValue $skuId -SkuLookup $SkuLookup
        }
        $assignedByGroup = [string](Get-ArrayaObjectValue -Object $stateRow -Names @('AssignedByGroup', 'assignedByGroup'))
        $assignmentSource = [string](Get-ArrayaObjectValue -Object $stateRow -Names @('AssignmentSource', 'Source'))
        if ([string]::IsNullOrWhiteSpace($assignmentSource)) {
            $assignmentSource = if ([string]::IsNullOrWhiteSpace($assignedByGroup) -or $assignedByGroup -eq '00000000-0000-0000-0000-000000000000') { 'Direct' } else { 'Group' }
        }
        $rows.Add([pscustomobject]@{
            SkuId            = $skuId
            SkuName          = if ([string]::IsNullOrWhiteSpace($skuName)) { $skuId } else { $skuName }
            AssignedByGroup  = $assignedByGroup
            AssignmentSource = $assignmentSource
            State            = [string](Get-ArrayaObjectValue -Object $stateRow -Names @('State', 'state'))
            Error            = [string](Get-ArrayaObjectValue -Object $stateRow -Names @('Error', 'error'))
            LastUpdatedDateTime = [string](Get-ArrayaObjectValue -Object $stateRow -Names @('LastUpdatedDateTime', 'lastUpdatedDateTime'))
        }) | Out-Null
    }

    return @($rows.ToArray())
}

function Get-ArrayaGroupLicenseAssignedUserCountLookup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Users = @(),
        [Parameter(Mandatory = $false)][object[]]$Licenses = @()
    )

    $skuLookup = Get-ArrayaLicenseSkuLookup -Licenses $Licenses
    $userKeysByGroupId = @{}
    $hasAssignmentStateData = $false
    foreach ($user in @($Users)) {
        if ($null -eq $user) { continue }

        $assignmentRows = @(Get-ArrayaUserLicenseAssignmentRows -User $user -SkuLookup $skuLookup)
        if ($assignmentRows.Count -gt 0) {
            $hasAssignmentStateData = $true
        }

        $userKey = [string](Get-ArrayaObjectValue -Object $user -Names @('Id', 'ID', 'UserPrincipalName', 'Mail', 'DisplayName'))
        if ([string]::IsNullOrWhiteSpace($userKey)) {
            $userKey = [guid]::NewGuid().ToString()
        }

        foreach ($assignmentRow in @($assignmentRows)) {
            $assignedByGroup = ([string]$assignmentRow.AssignedByGroup).Trim()
            if ([string]::IsNullOrWhiteSpace($assignedByGroup) -or $assignedByGroup -eq '00000000-0000-0000-0000-000000000000') {
                continue
            }

            if (-not $userKeysByGroupId.ContainsKey($assignedByGroup)) {
                $userKeysByGroupId[$assignedByGroup] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            }
            [void]$userKeysByGroupId[$assignedByGroup].Add($userKey)
        }
    }

    $countLookup = @{}
    foreach ($groupId in @($userKeysByGroupId.Keys)) {
        $countLookup[$groupId] = [int]$userKeysByGroupId[$groupId].Count
    }

    return [pscustomobject]@{
        Counts = $countLookup
        HasAssignmentStateData = $hasAssignmentStateData
    }
}

function Test-ArrayaGroupManagesLicenses {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Group
    )

    if ($null -eq $Group) {
        return $false
    }

    $assignedLicenseCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $Group -Names @('AssignedLicenseCount'))
    if ($null -ne $assignedLicenseCount) {
        return ($assignedLicenseCount -gt 0)
    }

    $assignedLicenseSignals = @(
        'AssignedLicenseSkuIds',
        'AssignedLicenseSkuPartNumbers',
        'AssignedLicenseFriendlyNames'
    ) | ForEach-Object {
        [string](Get-ArrayaObjectValue -Object $Group -Names @($_))
    } | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '^(?i)notcollected|not surfaced'
    }

    if (@($assignedLicenseSignals).Count -gt 0) {
        return $true
    }

    # Older snapshots only carried the boolean. Trust it only when newer explicit
    # assignment counts are not present, so stale replay data does not mark every
    # group as a licensing group.
    return ((Convert-ArrayaToBoolean (Get-ArrayaObjectValue -Object $Group -Names @('IsManagingLicenses'))) -eq $true)
}

function New-ArrayaGroupLicensingSummaryRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Groups = @(),
        [Parameter(Mandatory = $false)][object[]]$Users = @(),
        [Parameter(Mandatory = $false)][object[]]$Licenses = @()
    )

    $assignedUserCountLookup = Get-ArrayaGroupLicenseAssignedUserCountLookup -Users $Users -Licenses $Licenses
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($group in @($Groups)) {
        $isManagingLicenses = Test-ArrayaGroupManagesLicenses -Group $group
        $assignedLicenseCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $group -Names @('AssignedLicenseCount'))
        if ($isManagingLicenses -ne $true) { continue }
        $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $group -Names @('OwnerCount'))
        $groupId = [string](Get-ArrayaObjectValue -Object $group -Names @('Id', 'ID'))
        $memberCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $group -Names @('MemberCount'))
        $memberCountSource = [string](Get-ArrayaObjectValue -Object $group -Names @('MemberCountSource'))
        $licensedUserCount = $null
        if ($assignedUserCountLookup.HasAssignmentStateData -eq $true) {
            $licensedUserCount = if (-not [string]::IsNullOrWhiteSpace($groupId) -and $assignedUserCountLookup.Counts.ContainsKey($groupId)) {
                [int]$assignedUserCountLookup.Counts[$groupId]
            }
            else {
                0
            }
        }
        $effectiveMemberCount = if ($null -ne $licensedUserCount) { $licensedUserCount } else { $memberCount }
        $membershipCountSource = if ($null -ne $licensedUserCount) { 'User licenseAssignmentStates assignedByGroup' } else { 'Group member count' }
        $rows.Add([pscustomobject]@{
            GroupName                    = [string](Get-ArrayaObjectValue -Object $group -Names @('DisplayName', 'Name'))
            GroupId                      = $groupId
            OwnerCount                   = $ownerCount
            MemberCount                  = $memberCount
            MemberCountSource            = $memberCountSource
            LicensedUserCount            = $licensedUserCount
            EffectiveMemberCount         = $effectiveMemberCount
            MembershipCountSource        = $membershipCountSource
            AssignedLicenseCount         = if ($null -ne $assignedLicenseCount) { [int]$assignedLicenseCount } else { $null }
            AssignedLicenseSkuIds        = [string](Get-ArrayaObjectValue -Object $group -Names @('AssignedLicenseSkuIds'))
            AssignedLicenseSkuPartNumbers = [string](Get-ArrayaObjectValue -Object $group -Names @('AssignedLicenseSkuPartNumbers'))
            AssignedLicenseFriendlyNames = [string](Get-ArrayaObjectValue -Object $group -Names @('AssignedLicenseFriendlyNames'))
            LicenseProcessingState       = [string](Get-ArrayaObjectValue -Object $group -Names @('LicenseProcessingState', 'licenseProcessingState'))
            RiskSignal                   = if ($null -ne $ownerCount -and $ownerCount -le 0) { 'License-managing group without owner' } else { 'Group-based licensing in use' }
            Recommendation               = 'Document license group ownership, assignment intent, and error-review cadence before relying on the group as a lifecycle control.'
            SourceWorksheet              = 'EntraIDGroups'
        }) | Out-Null
    }

    return @($rows.ToArray())
}

function Add-ArrayaGroupLicensingEffectiveCounts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $false)][object[]]$Users = @(),
        [Parameter(Mandatory = $false)][object[]]$Licenses = @()
    )

    $assignedUserCountLookup = Get-ArrayaGroupLicenseAssignedUserCountLookup -Users $Users -Licenses $Licenses
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }

        $groupId = [string](Get-ArrayaObjectValue -Object $row -Names @('GroupId', 'Id', 'ID'))
        $memberCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $row -Names @('MemberCount'))
        $memberCountSource = [string](Get-ArrayaObjectValue -Object $row -Names @('MemberCountSource'))
        $licensedUserCount = $null
        if ($assignedUserCountLookup.HasAssignmentStateData -eq $true) {
            $licensedUserCount = if (-not [string]::IsNullOrWhiteSpace($groupId) -and $assignedUserCountLookup.Counts.ContainsKey($groupId)) {
                [int]$assignedUserCountLookup.Counts[$groupId]
            }
            else {
                0
            }
        }
        $effectiveMemberCount = if ($null -ne $licensedUserCount) { $licensedUserCount } else { $memberCount }
        $membershipCountSource = if ($null -ne $licensedUserCount) { 'User licenseAssignmentStates assignedByGroup' } else { 'Group member count' }

        $row | Add-Member -MemberType NoteProperty -Name 'MemberCountSource' -Value $memberCountSource -Force
        $row | Add-Member -MemberType NoteProperty -Name 'LicensedUserCount' -Value $licensedUserCount -Force
        $row | Add-Member -MemberType NoteProperty -Name 'EffectiveMemberCount' -Value $effectiveMemberCount -Force
        $row | Add-Member -MemberType NoteProperty -Name 'MembershipCountSource' -Value $membershipCountSource -Force
    }

    return @($Rows)
}

function New-ArrayaLicenseOptimizationCandidateRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Users = @(),
        [Parameter(Mandatory = $false)][object[]]$Groups = @(),
        [Parameter(Mandatory = $false)][object[]]$Licenses = @(),
        [Parameter(Mandatory = $false)][object[]]$EmailActivityRows = @(),
        [Parameter(Mandatory = $false)][object[]]$TeamsActivityRows = @(),
        [Parameter(Mandatory = $false)][object[]]$Microsoft365ActivityRows = @(),
        [Parameter(Mandatory = $false)][int]$InactiveDaysThreshold = 90
    )

    $skuLookup = Get-ArrayaLicenseSkuLookup -Licenses $Licenses
    $emailActivityLookup = New-ArrayaActivityLookup -Rows $EmailActivityRows -PrincipalNames @('UserPrincipalName', 'UPN', 'Mail', 'PrimarySmtpAddress')
    $teamsActivityLookup = New-ArrayaActivityLookup -Rows $TeamsActivityRows -PrincipalNames @('UserPrincipalName', 'UPN', 'Mail', 'PrimarySmtpAddress')
    $microsoft365ActivityLookup = New-ArrayaActivityLookup -Rows $Microsoft365ActivityRows -PrincipalNames @('UserPrincipalName', 'UPN', 'Mail', 'PrimarySmtpAddress')
    $inactiveCutoff = (Get-Date).AddDays(-1 * [Math]::Abs($InactiveDaysThreshold))
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($group in @($Groups)) {
        $isManagingLicenses = Test-ArrayaGroupManagesLicenses -Group $group
        if ($isManagingLicenses -eq $true) {
            $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $group -Names @('OwnerCount'))
            if ($null -eq $ownerCount -or $ownerCount -le 0) {
                $rows.Add([pscustomobject]@{ ObjectType = 'Group'; DisplayName = [string](Get-ArrayaObjectValue -Object $group -Names @('DisplayName', 'Name')); UserPrincipalName = $null; SkuFamily = 'Group-based licensing'; SkuNames = [string](Get-ArrayaObjectValue -Object $group -Names @('AssignedLicenseFriendlyNames', 'AssignedLicenseSkuPartNumbers')); AssignmentSource = 'Group'; LicenseProcessingState = [string](Get-ArrayaObjectValue -Object $group -Names @('LicenseProcessingState', 'licenseProcessingState')); Issue = 'License-managing group without owner'; RecommendedAction = 'Assign accountable owners to licensing groups and document change-control expectations.'; SourceWorksheet = 'EntraIDGroups' }) | Out-Null
            }
        }
    }

    foreach ($user in @($Users)) {
        $displayName = [string](Get-ArrayaObjectValue -Object $user -Names @('DisplayName'))
        $upn = [string](Get-ArrayaObjectValue -Object $user -Names @('UserPrincipalName'))
        $assignmentRows = @(Get-ArrayaUserLicenseAssignmentRows -User $user -SkuLookup $skuLookup)
        foreach ($skuGroup in @($assignmentRows | Group-Object { if (-not [string]::IsNullOrWhiteSpace([string]$_.SkuName)) { [string]$_.SkuName } else { [string]$_.SkuId } })) {
            if ([string]::IsNullOrWhiteSpace([string]$skuGroup.Name)) { continue }
            $hasDirect = @($skuGroup.Group | Where-Object { [string]$_.AssignmentSource -eq 'Direct' }).Count -gt 0
            $hasGroup = @($skuGroup.Group | Where-Object { [string]$_.AssignmentSource -eq 'Group' }).Count -gt 0
            if ($hasDirect -and $hasGroup) {
                $rows.Add([pscustomobject]@{ ObjectType = 'User'; DisplayName = $displayName; UserPrincipalName = $upn; SkuFamily = 'Duplicate assignment source'; SkuNames = [string]$skuGroup.Name; AssignmentSource = 'Direct + Group'; Issue = 'Same SKU assigned directly and by group'; RecommendedAction = 'Remove duplicate direct assignments after confirming the group-based assignment is intentional and healthy.'; SourceWorksheet = 'Users' }) | Out-Null
            }
        }
        foreach ($assignmentRow in @($assignmentRows)) {
            $assignmentError = [string]$assignmentRow.Error
            $assignmentState = [string]$assignmentRow.State
            $hasAssignmentError = (
                (-not [string]::IsNullOrWhiteSpace($assignmentError) -and $assignmentError -notmatch '^(?i)none|noerror|success$') -or
                (-not [string]::IsNullOrWhiteSpace($assignmentState) -and $assignmentState -match '(?i)error|failed|failure')
            )
            if ($hasAssignmentError) {
                $errorDetail = if (-not [string]::IsNullOrWhiteSpace($assignmentError) -and $assignmentError -notmatch '^(?i)none|noerror|success$') { $assignmentError } else { $assignmentState }
                $rows.Add([pscustomobject]@{
                    ObjectType          = 'User'
                    DisplayName         = $displayName
                    UserPrincipalName   = $upn
                    SkuFamily           = 'Assignment error'
                    SkuNames            = $assignmentRow.SkuName
                    AssignmentSource    = $assignmentRow.AssignmentSource
                    AssignmentState     = $assignmentState
                    ErrorCode           = $errorDetail
                    AssignedByGroup     = $assignmentRow.AssignedByGroup
                    LastUpdatedDateTime = $assignmentRow.LastUpdatedDateTime
                    Issue               = "License assignment error: $errorDetail"
                    RecommendedAction   = 'Review Microsoft Graph license assignment errors, resolve the assignment conflict or capacity/usage-location issue, and reprocess group-based license assignment where appropriate.'
                    SourceWorksheet     = 'Users'
                }) | Out-Null
            }
        }
        $assignedLicenseNames = @($assignmentRows | ForEach-Object { [string]$_.SkuName } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($assignedLicenseNames.Count -eq 0) {
            $assignedLicenseNames = Convert-ArrayaToStringList (Get-ArrayaObjectValue -Object $user -Names @('AssignedLicensesFriendly', 'AssignedLicenses'))
        }
        $suiteGroups = @($assignedLicenseNames | ForEach-Object { [pscustomobject]@{ Name = $_; Family = Get-ArrayaLicenseSuiteFamily -SkuName $_ } } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Family) } | Group-Object Family | Where-Object { $_.Count -gt 1 })
        foreach ($suiteGroup in @($suiteGroups)) {
            $rows.Add([pscustomobject]@{ ObjectType = 'User'; DisplayName = $displayName; UserPrincipalName = $upn; SkuFamily = [string]$suiteGroup.Name; SkuNames = (@($suiteGroup.Group | ForEach-Object { $_.Name }) -join '; '); AssignmentSource = 'Mixed/Unknown'; Issue = 'Likely duplicate suite assignment'; RecommendedAction = 'Review stacked Microsoft 365 / Office 365 suite assignments and retain the license set that matches the user role and required services.'; SourceWorksheet = 'Users' }) | Out-Null
        }

        if ($assignedLicenseNames.Count -gt 0) {
            $enabled = Convert-ArrayaToBoolean (Get-ArrayaObjectValue -Object $user -Names @('AccountEnabled', 'Enabled'))
            try {
                $activityState = Get-ArrayaUserLicenseActivityState `
                    -User $user `
                    -EmailActivityLookup $emailActivityLookup `
                    -TeamsActivityLookup $teamsActivityLookup `
                    -Microsoft365ActivityLookup $microsoft365ActivityLookup
            }
            catch {
                $activityState = [pscustomobject]@{
                    LastSignInDateTime           = $null
                    LastExchangeActivityDate     = $null
                    LastTeamsActivityDate        = $null
                    LastMicrosoft365ActivityDate = $null
                    LastActivityDate             = $null
                    AvailableSignalCount         = 0
                }
            }
            $isDisabledLicensedUser = ($enabled -eq $false)
            $isInactiveLicensedUser = (
                -not $isDisabledLicensedUser -and
                $activityState.AvailableSignalCount -gt 0 -and
                $activityState.LastActivityDate -and
                $activityState.LastActivityDate -lt $inactiveCutoff
            )

            if ($isDisabledLicensedUser -or $isInactiveLicensedUser) {
                $basis = if ($isDisabledLicensedUser) {
                    'Account disabled'
                }
                else {
                    "No recent activity in surfaced sign-in or workload signals for $InactiveDaysThreshold+ days"
                }
                $rows.Add([pscustomobject]@{
                    ObjectType                   = 'User'
                    DisplayName                  = $displayName
                    UserPrincipalName            = $upn
                    Enabled                      = $enabled
                    SkuFamily                    = 'Inactive licensed user'
                    SkuNames                     = (($assignedLicenseNames | Select-Object -Unique) -join '; ')
                    AssignmentSource             = if (@($assignmentRows | Where-Object { [string]$_.AssignmentSource -eq 'Group' }).Count -gt 0) { 'Direct/Group review' } else { 'Direct/Unknown' }
                    Issue                        = if ($isDisabledLicensedUser) { 'Disabled user with assigned licenses' } else { 'Inactive licensed user with assigned licenses' }
                    LastSignInDateTime           = Get-ArrayaDateText -Date $activityState.LastSignInDateTime
                    LastExchangeActivityDate     = Get-ArrayaDateText -Date $activityState.LastExchangeActivityDate
                    LastTeamsActivityDate        = Get-ArrayaDateText -Date $activityState.LastTeamsActivityDate
                    LastMicrosoft365ActivityDate = Get-ArrayaDateText -Date $activityState.LastMicrosoft365ActivityDate
                    LastActivityDate             = Get-ArrayaDateText -Date $activityState.LastActivityDate
                    InactivityBasis              = $basis
                    RecommendedAction            = 'Validate whether the user still requires paid Microsoft 365 licensing; reclaim or right-size licenses after confirming mailbox, OneDrive, Teams, and compliance hold requirements.'
                    SourceWorksheet              = 'Users'
                }) | Out-Null
            }
        }
    }

    return @($rows.ToArray())
}

