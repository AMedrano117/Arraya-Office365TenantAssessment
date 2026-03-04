function Get-ArrayaEntraGroupClassification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $GroupDetails
    )

    $membershipType = if ($GroupDetails.groupTypes -contains 'DynamicMembership') { 'Dynamic Group' } else { 'Assigned' }
    $membershipRule = if ($membershipType -eq 'Dynamic Group') { $GroupDetails.membershipRule } else { $null }

    if ($GroupDetails.groupTypes -contains 'Unified') {
        $groupType = 'Microsoft 365'
    }
    elseif ($GroupDetails.mailEnabled -eq $true -and $GroupDetails.securityEnabled -eq $true) {
        $groupType = 'Mail-enabled Security Group'
    }
    elseif ($GroupDetails.mailEnabled -eq $false -and $GroupDetails.securityEnabled -eq $true) {
        $groupType = 'Security Group'
    }
    elseif ($GroupDetails.mailEnabled -eq $true -and $GroupDetails.securityEnabled -eq $false) {
        $groupType = 'Distribution Group'
    }
    else {
        $groupType = 'Unknown'
    }

    $source = if ($GroupDetails.onPremisesSyncEnabled) { 'On-Premises' } else { 'Cloud' }

    return [PSCustomObject]@{
        MembershipType = $membershipType
        MembershipRule = $membershipRule
        GroupType      = $groupType
        Source         = $source
    }
}
