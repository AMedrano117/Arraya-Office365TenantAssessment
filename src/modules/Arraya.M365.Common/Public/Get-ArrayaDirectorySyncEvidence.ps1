function Get-ArrayaDirectorySyncEvidence {
    <#
    .SYNOPSIS
        Interprets directory-synchronization evidence without guessing.

    .DESCRIPTION
        Returns one of three states:

          Synchronized    - organization.onPremisesSyncEnabled is true, or collected user
                            objects carry onPremisesSyncEnabled = true.
          NotSynchronized - the organization read succeeded and reports no synchronization.
                            Only this state supports a "cloud-only" statement.
          Unknown         - the collector did not run, failed, or returned an empty flag with
                            nothing to show that the organization read succeeded.

        Microsoft Graph returns null for a tenant that never synchronized, which is the same
        value the collector stores when the organization read fails. The read is treated as
        successful when TenantInfo, which comes from the same organization call, carries a
        tenant id or display name.

        The evidence does not distinguish Microsoft Entra Connect Sync from Cloud Sync, and it
        cannot show an on-premises directory that is not synchronized to this tenant.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [System.Collections.IDictionary]$TenantStatsHash
    )

    $newResult = {
        param([string]$State, [string]$Reason, [int]$SyncedUserCount, $OrganizationFlag)
        [pscustomobject][ordered]@{
            State            = $State
            Reason           = $Reason
            SyncedUserCount  = $SyncedUserCount
            OrganizationFlag = $OrganizationFlag
        }
    }

    if ($null -eq $TenantStatsHash) {
        return (& $newResult 'Unknown' 'No assessment data was supplied.' 0 $null)
    }

    $syncedUserCount = 0
    if ($TenantStatsHash.Contains('Users') -and $null -ne $TenantStatsHash['Users']) {
        $userValue = $TenantStatsHash['Users']
        $userRows = if ($userValue -is [System.Collections.IDictionary]) { @($userValue.Values) } else { @($userValue) }
        $syncedUserCount = @(
            $userRows | Where-Object {
                $null -ne $_ -and
                (Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $_ -Names @('OnPremisesSyncEnabled', 'IsDirSynced'))) -eq $true
            }
        ).Count
    }

    $configuration = if ($TenantStatsHash.Contains('AdConnectConfiguration')) { $TenantStatsHash['AdConnectConfiguration'] } else { $null }
    $summary = Get-ArrayaObjectValue -Object $configuration -Names @('Summary')
    $organizationFlag = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $summary -Names @('OnPremisesSyncEnabled'))

    if ($organizationFlag -eq $true) {
        return (& $newResult 'Synchronized' 'organization.onPremisesSyncEnabled is true.' $syncedUserCount $true)
    }

    if ($syncedUserCount -gt 0) {
        return (& $newResult 'Synchronized' ("{0} collected user object(s) carry onPremisesSyncEnabled = true." -f $syncedUserCount) $syncedUserCount $organizationFlag)
    }

    $stepEvidence = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'AdConnectConfiguration' -CollectorStepName 'AD Connect sync details'
    if ($stepEvidence.State -in @('Failed', 'NotCollected') -and $null -eq $summary) {
        return (& $newResult 'Unknown' $stepEvidence.Reason 0 $null)
    }

    if ($null -eq $summary) {
        return (& $newResult 'Unknown' 'Directory synchronization details were not collected.' 0 $null)
    }

    if ($organizationFlag -eq $false) {
        return (& $newResult 'NotSynchronized' 'organization.onPremisesSyncEnabled is false (synchronization was enabled previously and is now off).' 0 $false)
    }

    $tenantInfo = if ($TenantStatsHash.Contains('TenantInfo')) { $TenantStatsHash['TenantInfo'] } else { $null }
    $organizationReadSucceeded = -not [string]::IsNullOrWhiteSpace([string](Get-ArrayaObjectValue -Object $tenantInfo -Names @('TenantId', 'TenantID', 'DisplayName')))
    if ($organizationReadSucceeded) {
        return (& $newResult 'NotSynchronized' 'The organization read succeeded and onPremisesSyncEnabled is not set; no collected user object is synchronized.' 0 $null)
    }

    return (& $newResult 'Unknown' 'onPremisesSyncEnabled is empty and nothing shows that the organization read succeeded.' 0 $null)
}
