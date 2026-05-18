function Get-ArrayaTenantSnapshotMetricSetFromContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Context,
        [Parameter(Mandatory = $false)]
        [ValidateRange(30, 1095)]
        [int]$StaleDeviceDays = 180
    )

    $metrics = Get-ArrayaTenantSnapshotMetricSet `
        -SecureScoreRows       (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $Context.LegacyData -Names @('SecuritySecureScore', 'SecureScore'))) `
        -ConditionalAccessRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $Context.LegacyData -Names @('ConditionalAccessPolicies', 'ConditionalAccess'))) `
        -AdminRows             (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $Context.LegacyData -Names @('AllOffice365Admins', 'Office365Admins', 'Admins'))) `
        -DomainRows            (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $Context.LegacyData -Names @('Domains'))) `
        -LicenseRows           (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $Context.LegacyData -Names @('LicenseSKUs', 'Licenses'))) `
        -DeviceRows            (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $Context.LegacyData -Names @('DeviceDetails', 'Devices'))) `
        -StaleDeviceDays       $StaleDeviceDays
    $metrics | Add-Member -MemberType NoteProperty -Name GeneratedAt -Value $Context.GeneratedAt -Force
    $metrics | Add-Member -MemberType NoteProperty -Name Path -Value $Context.Path -Force
    return $metrics
}
