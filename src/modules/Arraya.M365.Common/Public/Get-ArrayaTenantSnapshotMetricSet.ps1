function Get-ArrayaTenantSnapshotMetricSet {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $SecureScoreRows,
        [AllowNull()]
        $ConditionalAccessRows,
        [AllowNull()]
        $AdminRows,
        [AllowNull()]
        $DomainRows,
        [AllowNull()]
        $LicenseRows,
        [AllowNull()]
        $DeviceRows,
        [Parameter(Mandatory = $false)]
        [ValidateRange(30, 1095)]
        [int]$StaleDeviceDays = 180
    )

    $secureScoreRows = Convert-ArrayaObjectToArray $SecureScoreRows
    $conditionalAccessRows = Convert-ArrayaObjectToArray $ConditionalAccessRows
    $adminRows = Convert-ArrayaObjectToArray $AdminRows
    $domainRows = Convert-ArrayaObjectToArray $DomainRows
    $licenseRows = Convert-ArrayaObjectToArray $LicenseRows
    $deviceRows = Convert-ArrayaObjectToArray $DeviceRows

    $latestSecureScore = $secureScoreRows |
        Sort-Object { Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('CreatedDateTime', 'createdDateTime')) } -Descending |
        Select-Object -First 1
    $currentScore = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $latestSecureScore -Names @('CurrentScore', 'currentScore'))
    $maxScore = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $latestSecureScore -Names @('MaxScore', 'maxScore'))
    $secureScorePercent = if ($null -ne $currentScore -and $null -ne $maxScore -and $maxScore -gt 0) {
        [math]::Round(($currentScore / $maxScore) * 100, 2)
    }
    else {
        $null
    }

    $enabledConditionalAccessRows = @(
        $conditionalAccessRows | Where-Object {
            $state = Get-ArrayaObjectValue -Object $_ -Names @('State', 'state')
            $null -ne $state -and $state.ToString().ToLowerInvariant().Contains('enabled')
        }
    )

    $globalAdminRows = @(
        $adminRows | Where-Object {
            $role = Get-ArrayaObjectValue -Object $_ -Names @('Role', 'RoleName', 'DirectoryRole', 'AdminRole')
            $null -ne $role -and $role.ToString() -match 'Global Administrator|Company Administrator'
        }
    )
    $globalAdminCount = if ($globalAdminRows.Count -gt 0) {
        $globalAdminRows.Count
    }
    else {
        $adminRows.Count
    }

    $unverifiedDomainRows = @(
        $domainRows | Where-Object {
            $isVerified = Get-ArrayaObjectValue -Object $_ -Names @('IsVerified', 'Verified', 'isVerified')
            if ($isVerified -is [bool]) { return (-not $isVerified) }
            if ($null -eq $isVerified) { return $false }
            return ($isVerified.ToString().ToLowerInvariant() -notin @('true', 'verified'))
        }
    )

    $maxLicenseUtilizationPercent = $null
    $highUtilizationSkus = New-Object System.Collections.Generic.List[string]
    foreach ($sku in $licenseRows) {
        $skuName = Get-ArrayaObjectValue -Object $sku -Names @('SkuPartNumber', 'DisplayName', 'ProductName', 'SkuId')
        if ([string]::IsNullOrWhiteSpace([string]$skuName)) {
            $skuName = 'UnknownSKU'
        }

        $consumed = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sku -Names @('ConsumedUnits', 'Consumed', 'Assigned'))
        $active = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sku -Names @('ActiveUnits', 'EnabledUnits', 'TotalUnits'))
        if ($null -eq $active) {
            $prepaid = Get-ArrayaObjectValue -Object $sku -Names @('PrepaidUnits')
            if ($prepaid) {
                $active = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $prepaid -Names @('Enabled', 'enabled'))
            }
        }

        if ($null -ne $consumed -and $null -ne $active -and $active -gt 0) {
            $utilizationPercent = [math]::Round(($consumed / $active) * 100, 2)
            if ($null -eq $maxLicenseUtilizationPercent -or $utilizationPercent -gt $maxLicenseUtilizationPercent) {
                $maxLicenseUtilizationPercent = $utilizationPercent
            }
            if ($utilizationPercent -ge 95) {
                $highUtilizationSkus.Add("$skuName ($utilizationPercent%)") | Out-Null
            }
        }
    }

    $staleDevices = @()
    $staleDevicePercent = $null
    if ($deviceRows.Count -gt 0) {
        $cutoff = (Get-Date).AddDays(-1 * $StaleDeviceDays)
        $staleDevices = @(
            $deviceRows | Where-Object {
                $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('ApproximateLastSignInDateTime', 'LastSignInDateTime', 'LastSignInDate', 'LastLogonDateTime'))
                $null -ne $lastSignIn -and $lastSignIn -lt $cutoff
            }
        )
        $staleDevicePercent = [math]::Round((($staleDevices.Count / $deviceRows.Count) * 100), 2)
    }

    return [PSCustomObject]@{
        SecureScoreRows               = $secureScoreRows
        LatestSecureScore             = $latestSecureScore
        SecureScoreCurrentScore       = $currentScore
        SecureScoreMaxScore           = $maxScore
        SecureScorePercent            = $secureScorePercent
        ConditionalAccessRows         = $conditionalAccessRows
        ConditionalAccessPolicyCount  = $conditionalAccessRows.Count
        EnabledConditionalAccessRows  = $enabledConditionalAccessRows
        EnabledConditionalAccessCount = $enabledConditionalAccessRows.Count
        AdminRows                     = $adminRows
        GlobalAdminRows               = $globalAdminRows
        GlobalAdminCount              = $globalAdminCount
        DomainRows                    = $domainRows
        UnverifiedDomainRows          = $unverifiedDomainRows
        UnverifiedDomainCount         = $unverifiedDomainRows.Count
        LicenseRows                   = $licenseRows
        HighUtilizationSkus           = $highUtilizationSkus.ToArray()
        MaxLicenseUtilizationPercent  = $maxLicenseUtilizationPercent
        DeviceRows                    = $deviceRows
        DeviceCount                   = $deviceRows.Count
        StaleDevices                  = $staleDevices
        StaleDeviceCount              = $staleDevices.Count
        StaleDevicePercent            = $staleDevicePercent
    }
}
