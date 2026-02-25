[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$BaselineJsonPath,
    [Parameter(Mandatory = $true)]
    [string]$CurrentJsonPath,
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder,
    [Parameter(Mandatory = $false)]
    [string]$OutputPrefix = 'M365TenantSnapshotComparison',
    [Parameter(Mandatory = $false)]
    [ValidateRange(30, 1095)]
    [int]$StaleDeviceDays = 180,
    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'

function ConvertTo-Array {
    param($InputObject)

    if ($null -eq $InputObject) { return @() }
    if ($InputObject -is [System.Collections.IDictionary]) { return @($InputObject.Values) }
    if ($InputObject -is [string]) { return @($InputObject) }
    if ($InputObject -is [System.Collections.IEnumerable]) { return @($InputObject) }
    return @($InputObject)
}

function Get-Value {
    param(
        $Object,
        [string[]]$Names
    )

    if ($null -eq $Object) { return $null }

    foreach ($name in $Names) {
        if ($Object -is [System.Collections.IDictionary] -and $Object.Contains($name)) {
            return $Object[$name]
        }
        if ($Object.PSObject.Properties.Name -contains $name) {
            return $Object.$name
        }
    }
    return $null
}

function Convert-ToNumber {
    param($Value)
    if ($null -eq $Value) { return $null }
    $numeric = 0.0
    if ([double]::TryParse(($Value.ToString()), [ref]$numeric)) { return [double]$numeric }
    return $null
}

function Convert-ToDate {
    param($Value)
    if ($null -eq $Value) { return $null }
    $dateValue = [datetime]::MinValue
    if ([datetime]::TryParse(($Value.ToString()), [ref]$dateValue)) { return $dateValue }
    return $null
}

function Get-AssessmentContent {
    param([string]$Path)

    if (-not (Test-Path -Path $Path)) {
        throw "Snapshot JSON not found: $Path"
    }

    $jsonRoot = Get-Content -Raw -Path $Path | ConvertFrom-Json
    $data = if ($jsonRoot.PSObject.Properties.Name -contains 'Data') { $jsonRoot.Data } else { $jsonRoot }
    [PSCustomObject]@{
        Raw        = $jsonRoot
        Data       = $data
        GeneratedAt = (Get-Value -Object $jsonRoot -Names @('GeneratedAt'))
        Path       = (Resolve-Path -Path $Path).Path
    }
}

function Get-Metrics {
    param(
        [Parameter(Mandatory = $true)]
        $AssessmentData,
        [Parameter(Mandatory = $true)]
        [int]$StaleDeviceDays
    )

    $data = $AssessmentData.Data

    $secureScoreRows = ConvertTo-Array (Get-Value -Object $data -Names @('SecuritySecureScore', 'SecureScore'))
    $latestSecureScore = $secureScoreRows |
        Sort-Object { Convert-ToDate (Get-Value -Object $_ -Names @('CreatedDateTime', 'createdDateTime')) } -Descending |
        Select-Object -First 1
    $currentScore = Convert-ToNumber (Get-Value -Object $latestSecureScore -Names @('CurrentScore', 'currentScore'))
    $maxScore = Convert-ToNumber (Get-Value -Object $latestSecureScore -Names @('MaxScore', 'maxScore'))
    $secureScorePct = if ($null -ne $currentScore -and $null -ne $maxScore -and $maxScore -gt 0) {
        [math]::Round(($currentScore / $maxScore) * 100, 2)
    } else { $null }

    $caPolicies = ConvertTo-Array (Get-Value -Object $data -Names @('ConditionalAccessPolicies', 'ConditionalAccess'))
    $enabledCaPolicies = @(
        $caPolicies | Where-Object {
            $state = (Get-Value -Object $_ -Names @('State', 'state'))
            $null -ne $state -and $state.ToString().ToLower().Contains('enabled')
        }
    )

    $adminRows = ConvertTo-Array (Get-Value -Object $data -Names @('AllOffice365Admins', 'Office365Admins', 'Admins'))
    $globalAdmins = @(
        $adminRows | Where-Object {
            $role = (Get-Value -Object $_ -Names @('Role', 'RoleName', 'DirectoryRole', 'AdminRole'))
            $null -ne $role -and $role.ToString() -match 'Global Administrator|Company Administrator'
        }
    )
    $globalAdminCount = if ($globalAdmins.Count -gt 0) { $globalAdmins.Count } else { $adminRows.Count }

    $domainRows = ConvertTo-Array (Get-Value -Object $data -Names @('Domains'))
    $unverifiedDomains = @(
        $domainRows | Where-Object {
            $isVerified = Get-Value -Object $_ -Names @('IsVerified', 'Verified', 'isVerified')
            if ($isVerified -is [bool]) { return (-not $isVerified) }
            if ($null -eq $isVerified) { return $false }
            return ($isVerified.ToString().ToLower() -notin @('true', 'verified'))
        }
    )

    $licenseRows = ConvertTo-Array (Get-Value -Object $data -Names @('LicenseSKUs', 'Licenses'))
    $maxLicenseUtilizationPct = $null
    foreach ($sku in $licenseRows) {
        $consumed = Convert-ToNumber (Get-Value -Object $sku -Names @('ConsumedUnits', 'Consumed', 'Assigned'))
        $active = Convert-ToNumber (Get-Value -Object $sku -Names @('ActiveUnits', 'EnabledUnits', 'TotalUnits'))
        if ($null -eq $active) {
            $prepaid = Get-Value -Object $sku -Names @('PrepaidUnits')
            if ($prepaid) {
                $active = Convert-ToNumber (Get-Value -Object $prepaid -Names @('Enabled', 'enabled'))
            }
        }
        if ($null -ne $consumed -and $null -ne $active -and $active -gt 0) {
            $util = [math]::Round(($consumed / $active) * 100, 2)
            if ($null -eq $maxLicenseUtilizationPct -or $util -gt $maxLicenseUtilizationPct) {
                $maxLicenseUtilizationPct = $util
            }
        }
    }

    $deviceRows = ConvertTo-Array (Get-Value -Object $data -Names @('DeviceDetails', 'Devices'))
    $staleDevicePct = $null
    if ($deviceRows.Count -gt 0) {
        $cutoff = (Get-Date).AddDays(-1 * $StaleDeviceDays)
        $staleDevices = @(
            $deviceRows | Where-Object {
                $lastSignIn = Convert-ToDate (Get-Value -Object $_ -Names @('ApproximateLastSignInDateTime', 'LastSignInDateTime', 'LastSignInDate', 'LastLogonDateTime'))
                $null -ne $lastSignIn -and $lastSignIn -lt $cutoff
            }
        )
        $staleDevicePct = [math]::Round((($staleDevices.Count / $deviceRows.Count) * 100), 2)
    }

    [PSCustomObject]@{
        GeneratedAt                    = $AssessmentData.GeneratedAt
        Path                           = $AssessmentData.Path
        SecureScorePercent             = $secureScorePct
        ConditionalAccessPolicyCount   = $caPolicies.Count
        EnabledConditionalAccessCount  = $enabledCaPolicies.Count
        GlobalAdminCount               = $globalAdminCount
        UnverifiedDomainCount          = $unverifiedDomains.Count
        MaxLicenseUtilizationPercent   = $maxLicenseUtilizationPct
        StaleDevicePercent             = $staleDevicePct
        DeviceCount                    = $deviceRows.Count
    }
}

function New-MetricComparison {
    param(
        [string]$Metric,
        [Nullable[Double]]$Baseline,
        [Nullable[Double]]$Current,
        [ValidateSet('HigherIsBetter', 'LowerIsBetter')]
        [string]$Direction
    )

    $delta = $null
    $status = 'Unknown'

    if ($null -eq $Baseline -and $null -eq $Current) {
        $status = 'NoData'
    } elseif ($null -eq $Baseline -and $null -ne $Current) {
        $status = 'NewData'
    } elseif ($null -ne $Baseline -and $null -eq $Current) {
        $status = 'MissingData'
    } else {
        $delta = [math]::Round(($Current - $Baseline), 2)
        if ($delta -eq 0) {
            $status = 'NoChange'
        } elseif ($Direction -eq 'HigherIsBetter') {
            $status = if ($delta -gt 0) { 'Improved' } else { 'Regressed' }
        } else {
            $status = if ($delta -lt 0) { 'Improved' } else { 'Regressed' }
        }
    }

    [PSCustomObject]@{
        Metric    = $Metric
        Baseline  = $Baseline
        Current   = $Current
        Delta     = $delta
        Direction = $Direction
        Status    = $status
    }
}

$baseline = Get-AssessmentContent -Path $BaselineJsonPath
$current = Get-AssessmentContent -Path $CurrentJsonPath

$baselineMetrics = Get-Metrics -AssessmentData $baseline -StaleDeviceDays $StaleDeviceDays
$currentMetrics = Get-Metrics -AssessmentData $current -StaleDeviceDays $StaleDeviceDays

$comparisons = @(
    New-MetricComparison -Metric 'SecureScorePercent' -Baseline $baselineMetrics.SecureScorePercent -Current $currentMetrics.SecureScorePercent -Direction HigherIsBetter
    New-MetricComparison -Metric 'ConditionalAccessPolicyCount' -Baseline $baselineMetrics.ConditionalAccessPolicyCount -Current $currentMetrics.ConditionalAccessPolicyCount -Direction HigherIsBetter
    New-MetricComparison -Metric 'EnabledConditionalAccessCount' -Baseline $baselineMetrics.EnabledConditionalAccessCount -Current $currentMetrics.EnabledConditionalAccessCount -Direction HigherIsBetter
    New-MetricComparison -Metric 'GlobalAdminCount' -Baseline $baselineMetrics.GlobalAdminCount -Current $currentMetrics.GlobalAdminCount -Direction LowerIsBetter
    New-MetricComparison -Metric 'UnverifiedDomainCount' -Baseline $baselineMetrics.UnverifiedDomainCount -Current $currentMetrics.UnverifiedDomainCount -Direction LowerIsBetter
    New-MetricComparison -Metric 'MaxLicenseUtilizationPercent' -Baseline $baselineMetrics.MaxLicenseUtilizationPercent -Current $currentMetrics.MaxLicenseUtilizationPercent -Direction LowerIsBetter
    New-MetricComparison -Metric 'StaleDevicePercent' -Baseline $baselineMetrics.StaleDevicePercent -Current $currentMetrics.StaleDevicePercent -Direction LowerIsBetter
)

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Split-Path -Path $CurrentJsonPath -Parent
}
if (-not (Test-Path -Path $OutputFolder)) {
    $null = New-Item -ItemType Directory -Path $OutputFolder -Force
}

$jsonOutPath = Join-Path -Path $OutputFolder -ChildPath "$OutputPrefix.json"
$csvOutPath = Join-Path -Path $OutputFolder -ChildPath "$OutputPrefix.csv"
$mdOutPath = Join-Path -Path $OutputFolder -ChildPath "$OutputPrefix.md"

$payload = [PSCustomObject]@{
    GeneratedAt      = (Get-Date).ToString('o')
    BaselineSnapshot = $baselineMetrics
    CurrentSnapshot  = $currentMetrics
    Comparison       = $comparisons
}

$payload | ConvertTo-Json -Depth 8 | Set-Content -Path $jsonOutPath -Encoding UTF8
$comparisons | Export-Csv -Path $csvOutPath -NoTypeInformation -Encoding UTF8

$statusCounts = $comparisons | Group-Object -Property Status | Sort-Object Name
$lines = @(
    "# M365 Tenant Snapshot Comparison",
    "",
    "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    "Baseline: $($baselineMetrics.Path)",
    "Current: $($currentMetrics.Path)",
    "",
    "## Status Summary"
)
foreach ($group in $statusCounts) {
    $lines += "- $($group.Name): $($group.Count)"
}
$lines += @(
    "",
    "## Metric Comparison",
    "",
    "| Metric | Baseline | Current | Delta | Status |",
    "|---|---|---|---|---|"
)
foreach ($row in $comparisons) {
    $lines += "| $($row.Metric) | $($row.Baseline) | $($row.Current) | $($row.Delta) | $($row.Status) |"
}
Set-Content -Path $mdOutPath -Value ($lines -join [Environment]::NewLine) -Encoding UTF8

$result = [PSCustomObject]@{
    JsonPath = $jsonOutPath
    CsvPath  = $csvOutPath
    MdPath   = $mdOutPath
}

Write-Host "Snapshot comparison generated."
Write-Host "  JSON: $jsonOutPath"
Write-Host "  CSV : $csvOutPath"
Write-Host "  MD  : $mdOutPath"

if ($PassThru) {
    $result
}
