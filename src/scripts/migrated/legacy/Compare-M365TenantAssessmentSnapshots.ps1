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

function Import-ArrayaCommonModuleForSnapshotComparison {
    [CmdletBinding()]
    param()

    $commonModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'))
    if (-not (Test-Path -Path $commonModuleManifestPath)) {
        throw "Required common module manifest not found: $commonModuleManifestPath"
    }

    $resolvedCommonManifestPath = (Resolve-Path -Path $commonModuleManifestPath).Path
    $loadedCommonModule = Get-Module -Name 'Arraya.M365.Common' -ErrorAction SilentlyContinue | Select-Object -First 1
    $requiredCommonCommands = @(
        'Convert-ArrayaObjectToArray',
        'Get-ArrayaObjectValue',
        'Convert-ArrayaToNumber',
        'Convert-ArrayaToDate',
        'Get-ArrayaTenantSnapshotMetricSet',
        'Import-ArrayaTenantSnapshotContext',
        'Resolve-ArrayaSnapshotOutputContext'
    )
    $missingCommonCommands = @(
        $requiredCommonCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
    )
    if (
        -not $loadedCommonModule -or
        $loadedCommonModule.Path -ne $resolvedCommonManifestPath -or
        $missingCommonCommands.Count -gt 0
    ) {
        Import-Module -Name $resolvedCommonManifestPath -Force -ErrorAction Stop
    }
}
Import-ArrayaCommonModuleForSnapshotComparison

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

$baseline = Import-ArrayaTenantSnapshotContext -Path $BaselineJsonPath -Purpose Export
$current = Import-ArrayaTenantSnapshotContext -Path $CurrentJsonPath -Purpose Export

$baselineMetrics = Get-ArrayaTenantSnapshotMetricSet `
    -SecureScoreRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $baseline.LegacyData -Names @('SecuritySecureScore', 'SecureScore'))) `
    -ConditionalAccessRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $baseline.LegacyData -Names @('ConditionalAccessPolicies', 'ConditionalAccess'))) `
    -AdminRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $baseline.LegacyData -Names @('AllOffice365Admins', 'Office365Admins', 'Admins'))) `
    -DomainRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $baseline.LegacyData -Names @('Domains'))) `
    -LicenseRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $baseline.LegacyData -Names @('LicenseSKUs', 'Licenses'))) `
    -DeviceRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $baseline.LegacyData -Names @('DeviceDetails', 'Devices'))) `
    -StaleDeviceDays $StaleDeviceDays
$baselineMetrics | Add-Member -MemberType NoteProperty -Name GeneratedAt -Value $baseline.GeneratedAt -Force
$baselineMetrics | Add-Member -MemberType NoteProperty -Name Path -Value $baseline.Path -Force

$currentMetrics = Get-ArrayaTenantSnapshotMetricSet `
    -SecureScoreRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $current.LegacyData -Names @('SecuritySecureScore', 'SecureScore'))) `
    -ConditionalAccessRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $current.LegacyData -Names @('ConditionalAccessPolicies', 'ConditionalAccess'))) `
    -AdminRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $current.LegacyData -Names @('AllOffice365Admins', 'Office365Admins', 'Admins'))) `
    -DomainRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $current.LegacyData -Names @('Domains'))) `
    -LicenseRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $current.LegacyData -Names @('LicenseSKUs', 'Licenses'))) `
    -DeviceRows (Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $current.LegacyData -Names @('DeviceDetails', 'Devices'))) `
    -StaleDeviceDays $StaleDeviceDays
$currentMetrics | Add-Member -MemberType NoteProperty -Name GeneratedAt -Value $current.GeneratedAt -Force
$currentMetrics | Add-Member -MemberType NoteProperty -Name Path -Value $current.Path -Force

$comparisons = @(
    New-MetricComparison -Metric 'SecureScorePercent' -Baseline $baselineMetrics.SecureScorePercent -Current $currentMetrics.SecureScorePercent -Direction HigherIsBetter
    New-MetricComparison -Metric 'ConditionalAccessPolicyCount' -Baseline $baselineMetrics.ConditionalAccessPolicyCount -Current $currentMetrics.ConditionalAccessPolicyCount -Direction HigherIsBetter
    New-MetricComparison -Metric 'EnabledConditionalAccessCount' -Baseline $baselineMetrics.EnabledConditionalAccessCount -Current $currentMetrics.EnabledConditionalAccessCount -Direction HigherIsBetter
    New-MetricComparison -Metric 'GlobalAdminCount' -Baseline $baselineMetrics.GlobalAdminCount -Current $currentMetrics.GlobalAdminCount -Direction LowerIsBetter
    New-MetricComparison -Metric 'UnverifiedDomainCount' -Baseline $baselineMetrics.UnverifiedDomainCount -Current $currentMetrics.UnverifiedDomainCount -Direction LowerIsBetter
    New-MetricComparison -Metric 'MaxLicenseUtilizationPercent' -Baseline $baselineMetrics.MaxLicenseUtilizationPercent -Current $currentMetrics.MaxLicenseUtilizationPercent -Direction LowerIsBetter
    New-MetricComparison -Metric 'StaleDevicePercent' -Baseline $baselineMetrics.StaleDevicePercent -Current $currentMetrics.StaleDevicePercent -Direction LowerIsBetter
)

$outputContext = Resolve-ArrayaSnapshotOutputContext -PrimaryInputPath $current.Path -OutputFolder $OutputFolder -OutputPrefix $OutputPrefix
$OutputFolder = $outputContext.OutputFolder
$OutputPrefix = $outputContext.OutputPrefix

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
