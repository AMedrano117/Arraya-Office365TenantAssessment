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

$resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\shared\Resolve-ArrayaRepoRoot.ps1'))
if (-not (Test-Path -Path $resolveRepoRootHelperPath)) {
    throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
}
. $resolveRepoRootHelperPath

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$legacyScriptPath = Join-Path $repoRoot 'src\scripts\migrated\legacy\Compare-M365TenantAssessmentSnapshots.ps1'
if (-not (Test-Path -Path $legacyScriptPath)) {
    throw "Assessment-comparison implementation script not found: $legacyScriptPath"
}

$commonManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
if (-not (Get-Module -Name 'Arraya.M365.Common' -ErrorAction SilentlyContinue)) {
    Import-Module -Name $commonManifestPath -WarningAction SilentlyContinue -ErrorAction Stop
}

$baselineSnapshot = Import-ArrayaTenantSnapshot -Path $BaselineJsonPath -SkipValidation
$currentSnapshot  = Import-ArrayaTenantSnapshot -Path $CurrentJsonPath  -SkipValidation

if ($null -eq $baselineSnapshot) { throw "Baseline snapshot could not be loaded: $BaselineJsonPath" }
if ($null -eq $currentSnapshot)  { throw "Current snapshot could not be loaded: $CurrentJsonPath" }

$baselineVersion = if ($baselineSnapshot.Contains('SchemaVersion')) { [int]$baselineSnapshot['SchemaVersion'] } else { 0 }
$currentVersion  = if ($currentSnapshot.Contains('SchemaVersion'))  { [int]$currentSnapshot['SchemaVersion'] }  else { 0 }

if ($baselineVersion -ne $currentVersion) {
    throw "Schema version mismatch: baseline is v$baselineVersion, current is v$currentVersion. Both snapshots must use the same schema version for comparison."
}

$invokeParams = @{}
if ($PSBoundParameters.ContainsKey('BaselineJsonPath')) { $invokeParams.BaselineJsonPath = $BaselineJsonPath }
if ($PSBoundParameters.ContainsKey('CurrentJsonPath')) { $invokeParams.CurrentJsonPath = $CurrentJsonPath }
if ($PSBoundParameters.ContainsKey('OutputFolder')) { $invokeParams.OutputFolder = $OutputFolder }
if ($PSBoundParameters.ContainsKey('OutputPrefix')) { $invokeParams.OutputPrefix = $OutputPrefix }
if ($PSBoundParameters.ContainsKey('StaleDeviceDays')) { $invokeParams.StaleDeviceDays = $StaleDeviceDays }
if ($PSBoundParameters.ContainsKey('PassThru')) { $invokeParams.PassThru = $PassThru }

& {
    Set-StrictMode -Off
    & $legacyScriptPath @invokeParams
}
