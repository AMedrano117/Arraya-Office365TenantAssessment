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
