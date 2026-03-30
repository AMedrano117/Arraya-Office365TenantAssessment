[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssessmentJsonPath,
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder,
    [Parameter(Mandatory = $false)]
    [string]$OutputPrefix,
    [Parameter(Mandatory = $false)]
    [ValidateRange(30, 1095)]
    [int]$StaleDeviceDays = 180,
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$MaxGlobalAdmins = 5,
    [Parameter(Mandatory = $false)]
    [Alias('LiveRefresh')]
    [switch]$UseGraphFallback,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyArtifacts,
    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

$resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\shared\Resolve-ArrayaRepoRoot.ps1'))
if (-not (Test-Path -Path $resolveRepoRootHelperPath)) {
    throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
}
. $resolveRepoRootHelperPath

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$legacyScriptPath = Join-Path $repoRoot 'src\scripts\migrated\legacy\New-M365TenantImprovementPlan.ps1'
if (-not (Test-Path -Path $legacyScriptPath)) {
    throw "Improvement-plan implementation script not found: $legacyScriptPath"
}

$invokeParams = @{}
if ($PSBoundParameters.ContainsKey('AssessmentJsonPath')) { $invokeParams.AssessmentJsonPath = $AssessmentJsonPath }
if ($PSBoundParameters.ContainsKey('OutputFolder')) { $invokeParams.OutputFolder = $OutputFolder }
if ($PSBoundParameters.ContainsKey('OutputPrefix')) { $invokeParams.OutputPrefix = $OutputPrefix }
if ($PSBoundParameters.ContainsKey('StaleDeviceDays')) { $invokeParams.StaleDeviceDays = $StaleDeviceDays }
if ($PSBoundParameters.ContainsKey('MaxGlobalAdmins')) { $invokeParams.MaxGlobalAdmins = $MaxGlobalAdmins }
if ($PSBoundParameters.ContainsKey('UseGraphFallback')) { $invokeParams.UseGraphFallback = $UseGraphFallback }
if ($PSBoundParameters.ContainsKey('IncludeLegacyArtifacts')) { $invokeParams.IncludeLegacyArtifacts = $IncludeLegacyArtifacts }
if ($PSBoundParameters.ContainsKey('PassThru')) { $invokeParams.PassThru = $PassThru }

& {
    Set-StrictMode -Off
    & $legacyScriptPath @invokeParams
}
