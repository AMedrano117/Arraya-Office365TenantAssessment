[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssessmentJsonPath,
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder,
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
$runnerManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
Import-Module -Name $runnerManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop

$invokeParams = @{ AssessmentJsonPath = $AssessmentJsonPath }
if ($PSBoundParameters.ContainsKey('OutputFolder')) {
    $runStamp = Get-Date -Format 'yyyy-MM-dd-HHmm'
    $invokeParams.OutputFolder = Join-Path -Path $OutputFolder -ChildPath $runStamp
}
if ($PSBoundParameters.ContainsKey('UseGraphFallback')) { $invokeParams.UseGraphFallback = $UseGraphFallback }
if ($PSBoundParameters.ContainsKey('IncludeLegacyArtifacts')) { $invokeParams.IncludeLegacyArtifacts = $IncludeLegacyArtifacts }
if ($PSBoundParameters.ContainsKey('PassThru')) { $invokeParams.PassThru = $PassThru }
Invoke-M365ImprovementPlan @invokeParams
