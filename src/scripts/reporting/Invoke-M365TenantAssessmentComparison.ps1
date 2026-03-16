[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$BaselineJsonPath,
    [Parameter(Mandatory = $true)]
    [string]$CurrentJsonPath,
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder
)

$resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\shared\Resolve-ArrayaRepoRoot.ps1'))
if (-not (Test-Path -Path $resolveRepoRootHelperPath)) {
    throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
}
. $resolveRepoRootHelperPath

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$runnerManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
if (-not (Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue)) {
    Import-Module -Name $runnerManifestPath -ErrorAction Stop
}

$invokeParams = @{
    BaselineJsonPath = $BaselineJsonPath
    CurrentJsonPath  = $CurrentJsonPath
}
if ($PSBoundParameters.ContainsKey('OutputFolder')) { $invokeParams.OutputFolder = $OutputFolder }
Invoke-M365AssessmentComparison @invokeParams
