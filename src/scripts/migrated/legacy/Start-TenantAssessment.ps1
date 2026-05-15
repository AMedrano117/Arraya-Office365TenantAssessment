[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Full', 'Collect', 'Report', 'AD', 'Improve', 'Compare')]
    [string]$Action,
    [Parameter(Mandatory = $false)]
    [switch]$RunImprove,
    [Parameter(Mandatory = $false)]
    [switch]$SkipImprove,
    [Parameter(Mandatory = $false)]
    [string[]]$OutputProfile,
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,
    [Parameter(Mandatory = $false)]
    [Alias('LiveRefresh')]
    [switch]$UseGraphFallback,
    [Parameter(Mandatory = $false)]
    [string]$ImproveOutputFolder,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyArtifacts,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyAssessmentArtifacts
)

$resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\shared\Resolve-ArrayaRepoRoot.ps1'))
if (-not (Test-Path -Path $resolveRepoRootHelperPath)) {
    throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
}
. $resolveRepoRootHelperPath

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$newLauncher = Join-Path $repoRoot 'src\scripts\operations\Start-M365TenantAssessment.ps1'
if (-not (Test-Path -Path $newLauncher)) {
    throw "Launcher script not found: $newLauncher"
}

if ($PSBoundParameters.Count -gt 0) {
    & $newLauncher @PSBoundParameters
}
else {
    & $newLauncher
}
