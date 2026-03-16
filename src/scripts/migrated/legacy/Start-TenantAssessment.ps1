[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('M365', 'AD', 'Graph', 'Improve', 'Compare')]
    [string]$Action
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

if ($PSBoundParameters.ContainsKey('Action')) {
    & $newLauncher -Action $Action
} else {
    & $newLauncher
}
