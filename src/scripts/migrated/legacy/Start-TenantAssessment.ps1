[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('M365', 'AD', 'Graph', 'Improve', 'Compare')]
    [string]$Action
)

function Resolve-ArrayaRepoRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$StartPath
    )

    $candidate = (Resolve-Path -Path $StartPath).Path
    while ($true) {
        $launcherPath = Join-Path $candidate 'src\scripts\operations\Start-M365TenantAssessment.ps1'
        if (Test-Path -Path $launcherPath) {
            return $candidate
        }

        $parent = Split-Path -Path $candidate -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $candidate) {
            break
        }
        $candidate = $parent
    }

    throw "Could not resolve repository root from: $StartPath"
}

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
