function Import-Office365CustomLocal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$RepoRoot = $PSScriptRoot,
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredCommands = @()
    )

    $resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\shared\Resolve-ArrayaRepoRoot.ps1'))
    if (-not (Test-Path -Path $resolveRepoRootHelperPath)) {
        throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
    }
    . $resolveRepoRootHelperPath

    $resolvedRepoRoot = Resolve-ArrayaRepoRoot -StartPath $RepoRoot
    $commonManifestPath = Join-Path $resolvedRepoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
    $resolvedCommonManifestPath = (Resolve-Path -Path $commonManifestPath).Path

    $loadedCommonModule = Get-Module -Name 'Arraya.M365.Common' -ErrorAction SilentlyContinue | Select-Object -First 1
    $requiredCommonCommands = @('Import-ArrayaOffice365CustomLocal', 'Get-ArrayaAssessmentOutputRoot')
    $missingCommonCommands = @(
        $requiredCommonCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
    )

    if (
        -not $loadedCommonModule -or
        $loadedCommonModule.Path -ne $resolvedCommonManifestPath -or
        $missingCommonCommands.Count -gt 0
    ) {
        Import-Module -Name $resolvedCommonManifestPath -Force -DisableNameChecking -ErrorAction Stop
    }

    Import-ArrayaOffice365CustomLocal -RepoRoot $resolvedRepoRoot -RequiredCommands $RequiredCommands
}
