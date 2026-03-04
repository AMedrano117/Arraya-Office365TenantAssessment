function Import-Office365CustomLocal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$RepoRoot = $PSScriptRoot,
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredCommands = @()
    )

    function Resolve-ArrayaRepoRoot {
        param(
            [Parameter(Mandatory = $true)]
            [string]$StartPath
        )

        $candidate = (Resolve-Path -Path $StartPath).Path
        while ($true) {
            $commonManifestPath = Join-Path $candidate 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
            if (Test-Path -Path $commonManifestPath) {
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
        Import-Module -Name $resolvedCommonManifestPath -Force -ErrorAction Stop
    }

    Import-ArrayaOffice365CustomLocal -RepoRoot $resolvedRepoRoot -RequiredCommands $RequiredCommands
}
