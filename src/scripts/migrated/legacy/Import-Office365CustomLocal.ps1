function Import-Office365CustomLocal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$RepoRoot = $PSScriptRoot,
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredCommands = @()
    )

    function Test-ImportedModuleMatchesManifestPath {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [System.Management.Automation.PSModuleInfo]$Module,
            [Parameter(Mandatory = $true)]
            [string]$ManifestPath
        )

        if (-not $Module) {
            return $false
        }

        try {
            $resolvedManifestPath = (Resolve-Path -Path $ManifestPath -ErrorAction Stop).Path
        }
        catch {
            return $false
        }

        $candidatePaths = @(
            $Module.Path
            (Join-Path -Path $Module.ModuleBase -ChildPath ([System.IO.Path]::GetFileName($resolvedManifestPath)))
            $Module.ModuleBase
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

        foreach ($candidatePath in $candidatePaths) {
            try {
                $resolvedCandidatePath = [System.IO.Path]::GetFullPath($candidatePath)
            }
            catch {
                $resolvedCandidatePath = $candidatePath
            }

            if ([string]::Equals($resolvedCandidatePath, $resolvedManifestPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }

        $resolvedManifestDirectory = Split-Path -Path $resolvedManifestPath -Parent
        if (
            -not [string]::IsNullOrWhiteSpace([string]$Module.ModuleBase) -and
            [string]::Equals(
                ([System.IO.Path]::GetFullPath($Module.ModuleBase)),
                ([System.IO.Path]::GetFullPath($resolvedManifestDirectory)),
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            return $true
        }

        return $false
    }

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
        -not (Test-ImportedModuleMatchesManifestPath -Module $loadedCommonModule -ManifestPath $resolvedCommonManifestPath) -or
        $missingCommonCommands.Count -gt 0
    ) {
        Import-Module -Name $resolvedCommonManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
    }

    Import-ArrayaOffice365CustomLocal -RepoRoot $resolvedRepoRoot -RequiredCommands $RequiredCommands
}
