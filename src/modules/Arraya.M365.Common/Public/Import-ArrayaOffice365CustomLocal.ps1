function Import-ArrayaOffice365CustomLocal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$RepoRoot,
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredCommands = @()
    )

    if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
        $modulePath = Split-Path -Path $PSScriptRoot -Parent
        $resolvedRepoRoot = Resolve-Path -Path (Join-Path $modulePath '..\..\..')
        $RepoRoot = $resolvedRepoRoot.Path
    } else {
        $RepoRoot = (Resolve-Path -Path $RepoRoot).Path
    }

    $supportedVersion = '1.2.1'
    $manifestPathsToCheck = @(
        (Join-Path $RepoRoot "src/vendor/Office365Custom/$supportedVersion/Office365Custom.psd1"),
        (Join-Path $RepoRoot "Office365Custom/$supportedVersion/Office365Custom.psd1"),
        (Join-Path $RepoRoot 'src/vendor/Office365Custom/Office365Custom.psd1'),
        (Join-Path $RepoRoot 'Office365Custom/Office365Custom.psd1')
    )
    $manifestPath = $manifestPathsToCheck |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace([string]$manifestPath)) {
        throw "Could not find supported Office365Custom version $supportedVersion. Checked: $($manifestPathsToCheck -join ', ')"
    }
    $resolvedManifestPath = (Resolve-Path -LiteralPath $manifestPath -ErrorAction Stop).Path
    $manifestData = Import-PowerShellDataFile -LiteralPath $resolvedManifestPath -ErrorAction Stop
    if ([version]$manifestData.ModuleVersion -ne [version]$supportedVersion) {
        throw "Office365Custom manifest version '$($manifestData.ModuleVersion)' does not match supported version '$supportedVersion': $resolvedManifestPath"
    }

    $officeModule = Get-Module -Name 'Office365Custom' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($officeModule) {
        $loadedModuleBase = $null
        try {
            $loadedModuleBase = [System.IO.Path]::GetFullPath([string]$officeModule.ModuleBase)
        }
        catch {
            $loadedModuleBase = [string]$officeModule.ModuleBase
        }

        $targetModuleBase = Split-Path -Path $resolvedManifestPath -Parent
        if (-not [string]::Equals($loadedModuleBase, $targetModuleBase, [System.StringComparison]::OrdinalIgnoreCase)) {
            Remove-Module -Name 'Office365Custom' -Force -ErrorAction SilentlyContinue
        }
    }

    # Force import globally so code changes in the repo are reflected in the current session.
    Import-Module -Name $resolvedManifestPath -Global -Force -DisableNameChecking -ErrorAction Stop

    if ($RequiredCommands.Count -gt 0) {
        $missingAfterImport = @(
            $RequiredCommands | Where-Object {
                -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
            }
        )

        if ($missingAfterImport.Count -gt 0) {
            throw "Office365Custom import succeeded, but required commands are missing: $($missingAfterImport -join ', ')"
        }
    }

    return (Get-Module -Name 'Office365Custom' -ErrorAction SilentlyContinue)
}
