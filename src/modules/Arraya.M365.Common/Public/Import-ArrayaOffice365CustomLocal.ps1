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

    $moduleRootsToCheck = @(
        (Join-Path $RepoRoot 'src/vendor/Office365Custom'),
        (Join-Path $RepoRoot 'Office365Custom')
    )
    $existingRoots = @($moduleRootsToCheck | Where-Object { Test-Path -Path $_ })
    if ($existingRoots.Count -eq 0) {
        throw "Could not locate Office365Custom under $RepoRoot."
    }

    $manifests = foreach ($root in $existingRoots) {
        Get-ChildItem -Path $root -Filter 'Office365Custom.psd1' -Recurse -File -ErrorAction SilentlyContinue
    }
    if (-not $manifests) {
        throw "Could not find Office365Custom module manifest under: $($existingRoots -join ', ')"
    }

    $latestManifest = $manifests |
        Sort-Object {
            try { [version]$_.Directory.Name }
            catch { [version]'0.0.0' }
        } -Descending |
        Select-Object -First 1

    $officeModule = Get-Module -Name 'Office365Custom' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($officeModule) {
        $loadedModulePath = $null
        try {
            $loadedModulePath = (Resolve-Path -Path $officeModule.Path -ErrorAction Stop).Path
        }
        catch {
            $loadedModulePath = $officeModule.Path
        }

        $targetModulePath = (Resolve-Path -Path $latestManifest.FullName).Path
        if ($loadedModulePath -ne $targetModulePath) {
            Remove-Module -Name 'Office365Custom' -Force -ErrorAction SilentlyContinue
        }
    }

    # Force import globally so code changes in the repo are reflected in the current session.
    Import-Module -Name $latestManifest.FullName -Global -Force -ErrorAction Stop

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
