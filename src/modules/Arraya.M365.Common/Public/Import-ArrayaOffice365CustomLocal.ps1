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

    $officeModule = Get-Module -Name 'Office365Custom' -ErrorAction SilentlyContinue
    if ($officeModule) {
        $missingCommands = @(
            $RequiredCommands | Where-Object {
                -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
            }
        )

        if ($missingCommands.Count -gt 0) {
            throw "Office365Custom is loaded but missing required commands: $($missingCommands -join ', ')"
        }

        return $officeModule
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

    # Import globally so commands remain available after this helper function returns.
    Import-Module -Name $latestManifest.FullName -Global -ErrorAction Stop

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
