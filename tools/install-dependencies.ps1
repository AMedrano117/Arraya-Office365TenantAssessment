<#
.SYNOPSIS
    Installs or verifies the PowerShell modules declared in dependencies.psd1.

.DESCRIPTION
    Replaces the previous install-microsoft-modules.ps1 and install-required-modules.ps1
    scripts, which installed overlapping and unpinned "latest" module sets. Everything now
    comes from a single versioned manifest so a run on an engineer's laptop, a build agent,
    and a customer jump box resolve the same module versions.

.PARAMETER Scope
    Which dependency scopes to act on: Runtime, Onboarding, DevTools, or All.

.PARAMETER Validate
    Report what is missing or out of range without installing anything. Exits nonzero when
    the environment does not satisfy the manifest. Intended for the clean-machine CI job.

.PARAMETER Force
    Reinstall even when a satisfying version is already present.

.PARAMETER ManifestPath
    Path to the dependency manifest. Defaults to dependencies.psd1 at the repository root.

.EXAMPLE
    ./tools/install-dependencies.ps1 -Scope All

.EXAMPLE
    ./tools/install-dependencies.ps1 -Scope Runtime -Validate
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    'PSAvoidUsingWriteHost', '',
    Justification = 'This is an operator-facing installer; colored progress output is the intended interface, not pipeline data.'
)]
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Runtime', 'Onboarding', 'DevTools', 'All')]
    [string[]]$Scope = @('Runtime'),

    [Parameter(Mandatory = $false)]
    [switch]$Validate,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [string]$ManifestPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'dependencies.psd1'
}

if (-not (Test-Path -Path $ManifestPath)) {
    throw "Dependency manifest not found: $ManifestPath"
}

$manifest = Import-PowerShellDataFile -Path $ManifestPath

$requiredPSVersion = [version]$manifest.PowerShellVersion
if ($PSVersionTable.PSVersion -lt $requiredPSVersion) {
    throw "This repository requires PowerShell $requiredPSVersion or later (found $($PSVersionTable.PSVersion)). Run with 'pwsh'."
}

$selectedScopes = if ($Scope -contains 'All') {
    @('Runtime', 'Onboarding', 'DevTools')
}
else {
    @($Scope)
}

$selected = @(
    $manifest.Modules | Where-Object {
        $moduleScopes = @($_.Scope)
        @($moduleScopes | Where-Object { $selectedScopes -contains $_ }).Count -gt 0
    }
)

if ($selected.Count -eq 0) {
    throw "No modules matched the requested scope(s): $($selectedScopes -join ', ')"
}

Write-Host ("Dependency manifest: {0}" -f $ManifestPath) -ForegroundColor Cyan
Write-Host ("Scopes: {0} ({1} module(s))" -f ($selectedScopes -join ', '), $selected.Count) -ForegroundColor Cyan
Write-Host ("PowerShell: {0}" -f $PSVersionTable.PSVersion) -ForegroundColor Cyan
Write-Host ''

function Get-SatisfyingModule {
    param(
        [Parameter(Mandatory = $true)]
        $Requirement
    )

    $minimum = [version]$Requirement.MinimumVersion
    $maximum = [version]$Requirement.MaximumVersion

    return Get-Module -ListAvailable -Name $Requirement.Name |
        Where-Object { $_.Version -ge $minimum -and $_.Version -le $maximum } |
        Sort-Object Version -Descending |
        Select-Object -First 1
}

if (-not $Validate) {
    $gallery = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
    if ($gallery -and $gallery.InstallationPolicy -ne 'Trusted') {
        Write-Host 'Setting PSGallery to Trusted for this session...' -ForegroundColor DarkGray
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    }
}

$unsatisfied = New-Object System.Collections.Generic.List[string]

foreach ($requirement in $selected) {
    $name = [string]$requirement.Name
    $range = '{0} - {1}' -f $requirement.MinimumVersion, $requirement.MaximumVersion
    $satisfying = Get-SatisfyingModule -Requirement $requirement

    if ($satisfying -and -not $Force) {
        Write-Host ("  [ok]      {0} {1} (satisfies {2})" -f $name, $satisfying.Version, $range) -ForegroundColor Green
        continue
    }

    $installedAny = @(Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1)
    $currentText = if ($installedAny.Count -gt 0) { [string]$installedAny[0].Version } else { 'not installed' }

    if ($Validate) {
        Write-Host ("  [MISSING] {0} requires {1}, found {2}" -f $name, $range, $currentText) -ForegroundColor Red
        $unsatisfied.Add(("{0} (requires {1}, found {2})" -f $name, $range, $currentText)) | Out-Null
        continue
    }

    Write-Host ("  [install] {0} {1} (currently {2})" -f $name, $range, $currentText) -ForegroundColor Yellow

    $installParams = @{
        Name           = $name
        MinimumVersion = $requirement.MinimumVersion
        MaximumVersion = $requirement.MaximumVersion
        Scope          = 'CurrentUser'
        Repository     = 'PSGallery'
        AllowClobber   = $true
        ErrorAction    = 'Stop'
    }

    if ($Force) {
        $installParams.Force = $true
        $installParams.SkipPublisherCheck = $true
    }

    try {
        Install-Module @installParams
    }
    catch {
        Write-Host ("  [FAILED]  {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
        $unsatisfied.Add(("{0} (install failed: {1})" -f $name, $_.Exception.Message)) | Out-Null
        continue
    }

    $confirmed = Get-SatisfyingModule -Requirement $requirement
    if ($confirmed) {
        Write-Host ("  [ok]      {0} {1} installed" -f $name, $confirmed.Version) -ForegroundColor Green
    }
    else {
        Write-Host ("  [FAILED]  {0} still does not satisfy {1} after install" -f $name, $range) -ForegroundColor Red
        $unsatisfied.Add(("{0} (requires {1}, still unsatisfied)" -f $name, $range)) | Out-Null
    }
}

Write-Host ''

if ($unsatisfied.Count -gt 0) {
    Write-Host ("{0} dependency requirement(s) not satisfied:" -f $unsatisfied.Count) -ForegroundColor Red
    foreach ($item in $unsatisfied) {
        Write-Host ("  - {0}" -f $item) -ForegroundColor Red
    }

    if ($Validate) {
        Write-Host 'Run ./tools/install-dependencies.ps1 -Scope All to install them.' -ForegroundColor Yellow
    }

    exit 1
}

if ($Validate) {
    Write-Host 'All manifest dependencies are satisfied.' -ForegroundColor Green
}
else {
    Write-Host 'Dependency installation complete.' -ForegroundColor Green
}
