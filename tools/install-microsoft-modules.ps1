<#
.SYNOPSIS
    Deprecated. Use tools/install-dependencies.ps1 instead.

.DESCRIPTION
    Kept as a shim so existing runbooks and muscle memory keep working. Module versions are
    now declared in dependencies.psd1 rather than resolved as "latest" at install time.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDevTools,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

Write-Warning 'install-microsoft-modules.ps1 is deprecated. Use ./tools/install-dependencies.ps1 -Scope All instead.'

$scope = if ($IncludeDevTools) { @('Runtime', 'Onboarding', 'DevTools') } else { @('Runtime', 'Onboarding') }

$forward = @{ Scope = $scope }
if ($Force) {
    $forward.Force = $true
}

& (Join-Path -Path $PSScriptRoot -ChildPath 'install-dependencies.ps1') @forward
