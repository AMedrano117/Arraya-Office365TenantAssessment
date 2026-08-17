<#
.SYNOPSIS
    Deprecated. Use tools/install-dependencies.ps1 instead.

.DESCRIPTION
    Kept as a shim so existing runbooks and muscle memory keep working. This script used to
    install an unpinned subset of the required modules, which could leave a machine partially
    provisioned; the manifest-driven installer covers the full set.
#>
[CmdletBinding()]
param()

Write-Warning 'install-required-modules.ps1 is deprecated. Use ./tools/install-dependencies.ps1 -Scope All instead.'

& (Join-Path -Path $PSScriptRoot -ChildPath 'install-dependencies.ps1') -Scope All
