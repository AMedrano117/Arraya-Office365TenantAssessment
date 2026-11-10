function Convert-ArrayaToBoolean {
    <#
    .SYNOPSIS
        Normalizes a loosely-typed value to a boolean.

    .DESCRIPTION
        Snapshot round-trips and Graph payloads represent booleans inconsistently: as real
        booleans, as 'True'/'False' strings, as 'Enabled'/'Disabled', and as 1/0. Anything
        that is not recognizably true or false returns $null rather than $false, because a
        value that was never collected is not a value that was collected as false. Callers
        distinguish the two by comparing against $true or $false explicitly.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return [bool]$Value }

    switch -Regex (([string]$Value).Trim().ToLowerInvariant()) {
        '^(true|yes|enabled|on|1)$' { return $true }
        '^(false|no|disabled|off|0)$' { return $false }
        default { return $null }
    }
}
