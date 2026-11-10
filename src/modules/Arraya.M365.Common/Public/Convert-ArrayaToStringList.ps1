function Convert-ArrayaToStringList {
    <#
    .SYNOPSIS
        Normalizes a loosely-typed value to an array of non-empty strings.

    .DESCRIPTION
        Collected list values arrive as delimited strings in some code paths, as real arrays
        in others, and as dictionaries after a snapshot round-trip. Delimited strings are
        split on comma and semicolon; everything else is flattened through
        Convert-ArrayaObjectToArray. Empty and whitespace-only entries are dropped so callers
        can treat Count -eq 0 as "nothing recorded".
    #>
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
        return @(
            $Value -split '[,;]' |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
    }

    return @(
        (Convert-ArrayaObjectToArray $Value) |
            ForEach-Object { [string]$_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}
