function Convert-ArrayaToNumber {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $numeric = 0.0
    if ([double]::TryParse($Value.ToString(), [ref]$numeric)) {
        return [double]$numeric
    }

    return $null
}
