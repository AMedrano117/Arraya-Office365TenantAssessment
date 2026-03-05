function Convert-ArrayaToDate {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $dateValue = [datetime]::MinValue
    if ([datetime]::TryParse($Value.ToString(), [ref]$dateValue)) {
        return $dateValue
    }

    return $null
}
