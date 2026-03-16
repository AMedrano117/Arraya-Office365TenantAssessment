function Convert-ArrayaToNumber {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value,
        [Parameter(Mandatory = $false)]
        [switch]$AsInt64
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = $text -replace ',', ''
    $styles = [System.Globalization.NumberStyles]::Any
    $culture = [System.Globalization.CultureInfo]::InvariantCulture

    $numeric = 0.0
    if ([double]::TryParse($text, $styles, $culture, [ref]$numeric)) {
        if ($AsInt64) {
            return [int64][math]::Round($numeric, 0)
        }
        return [double]$numeric
    }

    return $null
}
