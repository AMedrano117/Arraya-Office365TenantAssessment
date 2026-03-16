function Convert-DataSizeToBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return [int64]0
    }

    if ($Value -is [int64] -or $Value -is [int32] -or $Value -is [double] -or $Value -is [decimal]) {
        try { return [int64]$Value } catch { return [int64]0 }
    }

    foreach ($propertyName in @('Bytes', 'ByteCount', 'Value')) {
        if ($Value.PSObject -and $Value.PSObject.Properties[$propertyName]) {
            $nestedValue = $Value.PSObject.Properties[$propertyName].Value
            if ($nestedValue -is [int64] -or $nestedValue -is [int32] -or $nestedValue -is [double] -or $nestedValue -is [decimal]) {
                try { return [int64]$nestedValue } catch {}
            }
        }
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [int64]0
    }

    if ($text -match '\((?<bytes>[0-9,]+)\s+bytes\)') {
        try { return [int64](($Matches['bytes'] -replace ',', '')) } catch {}
    }

    if ($text -match '^\s*(?<number>[0-9]+(?:\.[0-9]+)?)\s*(?<unit>KB|MB|GB|TB|PB)\b') {
        $number = [double]$Matches['number']
        $multiplier = switch ($Matches['unit'].ToUpperInvariant()) {
            'KB' { 1KB }
            'MB' { 1MB }
            'GB' { 1GB }
            'TB' { 1TB }
            'PB' { 1PB }
            default { 1 }
        }
        try { return [int64]($number * $multiplier) } catch { return [int64]0 }
    }

    return [int64]0
}
