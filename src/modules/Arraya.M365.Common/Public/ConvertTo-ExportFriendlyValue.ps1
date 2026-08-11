function ConvertTo-ExportFriendlyValue {
    param(
        $Value,
        [int]$Depth = 0
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Depth -ge 3) {
        return '[Nested]'
    }

    if (
        $Value -is [string] -or
        $Value -is [char] -or
        $Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [sbyte] -or
        $Value -is [int16] -or
        $Value -is [uint16] -or
        $Value -is [int32] -or
        $Value -is [uint32] -or
        $Value -is [int64] -or
        $Value -is [uint64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]
    ) {
        return $Value
    }

    if ($Value -is [datetime] -or $Value -is [datetimeoffset]) {
        return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm:ss')
    }

    if ($Value -is [timespan] -or $Value -is [guid] -or $Value -is [uri] -or $Value -is [version] -or $Value -is [enum]) {
        return $Value.ToString()
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $pairs = @()
        foreach ($key in $Value.Keys) {
            $pairs += "$key=$(ConvertTo-ExportFriendlyValue -Value $Value[$key] -Depth ($Depth + 1))"
        }
        return ($pairs -join '; ')
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @()
        foreach ($item in $Value) {
            if ($null -eq $item) { continue }

            if (
                $item -is [string] -or
                $item -is [ValueType]
            ) {
                $items += $item.ToString()
                continue
            }

            $identityValue = $null
            foreach ($identityProperty in @('DisplayName', 'Name', 'Title', 'Domain', 'UserPrincipalName', 'Mail', 'AppId', 'Id', 'SkuFriendlyName', 'SkuPartNumber', 'Value')) {
                $property = $item.PSObject.Properties[$identityProperty]
                if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                    $identityValue = $property.Value
                    break
                }
            }

            if ($null -ne $identityValue) {
                $items += $identityValue.ToString()
            } else {
                $items += (ConvertTo-ExportFriendlyValue -Value $item -Depth ($Depth + 1))
            }
        }
        return ($items -join '; ')
    }

    $properties = @(
        $Value.PSObject.Properties |
            Where-Object { $_.MemberType -in @('NoteProperty', 'AliasProperty') }
    )

    if ($properties.Count -gt 0) {
        $pairs = @()
        foreach ($property in $properties) {
            $pairs += "$($property.Name)=$(ConvertTo-ExportFriendlyValue -Value $property.Value -Depth ($Depth + 1))"
        }
        return ($pairs -join '; ')
    }

    # Graph SDK model objects expose their members as 'Property' rather than
    # 'NoteProperty', so the filter above finds nothing and ToString() falls back to the
    # class name. Emitting that would fill a worksheet column with
    # 'Microsoft.Graph.PowerShell.Models.MicrosoftGraphAuthentication' and no data.
    $text = $Value.ToString()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $typeName = $null
    try { $typeName = $Value.GetType().FullName } catch {}
    if ($typeName -and [string]::Equals($text, $typeName, [System.StringComparison]::Ordinal)) {
        return $null
    }

    if ($text -match '^(Microsoft\.Graph|Microsoft\.Online|System)\.[A-Za-z0-9_.`\[\]+]+$') {
        return $null
    }

    return $text
}