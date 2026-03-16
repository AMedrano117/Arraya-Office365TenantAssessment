function ConvertTo-ExportFriendlyRecord {
    param($InputObject)

    if ($null -eq $InputObject) {
        return [pscustomobject]@{ Value = $null }
    }

    if ($InputObject -is [hashtable] -or $InputObject -is [System.Collections.Specialized.OrderedDictionary]) {
        $result = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $result[[string]$key] = ConvertTo-ExportFriendlyValue -Value $InputObject[$key]
        }
        return [pscustomobject]$result
    }

    $properties = @(
        $InputObject.PSObject.Properties |
            Where-Object { $_.MemberType -in @('NoteProperty', 'AliasProperty', 'Property') }
    )

    if ($properties.Count -eq 0) {
        return [pscustomobject]@{ Value = ConvertTo-ExportFriendlyValue -Value $InputObject }
    }

    $result = [ordered]@{}
    foreach ($property in $properties) {
        $result[$property.Name] = ConvertTo-ExportFriendlyValue -Value $property.Value
    }

    return [pscustomobject]$result
}