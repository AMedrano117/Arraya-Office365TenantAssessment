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

    # Collections must be collapsed before the property branch below. Reflecting over an
    # array surfaces its own .NET members (Length, LongLength, Rank, SyncRoot, IsReadOnly,
    # IsFixedSize, IsSynchronized, Count) and those become the worksheet columns.
    if (($InputObject -is [System.Collections.IEnumerable]) -and
        -not ($InputObject -is [string]) -and
        -not ($InputObject -is [System.Collections.IDictionary])) {
        return [pscustomobject]@{ Value = ConvertTo-ExportFriendlyValue -Value $InputObject }
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