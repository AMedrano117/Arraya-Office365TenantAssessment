function Export-ArrayaTenantSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Snapshot,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    function ConvertTo-ArrayaJsonFriendlyValue {
        param(
            [Parameter(Mandatory = $false)]
            $Value,
            [Parameter(Mandatory = $false)]
            [int]$Depth = 0,
            [Parameter(Mandatory = $false)]
            [System.Collections.Generic.HashSet[int]]$Visited
        )

        if ($null -eq $Value) {
            return $null
        }

        if ($null -eq $Visited) {
            $Visited = [System.Collections.Generic.HashSet[int]]::new()
        }

        if ($Depth -ge 25) {
            return '[MaxDepthExceeded]'
        }

        $valueType = $Value.GetType()
        $isReferenceType = -not $valueType.IsValueType -and $Value -isnot [string]
        $referenceId = $null
        if ($isReferenceType) {
            $referenceId = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Value)
            if (-not $Visited.Add($referenceId)) {
                return '[CircularReference]'
            }
        }

        try {
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
                return ([datetime]$Value).ToString('o')
            }

            if ($Value -is [timespan] -or $Value -is [guid] -or $Value -is [uri] -or $Value -is [version]) {
                return $Value.ToString()
            }

            if ($Value -is [enum]) {
                return $Value.ToString()
            }

            if ($Value -is [securestring]) {
                return '[SecureString]'
            }

            if ($Value -is [System.Management.Automation.SwitchParameter]) {
                return [bool]$Value
            }

            if ($Value -is [System.Collections.IDictionary]) {
                $result = [ordered]@{}
                foreach ($key in $Value.Keys) {
                    $result[[string]$key] = ConvertTo-ArrayaJsonFriendlyValue -Value $Value[$key] -Depth ($Depth + 1) -Visited $Visited
                }
                return $result
            }

            if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
                $items = New-Object System.Collections.Generic.List[object]
                foreach ($item in $Value) {
                    $items.Add((ConvertTo-ArrayaJsonFriendlyValue -Value $item -Depth ($Depth + 1) -Visited $Visited))
                }
                return $items.ToArray()
            }

            $serializableProperties = @(
                $Value.PSObject.Properties |
                    Where-Object {
                        $_.MemberType -in @('NoteProperty', 'AliasProperty') -and
                        $_.Name -ne 'SyncRoot'
                    }
            )

            if ($serializableProperties.Count -gt 0) {
                $result = [ordered]@{}
                foreach ($property in $serializableProperties) {
                    try {
                        $result[$property.Name] = ConvertTo-ArrayaJsonFriendlyValue -Value $property.Value -Depth ($Depth + 1) -Visited $Visited
                    }
                    catch {
                        $result[$property.Name] = "[PropertyReadError] $($_.Exception.Message)"
                    }
                }
                return $result
            }

            return $Value.ToString()
        }
        finally {
            if ($isReferenceType -and $null -ne $referenceId) {
                $Visited.Remove($referenceId) | Out-Null
            }
        }
    }

    $validation = Test-ArrayaTenantSnapshot -Snapshot $Snapshot -Purpose Export
    if (-not $validation.Valid) {
        throw ("Snapshot validation failed: {0}" -f ($validation.Errors -join '; '))
    }

    $normalized = [ordered]@{}
    foreach ($key in $Snapshot.Keys) {
        $normalized[[string]$key] = $Snapshot[$key]
    }
    if (-not $normalized.Contains('SchemaVersion')) {
        $normalized['SchemaVersion'] = 2
    }

    $visited = [System.Collections.Generic.HashSet[int]]::new()
    $jsonFriendly = ConvertTo-ArrayaJsonFriendlyValue -Value $normalized -Visited $visited

    $jsonOptions = [System.Text.Json.JsonSerializerOptions]::new()
    $jsonOptions.WriteIndented = $true
    $jsonOptions.ReferenceHandler = [System.Text.Json.Serialization.ReferenceHandler]::IgnoreCycles
    $json = [System.Text.Json.JsonSerializer]::Serialize($jsonFriendly, $jsonOptions)
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.UTF8Encoding]::new($false))
}
