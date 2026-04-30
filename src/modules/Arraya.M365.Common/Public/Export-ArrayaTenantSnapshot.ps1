function Export-ArrayaTenantSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Snapshot,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    function Write-ArrayaJsonValue {
        param(
            [Parameter(Mandatory = $true)]
            [System.Text.Json.Utf8JsonWriter]$Writer,
            [Parameter(Mandatory = $false)]
            $Value,
            [Parameter(Mandatory = $false)]
            [int]$Depth = 0,
            [Parameter(Mandatory = $false)]
            [System.Collections.Generic.HashSet[int]]$Visited
        )

        if ($null -eq $Value) {
            $Writer.WriteNullValue()
            return
        }

        if ($null -eq $Visited) {
            $Visited = [System.Collections.Generic.HashSet[int]]::new()
        }

        if ($Depth -ge 25) {
            $Writer.WriteStringValue('[MaxDepthExceeded]')
            return
        }

        $valueType = $Value.GetType()
        $isReferenceType = -not $valueType.IsValueType -and $Value -isnot [string]
        $referenceId = $null
        if ($isReferenceType) {
            $referenceId = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Value)
            if (-not $Visited.Add($referenceId)) {
                $Writer.WriteStringValue('[CircularReference]')
                return
            }
        }

        try {
            if ($Value -is [string] -or $Value -is [char]) {
                $Writer.WriteStringValue([string]$Value)
                return
            }

            if ($Value -is [bool]) {
                $Writer.WriteBooleanValue([bool]$Value)
                return
            }

            if ($Value -is [byte] -or $Value -is [sbyte] -or $Value -is [int16] -or $Value -is [uint16] -or $Value -is [int32]) {
                $Writer.WriteNumberValue([int]$Value)
                return
            }

            if ($Value -is [uint32] -or $Value -is [int64]) {
                $Writer.WriteNumberValue([long]$Value)
                return
            }

            if ($Value -is [uint64]) {
                $Writer.WriteNumberValue([decimal]$Value)
                return
            }

            if ($Value -is [single] -or $Value -is [double]) {
                $numberValue = [double]$Value
                if ([double]::IsNaN($numberValue) -or [double]::IsInfinity($numberValue)) {
                    $Writer.WriteStringValue([string]$Value)
                }
                else {
                    $Writer.WriteNumberValue($numberValue)
                }
                return
            }

            if ($Value -is [decimal]) {
                $Writer.WriteNumberValue([decimal]$Value)
                return
            }

            if ($Value -is [datetime] -or $Value -is [datetimeoffset]) {
                $Writer.WriteStringValue(([datetime]$Value).ToString('o'))
                return
            }

            if ($Value -is [timespan] -or $Value -is [guid] -or $Value -is [uri] -or $Value -is [version] -or $Value -is [enum]) {
                $Writer.WriteStringValue($Value.ToString())
                return
            }

            if ($Value -is [securestring]) {
                $Writer.WriteStringValue('[SecureString]')
                return
            }

            if ($Value -is [System.Management.Automation.SwitchParameter]) {
                $Writer.WriteBooleanValue([bool]$Value)
                return
            }

            if ($Value -is [System.Collections.IDictionary]) {
                $Writer.WriteStartObject()
                foreach ($key in $Value.Keys) {
                    $Writer.WritePropertyName([string]$key)
                    Write-ArrayaJsonValue -Writer $Writer -Value $Value[$key] -Depth ($Depth + 1) -Visited $Visited
                }
                $Writer.WriteEndObject()
                return
            }

            if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
                $Writer.WriteStartArray()
                foreach ($item in $Value) {
                    Write-ArrayaJsonValue -Writer $Writer -Value $item -Depth ($Depth + 1) -Visited $Visited
                }
                $Writer.WriteEndArray()
                return
            }

            $serializableProperties = @(
                $Value.PSObject.Properties |
                    Where-Object {
                        $_.MemberType -in @('NoteProperty', 'AliasProperty') -and
                        $_.Name -ne 'SyncRoot'
                    }
            )

            if ($serializableProperties.Count -gt 0) {
                $Writer.WriteStartObject()
                foreach ($property in $serializableProperties) {
                    $Writer.WritePropertyName([string]$property.Name)
                    try {
                        Write-ArrayaJsonValue -Writer $Writer -Value $property.Value -Depth ($Depth + 1) -Visited $Visited
                    }
                    catch {
                        $Writer.WriteStringValue("[PropertyReadError] $($_.Exception.Message)")
                    }
                }
                $Writer.WriteEndObject()
                return
            }

            $Writer.WriteStringValue($Value.ToString())
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

    $resolvedPath = [System.IO.Path]::GetFullPath($Path)
    $resolvedDirectory = Split-Path -Path $resolvedPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($resolvedDirectory) -and -not (Test-Path -Path $resolvedDirectory)) {
        $null = New-Item -ItemType Directory -Path $resolvedDirectory -Force
    }

    $normalized = [ordered]@{}
    foreach ($key in $Snapshot.Keys) {
        $normalized[[string]$key] = $Snapshot[$key]
    }
    if (-not $normalized.Contains('SchemaVersion')) {
        $normalized['SchemaVersion'] = 2
    }

    $stream = $null
    $writer = $null
    $writeSucceeded = $false
    try {
        $stream = [System.IO.File]::Create($resolvedPath)
        $jsonWriterOptions = [System.Text.Json.JsonWriterOptions]::new()
        $jsonWriterOptions.Indented = $true
        $writer = [System.Text.Json.Utf8JsonWriter]::new($stream, $jsonWriterOptions)
        $visited = [System.Collections.Generic.HashSet[int]]::new()

        Write-ArrayaJsonValue -Writer $writer -Value $normalized -Visited $visited
        $writer.Flush()
        $stream.Flush()
        $writeSucceeded = $true
    }
    finally {
        if ($writer) {
            $writer.Dispose()
        }
        if ($stream) {
            $stream.Dispose()
        }
        if (-not $writeSucceeded -and (Test-Path -Path $resolvedPath -PathType Leaf)) {
            Remove-Item -Path $resolvedPath -Force -ErrorAction SilentlyContinue
        }
    }
}
