function Import-ArrayaTenantSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [switch]$SkipValidation
    )

    if (-not (Test-Path -Path $Path)) {
        return $null
    }

    $jsonRoot = Get-Content -Raw -Path $Path | ConvertFrom-Json -AsHashtable
    if (-not ($jsonRoot -is [hashtable])) {
        throw "Snapshot file did not deserialize to a hashtable: $Path"
    }

    $schemaVersion = 0
    if ($jsonRoot.Contains('SchemaVersion')) {
        try { $schemaVersion = [int]$jsonRoot['SchemaVersion'] } catch { $schemaVersion = 0 }
    }

    $snapshot = $null
    if ($schemaVersion -ge 2) {
        $snapshot = $jsonRoot
    }
    else {
        $legacyData = $null
        if ($jsonRoot.Contains('Data') -and ($jsonRoot['Data'] -is [System.Collections.IDictionary])) {
            $legacyData = @{}
            foreach ($entry in $jsonRoot['Data'].GetEnumerator()) {
                $legacyData[[string]$entry.Key] = $entry.Value
            }
        }
        else {
            $legacyData = @{}
            foreach ($entry in $jsonRoot.GetEnumerator()) {
                if ([string]$entry.Key -eq 'SchemaVersion') {
                    continue
                }
                $legacyData[[string]$entry.Key] = $entry.Value
            }
        }

        $metadata = [ordered]@{}
        if ($jsonRoot.Contains('GeneratedAt')) {
            $metadata['GeneratedAt'] = $jsonRoot['GeneratedAt']
        }
        if ($jsonRoot.Contains('OutputProfile')) {
            $metadata['OutputProfileLabel'] = $jsonRoot['OutputProfile']
        }
        if ($jsonRoot.Contains('ReportingMode')) {
            $metadata['ReportingMode'] = $jsonRoot['ReportingMode']
        }

        $snapshot = Convert-ArrayaLegacyTenantStatsToSnapshot -TenantStatsHash $legacyData -Metadata $metadata
    }

    if (-not $SkipValidation) {
        $validation = Test-ArrayaTenantSnapshot -Snapshot $snapshot -Purpose Export
        if (-not $validation.Valid) {
            throw ("Imported snapshot failed validation: {0}" -f ($validation.Errors -join '; '))
        }
    }

    return $snapshot
}
