function Import-ArrayaTenantSnapshotContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Export', 'ImprovementPlan', 'Collection')]
        [string]$Purpose = 'Export'
    )

    if (-not (Test-Path -Path $Path)) {
        throw "Snapshot JSON not found: $Path"
    }

    $resolvedPath = (Resolve-Path -Path $Path).Path
    $snapshot = Import-ArrayaTenantSnapshot -Path $resolvedPath
    $validation = Test-ArrayaTenantSnapshot -Snapshot $snapshot -Purpose $Purpose
    if (-not $validation.Valid) {
        throw "Snapshot failed validation for ${Purpose}: $($validation.Errors -join '; ')"
    }
    if ($validation.Warnings.Count -gt 0) {
        Write-Warning "Snapshot validation warnings for ${Purpose}: $($validation.Warnings -join '; ')"
    }

    $data = if ($snapshot.Contains('Data') -and ($snapshot['Data'] -is [System.Collections.IDictionary])) {
        $snapshot['Data']
    }
    else {
        $snapshot
    }

    $derived = @{}
    if ($snapshot.Contains('Derived') -and ($snapshot['Derived'] -is [System.Collections.IDictionary])) {
        foreach ($entry in $snapshot['Derived'].GetEnumerator()) {
            $derived[[string]$entry.Key] = $entry.Value
        }
    }

    $diagnostics = @{}
    if ($snapshot.Contains('Diagnostics') -and ($snapshot['Diagnostics'] -is [System.Collections.IDictionary])) {
        foreach ($entry in $snapshot['Diagnostics'].GetEnumerator()) {
            $diagnostics[[string]$entry.Key] = $entry.Value
        }
    }

    $generatedAt = $null
    if ($snapshot.Contains('Metadata') -and ($snapshot['Metadata'] -is [System.Collections.IDictionary])) {
        $generatedAt = Get-ArrayaObjectValue -Object $snapshot['Metadata'] -Names @('GeneratedAt')
    }

    return [PSCustomObject]@{
        Path        = $resolvedPath
        Snapshot    = $snapshot
        Data        = $data
        LegacyData  = Convert-ArrayaSnapshotToLegacyTenantStatsHash -Snapshot $snapshot
        Derived     = $derived
        Diagnostics = $diagnostics
        GeneratedAt = $generatedAt
    }
}
