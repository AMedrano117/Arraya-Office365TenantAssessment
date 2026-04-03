function Import-ArrayaTenantSnapshotContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Export', 'ImprovementPlan', 'Collection')]
        [string]$Purpose = 'Export'
    )

    function Resolve-SnapshotPathFromInput {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$InputPath
        )

        if (-not (Test-Path -Path $InputPath)) {
            throw "Snapshot JSON not found: $InputPath"
        }

        $resolvedInputPath = (Resolve-Path -Path $InputPath).Path
        if ($resolvedInputPath -notmatch '\.manifest\.json$') {
            return $resolvedInputPath
        }

        $manifest = Get-Content -Raw -Path $resolvedInputPath | ConvertFrom-Json -AsHashtable
        if (-not ($manifest -is [System.Collections.IDictionary])) {
            throw "Manifest did not deserialize to a hashtable: $resolvedInputPath"
        }

        $artifactRows = @()
        if ($manifest.Contains('Artifacts')) {
            $artifactRows = Convert-ArrayaObjectToArray $manifest['Artifacts']
        }

        $jsonArtifact = $artifactRows |
            Where-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('Type')) -in @('Assessment Snapshot JSON', 'JSON') } |
            Select-Object -First 1

        if (-not $jsonArtifact) {
            throw "Manifest does not contain a JSON snapshot artifact. Use a snapshot JSON directly or rerun with a JSON-enabled output profile: $resolvedInputPath"
        }

        $artifactPath = [string](Get-ArrayaObjectValue -Object $jsonArtifact -Names @('Path'))
        if ([string]::IsNullOrWhiteSpace($artifactPath)) {
            throw "Manifest JSON artifact is missing its path value: $resolvedInputPath"
        }

        $resolvedArtifactPath = if ([System.IO.Path]::IsPathRooted($artifactPath)) {
            $artifactPath
        }
        else {
            Join-Path -Path (Split-Path -Path $resolvedInputPath -Parent) -ChildPath $artifactPath
        }

        if (-not (Test-Path -Path $resolvedArtifactPath)) {
            throw "Manifest JSON artifact was not found on disk: $resolvedArtifactPath"
        }

        return (Resolve-Path -Path $resolvedArtifactPath).Path
    }

    $resolvedPath = Resolve-SnapshotPathFromInput -InputPath $Path
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
