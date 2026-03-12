function Convert-ArrayaSnapshotToLegacyTenantStatsHash {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Snapshot
    )

    $legacy = @{}
    $schemaVersion = 0
    if ($Snapshot.Contains('SchemaVersion')) {
        try { $schemaVersion = [int]$Snapshot['SchemaVersion'] } catch { $schemaVersion = 0 }
    }

    if ($schemaVersion -le 1) {
        if ($Snapshot.Contains('Data') -and ($Snapshot['Data'] -is [System.Collections.IDictionary])) {
            foreach ($entry in $Snapshot['Data'].GetEnumerator()) {
                $legacy[[string]$entry.Key] = $entry.Value
            }
            return $legacy
        }

        foreach ($entry in $Snapshot.GetEnumerator()) {
            if ([string]$entry.Key -eq 'SchemaVersion') {
                continue
            }
            $legacy[[string]$entry.Key] = $entry.Value
        }
        return $legacy
    }

    if ($Snapshot.Contains('Data') -and ($Snapshot['Data'] -is [System.Collections.IDictionary])) {
        $knownDomains = @('Exchange', 'Identity', 'Collaboration', 'Security', 'Tenant', 'Other')
        foreach ($domain in $knownDomains) {
            if (-not $Snapshot['Data'].Contains($domain)) {
                continue
            }
            $domainValue = $Snapshot['Data'][$domain]
            if (-not ($domainValue -is [System.Collections.IDictionary])) {
                continue
            }

            foreach ($entry in $domainValue.GetEnumerator()) {
                $key = [string]$entry.Key
                if ($legacy.Contains($key)) {
                    $legacy["$domain-$key"] = $entry.Value
                }
                else {
                    $legacy[$key] = $entry.Value
                }
            }
        }
    }

    if ($Snapshot.Contains('Derived') -and ($Snapshot['Derived'] -is [System.Collections.IDictionary])) {
        foreach ($entry in $Snapshot['Derived'].GetEnumerator()) {
            $legacy[[string]$entry.Key] = $entry.Value
        }
    }

    if (
        $Snapshot.Contains('Diagnostics') -and
        $Snapshot['Diagnostics'] -is [System.Collections.IDictionary] -and
        $Snapshot['Diagnostics'].Contains('CollectorStats') -and
        $Snapshot['Diagnostics']['CollectorStats'] -is [System.Collections.IDictionary]
    ) {
        foreach ($entry in $Snapshot['Diagnostics']['CollectorStats'].GetEnumerator()) {
            if (-not $legacy.Contains([string]$entry.Key)) {
                $legacy[[string]$entry.Key] = $entry.Value
            }
        }
    }

    if (
        -not $legacy.Contains('TenantInfo') -and
        $Snapshot.Contains('Metadata') -and
        $Snapshot['Metadata'] -is [System.Collections.IDictionary] -and
        $Snapshot['Metadata'].Contains('Tenant') -and
        $Snapshot['Metadata']['Tenant'] -is [System.Collections.IDictionary]
    ) {
        $legacy['TenantInfo'] = [PSCustomObject]$Snapshot['Metadata']['Tenant']
    }

    return $legacy
}
