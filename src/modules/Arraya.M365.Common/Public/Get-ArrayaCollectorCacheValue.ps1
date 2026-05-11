function Get-ArrayaCollectorCacheValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Context,
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [Parameter(Mandatory = $false)]
        [switch]$CacheNameOnly
    )

    if ($null -eq $Context) {
        return $null
    }
    if (-not ($Context.PSObject.Properties['Runtime']) -or -not ($Context.Runtime -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName Runtime -NotePropertyValue ([ordered]@{}) -Force
    }
    if (-not $Context.Runtime.Contains('CollectorCache')) {
        $Context.Runtime['CollectorCache'] = [ordered]@{}
    }
    if (-not $Context.Runtime.Contains('CollectorCacheStats')) {
        $Context.Runtime['CollectorCacheStats'] = [ordered]@{
            Hits   = 0
            Misses = 0
            Writes = 0
        }
    }

    if ($Context.Runtime['CollectorCache'].Contains($Key)) {
        if (-not $CacheNameOnly) {
            $Context.Runtime['CollectorCacheStats']['Hits'] = [int]$Context.Runtime['CollectorCacheStats']['Hits'] + 1
        }
        return $Context.Runtime['CollectorCache'][$Key]
    }

    if (-not $CacheNameOnly) {
        $Context.Runtime['CollectorCacheStats']['Misses'] = [int]$Context.Runtime['CollectorCacheStats']['Misses'] + 1
    }
    return $null
}
