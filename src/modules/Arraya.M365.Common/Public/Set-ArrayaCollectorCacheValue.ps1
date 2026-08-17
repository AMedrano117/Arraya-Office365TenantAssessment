function Set-ArrayaCollectorCacheValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Context,
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Context) {
        return $Value
    }
    if (-not ($Context.PSObject.Properties['Runtime']) -or -not ($Context.Runtime -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName Runtime -NotePropertyValue ([ordered]@{}) -Force
    }
    if (
        -not ([System.Collections.IDictionary]$Context.Runtime).Contains('CollectorCache') -or
        -not ($Context.Runtime['CollectorCache'] -is [System.Collections.IDictionary])
    ) {
        $Context.Runtime['CollectorCache'] = [ordered]@{}
    }
    if (
        -not ([System.Collections.IDictionary]$Context.Runtime).Contains('CollectorCacheStats') -or
        -not ($Context.Runtime['CollectorCacheStats'] -is [System.Collections.IDictionary])
    ) {
        $Context.Runtime['CollectorCacheStats'] = [ordered]@{
            Hits   = 0
            Misses = 0
            Writes = 0
        }
    }

    $Context.Runtime['CollectorCache'][$Key] = $Value
    $Context.Runtime['CollectorCacheStats']['Writes'] = [int]$Context.Runtime['CollectorCacheStats']['Writes'] + 1
    return $Value
}
