function Invoke-ArrayaGraphCollectionRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $false)]
        [ValidateSet('GET', 'POST')]
        [string]$Method = 'GET',
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Graph collection request',
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,
        [Parameter(Mandatory = $false)]
        [string]$GraphMode,
        [Parameter(Mandatory = $false)]
        [switch]$NoCache,
        [Parameter(Mandatory = $false)]
        $Context,
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    if ($null -ne $Context -and $Context.PSObject.Properties['Runtime'] -and ($Context.Runtime -is [System.Collections.IDictionary])) {
        if (-not ([System.Collections.IDictionary]$Context.Runtime).Contains('GraphRequestStats')) {
            $Context.Runtime['GraphRequestStats'] = [ordered]@{
                Requests = 0
                CacheHits = 0
                CacheWrites = 0
                BatchRequests = 0
                OptionalFailures = 0
            }
        }
        $Context.Runtime['GraphRequestStats']['Requests'] = [int]$Context.Runtime['GraphRequestStats']['Requests'] + 1
    }

    $headerKey = ''
    if ($Headers) {
        $headerKey = @(
            $Headers.GetEnumerator() |
                Where-Object { ([string]$_.Key) -ine 'Authorization' } |
                Sort-Object Name |
                ForEach-Object { '{0}={1}' -f $_.Name, $_.Value }
        ) -join ';'
    }
    $cacheKey = 'Graph:{0}:{1}:{2}:{3}' -f $Method.ToUpperInvariant(), $GraphMode, $Uri, $headerKey

    if (-not $NoCache -and $null -ne $Context) {
        $cachedValue = Get-ArrayaCollectorCacheValue -Context $Context -Key $cacheKey
        if ($null -ne $cachedValue) {
            if (
                $Context.Runtime -is [System.Collections.IDictionary] -and
                ([System.Collections.IDictionary]$Context.Runtime).Contains('GraphRequestStats')
            ) {
                $Context.Runtime['GraphRequestStats']['CacheHits'] = [int]$Context.Runtime['GraphRequestStats']['CacheHits'] + 1
            }
            return $cachedValue
        }
    }

    try {
        $value = & $ScriptBlock
        if (-not $NoCache -and $null -ne $Context) {
            Set-ArrayaCollectorCacheValue -Context $Context -Key $cacheKey -Value $value | Out-Null
            if (
                $Context.Runtime -is [System.Collections.IDictionary] -and
                ([System.Collections.IDictionary]$Context.Runtime).Contains('GraphRequestStats')
            ) {
                $Context.Runtime['GraphRequestStats']['CacheWrites'] = [int]$Context.Runtime['GraphRequestStats']['CacheWrites'] + 1
            }
        }
        return $value
    }
    catch {
        if (
            $null -ne $Context -and
            $Context.PSObject.Properties['Runtime'] -and
            $Context.Runtime -is [System.Collections.IDictionary] -and
            ([System.Collections.IDictionary]$Context.Runtime).Contains('GraphRequestStats')
        ) {
            $Context.Runtime['GraphRequestStats']['OptionalFailures'] = [int]$Context.Runtime['GraphRequestStats']['OptionalFailures'] + 1
        }
        throw
    }
}
