function Invoke-ArrayaGraphCollectionBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Requests,
        [Parameter(Mandatory = $false)]
        [string[]]$GraphAuthType = @('REST'),
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Graph batch request',
        [Parameter(Mandatory = $false)]
        [string]$ExportFileLocation,
        [Parameter(Mandatory = $false)]
        $Context
    )

    $result = @{}
    if (-not $Requests -or $Requests.Count -eq 0) {
        return $result
    }

    if ($Context -and $Context.PSObject.Properties['Runtime'] -and ($Context.Runtime -is [System.Collections.IDictionary])) {
        if (-not $Context.Runtime.Contains('GraphRequestStats')) {
            $Context.Runtime['GraphRequestStats'] = [ordered]@{
                Requests = 0
                CacheHits = 0
                CacheWrites = 0
                BatchRequests = 0
                OptionalFailures = 0
            }
        }
    }

    $batchEndpoint = 'https://graph.microsoft.com/v1.0/$batch'
    $chunkSize = 20
    for ($offset = 0; $offset -lt $Requests.Count; $offset += $chunkSize) {
        $chunk = @($Requests | Select-Object -Skip $offset -First $chunkSize)
        if ($chunk.Count -eq 0) {
            continue
        }

        if ($Context -and $Context.Runtime.Contains('GraphRequestStats')) {
            $Context.Runtime['GraphRequestStats']['BatchRequests'] = [int]$Context.Runtime['GraphRequestStats']['BatchRequests'] + 1
        }

        $payload = @{ requests = $chunk } | ConvertTo-Json -Depth 10 -Compress
        $batchResponse = $null
        try {
            $canUseSdkBatch = (
                ($GraphAuthType -contains 'SDK') -and
                (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue) -and
                (Get-MgContext -ErrorAction SilentlyContinue)
            )
            if ($canUseSdkBatch) {
                $batchResponse = Invoke-QuietCommand -ScriptBlock {
                    Invoke-MgGraphRequest -Method POST -Uri $batchEndpoint -Body $payload -OutputType PSObject -ProgressAction SilentlyContinue -ErrorAction Stop
                }
            }
            else {
                $headers = @{}
                if ($global:GraphHeaders) {
                    $headers = $global:GraphHeaders.Clone()
                }
                elseif ($global:GraphToken) {
                    $headers = @{
                        'Content-Type'     = 'application/json'
                        'Authorization'    = "Bearer $global:GraphToken"
                        'ConsistencyLevel' = 'eventual'
                    }
                }
                if (-not $headers.ContainsKey('Content-Type')) {
                    $headers['Content-Type'] = 'application/json'
                }
                if (-not $headers.ContainsKey('ConsistencyLevel')) {
                    $headers['ConsistencyLevel'] = 'eventual'
                }

                $batchResponse = Invoke-RestMethod -Uri $batchEndpoint -Headers $headers -Method POST -Body $payload -ContentType 'application/json' -ErrorAction Stop
            }
        }
        catch {
            if ($Context -and $Context.Runtime.Contains('GraphRequestStats')) {
                $Context.Runtime['GraphRequestStats']['OptionalFailures'] = [int]$Context.Runtime['GraphRequestStats']['OptionalFailures'] + 1
            }
            if (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
                Write-Log -Type WARNING -Message "[$Activity] Graph batch chunk starting at index $offset failed. $($_.Exception.Message)" -ExportFileLocation $ExportFileLocation
            }
            continue
        }

        foreach ($response in @($batchResponse.responses)) {
            if ($response -and $response.id) {
                $result[[string]$response.id] = $response
            }
        }
    }

    return $result
}
