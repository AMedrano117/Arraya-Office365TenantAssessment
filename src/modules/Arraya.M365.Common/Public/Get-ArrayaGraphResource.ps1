function Get-ArrayaGraphResource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,
        [Parameter(Mandatory = $false)]
        [int]$PageSize = 999,
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Fetching data from Microsoft Graph',
        [Parameter(Mandatory = $false)]
        [switch]$PreferRest,
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 5
    )

    $resolvedHeaders = @{}
    if ($Headers) {
        $resolvedHeaders = $Headers.Clone()
    }
    elseif ($global:GraphHeaders) {
        $resolvedHeaders = $global:GraphHeaders.Clone()
    }
    elseif ($global:GraphToken) {
        $resolvedHeaders = @{
            'Content-Type'     = 'application/json'
            'Authorization'    = "Bearer $global:GraphToken"
            'ConsistencyLevel' = 'eventual'
        }
    }

    $useSdk = $false
    if (-not $PreferRest -and (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            $useSdk = $true
        }
    }

    if (-not $useSdk -and $resolvedHeaders.Count -eq 0) {
        throw "No Graph authentication context is available for '$Activity'. Connect with Microsoft Graph SDK or provide REST headers."
    }

    $queryUri = $Uri
    if ($PageSize -gt 0 -and $queryUri -notmatch '\$top=' -and $queryUri -notmatch '/\$count(\?|$)') {
        $separator = if ($queryUri.Contains('?')) { '&' } else { '?' }
        $queryUri = "$queryUri${separator}`$top=$PageSize"
    }

    $results = New-Object System.Collections.Generic.List[object]
    $scalarResponse = $null
    $currentUri = $queryUri

    while (-not [string]::IsNullOrWhiteSpace($currentUri)) {
        $attempt = 0
        $response = $null
        $success = $false

        while (-not $success -and $attempt -lt $MaxRetries) {
            try {
                if ($useSdk) {
                    $response = Invoke-MgGraphRequest -Uri $currentUri -Method GET -OutputType PSObject -ErrorAction Stop
                }
                else {
                    $response = Invoke-RestMethod -Uri $currentUri -Headers $resolvedHeaders -Method GET -ContentType 'application/json' -ErrorAction Stop
                }
                $success = $true
            }
            catch {
                if ($useSdk -and $resolvedHeaders.Count -gt 0) {
                    $useSdk = $false
                    $attempt++
                    continue
                }

                $attempt++
                if ($attempt -ge $MaxRetries) {
                    throw
                }

                $delaySeconds = [math]::Min([math]::Pow(2, $attempt), 60)
                Start-Sleep -Seconds $delaySeconds
            }
        }

        if (-not $success) {
            break
        }

        $nextLink = $null
        if ($null -ne $response) {
            if ($response.PSObject.Properties['@odata.nextLink']) {
                $nextLink = [string]$response.'@odata.nextLink'
            }
            elseif ($response.PSObject.Properties['@odata.nextlink']) {
                $nextLink = [string]$response.'@odata.nextlink'
            }
        }

        if ($null -ne $response -and $response.PSObject.Properties['value']) {
            foreach ($item in @($response.value)) {
                $results.Add($item)
            }
            $scalarResponse = $null
        }
        elseif ($response -is [array] -or $response -is [System.Collections.IList]) {
            foreach ($item in $response) {
                $results.Add($item)
            }
            $scalarResponse = $null
        }
        else {
            $scalarResponse = $response
        }

        if ($nextLink) {
            $currentUri = $nextLink
        }
        else {
            $currentUri = $null
        }
    }

    if ($results.Count -gt 0) {
        return $results.ToArray()
    }

    return $scalarResponse
}
