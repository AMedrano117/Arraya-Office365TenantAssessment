<#
.SYNOPSIS
    Invokes REST calls with retry and optional re-authentication support.
.DESCRIPTION
    Performs REST requests with exponential retries for throttling and handles token renewal when unauthorized responses occur.
.PARAMETER Uri
    The target URI to invoke.
.PARAMETER Headers
    Optional headers for the request.
.PARAMETER Method
    HTTP method to use. Defaults to GET.
.PARAMETER MaxRetries
    Maximum number of retry attempts. Defaults to 5.
.PARAMETER RetryAfter
    Delay between retries when throttled. Defaults to 30 seconds.
.PARAMETER UseBasicParsing
    Switch to enable basic parsing on Invoke-RestMethod.
.PARAMETER ContentType
    Content type for the request. Defaults to application/json.
.PARAMETER TenantId
    Tenant identifier used when requesting a new access token.
.PARAMETER ClientId
    Client identifier for authentication.
.PARAMETER ClientSecret
    Client secret for authentication.
#>
function Invoke-WithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,

        [Parameter(Mandatory = $false)]
        [string]$Method = "GET",

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 5,

        [Parameter(Mandatory = $false)]
        [int]$RetryAfter = 30,

        [Parameter(Mandatory = $false)]
        [switch]$UseBasicParsing,

        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json",

        [Parameter(Mandatory = $false)]
        [string]$TenantId,

        [Parameter(Mandatory = $false)]
        [string]$ClientId,

        [Parameter(Mandatory = $false)]
        [string]$ClientSecret
    )

    $RetryCount = 0
    do {
        try {
            Write-Verbose "Calling $Uri (attempt: $RetryCount)..."
            $Response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -ContentType $ContentType `
                        -UseBasicParsing:$UseBasicParsing -ErrorAction Stop
            return $Response
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            if ($statusCode -eq 429) {
                Write-Warning "Throttled (429). Retrying in $RetryAfter seconds..."
                Start-Sleep -Seconds $RetryAfter
                $RetryCount++
            }
            elseif ($statusCode -eq 401) {
                Write-Warning "Unauthorized (401). Re-authenticating..."
                try {
                    if ($TenantId -and $ClientId -and $ClientSecret) {
                        $newToken = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
                    }
                    else {
                        $newToken = Get-AccessToken -TenantId $global:TenantId -ClientId $global:ClientId -ClientSecret $global:ClientSecret
                    }
                    $global:GraphToken = $newToken
                    $global:GraphHeaders = @{ Authorization = "Bearer $global:GraphToken" }
                    $Headers = $global:GraphHeaders
                }
                catch {
                    Write-Warning "Re-authentication failed. Incrementing retry count."
                    $RetryCount++
                }
            }
            elseif ($statusCode -eq 404) {
                throw "(404) Resource not found: $Uri"
            }
            else {
                throw "Error during request: $($_.Exception.Message)"
            }
        }
        if ($RetryCount -ge $MaxRetries) {
            throw "Max retry attempts reached for $Uri."
        }
    } while ($RetryCount -lt $MaxRetries)
}
