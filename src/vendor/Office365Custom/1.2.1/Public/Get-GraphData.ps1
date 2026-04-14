<#
.SYNOPSIS
Retrieves data from Microsoft Graph API with automatic pagination, throttling, and error handling.

.DESCRIPTION
Get-GraphData is a universal function for querying Microsoft Graph API endpoints using the modern Microsoft Graph
PowerShell SDK (Invoke-MgGraphRequest). It handles pagination automatically, manages throttling (429) errors with
exponential backoff, supports multiple authentication methods, and displays progress during long-running operations.

The function can gracefully fall back to Invoke-RestMethod if the SDK is unavailable or fails.

.PARAMETER Uri
The Microsoft Graph API endpoint URI to query (e.g., "https://graph.microsoft.com/v1.0/users").

.PARAMETER PageSize
Optional page size for pagination. Adds $top parameter to the URI if not already present.

.PARAMETER Activity
Activity name displayed in the progress bar. Defaults to "Fetching Data from Microsoft Graph".

.PARAMETER Operation
Current operation description displayed in the progress bar. Defaults to "Retrieving Data".

.PARAMETER Id
Progress bar ID for nested progress scenarios. Defaults to 1.

.PARAMETER ParentId
Parent progress bar ID for nested progress scenarios.

.PARAMETER UseRestMethod
Forces the use of Invoke-RestMethod instead of Invoke-MgGraphRequest. Useful for legacy scenarios or
when the Microsoft Graph PowerShell SDK is not available.

.PARAMETER AccessToken
Optional access token for authentication. If not provided, uses global tokens ($global:GraphHeaders or $global:GraphToken).

.PARAMETER MaxRetries
Maximum number of retry attempts for throttling or transient errors. Defaults to 5.

.EXAMPLE
# Get all users with automatic pagination
$users = Get-GraphData -Uri "https://graph.microsoft.com/v1.0/users" -PageSize 999

.EXAMPLE
# Get groups with progress tracking and nested progress bars
$groups = Get-GraphData -Uri "https://graph.microsoft.com/v1.0/groups" -Activity "Fetching Groups" -Id 2 -ParentId 1

.EXAMPLE
# Use REST method fallback with custom token
$token = Get-MgAccessToken
$data = Get-GraphData -Uri "https://graph.microsoft.com/v1.0/organization" -AccessToken $token -UseRestMethod

.NOTES
- Requires Microsoft Graph PowerShell SDK or valid authentication headers
- Automatically handles @odata.nextLink pagination
- Implements exponential backoff for throttling (429) errors
- Supports both v1.0 and beta Graph API endpoints
- Returns empty array @() when no results are found
#>
function Get-GraphData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [int]$PageSize,

        [Parameter(Mandatory = $false)]
        [string]$Activity = "Fetching Data from Microsoft Graph",

        [Parameter(Mandatory = $false)]
        [string]$Operation = "Retrieving Data",

        [Parameter(Mandatory = $false)]
        [int]$Id = 1,

        [Parameter(Mandatory = $false)]
        [int]$ParentId,

        [Parameter(Mandatory = $false)]
        [switch]$UseRestMethod,

        [Parameter(Mandatory = $false)]
        [switch]$SuppressProgress,

        [Parameter(Mandatory = $false)]
        [switch]$SuppressAccessDeniedWarning,

        [Parameter(Mandatory = $false)]
        [string]$AccessToken,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 5
    )

    begin {
        # Initialize progress tracking
        $pageCount = 0
        $totalRecordsRetrieved = 0

        # Prepare authentication headers for REST fallback
        $Headers = $null
        if ($global:GraphHeaders) {
            $Headers = $global:GraphHeaders
        } elseif ($global:GraphToken) {
            $Headers = @{
                'Content-Type'     = "application/json"
                'Authorization'    = "Bearer $global:GraphToken"
                'ConsistencyLevel' = "eventual"
            }
        } elseif ($AccessToken) {
            $Headers = @{
                'Content-Type'     = "application/json"
                'Authorization'    = "Bearer $AccessToken"
                'ConsistencyLevel' = "eventual"
            }
        }

        # Build query URI with pagination
        $QueryUri = $Uri
        if ($PageSize -and $QueryUri -notmatch '(?i)(?:[?&])\$top=') {
            $separator = if ($QueryUri.Contains('?')) { '&' } else { '?' }
            $QueryUri += "${separator}`$top=$PageSize"
        }

        # Initialize results collection
        $QueryResults = [System.Collections.Generic.List[PSObject]]::new()
        
        Write-Verbose "Starting Get-GraphData for URI: $QueryUri"
    }

    process {
        $CurrentUri = $QueryUri
        $MorePages = $true
        try {
            do {
                $pageCount++
                $Results = $null
                $RetryCount = 0
                $Success = $false

                # Retry loop for handling throttling and transient errors
                while (-not $Success -and $RetryCount -lt $MaxRetries) {
                    try {
                        # Primary method: Use Microsoft Graph PowerShell SDK
                        if (-not $UseRestMethod) {
                            Write-Verbose "Page $pageCount : Invoke-MgGraphRequest => $CurrentUri"
                            $Results = Invoke-MgGraphRequest -Uri $CurrentUri -Method GET -OutputType PSObject -ProgressAction SilentlyContinue -ErrorAction Stop
                            $Success = $true
                        } else {
                            # Fallback: Use REST method
                            Write-Verbose "Page $pageCount : Invoke-RestMethod => $CurrentUri"
                            if (-not $Headers) {
                                throw "No authentication headers available for Invoke-RestMethod. Please ensure you're connected to Microsoft Graph."
                            }
                            $Results = Invoke-RestMethod -Uri $CurrentUri -Headers $Headers -Method GET -ContentType "application/json" -UseBasicParsing -ErrorAction Stop
                            $Success = $true
                        }
                    }
                    catch {
                        $statusCode = $null
                        $retryAfter = 5

                        # Extract status code if available
                        if ($_.Exception.Response) {
                            $statusCode = $_.Exception.Response.StatusCode.value__
                            # Check for Retry-After header
                            if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers['Retry-After']) {
                                $retryAfter = [int]$_.Exception.Response.Headers['Retry-After']
                            }
                        }

                        # Handle specific error codes
                        switch ($statusCode) {
                            429 {
                                # Throttling - exponential backoff
                                $waitTime = [math]::Min(($retryAfter * [math]::Pow(2, $RetryCount)), 300) # Max 5 minutes
                                Write-Warning "Throttled (429). Waiting $waitTime seconds before retry $($RetryCount + 1)/$MaxRetries..."
                                Start-Sleep -Seconds $waitTime
                                $RetryCount++
                            }
                            401 {
                                # Unauthorized - token might be expired
                                Write-Warning "Authentication failed (401). Please reconnect to Microsoft Graph."
                                throw $_
                            }
                            403 {
                                # Forbidden - insufficient permissions
                                if (-not $SuppressAccessDeniedWarning) {
                                    Write-Warning "Access denied (403). Insufficient permissions for: $CurrentUri"
                                }
                                throw $_
                            }
                            404 {
                                # Not found - return empty
                                Write-Warning "Resource not found (404): $CurrentUri"
                                return @()
                            }
                            503 {
                                # Service unavailable - retry with backoff
                                $waitTime = [math]::Min((5 * [math]::Pow(2, $RetryCount)), 60)
                                Write-Warning "Service unavailable (503). Waiting $waitTime seconds before retry $($RetryCount + 1)/$MaxRetries..."
                                Start-Sleep -Seconds $waitTime
                                $RetryCount++
                            }
                            default {
                                # For SDK failure, try REST fallback once
                                if (-not $UseRestMethod -and $Headers) {
                                    Write-Warning "Invoke-MgGraphRequest failed: $($_.Exception.Message). Falling back to Invoke-RestMethod..."
                                    $UseRestMethod = $true
                                    $RetryCount++
                                } else {
                                    Write-Error "Graph API request failed: $($_.Exception.Message)"
                                    throw $_
                                }
                            }
                        }
                    }
                }

                # Check if max retries exceeded
                if (-not $Success) {
                    Write-Error "Failed to retrieve data after $MaxRetries attempts from: $CurrentUri"
                    break
                }

                # Process results
                if ($Results.value) {
                    # Standard Graph response with value property
                    foreach ($item in $Results.value) {
                        $QueryResults.Add([PSObject]$item)
                    }
                    $totalRecordsRetrieved += $Results.value.Count
                } elseif ($Results -is [System.Collections.IEnumerable] -and -not $Results.PSObject.Properties['value']) {
                    # Collection without value property (some reports)
                    foreach ($item in $Results) {
                        $QueryResults.Add([PSObject]$item)
                    }
                    $totalRecordsRetrieved += $Results.Count
                } elseif ($Results) {
                    # Single object response
                    $QueryResults.Add([PSObject]$Results)
                    $totalRecordsRetrieved++
                }

                # Update progress
                $progressParams = @{
                    Total     = [math]::Max($totalRecordsRetrieved, 1)  # Prevent divide by zero
                    Activity  = $Activity
                    Operation = "$Operation (Page $pageCount, $totalRecordsRetrieved records)"
                    Id        = $Id
                }
                if ($PSBoundParameters.ContainsKey('ParentId')) {
                    $progressParams.ParentId = $ParentId
                }
                if (-not $SuppressProgress) {
                    Write-ProgressHelper @progressParams
                }

                # Check for next page
                $NextLink = $null
                if ($Results.'@odata.nextLink') {
                    $NextLink = $Results.'@odata.nextLink'
                } elseif ($Results.PSObject.Properties['nextLink']) {
                    $NextLink = $Results.nextLink
                }

                if ($NextLink) {
                    $CurrentUri = $NextLink
                    Write-Verbose "Next page available: $NextLink"
                } else {
                    $MorePages = $false
                    Write-Verbose "No more pages. Total records retrieved: $totalRecordsRetrieved"
                }

            } while ($MorePages)
        }
        finally {
            if (-not $SuppressProgress) {
                Write-ProgressHelper -Total 1 -Activity $Activity -Id $Id -Completed
            }
        }
    }

    end {
        # Return results
        Write-Verbose "Returning $($QueryResults.Count) total results"
        if ($QueryResults.Count -eq 0) {
            return @()
        } else {
            return $QueryResults.ToArray()
        }
    }
}
