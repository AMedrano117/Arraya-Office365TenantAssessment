<#
.SYNOPSIS
    Retrieves an access token from Microsoft Graph.
.DESCRIPTION
    This helper function requests an access token using client credentials.
.PARAMETER TenantId
    The tenant ID.
.PARAMETER ClientId
    The client ID.
.PARAMETER ClientSecret
    The client secret.
.EXAMPLE
    Get-AccessToken -TenantId "12345" -ClientId "abcde" -ClientSecret "secret"
#>

function Get-AccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $body = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    $response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -ContentType "application/x-www-form-urlencoded" -Body $body
    return $response.access_token
}
