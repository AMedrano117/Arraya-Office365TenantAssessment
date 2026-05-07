[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$RedirectUri,

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9.-]*$|^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$Tenant = 'organizations',

    [Parameter(Mandatory = $false)]
    [AllowEmptyString()]
    [string]$State = '',

    [Parameter(Mandatory = $false)]
    [switch]$IncludeExchangeScope
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-M365ConsentRedirectUri {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri
    )

    $redirectUriObject = $null
    if (-not [Uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$redirectUriObject)) {
        throw "RedirectUri must be an absolute URI. Value: $Uri"
    }

    if ($redirectUriObject.Scheme -notin @('http', 'https')) {
        throw "RedirectUri must use http or https. Value: $Uri"
    }

    if ($redirectUriObject.Scheme -ne 'https' -and -not $redirectUriObject.IsLoopback) {
        throw "RedirectUri must use HTTPS unless it is a localhost or loopback URI. Value: $Uri"
    }

    return $true
}

$null = Test-M365ConsentRedirectUri -Uri $RedirectUri

$encodedRedirectUri = [Uri]::EscapeDataString($RedirectUri)
$encodedState = [Uri]::EscapeDataString($State)
$baseUrl = 'https://login.microsoftonline.com/{0}/adminconsent?client_id={1}' -f $Tenant, $ClientId

if ($IncludeExchangeScope) {
    $baseUrl = '{0}&scope=https%3A%2F%2Foutlook.office365.com%2F.default' -f $baseUrl
}

return '{0}&redirect_uri={1}&state={2}' -f $baseUrl, $encodedRedirectUri, $encodedState
