[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$RedirectUri,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CustomerName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [securestring]$StateSecret,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeExchangeScope,

    [Parameter(Mandatory = $false)]
    [switch]$OpenBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path -Path $PSScriptRoot -ChildPath 'M365ConsentState.ps1')

$stateEnvelope = ConvertTo-M365ConsentStateValue -CustomerName $CustomerName -StateSecret $StateSecret
Write-Verbose ("Generated signed consent state for customer '{0}' with nonce '{1}'." -f $CustomerName, $stateEnvelope.Nonce)

$urlScriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'Get-M365AdminConsentUrl.ps1'
if (-not (Test-Path -LiteralPath $urlScriptPath)) {
    throw "Required script not found: $urlScriptPath"
}

$urlParameters = @{
    ClientId    = $ClientId
    RedirectUri = $RedirectUri
    State       = $stateEnvelope.State
}

if ($IncludeExchangeScope) {
    $urlParameters.IncludeExchangeScope = $true
}

$adminConsentUrl = & $urlScriptPath @urlParameters

if ($OpenBrowser) {
    try {
        Start-Process $adminConsentUrl
    }
    catch {
        Write-Warning ("Could not open the admin consent URL in a browser. Open it manually. Details: {0}" -f $_.Exception.Message)
    }
}

return [string]$adminConsentUrl
