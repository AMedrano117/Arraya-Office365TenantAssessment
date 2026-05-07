[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$AppId,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{40}$')]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CustomerOrganization
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-M365ExchangeCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        throw "Exchange Online PowerShell cmdlet '$Name' is not available. Install or update the ExchangeOnlineManagement module, then retry."
    }
}

function Get-M365CertificateByThumbprint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[0-9a-fA-F]{40}$')]
        [string]$Thumbprint
    )

    $normalizedThumbprint = $Thumbprint.ToUpperInvariant()
    $certificatePaths = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')
    foreach ($certificatePath in $certificatePaths) {
        $certificate = Get-ChildItem -Path $certificatePath -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $normalizedThumbprint } |
            Select-Object -First 1

        if ($certificate) {
            return $certificate
        }
    }

    return $null
}

function Resolve-M365ExchangeAppOnlyError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Connect', 'AcceptedDomain')]
        [string]$Stage,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $message = [string]$ErrorRecord.Exception.Message
    $lowerMessage = $message.ToLowerInvariant()

    if ($lowerMessage -match 'certificate|thumbprint|private key|keyset') {
        return "Certificate authentication failed. Confirm thumbprint '$CertificateThumbprint' exists in CurrentUser\My or LocalMachine\My, includes the private key, and matches the certificate uploaded to the app registration. Details: $message"
    }

    if ($lowerMessage -match 'aadsts700016|application .*not.*found|service principal.*not.*found|not found in the directory') {
        return "Admin consent appears to be missing for appId '$AppId' in organization '$CustomerOrganization'. Have a customer admin grant tenant-wide admin consent so the Enterprise Application is created. Details: $message"
    }

    if ($lowerMessage -match 'aadsts65001|consent|invalid_grant') {
        return "Admin consent appears incomplete for appId '$AppId'. Re-run customer admin consent and include the Office 365 Exchange Online application permission Exchange.ManageAsApp. Details: $message"
    }

    if ($lowerMessage -match 'manageasapp|permission|scope|unauthorized_client') {
        return "Exchange.ManageAsApp may be missing or not admin-consented for appId '$AppId'. Confirm the app registration has Office 365 Exchange Online application permission Exchange.ManageAsApp and that the customer admin granted consent. Details: $message"
    }

    if ($lowerMessage -match 'role|rbac|not authorized|access denied|forbidden|insufficient') {
        return "The customer service principal may not have an Exchange-supported directory role. Assign 'Exchange Administrator' or another supported role to the service principal in the customer tenant, then retry. Details: $message"
    }

    if ($lowerMessage -match 'organization|tenant|domain|could not be resolved|invalid.*resource|realm') {
        return "The organization value '$CustomerOrganization' may be wrong. Use the customer's .onmicrosoft.com domain with Connect-ExchangeOnline -Organization. Details: $message"
    }

    if ($Stage -eq 'AcceptedDomain') {
        return "Connected to Exchange Online, but Get-AcceptedDomain failed. This commonly means the Exchange role was not assigned to the customer service principal. Details: $message"
    }

    return "Exchange app-only connection failed. Check admin consent, Exchange.ManageAsApp, certificate trust/private key, customer organization, and the Exchange role assignment. Details: $message"
}

Assert-M365ExchangeCommand -Name 'Connect-ExchangeOnline'

if (-not (Get-Command -Name 'Get-AcceptedDomain' -ErrorAction SilentlyContinue)) {
    Write-Verbose 'Get-AcceptedDomain is not loaded yet. It should become available after Connect-ExchangeOnline succeeds.'
}

if ($CustomerOrganization -notmatch '\.onmicrosoft\.com$') {
    Write-Warning "CustomerOrganization is usually the customer's .onmicrosoft.com domain for app-only Exchange Online connections. Current value: $CustomerOrganization"
}

$certificate = Get-M365CertificateByThumbprint -Thumbprint $CertificateThumbprint
if (-not $certificate) {
    throw "Certificate missing. Thumbprint '$CertificateThumbprint' was not found in CurrentUser\My or LocalMachine\My. Install the certificate with its private key on this machine before connecting."
}

if (-not $certificate.HasPrivateKey) {
    throw "Certificate private key missing. Thumbprint '$CertificateThumbprint' exists, but it does not have a private key available to this user or machine."
}

Write-Verbose ("Using certificate '{0}' expiring {1:u}." -f $certificate.Subject, $certificate.NotAfter)

try {
    Connect-ExchangeOnline `
        -AppId $AppId `
        -CertificateThumbprint $CertificateThumbprint `
        -Organization $CustomerOrganization `
        -ShowBanner:$false `
        -ErrorAction Stop
}
catch {
    throw (Resolve-M365ExchangeAppOnlyError -Stage Connect -ErrorRecord $_)
}

try {
    try {
        $acceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop)
    }
    catch {
        throw (Resolve-M365ExchangeAppOnlyError -Stage AcceptedDomain -ErrorRecord $_)
    }

    return $acceptedDomains |
        Select-Object -Property Name, DomainName, DomainType, Default
}
finally {
    if (Get-Command -Name 'Disconnect-ExchangeOnline' -ErrorAction SilentlyContinue) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
}
