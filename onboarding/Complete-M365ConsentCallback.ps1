[CmdletBinding(DefaultParameterSetName = 'Query')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Query')]
    [ValidateNotNullOrEmpty()]
    [string]$State,

    [Parameter(Mandatory = $false, ParameterSetName = 'Query')]
    [Alias('admin_consent')]
    [AllowEmptyString()]
    [string]$AdminConsent,

    [Parameter(Mandatory = $false, ParameterSetName = 'Query')]
    [AllowEmptyString()]
    [string]$Tenant,

    [Parameter(Mandatory = $false, ParameterSetName = 'Query')]
    [Alias('error')]
    [AllowEmptyString()]
    [string]$ErrorCode,

    [Parameter(Mandatory = $false, ParameterSetName = 'Query')]
    [Alias('error_description')]
    [AllowEmptyString()]
    [string]$ErrorDescription,

    [Parameter(Mandatory = $true, ParameterSetName = 'Listen')]
    [switch]$Listen,

    [Parameter(Mandatory = $true, ParameterSetName = 'Listen')]
    [ValidateNotNullOrEmpty()]
    [string]$RedirectUri,

    [Parameter(Mandatory = $false, ParameterSetName = 'Listen')]
    [ValidateRange(10, 3600)]
    [int]$TimeoutSeconds = 300,

    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [securestring]$StateSecret,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$CustomersPath = (Join-Path -Path $PSScriptRoot -ChildPath 'customers.json'),

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1440)]
    [int]$StateMaxAgeMinutes = 120
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path -Path $PSScriptRoot -ChildPath 'M365ConsentState.ps1')

function Test-M365ConsentGuid {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value
    )

    $guidValue = [Guid]::Empty
    return [Guid]::TryParse($Value, [ref]$guidValue)
}

function Get-M365ConsentCustomer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    $json = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($json)) {
        return @()
    }

    try {
        $customers = $json | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Could not parse customer consent data at '$Path'. Fix or remove the file before retrying."
    }

    if ($null -eq $customers) {
        return @()
    }

    return @($customers | Where-Object { $null -ne $_ })
}

function Copy-M365ConsentObjectProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$InputObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    if ($InputObject.PSObject.Properties[$Name]) {
        $InputObject.$Name = $Value
        return
    }

    $InputObject | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
}

function Save-M365ConsentCustomer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Record
    )

    $directory = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        $null = New-Item -Path $directory -ItemType Directory -Force
    }

    $customers = @(Get-M365ConsentCustomer -Path $Path)
    $existingCustomer = $null
    foreach ($customer in $customers) {
        $tenantMatches = $false
        $nameMatches = $false

        if ($Record.PSObject.Properties['TenantId'] -and -not [string]::IsNullOrWhiteSpace([string]$Record.TenantId)) {
            $tenantMatches = ($customer.PSObject.Properties['TenantId'] -and [string]$customer.TenantId -eq [string]$Record.TenantId)
        }

        if ($Record.PSObject.Properties['CustomerName'] -and -not [string]::IsNullOrWhiteSpace([string]$Record.CustomerName)) {
            $nameMatches = ($customer.PSObject.Properties['CustomerName'] -and [string]$customer.CustomerName -eq [string]$Record.CustomerName)
        }

        if ($tenantMatches -or $nameMatches) {
            $existingCustomer = $customer
            break
        }
    }

    if ($null -eq $existingCustomer) {
        $existingCustomer = [pscustomobject]@{}
        $customers += $existingCustomer
    }

    foreach ($property in $Record.PSObject.Properties) {
        Copy-M365ConsentObjectProperty -InputObject $existingCustomer -Name $property.Name -Value $property.Value
    }

    $customers |
        ConvertTo-Json -Depth 8 |
        Set-Content -LiteralPath $Path -Encoding utf8
}

function ConvertTo-M365ConsentResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Success,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$CustomerName,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$TenantId,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$ExchangeRoleStatus = 'NotStarted',

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ResultCustomersPath
    )

    return [pscustomobject]@{
        Success            = $Success
        Message            = $Message
        CustomerName       = $CustomerName
        TenantId           = $TenantId
        ExchangeRoleStatus = $ExchangeRoleStatus
        CustomersPath      = $ResultCustomersPath
    }
}

function Invoke-M365ConsentCallbackCompletion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$CallbackAdminConsent,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$CallbackTenant,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$CallbackState,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$CallbackErrorCode,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$CallbackErrorDescription,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [securestring]$CallbackStateSecret,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CallbackCustomersPath,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 1440)]
        [int]$CallbackStateMaxAgeMinutes
    )

    if ([string]::IsNullOrWhiteSpace($CallbackState)) {
        return ConvertTo-M365ConsentResult -Success $false -Message 'Consent callback did not include state. No customer data was stored.' -CustomerName $null -TenantId $null -ResultCustomersPath $CallbackCustomersPath
    }

    try {
        $validatedState = ConvertFrom-M365ConsentStateValue -State $CallbackState -StateSecret $CallbackStateSecret -MaxAgeMinutes $CallbackStateMaxAgeMinutes
    }
    catch {
        return ConvertTo-M365ConsentResult -Success $false -Message ("Consent state validation failed: {0} No customer data was stored." -f $_.Exception.Message) -CustomerName $null -TenantId $null -ResultCustomersPath $CallbackCustomersPath
    }

    $now = [DateTimeOffset]::UtcNow.ToString('o')

    if (-not [string]::IsNullOrWhiteSpace($CallbackErrorCode)) {
        $failureMessage = if ([string]::IsNullOrWhiteSpace($CallbackErrorDescription)) {
            $CallbackErrorCode
        }
        else {
            '{0}: {1}' -f $CallbackErrorCode, $CallbackErrorDescription
        }

        $failureRecord = [pscustomobject]@{
            CustomerName       = $validatedState.CustomerName
            TenantId           = $CallbackTenant
            ConsentStatus      = 'Failed'
            ConsentError       = $CallbackErrorCode
            ConsentErrorDetail = $CallbackErrorDescription
            ExchangeRoleStatus = 'NotStarted'
            StateNonce         = $validatedState.Nonce
            LastUpdatedUtc     = $now
        }
        Save-M365ConsentCustomer -Path $CallbackCustomersPath -Record $failureRecord

        return ConvertTo-M365ConsentResult -Success $false -Message ("Consent failed for '{0}'. {1}" -f $validatedState.CustomerName, $failureMessage) -CustomerName $validatedState.CustomerName -TenantId $CallbackTenant -ResultCustomersPath $CallbackCustomersPath
    }

    if ($CallbackAdminConsent -notmatch '^(?i:true)$') {
        $failureRecord = [pscustomobject]@{
            CustomerName       = $validatedState.CustomerName
            TenantId           = $CallbackTenant
            ConsentStatus      = 'Failed'
            ConsentError       = 'admin_consent_not_true'
            ConsentErrorDetail = 'The callback did not include admin_consent=True.'
            ExchangeRoleStatus = 'NotStarted'
            StateNonce         = $validatedState.Nonce
            LastUpdatedUtc     = $now
        }
        Save-M365ConsentCustomer -Path $CallbackCustomersPath -Record $failureRecord

        return ConvertTo-M365ConsentResult -Success $false -Message ("Consent was not granted for '{0}'. The callback did not include admin_consent=True." -f $validatedState.CustomerName) -CustomerName $validatedState.CustomerName -TenantId $CallbackTenant -ResultCustomersPath $CallbackCustomersPath
    }

    if ([string]::IsNullOrWhiteSpace($CallbackTenant)) {
        return ConvertTo-M365ConsentResult -Success $false -Message ("Consent succeeded for '{0}', but the callback did not include a tenant ID. No customer tenant was stored." -f $validatedState.CustomerName) -CustomerName $validatedState.CustomerName -TenantId $null -ResultCustomersPath $CallbackCustomersPath
    }

    if (-not (Test-M365ConsentGuid -Value $CallbackTenant)) {
        return ConvertTo-M365ConsentResult -Success $false -Message ("Consent callback returned an invalid tenant ID for '{0}'. No customer tenant was stored." -f $validatedState.CustomerName) -CustomerName $validatedState.CustomerName -TenantId $CallbackTenant -ResultCustomersPath $CallbackCustomersPath
    }

    $successRecord = [pscustomobject]@{
        CustomerName       = $validatedState.CustomerName
        TenantId           = $CallbackTenant
        ConsentStatus      = 'Granted'
        ConsentError       = $null
        ConsentErrorDetail = $null
        ExchangeRoleStatus = 'Pending'
        StateNonce         = $validatedState.Nonce
        ConsentedAtUtc     = $now
        LastUpdatedUtc     = $now
    }
    Save-M365ConsentCustomer -Path $CallbackCustomersPath -Record $successRecord

    return ConvertTo-M365ConsentResult -Success $true -Message ("Consent succeeded for '{0}'. Stored customer tenant '{1}'. Exchange role bootstrap is pending." -f $validatedState.CustomerName, $CallbackTenant) -CustomerName $validatedState.CustomerName -TenantId $CallbackTenant -ExchangeRoleStatus 'Pending' -ResultCustomersPath $CallbackCustomersPath
}

function Get-M365ConsentListenerPrefix {
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

    if ($redirectUriObject.Scheme -ne 'http' -or -not $redirectUriObject.IsLoopback) {
        throw 'The built-in callback listener only supports HTTP localhost or loopback redirect URIs. Use query mode behind a real HTTPS endpoint.'
    }

    $authority = $redirectUriObject.Host
    if (-not $redirectUriObject.IsDefaultPort) {
        $authority = '{0}:{1}' -f $redirectUriObject.Host, $redirectUriObject.Port
    }

    return '{0}://{1}/' -f $redirectUriObject.Scheme, $authority
}

function ConvertTo-M365ConsentHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Result
    )

    $title = if ($Result.Success) { 'Microsoft 365 consent complete' } else { 'Microsoft 365 consent failed' }
    $encodedTitle = [Net.WebUtility]::HtmlEncode($title)
    $encodedMessage = [Net.WebUtility]::HtmlEncode([string]$Result.Message)

    return @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>$encodedTitle</title>
  <style>
    body { font-family: "Segoe UI", Arial, sans-serif; margin: 3rem; line-height: 1.5; color: #1f2937; }
    main { max-width: 42rem; }
    h1 { font-size: 1.75rem; margin-bottom: 1rem; }
  </style>
</head>
<body>
  <main>
    <h1>$encodedTitle</h1>
    <p>$encodedMessage</p>
  </main>
</body>
</html>
"@
}

if ($Listen) {
    $listener = [Net.HttpListener]::new()
    $prefix = Get-M365ConsentListenerPrefix -Uri $RedirectUri
    $listener.Prefixes.Add($prefix)

    try {
        Write-Verbose ("Listening for one Microsoft admin consent callback at {0}" -f $prefix)
        $listener.Start()
        $contextTask = $listener.GetContextAsync()
        if (-not $contextTask.Wait([TimeSpan]::FromSeconds($TimeoutSeconds))) {
            throw "Timed out waiting $TimeoutSeconds seconds for the Microsoft admin consent callback."
        }

        $context = $contextTask.Result
        $query = $context.Request.QueryString
        $result = Invoke-M365ConsentCallbackCompletion `
            -CallbackAdminConsent $query['admin_consent'] `
            -CallbackTenant $query['tenant'] `
            -CallbackState $query['state'] `
            -CallbackErrorCode $query['error'] `
            -CallbackErrorDescription $query['error_description'] `
            -CallbackStateSecret $StateSecret `
            -CallbackCustomersPath $CustomersPath `
            -CallbackStateMaxAgeMinutes $StateMaxAgeMinutes

        $html = ConvertTo-M365ConsentHtml -Result $result
        [byte[]]$responseBytes = [Text.Encoding]::UTF8.GetBytes($html)
        $context.Response.StatusCode = if ($result.Success) { 200 } else { 400 }
        $context.Response.ContentType = 'text/html; charset=utf-8'
        $context.Response.ContentLength64 = $responseBytes.Length
        $context.Response.OutputStream.Write($responseBytes, 0, $responseBytes.Length)
        $context.Response.OutputStream.Close()

        return [string]$result.Message
    }
    catch {
        throw "Could not complete the local consent callback. $($_.Exception.Message)"
    }
    finally {
        if ($listener.IsListening) {
            $listener.Stop()
        }
        $listener.Close()
    }
}

$queryResult = Invoke-M365ConsentCallbackCompletion `
    -CallbackAdminConsent $AdminConsent `
    -CallbackTenant $Tenant `
    -CallbackState $State `
    -CallbackErrorCode $ErrorCode `
    -CallbackErrorDescription $ErrorDescription `
    -CallbackStateSecret $StateSecret `
    -CallbackCustomersPath $CustomersPath `
    -CallbackStateMaxAgeMinutes $StateMaxAgeMinutes

return [string]$queryResult.Message
