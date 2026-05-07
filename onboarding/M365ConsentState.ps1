Set-StrictMode -Version Latest

function ConvertTo-M365ConsentBase64Url {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function ConvertFrom-M365ConsentBase64Url {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Value
    )

    $base64 = $Value.Replace('-', '+').Replace('_', '/')
    switch ($base64.Length % 4) {
        0 { break }
        2 { $base64 = '{0}==' -f $base64; break }
        3 { $base64 = '{0}=' -f $base64; break }
        default { throw 'The consent state contains invalid base64url padding.' }
    }

    try {
        return [Convert]::FromBase64String($base64)
    }
    catch {
        throw 'The consent state contains invalid base64url data.'
    }
}

function ConvertTo-M365ConsentStateSecretByteArray {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [securestring]$StateSecret
    )

    $secretPointer = [IntPtr]::Zero
    try {
        $secretPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($StateSecret)
        $plainTextSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($secretPointer)

        if ([string]::IsNullOrWhiteSpace($plainTextSecret)) {
            throw 'StateSecret cannot be empty.'
        }

        if ($plainTextSecret.Length -lt 16) {
            throw 'StateSecret must be at least 16 characters long.'
        }

        return [Text.Encoding]::UTF8.GetBytes($plainTextSecret)
    }
    finally {
        if ($secretPointer -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($secretPointer)
        }
    }
}

function Get-M365ConsentNonce {
    [CmdletBinding()]
    param()

    [byte[]]$bytes = New-Object byte[] 32
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $rng.GetBytes($bytes)
    }
    finally {
        $rng.Dispose()
    }

    return ConvertTo-M365ConsentBase64Url -Bytes $bytes
}

function Get-M365ConsentSignature {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$PayloadBase64Url,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [securestring]$StateSecret
    )

    [byte[]]$secretBytes = ConvertTo-M365ConsentStateSecretByteArray -StateSecret $StateSecret
    try {
        $hmac = [Security.Cryptography.HMACSHA256]::new($secretBytes)
        try {
            [byte[]]$payloadBytes = [Text.Encoding]::UTF8.GetBytes($PayloadBase64Url)
            [byte[]]$signatureBytes = $hmac.ComputeHash($payloadBytes)
            return ConvertTo-M365ConsentBase64Url -Bytes $signatureBytes
        }
        finally {
            $hmac.Dispose()
        }
    }
    finally {
        [Array]::Clear($secretBytes, 0, $secretBytes.Length)
    }
}

function Test-M365ConsentSignatureMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Expected,

        [Parameter(Mandatory = $true)]
        [byte[]]$Actual
    )

    if ($Expected.Length -ne $Actual.Length) {
        return $false
    }

    $difference = 0
    for ($index = 0; $index -lt $Expected.Length; $index++) {
        $difference = $difference -bor ($Expected[$index] -bxor $Actual[$index])
    }

    return ($difference -eq 0)
}

function ConvertTo-M365ConsentStateValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CustomerName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [securestring]$StateSecret
    )

    $nonce = Get-M365ConsentNonce
    $issuedAtUtc = [DateTimeOffset]::UtcNow
    $payload = [ordered]@{
        version             = 1
        customerName        = $CustomerName
        nonce               = $nonce
        issuedAtUtc         = $issuedAtUtc.ToString('o')
        issuedAtUnixSeconds = $issuedAtUtc.ToUnixTimeSeconds()
    }

    $payloadJson = $payload | ConvertTo-Json -Compress -Depth 4
    [byte[]]$payloadBytes = [Text.Encoding]::UTF8.GetBytes($payloadJson)
    $payloadBase64Url = ConvertTo-M365ConsentBase64Url -Bytes $payloadBytes
    $signatureBase64Url = Get-M365ConsentSignature -PayloadBase64Url $payloadBase64Url -StateSecret $StateSecret

    return [pscustomobject]@{
        State       = '{0}.{1}' -f $payloadBase64Url, $signatureBase64Url
        CustomerName = $CustomerName
        Nonce       = $nonce
        IssuedAtUtc = $issuedAtUtc
    }
}

function ConvertFrom-M365ConsentStateValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$State,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [securestring]$StateSecret,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 1440)]
        [int]$MaxAgeMinutes = 120
    )

    $stateParts = $State.Split('.')
    if ($stateParts.Count -ne 2) {
        throw 'The consent state is malformed. Expected a signed state value with two parts.'
    }

    $payloadBase64Url = $stateParts[0]
    $providedSignatureBase64Url = $stateParts[1]
    $expectedSignatureBase64Url = Get-M365ConsentSignature -PayloadBase64Url $payloadBase64Url -StateSecret $StateSecret

    [byte[]]$expectedSignatureBytes = ConvertFrom-M365ConsentBase64Url -Value $expectedSignatureBase64Url
    [byte[]]$providedSignatureBytes = ConvertFrom-M365ConsentBase64Url -Value $providedSignatureBase64Url
    if (-not (Test-M365ConsentSignatureMatch -Expected $expectedSignatureBytes -Actual $providedSignatureBytes)) {
        throw 'The consent state signature is invalid.'
    }

    [byte[]]$payloadBytes = ConvertFrom-M365ConsentBase64Url -Value $payloadBase64Url
    $payloadJson = [Text.Encoding]::UTF8.GetString($payloadBytes)
    try {
        $payload = $payloadJson | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw 'The consent state payload is not valid JSON.'
    }

    foreach ($requiredProperty in @('version', 'customerName', 'nonce', 'issuedAtUnixSeconds')) {
        if (-not $payload.PSObject.Properties[$requiredProperty]) {
            throw ("The consent state payload is missing '{0}'." -f $requiredProperty)
        }
    }

    if ([int]$payload.version -ne 1) {
        throw ("Unsupported consent state version '{0}'." -f $payload.version)
    }

    try {
        $issuedAtUtc = [DateTimeOffset]::FromUnixTimeSeconds([int64]$payload.issuedAtUnixSeconds)
    }
    catch {
        throw 'The consent state issuedAtUnixSeconds value is invalid.'
    }

    $nowUtc = [DateTimeOffset]::UtcNow
    if ($issuedAtUtc -gt $nowUtc.AddMinutes(5)) {
        throw 'The consent state was issued in the future. Check clock skew before retrying.'
    }

    if (($nowUtc - $issuedAtUtc).TotalMinutes -gt $MaxAgeMinutes) {
        throw ("The consent state is older than {0} minutes. Start a new consent request." -f $MaxAgeMinutes)
    }

    return [pscustomobject]@{
        CustomerName = [string]$payload.customerName
        Nonce       = [string]$payload.nonce
        IssuedAtUtc = $issuedAtUtc
    }
}
