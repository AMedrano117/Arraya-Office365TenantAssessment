function Resolve-ExoStatisticsIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $MailboxObject
    )

    if ($null -eq $MailboxObject) {
        return $null
    }

    if ($MailboxObject -is [string]) {
        return $MailboxObject
    }

    foreach ($propertyName in @('ExchangeGuid', 'ExternalDirectoryObjectId', 'Guid', 'Identity', 'PrimarySmtpAddress', 'UserPrincipalName', 'WindowsEmailAddress', 'Alias', 'DistinguishedName')) {
        $property = $MailboxObject.PSObject.Properties[$propertyName]
        if (-not $property) {
            continue
        }

        $value = $property.Value
        if ($null -eq $value) {
            continue
        }

        $stringValue = [string]$value
        if (-not [string]::IsNullOrWhiteSpace($stringValue)) {
            return $stringValue
        }
    }

    return $null
}
