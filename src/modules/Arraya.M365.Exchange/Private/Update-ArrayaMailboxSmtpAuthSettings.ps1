function Get-ArrayaMailboxSmtpAuthIdentityKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Record
    )

    if ($null -eq $Record) {
        return @()
    }

    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($propertyName in @(
        'ExternalDirectoryObjectId',
        'UserPrincipalName',
        'PrimarySmtpAddress',
        'Identity',
        'Guid'
    )) {
        $value = Get-ArrayaObjectValue -Object $Record -Names @($propertyName)
        if ($null -eq $value) {
            continue
        }

        $normalized = ([string]$value).Trim().ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($normalized) -and -not $keys.Contains($normalized)) {
            $keys.Add($normalized) | Out-Null
        }
    }

    return @($keys.ToArray())
}

function Add-ArrayaMailboxSmtpAuthSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Mailboxes,
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $CasMailboxSettings
    )

    $settingsLookup = @{}
    foreach ($setting in @($CasMailboxSettings)) {
        if ($null -eq $setting) {
            continue
        }

        $hasSmtpAuthSetting = if ($setting -is [System.Collections.IDictionary]) {
            ([System.Collections.IDictionary]$setting).Contains('SmtpClientAuthenticationDisabled')
        }
        else {
            $setting.PSObject -and $null -ne $setting.PSObject.Properties['SmtpClientAuthenticationDisabled']
        }
        if (-not $hasSmtpAuthSetting) {
            continue
        }

        foreach ($key in @(Get-ArrayaMailboxSmtpAuthIdentityKey -Record $setting)) {
            $settingsLookup[$key] = $setting
        }
    }

    $matchedCount = 0
    foreach ($mailbox in @($Mailboxes)) {
        if ($null -eq $mailbox) {
            continue
        }

        $matchedSetting = $null
        foreach ($key in @(Get-ArrayaMailboxSmtpAuthIdentityKey -Record $mailbox)) {
            if ($settingsLookup.ContainsKey($key)) {
                $matchedSetting = $settingsLookup[$key]
                break
            }
        }
        if ($null -eq $matchedSetting) {
            continue
        }

        $settingValue = Get-ArrayaObjectValue -Object $matchedSetting -Names @('SmtpClientAuthenticationDisabled')
        $mailbox | Add-Member -MemberType NoteProperty -Name 'SmtpClientAuthenticationDisabled' -Value $settingValue -Force
        $matchedCount++
    }

    return [pscustomobject]@{
        MailboxCount       = @($Mailboxes).Count
        SettingsCount      = @($CasMailboxSettings).Count
        MatchedMailboxCount = $matchedCount
    }
}

function Get-ArrayaSmtpRelayTransportConfig {
    [CmdletBinding()]
    param()

    return Get-TransportConfig -ErrorAction SilentlyContinue
}

function Get-ArrayaSmtpRelayAcceptedDomain {
    [CmdletBinding()]
    param()

    return @(Get-AcceptedDomain -ErrorAction SilentlyContinue)
}
