function Get-ArrayaMailExchangerClassification {
    <#
    .SYNOPSIS
        Classifies the MX hosts of a domain as Microsoft 365, third-party, mixed, or unknown.

    .DESCRIPTION
        Third-party routing requires affirmative evidence: at least one well-formed MX host
        that is not a Microsoft 365 mail exchanger. A lookup that returned nothing, or returned
        only malformed values, is Unknown and never third-party.

          Microsoft365 - every MX host is an Exchange Online endpoint. IsMicrosoft365 = $true
          ThirdParty   - no MX host is an Exchange Online endpoint.     IsMicrosoft365 = $false
          Mixed        - Exchange Online and other hosts are published. IsMicrosoft365 = $false
          NoMxRecord   - the lookup answered and the domain has no MX.  IsMicrosoft365 = $null
          Unknown      - the lookup did not answer or was unusable.     IsMicrosoft365 = $null

        Exchange Online endpoints are matched by suffix on a label boundary: the worldwide
        mail.protection.outlook.com form, the mx.microsoft form used with inbound DNSSEC, the
        legacy mail.eo.outlook.com and mail.protection.outlook.de forms, and the US Government
        and 21Vianet forms.

    .PARAMETER LookupAnswered
        True when the DNS query returned any record set. Distinguishes "no MX record" from
        "no answer".
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$MxHost,
        [Parameter(Mandatory = $false)]
        [bool]$LookupAnswered = $false
    )

    $microsoftSuffixes = @(
        '.mail.protection.outlook.com',
        '.mx.microsoft',
        '.mail.eo.outlook.com',
        '.mail.protection.outlook.de',
        '.mail.protection.office365.us',
        '.mail.protection.partner.outlook.cn'
    )
    $hostPattern = '^(?=.{1,253}$)(?:[a-z0-9_](?:[a-z0-9_-]{0,61}[a-z0-9_])?\.)+[a-z][a-z0-9-]{0,61}[a-z0-9]$'

    $microsoftHosts = New-Object System.Collections.Generic.List[string]
    $otherHosts = New-Object System.Collections.Generic.List[string]
    $malformedCount = 0

    foreach ($candidate in @($MxHost)) {
        if ($null -eq $candidate) { continue }
        $hostName = ([string]$candidate).Trim().TrimEnd('.').ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($hostName)) { continue }

        if ($hostName -notmatch $hostPattern) {
            $malformedCount++
            continue
        }

        $isMicrosoft = $false
        foreach ($suffix in $microsoftSuffixes) {
            if ($hostName.EndsWith($suffix, [System.StringComparison]::Ordinal)) {
                $isMicrosoft = $true
                break
            }
        }

        if ($isMicrosoft) {
            if (-not $microsoftHosts.Contains($hostName)) { $microsoftHosts.Add($hostName) | Out-Null }
        }
        elseif (-not $otherHosts.Contains($hostName)) {
            $otherHosts.Add($hostName) | Out-Null
        }
    }

    $classification = 'Unknown'
    $isMicrosoft365 = $null
    $reason = $null
    if ($microsoftHosts.Count -eq 0 -and $otherHosts.Count -eq 0) {
        if ($malformedCount -gt 0) {
            $reason = 'The MX lookup returned only values that are not valid host names.'
        }
        elseif ($LookupAnswered) {
            $classification = 'NoMxRecord'
            $reason = 'The DNS lookup answered and the domain publishes no MX record.'
        }
        else {
            $reason = 'The MX lookup returned no answer; DNS may be unreachable from this workstation.'
        }
    }
    elseif ($otherHosts.Count -eq 0) {
        $classification = 'Microsoft365'
        $isMicrosoft365 = $true
        $reason = 'Every MX host is an Exchange Online endpoint.'
    }
    elseif ($microsoftHosts.Count -eq 0) {
        $classification = 'ThirdParty'
        $isMicrosoft365 = $false
        $reason = 'No MX host is an Exchange Online endpoint.'
    }
    else {
        $classification = 'Mixed'
        $isMicrosoft365 = $false
        $reason = 'Exchange Online and other MX hosts are both published.'
    }

    return [pscustomobject][ordered]@{
        Classification = $classification
        IsMicrosoft365 = $isMicrosoft365
        MicrosoftHosts = @($microsoftHosts.ToArray())
        OtherHosts     = @($otherHosts.ToArray())
        Reason         = $reason
    }
}
