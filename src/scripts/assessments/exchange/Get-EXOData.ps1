<#
.SYNOPSIS
Collect Exchange Online data via the ExchangeOnlineManagement PS module.
Outputs a JSON file at OutputPath with EXO-only datasets that are not
available via the Microsoft Graph API.

.NOTES
Required: ExchangeOnlineManagement module v3+
For Certificate/app-only auth: Exchange.ManageAsApp app permission required.
For Interactive auth: user must have Exchange Administrator or Global Reader role.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$Organization = '',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Certificate', 'Interactive')]
    [string]$AuthMode = 'Interactive',

    [Parameter(Mandatory = $false)]
    [string]$AppId = '',

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint = '',

    [Parameter(Mandatory = $false)]
    [string]$TenantId = ''
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------
# Module check
# -----------------------------------------------------------------------
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable -ErrorAction SilentlyContinue)) {
    Write-Error "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force"
    exit 1
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

# -----------------------------------------------------------------------
# Connect
# -----------------------------------------------------------------------
try {
    if ($AuthMode -eq 'Certificate') {
        if (-not $AppId -or -not $CertificateThumbprint -or -not $Organization) {
            Write-Error "Certificate auth requires AppId, CertificateThumbprint, and Organization."
            exit 1
        }
        Connect-ExchangeOnline `
            -AppId $AppId `
            -CertificateThumbprint $CertificateThumbprint `
            -Organization $Organization `
            -ShowBanner:$false `
            -ShowProgress:$false `
            -ErrorAction Stop
    } else {
        $connectParams = @{
            ShowBanner   = $false
            ShowProgress = $false
            ErrorAction  = 'Stop'
        }
        if ($Organization) { $connectParams['Organization'] = $Organization }
        Connect-ExchangeOnline @connectParams
    }
} catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit 1
}

$ErrorActionPreference = 'Continue'

$result = @{}

# -----------------------------------------------------------------------
# AllMailboxes
# -----------------------------------------------------------------------
try {
    $mbxProps = @(
        'DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress',
        'RecipientTypeDetails', 'ExchangeObjectId',
        'ForwardingAddress', 'ForwardingSmtpAddress',
        'LitigationHoldEnabled', 'ArchiveStatus',
        'HiddenFromAddressListsEnabled', 'AccountDisabled'
    )
    $allMbx  = @(Get-EXOMailbox -ResultSize Unlimited -Properties $mbxProps -ErrorAction Stop)
    $mbxDict = @{}
    foreach ($m in $allMbx) {
        $key = if ($m.ExchangeObjectId) { [string]$m.ExchangeObjectId } else { $m.UserPrincipalName }
        $mbxDict[$key] = [PSCustomObject]@{
            DisplayName                   = $m.DisplayName
            UserPrincipalName             = $m.UserPrincipalName
            PrimarySmtpAddress            = $m.PrimarySmtpAddress
            RecipientTypeDetails          = [string]$m.RecipientTypeDetails
            ExchangeObjectId              = [string]$m.ExchangeObjectId
            ForwardingAddress             = $m.ForwardingAddress
            ForwardingSmtpAddress         = $m.ForwardingSmtpAddress
            LitigationHoldEnabled         = [bool]$m.LitigationHoldEnabled
            ArchiveStatus                 = [string]$m.ArchiveStatus
            HiddenFromAddressListsEnabled = [bool]$m.HiddenFromAddressListsEnabled
            AccountDisabled               = [bool]$m.AccountDisabled
        }
    }
    $result['AllMailboxes'] = $mbxDict
    Write-Verbose "AllMailboxes: $($mbxDict.Count) mailbox(es)"
} catch {
    Write-Warning "AllMailboxes collection failed: $_"
    $result['AllMailboxes'] = @{}
}

# -----------------------------------------------------------------------
# MailFlowRules (transport rules)
# -----------------------------------------------------------------------
try {
    $result['MailFlowRules'] = @(
        Get-TransportRule -ErrorAction Stop |
            Select-Object Name, State, Priority, Mode, Description,
                          @{N = 'ActionCount'; E = { $_.Actions.Count } }
    )
    Write-Verbose "MailFlowRules: $($result['MailFlowRules'].Count) rule(s)"
} catch {
    Write-Warning "MailFlowRules collection failed: $_"
    $result['MailFlowRules'] = @()
}

# -----------------------------------------------------------------------
# MailFlowConnectors
# -----------------------------------------------------------------------
try {
    $inbound = @(
        Get-InboundConnector -ErrorAction Stop |
            Select-Object Name, ConnectorType, Enabled, SenderDomains, ConnectorSource,
                          @{N = 'ConnectorDirection'; E = { 'Inbound' } }
    )
    $outbound = @(
        Get-OutboundConnector -ErrorAction Stop |
            Select-Object Name, ConnectorType, Enabled, SmartHosts, RecipientDomains, ConnectorSource,
                          @{N = 'ConnectorDirection'; E = { 'Outbound' } }
    )
    $result['MailFlowConnectors'] = @($inbound + $outbound)
    Write-Verbose "MailFlowConnectors: $($result['MailFlowConnectors'].Count) connector(s)"
} catch {
    Write-Warning "MailFlowConnectors collection failed: $_"
    $result['MailFlowConnectors'] = @()
}

# -----------------------------------------------------------------------
# SpamFilteringConfig (inbound policies)
# -----------------------------------------------------------------------
try {
    $result['SpamFilteringConfig'] = @(
        Get-HostedContentFilterPolicy -ErrorAction Stop |
            Select-Object Name, IsDefault, SpamAction, HighConfidenceSpamAction,
                          PhishSpamAction, HighConfidencePhishAction,
                          BulkSpamAction, BulkThreshold, ZapEnabled
    )
    Write-Verbose "SpamFilteringConfig: $($result['SpamFilteringConfig'].Count) policy/policies"
} catch {
    Write-Warning "SpamFilteringConfig collection failed: $_"
    $result['SpamFilteringConfig'] = @()
}

# -----------------------------------------------------------------------
# RemoteDomains
# -----------------------------------------------------------------------
try {
    $result['RemoteDomains'] = @(
        Get-RemoteDomain -ErrorAction Stop |
            Select-Object DomainName, AutoForwardEnabled, AutoReplyEnabled,
                          AllowedOOFType, IsInternal, MeetingForwardNotificationEnabled
    )
    Write-Verbose "RemoteDomains: $($result['RemoteDomains'].Count) domain(s)"
} catch {
    Write-Warning "RemoteDomains collection failed: $_"
    $result['RemoteDomains'] = @()
}

# -----------------------------------------------------------------------
# PublicFolderDetails
# -----------------------------------------------------------------------
try {
    $pf = @(
        Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction Stop |
            Select-Object Identity, Name, MailEnabled, HasSubfolders, ContentMailboxName
    )
    $result['PublicFolderDetails'] = $pf
    Write-Verbose "PublicFolderDetails: $($pf.Count) folder(s)"
} catch {
    Write-Warning "PublicFolderDetails collection failed (public folders may not be enabled): $_"
    $result['PublicFolderDetails'] = @()
}

# -----------------------------------------------------------------------
# InactiveMailboxDetails
# -----------------------------------------------------------------------
try {
    $inactive = @(
        Get-EXOMailbox -InactiveMailboxOnly -ResultSize Unlimited -ErrorAction Stop |
            Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress,
                          RecipientTypeDetails, WhenSoftDeleted
    )
    $result['InactiveMailboxDetails'] = $inactive
    Write-Verbose "InactiveMailboxDetails: $($inactive.Count) mailbox(es)"
} catch {
    Write-Warning "InactiveMailboxDetails collection failed: $_"
    $result['InactiveMailboxDetails'] = @()
}

# -----------------------------------------------------------------------
# LitigationHoldMailboxes (derived from AllMailboxes)
# -----------------------------------------------------------------------
$litigList = @()
if ($result['AllMailboxes'] -and $result['AllMailboxes'].Count -gt 0) {
    $litigList = @($result['AllMailboxes'].Values | Where-Object { $_.LitigationHoldEnabled -eq $true })
}
$result['LitigationHoldMailboxes'] = $litigList
Write-Verbose "LitigationHoldMailboxes: $($litigList.Count)"

# -----------------------------------------------------------------------
# ForwardingPolicySummary (derived)
# -----------------------------------------------------------------------
try {
    $rdAllowing = @($result['RemoteDomains'] | Where-Object { $_.AutoForwardEnabled -eq $true })

    $outboundPolicies = @(Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop)
    $onPolicies       = @($outboundPolicies | Where-Object { $_.AutoForwardingMode -eq 'On' })
    $modeList         = @($outboundPolicies | Select-Object -ExpandProperty AutoForwardingMode -Unique)
    $modesStr         = $modeList -join ', '

    $fwdCount = 0
    if ($result['AllMailboxes'] -and $result['AllMailboxes'].Count -gt 0) {
        $fwdCount = @($result['AllMailboxes'].Values |
            Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress }).Count
    }

    $result['ForwardingPolicySummary'] = [PSCustomObject]@{
        RemoteDomainsAllowingAutoForwarding      = $rdAllowing.Count
        PoliciesExplicitlyAllowingAutoForwarding = $onPolicies.Count
        PolicyAutoForwardingModes                = $modesStr
        MailboxesWithForwardingConfigured         = $fwdCount
    }
    Write-Verbose "ForwardingPolicySummary: remote=$($rdAllowing.Count) policies=$($onPolicies.Count) mailboxes=$fwdCount"
} catch {
    Write-Warning "ForwardingPolicySummary derivation failed: $_"
    $result['ForwardingPolicySummary'] = [PSCustomObject]@{
        RemoteDomainsAllowingAutoForwarding      = 0
        PoliciesExplicitlyAllowingAutoForwarding = 0
        PolicyAutoForwardingModes                = ''
        MailboxesWithForwardingConfigured         = 0
    }
}

# -----------------------------------------------------------------------
# Write JSON output and disconnect
# -----------------------------------------------------------------------
try {
    $json = ConvertTo-Json -InputObject $result -Depth 10
    [System.IO.File]::WriteAllText($OutputPath, $json, [System.Text.Encoding]::UTF8)
    Write-Output "EXO data written to: $OutputPath"
} catch {
    Write-Error "Failed to write output file: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
} catch {}

exit 0
