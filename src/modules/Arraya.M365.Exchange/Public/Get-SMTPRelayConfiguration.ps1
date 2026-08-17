function Get-SMTPRelayConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Context
    )

    function Test-SmtpRelayProperty {
        param(
            [AllowNull()]
            $Object,
            [Parameter(Mandatory = $true)]
            [string]$Name
        )

        if ($null -eq $Object) { return $false }
        if ($Object -is [System.Collections.IDictionary]) {
            return ([System.Collections.IDictionary]$Object).Contains($Name)
        }
        return [bool]($Object.PSObject -and $null -ne $Object.PSObject.Properties[$Name])
    }

    function Get-SmtpRelayValue {
        param(
            [AllowNull()]
            $Object,
            [Parameter(Mandatory = $true)]
            [string[]]$Names
        )

        foreach ($name in $Names) {
            if (Test-SmtpRelayProperty -Object $Object -Name $name) {
                if ($Object -is [System.Collections.IDictionary]) {
                    return $Object[$name]
                }
                return $Object.$name
            }
        }
        return $null
    }

    function Convert-ToSmtpRelayNullableBoolean {
        param(
            [AllowNull()]
            $Value
        )

        if ($null -eq $Value) { return $null }
        if ($Value -is [bool]) { return $Value }

        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        switch -Regex ($text) {
            '^(?i:true|yes|enabled|on|1)$' { return $true }
            '^(?i:false|no|disabled|off|0)$' { return $false }
        }
        return $null
    }

    function Convert-ToSmtpRelayValueArray {
        param(
            [AllowNull()]
            $Value,
            [switch]$SplitText
        )

        if ($null -eq $Value) { return @() }
        if ($Value -is [string]) {
            if (-not $SplitText) {
                $normalizedText = $Value.Trim()
                if ([string]::IsNullOrWhiteSpace($normalizedText)) { return @() }
                return @($normalizedText)
            }
            return @(
                $Value -split '[,;]' |
                    ForEach-Object { ([string]$_).Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }
        if ($Value -is [System.Collections.IEnumerable]) {
            return @(
                $Value |
                    ForEach-Object { ([string]$_).Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }

        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return @() }
        return @($text)
    }

    function Convert-SmtpRelayContainerToRow {
        param(
            [AllowNull()]
            $Container
        )

        if ($null -eq $Container) { return @() }
        if ($Container -is [System.Collections.IDictionary]) {
            return @($Container.Values)
        }
        return @($Container)
    }

    function Write-SmtpRelayLog {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Type,
            [Parameter(Mandatory = $true)]
            [string]$Message,
            [string]$ExportFileLocation,
            [switch]$CaptureError,
            $ErrorRecordVar
        )

        if (-not (Get-Command -Name 'Write-Log' -ErrorAction SilentlyContinue)) {
            return
        }
        if ($CaptureError) {
            Write-Log -Type $Type -Message $Message -ExportFileLocation $ExportFileLocation -CaptureError -ErrorRecordVar $ErrorRecordVar
            return
        }
        Write-Log -Type $Type -Message $Message -ExportFileLocation $ExportFileLocation
    }

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $start = Get-Date
    $tenantStatsHash['SMTPRelayConfig'] = @{}
    $tenantStatsHash['SMTPRelayServiceAccounts'] = [ordered]@{}

    Write-ArrayaExchangeCollectorBanner -Message '[Get-SMTPRelayConfiguration] START: Checking SMTP Relay Configuration' -ExportFileLocation $exportDetails
    Write-SmtpRelayLog -Type INFO -Message '[Get-SMTPRelayConfiguration] START: Checking SMTP Relay Configuration' -ExportFileLocation $exportDetails

    try {
        $smtpConfig = [PSCustomObject]@{
            SMTPAuthEnabled                       = $null
            SMTPAuthUsers                         = 0
            SMTPAuthUsersIsMinimum                = $false
            SMTPAuthExplicitlyEnabledUsers        = 0
            SMTPAuthInheritedEnabledUsers         = 0
            SMTPAuthMailboxSettingsCollected      = 0
            SMTPAuthMailboxSettingsMissing        = 0
            SMTPAuthMailboxStateUnknown           = 0
            SMTPAuthEvidenceState                 = 'Unavailable'
            ConnectorBasedRelay                   = $null
            RelayConnectorEvidenceState           = 'Unavailable'
            DirectSendEnabled                     = $false
            RelayConnectors                       = @()
            SmtpClientAuthenticationDisabled      = $null
            TenantSmtpAuthSettingCollected        = $false
        }

        $tenantSmtpAuthDisabled = $null
        try {
            $orgConfig = Get-ArrayaSmtpRelayTransportConfig
            if ($orgConfig -and (Test-SmtpRelayProperty -Object $orgConfig -Name 'SmtpClientAuthenticationDisabled')) {
                $smtpConfig.TenantSmtpAuthSettingCollected = $true
                $tenantSmtpAuthDisabled = Convert-ToSmtpRelayNullableBoolean -Value (Get-SmtpRelayValue -Object $orgConfig -Names @('SmtpClientAuthenticationDisabled'))
                $smtpConfig.SmtpClientAuthenticationDisabled = $tenantSmtpAuthDisabled
            }
        }
        catch {
            Write-SmtpRelayLog -Type WARNING -Message '[Get-SMTPRelayConfiguration] Unable to check transport config' -ExportFileLocation $exportDetails
        }

        $mailboxInventoryAvailable = Test-SmtpRelayProperty -Object $tenantStatsHash -Name 'AllMailboxes'
        $mailboxRows = if ($mailboxInventoryAvailable) {
            @(Convert-SmtpRelayContainerToRow -Container $tenantStatsHash['AllMailboxes'])
        }
        else {
            @()
        }

        $smtpAuthServiceAccounts = [ordered]@{}
        $explicitlyEnabledCount = 0
        $inheritedEnabledCount = 0
        $unknownCount = 0
        $settingsCollectedCount = 0
        $settingsMissingCount = 0

        foreach ($mailbox in $mailboxRows) {
            if ($null -eq $mailbox) { continue }

            if (-not (Test-SmtpRelayProperty -Object $mailbox -Name 'SmtpClientAuthenticationDisabled')) {
                $settingsMissingCount++
                $unknownCount++
                continue
            }

            $settingsCollectedCount++
            $rawMailboxSetting = Get-SmtpRelayValue -Object $mailbox -Names @('SmtpClientAuthenticationDisabled')
            $mailboxSetting = Convert-ToSmtpRelayNullableBoolean -Value $rawMailboxSetting
            $isInheritedSetting = $null -eq $rawMailboxSetting -or [string]::IsNullOrWhiteSpace([string]$rawMailboxSetting)
            $effectiveDisabled = $mailboxSetting
            $settingSource = 'Explicit mailbox override'

            if ($isInheritedSetting) {
                $settingSource = 'Inherited tenant setting'
                if ($smtpConfig.TenantSmtpAuthSettingCollected -and $null -ne $tenantSmtpAuthDisabled) {
                    $effectiveDisabled = $tenantSmtpAuthDisabled
                }
                else {
                    $effectiveDisabled = $null
                }
            }
            elseif ($null -eq $mailboxSetting) {
                $settingSource = 'Unrecognized mailbox setting'
            }

            if ($null -eq $effectiveDisabled) {
                $unknownCount++
                continue
            }
            if ($effectiveDisabled -eq $true) {
                continue
            }

            if ($isInheritedSetting) {
                $inheritedEnabledCount++
            }
            else {
                $explicitlyEnabledCount++
            }

            $displayName = [string](Get-SmtpRelayValue -Object $mailbox -Names @('DisplayName'))
            $userPrincipalName = [string](Get-SmtpRelayValue -Object $mailbox -Names @('UserPrincipalName'))
            $primarySmtpAddress = [string](Get-SmtpRelayValue -Object $mailbox -Names @('PrimarySmtpAddress'))
            $lookupKey = @(
                $userPrincipalName,
                $primarySmtpAddress,
                [string](Get-SmtpRelayValue -Object $mailbox -Names @('ExternalDirectoryObjectId', 'Identity')),
                $displayName
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
            if ([string]::IsNullOrWhiteSpace($lookupKey)) {
                $lookupKey = "smtp-auth-candidate:$([guid]::NewGuid().Guid)"
            }
            $baseLookupKey = $lookupKey
            $duplicateIndex = 1
            while ($smtpAuthServiceAccounts.Contains($lookupKey)) {
                $duplicateIndex++
                $lookupKey = '{0}#{1}' -f $baseLookupKey, $duplicateIndex
            }

            $smtpAuthServiceAccounts[$lookupKey] = [pscustomobject]@{
                DisplayName                       = if ([string]::IsNullOrWhiteSpace($displayName)) { 'Not surfaced in current source' } else { $displayName }
                UserPrincipalName                 = if ([string]::IsNullOrWhiteSpace($userPrincipalName)) { $null } else { $userPrincipalName }
                PrimarySmtpAddress                = if ([string]::IsNullOrWhiteSpace($primarySmtpAddress)) { $null } else { $primarySmtpAddress }
                RecipientTypeDetails              = Get-SmtpRelayValue -Object $mailbox -Names @('RecipientTypeDetails')
                IsDirSynced                       = Get-SmtpRelayValue -Object $mailbox -Names @('IsDirSynced')
                AssignedLicensesFriendly          = Get-SmtpRelayValue -Object $mailbox -Names @('AssignedLicensesFriendly')
                SmtpClientAuthenticationDisabled  = $mailboxSetting
                EffectiveSmtpAuthEnabled          = $true
                SmtpAuthSettingSource             = $settingSource
                Notes                             = 'SMTP AUTH is effectively enabled for this mailbox. Validate whether it is used by an application, device, or service before cutover.'
            }
        }

        $enabledMailboxCount = $explicitlyEnabledCount + $inheritedEnabledCount
        $smtpConfig.SMTPAuthUsers = $enabledMailboxCount
        $smtpConfig.SMTPAuthExplicitlyEnabledUsers = $explicitlyEnabledCount
        $smtpConfig.SMTPAuthInheritedEnabledUsers = $inheritedEnabledCount
        $smtpConfig.SMTPAuthMailboxSettingsCollected = $settingsCollectedCount
        $smtpConfig.SMTPAuthMailboxSettingsMissing = $settingsMissingCount
        $smtpConfig.SMTPAuthMailboxStateUnknown = $unknownCount
        $smtpConfig.SMTPAuthUsersIsMinimum = $unknownCount -gt 0
        $smtpConfig.SMTPAuthEnabled = if ($enabledMailboxCount -gt 0) {
            $true
        }
        elseif ($unknownCount -gt 0 -or -not $mailboxInventoryAvailable) {
            $null
        }
        else {
            $false
        }
        $smtpConfig.SMTPAuthEvidenceState = if (-not $mailboxInventoryAvailable) {
            'Unavailable'
        }
        elseif ($unknownCount -gt 0) {
            'Partial'
        }
        else {
            'Complete'
        }
        $tenantStatsHash['SMTPRelayServiceAccounts'] = $smtpAuthServiceAccounts

        if ($enabledMailboxCount -gt 0) {
            Write-SmtpRelayLog -Type INFO -Message "[Get-SMTPRelayConfiguration] Found $enabledMailboxCount mailbox(es) with effective SMTP AUTH enabled ($explicitlyEnabledCount explicit; $inheritedEnabledCount inherited)." -ExportFileLocation $exportDetails
        }
        if ($unknownCount -gt 0) {
            Write-SmtpRelayLog -Type WARNING -Message "[Get-SMTPRelayConfiguration] Effective SMTP AUTH state is unknown for $unknownCount mailbox(es); enabled-user count is a minimum." -ExportFileLocation $exportDetails
        }

        $connectorInventoryAvailable = Test-SmtpRelayProperty -Object $tenantStatsHash -Name 'MailFlowConnectors'
        $connectorRows = if ($connectorInventoryAvailable) {
            @(Convert-SmtpRelayContainerToRow -Container $tenantStatsHash['MailFlowConnectors'])
        }
        else {
            @()
        }

        foreach ($connector in $connectorRows) {
            if ($null -eq $connector) { continue }

            $direction = [string](Get-SmtpRelayValue -Object $connector -Names @('ConnectorDirection', 'Direction'))
            $connectorType = [string](Get-SmtpRelayValue -Object $connector -Names @('ConnectorType'))
            $enabled = Convert-ToSmtpRelayNullableBoolean -Value (Get-SmtpRelayValue -Object $connector -Names @('Enabled'))
            $testMode = Convert-ToSmtpRelayNullableBoolean -Value (Get-SmtpRelayValue -Object $connector -Names @('TestMode'))
            if ($direction -ine 'Inbound' -or $connectorType -ine 'OnPremises' -or $enabled -ne $true -or $testMode -eq $true) {
                continue
            }

            $senderIpAddresses = @(Convert-ToSmtpRelayValueArray -Value (Get-SmtpRelayValue -Object $connector -Names @('SenderIPAddresses')) -SplitText)
            $tlsCertificateNames = @(Convert-ToSmtpRelayValueArray -Value (Get-SmtpRelayValue -Object $connector -Names @('TlsSenderCertificateName', 'TlsCertificateName')))
            if ($senderIpAddresses.Count -eq 0 -and $tlsCertificateNames.Count -eq 0) {
                continue
            }

            $authenticationEvidence = @()
            if ($senderIpAddresses.Count -gt 0) { $authenticationEvidence += 'Sender IP address' }
            if ($tlsCertificateNames.Count -gt 0) { $authenticationEvidence += 'TLS sender certificate' }
            $connectorName = [string](Get-SmtpRelayValue -Object $connector -Names @('Name', 'Id', 'Identity'))
            if ([string]::IsNullOrWhiteSpace($connectorName)) { $connectorName = 'Unnamed inbound connector' }

            $smtpConfig.RelayConnectors += [PSCustomObject]@{
                ConnectorName                   = $connectorName
                ConnectorDirection              = 'Inbound'
                ConnectorType                   = $connectorType
                AuthenticationEvidence          = $authenticationEvidence -join ', '
                SenderIPAddresses               = $senderIpAddresses -join ', '
                TlsSenderCertificateName        = $tlsCertificateNames -join ', '
                RestrictDomainsToIPAddresses    = Get-SmtpRelayValue -Object $connector -Names @('RestrictDomainsToIPAddresses')
                RestrictDomainsToCertificate    = Get-SmtpRelayValue -Object $connector -Names @('RestrictDomainsToCertificate')
                ConnectorSource                 = Get-SmtpRelayValue -Object $connector -Names @('ConnectorSource')
                CloudServicesMailEnabled        = Get-SmtpRelayValue -Object $connector -Names @('CloudServicesMailEnabled')
                TreatMessagesAsInternal         = Get-SmtpRelayValue -Object $connector -Names @('TreatMessagesAsInternal')
                Comment                         = Get-SmtpRelayValue -Object $connector -Names @('Comment', 'Description')
            }
        }

        $smtpConfig.ConnectorBasedRelay = if ($smtpConfig.RelayConnectors.Count -gt 0) {
            $true
        }
        elseif ($connectorInventoryAvailable) {
            $false
        }
        else {
            $null
        }
        $smtpConfig.RelayConnectorEvidenceState = if ($connectorInventoryAvailable) { 'Complete' } else { 'Unavailable' }
        if ($smtpConfig.RelayConnectors.Count -gt 0) {
            Write-SmtpRelayLog -Type INFO -Message "[Get-SMTPRelayConfiguration] Found $($smtpConfig.RelayConnectors.Count) enabled OnPremises inbound connector(s) with SMTP relay authentication evidence." -ExportFileLocation $exportDetails
        }

        $acceptedDomains = @(Get-ArrayaSmtpRelayAcceptedDomain)
        if ($acceptedDomains) {
            $authoritative = @($acceptedDomains | Where-Object { $_.DomainType -eq 'Authoritative' })
            if ($authoritative.Count -gt 0) {
                $smtpConfig.DirectSendEnabled = $true
                $smtpConfig | Add-Member -MemberType NoteProperty -Name 'AuthoritativeDomains' -Value ($authoritative.DomainName -join ',') -Force
            }
        }

        $tenantStatsHash['SMTPRelayConfig']['Configuration'] = $smtpConfig
    }
    catch {
        Write-SmtpRelayLog -Type ERROR -Message "[Get-SMTPRelayConfiguration] Error checking SMTP relay configuration: $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }

    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-ArrayaExchangeCollectorCompletionBanner -Message "[Get-SMTPRelayConfiguration] COMPLETED: Checking SMTP Relay Configuration in $CompletedTime" -ExportFileLocation $exportDetails
    Write-SmtpRelayLog -Type INFO -Message "[Get-SMTPRelayConfiguration] COMPLETED: Checking SMTP Relay Configuration in $CompletedTime" -ExportFileLocation $exportDetails
}
