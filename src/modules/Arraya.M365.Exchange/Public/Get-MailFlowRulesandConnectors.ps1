# Get Mail Flow Rules and Connectors
function Get-MailFlowRulesandConnectors {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Provide the level of detail')]
        [ValidateSet('minimum', 'operator', 'combined', 'automation', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory = $false)]
        $Context
    )

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context -DetailLevel $detailLevel
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $initialStart = Get-ArrayaExchangeCollectorStartTime -Context $Context
    $start = Get-Date
    $mailFlowProgressId = 34
    $cacheKey = "Exchange:MailFlowRulesConnectors:$detailLevel"
    $cachedMailFlow = Get-ArrayaCollectorCacheValue -Context $Context -Key $cacheKey
    if ($null -ne $cachedMailFlow) {
        $tenantStatsHash['MailFlowRules'] = $cachedMailFlow.MailFlowRules
        $tenantStatsHash['MailFlowConnectors'] = $cachedMailFlow.MailFlowConnectors
        if ($cachedMailFlow.PSObject.Properties['RemoteDomains']) {
            $tenantStatsHash['RemoteDomains'] = $cachedMailFlow.RemoteDomains
        }
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Reused cached mail flow rules/connectors for $detailLevel details." -ExportFileLocation $exportDetails
        return $cachedMailFlow
    }
    $tenantStatsHash['MailFlowRules'] = @{}
    $tenantStatsHash['MailFlowConnectors'] = @{}

    function Add-ConnectorToHash {
        param (
            [Parameter(Mandatory = $true)]
            $ConnectorList,
            [Parameter(Mandatory = $true)]
            [ValidateSet('Inbound', 'Outbound')]
            [string]$Direction
        )

        $progressCounter = 0
        $totalCount = ($ConnectorList | Measure-Object).Count
        foreach ($connector in $ConnectorList) {
            $progressCounter++
            Write-Log -Type DEBUG -Message ("[Add-ConnectorToHash] ({0}/{1}) Gathering '{2}' Mail Connector Details for {3}" -f $progressCounter, $totalCount, $Direction, $connector.ID) -ExportFileLocation $exportDetails

            $currentConnector = $connector | Select-Object *
            $currentConnector | Add-Member -MemberType NoteProperty -Name 'ConnectorDirection' -Value $Direction -Force
            if ($detailLevel -ne 'geek') {
                foreach ($prop in @('RecipientDomains', 'SmartHosts', 'ValidationRecipients', 'SenderDomains', 'SenderIPAddresses', 'TrustedOrganizations', 'EFSkipIPs', 'EFSkipMailGateway', 'EFUsers')) {
                    if ($connector.$prop) {
                        $currentConnector | Add-Member -MemberType NoteProperty -Name $prop -Value ($connector.$prop -join ',') -Force
                    }
                }
            }

            $tenantStatsHash['MailFlowConnectors'][$connector.Id] = $currentConnector
        }
    }

    Write-ArrayaExchangeCollectorBanner -Message '[Get-MailFlowRulesandConnectors] START: Gathering all Mail Flow Rules and Connectors' -ExportFileLocation $exportDetails
    Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] START: Gathering all Mail Flow Rules and Connectors with $detailLevel details" -ExportFileLocation $exportDetails
    Write-Progress -Id $mailFlowProgressId -Activity 'Getting all Mail Flow Rules details' -Status (((Get-Date) - $initialStart).ToString('hh\:mm\:ss'))

    try {
        Write-Log -Type INFO -Message 'Gathering all Mail Flow Rules' -ExportFileLocation $exportDetails
        $transportRuleCommand = Get-Command -Name 'Get-TransportRule' -ErrorAction Ignore
        $supportsTransportRuleTestModeConnectors = $transportRuleCommand -and $transportRuleCommand.Parameters.ContainsKey('IncludeTestModeConnectors')
        switch ($detailLevel) {
            { $_ -in 'minimum', 'operator', 'combined', 'automation', 'all' } {
                $desiredProperties = @('Name', 'State', 'Mode', 'Priority', 'Description')
                if ($supportsTransportRuleTestModeConnectors) {
                    $mailFlowRules = Get-TransportRule -IncludeTestModeConnectors -ErrorAction Continue | Select-Object $desiredProperties
                }
                else {
                    $mailFlowRules = Get-TransportRule -ErrorAction Continue | Select-Object $desiredProperties
                }
            }
            default {
                if ($supportsTransportRuleTestModeConnectors) {
                    $mailFlowRules = Get-TransportRule -IncludeTestModeConnectors -ErrorAction Continue
                }
                else {
                    $mailFlowRules = Get-TransportRule -ErrorAction Continue
                }
            }
        }

        $progresscounter = 0
        $totalCount = ($mailFlowRules | Measure-Object).Count
        foreach ($rule in $mailFlowRules) {
            $progresscounter++
            Write-Log -Type DEBUG -Message "[Get-MailFlowRulesandConnectors] Gathering Mail Flow Details for $($rule.Name): $progresscounter/$totalCount" -ExportFileLocation $exportDetails
            $tenantStatsHash['MailFlowRules'][$rule.Priority] = $rule
        }
        Write-Progress -Id $mailFlowProgressId -Activity 'Getting all Mail Flow Rules details' -Completed
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-MailFlowRulesandConnectors] An error occurred in gathering MailFlow Rules. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }

    try {
        Write-Log -Type INFO -Message '[Get-MailFlowRulesandConnectors] Gathering all Inbound Mail Connectors' -ExportFileLocation $exportDetails
        $mailFlowInboundConnectors = Get-InboundConnector -ErrorAction Stop
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Found $(($mailFlowInboundConnectors | Measure-Object).Count) Inbound Mail Connectors" -ExportFileLocation $exportDetails

        Write-Log -Type INFO -Message '[Get-MailFlowRulesandConnectors] Gathering all Outbound Mail Connectors' -ExportFileLocation $exportDetails
        $outboundConnectorCommand = Get-Command -Name 'Get-OutboundConnector' -ErrorAction Ignore
        if ($outboundConnectorCommand -and $outboundConnectorCommand.Parameters.ContainsKey('IncludeTestModeConnectors')) {
            $mailFlowOutboundConnectors = Get-OutboundConnector -IncludeTestModeConnectors $true -ErrorAction Stop
        }
        else {
            $mailFlowOutboundConnectors = Get-OutboundConnector -ErrorAction Stop
        }
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Found $(($mailFlowOutboundConnectors | Measure-Object).Count) Outbound Mail Connectors" -ExportFileLocation $exportDetails

        if ($detailLevel -in @('minimum', 'operator', 'combined', 'automation', 'all')) {
            $desiredProperties = @(
                'Id', 'ConnectorDirection', 'Comment', 'Enabled', 'TestMode', 'ConnectorType',
                'UseMXRecord', 'IsTransportRuleScoped', 'RecipientDomains',
                @{ Name = 'SmartHosts'; Expression = { if ($_.SmartHosts -is [array]) { $_.SmartHosts -join ',' } else { $_.SmartHosts } } },
                'AllAcceptedDomains', 'SenderRewritingEnabled',
                'RouteAllMessagesViaOnPremises', 'CloudServicesMailEnabled',
                'ValidationRecipients', 'Description', 'IsValidated',
                'LastValidationTimestamp',
                @{ Name = 'SenderDomains'; Expression = { if ($_.SenderDomains -is [array]) { $_.SenderDomains -join ',' } else { $_.SenderDomains } } },
                @{ Name = 'SenderIPAddresses'; Expression = { if ($_.SenderIPAddresses -is [array]) { $_.SenderIPAddresses -join ',' } else { $_.SenderIPAddresses } } },
                @{ Name = 'TrustedOrganizations'; Expression = { if ($_.TrustedOrganizations -is [array]) { $_.TrustedOrganizations -join ',' } else { $_.TrustedOrganizations } } },
                'RequireTls', 'TlsSettings', 'TlsDomain', 'TlsSenderCertificateName',
                'RestrictDomainsToCertificate', 'RestrictDomainsToIPAddresses', 'ConnectorSource',
                'TreatMessagesAsInternal', 'EFTestMode', 'EFSkipLastIP',
                @{ Name = 'EFSkipIPs'; Expression = { if ($_.EFSkipIPs -is [array]) { $_.EFSkipIPs -join ',' } else { $_.EFSkipIPs } } },
                @{ Name = 'EFSkipMailGateway'; Expression = { if ($_.EFSkipMailGateway -is [array]) { $_.EFSkipMailGateway -join ',' } else { $_.EFSkipMailGateway } } },
                @{ Name = 'EFUsers'; Expression = { if ($_.EFUsers -is [array]) { $_.EFUsers -join ',' } else { $_.EFUsers } } }
            )
            $mailFlowInboundConnectors = $mailFlowInboundConnectors | Select-Object $desiredProperties
            $mailFlowOutboundConnectors = $mailFlowOutboundConnectors | Select-Object $desiredProperties
        }

        $tenantStatsHash['MailFlowConnectors'] = @{}
        if ($mailFlowInboundConnectors) {
            Add-ConnectorToHash -ConnectorList $mailFlowInboundConnectors -Direction 'Inbound'
        }
        if ($mailFlowOutboundConnectors) {
            Add-ConnectorToHash -ConnectorList $mailFlowOutboundConnectors -Direction 'Outbound'
        }
        Set-ArrayaCollectorCacheValue -Context $Context -Key $cacheKey -Value ([pscustomobject]@{
            MailFlowRules      = $tenantStatsHash['MailFlowRules']
            MailFlowConnectors = $tenantStatsHash['MailFlowConnectors']
            RemoteDomains      = if ($tenantStatsHash.ContainsKey('RemoteDomains')) { $tenantStatsHash['RemoteDomains'] } else { $null }
        }) | Out-Null
        Write-Log -Type INFO -Message '[Get-MailFlowRulesandConnectors] Add Mail Connectors Details to Tenant Stats Hash' -ExportFileLocation $exportDetails
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-MailFlowRulesandConnectors] An error occurred in MailFlow Connectors. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-Progress -Id $mailFlowProgressId -Activity 'Getting all Mail Flow Rules details' -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-ArrayaExchangeCollectorCompletionBanner -Message "[Get-MailFlowRulesandConnectors] COMPLETED: Gathering All Mail Flow Rules and Connectors in $CompletedTime" -ExportFileLocation $exportDetails
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] COMPLETED: Gathering All Mail Flow Rules and Connectors in $CompletedTime" -ExportFileLocation $exportDetails
    }
}
