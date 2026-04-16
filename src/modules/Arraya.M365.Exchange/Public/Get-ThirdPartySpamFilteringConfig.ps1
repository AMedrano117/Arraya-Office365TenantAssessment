# Function to analyze Mail Flow Connectors for 3rd Party Spam Filtering
function Get-ThirdPartySpamFilteringConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Context
    )

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $start = Get-Date
    $tenantStatsHash['SpamFilteringConfig'] = @{}

    Write-ArrayaExchangeCollectorBanner -Message '[Get-ThirdPartySpamFilteringConfig] START: Analyzing mail flow connectors' -ExportFileLocation $exportDetails
    Write-Log -Type INFO -Message '[Get-ThirdPartySpamFilteringConfig] START: Analyzing mail flow connectors' -ExportFileLocation $exportDetails

    try {
        $spamFilterConfig = [PSCustomObject]@{
            Uses3rdPartyFiltering      = $false
            InboundConnectorCount      = 0
            OutboundConnectorCount     = 0
            InboundConnectors          = @()
            OutboundConnectors         = @()
            TransportRuleCount         = 0
            TransportRulesWithTrustedIPs = 0
            TransportRuleIndicators    = @()
            PotentialSpamFilters       = @()
        }

        $inboundConnectors = @($tenantStatsHash['MailFlowConnectors'].Values | Where-Object { $_.ConnectorDirection -eq 'Inbound' })
        $spamFilterConfig.InboundConnectorCount = $inboundConnectors.Count
        foreach ($connector in $inboundConnectors) {
            $spamFilterConfig.InboundConnectors += [PSCustomObject]@{
                Name                    = $connector.Id
                Enabled                 = $connector.Enabled
                SenderDomains           = $connector.SenderDomains
                SenderIPAddresses       = $connector.SenderIPAddresses
                RequireTls              = $connector.RequireTls
                TreatMessagesAsInternal = $connector.TreatMessagesAsInternal
            }
            if ($connector.SenderIPAddresses -or $connector.TlsSenderCertificateName) {
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $spamFilterConfig.PotentialSpamFilters += "Inbound: $($connector.Id)"
            }
        }

        $outboundConnectors = @($tenantStatsHash['MailFlowConnectors'].Values | Where-Object { $_.ConnectorDirection -eq 'Outbound' })
        $spamFilterConfig.OutboundConnectorCount = $outboundConnectors.Count
        foreach ($connector in $outboundConnectors) {
            $spamFilterConfig.OutboundConnectors += [PSCustomObject]@{
                Name                    = $connector.Id
                Enabled                 = $connector.Enabled
                SmartHosts              = $connector.SmartHosts
                UseMXRecord             = $connector.UseMXRecord
                CloudServicesMailEnabled = $connector.CloudServicesMailEnabled
                TlsSettings             = $connector.TlsSettings
            }
            if ($connector.SmartHosts -and $connector.Enabled) {
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $spamFilterConfig.PotentialSpamFilters += "Outbound: $($connector.Id) -> $($connector.SmartHosts)"
            }
            if ($connector.SmartHosts -match '(?i)proofpoint|ppe-hosted|pphosted|ironport|cisco|barracuda|mimecast|forcepoint|sophos|trendmicro|spamtitan') {
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $spamFilterConfig.PotentialSpamFilters += "Outbound SmartHost match: $($connector.Id) -> $($connector.SmartHosts)"
            }
        }

        $transportRules = @($tenantStatsHash['MailFlowRules'].Values)
        $spamFilterConfig.TransportRuleCount = $transportRules.Count
        foreach ($rule in $transportRules) {
            $indicatorHits = @()
            $trustedIpValue = $null
            if ($rule.PSObject.Properties['SenderIPRanges'] -and $rule.SenderIPRanges) {
                $trustedIpValue = $rule.SenderIPRanges
                $indicatorHits += 'SenderIPRanges'
            }
            elseif ($rule.PSObject.Properties['SenderIPAddresses'] -and $rule.SenderIPAddresses) {
                $trustedIpValue = $rule.SenderIPAddresses
                $indicatorHits += 'SenderIPAddresses'
            }

            if ($rule.PSObject.Properties['SetSCL'] -and $rule.SetSCL -eq -1) { $indicatorHits += 'SetSCL=-1' }
            if ($rule.PSObject.Properties['BypassSpamFiltering'] -and $rule.BypassSpamFiltering -eq $true) { $indicatorHits += 'BypassSpamFiltering' }
            if ($rule.PSObject.Properties['HeaderContainsWords'] -and $rule.HeaderContainsWords -match '(?i)proofpoint|ironport|mimecast|barracuda|sophos|spamtitan') { $indicatorHits += 'HeaderContainsWords' }
            if ($rule.PSObject.Properties['HeaderMatchesMessageHeader'] -and $rule.HeaderMatchesMessageHeader -match '(?i)proofpoint|ironport|mimecast|barracuda|sophos|spamtitan') { $indicatorHits += 'HeaderMatchesMessageHeader' }

            if ($indicatorHits.Count -gt 0) {
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $summary = "Rule: $($rule.Name) ($($indicatorHits -join ', '))"
                if ($trustedIpValue) {
                    $summary += " | Trusted IPs: $trustedIpValue"
                    $spamFilterConfig.TransportRulesWithTrustedIPs++
                }
                $spamFilterConfig.TransportRuleIndicators += $summary
            }
        }

        try {
            $hostedFilter = Get-HostedConnectionFilterPolicy -ErrorAction SilentlyContinue
            if ($hostedFilter -and $hostedFilter.IPAllowList.Count -gt 0) {
                $spamFilterConfig | Add-Member -MemberType NoteProperty -Name 'IPAllowList' -Value ($hostedFilter.IPAllowList -join ',') -Force
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $spamFilterConfig.TransportRulesWithTrustedIPs += $hostedFilter.IPAllowList.Count
                Write-Log -Type INFO -Message "[Get-ThirdPartySpamFilteringConfig] IP Allow List configured with $($hostedFilter.IPAllowList.Count) entries" -ExportFileLocation $exportDetails
            }
        }
        catch {
            Write-Log -Type WARNING -Message '[Get-ThirdPartySpamFilteringConfig] Unable to check hosted connection filter policy' -ExportFileLocation $exportDetails
        }

        $tenantStatsHash['SpamFilteringConfig']['Configuration'] = $spamFilterConfig
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-ThirdPartySpamFilteringConfig] Error analyzing spam filtering config: $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }

    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-ArrayaExchangeCollectorCompletionBanner -Message "[Get-ThirdPartySpamFilteringConfig] COMPLETED: Analyzing mail flow in $CompletedTime" -ExportFileLocation $exportDetails
    Write-Log -Type INFO -Message "[Get-ThirdPartySpamFilteringConfig] COMPLETED: Analyzing mail flow in $CompletedTime" -ExportFileLocation $exportDetails
}
