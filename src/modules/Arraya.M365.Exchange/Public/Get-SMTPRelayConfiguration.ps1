function Get-SMTPRelayConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Context
    )

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $start = Get-Date
    $tenantStatsHash['SMTPRelayConfig'] = @{}

    Write-Host 'Checking SMTP Relay Configuration ...' -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message '[Get-SMTPRelayConfiguration] START: Checking SMTP Relay Configuration' -ExportFileLocation $exportDetails

    try {
        $smtpConfig = [PSCustomObject]@{
            SMTPAuthEnabled     = $false
            ConnectorBasedRelay = $false
            DirectSendEnabled   = $false
            RelayConnectors     = @()
            SMTPAuthUsers       = 0
        }

        try {
            $smtpAuthUsers = @($tenantStatsHash['AllMailboxes'].Values | Where-Object { $_.SmtpClientAuthenticationDisabled -eq $false })
            if ($smtpAuthUsers.Count -gt 0) {
                $smtpConfig.SMTPAuthEnabled = $true
                $smtpConfig.SMTPAuthUsers = $smtpAuthUsers.Count
                Write-Log -Type INFO -Message "[Get-SMTPRelayConfiguration] Found $($smtpAuthUsers.Count) mailboxes with SMTP AUTH enabled" -ExportFileLocation $exportDetails
            }
        }
        catch {
            Write-Log -Type WARNING -Message '[Get-SMTPRelayConfiguration] Unable to check SMTP AUTH on mailboxes' -ExportFileLocation $exportDetails
        }

        $outboundConnectors = @($tenantStatsHash['MailFlowConnectors'].Values | Where-Object { $_.ConnectorDirection -eq 'Outbound' -and $_.Enabled -eq $true })
        foreach ($connector in $outboundConnectors) {
            if ($connector.SmartHosts -or $connector.CloudServicesMailEnabled) {
                $smtpConfig.ConnectorBasedRelay = $true
                $smtpConfig.RelayConnectors += [PSCustomObject]@{
                    ConnectorName            = $connector.Id
                    SmartHosts               = $connector.SmartHosts
                    UseMXRecord              = $connector.UseMXRecord
                    CloudServicesMailEnabled = $connector.CloudServicesMailEnabled
                    TlsSettings              = $connector.TlsSettings
                    ConnectorType            = $connector.ConnectorType
                }
            }
        }
        if ($smtpConfig.RelayConnectors.Count -gt 0) {
            Write-Log -Type INFO -Message "[Get-SMTPRelayConfiguration] Found $($smtpConfig.RelayConnectors.Count) relay connectors" -ExportFileLocation $exportDetails
        }

        $acceptedDomains = Get-AcceptedDomain -ErrorAction SilentlyContinue
        if ($acceptedDomains) {
            $authoritative = @($acceptedDomains | Where-Object { $_.DomainType -eq 'Authoritative' })
            if ($authoritative.Count -gt 0) {
                $smtpConfig.DirectSendEnabled = $true
                $smtpConfig | Add-Member -MemberType NoteProperty -Name 'AuthoritativeDomains' -Value ($authoritative.DomainName -join ',') -Force
            }
        }

        try {
            $orgConfig = Get-TransportConfig -ErrorAction SilentlyContinue
            if ($orgConfig) {
                $smtpConfig | Add-Member -MemberType NoteProperty -Name 'SmtpClientAuthenticationDisabled' -Value $orgConfig.SmtpClientAuthenticationDisabled -Force
            }
        }
        catch {
            Write-Log -Type WARNING -Message '[Get-SMTPRelayConfiguration] Unable to check transport config' -ExportFileLocation $exportDetails
        }

        $tenantStatsHash['SMTPRelayConfig']['Configuration'] = $smtpConfig
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-SMTPRelayConfiguration] Error checking SMTP relay configuration: $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }

    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $CompletedTime" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-SMTPRelayConfiguration] COMPLETED: Checking SMTP Relay Configuration in $CompletedTime" -ExportFileLocation $exportDetails
}
