function Get-ExchangeHybridConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('minimum', 'operator', 'combined', 'automation', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory = $false)]
        $Context
    )

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context -DetailLevel $detailLevel
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $start = Get-Date
    $tenantStatsHash['HybridConfiguration'] = @{}

    Write-Host 'Checking for Exchange Hybrid Configuration ...' -ForegroundColor Cyan -NoNewline
    Write-Log -Type INFO -Message '[Get-ExchangeHybridConfiguration] START' -ExportFileLocation $exportDetails

    try {
        $ioc = @()
        try {
            $ioc = @(Get-IntraOrganizationConnector -ErrorAction SilentlyContinue | Where-Object { $_.Enabled })
        }
        catch {
            $ioc = @()
        }

        $orgR = @(Get-OrganizationRelationship -ErrorAction SilentlyContinue | Where-Object { $_.TargetApplicationUri -or $_.TargetAutodiscoverEpr -or $_.DomainNames })
        $inb = @(Get-InboundConnector -ErrorAction SilentlyContinue | Where-Object { $_.ConnectorType -eq 'OnPremises' })
        $outb = @(Get-OutboundConnector -ErrorAction SilentlyContinue | Where-Object { $_.ConnectorType -eq 'OnPremises' })
        $mig = @(Get-MigrationEndpoint -ErrorAction SilentlyContinue | Where-Object { $_.EndpointType -in @('ExchangeRemote', 'ExchangeRemoteMove') })

        $hybridConfig = $null
        $hybridConfigurationCommand = Get-Command -Name 'Get-HybridConfiguration' -ErrorAction Ignore
        if ($hybridConfigurationCommand) {
            try { $hybridConfig = Get-HybridConfiguration -ErrorAction SilentlyContinue } catch {}
        }

        $orgConfig = $null
        try { $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue } catch {}

        $mailFlowConnectors = if ($tenantStatsHash.ContainsKey('MailFlowConnectors')) {
            @($tenantStatsHash['MailFlowConnectors'].Values)
        }
        else {
            @((Get-InboundConnector -ErrorAction SilentlyContinue), (Get-OutboundConnector -ErrorAction SilentlyContinue) | Where-Object { $_ })
        }

        $onPremFlowConnectors = @($mailFlowConnectors | Where-Object {
            $_.ConnectorType -eq 'OnPremises' -or $_.RouteAllMessagesViaOnPremises -eq $true
        })

        $signals = [ordered]@{
            OrganizationRelationships = ($orgR | Select-Object -ExpandProperty Name)
            InboundConnectorsOnPrem   = ($inb | Select-Object -ExpandProperty Name)
            OutboundConnectorsOnPrem  = ($outb | Select-Object -ExpandProperty Name)
            MigrationEndpoints        = ($mig | Select-Object -ExpandProperty RemoteServer)
            HybridConfigurationObject = if ($hybridConfig) { $hybridConfig.Name } else { $null }
            HybridDomains             = if ($orgConfig -and $orgConfig.PSObject.Properties['HybridDomains']) { ($orgConfig.HybridDomains -join ', ') } else { $null }
            MailFlowOnPremConnectors  = ($onPremFlowConnectors | Select-Object -ExpandProperty ID)
        }

        $evidence = New-Object System.Collections.Generic.List[string]
        if ($orgConfig -and $orgConfig.PSObject.Properties['HybridDomains'] -and $orgConfig.HybridDomains.Count -gt 0) {
            $null = $evidence.Add('HybridDomains configured in Get-OrganizationConfig')
        }
        if ($ioc.Count -gt 0) { $null = $evidence.Add("IntraOrganizationConnector enabled ($($ioc.Count))") }
        if ($orgR.Count -gt 0) { $null = $evidence.Add("OrganizationRelationship configured ($($orgR.Count))") }
        if ($inb.Count -gt 0) { $null = $evidence.Add("Inbound OnPremises connector(s) ($($inb.Count))") }
        if ($outb.Count -gt 0) { $null = $evidence.Add("Outbound OnPremises connector(s) ($($outb.Count))") }
        if ($mig.Count -gt 0) { $null = $evidence.Add("Migration endpoint(s) ($($mig.Count))") }
        if ($onPremFlowConnectors.Count -gt 0) { $null = $evidence.Add("Mail flow On-Premises connector(s) ($($onPremFlowConnectors.Count))") }

        $hasConnectorPair = (($inb.Count -gt 0) -and ($outb.Count -gt 0))
        $hasSingleConnector = (($inb.Count -gt 0) -xor ($outb.Count -gt 0))
        $hasMigrationEndpoint = ($mig.Count -gt 0)
        $hasHybridDomains = ($orgConfig -and $orgConfig.PSObject.Properties['HybridDomains'] -and $orgConfig.HybridDomains.Count -gt 0)
        $hasFederationOnly = ($ioc.Count -gt 0 -or $orgR.Count -gt 0)
        $isHybrid = $hasMigrationEndpoint -or $hasConnectorPair -or $hasHybridDomains

        $hybridStatus = if ($isHybrid) {
            'Hybrid'
        }
        elseif ($hasSingleConnector) {
            'Possible Hybrid'
        }
        elseif ($hasFederationOnly) {
            'Federation/Relationship'
        }
        else {
            'None'
        }

        $hybridType = if ($hasMigrationEndpoint) {
            'Migration/Remote Move'
        }
        elseif ($hasConnectorPair) {
            'Mail Flow (On-Premises connectors)'
        }
        elseif ($hasHybridDomains) {
            'Hybrid Domains'
        }
        elseif ($hasSingleConnector) {
            'Mail Flow (Single On-Premises connector)'
        }
        elseif ($hasFederationOnly) {
            'Free/Busy or OAuth Federation'
        }
        else {
            'None'
        }

        $details = [pscustomobject]@{
            IsHybridConfigured           = [bool]$isHybrid
            HybridStatus                 = $hybridStatus
            HybridType                   = $hybridType
            EvidenceCount                = $evidence.Count
            Evidence                     = ($evidence -join '; ')
            IntraOrgConnectorCount       = $ioc.Count
            OrgRelationshipCount         = $orgR.Count
            InboundOnPremConnectorCount  = $inb.Count
            OutboundOnPremConnectorCount = $outb.Count
            MigrationEndpointCount       = $mig.Count
            MailFlowOnPremConnectorCount = $onPremFlowConnectors.Count
        }

        if ($detailLevel -in @('operator', 'combined', 'automation', 'all', 'geek')) {
            $details | Add-Member NoteProperty OrganizationRelationships ($signals.OrganizationRelationships -join ', ')
            $details | Add-Member NoteProperty InboundOnPremConnectors ($signals.InboundConnectorsOnPrem -join ', ')
            $details | Add-Member NoteProperty OutboundOnPremConnectors ($signals.OutboundConnectorsOnPrem -join ', ')
            $details | Add-Member NoteProperty MigrationEndpoints ($signals.MigrationEndpoints -join ', ')
            $details | Add-Member NoteProperty HybridDomains ($signals.HybridDomains)
            $details | Add-Member NoteProperty MailFlowOnPremConnectors ($signals.MailFlowOnPremConnectors -join ', ')
        }

        $tenantStatsHash['HybridConfiguration']['ExchangeHybrid'] = $details
        Write-Log -Type INFO -Message "[Get-ExchangeHybridConfiguration] Hybrid=$($details.IsHybridConfigured) Type=$($details.HybridType) Evidence=$($details.EvidenceCount)" -ExportFileLocation $exportDetails
    }
    catch {
        Write-Log -Type WARNING -Message "[Get-ExchangeHybridConfiguration] Error: $($_.Exception.Message)" -ExportFileLocation $exportDetails
        $tenantStatsHash['HybridConfiguration']['ExchangeHybrid'] = [pscustomobject]@{
            IsHybridConfigured = $false
            Message            = 'Unable to determine hybrid configuration'
            Error              = $_.Exception.Message
        }
    }

    $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
    Write-Host " Completed in $elapsed" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-ExchangeHybridConfiguration] COMPLETED in $elapsed" -ExportFileLocation $exportDetails
}
