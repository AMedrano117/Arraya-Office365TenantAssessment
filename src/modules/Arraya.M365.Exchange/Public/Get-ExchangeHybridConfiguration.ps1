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
    $cacheKey = "Exchange:HybridConfiguration:$detailLevel"
    $cachedHybridConfiguration = Get-ArrayaCollectorCacheValue -Context $Context -Key $cacheKey
    if ($null -ne $cachedHybridConfiguration) {
        $tenantStatsHash['HybridConfiguration'] = $cachedHybridConfiguration
        Write-Log -Type INFO -Message "[Get-ExchangeHybridConfiguration] Reused cached Exchange hybrid configuration for $detailLevel details." -ExportFileLocation $exportDetails
        return $tenantStatsHash['HybridConfiguration']
    }
    $tenantStatsHash['HybridConfiguration'] = @{}

    function Get-ConnectorIdentityValue {
        param(
            [Parameter(Mandatory = $false)]
            $Connector
        )

        if ($null -eq $Connector) {
            return $null
        }

        foreach ($propertyName in @('Identity', 'Guid', 'Name', 'Id')) {
            if ($Connector.PSObject.Properties[$propertyName] -and -not [string]::IsNullOrWhiteSpace([string]$Connector.$propertyName)) {
                return [string]$Connector.$propertyName
            }
        }

        return $null
    }

    Write-ArrayaExchangeCollectorBanner -Message '[Get-ExchangeHybridConfiguration] START' -ExportFileLocation $exportDetails
    Write-Log -Type INFO -Message '[Get-ExchangeHybridConfiguration] START' -ExportFileLocation $exportDetails

    try {
        # Each probe records whether it answered. Suppressing errors here would make a
        # permission denial indistinguishable from a genuinely empty result.
        $probes = [ordered]@{}
        $probes['IntraOrganizationConnector'] = Invoke-ArrayaExchangeHybridProbe -Name 'IntraOrganizationConnector' -RequiresCommand 'Get-IntraOrganizationConnector' -ScriptBlock {
            @(Get-IntraOrganizationConnector -ErrorAction Stop | Where-Object { $_.Enabled })
        }
        $probes['OrganizationRelationship'] = Invoke-ArrayaExchangeHybridProbe -Name 'OrganizationRelationship' -RequiresCommand 'Get-OrganizationRelationship' -ScriptBlock {
            @(Get-OrganizationRelationship -ErrorAction Stop | Where-Object { $_.TargetApplicationUri -or $_.TargetAutodiscoverEpr -or $_.DomainNames })
        }
        $probes['InboundConnector'] = Invoke-ArrayaExchangeHybridProbe -Name 'InboundConnector' -RequiresCommand 'Get-InboundConnector' -ScriptBlock {
            @(Get-InboundConnector -ErrorAction Stop | Where-Object { $_.ConnectorType -eq 'OnPremises' })
        }
        $probes['OutboundConnector'] = Invoke-ArrayaExchangeHybridProbe -Name 'OutboundConnector' -RequiresCommand 'Get-OutboundConnector' -ScriptBlock {
            @(Get-OutboundConnector -ErrorAction Stop | Where-Object { $_.ConnectorType -eq 'OnPremises' })
        }
        $probes['MigrationEndpoint'] = Invoke-ArrayaExchangeHybridProbe -Name 'MigrationEndpoint' -RequiresCommand 'Get-MigrationEndpoint' -ScriptBlock {
            @(Get-MigrationEndpoint -ErrorAction Stop | Where-Object { $_.EndpointType -in @('ExchangeRemote', 'ExchangeRemoteMove') })
        }
        $probes['HybridConfiguration'] = Invoke-ArrayaExchangeHybridProbe -Name 'HybridConfiguration' -RequiresCommand 'Get-HybridConfiguration' -ScriptBlock {
            @(Get-HybridConfiguration -ErrorAction Stop)
        }
        $probes['OrganizationConfig'] = Invoke-ArrayaExchangeHybridProbe -Name 'OrganizationConfig' -RequiresCommand 'Get-OrganizationConfig' -ScriptBlock {
            @(Get-OrganizationConfig -ErrorAction Stop)
        }

        foreach ($probe in $probes.Values) {
            if (-not $probe.Succeeded) {
                Write-Log -Type WARNING -Message ("[Get-ExchangeHybridConfiguration] Probe '{0}' did not answer ({1}): {2}" -f $probe.Name, $probe.FailureKind, $probe.ErrorMessage) -ExportFileLocation $exportDetails
            }
        }

        $ioc = @($probes['IntraOrganizationConnector'].Value)
        $orgR = @($probes['OrganizationRelationship'].Value)
        $inb = @($probes['InboundConnector'].Value)
        $outb = @($probes['OutboundConnector'].Value)
        $mig = @($probes['MigrationEndpoint'].Value)
        $hybridConfig = @($probes['HybridConfiguration'].Value) | Select-Object -First 1
        $orgConfig = @($probes['OrganizationConfig'].Value) | Select-Object -First 1

        $usingCollectedMailFlowConnectors = ($tenantStatsHash.ContainsKey('MailFlowConnectors') -and $tenantStatsHash['MailFlowConnectors'])
        $mailFlowConnectors = if ($usingCollectedMailFlowConnectors) {
            @($tenantStatsHash['MailFlowConnectors'].Values)
        }
        else {
            Write-Log -Type INFO -Message '[Get-ExchangeHybridConfiguration] Reusing live connector queries because cached mail flow connectors are unavailable. Test-mode connectors may not be included in this fallback view.' -ExportFileLocation $exportDetails
            @(
                (Get-InboundConnector -ErrorAction SilentlyContinue -WarningAction SilentlyContinue),
                (Get-OutboundConnector -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) | Where-Object { $_ }
            )
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
            MailFlowOnPremConnectors  = @(
                $onPremFlowConnectors |
                    ForEach-Object { Get-ConnectorIdentityValue -Connector $_ } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
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

        # Probes whose emptiness is what lets us conclude "not hybrid". If any of them failed
        # to answer, an absence of evidence is not evidence of absence.
        $decisionProbeNames = @('InboundConnector', 'OutboundConnector', 'MigrationEndpoint', 'OrganizationConfig', 'IntraOrganizationConnector', 'OrganizationRelationship')
        $failedDecisionProbes = @(
            foreach ($probeName in $decisionProbeNames) {
                if ($probes.Contains($probeName) -and -not $probes[$probeName].Succeeded) {
                    $probes[$probeName]
                }
            }
        )
        $permissionDeniedProbes = @($failedDecisionProbes | Where-Object { [string]$_.FailureKind -eq 'PermissionDenied' })
        $evidenceIsComplete = ($failedDecisionProbes.Count -eq 0)

        $hybridStatus = if ($isHybrid) {
            # Positive evidence stands on its own: finding a migration endpoint proves hybrid
            # even if an unrelated probe failed.
            'Hybrid'
        }
        elseif ($hasSingleConnector) {
            'Possible Hybrid'
        }
        elseif ($hasFederationOnly) {
            'Federation/Relationship'
        }
        elseif ($permissionDeniedProbes.Count -gt 0) {
            'PermissionDenied'
        }
        elseif ($failedDecisionProbes.Count -gt 0) {
            'Unknown'
        }
        else {
            'None'
        }

        # Only a completed set of probes can support a definitive "false".
        $isHybridConfigured = if ($isHybrid) {
            $true
        }
        elseif ($evidenceIsComplete) {
            $false
        }
        else {
            $null
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
        elseif (-not $evidenceIsComplete) {
            'Undetermined'
        }
        else {
            'None'
        }

        $failedProbeSummary = @(
            $failedDecisionProbes | ForEach-Object { '{0} ({1})' -f $_.Name, $_.FailureKind }
        ) -join '; '

        $details = [pscustomobject]@{
            IsHybridConfigured           = $isHybridConfigured
            HybridStatus                 = $hybridStatus
            HybridType                   = $hybridType
            EvidenceComplete             = $evidenceIsComplete
            UnansweredProbes             = $failedProbeSummary
            UnansweredProbeCount         = $failedDecisionProbes.Count
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
        Set-ArrayaCollectorCacheValue -Context $Context -Key $cacheKey -Value $tenantStatsHash['HybridConfiguration'] | Out-Null
        $hybridLogValue = if ($null -eq $details.IsHybridConfigured) { 'Unknown' } else { [string]$details.IsHybridConfigured }
        $hybridLogType = if ($evidenceIsComplete) { 'INFO' } else { 'WARNING' }
        Write-Log -Type $hybridLogType -Message "[Get-ExchangeHybridConfiguration] Hybrid=$hybridLogValue Status=$($details.HybridStatus) Type=$($details.HybridType) Evidence=$($details.EvidenceCount) EvidenceComplete=$evidenceIsComplete UnansweredProbes=$failedProbeSummary" -ExportFileLocation $exportDetails
    }
    catch {
        Write-Log -Type WARNING -Message "[Get-ExchangeHybridConfiguration] Error: $($_.Exception.Message)" -ExportFileLocation $exportDetails
        # Null rather than false: the collection failed, so we do not know whether this
        # tenant is hybrid. Reporting false here would be an unearned answer.
        $tenantStatsHash['HybridConfiguration']['ExchangeHybrid'] = [pscustomobject]@{
            IsHybridConfigured   = $null
            HybridStatus         = if (Test-ArrayaExchangeProbePermissionError -ErrorRecord $_) { 'PermissionDenied' } else { 'Unknown' }
            HybridType           = 'Undetermined'
            EvidenceComplete     = $false
            UnansweredProbes     = 'All hybrid probes (collector error)'
            UnansweredProbeCount = 1
            EvidenceCount        = 0
            Message              = 'Unable to determine hybrid configuration'
            Error                = $_.Exception.Message
        }
    }

    $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
    Write-ArrayaExchangeCollectorCompletionBanner -Message "[Get-ExchangeHybridConfiguration] COMPLETED in $elapsed" -ExportFileLocation $exportDetails
    Write-Log -Type INFO -Message "[Get-ExchangeHybridConfiguration] COMPLETED in $elapsed" -ExportFileLocation $exportDetails
}
