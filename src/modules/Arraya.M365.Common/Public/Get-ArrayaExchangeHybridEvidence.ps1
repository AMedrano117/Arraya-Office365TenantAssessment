function Get-ArrayaExchangeHybridEvidence {
    <#
    .SYNOPSIS
        Interprets the Exchange hybrid collector output as a four-state result.

    .DESCRIPTION
          Configured    - the collector confirmed hybrid.
          Indicators    - partial signals exist (for example an organization relationship or a
                          single on-premises connector) without confirmed hybrid.
          NotConfigured - an explicit false. The collector only writes false when every
                          decision probe answered.
          Unknown       - the collector did not run, failed, or its probes were incomplete.

        A populated HybridConfiguration container is not evidence of hybrid: the collector
        always writes one. The verdict comes from its IsHybridConfigured value.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [System.Collections.IDictionary]$TenantStatsHash
    )

    $newResult = {
        param([string]$State, [string]$Reason, $Details)
        [pscustomobject][ordered]@{
            State   = $State
            Reason  = $Reason
            Details = $Details
        }
    }

    if ($null -eq $TenantStatsHash -or -not $TenantStatsHash.Contains('HybridConfiguration') -or $null -eq $TenantStatsHash['HybridConfiguration']) {
        return (& $newResult 'Unknown' 'Exchange hybrid configuration was not collected.' $null)
    }

    $container = $TenantStatsHash['HybridConfiguration']
    $details = Get-ArrayaObjectValue -Object $container -Names @('ExchangeHybrid', 'Summary')
    if ($null -eq $details) {
        # Flat shapes: a single record, or a one-element list holding it.
        if ($container -is [System.Collections.IDictionary]) {
            if ($container.Count -gt 0 -and ($container.Contains('IsHybridConfigured') -or $container.Contains('HybridConfigured') -or $container.Contains('IsConfigured'))) {
                $details = $container
            }
        }
        elseif (($container -is [System.Collections.IEnumerable]) -and -not ($container -is [string])) {
            $details = @($container | Where-Object { $null -ne $_ }) | Select-Object -First 1
        }
        else {
            $details = $container
        }
    }

    if ($null -eq $details) {
        $stepEvidence = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'HybridConfiguration' -CollectorStepName 'Exchange hybrid configuration'
        return (& $newResult 'Unknown' ("Exchange hybrid configuration produced no result. {0}" -f $stepEvidence.Reason) $null)
    }

    $isHybrid = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $details -Names @('IsHybridConfigured', 'HybridConfigured', 'IsConfigured'))
    $evidenceCount = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $details -Names @('EvidenceCount')) -AsInt64)
    $migrationEndpointCount = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $details -Names @('MigrationEndpointCount')) -AsInt64)
    $hybridStatus = [string](Get-ArrayaObjectValue -Object $details -Names @('HybridStatus'))
    $unansweredProbes = [string](Get-ArrayaObjectValue -Object $details -Names @('UnansweredProbes'))
    $evidenceCompleteValue = Get-ArrayaObjectValue -Object $details -Names @('EvidenceComplete')
    $evidenceIncomplete = (Convert-ArrayaToBoolean -Value $evidenceCompleteValue) -eq $false

    if ($isHybrid -eq $true) {
        return (& $newResult 'Configured' 'The Exchange hybrid collector confirmed hybrid configuration.' $details)
    }

    if ($evidenceCount -gt 0 -or $migrationEndpointCount -gt 0) {
        $statusText = if ([string]::IsNullOrWhiteSpace($hybridStatus)) { 'hybrid or coexistence signals' } else { $hybridStatus }
        return (& $newResult 'Indicators' ("{0} indicator(s) detected ({1}) without confirmed hybrid." -f [Math]::Max($evidenceCount, $migrationEndpointCount), $statusText) $details)
    }

    if ($isHybrid -eq $false -and -not $evidenceIncomplete) {
        return (& $newResult 'NotConfigured' 'Every hybrid probe answered and none found hybrid configuration.' $details)
    }

    $reasonParts = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($hybridStatus)) { $reasonParts.Add("status $hybridStatus") | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace($unansweredProbes)) { $reasonParts.Add("unanswered probes: $unansweredProbes") | Out-Null }
    $reasonText = if ($reasonParts.Count -gt 0) { $reasonParts.ToArray() -join '; ' } else { 'the collector did not return a verdict' }
    return (& $newResult 'Unknown' ("Exchange hybrid state could not be established ({0})." -f $reasonText) $details)
}
