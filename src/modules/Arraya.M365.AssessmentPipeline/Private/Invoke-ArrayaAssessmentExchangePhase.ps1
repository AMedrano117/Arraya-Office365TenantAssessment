function Invoke-ArrayaAssessmentExchangePhase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $false)]
        [string]$InputSnapshotPath
    )

    if (-not (Test-ArrayaAssessmentPipelineSessionReadiness -Context $Context -PhaseName 'Exchange')) {
        Invoke-ArrayaAssessmentConnectionPhase -Context $Context | Out-Null
    }

    $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Exchange'
    $detailLevel = ([string]$Context.ReportingMode).ToLowerInvariant()
    $profileCollectionPlan = Resolve-ArrayaAssessmentPipelineCollectionPlan -Context $Context
    $currentSnapshotPath = $InputSnapshotPath
    $currentSnapshot = if (-not [string]::IsNullOrWhiteSpace($currentSnapshotPath)) {
        Import-ArrayaTenantSnapshot -Path $currentSnapshotPath -SkipValidation
    }
    else {
        $null
    }

    $tenantStatsHash = if ($currentSnapshot) {
        Convert-ArrayaSnapshotToLegacyTenantStatsHash -Snapshot $currentSnapshot
    }
    else {
        @{}
    }

    $assessmentContext = New-ArrayaAssessmentContext `
        -ExportFileLocation $null `
        -ReportingMode ([string]$Context.ReportingMode) `
        -TenantStats $tenantStatsHash `
        -Policies ([ordered]@{}) `
        -Runtime ([ordered]@{}) `
        -Metadata ([ordered]@{ StartedAt = (Get-Date) })

    Write-Host ("Output profile: {0} (scope: {1})" -f $Context.OutputProfileLabel, ([string]$Context.ReportingMode).ToLowerInvariant()) -ForegroundColor DarkGray
    Write-Host 'Microsoft 365 Tenant Assessment' -ForegroundColor White
    Write-Host ''
    Write-Host '[3/6] Exchange' -ForegroundColor White
    Write-Host ('-' * 72) -ForegroundColor DarkGray

    $exchangeSteps = @(
        [pscustomobject]@{
            Index      = 1
            Name       = 'Exchange recipients'
            Mode       = 'Native'
            Enabled    = [bool]$profileCollectionPlan.CollectExchangeRecipients
            SkipReason = 'Not required for this profile output.'
            ScriptBlock = { Get-AllRecipientDetails -detailLevel $detailLevel -Context $assessmentContext | Out-Null }
        }
        [pscustomobject]@{
            Index      = 2
            Name       = 'Exchange mailboxes'
            Mode       = 'Native'
            Enabled    = $true
            SkipReason = $null
            ScriptBlock = { Get-AllExchangeMailboxDetails -detailLevel $detailLevel -Context $assessmentContext | Out-Null }
        }
        [pscustomobject]@{
            Index      = 3
            Name       = 'Exchange groups'
            Mode       = 'Native'
            Enabled    = [bool]$profileCollectionPlan.CollectExchangeGroups
            SkipReason = 'Not required for this profile output.'
            ScriptBlock = { Get-ExchangeGroupDetails -detailLevel $detailLevel -Context $assessmentContext | Out-Null }
        }
        [pscustomobject]@{
            Index      = 4
            Name       = 'Public folders'
            Mode       = 'Native'
            Enabled    = [bool]$profileCollectionPlan.CollectPublicFolders
            SkipReason = 'Not required for this profile output.'
            ScriptBlock = { Get-AllPublicFolderDetails -detailLevel $detailLevel -Context $assessmentContext | Out-Null }
        }
        [pscustomobject]@{
            Index      = 5
            Name       = 'Exchange hybrid configuration'
            Mode       = 'Native'
            Enabled    = $true
            SkipReason = $null
            ScriptBlock = { Get-ExchangeHybridConfiguration -detailLevel $detailLevel -Context $assessmentContext | Out-Null }
        }
        [pscustomobject]@{
            Index      = 6
            Name       = 'Mail flow rules/connectors'
            Mode       = 'Native'
            Enabled    = [bool]$profileCollectionPlan.CollectMailFlowRulesConnectors
            SkipReason = 'Skipped in best-practices-only profile to reduce runtime.'
            ScriptBlock = { Get-MailFlowRulesandConnectors -detailLevel $detailLevel -Context $assessmentContext | Out-Null }
        }
        [pscustomobject]@{
            Index      = 7
            Name       = 'Email activity insights'
            Mode       = 'Legacy'
            Enabled    = [bool]$profileCollectionPlan.CollectEmailActivityDetails
            SkipReason = 'Not required for this profile output.'
        }
        [pscustomobject]@{
            Index      = 8
            Name       = 'Third-party spam filtering configuration'
            Mode       = 'Native'
            Enabled    = [bool]$profileCollectionPlan.CollectThirdPartySpamFiltering
            SkipReason = 'Requires mail flow connector/rule collection, which is disabled for this profile.'
            ScriptBlock = { Get-ThirdPartySpamFilteringConfig -Context $assessmentContext | Out-Null }
        }
        [pscustomobject]@{
            Index      = 9
            Name       = 'SMTP relay configuration'
            Mode       = 'Native'
            Enabled    = [bool]$profileCollectionPlan.CollectSmtpRelayConfiguration
            SkipReason = 'Requires mail flow connector collection, which is disabled for this profile.'
            ScriptBlock = { Get-SMTPRelayConfiguration -Context $assessmentContext | Out-Null }
        }
        [pscustomobject]@{
            Index      = 10
            Name       = 'Exchange governance summaries'
            Mode       = 'Legacy'
            Enabled    = $true
            SkipReason = $null
        }
    )

    foreach ($step in $exchangeSteps) {
        if ($step.Mode -eq 'Native') {
            Invoke-ArrayaAssessmentPipelineProfileAwareStep `
                -Index $step.Index `
                -Total 10 `
                -Name $step.Name `
                -Enabled ([bool]$step.Enabled) `
                -SkipReason $step.SkipReason `
                -ScriptBlock $step.ScriptBlock | Out-Null

            if ($step.Enabled) {
                $currentSnapshot = Export-ArrayaAssessmentPipelineLegacyTenantStatsSnapshot -Context $Context -TenantStatsHash $tenantStatsHash -InputSnapshot $currentSnapshot -Path $checkpointPath
                $currentSnapshotPath = $checkpointPath
            }

            continue
        }

        $legacyStepResult = Invoke-ArrayaAssessmentLegacyPipelineStep `
            -Context $Context `
            -PhaseName 'Exchange' `
            -StepName $step.Name `
            -InputSnapshotPath $currentSnapshotPath `
            -CheckpointPath $checkpointPath `
            -ProgressTotalSteps 10 `
            -ProgressStartingStep ($step.Index - 1) `
            -ReuseCurrentSessions $true

        if (-not [string]::IsNullOrWhiteSpace([string]$legacyStepResult.CheckpointPath)) {
            $currentSnapshotPath = [string]$legacyStepResult.CheckpointPath
            $currentSnapshot = Import-ArrayaTenantSnapshot -Path $currentSnapshotPath -SkipValidation
            $tenantStatsHash = Convert-ArrayaSnapshotToLegacyTenantStatsHash -Snapshot $currentSnapshot
            $assessmentContext.TenantStats = $tenantStatsHash
        }
    }

    if ([string]::IsNullOrWhiteSpace($currentSnapshotPath) -or -not (Test-Path -Path $currentSnapshotPath)) {
        $currentSnapshot = Export-ArrayaAssessmentPipelineLegacyTenantStatsSnapshot -Context $Context -TenantStatsHash $tenantStatsHash -InputSnapshot $currentSnapshot -Path $checkpointPath
        $currentSnapshotPath = $checkpointPath
    }

    return [pscustomobject]@{
        Phase          = 'Exchange'
        CheckpointPath = $currentSnapshotPath
    }
}
