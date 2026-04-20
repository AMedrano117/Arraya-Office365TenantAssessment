function Invoke-ArrayaAssessmentEndpointPhase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $false)]
        [string]$InputSnapshotPath
    )

    if (-not (Test-ArrayaAssessmentPipelineSessionReadiness -Context $Context -PhaseName 'Endpoint')) {
        Invoke-ArrayaAssessmentConnectionPhase -Context $Context | Out-Null
    }

    $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Endpoint'
    $currentSnapshotPath = $InputSnapshotPath
    $currentSnapshot = if (-not [string]::IsNullOrWhiteSpace($currentSnapshotPath)) {
        Import-ArrayaTenantSnapshot -Path $currentSnapshotPath -SkipValidation
    }
    else {
        $null
    }

    Write-Host ("Output profile: {0} (scope: {1})" -f $Context.OutputProfileLabel, ([string]$Context.ReportingMode).ToLowerInvariant()) -ForegroundColor DarkGray
    Write-Host 'Microsoft 365 Tenant Assessment' -ForegroundColor White
    Write-Host ''
    Write-Host '[5/6] Endpoint' -ForegroundColor White
    Write-Host ('-' * 72) -ForegroundColor DarkGray

    $endpointSteps = @(
        [pscustomobject]@{ Index = 1; Name = 'Devices' }
        [pscustomobject]@{ Index = 2; Name = 'Endpoint operational summaries' }
    )

    foreach ($step in $endpointSteps) {
        $legacyStepResult = Invoke-ArrayaAssessmentLegacyPipelineStep `
            -Context $Context `
            -PhaseName 'Endpoint' `
            -StepName $step.Name `
            -InputSnapshotPath $currentSnapshotPath `
            -CheckpointPath $checkpointPath `
            -ProgressTotalSteps 2 `
            -ProgressStartingStep ($step.Index - 1) `
            -ReuseCurrentSessions $true

        if (-not [string]::IsNullOrWhiteSpace([string]$legacyStepResult.CheckpointPath)) {
            $currentSnapshotPath = [string]$legacyStepResult.CheckpointPath
        }
    }

    if ([string]::IsNullOrWhiteSpace($currentSnapshotPath) -or -not (Test-Path -Path $currentSnapshotPath)) {
        $tenantStatsHash = if ($currentSnapshot) {
            Convert-ArrayaSnapshotToLegacyTenantStatsHash -Snapshot $currentSnapshot
        }
        else {
            @{}
        }
        $currentSnapshot = Export-ArrayaAssessmentPipelineLegacyTenantStatsSnapshot -Context $Context -TenantStatsHash $tenantStatsHash -InputSnapshot $currentSnapshot -Path $checkpointPath
        $currentSnapshotPath = $checkpointPath
    }

    return [pscustomobject]@{
        Phase          = 'Endpoint'
        CheckpointPath = $currentSnapshotPath
    }
}
