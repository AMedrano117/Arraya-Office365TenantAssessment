function Get-ArrayaAssessmentPipelineCollaborationServiceSelection {
    [CmdletBinding()]
    param()

    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
    $teamsInventoryService = if ($graphContext) { 'MGGraph' } else { 'Teams' }
    $sharePointService = if ($graphContext) { 'API' } else { 'API' }

    return [pscustomobject]@{
        SharePointService    = $sharePointService
        TeamsInventoryService = $teamsInventoryService
    }
}

function Invoke-ArrayaAssessmentCollaborationPhase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $false)]
        [string]$InputSnapshotPath
    )

    if (-not (Test-ArrayaAssessmentPipelineSessionReadiness -Context $Context -PhaseName 'Collaboration')) {
        Invoke-ArrayaAssessmentConnectionPhase -Context $Context | Out-Null
    }

    $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Collaboration'
    $serviceSelection = Get-ArrayaAssessmentPipelineCollaborationServiceSelection
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
    Write-Host '[4/6] Collaboration' -ForegroundColor White
    Write-Host ('-' * 72) -ForegroundColor DarkGray

    $collaborationSteps = @(
        [pscustomobject]@{ Index = 1; Name = 'Unified groups' }
        [pscustomobject]@{ Index = 2; Name = "SharePoint/OneDrive sites ($($serviceSelection.SharePointService))" }
        [pscustomobject]@{ Index = 3; Name = 'Teams voice details' }
        [pscustomobject]@{ Index = 4; Name = "Teams inventory ($($serviceSelection.TeamsInventoryService))" }
    )

    foreach ($step in $collaborationSteps) {
        $legacyStepResult = Invoke-ArrayaAssessmentLegacyPipelineStep `
            -Context $Context `
            -PhaseName 'Collaboration' `
            -StepName $step.Name `
            -InputSnapshotPath $currentSnapshotPath `
            -CheckpointPath $checkpointPath `
            -ProgressTotalSteps 4 `
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
        Phase          = 'Collaboration'
        CheckpointPath = $currentSnapshotPath
    }
}
