function Invoke-ArrayaAssessmentIdentityPhase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $false)]
        [string]$InputSnapshotPath
    )

    if (-not (Test-ArrayaAssessmentPipelineSessionReadiness -Context $Context -PhaseName 'Identity')) {
        Invoke-ArrayaAssessmentConnectionPhase -Context $Context | Out-Null
    }

    $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Identity'
    $detailLevel = ([string]$Context.ReportingMode).ToLowerInvariant()
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
    Write-Host '[2/6] Identity' -ForegroundColor White
    Write-Host ('-' * 72) -ForegroundColor DarkGray

    $identitySteps = @(
        [pscustomobject]@{ Index = 1; Name = 'Users'; Mode = 'Legacy' }
        [pscustomobject]@{ Index = 2; Name = 'Admins'; Mode = 'Legacy' }
        [pscustomobject]@{ Index = 3; Name = 'Entra groups'; Mode = 'Native' }
        [pscustomobject]@{ Index = 4; Name = 'Domains'; Mode = 'Legacy' }
        [pscustomobject]@{ Index = 5; Name = 'Authentication/SSO configuration'; Mode = 'Legacy' }
        [pscustomobject]@{ Index = 6; Name = 'Federation/cross-tenant configuration'; Mode = 'Legacy' }
        [pscustomobject]@{ Index = 7; Name = 'Conditional Access policies'; Mode = 'Legacy' }
        [pscustomobject]@{ Index = 8; Name = 'MFA registration details'; Mode = 'Legacy' }
    )

    foreach ($step in $identitySteps) {
        if ($step.Mode -eq 'Native') {
            $tenantStatsHash = if ($currentSnapshot) {
                Convert-ArrayaSnapshotToLegacyTenantStatsHash -Snapshot $currentSnapshot
            }
            else {
                @{}
            }

            $graphContext = [pscustomobject]@{
                TenantStats        = $tenantStatsHash
                ExportFileLocation = $null
                Policies           = [ordered]@{
                    ReportingMode = [string]$Context.ReportingMode
                }
                Runtime            = [ordered]@{}
                Metadata           = [ordered]@{}
            }

            Invoke-ArrayaAssessmentPipelineStep -Index $step.Index -Total 8 -Name $step.Name -ScriptBlock {
                Get-EntraIDGroups -detailLevel $detailLevel -GraphAuthType SDK -Context $graphContext | Out-Null
            } | Out-Null

            $currentSnapshot = Export-ArrayaAssessmentPipelineLegacyTenantStatsSnapshot -Context $Context -TenantStatsHash $tenantStatsHash -InputSnapshot $currentSnapshot -Path $checkpointPath
            $currentSnapshotPath = $checkpointPath
            continue
        }

        $legacyStepResult = Invoke-ArrayaAssessmentLegacyPipelineStep `
            -Context $Context `
            -PhaseName 'Identity' `
            -StepName $step.Name `
            -InputSnapshotPath $currentSnapshotPath `
            -CheckpointPath $checkpointPath `
            -ProgressTotalSteps 8 `
            -ProgressStartingStep ($step.Index - 1) `
            -ReuseCurrentSessions $true

        if (-not [string]::IsNullOrWhiteSpace([string]$legacyStepResult.CheckpointPath)) {
            $currentSnapshotPath = [string]$legacyStepResult.CheckpointPath
            $currentSnapshot = Import-ArrayaTenantSnapshot -Path $currentSnapshotPath -SkipValidation
        }
    }

    return [pscustomobject]@{
        Phase          = 'Identity'
        CheckpointPath = $currentSnapshotPath
    }
}
