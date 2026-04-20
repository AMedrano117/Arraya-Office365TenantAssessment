function New-ArrayaAssessmentPipelineContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Full', 'CollectOnly', 'ExportOnly', 'PreflightOnly')]
        [string]$Mode,
        [Parameter(Mandatory = $false)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
        [string]$OutputProfile = 'SolutionsEngineer',
        [Parameter(Mandatory = $false)]
        [string]$OutputProfileLabel,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Minimum', 'Operator', 'Combined', 'Automation', 'Geek', 'All')]
        [string]$ReportingMode = 'Minimum',
        [Parameter(Mandatory = $false)]
        [string]$CheckpointRoot,
        [Parameter(Mandatory = $false)]
        [string]$AssessmentJsonPath
    )

    Import-AssessmentPipelineDependencies

    $resolvedExportPath = if ([string]::IsNullOrWhiteSpace($ExportPath)) {
        Get-ArrayaAssessmentOutputRoot -FallbackPath $script:RepoRoot
    }
    else {
        $ExportPath
    }

    $resolvedCheckpointRoot = Resolve-ArrayaAssessmentPipelineCheckpointRoot -CheckpointRoot $CheckpointRoot -ExportPath $resolvedExportPath
    $definitions = @(Get-ArrayaAssessmentPipelineDefinition | Sort-Object -Property Order)
    $phaseStates = [ordered]@{}
    foreach ($definition in $definitions) {
        $phaseStates[$definition.Name] = [ordered]@{
            Name           = $definition.Name
            Order          = [int]$definition.Order
            Status         = 'Pending'
            CheckpointPath = if ($definition.Checkpointed) { Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $resolvedCheckpointRoot -PhaseName $definition.Name } else { $null }
        }
    }

    return [ordered]@{
        SchemaVersion            = 1
        RunId                    = ([guid]::NewGuid()).Guid
        Mode                     = $Mode
        ExportPath               = [System.IO.Path]::GetFullPath($resolvedExportPath)
        OutputProfile            = $OutputProfile
        OutputProfileLabel       = if ([string]::IsNullOrWhiteSpace($OutputProfileLabel)) { $OutputProfile } else { $OutputProfileLabel }
        ReportingMode            = $ReportingMode
        CheckpointRoot           = $resolvedCheckpointRoot
        AssessmentJsonPath       = if ([string]::IsNullOrWhiteSpace($AssessmentJsonPath)) { $null } else { [System.IO.Path]::GetFullPath($AssessmentJsonPath) }
        StartedAt                = (Get-Date).ToString('o')
        LastUpdatedAt            = (Get-Date).ToString('o')
        CurrentPhase             = $null
        LastCheckpointPath       = $null
        PhaseStates              = $phaseStates
    }
}
