function Test-ArrayaAssessmentPipelineDependencies {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $true)]
        [string]$PhaseName,
        [Parameter(Mandatory = $false)]
        [string]$InputSnapshotPath
    )

    $phaseDefinition = Get-ArrayaAssessmentPipelineDefinition | Where-Object { $_.Name -eq $PhaseName } | Select-Object -First 1
    if (-not $phaseDefinition) {
        throw "Unknown assessment phase: $PhaseName"
    }

    if ($phaseDefinition.DependsOn.Count -eq 0) {
        return $true
    }

    $phaseStates = if ($Context.Contains('PhaseStates') -and $Context.PhaseStates) { $Context.PhaseStates } else { @{} }
    foreach ($dependency in @($phaseDefinition.DependsOn)) {
        if ($phaseStates.Contains($dependency) -and [string]$phaseStates[$dependency].Status -eq 'Completed') {
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($InputSnapshotPath)) {
            return $true
        }

        throw "Phase '$PhaseName' depends on '$dependency', but no completed checkpoint or input snapshot was available."
    }

    return $true
}
