function Get-ArrayaAssessmentPipelineCheckpointPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CheckpointRoot,
        [Parameter(Mandatory = $true)]
        [string]$PhaseName
    )

    $phaseDefinition = Get-ArrayaAssessmentPipelineDefinition | Where-Object { $_.Name -eq $PhaseName } | Select-Object -First 1
    if (-not $phaseDefinition) {
        throw "Unknown assessment phase: $PhaseName"
    }

    return (Join-Path -Path $CheckpointRoot -ChildPath ('{0:d2}-{1}.snapshot.json' -f [int]$phaseDefinition.Order, $PhaseName))
}
