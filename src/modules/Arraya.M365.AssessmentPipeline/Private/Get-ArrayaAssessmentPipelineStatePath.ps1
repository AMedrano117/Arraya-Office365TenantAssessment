function Get-ArrayaAssessmentPipelineStatePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CheckpointRoot
    )

    return (Join-Path -Path $CheckpointRoot -ChildPath 'pipeline.state.json')
}
