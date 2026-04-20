function Import-ArrayaAssessmentPipelineState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CheckpointRoot
    )

    $statePath = Get-ArrayaAssessmentPipelineStatePath -CheckpointRoot ([System.IO.Path]::GetFullPath($CheckpointRoot))
    if (-not (Test-Path -Path $statePath -PathType Leaf)) {
        throw "Pipeline state file not found: $statePath"
    }

    return (Get-Content -Path $statePath -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable)
}
