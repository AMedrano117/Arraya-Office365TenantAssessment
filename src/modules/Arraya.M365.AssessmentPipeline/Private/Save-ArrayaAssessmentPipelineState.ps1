function Save-ArrayaAssessmentPipelineState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$State
    )

    if (-not (Test-Path -Path $State.CheckpointRoot -PathType Container)) {
        $null = New-Item -Path $State.CheckpointRoot -ItemType Directory -Force
    }

    $State.LastUpdatedAt = (Get-Date).ToString('o')
    $statePath = Get-ArrayaAssessmentPipelineStatePath -CheckpointRoot $State.CheckpointRoot
    $json = $State | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($statePath, $json, [System.Text.UTF8Encoding]::new($false))
}
