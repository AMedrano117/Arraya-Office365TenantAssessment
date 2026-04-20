function Resume-ArrayaAssessmentPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CheckpointRoot,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Connection', 'TenantOverview', 'Identity', 'Exchange', 'Collaboration', 'Endpoint', 'Governance', 'Export')]
        [string]$ThroughPhase,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Connection', 'TenantOverview', 'Identity', 'Exchange', 'Collaboration', 'Endpoint', 'Governance', 'Export')]
        [string]$Phase
    )

    $invokeParams = @{
        ResumeFromCheckpointRoot = $CheckpointRoot
    }
    if ($PSBoundParameters.ContainsKey('ThroughPhase')) {
        $invokeParams.ThroughPhase = $ThroughPhase
    }
    if ($PSBoundParameters.ContainsKey('Phase')) {
        $invokeParams.Phase = $Phase
    }

    return Invoke-ArrayaAssessmentPipeline @invokeParams
}
