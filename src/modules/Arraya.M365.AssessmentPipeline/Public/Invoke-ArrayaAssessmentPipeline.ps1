function Invoke-ArrayaAssessmentPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('Full', 'CollectOnly', 'ExportOnly', 'PreflightOnly')]
        [string]$Mode = 'Full',
        [Parameter(Mandatory = $false)]
        [ValidateSet('Connection', 'TenantOverview', 'Identity', 'Exchange', 'Collaboration', 'Endpoint', 'Governance', 'Export')]
        [string]$ThroughPhase,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Connection', 'TenantOverview', 'Identity', 'Exchange', 'Collaboration', 'Endpoint', 'Governance', 'Export')]
        [string]$Phase,
        [Parameter(Mandatory = $false)]
        [string]$CheckpointRoot,
        [Parameter(Mandatory = $false)]
        [string]$ResumeFromCheckpointRoot,
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
        [string]$AssessmentJsonPath,
        [Parameter(Mandatory = $false)]
        [bool]$GenerateWorkbook = $true,
        [Parameter(Mandatory = $false)]
        [bool]$GenerateTechnicalHtml = $false,
        [Parameter(Mandatory = $false)]
        [bool]$GenerateBestPracticesHtml = $false,
        [Parameter(Mandatory = $false)]
        [bool]$GenerateQuestionnaire = $false,
        [Parameter(Mandatory = $false)]
        [bool]$GenerateJson = $true,
        [Parameter(Mandatory = $false)]
        [bool]$GeneratePdf = $false,
        [Parameter(Mandatory = $false)]
        [switch]$SkipHtmlReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPdfReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipJsonReport,
        [Parameter(Mandatory = $false)]
        [switch]$StoreTenantStatsGlobal,
        [Parameter(Mandatory = $false)]
        [string]$TenantStatsVariableName = 'ArrayaTenantStats',
        [Parameter(Mandatory = $false)]
        [switch]$SkipAuth,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPermissionPreflight,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthMode,
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret
    )

    Import-AssessmentPipelineDependencies

    $context = if (-not [string]::IsNullOrWhiteSpace($ResumeFromCheckpointRoot)) {
        Import-ArrayaAssessmentPipelineState -CheckpointRoot $ResumeFromCheckpointRoot
    }
    else {
        New-ArrayaAssessmentPipelineContext `
            -Mode $Mode `
            -ExportPath $ExportPath `
            -OutputProfile $OutputProfile `
            -OutputProfileLabel $OutputProfileLabel `
            -ReportingMode $ReportingMode `
            -CheckpointRoot $CheckpointRoot `
            -AssessmentJsonPath $AssessmentJsonPath
    }

    if ($PSBoundParameters.ContainsKey('GenerateWorkbook') -or -not $context.Contains('GenerateWorkbook')) { $context.GenerateWorkbook = [bool]$GenerateWorkbook }
    if ($PSBoundParameters.ContainsKey('GenerateTechnicalHtml') -or -not $context.Contains('GenerateTechnicalHtml')) { $context.GenerateTechnicalHtml = [bool]$GenerateTechnicalHtml }
    if ($PSBoundParameters.ContainsKey('GenerateBestPracticesHtml') -or -not $context.Contains('GenerateBestPracticesHtml')) { $context.GenerateBestPracticesHtml = [bool]$GenerateBestPracticesHtml }
    if ($PSBoundParameters.ContainsKey('GenerateQuestionnaire') -or -not $context.Contains('GenerateQuestionnaire')) { $context.GenerateQuestionnaire = [bool]$GenerateQuestionnaire }
    if ($PSBoundParameters.ContainsKey('GenerateJson') -or -not $context.Contains('GenerateJson')) { $context.GenerateJson = [bool]$GenerateJson }
    if ($PSBoundParameters.ContainsKey('GeneratePdf') -or -not $context.Contains('GeneratePdf')) { $context.GeneratePdf = [bool]$GeneratePdf }
    if ($PSBoundParameters.ContainsKey('SkipHtmlReport') -or -not $context.Contains('SkipHtmlReport')) { $context.SkipHtmlReport = [bool]$SkipHtmlReport }
    if ($PSBoundParameters.ContainsKey('SkipPdfReport') -or -not $context.Contains('SkipPdfReport')) { $context.SkipPdfReport = [bool]$SkipPdfReport }
    if ($PSBoundParameters.ContainsKey('SkipJsonReport') -or -not $context.Contains('SkipJsonReport')) { $context.SkipJsonReport = [bool]$SkipJsonReport }
    if ($PSBoundParameters.ContainsKey('StoreTenantStatsGlobal') -or -not $context.Contains('StoreTenantStatsGlobal')) { $context.StoreTenantStatsGlobal = [bool]$StoreTenantStatsGlobal }
    if ($PSBoundParameters.ContainsKey('TenantStatsVariableName') -or -not $context.Contains('TenantStatsVariableName')) { $context.TenantStatsVariableName = $TenantStatsVariableName }
    if ($PSBoundParameters.ContainsKey('SkipAuth') -or -not $context.Contains('SkipAuth')) { $context.SkipAuth = [bool]$SkipAuth }
    if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight') -or -not $context.Contains('SkipPermissionPreflight')) { $context.SkipPermissionPreflight = [bool]$SkipPermissionPreflight }
    if ($PSBoundParameters.ContainsKey('AuthMode') -or -not $context.Contains('AuthMode')) { $context.AuthMode = $AuthMode }
    if ($PSBoundParameters.ContainsKey('TenantId') -or -not $context.Contains('TenantId')) { $context.TenantId = $TenantId }
    if ($PSBoundParameters.ContainsKey('CertificateThumbprint') -or -not $context.Contains('CertificateThumbprint')) { $context.CertificateThumbprint = $CertificateThumbprint }
    if ($PSBoundParameters.ContainsKey('ClientId') -or -not $context.Contains('ClientId')) { $context.ClientId = $ClientId }
    if ($PSBoundParameters.ContainsKey('ClientSecret') -or -not $context.Contains('ClientSecret')) { $context.ClientSecret = $ClientSecret }

    Save-ArrayaAssessmentPipelineState -State $context

    $sequence = @(Resolve-ArrayaAssessmentPipelineInvocationSequence -Context $context -ThroughPhase $ThroughPhase -Phase $Phase)
    $completedConnectionInInvocation = $false
    foreach ($phaseName in $sequence) {
        if ($context.PhaseStates.Contains($phaseName) -and [string]$context.PhaseStates[$phaseName].Status -eq 'Completed') {
            if (-not [string]::IsNullOrWhiteSpace($context.PhaseStates[$phaseName].CheckpointPath)) {
                $context.LastCheckpointPath = [string]$context.PhaseStates[$phaseName].CheckpointPath
            }
            continue
        }

        $inputSnapshotPath = $null
        if ($phaseName -eq 'Export') {
            if (-not [string]::IsNullOrWhiteSpace([string]$context.LastCheckpointPath)) {
                $inputSnapshotPath = [string]$context.LastCheckpointPath
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$context.AssessmentJsonPath)) {
                $inputSnapshotPath = [string]$context.AssessmentJsonPath
            }
        }
        elseif ($phaseName -ne 'Connection') {
            if (-not [string]::IsNullOrWhiteSpace([string]$context.LastCheckpointPath)) {
                $inputSnapshotPath = [string]$context.LastCheckpointPath
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$context.AssessmentJsonPath)) {
                $inputSnapshotPath = [string]$context.AssessmentJsonPath
            }
        }

        Test-ArrayaAssessmentPipelineDependencies -Context $context -PhaseName $phaseName -InputSnapshotPath $inputSnapshotPath | Out-Null

        if (
            -not $completedConnectionInInvocation -and
            $phaseName -ne 'Connection' -and
            $context.PhaseStates.Contains('Connection') -and
            [string]$context.PhaseStates['Connection'].Status -eq 'Completed' -and
            (Test-ArrayaAssessmentPipelineSessionReadiness -Context $context -PhaseName $phaseName)
        ) {
            $completedConnectionInInvocation = $true
        }

        $context.CurrentPhase = $phaseName
        $context.PhaseStates[$phaseName].Status = 'InProgress'
        $context.PhaseStates[$phaseName].StartedAt = (Get-Date).ToString('o')
        Save-ArrayaAssessmentPipelineState -State $context

        $phaseStart = Get-Date
        switch ($phaseName) {
            'Connection' {
                $phaseResult = Invoke-ArrayaAssessmentConnectionPhase -Context $context
            }
            'TenantOverview' {
                $phaseResult = Invoke-ArrayaAssessmentTenantOverviewPhase -Context $context -InputSnapshotPath $inputSnapshotPath
            }
            'Identity' {
                $phaseResult = Invoke-ArrayaAssessmentIdentityPhase -Context $context -InputSnapshotPath $inputSnapshotPath
            }
            'Exchange' {
                $phaseResult = Invoke-ArrayaAssessmentExchangePhase -Context $context -InputSnapshotPath $inputSnapshotPath
            }
            'Collaboration' {
                $phaseResult = Invoke-ArrayaAssessmentCollaborationPhase -Context $context -InputSnapshotPath $inputSnapshotPath
            }
            'Endpoint' {
                $phaseResult = Invoke-ArrayaAssessmentEndpointPhase -Context $context -InputSnapshotPath $inputSnapshotPath
            }
            default {
                $phaseResult = Invoke-ArrayaLegacyAssessmentPhaseBridge `
                    -Context $context `
                    -PhaseName $phaseName `
                    -InputSnapshotPath $inputSnapshotPath `
                    -ReuseCurrentSessions:$completedConnectionInInvocation
            }
        }
        $phaseEnd = Get-Date

        $context.PhaseStates[$phaseName].Status = 'Completed'
        $context.PhaseStates[$phaseName].CompletedAt = if ($phaseResult.PSObject.Properties['CompletedAt'] -and $phaseResult.CompletedAt) { [datetime]$phaseResult.CompletedAt } else { $phaseEnd }
        $context.PhaseStates[$phaseName].DurationSeconds = if ($phaseResult.PSObject.Properties['DurationSeconds'] -and $null -ne $phaseResult.DurationSeconds) { [double]$phaseResult.DurationSeconds } else { [math]::Round((New-TimeSpan -Start $phaseStart -End $phaseEnd).TotalSeconds, 3) }
        $context.PhaseStates[$phaseName].CheckpointPath = $phaseResult.CheckpointPath
        if ($phaseResult.PSObject.Properties['GeneratedArtifacts'] -and $phaseResult.GeneratedArtifacts) {
            $context.PhaseStates[$phaseName].GeneratedArtifacts = $phaseResult.GeneratedArtifacts
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$phaseResult.CheckpointPath)) {
            $context.LastCheckpointPath = [string]$phaseResult.CheckpointPath
        }
        if ($phaseName -eq 'Connection') {
            $completedConnectionInInvocation = $true
        }

        Save-ArrayaAssessmentPipelineState -State $context
    }

    $context.CurrentPhase = $null
    Save-ArrayaAssessmentPipelineState -State $context
    return [pscustomobject]$context
}
