function Invoke-ArrayaLegacyAssessmentPhaseBridge {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $true)]
        [string]$PhaseName,
        [Parameter(Mandatory = $false)]
        [string]$InputSnapshotPath,
        [Parameter(Mandatory = $false)]
        [bool]$ReuseCurrentSessions = $false
    )

    $parameters = @{
        ExportPath                     = $Context.ExportPath
        OutputProfile                  = $Context.OutputProfile
        OutputProfileLabel             = $Context.OutputProfileLabel
        ReportingModeOverride          = $Context.ReportingMode
        GenerateWorkbookOverride       = [bool]$Context.GenerateWorkbook
        GenerateTechnicalHtmlOverride  = [bool]$Context.GenerateTechnicalHtml
        GenerateBestPracticesHtmlOverride = [bool]$Context.GenerateBestPracticesHtml
        GenerateQuestionnaireOverride  = [bool]$Context.GenerateQuestionnaire
        GenerateJsonOverride           = [bool]$Context.GenerateJson
        GeneratePdfOverride            = [bool]$Context.GeneratePdf
        PipelinePhase                  = $PhaseName
        PipelinePassThru               = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($Context.CheckpointRoot) -and $PhaseName -ne 'Export') {
        $parameters.PipelineCheckpointPath = (Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName $PhaseName)
    }
    if (-not [string]::IsNullOrWhiteSpace($InputSnapshotPath)) {
        $parameters.PipelineInputSnapshotPath = $InputSnapshotPath
    }
    if ($Context.Contains('SkipHtmlReport') -and [bool]$Context.SkipHtmlReport) { $parameters.SkipHtmlReport = $true }
    if ($Context.Contains('SkipPdfReport') -and [bool]$Context.SkipPdfReport) { $parameters.SkipPdfReport = $true }
    if ($Context.Contains('SkipJsonReport') -and [bool]$Context.SkipJsonReport) { $parameters.SkipJsonReport = $true }
    if ($Context.Contains('StoreTenantStatsGlobal') -and [bool]$Context.StoreTenantStatsGlobal) { $parameters.StoreTenantStatsGlobal = $true }
    if ($Context.Contains('TenantStatsVariableName')) { $parameters.TenantStatsVariableName = [string]$Context.TenantStatsVariableName }

    if ($PhaseName -ne 'Export') {
        $effectiveSkipAuth = if ($ReuseCurrentSessions) { $true } else { [bool]$Context.SkipAuth }
        $effectiveSkipPermissionPreflight = if ($ReuseCurrentSessions) { $true } else { [bool]$Context.SkipPermissionPreflight }
        if ($effectiveSkipAuth) { $parameters.SkipAuth = $true }
        if ($effectiveSkipPermissionPreflight) { $parameters.SkipPermissionPreflight = $true }
        if ($ReuseCurrentSessions) { $parameters.PipelineSkipConnectionBootstrap = $true }
        if (-not [string]::IsNullOrWhiteSpace([string]$Context.AuthMode)) { $parameters.AuthMode = [string]$Context.AuthMode }
        if (-not [string]::IsNullOrWhiteSpace([string]$Context.TenantId)) { $parameters.TenantId = [string]$Context.TenantId }
        if (-not [string]::IsNullOrWhiteSpace([string]$Context.CertificateThumbprint)) { $parameters.CertificateThumbprint = [string]$Context.CertificateThumbprint }
        if (-not [string]::IsNullOrWhiteSpace([string]$Context.ClientId)) { $parameters.ClientId = [string]$Context.ClientId }
        if (-not [string]::IsNullOrWhiteSpace([string]$Context.ClientSecret)) { $parameters.ClientSecret = [string]$Context.ClientSecret }
    }

    $phaseStart = Get-Date
    $rawResult = @(Invoke-AssessmentPipelineScript -ScriptPath $script:LegacyAssessmentScriptPath -Parameters $parameters)
    $phaseEnd = Get-Date
    $result = @($rawResult | Where-Object { $_ -is [psobject] -and $_.PSObject.Properties['Phase'] } | Select-Object -Last 1)
    $phaseResult = if ($result.Count -gt 0) { $result[0] } else { [pscustomobject]@{ Phase = $PhaseName } }

    return [pscustomobject]@{
        Phase           = $PhaseName
        CheckpointPath  = if ($phaseResult.PSObject.Properties['CheckpointPath']) { [string]$phaseResult.CheckpointPath } else { $null }
        ExportFilePath  = if ($phaseResult.PSObject.Properties['ExportFileLocation']) { [string]$phaseResult.ExportFileLocation } else { $null }
        GeneratedArtifacts = if ($phaseResult.PSObject.Properties['GeneratedArtifacts']) { $phaseResult.GeneratedArtifacts } else { $null }
        StartedAt       = $phaseStart
        CompletedAt     = $phaseEnd
        DurationSeconds = [math]::Round((New-TimeSpan -Start $phaseStart -End $phaseEnd).TotalSeconds, 3)
    }
}
