function Export-ArrayaAssessmentPipelineLegacyTenantStatsSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory = $false)]
        [hashtable]$InputSnapshot,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $metadata = [ordered]@{}
    $collectionPlan = [ordered]@{}
    $diagnostics = [ordered]@{}
    if ($InputSnapshot) {
        if ($InputSnapshot.Contains('Metadata') -and ($InputSnapshot['Metadata'] -is [System.Collections.IDictionary])) {
            foreach ($entry in $InputSnapshot['Metadata'].GetEnumerator()) {
                $metadata[[string]$entry.Key] = $entry.Value
            }
        }
        if ($InputSnapshot.Contains('CollectionPlan') -and ($InputSnapshot['CollectionPlan'] -is [System.Collections.IDictionary])) {
            foreach ($entry in $InputSnapshot['CollectionPlan'].GetEnumerator()) {
                $collectionPlan[[string]$entry.Key] = $entry.Value
            }
        }
        if ($InputSnapshot.Contains('Diagnostics') -and ($InputSnapshot['Diagnostics'] -is [System.Collections.IDictionary])) {
            foreach ($entry in $InputSnapshot['Diagnostics'].GetEnumerator()) {
                $diagnostics[[string]$entry.Key] = $entry.Value
            }
        }
    }

    $metadata['GeneratedAt'] = (Get-Date).ToString('o')
    $metadata['OutputProfile'] = $Context.OutputProfile
    $metadata['OutputProfileLabel'] = $Context.OutputProfileLabel
    $metadata['ReportingMode'] = $Context.ReportingMode

    $snapshot = Convert-ArrayaLegacyTenantStatsToSnapshot -TenantStatsHash $TenantStatsHash -Metadata $metadata -CollectionPlan $collectionPlan -Diagnostics $diagnostics
    Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $Path
    return $snapshot
}

function Invoke-ArrayaAssessmentLegacyPipelineStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Identity', 'Exchange', 'Collaboration', 'Endpoint', 'Governance')]
        [string]$PhaseName,
        [Parameter(Mandatory = $true)]
        [string]$StepName,
        [Parameter(Mandatory = $true)]
        [string]$InputSnapshotPath,
        [Parameter(Mandatory = $true)]
        [string]$CheckpointPath,
        [Parameter(Mandatory = $true)]
        [int]$ProgressTotalSteps,
        [Parameter(Mandatory = $true)]
        [int]$ProgressStartingStep,
        [Parameter(Mandatory = $false)]
        [bool]$ReuseCurrentSessions = $false
    )

    $parameters = @{
        ExportPath                        = $Context.ExportPath
        OutputProfile                     = $Context.OutputProfile
        OutputProfileLabel                = $Context.OutputProfileLabel
        ReportingModeOverride             = $Context.ReportingMode
        GenerateWorkbookOverride          = [bool]$Context.GenerateWorkbook
        GenerateTechnicalHtmlOverride     = [bool]$Context.GenerateTechnicalHtml
        GenerateBestPracticesHtmlOverride = [bool]$Context.GenerateBestPracticesHtml
        GenerateQuestionnaireOverride     = [bool]$Context.GenerateQuestionnaire
        GenerateJsonOverride              = [bool]$Context.GenerateJson
        GeneratePdfOverride               = [bool]$Context.GeneratePdf
        PipelinePhase                     = $PhaseName
        PipelinePassThru                  = $true
        PipelineInputSnapshotPath         = $InputSnapshotPath
        PipelineCheckpointPath            = $CheckpointPath
        PipelinePhaseSteps                = @($StepName)
        PipelineProgressTotalSteps        = $ProgressTotalSteps
        PipelineProgressStartingStep      = $ProgressStartingStep
        PipelineSuppressPhaseHeader       = $true
    }

    if ($Context.Contains('SkipHtmlReport') -and [bool]$Context.SkipHtmlReport) { $parameters.SkipHtmlReport = $true }
    if ($Context.Contains('SkipPdfReport') -and [bool]$Context.SkipPdfReport) { $parameters.SkipPdfReport = $true }
    if ($Context.Contains('SkipJsonReport') -and [bool]$Context.SkipJsonReport) { $parameters.SkipJsonReport = $true }
    if ($Context.Contains('StoreTenantStatsGlobal') -and [bool]$Context.StoreTenantStatsGlobal) { $parameters.StoreTenantStatsGlobal = $true }
    if ($Context.Contains('TenantStatsVariableName')) { $parameters.TenantStatsVariableName = [string]$Context.TenantStatsVariableName }

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

    $rawResult = @(Invoke-AssessmentPipelineScript -ScriptPath $script:LegacyAssessmentScriptPath -Parameters $parameters)
    $result = @($rawResult | Where-Object { $_ -is [psobject] -and $_.PSObject.Properties['Phase'] } | Select-Object -Last 1)
    if ($result.Count -gt 0) {
        return $result[0]
    }

    return [pscustomobject]@{
        Phase          = $PhaseName
        CheckpointPath = $CheckpointPath
    }
}

function Resolve-ArrayaAssessmentPipelineCollectionPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    $generateTechnicalHtml = if ($Context.Contains('GenerateTechnicalHtml')) { [bool]$Context.GenerateTechnicalHtml } else { $false }
    $generateWorkbook = if ($Context.Contains('GenerateWorkbook')) { [bool]$Context.GenerateWorkbook } else { $true }
    $generateBestPracticesHtml = if ($Context.Contains('GenerateBestPracticesHtml')) { [bool]$Context.GenerateBestPracticesHtml } else { $false }
    $generateJson = if ($Context.Contains('GenerateJson')) { [bool]$Context.GenerateJson } else { $true }
    $generateQuestionnaire = if ($Context.Contains('GenerateQuestionnaire')) { [bool]$Context.GenerateQuestionnaire } else { $false }
    $reportingMode = if ($Context.Contains('ReportingMode')) { [string]$Context.ReportingMode } else { 'Minimum' }
    $outputProfile = if ($Context.Contains('OutputProfile')) { [string]$Context.OutputProfile } else { 'SolutionsEngineer' }

    $plan = [ordered]@{
        CollectExchangeRecipients          = $true
        CollectEmailActivityDetails        = ($generateTechnicalHtml -or $generateWorkbook -or $generateJson)
        CollectExchangeGroups              = $true
        CollectMailFlowRulesConnectors     = $true
        CollectPublicFolders               = $true
        CollectThirdPartySpamFiltering     = $true
        CollectSmtpRelayConfiguration      = $true
        CollectGovernanceCompliancePolicies = ($generateTechnicalHtml -or $generateWorkbook -or $generateBestPracticesHtml -or $generateJson)
        CollectTeamsDetails                = ($generateTechnicalHtml -or $generateWorkbook -or $generateBestPracticesHtml -or $generateJson)
        CollectTeamsVoiceDetails           = $true
        CollectUnifiedGroups               = $true
        BuildOwnershipGovernanceTables     = ($generateTechnicalHtml -or $generateBestPracticesHtml -or $generateWorkbook -or $generateJson)
        BuildAssessmentReportTables        = ($generateBestPracticesHtml -or $generateWorkbook -or $generateQuestionnaire -or $generateJson)
        BuildConfigurationSummaryTables    = $generateJson
        BuildLicenseClassificationMetadata = ($generateWorkbook -or $generateTechnicalHtml -or $generateBestPracticesHtml -or $generateQuestionnaire -or $generateJson)
        BuildCombinedUserMailboxProjection = ($reportingMode -eq 'All')
    }

    if ($outputProfile -eq 'ExecutiveLevel') {
        $plan.CollectExchangeRecipients = $false
        $plan.CollectEmailActivityDetails = $true
        $plan.CollectExchangeGroups = $false
        $plan.CollectMailFlowRulesConnectors = $false
        $plan.CollectPublicFolders = $false
        $plan.CollectThirdPartySpamFiltering = $false
        $plan.CollectSmtpRelayConfiguration = $false
        $plan.CollectUnifiedGroups = $false
    }

    return $plan
}
