function Invoke-M365TenantAssessmentExportPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory = $false)]
        [hashtable]$ExportTenantStatsHash,
        [Parameter(Mandatory = $true)]
        [string]$ExportDetails,
        [Parameter(Mandatory = $true)]
        [bool]$SkipWorkbook,
        [Parameter(Mandatory = $true)]
        [bool]$SkipBestPracticesHtml,
        [Parameter(Mandatory = $true)]
        [bool]$SkipQuestionnaire,
        [Parameter(Mandatory = $true)]
        [bool]$SkipHtmlReport,
        [Parameter(Mandatory = $true)]
        [bool]$SkipPdfReport,
        [Parameter(Mandatory = $true)]
        [bool]$SkipJsonReport,
        [Parameter(Mandatory = $true)]
        [string]$OutputProfileLabel,
        [Parameter(Mandatory = $true)]
        [string]$ReportingMode,
        [Parameter(Mandatory = $true)]
        [bool]$CollectionOnly,
        [Parameter(Mandatory = $true)]
        [bool]$ExportOnly,
        [Parameter(Mandatory = $false)]
        [string]$LegacyScriptRoot
    )

    function Resolve-LegacyHelperRoot {
        if (-not [string]::IsNullOrWhiteSpace($LegacyScriptRoot)) {
            return $LegacyScriptRoot
        }
        return $PSScriptRoot
    }

    function Get-SupportDirectory {
        $baseDirectory = [System.IO.Path]::GetDirectoryName($ExportDetails)
        if ([string]::IsNullOrWhiteSpace($baseDirectory)) {
            $baseDirectory = (Get-Location).Path
        }

        $supportDirectory = Join-Path -Path $baseDirectory -ChildPath 'Support'
        if (-not (Test-Path -Path $supportDirectory)) {
            $null = New-Item -ItemType Directory -Path $supportDirectory -Force
        }

        return $supportDirectory
    }

    function Test-SuppressedConsoleWarningMessage {
        param([string]$Message)

        if ([string]::IsNullOrWhiteSpace($Message)) {
            return $false
        }

        return $Message -match '(?i)conditional access|inbox rule'
    }

    function Write-PipelineLog {
        param(
            [string]$Type = 'INFO',
            [string]$Message
        )

        if (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
            Write-Log -Type $Type -Message $Message -ExportFileLocation $ExportDetails
            return
        }

        switch ($Type.ToUpperInvariant()) {
            'ERROR' { Write-Error $Message }
            'WARNING' {
                if (Test-SuppressedConsoleWarningMessage -Message $Message) {
                    Write-Verbose $Message
                }
                else {
                    Write-Warning $Message
                }
            }
            default { Write-Host $Message }
        }
    }

    function Try-OpenHtmlArtifact {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,
            [Parameter(Mandatory = $true)]
            [string]$Label
        )

        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -Path $Path)) {
            return
        }

        try {
            Start-Process -FilePath $Path | Out-Null
            Write-PipelineLog -Type INFO -Message "$Label opened in the default browser: $Path"
        }
        catch {
            Write-PipelineLog -Type INFO -Message "Could not auto-open ${Label}: $($_.Exception.Message)"
        }
    }

    $generatedArtifacts = [ordered]@{
        Workbook              = $null
        'Best Practices HTML' = $null
        Questionnaire         = $null
        'Full HTML'           = $null
        PDF                   = $null
        JSON                  = $null
    }

    if (-not $ExportTenantStatsHash) {
        $ExportTenantStatsHash = $TenantStatsHash
    }

    if ($SkipWorkbook) {
        Write-PipelineLog -Type INFO -Message 'Skipping workbook generation because the selected output profile disables workbook output.'
    }
    else {
        try {
            if (Get-Command -Name Ensure-ImportExcelReady -ErrorAction SilentlyContinue) {
                Ensure-ImportExcelReady
            }
            elseif (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
                Import-Module ImportExcel -ErrorAction Stop
            }
            elseif (-not (Get-Module -Name ImportExcel -ErrorAction SilentlyContinue)) {
                Import-Module ImportExcel -ErrorAction Stop
            }

            if (-not (Get-Command -Name Export-HashTableToExcel -ErrorAction SilentlyContinue)) {
                throw 'Export-HashTableToExcel function is unavailable in the current session.'
            }

            Write-PipelineLog -Type INFO -Message "Exporting the Tenant Statistics to $ExportDetails."
            Export-HashTableToExcel -hashtable $ExportTenantStatsHash -ExportDetails $ExportDetails
            $generatedArtifacts['Workbook'] = $ExportDetails
        }
        catch {
            Write-PipelineLog -Type ERROR -Message "Workbook export failed: $($_.Exception.Message)"
        }
    }

    if ($SkipJsonReport) {
        Write-PipelineLog -Type INFO -Message 'Skipping JSON report generation because the selected output profile disables it or -SkipJsonReport was provided.'
    }
    else {
        try {
            if (-not (Get-Command -Name Export-TenantStatsJson -ErrorAction SilentlyContinue)) {
                throw 'Export-TenantStatsJson function is unavailable in the current session.'
            }
            $jsonExportPath = Join-Path -Path (Get-SupportDirectory) -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($ExportDetails) + '.json')
            # Preserve the full collection snapshot for re-export scenarios.
            Export-TenantStatsJson -TenantStatsHash $TenantStatsHash -Path $jsonExportPath
            $generatedArtifacts['JSON'] = $jsonExportPath
            Write-PipelineLog -Type INFO -Message "Exported Tenant Statistics JSON to $jsonExportPath"
        }
        catch {
            Write-PipelineLog -Type WARNING -Message "Unable to export Tenant Statistics JSON: $($_.Exception.Message)"
        }
    }

    if ($SkipQuestionnaire) {
        Write-PipelineLog -Type INFO -Message 'Skipping questionnaire export because the selected output profile disables questionnaire output.'
    }
    else {
        try {
            if (Get-Command -Name Ensure-TenantQuestionnaireHelperLoaded -ErrorAction SilentlyContinue) {
                Ensure-TenantQuestionnaireHelperLoaded
            }
            if (-not (Get-Command -Name Export-TenantToTenantQuestionnaireMarkdown -ErrorAction SilentlyContinue)) {
                $questionnaireHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path (Resolve-LegacyHelperRoot) -ChildPath 'Export-TenantToTenantQuestionnaireMarkdown.ps1'))
                if (Test-Path -Path $questionnaireHelperPath) {
                    . $questionnaireHelperPath
                }
            }
            if (Get-Command -Name Export-TenantToTenantQuestionnaireMarkdown -ErrorAction SilentlyContinue) {
                $resolvedLegacyRoot = if (-not [string]::IsNullOrWhiteSpace($LegacyScriptRoot)) {
                    $LegacyScriptRoot
                }
                else {
                    $PSScriptRoot
                }
                $questionnaireTemplateCandidates = @(
                    [System.IO.Path]::GetFullPath((Join-Path -Path $resolvedLegacyRoot -ChildPath '..\..\..\..\docs\Microsoft 365 Tenant to Tenant Questionnaire.md')),
                    [System.IO.Path]::GetFullPath((Join-Path -Path $resolvedLegacyRoot -ChildPath '..\..\..\..\docs\templates\Microsoft 365 Tenant to Tenant Questionnaire.md'))
                )
                $questionnaireTemplatePath = $questionnaireTemplateCandidates | Where-Object { Test-Path -Path $_ } | Select-Object -First 1
                if (-not $questionnaireTemplatePath) {
                    throw "Questionnaire template not found in expected locations: $($questionnaireTemplateCandidates -join '; ')"
                }

                $questionnaireExportPath = $ExportDetails -replace '\.xlsx$', '-TenantToTenantQuestionnaire.md'
                Export-TenantToTenantQuestionnaireMarkdown -TenantStatsHash $TenantStatsHash -TemplatePath $questionnaireTemplatePath -Path $questionnaireExportPath
                $generatedArtifacts['Questionnaire'] = $questionnaireExportPath
                Write-PipelineLog -Type INFO -Message "Exported Tenant to Tenant Questionnaire to $questionnaireExportPath"
            }
            else {
                Write-PipelineLog -Type WARNING -Message 'Skipping questionnaire export because Export-TenantToTenantQuestionnaireMarkdown is unavailable.'
            }
        }
        catch {
            Write-PipelineLog -Type WARNING -Message "Unable to export Tenant to Tenant Questionnaire: $($_.Exception.Message)"
        }
    }

    if ($SkipBestPracticesHtml) {
        Write-PipelineLog -Type INFO -Message 'Skipping Best Practices Analysis HTML generation because the selected output profile disables it.'
    }
    else {
        try {
            if (Get-Command -Name Ensure-TenantHtmlHelpersLoaded -ErrorAction SilentlyContinue) {
                Ensure-TenantHtmlHelpersLoaded
            }
            $bestPracticesHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path (Resolve-LegacyHelperRoot) -ChildPath 'New-TenantHtmlReport.ps1'))
            if (Test-Path -Path $bestPracticesHelperPath) {
                . $bestPracticesHelperPath
            }
            if (Get-Command -Name New-TenantAssessmentHtmlReport -ErrorAction SilentlyContinue) {
                $assessmentHtmlPath = $ExportDetails -replace '\.xlsx$', '-BestPracticesSnapshot.html'
                $assessmentHtmlResult = New-TenantAssessmentHtmlReport -TenantStatsHash $TenantStatsHash -OutputPath $assessmentHtmlPath
                if ($assessmentHtmlResult.Success) {
                    $generatedArtifacts['Best Practices HTML'] = $assessmentHtmlResult.OutputPath
                    Write-PipelineLog -Type INFO -Message "Best Practices Snapshot HTML report generated: $($assessmentHtmlResult.OutputPath)"
                    Try-OpenHtmlArtifact -Path $assessmentHtmlResult.OutputPath -Label 'Best Practices Snapshot HTML'
                }
                else {
                    Write-PipelineLog -Type WARNING -Message "Best Practices Snapshot HTML report generation failed: $($assessmentHtmlResult.Error)"
                }
            }
            else {
                Write-PipelineLog -Type WARNING -Message 'Skipping Best Practices Snapshot HTML generation because New-TenantAssessmentHtmlReport is unavailable.'
            }
        }
        catch {
            Write-PipelineLog -Type WARNING -Message "Error generating Best Practices Snapshot HTML report: $($_.Exception.Message)"
        }
    }

    if ($SkipHtmlReport) {
        Write-PipelineLog -Type INFO -Message 'Skipping full HTML report generation because the selected output profile disables it or -SkipHtmlReport was provided.'
    }
    else {
        try {
            if (Get-Command -Name Ensure-TenantHtmlHelpersLoaded -ErrorAction SilentlyContinue) {
                Ensure-TenantHtmlHelpersLoaded
            }
            $fullHtmlHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path (Resolve-LegacyHelperRoot) -ChildPath 'New-TenantHtmlReport.ps1'))
            if (Test-Path -Path $fullHtmlHelperPath) {
                . $fullHtmlHelperPath
            }
            if (-not (Get-Command -Name New-TenantHtmlReport -ErrorAction SilentlyContinue)) {
                Write-PipelineLog -Type WARNING -Message 'Skipping full HTML report generation because New-TenantHtmlReport is unavailable.'
            }
            else {
                $reportThresholds = @{
                    LicenseUtilization      = 85
                    MailboxSizeGB           = 50
                    ArchiveSizeGB           = 50
                    SharePointSiteGB        = 1024
                    OneDriveSiteGB          = 1024
                    DeviceStaleMonths       = 6
                    DeviceCompliancePercent = 80
                }
                $htmlExportPath = $ExportDetails -replace '\.xlsx$', '-TenantSnapshot.html'
                $htmlResult = New-TenantHtmlReport -TenantStatsHash $TenantStatsHash -Thresholds $reportThresholds -OutputPath $htmlExportPath
                if ($htmlResult.Success) {
                    $generatedArtifacts['Full HTML'] = $htmlResult.OutputPath
                    Write-PipelineLog -Type INFO -Message "HTML report generated: $($htmlResult.OutputPath)"
                    Try-OpenHtmlArtifact -Path $htmlResult.OutputPath -Label 'Technical HTML report'

                    if ($SkipPdfReport) {
                        # Intentionally no-op for profile-based PDF skips.
                    }
                    elseif (-not (Get-Command -Name Export-TenantHtmlReportPdf -ErrorAction SilentlyContinue)) {
                        Write-PipelineLog -Type WARNING -Message 'Skipping PDF report generation because Export-TenantHtmlReportPdf is unavailable.'
                    }
                    else {
                        try {
                            $pdfExportPath = $htmlResult.OutputPath -replace '\.html$', '.pdf'
                            $pdfResult = Export-TenantHtmlReportPdf -HtmlPath $htmlResult.OutputPath -PdfPath $pdfExportPath
                            if ($pdfResult.Success) {
                                $generatedArtifacts['PDF'] = $pdfResult.PdfPath
                                Write-PipelineLog -Type INFO -Message "PDF report generated: $($pdfResult.PdfPath) using $($pdfResult.Renderer)"
                            }
                            else {
                                Write-PipelineLog -Type WARNING -Message "PDF report generation failed: $($pdfResult.Error)"
                            }
                        }
                        catch {
                            Write-PipelineLog -Type WARNING -Message "Error generating PDF report: $($_.Exception.Message)"
                        }
                    }
                }
                else {
                    Write-PipelineLog -Type ERROR -Message "HTML report generation failed: $($htmlResult.Error)"
                }
            }
        }
        catch {
            Write-PipelineLog -Type ERROR -Message "HTML report generation error: $($_.Exception.Message)"
        }
    }

    try {
        $artifactManifestPath = Write-ArrayaAssessmentArtifactManifest `
            -BaseExportPath $ExportDetails `
            -Artifacts $generatedArtifacts `
            -OutputProfileLabel $OutputProfileLabel `
            -ReportingMode $ReportingMode `
            -CollectionOnly $CollectionOnly `
            -ExportOnly $ExportOnly
        if (-not [string]::IsNullOrWhiteSpace($artifactManifestPath)) {
            $generatedArtifacts['Manifest'] = $artifactManifestPath
        }
    }
    catch {
        Write-PipelineLog -Type WARNING -Message "Unable to write artifact manifest: $($_.Exception.Message)"
    }

    return $generatedArtifacts
}
