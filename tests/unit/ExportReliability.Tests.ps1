Describe 'Assessment export reliability' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:commonManifest = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        $script:pipelinePath = Join-Path $script:repoRoot 'src\scripts\reporting\Invoke-M365TenantAssessmentExportPipeline.ps1'

        Import-Module -Name $script:commonManifest -Force -WarningAction SilentlyContinue -DisableNameChecking
        . $script:pipelinePath

        $script:stubbedGlobalFunctionNames = @(
            'Write-Log',
            'Write-ProgressHelper',
            'Ensure-ImportExcelReady',
            'Ensure-TenantHtmlHelpersLoaded',
            'Export-PresalesMigrationScopeQuestionnaireMarkdown',
            'New-TenantMigrationScopeHtmlReport',
            'New-TenantAssessmentHtmlReport'
        )
        $script:originalGlobalFunctions = @{}
        foreach ($functionName in $script:stubbedGlobalFunctionNames) {
            $existingFunction = Get-Item -LiteralPath "Function:\$functionName" -ErrorAction SilentlyContinue
            if ($null -ne $existingFunction) {
                $script:originalGlobalFunctions[$functionName] = $existingFunction.ScriptBlock
            }
        }

        function global:Write-Log {
            param(
                [string]$Type,
                [string]$Message,
                [string]$ExportFileLocation,
                [switch]$CaptureError,
                $ErrorRecordVar
            )
        }

        function global:Write-ProgressHelper { param() }
        function global:Ensure-ImportExcelReady { param() }
        function global:Ensure-TenantHtmlHelpersLoaded { param() }
        function global:Export-PresalesMigrationScopeQuestionnaireMarkdown { param() }
        function global:New-TenantMigrationScopeHtmlReport { param() }
        function global:New-TenantAssessmentHtmlReport { param() }

        $script:baseParams = @{
            TenantStatsHash         = @{}
            ExportDetails           = $null
            SkipWorkbook            = $true
            SkipBestPracticesHtml   = $true
            SkipQuestionnaire       = $true
            SkipHtmlReport          = $true
            SkipPdfReport           = $true
            SkipJsonReport          = $true
            OutputProfileLabel      = 'ReliabilityTest'
            ReportingMode           = 'Standard'
            CollectionOnly          = $false
            ExportOnly              = $false
            WorkbookExportPolicy    = 'Default'
            TechnicalHtmlPolicy     = 'Default'
            GenerateMigrationPack   = $false
        }
    }

    AfterAll {
        foreach ($functionName in $script:stubbedGlobalFunctionNames) {
            Remove-Item -LiteralPath "Function:\$functionName" -Force -ErrorAction SilentlyContinue
            if ($script:originalGlobalFunctions.ContainsKey($functionName)) {
                Set-Item -LiteralPath "Function:\$functionName" -Value $script:originalGlobalFunctions[$functionName]
            }
        }
    }

    BeforeEach {
        Mock Test-ArrayaTenantSnapshot {
            [pscustomobject]@{
                Valid    = $true
                Errors   = @()
                Warnings = @()
            }
        }

        Mock Write-ArrayaAssessmentArtifactManifest {
            $manifestPath = Join-Path $TestDrive 'ArtifactManifest.json'
            '{}' | Set-Content -LiteralPath $manifestPath -Encoding UTF8
            return $manifestPath
        }
    }

    It 'fails before writing a manifest when the required workbook exporter throws' {
        Mock Export-HashTableToExcel { throw 'synthetic workbook failure' }

        $params = $script:baseParams.Clone()
        $params.ExportDetails = Join-Path $TestDrive 'WorkbookFailure.xlsx'
        $params.ExportTenantStatsHash = @{}
        $params.SkipWorkbook = $false

        $result = $null
        $caughtError = $null
        try {
            $result = Invoke-M365TenantAssessmentExportPipeline @params
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should -Not -BeNullOrEmpty
        $caughtError.Exception.Message | Should -Match 'Workbook export failed: synthetic workbook failure'
        $result | Should -BeNullOrEmpty
        Should -Invoke Write-ArrayaAssessmentArtifactManifest -Times 0 -Exactly
    }

    It 'fails before writing a manifest when the workbook exporter returns without creating its artifact' {
        Mock Export-HashTableToExcel { }

        $params = $script:baseParams.Clone()
        $params.ExportDetails = Join-Path $TestDrive 'MissingWorkbook.xlsx'
        $params.ExportTenantStatsHash = @{}
        $params.SkipWorkbook = $false

        $caughtError = $null
        try {
            Invoke-M365TenantAssessmentExportPipeline @params
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should -Not -BeNullOrEmpty
        $caughtError.Exception.Message | Should -Match 'Workbook export failed: Workbook export completed without producing the required artifact'
        Test-Path -LiteralPath $params.ExportDetails | Should -BeFalse
        Should -Invoke Write-ArrayaAssessmentArtifactManifest -Times 0 -Exactly
    }

    It 'fails before writing a manifest when the required cutover-pack exporter throws' {
        Mock Export-ArrayaTenantToTenantCutoverPack { throw 'synthetic cutover pack failure' }

        $params = $script:baseParams.Clone()
        $params.ExportDetails = Join-Path $TestDrive 'CutoverFailure.xlsx'
        $params.GenerateMigrationPack = $true

        $caughtError = $null
        try {
            Invoke-M365TenantAssessmentExportPipeline @params
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should -Not -BeNullOrEmpty
        $caughtError.Exception.Message | Should -Match 'cutover pack export failed: synthetic cutover pack failure'
        Should -Invoke Write-ArrayaAssessmentArtifactManifest -Times 0 -Exactly
    }

    It 'fails before writing a manifest when the required questionnaire exporter throws' {
        Mock Export-PresalesMigrationScopeQuestionnaireMarkdown { throw 'synthetic questionnaire failure' }

        $params = $script:baseParams.Clone()
        $params.ExportDetails = Join-Path $TestDrive 'QuestionnaireFailure.xlsx'
        $params.SkipQuestionnaire = $false
        $params.WorkbookExportPolicy = 'Presales'
        $params.TechnicalHtmlPolicy = 'Presales'

        $caughtError = $null
        try {
            Invoke-M365TenantAssessmentExportPipeline @params
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should -Not -BeNullOrEmpty
        $caughtError.Exception.Message | Should -Match 'Questionnaire export failed: synthetic questionnaire failure'
        Should -Invoke Write-ArrayaAssessmentArtifactManifest -Times 0 -Exactly
    }

    It 'fails before writing a manifest when the required technical HTML exporter throws' {
        Mock New-TenantMigrationScopeHtmlReport { throw 'synthetic technical HTML failure' }

        $params = $script:baseParams.Clone()
        $params.ExportDetails = Join-Path $TestDrive 'HtmlFailure.xlsx'
        $params.SkipHtmlReport = $false
        $params.TechnicalHtmlPolicy = 'Presales'

        $caughtError = $null
        try {
            Invoke-M365TenantAssessmentExportPipeline @params
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should -Not -BeNullOrEmpty
        $caughtError.Exception.Message | Should -Match 'Technical HTML export failed: synthetic technical HTML failure'
        Should -Invoke Write-ArrayaAssessmentArtifactManifest -Times 0 -Exactly
    }

    It 'fails when the required manifest writer throws' {
        Mock Write-ArrayaAssessmentArtifactManifest { throw 'synthetic manifest failure' }

        $params = $script:baseParams.Clone()
        $params.ExportDetails = Join-Path $TestDrive 'ManifestFailure.xlsx'

        $caughtError = $null
        try {
            Invoke-M365TenantAssessmentExportPipeline @params
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should -Not -BeNullOrEmpty
        $caughtError.Exception.Message | Should -Match 'Run manifest export failed: synthetic manifest failure'
        Should -Invoke Write-ArrayaAssessmentArtifactManifest -Times 1 -Exactly
    }

    It 'keeps optional best-practices HTML failures non-fatal' {
        Mock New-TenantAssessmentHtmlReport { throw 'synthetic optional HTML failure' }

        $params = $script:baseParams.Clone()
        $params.ExportDetails = Join-Path $TestDrive 'OptionalFailure.xlsx'
        $params.SkipBestPracticesHtml = $false

        $result = Invoke-M365TenantAssessmentExportPipeline @params

        $result['Best Practices HTML'] | Should -BeNullOrEmpty
        Test-Path -LiteralPath $result['Manifest'] -PathType Leaf | Should -BeTrue
    }

    It 'propagates a worksheet export failure instead of reporting workbook success' {
        Mock -ModuleName Arraya.M365.Common Open-ExcelPackage { [pscustomobject]@{} }
        Mock -ModuleName Arraya.M365.Common Export-Excel { throw 'synthetic worksheet failure' }
        Mock -ModuleName Arraya.M365.Common Close-ExcelPackage { }

        $caughtError = $null
        try {
            Export-HashTableToExcel `
                -ExportDetails (Join-Path $TestDrive 'WorksheetFailure.xlsx') `
                -hashtable @{ TenantInfo = [pscustomobject]@{ DisplayName = 'Contoso' } }
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should -Not -BeNullOrEmpty
        $caughtError.Exception.Message | Should -Match 'synthetic worksheet failure'
    }
}
