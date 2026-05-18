Describe 'Export Pipeline - Integration' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:commonManifest = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        Import-Module -Name $script:commonManifest -Force -WarningAction SilentlyContinue -DisableNameChecking

        $script:pipelinePath = Join-Path $script:repoRoot 'src\scripts\reporting\Invoke-M365TenantAssessmentExportPipeline.ps1'
        . $script:pipelinePath

        $script:baseInvokeParams = @{
            SkipWorkbook           = $true
            SkipBestPracticesHtml  = $true
            SkipQuestionnaire      = $true
            SkipHtmlReport         = $true
            SkipPdfReport          = $true
            SkipJsonReport         = $true
            OutputProfileLabel     = 'TestRun'
            ReportingMode          = 'Standard'
            CollectionOnly         = $false
            ExportOnly             = $false
        }
    }

    It 'rejects a null snapshot before writing any artifact' {
        $params = $script:baseInvokeParams.Clone()
        $params.TenantStatsHash = $null
        $params.ExportDetails   = Join-Path $TestDrive 'NullTest.xlsx'

        { Invoke-M365TenantAssessmentExportPipeline @params } | Should -Throw
    }

    It 'rejects a non-hashtable snapshot' {
        $params = $script:baseInvokeParams.Clone()
        $params.TenantStatsHash = 'not-a-hashtable'
        $params.ExportDetails   = Join-Path $TestDrive 'BadType.xlsx'

        { Invoke-M365TenantAssessmentExportPipeline @params } | Should -Throw
    }

    It 'completes without throwing for a minimal valid v2 snapshot with all outputs skipped' {
        $snapshot = New-ArrayaTenantSnapshot -Metadata @{ OutputProfileLabel = 'TestRun' }
        $params = $script:baseInvokeParams.Clone()
        $params.TenantStatsHash = $snapshot
        $params.ExportDetails   = Join-Path $TestDrive 'MinimalValid.xlsx'

        { Invoke-M365TenantAssessmentExportPipeline @params } | Should -Not -Throw
    }

    It 'creates a Support directory relative to the export path' {
        $snapshot   = New-ArrayaTenantSnapshot -Metadata @{ OutputProfileLabel = 'TestRun' }
        $exportDir  = Join-Path $TestDrive 'SupportDirTest'
        $null       = New-Item -ItemType Directory -Path $exportDir -Force
        $exportPath = Join-Path $exportDir 'Deliverables\Assessment.xlsx'

        $params = $script:baseInvokeParams.Clone()
        $params.TenantStatsHash = $snapshot
        $params.ExportDetails   = $exportPath

        Invoke-M365TenantAssessmentExportPipeline @params

        Test-Path (Join-Path $exportDir 'Support') | Should -BeTrue
    }

    It 'reports a missing data domain as a warning not an error' {
        # Build a snapshot that explicitly omits the Exchange domain
        $snapshot = New-ArrayaTenantSnapshot -Metadata @{ OutputProfileLabel = 'TestRun' }
        $snapshot['Data'].Remove('Exchange')

        $validation = Test-ArrayaTenantSnapshot -Snapshot $snapshot -Purpose Export

        $validation.Valid | Should -BeTrue
        ($validation.Warnings | Where-Object { $_ -match 'Exchange' }).Count | Should -BeGreaterThan 0
    }

    It 'is valid with no warnings when all data domains are present' {
        $snapshot = New-ArrayaTenantSnapshot -Metadata @{ OutputProfileLabel = 'TestRun' }
        $validation = Test-ArrayaTenantSnapshot -Snapshot $snapshot -Purpose Export

        $validation.Valid | Should -BeTrue
        ($validation.Warnings | Where-Object { $_ -match 'Data domain' }).Count | Should -Be 0
    }
}
