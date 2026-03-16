Describe 'Arraya.M365.Common' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        $script:expectedExports = @(
            'Convert-ArrayaLegacyTenantStatsToSnapshot'
            'Convert-ArrayaObjectToArray'
            'Convert-ArrayaSnapshotToLegacyTenantStatsHash'
            'Convert-ArrayaToDate'
            'Convert-ArrayaToNumber'
            'ConvertTo-ExportFriendlyRecord'
            'ConvertTo-ExportFriendlyValue'
            'Export-ArrayaErrorReports'
            'Export-ArrayaTenantSnapshot'
            'Export-HashTableToExcel'
            'Filter-TenantStatsHash'
            'Get-ArrayaAssessmentOutputProfilePolicy'
            'Get-ArrayaAssessmentOutputRoot'
            'Get-ArrayaObjectValue'
            'Import-ArrayaOffice365CustomLocal'
            'Import-ArrayaTenantSnapshot'
            'Invoke-ArrayaCollectionStepSafe'
            'Invoke-QuietCommand'
            'New-ArrayaAssessmentContext'
            'New-ArrayaTenantSnapshot'
            'Test-ArrayaTenantSnapshot'
            'Update-ArrayaTenantSnapshot'
            'Write-ArrayaAssessmentArtifactManifest'
        )
        $script:retiredPublicFiles = @(
            'src\modules\Arraya.M365.Common\Public\Clear-ArrayaAssessmentRuntimeState.ps1'
            'src\modules\Arraya.M365.Common\Public\Convert-HashToArray.ps1'
            'src\modules\Arraya.M365.Common\Public\Export-ArrayaGraphReportCsv.ps1'
            'src\modules\Arraya.M365.Common\Public\Export-ErrorReports.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaAssessmentRuntimeState.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaCollectionDepthPolicy.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaEntraGroupClassification.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaGraphAdminReportSettings.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaGraphResource.ps1'
            'src\modules\Arraya.M365.Common\Public\Invoke-ArrayaRetry.ps1'
            'src\modules\Arraya.M365.Common\Public\Set-ArrayaAssessmentRuntimeState.ps1'
            'src\modules\Arraya.M365.Common\Public\Write-ArrayaLog.ps1'
        )
    }

    It 'has a manifest' {
        Test-Path $script:manifestPath | Should -BeTrue
    }

    It 'exports the explicit common surface' {
        $manifest = Import-PowerShellDataFile -Path $script:manifestPath
        $manifest.FunctionsToExport | Should -Be $script:expectedExports
        $manifest.VariablesToExport | Should -Be @()
    }

    It 'does not export retired wrapper commands' {
        $manifest = Import-PowerShellDataFile -Path $script:manifestPath
        @(
            'Get-ArrayaGraphResource'
            'Get-ArrayaGraphAdminReportSettings'
            'Export-ArrayaGraphReportCsv'
            'Get-ArrayaAssessmentRuntimeState'
            'Set-ArrayaAssessmentRuntimeState'
            'Clear-ArrayaAssessmentRuntimeState'
            'Invoke-ArrayaRetry'
            'Write-ArrayaLog'
            'Convert-HashToArray'
            'Export-ErrorReports'
        ) | ForEach-Object {
            $manifest.FunctionsToExport -contains $_ | Should -BeFalse
        }
    }

    It 'does not keep retired common wrapper files' {
        foreach ($relativePath in $script:retiredPublicFiles) {
            Test-Path (Join-Path $script:repoRoot $relativePath) | Should -BeFalse
        }
    }

    It 'writes error exports into a Debugging folder beside the base artifact' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $exportFileLocation = Join-Path $TestDrive 'Tenant Discovery Report-SolutionsEngineer.xlsx'
        $errorSummary = Export-ArrayaErrorReports -ExportFileLocation $exportFileLocation -ErrorData @(
            [pscustomobject]@{
                Message = 'Example failure'
                Step    = 'UnitTest'
            }
        )

        $expectedDirectory = Join-Path (Split-Path -Path $exportFileLocation -Parent) 'Debugging'
        $errorSummary.FolderPath | Should -Be $expectedDirectory
        Split-Path -Path $errorSummary.JsonPath -Parent | Should -Be $expectedDirectory
        Split-Path -Path $errorSummary.LogPath -Parent | Should -Be $expectedDirectory
        Split-Path -Path $errorSummary.CsvPath -Parent | Should -Be $expectedDirectory
        Test-Path (Join-Path $expectedDirectory 'Tenant Discovery Report-SolutionsEngineer Error Reporting') | Should -BeFalse
    }
}
