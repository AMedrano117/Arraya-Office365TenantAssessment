Describe 'Arraya.M365.Reporting' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Reporting\Arraya.M365.Reporting.psd1'
        $script:moduleRoot = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Reporting'
        $script:expectedExports = @(
            'Get-ArrayaAssessmentRecommendationText'
            'Get-ArrayaAssessmentWorksheetName'
            'Get-ArrayaEmployeeExperienceInsightsAnalysis'
            'Get-GraphUserStats'
        )
    }

    It 'has a manifest' {
        Test-Path $script:manifestPath | Should -BeTrue
    }

    It 'exports the reduced reporting surface' {
        $manifest = Import-PowerShellDataFile -Path $script:manifestPath
        $manifest.FunctionsToExport | Should -Be $script:expectedExports
        $manifest.VariablesToExport | Should -Be @()
    }

    It 'does not keep retired reporting wrapper files' {
        @(
            'Public\Get-ArrayaReportingRuntimeContext.ps1'
            'Public\Set-ArrayaReportingRuntimeContext.ps1'
            'Public\New-ArrayaFinding.ps1'
        ) | ForEach-Object {
            Test-Path (Join-Path $script:moduleRoot $_) | Should -BeFalse
        }
    }

    It 'does not reference retired wrappers or hidden runtime globals in source' {
        $matches = Get-ChildItem -Path $script:moduleRoot -Recurse -Include *.ps1,*.psm1 |
            Select-String -Pattern @(
                '\$script:tenantStatsHash'
                '\$ExportDetails'
                '\$global:initialStart'
                'Get-ArrayaReportingRuntimeContext'
                'Set-ArrayaReportingRuntimeContext'
                'Invoke-QuietRestMethod'
            ) -CaseSensitive
        $matches | Should -BeNullOrEmpty
    }
}
