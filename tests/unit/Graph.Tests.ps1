Describe 'Arraya.M365.Graph' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Graph\Arraya.M365.Graph.psd1'
        $script:moduleRoot = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Graph'
    }

    It 'has a manifest' {
        Test-Path $script:manifestPath | Should -BeTrue
    }

    It 'exports only the graph collector entrypoint' {
        $manifest = Import-PowerShellDataFile -Path $script:manifestPath
        $manifest.FunctionsToExport | Should -Be @('Get-EntraIDGroups')
        $manifest.VariablesToExport | Should -Be @()
    }

    It 'does not keep retired graph wrapper files' {
        @(
            'Public\Get-ArrayaGraphRuntimeContext.ps1'
            'Public\Set-ArrayaGraphRuntimeContext.ps1'
            'Public\Get-GraphReportFieldValue.ps1'
            'Public\Normalize-GraphReportFieldName.ps1'
            'Public\Convert-GraphReportValueToInt64.ps1'
        ) | ForEach-Object {
            Test-Path (Join-Path $script:moduleRoot $_) | Should -BeFalse
        }
    }

    It 'does not reference legacy graph wrappers or runtime globals' {
        $matches = Get-ChildItem -Path $script:moduleRoot -Recurse -Include *.ps1,*.psm1 |
            Select-String -Pattern @(
                '\$script:tenantStatsHash'
                '\$ExportDetails'
                '\$global:initialStart'
                'Get-ArrayaGraphRuntimeContext'
                'Set-ArrayaGraphRuntimeContext'
                'Invoke-QuietRestMethod'
            ) -CaseSensitive
        $matches | Should -BeNullOrEmpty
    }
}
