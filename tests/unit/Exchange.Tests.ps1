Describe 'Arraya.M365.Exchange' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Arraya.M365.Exchange.psd1'
        $script:moduleRoot = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange'
        $script:expectedExports = @(
            'Get-AllExchangeMailboxDetails'
            'Get-AllPublicFolderDetails'
            'Get-AllRecipientDetails'
            'Get-ExchangeGroupDetails'
            'Get-ExchangeHybridConfiguration'
            'Get-MailFlowRulesandConnectors'
            'Get-SMTPRelayConfiguration'
            'Get-ThirdPartySpamFilteringConfig'
        )
    }

    It 'has a manifest' {
        Test-Path $script:manifestPath | Should -BeTrue
    }

    It 'exports the reduced exchange collector surface' {
        $manifest = Import-PowerShellDataFile -Path $script:manifestPath
        $manifest.FunctionsToExport | Should -Be $script:expectedExports
        $manifest.VariablesToExport | Should -Be @()
    }

    It 'does not keep retired public exchange wrapper files' {
        @(
            'Public\Get-ArrayaExchangeRuntimeContext.ps1'
            'Public\Set-ArrayaExchangeRuntimeContext.ps1'
            'Public\Get-ResolvedEmailAddresses.ps1'
            'Public\Get-ExoMailboxStatisticsSafe.ps1'
            'Public\Resolve-ExoStatisticsIdentity.ps1'
            'Public\Invoke-ExchangeMailboxStatHelpers.ps1'
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
                'Get-ArrayaExchangeRuntimeContext'
                'Set-ArrayaExchangeRuntimeContext'
                'Get-ArrayaGraphAdminReportSettings'
                'Export-ArrayaGraphReportCsv'
                'Invoke-QuietRestMethod'
            ) -CaseSensitive
        $matches | Should -BeNullOrEmpty
    }
}
