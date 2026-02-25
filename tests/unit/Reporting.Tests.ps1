Describe 'Arraya.M365.Reporting' {
    It 'has manifest' {
        (Test-Path './src/modules/Arraya.M365.Reporting/Arraya.M365.Reporting.psd1') | Should Be $true
    }

    It 'exports functions through manifest' {
        $manifest = Import-PowerShellDataFile './src/modules/Arraya.M365.Reporting/Arraya.M365.Reporting.psd1'
        $manifest.FunctionsToExport | Should Be '*'
    }
}

