Describe 'Arraya.M365.Graph' {
    It 'has manifest' {
        (Test-Path './src/modules/Arraya.M365.Graph/Arraya.M365.Graph.psd1') | Should Be $true
    }

    It 'exports functions through manifest' {
        $manifest = Import-PowerShellDataFile './src/modules/Arraya.M365.Graph/Arraya.M365.Graph.psd1'
        $manifest.FunctionsToExport | Should Be '*'
    }
}

