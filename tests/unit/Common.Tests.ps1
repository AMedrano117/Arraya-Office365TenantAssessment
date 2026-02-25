Describe 'Arraya.M365.Common' {
    It 'has manifest' {
        (Test-Path './src/modules/Arraya.M365.Common/Arraya.M365.Common.psd1') | Should Be $true
    }

    It 'exports functions through manifest' {
        $manifest = Import-PowerShellDataFile './src/modules/Arraya.M365.Common/Arraya.M365.Common.psd1'
        $manifest.FunctionsToExport | Should Be '*'
    }

    It 'has Office365Custom local loader function' {
        (Test-Path './src/modules/Arraya.M365.Common/Public/Import-ArrayaOffice365CustomLocal.ps1') | Should Be $true
    }
}

