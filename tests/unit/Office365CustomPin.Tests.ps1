Describe 'Office365Custom version pin' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $loaderPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Public\Import-ArrayaOffice365CustomLocal.ps1'
        . $loaderPath
    }

    BeforeEach {
        Remove-Module -Name 'Office365Custom' -Force -ErrorAction SilentlyContinue
    }

    It 'loads supported version 1.2.1 instead of a higher unapproved folder version' {
        $fakeRepoRoot = Join-Path $TestDrive 'repo'
        foreach ($version in @('1.2.1', '9.9.9')) {
            $moduleRoot = Join-Path $fakeRepoRoot "src\vendor\Office365Custom\$version"
            $null = New-Item -ItemType Directory -Path $moduleRoot -Force
            $modulePath = Join-Path $moduleRoot 'Office365Custom.psm1'
            $manifestPath = Join-Path $moduleRoot 'Office365Custom.psd1'
            $moduleSource = "function Get-PinnedOfficeCommand { '$version' }`nExport-ModuleMember -Function 'Get-PinnedOfficeCommand'"
            Set-Content -LiteralPath $modulePath -Value $moduleSource -Encoding utf8
            New-ModuleManifest `
                -Path $manifestPath `
                -RootModule 'Office365Custom.psm1' `
                -ModuleVersion $version `
                -FunctionsToExport @('Get-PinnedOfficeCommand')
        }

        $module = Import-ArrayaOffice365CustomLocal `
            -RepoRoot $fakeRepoRoot `
            -RequiredCommands @('Office365Custom\Get-PinnedOfficeCommand')

        $module.Version.ToString() | Should -Be '1.2.1'
        Office365Custom\Get-PinnedOfficeCommand | Should -Be '1.2.1'
        $module.ModuleBase | Should -Be (Join-Path $fakeRepoRoot 'src\vendor\Office365Custom\1.2.1')
    }

    AfterAll {
        Remove-Module -Name 'Office365Custom' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path Function:\Import-ArrayaOffice365CustomLocal -ErrorAction SilentlyContinue
    }
}
