BeforeAll {
    $script:RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $script:ManifestPath = Join-Path $script:RepoRoot 'dependencies.psd1'
    $script:Manifest = Import-PowerShellDataFile -Path $script:ManifestPath
}

Describe 'dependencies.psd1' {
    It 'exists at the repository root' {
        Test-Path $script:ManifestPath | Should -BeTrue
    }

    It 'declares the PowerShell version the module manifests require' {
        $script:Manifest.PowerShellVersion | Should -Be '7.0'

        foreach ($moduleManifest in @(Get-ChildItem (Join-Path $script:RepoRoot 'src/modules') -Filter '*.psd1' -Recurse)) {
            $data = Import-PowerShellDataFile -Path $moduleManifest.FullName
            $data.PowerShellVersion | Should -Be $script:Manifest.PowerShellVersion -Because "$($moduleManifest.Name) must agree with the dependency manifest"
        }
    }

    It 'pins every module to a bounded version range' {
        foreach ($module in @($script:Manifest.Modules)) {
            $module.MinimumVersion | Should -Not -BeNullOrEmpty -Because "$($module.Name) needs a minimum version"
            $module.MaximumVersion | Should -Not -BeNullOrEmpty -Because "$($module.Name) needs a maximum version"
            ([version]$module.MinimumVersion) | Should -BeLessOrEqual ([version]$module.MaximumVersion) -Because "$($module.Name) range must be valid"
        }
    }

    It 'gives every module at least one valid scope and a reason' {
        $validScopes = @('Runtime', 'Onboarding', 'DevTools')
        foreach ($module in @($script:Manifest.Modules)) {
            @($module.Scope).Count | Should -BeGreaterThan 0 -Because "$($module.Name) needs a scope"
            foreach ($scope in @($module.Scope)) {
                $scope | Should -BeIn $validScopes
            }
            $module.Reason | Should -Not -BeNullOrEmpty -Because "$($module.Name) should document why it is needed"
        }
    }

    It 'lists no duplicate module names' {
        $names = @($script:Manifest.Modules | ForEach-Object { $_.Name })
        ($names | Sort-Object -Unique).Count | Should -Be $names.Count
    }

    It 'covers the Graph submodules the onboarding script actually calls' {
        # The previous installer shipped only Microsoft.Graph.Authentication, so onboarding
        # failed on a clean machine at the first application cmdlet.
        $onboardingSource = Get-Content -Path (Join-Path $script:RepoRoot 'onboarding/New-AssessmentAppRegistration.ps1') -Raw
        $onboardingModules = @(
            $script:Manifest.Modules |
                Where-Object { @($_.Scope) -contains 'Onboarding' } |
                ForEach-Object { $_.Name }
        )

        if ($onboardingSource -match '-MgApplication|-MgServicePrincipal') {
            $onboardingModules | Should -Contain 'Microsoft.Graph.Applications'
        }
        if ($onboardingSource -match '-MgRoleManagementDirectory') {
            $onboardingModules | Should -Contain 'Microsoft.Graph.Identity.Governance'
        }
        $onboardingModules | Should -Contain 'Microsoft.Graph.Authentication'
    }

    It 'keeps the Graph SDK submodules on a single major line' {
        $graphModules = @($script:Manifest.Modules | Where-Object { $_.Name -like 'Microsoft.Graph.*' })
        $graphModules.Count | Should -BeGreaterThan 1

        $minimums = @($graphModules | ForEach-Object { [version]$_.MinimumVersion })
        @($minimums | ForEach-Object { $_.Major } | Sort-Object -Unique).Count | Should -Be 1 -Because 'mixing Graph SDK major versions in one session breaks assembly loading'
    }
}

Describe 'tools/install-dependencies.ps1' {
    BeforeAll {
        $script:InstallerPath = Join-Path $script:RepoRoot 'tools/install-dependencies.ps1'
    }

    It 'exists and parses' {
        Test-Path $script:InstallerPath | Should -BeTrue

        $parseErrors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($script:InstallerPath, [ref]$null, [ref]$parseErrors)
        @($parseErrors).Count | Should -Be 0
    }

    It 'reports satisfied requirements without installing when validating' {
        $output = & pwsh -NoProfile -File $script:InstallerPath -Scope DevTools -Validate 2>&1
        $LASTEXITCODE | Should -Be 0
        ($output -join "`n") | Should -Match 'All manifest dependencies are satisfied'
    }

    It 'exits nonzero when a requirement cannot be met' {
        $fakeManifest = Join-Path $TestDrive 'fake-dependencies.psd1'
        @'
@{
    PowerShellVersion = '7.0'
    Modules = @(
        @{ Name = 'TotallyFakeModule'; MinimumVersion = '1.0.0'; MaximumVersion = '1.99.99'; Scope = @('DevTools'); Reason = 'test' }
    )
}
'@ | Set-Content -Path $fakeManifest -Encoding UTF8

        $output = & pwsh -NoProfile -File $script:InstallerPath -Scope DevTools -Validate -ManifestPath $fakeManifest 2>&1
        $LASTEXITCODE | Should -Be 1
        ($output -join "`n") | Should -Match 'not satisfied'
    }
}

Describe 'Deprecated installer shims' {
    It 'forwards to the manifest-driven installer instead of installing latest' {
        foreach ($shim in @('tools/install-microsoft-modules.ps1', 'tools/install-required-modules.ps1')) {
            $source = Get-Content -Path (Join-Path $script:RepoRoot $shim) -Raw
            $source | Should -Match 'install-dependencies\.ps1' -Because "$shim should delegate"
            $source | Should -Not -Match "Install-Module\s" -Because "$shim should no longer install modules directly"
        }
    }
}
