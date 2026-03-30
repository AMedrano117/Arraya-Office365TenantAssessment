Describe 'Arraya.M365.AssessmentRunner' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
        $script:runnerPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psm1'
    }

    It 'has a manifest and top-level wrapper scripts' {
        Test-Path $script:manifestPath | Should -BeTrue
        @(
            'src\scripts\operations\Start-M365TenantAssessment.ps1'
            'src\scripts\assessments\tenant-wide\Invoke-M365FullTenantAssessment.ps1'
            'src\scripts\assessments\tenant-wide\Invoke-M365TenantDataCollection.ps1'
            'src\scripts\assessments\tenant-wide\Invoke-M365GraphActivityReport.ps1'
            'src\scripts\reporting\Invoke-M365TenantAssessmentExport.ps1'
            'src\scripts\reporting\Invoke-M365TenantImprovementPlan.ps1'
            'src\scripts\reporting\Invoke-M365TenantAssessmentComparison.ps1'
            'src\scripts\reporting\New-M365TenantImprovementPlan.ps1'
            'src\scripts\reporting\Compare-M365TenantAssessmentSnapshots.ps1'
        ) | ForEach-Object {
            Test-Path (Join-Path $script:repoRoot $_) | Should -BeTrue
        }
    }

    It 'uses assessment-script helpers and modern script roots' {
        $runnerSource = Get-Content -Raw -Path $script:runnerPath
        $runnerSource | Should -Match 'function Invoke-AssessmentScript'
        $runnerSource | Should -Match 'function Invoke-M365TenantWorkflow'
        $runnerSource | Should -Match 'function Resolve-AssessmentScriptPath'
        $runnerSource | Should -Match 'function Resolve-AssessmentExportPathInput'
        $runnerSource | Should -Match 'Get-ArrayaAssessmentOutputRoot -FallbackPath \$script:RepoRoot'
        $runnerSource | Should -Match 'src\\scripts\\reporting'
        $runnerSource | Should -Match 'src\\scripts\\migrated\\legacy'
        $runnerSource | Should -Not -Match 'function Invoke-LegacyScriptCompat'
        $runnerSource | Should -Not -Match 'function Get-LegacyScriptPath'
    }

    It 'routes the M365 family through the internal mode-based workflow helper' {
        $runnerSource = Get-Content -Raw -Path $script:runnerPath
        $runnerSource | Should -Match "ValidateSet\('Full', 'CollectOnly', 'ExportOnly'\)"
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode Full'
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode CollectOnly'
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode ExportOnly'
    }

    It 'surfaces AuthMode on tenant assessment entrypoints' {
        $runnerSource = Get-Content -Raw -Path $script:runnerPath
        $runnerSource | Should -Match '\[ValidateSet\(''Interactive'', ''Certificate'', ''ClientSecret''\)\]\s*\[string\]\$AuthMode'
        $runnerSource | Should -Match '\$invokeParams\.AuthMode = \$AuthMode'
    }
}
