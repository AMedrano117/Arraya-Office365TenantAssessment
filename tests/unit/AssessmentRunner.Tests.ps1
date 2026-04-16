Describe 'Arraya.M365.AssessmentRunner' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
        $script:runnerPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psm1'
        Import-Module -Name $script:manifestPath -Force -DisableNameChecking
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
        $runnerSource | Should -Match 'function Test-AssessmentRunnerModuleMatchesManifestPath'
        $runnerSource | Should -Match 'function Resolve-AssessmentExportPathInput'
        $runnerSource | Should -Match 'Get-ArrayaAssessmentOutputRoot -FallbackPath \$script:RepoRoot'
        $runnerSource | Should -Match 'src\\scripts\\reporting'
        $runnerSource | Should -Match 'src\\scripts\\migrated\\legacy'
        $runnerSource | Should -Match 'Join-Path -Path \$Module\.ModuleBase -ChildPath \(\[System\.IO\.Path\]::GetFileName\(\$resolvedManifestPath\)\)'
        $runnerSource | Should -Not -Match '\$loadedCommonModule\.Path -ne \$resolvedCommonManifestPath'
        $runnerSource | Should -Not -Match 'function Invoke-LegacyScriptCompat'
        $runnerSource | Should -Not -Match 'function Get-LegacyScriptPath'
    }

    It 'routes the M365 family through the internal mode-based workflow helper' {
        $runnerSource = Get-Content -Raw -Path $script:runnerPath
        $runnerSource | Should -Match "ValidateSet\('Full', 'CollectOnly', 'ExportOnly'\)"
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode Full'
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode CollectOnly'
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode ExportOnly'
        $runnerSource | Should -Match 'IncludeLegacyAssessmentArtifacts'
        $runnerSource | Should -Match 'Invoke-M365ImproveForAssessmentRun'
        $runnerSource | Should -Match 'Customer Assessment Report'
        $runnerSource | Should -Match 'Customer Assessment Report Markdown'
        $runnerSource | Should -Not -Match 'Customer Remediation HTML'
        $runnerSource | Should -Not -Match 'Customer Remediation Markdown'
        $runnerSource | Should -Match '\$invokeParams\.GenerateWorkbookOverride = \[bool\]\$plan\.GenerateWorkbook'
    }

    It 'normalizes improve outputs back into the assessment run when a broad export root is supplied' {
        $runnerSource = Get-Content -Raw -Path $script:runnerPath
        $runnerSource | Should -Match 'function Resolve-AssessmentImproveOutputFolder'
        $runnerSource | Should -Match 'Resolve-AssessmentImproveOutputFolder -ManifestPath \$manifestPath -ExportPath \$ExportPath -OutputFolder \$OutputFolder'
        $runnerSource | Should -Match '\[string\]::Equals\(\$resolvedOutputFolder, \$resolvedExportPath, \[System\.StringComparison\]::OrdinalIgnoreCase\)'
    }

    It 'surfaces auth and permission-preflight controls on tenant assessment entrypoints' {
        $runnerSource = Get-Content -Raw -Path $script:runnerPath
        $runnerSource | Should -Match '\[ValidateSet\(''Interactive'', ''Certificate'', ''ClientSecret''\)\]\s*\[string\]\$AuthMode'
        $runnerSource | Should -Match '\$invokeParams\.AuthMode = \$AuthMode'
        $runnerSource | Should -Match '\[switch\]\$SkipPermissionPreflight'
        $runnerSource | Should -Match '\$invokeParams\.SkipPermissionPreflight = \$SkipPermissionPreflight'
    }

    It 'updates the assessment manifest with improve artifacts including the markdown companion' {
        $manifestPath = Join-Path $TestDrive 'assessment.manifest.json'
        [pscustomobject]@{
            SchemaVersion = 2
            GeneratedAt   = (Get-Date).ToString('o')
            OutputProfile = 'SolutionsEngineer'
            Artifacts     = @(
                [pscustomobject]@{
                    Type      = 'Workbook'
                    Path      = 'C:\Temp\tenant.xlsx'
                    Exists    = $true
                    SizeBytes = 1234
                },
                [pscustomobject]@{
                    Type      = 'Assessment Snapshot JSON'
                    Path      = 'C:\Temp\tenant-Snapshot.json'
                    Exists    = $true
                    SizeBytes = 5678
                }
            )
        } | ConvertTo-Json -Depth 10 | Set-Content -Path $manifestPath -Encoding UTF8

        $improveResult = [pscustomobject]@{
            CustomerAssessmentReportPath         = 'C:\Temp\Contoso-CustRpt.docx'
            CustomerAssessmentReportMarkdownPath = 'C:\Temp\Contoso-CustRpt.md'
            EngineerActionPackPath               = 'C:\Temp\Contoso-EngPack.md'
            JsonPath                             = 'C:\Temp\Support\Contoso-Plan.json'
            RemediationPs1Path                   = 'C:\Temp\Support\Contoso-Snips.ps1'
            CsvPath                              = $null
            MarkdownPath                         = $null
        }

        $module = Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction Stop | Select-Object -First 1
        & $module {
            param($Path, $Result)
            Update-AssessmentArtifactManifestWithImproveOutputs -ManifestPath $Path -ImproveResult $Result
        } $manifestPath $improveResult

        $updatedManifest = Get-Content -Raw -Path $manifestPath | ConvertFrom-Json -Depth 10
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Workbook' }).Count | Should -Be 1
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Assessment Snapshot JSON' }).Count | Should -Be 1
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Customer Assessment Report' }).Count | Should -Be 1
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Customer Assessment Report Markdown' }).Count | Should -Be 1
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Engineer Action Pack' }).Count | Should -Be 1
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Improvement Plan JSON' }).Count | Should -Be 1
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Remediation Snippets' }).Count | Should -Be 1
    }

    It 'resolves the run manifest from a Tenant Details workbook path' {
        $runRoot = Join-Path $TestDrive 'Contoso SE'
        $supportPath = Join-Path $runRoot 'Support'
        $null = New-Item -ItemType Directory -Path $supportPath -Force

        $workbookPath = Join-Path $runRoot 'Contoso Ltd - Tenant Details.xlsx'
        Set-Content -Path $workbookPath -Value 'placeholder' -Encoding UTF8

        $manifestPath = Join-Path $supportPath 'Contoso Ltd-Run.manifest.json'
        Set-Content -Path $manifestPath -Value '{}' -Encoding UTF8

        $module = Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction Stop | Select-Object -First 1
        $resolvedManifest = & $module {
            param($Path)
            Resolve-AssessmentLatestManifestPath -ExportPath $Path
        } $workbookPath

        $resolvedManifest | Should -Be (Resolve-Path -Path $manifestPath).Path
    }

    It 'does not throw when the inferred Support folder does not exist yet' {
        $missingRunRoot = Join-Path $TestDrive 'Missing Run Root'
        $workbookPath = Join-Path $missingRunRoot 'Contoso Ltd - Tenant Details.xlsx'

        $module = Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction Stop | Select-Object -First 1
        {
            & $module {
                param($Path)
                Resolve-AssessmentLatestManifestPath -ExportPath $Path
            } $workbookPath
        } | Should -Not -Throw
    }
}
