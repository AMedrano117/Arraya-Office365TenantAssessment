Describe 'Arraya.M365.AssessmentRunner' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
        $script:runnerPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psm1'
        $script:launcherPath = Join-Path $script:repoRoot 'src\scripts\operations\Start-M365TenantAssessment.ps1'
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
        $runnerSource | Should -Match "ValidateSet\('Full', 'CollectOnly', 'ExportOnly', 'PreflightOnly'\)"
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode Full'
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode PreflightOnly'
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode CollectOnly'
        $runnerSource | Should -Match 'Invoke-M365TenantWorkflow -Mode ExportOnly'
        $runnerSource | Should -Match 'function Invoke-M365TenantConnectionPreflight'
        $runnerSource | Should -Match '\[switch\]\$UseExistingConnections'
        $runnerSource | Should -Match '\$invokeParams\.UseExistingConnections = \$UseExistingConnections'
        $runnerSource | Should -Match 'IncludeLegacyAssessmentArtifacts'
        $runnerSource | Should -Match 'Invoke-M365ImproveForAssessmentRun'
        $runnerSource | Should -Match 'Customer Assessment Report'
        $runnerSource | Should -Not -Match 'Customer Assessment Report Markdown'
        $runnerSource | Should -Match 'Roadmap Remediation Plan'
        $runnerSource | Should -Not -Match 'Customer Remediation HTML'
        $runnerSource | Should -Not -Match 'Customer Remediation Markdown'
        $runnerSource | Should -Match '\$invokeParams\.GenerateWorkbookOverride = \[bool\]\$plan\.GenerateWorkbook'
        $runnerSource | Should -Match 'WorkbookExportPolicy'
        $runnerSource | Should -Match 'TechnicalHtmlPolicy'
        $runnerSource | Should -Match 'GenerateMigrationPack'
        $runnerSource | Should -Match 'CollectionScopePolicy'
        $runnerSource | Should -Match '\$invokeParams\.WorkbookExportPolicyOverride = \[string\]\$plan\.WorkbookExportPolicy'
        $runnerSource | Should -Match '\$invokeParams\.TechnicalHtmlPolicyOverride = \[string\]\$plan\.TechnicalHtmlPolicy'
        $runnerSource | Should -Match '\$invokeParams\.GenerateMigrationPackOverride = \[bool\]\$plan\.GenerateMigrationPack'
        $runnerSource | Should -Match '\$invokeParams\.CollectionScopePolicyOverride = \[string\]\$plan\.CollectionScopePolicy'
        $runnerSource | Should -Match '\$shouldEnablePolicyTechnicalHtml = \(\[string\]\$plan\.TechnicalHtmlPolicy -eq ''TenantToTenantCutover''\)'
        $runnerSource | Should -Match 'if \(\$IncludeLegacyAssessmentArtifacts -or \$shouldEnablePolicyTechnicalHtml\)'
        $runnerSource | Should -Match 'if \(-not \$UseExistingConnections\) \{[\s\S]*\$invokeParams\.SkipAuth = \$SkipAuth[\s\S]*\$invokeParams\.SkipPermissionPreflight = \$SkipPermissionPreflight[\s\S]*\$invokeParams\.AuthMode = \$AuthMode'
        $runnerSource | Should -Match 'if \(-not \$UseExistingConnections\) \{[\s\S]*\$invokeParams\.CertificateThumbprint = \$CertificateThumbprint[\s\S]*\$invokeParams\.ClientSecretSecure = \$ClientSecretSecure'
        $runnerSource | Should -Match '\$invokeParams\.ExportOnly = \$true'
        $runnerSource | Should -Match '\$invokeParams\.TenantStatsJsonPath = \$AssessmentJsonPath'
        $runnerSource | Should -Match 'function Test-AssessmentPlanSkipsImproveByDefault'
        $runnerSource | Should -Match 'SkipImproveByDefault'
        $runnerSource | Should -Match 'Tenant-to-tenant migration profile defaults to workbook, cutover pack, technical HTML, and JSON snapshot output\. Skipping Improve/customer-style follow-up artifacts unless requested separately\.'
    }

    It 'validates the manifest contains a usable snapshot before launching Improve from a completed run' {
        $runnerSource = Get-Content -Raw -Path $script:runnerPath
        $runnerSource | Should -Match 'Import-ArrayaTenantSnapshotContext -Path \$manifestPath -Purpose ImprovementPlan'
        $runnerSource | Should -Match 'The completed assessment run did not produce a usable JSON snapshot, so Improve cannot continue from the manifest'
        $runnerSource | Should -Match 'This usually means snapshot export failed earlier in the run even if the workbook was created'
    }

    It 'uses the same usable-snapshot manifest guard in the launcher handoff' {
        $launcherSource = Get-Content -Raw -Path $script:launcherPath
        $launcherSource | Should -Match 'function Test-LauncherManifestHasSnapshotArtifact'
        $launcherSource | Should -Match 'Import-ArrayaTenantSnapshotContext -Path \$manifestPath -Purpose ImprovementPlan'
        $launcherSource | Should -Match 'The completed assessment run did not produce a usable JSON snapshot, so Improve cannot continue from the manifest'
    }

    It 'normalizes improve outputs back into the assessment run when a broad export root is supplied' {
        $runnerSource = Get-Content -Raw -Path $script:runnerPath
        $runnerSource | Should -Match 'function Resolve-AssessmentImproveOutputFolder'
        $runnerSource | Should -Match 'Resolve-AssessmentImproveOutputFolder -ManifestPath \$manifestPath -ExportPath \$ExportPath -OutputFolder \$OutputFolder'
        $runnerSource | Should -Match '\[string\]::Equals\(\$resolvedOutputFolder, \$resolvedExportPath, \[System\.StringComparison\]::OrdinalIgnoreCase\)'
    }

    It 'uses the primary profile for workbook and HTML policy, but falls back to default collection scope for merged selections' {
        $module = Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction Stop | Select-Object -First 1

        $singlePlan = & $module {
            Resolve-M365OutputProfileExecutionPlan -OutputProfile 'TenantToTenantMigration'
        }
        $singlePlan.WorkbookExportPolicy | Should -Be 'TenantToTenantCutover'
        $singlePlan.TechnicalHtmlPolicy | Should -Be 'TenantToTenantCutover'
        $singlePlan.GenerateMigrationPack | Should -BeTrue
        $singlePlan.CollectionScopePolicy | Should -Be 'TenantToTenantCutover'
        $singlePlan.SkipImproveByDefault | Should -BeTrue

        $mergedPlan = & $module {
            Resolve-M365OutputProfileExecutionPlan -OutputProfile @('TenantToTenantMigration', 'SolutionsEngineer')
        }
        $mergedPlan.PrimaryProfile | Should -Be 'TenantToTenantMigration'
        $mergedPlan.WorkbookExportPolicy | Should -Be 'TenantToTenantCutover'
        $mergedPlan.TechnicalHtmlPolicy | Should -Be 'TenantToTenantCutover'
        $mergedPlan.GenerateMigrationPack | Should -BeTrue
        $mergedPlan.CollectionScopePolicy | Should -Be 'Default'
        $mergedPlan.SkipImproveByDefault | Should -BeFalse
    }

    It 'surfaces auth and permission-preflight controls on tenant assessment entrypoints' {
        $runnerSource = Get-Content -Raw -Path $script:runnerPath
        $runnerSource | Should -Match '\[ValidateSet\(''Interactive'', ''Certificate'', ''ClientSecret''\)\]\s*\[string\]\$AuthMode'
        $runnerSource | Should -Match '\$invokeParams\.AuthMode = \$AuthMode'
        $runnerSource | Should -Match '\$invokeParams\.PreflightOnly = \$true'
        $runnerSource | Should -Match '\[switch\]\$SkipPermissionPreflight'
        $runnerSource | Should -Match '\$invokeParams\.SkipPermissionPreflight = \$SkipPermissionPreflight'
        $runnerSource | Should -Match '\[pscredential\]\$ClientSecretCredential'
        $runnerSource | Should -Match '\[securestring\]\$ClientSecretSecure'
        $runnerSource | Should -Match '\$invokeParams\.ClientSecretCredential = \$ClientSecretCredential'
        $runnerSource | Should -Match '\$invokeParams\.ClientSecretSecure = \$ClientSecretSecure'

        $launcherSource = Get-Content -Raw -Path $script:launcherPath
        $launcherSource | Should -Match '\[pscredential\]\$ClientSecretCredential'
        $launcherSource | Should -Match '\[securestring\]\$ClientSecretSecure'
        $launcherSource | Should -Match '\$invokeParams\.ClientSecretCredential = \$ClientSecretCredential'
        $launcherSource | Should -Match '\$invokeParams\.ClientSecretSecure = \$ClientSecretSecure'
    }

    It 'treats client secret auth as a compatibility path and gates the unsupported Entra reference script' {
        $launcherSource = Get-Content -Raw -Path $script:launcherPath
        $launcherSource | Should -Match 'Client secret auth is a compatibility path in this workflow'

        $legacyEntraReportPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Invoke-EntraAppReport.ps1'
        Test-Path $legacyEntraReportPath | Should -BeTrue

        $legacyEntraReportSource = Get-Content -Raw -Path $legacyEntraReportPath
        $legacyEntraReportSource | Should -Match '\[switch\]\$AllowUnsupportedLegacyExecution'
        $legacyEntraReportSource | Should -Match 'unsupported legacy reference script'
    }

    It 'updates the assessment manifest with improve artifacts without a customer markdown companion' {
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
            CustomerAssessmentReportPath         = 'C:\Temp\Deliverables\Contoso-CustomerReport.docx'
            RoadmapRemediationPlanPath          = 'C:\Temp\Deliverables\Contoso-Roadmap.docx'
            EngineerActionPackPath               = 'C:\Temp\Deliverables\Contoso-EngPack.md'
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
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Customer Assessment Report Markdown' }).Count | Should -Be 0
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Roadmap Remediation Plan' }).Count | Should -Be 1
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Engineer Action Pack' }).Count | Should -Be 1
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Improvement Plan JSON' }).Count | Should -Be 1
        @($updatedManifest.Artifacts | Where-Object { $_.Type -eq 'Remediation Snippets' }).Count | Should -Be 1
        $updatedManifest.OperatorSummary | Should -Not -BeNullOrEmpty
        @($updatedManifest.OperatorSummary.PrimaryDeliverables.Type) | Should -Contain 'Customer Assessment Report'
        @($updatedManifest.OperatorSummary.PrimaryDeliverables.Type) | Should -Contain 'Roadmap Remediation Plan'
        @($updatedManifest.OperatorSummary.PrimaryDeliverables.Type) | Should -Contain 'Engineer Action Pack'
        @($updatedManifest.OperatorSummary.SupportArtifacts.Type) | Should -Contain 'Assessment Snapshot JSON'
        @($updatedManifest.OperatorSummary.SupportArtifacts.Type) | Should -Contain 'Improvement Plan JSON'
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

    It 'resolves the run manifest from a Tenant Details workbook path in Deliverables' {
        $runRoot = Join-Path $TestDrive 'Contoso SE Split'
        $deliverablesPath = Join-Path $runRoot 'Deliverables'
        $supportPath = Join-Path $runRoot 'Support'
        $null = New-Item -ItemType Directory -Path $deliverablesPath -Force
        $null = New-Item -ItemType Directory -Path $supportPath -Force

        $workbookPath = Join-Path $deliverablesPath 'Contoso Ltd - Tenant Details.xlsx'
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

    It 'prefers a manifest with an existing assessment snapshot when multiple run manifests are present' {
        $runRoot = Join-Path $TestDrive 'Contoso SE With Snapshot'
        $supportPath = Join-Path $runRoot 'Support'
        $null = New-Item -ItemType Directory -Path $supportPath -Force

        $invalidManifestPath = Join-Path $supportPath 'Run.manifest.json'
        [pscustomobject]@{
            SchemaVersion = 2
            Artifacts     = @(
                [pscustomobject]@{
                    Type = 'Workbook'
                    Path = (Join-Path $runRoot 'Contoso.xlsx')
                }
            )
        } | ConvertTo-Json -Depth 10 | Set-Content -Path $invalidManifestPath -Encoding UTF8

        $snapshotPath = Join-Path $supportPath 'Contoso-Snap.json'
        Set-Content -Path $snapshotPath -Value '{}' -Encoding UTF8

        $validManifestPath = Join-Path $supportPath 'Contoso-Run.manifest.json'
        [pscustomobject]@{
            SchemaVersion = 2
            Artifacts     = @(
                [pscustomobject]@{
                    Type = 'Assessment Snapshot JSON'
                    Path = $snapshotPath
                }
            )
        } | ConvertTo-Json -Depth 10 | Set-Content -Path $validManifestPath -Encoding UTF8

        $module = Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction Stop | Select-Object -First 1
        $resolvedManifest = & $module {
            param($Path)
            Resolve-AssessmentLatestManifestPath -ExportPath $Path
        } $runRoot

        $resolvedManifest | Should -Be (Resolve-Path -Path $validManifestPath).Path
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
