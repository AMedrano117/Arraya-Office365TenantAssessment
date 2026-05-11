Describe 'Solutions Engineer evidence coverage validator' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:validatorPath = Join-Path $script:repoRoot 'src\scripts\reporting\Test-SolutionsEngineerAssessmentEvidence.ps1'
        $script:exportPipelinePath = Join-Path $script:repoRoot 'src\scripts\reporting\Invoke-M365TenantAssessmentExportPipeline.ps1'
        $script:manifestWriterPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Public\Write-ArrayaAssessmentArtifactManifest.ps1'
        $script:matrixPath = Join-Path $script:repoRoot 'src\config\baseline\solutions-engineer-assessment-objectives.json'
        $script:matrix = Get-Content -Raw -Path $script:matrixPath | ConvertFrom-Json -Depth 100

        function New-TestEvidenceSnapshotObject {
            param(
                [Parameter(Mandatory = $false)]
                [string[]]$RemoveDatasets = @(),

                [Parameter(Mandatory = $false)]
                [string[]]$EmptyDatasets = @()
            )

            $datasetTable = [ordered]@{}
            $allDatasets = @(
                $script:matrix.Objectives |
                    ForEach-Object { $_.ExpectedDatasets } |
                    Sort-Object -Unique
            )

            foreach ($dataset in $allDatasets) {
                $datasetName = [string]$dataset
                $datasetTable[$datasetName] = @(
                    [pscustomobject]@{
                        Name   = $datasetName
                        Source = 'UnitTest'
                    }
                )
            }

            foreach ($dataset in @($EmptyDatasets)) {
                if ($datasetTable.Contains($dataset)) {
                    $datasetTable[$dataset] = @()
                }
            }

            foreach ($dataset in @($RemoveDatasets)) {
                if ($datasetTable.Contains($dataset)) {
                    $datasetTable.Remove($dataset)
                }
            }

            return @{
                SchemaVersion = 2
                Data          = @{
                    Tenant = $datasetTable
                }
            }
        }

        function New-TestEvidenceSnapshotFile {
            param(
                [Parameter(Mandatory = $true)]
                [string]$Path,

                [Parameter(Mandatory = $false)]
                [string[]]$RemoveDatasets = @(),

                [Parameter(Mandatory = $false)]
                [string[]]$EmptyDatasets = @()
            )

            $snapshot = New-TestEvidenceSnapshotObject -RemoveDatasets $RemoveDatasets -EmptyDatasets $EmptyDatasets
            $snapshot | ConvertTo-Json -Depth 100 | Set-Content -Path $Path -Encoding UTF8
        }

        function Export-TenantStatsJson {
            param(
                [Parameter(Mandatory = $true)]
                [hashtable]$TenantStatsHash,

                [Parameter(Mandatory = $true)]
                [string]$Path
            )

            $TenantStatsHash | ConvertTo-Json -Depth 100 | Set-Content -Path $Path -Encoding UTF8
        }
    }

    It 'reports covered, review, partial, and missing objective states from a snapshot' {
        $snapshotPath = Join-Path $TestDrive 'snapshot.json'
        New-TestEvidenceSnapshotFile `
            -Path $snapshotPath `
            -RemoveDatasets @('DeviceDetails', 'DeviceManagementSummary', 'SMTPRelaySummary') `
            -EmptyDatasets @('DlpPolicies')

        $result = & $script:validatorPath -SnapshotPath $snapshotPath -PassThru

        $result.OutputProfile | Should -Be 'SolutionsEngineer'
        $result.ReportingMode | Should -Be 'Operator'
        $result.ObjectiveCount | Should -Be @($script:matrix.Objectives).Count
        $result.MissingDatasets | Should -Contain 'DeviceDetails'
        $result.MissingDatasets | Should -Contain 'DeviceManagementSummary'
        $result.MissingDatasets | Should -Contain 'SMTPRelaySummary'
        $result.EmptyDatasets | Should -Contain 'DlpPolicies'
        $result.UnexpectedEmptyDatasets | Should -Contain 'DlpPolicies'

        $deviceObjective = $result.Objectives | Where-Object { $_.ObjectiveId -eq 'SE-DEV-001' } | Select-Object -First 1
        $deviceObjective.Status | Should -Be 'Missing'
        $deviceObjective.Confidence | Should -Be 'Missing'

        $exchangeObjective = $result.Objectives | Where-Object { $_.ObjectiveId -eq 'SE-EX-001' } | Select-Object -First 1
        $exchangeObjective.Status | Should -Be 'Partial'
        $exchangeObjective.Confidence | Should -Be 'Partial'

        $governanceObjective = $result.Objectives | Where-Object { $_.ObjectiveId -eq 'SE-GOV-001' } | Select-Object -First 1
        $governanceObjective.Status | Should -Be 'Covered'
        $governanceObjective.Confidence | Should -Be 'Review'
    }

    It 'does not downgrade confidence for datasets that are valid when empty' {
        $snapshotPath = Join-Path $TestDrive 'expected-empty-snapshot.json'
        New-TestEvidenceSnapshotFile `
            -Path $snapshotPath `
            -EmptyDatasets @('AuthenticationSSOApplications', 'InboxRulesExternalForwarding', 'ExternalSharingSiteOverrides', 'OneDriveOwnerMismatches')

        $result = & $script:validatorPath -SnapshotPath $snapshotPath -PassThru

        $appObjective = $result.Objectives | Where-Object { $_.ObjectiveId -eq 'SE-APP-001' } | Select-Object -First 1
        $exchangeObjective = $result.Objectives | Where-Object { $_.ObjectiveId -eq 'SE-EX-001' } | Select-Object -First 1
        $collaborationObjective = $result.Objectives | Where-Object { $_.ObjectiveId -eq 'SE-COL-001' } | Select-Object -First 1

        $appObjective.Confidence | Should -Be 'Strong'
        $exchangeObjective.Confidence | Should -Be 'Strong'
        $collaborationObjective.Confidence | Should -Be 'Strong'
        $exchangeObjective.ExpectedEmptyDatasets | Should -Contain 'InboxRulesExternalForwarding'
        $collaborationObjective.ExpectedEmptyDatasets | Should -Contain 'ExternalSharingSiteOverrides'
        $collaborationObjective.ExpectedEmptyDatasets | Should -Contain 'OneDriveOwnerMismatches'
        $result.UnexpectedEmptyDatasets | Should -Not -Contain 'InboxRulesExternalForwarding'
        $result.UnexpectedEmptyDatasets | Should -Not -Contain 'ExternalSharingSiteOverrides'
        $result.UnexpectedEmptyDatasets | Should -Not -Contain 'OneDriveOwnerMismatches'
    }

    It 'writes a JSON coverage artifact when requested' {
        $snapshotPath = Join-Path $TestDrive 'snapshot.json'
        $outputPath = Join-Path $TestDrive 'coverage\se-evidence-coverage.json'
        New-TestEvidenceSnapshotFile -Path $snapshotPath

        $result = & $script:validatorPath -SnapshotPath $snapshotPath -OutputPath $outputPath -PassThru

        Test-Path $outputPath | Should -BeTrue
        $written = Get-Content -Raw -Path $outputPath | ConvertFrom-Json -Depth 100
        $written.SchemaVersion | Should -Be 1
        $written.ObjectiveCount | Should -Be $result.ObjectiveCount
        $written.MissingOrPartialObjectiveCount | Should -Be 0
    }

    It 'can fail when missing or partial evidence is present' {
        $snapshotPath = Join-Path $TestDrive 'snapshot.json'
        New-TestEvidenceSnapshotFile -Path $snapshotPath -RemoveDatasets @('DeviceDetails', 'DeviceManagementSummary')

        { & $script:validatorPath -SnapshotPath $snapshotPath -FailOnMissingEvidence -PassThru } |
            Should -Throw '*missing or partial objective*'
    }

    It 'discovers evidence datasets nested under snapshot domain objects' {
        $snapshotPath = Join-Path $TestDrive 'nested-snapshot.json'
        $snapshot = [ordered]@{
            SchemaVersion = 2
            Data          = [ordered]@{
                Identity = [ordered]@{
                    Users  = @([pscustomobject]@{ DisplayName = 'User One' })
                    Admins = @([pscustomobject]@{ DisplayName = 'Admin One' })
                }
                Exchange = [ordered]@{
                    AllMailboxes = [ordered]@{
                        'user-one' = [pscustomobject]@{ PrimarySmtpAddress = 'user.one@contoso.com' }
                        'user-two' = [pscustomobject]@{ PrimarySmtpAddress = 'user.two@contoso.com' }
                    }
                }
            }
        }
        $snapshot | ConvertTo-Json -Depth 100 | Set-Content -Path $snapshotPath -Encoding UTF8

        $result = & $script:validatorPath -SnapshotPath $snapshotPath -PassThru
        $usersEvidence = $result.Objectives.DatasetEvidence | Where-Object { $_.Name -eq 'Users' } | Select-Object -First 1
        $mailboxEvidence = $result.Objectives.DatasetEvidence | Where-Object { $_.Name -eq 'AllMailboxes' } | Select-Object -First 1

        $usersEvidence.Present | Should -BeTrue
        $usersEvidence.RecordCount | Should -Be 1
        $mailboxEvidence.Present | Should -BeTrue
        $mailboxEvidence.RecordCount | Should -Be 2
    }

    It 'adds Solutions Engineer evidence coverage to the support artifacts and run manifest' {
        . $script:manifestWriterPath
        . $script:exportPipelinePath

        $exportPath = Join-Path $TestDrive 'Contoso - Tenant Details.xlsx'
        $tenantStatsHash = New-TestEvidenceSnapshotObject

        $artifacts = Invoke-M365TenantAssessmentExportPipeline `
            -TenantStatsHash $tenantStatsHash `
            -ExportTenantStatsHash $tenantStatsHash `
            -ExportDetails $exportPath `
            -SkipWorkbook $true `
            -SkipBestPracticesHtml $true `
            -SkipQuestionnaire $true `
            -SkipHtmlReport $true `
            -SkipPdfReport $true `
            -SkipJsonReport $false `
            -OutputProfileLabel 'SolutionsEngineer' `
            -ReportingMode 'Operator' `
            -CollectionOnly $false `
            -ExportOnly $false

        $artifacts.Contains('Assessment Snapshot JSON') | Should -BeTrue
        $artifacts.Contains('Solutions Engineer Evidence Coverage') | Should -BeTrue
        Test-Path -Path $artifacts['Assessment Snapshot JSON'] | Should -BeTrue
        Test-Path -Path $artifacts['Solutions Engineer Evidence Coverage'] | Should -BeTrue

        $coverage = Get-Content -Raw -Path $artifacts['Solutions Engineer Evidence Coverage'] | ConvertFrom-Json -Depth 100
        $coverage.OutputProfile | Should -Be 'SolutionsEngineer'
        $coverage.ReportingMode | Should -Be 'Operator'
        $coverage.MissingOrPartialObjectiveCount | Should -Be 0

        $manifest = Get-Content -Raw -Path $artifacts['Manifest'] | ConvertFrom-Json -Depth 100
        @($manifest.Artifacts.Type) | Should -Contain 'Solutions Engineer Evidence Coverage'
        $manifest.OperatorSummary | Should -Not -BeNullOrEmpty
        $manifest.OperatorSummary.EvidenceCoverage.Status | Should -Be 'Covered'
        $manifest.OperatorSummary.EvidenceCoverage.ObjectiveCount | Should -Be @($script:matrix.Objectives).Count
        $manifest.OperatorSummary.SupportArtifacts.Type | Should -Contain 'Solutions Engineer Evidence Coverage'
    }

    It 'does not add the Solutions Engineer evidence coverage artifact for other output profiles' {
        . $script:manifestWriterPath
        . $script:exportPipelinePath

        $exportPath = Join-Path $TestDrive 'Executive - Tenant Details.xlsx'
        $tenantStatsHash = New-TestEvidenceSnapshotObject

        $artifacts = Invoke-M365TenantAssessmentExportPipeline `
            -TenantStatsHash $tenantStatsHash `
            -ExportTenantStatsHash $tenantStatsHash `
            -ExportDetails $exportPath `
            -SkipWorkbook $true `
            -SkipBestPracticesHtml $true `
            -SkipQuestionnaire $true `
            -SkipHtmlReport $true `
            -SkipPdfReport $true `
            -SkipJsonReport $false `
            -OutputProfileLabel 'ExecutiveLevel' `
            -ReportingMode 'Minimum' `
            -CollectionOnly $false `
            -ExportOnly $false

        $artifacts.Contains('Assessment Snapshot JSON') | Should -BeTrue
        $artifacts.Contains('Solutions Engineer Evidence Coverage') | Should -BeFalse
    }
}
