Describe 'Solutions Engineer assessment objective matrix' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:matrixPath = Join-Path $script:repoRoot 'src\config\baseline\solutions-engineer-assessment-objectives.json'
        $script:profilePolicyPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Public\Get-ArrayaAssessmentOutputProfilePolicy.ps1'
        $script:collectorPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        $script:improvePath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\New-M365TenantImprovementPlan.ps1'
        $script:exportExcelPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Public\Export-HashTableToExcel.ps1'
        $script:snapshotPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Public\Convert-ArrayaLegacyTenantStatsToSnapshot.ps1'
        $script:customerReportPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Private\CustomerAssessmentDocx.ps1'
        $script:htmlHelperPath = Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1'
        $script:graphGroupsPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Graph\Public\Get-EntraIDGroups.ps1'
        $script:exchangeGroupsPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Public\Get-ExchangeGroupDetails.ps1'

        $script:matrix = Get-Content -Raw -Path $script:matrixPath | ConvertFrom-Json -Depth 20
        $script:profilePolicySource = Get-Content -Raw -Path $script:profilePolicyPath
        $script:improveSource = Get-Content -Raw -Path $script:improvePath
        $script:sourceIndex = @(
            $script:collectorPath,
            $script:improvePath,
            $script:exportExcelPath,
            $script:snapshotPath,
            $script:customerReportPath,
            $script:htmlHelperPath,
            $script:graphGroupsPath,
            $script:exchangeGroupsPath
        ) | ForEach-Object { Get-Content -Raw -Path $_ } | Out-String

        $script:requiredAreas = @(
            'Tenant baseline and collection quality',
            'Identity and privileged access',
            'MFA and authentication methods',
            'Conditional Access quality',
            'Enterprise application governance',
            'Exchange and mail hygiene',
            'SharePoint, OneDrive, and external sharing',
            'Teams and Microsoft 365 Groups governance',
            'Licensing optimization',
            'Endpoint and device posture',
            'Purview, retention, DLP, and governance summaries'
        )

        $script:allowedReportUses = @(
            'CustomerAssessmentReport',
            'RemediationRoadmap',
            'EngineerActionPack',
            'Workbook',
            'JsonSnapshot',
            'ImprovementPlan'
        )
    }

    It 'keeps the SolutionsEngineer profile mapped to Operator without changing report output switches' {
        . $script:profilePolicyPath

        $policy = Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile SolutionsEngineer

        $policy.ReportingMode | Should -Be 'Operator'
        $policy.GenerateWorkbook | Should -BeTrue
        $policy.GenerateTechnicalHtml | Should -BeTrue
        $policy.GenerateJson | Should -BeTrue
        $policy.GenerateBestPracticesHtml | Should -BeFalse
        $policy.GenerateQuestionnaire | Should -BeFalse
        $policy.GeneratePdf | Should -BeFalse
        $script:profilePolicySource | Should -Match "'SolutionsEngineer'"
        $script:profilePolicySource | Should -Match "ReportingMode\s+=\s+'Operator'"
    }

    It 'has the required schema fields and objective families' {
        $script:matrix.SchemaVersion | Should -Be 1
        $script:matrix.OutputProfile | Should -Be 'SolutionsEngineer'
        $script:matrix.ReportingMode | Should -Be 'Operator'
        $script:matrix.GeneratedReportImpact | Should -Match 'None'
        @($script:matrix.Objectives).Count | Should -BeGreaterOrEqual $script:requiredAreas.Count

        foreach ($area in $script:requiredAreas) {
            @($script:matrix.ObjectiveFamilies) | Should -Contain $area
            @($script:matrix.Objectives.Area) | Should -Contain $area
        }
    }

    It 'defines complete objective rows with concrete evidence and bounded report use' {
        $objectiveIds = @{}

        foreach ($objective in @($script:matrix.Objectives)) {
            [string]$objective.ObjectiveId | Should -Match '^SE-[A-Z]+-\d{3}$'
            $objectiveIds.ContainsKey($objective.ObjectiveId) | Should -BeFalse
            $objectiveIds[$objective.ObjectiveId] = $true

            [string]$objective.Area | Should -Not -BeNullOrEmpty
            [string]$objective.Goal | Should -Not -BeNullOrEmpty
            [string]$objective.ConfidenceNotes | Should -Not -BeNullOrEmpty
            @($objective.EvidenceSources).Count | Should -BeGreaterThan 0
            @($objective.ExpectedDatasets).Count | Should -BeGreaterThan 0
            @($objective.ChecksFor).Count | Should -BeGreaterThan 0
            @($objective.ReportUse).Count | Should -BeGreaterThan 0

            $expectedEmptyDatasets = @()
            if ($objective.PSObject.Properties['ExpectedEmptyDatasets']) {
                $expectedEmptyDatasets = @($objective.ExpectedEmptyDatasets)
            }

            foreach ($expectedEmptyDataset in $expectedEmptyDatasets) {
                @($objective.ExpectedDatasets) | Should -Contain $expectedEmptyDataset
            }

            foreach ($reportUse in @($objective.ReportUse)) {
                $script:allowedReportUses | Should -Contain $reportUse
            }
        }
    }

    It 'references only datasets known to the collector, exporter, or reporting pipeline' {
        $datasets = @(
            $script:matrix.Objectives |
                ForEach-Object { $_.ExpectedDatasets } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $datasets.Count | Should -BeGreaterThan 0

        foreach ($dataset in $datasets) {
            $script:sourceIndex | Should -Match ([regex]::Escape([string]$dataset))
        }
    }

    It 'references only implemented finding rules or documentation-only rule IDs' {
        $ruleIds = @(
            $script:matrix.Objectives |
                ForEach-Object { $_.RelatedFindings } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $ruleIds.Count | Should -BeGreaterThan 0

        foreach ($ruleId in $ruleIds) {
            if ($ruleId -match '^DOC-') {
                $ruleId | Should -Match '^DOC-[A-Z0-9-]+$'
            }
            else {
                $script:improveSource | Should -Match ("RuleId\s+'{0}'" -f [regex]::Escape([string]$ruleId))
            }
        }
    }

    It 'represents the major Solutions Engineer finding rule prefixes' {
        $prefixes = @(
            $script:matrix.Objectives |
                ForEach-Object { $_.RelatedFindings } |
                Where-Object { $_ -match '^[A-Z]+-' } |
                ForEach-Object { ([string]$_).Split('-')[0] } |
                Sort-Object -Unique
        )

        foreach ($prefix in @('SEC', 'CA', 'MFA', 'ADMIN', 'ID', 'EX', 'COL', 'TM', 'LIC', 'DEV')) {
            $prefixes | Should -Contain $prefix
        }
    }

    It 'documents coverage gaps for live validation and heuristic confidence' {
        @($script:matrix.CoverageGaps).Count | Should -BeGreaterThan 0

        foreach ($gap in @($script:matrix.CoverageGaps)) {
            [string]$gap.GapId | Should -Match '^SE-GAP-\d{3}$'
            [string]$gap.Area | Should -Not -BeNullOrEmpty
            [string]$gap.CurrentLimitation | Should -Not -BeNullOrEmpty
            [string]$gap.Impact | Should -Not -BeNullOrEmpty
            [string]$gap.NextValidationLayer | Should -Not -BeNullOrEmpty
        }

        @($script:matrix.CoverageGaps.Area) | Should -Contain 'Live integration proof'
        @($script:matrix.CoverageGaps.Area) | Should -Contain 'Heuristic confidence'
    }
}
