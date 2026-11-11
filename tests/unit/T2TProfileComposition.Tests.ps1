# Scenario tests for profile scope, multi-profile composition, collector evidence states, hold
# evidence versus migration treatment, the narrow Presales blocker rule, and the
# MigrationCriteriaRegister projection. Fixtures are fictional and use reserved .example
# domains. Exchange cmdlets are replaced by local stubs; nothing contacts a tenant.
Describe 'T2T profile composition and evidence states' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        Import-Module -Name (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -WarningAction SilentlyContinue -ErrorAction Stop
        Import-Module -Name (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Arraya.M365.Exchange.psd1') -Force -WarningAction SilentlyContinue -ErrorAction Stop

        function global:Write-Log { param() }
        function global:Write-ProgressHelper { param() }
        function Write-AssessmentCollectorBanner { param() }
        function Write-AssessmentCollectorCompletionBanner { param() }

        $collectorPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($collectorPath, [ref]$tokens, [ref]$parseErrors)
        $wantedFunctions = @(
            'Get-AssessmentNormalizedSiteUrl', 'Get-AssessmentExportTableArray', 'Get-AssessmentExportPropertyValue',
            'Set-AssessmentExportProperty', 'Convert-AssessmentExportSizeToGb', 'Convert-AssessmentExportByteCountToGb',
            'Convert-DataSizeToBytes', 'Convert-AssessmentExportListToArray', 'Convert-AssessmentExportListToText',
            'Test-AssessmentMailboxHasArchiveEvidence', 'Get-AssessmentTeamEnrichmentEvidence',
            'Update-AssessmentTeamEnrichmentEvidenceRows', 'Ensure-MigrationGroupMailboxInventory',
            'Get-AssessmentMailboxPlanningRows', 'Update-MigrationGroupWorkloadReconciliation',
            'Get-MigrationShareGateScopeData', 'Update-MigrationPlanningRegisters', 'Get-AssessmentMailboxStatsLookup',
            'Add-TenantToTenantMailboxDerivedFields', 'Get-AssessmentBitTitanBlockSizeGB', 'Get-AssessmentMailboxLicenseType',
            'Get-AssessmentCustomDomainApplicationUris', 'Get-AssessmentExportDomainFromAddress',
            'Convert-AssessmentExportDisplayText', 'New-AssessmentExportLookup', 'New-TenantToTenantConfigurationRow',
            'Resolve-AssessmentExportRoutingSummary', 'Update-MigrationComplexityFlags', 'Update-MigrationScopeExportData',
            'Update-TenantToTenantMigrationExportData'
        )
        $definitions = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) |
            Where-Object { $wantedFunctions -contains $_.Name }
        foreach ($definition in $definitions) { . ([scriptblock]::Create($definition.Extent.Text)) }
        $script:AssessmentBitTitanLicenseModel = Get-ArrayaBitTitanLicenseModel

        $script:allCollectorStates = @('CompleteWithRecords', 'CompleteEmpty', 'Partial', 'Failed', 'Unavailable', 'NotRequested', 'Unsupported', 'NotApplicable', 'Unknown', 'SkippedByProfile')

        # ---- Fixture builders --------------------------------------------------------------
        function script:New-Step {
            param([string]$Name, [string]$Status = 'Completed', [string]$Message)
            [pscustomobject]@{ Name = $Name; Section = 'Test'; Workload = 'Test'; Status = $Status; Required = $false; Message = $Message; DurationSeconds = 0; ProducedKeys = @() }
        }
        function script:New-Mailbox {
            param([string]$Alias, [hashtable]$Holds = @{}, [switch]$OmitHoldFields, [switch]$WithDelegates)
            $row = [ordered]@{
                DisplayName = $Alias; UserPrincipalName = "$Alias@contoso.example"; PrimarySmtpAddress = "$Alias@contoso.example"
                RecipientTypeDetails = 'UserMailbox'; ExchangeGuid = ([guid]::NewGuid()).ToString(); MailboxSizeGB = 3; ArchiveStatus = 'None'
            }
            if (-not $OmitHoldFields) {
                $row.LitigationHoldEnabled = $false; $row.DelayHoldApplied = $false; $row.RetentionHoldEnabled = $false; $row.InPlaceHolds = ''
            }
            foreach ($key in $Holds.Keys) { $row[$key] = $Holds[$key] }
            if ($WithDelegates) {
                $row.FullAccessDelegateState = 'Collected'; $row.SendAsDelegateState = 'Collected'; $row.CalendarDelegateState = 'Collected'
            }
            [pscustomobject]$row
        }
        function script:New-SyncSummary {
            param($Flag)
            @{ Summary = [pscustomobject]@{ OnPremisesSyncEnabled = $Flag }; SyncServices = @(); RecentErrors = $null; ErrorCount = 0 }
        }
        function script:New-Hybrid {
            param($IsHybrid, [bool]$Complete = $true, [string]$Status = 'None')
            @{ ExchangeHybrid = [pscustomobject]@{ IsHybridConfigured = $IsHybrid; HybridStatus = $Status; EvidenceComplete = $Complete; EvidenceCount = 0; MigrationEndpointCount = 0; UnansweredProbes = $(if ($Complete) { '' } else { 'MigrationEndpoint (Error)' }) } }
        }
        # A tenant where every hold source and every Presales dataset was collected.
        function script:New-CompleteStore {
            param([object[]]$Mailboxes = @((New-Mailbox -Alias 'avery')))
            $store = @{
                TenantInfo             = [pscustomobject]@{ DisplayName = 'Contoso Synthetic'; TenantId = '00000000-0000-0000-0000-000000000001' }
                AllMailboxes           = @($Mailboxes)
                AdConnectConfiguration = New-SyncSummary -Flag $null
                HybridConfiguration    = New-Hybrid -IsHybrid $false
                RetentionPolicies      = @{}
                Domains                = @{ 'contoso.example' = [pscustomobject]@{ Domain = 'contoso.example'; Verified = $true; Office365MailExchanger = $true } }
                SharePoint             = @([pscustomobject]@{ Title = 'Site'; Url = 'https://contoso.sharepoint.example/sites/a'; StorageUsedGB = 1 })
                OneDrive               = @([pscustomobject]@{ Title = 'Avery'; Url = 'https://contoso-my.sharepoint.example/personal/avery'; StorageUsedGB = 1 })
                AllTeams               = @{}
                EnterpriseApplications = @{}
                LicenseSKUs            = @([pscustomobject]@{ SkuPartNumber = 'SYNTHETIC_SKU' })
                DeviceDetails          = @{}
                CollectorPlanResults   = @(
                    (New-Step 'Exchange mailboxes'), (New-Step 'AD Connect sync details'), (New-Step 'Exchange hybrid configuration'),
                    (New-Step 'Purview retention/DLP policies'), (New-Step 'Domains'), (New-Step 'SharePoint/OneDrive sites (API)'),
                    (New-Step 'Teams inventory (MGGraph)'), (New-Step 'Authentication/SSO configuration'), (New-Step 'License SKUs'),
                    (New-Step 'Devices')
                )
            }
            $store
        }
        function script:Get-Row {
            param([object[]]$Rows, [string]$CriterionId)
            @($Rows | Where-Object { $_.CriterionId -eq $CriterionId })[0]
        }
        function script:New-SyntheticCatalog {
            param([object[]]$Criteria)
            [pscustomobject]@{ Criteria = @($Criteria); HoldMethodRules = @() }
        }
        function script:New-SyntheticCriterion {
            param([string]$Id, [string]$Profile, [string]$Key = 'test.key', [string]$Meaning = 'Same meaning', [string]$Requirement = 'Required', [string]$Behavior = 'NeedsData', [string]$Role = 'Engineer', [string]$Resolution = 'Collect dataset X', [string]$DecisionKey = '')
            [pscustomobject]@{
                CriterionId = $Id; Profile = $Profile; SemanticKey = $Key; Criterion = "Criterion $Id"; Workloads = @('Exchange'); Requirement = $Requirement
                Phase = 'Design'; Scope = 'Tenant'; Population = 'Mailboxes'; Period = 'Current'; Meaning = $Meaning; Applicability = 'Always'; Coverage = 'Supported'
                Evidence = [pscustomobject]@{ Rule = 'Dataset'; Datasets = @('DatasetX'); CollectorSteps = @('Step X') }
                MissingEvidenceBehavior = $Behavior; Impact = 'Planning'; DecisionKey = $DecisionKey; RequiredDecision = 'Decide X'; DecisionCategory = 'Scope'
                BlockerImpact = 'Impact X'; ResolutionEvidence = $Resolution; ResponsibleRole = $Role
            }
        }
    }

    Context 'Profile scope: catalog contract' {
        It 'loads a valid catalog where every criterion states its scope attributes' {
            $catalog = Get-ArrayaMigrationCriteriaCatalog
            foreach ($criterion in $catalog.Criteria) {
                foreach ($field in @('Profile', 'Workloads', 'Requirement', 'Phase', 'Evidence', 'Applicability', 'MissingEvidenceBehavior', 'Impact', 'ResponsibleRole', 'ResolutionEvidence')) {
                    $criterion.$field | Should -Not -BeNullOrEmpty -Because "$($criterion.CriterionId).$field"
                }
            }
        }

        It 'keeps the TenantToTenantMigration coverage gaps explicit instead of inferring them' {
            $t2t = @((Get-ArrayaMigrationCriteriaCatalog).Criteria | Where-Object { $_.Profile -eq 'TenantToTenantMigration' })
            $expected = @{
                'EnterpriseApplications'    = 'Enterprise applications'
                'ConditionalAccessPolicies' = 'Conditional Access'
                'EntraIDGroups'             = 'Entra groups'
                'UnifiedGroups'             = 'Microsoft 365 Groups'
                'DeviceDetails'             = 'devices'
                'RetentionPolicies'         = 'Purview'
            }
            foreach ($dataset in $expected.Keys) {
                @($t2t | Where-Object { $_.Evidence.Datasets -contains $dataset -and $_.Evidence.Rule -eq 'Dataset' }).Count | Should -BeGreaterThan 0 -Because "$($expected[$dataset]) must be an explicit T2T criterion"
            }
            @($t2t | Where-Object { $_.SemanticKey -eq 'applications.sso-identity-configuration' -and $_.Coverage -eq 'Unsupported' }).Count | Should -Be 1
            @($t2t | Where-Object { $_.Evidence.Datasets -contains 'LicenseSKUs' }).Count | Should -Be 0 -Because 'no T2T coverage gap is answered from licensing'
        }

        It 'rejects a catalog whose PresalesBlocker lacks a blocked decision' {
            $path = Join-Path $TestDrive 'bad-catalog.json'
            $bad = New-SyntheticCatalog @((New-SyntheticCriterion -Id 'PS-X' -Profile 'Presales' -Behavior 'PresalesBlocker'))
            $bad.Criteria[0].RequiredDecision = ''
            $bad | ConvertTo-Json -Depth 8 | Set-Content -Path $path
            { Get-ArrayaMigrationCriteriaCatalog -Path $path } | Should -Throw '*must state*RequiredDecision*'
        }

        It 'resolves requested profiles from merged labels in canonical order' {
            Resolve-ArrayaRequestedOutputProfiles -OutputProfile 'Presales' -OutputProfileLabel 'Export-Merged(TenantToTenantMigration+Presales)' | Should -Be @('Presales', 'TenantToTenantMigration')
            Resolve-ArrayaRequestedOutputProfiles -OutputProfile 'TenantToTenantMigration' -OutputProfileLabel 'TenantToTenantMigration' | Should -Be @('TenantToTenantMigration')
            @(Resolve-ArrayaRequestedOutputProfiles -OutputProfile 'Unknown' -OutputProfileLabel '').Count | Should -Be 0
        }
    }

    Context 'Composition: single and multiple profiles' {
        It 'yields the same criteria in any request order' {
            $a = Resolve-ArrayaMigrationCriteria -OutputProfile @('TenantToTenantMigration', 'Presales') | ConvertTo-Json -Depth 6
            $b = Resolve-ArrayaMigrationCriteria -OutputProfile @('Presales', 'TenantToTenantMigration') | ConvertTo-Json -Depth 6
            $a | Should -Be $b
        }

        It 'merges identical semantic criteria and keeps owning and contributing profiles' {
            $both = @(Resolve-ArrayaMigrationCriteria -OutputProfile @('Presales', 'TenantToTenantMigration'))
            $single = @(Resolve-ArrayaMigrationCriteria -OutputProfile @('Presales')).Count + @(Resolve-ArrayaMigrationCriteria -OutputProfile @('TenantToTenantMigration')).Count
            $both.Count | Should -BeLessThan $single
            $sync = @($both | Where-Object { $_.SemanticKey -eq 'identity.directory-sync-state' })
            $sync.Count | Should -Be 1
            $sync[0].Profile | Should -Be 'Presales'
            $sync[0].ContributingProfiles | Should -Be @('TenantToTenantMigration')
            $sync[0].SourceCriterionIds | Should -Be @('PS-ID-001', 'T2T-ID-001')
        }

        It 'keeps criteria with the same key but a different scope or meaning separate' {
            $mailbox = @(Resolve-ArrayaMigrationCriteria -OutputProfile @('Presales', 'TenantToTenantMigration') | Where-Object { $_.SemanticKey -eq 'exchange.mailbox-inventory' })
            $mailbox.Count | Should -Be 2
            @($mailbox | ForEach-Object { $_.Profile }) | Should -Be @('Presales', 'TenantToTenantMigration')
        }

        It 'preserves the strictest requirement and evidence rule when merging' {
            $devices = @(Resolve-ArrayaMigrationCriteria -OutputProfile @('Presales', 'TenantToTenantMigration') | Where-Object { $_.SemanticKey -eq 'endpoints.device-inventory' })[0]
            $devices.Requirement | Should -Be 'Required'
            $devices.Profile | Should -Be 'TenantToTenantMigration'
            $devices.ContributingProfiles | Should -Be @('Presales')
            $devices.MissingEvidenceBehavior | Should -Be 'NeedsData'

            $catalog = New-SyntheticCatalog @(
                (New-SyntheticCriterion -Id 'A' -Profile 'Presales' -Requirement 'Optional' -Behavior 'PresalesBlocker'),
                (New-SyntheticCriterion -Id 'B' -Profile 'TenantToTenantMigration' -Requirement 'Required' -Behavior 'PlanningGap')
            )
            $merged = @(Resolve-ArrayaMigrationCriteria -OutputProfile @('Presales', 'TenantToTenantMigration') -Catalog $catalog)
            $merged.Count | Should -Be 1
            $merged[0].Requirement | Should -Be 'Required'
            $merged[0].MissingEvidenceBehavior | Should -Be 'PresalesBlocker'
        }

        It 'never lets another profile inherit a Presales blocker' {
            $catalog = New-SyntheticCatalog @((New-SyntheticCriterion -Id 'A' -Profile 'TenantToTenantMigration' -Behavior 'NeedsData'))
            $catalog.Criteria[0].MissingEvidenceBehavior = 'PresalesBlocker'
            (Resolve-ArrayaMigrationCriteria -OutputProfile @('TenantToTenantMigration') -Catalog $catalog)[0].MissingEvidenceBehavior | Should -Be 'NeedsData'
        }

        It 'does not merge criteria whose owner role or closure evidence differs' {
            $catalog = New-SyntheticCatalog @(
                (New-SyntheticCriterion -Id 'A' -Profile 'Presales' -Role 'Messaging engineer'),
                (New-SyntheticCriterion -Id 'B' -Profile 'TenantToTenantMigration' -Role 'Compliance owner'),
                (New-SyntheticCriterion -Id 'C' -Profile 'TenantToTenantMigration' -Role 'Messaging engineer' -Resolution 'A different closure')
            )
            @(Resolve-ArrayaMigrationCriteria -OutputProfile @('Presales', 'TenantToTenantMigration') -Catalog $catalog).Count | Should -Be 3
        }
    }

    Context 'Actions: one action per semantic condition and decision' {
        It 'produces one action for one merged condition' {
            $store = @{ AllMailboxes = @(New-Mailbox -Alias 'avery'); CollectorPlanResults = @((New-Step 'Exchange mailboxes'), (New-Step 'AD Connect sync details' -Status 'Failed' -Message 'synthetic 403')) }
            $rows = @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile @('Presales', 'TenantToTenantMigration'))
            $sync = @($rows | Where-Object { $_.SourceCriterionIds -match 'PS-ID-001' })
            $sync.Count | Should -Be 1
            $sync[0].SourceCriterionIds | Should -Be 'PS-ID-001; T2T-ID-001'
            @($rows | Where-Object { $_.ActionId -eq $sync[0].ActionId }).Count | Should -Be 1
        }

        It 'shares an action only when kind, role, scope, and closure evidence are the same' {
            $store = @{ CollectorPlanResults = @((New-Step 'Exchange mailboxes' -Status 'Failed' -Message 'synthetic EXO failure')) }
            $rows = @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile @('Presales', 'TenantToTenantMigration'))
            $psMailbox = Get-Row $rows 'PS-EXO-001'
            $t2tMailbox = Get-Row $rows 'T2T-EXO-001'
            $psSizing = Get-Row $rows 'PS-EXO-002'
            $psMailbox.ActionId | Should -Not -BeNullOrEmpty
            $t2tMailbox.ActionId | Should -Be $psMailbox.ActionId -Because 'both need the same completed mailbox collection by the same role'
            $psSizing.ActionId | Should -Not -Be $psMailbox.ActionId -Because 'mailbox statistics are different closure evidence'
        }
    }

    Context 'Collector evidence states' {
        It 'maps every collector state to an evidence status where only complete states are Ready' {
            foreach ($state in @('Partial', 'Failed', 'Unavailable', 'Unsupported')) {
                $store = @{}
                Set-ArrayaCollectorEvidenceState -TenantStatsHash $store -DatasetKey 'DatasetX' -State $state -Reason 'synthetic' | Out-Null
                $catalog = New-SyntheticCatalog @((New-SyntheticCriterion -Id 'A' -Profile 'TenantToTenantMigration'))
                $row = @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile @('TenantToTenantMigration') -Catalog $catalog)[0]
                $row.CollectorState | Should -Be $state
                $row.EvidenceStatus | Should -Not -Be 'Ready' -Because "$state is not complete evidence"
                $row.CurrentStatus | Should -Not -Be 'Ready'
            }
            foreach ($case in @(
                    @{ Store = @{}; Expected = 'Unknown' },
                    @{ Store = @{ CollectorPlanResults = @(New-Step 'Other step') }; Expected = 'NotRequested' },
                    @{ Store = @{ CollectorPlanResults = @(New-Step 'Step X' -Status 'Skipped' -Message 'Disabled for this profile.') }; Expected = 'SkippedByProfile' },
                    @{ Store = @{ CollectorPlanResults = @(New-Step 'Step X' -Status 'Skipped' -Message 'Missing dependency: Users') }; Expected = 'Unavailable' },
                    @{ Store = @{ CollectorPlanResults = @(New-Step 'Step X' -Status 'Skipped' -Message 'Requires Microsoft Graph SDK collection mode') }; Expected = 'Unsupported' }
                )) {
                $catalog = New-SyntheticCatalog @((New-SyntheticCriterion -Id 'A' -Profile 'TenantToTenantMigration'))
                $row = @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $case.Store -OutputProfile @('TenantToTenantMigration') -Catalog $catalog)[0]
                $row.CollectorState | Should -Be $case.Expected
                $row.EvidenceStatus | Should -Be 'NeedsData'
            }
        }

        It 'supports zero only for CompleteEmpty with known completeness' {
            $catalog = New-SyntheticCatalog @((New-SyntheticCriterion -Id 'A' -Profile 'TenantToTenantMigration'))
            $complete = @{ DatasetX = @{}; CollectorPlanResults = @(New-Step 'Step X') }
            (New-ArrayaMigrationCriteriaRegister -TenantStatsHash $complete -OutputProfile 'TenantToTenantMigration' -Catalog $catalog)[0].CollectorState | Should -Be 'CompleteEmpty'
            (New-ArrayaMigrationCriteriaRegister -TenantStatsHash $complete -OutputProfile 'TenantToTenantMigration' -Catalog $catalog)[0].EvidenceStatus | Should -Be 'Ready'

            $unknown = @{ DatasetX = @{} }
            (New-ArrayaMigrationCriteriaRegister -TenantStatsHash $unknown -OutputProfile 'TenantToTenantMigration' -Catalog $catalog)[0].EvidenceStatus | Should -Be 'NeedsData'

            $recordedFailure = @{ DatasetX = @{}; CollectorPlanResults = @(New-Step 'Step X') }
            Set-ArrayaCollectorEvidenceState -TenantStatsHash $recordedFailure -DatasetKey 'DatasetX' -State 'Failed' -Reason 'synthetic' | Out-Null
            (New-ArrayaMigrationCriteriaRegister -TenantStatsHash $recordedFailure -OutputProfile 'TenantToTenantMigration' -Catalog $catalog)[0].CollectorState | Should -Be 'Failed'

            (Get-ArrayaDatasetEvidenceState -TenantStatsHash @{ Domains = @{}; CollectorPlanResults = @(New-Step 'Domains') } -DatasetKey 'Domains' -CollectorStepName 'Domains' -EmptyIsUnexpected).CollectorState | Should -Be 'Unknown'
        }

        It 'keeps collectors skipped by the TenantToTenantMigration profile visible' {
            $store = @{
                AllMailboxes = @(New-Mailbox -Alias 'avery' -WithDelegates)
                CollectorPlanResults = @(
                    (New-Step 'Exchange mailboxes'),
                    (New-Step 'Entra groups' -Status 'Skipped' -Message 'Entra group inventory is disabled for this profile.'),
                    (New-Step 'Authentication/SSO configuration' -Status 'Skipped' -Message 'Authentication and SSO collection is disabled for this profile.'),
                    (New-Step 'Unified groups' -Status 'Skipped' -Message 'Not required for this profile output.'),
                    (New-Step 'Devices' -Status 'Skipped' -Message 'Device inventory is disabled for this profile.'),
                    (New-Step 'Conditional Access policies' -Status 'Skipped' -Message 'Conditional Access collection is disabled for this profile.'),
                    (New-Step 'Purview retention/DLP policies' -Status 'Skipped' -Message 'Governance compliance collection is disabled for this profile.')
                )
            }
            $rows = @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile @('TenantToTenantMigration'))
            foreach ($id in @('T2T-APP-001', 'T2T-GRP-001', 'T2T-GRP-002', 'T2T-DEV-001', 'T2T-CA-001', 'T2T-PUR-001')) {
                $row = Get-Row $rows $id
                $row | Should -Not -BeNullOrEmpty -Because "$id must stay visible"
                $row.CollectorState | Should -Be 'SkippedByProfile'
                $row.EvidenceStatus | Should -Be 'NeedsData'
                $row.CurrentStatus | Should -Be 'Needs Data'
                $row.PresalesClassification | Should -Be 'NeedsData'
            }
            (Get-Row $rows 'T2T-EXO-002').EvidenceStatus | Should -Be 'Ready'
        }

        It 'keeps unsupported evidence visible as a planning gap that needs customer input' {
            $store = New-CompleteStore
            $store.AdConnectConfiguration = New-SyncSummary -Flag $true
            $rows = @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile @('TenantToTenantMigration'))
            foreach ($id in @('T2T-SSO-001', 'T2T-ID-002')) {
                $row = Get-Row $rows $id
                $row.CollectorState | Should -Be 'Unsupported'
                $row.EvidenceStatus | Should -Be 'NeedsData'
                $row.PresalesClassification | Should -Be 'PlanningGap'
                $row.CurrentStatus | Should -Be 'Needs Input'
                $row.CollectorOrSource | Should -Match 'no collector'
            }
        }

        It 'reports a conditional criterion as not applicable only on affirmative evidence' {
            $cloudOnly = New-CompleteStore
            (Get-Row @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $cloudOnly -OutputProfile 'TenantToTenantMigration') 'T2T-ID-002').EvidenceStatus | Should -Be 'NotApplicable'

            $unknown = New-CompleteStore
            $unknown.Remove('TenantInfo')
            $row = Get-Row @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $unknown -OutputProfile 'TenantToTenantMigration') 'T2T-ID-002'
            $row.EvidenceStatus | Should -Be 'NeedsData'
            $row.MissingEvidence | Should -Match 'Applicability not established'
        }

        It 'records collector evidence for public folders (fictional cmdlet responses)' {
            function global:Get-PublicFolder { [CmdletBinding()] param([switch]$Recurse, $ResultSize) Write-Error 'No active public folder mailboxes were found for organization contoso.example.' }
            function global:Get-PublicFolderStatistics { [CmdletBinding()] param([Parameter(ValueFromPipeline = $true)]$InputObject) process { } }
            function global:get-publicfolderclientpermission { [CmdletBinding()] param([Parameter(ValueFromPipeline = $true)]$InputObject) process { } }
            try {
                $stats = [ordered]@{}
                Get-AllPublicFolderDetails -detailLevel 'all' -Context ([pscustomobject]@{ TenantStats = $stats; ExportFileLocation = $null; Policies = [ordered]@{}; Runtime = [ordered]@{}; Metadata = [ordered]@{} }) -ErrorAction SilentlyContinue
                $stats['CollectorEvidenceStates']['PublicFolderDetails'].State | Should -Be 'CompleteEmpty'

                function global:Get-PublicFolder { [CmdletBinding()] param([switch]$Recurse, $ResultSize) Write-Error 'Access denied: the user is not authorized to run Get-PublicFolder.' }
                $stats = [ordered]@{}
                Get-AllPublicFolderDetails -detailLevel 'all' -Context ([pscustomobject]@{ TenantStats = $stats; ExportFileLocation = $null; Policies = [ordered]@{}; Runtime = [ordered]@{}; Metadata = [ordered]@{} }) -ErrorAction SilentlyContinue
                $stats['CollectorEvidenceStates']['PublicFolderDetails'].State | Should -Be 'Unavailable'
                (Get-ArrayaDatasetEvidenceState -TenantStatsHash $stats -DatasetKey 'PublicFolderDetails').CollectorState | Should -Be 'Unavailable'
            }
            finally {
                Remove-Item Function:\global:Get-PublicFolder, Function:\global:Get-PublicFolderStatistics, Function:\global:get-publicfolderclientpermission -ErrorAction SilentlyContinue
            }
        }

        It 'records collector evidence for mail flow rules and connectors (fictional cmdlet responses)' {
            function global:Get-TransportRule { [CmdletBinding()] param([switch]$IncludeTestModeConnectors) Write-Error 'synthetic transport rule failure' }
            function global:Get-InboundConnector { [CmdletBinding()] param() throw 'synthetic connector failure' }
            function global:Get-OutboundConnector { [CmdletBinding()] param($IncludeTestModeConnectors) @() }
            try {
                $stats = [ordered]@{}
                Get-MailFlowRulesandConnectors -detailLevel 'all' -Context ([pscustomobject]@{ TenantStats = $stats; ExportFileLocation = $null; Policies = [ordered]@{}; Runtime = [ordered]@{}; Metadata = [ordered]@{ StartedAt = (Get-Date) } }) -ErrorAction SilentlyContinue
                $stats['CollectorEvidenceStates']['MailFlowRules'].State | Should -Be 'Failed'
                $stats['CollectorEvidenceStates']['MailFlowConnectors'].State | Should -Be 'Failed'
                (Get-ArrayaDatasetEvidenceState -TenantStatsHash $stats -DatasetKey 'MailFlowConnectors').State | Should -Be 'Failed'
            }
            finally {
                Remove-Item Function:\global:Get-TransportRule, Function:\global:Get-InboundConnector, Function:\global:Get-OutboundConnector -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'Holds: evidence status separate from migration treatment' {
        It 'reports no holds as Ready and Clear only with complete evidence' {
            $hold = Get-ArrayaMailboxHoldTreatment -TenantStatsHash (New-CompleteStore)
            $hold.EvidenceStatus | Should -Be 'Ready'
            $hold.MigrationTreatment | Should -Be 'Clear'

            $store = New-CompleteStore
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store
            @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Mailboxes on hold' })[0].Status | Should -Be 'Ready'
        }

        It 'keeps evidence Ready when holds are present, and does not treat it as clearance' {
            $store = New-CompleteStore -Mailboxes @((New-Mailbox -Alias 'policy' -Holds @{ InPlaceHolds = 'mbx0123456789abcdef0123456789abcdef:1' }))
            $hold = Get-ArrayaMailboxHoldTreatment -TenantStatsHash $store
            $hold.EvidenceStatus | Should -Be 'Ready'
            $hold.MigrationTreatment | Should -Be 'RequiresComplianceValidation'

            $row = Get-Row @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile 'TenantToTenantMigration') 'T2T-CMP-001'
            $row.EvidenceStatus | Should -Be 'Ready'
            $row.MigrationTreatment | Should -Be 'RequiresComplianceValidation'
            $row.CurrentStatus | Should -Not -Be 'Ready'
            $row.PresalesClassification | Should -Be 'PlanningGap'
            $row.Notes | Should -Match 'not migration clearance'
        }

        It 'treats an organization-wide Exchange retention policy as a retention-policy hold on every mailbox' {
            $store = New-CompleteStore
            $store.RetentionPolicies = @{ '001-All' = [pscustomobject]@{ PolicyName = 'Synthetic retain all'; Enabled = $true; ExchangeLocation = 'All' } }
            $hold = Get-ArrayaMailboxHoldTreatment -TenantStatsHash $store
            $hold.MigrationTreatment | Should -Be 'RequiresComplianceValidation'
            $hold.OrganizationWidePolicies | Should -Be @('Synthetic retain all')
        }

        It 'states that a legal decision is required without making it' {
            $store = New-CompleteStore -Mailboxes @((New-Mailbox -Alias 'lit' -Holds @{ LitigationHoldEnabled = $true }))
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store
            $row = Get-Row @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile 'TenantToTenantMigration') 'T2T-CMP-001'
            $row.MigrationTreatment | Should -Be 'RequiresLegalDecision'
            $row.CurrentStatus | Should -Be 'Needs Input'
            $row.Notes | Should -Match 'legal owner must decide'
            $decision = @($store['MigrationScopeDecisions'] | Where-Object { $_.DecisionKey -eq 'CP-01' })[0]
            $decision.Status | Should -Be 'Needs Input'
            $decision.CustomerConfirmedValue | Should -BeNullOrEmpty
            @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Mailboxes on hold' })[0].Status | Should -Be 'Review'
        }

        It 'reports incomplete hold evidence as NotEstablished, never Clear' {
            $missingField = New-CompleteStore -Mailboxes @((New-Mailbox -Alias 'bare' -OmitHoldFields))
            $hold = Get-ArrayaMailboxHoldTreatment -TenantStatsHash $missingField
            $hold.EvidenceStatus | Should -Be 'Partial'
            $hold.MigrationTreatment | Should -Be 'NotEstablished'

            $noPurview = New-CompleteStore
            $noPurview.Remove('RetentionPolicies')
            (Get-ArrayaMailboxHoldTreatment -TenantStatsHash $noPurview).MigrationTreatment | Should -Be 'NotEstablished'

            $noMailboxes = @{ CollectorPlanResults = @(New-Step 'Exchange mailboxes' -Status 'Failed' -Message 'synthetic') }
            $hold = Get-ArrayaMailboxHoldTreatment -TenantStatsHash $noMailboxes
            $hold.EvidenceStatus | Should -Be 'Failed'
            $hold.MigrationTreatment | Should -Be 'NotEstablished'
        }

        It 'blocks migration only when a confirmed method is documented as incompatible with the hold' {
            $store = New-CompleteStore -Mailboxes @((New-Mailbox -Alias 'lit' -Holds @{ LitigationHoldEnabled = $true }))
            $store.MigrationScopeDecisions = @([pscustomobject]@{ DecisionKey = 'MAILBOX:lit@contoso.example'; Status = 'Approved'; MigrationTool = 'Microsoft cross-tenant mailbox migration'; InScope = $true; TargetMapping = 'lit@target.example' })
            $hold = Get-ArrayaMailboxHoldTreatment -TenantStatsHash $store
            $hold.EvidenceStatus | Should -Be 'Ready'
            $hold.MigrationTreatment | Should -Be 'BlocksMigration'
            $hold.BlockingRuleIds | Should -Be @('HOLD-NATIVE-CTM')

            Update-TenantToTenantMigrationExportData -TenantStatsStore $store
            @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Mailboxes on hold' })[0].Status | Should -Be 'Blocker'
            (Get-Row @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile 'TenantToTenantMigration') 'T2T-CMP-001').CurrentStatus | Should -Be 'Blocker'

            $copyTool = New-CompleteStore -Mailboxes @((New-Mailbox -Alias 'lit' -Holds @{ LitigationHoldEnabled = $true }))
            $copyTool.MigrationScopeDecisions = @([pscustomobject]@{ DecisionKey = 'MAILBOX:lit@contoso.example'; Status = 'Approved'; MigrationTool = 'Third-party copy tool' })
            (Get-ArrayaMailboxHoldTreatment -TenantStatsHash $copyTool).MigrationTreatment | Should -Be 'RequiresLegalDecision'

            $unconfirmed = New-CompleteStore -Mailboxes @((New-Mailbox -Alias 'lit' -Holds @{ LitigationHoldEnabled = $true }))
            $unconfirmed.MigrationScopeDecisions = @([pscustomobject]@{ DecisionKey = 'MAILBOX:lit@contoso.example'; Status = 'Needs Input'; MigrationTool = 'Microsoft cross-tenant mailbox migration' })
            (Get-ArrayaMailboxHoldTreatment -TenantStatsHash $unconfirmed).MigrationTreatment | Should -Be 'RequiresLegalDecision'
        }
    }

    Context 'Presales classification' {
        It 'raises a PresalesBlocker that names every required element' {
            $store = @{ CollectorPlanResults = @(New-Step 'Exchange mailboxes' -Status 'Failed' -Message 'synthetic EXO failure') }
            $row = Get-Row @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile 'Presales') 'PS-EXO-001'
            $row.PresalesClassification | Should -Be 'PresalesBlocker'
            $row.CurrentStatus | Should -Be 'Blocker'
            $row.MissingEvidence | Should -Match 'synthetic EXO failure'
            $row.RequiredDecision | Should -Not -BeNullOrEmpty
            $row.BlockerImpact | Should -Not -BeNullOrEmpty
            $row.ResolutionEvidence | Should -Not -BeNullOrEmpty
            $row.ResponsibleRole | Should -Not -BeNullOrEmpty
            $row.Profile | Should -Be 'Presales'
            $row.Workloads | Should -Be 'Exchange'
        }

        It 'separates NeedsData from PlanningGap' {
            $store = New-CompleteStore
            $store.AdConnectConfiguration = New-SyncSummary -Flag $true
            $store.HybridConfiguration = New-Hybrid -IsHybrid $null -Complete $false -Status 'Unknown'
            $rows = @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile 'Presales')
            (Get-Row $rows 'PS-EXO-003').PresalesClassification | Should -Be 'NeedsData' -Because 'collectible evidence for a non-gating decision'
            (Get-Row $rows 'PS-ID-002').PresalesClassification | Should -Be 'PlanningGap' -Because 'not collectible; needs customer input'
            @($rows | Where-Object { $_.PresalesClassification -eq 'PresalesBlocker' }).Count | Should -Be 0
        }

        It 'never blocks Presales on a missing optional input' {
            $store = New-CompleteStore
            $store.Remove('DeviceDetails')
            $store.CollectorPlanResults = @($store.CollectorPlanResults | Where-Object { $_.Name -ne 'Devices' }) + @(New-Step 'Devices' -Status 'Skipped' -Message 'Device inventory is disabled for this profile.')
            $row = Get-Row @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile 'Presales') 'PS-DEV-001'
            $row.EvidenceStatus | Should -Be 'NeedsData'
            $row.PresalesClassification | Should -Be 'Nonblocking'
            $row.CurrentStatus | Should -Be 'Needs Data'
            $row.Notes | Should -Match 'Optional evidence'
        }

        It 'blocks the Presales quote on unknown synchronization but not on unknown Exchange hybrid' {
            $hybridUnknown = @{ TenantInfo = [pscustomobject]@{ TenantId = '00000000-0000-0000-0000-000000000001' }; AllMailboxes = @(New-Mailbox -Alias 'avery'); AdConnectConfiguration = New-SyncSummary -Flag $null; HybridConfiguration = New-Hybrid -IsHybrid $null -Complete $false -Status 'Unknown' }
            Update-MigrationScopeExportData -TenantStatsStore $hybridUnknown
            $hybridRow = @($hybridUnknown['MigrationQuoteReadiness'] | Where-Object { $_.DecisionKey -eq 'FLAG:exchange-hybrid' })[0]
            $hybridRow.Status | Should -Be 'Condition'
            $hybridRow.Blocking | Should -BeFalse

            $syncUnknown = @{ AllMailboxes = @(New-Mailbox -Alias 'avery'); HybridConfiguration = New-Hybrid -IsHybrid $false }
            Update-MigrationScopeExportData -TenantStatsStore $syncUnknown
            $syncRow = @($syncUnknown['MigrationQuoteReadiness'] | Where-Object { $_.DecisionKey -eq 'FLAG:ad-connect-directory-synchronization' })[0]
            $syncRow.Status | Should -Be 'Blocker'
            $syncRow.Blocking | Should -BeTrue
        }
    }

    Context 'Composition never fills evidence gaps' {
        It 'keeps the T2T SSO identity criterion unsupported when Presales collected the application inventory' {
            $store = New-CompleteStore
            $store.EnterpriseApplications = @{ app = [pscustomobject]@{ DisplayName = 'Synthetic SaaS'; PreferredSingleSignOnMode = 'saml' } }
            $rows = @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile @('Presales', 'TenantToTenantMigration'))
            (Get-Row $rows 'PS-APP-001').EvidenceStatus | Should -Be 'Ready'
            (Get-Row $rows 'T2T-SSO-001').CollectorState | Should -Be 'Unsupported'
        }

        It 'shows delegates as not requested in a combined run instead of borrowing other evidence' {
            $store = New-CompleteStore -Mailboxes @((New-Mailbox -Alias 'avery'))
            $row = Get-Row @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $store -OutputProfile @('Presales', 'TenantToTenantMigration')) 'T2T-EXO-002'
            $row.CollectorState | Should -Be 'NotRequested'
            $row.MissingEvidence | Should -Match 'only when TenantToTenantMigration runs on its own'
        }
    }

    Context 'Planning-register projection: JSON and workbook agree' {
        BeforeAll {
            $script:projectionStore = New-CompleteStore -Mailboxes @((New-Mailbox -Alias 'lit' -Holds @{ LitigationHoldEnabled = $true }))
            $script:projectionStore.CollectorPlanResults = @($script:projectionStore.CollectorPlanResults) + @(New-Step 'Entra groups' -Status 'Skipped' -Message 'Entra group inventory is disabled for this profile.')
            $script:projectionStore['MigrationCriteriaRegister'] = @(New-ArrayaMigrationCriteriaRegister -TenantStatsHash $script:projectionStore -OutputProfile @('Presales', 'TenantToTenantMigration'))
            $script:registerColumns = @('CriterionId', 'SourceCriterionIds', 'Profile', 'ContributingProfiles', 'Criterion', 'Workloads', 'Requirement', 'Phase', 'Applicability', 'Impact', 'CollectorOrSource', 'CollectorState', 'EvidenceStatus', 'MissingEvidence', 'PresalesClassification', 'MigrationTreatment', 'RequiredDecision', 'DecisionKey', 'BlockerImpact', 'ResolutionEvidence', 'ResponsibleRole', 'CurrentStatus', 'ActionId', 'Action', 'Notes')
            function script:ConvertTo-CellText { param($Value) if ($null -eq $Value) { '' } else { ([string]$Value).Trim() } }
        }

        It 'emits the documented column order' {
            @($script:projectionStore['MigrationCriteriaRegister'][0].PSObject.Properties.Name) | Should -Be $script:registerColumns
        }

        It 'writes the same rows to the T2T and Presales workbooks and the snapshot' {
            $snapshot = Convert-ArrayaLegacyTenantStatsToSnapshot -TenantStatsHash $script:projectionStore
            $jsonRows = @(($snapshot | ConvertTo-Json -Depth 20 | ConvertFrom-Json).Derived.MigrationCriteriaRegister)
            $jsonRows.Count | Should -Be @($script:projectionStore['MigrationCriteriaRegister']).Count

            foreach ($policy in @('TenantToTenantCutover', 'Presales')) {
                $path = Join-Path $TestDrive "register-$policy.xlsx"
                Export-HashTableToExcel -hashtable $script:projectionStore -ExportDetails $path -WorkbookExportPolicy $policy
                $sheet = @(Import-Excel -Path $path -WorksheetName 'MigrationCriteriaRegister')
                @($sheet[0].PSObject.Properties.Name) | Should -Be $script:registerColumns -Because "$policy keeps a stable register schema"
                $sheet.Count | Should -Be $jsonRows.Count
                for ($i = 0; $i -lt $jsonRows.Count; $i++) {
                    foreach ($column in $script:registerColumns) {
                        (ConvertTo-CellText $sheet[$i].$column) | Should -Be (ConvertTo-CellText $jsonRows[$i].$column) -Because "$policy row $i column $column"
                    }
                }
                @(Get-ExcelSheetInfo -Path $path | Select-Object -ExpandProperty Name) | Should -Not -Contain 'CollectorEvidenceStates'
            }
        }

        It 'restores collector evidence records from the snapshot diagnostics' {
            $store = New-CompleteStore
            Set-ArrayaCollectorEvidenceState -TenantStatsHash $store -DatasetKey 'PublicFolderDetails' -State 'Failed' -Reason 'synthetic' | Out-Null
            $snapshot = Convert-ArrayaLegacyTenantStatsToSnapshot -TenantStatsHash $store
            $snapshot['Diagnostics']['CollectorStats'].Contains('CollectorEvidenceStates') | Should -BeTrue
            $restored = Convert-ArrayaSnapshotToLegacyTenantStatsHash -Snapshot $snapshot
            (Get-ArrayaDatasetEvidenceState -TenantStatsHash $restored -DatasetKey 'PublicFolderDetails').CollectorState | Should -Be 'Failed'
        }
    }

    Context 'Compatibility of existing contracts' {
        It 'keeps the cutover pack CSV columns unchanged' {
            $store = New-CompleteStore
            $result = Export-ArrayaTenantToTenantCutoverPack -TenantStatsHash $store -BaseExportPath (Join-Path $TestDrive 'Contoso-Assess.xlsx')
            $expected = @{
                UserWaveTemplate            = 'SourceAccount,SourceUPN,SourcePrimarySmtpAddress,MailboxPrimarySmtpAddress,RecipientTypeDetails,HasMailbox,IsDirSynced,TargetAccount,TargetUPN,TargetPrimarySmtpAddress,TargetMail,Wave,MigrationState,CutoverDate,Notes'
                CalendarDelegateReference   = 'MailboxDisplayName,MailboxPrimarySmtpAddress,MailboxUserPrincipalName,RecipientTypeDetails,CalendarName,CalendarPath,PermissionTarget,PermissionTargetDisplayName,PermissionTargetType,AccessRights,SharingPermissionFlags,Wave,Notes'
                RecipientDomainReference    = 'DisplayName,PrimarySmtpAddress,RecipientTypeDetails,UserPrincipalName,WindowsEmailAddress,EmailAddresses,OnMicrosoftAlias,OnMicrosoftAliases,HiddenFromAddressListsEnabled,Wave,Notes'
                CollaborationAccessTemplate = 'SourceAccount,TargetAccount,SourceUPN,TargetUPN,SourcePrimarySmtpAddress,Direction,IncludeEntraGroups,IncludeTeams,IncludeSharePoint,ApprovalStatus,ApprovalNotes,Wave,Notes'
                OneDriveMigrationTemplate   = 'SourceUPN,SourceEmail,SourceOneDriveUrl,HasOneDrive,TargetMail,TargetUPN,TargetOneDriveUrl,TargetFolderName,Wave,Notes'
            }
            foreach ($sheet in $expected.Keys) {
                ((Get-Content -Path $result.CsvArtifacts[$sheet] -TotalCount 1) -replace '"', '') | Should -Be $expected[$sheet] -Because $sheet
            }
        }

        It 'keeps the CutoverPrepSummary and quote readiness columns unchanged' {
            $store = New-CompleteStore
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store
            @($store['CutoverPrepSummary'][0].PSObject.Properties.Name) | Should -Be @('Category', 'Item', 'Status', 'Value', 'Notes')
            @($store['MigrationQuoteReadiness'][0].PSObject.Properties.Name) | Should -Be @('RowType', 'ReadinessLevel', 'Category', 'DecisionKey', 'Requirement', 'Status', 'Blocking', 'DiscoveredValue', 'EvidenceSource', 'CustomerConfirmedValue', 'Resolution', 'DecisionOwner', 'DueDate', 'Notes')
        }
    }
}
