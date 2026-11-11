# Regression tests for the tenant-to-tenant P0 correctness batch (P0-01 to P0-07).
# See docs/t2t-migration-review. Every fixture is synthetic and uses reserved .example
# domains. No test contacts a tenant or performs a DNS query: Microsoft Graph, Exchange and
# DNS calls are replaced by local stub functions, and the production functions under test
# are loaded from the shipped scripts.
Describe 'T2T P0 correctness' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        Import-Module -Name (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -WarningAction SilentlyContinue -ErrorAction Stop

        # Quiet stand-ins for logging and console helpers used by the legacy functions.
        function Write-Log { param() }
        function Write-AssessmentCollectorBanner { param() }
        function Write-AssessmentCollectorCompletionBanner { param() }
        function Write-ProgressHelper { param() }
        function Get-MgDirectoryOnPremiseSynchronization { throw 'synthetic: directory synchronization features unavailable' }

        $script:collectorPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:collectorPath, [ref]$tokens, [ref]$parseErrors)
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
            'Update-TenantToTenantMigrationExportData', 'Get-TenantOverviewInfo', 'Get-AdConnectSyncDetails',
            'Get-AllOffice365Domains', 'Resolve-AssessmentProfileCollectionPlan', 'Get-ArrayaCollectionDepthPolicy',
            'Normalize-ArrayaReportingMode'
        )
        $definitions = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $true) | Where-Object { $wantedFunctions -contains $_.Name }
        foreach ($definition in $definitions) {
            . ([scriptblock]::Create($definition.Extent.Text))
        }
        $script:AssessmentBitTitanLicenseModel = Get-ArrayaBitTitanLicenseModel

        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')
        . (Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Export-TenantToTenantQuestionnaireMarkdown.ps1')
        $script:questionnaireTemplatePath = Join-Path $script:repoRoot 'docs\templates\Microsoft 365 Tenant to Tenant Questionnaire.md'

        # ---- Synthetic fixture builders -------------------------------------------------
        function script:New-T2TCollectorResult {
            param([string]$Name, [string]$Status = 'Completed', [string]$Message, [string[]]$ProducedKeys = @())
            [pscustomobject]@{ Name = $Name; Section = 'Test'; Workload = 'Test'; Status = $Status; Required = $false; Message = $Message; DurationSeconds = 0; ProducedKeys = $ProducedKeys }
        }

        function script:New-T2TMailbox {
            param(
                [string]$Alias = 'avery',
                [hashtable]$Extra = @{}
            )
            $row = [ordered]@{
                DisplayName          = $Alias
                UserPrincipalName    = "$Alias@contoso.example"
                PrimarySmtpAddress   = "$Alias@contoso.example"
                RecipientTypeDetails = 'UserMailbox'
                ExchangeGuid         = "aaaaaaaa-0000-0000-0000-$(([string]([math]::Abs($Alias.GetHashCode()))).PadLeft(12, '0').Substring(0, 12))"
                MailboxSizeGB        = 2
                ArchiveStatus        = 'None'
            }
            foreach ($key in $Extra.Keys) { $row[$key] = $Extra[$key] }
            [pscustomobject]$row
        }

        function script:New-T2TTenantInfo {
            [pscustomobject]@{ DisplayName = 'Contoso Synthetic'; TenantId = '00000000-0000-0000-0000-000000000001'; InitialDomain = 'contoso.onmicrosoft.com'; DefaultDomain = 'contoso.example' }
        }

        function script:New-T2TSyncConfiguration {
            param($OnPremisesSyncEnabled)
            @{
                Summary      = [pscustomobject]@{ OnPremisesSyncEnabled = $OnPremisesSyncEnabled; OnPremisesLastSyncDateTime = $null; PasswordSyncEnabled = $null }
                SyncServices = @()
                RecentErrors = $null
                ErrorCount   = 0
            }
        }

        function script:New-T2THybridConfiguration {
            param($IsHybridConfigured, [bool]$EvidenceComplete = $true, [int]$EvidenceCount = 0, [string]$HybridStatus = 'None', [string]$UnansweredProbes = '')
            @{
                ExchangeHybrid = [pscustomobject]@{
                    IsHybridConfigured     = $IsHybridConfigured
                    HybridStatus           = $HybridStatus
                    HybridType             = 'None'
                    EvidenceComplete       = $EvidenceComplete
                    UnansweredProbes       = $UnansweredProbes
                    EvidenceCount          = $EvidenceCount
                    MigrationEndpointCount = 0
                }
            }
        }

        function script:New-T2TDomain {
            param([string]$Name = 'contoso.example', $Office365MailExchanger = $true, [bool]$Verified = $true, [string]$MXRecords = 'contoso-example.mail.protection.outlook.com')
            [pscustomobject]@{ Domain = $Name; Name = $Name; Verified = $Verified; IsDefault = $true; AuthenticationType = 'Managed'; DomainType = 'Authoritative'; MXRecords = $MXRecords; Office365MailExchanger = $Office365MailExchanger }
        }

        function script:Get-T2TReadinessRow {
            param([hashtable]$Store, [string]$Item)
            @(@($Store['MigrationReadiness'].Values) | Where-Object { $_.Item -eq $Item })[0]
        }

        function script:Get-T2TQuestionnaireLine {
            param([hashtable]$Store, [string]$Pattern)
            $path = Join-Path $TestDrive ("questionnaire-{0}.md" -f ([guid]::NewGuid().ToString('N')))
            Export-TenantToTenantQuestionnaireMarkdown -TenantStatsHash $Store -TemplatePath $script:questionnaireTemplatePath -Path $path -PreparedBy 'synthetic'
            @(Get-Content -Path $path | Where-Object { $_ -match $Pattern })[0]
        }
    }

    Context 'Shared interpretation: dataset evidence state' {
        It 'reports Collected when the dataset has rows' {
            $state = Get-ArrayaDatasetEvidenceState -TenantStatsHash @{ PublicFolderDetails = @([pscustomobject]@{ Name = 'pf' }) } -DatasetKey 'PublicFolderDetails'
            $state.State | Should -Be 'Collected'
            $state.RowCount | Should -Be 1
        }

        It 'does not treat a present but empty dataset as collected when nothing shows the collection completed' {
            # Tightened in the profile-composition pass: CompleteEmpty needs a completed
            # collector step or the collector's own record, not just an empty container.
            $state = Get-ArrayaDatasetEvidenceState -TenantStatsHash @{ PublicFolderDetails = @{} } -DatasetKey 'PublicFolderDetails'
            $state.State | Should -Be 'Unknown'
            $state.CollectorState | Should -Be 'Unknown'
            $state.RowCount | Should -Be 0
        }

        It 'treats a present but empty dataset as complete when its collector step completed' {
            $store = @{ PublicFolderDetails = @{}; CollectorPlanResults = @(New-T2TCollectorResult -Name 'Public folders') }
            $state = Get-ArrayaDatasetEvidenceState -TenantStatsHash $store -DatasetKey 'PublicFolderDetails' -CollectorStepName 'Public folders'
            $state.State | Should -Be 'Collected'
            $state.CollectorState | Should -Be 'CompleteEmpty'
        }

        It 'reports NotCollected when the dataset is absent' {
            (Get-ArrayaDatasetEvidenceState -TenantStatsHash @{} -DatasetKey 'PublicFolderDetails').State | Should -Be 'NotCollected'
        }

        It 'reports NotCollected with the skip reason when the collector step was skipped' {
            $store = @{ CollectorPlanResults = @(New-T2TCollectorResult -Name 'Public folders' -Status 'Skipped' -Message 'Not required for this profile output.') }
            $state = Get-ArrayaDatasetEvidenceState -TenantStatsHash $store -DatasetKey 'PublicFolderDetails' -CollectorStepName 'Public folders'
            $state.State | Should -Be 'NotCollected'
            $state.Reason | Should -Match 'Not required for this profile output'
        }

        It 'reports Failed when the collector step failed, even if an empty container exists' {
            $store = @{ PublicFolderDetails = @{}; CollectorPlanResults = @(New-T2TCollectorResult -Name 'Public folders' -Status 'Failed' -Message 'synthetic 403') }
            $state = Get-ArrayaDatasetEvidenceState -TenantStatsHash $store -DatasetKey 'PublicFolderDetails' -CollectorStepName 'Public folders'
            $state.State | Should -Be 'Failed'
            $state.Reason | Should -Match 'synthetic 403'
        }

        It 'does not accept an empty dataset that cannot legitimately be empty' {
            (Get-ArrayaDatasetEvidenceState -TenantStatsHash @{ Domains = @{} } -DatasetKey 'Domains' -EmptyIsUnexpected).State | Should -Be 'Unknown'
        }

        It 'keeps collected rows authoritative when a different step was skipped' {
            $store = @{ RemoteDomains = @([pscustomobject]@{ DomainName = '*' }); CollectorPlanResults = @(New-T2TCollectorResult -Name 'Mail flow rules/connectors' -Status 'Skipped' -Message 'skipped') }
            (Get-ArrayaDatasetEvidenceState -TenantStatsHash $store -DatasetKey 'RemoteDomains' -CollectorStepName 'Mail flow rules/connectors').State | Should -Be 'Collected'
        }
    }

    Context 'P0-01 and P0-02: directory synchronization evidence' {
        It 'reports Synchronized when the organization flag is true (Entra Connect or Cloud Sync, not distinguished)' {
            $evidence = Get-ArrayaDirectorySyncEvidence -TenantStatsHash @{ TenantInfo = New-T2TTenantInfo; AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $true }
            $evidence.State | Should -Be 'Synchronized'
        }

        It 'reports NotSynchronized only when the organization read succeeded and reports no synchronization' {
            $evidence = Get-ArrayaDirectorySyncEvidence -TenantStatsHash @{ TenantInfo = New-T2TTenantInfo; AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $null }
            $evidence.State | Should -Be 'NotSynchronized'
        }

        It 'reports NotSynchronized for an explicit false' {
            (Get-ArrayaDirectorySyncEvidence -TenantStatsHash @{ AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $false }).State | Should -Be 'NotSynchronized'
        }

        It 'reports Unknown when the flag is empty and nothing shows the organization read succeeded' {
            (Get-ArrayaDirectorySyncEvidence -TenantStatsHash @{ AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $null }).State | Should -Be 'Unknown'
        }

        It 'reports Unknown when the collector did not run' {
            (Get-ArrayaDirectorySyncEvidence -TenantStatsHash @{ TenantInfo = New-T2TTenantInfo }).State | Should -Be 'Unknown'
        }

        It 'reports Unknown when the collector step failed' {
            $store = @{ TenantInfo = New-T2TTenantInfo; AdConnectConfiguration = @{}; CollectorPlanResults = @(New-T2TCollectorResult -Name 'AD Connect sync details' -Status 'Failed' -Message 'synthetic failure') }
            (Get-ArrayaDirectorySyncEvidence -TenantStatsHash $store).State | Should -Be 'Unknown'
        }

        It 'accepts synchronized user objects as affirmative evidence when the organization flag is unavailable' {
            $store = @{
                AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $null
                Users = @([pscustomobject]@{ UserPrincipalName = 'avery@contoso.example'; OnPremisesSyncEnabled = $true })
            }
            $evidence = Get-ArrayaDirectorySyncEvidence -TenantStatsHash $store
            $evidence.State | Should -Be 'Synchronized'
            $evidence.SyncedUserCount | Should -Be 1
        }
    }

    Context 'P0-01: Exchange hybrid evidence' {
        It 'reports Configured for confirmed hybrid' {
            (Get-ArrayaExchangeHybridEvidence -TenantStatsHash @{ HybridConfiguration = New-T2THybridConfiguration -IsHybridConfigured $true -EvidenceCount 2 -HybridStatus 'Hybrid' }).State | Should -Be 'Configured'
        }

        It 'reports NotConfigured only for an explicit false backed by complete probes' {
            (Get-ArrayaExchangeHybridEvidence -TenantStatsHash @{ HybridConfiguration = New-T2THybridConfiguration -IsHybridConfigured $false }).State | Should -Be 'NotConfigured'
        }

        It 'reports Indicators when partial signals exist without confirmed hybrid' {
            (Get-ArrayaExchangeHybridEvidence -TenantStatsHash @{ HybridConfiguration = New-T2THybridConfiguration -IsHybridConfigured $false -EvidenceCount 1 -HybridStatus 'Federation/Relationship' }).State | Should -Be 'Indicators'
        }

        It 'reports Unknown when probes did not answer' {
            $evidence = Get-ArrayaExchangeHybridEvidence -TenantStatsHash @{ HybridConfiguration = New-T2THybridConfiguration -IsHybridConfigured $null -EvidenceComplete $false -HybridStatus 'PermissionDenied' -UnansweredProbes 'InboundConnector (PermissionDenied)' }
            $evidence.State | Should -Be 'Unknown'
            $evidence.Reason | Should -Match 'InboundConnector'
        }

        It 'reports Unknown when the collector did not run' {
            (Get-ArrayaExchangeHybridEvidence -TenantStatsHash @{}).State | Should -Be 'Unknown'
        }
    }

    Context 'P0-01: migration complexity flags' {
        It 'raises neither flag for confirmed cloud-only identity without Exchange hybrid' {
            $store = @{
                TenantInfo             = New-T2TTenantInfo
                Users                  = @([pscustomobject]@{ UserPrincipalName = 'avery@contoso.example'; Mail = 'avery@contoso.example'; UserType = 'Member'; AccountEnabled = $true })
                AllMailboxes           = @(New-T2TMailbox)
                AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $null
                HybridConfiguration    = New-T2THybridConfiguration -IsHybridConfigured $false
            }
            Update-MigrationScopeExportData -TenantStatsStore $store

            $items = @($store['MigrationComplexityFlags'] | ForEach-Object { $_.Item })
            $items | Should -Not -Contain 'AD Connect directory synchronization'
            $items | Should -Not -Contain 'Exchange hybrid'
            @($store['MigrationQuoteReadiness'] | Where-Object { $_.DecisionKey -in @('FLAG:ad-connect-directory-synchronization', 'FLAG:exchange-hybrid') }).Count | Should -Be 0
        }

        It 'raises both flags as Review when synchronization and hybrid are confirmed' {
            $store = @{
                TenantInfo             = New-T2TTenantInfo
                AllMailboxes           = @(New-T2TMailbox)
                AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $true
                HybridConfiguration    = New-T2THybridConfiguration -IsHybridConfigured $true -EvidenceCount 2 -HybridStatus 'Hybrid'
            }
            Update-MigrationScopeExportData -TenantStatsStore $store

            $sync = @($store['MigrationComplexityFlags'] | Where-Object { $_.Item -eq 'AD Connect directory synchronization' })
            $hybrid = @($store['MigrationComplexityFlags'] | Where-Object { $_.Item -eq 'Exchange hybrid' })
            $sync.Count | Should -Be 1
            $sync[0].Status | Should -Be 'Review'
            $hybrid.Count | Should -Be 1
            $hybrid[0].Status | Should -Be 'Review'
        }

        It 'raises Needs Data, not Configured, when neither state can be established' {
            $store = @{ AllMailboxes = @(New-T2TMailbox) }
            Update-MigrationScopeExportData -TenantStatsStore $store

            $sync = @($store['MigrationComplexityFlags'] | Where-Object { $_.Item -eq 'AD Connect directory synchronization' })
            $hybrid = @($store['MigrationComplexityFlags'] | Where-Object { $_.Item -eq 'Exchange hybrid' })
            $sync[0].Status | Should -Be 'Needs Data'
            $sync[0].Value | Should -Not -Be 'Configured'
            $hybrid[0].Status | Should -Be 'Needs Data'
            $hybrid[0].Value | Should -Not -Be 'Configured'
        }

        It 'raises Needs Data when hybrid probes were incomplete' {
            $store = @{
                TenantInfo             = New-T2TTenantInfo
                AllMailboxes           = @(New-T2TMailbox)
                AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $null
                HybridConfiguration    = New-T2THybridConfiguration -IsHybridConfigured $null -EvidenceComplete $false -HybridStatus 'Unknown' -UnansweredProbes 'MigrationEndpoint (Error)'
            }
            Update-MigrationScopeExportData -TenantStatsStore $store

            @($store['MigrationComplexityFlags'] | Where-Object { $_.Item -eq 'Exchange hybrid' })[0].Status | Should -Be 'Needs Data'
            @($store['MigrationComplexityFlags'] | ForEach-Object { $_.Item }) | Should -Not -Contain 'AD Connect directory synchronization'
        }
    }

    Context 'P0-02 and P0-03: migration readiness rows' {
        It 'does not report cloud-only when the organization read failed (production collector, mocked Graph failure)' {
            $script:tenantStatsHash = @{}
            function Get-AssessmentTenantOrganization { throw 'synthetic Graph failure (403)' }
            Get-AdConnectSyncDetails
            $store = @{ AdConnectConfiguration = $script:tenantStatsHash['AdConnectConfiguration']; Domains = @{} }
            Update-AssessmentReportTables -TenantStatsHash $store -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true

            $row = Get-T2TReadinessRow -Store $store -Item 'Directory synchronization'
            $row.Status | Should -Be 'Needs Data'
            $row.Value | Should -Not -Match 'Cloud-only'
        }

        It 'reports cloud-only as Ready only with affirmative evidence' {
            $store = @{ TenantInfo = New-T2TTenantInfo; AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $null; Domains = @{ 'contoso.example' = New-T2TDomain } }
            Update-AssessmentReportTables -TenantStatsHash $store -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true

            $row = Get-T2TReadinessRow -Store $store -Item 'Directory synchronization'
            $row.Status | Should -Be 'Ready'
            $row.Value | Should -Match 'Cloud-only'
        }

        It 'reports synchronization as Review when confirmed' {
            $store = @{ TenantInfo = New-T2TTenantInfo; AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $true; Domains = @{ 'contoso.example' = New-T2TDomain } }
            Update-AssessmentReportTables -TenantStatsHash $store -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true

            (Get-T2TReadinessRow -Store $store -Item 'Directory synchronization').Status | Should -Be 'Review'
        }

        It 'scores domain rows Needs Data when no domain was collected' {
            $store = @{ Domains = @{} }
            Update-AssessmentReportTables -TenantStatsHash $store -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true

            (Get-T2TReadinessRow -Store $store -Item 'Verified custom domains').Status | Should -Be 'Needs Data'
            (Get-T2TReadinessRow -Store $store -Item 'Mail routing').Status | Should -Be 'Needs Data'
        }

        It 'scores domain rows Ready with complete Microsoft-routed evidence' {
            $store = @{ Domains = @{ 'contoso.example' = New-T2TDomain } }
            Update-AssessmentReportTables -TenantStatsHash $store -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true

            (Get-T2TReadinessRow -Store $store -Item 'Verified custom domains').Status | Should -Be 'Ready'
            (Get-T2TReadinessRow -Store $store -Item 'Mail routing').Status | Should -Be 'Ready'
        }

        It 'scores mail routing Needs Data when MX evidence is unknown and Review when third-party routing is affirmative' {
            $unknown = @{ Domains = @{ 'contoso.example' = New-T2TDomain -Office365MailExchanger $null -MXRecords '' } }
            Update-AssessmentReportTables -TenantStatsHash $unknown -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true
            (Get-T2TReadinessRow -Store $unknown -Item 'Mail routing').Status | Should -Be 'Needs Data'

            $thirdParty = @{ Domains = @{ 'contoso.example' = New-T2TDomain -Office365MailExchanger $false -MXRecords 'mx1.gateway.example' } }
            Update-AssessmentReportTables -TenantStatsHash $thirdParty -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true
            (Get-T2TReadinessRow -Store $thirdParty -Item 'Mail routing').Status | Should -Be 'Review'
        }

        It 'scores hybrid, public folder, and mail flow rows Needs Data when nothing was collected' {
            $store = @{ Domains = @{ 'contoso.example' = New-T2TDomain } }
            Update-AssessmentReportTables -TenantStatsHash $store -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true

            (Get-T2TReadinessRow -Store $store -Item 'Hybrid or coexistence indicators').Status | Should -Be 'Needs Data'
            (Get-T2TReadinessRow -Store $store -Item 'Public folders').Status | Should -Be 'Needs Data'
            (Get-T2TReadinessRow -Store $store -Item 'Mail flow dependencies').Status | Should -Be 'Needs Data'
        }

        It 'keeps Ready for datasets that were collected and are genuinely empty' {
            $store = @{
                Domains             = @{ 'contoso.example' = New-T2TDomain }
                PublicFolderDetails = @{}
                MailFlowConnectors  = @{}
                HybridConfiguration = New-T2THybridConfiguration -IsHybridConfigured $false
                CollectorPlanResults = @(
                    (New-T2TCollectorResult -Name 'Public folders'),
                    (New-T2TCollectorResult -Name 'Mail flow rules/connectors'),
                    (New-T2TCollectorResult -Name 'Exchange hybrid configuration')
                )
            }
            Update-AssessmentReportTables -TenantStatsHash $store -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true

            (Get-T2TReadinessRow -Store $store -Item 'Hybrid or coexistence indicators').Status | Should -Be 'Ready'
            (Get-T2TReadinessRow -Store $store -Item 'Public folders').Status | Should -Be 'Ready'
            (Get-T2TReadinessRow -Store $store -Item 'Mail flow dependencies').Status | Should -Be 'Ready'
        }
    }

    Context 'P0-03: tenant-to-tenant cutover preparation and checklist' {
        It 'never scores Ready when no Exchange evidence exists, and keeps the items visible' {
            $store = @{ Users = @([pscustomobject]@{ UserPrincipalName = 'avery@contoso.example'; Mail = 'avery@contoso.example'; UserType = 'Member'; AccountEnabled = $true }) }
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store

            @($store['CutoverPrepSummary'] | Where-Object { $_.Status -eq 'Ready' } | ForEach-Object { $_.Item }) | Should -BeNullOrEmpty

            foreach ($item in @(
                    'Mail flow connectors', 'Public folders', 'Mailboxes on hold', 'Mailboxes with forwarding',
                    'Remote domains allowing auto-forwarding', 'Teams with shared channels',
                    'Oversized mailbox batches', 'Oversized archives'
                )) {
                $row = @($store['MigrationReadinessChecklist'] | Where-Object { $_.Item -eq $item })
                $row.Count | Should -Be 1 -Because "'$item' has no evidence and must stay visible"
                $row[0].Status | Should -Be 'Needs Data' -Because "'$item' has no evidence"
            }
        }

        It 'keeps Ready, and hides the row, for collected-empty connectors and public folders' {
            $store = @{
                Users               = @()
                AllMailboxes        = @(New-T2TMailbox)
                MailFlowConnectors  = @{}
                PublicFolderDetails = @{}
                RemoteDomains       = @([pscustomobject]@{ DomainName = '*'; AutoForwardEnabled = $false })
                CollectorPlanResults = @(
                    (New-T2TCollectorResult -Name 'Mail flow rules/connectors'),
                    (New-T2TCollectorResult -Name 'Public folders'),
                    (New-T2TCollectorResult -Name 'Exchange mailboxes')
                )
            }
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store

            @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Mail flow connectors' })[0].Status | Should -Be 'Ready'
            @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Public folders' })[0].Status | Should -Be 'Ready'
            @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Mailboxes with forwarding' })[0].Status | Should -Be 'Ready'
            @($store['MigrationReadinessChecklist'] | Where-Object { $_.Item -in @('Mail flow connectors', 'Public folders', 'Mailboxes with forwarding') }).Count | Should -Be 0
        }

        It 'scores Needs Data when a collector failed even though an empty container exists' {
            $store = @{
                AllMailboxes        = @(New-T2TMailbox)
                PublicFolderDetails = @{}
                CollectorPlanResults = @((New-T2TCollectorResult -Name 'Public folders' -Status 'Failed' -Message 'synthetic 403'))
            }
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store

            @($store['MigrationReadinessChecklist'] | Where-Object { $_.Item -eq 'Public folders' })[0].Status | Should -Be 'Needs Data'
        }

        It 'lists migration-relevant collectors that the profile skipped' {
            $store = @{
                AllMailboxes = @(New-T2TMailbox)
                CollectorPlanResults = @(
                    (New-T2TCollectorResult -Name 'Entra groups' -Status 'Skipped' -Message 'Entra group inventory is disabled for this profile.'),
                    (New-T2TCollectorResult -Name 'Authentication/SSO configuration' -Status 'Skipped' -Message 'Authentication and SSO collection is disabled for this profile.'),
                    (New-T2TCollectorResult -Name 'Devices' -Status 'Skipped' -Message 'Device inventory is disabled for this profile.'),
                    (New-T2TCollectorResult -Name 'Secure Score report' -Status 'Skipped' -Message 'Secure Score collection is disabled for this profile.'),
                    (New-T2TCollectorResult -Name 'Exchange mailboxes')
                )
            }
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store

            $coverageRows = @($store['MigrationReadinessChecklist'] | Where-Object { $_.Category -eq 'Evidence coverage' })
            $coverageText = ($coverageRows | ForEach-Object { $_.Item }) -join ' | '
            $coverageText | Should -Match 'Entra groups'
            $coverageText | Should -Match 'Authentication/SSO configuration'
            $coverageText | Should -Match 'Devices'
            $coverageText | Should -Not -Match 'Secure Score'
            $coverageText | Should -Not -Match 'Exchange mailboxes'
            @($coverageRows | Where-Object { $_.Status -ne 'Needs Data' }).Count | Should -Be 0
        }
    }

    Context 'P0-04: Multi-Geo' {
        BeforeEach {
            $script:tenantStatsHash = @{}
            $script:AssessmentAuthWorkloadPlan = [pscustomobject]@{ AuthenticationType = 'Certificate' }
            function script:New-T2TOrganizationPayload {
                param([hashtable]$Extra = @{})
                $payload = [ordered]@{
                    id = '00000000-0000-0000-0000-000000000001'; displayName = 'Contoso Synthetic'; country = 'United States'; countryLetterCode = 'US'
                    verifiedDomains = @(
                        [pscustomobject]@{ name = 'contoso.onmicrosoft.com'; isInitial = $true; isDefault = $false },
                        [pscustomobject]@{ name = 'contoso.example'; isInitial = $false; isDefault = $true }
                    )
                }
                foreach ($key in $Extra.Keys) { $payload[$key] = $Extra[$key] }
                [pscustomobject]$payload
            }
        }

        It 'reads the documented Graph property when Multi-Geo is enabled' {
            function Get-AssessmentTenantOrganization { New-T2TOrganizationPayload -Extra @{ isMultipleDataLocationsForServicesEnabled = $true } }
            Get-TenantOverviewInfo

            $script:tenantStatsHash['TenantInfo'].MultiGeoEnabled | Should -BeTrue
            Get-T2TQuestionnaireLine -Store @{ TenantInfo = $script:tenantStatsHash['TenantInfo'] } -Pattern 'Multi-Geo environment' | Should -Match '\| Yes'
        }

        It 'reports not enabled only for an explicit false' {
            function Get-AssessmentTenantOrganization { New-T2TOrganizationPayload -Extra @{ isMultipleDataLocationsForServicesEnabled = $false } }
            Get-TenantOverviewInfo

            $script:tenantStatsHash['TenantInfo'].MultiGeoEnabled | Should -BeFalse
            $script:tenantStatsHash['TenantInfo'].MultiGeoEnabled | Should -Not -BeNullOrEmpty
            Get-T2TQuestionnaireLine -Store @{ TenantInfo = $script:tenantStatsHash['TenantInfo'] } -Pattern 'Multi-Geo environment' | Should -Match 'not Multi-Geo enabled'
        }

        It 'keeps the state unknown when the property is absent' {
            function Get-AssessmentTenantOrganization { New-T2TOrganizationPayload }
            Get-TenantOverviewInfo

            $script:tenantStatsHash['TenantInfo'].MultiGeoEnabled | Should -BeNullOrEmpty
            $line = Get-T2TQuestionnaireLine -Store @{ TenantInfo = $script:tenantStatsHash['TenantInfo'] } -Pattern 'Multi-Geo environment'
            $line | Should -Not -Match '\| No\b'
            $line | Should -Match 'Not reported'
            $line | Should -Match 'confirm'
        }

        It 'does not claim a Multi-Geo state when the tenant overview was not collected' {
            function Get-AssessmentTenantOrganization { throw 'synthetic Graph failure' }
            Get-TenantOverviewInfo

            $line = Get-T2TQuestionnaireLine -Store @{ TenantInfo = $script:tenantStatsHash['TenantInfo'] } -Pattern 'Multi-Geo environment'
            $line | Should -Not -Match '\| No\b'
            $line | Should -Match 'Not collected'
        }
    }

    Context 'P0-02 and P0-03: questionnaire answers without evidence' {
        It 'does not assert identity, sign-in, or inventory facts that were not collected' {
            $path = Join-Path $TestDrive 'questionnaire-empty.md'
            Export-TenantToTenantQuestionnaireMarkdown -TenantStatsHash @{ TenantInfo = New-T2TTenantInfo } -TemplatePath $script:questionnaireTemplatePath -Path $path -PreparedBy 'synthetic'
            $content = Get-Content -Path $path -Raw

            $content | Should -Not -Match 'No active on-premises directory synchronization detected'
            $content | Should -Not -Match 'appears to be the primary sign-on provider'
            $content | Should -Not -Match 'No enterprise SSO applications were detected'
            $content | Should -Not -Match '\| 0 active user mailbox\(es\)\.'
            $content | Should -Not -Match '\| 0 public folder object\(s\)\.'
            $content | Should -Not -Match '\| 0 SharePoint site\(s\)\.'
        }

        It 'reports synchronization and collected counts when evidence exists' {
            $store = @{
                TenantInfo             = New-T2TTenantInfo
                AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $true
                AllMailboxes           = @(New-T2TMailbox)
                PublicFolderDetails    = @{}
                CollectorPlanResults   = @((New-T2TCollectorResult -Name 'Public folders'), (New-T2TCollectorResult -Name 'Exchange mailboxes'))
            }
            Get-T2TQuestionnaireLine -Store $store -Pattern 'Are objects synchronized from on-prem' | Should -Match '\| Yes'
            Get-T2TQuestionnaireLine -Store $store -Pattern '^\| Active user mailboxes' | Should -Match '1 active user mailbox'
            Get-T2TQuestionnaireLine -Store $store -Pattern '^\| Public folders' | Should -Match '0 public folder object'
        }
    }

    Context 'P0-05: mail exchanger classification' {
        It 'classifies <Name>' -ForEach @(
            @{ Name = 'standard Exchange Online Protection MX'; Hosts = @('contoso-example.mail.protection.outlook.com'); Answered = $true; Classification = 'Microsoft365'; IsMicrosoft = $true }
            @{ Name = 'mx.microsoft MX used with DNSSEC'; Hosts = @('contoso-example.b-v1.mx.microsoft'); Answered = $true; Classification = 'Microsoft365'; IsMicrosoft = $true }
            @{ Name = 'Microsoft MX with trailing dot and mixed case'; Hosts = @('Contoso-Example.Mail.Protection.Outlook.com.'); Answered = $true; Classification = 'Microsoft365'; IsMicrosoft = $true }
            @{ Name = 'synthetic third-party MX'; Hosts = @('mx1.gateway.example', 'mx2.gateway.example'); Answered = $true; Classification = 'ThirdParty'; IsMicrosoft = $false }
            @{ Name = 'mixed Microsoft and third-party MX'; Hosts = @('contoso-example.mail.protection.outlook.com', 'mx1.gateway.example'); Answered = $true; Classification = 'Mixed'; IsMicrosoft = $false }
            @{ Name = 'answered lookup without MX records'; Hosts = @(); Answered = $true; Classification = 'NoMxRecord'; IsMicrosoft = $null }
            @{ Name = 'failed DNS lookup'; Hosts = $null; Answered = $false; Classification = 'Unknown'; IsMicrosoft = $null }
            @{ Name = 'malformed result'; Hosts = @('not a host name', '   '); Answered = $true; Classification = 'Unknown'; IsMicrosoft = $null }
            @{ Name = 'a look-alike suffix that is not Microsoft'; Hosts = @('mail.protection.outlook.com.gateway.example'); Answered = $true; Classification = 'ThirdParty'; IsMicrosoft = $false }
        ) {
            $result = Get-ArrayaMailExchangerClassification -MxHost $Hosts -LookupAnswered $Answered
            $result.Classification | Should -Be $Classification
            $result.IsMicrosoft365 | Should -Be $IsMicrosoft
        }

        It 'applies the classification in the production domain collector with mocked DNS' {
            $script:tenantStatsHash = @{
                AllRecipients = @{ r1 = [pscustomobject]@{ PrimarySmtpAddress = 'avery@contoso.example'; EmailAddresses = 'SMTP:avery@contoso.example' } }
            }
            function Get-MgDomain {
                [pscustomobject]@{ ID = 'contoso.example'; IsVerified = $true; IsDefault = $true; AuthenticationType = 'Managed' }
                [pscustomobject]@{ ID = 'source.example'; IsVerified = $true; IsDefault = $false; AuthenticationType = 'Managed' }
                [pscustomobject]@{ ID = 'target.example'; IsVerified = $true; IsDefault = $false; AuthenticationType = 'Managed' }
                [pscustomobject]@{ ID = 'nomx.example'; IsVerified = $true; IsDefault = $false; AuthenticationType = 'Managed' }
            }
            function Get-AcceptedDomain { @() }
            function Get-RemoteDomain { @() }
            function Resolve-DnsName {
                param($Name, $Server, $Type)
                if ($Type -ne 'MX') { return $null }
                switch ($Name) {
                    'contoso.example' { [pscustomobject]@{ Name = $Name; NameExchange = 'contoso-example.b-v1.mx.microsoft'; Preference = 0 } }
                    'source.example' { [pscustomobject]@{ Name = $Name; NameExchange = 'mx1.gateway.example'; Preference = 10 } }
                    'target.example' { $null }
                    'nomx.example' { [pscustomobject]@{ Name = $Name; PrimaryServer = 'ns1.nomx.example'; NameAdministrator = 'hostmaster.nomx.example' } }
                }
            }
            Get-AllOffice365Domains

            $domains = $script:tenantStatsHash['Domains']
            $domains['contoso.example'].Office365MailExchanger | Should -BeTrue
            $domains['source.example'].Office365MailExchanger | Should -BeFalse
            $domains['source.example'].Office365MailExchanger | Should -Not -BeNullOrEmpty
            $domains['target.example'].Office365MailExchanger | Should -BeNullOrEmpty -Because 'a lookup that returned nothing is unknown, not third-party routing'
            $domains['nomx.example'].Office365MailExchanger | Should -BeNullOrEmpty -Because 'a domain without MX records has no mail routing to classify'
        }
    }

    Context 'P0-07: mailbox hold evidence' {
        It 'counts litigation, delay, and in-place hold indicators and reports retention hold separately' {
            $rows = @(
                (New-T2TMailbox -Alias 'lit' -Extra @{ LitigationHoldEnabled = $true; DelayHoldApplied = $false; RetentionHoldEnabled = $false; InPlaceHolds = '' }),
                (New-T2TMailbox -Alias 'delay' -Extra @{ LitigationHoldEnabled = $false; DelayHoldApplied = $true; RetentionHoldEnabled = $true; InPlaceHolds = '' }),
                (New-T2TMailbox -Alias 'policy' -Extra @{ LitigationHoldEnabled = $false; DelayHoldApplied = $false; RetentionHoldEnabled = $false; InPlaceHolds = 'mbx0123456789abcdef0123456789abcdef:1' }),
                (New-T2TMailbox -Alias 'excluded' -Extra @{ LitigationHoldEnabled = $false; DelayHoldApplied = $false; RetentionHoldEnabled = $false; InPlaceHolds = '-mbx0123456789abcdef0123456789abcdef' }),
                (New-T2TMailbox -Alias 'none' -Extra @{ LitigationHoldEnabled = $false; DelayHoldApplied = $false; RetentionHoldEnabled = $false; InPlaceHolds = '' })
            )
            $evidence = Get-ArrayaMailboxHoldEvidence -Mailbox $rows

            $evidence.MailboxCount | Should -Be 5
            $evidence.LitigationHoldCount | Should -Be 1
            $evidence.DelayHoldCount | Should -Be 1
            $evidence.InPlaceHoldCount | Should -Be 1
            $evidence.RetentionHoldCount | Should -Be 1
            $evidence.AnyHoldCount | Should -Be 3
            $evidence.EvaluatedSignals | Should -Contain 'LitigationHoldEnabled'
            $evidence.EvaluatedSignals | Should -Contain 'InPlaceHolds'
            $evidence.UnavailableSignals | Should -BeNullOrEmpty
        }

        It 'reports unavailable signals instead of zero when hold fields were not collected' {
            $evidence = Get-ArrayaMailboxHoldEvidence -Mailbox @((New-T2TMailbox -Alias 'bare'))

            $evidence.AnyHoldCount | Should -Be 0
            $evidence.EvaluatedSignals | Should -BeNullOrEmpty
            $evidence.UnavailableSignals | Should -Contain 'LitigationHoldEnabled'
            $evidence.UnavailableSignals | Should -Contain 'InPlaceHolds'
        }

        It 'flags mailboxes preserved by non-litigation holds in the cutover summary and checklist' {
            $store = @{
                AllMailboxes = @(
                    (New-T2TMailbox -Alias 'policy' -Extra @{ LitigationHoldEnabled = $false; DelayHoldApplied = $false; RetentionHoldEnabled = $false; InPlaceHolds = 'mbx0123456789abcdef0123456789abcdef:1' }),
                    (New-T2TMailbox -Alias 'delay' -Extra @{ LitigationHoldEnabled = $false; DelayHoldApplied = $true; RetentionHoldEnabled = $true; InPlaceHolds = 'UniH0123456789abcdef0123456789abcdef' })
                )
            }
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store

            $summary = @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Mailboxes on hold' })[0]
            $summary.Status | Should -Be 'Review'
            [int]$summary.Value | Should -Be 2
            $checklist = @($store['MigrationReadinessChecklist'] | Where-Object { $_.Item -eq 'Mailboxes on hold' })
            $checklist.Count | Should -Be 1
            $checklist[0].Value | Should -Match '2 mailbox'
        }

        It 'flags litigation hold as Review' {
            $store = @{ AllMailboxes = @((New-T2TMailbox -Alias 'lit' -Extra @{ LitigationHoldEnabled = $true; DelayHoldApplied = $false; RetentionHoldEnabled = $false; InPlaceHolds = '' })) }
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store

            @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Mailboxes on hold' })[0].Status | Should -Be 'Review'
        }

        It 'never scores Ready when no indicator is found, because organization-wide holds are not evaluated' {
            $store = @{ AllMailboxes = @((New-T2TMailbox -Alias 'none' -Extra @{ LitigationHoldEnabled = $false; DelayHoldApplied = $false; RetentionHoldEnabled = $false; InPlaceHolds = '' })) }
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store

            $summary = @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Mailboxes on hold' })[0]
            $summary.Status | Should -Be 'Needs Data'
            $summary.Notes | Should -Match 'not evaluated'
            [int]$summary.Value | Should -Be 0
        }

        It 'states which hold fields were missing when evidence is incomplete' {
            $store = @{ AllMailboxes = @((New-T2TMailbox -Alias 'bare' -Extra @{ LitigationHoldEnabled = $false })) }
            Update-TenantToTenantMigrationExportData -TenantStatsStore $store

            $summary = @($store['CutoverPrepSummary'] | Where-Object { $_.Item -eq 'Mailboxes on hold' })[0]
            $summary.Status | Should -Be 'Needs Data'
            $summary.Notes | Should -Match 'InPlaceHolds'
        }
    }

    Context 'Render-only regression: tenant-to-tenant workbook and HTML' {
        BeforeAll {
            # Export-HashTableToExcel is a module function, so its logging helpers must be
            # visible at global scope, as the existing workbook tests do.
            function global:Write-Log { param() }
            function global:Write-ProgressHelper { param() }

            $script:renderStore = @{
                TenantInfo             = New-T2TTenantInfo
                Users                  = @([pscustomobject]@{ DisplayName = 'Avery'; UserPrincipalName = 'avery@contoso.example'; Mail = 'avery@contoso.example'; UserType = 'Member'; AccountEnabled = $true })
                AllMailboxes           = @((New-T2TMailbox -Alias 'avery' -Extra @{ LitigationHoldEnabled = $false; DelayHoldApplied = $true; RetentionHoldEnabled = $false; InPlaceHolds = '' }))
                Domains                = @{
                    'contoso.example' = (New-T2TDomain)
                    'source.example'  = (New-T2TDomain -Name 'source.example' -Office365MailExchanger $null -MXRecords '')
                }
                AdConnectConfiguration = New-T2TSyncConfiguration -OnPremisesSyncEnabled $null
                HybridConfiguration    = New-T2THybridConfiguration -IsHybridConfigured $null -EvidenceComplete $false -HybridStatus 'Unknown' -UnansweredProbes 'MigrationEndpoint (Error)'
                CollectorPlanResults   = @(
                    (New-T2TCollectorResult -Name 'Exchange mailboxes'),
                    (New-T2TCollectorResult -Name 'Public folders' -Status 'Failed' -Message 'synthetic 403'),
                    (New-T2TCollectorResult -Name 'Authentication/SSO configuration' -Status 'Skipped' -Message 'Authentication and SSO collection is disabled for this profile.')
                )
            }
            Update-AssessmentReportTables -TenantStatsHash $script:renderStore -IncludeBestPracticeTables $false -IncludeMigrationReadiness $true
            Update-TenantToTenantMigrationExportData -TenantStatsStore $script:renderStore
        }

        It 'writes the cutover workbook with unchanged sheet schemas and the corrected statuses' {
            $exportPath = Join-Path $TestDrive 'p0-render.xlsx'
            Export-HashTableToExcel -hashtable $script:renderStore -ExportDetails $exportPath -WorkbookExportPolicy TenantToTenantCutover

            $worksheetNames = @(Get-ExcelSheetInfo -Path $exportPath | Select-Object -ExpandProperty Name)
            $worksheetNames | Should -Contain 'CutoverPrepSummary'
            $worksheetNames | Should -Contain 'Domains'
            $worksheetNames | Should -Not -Contain 'CollectorPlanResults'

            $prepSheet = @(Import-Excel -Path $exportPath -WorksheetName 'CutoverPrepSummary')
            $prepSheet[0].PSObject.Properties.Name | Should -Be @('Category', 'Item', 'Status', 'Value', 'Notes')
            @($prepSheet | Where-Object { $_.Item -eq 'Public folders' })[0].Status | Should -Be 'Needs Data'
            @($prepSheet | Where-Object { $_.Item -eq 'Mailboxes on hold' })[0].Status | Should -Be 'Review'

            $domainSheet = @(Import-Excel -Path $exportPath -WorksheetName 'Domains')
            @($domainSheet | Where-Object { $_.Name -eq 'source.example' })[0].ThirdPartySpamFilterReview | Should -Be 'Needs Data'
            @($domainSheet | Where-Object { $_.Name -eq 'contoso.example' })[0].ThirdPartySpamFilterReview | Should -Be 'No immediate signal'
        }

        It 'renders the cutover HTML checklist with evidence gaps visible' {
            $htmlPath = Join-Path $TestDrive 'p0-render.html'
            $result = New-TenantMigrationCutoverHtmlReport -TenantStatsHash $script:renderStore -OutputPath $htmlPath -CollectionScopePolicy 'TenantToTenantCutover'

            $result.Success | Should -BeTrue
            $html = Get-Content -Path $htmlPath -Raw
            $html | Should -Match 'Evidence not collected: Authentication/SSO configuration'
            # The fixture confirms cloud-only identity (Ready, so hidden) and leaves hybrid unknown.
            $html | Should -Match 'Hybrid state not established'
            $html | Should -Match 'Public folder inventory not collected'
            $html | Should -Not -Match 'Directory synchronization state not established'
        }
    }

    Context 'P0-06: profile contract' {
        BeforeAll {
            $script:profileNames = @('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')
            $script:runGuide = Get-Content -Path (Join-Path $script:repoRoot 'RUN.md') -Raw
            $script:readme = Get-Content -Path (Join-Path $script:repoRoot 'README.md') -Raw

            function script:Get-T2TProfilePlan {
                param([string]$ProfileName)
                $policy = Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile $ProfileName
                Resolve-AssessmentProfileCollectionPlan `
                    -OutputProfile $ProfileName `
                    -EffectiveOutputProfileLabel $ProfileName `
                    -GenerateTechnicalHtml ([bool]$policy.GenerateTechnicalHtml) `
                    -GenerateWorkbook ([bool]$policy.GenerateWorkbook) `
                    -GenerateBestPracticesHtml ([bool]$policy.GenerateBestPracticesHtml) `
                    -GenerateJson ([bool]$policy.GenerateJson) `
                    -GenerateQuestionnaire ([bool]$policy.GenerateQuestionnaire) `
                    -IsMergedOutputProfileSelection $false `
                    -ReportingMode ([string]$policy.ReportingMode) `
                    -CollectionScopePolicy ([string]$policy.CollectionScopePolicy)
            }
        }

        It 'resolves a collection plan for every profile' {
            foreach ($profileName in $script:profileNames) {
                $plan = Get-T2TProfilePlan -ProfileName $profileName
                $plan.CollectUsers | Should -BeTrue -Because "$profileName must still collect users"
                $plan.CollectExchangeMailboxes | Should -BeTrue -Because "$profileName must still collect mailboxes"
            }
        }

        It 'keeps the default profiles collecting identity, application, group, and device evidence' {
            foreach ($profileName in @('SolutionsEngineer', 'Geek', 'Machine')) {
                $plan = Get-T2TProfilePlan -ProfileName $profileName
                foreach ($flag in @('CollectEntraGroups', 'CollectAuthenticationConfiguration', 'CollectConditionalAccessPolicies', 'CollectUnifiedGroups', 'CollectDevices')) {
                    $plan[$flag] | Should -BeTrue -Because "$profileName.$flag"
                }
            }
        }

        It 'pins the current TenantToTenantMigration contract so documentation cannot drift from it silently' {
            $plan = Get-T2TProfilePlan -ProfileName 'TenantToTenantMigration'
            foreach ($flag in @('CollectEntraGroups', 'CollectAuthenticationConfiguration', 'CollectConditionalAccessPolicies', 'CollectMfaRegistrationDetails', 'CollectAdmins', 'CollectUnifiedGroups', 'CollectDevices', 'CollectGovernanceCompliancePolicies')) {
                $plan[$flag] | Should -BeFalse -Because "TenantToTenantMigration.$flag is documented as not collected; update README.md, RUN.md, and this test together when the profile is expanded"
            }
            $plan.BuildMigrationReadinessTables | Should -BeTrue

            $depth = Get-ArrayaCollectionDepthPolicy -ReportingMode 'All' -CollectionScopePolicy 'TenantToTenantCutover'
            $depth.CollectMailboxDelegatePermissions | Should -BeTrue
            $depth.CollectMailboxCalendarDelegatePermissions | Should -BeTrue
            $depth.CollectFullSharePointDetail | Should -BeFalse
            $depth.CollectUnifiedGroupMailboxStats | Should -BeFalse
        }

        It 'collects mailbox delegates only under the tenant-to-tenant scope policy' {
            (Get-ArrayaCollectionDepthPolicy -ReportingMode 'Operator' -CollectionScopePolicy 'Default').CollectMailboxDelegatePermissions | Should -BeFalse
            (Get-ArrayaCollectionDepthPolicy -ReportingMode 'Operator' -CollectionScopePolicy 'Presales').CollectMailboxDelegatePermissions | Should -BeFalse
        }

        It 'documents the same collection depth in RUN.md as the profile policy implements' {
            foreach ($profileName in $script:profileNames) {
                $policy = Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile $profileName
                $rowMatch = [regex]::Match($script:runGuide, ('(?m)^\|\s*`{0}`\s*\|\s*(?<Depth>[A-Za-z]+)\s*\|' -f [regex]::Escape($profileName)))
                $rowMatch.Success | Should -BeTrue -Because "RUN.md must list $profileName in the output profile table"
                $rowMatch.Groups['Depth'].Value | Should -Be ([string]$policy.ReportingMode) -Because "RUN.md depth for $profileName must match Get-ArrayaAssessmentOutputProfilePolicy"
            }
        }

        It 'does not describe TenantToTenantMigration as comprehensive while it skips identity and application evidence' {
            $script:readme | Should -Not -Match 'TenantToTenantMigration`\s*\|\s*Deep migration readiness and cutover planning'
            $script:runGuide | Should -Not -Match 'TenantToTenantMigration`\s*\|\s*All\s*\|\s*Deep migration readiness\.'
            $script:readme | Should -Match 't2t-migration-review'
            $script:runGuide | Should -Match 't2t-migration-review'
        }

        It 'documents the multi-profile composition limitation' {
            $script:runGuide | Should -Match '(?s)TenantToTenantMigration.{0,600}(another profile|other profile|combined)'
            $script:runGuide | Should -Match 'delegate'
        }
    }
}
