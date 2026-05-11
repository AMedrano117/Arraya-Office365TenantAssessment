Describe 'Arraya.M365.Common' {
    BeforeAll {
        $script:previousImportModuleWarningPreference = $PSDefaultParameterValues['Import-Module:WarningAction']
        $PSDefaultParameterValues['Import-Module:WarningAction'] = 'SilentlyContinue'
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        $script:exportExcelPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Public\Export-HashTableToExcel.ps1'
        $script:exportExcelSource = Get-Content -Raw -Path $script:exportExcelPath
        $script:htmlHelperPath = Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1'
        $script:htmlHelperSource = Get-Content -Raw -Path $script:htmlHelperPath
        $script:tenantQuestionnairePath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Export-TenantToTenantQuestionnaireMarkdown.ps1'
        $script:tenantQuestionnaireSource = Get-Content -Raw -Path $script:tenantQuestionnairePath
        $script:chartBundlePath = Join-Path $script:repoRoot 'src\vendor\chart.js\4.4.0\chart.umd.min.js'
        $script:scriptAnalyzerWrapperPath = Join-Path $script:repoRoot 'tools\invoke-scriptanalyzer.ps1'
        $script:scriptAnalyzerWrapperSource = Get-Content -Raw -Path $script:scriptAnalyzerWrapperPath
        $script:expectedExports = @(
            'Convert-ArrayaLegacyTenantStatsToSnapshot'
            'Convert-ArrayaObjectToArray'
            'Convert-ArrayaSnapshotToLegacyTenantStatsHash'
            'Convert-ArrayaToDate'
            'Convert-ArrayaToNumber'
            'ConvertTo-ExportFriendlyRecord'
            'ConvertTo-ExportFriendlyValue'
            'Export-ArrayaErrorReports'
            'Export-ArrayaTenantToTenantCutoverPack'
            'Export-ArrayaTenantSnapshot'
            'Export-HashTableToExcel'
            'Filter-TenantStatsHash'
            'Get-ArrayaAssessmentOutputProfilePolicy'
            'Get-ArrayaAssessmentOutputRoot'
            'Get-ArrayaObjectValue'
            'Get-ArrayaTenantSnapshotMetricSet'
            'Import-ArrayaOffice365CustomLocal'
            'Import-ArrayaTenantSnapshotContext'
            'Import-ArrayaTenantSnapshot'
            'Invoke-ArrayaCollectionStepSafe'
            'Invoke-QuietCommand'
            'New-ArrayaAssessmentOperatorSummary'
            'New-ArrayaAssessmentContext'
            'New-ArrayaTenantSnapshot'
            'Resolve-ArrayaSnapshotOutputContext'
            'Test-ArrayaTenantSnapshot'
            'Update-ArrayaTenantSnapshot'
            'Write-ArrayaAssessmentArtifactManifest'
        )
        $script:retiredPublicFiles = @(
            'src\modules\Arraya.M365.Common\Public\Clear-ArrayaAssessmentRuntimeState.ps1'
            'src\modules\Arraya.M365.Common\Public\Convert-HashToArray.ps1'
            'src\modules\Arraya.M365.Common\Public\Export-ArrayaGraphReportCsv.ps1'
            'src\modules\Arraya.M365.Common\Public\Export-ErrorReports.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaAssessmentRuntimeState.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaCollectionDepthPolicy.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaEntraGroupClassification.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaGraphAdminReportSettings.ps1'
            'src\modules\Arraya.M365.Common\Public\Get-ArrayaGraphResource.ps1'
            'src\modules\Arraya.M365.Common\Public\Invoke-ArrayaRetry.ps1'
            'src\modules\Arraya.M365.Common\Public\Set-ArrayaAssessmentRuntimeState.ps1'
            'src\modules\Arraya.M365.Common\Public\Write-ArrayaLog.ps1'
        )
    }

    AfterAll {
        if ($null -ne $script:previousImportModuleWarningPreference) {
            $PSDefaultParameterValues['Import-Module:WarningAction'] = $script:previousImportModuleWarningPreference
        }
        else {
            $null = $PSDefaultParameterValues.Remove('Import-Module:WarningAction')
        }
    }

    It 'has a manifest' {
        Test-Path $script:manifestPath | Should -BeTrue
    }

    It 'exports the explicit common surface' {
        $manifest = Import-PowerShellDataFile -Path $script:manifestPath
        $manifest.FunctionsToExport | Should -Be $script:expectedExports
        $manifest.VariablesToExport | Should -Be @()
    }

    It 'does not export retired wrapper commands' {
        $manifest = Import-PowerShellDataFile -Path $script:manifestPath
        @(
            'Get-ArrayaGraphResource'
            'Get-ArrayaGraphAdminReportSettings'
            'Export-ArrayaGraphReportCsv'
            'Get-ArrayaAssessmentRuntimeState'
            'Set-ArrayaAssessmentRuntimeState'
            'Clear-ArrayaAssessmentRuntimeState'
            'Invoke-ArrayaRetry'
            'Write-ArrayaLog'
            'Convert-HashToArray'
            'Export-ErrorReports'
        ) | ForEach-Object {
            $manifest.FunctionsToExport -contains $_ | Should -BeFalse
        }
    }

    It 'does not keep retired common wrapper files' {
        foreach ($relativePath in $script:retiredPublicFiles) {
            Test-Path (Join-Path $script:repoRoot $relativePath) | Should -BeFalse
        }
    }

    It 'embeds a bundled Chart.js asset for HTML reports instead of a public CDN' {
        Test-Path $script:chartBundlePath | Should -BeTrue
        $script:htmlHelperSource | Should -Match 'function Get-AssessmentBundledChartJsScript'
        $script:htmlHelperSource | Should -Match 'chart\.umd\.min\.js'
        $script:htmlHelperSource | Should -Not -Match 'cdn\.jsdelivr\.net'
    }

    It 'surfaces cross-tenant mailbox cutover readiness in the migration summary outputs' {
        $script:htmlHelperSource | Should -Match 'Mailbox cutover attributes'
        $script:htmlHelperSource | Should -Match 'source \.onmicrosoft routing addresses'
        $script:htmlHelperSource | Should -Match 'canonical x500:<LegacyExchangeDN>'
        $script:htmlHelperSource | Should -Match 'Additional X500 proxies should also be preserved when present'
        $script:htmlHelperSource | Should -Match 'Calendar folder permissions do not move automatically'
        $script:htmlHelperSource | Should -Match 'Full Access permissions do not move automatically'
        $script:htmlHelperSource | Should -Match 'Send As permissions do not move automatically'
        $script:htmlHelperSource | Should -Match 'Send On Behalf permissions do not move automatically'
        $script:tenantQuestionnaireSource | Should -Match 'Exchange cutover-ready mailbox data'
        $script:tenantQuestionnaireSource | Should -Match 'Calendar delegates are present'
        $script:tenantQuestionnaireSource | Should -Match 'Full Access delegates are present'
        $script:tenantQuestionnaireSource | Should -Match 'Send As delegates are present'
        $script:tenantQuestionnaireSource | Should -Match 'Send-on-Behalf delegates are present'
    }

    It 'keeps report-building phases visible with optional console substeps' {
        $script:htmlHelperSource | Should -Match 'function Write-AssessmentOptionalSubstep'
        $script:htmlHelperSource | Should -Match 'Ownership governance: building owner state lookup'
        $script:htmlHelperSource | Should -Match 'Assessment reporting: building best-practice and migration summary tables'
        $script:htmlHelperSource | Should -Match 'Assessment reporting: building migration readiness checkpoints'
        $script:htmlHelperSource | Should -Match 'Configuration summaries: normalizing tenant, mail flow, federation, and voice rollups'
    }

    It 'includes a reduced-scope T2T HTML mode with explicit not-collected guidance' {
        $script:htmlHelperSource | Should -Match 'function New-TenantMigrationCutoverHtmlReport'
        $script:htmlHelperSource | Should -Match 'function New-MigrationReadinessChecklistRows'
        $script:htmlHelperSource | Should -Match "ConvertTo-MigrationHtmlArray -Key 'MigrationReadinessChecklist'"
        $script:htmlHelperSource | Should -Match 'function Format-MigrationHumanDataSize'
        $script:htmlHelperSource | Should -Match 'Executive Summary'
        $script:htmlHelperSource | Should -Match 'Domain and Mail Flow'
        $script:htmlHelperSource | Should -Match 'Recipient and Mailbox Footprint'
        $script:htmlHelperSource | Should -Match 'Mailbox Migration Sizing'
        $script:htmlHelperSource | Should -Match 'Delegates and Forwarding'
        $script:htmlHelperSource | Should -Match 'Collaboration Footprint'
        $script:htmlHelperSource | Should -Match 'Cutover Preparation'
        $script:htmlHelperSource | Should -Match 'Migration Readiness Checklist'
        $script:htmlHelperSource | Should -Match 'Mail flow connectors'
        $script:htmlHelperSource | Should -Match 'Remote domains allowing auto-forwarding'
        $script:htmlHelperSource | Should -Match 'Mailbox routing and proxy attributes'
        $script:htmlHelperSource | Should -Match 'Mailbox delegate reapplication'
        $script:htmlHelperSource | Should -Match 'Oversized mailbox batches'
        $script:htmlHelperSource | Should -Match 'Oversized archives'
        $script:htmlHelperSource | Should -Match 'Teams with shared channels'
        $script:htmlHelperSource | Should -Match 'Collaboration sizing gaps'
        $script:htmlHelperSource | Should -Match 'Current MX'
        $script:htmlHelperSource | Should -Match 'Teams with Shared Channels'
        $script:htmlHelperSource | Should -Match 'Detailed mailbox-by-mailbox sizing remains in the workbook'
        $script:htmlHelperSource | Should -Match 'Detailed routing-gap mailbox rows remain in the workbook and cutover pack'
        $script:htmlHelperSource | Should -Match 'All accepted domains collected in this run are listed below'
        $script:htmlHelperSource | Should -Match 'This section stays summary-focused; the workbook and cutover pack hold the mailbox-level detail used for execution'
        $script:htmlHelperSource | Should -Match 'Only the migration readiness items that still need validation, decision-making, or remediation are listed below'
        $script:htmlHelperSource | Should -Match 'largest SharePoint sites, OneDrives, and Teams-connected sites surfaced in this run'
        $script:htmlHelperSource | Should -Match "Status -notin @\('Ready', 'Info'\)"
        $script:htmlHelperSource | Should -Not -Match 'CollectionStates'
        $script:htmlHelperSource | Should -Not -Match 'Mailbox Cutover Inventory'
        $script:htmlHelperSource | Should -Match 'if \(-not \$isTenantToTenantCutover\) \{[\s\S]*?Add-MigrationRow -Category ''Identity'' -Item ''Guests and external identities'''
        $script:htmlHelperSource | Should -Match 'if \(-not \$isTenantToTenantCutover\) \{[\s\S]*?Add-MigrationRow -Category ''Collaboration'' -Item ''Teams workload data'''
        $script:htmlHelperSource | Should -Match 'if \(-not \$isTenantToTenantCutover\) \{[\s\S]*?Add-MigrationRow -Category ''Security'' -Item ''Conditional Access and MFA'''
        $script:htmlHelperSource | Should -Match 'if \(-not \$isTenantToTenantCutover\) \{[\s\S]*?Add-MigrationRow -Category ''External Access'' -Item ''Cross-tenant and B2B settings'''
        $script:htmlHelperSource | Should -Match 'if \(-not \$isTenantToTenantCutover\) \{[\s\S]*?Add-MigrationRow -Category ''Teams Voice'' -Item ''Voice workload readiness'''
    }

    It 'synthesizes the expanded T2T migration readiness checklist during HTML replay when the dedicated dataset is absent' {
        . $script:htmlHelperPath

        $outputPath = Join-Path $TestDrive 't2t-cutover.html'
        $tenantStats = @{
            TenantInfo = [pscustomobject]@{
                DisplayName   = 'Contoso'
                TenantId      = '11111111-1111-1111-1111-111111111111'
                DefaultDomain = 'contoso.com'
            }
            MigrationExecutiveSummary = @(
                [pscustomobject]@{ Section = 'Tenant'; Metric = 'Enabled member users'; Value = 10; Notes = 'Wave planning population.' },
                [pscustomobject]@{ Section = 'Messaging'; Metric = 'Recipients'; Value = 20; Notes = 'Recipient scope.' },
                [pscustomobject]@{ Section = 'Messaging'; Metric = 'Mailboxes'; Value = 2; Notes = 'Mailbox scope.' },
                [pscustomobject]@{ Section = 'Messaging'; Metric = 'Mailboxes with delegate dependencies'; Value = 2; Notes = 'Delegate cutover work.' },
                [pscustomobject]@{ Section = 'Messaging'; Metric = 'Grand total data to migrate (GB)'; Value = 250; Notes = 'Total data.' }
            )
            RecipientDomainSummary = @(
                [pscustomobject]@{ Domain = 'contoso.com'; RecipientCount = 20; UserMailboxCount = 2; SharedMailboxCount = 0; GroupRecipientCount = 0; HiddenFromAddressListsCount = 0 }
            )
            MailboxMigrationSummary = @(
                [pscustomobject]@{ RecipientTypeDetails = 'UserMailbox'; MailboxCount = 2; ActiveMailboxCount = 2; InactiveMailboxCount = 0; ArchiveEnabledCount = 1; ForwardingCount = 1; MailboxesWithDelegateDependencies = 2; TotalDataToMigrateGB = 250 }
            )
            BitTitanLicenseSummary = @(
                [pscustomobject]@{ Section = 'Counts'; Metric = 'Inactive mailboxes'; Value = 0; Notes = 'None inactive.' },
                [pscustomobject]@{ Section = 'Counts'; Metric = 'Archive-enabled mailboxes'; Value = 1; Notes = 'Archive enabled.' },
                [pscustomobject]@{ Section = 'Thresholds'; Metric = 'Mailboxes over 50 GB'; Value = 1; Notes = 'Primary content only.' },
                [pscustomobject]@{ Section = 'Thresholds'; Metric = 'Mailboxes over 100 GB'; Value = 0; Notes = 'Primary content only.' },
                [pscustomobject]@{ Section = 'Thresholds'; Metric = 'Mailboxes over 50 GB including deleted items'; Value = 2; Notes = 'Primary plus deleted items.' },
                [pscustomobject]@{ Section = 'Thresholds'; Metric = 'Archives over 100 GB'; Value = 1; Notes = 'Archive threshold.' },
                [pscustomobject]@{ Section = 'Data'; Metric = 'Grand total data to migrate (GB)'; Value = 250; Notes = 'Total data.' },
                [pscustomobject]@{ Section = 'Licensing'; Metric = 'MigrationWiz-Mailbox'; Value = 0; Notes = 'License summary.' },
                [pscustomobject]@{ Section = 'Licensing'; Metric = 'MigrationWiz-Mailbox x2'; Value = 1; Notes = 'License summary.' },
                [pscustomobject]@{ Section = 'Licensing'; Metric = 'User Migration Bundle'; Value = 1; Notes = 'License summary.' }
            )
            DelegateSummary = @(
                [pscustomobject]@{ PermissionType = 'FullAccess'; AffectedMailboxCount = 1; AssignmentCount = 1 }
            )
            CollaborationSummary = @(
                [pscustomobject]@{ Workload = 'Teams'; TotalCount = 1; TotalStorageGB = 0.5; UnknownStorageCount = 1; LargestObjectName = 'Ops Team'; LargestObjectSizeGB = 0.5; Notes = 'Some size fields were not surfaced in current source.' },
                [pscustomobject]@{ Workload = 'SharePoint'; TotalCount = 1; TotalStorageGB = 4.5; UnknownStorageCount = 2; LargestObjectName = 'Projects'; LargestObjectSizeGB = 4.5; Notes = 'Some size fields were not surfaced in current source.' },
                [pscustomobject]@{ Workload = 'OneDrive'; TotalCount = 1; TotalStorageGB = 8.0; UnknownStorageCount = 0; LargestObjectName = 'User OneDrive'; LargestObjectSizeGB = 8.0; Notes = $null }
            )
            CutoverPrepSummary = @(
                [pscustomobject]@{ Category = 'Routing'; Item = 'Mailboxes missing .onmicrosoft alias'; Status = 'Review'; Value = 1; Notes = 'Routing gap exists.' }
            )
            MigrationReadiness = @(
                [pscustomobject]@{ Category = 'Domains'; Item = 'Verified custom domains'; Status = 'Blocker'; Value = '1 verified / 1 unverified'; Notes = 'Validate domains.'; MigrationAction = 'Confirm accepted domains.'; SourceWorksheet = 'Domains' },
                [pscustomobject]@{ Category = 'Domains'; Item = 'Mail routing'; Status = 'Review'; Value = '1 domain(s) with non-Microsoft 365 MX'; Notes = 'Review MX routing.'; MigrationAction = 'Document MX path.'; SourceWorksheet = 'Domains' },
                [pscustomobject]@{ Category = 'Hybrid'; Item = 'Hybrid or coexistence indicators'; Status = 'Review'; Value = 'Hybrid signals detected; migration endpoints=0'; Notes = 'Hybrid impacts routing.'; MigrationAction = 'Validate hybrid scope.'; SourceWorksheet = 'HybridConfiguration' },
                [pscustomobject]@{ Category = 'Identity'; Item = 'Directory synchronization'; Status = 'Review'; Value = 'On-prem sync enabled'; Notes = 'DirSync affects authority.'; MigrationAction = 'Plan sync posture.'; SourceWorksheet = 'AdConnectConfiguration' },
                [pscustomobject]@{ Category = 'Messaging'; Item = 'Mail flow dependencies'; Status = 'Review'; Value = '1 connector(s), 1 remote domain(s)'; Notes = 'Legacy broad row.'; MigrationAction = 'Legacy broad action.'; SourceWorksheet = 'MailFlowConnectors' },
                [pscustomobject]@{ Category = 'Messaging'; Item = 'Mailbox cutover attributes'; Status = 'Review'; Value = 'Legacy broad row.'; Notes = 'Legacy broad row.'; MigrationAction = 'Legacy broad action.'; SourceWorksheet = 'AllMailboxes' },
                [pscustomobject]@{ Category = 'Messaging'; Item = 'Mailbox inventory'; Status = 'Info'; Value = '2 mailbox(es)'; Notes = 'Inventory row.'; MigrationAction = 'Use workbook.'; SourceWorksheet = 'AllMailboxes' },
                [pscustomobject]@{ Category = 'Licensing'; Item = 'Target licensing readiness'; Status = 'Info'; Value = '50% utilized'; Notes = 'License row.'; MigrationAction = 'Review licensing.'; SourceWorksheet = 'LicenseSKUs' }
            )
            Domains = @(
                [pscustomobject]@{ Name = 'contoso.com'; MXRecords = 'mx.contoso.com'; AuthenticationType = 'Managed'; Office365MailExchanger = $false; ThirdPartySpamFilterReview = 'Review'; HybridRoutingReview = 'Review'; RecipientCount = 20 },
                [pscustomobject]@{ Name = 'contoso.mail.onmicrosoft.com'; MXRecords = 'contoso.mail.protection.outlook.com'; AuthenticationType = 'Managed'; Office365MailExchanger = $true; ThirdPartySpamFilterReview = 'No immediate signal'; HybridRoutingReview = 'No immediate signal'; RecipientCount = 0 }
            )
            AllMailboxes = @(
                [pscustomobject]@{
                    DisplayName               = 'User One'
                    PrimarySmtpAddress        = 'user1@contoso.com'
                    UserPrincipalName         = 'user1@contoso.com'
                    RecipientTypeDetails      = 'UserMailbox'
                    OnMicrosoftAlias          = $null
                    LegacyExchangeDnX500      = 'x500:/o=Contoso/ou=Exchange/cn=Recipients/cn=user1'
                    X500AddressCount          = 0
                    X400AddressCount          = 0
                    FullAccessDelegateCount   = 1
                    FullAccessDelegateState   = 'Collected'
                    SendAsDelegateCount       = $null
                    SendAsDelegateState       = 'NotCollected'
                    GrantSendOnBehalfToCount  = 0
                    CalendarDelegateCount     = 1
                    CalendarDelegateState     = 'Collected'
                    ForwardingSmtpAddress     = 'forward@external.com'
                    ForwardingAddress         = $null
                    LitigationHoldEnabled     = $true
                },
                [pscustomobject]@{
                    DisplayName               = 'User Two'
                    PrimarySmtpAddress        = 'user2@contoso.com'
                    UserPrincipalName         = 'user2@contoso.com'
                    RecipientTypeDetails      = 'UserMailbox'
                    OnMicrosoftAlias          = 'user2@contoso.onmicrosoft.com'
                    LegacyExchangeDnX500      = $null
                    X500AddressCount          = 2
                    X400AddressCount          = 1
                    FullAccessDelegateCount   = 0
                    FullAccessDelegateState   = 'Collected'
                    SendAsDelegateCount       = 1
                    SendAsDelegateState       = 'Collected'
                    GrantSendOnBehalfToCount  = 1
                    CalendarDelegateCount     = 0
                    CalendarDelegateState     = 'Collected'
                    ForwardingSmtpAddress     = $null
                    ForwardingAddress         = $null
                    LitigationHoldEnabled     = $false
                }
            )
            MailboxFullDetails = @()
            RemoteDomains = @(
                [pscustomobject]@{ Name = 'partner.com'; AutoForwardEnabled = $true }
            )
            SMTPRelayServiceAccounts = @(
                [pscustomobject]@{ UserPrincipalName = 'relay@contoso.com' }
            )
            PublicFolderDetails = @(
                [pscustomobject]@{ Name = 'PF Root' }
            )
            AllTeams = @(
                [pscustomobject]@{ DisplayName = 'Ops Team'; SharePointSiteUrl = 'https://contoso.sharepoint.com/sites/ops'; SharedChannelCount = 2; SharedChannels = 'Vendor;Partner'; 'SiteSize-GB' = 0.5 }
            )
            SharePoint = @(
                [pscustomobject]@{ Title = 'Projects'; Owner = 'owner@contoso.com'; Url = 'https://contoso.sharepoint.com/sites/projects'; StorageUsedGB = 4.5 }
            )
            OneDrive = @(
                [pscustomobject]@{ Title = 'User OneDrive'; Owner = 'user1@contoso.com'; Url = 'https://contoso-my.sharepoint.com/personal/user1'; StorageUsedGB = 8.0 }
            )
            MailFlowConnectors = @(
                [pscustomobject]@{ Name = 'Inbound Connector' }
            )
        }

        $result = New-TenantMigrationCutoverHtmlReport -TenantStatsHash $tenantStats -OutputPath $outputPath

        $result.Success | Should -BeTrue
        Test-Path $outputPath | Should -BeTrue

        $html = Get-Content -Raw -Path $outputPath
        $html | Should -Match 'Migration Readiness Checklist'
        $html | Should -Match 'Mail flow connectors'
        $html | Should -Match 'Remote domains allowing auto-forwarding'
        $html | Should -Match 'Mailbox routing and proxy attributes'
        $html | Should -Match 'Mailbox delegate reapplication'
        $html | Should -Match 'Oversized mailbox batches'
        $html | Should -Match 'Oversized archives'
        $html | Should -Match 'Teams with shared channels'
        $html | Should -Match 'Collaboration sizing gaps'
        $html | Should -Not -Match '>Mail flow dependencies<'
        $html | Should -Not -Match '>Mailbox cutover attributes<'
        $html | Should -Not -Match '>Mailbox inventory<'
        $html | Should -Not -Match '>Target licensing readiness<'
    }

    It 'runs ScriptAnalyzer per target, excludes vendor code by default, and fails hard on the configured severity threshold' {
        $script:scriptAnalyzerWrapperSource | Should -Match 'Set-StrictMode -Version Latest'
        $script:scriptAnalyzerWrapperSource | Should -Match '\[ValidateSet\(''Error'', ''Warning'', ''Information''\)\]'
        $script:scriptAnalyzerWrapperSource | Should -Match 'foreach \(\$target in \$targets\)'
        $script:scriptAnalyzerWrapperSource | Should -Not -Match 'Invoke-ScriptAnalyzer -Path \$targets -Recurse'
        $script:scriptAnalyzerWrapperSource | Should -Match 'PSScriptAnalyzer found no valid targets to scan'
        $script:scriptAnalyzerWrapperSource | Should -Match 'ExcludePathPrefix = @\(''./src/vendor''\)'
        $script:scriptAnalyzerWrapperSource | Should -Match 'PSScriptAnalyzer found \$\(\$blockingIssues.Count\) issue\(s\) at severity'
        $script:scriptAnalyzerWrapperSource | Should -Match 'Write-Output ''PSScriptAnalyzer passed\.'''
    }

    It 'writes error exports into a Debugging folder beside the base artifact' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $exportFileLocation = Join-Path $TestDrive 'Tenant Discovery Report-SolutionsEngineer.xlsx'
        $errorSummary = Export-ArrayaErrorReports -ExportFileLocation $exportFileLocation -ErrorData @(
            [pscustomobject]@{
                Message = 'Example failure'
                Step    = 'UnitTest'
            }
        )

        $expectedDirectory = Join-Path (Split-Path -Path $exportFileLocation -Parent) 'Debugging'
        $errorSummary.FolderPath | Should -Be $expectedDirectory
        Split-Path -Path $errorSummary.JsonPath -Parent | Should -Be $expectedDirectory
        Split-Path -Path $errorSummary.LogPath -Parent | Should -Be $expectedDirectory
        Split-Path -Path $errorSummary.CsvPath -Parent | Should -Be $expectedDirectory
        Test-Path (Join-Path $expectedDirectory 'Tenant Discovery Report-SolutionsEngineer Error Reporting') | Should -BeFalse
    }

    It 'writes a tenant-prefixed run manifest when the workbook uses the Tenant Details suffix' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $baseExportPath = Join-Path $TestDrive 'Contoso Ltd - Tenant Details.xlsx'
        Set-Content -Path $baseExportPath -Value 'placeholder' -Encoding UTF8

        $manifestPath = Write-ArrayaAssessmentArtifactManifest `
            -BaseExportPath $baseExportPath `
            -Artifacts @{ Workbook = $baseExportPath } `
            -OutputProfileLabel 'SolutionsEngineer' `
            -ReportingMode 'Operator' `
            -CollectionOnly $false `
            -ExportOnly $false

        Split-Path -Path $manifestPath -Leaf | Should -Be 'Contoso Ltd-Run.manifest.json'
        Test-Path -Path $manifestPath | Should -BeTrue

        $manifest = Get-Content -Raw -Path $manifestPath | ConvertFrom-Json -Depth 10
        $manifest.OperatorSummary | Should -Not -BeNullOrEmpty
        $manifest.OperatorSummary.PrimaryDeliverables.Type | Should -Contain 'Workbook'
        $manifest.OperatorSummary.RecommendedStart | Should -Contain 'Workbook'
        $manifest.OperatorSummary.OperatorNote | Should -Match 'Start with PrimaryDeliverables'
    }

    It 'calculates shared snapshot metrics for improvement and comparison workflows' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $metrics = Get-ArrayaTenantSnapshotMetricSet `
            -SecureScoreRows @([pscustomobject]@{
                CreatedDateTime = '2024-01-01T00:00:00Z'
                CurrentScore    = 35
                MaxScore        = 100
            }) `
            -ConditionalAccessRows @(
                [pscustomobject]@{ State = 'enabled' },
                [pscustomobject]@{ State = 'disabled' }
            ) `
            -AdminRows @(
                [pscustomobject]@{ Role = 'Global Administrator' },
                [pscustomobject]@{ Role = 'Exchange Administrator' }
            ) `
            -DomainRows @(
                [pscustomobject]@{ IsVerified = $true },
                [pscustomobject]@{ IsVerified = $false }
            ) `
            -LicenseRows @(
                [pscustomobject]@{
                    SkuPartNumber = 'ENTERPRISEPACK'
                    ConsumedUnits = 98
                    ActiveUnits   = 100
                }
            ) `
            -DeviceRows @(
                [pscustomobject]@{ ApproximateLastSignInDateTime = (Get-Date).AddDays(-60).ToString('o') },
                [pscustomobject]@{ ApproximateLastSignInDateTime = (Get-Date).AddDays(-5).ToString('o') }
            ) `
            -StaleDeviceDays 30

        $metrics.SecureScorePercent | Should -Be 35
        $metrics.ConditionalAccessPolicyCount | Should -Be 2
        $metrics.EnabledConditionalAccessCount | Should -Be 1
        $metrics.GlobalAdminCount | Should -Be 1
        $metrics.UnverifiedDomainCount | Should -Be 1
        $metrics.MaxLicenseUtilizationPercent | Should -Be 98
        $metrics.HighUtilizationSkus | Should -Be @('ENTERPRISEPACK (98%)')
        $metrics.DeviceCount | Should -Be 2
        $metrics.StaleDeviceCount | Should -Be 1
        $metrics.StaleDevicePercent | Should -Be 50
    }

    It 'imports snapshot context and resolves default snapshot output paths' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $snapshot = New-ArrayaTenantSnapshot -Data @{
            Tenant = @{
                Domains = @(
                    [pscustomobject]@{
                        Id         = 'contoso.com'
                        IsVerified = $true
                    }
                )
            }
        }
        $snapshotPath = Join-Path $TestDrive 'tenant-snapshot.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $context = Import-ArrayaTenantSnapshotContext -Path $snapshotPath -Purpose Export
        $outputContext = Resolve-ArrayaSnapshotOutputContext -PrimaryInputPath $snapshotPath

        $context.Path | Should -Be (Resolve-Path $snapshotPath).Path
        $context.GeneratedAt | Should -Not -BeNullOrEmpty
        $outputContext.OutputFolder | Should -Be (Resolve-Path $TestDrive).Path
        $outputContext.OutputPrefix | Should -Be 'tenant-snapshot'
    }

    It 'streams tenant snapshot JSON directly to disk instead of materializing a giant JSON string' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $exportSnapshotSource = Get-Content -Raw -Path (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Public\Export-ArrayaTenantSnapshot.ps1')
        $exportSnapshotSource | Should -Match 'System\.Text\.Json\.Utf8JsonWriter'
        $exportSnapshotSource | Should -Match 'System\.IO\.File\]::Create'
        $exportSnapshotSource | Should -Not -Match 'JsonSerializer\]::Serialize'
        $exportSnapshotSource | Should -Not -Match '\$jsonFriendly ='
        $exportSnapshotSource | Should -Not -Match 'WriteAllText'

        $selfReference = [ordered]@{ Name = 'self' }
        $selfReference['Self'] = $selfReference
        $snapshot = New-ArrayaTenantSnapshot -Data @{
            Tenant = @{
                Domains = @(
                    [pscustomobject]@{
                        Id         = 'contoso.com'
                        IsVerified = $true
                    }
                )
            }
            Other = @{
                Circular = $selfReference
                Secure   = ConvertTo-SecureString -String 'secret' -AsPlainText -Force
            }
        }

        $snapshotPath = Join-Path $TestDrive 'streamed-snapshot.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        Test-Path -Path $snapshotPath | Should -BeTrue
        $parsed = Get-Content -Path $snapshotPath -Raw | ConvertFrom-Json -Depth 20
        $parsed.SchemaVersion | Should -Be 2
        $parsed.Data.Tenant.Domains[0].Id | Should -Be 'contoso.com'
        $parsed.Data.Other.Circular.Self | Should -Be '[CircularReference]'
        $parsed.Data.Other.Secure | Should -Be '[SecureString]'
    }

    It 'normalizes exported assessment snapshot prefixes for shorter improve outputs' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $snapshotPath = Join-Path $TestDrive 'Contoso Tenant Discovery Report-SolutionsEngineer_20260403_101500-Snapshot.json'
        Set-Content -Path $snapshotPath -Value '{}' -Encoding UTF8

        $outputContext = Resolve-ArrayaSnapshotOutputContext -PrimaryInputPath $snapshotPath
        $outputContext.OutputPrefix | Should -Be 'Contoso-SE_20260403_101500'
    }

    It 'resolves assessment snapshot JSON from a manifest artifact declaration' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $snapshot = New-ArrayaTenantSnapshot -Data @{
            Tenant = @{
                Domains = @(
                    [pscustomobject]@{
                        Id         = 'contoso.com'
                        IsVerified = $true
                    }
                )
            }
        }

        $supportFolder = Join-Path $TestDrive 'Support'
        $null = New-Item -ItemType Directory -Path $supportFolder -Force
        $snapshotPath = Join-Path $supportFolder 'tenant-AssessmentSnapshot.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $manifestPath = Join-Path $supportFolder 'tenant.manifest.json'
        $manifestContent = [ordered]@{ Artifacts = @([ordered]@{
            Type      = 'Assessment Snapshot JSON'
            Path      = $snapshotPath
            Exists    = $true
            SizeBytes = (Get-Item -Path $snapshotPath).Length
        }) } | ConvertTo-Json -Depth 5
        Set-Content -Path $manifestPath -Value $manifestContent -Encoding UTF8

        $context = Import-ArrayaTenantSnapshotContext -Path $manifestPath -Purpose ImprovementPlan
        $context.Path | Should -Be (Resolve-Path $snapshotPath).Path
    }

    It 'maps output profiles to the updated reporting modes' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile Presales).ReportingMode | Should -Be 'Minimum'
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile ExecutiveLevel).ReportingMode | Should -Be 'Minimum'
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile SolutionsEngineer).ReportingMode | Should -Be 'Operator'
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile SolutionsEngineer).GenerateWorkbook | Should -BeTrue
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile SolutionsEngineer).GenerateJson | Should -BeTrue
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile Machine).ReportingMode | Should -Be 'Automation'
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile Geek).ReportingMode | Should -Be 'Geek'
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile TenantToTenantMigration).ReportingMode | Should -Be 'All'
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile TenantToTenantMigration).GenerateWorkbook | Should -BeTrue
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile TenantToTenantMigration).GenerateJson | Should -BeTrue
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile TenantToTenantMigration).GenerateQuestionnaire | Should -BeFalse
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile TenantToTenantMigration).WorkbookExportPolicy | Should -Be 'TenantToTenantCutover'
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile TenantToTenantMigration).TechnicalHtmlPolicy | Should -Be 'TenantToTenantCutover'
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile TenantToTenantMigration).GenerateMigrationPack | Should -BeTrue
        (Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile TenantToTenantMigration).CollectionScopePolicy | Should -Be 'TenantToTenantCutover'
    }

    It 'normalizes Exchange mailbox-size wrapper values without throwing' {
        $htmlHelperPath = Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1'
        . $htmlHelperPath

        $wrappedSize = [pscustomobject]@{
            IsUnlimited = $false
            Value       = '33.47 KB (34,275 bytes)'
        }

        { Convert-MailboxSizeToGB -SizeValue $wrappedSize } | Should -Not -Throw
        (Convert-MailboxSizeToGB -SizeValue $wrappedSize) | Should -BeGreaterOrEqual 0
    }

    It 'maps governance and password lifecycle signals into the snapshot domains' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $legacyTenantStats = @{
            AllMailboxes = @(
                [pscustomobject]@{
                    DisplayName                       = 'Relay Mailbox'
                    UserPrincipalName                 = 'relay@contoso.com'
                    PrimarySmtpAddress                = 'relay@contoso.com'
                    SmtpClientAuthenticationDisabled  = $false
                    RetentionPolicy                   = 'Finance Hold'
                    LitigationHoldEnabled             = $true
                    RetentionHoldEnabled              = $false
                    DelayHoldApplied                  = $false
                }
            )
            EmailActivityTopSenders = @(
                [pscustomobject]@{
                    UserPrincipalName = 'relay@contoso.com'
                    SendCount         = 42
                    LastActivityDate  = '2026-01-01'
                }
            )
            DlpPolicies = @(
                [pscustomobject]@{
                    PolicyName = 'Credit Card DLP'
                }
            )
            AdConnectConfiguration = @{
                Summary = [pscustomobject]@{
                    PasswordWritebackEnabled        = $true
                    PassThroughAuthenticationEnabled = $false
                    SelfServicePasswordResetEnabled = $true
                    OnPremisesSyncEnabled           = $true
                    OnPremisesLastSyncDateTime      = '2026-01-01T00:00:00Z'
                }
            }
        }

        $snapshot = Convert-ArrayaLegacyTenantStatsToSnapshot -TenantStatsHash $legacyTenantStats

        $snapshot.Data.Governance.RetentionPolicies.Keys.Count | Should -Be 1
        $snapshot.Data.Governance.DlpPolicies.Count | Should -Be 1
        $snapshot.Data.Governance.PasswordLifecycleSummary.PasswordWritebackEnabled | Should -BeTrue
        $snapshot.Data.Governance.PasswordLifecycleSummary.SelfServicePasswordResetEnabled | Should -BeTrue
        $snapshot.Data.Security.SMTPRelayServiceAccounts.Keys.Count | Should -Be 1
    }

    It 'maps external sharing and guest access signals into the snapshot domains' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        $legacyTenantStats = @{
            SharePointSharingSummary = @{
                Summary = [pscustomobject]@{
                    CollectionSource                      = 'Microsoft Graph SharePoint tenant settings'
                    TenantSharingCapability               = 'ExternalUserAndGuestSharing'
                    OneDriveSharingCapability             = 'ExternalUserSharingOnly'
                    DeletedUserPersonalSiteRetentionPeriodInDays = 30
                    IsLegacyAuthProtocolsEnabled          = $true
                }
            }
            ExternalSharingSummary = @{
                Summary = [pscustomobject]@{
                    TenantSharingCapability      = 'ExternalUserAndGuestSharing'
                    DefaultSharingLinkType       = 'AnonymousAccess'
                    SharingDomainRestrictionMode = 'allowList'
                    SiteOverrideCount            = 1
                }
            }
            ExternalSharingSiteOverrides = @(
                [pscustomobject]@{
                    Title                  = 'Projects'
                    Url                    = 'https://contoso.sharepoint.com/sites/projects'
                    SharingCapability      = 'ExistingExternalUserSharingOnly'
                    DefaultSharingLinkType = 'SpecificPeople'
                    DefaultLinkPermission  = 'View'
                    OverrideReason         = 'Sharing capability differs from tenant setting'
                }
            )
            ExternalExposureFindings = @(
                [pscustomobject]@{
                    Workload         = 'SharePoint'
                    AssetType        = 'Site'
                    Title            = 'Projects'
                    UrlOrIdentifier  = 'https://contoso.sharepoint.com/sites/projects'
                    ExposureCategory = 'Stale externally shared content'
                    TenantBaseline   = 'SharingCapability=ExternalUserAndGuestSharing'
                    ObservedSetting  = 'SharingCapability=ExistingExternalUserSharingOnly'
                    OwnerSignal      = 'Owner=siteowner@contoso.com'
                    GuestSignal      = 'Not applicable'
                    ActivitySignal   = 'LastContentModifiedDate=2025-01-01'
                    StaleSignal      = 'Yes'
                    GapReason        = 'Site supports external sharing and last content activity is older than the 180-day stale threshold.'
                    ReviewPriority   = 'High'
                }
            )
            ExternalIdentityRestrictions = @{
                Summary = [pscustomobject]@{
                    AllowInvitesFrom        = 'adminsAndGuestInviters'
                    GuestUserRoleLabel      = 'Guest users have limited access to directory objects (default)'
                    CrossTenantPartnerCount = 2
                    DefaultInboundMfaTrust  = $true
                }
            }
            GuestAccessConfiguration = @{
                Summary = [pscustomobject]@{
                    GuestInvitationControl       = 'adminsAndGuestInviters'
                    GuestUserRoleLabel           = 'Guest users have limited access to directory objects (default)'
                    ConditionalAccessGuestCoverage = $true
                }
            }
            MfaEnrollmentSummary = [pscustomobject]@{
                RegistrationPercent        = 55
                RegisteredMethodBreakdown = 'Microsoft Authenticator=4; SMS / phone=3'
                WeakMethodBreakdown       = 'SMS / phone=3'
                UsersWithWeakMethodsOnly  = 2
            }
            MfaEnforcementSummary = [pscustomobject]@{
                EnabledPoliciesRequiringMfa    = 2
                UserCoveragePercent            = 78.5
                UsersCoveredByEnabledMfaPolicies = 11
                ReportOnlyPoliciesRequiringMfa = 1
                EnforcementState               = 'MFA enforcement is active through enabled Conditional Access policies.'
            }
            MfaEnforcementGapUsers = @(
                [pscustomobject]@{
                    DisplayName        = 'Uncovered Member'
                    UserPrincipalName  = 'uncovered.member@contoso.com'
                    UserType           = 'Member'
                    DirectoryObjectId  = 'user-001'
                    GapCategory        = 'Outside enabled MFA CA include scope'
                    GapReason          = 'User is outside the include scope of all enabled Conditional Access policies that currently require MFA.'
                    RelatedPolicies    = ''
                    AccountEnabled     = $true
                    LastSignInDateTime = '2026-01-01T00:00:00Z'
                }
            )
            MfaEnforcementScopeReview = @(
                [pscustomobject]@{
                    PolicyName          = 'Baseline MFA'
                    ScopeType           = 'Exclude'
                    ObjectType          = 'Group'
                    DisplayName         = 'Break Glass Exclusions'
                    Identifier          = 'group-001'
                    AffectedEnabledUsers = 2
                    DirectoryMemberCount = 2
                    Notes               = 'Members of this group are excluded from the enabled MFA enforcement policy scope.'
                }
            )
        }

        $snapshot = Convert-ArrayaLegacyTenantStatsToSnapshot -TenantStatsHash $legacyTenantStats

        $snapshot.Data.Collaboration.SharePointSharingSummary.Summary.OneDriveSharingCapability | Should -Be 'ExternalUserSharingOnly'
        $snapshot.Data.Collaboration.SharePointSharingSummary.Summary.DeletedUserPersonalSiteRetentionPeriodInDays | Should -Be 30
        $snapshot.Data.Tenant.ExternalSharingSummary.Summary.TenantSharingCapability | Should -Be 'ExternalUserAndGuestSharing'
        $snapshot.Data.Tenant.ExternalSharingSiteOverrides.Count | Should -Be 1
        $snapshot.Data.Tenant.ExternalExposureFindings.Count | Should -Be 1
        $snapshot.Data.Identity.ExternalIdentityRestrictions.Summary.AllowInvitesFrom | Should -Be 'adminsAndGuestInviters'
        $snapshot.Data.Identity.ExternalIdentityRestrictions.Summary.GuestUserRoleLabel | Should -Be 'Guest users have limited access to directory objects (default)'
        $snapshot.Data.Identity.GuestAccessConfiguration.Summary.GuestInvitationControl | Should -Be 'adminsAndGuestInviters'
        $snapshot.Data.Identity.GuestAccessConfiguration.Summary.GuestUserRoleLabel | Should -Be 'Guest users have limited access to directory objects (default)'
        $snapshot.Data.Identity.MfaEnrollmentSummary.UsersWithWeakMethodsOnly | Should -Be 2
        $snapshot.Data.Identity.MfaEnforcementSummary.EnabledPoliciesRequiringMfa | Should -Be 2
        $snapshot.Data.Identity.MfaEnforcementSummary.UsersCoveredByEnabledMfaPolicies | Should -Be 11
        $snapshot.Data.Identity.MfaEnforcementSummary.UserCoveragePercent | Should -Be 78.5
        $snapshot.Data.Identity.MfaEnforcementGapUsers.Count | Should -Be 1
        $snapshot.Data.Identity.MfaEnforcementGapUsers[0].GapCategory | Should -Be 'Outside enabled MFA CA include scope'
        $snapshot.Data.Identity.MfaEnforcementScopeReview.Count | Should -Be 1
        $snapshot.Data.Identity.MfaEnforcementScopeReview[0].DisplayName | Should -Be 'Break Glass Exclusions'
    }

    It 'exports the external exposure worksheets in stable order' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        if (-not (Get-Command -Name Get-ExcelSheetInfo -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because 'ImportExcel worksheet inspection is not available in this environment.'
            return
        }

        function global:Write-Log { param() }
        function global:Write-ProgressHelper { param() }

        $exportPath = Join-Path $TestDrive 'external-exposure.xlsx'
        $tenantStats = @{
            TenantInfo = [pscustomobject]@{
                DisplayName = 'Contoso'
            }
            LicenseSKUs = @(
                [pscustomobject]@{
                    SkuPartNumber = 'ENTERPRISEPACK'
                    ActiveUnits   = 100
                    ConsumedUnits = 42
                }
            )
            GroupLicensingSummary = @(
                [pscustomobject]@{
                    GroupName                    = 'Empty Assigned License Signal Group'
                    IsManagingLicenses           = $true
                    AssignedLicenseCount         = 0
                    AssignedLicenseFriendlyNames = ''
                },
                [pscustomobject]@{
                    GroupName                    = 'M365 E3 License Group'
                    IsManagingLicenses           = $true
                    AssignedLicenseCount         = 1
                    AssignedLicenseFriendlyNames = 'Microsoft 365 E3'
                }
            )
            LicenseOptimizationCandidates = @(
                [pscustomobject]@{
                    ObjectType  = 'Group'
                    DisplayName = 'Empty Assigned License Signal Group'
                    Issue       = 'Licensing group has no owner'
                    SkuNames    = ''
                },
                [pscustomobject]@{
                    ObjectType  = 'Group'
                    DisplayName = 'M365 E3 License Group'
                    Issue       = 'Licensing group has no owner'
                    SkuNames    = ''
                },
                [pscustomobject]@{
                    ObjectType  = 'User'
                    DisplayName = 'Licensed User'
                    Issue       = 'Same SKU assigned directly and by group'
                    SkuNames    = 'Microsoft 365 E3'
                }
            )
            AdConnectConfiguration = @(
                [pscustomobject]@{
                    DirectorySyncEnabled = $true
                }
            )
            GuestSignInSummary = @{
                Summary = [pscustomobject]@{
                    InactiveGuests90Days = 3
                }
            }
            GuestAccessConfiguration = @{
                Summary = [pscustomobject]@{
                    GuestInvitationControl = 'adminsAndGuestInviters'
                }
            }
            ExternalIdentityRestrictions = @{
                Summary = [pscustomobject]@{
                    CrossTenantPartnerCount = 2
                }
            }
            SharePoint = @(
                [pscustomobject]@{
                    Title = 'Projects'
                    Url   = 'https://contoso.sharepoint.com/sites/projects'
                }
            )
            SharePointSharingSummary = @{
                Summary = [pscustomobject]@{
                    TenantSharingCapability = 'ExternalUserAndGuestSharing'
                    DefaultSharingLinkType  = 'AnonymousAccess'
                }
            }
            MfaEnrollmentSummary = [pscustomobject]@{
                RegistrationPercent = 72
            }
            MfaEnforcementSummary = [pscustomobject]@{
                EnabledPoliciesRequiringMfa = 1
            }
            AuthenticationConfig = @(
                [pscustomobject]@{
                    DefaultMfaState = 'Enabled'
                }
            )
            MfaEnforcementGapUsers = @(
                [pscustomobject]@{
                    UserPrincipalName = 'uncovered.member@contoso.com'
                    GapCategory       = 'Outside enabled MFA CA include scope'
                }
            )
            MfaEnforcementScopeReview = @(
                [pscustomobject]@{
                    PolicyName  = 'Baseline MFA'
                    ScopeType   = 'Exclude'
                    ObjectType  = 'Group'
                    DisplayName = 'Break Glass Exclusions'
                }
            )
            ConditionalAccessPolicySummary = @{
                Summary = [pscustomobject]@{
                    EnabledPolicies = 2
                }
            }
            AllRecipients = @(
                [pscustomobject]@{
                    DisplayName = 'Shared Mailbox'
                }
            )
            HybridConfiguration = @(
                [pscustomobject]@{
                    HybridEnabled = $true
                }
            )
            MailFlowRules = @(
                [pscustomobject]@{
                    Name = 'Block External Auto Forwarding'
                }
            )
            EmailActivityTopSenders = @(
                [pscustomobject]@{
                    UserPrincipalName = 'sender@contoso.com'
                }
            )
            EmailActivityTopReceivers = @(
                [pscustomobject]@{
                    UserPrincipalName = 'receiver@contoso.com'
                }
            )
            UnifiedGroups = @(
                [pscustomobject]@{
                    DisplayName = 'Projects'
                }
            )
            DeviceDetails = @(
                [pscustomobject]@{
                    DeviceName = 'CONTOSO-LT-01'
                }
            )
            SecuritySecureScore = @(
                [pscustomobject]@{
                    CurrentScore = 55
                }
            )
            ExternalSharingSummary = @{
                Summary = [pscustomobject]@{
                    SharingDomainRestrictionMode = 'allowList'
                }
            }
            ExternalSharingSiteOverrides = @(
                [pscustomobject]@{
                    Title          = 'Projects'
                    Url            = 'https://contoso.sharepoint.com/sites/projects'
                    OverrideReason = 'Sharing capability differs from tenant setting'
                }
            )
            ExternalExposureFindings = @(
                [pscustomobject]@{
                    Workload       = 'SharePoint'
                    Title          = 'Projects'
                    ReviewPriority = 'High'
                }
            )
        }

        Export-HashTableToExcel -hashtable $tenantStats -ExportDetails $exportPath

        $worksheetNames = @(Get-ExcelSheetInfo -Path $exportPath | Select-Object -ExpandProperty Name)
        ($worksheetNames -contains 'GuestAccessConfiguration') | Should -BeTrue
        ($worksheetNames -contains 'ExternalIdentityRestrictions') | Should -BeTrue
        ($worksheetNames -contains 'MfaEnrollmentSummary') | Should -BeTrue
        ($worksheetNames -contains 'MfaEnforcementSummary') | Should -BeTrue
        ($worksheetNames -contains 'MfaEnforcementGapUsers') | Should -BeTrue
        ($worksheetNames -contains 'MfaEnforcementScopeReview') | Should -BeTrue
        ($worksheetNames -contains 'ExternalSharingSummary') | Should -BeTrue
        ($worksheetNames -contains 'ExternalSharingSiteOverrides') | Should -BeTrue
        ($worksheetNames -contains 'ExternalExposureFindings') | Should -BeTrue
        ($worksheetNames -contains 'LicenseSKUs') | Should -BeTrue
        ($worksheetNames -contains 'GroupLicensingSummary') | Should -BeTrue
        ($worksheetNames -contains 'LicenseOptimizationCandidates') | Should -BeTrue
        ($worksheetNames -contains 'HybridConfiguration') | Should -BeTrue
        $worksheetNames.IndexOf('TenantInfo') | Should -BeLessThan $worksheetNames.IndexOf('GuestSignInSummary')
        $worksheetNames.IndexOf('TenantInfo') | Should -BeLessThan $worksheetNames.IndexOf('LicenseSKUs')
        $worksheetNames.IndexOf('LicenseSKUs') | Should -BeLessThan $worksheetNames.IndexOf('GroupLicensingSummary')
        $worksheetNames.IndexOf('GroupLicensingSummary') | Should -BeLessThan $worksheetNames.IndexOf('LicenseOptimizationCandidates')
        $worksheetNames.IndexOf('LicenseSKUs') | Should -BeLessThan $worksheetNames.IndexOf('GuestSignInSummary')
        $worksheetNames.IndexOf('GuestSignInSummary') | Should -BeLessThan $worksheetNames.IndexOf('GuestAccessConfiguration')
        $worksheetNames.IndexOf('AuthenticationConfig') | Should -BeLessThan $worksheetNames.IndexOf('MfaEnrollmentSummary')
        $worksheetNames.IndexOf('MfaEnrollmentSummary') | Should -BeLessThan $worksheetNames.IndexOf('MfaEnforcementSummary')
        $worksheetNames.IndexOf('MfaEnforcementSummary') | Should -BeLessThan $worksheetNames.IndexOf('MfaEnforcementGapUsers')
        $worksheetNames.IndexOf('MfaEnforcementGapUsers') | Should -BeLessThan $worksheetNames.IndexOf('MfaEnforcementScopeReview')
        $worksheetNames.IndexOf('MfaEnforcementScopeReview') | Should -BeLessThan $worksheetNames.IndexOf('ConditionalAccessPolicySummary')
        $worksheetNames.IndexOf('GuestAccessConfiguration') | Should -BeLessThan $worksheetNames.IndexOf('ExternalIdentityRestrictions')
        $worksheetNames.IndexOf('HybridConfiguration') | Should -BeLessThan $worksheetNames.IndexOf('AllRecipients')
        $worksheetNames.IndexOf('AllRecipients') | Should -BeLessThan $worksheetNames.IndexOf('MailFlowRules')
        $worksheetNames.IndexOf('MailFlowRules') | Should -BeLessThan $worksheetNames.IndexOf('EmailActivityTopSenders')
        $worksheetNames.IndexOf('EmailActivityTopSenders') | Should -BeLessThan $worksheetNames.IndexOf('EmailActivityTopReceivers')
        $worksheetNames.IndexOf('EmailActivityTopReceivers') | Should -BeLessThan $worksheetNames.IndexOf('UnifiedGroups')
        $worksheetNames.IndexOf('SharePointSharingSummary') | Should -BeLessThan $worksheetNames.IndexOf('DeviceDetails')
        $worksheetNames.IndexOf('DeviceDetails') | Should -BeLessThan $worksheetNames.IndexOf('SecuritySecureScore')
        $worksheetNames.IndexOf('SharePointSharingSummary') | Should -BeLessThan $worksheetNames.IndexOf('ExternalSharingSummary')
        $worksheetNames.IndexOf('ExternalSharingSummary') | Should -BeLessThan $worksheetNames.IndexOf('ExternalSharingSiteOverrides')
        $worksheetNames.IndexOf('ExternalSharingSiteOverrides') | Should -BeLessThan $worksheetNames.IndexOf('ExternalExposureFindings')

        $groupLicensingRows = @(Import-Excel -Path $exportPath -WorksheetName 'GroupLicensingSummary')
        $groupLicensingRows.Count | Should -Be 1
        $groupLicensingRows[0].GroupName | Should -Be 'M365 E3 License Group'

        $licenseOptimizationRows = @(Import-Excel -Path $exportPath -WorksheetName 'LicenseOptimizationCandidates')
        $licenseOptimizationRows.Count | Should -Be 2
        @($licenseOptimizationRows | Where-Object { $_.DisplayName -eq 'Empty Assigned License Signal Group' }).Count | Should -Be 0
        @($licenseOptimizationRows | Where-Object { $_.ObjectType -eq 'User' }).Count | Should -Be 1
    }

    It 'uses a single workbook package session for faster Excel export' {
        $script:exportExcelSource | Should -Match 'Open-ExcelPackage -Path \$ExportDetails -Create'
        $script:exportExcelSource | Should -Match 'Export-Excel -ExcelPackage \$excelPackage'
        $script:exportExcelSource | Should -Match 'Close-ExcelPackage -ExcelPackage \$excelPackage'
        $script:exportExcelSource | Should -Match '\$autoSizeRowLimit = 1000'
        $script:exportExcelSource | Should -Match "\$singleRecordWorksheets = @\('TenantInfo'\)"
        $script:exportExcelSource | Should -Match '\$autoSizeSheet = \(\$sourceCount -le \$autoSizeRowLimit -and \$WorkbookExportPolicy -ne ''TenantToTenantCutover''\)'
        $script:exportExcelSource | Should -Match 'Set-TenantToTenantWorksheetColumnWidths -ExcelPackage \$excelPackage -WorksheetName \$worksheetName -Columns \$explicitColumns'
        $script:exportExcelSource | Should -Match "'PrivilegedAccessRemediationSummary' = 'PrivilegedAccessRemediation'"
        $script:exportExcelSource | Should -Match "'PrivilegedAccessRemediationSummary' = @\('PrivilegedAccessRemediationSummary', 'PrivilegedAccessRemediationSumm', 'PrivilegedAccessRemediation'\)"
    }

    It 'adds full-assessment governance datasets to workbook export while excluding them from T2T output' {
        $script:htmlHelperSource | Should -Match '\$isFullAssessmentGovernanceScope = \('
        $script:htmlHelperSource | Should -Match "\$collectionDepthMode -in @\('Operator', 'Automation', 'Geek', 'All'\)"
        $script:htmlHelperSource | Should -Match 'if \(\$IncludeBestPracticeTables -and \$isFullAssessmentGovernanceScope\)'

        @(
            'ConditionalAccessOptimization'
            'MfaMethodPostureSummary'
            'PrivilegedAccessRemediationSummary'
            'TeamsGroupsCleanupCandidates'
            'GroupLicensingSummary'
            'LicenseOptimizationCandidates'
        ) | ForEach-Object {
            $script:exportExcelSource | Should -Match ([regex]::Escape("'$_'"))
            $script:exportExcelSource | Should -Match ([regex]::Escape("""$_"""))
        }

        $script:exportExcelSource | Should -Match "'SharePointSharingSummary', 'AllExchangeGroups', 'MigrationReadiness', 'ConditionalAccessOptimization'"
        $script:exportExcelSource | Should -Match "'MfaMethodPostureSummary', 'PrivilegedAccessRemediationSummary', 'TeamsGroupsCleanupCandidates'"
        $script:exportExcelSource | Should -Match "'GroupLicensingSummary', 'LicenseOptimizationCandidates'"
    }

    It 'narrows the workbook to the migration-focused sheet set for tenant-to-tenant exports' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        if (
            -not (Get-Command -Name Get-ExcelSheetInfo -ErrorAction SilentlyContinue) -or
            -not (Get-Command -Name Import-Excel -ErrorAction SilentlyContinue)
        ) {
            Set-ItResult -Skipped -Because 'ImportExcel worksheet inspection is not available in this environment.'
            return
        }

        function global:Write-Log { param() }
        function global:Write-ProgressHelper { param() }

        $exportPath = Join-Path $TestDrive 't2t-cutover.xlsx'
        $tenantStats = @{
            TenantInfo = [pscustomobject]@{
                DisplayName = 'Contoso'
                SelfServicePurchase = [pscustomobject]@{
                    Answer = 'Unknown'
                    Notes = 'Raw nested object should not leak into the T2T workbook.'
                }
            }
            TenantInfoSummary = @{
                Summary = [pscustomobject]@{
                    DisplayName               = 'Contoso'
                    TenantId                  = '11111111-1111-1111-1111-111111111111'
                    InitialDomain             = 'contoso.onmicrosoft.com'
                    DefaultDomain             = 'contoso.com'
                    Country                   = 'US'
                    CountryLetterCode         = 'US'
                    PreferredDataLocation     = 'NAM'
                    MultiGeoEnabled           = $false
                    MultiGeoAllowed           = $null
                    MultiGeoCentral           = $null
                    SelfServicePurchase       = 'No (Disabled for all 3 products)'
                    SelfServicePurchaseNotes  = 'Disabled: 3; Trial-only: 0'
                    AzureResourceUsage        = 'Not collected'
                    AzureResourceUsageNotes   = 'Azure module not used in this report'
                }
            }
            MigrationExecutiveSummary = @([pscustomobject]@{ Section = 'Tenant'; Metric = 'Enabled member users'; Value = 10; Notes = 'Wave planning population.' })
            RecipientDomainSummary = @([pscustomobject]@{ Domain = 'contoso.com'; RecipientCount = 25 })
            MailboxMigrationSummary = @([pscustomobject]@{ RecipientTypeDetails = 'UserMailbox'; MailboxCount = 20 })
            BitTitanLicenseSummary = @([pscustomobject]@{ Section = 'Licensing'; Metric = 'MigrationWiz-Mailbox'; Value = 15 })
            DelegateSummary = @([pscustomobject]@{ PermissionType = 'FullAccess'; AffectedMailboxCount = 4; AssignmentCount = 7 })
            CollaborationSummary = @([pscustomobject]@{ Workload = 'SharePoint'; TotalCount = 12; TotalStorageGB = 88.2 })
            CutoverPrepSummary = @([pscustomobject]@{ Category = 'Routing'; Item = 'Mailboxes missing .onmicrosoft alias'; Status = 'Review'; Value = 2 })
            Domains = @([pscustomobject]@{ Name = 'contoso.com'; AuthenticationType = 'Managed'; SupportedServices = 'Email'; MXRecords = 'mail.protection.outlook.com'; Office365MailExchanger = $true; ThirdPartySpamFilterReview = 'No immediate signal'; HybridRoutingReview = 'Review'; RecipientCount = 25 })
            AdConnectConfiguration = @([pscustomobject]@{ DirectorySyncEnabled = $true })
            TenantToTenantAdConnectConfiguration = @([pscustomobject]@{ Section = 'Directory sync'; Item = 'On-prem sync enabled'; Value = 'True'; Notes = 'Microsoft Entra Connect detected' })
            HybridConfiguration = @([pscustomobject]@{ IsConfigured = $true })
            TenantToTenantHybridConfiguration = @([pscustomobject]@{ Section = 'Hybrid posture'; Item = 'Hybrid configured'; Value = 'True'; Notes = 'Mail flow connector evidence detected' })
            Users = @([pscustomobject]@{ DisplayName = 'User One'; UserPrincipalName = 'user1@contoso.com'; Mail = 'user1@contoso.com'; AccountEnabled = $true; OnPremisesSyncEnabled = $true; AssignedLicensesFriendly = 'ENTERPRISEPACK'; HasMailbox = $true; MailboxPrimarySmtpAddress = 'user1@contoso.com'; RecipientTypeDetails = 'UserMailbox' })
            AllRecipients = @([pscustomobject]@{ DisplayName = 'User One'; PrimarySmtpAddress = 'user1@contoso.com'; RecipientTypeDetails = 'UserMailbox'; UserPrincipalName = 'user1@contoso.com'; WindowsEmailAddress = 'user1@contoso.com'; PrimaryDomain = 'contoso.com'; EmailAddresses = 'SMTP:user1@contoso.com;smtp:user1@contoso.onmicrosoft.com'; OnMicrosoftAlias = 'user1@contoso.onmicrosoft.com'; OnMicrosoftAliases = 'user1@contoso.onmicrosoft.com'; HiddenFromAddressListsEnabled = $false })
            AllMailboxes = @([pscustomobject]@{ DisplayName = 'Raw User One'; PrimarySmtpAddress = 'raw-user1@contoso.com' })
            MailboxFullDetails = @([pscustomobject]@{
                DisplayName = 'User One'; UserPrincipalName = 'user1@contoso.com'; PrimarySmtpAddress = 'user1@contoso.com'; RecipientTypeDetails = 'UserMailbox'; IsInactiveMailbox = $false;
                OnPremisesSyncEnabled = $true; ExchangeGuid = '11111111-1111-1111-1111-111111111111'; ArchiveGuid = '22222222-2222-2222-2222-222222222222';
                OnMicrosoftAlias = 'user1@contoso.onmicrosoft.com'; OnMicrosoftAliases = 'user1@contoso.onmicrosoft.com'; LegacyExchangeDn = '/o=Contoso/ou=Exchange Administrative Group/cn=Recipients/cn=user1';
                LegacyExchangeDnX500 = 'x500:/o=Contoso/ou=Exchange Administrative Group/cn=Recipients/cn=user1'; X500Addresses = 'x500:/o=Contoso/ou=Exchange Administrative Group/cn=Recipients/cn=user1';
                X400Addresses = 'x400:c=US;a= ;p=Contoso;o=Exchange;s=User1;'; MailboxSizeGB = 48.5; DeletedItemsGB = 3.2; ArchiveStatus = 'Active'; ArchiveSizeGB = 10.1;
                ArchiveDeletedItemsGB = 0.5; TotalDataToMigrateGB = 62.3; BitTitanLicenseType = 'User Migration Bundle'; BitTitanLicenseCount = 1; FullAccessDelegateCount = 1;
                SendAsDelegateCount = 0; GrantSendOnBehalfToCount = 0; CalendarDelegateCount = 2
            })
            MailboxDelegateAssignments = @([pscustomobject]@{ MailboxDisplayName = 'User One'; MailboxPrimarySmtpAddress = 'user1@contoso.com'; MailboxUserPrincipalName = 'user1@contoso.com'; RecipientTypeDetails = 'UserMailbox'; PermissionType = 'FullAccess'; Delegate = 'assistant@contoso.com'; DelegateCountSource = 1; CollectionState = 'Collected' })
            MailboxCalendarDelegatePermissions = @([pscustomobject]@{ MailboxPrimarySmtpAddress = 'user1@contoso.com'; PermissionTarget = 'delegate@contoso.com'; AccessRights = 'Editor' })
            InactiveMailboxDetails = @([pscustomobject]@{ DisplayName = 'Former User'; UserPrincipalName = 'former@contoso.com'; PrimarySmtpAddress = 'former@contoso.com'; RecipientTypeDetails = 'InactiveMailbox'; IsInactiveMailbox = $true; MailboxSizeGB = 12.1; TotalDataToMigrateGB = 12.1 })
            PublicFolderDetails = @([pscustomobject]@{ Name = 'PF Root' })
            MailFlowConnectors = @([pscustomobject]@{ Name = 'Inbound Connector' })
            RemoteDomains = @([pscustomobject]@{ Name = 'partner.com' })
            SMTPRelayServiceAccounts = @([pscustomobject]@{ UserPrincipalName = 'relay@contoso.com' })
            AllTeams = @([pscustomobject]@{ DisplayName = 'Projects Team'; Visibility = 'Private'; IsArchived = $false; SharePointSiteUrl = 'https://contoso.sharepoint.com/sites/projects'; 'SiteSize-GB' = 44.5; TotalChannels = 12; SharedChannelCount = 2; SharedChannels = 'Vendors; Program Office'; OwnerCount = 2; MemberCount = 16; GuestCount = 1; LastActivityDate = '2026-04-01' })
            SharePoint = @([pscustomobject]@{ Title = 'Projects'; Url = 'https://contoso.sharepoint.com/sites/projects'; Template = 'TEAMSITE'; Owner = 'owner@contoso.com'; StorageUsedGB = 44.5; StorageQuota = 1024; LastContentModifiedDate = '2026-04-01'; LockState = 'Unlock'; ArchiveStatus = 'NotArchived'; SharingCapability = 'ExternalUserAndGuestSharing'; IsTeamsConnected = $true })
            OneDrive = @([pscustomobject]@{ Title = 'User One OneDrive'; Url = 'https://contoso-my.sharepoint.com/personal/user1_contoso_com'; Template = 'SPSPERS'; Owner = 'user1@contoso.com'; StorageUsedGB = 18.4; StorageQuota = 1024; LastContentModifiedDate = '2026-04-02'; LockState = 'Unlock'; ArchiveStatus = 'NotArchived'; SharingCapability = 'Disabled'; IsTeamsConnected = $false })
            AuthenticationConfig = @([pscustomobject]@{ DefaultMfaState = 'Enabled' })
            AuthenticationSSOApplications = @([pscustomobject]@{ DisplayName = 'Salesforce' })
            ExternalSharingSummary = @{ Summary = [pscustomobject]@{ TenantSharingCapability = 'ExternalUserAndGuestSharing' } }
            SecuritySecureScore = @([pscustomobject]@{ CurrentScore = 55 })
        }

        Export-HashTableToExcel -hashtable $tenantStats -ExportDetails $exportPath -WorkbookExportPolicy TenantToTenantCutover

        $worksheetNames = @(Get-ExcelSheetInfo -Path $exportPath | Select-Object -ExpandProperty Name)
        @('AllMailboxes', 'TenantInfo', 'MigrationExecutiveSummary', 'RecipientDomainSummary', 'MailboxMigrationSummary', 'BitTitanLicenseSummary', 'DelegateSummary', 'CollaborationSummary', 'CutoverPrepSummary', 'Domains', 'AdConnectConfiguration', 'HybridConfiguration', 'Users', 'AllRecipients', 'MailboxDelegateAssignments', 'MailboxCalendarDelegatePerms', 'InactiveMailboxDetails', 'PublicFolderDetails', 'MailFlowConnectors', 'RemoteDomains', 'SMTPRelayServiceAccounts', 'AllTeams', 'SharePoint', 'OneDrive') | ForEach-Object {
            ($worksheetNames -contains $_) | Should -BeTrue
        }
        @('AuthenticationConfig', 'AuthenticationSSOApplications', 'ExternalSharingSummary', 'SecuritySecureScore', 'MailboxFullDetails') | ForEach-Object {
            ($worksheetNames -contains $_) | Should -BeFalse
        }
        $worksheetNames[0] | Should -Be 'AllMailboxes'

        $tenantInfoSheet = @(Import-Excel -Path $exportPath -WorksheetName 'TenantInfo')
        $tenantInfoSheet[0].PSObject.Properties.Name | Should -Be @('DisplayName', 'TenantId', 'InitialDomain', 'DefaultDomain', 'Country', 'CountryLetterCode', 'PreferredDataLocation', 'MultiGeoEnabled', 'MultiGeoAllowed', 'MultiGeoCentral', 'SelfServicePurchase', 'SelfServicePurchaseNotes', 'AzureResourceUsage', 'AzureResourceUsageNotes')
        $tenantInfoSheet[0].SelfServicePurchase | Should -Be 'No (Disabled for all 3 products)'
        $tenantInfoSheet[0].SelfServicePurchaseNotes | Should -Be 'Disabled: 3; Trial-only: 0'
        $tenantInfoSheet[0].PSObject.Properties.Name | Should -Not -Contain 'Answer'
        $tenantInfoSheet[0].PSObject.Properties.Name | Should -Not -Contain 'Notes'

        $adConnectSheet = @(Import-Excel -Path $exportPath -WorksheetName 'AdConnectConfiguration')
        $adConnectSheet[0].PSObject.Properties.Name | Should -Be @('Section', 'Item', 'Value', 'Notes')
        $adConnectSheet[0].Item | Should -Be 'On-prem sync enabled'

        $hybridSheet = @(Import-Excel -Path $exportPath -WorksheetName 'HybridConfiguration')
        $hybridSheet[0].PSObject.Properties.Name | Should -Be @('Section', 'Item', 'Value', 'Notes')
        $hybridSheet[0].Item | Should -Be 'Hybrid configured'

        $userSheet = @(Import-Excel -Path $exportPath -WorksheetName 'Users')
        $userSheet[0].PSObject.Properties.Name | Should -Be @('DisplayName', 'UserPrincipalName', 'Mail', 'AccountEnabled', 'OnPremisesSyncEnabled', 'AssignedLicensesFriendly', 'HasMailbox', 'MailboxPrimarySmtpAddress', 'RecipientTypeDetails', 'TargetUPN', 'TargetPrimarySmtpAddress', 'Wave', 'MigrationState', 'CutoverDate', 'Notes')

        $mailboxSheet = @(Import-Excel -Path $exportPath -WorksheetName 'AllMailboxes')
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'MailboxSizeGB'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'TotalDataToMigrateGB'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'BitTitanLicenseType'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'ForwardingAddress'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Not -Contain 'ArchiveName'

        $domainSheet = @(Import-Excel -Path $exportPath -WorksheetName 'Domains')
        $domainSheet[0].PSObject.Properties.Name | Should -Contain 'MXRecords'
        $domainSheet[0].MXRecords | Should -Be 'mail.protection.outlook.com'

        $teamsSheet = @(Import-Excel -Path $exportPath -WorksheetName 'AllTeams')
        $teamsSheet[0].PSObject.Properties.Name | Should -Contain 'SharedChannelCount'
        $teamsSheet[0].PSObject.Properties.Name | Should -Contain 'SharedChannels'
        $teamsSheet[0].SharedChannelCount | Should -Be 2

        $delegateSheet = @(Import-Excel -Path $exportPath -WorksheetName 'MailboxDelegateAssignments')
        $delegateSheet[0].PSObject.Properties.Name | Should -Be @('MailboxDisplayName', 'MailboxPrimarySmtpAddress', 'MailboxUserPrincipalName', 'RecipientTypeDetails', 'PermissionType', 'Delegate', 'DelegateCountSource', 'CollectionState', 'Wave', 'Notes')
    }

    It 'exports a tenant-to-tenant cutover pack workbook and CSV companions with the required mailbox migration columns' {
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        if (
            -not (Get-Command -Name Get-ExcelSheetInfo -ErrorAction SilentlyContinue) -or
            -not (Get-Command -Name Import-Excel -ErrorAction SilentlyContinue)
        ) {
            Set-ItResult -Skipped -Because 'ImportExcel workbook inspection commands are not available in this environment.'
            return
        }

        $tenantStats = @{
            TenantInfo = [pscustomobject]@{ DisplayName = 'Contoso' }
            Users = @(
                [pscustomobject]@{
                    UserPrincipalName   = 'user1@contoso.com'
                    AccountEnabled      = $true
                    Mail                = 'user1@contoso.com'
                    OnPremisesSyncEnabled = $true
                },
                [pscustomobject]@{
                    UserPrincipalName   = 'nomailbox@contoso.com'
                    AccountEnabled      = $true
                    Mail                = 'nomailbox@contoso.com'
                    OnPremisesSyncEnabled = $false
                },
                [pscustomobject]@{
                    UserPrincipalName   = 'upnonly@contoso.com'
                    AccountEnabled      = $true
                    Mail                = $null
                    OnPremisesSyncEnabled = $false
                }
            )
            AllRecipients = @(
                [pscustomobject]@{
                    DisplayName                  = 'User One'
                    PrimarySmtpAddress           = 'user1@contoso.com'
                    RecipientTypeDetails         = 'UserMailbox'
                    UserPrincipalName            = 'user1@contoso.com'
                    WindowsEmailAddress          = 'user1@contoso.com'
                    EmailAddresses               = @('SMTP:user1@contoso.com', 'smtp:user1@contoso.onmicrosoft.com')
                    HiddenFromAddressListsEnabled = $false
                }
            )
            MailboxFullDetails = @(
                [pscustomobject]@{
                    DisplayName              = 'User One'
                    UserPrincipalName        = 'user1@contoso.com'
                    PrimarySmtpAddress       = 'user1@contoso.com'
                    RecipientTypeDetails     = 'UserMailbox'
                    ExchangeGuid             = '11111111-1111-1111-1111-111111111111'
                    ArchiveGuid              = '22222222-2222-2222-2222-222222222222'
                    OnMicrosoftAlias         = 'user1@contoso.onmicrosoft.com'
                    OnMicrosoftAliases       = 'user1@contoso.onmicrosoft.com'
                    LegacyExchangeDn         = '/o=Contoso/ou=Exchange Administrative Group/cn=Recipients/cn=user1'
                    LegacyExchangeDnX500     = 'x500:/o=Contoso/ou=Exchange Administrative Group/cn=Recipients/cn=user1'
                    X500Addresses            = 'x500:/o=Contoso/ou=Exchange Administrative Group/cn=Recipients/cn=user1'
                    X400Addresses            = 'x400:c=US;a= ;p=Contoso;o=Exchange;s=User1;'
                    CalendarDelegates        = 'calendar.delegate@contoso.com'
                    CalendarDelegateCount    = 1
                    CalendarPermissionEntryCount = 1
                    CalendarDelegateState    = 'Collected'
                    FullAccessDelegates      = 'assistant@contoso.com'
                    FullAccessDelegateCount  = 1
                    FullAccessDelegateState  = 'Collected'
                    SendAsDelegates          = 'helpdesk@contoso.com'
                    SendAsDelegateCount      = 1
                    SendAsDelegateState      = 'Collected'
                    GrantSendOnBehalfTo      = 'delegate@contoso.com'
                    GrantSendOnBehalfToCount = 1
                    ForwardingAddress        = 'forwarding-mailuser'
                    ForwardingSmtpAddress    = $null
                    DeliverToMailboxAndForward = $false
                    IsInactiveMailbox        = $false
                    ArchiveStatus            = 'Active'
                    LitigationHoldEnabled    = $false
                    RetentionPolicy          = 'Default MRM Policy'
                    EmailAddresses           = @('SMTP:user1@contoso.com', 'smtp:user1@contoso.onmicrosoft.com')
                }
            )
            MailboxCalendarDelegatePermissions = @(
                [pscustomobject]@{
                    MailboxDisplayName        = 'User One'
                    MailboxPrimarySmtpAddress = 'user1@contoso.com'
                    MailboxUserPrincipalName  = 'user1@contoso.com'
                    RecipientTypeDetails      = 'UserMailbox'
                    CalendarName              = 'Calendar'
                    CalendarPath              = '/Calendar'
                    PermissionTarget          = 'calendar.delegate@contoso.com'
                    PermissionTargetDisplayName = 'Calendar Delegate'
                    PermissionTargetType      = 'UserMailbox'
                    AccessRights              = 'Editor'
                    SharingPermissionFlags    = 'Delegate'
                }
            )
            OneDrive = @(
                [pscustomobject]@{
                    Url   = 'https://contoso-my.sharepoint.com/personal/user1_contoso_com'
                    Owner = 'user1@contoso.com'
                }
            )
        }

        $baseExportPath = Join-Path $TestDrive 'Contoso - Tenant Details.xlsx'
        $result = Export-ArrayaTenantToTenantCutoverPack -TenantStatsHash $tenantStats -BaseExportPath $baseExportPath

        Test-Path $result.WorkbookPath | Should -BeTrue
        @('PackSummary', 'UserWaveTemplate', 'MailboxCutoverReference', 'CalendarDelegateReference', 'RecipientDomainReference', 'CollaborationAccessTemplate', 'OneDriveMigrationTemplate') | ForEach-Object {
            ($result.CsvArtifacts.Contains($_)) | Should -BeTrue
            Test-Path $result.CsvArtifacts[$_] | Should -BeTrue
        }

        $sheetNames = @(Get-ExcelSheetInfo -Path $result.WorkbookPath | Select-Object -ExpandProperty Name)
        ($sheetNames -contains 'MailboxCutoverReference') | Should -BeTrue
        ($sheetNames -contains 'CalendarDelegateReference') | Should -BeTrue
        ($sheetNames -contains 'UserWaveTemplate') | Should -BeTrue

        $mailboxSheet = @(Import-Excel -Path $result.WorkbookPath -WorksheetName 'MailboxCutoverReference')
        $mailboxSheet.Count | Should -BeGreaterThan 0
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'OnMicrosoftAlias'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'LegacyExchangeDnX500'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'X500Addresses'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'X400Addresses'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'MailboxSizeGB'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'DeletedItemsGB'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'ArchiveSizeGB'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'ArchiveDeletedItemsGB'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'TotalDataToMigrateGB'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'BitTitanLicenseType'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'BitTitanLicenseCount'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'ForwardingAddress'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'CalendarDelegateCount'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'FullAccessDelegateCount'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'SendAsDelegateCount'
        $mailboxSheet[0].PSObject.Properties.Name | Should -Contain 'GrantSendOnBehalfToCount'

        $calendarSheet = @(Import-Excel -Path $result.WorkbookPath -WorksheetName 'CalendarDelegateReference')
        $calendarSheet.Count | Should -BeGreaterThan 0
        $calendarSheet[0].PSObject.Properties.Name | Should -Contain 'PermissionTarget'
        $calendarSheet[0].PSObject.Properties.Name | Should -Contain 'AccessRights'
        $calendarSheet[0].PSObject.Properties.Name | Should -Contain 'SharingPermissionFlags'

        $userWaveSheet = @(Import-Excel -Path $result.WorkbookPath -WorksheetName 'UserWaveTemplate')
        (@($userWaveSheet | Where-Object { $_.SourceUPN -eq 'nomailbox@contoso.com' })).Count | Should -Be 1
        (@($userWaveSheet | Where-Object { $_.SourceUPN -eq 'upnonly@contoso.com' })).Count | Should -Be 1
    }

    It 'does not throw when Write-Log captures an error message without an ErrorRecordVar' {
        $writeLogPath = Join-Path $script:repoRoot 'src\vendor\Office365Custom\1.2.1\Public\Write-Log.ps1'
        Test-Path $writeLogPath | Should -BeTrue

        $script:captureErrorHelperInvocations = 0
        function global:Capture-ErrorHelper {
            param(
                [Parameter(Mandatory = $true)]
                $ErrorRecordVar,
                [Parameter(Mandatory = $true)]
                [string]$errorMessage
            )

            $script:captureErrorHelperInvocations++
        }

        . $writeLogPath

        { Write-Log -Type ERROR -Message 'Synthetic error without ErrorRecordVar' } | Should -Not -Throw
        $script:captureErrorHelperInvocations | Should -Be 0

        Remove-Item Function:\Write-Log -ErrorAction SilentlyContinue
        Remove-Item Function:\Capture-ErrorHelper -ErrorAction SilentlyContinue
    }

    It 'does not emit a redundant Write-Error before rethrowing Graph request failures' {
        $graphDataPath = Join-Path $script:repoRoot 'src\vendor\Office365Custom\1.2.1\Public\Get-GraphData.ps1'
        Test-Path $graphDataPath | Should -BeTrue

        $graphDataSource = Get-Content -Raw -Path $graphDataPath
        $graphDataSource | Should -Not -Match 'Write-Error "Graph API request failed:'
        $graphDataSource | Should -Match 'Write-Verbose "Graph API request failed:'
        $graphDataSource | Should -Match 'throw "Failed to retrieve data after \$MaxRetries attempts from: \$CurrentUri"'
    }

    It 'skips SharePoint module import for app-based auth and relies on Graph collection instead' {
        $connectOffice365Path = Join-Path $script:repoRoot 'src\vendor\Office365Custom\1.2.1\Public\Connect-Office365.ps1'
        Test-Path $connectOffice365Path | Should -BeTrue

        $connectOffice365Source = Get-Content -Raw -Path $connectOffice365Path
        $connectOffice365Source | Should -Match "SharePoint Online SPO cmdlets are skipped for app-based authentication"
        $connectOffice365Source | Should -Match "Skipping SharePoint module import and Connect-SPOService for authentication type"
    }

    It 'prefers lightweight Exchange connection metadata before probing remote Exchange cmdlets' {
        $connectOffice365Path = Join-Path $script:repoRoot 'src\vendor\Office365Custom\1.2.1\Public\Connect-Office365.ps1'
        Test-Path $connectOffice365Path | Should -BeTrue

        $connectOffice365Source = Get-Content -Raw -Path $connectOffice365Path
        $connectOffice365Source | Should -Match "Get-ConnectionInformation"
        $connectOffice365Source | Should -Match "Get-OrganizationConfig"
    }

    It 'aligns the default delegated Graph scope ask to the standard assessment core set' {
        $connectOffice365Path = Join-Path $script:repoRoot 'src\vendor\Office365Custom\1.2.1\Public\Connect-Office365.ps1'
        $graphSdkHelperPath = Join-Path $script:repoRoot 'src\vendor\Office365Custom\1.2.1\Private\Connect-MicrosoftGraphSDK.ps1'
        Test-Path $connectOffice365Path | Should -BeTrue
        Test-Path $graphSdkHelperPath | Should -BeTrue

        $connectOffice365Source = Get-Content -Raw -Path $connectOffice365Path
        $graphSdkHelperSource = Get-Content -Raw -Path $graphSdkHelperPath

        $connectOffice365Source | Should -Match 'SharePointTenantSettings\.Read\.All'
        $connectOffice365Source | Should -Not -Match 'Files\.Read\.All'
        $graphSdkHelperSource | Should -Match 'SharePointTenantSettings\.Read\.All'
        $graphSdkHelperSource | Should -Not -Match 'Files\.Read\.All'
    }

    It 'loads the common module cleanly even when the Private folder is absent' {
        $commonModulePath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psm1'
        Test-Path $commonModulePath | Should -BeTrue

        $commonModuleSource = Get-Content -Raw -Path $commonModulePath
        $commonModuleSource | Should -Match '\$privatePath = Join-Path \$PSScriptRoot ''Private'''
        $commonModuleSource | Should -Match 'if \(Test-Path -Path \$privatePath -PathType Container\)'
    }

    It 'guards Exchange compatibility differences before using hybrid and test-mode connector parameters' {
        $hybridPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Public\Get-ExchangeHybridConfiguration.ps1'
        $mailFlowPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Public\Get-MailFlowRulesandConnectors.ps1'
        Test-Path $hybridPath | Should -BeTrue
        Test-Path $mailFlowPath | Should -BeTrue

        $hybridSource = Get-Content -Raw -Path $hybridPath
        $mailFlowSource = Get-Content -Raw -Path $mailFlowPath

        $hybridSource | Should -Match "Get-Command -Name 'Get-HybridConfiguration' -ErrorAction Ignore"
        $mailFlowSource | Should -Match "Get-Command -Name 'Get-TransportRule' -ErrorAction Ignore"
        $mailFlowSource | Should -Match 'Parameters\.ContainsKey\(''IncludeTestModeConnectors''\)'
        $mailFlowSource | Should -Match "Get-Command -Name 'Get-OutboundConnector' -ErrorAction Ignore"
    }
}
