Describe 'Presales migration scope overview' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:exporterPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Export-PresalesMigrationScopeQuestionnaireMarkdown.ps1'
        $script:pipelinePath = Join-Path $script:repoRoot 'src\scripts\reporting\Invoke-M365TenantAssessmentExportPipeline.ps1'
        $script:toolingCatalogPath = Join-Path $script:repoRoot 'src\config\baseline\migration-tooling-options.json'
        . $script:exporterPath

        function Export-TestMigrationOverview {
            param(
                [hashtable]$Snapshot,
                [string]$OutputPath
            )

            Export-PresalesMigrationScopeQuestionnaireMarkdown `
                -TenantStatsHash $Snapshot `
                -Path $OutputPath `
                -PreparedBy 'Unit Test' `
                -AsOfDate ([datetime]'2026-08-14')
            return (Get-Content -LiteralPath $OutputPath -Raw)
        }

        $snapshot = @{
            TenantInfo = [pscustomobject]@{
                DisplayName   = 'Contoso'
                TenantId      = 'INTERNAL-TENANT-ID-MUST-NOT-LEAK'
                ObjectId      = 'INTERNAL-OBJECT-ID-MUST-NOT-LEAK'
                DefaultDomain = 'contoso.com'
            }
            MigrationExecutiveSummary = @(
                [pscustomobject]@{ Section = 'Tenant'; Metric = 'Enabled member users'; Value = 180 },
                [pscustomobject]@{ Section = 'Messaging'; Metric = 'Mailboxes'; Value = 160 },
                [pscustomobject]@{ Section = 'Messaging'; Metric = 'Grand total data to migrate (GB)'; Value = 1500 },
                [pscustomobject]@{ Section = 'Collaboration'; Metric = 'Teams'; Value = 2 },
                [pscustomobject]@{ Section = 'Collaboration'; Metric = 'SharePoint sites'; Value = 1 },
                [pscustomobject]@{ Section = 'Collaboration'; Metric = 'OneDrive sites'; Value = 1 }
            )
            MigrationScopeSummary = @(
                [pscustomobject]@{ Tool = 'BitTitan'; Section = 'Data'; Metric = 'Total mailbox data to migrate (GB)'; Value = 1500 },
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'Sites'; Metric = 'SharePoint sites (total)'; Value = 1 },
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'Data'; Metric = 'SharePoint storage (GB)'; Value = 18 },
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'OneDrive'; Metric = 'OneDrive sites'; Value = 1 },
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'OneDrive'; Metric = 'OneDrive storage (GB)'; Value = 10 },
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'Teams'; Metric = 'Teams'; Value = 2 },
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'Teams'; Metric = 'Microsoft 365 Groups'; Value = 3 }
            )
            ShareGateScopeSummary = @(
                [pscustomobject]@{
                    Workload = 'All collaboration sites (non-duplicated)'; ObjectCount = 2; TotalStorageGB = 28
                    LargestObjectName = 'Operations'; LargestObjectSizeGB = 18
                },
                [pscustomobject]@{ Workload = 'SharePoint (total)'; ObjectCount = 1; TotalStorageGB = 18 },
                [pscustomobject]@{ Workload = 'OneDrive'; ObjectCount = 1; TotalStorageGB = 10 }
            )
            InactiveMailboxDetails = @(
                [pscustomobject]@{
                    DisplayName = 'INACTIVE-NAME-MUST-NOT-LEAK'; PrimarySmtpAddress = 'archive-sentinel@contoso.com'
                    ExchangeGuid = 'INTERNAL-EXCHANGE-GUID-MUST-NOT-LEAK'; ArchiveStatus = 'Active'; ArchiveSizeGB = 44
                },
                [pscustomobject]@{ DisplayName = 'Unknown Archive'; PrimarySmtpAddress = 'unknown@contoso.com'; ArchiveStatus = 'Unknown' },
                [pscustomobject]@{ DisplayName = 'Unavailable Archive'; PrimarySmtpAddress = 'unavailable@contoso.com'; ArchiveStatus = 'Unavailable' },
                [pscustomobject]@{ DisplayName = 'Blank Archive'; PrimarySmtpAddress = 'blank@contoso.com' },
                [pscustomobject]@{ DisplayName = 'Conflicting Archive'; PrimarySmtpAddress = 'conflict@contoso.com'; HasArchive = $false; ArchiveStatus = 'Active'; ArchiveSizeGB = 12 }
            )
            BitTitanLicenseSummary = @(
                [pscustomobject]@{ Section = 'Licensing'; Metric = 'Mailbox Migration licenses'; Value = 12 },
                [pscustomobject]@{ Section = 'Licensing'; Metric = 'User Migration Bundles'; Value = 3 },
                [pscustomobject]@{ Section = 'Counts'; Metric = 'Archive-enabled mailboxes'; Value = 2 },
                [pscustomobject]@{ Section = 'Counts'; Metric = 'Microsoft 365 Group mailboxes needing evidence'; Value = 2 },
                [pscustomobject]@{ Section = 'Thresholds'; Metric = 'Mailboxes over 50 GB'; Value = 2 },
                [pscustomobject]@{ Section = 'Thresholds'; Metric = 'Mailboxes over 100 GB'; Value = 1 },
                [pscustomobject]@{ Section = 'Thresholds'; Metric = 'Archives over 100 GB'; Value = 1 }
            )
            GroupWorkloadReconciliation = @(
                [pscustomobject]@{
                    DisplayName = 'GROUP-NAME-MUST-NOT-LEAK'; PrimarySmtpAddress = 'group-sentinel@contoso.com'
                    DirectoryObjectId = 'INTERNAL-GROUP-ID-MUST-NOT-LEAK'; MailboxSizeGB = 2; MailboxEvidenceStatus = 'Measured data'
                    SiteStorageGB = 18; IsTeam = $true
                },
                [pscustomobject]@{ DisplayName = 'Unknown Group One'; MailboxSizeGB = $null; MailboxEvidenceStatus = 'Needs Data' },
                [pscustomobject]@{ DisplayName = 'Unknown Group Two'; MailboxSizeGB = $null; MailboxEvidenceStatus = 'Needs Data' }
            )
            AllTeams = @(
                [pscustomobject]@{
                    DisplayName = 'Operations'; SharedChannelCount = 1
                    ChannelInventoryStatus = 'Measured data'; MemberInventoryStatus = 'Measured data'
                },
                [pscustomobject]@{
                    DisplayName = 'Evidence Gap Team'; SharedChannelCount = 0
                    ChannelInventoryStatus = 'Needs Data'; MemberInventoryStatus = 'Needs Data'
                }
            )
            PublicFolderDetails = @(
                [pscustomobject]@{ Name = 'PUBLIC-FOLDER-NAME-MUST-NOT-LEAK'; EntryId = 'INTERNAL-PUBLIC-FOLDER-ID-MUST-NOT-LEAK' }
            )
            Domains = @(
                [pscustomobject]@{ Name = 'contoso.com' },
                [pscustomobject]@{ Name = 'contoso.onmicrosoft.com' }
            )
            MigrationComplexityFlags = @(
                [pscustomobject]@{
                    Status = 'Blocker'; Item = 'Applications with a custom-domain application URI'; Value = 2
                    Notes = 'A custom domain creates a hard sequencing dependency.'
                    MigrationAction = 'Sequence domain cutover before application re-registration.'
                    DecisionKey = 'INTERNAL-COMPLEXITY-KEY-MUST-NOT-LEAK'
                }
            )
            MigrationScopeDecisions = @(
                [pscustomobject]@{
                    DecisionKey = 'INTERNAL-DECISION-KEY-MUST-NOT-LEAK'; Status = 'Approved'
                    CustomerConfirmedValue = 'CUSTOMER-DECISION-MUST-NOT-LEAK'
                    TargetMapping = 'TARGET-MAPPING-MUST-NOT-LEAK'; DecisionOwner = 'OWNER-MUST-NOT-LEAK'
                }
            )
            MigrationTargetReadiness = @(
                [pscustomobject]@{ Item = 'TARGET-READINESS-MUST-NOT-LEAK'; Status = 'Blocked' }
            )
            MigrationWavePlan = @(
                [pscustomobject]@{ Wave = 'WAVE-BREAKDOWN-MUST-NOT-LEAK'; CandidateObjectCount = 99 }
            )
        }

        $script:outputPath = Join-Path $TestDrive 'Contoso-MigrationScopeQuestionnaire.md'
        $script:content = Export-TestMigrationOverview -Snapshot $snapshot -OutputPath $script:outputPath
        $script:pipelineSource = Get-Content -LiteralPath $script:pipelinePath -Raw
    }

    It 'renders a compact source-only overview instead of an interactive decision register' {
        Test-Path $script:outputPath | Should -BeTrue
        $script:content | Should -Match '# Microsoft 365 Migration Scope Overview'
        $script:content | Should -Match 'source-tenant-only overview'
        $script:content | Should -Match 'Full object inventories remain in the supporting workbook'
        $script:content | Should -Not -Match '(?i)Decision status|Decision key|Customer-confirmed value|Current Quote Readiness|Quote-readiness scale|Target readiness|Open readiness items'
    }

    It 'shows aggregate workload sizing without exposing tenant or object identifiers' {
        $script:content | Should -Match '\*\*Client:\*\* Contoso'
        $script:content | Should -Match '\*\*Source tenant domain:\*\* contoso\.com'
        $script:content | Should -Match '180 enabled member user\(s\)'
        $script:content | Should -Match '160 mailbox object\(s\)'
        $script:content | Should -Match '1\.46 TB known mailbox'
        $script:content | Should -Match '2 Team\(s\)'
        $script:content | Should -Match '28\.00 GB known, non-duplicated collaboration data'

        @(
            'INTERNAL-TENANT-ID-MUST-NOT-LEAK',
            'INTERNAL-OBJECT-ID-MUST-NOT-LEAK',
            'INTERNAL-EXCHANGE-GUID-MUST-NOT-LEAK',
            'INTERNAL-GROUP-ID-MUST-NOT-LEAK',
            'INTERNAL-PUBLIC-FOLDER-ID-MUST-NOT-LEAK',
            'INTERNAL-COMPLEXITY-KEY-MUST-NOT-LEAK',
            'INTERNAL-DECISION-KEY-MUST-NOT-LEAK'
        ) | ForEach-Object { $script:content | Should -Not -Match ([regex]::Escape($_)) }
    }

    It 'summarizes only exceptional mailbox, group, Teams, and public-folder conditions' {
        $script:content | Should -Match 'Inactive mailboxes and archives'
        $script:content | Should -Match '5 inactive mailbox\(es\); 2 have archive evidence and 3 have an unknown archive state'
        $script:content | Should -Match 'Teams shared channels'
        $script:content | Should -Match 'Incomplete Teams evidence'
        $script:content | Should -Match 'Group mailbox conversation evidence'
        $script:content | Should -Match 'Large mailbox content'
        $script:content | Should -Match 'Public folders'
        $script:content | Should -Match 'custom-domain application URI'

        @(
            'INACTIVE-NAME-MUST-NOT-LEAK',
            'archive-sentinel@contoso.com',
            'GROUP-NAME-MUST-NOT-LEAK',
            'group-sentinel@contoso.com',
            'PUBLIC-FOLDER-NAME-MUST-NOT-LEAK',
            'CUSTOMER-DECISION-MUST-NOT-LEAK',
            'TARGET-MAPPING-MUST-NOT-LEAK',
            'OWNER-MUST-NOT-LEAK',
            'TARGET-READINESS-MUST-NOT-LEAK',
            'WAVE-BREAKDOWN-MUST-NOT-LEAK'
        ) | ForEach-Object { $script:content | Should -Not -Match ([regex]::Escape($_)) }
    }

    It 'keeps every Markdown table to three columns or fewer' {
        $tableLines = @($script:content -split '\r?\n' | Where-Object { $_ -match '^\|' })
        $tableLines.Count | Should -BeGreaterThan 0
        foreach ($line in $tableLines) {
            @($line.ToCharArray() | Where-Object { $_ -eq '|' }).Count | Should -BeLessOrEqual 4
        }
    }

    It 'recommends phased delivery only as an overview when source sequencing makes it beneficial' {
        $script:content | Should -Match 'Recommended planning baseline:\*\* Phased migration'
        $script:content | Should -Match 'Why phased delivery may help'
        $script:content | Should -Match 'does not manufacture wave counts or memberships'
        $script:content | Should -Match 'Production batch acceptance'
        $script:content | Should -Not -Match 'WAVE-BREAKDOWN-MUST-NOT-LEAK'
    }

    It 'recommends a pilot plus single cutover for a small simple source and ignores seeded wave rows' {
        $simplePath = Join-Path $TestDrive 'simple-overview.md'
        $simpleContent = Export-TestMigrationOverview -Snapshot @{
            TenantInfo = [pscustomobject]@{ DisplayName = 'Simple Tenant'; DefaultDomain = 'simple.example' }
            MigrationExecutiveSummary = @(
                [pscustomobject]@{ Metric = 'Enabled member users'; Value = 100 },
                [pscustomobject]@{ Metric = 'Mailboxes'; Value = 100 },
                [pscustomobject]@{ Metric = 'Grand total data to migrate (GB)'; Value = 500 },
                [pscustomobject]@{ Metric = 'Teams'; Value = 10 },
                [pscustomobject]@{ Metric = 'SharePoint sites'; Value = 20 },
                [pscustomobject]@{ Metric = 'OneDrive sites'; Value = 100 }
            )
            MigrationScopeSummary = @(
                [pscustomobject]@{ Metric = 'SharePoint storage (GB)'; Value = 200 },
                [pscustomobject]@{ Metric = 'OneDrive storage (GB)'; Value = 200 }
            )
            ShareGateScopeSummary = @(
                [pscustomobject]@{ Workload = 'All collaboration sites (non-duplicated)'; ObjectCount = 120; TotalStorageGB = 400 }
            )
            MigrationWavePlan = @([pscustomobject]@{ Wave = 'SEEDED-WAVE-MUST-NOT-APPEAR'; CandidateObjectCount = 100 })
        } -OutputPath $simplePath

        $simpleContent | Should -Match 'Recommended planning baseline:\*\* Technical pilot followed by a single production cutover'
        $simpleContent | Should -Match 'Production cutover acceptance'
        $simpleContent | Should -Not -Match 'Why phased delivery may help|Production batch acceptance|SEEDED-WAVE-MUST-NOT-APPEAR'
        $simpleContent | Should -Match 'ShareGate Migrate Essentials'
    }

    It 'selects ShareGate Pro or Enterprise from source scale without emitting a detailed wave plan' {
        $proPath = Join-Path $TestDrive 'pro-overview.md'
        $proContent = Export-TestMigrationOverview -Snapshot @{
            TenantInfo = [pscustomobject]@{ DisplayName = 'Medium Tenant'; DefaultDomain = 'medium.example' }
            MigrationExecutiveSummary = @(
                [pscustomobject]@{ Metric = 'Enabled member users'; Value = 600 },
                [pscustomobject]@{ Metric = 'Mailboxes'; Value = 580 },
                [pscustomobject]@{ Metric = 'Teams'; Value = 20 },
                [pscustomobject]@{ Metric = 'SharePoint sites'; Value = 80 },
                [pscustomobject]@{ Metric = 'OneDrive sites'; Value = 600 }
            )
            MigrationScopeSummary = @([pscustomobject]@{ Metric = 'SharePoint storage (GB)'; Value = 500 })
            MigrationWavePlan = @([pscustomobject]@{ Wave = 'PRO-WAVE-DETAIL-MUST-NOT-APPEAR' })
        } -OutputPath $proPath

        $enterprisePath = Join-Path $TestDrive 'enterprise-overview.md'
        $enterpriseContent = Export-TestMigrationOverview -Snapshot @{
            TenantInfo = [pscustomobject]@{ DisplayName = 'Large Tenant'; DefaultDomain = 'large.example' }
            MigrationExecutiveSummary = @(
                [pscustomobject]@{ Metric = 'Enabled member users'; Value = 1201 },
                [pscustomobject]@{ Metric = 'Mailboxes'; Value = 1150 },
                [pscustomobject]@{ Metric = 'Teams'; Value = 100 },
                [pscustomobject]@{ Metric = 'SharePoint sites'; Value = 400 },
                [pscustomobject]@{ Metric = 'OneDrive sites'; Value = 1201 }
            )
            MigrationScopeSummary = @([pscustomobject]@{ Metric = 'SharePoint storage (GB)'; Value = 3000 })
            MigrationWavePlan = @([pscustomobject]@{ Wave = 'ENTERPRISE-WAVE-DETAIL-MUST-NOT-APPEAR' })
        } -OutputPath $enterprisePath

        $proContent | Should -Match 'ShareGate Migrate Pro'
        $proContent | Should -Match 'Recommended planning baseline:\*\* Phased migration'
        $proContent | Should -Not -Match 'PRO-WAVE-DETAIL-MUST-NOT-APPEAR'
        $enterpriseContent | Should -Match 'ShareGate Migrate Enterprise'
        $enterpriseContent | Should -Match 'Recommended planning baseline:\*\* Phased migration'
        $enterpriseContent | Should -Not -Match 'ENTERPRISE-WAVE-DETAIL-MUST-NOT-APPEAR'
    }

    It 'defers the production method when aggregate sizing evidence is absent' {
        $unknownPath = Join-Path $TestDrive 'unknown-overview.md'
        $unknownContent = Export-TestMigrationOverview -Snapshot @{
            TenantInfo = [pscustomobject]@{ DisplayName = 'Unknown Scope'; DefaultDomain = 'unknown.example' }
            MigrationWavePlan = @([pscustomobject]@{ Wave = 'FABRICATED-WAVE-MUST-NOT-APPEAR' })
        } -OutputPath $unknownPath

        $unknownContent | Should -Match 'Recommended planning baseline:\*\* Select the method during solution design'
        $unknownContent | Should -Match 'does not contain enough population or data-volume evidence'
        $unknownContent | Should -Match 'Production execution acceptance'
        $unknownContent | Should -Not -Match 'Production cutover acceptance|Production batch acceptance|FABRICATED-WAVE-MUST-NOT-APPEAR'
    }

    It 'identifies current tool choices, purchase categories, and the best-fit ShareGate tier' {
        $script:content | Should -Match 'Tooling and additional purchasing'
        $script:content | Should -Match 'Best-fit ShareGate tier from the source-only sizing proxy:\*\* ShareGate Migrate Essentials'
        $script:content | Should -Match 'Annual commercial subscription'
        $script:content | Should -Match 'Cross-Tenant User Data Migration add-on'
        $script:content | Should -Match 'Cross-Tenant Shared Data Migration SKU in 100 GB units'
        $script:content | Should -Match 'Currently documented for Enterprise Agreement customers'
        $script:content | Should -Match 'Microsoft-plus-ShareGate hybrid design'
        $script:content | Should -Match '12 mailbox license unit\(s\) and 3 User Migration Bundle unit\(s\)'
        $script:content | Should -Match 'Microsoft FastTrack cross-tenant service'
        $script:content | Should -Match 'ShareGate Protect is a separate governance/security product'
        $script:content | Should -Match 'sharegate\.com/pricing'
        $script:content | Should -Match 'migration-orchestrator-1-overview'
        $script:content | Should -Match 'Vendor-listed starting price: USD 5,995/year \(catalog reviewed 2026-08-14\)'
        $script:content | Should -Match 'Tooling catalog last reviewed: 2026-08-14'
    }

    It 'keeps reviewed vendor tiers and source links in a versioned tooling catalog' {
        Test-Path -LiteralPath $script:toolingCatalogPath | Should -BeTrue
        $catalog = Get-Content -LiteralPath $script:toolingCatalogPath -Raw | ConvertFrom-Json
        $catalog.SchemaVersion | Should -Be 1
        $catalog.LastReviewed | Should -Be '2026-08-14'
        (@($catalog.ShareGate.Plans | Select-Object -ExpandProperty Name) -join '|') | Should -Be 'ShareGate Migrate Essentials|ShareGate Migrate Pro|ShareGate Migrate Enterprise'
        (@($catalog.ShareGate.Plans | Select-Object -ExpandProperty StartingPriceUsdPerYear) -join '|') | Should -Be '5995|9995|17995'
        $catalog.ShareGate.PricingUrl | Should -Be 'https://sharegate.com/pricing'
        $catalog.Microsoft.MigrationOrchestrator.SourceUrl | Should -Match 'learn\.microsoft\.com'
    }

    It 'presents billable milestones without pretending the script collected progress input' {
        $script:content | Should -Match 'Suggested project milestones'
        $script:content | Should -Match 'suggested billable outcome points, not progress states inferred by the assessment script'
        $script:content | Should -Match 'Assessment readout and approach selection'
        $script:content | Should -Match 'Solution design and tooling procurement'
        $script:content | Should -Match 'Target foundation and cutover preparation'
        $script:content | Should -Match 'Pilot migration and acceptance'
        $script:content | Should -Match 'Post-migration validation and hypercare exit'
        $script:content | Should -Match 'Source decommission and compliance closeout'
    }

    It 'does not grow linearly or expose names when full object inventories are supplied' {
        $oneMailbox = @([pscustomobject]@{ DisplayName = 'Inventory Sentinel 1'; PrimarySmtpAddress = 'sentinel1@example.com'; MailboxSizeGB = 1 })
        $manyMailboxes = @(1..200 | ForEach-Object {
                [pscustomobject]@{ DisplayName = "Inventory Sentinel $_"; PrimarySmtpAddress = "sentinel$_@example.com"; MailboxSizeGB = 1 }
            })
        $oneGroup = @([pscustomobject]@{ DisplayName = 'Group Inventory Sentinel 1'; MailboxSizeGB = 1; MailboxEvidenceStatus = 'Measured data' })
        $manyGroups = @(1..200 | ForEach-Object {
                [pscustomobject]@{ DisplayName = "Group Inventory Sentinel $_"; MailboxSizeGB = 1; MailboxEvidenceStatus = 'Measured data' }
            })

        $singleContent = Export-TestMigrationOverview -Snapshot @{
            TenantInfo = [pscustomobject]@{ DisplayName = 'Single Inventory'; DefaultDomain = 'single.example' }
            MigrationExecutiveSummary = @([pscustomobject]@{ Metric = 'Enabled member users'; Value = 100 })
            AllMailboxes = $oneMailbox
            GroupWorkloadReconciliation = $oneGroup
        } -OutputPath (Join-Path $TestDrive 'single-inventory.md')
        $manyContent = Export-TestMigrationOverview -Snapshot @{
            TenantInfo = [pscustomobject]@{ DisplayName = 'Many Inventory'; DefaultDomain = 'many.example' }
            MigrationExecutiveSummary = @([pscustomobject]@{ Metric = 'Enabled member users'; Value = 100 })
            AllMailboxes = $manyMailboxes
            GroupWorkloadReconciliation = $manyGroups
        } -OutputPath (Join-Path $TestDrive 'many-inventory.md')

        $singleLineCount = @($singleContent -split '\r?\n').Count
        $manyLineCount = @($manyContent -split '\r?\n').Count
        [math]::Abs($manyLineCount - $singleLineCount) | Should -BeLessOrEqual 2
        $manyContent | Should -Not -Match 'Inventory Sentinel 199|sentinel199@example\.com|Group Inventory Sentinel 199'
    }

    It 'routes only the Presales policy to this artifact and retains the legacy exporter' {
        $script:pipelineSource | Should -Match '\$isPresalesScopeQuestionnaire'
        $script:pipelineSource | Should -Match 'Export-PresalesMigrationScopeQuestionnaireMarkdown'
        $script:pipelineSource | Should -Match '-MigrationScopeQuestionnaire\.md'
        $script:pipelineSource | Should -Match 'Export-TenantToTenantQuestionnaireMarkdown'
        $script:pipelineSource | Should -Match '-TenantToTenantQuestionnaire\.md'
    }
}
