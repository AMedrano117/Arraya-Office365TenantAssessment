Describe 'Presales migration scope questionnaire' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:exporterPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Export-PresalesMigrationScopeQuestionnaireMarkdown.ps1'
        $script:pipelinePath = Join-Path $script:repoRoot 'src\scripts\reporting\Invoke-M365TenantAssessmentExportPipeline.ps1'
        . $script:exporterPath

        $snapshot = @{
            TenantInfo = [pscustomobject]@{
                DisplayName   = 'Contoso'
                TenantId      = 'source-tenant-id'
                DefaultDomain = 'contoso.com'
            }
            InactiveMailboxDetails = @(
                [pscustomobject]@{
                    DisplayName           = 'Inactive With Archive'
                    PrimarySmtpAddress     = 'archive@contoso.com'
                    RecipientTypeDetails  = 'InactiveMailbox'
                    IsInactiveMailbox     = $true
                    MailboxSizeGB          = 22
                    DeletedItemsGB         = 3
                    ArchiveStatus          = 'Active'
                    ArchiveSizeGB          = 44
                    ArchiveDeletedItemsGB  = 5
                    TotalDataToMigrateGB   = 74
                    LitigationHoldEnabled = $true
                    RetentionPolicy        = 'Retain'
                },
                [pscustomobject]@{
                    DisplayName          = 'Inactive Unknown Archive'
                    PrimarySmtpAddress    = 'unknown@contoso.com'
                    RecipientTypeDetails = 'InactiveMailbox'
                    IsInactiveMailbox    = $true
                    MailboxSizeGB         = 8
                    ArchiveStatus        = 'Unknown'
                },
                [pscustomobject]@{
                    DisplayName          = 'Inactive Unavailable Archive'
                    PrimarySmtpAddress    = 'unavailable@contoso.com'
                    RecipientTypeDetails = 'InactiveMailbox'
                    IsInactiveMailbox    = $true
                    MailboxSizeGB         = 7
                    ArchiveStatus        = 'Unavailable'
                },
                [pscustomobject]@{
                    DisplayName          = 'Inactive Blank Archive State'
                    PrimarySmtpAddress    = 'blank@contoso.com'
                    RecipientTypeDetails = 'InactiveMailbox'
                    IsInactiveMailbox    = $true
                    MailboxSizeGB         = 6
                },
                [pscustomobject]@{
                    DisplayName          = 'Inactive Conflicting Archive Evidence'
                    PrimarySmtpAddress    = 'conflict@contoso.com'
                    RecipientTypeDetails = 'InactiveMailbox'
                    IsInactiveMailbox    = $true
                    MailboxSizeGB         = 9
                    HasArchive           = $false
                    ArchiveStatus        = 'Active'
                    ArchiveSizeGB        = 12
                }
            )
            BitTitanLicenseSummary = @(
                [pscustomobject]@{ Section = 'Licensing'; Metric = 'Mailbox Migration licenses'; Value = 12 },
                [pscustomobject]@{ Section = 'Licensing'; Metric = 'User Migration Bundles'; Value = 3 }
            )
            BitTitanLicenseDetail = @(
                [pscustomobject]@{ ObjectType = 'Mailbox'; Identity = 'user@contoso.com'; BundleEligible = $true }
            )
            GroupWorkloadReconciliation = @(
                [pscustomobject]@{
                    DisplayName       = 'Operations'
                    PrimarySmtpAddress = 'operations@contoso.com'
                    MailboxSizeGB      = 2
                    SiteStorageGB     = 18
                    IsTeam            = $true
                    Classification    = 'Mailbox and site content'
                },
                [pscustomobject]@{
                    DisplayName           = 'Unknown Group Mailbox'
                    PrimarySmtpAddress     = 'unknown-group@contoso.com'
                    MailboxSizeGB          = $null
                    MailboxEvidenceStatus  = 'Needs Data'
                    SiteStorageGB          = 4
                    IsTeam                 = $false
                    Classification         = 'Group, site only - mail evidence unavailable'
                },
                [pscustomobject]@{
                    DisplayName       = 'Legacy Snapshot Group'
                    PrimarySmtpAddress = 'legacy-group@contoso.com'
                    MailboxSizeGB      = $null
                    SiteStorageGB      = 0
                    IsTeam            = $false
                    Classification    = 'Empty shell'
                }
            )
            AllTeams = @(
                [pscustomobject]@{
                    DisplayName = 'Operations'; SharedChannelCount = 1; MemberCount = 10
                    ChannelInventoryStatus = 'Measured data'; MemberInventoryStatus = 'Measured data'
                },
                [pscustomobject]@{
                    DisplayName = 'Legacy Failed Team'; SharedChannelCount = 0; MemberCount = 0
                    ChannelInventoryStatus = 'Needs Data'; MemberInventoryStatus = 'Needs Data'
                }
            )
            SharePoint = @([pscustomobject]@{ Title = 'Operations'; StorageUsedGB = 18 })
            OneDrive = @([pscustomobject]@{ Title = 'User OneDrive'; StorageUsedGB = 10 })
            MigrationScopeDecisions = @(
                [pscustomobject]@{
                    DecisionKey           = 'BT-06'
                    CustomerConfirmedValue = 'Evaluate as an alternative'
                    Status                = 'Open'
                    InScope               = $true
                    TargetMapping         = 'Operations Team'
                    MigrationTool         = 'Optional BitTitan TMB'
                    DecisionOwner         = 'Solution Engineer'
                    DueDate               = '2026-08-20'
                    Notes                 = 'Do not duplicate ShareGate scope.'
                },
                [pscustomobject]@{
                    DecisionKey   = 'GROUP:operations@contoso.com'
                    Status        = 'Open'
                    InScope       = $false
                    TargetMapping = 'Operations-New'
                    MigrationTool = 'Retain source conversations'
                    DecisionOwner = 'Messaging Lead'
                    Notes         = 'Site content remains in ShareGate scope.'
                },
                [pscustomobject]@{
                    DecisionKey           = 'MAILBOX:archive@contoso.com'
                    CustomerConfirmedValue = 'Restore and migrate primary plus archive'
                    Status                = 'Approved'
                    InScope               = $true
                    TargetMapping         = 'archive@target.example'
                    MigrationTool         = 'BitTitan MigrationWiz'
                    DecisionOwner         = 'Compliance Lead'
                    Notes                 = 'Preserve hold evidence before restoration.'
                }
            )
            MigrationQuoteReadiness = @(
                [pscustomobject]@{ RowType = 'Summary'; ReadinessLevel = 'ROM'; Status = 'ROM'; DecisionKey = 'QR-01' }
                [pscustomobject]@{ RowType = 'Check'; Status = 'Blocker'; DecisionKey = 'FLAG:custom-domain-app'; Requirement = 'Custom-domain application dependency' }
            )
            MigrationComplexityFlags = @(
                [pscustomobject]@{ Status = 'Blocker'; Item = 'Custom-domain application dependency' }
            )
        }

        $script:outputPath = Join-Path $TestDrive 'Contoso-MigrationScopeQuestionnaire.md'
        Export-PresalesMigrationScopeQuestionnaireMarkdown -TenantStatsHash $snapshot -Path $script:outputPath -PreparedBy 'Unit Test' -AsOfDate ([datetime]'2026-08-12')
        $script:content = Get-Content -Raw -Path $script:outputPath
        $script:pipelineSource = Get-Content -Raw -Path $script:pipelinePath
    }

    It 'separates discovered evidence from customer decisions and omits pricing' {
        Test-Path $script:outputPath | Should -BeTrue
        $script:content | Should -Match 'Discovered value'
        $script:content | Should -Match 'Customer-confirmed value'
        $script:content | Should -Match 'Decision status'
        $script:content | Should -Match 'Current Quote Readiness:\*\* ROM'
        $script:content | Should -Match 'Pricing is intentionally omitted'
        $script:content | Should -Not -Match 'Unit Cost|Extended Cost|License Price'
    }

    It 'makes inactive mailbox restoration and archive evidence explicit' {
        $script:content | Should -Match 'inactive mailboxes will be restored or recovered'
        $script:content | Should -Match 'Inactive mailbox and archive confirmation'
        $script:content | Should -Match 'Inactive With Archive'
        $script:content | Should -Match 'archive@contoso\.com'
        $script:content | Should -Match '\| Yes \| Active \| 44 \| 5 \|'
        $script:content | Should -Match 'Inactive Unknown Archive'
        $unknownLine = @($script:content -split '\r?\n' | Where-Object { $_ -match 'unknown@contoso\.com' } | Select-Object -First 1)
        $unavailableLine = @($script:content -split '\r?\n' | Where-Object { $_ -match 'unavailable@contoso\.com' } | Select-Object -First 1)
        $blankLine = @($script:content -split '\r?\n' | Where-Object { $_ -match 'blank@contoso\.com' } | Select-Object -First 1)
        $unknownLine | Should -Match '\| Unknown - validate \| Unknown \|'
        $unavailableLine | Should -Match '\| Unknown - validate \| Unavailable \|'
        $blankLine | Should -Match '\| Unknown - validate \|  \|'
        $conflictLine = @($script:content -split '\r?\n' | Where-Object { $_ -match 'conflict@contoso\.com' } | Select-Object -First 1)
        $conflictLine | Should -Match '\| Yes \| Active \| 12 \|'
    }

    It 'maps inactive mailbox decisions by the object-level MAILBOX key' {
        $inactiveDecisionLine = @($script:content -split '\r?\n' | Where-Object { $_ -match '^\| MAILBOX:archive@contoso\.com \|' } | Select-Object -First 1)
        $inactiveDecisionLine | Should -Not -BeNullOrEmpty
        $inactiveDecisionLine | Should -Match 'Restore and migrate primary plus archive'
        $inactiveDecisionLine | Should -Match '\| Yes \| Approved \| archive@target\.example \| BitTitan MigrationWiz \|'
        $inactiveDecisionLine | Should -Match 'Compliance Lead; Preserve hold evidence before restoration\.'
    }

    It 'includes optional TMB guardrails and current MigrationWiz readiness questions' {
        $script:content | Should -Match 'Optional Tenant Migration Bundle decision guide'
        $script:content | Should -Match 'one UMB plus one FCL'
        $script:content | Should -Match 'one Team or SharePoint library up to 100 GB'
        $script:content | Should -Match 'TMB does not migrate public folders'
        $script:content | Should -Match 'Teams Private Chat \(PCH\) uses a separate Collaboration \(Private Chats\) project'
        $script:content | Should -Match 'validate current capability, limitations, and licensing entitlement'
        $script:content | Should -Match 'rather than assuming TMB/UMB inclusion or exclusion'
        $script:content | Should -Match '25603590557979-Teams-Private-Chat-Migration-Guide'
        $script:content | Should -Match 'ShareGate remains the default tool'
        $script:content | Should -Match 'EWSAllowedAppIDs'
        $script:content | Should -Match '2026-10-01'
        $script:content | Should -Match '2027-04-01'
    }

    It 'does not treat failed Team channel or member enrichment as measured zero' {
        $sg03Line = @($script:content -split '\r?\n' | Where-Object { $_ -match '^\| SG-03 \|' } | Select-Object -First 1)
        $sg03Line | Should -Match 'Needs Data: channel inventory unavailable for 1 of 2 Team\(s\)'
        $sg03Line | Should -Match 'member/guest inventory unavailable for 1'
        $sg03Line | Should -Match '\| Medium \|'
        $sg03Line | Should -Not -Match '0 Team/group row\(s\) have a shared-channel signal'
    }

    It 'maps ShareGate endpoint readiness to the SG-04 blocker decision key' {
        $sg04Line = @($script:content -split '\r?\n' | Where-Object { $_ -match '^\| SG-04 \|' } | Select-Object -First 1)
        $sg04Line | Should -Not -BeNullOrEmpty
        $sg04Line | Should -Match 'source and destination connectivity'
        $sg04Line | Should -Match 'admin roles'
        $sg04Line | Should -Match 'application consent'
        $sg04Line | Should -Match 'representative test access'
    }

    It 'merges customer decisions by DecisionKey without replacing discovered evidence' {
        $script:content | Should -Match '\| BT-06 \|'
        $script:content | Should -Match 'Evaluate as an alternative'
        $script:content | Should -Match 'Optional BitTitan TMB'
        $script:content | Should -Match 'Solution Engineer'
        $script:content | Should -Match 'Do not duplicate ShareGate scope'
    }

    It 'does not double-count complexity blockers already projected into quote readiness' {
        $qrLine = @($script:content -split '\r?\n' | Where-Object { $_ -match '^\| QR-01 \|' } | Select-Object -First 1)
        $qrLine | Should -Match 'blocker or missing-data rows: 1;'
    }

    It 'summarizes reconciled group mailbox evidence even without a separate GroupMailboxes table' {
        $bt04Line = @($script:content -split '\r?\n' | Where-Object { $_ -match '^\| BT-04 \|' } | Select-Object -First 1)
        $bt04Line | Should -Match 'measured mail data: 1; measured empty: 0; needs data: 2'
        $bt04Line | Should -Not -Match '0 group mailbox inventory row'
    }

    It 'uses the seeded DecisionPrompt as the cross-artifact question contract' {
        $contractPath = Join-Path $TestDrive 'canonical-decision-contract.md'
        Export-PresalesMigrationScopeQuestionnaireMarkdown -TenantStatsHash @{
            TenantInfo = [pscustomobject]@{ DisplayName = 'Contract Test'; TenantId = 'source-id' }
            MigrationScopeDecisions = @(
                [pscustomobject]@{
                    DecisionKey = 'BT-07'; DecisionPrompt = 'Canonical BT-07 prompt supplied by the planning register.'; Status = 'Needs Input'
                }
            )
        } -Path $contractPath -PreparedBy 'Unit Test' -AsOfDate ([datetime]'2026-08-12')

        $contractContent = Get-Content -LiteralPath $contractPath -Raw
        $bt07Line = @($contractContent -split '\r?\n' | Where-Object { $_ -match '^\| BT-07 \|' } | Select-Object -First 1)
        $bt07Line | Should -Match 'Canonical BT-07 prompt supplied by the planning register\.'
        $bt07Line | Should -Not -Match 'If TMB is selected'
    }

    It 'provides an object-level group mailbox conversation disposition register' {
        $script:content | Should -Match 'Group mailbox conversation disposition'
        $script:content | Should -Match 'Report only / customer decision'
        $script:content | Should -Match 'GROUP:operations@contoso\.com'
        $script:content | Should -Match 'operations@contoso\.com'
        $script:content | Should -Match 'Mailbox and site content'
        $script:content | Should -Match 'Operations-New'
        $script:content | Should -Match 'Retain source conversations'
        $script:content | Should -Match 'Site content remains in ShareGate scope'
    }

    It 'does not treat missing group mailbox size evidence as an empty mailbox' {
        $explicitUnknownLine = @($script:content -split '\r?\n' | Where-Object { $_ -match 'unknown-group@contoso\.com' } | Select-Object -First 1)
        $legacyUnknownLine = @($script:content -split '\r?\n' | Where-Object { $_ -match 'legacy-group@contoso\.com' } | Select-Object -First 1)
        $explicitUnknownLine | Should -Match '\| Unknown / needs data \| Needs Data \|'
        $explicitUnknownLine | Should -Match 'Group, site only - mail evidence unavailable'
        $explicitUnknownLine | Should -Not -Match 'mail evidence unavailable \(mailbox evidence unknown\)'
        $legacyUnknownLine | Should -Match '\| Unknown / needs data \| Unknown / needs data \|'
        $legacyUnknownLine | Should -Match 'Empty shell \(mailbox evidence unknown\)'
    }

    It 'routes only the Presales policy to the new artifact and retains the legacy exporter' {
        $script:pipelineSource | Should -Match '\$isPresalesScopeQuestionnaire'
        $script:pipelineSource | Should -Match 'Export-PresalesMigrationScopeQuestionnaireMarkdown'
        $script:pipelineSource | Should -Match '-MigrationScopeQuestionnaire\.md'
        $script:pipelineSource | Should -Match 'Export-TenantToTenantQuestionnaireMarkdown'
        $script:pipelineSource | Should -Match '-TenantToTenantQuestionnaire\.md'
    }
}
