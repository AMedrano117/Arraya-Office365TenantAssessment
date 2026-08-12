Describe 'Migration scope reconciliation' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        Import-Module -Name (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop

        # The reconciliation and scope builders live in the legacy collector script, which
        # cannot be dot-sourced without a live tenant. Extract just the function definitions
        # under test, plus the small helpers they call, into this session.
        $script:collectorPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:collectorPath, [ref]$tokens, [ref]$parseErrors)

        $wantedFunctions = @(
            'Get-AssessmentNormalizedSiteUrl'
            'Get-AssessmentExportTableArray'
            'Get-AssessmentExportPropertyValue'
            'Set-AssessmentExportProperty'
            'Convert-AssessmentExportSizeToGb'
            'Convert-AssessmentExportByteCountToGb'
            'Convert-DataSizeToBytes'
            'Test-AssessmentMailboxHasArchiveEvidence'
            'Get-AssessmentTeamEnrichmentEvidence'
            'Update-AssessmentTeamEnrichmentEvidenceRows'
            'Ensure-MigrationGroupMailboxInventory'
            'Get-AssessmentMailboxPlanningRows'
            'Update-MigrationGroupWorkloadReconciliation'
            'Get-MigrationShareGateScopeData'
            'Update-MigrationPlanningRegisters'
        )

        $definitions = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $true) | Where-Object { $wantedFunctions -contains $_.Name }

        foreach ($definition in $definitions) {
            . ([scriptblock]::Create($definition.Extent.Text))
        }

        $script:AssessmentBitTitanLicenseModel = Get-ArrayaBitTitanLicenseModel

        function script:New-TestSite {
            param([string]$Title, [string]$Url, $StorageGB)
            [pscustomobject]@{ Title = $Title; Url = $Url; StorageUsedGB = $StorageGB }
        }

        function script:New-TestGroup {
            param([string]$Name, [string]$Smtp, [string]$SiteUrl)
            [pscustomobject]@{ DisplayName = $Name; PrimarySmtpAddress = $Smtp; SharePointSiteUrl = $SiteUrl; ExchangeGuid = $Smtp }
        }

        function script:New-TestTeam {
            param(
                [string]$Name,
                [string]$SiteUrl,
                [int]$SharedChannelCount = 0,
                [string]$ChannelInventoryStatus = 'Measured data',
                [string]$MemberInventoryStatus = 'Measured data'
            )
            [pscustomobject]@{
                DisplayName = $Name
                SharePointSiteUrl = $SiteUrl
                TotalChannels = [Math]::Max(1, $SharedChannelCount)
                SharedChannelCount = $SharedChannelCount
                ChannelInventoryStatus = $ChannelInventoryStatus
                MemberCount = 5
                GuestCount = 0
                MemberInventoryStatus = $MemberInventoryStatus
                'SiteSize-GB' = 1
            }
        }

        function script:New-TestGroupMailbox {
            param([string]$Name, [string]$Smtp, $SizeGB)
            [pscustomobject]@{ DisplayName = $Name; PrimarySmtpAddress = $Smtp; ExchangeGuid = $Smtp; TotalDataToMigrateGB = $SizeGB; MailboxSizeGB = $SizeGB }
        }

        function script:New-TestStore {
            # A tenant exercising every classification the reconciliation has to distinguish,
            # including a URL that differs only by case and trailing slash.
            @{
                UnifiedGroups = @(
                    New-TestGroup -Name 'Team with mail and site' -Smtp 'team1@contoso.com' -SiteUrl 'https://contoso.sharepoint.com/sites/team1'
                    New-TestGroup -Name 'Team with empty mailbox' -Smtp 'team2@contoso.com' -SiteUrl 'https://contoso.sharepoint.com/sites/team2'
                    New-TestGroup -Name 'Plain group with site' -Smtp 'group1@contoso.com' -SiteUrl 'https://CONTOSO.sharepoint.com/sites/group1/'
                    New-TestGroup -Name 'Empty shell group' -Smtp 'group2@contoso.com' -SiteUrl 'https://contoso.sharepoint.com/sites/group2'
                )
                AllTeams = @(
                    New-TestTeam -Name 'Team with mail and site' -SiteUrl 'https://contoso.sharepoint.com/sites/team1' -SharedChannelCount 2
                    New-TestTeam -Name 'Team with empty mailbox' -SiteUrl 'https://contoso.sharepoint.com/sites/team2'
                )
                SharePoint = @(
                    New-TestSite -Title 'Team 1 site' -Url 'https://contoso.sharepoint.com/sites/team1' -StorageGB 10
                    New-TestSite -Title 'Team 2 site' -Url 'https://contoso.sharepoint.com/sites/team2' -StorageGB 5
                    New-TestSite -Title 'Group 1 site' -Url 'https://contoso.sharepoint.com/sites/group1' -StorageGB 20
                    New-TestSite -Title 'Group 2 site' -Url 'https://contoso.sharepoint.com/sites/group2' -StorageGB 0
                    New-TestSite -Title 'Communication site' -Url 'https://contoso.sharepoint.com/sites/comms' -StorageGB 40
                    New-TestSite -Title 'Classic site' -Url 'https://contoso.sharepoint.com/sites/classic' -StorageGB 25
                )
                GroupMailboxes = @(
                    New-TestGroupMailbox -Name 'Team with mail and site' -Smtp 'team1@contoso.com' -SizeGB 8
                    New-TestGroupMailbox -Name 'Team with empty mailbox' -Smtp 'team2@contoso.com' -SizeGB 0
                    New-TestGroupMailbox -Name 'Plain group with site' -Smtp 'group1@contoso.com' -SizeGB 3
                    New-TestGroupMailbox -Name 'Empty shell group' -Smtp 'group2@contoso.com' -SizeGB 0
                )
            }
        }
    }

    Context 'Site attribution' {
        BeforeAll {
            $script:store = New-TestStore
            $script:result = Update-MigrationGroupWorkloadReconciliation -TenantStatsStore $script:store
            $script:rows = @($script:store['GroupWorkloadReconciliation'])
        }

        It 'produces exactly one row per Microsoft 365 Group' {
            $script:rows.Count | Should -Be 4
        }

        It 'attributes every site URL to exactly one owner' {
            # This is what stops the ShareGate total double counting: a site claimed by a group
            # must not also appear in the standalone bucket.
            $claimedCount = $script:result.ClaimedSiteUrls.Count
            $standaloneCount = @($script:result.StandaloneSiteRows).Count
            ($claimedCount + $standaloneCount) | Should -Be 6
        }

        It 'leaves only genuinely standalone sites unclaimed' {
            $standaloneTitles = @($script:result.StandaloneSiteRows | ForEach-Object { $_.Title } | Sort-Object)
            $standaloneTitles | Should -Be @('Classic site', 'Communication site')
        }

        It 'matches sites whose URL differs only by case and trailing slash' {
            $groupRow = @($script:rows | Where-Object { $_.DisplayName -eq 'Plain group with site' } | Select-Object -First 1)
            $groupRow.SiteStorageGB | Should -Be 20
        }

        It 'balances group-site storage plus standalone storage against the SharePoint total' {
            $groupSiteTotal = @($script:rows | ForEach-Object { [double]($_.SiteStorageGB) } | Measure-Object -Sum).Sum
            $standaloneTotal = @($script:result.StandaloneSiteRows | ForEach-Object { [double]$_.StorageUsedGB } | Measure-Object -Sum).Sum
            $sharePointTotal = @($script:result.SharePointRows | ForEach-Object { [double]$_.StorageUsedGB } | Measure-Object -Sum).Sum

            ($groupSiteTotal + $standaloneTotal) | Should -Be $sharePointTotal
        }

        It 'splits Teams, other group, standalone, and OneDrive storage while preserving one non-duplicated total' {
            $oneDrives = @(
                New-TestSite -Title 'User OneDrive' -Url 'https://contoso-my.sharepoint.com/personal/user' -StorageGB 30
            )
            $scope = Get-MigrationShareGateScopeData -Reconciliation $script:result -OneDriveRows $oneDrives
            $byWorkload = @{}
            foreach ($row in @($scope.SummaryRows)) { $byWorkload[[string]$row.Workload] = $row }

            $byWorkload['Teams-connected SharePoint'].ObjectCount | Should -Be 2
            $byWorkload['Teams-connected SharePoint'].TotalStorageGB | Should -Be 15
            $byWorkload['Microsoft 365 Group SharePoint (non-Team)'].ObjectCount | Should -Be 2
            $byWorkload['Microsoft 365 Group SharePoint (non-Team)'].TotalStorageGB | Should -Be 20
            $byWorkload['Standalone SharePoint'].ObjectCount | Should -Be 2
            $byWorkload['Standalone SharePoint'].TotalStorageGB | Should -Be 65
            $byWorkload['SharePoint (total)'].TotalStorageGB | Should -Be 100
            $byWorkload['OneDrive'].TotalStorageGB | Should -Be 30
            $byWorkload['All collaboration sites (non-duplicated)'].ObjectCount | Should -Be 7
            $byWorkload['All collaboration sites (non-duplicated)'].TotalStorageGB | Should -Be 130

            $componentTotal = [double]$byWorkload['Teams-connected SharePoint'].TotalStorageGB +
                [double]$byWorkload['Microsoft 365 Group SharePoint (non-Team)'].TotalStorageGB +
                [double]$byWorkload['Standalone SharePoint'].TotalStorageGB +
                [double]$byWorkload['OneDrive'].TotalStorageGB
            $componentTotal | Should -Be $byWorkload['All collaboration sites (non-duplicated)'].TotalStorageGB
        }
    }

    Context 'Teams are a subset of groups' {
        BeforeAll {
            $script:store = New-TestStore
            $script:result = Update-MigrationGroupWorkloadReconciliation -TenantStatsStore $script:store
            $script:rows = @($script:store['GroupWorkloadReconciliation'])
        }

        It 'never reports more Teams than there are Teams collected' {
            $teamRowCount = @($script:rows | Where-Object { $_.IsTeam -eq $true }).Count
            $teamRowCount | Should -BeLessOrEqual @($script:result.TeamRows).Count
        }

        It 'marks only the groups that actually back a Team' {
            $teamNames = @($script:rows | Where-Object { $_.IsTeam -eq $true } | ForEach-Object { $_.DisplayName } | Sort-Object)
            $teamNames | Should -Be @('Team with empty mailbox', 'Team with mail and site')
        }

        It 'flags shared channels on the backing group row' {
            $row = @($script:rows | Where-Object { $_.DisplayName -eq 'Team with mail and site' } | Select-Object -First 1)
            $row.HasSharedChannels | Should -BeTrue
        }

        It 'keeps failed Team enrichment unknown instead of turning zero into measured evidence' {
            $store = New-TestStore
            $store.AllTeams = @(
                New-TestTeam -Name 'Team with mail and site' -SiteUrl 'https://contoso.sharepoint.com/sites/team1' -SharedChannelCount 2
                [pscustomobject]@{
                    DisplayName = 'Team with empty mailbox'
                    SharePointSiteUrl = 'https://contoso.sharepoint.com/sites/team2'
                    TotalChannels = 0
                    SharedChannelCount = 0
                    ChannelInventoryStatus = 'Needs Data'
                    MemberCount = 0
                    GuestCount = 0
                    MemberInventoryStatus = 'Needs Data'
                }
            )

            $null = Update-MigrationGroupWorkloadReconciliation -TenantStatsStore $store
            $row = @($store.GroupWorkloadReconciliation | Where-Object DisplayName -eq 'Team with empty mailbox' | Select-Object -First 1)

            $row.ChannelInventoryStatus | Should -Be 'Needs Data'
            $row.MemberInventoryStatus | Should -Be 'Needs Data'
            $row.HasSharedChannels | Should -BeNullOrEmpty
        }

        It 'reports Teams that did not join to a collected site' {
            $script:result.UnmatchedTeamCount | Should -Be 0
        }
    }

    Context 'Legacy Teams evidence normalization' {
        It 'clears unqualified legacy zeroes while preserving explicit measured empty evidence' {
            $store = @{
                AllTeams = @(
                    [pscustomobject]@{
                        DisplayName = 'Legacy failed Team'
                        TotalChannels = 0; PublicChannelCount = 0; PrivateChannelCount = 0; SharedChannelCount = 0
                        MemberCount = 0; GuestCount = 0
                    }
                    [pscustomobject]@{
                        DisplayName = 'Measured empty Team'
                        TotalChannels = 0; PublicChannelCount = 0; PrivateChannelCount = 0; SharedChannelCount = 0
                        ChannelInventoryStatus = 'Measured empty'
                        MemberCount = 0; GuestCount = 0; MemberInventoryStatus = 'Measured empty'
                    }
                    [pscustomobject]@{
                        DisplayName = 'Legacy measured Team'
                        TotalChannels = 4; PublicChannelCount = 3; PrivateChannelCount = 0; SharedChannelCount = 1
                        MemberCount = 12; GuestCount = 0
                    }
                )
            }

            $summary = Update-AssessmentTeamEnrichmentEvidenceRows -TenantStatsStore $store
            $legacyFailed = @($store.AllTeams | Where-Object DisplayName -eq 'Legacy failed Team' | Select-Object -First 1)
            $measuredEmpty = @($store.AllTeams | Where-Object DisplayName -eq 'Measured empty Team' | Select-Object -First 1)
            $legacyMeasured = @($store.AllTeams | Where-Object DisplayName -eq 'Legacy measured Team' | Select-Object -First 1)

            $summary.TeamCount | Should -Be 3
            $summary.ChannelNeedsDataCount | Should -Be 1
            $summary.MemberNeedsDataCount | Should -Be 1

            $legacyFailed.ChannelInventoryStatus | Should -Be 'Needs Data'
            $legacyFailed.MemberInventoryStatus | Should -Be 'Needs Data'
            $legacyFailed.TotalChannels | Should -BeNullOrEmpty
            $legacyFailed.SharedChannelCount | Should -BeNullOrEmpty
            $legacyFailed.MemberCount | Should -BeNullOrEmpty
            $legacyFailed.GuestCount | Should -BeNullOrEmpty

            $measuredEmpty.ChannelInventoryStatus | Should -Be 'Measured empty'
            $measuredEmpty.MemberInventoryStatus | Should -Be 'Measured empty'
            $measuredEmpty.TotalChannels | Should -Be 0
            $measuredEmpty.SharedChannelCount | Should -Be 0
            $measuredEmpty.MemberCount | Should -Be 0
            $measuredEmpty.GuestCount | Should -Be 0

            $legacyMeasured.ChannelInventoryStatus | Should -Be 'Measured data'
            $legacyMeasured.MemberInventoryStatus | Should -Be 'Measured data'
            $legacyMeasured.TotalChannels | Should -Be 4
            $legacyMeasured.SharedChannelCount | Should -Be 1
            $legacyMeasured.MemberCount | Should -Be 12
        }
    }

    Context 'Classification and tool ownership' {
        BeforeAll {
            $script:store = New-TestStore
            $null = Update-MigrationGroupWorkloadReconciliation -TenantStatsStore $script:store
            $script:rows = @($script:store['GroupWorkloadReconciliation'])
        }

        It 'classifies a Team carrying both mail and site content' {
            $row = @($script:rows | Where-Object { $_.DisplayName -eq 'Team with mail and site' } | Select-Object -First 1)
            $row.Classification | Should -Be 'Team with mail and site'
            $row.MailboxScope | Should -Be 'BitTitan (optional)'
            $row.SiteScope | Should -Be 'ShareGate'
        }

        It 'classifies a Team whose mailbox was never used as site-only' {
            $row = @($script:rows | Where-Object { $_.DisplayName -eq 'Team with empty mailbox' } | Select-Object -First 1)
            $row.Classification | Should -Be 'Team, site only'
            $row.MailboxEvidenceStatus | Should -Be 'Measured empty'
            $row.MailboxEvidenceSource | Should -Be 'GroupMailboxes'
            $row.MailboxSizeGB | Should -Be 0
            $row.MailboxScope | Should -Be 'None'
            $row.SiteScope | Should -Be 'ShareGate'
        }

        It 'classifies a plain group that is not a Team' {
            $row = @($script:rows | Where-Object { $_.DisplayName -eq 'Plain group with site' } | Select-Object -First 1)
            $row.IsTeam | Should -BeFalse
            $row.Classification | Should -Be 'Group with mail and site'
        }

        It 'classifies a group with neither mail nor site content as an empty shell' {
            $row = @($script:rows | Where-Object { $_.DisplayName -eq 'Empty shell group' } | Select-Object -First 1)
            $row.Classification | Should -Be 'Empty shell'
            $row.MailboxScope | Should -Be 'None'
            $row.SiteScope | Should -Be 'None'
        }

        It 'never assigns the same half of a group to both migration tools' {
            foreach ($row in $script:rows) {
                $row.MailboxScope | Should -BeIn @('BitTitan (optional)', 'None', 'Needs Data')
                $row.SiteScope | Should -BeIn @('ShareGate', 'None')
            }
        }
    }

    Context 'Degraded input' {
        It 'returns an empty reconciliation rather than throwing when nothing was collected' {
            $store = @{}
            $result = Update-MigrationGroupWorkloadReconciliation -TenantStatsStore $store

            @($store['GroupWorkloadReconciliation']).Count | Should -Be 0
            @($result.StandaloneSiteRows).Count | Should -Be 0
        }

        It 'handles a group whose site URL matches nothing that was collected' {
            $store = @{
                UnifiedGroups = @(New-TestGroup -Name 'Orphan' -Smtp 'orphan@contoso.com' -SiteUrl 'https://contoso.sharepoint.com/sites/missing')
                SharePoint    = @(New-TestSite -Title 'Other' -Url 'https://contoso.sharepoint.com/sites/other' -StorageGB 5)
            }
            $null = Update-MigrationGroupWorkloadReconciliation -TenantStatsStore $store

            $row = @($store['GroupWorkloadReconciliation'] | Select-Object -First 1)
            $row.SiteStorageGB | Should -BeNullOrEmpty
            $row.Notes | Should -Match 'did not match|inventory/statistics'
        }

        It 'marks an older snapshot without group mailbox evidence as Needs Data' {
            $store = @{
                UnifiedGroups = @(New-TestGroup -Name 'Legacy Group' -Smtp 'legacy@contoso.com' -SiteUrl 'https://contoso.sharepoint.com/sites/legacy')
                SharePoint = @(New-TestSite -Title 'Legacy site' -Url 'https://contoso.sharepoint.com/sites/legacy' -StorageGB 12)
            }

            $result = Update-MigrationGroupWorkloadReconciliation -TenantStatsStore $store
            $row = @($store.GroupWorkloadReconciliation | Select-Object -First 1)

            $row.MailboxEvidenceStatus | Should -Be 'Needs Data'
            $row.MailboxEvidenceSource | Should -Be 'Not collected'
            $row.MailboxSizeGB | Should -BeNullOrEmpty
            $row.MailboxScope | Should -Be 'Needs Data'
            $row.Classification | Should -Match 'mail evidence unavailable'
            $row.Notes | Should -Match 'do not infer zero'
            $result.UnknownMailboxEvidenceCount | Should -Be 1
            $result.MeasuredEmptyMailboxCount | Should -Be 0
        }

        It 'uses joined live Graph or EXO statistics to distinguish measured zero from measured data' {
            $store = @{
                UnifiedGroups = @(
                    New-TestGroup -Name 'Measured Zero' -Smtp 'zero@contoso.com' -SiteUrl $null
                    New-TestGroup -Name 'Measured Data' -Smtp 'data@contoso.com' -SiteUrl $null
                )
                PrimaryMailboxStats = @(
                    [pscustomobject]@{ MailboxGuid = 'zero@contoso.com'; TotalItemSizeBytes = [int64]0; TotalDeletedItemSizeBytes = [int64]0 }
                    [pscustomobject]@{ MailboxGuid = 'data@contoso.com'; TotalItemSizeBytes = [int64](2GB); TotalDeletedItemSizeBytes = [int64](1GB) }
                )
            }

            $result = Update-MigrationGroupWorkloadReconciliation -TenantStatsStore $store
            $zero = @($store.GroupWorkloadReconciliation | Where-Object DisplayName -eq 'Measured Zero' | Select-Object -First 1)
            $data = @($store.GroupWorkloadReconciliation | Where-Object DisplayName -eq 'Measured Data' | Select-Object -First 1)

            $zero.MailboxEvidenceStatus | Should -Be 'Measured empty'
            $zero.MailboxEvidenceSource | Should -Be 'PrimaryMailboxStats (Graph/EXO)'
            $zero.MailboxSizeGB | Should -Be 0
            $zero.MailboxScope | Should -Be 'None'
            $data.MailboxEvidenceStatus | Should -Be 'Measured data'
            $data.MailboxSizeGB | Should -Be 3
            $data.MailboxScope | Should -Be 'BitTitan (optional)'
            $result.MeasuredEmptyMailboxCount | Should -Be 1
            $result.MeasuredMailDataCount | Should -Be 1
            $result.UnknownMailboxEvidenceCount | Should -Be 0
        }

        It 'retains an unmatched dedicated group mailbox as a report-only disposition row' {
            $store = @{ GroupMailboxes = @(New-TestGroupMailbox -Name 'Exchange Only' -Smtp 'exchangeonly@contoso.com' -SizeGB 4) }

            $result = Update-MigrationGroupWorkloadReconciliation -TenantStatsStore $store
            $row = @($store.GroupWorkloadReconciliation | Select-Object -First 1)

            @($result.ReconciliationRows).Count | Should -Be 1
            $row.Classification | Should -Be 'Unmatched group mailbox with measured mail data'
            $row.MailboxEvidenceStatus | Should -Be 'Measured data'
            $row.MailboxScope | Should -Be 'BitTitan (optional)'
            $row.SiteScope | Should -Be 'Needs Data'
            $row.Notes | Should -Match 'did not match a collected UnifiedGroups row'
        }
    }

    Context 'Archive evidence consistency' {
        It 'recognizes status, size, deleted-item, and nonzero GUID archive evidence' {
            Test-AssessmentMailboxHasArchiveEvidence -MailboxRecord ([pscustomobject]@{ ArchiveStatus = 'Enabled' }) | Should -BeTrue
            Test-AssessmentMailboxHasArchiveEvidence -MailboxRecord ([pscustomobject]@{ ArchiveSizeGB = 1 }) | Should -BeTrue
            Test-AssessmentMailboxHasArchiveEvidence -MailboxRecord ([pscustomobject]@{ ArchiveDeletedItemsGB = 0.5 }) | Should -BeTrue
            Test-AssessmentMailboxHasArchiveEvidence -MailboxRecord ([pscustomobject]@{ ArchiveGuid = '11111111-1111-1111-1111-111111111111' }) | Should -BeTrue
        }

        It 'does not treat missing, disabled, or zero GUID archive state as positive evidence' {
            Test-AssessmentMailboxHasArchiveEvidence -MailboxRecord ([pscustomobject]@{ ArchiveStatus = 'None'; ArchiveSizeGB = 0; ArchiveDeletedItemsGB = 0; ArchiveGuid = '00000000-0000-0000-0000-000000000000' }) | Should -BeFalse
        }
    }

    Context 'Dedicated group mailbox inventory fallback' {
        It 'derives the dedicated inventory from unified groups and joins collected mailbox statistics' {
            $store = @{
                UnifiedGroups = @(
                    New-TestGroup -Name 'Measured Zero' -Smtp 'zero@contoso.com' -SiteUrl $null
                    New-TestGroup -Name 'Measured Data' -Smtp 'data@contoso.com' -SiteUrl $null
                )
                GroupMailboxes = @()
                PrimaryMailboxStats = @(
                    [pscustomobject]@{ MailboxGuid = 'zero@contoso.com'; TotalItemSizeBytes = [int64]0; TotalDeletedItemSizeBytes = [int64]0 }
                    [pscustomobject]@{ MailboxGuid = 'data@contoso.com'; TotalItemSizeBytes = [int64](2GB); TotalDeletedItemSizeBytes = [int64](1GB) }
                )
            }

            $result = Ensure-MigrationGroupMailboxInventory -TenantStatsStore $store
            $rows = @(Get-AssessmentExportTableArray -TenantStatsStore $store -Key 'GroupMailboxes')
            $zero = @($rows | Where-Object DisplayName -eq 'Measured Zero' | Select-Object -First 1)
            $data = @($rows | Where-Object DisplayName -eq 'Measured Data' | Select-Object -First 1)

            $result.WasDerived | Should -BeTrue
            $result.RowCount | Should -Be 2
            $result.MeasuredCount | Should -Be 2
            $result.NeedsDataCount | Should -Be 0
            $rows.Count | Should -Be 2
            $zero.RecipientTypeDetails | Should -Be 'GroupMailbox'
            $zero.MailboxEvidenceStatus | Should -Be 'Measured empty'
            $zero.TotalDataToMigrateGB | Should -Be 0
            $data.MailboxEvidenceStatus | Should -Be 'Measured data'
            $data.TotalDataToMigrateGB | Should -Be 3
            $data.InventorySource | Should -Match 'Derived from UnifiedGroups'
            $data.BitTitanLicenseType | Should -Be 'Report only - customer disposition required'
            $data.BitTitanLicenseCount | Should -BeNullOrEmpty
        }

        It 'keeps a genuine Exchange group-mailbox inventory unchanged' {
            $existing = New-TestGroupMailbox -Name 'EXO Group' -Smtp 'exo@contoso.com' -SizeGB 4
            $existing | Add-Member -NotePropertyName InventorySource -NotePropertyValue 'Exchange Online'
            $store = @{
                UnifiedGroups = @(New-TestGroup -Name 'Graph Group' -Smtp 'graph@contoso.com' -SiteUrl $null)
                GroupMailboxes = @($existing)
            }

            $result = Ensure-MigrationGroupMailboxInventory -TenantStatsStore $store
            $rows = @(Get-AssessmentExportTableArray -TenantStatsStore $store -Key 'GroupMailboxes')

            $result.WasDerived | Should -BeFalse
            $rows.Count | Should -Be 1
            $rows[0].DisplayName | Should -Be 'EXO Group'
            $rows[0].InventorySource | Should -Be 'Exchange Online'
        }

        It 'creates a Needs Data inventory row when the unified group has no joinable statistics' {
            $store = @{ UnifiedGroups = @(New-TestGroup -Name 'Unknown Group' -Smtp 'unknown@contoso.com' -SiteUrl $null) }

            $result = Ensure-MigrationGroupMailboxInventory -TenantStatsStore $store
            $row = @(Get-AssessmentExportTableArray -TenantStatsStore $store -Key 'GroupMailboxes' | Select-Object -First 1)

            $result.NeedsDataCount | Should -Be 1
            $row.MailboxEvidenceStatus | Should -Be 'Needs Data'
            $row.MailboxEvidenceSource | Should -Be 'Not collected'
            $row.TotalDataToMigrateGB | Should -BeNullOrEmpty
        }
    }

    Context 'Quote readiness and planning registers' {
        BeforeAll {
            $script:planningStore = @{
                AllMailboxes = @(
                    [pscustomobject]@{
                        DisplayName = 'Inactive Archive User'; UserPrincipalName = 'old@contoso.com'; PrimarySmtpAddress = 'old@contoso.com'
                        RecipientTypeDetails = 'UserMailbox'; IsInactiveMailbox = $true; MailboxSizeGB = 10; ArchiveStatus = 'Active'
                        ArchiveSizeGB = 25; TotalDataToMigrateGB = 35; ExchangeGuid = 'mailbox-1'
                    }
                )
                Users = @(
                    [pscustomobject]@{ DisplayName = 'Inactive Archive User'; UserPrincipalName = 'old@contoso.com'; Mail = 'old@contoso.com'; Id = 'user-1' }
                )
                AllRecipients = @(
                    [pscustomobject]@{ DisplayName = 'Inactive Archive User'; PrimarySmtpAddress = 'old@contoso.com'; RecipientTypeDetails = 'UserMailbox' }
                )
                Domains = @(
                    [pscustomobject]@{ Name = 'contoso.com'; IsDefault = $true; IsInitial = $false; Verified = $true; AuthenticationType = 'Managed' }
                    [pscustomobject]@{ Name = 'contoso.onmicrosoft.com'; IsDefault = $false; IsInitial = $true; Verified = $true; AuthenticationType = 'Managed' }
                )
                GroupWorkloadReconciliation = @(
                    [pscustomobject]@{
                        DisplayName = 'Operations'; PrimarySmtpAddress = 'operations@contoso.com'; MailboxSizeGB = 8; SiteStorageGB = 20
                        MailboxEvidenceStatus = 'Measured data'; MailboxScope = 'BitTitan (optional)'; SiteScope = 'ShareGate'; Classification = 'Group with mail and site'; IsTeam = $false
                    }
                )
                SharePoint = @(
                    [pscustomobject]@{ Title = 'Operations'; Url = 'https://contoso.sharepoint.com/sites/operations'; StorageUsedGB = 20 }
                )
                OneDrive = @(
                    [pscustomobject]@{ Title = 'Old User'; Owner = 'old@contoso.com'; Url = 'https://contoso-my.sharepoint.com/personal/old'; StorageUsedGB = 5 }
                )
                AllTeams = @(
                    [pscustomobject]@{
                        DisplayName = 'Operations'; SharePointSiteUrl = 'https://contoso.sharepoint.com/sites/operations'
                        TotalChannels = 3; SharedChannelCount = 1; ChannelInventoryStatus = 'Measured data'
                        MemberCount = 8; GuestCount = 1; MemberInventoryStatus = 'Measured data'
                    }
                )
                BitTitanLicenseSummary = @(
                    [pscustomobject]@{ Section = 'Totals'; Metric = 'Objects needing license evidence'; Value = 0 }
                )
                BitTitanLicenseDetail = @(
                    [pscustomobject]@{ Identity = 'old@contoso.com'; ObjectType = 'Mailbox'; LicenseUnits = 1 }
                )
                MigrationComplexityFlags = @()
            }

            Update-MigrationPlanningRegisters -TenantStatsStore $script:planningStore
        }

        It 'keeps a source-only assessment at ROM and surfaces MigrationWiz endpoint blockers' {
            $summary = @($script:planningStore.MigrationQuoteReadiness | Where-Object RowType -eq 'Summary' | Select-Object -First 1)
            $summary.ReadinessLevel | Should -Be 'ROM'

            $endpointChecks = @($script:planningStore.MigrationQuoteReadiness | Where-Object DecisionKey -in @('BT-09', 'BT-10'))
            $endpointChecks.Count | Should -Be 2
            @($endpointChecks | Where-Object Status -eq 'Blocker').Count | Should -Be 2
            ($endpointChecks.Notes -join ' ') | Should -Match 'assessment application is not the MigrationWiz application'
            ($endpointChecks.Resolution -join ' ') | Should -Match '24 hours'
        }

        It 'makes Teams effort and SG-03 readiness Needs Data when enrichment failed' {
            $store = $script:planningStore.Clone()
            $store.AllTeams = @(
                [pscustomobject]@{
                    DisplayName = 'Operations'; SharePointSiteUrl = 'https://contoso.sharepoint.com/sites/operations'
                    TotalChannels = 0; SharedChannelCount = 0; ChannelInventoryStatus = 'Needs Data'
                    MemberCount = 0; GuestCount = 0; MemberInventoryStatus = 'Needs Data'
                }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            $effort = @($store.MigrationWorkloadEffort | Where-Object DecisionKey -eq 'EFFORT:TEAMS' | Select-Object -First 1)
            $effort.ProvisionalEffortBand | Should -Be 'Needs Data'
            $effort.Status | Should -Be 'Needs Data'
            $effort.UnknownEvidenceCount | Should -Be 1
            $effort.ComplexityDrivers | Should -Match 'channel inventory Needs Data for 1 Team\(s\)'
            $effort.ComplexityDrivers | Should -Match 'member/guest inventory Needs Data for 1 Team\(s\)'

            $readiness = @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'SG-03' | Select-Object -First 1)
            $readiness.Status | Should -Be 'Needs Data'
            $readiness.Blocking | Should -BeTrue
            $readiness.DiscoveredValue | Should -Match 'Channel inventory Needs Data for 1 Team\(s\)'
            $readiness.Notes | Should -Match 'zero counts are not treated as measured evidence'
        }

        It 'carries inactive archive evidence and the restore-before-migrate prerequisite into object disposition' {
            $row = @($script:planningStore.MigrationObjectDisposition | Where-Object SourceIdentity -eq 'old@contoso.com' | Select-Object -First 1)
            $row.State | Should -Be 'Inactive'
            $row.HasArchive | Should -BeTrue
            $row.ArchiveSizeGB | Should -Be 25
            $row.ProposedDisposition | Should -Match 'Restore/recover, then migrate'
            $row.MigrationPrerequisite | Should -Match 'primary mailbox and its archive'
        }

        It 'creates target, identity, domain, and object-disposition registers' {
            @($script:planningStore.MigrationTargetReadiness).Count | Should -Be 5
            @($script:planningStore.MigrationIdentityMapping).Count | Should -BeGreaterOrEqual 2
            @($script:planningStore.MigrationDomainDependencies).Count | Should -Be 2
            @($script:planningStore.MigrationObjectDisposition).Count | Should -BeGreaterOrEqual 4

            $customDomain = @($script:planningStore.MigrationDomainDependencies | Where-Object Domain -eq 'contoso.com' | Select-Object -First 1)
            $customDomain.Status | Should -Be 'Needs Input'
            $customDomain.ApplicationUriReferenceCount | Should -Be 0
        }

        It 'seeds an editable customer-decision table without marking decisions complete' {
            $decisionRows = @($script:planningStore.MigrationScopeDecisions)
            $decisionRows.Count | Should -BeGreaterThan 25
            @($decisionRows | Where-Object DecisionKey -eq 'BT-09' | Select-Object -ExpandProperty Status) | Should -Be 'Needs Input'
            @($decisionRows | Where-Object DecisionKey -eq 'BT-09' | Select-Object -ExpandProperty DecisionPrompt) | Should -Match 'MigrationWiz'
            @($decisionRows | Where-Object DecisionKey -like 'IDENTITY:*').Count | Should -BeGreaterOrEqual 2
            @($decisionRows | Where-Object DecisionKey -eq 'DOMAIN:contoso.com').Count | Should -Be 1
            @($decisionRows | Where-Object DecisionKey -eq 'MAILBOX:old@contoso.com').Count | Should -Be 1
        }

        It 'keeps canonical decision prompts aligned with the actual decision semantics' {
            $expectedPrompts = [ordered]@{
                'BT-07' = 'If TMB is selected, map each Flex Collaboration entitlement to one Team or one SharePoint library and validate the 100 GB allowance.'
                'DC-03' = 'Confirm DNS TTL, MX/Autodiscover/SPF/DKIM/DMARC, freeze, rollback, and communications ownership.'
                'CP-01' = 'Confirm legal, hold, retention, eDiscovery, and records requirements before restoring inactive mailboxes.'
                'CP-02' = 'Confirm treatment for encrypted, recoverable, oversized, corrupt, and unsupported data.'
                'CP-03' = 'Confirm destination retention, sensitivity, DLP, sharing, and residency requirements.'
            }

            foreach ($entry in $expectedPrompts.GetEnumerator()) {
                @($script:planningStore.MigrationScopeDecisions |
                    Where-Object DecisionKey -eq $entry.Key |
                    Select-Object -ExpandProperty DecisionPrompt) | Should -Be $entry.Value
            }

            # The effort register intentionally reuses BT-07. It must not replace the
            # library/FCL mapping decision with a generic effort-band acknowledgement.
            @($script:planningStore.MigrationScopeDecisions |
                Where-Object DecisionKey -eq 'BT-07' |
                Select-Object -ExpandProperty DecisionPrompt) | Should -Not -Match 'Review the provisional effort band'
        }

        It 'refreshes stale generated prompts while preserving customer-entered values' {
            $store = $script:planningStore.Clone()
            $store.MigrationScopeDecisions = @(
                [pscustomobject]@{
                    DecisionKey = 'BT-07'; DecisionPrompt = 'Review the provisional effort band.'
                    CustomerConfirmedValue = 'Library map attached'; Status = 'Approved'; InScope = $true
                    TargetMapping = 'Library map v1'; MigrationTool = 'Tenant Migration Bundle'; DecisionOwner = 'SE'; Notes = 'Validated under 100 GB'
                }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            $decision = @($store.MigrationScopeDecisions | Where-Object DecisionKey -eq 'BT-07' | Select-Object -First 1)
            $decision.DecisionPrompt | Should -Be 'If TMB is selected, map each Flex Collaboration entitlement to one Team or one SharePoint library and validate the 100 GB allowance.'
            $decision.CustomerConfirmedValue | Should -Be 'Library map attached'
            $decision.InScope | Should -BeTrue
            $decision.TargetMapping | Should -Be 'Library map v1'
            $decision.DecisionOwner | Should -Be 'SE'
        }

        It 'does not substitute site count for optional Tenant Migration Bundle FCL demand' {
            $row = @($script:planningStore.MigrationWorkloadEffort | Where-Object DecisionKey -eq 'BT-07' | Select-Object -First 1)
            $row.ObjectCount | Should -BeNullOrEmpty
            $row.ProvisionalEffortBand | Should -Be 'Needs Library Mapping'
            ($row.ComplexityDrivers + ' ' + $row.EstimateBasis) | Should -Match 'site count is not a valid substitute'
        }

        It 'counts measured group conversations under the optional BitTitan workload' {
            $row = @($script:planningStore.MigrationWorkloadEffort | Where-Object DecisionKey -eq 'EFFORT:GROUPMAIL' | Select-Object -First 1)
            $row.ObjectCount | Should -Be 1
            $row.DataGB | Should -Be 8
            $row.UnknownEvidenceCount | Should -Be 0
            $row.ProvisionalEffortBand | Should -Be 'Medium'
            @($script:planningStore.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'BT-04').Count | Should -Be 1
        }

        It 'keeps unknown group mailbox evidence at Needs Data instead of treating it as empty' {
            $store = $script:planningStore.Clone()
            $store.GroupWorkloadReconciliation = @(
                [pscustomobject]@{
                    DisplayName = 'Unknown Conversations'; PrimarySmtpAddress = 'unknown@contoso.com'; MailboxSizeGB = $null; SiteStorageGB = 20
                    MailboxEvidenceStatus = 'Needs Data'; MailboxScope = 'Needs Data'; SiteScope = 'ShareGate'; Classification = 'Group, site only - mailbox evidence unavailable'; IsTeam = $false
                }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            $row = @($store.MigrationWorkloadEffort | Where-Object DecisionKey -eq 'EFFORT:GROUPMAIL' | Select-Object -First 1)
            $row.ObjectCount | Should -Be 0
            $row.UnknownEvidenceCount | Should -Be 1
            $row.ProvisionalEffortBand | Should -Be 'Needs Data'
            $row.ComplexityDrivers | Should -Match '1 group mailbox\(es\) need statistics'
            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'AUTO-GROUP-MAIL-EVIDENCE' | Select-Object -ExpandProperty Status) | Should -Be 'Blocker'
        }

        It 'prevents Firm readiness from bypassing compliance decisions' {
            @($script:planningStore.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'CP-01' | Select-Object -ExpandProperty Status) | Should -Be 'Blocker'
            @($script:planningStore.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'CP-02' | Select-Object -ExpandProperty Status) | Should -Be 'Condition'
            @($script:planningStore.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'CP-03' | Select-Object -ExpandProperty Status) | Should -Be 'Condition'
            @($script:planningStore.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'CP-02' | Select-Object -ExpandProperty Requirement) | Should -Match 'encrypted, recoverable, oversized, corrupt, and unsupported data'
            @($script:planningStore.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'CP-03' | Select-Object -ExpandProperty Requirement) | Should -Match 'destination retention, sensitivity, DLP, sharing, and residency'
        }

        It 'requires BT-07 library mapping when approved BT-06 explicitly enables TMB' {
            $store = $script:planningStore.Clone()
            $store.MigrationScopeDecisions = @(
                [pscustomobject]@{ DecisionKey = 'BT-06'; Status = 'Approved'; InScope = $true; MigrationTool = 'Tenant Migration Bundle' }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'BT-06' | Select-Object -ExpandProperty Status) | Should -Be 'Complete'
            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'BT-07' | Select-Object -ExpandProperty Status) | Should -Be 'Blocker'
            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'AUTO-TMB-CHOICE').Count | Should -Be 0
        }

        It 'does not treat an approved but blank TMB decision as an explicit choice' {
            $store = $script:planningStore.Clone()
            $store.MigrationScopeDecisions = @(
                [pscustomobject]@{ DecisionKey = 'BT-06'; Status = 'Approved'; CustomerConfirmedValue = 'Approved' }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'AUTO-TMB-CHOICE' | Select-Object -ExpandProperty Status) | Should -Be 'Blocker'
            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'BT-07').Count | Should -Be 0
        }

        It 'creates transparent provisional wave estimates' {
            $rows = @($script:planningStore.MigrationWavePlan)
            $rows.Count | Should -Be 5
            @($rows | Where-Object Status -eq 'Provisional').Count | Should -Be 5
            @($rows | Where-Object Wave -eq '1 - Pilot' | Select-Object -ExpandProperty PlanningAssumption) | Should -Match '5% capped at 25'
            @($rows | Where-Object Wave -eq '2-N - Mailbox production' | Select-Object -ExpandProperty PlanningAssumption) | Should -Match 'up to 100 mailbox objects per wave'
        }

        It 'merges completed questionnaire decisions by DecisionKey' {
            $store = $script:planningStore.Clone()
            $store.MigrationScopeDecisions = @(
                [pscustomobject]@{ DecisionKey = 'BT-09'; CustomerConfirmedValue = 'Validated'; Status = 'Approved'; DecisionOwner = 'Messaging Lead' }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            $targetRow = @($store.MigrationTargetReadiness | Where-Object DecisionKey -eq 'BT-09' | Select-Object -First 1)
            $quoteRow = @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'BT-09' | Select-Object -First 1)
            $targetRow.Status | Should -Be 'Confirmed'
            $quoteRow.Status | Should -Be 'Complete'
            $quoteRow.DecisionOwner | Should -Be 'Messaging Lead'
        }

        It 'does not let broad section approvals confirm child identity, disposition, or domain rows' {
            $store = $script:planningStore.Clone()
            $store.MigrationScopeDecisions = @(
                [pscustomobject]@{ DecisionKey = 'ID-02'; Status = 'Approved'; CustomerConfirmedValue = 'Mappings approved'; TargetMapping = 'Broad mapping text' }
                [pscustomobject]@{ DecisionKey = 'ID-04'; Status = 'Approved'; CustomerConfirmedValue = 'Dispositions approved'; InScope = $true; TargetMapping = 'Broad target'; MigrationTool = 'Broad tool' }
                [pscustomobject]@{ DecisionKey = 'DC-01'; Status = 'Approved'; CustomerConfirmedValue = 'Move domains'; InScope = $true; TargetMapping = 'First'; DecisionOwner = 'DNS Owner' }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            @($store.MigrationIdentityMapping | Where-Object MappingStatus -eq 'Confirmed').Count | Should -Be 0
            @($store.MigrationObjectDisposition | Where-Object Status -eq 'Confirmed').Count | Should -Be 0
            @($store.MigrationDomainDependencies | Where-Object Domain -eq 'contoso.com' | Select-Object -ExpandProperty Status) | Should -Be 'Needs Input'
            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'AUTO-IDENTITY-MAPPINGS' | Select-Object -ExpandProperty Status) | Should -Be 'Blocker'
            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'AUTO-OBJECT-DISPOSITIONS' | Select-Object -ExpandProperty Status) | Should -Be 'Blocker'
            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'AUTO-DOMAIN-DISPOSITIONS' | Select-Object -ExpandProperty Status) | Should -Be 'Blocker'
            @($store.MigrationQuoteReadiness | Where-Object RowType -eq 'Summary' | Select-Object -ExpandProperty ReadinessLevel) | Should -Be 'ROM'
        }

        It 'requires populated object-specific fields in addition to an approved status' {
            $store = $script:planningStore.Clone()
            $store.MigrationScopeDecisions = @(
                [pscustomobject]@{ DecisionKey = 'IDENTITY:old@contoso.com'; Status = 'Approved' }
                [pscustomobject]@{ DecisionKey = 'MAILBOX:old@contoso.com'; Status = 'Approved'; InScope = $true }
                [pscustomobject]@{ DecisionKey = 'DOMAIN:contoso.com'; Status = 'Approved'; InScope = $true; CustomerConfirmedValue = 'Move'; DecisionOwner = 'DNS Owner' }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            @($store.MigrationIdentityMapping | Where-Object DecisionKey -eq 'IDENTITY:old@contoso.com' | Select-Object -ExpandProperty MappingStatus) | Should -Be 'Incomplete - target mapping required'
            @($store.MigrationObjectDisposition | Where-Object DecisionKey -eq 'MAILBOX:old@contoso.com' | Select-Object -ExpandProperty Status) | Should -Be 'Incomplete - target mapping required'
            @($store.MigrationDomainDependencies | Where-Object DecisionKey -eq 'DOMAIN:contoso.com' | Select-Object -ExpandProperty Status) | Should -Be 'Incomplete Decision'
        }

        It 'subtracts only confirmed object-specific exclusions from effort data and provisional waves' {
            $store = $script:planningStore.Clone()
            $store.MigrationScopeDecisions = @(
                [pscustomobject]@{ DecisionKey = 'MAILBOX:old@contoso.com'; Status = 'Approved'; InScope = $false }
                [pscustomobject]@{ DecisionKey = 'SITE:https://contoso.sharepoint.com/sites/operations'; Status = 'Approved'; InScope = $false }
                [pscustomobject]@{ DecisionKey = 'ONEDRIVE:https://contoso-my.sharepoint.com/personal/old'; Status = 'Approved'; InScope = $false }
                [pscustomobject]@{ DecisionKey = 'TEAM:https://contoso.sharepoint.com/sites/operations'; Status = 'Approved'; InScope = $false }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            $mailEffort = @($store.MigrationWorkloadEffort | Where-Object DecisionKey -eq 'EFFORT:MAILBOXES' | Select-Object -First 1)
            $mailEffort.DiscoveredObjectCount | Should -Be 1
            $mailEffort.ObjectCount | Should -Be 0
            $mailEffort.DataGB | Should -Be 0
            $mailEffort.ExplicitExcludedCount | Should -Be 1
            $mailEffort.ExplicitExcludedDataGB | Should -Be 35
            $mailEffort.UndecidedObjectCount | Should -Be 0

            foreach ($key in @('EFFORT:SHAREPOINT', 'EFFORT:ONEDRIVE', 'EFFORT:TEAMS')) {
                @($store.MigrationWorkloadEffort | Where-Object DecisionKey -eq $key | Select-Object -ExpandProperty ObjectCount) | Should -Be 0
            }

            @($store.MigrationWavePlan | Where-Object Wave -eq '1 - Pilot' | Select-Object -ExpandProperty CandidateObjectCount) | Should -Be 0
            @($store.MigrationWavePlan | Where-Object Wave -eq '2-N - Mailbox production' | Select-Object -ExpandProperty CandidateObjectCount) | Should -Be 0
            @($store.MigrationWavePlan | Where-Object Wave -eq 'Parallel - Collaboration' | Select-Object -ExpandProperty CandidateObjectCount) | Should -Be 0

            $mailDisposition = @($store.MigrationObjectDisposition | Where-Object DecisionKey -eq 'MAILBOX:old@contoso.com' | Select-Object -First 1)
            $mailDisposition.Status | Should -Be 'Confirmed'
            $mailDisposition.ScopeAccountingState | Should -Be 'Confirmed Excluded'
        }

        It 'keeps an unconfirmed false and every undecided object in conservative planning counts' {
            $store = $script:planningStore.Clone()
            $store.AllMailboxes = @($store.AllMailboxes) + @(
                [pscustomobject]@{
                    DisplayName = 'Second User'; UserPrincipalName = 'second@contoso.com'; PrimarySmtpAddress = 'second@contoso.com'
                    RecipientTypeDetails = 'UserMailbox'; IsInactiveMailbox = $false; MailboxSizeGB = 10; ArchiveStatus = 'None'
                    ArchiveSizeGB = 0; TotalDataToMigrateGB = 10; ExchangeGuid = 'mailbox-2'
                }
            )
            $store.MigrationScopeDecisions = @(
                [pscustomobject]@{ DecisionKey = 'MAILBOX:old@contoso.com'; Status = 'Approved'; InScope = $true; TargetMapping = 'old@target.com'; MigrationTool = 'BitTitan MigrationWiz' }
                [pscustomobject]@{ DecisionKey = 'MAILBOX:second@contoso.com'; Status = 'Needs Input'; InScope = $false }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            $mailEffort = @($store.MigrationWorkloadEffort | Where-Object DecisionKey -eq 'EFFORT:MAILBOXES' | Select-Object -First 1)
            $mailEffort.DiscoveredObjectCount | Should -Be 2
            $mailEffort.ObjectCount | Should -Be 2
            $mailEffort.DataGB | Should -Be 45
            $mailEffort.ExplicitIncludedCount | Should -Be 1
            $mailEffort.ExplicitExcludedCount | Should -Be 0
            $mailEffort.UndecidedObjectCount | Should -Be 1

            $secondDisposition = @($store.MigrationObjectDisposition | Where-Object DecisionKey -eq 'MAILBOX:second@contoso.com' | Select-Object -First 1)
            $secondDisposition.ScopeAccountingState | Should -Be 'Undecided - included in provisional counts'
        }

        It 'treats a properly confirmed identity exclusion as resolved without requiring a target mapping' {
            $store = $script:planningStore.Clone()
            $store.MigrationScopeDecisions = @(
                [pscustomobject]@{ DecisionKey = 'IDENTITY:old@contoso.com'; Status = 'Approved'; InScope = $false }
            )

            Update-MigrationPlanningRegisters -TenantStatsStore $store

            @($store.MigrationIdentityMapping | Where-Object DecisionKey -eq 'IDENTITY:old@contoso.com' | Select-Object -ExpandProperty MappingStatus) | Should -Be 'Confirmed Excluded'
            @($store.MigrationQuoteReadiness | Where-Object DecisionKey -eq 'AUTO-IDENTITY-MAPPINGS' | Select-Object -ExpandProperty DiscoveredValue) | Should -Be 1
        }
    }
}
