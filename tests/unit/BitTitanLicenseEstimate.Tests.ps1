Describe 'Get-ArrayaBitTitanLicenseEstimate' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        Import-Module -Name $script:manifestPath -Force -ErrorAction Stop

        function script:New-TestUser {
            param([string]$Upn)
            [pscustomobject]@{ UserPrincipalName = $Upn; Mail = $Upn }
        }

        function script:New-TestMailbox {
            param(
                [string]$Name,
                [string]$Smtp,
                [string]$Type = 'UserMailbox',
                $SizeGB = $null,
                $DeletedGB = $null,
                $ArchiveGB = $null,
                $ArchiveDeletedGB = $null,
                [string]$ArchiveStatus = 'None',
                [bool]$IsInactive = $false
            )
            [pscustomobject]@{
                DisplayName          = $Name
                PrimarySmtpAddress   = $Smtp
                RecipientTypeDetails = $Type
                MailboxSizeGB        = $SizeGB
                DeletedItemsGB       = $DeletedGB
                ArchiveSizeGB        = $ArchiveGB
                ArchiveDeletedItemsGB = $ArchiveDeletedGB
                ArchiveStatus        = $ArchiveStatus
                IsInactiveMailbox    = $IsInactive
            }
        }

        function script:Get-Line {
            param($Estimate, [string]$DisplayName)
            @($Estimate.ObjectLines | Where-Object { $_.DisplayName -eq $DisplayName } | Select-Object -First 1) | Select-Object -First 1
        }
    }

    Context 'Block arithmetic' {
        BeforeAll {
            $script:users = @(
                New-TestUser -Upn 'small@contoso.com'
                New-TestUser -Upn 'exactly50@contoso.com'
                New-TestUser -Upn 'justover@contoso.com'
                New-TestUser -Upn 'archived@contoso.com'
            )
            $script:mailboxes = @(
                New-TestMailbox -Name 'Small' -Smtp 'small@contoso.com' -SizeGB 10
                New-TestMailbox -Name 'Exactly 50' -Smtp 'exactly50@contoso.com' -SizeGB 50
                New-TestMailbox -Name 'Just over' -Smtp 'justover@contoso.com' -SizeGB 50.5
                New-TestMailbox -Name 'Archived' -Smtp 'archived@contoso.com' -SizeGB 100 -ArchiveGB 300 -ArchiveStatus 'Active'
            )
            $script:estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $script:mailboxes -UserRows $script:users
        }

        It 'charges one license block for a mailbox well under the threshold' {
            (Get-Line -Estimate $script:estimate -DisplayName 'Small').LicenseUnits | Should -Be 1
        }

        It 'treats exactly one block as one license, not two' {
            # The previous model used "-le 50 => 1, else 2", so the boundary itself was correct
            # but everything past 100 GB silently under-counted.
            (Get-Line -Estimate $script:estimate -DisplayName 'Exactly 50').LicenseUnits | Should -Be 1
        }

        It 'rolls to a second block as soon as the threshold is crossed' {
            (Get-Line -Estimate $script:estimate -DisplayName 'Just over').LicenseUnits | Should -Be 2
        }

        It 'sizes archive content and maps the object to one User Migration Bundle' {
            $line = Get-Line -Estimate $script:estimate -DisplayName 'Archived'
            $line.SizeGB | Should -Be 400
            $line.LicenseUnits | Should -Be 1
            $line.RecommendedSku | Should -Be 'User Migration Bundle'
            $line.HasActiveArchive | Should -BeTrue
            $line.UmbEligibilityValidationRequired | Should -BeFalse
        }
    }

    Context 'Size parsing' {
        It 'parses Exchange-style size strings consistently with numeric values' {
            $users = @(New-TestUser -Upn 'a@contoso.com', 'b@contoso.com')
            $mailboxes = @(
                New-TestMailbox -Name 'String size' -Smtp 'a@contoso.com' -SizeGB '120 GB'
                New-TestMailbox -Name 'Numeric size' -Smtp 'b@contoso.com' -SizeGB 120
            )
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $mailboxes -UserRows $users

            (Get-Line -Estimate $estimate -DisplayName 'String size').SizeGB | Should -Be 120
            (Get-Line -Estimate $estimate -DisplayName 'String size').LicenseUnits |
                Should -Be (Get-Line -Estimate $estimate -DisplayName 'Numeric size').LicenseUnits
        }

        It 'converts non-GB units rather than reading the number as GB' {
            $mailboxes = @(New-TestMailbox -Name 'In MB' -Smtp 'mb@contoso.com' -SizeGB '512 MB')
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $mailboxes

            (Get-Line -Estimate $estimate -DisplayName 'In MB').SizeGB | Should -Be 0.5
        }
    }

    Context 'Objects that cannot be priced' {
        It 'buckets a mailbox with no size data instead of dropping it from the totals' {
            # The previous model returned a null license type here, and the mailbox then
            # vanished from every count with no trace.
            $mailboxes = @(New-TestMailbox -Name 'No stats' -Smtp 'nostats@contoso.com' -SizeGB $null)
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $mailboxes

            $line = Get-Line -Estimate $estimate -DisplayName 'No stats'
            $line | Should -Not -BeNullOrEmpty
            $line.RecommendedSku | Should -Be 'Needs data'
            $line.LicenseUnits | Should -BeNullOrEmpty
            $estimate.NeedsDataCount | Should -Be 1
        }

        It 'makes unknown archive state a needs-data quote condition' {
            $mailbox = New-TestMailbox -Name 'Archive unknown' -Smtp 'unknown@contoso.com' -SizeGB 12
            $mailbox.ArchiveStatus = $null
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows @($mailbox)
            $line = Get-Line -Estimate $estimate -DisplayName 'Archive unknown'

            $line.RecommendedSku | Should -Be 'Needs data'
            $line.EstimateStatus | Should -Be 'Needs archive status'
            $line.LicenseUnits | Should -BeNullOrEmpty
            $line.Notes | Should -Match 'archive content requires UMB'
        }

        It 'keeps Microsoft 365 Group mailboxes visible but excludes all group license units by default' {
            $groups = @(
                New-TestMailbox -Name 'Empty group' -Smtp 'g1@contoso.com' -Type 'GroupMailbox' -SizeGB 0
                New-TestMailbox -Name 'Active group' -Smtp 'g2@contoso.com' -Type 'GroupMailbox' -SizeGB 12
            )
            $estimate = Get-ArrayaBitTitanLicenseEstimate -GroupMailboxRows $groups

            (Get-Line -Estimate $estimate -DisplayName 'Empty group').RecommendedSku | Should -Be 'No migration'
            (Get-Line -Estimate $estimate -DisplayName 'Empty group').LicenseUnits | Should -BeNullOrEmpty
            (Get-Line -Estimate $estimate -DisplayName 'Active group').RecommendedSku | Should -Be 'Report only'
            (Get-Line -Estimate $estimate -DisplayName 'Active group').LicenseUnits | Should -BeNullOrEmpty
            (Get-Line -Estimate $estimate -DisplayName 'Active group').SizeGB | Should -Be 12
            $estimate.EmptyShellCount | Should -Be 1
            $estimate.GroupMailboxDataCount | Should -Be 1
            $estimate.ExcludedGroupMailboxCount | Should -Be 1
            $estimate.GroupMailboxLicensingMode | Should -Be 'Report only (practice default)'
        }

        It 'only counts a group mailbox when the caller explicitly overrides the practice default' {
            $groups = @(New-TestMailbox -Name 'Approved group exception' -Smtp 'approved@contoso.com' -Type 'GroupMailbox' -SizeGB 12)
            $estimate = Get-ArrayaBitTitanLicenseEstimate -GroupMailboxRows $groups -IncludeGroupMailboxLicenseUnits

            (Get-Line -Estimate $estimate -DisplayName 'Approved group exception').LicenseUnits | Should -Be 1
            (Get-Line -Estimate $estimate -DisplayName 'Approved group exception').EstimateStatus | Should -Be 'Included in unit estimate'
            $estimate.ExcludedGroupMailboxCount | Should -Be 0
            $estimate.GroupMailboxLicensingMode | Should -Be 'Included by explicit override'
        }

        It 'marks inactive mailboxes as restore-then-migrate and exposes archive evidence per mailbox' {
            $mailboxes = @(
                New-TestMailbox -Name 'Inactive with archive' -Smtp 'former@contoso.com' -SizeGB 8 -ArchiveGB 22 -ArchiveDeletedGB 3 -ArchiveStatus 'Active' -IsInactive $true
                New-TestMailbox -Name 'Inactive without archive' -Smtp 'former2@contoso.com' -SizeGB 4 -IsInactive $true
            )
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $mailboxes
            $withArchive = Get-Line -Estimate $estimate -DisplayName 'Inactive with archive'
            $withoutArchive = Get-Line -Estimate $estimate -DisplayName 'Inactive without archive'

            $withArchive.IsInactiveMailbox | Should -BeTrue
            $withArchive.HasArchive | Should -BeTrue
            $withArchive.ArchiveStatus | Should -Be 'Active'
            $withArchive.ArchiveSizeGB | Should -Be 22
            $withArchive.ArchiveDeletedItemsGB | Should -Be 3
            $withArchive.RecommendedSku | Should -Be 'User Migration Bundle'
            $withArchive.LicenseUnits | Should -Be 1
            $withArchive.UmbEligibilityValidationRequired | Should -BeTrue
            $withArchive.EstimateStatus | Should -Match 'validate eligibility'
            $withArchive.MigrationPrerequisite | Should -Match 'Restore or activate'
            $withoutArchive.HasArchive | Should -BeFalse
            $estimate.InactiveMailboxCount | Should -Be 2
            $estimate.InactiveMailboxWithArchiveCount | Should -Be 1
            $estimate.InactiveArchiveDataGB | Should -Be 25
        }

        It 'recognizes enabled, deleted-item-only, and nonzero-GUID archive evidence' {
            $enabled = New-TestMailbox -Name 'Enabled archive' -Smtp 'enabled@contoso.com' -SizeGB 4 -ArchiveStatus 'Enabled'
            $deletedOnly = New-TestMailbox -Name 'Deleted archive evidence' -Smtp 'deleted@contoso.com' -SizeGB 4 -ArchiveDeletedGB 2
            $guidOnly = New-TestMailbox -Name 'GUID archive evidence' -Smtp 'guid@contoso.com' -SizeGB 4
            $guidOnly | Add-Member -NotePropertyName ArchiveGuid -NotePropertyValue '11111111-1111-1111-1111-111111111111'

            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows @($enabled, $deletedOnly, $guidOnly)

            foreach ($name in @('Enabled archive', 'Deleted archive evidence', 'GUID archive evidence')) {
                (Get-Line -Estimate $estimate -DisplayName $name).HasArchive | Should -BeTrue
                (Get-Line -Estimate $estimate -DisplayName $name).RecommendedSku | Should -Be 'User Migration Bundle'
            }
        }

        It 'counts an archive-bearing non-user object as a UMB planning unit but requires eligibility validation' {
            $mailboxes = @(New-TestMailbox -Name 'Shared archive' -Smtp 'sharedarchive@contoso.com' -Type 'SharedMailbox' -SizeGB 20 -ArchiveGB 30 -ArchiveStatus 'Active')
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $mailboxes
            $line = Get-Line -Estimate $estimate -DisplayName 'Shared archive'

            $line.RecommendedSku | Should -Be 'User Migration Bundle'
            $line.LicenseUnits | Should -Be 1
            $line.BundleEligible | Should -BeFalse
            $line.UmbEligibilityValidationRequired | Should -BeTrue
            $line.EstimateStatus | Should -Be 'Included as UMB planning unit - validate eligibility'
            $line.Notes | Should -Match 'Validate UMB eligibility'
        }
    }

    Context 'Bundle eligibility' {
        BeforeAll {
            $script:users = @(New-TestUser -Upn 'user@contoso.com')
            $script:mixed = @(
                New-TestMailbox -Name 'Real user' -Smtp 'user@contoso.com' -Type 'UserMailbox' -SizeGB 10
                New-TestMailbox -Name 'Shared' -Smtp 'shared@contoso.com' -Type 'SharedMailbox' -SizeGB 10
                New-TestMailbox -Name 'Room' -Smtp 'room@contoso.com' -Type 'RoomMailbox' -SizeGB 1
            )
            $script:groups = @(New-TestMailbox -Name 'Group' -Smtp 'group@contoso.com' -Type 'GroupMailbox' -SizeGB 5)
            $script:estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $script:mixed -GroupMailboxRows $script:groups -UserRows $script:users
        }

        It 'treats only mailboxes with a backing licensable user as bundle eligible' {
            (Get-Line -Estimate $script:estimate -DisplayName 'Real user').BundleEligible | Should -BeTrue
            (Get-Line -Estimate $script:estimate -DisplayName 'Shared').BundleEligible | Should -BeFalse
            (Get-Line -Estimate $script:estimate -DisplayName 'Room').BundleEligible | Should -BeFalse
            (Get-Line -Estimate $script:estimate -DisplayName 'Group').BundleEligible | Should -BeFalse
        }

        It 'counts non-bundle-eligible mailboxes a-la-carte and leaves groups report-only on the bundle path' {
            $bundlePath = @($script:estimate.PathComparison | Where-Object { $_.Path -eq 'UMB-led' } | Select-Object -First 1)
            $bundlePath.PrimaryUnits | Should -Be 1
            $bundlePath.SecondaryUnits | Should -Be 2
            $bundlePath.TotalUnits | Should -Be 3
        }
    }

    Context 'Path totals' {
        It 'totals the a-la-carte path as the sum of every billable line' {
            $users = @(New-TestUser -Upn 'a@contoso.com')
            $mailboxes = @(
                New-TestMailbox -Name 'A' -Smtp 'a@contoso.com' -SizeGB 10     # 1 unit
                New-TestMailbox -Name 'B' -Smtp 'b@contoso.com' -SizeGB 120    # 3 units
                New-TestMailbox -Name 'C' -Smtp 'c@contoso.com' -SizeGB $null  # excluded
            )
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $mailboxes -UserRows $users

            $workloadFit = @($estimate.PathComparison | Where-Object { $_.Path -eq 'Workload fit' } | Select-Object -First 1)
            $workloadFit.PrimaryUnits | Should -Be 4
            $workloadFit.SecondaryUnits | Should -Be 0
            $workloadFit.TotalUnits | Should -Be 4
        }

        It 'does not make a cost-based recommendation without commercial rates' {
            $mailboxes = @(New-TestMailbox -Name 'A' -Smtp 'a@contoso.com' -SizeGB 10)
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $mailboxes

            @($estimate.PathComparison | Where-Object { $_.IsRecommended }).Count | Should -Be 0
            $estimate.RecommendedPath | Should -Be 'SE selection required'
            @($estimate.PathComparison | Where-Object { $_.SelectionStatus -eq 'SE commercial comparison required' }).Count | Should -Be 2
            $estimate.PSObject.Properties.Name | Should -Not -Contain 'Cost'
            $estimate.PSObject.Properties.Name | Should -Not -Contain 'Price'
        }

        It 'surfaces the unused bundle entitlement caveat on the bundle path' {
            # BitTitan is used for mailbox data only here; ShareGate handles files. An engineer
            # reading a bundle recommendation must not assume OneDrive is covered.
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows @(New-TestMailbox -Name 'A' -Smtp 'a@contoso.com' -SizeGB 10)
            $bundlePath = @($estimate.PathComparison | Where-Object { $_.Path -eq 'UMB-led' } | Select-Object -First 1)

            $bundlePath.Notes | Should -Match 'ShareGate'
            $bundlePath.Notes | Should -Match 'unused'
        }

        It 'surfaces Tenant Migration Bundle as optional guidance without fabricating units' {
            $estimate = Get-ArrayaBitTitanLicenseEstimate `
                -MailboxRows @(New-TestMailbox -Name 'A' -Smtp 'a@contoso.com' -SizeGB 10) `
                -UserRows @(New-TestUser -Upn 'a@contoso.com')
            $tenantBundle = @($estimate.PlanningOptions | Where-Object { $_.Option -eq 'Tenant Migration Bundle' } | Select-Object -First 1)

            $tenantBundle | Should -Not -BeNullOrEmpty
            $tenantBundle.DefaultPractice | Should -Be 'Not used by default'
            $tenantBundle.PlanningUnits | Should -Be 1
            $tenantBundle.WorkloadCoverage | Should -Match 'User Migration Bundle'
            $tenantBundle.WorkloadCoverage | Should -Match 'Flex Collaboration'
            $tenantBundle.UnitBasis | Should -Match '100 GB'
            $tenantBundle.Decision | Should -Match 'Needs Data'
            $tenantBundle.WhyNormallyExcluded | Should -Match 'ShareGate'
            $tenantBundle.Exclusions | Should -Match 'Public folders'
            $tenantBundle.Exclusions | Should -Not -Match 'Teams private chat'
            $tenantBundle.PrivateChatValidation | Should -Match 'separate Collaboration \(Private Chats\) project'
            $tenantBundle.PrivateChatValidation | Should -Match 'do not assume.*inclusion or exclusion'
            $tenantBundle.PrivateChatGuideUrl | Should -Be 'https://help.bittitan.com/hc/en-us/articles/25603590557979-Teams-Private-Chat-Migration-Guide'
            $tenantBundle.GettingStartedUrl | Should -Be 'https://help.bittitan.com/hc/en-us/articles/1260806785629-Getting-Started-with-Migrations'
            $tenantBundle.Decision | Should -Match 'PCH capability and licensing must be revalidated'
            $tenantBundle.GuidanceParagraph | Should -Match 'optional and is not Arraya''s default'
            $tenantBundle.GuidanceParagraph | Should -Match 'Do not add its .* candidate unit\(s\) to the practical mix above'
            $tenantBundle.GuidanceParagraph | Should -Match 'alternate packaging option'
            $tenantBundle.GuidanceParagraph | Should -Match 'site count alone is not enough'
            $tenantBundle.GuidanceParagraph | Should -Match 'official guidance is changing and currently inconsistent'
            $tenantBundle.GuidanceParagraph | Should -Match 'no prices or automatic cost-based recommendation'
        }
    }

    Context 'Canonical practical license mix' {
        It 'groups Mailbox Migration and UMB units by the requested recipient categories without counting group mailboxes' {
            $mailboxes = @(
                New-TestMailbox -Name 'Room' -Smtp 'room@contoso.com' -Type 'RoomMailbox' -SizeGB 60
                New-TestMailbox -Name 'Shared' -Smtp 'shared@contoso.com' -Type 'SharedMailbox' -SizeGB 10
                New-TestMailbox -Name 'Shared archive' -Smtp 'sharedarchive@contoso.com' -Type 'SharedMailbox' -SizeGB 10 -ArchiveGB 5 -ArchiveStatus 'Active'
                New-TestMailbox -Name 'Shared missing stats' -Smtp 'sharedmissing@contoso.com' -Type 'SharedMailbox' -SizeGB $null
                New-TestMailbox -Name 'Inactive' -Smtp 'inactive@contoso.com' -Type 'SharedMailbox' -SizeGB 5 -IsInactive $true
                New-TestMailbox -Name 'Inactive archive' -Smtp 'inactivearchive@contoso.com' -Type 'SharedMailbox' -SizeGB 5 -ArchiveGB 2 -ArchiveStatus 'Active' -IsInactive $true
                New-TestMailbox -Name 'User' -Smtp 'user@contoso.com' -Type 'UserMailbox' -SizeGB 10
                New-TestMailbox -Name 'User archive' -Smtp 'userarchive@contoso.com' -Type 'UserMailbox' -SizeGB 10 -ArchiveGB 20 -ArchiveStatus 'Active'
            )
            $groups = @(
                New-TestMailbox -Name 'Group data' -Smtp 'groupdata@contoso.com' -Type 'GroupMailbox' -SizeGB 12
                New-TestMailbox -Name 'Group empty' -Smtp 'groupempty@contoso.com' -Type 'GroupMailbox' -SizeGB 0
                New-TestMailbox -Name 'Group unknown' -Smtp 'groupunknown@contoso.com' -Type 'GroupMailbox' -SizeGB $null
            )
            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows $mailboxes -GroupMailboxRows $groups -UserRows @(
                New-TestUser -Upn 'user@contoso.com'
                New-TestUser -Upn 'userarchive@contoso.com'
            )
            $mix = @($estimate.LicenseMixBreakdown)
            $byCategory = @{}
            foreach ($row in $mix) { $byCategory[[string]$row.Category] = $row }

            $mix.Count | Should -Be 7
            @($mix.Category) | Should -Be @(
                'Office 365 Groups (report-only / optional)', 'Resources', 'Shared Mailboxes', 'Inactive',
                'Inactive + archive', 'User Mailboxes', 'User Mailboxes + archive'
            )
            $byCategory['Office 365 Groups (report-only / optional)'].ObjectCount | Should -Be 3
            $byCategory['Office 365 Groups (report-only / optional)'].MailboxMigrationLicenseUnits | Should -Be 0
            $byCategory['Office 365 Groups (report-only / optional)'].UserMigrationBundleUnits | Should -Be 0
            $byCategory['Office 365 Groups (report-only / optional)'].ReportOnlyObjectCount | Should -Be 3
            $byCategory['Office 365 Groups (report-only / optional)'].OptionalMigrationObjectCount | Should -Be 1
            $byCategory['Office 365 Groups (report-only / optional)'].NeedsDataObjectCount | Should -Be 1
            $byCategory['Resources'].MailboxMigrationLicenseUnits | Should -Be 2
            $byCategory['Shared Mailboxes'].ObjectCount | Should -Be 3
            $byCategory['Shared Mailboxes'].MailboxMigrationLicenseUnits | Should -Be 1
            $byCategory['Shared Mailboxes'].UserMigrationBundleUnits | Should -Be 1
            $byCategory['Shared Mailboxes'].NeedsDataObjectCount | Should -Be 1
            $byCategory['Inactive'].MailboxMigrationLicenseUnits | Should -Be 1
            $byCategory['Inactive + archive'].UserMigrationBundleUnits | Should -Be 1
            $byCategory['User Mailboxes'].MailboxMigrationLicenseUnits | Should -Be 1
            $byCategory['User Mailboxes + archive'].UserMigrationBundleUnits | Should -Be 1
            [int](@($mix.MailboxMigrationLicenseUnits | Measure-Object -Sum).Sum) | Should -Be 5
            [int](@($mix.UserMigrationBundleUnits | Measure-Object -Sum).Sum) | Should -Be 3
        }
    }

    Context 'Mailbox-only boundary' {
        It 'exposes no parameter for site or drive rows' {
            # BitTitan migrates mailbox data only. Enforcing that at the signature means the
            # engine cannot be accidentally handed ShareGate scope.
            $command = Get-Command -Name Get-ArrayaBitTitanLicenseEstimate
            $command.Parameters.Keys | Should -Not -Contain 'SharePointRows'
            $command.Parameters.Keys | Should -Not -Contain 'OneDriveRows'
            $command.Parameters.Keys | Should -Not -Contain 'TeamsRows'
        }
    }

    Context 'License model configuration' {
        It 'loads the shipped model with mailbox SKUs only' {
            $model = Get-ArrayaBitTitanLicenseModel

            $model.Skus.MigrationWizMailbox.BlockSizeGB | Should -Be 50
            $model.Skus.MigrationWizMailbox.PassesPerMailbox | Should -Be 10
            $model.Skus.MigrationWizMailbox.Notes | Should -Match 'cannot be pooled'
            $model.Skus.PSObject.Properties.Name | Should -Not -Contain 'Collaboration'
            $model.Skus.PSObject.Properties.Name | Should -Not -Contain 'Document'
            $model.OptionalCommercialOptions.TenantMigrationBundle.DefaultIncluded | Should -BeFalse
            $model.OptionalCommercialOptions.TenantMigrationBundle.Excludes | Should -Contain 'Public folders'
            $model.OptionalCommercialOptions.TenantMigrationBundle.Excludes | Should -Not -Contain 'Teams private chat'
            $model.OptionalCommercialOptions.TenantMigrationBundle.RequiresCurrentValidation | Should -Contain 'Teams private chat (PCH) capability and licensing'
            $model.OptionalCommercialOptions.TenantMigrationBundle.PrivateChatProjectType | Should -Be 'Collaboration (Private Chats)'
            $model.PracticePolicy.GroupMailboxLicensing | Should -Be 'ReportOnly'
            $model.PracticePolicy.EmitDollarCosts | Should -BeFalse
        }

        It 'honours a caller-supplied block size' {
            $model = Get-ArrayaBitTitanLicenseModel
            $model.Skus.MigrationWizMailbox.BlockSizeGB = 100

            $estimate = Get-ArrayaBitTitanLicenseEstimate -MailboxRows @(New-TestMailbox -Name 'A' -Smtp 'a@contoso.com' -SizeGB 150) -Model $model
            (Get-Line -Estimate $estimate -DisplayName 'A').LicenseUnits | Should -Be 2
        }
    }
}
