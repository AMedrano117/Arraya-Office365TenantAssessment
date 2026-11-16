# Group and guest recreation inventories for the tenant-to-tenant profile. Fixtures are
# fictional, use reserved .example domains, and replace Exchange cmdlets with local stubs;
# nothing contacts a tenant.
Describe 'Tenant-to-tenant identity inventory' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        Import-Module -Name (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -WarningAction SilentlyContinue -ErrorAction Stop
        Import-Module -Name (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Arraya.M365.Exchange.psd1') -Force -WarningAction SilentlyContinue -ErrorAction Stop

        function global:Write-Log { param() }
        function global:Write-ProgressHelper { param() }

        function script:New-Step {
            param([string]$Name, [string]$Status = 'Completed', [string]$Message)
            [pscustomobject]@{ Name = $Name; Section = 'Test'; Workload = 'Test'; Status = $Status; Required = $false; Message = $Message; DurationSeconds = 0; ProducedKeys = @() }
        }
        function script:Get-Group {
            param($Inventory, [string]$Name)
            @($Inventory.GroupInventory | Where-Object { $_.DisplayName -eq $Name })[0]
        }

        $script:m365Id = '10000000-0000-0000-0000-000000000001'
        $script:store = @{
            EntraIDGroups     = @{
                $script:m365Id = [pscustomobject]@{ ID = $script:m365Id; DisplayName = 'Projects'; GroupType = 'Microsoft 365'; MembershipType = 'Assigned'; MailEnabled = $true; SecurityEnabled = $false; Mail = 'projects@contoso.example'; Visibility = 'Private'; Source = 'Cloud'; MemberCount = 12; OwnerCount = 2; IsAssignableToRole = $false; IsManagingLicenses = $false }
                'g-2'          = [pscustomobject]@{ ID = 'g-2'; DisplayName = 'Finance Access'; GroupType = 'Security Group'; MembershipType = 'Assigned'; MailEnabled = $false; SecurityEnabled = $true; Source = 'Cloud'; MemberCount = 4; OwnerCount = 0; IsAssignableToRole = $true; IsManagingLicenses = $true; AssignedLicenseFriendlyNames = 'Synthetic E5' }
                'g-3'          = [pscustomobject]@{ ID = 'g-3'; DisplayName = 'Sales Dynamic'; GroupType = 'Security Group'; MembershipType = 'Dynamic Group'; MembershipRule = '(user.department -eq "Sales")'; MailEnabled = $false; SecurityEnabled = $true; Source = 'Cloud'; MemberCount = 'NotCollected (minimum mode)'; OwnerCount = 1 }
                'g-4'          = [pscustomobject]@{ ID = 'g-4'; DisplayName = 'Helpdesk'; GroupType = 'Mail-enabled Security Group'; MembershipType = 'Assigned'; MailEnabled = $true; SecurityEnabled = $true; Mail = 'helpdesk@contoso.example'; Source = 'On-Premises'; OnPremisesSyncEnabled = $true; MemberCount = 3; OwnerCount = 1 }
            }
            AllExchangeGroups = @{
                'Projects'  = [pscustomobject]@{ DisplayName = 'Projects'; ExternalDirectoryObjectId = $script:m365Id; PrimarySMTPAddress = 'projects@contoso.example'; RecipientTypeDetails = 'GroupMailbox'; ResourceProvisioningOptions = 'Team'; Owners = 'avery; blake'; OwnersCount = 2; MembersCount = 12; MemberInventoryStatus = 'Collected'; SharePointSiteUrl = 'https://contoso.sharepoint.example/sites/projects'; AccessType = 'Private'; RequireSenderAuthenticationEnabled = $true }
                'Helpdesk'  = [pscustomobject]@{ DisplayName = 'Helpdesk'; PrimarySMTPAddress = 'helpdesk@contoso.example'; RecipientTypeDetails = 'MailUniversalSecurityGroup'; IsDirSynced = $true; Owners = 'casey'; OwnersCount = 1; MembersCount = 3; MemberInventoryStatus = 'Collected'; RequireSenderAuthenticationEnabled = $false }
                'All Staff' = [pscustomobject]@{ DisplayName = 'All Staff'; PrimarySMTPAddress = 'allstaff@contoso.example'; RecipientTypeDetails = 'DynamicDistributionGroup'; RecipientFilter = "((RecipientType -eq 'UserMailbox'))"; OwnersCount = 0; MembersCount = 40; MemberInventoryStatus = 'Collected'; RequireSenderAuthenticationEnabled = $true }
                'Newsletter' = [pscustomobject]@{ DisplayName = 'Newsletter'; PrimarySMTPAddress = 'news@contoso.example'; RecipientTypeDetails = 'MailUniversalDistributionGroup'; Owners = 'drew'; OwnersCount = 1; MembersCount = 0; MemberInventoryStatus = 'NotCollected'; RequireSenderAuthenticationEnabled = $true; AcceptMessagesOnlyFrom = 'drew' }
            }
            Users             = @(
                [pscustomobject]@{ DisplayName = 'Avery'; UserPrincipalName = 'avery@contoso.example'; Mail = 'avery@contoso.example'; UserType = 'Member'; AccountEnabled = $true }
                [pscustomobject]@{ DisplayName = 'Partner One'; UserPrincipalName = 'p1_fabrikam.example#EXT#@contoso.onmicrosoft.example'; Mail = 'p1@fabrikam.example'; UserType = 'Guest'; AccountEnabled = $true; ExternalUserState = 'Accepted'; SignInActivity = [pscustomobject]@{ LastSignInDateTime = '2026-10-01T00:00:00Z' } }
                [pscustomobject]@{ DisplayName = 'Partner Two'; UserPrincipalName = 'p2_tailspin.example#EXT#@contoso.onmicrosoft.example'; Mail = 'p2_tailspin.example#EXT#@contoso.onmicrosoft.example'; AccountEnabled = $false; ExternalUserState = 'PendingAcceptance' }
            )
        }
        $script:inventory = New-ArrayaTenantToTenantIdentityInventory -TenantStatsHash $script:store
    }

    Context 'Group inventory' {
        It 'lists each group once, joining Entra and Exchange evidence' {
            @($script:inventory.GroupInventory).Count | Should -Be 6
            @($script:inventory.GroupInventory | Where-Object { $_.DisplayName -eq 'Projects' }).Count | Should -Be 1
            @($script:inventory.GroupInventory | Where-Object { $_.DisplayName -eq 'Helpdesk' }).Count | Should -Be 1
        }

        It 'joins a Microsoft 365 group by object id and keeps the Exchange owners, site, and Team flag' {
            $row = Get-Group $script:inventory 'Projects'
            $row.GroupType | Should -Be 'Microsoft 365 group'
            $row.InventorySource | Should -Be 'Entra ID and Exchange Online'
            $row.MailEnabled | Should -BeTrue
            $row.IsTeam | Should -BeTrue
            $row.Owners | Should -Be 'avery; blake'
            $row.SharePointSiteUrl | Should -Be 'https://contoso.sharepoint.example/sites/projects'
            $row.RecreateIn | Should -Be 'Microsoft 365 (Entra ID or Exchange Online)'
            $row.MemberList | Should -Be 'Listed in GroupMembers'
        }

        It 'reports a cloud security group as not mail-enabled, created in Entra ID, with role and license notes' {
            $row = Get-Group $script:inventory 'Finance Access'
            $row.GroupType | Should -Be 'Security group'
            $row.MailEnabled | Should -BeFalse
            $row.RecreateIn | Should -Be 'Entra ID'
            $row.IsAssignableToRole | Should -BeTrue
            $row.AssignsLicenses | Should -BeTrue
            $row.AssignedLicenses | Should -Be 'Synthetic E5'
            $row.RecreationConsiderations | Should -Match 'Role-assignable'
            $row.RecreationConsiderations | Should -Match 'No owner recorded'
        }

        It 'carries an Entra dynamic membership rule and does not treat a text placeholder as a member count' {
            $row = Get-Group $script:inventory 'Sales Dynamic'
            $row.MembershipType | Should -Be 'Dynamic (Entra membership rule)'
            $row.MembershipRule | Should -Be '(user.department -eq "Sales")'
            $row.MemberCount | Should -BeNullOrEmpty
            $row.MemberList | Should -Be 'Not collected'
        }

        It 'joins a synced mail-enabled security group by address and sends it to on-premises AD' {
            $row = Get-Group $script:inventory 'Helpdesk'
            $row.GroupType | Should -Be 'Mail-enabled security group'
            $row.Source | Should -Be 'Synced from on-premises'
            $row.RecreateIn | Should -Be 'On-premises AD (synced)'
            $row.RecreationConsiderations | Should -Match 'external senders'
        }

        It 'includes Exchange-only dynamic distribution lists with their recipient filter' {
            $row = Get-Group $script:inventory 'All Staff'
            $row.GroupType | Should -Be 'Dynamic distribution list'
            $row.MembershipType | Should -Be 'Dynamic (Exchange recipient filter)'
            $row.MembershipRule | Should -Be "((RecipientType -eq 'UserMailbox'))"
            $row.RecreateIn | Should -Be 'Exchange Online'
            $row.InventorySource | Should -Be 'Exchange Online'
            $row.RecreationConsiderations | Should -Not -Match 'No owner'
        }

        It 'does not report a zero member count when membership was not expanded' {
            $row = Get-Group $script:inventory 'Newsletter'
            $row.GroupType | Should -Be 'Distribution list'
            $row.MemberCount | Should -BeNullOrEmpty
            $row.MemberList | Should -Be 'Not collected'
            $row.RecreationConsiderations | Should -Match 'Restricts who can send'
        }

        It 'summarizes types, mail-enabled, dynamic, synced, role-assignable, and license groups' {
            $script:inventory.GroupCountByType['Microsoft 365 group'] | Should -Be 1
            $script:inventory.GroupCountByType['Security group'] | Should -Be 2
            $script:inventory.GroupCountByType['Mail-enabled security group'] | Should -Be 1
            $script:inventory.GroupCountByType['Distribution list'] | Should -Be 1
            $script:inventory.GroupCountByType['Dynamic distribution list'] | Should -Be 1
            $script:inventory.MailEnabledGroupCount | Should -Be 4
            $script:inventory.DynamicGroupCount | Should -Be 2
            $script:inventory.SyncedGroupCount | Should -Be 1
            $script:inventory.RoleAssignableCount | Should -Be 1
            $script:inventory.LicenseGroupCount | Should -Be 1
            $script:inventory.GroupStatus | Should -Be 'Review'
        }
    }

    Context 'Guest inventory' {
        It 'lists guests by UserType and by the #EXT# marker, never members' {
            @($script:inventory.GuestUsers).Count | Should -Be 2
            @($script:inventory.GuestUsers | Where-Object { $_.DisplayName -eq 'Avery' }).Count | Should -Be 0
        }

        It 'takes the home domain from the mail address, or from the UPN when mail is a resource-tenant address' {
            (@($script:inventory.GuestUsers | Where-Object { $_.DisplayName -eq 'Partner One' })[0]).HomeDomain | Should -Be 'fabrikam.example'
            $second = @($script:inventory.GuestUsers | Where-Object { $_.DisplayName -eq 'Partner Two' })[0]
            $second.HomeDomain | Should -Be 'tailspin.example'
            $second.Mail | Should -BeNullOrEmpty
        }

        It 'reads the last sign-in from the sign-in activity and counts pending and enabled guests' {
            (@($script:inventory.GuestUsers | Where-Object { $_.DisplayName -eq 'Partner One' })[0]).LastSignInDateTime | Should -Be '2026-10-01T00:00:00Z'
            $script:inventory.EnabledGuestCount | Should -Be 1
            $script:inventory.PendingGuestCount | Should -Be 1
            $script:inventory.GuestStatus | Should -Be 'Review'
        }
    }

    Context 'Evidence status' {
        It 'reports Needs Data, not Ready, when neither group collector ran' {
            $result = New-ArrayaTenantToTenantIdentityInventory -TenantStatsHash @{ Users = @() }
            $result.GroupStatus | Should -Be 'Needs Data'
            $result.GroupEvidenceGap | Should -Match 'Entra groups'
            $result.GuestStatus | Should -Be 'Needs Data'
        }

        It 'reports Ready only when both group collectors and the user collector completed empty' {
            $result = New-ArrayaTenantToTenantIdentityInventory -TenantStatsHash @{
                EntraIDGroups        = @{}
                AllExchangeGroups    = @{}
                Users                = @{}
                CollectorPlanResults = @((New-Step 'Entra groups'), (New-Step 'Exchange groups'), (New-Step 'Users'))
            }
            $result.GroupStatus | Should -Be 'Ready'
            $result.GroupEvidenceGap | Should -BeNullOrEmpty
            $result.GuestStatus | Should -Be 'Ready'
        }

        It 'keeps the Exchange groups but flags the gap when the Entra collector was skipped' {
            $result = New-ArrayaTenantToTenantIdentityInventory -TenantStatsHash @{
                AllExchangeGroups    = $script:store.AllExchangeGroups
                CollectorPlanResults = @((New-Step 'Entra groups' -Status 'Skipped' -Message 'Entra group inventory is disabled for this profile.'), (New-Step 'Exchange groups'))
            }
            @($result.GroupInventory).Count | Should -Be 4
            $result.GroupStatus | Should -Be 'Review'
            $result.GroupEvidenceGap | Should -Match 'Entra groups'
        }
    }

    Context 'Group membership and contacts' {
        BeforeAll {
            $script:memberStore = @{
                EntraIDGroups        = @{
                    'g-sec' = [pscustomobject]@{ ID = 'g-sec'; DisplayName = 'Finance Access'; GroupType = 'Security Group'; MembershipType = 'Assigned'; MailEnabled = $false; SecurityEnabled = $true; Source = 'Cloud'; MemberCount = 2; OwnerCount = 1; Owners = 'owner@contoso.example'; MemberListState = 'Collected' }
                    'g-m365' = [pscustomobject]@{ ID = 'g-m365'; DisplayName = 'Projects'; GroupType = 'Microsoft 365'; MembershipType = 'Assigned'; MailEnabled = $true; SecurityEnabled = $false; Mail = 'projects@contoso.example'; Source = 'Cloud'; MemberCount = 1; OwnerCount = 1; MemberListState = 'Collected' }
                    'g-fail' = [pscustomobject]@{ ID = 'g-fail'; DisplayName = 'Broken Lookup'; GroupType = 'Security Group'; MembershipType = 'Assigned'; MailEnabled = $false; SecurityEnabled = $true; Source = 'Cloud'; MemberCount = 5; MemberListState = 'LookupFailed' }
                }
                EntraGroupMembers    = @(
                    [pscustomobject]@{ GroupId = 'g-sec'; GroupDisplayName = 'Finance Access'; Relationship = 'Owner'; MemberDisplayName = 'Owner'; MemberAddress = 'owner@contoso.example'; MemberType = 'user'; MembershipSource = 'Assigned' }
                    [pscustomobject]@{ GroupId = 'g-sec'; GroupDisplayName = 'Finance Access'; Relationship = 'Member'; MemberDisplayName = 'Avery'; MemberAddress = 'avery@contoso.example'; MemberType = 'user'; MembershipSource = 'Assigned' }
                    [pscustomobject]@{ GroupId = 'g-m365'; GroupDisplayName = 'Projects'; GroupMail = 'projects@contoso.example'; Relationship = 'Member'; MemberDisplayName = 'Blake'; MemberAddress = 'blake@contoso.example'; MemberType = 'user'; MembershipSource = 'Assigned' }
                )
                AllExchangeGroups    = @{
                    'projects' = [pscustomobject]@{ DisplayName = 'Projects'; ExternalDirectoryObjectId = 'g-m365'; PrimarySMTPAddress = 'projects@contoso.example'; RecipientTypeDetails = 'GroupMailbox'; MembersCount = 1; MemberInventoryStatus = 'Collected'; SharePointSiteUrl = 'https://contoso.sharepoint.example/sites/projects/' }
                    'sales'    = [pscustomobject]@{ DisplayName = 'Sales'; PrimarySMTPAddress = 'sales@contoso.example'; RecipientTypeDetails = 'MailUniversalDistributionGroup'; MembersCount = 2; MemberInventoryStatus = 'Collected'; Owners = 'avery' }
                }
                ExchangeGroupMembers = @(
                    [pscustomobject]@{ GroupDisplayName = 'Projects'; GroupPrimarySmtpAddress = 'projects@contoso.example'; GroupRecipientTypeDetails = 'GroupMailbox'; MemberDisplayName = 'Blake'; MemberPrimarySmtpAddress = 'blake@contoso.example'; MemberRecipientTypeDetails = 'UserMailbox'; MembershipSource = 'Assigned' }
                    [pscustomobject]@{ GroupDisplayName = 'Sales'; GroupPrimarySmtpAddress = 'sales@contoso.example'; GroupRecipientTypeDetails = 'MailUniversalDistributionGroup'; MemberDisplayName = 'Avery'; MemberPrimarySmtpAddress = 'avery@contoso.example'; MemberRecipientTypeDetails = 'UserMailbox'; MembershipSource = 'Assigned' }
                    [pscustomobject]@{ GroupDisplayName = 'Sales'; GroupPrimarySmtpAddress = 'sales@contoso.example'; GroupRecipientTypeDetails = 'MailUniversalDistributionGroup'; MemberDisplayName = 'Vendor'; MemberPrimarySmtpAddress = 'vendor@fabrikam.example'; MemberRecipientTypeDetails = 'MailContact'; MembershipSource = 'Assigned' }
                )
                AllTeams             = @([pscustomobject]@{ DisplayName = 'Projects'; SharePointSiteUrl = 'https://contoso.sharepoint.example/sites/projects' })
                MailContacts         = @{
                    'vendor@fabrikam.example' = [pscustomobject]@{ DisplayName = 'Vendor'; RecipientTypeDetails = 'MailContact'; PrimarySmtpAddress = 'vendor@fabrikam.example'; ExternalEmailAddress = 'vendor@fabrikam.example'; IsDirSynced = $false }
                    'alumni@contoso.example'  = [pscustomobject]@{ DisplayName = 'Alumni Forward'; RecipientTypeDetails = 'MailUser'; PrimarySmtpAddress = 'alumni@contoso.example'; ExternalEmailAddress = 'alumni@contoso.example'; IsDirSynced = $true }
                }
                Domains              = @{ 'contoso.example' = [pscustomobject]@{ Name = 'contoso.example' } }
            }
            $script:memberInventory = New-ArrayaTenantToTenantIdentityInventory -TenantStatsHash $script:memberStore
        }

        It 'merges Graph and Exchange member lists without repeating a group Graph already listed' {
            $rows = @($script:memberInventory.GroupMembers)
            @($rows | Where-Object { $_.GroupDisplayName -eq 'Projects' }).Count | Should -Be 1
            (@($rows | Where-Object { $_.GroupDisplayName -eq 'Projects' })[0]).Source | Should -Be 'Entra ID'
            @($rows | Where-Object { $_.GroupDisplayName -eq 'Sales' }).Count | Should -Be 2
            (@($rows | Where-Object { $_.GroupDisplayName -eq 'Sales' })[0]).GroupType | Should -Be 'Distribution list'
            @($rows | Where-Object { $_.GroupDisplayName -eq 'Finance Access' -and $_.Relationship -eq 'Owner' }).Count | Should -Be 1
        }

        It 'uses Graph owners for a security group and reports a failed Graph lookup' {
            $finance = Get-Group $script:memberInventory 'Finance Access'
            $finance.Owners | Should -Be 'owner@contoso.example'
            $finance.MemberList | Should -Be 'Listed in GroupMembers'
            (Get-Group $script:memberInventory 'Broken Lookup').MemberList | Should -Be 'Lookup failed'
        }

        It 'recognizes a Team from the Teams inventory when the group has no provisioning options' {
            (Get-Group $script:memberInventory 'Projects').IsTeam | Should -BeTrue
        }

        It 'lists contacts and mail users with the groups they belong to' {
            $vendor = @($script:memberInventory.ContactInventory | Where-Object { $_.DisplayName -eq 'Vendor' })[0]
            $vendor.ContactType | Should -Be 'Mail contact'
            $vendor.MemberOfGroups | Should -Be 'Sales'
            $vendor.RecreateIn | Should -Be 'Exchange Online'
            $vendor.RecreationConsiderations | Should -Match 'Member of 1 group'
            $script:memberInventory.MailContactCount | Should -Be 1
            $script:memberInventory.MailUserCount | Should -Be 1
            $script:memberInventory.ContactStatus | Should -Be 'Review'
        }

        It 'flags a mail user, a synced object, and an external address in a domain the tenant owns' {
            $alumni = @($script:memberInventory.ContactInventory | Where-Object { $_.DisplayName -eq 'Alumni Forward' })[0]
            $alumni.ContactType | Should -Be 'Mail user'
            $alumni.RecreateIn | Should -Be 'On-premises AD (synced)'
            $alumni.RecreationConsiderations | Should -Match 'sign-in account'
            $alumni.RecreationConsiderations | Should -Match 'contoso\.example, a domain this tenant owns'
        }

        It 'reports contacts as Needs Data when the contact collector did not run' {
            (New-ArrayaTenantToTenantIdentityInventory -TenantStatsHash @{ Users = @() }).ContactStatus | Should -Be 'Needs Data'
        }
    }

    Context 'Entra group collector member lists' {
        BeforeAll {
            Import-Module -Name (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Graph\Arraya.M365.Graph.psd1') -Force -WarningAction SilentlyContinue -ErrorAction Stop
            $script:graphGroups = @(
                [pscustomobject]@{ id = 'sec-1'; displayName = 'Finance Access'; groupTypes = @(); mailEnabled = $false; securityEnabled = $true }
                [pscustomobject]@{ id = 'm365-1'; displayName = 'Projects'; groupTypes = @('Unified'); mailEnabled = $true; securityEnabled = $false; mail = 'projects@contoso.example' }
                [pscustomobject]@{ id = 'dl-1'; displayName = 'Sales'; groupTypes = @(); mailEnabled = $true; securityEnabled = $false; mail = 'sales@contoso.example' }
                [pscustomobject]@{ id = 'fail-1'; displayName = 'Broken'; groupTypes = @(); mailEnabled = $false; securityEnabled = $true }
            )
            Mock -ModuleName Arraya.M365.Graph Invoke-ArrayaGraphCollectionRequest { $script:graphGroups }
            Mock -ModuleName Arraya.M365.Graph Write-Host { }
            Mock -ModuleName Arraya.M365.Graph Invoke-ArrayaGraphCollectionBatch {
                $result = @{}
                foreach ($request in $Requests) {
                    $id = [string]$request.id
                    if ($id -like 'l*:fail-1') {
                        $result[$id] = [pscustomobject]@{ id = $id; status = 403; body = $null }
                        continue
                    }
                    $body = switch -Wildcard ($id) {
                        'm:*' { 1 }
                        'o:*' { 1 }
                        'lm:sec-1' { [pscustomobject]@{ value = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'u1'; displayName = 'Avery'; userPrincipalName = 'avery@contoso.example' }, [pscustomobject]@{ '@odata.type' = '#microsoft.graph.group'; id = 'g9'; displayName = 'Nested'; mail = $null }) } }
                        'lo:sec-1' { [pscustomobject]@{ value = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = 'u2'; displayName = 'Owner'; userPrincipalName = 'owner@contoso.example' }) } }
                        'lm:m365-1' { [pscustomobject]@{ value = @() } }
                        'lo:m365-1' { [pscustomobject]@{ value = @() } }
                        default { [pscustomobject]@{ value = @() } }
                    }
                    $result[$id] = [pscustomobject]@{ id = $id; status = 200; body = $body }
                }
                $result
            }

            $depth = [pscustomobject]@{ CollectEntraGroupDeepDetails = $false; CollectEntraGroupMemberCounts = $true; CollectEntraGroupOwnerCounts = $true; CollectEntraGroupLicenseChecks = $false; CollectEntraGroupMemberLists = $true; IsMinimum = $false }
            $context = New-ArrayaAssessmentContext -ReportingMode 'All' -Policies @{ ReportingMode = 'All'; CollectionDepth = $depth }
            Get-EntraIDGroups -detailLevel 'all' -GraphAuthType REST -Context $context | Out-Null
            $script:graphStats = $context.TenantStats
        }

        It 'lists members and owners of security and Microsoft 365 groups, not of distribution lists' {
            $rows = @($script:graphStats['EntraGroupMembers'])
            @($rows | Where-Object { $_.GroupId -eq 'sec-1' -and $_.Relationship -eq 'Member' }).Count | Should -Be 2
            @($rows | Where-Object { $_.GroupId -eq 'sec-1' -and $_.Relationship -eq 'Owner' })[0].MemberAddress | Should -Be 'owner@contoso.example'
            @($rows | Where-Object { $_.MemberDisplayName -eq 'Nested' })[0].MemberType | Should -Be 'group'
            @($rows | Where-Object { $_.GroupId -eq 'dl-1' }).Count | Should -Be 0
        }

        It 'records the list state per group, including an empty list and a failed lookup' {
            $script:graphStats['EntraIDGroups']['sec-1'].MemberListState | Should -Be 'Collected'
            $script:graphStats['EntraIDGroups']['sec-1'].Owners | Should -Be 'owner@contoso.example'
            $script:graphStats['EntraIDGroups']['m365-1'].MemberListState | Should -Be 'Collected'
            $script:graphStats['EntraIDGroups']['dl-1'].MemberListState | Should -Be 'ListedByExchange'
            $script:graphStats['EntraIDGroups']['fail-1'].MemberListState | Should -Be 'LookupFailed'
        }
    }

    Context 'Exchange mail contact collector' {
        BeforeAll {
            function global:Get-EXORecipient {
                param($RecipientTypeDetails, $Properties, $ResultSize)
                [pscustomobject]@{ DisplayName = 'Vendor'; Alias = 'vendor'; RecipientTypeDetails = 'MailContact'; PrimarySmtpAddress = 'vendor@fabrikam.example'; ExternalEmailAddress = 'SMTP:vendor@fabrikam.example'; EmailAddresses = @('SMTP:vendor@fabrikam.example'); IsDirSynced = $false; ExternalDirectoryObjectId = 'c-1' }
            }
            $context = New-ArrayaAssessmentContext -ReportingMode 'All' -TenantStats @{}
            Get-ExchangeMailContactInventory -detailLevel 'all' -Context $context 6>$null | Out-Null
            $script:contactStats = $context.TenantStats
        }

        AfterAll {
            Remove-Item -Path 'Function:\global:Get-EXORecipient' -ErrorAction SilentlyContinue
        }

        It 'stores contacts with the SMTP prefix removed from the external address and records complete evidence' {
            $row = @($script:contactStats['MailContacts'].Values)[0]
            $row.ExternalEmailAddress | Should -Be 'vendor@fabrikam.example'
            $row.RecipientTypeDetails | Should -Be 'MailContact'
            (Get-ArrayaDatasetEvidenceState -TenantStatsHash $script:contactStats -DatasetKey 'MailContacts').State | Should -Be 'Collected'
        }
    }

    Context 'Exchange group collector' {
        BeforeAll {
            function global:Get-DistributionGroup { param($Identity) [pscustomobject]@{ IsDirSynced = $false; RequireSenderAuthenticationEnabled = $true; MemberJoinRestriction = 'Closed'; MemberDepartRestriction = 'Closed'; ModerationEnabled = $true; ModeratedBy = @('drew') } }
            function global:Get-DistributionGroupMember {
                param($Identity)
                if ($Identity -eq 'broken') { throw 'synthetic member lookup failure' }
                if ($Identity -eq 'empty') { return }
                [pscustomobject]@{ DisplayName = 'Avery'; PrimarySmtpAddress = 'avery@contoso.example'; RecipientTypeDetails = 'UserMailbox' }
                [pscustomobject]@{ DisplayName = 'Blake'; PrimarySmtpAddress = 'blake@contoso.example'; RecipientTypeDetails = 'UserMailbox' }
            }
            function global:Get-DynamicDistributionGroup { param($Identity) [pscustomobject]@{ RecipientFilter = "((Department -eq 'Sales'))"; RequireSenderAuthenticationEnabled = $true } }
            function global:Get-DynamicDistributionGroupMember { param($Identity) [pscustomobject]@{ DisplayName = 'Casey'; PrimarySmtpAddress = 'casey@contoso.example'; RecipientTypeDetails = 'UserMailbox' } }

            $recipients = @{
                'sales@contoso.example'  = [pscustomobject]@{ Identity = 'sales'; DisplayName = 'Sales'; PrimarySMTPAddress = 'sales@contoso.example'; RecipientTypeDetails = 'MailUniversalDistributionGroup'; ManagedBy = 'avery,blake'; ExternalDirectoryObjectId = 'g-sales' }
                'empty@contoso.example'  = [pscustomobject]@{ Identity = 'empty'; DisplayName = 'Empty'; PrimarySMTPAddress = 'empty@contoso.example'; RecipientTypeDetails = 'MailUniversalDistributionGroup'; ManagedBy = '' }
                'broken@contoso.example' = [pscustomobject]@{ Identity = 'broken'; DisplayName = 'Broken'; PrimarySMTPAddress = 'broken@contoso.example'; RecipientTypeDetails = 'MailUniversalDistributionGroup'; ManagedBy = 'casey' }
                'dyn@contoso.example'    = [pscustomobject]@{ Identity = 'dyn'; DisplayName = 'Sales Everyone'; PrimarySMTPAddress = 'dyn@contoso.example'; RecipientTypeDetails = 'DynamicDistributionGroup'; ManagedBy = '' }
                'avery@contoso.example'  = [pscustomobject]@{ Identity = 'avery'; DisplayName = 'Avery'; PrimarySMTPAddress = 'avery@contoso.example'; RecipientTypeDetails = 'UserMailbox' }
            }
            $context = New-ArrayaAssessmentContext -ReportingMode 'All' -TenantStats @{ AllRecipients = $recipients }
            Get-ExchangeGroupDetails -detailLevel 'all' -Context $context 6>$null | Out-Null
            $script:collectorStats = $context.TenantStats
            $script:groups = $script:collectorStats['AllExchangeGroups']
        }

        AfterAll {
            foreach ($name in @('Get-DistributionGroup', 'Get-DistributionGroupMember', 'Get-DynamicDistributionGroup', 'Get-DynamicDistributionGroupMember')) {
                Remove-Item -Path ("Function:\global:{0}" -f $name) -ErrorAction SilentlyContinue
            }
        }

        It 'splits the joined ManagedBy string into owners instead of counting it as one' {
            $sales = $script:groups['sales']
            $sales.OwnersCount | Should -Be 2
            $sales.Owners | Should -Be 'avery; blake'
            $sales.ExternalDirectoryObjectId | Should -Be 'g-sales'
        }

        It 'keeps the moderation and join settings from the group detail' {
            $sales = $script:groups['sales']
            $sales.ModeratedBy | Should -Be 'drew'
            $sales.ModerationEnabled | Should -BeTrue
            $sales.MemberJoinRestriction | Should -Be 'Closed'
        }

        It 'records the member list and tells an empty group from a failed lookup' {
            $script:groups['sales'].MemberInventoryStatus | Should -Be 'Collected'
            $script:groups['sales'].MembersCount | Should -Be 2
            $script:groups['empty'].MemberInventoryStatus | Should -Be 'Collected'
            $script:groups['empty'].MembersCount | Should -Be 0
            $script:groups['broken'].MemberInventoryStatus | Should -Be 'LookupFailed'

            $memberRows = @($script:collectorStats['ExchangeGroupMembers'])
            @($memberRows | Where-Object { $_.GroupPrimarySmtpAddress -eq 'sales@contoso.example' } | ForEach-Object { $_.MemberPrimarySmtpAddress }) | Should -Be @('avery@contoso.example', 'blake@contoso.example')
            @($memberRows | Where-Object { $_.GroupPrimarySmtpAddress -eq 'broken@contoso.example' }).Count | Should -Be 0
        }

        It 'keeps the dynamic distribution list recipient filter and labels its computed members' {
            $script:groups['dyn'].RecipientFilter | Should -Be "((Department -eq 'Sales'))"
            $dynamicRow = @($script:collectorStats['ExchangeGroupMembers'] | Where-Object { $_.GroupPrimarySmtpAddress -eq 'dyn@contoso.example' })[0]
            $dynamicRow.MembershipSource | Should -Be 'Computed from recipient filter'
        }
    }
}
