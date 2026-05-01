Describe 'Improve workflow' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:commonManifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        $script:improveScriptPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\New-M365TenantImprovementPlan.ps1'

        Import-Module -Name $script:commonManifestPath -Force -ErrorAction Stop

        function Get-TestDocxDocumentXmlText {
            param([Parameter(Mandatory = $true)][string]$Path)

            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
            try {
                $entry = $archive.GetEntry('word/document.xml')
                if ($null -eq $entry) {
                    throw "DOCX did not contain word/document.xml: $Path"
                }

                $reader = New-Object System.IO.StreamReader($entry.Open())
                try {
                    return $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $archive.Dispose()
            }
        }

        function Get-TestDocxDocumentXml {
            param([Parameter(Mandatory = $true)][string]$Path)

            [xml](Get-TestDocxDocumentXmlText -Path $Path)
        }

        function Get-TestDocxEntryNames {
            param([Parameter(Mandatory = $true)][string]$Path)

            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
            try {
                return @($archive.Entries | ForEach-Object { $_.FullName })
            }
            finally {
                $archive.Dispose()
            }
        }

        function New-TestEnterpriseApplicationSnapshot {
            param(
                [Parameter(Mandatory = $true)]
                $EnterpriseApplications,
                [Parameter(Mandatory = $false)]
                $EnterpriseApplicationSummary,
                [Parameter(Mandatory = $false)]
                $AuthenticationSsoApplications = @()
            )

            function Convert-TestApplicationCollection {
                param(
                    [Parameter(Mandatory = $false)]
                    $Rows
                )

                if ($null -eq $Rows) {
                    return @{}
                }

                if ($Rows -is [System.Collections.IDictionary]) {
                    return $Rows
                }

                $normalized = [ordered]@{}
                $rowIndex = 0
                foreach ($row in @(Convert-ArrayaObjectToArray $Rows)) {
                    if ($null -eq $row) {
                        continue
                    }

                    $displayName = [string](Get-ArrayaObjectValue -Object $row -Names @('DisplayName'))
                    if ([string]::IsNullOrWhiteSpace($displayName)) {
                        $displayName = ('App-{0:D3}' -f ($rowIndex + 1))
                    }

                    $normalized[('{0:D3}-{1}' -f ($rowIndex + 1), $displayName)] = $row
                    $rowIndex++
                }

                return $normalized
            }

            $identityData = @{
                Admins = @()
                Users = @()
                ConditionalAccessPolicies = @()
                ConditionalAccessPolicySummary = @{
                    Summary = [pscustomobject]@{
                        TotalPolicies                 = 0
                        EnabledPolicies               = 0
                        ReportOnlyPolicies            = 0
                        HasGuestCoverage              = $false
                        HasPrivilegedRoleCoverage     = $false
                        HasCompliantDeviceRequirement = $false
                        HasRiskBasedCoverage          = $false
                        PoliciesWithExclusions        = 0
                    }
                }
                SecurityDefaultsPolicy = @{
                    Summary = [pscustomobject]@{
                        IsEnabled = $false
                    }
                }
                AuthenticationConfig = [pscustomobject]@{
                    AdminConsentWorkflowEnabled   = $false
                    PermissionGrantPoliciesAssigned = @()
                    PasswordlessMethods           = @()
                }
                EnterpriseApplications = Convert-TestApplicationCollection -Rows $EnterpriseApplications
                AuthenticationSSOApplications = Convert-TestApplicationCollection -Rows $AuthenticationSsoApplications
            }

            if ($null -ne $EnterpriseApplicationSummary) {
                $identityData['EnterpriseApplicationSummary'] = @{
                    Summary = [pscustomobject]$EnterpriseApplicationSummary
                }
            }

            return New-ArrayaTenantSnapshot -Data @{
                Identity = $identityData
            }
        }
    }

    It 'generates normalized findings, customer deliverables, and fixed remediation snippets' {
        $snapshot = New-ArrayaTenantSnapshot `
            -Data @{
                Identity = @{
                    Admins = @(
                        [pscustomobject]@{
                            Role                = 'Global Administrator'
                            AccountEnabled      = $true
                            OnPremisesSyncEnabled = $false
                            LastSignInDateTime  = (Get-Date).AddDays(-120).ToString('o')
                        },
                        [pscustomobject]@{
                            Role                = 'Global Administrator'
                            AccountEnabled      = $true
                            OnPremisesSyncEnabled = $false
                            LastSignInDateTime  = (Get-Date).AddDays(-121).ToString('o')
                        },
                        [pscustomobject]@{
                            Role                = 'Global Administrator'
                            AccountEnabled      = $true
                            OnPremisesSyncEnabled = $false
                            LastSignInDateTime  = (Get-Date).AddDays(-122).ToString('o')
                        },
                        [pscustomobject]@{
                            Role                = 'Global Administrator'
                            AccountEnabled      = $true
                            OnPremisesSyncEnabled = $false
                            LastSignInDateTime  = (Get-Date).AddDays(-123).ToString('o')
                        },
                        [pscustomobject]@{
                            Role                = 'Global Administrator'
                            AccountEnabled      = $true
                            OnPremisesSyncEnabled = $false
                            LastSignInDateTime  = (Get-Date).AddDays(-124).ToString('o')
                        },
                        [pscustomobject]@{
                            Role                = 'Global Administrator'
                            AccountEnabled      = $true
                            OnPremisesSyncEnabled = $false
                            LastSignInDateTime  = (Get-Date).AddDays(-125).ToString('o')
                        }
                    )
                    Users = @(
                        [pscustomobject]@{
                            UserType          = 'Guest'
                            AccountEnabled    = $true
                            LastSignInDateTime = (Get-Date).AddDays(-200).ToString('o')
                            AssignedLicensesFriendly = @()
                        }
                    )
                    ConditionalAccessPolicies = @(
                        [pscustomobject]@{ DisplayName = 'Baseline MFA'; State = 'enabled' },
                        [pscustomobject]@{ DisplayName = 'Admins report-only'; State = 'reportOnly' }
                    )
                    LicenseSKUs = @(
                        [pscustomobject]@{
                            SkuPartNumber = 'ENTERPRISEPACK'
                            ConsumedUnits = 99
                            ActiveUnits   = 100
                        }
                    )
                    DeviceDetails = @(
                        [pscustomobject]@{
                            ApproximateLastSignInDateTime = (Get-Date).AddDays(-200).ToString('o')
                            IsCompliant                   = $false
                            MDMSolution                   = ''
                        }
                    )
                    AuthenticationConfig = [pscustomobject]@{
                        AdminConsentWorkflowEnabled = $false
                        PermissionGrantPoliciesAssigned = @('ManagePermissionGrantsForSelf.microsoft-user-default-allow-consent-apps')
                        PasswordlessMethods = @()
                    }
                    ConditionalAccessPolicySummary = @{
                        Summary = [pscustomobject]@{
                            TotalPolicies                 = 8
                            EnabledPolicies               = 1
                            ReportOnlyPolicies            = 1
                            HasGuestCoverage              = $true
                            HasPrivilegedRoleCoverage     = $true
                            HasCompliantDeviceRequirement = $false
                            HasRiskBasedCoverage          = $true
                            PoliciesWithExclusions        = 7
                        }
                    }
                    SecurityDefaultsPolicy = @{
                        Summary = [pscustomobject]@{
                            IsEnabled = $false
                        }
                    }
                    EnterpriseApplications = @{
                        '001-HighPrivApp' = [pscustomobject]@{
                            AppId                        = '11111111-1111-1111-1111-111111111111'
                            ServicePrincipalId           = '22222222-2222-2222-2222-222222222222'
                            DisplayName                   = 'High Priv App'
                            ApplicationSource             = 'First Party'
                            ApplicationSourceState        = 'Collected'
                            SsoEnabled                    = $true
                            PreferredSingleSignOnMode     = 'saml'
                            SSOMode                       = 'saml'
                            HighPrivilegePermissionCount  = 2
                            HighPrivilegePermissions      = 'Application.ReadWrite.All,Directory.ReadWrite.All'
                            ApplicationPermissionCount    = 1
                            DelegatedPermissionGrantCount = 1
                            AppRoleAssignmentRequired     = $true
                            OwnerSignalState              = 'Collected'
                            OwnerCount                    = 0
                            AppCredentials                = 'Client Secret'
                            CredentialReviewState         = 'Collected'
                            HasExpiredCredentials         = $true
                            ExpiredCredentialCount        = 1
                            ExpiredPasswordCredentialCount = 1
                            HasCredentialsExpiringSoon    = $false
                            ExpiringCredentialCount       = 0
                            CredentialIssueSummary        = '1 expired client secret'
                            RedirectUriSignalState        = 'Collected'
                            InsecureRedirectUriCount      = 1
                            HasInsecureRedirectUris       = $true
                            ActivitySignalState           = 'Collected'
                            HasRecentActivity             = $false
                            DelegatedLastSignIn           = $null
                            ApplicationLastSignIn         = $null
                            LastSignInDateTime            = '2026-04-12T09:15:00Z'
                            LastSignInUserDisplayName     = 'Adele Vance'
                            LastSignInUserPrincipalName   = 'adele.vance@contoso.com'
                            LastConditionalAccessStatus   = 'success'
                            LastClientAppUsed             = 'Browser'
                        }
                        '002-ThirdPartyApp' = [pscustomobject]@{
                            AppId                        = '33333333-3333-3333-3333-333333333333'
                            ServicePrincipalId           = '44444444-4444-4444-4444-444444444444'
                            DisplayName                  = 'Third Party App'
                            ApplicationSource            = 'Third Party'
                            ApplicationSourceState       = 'Collected'
                            SsoEnabled                   = $false
                            HighPrivilegePermissionCount = 0
                            ApplicationPermissionCount   = 1
                            DelegatedPermissionGrantCount = 0
                            OwnerSignalState             = 'NotApplicable'
                            OwnerCount                   = $null
                            RedirectUriSignalState       = 'Partial'
                            HasInsecureRedirectUris      = $null
                            ActivitySignalState          = 'Collected'
                            HasRecentActivity            = $true
                        }
                    }
                    EnterpriseApplicationSummary = @{
                        Summary = [pscustomobject]@{
                            TotalEnterpriseApplications   = 2
                            FirstPartyApplications        = 1
                            ThirdPartyApplications        = 1
                            UnknownSourceApplications     = 0
                            ApplicationsWithHighPrivilege = 1
                            FirstPartyAppsWithoutOwners   = 1
                            ThirdPartyAppsWithApplicationPerms = 1
                            ApplicationsWithInsecureRedirectUris = 1
                            ApplicationsWithNoRecentActivity = 1
                            SsoEnabledApplications        = 1
                            ApplicationsWithExpiredCredentials = 1
                            ApplicationsWithCredentialsExpiringSoon = 0
                            ApplicationSourceCoverageState = 'Collected'
                            ApplicationSourceCollectedCount = 2
                            ApplicationSourceUnavailableCount = 0
                            OwnerSignalCoverageState      = 'Collected'
                            OwnerSignalCollectedCount     = 1
                            OwnerSignalUnavailableCount   = 0
                            OwnerSignalNotApplicableCount = 1
                            RedirectUriSignalCoverageState = 'Partial'
                            RedirectUriSignalCollectedCount = 1
                            RedirectUriSignalPartialCount = 1
                            RedirectUriSignalUnavailableCount = 0
                            ActivitySignalCoverageState   = 'Collected'
                            ActivitySignalCollectedCount  = 2
                            ActivitySignalPartialCount    = 0
                            ActivitySignalUnavailableCount = 0
                        }
                    }
                    GuestSignInSummary = @{
                        Summary = [pscustomobject]@{
                            InactiveGuests90Days = 12
                        }
                    }
                    ExternalIdentityRestrictions = @{
                        Summary = [pscustomobject]@{
                            AllowInvitesFrom           = 'adminsAndGuestInviters'
                            CrossTenantPartnerCount    = 2
                            HasCrossTenantAccessPolicy = $true
                            DefaultInboundMfaTrust     = $true
                            DefaultOutboundMfaTrust    = 'Not configured'
                        }
                    }
                    GuestAccessConfiguration = @{
                        Summary = [pscustomobject]@{
                            GuestInvitationControl      = 'adminsAndGuestInviters'
                            ConditionalAccessGuestCoverage = $true
                        }
                    }
                    DeviceManagementSummary = @{
                        Summary = [pscustomobject]@{
                            TotalDevices         = 10
                            UnmanagedDevices     = 6
                            UnsupportedOsDevices = 2
                        }
                    }
                    MfaRegistrationSummary = [pscustomobject]@{
                        TotalUsers          = 10
                        RegisteredUsers     = 3
                        NotRegisteredUsers  = 7
                        RegistrationPercent = 30
                        MethodCounts        = @{ 'Microsoft Authenticator' = 2; 'SMS / phone' = 2 }
                    }
                    MfaEnrollmentSummary = [pscustomobject]@{
                        TotalUsers                        = 10
                        RegisteredUsers                   = 3
                        NotRegisteredUsers                = 7
                        RegistrationPercent               = 30
                        RegisteredMethodBreakdown         = 'Microsoft Authenticator=2; s of tw ar eO ne Ti me Pa ss co de=2; SMS / phone=2'
                        WeakMethodBreakdown               = 'SMS / phone=2'
                        PhishingResistantMethodBreakdown  = 'FIDO2 security key / passkey=1'
                        UsersWithWeakMethodsOnly          = 1
                        UsersWithWeakDefaultMethod        = 1
                        UsersWithPhishingResistantMethods = 1
                    }
                    MfaRegistrationDetails = @(
                        [pscustomobject]@{
                            UserPrincipalName             = 'registered.member@contoso.com'
                            IsMfaRegistered               = $true
                            IsMfaCapable                  = $true
                            MethodsRegistered             = @('Microsoft Authenticator', 'SMS / phone')
                            DefaultMfaMethod              = 'Microsoft Authenticator'
                            HasWeakMethod                 = $true
                            HasStrongMethod               = $true
                            HasPhishingResistantMethod    = $false
                        },
                        [pscustomobject]@{
                            UserPrincipalName             = 'unregistered.member@contoso.com'
                            IsMfaRegistered               = $false
                            IsMfaCapable                  = $false
                            MethodsRegistered             = @()
                            DefaultMfaMethod              = ''
                            HasWeakMethod                 = $false
                            HasStrongMethod               = $false
                            HasPhishingResistantMethod    = $false
                        }
                    )
                    MfaEnforcementSummary = [pscustomobject]@{
                        ConditionalAccessPoliciesReviewed   = 2
                        EnabledPoliciesRequiringMfa         = 1
                        ReportOnlyPoliciesRequiringMfa      = 1
                        PoliciesWithExclusions              = 2
                        EnabledUsersReviewed                = 80
                        UsersCoveredByEnabledMfaPolicies    = 60
                        UserCoveragePercent                 = 75
                        EnabledMemberUsersReviewed          = 70
                        MemberUsersCoveredByEnabledMfaPolicies = 56
                        MemberUserCoveragePercent           = 80
                        EnabledGuestUsersReviewed           = 10
                        GuestUsersCoveredByEnabledMfaPolicies = 4
                        GuestUserCoveragePercent            = 40
                        CoverageCalculationNote             = 'Estimate is based on enabled reviewed users and enabled Conditional Access policies that require MFA, expanded across direct users, targeted groups, targeted roles, and guest/external-user scope where supported.'
                        GuestOrExternalCoverage             = $true
                        PrivilegedRoleCoverage              = $false
                        RiskBasedCoverage                   = $true
                        CompliantDeviceRequirement          = $false
                        SecurityDefaultsEnabled             = $false
                        EnforcementState                    = 'MFA enforcement is active through enabled Conditional Access policies.'
                        GuestUserEnforcementState           = 'Guest users are partially covered by the active MFA enforcement baseline.'
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
                        },
                        [pscustomobject]@{
                            DisplayName        = 'Excluded Member'
                            UserPrincipalName  = 'excluded.member@contoso.com'
                            UserType           = 'Member'
                            DirectoryObjectId  = 'user-002'
                            GapCategory        = 'Excluded from all enabled MFA CA policies that otherwise target the user'
                            GapReason          = 'Excluded through group ''Break Glass Exclusions'' in enabled MFA policy ''Baseline MFA''.'
                            RelatedPolicies    = 'Baseline MFA'
                            AccountEnabled     = $true
                            LastSignInDateTime = '2026-01-05T00:00:00Z'
                        },
                        [pscustomobject]@{
                            DisplayName        = 'Uncovered Guest'
                            UserPrincipalName  = 'uncovered.guest_contoso.com#EXT#@contoso.onmicrosoft.com'
                            UserType           = 'Guest'
                            DirectoryObjectId  = 'guest-001'
                            GapCategory        = 'Outside enabled MFA CA include scope'
                            GapReason          = 'Guest user is outside the include scope of all enabled Conditional Access policies that currently require MFA.'
                            RelatedPolicies    = ''
                            AccountEnabled     = $true
                            LastSignInDateTime = '2026-01-01T00:00:00Z'
                        }
                    )
                    MfaEnforcementScopeReview = @(
                        [pscustomobject]@{
                            PolicyName           = 'Baseline MFA'
                            ScopeType            = 'Exclude'
                            ObjectType           = 'Group'
                            DisplayName          = 'Break Glass Exclusions'
                            Identifier           = 'group-001'
                            AffectedEnabledUsers = 2
                            DirectoryMemberCount = 2
                            Notes                = 'Members of this group are excluded from the enabled MFA enforcement policy scope.'
                        }
                    )
                    AdminMfaSummary = [pscustomobject]@{
                        EnabledAdminUsersReviewed            = 6
                        AdminUsersRegisteredForMfa           = 4
                        AdminUsersNotRegisteredForMfa        = 2
                        AdminUsersCoveredByMfaEnforcement    = 5
                        AdminUsersNotCoveredByMfaEnforcement = 1
                    }
                    AdminMfaRegistrationGaps = @(
                        [pscustomobject]@{
                            DisplayName          = 'Stale Global Admin'
                            UserPrincipalName    = 'stale.admin@contoso.com'
                            Role                 = 'Global Administrator'
                            AccountEnabled       = $true
                            UserType             = 'Member'
                            MfaRegistrationState = 'Not registered'
                            MethodsRegistered    = ''
                            DefaultMfaMethod     = ''
                            LastSignInDateTime   = '2026-01-02T00:00:00Z'
                        }
                    )
                    AdminMfaEnforcementGaps = @(
                        [pscustomobject]@{
                            DisplayName         = 'Excluded Global Admin'
                            UserPrincipalName   = 'excluded.admin@contoso.com'
                            Role                = 'Global Administrator'
                            AccountEnabled      = $true
                            UserType            = 'Member'
                            MfaEnforcementState = 'Not covered by the current enabled MFA enforcement baseline'
                            GapCategory         = 'Excluded from all enabled MFA CA policies that otherwise target the user'
                            GapReason           = 'Excluded through group ''Break Glass Exclusions'' in enabled MFA policy ''Baseline MFA''.'
                            RelatedPolicies     = 'Baseline MFA'
                            LastSignInDateTime  = '2026-01-03T00:00:00Z'
                        }
                    )
                }
                Exchange = @{
                    AllRecipients = @(
                        [pscustomobject]@{
                            DisplayName          = 'Forwarded Mailbox'
                            RecipientTypeDetails = 'UserMailbox'
                            PrimarySmtpAddress   = 'forwarded@contoso.com'
                        },
                        [pscustomobject]@{
                            DisplayName          = 'Shared Projects'
                            RecipientTypeDetails = 'SharedMailbox'
                            PrimarySmtpAddress   = 'shared@contoso.com'
                        },
                        [pscustomobject]@{
                            DisplayName          = 'Executive Team'
                            RecipientTypeDetails = 'GroupMailbox'
                            PrimarySmtpAddress   = 'executive@contoso.com'
                        }
                    )
                    AllMailboxes = @{
                        '001-forwarded-mailbox' = [pscustomobject]@{
                            DisplayName             = 'Forwarded Mailbox'
                            DeliverToMailboxAndForward = $true
                            ForwardingSmtpAddress   = 'external@example.com'
                            RecipientTypeDetails    = 'UserMailbox'
                            PrimarySmtpAddress      = 'forwarded@contoso.com'
                        }
                    }
                    PrimaryMailboxStats = @(
                        [pscustomobject]@{
                            MailboxType        = 'UserMailbox'
                            TotalItemSizeBytes = 50GB
                        },
                        [pscustomobject]@{
                            MailboxType        = 'SharedMailbox'
                            TotalItemSizeBytes = 10GB
                        },
                        [pscustomobject]@{
                            MailboxType        = 'GroupMailbox'
                            TotalItemSizeBytes = 5GB
                        }
                    )
                    ArchiveMailboxes = @(
                        [pscustomobject]@{
                            RecipientTypeDetails = 'UserMailbox'
                        }
                    )
                    ArchiveMailboxStats = @(
                        [pscustomobject]@{
                            MailboxType        = 'UserMailbox'
                            TotalItemSizeBytes = 25GB
                        }
                    )
                    EmailActivityTopSenders = @(
                        [pscustomobject]@{
                            DisplayName       = 'Forwarded Mailbox'
                            UserPrincipalName = 'forwarded@contoso.com'
                            SendCount         = 245
                            LastActivityDate  = '2026-04-25'
                        },
                        [pscustomobject]@{
                            DisplayName       = 'Shared Projects'
                            UserPrincipalName = 'shared@contoso.com'
                            SendCount         = 120
                            LastActivityDate  = '2026-04-24'
                        }
                    )
                    EmailActivityTopReceivers = @(
                        [pscustomobject]@{
                            DisplayName       = 'Executive Team'
                            UserPrincipalName = 'executive@contoso.com'
                            ReceiveCount      = 310
                            LastActivityDate  = '2026-04-25'
                        },
                        [pscustomobject]@{
                            DisplayName       = 'Forwarded Mailbox'
                            UserPrincipalName = 'forwarded@contoso.com'
                            ReceiveCount      = 210
                            LastActivityDate  = '2026-04-23'
                        }
                    )
                    MailFlowConnectors = @(
                        [pscustomobject]@{
                            Enabled                = $true
                            RequireTls             = $false
                            TreatMessagesAsInternal = $true
                            SenderIPAddresses      = @('10.0.0.1')
                        }
                    )
                    PublicFolderDetails = @(
                        [pscustomobject]@{ Identity = '\Root\Legacy' }
                    )
                    SharedMailboxGovernanceSummary = @{
                        Summary = [pscustomobject]@{
                            OversizedSharedMailboxes         = 0
                            SharedMailboxesWithoutOwnerSignal = 12
                        }
                    }
                    InboxRulesExternalForwarding = @{
                        '001-forward-rule' = [pscustomobject]@{
                            Mailbox            = 'forwarded@contoso.com'
                            RuleName           = 'Forward Externally'
                            Enabled            = $true
                            ExternalTargets    = 'external@example.com'
                            ForwardTargetCount = 1
                        }
                    }
                    InboxRuleForwardingSummary = @{
                        Summary = [pscustomobject]@{
                            InspectedMailboxCount       = 1
                            ExternalForwardingRuleCount = 1
                            CollectionState             = 'Collected'
                        }
                    }
                    ForwardingPolicySummary = @{
                        Summary = [pscustomobject]@{
                            PolicyCollectionState                    = 'Collected'
                            HostedOutboundPolicyCount                = 2
                            HostedOutboundRuleCount                  = 1
                            PoliciesExplicitlyAllowingAutoForwarding = 1
                            PoliciesRestrictingAutoForwarding        = 1
                            PolicyAutoForwardingModes                = 'Default=Off; Pilot Allow External=On'
                            PoliciesReferencedByRules                = 'Pilot Allow External'
                            RemoteDomainCount                        = 2
                            RemoteDomainsAllowingAutoForwarding      = 1
                            RemoteDomainsAllowingAutoForwardingList  = 'partner.example'
                            DefaultRemoteDomainAllowsAutoForwarding  = $false
                        }
                    }
                }
                Collaboration = @{
                    AllTeams = @(
                        [pscustomobject]@{
                            OwnerCount          = 0
                            PrivateChannelCount = 6
                            SharedChannelCount  = 0
                        }
                    )
                    UnifiedGroups = @(
                        [pscustomobject]@{ OwnerCount = 0 }
                    )
                    SharePoint = @(
                        [pscustomobject]@{
                            Title                   = 'Projects Hub'
                            IsTeamsConnected        = $true
                            LastContentModifiedDate = (Get-Date).AddDays(-220).ToString('o')
                            StorageUsedGB           = 1200
                        },
                        [pscustomobject]@{
                            Title                   = 'Standalone PMO'
                            IsTeamsConnected        = $false
                            LastContentModifiedDate = (Get-Date).AddDays(-90).ToString('o')
                            StorageUsedGB           = 300
                        }
                    )
                    OneDrive = @(
                        [pscustomobject]@{
                            LastContentModifiedDate = (Get-Date).AddDays(-221).ToString('o')
                            StorageUsedGB           = 10
                        }
                    )
                    SharePointSharingSummary = @{
                        Summary = [pscustomobject]@{
                            CollectionSource                        = 'Microsoft Graph SharePoint tenant settings'
                            SettingsApiVersion                      = 'v1.0'
                            TenantSharingCapability                 = 'ExternalUserAndGuestSharing'
                            OneDriveSharingCapability               = 'ExternalUserSharingOnly'
                            DefaultSharingLinkType                  = 'AnonymousAccess'
                            DeletedUserPersonalSiteRetentionPeriodInDays = 30
                            OneDriveStorageQuotaMB                  = 1048576
                            SiteCreationDefaultStorageQuotaMB       = 26214400
                            IsSitesStorageLimitAutomatic            = $true
                            IsLegacyAuthProtocolsEnabled            = $true
                            DisableCustomAppAuthentication          = $false
                            IsUnmanagedSyncAppForTenantRestricted   = $true
                            IsSyncButtonHiddenOnPersonalSite        = $false
                            IsSiteCreationEnabled                   = $true
                            IsSiteCreationUIEnabled                 = $true
                            IsLoopEnabled                           = $true
                        }
                    }
                    ExternalSharingSummary = @{
                        Summary = [pscustomobject]@{
                            SharingDomainRestrictionMode = 'allowList'
                            SharingAllowedDomainList     = 'contoso.com'
                            SharingBlockedDomainList     = 'fabrikam.com'
                            AnonymousLinkExpirationInDays = 5
                            GuestExpirationInDays        = 30
                            RequireInvitedUserMatch      = $true
                            PreventExternalUsersFromResharing = $true
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
                            Workload         = 'OneDrive'
                            AssetType        = 'Personal Site'
                            Title            = 'Jefferson Test4'
                            UrlOrIdentifier  = 'https://contoso-my.sharepoint.com/personal/jtest4'
                            ExposureCategory = 'Externally sharable OneDrive with ownership mismatch'
                            TenantBaseline   = 'SharingCapability=ExternalUserAndGuestSharing; DefaultSharingLinkType=AnonymousAccess'
                            ObservedSetting  = 'SharingCapability=ExistingExternalUserSharingOnly; DefaultSharingLinkType=SpecificPeople'
                            OwnerSignal      = 'Owner=jtest4@contoso.com; ownership mismatch surfaced'
                            GuestSignal      = 'Not applicable'
                            ActivitySignal   = 'LastContentModifiedDate=2025-01-01'
                            StaleSignal      = 'Yes'
                            GapReason        = 'Externally sharable OneDrive also appears in the ownership-mismatch review.'
                            ReviewPriority   = 'High'
                        },
                        [pscustomobject]@{
                            Workload         = 'Teams'
                            AssetType        = 'Team'
                            Title            = 'Ownerless Team'
                            UrlOrIdentifier  = 'team-001'
                            ExposureCategory = 'Guest-heavy dormant Team'
                            TenantBaseline   = 'GuestInvitationControl=adminsAndGuestInviters; ConditionalAccessGuestCoverage=True'
                            ObservedSetting  = 'Guests=5; Members=10; GuestRatio=50%'
                            OwnerSignal      = '0 owner(s)'
                            GuestSignal      = '5 guest(s)'
                            ActivitySignal   = 'LastActivityDate=2025-01-01'
                            StaleSignal      = 'Yes'
                            GapReason        = 'Team is guest-heavy and its recorded activity is older than the 90-day dormant threshold.'
                            ReviewPriority   = 'High'
                        }
                    )
                }
                Security = @{
                    SecuritySecureScore = @(
                        [pscustomobject]@{
                            CreatedDateTime = (Get-Date).ToString('o')
                            CurrentScore    = 35
                            MaxScore        = 100
                        }
                    )
                    SMTPRelaySummary = [pscustomobject]@{
                        SMTPAuthEnabled                  = $true
                        SMTPAuthUsers                    = 2
                        SmtpClientAuthenticationDisabled = $false
                    }
                }
                Tenant = @{
                    Domains = @(
                        [pscustomobject]@{
                            Domain                = 'contoso.com'
                            Id                    = 'contoso.com'
                            IsVerified            = $false
                            Verified              = $false
                            TotalDomainRecipients = 3
                            PrimarySMTPRecipients = 2
                            AliasOnlyRecipients   = 1
                            DmarcConfigured       = $false
                            DkimConfigured        = $true
                        }
                    )
                }
            } `
            -Derived @{
                BestPracticeFindings = @{
                    '001-Secure Score' = [pscustomobject]@{
                        Area              = 'Secure Score'
                        Severity          = 'Risk'
                        Category          = 'Security Configuration'
                        Message           = 'Microsoft Secure Score is below the target range.'
                        Priority          = 'High'
                        RecommendedAction = 'Use the Secure Score improvement action plan to raise the baseline.'
                        RelatedWorksheet  = 'SecuritySecureScore'
                        RelatedSection    = 'Secure Score'
                    }
                    '002-Privileged Access' = [pscustomobject]@{
                        Area              = 'Identity & Admins'
                        Severity          = 'Risk'
                        Category          = 'Global Admin Count'
                        Message           = 'Global Administrator count is higher than the recommended operating threshold.'
                        Priority          = 'High'
                        RecommendedAction = 'Reduce standing Global Administrators and document emergency access handling.'
                        RelatedWorksheet  = 'Admins'
                        RelatedSection    = 'Identity & Admins'
                    }
                }
                BestPractices = @{
                    '001-Collaboration Summary' = [pscustomobject]@{
                        Area             = 'Ownership & Stewardship'
                        Severity         = 'Critical'
                        Status           = 'Critical'
                        Message          = '1 critical, 2 warning, 0 informational finding(s). Top signals: OneDrive Ownership Mismatch: 1; Owner Health: 1; Unowned Objects: 1.'
                        TotalFindings    = 3
                        CriticalFindings = 1
                        WarningFindings  = 2
                        InfoFindings     = 0
                        RelatedWorksheet = 'BestPractices'
                    }
                }
                AuthenticationConfigSummary = @{
                    Summary = [pscustomobject]@{
                        PasswordlessMethods         = @()
                        AdminConsentWorkflowEnabled = $false
                        PermissionGrantPolicies     = @('ManagePermissionGrantsForSelf.microsoft-user-default-allow-consent-apps')
                    }
                }
                OwnershipGovernanceSummary = @{
                    Summary = [pscustomobject]@{
                        MissingOwnerCount          = 2
                        OneDriveOwnerMismatchCount = 1
                    }
                }
                OneDriveOwnerMismatches = @(
                    [pscustomobject]@{ SiteUrl = 'https://contoso-my.sharepoint.com/personal/user' }
                )
                UnmanagedObjects = @(
                    [pscustomobject]@{ DisplayName = 'Ownerless Team'; OwnerState = 'Missing' }
                )
            } `
            -Diagnostics @{
                WarningCount = 1
                ErrorCount   = 0
            }

        $snapshotPath = Join-Path $TestDrive 'assessment.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $result = & $script:improveScriptPath -AssessmentJsonPath $snapshotPath -OutputFolder $TestDrive -PassThru

        Test-Path $result.JsonPath | Should -BeTrue
        Split-Path -Path $result.JsonPath -Parent | Should -Be (Join-Path $TestDrive 'Support')
        [System.IO.Path]::GetFileName($result.JsonPath) | Should -Match '.+-Plan\.json$'
        $result.CsvPath | Should -BeNullOrEmpty
        $result.MarkdownPath | Should -BeNullOrEmpty
        Test-Path $result.CustomerAssessmentReportPath | Should -BeTrue
        [System.IO.Path]::GetFileName($result.CustomerAssessmentReportPath) | Should -Match '.+-Microsoft 365 Tenant Best Practices Assessment-\d{4}-\d{2}-\d{2}\.docx$'
        Test-Path $result.CustomerAssessmentReportMarkdownPath | Should -BeTrue
        [System.IO.Path]::GetFileName($result.CustomerAssessmentReportMarkdownPath) | Should -Match '.+-CustRpt\.md$'
        if ($result.PSObject.Properties.Name -contains 'CustomerRemediationReportPath') {
            $result.CustomerRemediationReportPath | Should -BeNullOrEmpty
        }
        if ($result.PSObject.Properties.Name -contains 'CustomerRemediationReportMarkdownPath') {
            $result.CustomerRemediationReportMarkdownPath | Should -BeNullOrEmpty
        }
        Test-Path $result.EngineerActionPackPath | Should -BeTrue
        [System.IO.Path]::GetFileName($result.EngineerActionPackPath) | Should -Match '.+-EngPack\.md$'
        Test-Path $result.RemediationPs1Path | Should -BeTrue
        [System.IO.Path]::GetFileName($result.RemediationPs1Path) | Should -Match '.+-Snips\.ps1$'
        Split-Path -Path $result.RemediationPs1Path -Parent | Should -Be (Join-Path $TestDrive 'Support')
        $result.SupportFolderPath | Should -Be (Join-Path $TestDrive 'Support')

        $payload = Get-Content -Raw $result.JsonPath | ConvertFrom-Json -Depth 20
        $roadmapActionTitles = @($payload.RoadmapActions | ForEach-Object { [string]$_.ActionTitle } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $payload.Findings.Count | Should -BeGreaterThan 0
        $roadmapActionTitles.Count | Should -Be (@($roadmapActionTitles | Select-Object -Unique).Count)
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'Source'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'PriorityBand'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'OwnerTeam'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'RoadmapPhase'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'CustomerSummary'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'EngineerNotes'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'WhyFlagged'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'ExampleAction'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'EvidenceLocation'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'TechnicalRemediation'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'RelatedWorksheet'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'RelatedSection'
        $payload.PSObject.Properties.Name | Should -Contain 'WorkstreamSummaries'
        $payload.PSObject.Properties.Name | Should -Contain 'ExternalExposureFindings'
        $payload.PSObject.Properties.Name | Should -Contain 'ConsultativeSummaries'
        @($payload.WorkstreamSummaries).Count | Should -BeGreaterThan 0
        @($payload.ExternalExposureFindings).Count | Should -Be 2
        $payload.ConsultativeSummaries.PSObject.Properties.Name | Should -Contain 'ExecutiveDecisionSummary'
        $payload.ConsultativeSummaries.PSObject.Properties.Name | Should -Contain 'MessagingConsultativeSummary'
        $payload.ConsultativeSummaries.PSObject.Properties.Name | Should -Contain 'IdentityConsultativeSummary'
        $payload.ConsultativeSummaries.PSObject.Properties.Name | Should -Contain 'CollaborationConsultativeSummary'
        $payload.ConsultativeSummaries.PSObject.Properties.Name | Should -Contain 'GovernanceConsultativeSummary'
        $payload.ConsultativeSummaries.PSObject.Properties.Name | Should -Contain 'LifecycleConsultativeSummary'
        @($payload.ConsultativeSummaries.ExecutiveDecisionSummary.RiskRows).Count | Should -BeGreaterThan 0
        @($payload.ConsultativeSummaries.MessagingConsultativeSummary.SnapshotRows).Count | Should -BeGreaterThan 0
        @($payload.ConsultativeSummaries.IdentityConsultativeSummary.SnapshotRows).Count | Should -BeGreaterThan 0
        @($payload.ConsultativeSummaries.CollaborationConsultativeSummary.SnapshotRows).Count | Should -BeGreaterThan 0

        @($payload.Findings | Where-Object { $_.RuleId -eq 'SEC-001' }).Count | Should -Be 0
        @($payload.Findings | Where-Object { $_.RuleId -like 'AREA-*' }).Count | Should -Be 0
        @($payload.Findings | Where-Object { $_.Source -eq 'Derived/BestPracticeFindings' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-007' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-009' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-010' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-011' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-012' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-005' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'CA-002' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'CA-010' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'CA-012' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'EX-001' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'EX-006' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'EX-007' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'DEV-005' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'DEV-006' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'COL-005' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'COL-006' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.TargetValue -eq 'Reduce open findings in this workstream' }).Count | Should -Be 0
        @($payload.Findings | Where-Object { $_.TargetValue -match 'unauthorized mailbox forwarding|unauthorized inbox-rule forwarding' }).Count | Should -Be 0
        @($payload.Findings | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.ExampleAction) }).Count | Should -Be 0

        $id007 = @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-007' }) | Select-Object -First 1
        $id007.WhyFlagged | Should -Match 'permission'
        $id007.EvidenceLocation | Should -Match 'EnterpriseApplications'
        $id007.TechnicalRemediation | Should -Match 'business owner|permissions|broad consent'

        $id010 = @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-010' }) | Select-Object -First 1
        $id010.CurrentValue | Should -Match 'owner signal'

        $id011 = @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-011' }) | Select-Object -First 1
        $id011.CurrentValue | Should -Match 'redirect URI'

        $col002 = @($payload.Findings | Where-Object { $_.RuleId -eq 'COL-002' }) | Select-Object -First 1
        $col002.Finding | Should -Be 'OneDrive delegated ownership review is required.'

        $ex001 = @($payload.Findings | Where-Object { $_.RuleId -eq 'EX-001' }) | Select-Object -First 1
        $ex001.CurrentValue | Should -Match '^1 mailbox\(es\) with forwarding configured'
        $ex001.CurrentValue | Should -Match '1 hosted outbound policy/policies explicitly allow auto-forwarding'
        $ex001.CurrentValue | Should -Match '1 remote domain\(s\) have AutoForwardEnabled'
        $ex001.CurrentValue | Should -Match 'Default remote domain AutoForwardEnabled=False'
        $ex001.TargetValue | Should -Be 'All mailbox forwarding configurations reviewed and either approved or removed, with tenant forwarding policy aligned to the approved baseline'

        $ex006 = @($payload.Findings | Where-Object { $_.RuleId -eq 'EX-006' }) | Select-Object -First 1
        $ex006.CurrentValue | Should -Match 'external-forwarding inbox rule\(s\)'
        $ex006.CurrentValue | Should -Match 'Observed policy modes: Default=Off; Pilot Allow External=On'
        $ex006.TargetValue | Should -Be 'All external inbox-rule forwarding paths reviewed and either approved or removed, with tenant forwarding policy aligned to the approved baseline'

        $customerReportDocument = Get-TestDocxDocumentXmlText -Path $result.CustomerAssessmentReportPath
        $customerReportXml = Get-TestDocxDocumentXml -Path $result.CustomerAssessmentReportPath
        $customerReportEntryNames = Get-TestDocxEntryNames -Path $result.CustomerAssessmentReportPath
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath
        $ns = New-Object System.Xml.XmlNamespaceManager($customerReportXml.NameTable)
        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $customerTables = $customerReportXml.SelectNodes('//w:tbl', $ns)
        $customerReportDocument | Should -Match '1\.0 Introduction'
        $customerReportDocument | Should -Match 'Assessment Snapshot At A Glance'
        $customerReportDocument | Should -Match '2\.0 Project Scope'
        $customerReportDocument | Should -Match '3\.0 Executive Summary'
        $customerReportDocument | Should -Match 'Overall Findings Summary'
        $customerReportDocument | Should -Match 'Risk Clusters'
        $customerReportDocument | Should -Match '4\.0 Modern Workplace Recommendations'
        $customerReportDocument | Should -Match '5\.0 Entra ID Review: User and Device Inventory'
        $customerReportDocument | Should -Match '5\.1 Entra User'
        $customerReportDocument | Should -Match '5\.2 Entra Device'
        $customerReportDocument | Should -Match 'Device Registration and Compliance Gaps'
        $customerReportDocument | Should -Match 'Device Platform Distribution'
        $customerReportDocument | Should -Match '5\.3 Entra Guest Access Configuration'
        $customerReportDocument | Should -Match '5\.4 Entra Applications and Access Review'
        $customerReportDocument | Should -Match 'Application Spotlight'
        $customerReportDocument | Should -Match 'SSO-Enabled Applications'
        $customerReportDocument | Should -Match 'Inactive or High-Privilege Applications'
        $customerReportDocument | Should -Match 'Credential Cleanup Opportunities'
        $customerReportDocument | Should -Match '6\.0 Authentication Methods, MFA Enrollment, and MFA Enforcement'
        $customerReportDocument | Should -Match 'MFA Enrollment'
        $customerReportDocument | Should -Match 'MFA Enrollment Status'
        $customerReportDocument | Should -Match 'Registered MFA Method Mix'
        $customerReportDocument | Should -Match 'MFA Enforcement'
        $customerReportDocument | Should -Match 'Covered vs Not Covered by Active MFA Enforcement'
        $customerReportDocument | Should -Match 'MFA Enforcement Driver Breakdown'
        $customerReportDocument | Should -Match 'Software one-time passcode'
        $customerReportDocument | Should -Match 'Users with weak MFA methods only'
        $customerReportDocument | Should -Match 'Users with weak default MFA method'
        $customerReportDocument | Should -Match 'Enabled MFA enforcement policies'
        $customerReportDocument | Should -Match 'Report-only MFA enforcement policies'
        $customerReportDocument | Should -Match 'Guest-user MFA enforcement summary'
        $customerReportDocument | Should -Match 'Admin users not registered for MFA'
        $customerReportDocument | Should -Match 'Admin users not covered by enabled MFA enforcement policies'
        $customerReportDocument | Should -Match 'Users not covered by enabled MFA enforcement policies'
        $customerReportDocument | Should -Match 'MFA enforcement state'
        $customerReportDocument | Should -Match 'MFA Enforcement Gap Summary'
        $customerReportDocument | Should -Match 'Common Coverage Drivers'
        $customerReportDocument | Should -Match 'Internal Member Users Outside Active MFA Include Scope'
        $customerReportDocument | Should -Match 'Internal Member Users Explicitly Excluded from Active MFA Policies'
        $customerReportDocument | Should -Match 'Guest MFA Coverage Drivers'
        $customerReportDocument | Should -Match 'Representative MFA Scope Examples'
        $customerReportDocument | Should -Match 'Coverage Gap Signal'
        $customerReportDocument | Should -Match 'Uncovered Member'
        $customerReportDocument | Should -Match 'Excluded Member'
        $customerReportDocument | Should -Not -Match 'Uncovered Guest'
        $customerReportDocument | Should -Match 'Outside include scope'
        $customerReportDocument | Should -Match 'outside the active MFA include scope'
        $customerReportDocument | Should -Match 'Enrollment and enforcement are intentionally reported as separate views'
        $customerReportDocument | Should -Match 'Some Conditional Access policies naturally target employee, admin, or workload-specific populations'
        $customerReportDocument | Should -Match 'partner tenants with the most collaboration and a validated trust relationship'
        $customerReportDocument | Should -Match 'MfaEnforcementGapUsers'
        $customerReportDocument | Should -Match 'MfaEnforcementScopeReview'
        $customerReportDocument | Should -Match '7\.0 Password Writeback and Self-Service Password Reset'
        $customerReportDocument | Should -Match '8\.0 Authorization: Admin Access and Role Assignments'
        $customerReportDocument | Should -Match 'Recent vs Stale Admins'
        $customerReportDocument | Should -Match 'Admins Covered vs Not Covered by Active MFA Enforcement'
        $customerReportDocument | Should -Match 'Service Accounts'
        $customerReportDocument | Should -Match '9\.0 Exchange Online: Mailboxes and Storage Overview'
        $customerReportDocument | Should -Match 'Data Footprint by Workload'
        $customerReportDocument | Should -Match '9\.1 Recipient and Mailbox Footprint'
        $customerReportDocument | Should -Match 'Recipient Mix by Type'
        $customerReportDocument | Should -Match 'Top Recipient Domains'
        $customerReportDocument | Should -Match 'Top Senders'
        $customerReportDocument | Should -Match 'Top Receivers'
        $customerReportDocument | Should -Match 'Mailbox Lifecycle Summary'
        $customerReportDocument | Should -Match 'Transport Exposure Summary'
        $customerReportDocument | Should -Match 'User Mailbox Growth'
        $customerReportDocument | Should -Match 'Shared Mailbox Review'
        $customerReportDocument | Should -Match 'Inactive Mailboxes'
        $customerReportDocument | Should -Match 'Group Mailbox Utilization'
        $customerReportDocument | Should -Match 'Archive Mailbox Usage and Licensing'
        $customerReportDocument | Should -Match '9\.2 SMTP Relay Usage'
        $customerReportDocument | Should -Match '10\. Microsoft Teams Governance and Cleanup'
        $customerReportDocument | Should -Match 'Key Findings'
        $customerReportDocument | Should -Match 'Opportunities for Cleanup'
        $customerReportDocument | Should -Match '11\.0 SharePoint Online Storage and External Sharing'
        $customerReportDocument | Should -Match 'Largest Collaboration Sites to Review'
        $customerReportDocument | Should -Match '12\.0 Retention Policies and Data Loss Prevention'
        $customerReportDocument | Should -Match '13\.0 Domain Configuration and DNS Overview'
        $customerReportDocument | Should -Match '14\.0 Offboarding Recommendation'
        $customerReportDocument | Should -Match '15\.0 Appendix'
        $customerReportDocument | Should -Match '15\.10 Full Findings Inventory'
        $customerReportDocument | Should -Match 'Application User Consent Management'
        $customerReportDocument | Should -Match 'Authentication Methods Migration'
        $customerReportDocument | Should -Match 'DMARC Records'
        $customerReportDocument | Should -Match 'Password Writeback with AD Sync'
        $customerReportDocument | Should -Match 'Version History'
        $customerReportDocument | Should -Match 'Arraya Solutions'
        $customerReportDocument | Should -Match 'Not validated from the reviewed data'
        $customerReportDocument | Should -Match 'Configuration Signal'
        $customerReportDocument | Should -Match 'Current State'
        $customerReportDocument | Should -Match 'Leadership Decision Brief'
        $customerReportDocument | Should -Match 'Guest invitation control'
        $customerReportDocument | Should -Match 'Cross-tenant partner count'
        $customerReportDocument | Should -Match 'Default inbound MFA trust'
        $customerReportDocument | Should -Match 'External Access Snapshot'
        $customerReportDocument | Should -Match 'Observed Guest and B2B Posture'
        $customerReportDocument | Should -Match 'Recommended Next Step'
        $customerReportDocument | Should -Match 'Tenant External Sharing Snapshot'
        $customerReportDocument | Should -Match 'SharePoint Tenant Controls Snapshot'
        $customerReportDocument | Should -Match 'External Exposure Review'
        $customerReportDocument | Should -Match 'Exposure Category'
        $customerReportDocument | Should -Match 'Review Priority'
        $customerReportDocument | Should -Match 'Guest-heavy dormant Team'
        $customerReportDocument | Should -Match 'Admins and approved guest inviters can invite guests'
        $customerReportDocument | Should -Match 'Only explicitly allowed domains may be shared externally'
        $customerReportDocument | Should -Match 'Prevent external users from resharing'
        $customerReportDocument | Should -Match 'Legacy auth protocols enabled'
        $customerReportDocument | Should -Match 'Deleted user personal site retention \(days\)'
        $customerReportDocument | Should -Match 'Why it is relevant:'
        $customerReportDocument | Should -Match 'Recommendation'
        $customerReportDocument | Should -Match 'Criticality'
        $customerReportDocument | Should -Match 'Level of Effort'
        $customerReportDocument | Should -Match 'initial delivery-planning estimate'
        $customerReportDocument | Should -Match 'w:fill="FCE5CD"'
        $customerReportDocument | Should -Match 'w:fill="FFF2CC"'
        $customerReportDocument | Should -Match 'Workstream'
        $customerReportDocument | Should -Match 'Severity / Impact'
        $customerReportDocument | Should -Match 'Open Findings'
        $customerReportDocument | Should -Match 'What Stands Out'
        $customerReportDocument | Should -Match 'Open Findings'
        $customerReportDocument | Should -Match 'Rule / Finding'
        $customerReportDocument | Should -Match 'Why Flagged'
        $customerReportDocument | Should -Match 'Success Criteria'
        $customerReportDocument | Should -Match 'DMARC'
        $customerReportDocument | Should -Not -Match 'The clearest concentration in this tenant appears in'
        $customerReportDocument | Should -Not -Match 'AREA-'
        $customerReportDocument | Should -Not -Match 'CustomerRemediationReport'
        $customerReportDocument | Should -Not -Match 'unauthorized mailbox forwarding'
        $customerReportDocument | Should -Not -Match 'unauthorized inbox-rule forwarding'
        $customerReportDocument | Should -Not -Match 'licensing headroom appear together in this section'
        $customerReportDocument | Should -Not -Match '10dae51f-b6af-4016-8d66-8c2a99b929b3'
        $customerReportDocument | Should -Not -Match 'Conditional Access - Key Recommendations'
        $customerReportDocument | Should -Not -Match 'Conditional Access - Current State Analysis'
        $customerReportDocument | Should -Not -Match 'Recommended Next Steps'
        $customerReportDocument | Should -Not -Match 'Domain Recommendations'
        $customerReportDocument | Should -Not -Match 'DNS Recommendations'
        $customerReportDocument | Should -Not -Match 'Enabled MFA Policy Scope Review'
        $customerReportDocument | Should -Not -Match 'Top Internal Member Users Not Covered by Enabled MFA Enforcement'
        $customerReportDocument | Should -Not -Match 'Findings Legend'
        $customerReportDocument | Should -Not -Match 'What Should Happen Next'
        $customerReportDocument | Should -Not -Match 'Priority / Impact'
        $customerReportDocument | Should -Not -Match 'First Step'
        $customerReportDocument | Should -Not -Match 'Success Check'
        $customerReportDocument | Should -Not -Match 'User Impact / Expected Experience'
        $customerReportDocument | Should -Not -Match 'SSO Mode'
        $customerReportDocument | Should -Not -Match 'Tier B'
        $customerReportDocument | Should -Not -Match 's of tw ar eO ne Ti me Pa ss co de'
        $customerReportDocument | Should -Not -Match 'Action:'
        $customerReportDocument | Should -Not -Match 'What It Means In This Report'
        $customerReportDocument | Should -Match 'w:tblBorders'
        $customerTables.Count | Should -BeGreaterThan 12
        $customerTables[0].SelectNodes('./w:tr[1]/w:tc', $ns).Count | Should -Be 5
        $customerTables[0].SelectNodes('./w:tr[2]/w:tc', $ns).Count | Should -Be 5
        @($customerTables | Where-Object { $_.SelectNodes('./w:tr[1]/w:tc', $ns).Count -eq 4 }).Count | Should -BeGreaterThan 0
        @($customerTables | Where-Object { $_.SelectNodes('./w:tr[1]/w:tc', $ns).Count -eq 3 }).Count | Should -BeGreaterThan 0
        $customerReportMarkdown | Should -Match '# .+ Microsoft 365 Tenant Best Practices Assessment'
        $customerReportMarkdown | Should -Match '## 1\.0 Introduction'
        $customerReportMarkdown | Should -Match '### Assessment Snapshot At A Glance'
        $customerReportMarkdown | Should -Match '## 3\.0 Executive Summary'
        $customerReportMarkdown | Should -Match '### Overall Findings Summary'
        $customerReportMarkdown | Should -Match '### Leadership Decision Brief'
        $customerReportMarkdown | Should -Match '## 4\.0 Modern Workplace Recommendations'
        $customerReportMarkdown | Should -Match '\| Recommendation \| Criticality \| Level of Effort \|'
        $customerReportMarkdown | Should -Match '\| .+ \| (Critical|High|Medium|Low) \| (High|Medium|Low) \|'
        $customerReportMarkdown | Should -Match 'Level of Effort is an initial delivery-planning estimate'
        $customerReportMarkdown | Should -Match '\| Workstream \| Severity / Impact \| Open Findings \| What Stands Out \|'
        $customerReportMarkdown | Should -Match '### Risk Clusters'
        $customerReportMarkdown | Should -Not -Match '### Findings Legend'
        $customerReportMarkdown | Should -Not -Match '- Open Findings:'
        $customerReportMarkdown | Should -Not -Match '- Severity / Impact:'
        $customerReportMarkdown | Should -Match '### 5\.3 Entra Guest Access Configuration'
        $customerReportMarkdown | Should -Match '#### External Access Snapshot'
        $customerReportMarkdown | Should -Match '\| Guest invitation control \|'
        $customerReportMarkdown | Should -Match '\| Cross-tenant partner count \|'
        $customerReportMarkdown | Should -Match '#### Observed Guest and B2B Posture'
        $customerReportMarkdown | Should -Match '#### Recommended Next Step'
        $customerReportMarkdown | Should -Match 'Admins and approved guest inviters can invite guests'
        $customerReportMarkdown | Should -Match 'Guest MFA enforcement matters because guest identities are external accounts with access into this tenant''s resources'
        $customerReportMarkdown | Should -Match 'home tenant instead of registering separately in this tenant'
        $customerReportMarkdown | Should -Not -Match 'What should happen next:'
        $customerReportMarkdown | Should -Match 'Detailed admin gap rows are available in the workbook tabs AdminMfaRegistrationGaps and AdminMfaEnforcementGaps'
        $customerReportMarkdown | Should -Match '### 5\.4 Entra Applications and Access Review'
        $customerReportMarkdown | Should -Match '### Application Spotlight'
        $customerReportMarkdown | Should -Match '- Applications: High Priv App'
        $customerReportMarkdown | Should -Match 'Primary risk signal:'
        $customerReportMarkdown | Should -Match 'Latest activity:'
        $customerReportMarkdown | Should -Not -Match '\| Application \| SSO Mode \| Primary Risk Signal \| Latest Activity \|'
        $customerReportMarkdown | Should -Match '### SSO-Enabled Applications'
        $customerReportMarkdown | Should -Match '### Inactive or High-Privilege Applications'
        $customerReportMarkdown | Should -Match '### Credential Cleanup Opportunities'
        $customerReportMarkdown | Should -Match '- High Priv App - SAML;'
        $customerReportMarkdown | Should -Match '1 expired client secret'
        $customerReportMarkdown | Should -Match 'Latest activity 2026-04-12'
        $customerReportMarkdown | Should -Match '## 6\.0 Authentication Methods, MFA Enrollment, and MFA Enforcement'
        $customerReportMarkdown | Should -Match '### MFA Enrollment'
        $customerReportMarkdown | Should -Match '#### MFA Enrollment Status'
        $customerReportMarkdown | Should -Match '#### Registered MFA Method Mix'
        $customerReportMarkdown | Should -Match '### MFA Enforcement'
        $customerReportMarkdown | Should -Match '#### Covered vs Not Covered by Active MFA Enforcement'
        $customerReportMarkdown | Should -Match '#### MFA Enforcement Driver Breakdown'
        $customerReportMarkdown | Should -Match 'Software one-time passcode'
        $customerReportMarkdown | Should -Match '\| Conditional Access policies reviewed \| 8 \|'
        $customerReportMarkdown | Should -Match '\| Policies with exclusions \| 7 \|'
        $customerReportMarkdown | Should -Match '\| Users covered by enabled MFA enforcement policies \| 60 \|'
        $customerReportMarkdown | Should -Match '\| Users not covered by enabled MFA enforcement policies \| 20 \|'
        $customerReportMarkdown | Should -Match '\| Estimated enabled-user CA MFA coverage \| 75% \|'
        $customerReportMarkdown | Should -Match '\| Estimated member-user CA MFA coverage \| 80% \|'
        $customerReportMarkdown | Should -Match '\| Estimated guest-user CA MFA coverage \| 40% \|'
        $customerReportMarkdown | Should -Match 'Guest-user MFA enforcement summary'
        $customerReportMarkdown | Should -Match '\| Admin users not registered for MFA \| 2 \|'
        $customerReportMarkdown | Should -Match '\| Admin users not covered by enabled MFA enforcement policies \| 1 \|'
        $customerReportMarkdown | Should -Match '#### MFA Enforcement Gap Summary'
        $customerReportMarkdown | Should -Match '\| Coverage Gap Signal \| Current State \|'
        $customerReportMarkdown | Should -Match '\| Users outside enabled MFA CA include scope \| 2 \|'
        $customerReportMarkdown | Should -Match '\| Users explicitly excluded from enabled MFA CA policies \| 1 \|'
        $customerReportMarkdown | Should -Match '#### Common Coverage Drivers'
        $customerReportMarkdown | Should -Match '- Outside include scope: 1 uncovered internal member user\(s\)\.'
        $customerReportMarkdown | Should -Match '- Policy: Baseline MFA: 1 uncovered internal member user\(s\)\.'
        $customerReportMarkdown | Should -Match '#### Internal Member Users Outside Active MFA Include Scope'
        $customerReportMarkdown | Should -Match '#### Internal Member Users Explicitly Excluded from Active MFA Policies'
        $customerReportMarkdown | Should -Match '\| Display Name \| User Principal Name \|'
        $customerReportMarkdown | Should -Match 'Uncovered Member'
        $customerReportMarkdown | Should -Match 'Excluded Member'
        $customerReportMarkdown | Should -Not -Match 'Uncovered Guest'
        $customerReportMarkdown | Should -Match 'Outside include scope'
        $customerReportMarkdown | Should -Match 'Enrollment and enforcement are intentionally reported as separate views in this report'
        $customerReportMarkdown | Should -Not -Match 'Related Policy / Scope'
        $customerReportMarkdown | Should -Match 'The desired baseline is to require strong guest authentication through a guest-specific Conditional Access policy and to trust the guest home-tenant MFA where supported and approved'
        $customerReportMarkdown | Should -Match '#### Guest MFA Coverage Drivers'
        $customerReportMarkdown | Should -Match '#### Representative MFA Scope Examples'
        $customerReportMarkdown | Should -Match 'Some Conditional Access policies naturally target employee, admin, or workload-specific populations'
        $customerReportMarkdown | Should -Match 'partner tenants with the most collaboration and a validated trust relationship'
        $customerReportMarkdown | Should -Match '\| Policy \| Scope Type \| Scope Signal \| Why It Matters \|'
        $customerReportMarkdown | Should -Match 'Break Glass Exclusions'
        $customerReportDocument | Should -Match 'Global Reader'
        $customerReportDocument | Should -Match 'eligible role that requires approval before activation'
        $customerReportMarkdown | Should -Match '\| Reduce privileged access and strengthen identity controls \| High \|'
        $customerReportMarkdown | Should -Match '\| Users with weak MFA methods only \|'
        $customerReportMarkdown | Should -Match '\| Enabled MFA enforcement policies \|'
        $customerReportMarkdown | Should -Match 'MfaEnforcementGapUsers'
        $customerReportMarkdown | Should -Match 'MfaEnforcementScopeReview'
        $customerReportMarkdown | Should -Match '## 11\.0 SharePoint Online Storage and External Sharing'
        $customerReportMarkdown | Should -Match '### Tenant External Sharing Snapshot'
        $customerReportMarkdown | Should -Match '### SharePoint Tenant Controls Snapshot'
        $customerReportMarkdown | Should -Match '### External Exposure by Category'
        $customerReportMarkdown | Should -Match '\| Prevent external users from resharing \| Prevented \|'
        $customerReportMarkdown | Should -Match '\| Legacy auth protocols enabled \| Enabled \|'
        $customerReportMarkdown | Should -Match '\| Deleted user personal site retention \(days\) \| 30 \|'
        $customerReportMarkdown | Should -Match '### External Exposure Review'
        $customerReportMarkdown | Should -Match '\| Workload \| Asset Type \| Title \| Exposure Category \| Gap Reason \| Review Priority \|'
        $customerReportMarkdown | Should -Match '## 9\.0 Exchange Online: Mailboxes and Storage Overview'
        $customerReportMarkdown | Should -Match '### Data Footprint by Workload'
        $customerReportMarkdown | Should -Match '### 9\.1 Recipient and Mailbox Footprint'
        $customerReportMarkdown | Should -Match '### Recipient Mix by Type'
        $customerReportMarkdown | Should -Match '### Top Recipient Domains'
        $customerReportMarkdown | Should -Match '\| Domain \| Total \| Primary \| Alias-only \|'
        $customerReportMarkdown | Should -Match 'Teams storage is represented here through team-connected SharePoint site storage'
        $customerReportMarkdown | Should -Match '#### Top Senders'
        $customerReportMarkdown | Should -Match '#### Top Receivers'
        $customerReportMarkdown | Should -Match '\| Display Name \| User Principal Name \| Send Count \| Last Activity \|'
        $customerReportMarkdown | Should -Match '\| Display Name \| User Principal Name \| Receive Count \| Last Activity \|'
        $customerReportMarkdown | Should -Match '### Largest Collaboration Sites to Review'
        $customerReportMarkdown | Should -Match '- '
        $customerReportMarkdown | Should -Match '## 12\.0 Retention Policies and Data Loss Prevention'
        $customerReportMarkdown | Should -Match '## 14\.0 Offboarding Recommendation'
        $customerReportMarkdown | Should -Match '### 15\.10 Full Findings Inventory'
        $customerReportMarkdown | Should -Match '#### Mailbox Lifecycle Summary'
        $customerReportMarkdown | Should -Match '#### Transport Exposure Summary'
        $customerReportMarkdown | Should -Match '\| Severity \| Priority \| Workstream \| Rule / Finding \| Why Flagged \| Recommended Action \| Success Criteria \|'
        $customerReportMarkdown | Should -Match 'Only explicitly allowed domains may be shared externally'
        $customerReportMarkdown | Should -Match 'Guest invitation control'
        $customerReportMarkdown | Should -Not -Match 'Priority: Near Term'
        $customerReportMarkdown | Should -Not -Match 'The clearest concentration in this tenant appears in'
        $customerReportMarkdown | Should -Not -Match 'CustomerRemediationReport'
        $customerReportMarkdown | Should -Not -Match '10dae51f-b6af-4016-8d66-8c2a99b929b3'
        $customerReportMarkdown | Should -Not -Match '#### Enabled MFA Policy Scope Review'
        $customerReportMarkdown | Should -Not -Match '#### Top Internal Member Users Not Covered by Enabled MFA Enforcement'
        $customerReportMarkdown | Should -Not -Match '\| Display Name \| User Principal Name \| Gap Category \| Coverage Driver \|'
        $customerReportMarkdown | Should -Not -Match 'Tier B'
        $customerReportMarkdown | Should -Not -Match 's of tw ar eO ne Ti me Pa ss co de'
        $customerReportMarkdown | Should -Not -Match 'Primary Owner'
        $customerReportMarkdown | Should -Not -Match 'Action:'
        $customerReportMarkdown | Should -Not -Match 'Recommended Next Steps'
        $customerReportMarkdown | Should -Not -Match 'must register separately in this tenant'
        $customerReportMarkdown | Should -Not -Match '\| Term \| What It Means In This Report \|'
        $customerReportMarkdown | Should -Not -Match '\| Application \| SSO Enabled \| SSO Mode \| Observation \|'
        $customerReportMarkdown | Should -Not -Match '\| Application \| SSO Mode \| Privilege / Activity \| Observation \|'
        $customerReportMarkdown | Should -Not -Match '\| Application \| Cleanup Signal \| Latest Activity \| Observation \|'
        $customerReportMarkdown | Should -Not -Match '\| Application \| Credential State \| Latest Activity \| Observation \|'
        $customerReportMarkdown | Should -Not -Match '#### Messaging Snapshot At A Glance'

        $engineerPack = Get-Content -Raw $result.EngineerActionPackPath
        $engineerPack | Should -Match '## Engineering Summary'
        $engineerPack | Should -Match '## Findings To Work'
        $engineerPack | Should -Match '## Supporting Files'
        $engineerPack | Should -Match 'Why It Matters'
        $engineerPack | Should -Match 'Where To Verify'
        $engineerPack | Should -Match 'Technical Remediation'
        $engineerPack | Should -Match 'What Needs Attention'
        $engineerPack | Should -Match 'Done When'
        $engineerPack | Should -Not -Match '## Engineer Notes'
        $engineerPack | Should -Not -Match 'Action Path'
        $engineerPack | Should -Not -Match 'Normalized Findings'
        $engineerPack | Should -Not -Match 'Command References'
        $engineerPack | Should -Not -Match '\| Example \|'
        $engineerPack | Should -Not -Match 'CSV output:'
        $engineerPack | Should -Match 'Support folder:'

        $snippetContent = Get-Content -Raw $result.RemediationPs1Path
        $snippetContent | Should -Match '\$gaRole = Get-MgDirectoryRole'
        $snippetContent | Should -Match "Connect-MgGraph -Scopes 'SecurityEvents.Read.All'"
        $snippetContent | Should -Match "Connect-MgGraph -Scopes 'Policy.Read.All'"
        $snippetContent | Should -Not -Match 'SecurityEvents.ReadWrite.All'
        $snippetContent | Should -Not -Match 'Policy.ReadWrite.ConditionalAccess'
    }

    It 'embeds customer-report chart images in the DOCX and skips them in markdown when chart datasets are renderable' {
        $snapshot = New-ArrayaTenantSnapshot -Data @{
            Identity = @{
                Admins = @(
                    [pscustomobject]@{
                        DisplayName       = 'Recent Admin'
                        UserPrincipalName = 'recent.admin@contoso.com'
                        AccountEnabled    = $true
                        LastSignInDateTime = (Get-Date).AddDays(-20).ToString('o')
                    },
                    [pscustomobject]@{
                        DisplayName       = 'Stale Admin'
                        UserPrincipalName = 'stale.admin@contoso.com'
                        AccountEnabled    = $true
                        LastSignInDateTime = (Get-Date).AddDays(-240).ToString('o')
                    }
                )
                Users = @(
                    [pscustomobject]@{
                        UserPrincipalName = 'registered.user@contoso.com'
                        UserType          = 'Member'
                        AccountEnabled    = $true
                    },
                    [pscustomobject]@{
                        UserPrincipalName = 'not.registered@contoso.com'
                        UserType          = 'Member'
                        AccountEnabled    = $true
                    }
                )
                ConditionalAccessPolicies = @()
                ConditionalAccessPolicySummary = @{
                    Summary = [pscustomobject]@{
                        TotalPolicies                 = 0
                        EnabledPolicies               = 0
                        ReportOnlyPolicies            = 0
                        HasGuestCoverage              = $false
                        HasPrivilegedRoleCoverage     = $false
                        HasCompliantDeviceRequirement = $false
                        HasRiskBasedCoverage          = $false
                        PoliciesWithExclusions        = 0
                    }
                }
                SecurityDefaultsPolicy = @{
                    Summary = [pscustomobject]@{
                        IsEnabled = $false
                    }
                }
                AuthenticationConfig = [pscustomobject]@{
                    AdminConsentWorkflowEnabled    = $false
                    PermissionGrantPoliciesAssigned = @()
                    PasswordlessMethods            = @()
                }
                MfaRegistrationSummary = [pscustomobject]@{
                    TotalUsers         = 5
                    RegisteredUsers    = 3
                    NotRegisteredUsers = 2
                    MethodCounts       = @{ 'Microsoft Authenticator' = 2; 'SMS / phone' = 1 }
                }
                MfaEnrollmentSummary = [pscustomobject]@{
                    TotalUsers                        = 5
                    RegisteredUsers                   = 3
                    NotRegisteredUsers                = 2
                    RegistrationPercent               = 60
                    RegisteredMethodBreakdown         = 'Microsoft Authenticator=2; SMS / phone=1'
                    WeakMethodBreakdown               = 'SMS / phone=1'
                    PhishingResistantMethodBreakdown  = 'FIDO2 security key / passkey=1'
                    UsersWithWeakMethodsOnly          = 1
                    UsersWithWeakDefaultMethod        = 1
                    UsersWithPhishingResistantMethods = 1
                }
                MfaEnforcementSummary = [pscustomobject]@{
                    ConditionalAccessPoliciesReviewed      = 2
                    EnabledPoliciesRequiringMfa            = 1
                    ReportOnlyPoliciesRequiringMfa         = 0
                    PoliciesWithExclusions                 = 1
                    EnabledUsersReviewed                   = 5
                    UsersCoveredByEnabledMfaPolicies       = 3
                    UsersNotCoveredByEnabledMfaPolicies    = 2
                    UserCoveragePercent                    = 60
                    EnabledMemberUsersReviewed             = 4
                    MemberUsersCoveredByEnabledMfaPolicies = 2
                    MemberUserCoveragePercent              = 50
                    EnabledGuestUsersReviewed              = 1
                    GuestUsersCoveredByEnabledMfaPolicies  = 1
                    GuestUserCoveragePercent               = 100
                    GuestUserEnforcementState              = 'Guest users appear covered in the reviewed baseline.'
                    EnforcementState                       = 'MFA enforcement is active through enabled Conditional Access policies.'
                }
                MfaEnforcementGapUsers = @(
                    [pscustomobject]@{
                        DisplayName       = 'Outside Scope User'
                        UserPrincipalName = 'outside.scope@contoso.com'
                        UserType          = 'Member'
                        GapCategory       = 'Outside enabled MFA CA include scope'
                        GapReason         = 'User is outside the include scope of all enabled Conditional Access policies that currently require MFA.'
                        RelatedPolicies   = ''
                    },
                    [pscustomobject]@{
                        DisplayName       = 'Excluded Scope User'
                        UserPrincipalName = 'excluded.scope@contoso.com'
                        UserType          = 'Member'
                        GapCategory       = 'Excluded from all enabled MFA CA policies that otherwise target the user'
                        GapReason         = 'Excluded through group ''Break Glass Exclusions'' in enabled MFA policy ''Baseline MFA''.'
                        RelatedPolicies   = 'Baseline MFA'
                    }
                )
                AdminMfaSummary = [pscustomobject]@{
                    EnabledAdminUsersReviewed            = 2
                    AdminUsersRegisteredForMfa           = 1
                    AdminUsersNotRegisteredForMfa        = 1
                    AdminUsersCoveredByMfaEnforcement    = 1
                    AdminUsersNotCoveredByMfaEnforcement = 1
                }
                LicenseSKUs = @()
                DeviceDetails = @(
                    [pscustomobject]@{
                        OperatingSystem               = 'Windows'
                        ApproximateLastSignInDateTime = (Get-Date).AddDays(-30).ToString('o')
                        IsCompliant                   = $true
                        MDMSolution                   = 'Intune'
                    },
                    [pscustomobject]@{
                        OperatingSystem               = 'iOS'
                        ApproximateLastSignInDateTime = (Get-Date).AddDays(-40).ToString('o')
                        IsCompliant                   = $false
                        MDMSolution                   = ''
                    }
                )
                DeviceManagementSummary = @{
                    Summary = [pscustomobject]@{
                        TotalDevices         = 2
                        UnmanagedDevices     = 1
                        UnsupportedOsDevices = 0
                    }
                }
                EnterpriseApplications = @{}
                GuestAccessConfiguration = @{
                    Summary = [pscustomobject]@{
                        GuestInvitationControl       = 'adminsAndGuestInviters'
                        ConditionalAccessGuestCoverage = $false
                    }
                }
                ExternalIdentityRestrictions = @{
                    Summary = [pscustomobject]@{
                        AllowInvitesFrom           = 'adminsAndGuestInviters'
                        CrossTenantPartnerCount    = 0
                        HasCrossTenantAccessPolicy = $false
                        DefaultInboundMfaTrust     = $false
                        DefaultOutboundMfaTrust    = 'Not configured'
                    }
                }
            }
            Exchange = @{
                AllRecipients = @(
                    [pscustomobject]@{
                        DisplayName          = 'User A'
                        RecipientTypeDetails = 'UserMailbox'
                        PrimarySmtpAddress   = 'usera@contoso.com'
                    },
                    [pscustomobject]@{
                        DisplayName          = 'Shared A'
                        RecipientTypeDetails = 'SharedMailbox'
                        PrimarySmtpAddress   = 'shared@fabrikam.com'
                    },
                    [pscustomobject]@{
                        DisplayName          = 'Group A'
                        RecipientTypeDetails = 'GroupMailbox'
                        PrimarySmtpAddress   = 'group@tailspintoys.com'
                    }
                )
                AllMailboxes = @{}
                PrimaryMailboxStats = @(
                    [pscustomobject]@{
                        MailboxType        = 'UserMailbox'
                        TotalItemSizeBytes = 15GB
                    },
                    [pscustomobject]@{
                        MailboxType        = 'GroupMailbox'
                        TotalItemSizeBytes = 4GB
                    }
                )
                ArchiveMailboxStats = @(
                    [pscustomobject]@{
                        MailboxType        = 'UserMailbox'
                        TotalItemSizeBytes = 6GB
                    }
                )
                EmailActivityTopSenders = @(
                    [pscustomobject]@{
                        DisplayName       = 'User A'
                        UserPrincipalName = 'usera@contoso.com'
                        SendCount         = 42
                        LastActivityDate  = '2026-04-20'
                    }
                )
                EmailActivityTopReceivers = @(
                    [pscustomobject]@{
                        DisplayName       = 'Shared A'
                        UserPrincipalName = 'shared@fabrikam.com'
                        ReceiveCount      = 55
                        LastActivityDate  = '2026-04-21'
                    }
                )
                MailFlowConnectors = @()
                PublicFolderDetails = @()
            }
            Collaboration = @{
                AllTeams    = @()
                SharePoint  = @(
                    [pscustomobject]@{
                        Title                   = 'Projects'
                        IsTeamsConnected        = $true
                        SharingCapability       = 'ExternalUserSharingOnly'
                        StorageUsedGB           = 24
                        LastContentModifiedDate = (Get-Date).AddDays(-12).ToString('o')
                    }
                )
                OneDrive    = @(
                    [pscustomobject]@{
                        Title         = 'User A OneDrive'
                        StorageUsedGB = 12
                    }
                )
                ExternalExposureFindings = @(
                    [pscustomobject]@{
                        Workload         = 'SharePoint'
                        AssetType        = 'Site'
                        Title            = 'Projects'
                        ExposureCategory = 'Stale externally shared content'
                        GapReason        = 'Dormant site still externally available'
                        ReviewPriority   = 'High'
                    },
                    [pscustomobject]@{
                        Workload         = 'OneDrive'
                        AssetType        = 'Personal Site'
                        Title            = 'User A'
                        ExposureCategory = 'Externally sharable OneDrive with ownership mismatch'
                        GapReason        = 'Owner mismatch surfaced'
                        ReviewPriority   = 'High'
                    }
                )
            }
        }

        $snapshotPath = Join-Path $TestDrive 'chart-enabled-assessment.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $result = & $script:improveScriptPath -AssessmentJsonPath $snapshotPath -OutputFolder $TestDrive -PassThru
        $customerReportEntryNames = Get-TestDocxEntryNames -Path $result.CustomerAssessmentReportPath
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath

        @($customerReportEntryNames | Where-Object { $_ -match '^word/media/customer-chart-\d+\.png$' }).Count | Should -BeGreaterOrEqual 8
        $customerReportMarkdown | Should -Match '#### MFA Enrollment Status'
        $customerReportMarkdown | Should -Match '#### Registered MFA Method Mix'
        $customerReportMarkdown | Should -Match '#### Covered vs Not Covered by Active MFA Enforcement'
        $customerReportMarkdown | Should -Match '#### MFA Enforcement Driver Breakdown'
        $customerReportMarkdown | Should -Match '#### Recent vs Stale Admins'
        $customerReportMarkdown | Should -Match '#### Admins Covered vs Not Covered by Active MFA Enforcement'
        $customerReportMarkdown | Should -Match '### Data Footprint by Workload'
        $customerReportMarkdown | Should -Match '### Device Platform Distribution'
        $customerReportMarkdown | Should -Match '### Recipient Mix by Type'
        $customerReportMarkdown | Should -Match '### Top Recipient Domains'
        $customerReportMarkdown | Should -Match '### External Exposure by Category'
        $customerReportMarkdown | Should -Not -Match '!\['
        $customerReportMarkdown | Should -Not -Match 'word/media/customer-chart'
    }

    It 'does not raise ID-010 and labels owner validation as incomplete when owner enrichment is unavailable' {
        $snapshot = New-TestEnterpriseApplicationSnapshot -EnterpriseApplications @{
            '001-OwnerValidationPending' = [pscustomobject]@{
                AppId                  = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                ServicePrincipalId     = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
                DisplayName            = 'Owner Validation Pending'
                ApplicationSource      = 'First Party'
                ApplicationSourceState = 'Collected'
                OwnerSignalState       = 'Unavailable'
                OwnerCount             = $null
                RedirectUriSignalState = 'Collected'
                InsecureRedirectUriCount = 0
                HasInsecureRedirectUris = $false
                ActivitySignalState    = 'Collected'
                HasRecentActivity      = $true
                ApplicationPermissionCount = 0
                DelegatedPermissionGrantCount = 0
                HighPrivilegePermissionCount = 0
                SsoEnabled             = $false
            }
        }

        $snapshotPath = Join-Path $TestDrive 'owner-validation-pending.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $result = & $script:improveScriptPath -AssessmentJsonPath $snapshotPath -OutputFolder $TestDrive -PassThru
        $payload = Get-Content -Raw $result.JsonPath | ConvertFrom-Json -Depth 20
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath

        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-010' }).Count | Should -Be 0
        $customerReportMarkdown | Should -Match 'Owner validation was not fully validated in this run\.'
        $customerReportMarkdown | Should -Not -Match 'without owner coverage'
    }

    It 'does not raise ID-012 and avoids inactivity count claims when activity enrichment is unavailable' {
        $snapshot = New-TestEnterpriseApplicationSnapshot -EnterpriseApplications @{
            '001-ActivityValidationPending' = [pscustomobject]@{
                AppId                  = 'cccccccc-cccc-cccc-cccc-cccccccccccc'
                ServicePrincipalId     = 'dddddddd-dddd-dddd-dddd-dddddddddddd'
                DisplayName            = 'Activity Validation Pending'
                ApplicationSource      = 'First Party'
                ApplicationSourceState = 'Collected'
                OwnerSignalState       = 'Collected'
                OwnerCount             = 1
                RedirectUriSignalState = 'Collected'
                InsecureRedirectUriCount = 0
                HasInsecureRedirectUris = $false
                ActivitySignalState    = 'Unavailable'
                HasRecentActivity      = $null
                ApplicationPermissionCount = 0
                DelegatedPermissionGrantCount = 0
                HighPrivilegePermissionCount = 0
                SsoEnabled             = $false
            }
        }

        $snapshotPath = Join-Path $TestDrive 'activity-validation-pending.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $result = & $script:improveScriptPath -AssessmentJsonPath $snapshotPath -OutputFolder $TestDrive -PassThru
        $payload = Get-Content -Raw $result.JsonPath | ConvertFrom-Json -Depth 20
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath

        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-012' }).Count | Should -Be 0
        $customerReportMarkdown | Should -Match 'Recent activity validation was not fully validated in this run\.'
        $customerReportMarkdown | Should -Not -Match 'with no recent activity signal'
    }

    It 'does not classify unknown-source apps as first-party or third-party findings' {
        $snapshot = New-TestEnterpriseApplicationSnapshot -EnterpriseApplications @{
            '001-UnknownSourceApp' = [pscustomobject]@{
                AppId                        = 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'
                ServicePrincipalId           = 'ffffffff-ffff-ffff-ffff-ffffffffffff'
                DisplayName                  = 'Unknown Source App'
                ApplicationSource            = 'Unknown'
                ApplicationSourceState       = 'Unavailable'
                OwnerSignalState             = 'Unavailable'
                OwnerCount                   = $null
                RedirectUriSignalState       = 'Unavailable'
                InsecureRedirectUriCount     = $null
                HasInsecureRedirectUris      = $null
                ActivitySignalState          = 'Unavailable'
                HasRecentActivity            = $null
                ApplicationPermissionCount   = 2
                DelegatedPermissionGrantCount = 0
                HighPrivilegePermissionCount = 0
                SsoEnabled                   = $false
            }
        }

        $snapshotPath = Join-Path $TestDrive 'unknown-source-app.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $result = & $script:improveScriptPath -AssessmentJsonPath $snapshotPath -OutputFolder $TestDrive -PassThru
        $payload = Get-Content -Raw $result.JsonPath | ConvertFrom-Json -Depth 20
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath

        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-009' }).Count | Should -Be 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-010' }).Count | Should -Be 0
        $customerReportMarkdown | Should -Not -Match 'third-party app\(s\) with application permissions'
        $customerReportMarkdown | Should -Not -Match 'without owner coverage'
    }

    It 'merges AuthenticationSSOApplications into the customer report app section when raw enterprise rows are missing' {
        $snapshot = New-TestEnterpriseApplicationSnapshot `
            -EnterpriseApplications @{} `
            -AuthenticationSsoApplications @(
                [pscustomobject]@{
                    AppId                     = '55555555-5555-5555-5555-555555555555'
                    ServicePrincipalId        = '66666666-6666-6666-6666-666666666666'
                    DisplayName               = 'Salesforce'
                    SsoEnabled                = $true
                    PreferredSingleSignOnMode = 'saml'
                    SSOMode                   = 'saml'
                    ActivitySignalState       = 'Collected'
                    HasRecentActivity         = $true
                }
            )

        $snapshotPath = Join-Path $TestDrive 'sso-merged-apps.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $result = & $script:improveScriptPath -AssessmentJsonPath $snapshotPath -OutputFolder $TestDrive -PassThru
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath

        $customerReportMarkdown | Should -Match '### SSO-Enabled Applications'
        $customerReportMarkdown | Should -Match '- Salesforce - SAML;'
        $customerReportMarkdown | Should -Match 'Applications: Salesforce; SSO: SAML;'
        $customerReportMarkdown | Should -Match '### SSO-Enabled Applications'
    }

    It 'surfaces a noteworthy app outside the alphabetical first eight in the customer report sample' {
        $enterpriseApplications = [ordered]@{}
        foreach ($name in @('Alpha App', 'Bravo App', 'Charlie App', 'Delta App', 'Echo App', 'Foxtrot App', 'Golf App', 'Hotel App')) {
            $key = ('app-{0}' -f ($enterpriseApplications.Count + 1).ToString('00'))
            $enterpriseApplications[$key] = [pscustomobject]@{
                AppId                        = ('00000000-0000-0000-0000-{0}' -f ($enterpriseApplications.Count + 1).ToString('000000000001'))
                ServicePrincipalId           = ('10000000-0000-0000-0000-{0}' -f ($enterpriseApplications.Count + 1).ToString('000000000001'))
                DisplayName                  = $name
                ApplicationSource            = 'First Party'
                ApplicationSourceState       = 'Collected'
                OwnerSignalState             = 'Collected'
                OwnerCount                   = 1
                RedirectUriSignalState       = 'Collected'
                InsecureRedirectUriCount     = 0
                HasInsecureRedirectUris      = $false
                ActivitySignalState          = 'Collected'
                HasRecentActivity            = $true
                ApplicationPermissionCount   = 0
                DelegatedPermissionGrantCount = 0
                HighPrivilegePermissionCount = 0
                SsoEnabled                   = $false
            }
        }
        $enterpriseApplications['app-09'] = [pscustomobject]@{
            AppId                        = '99999999-9999-9999-9999-999999999999'
            ServicePrincipalId           = '88888888-8888-8888-8888-888888888888'
            DisplayName                  = 'Zulu Risky App'
            ApplicationSource            = 'First Party'
            ApplicationSourceState       = 'Collected'
            SsoEnabled                   = $true
            SSOMode                      = 'saml'
            PreferredSingleSignOnMode    = 'saml'
            OwnerSignalState             = 'Collected'
            OwnerCount                   = 0
            RedirectUriSignalState       = 'Collected'
            InsecureRedirectUriCount     = 2
            HasInsecureRedirectUris      = $true
            ActivitySignalState          = 'Collected'
            HasRecentActivity            = $false
            ApplicationPermissionCount   = 1
            DelegatedPermissionGrantCount = 1
            HighPrivilegePermissionCount = 3
            AppRoleAssignmentRequired    = $true
        }

        $snapshot = New-TestEnterpriseApplicationSnapshot -EnterpriseApplications $enterpriseApplications
        $snapshotPath = Join-Path $TestDrive 'ranked-enterprise-apps.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $result = & $script:improveScriptPath -AssessmentJsonPath $snapshotPath -OutputFolder $TestDrive -PassThru
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath

        $customerReportMarkdown | Should -Match 'Applications: Zulu Risky App; SSO: SAML;'
        $customerReportMarkdown | Should -Match '- Zulu Risky App -'
    }

    It 'allows redirect-risk findings from partial data without emitting a zero-risk narrative' {
        $snapshot = New-TestEnterpriseApplicationSnapshot -EnterpriseApplications @{
            '001-RedirectRiskPartial' = [pscustomobject]@{
                AppId                        = '12121212-1212-1212-1212-121212121212'
                ServicePrincipalId           = '34343434-3434-3434-3434-343434343434'
                DisplayName                  = 'Partial Redirect Review App'
                ApplicationSource            = 'Third Party'
                ApplicationSourceState       = 'Collected'
                OwnerSignalState             = 'NotApplicable'
                OwnerCount                   = $null
                RedirectUriSignalState       = 'Partial'
                InsecureRedirectUriCount     = 1
                HasInsecureRedirectUris      = $true
                ActivitySignalState          = 'Collected'
                HasRecentActivity            = $true
                ApplicationPermissionCount   = 1
                DelegatedPermissionGrantCount = 0
                HighPrivilegePermissionCount = 0
                SsoEnabled                   = $false
            }
        }

        $snapshotPath = Join-Path $TestDrive 'partial-redirect-review.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $result = & $script:improveScriptPath -AssessmentJsonPath $snapshotPath -OutputFolder $TestDrive -PassThru
        $payload = Get-Content -Raw $result.JsonPath | ConvertFrom-Json -Depth 20
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath

        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-011' }).Count | Should -BeGreaterThan 0
        $customerReportMarkdown | Should -Match 'Redirect URI review was not fully validated in this run\.'
        $customerReportMarkdown | Should -Not -Match '1 redirect URI review flag\(s\)'
    }

    It 'derives MFA enrollment output from registration details when summary objects are missing' {
        $snapshot = New-ArrayaTenantSnapshot `
            -Data @{
                Identity = @{
                    Admins = @(
                        [pscustomobject]@{
                            Role                  = 'Global Administrator'
                            AccountEnabled        = $true
                            OnPremisesSyncEnabled = $false
                            LastSignInDateTime    = (Get-Date).AddDays(-10).ToString('o')
                            UserPrincipalName     = 'admin@contoso.com'
                        }
                    )
                    Users = @(
                        [pscustomobject]@{
                            UserType                 = 'Member'
                            AccountEnabled           = $true
                            UserPrincipalName        = 'member.one@contoso.com'
                            DisplayName              = 'Member One'
                            AssignedLicensesFriendly = @('Microsoft 365 E3')
                            LastSignInDateTime       = (Get-Date).AddDays(-5).ToString('o')
                        },
                        [pscustomobject]@{
                            UserType                 = 'Member'
                            AccountEnabled           = $true
                            UserPrincipalName        = 'member.two@contoso.com'
                            DisplayName              = 'Member Two'
                            AssignedLicensesFriendly = @('Microsoft 365 E3')
                            LastSignInDateTime       = (Get-Date).AddDays(-7).ToString('o')
                        }
                    )
                    ConditionalAccessPolicies = @(
                        [pscustomobject]@{
                            DisplayName = 'Baseline MFA'
                            State       = 'enabled'
                        }
                    )
                    ConditionalAccessPolicySummary = @{
                        Summary = [pscustomobject]@{
                            TotalPolicies                 = 1
                            EnabledPolicies               = 1
                            ReportOnlyPolicies            = 0
                            HasGuestCoverage              = $false
                            HasPrivilegedRoleCoverage     = $true
                            HasCompliantDeviceRequirement = $false
                            HasRiskBasedCoverage          = $false
                            PoliciesWithExclusions        = 0
                        }
                    }
                    SecurityDefaultsPolicy = @{
                        Summary = [pscustomobject]@{
                            IsEnabled = $false
                        }
                    }
                    MfaRegistrationDetails = @(
                        [pscustomobject]@{
                            UserPrincipalName          = 'member.one@contoso.com'
                            IsMfaRegistered            = $true
                            IsMfaCapable               = $true
                            MethodsRegistered          = @('Microsoft Authenticator')
                            DefaultMfaMethod           = 'Microsoft Authenticator'
                            HasWeakMethod              = $false
                            HasStrongMethod            = $true
                            HasPhishingResistantMethod = $false
                        },
                        [pscustomobject]@{
                            UserPrincipalName          = 'member.two@contoso.com'
                            IsMfaRegistered            = $false
                            IsMfaCapable               = $false
                            MethodsRegistered          = @()
                            DefaultMfaMethod           = ''
                            HasWeakMethod              = $false
                            HasStrongMethod            = $false
                            HasPhishingResistantMethod = $false
                        }
                    )
                }
            }

        $snapshotPath = Join-Path $TestDrive 'mfa-fallback-assessment.json'
        Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $snapshotPath

        $result = & $script:improveScriptPath -AssessmentJsonPath $snapshotPath -OutputFolder $TestDrive -PassThru

        Test-Path $result.CustomerAssessmentReportMarkdownPath | Should -BeTrue
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath

        $customerReportMarkdown | Should -Match '## 6\.0 Authentication Methods, MFA Enrollment, and MFA Enforcement'
        $customerReportMarkdown | Should -Match '\| Users reviewed \| 2 \|'
        $customerReportMarkdown | Should -Match '\| Registered for MFA \| 1 \|'
        $customerReportMarkdown | Should -Match '\| Not registered for MFA \| 1 \|'
        $customerReportMarkdown | Should -Match '\| MFA enrollment rate \| 50% \|'
        $customerReportMarkdown | Should -Match '\| Registered method mix \|'
        $customerReportMarkdown | Should -Match '\| Enabled users reviewed \| 2 \|'
        $customerReportMarkdown | Should -Match '\| Enabled member users reviewed \| 2 \|'
        $customerReportMarkdown | Should -Match 'The reviewed enabled-user inventory included 2 enabled account\(s\), but the current source did not surface a complete covered-user total\.'
        $customerReportMarkdown | Should -Match 'Enrollment shows readiness, not enforcement\. In the reviewed data, MFA enrollment is 50%'
        $customerReportMarkdown | Should -Not -Match '\| MFA enrollment rate \| Not validated from the reviewed data \|'
        $customerReportMarkdown | Should -Not -Match 'an unconfirmed number of an unconfirmed number'
    }
}
