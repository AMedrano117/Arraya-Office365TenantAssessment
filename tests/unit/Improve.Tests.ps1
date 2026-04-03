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
                            DisplayName                  = 'High Priv App'
                            HighPrivilegePermissionCount = 2
                            HighPrivilegePermissions     = 'Application.ReadWrite.All,Directory.ReadWrite.All'
                            ApplicationPermissionCount   = 1
                            DelegatedPermissionGrantCount = 1
                        }
                    }
                    EnterpriseApplicationSummary = @{
                        Summary = [pscustomobject]@{
                            TotalEnterpriseApplications   = 1
                            ApplicationsWithHighPrivilege = 1
                        }
                    }
                    GuestSignInSummary = @{
                        Summary = [pscustomobject]@{
                            InactiveGuests90Days = 12
                        }
                    }
                    ExternalIdentityRestrictions = @{
                        Summary = [pscustomobject]@{
                            AllowInvitesFrom        = 'adminsAndGuestInviters'
                            CrossTenantPartnerCount = 2
                            DefaultInboundMfaTrust  = $true
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
                    }
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
                            LastContentModifiedDate = (Get-Date).AddDays(-220).ToString('o')
                            StorageUsedGB           = 1200
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
                            TenantSharingCapability = 'ExternalUserAndGuestSharing'
                            DefaultSharingLinkType  = 'AnonymousAccess'
                        }
                    }
                    ExternalSharingSummary = @{
                        Summary = [pscustomobject]@{
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
        $result.CsvPath | Should -BeNullOrEmpty
        $result.MarkdownPath | Should -BeNullOrEmpty
        Test-Path $result.CustomerAssessmentReportPath | Should -BeTrue
        $result.CustomerAssessmentReportPath | Should -Match '\.docx$'
        Test-Path $result.CustomerAssessmentReportMarkdownPath | Should -BeTrue
        $result.CustomerAssessmentReportMarkdownPath | Should -Match '\.md$'
        if ($result.PSObject.Properties.Name -contains 'CustomerRemediationReportPath') {
            $result.CustomerRemediationReportPath | Should -BeNullOrEmpty
        }
        if ($result.PSObject.Properties.Name -contains 'CustomerRemediationReportMarkdownPath') {
            $result.CustomerRemediationReportMarkdownPath | Should -BeNullOrEmpty
        }
        Test-Path $result.EngineerActionPackPath | Should -BeTrue
        Test-Path $result.RemediationPs1Path | Should -BeTrue
        Split-Path -Path $result.RemediationPs1Path -Parent | Should -Be (Join-Path $TestDrive 'Support')
        $result.SupportFolderPath | Should -Be (Join-Path $TestDrive 'Support')

        $payload = Get-Content -Raw $result.JsonPath | ConvertFrom-Json -Depth 20
        $payload.Findings.Count | Should -BeGreaterThan 0
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
        @($payload.WorkstreamSummaries).Count | Should -BeGreaterThan 0

        @($payload.Findings | Where-Object { $_.RuleId -eq 'SEC-001' }).Count | Should -Be 0
        @($payload.Findings | Where-Object { $_.RuleId -like 'AREA-*' }).Count | Should -Be 0
        @($payload.Findings | Where-Object { $_.Source -eq 'Derived/BestPracticeFindings' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-007' }).Count | Should -BeGreaterThan 0
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
        $customerReportMarkdown = Get-Content -Raw $result.CustomerAssessmentReportMarkdownPath
        $ns = New-Object System.Xml.XmlNamespaceManager($customerReportXml.NameTable)
        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $customerTables = $customerReportXml.SelectNodes('//w:tbl', $ns)
        $customerReportDocument | Should -Match 'Introduction'
        $customerReportDocument | Should -Match 'Project Scope'
        $customerReportDocument | Should -Match 'Executive Summary'
        $customerReportDocument | Should -Match 'Modern Workplace Recommendations'
        $customerReportDocument | Should -Match 'Entra ID Review: User and Device Inventory'
        $customerReportDocument | Should -Match 'Device Registration and Compliance Gaps'
        $customerReportDocument | Should -Match 'Entra Guest Access Configuration'
        $customerReportDocument | Should -Match 'Entra Applications and Access Review'
        $customerReportDocument | Should -Match 'Modernizing Authentication: Duo Integration, Conditional Access, and MFA Coverage'
        $customerReportDocument | Should -Match 'Password Writeback and Self-Service Password Reset'
        $customerReportDocument | Should -Match 'Authorization: Admin Access and Role Assignments'
        $customerReportDocument | Should -Match 'Exchange Online: Mailboxes and Storage Overview'
        $customerReportDocument | Should -Match 'User Mailbox Growth'
        $customerReportDocument | Should -Match 'Shared Mailbox Review'
        $customerReportDocument | Should -Match 'Inactive Mailboxes'
        $customerReportDocument | Should -Match 'Group Mailbox Utilization'
        $customerReportDocument | Should -Match 'Archive Mailbox Usage and Licensing'
        $customerReportDocument | Should -Match 'SMTP Relay Usage and Service Accounts'
        $customerReportDocument | Should -Match 'Microsoft Teams Governance and Cleanup'
        $customerReportDocument | Should -Match 'SharePoint Online Storage and External Sharing'
        $customerReportDocument | Should -Match 'Retention Policies and Data Loss Prevention'
        $customerReportDocument | Should -Match 'Domain Configuration and DNS Overview'
        $customerReportDocument | Should -Match 'Offboarding Recommendation'
        $customerReportDocument | Should -Match 'Application User Consent Management'
        $customerReportDocument | Should -Match 'Authentication Methods Migration'
        $customerReportDocument | Should -Match 'DMARC Records'
        $customerReportDocument | Should -Match 'Password Writeback with AD Sync'
        $customerReportDocument | Should -Match 'Version History'
        $customerReportDocument | Should -Match 'Arraya Solutions'
        $customerReportDocument | Should -Match 'Not surfaced in current source'
        $customerReportDocument | Should -Match 'Configuration Signal'
        $customerReportDocument | Should -Match 'Current State'
        $customerReportDocument | Should -Match 'Guest invitation control'
        $customerReportDocument | Should -Match 'Cross-tenant partner count'
        $customerReportDocument | Should -Match 'Default inbound MFA trust'
        $customerReportDocument | Should -Match 'adminsAndGuestInviters'
        $customerReportDocument | Should -Match 'allowList'
        $customerReportDocument | Should -Match 'Microsoft Guidance'
        $customerReportDocument | Should -Match 'Why It Is Relevant'
        $customerReportDocument | Should -Match 'Recommendation'
        $customerReportDocument | Should -Match 'Criticality / Impact'
        $customerReportDocument | Should -Match 'Level of Effort'
        $customerReportDocument | Should -Match 'Action:'
        $customerReportDocument | Should -Not -Match 'AREA-'
        $customerReportDocument | Should -Not -Match 'CustomerRemediationReport'
        $customerReportDocument | Should -Not -Match 'unauthorized mailbox forwarding'
        $customerReportDocument | Should -Not -Match 'unauthorized inbox-rule forwarding'
        $customerTables.Count | Should -BeGreaterThan 10
        $customerTables[0].SelectNodes('./w:tr[1]/w:tc', $ns).Count | Should -Be 5
        $customerTables[0].SelectNodes('./w:tr[2]/w:tc', $ns).Count | Should -Be 5
        $customerTables[1].SelectNodes('./w:tr[1]/w:tc', $ns).Count | Should -Be 3
        $customerTables[1].SelectNodes('./w:tr[2]/w:tc', $ns).Count | Should -Be 3
        $customerReportMarkdown | Should -Match '# .+ Microsoft 365 Tenant Best Practices Assessment'
        $customerReportMarkdown | Should -Match '## Executive Summary'
        $customerReportMarkdown | Should -Match '## Modern Workplace Recommendations'
        $customerReportMarkdown | Should -Match '\| Recommendation \| Criticality / Impact \| Level of Effort \|'
        $customerReportMarkdown | Should -Match '## Entra Guest Access Configuration'
        $customerReportMarkdown | Should -Match '## SharePoint Online Storage and External Sharing'
        $customerReportMarkdown | Should -Match '## Retention Policies and Data Loss Prevention'
        $customerReportMarkdown | Should -Match '## Offboarding Recommendation'
        $customerReportMarkdown | Should -Match 'adminsAndGuestInviters'
        $customerReportMarkdown | Should -Match 'Guest invitation control'
        $customerReportMarkdown | Should -Not -Match 'CustomerRemediationReport'

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
    }
}
