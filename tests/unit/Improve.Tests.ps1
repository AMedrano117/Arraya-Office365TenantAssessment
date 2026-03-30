Describe 'Improve workflow' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:commonManifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        $script:improveScriptPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\New-M365TenantImprovementPlan.ps1'

        Import-Module -Name $script:commonManifestPath -Force -ErrorAction Stop
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
                            Id         = 'contoso.com'
                            IsVerified = $false
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
        $result.CsvPath | Should -BeNullOrEmpty
        $result.MarkdownPath | Should -BeNullOrEmpty
        Test-Path $result.CustomerRemediationReportPath | Should -BeTrue
        Test-Path $result.EngineerActionPackPath | Should -BeTrue
        Test-Path $result.RemediationPs1Path | Should -BeTrue

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

        $id007 = @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-007' }) | Select-Object -First 1
        $id007.WhyFlagged | Should -Match 'permission'
        $id007.ExampleAction | Should -Match 'HighPrivilegePermissions|enterprise apps'

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

        $customerReport = Get-Content -Raw $result.CustomerRemediationReportPath
        $customerReport | Should -Match '## Executive Summary'
        $customerReport | Should -Match '## Phased Roadmap'
        $customerReport | Should -Match '## Workstream Summary'
        $customerReport | Should -Not -Match 'AREA-'
        $customerReport | Should -Match 'Current State'
        $customerReport | Should -Match '1 mailbox\(es\) with forwarding configured'
        $customerReport | Should -Match 'external-forwarding inbox rule\(s\)'
        $customerReport | Should -Match 'hosted outbound policy/policies explicitly allow auto-forwarding'
        $customerReport | Should -Match 'remote domain\(s\) have AutoForwardEnabled'
        $customerReport | Should -Not -Match 'unauthorized mailbox forwarding'
        $customerReport | Should -Not -Match 'unauthorized inbox-rule forwarding'

        $engineerPack = Get-Content -Raw $result.EngineerActionPackPath
        $engineerPack | Should -Match '## Normalized Findings'
        $engineerPack | Should -Match '## Command References'
        $engineerPack | Should -Match 'Why Flagged'
        $engineerPack | Should -Match 'Example'
        $engineerPack | Should -Not -Match 'CSV output:'

        $snippetContent = Get-Content -Raw $result.RemediationPs1Path
        $snippetContent | Should -Match '\$gaRole = Get-MgDirectoryRole'
    }
}
