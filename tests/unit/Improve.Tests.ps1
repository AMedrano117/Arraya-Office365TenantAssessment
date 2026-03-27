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
                    EnterpriseApplications = @(
                        [pscustomobject]@{
                            DisplayName                  = 'High Priv App'
                            HighPrivilegePermissionCount = 2
                        }
                    )
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
                    AllMailboxes = @(
                        [pscustomobject]@{
                            DisplayName             = 'Forwarded Mailbox'
                            DeliverToMailboxAndForward = $true
                            ForwardingSmtpAddress   = 'external@example.com'
                            RecipientTypeDetails    = 'UserMailbox'
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
        Test-Path $result.CsvPath | Should -BeTrue
        Test-Path $result.MarkdownPath | Should -BeTrue
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
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'RelatedWorksheet'
        ($payload.Findings | Select-Object -First 1).PSObject.Properties.Name | Should -Contain 'RelatedSection'

        @($payload.Findings | Where-Object { $_.RuleId -eq 'SEC-001' }).Count | Should -Be 0
        @($payload.Findings | Where-Object { $_.Source -eq 'Derived/BestPracticeFindings' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-007' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'ID-005' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'CA-002' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'CA-010' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'CA-012' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'EX-007' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'DEV-005' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'DEV-006' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'COL-005' }).Count | Should -BeGreaterThan 0
        @($payload.Findings | Where-Object { $_.RuleId -eq 'COL-006' }).Count | Should -BeGreaterThan 0

        $customerReport = Get-Content -Raw $result.CustomerRemediationReportPath
        $customerReport | Should -Match '## Executive Summary'
        $customerReport | Should -Match '## Phased Roadmap'

        $engineerPack = Get-Content -Raw $result.EngineerActionPackPath
        $engineerPack | Should -Match '## Normalized Findings'
        $engineerPack | Should -Match '## Command References'

        $snippetContent = Get-Content -Raw $result.RemediationPs1Path
        $snippetContent | Should -Match '\$gaRole = Get-MgDirectoryRole'
    }
}
