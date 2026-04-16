Describe 'Arraya.M365.Common' {
    BeforeAll {
        $script:previousImportModuleWarningPreference = $PSDefaultParameterValues['Import-Module:WarningAction']
        $PSDefaultParameterValues['Import-Module:WarningAction'] = 'SilentlyContinue'
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        $script:exportExcelPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Public\Export-HashTableToExcel.ps1'
        $script:exportExcelSource = Get-Content -Raw -Path $script:exportExcelPath
        $script:expectedExports = @(
            'Convert-ArrayaLegacyTenantStatsToSnapshot'
            'Convert-ArrayaObjectToArray'
            'Convert-ArrayaSnapshotToLegacyTenantStatsHash'
            'Convert-ArrayaToDate'
            'Convert-ArrayaToNumber'
            'ConvertTo-ExportFriendlyRecord'
            'ConvertTo-ExportFriendlyValue'
            'Export-ArrayaErrorReports'
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
        ($worksheetNames -contains 'HybridConfiguration') | Should -BeTrue
        $worksheetNames.IndexOf('TenantInfo') | Should -BeLessThan $worksheetNames.IndexOf('GuestSignInSummary')
        $worksheetNames.IndexOf('TenantInfo') | Should -BeLessThan $worksheetNames.IndexOf('LicenseSKUs')
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
    }

    It 'uses a single workbook package session for faster Excel export' {
        $script:exportExcelSource | Should -Match 'Open-ExcelPackage -Path \$ExportDetails -Create'
        $script:exportExcelSource | Should -Match 'Export-Excel -ExcelPackage \$excelPackage'
        $script:exportExcelSource | Should -Match 'Close-ExcelPackage -ExcelPackage \$excelPackage'
        $script:exportExcelSource | Should -Match '\$autoSizeRowLimit = 1000'
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
