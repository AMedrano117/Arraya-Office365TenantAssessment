Describe 'Arraya.M365.Common' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
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
            ExternalIdentityRestrictions = @{
                Summary = [pscustomobject]@{
                    AllowInvitesFrom        = 'adminsAndGuestInviters'
                    CrossTenantPartnerCount = 2
                    DefaultInboundMfaTrust  = $true
                }
            }
            GuestAccessConfiguration = @{
                Summary = [pscustomobject]@{
                    GuestInvitationControl       = 'adminsAndGuestInviters'
                    ConditionalAccessGuestCoverage = $true
                }
            }
        }

        $snapshot = Convert-ArrayaLegacyTenantStatsToSnapshot -TenantStatsHash $legacyTenantStats

        $snapshot.Data.Tenant.ExternalSharingSummary.Summary.TenantSharingCapability | Should -Be 'ExternalUserAndGuestSharing'
        $snapshot.Data.Tenant.ExternalSharingSiteOverrides.Count | Should -Be 1
        $snapshot.Data.Identity.ExternalIdentityRestrictions.Summary.AllowInvitesFrom | Should -Be 'adminsAndGuestInviters'
        $snapshot.Data.Identity.GuestAccessConfiguration.Summary.GuestInvitationControl | Should -Be 'adminsAndGuestInviters'
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
        }

        Export-HashTableToExcel -hashtable $tenantStats -ExportDetails $exportPath

        $worksheetNames = @(Get-ExcelSheetInfo -Path $exportPath | Select-Object -ExpandProperty Name)
        ($worksheetNames -contains 'GuestAccessConfiguration') | Should -BeTrue
        ($worksheetNames -contains 'ExternalIdentityRestrictions') | Should -BeTrue
        ($worksheetNames -contains 'ExternalSharingSummary') | Should -BeTrue
        ($worksheetNames -contains 'ExternalSharingSiteOverrides') | Should -BeTrue
        $worksheetNames.IndexOf('GuestSignInSummary') | Should -BeLessThan $worksheetNames.IndexOf('GuestAccessConfiguration')
        $worksheetNames.IndexOf('GuestAccessConfiguration') | Should -BeLessThan $worksheetNames.IndexOf('ExternalIdentityRestrictions')
        $worksheetNames.IndexOf('SharePointSharingSummary') | Should -BeLessThan $worksheetNames.IndexOf('ExternalSharingSummary')
        $worksheetNames.IndexOf('ExternalSharingSummary') | Should -BeLessThan $worksheetNames.IndexOf('ExternalSharingSiteOverrides')
    }
}
