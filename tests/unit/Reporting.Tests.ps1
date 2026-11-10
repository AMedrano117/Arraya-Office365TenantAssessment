Describe 'Arraya.M365.Reporting' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Reporting\Arraya.M365.Reporting.psd1'
        $script:moduleRoot = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Reporting'
        $script:expectedExports = @(
            'Get-ArrayaAssessmentRecommendationText'
            'Get-ArrayaAssessmentWorksheetName'
            'Get-ArrayaEmployeeExperienceInsightsAnalysis'
            'Get-GraphUserStats'
        )
    }

    It 'has a manifest' {
        Test-Path $script:manifestPath | Should -BeTrue
    }

    It 'exports the reduced reporting surface' {
        $manifest = Import-PowerShellDataFile -Path $script:manifestPath
        $manifest.FunctionsToExport | Should -Be $script:expectedExports
        $manifest.VariablesToExport | Should -Be @()
    }

    It 'does not keep retired reporting wrapper files' {
        @(
            'Public\Get-ArrayaReportingRuntimeContext.ps1'
            'Public\Set-ArrayaReportingRuntimeContext.ps1'
            'Public\New-ArrayaFinding.ps1'
        ) | ForEach-Object {
            Test-Path (Join-Path $script:moduleRoot $_) | Should -BeFalse
        }
    }

    It 'does not reference retired wrappers or hidden runtime globals in source' {
        $matches = Get-ChildItem -Path $script:moduleRoot -Recurse -Include *.ps1,*.psm1 |
            Select-String -Pattern @(
                '\$script:tenantStatsHash'
                '\$ExportDetails'
                '\$global:initialStart'
                'Get-ArrayaReportingRuntimeContext'
                'Set-ArrayaReportingRuntimeContext'
                'Invoke-QuietRestMethod'
            ) -CaseSensitive
        $matches | Should -BeNullOrEmpty
    }

    It 'analyzes generic and concurrent summary dictionaries with present and missing keys' {
        Import-Module $script:manifestPath -Force -ErrorAction Stop

        $emailSummary = [System.Collections.Generic.Dictionary[string, object]]::new()
        $emailSummary['ReportRows'] = 7
        $emailSummary['ActiveUsers'] = 3
        $emailSummary['PeriodDuration'] = 'D30'

        $employeeSummary = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
        $employeeSummary['TeamsReportRows'] = 2

        $authConfig = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
        $authConfig['MFAEnabled'] = $true

        $result = Get-ArrayaEmployeeExperienceInsightsAnalysis `
            -EmailActivitySummary $emailSummary `
            -EmployeeExperienceInsightsSummary $employeeSummary `
            -AuthConfig $authConfig `
            -TotalUsers 10

        $result | Should -BeOfType ([hashtable])
        $result.ContainsKey('Findings') | Should -BeTrue
        @($result.Findings).Count | Should -BeGreaterThan 0
        @($result.Findings | ForEach-Object { $_.Category }) | Should -Contain 'Conditional Access Enforcement'
        @($result.Findings | ForEach-Object { $_.Category }) | Should -Contain 'MFA Coverage'
    }

    It 'classifies lowercase generic-dictionary report-only MFA policies as staged rather than enabled' {
        Import-Module $script:manifestPath -Force -ErrorAction Stop

        $grantControls = [System.Collections.Generic.Dictionary[string, object]]::new()
        $grantControls['builtInControls'] = @('mfa')

        $reportOnlyPolicy = [System.Collections.Generic.Dictionary[string, object]]::new()
        $reportOnlyPolicy['displayName'] = 'Report-only MFA pilot'
        $reportOnlyPolicy['state'] = 'enabledForReportingButNotEnforced'
        $reportOnlyPolicy['grantControls'] = $grantControls

        $result = Get-ArrayaEmployeeExperienceInsightsAnalysis -ConditionalAccessPolicies @($reportOnlyPolicy)
        $enabledFindings = @($result.Findings | Where-Object { $_.Category -eq 'Conditional Access Enforcement' })
        $stagingFindings = @($result.Findings | Where-Object { $_.Category -eq 'Conditional Access Staging' })
        $mfaFindings = @($result.Findings | Where-Object { $_.Category -eq 'MFA Coverage' })

        $enabledFindings.Count | Should -Be 1
        $enabledFindings[0].Message | Should -Match 'none are in an enabled state'
        $enabledFindings[0].Message | Should -Not -Match 'enabled policies detected: 1'
        $stagingFindings.Count | Should -Be 1
        $stagingFindings[0].Message | Should -Match '^1 Conditional Access policy/policies are in report-only state'
        $mfaFindings.Count | Should -Be 1
        $mfaFindings[0].Message | Should -Match '^1 MFA-grant Conditional Access policy/policies are report-only'
        $mfaFindings[0].Message | Should -Match 'enabled MFA Conditional Access enforcement evidence was not detected'
    }

    It 'counts the normalized Graph report-only state separately in Conditional Access KPIs' {
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $reportOnlyGrantControls = [System.Collections.Generic.Dictionary[string, object]]::new()
        $reportOnlyGrantControls['builtInControls'] = @('mfa')
        $reportOnlyPolicy = [System.Collections.Generic.Dictionary[string, object]]::new()
        $reportOnlyPolicy['state'] = 'enabledForReportingButNotEnforced'
        $reportOnlyPolicy['grantControls'] = $reportOnlyGrantControls

        $policies = @(
            [pscustomobject]@{ State = 'enabled'; GrantControls_BuiltInControls = 'compliantDevice' }
            $reportOnlyPolicy
            [pscustomobject]@{ State = 'disabled'; GrantControls_BuiltInControls = 'mfa' }
        )
        $html = Build-ConditionalAccessMfaSection `
            -ConditionalAccessPolicies $policies `
            -AuthConfig ([pscustomobject]@{ MFAEnabled = $true; MFAMethods = @('email') }) `
            -Users @() `
            -MfaRegistrationSummary $null

        $html | Should -Match '(?s)<div class="kpi-title">CA Policies</div>\s*<div class="kpi-value">3\s*</div>'
        $html | Should -Match '(?s)<div class="kpi-title">Enabled</div>\s*<div class="kpi-value">1\s*</div>'
        $html | Should -Match '(?s)<div class="kpi-title">Report-Only</div>\s*<div class="kpi-value">1\s*</div>'
        $html | Should -Match '(?s)<div class="kpi-title">MFA CA Enforcement</div>\s*<div class="kpi-value">Not Evidenced\s*</div>'
    }

    It 'does not treat a disabled MFA-grant Conditional Access policy as enforcement evidence' {
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $analysis = Get-ConditionalAccessMfaAnalysis `
            -ConditionalAccessPolicies @([pscustomobject]@{
                    State = 'disabled'
                    GrantControls_BuiltInControls = 'mfa'
                }) `
            -AuthConfig $null `
            -TotalUsers 1

        $mfaFindings = @($analysis.Findings | Where-Object { $_.Category -eq 'MFA Enforcement' })
        $mfaFindings.Count | Should -Be 1
        $mfaFindings[0].Message | Should -Match 'No enabled Conditional Access MFA-enforcement evidence'
    }

    It 'does not infer MFA CA enforcement from an unrelated enabled authentication method' {
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $authConfig = [pscustomobject]@{
            MFAEnabled = $true
            MFAMethods = @('email')
        }
        $enabledNonMfaPolicy = [pscustomobject]@{
            State = 'enabled'
            GrantControls_BuiltInControls = 'compliantDevice'
        }
        $analysis = Get-ConditionalAccessMfaAnalysis -ConditionalAccessPolicies @($enabledNonMfaPolicy) -AuthConfig $authConfig -TotalUsers 1
        $html = Build-ConditionalAccessMfaSection -ConditionalAccessPolicies @($enabledNonMfaPolicy) -AuthConfig $authConfig -Users @() -MfaRegistrationSummary $null

        @($analysis.Findings | Where-Object { $_.Category -eq 'MFA Enforcement' }).Count | Should -Be 1
        $html | Should -Match '(?s)<div class="kpi-title">MFA CA Enforcement</div>\s*<div class="kpi-value">Not Evidenced\s*</div>'
        $html | Should -Match 'Authentication Methods Available'
        $html | Should -Not -Match '<div class="kpi-title">MFA Enforced</div>'
    }

    It 'asserts MFA CA enforcement only for an enabled MFA-grant policy' {
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $enabledMfaPolicy = [pscustomobject]@{
            State = 'enabled'
            GrantControls_BuiltInControls = 'mfa'
        }
        $analysis = Get-ConditionalAccessMfaAnalysis -ConditionalAccessPolicies @($enabledMfaPolicy) -AuthConfig $null -TotalUsers 1
        $html = Build-ConditionalAccessMfaSection -ConditionalAccessPolicies @($enabledMfaPolicy) -AuthConfig $null -Users @() -MfaRegistrationSummary $null

        @($analysis.Findings | Where-Object { $_.Category -eq 'MFA Enforcement' }).Count | Should -Be 0
        $html | Should -Match '(?s)<div class="kpi-title">MFA CA Enforcement</div>\s*<div class="kpi-value">Evidenced\s*</div>'
    }

    It 'reads generic and concurrent dictionary values through HTML helpers' {
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $genericRecord = [System.Collections.Generic.Dictionary[string, object]]::new()
        $genericRecord['DisplayName'] = 'Generic tenant'

        Get-RecordValue -Record $genericRecord -Key 'DisplayName' | Should -Be 'Generic tenant'
        Get-RecordValue -Record $genericRecord -Key 'Missing' | Should -BeNullOrEmpty

        $licenseState = [System.Collections.Generic.Dictionary[string, object]]::new()
        $licenseState['SkuId'] = 'sku-1'
        $licenseState['State'] = 'Active'

        $concurrentUser = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
        $concurrentUser['LicenseAssignmentStates'] = $licenseState

        Get-AssessmentRuleValue -Record $concurrentUser -Names @('LicenseAssignmentStates') | Should -Be $licenseState
        Get-AssessmentRuleValue -Record $concurrentUser -Names @('Missing') | Should -BeNullOrEmpty

        $assignmentRows = @(Get-ArrayaUserLicenseAssignmentRows -User $concurrentUser)
        $assignmentRows.Count | Should -Be 1
        $assignmentRows[0].SkuId | Should -Be 'sku-1'
        $assignmentRows[0].State | Should -Be 'Active'
    }

    It 'renders absent Presales readiness tables as warning Needs Data values instead of green zeros' {
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $outputPath = Join-Path $TestDrive 'presales-missing-readiness.html'
        $result = New-TenantMigrationScopeHtmlReport -TenantStatsHash @{
            TenantInfo = @([pscustomobject]@{ DisplayName = 'Legacy Snapshot'; DefaultDomain = 'legacy.example' })
        } -OutputPath $outputPath

        $result.Success | Should -BeTrue
        $html = Get-Content -LiteralPath $outputPath -Raw
        $html | Should -Match '(?s)<div class="kpi-card kpi-warning">\s*<div class="kpi-title">Quote Readiness</div>\s*<div class="kpi-value">Needs Data'
        $html | Should -Match '(?s)<div class="kpi-card kpi-warning">\s*<div class="kpi-title">Unresolved Blockers</div>\s*<div class="kpi-value">Needs Data'
        $html | Should -Match '(?s)<div class="kpi-card kpi-warning">\s*<div class="kpi-title">Customer Inputs Needed</div>\s*<div class="kpi-value">Needs Data'
        $html | Should -Not -Match '(?s)<div class="kpi-card kpi-success">\s*<div class="kpi-title">Unresolved Blockers</div>\s*<div class="kpi-value">0'
    }

    It 'reports source tenant licensing, public folders, and the device estate in the Presales brief' {
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $outputPath = Join-Path $TestDrive 'presales-source-insight.html'
        $result = New-TenantMigrationScopeHtmlReport -TenantStatsHash @{
            TenantInfo = [ordered]@{ DisplayName = 'Source Insight'; DefaultDomain = 'insight.example' }
            LicenseSKUs = @{
                'sku-1' = [pscustomobject]@{ SkuPartNumber = 'SPE_E3'; SkuFriendlyName = 'Microsoft 365 E3'; PurchasedUnits = 250; ConsumedUnits = 238; LicenseClass = 'Paid'; NextLifecycleDateTime = '2027-03-01T00:00:00Z' }
                'sku-2' = [pscustomobject]@{ SkuPartNumber = 'EXCHANGESTANDARD'; SkuFriendlyName = 'Exchange Online (Plan 1)'; PurchasedUnits = 40; ConsumedUnits = 44; LicenseClass = 'Paid'; NextLifecycleDateTime = '9999-12-31T23:59:59Z' }
                'sku-3' = [pscustomobject]@{ SkuPartNumber = 'FLOW_FREE'; SkuFriendlyName = 'Microsoft Power Automate Free'; PurchasedUnits = 10000; ConsumedUnits = 12; LicenseClass = 'FreeTrialBenefit'; NextLifecycleDateTime = $null }
            }
            PublicFolderDetails = @{
                'pf1' = [pscustomobject]@{ Identity = 'Shared\Archive'; FolderPath = '\Shared\Archive'; MailEnabled = $false; ItemCount = 41200; TotalItemSize = '18.4 GB (19,756,834,816 bytes)'; TotalDeletedItemSize = '1.1 GB (1,181,116,006 bytes)' }
                'pf2' = [pscustomobject]@{ Identity = 'Shared\Orders'; FolderPath = '\Shared\Orders'; MailEnabled = $true; ItemCount = 8100; TotalItemSize = '2.6 GB (2,791,728,742 bytes)'; TotalDeletedItemSize = $null }
                'pf3' = [pscustomobject]@{ Identity = 'Shared\Legacy'; FolderPath = '\Shared\Legacy'; MailEnabled = $false; ItemCount = $null; TotalItemSize = $null; TotalDeletedItemSize = $null }
            }
            DeviceDetails = @{
                'd1' = [pscustomobject]@{ DisplayName = 'WS-001'; OperatingSystem = 'Windows'; DeviceJoinType = 'Entra AD joined'; IsManaged = $true; DeviceStale = $false }
                'd2' = [pscustomobject]@{ DisplayName = 'WS-002'; OperatingSystem = 'Windows'; DeviceJoinType = 'Hybrid Entra AD joined'; IsManaged = $true; DeviceStale = $false }
                'd3' = [pscustomobject]@{ DisplayName = 'WS-003'; OperatingSystem = 'Windows'; DeviceJoinType = 'Hybrid Entra AD joined'; IsManaged = $false; DeviceStale = $true }
                'd4' = [pscustomobject]@{ DisplayName = 'MAC-01'; OperatingSystem = 'MacMDM'; DeviceJoinType = 'Entra AD registered'; IsManaged = $true; DeviceStale = $false }
            }
        } -OutputPath $outputPath

        $result.Success | Should -BeTrue
        $html = Get-Content -LiteralPath $outputPath -Raw

        # Sections exist and are navigable.
        $html | Should -Match '<h2>Source Tenant Licensing</h2>'
        $html | Should -Match '<h2>Public Folders</h2>'
        $html | Should -Match '<h2>Device Estate</h2>'

        # Licensing totals count paid SKUs only and flag the over-assigned subscription.
        $html | Should -Match '(?s)<div class="kpi-title">Paid Seats Assigned</div>\s*<div class="kpi-value">282'
        $html | Should -Match '(?s)<div class="kpi-title">Paid Seats Purchased</div>\s*<div class="kpi-value">290'
        $html | Should -Match '(?s)<div class="kpi-title">Free or Trial SKUs</div>\s*<div class="kpi-value">1'
        $html | Should -Match 'Assigned seats exceed purchased seats'
        $html | Should -Match '<td class="risk-cell">-4</td>'

        # Lifecycle dates are published in UTC and must not shift into the report machine's zone.
        $html | Should -Match '<td>2027-03-01</td>'
        $html | Should -Match '<td>No scheduled expiry</td>'

        # Public folders are sized and item-counted, not just counted.
        $html | Should -Match '(?s)<div class="kpi-title">Public Folders</div>\s*<div class="kpi-value">3'
        $html | Should -Match '(?s)<div class="kpi-title">Mail-Enabled Folders</div>\s*<div class="kpi-value">1'
        $html | Should -Match '(?s)<div class="kpi-title">Items</div>\s*<div class="kpi-value">49,300'
        $html | Should -Match 'Public folders are a separate priced workstream'

        # A folder with no statistics is reported as needing data, never as a measured zero.
        $html | Should -Match '(?s)<td>\\Shared\\Legacy</td>.*?<td>Needs data</td><td>Needs data</td><td>Needs data</td>'

        # Device counts carry an OS and a join-type breakdown.
        $html | Should -Match '(?s)<div class="kpi-title">Total Devices</div>\s*<div class="kpi-value">4'
        $html | Should -Match '(?s)<div class="kpi-title">Stale Devices</div>\s*<div class="kpi-value">1'
        $html | Should -Match '(?s)<div class="kpi-title">Hybrid Joined</div>\s*<div class="kpi-value">2'
        $html | Should -Match 'Breakdown by operating system'
        $html | Should -Match 'Breakdown by join type'
        $html | Should -Match '<tr><td>Windows</td><td>3</td><td>75.0%</td><td>2</td><td>1</td></tr>'
        $html | Should -Match '<tr><td>Hybrid Entra AD joined</td><td>2</td><td>50.0%</td><td>1</td><td>1</td></tr>'
    }

    It 'falls back to TrustType for the join-type breakdown when snapshots record Unknown' {
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        # Replaying an older snapshot must not report an entire estate as unclassified just
        # because DeviceJoinType was baked in before the classification bug was fixed.
        $outputPath = Join-Path $TestDrive 'presales-legacy-devices.html'
        $result = New-TenantMigrationScopeHtmlReport -TenantStatsHash @{
            TenantInfo = [ordered]@{ DisplayName = 'Legacy Devices'; DefaultDomain = 'legacy.example' }
            DeviceDetails = @{
                'd1' = [pscustomobject]@{ DisplayName = 'D1'; OperatingSystem = 'Windows'; TrustType = 'AzureAd'; DeviceJoinType = 'Unknown'; IsManaged = $true; DeviceStale = $false }
                'd2' = [pscustomobject]@{ DisplayName = 'D2'; OperatingSystem = 'Windows'; TrustType = 'ServerAd'; DeviceJoinType = 'Unknown'; IsManaged = $false; DeviceStale = $true }
                'd3' = [pscustomobject]@{ DisplayName = 'D3'; OperatingSystem = 'iOS'; TrustType = 'Workplace'; DeviceJoinType = 'Unknown'; IsManaged = $false; DeviceStale = $false }
            }
        } -OutputPath $outputPath

        $result.Success | Should -BeTrue
        $html = Get-Content -LiteralPath $outputPath -Raw

        $html | Should -Match '<tr><td>Entra AD joined</td><td>1</td>'
        $html | Should -Match '<tr><td>Hybrid Entra AD joined</td><td>1</td>'
        $html | Should -Match '<tr><td>Entra AD registered</td><td>1</td>'
        $html | Should -Not -Match '<tr><td>Unknown</td><td>3</td>'
        $html | Should -Match '(?s)<div class="kpi-title">Hybrid Joined</div>\s*<div class="kpi-value">1'
    }

    It 'distinguishes uncollected licensing, public folder, and device inventories from measured zero' {
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $absentPath = Join-Path $TestDrive 'presales-inventory-absent.html'
        $absent = New-TenantMigrationScopeHtmlReport -TenantStatsHash @{
            TenantInfo = [ordered]@{ DisplayName = 'Legacy Snapshot'; DefaultDomain = 'legacy.example' }
        } -OutputPath $absentPath
        $absent.Success | Should -BeTrue
        $absentHtml = Get-Content -LiteralPath $absentPath -Raw

        $absentHtml | Should -Match 'License inventory was not collected in this run'
        $absentHtml | Should -Match 'Public folder inventory was not collected in this run'
        $absentHtml | Should -Match 'Device inventory was not collected in this run'
        $absentHtml | Should -Match '(?s)<div class="kpi-card kpi-warning">\s*<div class="kpi-title">Total Devices</div>\s*<div class="kpi-value">Not collected'
        $absentHtml | Should -Not -Match '(?s)<div class="kpi-card kpi-success">\s*<div class="kpi-title">Public Folders</div>\s*<div class="kpi-value">0'

        # A collector that ran and found nothing is a different, reportable answer.
        $emptyPath = Join-Path $TestDrive 'presales-inventory-empty.html'
        $empty = New-TenantMigrationScopeHtmlReport -TenantStatsHash @{
            TenantInfo = [ordered]@{ DisplayName = 'Clean Tenant'; DefaultDomain = 'clean.example' }
            PublicFolderDetails = @{}
            DeviceDetails = @{}
            LicenseSKUs = @{}
        } -OutputPath $emptyPath
        $empty.Success | Should -BeTrue
        $emptyHtml = Get-Content -LiteralPath $emptyPath -Raw

        $emptyHtml | Should -Match 'No public folders were found'
        $emptyHtml | Should -Not -Match 'Public folder inventory was not collected in this run'
        $emptyHtml | Should -Not -Match 'Device inventory was not collected in this run'
        $emptyHtml | Should -Match 'The device collector ran and returned no device objects'
        $emptyHtml | Should -Match 'The license collector ran and returned no subscribed SKUs'
    }

    It 'presents Presales HTML as source-only effort scoping and prioritizes discovered blockers' {
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $outputPath = Join-Path $TestDrive 'presales-source-only.html'
        $result = New-TenantMigrationScopeHtmlReport -TenantStatsHash @{
            TenantInfo = @([pscustomobject]@{ DisplayName = 'Source Tenant'; DefaultDomain = 'source.example' })
            MigrationQuoteReadiness = @(
                [pscustomobject]@{ RowType = 'Summary'; ReadinessLevel = 'ROM'; Status = 'ROM'; DecisionKey = 'QR-01' }
                [pscustomobject]@{ RowType = 'Check'; Category = 'Target'; DecisionKey = 'ID-01'; Requirement = 'Generic destination confirmation'; Status = 'Blocker'; DiscoveredValue = 'Needs Input' }
                [pscustomobject]@{ RowType = 'Check'; Category = 'BitTitan'; DecisionKey = 'BT-02'; Requirement = 'Inactive mailbox restore plan'; Status = 'Blocker'; DiscoveredValue = 14 }
            )
            MigrationTargetReadiness = @(
                [pscustomobject]@{ Area = 'Destination tenant'; DecisionKey = 'ID-01'; ReadinessCheck = 'Customer confirmation'; Status = 'Needs Input' }
            )
            MigrationWorkloadEffort = @(
                [pscustomobject]@{
                    Workload = 'Mailboxes'; MigrationTool = 'BitTitan'; ObjectCount = 14; DataGB = 1.25
                    ComplexityDrivers = 'VERBOSE-DRIVER-SENTINEL'; ProvisionalEffortBand = 'Medium'
                    Status = 'Provisional'; Assumptions = 'VERBOSE-ASSUMPTION-SENTINEL'
                }
            )
            MigrationWavePlan = @(
                [pscustomobject]@{ Wave = '1 - Pilot'; CandidateObjectCount = 1; EstimatedWaveCount = 1; Status = 'Provisional' }
            )
        } -OutputPath $outputPath

        $result.Success | Should -BeTrue
        $html = Get-Content -LiteralPath $outputPath -Raw
        $html | Should -Match 'This is a source-only snapshot'
        $html | Should -Match 'customer-provided destination inputs, not queried'
        $html | Should -Match 'Customer-Provided Destination and Tool Inputs \(Not Queried\)'
        $html | Should -Match 'Provisional Workload Effort'
        $html | Should -Not -Match 'Provisional Migration Waves|Provisional Wave Units|Provisional Effort and Wave Plan'
        $html | Should -Match '(?s)Provisional Workload Effort.*<thead><tr><th>Workload</th><th>MigrationTool</th><th>ObjectCount</th><th>Measured Data</th><th>Effort Band</th><th>Status</th></tr></thead>'
        $html | Should -Not -Match 'VERBOSE-DRIVER-SENTINEL|VERBOSE-ASSUMPTION-SENTINEL|<th>ComplexityDrivers</th>|<th>Assumptions</th>'
        $html | Should -Match '(?s)Highest-Priority Unresolved Blockers.*<td>BT-02</td>.*<td>ID-01</td>'
    }

    It 'renders one practical BitTitan mix and paragraph-ready optional TMB guidance' {
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $outputPath = Join-Path $TestDrive 'presales-practical-license-mix.html'
        $result = New-TenantMigrationScopeHtmlReport -TenantStatsHash @{
            TenantInfo = @([pscustomobject]@{ DisplayName = 'License Mix Tenant'; DefaultDomain = 'mix.example' })
            BitTitanLicenseMixBreakdown = @(
                [pscustomobject]@{ Category = 'Office 365 Groups (report-only / optional)'; ObjectCount = 37; MailboxMigrationLicenseUnits = 0; UserMigrationBundleUnits = 0; ReportOnlyObjectCount = 37; OptionalMigrationObjectCount = 10; NeedsDataObjectCount = 0; DataGB = 0.01; UmbEligibilityValidationCount = 0; Treatment = 'Report only by default.' }
                [pscustomobject]@{ Category = 'Shared Mailboxes'; ObjectCount = 281; MailboxMigrationLicenseUnits = 280; UserMigrationBundleUnits = 1; ReportOnlyObjectCount = 0; OptionalMigrationObjectCount = 0; NeedsDataObjectCount = 0; DataGB = 1.2; UmbEligibilityValidationCount = 1; Treatment = 'Mailbox Migration or UMB.' }
                [pscustomobject]@{ Category = 'User Mailboxes + archive'; ObjectCount = 2; MailboxMigrationLicenseUnits = 0; UserMigrationBundleUnits = 2; ReportOnlyObjectCount = 0; OptionalMigrationObjectCount = 0; NeedsDataObjectCount = 0; DataGB = 0.1; UmbEligibilityValidationCount = 0; Treatment = 'UMB.' }
            )
            BitTitanLicenseSummary = @(
                [pscustomobject]@{ Section = 'Alternative scenarios'; Metric = 'Workload-fit: MigrationWiz-Mailbox units'; Value = 318; Notes = 'Workbook-only alternative.' }
                [pscustomobject]@{ Section = 'Alternative scenarios'; Metric = 'UMB-led: User Migration Bundles'; Value = 26; Notes = 'Workbook-only alternative.' }
                [pscustomobject]@{ Section = 'Totals'; Metric = 'License path selection'; Value = 'SE selection required'; Notes = 'Workbook-only decision.' }
                [pscustomobject]@{ Section = 'Policy'; Metric = 'Group mailbox licensing'; Value = 'Report only'; Notes = 'Visible policy.' }
            )
            BitTitanPlanningOptions = @(
                [pscustomobject]@{
                    Option = 'Tenant Migration Bundle'
                    GuidanceParagraph = 'Tenant Migration Bundle is optional and is not Arraya''s default. Private-chat guidance is changing; validate it before ordering.'
                    PrivateChatGuideUrl = 'https://help.bittitan.com/hc/en-us/articles/25603590557979-Teams-Private-Chat-Migration-Guide'
                    GettingStartedUrl = 'https://help.bittitan.com/hc/en-us/articles/1260806785629-Getting-Started-with-Migrations'
                }
            )
        } -OutputPath $outputPath

        $result.Success | Should -BeTrue
        $html = Get-Content -LiteralPath $outputPath -Raw
        $html | Should -Match 'Practical License Mix by Recipient Category'
        $html | Should -Match 'Mailbox Migration Licenses'
        $html | Should -Match 'User Migration Bundles'
        $html | Should -Match '<td>Office 365 Groups</td>'
        $html | Should -Not -Match '<td>Office 365 Groups \(report-only / optional\)</td>'
        $html | Should -Match 'Tenant Migration Bundle is optional and is not Arraya&#39;s default'
        $html | Should -Match 'Current Teams Private Chat guide'
        $html | Should -Not -Match 'Workload-Fit|UMB-Led|Path Selection|SE selection required'
        $html | Should -Match '<th>Category</th><th>Recipients</th><th>Mailbox Licenses</th><th>UMB</th><th>Data</th>'
        $html | Should -Not -Match '<th>Report-Only Objects</th>|<th>Optional Migration Objects</th>|<th>Needs Data</th>|<th>UMB Eligibility Checks</th>|<th>Treatment</th>'
    }

    It 'renders tenant and domain scope without mailbox or group object-detail tables' {
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop
        . (Join-Path $script:repoRoot 'src\scripts\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1')

        $outputPath = Join-Path $TestDrive 'presales-concise-scope.html'
        $result = New-TenantMigrationScopeHtmlReport -TenantStatsHash @{
            TenantInfo = [ordered]@{ DisplayName = 'Source Tenant'; DefaultDomain = 'contoso.com' }
            MigrationScopeSummary = @(
                [pscustomobject]@{ Tool = 'BitTitan'; Section = 'Data'; Metric = 'Total mailbox data to migrate (GB)'; Value = 2.5 }
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'Sites'; Metric = 'SharePoint sites (total)'; Value = 12 }
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'Data'; Metric = 'SharePoint storage (GB)'; Value = 4.25 }
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'OneDrive'; Metric = 'OneDrive sites'; Value = 8 }
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'OneDrive'; Metric = 'OneDrive storage (GB)'; Value = 1.5 }
                [pscustomobject]@{ Tool = 'ShareGate'; Section = 'Teams'; Metric = 'Teams'; Value = 3 }
            )
            Domains = @(
                [pscustomobject]@{ Domain = 'contoso.com'; DomainType = 'Authoritative'; Verified = $true; IsDefault = $true; DNSCompanies = 'Cloudflare'; NSRecords = 'ada.ns.cloudflare.com,bob.ns.cloudflare.com'; PrimarySMTPRecipients = 6; AliasOnlyRecipients = 2; TotalDomainRecipients = 8 }
                [pscustomobject]@{ Domain = 'contoso.onmicrosoft.com'; DomainType = 'Authoritative'; Verified = $true; IsDefault = $false; DNSCompanies = ''; NSRecords = ''; PrimarySMTPRecipients = 2; AliasOnlyRecipients = 4; TotalDomainRecipients = 6 }
            )
            BitTitanLicenseMixBreakdown = @(
                [pscustomobject]@{ Category = 'Office 365 Groups (report-only / optional)'; ObjectCount = 2; MailboxMigrationLicenseUnits = 0; UserMigrationBundleUnits = 0; ReportOnlyObjectCount = 2; OptionalMigrationObjectCount = 1; NeedsDataObjectCount = 0; DataGB = 0.1; UmbEligibilityValidationCount = 0; Treatment = 'Report only' }
                [pscustomobject]@{ Category = 'Shared Mailboxes'; ObjectCount = 2; MailboxMigrationLicenseUnits = 2; UserMigrationBundleUnits = 0; ReportOnlyObjectCount = 0; OptionalMigrationObjectCount = 0; NeedsDataObjectCount = 0; DataGB = 1.0; UmbEligibilityValidationCount = 0; Treatment = 'Migrate' }
                [pscustomobject]@{ Category = 'Inactive'; ObjectCount = 1; MailboxMigrationLicenseUnits = 1; UserMigrationBundleUnits = 0; ReportOnlyObjectCount = 0; OptionalMigrationObjectCount = 0; NeedsDataObjectCount = 0; DataGB = 0.5; UmbEligibilityValidationCount = 0; Treatment = 'Restore then migrate' }
                [pscustomobject]@{ Category = 'User Mailboxes + archive'; ObjectCount = 1; MailboxMigrationLicenseUnits = 0; UserMigrationBundleUnits = 1; ReportOnlyObjectCount = 0; OptionalMigrationObjectCount = 0; NeedsDataObjectCount = 0; DataGB = 0.9; UmbEligibilityValidationCount = 0; Treatment = 'Migrate' }
            )
            BitTitanLicenseDetail = @(
                [pscustomobject]@{ DisplayName = 'Mailbox Detail Must Stay Out'; IsInactiveMailbox = $true; SizeGB = 100; LicenseUnits = 2 }
            )
            GroupWorkloadReconciliation = @(
                [pscustomobject]@{ DisplayName = 'Group Detail Must Stay Out'; Classification = 'Mailbox and site content'; IsTeam = $true; MailboxScope = 'BitTitan (optional)'; SiteScope = 'ShareGate'; MailboxEvidenceStatus = 'Measured data' }
                [pscustomobject]@{ DisplayName = 'Second Group Must Stay Out'; Classification = 'Empty shell'; IsTeam = $false; MailboxScope = 'None'; SiteScope = 'None'; MailboxEvidenceStatus = 'Measured empty' }
            )
            ShareGateScopeSummary = @(
                [pscustomobject]@{ Workload = 'All collaboration sites (non-duplicated)'; ObjectCount = 20; TotalStorageGB = 12; LargestObjectName = 'Largest OneDrive'; LargestObjectSizeGB = 4 }
                [pscustomobject]@{ Workload = 'Teams-connected SharePoint'; ObjectCount = 3; TotalStorageGB = 2; LargestObjectName = 'Largest Team site'; LargestObjectSizeGB = 1 }
                [pscustomobject]@{ Workload = 'Microsoft 365 Group SharePoint (non-Team)'; ObjectCount = 2; TotalStorageGB = 1; LargestObjectName = 'Largest group site'; LargestObjectSizeGB = 0.75 }
                [pscustomobject]@{ Workload = 'Standalone SharePoint'; ObjectCount = 7; TotalStorageGB = 5; LargestObjectName = 'Largest standalone site'; LargestObjectSizeGB = 2 }
                [pscustomobject]@{ Workload = 'SharePoint (total)'; ObjectCount = 12; TotalStorageGB = 8; LargestObjectName = 'Largest standalone site'; LargestObjectSizeGB = 2 }
                [pscustomobject]@{ Workload = 'OneDrive'; ObjectCount = 8; TotalStorageGB = 4; LargestObjectName = 'Largest OneDrive'; LargestObjectSizeGB = 4 }
            )
        } -OutputPath $outputPath

        $result.Success | Should -BeTrue
        $html = Get-Content -LiteralPath $outputPath -Raw
        $html | Should -Match 'Tenant Scope at a Glance'
        $html | Should -Match '<title>Source Tenant - Migration Scoping Brief'
        $html | Should -Match 'Source Tenant</h1>'
        $html | Should -Match 'Mailbox Recipients'
        $html | Should -Not -Match 'Licensable Users'
        $html | Should -Match 'Domain and DNS Inventory'
        $html | Should -Match 'contoso\.com'
        $html | Should -Match 'Cloudflare'
        $html | Should -Match 'ada\.ns\.cloudflare\.com'
        $html | Should -Match 'Primary SMTP Recipients'
        $html | Should -Match 'Alias-Only Recipients'
        $html | Should -Not -Match 'Largest Mailboxes|Inactive Mailboxes: Restore and Archive Scope'
        $html | Should -Not -Match 'Mailbox Detail Must Stay Out|Group Detail Must Stay Out|Second Group Must Stay Out'
        $html | Should -Match 'Mailbox and site content'
        $html | Should -Match 'Empty shell'
        $html | Should -Match 'All collaboration sites \(non-duplicated\)'
        $html | Should -Match 'Teams-connected SharePoint'
        $html | Should -Match 'Microsoft 365 Group SharePoint \(non-Team\)'
        $html | Should -Match 'Total Collaboration Storage'
        $html | Should -Match 'How the storage total works'
        $html | Should -Match '<th>Workload</th><th>Sites</th><th>Storage</th><th>Largest Site</th><th>Largest Site Size</th>'
        $html | Should -Not -Match '<th>UnknownStorageCount</th>|<th>Notes</th>'
    }
}
