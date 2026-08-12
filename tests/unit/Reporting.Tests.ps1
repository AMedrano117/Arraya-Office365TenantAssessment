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
