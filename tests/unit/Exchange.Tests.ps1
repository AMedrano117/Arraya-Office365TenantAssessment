Describe 'Arraya.M365.Exchange' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Arraya.M365.Exchange.psd1'
        $script:moduleRoot = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange'
        $script:hybridPath = Join-Path $script:moduleRoot 'Public\Get-ExchangeHybridConfiguration.ps1'
        $script:hybridSource = Get-Content -Raw -Path $script:hybridPath
        $script:mailboxPath = Join-Path $script:moduleRoot 'Public\Get-AllExchangeMailboxDetails.ps1'
        $script:mailboxSource = Get-Content -Raw -Path $script:mailboxPath
        $script:mailboxStatHelperPath = Join-Path $script:moduleRoot 'Private\Invoke-ExchangeMailboxStatHelpers.ps1'
        $script:mailboxStatHelperSource = Get-Content -Raw -Path $script:mailboxStatHelperPath
        $script:smtpRelayPath = Join-Path $script:moduleRoot 'Public\Get-SMTPRelayConfiguration.ps1'
        $script:smtpRelaySource = Get-Content -Raw -Path $script:smtpRelayPath
        $script:mailFlowPath = Join-Path $script:moduleRoot 'Public\Get-MailFlowRulesandConnectors.ps1'
        $script:mailFlowSource = Get-Content -Raw -Path $script:mailFlowPath
        $script:expectedExports = @(
            'Get-AllExchangeMailboxDetails'
            'Get-AllPublicFolderDetails'
            'Get-AllRecipientDetails'
            'Get-ExchangeGroupDetails'
            'Get-ExchangeHybridConfiguration'
            'Get-MailFlowRulesandConnectors'
            'Get-SMTPRelayConfiguration'
            'Get-ThirdPartySpamFilteringConfig'
        )
    }

    It 'has a manifest' {
        Test-Path $script:manifestPath | Should -BeTrue
    }

    It 'exports the reduced exchange collector surface' {
        $manifest = Import-PowerShellDataFile -Path $script:manifestPath
        $manifest.FunctionsToExport | Should -Be $script:expectedExports
        $manifest.VariablesToExport | Should -Be @()
    }

    It 'does not keep retired public exchange wrapper files' {
        @(
            'Public\Get-ArrayaExchangeRuntimeContext.ps1'
            'Public\Set-ArrayaExchangeRuntimeContext.ps1'
            'Public\Get-ResolvedEmailAddresses.ps1'
            'Public\Get-ExoMailboxStatisticsSafe.ps1'
            'Public\Resolve-ExoStatisticsIdentity.ps1'
            'Public\Invoke-ExchangeMailboxStatHelpers.ps1'
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
                'Get-ArrayaExchangeRuntimeContext'
                'Set-ArrayaExchangeRuntimeContext'
                'Get-ArrayaGraphAdminReportSettings'
                'Export-ArrayaGraphReportCsv'
                'Invoke-QuietRestMethod'
            ) -CaseSensitive
        $matches | Should -BeNullOrEmpty
    }

    It 'uses safe connector identity fallback logic for hybrid detection and suppresses noisy fallback warnings' {
        Test-Path $script:hybridPath | Should -BeTrue
        $script:hybridSource | Should -Match 'function Get-ConnectorIdentityValue'
        $script:hybridSource | Should -Match "\('Identity', 'Guid', 'Name', 'Id'\)"
        $script:hybridSource | Should -Not -Match 'Select-Object -ExpandProperty ID'
        $script:hybridSource | Should -Match 'MailFlowConnectors'
        $script:hybridSource | Should -Match 'Test-mode connectors may not be included in this fallback view'
        $script:hybridSource | Should -Match 'Get-InboundConnector -ErrorAction SilentlyContinue -WarningAction SilentlyContinue'
        $script:hybridSource | Should -Match 'Get-OutboundConnector -ErrorAction SilentlyContinue -WarningAction SilentlyContinue'
    }

    It 'stamps migration-ready mailbox address properties for cutover planning' {
        Test-Path $script:mailboxPath | Should -BeTrue
        $script:mailboxSource | Should -Match 'function Add-MailboxMigrationAddressProperties'
        $script:mailboxSource | Should -Match 'function Convert-MailboxProxyAddressCollectionToArray'
        $script:mailboxSource | Should -Match '"LegacyExchangeDN"'
        $script:mailboxSource | Should -Match 'LegacyExchangeDn'
        $script:mailboxSource | Should -Match 'LegacyExchangeDnX500'
        $script:mailboxSource | Should -Match 'OnMicrosoftAlias'
        $script:mailboxSource | Should -Match 'OnMicrosoftAliases'
        $script:mailboxSource | Should -Match 'OnMicrosoftAliasCount'
        $script:mailboxSource | Should -Match 'X500Addresses'
        $script:mailboxSource | Should -Match 'X500AddressCount'
        $script:mailboxSource | Should -Match 'X400Addresses'
        $script:mailboxSource | Should -Match 'X400AddressCount'
        $script:mailboxSource | Should -Match 'GrantSendOnBehalfToCount'
        $script:mailboxSource | Should -Match 'FullAccessDelegates'
        $script:mailboxSource | Should -Match 'FullAccessDelegateCount'
        $script:mailboxSource | Should -Match 'FullAccessDelegateState'
        $script:mailboxSource | Should -Match 'SendAsDelegates'
        $script:mailboxSource | Should -Match 'SendAsDelegateCount'
        $script:mailboxSource | Should -Match 'SendAsDelegateState'
        $script:mailboxSource | Should -Match 'CalendarDelegates'
        $script:mailboxSource | Should -Match 'CalendarDelegateCount'
        $script:mailboxSource | Should -Match 'CalendarPermissionEntryCount'
        $script:mailboxSource | Should -Match 'CalendarDelegateState'
        $script:mailboxSource | Should -Match 'Add-MailboxMigrationAddressProperties -Mailbox \$mailbox'
        $script:mailboxSource | Should -Match 'function Update-MailboxDelegatePermissionInventory'
        $script:mailboxSource | Should -Match 'function Update-MailboxCalendarDelegateInventory'
        $script:mailboxSource | Should -Match 'function Resolve-MailboxCalendarQueryIdentity'
        $script:mailboxSource | Should -Match 'Get-EXOMailboxPermission'
        $script:mailboxSource | Should -Match "Parameters\.ContainsKey\('GroupMailbox'\)"
        $script:mailboxSource | Should -Match "RecipientTypeDetails -eq 'GroupMailbox'"
        $script:mailboxSource | Should -Match 'Get-EXORecipientPermission -AccessRights SendAs -ResultSize Unlimited'
        $script:mailboxSource | Should -Match 'Get-EXOMailboxFolderStatistics -Identity \$mailboxCalendarQueryIdentity -FolderScope Calendar'
        $script:mailboxSource | Should -Match 'Get-EXOMailboxFolderPermission \$calendarPermissionIdentity'
        $script:mailboxSource | Should -Match 'Get-MailboxPermission -Identity \$mailbox\.Identity'
        $script:mailboxSource | Should -Match 'Get-RecipientPermission -Identity \$mailbox\.Identity'
        $script:mailboxSource | Should -Match 'Get-MailboxFolderStatistics -Identity \$mailboxCalendarQueryIdentity'
        $script:mailboxSource | Should -Match 'Get-MailboxFolderPermission \$calendarPermissionIdentity'
        $script:mailboxSource | Should -Match 'Gathering mailbox delegate permissions'
        $script:mailboxSource | Should -Match 'Gathering mailbox calendar delegate permissions'
        $script:mailboxSource | Should -Match 'Enumerating permissions:'
        $script:mailboxSource | Should -Match 'Calendar delegate enumeration for mailbox'
        $script:mailboxSource | Should -Match 'MailboxCalendarDelegatePermissions'
        $script:mailboxSource | Should -Match 'GrantSendOnBehalfTo'
        $script:mailboxSource | Should -Match '\^\(\?i\)x500:'
        $script:mailboxSource | Should -Match '\^\(\?i\)x400:'
    }

    It 'reuses run-scoped collector cache for expensive Exchange activity reports' {
        Test-Path $script:mailboxPath | Should -BeTrue
        Test-Path $script:mailboxStatHelperPath | Should -BeTrue
        $script:mailboxSource | Should -Match 'GraphActivityReport:MailboxUsage:D180'
        $script:mailboxSource | Should -Match 'Get-ArrayaCollectorCacheValue -Context \$Context -Key \$graphReportCacheKey'
        $script:mailboxSource | Should -Match 'Set-ArrayaCollectorCacheValue -Context \$Context -Key \$graphReportCacheKey -Value \$graphReportData'
        $script:mailboxStatHelperSource | Should -Match 'GraphActivityReport:Office365GroupsActivity:D180'
        $script:mailboxStatHelperSource | Should -Match 'Get-ArrayaCollectorCacheValue -Context \$Context -Key \$reportCacheKey'
        $script:mailboxStatHelperSource | Should -Match 'Set-ArrayaCollectorCacheValue -Context \$Context -Key \$reportCacheKey -Value \$rows'
    }

    It 'collects and retains SMTP AUTH and inbound relay authentication evidence for migration profiles' {
        $script:mailboxSource | Should -Match "detailLevel -in @\('all', 'geek'\)"
        $script:mailboxSource | Should -Match 'Get-EXOCASMailbox -ResultSize Unlimited -Properties SmtpClientAuthenticationDisabled'
        $script:mailboxSource | Should -Match 'Add-ArrayaMailboxSmtpAuthSetting -Mailboxes \$exoMailboxes'
        $script:mailFlowSource | Should -Match "'TlsSenderCertificateName'"
        $script:mailFlowSource | Should -Match "'RestrictDomainsToCertificate'"
        $script:mailFlowSource | Should -Match "'RestrictDomainsToIPAddresses'"
        $script:smtpRelaySource | Should -Match "ConnectorType -ine 'OnPremises'"
        $script:smtpRelaySource | Should -Match "ConnectorDirection', 'Direction'"
    }

    Context 'SMTP relay evidence classification' {
        BeforeAll {
            Import-Module -Name $script:manifestPath -Force -ErrorAction Stop
            $script:newSmtpRelayTestContext = {
                param(
                    [System.Collections.IDictionary]$TenantStats
                )

                [pscustomobject]@{
                    TenantStats        = $TenantStats
                    ExportFileLocation = $null
                    Policies           = [ordered]@{}
                    Runtime            = [ordered]@{}
                    Metadata           = [ordered]@{ StartedAt = Get-Date }
                }
            }
        }

        BeforeEach {
            Mock -CommandName Get-ArrayaSmtpRelayTransportConfig -ModuleName Arraya.M365.Exchange -MockWith {
                [pscustomobject]@{ SmtpClientAuthenticationDisabled = $true }
            }
            Mock -CommandName Get-ArrayaSmtpRelayAcceptedDomain -ModuleName Arraya.M365.Exchange -MockWith { @() }
        }

        It 'projects CAS mailbox settings onto matching mailbox inventory records' {
            $mailboxes = @(
                [pscustomobject]@{ ExternalDirectoryObjectId = 'object-1'; UserPrincipalName = 'relay@contoso.com' },
                [pscustomobject]@{ ExternalDirectoryObjectId = 'object-2'; UserPrincipalName = 'unmatched@contoso.com' }
            )
            $casSettings = @(
                [pscustomobject]@{
                    ExternalDirectoryObjectId          = 'object-1'
                    UserPrincipalName                  = 'relay@contoso.com'
                    SmtpClientAuthenticationDisabled  = $false
                }
            )

            $module = Get-Module -Name Arraya.M365.Exchange
            $projection = & $module {
                param($MailboxRows, $SettingRows)
                Add-ArrayaMailboxSmtpAuthSetting -Mailboxes $MailboxRows -CasMailboxSettings $SettingRows
            } $mailboxes $casSettings

            $projection.MatchedMailboxCount | Should -Be 1
            $mailboxes[0].SmtpClientAuthenticationDisabled | Should -BeFalse
            $mailboxes[1].PSObject.Properties['SmtpClientAuthenticationDisabled'] | Should -BeNullOrEmpty
        }

        It 'resolves explicit and inherited effective SMTP AUTH settings into service-account evidence' {
            Mock -CommandName Get-ArrayaSmtpRelayTransportConfig -ModuleName Arraya.M365.Exchange -MockWith {
                [pscustomobject]@{ SmtpClientAuthenticationDisabled = $false }
            }
            $tenantStats = [ordered]@{
                AllMailboxes = [ordered]@{
                    explicit = [pscustomobject]@{
                        DisplayName                      = 'Explicit Relay'
                        UserPrincipalName                = 'explicit-relay@contoso.com'
                        PrimarySmtpAddress               = 'explicit-relay@contoso.com'
                        SmtpClientAuthenticationDisabled = $false
                    }
                    inherited = [pscustomobject]@{
                        DisplayName                      = 'Inherited Relay'
                        UserPrincipalName                = 'inherited-relay@contoso.com'
                        PrimarySmtpAddress               = 'inherited-relay@contoso.com'
                        SmtpClientAuthenticationDisabled = $null
                    }
                }
                MailFlowConnectors = [ordered]@{}
            }
            $context = & $script:newSmtpRelayTestContext $tenantStats

            Get-SMTPRelayConfiguration -Context $context

            $config = $tenantStats['SMTPRelayConfig']['Configuration']
            $config.SMTPAuthEnabled | Should -BeTrue
            $config.SMTPAuthUsers | Should -Be 2
            $config.SMTPAuthExplicitlyEnabledUsers | Should -Be 1
            $config.SMTPAuthInheritedEnabledUsers | Should -Be 1
            $config.SMTPAuthEvidenceState | Should -Be 'Complete'
            $tenantStats['SMTPRelayServiceAccounts'].Count | Should -Be 2
        }

        It 'keeps a missing mailbox SMTP AUTH property unknown instead of treating it as enabled' {
            Mock -CommandName Get-ArrayaSmtpRelayTransportConfig -ModuleName Arraya.M365.Exchange -MockWith {
                [pscustomobject]@{ SmtpClientAuthenticationDisabled = $false }
            }
            $tenantStats = [ordered]@{
                AllMailboxes = [ordered]@{
                    missing = [pscustomobject]@{
                        DisplayName       = 'Uncollected Setting'
                        UserPrincipalName = 'unknown@contoso.com'
                    }
                }
                MailFlowConnectors = [ordered]@{}
            }
            $context = & $script:newSmtpRelayTestContext $tenantStats

            Get-SMTPRelayConfiguration -Context $context

            $config = $tenantStats['SMTPRelayConfig']['Configuration']
            $config.SMTPAuthEnabled | Should -BeNullOrEmpty
            $config.SMTPAuthUsers | Should -Be 0
            $config.SMTPAuthMailboxSettingsMissing | Should -Be 1
            $config.SMTPAuthMailboxStateUnknown | Should -Be 1
            $config.SMTPAuthUsersIsMinimum | Should -BeTrue
            $config.SMTPAuthEvidenceState | Should -Be 'Partial'
            $tenantStats['SMTPRelayServiceAccounts'].Count | Should -Be 0
        }

        It 'detects enabled OnPremises inbound relay connectors authenticated by IP or certificate' {
            $tenantStats = [ordered]@{
                AllMailboxes = [ordered]@{}
                MailFlowConnectors = [ordered]@{
                    ipRelay = [pscustomobject]@{
                        Id                    = 'IP relay'
                        ConnectorDirection    = 'Inbound'
                        ConnectorType         = 'OnPremises'
                        Enabled               = $true
                        TestMode              = $false
                        SenderIPAddresses     = @('203.0.113.10')
                    }
                    certificateRelay = [pscustomobject]@{
                        Id                       = 'Certificate relay'
                        ConnectorDirection       = 'Inbound'
                        ConnectorType            = 'OnPremises'
                        Enabled                  = $true
                        TestMode                 = $false
                        TlsSenderCertificateName = 'smtp.contoso.com'
                    }
                }
            }
            $context = & $script:newSmtpRelayTestContext $tenantStats

            Get-SMTPRelayConfiguration -Context $context

            $config = $tenantStats['SMTPRelayConfig']['Configuration']
            $config.ConnectorBasedRelay | Should -BeTrue
            $config.RelayConnectorEvidenceState | Should -Be 'Complete'
            @($config.RelayConnectors).Count | Should -Be 2
            @($config.RelayConnectors.ConnectorName) | Should -Contain 'IP relay'
            @($config.RelayConnectors.ConnectorName) | Should -Contain 'Certificate relay'
        }

        It 'does not classify outbound, partner, disabled, test-mode, or unauthenticated connectors as relay evidence' {
            $tenantStats = [ordered]@{
                AllMailboxes = [ordered]@{}
                MailFlowConnectors = [ordered]@{
                    outbound = [pscustomobject]@{ Id = 'Outbound'; ConnectorDirection = 'Outbound'; ConnectorType = 'OnPremises'; Enabled = $true; SenderIPAddresses = '203.0.113.11' }
                    partner = [pscustomobject]@{ Id = 'Partner'; ConnectorDirection = 'Inbound'; ConnectorType = 'Partner'; Enabled = $true; SenderIPAddresses = '203.0.113.12' }
                    disabled = [pscustomobject]@{ Id = 'Disabled'; ConnectorDirection = 'Inbound'; ConnectorType = 'OnPremises'; Enabled = $false; SenderIPAddresses = '203.0.113.13' }
                    testMode = [pscustomobject]@{ Id = 'Test'; ConnectorDirection = 'Inbound'; ConnectorType = 'OnPremises'; Enabled = $true; TestMode = $true; SenderIPAddresses = '203.0.113.14' }
                    noAuthentication = [pscustomobject]@{ Id = 'No auth'; ConnectorDirection = 'Inbound'; ConnectorType = 'OnPremises'; Enabled = $true; SenderDomains = '*' }
                }
            }
            $context = & $script:newSmtpRelayTestContext $tenantStats

            Get-SMTPRelayConfiguration -Context $context

            $config = $tenantStats['SMTPRelayConfig']['Configuration']
            $config.ConnectorBasedRelay | Should -BeFalse
            @($config.RelayConnectors).Count | Should -Be 0
        }
    }
}
