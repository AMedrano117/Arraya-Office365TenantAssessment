function Export-TenantToTenantQuestionnaireMarkdown {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory)]
        [string]$TemplatePath,
        [Parameter(Mandatory)]
        [string]$Path,
        [string]$PreparedBy = "$($env:USERNAME) (automated assessment)"
    )

    if (-not (Test-Path -Path $TemplatePath)) {
        throw "Questionnaire template not found: $TemplatePath"
    }

    function Arr([string]$Key) {
        if (-not $TenantStatsHash.ContainsKey($Key) -or $null -eq $TenantStatsHash[$Key]) { return @() }
        $value = $TenantStatsHash[$Key]
        if ($value -is [hashtable] -or $value -is [System.Collections.Specialized.OrderedDictionary]) { return @($value.Values) }
        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) { return @($value) }
        return @($value)
    }

    function Val([string]$Key) {
        if ($TenantStatsHash.ContainsKey($Key)) { return $TenantStatsHash[$Key] }
        return $null
    }

    function Prop($Object, [string]$Name, $Default = $null) {
        if ($null -eq $Object) { return $Default }
        if ($Object -is [hashtable] -and $Object.ContainsKey($Name)) { return $Object[$Name] }
        $property = $Object.PSObject.Properties[$Name]
        if ($property) { return $property.Value }
        return $Default
    }

    function Txt($Value, [string]$Default = 'Not collected by current script; manual validation required.') {
        if ($null -eq $Value) { return $Default }
        if ($Value -is [bool]) { return $(if ($Value) { 'Yes' } else { 'No' }) }
        if ($Value -is [datetime] -or $Value -is [datetimeoffset]) { return ([datetime]$Value).ToString('yyyy-MM-dd') }
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) { return $Default }
        return $text.Trim()
    }

    function Truthy($Value) {
        if ($Value -is [bool]) { return $Value }
        $text = Txt $Value ''
        return $text -match '^(?i:true|yes|enabled|configured|present)$'
    }

    function QKey([string]$Text) {
        $normalized = $Text.ToLowerInvariant()
        $normalized = $normalized -replace '[\u2010-\u2015]', '-'
        $normalized = $normalized -replace '[^a-z0-9]+', '-'
        return $normalized.Trim('-')
    }

    function Cell([string]$Value) {
        $text = Txt $Value
        $text = $text -replace '\|', '\|'
        $text = $text -replace "`r?`n", ' '
        return $text.Trim()
    }

    function Summary([array]$Items, [string]$PropertyName, [int]$MaxItems = 8) {
        $values = @($Items | ForEach-Object { Prop $_ $PropertyName } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
        if ($values.Count -eq 0) { return $null }
        if ($values.Count -le $MaxItems) { return ($values -join ', ') }
        return (($values | Select-Object -First $MaxItems) -join ', ') + " (+$($values.Count - $MaxItems) more)"
    }

    function SizeGb($Object) {
        foreach ($name in @('TotalItemSizeGB','StorageUsedGB','StorageGB','SizeGB')) {
            $value = Prop $Object $name
            if ($null -ne $value) {
                $number = 0.0
                if ([double]::TryParse([string]$value, [ref]$number)) { return [math]::Round($number, 2) }
            }
        }
        $sizeText = Txt (Prop $Object 'TotalItemSize') ''
        if ($sizeText -match '([\d\.]+)\s*GB') { return [math]::Round([double]$Matches[1], 2) }
        return $null
    }

    function LicenseMatches([string[]]$Patterns) {
        return @(
            $licenses |
                Where-Object {
                    $haystack = @((Prop $_ 'SkuFriendlyName'), (Prop $_ 'SkuPartNumber'), (Prop $_ 'ServicePlans')) -join ' '
                    foreach ($pattern in $Patterns) { if ($haystack -match $pattern) { return $true } }
                    return $false
                } |
                Sort-Object { Txt (Prop $_ 'SkuFriendlyName') } -Unique
        )
    }

    function LicenseAnswer([string]$CapabilityName, [string[]]$Patterns) {
        $matches = LicenseMatches $Patterns
        if ($matches.Count -eq 0) { return 'Not detected in discovered license inventory.' }
        $sample = @(
            $matches | Select-Object -First 4 | ForEach-Object {
                "$(Txt (Prop $_ 'SkuFriendlyName') (Prop $_ 'SkuPartNumber')) ($(Txt (Prop $_ 'ConsumedUnits') '0')/$(Txt (Prop $_ 'PurchasedUnits') '0') consumed)"
            }
        ) -join '; '
        return "Likely yes for $CapabilityName; matched license(s): $sample"
    }

    function StorageAnswer([array]$Items, [string]$Label) {
        if ($Items.Count -eq 0) { return "No $Label inventory collected." }
        $sizes = @($Items | ForEach-Object { SizeGb $_ } | Where-Object { $null -ne $_ })
        if ($sizes.Count -eq 0) { return "$($Items.Count) $Label site(s) collected, but size data was unavailable." }
        return "$([math]::Round((($sizes | Measure-Object -Sum).Sum), 2)) GB across $($Items.Count) $Label site(s)."
    }

    $licenses = Arr 'LicenseSKUs'
    $users = Arr 'Users'
    $mailboxes = Arr 'MailboxFullDetails'
    if ($mailboxes.Count -eq 0) { $mailboxes = Arr 'AllMailboxes' }
    $inactiveMailboxes = Arr 'InactiveMailboxDetails'
    if ($inactiveMailboxes.Count -eq 0) { $inactiveMailboxes = @($mailboxes | Where-Object { (Prop $_ 'IsInactiveMailbox') -eq $true }) }
    $primaryMailboxStats = Arr 'PrimaryMailboxStats'
    $recipients = Arr 'AllRecipients'
    $publicFolders = Arr 'PublicFolderDetails'
    if ($publicFolders.Count -eq 0) { $publicFolders = Arr 'PublicFolders' }
    $sharePointSites = Arr 'SharePoint'
    $oneDriveSites = Arr 'OneDrive'
    $domains = Arr 'Domains'
    $devices = Arr 'DeviceDetails'
    $teams = Arr 'AllTeams'
    $groups = Arr 'EntraIDGroups'
    $conditionalAccessPolicies = Arr 'ConditionalAccessPolicies'
    $mailFlowRules = Arr 'MailFlowRules'
    $mailFlowConnectors = Arr 'MailFlowConnectors'
    $remoteDomains = Arr 'RemoteDomains'
    $migrationReadiness = Arr 'MigrationReadiness'
    $bestPracticeFindings = Arr 'BestPracticeFindings'

    $tenantInfo = Val 'TenantInfo'
    $authContainer = Val 'AuthenticationConfig'
    $authConfig = if ($authContainer -is [hashtable]) { $authContainer['Configuration'] } else { Prop $authContainer 'Configuration' }
    $adConnect = Val 'AdConnectConfiguration'
    $adConnectSummary = if ($adConnect -is [hashtable]) { $adConnect['Summary'] } else { Prop $adConnect 'Summary' }
    $hybridContainer = Val 'HybridConfiguration'
    $hybridInfo = if ($hybridContainer -is [hashtable]) { $hybridContainer['ExchangeHybrid'] } else { Prop $hybridContainer 'ExchangeHybrid' }
    $spamFiltering = Val 'SpamFilteringConfig'
    $spamFilteringSummary = if ($spamFiltering -is [hashtable]) { $spamFiltering['Configuration'] } else { Prop $spamFiltering 'Configuration' }
    $federation = Val 'FederationConfiguration'
    $crossTenantAccess = if ($federation -is [hashtable]) { $federation['CrossTenantAccess'] } else { Prop $federation 'CrossTenantAccess' }
    $mfaRegistrationSummary = Val 'MfaRegistrationSummary'
    $teamsVoice = Val 'TeamsVoice'
    $teamsVoiceSummary = if ($teamsVoice -is [hashtable]) { $teamsVoice['Summary'] } else { Prop $teamsVoice 'Summary' }

    $tenantDisplayName = Txt (Prop $tenantInfo 'DisplayName') 'Unknown tenant'
    $tenantDefaultDomain = Txt (Prop $tenantInfo 'DefaultDomain') ''
    $tenantInitialDomain = Txt (Prop $tenantInfo 'InitialDomain') ''
    $tenantCountryCode = Txt (Prop $tenantInfo 'CountryLetterCode') ''
    $tenantCountry = Txt (Prop $tenantInfo 'Country') ''
    $tenantPreferredDataLocation = Txt (Prop $tenantInfo 'PreferredDataLocation') ''

    $licenseSummary = if ($licenses.Count -gt 0) {
        @($licenses | Sort-Object { Prop $_ 'SkuFriendlyName' } | ForEach-Object { "$(Txt (Prop $_ 'SkuFriendlyName') (Prop $_ 'SkuPartNumber')): $(Txt (Prop $_ 'ConsumedUnits') '0')/$(Txt (Prop $_ 'PurchasedUnits') '0')" }) -join '; '
    } else {
        'No license inventory collected.'
    }

    $customDomains = @($domains | Where-Object { $domainName = Txt (Prop $_ 'Domain') ''; -not [string]::IsNullOrWhiteSpace($domainName) -and $domainName -notlike '*.onmicrosoft.com' })
    $m365GroupCount = @($groups | Where-Object { (Prop $_ 'GroupType') -eq 'Microsoft 365' }).Count
    $guestCount = @($users | Where-Object { (Txt (Prop $_ 'UserType') '') -match 'Guest' -or (Txt (Prop $_ 'UserPrincipalName') '') -like '*#EXT#*' }).Count
    $managedDevices = @($devices | Where-Object { Truthy (Prop $_ 'IsManaged') }).Count
    $intuneOrSccmDevices = @($devices | Where-Object { (Txt (Prop $_ 'MDMSolution') '') -match 'Intune|SCCM' }).Count
    $windowsDevices = @($devices | Where-Object { (Txt (Prop $_ 'OperatingSystem') '') -match '^Windows' }).Count
    $macDevices = @($devices | Where-Object { (Txt (Prop $_ 'OperatingSystem') '') -match 'Mac' }).Count
    $iosDevices = @($devices | Where-Object { (Txt (Prop $_ 'OperatingSystem') '') -match '^iOS' }).Count
    $androidDevices = @($devices | Where-Object { (Txt (Prop $_ 'OperatingSystem') '') -match 'Android' }).Count
    $mobileDeviceCount = $iosDevices + $androidDevices
    $trustTypeSummary = Summary $devices 'TrustType' 6
    $mdmSolutionSummary = Summary $devices 'MDMSolution' 4

    $mailboxSizeRows = @(
        $primaryMailboxStats |
            ForEach-Object {
                $size = SizeGb $_
                if ($null -ne $size) {
                    [PSCustomObject]@{
                        DisplayName = Txt (Prop $_ 'DisplayName') (Prop $_ 'MailboxGuid')
                        SizeGB      = $size
                    }
                }
            } |
            Where-Object { $null -ne $_ }
    )
    $averageMailboxSizeGb = if ($mailboxSizeRows.Count -gt 0) { [math]::Round((($mailboxSizeRows | Measure-Object SizeGB -Average).Average), 2) } else { $null }
    $largestMailbox = @($mailboxSizeRows | Sort-Object SizeGB -Descending | Select-Object -First 1)
    $crossTenantMailboxRows = @(
        $mailboxes | Where-Object {
            (Txt (Prop $_ 'RecipientTypeDetails') '') -ne 'GroupMailbox'
        }
    )
    $mailboxesWithOnMicrosoftAlias = @(
        $crossTenantMailboxRows | Where-Object {
            -not [string]::IsNullOrWhiteSpace((Txt (Prop $_ 'OnMicrosoftAlias') ''))
        }
    ).Count
    $mailboxesWithLegacyExchangeDn = @(
        $crossTenantMailboxRows | Where-Object {
            -not [string]::IsNullOrWhiteSpace((Txt (Prop $_ 'LegacyExchangeDn') ''))
        }
    ).Count
    $mailboxesWithExistingX500 = @(
        $crossTenantMailboxRows | Where-Object {
            $countValue = 0
            [void][int]::TryParse((Txt (Prop $_ 'X500AddressCount') '0'), [ref]$countValue)
            $countValue -gt 0
        }
    ).Count
    $mailboxesWithExistingX400 = @(
        $crossTenantMailboxRows | Where-Object {
            $countValue = 0
            [void][int]::TryParse((Txt (Prop $_ 'X400AddressCount') '0'), [ref]$countValue)
            $countValue -gt 0
        }
    ).Count
    $calendarDelegateSignalMailboxCount = @(
        $crossTenantMailboxRows | Where-Object {
            (Txt (Prop $_ 'CalendarDelegateState') '') -in @('Collected', 'Partial')
        }
    ).Count
    $mailboxesWithCalendarDelegates = @(
        $crossTenantMailboxRows | Where-Object {
            $countValue = 0
            [void][int]::TryParse((Txt (Prop $_ 'CalendarDelegateCount') '0'), [ref]$countValue)
            $countValue -gt 0
        }
    ).Count
    $fullAccessSignalMailboxCount = @(
        $crossTenantMailboxRows | Where-Object {
            (Txt (Prop $_ 'FullAccessDelegateState') '') -eq 'Collected'
        }
    ).Count
    $mailboxesWithFullAccessDelegates = @(
        $crossTenantMailboxRows | Where-Object {
            $countValue = 0
            [void][int]::TryParse((Txt (Prop $_ 'FullAccessDelegateCount') '0'), [ref]$countValue)
            $countValue -gt 0
        }
    ).Count
    $sendAsSignalMailboxCount = @(
        $crossTenantMailboxRows | Where-Object {
            (Txt (Prop $_ 'SendAsDelegateState') '') -eq 'Collected'
        }
    ).Count
    $mailboxesWithSendAsDelegates = @(
        $crossTenantMailboxRows | Where-Object {
            $countValue = 0
            [void][int]::TryParse((Txt (Prop $_ 'SendAsDelegateCount') '0'), [ref]$countValue)
            $countValue -gt 0
        }
    ).Count
    $sendOnBehalfSignalMailboxCount = @(
        $crossTenantMailboxRows | Where-Object {
            $null -ne (Prop $_ 'GrantSendOnBehalfToCount')
        }
    ).Count
    $mailboxesWithSendOnBehalfDelegates = @(
        $crossTenantMailboxRows | Where-Object {
            $countValue = 0
            [void][int]::TryParse((Txt (Prop $_ 'GrantSendOnBehalfToCount') '0'), [ref]$countValue)
            $countValue -gt 0
        }
    ).Count

    $connectorNames = Summary $mailFlowConnectors 'Name' 5
    $inboundConnectors = @($mailFlowConnectors | Where-Object { (Prop $_ 'ConnectorDirection') -eq 'Inbound' }).Count
    $outboundConnectors = @($mailFlowConnectors | Where-Object { (Prop $_ 'ConnectorDirection') -eq 'Outbound' }).Count
    $mailConnectorSummary = if ($mailFlowConnectors.Count -gt 0) {
        "$inboundConnectors inbound and $outboundConnectors outbound connector(s)" + $(if ($connectorNames) { "; names: $connectorNames" } else { '' })
    } else {
        'No send/receive connectors collected.'
    }

    $usesThirdPartyFiltering = Truthy (Prop $spamFilteringSummary 'Uses3rdPartyFiltering')
    $selfServicePurchase = Prop $tenantInfo 'SelfServicePurchase'
    $azureResourceUsage = Prop $tenantInfo 'AzureResourceUsage'
    $ssoApps = Prop $authConfig 'SSOApplications'
    $federatedDomains = Prop $authConfig 'FederatedDomains'
    $mfaMethods = Prop $authConfig 'MFAMethods'
    $partnerCount = 0
    [void][int]::TryParse([string](Prop $crossTenantAccess 'PartnerCount' 0), [ref]$partnerCount)

    $comments = New-Object System.Collections.Generic.List[string]
    $comments.Add("- Automated tenant discovery generated on $(Get-Date -Format 'yyyy-MM-dd') for $tenantDisplayName.") | Out-Null
    $priorityMigration = @($migrationReadiness | Where-Object { (Prop $_ 'Status') -in @('Blocker','Review','Needs Data') } | Select-Object -First 6)
    if ($priorityMigration.Count -gt 0) {
        $summary = @($priorityMigration | ForEach-Object { "$(Txt (Prop $_ 'Item') 'Unknown item') ($(Txt (Prop $_ 'Status') 'Review'))" }) -join '; '
        $comments.Add("- Primary migration review items: $summary.") | Out-Null
    }
    $priorityFindings = @($bestPracticeFindings | Where-Object { (Prop $_ 'Severity') -in @('Risk','Warning') } | Select-Object -First 4)
    if ($priorityFindings.Count -gt 0) {
        $summary = @($priorityFindings | ForEach-Object { Txt (Prop $_ 'Message') '' }) -join ' '
        $comments.Add("- Current assessment findings to review: $summary") | Out-Null
    }
    if ($crossTenantMailboxRows.Count -gt 0) {
        $mailboxCutoverComment = "- Exchange cutover-ready mailbox data: $mailboxesWithOnMicrosoftAlias/$($crossTenantMailboxRows.Count) include a source .onmicrosoft alias, $mailboxesWithLegacyExchangeDn/$($crossTenantMailboxRows.Count) include LegacyExchangeDN, X500 proxies are surfaced on $mailboxesWithExistingX500 mailbox(es), and X400 proxies are surfaced on $mailboxesWithExistingX400 mailbox(es)."
        if ($calendarDelegateSignalMailboxCount -gt 0) {
            $mailboxCutoverComment += " Calendar delegates are present on $mailboxesWithCalendarDelegates mailbox(es) and should be rebuilt after cutover."
        }
        else {
            $mailboxCutoverComment += ' Calendar delegate state was not collected in this run.'
        }
        if ($fullAccessSignalMailboxCount -gt 0) {
            $mailboxCutoverComment += " Full Access delegates are present on $mailboxesWithFullAccessDelegates mailbox(es) and should be restamped after cutover."
        }
        else {
            $mailboxCutoverComment += ' Full Access state was not collected in this run.'
        }
        if ($sendAsSignalMailboxCount -gt 0) {
            $mailboxCutoverComment += " Send As delegates are present on $mailboxesWithSendAsDelegates mailbox(es) and should be restamped after cutover."
        }
        else {
            $mailboxCutoverComment += ' Send As state was not collected in this run.'
        }
        if ($sendOnBehalfSignalMailboxCount -gt 0) {
            $mailboxCutoverComment += " Send-on-Behalf delegates are present on $mailboxesWithSendOnBehalfDelegates mailbox(es) and should be restamped after cutover."
        }
        else {
            $mailboxCutoverComment += ' Send-on-Behalf state was not collected in this run.'
        }
        $comments.Add($mailboxCutoverComment) | Out-Null
    }
    $comments.Add("- Manual validation is still required for Teams topology details, SharePoint external sharing posture, App Proxy/Private Access/PIM, compliance policy inventory, PST usage, and browser standards.") | Out-Null

    $answers = [ordered]@{}

    $answers['what-licensing-skus-exist-in-the-tenant-and-how-many-of-each'] = $licenseSummary
    $answers['what-is-your-tenant-address'] = @(
        $(if ($tenantDefaultDomain) { "Default domain: $tenantDefaultDomain" }),
        $(if ($tenantInitialDomain) { "Initial domain: $tenantInitialDomain" })
    ) -join '; '
    $answers['is-your-tenant-homed-in-the-united-states-if-not-where'] = if ($tenantCountryCode -eq 'US' -or $tenantPreferredDataLocation -eq 'NAM') { 'Yes. Tenant country/data location indicates United States/North America.' } elseif ($tenantCountry -or $tenantPreferredDataLocation) { "No or unknown. Country: $(Txt $tenantCountry 'Unknown'); preferred data location: $(Txt $tenantPreferredDataLocation 'Unknown')." } else { 'Unknown. Country metadata was not fully populated in this run.' }
    $answers['is-your-tenant-a-multi-geo-environment-if-so-where-is-your-central-location'] = if (Truthy (Prop $tenantInfo 'MultiGeoEnabled')) { "Yes. Central location: $(Txt (Prop $tenantInfo 'MultiGeoCentral') 'not provided by current discovery output')." } else { 'No multi-geo configuration detected in tenant overview data.' }
    $answers['does-your-organization-utilize-microsoft-copilot'] = LicenseAnswer 'Microsoft Copilot' @('(?i)copilot|m365[_\s]*copilot')
    $answers['does-your-organization-utilize-copilot-studio'] = LicenseAnswer 'Copilot Studio' @('(?i)copilot[\s_]*studio|power[\s_]*virtual[\s_]*agents|pva|bots')
    $answers['do-your-users-utilize-microsoft-power-automate'] = LicenseAnswer 'Power Automate' @('(?i)power[\s_]*automate|flow_')
    $answers['do-your-users-utilize-microsoft-power-apps'] = LicenseAnswer 'Power Apps' @('(?i)power[\s_]*apps|powerapps')
    $answers['does-your-organization-use-power-pages'] = LicenseAnswer 'Power Pages' @('(?i)power[\s_]*pages')
    $answers['does-your-organization-use-dynamics-365'] = LicenseAnswer 'Dynamics 365' @('(?i)dynamics[\s_]*365|dynamics')
    $answers['does-your-organization-use-power-bi'] = LicenseAnswer 'Power BI' @('(?i)power[\s_]*bi|power_bi|pbi_')
    $answers['do-you-have-self-service-purchases-enabled'] = (Txt (Prop $selfServicePurchase 'Answer') 'Unknown') + $(if (Prop $selfServicePurchase 'Notes') { "; $(Txt (Prop $selfServicePurchase 'Notes') '')" } else { '' })
    $answers['does-the-organization-utilize-the-source-entra-tenant-for-azure-resources'] = (Txt (Prop $azureResourceUsage 'Answer') 'Not collected') + $(if (Prop $azureResourceUsage 'Notes') { "; $(Txt (Prop $azureResourceUsage 'Notes') '')" } else { '' })

    $answers['what-identity-provider-is-used-for-sign-on'] = if ($federatedDomains) { "Federated sign-on detected. Federated domains: $(Txt ($federatedDomains -join ', ') 'Unknown')." } else { 'Entra ID cloud authentication appears to be the primary sign-on provider.' }
    $answers['are-objects-synchronized-from-on-prem-active-directory'] = if (Truthy (Prop $adConnectSummary 'OnPremisesSyncEnabled')) { "Yes. Directory synchronization is enabled. Last reported sync: $(Txt (Prop $adConnectSummary 'OnPremisesLastSyncDateTime') 'Unknown')." } else { 'No active on-premises directory synchronization detected.' }
    $answers['is-self-service-password-reset-or-password-writeback-enabled'] = 'Not collected by current script; manual validation required.'
    $answers['what-mfa-solution-is-used'] = if (Truthy (Prop $authConfig 'MFAEnabled')) { "Entra ID MFA. Methods: $(Txt (($mfaMethods -join ', ')) 'Not listed'); CA policies requiring MFA: $(Txt (Prop $authConfig 'MFAConditionalAccessPolicies') '0'); registered users: $(Txt (Prop $mfaRegistrationSummary 'RegistrationPercent') 'Unknown')%." } else { 'MFA not clearly detected in current auth policy data; manual validation required.' }
    $answers['is-entra-id-enterprise-sso-used-for-saas-applications'] = if (Truthy (Prop $authConfig 'SSOEnabled')) { "Yes. $(Txt ($ssoApps.Count) '0') SSO-enabled application(s) were discovered." } else { 'No enterprise SSO applications were detected in the current run.' }
    $answers['if-yes-provide-a-list-of-applications'] = if ($ssoApps -and $ssoApps.Count -gt 0) { Summary @($ssoApps) 'DisplayName' 12 } else { 'No SSO application list collected.' }
    $answers['is-entra-id-conditional-access-used'] = if ($conditionalAccessPolicies.Count -gt 0) { "Yes. $($conditionalAccessPolicies.Count) Conditional Access policy/policies collected." } else { 'No Conditional Access policies were collected in this run.' }
    $answers['is-entra-id-app-proxy-used'] = 'Not collected by current script; manual validation required.'
    $answers['is-entra-id-private-access-used'] = 'Not collected by current script; manual validation required.'
    $answers['is-privileged-identity-management-pim-used'] = 'Not collected by current script; manual validation required.'

    $answers['active-user-mailboxes-to-migrate'] = "$(@($mailboxes | Where-Object { (Prop $_ 'RecipientTypeDetails') -eq 'UserMailbox' -and (Prop $_ 'IsInactiveMailbox') -ne $true }).Count) active user mailbox(es)."
    $answers['inactive-user-mailboxes-to-migrate'] = "$($inactiveMailboxes.Count) inactive mailbox(es)."
    $answers['shared-mailboxes-to-migrate'] = "$(@($mailboxes | Where-Object { (Txt (Prop $_ 'RecipientTypeDetails') '') -match 'SharedMailbox' }).Count) shared mailbox(es)."
    $answers['resource-mailboxes-to-migrate'] = "$(@($mailboxes | Where-Object { Truthy (Prop $_ 'IsResource') -or (Txt (Prop $_ 'RecipientTypeDetails') '') -match 'Room|Equipment|Scheduling' }).Count) resource mailbox(es)."
    $answers['archive-mailboxes-to-migrate'] = "$(@($mailboxes | Where-Object { $archiveStatus = Txt (Prop $_ 'ArchiveStatus') ''; -not [string]::IsNullOrWhiteSpace($archiveStatus) -and $archiveStatus -ne 'None' }).Count) archive-enabled mailbox(es)."
    $answers['distribution-lists-to-migrate'] = "$(@($recipients | Where-Object { (Prop $_ 'RecipientTypeDetails') -eq 'MailUniversalDistributionGroup' }).Count) distribution list(s)."
    $answers['dynamic-distribution-lists'] = "$(@($recipients | Where-Object { (Txt (Prop $_ 'RecipientTypeDetails') '') -match 'DynamicDistributionGroup' }).Count) dynamic distribution group(s)."
    $answers['external-contacts-to-migrate'] = "$(@($recipients | Where-Object { (Txt (Prop $_ 'RecipientTypeDetails') '') -match 'Contact|MailUser' }).Count) external contact/mail user object(s)."
    $answers['email-domains-to-migrate'] = if ($customDomains.Count -gt 0) { "$($customDomains.Count) custom domain(s): " + (($customDomains | ForEach-Object { Prop $_ 'Domain' }) -join ', ') } else { 'No custom email domains were collected.' }
    $answers['public-folders'] = "$($publicFolders.Count) public folder object(s)."
    $answers['average-mailbox-size'] = if ($null -ne $averageMailboxSizeGb) { "$averageMailboxSizeGb GB average primary mailbox size." } else { 'Mailbox size statistics were not available in this run.' }
    $answers['largest-mailbox-size'] = if ($largestMailbox.Count -gt 0) { "$($largestMailbox[0].SizeGB) GB ($($largestMailbox[0].DisplayName))." } else { 'Mailbox size statistics were not available in this run.' }
    $answers['hybrid-exchange-configuration'] = if ($hybridInfo) { "$(Txt (Prop $hybridInfo 'HybridStatus') 'Unknown'); type: $(Txt (Prop $hybridInfo 'HybridType') 'Unknown'); evidence: $(Txt (Prop $hybridInfo 'Evidence') 'Not provided')." } else { 'No hybrid Exchange summary was collected.' }
    $answers['on-prem-exchange-servers'] = 'Not collected by current script; manual validation required.'
    $answers['edge-transport-servers'] = 'Not collected by current script; manual validation required.'
    $answers['mail-transport-rules'] = "$($mailFlowRules.Count) transport rule(s) collected."
    $answers['exchange-online-message-protection-encryption'] = if ((LicenseMatches @('(?i)rms|mip|purview|encryption|defender|atp')).Count -gt 0) { 'Microsoft-native protection/encryption capabilities are indicated by licensing, but specific OME or mail-flow encryption configuration was not collected.' } else { 'Exchange Online protection specifics were not collected; manual validation required.' }
    $answers['pst-usage'] = 'Not collected by current script; manual validation required.'
    $answers['3rd-party-mail-hygiene'] = if ($spamFilteringSummary) { if ($usesThirdPartyFiltering) { "Likely yes; mail flow analysis indicates third-party filtering. Inbound connectors: $(Txt (Prop $spamFilteringSummary 'InboundConnectorCount') '0'), outbound connectors: $(Txt (Prop $spamFilteringSummary 'OutboundConnectorCount') '0')." } else { "Not detected from current mail flow analysis. Inbound connectors: $(Txt (Prop $spamFilteringSummary 'InboundConnectorCount') '0'), outbound connectors: $(Txt (Prop $spamFilteringSummary 'OutboundConnectorCount') '0')." } } else { 'Not collected by current script; manual validation required.' }
    $answers['3rd-party-archiving'] = 'Not collected by current script; manual validation required.'
    $answers['additional-address-lists'] = 'Not collected by current script; manual validation required.'
    $answers['send-receive-connectors'] = $mailConnectorSummary

    $answers['onedrive-storage-in-use'] = StorageAnswer $oneDriveSites 'OneDrive'
    $answers['sharepoint-online-storage-in-use'] = StorageAnswer $sharePointSites 'SharePoint'
    $answers['external-sharing-enabled'] = 'Not collected by current script; manual validation required.'
    $answers['data-sync-restricted-to-managed-devices'] = 'Not collected by current script; manual validation required.'
    $answers['sharepoint-sites-to-migrate'] = "$($sharePointSites.Count) SharePoint site(s)."
    $answers['all-sites-modern'] = 'Undetermined from current site inventory alone; manual validation recommended.'
    $answers['user-created-sharepoint-sites-allowed'] = 'Not collected by current script; manual validation required.'
    $answers['sharepoint-hybrid-configuration'] = 'Not collected by current script; manual validation required.'
    $answers['active-onedrive-profiles'] = "$(@($oneDriveSites | Where-Object { $archiveStatus = Txt (Prop $_ 'ArchiveStatus') ''; [string]::IsNullOrWhiteSpace($archiveStatus) -or $archiveStatus -eq 'Active' }).Count) active OneDrive profile(s)."
    $answers['archived-onedrive-profiles'] = "$(@($oneDriveSites | Where-Object { $archiveStatus = Txt (Prop $_ 'ArchiveStatus') ''; -not [string]::IsNullOrWhiteSpace($archiveStatus) -and $archiveStatus -ne 'Active' }).Count) archived OneDrive profile(s)."
    $answers['sharepoint-workflows'] = 'Not collected by current script; manual validation required.'

    $answers['microsoft-teams-to-migrate'] = if ($teams.Count -gt 0) { "$($teams.Count) Microsoft Teams team(s) collected." } else { "Teams topology was not collected in the current PowerShell 7 app-only run. Microsoft 365 groups detected: $m365GroupCount. Manual validation required." }
    $answers['non-teams-microsoft-365-groups'] = if ($teams.Count -gt 0) { "$([math]::Max($m365GroupCount - $teams.Count, 0)) non-Team Microsoft 365 group(s)." } else { "$m365GroupCount Microsoft 365 group(s) detected in Entra. Team association is incomplete in this run, so manual validation is required." }
    $answers['users-can-create-teams'] = 'Not collected by current script; manual validation required.'
    $answers['external-guest-access-enabled'] = if ($guestCount -gt 0 -or $partnerCount -gt 0) { "Yes or likely yes. Guest/external users detected: $guestCount; cross-tenant partner relationships: $partnerCount." } else { 'No guest or external access indicators were detected in the collected data.' }
    $answers['teams-audio-conferencing'] = LicenseAnswer 'Teams Audio Conferencing' @('(?i)audio\s*conferencing|mcomeetadv|mcoaudio')
    $answers['teams-pstn-calling'] = if ($teamsVoiceSummary) { "Voice users: $(Txt (Prop $teamsVoiceSummary 'VoiceUserCount') '0'); data source: $(Txt (Prop $teamsVoiceSummary 'DataSource') 'Unknown'); notes: $(Txt (Prop $teamsVoiceSummary 'Notes') 'None')." } else { LicenseAnswer 'Teams PSTN/Phone System' @('(?i)pstn|phone\s*system|calling\s*plan|mcoev|mcopstn') }
    $answers['teams-rooms'] = LicenseAnswer 'Teams Rooms' @('(?i)teams\s*rooms|meeting\s*room|mtr')
    $answers['teams-room-hardware'] = 'Not collected by current script; manual validation required.'

    $answers['microsoft-defender-products-in-use'] = LicenseAnswer 'Microsoft Defender' @('(?i)defender|atp|mde|endpoint|adallom')
    $answers['non-outlook-client-access-allowed'] = 'Not collected by current script; manual validation required.'
    $answers['mfa-usage-scenarios'] = "Registered users: $(Txt (Prop $mfaRegistrationSummary 'RegisteredUsers') '0')/$(Txt (Prop $mfaRegistrationSummary 'TotalUsers') $users.Count); methods enabled: $(Txt (($mfaMethods -join ', ')) 'Not listed')."
    $answers['insider-risk-management'] = 'Not collected by current script; manual validation required.'
    $answers['customer-lockbox'] = 'Not collected by current script; manual validation required.'
    $answers['background-checks-required-for-arraya-access'] = 'Customer policy item; manual validation required.'

    $answers['ediscovery-solution'] = if ((LicenseMatches @('(?i)purview|ediscovery|discovery|mip|rms')).Count -gt 0) { 'Microsoft Purview/eDiscovery capability appears licensed. Active case inventory was not collected in this run.' } else { 'Not clearly detected in current license inventory; manual validation required.' }
    $answers['active-ediscovery-cases'] = 'Not collected by current script; manual validation required.'
    $answers['data-labeling-classification'] = if ((LicenseMatches @('(?i)mip|rms|purview|label')).Count -gt 0) { 'Labeling and classification capability is indicated by licensing, but policy inventory was not collected.' } else { 'Not clearly detected in current license inventory; manual validation required.' }
    $answers['automatic-classification'] = 'Not collected by current script; manual validation required.'
    $answers['encryption-of-data-at-rest'] = 'Microsoft 365 encrypts customer data at rest by platform default. Customer-managed key usage was not collected in this run.'
    $answers['retention-period'] = 'Not collected by current script; manual validation required.'
    $answers['regulatory-requirements-hipaa-itar-etc'] = 'Not collected by current script; manual validation required.'
    $answers['data-loss-prevention-dlp'] = if ((LicenseMatches @('(?i)dlp|purview|mip|rms')).Count -gt 0) { 'Potential DLP capability is indicated by licensing, but actual DLP policy inventory was not collected.' } else { 'Not clearly detected in current license inventory; manual validation required.' }
    $answers['endpoint-dlp'] = if ((LicenseMatches @('(?i)defender|endpoint|purview|mde')).Count -gt 0) { 'Potential endpoint DLP capability is indicated by licensing, but endpoint DLP policy inventory was not collected.' } else { 'Not clearly detected in current license inventory; manual validation required.' }
    $answers['other-purview-capabilities'] = if ((LicenseMatches @('(?i)purview|mip|rms|contentexplorer|discovery')).Count -gt 0) { 'Detected Purview-related licensing. Workbook and JSON outputs should be used for detailed follow-up.' } else { 'No additional Purview-specific capability was clearly detected from licenses alone.' }
    $answers['encryption-key-type'] = 'Not collected by current script. Default assumption is Microsoft-managed keys unless Customer Key has been separately configured.'

    $answers['microsoft-intune-usage'] = if ($managedDevices -gt 0 -or $intuneOrSccmDevices -gt 0) { "Likely yes. Managed devices: $managedDevices; devices reporting MDM values of $(Txt $mdmSolutionSummary 'Unknown'): $intuneOrSccmDevices." } else { 'No Intune or SCCM managed devices were clearly identified in the current device inventory.' }
    $answers['device-configuration-policies'] = 'Not collected by current script; manual validation required.'
    $answers['device-compliance-policies'] = 'Not collected by current script; manual validation required.'
    $answers['defender-for-endpoint-via-intune'] = if ((LicenseMatches @('(?i)defender|mde|endpoint|atp')).Count -gt 0) { 'Defender licensing is present, but Defender for Endpoint deployment or configuration via Intune was not collected.' } else { 'Not clearly detected in current license inventory; manual validation required.' }
    $answers['on-prem-endpoint-manager-sccm'] = if ($intuneOrSccmDevices -gt 0) { "Possible. $intuneOrSccmDevices device(s) reported MDM values of $(Txt $mdmSolutionSummary 'Unknown'), but SCCM-specific inventory was not collected." } else { 'Not collected by current script; manual validation required.' }
    $answers['applications-deployed-via-intune'] = 'Not collected by current script; manual validation required.'
    $answers['autopilot-usage'] = 'Not collected by current script; manual validation required.'
    $answers['apple-business-manager-android-enterprise'] = 'Not collected by current script; manual validation required.'
    $answers['co-management-enabled'] = 'Not collected by current script; manual validation required.'
    $answers['supported-workstations'] = "$windowsDevices Windows and $macDevices macOS workstation device(s) discovered."
    $answers['corporate-mobile-devices'] = "Mobile devices discovered: $mobileDeviceCount (iOS: $iosDevices, Android: $androidDevices). Ownership or corporate classification was not collected."
    $answers['byod-mobile-devices'] = 'Not collected by current script; manual validation required.'
    $answers['employee-owned-device-management'] = 'Not collected by current script; manual validation required.'
    $answers['ios-vs-android-percentages'] = if ($mobileDeviceCount -gt 0) { "iOS: $([math]::Round((($iosDevices / [double]$mobileDeviceCount) * 100), 1))%; Android: $([math]::Round((($androidDevices / [double]$mobileDeviceCount) * 100), 1))%." } else { 'No mobile device inventory was collected.' }
    $answers['device-join-type-hybrid-cloud'] = if ($trustTypeSummary) { "Trust or join types observed: $trustTypeSummary." } else { 'Device join type was not clearly populated in the current device inventory.' }
    $answers['standard-browser'] = 'Not collected by current script; manual validation required.'

    $answers['microsoft-viva'] = LicenseAnswer 'Microsoft Viva' @('(?i)viva')
    $answers['microsoft-forms'] = LicenseAnswer 'Microsoft Forms' @('(?i)forms')
    $answers['microsoft-planner'] = LicenseAnswer 'Microsoft Planner' @('(?i)planner|projectworkmanagement')
    $answers['microsoft-yammer'] = LicenseAnswer 'Microsoft Yammer/Viva Engage' @('(?i)yammer|engage')
    $answers['microsoft-stream'] = LicenseAnswer 'Microsoft Stream' @('(?i)stream')
    $answers['project-visio-licensing'] = LicenseAnswer 'Project/Visio' @('(?i)project|visio')

    $answers['3rd-party-email-services'] = if ($mailFlowConnectors.Count -gt 0 -or $remoteDomains.Count -gt 0) { "Review required. Connectors: $($mailFlowConnectors.Count); remote domains: $($remoteDomains.Count); third-party filtering detected: $(Txt (Prop $spamFilteringSummary 'Uses3rdPartyFiltering') 'Unknown')." } else { 'No clear third-party email service dependency was detected from current mail flow data.' }
    $answers['3rd-party-device-management'] = if ($mdmSolutionSummary -and $mdmSolutionSummary -notmatch '^Not Managed$') { "Current device inventory reports these MDM values: $mdmSolutionSummary." } else { 'No third-party device management platform was clearly identified.' }
    $answers['3rd-party-cloud-storage'] = 'Not collected by current script; manual validation required.'

    $templateContent = Get-Content -Raw -Path $TemplatePath
    $templateContent = [regex]::Replace($templateContent, '(?m)^\*\*Date:\*\*\s*$', "**Date:** $(Get-Date -Format 'yyyy-MM-dd')")
    $templateContent = [regex]::Replace($templateContent, '(?m)^\*\*Prepared By:\*\*\s*$', "**Prepared By:** $PreparedBy")
    $templateContent = [regex]::Replace($templateContent, '(?m)^\*\*Client:\*\*\s*$', "**Client:** $tenantDisplayName")

    $updatedLines = foreach ($line in ($templateContent -split "`r?`n")) {
        if ($line -match '^\|\s*(.+?)\s*\|\s*\|\s*$' -and $line -notmatch '^\|\s*Discovery Item\s*\|') {
            $questionText = $matches[1].Trim()
            $key = QKey $questionText
            $answer = if ($answers.Contains($key)) { $answers[$key] } else { 'Not collected by current script; manual validation required.' }
            "| $questionText | $(Cell $answer) |"
        } else {
            $line
        }
    }

    $finalLines = New-Object System.Collections.Generic.List[string]
    $commentsInjected = $false
    foreach ($line in $updatedLines) {
        $finalLines.Add($line) | Out-Null
        if (-not $commentsInjected -and $line -eq '## Additional Comments or Concerns') {
            $finalLines.Add('') | Out-Null
            foreach ($comment in $comments) {
                $finalLines.Add($comment) | Out-Null
            }
            $finalLines.Add('') | Out-Null
            $commentsInjected = $true
        }
    }

    [System.IO.File]::WriteAllText($Path, ($finalLines -join [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
}
