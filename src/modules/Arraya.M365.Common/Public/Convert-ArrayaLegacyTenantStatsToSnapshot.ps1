function Convert-ArrayaLegacyTenantStatsToSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory = $false)]
        [hashtable]$Metadata,
        [Parameter(Mandatory = $false)]
        [hashtable]$CollectionPlan,
        [Parameter(Mandatory = $false)]
        [hashtable]$Diagnostics
    )

    $snapshot = New-ArrayaTenantSnapshot -Metadata $Metadata -CollectionPlan $CollectionPlan -Diagnostics $Diagnostics

    $normalizedTenantStatsHash = @{}
    foreach ($key in $TenantStatsHash.Keys) {
        $normalizedTenantStatsHash[[string]$key] = $TenantStatsHash[$key]
    }

    $getObjectValue = {
        param(
            $Object,
            [string[]]$Names
        )

        if ($null -eq $Object) {
            return $null
        }

        foreach ($name in $Names) {
            if ($Object -is [System.Collections.IDictionary] -and $Object.Contains($name)) {
                return $Object[$name]
            }
            if ($Object.PSObject -and $Object.PSObject.Properties[$name]) {
                return $Object.$name
            }
        }

        return $null
    }

    $toBoolean = {
        param($Value)

        if ($null -eq $Value) {
            return $null
        }

        if ($Value -is [bool]) {
            return $Value
        }

        try {
            return [System.Convert]::ToBoolean($Value)
        }
        catch {
            switch -Regex ([string]$Value) {
                '^(true|yes|enabled|on|1)$' { return $true }
                '^(false|no|disabled|off|0)$' { return $false }
            }
        }

        return $null
    }

    $allMailboxRows = @()
    if ($normalizedTenantStatsHash.ContainsKey('AllMailboxes') -and $normalizedTenantStatsHash['AllMailboxes']) {
        $allMailboxSource = $normalizedTenantStatsHash['AllMailboxes']
        if ($allMailboxSource -is [System.Collections.IDictionary]) {
            $allMailboxRows = @($allMailboxSource.Values)
        }
        else {
            $allMailboxRows = @($allMailboxSource)
        }
    }

    if (-not $normalizedTenantStatsHash.ContainsKey('RetentionPolicies') -and $allMailboxRows.Count -gt 0) {
        $retentionPolicies = [ordered]@{}
        foreach ($mailbox in $allMailboxRows) {
            $policyName = [string](& $getObjectValue $mailbox @('RetentionPolicy'))
            $hasLitigationHold = (& $toBoolean (& $getObjectValue $mailbox @('LitigationHoldEnabled'))) -eq $true
            $hasRetentionHold = (& $toBoolean (& $getObjectValue $mailbox @('RetentionHoldEnabled'))) -eq $true
            $hasDelayHold = (& $toBoolean (& $getObjectValue $mailbox @('DelayHoldApplied'))) -eq $true

            if ([string]::IsNullOrWhiteSpace($policyName) -and -not ($hasLitigationHold -or $hasRetentionHold -or $hasDelayHold)) {
                continue
            }

            $policyKey = if ([string]::IsNullOrWhiteSpace($policyName)) { 'HoldSignalsWithoutNamedPolicy' } else { $policyName }
            if (-not $retentionPolicies.Contains($policyKey)) {
                $retentionPolicies[$policyKey] = [pscustomobject]@{
                    PolicyName                 = $policyKey
                    MailboxCount               = 0
                    LitigationHoldMailboxCount = 0
                    RetentionHoldMailboxCount  = 0
                    DelayHoldMailboxCount      = 0
                }
            }

            $summary = $retentionPolicies[$policyKey]
            $summary.MailboxCount++
            if ($hasLitigationHold) { $summary.LitigationHoldMailboxCount++ }
            if ($hasRetentionHold) { $summary.RetentionHoldMailboxCount++ }
            if ($hasDelayHold) { $summary.DelayHoldMailboxCount++ }
        }

        if ($retentionPolicies.Count -gt 0) {
            $normalizedTenantStatsHash['RetentionPolicies'] = $retentionPolicies
        }
    }

    if (-not $normalizedTenantStatsHash.ContainsKey('PasswordLifecycleSummary') -and $normalizedTenantStatsHash.ContainsKey('AdConnectConfiguration')) {
        $adConnectConfig = $normalizedTenantStatsHash['AdConnectConfiguration']
        $adConnectSummary = & $getObjectValue $adConnectConfig @('Summary')
        if ($adConnectSummary) {
            $passwordLifecycleSummary = [ordered]@{}
            foreach ($field in @(
                'PasswordWriteback',
                'PasswordWritebackEnabled',
                'PassThroughAuthentication',
                'PassThroughAuthenticationEnabled',
                'SelfServicePasswordReset',
                'SelfServicePasswordResetEnabled',
                'OnPremisesSyncEnabled',
                'OnPremisesLastSyncDateTime'
            )) {
                $value = & $getObjectValue $adConnectSummary @($field)
                if ($null -ne $value) {
                    $passwordLifecycleSummary[$field] = $value
                }
            }

            if ($passwordLifecycleSummary.Count -gt 0) {
                $normalizedTenantStatsHash['PasswordLifecycleSummary'] = [pscustomobject]$passwordLifecycleSummary
            }
        }
    }

    if (-not $normalizedTenantStatsHash.ContainsKey('SMTPRelayServiceAccounts') -and $allMailboxRows.Count -gt 0) {
        $smtpAuthMailboxes = @(
            $allMailboxRows | Where-Object {
                (& $toBoolean (& $getObjectValue $_ @('SmtpClientAuthenticationDisabled'))) -eq $false
            }
        )

        if ($smtpAuthMailboxes.Count -gt 0) {
            $emailActivityRows = @()
            if ($normalizedTenantStatsHash.ContainsKey('EmailActivityTopSenders') -and $normalizedTenantStatsHash['EmailActivityTopSenders']) {
                $emailActivitySource = $normalizedTenantStatsHash['EmailActivityTopSenders']
                if ($emailActivitySource -is [System.Collections.IDictionary]) {
                    $emailActivityRows = @($emailActivitySource.Values)
                }
                else {
                    $emailActivityRows = @($emailActivitySource)
                }
            }

            $senderLookup = @{}
            foreach ($sender in $emailActivityRows) {
                foreach ($candidate in @(
                    [string](& $getObjectValue $sender @('UserPrincipalName')),
                    [string](& $getObjectValue $sender @('PrimarySmtpAddress', 'Mail')),
                    [string](& $getObjectValue $sender @('User'))
                )) {
                    if ([string]::IsNullOrWhiteSpace($candidate)) {
                        continue
                    }
                    $senderLookup[$candidate.ToLowerInvariant()] = $sender
                }
            }

            $smtpRelayServiceAccounts = [ordered]@{}
            foreach ($mailbox in $smtpAuthMailboxes) {
                $displayName = [string](& $getObjectValue $mailbox @('DisplayName'))
                $userPrincipalName = [string](& $getObjectValue $mailbox @('UserPrincipalName'))
                $primarySmtpAddress = [string](& $getObjectValue $mailbox @('PrimarySmtpAddress'))
                $lookupKey = @($userPrincipalName, $primarySmtpAddress) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                if ([string]::IsNullOrWhiteSpace($lookupKey)) {
                    $lookupKey = if (-not [string]::IsNullOrWhiteSpace($displayName)) { $displayName } else { [guid]::NewGuid().ToString() }
                }

                $senderRow = $null
                foreach ($candidate in @($userPrincipalName, $primarySmtpAddress, $lookupKey)) {
                    if ([string]::IsNullOrWhiteSpace($candidate)) {
                        continue
                    }
                    $normalizedCandidate = $candidate.ToLowerInvariant()
                    if ($senderLookup.ContainsKey($normalizedCandidate)) {
                        $senderRow = $senderLookup[$normalizedCandidate]
                        break
                    }
                }

                $smtpRelayServiceAccounts[$lookupKey] = [pscustomobject]@{
                    DisplayName        = if ([string]::IsNullOrWhiteSpace($displayName)) { 'Not surfaced in current source' } else { $displayName }
                    UserPrincipalName  = if ([string]::IsNullOrWhiteSpace($userPrincipalName)) { $null } else { $userPrincipalName }
                    PrimarySmtpAddress = if ([string]::IsNullOrWhiteSpace($primarySmtpAddress)) { $null } else { $primarySmtpAddress }
                    SendCount          = & $getObjectValue $senderRow @('SendCount', 'MessageCount')
                    LastActivityDate   = & $getObjectValue $senderRow @('LastActivityDate', 'LastActivityDateTime')
                }
            }

            if ($smtpRelayServiceAccounts.Count -gt 0) {
                $normalizedTenantStatsHash['SMTPRelayServiceAccounts'] = $smtpRelayServiceAccounts
            }
        }
    }

    $sectionMap = @{
        Exchange = @(
            'AllRecipients', 'AllMailboxes', 'AllMailboxes-MailIdentity', 'AllMailboxes-UserPrincipalName',
            'AllMailboxes-PrimarySmtpAddress', 'PrimaryMailboxStats', 'ArchiveMailboxes', 'ArchiveMailboxStats',
            'InactiveMailboxes', 'InactiveMailboxDetails', 'LitigationHoldMailboxes', 'NonUserMailboxes',
            'AllExchangeGroups', 'MailFlowRules', 'MailFlowConnectors', 'PublicFolderDetails', 'PublicFolderPerms',
            'RemoteDomains', 'EmailActivityTopSenders', 'EmailActivityTopReceivers', 'EmailActivitySummary',
            'MailboxFullDetails', 'InboxRulesExternalForwarding', 'InboxRuleForwardingSummary', 'SharedMailboxGovernanceSummary',
            'ForwardingPolicySummary'
        )
        Identity = @(
            'Users', 'UserFullDetails', 'Admins', 'EntraIDGroups', 'DeviceDetails', 'ConditionalAccessPolicies',
            'AuthenticationConfig', 'AuthenticationConfigSummary', 'AuthenticationMethods',
            'AuthenticationSSOApplications', 'LicenseSKUs', 'MfaRegistrationDetails', 'MfaRegistrationSummary',
            'MfaEnrollmentSummary', 'MfaEnforcementSummary', 'MfaEnforcementGapUsers', 'MfaEnforcementScopeReview',
            'ConditionalAccessPolicySummary', 'EnterpriseApplications', 'EnterpriseApplicationSummary',
            'SecurityDefaultsPolicy', 'GuestSignInSummary', 'PrivilegedAccessSummary', 'DeviceManagementSummary',
            'ExternalIdentityRestrictions', 'GuestAccessConfiguration',
            'ConditionalAccessOptimization', 'MfaMethodPostureSummary', 'PrivilegedAccessRemediationSummary',
            'GroupLicensingSummary', 'LicenseOptimizationCandidates'
        )
        Collaboration = @(
            'UnifiedGroups', 'AllTeams', 'TeamsVoice', 'TeamsVoiceSummary',
            'SharePoint', 'OneDrive', 'TeamsActivityTopUsers', 'Office365GroupsActivityTopGroups',
            'UnmanagedObjects', 'OneDriveOwnerMismatches', 'OwnershipGovernanceSummary',
            'SharePointSharingSummary', 'CollaborationActivitySummary', 'TeamsGroupsCleanupCandidates'
        )
        Security = @(
            'SecuritySecureScore', 'SecureScoreActions', 'SpamFilteringConfig', 'SpamFilteringSummary',
            'SMTPRelayConfig', 'SMTPRelaySummary', 'SMTPRelayServiceAccounts'
        )
        Governance = @(
            'RetentionPolicies', 'DlpPolicies', 'PasswordLifecycleSummary'
        )
        Tenant = @(
            'TenantInfo', 'TenantInfoSummary', 'Domains', 'HybridConfiguration',
            'FederationConfiguration', 'FederationSummary', 'AdConnectConfiguration',
            'ExternalSharingSummary', 'ExternalSharingSiteOverrides', 'ExternalExposureFindings'
        )
    }

    $derivedKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in @(
        'BestPractices', 'BestPracticeFindings', 'Findings', 'MigrationReadiness',
        'OwnershipGovernanceSummary', 'UnmanagedObjects', 'OneDriveOwnerMismatches',
        'EmployeeExperienceInsightsSummary', 'LicenseClassificationMetadata',
        'TenantInfoSummary', 'AuthenticationConfigSummary',
        'SpamFilteringSummary', 'FederationSummary', 'MfaRegistrationSummary'
    )) {
        $null = $derivedKeys.Add($key)
    }

    $diagnosticKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in @(
        'PrimaryMailboxStatsCollectionSummary', 'UnifiedGroupMailboxStatsCollectionSummary',
        'CollectorInventory', 'CollectorPerformance', 'CollectorMemoryImpact'
    )) {
        $null = $diagnosticKeys.Add($key)
    }

    $sectionLookup = @{}
    foreach ($domain in $sectionMap.Keys) {
        foreach ($key in $sectionMap[$domain]) {
            $sectionLookup[[string]$key] = $domain
        }
    }

    foreach ($key in $normalizedTenantStatsHash.Keys) {
        $stringKey = [string]$key
        $value = $normalizedTenantStatsHash[$key]

        if ($derivedKeys.Contains($stringKey)) {
            $snapshot['Derived'][$stringKey] = $value
            continue
        }
        if ($diagnosticKeys.Contains($stringKey)) {
            $snapshot['Diagnostics']['CollectorStats'][$stringKey] = $value
            continue
        }

        if ($sectionLookup.Contains($stringKey)) {
            $targetSection = $sectionLookup[$stringKey]
            if (-not $snapshot['Data'].Contains($targetSection) -or -not ($snapshot['Data'][$targetSection] -is [System.Collections.IDictionary])) {
                $snapshot['Data'][$targetSection] = [ordered]@{}
            }
            $snapshot['Data'][$targetSection][$stringKey] = $value
        }
        else {
            if (-not $snapshot['Data'].Contains('Other') -or -not ($snapshot['Data']['Other'] -is [System.Collections.IDictionary])) {
                $snapshot['Data']['Other'] = [ordered]@{}
            }
            $snapshot['Data']['Other'][$stringKey] = $value
        }
    }

    if (
        $snapshot['Data'].Contains('Tenant') -and
        $snapshot['Data']['Tenant'].Contains('TenantInfo') -and
        $snapshot['Data']['Tenant']['TenantInfo']
    ) {
        $tenantInfo = $snapshot['Data']['Tenant']['TenantInfo']
        if (-not $snapshot['Metadata'].Contains('Tenant') -or -not ($snapshot['Metadata']['Tenant'] -is [System.Collections.IDictionary])) {
            $snapshot['Metadata']['Tenant'] = [ordered]@{}
        }
        foreach ($name in @('TenantID', 'TenantId', 'DisplayName', 'DefaultDomainName')) {
            if ($tenantInfo.PSObject.Properties[$name]) {
                $snapshot['Metadata']['Tenant'][$name] = $tenantInfo.$name
            }
            elseif ($tenantInfo -is [System.Collections.IDictionary] -and $tenantInfo.Contains($name)) {
                $snapshot['Metadata']['Tenant'][$name] = $tenantInfo[$name]
            }
        }
    }

    if (-not $snapshot['Derived'].Contains('Findings') -and $snapshot['Derived'].Contains('BestPracticeFindings')) {
        $snapshot['Derived']['Findings'] = $snapshot['Derived']['BestPracticeFindings']
    }

    return $snapshot
}
