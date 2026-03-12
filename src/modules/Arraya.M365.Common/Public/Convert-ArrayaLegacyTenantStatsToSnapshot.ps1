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

    $sectionMap = @{
        Exchange = @(
            'AllRecipients', 'AllMailboxes', 'AllMailboxes-MailIdentity', 'AllMailboxes-UserPrincipalName',
            'AllMailboxes-PrimarySmtpAddress', 'PrimaryMailboxStats', 'ArchiveMailboxes', 'ArchiveMailboxStats',
            'InactiveMailboxes', 'InactiveMailboxDetails', 'LitigationHoldMailboxes', 'NonUserMailboxes',
            'AllExchangeGroups', 'MailFlowRules', 'MailFlowConnectors', 'PublicFolderDetails', 'PublicFolderPerms',
            'RemoteDomains', 'EmailActivityTopSenders', 'EmailActivityTopReceivers', 'EmailActivitySummary',
            'MailboxFullDetails'
        )
        Identity = @(
            'Users', 'UserFullDetails', 'Admins', 'EntraIDGroups', 'DeviceDetails', 'ConditionalAccessPolicies',
            'AuthenticationConfig', 'AuthenticationConfigSummary', 'AuthenticationMethods',
            'AuthenticationSSOApplications', 'LicenseSKUs', 'MfaRegistrationDetails', 'MfaRegistrationSummary'
        )
        Collaboration = @(
            'UnifiedGroups', 'AllTeams', 'TeamsVoice', 'TeamsVoiceSummary',
            'SharePoint', 'OneDrive', 'TeamsActivityTopUsers', 'Office365GroupsActivityTopGroups',
            'UnmanagedObjects', 'OneDriveOwnerMismatches', 'OwnershipGovernanceSummary'
        )
        Security = @(
            'SecuritySecureScore', 'SecureScoreActions', 'SpamFilteringConfig', 'SpamFilteringSummary',
            'SMTPRelayConfig', 'SMTPRelaySummary'
        )
        Tenant = @(
            'TenantInfo', 'TenantInfoSummary', 'Domains', 'HybridConfiguration',
            'FederationConfiguration', 'FederationSummary', 'AdConnectConfiguration'
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

    foreach ($key in $TenantStatsHash.Keys) {
        $stringKey = [string]$key
        $value = $TenantStatsHash[$key]

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
