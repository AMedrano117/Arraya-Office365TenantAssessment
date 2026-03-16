function Filter-TenantStatsHash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantStatsStore,
        [Parameter(Mandatory = $true)]
        [string]$reportingMode,
        [Parameter(Mandatory = $false)]
        $GraphTest
    )

    # Common tables to remove for both 'combined' and 'minimum' detail levels
    $commonTablesToRemove = @(
        'AllMailboxes-MailIdentity',
        'AllMailboxes-UserPrincipalName',
        'AllMailboxes-PrimarySmtpAddress'
    )

    # Determine additional tables to remove based on the reporting mode
    $additionalTables = switch ($reportingMode) {
        "combined" {
            @(
                'Users', 'ArchiveMailboxes', 'ArchiveMailboxStats', 'NonUserMailboxes',
                'InactiveMailboxes', 'LitigationHoldMailboxes', 'UnifiedGroups'
            )
        }
        "minimum" {
            @(
                'InactiveMailboxes','ArchiveMailboxStats', 'NonUserMailboxes',
                'LitigationHoldMailboxes',
                'RemoteDomains','UnifiedGroups', 'MailFlowConnectors',
                'PublicFolderPerms', 'AuthenticationConfig', 'TenantInfo', 'SpamFilteringConfig',
                'SMTPRelayConfig', 'FederationConfiguration', 'TeamsVoice', 'MfaRegistrationDetails'
            )
        }
        default {
            @() # No additional tables for other reporting modes
        }
    }
    if ($GraphTest -eq "REST") {
        $additionalTables += @(
            "EmailData", "OneDriveData", "SharePointData", "TeamsUserData", "MailboxUsage", 
            "SignInData", "UserSignIns", "SPOUsage", "YammerUsage"
        )
    }

    # Combine common and additional tables to form the full exclusion list
    $tablesToRemove = $commonTablesToRemove + $additionalTables

    $excludeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in $tablesToRemove) {
        if (-not [string]::IsNullOrWhiteSpace([string]$key)) {
            $null = $excludeSet.Add([string]$key)
        }
    }

    # Build filtered output directly to avoid carrying duplicate hash table state.
    $filteredStatsHash = @{}
    foreach ($entry in $TenantStatsStore.GetEnumerator()) {
        if ($excludeSet.Contains([string]$entry.Key)) {
            continue
        }
        $filteredStatsHash[$entry.Key] = $entry.Value
    }

    # Return the filtered hash table
    return $filteredStatsHash
}
