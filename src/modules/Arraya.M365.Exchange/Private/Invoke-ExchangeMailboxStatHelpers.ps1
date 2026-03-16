function Write-ExoStatisticsFailureSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OperationName,
        [Parameter(Mandatory = $false)]
        $Context,
        [AllowNull()]
        [array]$Failures
    )

    if (-not $Failures -or $Failures.Count -eq 0) {
        return
    }

    $sampleText = @(
        $Failures |
            Select-Object -First 5 |
            ForEach-Object {
                $label = if ([string]::IsNullOrWhiteSpace([string]$_.DisplayName)) {
                    if ([string]::IsNullOrWhiteSpace([string]$_.Identity)) { '<unknown>' } else { $_.Identity }
                } else {
                    $_.DisplayName
                }
                "$label ($($_.Reason))"
            }
    ) -join '; '

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context
    $exportDetails = $Context.ExportFileLocation
    Write-Log -Type WARNING -Message "[$OperationName] Skipped $($Failures.Count) object(s). Sample failures: $sampleText" -ExportFileLocation $exportDetails
}

function Convert-ToMailboxGuidKey {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $GuidValue
    )

    if ($null -eq $GuidValue) {
        return $null
    }

    $text = [string]$GuidValue
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return $text.Trim().Trim('{}').ToLowerInvariant()
}

function Get-MailboxGuidCandidateKeys {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $MailboxRecord,
        [switch]$IncludeArchiveGuid
    )

    if ($null -eq $MailboxRecord) {
        return @()
    }

    $candidateKeys = New-Object System.Collections.Generic.List[string]
    if (
        $IncludeArchiveGuid -and
        $MailboxRecord.PSObject.Properties['ArchiveGuid'] -and
        $MailboxRecord.ArchiveGuid
    ) {
        $archiveKey = Convert-ToMailboxGuidKey -GuidValue $MailboxRecord.ArchiveGuid
        if ($archiveKey -and -not $candidateKeys.Contains($archiveKey)) {
            $candidateKeys.Add($archiveKey)
        }
    }

    foreach ($propertyName in @('ExchangeGuid', 'Guid', 'MailboxGuid')) {
        if (-not ($MailboxRecord.PSObject.Properties[$propertyName])) {
            continue
        }

        $value = $MailboxRecord.PSObject.Properties[$propertyName].Value
        if (-not $value) {
            continue
        }

        $key = Convert-ToMailboxGuidKey -GuidValue $value
        if ($key -and -not $candidateKeys.Contains($key)) {
            $candidateKeys.Add($key)
        }
    }

    return @($candidateKeys)
}

function Test-MailboxStatCached {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $MailboxRecord,
        [AllowNull()]
        $StatsHash,
        [switch]$IncludeArchiveGuid
    )

    if ($null -eq $MailboxRecord -or $null -eq $StatsHash) {
        return $false
    }

    $candidateKeys = Get-MailboxGuidCandidateKeys -MailboxRecord $MailboxRecord -IncludeArchiveGuid:$IncludeArchiveGuid
    if (-not $candidateKeys -or $candidateKeys.Count -eq 0) {
        return $false
    }

    foreach ($key in $candidateKeys) {
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }

        if (
            ($StatsHash.PSObject.Methods['ContainsKey'] -and $StatsHash.ContainsKey($key)) -or
            ($StatsHash -is [System.Collections.IDictionary] -and $StatsHash.Contains($key))
        ) {
            return $true
        }
    }

    return $false
}

function Test-ShouldCollectUnifiedGroupMailboxStats {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Context,
        [Parameter(Mandatory = $false)]
        [string]$DetailLevel
    )

    if ($Context) {
        $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context -DetailLevel $DetailLevel
    }

    $depthPolicy = if ($Context -and $Context.Policies.Contains('CollectionDepth')) {
        $Context.Policies['CollectionDepth']
    }
    else {
        Get-ArrayaExchangeCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($DetailLevel.ToLowerInvariant()))
    }

    if ($depthPolicy -and $depthPolicy.PSObject.Properties['CollectUnifiedGroupMailboxStats']) {
        return [bool]$depthPolicy.CollectUnifiedGroupMailboxStats
    }

    $isMinimum = $false
    if ($depthPolicy -and $depthPolicy.PSObject.Properties['IsMinimum']) {
        $isMinimum = ($depthPolicy.IsMinimum -eq $true)
    }
    elseif (-not [string]::IsNullOrWhiteSpace($DetailLevel) -and $DetailLevel.ToLowerInvariant() -eq 'minimum') {
        $isMinimum = $true
    }
    if ($isMinimum) {
        return $false
    }

    return $true
}

function Get-Office365GroupsActivityMailboxLookup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Context
    )

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context

    $runtime = $Context.Runtime
    $exportDetails = $Context.ExportFileLocation

    if (
        $runtime.Contains('Office365GroupsActivityMailboxLookup') -and
        $runtime['Office365GroupsActivityMailboxLookup'] -and
        $runtime['Office365GroupsActivityMailboxLookup'].PSObject.Properties['ByGroupId'] -and
        $runtime['Office365GroupsActivityMailboxLookup'].PSObject.Properties['ByPrimarySmtpAddress']
    ) {
        return $runtime['Office365GroupsActivityMailboxLookup']
    }

    $lookup = [ordered]@{
        Rows                 = 0
        DownloadSucceeded    = $false
        ByGroupId            = @{}
        ByPrimarySmtpAddress = @{}
    }

    try {
        $rows = @(Office365Custom\Get-GraphAPIActivityReport -ServiceName 'Office365GroupsActivity' -PeriodDuration 'D180')
        $lookup.Rows = @($rows).Count

        foreach ($row in $rows) {
            $groupId = [string](Get-ArrayaObjectValue -Object $row -Names @('Group Id', 'GroupId'))
            if (-not [string]::IsNullOrWhiteSpace($groupId)) {
                $groupIdKey = $groupId.Trim().ToLowerInvariant()
                if (-not $lookup.ByGroupId.ContainsKey($groupIdKey)) {
                    $lookup.ByGroupId[$groupIdKey] = $row
                }
            }

            $primarySmtp = [string](Get-ArrayaObjectValue -Object $row -Names @(
                'Group Principal Name',
                'Group Email',
                'Group Email Address',
                'Group Primary SMTP Address',
                'Group SMTP Address'
            ))
            if (-not [string]::IsNullOrWhiteSpace($primarySmtp)) {
                $smtpKey = $primarySmtp.Trim().ToLowerInvariant()
                if (-not $lookup.ByPrimarySmtpAddress.ContainsKey($smtpKey)) {
                    $lookup.ByPrimarySmtpAddress[$smtpKey] = $row
                }
            }
        }

        $lookup.DownloadSucceeded = $true
    }
    catch {
        Write-Log -Type WARNING -Message "[Get-Office365GroupsActivityMailboxLookup] Unable to download Office 365 Groups activity detail report: $($_.Exception.Message)" -ExportFileLocation $exportDetails
    }

    $resolvedLookup = [pscustomobject]$lookup
    $Context.Runtime['Office365GroupsActivityMailboxLookup'] = $resolvedLookup
    return $resolvedLookup
}

function Get-GroupMailboxActivityRowForUnifiedGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$GroupRecord,
        [Parameter(Mandatory = $true)]
        [object]$ActivityLookup
    )

    if (-not $GroupRecord -or -not $ActivityLookup) {
        return $null
    }

    if ($ActivityLookup.PSObject.Properties['ByGroupId'] -and $ActivityLookup.ByGroupId) {
        $groupIdCandidates = @(
            if ($GroupRecord.PSObject.Properties['ExternalDirectoryObjectId']) { [string]$GroupRecord.ExternalDirectoryObjectId }
            if ($GroupRecord.PSObject.Properties['ExternalDirectoryObjectID']) { [string]$GroupRecord.ExternalDirectoryObjectID }
            if ($GroupRecord.PSObject.Properties['GroupId']) { [string]$GroupRecord.GroupId }
            if ($GroupRecord.PSObject.Properties['Id']) { [string]$GroupRecord.Id }
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($candidateId in $groupIdCandidates) {
            $groupIdKey = $candidateId.Trim().ToLowerInvariant()
            if ($ActivityLookup.ByGroupId.ContainsKey($groupIdKey)) {
                return $ActivityLookup.ByGroupId[$groupIdKey]
            }
        }
    }

    if ($ActivityLookup.PSObject.Properties['ByPrimarySmtpAddress'] -and $ActivityLookup.ByPrimarySmtpAddress) {
        $smtpCandidates = @(
            if ($GroupRecord.PSObject.Properties['PrimarySmtpAddress']) { [string]$GroupRecord.PrimarySmtpAddress }
            if ($GroupRecord.PSObject.Properties['WindowsEmailAddress']) { [string]$GroupRecord.WindowsEmailAddress }
            if ($GroupRecord.PSObject.Properties['UserPrincipalName']) { [string]$GroupRecord.UserPrincipalName }
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($candidateSmtp in $smtpCandidates) {
            $smtpKey = $candidateSmtp.Trim().ToLowerInvariant()
            if ($ActivityLookup.ByPrimarySmtpAddress.ContainsKey($smtpKey)) {
                return $ActivityLookup.ByPrimarySmtpAddress[$smtpKey]
            }
        }
    }

    return $null
}

function New-GroupMailboxStatFromActivityRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$GroupRecord,
        [Parameter(Mandatory = $true)]
        [object]$ActivityRow
    )

    if (-not $GroupRecord -or -not $ActivityRow) {
        return $null
    }

    $storageBytes = Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $ActivityRow -Names @(
        'Exchange Mailbox Storage Used (Byte)',
        'Exchange Mailbox Storage Used (Bytes)',
        'Exchange Mailbox Storage Used'
    )) -AsInt64

    $itemCountValue = Get-ArrayaObjectValue -Object $ActivityRow -Names @(
        'Exchange Mailbox Total Item Count',
        'Exchange Mailbox Item Count'
    )
    if ([string]::IsNullOrWhiteSpace($itemCountValue)) {
        $itemCountValue = '0'
    }

    $displayName = if (
        $GroupRecord.PSObject.Properties['DisplayName'] -and
        -not [string]::IsNullOrWhiteSpace([string]$GroupRecord.DisplayName)
    ) {
        [string]$GroupRecord.DisplayName
    }
    else {
        $graphDisplayName = Get-ArrayaObjectValue -Object $ActivityRow -Names @('Group Display Name')
        if (-not [string]::IsNullOrWhiteSpace($graphDisplayName)) { $graphDisplayName } else { '<Unified Group>' }
    }

    $mailboxGuid = if ($GroupRecord.PSObject.Properties['ExchangeGuid'] -and $GroupRecord.ExchangeGuid) {
        $GroupRecord.ExchangeGuid
    }
    elseif ($GroupRecord.PSObject.Properties['Guid'] -and $GroupRecord.Guid) {
        $GroupRecord.Guid
    }
    else {
        $null
    }

    return [PSCustomObject]@{
        DisplayName               = $displayName
        TotalItemSize             = "$([math]::Round($storageBytes / 1GB, 4)) GB ($storageBytes bytes)"
        TotalItemSizeBytes        = [int64]$storageBytes
        ItemCount                 = [string]$itemCountValue
        TotalDeletedItemSize      = "0 GB (0 bytes)"
        TotalDeletedItemSizeBytes = [int64]0
        MailboxType               = 'GroupMailbox'
        MailboxGuid               = $mailboxGuid
    }
}
