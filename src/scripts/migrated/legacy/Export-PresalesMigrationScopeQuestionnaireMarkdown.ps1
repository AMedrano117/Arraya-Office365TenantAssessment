function Export-PresalesMigrationScopeQuestionnaireMarkdown {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$PreparedBy = "$($env:USERNAME) (automated assessment)",
        [datetime]$AsOfDate = (Get-Date)
    )

    function Rows([string]$Key) {
        if (-not $TenantStatsHash.ContainsKey($Key) -or $null -eq $TenantStatsHash[$Key]) { return @() }
        $value = $TenantStatsHash[$Key]
        if ($value -is [System.Collections.IDictionary]) { return @($value.Values | Where-Object { $null -ne $_ }) }
        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) { return @($value | Where-Object { $null -ne $_ }) }
        return @($value)
    }

    function Prop($Record, [string[]]$Names, $Default = $null) {
        if ($null -eq $Record) { return $Default }
        foreach ($name in $Names) {
            if ($Record -is [System.Collections.IDictionary]) {
                foreach ($key in $Record.Keys) {
                    if ([string]::Equals([string]$key, $name, [System.StringComparison]::OrdinalIgnoreCase)) {
                        return $Record[$key]
                    }
                }
            }
            else {
                $property = $Record.PSObject.Properties |
                    Where-Object { [string]::Equals($_.Name, $name, [System.StringComparison]::OrdinalIgnoreCase) } |
                    Select-Object -First 1
                if ($property) { return $property.Value }
            }
        }
        return $Default
    }

    function Text($Value, [string]$Default = 'Not collected') {
        if ($null -eq $Value) { return $Default }
        if ($Value -is [bool]) { return $(if ($Value) { 'Yes' } else { 'No' }) }
        if ($Value -is [datetime] -or $Value -is [datetimeoffset]) { return ([datetime]$Value).ToString('yyyy-MM-dd') }
        if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
            $items = @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($items.Count -eq 0) { return $Default }
            return ($items -join ', ')
        }
        $result = [string]$Value
        if ([string]::IsNullOrWhiteSpace($result)) { return $Default }
        return $result.Trim()
    }

    function Cell($Value, [string]$Default = '') {
        $result = Text $Value $Default
        $result = $result -replace '\|', '\|'
        $result = $result -replace '[\r\n]+', '<br>'
        return $result.Trim()
    }

    function Number($Value) {
        if ($null -eq $Value) { return $null }
        if ($Value -is [ValueType]) {
            try { return [double]$Value } catch { return $null }
        }

        $result = 0.0
        if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$result)) {
            return $result
        }
        if ([string]$Value -match '(?i)([\d,]+(?:\.\d+)?)\s*(TB|GB|MB|KB)') {
            $size = [double]($Matches[1] -replace ',', '')
            switch ($Matches[2].ToUpperInvariant()) {
                'TB' { return ($size * 1024) }
                'GB' { return $size }
                'MB' { return ($size / 1024) }
                'KB' { return ($size / 1MB) }
            }
        }
        return $null
    }

    function MetricNumber {
        param(
            [object[]]$InputRows,
            [string[]]$MetricNames,
            [string[]]$ValueNames = @('Value', 'Count', 'TotalCount', 'ObjectCount', 'TotalStorageGB')
        )

        foreach ($row in @($InputRows)) {
            $label = Text (Prop $row @('Metric', 'Item', 'Workload', 'Name')) ''
            foreach ($metricName in $MetricNames) {
                if ([string]::Equals($label, $metricName, [System.StringComparison]::OrdinalIgnoreCase)) {
                    return (Number (Prop $row $ValueNames))
                }
            }
        }
        return $null
    }

    function SumSizes {
        param(
            [object[]]$InputRows,
            [string[]]$FieldNames
        )

        $total = 0.0
        $hasKnownValue = $false
        foreach ($row in @($InputRows)) {
            foreach ($fieldName in $FieldNames) {
                $value = Number (Prop $row @($fieldName))
                if ($null -ne $value) {
                    $total += $value
                    $hasKnownValue = $true
                }
            }
        }
        if (-not $hasKnownValue) { return $null }
        return [math]::Round($total, 2)
    }

    function Format-MigrationCount($Value) {
        $numberValue = Number $Value
        if ($null -eq $numberValue) { return 'Not collected' }
        return ('{0:N0}' -f $numberValue)
    }

    function Format-MigrationDataSize($Value) {
        $numberValue = Number $Value
        if ($null -eq $numberValue) { return 'Not collected' }
        if ($numberValue -ge 1024) { return ('{0:N2} TB' -f ($numberValue / 1024)) }
        return ('{0:N2} GB' -f $numberValue)
    }

    function HasArchive($Mailbox) {
        $explicit = Prop $Mailbox @('HasActiveArchive', 'HasArchive')
        if ($explicit -eq $true) { return $true }
        if ((Text $explicit '') -match '^(?i:true|yes|active|enabled)$') { return $true }
        if ((Text (Prop $Mailbox @('ArchiveStatus')) '') -match '(?i)active|enabled') { return $true }
        $archiveGuid = Text (Prop $Mailbox @('ArchiveGuid')) ''
        if ($archiveGuid -and $archiveGuid -notmatch '^0{8}-0{4}-0{4}-0{4}-0{12}$') { return $true }
        $archiveSize = Number (Prop $Mailbox @('ArchiveSizeGB', 'ArchiveGB'))
        return ($null -ne $archiveSize -and $archiveSize -gt 0)
    }

    function ArchiveUnknown($Mailbox) {
        if (HasArchive $Mailbox) { return $false }
        $archiveStatus = Text (Prop $Mailbox @('ArchiveStatus')) ''
        return (
            [string]::IsNullOrWhiteSpace($archiveStatus) -or
            $archiveStatus -match '^(?i:unknown|unavailable|not available|not collected|n/?a)(?:\b|$)'
        )
    }

    function TeamEvidence($TeamRow) {
        $channelStatus = Text (Prop $TeamRow @('ChannelInventoryStatus', 'ChannelEvidenceStatus')) ''
        $memberStatus = Text (Prop $TeamRow @('MemberInventoryStatus', 'MemberEvidenceStatus')) ''
        $sharedChannelCount = Number (Prop $TeamRow @('SharedChannelCount'))
        if ($null -eq $sharedChannelCount) { $sharedChannelCount = 0 }

        [pscustomobject]@{
            ChannelMissing = ($channelStatus -match '(?i)needs?\s+data|unknown|unavailable|failed|not\s+(?:collected|requested)|partial')
            MemberMissing = ($memberStatus -match '(?i)needs?\s+data|unknown|unavailable|failed|not\s+(?:collected|requested)|partial')
            SharedChannelCount = [int]$sharedChannelCount
        }
    }

    function ShortText($Value, [int]$MaximumLength = 320) {
        $result = Text $Value ''
        if ($result.Length -le $MaximumLength) { return $result }
        return ($result.Substring(0, ($MaximumLength - 3)).TrimEnd() + '...')
    }

    # Vendor packaging changes faster than the exporter. Keep the reviewed facts in a
    # versioned catalog so plan thresholds, prices, and source links can be updated without
    # rebuilding the report layout. Safe defaults preserve replay compatibility if the
    # standalone helper is copied without the repository config.
    $toolingCatalog = $null
    $toolingCatalogPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\config\baseline\migration-tooling-options.json'))
    if (Test-Path -LiteralPath $toolingCatalogPath) {
        try {
            $toolingCatalog = Get-Content -LiteralPath $toolingCatalogPath -Raw | ConvertFrom-Json
        }
        catch {
            Write-Verbose "Unable to load migration tooling catalog '$toolingCatalogPath': $($_.Exception.Message)"
        }
    }

    $toolingReferenceDate = Text (Prop $toolingCatalog @('LastReviewed')) 'not recorded'
    $toolingReviewPolicy = Text (Prop $toolingCatalog @('ReviewPolicy')) 'Revalidate vendor terms during solution design and immediately before ordering.'
    $shareGateCatalog = Prop $toolingCatalog @('ShareGate')
    $shareGatePlans = @((Prop $shareGateCatalog @('Plans') @()))
    function ShareGatePlan([string]$Key) {
        return @($shareGatePlans | Where-Object {
                [string]::Equals((Text (Prop $_ @('Key')) ''), $Key, [System.StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1) | Select-Object -First 1
    }

    $shareGateEssentialsCatalog = ShareGatePlan 'Essentials'
    $shareGateProCatalog = ShareGatePlan 'Pro'
    $shareGateEnterpriseCatalog = ShareGatePlan 'Enterprise'
    $shareGateEssentialsMaximum = Number (Prop $shareGateEssentialsCatalog @('MaximumEmployees'))
    if ($null -eq $shareGateEssentialsMaximum) { $shareGateEssentialsMaximum = 250 }
    $shareGateProMaximum = Number (Prop $shareGateProCatalog @('MaximumEmployees'))
    if ($null -eq $shareGateProMaximum) { $shareGateProMaximum = 1000 }
    $shareGatePricingUrl = Text (Prop $shareGateCatalog @('PricingUrl')) 'https://sharegate.com/pricing'

    $microsoftCatalog = Prop $toolingCatalog @('Microsoft')
    $orchestratorCatalog = Prop $microsoftCatalog @('MigrationOrchestrator')
    $nativeUserLicenseCatalog = Prop $microsoftCatalog @('CrossTenantUserDataMigration')
    $nativeSharedDataCatalog = Prop $microsoftCatalog @('CrossTenantSharedDataMigration')
    $fastTrackCatalog = Prop $microsoftCatalog @('FastTrack')
    $bitTitanCatalog = Prop $toolingCatalog @('BitTitan')

    $orchestratorUrl = Text (Prop $orchestratorCatalog @('SourceUrl')) 'https://learn.microsoft.com/en-us/microsoft-365/migration/migration-orchestrator-1-overview?view=o365-worldwide'
    $nativeMailboxUrl = Text (Prop $nativeUserLicenseCatalog @('SourceUrl')) 'https://learn.microsoft.com/en-us/microsoft-365/migration/cross-tenant-mailbox-migration?view=o365-worldwide'
    $nativeSharePointUrl = Text (Prop $nativeSharedDataCatalog @('SourceUrl')) 'https://learn.microsoft.com/en-us/microsoft-365/migration/cross-tenant-sharepoint-migration?view=o365-worldwide'
    $fastTrackUrl = Text (Prop $fastTrackCatalog @('SourceUrl')) 'https://learn.microsoft.com/en-us/microsoft-365/fasttrack/cross-tenant-migration'
    $bitTitanLicenseUrl = Text (Prop $bitTitanCatalog @('LicenseGuideUrl')) 'https://help.bittitan.com/hc/en-us/articles/36203678661531-Which-Migration-License-Do-I-Need'

    $tenantInfo = if ($TenantStatsHash.ContainsKey('TenantInfo')) { $TenantStatsHash['TenantInfo'] } else { $null }
    $tenantName = Text (Prop $tenantInfo @('DisplayName', 'TenantName')) 'Unknown tenant'
    $defaultDomain = Text (Prop $tenantInfo @('DefaultDomain', 'InitialDomain')) 'Not collected'

    $executiveRows = @(Rows 'MigrationExecutiveSummary')
    $scopeRows = @(Rows 'MigrationScopeSummary')
    $licenseSummaryRows = @(Rows 'BitTitanLicenseSummary')
    $licenseDetailRows = @(Rows 'BitTitanLicenseDetail')
    $collaborationRows = @(Rows 'ShareGateScopeSummary')
    if ($collaborationRows.Count -eq 0) { $collaborationRows = @(Rows 'CollaborationSummary') }
    $complexityRows = @(Rows 'MigrationComplexityFlags')
    $users = @(Rows 'Users')
    $mailboxes = @(Rows 'AllMailboxes')
    if ($mailboxes.Count -eq 0) { $mailboxes = @(Rows 'MailboxFullDetails') }
    $inactiveMailboxes = @(Rows 'InactiveMailboxDetails')
    if ($inactiveMailboxes.Count -eq 0) {
        $inactiveMailboxes = @($mailboxes | Where-Object {
                (Prop $_ @('IsInactiveMailbox')) -eq $true -or
                (Text (Prop $_ @('IsInactiveMailbox')) '') -match '^(?i:true|yes)$'
            })
    }
    $groupRows = @(Rows 'GroupWorkloadReconciliation')
    if ($groupRows.Count -eq 0) { $groupRows = @(Rows 'GroupMailboxes') }
    $teams = @(Rows 'AllTeams')
    $sharePointSites = @(Rows 'SharePoint')
    $oneDriveSites = @(Rows 'OneDrive')
    $publicFolders = @(Rows 'PublicFolderDetails')
    if ($publicFolders.Count -eq 0) { $publicFolders = @(Rows 'PublicFolders') }
    $domains = @(Rows 'Domains')

    $enabledMemberUsers = MetricNumber -InputRows $executiveRows -MetricNames @('Enabled member users')
    if ($null -eq $enabledMemberUsers -and $users.Count -gt 0) {
        $enabledMemberUsers = @($users | Where-Object {
                (Text (Prop $_ @('UserType')) '') -ne 'Guest' -and
                (Text (Prop $_ @('UserPrincipalName')) '') -notlike '*#EXT#*' -and
                ((Prop $_ @('AccountEnabled')) -ne $false)
            }).Count
    }
    if ($null -eq $enabledMemberUsers) {
        $enabledMemberUsers = MetricNumber -InputRows $scopeRows -MetricNames @('Licensable users with a mailbox')
    }

    $mailboxCount = MetricNumber -InputRows $executiveRows -MetricNames @('Mailboxes')
    if ($null -eq $mailboxCount) {
        if ($mailboxes.Count -gt 0) { $mailboxCount = $mailboxes.Count }
        elseif ($licenseDetailRows.Count -gt 0) { $mailboxCount = $licenseDetailRows.Count }
    }

    $mailboxDataGB = MetricNumber -InputRows $scopeRows -MetricNames @('Total mailbox data to migrate (GB)')
    if ($null -eq $mailboxDataGB) {
        $mailboxDataGB = MetricNumber -InputRows $licenseSummaryRows -MetricNames @('Grand total data to migrate (GB)')
    }
    if ($null -eq $mailboxDataGB) {
        $mailboxDataGB = MetricNumber -InputRows $executiveRows -MetricNames @('Grand total data to migrate (GB)')
    }
    if ($null -eq $mailboxDataGB -and $mailboxes.Count -gt 0) {
        $mailboxDataGB = SumSizes -InputRows $mailboxes -FieldNames @('MailboxSizeGB', 'DeletedItemsGB', 'ArchiveSizeGB', 'ArchiveDeletedItemsGB')
    }

    $sharePointCount = MetricNumber -InputRows $scopeRows -MetricNames @('SharePoint sites (total)')
    if ($null -eq $sharePointCount) {
        $sharePointCount = MetricNumber -InputRows $executiveRows -MetricNames @('SharePoint sites')
    }
    if ($null -eq $sharePointCount -and $sharePointSites.Count -gt 0) { $sharePointCount = $sharePointSites.Count }

    $sharePointStorageGB = MetricNumber -InputRows $scopeRows -MetricNames @('SharePoint storage (GB)')
    if ($null -eq $sharePointStorageGB) {
        $sharePointStorageGB = MetricNumber -InputRows $collaborationRows -MetricNames @('SharePoint (total)', 'SharePoint') -ValueNames @('TotalStorageGB', 'Value')
    }
    if ($null -eq $sharePointStorageGB -and $sharePointSites.Count -gt 0) {
        $sharePointStorageGB = SumSizes -InputRows $sharePointSites -FieldNames @('StorageUsedGB', 'SiteSizeGB')
    }

    $oneDriveCount = MetricNumber -InputRows $scopeRows -MetricNames @('OneDrive sites')
    if ($null -eq $oneDriveCount) {
        $oneDriveCount = MetricNumber -InputRows $executiveRows -MetricNames @('OneDrive sites')
    }
    if ($null -eq $oneDriveCount -and $oneDriveSites.Count -gt 0) { $oneDriveCount = $oneDriveSites.Count }

    $oneDriveStorageGB = MetricNumber -InputRows $scopeRows -MetricNames @('OneDrive storage (GB)')
    if ($null -eq $oneDriveStorageGB) {
        $oneDriveStorageGB = MetricNumber -InputRows $collaborationRows -MetricNames @('OneDrive') -ValueNames @('TotalStorageGB', 'Value')
    }
    if ($null -eq $oneDriveStorageGB -and $oneDriveSites.Count -gt 0) {
        $oneDriveStorageGB = SumSizes -InputRows $oneDriveSites -FieldNames @('StorageUsedGB', 'SiteSizeGB')
    }

    $teamCount = MetricNumber -InputRows $scopeRows -MetricNames @('Teams')
    if ($null -eq $teamCount) { $teamCount = MetricNumber -InputRows $executiveRows -MetricNames @('Teams') }
    if ($null -eq $teamCount -and $teams.Count -gt 0) { $teamCount = $teams.Count }

    $groupCount = MetricNumber -InputRows $scopeRows -MetricNames @('Microsoft 365 Groups')
    if ($null -eq $groupCount -and $groupRows.Count -gt 0) { $groupCount = $groupRows.Count }

    $collaborationObjectCount = MetricNumber -InputRows $collaborationRows -MetricNames @('All collaboration sites (non-duplicated)') -ValueNames @('ObjectCount', 'TotalCount', 'Value')
    if ($null -eq $collaborationObjectCount) {
        $knownSiteCounts = @($sharePointCount, $oneDriveCount | Where-Object { $null -ne $_ })
        if ($knownSiteCounts.Count -gt 0) { $collaborationObjectCount = [double](($knownSiteCounts | Measure-Object -Sum).Sum) }
    }

    $collaborationStorageGB = MetricNumber -InputRows $collaborationRows -MetricNames @('All collaboration sites (non-duplicated)') -ValueNames @('TotalStorageGB', 'Value')
    if ($null -eq $collaborationStorageGB) {
        $knownStorageValues = @($sharePointStorageGB, $oneDriveStorageGB | Where-Object { $null -ne $_ })
        if ($knownStorageValues.Count -gt 0) { $collaborationStorageGB = [double](($knownStorageValues | Measure-Object -Sum).Sum) }
    }

    $inactiveWithArchive = @($inactiveMailboxes | Where-Object { HasArchive $_ })
    $inactiveUnknownArchive = @($inactiveMailboxes | Where-Object { ArchiveUnknown $_ })
    $archiveEnabledCount = MetricNumber -InputRows $licenseSummaryRows -MetricNames @('Archive-enabled mailboxes')
    if ($null -eq $archiveEnabledCount -and $mailboxes.Count -gt 0) {
        $archiveEnabledCount = @($mailboxes | Where-Object { HasArchive $_ }).Count
    }
    $mailboxesOver50 = MetricNumber -InputRows $licenseSummaryRows -MetricNames @('Mailboxes over 50 GB')
    $mailboxesOver100 = MetricNumber -InputRows $licenseSummaryRows -MetricNames @('Mailboxes over 100 GB')
    $archivesOver100 = MetricNumber -InputRows $licenseSummaryRows -MetricNames @('Archives over 100 GB')
    $groupMailboxNeedsEvidence = MetricNumber -InputRows $licenseSummaryRows -MetricNames @('Microsoft 365 Group mailboxes needing evidence')
    if ($null -eq $groupMailboxNeedsEvidence -and $groupRows.Count -gt 0) {
        $groupMailboxNeedsEvidence = @($groupRows | Where-Object {
                $status = Text (Prop $_ @('MailboxEvidenceStatus', 'MailboxSizingStatus', 'MailboxDataStatus')) ''
                $size = Number (Prop $_ @('MailboxSizeGB', 'SizeGB'))
                $status -match '(?i)unknown|unavailable|needs\s+data|not\s+collected|missing' -or
                ([string]::IsNullOrWhiteSpace($status) -and $null -eq $size)
            }).Count
    }

    $teamEvidence = @($teams | ForEach-Object { TeamEvidence $_ })
    $sharedChannelCount = [int](($teamEvidence | Measure-Object -Property SharedChannelCount -Sum).Sum)
    $teamsMissingChannels = @($teamEvidence | Where-Object { $_.ChannelMissing }).Count
    $teamsMissingMembers = @($teamEvidence | Where-Object { $_.MemberMissing }).Count

    $sourceDomainCount = if ($domains.Count -gt 0) { $domains.Count } else { $null }
    $employeeBasis = $enabledMemberUsers
    if ($null -eq $employeeBasis -or $employeeBasis -le 0) { $employeeBasis = $mailboxCount }

    $callouts = New-Object System.Collections.Generic.List[object]
    $calloutKeys = @{}
    function Add-Callout {
        param([string]$Title, [string]$Observation, [string]$Response)
        if ($callouts.Count -ge 8 -or [string]::IsNullOrWhiteSpace($Title)) { return }
        $key = ($Title -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
        if ($calloutKeys.ContainsKey($key)) { return }
        $calloutKeys[$key] = $true
        $callouts.Add([pscustomobject]@{
                Title = $Title
                Observation = ShortText $Observation 220
                Response = ShortText $Response 260
            }) | Out-Null
    }

    if ($inactiveMailboxes.Count -gt 0) {
        Add-Callout -Title 'Inactive mailboxes and archives' `
            -Observation ("{0} inactive mailbox(es); {1} have archive evidence and {2} have an unknown archive state." -f $inactiveMailboxes.Count, $inactiveWithArchive.Count, $inactiveUnknownArchive.Count) `
            -Response 'Treat restoration, holds, archive handling, and licensing as an exception workstream before scheduling migration.'
    }
    if ($publicFolders.Count -gt 0) {
        Add-Callout -Title 'Public folders' `
            -Observation ("{0} public folder object(s) were discovered." -f $publicFolders.Count) `
            -Response 'Scope public folders separately; normal user-mailbox and collaboration migration paths do not automatically cover them.'
    }
    if ($sharedChannelCount -gt 0) {
        Add-Callout -Title 'Teams shared channels' `
            -Observation ("{0} shared channel(s) were measured across the source Teams inventory." -f $sharedChannelCount) `
            -Response 'Validate current tool support and plan recreation, membership, and external-access handling before collaboration cutover.'
    }
    if ($teamsMissingChannels -gt 0 -or $teamsMissingMembers -gt 0) {
        Add-Callout -Title 'Incomplete Teams evidence' `
            -Observation ("Channel details were unavailable for {0} Team(s); member or guest details were unavailable for {1}." -f $teamsMissingChannels, $teamsMissingMembers) `
            -Response 'Collect the missing collaboration evidence during design before committing feature-remediation effort.'
    }
    if ($groupMailboxNeedsEvidence -gt 0) {
        Add-Callout -Title 'Group mailbox conversation evidence' `
            -Observation ("{0} Microsoft 365 Group mailbox(es) lack reliable conversation sizing." -f (Format-MigrationCount $groupMailboxNeedsEvidence)) `
            -Response 'Confirm whether group conversations must migrate; do not infer that missing statistics means an empty mailbox.'
    }
    if (($mailboxesOver100 -gt 0) -or ($archivesOver100 -gt 0) -or ($mailboxesOver50 -gt 0)) {
        Add-Callout -Title 'Large mailbox content' `
            -Observation ("{0} mailbox(es) exceed 50 GB, {1} exceed 100 GB, and {2} archive(s) exceed 100 GB." -f (Format-MigrationCount $mailboxesOver50), (Format-MigrationCount $mailboxesOver100), (Format-MigrationCount $archivesOver100)) `
            -Response 'Validate license entitlement, pre-stage duration, and exception batching before the production cutover.'
    }

    $largestCollaborationRow = @($collaborationRows |
            Where-Object { $null -ne (Number (Prop $_ @('LargestObjectSizeGB'))) } |
            Sort-Object @{ Expression = { Number (Prop $_ @('LargestObjectSizeGB')) }; Descending = $true } |
            Select-Object -First 1) | Select-Object -First 1
    $largestCollaborationSize = Number (Prop $largestCollaborationRow @('LargestObjectSizeGB'))
    if ($null -ne $largestCollaborationSize -and $largestCollaborationSize -ge 250) {
        $largestCollaborationName = Text (Prop $largestCollaborationRow @('LargestObjectName')) 'Largest discovered collaboration site'
        Add-Callout -Title 'Large collaboration site' `
            -Observation ("{0} is approximately {1}." -f $largestCollaborationName, (Format-MigrationDataSize $largestCollaborationSize)) `
            -Response 'Run a representative throughput test and validate item-count, path, sharing, and protected-content constraints.'
    }

    $orderedComplexityRows = @($complexityRows | Sort-Object @{ Expression = {
                    $status = Text (Prop $_ @('Status', 'ReadinessStatus')) ''
                    if ($status -match '(?i)block') { 0 }
                    elseif ($status -match '(?i)needs|missing') { 1 }
                    elseif ($status -match '(?i)review|conditional|open') { 2 }
                    else { 3 }
                } })
    foreach ($row in $orderedComplexityRows) {
        $status = Text (Prop $row @('Status', 'ReadinessStatus')) ''
        if ($status -notmatch '(?i)block|needs|missing|review|conditional|open') { continue }
        $title = Text (Prop $row @('Item', 'Check', 'Metric')) ''
        $value = Text (Prop $row @('Value', 'DiscoveredValue')) ''
        $observation = @($value, (Prop $row @('Notes'))) |
            ForEach-Object { Text $_ '' } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
        $response = Text (Prop $row @('MigrationAction', 'Resolution', 'RequiredAction')) 'Validate this dependency during migration design.'
        Add-Callout -Title $title -Observation ($observation -join '; ') -Response $response
    }

    $totalKnownDataGB = $null
    $knownDataValues = @($mailboxDataGB, $collaborationStorageGB | Where-Object { $null -ne $_ })
    if ($knownDataValues.Count -gt 0) { $totalKnownDataGB = [double](($knownDataValues | Measure-Object -Sum).Sum) }

    $phasedReasons = New-Object System.Collections.Generic.List[string]
    if ($null -ne $employeeBasis -and $employeeBasis -gt 250) {
        $phasedReasons.Add(("The source population is approximately {0} users/mailbox users, above a typical small-project single-event baseline." -f (Format-MigrationCount $employeeBasis))) | Out-Null
    }
    if ($null -ne $collaborationObjectCount -and $collaborationObjectCount -gt 250) {
        $phasedReasons.Add(("The source contains approximately {0} SharePoint and OneDrive sites." -f (Format-MigrationCount $collaborationObjectCount))) | Out-Null
    }
    if ($null -ne $totalKnownDataGB -and $totalKnownDataGB -gt 2048) {
        $phasedReasons.Add(("Known source content is approximately {0}, which warrants staged throughput validation." -f (Format-MigrationDataSize $totalKnownDataGB))) | Out-Null
    }
    $sequencingDependencies = @($complexityRows | Where-Object {
            $status = Text (Prop $_ @('Status', 'ReadinessStatus')) ''
            $item = Text (Prop $_ @('Item', 'Check', 'Metric')) ''
            $status -match '(?i)block' -or
            $item -match '(?i)exchange hybrid|directory synchronization|ad connect|custom-domain application|multi-geo'
        })
    if ($sequencingDependencies.Count -gt 0) {
        $phasedReasons.Add('Source-observed identity, hybrid, domain, or application dependencies require deliberate sequencing.') | Out-Null
    }

    $hasSizingBasis = ($null -ne $employeeBasis -and $employeeBasis -gt 0) -or ($null -ne $totalKnownDataGB)
    $phasedRecommended = ($phasedReasons.Count -gt 0)
    if ($phasedRecommended) {
        $recommendedMethod = 'Phased migration; define production batches during solution design'
        $methodNarrative = 'A phased approach is likely beneficial. Use a technical pilot first, then size production batches from measured throughput, business groupings, and coexistence constraints. This source-only assessment intentionally does not manufacture wave counts or memberships.'
    }
    elseif ($hasSizingBasis) {
        $recommendedMethod = 'Technical pilot followed by a single production cutover'
        $methodNarrative = 'A single-event (big bang) production cutover is a reasonable planning baseline for the discovered scale, provided the target design, pilot throughput, support coverage, and business outage window validate it. A technical pilot is still recommended.'
    }
    else {
        $recommendedMethod = 'Select the method during solution design'
        $methodNarrative = 'The source snapshot does not contain enough population or data-volume evidence to recommend a single-event or phased production approach. Complete the aggregate sizing before committing a delivery method.'
    }

    $hasCollaborationScope = (($teamCount -gt 0) -or ($sharePointCount -gt 0) -or ($oneDriveCount -gt 0))
    $shareGatePlan = 'Not indicated by collected scope'
    $shareGateReason = 'No Teams, SharePoint, or OneDrive objects were surfaced in this source snapshot.'
    $shareGateSelectedCatalog = $null
    if ($hasCollaborationScope) {
        if ($null -ne $employeeBasis -and $employeeBasis -gt $shareGateProMaximum) {
            $shareGateSelectedCatalog = $shareGateEnterpriseCatalog
            $shareGatePlan = Text (Prop $shareGateSelectedCatalog @('Name')) 'ShareGate Migrate Enterprise'
            $machineActivations = Text (Prop $shareGateSelectedCatalog @('MachineActivations')) '25'
            $shareGateReason = "The source population exceeds $(Format-MigrationCount $shareGateProMaximum); Enterprise is the vendor-positioned tier for larger tenants and includes up to $machineActivations migration workstations plus 24/7 prioritized support."
        }
        elseif (
            ($null -ne $employeeBasis -and $employeeBasis -gt $shareGateEssentialsMaximum) -or
            ($null -ne $collaborationObjectCount -and $collaborationObjectCount -gt 250) -or
            ($null -ne $collaborationStorageGB -and $collaborationStorageGB -gt 2048)
        ) {
            $shareGateSelectedCatalog = $shareGateProCatalog
            $shareGatePlan = Text (Prop $shareGateSelectedCatalog @('Name')) 'ShareGate Migrate Pro'
            $machineActivations = Text (Prop $shareGateSelectedCatalog @('MachineActivations')) '5'
            $shareGateReason = "The source is beyond the simple-project range; Pro supports up to $machineActivations migration workstations and parallel execution. It also adds Exchange Online, Entra ID preview, sensitivity-label, and automation capabilities."
        }
        else {
            $shareGateSelectedCatalog = $shareGateEssentialsCatalog
            $shareGatePlan = Text (Prop $shareGateSelectedCatalog @('Name')) 'ShareGate Migrate Essentials'
            $machineActivations = Text (Prop $shareGateSelectedCatalog @('MachineActivations')) '1'
            $shareGateReason = "Best fit for a straightforward collaboration migration with up to $(Format-MigrationCount $shareGateEssentialsMaximum) employees and $machineActivations migration workstation. Essentials covers SharePoint, Teams, Planner, OneDrive, and file shares, but not Exchange Online or Entra ID migration."
        }
    }

    $migrationWizMailboxUnits = $null
    $migrationWizUmbUnits = $null
    foreach ($row in $licenseSummaryRows) {
        $metric = Text (Prop $row @('Metric', 'Item')) ''
        $value = Number (Prop $row @('Value', 'Count'))
        if ($null -eq $migrationWizMailboxUnits -and $metric -match '(?i)^Mailbox Migration licenses$|^Workload-fit: MigrationWiz-Mailbox units$') {
            $migrationWizMailboxUnits = $value
        }
        if ($null -eq $migrationWizUmbUnits -and $metric -match '(?i)^User Migration Bundles$|^Workload-fit: User Migration Bundles$') {
            $migrationWizUmbUnits = $value
        }
    }
    $migrationWizBasis = if ($null -ne $migrationWizMailboxUnits -or $null -ne $migrationWizUmbUnits) {
        "Current source estimate: $(Format-MigrationCount $migrationWizMailboxUnits) mailbox license unit(s) and $(Format-MigrationCount $migrationWizUmbUnits) User Migration Bundle unit(s). Recalculate after final scope and current vendor entitlement validation."
    }
    else {
        'Per-user or per-workload licenses are an additional purchase; calculate the basket after final mailbox, archive, and exception scope is confirmed.'
    }

    $nativeUserLicenseBasis = if ($null -ne $employeeBasis -and $employeeBasis -gt 0) {
        "Up to $(Format-MigrationCount $employeeBasis) Cross-Tenant User Data Migration add-on license(s) on the current source-population basis; final quantity is the users actually migrated."
    }
    else {
        'Cross-Tenant User Data Migration add-on licenses are purchased once per migrating user; quantity requires confirmed user scope.'
    }
    $sharedDataUnitSizeGB = Number (Prop $nativeSharedDataCatalog @('UnitSizeGB'))
    if ($null -eq $sharedDataUnitSizeGB -or $sharedDataUnitSizeGB -le 0) { $sharedDataUnitSizeGB = 100 }
    $sharedDataAvailability = Text (Prop $nativeSharedDataCatalog @('Availability')) 'Validate current purchasing-channel eligibility'
    $nativeSharePointLicenseBasis = if ($null -ne $sharePointStorageGB) {
        $sharedDataUnits = [math]::Ceiling($sharePointStorageGB / $sharedDataUnitSizeGB)
        "Native SharePoint migration currently uses the separate Cross-Tenant Shared Data Migration SKU in $(Format-MigrationCount $sharedDataUnitSizeGB) GB units; $sharedDataAvailability. Approximately $sharedDataUnits unit(s) from known SharePoint storage, before eligibility and growth validation."
    }
    else {
        "Native SharePoint migration currently uses a separate Cross-Tenant Shared Data Migration SKU in $(Format-MigrationCount $sharedDataUnitSizeGB) GB units; $sharedDataAvailability. Quantity requires known SharePoint storage."
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# Microsoft 365 Migration Scope Overview') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("**Client:** $(Cell $tenantName)") | Out-Null
    $lines.Add("**Source tenant domain:** $(Cell $defaultDomain)") | Out-Null
    $lines.Add("**Assessment date:** $($AsOfDate.ToString('yyyy-MM-dd'))") | Out-Null
    $lines.Add("**Prepared by:** $(Cell $PreparedBy)") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('> This is a source-tenant-only overview. It summarizes scale and exceptional complexity; it does not assess the destination tenant, collect interactive responses, or replace detailed project design. Full object inventories remain in the supporting workbook and snapshot and are intentionally not repeated here.') | Out-Null

    $lines.Add('') | Out-Null
    $lines.Add('## Scope at a glance') | Out-Null
    $lines.Add('') | Out-Null
    $identitySummary = "$(Format-MigrationCount $enabledMemberUsers) enabled member user(s)"
    if ($null -ne $sourceDomainCount) { $identitySummary += "; $(Format-MigrationCount $sourceDomainCount) source domain record(s)" }
    $lines.Add("- **Identity:** $identitySummary.") | Out-Null

    $mailSummaryParts = New-Object System.Collections.Generic.List[string]
    $mailSummaryParts.Add("$(Format-MigrationCount $mailboxCount) mailbox object(s)") | Out-Null
    $mailSummaryParts.Add("$(Format-MigrationDataSize $mailboxDataGB) known mailbox, deleted-item, and archive data") | Out-Null
    if ($null -ne $archiveEnabledCount) { $mailSummaryParts.Add("$(Format-MigrationCount $archiveEnabledCount) archive-enabled mailbox(es)") | Out-Null }
    if ($inactiveMailboxes.Count -gt 0) { $mailSummaryParts.Add("$(Format-MigrationCount $inactiveMailboxes.Count) inactive mailbox(es)") | Out-Null }
    $lines.Add("- **Exchange:** $($mailSummaryParts -join '; ').") | Out-Null

    $collaborationSummaryParts = New-Object System.Collections.Generic.List[string]
    $collaborationSummaryParts.Add("$(Format-MigrationCount $teamCount) Team(s)") | Out-Null
    $collaborationSummaryParts.Add("$(Format-MigrationCount $sharePointCount) SharePoint site(s)") | Out-Null
    $collaborationSummaryParts.Add("$(Format-MigrationCount $oneDriveCount) OneDrive site(s)") | Out-Null
    if ($null -ne $groupCount) { $collaborationSummaryParts.Add("$(Format-MigrationCount $groupCount) Microsoft 365 Group(s)") | Out-Null }
    $collaborationSummaryParts.Add("$(Format-MigrationDataSize $collaborationStorageGB) known, non-duplicated collaboration data") | Out-Null
    $lines.Add("- **Collaboration:** $($collaborationSummaryParts -join '; '). Teams file storage is already represented by its backing SharePoint site.") | Out-Null
    if ($publicFolders.Count -gt 0) {
        $lines.Add("- **Special content:** $(Format-MigrationCount $publicFolders.Count) public folder object(s), which require a separate migration decision and tool validation.") | Out-Null
    }

    $lines.Add('') | Out-Null
    $lines.Add('## Material complexity callouts') | Out-Null
    $lines.Add('') | Out-Null
    if ($callouts.Count -eq 0) {
        $lines.Add('No exceptional source-side complexity was surfaced by the collected datasets. This is not evidence about the destination tenant or unsupported data that was not collected.') | Out-Null
    }
    else {
        foreach ($callout in $callouts) {
            $lines.Add("- **$(Cell $callout.Title):** $(Cell $callout.Observation) **Suggested response:** $(Cell $callout.Response)") | Out-Null
        }
    }

    $lines.Add('') | Out-Null
    $lines.Add('## Suggested migration approach') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("**Recommended planning baseline:** $(Cell $recommendedMethod)") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add($methodNarrative) | Out-Null
    if ($phasedRecommended) {
        $lines.Add('') | Out-Null
        $lines.Add('Why phased delivery may help:') | Out-Null
        $lines.Add('') | Out-Null
        foreach ($reason in $phasedReasons) { $lines.Add("- $(Cell $reason)") | Out-Null }
    }
    $lines.Add('') | Out-Null
    $lines.Add('Methods to evaluate during design:') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('- **Single-event / big bang:** lowest coexistence duration and operational overhead; best when scope is small or simple and the validated cutover window can absorb the move.') | Out-Null
    $lines.Add('- **Phased / waves:** reduces per-event impact and allows measured learning; adds coexistence, routing, communications, and repeated validation effort.') | Out-Null
    $lines.Add('- **Microsoft-supported native path:** strong fit for lift-and-shift workloads that meet current native constraints and do not need copy-and-delta behavior.') | Out-Null
    $lines.Add('- **Commercial third-party or hybrid path:** useful for incremental staging, destination restructuring, broader shared-content coverage, or exceptions that native tooling does not cover.') | Out-Null

    $lines.Add('') | Out-Null
    $lines.Add('## Tooling and additional purchasing') | Out-Null
    $lines.Add('') | Out-Null
    if ($hasCollaborationScope) {
        $lines.Add("**Best-fit ShareGate tier from the source-only sizing proxy:** $(Cell $shareGatePlan)") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add($shareGateReason) | Out-Null
    }
    else {
        $lines.Add('No ShareGate subscription is indicated by the collaboration data collected in this run. Confirm that the zero scope is measured rather than a collection gap before removing it from the commercial plan.') | Out-Null
    }
    $shareGateStartingPrice = Number (Prop $shareGateSelectedCatalog @('StartingPriceUsdPerYear'))
    $shareGatePriceReference = if ($null -ne $shareGateStartingPrice) {
        " Vendor-listed starting price: USD $('{0:N0}' -f $shareGateStartingPrice)/year (catalog reviewed $toolingReferenceDate)."
    }
    else { '' }
    $shareGatePurchase = if ($hasCollaborationScope) { "Annual commercial subscription if not already owned. $shareGatePlan is the current best-fit tier.$shareGatePriceReference Obtain a current quote." } else { 'None indicated unless collaboration scope is later added.' }
    $lines.Add('') | Out-Null
    $lines.Add("**Purchase:** $(Cell $shareGatePurchase)") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('ShareGate is strongest when the project needs collaboration copy, incremental passes, restructuring, supported Teams/channel and Planner work, or an existing destination. Validate sharing-link fidelity, shared channels, and private-chat support. Pro or Enterprise is required if ShareGate will also handle Exchange Online, Entra ID preview, sensitivity labels, or parallel workstations.') | Out-Null

    $lines.Add('') | Out-Null
    $lines.Add('### Microsoft 365 native cross-tenant tools') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Migration Orchestrator currently covers Exchange mailboxes, OneDrive, Teams chats, and Teams meetings; native SharePoint is a separate workflow. Shared Teams/channel structure is outside the orchestrator scope. Verify preview/GA status and exact workload limits at project start.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('The native OneDrive and SharePoint paths are lift-and-shift moves without incremental delta or destination merge behavior; they are strongest when Microsoft-hosted moves and source-link redirects matter more than restructuring.') | Out-Null
    if ($teamCount -gt 0) {
        $lines.Add('') | Out-Null
        $lines.Add('If Teams chats or meetings and shared Teams/channel content are both required, evaluate a Microsoft-plus-ShareGate hybrid design and prove the fidelity and sequencing in the pilot.') | Out-Null
    }
    $lines.Add('') | Out-Null
    $lines.Add("**Purchase:** $(Cell $nativeUserLicenseBasis) $(Cell $nativeSharePointLicenseBasis)") | Out-Null

    $lines.Add('') | Out-Null
    $lines.Add('### BitTitan MigrationWiz') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Commercial alternative for mailbox and archive copy/pre-stage projects, especially when a copy model is preferred over a native move. Collaboration entitlements should not duplicate ShareGate or Microsoft-native scope.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("**Purchase:** $(Cell $migrationWizBasis)") | Out-Null

    $fastTrackMinimum = Number (Prop $fastTrackCatalog @('MinimumUserLicenses'))
    if ($null -eq $fastTrackMinimum) { $fastTrackMinimum = 150 }
    $fastTrackFit = if ($null -ne $employeeBasis -and $employeeBasis -ge $fastTrackMinimum) { 'Ask Microsoft whether the project qualifies for the invitation-only FastTrack cross-tenant service. It covers Exchange, SharePoint, and OneDrive data movement, not project architecture, identities, Teams, or post-migration orchestration.' } else { "Not a planning assumption: the current service is invitation-only and requires a minimum purchase of $(Format-MigrationCount $fastTrackMinimum) Cross-Tenant User Data Migration licenses." }
    $lines.Add('') | Out-Null
    $lines.Add('### Microsoft FastTrack cross-tenant service') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add((Cell $fastTrackFit)) | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('**Purchase:** Eligibility review plus the qualifying Microsoft migration licenses; do not assume acceptance or coverage.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Procurement notes:') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('- Target Microsoft 365 workload licenses and capacity are required but cannot be sized reliably from a source-only discovery run.') | Out-Null
    $lines.Add('- ShareGate Migrate is the migration product. ShareGate Protect is a separate governance/security product and is not required solely to perform a migration.') | Out-Null
    $lines.Add('- Microsoft native user-data and shared-data migration SKUs are separate from the ordinary target workload licenses. Confirm purchasing channel, eligibility, and assignment rules with Microsoft.') | Out-Null
    $lines.Add("- Tooling catalog last reviewed: $toolingReferenceDate. $toolingReviewPolicy") | Out-Null

    $lines.Add('') | Out-Null
    $lines.Add('## Suggested project milestones') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('These are suggested billable outcome points, not progress states inferred by the assessment script.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('| Milestone | Billable outcome / acceptance evidence |') | Out-Null
    $lines.Add('|---|---|') | Out-Null
    $lines.Add('| Assessment readout and approach selection | Source scope, material exceptions, delivery method, and planning assumptions reviewed with stakeholders. |') | Out-Null
    $lines.Add('| Solution design and tooling procurement | Workload architecture, fidelity requirements, tool path, license quantities, responsibilities, and commercial purchases approved. |') | Out-Null
    $lines.Add('| Target foundation and cutover preparation | Target identities, licenses, domains, coexistence, security, communications, rollback, and test prerequisites implemented. |') | Out-Null
    $lines.Add('| Pilot migration and acceptance | Representative pilot completed; throughput, fidelity, user experience, exception process, and go/no-go criteria validated. |') | Out-Null
    if ($phasedRecommended) {
        $lines.Add('| Production batch acceptance | Each agreed production batch completed and validated; detailed batch membership is established during the project, not by this assessment. |') | Out-Null
    }
    elseif ($hasSizingBasis) {
        $lines.Add('| Production cutover acceptance | Single production migration and domain/cutover activities completed with agreed validation evidence. |') | Out-Null
    }
    else {
        $lines.Add('| Production execution acceptance | Delivery structure is selected during solution design; the agreed production event or batches are completed with validation evidence. |') | Out-Null
    }
    $lines.Add('| Post-migration validation and hypercare exit | Data reconciliation, remediation, user support, and agreed stabilization criteria completed. |') | Out-Null
    $lines.Add('| Source decommission and compliance closeout | Retention, legal, rollback, license-recovery, and source shutdown obligations completed or handed off. |') | Out-Null

    $lines.Add('') | Out-Null
    $lines.Add('## Project planning inputs to confirm later') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('- Destination identity, tenant/domain ownership, target object mappings, and conflict handling.') | Out-Null
    $lines.Add('- Workload fidelity requirements: permissions, sharing links, Teams chats/channels, Planner, apps/workflows, labels, holds, archives, and public folders.') | Out-Null
    $lines.Add('- Business cutover window, coexistence duration, blackout dates, communications, support model, rollback, and acceptance criteria.') | Out-Null
    $lines.Add('- Final in-scope users and objects, exclusions, target licensing/capacity, vendor quote, and procurement lead time.') | Out-Null

    $lines.Add('') | Out-Null
    $lines.Add('## Current vendor guidance') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Re-check these primary sources during solution design and procurement:') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("- [Microsoft 365 tenant-to-tenant Migration Orchestrator]($orchestratorUrl)") | Out-Null
    $lines.Add("- [Microsoft cross-tenant mailbox migration]($nativeMailboxUrl)") | Out-Null
    $lines.Add('- [Microsoft cross-tenant OneDrive migration](https://learn.microsoft.com/en-us/microsoft-365/migration/cross-tenant-onedrive-migration?view=o365-worldwide)') | Out-Null
    $lines.Add("- [Microsoft cross-tenant SharePoint migration]($nativeSharePointUrl)") | Out-Null
    $lines.Add("- [Microsoft FastTrack cross-tenant migration service]($fastTrackUrl)") | Out-Null
    $lines.Add("- [ShareGate Migrate pricing and plan comparison]($shareGatePricingUrl)") | Out-Null
    $lines.Add("- [BitTitan migration license selection]($bitTitanLicenseUrl)") | Out-Null

    $parentDirectory = Split-Path -Path $Path -Parent
    if ($parentDirectory -and -not (Test-Path -Path $parentDirectory)) {
        New-Item -Path $parentDirectory -ItemType Directory -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, ($lines -join [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
}
