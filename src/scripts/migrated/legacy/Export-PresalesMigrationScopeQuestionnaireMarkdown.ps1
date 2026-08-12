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

    function HasArchive($Mailbox) {
        $explicit = Prop $Mailbox @('HasActiveArchive', 'HasArchive')
        if ($explicit -eq $true) { return $true }
        if ((Text $explicit '') -match '^(?i:true|yes|active|enabled)$') { return $true }
        if ((Text (Prop $Mailbox @('ArchiveStatus')) '') -match '(?i)active|enabled') { return $true }
        $archiveGuid = Text (Prop $Mailbox @('ArchiveGuid')) ''
        if ($archiveGuid -and $archiveGuid -notmatch '^0{8}-0{4}-0{4}-0{4}-0{12}$') { return $true }
        $archiveSize = Number (Prop $Mailbox @('ArchiveSizeGB', 'ArchiveGB'))
        if ($null -ne $archiveSize -and $archiveSize -gt 0) { return $true }
        $archiveDeletedSize = Number (Prop $Mailbox @('ArchiveDeletedItemsGB'))
        return ($null -ne $archiveDeletedSize -and $archiveDeletedSize -gt 0)
    }

    function ArchiveState($Mailbox) {
        if (HasArchive $Mailbox) { return 'Yes' }
        $archiveStatus = Text (Prop $Mailbox @('ArchiveStatus')) ''
        if (
            [string]::IsNullOrWhiteSpace($archiveStatus) -or
            $archiveStatus -match '^(?i:unknown|unavailable|not available|not collected|n/?a)(?:\b|$)'
        ) {
            return 'Unknown - validate'
        }

        # A collected, non-unknown status with no positive archive evidence is a defensible
        # negative. Do not infer "No" from placeholder statuses or missing collection data.
        return 'No'
    }

    function GroupMailboxEvidenceStatus($GroupRow) {
        $explicit = Text (Prop $GroupRow @('MailboxEvidenceStatus', 'MailboxSizingStatus', 'MailboxDataStatus')) ''
        if ($explicit) { return $explicit }

        $mailboxSize = Number (Prop $GroupRow @('MailboxSizeGB', 'SizeGB'))
        if ($null -eq $mailboxSize) { return 'Unknown / needs data' }
        return 'Measured'
    }

    function TeamEnrichmentEvidence($TeamRow) {
        $normalizeStatus = {
            param([string]$ExplicitStatus, [bool]$HasPositiveEvidence)
            if ($ExplicitStatus -match '(?i)needs?\s+data|unknown|unavailable|failed|not\s+(?:collected|requested)|partial') { return 'Needs Data' }
            if ($ExplicitStatus -match '(?i)measured\s+data') { return 'Measured data' }
            if ($ExplicitStatus -match '(?i)measured\s+empty') { return 'Measured empty' }
            if ($ExplicitStatus -match '(?i)measured|collected|complete|success') {
                return $(if ($HasPositiveEvidence) { 'Measured data' } else { 'Measured empty' })
            }
            if ($HasPositiveEvidence) { return 'Measured data' }
            return 'Needs Data'
        }

        $channelPositive = $false
        foreach ($name in @('TotalChannels', 'PublicChannelCount', 'PrivateChannelCount', 'SharedChannelCount')) {
            $count = Number (Prop $TeamRow @($name))
            if ($null -ne $count -and $count -gt 0) { $channelPositive = $true; break }
        }
        if (-not $channelPositive) {
            foreach ($name in @('PublicChannels', 'PrivateChannels', 'SharedChannels')) {
                if (-not [string]::IsNullOrWhiteSpace((Text (Prop $TeamRow @($name)) ''))) { $channelPositive = $true; break }
            }
        }

        $memberPositive = $false
        foreach ($name in @('MemberCount', 'GuestCount')) {
            $count = Number (Prop $TeamRow @($name))
            if ($null -ne $count -and $count -gt 0) { $memberPositive = $true; break }
        }

        [pscustomobject]@{
            Team = $TeamRow
            ChannelStatus = & $normalizeStatus (Text (Prop $TeamRow @('ChannelInventoryStatus', 'ChannelEvidenceStatus')) '') $channelPositive
            MemberStatus = & $normalizeStatus (Text (Prop $TeamRow @('MemberInventoryStatus', 'MemberEvidenceStatus')) '') $memberPositive
            HasSharedChannels = $channelPositive -and (
                ((Number (Prop $TeamRow @('SharedChannelCount'))) -gt 0) -or
                -not [string]::IsNullOrWhiteSpace((Text (Prop $TeamRow @('SharedChannels')) ''))
            )
        }
    }

    function StatusSummary([object[]]$InputRows) {
        if ($InputRows.Count -eq 0) { return 'No rows available.' }
        $groups = @(
            $InputRows | ForEach-Object {
                Text (Prop $_ @('Status', 'MappingStatus', 'DecisionStatus', 'ReadinessStatus')) 'Unspecified'
            } | Group-Object | Sort-Object Name
        )
        return (($groups | ForEach-Object { "$($_.Name): $($_.Count)" }) -join '; ')
    }

    $questions = New-Object System.Collections.Generic.List[object]
    $decisionLookup = @{}
    function Q {
        param(
            [string]$Id,
            [string]$Section,
            [string]$Question,
            [string]$Evidence,
            [string]$Source,
            [ValidateSet('High', 'Medium', 'Low')][string]$Confidence = 'Low',
            [string]$Tool = '',
            [string]$Assumption = ''
        )
        $decision = if ($decisionLookup.ContainsKey($Id)) { $decisionLookup[$Id] } else { $null }
        $decisionPrompt = Text (Prop $decision @('DecisionPrompt')) ''
        $decisionTool = Text (Prop $decision @('MigrationTool')) ''
        $decisionNotes = Text (Prop $decision @('Notes')) ''
        $mergedAssumption = @($Assumption, $decisionNotes) |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Select-Object -Unique
        $questions.Add([pscustomobject][ordered]@{
                Id = $Id
                Section = $Section
                # Planning-register DecisionPrompt is the canonical cross-artifact contract.
                # The local question remains a fallback for legacy snapshots without a seeded
                # decision table.
                Question = $(if ($decisionPrompt) { $decisionPrompt } else { $Question })
                Evidence = $Evidence
                Source = $Source
                Confidence = $Confidence
                CustomerConfirmedValue = Text (Prop $decision @('CustomerConfirmedValue')) ''
                DecisionStatus = Text (Prop $decision @('Status')) ''
                InScope = Text (Prop $decision @('InScope')) ''
                TargetMapping = Text (Prop $decision @('TargetMapping')) ''
                Tool = $(if ($decisionTool) { $decisionTool } else { $Tool })
                Assumption = ($mergedAssumption -join ' ')
                Owner = Text (Prop $decision @('DecisionOwner')) ''
                DueDate = Text (Prop $decision @('DueDate')) ''
            }) | Out-Null
    }

    $tenantInfo = if ($TenantStatsHash.ContainsKey('TenantInfo')) { $TenantStatsHash['TenantInfo'] } else { $null }
    $tenantName = Text (Prop $tenantInfo @('DisplayName', 'TenantName')) 'Unknown tenant'
    $tenantId = Text (Prop $tenantInfo @('TenantId', 'Id')) 'Not collected'
    $defaultDomain = Text (Prop $tenantInfo @('DefaultDomain', 'InitialDomain')) 'Not collected'

    $scope = @(Rows 'MigrationScopeSummary')
    $licenseSummary = @(Rows 'BitTitanLicenseSummary')
    $licenseDetail = @(Rows 'BitTitanLicenseDetail')
    $shareGate = @(Rows 'ShareGateScopeSummary')
    $groupReconciliation = @(Rows 'GroupWorkloadReconciliation')
    $complexity = @(Rows 'MigrationComplexityFlags')
    $quoteRows = @(Rows 'MigrationQuoteReadiness')
    $targetRows = @(Rows 'MigrationTargetReadiness')
    $identityRows = @(Rows 'MigrationIdentityMapping')
    $domainRows = @(Rows 'MigrationDomainDependencies')
    $dispositionRows = @(Rows 'MigrationObjectDisposition')
    $effortRows = @(Rows 'MigrationWorkloadEffort')
    $waveRows = @(Rows 'MigrationWavePlan')
    $decisionRows = @(Rows 'MigrationScopeDecisions')
    foreach ($decisionRow in $decisionRows) {
        $decisionKey = Text (Prop $decisionRow @('DecisionKey', 'Id')) ''
        if ($decisionKey) { $decisionLookup[$decisionKey] = $decisionRow }
    }

    $mailboxes = @(Rows 'AllMailboxes')
    if ($mailboxes.Count -eq 0) { $mailboxes = @(Rows 'MailboxFullDetails') }
    $inactive = @(Rows 'InactiveMailboxDetails')
    if ($inactive.Count -eq 0) {
        $inactive = @($mailboxes | Where-Object {
                (Prop $_ @('IsInactiveMailbox')) -eq $true -or (Text (Prop $_ @('IsInactiveMailbox')) '') -match '^(?i:true|yes)$'
            })
    }
    $inactiveWithArchive = @($inactive | Where-Object { HasArchive $_ })
    $inactiveUnknownArchive = @($inactive | Where-Object { (ArchiveState $_) -match '^Unknown' })
    $inactiveArchiveGB = 0.0
    foreach ($mailbox in $inactiveWithArchive) {
        $size = Number (Prop $mailbox @('ArchiveSizeGB', 'ArchiveGB'))
        if ($null -ne $size) { $inactiveArchiveGB += $size }
    }
    $inactiveArchiveGB = [math]::Round($inactiveArchiveGB, 2)

    $groups = @(Rows 'GroupMailboxes')
    $teams = @(Rows 'AllTeams')
    $sites = @(Rows 'SharePoint')
    $oneDrives = @(Rows 'OneDrive')
    $publicFolders = @(Rows 'PublicFolderDetails')
    if ($publicFolders.Count -eq 0) { $publicFolders = @(Rows 'PublicFolders') }
    $domains = @(Rows 'Domains')
    $teamEvidence = @($teams | ForEach-Object { TeamEnrichmentEvidence $_ })
    $unknownChannelTeams = @($teamEvidence | Where-Object { $_.ChannelStatus -eq 'Needs Data' })
    $unknownMemberTeams = @($teamEvidence | Where-Object { $_.MemberStatus -eq 'Needs Data' })
    $sharedChannels = @($teamEvidence | Where-Object { $_.HasSharedChannels } | ForEach-Object { $_.Team })
    if ($sharedChannels.Count -eq 0) {
        $sharedChannels = @($groupReconciliation | Where-Object {
                (Text (Prop $_ @('HasSharedChannels')) '') -match '^(?i:true|yes)$'
            })
    }

    $readiness = 'ROM'
    foreach ($row in $quoteRows) {
        $candidate = Text (Prop $row @('ReadinessLevel', 'QuoteReadiness')) ''
        if ($candidate -in @('ROM', 'Conditional', 'Firm')) { $readiness = $candidate; break }
    }
    if ($readiness -eq 'Firm' -and @(Rows 'MigrationScopeDecisions').Count -eq 0) { $readiness = 'Conditional' }
    # MigrationQuoteReadiness already projects blocking complexity flags. Treat it as the
    # authoritative count when present; adding MigrationComplexityFlags again double-counts
    # those same dependencies. Complexity is only a legacy-snapshot fallback.
    $blockingRows = if ($quoteRows.Count -gt 0) {
        @($quoteRows | Where-Object {
                (Text (Prop $_ @('Status', 'ReadinessStatus')) '') -match '^(?i:blocker|blocked|needs data|missing)$'
            })
    }
    else {
        @($complexity | Where-Object {
                (Text (Prop $_ @('Status', 'ReadinessStatus')) '') -match '^(?i:blocker|blocked|needs data|missing)$'
            })
    }

    $groupMailMeasuredData = @($groupReconciliation | Where-Object {
            $status = GroupMailboxEvidenceStatus $_
            $mailboxSize = Number (Prop $_ @('MailboxSizeGB', 'SizeGB'))
            $status -eq 'Measured data' -or ($status -eq 'Measured' -and $null -ne $mailboxSize -and $mailboxSize -gt 0)
        }).Count
    $groupMailMeasuredEmpty = @($groupReconciliation | Where-Object {
            $status = GroupMailboxEvidenceStatus $_
            $mailboxSize = Number (Prop $_ @('MailboxSizeGB', 'SizeGB'))
            $status -eq 'Measured empty' -or ($status -eq 'Measured' -and $null -ne $mailboxSize -and $mailboxSize -le 0)
        }).Count
    $groupMailNeedsData = @($groupReconciliation | Where-Object {
            (GroupMailboxEvidenceStatus $_) -match '(?i)unknown|unavailable|needs\s+data|not\s+collected|missing'
        }).Count

    $licenseEvidence = if ($licenseSummary.Count -gt 0) {
        @($licenseSummary | Select-Object -First 10 | ForEach-Object {
                "$(Text (Prop $_ @('Metric', 'Item')) 'Metric')=$(Text (Prop $_ @('Value', 'Count')) 'Unknown')"
            }) -join '; '
    }
    elseif ($licenseDetail.Count -gt 0) { "$($licenseDetail.Count) BitTitan object planning row(s)." }
    else { 'No BitTitan licensing table was present.' }

    Q 'QR-01' 'Quote readiness' 'Confirm the quote-readiness level and resolve every blocker before changing it.' "$readiness; blocker or missing-data rows: $($blockingRows.Count); scope metrics: $($scope.Count)." 'MigrationQuoteReadiness; MigrationComplexityFlags; MigrationScopeSummary' $(if ($quoteRows.Count -gt 0) { 'High' } else { 'Medium' }) '' 'Firm requires resolved blockers plus merged, customer-confirmed scope and target decisions.'
    Q 'QR-02' 'Quote readiness' 'Confirm pricing will be applied separately by the solution engineer.' 'This artifact contains quantities and licensing guidance only; no costs or prices are included.' 'Presales output policy' 'High' '' 'Current vendor pricing is supplied outside this artifact.'

    Q 'BT-01' 'BitTitan mail scope and licensing' 'Confirm the mailbox objects that will be migrated with BitTitan.' $licenseEvidence 'BitTitanLicenseSummary; BitTitanLicenseDetail; AllMailboxes' $(if ($licenseDetail.Count -gt 0) { 'High' } else { 'Low' }) 'BitTitan MigrationWiz' 'Group mailbox licensing is not included automatically; group mailbox content requires an explicit disposition.'
    Q 'BT-02' 'BitTitan mail scope and licensing' 'Confirm inactive mailboxes will be restored or recovered before a MigrationWiz job is submitted.' "$($inactive.Count) inactive mailbox(es); $($inactiveWithArchive.Count) with an archive signal; $($inactiveUnknownArchive.Count) with unknown archive state; known archive data: $inactiveArchiveGB GB." 'InactiveMailboxDetails; AllMailboxes; BitTitanLicenseDetail' $(if ($inactiveUnknownArchive.Count -eq 0) { 'High' } else { 'Medium' }) 'Restore/recover in Exchange Online, then migrate with BitTitan' 'Inactive mailboxes are not migrated in-place. Validate restoration, target provisioning, hold implications, and archives.'
    Q 'BT-03' 'BitTitan mail scope and licensing' 'Confirm every inactive mailbox archive will be restored and migrated as a separate archive workload where required.' "See the inactive mailbox appendix. Archive-present: $($inactiveWithArchive.Count); archive state unknown: $($inactiveUnknownArchive.Count)." 'InactiveMailboxDetails archive fields' $(if ($inactiveUnknownArchive.Count -eq 0) { 'High' } else { 'Medium' }) 'BitTitan mailbox and archive projects as applicable' 'Unknown archive state is not treated as no archive.'
    Q 'BT-04' 'BitTitan mail scope and licensing' 'Confirm the disposition of Microsoft 365 group mailbox conversations.' "$($groupReconciliation.Count) reconciled group workload row(s); measured mail data: $groupMailMeasuredData; measured empty: $groupMailMeasuredEmpty; needs data: $groupMailNeedsData." 'GroupWorkloadReconciliation; GroupMailboxes' $(if ($groupReconciliation.Count -gt 0 -and $groupMailNeedsData -eq 0) { 'High' } elseif ($groupReconciliation.Count -gt 0) { 'Medium' } else { 'Low' }) 'Customer decision; BitTitan only when explicitly approved' 'Arraya normally avoids an automatic group-mailbox license allowance. Decide migrate, retain, archive, or exclude per group.'
    Q 'BT-05' 'BitTitan mail scope and licensing' 'Select the applicable BitTitan mail licensing approach.' $licenseEvidence 'BitTitanLicenseSummary; BitTitanLicenseDetail' $(if ($licenseSummary.Count -gt 0) { 'High' } else { 'Low' }) 'Mailbox Migration license and/or User Migration Bundle' 'Archive-enabled or multi-workload users may require UMB. Validate current terms; pricing remains separate.'
    Q 'BT-06' 'BitTitan mail scope and licensing' 'Should Tenant Migration Bundles be evaluated as an optional alternative?' "$($licenseDetail.Count) mail planning row(s), $($oneDrives.Count) OneDrive site(s), $($teams.Count) Team(s), and $($sites.Count) SharePoint site(s)." 'BitTitanLicenseDetail; OneDrive; AllTeams; SharePoint' 'Medium' 'Optional BitTitan Tenant Migration Bundle' 'TMB is not the default. Evaluate only when combined entitlements map cleanly and do not duplicate ShareGate.'
    Q 'BT-07' 'BitTitan mail scope and licensing' 'If TMB is selected, map each Flex Collaboration entitlement to one Team or one SharePoint library and validate the 100 GB allowance.' "$($teams.Count) Team(s) and $($sites.Count) SharePoint site(s); library-level mappings and sizes require confirmation." 'AllTeams; SharePoint; ShareGateScopeSummary' 'Medium' 'TMB Flex Collaboration entitlement' 'Each FCL maps to one Team or one SharePoint library, up to 100 GB. Additional entitlements may be required.'
    Q 'BT-08' 'BitTitan mail scope and licensing' 'Confirm public folders are outside TMB and have a separate disposition and licensing plan.' "$($publicFolders.Count) public folder record(s)." 'PublicFolderDetails; PublicFolders' $(if ($publicFolders.Count -gt 0) { 'High' } else { 'Medium' }) 'Separate public-folder plan' 'TMB does not migrate public folders.'
    Q 'BT-09' 'BitTitan mail scope and licensing' 'Validate source MigrationWiz authentication, organization/user EWS access, and the EWS application allow-list.' 'Not confirmed by discovery. Validate organization and mailbox EWSEnabled, modern-auth permissions, and the applicable MigrationWiz application ID in EWSAllowedAppIDs.' 'Source Exchange Online and Entra configuration' 'Low' 'MigrationWiz source endpoint' 'Allow up to 24 hours for Exchange configuration propagation; validate the endpoint before pilot.'
    Q 'BT-10' 'BitTitan mail scope and licensing' 'Validate destination MigrationWiz authentication, organization/user EWS access, and the EWS application allow-list.' 'Destination was not assessed. Validate organization and mailbox EWSEnabled, modern-auth permissions, and the applicable MigrationWiz application ID in EWSAllowedAppIDs.' 'Destination Exchange Online and Entra configuration' 'Low' 'MigrationWiz destination endpoint' 'Phased EWS blocking begins 2026-10-01; final retirement is listed for 2027-04-01. Re-check current guidance before execution.'

    Q 'SG-01' 'ShareGate collaboration scope' 'Confirm ShareGate remains the default tool for Teams, SharePoint, and OneDrive content.' "$($shareGate.Count) ShareGate metric(s); $($teams.Count) Team(s); $($sites.Count) SharePoint site(s); $($oneDrives.Count) OneDrive site(s)." 'ShareGateScopeSummary; AllTeams; SharePoint; OneDrive' $(if ($shareGate.Count -gt 0) { 'High' } else { 'Medium' }) 'ShareGate' 'Remove any collaboration workload assigned to TMB from ShareGate quantities.'
    Q 'SG-02' 'ShareGate collaboration scope' 'Confirm target mappings for Teams, non-Team groups, and group-connected sites.' "$($groupReconciliation.Count) group reconciliation row(s)." 'GroupWorkloadReconciliation; AllTeams; SharePoint' $(if ($groupReconciliation.Count -gt 0) { 'High' } else { 'Low' }) 'ShareGate plus manual provisioning as needed' 'Group conversations and group-connected sites are distinct workloads and must not be double-counted.'
    $sg03Evidence = if ($unknownChannelTeams.Count -gt 0 -or $unknownMemberTeams.Count -gt 0) {
        "Needs Data: channel inventory unavailable for $($unknownChannelTeams.Count) of $($teams.Count) Team(s); member/guest inventory unavailable for $($unknownMemberTeams.Count); $($sharedChannels.Count) Team/group row(s) have a measured shared-channel signal."
    }
    else {
        "$($sharedChannels.Count) Team/group row(s) have a shared-channel signal; channel and member inventories were measured for $($teams.Count) Team(s)."
    }
    Q 'SG-03' 'ShareGate collaboration scope' 'Review shared channels and unsupported or manually rebuilt collaboration features.' $sg03Evidence 'AllTeams; GroupWorkloadReconciliation; MigrationComplexityFlags' $(if ($teams.Count -eq 0) { 'Low' } elseif ($unknownChannelTeams.Count -gt 0 -or $unknownMemberTeams.Count -gt 0) { 'Medium' } else { 'High' }) 'ShareGate plus exception plan' 'Missing Team enrichment is not treated as zero. Confirm current product support before committing scope.'
    Q 'SG-04' 'ShareGate collaboration scope' 'Validate ShareGate source and destination connectivity, required admin roles, application consent, and representative test access.' 'Not confirmed by source discovery. Validate both tenant connections and run a representative source-to-target access test.' 'Source and destination ShareGate connection validation' 'Low' 'ShareGate source and destination endpoints' 'Complete this validation before the pilot; the Arraya assessment application and certificate do not prove ShareGate endpoint readiness.'

    Q 'ID-01' 'Identity and target readiness' 'Confirm destination tenant identity, domain, geography, access, and migration ownership.' "Source: $tenantName; tenant ID: $tenantId; default domain: $defaultDomain. Destination not assessed." 'TenantInfo; customer destination details' 'High' '' 'This assessment contains source-tenant evidence.'
    Q 'ID-02' 'Identity and target readiness' 'Approve source-to-target UPN, SMTP, alias, anchor, and conflict mappings.' "$($identityRows.Count) mapping row(s); $(StatusSummary $identityRows)" 'MigrationIdentityMapping; Users; AllRecipients' $(if ($identityRows.Count -gt 0) { 'High' } else { 'Low' }) '' 'Source discovery cannot confirm target identities.'
    Q 'ID-03' 'Identity and target readiness' 'Confirm target provisioning, licenses, mailbox capacity, and application consent.' "$($targetRows.Count) readiness row(s); $(StatusSummary $targetRows)" 'MigrationTargetReadiness; destination evidence' $(if ($targetRows.Count -gt 0) { 'Medium' } else { 'Low' }) '' 'Endpoint and pilot validation are required before a firm quote.'
    Q 'ID-04' 'Identity and target readiness' 'Approve a disposition for every included and excluded object.' "$($dispositionRows.Count) disposition row(s); $(StatusSummary $dispositionRows)" 'MigrationObjectDisposition' $(if ($dispositionRows.Count -gt 0) { 'High' } else { 'Low' }) '' 'Unconfirmed objects remain open scope.'

    Q 'DC-01' 'Domain, cutover, and coexistence' 'Approve domain disposition, DNS ownership, release order, and dependencies.' "$($domains.Count) domain row(s); $($domainRows.Count) dependency row(s); $(StatusSummary $domainRows)" 'Domains; MigrationDomainDependencies' $(if ($domainRows.Count -gt 0) { 'High' } else { 'Medium' }) '' 'Destination readiness and DNS access require confirmation.'
    Q 'DC-02' 'Domain, cutover, and coexistence' 'Select big-bang or waves and confirm coexistence, forwarding, and routing.' "$($waveRows.Count) wave row(s); $($mailboxes.Count) mailbox row(s)." 'MigrationWavePlan; AllMailboxes; mail flow tables' $(if ($waveRows.Count -gt 0) { 'Medium' } else { 'Low' }) 'BitTitan coexistence/cutover as approved' 'MigrationWiz is not bidirectional live synchronization.'
    Q 'DC-03' 'Domain, cutover, and coexistence' 'Confirm DNS TTL, MX/Autodiscover/SPF/DKIM/DMARC, freeze, rollback, and communications ownership.' 'Not fully discoverable from source assessment.' 'Customer messaging, DNS, security, and change owners' 'Low' '' 'Execution decisions are not inferred from source configuration.'

    Q 'CP-01' 'Compliance and data handling' 'Confirm legal, hold, retention, eDiscovery, and records requirements before restoring inactive mailboxes.' "$($inactive.Count) inactive mailbox(es); $($inactiveWithArchive.Count) with archive signal." 'InactiveMailboxDetails; retention and hold fields' 'Medium' '' 'Customer legal/records owners must approve restoration and disposition.'
    Q 'CP-02' 'Compliance and data handling' 'Confirm treatment for encrypted, recoverable, oversized, corrupt, and unsupported data.' "$($complexity.Count) complexity row(s); $(StatusSummary $complexity)" 'MigrationComplexityFlags; BitTitanLicenseSummary' $(if ($complexity.Count -gt 0) { 'Medium' } else { 'Low' }) '' 'Define exception and acceptance thresholds before production.'
    Q 'CP-03' 'Compliance and data handling' 'Confirm destination retention, sensitivity, DLP, sharing, and residency requirements.' 'Destination compliance configuration was not assessed.' 'Destination Purview, SharePoint, Teams, and tenant configuration' 'Low' '' 'Source inventory cannot validate destination governance.'

    Q 'PM-01' 'Project constraints, effort, and waves' 'Confirm blackout dates, completion date, work hours, change windows, and decision owners.' 'Not collected by tenant discovery.' 'Customer and project stakeholders' 'Low' '' 'Open dates and owners keep readiness at ROM or Conditional.'
    Q 'PM-02' 'Project constraints, effort, and waves' 'Review workload effort bands and validate their assumptions.' "$($effortRows.Count) effort row(s); $(StatusSummary $effortRows)" 'MigrationWorkloadEffort; MigrationComplexityFlags' $(if ($effortRows.Count -gt 0) { 'Medium' } else { 'Low' }) '' 'Effort is provisional until decisions, mappings, tools, and exceptions are merged.'
    Q 'PM-03' 'Project constraints, effort, and waves' 'Approve pilot criteria, waves, migration passes, acceptance, hypercare, and rollback.' "$($waveRows.Count) wave row(s); $(StatusSummary $waveRows)" 'MigrationWavePlan; MigrationWorkloadEffort' $(if ($waveRows.Count -gt 0) { 'Medium' } else { 'Low' }) '' 'Wave guidance is not a committed schedule until dependencies are approved.'

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# Microsoft 365 Tenant-to-Tenant Migration Scope Confirmation') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("**Client:** $(Cell $tenantName)") | Out-Null
    $lines.Add("**Source Tenant ID:** $(Cell $tenantId)") | Out-Null
    $lines.Add("**Assessment Date:** $($AsOfDate.ToString('yyyy-MM-dd'))") | Out-Null
    $lines.Add("**Prepared By:** $(Cell $PreparedBy)") | Out-Null
    $lines.Add("**Current Quote Readiness:** $readiness") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('This working questionnaire separates source-tenant discovery from customer decisions. Treat **Discovered value** as evidence, not approval. Complete the confirmation, status, scope, mapping, owner, and date fields. Existing values are merged from MigrationScopeDecisions by DecisionKey. Pricing is intentionally omitted.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## Quote-readiness scale') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('| Level | Meaning |') | Out-Null
    $lines.Add('|---|---|') | Out-Null
    $lines.Add('| ROM | Source discovery supports an initial range; target evidence, decisions, or blockers remain open. |') | Out-Null
    $lines.Add('| Conditional | Major quantities and prerequisites are known; named assumptions or exceptions remain. |') | Out-Null
    $lines.Add('| Firm | Source and target evidence is present, blockers are resolved, and customer decisions and mappings are merged. |') | Out-Null

    foreach ($section in @('Quote readiness', 'BitTitan mail scope and licensing', 'ShareGate collaboration scope', 'Identity and target readiness', 'Domain, cutover, and coexistence', 'Compliance and data handling', 'Project constraints, effort, and waves')) {
        $lines.Add('') | Out-Null
        $lines.Add("## $section") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add('| Decision key | Scope question / decision | Discovered value | Evidence source | Confidence | Customer-confirmed value | Decision status | In scope? | Target mapping / approach | Tool / method | Assumption / exception | Owner | Due date |') | Out-Null
        $lines.Add('|---|---|---|---|---|---|---|---|---|---|---|---|---|') | Out-Null
        foreach ($question in @($questions | Where-Object { $_.Section -eq $section })) {
            $cells = @($question.Id, $question.Question, $question.Evidence, $question.Source, $question.Confidence, $question.CustomerConfirmedValue, $question.DecisionStatus, $question.InScope, $question.TargetMapping, $question.Tool, $question.Assumption, $question.Owner, $question.DueDate) |
                ForEach-Object { Cell $_ '' }
            $lines.Add('| ' + ($cells -join ' | ') + ' |') | Out-Null
        }
    }

    $lines.Add('') | Out-Null
    $lines.Add('## Optional Tenant Migration Bundle decision guide') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Arraya''s default planning path remains BitTitan for approved mailbox/archive workloads and ShareGate for collaboration. TMB is an optional licensing and delivery alternative, not an automatic recommendation.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('| Evaluate TMB when... | Keep the default split-tool approach when... | Guardrail |') | Out-Null
    $lines.Add('|---|---|---|') | Out-Null
    $lines.Add('| A qualifying user needs UMB and there is a clearly mapped Team or SharePoint library for the paired FCL. | ShareGate is selected, FCLs would be stranded, or collaboration mappings are unfinished. | Current TMB guidance states one UMB plus one FCL; each FCL maps to one Team or SharePoint library up to 100 GB. |') | Out-Null
    $lines.Add('| Consolidated MigrationWiz licensing and operations has a validated project benefit. | The project is mailbox-only or TMB would duplicate ShareGate scope. | Public folders are excluded. Teams Private Chat (PCH) uses a separate Collaboration (Private Chats) project; validate current capability, limitations, and licensing entitlement rather than assuming TMB/UMB inclusion or exclusion. |') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Official guidance to re-check at quote and execution time:') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('- [BitTitan: Which Migration License Do I Need?](https://help.bittitan.com/hc/en-us/articles/36203678661531-Which-Migration-License-Do-I-Need)') | Out-Null
    $lines.Add('- [BitTitan: Getting Started with Migrations](https://help.bittitan.com/hc/en-us/articles/1260806785629-Getting-Started-with-Migrations)') | Out-Null
    $lines.Add('- [BitTitan: Teams Private Chat Migration Guide](https://help.bittitan.com/hc/en-us/articles/25603590557979-Teams-Private-Chat-Migration-Guide)') | Out-Null
    $lines.Add('- [BitTitan: tenant-to-tenant mailbox migration with coexistence](https://help.bittitan.com/hc/en-us/articles/360045004254-Exchange-Online-Microsoft-365-to-Exchange-Online-Microsoft-365-Migration-Guide-Using-Coexistence-Different-Domain)') | Out-Null

    $lines.Add('') | Out-Null
    $lines.Add('## Inactive mailbox and archive confirmation') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Inactive mailboxes are assumed to be restored/recovered before migration. Archive state is explicit: unknown must be validated, and an existing archive may require its own workload and licensing decision.') | Out-Null
    $lines.Add('') | Out-Null
    if ($inactive.Count -eq 0) {
        $lines.Add('_No inactive mailbox records were present._') | Out-Null
    }
    else {
        $lines.Add('| Decision key | Display name | Primary SMTP | Recipient type | Mailbox GB | Deleted GB | Archive present | Archive status | Archive GB | Archive deleted GB | Total GB | Hold / retention | Migration prerequisite | Customer-confirmed value | In scope? | Decision status | Target mailbox | Tool / method | Owner / notes |') | Out-Null
        $lines.Add('|---|---|---|---|---:|---:|---|---|---:|---:|---:|---|---|---|---|---|---|---|---|') | Out-Null
        foreach ($mailbox in @($inactive | Sort-Object { Text (Prop $_ @('PrimarySmtpAddress', 'UserPrincipalName', 'DisplayName')) '' })) {
            $hold = Text (Prop $mailbox @('LitigationHoldEnabled')) 'Unknown'
            $retention = Text (Prop $mailbox @('RetentionPolicy')) 'Unknown'
            $identity = Text (Prop $mailbox @('PrimarySmtpAddress', 'UserPrincipalName', 'Identity', 'ExchangeGuid', 'DisplayName')) ''
            $mailboxDecisionKey = Text (Prop $mailbox @('DecisionKey')) ''
            if (-not $mailboxDecisionKey) { $mailboxDecisionKey = "MAILBOX:$($identity.Trim().ToLowerInvariant())" }
            $mailboxDecision = if ($decisionLookup.ContainsKey($mailboxDecisionKey)) { $decisionLookup[$mailboxDecisionKey] } else { $null }
            $decisionOwner = Text (Prop $mailboxDecision @('DecisionOwner')) ''
            $decisionNotes = Text (Prop $mailboxDecision @('Notes')) ''
            $ownerNotes = @($decisionOwner, $decisionNotes) |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Select-Object -Unique
            $cells = @(
                $mailboxDecisionKey,
                (Prop $mailbox @('DisplayName', 'Name')),
                (Prop $mailbox @('PrimarySmtpAddress', 'UserPrincipalName', 'Identity')),
                (Prop $mailbox @('RecipientTypeDetails', 'RecipientType')),
                (Prop $mailbox @('MailboxSizeGB', 'SizeGB')),
                (Prop $mailbox @('DeletedItemsGB')),
                (ArchiveState $mailbox),
                (Prop $mailbox @('ArchiveStatus')),
                (Prop $mailbox @('ArchiveSizeGB', 'ArchiveGB')),
                (Prop $mailbox @('ArchiveDeletedItemsGB')),
                (Prop $mailbox @('TotalDataToMigrateGB')),
                "Litigation hold: $hold; retention: $retention",
                'Restore/recover, validate hold and archive, provision target, then migrate',
                (Prop $mailboxDecision @('CustomerConfirmedValue')),
                (Prop $mailboxDecision @('InScope')),
                (Prop $mailboxDecision @('Status')),
                (Prop $mailboxDecision @('TargetMapping')),
                (Prop $mailboxDecision @('MigrationTool')),
                ($ownerNotes -join '; ')
            ) | ForEach-Object { Cell $_ '' }
            $lines.Add('| ' + ($cells -join ' | ') + ' |') | Out-Null
        }
    }

    $groupPlanningRows = New-Object System.Collections.Generic.List[object]
    $groupPlanningKeys = @{}
    foreach ($groupRow in @($groupReconciliation + $groups)) {
        $smtp = Text (Prop $groupRow @('PrimarySmtpAddress', 'PrimarySMTPAddress', 'Mail')) ''
        $displayName = Text (Prop $groupRow @('DisplayName', 'Name')) ''
        $dedupeKey = if ($smtp) { $smtp.ToLowerInvariant() } else { $displayName.ToLowerInvariant() }
        if (-not $dedupeKey -or $groupPlanningKeys.ContainsKey($dedupeKey)) { continue }
        $groupPlanningKeys[$dedupeKey] = $true
        $groupPlanningRows.Add($groupRow) | Out-Null
    }

    $lines.Add('') | Out-Null
    $lines.Add('## Group mailbox conversation disposition') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Group mailbox conversations are reported for scoping but are not included automatically in the BitTitan license allowance. Confirm the disposition of each group separately from its Team or SharePoint site content.') | Out-Null
    $lines.Add('') | Out-Null
    if ($groupPlanningRows.Count -eq 0) {
        $lines.Add('_No Microsoft 365 group mailbox records were present._') | Out-Null
    }
    else {
        $lines.Add('| Decision key | Display name | Primary SMTP | Team? | Mailbox GB | Mailbox evidence | Site GB | Classification | Default practice status | Migrate conversations? | Target group | Tool / method | Owner / notes |') | Out-Null
        $lines.Add('|---|---|---|---|---:|---|---:|---|---|---|---|---|---|') | Out-Null
        foreach ($groupRow in @($groupPlanningRows | Sort-Object { Text (Prop $_ @('PrimarySmtpAddress', 'DisplayName')) '' })) {
            $smtp = Text (Prop $groupRow @('PrimarySmtpAddress', 'PrimarySMTPAddress', 'Mail')) ''
            $displayName = Text (Prop $groupRow @('DisplayName', 'Name')) ''
            $groupDecisionKey = Text (Prop $groupRow @('DecisionKey')) ''
            if (-not $groupDecisionKey) {
                $groupDecisionKey = if ($smtp) { "GROUP:$($smtp.ToLowerInvariant())" } else { "GROUP:$($displayName.ToLowerInvariant())" }
            }
            $groupDecision = if ($decisionLookup.ContainsKey($groupDecisionKey)) { $decisionLookup[$groupDecisionKey] } else { $null }
            $migrateConversations = Text (Prop $groupDecision @('InScope', 'CustomerConfirmedValue')) ''
            $targetGroup = Text (Prop $groupDecision @('TargetMapping')) ''
            $groupTool = Text (Prop $groupDecision @('MigrationTool')) 'Customer decision; BitTitan only when approved'
            $decisionOwner = Text (Prop $groupDecision @('DecisionOwner')) ''
            $sourceNotes = Text (Prop $groupRow @('Notes')) ''
            $decisionNotes = Text (Prop $groupDecision @('Notes')) ''
            $mailboxSize = Number (Prop $groupRow @('MailboxSizeGB', 'SizeGB'))
            $mailboxEvidence = GroupMailboxEvidenceStatus $groupRow
            $mailboxSizeDisplay = if ($null -eq $mailboxSize) { 'Unknown / needs data' } else { $mailboxSize }
            $classification = Text (Prop $groupRow @('Classification')) ''
            if (
                $mailboxEvidence -match '(?i)unknown|unavailable|needs\s+data|not\s+collected|missing' -and
                $classification -notmatch '(?i)unknown|unavailable|needs\s+data'
            ) {
                $classification = if ($classification) { "$classification (mailbox evidence unknown)" } else { 'Mailbox evidence unknown' }
            }
            $ownerNotes = @($decisionOwner, $sourceNotes, $decisionNotes) |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Select-Object -Unique
            $cells = @(
                $groupDecisionKey,
                $displayName,
                $smtp,
                (Prop $groupRow @('IsTeam')),
                $mailboxSizeDisplay,
                $mailboxEvidence,
                (Prop $groupRow @('SiteStorageGB', 'StorageUsedGB', 'SiteSizeGB')),
                $classification,
                'Report only / customer decision',
                $migrateConversations,
                $targetGroup,
                $groupTool,
                ($ownerNotes -join '; ')
            ) | ForEach-Object { Cell $_ '' }
            $lines.Add('| ' + ($cells -join ' | ') + ' |') | Out-Null
        }
    }

    $openRows = @(@($quoteRows + $complexity) | Where-Object {
            (Text (Prop $_ @('Status', 'ReadinessStatus')) '') -match '(?i)block|review|open|needs|missing|conditional'
        })
    $lines.Add('') | Out-Null
    $lines.Add('## Open readiness items') | Out-Null
    $lines.Add('') | Out-Null
    if ($openRows.Count -eq 0) {
        $lines.Add('_No explicit blocker/review rows were present. This does not replace customer confirmation._') | Out-Null
    }
    else {
        $lines.Add('| Category | Item | Status | Value | Required action | Source worksheet | Owner | Due date |') | Out-Null
        $lines.Add('|---|---|---|---|---|---|---|---|') | Out-Null
        foreach ($row in $openRows) {
            $cells = @(
                (Prop $row @('Category', 'Section')),
                (Prop $row @('Item', 'Check', 'Metric')),
                (Prop $row @('Status', 'ReadinessStatus')),
                (Prop $row @('Value', 'DiscoveredValue')),
                (Prop $row @('MigrationAction', 'Resolution', 'RequiredAction', 'Notes')),
                (Prop $row @('SourceWorksheet', 'EvidenceSource')),
                '',
                ''
            ) | ForEach-Object { Cell $_ '' }
            $lines.Add('| ' + ($cells -join ' | ') + ' |') | Out-Null
        }
    }

    $parentDirectory = Split-Path -Path $Path -Parent
    if ($parentDirectory -and -not (Test-Path -Path $parentDirectory)) {
        New-Item -Path $parentDirectory -ItemType Directory -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, ($lines -join [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
}
