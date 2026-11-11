function New-ArrayaMigrationCriteriaRegister {
    <#
    .SYNOPSIS
        Builds the operator-readable migration criteria register for the requested profiles.

    .DESCRIPTION
        One row per composed criterion (see Resolve-ArrayaMigrationCriteria). Each row reports:

          CollectorState         - the detailed collector evidence state, from
                                   Get-ArrayaDatasetEvidenceState or the rule-specific helper.
          EvidenceStatus         - Ready, Partial, Failed, NeedsData, or NotApplicable. Only a
                                   complete collector state is Ready; Unknown, Unavailable,
                                   Unsupported, NotRequested, and SkippedByProfile are NeedsData.
          PresalesClassification - PresalesBlocker, NeedsData, PlanningGap, or Nonblocking.
          MigrationTreatment     - for the mailbox hold criterion only; empty otherwise.
          CurrentStatus          - Blocker, Needs Data, Needs Input, Review, Ready, Confirmed,
                                   or Not Applicable.
          ActionId and Action    - one action per distinct condition. Rows share an action only
                                   when kind, responsible role, scope, and decision or closure
                                   evidence are all the same.

        Presales classification rules:
          - Not applicable, or evidence ready with nothing open: Nonblocking.
          - Evidence missing for an optional criterion: Nonblocking (the gap is still shown).
          - Evidence missing and not automatically collectible: PlanningGap.
          - Evidence missing for a required criterion whose missing-evidence behavior is
            PresalesBlocker, with Presales among the composed profiles: PresalesBlocker.
          - Other missing required evidence: NeedsData (or PlanningGap when the catalog says so).
          - Evidence ready but a decision or hold treatment is still open: PlanningGap.

        Legal or compliance outcomes are never decided here. Decisions count as confirmed only
        when MigrationScopeDecisions carries the exact DecisionKey with an approved status.
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Builds and returns in-memory rows; it changes no system state.')]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$TenantStatsHash,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$OutputProfile,
        [Parameter(Mandatory = $false)]
        $Catalog
    )

    if ($null -eq $Catalog) { $Catalog = Get-ArrayaMigrationCriteriaCatalog }
    $criteria = @(Resolve-ArrayaMigrationCriteria -OutputProfile $OutputProfile -Catalog $Catalog)
    if ($criteria.Count -eq 0) { return @() }

    $toRows = {
        param($Value)
        if ($null -eq $Value) { return @() }
        if ($Value -is [System.Collections.IDictionary]) { return @($Value.Values | Where-Object { $null -ne $_ }) }
        if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) { return @($Value | Where-Object { $null -ne $_ }) }
        return @($Value)
    }
    $completeStates = @('CompleteWithRecords', 'CompleteEmpty')
    $stateRank = @{ CompleteWithRecords = 0; CompleteEmpty = 0; Partial = 1; Unknown = 2; NotRequested = 3; SkippedByProfile = 4; Unsupported = 5; Unavailable = 6; Failed = 7 }

    # ---- Shared evidence -----------------------------------------------------------------
    $decisionLookup = @{}
    foreach ($decision in @(& $toRows $(if ($TenantStatsHash.Contains('MigrationScopeDecisions')) { $TenantStatsHash['MigrationScopeDecisions'] } else { $null }))) {
        $key = [string](Get-ArrayaObjectValue -Object $decision -Names @('DecisionKey'))
        if (-not [string]::IsNullOrWhiteSpace($key)) { $decisionLookup[$key.Trim().ToLowerInvariant()] = $decision }
    }
    $isDecisionConfirmed = {
        param([string]$DecisionKey)
        if ([string]::IsNullOrWhiteSpace($DecisionKey)) { return $false }
        $decision = $decisionLookup[$DecisionKey.Trim().ToLowerInvariant()]
        if ($null -eq $decision) { return $false }
        return ([string](Get-ArrayaObjectValue -Object $decision -Names @('Status'))) -match '^(?i)(confirmed|approved|complete|completed|resolved)$'
    }

    $mailboxRows = if ($TenantStatsHash.Contains('MailboxFullDetails') -and @(& $toRows $TenantStatsHash['MailboxFullDetails']).Count -gt 0) {
        @(& $toRows $TenantStatsHash['MailboxFullDetails'])
    }
    else {
        @(& $toRows $(if ($TenantStatsHash.Contains('AllMailboxes')) { $TenantStatsHash['AllMailboxes'] } else { $null }))
    }
    $mailboxEvidence = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'AllMailboxes' -CollectorStepName 'Exchange mailboxes'
    $directorySync = Get-ArrayaDirectorySyncEvidence -TenantStatsHash $TenantStatsHash
    $exchangeHybrid = Get-ArrayaExchangeHybridEvidence -TenantStatsHash $TenantStatsHash
    $evidenceCache = @{}

    $newEvidence = {
        param([string]$CollectorState, [string]$Reason, [string]$Treatment, [string]$Notes, [string]$EvidenceStatusOverride)
        [pscustomobject]@{ CollectorState = $CollectorState; Reason = $Reason; Treatment = $Treatment; Notes = $Notes; EvidenceStatus = $EvidenceStatusOverride }
    }
    $fromDatasetState = {
        param($State, [string]$Prefix)
        $collectorState = if ($State.CollectorState -in $completeStates) { 'Unknown' } else { [string]$State.CollectorState }
        & $newEvidence $collectorState ('{0}{1}' -f $Prefix, $State.Reason) '' '' ''
    }

    $evaluate = {
        param($Criterion)

        if ($Criterion.Applicability -eq 'WhenDirectorySynchronized') {
            if ($directorySync.State -eq 'NotSynchronized') {
                return (& $newEvidence 'NotApplicable' '' '' ("Not applicable: {0}" -f $directorySync.Reason) '')
            }
            if ($directorySync.State -eq 'Unknown') {
                return (& $newEvidence 'Unknown' ("Applicability not established: directory synchronization state is unknown. {0}" -f $directorySync.Reason) '' '' '')
            }
        }

        switch ($Criterion.EvidenceRule) {
            'Unsupported' {
                return (& $newEvidence 'Unsupported' 'No collector in this assessment gathers this evidence; it must come from the customer or a vendor.' '' '' '')
            }
            'Dataset' {
                $states = @(foreach ($datasetKey in $Criterion.Datasets) {
                        Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey $datasetKey -CollectorStepName $Criterion.CollectorSteps -EmptyIsUnexpected:([bool]$Criterion.EmptyIsUnexpected)
                    })
                $incomplete = @($states | Where-Object { $_.CollectorState -notin $completeStates })
                if ($incomplete.Count -eq 0) {
                    $combined = if (@($states | Where-Object { $_.CollectorState -eq 'CompleteWithRecords' }).Count -gt 0) { 'CompleteWithRecords' } else { 'CompleteEmpty' }
                    $rowText = (@($states | ForEach-Object { '{0}: {1} row(s)' -f $_.DatasetKey, $_.RowCount }) -join '; ')
                    return (& $newEvidence $combined '' '' $rowText '')
                }
                $worst = @($incomplete | Sort-Object @{ Expression = { $stateRank[[string]$_.CollectorState] }; Descending = $true } | Select-Object -First 1)[0]
                return (& $newEvidence ([string]$worst.CollectorState) ((@($incomplete | ForEach-Object { '{0}: {1}' -f $_.DatasetKey, $_.Reason }) -join ' ')) '' '' '')
            }
            'DirectorySync' {
                if ($directorySync.State -in @('Synchronized', 'NotSynchronized')) {
                    return (& $newEvidence 'CompleteWithRecords' '' '' ("{0}: {1}" -f $directorySync.State, $directorySync.Reason) '')
                }
                $state = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'AdConnectConfiguration' -CollectorStepName $Criterion.CollectorSteps
                $collectorState = if ($state.CollectorState -in $completeStates) { 'Unknown' } else { [string]$state.CollectorState }
                return (& $newEvidence $collectorState ("Directory synchronization state not established. {0}" -f $directorySync.Reason) '' '' '')
            }
            'ExchangeHybrid' {
                if ($exchangeHybrid.State -in @('Configured', 'Indicators', 'NotConfigured')) {
                    return (& $newEvidence 'CompleteWithRecords' '' '' ("{0}: {1}" -f $exchangeHybrid.State, $exchangeHybrid.Reason) '')
                }
                if ($null -ne $exchangeHybrid.Details) {
                    $hybridStatus = [string](Get-ArrayaObjectValue -Object $exchangeHybrid.Details -Names @('HybridStatus'))
                    $collectorState = if ($hybridStatus -eq 'PermissionDenied') { 'Unavailable' } else { 'Partial' }
                    return (& $newEvidence $collectorState $exchangeHybrid.Reason '' '' '')
                }
                $state = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'HybridConfiguration' -CollectorStepName $Criterion.CollectorSteps
                return (& $fromDatasetState $state 'Exchange hybrid state not established. ')
            }
            'MailboxHolds' {
                if (-not $evidenceCache.ContainsKey('MailboxHolds')) {
                    $evidenceCache['MailboxHolds'] = Get-ArrayaMailboxHoldTreatment -TenantStatsHash $TenantStatsHash -Mailbox $mailboxRows -Catalog $Catalog
                }
                $hold = $evidenceCache['MailboxHolds']
                $collectorState = switch ($hold.EvidenceStatus) {
                    'Ready' { 'CompleteWithRecords' }
                    'Partial' { 'Partial' }
                    'Failed' { 'Failed' }
                    'NotApplicable' { 'NotApplicable' }
                    default { if ($mailboxEvidence.CollectorState -in $completeStates) { 'Unknown' } else { [string]$mailboxEvidence.CollectorState } }
                }
                return (& $newEvidence $collectorState $hold.MissingEvidence $hold.MigrationTreatment $hold.Summary ([string]$hold.EvidenceStatus))
            }
            'MailboxDelegates' {
                if ($mailboxRows.Count -eq 0) {
                    if ($mailboxEvidence.CollectorState -eq 'CompleteEmpty') {
                        return (& $newEvidence 'NotApplicable' '' '' 'No mailboxes were found.' '')
                    }
                    return (& $fromDatasetState $mailboxEvidence 'Mailbox inventory not collected. ')
                }
                $stateFields = @('FullAccessDelegateState', 'SendAsDelegateState', 'CalendarDelegateState')
                $values = @(foreach ($row in $mailboxRows) { foreach ($field in $stateFields) { [string](Get-ArrayaObjectValue -Object $row -Names @($field)) } })
                $collected = @($values | Where-Object { $_ -eq 'Collected' }).Count
                $notRequested = @($values | Where-Object { [string]::IsNullOrWhiteSpace($_) -or $_ -eq 'NotCollected' }).Count
                if ($collected -eq $values.Count) {
                    return (& $newEvidence 'CompleteWithRecords' '' '' ("Delegate permissions collected for {0} mailbox(es)." -f $mailboxRows.Count) '')
                }
                if ($notRequested -eq $values.Count) {
                    return (& $newEvidence 'NotRequested' 'Delegate permissions were not requested in this run. They are collected only when TenantToTenantMigration runs on its own.' '' '' '')
                }
                return (& $newEvidence 'Partial' ("Delegate permission state is incomplete: {0} of {1} mailbox permission lookups collected." -f $collected, $values.Count) '' '' '')
            }
            'MailboxSizing' {
                if ($mailboxRows.Count -eq 0) {
                    if ($mailboxEvidence.CollectorState -eq 'CompleteEmpty') {
                        return (& $newEvidence 'NotApplicable' '' '' 'No mailboxes were found.' '')
                    }
                    return (& $fromDatasetState $mailboxEvidence 'Mailbox inventory not collected. ')
                }
                $sized = @($mailboxRows | Where-Object { $null -ne (Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $_ -Names @('MailboxSizeGB', 'TotalItemSizeGB', 'TotalDataToMigrateGB'))) }).Count
                if ($sized -eq $mailboxRows.Count) {
                    return (& $newEvidence 'CompleteWithRecords' '' '' ("Size statistics for {0} mailbox(es)." -f $mailboxRows.Count) '')
                }
                $collectorState = if ($sized -gt 0) { 'Partial' } else { 'Unknown' }
                return (& $newEvidence $collectorState ("Size statistics missing for {0} of {1} mailbox(es)." -f ($mailboxRows.Count - $sized), $mailboxRows.Count) '' '' '')
            }
        }

        return (& $newEvidence 'Unknown' ("Evidence rule '{0}' is not implemented." -f $Criterion.EvidenceRule) '' '' '')
    }

    # ---- Rows ----------------------------------------------------------------------------
    $rows = New-Object System.Collections.Generic.List[object]
    $actionIds = @{}
    $actionCounter = 0
    foreach ($criterion in $criteria) {
            $evidence = & $evaluate $criterion
            $evidenceStatus = if (-not [string]::IsNullOrWhiteSpace($evidence.EvidenceStatus)) {
                $evidence.EvidenceStatus
            }
            else {
                switch ($evidence.CollectorState) {
                    { $_ -in $completeStates } { 'Ready' }
                    'Partial' { 'Partial' }
                    'Failed' { 'Failed' }
                    'NotApplicable' { 'NotApplicable' }
                    default { 'NeedsData' }
                }
            }
            $evidenceReady = ($evidenceStatus -eq 'Ready')
            $isUnsupported = ($criterion.Coverage -eq 'Unsupported' -or $evidence.CollectorState -eq 'Unsupported')
            $presalesInvolved = (@($criterion.Profile) + @($criterion.ContributingProfiles)) -contains 'Presales'
            $treatment = [string]$evidence.Treatment
            $treatmentNeedsDecision = $treatment -in @('RequiresPreservationPlan', 'RequiresLegalDecision', 'RequiresComplianceValidation', 'BlocksMigration')
            $decisionConfirmed = & $isDecisionConfirmed $criterion.DecisionKey
            $decisionOpen = -not [string]::IsNullOrWhiteSpace($criterion.DecisionKey) -and -not $decisionConfirmed

            $classification = if ($evidenceStatus -eq 'NotApplicable') {
                'Nonblocking'
            }
            elseif ($evidenceReady) {
                if ($treatment -eq 'BlocksMigration' -or ($treatmentNeedsDecision -and -not $decisionConfirmed) -or $decisionOpen) { 'PlanningGap' } else { 'Nonblocking' }
            }
            elseif ($treatment -eq 'BlocksMigration') {
                'PlanningGap'
            }
            elseif ($criterion.Requirement -eq 'Optional') {
                'Nonblocking'
            }
            elseif ($isUnsupported) {
                'PlanningGap'
            }
            elseif ($criterion.MissingEvidenceBehavior -eq 'PresalesBlocker' -and $presalesInvolved) {
                'PresalesBlocker'
            }
            elseif ($criterion.MissingEvidenceBehavior -eq 'PlanningGap') {
                'PlanningGap'
            }
            else {
                'NeedsData'
            }

            $currentStatus = if ($evidenceStatus -eq 'NotApplicable') {
                'Not Applicable'
            }
            elseif ($treatment -eq 'BlocksMigration' -or $classification -eq 'PresalesBlocker') {
                'Blocker'
            }
            elseif (-not $evidenceReady) {
                if ($isUnsupported) { 'Needs Input' } else { 'Needs Data' }
            }
            elseif ($decisionConfirmed) {
                'Confirmed'
            }
            elseif ($decisionOpen) {
                'Needs Input'
            }
            elseif ($treatmentNeedsDecision) {
                'Review'
            }
            else {
                'Ready'
            }

            $actionKind = if ($currentStatus -in @('Ready', 'Confirmed', 'Not Applicable')) {
                ''
            }
            elseif ($treatment -eq 'BlocksMigration') {
                'Resolve'
            }
            elseif (-not $evidenceReady) {
                if ($isUnsupported) { 'Obtain' } else { 'Collect' }
            }
            else {
                'Decide'
            }

            $actionId = ''
            $actionText = ''
            if (-not [string]::IsNullOrWhiteSpace($actionKind)) {
                $closure = if ($actionKind -in @('Collect', 'Obtain')) { $criterion.ResolutionEvidence } else { '{0}|{1}' -f $criterion.DecisionKey, $criterion.RequiredDecision }
                $actionKey = (@($actionKind, $criterion.ResponsibleRole, $criterion.Scope, $closure) | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() }) -join '|'
                if (-not $actionIds.ContainsKey($actionKey)) {
                    $actionCounter++
                    $actionIds[$actionKey] = 'ACT-{0:D3}' -f $actionCounter
                }
                $actionId = $actionIds[$actionKey]
                $decisionLabel = if ([string]::IsNullOrWhiteSpace($criterion.DecisionKey)) { 'Record the decision' } else { 'Record decision {0}' -f $criterion.DecisionKey }
                $decisionText = if ([string]::IsNullOrWhiteSpace($criterion.RequiredDecision)) { $criterion.Criterion } else { $criterion.RequiredDecision }
                $actionText = switch ($actionKind) {
                    'Collect' { 'Collect: {0}.' -f $criterion.ResolutionEvidence }
                    'Obtain' { 'Obtain from the customer or vendor: {0}.' -f $criterion.ResolutionEvidence }
                    'Resolve' { 'Change the migration method or the hold for the affected mailboxes before they move; {0}: {1}.' -f $decisionLabel.ToLowerInvariant(), $decisionText }
                    default { '{0}: {1}.' -f $decisionLabel, $decisionText }
                }
            }

            $missingEvidence = if ($evidenceReady -or $evidenceStatus -eq 'NotApplicable') { '' } else { [string]$evidence.Reason }
            $notes = New-Object System.Collections.Generic.List[string]
            if (-not [string]::IsNullOrWhiteSpace($evidence.Notes)) { $notes.Add([string]$evidence.Notes) | Out-Null }
            if (-not $evidenceReady -and $criterion.Requirement -eq 'Optional') { $notes.Add('Optional evidence: its absence is shown but does not block.') | Out-Null }
            if ($evidenceReady -and $treatmentNeedsDecision) { $notes.Add('Ready evidence is not migration clearance; the treatment above must be resolved.') | Out-Null }
            if ($treatment -eq 'RequiresLegalDecision') { $notes.Add('A legal owner must decide the disposition; the assessment does not.') | Out-Null }

            $source = if ($isUnsupported) {
                'Customer or vendor input (no collector)'
            }
            else {
                $parts = @()
                if (@($criterion.CollectorSteps).Count -gt 0) { $parts += ('Collector: {0}' -f (@($criterion.CollectorSteps) -join ', ')) }
                if (@($criterion.Datasets).Count -gt 0) { $parts += ('Datasets: {0}' -f (@($criterion.Datasets) -join ', ')) }
                $parts -join '; '
            }

            $rows.Add([pscustomobject][ordered]@{
                CriterionId            = $criterion.CriterionId
                SourceCriterionIds     = (@($criterion.SourceCriterionIds) -join '; ')
                Profile                = $criterion.Profile
                ContributingProfiles   = (@($criterion.ContributingProfiles) -join '; ')
                Criterion              = $criterion.Criterion
                Workloads              = (@($criterion.Workloads) -join '; ')
                Requirement            = $criterion.Requirement
                Phase                  = (@($criterion.Phase) -join '; ')
                Applicability          = $criterion.Applicability
                Impact                 = $criterion.Impact
                CollectorOrSource      = $source
                CollectorState         = [string]$evidence.CollectorState
                EvidenceStatus         = $evidenceStatus
                MissingEvidence        = $missingEvidence
                PresalesClassification = $classification
                MigrationTreatment     = $treatment
                RequiredDecision       = $criterion.RequiredDecision
                DecisionKey            = $criterion.DecisionKey
                BlockerImpact          = $(if ($classification -eq 'PresalesBlocker' -or -not $evidenceReady) { $criterion.BlockerImpact } else { '' })
                ResolutionEvidence     = $criterion.ResolutionEvidence
                ResponsibleRole        = $criterion.ResponsibleRole
                CurrentStatus          = $currentStatus
                ActionId               = $actionId
                Action                 = $actionText
                Notes                  = (@($notes.ToArray()) -join ' ')
            }) | Out-Null
    }

    return @($rows.ToArray())
}
