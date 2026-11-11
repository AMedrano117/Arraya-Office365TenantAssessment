function Get-ArrayaMailboxHoldTreatment {
    <#
    .SYNOPSIS
        Separates mailbox hold evidence from the migration treatment it calls for.

    .DESCRIPTION
        Two independent answers:

        EvidenceStatus - how complete the hold evidence is.
          Ready         - every source in the evidence scope is complete: all four hold fields
                          on every mailbox, the Purview retention policy inventory, or both
                          (see EvidenceScope).
          Partial       - mailboxes were collected, but some hold fields or the retention
                          policy inventory are missing (or, for OrganizationPolicies, the
                          retention collection was partial).
          NeedsData     - mailbox inventory was not collected.
          Failed        - mailbox collection failed.
          NotApplicable - mailbox collection completed and found no mailboxes.

        MigrationTreatment - what the evidence requires before migration.
          Clear                        - no hold found in the evaluated sources.
          RequiresPreservationPlan     - delay hold or retention hold only.
          RequiresComplianceValidation - retention-policy holds (mailbox or organization-wide).
          RequiresLegalDecision        - litigation hold or eDiscovery hold. A legal owner must
                                         decide; the assessment does not.
          BlocksMigration              - a held mailbox has a confirmed migration method that a
                                         HoldMethodRules entry documents as incompatible.
          NotEstablished               - evidence is incomplete.

        Hold presence does not make evidence unready, and Ready evidence does not mean the
        mailboxes are clear to migrate. A block confirmed for specific mailboxes is reported even
        when evidence for other mailboxes is incomplete, because that block is itself evidenced.

        Compliance tag holds and eDiscovery case detail are not collected and never evaluated.

        In-place hold identifiers are classified by their documented prefix: 'UniH' and bare
        GUIDs are eDiscovery holds; 'mbx', 'skp', 'grp', and 'cld' are retention-policy holds;
        entries starting with '-' are exclusions and are ignored.

    .PARAMETER Mailbox
        Mailbox rows to evaluate. Defaults to MailboxFullDetails, then AllMailboxes.

    .PARAMETER Catalog
        Criteria catalog that holds HoldMethodRules. Defaults to the shipped catalog.

    .PARAMETER EvidenceScope
        Which hold sources to evaluate.
          All                  - mailbox hold fields and organization-wide retention policies
                                 together (the default; used by the cutover summary).
          MailboxFields        - mailbox hold fields only. Evidence is Ready from a complete
                                 Exchange mailbox collection; Purview is not required.
          OrganizationPolicies - organization-wide Exchange retention policies only. Evidence is
                                 Ready from a complete Purview retention policy collection.
        The two partial scopes let a profile that does not collect Purview still establish
        mailbox-level hold evidence, while the organization-wide gap stays visible on its own.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$TenantStatsHash,
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Mailbox,
        [Parameter(Mandatory = $false)]
        $Catalog,
        [Parameter(Mandatory = $false)]
        [ValidateSet('All', 'MailboxFields', 'OrganizationPolicies')]
        [string]$EvidenceScope = 'All'
    )

    $includeMailboxFields = $EvidenceScope -in @('All', 'MailboxFields')
    $includeOrganizationPolicies = $EvidenceScope -in @('All', 'OrganizationPolicies')

    $toRows = {
        param($Value)
        if ($null -eq $Value) { return @() }
        if ($Value -is [System.Collections.IDictionary]) { return @($Value.Values | Where-Object { $null -ne $_ }) }
        if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) { return @($Value | Where-Object { $null -ne $_ }) }
        return @($Value)
    }

    if (-not $PSBoundParameters.ContainsKey('Mailbox')) {
        $Mailbox = if ($TenantStatsHash.Contains('MailboxFullDetails') -and @(& $toRows $TenantStatsHash['MailboxFullDetails']).Count -gt 0) {
            @(& $toRows $TenantStatsHash['MailboxFullDetails'])
        }
        else {
            @(& $toRows $(if ($TenantStatsHash.Contains('AllMailboxes')) { $TenantStatsHash['AllMailboxes'] } else { $null }))
        }
    }
    $rows = @($Mailbox | Where-Object { $null -ne $_ })
    if ($null -eq $Catalog) { $Catalog = Get-ArrayaMigrationCriteriaCatalog }

    $mailboxEvidence = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'AllMailboxes' -CollectorStepName 'Exchange mailboxes'
    $retentionEvidence = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'RetentionPolicies' -CollectorStepName 'Purview retention/DLP policies'
    $holdEvidence = Get-ArrayaMailboxHoldEvidence -Mailbox $rows

    # ---- Organization-wide retention policies ---------------------------------------------
    $retentionComplete = $retentionEvidence.CollectorState -in @('CompleteWithRecords', 'CompleteEmpty')
    $orgWidePolicyNames = @()
    if ($includeOrganizationPolicies -and $retentionComplete) {
        $orgWidePolicyNames = @(
            & $toRows $TenantStatsHash['RetentionPolicies'] | Where-Object {
                $enabled = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $_ -Names @('Enabled'))
                $location = [string](Get-ArrayaObjectValue -Object $_ -Names @('ExchangeLocation'))
                $enabled -ne $false -and $location -match '(?i)(^|[;,\s])All($|[;,\s])'
            } | ForEach-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('PolicyName', 'Name')) }
        )
    }

    # ---- Per-mailbox hold types ------------------------------------------------------------
    $decisionRows = @(& $toRows $(if ($TenantStatsHash.Contains('MigrationScopeDecisions')) { $TenantStatsHash['MigrationScopeDecisions'] } else { $null }))
    $confirmedMethods = @{}
    foreach ($decision in $decisionRows) {
        $key = [string](Get-ArrayaObjectValue -Object $decision -Names @('DecisionKey'))
        $status = [string](Get-ArrayaObjectValue -Object $decision -Names @('Status'))
        $tool = [string](Get-ArrayaObjectValue -Object $decision -Names @('MigrationTool'))
        if ($key -match '^(?i)MAILBOX:' -and $status -match '^(?i)(confirmed|approved|complete|completed|resolved)$' -and -not [string]::IsNullOrWhiteSpace($tool)) {
            $confirmedMethods[$key.Trim().ToLowerInvariant()] = $tool
        }
    }
    $methodRules = @(Get-ArrayaObjectValue -Object $Catalog -Names @('HoldMethodRules') | Where-Object { $null -ne $_ })

    $counts = [ordered]@{ LitigationHold = 0; eDiscoveryHold = 0; RetentionPolicyHold = 0; DelayHold = 0; RetentionHold = 0 }
    $blockingMailboxes = New-Object System.Collections.Generic.List[string]
    $blockingRuleIds = New-Object System.Collections.Generic.List[string]
    foreach ($row in $rows) {
        $types = New-Object System.Collections.Generic.List[string]
        if ($includeMailboxFields) {
            if ((Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $row -Names @('LitigationHoldEnabled'))) -eq $true) { $types.Add('LitigationHold') | Out-Null }
            if ((Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $row -Names @('DelayHoldApplied'))) -eq $true) { $types.Add('DelayHold') | Out-Null }
            if ((Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $row -Names @('RetentionHoldEnabled'))) -eq $true) { $types.Add('RetentionHold') | Out-Null }

            $inPlaceValue = Get-ArrayaObjectValue -Object $row -Names @('InPlaceHolds')
            $entries = if ($null -eq $inPlaceValue) { @() } elseif ($inPlaceValue -is [string]) { @($inPlaceValue -split '[,;]') } else { @($inPlaceValue | ForEach-Object { [string]$_ }) }
            foreach ($entry in @($entries | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith('-') })) {
                $type = if ($entry -match '^(?i)UniH' -or $entry -match '^[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}$') { 'eDiscoveryHold' } else { 'RetentionPolicyHold' }
                if (-not $types.Contains($type)) { $types.Add($type) | Out-Null }
            }
        }
        if ($orgWidePolicyNames.Count -gt 0 -and -not $types.Contains('RetentionPolicyHold')) { $types.Add('RetentionPolicyHold') | Out-Null }

        foreach ($type in $types) { $counts[$type]++ }

        $identity = [string](Get-ArrayaObjectValue -Object $row -Names @('PrimarySmtpAddress', 'UserPrincipalName', 'ExchangeGuid'))
        $methodKey = ('MAILBOX:{0}' -f $identity.Trim().ToLowerInvariant())
        if ($types.Count -gt 0 -and $confirmedMethods.ContainsKey($methodKey)) {
            $method = [string]$confirmedMethods[$methodKey]
            foreach ($rule in $methodRules) {
                $patterns = @(Get-ArrayaObjectValue -Object $rule -Names @('MethodPatterns'))
                $blockedTypes = @(Get-ArrayaObjectValue -Object $rule -Names @('BlockedHoldTypes'))
                $methodMatches = @($patterns | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) -and $method.IndexOf([string]$_, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count -gt 0
                $typeMatches = @($types | Where-Object { $blockedTypes -contains $_ }).Count -gt 0
                if ($methodMatches -and $typeMatches) {
                    if (-not $blockingMailboxes.Contains($identity)) { $blockingMailboxes.Add($identity) | Out-Null }
                    $ruleId = [string](Get-ArrayaObjectValue -Object $rule -Names @('RuleId'))
                    if (-not $blockingRuleIds.Contains($ruleId)) { $blockingRuleIds.Add($ruleId) | Out-Null }
                }
            }
        }
    }

    # ---- Evidence status -------------------------------------------------------------------
    # Each entry is one sentence, so joined entries never run together.
    $missing = New-Object System.Collections.Generic.List[string]
    $addMissing = {
        param([string]$Text)
        $sentence = ([string]$Text).Trim()
        if ([string]::IsNullOrWhiteSpace($sentence)) { return }
        if ($sentence -notmatch '[.!?]$') { $sentence += '.' }
        $missing.Add($sentence) | Out-Null
    }
    $retentionGap = "Organization-wide retention policies were not evaluated: $($retentionEvidence.Reason)"
    $evidenceStatus = 'Ready'
    if ($EvidenceScope -eq 'OrganizationPolicies') {
        $evidenceStatus = switch ($retentionEvidence.CollectorState) {
            { $_ -in @('CompleteWithRecords', 'CompleteEmpty') } { 'Ready' }
            'Partial' { 'Partial' }
            { $_ -in @('Failed', 'Unavailable') } { 'Failed' }
            default { 'NeedsData' }
        }
        if ($evidenceStatus -ne 'Ready') { & $addMissing $retentionGap }
    }
    elseif ($rows.Count -eq 0) {
        $evidenceStatus = switch ($mailboxEvidence.CollectorState) {
            'CompleteEmpty' { 'NotApplicable' }
            { $_ -in @('Failed', 'Unavailable') } { 'Failed' }
            default { 'NeedsData' }
        }
        if ($evidenceStatus -ne 'NotApplicable') { & $addMissing "Mailbox inventory: $($mailboxEvidence.Reason)" }
    }
    else {
        if (@($holdEvidence.UnavailableSignals).Count -gt 0) {
            $evidenceStatus = 'Partial'
            & $addMissing ('Mailbox hold fields not collected for every mailbox: {0}' -f (@($holdEvidence.UnavailableSignals) -join ', '))
        }
        if ($includeOrganizationPolicies -and -not $retentionComplete) {
            $evidenceStatus = 'Partial'
            & $addMissing $retentionGap
        }
    }

    # ---- Migration treatment ---------------------------------------------------------------
    $treatment = if ($blockingMailboxes.Count -gt 0) {
        'BlocksMigration'
    }
    elseif ($evidenceStatus -eq 'NotApplicable') {
        'Clear'
    }
    elseif ($evidenceStatus -ne 'Ready') {
        'NotEstablished'
    }
    elseif ($EvidenceScope -eq 'OrganizationPolicies') {
        # An organization-wide policy places every mailbox on a retention-policy hold, whether
        # or not the mailbox inventory was collected.
        if ($orgWidePolicyNames.Count -gt 0) { 'RequiresComplianceValidation' } else { 'Clear' }
    }
    elseif ($counts.LitigationHold -gt 0 -or $counts.eDiscoveryHold -gt 0) {
        'RequiresLegalDecision'
    }
    elseif ($counts.RetentionPolicyHold -gt 0) {
        'RequiresComplianceValidation'
    }
    elseif ($counts.DelayHold -gt 0 -or $counts.RetentionHold -gt 0) {
        'RequiresPreservationPlan'
    }
    else {
        'Clear'
    }

    $breakdown = 'litigation hold: {0}; eDiscovery hold: {1}; retention-policy hold: {2}; delay hold: {3}; retention hold (retention processing suspended): {4}' -f $counts.LitigationHold, $counts.eDiscoveryHold, $counts.RetentionPolicyHold, $counts.DelayHold, $counts.RetentionHold
    $notEvaluated = 'Compliance tag holds and eDiscovery case detail are not evaluated.'
    $orgWideText = if ($orgWidePolicyNames.Count -gt 0) { ' Organization-wide Exchange retention policies: {0}.' -f ($orgWidePolicyNames -join ', ') } else { '' }
    $blockText = if ($blockingMailboxes.Count -gt 0) { ' {0} held mailbox(es) have a confirmed migration method that does not support holds ({1}).' -f $blockingMailboxes.Count, ($blockingRuleIds -join ', ') } else { '' }
    $summary = switch ($EvidenceScope) {
        'OrganizationPolicies' {
            $policyText = if ($orgWidePolicyNames.Count -gt 0) {
                'Organization-wide Exchange retention policies: {0}.' -f ($orgWidePolicyNames -join ', ')
            }
            elseif ($evidenceStatus -eq 'Ready') {
                'No enabled retention policy applies to all Exchange mailboxes.'
            }
            else {
                'Organization-wide retention policies not evaluated.'
            }
            ('{0}{1} Mailbox-level holds are reported separately.' -f $policyText, $blockText).Trim()
        }
        'MailboxFields' {
            ('{0}.{1} Organization-wide retention policies are reported separately. {2}' -f $breakdown, $blockText, $notEvaluated).Trim()
        }
        default {
            ('{0}.{1}{2} {3}' -f $breakdown, $orgWideText, $blockText, $notEvaluated).Trim()
        }
    }

    return [pscustomobject][ordered]@{
        EvidenceScope             = $EvidenceScope
        EvidenceStatus            = $evidenceStatus
        MigrationTreatment        = $treatment
        MailboxCount              = $rows.Count
        MailboxesWithHoldIndicators = [int]$holdEvidence.AnyHoldCount
        HoldTypeCounts            = [pscustomobject]$counts
        OrganizationWidePolicies  = $orgWidePolicyNames
        BlockingMailboxCount      = $blockingMailboxes.Count
        BlockingRuleIds           = @($blockingRuleIds.ToArray())
        EvaluatedSignals          = @($holdEvidence.EvaluatedSignals)
        UnavailableSignals        = @($holdEvidence.UnavailableSignals)
        MissingEvidence           = (@($missing.ToArray()) -join ' ')
        DecisionKey               = 'CP-01'
        Summary                   = $summary
    }
}
