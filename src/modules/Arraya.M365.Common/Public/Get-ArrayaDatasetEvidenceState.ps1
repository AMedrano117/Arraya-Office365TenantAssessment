function Get-ArrayaDatasetEvidenceState {
    <#
    .SYNOPSIS
        Tells whether a dataset in the assessment store was actually collected.

    .DESCRIPTION
        A count of zero only means "none exist" when the data behind it was collected
        completely. This function returns two views of the same evidence:

        CollectorState, the detailed state:

          CompleteWithRecords - rows exist.
          CompleteEmpty       - the collector answered with no rows, and completeness is known
                                from a completed collector step or the collector's own record.
          Partial             - the collector recorded that only part of the source answered.
          Failed              - the collector step failed, or the collector recorded a failure.
          Unavailable         - the source could not be reached (permission, session, missing
                                dependency).
          NotRequested        - collector results exist, but no step produced this dataset.
          Unsupported         - the dataset is not supported in this collection mode.
          SkippedByProfile    - the output profile skipped the collector step.
          Unknown             - nothing establishes whether the dataset was collected.

        State, the summary kept for existing callers:

          Collected (CompleteWithRecords, CompleteEmpty, or rows present), Failed (Failed,
          Unavailable), NotCollected (NotRequested, Unsupported, SkippedByProfile), Unknown.

        Inputs, in order of authority: rows in the dataset; the collector's record in
        CollectorEvidenceStates (see Set-ArrayaCollectorEvidenceState); the matching
        CollectorPlanResults step; and presence of the dataset. A present but empty dataset
        with no collector record and no step result is Unknown: nothing shows the collection
        was complete.

    .PARAMETER CollectorStepName
        Name, or leading part of the name, of the collector step(s) that produce the dataset.
        Steps are also matched through their ProducedKeys.

    .PARAMETER EmptyIsUnexpected
        Treat a collected but empty dataset as Unknown. Use for datasets such as Domains or
        RemoteDomains where every tenant has at least one entry.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [System.Collections.IDictionary]$TenantStatsHash,
        [Parameter(Mandatory = $true)]
        [string]$DatasetKey,
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]]$CollectorStepName = @(),
        [Parameter(Mandatory = $false)]
        [switch]$EmptyIsUnexpected
    )

    $legacyStateMap = @{
        CompleteWithRecords = 'Collected'
        CompleteEmpty       = 'Collected'
        Partial             = 'Unknown'
        Failed              = 'Failed'
        Unavailable         = 'Failed'
        NotRequested        = 'NotCollected'
        Unsupported         = 'NotCollected'
        SkippedByProfile    = 'NotCollected'
        Unknown             = 'Unknown'
    }

    $newResult = {
        param([string]$CollectorState, [int]$RowCount, [string]$Reason, [string]$StepName, [string]$StepStatus, [string]$StateOverride)
        $state = if ([string]::IsNullOrWhiteSpace($StateOverride)) { $legacyStateMap[$CollectorState] } else { $StateOverride }
        [pscustomobject][ordered]@{
            DatasetKey      = $DatasetKey
            State           = $state
            CollectorState  = $CollectorState
            IsCollected     = ($state -eq 'Collected')
            RowCount        = $RowCount
            Reason          = $Reason
            CollectorStep   = $StepName
            CollectorStatus = $StepStatus
        }
    }

    if ($null -eq $TenantStatsHash) {
        return (& $newResult 'Unknown' 0 'No assessment data was supplied.' $null $null 'NotCollected')
    }

    $isPresent = $TenantStatsHash.Contains($DatasetKey) -and $null -ne $TenantStatsHash[$DatasetKey]
    $rowCount = 0
    if ($isPresent) {
        $value = $TenantStatsHash[$DatasetKey]
        if ($value -is [System.Collections.IDictionary]) {
            $rowCount = $value.Count
        }
        elseif (($value -is [System.Collections.IEnumerable]) -and -not ($value -is [string])) {
            $rowCount = @($value | Where-Object { $null -ne $_ }).Count
        }
        else {
            $rowCount = 1
        }
    }

    $record = $null
    if ($TenantStatsHash.Contains('CollectorEvidenceStates') -and ($TenantStatsHash['CollectorEvidenceStates'] -is [System.Collections.IDictionary])) {
        $recordContainer = $TenantStatsHash['CollectorEvidenceStates']
        if ($recordContainer.Contains($DatasetKey)) {
            $record = $recordContainer[$DatasetKey]
        }
    }
    $recordedState = [string](Get-ArrayaObjectValue -Object $record -Names @('State'))
    $recordedReason = [string](Get-ArrayaObjectValue -Object $record -Names @('Reason'))
    $recordedCollector = [string](Get-ArrayaObjectValue -Object $record -Names @('Collector'))
    $recordSuffix = if ([string]::IsNullOrWhiteSpace($recordedReason)) { '' } else { ": $recordedReason" }

    # Rows are authoritative that data exists, but a collector that recorded a partial or
    # failed run still has gaps behind those rows.
    if ($rowCount -gt 0) {
        if ($recordedState -in @('Partial', 'Failed', 'Unavailable')) {
            return (& $newResult 'Partial' $rowCount ("Rows were collected, but the collector recorded an incomplete run{0}" -f $recordSuffix) $recordedCollector $recordedState 'Collected')
        }
        return (& $newResult 'CompleteWithRecords' $rowCount 'Dataset contains collected rows.' $null $null $null)
    }

    if (-not [string]::IsNullOrWhiteSpace($recordedState)) {
        switch ($recordedState) {
            { $_ -in @('CompleteWithRecords', 'CompleteEmpty') } {
                if ($EmptyIsUnexpected) {
                    return (& $newResult 'Unknown' 0 ("'{0}' is empty although every tenant has at least one entry." -f $DatasetKey) $recordedCollector $recordedState $null)
                }
                return (& $newResult 'CompleteEmpty' 0 ("The collector answered and returned no rows{0}" -f $recordSuffix) $recordedCollector $recordedState $null)
            }
            'Partial' { return (& $newResult 'Partial' 0 ("The collector recorded a partial run{0}" -f $recordSuffix) $recordedCollector $recordedState $null) }
            'Failed' { return (& $newResult 'Failed' 0 ("The collector recorded a failure{0}" -f $recordSuffix) $recordedCollector $recordedState $null) }
            'Unavailable' { return (& $newResult 'Unavailable' 0 ("The source was unavailable{0}" -f $recordSuffix) $recordedCollector $recordedState $null) }
            'Unsupported' { return (& $newResult 'Unsupported' 0 ("The source is not supported in this collection mode{0}" -f $recordSuffix) $recordedCollector $recordedState $null) }
        }
    }

    $planRows = @()
    if ($TenantStatsHash.Contains('CollectorPlanResults') -and $null -ne $TenantStatsHash['CollectorPlanResults']) {
        $planValue = $TenantStatsHash['CollectorPlanResults']
        $planRows = @(if ($planValue -is [System.Collections.IDictionary]) { $planValue.Values } else { $planValue }) | Where-Object { $null -ne $_ }
    }

    $matchedStep = $null
    foreach ($planRow in $planRows) {
        $stepName = [string](Get-ArrayaObjectValue -Object $planRow -Names @('Name'))
        $producedKeys = @(Get-ArrayaObjectValue -Object $planRow -Names @('ProducedKeys')) | ForEach-Object { [string]$_ }
        $nameMatches = $false
        foreach ($candidate in @($CollectorStepName)) {
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            if ($stepName.StartsWith($candidate, [System.StringComparison]::OrdinalIgnoreCase)) {
                $nameMatches = $true
                break
            }
        }
        if ($nameMatches -or ($producedKeys -contains $DatasetKey)) {
            $matchedStep = $planRow
            break
        }
    }

    if ($null -ne $matchedStep) {
        $stepName = [string](Get-ArrayaObjectValue -Object $matchedStep -Names @('Name'))
        $stepStatus = [string](Get-ArrayaObjectValue -Object $matchedStep -Names @('Status'))
        $stepMessage = [string](Get-ArrayaObjectValue -Object $matchedStep -Names @('Message'))
        $messageSuffix = if ([string]::IsNullOrWhiteSpace($stepMessage)) { '' } else { ": $stepMessage" }
        switch ($stepStatus) {
            'Skipped' {
                if ($stepMessage -match '(?i)missing dependency') {
                    return (& $newResult 'Unavailable' 0 ("Collector step '{0}' could not run{1}" -f $stepName, $messageSuffix) $stepName $stepStatus $null)
                }
                if ($stepMessage -match '(?i)requires microsoft graph sdk collection mode|not supported') {
                    return (& $newResult 'Unsupported' 0 ("Collector step '{0}' is not supported in this collection mode{1}" -f $stepName, $messageSuffix) $stepName $stepStatus $null)
                }
                return (& $newResult 'SkippedByProfile' 0 ("Collector step '{0}' was skipped{1}" -f $stepName, $messageSuffix) $stepName $stepStatus $null)
            }
            'Failed' {
                return (& $newResult 'Failed' 0 ("Collector step '{0}' failed{1}" -f $stepName, $messageSuffix) $stepName $stepStatus $null)
            }
            'Completed' {
                if (-not $isPresent) {
                    return (& $newResult 'Unknown' 0 ("Collector step '{0}' completed but did not produce '{1}'." -f $stepName, $DatasetKey) $stepName $stepStatus $null)
                }
                if ($EmptyIsUnexpected) {
                    return (& $newResult 'Unknown' 0 ("'{0}' is empty although every tenant has at least one entry; treat it as not collected." -f $DatasetKey) $stepName $stepStatus $null)
                }
                return (& $newResult 'CompleteEmpty' 0 ("Collector step '{0}' completed and returned no rows." -f $stepName) $stepName $stepStatus $null)
            }
            default {
                return (& $newResult 'Unknown' 0 ("Collector step '{0}' reported status '{1}'." -f $stepName, $stepStatus) $stepName $stepStatus $null)
            }
        }
    }

    if ($planRows.Count -gt 0 -and -not $isPresent) {
        return (& $newResult 'NotRequested' 0 ("No collector step in this run produced '{0}'." -f $DatasetKey) $null $null $null)
    }

    if (-not $isPresent) {
        return (& $newResult 'Unknown' 0 ("'{0}' is not present in the assessment data." -f $DatasetKey) $null $null 'NotCollected')
    }

    return (& $newResult 'Unknown' 0 ("'{0}' is present and empty, but no collector result shows the collection was complete." -f $DatasetKey) $null $null $null)
}
