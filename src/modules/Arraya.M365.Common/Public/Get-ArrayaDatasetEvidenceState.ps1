function Get-ArrayaDatasetEvidenceState {
    <#
    .SYNOPSIS
        Tells whether a dataset in the assessment store was actually collected.

    .DESCRIPTION
        A count of zero only means "none exist" when the data behind it was collected. This
        function separates the cases that a bare row count hides:

          Collected    - rows exist, or the dataset is present and nothing contradicts it.
          NotCollected - the dataset is absent, or its collector step was skipped.
          Failed       - the collector step failed.
          Unknown      - the collector step finished without producing the dataset, or the
                         dataset is empty although it cannot legitimately be empty.

        Inputs are the dataset itself and, when present, the CollectorPlanResults rows that
        Invoke-ArrayaCollectorPlan records. Stores without collector results (older snapshots,
        hand-built data) fall back to dataset presence.

        Limitation: a collector that catches its own errors still reports Completed, so an
        internal failure that leaves an empty container is indistinguishable from an empty
        result here. Use -EmptyIsUnexpected for datasets that always have at least one entry.

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
        [string[]]$CollectorStepName = @(),
        [Parameter(Mandatory = $false)]
        [switch]$EmptyIsUnexpected
    )

    $newResult = {
        param([string]$State, [int]$RowCount, [string]$Reason, [string]$StepName, [string]$StepStatus)
        [pscustomobject][ordered]@{
            DatasetKey   = $DatasetKey
            State        = $State
            IsCollected  = ($State -eq 'Collected')
            RowCount     = $RowCount
            Reason       = $Reason
            CollectorStep = $StepName
            CollectorStatus = $StepStatus
        }
    }

    if ($null -eq $TenantStatsHash) {
        return (& $newResult 'NotCollected' 0 'No assessment data was supplied.' $null $null)
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

    # Rows are authoritative: whatever a step reported, the data exists.
    if ($rowCount -gt 0) {
        return (& $newResult 'Collected' $rowCount 'Dataset contains collected rows.' $null $null)
    }

    $planRows = @()
    if ($TenantStatsHash.Contains('CollectorPlanResults') -and $null -ne $TenantStatsHash['CollectorPlanResults']) {
        $planValue = $TenantStatsHash['CollectorPlanResults']
        $planRows = if ($planValue -is [System.Collections.IDictionary]) { @($planValue.Values) } else { @($planValue) }
    }

    $matchedStep = $null
    foreach ($planRow in $planRows) {
        if ($null -eq $planRow) { continue }
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
                return (& $newResult 'NotCollected' 0 ("Collector step '{0}' was skipped{1}" -f $stepName, $messageSuffix) $stepName $stepStatus)
            }
            'Failed' {
                return (& $newResult 'Failed' 0 ("Collector step '{0}' failed{1}" -f $stepName, $messageSuffix) $stepName $stepStatus)
            }
            'Completed' {
                if (-not $isPresent) {
                    return (& $newResult 'Unknown' 0 ("Collector step '{0}' completed but did not produce '{1}'." -f $stepName, $DatasetKey) $stepName $stepStatus)
                }
            }
            default {
                return (& $newResult 'Unknown' 0 ("Collector step '{0}' reported status '{1}'." -f $stepName, $stepStatus) $stepName $stepStatus)
            }
        }

        if ($EmptyIsUnexpected) {
            return (& $newResult 'Unknown' 0 ("'{0}' is empty although every tenant has at least one entry; treat it as not collected." -f $DatasetKey) $stepName $stepStatus)
        }
        return (& $newResult 'Collected' 0 ("Collector step '{0}' completed and returned no rows." -f $stepName) $stepName $stepStatus)
    }

    if (-not $isPresent) {
        return (& $newResult 'NotCollected' 0 ("'{0}' is not present in the assessment data." -f $DatasetKey) $null $null)
    }

    if ($EmptyIsUnexpected) {
        return (& $newResult 'Unknown' 0 ("'{0}' is empty although every tenant has at least one entry; treat it as not collected." -f $DatasetKey) $null $null)
    }

    return (& $newResult 'Collected' 0 ("'{0}' is present and empty; no collector status was recorded." -f $DatasetKey) $null $null)
}
