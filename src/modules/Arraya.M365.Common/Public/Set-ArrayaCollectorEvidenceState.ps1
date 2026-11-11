function Set-ArrayaCollectorEvidenceState {
    <#
    .SYNOPSIS
        Records how completely a collector gathered one dataset.

    .DESCRIPTION
        Collectors that catch their own errors leave an empty container behind and report
        their step as completed, which makes a failure look like "none exist". A collector
        calls this function to record what actually happened, and
        Get-ArrayaDatasetEvidenceState reads the record before it trusts a row count.

        Records are stored in $TenantStatsHash['CollectorEvidenceStates'], keyed by dataset.
        The snapshot keeps them under Diagnostics, and the workbook does not export them.

    .PARAMETER State
        CompleteWithRecords - collected, rows returned.
        CompleteEmpty       - collected, the source answered with no rows.
        Partial             - some of the source answered.
        Failed              - the source did not answer.
        Unavailable         - the source could not be reached (permission, session, command).
        Unsupported         - the source is not supported in this collection mode.
    #>
    [CmdletBinding(SupportsShouldProcess = $false)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Writes to the in-memory assessment store only.')]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$TenantStatsHash,
        [Parameter(Mandatory = $true)]
        [string]$DatasetKey,
        [Parameter(Mandatory = $true)]
        [ValidateSet('CompleteWithRecords', 'CompleteEmpty', 'Partial', 'Failed', 'Unavailable', 'Unsupported')]
        [string]$State,
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Reason = '',
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Collector = ''
    )

    if (-not $TenantStatsHash.Contains('CollectorEvidenceStates') -or -not ($TenantStatsHash['CollectorEvidenceStates'] -is [System.Collections.IDictionary])) {
        $TenantStatsHash['CollectorEvidenceStates'] = @{}
    }

    $record = [pscustomobject][ordered]@{
        DatasetKey = $DatasetKey
        State      = $State
        Reason     = $Reason
        Collector  = $Collector
    }
    $TenantStatsHash['CollectorEvidenceStates'][$DatasetKey] = $record
    return $record
}
