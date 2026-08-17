function Invoke-ArrayaExchangeHybridProbe {
    <#
    .SYNOPSIS
        Runs a single hybrid-detection query and records whether it actually answered.

    .DESCRIPTION
        Hybrid detection reasons from the absence of evidence ("no connectors were found,
        therefore this tenant is not hybrid"). That inference is only valid when the query
        succeeded. A suppressed error looks identical to a genuine empty result, which lets a
        permission denial masquerade as a confident "not hybrid" answer.

        This wrapper separates the two: it reports Succeeded plus a FailureKind so callers can
        distinguish "asked and found nothing" from "could not ask".

    .PARAMETER Name
        Probe name used in evidence and log messages.

    .PARAMETER ScriptBlock
        The query to run. Should use -ErrorAction Stop so real failures surface.

    .PARAMETER RequiresCommand
        Optional cmdlet name. When absent from the session, the probe is reported as
        Unavailable rather than Failed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        [Parameter(Mandatory = $false)]
        [string]$RequiresCommand
    )

    if (-not [string]::IsNullOrWhiteSpace($RequiresCommand)) {
        if (-not (Get-Command -Name $RequiresCommand -ErrorAction Ignore)) {
            return [pscustomobject][ordered]@{
                Name         = $Name
                Succeeded    = $false
                FailureKind  = 'Unavailable'
                ErrorMessage = "Cmdlet '$RequiresCommand' is not available in this session."
                Value        = @()
            }
        }
    }

    try {
        $value = & $ScriptBlock
        return [pscustomobject][ordered]@{
            Name         = $Name
            Succeeded    = $true
            FailureKind  = $null
            ErrorMessage = $null
            Value        = @($value)
        }
    }
    catch {
        $message = [string]$_.Exception.Message
        $failureKind = if (Test-ArrayaExchangeProbePermissionError -ErrorRecord $_) {
            'PermissionDenied'
        }
        elseif ($_.CategoryInfo -and [string]$_.CategoryInfo.Category -eq 'ObjectNotFound' -and $message -match "is not recognized") {
            'Unavailable'
        }
        else {
            'Error'
        }

        return [pscustomobject][ordered]@{
            Name         = $Name
            Succeeded    = $false
            FailureKind  = $failureKind
            ErrorMessage = $message
            Value        = @()
        }
    }
}

function Test-ArrayaExchangeProbePermissionError {
    <#
    .SYNOPSIS
        Heuristically identifies access-denied failures from Exchange Online cmdlets.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $ErrorRecord
    )

    $message = [string]$ErrorRecord.Exception.Message
    if ([string]::IsNullOrWhiteSpace($message)) {
        return $false
    }

    $permissionPatterns = @(
        'access\s+is\s+denied',
        'access\s+denied',
        'unauthorized',
        'not\s+authorized',
        'insufficient\s+(privileges|permission)',
        'forbidden',
        "doesn't have permission",
        'does not have permission',
        'you don''t have sufficient permissions',
        'operation is not allowed',
        'RBAC',
        'management role'
    )

    foreach ($pattern in $permissionPatterns) {
        if ($message -match $pattern) {
            return $true
        }
    }

    if ($ErrorRecord.CategoryInfo -and [string]$ErrorRecord.CategoryInfo.Category -eq 'PermissionDenied') {
        return $true
    }

    return $false
}
