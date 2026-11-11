function Get-ArrayaMailboxHoldEvidence {
    <#
    .SYNOPSIS
        Summarizes the mailbox-level hold indicators that the mailbox collector gathers.

    .DESCRIPTION
        Counts three preservation indicators and reports a fourth signal separately:

          LitigationHoldEnabled - litigation hold.
          DelayHoldApplied      - delay hold applied after a hold was removed.
          InPlaceHolds          - eDiscovery or retention-policy hold identifiers stamped on
                                  the mailbox. Entries that start with '-' are exclusions from
                                  an organization-wide policy and are not counted as holds.
          RetentionHoldEnabled  - retention processing is suspended. This is not a
                                  preservation hold, so it is counted on its own.

        A signal is "evaluated" only when every supplied mailbox row carries the property.
        Signals missing from some or all rows are listed as unavailable so a zero is never
        mistaken for "no hold". InPlaceHolds is collected only at the deepest mailbox detail
        level.

        Not evaluated here, because the collector does not gather them: compliance tag holds,
        organization-wide retention policies, and eDiscovery case detail. A result of zero
        therefore never establishes that a mailbox is free of holds.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Mailbox
    )

    $rows = @($Mailbox | Where-Object { $null -ne $_ })
    $signalNames = @('LitigationHoldEnabled', 'DelayHoldApplied', 'InPlaceHolds', 'RetentionHoldEnabled')
    $presentCounts = @{}
    foreach ($signalName in $signalNames) { $presentCounts[$signalName] = 0 }

    $testProperty = {
        param($Record, [string]$Name)
        if ($Record -is [System.Collections.IDictionary]) { return ([System.Collections.IDictionary]$Record).Contains($Name) }
        return [bool]($Record.PSObject -and $null -ne $Record.PSObject.Properties[$Name])
    }

    $litigationCount = 0
    $delayCount = 0
    $inPlaceCount = 0
    $retentionHoldCount = 0
    $anyHoldCount = 0

    foreach ($row in $rows) {
        foreach ($signalName in $signalNames) {
            if (& $testProperty $row $signalName) { $presentCounts[$signalName]++ }
        }

        $hasLitigation = (Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $row -Names @('LitigationHoldEnabled'))) -eq $true
        $hasDelay = (Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $row -Names @('DelayHoldApplied'))) -eq $true
        $hasRetentionHold = (Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $row -Names @('RetentionHoldEnabled'))) -eq $true

        $inPlaceValue = Get-ArrayaObjectValue -Object $row -Names @('InPlaceHolds')
        $inPlaceEntries = @()
        if ($null -ne $inPlaceValue) {
            $inPlaceEntries = if ($inPlaceValue -is [string]) {
                @($inPlaceValue -split '[,;]')
            }
            elseif ($inPlaceValue -is [System.Collections.IEnumerable]) {
                @($inPlaceValue | ForEach-Object { [string]$_ })
            }
            else {
                @([string]$inPlaceValue)
            }
        }
        $hasInPlace = @(
            $inPlaceEntries |
                ForEach-Object { ([string]$_).Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith('-') }
        ).Count -gt 0

        if ($hasLitigation) { $litigationCount++ }
        if ($hasDelay) { $delayCount++ }
        if ($hasInPlace) { $inPlaceCount++ }
        if ($hasRetentionHold) { $retentionHoldCount++ }
        if ($hasLitigation -or $hasDelay -or $hasInPlace) { $anyHoldCount++ }
    }

    $evaluatedSignals = @($signalNames | Where-Object { $rows.Count -gt 0 -and $presentCounts[$_] -eq $rows.Count })
    $unavailableSignals = @($signalNames | Where-Object { $rows.Count -eq 0 -or $presentCounts[$_] -lt $rows.Count })

    return [pscustomobject][ordered]@{
        MailboxCount        = $rows.Count
        AnyHoldCount        = $anyHoldCount
        LitigationHoldCount = $litigationCount
        DelayHoldCount      = $delayCount
        InPlaceHoldCount    = $inPlaceCount
        RetentionHoldCount  = $retentionHoldCount
        EvaluatedSignals    = $evaluatedSignals
        UnavailableSignals  = $unavailableSignals
        NotEvaluated        = 'Compliance tag holds, organization-wide retention policies, and eDiscovery case detail are not evaluated.'
    }
}
