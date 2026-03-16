function Get-ExoMailboxStatisticsSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $MailboxObjects,
        [Parameter(Mandatory = $false)]
        $Context,
        [switch]$Archive,
        [Parameter(Mandatory = $false)]
        [string]$ProgressActivity,
        [Parameter(Mandatory = $false)]
        [int]$ProgressId = 0
    )

    $results = New-Object System.Collections.Generic.List[object]
    $failures = New-Object System.Collections.Generic.List[object]
    $normalizedMailboxObjects = New-Object System.Collections.Generic.List[object]
    if ($Context) {
        $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context
    }

    if ($null -ne $MailboxObjects) {
        foreach ($mailboxItem in $MailboxObjects) {
            $normalizedMailboxObjects.Add($mailboxItem)
        }
    }

    $totalMailboxCount = $normalizedMailboxObjects.Count
    $processedMailboxCount = 0
    $progressStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($mailbox in $normalizedMailboxObjects) {
        $processedMailboxCount++
        if (-not [string]::IsNullOrWhiteSpace($ProgressActivity) -and $totalMailboxCount -gt 0) {
            $label = if ($mailbox -and $mailbox.PSObject.Properties['DisplayName'] -and -not [string]::IsNullOrWhiteSpace([string]$mailbox.DisplayName)) {
                [string]$mailbox.DisplayName
            }
            elseif ($mailbox -and $mailbox.PSObject.Properties['PrimarySmtpAddress'] -and -not [string]::IsNullOrWhiteSpace([string]$mailbox.PrimarySmtpAddress)) {
                [string]$mailbox.PrimarySmtpAddress
            }
            else {
                'Processing mailbox'
            }

            $percentComplete = [math]::Round(($processedMailboxCount / $totalMailboxCount) * 100, 2)
            $elapsedText = if ($progressStopwatch) { $progressStopwatch.Elapsed.ToString('hh\:mm\:ss') } else { '00:00:00' }
            $progressStatus = "[{0} / {1}] {2} | {3} elapsed | results={4}, failures={5}" -f $processedMailboxCount, $totalMailboxCount, $label, $elapsedText, $results.Count, $failures.Count
            Write-Progress -Id $ProgressId -Activity $ProgressActivity -Status $progressStatus -PercentComplete $percentComplete
        }

        $identity = Resolve-ExoStatisticsIdentity -MailboxObject $mailbox
        $displayName = $null
        if ($null -ne $mailbox -and $mailbox.PSObject.Properties['DisplayName']) {
            $displayName = [string]$mailbox.DisplayName
        }

        if ([string]::IsNullOrWhiteSpace($identity)) {
            $failures.Add([PSCustomObject]@{
                DisplayName = $displayName
                Identity    = $null
                Reason      = 'No supported mailbox identity was available.'
            })
            continue
        }

        try {
            $statsParams = @{
                Identity    = $identity
                ErrorAction = 'Stop'
            }

            if ($Archive) {
                $statsParams.Archive = $true
                $statsParams.Properties = 'MailboxGuid'
                $statsParams.IncludeSoftDeletedRecipients = $true
            }
            else {
                $statsParams.IncludeSoftDeletedRecipient = $true
            }

            $statResults = Invoke-QuietCommand -ScriptBlock { Get-EXOMailboxStatistics @statsParams }
            if ($statResults.Count -eq 0) {
                $failures.Add([PSCustomObject]@{
                    DisplayName = $displayName
                    Identity    = $identity
                    Reason      = 'No mailbox statistics were returned.'
                })
                continue
            }

            foreach ($stat in $statResults) {
                $totalItemBytes = Convert-DataSizeToBytes -Value $stat.TotalItemSize
                $deletedItemBytes = Convert-DataSizeToBytes -Value $stat.TotalDeletedItemSize
                if (-not $stat.PSObject.Properties['TotalItemSizeBytes']) {
                    $stat | Add-Member -MemberType NoteProperty -Name 'TotalItemSizeBytes' -Value $totalItemBytes -Force
                }
                if (-not $stat.PSObject.Properties['TotalDeletedItemSizeBytes']) {
                    $stat | Add-Member -MemberType NoteProperty -Name 'TotalDeletedItemSizeBytes' -Value $deletedItemBytes -Force
                }
                $results.Add($stat)
            }
        }
        catch {
            $failures.Add([PSCustomObject]@{
                DisplayName = $displayName
                Identity    = $identity
                Reason      = $_.Exception.Message
            })
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ProgressActivity)) {
        Write-Progress -Id $ProgressId -Activity $ProgressActivity -Completed
    }
    if ($progressStopwatch -and $progressStopwatch.IsRunning) {
        $progressStopwatch.Stop()
    }

    return [PSCustomObject]@{
        Results  = $results.ToArray()
        Failures = $failures.ToArray()
    }
}
