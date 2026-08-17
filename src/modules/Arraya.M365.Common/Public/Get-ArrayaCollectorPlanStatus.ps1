function Get-ArrayaCollectorPlanStatus {
    <#
    .SYNOPSIS
        Reduces collector plan step results to a single run status.

    .DESCRIPTION
        Invoke-ArrayaCollectorPlan records a Status per step but does not judge the run as a
        whole. This function applies that judgement so callers can decide whether the
        collected data is fit to export and what exit code to return.

        Status values:
          Passed   - no collector failed.
          Degraded - one or more optional collectors failed. Deliverables are still usable,
                     but they have gaps that must be disclosed.
          Failed   - at least one collector marked Required failed. Data the deliverables
                     depend on is missing or untrustworthy.

        A step that was skipped (disabled by profile, or a missing dependency) is not a
        failure: skipping is a deliberate choice, not a collection error.

    .PARAMETER Result
        Step result objects as returned by Invoke-ArrayaCollectorPlan.

    .EXAMPLE
        $results = Invoke-ArrayaCollectorPlan -Sections $s -Steps $p
        $status = Get-ArrayaCollectorPlanStatus -Result $results
        if ($status.Status -eq 'Failed') { exit $status.SuggestedExitCode }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
        [AllowNull()]
        [object[]]$Result
    )

    begin {
        $collected = New-Object System.Collections.Generic.List[object]
    }

    process {
        foreach ($item in @($Result)) {
            if ($null -ne $item) {
                $collected.Add($item) | Out-Null
            }
        }
    }

    end {
        $all = @($collected.ToArray())

        $failed = @($all | Where-Object { [string]$_.Status -eq 'Failed' })
        $skipped = @($all | Where-Object { [string]$_.Status -eq 'Skipped' })
        $completed = @($all | Where-Object { [string]$_.Status -eq 'Completed' })

        $requiredFailed = @(
            $failed | Where-Object {
                $_.PSObject.Properties['Required'] -and [bool]$_.Required
            }
        )
        $optionalFailed = @(
            $failed | Where-Object {
                -not ($_.PSObject.Properties['Required'] -and [bool]$_.Required)
            }
        )

        $status = if ($requiredFailed.Count -gt 0) {
            'Failed'
        }
        elseif ($optionalFailed.Count -gt 0) {
            'Degraded'
        }
        else {
            'Passed'
        }

        # 0 = clean, 1 is reserved for the harness/unhandled errors, 2 = required data missing.
        $exitCode = if ($status -eq 'Failed') { 2 } else { 0 }

        $summary = switch ($status) {
            'Failed' {
                'Required collector(s) failed: {0}' -f (@($requiredFailed | ForEach-Object { [string]$_.Name }) -join ', ')
            }
            'Degraded' {
                'Optional collector(s) failed: {0}' -f (@($optionalFailed | ForEach-Object { [string]$_.Name }) -join ', ')
            }
            default {
                'All {0} collector step(s) completed or were intentionally skipped.' -f $all.Count
            }
        }

        return [pscustomobject][ordered]@{
            Status               = $status
            Summary              = $summary
            SuggestedExitCode    = $exitCode
            TotalSteps           = $all.Count
            CompletedCount       = $completed.Count
            SkippedCount         = $skipped.Count
            FailedCount          = $failed.Count
            RequiredFailedCount  = $requiredFailed.Count
            OptionalFailedCount  = $optionalFailed.Count
            RequiredFailedSteps  = @($requiredFailed | ForEach-Object { [string]$_.Name })
            OptionalFailedSteps  = @($optionalFailed | ForEach-Object { [string]$_.Name })
            FailedSteps          = @(
                $failed | ForEach-Object {
                    [pscustomobject][ordered]@{
                        Name     = [string]$_.Name
                        Section  = [string]$_.Section
                        Workload = [string]$_.Workload
                        Required = [bool]($_.PSObject.Properties['Required'] -and [bool]$_.Required)
                        Message  = [string]$_.Message
                    }
                }
            )
        }
    }
}
