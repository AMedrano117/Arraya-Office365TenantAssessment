function Invoke-ArrayaCollectorPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Sections,
        [Parameter(Mandatory = $true)]
        [array]$Steps,
        [Parameter(Mandatory = $false)]
        [scriptblock]$OnSection,
        [Parameter(Mandatory = $false)]
        [scriptblock]$OnStep,
        [Parameter(Mandatory = $false)]
        $Context
    )

    $results = New-Object System.Collections.Generic.List[object]
    $completed = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($section in @($Sections)) {
        if ($OnSection) {
            & $OnSection $section
        }

        $sectionName = [string]$section.Name
        foreach ($step in @($Steps | Where-Object { [string]$_.Section -eq $sectionName })) {
            $start = Get-Date
            $status = 'Completed'
            $message = $null
            $missingDependencies = @(
                @($step.DependsOn) | Where-Object {
                    -not [string]::IsNullOrWhiteSpace([string]$_) -and
                    -not $completed.Contains([string]$_)
                }
            )
            if ($missingDependencies.Count -gt 0) {
                $status = 'Skipped'
                $message = 'Missing dependency: {0}' -f ($missingDependencies -join ', ')
            }

            try {
                if ($status -eq 'Skipped') {
                    if ($OnStep) {
                        & $OnStep $step $false $message | Out-Null
                    }
                }
                elseif ($OnStep) {
                    & $OnStep $step ([bool]$step.Enabled) ([string]$step.SkipReason) | Out-Null
                    if (-not [bool]$step.Enabled) {
                        $status = 'Skipped'
                        $message = [string]$step.SkipReason
                    }
                }
                elseif ([bool]$step.Enabled) {
                    & $step.ScriptBlock | Out-Null
                }
                else {
                    $status = 'Skipped'
                    $message = [string]$step.SkipReason
                }
            }
            catch {
                $status = 'Failed'
                $message = $_.Exception.Message
            }
            finally {
                $elapsed = (Get-Date) - $start
                if ($status -eq 'Completed') {
                    $completed.Add([string]$step.Name) | Out-Null
                }

                $isRequired = $false
                if ($step.PSObject.Properties['Required']) {
                    $isRequired = [bool]$step.Required
                }

                $results.Add([pscustomobject][ordered]@{
                    Name            = [string]$step.Name
                    Section         = $sectionName
                    Workload        = [string]$step.Workload
                    Status          = $status
                    Required        = $isRequired
                    Message         = $message
                    DurationSeconds = [math]::Round($elapsed.TotalSeconds, 3)
                    ProducedKeys    = @($step.Produces)
                }) | Out-Null
            }
        }
    }

    if ($Context -and $Context.PSObject.Properties['Runtime'] -and ($Context.Runtime -is [System.Collections.IDictionary])) {
        $Context.Runtime['CollectorPlanResults'] = @($results.ToArray())
    }

    return @($results.ToArray())
}
