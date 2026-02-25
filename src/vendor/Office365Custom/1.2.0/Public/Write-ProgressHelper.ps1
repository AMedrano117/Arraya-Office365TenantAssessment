function Write-ProgressHelper {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$Total,
        [int]$Index,
        [string]$Activity = 'Processing',
        [string]$Operation,
        [int]$Id = 1,
        [int]$ParentId,
        [switch]$Completed
    )

    if ($ProgressPreference -eq 'SilentlyContinue') { $ProgressPreference = 'Continue' }

    # Always ensure required progress dictionaries exist before any use
    if (-not $script:ProgressStartTimes) { $script:ProgressStartTimes = @{} }
    if (-not $script:ProgressIndices)    { $script:ProgressIndices    = @{} }
    if (-not $script:ProgressTotals)     { $script:ProgressTotals     = @{} }

    if (-not $script:ProgressStartTimes.ContainsKey($Id)) { $script:ProgressStartTimes[$Id] = Get-Date }
    if (-not $script:ProgressIndices.ContainsKey($Id))    { $script:ProgressIndices[$Id]    = 0 }
    if (-not $script:ProgressTotals.ContainsKey($Id))     { $script:ProgressTotals[$Id]     = $Total }

    # If Total changed for this Id, reset timer/index so batches don't bleed together
    if ($script:ProgressTotals[$Id] -ne $Total) {
        $script:ProgressStartTimes[$Id] = Get-Date
        $script:ProgressIndices[$Id]    = 0
        $script:ProgressTotals[$Id]     = $Total
    }

    # Auto-increment when Index isn't provided
    if (-not $PSBoundParameters.ContainsKey('Index')) {
        $script:ProgressIndices[$Id]++
        $Index = $script:ProgressIndices[$Id]
    } else {
        $script:ProgressIndices[$Id] = $Index
    }

    $startTime = $script:ProgressStartTimes[$Id]
    $elapsed   = (Get-Date) - $startTime

    # Clamp index within [0, Total] so percent never exceeds 100
    if ($Total -gt 0) {
        if     ($Index -lt 0)      { $Index = 0 }
        elseif ($Index -gt $Total) { $Index = $Total }
    }

    $percent = if ($Total -gt 0) {
        $raw = (($Index / $Total) * 100)
        [math]::Round([math]::Min(100,[math]::Max(0,$raw)), 2)   # <--- CLAMP
    } else { $null }

    $etaSec  = if ($Index -gt 0 -and $Total -ge $Index) {
        $rate = $elapsed.TotalSeconds / $Index
        [int][math]::Max(0, [math]::Round($rate * ($Total - $Index)))
    } else { $null }

    $status = if ($Total -gt 0) { "[${Index} / ${Total}] $($elapsed.ToString('hh\:mm\:ss')) elapsed" }
              else              { "$($elapsed.ToString('hh\:mm\:ss')) elapsed" }

    $splat = @{ Activity=$Activity; Status=$status; Id=$Id }
    if ($null -ne $percent)   { $splat.PercentComplete  = $percent }
    if ($null -ne $etaSec)    { $splat.SecondsRemaining = $etaSec }
    if ($Operation)           { $splat.CurrentOperation = $Operation }
    if ($PSBoundParameters.ContainsKey('ParentId')) { $splat.ParentId = $ParentId }
    if ($Completed.IsPresent) { $splat.Completed = $true }

    Write-Progress @splat

    if ($Completed.IsPresent) {
        $script:ProgressStartTimes.Remove($Id) | Out-Null
        $script:ProgressIndices.Remove($Id)    | Out-Null
        $script:ProgressTotals.Remove($Id)     | Out-Null
    }
}