[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Error', 'Warning', 'Information')]
    [string]$FailOnSeverity = 'Error',
    [Parameter(Mandatory = $false)]
    [string[]]$TargetPath = @('./src', './tools'),
    [Parameter(Mandatory = $false)]
    [string[]]$ExcludePathPrefix = @('./src/vendor')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-SeverityRank {
    param([string]$Severity)

    switch ($Severity) {
        'ParseError' { return 4 }
        'Error' { return 3 }
        'Warning' { return 2 }
        default { return 1 }
    }
}

$targets = @($TargetPath) | Where-Object { Test-Path $_ }
if ($targets.Count -eq 0) {
    throw 'PSScriptAnalyzer found no valid targets to scan.'
}

$resolvedExcludePrefixes = @(
    foreach ($prefix in @($ExcludePathPrefix)) {
        if ([string]::IsNullOrWhiteSpace([string]$prefix)) {
            continue
        }

    if (Test-Path $prefix) {
        (Resolve-Path -Path $prefix -ErrorAction Stop).Path
        continue
    }

        [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location).Path -ChildPath $prefix))
    }
)

$issues = New-Object System.Collections.Generic.List[object]
foreach ($target in $targets) {
    foreach ($issue in @(Invoke-ScriptAnalyzer -Path $target -Recurse -ErrorAction Stop)) {
        if ($issue -and $issue.PSObject.Properties['ScriptPath'] -and -not [string]::IsNullOrWhiteSpace([string]$issue.ScriptPath)) {
            $resolvedIssuePath = [System.IO.Path]::GetFullPath([string]$issue.ScriptPath)
            $isExcluded = $false
            foreach ($excludePrefix in @($resolvedExcludePrefixes)) {
                if ($resolvedIssuePath.StartsWith($excludePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $isExcluded = $true
                    break
                }
            }

            if ($isExcluded) {
                continue
            }
        }

        $issues.Add($issue)
    }
}

$blockingSeverityRank = Get-SeverityRank -Severity $FailOnSeverity
$blockingIssues = @(
    $issues | Where-Object {
        (Get-SeverityRank -Severity ([string]$_.Severity)) -ge $blockingSeverityRank
    }
)

$summary = if ($issues.Count -gt 0) {
    ($issues | Group-Object Severity | Sort-Object Name | ForEach-Object { '{0}={1}' -f $_.Name, $_.Count }) -join ', '
}
else {
    'none'
}

Write-Output ("PSScriptAnalyzer summary: {0}" -f $summary)
if ($resolvedExcludePrefixes.Count -gt 0) {
    Write-Output ("Excluded path prefixes: {0}" -f ($resolvedExcludePrefixes -join '; '))
}

if ($blockingIssues.Count -gt 0) {
    $blockingIssues | Sort-Object -Property ScriptPath, Line, RuleName | Format-Table -AutoSize
    throw "PSScriptAnalyzer found $($blockingIssues.Count) issue(s) at severity '$FailOnSeverity' or higher."
}

if ($issues.Count -gt 0) {
    Write-Warning ("PSScriptAnalyzer found {0} non-blocking issue(s) below the configured fail threshold '{1}'." -f $issues.Count, $FailOnSeverity)
}

Write-Output 'PSScriptAnalyzer passed.'
