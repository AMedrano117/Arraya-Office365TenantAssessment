[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Path = './tests',

    [Parameter(Mandatory = $false)]
    [string]$TestResultPath = './TestResults.xml',

    [Parameter(Mandatory = $false)]
    [string]$CoverageResultPath = './CoverageResults.xml',

    # Guards against a silently empty run (bad path, discovery failure, filter typo).
    [Parameter(Mandatory = $false)]
    [int]$MinimumTestCount = 1,

    # Coverage floor to ratchet upward over time. 0 disables the gate.
    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 100)]
    [double]$MinimumCoveragePercent = 0,

    [Parameter(Mandatory = $false)]
    [switch]$SkipCodeCoverage
)

# Deliberately no Set-StrictMode / $ErrorActionPreference here: both leak from this script
# scope into the Pester run and change what the tests exercise. Setup calls below pass
# -ErrorAction Stop explicitly instead.

# The module manifests under src/modules all declare PowerShellVersion = '7.0'. Running the
# suite under Windows PowerShell 5.1 exercises a runtime the product does not support, so
# fail loudly instead of reporting a green run against the wrong engine.
if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw "This test suite requires PowerShell 7 or later (found $($PSVersionTable.PSVersion)). Run it with 'pwsh'."
}

Import-Module Pester -MinimumVersion 5.0 -ErrorAction Stop

$config = New-PesterConfiguration
$config.Run.Path = $Path
$config.Run.PassThru = $true
$config.Output.Verbosity = 'Detailed'
$config.TestResult.Enabled = $true
$config.TestResult.OutputPath = $TestResultPath
$config.TestResult.OutputFormat = 'NUnitXml'

if (-not $SkipCodeCoverage) {
    $coverageFiles = @(Get-ChildItem './src/modules' -Recurse -Filter '*.ps1' -ErrorAction Stop).FullName
    if ($coverageFiles.Count -eq 0) {
        throw 'Code coverage was requested but no .ps1 files were found under ./src/modules.'
    }

    $config.CodeCoverage.Enabled = $true
    $config.CodeCoverage.Path = $coverageFiles
    $config.CodeCoverage.OutputPath = $CoverageResultPath
}

$result = Invoke-Pester -Configuration $config

$failures = New-Object System.Collections.Generic.List[string]

# Pester reports discovery and container-level errors here rather than as failed tests, so a
# run can report FailedCount = 0 while an entire test file never executed.
if ($result.FailedContainersCount -gt 0) {
    $failedContainerNames = @(
        $result.Containers |
            Where-Object { $_.Result -ne 'Passed' } |
            ForEach-Object { [string]$_.Item }
    )
    $failures.Add(("{0} test container(s) failed to run: {1}" -f $result.FailedContainersCount, ($failedContainerNames -join '; ')))
}

if ($result.FailedCount -gt 0) {
    $failures.Add("$($result.FailedCount) test(s) failed.")
}

# Pester returns a result object with empty properties when the path matches no test files,
# so coerce before comparing and reporting.
$totalCount = [int]($result.TotalCount | Select-Object -First 1)
if ($totalCount -lt $MinimumTestCount) {
    $failures.Add("Only $totalCount test(s) ran; expected at least $MinimumTestCount. Check the test path and discovery output.")
}

# Catches states the individual counters above do not model (for example Inconclusive).
$overallResult = [string]$result.Result
if ($overallResult -ne 'Passed') {
    $label = if ([string]::IsNullOrWhiteSpace($overallResult)) { '<none>' } else { $overallResult }
    $failures.Add("Overall Pester result was '$label' rather than 'Passed'.")
}

$coveragePercent = $null
if (-not $SkipCodeCoverage -and $result.PSObject.Properties['CodeCoverage'] -and $result.CodeCoverage) {
    $coveragePercent = [math]::Round([double]$result.CodeCoverage.CoveragePercent, 2)
    if ($MinimumCoveragePercent -gt 0 -and $coveragePercent -lt $MinimumCoveragePercent) {
        $failures.Add("Code coverage $coveragePercent% is below the required minimum of $MinimumCoveragePercent%.")
    }
}

Write-Output ''
Write-Output ("PowerShell runtime: {0}" -f $PSVersionTable.PSVersion)
Write-Output ("Pester result: {0}" -f $result.Result)
Write-Output ("Tests: {0} total, {1} passed, {2} failed, {3} skipped" -f $result.TotalCount, $result.PassedCount, $result.FailedCount, $result.SkippedCount)
Write-Output ("Containers: {0} total, {1} failed" -f @($result.Containers).Count, $result.FailedContainersCount)
if ($null -ne $coveragePercent) {
    Write-Output ("Code coverage: {0}%{1}" -f $coveragePercent, $(if ($MinimumCoveragePercent -gt 0) { " (minimum $MinimumCoveragePercent%)" } else { ' (no minimum enforced)' }))
}

if ($failures.Count -gt 0) {
    Write-Output ''
    foreach ($failure in $failures) {
        Write-Output ("FAIL: {0}" -f $failure)
    }

    exit 1
}

Write-Output 'Pester passed.'
