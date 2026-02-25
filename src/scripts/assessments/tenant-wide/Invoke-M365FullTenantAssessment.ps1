[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Minimum', 'Combined', 'All', 'Geek')]
    [string]$ReportingMode,
    [Parameter(Mandatory = $false)]
    [switch]$SkipHtmlReport
)

function Resolve-ArrayaRepoRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$StartPath
    )

    $candidate = (Resolve-Path -Path $StartPath).Path
    while ($true) {
        $runnerManifestPath = Join-Path $candidate 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
        if (Test-Path -Path $runnerManifestPath) {
            return $candidate
        }

        $parent = Split-Path -Path $candidate -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $candidate) {
            break
        }
        $candidate = $parent
    }

    throw "Could not resolve repository root from: $StartPath"
}

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$runnerManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
if (-not (Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue)) {
    Import-Module -Name $runnerManifestPath -ErrorAction Stop
}

$invokeParams = @{}
if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
if ($PSBoundParameters.ContainsKey('ReportingMode')) { $invokeParams.ReportingMode = $ReportingMode }
if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }

Invoke-M365TenantAssessment @invokeParams
