[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('M365', 'AD', 'Graph', 'Improve', 'Compare')]
    [string]$Action,
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Lean', 'Standard', 'Full')]
    [string]$OutputProfile = 'Lean',
    [Parameter(Mandatory = $false)]
    [switch]$UseGraphFallback,
    [Parameter(Mandatory = $false)]
    [switch]$SkipPdfReport,
    [Parameter(Mandatory = $false)]
    [switch]$SkipJsonReport
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
$commonManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
$runnerManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'

if (-not (Test-Path -Path $commonManifestPath)) {
    throw "Common module manifest not found: $commonManifestPath"
}
$resolvedCommonManifestPath = (Resolve-Path -Path $commonManifestPath).Path
$loadedCommonModule = Get-Module -Name 'Arraya.M365.Common' -ErrorAction SilentlyContinue | Select-Object -First 1
$requiredCommonCommands = @('Get-ArrayaAssessmentOutputRoot')
$missingCommonCommands = @(
    $requiredCommonCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
)
if (
    -not $loadedCommonModule -or
    $loadedCommonModule.Path -ne $resolvedCommonManifestPath -or
    $missingCommonCommands.Count -gt 0
) {
    Import-Module -Name $resolvedCommonManifestPath -Force -ErrorAction Stop
}

if (-not (Test-Path -Path $runnerManifestPath)) {
    throw "Runner module manifest not found: $runnerManifestPath"
}
$resolvedRunnerManifestPath = (Resolve-Path -Path $runnerManifestPath).Path
$loadedRunnerModule = Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue | Select-Object -First 1
$requiredRunnerCommands = @('Invoke-M365TenantAssessment')
$missingRunnerCommands = @(
    $requiredRunnerCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
)
if (
    -not $loadedRunnerModule -or
    $loadedRunnerModule.Path -ne $resolvedRunnerManifestPath -or
    $missingRunnerCommands.Count -gt 0
) {
    Import-Module -Name $resolvedRunnerManifestPath -Force -ErrorAction Stop
}

if ([string]::IsNullOrWhiteSpace($Action)) {
    Write-Host ''
    Write-Host 'Tenant Assessment Launcher' -ForegroundColor Cyan
    Write-Host '1. Microsoft 365 Full Tenant Assessment'
    Write-Host '2. Active Directory Assessment'
    Write-Host '3. Microsoft Graph Activity Report'
    Write-Host '4. Build Improvement Plan from Tenant JSON'
    Write-Host '5. Compare Two Tenant JSON Snapshots'
    Write-Host ''

    $choice = Read-Host 'Select an option (1-5)'
    switch ($choice) {
        '1' { $Action = 'M365' }
        '2' { $Action = 'AD' }
        '3' { $Action = 'Graph' }
        '4' { $Action = 'Improve' }
        '5' { $Action = 'Compare' }
        default { throw "Invalid selection: $choice" }
    }
}

switch ($Action) {
    'M365' {
        $defaultOutputRoot = Get-ArrayaAssessmentOutputRoot -FallbackPath $repoRoot
        $reportingModeInput = Read-Host 'Reporting scope (Minimum, Combined, All, Geek) - default Minimum'
        $outputProfileInput = Read-Host 'Output set (Lean, Standard, Full) - default Lean'
        $exportPathInput = Read-Host "Export path (.xlsx or folder) - default $defaultOutputRoot"

        $invokeParams = @{}
        $invokeParams.ReportingMode = if ([string]::IsNullOrWhiteSpace($reportingModeInput)) { 'Minimum' } else { $reportingModeInput }
        $invokeParams.OutputProfile = if (-not [string]::IsNullOrWhiteSpace($outputProfileInput)) { $outputProfileInput } else { $OutputProfile }
        $invokeParams.ExportPath = if (-not [string]::IsNullOrWhiteSpace($exportPathInput)) { $exportPathInput } else { $defaultOutputRoot }
        if ($SkipPdfReport) { $invokeParams.SkipPdfReport = $true }
        if ($SkipJsonReport) { $invokeParams.SkipJsonReport = $true }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $invokeParams.TenantId = $TenantId }
        if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $invokeParams.ClientId = $ClientId }
        if (-not [string]::IsNullOrWhiteSpace($ClientSecret)) { $invokeParams.ClientSecret = $ClientSecret }

        Invoke-M365TenantAssessment @invokeParams
    }
    'AD' {
        Invoke-ADTenantAssessment
    }
    'Graph' {
        $serviceName = Read-Host 'ServiceName (example: Office365ActiveUser, SharePointSites, TeamsUser)'
        $period = Read-Host 'Period (D7, D30, D90, D180) - default D90'
        if ([string]::IsNullOrWhiteSpace($period)) { $period = 'D90' }
        $useBetaInput = Read-Host 'Use beta endpoint? (Y/N, default N)'
        $useBeta = $useBetaInput -match '^(y|yes)$'

        $graphData = Invoke-GraphActivityAssessment -ServiceName $serviceName -PeriodDuration $period -UseBeta:$useBeta
        $graphData | Format-Table -AutoSize
    }
    'Improve' {
        $jsonPath = Read-Host 'Path to tenant assessment JSON'
        $outputFolder = Read-Host 'Output folder (leave blank to use JSON folder)'

        $invokeParams = @{ AssessmentJsonPath = $jsonPath }
        if (-not [string]::IsNullOrWhiteSpace($outputFolder)) { $invokeParams.OutputFolder = $outputFolder }
        if ($UseGraphFallback) { $invokeParams.UseGraphFallback = $true }

        Invoke-M365ImprovementPlan @invokeParams
    }
    'Compare' {
        $baselinePath = Read-Host 'Path to BASELINE JSON'
        $currentPath = Read-Host 'Path to CURRENT JSON'
        $outputFolder = Read-Host 'Output folder (leave blank to use current JSON folder)'

        $invokeParams = @{
            BaselineJsonPath = $baselinePath
            CurrentJsonPath  = $currentPath
        }
        if (-not [string]::IsNullOrWhiteSpace($outputFolder)) { $invokeParams.OutputFolder = $outputFolder }

        Invoke-M365AssessmentComparison @invokeParams
    }
}
