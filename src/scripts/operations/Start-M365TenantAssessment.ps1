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
    [ValidateSet('Lean', 'Standard', 'Full')]
    [string]$OutputProfile,
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
$runnerManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
if (-not (Test-Path -Path $runnerManifestPath)) {
    throw "Runner module manifest not found: $runnerManifestPath"
}

if (-not (Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue)) {
    Import-Module -Name $runnerManifestPath -ErrorAction Stop
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
        $reportingMode = Read-Host 'Reporting mode (Minimum, Combined, All, Geek) - leave blank for prompt'
        $outputProfileInput = Read-Host 'Output profile (Lean, Standard, Full) - default Standard'
        $exportPath = Read-Host 'Export path (.xlsx or folder) - leave blank for prompt'
        $skipHtmlInput = Read-Host 'Skip full HTML report? (Y/N, default by profile)'
        $skipPdfInput = Read-Host 'Skip PDF report? (Y/N, default by profile)'
        $skipJsonInput = Read-Host 'Skip JSON snapshot? (Y/N, default by profile)'

        $invokeParams = @{}
        if (-not [string]::IsNullOrWhiteSpace($reportingMode)) { $invokeParams.ReportingMode = $reportingMode }
        if (-not [string]::IsNullOrWhiteSpace($outputProfileInput)) { $invokeParams.OutputProfile = $outputProfileInput }
        elseif (-not [string]::IsNullOrWhiteSpace($OutputProfile)) { $invokeParams.OutputProfile = $OutputProfile }
        if (-not [string]::IsNullOrWhiteSpace($exportPath)) { $invokeParams.ExportPath = $exportPath }
        if ($skipHtmlInput -match '^(y|yes)$') { $invokeParams.SkipHtmlReport = $true }
        if ($skipPdfInput -match '^(y|yes)$' -or $SkipPdfReport) { $invokeParams.SkipPdfReport = $true }
        if ($skipJsonInput -match '^(y|yes)$' -or $SkipJsonReport) { $invokeParams.SkipJsonReport = $true }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $invokeParams.TenantId = $TenantId }
        if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $invokeParams.ClientId = $ClientId }

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
