function Import-AssessmentPipelineDependencies {
    [CmdletBinding()]
    param()

    $repoRootVariable = Get-Variable -Name RepoRoot -Scope Script -ErrorAction SilentlyContinue
    if (-not $repoRootVariable -or [string]::IsNullOrWhiteSpace([string]$repoRootVariable.Value)) {
        $script:RepoRoot = (Resolve-Path -Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
    }

    $commonManifestPath = Join-Path $script:RepoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
    if (-not (Test-Path -Path $commonManifestPath)) {
        throw "Common module manifest not found: $commonManifestPath"
    }

    Import-Module -Name $commonManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop

    $graphManifestPath = Join-Path $script:RepoRoot 'src\modules\Arraya.M365.Graph\Arraya.M365.Graph.psd1'
    if (-not (Test-Path -Path $graphManifestPath)) {
        throw "Graph module manifest not found: $graphManifestPath"
    }
    Import-Module -Name $graphManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop

    $exchangeManifestPath = Join-Path $script:RepoRoot 'src\modules\Arraya.M365.Exchange\Arraya.M365.Exchange.psd1'
    if (-not (Test-Path -Path $exchangeManifestPath)) {
        throw "Exchange module manifest not found: $exchangeManifestPath"
    }
    Import-Module -Name $exchangeManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop

    Import-ArrayaOffice365CustomLocal -RepoRoot $script:RepoRoot -RequiredCommands @(
        'Office365Custom\Get-GraphData',
        'Write-ProgressHelper'
    ) | Out-Null

    $legacyScriptVariable = Get-Variable -Name LegacyAssessmentScriptPath -Scope Script -ErrorAction SilentlyContinue
    if (-not $legacyScriptVariable -or [string]::IsNullOrWhiteSpace([string]$legacyScriptVariable.Value)) {
        $script:LegacyAssessmentScriptPath = Join-Path $script:RepoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
    }
    if (-not (Test-Path -Path $script:LegacyAssessmentScriptPath)) {
        throw "Legacy assessment script not found: $script:LegacyAssessmentScriptPath"
    }
}
