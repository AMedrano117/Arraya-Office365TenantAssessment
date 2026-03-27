[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('M365', 'M365Collect', 'M365Export', 'AD', 'Improve', 'Compare')]
    [string]$Action,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
    [string]$AuthMode,
    [Parameter(Mandatory = $false)]
    [switch]$StoreTenantStatsGlobal,
    [Parameter(Mandatory = $false)]
    [string]$TenantStatsVariableName = 'ArrayaTenantStats',
    [Parameter(Mandatory = $false)]
    [switch]$SkipAuth,
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
    [string[]]$OutputProfile = @('SolutionsEngineer'),
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,
    [Parameter(Mandatory = $false)]
    [switch]$UseGraphFallback,
    [Parameter(Mandatory = $false)]
    [switch]$RunImprove,
    [Parameter(Mandatory = $false)]
    [string]$ImproveOutputFolder,
    [Parameter(Mandatory = $false)]
    [switch]$SkipPdfReport,
    [Parameter(Mandatory = $false)]
    [switch]$SkipJsonReport
)

$resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\shared\Resolve-ArrayaRepoRoot.ps1'))
if (-not (Test-Path -Path $resolveRepoRootHelperPath)) {
    throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
}
. $resolveRepoRootHelperPath

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
    Import-Module -Name $resolvedCommonManifestPath -Force -DisableNameChecking -ErrorAction Stop
}

if (-not (Test-Path -Path $runnerManifestPath)) {
    throw "Runner module manifest not found: $runnerManifestPath"
}
$resolvedRunnerManifestPath = (Resolve-Path -Path $runnerManifestPath).Path
$loadedRunnerModule = Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue | Select-Object -First 1
$requiredRunnerCommands = @(
    'Invoke-M365TenantAssessment',
    'Invoke-M365TenantDataCollection',
    'Invoke-M365TenantAssessmentExport'
)
$missingRunnerCommands = @(
    $requiredRunnerCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
)
if (
    -not $loadedRunnerModule -or
    $loadedRunnerModule.Path -ne $resolvedRunnerManifestPath -or
    $missingRunnerCommands.Count -gt 0
) {
    Import-Module -Name $resolvedRunnerManifestPath -Force -DisableNameChecking -ErrorAction Stop
}

function Test-LauncherOutputProfilesGenerateJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$OutputProfile = @('SolutionsEngineer')
    )

    foreach ($rawProfile in @($OutputProfile)) {
        if ([string]::IsNullOrWhiteSpace([string]$rawProfile)) {
            continue
        }

        foreach ($token in ([string]$rawProfile -split ',')) {
            $profile = $token.Trim()
            if ([string]::IsNullOrWhiteSpace($profile)) {
                continue
            }

            $policy = Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile $profile
            if ([bool]$policy.GenerateJson) {
                return $true
            }
        }
    }

    return $false
}

function Add-LauncherOutputProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$OutputProfile = @('SolutionsEngineer'),
        [Parameter(Mandatory = $true)]
        [string]$ProfileToAdd
    )

    $resolvedProfiles = New-Object System.Collections.Generic.List[string]
    $profileSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($rawProfile in @($OutputProfile)) {
        if ([string]::IsNullOrWhiteSpace([string]$rawProfile)) {
            continue
        }

        foreach ($token in ([string]$rawProfile -split ',')) {
            $profile = $token.Trim()
            if ([string]::IsNullOrWhiteSpace($profile)) {
                continue
            }

            $null = Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile $profile
            if ($profileSet.Add($profile)) {
                $resolvedProfiles.Add($profile)
            }
        }
    }

    if ($profileSet.Add($ProfileToAdd)) {
        $resolvedProfiles.Add($ProfileToAdd)
    }

    if ($resolvedProfiles.Count -eq 0) {
        $resolvedProfiles.Add('SolutionsEngineer')
    }

    return $resolvedProfiles.ToArray()
}

function Resolve-LauncherLatestManifestPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportPath
    )

    $candidateManifestPaths = New-Object System.Collections.Generic.List[string]
    $fullExportPath = [System.IO.Path]::GetFullPath($ExportPath)

    if (Test-Path -Path $fullExportPath -PathType Leaf) {
        if ($fullExportPath -match '\.manifest\.json$') {
            $candidateManifestPaths.Add($fullExportPath)
        }
        elseif ($fullExportPath -match '\.xlsx$') {
            $candidateManifestPaths.Add(($fullExportPath -replace '\.xlsx$', '.manifest.json'))
        }
        else {
            $candidateManifestPaths.Add("$fullExportPath.manifest.json")
        }
    }
    else {
        if ([System.IO.Path]::HasExtension($fullExportPath)) {
            if ($fullExportPath -match '\.xlsx$') {
                $candidateManifestPaths.Add(($fullExportPath -replace '\.xlsx$', '.manifest.json'))
            }
            elseif ($fullExportPath -match '\.manifest\.json$') {
                $candidateManifestPaths.Add($fullExportPath)
            }
            else {
                $candidateManifestPaths.Add("$fullExportPath.manifest.json")
            }
        }
    }

    foreach ($candidatePath in $candidateManifestPaths) {
        if (Test-Path -Path $candidatePath -PathType Leaf) {
            return (Resolve-Path -Path $candidatePath).Path
        }
    }

    $searchRoot = if (Test-Path -Path $fullExportPath -PathType Container) {
        $fullExportPath
    }
    else {
        $parent = Split-Path -Path $fullExportPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($parent) -and (Test-Path -Path $parent -PathType Container)) {
            $parent
        }
        else {
            $null
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($searchRoot)) {
        $latestManifest = Get-ChildItem -Path $searchRoot -Recurse -Filter '*.manifest.json' -File -ErrorAction SilentlyContinue |
            Sort-Object -Property LastWriteTimeUtc -Descending |
            Select-Object -First 1
        if ($latestManifest) {
            return $latestManifest.FullName
        }
    }

    return $null
}

function Invoke-LauncherImproveFromLatestRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [string]$OutputFolder,
        [Parameter(Mandatory = $false)]
        [switch]$UseGraphFallback
    )

    $manifestPath = Resolve-LauncherLatestManifestPath -ExportPath $ExportPath
    if ([string]::IsNullOrWhiteSpace($manifestPath)) {
        throw "Could not find a manifest for the completed assessment run under: $ExportPath"
    }

    $resolvedOutputFolder = $OutputFolder
    if ([string]::IsNullOrWhiteSpace($resolvedOutputFolder)) {
        $manifestDirectory = Split-Path -Path $manifestPath -Parent
        $resolvedOutputFolder = Join-Path -Path $manifestDirectory -ChildPath 'Improve'
    }

    Write-Host ("Launching Improve from manifest: {0}" -f $manifestPath) -ForegroundColor Cyan
    Invoke-M365ImprovementPlan -AssessmentJsonPath $manifestPath -OutputFolder $resolvedOutputFolder -UseGraphFallback:$UseGraphFallback
}

if ([string]::IsNullOrWhiteSpace($Action)) {
    Write-Host ''
    Write-Host 'Tenant Assessment Launcher' -ForegroundColor Cyan
    Write-Host '1. Microsoft 365 Full Tenant Assessment'
    Write-Host '2. Microsoft 365 Data Collection Only (JSON snapshot)'
    Write-Host '3. Microsoft 365 Export from JSON Snapshot'
    Write-Host '4. Active Directory Assessment'
    Write-Host '5. Build Improvement Plan from Tenant JSON'
    Write-Host '6. Compare Two Tenant JSON Snapshots'
    Write-Host ''

    $choice = Read-Host 'Select an option (1-6)'
    switch ($choice) {
        '1' { $Action = 'M365' }
        '2' { $Action = 'M365Collect' }
        '3' { $Action = 'M365Export' }
        '4' { $Action = 'AD' }
        '5' { $Action = 'Improve' }
        '6' { $Action = 'Compare' }
        default { throw "Invalid selection: $choice" }
    }
}

switch ($Action) {
    'M365' {
        $defaultOutputRoot = Get-ArrayaAssessmentOutputRoot -FallbackPath $repoRoot
        $exportPathInput = if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
            $ExportPath
        } else {
            Read-Host "Export path (.xlsx or folder) - default $defaultOutputRoot"
        }
        $selectedOutputProfiles = @($OutputProfile)
        if (-not $PSBoundParameters.ContainsKey('OutputProfile')) {
            $outputProfileInput = Read-Host 'Output profile(s) (Presales, SolutionsEngineer, ExecutiveLevel, TenantToTenantMigration, Geek, Machine) - comma-separated, default SolutionsEngineer'
            if (-not [string]::IsNullOrWhiteSpace($outputProfileInput)) {
                $selectedOutputProfiles = @(
                    $outputProfileInput -split ',' |
                        ForEach-Object { $_.Trim() } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                )
            }
        }

        if (-not $selectedOutputProfiles -or $selectedOutputProfiles.Count -eq 0) {
            $selectedOutputProfiles = @('SolutionsEngineer')
        }

        if ($RunImprove -and -not (Test-LauncherOutputProfilesGenerateJson -OutputProfile $selectedOutputProfiles)) {
            $selectedOutputProfiles = Add-LauncherOutputProfile -OutputProfile $selectedOutputProfiles -ProfileToAdd 'Machine'
            Write-Warning "RunImprove requested. Added output profile 'Machine' so the run preserves a reusable JSON snapshot."
        }

        $invokeParams = @{}
        $invokeParams.OutputProfile = $selectedOutputProfiles
        $invokeParams.ExportPath = if (-not [string]::IsNullOrWhiteSpace($exportPathInput)) { $exportPathInput } else { $defaultOutputRoot }
        if ($SkipPdfReport) { $invokeParams.SkipPdfReport = $true }
        if ($SkipJsonReport -and -not $RunImprove) {
            $invokeParams.SkipJsonReport = $true
        }
        elseif ($SkipJsonReport -and $RunImprove) {
            Write-Warning 'RunImprove requires a JSON snapshot. Ignoring -SkipJsonReport for this run.'
        }
        if ($StoreTenantStatsGlobal) { $invokeParams.StoreTenantStatsGlobal = $true }
        if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
        if ($SkipAuth) { $invokeParams.SkipAuth = $true }
        if (-not [string]::IsNullOrWhiteSpace($AuthMode)) { $invokeParams.AuthMode = $AuthMode }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $invokeParams.TenantId = $TenantId }
        if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $invokeParams.ClientId = $ClientId }
        if (-not [string]::IsNullOrWhiteSpace($ClientSecret)) { $invokeParams.ClientSecret = $ClientSecret }

        Invoke-M365TenantAssessment @invokeParams
        if ($RunImprove) {
            Invoke-LauncherImproveFromLatestRun -ExportPath $invokeParams.ExportPath -OutputFolder $ImproveOutputFolder -UseGraphFallback:$UseGraphFallback
        }
    }
    'M365Collect' {
        $defaultOutputRoot = Get-ArrayaAssessmentOutputRoot -FallbackPath $repoRoot
        $exportPathInput = if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
            $ExportPath
        } else {
            Read-Host "Snapshot output path (.xlsx or folder) - default $defaultOutputRoot"
        }
        $selectedOutputProfiles = @($OutputProfile)
        if (-not $PSBoundParameters.ContainsKey('OutputProfile')) {
            $outputProfileInput = Read-Host 'Collection profile(s) (Presales, SolutionsEngineer, ExecutiveLevel, TenantToTenantMigration, Geek, Machine) - comma-separated, default SolutionsEngineer'
            if (-not [string]::IsNullOrWhiteSpace($outputProfileInput)) {
                $selectedOutputProfiles = @(
                    $outputProfileInput -split ',' |
                        ForEach-Object { $_.Trim() } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                )
            }
        }

        if (-not $selectedOutputProfiles -or $selectedOutputProfiles.Count -eq 0) {
            $selectedOutputProfiles = @('SolutionsEngineer')
        }

        if ($RunImprove -and -not (Test-LauncherOutputProfilesGenerateJson -OutputProfile $selectedOutputProfiles)) {
            $selectedOutputProfiles = Add-LauncherOutputProfile -OutputProfile $selectedOutputProfiles -ProfileToAdd 'Machine'
            Write-Warning "RunImprove requested. Added output profile 'Machine' so the run preserves a reusable JSON snapshot."
        }

        $invokeParams = @{}
        $invokeParams.OutputProfile = $selectedOutputProfiles
        $invokeParams.ExportPath = if (-not [string]::IsNullOrWhiteSpace($exportPathInput)) { $exportPathInput } else { $defaultOutputRoot }
        if ($StoreTenantStatsGlobal) { $invokeParams.StoreTenantStatsGlobal = $true }
        if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
        if ($SkipAuth) { $invokeParams.SkipAuth = $true }
        if (-not [string]::IsNullOrWhiteSpace($AuthMode)) { $invokeParams.AuthMode = $AuthMode }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $invokeParams.TenantId = $TenantId }
        if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $invokeParams.ClientId = $ClientId }
        if (-not [string]::IsNullOrWhiteSpace($ClientSecret)) { $invokeParams.ClientSecret = $ClientSecret }

        Invoke-M365TenantDataCollection @invokeParams
        if ($RunImprove) {
            Invoke-LauncherImproveFromLatestRun -ExportPath $invokeParams.ExportPath -OutputFolder $ImproveOutputFolder -UseGraphFallback:$UseGraphFallback
        }
    }
    'M365Export' {
        $defaultOutputRoot = Get-ArrayaAssessmentOutputRoot -FallbackPath $repoRoot
        $jsonPath = Read-Host 'Path to collected tenant JSON snapshot'
        if ([string]::IsNullOrWhiteSpace($jsonPath)) {
            throw 'A JSON snapshot path is required for export.'
        }
        $exportPathInput = if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
            $ExportPath
        } else {
            Read-Host "Export path (.xlsx or folder) - default $defaultOutputRoot"
        }

        $selectedOutputProfiles = @($OutputProfile)
        if (-not $PSBoundParameters.ContainsKey('OutputProfile')) {
            $outputProfileInput = Read-Host 'Export profile(s) (Presales, SolutionsEngineer, ExecutiveLevel, TenantToTenantMigration, Geek, Machine) - comma-separated, default SolutionsEngineer'
            if (-not [string]::IsNullOrWhiteSpace($outputProfileInput)) {
                $selectedOutputProfiles = @(
                    $outputProfileInput -split ',' |
                        ForEach-Object { $_.Trim() } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                )
            }
        }

        if (-not $selectedOutputProfiles -or $selectedOutputProfiles.Count -eq 0) {
            $selectedOutputProfiles = @('SolutionsEngineer')
        }

        $invokeParams = @{ AssessmentJsonPath = $jsonPath }
        $invokeParams.OutputProfile = $selectedOutputProfiles
        $invokeParams.ExportPath = if (-not [string]::IsNullOrWhiteSpace($exportPathInput)) { $exportPathInput } else { $defaultOutputRoot }
        if ($SkipPdfReport) { $invokeParams.SkipPdfReport = $true }
        if ($SkipJsonReport) { $invokeParams.SkipJsonReport = $true }

        Invoke-M365TenantAssessmentExport @invokeParams
    }
    'AD' {
        Invoke-ADTenantAssessment
    }
    'Improve' {
        $jsonPath = Read-Host 'Path to tenant assessment JSON or manifest JSON'
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
