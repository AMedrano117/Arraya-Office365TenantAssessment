[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Full', 'Preflight', 'Collect', 'Report', 'AD', 'Improve', 'Compare')]
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
    [switch]$SkipPermissionPreflight,
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    [Parameter(Mandatory = $false)]
    [pscredential]$ClientSecretCredential,
    [Parameter(Mandatory = $false)]
    [securestring]$ClientSecretSecure,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
    [string[]]$OutputProfile = @('SolutionsEngineer'),
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,
    [Parameter(Mandatory = $false)]
    [Alias('LiveRefresh')]
    [switch]$UseGraphFallback,
    [Parameter(Mandatory = $false)]
    [switch]$RunImprove,
    [Parameter(Mandatory = $false)]
    [switch]$SkipImprove,
    [Parameter(Mandatory = $false)]
    [string]$ImproveOutputFolder,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyArtifacts,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyAssessmentArtifacts,
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

function Test-LauncherModuleMatchesManifestPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSModuleInfo]$Module,
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    if (-not $Module) {
        return $false
    }

    try {
        $resolvedManifestPath = (Resolve-Path -Path $ManifestPath -ErrorAction Stop).Path
    }
    catch {
        return $false
    }

    $candidatePaths = @(
        $Module.Path
        (Join-Path -Path $Module.ModuleBase -ChildPath ([System.IO.Path]::GetFileName($resolvedManifestPath)))
        $Module.ModuleBase
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

    foreach ($candidatePath in $candidatePaths) {
        try {
            $resolvedCandidatePath = [System.IO.Path]::GetFullPath($candidatePath)
        }
        catch {
            $resolvedCandidatePath = $candidatePath
        }

        if ([string]::Equals($resolvedCandidatePath, $resolvedManifestPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    $resolvedManifestDirectory = Split-Path -Path $resolvedManifestPath -Parent
    if (
        -not [string]::IsNullOrWhiteSpace([string]$Module.ModuleBase) -and
        [string]::Equals(
            ([System.IO.Path]::GetFullPath($Module.ModuleBase)),
            ([System.IO.Path]::GetFullPath($resolvedManifestDirectory)),
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        return $true
    }

    return $false
}

if (-not (Test-Path -Path $commonManifestPath)) {
    throw "Common module manifest not found: $commonManifestPath"
}
$resolvedCommonManifestPath = (Resolve-Path -Path $commonManifestPath).Path
$loadedCommonModule = Get-Module -Name 'Arraya.M365.Common' -ErrorAction SilentlyContinue | Select-Object -First 1
$requiredCommonCommands = @(
    'Get-ArrayaAssessmentOutputRoot',
    'Import-ArrayaTenantSnapshotContext'
)
$missingCommonCommands = @(
    $requiredCommonCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
)
if (
    -not $loadedCommonModule -or
    -not (Test-LauncherModuleMatchesManifestPath -Module $loadedCommonModule -ManifestPath $resolvedCommonManifestPath) -or
    $missingCommonCommands.Count -gt 0
) {
    Import-Module -Name $resolvedCommonManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
}

if (-not (Test-Path -Path $runnerManifestPath)) {
    throw "Runner module manifest not found: $runnerManifestPath"
}
$resolvedRunnerManifestPath = (Resolve-Path -Path $runnerManifestPath).Path
$loadedRunnerModule = Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue | Select-Object -First 1
$requiredRunnerCommands = @(
    'Invoke-M365TenantAssessment',
    'Invoke-M365TenantConnectionPreflight',
    'Invoke-M365TenantDataCollection',
    'Invoke-M365TenantAssessmentExport'
)
$missingRunnerCommands = @(
    $requiredRunnerCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
)
if (
    -not $loadedRunnerModule -or
    -not (Test-LauncherModuleMatchesManifestPath -Module $loadedRunnerModule -ManifestPath $resolvedRunnerManifestPath) -or
    $missingRunnerCommands.Count -gt 0
) {
    Import-Module -Name $resolvedRunnerManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
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

# The manifest-resolution helpers this script calls below are imported from
# Arraya.M365.AssessmentRunner. Local Launcher-prefixed copies of them used to live here
# and had already begun to drift from the module originals, so the module now exports them
# and the launcher and the runner resolve a prior run through one implementation.
# Test-LauncherModuleMatchesManifestPath above stays local on purpose: it runs before the
# module is imported and cannot come from it.

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

    $manifestPath = Resolve-AssessmentLatestManifestPath -ExportPath $ExportPath
    if ([string]::IsNullOrWhiteSpace($manifestPath)) {
        throw "Could not find a run manifest for the completed assessment under: $ExportPath. The workbook alone is not enough for Improve; rerun a JSON-enabled output profile and check the export log for manifest write warnings."
    }

    try {
        Import-ArrayaTenantSnapshotContext -Path $manifestPath -Purpose ImprovementPlan | Out-Null
    }
    catch {
        throw ("The completed assessment run did not produce a usable JSON snapshot, so Improve cannot continue from the manifest. This usually means snapshot export failed earlier in the run even if the workbook was created. Review the earlier export errors in the run log and rerun after the snapshot issue is corrected. Manifest: {0} Underlying error: {1}" -f $manifestPath, $_.Exception.Message)
    }

    $resolvedOutputFolder = $OutputFolder
    if ([string]::IsNullOrWhiteSpace($resolvedOutputFolder)) {
        $resolvedOutputFolder = Resolve-AssessmentRunRootFromManifestPath -ManifestPath $manifestPath
    }

    Write-Host ("Launching Improve from manifest: {0}" -f $manifestPath) -ForegroundColor Cyan
    Invoke-M365ImprovementPlan -AssessmentJsonPath $manifestPath -OutputFolder $resolvedOutputFolder -UseGraphFallback:$UseGraphFallback
}

function Test-LauncherShouldRunImprove {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action,
        [Parameter(Mandatory = $false)]
        [switch]$RunImprove,
        [Parameter(Mandatory = $false)]
        [switch]$SkipImprove
    )

    switch ($Action) {
        'Full'    { return (-not $SkipImprove) }
        'Collect' { return ($RunImprove -and -not $SkipImprove) }
        default { return $false }
    }
}

function Write-LauncherBanner {
    [CmdletBinding()]
    param()
    $tl = [char]0x2554; $tr = [char]0x2557
    $bl = [char]0x255A; $br = [char]0x255D
    $h  = [char]0x2550; $v  = [char]0x2551
    $top    = "  $tl$([string]::new($h, 54))$tr"
    $middle = "  $v       Arraya M365 Tenant Assessment Launcher         $v"
    $bottom = "  $bl$([string]::new($h, 54))$br"
    Write-Host ''
    Write-Host $top    -ForegroundColor Cyan
    Write-Host $middle -ForegroundColor Cyan
    Write-Host $bottom -ForegroundColor Cyan
    Write-Host ("  {0}" -f (Get-Date -Format 'yyyy-MM-dd  HH:mm')) -ForegroundColor DarkGray
    Write-Host ''
}

function Write-LauncherMenuGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Label
    )
    Write-Host ("  {0} {1}" -f ([string]::new([char]0x2500, 3)), $Label) -ForegroundColor DarkGray
}

function Read-LauncherMenuChoice {
    [CmdletBinding()]
    param()

    Write-LauncherBanner

    function Write-MenuItem {
        param([string]$Number, [string]$Action, [string]$Description)
        Write-Host -NoNewline '    '
        Write-Host -NoNewline $Number    -ForegroundColor Cyan
        Write-Host -NoNewline ('  {0,-10}' -f $Action) -ForegroundColor DarkGray
        Write-Host $Description
    }

    Write-LauncherMenuGroup -Label 'Live'
    Write-MenuItem '1' 'Full'      'Microsoft 365 full assessment + improvement plan'
    Write-MenuItem '2' 'Preflight' 'Test M365 connections and permissions only'
    Write-MenuItem '3' 'Collect'   'Collect M365 data to a JSON snapshot'
    Write-Host ''
    Write-LauncherMenuGroup -Label 'From Snapshot'
    Write-MenuItem '4' 'Report'  'Generate Excel report + improvement plan from snapshot'
    Write-MenuItem '5' 'Improve' 'Build improvement plan docs only from snapshot'
    Write-MenuItem '6' 'Compare' 'Compare two snapshots'
    Write-Host ''
    Write-LauncherMenuGroup -Label 'Other'
    Write-MenuItem '7' 'AD' 'Active Directory assessment'
    Write-Host ''

    $actionMap = @{
        '1' = 'Full'; '2' = 'Preflight'; '3' = 'Collect'
        '4' = 'Report'; '5' = 'Improve'; '6' = 'Compare'; '7' = 'AD'
    }
    while ($true) {
        $choice = Read-Host 'Select an option (1-7)'
        if ($actionMap.ContainsKey($choice)) {
            return $actionMap[$choice]
        }
        Write-Warning "Invalid choice '$choice'. Please enter a number from 1 to 7."
        Write-Host ''
    }
}

function Read-LauncherExportPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DefaultPath,
        [Parameter(Mandatory = $false)]
        [string]$Label = 'Export path (folder or .xlsx)'
    )
    $value = Read-Host ("{0} [default: {1}]" -f $Label, $DefaultPath)
    if ([string]::IsNullOrWhiteSpace($value)) {
        Write-Host ("  Using default: {0}" -f $DefaultPath) -ForegroundColor DarkCyan
        return $DefaultPath
    }
    Write-Host ("  Export path: {0}" -f $value) -ForegroundColor DarkCyan
    return $value
}

function Read-LauncherOutputProfile {
    [CmdletBinding()]
    param()
    Write-Host '  Profiles: Presales | SolutionsEngineer | ExecutiveLevel | TenantToTenantMigration | Geek | Machine' -ForegroundColor DarkGray
    $value = Read-Host 'Profile(s) [default: SolutionsEngineer]  comma-separated'
    if ([string]::IsNullOrWhiteSpace($value)) {
        Write-Host '  Using profile: SolutionsEngineer' -ForegroundColor DarkCyan
        return @('SolutionsEngineer')
    }
    $profiles = @(
        $value -split ',' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
    Write-Host ("  Profile(s): {0}" -f ($profiles -join ', ')) -ForegroundColor DarkCyan
    return $profiles
}

function Read-LauncherFilePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,
        [Parameter(Mandatory = $false)]
        [switch]$CheckExists
    )
    $value = Read-Host $Prompt
    if ($CheckExists -and -not [string]::IsNullOrWhiteSpace($value) -and -not (Test-Path -Path $value)) {
        Write-Warning "File not found: $value. Continuing - validation will occur during the run."
    }
    return $value
}

function Write-LauncherActionHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Description
    )
    $line = [string]::new([char]0x2500, 58)
    Write-Host ''
    Write-Host ("  {0}" -f $line) -ForegroundColor DarkGray
    Write-Host ("  Running: {0}" -f $Description) -ForegroundColor Cyan
    Write-Host ("  {0}" -f $line) -ForegroundColor DarkGray
    Write-Host ''
}

function Write-LauncherRunSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Fields
    )
    Write-Host '  Run parameters:' -ForegroundColor DarkGray
    foreach ($key in $Fields.Keys) {
        $value = [string]$Fields[$key]
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            Write-Host -NoNewline ("    {0,-14}: " -f $key) -ForegroundColor DarkGray
            Write-Host $value
        }
    }
    Write-Host ''
}

if ([string]::IsNullOrWhiteSpace($Action)) {
    $Action = Read-LauncherMenuChoice
}

$clientSecretRequested = (
    [string]::Equals([string]$AuthMode, 'ClientSecret', [System.StringComparison]::OrdinalIgnoreCase) -or
    -not [string]::IsNullOrWhiteSpace($ClientSecret) -or
    $null -ne $ClientSecretCredential -or
    $null -ne $ClientSecretSecure
)
if ($clientSecretRequested) {
    Write-Warning 'Client secret auth is a compatibility path in this workflow. Microsoft Graph app auth remains available, but Exchange Online falls back to delegated sign-in and Purview compliance app auth is not supported. Prefer certificate auth for unattended production runs.'
}

switch ($Action) {
    'Full' {
        $defaultOutputRoot = Get-ArrayaAssessmentOutputRoot -FallbackPath $repoRoot
        $exportPathInput = if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
            $ExportPath
        } else {
            Read-LauncherExportPath -DefaultPath $defaultOutputRoot
        }
        $selectedOutputProfiles = @($OutputProfile)
        if (-not $PSBoundParameters.ContainsKey('OutputProfile')) {
            $selectedOutputProfiles = Read-LauncherOutputProfile
        }

        if (-not $selectedOutputProfiles -or $selectedOutputProfiles.Count -eq 0) {
            $selectedOutputProfiles = @('SolutionsEngineer')
        }

        $invokeParams = @{}
        $invokeParams.OutputProfile = $selectedOutputProfiles
        $invokeParams.ExportPath = if (-not [string]::IsNullOrWhiteSpace($exportPathInput)) { $exportPathInput } else { $defaultOutputRoot }
        if ($SkipPdfReport) { $invokeParams.SkipPdfReport = $true }
        if ($SkipJsonReport -and $IncludeLegacyAssessmentArtifacts) {
            $invokeParams.SkipJsonReport = $true
        }
        elseif ($SkipJsonReport) {
            Write-Warning 'The consolidated M365 workflow preserves a JSON snapshot for Improve and export replay. Ignoring -SkipJsonReport unless -IncludeLegacyAssessmentArtifacts is used.'
        }
        if ($SkipImprove) { $invokeParams.SkipImprove = $true }
        if ($IncludeLegacyArtifacts) { $invokeParams.IncludeLegacyArtifacts = $true }
        if ($IncludeLegacyAssessmentArtifacts) { $invokeParams.IncludeLegacyAssessmentArtifacts = $true }
        if (-not [string]::IsNullOrWhiteSpace($ImproveOutputFolder)) { $invokeParams.ImproveOutputFolder = $ImproveOutputFolder }
        if ($UseGraphFallback) { $invokeParams.UseGraphFallback = $true }
        if ($StoreTenantStatsGlobal) { $invokeParams.StoreTenantStatsGlobal = $true }
        if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
        if ($SkipAuth) { $invokeParams.SkipAuth = $true }
        if ($SkipPermissionPreflight) { $invokeParams.SkipPermissionPreflight = $true }
        if (-not [string]::IsNullOrWhiteSpace($AuthMode)) { $invokeParams.AuthMode = $AuthMode }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $invokeParams.TenantId = $TenantId }
        if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $invokeParams.ClientId = $ClientId }
        if (-not [string]::IsNullOrWhiteSpace($ClientSecret)) { $invokeParams.ClientSecret = $ClientSecret }
        if ($null -ne $ClientSecretCredential) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
        if ($null -ne $ClientSecretSecure) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }

        Write-LauncherActionHeader -Description 'Microsoft 365 Full Assessment + Improvement Plan'
        Write-LauncherRunSummary -Fields ([ordered]@{
            'Export path' = $invokeParams.ExportPath
            'Profile(s)'  = ($invokeParams.OutputProfile -join ', ')
            'Auth mode'   = if (-not [string]::IsNullOrWhiteSpace([string]$AuthMode)) { $AuthMode } else { 'Interactive' }
            'Improve'     = if ($SkipImprove) { 'No (skipped)' } else { 'Yes (auto)' }
        })
        Invoke-M365TenantAssessment @invokeParams
    }
    'Preflight' {
        $selectedOutputProfiles = @($OutputProfile)
        if (-not $selectedOutputProfiles -or $selectedOutputProfiles.Count -eq 0) {
            $selectedOutputProfiles = @('SolutionsEngineer')
        }

        $invokeParams = @{}
        $invokeParams.OutputProfile = $selectedOutputProfiles
        if (-not [string]::IsNullOrWhiteSpace($ExportPath)) { $invokeParams.ExportPath = $ExportPath }
        if ($SkipAuth) { $invokeParams.SkipAuth = $true }
        if ($SkipPermissionPreflight) { $invokeParams.SkipPermissionPreflight = $true }
        if (-not [string]::IsNullOrWhiteSpace($AuthMode)) { $invokeParams.AuthMode = $AuthMode }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $invokeParams.TenantId = $TenantId }
        if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $invokeParams.ClientId = $ClientId }
        if (-not [string]::IsNullOrWhiteSpace($ClientSecret)) { $invokeParams.ClientSecret = $ClientSecret }
        if ($null -ne $ClientSecretCredential) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
        if ($null -ne $ClientSecretSecure) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }

        Write-LauncherActionHeader -Description 'Microsoft 365 Connection / Preflight Check'
        Invoke-M365TenantConnectionPreflight @invokeParams
    }
    'Collect' {
        $defaultOutputRoot = Get-ArrayaAssessmentOutputRoot -FallbackPath $repoRoot
        $exportPathInput = if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
            $ExportPath
        } else {
            Read-LauncherExportPath -DefaultPath $defaultOutputRoot -Label 'Snapshot output path (folder or .xlsx)'
        }
        $selectedOutputProfiles = @($OutputProfile)
        if (-not $PSBoundParameters.ContainsKey('OutputProfile')) {
            $selectedOutputProfiles = Read-LauncherOutputProfile
        }

        if (-not $selectedOutputProfiles -or $selectedOutputProfiles.Count -eq 0) {
            $selectedOutputProfiles = @('SolutionsEngineer')
        }

        $invokeParams = @{}
        $invokeParams.OutputProfile = $selectedOutputProfiles
        $invokeParams.ExportPath = if (-not [string]::IsNullOrWhiteSpace($exportPathInput)) { $exportPathInput } else { $defaultOutputRoot }
        if ($RunImprove) { $invokeParams.RunImprove = $true }
        if ($IncludeLegacyArtifacts) { $invokeParams.IncludeLegacyArtifacts = $true }
        if (-not [string]::IsNullOrWhiteSpace($ImproveOutputFolder)) { $invokeParams.ImproveOutputFolder = $ImproveOutputFolder }
        if ($UseGraphFallback) { $invokeParams.UseGraphFallback = $true }
        if ($StoreTenantStatsGlobal) { $invokeParams.StoreTenantStatsGlobal = $true }
        if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
        if ($SkipAuth) { $invokeParams.SkipAuth = $true }
        if ($SkipPermissionPreflight) { $invokeParams.SkipPermissionPreflight = $true }
        if (-not [string]::IsNullOrWhiteSpace($AuthMode)) { $invokeParams.AuthMode = $AuthMode }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $invokeParams.TenantId = $TenantId }
        if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $invokeParams.ClientId = $ClientId }
        if (-not [string]::IsNullOrWhiteSpace($ClientSecret)) { $invokeParams.ClientSecret = $ClientSecret }
        if ($null -ne $ClientSecretCredential) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
        if ($null -ne $ClientSecretSecure) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }

        Write-LauncherActionHeader -Description 'Microsoft 365 Data Collection'
        Write-LauncherRunSummary -Fields ([ordered]@{
            'Snapshot path' = $invokeParams.ExportPath
            'Profile(s)'    = ($invokeParams.OutputProfile -join ', ')
            'Auth mode'     = if (-not [string]::IsNullOrWhiteSpace([string]$AuthMode)) { $AuthMode } else { 'Interactive' }
            'Run Improve'   = if ($RunImprove) { 'Yes' } else { 'No' }
        })
        Invoke-M365TenantDataCollection @invokeParams
    }
    'Report' {
        $defaultOutputRoot = Get-ArrayaAssessmentOutputRoot -FallbackPath $repoRoot
        $jsonPath = Read-LauncherFilePath -Prompt 'Path to collected tenant JSON snapshot' -CheckExists
        if ([string]::IsNullOrWhiteSpace($jsonPath)) {
            throw 'A JSON snapshot path is required for export.'
        }
        $exportPathInput = if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
            $ExportPath
        } else {
            Read-LauncherExportPath -DefaultPath $defaultOutputRoot
        }

        $selectedOutputProfiles = @($OutputProfile)
        if (-not $PSBoundParameters.ContainsKey('OutputProfile')) {
            $selectedOutputProfiles = Read-LauncherOutputProfile
        }

        if (-not $selectedOutputProfiles -or $selectedOutputProfiles.Count -eq 0) {
            $selectedOutputProfiles = @('SolutionsEngineer')
        }

        $invokeParams = @{ AssessmentJsonPath = $jsonPath }
        $invokeParams.OutputProfile = $selectedOutputProfiles
        $invokeParams.ExportPath = if (-not [string]::IsNullOrWhiteSpace($exportPathInput)) { $exportPathInput } else { $defaultOutputRoot }
        if ($SkipPdfReport) { $invokeParams.SkipPdfReport = $true }
        if ($SkipJsonReport -and $IncludeLegacyAssessmentArtifacts) {
            $invokeParams.SkipJsonReport = $true
        }
        elseif ($SkipJsonReport) {
            Write-Warning 'The consolidated export workflow preserves a JSON snapshot for Improve and replay. Ignoring -SkipJsonReport unless -IncludeLegacyAssessmentArtifacts is used.'
        }
        if ($SkipImprove) { $invokeParams.SkipImprove = $true }
        if ($IncludeLegacyArtifacts) { $invokeParams.IncludeLegacyArtifacts = $true }
        if ($IncludeLegacyAssessmentArtifacts) { $invokeParams.IncludeLegacyAssessmentArtifacts = $true }
        if (-not [string]::IsNullOrWhiteSpace($ImproveOutputFolder)) { $invokeParams.ImproveOutputFolder = $ImproveOutputFolder }
        if ($UseGraphFallback) { $invokeParams.UseGraphFallback = $true }

        Write-LauncherActionHeader -Description 'Report + Improvement Plan from Existing Snapshot'
        Write-LauncherRunSummary -Fields ([ordered]@{
            'Snapshot'    = $jsonPath
            'Export path' = $invokeParams.ExportPath
            'Profile(s)'  = ($invokeParams.OutputProfile -join ', ')
            'Improve'     = if ($SkipImprove) { 'No (skipped)' } else { 'Yes (auto)' }
        })
        Invoke-M365TenantAssessmentExport @invokeParams
    }
    'AD' {
        Write-LauncherActionHeader -Description 'Active Directory Assessment'
        Invoke-ADTenantAssessment
    }
    'Improve' {
        $jsonPath = Read-LauncherFilePath -Prompt 'Path to tenant assessment JSON or manifest JSON' -CheckExists
        $outputFolder = Read-Host 'Output folder (leave blank to use JSON folder)'

        $invokeParams = @{ AssessmentJsonPath = $jsonPath }
        if (-not [string]::IsNullOrWhiteSpace($outputFolder)) { $invokeParams.OutputFolder = $outputFolder }
        if ($UseGraphFallback) { $invokeParams.UseGraphFallback = $true }
        if ($IncludeLegacyArtifacts) { $invokeParams.IncludeLegacyArtifacts = $true }

        Write-LauncherActionHeader -Description 'Improvement Plan Generation'
        Invoke-M365ImprovementPlan @invokeParams
    }
    'Compare' {
        $baselinePath = Read-LauncherFilePath -Prompt 'Path to BASELINE JSON' -CheckExists
        $currentPath  = Read-LauncherFilePath -Prompt 'Path to CURRENT JSON'  -CheckExists
        $outputFolder = Read-Host 'Output folder (leave blank to use current JSON folder)'

        $invokeParams = @{
            BaselineJsonPath = $baselinePath
            CurrentJsonPath  = $currentPath
        }
        if (-not [string]::IsNullOrWhiteSpace($outputFolder)) { $invokeParams.OutputFolder = $outputFolder }

        Write-LauncherActionHeader -Description 'Snapshot Comparison'
        Invoke-M365AssessmentComparison @invokeParams
    }
}
