[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('M365', 'M365Preflight', 'M365Collect', 'M365Export', 'AD', 'Improve', 'Compare')]
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

function Test-LauncherManifestHasSnapshotArtifact {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    if (-not (Test-Path -Path $ManifestPath -PathType Leaf)) {
        return $false
    }

    try {
        $resolvedManifestPath = (Resolve-Path -Path $ManifestPath -ErrorAction Stop).Path
        $manifest = Get-Content -Path $resolvedManifestPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -Depth 10 -ErrorAction Stop
    }
    catch {
        return $false
    }

    if (-not $manifest -or -not $manifest.PSObject.Properties['Artifacts']) {
        return $false
    }

    foreach ($artifact in @($manifest.Artifacts)) {
        if (-not $artifact) {
            continue
        }

        $artifactType = [string]$artifact.Type
        if ($artifactType -notin @('Assessment Snapshot JSON', 'JSON')) {
            continue
        }

        $artifactPath = [string]$artifact.Path
        if ([string]::IsNullOrWhiteSpace($artifactPath)) {
            continue
        }

        $resolvedArtifactPath = if ([System.IO.Path]::IsPathRooted($artifactPath)) {
            $artifactPath
        }
        else {
            Join-Path -Path (Split-Path -Path $resolvedManifestPath -Parent) -ChildPath $artifactPath
        }

        if (Test-Path -Path $resolvedArtifactPath -PathType Leaf) {
            return $true
        }
    }

    return $false
}

function Resolve-LauncherLatestManifestPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportPath
    )

    function Add-LauncherManifestCandidatesFromSupportDirectory {
        param(
            [Parameter(Mandatory = $true)]
            [string]$SupportPath,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.Generic.List[string]]$Candidates
        )

        $Candidates.Add((Join-Path -Path $SupportPath -ChildPath 'Run.manifest.json'))
        if (Test-Path -Path $SupportPath -PathType Container) {
            foreach ($manifest in @(Get-ChildItem -Path $SupportPath -Filter '*-Run.manifest.json' -File -ErrorAction SilentlyContinue)) {
                $Candidates.Add($manifest.FullName)
            }
        }
    }

    function Add-LauncherManifestCandidatesFromDirectory {
        param(
            [Parameter(Mandatory = $true)]
            [string]$DirectoryPath,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.Generic.List[string]]$Candidates
        )

        Add-LauncherManifestCandidatesFromSupportDirectory -SupportPath (Join-Path -Path $DirectoryPath -ChildPath 'Support') -Candidates $Candidates
        if ([string]::Equals((Split-Path -Path $DirectoryPath -Leaf), 'Deliverables', [System.StringComparison]::OrdinalIgnoreCase)) {
            $runRoot = Split-Path -Path $DirectoryPath -Parent
            if (-not [string]::IsNullOrWhiteSpace($runRoot)) {
                Add-LauncherManifestCandidatesFromSupportDirectory -SupportPath (Join-Path -Path $runRoot -ChildPath 'Support') -Candidates $Candidates
            }
        }
    }

    $candidateManifestPaths = New-Object System.Collections.Generic.List[string]
    $fullExportPath = [System.IO.Path]::GetFullPath($ExportPath)
    if (Test-Path -Path $fullExportPath -PathType Container) {
        Add-LauncherManifestCandidatesFromDirectory -DirectoryPath $fullExportPath -Candidates $candidateManifestPaths
    }
    $fullExportPathParent = Split-Path -Path $fullExportPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($fullExportPathParent)) {
        Add-LauncherManifestCandidatesFromDirectory -DirectoryPath $fullExportPathParent -Candidates $candidateManifestPaths
    }

    if (Test-Path -Path $fullExportPath -PathType Leaf) {
        if ($fullExportPath -match '\.manifest\.json$') {
            $candidateManifestPaths.Add($fullExportPath)
        }
        elseif ($fullExportPath -match '\.xlsx$') {
            $candidateManifestPaths.Add(($fullExportPath -replace '\.xlsx$', '.manifest.json'))
            $leafBaseName = [System.IO.Path]::GetFileNameWithoutExtension($fullExportPath)
            foreach ($suffix in @('-Assess', ' - Tenant Details', '-Tenant Details')) {
                if ($leafBaseName.EndsWith($suffix, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $manifestLeaf = '{0}-Run.manifest.json' -f $leafBaseName.Substring(0, $leafBaseName.Length - $suffix.Length)
                    $candidateManifestPaths.Add((Join-Path -Path (Join-Path -Path (Split-Path -Path $fullExportPath -Parent) -ChildPath 'Support') -ChildPath $manifestLeaf))
                    if (-not [string]::IsNullOrWhiteSpace($fullExportPathParent) -and [string]::Equals((Split-Path -Path $fullExportPathParent -Leaf), 'Deliverables', [System.StringComparison]::OrdinalIgnoreCase)) {
                        $runRoot = Split-Path -Path $fullExportPathParent -Parent
                        if (-not [string]::IsNullOrWhiteSpace($runRoot)) {
                            $candidateManifestPaths.Add((Join-Path -Path (Join-Path -Path $runRoot -ChildPath 'Support') -ChildPath $manifestLeaf))
                        }
                    }
                    break
                }
            }
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

    $existingManifestPaths = New-Object System.Collections.Generic.List[string]
    $seenManifestPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($candidatePath in $candidateManifestPaths) {
        if (Test-Path -Path $candidatePath -PathType Leaf) {
            $resolvedCandidatePath = (Resolve-Path -Path $candidatePath).Path
            if ($seenManifestPaths.Add($resolvedCandidatePath)) {
                $existingManifestPaths.Add($resolvedCandidatePath)
            }
        }
    }

    foreach ($candidatePath in $existingManifestPaths) {
        if (Test-LauncherManifestHasSnapshotArtifact -ManifestPath $candidatePath) {
            return $candidatePath
        }
    }

    if ($existingManifestPaths.Count -gt 0) {
        return $existingManifestPaths[0]
    }

    $searchedManifestPaths = New-Object System.Collections.Generic.List[string]
    $seenSearchedManifestPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
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
        foreach ($manifest in @(Get-ChildItem -Path $searchRoot -Recurse -Filter '*.manifest.json' -File -ErrorAction SilentlyContinue | Sort-Object -Property LastWriteTimeUtc -Descending)) {
            if ($seenSearchedManifestPaths.Add($manifest.FullName)) {
                $searchedManifestPaths.Add($manifest.FullName)
            }
        }

        foreach ($manifestPath in $searchedManifestPaths) {
            if (Test-LauncherManifestHasSnapshotArtifact -ManifestPath $manifestPath) {
                return $manifestPath
            }
        }

        if ($searchedManifestPaths.Count -gt 0) {
            return $searchedManifestPaths[0]
        }
    }

    return $null
}

function Resolve-LauncherRunRootFromManifestPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    $manifestDirectory = Split-Path -Path $ManifestPath -Parent
    if ([string]::Equals((Split-Path -Path $manifestDirectory -Leaf), 'Support', [System.StringComparison]::OrdinalIgnoreCase)) {
        return (Split-Path -Path $manifestDirectory -Parent)
    }

    return $manifestDirectory
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
        $resolvedOutputFolder = Resolve-LauncherRunRootFromManifestPath -ManifestPath $manifestPath
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
        'M365' { return (-not $SkipImprove) }
        'M365Collect' { return ($RunImprove -and -not $SkipImprove) }
        default { return $false }
    }
}

if ([string]::IsNullOrWhiteSpace($Action)) {
    Write-Host ''
    Write-Host 'Tenant Assessment Launcher' -ForegroundColor Cyan
    Write-Host '1. Microsoft 365 Full Tenant Assessment + Improvement Plan'
    Write-Host '2. Microsoft 365 Connection / Preflight Only'
    Write-Host '3. Microsoft 365 Data Collection Only (JSON snapshot)'
    Write-Host '4. Microsoft 365 Export from JSON Snapshot + Improvement Plan'
    Write-Host '5. Active Directory Assessment'
    Write-Host '6. Build Improvement Plan from Tenant JSON'
    Write-Host '7. Compare Two Tenant JSON Snapshots'
    Write-Host ''

    $choice = Read-Host 'Select an option (1-7)'
    switch ($choice) {
        '1' { $Action = 'M365' }
        '2' { $Action = 'M365Preflight' }
        '3' { $Action = 'M365Collect' }
        '4' { $Action = 'M365Export' }
        '5' { $Action = 'AD' }
        '6' { $Action = 'Improve' }
        '7' { $Action = 'Compare' }
        default { throw "Invalid selection: $choice" }
    }
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

        Invoke-M365TenantAssessment @invokeParams
    }
    'M365Preflight' {
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

        Invoke-M365TenantConnectionPreflight @invokeParams
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

        Invoke-M365TenantDataCollection @invokeParams
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
        if ($IncludeLegacyArtifacts) { $invokeParams.IncludeLegacyArtifacts = $true }

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
