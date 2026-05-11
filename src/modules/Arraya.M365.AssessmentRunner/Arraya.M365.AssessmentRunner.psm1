Set-StrictMode -Version Latest

$script:RepoRoot = (Resolve-Path -Path (Join-Path $PSScriptRoot '..\..\..')).Path
$script:AssessmentScriptRoots = @(
    (Join-Path $script:RepoRoot 'src\scripts\reporting'),
    (Join-Path $script:RepoRoot 'src\scripts\migrated\legacy')
)

function Test-AssessmentRunnerModuleMatchesManifestPath {
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

function Invoke-AssessmentScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        [Parameter(Mandatory = $false)]
        [hashtable]$Parameters
    )

    & {
        # Legacy assessment scripts were authored without strict mode and expect nullable properties.
        Set-StrictMode -Off

        if ($Parameters) {
            & $ScriptPath @Parameters
        }
        else {
            & $ScriptPath
        }
    }
}

function Import-AssessmentRunnerDependencies {
    [CmdletBinding()]
    param(
        [string[]]$RequiredCommands = @()
    )

    $commonManifestPath = Join-Path $script:RepoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
    $resolvedCommonManifestPath = (Resolve-Path -Path $commonManifestPath).Path
    $loadedCommonModule = Get-Module -Name 'Arraya.M365.Common' -ErrorAction SilentlyContinue | Select-Object -First 1
    $requiredCommonCommands = @(
        'Import-ArrayaOffice365CustomLocal',
        'Get-ArrayaAssessmentOutputRoot',
        'Get-ArrayaAssessmentOutputProfilePolicy'
    )
    $missingCommonCommands = @(
        $requiredCommonCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
    )
    if (
        -not $loadedCommonModule -or
        -not (Test-AssessmentRunnerModuleMatchesManifestPath -Module $loadedCommonModule -ManifestPath $resolvedCommonManifestPath) -or
        $missingCommonCommands.Count -gt 0
    ) {
        Import-Module -Name $resolvedCommonManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
    }

    Import-ArrayaOffice365CustomLocal -RepoRoot $script:RepoRoot -RequiredCommands $RequiredCommands | Out-Null
}

function Resolve-AssessmentScriptPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    foreach ($scriptRoot in $script:AssessmentScriptRoots) {
        $scriptPath = Join-Path $scriptRoot $Name
        if (Test-Path -Path $scriptPath) {
            return (Resolve-Path -Path $scriptPath).Path
        }
    }

    throw "Assessment script not found in configured roots: $Name"
}

function Resolve-AssessmentExportPathInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExportPath
    )

    if (-not (Get-Command -Name 'Get-ArrayaAssessmentOutputRoot' -ErrorAction SilentlyContinue)) {
        Import-AssessmentRunnerDependencies
    }

    if ([string]::IsNullOrWhiteSpace($ExportPath)) {
        return (Get-ArrayaAssessmentOutputRoot -FallbackPath $script:RepoRoot)
    }

    return $ExportPath
}

function Test-AssessmentManifestHasSnapshotArtifact {
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

function Resolve-AssessmentLatestManifestPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportPath
    )

    function Add-AssessmentManifestCandidatesFromSupportDirectory {
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

    function Add-AssessmentManifestCandidatesFromDirectory {
        param(
            [Parameter(Mandatory = $true)]
            [string]$DirectoryPath,
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.Generic.List[string]]$Candidates
        )

        Add-AssessmentManifestCandidatesFromSupportDirectory -SupportPath (Join-Path -Path $DirectoryPath -ChildPath 'Support') -Candidates $Candidates
        if ([string]::Equals((Split-Path -Path $DirectoryPath -Leaf), 'Deliverables', [System.StringComparison]::OrdinalIgnoreCase)) {
            $runRoot = Split-Path -Path $DirectoryPath -Parent
            if (-not [string]::IsNullOrWhiteSpace($runRoot)) {
                Add-AssessmentManifestCandidatesFromSupportDirectory -SupportPath (Join-Path -Path $runRoot -ChildPath 'Support') -Candidates $Candidates
            }
        }
    }

    $candidateManifestPaths = New-Object System.Collections.Generic.List[string]
    $fullExportPath = [System.IO.Path]::GetFullPath($ExportPath)
    if (Test-Path -Path $fullExportPath -PathType Container) {
        Add-AssessmentManifestCandidatesFromDirectory -DirectoryPath $fullExportPath -Candidates $candidateManifestPaths
    }
    $fullExportPathParent = Split-Path -Path $fullExportPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($fullExportPathParent)) {
        Add-AssessmentManifestCandidatesFromDirectory -DirectoryPath $fullExportPathParent -Candidates $candidateManifestPaths
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
        if (Test-AssessmentManifestHasSnapshotArtifact -ManifestPath $candidatePath) {
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
            if (Test-AssessmentManifestHasSnapshotArtifact -ManifestPath $manifestPath) {
                return $manifestPath
            }
        }

        if ($searchedManifestPaths.Count -gt 0) {
            return $searchedManifestPaths[0]
        }
    }

    return $null
}

function Resolve-AssessmentRunRootFromManifestPath {
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

function Resolve-AssessmentImproveOutputFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,
        [Parameter(Mandatory = $true)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [string]$OutputFolder
    )

    $runRoot = Resolve-AssessmentRunRootFromManifestPath -ManifestPath $ManifestPath
    $resolvedRunRoot = if (Test-Path -Path $runRoot) {
        (Resolve-Path -Path $runRoot).Path
    }
    else {
        [System.IO.Path]::GetFullPath($runRoot)
    }

    if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
        return $resolvedRunRoot
    }

    $resolvedOutputFolder = if (Test-Path -Path $OutputFolder) {
        (Resolve-Path -Path $OutputFolder).Path
    }
    else {
        [System.IO.Path]::GetFullPath($OutputFolder)
    }

    $resolvedExportPath = [System.IO.Path]::GetFullPath($ExportPath)
    if ([string]::Equals($resolvedOutputFolder, $resolvedExportPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $resolvedRunRoot
    }

    return $resolvedOutputFolder
}

function Update-AssessmentArtifactManifestWithImproveOutputs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$ImproveResult,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLegacyArtifacts
    )

    if (-not (Test-Path -Path $ManifestPath -PathType Leaf)) {
        return
    }

    $manifest = Get-Content -Path $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 10
    if (-not $manifest) {
        return
    }

    $artifactEntries = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$ImproveResult.CustomerAssessmentReportPath)) {
        $artifactEntries += [pscustomobject][ordered]@{
            Type = 'Customer Assessment Report'
            Path = [string]$ImproveResult.CustomerAssessmentReportPath
        }
    }
    elseif ($ImproveResult.PSObject.Properties.Name -contains 'CustomerAssessmentReportPaths') {
        $reportIndex = 0
        foreach ($reportPath in @($ImproveResult.CustomerAssessmentReportPaths | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)) {
            $reportIndex++
            $artifactEntries += [pscustomobject][ordered]@{
                Type = ('Customer Assessment Report {0}' -f $reportIndex)
                Path = [string]$reportPath
            }
        }
    }
    if ($ImproveResult.PSObject.Properties.Name -contains 'RoadmapRemediationPlanPath' -and -not [string]::IsNullOrWhiteSpace([string]$ImproveResult.RoadmapRemediationPlanPath)) {
        $artifactEntries += [pscustomobject][ordered]@{
            Type = 'Roadmap Remediation Plan'
            Path = [string]$ImproveResult.RoadmapRemediationPlanPath
        }
    }
    elseif ($ImproveResult.PSObject.Properties.Name -contains 'RoadmapRemediationPlanPaths') {
        $roadmapIndex = 0
        foreach ($roadmapPath in @($ImproveResult.RoadmapRemediationPlanPaths | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)) {
            $roadmapIndex++
            $artifactEntries += [pscustomobject][ordered]@{
                Type = ('Roadmap Remediation Plan {0}' -f $roadmapIndex)
                Path = [string]$roadmapPath
            }
        }
    }
    $artifactEntries += [pscustomobject][ordered]@{
        Type = 'Engineer Action Pack'
        Path = [string]$ImproveResult.EngineerActionPackPath
    }
    $artifactEntries += [pscustomobject][ordered]@{
        Type = 'Improvement Plan JSON'
        Path = [string]$ImproveResult.JsonPath
    }
    $artifactEntries += [pscustomobject][ordered]@{
        Type = 'Remediation Snippets'
        Path = [string]$ImproveResult.RemediationPs1Path
    }
    if ($IncludeLegacyArtifacts) {
        if (-not [string]::IsNullOrWhiteSpace([string]$ImproveResult.CsvPath)) {
            $artifactEntries += [pscustomobject][ordered]@{
                Type = 'Improvement Plan CSV'
                Path = [string]$ImproveResult.CsvPath
            }
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$ImproveResult.MarkdownPath)) {
            $artifactEntries += [pscustomobject][ordered]@{
                Type = 'Improvement Plan Markdown'
                Path = [string]$ImproveResult.MarkdownPath
            }
        }
    }

    $artifactRows = @()
    $newArtifactPaths = @{}
    foreach ($entry in @($artifactEntries)) {
        if ($null -eq $entry) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$entry.Path)) { continue }
        $newArtifactPaths[[string]$entry.Path] = $true
    }
    if ($manifest.Artifacts) {
        foreach ($artifact in @($manifest.Artifacts)) {
            if (-not $artifact) { continue }
            if ($newArtifactPaths.ContainsKey([string]$artifact.Path)) { continue }
            $artifactRows += [pscustomobject][ordered]@{
                Type      = [string]$artifact.Type
                Path      = [string]$artifact.Path
                Exists    = [bool]$artifact.Exists
                SizeBytes = $artifact.SizeBytes
            }
        }
    }

    foreach ($entry in @($artifactEntries)) {
        if ($null -eq $entry) {
            continue
        }
        if ([string]::IsNullOrWhiteSpace([string]$entry.Path)) {
            continue
        }

        $artifactPath = [string]$entry.Path
        $exists = Test-Path -Path $artifactPath
        $sizeBytes = $null
        if ($exists) {
            try {
                $sizeBytes = (Get-Item -Path $artifactPath -ErrorAction Stop).Length
            }
            catch {}
        }

        $artifactRows += [pscustomobject][ordered]@{
            Type      = [string]$entry.Type
            Path      = $artifactPath
            Exists    = [bool]$exists
            SizeBytes = $sizeBytes
        }
    }

    $manifestData = [ordered]@{}
    foreach ($property in $manifest.PSObject.Properties) {
        if ([string]$property.Name -eq 'Artifacts') {
            continue
        }
        $manifestData[[string]$property.Name] = $property.Value
    }
    $manifestData['Artifacts'] = @($artifactRows)
    try {
        if (-not (Get-Command -Name 'New-ArrayaAssessmentOperatorSummary' -ErrorAction SilentlyContinue)) {
            $commonManifestPath = Join-Path $script:RepoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
            if (Test-Path -Path $commonManifestPath -PathType Leaf) {
                Import-Module -Name $commonManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
            }
        }

        if (Get-Command -Name 'New-ArrayaAssessmentOperatorSummary' -ErrorAction SilentlyContinue) {
            $outputProfileLabel = if ($manifest.PSObject.Properties['OutputProfile']) { [string]$manifest.OutputProfile } else { '' }
            $reportingMode = if ($manifest.PSObject.Properties['ReportingMode']) { [string]$manifest.ReportingMode } else { '' }
            $collectionOnly = if ($manifest.PSObject.Properties['CollectionOnly']) { [bool]$manifest.CollectionOnly } else { $false }
            $exportOnly = if ($manifest.PSObject.Properties['ExportOnly']) { [bool]$manifest.ExportOnly } else { $false }

            $manifestData['OperatorSummary'] = New-ArrayaAssessmentOperatorSummary `
                -Artifacts $artifactRows `
                -OutputProfileLabel $outputProfileLabel `
                -ReportingMode $reportingMode `
                -CollectionOnly $collectionOnly `
                -ExportOnly $exportOnly
        }
    }
    catch {
        # Preserve manifest update behavior even if the operator summary cannot be refreshed.
    }

    $json = ([pscustomobject]$manifestData) | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($ManifestPath, $json, [System.Text.UTF8Encoding]::new($false))
}

function Invoke-M365ImproveForAssessmentRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [string]$OutputFolder,
        [Parameter(Mandatory = $false)]
        [Alias('LiveRefresh')]
        [switch]$UseGraphFallback,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLegacyArtifacts
    )

    $manifestPath = Resolve-AssessmentLatestManifestPath -ExportPath $ExportPath
    if ([string]::IsNullOrWhiteSpace($manifestPath)) {
        throw "Could not find a run manifest for the completed assessment under: $ExportPath. The workbook alone is not enough for Improve; rerun a JSON-enabled output profile and check the export log for manifest write warnings."
    }

    try {
        if (-not (Get-Command -Name 'Import-ArrayaTenantSnapshotContext' -ErrorAction SilentlyContinue)) {
            Import-AssessmentRunnerDependencies -RequiredCommands @('Import-ArrayaTenantSnapshotContext')
        }

        Import-ArrayaTenantSnapshotContext -Path $manifestPath -Purpose ImprovementPlan | Out-Null
    }
    catch {
        throw ("The completed assessment run did not produce a usable JSON snapshot, so Improve cannot continue from the manifest. This usually means snapshot export failed earlier in the run even if the workbook was created. Review the earlier export errors in the run log and rerun after the snapshot issue is corrected. Manifest: {0} Underlying error: {1}" -f $manifestPath, $_.Exception.Message)
    }

    $resolvedOutputFolder = Resolve-AssessmentImproveOutputFolder -ManifestPath $manifestPath -ExportPath $ExportPath -OutputFolder $OutputFolder

    $improveResult = Invoke-M365ImprovementPlan `
        -AssessmentJsonPath $manifestPath `
        -OutputFolder $resolvedOutputFolder `
        -UseGraphFallback:$UseGraphFallback `
        -IncludeLegacyArtifacts:$IncludeLegacyArtifacts `
        -PassThru

    if ($improveResult) {
        Update-AssessmentArtifactManifestWithImproveOutputs -ManifestPath $manifestPath -ImproveResult $improveResult -IncludeLegacyArtifacts:$IncludeLegacyArtifacts
    }

    return $improveResult
}

function Resolve-M365OutputProfileExecutionPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$OutputProfile = @('SolutionsEngineer')
    )

    if (-not (Get-Command -Name 'Get-ArrayaAssessmentOutputProfilePolicy' -ErrorAction SilentlyContinue)) {
        Import-AssessmentRunnerDependencies
    }

    $selectedOutputProfiles = New-Object System.Collections.Generic.List[string]
    $selectedOutputProfileSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
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
            if ($selectedOutputProfileSet.Add($profile)) {
                $selectedOutputProfiles.Add($profile)
            }
        }
    }

    if ($selectedOutputProfiles.Count -eq 0) {
        $selectedOutputProfiles.Add('SolutionsEngineer')
    }

    $selectedPolicies = @()
    foreach ($profile in $selectedOutputProfiles) {
        $selectedPolicies += Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile $profile
    }

    $reportingModeRank = @{
        'Minimum' = 1
        'Operator' = 2
        'Combined' = 2
        'Automation' = 3
        'Geek' = 4
        'All' = 5
    }
    $resolvedReportingMode = 'Minimum'
    $resolvedReportingRank = 0
    foreach ($policy in $selectedPolicies) {
        $mode = [string]$policy.ReportingMode
        if (-not $reportingModeRank.ContainsKey($mode)) {
            continue
        }

        $rank = [int]$reportingModeRank[$mode]
        if ($rank -gt $resolvedReportingRank) {
            $resolvedReportingRank = $rank
            $resolvedReportingMode = $mode
        }
    }

    $mergedGenerateWorkbook = $false
    $mergedGenerateTechnicalHtml = $false
    $mergedGenerateBestPracticesHtml = $false
    $mergedGenerateQuestionnaire = $false
    $mergedGenerateJson = $false
    $mergedGeneratePdf = $false
    $mergedGenerateMigrationPack = $false
    foreach ($policy in $selectedPolicies) {
        $mergedGenerateWorkbook = $mergedGenerateWorkbook -or [bool]$policy.GenerateWorkbook
        $mergedGenerateTechnicalHtml = $mergedGenerateTechnicalHtml -or [bool]$policy.GenerateTechnicalHtml
        $mergedGenerateBestPracticesHtml = $mergedGenerateBestPracticesHtml -or [bool]$policy.GenerateBestPracticesHtml
        $mergedGenerateQuestionnaire = $mergedGenerateQuestionnaire -or [bool]$policy.GenerateQuestionnaire
        $mergedGenerateJson = $mergedGenerateJson -or [bool]$policy.GenerateJson
        $mergedGeneratePdf = $mergedGeneratePdf -or [bool]$policy.GeneratePdf
        $mergedGenerateMigrationPack = $mergedGenerateMigrationPack -or [bool]$policy.GenerateMigrationPack
    }

    $primaryProfile = [string]$selectedOutputProfiles[0]
    $primaryPolicy = $selectedPolicies | Select-Object -First 1
    $profileLabel = if ($selectedOutputProfiles.Count -gt 1) {
        "Merged({0})" -f (($selectedOutputProfiles.ToArray() -join '+'))
    } else {
        $primaryProfile
    }

    return [PSCustomObject]@{
        SelectedOutputProfiles      = $selectedOutputProfiles.ToArray()
        IsMergedSelection           = ($selectedOutputProfiles.Count -gt 1)
        PrimaryProfile              = $primaryProfile
        ProfileLabel                = $profileLabel
        ReportingMode               = $resolvedReportingMode
        GenerateWorkbook            = $mergedGenerateWorkbook
        GenerateTechnicalHtml       = $mergedGenerateTechnicalHtml
        GenerateBestPracticesHtml   = $mergedGenerateBestPracticesHtml
        GenerateQuestionnaire       = $mergedGenerateQuestionnaire
        GenerateJson                = $mergedGenerateJson
        GeneratePdf                 = $mergedGeneratePdf
        WorkbookExportPolicy        = if ($primaryPolicy -and $primaryPolicy.PSObject.Properties['WorkbookExportPolicy']) { [string]$primaryPolicy.WorkbookExportPolicy } else { 'Default' }
        TechnicalHtmlPolicy         = if ($primaryPolicy -and $primaryPolicy.PSObject.Properties['TechnicalHtmlPolicy']) { [string]$primaryPolicy.TechnicalHtmlPolicy } else { 'Default' }
        GenerateMigrationPack       = $mergedGenerateMigrationPack
        CollectionScopePolicy       = if ($selectedOutputProfiles.Count -gt 1) { 'Default' } elseif ($primaryPolicy -and $primaryPolicy.PSObject.Properties['CollectionScopePolicy']) { [string]$primaryPolicy.CollectionScopePolicy } else { 'Default' }
        SkipImproveByDefault        = ($selectedOutputProfiles.Count -eq 1 -and [string]::Equals($primaryProfile, 'TenantToTenantMigration', [System.StringComparison]::OrdinalIgnoreCase))
    }
}

function Test-AssessmentPlanSkipsImproveByDefault {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Plan
    )

    if (-not $Plan) {
        return $false
    }

    if ($Plan.PSObject.Properties.Name -contains 'SkipImproveByDefault') {
        return [bool]$Plan.SkipImproveByDefault
    }

    return (
        -not [bool]$Plan.IsMergedSelection -and
        [string]::Equals([string]$Plan.PrimaryProfile, 'TenantToTenantMigration', [System.StringComparison]::OrdinalIgnoreCase)
    )
}

function Invoke-M365TenantWorkflow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Full', 'CollectOnly', 'ExportOnly', 'PreflightOnly')]
        [string]$Mode,
        [Parameter(Mandatory = $false)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
        [string[]]$OutputProfile = @('SolutionsEngineer'),
        [Parameter(Mandatory = $false)]
        [string]$AssessmentJsonPath,
        [Parameter(Mandatory = $false)]
        [switch]$SkipHtmlReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPdfReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipJsonReport,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLegacyAssessmentArtifacts,
        [Parameter(Mandatory = $false)]
        [switch]$StoreTenantStatsGlobal,
        [Parameter(Mandatory = $false)]
        [string]$TenantStatsVariableName = 'ArrayaTenantStats',
        [Parameter(Mandatory = $false)]
        [switch]$SkipAuth,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPermissionPreflight,
        [Parameter(Mandatory = $false)]
        [switch]$UseExistingConnections,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthMode,
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
        [switch]$PassThruInvocation
    )

    $plan = Resolve-M365OutputProfileExecutionPlan -OutputProfile $OutputProfile
    $scriptPath = Resolve-AssessmentScriptPath -Name 'Get-FullTenantReportDetails.ps1'

    if ($plan.IsMergedSelection) {
        $modeLabel = switch ($Mode) {
            'CollectOnly' { 'collection pass' }
            'ExportOnly' { 'export pass' }
            'PreflightOnly' { 'preflight pass' }
            default { 'pass' }
        }
        Write-Host ("Running merged profile {0} for: {1}" -f $modeLabel, ($plan.SelectedOutputProfiles -join ', ')) -ForegroundColor Cyan
        Write-Host ("Merged reporting mode: {0}" -f $plan.ReportingMode) -ForegroundColor DarkCyan
    }

    $invokeParams = @{}
    $invokeParams.ExportPath = Resolve-AssessmentExportPathInput -ExportPath $ExportPath
    $invokeParams.OutputProfile = $plan.PrimaryProfile
    $invokeParams.ReportingModeOverride = $plan.ReportingMode
    $invokeParams.WorkbookExportPolicyOverride = [string]$plan.WorkbookExportPolicy
    $invokeParams.TechnicalHtmlPolicyOverride = [string]$plan.TechnicalHtmlPolicy
    $invokeParams.GenerateMigrationPackOverride = [bool]$plan.GenerateMigrationPack
    $invokeParams.CollectionScopePolicyOverride = [string]$plan.CollectionScopePolicy

    $shouldEnablePolicyTechnicalHtml = ([string]$plan.TechnicalHtmlPolicy -eq 'TenantToTenantCutover')

    switch ($Mode) {
        'Full' {
            $invokeParams.OutputProfileLabel = $plan.ProfileLabel
            $invokeParams.GenerateWorkbookOverride = [bool]$plan.GenerateWorkbook
            $invokeParams.GenerateTechnicalHtmlOverride = if ($IncludeLegacyAssessmentArtifacts -or $shouldEnablePolicyTechnicalHtml) { [bool]$plan.GenerateTechnicalHtml } else { $false }
            $invokeParams.GenerateBestPracticesHtmlOverride = if ($IncludeLegacyAssessmentArtifacts) { [bool]$plan.GenerateBestPracticesHtml } else { $false }
            $invokeParams.GenerateQuestionnaireOverride = if ($IncludeLegacyAssessmentArtifacts) { [bool]$plan.GenerateQuestionnaire } else { $false }
            $invokeParams.GenerateJsonOverride = [bool]$plan.GenerateJson
            if (-not $invokeParams.GenerateJsonOverride) {
                $invokeParams.GenerateJsonOverride = $true
            }
            $invokeParams.GeneratePdfOverride = if ($IncludeLegacyAssessmentArtifacts) { [bool]$plan.GeneratePdf } else { $false }
            if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
            if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
            if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }
            if ($PSBoundParameters.ContainsKey('StoreTenantStatsGlobal')) { $invokeParams.StoreTenantStatsGlobal = $StoreTenantStatsGlobal }
            if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
            if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
            if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight')) { $invokeParams.SkipPermissionPreflight = $SkipPermissionPreflight }
            if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
            if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
            if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
            if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
            if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }
            if ($PSBoundParameters.ContainsKey('ClientSecretCredential')) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
            if ($PSBoundParameters.ContainsKey('ClientSecretSecure')) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }
        }
        'CollectOnly' {
            $invokeParams.OutputProfileLabel = "Collection-$($plan.ProfileLabel)"
            $invokeParams.GenerateWorkbookOverride = $false
            $invokeParams.GenerateTechnicalHtmlOverride = $false
            $invokeParams.GenerateBestPracticesHtmlOverride = $false
            $invokeParams.GenerateQuestionnaireOverride = $false
            $invokeParams.GenerateJsonOverride = $true
            $invokeParams.GeneratePdfOverride = $false
            $invokeParams.GenerateMigrationPackOverride = $false
            $invokeParams.DataCollectionOnly = $true
            if ($PSBoundParameters.ContainsKey('UseExistingConnections')) { $invokeParams.UseExistingConnections = $UseExistingConnections }
            if ($PSBoundParameters.ContainsKey('StoreTenantStatsGlobal')) { $invokeParams.StoreTenantStatsGlobal = $StoreTenantStatsGlobal }
            if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
            if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
            if (-not $UseExistingConnections) {
                if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
                if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight')) { $invokeParams.SkipPermissionPreflight = $SkipPermissionPreflight }
                if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
                if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
                if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
                if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }
                if ($PSBoundParameters.ContainsKey('ClientSecretCredential')) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
                if ($PSBoundParameters.ContainsKey('ClientSecretSecure')) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }
            }
        }
        'ExportOnly' {
            if ([string]::IsNullOrWhiteSpace($AssessmentJsonPath)) {
                throw 'AssessmentJsonPath is required when Mode is ExportOnly.'
            }
            $invokeParams.OutputProfileLabel = "Export-$($plan.ProfileLabel)"
            $invokeParams.GenerateWorkbookOverride = [bool]$plan.GenerateWorkbook
            $invokeParams.GenerateTechnicalHtmlOverride = if ($IncludeLegacyAssessmentArtifacts -or $shouldEnablePolicyTechnicalHtml) { [bool]$plan.GenerateTechnicalHtml } else { $false }
            $invokeParams.GenerateBestPracticesHtmlOverride = if ($IncludeLegacyAssessmentArtifacts) { [bool]$plan.GenerateBestPracticesHtml } else { $false }
            $invokeParams.GenerateQuestionnaireOverride = if ($IncludeLegacyAssessmentArtifacts) { [bool]$plan.GenerateQuestionnaire } else { $false }
            $invokeParams.GenerateJsonOverride = [bool]$plan.GenerateJson
            if (-not $invokeParams.GenerateJsonOverride) {
                $invokeParams.GenerateJsonOverride = $true
            }
            $invokeParams.GeneratePdfOverride = if ($IncludeLegacyAssessmentArtifacts) { [bool]$plan.GeneratePdf } else { $false }
            $invokeParams.ExportOnly = $true
            $invokeParams.TenantStatsJsonPath = $AssessmentJsonPath
            if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
            if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
            if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }
        }
        'PreflightOnly' {
            $invokeParams.OutputProfileLabel = "Preflight-$($plan.ProfileLabel)"
            $invokeParams.GenerateWorkbookOverride = $false
            $invokeParams.GenerateTechnicalHtmlOverride = $false
            $invokeParams.GenerateBestPracticesHtmlOverride = $false
            $invokeParams.GenerateQuestionnaireOverride = $false
            $invokeParams.GenerateJsonOverride = $false
            $invokeParams.GeneratePdfOverride = $false
            $invokeParams.GenerateMigrationPackOverride = $false
            $invokeParams.PreflightOnly = $true
            if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
            if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight')) { $invokeParams.SkipPermissionPreflight = $SkipPermissionPreflight }
            if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
            if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
            if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
            if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
            if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }
            if ($PSBoundParameters.ContainsKey('ClientSecretCredential')) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
            if ($PSBoundParameters.ContainsKey('ClientSecretSecure')) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }
        }
    }

    if ($PassThruInvocation) {
        return [pscustomobject]@{
            Mode       = $Mode
            ScriptPath = $scriptPath
            Parameters = $invokeParams
            Plan       = $plan
        }
    }

    Invoke-AssessmentScript -ScriptPath $scriptPath -Parameters $invokeParams
}

function Invoke-M365TenantAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
        [string[]]$OutputProfile = @('SolutionsEngineer'),
        [Parameter(Mandatory = $false)]
        [switch]$SkipHtmlReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPdfReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipJsonReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipImprove,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLegacyArtifacts,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLegacyAssessmentArtifacts,
        [Parameter(Mandatory = $false)]
        [string]$ImproveOutputFolder,
        [Parameter(Mandatory = $false)]
        [Alias('LiveRefresh')]
        [switch]$UseGraphFallback,
        [Parameter(Mandatory = $false)]
        [switch]$StoreTenantStatsGlobal,
        [Parameter(Mandatory = $false)]
        [string]$TenantStatsVariableName = 'ArrayaTenantStats',
        [Parameter(Mandatory = $false)]
        [switch]$SkipAuth,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPermissionPreflight,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthMode,
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
        [securestring]$ClientSecretSecure
    )

    $plan = Resolve-M365OutputProfileExecutionPlan -OutputProfile $OutputProfile
    $invokeParams = @{}
    if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
    if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
    if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
    if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
    if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }
    if ($PSBoundParameters.ContainsKey('StoreTenantStatsGlobal')) { $invokeParams.StoreTenantStatsGlobal = $StoreTenantStatsGlobal }
    if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
    if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
    if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight')) { $invokeParams.SkipPermissionPreflight = $SkipPermissionPreflight }
    if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
    if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
    if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
    if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
    if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }
    if ($PSBoundParameters.ContainsKey('ClientSecretCredential')) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
    if ($PSBoundParameters.ContainsKey('ClientSecretSecure')) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }

    if ($PSBoundParameters.ContainsKey('IncludeLegacyAssessmentArtifacts')) { $invokeParams.IncludeLegacyAssessmentArtifacts = $IncludeLegacyAssessmentArtifacts }
    Invoke-M365TenantWorkflow -Mode Full @invokeParams

    if (Test-AssessmentPlanSkipsImproveByDefault -Plan $plan) {
        if (-not $SkipImprove) {
            Write-Host 'Tenant-to-tenant migration profile defaults to workbook, cutover pack, technical HTML, and JSON snapshot output. Skipping Improve/customer-style follow-up artifacts unless requested separately.' -ForegroundColor DarkCyan
        }
        return
    }

    if (-not $SkipImprove) {
        $resolvedExportPath = Resolve-AssessmentExportPathInput -ExportPath $ExportPath
        Invoke-M365ImproveForAssessmentRun -ExportPath $resolvedExportPath -OutputFolder $ImproveOutputFolder -UseGraphFallback:$UseGraphFallback -IncludeLegacyArtifacts:$IncludeLegacyArtifacts | Out-Null
    }
}

function Invoke-M365TenantConnectionPreflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
        [string[]]$OutputProfile = @('SolutionsEngineer'),
        [Parameter(Mandatory = $false)]
        [switch]$SkipAuth,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPermissionPreflight,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthMode,
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
        [securestring]$ClientSecretSecure
    )

    $invokeParams = @{}
    if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
    if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
    if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
    if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight')) { $invokeParams.SkipPermissionPreflight = $SkipPermissionPreflight }
    if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
    if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
    if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
    if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
    if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }
    if ($PSBoundParameters.ContainsKey('ClientSecretCredential')) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
    if ($PSBoundParameters.ContainsKey('ClientSecretSecure')) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }

    Invoke-M365TenantWorkflow -Mode PreflightOnly @invokeParams
}

function Invoke-M365TenantDataCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
        [string[]]$OutputProfile = @('SolutionsEngineer'),
        [Parameter(Mandatory = $false)]
        [switch]$RunImprove,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLegacyArtifacts,
        [Parameter(Mandatory = $false)]
        [string]$ImproveOutputFolder,
        [Parameter(Mandatory = $false)]
        [Alias('LiveRefresh')]
        [switch]$UseGraphFallback,
        [Parameter(Mandatory = $false)]
        [switch]$StoreTenantStatsGlobal,
        [Parameter(Mandatory = $false)]
        [string]$TenantStatsVariableName = 'ArrayaTenantStats',
        [Parameter(Mandatory = $false)]
        [switch]$SkipAuth,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPermissionPreflight,
        [Parameter(Mandatory = $false)]
        [switch]$UseExistingConnections,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthMode,
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
        [securestring]$ClientSecretSecure
    )

    $invokeParams = @{}
    if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
    if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
    if ($PSBoundParameters.ContainsKey('StoreTenantStatsGlobal')) { $invokeParams.StoreTenantStatsGlobal = $StoreTenantStatsGlobal }
    if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
    if ($PSBoundParameters.ContainsKey('UseExistingConnections')) { $invokeParams.UseExistingConnections = $UseExistingConnections }
    if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
    if (-not $UseExistingConnections) {
        if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
        if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight')) { $invokeParams.SkipPermissionPreflight = $SkipPermissionPreflight }
        if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
        if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
        if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
        if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }
        if ($PSBoundParameters.ContainsKey('ClientSecretCredential')) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
        if ($PSBoundParameters.ContainsKey('ClientSecretSecure')) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }
    }

    Invoke-M365TenantWorkflow -Mode CollectOnly @invokeParams

    if ($RunImprove) {
        $resolvedExportPath = Resolve-AssessmentExportPathInput -ExportPath $ExportPath
        Invoke-M365ImproveForAssessmentRun -ExportPath $resolvedExportPath -OutputFolder $ImproveOutputFolder -UseGraphFallback:$UseGraphFallback -IncludeLegacyArtifacts:$IncludeLegacyArtifacts | Out-Null
    }
}

function Invoke-M365TenantAssessmentExport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AssessmentJsonPath,
        [Parameter(Mandatory = $false)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
        [string[]]$OutputProfile = @('SolutionsEngineer'),
        [Parameter(Mandatory = $false)]
        [switch]$SkipHtmlReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPdfReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipJsonReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipImprove,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLegacyArtifacts,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLegacyAssessmentArtifacts,
        [Parameter(Mandatory = $false)]
        [string]$ImproveOutputFolder,
        [Parameter(Mandatory = $false)]
        [Alias('LiveRefresh')]
        [switch]$UseGraphFallback
    )

    $plan = Resolve-M365OutputProfileExecutionPlan -OutputProfile $OutputProfile
    $invokeParams = @{ AssessmentJsonPath = $AssessmentJsonPath }
    if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
    if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
    if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
    if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
    if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }
    if ($PSBoundParameters.ContainsKey('IncludeLegacyAssessmentArtifacts')) { $invokeParams.IncludeLegacyAssessmentArtifacts = $IncludeLegacyAssessmentArtifacts }
    Invoke-M365TenantWorkflow -Mode ExportOnly @invokeParams

    if (Test-AssessmentPlanSkipsImproveByDefault -Plan $plan) {
        if (-not $SkipImprove) {
            Write-Host 'Tenant-to-tenant migration export defaults to workbook, cutover pack, technical HTML, and JSON snapshot output. Skipping Improve/customer-style follow-up artifacts unless requested separately.' -ForegroundColor DarkCyan
        }
        return
    }

    if (-not $SkipImprove) {
        $resolvedExportPath = Resolve-AssessmentExportPathInput -ExportPath $ExportPath
        Invoke-M365ImproveForAssessmentRun -ExportPath $resolvedExportPath -OutputFolder $ImproveOutputFolder -UseGraphFallback:$UseGraphFallback -IncludeLegacyArtifacts:$IncludeLegacyArtifacts | Out-Null
    }
}

function Invoke-ADTenantAssessment {
    [CmdletBinding()]
    param()

    $scriptPath = Resolve-AssessmentScriptPath -Name 'Get-ActiveDirectoryReport.ps1'
    Invoke-AssessmentScript -ScriptPath $scriptPath
}

function Invoke-GraphActivityAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet(
            'SharePointSites',
            'MailboxUsage',
            'SharePointUser',
            'TeamsUser',
            'OneDriveUsage',
            'OneDriveActivity',
            'EmailActivity',
            'YammerActivity',
            'BrowserUsage',
            'SkypeForBusinessActivity',
            'Office365ActiveUser',
            'Office365ServicesUserCounts',
            'Office365ActivationCounts',
            'Office365ActivationUser',
            'Office365ActiveUserCounts',
            'Office365GroupsActivity',
            'Office365GroupsActivityCounts',
            'Office365GroupsActivityFileCounts',
            'Office365GroupsActivityGroupCounts',
            'Office365GroupsActivityStorage',
            'Office365GroupsActivityUser'
        )]
        [string]$ServiceName,
        [Parameter(Mandatory = $false)]
        [ValidateSet('D7', 'D30', 'D90', 'D180')]
        [string]$PeriodDuration = 'D90',
        [Parameter(Mandatory = $false)]
        [switch]$UseBeta
    )

    Import-AssessmentRunnerDependencies -RequiredCommands @('Office365Custom\Get-GraphAPIActivityReport')
    Office365Custom\Get-GraphAPIActivityReport -ServiceName $ServiceName -PeriodDuration $PeriodDuration -UseBeta:$UseBeta
}

function Invoke-M365ImprovementPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AssessmentJsonPath,
        [Parameter(Mandatory = $false)]
        [string]$OutputFolder,
        [Parameter(Mandatory = $false)]
        [Alias('LiveRefresh')]
        [switch]$UseGraphFallback,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLegacyArtifacts,
        [Parameter(Mandatory = $false)]
        [switch]$PassThru
    )

    $scriptPath = Resolve-AssessmentScriptPath -Name 'New-M365TenantImprovementPlan.ps1'
    $invokeParams = @{ AssessmentJsonPath = $AssessmentJsonPath }
    if ($PSBoundParameters.ContainsKey('OutputFolder')) { $invokeParams.OutputFolder = $OutputFolder }
    if ($PSBoundParameters.ContainsKey('UseGraphFallback')) { $invokeParams.UseGraphFallback = $UseGraphFallback }
    if ($PSBoundParameters.ContainsKey('IncludeLegacyArtifacts')) { $invokeParams.IncludeLegacyArtifacts = $IncludeLegacyArtifacts }
    if ($PSBoundParameters.ContainsKey('PassThru')) { $invokeParams.PassThru = $PassThru }

    Invoke-AssessmentScript -ScriptPath $scriptPath -Parameters $invokeParams
}

function Invoke-M365AssessmentComparison {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaselineJsonPath,
        [Parameter(Mandatory = $true)]
        [string]$CurrentJsonPath,
        [Parameter(Mandatory = $false)]
        [string]$OutputFolder
    )

    $scriptPath = Resolve-AssessmentScriptPath -Name 'Compare-M365TenantAssessmentSnapshots.ps1'
    $invokeParams = @{
        BaselineJsonPath = $BaselineJsonPath
        CurrentJsonPath  = $CurrentJsonPath
    }
    if ($PSBoundParameters.ContainsKey('OutputFolder')) { $invokeParams.OutputFolder = $OutputFolder }

    Invoke-AssessmentScript -ScriptPath $scriptPath -Parameters $invokeParams
}

Export-ModuleMember -Function @(
    'Invoke-M365TenantAssessment',
    'Invoke-M365TenantConnectionPreflight',
    'Invoke-M365TenantDataCollection',
    'Invoke-M365TenantAssessmentExport',
    'Invoke-ADTenantAssessment',
    'Invoke-GraphActivityAssessment',
    'Invoke-M365ImprovementPlan',
    'Invoke-M365AssessmentComparison'
)
