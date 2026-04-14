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
        Import-Module -Name $resolvedCommonManifestPath -Force -DisableNameChecking -ErrorAction Stop
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

function Resolve-AssessmentLatestManifestPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportPath
    )

    $candidateManifestPaths = New-Object System.Collections.Generic.List[string]
    $fullExportPath = [System.IO.Path]::GetFullPath($ExportPath)
    if (Test-Path -Path $fullExportPath -PathType Container) {
        $candidateManifestPaths.Add((Join-Path -Path $fullExportPath -ChildPath 'Support\Run.manifest.json'))
        foreach ($manifest in @(Get-ChildItem -Path (Join-Path -Path $fullExportPath -ChildPath 'Support') -Filter '*-Run.manifest.json' -File -ErrorAction SilentlyContinue)) {
            $candidateManifestPaths.Add($manifest.FullName)
        }
    }
    $fullExportPathParent = Split-Path -Path $fullExportPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($fullExportPathParent)) {
        $candidateManifestPaths.Add((Join-Path -Path $fullExportPathParent -ChildPath 'Support\Run.manifest.json'))
        foreach ($manifest in @(Get-ChildItem -Path (Join-Path -Path $fullExportPathParent -ChildPath 'Support') -Filter '*-Run.manifest.json' -File -ErrorAction SilentlyContinue)) {
            $candidateManifestPaths.Add($manifest.FullName)
        }
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
    if ($ImproveResult.PSObject.Properties.Name -contains 'CustomerAssessmentReportMarkdownPath' -and -not [string]::IsNullOrWhiteSpace([string]$ImproveResult.CustomerAssessmentReportMarkdownPath)) {
        $artifactEntries += [pscustomobject][ordered]@{
            Type = 'Customer Assessment Report Markdown'
            Path = [string]$ImproveResult.CustomerAssessmentReportMarkdownPath
        }
    }
    elseif ($ImproveResult.PSObject.Properties.Name -contains 'CustomerAssessmentReportMarkdownPaths') {
        $markdownIndex = 0
        foreach ($markdownPath in @($ImproveResult.CustomerAssessmentReportMarkdownPaths | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)) {
            $markdownIndex++
            $artifactEntries += [pscustomobject][ordered]@{
                Type = ('Customer Assessment Report Markdown {0}' -f $markdownIndex)
                Path = [string]$markdownPath
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
        throw "Could not find a manifest for the completed assessment run under: $ExportPath"
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
    foreach ($policy in $selectedPolicies) {
        $mergedGenerateWorkbook = $mergedGenerateWorkbook -or [bool]$policy.GenerateWorkbook
        $mergedGenerateTechnicalHtml = $mergedGenerateTechnicalHtml -or [bool]$policy.GenerateTechnicalHtml
        $mergedGenerateBestPracticesHtml = $mergedGenerateBestPracticesHtml -or [bool]$policy.GenerateBestPracticesHtml
        $mergedGenerateQuestionnaire = $mergedGenerateQuestionnaire -or [bool]$policy.GenerateQuestionnaire
        $mergedGenerateJson = $mergedGenerateJson -or [bool]$policy.GenerateJson
        $mergedGeneratePdf = $mergedGeneratePdf -or [bool]$policy.GeneratePdf
    }

    $primaryProfile = [string]$selectedOutputProfiles[0]
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
    }
}

function Invoke-M365TenantWorkflow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Full', 'CollectOnly', 'ExportOnly')]
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
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthMode,
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret
    )

    $plan = Resolve-M365OutputProfileExecutionPlan -OutputProfile $OutputProfile
    $scriptPath = Resolve-AssessmentScriptPath -Name 'Get-FullTenantReportDetails.ps1'

    if ($plan.IsMergedSelection) {
        $modeLabel = switch ($Mode) {
            'CollectOnly' { 'collection pass' }
            'ExportOnly' { 'export pass' }
            default { 'pass' }
        }
        Write-Host ("Running merged profile {0} for: {1}" -f $modeLabel, ($plan.SelectedOutputProfiles -join ', ')) -ForegroundColor Cyan
        Write-Host ("Merged reporting mode: {0}" -f $plan.ReportingMode) -ForegroundColor DarkCyan
    }

    $invokeParams = @{}
    $invokeParams.ExportPath = Resolve-AssessmentExportPathInput -ExportPath $ExportPath
    $invokeParams.OutputProfile = $plan.PrimaryProfile
    $invokeParams.ReportingModeOverride = $plan.ReportingMode

    switch ($Mode) {
        'Full' {
            $invokeParams.OutputProfileLabel = $plan.ProfileLabel
            $invokeParams.GenerateWorkbookOverride = [bool]$plan.GenerateWorkbook
            $invokeParams.GenerateTechnicalHtmlOverride = if ($IncludeLegacyAssessmentArtifacts) { [bool]$plan.GenerateTechnicalHtml } else { $false }
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
        }
        'CollectOnly' {
            $invokeParams.OutputProfileLabel = "Collection-$($plan.ProfileLabel)"
            $invokeParams.GenerateWorkbookOverride = $false
            $invokeParams.GenerateTechnicalHtmlOverride = $false
            $invokeParams.GenerateBestPracticesHtmlOverride = $false
            $invokeParams.GenerateQuestionnaireOverride = $false
            $invokeParams.GenerateJsonOverride = $true
            $invokeParams.GeneratePdfOverride = $false
            $invokeParams.DataCollectionOnly = $true
            if ($PSBoundParameters.ContainsKey('StoreTenantStatsGlobal')) { $invokeParams.StoreTenantStatsGlobal = $StoreTenantStatsGlobal }
            if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
            if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
            if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight')) { $invokeParams.SkipPermissionPreflight = $SkipPermissionPreflight }
            if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
            if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
            if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
            if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
            if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }
        }
        'ExportOnly' {
            if ([string]::IsNullOrWhiteSpace($AssessmentJsonPath)) {
                throw 'AssessmentJsonPath is required when Mode is ExportOnly.'
            }
            $invokeParams.OutputProfileLabel = "Export-$($plan.ProfileLabel)"
            $invokeParams.GenerateWorkbookOverride = [bool]$plan.GenerateWorkbook
            $invokeParams.GenerateTechnicalHtmlOverride = if ($IncludeLegacyAssessmentArtifacts) { [bool]$plan.GenerateTechnicalHtml } else { $false }
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
        [string]$ClientSecret
    )

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

    if ($PSBoundParameters.ContainsKey('IncludeLegacyAssessmentArtifacts')) { $invokeParams.IncludeLegacyAssessmentArtifacts = $IncludeLegacyAssessmentArtifacts }
    Invoke-M365TenantWorkflow -Mode Full @invokeParams

    if (-not $SkipImprove) {
        $resolvedExportPath = Resolve-AssessmentExportPathInput -ExportPath $ExportPath
        Invoke-M365ImproveForAssessmentRun -ExportPath $resolvedExportPath -OutputFolder $ImproveOutputFolder -UseGraphFallback:$UseGraphFallback -IncludeLegacyArtifacts:$IncludeLegacyArtifacts | Out-Null
    }
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
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthMode,
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret
    )

    $invokeParams = @{}
    if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
    if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
    if ($PSBoundParameters.ContainsKey('StoreTenantStatsGlobal')) { $invokeParams.StoreTenantStatsGlobal = $StoreTenantStatsGlobal }
    if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
    if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
    if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight')) { $invokeParams.SkipPermissionPreflight = $SkipPermissionPreflight }
    if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
    if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
    if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
    if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
    if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }

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

    $invokeParams = @{ AssessmentJsonPath = $AssessmentJsonPath }
    if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
    if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
    if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
    if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
    if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }
    if ($PSBoundParameters.ContainsKey('IncludeLegacyAssessmentArtifacts')) { $invokeParams.IncludeLegacyAssessmentArtifacts = $IncludeLegacyAssessmentArtifacts }
    Invoke-M365TenantWorkflow -Mode ExportOnly @invokeParams

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
    'Invoke-M365TenantDataCollection',
    'Invoke-M365TenantAssessmentExport',
    'Invoke-ADTenantAssessment',
    'Invoke-GraphActivityAssessment',
    'Invoke-M365ImprovementPlan',
    'Invoke-M365AssessmentComparison'
)
