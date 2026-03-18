Set-StrictMode -Version Latest

$script:RepoRoot = (Resolve-Path -Path (Join-Path $PSScriptRoot '..\..\..')).Path
$script:AssessmentScriptRoots = @(
    (Join-Path $script:RepoRoot 'src\scripts\reporting'),
    (Join-Path $script:RepoRoot 'src\scripts\migrated\legacy')
)

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
        $loadedCommonModule.Path -ne $resolvedCommonManifestPath -or
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
        'Combined' = 2
        'Geek' = 3
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
        [switch]$StoreTenantStatsGlobal,
        [Parameter(Mandatory = $false)]
        [string]$TenantStatsVariableName = 'ArrayaTenantStats',
        [Parameter(Mandatory = $false)]
        [switch]$SkipAuth,
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
        Write-Host ("Running merged profile pass for: {0}" -f ($plan.SelectedOutputProfiles -join ', ')) -ForegroundColor Cyan
        Write-Host ("Merged reporting mode: {0}" -f $plan.ReportingMode) -ForegroundColor DarkCyan
    }

    $invokeParams = @{}
    $invokeParams.ExportPath = Resolve-AssessmentExportPathInput -ExportPath $ExportPath
    $invokeParams.OutputProfile = $plan.PrimaryProfile
    $invokeParams.OutputProfileLabel = $plan.ProfileLabel
    $invokeParams.ReportingModeOverride = $plan.ReportingMode
    $invokeParams.GenerateWorkbookOverride = [bool]$plan.GenerateWorkbook
    $invokeParams.GenerateTechnicalHtmlOverride = [bool]$plan.GenerateTechnicalHtml
    $invokeParams.GenerateBestPracticesHtmlOverride = [bool]$plan.GenerateBestPracticesHtml
    $invokeParams.GenerateQuestionnaireOverride = [bool]$plan.GenerateQuestionnaire
    $invokeParams.GenerateJsonOverride = [bool]$plan.GenerateJson
    $invokeParams.GeneratePdfOverride = [bool]$plan.GeneratePdf
    if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
    if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
    if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }
    if ($PSBoundParameters.ContainsKey('StoreTenantStatsGlobal')) { $invokeParams.StoreTenantStatsGlobal = $StoreTenantStatsGlobal }
    if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
    if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
    if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
    if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
    if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
    if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
    if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }

    Invoke-AssessmentScript -ScriptPath $scriptPath -Parameters $invokeParams
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
        [switch]$StoreTenantStatsGlobal,
        [Parameter(Mandatory = $false)]
        [string]$TenantStatsVariableName = 'ArrayaTenantStats',
        [Parameter(Mandatory = $false)]
        [switch]$SkipAuth,
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
        Write-Host ("Running merged profile collection pass for: {0}" -f ($plan.SelectedOutputProfiles -join ', ')) -ForegroundColor Cyan
        Write-Host ("Merged reporting mode: {0}" -f $plan.ReportingMode) -ForegroundColor DarkCyan
    }

    $invokeParams = @{}
    $invokeParams.ExportPath = Resolve-AssessmentExportPathInput -ExportPath $ExportPath
    $invokeParams.OutputProfile = $plan.PrimaryProfile
    $invokeParams.OutputProfileLabel = "Collection-$($plan.ProfileLabel)"
    $invokeParams.ReportingModeOverride = $plan.ReportingMode
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
    if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
    if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
    if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
    if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
    if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }

    Invoke-AssessmentScript -ScriptPath $scriptPath -Parameters $invokeParams
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
        [switch]$SkipJsonReport
    )

    $plan = Resolve-M365OutputProfileExecutionPlan -OutputProfile $OutputProfile
    $scriptPath = Resolve-AssessmentScriptPath -Name 'Get-FullTenantReportDetails.ps1'

    $invokeParams = @{}
    $invokeParams.ExportPath = Resolve-AssessmentExportPathInput -ExportPath $ExportPath
    $invokeParams.OutputProfile = $plan.PrimaryProfile
    $invokeParams.OutputProfileLabel = "Export-$($plan.ProfileLabel)"
    $invokeParams.ReportingModeOverride = $plan.ReportingMode
    $invokeParams.GenerateWorkbookOverride = [bool]$plan.GenerateWorkbook
    $invokeParams.GenerateTechnicalHtmlOverride = [bool]$plan.GenerateTechnicalHtml
    $invokeParams.GenerateBestPracticesHtmlOverride = [bool]$plan.GenerateBestPracticesHtml
    $invokeParams.GenerateQuestionnaireOverride = [bool]$plan.GenerateQuestionnaire
    $invokeParams.GenerateJsonOverride = [bool]$plan.GenerateJson
    $invokeParams.GeneratePdfOverride = [bool]$plan.GeneratePdf
    $invokeParams.ExportOnly = $true
    $invokeParams.TenantStatsJsonPath = $AssessmentJsonPath
    if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
    if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
    if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }

    Invoke-AssessmentScript -ScriptPath $scriptPath -Parameters $invokeParams
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
        [switch]$UseGraphFallback
    )

    $scriptPath = Resolve-AssessmentScriptPath -Name 'New-M365TenantImprovementPlan.ps1'
    $invokeParams = @{ AssessmentJsonPath = $AssessmentJsonPath }
    if ($PSBoundParameters.ContainsKey('OutputFolder')) { $invokeParams.OutputFolder = $OutputFolder }
    if ($PSBoundParameters.ContainsKey('UseGraphFallback')) { $invokeParams.UseGraphFallback = $UseGraphFallback }

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
