Set-StrictMode -Version Latest

$script:RepoRoot = (Resolve-Path -Path (Join-Path $PSScriptRoot '..\..\..')).Path
$script:LegacyScriptRoot = Join-Path $script:RepoRoot 'src\scripts\migrated\legacy'

function Invoke-LegacyScriptCompat {
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
        Import-Module -Name $resolvedCommonManifestPath -Force -ErrorAction Stop
    }

    Import-ArrayaOffice365CustomLocal -RepoRoot $script:RepoRoot -RequiredCommands $RequiredCommands | Out-Null
}

function Get-LegacyScriptPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $scriptPath = Join-Path $script:LegacyScriptRoot $Name
    if (-not (Test-Path -Path $scriptPath)) {
        throw "Legacy script not found: $scriptPath"
    }

    return $scriptPath
}

function Invoke-M365TenantAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExportPath,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
        [string]$OutputProfile = 'SolutionsEngineer',
        [Parameter(Mandatory = $false)]
        [switch]$SkipHtmlReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipPdfReport,
        [Parameter(Mandatory = $false)]
        [switch]$SkipJsonReport,
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret
    )

    if (-not (Get-Command -Name 'Get-ArrayaAssessmentOutputProfilePolicy' -ErrorAction SilentlyContinue)) {
        Import-AssessmentRunnerDependencies
    }
    $null = Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile $OutputProfile

    $scriptPath = Get-LegacyScriptPath -Name 'Get-FullTenantReportDetails.ps1'
    $invokeParams = @{}
    if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
    if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
    if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
    if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
    if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }
    if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
    if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
    if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
    if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }

    Invoke-LegacyScriptCompat -ScriptPath $scriptPath -Parameters $invokeParams
}

function Invoke-ADTenantAssessment {
    [CmdletBinding()]
    param()

    $scriptPath = Get-LegacyScriptPath -Name 'Get-ActiveDirectoryReport.ps1'
    Invoke-LegacyScriptCompat -ScriptPath $scriptPath
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

    $scriptPath = Get-LegacyScriptPath -Name 'New-M365TenantImprovementPlan.ps1'
    $invokeParams = @{ AssessmentJsonPath = $AssessmentJsonPath }
    if ($PSBoundParameters.ContainsKey('OutputFolder')) { $invokeParams.OutputFolder = $OutputFolder }
    if ($PSBoundParameters.ContainsKey('UseGraphFallback')) { $invokeParams.UseGraphFallback = $UseGraphFallback }

    Invoke-LegacyScriptCompat -ScriptPath $scriptPath -Parameters $invokeParams
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

    $scriptPath = Get-LegacyScriptPath -Name 'Compare-M365TenantAssessmentSnapshots.ps1'
    $invokeParams = @{
        BaselineJsonPath = $BaselineJsonPath
        CurrentJsonPath  = $CurrentJsonPath
    }
    if ($PSBoundParameters.ContainsKey('OutputFolder')) { $invokeParams.OutputFolder = $OutputFolder }

    Invoke-LegacyScriptCompat -ScriptPath $scriptPath -Parameters $invokeParams
}

Export-ModuleMember -Function @(
    'Invoke-M365TenantAssessment',
    'Invoke-ADTenantAssessment',
    'Invoke-GraphActivityAssessment',
    'Invoke-M365ImprovementPlan',
    'Invoke-M365AssessmentComparison'
)
