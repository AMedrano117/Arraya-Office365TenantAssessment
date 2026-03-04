Set-StrictMode -Version Latest

$script:RepoRoot = (Resolve-Path -Path (Join-Path $PSScriptRoot '..\..\..')).Path
$script:LegacyScriptRoot = Join-Path $script:RepoRoot 'src\scripts\migrated\legacy'

function Import-AssessmentRunnerDependencies {
    [CmdletBinding()]
    param(
        [string[]]$RequiredCommands = @()
    )

    $commonManifestPath = Join-Path $script:RepoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
    if (-not (Get-Module -Name 'Arraya.M365.Common' -ErrorAction SilentlyContinue)) {
        Import-Module -Name $commonManifestPath -ErrorAction Stop
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
        [ValidateSet('Minimum', 'Combined', 'All', 'Geek')]
        [string]$ReportingMode,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Lean', 'Standard', 'Full')]
        [string]$OutputProfile = 'Lean',
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

    $scriptPath = Get-LegacyScriptPath -Name 'Get-FullTenantReportDetails.ps1'
    $invokeParams = @{}
    if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
    if ($PSBoundParameters.ContainsKey('ReportingMode')) { $invokeParams.ReportingMode = $ReportingMode }
    if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
    if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
    if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
    if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }
    if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
    if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
    if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
    if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }

    & $scriptPath @invokeParams
}

function Invoke-ADTenantAssessment {
    [CmdletBinding()]
    param()

    $scriptPath = Get-LegacyScriptPath -Name 'Get-ActiveDirectoryReport.ps1'
    & $scriptPath
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
        [string]$OutputFolder
    )

    $scriptPath = Get-LegacyScriptPath -Name 'New-M365TenantImprovementPlan.ps1'
    $invokeParams = @{ AssessmentJsonPath = $AssessmentJsonPath }
    if ($PSBoundParameters.ContainsKey('OutputFolder')) { $invokeParams.OutputFolder = $OutputFolder }

    & $scriptPath @invokeParams
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

    & $scriptPath @invokeParams
}

Export-ModuleMember -Function @(
    'Invoke-M365TenantAssessment',
    'Invoke-ADTenantAssessment',
    'Invoke-GraphActivityAssessment',
    'Invoke-M365ImprovementPlan',
    'Invoke-M365AssessmentComparison'
)
