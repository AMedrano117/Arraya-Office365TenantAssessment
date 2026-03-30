<#
.SYNOPSIS
This script is designed to automate the process of gathering, processing, and exporting data related to a Microsoft 365 tenant environment.

.DESCRIPTION
The script follows a structured approach comprising the following main sections:

1. Initialization:
   - Establishes connections to all necessary Office 365 services.
   - Checks and installs the ImportExcel module if needed.
   - Retrieves the export path, initializes various variables and arrays.
   - This will prioritize running the Microsoft Graph PowerShell SDK, especially if running Windows 7 PowerShell.
   - If a lesser version of PowerShell or Microsoft Graph PowerShell SDK is not installed, the other modules of SharePointOnline, MicrosoftTeams will be utilized along with ExchangeOnlineManagement

2. Data Gathering:
   - Collects data from Exchange Online, SharePoint/OneDrive, Entra ID, and other tenant objects.
   - Includes hybrid/federation signals, conditional access, MFA registration, admin role assignments, and AD Connect sync status (if available).
   - The data collection functions are parameterized to control the depth or scope of the data gathered based on the `detailLevel` argument.

3. Data Consolidation (Optional):
   - Optionally consolidates discovery report data into one file based on the reporting mode.

4. Export Preprocessing:
   - Processes the collected data to remove certain tables based on the reporting mode.
   - Creates a new hashtable to hold the data for export, excluding specified tables.

5. Data Export:
   - Exports the processed data to Excel and a JSON snapshot (for reuse/report-only runs).

6. Error Export:
   - Exports any captured errors to a separate file, and displays a summary of error details to the user.

7. Final Reporting:
   - Calculates and displays the total execution time.
   - Generates a final summary table of object counts, which is output to the console.

8. Logging and Error Handling:
   - Utilizes extensive logging and error handling to ensure that issues are captured, logged, and reported in a user-friendly manner.
9. HTML Reporting:
   - Generates a tenant HTML report with findings, KPIs, and section summaries.

.FUNCTIONS

- Get-AllRecipientDetails, Get-AllExchangeMailboxDetails, Get-ExchangeGroupDetails, Get-MailFlowRulesandConnectors, Get-AllPublicFolderDetails:
   Functions for gathering data from Exchange Online.

- Get-AllUnifiedGroups, Get-SPOAndOneDriveDetails, Get-TeamsDetails:
   Functions for gathering data from SharePoint and Teams.

- Get-AllLicenseSKUs, Get-allUserDetails, Get-AllOffice365Domains, Get-AllOffice365Admins:
   Functions for gathering tenant-wide data and license details.
   These Functions support Microsoft Graph and MSOnline connections

- Get-ExchangeHybridConfiguration, Get-FederationAndCrossTenantConfiguration, Get-AdConnectSyncDetails, Get-MfaRegistrationDetails:
   Functions for hybrid/federation detection, AD Connect sync status, and MFA registration reporting.

- New-TenantHtmlReport:
   Generates a consolidated HTML report with findings and KPIs.

- Combine-AllMailboxStats:
   Optional function to consolidate discovery report data into a single file.

.REQUIREMENTS
- Microsoft Graph PowerShell SDK
- ImportExcel
- Import Custom Module - Office365CustomModule


.NOTES
The script utilizes modular functions to perform specific tasks and employs logging and error handling to provide a robust solution for data collection and reporting.
The user-friendly prompts and color-coded console output enhance the user experience when running the script.

.AUTHOR
Aaron Medrano

.DATE
Created Date: 2021-04-19
Last Modified Date: 2024-06-06

.EXAMPLE
# How to run the script interactively
1. Save the script to a file named Get-FullTenantReportDetails.ps1.
2. Open PowerShell as an Administrator.
3. Navigate to the directory containing the script:
   cd "path\to\your\script\directory"
4. Run the script:
   .\Get-FullTenantReportDetails.ps1
5. Follow the prompts to provide necessary input such as OutputProfile, ExportPath, Authentication, and others as prompted.
6. Check the specified export path for the generated reports and error logs.
7. Review the console output for any errors or summary information provided by the script.

# The script will connect to all Office 365 services, gather data, and export it to the specified path.
# Errors will be logged and a summary of the data gathered will be output to the console.

#>
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
    [string]$ClientSecret,
    [Parameter(Mandatory = $false)]
    [string]$OutputProfileLabel,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Minimum', 'Combined', 'Geek')]
    [string]$ReportingModeOverride,
    [Parameter(Mandatory = $false)]
    [bool]$GenerateWorkbookOverride,
    [Parameter(Mandatory = $false)]
    [bool]$GenerateTechnicalHtmlOverride,
    [Parameter(Mandatory = $false)]
    [bool]$GenerateBestPracticesHtmlOverride,
    [Parameter(Mandatory = $false)]
    [bool]$GenerateQuestionnaireOverride,
    [Parameter(Mandatory = $false)]
    [bool]$GenerateJsonOverride,
    [Parameter(Mandatory = $false)]
    [bool]$GeneratePdfOverride,
    [Parameter(Mandatory = $false)]
    [switch]$DataCollectionOnly,
    [Parameter(Mandatory = $false)]
    [switch]$ExportOnly,
    [Parameter(Mandatory = $false)]
    [string]$TenantStatsJsonPath
)

# Strict-mode safety: ensure legacy Graph globals exist even when SDK auth is used.
if (-not (Get-Variable -Name GraphHeaders -Scope Global -ErrorAction SilentlyContinue)) {
    $global:GraphHeaders = $null
}
if (-not (Get-Variable -Name GraphToken -Scope Global -ErrorAction SilentlyContinue)) {
    $global:GraphToken = $null
}

# Suppress Microsoft Graph SDK progress records so the assessment's own
# progress bars remain readable and transient Graph SDK bars do not stick.
$global:PSDefaultParameterValues['Get-Mg*:ProgressAction'] = 'SilentlyContinue'
$global:PSDefaultParameterValues['Find-Mg*:ProgressAction'] = 'SilentlyContinue'
$global:PSDefaultParameterValues['Invoke-MgGraphRequest:ProgressAction'] = 'SilentlyContinue'

$effectiveSkipWorkbook = $false
$effectiveSkipBestPracticesHtml = $false
$effectiveSkipQuestionnaire = $false
$effectiveSkipHtmlReport = $false
$effectiveSkipPdfReport = $false
$effectiveSkipJsonReport = $false
$effectiveGenerateWorkbook = $false
$effectiveGenerateTechnicalHtml = $false
$effectiveGenerateBestPracticesHtml = $false
$effectiveGenerateQuestionnaire = $false
$effectiveGenerateJson = $false
$effectiveGeneratePdf = $false
$reportingMode = 'minimum'
$script:EffectiveOutputProfileLabel = $OutputProfile
$isMergedOutputProfileSelection = $false
$runExportOnly = $ExportOnly.IsPresent
$runCollectionOnly = $DataCollectionOnly.IsPresent
$script:LoadedTenantSnapshot = $null
$script:CurrentGraphMode = 'UNKNOWN'

if ($runCollectionOnly -and $runExportOnly) {
    throw "Data collection only mode and export only mode cannot be used together."
}

if ($runExportOnly -and [string]::IsNullOrWhiteSpace($TenantStatsJsonPath)) {
    throw "Export only mode requires -TenantStatsJsonPath."
}

if ($runCollectionOnly -and $SkipJsonReport.IsPresent) {
    throw "Data collection only mode requires JSON output. Remove -SkipJsonReport."
}

$officeModuleLoaderPath = Join-Path -Path $PSScriptRoot -ChildPath 'Import-Office365CustomLocal.ps1'
if (-not (Test-Path -Path $officeModuleLoaderPath)) {
    throw "Required loader script not found: $officeModuleLoaderPath"
}
. $officeModuleLoaderPath
Import-Office365CustomLocal -RepoRoot $PSScriptRoot -RequiredCommands @(
    'Connect-Office365',
    'Write-ProgressHelper',
    'Office365Custom\Write-Log',
    'Office365Custom\Capture-ErrorHelper',
    'Office365Custom\Get-ExportPath',
    'Office365Custom\Get-GraphData'
) | Out-Null

$commonModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'))
if (-not (Test-Path -Path $commonModuleManifestPath)) {
    throw "Required common module manifest not found: $commonModuleManifestPath"
}
$resolvedCommonManifestPath = (Resolve-Path -Path $commonModuleManifestPath).Path
$loadedCommonModule = Get-Module -Name 'Arraya.M365.Common' -ErrorAction SilentlyContinue | Select-Object -First 1
$requiredCommonCommands = @(
    'Get-ArrayaAssessmentOutputRoot',
    'Get-ArrayaAssessmentOutputProfilePolicy',
    'Invoke-ArrayaCollectionStepSafe',
    'Export-ArrayaErrorReports',
    'Convert-ArrayaObjectToArray',
    'Get-ArrayaObjectValue',
    'Import-ArrayaTenantSnapshotContext',
    'Convert-ArrayaToNumber',
    'New-ArrayaAssessmentContext',
    'New-ArrayaTenantSnapshot',
    'Update-ArrayaTenantSnapshot',
    'Test-ArrayaTenantSnapshot',
    'Convert-ArrayaLegacyTenantStatsToSnapshot',
    'Convert-ArrayaSnapshotToLegacyTenantStatsHash',
    'Import-ArrayaTenantSnapshot',
    'Export-ArrayaTenantSnapshot',
    'Write-ArrayaAssessmentArtifactManifest'
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

$reportingModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\modules\Arraya.M365.Reporting\Arraya.M365.Reporting.psd1'))
if (-not (Test-Path -Path $reportingModuleManifestPath)) {
    throw "Required reporting module manifest not found: $reportingModuleManifestPath"
}
$resolvedReportingManifestPath = (Resolve-Path -Path $reportingModuleManifestPath).Path
$loadedReportingModule = Get-Module -Name 'Arraya.M365.Reporting' -ErrorAction SilentlyContinue | Select-Object -First 1
    $requiredReportingCommands = @(
        'Get-ArrayaAssessmentWorksheetName',
        'Get-ArrayaAssessmentRecommendationText',
        'Get-ArrayaEmployeeExperienceInsightsAnalysis',
        'Get-GraphUserStats'
    )
$missingReportingCommands = @(
    $requiredReportingCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
)
if (
    -not $loadedReportingModule -or
    $loadedReportingModule.Path -ne $resolvedReportingManifestPath -or
    $missingReportingCommands.Count -gt 0
) {
    Import-Module -Name $resolvedReportingManifestPath -Force -DisableNameChecking -ErrorAction Stop
}

function Import-AssessmentCollectorModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LegacyScriptRoot
    )

    $graphModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $LegacyScriptRoot -ChildPath '..\..\..\modules\Arraya.M365.Graph\Arraya.M365.Graph.psd1'))
    if (-not (Test-Path -Path $graphModuleManifestPath)) {
        throw "Required Graph collector module manifest not found: $graphModuleManifestPath"
    }

    $resolvedGraphManifestPath = (Resolve-Path -Path $graphModuleManifestPath).Path
    $loadedGraphModule = Get-Module -Name 'Arraya.M365.Graph' -ErrorAction SilentlyContinue | Select-Object -First 1
    $requiredGraphCommands = @('Get-EntraIDGroups')
    $missingGraphCommands = @(
        $requiredGraphCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
    )
    if (
        -not $loadedGraphModule -or
        $loadedGraphModule.Path -ne $resolvedGraphManifestPath -or
        $missingGraphCommands.Count -gt 0
    ) {
        Import-Module -Name $resolvedGraphManifestPath -Force -ErrorAction Stop
    }

    $exchangeModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $LegacyScriptRoot -ChildPath '..\..\..\modules\Arraya.M365.Exchange\Arraya.M365.Exchange.psd1'))
    if (-not (Test-Path -Path $exchangeModuleManifestPath)) {
        throw "Required Exchange collector module manifest not found: $exchangeModuleManifestPath"
    }

    $resolvedExchangeManifestPath = (Resolve-Path -Path $exchangeModuleManifestPath).Path
    $loadedExchangeModule = Get-Module -Name 'Arraya.M365.Exchange' -ErrorAction SilentlyContinue | Select-Object -First 1
    $requiredExchangeCommands = @(
        'Get-AllRecipientDetails',
        'Get-AllExchangeMailboxDetails',
        'Get-ExchangeGroupDetails',
        'Get-AllPublicFolderDetails',
        'Get-ExchangeHybridConfiguration',
        'Get-MailFlowRulesandConnectors',
        'Get-ThirdPartySpamFilteringConfig',
        'Get-SMTPRelayConfiguration'
    )
    $missingExchangeCommands = @(
        $requiredExchangeCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
    )
    if (
        -not $loadedExchangeModule -or
        $loadedExchangeModule.Path -ne $resolvedExchangeManifestPath -or
        $missingExchangeCommands.Count -gt 0
    ) {
        Import-Module -Name $resolvedExchangeManifestPath -Force -ErrorAction Stop
    }
}

function Test-CollectorCommandAvailable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Get-Command -Name $Name -ErrorAction SilentlyContinue) {
        return $true
    }

    if (Get-Command -Name ("Office365Custom\{0}" -f $Name) -ErrorAction SilentlyContinue) {
        return $true
    }

    return $false
}

Import-AssessmentCollectorModules -LegacyScriptRoot $PSScriptRoot
$requiredCollectorCommands = @(
    'Get-AllRecipientDetails',
    'Get-AllExchangeMailboxDetails',
    'Get-ExchangeGroupDetails',
    'Get-AllPublicFolderDetails',
    'Get-ExchangeHybridConfiguration',
    'Get-MailFlowRulesandConnectors',
    'Get-ThirdPartySpamFilteringConfig',
    'Get-SMTPRelayConfiguration',
    'Get-EntraIDGroups',
    'Get-GraphUserStats'
)
$missingCollectorCommands = @(
    $requiredCollectorCommands | Where-Object { -not (Test-CollectorCommandAvailable -Name $_) }
)
if ($missingCollectorCommands.Count -gt 0) {
    throw "Required O365 assessment collector command(s) are unavailable: $($missingCollectorCommands -join ', '). Validate module wiring for Arraya.M365.Exchange/Graph/Reporting or Office365Custom."
}

$profilePolicy = Get-ArrayaAssessmentOutputProfilePolicy -OutputProfile $OutputProfile
$effectiveOutputProfileLabel = if (-not [string]::IsNullOrWhiteSpace($OutputProfileLabel)) { $OutputProfileLabel } else { $OutputProfile }
$script:EffectiveOutputProfileLabel = $effectiveOutputProfileLabel
$isMergedOutputProfileSelection = $effectiveOutputProfileLabel -like 'Merged(*'

$resolvedReportingMode = if ($PSBoundParameters.ContainsKey('ReportingModeOverride') -and -not [string]::IsNullOrWhiteSpace($ReportingModeOverride)) {
    $ReportingModeOverride
}
else {
    [string]$profilePolicy.ReportingMode
}
$reportingMode = $resolvedReportingMode.ToLowerInvariant()

$effectiveGenerateWorkbook = if ($PSBoundParameters.ContainsKey('GenerateWorkbookOverride')) {
    [bool]$GenerateWorkbookOverride
}
else {
    [bool]$profilePolicy.GenerateWorkbook
}
$effectiveGenerateTechnicalHtml = if ($PSBoundParameters.ContainsKey('GenerateTechnicalHtmlOverride')) {
    [bool]$GenerateTechnicalHtmlOverride
}
else {
    [bool]$profilePolicy.GenerateTechnicalHtml
}
$effectiveGenerateBestPracticesHtml = if ($PSBoundParameters.ContainsKey('GenerateBestPracticesHtmlOverride')) {
    [bool]$GenerateBestPracticesHtmlOverride
}
else {
    [bool]$profilePolicy.GenerateBestPracticesHtml
}
$effectiveGenerateQuestionnaire = if ($PSBoundParameters.ContainsKey('GenerateQuestionnaireOverride')) {
    [bool]$GenerateQuestionnaireOverride
}
else {
    [bool]$profilePolicy.GenerateQuestionnaire
}
$effectiveGenerateJson = if ($PSBoundParameters.ContainsKey('GenerateJsonOverride')) {
    [bool]$GenerateJsonOverride
}
else {
    [bool]$profilePolicy.GenerateJson
}
$effectiveGeneratePdf = if ($PSBoundParameters.ContainsKey('GeneratePdfOverride')) {
    [bool]$GeneratePdfOverride
}
else {
    [bool]$profilePolicy.GeneratePdf
}

if ($runCollectionOnly) {
    $effectiveGenerateWorkbook = $false
    $effectiveGenerateTechnicalHtml = $false
    $effectiveGenerateBestPracticesHtml = $false
    $effectiveGenerateQuestionnaire = $false
    $effectiveGenerateJson = $true
    $effectiveGeneratePdf = $false
}

$effectiveSkipWorkbook = (-not $effectiveGenerateWorkbook)
$effectiveSkipBestPracticesHtml = (-not $effectiveGenerateBestPracticesHtml)
$effectiveSkipQuestionnaire = (-not $effectiveGenerateQuestionnaire)
$effectiveSkipHtmlReport = (-not $effectiveGenerateTechnicalHtml) -or $SkipHtmlReport.IsPresent
$effectiveSkipPdfReport = (-not $effectiveGeneratePdf) -or $SkipPdfReport.IsPresent
$effectiveSkipJsonReport = (-not $effectiveGenerateJson) -or $SkipJsonReport.IsPresent

$tenantHtmlReportPath = Join-Path -Path $PSScriptRoot -ChildPath 'New-TenantHtmlReport.ps1'
$tenantHtmlHelperFunctionsPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1'))
$tenantQuestionnairePath = Join-Path -Path $PSScriptRoot -ChildPath 'Export-TenantToTenantQuestionnaireMarkdown.ps1'
$tenantExportPipelinePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\reporting\Invoke-M365TenantAssessmentExportPipeline.ps1'))

########################################################
# Functions
########################################################
# ----------------------------------
# Default Script Helper Functions
# ----------------------------------

function Write-ConsoleSection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Step,
        [Parameter(Mandatory)]
        [string]$Title
    )

    Write-Host ""
    Write-Host "[$Step] $Title" -ForegroundColor Cyan
    Write-Host ('-' * 72) -ForegroundColor DarkCyan
}

function Invoke-QuietRestMethod {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters
    )

    $savedProgressPreference = $ProgressPreference
    try {
        $ProgressPreference = 'SilentlyContinue'
        return Invoke-RestMethod @Parameters
    }
    finally {
        $ProgressPreference = $savedProgressPreference
    }
}

function Invoke-QuietWebRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters
    )

    $savedProgressPreference = $ProgressPreference
    try {
        $ProgressPreference = 'SilentlyContinue'
        return Invoke-WebRequest @Parameters
    }
    finally {
        $ProgressPreference = $savedProgressPreference
    }
}

function Invoke-QuietCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $savedProgressPreference = $ProgressPreference
    try {
        $ProgressPreference = 'SilentlyContinue'
        return (& $ScriptBlock)
    }
    finally {
        $ProgressPreference = $savedProgressPreference
    }
}

function Convert-DataSizeToBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return [int64]0
    }

    if ($Value -is [int64] -or $Value -is [int32] -or $Value -is [double] -or $Value -is [decimal]) {
        try { return [int64]$Value } catch { return [int64]0 }
    }

    foreach ($propertyName in @('Bytes', 'ByteCount', 'Value')) {
        if ($Value.PSObject -and $Value.PSObject.Properties[$propertyName]) {
            $nestedValue = $Value.PSObject.Properties[$propertyName].Value
            if ($nestedValue -is [int64] -or $nestedValue -is [int32] -or $nestedValue -is [double] -or $nestedValue -is [decimal]) {
                try { return [int64]$nestedValue } catch {}
            }
        }
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [int64]0
    }

    if ($text -match '\((?<bytes>[0-9,]+)\s+bytes\)') {
        try { return [int64](($Matches['bytes'] -replace ',', '')) } catch {}
    }

    if ($text -match '^\s*(?<number>[0-9]+(?:\.[0-9]+)?)\s*(?<unit>KB|MB|GB|TB|PB)\b') {
        $number = [double]$Matches['number']
        $multiplier = switch ($Matches['unit'].ToUpperInvariant()) {
            'KB' { 1KB }
            'MB' { 1MB }
            'GB' { 1GB }
            'TB' { 1TB }
            'PB' { 1PB }
            default { 1 }
        }
        try { return [int64]($number * $multiplier) } catch { return [int64]0 }
    }

    return [int64]0
}

function Get-ArrayaCollectionDepthPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('Minimum', 'Combined', 'All', 'Geek')]
        [string]$ReportingMode = 'Minimum'
    )

    $mode = $ReportingMode.ToLowerInvariant()
    $isMinimum = $mode -eq 'minimum'
    $isCombined = $mode -eq 'combined'
    $isAll = $mode -eq 'all'
    $isGeek = $mode -eq 'geek'

    return [PSCustomObject]@{
        ReportingMode                   = $ReportingMode
        CollectEntraGroupDeepDetails    = (-not $isMinimum)
        CollectEntraGroupMemberCounts   = (-not $isMinimum)
        CollectEntraGroupOwnerCounts    = (-not $isMinimum)
        CollectEntraGroupLicenseChecks  = (-not $isMinimum)
        CollectSsoApplicationDetails    = (-not $isMinimum)
        CollectAuthenticationMethodRows = $true
        CollectSecureScoreMappings      = $true
        CollectFullMailboxNormalization = $true
        CollectUnifiedGroupMailboxStats = (-not $isMinimum)
        CollectFullSharePointDetail     = ($isAll -or $isGeek)
        CollectExtendedGraphEnrichment  = ($isAll -or $isGeek)
        IsMinimum                       = $isMinimum
        IsCombined                      = $isCombined
        IsAll                           = $isAll
        IsGeek                          = $isGeek
    }
}

function Normalize-GraphReportFieldName {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }

    return (($Name -replace '[^A-Za-z0-9]', '').ToLowerInvariant())
}

function Get-GraphReportFieldValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Row,
        [Parameter(Mandatory = $true)]
        [string[]]$FieldNames
    )

    if (-not $Row) {
        return $null
    }

    return Get-ArrayaObjectValue -Object $Row -Names $FieldNames
}

function Convert-GraphReportValueToInt64 {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    return [int64](Convert-ArrayaToNumber -Value $Value -AsInt64)
}

function Get-ArrayaGraphResource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,
        [Parameter(Mandatory = $false)]
        [int]$PageSize = 999,
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Fetching data from Microsoft Graph',
        [Parameter(Mandatory = $false)]
        [switch]$PreferRest,
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 5
    )

    return Office365Custom\Get-GraphData -Uri $Uri -PageSize $PageSize -Activity $Activity -UseRestMethod:$PreferRest -MaxRetries $MaxRetries
}

function Export-ArrayaGraphReportCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Downloading Microsoft Graph report CSV',
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers
    )

    $tempCsvPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("arraya-graph-report-" + [guid]::NewGuid().ToString('N') + '.csv')
    $resolvedHeaders = @{}
    if ($Headers) {
        $resolvedHeaders = $Headers.Clone()
    }
    elseif ($global:GraphHeaders) {
        $resolvedHeaders = $global:GraphHeaders.Clone()
    }
    elseif ($global:GraphToken) {
        $resolvedHeaders = @{
            'Content-Type'     = 'application/json'
            'Authorization'    = "Bearer $global:GraphToken"
            'ConsistencyLevel' = 'eventual'
        }
    }

    $useSdk = $false
    if (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            $useSdk = $true
        }
    }

    try {
        if ($useSdk) {
            try {
                Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputFilePath $tempCsvPath -ProgressAction SilentlyContinue -ErrorAction Stop | Out-Null
            }
            catch {
                if ($resolvedHeaders.Count -eq 0) {
                    throw
                }
                $useSdk = $false
            }
        }

        if (-not $useSdk) {
            if ($resolvedHeaders.Count -eq 0) {
                throw "No Graph authentication context is available for '$Activity'."
            }

            $response = Invoke-QuietWebRequest -Parameters @{
                Uri                = $Uri
                Headers            = $resolvedHeaders
                Method             = 'GET'
                MaximumRedirection = 5
                ErrorAction        = 'Stop'
            }
            if ($null -eq $response -or [string]::IsNullOrWhiteSpace($response.Content)) {
                throw "No CSV content returned for '$Activity'."
            }
            Set-Content -Path $tempCsvPath -Value $response.Content -Encoding UTF8 -Force
        }

        if (-not (Test-Path -Path $tempCsvPath)) {
            throw "CSV report file was not created for '$Activity'."
        }

        return @(Import-Csv -Path $tempCsvPath -ErrorAction Stop)
    }
    finally {
        if (Test-Path -Path $tempCsvPath) {
            Remove-Item -Path $tempCsvPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-ArrayaGraphAdminReportSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,
        [Parameter(Mandatory = $false)]
        [switch]$PreferRest
    )

    try {
        $response = Get-ArrayaGraphResource -Uri 'https://graph.microsoft.com/v1.0/admin/reportSettings' -Activity 'Admin report settings' -PreferRest:$PreferRest -Headers $Headers
    }
    catch {
        return [PSCustomObject]@{
            Available                  = $false
            DisplayConcealedNames      = $null
            IdentifiableNamesInReports = $null
            Source                     = 'Graph /admin/reportSettings'
            ErrorMessage               = $_.Exception.Message
            RetrievedAt                = (Get-Date).ToString('o')
        }
    }

    $displayConcealedNames = $null
    if ($response -is [System.Collections.IDictionary] -and $response.Contains('displayConcealedNames')) {
        try { $displayConcealedNames = [bool]$response['displayConcealedNames'] } catch { $displayConcealedNames = $null }
    }
    elseif ($response.PSObject -and $response.PSObject.Properties['displayConcealedNames']) {
        try { $displayConcealedNames = [bool]$response.displayConcealedNames } catch { $displayConcealedNames = $null }
    }

    return [PSCustomObject]@{
        Available                  = ($null -ne $response)
        DisplayConcealedNames      = $displayConcealedNames
        IdentifiableNamesInReports = if ($displayConcealedNames -is [bool]) { -not $displayConcealedNames } else { $null }
        Source                     = 'Graph /admin/reportSettings'
        ErrorMessage               = $null
        RetrievedAt                = (Get-Date).ToString('o')
    }
}

function Write-ConsoleArtifactSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Artifacts,
        [Parameter(Mandatory)]
        [string]$DurationText,
        [int]$CapturedErrorCount = 0
    )

    Write-Host ""
    Write-Host "Assessment complete in $DurationText" -ForegroundColor Green

    $primaryArtifactLabels = @(
        'Customer Remediation HTML',
        'Engineer Action Pack'
    )
    $primaryArtifacts = @()
    $supportFolders = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $debugFolders = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $fallbackArtifacts = New-Object System.Collections.Generic.List[object]

    foreach ($entry in $Artifacts.GetEnumerator()) {
        $artifactPath = [string]$entry.Value
        if ([string]::IsNullOrWhiteSpace($artifactPath)) {
            continue
        }

        $parentDirectory = Split-Path -Path $artifactPath -Parent
        $leafDirectory = Split-Path -Path $parentDirectory -Leaf
        $artifactRecord = [PSCustomObject]@{
            Label = [string]$entry.Key
            Path  = $artifactPath
        }

        if ($primaryArtifactLabels -contains [string]$entry.Key) {
            $primaryArtifacts += $artifactRecord
            continue
        }

        if ([string]::Equals($leafDirectory, 'Support', [System.StringComparison]::OrdinalIgnoreCase)) {
            [void]$supportFolders.Add($parentDirectory)
            continue
        }

        if ([string]::Equals($leafDirectory, 'Debugging', [System.StringComparison]::OrdinalIgnoreCase)) {
            [void]$debugFolders.Add($parentDirectory)
            continue
        }

        $fallbackArtifacts.Add($artifactRecord) | Out-Null
    }

    if ($primaryArtifacts.Count -gt 0) {
        foreach ($artifact in $primaryArtifacts) {
            $foreground = if ($artifact.Label -eq 'Customer Remediation HTML') { 'Green' } else { 'Cyan' }
            Write-Host ("  {0}: {1}" -f $artifact.Label, $artifact.Path) -ForegroundColor $foreground
        }

        foreach ($supportFolder in ($supportFolders | Sort-Object)) {
            Write-Host ("  Support artifacts: {0}" -f $supportFolder) -ForegroundColor DarkGray
        }

        foreach ($debugFolder in ($debugFolders | Sort-Object)) {
            Write-Host ("  Debugging logs : {0}" -f $debugFolder) -ForegroundColor DarkGray
        }
    }
    else {
        foreach ($artifact in $fallbackArtifacts.ToArray()) {
            Write-Host ("  {0}: {1}" -f $artifact.Label, $artifact.Path) -ForegroundColor Gray
        }
    }

    if ($CapturedErrorCount -gt 0) {
        Write-Host ("  Captured errors: {0} (see log/error reports)" -f $CapturedErrorCount) -ForegroundColor Yellow
    }
}

function Get-ExportPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName,
        [Parameter(Mandatory = $false)]
        [string]$DefaultExtension = '.xlsx',
        [Parameter(Mandatory = $false)]
        [string]$UserInputPath
    )

    if ([string]::IsNullOrWhiteSpace($UserInputPath)) {
        return Office365Custom\Get-ExportPath -FileName $FileName -DefaultExtension $DefaultExtension
    }

    $requestedPath = ($UserInputPath -replace '"', '').Trim()
    if ([string]::IsNullOrWhiteSpace($requestedPath)) {
        return Office365Custom\Get-ExportPath -FileName $FileName -DefaultExtension $DefaultExtension
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $defaultFileName = "{0}_{1}{2}" -f $FileName, $timestamp, $DefaultExtension

    try {
        if (Test-Path -Path $requestedPath -PathType Container) {
            $resolvedDirectory = (Resolve-Path -Path $requestedPath).Path
            return Join-Path -Path $resolvedDirectory -ChildPath $defaultFileName
        }

        $leafName = [System.IO.Path]::GetFileName($requestedPath)
        $parentDirectory = [System.IO.Path]::GetDirectoryName($requestedPath)
        $leafExtension = [System.IO.Path]::GetExtension($leafName)
        $hasFileExtension = -not [string]::IsNullOrWhiteSpace($leafExtension)
        $treatAsDirectory = -not $hasFileExtension

        if ([string]::IsNullOrWhiteSpace($parentDirectory)) {
            if ($hasFileExtension) {
                $parentDirectory = (Resolve-Path -Path '.').Path
            }
            else {
                $parentDirectory = $requestedPath
                $leafName = $null
            }
        }
        elseif ($treatAsDirectory) {
            $parentDirectory = $requestedPath
            $leafName = $null
        }

        if (-not (Test-Path -Path $parentDirectory)) {
            New-Item -Path $parentDirectory -ItemType Directory -Force | Out-Null
        }

        $resolvedDirectory = (Resolve-Path -Path $parentDirectory).Path
        if ($hasFileExtension -and -not [string]::IsNullOrWhiteSpace($leafName)) {
            return Join-Path -Path $resolvedDirectory -ChildPath $leafName
        }

        return Join-Path -Path $resolvedDirectory -ChildPath $defaultFileName
    }
    catch {
        throw "Failed to resolve export path from '$UserInputPath': $($_.Exception.Message)"
    }
}

function Convert-ToAssessmentPathComponent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Value,
        [Parameter(Mandatory = $false)]
        [string]$Fallback = 'Value'
    )

    $candidate = if ([string]::IsNullOrWhiteSpace($Value)) { $Fallback } else { $Value }
    $invalidCharacterPattern = '[{0}]' -f [regex]::Escape((-join [System.IO.Path]::GetInvalidFileNameChars()))
    $sanitizedValue = [regex]::Replace($candidate, $invalidCharacterPattern, ' ')
    $sanitizedValue = ($sanitizedValue -replace '\s+', ' ').Trim()

    if ([string]::IsNullOrWhiteSpace($sanitizedValue)) {
        return $Fallback
    }

    return $sanitizedValue
}

function Resolve-AssessmentExportTargetPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$RequestedExportPath,
        [Parameter(Mandatory = $true)]
        [string]$DefaultOutputRoot,
        [Parameter(Mandatory = $true)]
        [string]$TenantDisplayName,
        [Parameter(Mandatory = $true)]
        [string]$OutputProfileFolderLabel
    )

    $requestedPath = if ([string]::IsNullOrWhiteSpace($RequestedExportPath)) {
        $DefaultOutputRoot
    }
    else {
        ($RequestedExportPath -replace '"', '').Trim()
    }

    if ([string]::IsNullOrWhiteSpace($requestedPath)) {
        $requestedPath = $DefaultOutputRoot
    }

    $requestedLeaf = [System.IO.Path]::GetFileName($requestedPath)
    $requestedLeafExtension = [System.IO.Path]::GetExtension($requestedLeaf)
    if (-not [string]::IsNullOrWhiteSpace($requestedLeafExtension)) {
        return $requestedPath
    }

    $resolvedBaseRoot = if ([System.IO.Path]::IsPathRooted($requestedPath)) {
        [System.IO.Path]::GetFullPath($requestedPath)
    }
    else {
        [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location).Path -ChildPath $requestedPath))
    }

    if (-not (Test-Path -Path $resolvedBaseRoot)) {
        New-Item -Path $resolvedBaseRoot -ItemType Directory -Force | Out-Null
    }

    $resolvedBaseRoot = (Resolve-Path -Path $resolvedBaseRoot).Path
    $safeTenantName = Convert-ToAssessmentPathComponent -Value $TenantDisplayName -Fallback 'Tenant'
    $safeOutputProfileLabel = Convert-ToAssessmentPathComponent -Value $OutputProfileFolderLabel -Fallback 'Profile'
    $expectedFolderName = "{0} Assessment Reporting {1}" -f $safeTenantName, $safeOutputProfileLabel

    if ([string]::Equals((Split-Path -Path $resolvedBaseRoot -Leaf), $expectedFolderName, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $resolvedBaseRoot
    }

    $resolvedAssessmentFolder = Join-Path -Path $resolvedBaseRoot -ChildPath $expectedFolderName
    if (-not (Test-Path -Path $resolvedAssessmentFolder)) {
        New-Item -Path $resolvedAssessmentFolder -ItemType Directory -Force | Out-Null
    }

    return (Resolve-Path -Path $resolvedAssessmentFolder).Path
}

$script:ImportExcelReady = $false
$script:TenantHtmlHelpersLoaded = $false
$script:TenantQuestionnaireHelperLoaded = $false

function Ensure-ImportExcelReady {
    [CmdletBinding()]
    param()

    if ($script:ImportExcelReady) {
        return
    }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
    }

    Import-Module ImportExcel -ErrorAction Stop
    $script:ImportExcelReady = $true
}

function Ensure-TenantHtmlHelpersLoaded {
    [CmdletBinding()]
    param()

    if ($script:TenantHtmlHelpersLoaded) {
        return
    }

    if (Test-Path -Path $tenantHtmlHelperFunctionsPath) {
        . $tenantHtmlHelperFunctionsPath
    }
    else {
        Write-Warning "Optional HTML helper function script not found: $tenantHtmlHelperFunctionsPath."
    }

    if (Test-Path -Path $tenantHtmlReportPath) {
        . $tenantHtmlReportPath
    }
    else {
        Write-Warning "Optional HTML helper script not found: $tenantHtmlReportPath. Built-in technical HTML generation remains available, but PDF export helpers may be unavailable."
    }

    $script:TenantHtmlHelpersLoaded = $true
}

function Test-AssessmentHelperCommands {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$CommandName
    )

    $missing = @(
        $CommandName | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
    )
    if ($missing.Count -gt 0) {
        throw ("Required assessment helper command(s) are unavailable: {0}. Validate HTML helper script wiring in '{1}' and '{2}'." -f ($missing -join ', '), $tenantHtmlHelperFunctionsPath, $tenantHtmlReportPath)
    }
}

function Ensure-TenantQuestionnaireHelperLoaded {
    [CmdletBinding()]
    param()

    if ($script:TenantQuestionnaireHelperLoaded) {
        return
    }

    if (-not (Test-Path -Path $tenantQuestionnairePath)) {
        Write-Warning "Optional questionnaire helper script not found: $tenantQuestionnairePath. Questionnaire export will be skipped."
        return
    }

    . $tenantQuestionnairePath
    $script:TenantQuestionnaireHelperLoaded = $true
}

function Write-AssessmentArtifactManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BaseExportPath,
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Artifacts,
        [Parameter(Mandatory)]
        [string]$OutputProfileLabel,
        [Parameter(Mandatory)]
        [string]$ReportingMode,
        [Parameter(Mandatory)]
        [bool]$CollectionOnly,
        [Parameter(Mandatory)]
        [bool]$ExportOnly
    )

    return Write-ArrayaAssessmentArtifactManifest `
        -BaseExportPath $BaseExportPath `
        -Artifacts $Artifacts `
        -OutputProfileLabel $OutputProfileLabel `
        -ReportingMode $ReportingMode `
        -CollectionOnly $CollectionOnly `
        -ExportOnly $ExportOnly
}

function Get-CurrentProcessMemorySnapshot {
    [CmdletBinding()]
    param()

    try {
        $proc = Get-Process -Id $PID -ErrorAction Stop
        return [PSCustomObject]@{
            WorkingSetMB = [math]::Round(($proc.WorkingSet64 / 1MB), 2)
            PrivateMB    = [math]::Round(($proc.PrivateMemorySize64 / 1MB), 2)
            PagedMB      = [math]::Round(($proc.PagedMemorySize64 / 1MB), 2)
            HeapMB       = [math]::Round(([GC]::GetTotalMemory($false) / 1MB), 2)
        }
    }
    catch {
        return [PSCustomObject]@{
            WorkingSetMB = 0
            PrivateMB    = 0
            PagedMB      = 0
            HeapMB       = 0
        }
    }
}

function Ensure-AssessmentCollectorContext {
    [CmdletBinding()]
    param()

    if (-not ($script:tenantStatsHash -is [System.Collections.IDictionary])) {
        $script:tenantStatsHash = @{}
    }

    if ($null -eq $script:AssessmentContext) {
        $script:AssessmentContext = New-ArrayaAssessmentContext `
            -ExportFileLocation $ExportDetails `
            -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($reportingMode)) `
            -TenantStats $script:tenantStatsHash `
            -Policies ([ordered]@{}) `
            -Runtime ([ordered]@{}) `
            -Metadata ([ordered]@{ StartedAt = $global:InitialStart })
    }

    $script:AssessmentContext.TenantStats = $script:tenantStatsHash
    $script:AssessmentContext.ExportFileLocation = $ExportDetails
    $script:AssessmentContext.Policies['ReportingMode'] = ((Get-Culture).TextInfo.ToTitleCase($reportingMode))
    $script:AssessmentContext.Policies['CollectionDepth'] = $script:CollectionDepthPolicy
    $script:AssessmentContext.Metadata['StartedAt'] = $global:InitialStart
    $script:AssessmentContext.Runtime['MailboxUsageGraphLookup'] = $script:MailboxUsageGraphLookup
    $script:AssessmentContext.Runtime['UnifiedGroupsInventoryCache'] = $script:UnifiedGroupsInventoryCache
    $script:AssessmentContext.Runtime['Office365GroupsActivityMailboxLookup'] = $script:Office365GroupsActivityMailboxLookup

    return $script:AssessmentContext
}

function Sync-CollectorModuleRuntimeContext {
    [CmdletBinding()]
    param()

    Ensure-AssessmentCollectorContext | Out-Null
}

function Sync-CollectorModuleRuntimeContextBack {
    [CmdletBinding()]
    param()

    $context = Ensure-AssessmentCollectorContext
    $script:tenantStatsHash = $context.TenantStats
    if ($context.Runtime.Contains('MailboxUsageGraphLookup')) {
        $script:MailboxUsageGraphLookup = $context.Runtime['MailboxUsageGraphLookup']
    }
    if ($context.Runtime.Contains('UnifiedGroupsInventoryCache')) {
        $script:UnifiedGroupsInventoryCache = $context.Runtime['UnifiedGroupsInventoryCache']
    }
    if ($context.Runtime.Contains('Office365GroupsActivityMailboxLookup')) {
        $script:Office365GroupsActivityMailboxLookup = $context.Runtime['Office365GroupsActivityMailboxLookup']
    }
}

function Initialize-AssessmentProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$TotalSteps
    )

    $script:AssessmentProgressId = 90
    $script:AssessmentProgressState = [ordered]@{
        Total   = [Math]::Max($TotalSteps, 1)
        Current = 0
    }
    $script:AssessmentStepMetrics = New-Object System.Collections.Generic.List[object]

    if (Test-ShowAssessmentProgress) {
        Write-Progress -Id $script:AssessmentProgressId -Activity 'Assessment progress' -Status "[0/$($script:AssessmentProgressState.Total)] Starting" -PercentComplete 0
    }
}

function Clear-TransientGraphProgress {
    [CmdletBinding()]
    param()

    $overallProgressId = if ($script:AssessmentProgressId) { [int]$script:AssessmentProgressId } else { 90 }
    for ($progressId = 0; $progressId -lt $overallProgressId; $progressId++) {
        try { Write-Progress -Id $progressId -Completed } catch {}
    }
}

function Invoke-AssessmentProgressStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock
    )

    if (-not $script:AssessmentProgressState) {
        return (& $ScriptBlock)
    }

    $script:AssessmentProgressState.Current++
    $current = $script:AssessmentProgressState.Current
    $total = $script:AssessmentProgressState.Total
    $percent = [math]::Round(($current / $total) * 100, 2)
    $memoryBefore = Get-CurrentProcessMemorySnapshot
    $stepStart = Get-Date

    if (Test-ShowAssessmentProgress) {
        Write-Progress -Id $script:AssessmentProgressId -Activity 'Assessment progress' -Status "[$current/$total] $Name" -PercentComplete $percent
    }
    try {
        Sync-CollectorModuleRuntimeContext
        $null = & $ScriptBlock
        Sync-CollectorModuleRuntimeContextBack
        Write-Host ("  Overall progress: {0}/{1} ({2}%) - {3}" -f $current, $total, $percent, $Name) -ForegroundColor Cyan
    }
    finally {
        Sync-CollectorModuleRuntimeContextBack
        $memoryAfter = Get-CurrentProcessMemorySnapshot
        $elapsed = (Get-Date) - $stepStart
        $script:AssessmentStepMetrics.Add([PSCustomObject]@{
            StepName             = $Name
            StartedAt            = $stepStart
            DurationSeconds      = [math]::Round($elapsed.TotalSeconds, 3)
            WorkingSetStartMB    = $memoryBefore.WorkingSetMB
            WorkingSetEndMB      = $memoryAfter.WorkingSetMB
            WorkingSetDeltaMB    = [math]::Round(($memoryAfter.WorkingSetMB - $memoryBefore.WorkingSetMB), 2)
            PrivateStartMB       = $memoryBefore.PrivateMB
            PrivateEndMB         = $memoryAfter.PrivateMB
            PrivateDeltaMB       = [math]::Round(($memoryAfter.PrivateMB - $memoryBefore.PrivateMB), 2)
            ManagedHeapStartMB   = $memoryBefore.HeapMB
            ManagedHeapEndMB     = $memoryAfter.HeapMB
            ManagedHeapDeltaMB   = [math]::Round(($memoryAfter.HeapMB - $memoryBefore.HeapMB), 2)
        }) | Out-Null
        Clear-TransientGraphProgress
    }
}

function Invoke-ProfileAwareAssessmentStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [bool]$Enabled,
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,
        [Parameter(Mandatory = $false)]
        [string]$SkipReason = 'Disabled by output profile'
    )

    if ($Enabled) {
        Invoke-AssessmentProgressStep -Name $Name -ScriptBlock $ScriptBlock
        return
    }

    $profileLabelForSkip = if ([string]::IsNullOrWhiteSpace($script:EffectiveOutputProfileLabel)) { $OutputProfile } else { $script:EffectiveOutputProfileLabel }
    $skipMessage = "Skipped by output profile '$profileLabelForSkip': $SkipReason"
    Invoke-AssessmentProgressStep -Name $Name -ScriptBlock {
        Write-Host ("{0} ...Skipped" -f $Name) -ForegroundColor DarkYellow
        Write-Log -Type INFO -Message ("[{0}] {1}" -f $Name, $skipMessage) -ExportFileLocation $ExportDetails
    }
}

function Complete-AssessmentProgress {
    [CmdletBinding()]
    param()

    if (Test-ShowAssessmentProgress) {
        Write-Progress -Id $script:AssessmentProgressId -Activity 'Assessment progress' -Completed
    }
    Clear-TransientGraphProgress
}

function Test-ShowAssessmentProgress {
    [CmdletBinding()]
    param()

    return ($env:ARRAYA_SHOW_ASSESSMENT_PROGRESS -match '^(1|true|yes)$')
}

function Test-ShowCollectorDiagnostics {
    [CmdletBinding()]
    param()

    if ($env:ARRAYA_SHOW_COLLECTOR_DIAGNOSTICS -match '^(1|true|yes)$') {
        return $true
    }

    return (
        $DebugPreference -ne [System.Management.Automation.ActionPreference]::SilentlyContinue -or
        $VerbosePreference -ne [System.Management.Automation.ActionPreference]::SilentlyContinue
    )
}

function Write-AssessmentStepMetricsSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExportFileLocation
    )

    if (-not $script:AssessmentStepMetrics -or $script:AssessmentStepMetrics.Count -eq 0) {
        return
    }

    $topSlow = @($script:AssessmentStepMetrics | Sort-Object DurationSeconds -Descending | Select-Object -First 10)
    $topMemory = @($script:AssessmentStepMetrics | Sort-Object PrivateDeltaMB -Descending | Select-Object -First 10)
    $showConsoleDiagnostics = Test-ShowCollectorDiagnostics

    if ($showConsoleDiagnostics) {
        Write-Host ""
        Write-Host "Collector performance (top 10 by duration):" -ForegroundColor DarkCyan
    }
    foreach ($item in $topSlow) {
        if ($showConsoleDiagnostics) {
            Write-Host ("  {0}: {1}s (Private Δ {2} MB, Heap Δ {3} MB)" -f $item.StepName, $item.DurationSeconds, $item.PrivateDeltaMB, $item.ManagedHeapDeltaMB) -ForegroundColor DarkGray
        }
        Write-Log -Type INFO -Message ("[CollectorMetrics][Duration] Step='{0}' DurationSeconds={1} PrivateDeltaMB={2} ManagedHeapDeltaMB={3}" -f $item.StepName, $item.DurationSeconds, $item.PrivateDeltaMB, $item.ManagedHeapDeltaMB) -ExportFileLocation $ExportFileLocation
    }

    if ($showConsoleDiagnostics) {
        Write-Host ""
        Write-Host "Collector memory impact (top 10 by private delta):" -ForegroundColor DarkCyan
    }
    foreach ($item in $topMemory) {
        if ($showConsoleDiagnostics) {
            Write-Host ("  {0}: Private Δ {1} MB (Duration {2}s, Heap Δ {3} MB)" -f $item.StepName, $item.PrivateDeltaMB, $item.DurationSeconds, $item.ManagedHeapDeltaMB) -ForegroundColor DarkGray
        }
        Write-Log -Type INFO -Message ("[CollectorMetrics][Memory] Step='{0}' PrivateDeltaMB={1} DurationSeconds={2} ManagedHeapDeltaMB={3}" -f $item.StepName, $item.PrivateDeltaMB, $item.DurationSeconds, $item.ManagedHeapDeltaMB) -ExportFileLocation $ExportFileLocation
    }
}

function Write-CollectorInventoryMatrix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory = $false)]
        [string]$ExportFileLocation
    )

    $inventory = @(
        [PSCustomObject]@{ Collector = 'Users'; Key = 'Users'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Entra Groups'; Key = 'EntraIDGroups'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Mailboxes'; Key = 'AllMailboxes'; Source = 'EXO'; Consumers = 'Workbook, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Primary Mailbox Stats'; Key = 'PrimaryMailboxStats'; Source = 'Graph+EXO'; Consumers = 'Workbook, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'SharePoint Sites'; Key = 'SharePoint'; Source = 'Graph/SPO'; Consumers = 'Workbook, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'OneDrive Sites'; Key = 'OneDrive'; Source = 'Graph/SPO'; Consumers = 'Workbook, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Teams Inventory'; Key = 'AllTeams'; Source = 'Graph/Teams'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Conditional Access'; Key = 'ConditionalAccessPolicies'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Secure Score'; Key = 'SecuritySecureScore'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Secure Score Actions'; Key = 'SecureScoreActions'; Source = 'Graph'; Consumers = 'Workbook, BestPractices'; Count = 0 }
        [PSCustomObject]@{ Collector = 'License SKUs'; Key = 'LicenseSKUs'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Email Activity Top Senders'; Key = 'EmailActivityTopSenders'; Source = 'Graph'; Consumers = 'Workbook, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Email Activity Top Receivers'; Key = 'EmailActivityTopReceivers'; Source = 'Graph'; Consumers = 'Workbook, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Teams Activity Top Users'; Key = 'TeamsActivityTopUsers'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Office 365 Groups Activity Top Groups'; Key = 'Office365GroupsActivityTopGroups'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Employee Experience Summary'; Key = 'EmployeeExperienceInsightsSummary'; Source = 'Graph'; Consumers = 'Workbook, BestPractices'; Count = 0 }
    )

    foreach ($entry in $inventory) {
        if (-not $TenantStatsHash.ContainsKey($entry.Key) -or $null -eq $TenantStatsHash[$entry.Key]) {
            $entry.Count = 0
            continue
        }

        $value = $TenantStatsHash[$entry.Key]
        if ($value -is [hashtable] -or $value -is [System.Collections.Specialized.OrderedDictionary]) {
            $entry.Count = @($value.Keys).Count
        }
        elseif ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
            $entry.Count = @($value).Count
        }
        else {
            $entry.Count = 1
        }
    }

    $showConsoleDiagnostics = Test-ShowCollectorDiagnostics
    if ($showConsoleDiagnostics) {
        Write-Host ""
        Write-Host "Collector inventory (row counts):" -ForegroundColor DarkCyan
    }
    foreach ($entry in $inventory) {
        if ($showConsoleDiagnostics) {
            Write-Host ("  {0}: {1} row(s) [{2}] -> {3}" -f $entry.Collector, $entry.Count, $entry.Source, $entry.Consumers) -ForegroundColor DarkGray
        }
        Write-Log -Type INFO -Message ("[CollectorInventory] Collector='{0}' Key='{1}' Count={2} Source='{3}' Consumers='{4}'" -f $entry.Collector, $entry.Key, $entry.Count, $entry.Source, $entry.Consumers) -ExportFileLocation $ExportFileLocation
    }
}

function Write-ExoStatisticsFailureSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OperationName,
        [AllowNull()]
        [array]$Failures
    )

    if (-not $Failures -or $Failures.Count -eq 0) {
        return
    }

    $sampleText = @(
        $Failures |
            Select-Object -First 5 |
            ForEach-Object {
                $label = if ([string]::IsNullOrWhiteSpace([string]$_.DisplayName)) {
                    if ([string]::IsNullOrWhiteSpace([string]$_.Identity)) { '<unknown>' } else { $_.Identity }
                } else {
                    $_.DisplayName
                }
                "$label ($($_.Reason))"
            }
    ) -join '; '

    Write-Log -Type WARNING -Message "[$OperationName] Skipped $($Failures.Count) object(s). Sample failures: $sampleText" -ExportFileLocation $ExportDetails
}

function Convert-ToMailboxGuidKey {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $GuidValue
    )

    if ($null -eq $GuidValue) {
        return $null
    }

    $text = [string]$GuidValue
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return $text.Trim().Trim('{}').ToLowerInvariant()
}

function Get-MailboxGuidCandidateKeys {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $MailboxRecord,
        [switch]$IncludeArchiveGuid
    )

    if ($null -eq $MailboxRecord) {
        return @()
    }

    $candidateKeys = New-Object System.Collections.Generic.List[string]
    if (
        $IncludeArchiveGuid -and
        $MailboxRecord.PSObject.Properties['ArchiveGuid'] -and
        $MailboxRecord.ArchiveGuid
    ) {
        $archiveKey = Convert-ToMailboxGuidKey -GuidValue $MailboxRecord.ArchiveGuid
        if ($archiveKey -and -not $candidateKeys.Contains($archiveKey)) {
            $candidateKeys.Add($archiveKey)
        }
    }

    foreach ($propertyName in @('ExchangeGuid', 'Guid', 'MailboxGuid')) {
        if (-not ($MailboxRecord.PSObject.Properties[$propertyName])) {
            continue
        }

        $value = $MailboxRecord.PSObject.Properties[$propertyName].Value
        if (-not $value) {
            continue
        }

        $key = Convert-ToMailboxGuidKey -GuidValue $value
        if ($key -and -not $candidateKeys.Contains($key)) {
            $candidateKeys.Add($key)
        }
    }

    return @($candidateKeys)
}

function Test-MailboxStatCached {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $MailboxRecord,
        [AllowNull()]
        $StatsHash,
        [switch]$IncludeArchiveGuid
    )

    if ($null -eq $MailboxRecord -or $null -eq $StatsHash) {
        return $false
    }

    $candidateKeys = Get-MailboxGuidCandidateKeys -MailboxRecord $MailboxRecord -IncludeArchiveGuid:$IncludeArchiveGuid
    if (-not $candidateKeys -or $candidateKeys.Count -eq 0) {
        return $false
    }

    foreach ($key in $candidateKeys) {
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }

        if (
            ($StatsHash.PSObject.Methods['ContainsKey'] -and $StatsHash.ContainsKey($key)) -or
            ($StatsHash -is [System.Collections.IDictionary] -and $StatsHash.Contains($key))
        ) {
            return $true
        }
    }

    return $false
}

function Test-ShouldCollectUnifiedGroupMailboxStats {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$DetailLevel
    )

    if ($script:CollectionDepthPolicy -and $script:CollectionDepthPolicy.PSObject.Properties['CollectUnifiedGroupMailboxStats']) {
        return [bool]$script:CollectionDepthPolicy.CollectUnifiedGroupMailboxStats
    }

    $isMinimum = $false
    if ($script:CollectionDepthPolicy -and $script:CollectionDepthPolicy.PSObject.Properties['IsMinimum']) {
        $isMinimum = ($script:CollectionDepthPolicy.IsMinimum -eq $true)
    }
    elseif (-not [string]::IsNullOrWhiteSpace($DetailLevel) -and $DetailLevel.ToLowerInvariant() -eq 'minimum') {
        $isMinimum = $true
    }
    if ($isMinimum) {
        return $false
    }

    return $true
}

function Get-Office365GroupsActivityMailboxLookup {
    [CmdletBinding()]
    param()

    if (
        $script:Office365GroupsActivityMailboxLookup -and
        $script:Office365GroupsActivityMailboxLookup.PSObject.Properties['ByGroupId'] -and
        $script:Office365GroupsActivityMailboxLookup.PSObject.Properties['ByPrimarySmtpAddress']
    ) {
        return $script:Office365GroupsActivityMailboxLookup
    }

    $lookup = [ordered]@{
        Rows               = 0
        DownloadSucceeded  = $false
        ByGroupId          = @{}
        ByPrimarySmtpAddress = @{}
    }

    try {
        $uri = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D180')"
        $rows = @(Export-ArrayaGraphReportCsv -Uri $uri -Activity 'Office 365 Groups activity detail report' -Headers $global:GraphHeaders)
        $lookup.Rows = @($rows).Count

        foreach ($row in $rows) {
            $groupId = Get-GraphReportFieldValue -Row $row -FieldNames @('Group Id', 'GroupId')
            if (-not [string]::IsNullOrWhiteSpace($groupId)) {
                $groupIdKey = $groupId.Trim().ToLowerInvariant()
                if (-not $lookup.ByGroupId.ContainsKey($groupIdKey)) {
                    $lookup.ByGroupId[$groupIdKey] = $row
                }
            }

            $primarySmtp = Get-GraphReportFieldValue -Row $row -FieldNames @(
                'Group Principal Name',
                'Group Email',
                'Group Email Address',
                'Group Primary SMTP Address',
                'Group SMTP Address'
            )
            if (-not [string]::IsNullOrWhiteSpace($primarySmtp)) {
                $smtpKey = $primarySmtp.Trim().ToLowerInvariant()
                if (-not $lookup.ByPrimarySmtpAddress.ContainsKey($smtpKey)) {
                    $lookup.ByPrimarySmtpAddress[$smtpKey] = $row
                }
            }
        }

        $lookup.DownloadSucceeded = $true
    }
    catch {
        Write-Log -Type WARNING -Message "[Get-Office365GroupsActivityMailboxLookup] Unable to download Office 365 Groups activity detail report: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }

    $script:Office365GroupsActivityMailboxLookup = [pscustomobject]$lookup
    return $script:Office365GroupsActivityMailboxLookup
}

function Get-GroupMailboxActivityRowForUnifiedGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$GroupRecord,
        [Parameter(Mandatory = $true)]
        [object]$ActivityLookup
    )

    if (-not $GroupRecord -or -not $ActivityLookup) {
        return $null
    }

    if ($ActivityLookup.PSObject.Properties['ByGroupId'] -and $ActivityLookup.ByGroupId) {
        $groupIdCandidates = @(
            if ($GroupRecord.PSObject.Properties['ExternalDirectoryObjectId']) { [string]$GroupRecord.ExternalDirectoryObjectId }
            if ($GroupRecord.PSObject.Properties['ExternalDirectoryObjectID']) { [string]$GroupRecord.ExternalDirectoryObjectID }
            if ($GroupRecord.PSObject.Properties['GroupId']) { [string]$GroupRecord.GroupId }
            if ($GroupRecord.PSObject.Properties['Id']) { [string]$GroupRecord.Id }
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($candidateId in $groupIdCandidates) {
            $groupIdKey = $candidateId.Trim().ToLowerInvariant()
            if ($ActivityLookup.ByGroupId.ContainsKey($groupIdKey)) {
                return $ActivityLookup.ByGroupId[$groupIdKey]
            }
        }
    }

    if ($ActivityLookup.PSObject.Properties['ByPrimarySmtpAddress'] -and $ActivityLookup.ByPrimarySmtpAddress) {
        $smtpCandidates = @(
            if ($GroupRecord.PSObject.Properties['PrimarySmtpAddress']) { [string]$GroupRecord.PrimarySmtpAddress }
            if ($GroupRecord.PSObject.Properties['WindowsEmailAddress']) { [string]$GroupRecord.WindowsEmailAddress }
            if ($GroupRecord.PSObject.Properties['UserPrincipalName']) { [string]$GroupRecord.UserPrincipalName }
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($candidateSmtp in $smtpCandidates) {
            $smtpKey = $candidateSmtp.Trim().ToLowerInvariant()
            if ($ActivityLookup.ByPrimarySmtpAddress.ContainsKey($smtpKey)) {
                return $ActivityLookup.ByPrimarySmtpAddress[$smtpKey]
            }
        }
    }

    return $null
}

function New-GroupMailboxStatFromActivityRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$GroupRecord,
        [Parameter(Mandatory = $true)]
        [object]$ActivityRow
    )

    if (-not $GroupRecord -or -not $ActivityRow) {
        return $null
    }

    $storageBytes = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $ActivityRow -FieldNames @(
        'Exchange Mailbox Storage Used (Byte)',
        'Exchange Mailbox Storage Used (Bytes)',
        'Exchange Mailbox Storage Used'
    ))

    $itemCountValue = Get-GraphReportFieldValue -Row $ActivityRow -FieldNames @(
        'Exchange Mailbox Total Item Count',
        'Exchange Mailbox Item Count'
    )
    if ([string]::IsNullOrWhiteSpace($itemCountValue)) {
        $itemCountValue = '0'
    }

    $displayName = if (
        $GroupRecord.PSObject.Properties['DisplayName'] -and
        -not [string]::IsNullOrWhiteSpace([string]$GroupRecord.DisplayName)
    ) {
        [string]$GroupRecord.DisplayName
    }
    else {
        $graphDisplayName = Get-GraphReportFieldValue -Row $ActivityRow -FieldNames @('Group Display Name')
        if (-not [string]::IsNullOrWhiteSpace($graphDisplayName)) { $graphDisplayName } else { '<Unified Group>' }
    }

    $mailboxGuid = if ($GroupRecord.PSObject.Properties['ExchangeGuid'] -and $GroupRecord.ExchangeGuid) {
        $GroupRecord.ExchangeGuid
    }
    elseif ($GroupRecord.PSObject.Properties['Guid'] -and $GroupRecord.Guid) {
        $GroupRecord.Guid
    }
    else {
        $null
    }

    return [PSCustomObject]@{
        DisplayName               = $displayName
        TotalItemSize             = "$([math]::Round($storageBytes / 1GB, 4)) GB ($storageBytes bytes)"
        TotalItemSizeBytes        = [int64]$storageBytes
        ItemCount                 = [string]$itemCountValue
        TotalDeletedItemSize      = "0 GB (0 bytes)"
        TotalDeletedItemSizeBytes = [int64]0
        MailboxType               = 'GroupMailbox'
        MailboxGuid               = $mailboxGuid
    }
}

# ----------------------------------
# Exchange Specific Functions
# ----------------------------------
# NEW - Convert Names to EmailAddresses loop







function Get-EmailActivityInsights {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = 'Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel
    )

    $start = Get-Date
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash['EmailActivityTopSenders'] = @{}
    $script:tenantStatsHash['EmailActivityTopReceivers'] = @{}
    $script:tenantStatsHash['EmailActivitySummary'] = @{}
    $script:tenantStatsHash['TeamsActivityTopUsers'] = @{}
    $script:tenantStatsHash['Office365GroupsActivityTopGroups'] = @{}
    $script:tenantStatsHash['EmployeeExperienceInsightsSummary'] = @{}
    $script:tenantStatsHash['CollaborationActivitySummary'] = @{}
    $adminReportSettings = $null

    Write-Host "Getting email activity details ..." -ForegroundColor Cyan -NoNewline
    Write-Log -Type INFO -Message "[Get-EmailActivityInsights] START: Gathering email activity details from Microsoft Graph" -ExportFileLocation $ExportDetails

    function Convert-EmailActivityDateValue {
        param([string]$Value)

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return $null
        }

        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($Value, [ref]$parsed)) {
            return $parsed
        }

        return $null
    }

    function Get-ActivityReportRows {
        param(
            [Parameter(Mandatory)]
            [string]$ServiceName,
            [Parameter(Mandatory)]
            [string]$PeriodDuration,
            [Parameter(Mandatory)]
            [string]$FallbackUri,
            [Parameter(Mandatory)]
            [string]$FallbackActivity
        )

        $rows = @()
        $source = 'Export-ArrayaGraphReportCsv'
        $graphActivityCommand = Get-Command -Name 'Office365Custom\Get-GraphAPIActivityReport' -ErrorAction SilentlyContinue
        if ($graphActivityCommand) {
            try {
                $savedProgressPreference = $ProgressPreference
                try {
                    $ProgressPreference = 'SilentlyContinue'
                    $rows = @(
                        Office365Custom\Get-GraphAPIActivityReport -ServiceName $ServiceName -PeriodDuration $PeriodDuration -ErrorAction Stop
                    )
                }
                finally {
                    $ProgressPreference = $savedProgressPreference
                }

                if ($rows.Count -gt 0) {
                    $source = 'Office365Custom.Get-GraphAPIActivityReport'
                }
            }
            catch {
                Write-Log -Type WARNING -Message "[Get-EmailActivityInsights] Office365Custom\\Get-GraphAPIActivityReport failed for service '$ServiceName': $($_.Exception.Message). Falling back to direct Graph report URI." -ExportFileLocation $ExportDetails
                $rows = @()
            }
        }

        if ($rows.Count -eq 0) {
            $rows = @(
                Export-ArrayaGraphReportCsv -Uri $FallbackUri -Activity $FallbackActivity -Headers $global:GraphHeaders
            )
            $source = 'Export-ArrayaGraphReportCsv'
        }

        return [PSCustomObject]@{
            Rows   = @($rows)
            Source = $source
        }
    }

    try {
        $periodDuration = if ($detailLevel -eq 'minimum') { 'D90' } else { 'D180' }
        $topLimit = if ($detailLevel -eq 'minimum') { 10 } else { 25 }
        $emailActivityRows = @()
        $emailActivitySource = 'Export-ArrayaGraphReportCsv'
        $teamsNormalizedUsers = @()
        $normalizedGroups = @()
        $adminReportSettings = Get-ArrayaGraphAdminReportSettings -Headers $global:GraphHeaders
        if ($adminReportSettings) {
            $script:tenantStatsHash['EmailActivitySummary']['AdminReportSettings'] = $adminReportSettings
            if (
                $adminReportSettings.PSObject.Properties['Available'] -and
                $adminReportSettings.Available -eq $false -and
                $adminReportSettings.PSObject.Properties['ErrorMessage'] -and
                -not [string]::IsNullOrWhiteSpace([string]$adminReportSettings.ErrorMessage)
            ) {
                Write-Log -Type WARNING -Message "[Get-EmailActivityInsights] Unable to read admin report settings: $($adminReportSettings.ErrorMessage)" -ExportFileLocation $ExportDetails
            }
        }

        $emailActivityReport = Get-ActivityReportRows `
            -ServiceName 'EmailActivity' `
            -PeriodDuration $periodDuration `
            -FallbackUri "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='$periodDuration')" `
            -FallbackActivity "Email activity user detail report ($periodDuration)"
        $emailActivityRows = @($emailActivityReport.Rows)
        $emailActivitySource = [string]$emailActivityReport.Source

        if ($emailActivityRows.Count -eq 0) {
            $script:tenantStatsHash['EmailActivitySummary']['Summary'] = [PSCustomObject]@{
                ReportRows        = 0
                PeriodDuration    = $periodDuration
                Source            = $emailActivitySource
                UsersNormalized   = 0
                ActiveUsers       = 0
                TotalSendCount    = 0
                TotalReceiveCount = 0
                ReportRefreshDate = $null
                DisplayConcealedNames = if ($adminReportSettings -and $adminReportSettings.PSObject.Properties['DisplayConcealedNames']) { $adminReportSettings.DisplayConcealedNames } else { $null }
            }
            Write-Log -Type INFO -Message "[Get-EmailActivityInsights] No rows returned for email activity report." -ExportFileLocation $ExportDetails
        }

        $activityByUpn = @{}
        $reportRefreshDates = New-Object System.Collections.Generic.List[datetime]

        foreach ($row in $emailActivityRows) {
            $upn = Get-GraphReportFieldValue -Row $row -FieldNames @('User Principal Name', 'UserPrincipalName')
            if ([string]::IsNullOrWhiteSpace($upn)) {
                continue
            }
            $upnKey = $upn.Trim().ToLowerInvariant()
            $displayName = Get-GraphReportFieldValue -Row $row -FieldNames @('Display Name', 'User Display Name')
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                $displayName = $upn
            }

            $sendCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Send Count', 'SendCount'))
            $receiveCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Receive Count', 'ReceiveCount'))
            $readCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Read Count', 'ReadCount'))
            $lastActivityDate = Get-GraphReportFieldValue -Row $row -FieldNames @('Last Activity Date', 'LastActivityDate')
            $reportRefreshDate = Get-GraphReportFieldValue -Row $row -FieldNames @('Report Refresh Date', 'ReportRefreshDate')
            $isDeletedRaw = Get-GraphReportFieldValue -Row $row -FieldNames @('Is Deleted', 'IsDeleted')

            $isDeleted = $false
            if (-not [string]::IsNullOrWhiteSpace($isDeletedRaw)) {
                $parsedBool = $false
                if ([bool]::TryParse($isDeletedRaw, [ref]$parsedBool)) {
                    $isDeleted = $parsedBool
                }
                elseif ($isDeletedRaw -match '^(1|yes|y|true)$') {
                    $isDeleted = $true
                }
            }

            $parsedRefreshDate = Convert-EmailActivityDateValue -Value $reportRefreshDate
            if ($parsedRefreshDate) {
                $reportRefreshDates.Add($parsedRefreshDate) | Out-Null
            }

            if ($activityByUpn.ContainsKey($upnKey)) {
                $existing = $activityByUpn[$upnKey]
                $existing.SendCount += [int64]$sendCount
                $existing.ReceiveCount += [int64]$receiveCount
                $existing.ReadCount += [int64]$readCount
                $existing.IsDeleted = ($existing.IsDeleted -or $isDeleted)

                $existingLast = Convert-EmailActivityDateValue -Value $existing.LastActivityDate
                $currentLast = Convert-EmailActivityDateValue -Value $lastActivityDate
                if ($currentLast -and (($null -eq $existingLast) -or $currentLast -gt $existingLast)) {
                    $existing.LastActivityDate = $currentLast.ToString('yyyy-MM-dd')
                }
            }
            else {
                $activityByUpn[$upnKey] = [PSCustomObject]@{
                    UserPrincipalName = $upn
                    DisplayName       = $displayName
                    SendCount         = [int64]$sendCount
                    ReceiveCount      = [int64]$receiveCount
                    ReadCount         = [int64]$readCount
                    LastActivityDate  = $lastActivityDate
                    IsDeleted         = $isDeleted
                }
            }
        }

        $normalizedUsers = @($activityByUpn.Values)
        $topSenders = @(
            $normalizedUsers |
                Where-Object { [int64]$_.SendCount -gt 0 } |
                Sort-Object SendCount, ReceiveCount -Descending |
                Select-Object -First $topLimit
        )
        $topReceivers = @(
            $normalizedUsers |
                Where-Object { [int64]$_.ReceiveCount -gt 0 } |
                Sort-Object ReceiveCount, SendCount -Descending |
                Select-Object -First $topLimit
        )

        $senderRank = 0
        foreach ($sender in $topSenders) {
            $senderRank++
            $key = "{0:D3}-{1}" -f $senderRank, (($sender.UserPrincipalName -replace '[^a-zA-Z0-9@._-]', '_').ToLowerInvariant())
            $script:tenantStatsHash['EmailActivityTopSenders'][$key] = [PSCustomObject]@{
                Rank              = $senderRank
                UserPrincipalName = $sender.UserPrincipalName
                DisplayName       = $sender.DisplayName
                SendCount         = [int64]$sender.SendCount
                ReceiveCount      = [int64]$sender.ReceiveCount
                ReadCount         = [int64]$sender.ReadCount
                LastActivityDate  = $sender.LastActivityDate
                IsDeleted         = $sender.IsDeleted
            }
        }

        $receiverRank = 0
        foreach ($receiver in $topReceivers) {
            $receiverRank++
            $key = "{0:D3}-{1}" -f $receiverRank, (($receiver.UserPrincipalName -replace '[^a-zA-Z0-9@._-]', '_').ToLowerInvariant())
            $script:tenantStatsHash['EmailActivityTopReceivers'][$key] = [PSCustomObject]@{
                Rank              = $receiverRank
                UserPrincipalName = $receiver.UserPrincipalName
                DisplayName       = $receiver.DisplayName
                ReceiveCount      = [int64]$receiver.ReceiveCount
                SendCount         = [int64]$receiver.SendCount
                ReadCount         = [int64]$receiver.ReadCount
                LastActivityDate  = $receiver.LastActivityDate
                IsDeleted         = $receiver.IsDeleted
            }
        }

        $latestRefresh = $null
        if ($reportRefreshDates.Count -gt 0) {
            $latestRefresh = ($reportRefreshDates | Sort-Object -Descending | Select-Object -First 1).ToString('yyyy-MM-dd')
        }

        $activeUsers = @(
            $normalizedUsers | Where-Object { ([int64]$_.SendCount + [int64]$_.ReceiveCount + [int64]$_.ReadCount) -gt 0 }
        ).Count

        $script:tenantStatsHash['EmailActivitySummary']['Summary'] = [PSCustomObject]@{
            ReportRows        = [int]$emailActivityRows.Count
            PeriodDuration    = $periodDuration
            Source            = $emailActivitySource
            UsersNormalized   = [int]$normalizedUsers.Count
            ActiveUsers       = [int]$activeUsers
            TotalSendCount    = [int64](($normalizedUsers | Measure-Object -Property SendCount -Sum).Sum)
            TotalReceiveCount = [int64](($normalizedUsers | Measure-Object -Property ReceiveCount -Sum).Sum)
            ReportRefreshDate = $latestRefresh
            TopSenderCount    = [int]$topSenders.Count
            TopReceiverCount  = [int]$topReceivers.Count
            DisplayConcealedNames = if ($adminReportSettings -and $adminReportSettings.PSObject.Properties['DisplayConcealedNames']) { $adminReportSettings.DisplayConcealedNames } else { $null }
        }

        $teamsReportRows = @()
        $teamsReportSource = 'Unavailable'
        $teamsActiveUsers = 0
        try {
            $teamsActivityReport = Get-ActivityReportRows `
                -ServiceName 'TeamsUserActivity' `
                -PeriodDuration $periodDuration `
                -FallbackUri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='$periodDuration')" `
                -FallbackActivity "Teams user activity detail report ($periodDuration)"
            $teamsReportRows = @($teamsActivityReport.Rows)
            $teamsReportSource = [string]$teamsActivityReport.Source

            $teamsNormalizedUsers = @()
            foreach ($row in $teamsReportRows) {
                $upn = Get-GraphReportFieldValue -Row $row -FieldNames @('User Principal Name', 'UserPrincipalName')
                $displayName = Get-GraphReportFieldValue -Row $row -FieldNames @('Display Name', 'User Display Name')
                if ([string]::IsNullOrWhiteSpace($upn) -and [string]::IsNullOrWhiteSpace($displayName)) {
                    continue
                }
                $teamChatCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Team Chat Message Count', 'TeamChatMessageCount'))
                $privateChatCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Private Chat Message Count', 'PrivateChatMessageCount'))
                $callCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Call Count', 'CallCount'))
                $meetingCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Meeting Count', 'MeetingCount'))
                $activityScore = [int64]($teamChatCount + $privateChatCount + $callCount + $meetingCount)

                $teamsNormalizedUsers += [PSCustomObject]@{
                    UserPrincipalName       = if ([string]::IsNullOrWhiteSpace($upn)) { $displayName } else { $upn }
                    DisplayName             = if ([string]::IsNullOrWhiteSpace($displayName)) { $upn } else { $displayName }
                    TeamChatMessageCount    = [int64]$teamChatCount
                    PrivateChatMessageCount = [int64]$privateChatCount
                    CallCount               = [int64]$callCount
                    MeetingCount            = [int64]$meetingCount
                    ActivityScore           = [int64]$activityScore
                    LastActivityDate        = Get-GraphReportFieldValue -Row $row -FieldNames @('Last Activity Date', 'LastActivityDate')
                }
            }

            $teamsActiveUsers = @($teamsNormalizedUsers | Where-Object { [int64]$_.ActivityScore -gt 0 }).Count
            $teamsTopUsers = @(
                $teamsNormalizedUsers |
                    Where-Object { [int64]$_.ActivityScore -gt 0 } |
                    Sort-Object ActivityScore, TeamChatMessageCount, PrivateChatMessageCount -Descending |
                    Select-Object -First $topLimit
            )

            $teamsRank = 0
            foreach ($teamsUser in $teamsTopUsers) {
                $teamsRank++
                $identityForKey = if ([string]::IsNullOrWhiteSpace([string]$teamsUser.UserPrincipalName)) { [string]$teamsUser.DisplayName } else { [string]$teamsUser.UserPrincipalName }
                $key = "{0:D3}-{1}" -f $teamsRank, (($identityForKey -replace '[^a-zA-Z0-9@._-]', '_').ToLowerInvariant())
                $script:tenantStatsHash['TeamsActivityTopUsers'][$key] = [PSCustomObject]@{
                    Rank                    = $teamsRank
                    UserPrincipalName       = $teamsUser.UserPrincipalName
                    DisplayName             = $teamsUser.DisplayName
                    TeamChatMessageCount    = [int64]$teamsUser.TeamChatMessageCount
                    PrivateChatMessageCount = [int64]$teamsUser.PrivateChatMessageCount
                    CallCount               = [int64]$teamsUser.CallCount
                    MeetingCount            = [int64]$teamsUser.MeetingCount
                    ActivityScore           = [int64]$teamsUser.ActivityScore
                    LastActivityDate        = $teamsUser.LastActivityDate
                }
            }
        }
        catch {
            Write-Log -Type WARNING -Message "[Get-EmailActivityInsights] Unable to process Teams activity insight report. $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        $groupsReportRows = @()
        $groupsReportSource = 'Unavailable'
        $activeGroups = 0
        try {
            $groupsActivityReport = Get-ActivityReportRows `
                -ServiceName 'Office365GroupsActivity' `
                -PeriodDuration $periodDuration `
                -FallbackUri "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='$periodDuration')" `
                -FallbackActivity "Office 365 groups activity detail report ($periodDuration)"
            $groupsReportRows = @($groupsActivityReport.Rows)
            $groupsReportSource = [string]$groupsActivityReport.Source

            $normalizedGroups = @()
            foreach ($row in $groupsReportRows) {
                $groupDisplayName = Get-GraphReportFieldValue -Row $row -FieldNames @('Group Display Name', 'GroupDisplayName')
                $groupId = Get-GraphReportFieldValue -Row $row -FieldNames @('Group Id', 'GroupId')
                if ([string]::IsNullOrWhiteSpace($groupDisplayName) -and [string]::IsNullOrWhiteSpace($groupId)) {
                    continue
                }
                $receivedEmailCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Exchange Received Email Count', 'ExchangeReceivedEmailCount'))
                $mailboxItemCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Exchange Mailbox Total Item Count', 'ExchangeMailboxTotalItemCount'))
                $sharePointActiveFileCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('SharePoint Active File Count', 'SharePointActiveFileCount'))
                $yammerPostedMessageCount = Convert-GraphReportValueToInt64 -Value (Get-GraphReportFieldValue -Row $row -FieldNames @('Yammer Posted Message Count', 'YammerPostedMessageCount'))
                $activityScore = [int64]($receivedEmailCount + $mailboxItemCount + $sharePointActiveFileCount + $yammerPostedMessageCount)

                $normalizedGroups += [PSCustomObject]@{
                    GroupDisplayName              = if ([string]::IsNullOrWhiteSpace($groupDisplayName)) { $groupId } else { $groupDisplayName }
                    GroupId                       = $groupId
                    ExchangeReceivedEmailCount    = [int64]$receivedEmailCount
                    ExchangeMailboxTotalItemCount = [int64]$mailboxItemCount
                    SharePointActiveFileCount     = [int64]$sharePointActiveFileCount
                    YammerPostedMessageCount      = [int64]$yammerPostedMessageCount
                    ActivityScore                 = [int64]$activityScore
                    LastActivityDate              = Get-GraphReportFieldValue -Row $row -FieldNames @('Last Activity Date', 'LastActivityDate')
                }
            }

            $activeGroups = @($normalizedGroups | Where-Object { [int64]$_.ActivityScore -gt 0 }).Count
            $topGroups = @(
                $normalizedGroups |
                    Where-Object { [int64]$_.ActivityScore -gt 0 } |
                    Sort-Object ActivityScore, ExchangeReceivedEmailCount -Descending |
                    Select-Object -First $topLimit
            )

            $groupsRank = 0
            foreach ($group in $topGroups) {
                $groupsRank++
                $identityForKey = if ([string]::IsNullOrWhiteSpace([string]$group.GroupId)) { [string]$group.GroupDisplayName } else { [string]$group.GroupId }
                $key = "{0:D3}-{1}" -f $groupsRank, (($identityForKey -replace '[^a-zA-Z0-9@._-]', '_').ToLowerInvariant())
                $script:tenantStatsHash['Office365GroupsActivityTopGroups'][$key] = [PSCustomObject]@{
                    Rank                           = $groupsRank
                    GroupDisplayName               = $group.GroupDisplayName
                    GroupId                        = $group.GroupId
                    ActivityScore                  = [int64]$group.ActivityScore
                    ExchangeReceivedEmailCount     = [int64]$group.ExchangeReceivedEmailCount
                    ExchangeMailboxTotalItemCount  = [int64]$group.ExchangeMailboxTotalItemCount
                    SharePointActiveFileCount      = [int64]$group.SharePointActiveFileCount
                    YammerPostedMessageCount       = [int64]$group.YammerPostedMessageCount
                    LastActivityDate               = $group.LastActivityDate
                }
            }
        }
        catch {
            Write-Log -Type WARNING -Message "[Get-EmailActivityInsights] Unable to process Office 365 groups activity insight report. $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        $script:tenantStatsHash['EmployeeExperienceInsightsSummary']['Summary'] = [PSCustomObject]@{
            PeriodDuration               = $periodDuration
            EmailReportRows              = [int]$emailActivityRows.Count
            EmailActiveUsers             = [int]$activeUsers
            TeamsReportRows              = [int]$teamsReportRows.Count
            TeamsActiveUsers             = [int]$teamsActiveUsers
            GroupsReportRows             = [int]$groupsReportRows.Count
            ActiveGroups                 = [int]$activeGroups
            EmailSource                  = $emailActivitySource
            TeamsSource                  = $teamsReportSource
            GroupsSource                 = $groupsReportSource
            ReportRefreshDate            = $latestRefresh
        }
        $dormantGroups = @($normalizedGroups | Where-Object { [int64]$_.ActivityScore -le 0 }).Count
        $inactiveTeamsUsers = @($teamsNormalizedUsers | Where-Object { [int64]$_.ActivityScore -le 0 }).Count
        $script:tenantStatsHash['CollaborationActivitySummary']['Summary'] = [PSCustomObject]@{
            PeriodDuration     = $periodDuration
            TeamsReportRows    = [int]$teamsReportRows.Count
            TeamsActiveUsers   = [int]$teamsActiveUsers
            TeamsInactiveUsers = [int]$inactiveTeamsUsers
            GroupsReportRows   = [int]$groupsReportRows.Count
            ActiveGroups       = [int]$activeGroups
            DormantGroups      = [int]$dormantGroups
            TeamsSource        = $teamsReportSource
            GroupsSource       = $groupsReportSource
            ReportRefreshDate  = $latestRefresh
        }

        #Write-Host ("Top senders={0}, top receivers={1}" -f $topSenders.Count, $topReceivers.Count) -ForegroundColor DarkGray -NoNewline
        Write-Log -Type INFO -Message "[Get-EmailActivityInsights] Email activity summary: source=$emailActivitySource; period=$periodDuration; reportRows=$($emailActivityRows.Count); usersNormalized=$($normalizedUsers.Count); activeUsers=$activeUsers; topSenders=$($topSenders.Count); topReceivers=$($topReceivers.Count)." -ExportFileLocation $ExportDetails
        Write-Log -Type INFO -Message "[Get-EmailActivityInsights] Collaboration insight summary: teamsRows=$($teamsReportRows.Count); teamsActiveUsers=$teamsActiveUsers; groupsRows=$($groupsReportRows.Count); activeGroups=$activeGroups; teamsSource=$teamsReportSource; groupsSource=$groupsReportSource." -ExportFileLocation $ExportDetails
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-EmailActivityInsights] An error occurred while gathering email activity details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        $completedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($completedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-EmailActivityInsights] COMPLETED: Gathering email activity details in $($completedTime)" -ExportFileLocation $ExportDetails
    }
}










# ----------------------------------
# Collaboration Specific Functions
# ----------------------------------
function Get-AllUnifiedGroups {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    #Get Office 365 Group / Group Mailbox data with SharePoint URL data
    $start = Get-Date
    $CompletedTime = $null
    $fetchProgressId = 40
    $hashProgressId = 41
    $statsProgressId = 42
    try {
        # Ensure global hash table structure
        if (-not $script:tenantStatsHash) {
            $script:tenantStatsHash = @{}
        }
        $script:tenantStatsHash["UnifiedGroups"] = @{}
        Write-Host "Getting all unified groups (including soft deleted)..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] START: Gathering all Unified with $($detailLevel) details" -ExportFileLocation $ExportDetails
        $allUnifiedGroups = New-Object System.Collections.Generic.List[object]
        $cachedUnifiedGroups = if ($script:UnifiedGroupsInventoryCache) { @($script:UnifiedGroupsInventoryCache) } else { @() }
        if ($cachedUnifiedGroups.Count -gt 0) {
            Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Using pre-cached unified group inventory from mailbox stats phase ($($cachedUnifiedGroups.Count) group(s))." -ExportFileLocation $ExportDetails
            foreach ($group in $cachedUnifiedGroups) {
                [void]$allUnifiedGroups.Add($group)
            }
        }
        else {
            Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Querying unified groups from Exchange Online" -ExportFileLocation $ExportDetails
            Write-Host "  Querying unified groups from Exchange Online..." -ForegroundColor DarkGray
            Write-Progress -Id $fetchProgressId -Activity "Querying unified groups from Exchange Online" -Status "Starting query"
            switch ($detailLevel) {
                {$_ -in "minimum", "combined", "all"} { 
                    $DesiredProperties = @(
                        "PrimarySmtpAddress", "DisplayName", "AccessType", "RecipientTypeDetails",
                        "ExternalDirectoryObjectId",
                        "ExchangeGuid", @{Name="ManagedByDetails"; Expression={$_.ManagedByDetails -join ','}}, "Notes",
                        "SharePointSiteUrl", "ContentMailboxName", "GroupMemberCount",
                        "AllowAddGuests", "WhenSoftDeleted", "HiddenFromExchangeClientsEnabled",
                        @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, @{Name="ModeratedBy"; Expression={$_.ModeratedBy -join ','}}, "FolderPath",
                        @{Name="Description"; Expression={$_.Description -join ','}}, "WhenCreated"
                    )

                    $fetchedGroups = 0
                    Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue } |
                        Select-Object $DesiredProperties |
                        ForEach-Object {
                            [void]$allUnifiedGroups.Add($_)
                            $fetchedGroups++
                            if ($fetchedGroups -eq 1 -or ($fetchedGroups % 50) -eq 0) {
                                Write-Progress -Id $fetchProgressId -Activity "Querying unified groups from Exchange Online" -Status "Fetched $fetchedGroups groups (continuing...)"
                            }
                        }
                }
                geek {
                    $fetchedGroups = 0
                    Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue } |
                        ForEach-Object {
                            [void]$allUnifiedGroups.Add($_)
                            $fetchedGroups++
                            if ($fetchedGroups -eq 1 -or ($fetchedGroups % 50) -eq 0) {
                                Write-Progress -Id $fetchProgressId -Activity "Querying unified groups from Exchange Online" -Status "Fetched $fetchedGroups groups (continuing...)"
                            }
                        }
                }
            }
        }
        Write-Progress -Id $fetchProgressId -Activity "Querying unified groups from Exchange Online" -Completed
        #Write-Host ("  Phase 1/3 complete: fetched {0} unified groups" -f $allUnifiedGroups.Count) -ForegroundColor DarkGray
        
        $totalUnifiedGroups = $allUnifiedGroups.Count
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Adding Unified Group data to Hash" -ExportFileLocation $ExportDetails
        #Write-Host "  Phase 2/3: Adding unified group data to hash..." -ForegroundColor DarkGray
        $progressTotal = [Math]::Max($totalUnifiedGroups, 1)
        $progressIndex = 0
        foreach ($group in $allUnifiedGroups) {
            $progressIndex++
            $groupLabel = if ([string]::IsNullOrWhiteSpace([string]$group.DisplayName)) { [string]$group.PrimarySmtpAddress } else { [string]$group.DisplayName }
            Write-ProgressHelper -Total $progressTotal -Id $hashProgressId -Index $progressIndex -Activity "Adding Unified Group data to Hash" -Operation $groupLabel
            #$key = $group.ExchangeGuid.ToString()
            $script:tenantStatsHash["UnifiedGroups"][$group.PrimarySmtpAddress] = $group
        }
        Write-ProgressHelper -Total $progressTotal -Id $hashProgressId -Activity "Adding Unified Group data to Hash" -Completed
        #Write-Host "  Phase 2/3 complete: unified group hash populated" -ForegroundColor DarkGray

        # Get Unified Group Statistics
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Gathering all Unified Group Statistics" -ExportFileLocation $ExportDetails
        if (-not $script:tenantStatsHash.ContainsKey("PrimaryMailboxStats")) {
            $script:tenantStatsHash["PrimaryMailboxStats"] = @{}
        }
        Write-Host "  Gathering unified group mailbox statistics..." -ForegroundColor DarkGray

        if ($totalUnifiedGroups -gt 0) {
            $collectUnifiedGroupMailboxStats = Test-ShouldCollectUnifiedGroupMailboxStats -DetailLevel $detailLevel
            $cachedStatsCount = 0
            $graphFilledCount = 0
            $graphGroupsActivityFilledCount = 0
            $graphMailboxUsageFilledCount = 0
            $fetchedStatsCount = 0
            $exoFilledCount = 0
            $exoFallbackRequestedCount = 0
            $unresolvedStatsCount = 0
            $groupsNeedingStats = New-Object System.Collections.Generic.List[object]
            $groupsWithoutGuidCount = 0
            $groupActivityLookup = Get-Office365GroupsActivityMailboxLookup
            $groupActivityReportRows = if ($groupActivityLookup -and $groupActivityLookup.PSObject.Properties['Rows']) { [int]$groupActivityLookup.Rows } else { 0 }

            foreach ($group in $allUnifiedGroups) {
                $groupGuidKey = $null
                if ($group -and $group.PSObject.Properties['ExchangeGuid'] -and $group.ExchangeGuid) {
                    $groupGuidKey = Convert-ToMailboxGuidKey -GuidValue $group.ExchangeGuid
                }
                if ([string]::IsNullOrWhiteSpace($groupGuidKey)) {
                    $groupsWithoutGuidCount++
                    continue
                }

                if (-not [string]::IsNullOrWhiteSpace($groupGuidKey) -and $script:tenantStatsHash["PrimaryMailboxStats"].ContainsKey($groupGuidKey)) {
                    $cachedStatsCount++
                    continue
                }

                $groupActivityRow = Get-GroupMailboxActivityRowForUnifiedGroup -GroupRecord $group -ActivityLookup $groupActivityLookup
                if ($groupActivityRow) {
                    $groupActivityStat = New-GroupMailboxStatFromActivityRow -GroupRecord $group -ActivityRow $groupActivityRow
                    if ($groupActivityStat) {
                        $script:tenantStatsHash["PrimaryMailboxStats"][$groupGuidKey] = $groupActivityStat
                        $graphGroupsActivityFilledCount++
                        $graphFilledCount++
                        continue
                    }
                }

                $graphLookupKey = if (-not [string]::IsNullOrWhiteSpace([string]$group.PrimarySmtpAddress)) { ([string]$group.PrimarySmtpAddress).ToLowerInvariant() } else { $null }
                if (
                    -not [string]::IsNullOrWhiteSpace($groupGuidKey) -and
                    $script:MailboxUsageGraphLookup -and
                    $graphLookupKey -and
                    $script:MailboxUsageGraphLookup.ContainsKey($graphLookupKey)
                ) {
                    $graphData = $script:MailboxUsageGraphLookup[$graphLookupKey]
                    $storageBytes = 0
                    $deletedBytes = 0
                    $itemCount = "0"
                    if ($graphData.'Storage Used (Byte)') { $storageBytes = [double]$graphData.'Storage Used (Byte)' }
                    if ($graphData.'Deleted Item Size (Byte)') { $deletedBytes = [double]$graphData.'Deleted Item Size (Byte)' }
                    if ($graphData.'Item Count') { $itemCount = $graphData.'Item Count' }

                    $script:tenantStatsHash["PrimaryMailboxStats"][$groupGuidKey] = [PSCustomObject]@{
                        DisplayName                = $group.DisplayName
                        TotalItemSize              = "$([math]::Round($storageBytes / 1GB, 4)) GB ($storageBytes bytes)"
                        TotalItemSizeBytes         = [int64]$storageBytes
                        ItemCount                  = $itemCount
                        TotalDeletedItemSize       = "$([math]::Round($deletedBytes / 1GB, 4)) GB ($deletedBytes bytes)"
                        TotalDeletedItemSizeBytes  = [int64]$deletedBytes
                        MailboxType                = 'GroupMailbox'
                        MailboxGuid                = $group.ExchangeGuid
                    }
                    $graphMailboxUsageFilledCount++
                    $graphFilledCount++
                    continue
                }

                [void]$groupsNeedingStats.Add($group)
            }

            if ($groupsNeedingStats.Count -gt 0 -and $collectUnifiedGroupMailboxStats) {
                $exoFallbackRequestedCount = $groupsNeedingStats.Count
                Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Reused $cachedStatsCount cached mailbox stat(s), populated $graphGroupsActivityFilledCount via Office 365 Groups activity report and $graphMailboxUsageFilledCount via mailbox usage report, then fetching EXO stats for $($groupsNeedingStats.Count) unified group(s) still missing stats." -ExportFileLocation $ExportDetails
                $allUnifiedGroupStatisticsResult = Get-ExoMailboxStatisticsSafe -MailboxObjects $groupsNeedingStats.ToArray() -ProgressActivity "Gathering Unified Group Mailbox Statistics" -ProgressId $statsProgressId
                foreach ($groupStat in $allUnifiedGroupStatisticsResult.Results) {
                    $key = Convert-ToMailboxGuidKey -GuidValue $groupStat.MailboxGuid
                    if (-not $key) { continue }
                    if (-not $script:tenantStatsHash["PrimaryMailboxStats"].ContainsKey($key)) {
                        $script:tenantStatsHash["PrimaryMailboxStats"][$key] = $groupStat
                        $exoFilledCount++
                    }
                }
                $fetchedStatsCount = @($allUnifiedGroupStatisticsResult.Results).Count
                Write-ExoStatisticsFailureSummary -OperationName 'Get-AllUnifiedGroups mailbox statistics' -Failures $allUnifiedGroupStatisticsResult.Failures
            
            }
            elseif ($groupsNeedingStats.Count -gt 0) {
                Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Unified group mailbox EXO fallback is disabled for this profile/depth. Reused $cachedStatsCount cached stat(s), populated $graphGroupsActivityFilledCount via Office 365 Groups activity report and $graphMailboxUsageFilledCount via mailbox usage report, leaving $($groupsNeedingStats.Count) group(s) unresolved." -ExportFileLocation $ExportDetails
            }
            else {
                Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] All unified group mailbox stats were already available in PrimaryMailboxStats; skipping EXO stats retrieval." -ExportFileLocation $ExportDetails
            }

            if ($groupsWithoutGuidCount -gt 0) {
                Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] $groupsWithoutGuidCount unified group(s) did not expose ExchangeGuid and were excluded from mailbox statistics lookup." -ExportFileLocation $ExportDetails
            }
            if ($groupsNeedingStats.Count -gt 0 -and $collectUnifiedGroupMailboxStats) {
                $unresolvedStatsCount = [Math]::Max(($groupsNeedingStats.Count - $exoFilledCount), 0)
            } elseif ($groupsNeedingStats.Count -gt 0) {
                $unresolvedStatsCount = $groupsNeedingStats.Count
            }
            $script:tenantStatsHash["UnifiedGroupMailboxStatsCollectionSummary"] = [PSCustomObject]@{
                ReusedFromCache             = [int]$cachedStatsCount
                GraphPopulated              = [int]$graphFilledCount
                GraphGroupsActivityRows     = [int]$groupActivityReportRows
                GraphGroupsActivityPopulated = [int]$graphGroupsActivityFilledCount
                GraphMailboxUsagePopulated  = [int]$graphMailboxUsageFilledCount
                ExoFallbackRequested        = [int]$exoFallbackRequestedCount
                ExoFallbackReturned         = [int]$fetchedStatsCount
                ExoFallbackPopulated        = [int]$exoFilledCount
                Missing                     = [int]$unresolvedStatsCount
                WithoutGuid                 = [int]$groupsWithoutGuidCount
                TotalGroups                 = [int]$totalUnifiedGroups
            }
            Write-Host ("  Unified group mailbox stats source breakdown: Cache={0}, Graph(Activity={1}, MailboxUsage={2}), EXO fallback={3}, Missing={4}" -f $cachedStatsCount, $graphGroupsActivityFilledCount, $graphMailboxUsageFilledCount, $exoFilledCount, $unresolvedStatsCount) -ForegroundColor DarkGray
            Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Mailbox stats source breakdown: Cache=$cachedStatsCount; GraphGroupsActivityRows=$groupActivityReportRows; GraphGroupsActivityPopulated=$graphGroupsActivityFilledCount; GraphMailboxUsagePopulated=$graphMailboxUsageFilledCount; GraphPopulated=$graphFilledCount; EXOFallbackRequested=$exoFallbackRequestedCount; EXOFallbackReturned=$fetchedStatsCount; EXOFallbackPopulated=$exoFilledCount; Missing=$unresolvedStatsCount; WithoutGuid=$groupsWithoutGuidCount; TotalGroups=$totalUnifiedGroups." -ExportFileLocation $ExportDetails
        }
        else {
            Write-Host "  No unified groups found" -ForegroundColor DarkGray
        }
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    }    
    catch {
        Write-Log -Type ERROR -Message "[Get-AllUnifiedGroups] An error occurred in running Get-AllUnifiedGroups function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        if (-not $CompletedTime) {
            $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        }
        $script:UnifiedGroupsInventoryCache = $null
        $script:Office365GroupsActivityMailboxLookup = $null
        Write-Progress -Id $fetchProgressId -Activity "Querying unified groups from Exchange Online" -Completed
        Write-ProgressHelper -Total 1 -Id $hashProgressId -Activity "Adding Unified Group data to Hash" -Completed
        Write-Progress -Id $statsProgressId -Activity "Gathering Unified Group Mailbox Statistics" -Completed
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] COMPLETED: Gathering all Unified Groups in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }
}



#Get SharePoint Site Details - Functional
function Get-SharePointAndOneDriveSites {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, HelpMessage = 'Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,

        [Parameter(Mandatory = $True, HelpMessage = 'Provide the service name')]
        [ValidateSet('MGGraph', 'SPO', 'API')]
        [string]$ServiceName
    )

    $start = Get-Date
    $depthPolicy = if ($script:CollectionDepthPolicy) {
        $script:CollectionDepthPolicy
    }
    else {
        Get-ArrayaCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($detailLevel.ToLowerInvariant()))
    }
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash['SharePoint'] = @{}
    $script:tenantStatsHash['OneDrive'] = @{}
    Write-Host "Getting all $($ServiceName) SharePoint Online and OneDrive Sites with $($detailLevel) ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type Info -Message "[Get-SharePointAndOneDriveSites] START: Getting all SharePoint Online and OneDrive $($detailLevel) details ($($ServiceName))" -ExportFileLocation $ExportDetails
    $graphSitesProgressId = 51
    $spoSitesProgressId = 52
    $siteDetailsProgressId = 53
    $restSitesProgressId = 54
    $siteUsageCoverage = [ordered]@{
        SharePointTotal   = 0
        SharePointMatched = 0
        SharePointMissing = 0
        OneDriveTotal     = 0
        OneDriveMatched   = 0
        OneDriveMissing   = 0
    }

    function Get-GraphCsvReportLookup {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Uri,
            [Parameter(Mandatory = $true)]
            [string]$LookupName,
            [Parameter(Mandatory = $false)]
            [hashtable]$UrlLookup
        )

        try {
            $rows = @(Export-ArrayaGraphReportCsv -Uri $Uri -Activity "$LookupName report" -Headers $global:GraphHeaders)
            $lookup = @{}
            foreach ($row in $rows) {
                $siteId = $row.'Site Id'
                if (-not [string]::IsNullOrWhiteSpace($siteId) -and -not $lookup.ContainsKey($siteId)) {
                    $lookup[$siteId] = $row
                }

                if ($UrlLookup) {
                    $siteUrl = $row.'Site URL'
                    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                        $siteUrl = $row.'Site Url'
                    }
                    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                        $siteUrl = $row.URL
                    }

                    if (-not [string]::IsNullOrWhiteSpace($siteUrl)) {
                        $normalizedSiteUrl = $siteUrl.Trim().TrimEnd('/').ToLowerInvariant()
                        if (-not $UrlLookup.ContainsKey($normalizedSiteUrl)) {
                            $UrlLookup[$normalizedSiteUrl] = $row
                        }
                    }
                }
            }
            return $lookup
        } catch {
            Write-Log -Type WARNING -Message "[Get-SharePointAndOneDriveSites] Unable to download $LookupName report: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            return @{}
        }
    }

    function Get-GraphSiteReportId {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$CompositeSiteId
        )

        $siteIdParts = $CompositeSiteId -split ','
        if ($siteIdParts.Count -ge 2) {
            return $siteIdParts[1]
        }

        return $CompositeSiteId
    }

    function Update-SiteUsageCoverage {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [bool]$IsOneDrive,
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $UsageReport
        )

        if ($IsOneDrive) {
            $siteUsageCoverage.OneDriveTotal++
            if ($UsageReport) {
                $siteUsageCoverage.OneDriveMatched++
            }
            else {
                $siteUsageCoverage.OneDriveMissing++
            }
        }
        else {
            $siteUsageCoverage.SharePointTotal++
            if ($UsageReport) {
                $siteUsageCoverage.SharePointMatched++
            }
            else {
                $siteUsageCoverage.SharePointMissing++
            }
        }
    }

    function ConvertTo-OneDriveOwnerKey {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [string]$Owner,
            [Parameter(Mandatory = $false)]
            [string]$Url
        )

        if (-not [string]::IsNullOrWhiteSpace($Owner)) {
            return $Owner.ToLowerInvariant()
        }

        if ([string]::IsNullOrWhiteSpace($Url)) {
            return $null
        }

        $pathSegment = ($Url.TrimEnd('/') -split '/')[-1]
        if ([string]::IsNullOrWhiteSpace($pathSegment)) {
            return $null
        }

        $firstSeparatorIndex = $pathSegment.IndexOf('_')
        if ($firstSeparatorIndex -lt 1) {
            return $pathSegment.ToLowerInvariant()
        }

        $localPart = $pathSegment.Substring(0, $firstSeparatorIndex)
        $domainPart = $pathSegment.Substring($firstSeparatorIndex + 1) -replace '_', '.'
        return ("{0}@{1}" -f $localPart, $domainPart).ToLowerInvariant()
    }

    function ConvertTo-OneDriveSiteKey {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [string]$Url,
            [Parameter(Mandatory = $false)]
            [string]$SiteId
        )

        if (-not [string]::IsNullOrWhiteSpace($Url)) {
            return $Url.Trim().TrimEnd('/').ToLowerInvariant()
        }

        if (-not [string]::IsNullOrWhiteSpace($SiteId)) {
            return $SiteId.Trim().ToLowerInvariant()
        }

        return $null
    }

    function Get-LikelyOneDriveOwnerFromUrl {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [string]$Url
        )

        if ([string]::IsNullOrWhiteSpace($Url)) {
            return $null
        }

        $segment = $null
        try {
            $segment = ($Url.TrimEnd('/') -split '/')[-1]
        }
        catch {
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($segment)) {
            return $null
        }

        $tokens = @($segment.Trim().ToLowerInvariant() -split '_' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($tokens.Count -lt 3) {
            return $null
        }

        # Most OneDrive personal paths map to:
        #   local_part_domain_tld
        # with onmicrosoft domains represented as:
        #   local_part_tenant_onmicrosoft_com
        $domainStartIndex = $tokens.Count - 2
        if ($tokens.Count -ge 3 -and $tokens[$tokens.Count - 2] -eq 'onmicrosoft') {
            $domainStartIndex = $tokens.Count - 3
        }

        if ($domainStartIndex -lt 1) {
            return $null
        }

        $localPart = ($tokens[0..($domainStartIndex - 1)] -join '.')
        $domainPart = ($tokens[$domainStartIndex..($tokens.Count - 1)] -join '.')
        if ([string]::IsNullOrWhiteSpace($localPart) -or [string]::IsNullOrWhiteSpace($domainPart)) {
            return $null
        }

        return ("{0}@{1}" -f $localPart, $domainPart).ToLowerInvariant()
    }

    function ConvertTo-NormalizedSiteData {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [object]$Site,
            [Parameter(Mandatory = $true)]
            [bool]$IsOneDrive,
            [Parameter(Mandatory = $false)]
            [ValidateSet('SPO', 'MGGraph', 'API')]
            [string]$Source = 'MGGraph',
            [Parameter(Mandatory = $false)]
            [psobject]$UsageReport
        )

        $storageUsageCurrent = 0
        $storageUsageCurrentGB = 0
        $owner = $null
        $groupId = $null
        $isTeamsChannelConnected = $false
        $isTeamsConnected = $false
        $title = $null
        $url = $null
        $template = $null
        $lastContentModifiedDate = $null
        $status = $null
        $archiveStatus = $null
        $lockState = $null
        $storageQuota = $null
        $sharingCapability = $null
        $defaultLinkPermission = $null
        $defaultSharingLinkType = $null

        switch ($Source) {
            'SPO' {
                $storageUsageCurrent = [int]($Site.StorageUsageCurrent ?? 0)
                $owner = $Site.Owner
                $groupId = $Site.GroupId
                $isTeamsChannelConnected = ($Site.IsTeamsChannelConnected -eq $true)
                $isTeamsConnected = ($Site.IsTeamsConnected -eq $true)
                $title = $Site.Title
                $url = $Site.Url
                $template = $Site.Template
                $lastContentModifiedDate = $Site.LastContentModifiedDate
                $status = $Site.Status
                $archiveStatus = $Site.ArchiveStatus
                $lockState = $Site.LockState
                $storageQuota = $Site.StorageQuota
                $sharingCapability = $Site.SharingCapability
                $defaultLinkPermission = $Site.DefaultLinkPermission
                $defaultSharingLinkType = $Site.DefaultSharingLinkType
            }
            'API' {
                $additionalProperties = $Site.additionalProperties
                $storageUsageCurrent = [int]($additionalProperties.storageUsage ?? 0)
                $owner = $additionalProperties.owner
                $groupId = $Site.groupId
                $isTeamsChannelConnected = ($additionalProperties.isTeamsChannelConnected -eq $true)
                $isTeamsConnected = ($additionalProperties.isTeamsConnected -eq $true)
                $title = $Site.displayName
                $url = $Site.webUrl
                $template = $additionalProperties.template
                $lastContentModifiedDate = $Site.lastModifiedDateTime
                $status = $additionalProperties.status
                $archiveStatus = $additionalProperties.archiveStatus
                $lockState = $additionalProperties.lockState
                $storageQuota = $additionalProperties.storageQuota
                $sharingCapability = $additionalProperties.sharingCapability
                $defaultLinkPermission = $additionalProperties.defaultLinkPermission
                $defaultSharingLinkType = $additionalProperties.defaultSharingLinkType
            }
            default {
                $storageUsageCurrent = [int]($Site.Usage.Storage ?? 0)
                if (-not $storageUsageCurrent -and $Site.Drive -and $Site.Drive.Quota) {
                    $storageUsageCurrent = [int]($Site.Drive.Quota.Used ?? 0)
                }

                $owner = $Site.Owner.UserPrincipalName
                if (-not $owner -and $Site.CreatedByUser) {
                    $owner = $Site.CreatedByUser.UserPrincipalName
                }
                if (-not $owner -and $Site.CreatedBy -and $Site.CreatedBy.User) {
                    $owner = $Site.CreatedBy.User.UserPrincipalName
                }

                $groupId = $Site.GroupId
                if (-not $groupId -and $Site.SharepointIds -and $Site.SharepointIds.SiteId) {
                    $groupId = $null
                }

                $isTeamsChannelConnected = ($Site.AdditionalProperties.IsTeamsChannelConnected -eq $true)
                $isTeamsConnected = ($Site.AdditionalProperties.IsTeamsConnected -eq $true) -or [bool]$groupId
                $title = if ($Site.DisplayName) { $Site.DisplayName } else { $Site.Name }
                $url = $Site.WebUrl
                $template = $Site.AdditionalProperties.Template
                $lastContentModifiedDate = $Site.LastModifiedDateTime
                $status = $Site.AdditionalProperties.Status
                $archiveStatus = $Site.AdditionalProperties.ArchiveStatus
                $lockState = $Site.AdditionalProperties.LockState
                $storageQuota = $Site.AdditionalProperties.StorageQuota
                $sharingCapability = $Site.AdditionalProperties.SharingCapability
                $defaultLinkPermission = $Site.AdditionalProperties.DefaultLinkPermission
                $defaultSharingLinkType = $Site.AdditionalProperties.DefaultSharingLinkType
            }
        }

        if ($UsageReport) {
            if ($UsageReport.'Storage Used (Byte)') {
                $storageUsageCurrent = [int64]$UsageReport.'Storage Used (Byte)'
            }
            if ($UsageReport.'Storage Allocated (Byte)') {
                $storageQuota = [int64]$UsageReport.'Storage Allocated (Byte)'
            }
            if ($UsageReport.'Owner Principal Name') {
                $owner = $UsageReport.'Owner Principal Name'
            }
            if (-not $lastContentModifiedDate -and $UsageReport.'Last Activity Date') {
                $lastContentModifiedDate = $UsageReport.'Last Activity Date'
            }
            if (-not $template -and $UsageReport.'Root Web Template') {
                $template = switch ($UsageReport.'Root Web Template') {
                    'Group' { 'GROUP#0' }
                    'TeamChannel' { 'TEAMCHANNEL#0' }
                    default { $UsageReport.'Root Web Template' }
                }
            }
        }

        if ($UsageReport -and $UsageReport.'Storage Used (Byte)') {
            # Graph usage reports return bytes.
            $storageUsageCurrentGB = [double]$storageUsageCurrent / 1GB
        }
        elseif ($Source -eq 'SPO') {
            # SPO cmdlets return storage in MB.
            $storageUsageCurrentGB = [double]$storageUsageCurrent / 1024
        }
        else {
            # Graph site usage and drive quota values are bytes.
            $storageUsageCurrentGB = [double]$storageUsageCurrent / 1GB
        }

        if (-not $template) {
            if ($IsOneDrive) {
                $template = 'SPSPERS#10'
            }
            elseif ($isTeamsChannelConnected) {
                $template = 'TEAMCHANNEL#0'
            }
            elseif ($groupId -and $groupId -ne '00000000-0000-0000-0000-000000000000') {
                $template = 'GROUP#0'
            }
            else {
                $template = 'STS#3'
            }
        }

        if (-not $url) {
            $url = if ($Site.WebUrl) { $Site.WebUrl } else { $Site.webUrl }
        }

        if (-not $title) {
            $title = if ($Site.DisplayName) { $Site.DisplayName } else { $url }
        }

        if (-not $owner -and $IsOneDrive -and $url) {
            $owner = Get-LikelyOneDriveOwnerFromUrl -Url $url
        }

        return [PSCustomObject]@{
            SiteId                    = $(if ($Site.PSObject.Properties['Id']) { [string]$Site.Id } elseif ($Site.PSObject.Properties['id']) { [string]$Site.id } else { $null })
            Template                  = $template
            IsHubSite                 = ($Site.IsHubSite -eq $true)
            Title                     = $title
            LastContentModifiedDate   = $lastContentModifiedDate
            Status                    = $status
            ArchiveStatus             = $archiveStatus
            StorageUsageCurrent       = $storageUsageCurrent
            LockState                 = $lockState
            Url                       = $url
            WebUrl                    = $url
            Owner                     = $owner
            StorageQuota              = $storageQuota
            GroupId                   = $groupId
            IsTeamsConnected          = $isTeamsConnected
            IsTeamsChannelConnected   = $isTeamsChannelConnected
            StorageUsedGB             = [math]::Round($storageUsageCurrentGB, 3)
            IsOffice365GroupsConnected = ($groupId -and $groupId -ne '00000000-0000-0000-0000-000000000000')
            IsOneDrive                = $IsOneDrive
            SharingCapability         = $sharingCapability
            DefaultLinkPermission     = $defaultLinkPermission
            DefaultSharingLinkType    = $defaultSharingLinkType
        }
    }

    $sharePointUsageBySiteId = @{}
    $oneDriveUsageBySiteId = @{}
    $sharePointUsageByUrl = @{}
    $oneDriveUsageByUrl = @{}
    if ($ServiceName -in @('MGGraph', 'API')) {
        $sharePointUsageBySiteId = Get-GraphCsvReportLookup -Uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D7')" -LookupName 'SharePointSiteUsageDetail' -UrlLookup $sharePointUsageByUrl
        $oneDriveUsageBySiteId = Get-GraphCsvReportLookup -Uri "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D7')" -LookupName 'OneDriveUsageAccountDetail' -UrlLookup $oneDriveUsageByUrl
    }

    # Option 1: Use Microsoft Graph PowerShell SDK
    function Get-SharePointAndOneDriveSitesFromGraphSdk {
        $totalCount = 0
        try {
            Write-Verbose "Fetching SharePoint and OneDrive sites using Microsoft Graph SDK"
            Write-Progress -Id $graphSitesProgressId -Activity "Gather all SharePoint Online Sites with OneDrives" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
            $savedProgressPreference = $ProgressPreference
            try {
                $ProgressPreference = 'SilentlyContinue'
                $graphSiteProperties = 'id,displayName,name,webUrl,lastModifiedDateTime,siteCollection,sharepointIds'
                foreach ($site in (Get-MgSite -All -Property $graphSiteProperties -ProgressAction SilentlyContinue -ErrorAction Stop)) {
                    $totalCount++
                    $isOneDrive = ($site.WebUrl -like "*-my.sharepoint.com*")
                    $reportSiteId = Get-GraphSiteReportId -CompositeSiteId $site.Id
                    $siteUrlKey = if ($site.WebUrl) { $site.WebUrl.TrimEnd('/').ToLowerInvariant() } else { $null }
                    $usageReport = if ($isOneDrive) { $oneDriveUsageBySiteId[$reportSiteId] } else { $sharePointUsageBySiteId[$reportSiteId] }
                    if (-not $usageReport -and $siteUrlKey) {
                        $usageReport = if ($isOneDrive) { $oneDriveUsageByUrl[$siteUrlKey] } else { $sharePointUsageByUrl[$siteUrlKey] }
                    }
                    Update-SiteUsageCoverage -IsOneDrive:$isOneDrive -UsageReport $usageReport

                    if ($totalCount -eq 1 -or ($totalCount % 50) -eq 0) {
                        Write-Progress -Id $siteDetailsProgressId -Activity "Gather Additional Site Details" -Status "Processed $totalCount site(s): $($site.DisplayName)"
                    }
                    $siteData = ConvertTo-NormalizedSiteData -Site $site -IsOneDrive:$isOneDrive -Source MGGraph -UsageReport $usageReport

                    if ($isOneDrive) {
                        $oneDriveKey = ConvertTo-OneDriveSiteKey -Url $siteData.Url -SiteId $siteData.SiteId
                        if ($oneDriveKey) {
                            $script:tenantStatsHash['OneDrive'][$oneDriveKey] = $siteData
                        }
                    } else {
                        $script:tenantStatsHash['SharePoint'][$siteData.Url] = $siteData
                    }
                }
            }
            finally {
                $ProgressPreference = $savedProgressPreference
            }
        } catch {
            Write-Error "Error fetching SharePoint and OneDrive sites with Graph SDK: $($_.Exception.Message)"
        } finally {
            Write-Progress -Id $siteDetailsProgressId -Activity "Gather Additional Site Details" -Completed
            Write-Progress -Id $graphSitesProgressId -Activity "Gather all SharePoint Online Sites with OneDrives" -Completed
        }
    }

    # Option 2: Use SharePoint Online PowerShell Module
    function Get-SharePointAndOneDriveSitesFromSPO {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $True, HelpMessage = 'Provide the level of detail')]
            [ValidateSet('minimum', 'combined', 'all', 'geek')]
            [string]$detailLevel
        )
        $totalCount = 1
        try {
            Write-Log -Type INFO -Message "[Get-SharePointAndOneDriveSitesFromSPO] START: Gathering all SharePoint Online Sites with OneDrives with $($detailLevel) details" -ExportFileLocation $ExportDetails
            Write-Progress -Id $spoSitesProgressId -Activity "Gather all SharePoint Online Sites with OneDrives" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
            switch ($detailLevel) {
                geek { $sites = Get-SPOSite -IncludePersonalSite $True -Limit All }
                Default {  
                    $DesiredProperties = @("Template", "IsHubSite", "Title", "LastContentModifiedDate", "Status", "ArchiveStatus", "StorageUsageCurrent", "LockState", "Url", "Owner", "StorageQuota", "GroupId", "IsTeamsConnected", "IsTeamsChannelConnected", "SharingCapability", "DefaultLinkPermission", "DefaultSharingLinkType")
                    $sites = Get-SPOSite -IncludePersonalSite $True -Limit All | Select-Object -Property $DesiredProperties
                }
            }
            
            $totalCount = ($sites | Measure-Object).count
            foreach ($site in $sites) {
                Write-ProgressHelper -Total $totalCount -Id $siteDetailsProgressId -Activity "Gather Additional Site Details" -Operation "Gathering Site Details for $($site.Title)"
                # Determine if the site is a OneDrive or a standard SharePoint site
                $isOneDrive = ($site.Url -like "*-my.sharepoint.com*")
                $siteData = ConvertTo-NormalizedSiteData -Site $site -IsOneDrive:$isOneDrive -Source SPO

                # Store data in appropriate hashtable
                if ($isOneDrive) {
                    $oneDriveKey = ConvertTo-OneDriveSiteKey -Url $siteData.Url -SiteId $siteData.SiteId
                    if ($oneDriveKey) {
                        $script:tenantStatsHash['OneDrive'][$oneDriveKey] = $siteData
                    }
                } else {
                    $script:tenantStatsHash['SharePoint'][$siteData.Url] = $siteData
                }
            }
        } catch {
            Write-Log -Type ERROR -Message "[Get-SharePointAndOneDriveSitesFromSPO] An error occurred in running Get-SharePointAndOneDriveSitesFromSPO function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        } finally {
            Write-ProgressHelper -Total ([Math]::Max($totalCount, 1)) -Id $siteDetailsProgressId -Activity "Gather Additional Site Details" -Completed
            Write-Progress -Id $spoSitesProgressId -Activity "Gather all SharePoint Online Sites with OneDrives" -Completed
        }
    }

    # Option 3: Use REST API
    function Get-SharePointAndOneDriveSitesFromRESTAPI {
        $allSitesUri = "https://graph.microsoft.com/v1.0/sites/getAllSites?`$top=200"
        $pageCount = 0
        $useSdkForPaging = (-not $global:GraphHeaders) -and (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) -and (Get-MgContext -ErrorAction SilentlyContinue)

        function Get-SitePageResponse {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$Uri
            )

            if ($useSdkForPaging) {
                $savedProgressPreference = $ProgressPreference
                try {
                    $ProgressPreference = 'SilentlyContinue'
                    return Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputType PSObject -ProgressAction SilentlyContinue -ErrorAction Stop
                }
                finally {
                    $ProgressPreference = $savedProgressPreference
                }
            }

            return Invoke-QuietRestMethod -Parameters @{
                Uri         = $Uri
                Headers     = $global:GraphHeaders
                Method      = 'Get'
                ContentType = 'application/json'
                ErrorAction = 'Stop'
            }
        }
        
        try {
            Write-Verbose "Fetching initial SharePoint and OneDrive sites via Graph getAllSites API"
            $pageCount++
            $response = Get-SitePageResponse -Uri $allSitesUri
            Write-Progress -Activity "Fetching Sites" -Id $restSitesProgressId -Status "Processing page $pageCount"

            while ($true) {
                foreach ($site in @($response.value)) {
                    $isOneDrive = ($site.webUrl -like "*-my.sharepoint.com*")
                    $reportSiteId = if ($site.id) { Get-GraphSiteReportId -CompositeSiteId $site.id } else { $null }
                    $siteUrlKey = if ($site.webUrl) { $site.webUrl.TrimEnd('/').ToLowerInvariant() } else { $null }
                    $usageReport = if ($isOneDrive) { $oneDriveUsageBySiteId[$reportSiteId] } else { $sharePointUsageBySiteId[$reportSiteId] }
                    if (-not $usageReport -and $siteUrlKey) {
                        $usageReport = if ($isOneDrive) { $oneDriveUsageByUrl[$siteUrlKey] } else { $sharePointUsageByUrl[$siteUrlKey] }
                    }
                    Update-SiteUsageCoverage -IsOneDrive:$isOneDrive -UsageReport $usageReport

                    $siteData = ConvertTo-NormalizedSiteData -Site $site -IsOneDrive:$isOneDrive -Source API -UsageReport $usageReport

                    if ($isOneDrive) {
                        $oneDriveKey = ConvertTo-OneDriveSiteKey -Url $siteData.Url -SiteId $siteData.SiteId
                        if ($oneDriveKey) {
                            $script:tenantStatsHash['OneDrive'][$oneDriveKey] = $siteData
                        }
                    } else {
                        $script:tenantStatsHash['SharePoint'][$siteData.Url] = $siteData
                    }
                }

                $allSitesUri = $response.'@odata.nextLink'
                if (-not $allSitesUri) {
                    break
                }

                $pageCount++
                Write-Verbose "Fetching next page of sites via REST API. Page $pageCount"
                Write-Progress -Activity "Fetching Sites" -Id $restSitesProgressId -Status "Processing page $pageCount"
                $response = Get-SitePageResponse -Uri $allSitesUri
            }
        } catch {
            $statusCode = $null
            if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response) {
                $statusCode = $_.Exception.Response.StatusCode
            }
            if ($statusCode -eq 429) {
                Write-Host "Throttling detected. Please try again later." -ForegroundColor Yellow
            } else {
                Write-Host "Error fetching sites via REST API: $($_.Exception.Message)" -ForegroundColor Red
            }
        } finally {
            Write-Progress -Activity "Fetching Sites" -Id $restSitesProgressId -Completed
        }
    }

# Determine which method to use based on ServiceName
    switch ($ServiceName) {
        'MGGraph' { Get-SharePointAndOneDriveSitesFromGraphSdk }
        'SPO'     { Get-SharePointAndOneDriveSitesFromSPO -detailLevel $detailLevel }
        'API'     { Get-SharePointAndOneDriveSitesFromRESTAPI }
    }
    if ($ServiceName -in @('MGGraph', 'API')) {
        $sharePointUsageRows = if ($sharePointUsageBySiteId) { $sharePointUsageBySiteId.Count } else { 0 }
        $oneDriveUsageRows = if ($oneDriveUsageBySiteId) { $oneDriveUsageBySiteId.Count } else { 0 }
        Write-Host ("  Site usage report coverage: SharePoint {0}/{1}, OneDrive {2}/{3}" -f $siteUsageCoverage.SharePointMatched, $siteUsageCoverage.SharePointTotal, $siteUsageCoverage.OneDriveMatched, $siteUsageCoverage.OneDriveTotal) -ForegroundColor DarkGray
        Write-Log -Type INFO -Message "[Get-SharePointAndOneDriveSites] Site usage report coverage: SharePointMatched=$($siteUsageCoverage.SharePointMatched); SharePointMissing=$($siteUsageCoverage.SharePointMissing); SharePointTotal=$($siteUsageCoverage.SharePointTotal); OneDriveMatched=$($siteUsageCoverage.OneDriveMatched); OneDriveMissing=$($siteUsageCoverage.OneDriveMissing); OneDriveTotal=$($siteUsageCoverage.OneDriveTotal); SharePointReportRows=$sharePointUsageRows; OneDriveReportRows=$oneDriveUsageRows." -ExportFileLocation $ExportDetails
    }
    if ($sharePointUsageBySiteId) { $sharePointUsageBySiteId.Clear() }
    if ($oneDriveUsageBySiteId) { $oneDriveUsageBySiteId.Clear() }
    if ($sharePointUsageByUrl) { $sharePointUsageByUrl.Clear() }
    if ($oneDriveUsageByUrl) { $oneDriveUsageByUrl.Clear() }
    $sharePointUsageBySiteId = $null
    $oneDriveUsageBySiteId = $null
    $sharePointUsageByUrl = $null
    $oneDriveUsageByUrl = $null
    if ($ServiceName -in @('MGGraph', 'API')) {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
   
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type Info -Message "[Get-SharePointAndOneDriveSites] COMPLETED: Gathering all SharePoint and OneDrive Details ($($ServiceName)) in $($CompletedTime)" -ExportFileLocation $ExportDetails
}

# Get Teams inventory and channel posture (Graph-first with Teams PS fallback)

function Get-TeamsDetails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = 'Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory = $false, HelpMessage = 'Provide the service name')]
        [ValidateSet('MGGraph', 'Teams')]
        [string]$ServiceName = 'MGGraph'
    )

    $start = Get-Date
    $teamsProgressId = 43
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash['AllTeams'] = @{}

    Write-Host "Getting all Microsoft Teams details ..." -ForegroundColor Cyan -NoNewline
    Write-Log -Type INFO -Message "[Get-TeamsDetails] START: Gathering Teams inventory with $detailLevel details" -ExportFileLocation $ExportDetails

    try {
        $allTeams = @()
        $selectedSource = 'Unavailable'
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        $canUseGraphSdk = ($ServiceName -eq 'MGGraph') -and $graphContext -and (Get-Command -Name 'Get-MgTeam' -ErrorAction SilentlyContinue)
        $canUseGraphRest = ($ServiceName -eq 'MGGraph') -and $global:GraphHeaders
        $teamsConnected = $false
        $scriptConnectionResult = Get-Variable -Name connectionResult -Scope Script -ErrorAction SilentlyContinue
        if ($scriptConnectionResult -and $scriptConnectionResult.Value) {
            $teamsConnected = [bool]$scriptConnectionResult.Value.Teams
        }

        if ($canUseGraphSdk) {
            $allTeams = @(Invoke-QuietCommand -ScriptBlock { Get-MgTeam -All -ErrorAction Stop })
            $selectedSource = 'MGGraph-SDK'
        }
        elseif ($canUseGraphRest) {
            $allTeams = @(Get-ArrayaGraphResource -Uri "https://graph.microsoft.com/v1.0/teams?`$select=id,displayName,description,visibility,isArchived" -Activity 'Teams inventory (Graph REST)' -PreferRest -Headers $global:GraphHeaders)
            $selectedSource = 'MGGraph-REST'
        }
        elseif ($teamsConnected -and (Get-Command -Name 'Get-Team' -ErrorAction SilentlyContinue)) {
            $previousProgressPreference = $ProgressPreference
            try {
                $ProgressPreference = 'SilentlyContinue'
                $allTeams = @(Invoke-QuietCommand -ScriptBlock { Get-Team -ErrorAction Stop })
            }
            finally {
                $ProgressPreference = $previousProgressPreference
            }
            $selectedSource = 'TeamsPowerShell'
        }
        else {
            throw "No Teams collection source is available. Ensure Graph app permissions for Teams or an active Teams PowerShell session."
        }

        $allTeams = @($allTeams | Where-Object { $null -ne $_ })
        if ($allTeams.Count -eq 0) {
            Write-Log -Type INFO -Message "[Get-TeamsDetails] No teams returned from source $selectedSource." -ExportFileLocation $ExportDetails
            return
        }

        $sharePointRows = if ($script:tenantStatsHash.ContainsKey('SharePoint') -and $script:tenantStatsHash['SharePoint'] -is [System.Collections.IDictionary]) {
            @($script:tenantStatsHash['SharePoint'].Values)
        } else {
            @()
        }
        $unifiedGroupsById = @{}
        if ($script:tenantStatsHash.ContainsKey('UnifiedGroups') -and $script:tenantStatsHash['UnifiedGroups'] -is [System.Collections.IDictionary]) {
            foreach ($group in @($script:tenantStatsHash['UnifiedGroups'].Values)) {
                if ($group -and $group.PSObject.Properties['ExternalDirectoryObjectId'] -and -not [string]::IsNullOrWhiteSpace([string]$group.ExternalDirectoryObjectId)) {
                    $unifiedGroupsById[[string]$group.ExternalDirectoryObjectId] = $group
                }
            }
        }
        $groupsActivityById = @{}
        if ($script:tenantStatsHash.ContainsKey('Office365GroupsActivityTopGroups') -and $script:tenantStatsHash['Office365GroupsActivityTopGroups'] -is [System.Collections.IDictionary]) {
            foreach ($groupActivity in @($script:tenantStatsHash['Office365GroupsActivityTopGroups'].Values)) {
                if ($groupActivity -and $groupActivity.PSObject.Properties['GroupId'] -and -not [string]::IsNullOrWhiteSpace([string]$groupActivity.GroupId)) {
                    $groupsActivityById[[string]$groupActivity.GroupId] = $groupActivity
                }
            }
        }

        $collectChannelDetails = ($detailLevel -ne 'minimum')
        $totalCount = [Math]::Max($allTeams.Count, 1)
        $index = 0
        $collectedCount = 0

        foreach ($team in $allTeams) {
            $index++
            $teamId = if ($team.PSObject.Properties['Id'] -and $team.Id) { [string]$team.Id } elseif ($team.PSObject.Properties['GroupId'] -and $team.GroupId) { [string]$team.GroupId } else { $null }
            $displayName = if ($team.PSObject.Properties['DisplayName'] -and $team.DisplayName) { [string]$team.DisplayName } else { if ($teamId) { $teamId } else { "Team-$index" } }
            Write-ProgressHelper -Total $totalCount -Id $teamsProgressId -Index $index -Activity "Gathering Teams inventory ($detailLevel)" -Operation $displayName

            $currentTeam = Invoke-ArrayaCollectionStepSafe -OperationName "Get-TeamsDetails [$displayName]" -DefaultValue $null -ExportFileLocation $ExportDetails -ScriptBlock {
                $spoSiteDetails = $null
                if ($sharePointRows.Count -gt 0) {
                    if (-not [string]::IsNullOrWhiteSpace($teamId)) {
                        $spoSiteDetails = $sharePointRows | Where-Object { $_.GroupId -and ([string]$_.GroupId -eq $teamId) } | Select-Object -First 1
                    }
                    if (-not $spoSiteDetails) {
                        $spoSiteDetails = $sharePointRows | Where-Object { $_.Title -and ([string]$_.Title -eq $displayName) } | Select-Object -First 1
                    }
                }

                $siteSizeGB = 0
                if ($spoSiteDetails) {
                    if ($spoSiteDetails.PSObject.Properties['StorageUsedGB'] -and $spoSiteDetails.StorageUsedGB) {
                        try { $siteSizeGB = [math]::Round([double]$spoSiteDetails.StorageUsedGB, 3) } catch { $siteSizeGB = 0 }
                    }
                    elseif ($spoSiteDetails.PSObject.Properties['StorageUsageCurrent'] -and $spoSiteDetails.StorageUsageCurrent) {
                        try { $siteSizeGB = [math]::Round(([double]$spoSiteDetails.StorageUsageCurrent / 1024), 3) } catch { $siteSizeGB = 0 }
                    }
                }
                $siteSizeMb = [math]::Round(($siteSizeGB * 1024), 3)

                $publicChannels = @()
                $privateChannels = @()
                $sharedChannels = @()
                $memberCount = $null
                $guestCount = $null
                $lastActivityDate = if ($spoSiteDetails -and $spoSiteDetails.PSObject.Properties['LastContentModifiedDate']) { $spoSiteDetails.LastContentModifiedDate } else { $null }
                if (-not $lastActivityDate -and -not [string]::IsNullOrWhiteSpace($teamId) -and $groupsActivityById.ContainsKey($teamId)) {
                    $lastActivityDate = $groupsActivityById[$teamId].LastActivityDate
                }

                if ($collectChannelDetails -and -not [string]::IsNullOrWhiteSpace($teamId)) {
                    $channels = @()
                    if ($selectedSource -eq 'MGGraph-SDK') {
                        if (Get-Command -Name 'Get-MgTeamAllChannel' -ErrorAction SilentlyContinue) {
                            $channels = @(Get-MgTeamAllChannel -TeamId $teamId -ProgressAction SilentlyContinue -ErrorAction SilentlyContinue)
                        }
                        else {
                            $channels = @(Get-MgTeamChannel -TeamId $teamId -ProgressAction SilentlyContinue -ErrorAction SilentlyContinue)
                        }
                    }
                    elseif ($selectedSource -eq 'MGGraph-REST') {
                        $channels = @(Get-ArrayaGraphResource -Uri ("https://graph.microsoft.com/v1.0/teams/{0}/allChannels?`$select=displayName,membershipType" -f $teamId) -Activity "Teams channels [$displayName]" -PreferRest -Headers $global:GraphHeaders)
                    }
                    elseif ($selectedSource -eq 'TeamsPowerShell') {
                        $channels = @(Get-TeamChannel -GroupId $teamId -ErrorAction SilentlyContinue)
                    }

                    foreach ($channel in $channels) {
                        if ($null -eq $channel) { continue }
                        $channelName = if ($channel.PSObject.Properties['DisplayName']) { [string]$channel.DisplayName } else { [string]$channel }
                        if ([string]::IsNullOrWhiteSpace($channelName)) { continue }
                        $membershipType = 'standard'
                        if ($channel.PSObject.Properties['MembershipType'] -and $channel.MembershipType) {
                            $membershipType = [string]$channel.MembershipType
                        }

                        switch ($membershipType.ToLowerInvariant()) {
                            'private' { $privateChannels += $channelName }
                            'shared' { $sharedChannels += $channelName }
                            default { $publicChannels += $channelName }
                        }
                    }
                }

                if ($detailLevel -ne 'minimum' -and -not [string]::IsNullOrWhiteSpace($teamId)) {
                    $memberObjects = @()
                    if ($selectedSource -eq 'MGGraph-SDK') {
                        if (Get-Command -Name 'Get-MgTeamMember' -ErrorAction SilentlyContinue) {
                            $memberObjects = @(Get-MgTeamMember -TeamId $teamId -All -ProgressAction SilentlyContinue -ErrorAction SilentlyContinue)
                        }
                    }
                    elseif ($selectedSource -eq 'MGGraph-REST') {
                        $memberObjects = @(Get-ArrayaGraphResource -Uri ("https://graph.microsoft.com/v1.0/teams/{0}/members?`$select=id,email,userId,roles" -f $teamId) -Activity "Teams members [$displayName]" -PreferRest -Headers $global:GraphHeaders)
                    }
                    elseif ($selectedSource -eq 'TeamsPowerShell' -and (Get-Command -Name 'Get-TeamUser' -ErrorAction SilentlyContinue)) {
                        $memberObjects = @(Get-TeamUser -GroupId $teamId -ErrorAction SilentlyContinue)
                    }

                    if ($memberObjects.Count -gt 0) {
                        $memberCount = $memberObjects.Count
                        $guestCount = @(
                            $memberObjects | Where-Object {
                                $odataType = if ($_.AdditionalProperties -and $_.AdditionalProperties.ContainsKey('@odata.type')) { [string]$_.AdditionalProperties['@odata.type'] } else { '' }
                                $emailValue = if ($_.PSObject.Properties['Email']) { [string]$_.Email } elseif ($_.AdditionalProperties -and $_.AdditionalProperties.ContainsKey('email')) { [string]$_.AdditionalProperties['email'] } else { '' }
                                $userValue = if ($_.PSObject.Properties['User']) { [string]$_.User } else { '' }
                                $upnValue = if ($_.AdditionalProperties -and $_.AdditionalProperties.ContainsKey('userPrincipalName')) { [string]$_.AdditionalProperties['userPrincipalName'] } else { '' }
                                ($odataType -match 'aadUserConversationMember') -and (($emailValue -like '*#EXT#*') -or ($userValue -like '*#EXT#*') -or ($upnValue -like '*#EXT#*'))
                            }
                        ).Count
                    }
                }

                $ownerCount = $null
                if (-not [string]::IsNullOrWhiteSpace($teamId) -and $unifiedGroupsById.ContainsKey($teamId)) {
                    $groupRecord = $unifiedGroupsById[$teamId]
                    $managedByDetails = if ($groupRecord.PSObject.Properties['ManagedByDetails']) { [string]$groupRecord.ManagedByDetails } else { $null }
                    if ($null -ne $managedByDetails) {
                        $ownerCount = @(
                            $managedByDetails -split ',' |
                                ForEach-Object { $_.Trim() } |
                                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                                Select-Object -Unique
                        ).Count
                    }
                }

                [PSCustomObject]@{
                    TeamId              = $teamId
                    DisplayName         = $displayName
                    Description         = if ($team.PSObject.Properties['Description']) { [string]$team.Description } else { $null }
                    Visibility          = if ($team.PSObject.Properties['Visibility']) { [string]$team.Visibility } else { 'Unknown' }
                    IsArchived          = if ($team.PSObject.Properties['IsArchived']) { [bool]$team.IsArchived } else { $false }
                    SharePointSiteUrl   = if ($spoSiteDetails -and $spoSiteDetails.PSObject.Properties['Url']) { [string]$spoSiteDetails.Url } else { $null }
                    'SiteSize-GB'       = $siteSizeGB
                    SiteSize            = $siteSizeMb
                    TotalChannels       = $publicChannels.Count + $privateChannels.Count + $sharedChannels.Count
                    PublicChannels      = ($publicChannels -join ',')
                    PrivateChannels     = ($privateChannels -join ',')
                    SharedChannels      = ($sharedChannels -join ',')
                    PublicChannelCount  = [int]$publicChannels.Count
                    PrivateChannelCount = [int]$privateChannels.Count
                    SharedChannelCount  = [int]$sharedChannels.Count
                    OwnerCount          = $ownerCount
                    OwnershipState      = if ($null -eq $ownerCount) { 'Unknown' } elseif ($ownerCount -eq 0) { 'Unowned' } else { 'Owned' }
                    MemberCount         = $memberCount
                    GuestCount          = $guestCount
                    LastActivityDate    = $lastActivityDate
                    DataSource          = $selectedSource
                }
            }

            if ($currentTeam) {
                $teamKeyBase = if (-not [string]::IsNullOrWhiteSpace($currentTeam.TeamId)) { [string]$currentTeam.TeamId } else { [string]$currentTeam.DisplayName }
                if ([string]::IsNullOrWhiteSpace($teamKeyBase)) {
                    $teamKeyBase = "team-$index"
                }
                $teamKey = $teamKeyBase
                $suffix = 2
                while ($script:tenantStatsHash['AllTeams'].ContainsKey($teamKey)) {
                    $teamKey = "{0}#{1}" -f $teamKeyBase, $suffix
                    $suffix++
                }
                $script:tenantStatsHash['AllTeams'][$teamKey] = $currentTeam
                $collectedCount++
            }
        }

        Write-Log -Type INFO -Message "[Get-TeamsDetails] Teams inventory collected: source=$selectedSource; discovered=$($allTeams.Count); stored=$collectedCount; channelDetailsCollected=$collectChannelDetails." -ExportFileLocation $ExportDetails
    }
    catch {
        Write-Log -Type WARNING -Message "[Get-TeamsDetails] Teams collection failed: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }
    finally {
        $completedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-ProgressHelper -Total 1 -Id $teamsProgressId -Activity "Gathering Teams inventory ($detailLevel)" -Completed
        Write-Host "Completed in $completedTime" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-TeamsDetails] COMPLETED in $completedTime" -ExportFileLocation $ExportDetails
    }
}

function Get-TeamsVoiceDetails {
    [CmdletBinding()]
    param()

    $start = Get-Date
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["TeamsVoice"] = @{}

    Write-Host "Gathering Teams Voice details ..." -ForegroundColor Cyan -NoNewline
    Write-Log -Type INFO -Message "[Get-TeamsVoiceDetails] START: Gathering Teams Voice details" -ExportFileLocation $ExportDetails

    $teamsConnected = $false
    $scriptConnectionResult = Get-Variable -Name connectionResult -Scope Script -ErrorAction SilentlyContinue
    if ($scriptConnectionResult -and $scriptConnectionResult.Value) {
        $teamsConnected = [bool]$scriptConnectionResult.Value.Teams
    }

    $pstnUsage = @()
    $callingPolicies = @()
    $phoneAssignments = @()
    $voiceUsers = @()
    $summary = [pscustomobject]@{
        PstnUsageDays = 30
        PstnTotalMinutes = 0
        PstnTotalCalls = 0
        CallingPolicyCount = 0
        PhoneNumberCount = 0
        VoiceUserCount = 0
        DataSource = 'TeamsPowerShell'
        Notes = $null
    }

    if (-not $teamsConnected) {
        $voicePlans = @('MCOEV','MCOPSTN1','MCOPSTN2','MCOEV_VIRTUALUSER','MCOEV_DOD','MCOPSTNC')
        $users = @()
        if ($script:tenantStatsHash.ContainsKey('Users')) {
            $users = @($script:tenantStatsHash['Users'].Values)
        }

        $voiceLicensedUsers = @(
            $users | Where-Object {
                $assignedLicenses = @()
                $enabledServicePlans = @()
                if ($_.PSObject.Properties['AssignedLicenses'] -and $_.AssignedLicenses) {
                    $assignedLicenses = @($_.AssignedLicenses -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                }
                if ($_.PSObject.Properties['EnabledServicePlans'] -and $_.EnabledServicePlans) {
                    $enabledServicePlans = @($_.EnabledServicePlans -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                }

                (@($assignedLicenses + $enabledServicePlans) | Where-Object { $voicePlans -contains $_ }).Count -gt 0
            }
        )

        $summary.PstnTotalMinutes = 'Unavailable'
        $summary.PstnTotalCalls = 'Unavailable'
        $summary.CallingPolicyCount = 'Unavailable'
        $summary.PhoneNumberCount = 'Unavailable'
        $summary.VoiceUserCount = $voiceLicensedUsers.Count
        $summary.DataSource = 'GraphLicenseInference'
        $summary.Notes = 'Teams PowerShell is not connected. Voice-enabled users are inferred from assigned voice licenses and enabled service plans.'

        $script:tenantStatsHash["TeamsVoice"]["Summary"] = $summary
        $script:tenantStatsHash["TeamsVoice"]["PstnUsage"] = $pstnUsage
        $script:tenantStatsHash["TeamsVoice"]["CallingPolicies"] = $callingPolicies
        $script:tenantStatsHash["TeamsVoice"]["PhoneNumbers"] = $phoneAssignments
        $script:tenantStatsHash["TeamsVoice"]["VoiceUsers"] = $voiceUsers

        $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
        Write-Host " Skipped in $elapsed" -ForegroundColor Yellow
        Write-Log -Type INFO -Message "[Get-TeamsVoiceDetails] Skipped because Teams PowerShell is not connected in the current session. Inferred $($voiceLicensedUsers.Count) voice-licensed users from Graph user licensing." -ExportFileLocation $ExportDetails
        Write-Log -Type INFO -Message "[Get-TeamsVoiceDetails] COMPLETED in $elapsed" -ExportFileLocation $ExportDetails
        return
    }

    try {
        if (Get-Command Get-CsOnlinePstnUsageReport -ErrorAction SilentlyContinue) {
            $endDate = Get-Date
            $startDate = $endDate.AddDays(-30)
            $pstnUsage = Get-CsOnlinePstnUsageReport -StartDate $startDate -EndDate $endDate -ErrorAction SilentlyContinue
            if ($pstnUsage) {
                $summary.PstnTotalMinutes = ($pstnUsage | Measure-Object -Property Minutes -Sum).Sum
                $summary.PstnTotalCalls = ($pstnUsage | Measure-Object -Property TotalCalls -Sum).Sum
            }
        }
    } catch {
        Write-Log -Type WARNING -Message "[Get-TeamsVoiceDetails] Unable to retrieve PSTN usage: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }

    try {
        if (Get-Command Get-CsTeamsCallingPolicy -ErrorAction SilentlyContinue) {
            $callingPolicies = Get-CsTeamsCallingPolicy -ErrorAction SilentlyContinue
        } elseif (Get-Command Get-CsCallingPolicy -ErrorAction SilentlyContinue) {
            $callingPolicies = Get-CsCallingPolicy -ErrorAction SilentlyContinue
        }
        if ($callingPolicies) {
            $summary.CallingPolicyCount = $callingPolicies.Count
        }
    } catch {
        Write-Log -Type WARNING -Message "[Get-TeamsVoiceDetails] Unable to retrieve calling policies: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }

    try {
        if (Get-Command Get-CsPhoneNumberAssignment -ErrorAction SilentlyContinue) {
            $phoneAssignments = Get-CsPhoneNumberAssignment -ErrorAction SilentlyContinue
        }
        if ($phoneAssignments) {
            $summary.PhoneNumberCount = $phoneAssignments.Count
        }
    } catch {
        Write-Log -Type WARNING -Message "[Get-TeamsVoiceDetails] Unable to retrieve phone number assignments: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }

    try {
        if (Get-Command Get-CsOnlineUser -ErrorAction SilentlyContinue) {
            $voiceUsers = Get-CsOnlineUser -WarningAction SilentlyContinue | Where-Object { $_.LineURI -and $_.LineURI -ne '' }
            if ($voiceUsers) {
                $summary.VoiceUserCount = $voiceUsers.Count
            }
        }
    } catch {
        Write-Log -Type WARNING -Message "[Get-TeamsVoiceDetails] Unable to retrieve voice-enabled users: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }

    $script:tenantStatsHash["TeamsVoice"]["Summary"] = $summary
    $script:tenantStatsHash["TeamsVoice"]["PstnUsage"] = $pstnUsage
    $script:tenantStatsHash["TeamsVoice"]["CallingPolicies"] = $callingPolicies
    $script:tenantStatsHash["TeamsVoice"]["PhoneNumbers"] = $phoneAssignments
    $script:tenantStatsHash["TeamsVoice"]["VoiceUsers"] = $voiceUsers

    $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
    Write-Host " Completed in $elapsed" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-TeamsVoiceDetails] COMPLETED in $elapsed" -ExportFileLocation $ExportDetails
}

# ----------------------------------
# Tenant Specific Functions
# ----------------------------------
$script:AssessmentTenantOrganization = $null
$script:AssessmentTenantOrganizationResolved = $false

function Get-AssessmentTenantOrganization {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$OrganizationId,
        [Parameter(Mandatory = $false)]
        [switch]$Refresh
    )

    if (-not $Refresh -and [string]::IsNullOrWhiteSpace($OrganizationId) -and $script:AssessmentTenantOrganizationResolved) {
        return $script:AssessmentTenantOrganization
    }

    $organizationUri = if ([string]::IsNullOrWhiteSpace($OrganizationId)) {
        'https://graph.microsoft.com/v1.0/organization'
    } else {
        "https://graph.microsoft.com/v1.0/organization/$OrganizationId"
    }

    $orgData = Get-ArrayaGraphResource -Uri $organizationUri -PageSize 50 -Activity 'Fetching tenant organization metadata' -Headers $global:GraphHeaders
    $organization = $null
    if ($orgData -is [array]) {
        $organization = $orgData | Select-Object -First 1
    }
    else {
        $organization = $orgData
    }

    if ([string]::IsNullOrWhiteSpace($OrganizationId)) {
        $script:AssessmentTenantOrganization = $organization
        $script:AssessmentTenantOrganizationResolved = $true
    }

    return $organization
}

function Initialize-MicrosoftLicenseReferenceMap {
    [CmdletBinding()]
    param()

    if ($script:MicrosoftProductNameMapFromReferenceInitialized -and $script:MicrosoftProductNameMapFromReference) {
        return $script:MicrosoftProductNameMapFromReference
    }

    $script:MicrosoftProductNameMapFromReference = @{}
    $script:MicrosoftProductNameMapFromReferenceInitialized = $true

    $knownCsvUri = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    $docsUri = 'https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference'
    $cacheFile = Join-Path -Path $env:TEMP -ChildPath 'arraya-license-service-plan-reference.csv'

    try {
        $useCachedCsv = $false
        if (Test-Path -Path $cacheFile) {
            $cacheAge = (Get-Date) - (Get-Item -Path $cacheFile).LastWriteTime
            if ($cacheAge.TotalDays -lt 7) {
                $useCachedCsv = $true
            }
        }

        if (-not $useCachedCsv) {
            $csvUri = $knownCsvUri
            try {
                $docsResponse = Invoke-WebRequest -Uri $docsUri -UseBasicParsing -ErrorAction Stop
                $pattern = "https://download\.microsoft\.com/download/[^\s""'<>]+Product%20names%20and%20service%20plan%20identifiers%20for%20licensing\.csv"
                $match = [regex]::Match($docsResponse.Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($match.Success -and -not [string]::IsNullOrWhiteSpace($match.Value)) {
                    $csvUri = $match.Value
                }
            }
            catch {
                $csvUri = $knownCsvUri
            }

            Invoke-WebRequest -Uri $csvUri -OutFile $cacheFile -UseBasicParsing -ErrorAction Stop
        }

        $csvRows = Import-Csv -Path $cacheFile -ErrorAction Stop
        foreach ($row in $csvRows) {
            $stringId = [string]$row.String_Id
            $displayName = [string]$row.Product_Display_Name
            if ([string]::IsNullOrWhiteSpace($stringId) -or [string]::IsNullOrWhiteSpace($displayName)) {
                continue
            }

            if (-not $script:MicrosoftProductNameMapFromReference.ContainsKey($stringId)) {
                $script:MicrosoftProductNameMapFromReference[$stringId] = $displayName
            }

            $normalizedStringId = ($stringId -replace '\s*_\s*', '_').Trim()
            if (-not [string]::IsNullOrWhiteSpace($normalizedStringId) -and -not $script:MicrosoftProductNameMapFromReference.ContainsKey($normalizedStringId)) {
                $script:MicrosoftProductNameMapFromReference[$normalizedStringId] = $displayName
            }
        }
    }
    catch {
        # Keep map empty on failure; caller falls back to static mappings and heuristics.
        $script:MicrosoftProductNameMapFromReference = @{}
    }

    return $script:MicrosoftProductNameMapFromReference
}

function Get-FriendlyProductName {
    param(
        [string]$SkuPartNumber
    )
    if ([string]::IsNullOrWhiteSpace($SkuPartNumber)) {
        return $SkuPartNumber
    }

    if (-not $script:MicrosoftProductNameMapFromReferenceInitialized) {
        Initialize-MicrosoftLicenseReferenceMap | Out-Null
    }

    # 1) Exact/raw lookup first.
    if ($script:CommonProductNameMapStatic) {
        if ($script:CommonProductNameMapStatic.ContainsKey($SkuPartNumber)) {
            return $script:CommonProductNameMapStatic[$SkuPartNumber]
        }
    }

    if ($script:MicrosoftProductNameMapFromReference) {
        if ($script:MicrosoftProductNameMapFromReference.ContainsKey($SkuPartNumber)) {
            return $script:MicrosoftProductNameMapFromReference[$SkuPartNumber]
        }
    }

    # 2) If exact lookup misses, normalize and retry.
    $normalizedSkuPartNumber = ($SkuPartNumber -replace '\s*_\s*', '_').Trim()
    if (
        -not [string]::IsNullOrWhiteSpace($normalizedSkuPartNumber) -and
        $normalizedSkuPartNumber -ne $SkuPartNumber
    ) {
        if ($script:CommonProductNameMapStatic -and $script:CommonProductNameMapStatic.ContainsKey($normalizedSkuPartNumber)) {
            return $script:CommonProductNameMapStatic[$normalizedSkuPartNumber]
        }
        if ($script:MicrosoftProductNameMapFromReference -and $script:MicrosoftProductNameMapFromReference.ContainsKey($normalizedSkuPartNumber)) {
            return $script:MicrosoftProductNameMapFromReference[$normalizedSkuPartNumber]
        }
    }

    switch -Regex ($normalizedSkuPartNumber.ToUpperInvariant()) {
        '^SPE_E([35])$' { return "Microsoft 365 E$($Matches[1])" }
        '^SPE_F([13])$' { return "Microsoft 365 F$($Matches[1])" }
        '^MCOPSTN1$' { return 'Microsoft Teams Phone Standard' }
        '^MCOCAP$' { return 'Microsoft Teams Shared Space' }
        '^TEAMS_SHARED_SPACE$' { return 'Microsoft Teams Shared Space' }
        '^MICROSOFT_TEAMS_ENTERPRISE_NEW$' { return 'Microsoft Teams Enterprise' }
        '^POWERAPPS_PER_USER$' { return 'Power Apps Premium' }
        '^EXCHANGEENTERPRISE$' { return 'Exchange Online (Plan 2)' }
        '^CPC_E_(\d+)C_(\d+)GB_(\d+)GB$' { return "Windows 365 Enterprise $($Matches[1]) vCPU $($Matches[2]) GB $($Matches[3]) GB" }
        '^WINDOWS_365_S_(\d+)VCPU_(\d+)GB_(\d+)GB$' { return "Windows 365 Shared Use $($Matches[1]) vCPU $($Matches[2]) GB $($Matches[3]) GB" }
        '^CPC_LVL_1$' { return 'Windows 365 Enterprise 2 vCPU 4 GB 128 GB (Preview)' }
        '^CPC_LVL_3$' { return 'Windows 365 Enterprise 4 vCPU 16 GB 256 GB (Preview)' }
        '^WIN10_PRO_ENT_SUB$' { return 'Windows 10/11 Enterprise E3' }
        '^WIN10_VDA_E3$' { return 'Windows 10/11 Enterprise VDA E3' }
        '^WIN10_VDA_E5$' { return 'Windows 10/11 Enterprise VDA E5' }
        '^WINE5_GCC_COMPAT$' { return 'Windows 10/11 Enterprise E5 Commercial (GCC Compatible)' }
        '^E3_VDA_ONLY$' { return 'Windows 10/11 Enterprise VDA E3 (VDA only)' }
    }

    $friendly = $normalizedSkuPartNumber -replace '[_\-]+', ' '
    $friendly = $friendly -replace '\bM365\b', 'Microsoft 365'
    $friendly = $friendly -replace '\bO365\b', 'Office 365'
    $friendly = $friendly -replace '\bEXCHANGE\b', 'Exchange'
    $friendly = $friendly -replace '\bPOWERAPPS\b', 'Power Apps'

    $tokens = $friendly -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
        if ($_ -match '^[A-Z0-9]{2,}$' -or $_ -match '^[EF]\d$') {
            $_
        }
        else {
            (Get-Culture).TextInfo.ToTitleCase($_.ToLowerInvariant())
        }
    }

    $friendly = ($tokens -join ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($friendly)) {
        return $normalizedSkuPartNumber
    }

    return $friendly
}

function Get-LicenseClassification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$License,
        [int]$UserCount = 0
    )

    $skuPartNumber = [string]$License.SkuPartNumber
    $appliesTo = [string]$License.AppliesTo
    $isTrialFlag = [bool]($License.PSObject.Properties['IsTrial'] -and $License.IsTrial -eq $true)
    $isFreeOrTrialFlag = [bool]($License.PSObject.Properties['IsFreeOrTrial'] -and $License.IsFreeOrTrial -eq $true)
    $ignoreLifecycle = [bool]($License.PSObject.Properties['IgnoreLifecycle'] -and $License.IgnoreLifecycle -eq $true)
    $purchased = 0
    $consumed = 0

    if ($License.PSObject.Properties['PurchasedUnits'] -and $null -ne $License.PurchasedUnits -and $License.PurchasedUnits -ne 'N/A' -and $License.PurchasedUnits -ne '') {
        try { $purchased = [int64]$License.PurchasedUnits } catch { $purchased = 0 }
    }
    if ($License.PSObject.Properties['ConsumedUnits'] -and $null -ne $License.ConsumedUnits -and $License.ConsumedUnits -ne 'N/A' -and $License.ConsumedUnits -ne '') {
        try { $consumed = [int64]$License.ConsumedUnits } catch { $consumed = 0 }
    }

    if ($isTrialFlag) {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Marked as trial by subscription metadata' }
    }

    if ($isFreeOrTrialFlag) {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Marked as free or trial by lifecycle metadata' }
    }

    if ($ignoreLifecycle) {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Ignored due to lifecycle metadata anomaly' }
    }

    $alwaysExcludeSkus = @(
        'MCOPSTNC',
        'STREAM',
        'FORMS_PRO',
        'POWER_BI_STANDARD',
        'RIGHTSMANAGEMENT_ADHOC',
        'PROJECT_MADEIRA_PREVIEW_IW_SKU',
        'DYN365_ENTERPRISE_P1_IW',
        'POWERAPPS_INDIVIDUAL_USER',
        'POWERAPPS_DEV',
        'POWERAPPS_VIRAL',
        'FLOW_FREE',
        'CCIBOTS_PRIVPREV_VIRAL',
        'Power_Pages_vTrial_for_Makers',
        'WINDOWS_STORE'
    )

    if ($alwaysExcludeSkus -contains $skuPartNumber) {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Excluded by known freemium or benefit SKU list' }
    }

    if ($skuPartNumber -match 'TRIAL|EXPLORATORY|FREE|VIRAL|_FACULTY|_STUDENT|PREVIEW|_IW($|_)|ADHOC|INDIVIDUAL|_DEV($|_)|_TRIAL($|_)|FOR_MAKERS') {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Excluded by SKU naming pattern' }
    }

    if ($UserCount -gt 0 -and $appliesTo -eq 'User') {
        $seatThreshold = [math]::Max(($UserCount * 10), 5000)
        $consumedThreshold = [math]::Max(($UserCount * 2), 250)
        if ($purchased -ge $seatThreshold -and $consumed -le $consumedThreshold) {
            return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = "Excluded by tenant user-count heuristic ($purchased seats for $UserCount users)" }
        }
    }

    return [pscustomobject]@{ IsPaid = $true; LicenseClass = 'Paid'; Reason = 'Included as paid license inventory' }
}

function Test-IsPaidLicenseSku {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$License,
        [int]$UserCount = 0
    )
    return (Get-LicenseClassification -License $License -UserCount $UserCount).IsPaid
}

function Get-LicenseInventoryRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$License,
        [int]$UserCount = 0
    )

    $purchased = 0
    $consumed = 0

    if ($null -ne $License.PurchasedUnits -and $License.PurchasedUnits -ne 'N/A' -and $License.PurchasedUnits -ne '') {
        try { $purchased = [int64]$License.PurchasedUnits } catch { $purchased = 0 }
    }

    if ($null -ne $License.ConsumedUnits -and $License.ConsumedUnits -ne 'N/A' -and $License.ConsumedUnits -ne '') {
        try { $consumed = [int64]$License.ConsumedUnits } catch { $consumed = 0 }
    }

    $friendlyName = if ($License.PSObject.Properties['SkuFriendlyName'] -and $License.SkuFriendlyName) {
        $License.SkuFriendlyName
    } else {
        Get-FriendlyProductName -SkuPartNumber $License.SkuPartNumber
    }

    $classification = Get-LicenseClassification -License $License -UserCount $UserCount

    [PSCustomObject]@{
        SkuPartNumber    = $License.SkuPartNumber
        SkuFriendlyName  = $friendlyName
        PurchasedUnits   = $purchased
        ConsumedUnits    = $consumed
        RemainingUnits   = ($purchased - $consumed)
        Utilization      = $(if ($purchased -gt 0) { ($consumed / $purchased) * 100 } else { 0 })
        IsPaid           = $classification.IsPaid
        LicenseClass     = $classification.LicenseClass
        LicenseClassificationReason = $classification.Reason
        ServicePlans     = $(if ($License.PSObject.Properties['ServicePlans']) { $License.ServicePlans } else { $null })
    }
}

# License SKUs and Service Plan IDs to HASH - MGGraph
function Get-AllLicenseSKUs {
    [CmdletBinding()]
    param ()

    # Common Product Name Map Static
    $script:CommonProductNameMapStatic = @{
        'AAD_PREMIUM'                    = 'Microsoft Entra ID P1'
        'AAD_PREMIUM_P2'                 = 'Microsoft Entra ID P2'
        'ATP_ENTERPRISE'                 = 'Microsoft Defender for Office 365 (Plan 1)'
        'DYN365_ENTERPRISE_SALES'        = 'Dynamics 365 for Sales Enterprise Edition'
        'EMS'                            = 'Enterprise Mobility + Security E3'
        'EXCHANGESTANDARD'               = 'Exchange Online (Plan 1)'
        'EXCHANGE_PREMIUM'               = 'Exchange Online (Plan 2)'
        'ENTERPRISEPACK' = 'Office 365 E3'
        'ENTERPRISEPREMIUM' = 'Office 365 E5'
        'FLOW_FREE'                      = 'Microsoft Power Automate Free'
        'M365_BUSINESS_BASIC'            = 'Microsoft 365 Business Basic'
        'M365_BUSINESS_STANDARD'         = 'Microsoft 365 Business Standard'
        'MCOMEETADV'                     = 'Microsoft 365 Audio Conferencing'
        'MICROSOFT_BUSINESS_CENTER'      = 'Microsoft Business Center'
        'Microsoft_Teams_Premium'        = 'Microsoft Teams Premium'
        'Microsoft_Teams_Audio_Conferencing_select_dial_out' = 'Microsoft Teams Audio Conferencing with dial-out to USA/CAN'
        'O365_BUSINESS_ESSENTIALS'       = 'Microsoft 365 Business Basic'
        'O365_BUSINESS_PREMIUM'          = 'Microsoft 365 Business Standard'
        'OFFICESUBSCRIPTION'             = 'Microsoft 365 Apps for Enterprise'
        'PBI_PREMIUM_PER_USER'           = 'Power BI Premium Per User'
        'POWERAPPS_VIRAL'                = 'Microsoft Power Apps Plan 2 Trial'
        'POWER_BI_PRO'                   = 'Power BI Pro'
        'PROJECTPROFESSIONAL'            = 'Planner and Project Plan 3'
        'SPB'                            = 'Microsoft 365 Business Premium'
        'Teams_Exploratory'              = 'Microsoft Teams Exploratory'
        'Teams_Premium_(for_Departments)'= 'Teams Premium (for Departments)'
        'THREAT_INTELLIGENCE'            = 'Microsoft Defender for Office 365 (Plan 2)'
        'VISIOCLIENT'                    = 'Visio Plan 2'
        'VISIO_PLAN2_DEPT'               = 'Visio Plan 2'
        'WINDOWS_STORE'                  = 'Windows Store for Business'
        'M365EDU_A5_STUUSEBNFT'          = 'Microsoft 365 A5 Student Use Benefit'
        'MCOMEETACPEA_FACULTY'           = 'Microsoft 365 Audio Conferencing for Faculty'
        'Power_Pages_vTrial_for_Makers'  = 'Power Pages vTrial for Makers'
        'STANDARDWOFFPACK_FACULTY'       = 'Microsoft 365 Standalone Plan 2 for Faculty'
        'POWERAPPS_DEV'                  = 'Microsoft Power Apps Plan 2 for Developers'
        'SPZA_IW'                        = 'Microsoft Search AI with IW'
        'RMSBASIC'                       = 'Rights Management Basic'
        'Dynamics_365_Onboarding_SKU'    = 'Dynamics 365 Onboarding SKU'
        'M365EDU_A5_FACULTY'             = 'Microsoft 365 A5 Faculty'
        'POWER_BI_STANDARD'              = 'Power BI Standard'
        'FORMS_PRO'                      = 'Forms Pro'
        'CCIBOTS_PRIVPREV_VIRAL'         = 'CCI Bots Private Preview'
        'PROJECT_MADEIRA_PREVIEW_IW_SKU' = 'Project Madeira Preview IW SKU'
        'RIGHTSMANAGEMENT_ADHOC'         = 'Rights Management Adhoc'
        'SPE_E5'                         = 'Microsoft 365 E5'
        'SPE_E3'                         = 'Microsoft 365 E3'
        'SPE_F3'                         = 'Microsoft 365 F3'
        'SPE_F1'                         = 'Microsoft 365 F1'
        'MCOPSTN1'                       = 'Microsoft Teams Phone Standard'
        'MCOCAP'                         = 'Microsoft Teams Shared Space'
        'TEAMS_SHARED_SPACE'             = 'Microsoft Teams Shared Space'
        'MICROSOFT_TEAMS_ENTERPRISE_NEW' = 'Microsoft Teams Enterprise'
        'POWERAPPS_PER_USER'             = 'Power Apps Premium'
        'EXCHANGEENTERPRISE'             = 'Exchange Online (Plan 2)'
        'EXCHANGEDESKLESS'               = 'Exchange Online Kiosk'
        'EMSPREMIUM'                     = 'Enterprise Mobility + Security E5'
        'MCOEV'                          = 'Microsoft Teams Phone Resource Account'
        'PROJECTPLAN3'                   = 'Planner and Project Plan 3'
        'PROJECTPLAN5'                   = 'Planner and Project Plan 5'
        'VISIO_PLAN1_DEPT'               = 'Visio Plan 1'
        'MICROSOFT_365_COPILOT'          = 'Microsoft 365 Copilot'
        'MICROSOFT_365_BUSINESS_PREMIUM_(NO_TEAMS)' = 'Microsoft 365 Business Premium (No Teams)'
        'CPC_LVL_1'                      = 'Windows 365 Enterprise 2 vCPU 4 GB 128 GB (Preview)'
        'CPC_LVL_3'                      = 'Windows 365 Enterprise 4 vCPU 16 GB 256 GB (Preview)'
        'CPC_E_4C_16GB_512GB'            = 'Windows 365 Enterprise 4 vCPU 16 GB 512 GB'
        'CPC_E_8C_32GB_128GB'            = 'Windows 365 Enterprise 8 vCPU 32 GB 128 GB'
        'CPC_E_8C_32GB_256GB'            = 'Windows 365 Enterprise 8 vCPU 32 GB 256 GB'
        'CPC_E_8C_32GB_512GB'            = 'Windows 365 Enterprise 8 vCPU 32 GB 512 GB'
        'WINDOWS_365_S_2VCPU_4GB_64GB'   = 'Windows 365 Shared Use 2 vCPU 4 GB 64 GB'
        'WINDOWS_365_S_2VCPU_4GB_128GB'  = 'Windows 365 Shared Use 2 vCPU 4 GB 128 GB'
        'WINDOWS_365_S_2VCPU_4GB_256GB'  = 'Windows 365 Shared Use 2 vCPU 4 GB 256 GB'
        'WIN10_PRO_ENT_SUB'              = 'Windows 10/11 Enterprise E3'
        'WIN10_VDA_E3'                   = 'Windows 10/11 Enterprise VDA E3'
        'WIN10_VDA_E5'                   = 'Windows 10/11 Enterprise VDA E5'
        'WINE5_GCC_COMPAT'               = 'Windows 10/11 Enterprise E5 Commercial (GCC Compatible)'
        'E3_VDA_ONLY'                    = 'Windows 10/11 Enterprise VDA E3 (VDA only)'
    }

    # Build a hashtable for license sku. Create start time of the function
    $start = Get-Date
    Initialize-MicrosoftLicenseReferenceMap | Out-Null
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["LicenseSKUs"] = @{}
    Write-Log -Type Info -Message "[Get-AllLicenseSKUs] Gathering all License SKUs from tenant" -ExportFileLocation $ExportDetails
    # Get License SKUs using MGGraph   
    $skus = Get-MgSubscribedSku -ProgressAction SilentlyContinue -ErrorAction Continue | ? {$_.AppliesTo}
    $script:SkuLookupById = @{}
    $script:ServicePlanLookupById = @{}

    # Get subscription metadata to identify trials
    $subscriptions = @()
    try {
        $subscriptions = Get-MgDirectorySubscription -All -ProgressAction SilentlyContinue -ErrorAction SilentlyContinue
    } catch {}
    $subscriptionLookup = @{}
    foreach ($sub in $subscriptions) {
        if ($sub.SkuId) {
            $subscriptionLookup[$sub.SkuId.ToString()] = $sub
        }
    }
    
    Write-Log -Type Info -Message "[Get-AllLicenseSKUs] Found $($skus.count) License SKUs from tenant" -ExportFileLocation $ExportDetails

    #Add License SKUs to Hash Table and Service Plans under each SKU to another Hash Table
    Write-Log -Type Info -Message "[Get-AllLicenseSKUs] START: Add License SKUs and Service Plans under each SKU to Tenant Hash Table" -ExportFileLocation $ExportDetails
    foreach ($sku in $skus) {
        $skuIdText = $sku.SkuId.ToString()
        $subMatch = $null
        if ($subscriptionLookup.ContainsKey($skuIdText)) {
            $subMatch = $subscriptionLookup[$skuIdText]
        }

        $friendlySkuName = Get-FriendlyProductName -SkuPartNumber $sku.SkuPartNumber
        $script:SkuLookupById[$skuIdText] = [PSCustomObject]@{
            SkuId = $skuIdText
            SkuPartNumber = $sku.SkuPartNumber
            FriendlyName = $friendlySkuName
            ServicePlans = @($sku.ServicePlans)
        }
        foreach ($servicePlan in @($sku.ServicePlans)) {
            if ($servicePlan.ServicePlanId) {
                $script:ServicePlanLookupById[$servicePlan.ServicePlanId.ToString()] = $servicePlan.ServicePlanName
            }
        }

        $skuDetails = [PSCustomObject]@{
            AccountSkuId = $sku.AccountSkuId
            AccountName = $sku.AccountName
            AppliesTo = $sku.AppliesTo
            CapabilityStatus = $sku.CapabilityStatus
            SkuId = $sku.SkuId
            SkuPartNumber = $sku.SkuPartNumber
            SkuFriendlyName = $friendlySkuName
            ConsumedUnits = $sku.ConsumedUnits
        }

        #Add Prepaid Units
        $skuDetails | Add-Member -MemberType NoteProperty -Name PurchasedUnits -Value $sku.PrepaidUnits.Enabled -Force
        #Add Remaining/Available Units
        $skuDetails | Add-Member -MemberType NoteProperty -Name RemainingUnits -Value ($sku.PrepaidUnits.Enabled - $sku.ConsumedUnits) -Force
        $skuDetails | Add-Member -MemberType NoteProperty -Name ServicePlansCount -Value (($sku.ServicePlans.ServicePlanName).count) -Force
        $skuDetails | Add-Member -MemberType NoteProperty -Name ServicePlans -Value ($sku.ServicePlans.ServicePlanName -join ",") -Force
        if ($subMatch) {
            $skuDetails | Add-Member -MemberType NoteProperty -Name IsTrial -Value $subMatch.IsTrial -Force
            $skuDetails | Add-Member -MemberType NoteProperty -Name SubscriptionStatus -Value $subMatch.Status -Force
            $skuDetails | Add-Member -MemberType NoteProperty -Name NextLifecycleDateTime -Value $subMatch.NextLifecycleDateTime -Force
            $skuDetails | Add-Member -MemberType NoteProperty -Name TotalLicenses -Value $subMatch.TotalLicenses -Force

            $nextLifecycle = $subMatch.NextLifecycleDateTime
            $isFreeOrTrial = ($null -eq $nextLifecycle -or $nextLifecycle -eq '')
            $isIgnoredLifecycle = $false
            if ($nextLifecycle) {
                try {
                    $dt = [datetime]$nextLifecycle
                    if ($dt.Year -eq 9999) { $isIgnoredLifecycle = $true }
                } catch {}
            }
            if ($sku.SkuPartNumber -eq 'MCOPSTNC') {
                $isFreeOrTrial = $true
            }
            $skuDetails | Add-Member -MemberType NoteProperty -Name IsFreeOrTrial -Value $isFreeOrTrial -Force
            $skuDetails | Add-Member -MemberType NoteProperty -Name IgnoreLifecycle -Value $isIgnoredLifecycle -Force
        }

        $skuClassification = Get-LicenseClassification -License $skuDetails
        $skuDetails | Add-Member -MemberType NoteProperty -Name IsPaid -Value $skuClassification.IsPaid -Force
        $skuDetails | Add-Member -MemberType NoteProperty -Name LicenseClass -Value $skuClassification.LicenseClass -Force
        $skuDetails | Add-Member -MemberType NoteProperty -Name LicenseClassificationReason -Value $skuClassification.Reason -Force

        Write-Log -Type DEBUG -Message "[Get-AllLicenseSKUs] Gathering License details for $($AccountSkuId)" -ExportFileLocation $ExportDetails

        #Create Hash Table for License SKUs
        $script:tenantStatsHash["LicenseSKUs"][$sku.SkuId.tostring()] = $skuDetails

    }
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Log -Type Info -Message "[Get-AllLicenseSKUs] COMPLETED: Gathering all License Details User Details in $($CompletedTime)" -ExportFileLocation $ExportDetails
}
function Get-AllUserDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel
    )
    $start = Get-Date
    $graphUsersProgressId = 71
    $userDetailsProgressId = 72
    $DesiredProperties = @(
        "DisplayName", "AssignedLicenses", "UserPrincipalName"
        "UserType", "Id", "AccountEnabled"
        "CreatedDateTime", "Mail", "JobTitle"
        "Department", "CompanyName", "OfficeLocation"
        "City", "State", "Country"
        "OnPremisesSyncEnabled", "OnPremisesDistinguishedName", "OnPremisesLastSyncDateTime"
        "UsageLocation", "SignInActivity", "ProxyAddresses"
    )
    $minimumModeMessage = 'NotCollected (minimum mode)'
    $progressStatusInterval = 25
    $isGeekDetail = ($detailLevel -eq 'geek')
    $logPerUserDebug = $isGeekDetail
    $BasicMGDetails = $false
    $userCollectionState = [ordered]@{
        ProcessedUserCount = 0
        LicensedUserCount = 0
        UnlicensedUserCount = 0
        LicenseFallbackUserCount = 0
        UnresolvedSkuCount = 0
    }

    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["Users"] = @{}

    $depthPolicy = if ($script:CollectionDepthPolicy) {
        $script:CollectionDepthPolicy
    }
    else {
        Get-ArrayaCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($detailLevel.ToLowerInvariant()))
    }
    $collectExtendedUserDetails = if ($null -ne $depthPolicy.CollectExtendedGraphEnrichment) { [bool]$depthPolicy.CollectExtendedGraphEnrichment } else { $true }
    if (-not $collectExtendedUserDetails) {
        Write-Log -Type INFO -Message "[Get-allUserDetails] Minimum mode optimization active. Skipping per-user service-plan expansion and request-id sign-in fields." -ExportFileLocation $ExportDetails
    }

    function Process-TenantUserRecord {
        param(
            [Parameter(Mandatory = $true)]
            [object]$UserRecord
        )

        $scriptLabel = if ($UserRecord.DisplayName) { [string]$UserRecord.DisplayName } elseif ($UserRecord.UserPrincipalName) { [string]$UserRecord.UserPrincipalName } else { 'User record' }
        try {
            if ($null -eq $script:tenantStatsHash -or -not ($script:tenantStatsHash -is [System.Collections.IDictionary])) {
                $script:tenantStatsHash = @{}
            }
            if ($null -eq $script:tenantStatsHash['Users'] -or -not ($script:tenantStatsHash['Users'] -is [System.Collections.IDictionary])) {
                $script:tenantStatsHash['Users'] = @{}
            }
            if ($null -eq $script:SkuLookupById -or -not ($script:SkuLookupById -is [System.Collections.IDictionary])) {
                $script:SkuLookupById = @{}
            }
            if ($null -eq $script:ServicePlanLookupById -or -not ($script:ServicePlanLookupById -is [System.Collections.IDictionary])) {
                $script:ServicePlanLookupById = @{}
            }

            $userCollectionState.ProcessedUserCount++
            if ($userCollectionState.ProcessedUserCount -eq 1 -or ($userCollectionState.ProcessedUserCount % $progressStatusInterval) -eq 0) {
                Write-Progress -Id $userDetailsProgressId -Activity "Gathering Tenant User Details" -Status "Processed $($userCollectionState.ProcessedUserCount) user(s): $scriptLabel"
            }

            $userPrincipalName = [string]$UserRecord.UserPrincipalName
            if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
                $userPrincipalName = if ($UserRecord.Id) { "id:$($UserRecord.Id)" } else { "unknown:$([guid]::NewGuid().Guid)" }
            }
            if ($logPerUserDebug) {
                Write-Log -Type DEBUG -Message ("[Get-allUserDetails] Creating Hash for '{0}'" -f $userPrincipalName) -ExportFileLocation $ExportDetails
            }

            $signInActivity = if ($UserRecord.PSObject.Properties['SignInActivity']) { $UserRecord.SignInActivity } else { $null }
            $userProperties = [ordered]@{}
            if ($isGeekDetail) {
                foreach ($property in $UserRecord.PSObject.Properties) {
                    $userProperties[$property.Name] = $property.Value
                }
            }
            else {
                $userProperties['DisplayName'] = $UserRecord.DisplayName
                $userProperties['AssignedLicenses'] = $UserRecord.AssignedLicenses
                $userProperties['UserPrincipalName'] = $UserRecord.UserPrincipalName
                $userProperties['UserType'] = $UserRecord.UserType
                $userProperties['Id'] = $UserRecord.Id
                $userProperties['AccountEnabled'] = $UserRecord.AccountEnabled
                $userProperties['CreatedDateTime'] = $UserRecord.CreatedDateTime
                $userProperties['Mail'] = $UserRecord.Mail
                $userProperties['JobTitle'] = $UserRecord.JobTitle
                $userProperties['Department'] = $UserRecord.Department
                $userProperties['CompanyName'] = $UserRecord.CompanyName
                $userProperties['OfficeLocation'] = $UserRecord.OfficeLocation
                $userProperties['City'] = $UserRecord.City
                $userProperties['State'] = $UserRecord.State
                $userProperties['Country'] = $UserRecord.Country
                $userProperties['OnPremisesSyncEnabled'] = $UserRecord.OnPremisesSyncEnabled
                $userProperties['OnPremisesDistinguishedName'] = $UserRecord.OnPremisesDistinguishedName
                $userProperties['OnPremisesLastSyncDateTime'] = $UserRecord.OnPremisesLastSyncDateTime
                $userProperties['UsageLocation'] = $UserRecord.UsageLocation
                $userProperties['SignInActivity'] = $signInActivity
                $userProperties['ProxyAddresses'] = $UserRecord.ProxyAddresses
            }

            $combinedProxyAddresses = ($UserRecord.ProxyAddresses -replace '^[sS][mM][tT][pP]:') -join ';'
            $userProperties['ProxyAddresses'] = $combinedProxyAddresses

            if ($BasicMGDetails) {
                if ($logPerUserDebug) {
                    Write-Log -Type DEBUG -Message ("[Get-allUserDetails] Updating '{0}' UserType to HashTable if Basic Details" -f $userPrincipalName) -ExportFileLocation $ExportDetails
                }
                $userProperties['UserType'] = (if ($UserRecord.UserPrincipalName -like "*#EXT#*") { "GuestUser" } else { "User" })
            }
            else {
                if ($logPerUserDebug) {
                    Write-Log -Type DEBUG -Message ("[Get-allUserDetails] Gather '{0}' License Friendly Names" -f $userPrincipalName) -ExportFileLocation $ExportDetails
                }
                $assignedLicensesString = $null
                $assignedLicensesFriendlyString = $null
                $disabledPlans = $null
                $enabledServicePlans = $null
                $assignedLicenseEntries = @(
                    @($UserRecord.AssignedLicenses) | Where-Object { $null -ne $_ }
                )
                if ($assignedLicenseEntries.Count -gt 0) {
                    $userCollectionState.LicensedUserCount++
                    $resolvedSkuParts = New-Object 'System.Collections.Generic.HashSet[string]'
                    $resolvedFriendlyNames = New-Object 'System.Collections.Generic.HashSet[string]'
                    $resolvedDisabledPlans = if ($collectExtendedUserDetails) { New-Object 'System.Collections.Generic.HashSet[string]' } else { $null }
                    $resolvedEnabledPlans = if ($collectExtendedUserDetails) { New-Object 'System.Collections.Generic.HashSet[string]' } else { $null }

                    foreach ($assignedLicenseEntry in $assignedLicenseEntries) {
                        $skuId = $null
                        if ($assignedLicenseEntry.PSObject.Properties['SkuId']) {
                            $skuId = $assignedLicenseEntry.SkuId
                        }
                        elseif ($assignedLicenseEntry -is [guid]) {
                            $skuId = $assignedLicenseEntry
                        }

                        $skuIdText = if ($skuId) { $skuId.ToString() } else { $null }
                        $skuLookup = if ($skuIdText -and $script:SkuLookupById.ContainsKey($skuIdText)) { $script:SkuLookupById[$skuIdText] } else { $null }
                        if ($skuLookup) {
                            [void]$resolvedSkuParts.Add($skuLookup.SkuPartNumber)
                            [void]$resolvedFriendlyNames.Add($skuLookup.FriendlyName)

                            if ($collectExtendedUserDetails) {
                                $disabledPlanIds = @()
                                if ($assignedLicenseEntry.PSObject.Properties['DisabledPlans'] -and $assignedLicenseEntry.DisabledPlans) {
                                    $disabledPlanIds = @($assignedLicenseEntry.DisabledPlans | ForEach-Object { $_.ToString() })
                                }

                                $disabledPlanNamesForSku = @()
                                foreach ($disabledPlanId in $disabledPlanIds) {
                                    if ($script:ServicePlanLookupById.ContainsKey($disabledPlanId)) {
                                        $disabledPlanName = $script:ServicePlanLookupById[$disabledPlanId]
                                        $disabledPlanNamesForSku += $disabledPlanName
                                        [void]$resolvedDisabledPlans.Add($disabledPlanName)
                                    }
                                }

                                $enabledPlanNamesForSku = @(
                                    @($skuLookup.ServicePlans) |
                                        Where-Object {
                                            $_.ServicePlanName -and
                                            ($disabledPlanNamesForSku -notcontains $_.ServicePlanName)
                                        } |
                                        ForEach-Object { $_.ServicePlanName }
                                )
                                foreach ($enabledPlanName in $enabledPlanNamesForSku) {
                                    [void]$resolvedEnabledPlans.Add($enabledPlanName)
                                }
                            }
                        }
                        else {
                            $userCollectionState.UnresolvedSkuCount++
                            if ($skuIdText) {
                                [void]$resolvedSkuParts.Add($skuIdText)
                                [void]$resolvedFriendlyNames.Add($skuIdText)
                            }
                        }
                    }

                    $userCollectionState.LicenseFallbackUserCount++
                    $assignedLicensesString = ($resolvedSkuParts | ForEach-Object { $_ }) -join ","
                    $assignedLicensesFriendlyString = ($resolvedFriendlyNames | ForEach-Object { $_ }) -join ","
                    if ($collectExtendedUserDetails) {
                        $disabledPlans = @($resolvedDisabledPlans | ForEach-Object { $_ })
                        $enabledServicePlans = ($resolvedEnabledPlans | ForEach-Object { $_ }) -join ","
                    }
                    else {
                        $disabledPlans = @()
                        $enabledServicePlans = $minimumModeMessage
                    }
                }
                else {
                    $userCollectionState.UnlicensedUserCount++
                    if (-not $collectExtendedUserDetails) {
                        $disabledPlans = @()
                        $enabledServicePlans = $minimumModeMessage
                    }
                }
                $userProperties['AssignedLicenses'] = $assignedLicensesString
                $userProperties['AssignedLicensesFriendly'] = $assignedLicensesFriendlyString
                $userProperties['License-DisabledArray'] = $disabledPlans
                $userProperties['EnabledServicePlans'] = $enabledServicePlans
                $userProperties['LastNonInteractiveSignInDateTime'] = if ($signInActivity) { $signInActivity.LastNonInteractiveSignInDateTime } else { $null }
                if ($collectExtendedUserDetails) {
                    $userProperties['LastNonInteractiveSignInRequestId'] = if ($signInActivity) { $signInActivity.LastNonInteractiveSignInRequestId } else { $null }
                    $userProperties['LastSignInRequestId'] = if ($signInActivity) { $signInActivity.LastSignInRequestId } else { $null }
                }
                else {
                    $userProperties['LastNonInteractiveSignInRequestId'] = $minimumModeMessage
                    $userProperties['LastSignInRequestId'] = $minimumModeMessage
                }
                $userProperties['LastSignInDateTime'] = if ($signInActivity) { $signInActivity.LastSignInDateTime } else { $null }
            }

            $script:tenantStatsHash["Users"][$userPrincipalName] = [PSCustomObject]$userProperties
        }
        catch {
            $lineNumber = if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) { $_.InvocationInfo.ScriptLineNumber } else { 'unknown' }
            Write-Log -Type ERROR -Message ("[Get-allUserDetails] An error occurred in Creating User Hash for user '{0}'. Line={1}. $($_.Exception.Message)" -f $scriptLabel, $lineNumber) -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }
    }

    function Invoke-UserCollectionQuery {
        param(
            [switch]$BasicMode
        )

        if ($BasicMode) {
            Invoke-QuietCommand -ScriptBlock { Get-MgUser -All -ErrorAction Stop } |
                Where-Object { $null -ne $_.ID } |
                ForEach-Object { Process-TenantUserRecord -UserRecord $_ }
            return
        }

        if ($detailLevel -eq 'geek') {
            Invoke-QuietCommand -ScriptBlock { Get-MgUser -All -ErrorAction Stop } |
                Where-Object { $null -ne $_.ID } |
                ForEach-Object { Process-TenantUserRecord -UserRecord $_ }
            return
        }

        Invoke-QuietCommand -ScriptBlock { Get-MgUser -All -Property $DesiredProperties -ErrorAction Stop } |
            Where-Object { $null -ne $_.ID } |
            ForEach-Object { Process-TenantUserRecord -UserRecord $_ }
    }

    try {
        Write-Host "Getting all Microsoft Graph $($detailLevel) User data..." -ForegroundColor Cyan -nonewline
        Write-Log -Type Info -Message "[Get-allUserDetails] START: Getting all Microsoft Graph $($detailLevel) User data" -ExportFileLocation $ExportDetails
        Write-Progress -Id $graphUsersProgressId -Activity "Getting all Microsoft Graph User Data" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))

        $maxRetries = 3
        $collectionSucceeded = $false
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                $script:tenantStatsHash["Users"] = @{}
                $userCollectionState.ProcessedUserCount = 0
                $userCollectionState.LicensedUserCount = 0
                $userCollectionState.UnlicensedUserCount = 0
                $userCollectionState.LicenseFallbackUserCount = 0
                $userCollectionState.UnresolvedSkuCount = 0
                Invoke-UserCollectionQuery
                $collectionSucceeded = $true
                break
            }
            catch {
                if ($attempt -lt $maxRetries) {
                    $waitSeconds = [math]::Min([math]::Pow(2, $attempt), 30)
                    Write-Log -Type WARNING -Message "[Get-allUserDetails] Attempt $attempt failed: $($_.Exception.Message). Retrying in $waitSeconds seconds..." -ExportFileLocation $ExportDetails
                    Start-Sleep -Seconds $waitSeconds
                }
                else {
                    throw
                }
            }
        }
        if (-not $collectionSucceeded) {
            throw "Graph user collection did not complete successfully."
        }
    }
    catch {
        if ($_.Exception.Message -like "*Neither tenant is B2C or tenant doesn't have premium license*") {
            Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_

            Write-Host
            Write-Host "Caught a tenant license exception. Getting all Microsoft Graph User data without licenses and sign in activity..." -ForegroundColor Yellow -nonewline
            try {
                Write-Log -Type Info -Message "[Get-allUserDetails] Attempt 2. Getting all Microsoft Graph $($detailLevel) with limited User Details" -ExportFileLocation $ExportDetails
                $script:tenantStatsHash["Users"] = @{}
                $userCollectionState.ProcessedUserCount = 0
                $userCollectionState.LicensedUserCount = 0
                $userCollectionState.UnlicensedUserCount = 0
                $userCollectionState.LicenseFallbackUserCount = 0
                $userCollectionState.UnresolvedSkuCount = 0
                $BasicMGDetails = $true
                Invoke-UserCollectionQuery -BasicMode
            }
            catch {
                Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
        }
        else {
            Write-Log -Type Error -Message "[Get-allUserDetails] An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            Write-Log -Type WARNING -Message "[Get-allUserDetails] Continuing without user details due to Graph request failure." -ExportFileLocation $ExportDetails
            return
        }
    }
    finally {
        Write-Progress -Id $graphUsersProgressId -Activity "Getting all Microsoft Graph User Data" -Completed
    }

    Write-Log -Type Info -Message "[Get-allUserDetails] Added additional properties for $($userCollectionState.ProcessedUserCount) users" -ExportFileLocation $ExportDetails
    try {
        Write-Progress -Id $userDetailsProgressId -Activity "Gathering Tenant User Details" -Status "Processed $($userCollectionState.ProcessedUserCount) user(s)"
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-allUserDetails] Licensing summary: LicensedUsers=$($userCollectionState.LicensedUserCount) UnlicensedUsers=$($userCollectionState.UnlicensedUserCount) LicenseLookupUsers=$($userCollectionState.LicenseFallbackUserCount) UnresolvedSkuReferences=$($userCollectionState.UnresolvedSkuCount)" -ExportFileLocation $ExportDetails
        Write-Log -Type Info -Message "[Get-allUserDetails] COMPLETED: Gathering all  User Details in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }     
    catch {
        Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-Progress -Id $userDetailsProgressId -Activity "Gathering Tenant User Details" -Completed
    }
}

#Gather all Office 365 Admins
function Get-AllOffice365Admins {
    param ()
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["Admins"] = @{}
    $adminResults = New-Object System.Collections.Generic.List[object]
    Write-Host "Gathering All Admins ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllOffice365Admins] START: Gathering All Admins from Tenant" -ExportFileLocation $ExportDetails

    try {
        # Gather all Admin Roles
        $adminRoles = Get-MgDirectoryRole -ProgressAction SilentlyContinue | Select DisplayName, ID, Description | ? {$null -ne $_.DisplayName}
        $totalCount = $adminRoles.count
        Write-Log -Type INFO -Message "[Get-AllOffice365Admins] Admin Roles Found: $($adminRoles.count)" -ExportFileLocation $ExportDetails
        foreach ($role in $adminRoles) {
            $roleName = $role.DisplayName
            Write-Log -Type DEBUG -Message "[Get-AllOffice365Admins] $($roleName): Gathering Admins in Role" -ExportFileLocation $ExportDetails
            Write-ProgressHelper -Total $totalCount -Id 1 -Activity "Gathering Admins in Roles" -Operation "Checking Role: $($roleName)" 
            $roleMemberList = Invoke-ArrayaCollectionStepSafe -OperationName "Get-AllOffice365Admins role membership for $roleName" -DefaultValue @() -ExportFileLocation $ExportDetails -ScriptBlock {
                @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ProgressAction SilentlyContinue -ErrorAction Stop | Where-Object { $null -ne $_.Id })
            }
            if ($roleMemberList) {
                $totalCount2 = $roleMemberList.count
                Write-Log -Type INFO -Message "[Get-AllOffice365Admins] $($roleName) Users Found: $($roleMemberList.count)" -ExportFileLocation $ExportDetails
                foreach ($roleMember in $roleMemberList) {
                    $Name = $null
                    $UPN = $null
                    $mail = $null
                    $jobTitle = $null
                    $UserType = $null
                    $AccountEnabled = $null
                    $LastSignInDateTime = $null
                    $CreatedDate = $null
                    $GroupMailEnabled = $null
                    $GroupMailNickname = $null
                    $GroupType = $null
                    $Name = $roleMember.AdditionalProperties['displayName']
                    if ($roleMember.AdditionalProperties['@odata.type'] -eq "#microsoft.graph.group") {
                        # Group Specific Values
                        $GroupMailEnabled = $roleMember.AdditionalProperties['mailEnabled']
                        $GroupMailNickname = $roleMember.AdditionalProperties['mailNickname']
                        $GroupType = ($roleMember.AdditionalProperties['groupTypes'] -join ",")
                        $CreatedDate = [datetime]::Parse($roleMember.AdditionalProperties['createdDateTime']).ToShortDateString()
                    } elseif ($roleMember.AdditionalProperties['@odata.type'] -eq "#microsoft.graph.user") {
                        # User Specific Values
                        $UPN = $roleMember.AdditionalProperties['userPrincipalName']
                        $mail = $roleMember.AdditionalProperties['mail']
                        $jobTitle = $roleMember.AdditionalProperties['jobTitle']

                        # Check if user is in tenant stats hash (guarded for profile/collector skips)
                        $usersLookup = if (
                            $script:tenantStatsHash.ContainsKey('Users') -and
                            $script:tenantStatsHash['Users'] -is [System.Collections.IDictionary]
                        ) {
                            $script:tenantStatsHash['Users']
                        }
                        else {
                            $null
                        }
                        if ($usersLookup -and $usersLookup.ContainsKey($UPN)) {
                            $userMatch = $usersLookup[$UPN]
                        } else { $userMatch = $null }
                        $UserType = $userMatch.UserType
                        $AccountEnabled = $userMatch.AccountEnabled
                        $LastSignInDateTime = $userMatch.LastSignInDateTime
                        $CreatedDate = $userMatch.CreatedDateTime

                    }
                    Write-Log -Type DEBUG "[Get-AllOffice365Admins] Gathering MGGraph Role Details for User: $($Name)" -ExportFileLocation $ExportDetails
                    $currentAdmin = New-Object PSObject -Property ([ordered]@{
                        # Default Values
                        Role = $roleName
                        ObjectType = $roleMember.AdditionalProperties['@odata.type']
                        DisplayName = $Name
                        Mail = $mail
                        CreatedDate = if ($CreatedDate) { $CreatedDate } else { $null }
                        # User Specific Values
                        UserPrincipalName = if ($UPN) { $UPN } else { $null }
                        AccountEnabled = if ($AccountEnabled) { $AccountEnabled } else { $null }
                        userType = if ($UserType) { $UserType } else { $null }
                        JobTitle = if ($jobTitle) { $jobTitle } else { $null }
                        LastSignInDateTime = if ($LastSignInDateTime) { $LastSignInDateTime } else { $null }
                        # Group Specific Values
                        GroupMailEnabled = if ($null -ne $GroupMailEnabled) { $GroupMailEnabled } else { $null }
                        GroupMailNickname = if ($GroupMailNickname) { $GroupMailNickname } else { $null }
                        GroupType = if ($GroupType) { $GroupType } else { $null }
                    })
                        
                    Write-ProgressHelper -Total $totalCount2 -Id 2 -ParentId 1 -Activity "Gathering Admin Details" -Operation "Gathering Admin Role Details: $($roleMember.DisplayName)"
    
                    $adminResults.Add($currentAdmin)
                }
            }
            
        }
        
        #Group by DisplayName or UserPrincipalName and combine roles into a comma-separated list
        $groupedResults = $adminResults | Group-Object -Property DisplayName,UserPrincipalName
        Write-Log -Type INFO -Message "[Get-AllOffice365Admins] Combining roles across $($groupedResults.Count) grouped admin entries" -ExportFileLocation $ExportDetails
        $finalResults = $groupedResults | ForEach-Object {
            $group = $_.Group
            $roles = ($group.Role -join ', ')
            # Use the properties of the first user in each group, but replace the Role with the combined roles
            $group[0] | Add-Member -MemberType NoteProperty -Name 'Role' -Value $roles -Force
            $group[0] | Add-Member -MemberType NoteProperty -Name 'RolesAssigned' -Value (($group.Role | Measure-Object).count) -Force
            $group[0]
        }

        foreach ($result in $finalResults) {
            $adminKeyBase = if (-not [string]::IsNullOrWhiteSpace([string]$result.UserPrincipalName)) {
                ([string]$result.UserPrincipalName).Trim().ToLowerInvariant()
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$result.Mail)) {
                ([string]$result.Mail).Trim().ToLowerInvariant()
            }
            else {
                # Keep this deterministic for non-user principals.
                ("{0}|{1}|{2}" -f [string]$result.ObjectType, [string]$result.DisplayName, [string]$result.Role).ToLowerInvariant()
            }

            $adminKey = $adminKeyBase
            $duplicateSuffix = 2
            while ($script:tenantStatsHash["Admins"].ContainsKey($adminKey)) {
                $adminKey = "{0}#{1}" -f $adminKeyBase, $duplicateSuffix
                $duplicateSuffix++
            }

            $script:tenantStatsHash["Admins"][$adminKey] = $result
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllOffice365Admins] An error occurred in running Get-AllOffice365Admins function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-ProgressHelper -Total $totalCount -Id 1 -Activity "Gathering Admins in Roles" -Completed
        Write-ProgressHelper -Total 1 -Id 2 -Activity "Gathering Admin Details" -Completed
        Write-Log -Type INFO -Message "[Get-AllOffice365Admins] COMPLETED: Gathering All Admins in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }
}

# Get all Office 365 Domains
function Get-AllOffice365Domains {
    param ()
    function Get-RecipientCounts {
        param (
            [Parameter(Mandatory = $true)]
            [array]$Recipients, # Array of recipient objects
    
            [Parameter(Mandatory = $true)]
            [string]$DomainName # The domain name to filter on
        )
    
        # Initialize counts
        $RecipientsWithPrimarySMTPCount = 0
        $DomainRecipientsCount = 0
        $RecipientsAliasOnlyCount = 0
    
        try {
            Write-Log -Type INFO -Message "[Get-AllOffice365Domains] Calculating recipient counts for domain: $DomainName" -ExportFileLocation $ExportDetails
            # Count recipients with Primary SMTP using the domain
            $RecipientsWithPrimarySMTPCount = ($Recipients | Where-Object {
                $_.PrimarySmtpAddress -like "*@$($DomainName)"
            }).Count
            Write-Log -Type INFO -Message "[Get-AllOffice365Domains] '$($DomainName)' Recipients with Primary SMTP Count: $RecipientsWithPrimarySMTPCount" -ExportFileLocation $ExportDetails
    
            # Count recipients with the domain in any of their aliases (EmailAddresses)
            $DomainRecipientsCount = ($Recipients | Where-Object {
                # Split EmailAddresses into an array
                $emailArray = $_.EmailAddresses -split ','
                # Check if any of the addresses contain the domain
                $emailArray -like "*@$($DomainName)*"
            }).Count
            Write-Log -Type INFO -Message "[Get-AllOffice365Domains] '$($DomainName)' Recipients with Alias Count: $DomainRecipientsCount" -ExportFileLocation $ExportDetails
    
            # Count recipients with the domain as alias only (not primary)
            $RecipientsAliasOnlyCount = ($Recipients | Where-Object {
                # Split EmailAddresses into an array
                $emailArray = $_.EmailAddresses -split ','
    
                # Check if there is at least one lowercase smtp alias with the domain
                $hasAlias = $emailArray -like "smtp:*@$($DomainName)"
                # Ensure the primary SMTP is not the domain
                $notPrimarySMTP = $_.PrimarySmtpAddress -notlike "*@$($DomainName)"
    
                # Return true only if the alias exists and it's not the primary SMTP
                $hasAlias -and $notPrimarySMTP
            }).Count
            Write-Log -Type INFO -Message "[Get-AllOffice365Domains] '$($DomainName)' Recipients with Alias Only Count: $RecipientsAliasOnlyCount" -ExportFileLocation $ExportDetails
    
        } catch {
            Write-Host "An error occurred while calculating recipient counts: $($_.Exception.Message)" -ForegroundColor Red
        }
    
        # Return results as a custom object
        return [PSCustomObject]@{
            PrimarySMTPCount = $RecipientsWithPrimarySMTPCount
            AliasOnlyCount   = $RecipientsAliasOnlyCount
            TotalDomainRecipientsCount = $DomainRecipientsCount
        }
    }

    # Resolve DNS
    function Get-DNSHostCompanyName {
        param (
            [Parameter(Mandatory=$true)]
            [object]$nsRecords
        )
        # Define a mapping of common DNS hosts to company names
        $dnsHostMapping = @{
            'ptd.net' = 'PenTeleData'
            'comcast.net' = 'Comcast'
            'charter.com' = 'Charter Communications'
            'rr.com' = 'Road Runner'
            'verizon.net' = 'Verizon'
            'cox.net' = 'Cox Communications'
            'sbcglobal.net' = 'AT&T'
            'frontiernet.net' = 'Frontier Communications'
            'earthlink.net' = 'EarthLink'
            'oraclecloud.net' = 'Oracle'
            'microsoft.com' = 'Microsoft'
            'google.com' = 'Google'
            'worldnic.com' = 'Network Solutions'
            'cloudflare.com' = 'Cloudflare'
            'domaincontrol.com' = 'GoDaddy'
            'namecheaphosting.com' = 'Namecheap'
            # Add more mappings as needed
        }
    
        # Function to get the main domain from a DNS host
        function Get-MainDomain {
            param (
                [string]$DNSHost
            )
            $parts = $DNSHost -split '\.'
            return ($parts[-2] + '.' + $parts[-1])
        }

        # Return if NS records are empty
        if (-not $nsRecords) {
            return @()
        }
    
        # Ensure nsRecords is an array
        if (-not ($nsRecords -is [array])) {
            $nsRecords = @($nsRecords)
        }
    
        # Get NS records and company names
        $mainDomains = $nsRecords | ForEach-Object { Get-MainDomain -DNSHost $_.NameHost }
        $uniqueMainDomains = $mainDomains | Select-Object -Unique
        $nsCompanies = ($uniqueMainDomains | ForEach-Object {
            if ($dnsHostMapping.ContainsKey($_)) {
                $dnsHostMapping[$_]
            } else {
                $_
            }
        })
    
        return $nsCompanies
    }

    # Initialize the hash tables for the domains and remote domains.
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["Domains"] = @{}
    $script:tenantStatsHash["RemoteDomains"] = @{}
    Write-Host "Gathering All Domains ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllOffice365Domains] START: Gathering All Domains from tenant" -ExportFileLocation $ExportDetails
    try {
        # Get all the domains
        $domains = Get-MgDomain -ProgressAction SilentlyContinue | ? {$null -ne $_.ID}
        #Gather Exchange Online Domain Details
        try {
            Write-Log -Type INFO -Message "[Get-AllOffice365Domains] Gathering Exchange Online Domain Details" -ExportFileLocation $ExportDetails
            $acceptedDomains = Get-AcceptedDomain
            $remoteDomains = Get-RemoteDomain | select Identity, DomainName, IsInternal, TargetDeliveryDomain, AllowedOOFType, AutoReplyEnabled, AutoForwardEnabled, DeliveryReportEnabled, NDREnabled, MeetingForwardNotificationEnabled, ContentType, TNEFEnabled, TrustedMailOutboundEnabled, TrustedMailInboundEnabled
            if ($script:tenantStatsHash['AllRecipients']) {
                $recipients = $script:tenantStatsHash['AllRecipients'].values
            } else {$recipients = Get-EXORecipient -ResultSize Unlimited}
                        
            $exchangeOnline = $True
        }
        catch {  
            if ($_.Exception.Message -like "*The term 'Get-AcceptedDomain' is not recognized*") {
                Write-Log -Type ERROR -Message "An error occurred in running Get-AllOffice365Domains function. Exception: Not Properly Connected to Exchange Online Tenant. Check Permissions and reconnect to Exchange Online. Skipping Exchange portion" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                
                $exchangeOnline = $False
            }
            else {
                Write-Log -Type Error -Message "[Get-AllOffice365Domains] Exchange Online not connected for other reasons. $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                Write-Error $_.Exception.Message
            }
        }
        # Get all Remote and Accepted domains
        $totalCount = $domains.count
        Write-Log  -Type INFO -Message "[Get-AllOffice365Domains] $($domains.count) domains found in tenant" -ExportFileLocation $ExportDetails

        foreach ($domain in $domains) {
            try {
                # Get the DNS records
                $domainName = $domain.ID
                $domainVerified = $domain.IsVerified       
                $AuthenticationType = if ($null -ne $domain.AuthenticationType -and -not [string]::IsNullOrWhiteSpace([string]$domain.AuthenticationType)) { [string]$domain.AuthenticationType } else { 'Unknown' }
                Write-Log -Type DEBUG -Message "[Get-AllOffice365Domains] Gathering '$($domainName)' domain details" -ExportFileLocation $ExportDetails
                Write-ProgressHelper -Total $totalCount -Id 1 -Activity "Gathering Domain Details" -Operation "'$($domainName)'"

                #Gather DNS Records from 1.1.1.1
                Write-Log -Type DEBUG -Message "[Get-AllOffice365Domains] Gathering DNS Records for '$($domainName)'" -ExportFileLocation $ExportDetails
                $aRecords = Resolve-DnsName -Name $domainName -Server 1.1.1.1 -Type A -ErrorAction SilentlyContinue -verbose:$false
                $mxRecords = Resolve-DnsName -Name $domainName -Server 1.1.1.1 -Type MX -ErrorAction SilentlyContinue -verbose:$false
                $NSRecords = Resolve-DnsName -Name $domainName -Server 1.1.1.1 -Type NS -ErrorAction SilentlyContinue -verbose:$false
                $txtRecords = Resolve-DnsName -Name $domainName -Server 1.1.1.1 -Type TXT -ErrorAction SilentlyContinue -verbose:$false
                $dmarcRecords = Resolve-DnsName -Name ("_dmarc.{0}" -f $domainName) -Server 1.1.1.1 -Type TXT -ErrorAction SilentlyContinue -verbose:$false

                $selectorRecords = @()
                foreach ($selector in @('selector1', 'selector2')) {
                    $selectorRecords += @(Resolve-DnsName -Name ("{0}._domainkey.{1}" -f $selector, $domainName) -Server 1.1.1.1 -Type CNAME -ErrorAction SilentlyContinue -verbose:$false)
                }

                $spfConfigured = $false
                $dmarcConfigured = $false
                $dkimSelectorCount = 0
                $dkimConfigured = $false

                $txtValues = @($txtRecords | ForEach-Object { @($_.Strings) -join '' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $spfRecord = $null
                $spfPolicyMode = $null
                if ($txtValues.Count -gt 0) {
                    $spfRecord = @($txtValues | Where-Object { $_ -match '(?i)^v=spf1' } | Select-Object -First 1)
                    if ($spfRecord.Count -gt 0) {
                        $spfRecord = $spfRecord[0]
                    } else {
                        $spfRecord = $null
                    }
                    $spfConfigured = @($txtValues | Where-Object { $_ -match '(?i)^v=spf1' -and $_ -match '(?i)include:spf\.protection\.outlook\.com' }).Count -gt 0
                    if ($spfRecord) {
                        if ($spfRecord -match '(?i)\s-all\b') { $spfPolicyMode = 'HardFail (-all)' }
                        elseif ($spfRecord -match '(?i)\s~all\b') { $spfPolicyMode = 'SoftFail (~all)' }
                        elseif ($spfRecord -match '(?i)\s\+all\b') { $spfPolicyMode = 'AllowAll (+all)' }
                        elseif ($spfRecord -match '(?i)\s\?all\b') { $spfPolicyMode = 'Neutral (?all)' }
                        else { $spfPolicyMode = 'Unspecified' }
                    }
                }

                $dmarcValues = @($dmarcRecords | ForEach-Object { @($_.Strings) -join '' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $dmarcRecord = $null
                $dmarcPolicy = $null
                $dmarcPct = $null
                if ($dmarcValues.Count -gt 0) {
                    $dmarcConfigured = @($dmarcValues | Where-Object { $_ -match '(?i)^v=dmarc1' }).Count -gt 0
                    $dmarcRecord = @($dmarcValues | Where-Object { $_ -match '(?i)^v=dmarc1' } | Select-Object -First 1)
                    if ($dmarcRecord.Count -gt 0) {
                        $dmarcRecord = $dmarcRecord[0]
                    } else {
                        $dmarcRecord = $null
                    }
                    if ($dmarcRecord) {
                        $policyMatch = [regex]::Match($dmarcRecord, '(?i)\bp=([a-z]+)')
                        if ($policyMatch.Success) {
                            $dmarcPolicy = $policyMatch.Groups[1].Value.ToLowerInvariant()
                        }
                        $pctMatch = [regex]::Match($dmarcRecord, '(?i)\bpct=(\d{1,3})')
                        if ($pctMatch.Success) {
                            $dmarcPct = [int]$pctMatch.Groups[1].Value
                        } else {
                            $dmarcPct = 100
                        }
                    }
                }

                $dkimSelectorCount = @($selectorRecords | Where-Object { $_.NameHost }).Count
                $dkimConfigured = ($dkimSelectorCount -ge 2)
                $selector1Host = @($selectorRecords | Where-Object { $_.Name -like 'selector1._domainkey*' -and $_.NameHost } | Select-Object -First 1 -ExpandProperty NameHost)
                $selector2Host = @($selectorRecords | Where-Object { $_.Name -like 'selector2._domainkey*' -and $_.NameHost } | Select-Object -First 1 -ExpandProperty NameHost)
                if ($selector1Host.Count -gt 0) { $selector1Host = $selector1Host[0] } else { $selector1Host = $null }
                if ($selector2Host.Count -gt 0) { $selector2Host = $selector2Host[0] } else { $selector2Host = $null }

                if ($NSRecords) {
                    $DNSCompanies =  Get-DNSHostCompanyName -nsRecords $NSRecords -ErrorAction SilentlyContinue
                }
                else {
                    $DNSCompanies = $null
                }
                Write-Log -Type DEBUG -Message "[Get-AllOffice365Domains] Completed Gathering DNS Records for '$($domainName)'"

                $DomainType = 'Unknown'
                $RecipientCounts = [PSCustomObject]@{
                    PrimarySMTPCount = 0
                    AliasOnlyCount = 0
                    TotalDomainRecipientsCount = 0
                }

                #Check for Exchange Online Dependencies
                if ($exchangeOnline -eq $true) {
                    Write-Log -Type DEBUG -Message "[Get-AllOffice365Domains] Gathering Exchange Online Domain Details for '$($domainName)'" -ExportFileLocation $ExportDetails
                    $domainTypeValue = @($acceptedDomains | Where-Object { $_.DomainName -eq $domainName } | Select-Object -First 1 -ExpandProperty "DomainType")
                    if ($domainTypeValue.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$domainTypeValue[0])) {
                        $DomainType = [string]$domainTypeValue[0]
                    }
                    else {
                        $DomainType = 'Unknown'
                    }
                    #Domain Dependencies
                    $RecipientCounts = Get-RecipientCounts -Recipients $recipients -DomainName $domainName
                }
                else { Write-Log -Type Error -Message "[Get-AllOffice365Domains] Skipping Exchange Online Domain Details for '$($domainName)'" -ExportFileLocation $ExportDetails }
                # Add to the results array
                Write-Log -Type INFO -Message "[Get-AllOffice365Domains] Adding '$($domainName)' to results array" -ExportFileLocation $ExportDetails
                $currentDomain = [PSCustomObject]@{
                    # ===== Domain Identity =====
                    Domain             = $domainName
                    DomainType         = $DomainType
                    Verified           = $domainVerified
                    IsDefault          = $domain.IsDefault
                    AuthenticationType = $AuthenticationType
            
                    # ===== DNS Records =====
                    DNSCompanies           = $DNSCompanies -join ","
                    NSRecords              = ($NSRecords.NameHost -join ",") -replace "`n|`r",     ""
                    ARecords               = ($aRecords.IPAddress -join ",") -replace "`n|`r",     ""
                    MXRecords              = ($mxRecords.NameExchange -join ",") -replace "`n|`r", ""
                    Office365MailExchanger = ($mxRecords.NameExchange -join "," -like "*protection.outlook.com")
                    SpfRecord              = $spfRecord
                    SpfIncludesM365        = $spfConfigured
                    SpfPolicyMode          = $spfPolicyMode
                    DmarcRecord            = $dmarcRecord
                    DmarcConfigured        = $dmarcConfigured
                    DmarcPolicy            = $dmarcPolicy
                    DmarcPercent           = $dmarcPct
                    DkimConfigured         = $dkimConfigured
                    DkimSelectorsConfigured = $dkimSelectorCount
                    DkimSelector1          = $selector1Host
                    DkimSelector2          = $selector2Host
                    DkimSelectorRecords    = (($selectorRecords | Where-Object { $_.NameHost } | Select-Object -ExpandProperty NameHost) -join ",")
            
                    # ===== Recipient Counts =====
                    PrimarySMTPRecipients = $RecipientCounts.PrimarySMTPCount
                    AliasOnlyRecipients   = $RecipientCounts.AliasOnlyCount
                    TotalDomainRecipients = $RecipientCounts.TotalDomainRecipientsCount
                }
                $script:tenantStatsHash["Domains"][$domainName] = $currentDomain
            }
            catch {
                Write-Log -Type ERROR -Message "An error occurred in running Get-AllOffice365Domains function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
        }
        foreach ($domain in $remoteDomains) {
            # Add to the results Hash Table
            $script:tenantStatsHash["RemoteDomains"][$domain.Identity] = $domain
        }
    }
    catch {
        Write-Log -Type ERROR -Message "An error occurred in running Get-AllOffice365Domains function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-ProgressHelper -Total $totalCount -Id 1 -Activity "Gathering Domain Details" -Completed
        Write-Log -Type INFO -Message "[Get-AllOffice365Domains] COMPLETED: Gathering All Domain Details" -ExportFileLocation $ExportDetails
    }
}
# ----------------------------------
# Combine Mailbox Details Specific Functions
# ----------------------------------

# Function to combine all user and mailbox statistics - User Priority and Mailbox Priority
function Report-UserAndMailboxStats {
    [CmdletBinding()]
    param ()

    $logPerRecordDebug = Test-ShowCollectorDiagnostics
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $tenantStatsStore = $script:tenantStatsHash

    function Get-LookupTable {
        param(
            [Parameter(Mandatory = $true)]
            [hashtable]$TenantStatsStore,
            [Parameter(Mandatory = $true)]
            [string]$Key
        )

        if (-not $TenantStatsStore.ContainsKey($Key)) {
            return $null
        }

        $candidate = $TenantStatsStore[$Key]
        if ($candidate -is [System.Collections.IDictionary]) {
            return $candidate
        }

        return $null
    }

    function Get-DictionaryValue {
        param(
            [Parameter(Mandatory = $false)]
            [object]$Dictionary,
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Key
        )

        if ($null -eq $Dictionary -or $null -eq $Key) {
            return $null
        }

        if ($Dictionary -is [System.Collections.IDictionary]) {
            if ($Dictionary.Contains($Key)) {
                return $Dictionary[$Key]
            }
            return $null
        }

        return $null
    }

    function Resolve-StatsSizeGb {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Stats,
            [Parameter(Mandatory = $true)]
            [string]$CachePrefix,
            [Parameter(Mandatory = $true)]
            [hashtable]$Cache
        )

        if ($null -eq $Stats) {
            return 0
        }

        $cacheKey = $null
        if ($Stats.PSObject.Properties['MailboxGuid'] -and $Stats.MailboxGuid) {
            $cacheKey = "{0}:{1}" -f $CachePrefix, [string]$Stats.MailboxGuid
        }

        if ($cacheKey -and $Cache.ContainsKey($cacheKey)) {
            return $Cache[$cacheKey]
        }

        $sizeBytes = [int64]0
        if ($Stats.PSObject.Properties['TotalItemSizeBytes'] -and $Stats.TotalItemSizeBytes) {
            $sizeBytes = [int64]$Stats.TotalItemSizeBytes
        }
        elseif ($Stats.PSObject.Properties['TotalItemSize'] -and $Stats.TotalItemSize) {
            $sizeBytes = Convert-DataSizeToBytes -Value $Stats.TotalItemSize
        }

        $sizeGb = [math]::Round(($sizeBytes / 1GB), 3)
        if ($cacheKey) {
            $Cache[$cacheKey] = $sizeGb
        }

        return $sizeGb
    }

    function Populate-Details {
        param (
            [Parameter(Mandatory = $true)]
            [object]$entity,
            [Parameter(Mandatory = $false)]
            [switch]$IsMailbox
        )

        $detailsMap = [ordered]@{}
        try {
            foreach ($property in $entity.PSObject.Properties) {
                $detailsMap[$property.Name] = $property.Value
            }

            $mailboxDetails = $null
            if ($IsMailbox -and $entity.RecipientTypeDetails -eq "GroupMailbox" -and $lookupContext.UnifiedGroups) {
                $mailboxDetails = Get-DictionaryValue -Dictionary $lookupContext.UnifiedGroups -Key $entity.PrimarySMTPAddress
                if ($mailboxDetails) {
                    Write-Verbose "Unified Group Details Found: $($entity.PrimarySMTPAddress)"
                }
            }
            elseif ($IsMailbox -and $lookupContext.AllMailboxesByIdentity) {
                $mailboxDetails = Get-DictionaryValue -Dictionary $lookupContext.AllMailboxesByIdentity -Key $entity.Identity
                if ($mailboxDetails) {
                    Write-Verbose "Mailbox Details Found: $($entity.Identity)"
                }
            }

            if (-not $mailboxDetails -and $lookupContext.AllMailboxesByUserPrincipalName -and $entity.UserPrincipalName) {
                $mailboxDetails = Get-DictionaryValue -Dictionary $lookupContext.AllMailboxesByUserPrincipalName -Key $entity.UserPrincipalName
                if ($mailboxDetails) {
                    Write-Verbose "Mailbox Details Found (UPN): $($entity.UserPrincipalName)"
                }
            }

            if (-not $mailboxDetails) {
                Write-Verbose "Mailbox Details Not Found"
            }

            $mailboxStats = $null
            if ($mailboxDetails -and $lookupContext.PrimaryMailboxStats) {
                $mailboxStats = Get-MailboxStatForRecord -MailboxRecord $mailboxDetails -StatsHash $lookupContext.PrimaryMailboxStats
                if ($mailboxStats) {
                    Write-Verbose "Mailbox Stats Found for '$($mailboxDetails.DisplayName)'"
                }
            }
            if (-not $mailboxStats) {
                Write-Verbose "Mailbox Stats Not Found"
            }

            $archiveStats = $null
            if ($mailboxDetails -and $lookupContext.ArchiveMailboxStats -and $mailboxDetails.ArchiveGuid) {
                $archiveStats = Get-DictionaryValue -Dictionary $lookupContext.ArchiveMailboxStats -Key (Convert-ToMailboxGuidKey -GuidValue $mailboxDetails.ArchiveGuid)
                if ($archiveStats) {
                    Write-Verbose "Archive Stats Found: $($mailboxDetails.ArchiveGuid.ToString())"
                }
            }
            if (-not $archiveStats) {
                Write-Verbose "Archive Stats Not Found"
            }

            $driveData = $null
            if ($IsMailbox -and $entity.RecipientTypeDetails -eq "GroupMailbox" -and $lookupContext.SharePoint -and $mailboxDetails -and $mailboxDetails.SharePointSiteUrl) {
                $driveData = Get-DictionaryValue -Dictionary $lookupContext.SharePoint -Key $mailboxDetails.SharePointSiteUrl
                if ($driveData) {
                    Write-Verbose "SharePoint Details Found: $($mailboxDetails.SharePointSiteUrl)"
                }
            }
            elseif ($IsMailbox -and $lookupContext.OneDrive -and $mailboxDetails -and $mailboxDetails.UserPrincipalName) {
                $driveData = Get-DictionaryValue -Dictionary $lookupContext.OneDrive -Key $mailboxDetails.UserPrincipalName
                if ($driveData) {
                    Write-Verbose "OneDrive Details Found: $($mailboxDetails.UserPrincipalName)"
                }
            }
            elseif ($lookupContext.OneDrive -and $entity.UserPrincipalName) {
                $driveData = Get-DictionaryValue -Dictionary $lookupContext.OneDrive -Key $entity.UserPrincipalName
                if ($driveData) {
                    Write-Verbose "OneDrive Details Found: $($entity.UserPrincipalName)"
                }
            }

            $MBXSizeGB = Resolve-StatsSizeGb -Stats $mailboxStats -CachePrefix 'MBX' -Cache $statsSizeCache
            $MBXItemCount = if ($mailboxStats -and $mailboxStats.ItemCount) { $mailboxStats.ItemCount } else { 0 }
            $ArchiveSizeGB = Resolve-StatsSizeGb -Stats $archiveStats -CachePrefix 'ARC' -Cache $statsSizeCache
            $ArchiveItemCount = if ($archiveStats -and $archiveStats.ItemCount) { $archiveStats.ItemCount } else { 0 }
            $DriveURL = if ($driveData -and $driveData.URL) { $driveData.URL } else { $null }
            $DriveStorageGB = if ($driveData -and $driveData.StorageUsageCurrent) { [math]::Round($driveData.StorageUsageCurrent / 1024, 3) } else { 0 }

            $detailsMap["MBXSizeGB"] = $MBXSizeGB
            $detailsMap["MBXItemCount"] = $MBXItemCount
            $detailsMap["ArchiveSizeGB"] = $ArchiveSizeGB
            $detailsMap["ArchiveItemCount"] = $ArchiveItemCount
            $detailsMap["DriveURL"] = $DriveURL
            $detailsMap["DriveStorageGB"] = $DriveStorageGB
        }
        catch {
            Write-Log -Type ERROR -Message "[Populate-Details] An error occurred in Populating Details for $($entity.PrimarySMTPAddress). $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }

        return [PSCustomObject]$detailsMap
    }

    $lookupContext = [ordered]@{
        UnifiedGroups                  = Get-LookupTable -TenantStatsStore $tenantStatsStore -Key 'UnifiedGroups'
        AllMailboxesByIdentity         = Get-LookupTable -TenantStatsStore $tenantStatsStore -Key 'AllMailboxes-MailIdentity'
        AllMailboxesByUserPrincipalName = Get-LookupTable -TenantStatsStore $tenantStatsStore -Key 'AllMailboxes-UserPrincipalName'
        PrimaryMailboxStats            = Get-LookupTable -TenantStatsStore $tenantStatsStore -Key 'PrimaryMailboxStats'
        ArchiveMailboxStats            = Get-LookupTable -TenantStatsStore $tenantStatsStore -Key 'ArchiveMailboxStats'
        SharePoint                     = Get-LookupTable -TenantStatsStore $tenantStatsStore -Key 'SharePoint'
        OneDrive                       = Get-LookupTable -TenantStatsStore $tenantStatsStore -Key 'OneDrive'
    }

    $statsSizeCache = @{}
    
    # Combine User and Mailbox Stats with Progress
    function Combine-UserAndMailboxStats {
        param (
            [hashtable]$TenantStatsStore
        )

        # Initialize variables. Include the start time for the script and logging
        $start = Get-Date
        if (-not $TenantStatsStore) {
            throw "[Combine-UserAndMailboxStats] TenantStatsStore was not provided."
        }
        # Create hash tables for user and mailbox details
        $TenantStatsStore["UserFullDetails"] = @{}
        $TenantStatsStore["MailboxFullDetails"] = @{}
        $combinedUserProgressId = 81
        $combinedMailboxProgressId = 82
        
        # Process Users
        $usersTable = Get-LookupTable -TenantStatsStore $TenantStatsStore -Key 'Users'
        $userEntries = @()
        if ($usersTable -is [System.Collections.IDictionary]) {
            $userEntries = @($usersTable.GetEnumerator())
        }
        $userCount = $userEntries.Count

        foreach ($userEntry in $userEntries) {
            $userKey = $userEntry.Key
            $user = $userEntry.Value
            if ($logPerRecordDebug) {
                Write-Log -Type DEBUG -Message ("[Combine-UserAndMailboxStats] Combining User '{0}' Details" -f $user.DisplayName) -ExportFileLocation $ExportDetails
            }

            Write-ProgressHelper -Total ([Math]::Max($userCount, 1)) -Id $combinedUserProgressId -Activity "Processing User Data" -Operation "Processing user: $($user.DisplayName)"

            $userDetails = Populate-Details -entity $user
            $userDetailsKey = @(
                [string]$user.UserPrincipalName
                [string]$user.PrimarySmtpAddress
                [string]$user.ExternalDirectoryObjectId
                [string]$user.Id
                [string]$userKey
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
            if ([string]::IsNullOrWhiteSpace($userDetailsKey)) {
                $userDetailsKey = "user:$([guid]::NewGuid().Guid)"
            }
            $TenantStatsStore["UserFullDetails"][$userDetailsKey] = $userDetails
        }
        Write-ProgressHelper -Total ([Math]::Max($userCount, 1)) -Id $combinedUserProgressId -Activity "Processing User Data" -Completed

        # Process Mailboxes
        $allRecipientRows = @()
        $allRecipientsTable = Get-LookupTable -TenantStatsStore $TenantStatsStore -Key 'AllRecipients'
        if ($allRecipientsTable -is [System.Collections.IDictionary]) {
            $allRecipientRows = @($allRecipientsTable.Values)
        }
        elseif ($null -ne $TenantStatsStore['AllRecipients'] -and $TenantStatsStore['AllRecipients'] -is [System.Collections.IEnumerable] -and -not ($TenantStatsStore['AllRecipients'] -is [string])) {
            $allRecipientRows = @($TenantStatsStore['AllRecipients'])
        }

        if ($allRecipientRows.Count -eq 0) {
            $allMailboxesTable = Get-LookupTable -TenantStatsStore $TenantStatsStore -Key 'AllMailboxes'
            if ($allMailboxesTable -is [System.Collections.IDictionary]) {
                $allRecipientRows = @($allMailboxesTable.Values)
            }
        }

        $allMailboxes = @($allRecipientRows | Where-Object {
            $_ -and $_.PSObject -and $_.PSObject.Properties['RecipientTypeDetails'] -and [string]$_.RecipientTypeDetails -like "*Mailbox"
        })
        if ($allMailboxes.Count -eq 0 -and $allRecipientRows.Count -gt 0) {
            $allMailboxes = @($allRecipientRows)
        }
        $mailboxCount = $allMailboxes.Count
        Write-Log -Type INFO -Message "[Combine-UserAndMailboxStats] Combining User and Mailbox Details - Processing Mailboxes" -ExportFileLocation $ExportDetails
        foreach ($mailbox in $allMailboxes) {
            #$mailbox = $TenantStatsStore["AllMailboxes"][$mailbox.PrimarySMTPAddress]
            if ($logPerRecordDebug) {
                Write-Log -Type DEBUG -Message ("[Combine-UserAndMailboxStats] Combining Mailbox '{0}' Details" -f $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails
            }
            Write-ProgressHelper -Total ([Math]::Max($mailboxCount, 1)) -Id $combinedMailboxProgressId -Activity "Processing Mailbox Data" -Operation "Processing mailbox: $($mailbox.PrimarySMTPAddress)"

            $mailboxDetails = Populate-Details -entity $mailbox -IsMailbox
            $mailboxDetailsKey = @(
                [string]$mailbox.PrimarySMTPAddress
                [string]$mailbox.PrimarySmtpAddress
                [string]$mailbox.UserPrincipalName
                [string]$mailbox.ExternalDirectoryObjectId
                [string]$mailbox.Guid
                [string]$mailbox.Identity
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
            if ([string]::IsNullOrWhiteSpace($mailboxDetailsKey)) {
                $mailboxDetailsKey = "mailbox:$([guid]::NewGuid().Guid)"
            }
            $TenantStatsStore["MailboxFullDetails"][$mailboxDetailsKey] = $mailboxDetails
        }

        Write-ProgressHelper -Total ([Math]::Max($mailboxCount, 1)) -Id $combinedMailboxProgressId -Activity "Processing Mailbox Data" -Completed
    }
    # New function to generate Inactive Mailboxes report
    function Report-InactiveMailboxes {
        [CmdletBinding()]
        param (
            [hashtable]$TenantStatsStore
        )

        # Initialize variables. Include the start time for the script and logging
        $start = Get-Date
        if (-not $TenantStatsStore) {
            throw "[Report-InactiveMailboxes] TenantStatsStore was not provided."
        }

        # Create hash table for inactive mailbox details
        $TenantStatsStore["InactiveMailboxDetails"] = @{}
        $inactiveMailboxProgressId = 83

        # Process Inactive Mailboxes
        $inactiveMailboxes = @()
        $inactiveMailboxTable = Get-LookupTable -TenantStatsStore $TenantStatsStore -Key 'InactiveMailboxes'
        if ($inactiveMailboxTable -is [System.Collections.IDictionary]) {
            $inactiveMailboxes = @($inactiveMailboxTable.Values)
        }
        elseif ($TenantStatsStore.ContainsKey('AllMailboxes') -and $TenantStatsStore['AllMailboxes'] -is [System.Collections.IDictionary]) {
            $inactiveMailboxes = @(
                $TenantStatsStore['AllMailboxes'].Values | Where-Object {
                    $_ -and $_.PSObject -and $_.PSObject.Properties['IsInactiveMailbox'] -and $_.IsInactiveMailbox -eq $true
                }
            )
        }
        $inactiveMailboxCount = $inactiveMailboxes.Count
        Write-Log -Type INFO -Message "[Report-InactiveMailboxes] Processing Inactive Mailboxes" -ExportFileLocation $ExportDetails
        foreach ($mailbox in $inactiveMailboxes) {
            if ($logPerRecordDebug) {
                Write-Log -Type DEBUG -Message ("[Report-InactiveMailboxes] Processing Inactive Mailbox '{0}'" -f $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails
            }
            Write-ProgressHelper -Total ([Math]::Max($inactiveMailboxCount, 1)) -Id $inactiveMailboxProgressId -Activity "Processing Inactive Mailbox Data" -Operation "Processing inactive mailbox: $($mailbox.PrimarySMTPAddress)"

            $existingMailboxDetails = $null
            if ($TenantStatsStore.ContainsKey("MailboxFullDetails") -and $TenantStatsStore["MailboxFullDetails"]) {
                $existingMailboxDetails = Get-DictionaryValue -Dictionary $TenantStatsStore["MailboxFullDetails"] -Key $mailbox.PrimarySMTPAddress
            }

            $mailboxDetails = if ($existingMailboxDetails) {
                $existingMailboxDetails
            } else {
                Populate-Details -entity $mailbox -IsMailbox
            }
            $TenantStatsStore["InactiveMailboxDetails"][$mailbox.PrimarySMTPAddress] = $mailboxDetails
        }

        Write-ProgressHelper -Total ([Math]::Max($inactiveMailboxCount, 1)) -Id $inactiveMailboxProgressId -Activity "Processing Inactive Mailbox Data" -Completed

        $completedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed Inactive Mailbox Report in $completedTime" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Report-InactiveMailboxes] COMPLETED: Processed Inactive Mailboxes in $completedTime" -ExportFileLocation $ExportDetails
    }

    # Combine user and mailbox stats
    Write-Log -Type INFO -Message "[Combine-UserAndMailboxStats] Combining User and Mailbox Details" -ExportFileLocation $ExportDetails
    try {
        $start = Get-Date
        Write-Host "Combining User and Mailbox Details..." -ForegroundColor Cyan -NoNewline
        Combine-UserAndMailboxStats -TenantStatsStore $tenantStatsStore
        Write-Host "Generating Inactive Mailbox Report..." -ForegroundColor Cyan -NoNewline
        Report-InactiveMailboxes -TenantStatsStore $tenantStatsStore
    } catch {
        Write-Log -Type ERROR -Message "[Combine-UserAndMailboxStats] An error occurred while combining User and Mailbox Details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    } finally {
        $completedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $completedTime" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Combine-UserAndMailboxStats] COMPLETED: Combined User and Mailbox Details in $completedTime" -ExportFileLocation $ExportDetails
    }
}



# ----------------------------------
# Entra Report Details Specific Functions
# ----------------------------------

# Device Report Details
function Get-AllDevicesReport {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel
    )

    # Initialize the tenant statistics hash table. Include the start time for the script and logging
    $start = Get-Date
    $deviceProgressId = 73
    $deviceProgressTotal = 1
        # Ensure global hash table structure
        if (-not $script:tenantStatsHash) {
            $script:tenantStatsHash = @{}
        }
    $script:tenantStatsHash["DeviceDetails"] = @{}

    Write-Host "Getting Device Details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllDevicesReport] START: Gathering all Device with $($detailLevel) details" -ExportFileLocation $ExportDetails
    
    try {
        $savedProgressPreference = $ProgressPreference
        try {
            $ProgressPreference = 'SilentlyContinue'
            $devices = Get-MgDevice -All -ProgressAction SilentlyContinue -ErrorAction Stop | ? {$null -ne $_.ID}
        }
        finally {
            $ProgressPreference = $savedProgressPreference
        }
        $DesiredProperties = @(
            "DisplayName", "AccountEnabled", "OperatingSystem", "OperatingSystemVersion", "DeviceId", "ID"
            "ApproximateLastSignInDateTime", "TrustType", "DirSyncEnabled", "LastDirSyncTime"
            "IsCompliant", "IsManaged", "ProfileType"
            )

        # Filter for Desired Attributes
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $devices = $devices | Select $DesiredProperties
            }
            geek {$devices = $devices}
        }

        # Add Additional Properties - MDM Solution, Join Type, Stale Device
        $totalCount = $devices.count
        $deviceProgressTotal = [Math]::Max($totalCount, 1)
        foreach ($device in $devices) {
            Write-ProgressHelper -Total $deviceProgressTotal -Id $deviceProgressId -Activity "Processing all Found Devices" -Operation "Updating Device Details for $($device.DisplayName)"
            try {
                #Add MDM Solution
                if ($device.ID) {
                    $device | Add-Member -MemberType NoteProperty -Name "ObjectID" -Value $device.ID -Force
                }
                #Add MDM Solution
                    if ($device.IsManaged -eq $true) {
                        $device | Add-Member -MemberType NoteProperty -Name "MDMSolution" -Value "Intune or SCCM" -Force
                    } else {
                        $device | Add-Member -MemberType NoteProperty -Name "MDMSolution" -Value "Not Managed" -Force
                    }
                Write-Log -Type DEBUG -Message ("[Get-AllDevicesReport] Adding Device Details for '{0}': Added {1} MDM Solution value" -f $device.DisplayName, $device.MDMSolution) -ExportFileLocation $ExportDetails
        
                #Add Join Type
                switch ($device.DeviceTrustType) {
                    AzureAD {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceJoinType" -Value "Entra AD joined" -Force
                    }
                    WorkPlace {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceJoinType" -Value "Entra AD registered" -Force
                    }
                    ServerAd {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceJoinType" -Value "Hybrid Entra AD joined" -Force
                    }
                    Default {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceJoinType" -Value "Unknown" -Force
                    }
                }
                Write-Log -Type DEBUG -Message ("[Get-AllDevicesReport] ({0}/{1}) Adding Device Details for '{2}': Added {3} Join Type value" -f $progressCounter, $devices.count, $device.DisplayName, $device.MDMSolution) -ExportFileLocation $ExportDetails
        
                #Add if Stale Device (Approximate login time older than 6 months)
                $lastLogin = $device.ApproximateLastSignInDateTime
                if ($lastLogin) {
                    $sixMonthsAgo = (Get-Date).AddMonths(-6)
                    $timeSinceLastLogon = $sixMonthsAgo - $lastLogin
                    $device | Add-Member -MemberType NoteProperty -Name "DaysDeviceInactiveFrom6MonthsAgo" -Value $timeSinceLastLogon.Days -Force
                    if ($lastLogin -le $sixMonthsAgo) {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceStale" -Value $true -Force
                    } else {
                        $device | Add-Member -MemberType NoteProperty -Name "DeviceStale" -Value $false -Force
                    }
                }
                elseif ($lastLogin) {
                    $device | Add-Member -MemberType NoteProperty -Name "DaysDeviceInactiveFrom6MonthsAgo" -Value "N/A" -Force
                    $device | Add-Member -MemberType NoteProperty -Name "DeviceStale" -Value "N/A" -Force
                }
                Write-Log -Type DEBUG -Message ("[Get-AllDevicesReport] ({0}/{1}) Adding Device Details for '{2}': Added {3} Device Stale value" -f $progressCounter, $devices.count, $device.DisplayName, $device.MDMSolution) -ExportFileLocation $ExportDetails
                
                #Add Compliant State
                if ($null -eq $device.IsCompliant) {
                    $device | Add-Member -MemberType NoteProperty -Name "IsCompliant" -Value "NotApplied" -Force
                }
                Write-Log -Type DEBUG -Message ("[Get-AllDevicesReport] ({0}/{1}) Adding Device Details for '{2}': Added {3} Device Compliant value" -f $progressCounter, $devices.count, $device.DisplayName, $device.MDMSolution) -ExportFileLocation $ExportDetails

            }
            catch {
                Write-Log -Type Error -Message ("[Get-AllDevicesReport] ({0}/{1}) Failed to add Device Details for '{2}'. $($_.Exception.Message)" -f $progressCounter, $devices.count, $device.DisplayName) -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
            finally {
                #Add to Hash Table
                $script:tenantStatsHash["DeviceDetails"][$device.ObjectID] = $device
            }
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllDevicesReport] An error occurred in running Get-AllDeviceReport function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-ProgressHelper -Total $deviceProgressTotal -Id $deviceProgressId -Activity "Processing all Found Devices" -Completed
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllDevicesReport] COMPLETED: Gathering all Entra Devices with $($detailLevel) details" -ExportFileLocation $ExportDetails
    }   
}

# Function to Retrieve Conditional Access Policies from Microsoft Graph
function Get-ConditionalAccessPoliciesReport {
    param (
        [Parameter(Mandatory = $True, HelpMessage = 'Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel
    )

    $start = Get-Date
    $conditionalAccessProgressId = 74
    $conditionalAccessProgressTotal = 1

    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["ConditionalAccessPolicies"] = @{}
    $script:tenantStatsHash["ConditionalAccessPolicySummary"] = @{}

    Write-Host "Getting Entra Conditional Access Policies Details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-ConditionalAccessPoliciesReport] START: Gathering all Entra Conditional Access Policies  with $($detailLevel) details" -ExportFileLocation $ExportDetails

    function Get-PolicyItemCounts {
        param (
            [object]$policyCondition
        )
        if ($null -eq $policyCondition) {
            return 0
        } elseif ($policyCondition -eq "All") {
            return "All"
        } else {
            return ($policyCondition | Measure-Object).Count
        }
    }

    try {
        # Use Get-MgIdentityConditionalAccessPolicy as primary cmdlet
        $savedProgressPreference = $ProgressPreference
        try {
            $ProgressPreference = 'SilentlyContinue'
            $conditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy -All -ProgressAction SilentlyContinue -ErrorAction Stop
        }
        finally {
            $ProgressPreference = $savedProgressPreference
        }

        switch ($detailLevel) {
            { $_ -in "minimum", "combined", "all" } {
                $DesiredProperties = @(
                    "Id", "DisplayName", "CreatedDateTime", "ModifiedDateTime", "Description", "State",
                    "Conditions", "GrantControls", "SessionControls"
                )
                $conditionalAccessPolicies = $conditionalAccessPolicies | Select-Object $DesiredProperties
            }
        }

        $totalCount = $conditionalAccessPolicies.Count
        $conditionalAccessProgressTotal = [Math]::Max($totalCount, 1)

        $summaryRows = New-Object System.Collections.Generic.List[object]
        foreach ($policy in $conditionalAccessPolicies) {
            Write-ProgressHelper -Total $conditionalAccessProgressTotal -Id $conditionalAccessProgressId -Activity "Processing all Conditional Access Policies" -Operation "Expanding Policy Details for $($policy.DisplayName)"
            Write-Log -Type INFO -Message "[Get-ConditionalAccessPoliciesReport] Expanding Conditional Access Policies for $($policy.DisplayName)" -ExportFileLocation $ExportDetails

            $includeGuestsOrExternalUsers = $null
            $excludeGuestsOrExternalUsers = $null
            if ($policy.Conditions.Users) {
                if ($policy.Conditions.Users.PSObject.Properties['IncludeGuestsOrExternalUsers']) {
                    $includeGuestsOrExternalUsers = $policy.Conditions.Users.IncludeGuestsOrExternalUsers
                }
                if ($policy.Conditions.Users.PSObject.Properties['ExcludeGuestsOrExternalUsers']) {
                    $excludeGuestsOrExternalUsers = $policy.Conditions.Users.ExcludeGuestsOrExternalUsers
                }
            }

            $clientAppTypes = @($policy.Conditions.ClientAppTypes)
            $grantControlsBuiltIn = @($policy.GrantControls.BuiltInControls)
            $includeUsers = @($policy.Conditions.Users.IncludeUsers)
            $excludeUsers = @($policy.Conditions.Users.ExcludeUsers)
            $includeGroups = @($policy.Conditions.Users.IncludeGroups)
            $excludeGroups = @($policy.Conditions.Users.ExcludeGroups)
            $includeRoles = @($policy.Conditions.Users.IncludeRoles)
            $excludeRoles = @($policy.Conditions.Users.ExcludeRoles)
            $includeLocations = @($policy.Conditions.Locations.IncludeLocations)
            $excludeLocations = @($policy.Conditions.Locations.ExcludeLocations)
            $includeApps = @($policy.Conditions.Applications.IncludeApplications)
            $excludeApps = @($policy.Conditions.Applications.ExcludeApplications)
            $signInRiskInclude = @($policy.Conditions.SignInRiskLevels.IncludeLevels)
            $servicePrincipalRiskInclude = @($policy.Conditions.ServicePrincipalRiskLevels.IncludeLevels)

            $hasExclusions = (
                $excludeUsers.Count -gt 0 -or
                $excludeGroups.Count -gt 0 -or
                $excludeRoles.Count -gt 0 -or
                $excludeLocations.Count -gt 0 -or
                $excludeApps.Count -gt 0 -or
                ($null -ne $excludeGuestsOrExternalUsers)
            )
            $targetsGuestsOrExternal = (
                ($null -ne $includeGuestsOrExternalUsers) -or
                (($includeUsers -join ',') -match 'guest|external')
            )
            $targetsPrivilegedRoles = ($includeRoles.Count -gt 0 -or $excludeRoles.Count -gt 0)
            $blocksLegacyAuth = (($clientAppTypes -join ',') -match 'exchangeactivesync|other' -and (($grantControlsBuiltIn -join ',') -match 'block'))
            $requiresCompliantDevice = (($grantControlsBuiltIn -join ',') -match 'compliantdevice|domainjoineddevice')
            $usesRiskSignals = ($signInRiskInclude.Count -gt 0 -or $servicePrincipalRiskInclude.Count -gt 0)

            $policyDetailsHash = [ordered]@{
                PolicyID         = $policy.Id
                DisplayName      = $policy.DisplayName
                State            = $policy.State
                CreatedDateTime  = $policy.CreatedDateTime
                ModifiedDateTime = $policy.ModifiedDateTime
                Description      = $policy.Description

                # Grant Controls
                GrantControls_BuiltInControls  = $policy.GrantControls.BuiltInControls -join ','
                GrantControls_CustomControls   = $policy.GrantControls.CustomControls -join ','
                GrantControls_Operator         = $policy.GrantControls.Operator

                # Session Controls
                SessionControls_ApplicationEnforcedRestrictions = $(if ($policy.SessionControls) { $policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled } else { $null })
                SessionControls_PersistentBrowser               = $(if ($policy.SessionControls) { $policy.SessionControls.PersistentBrowser.IsEnabled } else { $null })
                SessionControls_PersistentBrowserMode           = $(if ($policy.SessionControls) { $policy.SessionControls.PersistentBrowser.Mode } else { $null })
                SessionControls_SignInFrequency                 = $(if ($policy.SessionControls) { $policy.SessionControls.SignInFrequency.IsEnabled } else { $null })
                FrequencyInterval                               = $(if ($policy.SessionControls) { $policy.SessionControls.SignInFrequency.FrequencyInterval } else { $null })

                # Conditions - Users, Groups, Roles
                IncludedUsers      = $policy.Conditions.Users.IncludeUsers -join ','
                ExcludedUsers      = $policy.Conditions.Users.ExcludeUsers -join ','
                IncludedUsersCount = Get-PolicyItemCounts $policy.Conditions.Users.IncludeUsers
                ExcludedUsersCount = Get-PolicyItemCounts $policy.Conditions.Users.ExcludeUsers

                IncludedGroups      = $policy.Conditions.Users.IncludeGroups -join ','
                ExcludedGroups      = $policy.Conditions.Users.ExcludeGroups -join ','
                IncludedGroupsCount = Get-PolicyItemCounts $policy.Conditions.Users.IncludeGroups
                ExcludedGroupsCount = Get-PolicyItemCounts $policy.Conditions.Users.ExcludeGroups

                IncludedRoles      = $policy.Conditions.Users.IncludeRoles -join ','
                ExcludedRoles      = $policy.Conditions.Users.ExcludeRoles -join ','
                IncludedRolesCount = Get-PolicyItemCounts $policy.Conditions.Users.IncludeRoles
                ExcludedRolesCount = Get-PolicyItemCounts $policy.Conditions.Users.ExcludeRoles

                # Platforms
                IncludePlatforms = $policy.Conditions.Platforms.IncludePlatforms -join ','
                ExcludePlatforms = $policy.Conditions.Platforms.ExcludePlatforms -join ','

                # Locations
                IncludeLocations = $policy.Conditions.Locations.IncludeLocations -join ','
                ExcludeLocations = $policy.Conditions.Locations.ExcludeLocations -join ','

                # Devices/Device States/Filter
                IncludeDeviceStates = $policy.Conditions.Devices.IncludeDeviceStates -join ','
                ExcludeDeviceStates = $policy.Conditions.Devices.ExcludeDeviceStates -join ','
                DeviceFilterMode    = $(if ($policy.Conditions.Devices.DeviceFilter) { $policy.Conditions.Devices.DeviceFilter.Mode } else { '' })
                DeviceFilterRule    = $(if ($policy.Conditions.Devices.DeviceFilter) { $policy.Conditions.Devices.DeviceFilter.Rule } else { '' })

                # Applications/User Actions
                IncludeApplications      = $policy.Conditions.Applications.IncludeApplications -join ','
                ExcludeApplications      = $policy.Conditions.Applications.ExcludeApplications -join ','
                IncludedApplicationsCount = Get-PolicyItemCounts $policy.Conditions.Applications.IncludeApplications
                ExcludedApplicationsCount = Get-PolicyItemCounts $policy.Conditions.Applications.ExcludeApplications

                IncludeUserActions = $policy.Conditions.Applications.IncludeUserActions -join ','

                # Client Apps
                ClientAppTypes = $policy.Conditions.ClientAppTypes -join ','

                # Risk Levels & Others
                SignInRiskLevels_IncludeLevels = $policy.Conditions.SignInRiskLevels.IncludeLevels -join ','
                SignInRiskLevels_ExcludeLevels = $policy.Conditions.SignInRiskLevels.ExcludeLevels -join ','
                ServicePrincipalRiskLevels_IncludeLevels = $policy.Conditions.ServicePrincipalRiskLevels.IncludeLevels -join ','
                ServicePrincipalRiskLevels_ExcludeLevels = $policy.Conditions.ServicePrincipalRiskLevels.ExcludeLevels -join ','
                InsiderRiskLevels = $policy.Conditions.InsiderRiskLevels -join ','

                # Top-level DeviceStates
                DeviceStates_IncludeDeviceStates = $(if ($policy.Conditions.DeviceStates) { $policy.Conditions.DeviceStates.IncludeDeviceStates -join ',' } else { '' })
                DeviceStates_ExcludeDeviceStates = $(if ($policy.Conditions.DeviceStates) { $policy.Conditions.DeviceStates.ExcludeDeviceStates -join ',' } else { '' })
                IncludeGuestsOrExternalUsers = $(if ($null -ne $includeGuestsOrExternalUsers) { ($includeGuestsOrExternalUsers | ConvertTo-Json -Compress -Depth 5) } else { '' })
                ExcludeGuestsOrExternalUsers = $(if ($null -ne $excludeGuestsOrExternalUsers) { ($excludeGuestsOrExternalUsers | ConvertTo-Json -Compress -Depth 5) } else { '' })
                IsReportOnly                = (($policy.State -as [string]) -match 'report')
                HasExclusions               = [bool]$hasExclusions
                TargetsGuestsOrExternalUsers = [bool]$targetsGuestsOrExternal
                TargetsPrivilegedRoles      = [bool]$targetsPrivilegedRoles
                BlocksLegacyAuth            = [bool]$blocksLegacyAuth
                RequiresCompliantDevice     = [bool]$requiresCompliantDevice
                UsesRiskSignals             = [bool]$usesRiskSignals
            }

            $script:tenantStatsHash["ConditionalAccessPolicies"][$policy.DisplayName] = $policyDetailsHash
            $summaryRows.Add([pscustomobject]@{
                DisplayName                 = $policy.DisplayName
                State                       = $policy.State
                IsEnabled                   = ([string]$policy.State).ToLowerInvariant() -eq 'enabled'
                IsReportOnly                = (($policy.State -as [string]) -match 'report')
                HasExclusions               = [bool]$hasExclusions
                TargetsGuestsOrExternalUsers = [bool]$targetsGuestsOrExternal
                TargetsPrivilegedRoles      = [bool]$targetsPrivilegedRoles
                BlocksLegacyAuth            = [bool]$blocksLegacyAuth
                RequiresCompliantDevice     = [bool]$requiresCompliantDevice
                UsesRiskSignals             = [bool]$usesRiskSignals
            }) | Out-Null
        }

        $enabledPolicies = @($summaryRows | Where-Object { $_.IsEnabled })
        $reportOnlyPolicies = @($summaryRows | Where-Object { $_.IsReportOnly })
        $policiesWithExclusions = @($summaryRows | Where-Object { $_.HasExclusions })
        $guestPolicies = @($summaryRows | Where-Object { $_.TargetsGuestsOrExternalUsers })
        $privilegedPolicies = @($summaryRows | Where-Object { $_.TargetsPrivilegedRoles })
        $legacyAuthPolicies = @($summaryRows | Where-Object { $_.BlocksLegacyAuth })
        $compliantDevicePolicies = @($summaryRows | Where-Object { $_.RequiresCompliantDevice })
        $riskPolicies = @($summaryRows | Where-Object { $_.UsesRiskSignals })

        $script:tenantStatsHash["ConditionalAccessPolicySummary"]["Summary"] = [pscustomobject]@{
            TotalPolicies                     = $summaryRows.Count
            EnabledPolicies                   = $enabledPolicies.Count
            ReportOnlyPolicies                = $reportOnlyPolicies.Count
            PoliciesWithExclusions            = $policiesWithExclusions.Count
            PoliciesTargetingGuestsOrExternal = $guestPolicies.Count
            PoliciesTargetingPrivilegedRoles  = $privilegedPolicies.Count
            PoliciesBlockingLegacyAuth        = $legacyAuthPolicies.Count
            PoliciesRequiringCompliantDevice  = $compliantDevicePolicies.Count
            PoliciesUsingRiskSignals          = $riskPolicies.Count
            HasGuestCoverage                  = ($guestPolicies.Count -gt 0)
            HasPrivilegedRoleCoverage         = ($privilegedPolicies.Count -gt 0)
            HasLegacyAuthProtection           = ($legacyAuthPolicies.Count -gt 0)
            HasCompliantDeviceRequirement     = ($compliantDevicePolicies.Count -gt 0)
            HasRiskBasedCoverage              = ($riskPolicies.Count -gt 0)
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-ConditionalAccessPoliciesReport] An error occurred in running Get-ConditionalAccessPoliciesReport function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-ProgressHelper -Total $conditionalAccessProgressTotal -Id $conditionalAccessProgressId -Activity "Processing all Conditional Access Policies" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-ConditionalAccessPoliciesReport] COMPLETED: Gathering all Entra Conditional Access Policies with $($detailLevel) details" -ExportFileLocation $ExportDetails
    }
}

# Secure Score Report Details
function Get-SecuritySecureScoreReport {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$false,HelpMessage='Should It Only Pull the Most Recent?')]
        [Switch]$MostRecent 
    )
    $start = Get-Date
    $secureScoreProgressId = 75
    $secureScoreProgressTotal = 1
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["SecuritySecureScore"] = @{}
    $script:tenantStatsHash["SecureScoreActions"] = @{}
    $depthPolicy = if ($script:CollectionDepthPolicy) {
        $script:CollectionDepthPolicy
    } else {
        Get-ArrayaCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($detailLevel.ToLowerInvariant()))
    }
    $collectSecureScoreMappings = ($depthPolicy.CollectSecureScoreMappings -eq $true)

    Write-Host "Getting Security Score Details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] START: Gathering all Security Score details" -ExportFileLocation $ExportDetails
    try {
        $controlProfileLookup = @{}
        if ($collectSecureScoreMappings) {
            try {
                $controlProfileUri = "https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles?`$select=id,title,maxScore,rank,controlCategory,service,actionUrl,remediation,description,threats,userImpact,implementationCost,tier"
                $controlProfiles = @(Get-ArrayaGraphResource -Uri $controlProfileUri -PageSize 250 -Activity 'Secure Score control profile mappings' -Headers $global:GraphHeaders)
                foreach ($profile in $controlProfiles) {
                    if ($profile.Id) {
                        $controlProfileLookup[$profile.Id] = [PSCustomObject]@{
                            Id                 = $profile.Id
                            Title              = $profile.Title
                            MaxScore           = $profile.MaxScore
                            Rank               = $profile.Rank
                            ControlCategory    = $profile.ControlCategory
                            Service            = $profile.Service
                            ActionUrl          = $profile.ActionUrl
                            Remediation        = $profile.Remediation
                            Description        = $profile.Description
                            Threats            = @($profile.Threats)
                            UserImpact         = $profile.UserImpact
                            ImplementationCost = $profile.ImplementationCost
                            Tier               = $profile.Tier
                        }
                    }
                }
                Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Loaded $($controlProfileLookup.Count) Secure Score control profile mappings" -ExportFileLocation $ExportDetails
                $controlProfiles = $null
            } catch {
                Write-Log -Type WARNING -Message "[Get-SecuritySecureScoreReport] Unable to load Secure Score control profiles for source mapping: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }
        } else {
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Skipping Secure Score control profile mappings for this output profile." -ExportFileLocation $ExportDetails
        }

        if ($MostRecent) {
            $secureScoreUri = "https://graph.microsoft.com/v1.0/security/secureScores?`$orderby=createdDateTime desc&`$top=1"
            $secureScoreResponse = $null
            if ((Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) -and (Get-MgContext -ErrorAction SilentlyContinue)) {
                $savedProgressPreference = $ProgressPreference
                try {
                    $ProgressPreference = 'SilentlyContinue'
                    $secureScoreResponse = Invoke-MgGraphRequest -Uri $secureScoreUri -Method GET -OutputType PSObject -ProgressAction SilentlyContinue -ErrorAction Stop
                }
                finally {
                    $ProgressPreference = $savedProgressPreference
                }
            }
            else {
                $secureScoreResponse = Invoke-QuietRestMethod -Parameters @{
                    Uri         = $secureScoreUri
                    Headers     = $global:GraphHeaders
                    Method      = 'Get'
                    ContentType = 'application/json'
                    ErrorAction = 'Stop'
                }
            }
            $secureScore = @($secureScoreResponse.value | Where-Object { $null -ne $_.ID } | Select-Object -First 1)
        } else {
            $secureScore = @(Get-ArrayaGraphResource -Uri 'https://graph.microsoft.com/v1.0/security/secureScores' -PageSize 250 -Activity 'Secure Score history' -Headers $global:GraphHeaders | Where-Object { $null -ne $_.ID })
        }

        # Add Additional Properties
        $totalCount = $secureScore.count
        $secureScoreProgressTotal = [Math]::Max($totalCount, 1)
        foreach ($score in $secureScore) {
            Write-ProgressHelper -Total $secureScoreProgressTotal -Id $secureScoreProgressId -Activity "Processing all Security Score Details" -Operation "Gathering Score Details for $($score.ID)"
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Gathering Score Details for $($score.ID)" -ExportFileLocation $ExportDetails
            #Create Current Security Score Hash Table
            $currentSecurityScores = [PSCustomObject]@{
                ID = $score.ID
                CreatedDateTime = $score.CreatedDateTime
                # Add Current Security Score Percentage
                SecurityScorePercentage = ("{0:F2}" -f (($score.currentScore / $score.maxScore) * 100))
                CurrentScore = $score.CurrentScore
                MaxScore = $score.MaxScore
                LicensedUserCount = $score.LicensedUserCount
                # Add Enabled Services
                EnabledServicesCount = ($score.EnabledServices | Measure-Object).Count
                EnabledServices = ($score.EnabledServices -join ",")
                VendorProvider = $score.VendorInformation.Provider
                VendorName = $score.VendorInformation.Vendor
            }
            
            # Add Comparitive Scores - TotalSeats, AllTenants
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Gathering Score Details for $($score.ID): Add Comparitive Scores" -ExportFileLocation $ExportDetails
            $comparativeScores = $score.AverageComparativeScores
            $TotalSeatsScoreFullDetails = $comparativeScores | Where-Object { $_.Basis -eq "TotalSeats" }
            $currentSecurityScores | Add-Member -MemberType NoteProperty -Name "SimilarSeatSizeRangeLowerValue" -Value $(if ($TotalSeatsScoreFullDetails) { $TotalSeatsScoreFullDetails["SeatSizeRangeLowerValue"] } else { $null })
            $currentSecurityScores | Add-Member -MemberType NoteProperty -Name "SimilarSeatSizeRangeUpperValue" -Value $(if ($TotalSeatsScoreFullDetails) { $TotalSeatsScoreFullDetails["SeatSizeRangeUpperValue"] } else { $null })
            foreach ($ComparisonScore in ($comparativeScores | Where-Object { $_ })) {
                if ($ComparisonScore.Basis -eq "TotalSeats") {
                    $ComparisonBasisName = "SimilarSizeOrg"
                } elseif ($ComparisonScore.Basis -eq "AllTenants") {
                    $ComparisonBasisName = "AllTenants"
                }
                Write-Log -Type DEBUG -Message "[Get-SecuritySecureScoreReport] Gathering $($ComparisonBasisName) ComparisonScore Details Details for $($score.ID)" -ExportFileLocation $ExportDetails
                $currentSecurityScores | Add-Member -MemberType NoteProperty -Name "$($ComparisonBasisName)_ComparitiveScore" -Value $ComparisonScore.ComparativeScore
                if ($detailLevel -in "all", "geek") {
                    $comparisonAdditional = @{}
                    if ($ComparisonScore.PSObject.Properties['AdditionalProperties'] -and $ComparisonScore.AdditionalProperties) {
                        $comparisonAdditional = $ComparisonScore.AdditionalProperties
                    }
                    elseif ($ComparisonScore.PSObject.Properties['AdditionalData'] -and $ComparisonScore.AdditionalData) {
                        $comparisonAdditional = $ComparisonScore.AdditionalData
                    }

                    foreach ($individualScore in @($comparisonAdditional.Keys)) {
                        Write-Log -Type DEBUG -Message "[Get-SecuritySecureScoreReport] Gathering $($ComparisonScore.Basis):$($individualScore) Score Details for $($score.ID): Add Comparitive Scores" -ExportFileLocation $ExportDetails
                        $currentSecurityScores | Add-Member -MemberType NoteProperty -Name "$($ComparisonBasisName)_$($individualScore)" -Value $comparisonAdditional[$individualScore]
                    }
                }
            }

            
            #Add to Hash Table
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Gathering Score Details for $($score.ID): Add Score to Tenant Stats Hash Table" -ExportFileLocation $ExportDetails
            $script:tenantStatsHash["SecuritySecureScore"][$score.ID] = $currentSecurityScores
        }

        $latestScore = @($secureScore | Sort-Object CreatedDateTime -Descending | Select-Object -First 1)
        if ($latestScore.Count -gt 0 -and $collectSecureScoreMappings) {
            $actionIndex = 0
            foreach ($controlScore in @($latestScore[0].ControlScores)) {
                $actionIndex++
                $controlName = $controlScore.ControlName
                if ([string]::IsNullOrWhiteSpace($controlName)) {
                    continue
                }

                $profile = $null
                if ($controlProfileLookup.ContainsKey($controlName)) {
                    $profile = $controlProfileLookup[$controlName]
                }

                $currentControlScore = 0
                try { $currentControlScore = [double]$controlScore.Score } catch { $currentControlScore = 0 }

                $maxControlScore = 0
                if ($null -ne $controlScore.MaxScore -and $controlScore.MaxScore -ne '') {
                    try { $maxControlScore = [double]$controlScore.MaxScore } catch { $maxControlScore = 0 }
                } elseif ($profile -and $null -ne $profile.MaxScore -and $profile.MaxScore -ne '') {
                    try { $maxControlScore = [double]$profile.MaxScore } catch { $maxControlScore = 0 }
                }

                $scoreGap = [math]::Round(($maxControlScore - $currentControlScore), 2)
                $implementationStatus = if ($maxControlScore -gt 0 -and $scoreGap -le 0) {
                    'Implemented'
                } elseif ($currentControlScore -gt 0) {
                    'Partially Implemented'
                } else {
                    'Not Implemented'
                }

                $threats = @()
                if ($profile -and $profile.Threats) {
                    $threats = @($profile.Threats)
                }

                $actionRow = [PSCustomObject]@{
                    SecureScoreSnapshotId = $latestScore[0].Id
                    ControlId             = $controlName
                    RecommendationTitle   = if ($profile -and $profile.Title) { $profile.Title } else { $controlName }
                    Status                = $implementationStatus
                    CurrentScore          = $currentControlScore
                    MaxScore              = $maxControlScore
                    ScoreGap              = $scoreGap
                    Rank                  = if ($profile -and $profile.Rank) { $profile.Rank } else { $controlScore.Rank }
                    Category              = if ($profile -and $profile.ControlCategory) { $profile.ControlCategory } else { $controlScore.ControlCategory }
                    Service               = if ($profile -and $profile.Service) { $profile.Service } else { $controlScore.Service }
                    ActionUrl             = if ($profile) { $profile.ActionUrl } else { $null }
                    Remediation           = if ($profile) { $profile.Remediation } else { $null }
                    Description           = if ($profile -and $profile.Description) { $profile.Description } else { $controlScore.Description }
                    Threats               = if ($threats.Count -gt 0) { $threats -join ', ' } else { $null }
                    UserImpact            = if ($profile) { $profile.UserImpact } else { $null }
                    ImplementationCost    = if ($profile) { $profile.ImplementationCost } else { $null }
                    Tier                  = if ($profile) { $profile.Tier } else { $null }
                    SourceMapping         = if ($profile) { 'Microsoft Secure Score Control Profile' } else { 'Secure Score Snapshot Only' }
                }

                $script:tenantStatsHash["SecureScoreActions"][("{0:D3}-{1}" -f $actionIndex, $controlName)] = $actionRow
            }
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Added $($script:tenantStatsHash['SecureScoreActions'].Count) Secure Score recommendation mappings" -ExportFileLocation $ExportDetails
        } elseif (-not $collectSecureScoreMappings) {
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Secure Score recommendation mappings were skipped for this output profile." -ExportFileLocation $ExportDetails
        }

        $latestScore = $null
        $secureScore = $null
        $controlProfileLookup = $null

    }
    catch {
        Write-Log -Type ERROR -Message "[Get-SecuritySecureScoreReport] An error occurred in running Get-SecuritySecureScoreReport function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Write-ProgressHelper -Total $secureScoreProgressTotal -Id $secureScoreProgressId -Activity "Processing all Security Score Details" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] COMPLETED: Gathering Security Score details" -ExportFileLocation $ExportDetails
    }   
}


# Function to check Authentication Methods and SSO Configuration
function Get-AuthenticationConfiguration {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel
    )
    
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["AuthenticationConfig"] = @{}
    $script:tenantStatsHash["AuthenticationConfigSummary"] = @{}
    $script:tenantStatsHash["AuthenticationMethods"] = @{}
    $script:tenantStatsHash["AuthenticationSSOApplications"] = @{}
    $script:tenantStatsHash["EnterpriseApplications"] = @{}
    $script:tenantStatsHash["EnterpriseApplicationSummary"] = @{}
    $script:tenantStatsHash["SecurityDefaultsPolicy"] = @{}
    $script:tenantStatsHash["GuestSignInSummary"] = @{}
    $script:tenantStatsHash["PrivilegedAccessSummary"] = @{}
    
    Write-Host "Checking Authentication and SSO Configuration ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] START: Checking Authentication Configuration" -ExportFileLocation $ExportDetails
    
    try {
        $depthPolicy = if ($script:CollectionDepthPolicy) {
            $script:CollectionDepthPolicy
        } else {
            Get-ArrayaCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($detailLevel.ToLowerInvariant()))
        }
        $collectSsoAppDetails = ($depthPolicy.CollectSsoApplicationDetails -eq $true)
        $collectExtendedIdentityTierB = ($depthPolicy.CollectExtendedGraphEnrichment -eq $true)

        # Get authentication methods policy
        $authMethodsPolicy = [PSCustomObject]@{
            MFAEnabled = $false
            MFAMethods = @()
            SSOEnabled = $false
            SSOApplications = @()
            FederatedDomains = @()
            PasswordlessMethods = @()
            DefaultUserCanCreateApps = $null
            PermissionGrantPoliciesAssigned = @()
            AdminConsentWorkflowEnabled = $null
            AdminConsentWorkflowReviewerCount = 0
        }
        
        # Check for federated domains (indicates SSO)
        $domains = $script:tenantStatsHash["Domains"].Values
        $federatedDomains = $domains | Where-Object {$_.AuthenticationType -eq "Federated"}
        
        if ($federatedDomains) {
            $authMethodsPolicy.SSOEnabled = $true
            $authMethodsPolicy.FederatedDomains = $federatedDomains.Domain
            Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Found $($federatedDomains.Count) federated domains" -ExportFileLocation $ExportDetails
        }
        
        # Get Enterprise Applications configured for SSO
        if ($collectSsoAppDetails) {
            try {
                Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Checking Enterprise Applications for SSO" -ExportFileLocation $ExportDetails

                $ssoAppDetails = New-Object System.Collections.Generic.List[object]
                foreach ($app in (Get-MgServicePrincipal -All -Filter "tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')" -ProgressAction SilentlyContinue -ErrorAction SilentlyContinue)) {
                    if ($null -eq $app.PreferredSingleSignOnMode -or $app.PreferredSingleSignOnMode -eq "notSupported") {
                        continue
                    }
                    $ssoAppDetails.Add([PSCustomObject]@{
                        DisplayName = $app.DisplayName
                        AppId = $app.AppId
                        SSOMode = $app.PreferredSingleSignOnMode
                        ServicePrincipalType = $app.ServicePrincipalType
                        AccountEnabled = $app.AccountEnabled
                    })
                }

                if ($ssoAppDetails.Count -gt 0) {
                    $authMethodsPolicy.SSOEnabled = $true
                    $authMethodsPolicy.SSOApplications = $ssoAppDetails.ToArray()
                    Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Found $($ssoAppDetails.Count) SSO-enabled applications" -ExportFileLocation $ExportDetails
                }
                
            } catch {
                Write-Log -Type WARNING -Message "[Get-AuthenticationConfiguration] Unable to retrieve SSO applications: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }
        } else {
            Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Skipping detailed SSO application inventory in minimum reporting mode." -ExportFileLocation $ExportDetails
        }
        
        # Check MFA/Authentication methods policies
        try {
            $authMethodsPolicyUri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy"
            $authPolicyData = Office365Custom\Get-GraphData -Uri $authMethodsPolicyUri -Activity "Fetching Authentication Methods Policy"
            
            if ($authPolicyData) {
                # Extract enabled authentication methods
                foreach ($method in $authPolicyData.authenticationMethodConfigurations) {
                    if ($method.state -eq "enabled") {
                        $authMethodsPolicy.MFAMethods += $method.'@odata.type'.Replace('#microsoft.graph.', '')
                        
                        # Check for passwordless methods
                        if ($method.'@odata.type' -match "(fido2|windowsHello|microsoftAuthenticator)") {
                            $authMethodsPolicy.PasswordlessMethods += $method.'@odata.type'.Replace('#microsoft.graph.', '')
                        }
                    }
                }
                
                if ($authMethodsPolicy.MFAMethods.Count -gt 0) {
                    $authMethodsPolicy.MFAEnabled = $true
                    Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] MFA is enabled with $($authMethodsPolicy.MFAMethods.Count) methods" -ExportFileLocation $ExportDetails
                }
            }
            
        } catch {
            Write-Log -Type WARNING -Message "[Get-AuthenticationConfiguration] Unable to retrieve authentication methods policy: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }
        
        # Check Conditional Access policies for MFA requirements
        if ($script:tenantStatsHash["ConditionalAccessPolicies"]) {
            $mfaPolicies = $script:tenantStatsHash["ConditionalAccessPolicies"].Values | 
                Where-Object {$_.GrantControls_BuiltInControls -match "(?i)mfa"}
            
            if ($mfaPolicies) {
                $authMethodsPolicy.MFAEnabled = $true
                $authMethodsPolicy | Add-Member -MemberType NoteProperty -Name "MFAConditionalAccessPolicies" -Value $mfaPolicies.Count
                Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Found $($mfaPolicies.Count) CA policies requiring MFA" -ExportFileLocation $ExportDetails
            }
        }

        # Check authorization policy (app consent posture)
        try {
            $authorizationPolicyResponse = Office365Custom\Get-GraphData -Uri "https://graph.microsoft.com/v1.0/policies/authorizationPolicy" -Activity "Fetching authorization policy"
            $authorizationPolicy = @($authorizationPolicyResponse | Select-Object -First 1)
            if ($authorizationPolicy.Count -gt 0 -and $authorizationPolicy[0]) {
                $defaultPermissions = $null
                if ($authorizationPolicy[0].PSObject.Properties['defaultUserRolePermissions']) {
                    $defaultPermissions = $authorizationPolicy[0].defaultUserRolePermissions
                } elseif ($authorizationPolicy[0].PSObject.Properties['DefaultUserRolePermissions']) {
                    $defaultPermissions = $authorizationPolicy[0].DefaultUserRolePermissions
                }

                if ($defaultPermissions) {
                    if ($defaultPermissions.PSObject.Properties['allowedToCreateApps']) {
                        $authMethodsPolicy.DefaultUserCanCreateApps = [bool]$defaultPermissions.allowedToCreateApps
                    } elseif ($defaultPermissions.PSObject.Properties['AllowedToCreateApps']) {
                        $authMethodsPolicy.DefaultUserCanCreateApps = [bool]$defaultPermissions.AllowedToCreateApps
                    }

                    $assignedGrantPolicies = @()
                    if ($defaultPermissions.PSObject.Properties['permissionGrantPoliciesAssigned']) {
                        $assignedGrantPolicies = @($defaultPermissions.permissionGrantPoliciesAssigned)
                    } elseif ($defaultPermissions.PSObject.Properties['PermissionGrantPoliciesAssigned']) {
                        $assignedGrantPolicies = @($defaultPermissions.PermissionGrantPoliciesAssigned)
                    }
                    $authMethodsPolicy.PermissionGrantPoliciesAssigned = @($assignedGrantPolicies | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                }
            }
        } catch {
            Write-Log -Type WARNING -Message "[Get-AuthenticationConfiguration] Unable to retrieve authorization policy details: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        # Check admin consent request workflow
        try {
            $adminConsentPolicyResponse = Office365Custom\Get-GraphData -Uri "https://graph.microsoft.com/v1.0/policies/adminConsentRequestPolicy" -Activity "Fetching admin consent request policy"
            $adminConsentPolicy = @($adminConsentPolicyResponse | Select-Object -First 1)
            if ($adminConsentPolicy.Count -gt 0 -and $adminConsentPolicy[0]) {
                if ($adminConsentPolicy[0].PSObject.Properties['isEnabled']) {
                    $authMethodsPolicy.AdminConsentWorkflowEnabled = [bool]$adminConsentPolicy[0].isEnabled
                } elseif ($adminConsentPolicy[0].PSObject.Properties['IsEnabled']) {
                    $authMethodsPolicy.AdminConsentWorkflowEnabled = [bool]$adminConsentPolicy[0].IsEnabled
                }

                $reviewers = @()
                if ($adminConsentPolicy[0].PSObject.Properties['reviewers']) {
                    $reviewers = @($adminConsentPolicy[0].reviewers)
                } elseif ($adminConsentPolicy[0].PSObject.Properties['Reviewers']) {
                    $reviewers = @($adminConsentPolicy[0].Reviewers)
                }
                $authMethodsPolicy.AdminConsentWorkflowReviewerCount = $reviewers.Count
            }
        } catch {
            Write-Log -Type WARNING -Message "[Get-AuthenticationConfiguration] Unable to retrieve admin consent workflow policy: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        # Check Security Defaults policy state
        try {
            $securityDefaultsResponse = Office365Custom\Get-GraphData -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy" -Activity "Fetching security defaults policy"
            $securityDefaultsPolicy = @($securityDefaultsResponse | Select-Object -First 1)
            if ($securityDefaultsPolicy.Count -gt 0 -and $securityDefaultsPolicy[0]) {
                $isEnabled = $null
                if ($securityDefaultsPolicy[0].PSObject.Properties['isEnabled']) {
                    $isEnabled = [bool]$securityDefaultsPolicy[0].isEnabled
                } elseif ($securityDefaultsPolicy[0].PSObject.Properties['IsEnabled']) {
                    $isEnabled = [bool]$securityDefaultsPolicy[0].IsEnabled
                }

                $script:tenantStatsHash["SecurityDefaultsPolicy"]["Configuration"] = [PSCustomObject]@{
                    IsEnabled   = $isEnabled
                    Description = $(if ($isEnabled -eq $true) { 'Security Defaults are enabled.' } elseif ($isEnabled -eq $false) { 'Security Defaults are disabled.' } else { 'Security Defaults state unavailable.' })
                }
            }
        } catch {
            Write-Log -Type WARNING -Message "[Get-AuthenticationConfiguration] Unable to retrieve Security Defaults policy: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        if ($collectExtendedIdentityTierB) {
            try {
                Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Collecting enterprise application permission posture for Tier B findings" -ExportFileLocation $ExportDetails
                $resourceServicePrincipalCache = @{}
                $highPrivilegePatterns = @(
                    'Directory.ReadWrite.All', 'Directory.AccessAsUser.All', 'RoleManagement.ReadWrite.Directory',
                    'AppRoleAssignment.ReadWrite.All', 'Application.ReadWrite.All', 'Group.ReadWrite.All',
                    'User.ReadWrite.All', 'Mail.ReadWrite', 'Files.ReadWrite.All', 'Sites.FullControl.All',
                    'Sites.ReadWrite.All', 'Exchange.ManageAsApp', 'Policy.ReadWrite.ConditionalAccess',
                    'DeviceManagementManagedDevices.ReadWrite.All', 'DeviceManagementConfiguration.ReadWrite.All'
                )
                $servicePrincipals = @(Get-ArrayaGraphResource -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=id,displayName,appId,servicePrincipalType,accountEnabled,preferredSingleSignOnMode,tags,appRoleAssignmentRequired&`$top=250" -Activity 'Enterprise applications inventory' -Headers $global:GraphHeaders)
                $enterpriseAppIndex = 0
                $highPrivilegeAppCount = 0
                foreach ($servicePrincipal in $servicePrincipals) {
                    if ($null -eq $servicePrincipal -or [string]::IsNullOrWhiteSpace([string]$servicePrincipal.Id)) {
                        continue
                    }
                    $enterpriseAppIndex++
                    $spId = [string]$servicePrincipal.Id
                    $delegatedGrants = @()
                    $applicationPermissionValues = @()
                    try {
                        $delegatedGrants = @(Get-ArrayaGraphResource -Uri ("https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '{0}'&`$top=50" -f $spId) -Activity "Enterprise app delegated grants [$($servicePrincipal.DisplayName)]" -Headers $global:GraphHeaders)
                    } catch {
                        Write-Log -Type DEBUG -Message "[Get-AuthenticationConfiguration] Delegated grant lookup failed for '$($servicePrincipal.DisplayName)': $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                    }

                    try {
                        $appRoleAssignments = @(Get-ArrayaGraphResource -Uri ("https://graph.microsoft.com/v1.0/servicePrincipals/{0}/appRoleAssignments?`$top=50" -f $spId) -Activity "Enterprise app role assignments [$($servicePrincipal.DisplayName)]" -Headers $global:GraphHeaders)
                        foreach ($assignment in $appRoleAssignments) {
                            $resourceId = if ($assignment.PSObject.Properties['ResourceId']) { [string]$assignment.ResourceId } else { $null }
                            $appRoleId = if ($assignment.PSObject.Properties['AppRoleId']) { [string]$assignment.AppRoleId } else { $null }
                            if (-not $resourceId) { continue }

                            if (-not $resourceServicePrincipalCache.ContainsKey($resourceId)) {
                                try {
                                    $resourceServicePrincipalCache[$resourceId] = Get-ArrayaGraphResource -Uri ("https://graph.microsoft.com/v1.0/servicePrincipals/{0}?`$select=id,displayName,appRoles" -f $resourceId) -Activity "Enterprise app resource lookup [$resourceId]" -Headers $global:GraphHeaders
                                } catch {
                                    $resourceServicePrincipalCache[$resourceId] = $null
                                }
                            }

                            $resourcePrincipal = $resourceServicePrincipalCache[$resourceId]
                            $resourceDisplayName = if ($resourcePrincipal -and $resourcePrincipal.PSObject.Properties['DisplayName']) { [string]$resourcePrincipal.DisplayName } else { $resourceId }
                            $permissionValue = $null
                            if ($resourcePrincipal -and $resourcePrincipal.PSObject.Properties['AppRoles']) {
                                $permissionValue = @($resourcePrincipal.AppRoles | Where-Object { $_.Id -and ([string]$_.Id -eq $appRoleId) } | Select-Object -First 1 | ForEach-Object { $_.Value }) | Select-Object -First 1
                            }
                            if ([string]::IsNullOrWhiteSpace([string]$permissionValue)) {
                                $permissionValue = $appRoleId
                            }
                            $applicationPermissionValues += ("{0}:{1}" -f $resourceDisplayName, $permissionValue)
                        }
                    } catch {
                        Write-Log -Type DEBUG -Message "[Get-AuthenticationConfiguration] App-role assignment lookup failed for '$($servicePrincipal.DisplayName)': $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                    }

                    $delegatedScopes = @(
                        $delegatedGrants |
                            ForEach-Object {
                                if ($_.PSObject.Properties['Scope']) { [string]$_.Scope } elseif ($_.PSObject.Properties['scope']) { [string]$_.scope } else { $null }
                            } |
                            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    )
                    $flattenedDelegatedScopes = @(
                        $delegatedScopes |
                            ForEach-Object { $_ -split ' ' } |
                            ForEach-Object { $_.Trim() } |
                            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    )
                    $combinedPermissionText = (@($applicationPermissionValues) + @($flattenedDelegatedScopes)) -join ';'
                    $highPrivilegeMatches = @(
                        $highPrivilegePatterns | Where-Object { $combinedPermissionText -match [regex]::Escape($_) }
                    )
                    if ($highPrivilegeMatches.Count -gt 0) {
                        $highPrivilegeAppCount++
                    }

                    $script:tenantStatsHash["EnterpriseApplications"][("{0:D3}-{1}" -f $enterpriseAppIndex, ($servicePrincipal.DisplayName -replace '[^a-zA-Z0-9._-]', '_'))] = [PSCustomObject]@{
                        DisplayName                     = $servicePrincipal.DisplayName
                        AppId                           = $servicePrincipal.AppId
                        ServicePrincipalId              = $spId
                        ServicePrincipalType            = $servicePrincipal.ServicePrincipalType
                        AccountEnabled                  = $servicePrincipal.AccountEnabled
                        PreferredSingleSignOnMode       = $servicePrincipal.PreferredSingleSignOnMode
                        AppRoleAssignmentRequired       = $servicePrincipal.AppRoleAssignmentRequired
                        Tags                            = @($servicePrincipal.Tags) -join ','
                        DelegatedPermissionScopes       = $flattenedDelegatedScopes -join ','
                        DelegatedPermissionGrantCount   = @($delegatedGrants).Count
                        ApplicationPermissions          = $applicationPermissionValues -join ','
                        ApplicationPermissionCount      = @($applicationPermissionValues).Count
                        HighPrivilegePermissionCount    = $highPrivilegeMatches.Count
                        HighPrivilegePermissions        = $highPrivilegeMatches -join ','
                    }
                }

                $script:tenantStatsHash["EnterpriseApplicationSummary"]["Summary"] = [PSCustomObject]@{
                    TotalEnterpriseApplications       = $script:tenantStatsHash["EnterpriseApplications"].Count
                    ApplicationsWithHighPrivilege     = $highPrivilegeAppCount
                    ApplicationsWithDelegatedGrants   = @($script:tenantStatsHash["EnterpriseApplications"].Values | Where-Object { $_.DelegatedPermissionGrantCount -gt 0 }).Count
                    ApplicationsWithApplicationPerms  = @($script:tenantStatsHash["EnterpriseApplications"].Values | Where-Object { $_.ApplicationPermissionCount -gt 0 }).Count
                }
            } catch {
                Write-Log -Type WARNING -Message "[Get-AuthenticationConfiguration] Unable to collect enterprise application permission posture: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }
        }
        
        $script:tenantStatsHash["AuthenticationConfig"]["Configuration"] = $authMethodsPolicy

        $ssoAppNames = @($authMethodsPolicy.SSOApplications | ForEach-Object { $_.DisplayName } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $script:tenantStatsHash["AuthenticationConfigSummary"]["Summary"] = [PSCustomObject]@{
            MFAEnabled                   = $authMethodsPolicy.MFAEnabled
            MFAMethods                   = ($authMethodsPolicy.MFAMethods -join ', ')
            SSOEnabled                   = $authMethodsPolicy.SSOEnabled
            SSOApplicationsCount         = @($authMethodsPolicy.SSOApplications).Count
            SSOApplications              = ($ssoAppNames -join ', ')
            FederatedDomains             = ($authMethodsPolicy.FederatedDomains -join ', ')
            PasswordlessMethods          = ($authMethodsPolicy.PasswordlessMethods -join ', ')
            MFAConditionalAccessPolicies = $(if ($authMethodsPolicy.PSObject.Properties['MFAConditionalAccessPolicies']) { $authMethodsPolicy.MFAConditionalAccessPolicies } else { 0 })
            DefaultUserCanCreateApps     = $(if ($null -eq $authMethodsPolicy.DefaultUserCanCreateApps) { 'Not available' } elseif ($authMethodsPolicy.DefaultUserCanCreateApps) { 'Yes' } else { 'No' })
            PermissionGrantPolicies      = $(if (@($authMethodsPolicy.PermissionGrantPoliciesAssigned).Count -gt 0) { $authMethodsPolicy.PermissionGrantPoliciesAssigned -join ', ' } else { 'Not available' })
            AdminConsentWorkflowEnabled  = $(if ($null -eq $authMethodsPolicy.AdminConsentWorkflowEnabled) { 'Not available' } elseif ($authMethodsPolicy.AdminConsentWorkflowEnabled) { 'Enabled' } else { 'Disabled' })
            AdminConsentWorkflowReviewers = $authMethodsPolicy.AdminConsentWorkflowReviewerCount
        }

        $methodIndex = 0
        foreach ($method in @($authMethodsPolicy.MFAMethods)) {
            $methodIndex++
            $script:tenantStatsHash["AuthenticationMethods"][("{0:D3}-MFA" -f $methodIndex)] = [PSCustomObject]@{
                Category = 'MFA Method'
                Value    = $method
            }
        }

        foreach ($method in @($authMethodsPolicy.PasswordlessMethods)) {
            $methodIndex++
            $script:tenantStatsHash["AuthenticationMethods"][("{0:D3}-Passwordless" -f $methodIndex)] = [PSCustomObject]@{
                Category = 'Passwordless Method'
                Value    = $method
            }
        }

        foreach ($domainName in @($authMethodsPolicy.FederatedDomains)) {
            $methodIndex++
            $script:tenantStatsHash["AuthenticationMethods"][("{0:D3}-Federated" -f $methodIndex)] = [PSCustomObject]@{
                Category = 'Federated Domain'
                Value    = $domainName
            }
        }

        $ssoIndex = 0
        foreach ($app in @($authMethodsPolicy.SSOApplications)) {
            $ssoIndex++
            $script:tenantStatsHash["AuthenticationSSOApplications"][("{0:D3}-{1}" -f $ssoIndex, $app.DisplayName)] = $app
        }

        $guestUsers = @(
            $script:tenantStatsHash["Users"].Values |
                Where-Object { ([string]$_.UserType).ToLowerInvariant() -eq 'guest' -or ([string]$_.UserPrincipalName -like '*#EXT#*') }
        )
        $inactiveGuests = @(
            $guestUsers |
                Where-Object {
                    $lastSignIn = $null
                    try { $lastSignIn = [datetime]$_.LastSignInDateTime } catch { $lastSignIn = $null }
                    ($null -eq $lastSignIn) -or ($lastSignIn -lt (Get-Date).AddDays(-90))
                }
        )
        $script:tenantStatsHash["GuestSignInSummary"]["Summary"] = [PSCustomObject]@{
            TotalGuests            = $guestUsers.Count
            InactiveGuests90Days   = $inactiveGuests.Count
            GuestsNeverSignedIn    = @($guestUsers | Where-Object { -not $_.LastSignInDateTime }).Count
            GuestAccountsEnabled   = @($guestUsers | Where-Object { $_.AccountEnabled -eq $true }).Count
        }

        $privilegedAdmins = @($script:tenantStatsHash["Admins"].Values)
        $staleAdmins = @(
            $privilegedAdmins |
                Where-Object {
                    $lastSignIn = $null
                    try { $lastSignIn = [datetime]$_.LastSignInDateTime } catch { $lastSignIn = $null }
                    ($_.AccountEnabled -eq $true) -and $lastSignIn -and ($lastSignIn -lt (Get-Date).AddDays(-90))
                }
        )
        $guestAdmins = @(
            $privilegedAdmins |
                Where-Object {
                    ([string]$_.UserPrincipalName -like '*#EXT#*') -or ([string]$_.UserType).ToLowerInvariant() -eq 'guest'
                }
        )
        $script:tenantStatsHash["PrivilegedAccessSummary"]["Summary"] = [PSCustomObject]@{
            TotalPrivilegedIdentities = $privilegedAdmins.Count
            StalePrivilegedAccounts90Days = $staleAdmins.Count
            GuestPrivilegedAccounts   = $guestAdmins.Count
            GlobalAdministratorCount  = @($privilegedAdmins | Where-Object { ([string]$_.Role) -match 'Global Administrator' }).Count
        }
        
    } catch {
        Write-Log -Type ERROR -Message "An error occurred checking authentication configuration. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] COMPLETED: Checking Authentication Configuration in $($CompletedTime)" -ExportFileLocation $ExportDetails
}

function Get-TenantOverviewInfo {
    [CmdletBinding()]
    param ()

    $start = Get-Date
    if (-not $script:tenantStatsHash) { $script:tenantStatsHash = @{} }
    $script:tenantStatsHash["TenantInfo"] = @{}

    Write-Host "Gathering Tenant Overview Info ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-TenantOverviewInfo] START: Gathering tenant overview information" -ExportFileLocation $ExportDetails

    try {
        $org = Get-AssessmentTenantOrganization

        $initialDomain = $null
        $defaultDomain = $null
        $verifiedDomains = @()
        if ($org -and $org.PSObject.Properties['VerifiedDomains']) {
            $verifiedDomains = $org.VerifiedDomains
            $initial = $verifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1
            $default = $verifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -First 1
            if ($initial) { $initialDomain = if ($initial.PSObject.Properties['Name']) { $initial.Name } else { $initial.Id } }
            if ($default) { $defaultDomain = if ($default.PSObject.Properties['Name']) { $default.Name } else { $default.Id } }
        }

        $multiGeoEnabled = $null
        $multiGeoAllowed = @()
        $multiGeoCentral = $null
        if ($org.PSObject.Properties['IsMultipleGeolocationEnabled']) {
            $multiGeoEnabled = [bool]$org.IsMultipleGeolocationEnabled
        }
        if ($org.PSObject.Properties['MultiGeoConfiguration'] -and $org.MultiGeoConfiguration) {
            $cfg = $org.MultiGeoConfiguration
            if ($cfg.PSObject.Properties['AllowedDataLocations'] -and $cfg.AllowedDataLocations) {
                $multiGeoAllowed = @($cfg.AllowedDataLocations)
                if ($multiGeoAllowed.Count -gt 1) { $multiGeoEnabled = $true }
            }
            if ($cfg.PSObject.Properties['PreferredDataLocation'] -and $cfg.PreferredDataLocation) {
                $multiGeoCentral = $cfg.PreferredDataLocation
            }
        }

        $selfServiceAnswer = "Not collected"
        $selfServiceNotes = "MSCommerce module not installed"
        $selfServiceDetails = $null

        try {
            if (-not (Get-Module -ListAvailable -Name MSCommerce)) {
                Write-Log -Type INFO -Message "[Get-TenantOverviewInfo] Installing MSCommerce module (CurrentUser)" -ExportFileLocation $ExportDetails
                Install-Module -Name MSCommerce -Scope CurrentUser -Force -ErrorAction Stop
            }

            if ($PSVersionTable.PSVersion.Major -ge 6) {
                Import-Module MSCommerce -UseWindowsPowerShell -ErrorAction Stop
            } else {
                Import-Module MSCommerce -ErrorAction Stop
            }

            try {
                Connect-MSCommerce -ErrorAction Stop | Out-Null
                $sspPolicies = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase -ErrorAction Stop
                if ($sspPolicies -and $sspPolicies.Count -gt 0) {
                    $enabled = ($sspPolicies | Where-Object { $_.Value -eq 'Enabled' }).Count
                    $disabled = ($sspPolicies | Where-Object { $_.Value -eq 'Disabled' }).Count
                    $trialOnly = ($sspPolicies | Where-Object { $_.Value -eq 'OnlyTrialsWithoutPaymentMethod' }).Count
                    $total = $sspPolicies.Count
                    if ($enabled -gt 0) {
                        $selfServiceAnswer = "Yes (Enabled for $enabled/$total products)"
                    } else {
                        $selfServiceAnswer = "No (Disabled for all $total products)"
                    }
                    $selfServiceNotes = "Disabled: $disabled; Trial-only: $trialOnly"
                    $selfServiceDetails = [PSCustomObject]@{
                        EnabledCount = $enabled
                        DisabledCount = $disabled
                        TrialOnlyCount = $trialOnly
                        TotalCount = $total
                    }
                } else {
                    $selfServiceAnswer = "Unknown"
                    $selfServiceNotes = "No policy data returned"
                }
            } catch {
                $selfServiceAnswer = "Unknown"
                $selfServiceNotes = "MSCommerce connection or policy retrieval failed: $($_.Exception.Message)"
            }
        } catch {
            $selfServiceAnswer = "Unknown"
            $selfServiceNotes = "MSCommerce module install/load failed: $($_.Exception.Message)"
        }

        $azureAnswer = "Not collected"
        $azureNotes = "Azure module not used in this report"
        $azureDetails = $null

        $script:tenantStatsHash["TenantInfo"] = [PSCustomObject]@{
            DisplayName           = $org.DisplayName
            TenantId              = $org.Id
            InitialDomain         = $initialDomain
            DefaultDomain         = $defaultDomain
            PreferredDataLocation = $org.PreferredDataLocation
            Country               = $org.Country
            CountryLetterCode     = $org.CountryLetterCode
            MultiGeoEnabled       = $multiGeoEnabled
            MultiGeoAllowed       = $multiGeoAllowed
            MultiGeoCentral       = $multiGeoCentral
            SelfServicePurchase   = [PSCustomObject]@{
                Answer = $selfServiceAnswer
                Notes = $selfServiceNotes
                Details = $selfServiceDetails
            }
            AzureResourceUsage    = [PSCustomObject]@{
                Answer = $azureAnswer
                Notes = $azureNotes
                Details = $azureDetails
            }
        }

        Write-Log -Type INFO -Message "[Get-TenantOverviewInfo] Retrieved tenant info for $($org.DisplayName)" -ExportFileLocation $ExportDetails
    } catch {
        Write-Log -Type WARNING -Message "[Get-TenantOverviewInfo] Unable to retrieve tenant overview info: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }

    $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
    Write-Host "Completed in $($elapsed)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-TenantOverviewInfo] COMPLETED in $elapsed" -ExportFileLocation $ExportDetails
}

function Get-AdConnectSyncDetails {
    [CmdletBinding()]
    param ()
    
    $start = Get-Date
    if (-not $script:tenantStatsHash) { $script:tenantStatsHash = @{} }
    $script:tenantStatsHash["AdConnectConfiguration"] = @{}
    
    Write-Host "Checking AD Connect/Sync status ..." -ForegroundColor Cyan -NoNewline
    Write-Log -Type INFO -Message "[Get-AdConnectSyncDetails] START" -ExportFileLocation $ExportDetails
    
    try {
        $org = $null
        try { $org = Get-AssessmentTenantOrganization } catch {}
        
        $summary = [pscustomobject]@{
            OnPremisesSyncEnabled = if ($org) { $org.OnPremisesSyncEnabled } else { $null }
            OnPremisesLastSyncDateTime = if ($org) { $org.OnPremisesLastSyncDateTime } else { $null }
        }
        
        $serviceDetails = @()
        $syncErrors = @()
        
        if (Get-Command Get-AzureADConnectHealthSyncServices -ErrorAction SilentlyContinue) {
            try {
                $services = Get-AzureADConnectHealthSyncServices -ErrorAction SilentlyContinue
                foreach ($svc in $services) {
                    $serviceName = if ($svc.PSObject.Properties['ServiceName']) { $svc.ServiceName } elseif ($svc.PSObject.Properties['Name']) { $svc.Name } else { $null }
                    $serverName = if ($svc.PSObject.Properties['ServerName']) { $svc.ServerName } elseif ($svc.PSObject.Properties['Server']) { $svc.Server } else { $null }
                    $lastSync = if ($svc.PSObject.Properties['LastSyncTime']) { $svc.LastSyncTime } elseif ($svc.PSObject.Properties['LastSyncDateTime']) { $svc.LastSyncDateTime } else { $null }
                    $serviceDetails += [pscustomobject]@{
                        ServiceName = $serviceName
                        ServerName = $serverName
                        LastSyncTime = $lastSync
                    }
                }
            } catch {}
        }
        
        if (Get-Command Get-AzureADConnectHealthSyncErrors -ErrorAction SilentlyContinue) {
            try {
                $syncErrors = Get-AzureADConnectHealthSyncErrors -ErrorAction SilentlyContinue
            } catch {}
        } elseif (Get-Command Get-AzureADConnectHealthSyncAlert -ErrorAction SilentlyContinue) {
            try {
                $syncErrors = Get-AzureADConnectHealthSyncAlert -ErrorAction SilentlyContinue
            } catch {}
        }
        
        $script:tenantStatsHash["AdConnectConfiguration"]["Summary"] = $summary
        $script:tenantStatsHash["AdConnectConfiguration"]["SyncServices"] = $serviceDetails
        $script:tenantStatsHash["AdConnectConfiguration"]["RecentErrors"] = $syncErrors | Select-Object -First 10
        $script:tenantStatsHash["AdConnectConfiguration"]["ErrorCount"] = ($syncErrors | Measure-Object).Count
        
        Write-Log -Type INFO -Message "[Get-AdConnectSyncDetails] DirSyncEnabled=$($summary.OnPremisesSyncEnabled) LastSync=$($summary.OnPremisesLastSyncDateTime) Services=$($serviceDetails.Count) Errors=$(($syncErrors | Measure-Object).Count)" -ExportFileLocation $ExportDetails
    } catch {
        Write-Log -Type WARNING -Message "[Get-AdConnectSyncDetails] Error: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }
    
    $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
    Write-Host " Completed in $elapsed" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AdConnectSyncDetails] COMPLETED in $elapsed" -ExportFileLocation $ExportDetails
}

function Get-MfaRegistrationDetails {
    [CmdletBinding()]
    param ()
    
    $start = Get-Date
    if (-not $script:tenantStatsHash) { $script:tenantStatsHash = @{} }
    $script:tenantStatsHash["MfaRegistrationDetails"] = @{}
    $script:tenantStatsHash["MfaRegistrationSummary"] = $null
    
    Write-Host "Gathering MFA registration details ..." -ForegroundColor Cyan -NoNewline
    Write-Log -Type INFO -Message "[Get-MfaRegistrationDetails] START: Gathering MFA registration details" -ExportFileLocation $ExportDetails
    
    try {
        $regData = @()
        try {
            if (Get-Command Get-MgReportAuthenticationMethodUserRegistrationDetail -ErrorAction SilentlyContinue) {
                $regData = Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ProgressAction SilentlyContinue -ErrorAction Stop
            }
        } catch {}
        
        if (-not $regData -or $regData.Count -eq 0) {
            try {
                $uri = "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails"
                $regData = Office365Custom\Get-GraphData -Uri $uri -Activity "Fetching MFA Registration Details"
            } catch {
                Write-Log -Type WARNING -Message "[Get-MfaRegistrationDetails] Unable to retrieve registration report: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }
        }
        
        $methodCounts = @{}
        $registeredUsers = 0
        $totalUsers = 0
        
        foreach ($user in $regData) {
            $totalUsers++
            $upn = $user.UserPrincipalName
            $methods = @()
            if ($user.MethodsRegistered) {
                $methods = @($user.MethodsRegistered)
            }
            foreach ($m in $methods) {
                if (-not $methodCounts.ContainsKey($m)) { $methodCounts[$m] = 0 }
                $methodCounts[$m]++
            }
            if ($user.IsMfaRegistered -eq $true) { $registeredUsers++ }
            
            if ($upn) {
                $script:tenantStatsHash["MfaRegistrationDetails"][$upn] = [pscustomobject]@{
                    UserPrincipalName = $upn
                    IsMfaRegistered = $user.IsMfaRegistered
                    IsMfaCapable = $user.IsMfaCapable
                    MethodsRegistered = if ($methods) { $methods -join ', ' } else { $null }
                    DefaultMfaMethod = $user.DefaultMfaMethod
                }
            }
        }
        
        if ($totalUsers -gt 0) {
            $summary = [pscustomobject]@{
                TotalUsers = $totalUsers
                RegisteredUsers = $registeredUsers
                NotRegisteredUsers = ($totalUsers - $registeredUsers)
                RegistrationPercent = [math]::Round(($registeredUsers / $totalUsers) * 100, 1)
                MethodCounts = $methodCounts
            }
            $script:tenantStatsHash["MfaRegistrationSummary"] = $summary
        }
    } catch {
        Write-Log -Type WARNING -Message "[Get-MfaRegistrationDetails] Error: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }
    
    $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
    Write-Host " Completed in $elapsed" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-MfaRegistrationDetails] COMPLETED in $elapsed" -ExportFileLocation $ExportDetails
}

# Federation and Cross-Tenant Configuration (External Identities)
function Get-FederationAndCrossTenantConfiguration {
    [CmdletBinding()]
    param()

    $start = Get-Date
    if (-not $script:tenantStatsHash) { $script:tenantStatsHash = @{} }
    $script:tenantStatsHash["FederationConfiguration"] = @{}

    Write-Host "Checking Federation and Cross-Tenant Configuration ..." -ForegroundColor Cyan -NoNewline
    Write-Log -Type INFO -Message "[Get-FederationAndCrossTenantConfiguration] START" -ExportFileLocation $ExportDetails

    try {
        # Exchange federation signals
        $orgR = Get-OrganizationRelationship -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.TargetApplicationUri -or $_.TargetAutodiscoverEpr -or $_.DomainNames
                }
        $ioc  = Get-IntraOrganizationConnector -ErrorAction SilentlyContinue | Where-Object { $_.Enabled }

        $exchangeFed = [pscustomobject]@{
            OrganizationRelationshipCount = $orgR.Count
            IntraOrgConnectorCount         = $ioc.Count
            OrganizationRelationships      = ($orgR | Select-Object -ExpandProperty Name) -join ', '
            IntraOrgConnectors             = ($ioc  | Select-Object -ExpandProperty Name) -join ', '
        }

        # External identities / cross-tenant access (Graph)
        $crossTenantPolicy = $null
        $crossTenantPartners = @()
        $b2bPolicy = $null

        try {
            $crossTenantPolicy = Get-MgPolicyCrossTenantAccessPolicy -ProgressAction SilentlyContinue -ErrorAction Stop
        } catch {
            Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] Unable to retrieve CrossTenantAccessPolicy: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        try {
            $crossTenantPartners = Get-MgPolicyCrossTenantAccessPolicyPartner -All -ProgressAction SilentlyContinue -ErrorAction Stop
        } catch {
            Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] Unable to retrieve CrossTenantAccessPolicy partners: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        try {
            $b2bPolicy = Get-MgPolicyB2BManagementPolicy -ProgressAction SilentlyContinue -ErrorAction Stop
        } catch {
            if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
                try {
                    $b2bPolicy = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/b2bManagementPolicy' -ProgressAction SilentlyContinue -ErrorAction Stop
                } catch {
                    $b2bMessage = $_.Exception.Message
                    $b2bLogType = if ($b2bMessage -match 'BadRequest') { 'INFO' } else { 'WARNING' }
                    Write-Log -Type $b2bLogType -Message "[Get-FederationAndCrossTenantConfiguration] Unable to retrieve B2B management policy: $b2bMessage" -ExportFileLocation $ExportDetails
                }
            } else {
                $b2bMessage = $_.Exception.Message
                $b2bLogType = if ($b2bMessage -match 'BadRequest') { 'INFO' } else { 'WARNING' }
                Write-Log -Type $b2bLogType -Message "[Get-FederationAndCrossTenantConfiguration] Unable to retrieve B2B management policy: $b2bMessage" -ExportFileLocation $ExportDetails
            }
        }

        $defaultInboundMfa = if ($crossTenantPolicy -and $crossTenantPolicy.Default -and $crossTenantPolicy.Default.InboundTrust) {
            $crossTenantPolicy.Default.InboundTrust.IsMfaAccepted
        } else { $null }
        $defaultOutboundMfa = if ($crossTenantPolicy -and $crossTenantPolicy.Default -and $crossTenantPolicy.Default.OutboundTrust) {
            $crossTenantPolicy.Default.OutboundTrust.IsMfaAccepted
        } else { $null }

        function Get-B2BDirectConnectSummary {
            param([object]$Setting)
            if (-not $Setting) { return "Not configured" }
            $usersAndGroups = $null
            $applications = $null
            if ($Setting.PSObject.Properties['UsersAndGroups']) { $usersAndGroups = $Setting.UsersAndGroups }
            if ($Setting.PSObject.Properties['Applications']) { $applications = $Setting.Applications }

            $uAccess = if ($usersAndGroups -and $usersAndGroups.PSObject.Properties['AccessType']) { $usersAndGroups.AccessType } else { $null }
            $uTargets = if ($usersAndGroups -and $usersAndGroups.PSObject.Properties['Targets']) { $usersAndGroups.Targets } else { $null }
            $uTargetText = if ($uTargets) {
                if ($uTargets -is [string]) { $uTargets } else { "$(($uTargets | Measure-Object).Count) targets" }
            } else { $null }

            $aAccess = if ($applications -and $applications.PSObject.Properties['AccessType']) { $applications.AccessType } else { $null }
            $aTargets = if ($applications -and $applications.PSObject.Properties['Targets']) { $applications.Targets } else { $null }
            $aTargetText = if ($aTargets) {
                if ($aTargets -is [string]) { $aTargets } else { "$(($aTargets | Measure-Object).Count) targets" }
            } else { $null }

            $parts = @()
            if ($uAccess -or $uTargetText) { $parts += "Users: $uAccess $uTargetText".Trim() }
            if ($aAccess -or $aTargetText) { $parts += "Apps: $aAccess $aTargetText".Trim() }
            if ($parts.Count -gt 0) { return ($parts -join '; ') }
            return "Configured (details unavailable)"
        }

        function Resolve-TenantDisplayName {
            param(
                [Parameter(Mandatory = $true)][string]$TenantId,
                [Parameter(Mandatory = $false)][object]$PartnerObject
            )
            
            # 1) Partner object might already include a display name
            if ($PartnerObject -and $PartnerObject.PSObject.Properties['DisplayName']) {
                $name = $PartnerObject.DisplayName
                if (-not [string]::IsNullOrWhiteSpace($name)) { return $name }
            }
            
            # 2) Try Graph v1.0 findTenantInformationByTenantId (preferred)
            $tenantLookupUri = "https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByTenantId(tenantId='$TenantId')"
            try {
                if (Get-Command Find-MgTenantRelationshipTenantInformationByTenantId -ErrorAction SilentlyContinue) {
                    $lookup = Find-MgTenantRelationshipTenantInformationByTenantId -TenantId $TenantId -ProgressAction SilentlyContinue -ErrorAction Stop
                    if ($lookup.DisplayName) { return $lookup.DisplayName }
                }
                if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
                    $lookup = Office365Custom\Get-GraphData -Uri $tenantLookupUri -Activity "Tenant lookup (v1.0)"
                    $lookupItem = $lookup | Select-Object -First 1
                    if ($lookupItem.displayName) { return $lookupItem.displayName }
                } elseif ($global:GraphHeaders) {
                    $lookup = Invoke-QuietRestMethod -Parameters @{
                        Uri         = $tenantLookupUri
                        Headers     = $global:GraphHeaders
                        Method      = 'GET'
                        ContentType = 'application/json'
                        ErrorAction = 'Stop'
                    }
                    if ($lookup.displayName) { return $lookup.displayName }
                }
            } catch {
                Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] Tenant name lookup (v1.0) failed for $($TenantId): $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }

            # 3) Try TenantInfoApp lookup (public)
            try {
                $tenantInfoAppUris = @(
                    "https://tenantinfoapp.azurewebsites.us/tenantlookup?tenantId=$TenantId",
                    "https://tenantinfoapp.azurewebsites.us/?tenantId=$TenantId"
                )
                foreach ($uri in $tenantInfoAppUris) {
                    try {
                        $response = Invoke-QuietWebRequest -Parameters @{
                            Uri           = $uri
                            UseBasicParsing = $true
                            ErrorAction   = 'Stop'
                        }
                        if ($response -and $response.Content) {
                            # Try JSON first
                            $json = $null
                            try { $json = $response.Content | ConvertFrom-Json -ErrorAction Stop } catch {}
                            if ($json) {
                                foreach ($prop in @('displayName','tenantName','domainName','primaryDomain')) {
                                    if ($json.PSObject.Properties[$prop] -and $json.$prop) { return $json.$prop }
                                }
                            }

                            # Try HTML label parsing (best-effort)
                            if ($response.Content -match "Domain Name:\s*</[^>]+>\s*([^<]+)") { return $matches[1].Trim() }
                            if ($response.Content -match "Tenant Name:\s*</[^>]+>\s*([^<]+)") { return $matches[1].Trim() }
                        }
                    } catch {}
                }
            } catch {
                Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] TenantInfoApp lookup failed for $($TenantId): $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }
            
            # 4) Try Graph beta findTenantInformationByTenantId
            $tenantLookupUri = "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId='$TenantId')"
            try {
                if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
                    $lookup = Office365Custom\Get-GraphData -Uri $tenantLookupUri -Activity "Tenant lookup (beta)"
                    $lookupItem = $lookup | Select-Object -First 1
                    if ($lookupItem.displayName) { return $lookupItem.displayName }
                } elseif ($global:GraphHeaders) {
                    $lookup = Invoke-QuietRestMethod -Parameters @{
                        Uri         = $tenantLookupUri
                        Headers     = $global:GraphHeaders
                        Method      = 'GET'
                        ContentType = 'application/json'
                        ErrorAction = 'Stop'
                    }
                    if ($lookup.displayName) { return $lookup.displayName }
                }
            } catch {
                Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] Tenant name lookup (beta) failed for $($TenantId): $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }
            
            # 5) Last resort - this often fails cross-tenant
            try {
                $orgInfo = Get-AssessmentTenantOrganization -OrganizationId $TenantId
                if ($orgInfo) { return $orgInfo.DisplayName }
            } catch {
                Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] Unable to resolve tenant name for $($TenantId): $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }
            
            return $null
        }

        $partnerDetails = @()
        foreach ($partner in $crossTenantPartners) {
            $tenantId = $partner.TenantId
            $displayName = Resolve-TenantDisplayName -TenantId $tenantId -PartnerObject $partner
            $b2bInboundSummary = Get-B2BDirectConnectSummary -Setting $partner.B2bDirectConnectInbound
            $b2bOutboundSummary = Get-B2BDirectConnectSummary -Setting $partner.B2bDirectConnectOutbound
            $inboundTrust = $partner.InboundTrust
            $partnerDetails += [pscustomobject]@{
                TenantId = $tenantId
                DisplayName = $displayName
                B2BDirectConnectInbound = $b2bInboundSummary
                B2BDirectConnectOutbound = $b2bOutboundSummary
                TrustMfa = if ($inboundTrust) { $inboundTrust.IsMfaAccepted } else { $null }
                TrustCompliantDevices = if ($inboundTrust) { $inboundTrust.IsCompliantDeviceAccepted } else { $null }
                TrustHybridJoinedDevices = if ($inboundTrust) { $inboundTrust.IsHybridAzureAdJoinedDeviceAccepted } else { $null }
            }
        }

        $defaultB2BInbound = $null
        $defaultB2BOutbound = $null
        if ($crossTenantPolicy -and $crossTenantPolicy.Default) {
            $defaultB2BInbound = Get-B2BDirectConnectSummary -Setting $crossTenantPolicy.Default.B2bDirectConnectInbound
            $defaultB2BOutbound = Get-B2BDirectConnectSummary -Setting $crossTenantPolicy.Default.B2bDirectConnectOutbound
        }

        $crossTenantSummary = [pscustomobject]@{
            HasCrossTenantAccessPolicy = [bool]($crossTenantPolicy)
            PartnerCount               = ($crossTenantPartners | Measure-Object).Count
            PartnerTenants             = ($crossTenantPartners | ForEach-Object { $_.TenantId }) -join ', '
            PartnerTenantNames         = ($partnerDetails | Where-Object { $_.DisplayName } | ForEach-Object { $_.DisplayName }) -join ', '
            PartnerTenantDetails       = $partnerDetails
            DefaultInboundAccess       = $defaultInboundMfa
            DefaultOutboundAccess      = $defaultOutboundMfa
            DefaultB2BDirectConnectInbound  = $defaultB2BInbound
            DefaultB2BDirectConnectOutbound = $defaultB2BOutbound
        }

        $externalIdentities = [pscustomobject]@{
            B2BManagementPolicyPresent = [bool]($b2bPolicy)
            GuestUserRole              = if ($b2bPolicy) { $b2bPolicy.GuestUserRoleId } else { $null }
            InvitationsAllowed         = if ($b2bPolicy -and $b2bPolicy.InvitationPolicy) { $b2bPolicy.InvitationPolicy.AllowedToInvite } else { $null }
        }

        $script:tenantStatsHash["FederationConfiguration"]["ExchangeFederation"] = $exchangeFed
        $script:tenantStatsHash["FederationConfiguration"]["CrossTenantAccess"] = $crossTenantSummary
        $script:tenantStatsHash["FederationConfiguration"]["ExternalIdentities"] = $externalIdentities

        Write-Log -Type INFO -Message "[Get-FederationAndCrossTenantConfiguration] OrgRel=$($exchangeFed.OrganizationRelationshipCount) IOC=$($exchangeFed.IntraOrgConnectorCount) Partners=$($crossTenantSummary.PartnerCount)" -ExportFileLocation $ExportDetails
    }
    catch {
        Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] Error: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }

    $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
    Write-Host " Completed in $elapsed" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-FederationAndCrossTenantConfiguration] COMPLETED in $elapsed" -ExportFileLocation $ExportDetails
}

function Export-TenantStatsJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory)]
        [string]$Path
    )

    function Get-TenantStatsExportCount {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Key
        )

        if (-not ($TenantStatsHash -is [System.Collections.IDictionary]) -or -not $TenantStatsHash.ContainsKey($Key)) {
            return 0
        }

        $value = $TenantStatsHash[$Key]
        if ($value -is [System.Collections.IDictionary]) {
            return $value.Count
        }
        if ($null -ne $value -and $value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
            return @($value).Count
        }

        return 0
    }

    $metadata = [ordered]@{
        GeneratedAt        = (Get-Date).ToString('o')
        OutputProfile      = $OutputProfile
        OutputProfileLabel = $script:EffectiveOutputProfileLabel
        ReportingMode      = $reportingMode
        AuthMode           = $script:CurrentGraphMode
        CollectionOnly     = [bool]$runCollectionOnly
        ExportOnly         = [bool]$runExportOnly
    }

    if ($TenantStatsHash.ContainsKey('TenantInfo') -and $TenantStatsHash['TenantInfo']) {
        $tenantInfo = $TenantStatsHash['TenantInfo']
        $metadata['Tenant'] = [ordered]@{
            TenantId          = $tenantInfo.TenantID
            DisplayName       = $tenantInfo.DisplayName
            DefaultDomainName = $tenantInfo.DefaultDomainName
        }
    }

    $collectionPlan = [ordered]@{}
    if ($script:SnapshotCollectionPlan -is [System.Collections.IDictionary]) {
        foreach ($entry in $script:SnapshotCollectionPlan.GetEnumerator()) {
            $collectionPlan[[string]$entry.Key] = $entry.Value
        }
    }
    elseif ($script:ProfileCollectionPlan -is [System.Collections.IDictionary]) {
        foreach ($entry in $script:ProfileCollectionPlan.GetEnumerator()) {
            $collectionPlan[[string]$entry.Key] = [ordered]@{
                Status             = if ([bool]$entry.Value) { 'Collected' } else { 'Skipped' }
                NotCollectedReason = if ([bool]$entry.Value) { $null } else { 'Disabled by output profile policy.' }
            }
        }
    }

    $warningSummary = @()
    $errorSummary = @()
    if ($global:AllDiscoveryErrors) {
        foreach ($errorRow in $global:AllDiscoveryErrors) {
            $msg = [string]$errorRow.ErrorMessage
            if ([string]::IsNullOrWhiteSpace($msg)) {
                continue
            }
            $errorSummary += $msg
        }
    }

    $sourceCoverage = [ordered]@{}
    if ($TenantStatsHash.ContainsKey('PrimaryMailboxStatsCollectionSummary') -and $TenantStatsHash['PrimaryMailboxStatsCollectionSummary']) {
        $sourceCoverage['PrimaryMailboxStats'] = $TenantStatsHash['PrimaryMailboxStatsCollectionSummary']
    }
    if ($TenantStatsHash.ContainsKey('UnifiedGroupMailboxStatsCollectionSummary') -and $TenantStatsHash['UnifiedGroupMailboxStatsCollectionSummary']) {
        $sourceCoverage['UnifiedGroupMailboxStats'] = $TenantStatsHash['UnifiedGroupMailboxStatsCollectionSummary']
    }

    $diagnostics = [ordered]@{
        WarningCount   = $warningSummary.Count
        ErrorCount     = $errorSummary.Count
        WarningSummary = $warningSummary
        ErrorSummary   = $errorSummary
        SourceCoverage = $sourceCoverage
        CollectorStats = [ordered]@{
            StepMetrics = $script:AssessmentStepMetrics
        }
    }

    Write-Log -Type INFO -Message ("[Export-TenantStatsJson] Snapshot table counts before export: AllRecipients={0}; AllMailboxes={1}; MailboxFullDetails={2}; PrimaryMailboxStats={3}; ArchiveMailboxStats={4}; InactiveMailboxDetails={5}" -f (Get-TenantStatsExportCount -Key 'AllRecipients'), (Get-TenantStatsExportCount -Key 'AllMailboxes'), (Get-TenantStatsExportCount -Key 'MailboxFullDetails'), (Get-TenantStatsExportCount -Key 'PrimaryMailboxStats'), (Get-TenantStatsExportCount -Key 'ArchiveMailboxStats'), (Get-TenantStatsExportCount -Key 'InactiveMailboxDetails')) -ExportFileLocation $ExportDetails

    $snapshot = Convert-ArrayaLegacyTenantStatsToSnapshot `
        -TenantStatsHash $TenantStatsHash `
        -Metadata $metadata `
        -CollectionPlan $collectionPlan `
        -Diagnostics $diagnostics

    Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $Path
}

function Import-TenantStatsJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        return $null
    }

    $snapshotContext = Import-ArrayaTenantSnapshotContext -Path $Path -Purpose Export
    $script:LoadedTenantSnapshot = $snapshotContext.Snapshot

    return $snapshotContext.LegacyData
}

#region HTML Report Helpers
# HTML helper implementations are maintained in New-TenantHtmlReport.ps1
#region Configuration and Defaults

# Default thresholds for risk detection
$script:DefaultThresholds = @{
    LicenseUtilization = 85           # Percent
    MailboxSizeGB = 50                # GB
    ArchiveSizeGB = 50                # GB
    SharePointSiteGB = 1024           # GB (1 TB)
    OneDriveSiteGB = 1024             # GB (1 TB)
    DeviceStaleMonths = 6             # Months
    DeviceCompliancePercent = 80      # Percent
    InactiveMailboxAgeMonths = 12     # Months
    DomainVerificationRequired = $true
    MXRecordValidation = $true
}

# Merge custom thresholds if provided
$thresholdsVariable = Get-Variable -Name Thresholds -ErrorAction SilentlyContinue
if ($thresholdsVariable -and $null -ne $thresholdsVariable.Value) {
    $Thresholds = $thresholdsVariable.Value
    foreach ($key in $Thresholds.Keys) {
        $script:DefaultThresholds[$key] = $Thresholds[$key]
    }
}

# Available sections and their hashtable keys
$script:SectionMapping = @{
    'Licenses' = 'LicenseSKUs'
    'Recipients' = 'AllRecipients'
    'Mailboxes' = 'MailboxFullDetails'
    'InactiveMailboxes' = 'InactiveMailboxDetails'
    'SharePoint' = 'SharePoint'
    'OneDrive' = 'OneDrive'
    'Domains' = 'Domains'
    'DomainsDNS' = 'Domains'
    'DomainsRecipients' = 'Domains'
    'Devices' = 'DeviceDetails'
    'Teams' = 'AllTeams'
    'MailFlow' = 'MailFlowConnectors'
    'RemoteDomains' = 'RemoteDomains'
    'SecureScore' = 'SecuritySecureScore'
    'ExchangeHybrid' = 'HybridConfiguration'
    'Federation' = 'FederationConfiguration'
}

#endregion


########################################################
# Initialization (Beginning)
########################################################

# Import Office 365 Custom Module
try {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Verbose "Importing Office 365 Custom Module for PowerShell Version 5.1"
        #Import-Module -Name Office365Custom
        
    } else {
        Write-Verbose "Importing Office 365 Custom Module for PowerShell Version 7"
        #Import-Module -Name Office365Custom -UseWindowsPowerShell
    }
}
catch {
    Write-Error "You may need to download and install the custom module 'Office365CustomModule' from Aaron. Will need to install and place in an appropriate PS Path."
    throw $_
}

# Hash table and per-run caches
$script:tenantStatsHash = @{}
$script:MailboxUsageGraphLookup = $null
$script:UnifiedGroupsInventoryCache = $null
$script:Office365GroupsActivityMailboxLookup = $null
$script:AssessmentContext = $null

# Load HTML/report helper functions in script scope so they remain available across the run.
if (Test-Path -Path $tenantHtmlHelperFunctionsPath) {
    . $tenantHtmlHelperFunctionsPath
}
else {
    Write-Warning "Optional HTML helper function script not found: $tenantHtmlHelperFunctionsPath."
}
if (Test-Path -Path $tenantHtmlReportPath) {
    . $tenantHtmlReportPath
}
else {
    Write-Warning "Optional HTML helper script not found: $tenantHtmlReportPath. Built-in technical HTML generation remains available, but PDF export helpers may be unavailable."
}
$script:TenantHtmlHelpersLoaded = $true

Test-AssessmentHelperCommands -CommandName @(
    'Update-OwnershipGovernanceTables',
    'Update-LicenseClassificationMetadata',
    'Update-AssessmentReportTables',
    'Update-ConfigurationSummaryTables'
)

function Ensure-AssessmentServiceContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret,
        [Parameter(Mandatory = $false)]
        [string]$InitialDomain
    )

    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $mgContext) {
        try {
            if (
                -not [string]::IsNullOrWhiteSpace($TenantId) -and
                -not [string]::IsNullOrWhiteSpace($ClientId) -and
                -not [string]::IsNullOrWhiteSpace($CertificateThumbprint)
            ) {
                Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop | Out-Null
                Write-Log -Type INFO -Message "[Ensure-AssessmentServiceContext] Reconnected Microsoft Graph context using certificate auth in caller scope." -ExportFileLocation $ExportDetails
            }
            elseif (
                -not [string]::IsNullOrWhiteSpace($TenantId) -and
                -not [string]::IsNullOrWhiteSpace($ClientId) -and
                -not [string]::IsNullOrWhiteSpace($ClientSecret)
            ) {
                $secureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
                $clientSecretCredential = [System.Management.Automation.PSCredential]::new($ClientId, $secureClientSecret)
                Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $clientSecretCredential -NoWelcome -ErrorAction Stop | Out-Null
                Write-Log -Type INFO -Message "[Ensure-AssessmentServiceContext] Reconnected Microsoft Graph context using client secret auth in caller scope." -ExportFileLocation $ExportDetails
            }
        }
        catch {
            Write-Log -Type WARNING -Message "[Ensure-AssessmentServiceContext] Unable to reconnect Microsoft Graph context in caller scope: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }
    }

    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($mgContext) {
        try {
            $graphAccessToken = Get-MgAccessToken -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace([string]$graphAccessToken)) {
                $global:GraphToken = [string]$graphAccessToken
                $global:GraphHeaders = @{
                    'Content-Type'     = 'application/json'
                    'Authorization'    = "Bearer $graphAccessToken"
                    'ConsistencyLevel' = 'eventual'
                }
            }
        }
        catch {
            Write-Log -Type WARNING -Message "[Ensure-AssessmentServiceContext] Unable to cache Graph access token for REST fallback headers: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }
    }
    else {
        $graphMessage = "[Ensure-AssessmentServiceContext] Microsoft Graph context is unavailable after connection bootstrap."
        Write-Log -Type ERROR -Message $graphMessage -ExportFileLocation $ExportDetails
        throw $graphMessage
    }

    $hasExoMailboxCommand = [bool](Get-Command -Name 'Get-EXOMailbox' -ErrorAction SilentlyContinue)
    $hasUnifiedGroupCommand = [bool](Get-Command -Name 'Get-UnifiedGroup' -ErrorAction SilentlyContinue)
    if (-not ($hasExoMailboxCommand -and $hasUnifiedGroupCommand)) {
        try {
            if (
                -not [string]::IsNullOrWhiteSpace($ClientId) -and
                -not [string]::IsNullOrWhiteSpace($CertificateThumbprint)
            ) {
                $organization = $InitialDomain
                if ([string]::IsNullOrWhiteSpace($organization) -and $mgContext) {
                    try {
                        $organizationResponse = Get-ArrayaGraphResource -Uri 'https://graph.microsoft.com/v1.0/organization' -Activity 'Resolving tenant initial domain'
                        $organizationObject = @($organizationResponse) | Select-Object -First 1
                        if ($organizationObject -and $organizationObject.PSObject.Properties['verifiedDomains']) {
                            $initialDomain = @($organizationObject.verifiedDomains | Where-Object { $_.isInitial -eq $true } | Select-Object -First 1)
                            if ($initialDomain -and $initialDomain[0].name) {
                                $organization = [string]$initialDomain[0].name
                            }
                        }
                    }
                    catch {}
                }

                if (-not [string]::IsNullOrWhiteSpace($organization)) {
                    Connect-ExchangeOnline -AppId $ClientId -Organization $organization -CertificateThumbprint $CertificateThumbprint -ShowBanner:$false -ErrorAction Stop | Out-Null
                    Write-Log -Type INFO -Message "[Ensure-AssessmentServiceContext] Reconnected Exchange Online cmdlets in caller scope for collector visibility." -ExportFileLocation $ExportDetails
                }
            }
        }
        catch {
            Write-Log -Type WARNING -Message "[Ensure-AssessmentServiceContext] Unable to reconnect Exchange Online cmdlets in caller scope: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }
    }

    $hasExoMailboxCommand = [bool](Get-Command -Name 'Get-EXOMailbox' -ErrorAction SilentlyContinue)
    $hasUnifiedGroupCommand = [bool](Get-Command -Name 'Get-UnifiedGroup' -ErrorAction SilentlyContinue)
    if (-not ($hasExoMailboxCommand -and $hasUnifiedGroupCommand)) {
        $exchangeMessage = "[Ensure-AssessmentServiceContext] Exchange Online cmdlets are unavailable after connection bootstrap."
        Write-Log -Type ERROR -Message $exchangeMessage -ExportFileLocation $ExportDetails
        throw $exchangeMessage
    }
}

$connectionResult = $null
$defaultTenantDisplayName = 'Tenant'

if ($runExportOnly) {
    $resolvedTenantStatsJsonPath = [System.IO.Path]::GetFullPath($TenantStatsJsonPath)
    $loadedTenantStats = Import-TenantStatsJson -Path $resolvedTenantStatsJsonPath
    if (-not $loadedTenantStats) {
        throw "Export only mode could not load tenant data snapshot: $resolvedTenantStatsJsonPath"
    }
    if (-not ($loadedTenantStats -is [hashtable])) {
        throw "Export only mode requires a hashtable-compatible tenant data snapshot. Snapshot format was not compatible: $resolvedTenantStatsJsonPath"
    }

    $script:tenantStatsHash = $loadedTenantStats
    if (
        $script:LoadedTenantSnapshot -and
        $script:LoadedTenantSnapshot.Contains('CollectionPlan') -and
        $script:LoadedTenantSnapshot['CollectionPlan'] -is [System.Collections.IDictionary]
    ) {
        $script:SnapshotCollectionPlan = [ordered]@{}
        foreach ($entry in $script:LoadedTenantSnapshot['CollectionPlan'].GetEnumerator()) {
            $script:SnapshotCollectionPlan[[string]$entry.Key] = $entry.Value
        }
    }

    if (
        $script:tenantStatsHash.ContainsKey('TenantInfo') -and
        $script:tenantStatsHash['TenantInfo'] -and
        $script:tenantStatsHash['TenantInfo'].PSObject.Properties['DisplayName'] -and
        -not [string]::IsNullOrWhiteSpace([string]$script:tenantStatsHash['TenantInfo'].DisplayName)
    ) {
        $defaultTenantDisplayName = [string]$script:tenantStatsHash['TenantInfo'].DisplayName
    }
}
else {
    if ($SkipAuth) {
        Write-Host "Skipping authentication bootstrap and validating existing sessions..." -ForegroundColor Cyan
        $existingGraphContext = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $existingGraphContext) {
            throw "SkipAuth was requested, but no existing Microsoft Graph session was found. Connect first, then rerun with -SkipAuth."
        }

        $hasExoMailboxCommand = [bool](Get-Command -Name 'Get-EXOMailbox' -ErrorAction SilentlyContinue)
        $hasUnifiedGroupCommand = [bool](Get-Command -Name 'Get-UnifiedGroup' -ErrorAction SilentlyContinue)
        if (-not ($hasExoMailboxCommand -and $hasUnifiedGroupCommand)) {
            throw "SkipAuth was requested, but Exchange Online cmdlets are not available in the current session. Connect to Exchange Online first, then rerun with -SkipAuth."
        }

        $connectionResult = [pscustomobject]@{
            Graph          = $true
            ExchangeOnline = $true
            AuthenticationType = 'ExistingSession'
            InitialDomain  = $null
        }
    }
    else {
        # Connect to Microsoft Office 365 Services
        $resolvedAuthMode = if ($PSBoundParameters.ContainsKey('AuthMode') -and -not [string]::IsNullOrWhiteSpace($AuthMode)) {
            $AuthMode
        }
        elseif (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
            'Certificate'
        }
        elseif (-not [string]::IsNullOrWhiteSpace($ClientSecret)) {
            'ClientSecret'
        }
        else {
            'Interactive'
        }

        $connectOffice365Params = @{}
        $connectOffice365Params.AuthMode = $resolvedAuthMode
        if ($PSBoundParameters.ContainsKey('TenantId')) { $connectOffice365Params.TenantId = $TenantId }
        if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $connectOffice365Params.CertificateThumbprint = $CertificateThumbprint }
        if ($PSBoundParameters.ContainsKey('ClientId')) { $connectOffice365Params.ClientId = $ClientId }
        if ($PSBoundParameters.ContainsKey('ClientSecret')) {
            if ([string]::IsNullOrWhiteSpace($ClientId)) {
                throw "Client secret authentication requires -ClientId."
            }

            $secureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            $connectOffice365Params.ClientSecretCredential = [System.Management.Automation.PSCredential]::new($ClientId, $secureClientSecret)
        }
        $connectionResult = Connect-Office365 @connectOffice365Params -ErrorAction Stop
        if (
            -not $connectionResult -or
            -not $connectionResult.Graph -or
            -not $connectionResult.ExchangeOnline
        ) {
            throw "Authentication bootstrap did not complete successfully. Graph=$($connectionResult.Graph); ExchangeOnline=$($connectionResult.ExchangeOnline)"
        }
    }

    # Get default tenant display name from live connection
    $defaultOrganization = $null
    try {
        $defaultOrganization = Get-AssessmentTenantOrganization
    } catch {}
    $defaultTenantDisplayName = if ($defaultOrganization -and $defaultOrganization.DisplayName) { $defaultOrganization.DisplayName } else { 'Tenant' }

    $initialDomainForContext = $null
    if ($connectionResult -and $connectionResult.PSObject.Properties['InitialDomain'] -and $connectionResult.InitialDomain) {
        $initialDomainForContext = [string]$connectionResult.InitialDomain
    }
    Ensure-AssessmentServiceContext -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ClientSecret $ClientSecret -InitialDomain $initialDomainForContext
}

#Get Export Path
$profileFileTagSource = if ([string]::IsNullOrWhiteSpace($effectiveOutputProfileLabel)) { $OutputProfile } else { $effectiveOutputProfileLabel }
$profileFileTag = if ([string]::IsNullOrWhiteSpace($profileFileTagSource)) { 'Profile' } else { ($profileFileTagSource -replace '[^A-Za-z0-9_-]', '') }
$defaultReportFileName = ("{0} Tenant Discovery Report-{1}" -f $defaultTenantDisplayName, $profileFileTag)
$outputFolderProfileLabel = if ($isMergedOutputProfileSelection) { $effectiveOutputProfileLabel } else { $OutputProfile }
$defaultOutputRoot = Get-ArrayaAssessmentOutputRoot -FallbackPath $PSScriptRoot
$resolvedExportTargetPath = Resolve-AssessmentExportTargetPath `
    -RequestedExportPath $ExportPath `
    -DefaultOutputRoot $defaultOutputRoot `
    -TenantDisplayName $defaultTenantDisplayName `
    -OutputProfileFolderLabel $outputFolderProfileLabel
$ExportDetails = Get-ExportPath -FileName $defaultReportFileName -UserInputPath $resolvedExportTargetPath

# Initialize list to store all discovery errors
$global:AllDiscoveryErrors = New-Object System.Collections.Generic.List[pscustomobject]

# Resolve collection depth and output behavior from the selected profile.
Write-Host "Output profile: $effectiveOutputProfileLabel (scope: $reportingMode)" -ForegroundColor Green

$script:CollectionDepthPolicy = Get-ArrayaCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($reportingMode))

$collectSecureScoreMappings = ($effectiveGenerateBestPracticesHtml -or $effectiveGenerateWorkbook -or $effectiveGenerateJson)
if ($script:CollectionDepthPolicy.PSObject.Properties['CollectSecureScoreMappings']) {
    $script:CollectionDepthPolicy | Add-Member -MemberType NoteProperty -Name CollectSecureScoreMappings -Value ([bool]$collectSecureScoreMappings) -Force
}

# Tenant-to-tenant migration profile needs richer user licensing detail even in combined mode.
if ($OutputProfile -eq 'TenantToTenantMigration' -or $effectiveOutputProfileLabel -match '(?i)\bTenantToTenantMigration\b') {
    $script:CollectionDepthPolicy | Add-Member -MemberType NoteProperty -Name CollectExtendedGraphEnrichment -Value $true -Force
}

$script:ProfileCollectionPlan = [ordered]@{
    CollectExchangeRecipients        = $true
    CollectEmailActivityDetails      = ($effectiveGenerateTechnicalHtml -or $effectiveGenerateWorkbook -or $effectiveGenerateJson)
    CollectExchangeGroups            = $true
    CollectMailFlowRulesConnectors   = $true
    CollectPublicFolders             = $true
    CollectThirdPartySpamFiltering   = $true
    CollectSmtpRelayConfiguration    = $true
    CollectTeamsDetails              = ($effectiveGenerateTechnicalHtml -or $effectiveGenerateWorkbook -or $effectiveGenerateBestPracticesHtml -or $effectiveGenerateJson)
    CollectTeamsVoiceDetails         = $true
    CollectUnifiedGroups             = $true
    BuildOwnershipGovernanceTables   = ($effectiveGenerateTechnicalHtml -or $effectiveGenerateBestPracticesHtml -or $effectiveGenerateWorkbook -or $effectiveGenerateJson)
    BuildAssessmentReportTables      = ($effectiveGenerateBestPracticesHtml -or $effectiveGenerateWorkbook -or $effectiveGenerateQuestionnaire -or $effectiveGenerateJson)
    BuildConfigurationSummaryTables  = [bool]$effectiveGenerateJson
    BuildLicenseClassificationMetadata = ($effectiveGenerateWorkbook -or $effectiveGenerateTechnicalHtml -or $effectiveGenerateBestPracticesHtml -or $effectiveGenerateQuestionnaire -or $effectiveGenerateJson)
}

if (-not $isMergedOutputProfileSelection) {
    switch ($OutputProfile) {
        'ExecutiveLevel' {
            # Best-practices only profile: trim technical/deep transport collectors.
            $script:ProfileCollectionPlan.CollectExchangeRecipients = $false
            $script:ProfileCollectionPlan.CollectEmailActivityDetails = $true
            $script:ProfileCollectionPlan.CollectExchangeGroups = $false
            $script:ProfileCollectionPlan.CollectMailFlowRulesConnectors = $false
            $script:ProfileCollectionPlan.CollectPublicFolders = $false
            $script:ProfileCollectionPlan.CollectThirdPartySpamFiltering = $false
            $script:ProfileCollectionPlan.CollectSmtpRelayConfiguration = $false
            $script:ProfileCollectionPlan.CollectUnifiedGroups = $false
        }
    }
}

$script:SnapshotCollectionPlan = [ordered]@{
    Exchange = [ordered]@{
        Status = if (
            $script:ProfileCollectionPlan.CollectExchangeRecipients -or
            $script:ProfileCollectionPlan.CollectExchangeGroups -or
            $script:ProfileCollectionPlan.CollectMailFlowRulesConnectors -or
            $script:ProfileCollectionPlan.CollectPublicFolders -or
            $script:ProfileCollectionPlan.CollectEmailActivityDetails
        ) { 'Collected' } else { 'Skipped' }
        Collectors = [ordered]@{
            ExchangeRecipients      = [bool]$script:ProfileCollectionPlan.CollectExchangeRecipients
            ExchangeGroups          = [bool]$script:ProfileCollectionPlan.CollectExchangeGroups
            MailFlowRulesConnectors = [bool]$script:ProfileCollectionPlan.CollectMailFlowRulesConnectors
            PublicFolders           = [bool]$script:ProfileCollectionPlan.CollectPublicFolders
            EmailActivity           = [bool]$script:ProfileCollectionPlan.CollectEmailActivityDetails
        }
        NotCollectedReason = if (
            $script:ProfileCollectionPlan.CollectExchangeRecipients -or
            $script:ProfileCollectionPlan.CollectExchangeGroups -or
            $script:ProfileCollectionPlan.CollectMailFlowRulesConnectors -or
            $script:ProfileCollectionPlan.CollectPublicFolders -or
            $script:ProfileCollectionPlan.CollectEmailActivityDetails
        ) { $null } else { 'All Exchange collectors were disabled by output profile policy.' }
    }
    Identity = [ordered]@{
        Status = 'Collected'
        Collectors = [ordered]@{
            Users               = $true
            Admins              = $true
            Devices             = $true
            ConditionalAccess   = $true
            Authentication      = $true
            SecureScore         = $true
            LicenseSkus         = $true
            Domains             = $true
        }
        NotCollectedReason = $null
    }
    Collaboration = [ordered]@{
        Status = if (
            $script:ProfileCollectionPlan.CollectUnifiedGroups -or
            $script:ProfileCollectionPlan.CollectTeamsDetails -or
            $script:ProfileCollectionPlan.CollectTeamsVoiceDetails
        ) { 'Collected' } else { 'Skipped' }
        Collectors = [ordered]@{
            UnifiedGroups  = [bool]$script:ProfileCollectionPlan.CollectUnifiedGroups
            TeamsInventory = [bool]$script:ProfileCollectionPlan.CollectTeamsDetails
            TeamsVoice     = [bool]$script:ProfileCollectionPlan.CollectTeamsVoiceDetails
            SharePointSite = $true
            Teams          = [bool]$script:ProfileCollectionPlan.CollectTeamsDetails
        }
        NotCollectedReason = if (
            $script:ProfileCollectionPlan.CollectUnifiedGroups -or
            $script:ProfileCollectionPlan.CollectTeamsDetails -or
            $script:ProfileCollectionPlan.CollectTeamsVoiceDetails
        ) { $null } else { 'Collaboration enrichment collectors were disabled by output profile policy.' }
    }
    Security = [ordered]@{
        Status = if (
            $script:ProfileCollectionPlan.CollectThirdPartySpamFiltering -or
            $script:ProfileCollectionPlan.CollectSmtpRelayConfiguration
        ) { 'Collected' } else { 'Partial' }
        Collectors = [ordered]@{
            ThirdPartySpamFiltering = [bool]$script:ProfileCollectionPlan.CollectThirdPartySpamFiltering
            SMTPRelayConfiguration  = [bool]$script:ProfileCollectionPlan.CollectSmtpRelayConfiguration
            SecureScore             = $true
            ConditionalAccess       = $true
            Authentication          = $true
        }
        NotCollectedReason = $null
    }
    Tenant = [ordered]@{
        Status = 'Collected'
        Collectors = [ordered]@{
            TenantInfo          = $true
            Federation          = $true
            AdConnect           = $true
            OwnershipGovernance = [bool]$script:ProfileCollectionPlan.BuildOwnershipGovernanceTables
            ConfigurationTables = [bool]$script:ProfileCollectionPlan.BuildConfigurationSummaryTables
            AssessmentTables    = [bool]$script:ProfileCollectionPlan.BuildAssessmentReportTables
        }
        NotCollectedReason = $null
    }
}

if (
    $runExportOnly -and
    -not (
        $script:LoadedTenantSnapshot -and
        $script:LoadedTenantSnapshot.Contains('CollectionPlan') -and
        $script:LoadedTenantSnapshot['CollectionPlan'] -is [System.Collections.IDictionary]
    )
) {
    foreach ($domainKey in $script:SnapshotCollectionPlan.Keys) {
        $script:SnapshotCollectionPlan[$domainKey].Status = 'NotCollected'
        $script:SnapshotCollectionPlan[$domainKey].NotCollectedReason = 'Export-only mode consumed a previously collected snapshot.'
    }
}

Write-Log -Type INFO -Message ("Collection depth policy: Mode={0}; EntraDeep={1}; GroupCounts={2}; GroupLicenses={3}; SSOAppDetails={4}; ExtendedGraphEnrichment={5}; SecureScoreMappings={6}" -f $script:CollectionDepthPolicy.ReportingMode, $script:CollectionDepthPolicy.CollectEntraGroupDeepDetails, $script:CollectionDepthPolicy.CollectEntraGroupMemberCounts, $script:CollectionDepthPolicy.CollectEntraGroupLicenseChecks, $script:CollectionDepthPolicy.CollectSsoApplicationDetails, $script:CollectionDepthPolicy.CollectExtendedGraphEnrichment, $script:CollectionDepthPolicy.CollectSecureScoreMappings) -ExportFileLocation $ExportDetails
Write-Log -Type INFO -Message ("Profile collection plan ({0}): ExchangeRecipients={1}; EmailActivity={2}; ExchangeGroups={3}; MailFlow={4}; PublicFolders={5}; SpamFiltering={6}; SMTPRelay={7}; TeamsDetails={8}; TeamsVoice={9}; UnifiedGroups={10}; OwnershipTables={11}; AssessmentTables={12}; ConfigSummaryTables={13}; LicenseMetadata={14}" -f $effectiveOutputProfileLabel, $script:ProfileCollectionPlan.CollectExchangeRecipients, $script:ProfileCollectionPlan.CollectEmailActivityDetails, $script:ProfileCollectionPlan.CollectExchangeGroups, $script:ProfileCollectionPlan.CollectMailFlowRulesConnectors, $script:ProfileCollectionPlan.CollectPublicFolders, $script:ProfileCollectionPlan.CollectThirdPartySpamFiltering, $script:ProfileCollectionPlan.CollectSmtpRelayConfiguration, $script:ProfileCollectionPlan.CollectTeamsDetails, $script:ProfileCollectionPlan.CollectTeamsVoiceDetails, $script:ProfileCollectionPlan.CollectUnifiedGroups, $script:ProfileCollectionPlan.BuildOwnershipGovernanceTables, $script:ProfileCollectionPlan.BuildAssessmentReportTables, $script:ProfileCollectionPlan.BuildConfigurationSummaryTables, $script:ProfileCollectionPlan.BuildLicenseClassificationMetadata) -ExportFileLocation $ExportDetails

#Global Start Time for Script
$global:InitialStart = Get-Date
Sync-CollectorModuleRuntimeContext

function Update-ExchangeGovernanceTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory = $true)]
        [string]$DetailLevel
    )

    if (-not $TenantStatsHash.ContainsKey('InboxRulesExternalForwarding')) { $TenantStatsHash['InboxRulesExternalForwarding'] = @{} }
    if (-not $TenantStatsHash.ContainsKey('InboxRuleForwardingSummary')) { $TenantStatsHash['InboxRuleForwardingSummary'] = @{} }
    if (-not $TenantStatsHash.ContainsKey('SharedMailboxGovernanceSummary')) { $TenantStatsHash['SharedMailboxGovernanceSummary'] = @{} }
    if (-not $TenantStatsHash.ContainsKey('ForwardingPolicySummary')) { $TenantStatsHash['ForwardingPolicySummary'] = @{} }

    $allMailboxRows = if ($TenantStatsHash.ContainsKey('AllMailboxes') -and $TenantStatsHash['AllMailboxes'] -is [System.Collections.IDictionary]) { @($TenantStatsHash['AllMailboxes'].Values) } else { @() }
    $mailboxFullRows = if ($TenantStatsHash.ContainsKey('MailboxFullDetails') -and $TenantStatsHash['MailboxFullDetails'] -is [System.Collections.IDictionary]) { @($TenantStatsHash['MailboxFullDetails'].Values) } else { @() }
    $sharedMailboxRows = @(
        $allMailboxRows | Where-Object {
            $_ -and $_.PSObject.Properties['RecipientTypeDetails'] -and ([string]$_.RecipientTypeDetails -match 'SharedMailbox')
        }
    )
    $oversizedSharedMailboxes = @(
        $mailboxFullRows | Where-Object {
            $_ -and $_.PSObject.Properties['RecipientTypeDetails'] -and ([string]$_.RecipientTypeDetails -match 'SharedMailbox') -and
            $_.PSObject.Properties['TotalItemSizeGB'] -and $null -ne $_.TotalItemSizeGB -and ([double]$_.TotalItemSizeGB -gt 50)
        }
    )
    $ownerSignalMissing = @(
        $sharedMailboxRows | Where-Object {
            $grantSendOnBehalf = if ($_.PSObject.Properties['GrantSendOnBehalfTo']) { [string]$_.GrantSendOnBehalfTo } else { '' }
            [string]::IsNullOrWhiteSpace($grantSendOnBehalf)
        }
    )
    $TenantStatsHash['SharedMailboxGovernanceSummary']['Summary'] = [pscustomobject]@{
        SharedMailboxCount          = $sharedMailboxRows.Count
        OversizedSharedMailboxes    = $oversizedSharedMailboxes.Count
        SharedMailboxesWithoutOwnerSignal = $ownerSignalMissing.Count
    }

    $remoteDomainRows = if ($TenantStatsHash.ContainsKey('RemoteDomains') -and $TenantStatsHash['RemoteDomains'] -is [System.Collections.IDictionary]) { @($TenantStatsHash['RemoteDomains'].Values) } else { @() }
    $remoteDomainsWithForwardingEnabled = @(
        $remoteDomainRows |
            Where-Object {
                $_ -and
                $_.PSObject.Properties['AutoForwardEnabled'] -and
                $_.AutoForwardEnabled -eq $true
            }
    )
    $defaultRemoteDomain = @(
        $remoteDomainRows |
            Where-Object {
                $identity = if ($_.PSObject.Properties['Identity']) { [string]$_.Identity } else { '' }
                $domainName = if ($_.PSObject.Properties['DomainName']) { [string]$_.DomainName } else { '' }
                $identity -match '^default$' -or $domainName -eq '*'
            }
    ) | Select-Object -First 1
    $defaultRemoteDomainForwarding = if ($defaultRemoteDomain) {
        if ($defaultRemoteDomain.PSObject.Properties['AutoForwardEnabled']) { [bool]$defaultRemoteDomain.AutoForwardEnabled } else { $null }
    } else {
        $null
    }

    $hostedOutboundPolicies = @()
    $hostedOutboundRules = @()
    $forwardingPolicyCollectionState = 'Unavailable'
    try {
        if (Get-Command -Name 'Get-HostedOutboundSpamFilterPolicy' -ErrorAction SilentlyContinue) {
            $hostedOutboundPolicies = @(
                Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop |
                    Select-Object Name, Identity, IsDefault, AutoForwardingMode
            )
            $forwardingPolicyCollectionState = 'Collected'
        }
    }
    catch {
        $forwardingPolicyCollectionState = 'PolicyLookupFailed'
        Write-Log -Type DEBUG -Message "[Update-ExchangeGovernanceTables] Hosted outbound spam filter policy lookup failed: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }

    try {
        if (Get-Command -Name 'Get-HostedOutboundSpamFilterRule' -ErrorAction SilentlyContinue) {
            $hostedOutboundRules = @(
                Get-HostedOutboundSpamFilterRule -ErrorAction Stop |
                    Select-Object Name, HostedOutboundSpamFilterPolicy, State, Priority
            )
            if ($forwardingPolicyCollectionState -eq 'Unavailable') {
                $forwardingPolicyCollectionState = 'RulesOnly'
            }
        }
    }
    catch {
        if ($forwardingPolicyCollectionState -eq 'Unavailable') {
            $forwardingPolicyCollectionState = 'RuleLookupFailed'
        }
        Write-Log -Type DEBUG -Message "[Update-ExchangeGovernanceTables] Hosted outbound spam filter rule lookup failed: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }

    $policiesExplicitlyAllowingAutoForwarding = @(
        $hostedOutboundPolicies |
            Where-Object {
                $mode = if ($_.PSObject.Properties['AutoForwardingMode']) { [string]$_.AutoForwardingMode } else { '' }
                $mode -match '^(On|Automatic)$'
            }
    )
    $policiesRestrictingAutoForwarding = @(
        $hostedOutboundPolicies |
            Where-Object {
                $mode = if ($_.PSObject.Properties['AutoForwardingMode']) { [string]$_.AutoForwardingMode } else { '' }
                $mode -match '^(Off|InternalOnly)$'
            }
    )
    $autoForwardModeSummary = @(
        $hostedOutboundPolicies |
            ForEach-Object {
                $policyName = if ($_.PSObject.Properties['Name']) { [string]$_.Name } elseif ($_.PSObject.Properties['Identity']) { [string]$_.Identity } else { 'UnnamedPolicy' }
                $mode = if ($_.PSObject.Properties['AutoForwardingMode']) { [string]$_.AutoForwardingMode } else { 'Unknown' }
                "{0}={1}" -f $policyName, $mode
            }
    )
    $rulePolicyAssignments = @(
        $hostedOutboundRules |
            ForEach-Object {
                $policyName = if ($_.PSObject.Properties['HostedOutboundSpamFilterPolicy']) { [string]$_.HostedOutboundSpamFilterPolicy } else { '' }
                if (-not [string]::IsNullOrWhiteSpace($policyName)) { $policyName }
            } |
            Sort-Object -Unique
    )

    $TenantStatsHash['ForwardingPolicySummary']['Summary'] = [pscustomobject]@{
        PolicyCollectionState                    = $forwardingPolicyCollectionState
        HostedOutboundPolicyCount                = $hostedOutboundPolicies.Count
        HostedOutboundRuleCount                  = $hostedOutboundRules.Count
        PoliciesExplicitlyAllowingAutoForwarding = $policiesExplicitlyAllowingAutoForwarding.Count
        PoliciesRestrictingAutoForwarding        = $policiesRestrictingAutoForwarding.Count
        PolicyAutoForwardingModes                = if ($autoForwardModeSummary.Count -gt 0) { $autoForwardModeSummary -join '; ' } else { $null }
        PoliciesReferencedByRules                = if ($rulePolicyAssignments.Count -gt 0) { $rulePolicyAssignments -join '; ' } else { $null }
        RemoteDomainCount                        = $remoteDomainRows.Count
        RemoteDomainsAllowingAutoForwarding      = $remoteDomainsWithForwardingEnabled.Count
        RemoteDomainsAllowingAutoForwardingList  = if ($remoteDomainsWithForwardingEnabled.Count -gt 0) {
            @(
                $remoteDomainsWithForwardingEnabled |
                    ForEach-Object {
                        if ($_.PSObject.Properties['Identity']) { [string]$_.Identity } elseif ($_.PSObject.Properties['DomainName']) { [string]$_.DomainName } else { '' }
                    } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique
            ) -join '; '
        } else { $null }
        DefaultRemoteDomainAllowsAutoForwarding  = $defaultRemoteDomainForwarding
    }

    if ($DetailLevel -eq 'minimum') {
        $TenantStatsHash['InboxRuleForwardingSummary']['Summary'] = [pscustomobject]@{
            InspectedMailboxCount = 0
            ExternalForwardingRuleCount = 0
            CollectionState = 'Skipped in minimum mode'
        }
        return
    }

    $acceptedDomains = @()
    if ($TenantStatsHash.ContainsKey('Domains') -and $TenantStatsHash['Domains'] -is [System.Collections.IDictionary]) {
        $acceptedDomains = @(
            $TenantStatsHash['Domains'].Values |
                ForEach-Object { if ($_.PSObject.Properties['Domain']) { [string]$_.Domain } elseif ($_.PSObject.Properties['Id']) { [string]$_.Id } } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { $_.Trim().ToLowerInvariant() } |
                Select-Object -Unique
        )
    }

    $mailboxesToInspect = @(
        $allMailboxRows | Where-Object {
            $_ -and $_.PSObject.Properties['PrimarySmtpAddress'] -and $_.PrimarySmtpAddress -and
            ([string]$_.RecipientTypeDetails -in @('UserMailbox', 'SharedMailbox'))
        }
    )
    $ruleIndex = 0
    foreach ($mailbox in $mailboxesToInspect) {
        $mailboxAddress = [string]$mailbox.PrimarySmtpAddress
        $previousWarningPreference = $WarningPreference
        try {
            $WarningPreference = 'SilentlyContinue'
            $inboxRules = @(Get-InboxRule -Mailbox $mailboxAddress -WarningAction SilentlyContinue -ErrorAction Stop)
            foreach ($rule in $inboxRules) {
                $forwardTargets = @()
                foreach ($propertyName in @('ForwardTo', 'ForwardAsAttachmentTo', 'RedirectTo')) {
                    if ($rule.PSObject.Properties[$propertyName] -and $rule.$propertyName) {
                        $forwardTargets += @($rule.$propertyName | ForEach-Object { [string]$_ })
                    }
                }
                if ($forwardTargets.Count -eq 0) { continue }

                $externalTargets = @(
                    $forwardTargets |
                        ForEach-Object {
                            $addressText = $_
                            $domainPart = $null
                            if ($addressText -match '@') {
                                $domainPart = ($addressText -split '@')[-1].Trim().Trim('>',';').ToLowerInvariant()
                            }
                            if (-not [string]::IsNullOrWhiteSpace($domainPart) -and ($acceptedDomains -notcontains $domainPart)) {
                                $addressText
                            }
                        } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                )
                if ($externalTargets.Count -eq 0) { continue }

                $ruleIndex++
                $TenantStatsHash['InboxRulesExternalForwarding'][("{0:D4}-{1}" -f $ruleIndex, ($mailboxAddress -replace '[^a-zA-Z0-9@._-]', '_'))] = [pscustomobject]@{
                    Mailbox            = $mailboxAddress
                    RuleName           = $rule.Name
                    Enabled            = $rule.Enabled
                    Description        = $rule.Description
                    ExternalTargets    = ($externalTargets -join ',')
                    ForwardTargetCount = $externalTargets.Count
                }
            }
        }
        catch {
            Write-Log -Type DEBUG -Message "[Update-ExchangeGovernanceTables] Inbox rule lookup failed for ${mailboxAddress}: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }
        finally {
            $WarningPreference = $previousWarningPreference
        }
    }

    $TenantStatsHash['InboxRuleForwardingSummary']['Summary'] = [pscustomobject]@{
        InspectedMailboxCount        = $mailboxesToInspect.Count
        ExternalForwardingRuleCount  = $TenantStatsHash['InboxRulesExternalForwarding'].Count
        CollectionState              = 'Collected'
    }
}

function Update-TierBOperationalSummaries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantStatsHash
    )

    if (-not $TenantStatsHash.ContainsKey('SharePointSharingSummary')) { $TenantStatsHash['SharePointSharingSummary'] = @{} }
    if (-not $TenantStatsHash.ContainsKey('DeviceManagementSummary')) { $TenantStatsHash['DeviceManagementSummary'] = @{} }

    $deviceRows = if ($TenantStatsHash.ContainsKey('DeviceDetails') -and $TenantStatsHash['DeviceDetails'] -is [System.Collections.IDictionary]) { @($TenantStatsHash['DeviceDetails'].Values) } else { @() }
    $sharePointRows = if ($TenantStatsHash.ContainsKey('SharePoint') -and $TenantStatsHash['SharePoint'] -is [System.Collections.IDictionary]) { @($TenantStatsHash['SharePoint'].Values) } else { @() }
    $oneDriveRows = if ($TenantStatsHash.ContainsKey('OneDrive') -and $TenantStatsHash['OneDrive'] -is [System.Collections.IDictionary]) { @($TenantStatsHash['OneDrive'].Values) } else { @() }
    $managedDeviceCount = @($deviceRows | Where-Object { $_.PSObject.Properties['IsManaged'] -and $_.IsManaged -eq $true }).Count
    $compliantDeviceCount = @($deviceRows | Where-Object { $_.PSObject.Properties['IsCompliant'] -and $_.IsCompliant -eq $true }).Count
    $unsupportedOsCount = @(
        $deviceRows | Where-Object {
            $os = [string]$_.OperatingSystem
            $version = [string]$_.OperatingSystemVersion
            ($os -match 'Windows' -and $version -match '^10\.0\.(1[0-8]\d{3}|9\d{3})') -or
            ($os -match 'Windows' -and $version -match '^6\.') -or
            ($os -match 'Android' -and $version -match '^[0-9]+(\.[0-9]+)?' -and ([double]($version.Split('.')[0]) -lt 10)) -or
            ($os -match 'iOS' -and $version -match '^[0-9]+(\.[0-9]+)?' -and ([double]($version.Split('.')[0]) -lt 16))
        }
    ).Count
    $TenantStatsHash['DeviceManagementSummary']['Summary'] = [pscustomobject]@{
        TotalDevices          = $deviceRows.Count
        ManagedDevices        = $managedDeviceCount
        UnmanagedDevices      = $deviceRows.Count - $managedDeviceCount
        CompliantDevices      = $compliantDeviceCount
        NonCompliantDevices   = $deviceRows.Count - $compliantDeviceCount
        UnsupportedOsDevices  = $unsupportedOsCount
    }

    $sharePointSummary = [ordered]@{
        TenantSharingCapability          = 'Not collected'
        DefaultSharingLinkType           = 'Not collected'
        DefaultLinkPermission            = 'Not collected'
        FileAnonymousLinkType            = 'Not collected'
        AnonymousLinkExpirationInDays    = 'Not collected'
    }
    if (Get-Command -Name 'Get-SPOTenant' -ErrorAction SilentlyContinue) {
        try {
            $spoTenant = Get-SPOTenant -ErrorAction Stop
            foreach ($propertyName in $sharePointSummary.Keys) {
                if ($spoTenant.PSObject.Properties[$propertyName]) {
                    $sharePointSummary[$propertyName] = $spoTenant.$propertyName
                }
            }
        }
        catch {
            Write-Log -Type DEBUG -Message "[Update-TierBOperationalSummaries] SharePoint tenant sharing summary lookup failed: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }
    }

    $sharePointInventoryRows = @($sharePointRows) + @($oneDriveRows)
    if ($sharePointInventoryRows.Count -gt 0) {
        $sharingCapabilities = @(
            $sharePointInventoryRows |
                ForEach-Object {
                    if ($_.PSObject.Properties['SharingCapability']) { [string]$_.SharingCapability } else { $null }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique
        )
        $defaultSharingLinkTypes = @(
            $sharePointInventoryRows |
                ForEach-Object {
                    if ($_.PSObject.Properties['DefaultSharingLinkType']) { [string]$_.DefaultSharingLinkType } else { $null }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique
        )
        $defaultLinkPermissions = @(
            $sharePointInventoryRows |
                ForEach-Object {
                    if ($_.PSObject.Properties['DefaultLinkPermission']) { [string]$_.DefaultLinkPermission } else { $null }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique
        )

        if ($sharingCapabilities.Count -gt 0 -and [string]$sharePointSummary['TenantSharingCapability'] -eq 'Not collected') {
            $sharePointSummary['TenantSharingCapability'] = ($sharingCapabilities -join ', ')
        }
        if ($defaultSharingLinkTypes.Count -gt 0 -and [string]$sharePointSummary['DefaultSharingLinkType'] -eq 'Not collected') {
            $sharePointSummary['DefaultSharingLinkType'] = ($defaultSharingLinkTypes -join ', ')
        }
        if ($defaultLinkPermissions.Count -gt 0 -and [string]$sharePointSummary['DefaultLinkPermission'] -eq 'Not collected') {
            $sharePointSummary['DefaultLinkPermission'] = ($defaultLinkPermissions -join ', ')
        }
        if ($sharingCapabilities.Count -gt 0 -or $defaultSharingLinkTypes.Count -gt 0 -or $defaultLinkPermissions.Count -gt 0) {
            $sharePointSummary['DerivedFromSiteInventory'] = $true
            $sharePointSummary['DerivedSiteCount'] = $sharePointInventoryRows.Count
        }
    }

    $TenantStatsHash['SharePointSharingSummary']['Summary'] = [pscustomobject]$sharePointSummary
}

########################################################
# Main Execution (Main Block)
########################################################
try {
    if (
        $Host.Name -eq 'ConsoleHost' -and
        -not [Console]::IsInputRedirected -and
        -not [Console]::IsOutputRedirected -and
        -not [Console]::IsErrorRedirected
    ) {
        Clear-Host
    }
} catch {
    Write-Verbose "Skipping Clear-Host in non-interactive session: $($_.Exception.Message)"
}
Write-Host "Microsoft 365 Tenant Assessment" -ForegroundColor Cyan
Write-Host "Progress view: overall step completion is shown after each major task." -ForegroundColor DarkCyan
Write-Host "Legend: cyan=section/progress, green=completed, yellow=warnings/skips." -ForegroundColor DarkCyan

$GraphTest = if ($runExportOnly) {
    'CACHE'
}
elseif (Get-MgContext -ErrorAction SilentlyContinue) {
    'SDK'
}
elseif ($global:GraphHeaders) {
    'REST'
}
else {
    'SDK'
}
$script:CurrentGraphMode = $GraphTest

if ($runExportOnly) {
    Write-Host ("Loaded tenant data snapshot for export: {0}" -f ([System.IO.Path]::GetFullPath($TenantStatsJsonPath))) -ForegroundColor DarkCyan
    if ($script:ProfileCollectionPlan.BuildOwnershipGovernanceTables) {
        Update-OwnershipGovernanceTables -TenantStatsHash $script:tenantStatsHash
    }
    if ($script:ProfileCollectionPlan.BuildLicenseClassificationMetadata) {
        Update-LicenseClassificationMetadata -TenantStatsHash $script:tenantStatsHash
    }
    if ($script:ProfileCollectionPlan.BuildAssessmentReportTables) {
        Update-AssessmentReportTables -TenantStatsHash $script:tenantStatsHash
    }
    if ($script:ProfileCollectionPlan.BuildConfigurationSummaryTables) {
        Update-ConfigurationSummaryTables -TenantStatsHash $script:tenantStatsHash
    }
}
else {
    $baseCollectionSteps = 13 # Exchange(6) + Hybrid(4) + Collaboration(3)
    $identitySteps = if ($GraphTest -eq 'REST') { 2 } else { 13 }
    $combineSteps = 3
    $postProcessingSteps = 4
    $overallCollectionSteps = $baseCollectionSteps + $identitySteps + $combineSteps + $postProcessingSteps
    Initialize-AssessmentProgress -TotalSteps $overallCollectionSteps

    Write-ConsoleSection -Step '1/5' -Title 'Exchange inventory'
    Invoke-ProfileAwareAssessmentStep -Name 'Exchange recipients' -Enabled $script:ProfileCollectionPlan.CollectExchangeRecipients -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-AllRecipientDetails -detailLevel $reportingMode -Context $script:AssessmentContext }
    Invoke-AssessmentProgressStep -Name 'Exchange mailboxes' -ScriptBlock { Get-AllExchangeMailboxDetails -detailLevel $reportingMode -Context $script:AssessmentContext }
    Invoke-ProfileAwareAssessmentStep -Name 'Email activity insights' -Enabled $script:ProfileCollectionPlan.CollectEmailActivityDetails -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-EmailActivityInsights -detailLevel $reportingMode }
    Invoke-ProfileAwareAssessmentStep -Name 'Exchange groups' -Enabled $script:ProfileCollectionPlan.CollectExchangeGroups -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-ExchangeGroupDetails -detailLevel $reportingMode -Context $script:AssessmentContext }
    Invoke-ProfileAwareAssessmentStep -Name 'Mail flow rules/connectors' -Enabled $script:ProfileCollectionPlan.CollectMailFlowRulesConnectors -SkipReason 'Skipped in best-practices-only profile to reduce runtime.' -ScriptBlock { Get-MailFlowRulesandConnectors -detailLevel $reportingMode -Context $script:AssessmentContext }
    Invoke-ProfileAwareAssessmentStep -Name 'Public folders' -Enabled $script:ProfileCollectionPlan.CollectPublicFolders -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-AllPublicFolderDetails -detailLevel $reportingMode -Context $script:AssessmentContext }

    Write-ConsoleSection -Step '2/5' -Title 'Hybrid and configuration'
    Invoke-AssessmentProgressStep -Name 'Exchange hybrid configuration' -ScriptBlock { Get-ExchangeHybridConfiguration -detailLevel $reportingMode -Context $script:AssessmentContext }
    Invoke-AssessmentProgressStep -Name 'Federation/cross-tenant configuration' -ScriptBlock { Get-FederationAndCrossTenantConfiguration }
    Invoke-ProfileAwareAssessmentStep -Name 'Third-party spam filtering configuration' -Enabled $script:ProfileCollectionPlan.CollectThirdPartySpamFiltering -SkipReason 'Requires mail flow connector/rule collection, which is disabled for this profile.' -ScriptBlock { Get-ThirdPartySpamFilteringConfig -Context $script:AssessmentContext }
    Invoke-ProfileAwareAssessmentStep -Name 'SMTP relay configuration' -Enabled $script:ProfileCollectionPlan.CollectSmtpRelayConfiguration -SkipReason 'Requires mail flow connector collection, which is disabled for this profile.' -ScriptBlock { Get-SMTPRelayConfiguration -Context $script:AssessmentContext }

    Write-ConsoleSection -Step '3/5' -Title 'Identity, devices, and licensing'
    # Determine if using REST or SDK Graph API
    switch ($GraphTest) {
        "REST" {
            Write-Verbose "Attempting to use Microsoft Graph REST API for Tenant Object and License details"
            Invoke-AssessmentProgressStep -Name 'Graph user statistics' -ScriptBlock { Get-GraphUserStats -Context $script:AssessmentContext }
            Invoke-AssessmentProgressStep -Name 'Entra groups (REST)' -ScriptBlock { Get-EntraIDGroups -detailLevel $reportingMode -GraphAuthType REST -Context $script:AssessmentContext }
         }
        "SDK" {
            Write-Verbose "Attempting to use Microsoft Graph SDK for Tenant Object and License details"
            Invoke-AssessmentProgressStep -Name 'Conditional Access policies' -ScriptBlock { Get-ConditionalAccessPoliciesReport -detailLevel $reportingMode }
            Invoke-AssessmentProgressStep -Name 'License SKUs' -ScriptBlock { Get-AllLicenseSKUs }
            Invoke-AssessmentProgressStep -Name 'Users' -ScriptBlock { Get-AllUserDetails -detailLevel $reportingMode }
            Invoke-ProfileAwareAssessmentStep -Name 'Teams voice details' -Enabled $script:ProfileCollectionPlan.CollectTeamsVoiceDetails -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-TeamsVoiceDetails }
            Invoke-AssessmentProgressStep -Name 'Entra groups (SDK)' -ScriptBlock { Get-EntraIDGroups -detailLevel $reportingMode -GraphAuthType SDK -Context $script:AssessmentContext }
            Invoke-AssessmentProgressStep -Name 'Domains' -ScriptBlock { Get-AllOffice365Domains }
            Invoke-AssessmentProgressStep -Name 'Admins' -ScriptBlock { Get-AllOffice365Admins }
            Invoke-AssessmentProgressStep -Name 'Devices' -ScriptBlock { Get-AllDevicesReport -detailLevel $reportingMode }
            Invoke-AssessmentProgressStep -Name 'Tenant overview' -ScriptBlock { Get-TenantOverviewInfo }
            Invoke-AssessmentProgressStep -Name 'Authentication/SSO configuration' -ScriptBlock { Get-AuthenticationConfiguration -detailLevel $reportingMode }
            Invoke-AssessmentProgressStep -Name 'AD Connect sync details' -ScriptBlock { Get-AdConnectSyncDetails }
            Invoke-AssessmentProgressStep -Name 'MFA registration details' -ScriptBlock { Get-MfaRegistrationDetails }
            Invoke-AssessmentProgressStep -Name 'Secure Score report' -ScriptBlock { Get-SecuritySecureScoreReport -detailLevel $reportingMode -MostRecent }
         }
    }

    Write-ConsoleSection -Step '4/5' -Title 'Collaboration and SharePoint'
    Invoke-ProfileAwareAssessmentStep -Name 'Unified groups' -Enabled $script:ProfileCollectionPlan.CollectUnifiedGroups -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-AllUnifiedGroups -detailLevel $reportingMode }
    $sharePointDiscoveryService = if ($GraphTest -in @('REST', 'SDK')) {
        'API'
    }
    elseif ($connectionResult -and $connectionResult.SharePointOnline) {
        'SPO'
    }
    else {
        'API'
    }
    Invoke-AssessmentProgressStep -Name "SharePoint/OneDrive sites ($sharePointDiscoveryService)" -ScriptBlock { Get-SharePointAndOneDriveSites -detailLevel $reportingMode -ServiceName $sharePointDiscoveryService }
    $teamsDiscoveryService = if ($GraphTest -in @('SDK', 'REST')) { 'MGGraph' } else { 'Teams' }
    Invoke-ProfileAwareAssessmentStep -Name "Teams inventory ($teamsDiscoveryService)" -Enabled $script:ProfileCollectionPlan.CollectTeamsDetails -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-TeamsDetails -detailLevel $reportingMode -ServiceName $teamsDiscoveryService }
    Write-Host


    # Combine reporting data for all profiles so mailbox/detail stats are available across every output type.
    Write-Host
    Write-Host "Consolidating Discovery Report data for each user / object into one file" -ForegroundColor Black -BackgroundColor Green
    Invoke-AssessmentProgressStep -Name 'Combined user/mailbox reporting' -ScriptBlock { Report-UserAndMailboxStats }
    Invoke-AssessmentProgressStep -Name 'Exchange governance Tier B summaries' -ScriptBlock { Update-ExchangeGovernanceTables -TenantStatsHash $script:tenantStatsHash -DetailLevel $reportingMode }
    Invoke-AssessmentProgressStep -Name 'Operational Tier B summaries' -ScriptBlock { Update-TierBOperationalSummaries -TenantStatsHash $script:tenantStatsHash }

    Invoke-ProfileAwareAssessmentStep -Name 'Ownership governance tables' -Enabled $script:ProfileCollectionPlan.BuildOwnershipGovernanceTables -SkipReason 'Ownership governance table build is disabled for this profile.' -ScriptBlock { Update-OwnershipGovernanceTables -TenantStatsHash $script:tenantStatsHash }
    Invoke-ProfileAwareAssessmentStep -Name 'License classification metadata' -Enabled $script:ProfileCollectionPlan.BuildLicenseClassificationMetadata -SkipReason 'License classification metadata is disabled for this profile.' -ScriptBlock { Update-LicenseClassificationMetadata -TenantStatsHash $script:tenantStatsHash }
    Invoke-ProfileAwareAssessmentStep -Name 'Best-practice assessment tables' -Enabled $script:ProfileCollectionPlan.BuildAssessmentReportTables -SkipReason 'Best-practice table build is disabled for this profile.' -ScriptBlock { Update-AssessmentReportTables -TenantStatsHash $script:tenantStatsHash }
    Invoke-ProfileAwareAssessmentStep -Name 'Configuration summary tables' -Enabled $script:ProfileCollectionPlan.BuildConfigurationSummaryTables -SkipReason 'Configuration summary tables are disabled for this profile.' -ScriptBlock { Update-ConfigurationSummaryTables -TenantStatsHash $script:tenantStatsHash }
    Complete-AssessmentProgress
    Write-CollectorInventoryMatrix -TenantStatsHash $script:tenantStatsHash -ExportFileLocation $ExportDetails
    Write-AssessmentStepMetricsSummary -ExportFileLocation $ExportDetails
}


########################################################
### Export Reports ###
########################################################

#Exclude specific reports from Export
Write-ConsoleSection -Step '5/5' -Title 'Exporting results'
$requiresFilteredExportSnapshot = (-not $effectiveSkipWorkbook)
$ExportTenantStatsHash = $null
if ($requiresFilteredExportSnapshot) {
    $ExportTenantStatsHash = Filter-TenantStatsHash -TenantStatsStore $script:tenantStatsHash -reportingMode $reportingMode -GraphTest $GraphTest
} else {
    Write-Log -Type INFO -Message "Skipping filtered export snapshot build because workbook output is disabled for this profile." -ExportFileLocation $ExportDetails
}
$generatedArtifacts = [ordered]@{}
try {
    if (Test-Path -Path $tenantExportPipelinePath) {
        . $tenantExportPipelinePath
    }
    else {
        throw "Export pipeline script was not found: $tenantExportPipelinePath"
    }

    if (-not (Get-Command -Name Invoke-M365TenantAssessmentExportPipeline -ErrorAction SilentlyContinue)) {
        throw "Export pipeline function is unavailable. Expected loader path: $tenantExportPipelinePath"
    }

    $generatedArtifacts = Invoke-M365TenantAssessmentExportPipeline `
        -TenantStatsHash $script:tenantStatsHash `
        -ExportTenantStatsHash $ExportTenantStatsHash `
        -ExportDetails $ExportDetails `
        -SkipWorkbook $effectiveSkipWorkbook `
        -SkipBestPracticesHtml $effectiveSkipBestPracticesHtml `
        -SkipQuestionnaire $effectiveSkipQuestionnaire `
        -SkipHtmlReport $effectiveSkipHtmlReport `
        -SkipPdfReport $effectiveSkipPdfReport `
        -SkipJsonReport $effectiveSkipJsonReport `
        -OutputProfileLabel $effectiveOutputProfileLabel `
        -ReportingMode $reportingMode `
        -CollectionOnly ([bool]$runCollectionOnly) `
        -ExportOnly ([bool]$runExportOnly) `
        -LegacyScriptRoot $PSScriptRoot
}
catch {
    Write-Log -Type ERROR -Message "Export pipeline execution failed: $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    $generatedArtifacts = [ordered]@{}
}

$runLogDirectory = [System.IO.Path]::GetDirectoryName($ExportDetails)
if ([string]::IsNullOrWhiteSpace($runLogDirectory)) {
    $runLogDirectory = (Get-Location).Path
}
$runLogDirectory = Join-Path -Path $runLogDirectory -ChildPath 'Debugging'
$runLogBaseName = [System.IO.Path]::GetFileNameWithoutExtension($ExportDetails)
if (-not [string]::IsNullOrWhiteSpace($runLogBaseName)) {
    $runLogPath = Join-Path -Path $runLogDirectory -ChildPath ($runLogBaseName + '-FullReportLog.txt')
    if (Test-Path -Path $runLogPath) {
        $generatedArtifacts['Run Log'] = $runLogPath
    }
}

Write-Host ""
if ($global:AllDiscoveryErrors.Count -gt 0) {
    try {
        $errorReportSummary = Export-ArrayaErrorReports -ExportFileLocation $ExportDetails -ErrorData $global:AllDiscoveryErrors -LogReportDirectory $ExportDetails
        if ($errorReportSummary -and -not [string]::IsNullOrWhiteSpace([string]$errorReportSummary.JsonPath)) {
            $generatedArtifacts['Error Report'] = $errorReportSummary.JsonPath
        }
    }
    catch {
        Write-Log -Type ERROR -Message "An error occurred in Exporting the Error Reports. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
}

$scriptEndTime = Get-Date
$totalTime = $scriptEndTime - $global:InitialStart
$totalHours = [math]::Floor($totalTime.TotalHours)
$totalMinutes = $totalTime.Minutes
$totalSeconds = $totalTime.Seconds
$timeString = if ($totalHours -gt 0) {
    "$totalHours hour(s), $totalMinutes minute(s), $totalSeconds second(s)"
} elseif ($totalMinutes -gt 0) {
    "$totalMinutes minute(s), $totalSeconds second(s)"
} else {
    "$totalSeconds second(s)"
}

Write-ConsoleArtifactSummary -Artifacts $generatedArtifacts -DurationText $timeString -CapturedErrorCount $global:AllDiscoveryErrors.Count

$finalMemory = Get-CurrentProcessMemorySnapshot
Write-Log -Type INFO -Message ("Final process memory snapshot: WorkingSetMB={0}; PrivateMB={1}; PagedMB={2}; ManagedHeapMB={3}" -f $finalMemory.WorkingSetMB, $finalMemory.PrivateMB, $finalMemory.PagedMB, $finalMemory.HeapMB) -ExportFileLocation $ExportDetails

if ($StoreTenantStatsGlobal) {
    $resolvedTenantStatsVariableName = if ([string]::IsNullOrWhiteSpace($TenantStatsVariableName)) { 'ArrayaTenantStats' } else { $TenantStatsVariableName }
    Set-Variable -Scope Global -Name $resolvedTenantStatsVariableName -Value $script:tenantStatsHash -Force
    Write-Host ("Published tenant stats to global variable `${0}" -f $resolvedTenantStatsVariableName) -ForegroundColor DarkCyan
    Write-Log -Type INFO -Message ("Published tenant stats to global variable '{0}' for interactive inspection." -f $resolvedTenantStatsVariableName) -ExportFileLocation $ExportDetails
}

########################################################
### End of HTML Report Integration ###
########################################################

Write-Log -Type INFO -Message "COMPLETED: Gathered Tenant Details. Completed Time: $($timeString)" -ExportFileLocation $ExportDetails

# Encourage GC after large export/report generation to reduce retained working set in long-lived shells.
$ExportTenantStatsHash = $null
$script:AssessmentStepMetrics = $null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
