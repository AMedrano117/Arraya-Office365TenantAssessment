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
    'Write-ProgressHelper',
    'Office365Custom\Write-Log',
    'Office365Custom\Capture-ErrorHelper',
    'Office365Custom\Get-ExportPath'
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
    'Get-ArrayaGraphResource',
    'Export-ArrayaGraphReportCsv',
    'Get-ArrayaEntraGroupClassification',
    'Invoke-ArrayaCollectionStepSafe',
    'Export-ArrayaErrorReports',
    'Get-ArrayaCollectionDepthPolicy',
    'Convert-ArrayaObjectToArray',
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
    Import-Module -Name $resolvedCommonManifestPath -Force -ErrorAction Stop
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
$tenantQuestionnairePath = Join-Path -Path $PSScriptRoot -ChildPath 'Export-TenantToTenantQuestionnaireMarkdown.ps1'
$tenantExportPipelinePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\reporting\Invoke-M365TenantAssessmentExportPipeline.ps1'))

########################################################
# Functions
########################################################
# ----------------------------------
# Default Script Helper Functions
# ----------------------------------
function Capture-ErrorHelper {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [AllowNull()]
        [System.Management.Automation.ErrorRecord]$ErrorRecordVar,
        [Parameter(Mandatory=$true)]
        [string]$errorMessage
    )

    # Use Office365Custom implementation when an ErrorRecord is present.
    if ($null -ne $ErrorRecordVar) {
        Office365Custom\Capture-ErrorHelper -ErrorRecordVar $ErrorRecordVar -errorMessage $errorMessage
        return
    }

    # Compatibility fallback: some Write-Log calls emit ERROR without an ErrorRecord.
    Write-Host $errorMessage -ForegroundColor Red
    $currentError = [PSCustomObject]@{
        TimeStamp           = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        ErrorMessage        = $errorMessage
        Commandlet          = $null
        Reason              = $null
        "Exception-Message" = $null
        Exception           = $null
        Recipient           = $null
        TargetObject        = $null
    }

    if (!$global:AllDiscoveryErrors) {
        $global:AllDiscoveryErrors = New-Object System.Collections.Generic.List[pscustomobject]
    }
    $global:AllDiscoveryErrors.Add($currentError)

    return $currentError
}

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Type = "INFO",

        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [switch]$CaptureError,

        [Parameter(Mandatory=$false)]
        [System.Management.Automation.ErrorRecord]$ErrorRecordVar,

        [Parameter(Mandatory=$false)]
        [string]$LogPath,

        [Parameter(Mandatory=$false)]
        [string]$ExportFileLocation

    )
    $effectiveCapture = $CaptureError.IsPresent -or ($Type -eq 'ERROR')
    $hasErrorRecord = $PSBoundParameters.ContainsKey('ErrorRecordVar') -and $null -ne $ErrorRecordVar

    # Use Office365Custom Write-Log whenever signatures align cleanly.
    if ($Type -ne 'ERROR' -or $hasErrorRecord) {
        $logSplat = @{
            Type    = $Type
            Message = $Message
        }
        if ($LogPath) {
            $logSplat.LogPath = $LogPath
        }
        if ($ExportFileLocation) {
            $logSplat.ExportFileLocation = $ExportFileLocation
        }
        if ($effectiveCapture -and $hasErrorRecord) {
            $logSplat.CaptureError = $true
            $logSplat.ErrorRecordVar = $ErrorRecordVar
        }

        Office365Custom\Write-Log @logSplat

        if ($effectiveCapture -and -not $hasErrorRecord) {
            Capture-ErrorHelper -ErrorRecordVar $null -errorMessage $Message | Out-Null
        }
        return
    }

    # Compatibility fallback for ERROR logs without ErrorRecordVar.
    if ($LogPath) {
        if (-not (Test-Path $LogPath)) {
            $null = New-Item -Path $LogPath -ItemType Directory
        }
        try {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "$timestamp : [$Type] $Message"
            $LogFile = $LogPath + "\FullTenantReportLog.txt"
            $logMessage | Out-File -Append -FilePath $LogFile
        } catch {
            Write-Error "Failed to write to log file at '$LogFile': $_"
        }
    }
    elseif ($ExportFileLocation) {
        $directory = [System.IO.Path]::GetDirectoryName($ExportFileLocation)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ExportFileLocation)
        $txtFileName = $baseName + "-FullReportLog.txt"
        $newLogFolder = $baseName + " Reporting"
        $newDirectory = Join-Path -Path $directory -ChildPath $newLogFolder
        $LogFile = Join-Path -Path $newDirectory -ChildPath $txtFileName

        if (-not (Test-Path $newDirectory)) {
            $null = New-Item -Path $newDirectory -ItemType Directory
        }
        try {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "$timestamp : [$Type] $Message"
            $logMessage | Out-File -Append -FilePath $LogFile
        } catch {
            Write-Error "Failed to write to log file at '$LogFile': $_"
        }
    }

    if ($VerbosePreference -eq 'Continue') {
        Write-Verbose $Message
    }

    if ($effectiveCapture) {
        Capture-ErrorHelper -ErrorRecordVar $null -errorMessage $Message | Out-Null
    }
}

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

    foreach ($entry in $Artifacts.GetEnumerator()) {
        if ([string]::IsNullOrWhiteSpace([string]$entry.Value)) {
            continue
        }

        Write-Host ("  {0}: {1}" -f $entry.Key, $entry.Value) -ForegroundColor Gray
    }

    if ($CapturedErrorCount -gt 0) {
        Write-Host ("  Captured errors: {0} (see log/error reports)" -f $CapturedErrorCount) -ForegroundColor Yellow
    }
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

    if (-not (Test-Path -Path $tenantHtmlReportPath)) {
        Write-Warning "Optional HTML helper script not found: $tenantHtmlReportPath. Built-in technical HTML generation remains available, but PDF export helpers may be unavailable."
        return
    }

    . $tenantHtmlReportPath
    $script:TenantHtmlHelpersLoaded = $true
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

    Write-Progress -Id $script:AssessmentProgressId -Activity 'Assessment progress' -Status "[0/$($script:AssessmentProgressState.Total)] Starting" -PercentComplete 0
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

    Write-Progress -Id $script:AssessmentProgressId -Activity 'Assessment progress' -Status "[$current/$total] $Name" -PercentComplete $percent
    try {
        & $ScriptBlock
        Write-Host ("  Overall progress: {0}/{1} ({2}%) - {3}" -f $current, $total, $percent, $Name) -ForegroundColor DarkGray
    }
    finally {
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

    Write-Progress -Id $script:AssessmentProgressId -Activity 'Assessment progress' -Completed
    Clear-TransientGraphProgress
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
        [PSCustomObject]@{ Collector = 'Conditional Access'; Key = 'ConditionalAccessPolicies'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Secure Score'; Key = 'SecuritySecureScore'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Secure Score Actions'; Key = 'SecureScoreActions'; Source = 'Graph'; Consumers = 'Workbook, BestPractices'; Count = 0 }
        [PSCustomObject]@{ Collector = 'License SKUs'; Key = 'LicenseSKUs'; Source = 'Graph'; Consumers = 'Workbook, BestPractices, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Email Activity Top Senders'; Key = 'EmailActivityTopSenders'; Source = 'Graph'; Consumers = 'Workbook, HTML'; Count = 0 }
        [PSCustomObject]@{ Collector = 'Email Activity Top Receivers'; Key = 'EmailActivityTopReceivers'; Source = 'Graph'; Consumers = 'Workbook, HTML'; Count = 0 }
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

function Resolve-ExoStatisticsIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $MailboxObject
    )

    if ($null -eq $MailboxObject) {
        return $null
    }

    if ($MailboxObject -is [string]) {
        return $MailboxObject
    }

    foreach ($propertyName in @('ExchangeGuid', 'ExternalDirectoryObjectId', 'Guid', 'Identity', 'PrimarySmtpAddress', 'UserPrincipalName', 'WindowsEmailAddress', 'Alias', 'DistinguishedName')) {
        $property = $MailboxObject.PSObject.Properties[$propertyName]
        if (-not $property) {
            continue
        }

        $value = $property.Value
        if ($null -eq $value) {
            continue
        }

        $stringValue = [string]$value
        if (-not [string]::IsNullOrWhiteSpace($stringValue)) {
            return $stringValue
        }
    }

    return $null
}

function Get-ExoMailboxStatisticsSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $MailboxObjects,
        [switch]$Archive,
        [Parameter(Mandatory = $false)]
        [string]$ProgressActivity,
        [Parameter(Mandatory = $false)]
        [int]$ProgressId = 0
    )

    $results = New-Object System.Collections.Generic.List[object]
    $failures = New-Object System.Collections.Generic.List[object]
    $normalizedMailboxObjects = New-Object System.Collections.Generic.List[object]

    if ($null -ne $MailboxObjects) {
        foreach ($mailboxItem in $MailboxObjects) {
            $normalizedMailboxObjects.Add($mailboxItem)
        }
    }

    $totalMailboxCount = $normalizedMailboxObjects.Count
    $processedMailboxCount = 0
    $progressStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($mailbox in $normalizedMailboxObjects) {
        $processedMailboxCount++
        if (-not [string]::IsNullOrWhiteSpace($ProgressActivity) -and $totalMailboxCount -gt 0) {
            $label = if ($mailbox -and $mailbox.PSObject.Properties['DisplayName'] -and -not [string]::IsNullOrWhiteSpace([string]$mailbox.DisplayName)) {
                [string]$mailbox.DisplayName
            }
            elseif ($mailbox -and $mailbox.PSObject.Properties['PrimarySmtpAddress'] -and -not [string]::IsNullOrWhiteSpace([string]$mailbox.PrimarySmtpAddress)) {
                [string]$mailbox.PrimarySmtpAddress
            }
            else {
                'Processing mailbox'
            }

            $percentComplete = [math]::Round(($processedMailboxCount / $totalMailboxCount) * 100, 2)
            $elapsedText = if ($progressStopwatch) { $progressStopwatch.Elapsed.ToString('hh\:mm\:ss') } else { '00:00:00' }
            $progressStatus = "[{0} / {1}] {2} | {3} elapsed | results={4}, failures={5}" -f $processedMailboxCount, $totalMailboxCount, $label, $elapsedText, $results.Count, $failures.Count
            Write-Progress -Id $ProgressId -Activity $ProgressActivity -Status $progressStatus -PercentComplete $percentComplete
        }

        $identity = Resolve-ExoStatisticsIdentity -MailboxObject $mailbox
        $displayName = $null
        if ($null -ne $mailbox -and $mailbox.PSObject.Properties['DisplayName']) {
            $displayName = [string]$mailbox.DisplayName
        }

        if ([string]::IsNullOrWhiteSpace($identity)) {
            $failures.Add([PSCustomObject]@{
                DisplayName = $displayName
                Identity    = $null
                Reason      = 'No supported mailbox identity was available.'
            })
            continue
        }

        try {
            $statsParams = @{
                Identity    = $identity
                ErrorAction = 'Stop'
            }

            if ($Archive) {
                $statsParams.Archive = $true
                $statsParams.Properties = 'MailboxGuid'
                $statsParams.IncludeSoftDeletedRecipients = $true
            }
            else {
                $statsParams.IncludeSoftDeletedRecipient = $true
            }

            $statResults = Invoke-QuietCommand -ScriptBlock { Get-EXOMailboxStatistics @statsParams }
            if ($statResults.Count -eq 0) {
                $failures.Add([PSCustomObject]@{
                    DisplayName = $displayName
                    Identity    = $identity
                    Reason      = 'No mailbox statistics were returned.'
                })
                continue
            }

            foreach ($stat in $statResults) {
                $totalItemBytes = Convert-DataSizeToBytes -Value $stat.TotalItemSize
                $deletedItemBytes = Convert-DataSizeToBytes -Value $stat.TotalDeletedItemSize
                if (-not $stat.PSObject.Properties['TotalItemSizeBytes']) {
                    $stat | Add-Member -MemberType NoteProperty -Name 'TotalItemSizeBytes' -Value $totalItemBytes -Force
                }
                if (-not $stat.PSObject.Properties['TotalDeletedItemSizeBytes']) {
                    $stat | Add-Member -MemberType NoteProperty -Name 'TotalDeletedItemSizeBytes' -Value $deletedItemBytes -Force
                }
                $results.Add($stat)
            }
        }
        catch {
            $failures.Add([PSCustomObject]@{
                DisplayName = $displayName
                Identity    = $identity
                Reason      = $_.Exception.Message
            })
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ProgressActivity)) {
        Write-Progress -Id $ProgressId -Activity $ProgressActivity -Completed
    }
    if ($progressStopwatch -and $progressStopwatch.IsRunning) {
        $progressStopwatch.Stop()
    }

    return [PSCustomObject]@{
        Results  = $results.ToArray()
        Failures = $failures.ToArray()
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

function Normalize-GraphReportFieldName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ''
    }

    return (($Name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant())
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

    $normalizedPropertyMap = @{}
    foreach ($property in $Row.PSObject.Properties) {
        $normalizedName = Normalize-GraphReportFieldName -Name $property.Name
        if (-not [string]::IsNullOrWhiteSpace($normalizedName) -and -not $normalizedPropertyMap.ContainsKey($normalizedName)) {
            $normalizedPropertyMap[$normalizedName] = $property.Value
        }
    }

    foreach ($fieldName in $FieldNames) {
        if ($Row.PSObject.Properties[$fieldName]) {
            $rawValue = [string]$Row.PSObject.Properties[$fieldName].Value
            if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
                return $rawValue.Trim()
            }
        }

        $normalizedFieldName = Normalize-GraphReportFieldName -Name $fieldName
        if (
            -not [string]::IsNullOrWhiteSpace($normalizedFieldName) -and
            $normalizedPropertyMap.ContainsKey($normalizedFieldName)
        ) {
            $rawValue = [string]$normalizedPropertyMap[$normalizedFieldName]
            if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
                return $rawValue.Trim()
            }
        }
    }

    return $null
}

function Convert-GraphReportValueToInt64 {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return [int64]0
    }

    if ($Value -is [int64] -or $Value -is [int32] -or $Value -is [long]) {
        return [int64]$Value
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [int64]0
    }

    $text = $text -replace ',', ''
    $styles = [System.Globalization.NumberStyles]::Any
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $parsedInt64 = [int64]0
    if ([int64]::TryParse($text, $styles, $culture, [ref]$parsedInt64)) {
        return $parsedInt64
    }

    $parsedDouble = [double]0
    if ([double]::TryParse($text, $styles, $culture, [ref]$parsedDouble)) {
        return [int64][math]::Round($parsedDouble, 0)
    }

    return [int64]0
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
# Export Script Functions
# ----------------------------------
#Function to get Export Path
function Get-ExportPath {
    [CmdletBinding()]
    param (
        [string]$FileName,  # File name to use (can include -TenantReport or other custom string)
        [string]$DefaultExtension = ".xlsx",
        [string]$UserInputPath
    )

    if ([string]::IsNullOrWhiteSpace($UserInputPath)) {
        # Use Office365Custom implementation for interactive input flow.
        $resolvedName = if ([string]::IsNullOrWhiteSpace($FileName)) { "Report" } else { $FileName }
        return Office365Custom\Get-ExportPath -FileName $resolvedName -DefaultExtension $DefaultExtension
    }

    # Non-interactive shim for automation paths.
    $userInput = $UserInputPath

    # Handle quotes in input
    $userInput = $userInput -replace '"', ''

    if (-not [string]::IsNullOrWhiteSpace($userInput) -and -not [System.IO.Path]::IsPathRooted($userInput)) {
        $userInput = Join-Path -Path (Get-Location).Path -ChildPath $userInput
    }

    # If user input is empty, default to the local non-repo output root
    if ([string]::IsNullOrEmpty($userInput)) {
        $userInput = Get-ArrayaAssessmentOutputRoot -FallbackPath $PSScriptRoot
    }

    # File path processing
    $folderPath = ""
    $inputFileName = ""

    if ((Test-Path $userInput) -and (Get-Item -Path $userInput -ErrorAction SilentlyContinue).PSIsContainer) {
        # If the user input is a folder, assign the folder path, and use the provided $FileName if supplied
        $folderPath = $userInput
        if ([string]::IsNullOrEmpty($FileName)) {
            $fileName = "Report$DefaultExtension"
        } else {
            $fileName = $FileName
        }
    } else {
        # It's a file path (or just a file name), so split and use
        $folderPath = Split-Path -Path $userInput -Parent
        $inputFileName = Split-Path -Path $userInput -Leaf

        if (-not [string]::IsNullOrEmpty($inputFileName)) {
            # User is overriding the FileName via path
            $fileName = $inputFileName
            # If no folder path (i.e., just a file name), use the local output root
            if ([string]::IsNullOrWhiteSpace($folderPath)) {
                $folderPath = Get-ArrayaAssessmentOutputRoot -FallbackPath $PSScriptRoot
            }
        } else {
            # User entered something ambiguous, fallback
            $folderPath = $PSScriptRoot
            if ([string]::IsNullOrEmpty($FileName)) {
                $fileName = "Report$DefaultExtension"
            } else {
                $fileName = $FileName
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($folderPath) -and -not (Test-Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory -Force | Out-Null
    }

    # If folderPath is empty or invalid, default to script root
    if ([string]::IsNullOrEmpty($folderPath) -or !(Test-Path $folderPath)) {
        $folderPath = $PSScriptRoot
    }

    # Ensure the file has the correct extension
    $extension = [IO.Path]::GetExtension($fileName)
    if ([string]::IsNullOrEmpty($extension)) {
        $fileName += $DefaultExtension
    }

    # Full path to return
    $fullPath = Join-Path -Path $folderPath -ChildPath $fileName

    Write-Host "The file will be saved to: $fullPath" -ForegroundColor Green
    return $fullPath
}

# Convert Hash Table to Custom Object Array for Export
function Convert-HashToArray {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [Hashtable]$HashToConvert,
        [Parameter(Mandatory=$False)]
        [String]$tenant
    )

    # Compatibility shim: route legacy helper usage through shared converter.
    return Convert-ArrayaObjectToArray -InputObject $HashToConvert
}

function Filter-TenantStatsHash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantStatsStore,
        [Parameter(Mandatory = $true)]
        [string]$reportingMode,
        [Parameter(Mandatory = $false)]
        $GraphTest
    )

    # Common tables to remove for both 'combined' and 'minimum' detail levels
    $commonTablesToRemove = @(
        'AllMailboxes-MailIdentity',
        'AllMailboxes-UserPrincipalName',
        'AllMailboxes-PrimarySmtpAddress'
    )

    # Determine additional tables to remove based on the reporting mode
    $additionalTables = switch ($reportingMode) {
        "combined" {
            @(
                'Users', 'ArchiveMailboxes', 'ArchiveMailboxStats', 'NonUserMailboxes',
                'PrimaryMailboxStats', 'AllMailboxes', 'InactiveMailboxes', 
                'LitigationHoldMailboxes', 'UnifiedGroups'
            )
        }
        "minimum" {
            @(
                'InactiveMailboxes','ArchiveMailboxStats', 'NonUserMailboxes',
                'PrimaryMailboxStats', 'AllMailboxes', 'LitigationHoldMailboxes',
                'RemoteDomains','UnifiedGroups', 'MailFlowConnectors',
                'PublicFolderPerms', 'AuthenticationConfig', 'TenantInfo', 'SpamFilteringConfig',
                'SMTPRelayConfig', 'FederationConfiguration', 'TeamsVoice', 'MfaRegistrationDetails'
            )
        }
        default {
            @() # No additional tables for other reporting modes
        }
    }
    if ($GraphTest -eq "REST") {
        $additionalTables += @(
            "EmailData", "OneDriveData", "SharePointData", "TeamsUserData", "MailboxUsage", 
            "SignInData", "UserSignIns", "SPOUsage", "YammerUsage"
        )
    }

    # Combine common and additional tables to form the full exclusion list
    $tablesToRemove = $commonTablesToRemove + $additionalTables

    $excludeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in $tablesToRemove) {
        if (-not [string]::IsNullOrWhiteSpace([string]$key)) {
            $null = $excludeSet.Add([string]$key)
        }
    }

    # Build filtered output directly to avoid carrying duplicate hash table state.
    $filteredStatsHash = @{}
    foreach ($entry in $TenantStatsStore.GetEnumerator()) {
        if ($excludeSet.Contains([string]$entry.Key)) {
            continue
        }
        $filteredStatsHash[$entry.Key] = $entry.Value
    }

    # Return the filtered hash table
    return $filteredStatsHash
}

function ConvertTo-ExportFriendlyValue {
    param(
        $Value,
        [int]$Depth = 0
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Depth -ge 3) {
        return '[Nested]'
    }

    if (
        $Value -is [string] -or
        $Value -is [char] -or
        $Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [sbyte] -or
        $Value -is [int16] -or
        $Value -is [uint16] -or
        $Value -is [int32] -or
        $Value -is [uint32] -or
        $Value -is [int64] -or
        $Value -is [uint64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]
    ) {
        return $Value
    }

    if ($Value -is [datetime] -or $Value -is [datetimeoffset]) {
        return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm:ss')
    }

    if ($Value -is [timespan] -or $Value -is [guid] -or $Value -is [uri] -or $Value -is [version] -or $Value -is [enum]) {
        return $Value.ToString()
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $pairs = @()
        foreach ($key in $Value.Keys) {
            $pairs += "$key=$(ConvertTo-ExportFriendlyValue -Value $Value[$key] -Depth ($Depth + 1))"
        }
        return ($pairs -join '; ')
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @()
        foreach ($item in $Value) {
            if ($null -eq $item) { continue }

            if (
                $item -is [string] -or
                $item -is [ValueType]
            ) {
                $items += $item.ToString()
                continue
            }

            $identityValue = $null
            foreach ($identityProperty in @('DisplayName', 'Name', 'Title', 'Domain', 'UserPrincipalName', 'Mail', 'AppId', 'Id', 'SkuFriendlyName', 'SkuPartNumber', 'Value')) {
                $property = $item.PSObject.Properties[$identityProperty]
                if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                    $identityValue = $property.Value
                    break
                }
            }

            if ($null -ne $identityValue) {
                $items += $identityValue.ToString()
            } else {
                $items += (ConvertTo-ExportFriendlyValue -Value $item -Depth ($Depth + 1))
            }
        }
        return ($items -join '; ')
    }

    $properties = @(
        $Value.PSObject.Properties |
            Where-Object { $_.MemberType -in @('NoteProperty', 'AliasProperty') }
    )

    if ($properties.Count -gt 0) {
        $pairs = @()
        foreach ($property in $properties) {
            $pairs += "$($property.Name)=$(ConvertTo-ExportFriendlyValue -Value $property.Value -Depth ($Depth + 1))"
        }
        return ($pairs -join '; ')
    }

    return $Value.ToString()
}

function ConvertTo-ExportFriendlyRecord {
    param($InputObject)

    if ($null -eq $InputObject) {
        return [pscustomobject]@{ Value = $null }
    }

    if ($InputObject -is [hashtable] -or $InputObject -is [System.Collections.Specialized.OrderedDictionary]) {
        $result = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $result[[string]$key] = ConvertTo-ExportFriendlyValue -Value $InputObject[$key]
        }
        return [pscustomobject]$result
    }

    $properties = @(
        $InputObject.PSObject.Properties |
            Where-Object { $_.MemberType -in @('NoteProperty', 'AliasProperty', 'Property') }
    )

    if ($properties.Count -eq 0) {
        return [pscustomobject]@{ Value = ConvertTo-ExportFriendlyValue -Value $InputObject }
    }

    $result = [ordered]@{}
    foreach ($property in $properties) {
        $result[$property.Name] = ConvertTo-ExportFriendlyValue -Value $property.Value
    }

    return [pscustomobject]$result
}

## Export Hash Table to Excel
function Export-HashTableToExcel {
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True)] 
        [Hashtable]$hashtable,
        [Parameter(Mandatory=$True)] 
        [string]$ExportDetails
        #[Parameter(Mandatory=$false)] [switch]$tenant
    )

    # Ensure ExportDetails has .xlsx in the path name
    if ($ExportDetails -notmatch '\.xlsx$') {
        $ExportDetails += ".xlsx"
    }

    function Test-FileLocked {
        param([string]$Path)
        if (-not (Test-Path $Path)) { return $false }
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $stream.Close()
            return $false
        } catch {
            return $true
        }
    }

    $optionalEmptySheets = @(
        'AuthenticationSSOApplications',
        'SMTPRelaySummary',
        'TeamsVoiceSummary',
        'UnmanagedObjects',
        'OneDriveOwnerMismatches'
    )

    $excludedWorksheets = @(
        'OwnershipGovernanceSummary',
        'TenantInfoSummary',
        'AuthenticationConfigSummary',
        'SpamFilteringSummary',
        'FederationSummary',
        'MfaRegistrationSummary',
        'EmailActivitySummary',
        'PrimaryMailboxStatsCollectionSummary',
        'UnifiedGroupMailboxStatsCollectionSummary',
        'PrimaryMailboxStatsCollectionSu',
        'UnifiedGroupMailboxStatsCollect'
    )

    if (Test-Path -Path $ExportDetails) {
        try {
            $existingSheetNames = @()
            if (Get-Command -Name Get-ExcelSheetInfo -ErrorAction SilentlyContinue) {
                $existingSheetNames = @(Get-ExcelSheetInfo -Path $ExportDetails | Select-Object -ExpandProperty Name)
            }

            $sheetsToRemove = @($excludedWorksheets | Where-Object { $existingSheetNames -contains $_ })
            if ($sheetsToRemove.Count -gt 0) {
                Remove-Worksheet -Path $ExportDetails -WorksheetName $sheetsToRemove
                Write-Log -Type INFO -Message "Removed excluded worksheet(s) from existing workbook: $([string]::Join(', ', $sheetsToRemove))" -ExportFileLocation $ExportDetails
            }
        }
        catch {
            Write-Log -Type WARNING -Message "Unable to remove excluded worksheets from existing workbook before export: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }
    }
    
    # === Sheet ordering ===
    $desiredOrder = @(
        # Assessment Outputs
        "BestPractices", "BestPracticeFindings", "MigrationReadiness", "SecureScoreActions", "UnmanagedObjects", "OneDriveOwnerMismatches",

        # Licensing & Tenant Info
        "LicenseSKUs", "Domains", "AuthenticationMethods", "AuthenticationSSOApplications", "AuthenticationConfig", "Admins",

        # Users
        "Users", "UserFullDetails", "DeviceDetails",

        # Mailboxes
        "AllMailboxes", "PrimaryMailboxStats", "MailboxFullDetails", "ArchiveMailboxes", "ArchiveMailboxStats", "LitigationHoldMailboxes", "InactiveMailboxes", "InactiveMailboxDetails", "EmailActivityTopSenders", "EmailActivityTopReceivers", "NonUserMailboxes", "AllRecipients",

        # Groups
        "AllExchangeGroups", "UnifiedGroups", "EntraIDGroups",

        # Public Folders
        "PublicFolderDetails", "PublicFolderPerms",

        # Mail Flow
        "MailFlowRules", "MailFlowConnectors", "RemoteDomains", "SMTPRelayConfig",

        # Security & Compliance
        "SecuritySecureScore", "ConditionalAccessPolicies", "SMTPRelaySummary", "TeamsVoiceSummary", "SpamFilteringConfig",

        # Cloud Services
        "OneDrive",
        "SharePoint",

        # Hybrid / Infra
        "HybridConfiguration", "TenantInfo"
    )

    $orderedTables = @()
    foreach ($name in $desiredOrder) {
        if (($excludedWorksheets -notcontains $name) -and $hashtable.ContainsKey($name)) {
            $orderedTables += $name
        }
    }
    $orderedTables += ($hashtable.Keys | Where-Object { ($orderedTables -notcontains $_) -and ($excludedWorksheets -notcontains $_) } | Sort-Object)

    foreach ($excludedSheet in $excludedWorksheets) {
        if ($hashtable.ContainsKey($excludedSheet)) {
            Write-Log -Type INFO -Message "Skipping worksheet '$excludedSheet' by export policy." -ExportFileLocation $ExportDetails
        }
    }

    $totalCount = ($orderedTables | Measure-Object).Count
    foreach ($table in $orderedTables) {
        try {
            Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash To Excel"
            Write-Log -Type DEBUG -Message ("Exporting '{0}' Hash Table to '{1}'" -f $table, $ExportDetails) -ExportFileLocation $ExportDetails

            $tableValue = $hashtable[$table]
            $sourceEnumerable = $null
            $sourceCount = 0
            $singleValue = $null
            if ($tableValue -is [hashtable] -or $tableValue -is [System.Collections.Specialized.OrderedDictionary]) {
                $sourceCount = $tableValue.Count
                $sourceEnumerable = $tableValue.Values
            } elseif ($tableValue -is [System.Collections.IEnumerable] -and -not ($tableValue -is [string])) {
                if ($tableValue.PSObject.Properties['Count']) {
                    try { $sourceCount = [int]$tableValue.Count } catch { $sourceCount = -1 }
                } else {
                    $sourceCount = -1
                }
                $sourceEnumerable = $tableValue
            } else {
                if ($null -ne $tableValue) {
                    $singleValue = $tableValue
                    $sourceCount = 1
                }
            }

            if ($sourceCount -lt 0 -and $null -ne $sourceEnumerable) {
                $sourceEnumerable = @($sourceEnumerable)
                $sourceCount = $sourceEnumerable.Count
            }

            if ($sourceCount -gt 0) {
                $attempt = 0
                $maxAttempts = 3
                $saved = $false
                $exportSource = if ($null -ne $singleValue) { @($singleValue) } else { $sourceEnumerable }
                $autoSizeSheet = ($sourceCount -le 5000)
                if (-not $autoSizeSheet) {
                    Write-Log -Type INFO -Message "Skipping AutoSize for worksheet '$table' due to row count ($sourceCount) to reduce export runtime/memory pressure." -ExportFileLocation $ExportDetails
                }
                while (-not $saved -and $attempt -lt $maxAttempts) {
                    $attempt++
                    if (Test-FileLocked -Path $ExportDetails) {
                        Write-Log -Type WARNING -Message "Excel file is locked (attempt $attempt/$maxAttempts). Close the file and retrying in 5 seconds..." -ExportFileLocation $ExportDetails
                        Start-Sleep -Seconds 5
                        continue
                    }
                    try {
                        $excelSplat = @{
                            Path          = $ExportDetails
                            WorksheetName = $table
                            ClearSheet    = $true
                            BoldTopRow    = $true
                        }
                        if ($autoSizeSheet) {
                            $excelSplat.AutoSize = $true
                        }

                        $exportSource |
                            ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                            Export-Excel @excelSplat
                        $saved = $true
                    } catch {
                        if ($attempt -lt $maxAttempts) {
                            Write-Log -Type WARNING -Message "Save failed for '$table' (attempt $attempt/$maxAttempts): $($_.Exception.Message). Retrying in 5 seconds..." -ExportFileLocation $ExportDetails
                            Start-Sleep -Seconds 5
                        } else {
                            throw
                        }
                    }
                }
            } else {
                if ($table -eq 'OneDriveOwnerMismatches') {
                    try {
                        $placeholderRows = @(
                            [PSCustomObject]@{
                                DisplayName          = 'N/A'
                                SiteUrl              = 'N/A'
                                CurrentOwner         = 'N/A'
                                ExpectedDefaultOwner = 'N/A'
                                Notes                = 'No owner mismatches detected in this run.'
                            }
                        )
                        $excelSplat = @{
                            Path          = $ExportDetails
                            WorksheetName = $table
                            ClearSheet    = $true
                            BoldTopRow    = $true
                            AutoSize      = $true
                        }
                        $placeholderRows |
                            ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                            Export-Excel @excelSplat
                        Write-Log -Type INFO -Message "No owner mismatch rows found; exported placeholder row to worksheet '$table'." -ExportFileLocation $ExportDetails
                        continue
                    }
                    catch {
                        Write-Log -Type WARNING -Message "Unable to export placeholder row for '$table': $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                    }
                }

                $logType = if ($optionalEmptySheets -contains $table) { 'INFO' } else { 'WARNING' }
                Write-Log -Type $logType -Message "No data found for $($table) to export to Excel" -ExportFileLocation $ExportDetails

                if (Test-Path -Path $ExportDetails) {
                    try {
                        $existingSheetNames = @()
                        if (Get-Command -Name Get-ExcelSheetInfo -ErrorAction SilentlyContinue) {
                            $existingSheetNames = @(Get-ExcelSheetInfo -Path $ExportDetails | Select-Object -ExpandProperty Name)
                        }
                        if ($existingSheetNames -contains $table) {
                            Remove-Worksheet -Path $ExportDetails -WorksheetName $table
                            Write-Log -Type INFO -Message "Removed stale worksheet '$table' because the current run produced no rows." -ExportFileLocation $ExportDetails
                        }
                    }
                    catch {
                        Write-Log -Type WARNING -Message "Unable to remove stale worksheet '$table': $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                    }
                }
            }
        }
        catch {
            Write-Log -Type Error -Message "An error occurred in Exporting Hash To Excel for $($table) to '$($ExportDetails)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }
    }
    Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash To Excel" -Completed
    Write-Host "The report has been exported to: $($ExportDetails)" -ForegroundColor Green
}

# Export Errors
function Export-ErrorReports {
    [CmdletBinding()]
    param(
        [string]$BaseName = "ErrorReport",
        [string]$ExportFileLocation,
        [Parameter(Mandatory=$true)]
        [array]$ErrorData,
        [string]$ErrorReportFolderName,
        [string]$ErrorReportFolderDirectory,
        [Parameter(Mandatory=$True)]
        [string]$logReportDirectory
    )
    return Export-ArrayaErrorReports @PSBoundParameters
}

# ----------------------------------
# Exchange Specific Functions
# ----------------------------------
# NEW - Convert Names to EmailAddresses loop

function Get-ResolvedEmailAddresses {
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Input Recipient. Accepts Identity, DisplayName, PrimarySMTPAddress, or Name')]
        [array]$recipient,

        [Parameter(Mandatory = $true, HelpMessage = 'Specify Exchange Environment')]
        [ValidateSet('ExchangeOnline', 'EXO', 'On-Premises', 'OnPremises')]
        [string]$ExchangeEnvironment,

        [Parameter(Mandatory = $false, HelpMessage = 'Hash Table of Recipient Mail Objects')]
        [hashtable]$MailObjectHash
    )
    
    # Initialize default return values
    $recipientName = $recipient.ToString()
    $DisplayName = $recipientName
    $PrimarySMTPAddress = $recipientName
    $RecipientTypeDetails = "Unknown"

    Write-Verbose "Processing recipient: $recipientName"

    # Check in hash table first
    if ($MailObjectHash) {
        $matchingRecipient = $null
        $allRecipients = if ($MailObjectHash.ContainsKey('AllRecipients') -and $MailObjectHash['AllRecipients'] -is [System.Collections.IDictionary]) { $MailObjectHash['AllRecipients'] } else { $null }
        $allMailboxes = if ($MailObjectHash.ContainsKey('AllMailboxes') -and $MailObjectHash['AllMailboxes'] -is [System.Collections.IDictionary]) { $MailObjectHash['AllMailboxes'] } else { $null }
        $mailboxesByIdentity = if ($MailObjectHash.ContainsKey('AllMailboxes-MailIdentity') -and $MailObjectHash['AllMailboxes-MailIdentity'] -is [System.Collections.IDictionary]) { $MailObjectHash['AllMailboxes-MailIdentity'] } else { $null }
        $mailboxesByUpn = if ($MailObjectHash.ContainsKey('AllMailboxes-UserPrincipalName') -and $MailObjectHash['AllMailboxes-UserPrincipalName'] -is [System.Collections.IDictionary]) { $MailObjectHash['AllMailboxes-UserPrincipalName'] } else { $null }
        $mailboxesBySmtp = if ($MailObjectHash.ContainsKey('AllMailboxes-PrimarySmtpAddress') -and $MailObjectHash['AllMailboxes-PrimarySmtpAddress'] -is [System.Collections.IDictionary]) { $MailObjectHash['AllMailboxes-PrimarySmtpAddress'] } else { $null }

        if ($allRecipients -and $allRecipients.ContainsKey($recipientName)) {
            $matchingRecipient = $allRecipients[$recipientName]
        } elseif ($allMailboxes -and $allMailboxes.ContainsKey($recipientName)) {
            $matchingRecipient = $allMailboxes[$recipientName]
        } elseif ($mailboxesByIdentity -and $mailboxesByIdentity.ContainsKey($recipientName)) {
            $matchingRecipient = $mailboxesByIdentity[$recipientName]
        } elseif ($mailboxesByUpn -and $mailboxesByUpn.ContainsKey($recipientName)) {
            $matchingRecipient = $mailboxesByUpn[$recipientName]
        } elseif ($mailboxesBySmtp -and $mailboxesBySmtp.ContainsKey($recipientName)) {
            $matchingRecipient = $mailboxesBySmtp[$recipientName]
        }
        Write-Verbose "Matched '$recipientName' in hash table"       
    }
    if (-not $matchingRecipient) {
        # If check in Exchange based on environment
        Write-Verbose "Looking up $recipientName in $ExchangeEnvironment"
        switch ($ExchangeEnvironment) {
            {$_ -in 'On-Premises', 'OnPremises'} { 
                $matchingRecipient = Get-Recipient -Identity $recipientName -ErrorAction SilentlyContinue 
            }
            {$_ -in 'ExchangeOnline', 'EXO'} {
                $matchingRecipient = Get-EXORecipient -Identity $recipientName -ErrorAction SilentlyContinue
            }
            Default { 
                $matchingRecipient = Get-Recipient -Identity $recipientName -ErrorAction SilentlyContinue 
            }
        }
    }

    # If a match was found, assign details; otherwise, retain defaults
    if ($matchingRecipient) {
        $DisplayName = $matchingRecipient.DisplayName
        $PrimarySMTPAddress = $matchingRecipient.PrimarySMTPAddress
        $RecipientTypeDetails = $matchingRecipient.RecipientTypeDetails
        Write-Verbose "[Match-ExchangeRecipientFromHash] Matched User '$($recipient)' to '$($matchingRecipient.DisplayName)'"
    } else {
        Write-Verbose "[Match-ExchangeRecipientFromHash] No Match found for User '$recipient'"
    }

    # Return a structured object with the results
    return [PSCustomObject]@{
        DisplayName = $DisplayName
        PrimarySMTPAddress = $PrimarySMTPAddress
        RecipientTypeDetails = $RecipientTypeDetails
    }
}

#Gather all Exchange mailboxes and mailbox statistics
function Get-AllExchangeMailboxDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    $start = Get-Date
    $mailboxInventoryProgressId = 30
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    Write-Host "Getting all mailboxes and inactive mailboxes with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] START: Getting all mailboxes with $($detailLevel) details" -ExportFileLocation $ExportDetails
    try {
        # Gather Mailboxes - Include InActive Mailboxes
        switch ($detailLevel) {
            "minimum" {
                $Properties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress"
                    "Identity", "Guid", "ExchangeGuid", "ArchiveStatus", "ArchiveState", "ArchiveGuid", "ArchiveName"
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted"
                    "LitigationHoldEnabled", "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias", "EmailAddresses"
                )

                $DesiredProperties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress",
                    "Identity", "Guid", "ExchangeGuid", "ArchiveStatus", "ArchiveState", "ArchiveGuid",
                    @{Name="ArchiveName"; Expression={$_.ArchiveName -join ","}},
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted",
                    "LitigationHoldEnabled", "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias",
                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}
                )

                $exoMailboxes = Invoke-QuietCommand -ScriptBlock {
                    Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -Properties $Properties -IncludeInactiveMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Select-Object $DesiredProperties
                }
            }
            "combined" {
                # Combined mode optimization: request a slimmer property set to reduce EXO payload and local projection overhead.
                $Properties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Office", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress"
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted"
                    "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias", "EmailAddresses"
                    "Identity", "WhenCreated", "Guid", "DeliverToMailboxAndForward", "ForwardingAddress"
                    "ForwardingSmtpAddress", "LitigationHoldEnabled", "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy"
                    "ExchangeGuid", "ArchiveStatus", "ArchiveState", "ArchiveGuid", "ArchiveName", "AutoExpandingArchiveEnabled"
                )

                $DesiredProperties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Office", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress",
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted",
                    "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias",
                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}},
                    "Identity", "WhenCreated", "Guid", "DeliverToMailboxAndForward", "ForwardingAddress", "ForwardingSmtpAddress",
                    "LitigationHoldEnabled", "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy",
                    "ExchangeGuid", "ArchiveStatus", "ArchiveState", "ArchiveGuid",
                    @{Name="ArchiveName"; Expression={$_.ArchiveName -join ","}}, "AutoExpandingArchiveEnabled"
                )

                $exoMailboxes = Invoke-QuietCommand -ScriptBlock {
                    Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -Properties $Properties -IncludeInactiveMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Select-Object $DesiredProperties
                }
            }
            "all" {
                $Properties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Office", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress"
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted"
                    "InPlaceHolds", "AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias"
                    "EmailAddresses", "GrantSendOnBehalfTo", "AcceptMessagesOnlyFrom", "AcceptMessagesOnlyFromDLMembers", "AcceptMessagesOnlyFromSendersOrMembers"
                    "RejectMessagesFrom", "RejectMessagesFromDLMembers", "RejectMessagesFromSendersOrMembers", "RequireSenderAuthenticationEnabled", "WindowsEmailAddress"
                    "DistinguishedName", "Identity", "WhenChanged", "WhenCreated", "ExchangeObjectId"
                    "Guid", "DeliverToMailboxAndForward", "ForwardingAddress", "ForwardingSmtpAddress", "LitigationHoldEnabled"
                    "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy", "ExchangeGuid", "IsResource"
                    "IsShared", "ResourceType", "RoomMailboxAccountEnabled", "WindowsLiveID", "MicrosoftOnlineServicesID"
                    "EffectivePublicFolderMailbox", "MailboxPlan", "ArchiveStatus", "ArchiveState", "ArchiveName"
                    "ArchiveGuid", "AutoExpandingArchiveEnabled", "DisabledArchiveGuid", "PersistedCapabilities"
                )

                $DesiredProperties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Office", "UserPrincipalName", "RecipientTypeDetails", "PrimarySmtpAddress",
                    "WhenMailboxCreated", "UsageLocation", "IsInactiveMailbox", "WasInactiveMailbox", "WhenSoftDeleted",
                    @{Name="InPlaceHolds"; Expression={$_.InPlaceHolds -join ","}},"AccountDisabled", "IsDirSynced", "HiddenFromAddressListsEnabled", "Alias",
                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, 
                    @{Name="GrantSendOnBehalfTo"; Expression={$_.GrantSendOnBehalfTo -join ","}}, 
                    @{Name="AcceptMessagesOnlyFrom"; Expression={$_.AcceptMessagesOnlyFrom -join ","}}, 
                    @{Name="AcceptMessagesOnlyFromDLMembers"; Expression={$_.AcceptMessagesOnlyFromDLMembers -join ","}}, 
                    @{Name="AcceptMessagesOnlyFromSendersOrMembers"; Expression={$_.AcceptMessagesOnlyFromSendersOrMembers -join ","}}, 
                    @{Name="RejectMessagesFrom"; Expression={$_.RejectMessagesFrom -join ","}}, 
                    @{Name="RejectMessagesFromDLMembers"; Expression={$_.RejectMessagesFromDLMembers -join ","}}, 
                    @{Name="RejectMessagesFromSendersOrMembers"; Expression={$_.RejectMessagesFromSendersOrMembers -join ","}}, 
                    "RequireSenderAuthenticationEnabled", "WindowsEmailAddress",
                    "DistinguishedName", "Identity", "WhenChanged", "WhenCreated", "ExchangeObjectId",
                    "Guid", "DeliverToMailboxAndForward", "ForwardingAddress", "ForwardingSmtpAddress", "LitigationHoldEnabled",
                    "RetentionHoldEnabled", "DelayHoldApplied", "RetentionPolicy", "ExchangeGuid", "IsResource",
                    "IsShared", "ResourceType", "RoomMailboxAccountEnabled", "WindowsLiveID", "MicrosoftOnlineServicesID", "EffectivePublicFolderMailbox", "MailboxPlan", 
                    "ArchiveStatus", "ArchiveState", @{Name="ArchiveName"; Expression={$_.ArchiveName -join ","}}, "ArchiveGuid", "AutoExpandingArchiveEnabled", "DisabledArchiveGuid",
                    @{Name="PersistedCapabilities"; Expression={$_.PersistedCapabilities -join ","}}
                )

                $exoMailboxes = Invoke-QuietCommand -ScriptBlock {
                    Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -Properties $Properties -IncludeInactiveMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Select-Object $DesiredProperties
                }
            }
            geek {
                $exoMailboxes = Invoke-QuietCommand -ScriptBlock {
                    Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -IncludeInactiveMailbox -PropertySets All -ResultSize Unlimited -ErrorAction SilentlyContinue
                }
            }
        }
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering all mailboxes (Get-EXOMailbox) including Inactive Mailboxes" -ExportFileLocation $ExportDetails
        
        #Create Hash Table to store mailboxes
        # Ensure global hash table structure
        if (-not $script:tenantStatsHash) {
            $script:tenantStatsHash = @{}
        }
        $script:tenantStatsHash["AllMailboxes"] = @{}
        $script:tenantStatsHash["AllMailboxes-MailIdentity"] = @{}
        $script:tenantStatsHash["AllMailboxes-UserPrincipalName"] = @{}
        $script:tenantStatsHash["AllMailboxes-PrimarySmtpAddress"] = @{}
        $script:tenantStatsHash["NonUserMailboxes"] = @{}
        
        $script:tenantStatsHash["ArchiveMailboxes"] = @{}
        $script:tenantStatsHash["InactiveMailboxes"] = @{}
        $script:tenantStatsHash["LitigationHoldMailboxes"] = @{}

        # Insert individual mailboxes into the hashtable
        $totalCount = $exoMailboxes.Count
        foreach ($mailbox in $exoMailboxes) {
            $key = @(
                [string]$mailbox.ExternalDirectoryObjectId
                if ($mailbox.ExchangeGuid) { [string]$mailbox.ExchangeGuid }
                if ($mailbox.Guid) { [string]$mailbox.Guid }
                [string]$mailbox.UserPrincipalName
                [string]$mailbox.PrimarySmtpAddress
                [string]$mailbox.Identity
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
            if ([string]::IsNullOrWhiteSpace($key)) {
                $key = "mailbox:$([guid]::NewGuid().Guid)"
            }
            Write-ProgressHelper -Total $totalCount -Id $mailboxInventoryProgressId -Activity "Gathering All Exchange Mailbox Details" -Operation "Gathering Mailbox Details for $($key)"
            
            # Set the key based on the mailbox type
            #$MailboxTypeKey = $mailbox.RecipientTypeDetails.tostring() # Use RecipientTypeDetails as the key
            $script:tenantStatsHash["AllMailboxes"][$key] = $mailbox
            $script:tenantStatsHash["AllMailboxes-MailIdentity"][$mailbox.Identity] = $mailbox
            if ($mailbox.UserPrincipalName) {
                $script:tenantStatsHash["AllMailboxes-UserPrincipalName"][[string]$mailbox.UserPrincipalName] = $mailbox
            }
            if ($mailbox.PrimarySmtpAddress) {
                $script:tenantStatsHash["AllMailboxes-PrimarySmtpAddress"][[string]$mailbox.PrimarySmtpAddress] = $mailbox
            }

            if ($mailbox.RecipientTypeDetails -ne "UserMailbox" -and $mailbox.RecipientTypeDetails -ne "GroupMailbox") {
                #$MailboxTypeKey = "NonUserMailboxes"
                $script:tenantStatsHash["NonUserMailboxes"][$key] = $mailbox
            }
            if ($mailbox.ArchiveStatus -eq "Active") {
                #$MailboxTypeKey = "ArchiveMailboxes"
                $script:tenantStatsHash["ArchiveMailboxes"][$key] = $mailbox
            }
            if ($mailbox.IsInactiveMailbox -eq $true) {
                #$MailboxTypeKey = "InactiveMailboxes"
                $script:tenantStatsHash["InactiveMailboxes"][$key] = $mailbox
            }
            if ($mailbox.LitigationHoldEnabled -eq $true) {
                #$MailboxTypeKey = "LitigationHoldMailboxes"
                $script:tenantStatsHash["LitigationHoldMailboxes"][$key] = $mailbox
            }
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Mailbox Details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    } finally {
        Write-ProgressHelper -Total ([Math]::Max($totalCount, 1)) -Id $mailboxInventoryProgressId -Activity "Gathering All Exchange Mailbox Details" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] COMPLETED: Gathering All Mailbox Details in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }

    #Mailbox Statistics to Hash Table
    ###########################################################################################################################################
    $primaryStatsProgressId = 31
    $archiveStatsProgressId = 32

    ## Primary Mailbox Stats
    try {
        $start = Get-Date
        Write-Progress -Id $primaryStatsProgressId -Activity "Gathering All Primary Mailbox Statistics" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        Write-Host "  Getting primary mailbox stats..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering all primary mailbox statistics for collected mailboxes (including inactive where available)." -ExportFileLocation $ExportDetails

        $script:tenantStatsHash["PrimaryMailboxStats"] = @{}
        $script:MailboxUsageGraphLookup = @{}
        $activeMailboxes = $script:tenantStatsHash['AllMailboxes'].Values | Where-Object { $_.IsInactiveMailbox -ne $true }
        $graphLikelyEligibleMailboxCount = @(
            $activeMailboxes | Where-Object {
                $_.PSObject.Properties['RecipientTypeDetails'] -and
                [string]$_.RecipientTypeDetails -eq 'UserMailbox'
            }
        ).Count
        $activeMailboxTypeBreakdown = @(
            $activeMailboxes |
                Group-Object -Property RecipientTypeDetails |
                Sort-Object -Property Count -Descending |
                ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }
        ) -join '; '
        $activeMailboxUnmatchedAfterGraphCount = @($activeMailboxes).Count
        $graphStatsCount = 0
        $graphReportRowCount = 0
        $graphMatchedMailboxCount = 0
        $graphUsablePrincipalRowCount = 0
        $graphRowsWithoutPrincipalCount = 0
        $graphRowsPotentiallyObscuredPrincipalCount = 0
        $graphRowsDuplicatePrincipalCount = 0
        $activeMailboxWithoutLookupKeyCount = 0
        $activeMailboxLookupMissCount = 0
        $activeMailboxMatchedByGraphCount = 0
        $exoFallbackRequestedCount = 0
        $exoFallbackReturnedCount = 0
        $graphPhaseSeconds = 0
        $exoFallbackPhaseSeconds = 0
        $unifiedPreCachePhaseSeconds = 0
        $graphCoverageWarning = $null
        $shouldPreCacheUnifiedGroupStats = Test-ShouldCollectUnifiedGroupMailboxStats -DetailLevel $detailLevel
        $estimatedGroupMailboxCount = @(
            $script:tenantStatsHash['AllMailboxes'].Values | Where-Object {
                $_.PSObject.Properties['RecipientTypeDetails'] -and [string]$_.RecipientTypeDetails -eq 'GroupMailbox'
            }
        ).Count
        if ($shouldPreCacheUnifiedGroupStats) {
            Write-Host ("    Step 3/3 Unified group pre-cache: scheduled | estimated group mailboxes in inventory={0}" -f $estimatedGroupMailboxCount) -ForegroundColor DarkGray
        }

        # Try Graph mailbox usage report (fast) for active mailboxes
        $graphPhaseStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $mailboxUsageUri = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D180')"
            $graphReportData = @(Export-ArrayaGraphReportCsv -Uri $mailboxUsageUri -Activity 'Mailbox usage detail report' -Headers $global:GraphHeaders)
            $graphReportRowCount = @($graphReportData).Count
            if ($graphReportData -and $graphReportData.Count -gt 0) {
                foreach ($item in $graphReportData) {
                    $principalValue = Get-GraphReportFieldValue -Row $item -FieldNames @(
                        'User Principal Name',
                        'User Principal Name ',
                        'User Principal Name (UPN)',
                        'Owner Principal Name',
                        'Account'
                    )
                    if ([string]::IsNullOrWhiteSpace($principalValue)) {
                        $graphRowsWithoutPrincipalCount++
                        continue
                    }

                    $principalKey = $principalValue.Trim().ToLowerInvariant()
                    if (-not ($principalValue -match '@')) {
                        $graphRowsPotentiallyObscuredPrincipalCount++
                    }
                    if ($script:MailboxUsageGraphLookup.ContainsKey($principalKey)) {
                        $graphRowsDuplicatePrincipalCount++
                        continue
                    }

                    $script:MailboxUsageGraphLookup[$principalKey] = $item
                    $graphUsablePrincipalRowCount++
                }

                foreach ($mailbox in $activeMailboxes) {
                    $mailboxLookupKey = $null
                    if (-not [string]::IsNullOrWhiteSpace([string]$mailbox.UserPrincipalName)) {
                        $mailboxLookupKey = ([string]$mailbox.UserPrincipalName).ToLowerInvariant()
                    }
                    elseif (-not [string]::IsNullOrWhiteSpace([string]$mailbox.PrimarySmtpAddress)) {
                        $mailboxLookupKey = ([string]$mailbox.PrimarySmtpAddress).ToLowerInvariant()
                    }

                    if (-not $mailboxLookupKey) {
                        $activeMailboxWithoutLookupKeyCount++
                        continue
                    }
                    if (-not $script:MailboxUsageGraphLookup.ContainsKey($mailboxLookupKey)) {
                        $activeMailboxLookupMissCount++
                        continue
                    }

                    $graphData = $script:MailboxUsageGraphLookup[$mailboxLookupKey]
                    $storageBytes = 0
                    $deletedBytes = 0
                    $itemCount = "0"
                    if ($graphData.'Storage Used (Byte)') { $storageBytes = [double]$graphData.'Storage Used (Byte)' }
                    if ($graphData.'Deleted Item Size (Byte)') { $deletedBytes = [double]$graphData.'Deleted Item Size (Byte)' }
                    if ($graphData.'Item Count') { $itemCount = $graphData.'Item Count' }
                    $guidKey = if ($mailbox.ExchangeGuid) { Convert-ToMailboxGuidKey -GuidValue $mailbox.ExchangeGuid } elseif ($mailbox.Guid) { Convert-ToMailboxGuidKey -GuidValue $mailbox.Guid } else { $null }
                    if (-not $guidKey) { continue }
                    $stats = [PSCustomObject]@{
                        DisplayName = $mailbox.DisplayName
                        TotalItemSize = "$([math]::Round($storageBytes / 1GB, 4)) GB ($storageBytes bytes)"
                        TotalItemSizeBytes = [int64]$storageBytes
                        ItemCount = $itemCount
                        TotalDeletedItemSize = "$([math]::Round($deletedBytes / 1GB, 4)) GB ($deletedBytes bytes)"
                        TotalDeletedItemSizeBytes = [int64]$deletedBytes
                        MailboxType = $mailbox.RecipientTypeDetails
                        MailboxGuid = if ($mailbox.ExchangeGuid) { $mailbox.ExchangeGuid } else { $mailbox.Guid }
                    }
                    $script:tenantStatsHash["PrimaryMailboxStats"][$guidKey] = $stats
                    $activeMailboxMatchedByGraphCount++
                }
            }
            $graphStatsCount = $script:tenantStatsHash["PrimaryMailboxStats"].Count
            $graphMatchedMailboxCount = $graphStatsCount
            $activeMailboxUnmatchedAfterGraphCount = [Math]::Max(($activeMailboxes.Count - $activeMailboxMatchedByGraphCount), 0)
            if (
                $graphReportRowCount -gt 0 -and
                $graphMatchedMailboxCount -eq 0 -and
                $graphRowsPotentiallyObscuredPrincipalCount -gt 0
            ) {
                $graphCoverageWarning = "Graph mailbox report appears to contain obfuscated principals (non-UPN rows=$graphRowsPotentiallyObscuredPrincipalCount), which can force EXO fallback."
            }
        } catch {
            Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] Graph mailbox usage report failed: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }
        finally {
            if ($graphPhaseStopwatch -and $graphPhaseStopwatch.IsRunning) {
                $graphPhaseStopwatch.Stop()
            }
            if ($graphPhaseStopwatch) {
                $graphPhaseSeconds = [math]::Round($graphPhaseStopwatch.Elapsed.TotalSeconds, 2)
            }
        }
        Write-Host ("    Step 1/3 Graph mailbox usage: {0}s | rows={1}, usable={2}, matched={3}/{4}, unresolvedActive={5}" -f $graphPhaseSeconds, $graphReportRowCount, $graphUsablePrincipalRowCount, $graphMatchedMailboxCount, $activeMailboxes.Count, $activeMailboxUnmatchedAfterGraphCount) -ForegroundColor DarkGray
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 1/3 (Graph mailbox usage) completed in $graphPhaseSeconds sec; reportRows=$graphReportRowCount; usablePrincipalRows=$graphUsablePrincipalRowCount; missingPrincipalRows=$graphRowsWithoutPrincipalCount; nonUpnRows=$graphRowsPotentiallyObscuredPrincipalCount; duplicatePrincipalRows=$graphRowsDuplicatePrincipalCount; activeMailboxCount=$($activeMailboxes.Count); activeMatched=$activeMailboxMatchedByGraphCount; activeWithoutLookupKey=$activeMailboxWithoutLookupKeyCount; activeLookupMiss=$activeMailboxLookupMissCount; populated=$graphMatchedMailboxCount; unresolvedActive=$activeMailboxUnmatchedAfterGraphCount." -ExportFileLocation $ExportDetails
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Graph mailbox usage endpoint scope diagnostics: likelyEligibleUserMailboxes=$graphLikelyEligibleMailboxCount; activeMailboxCount=$($activeMailboxes.Count); activeMailboxTypeBreakdown='$activeMailboxTypeBreakdown'. This Graph report is user-mailbox activity data and does not fully represent shared/resource/group/system mailbox inventory." -ExportFileLocation $ExportDetails
        if (-not [string]::IsNullOrWhiteSpace($graphCoverageWarning)) {
            Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] $graphCoverageWarning" -ExportFileLocation $ExportDetails
        }

        Write-Progress -Id $primaryStatsProgressId -Activity "Gathering All Primary Mailbox Statistics" -Completed

        # Fallback to EXO stats for unresolved mailboxes only (Graph-first, profile-aware filtering)
        $mailboxesNeedingStatsList = New-Object System.Collections.Generic.List[object]
        $profileSkippedForFallbackCount = 0
        $deferredUnifiedGroupMailboxCount = 0
        $exoFallbackActiveCandidateCount = 0
        $exoFallbackInactiveCandidateCount = 0
        $minimumExcludedRecipientTypesForExoStats = @(
            'AuditLogMailbox',
            'AuxAuditLogMailbox',
            'ArbitrationMailbox',
            'DiscoveryMailbox',
            'MonitoringMailbox',
            'PublicFolderMailbox',
            'SupervisoryReviewPolicyMailbox'
        )
        $isMinimumDepth = ($script:CollectionDepthPolicy -and $script:CollectionDepthPolicy.IsMinimum)
        foreach ($mailbox in $script:tenantStatsHash['AllMailboxes'].Values) {
            if (Test-MailboxStatCached -MailboxRecord $mailbox -StatsHash $script:tenantStatsHash["PrimaryMailboxStats"]) {
                continue
            }

            if (
                $shouldPreCacheUnifiedGroupStats -and
                $mailbox.PSObject.Properties['RecipientTypeDetails'] -and
                [string]$mailbox.RecipientTypeDetails -eq 'GroupMailbox'
            ) {
                $deferredUnifiedGroupMailboxCount++
                continue
            }

            if (
                $isMinimumDepth -and
                $mailbox.PSObject.Properties['RecipientTypeDetails'] -and
                $minimumExcludedRecipientTypesForExoStats -contains ([string]$mailbox.RecipientTypeDetails)
            ) {
                $profileSkippedForFallbackCount++
                continue
            }

            [void]$mailboxesNeedingStatsList.Add($mailbox)
            if ($mailbox.PSObject.Properties['IsInactiveMailbox'] -and $mailbox.IsInactiveMailbox -eq $true) {
                $exoFallbackInactiveCandidateCount++
            }
            else {
                $exoFallbackActiveCandidateCount++
            }
        }
        $mailboxesNeedingStats = @($mailboxesNeedingStatsList.ToArray())
        $exoFilledCount = 0
        if ($profileSkippedForFallbackCount -gt 0) {
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Minimum mode optimization skipped EXO mailbox stats fallback for $profileSkippedForFallbackCount system mailbox(es)." -ExportFileLocation $ExportDetails
        }
        if ($deferredUnifiedGroupMailboxCount -gt 0) {
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Deferred EXO mailbox stats fallback for $deferredUnifiedGroupMailboxCount unified group mailbox(es) to unified-group pre-cache stage." -ExportFileLocation $ExportDetails
        }
        if ($mailboxesNeedingStats.Count -gt 0) {
            $exoFallbackStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $exoFallbackRequestedCount = $mailboxesNeedingStats.Count
            #$inactiveMBXTest = ($script:tenantStatsHash['AllMailboxes'].Values | Where-Object { $_.IsInactiveMailbox -eq $true }).Count -gt 0
            Write-Host ("    Step 2/3 EXO fallback: starting | unresolved={0} (active={1}, inactive={2}, deferredGroups={3}, profileSkipped={4})" -f $mailboxesNeedingStats.Count, $exoFallbackActiveCandidateCount, $exoFallbackInactiveCandidateCount, $deferredUnifiedGroupMailboxCount, $profileSkippedForFallbackCount) -ForegroundColor DarkGray
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Graph report covered $($script:tenantStatsHash['PrimaryMailboxStats'].Count) mailboxes; fetching EXO stats for $($mailboxesNeedingStats.Count) missing/inactive." -ExportFileLocation $ExportDetails

            $allRemainingMBXStatsResult = Get-ExoMailboxStatisticsSafe -MailboxObjects $mailboxesNeedingStats -ProgressActivity "Gathering All Primary Mailbox Statistics" -ProgressId $primaryStatsProgressId
            $allRemainingMBXStats = @($allRemainingMBXStatsResult.Results)
            $exoFallbackReturnedCount = $allRemainingMBXStats.Count
            $allRemainingMBXStats | ForEach-Object {
                $key = Convert-ToMailboxGuidKey -GuidValue $_.MailboxGuid
                if (-not $key) { continue }
                if (-not $script:tenantStatsHash["PrimaryMailboxStats"].ContainsKey($key)) {
                    $script:tenantStatsHash["PrimaryMailboxStats"][$key] = $_
                    $exoFilledCount++
                }
            }
            Write-ExoStatisticsFailureSummary -OperationName 'Get-AllExchangeMailboxDetails primary mailbox statistics' -Failures $allRemainingMBXStatsResult.Failures
            $exoFallbackStopwatch.Stop()
            $exoFallbackPhaseSeconds = [math]::Round($exoFallbackStopwatch.Elapsed.TotalSeconds, 2)
            Write-Host ("    Step 2/3 EXO fallback: {0}s | requested={1}, returned={2}, populated={3}" -f $exoFallbackPhaseSeconds, $exoFallbackRequestedCount, $exoFallbackReturnedCount, $exoFilledCount) -ForegroundColor DarkGray
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 2/3 (EXO fallback) completed in $exoFallbackPhaseSeconds sec; requested=$exoFallbackRequestedCount; returned=$exoFallbackReturnedCount; populated=$exoFilledCount." -ExportFileLocation $ExportDetails
        }
        else {
            Write-Host ("    Step 2/3 EXO fallback: skipped | no unresolved mailboxes (deferredGroups={0}, profileSkipped={1})" -f $deferredUnifiedGroupMailboxCount, $profileSkippedForFallbackCount) -ForegroundColor DarkGray
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 2/3 (EXO fallback) skipped; no unresolved mailboxes. deferredUnifiedGroups=$deferredUnifiedGroupMailboxCount; profileSkipped=$profileSkippedForFallbackCount." -ExportFileLocation $ExportDetails
        }

        $finalStatsCount = $script:tenantStatsHash["PrimaryMailboxStats"].Count
        if ($exoFilledCount -eq 0 -and $finalStatsCount -gt $graphStatsCount) {
            $exoFilledCount = $finalStatsCount - $graphStatsCount
        }
        Write-Verbose "exoStatsFilledCount: $exoFilledCount"
        Write-Verbose "graphStatsCount: $graphStatsCount"
        Write-Verbose "finalStatsCount: $finalStatsCount"

        if ($exoFilledCount -lt 0) { $exoFilledCount = 0 }
        $totalMailboxes = $script:tenantStatsHash['AllMailboxes'].Values.Count
        $missingMailboxStatsCount = [Math]::Max(($totalMailboxes - $finalStatsCount), 0)
        $script:tenantStatsHash["PrimaryMailboxStatsCollectionSummary"] = [PSCustomObject]@{
            GraphReportRows        = [int]$graphReportRowCount
            GraphUsablePrincipalRows = [int]$graphUsablePrincipalRowCount
            GraphRowsWithoutPrincipal = [int]$graphRowsWithoutPrincipalCount
            GraphRowsPotentiallyObscuredPrincipal = [int]$graphRowsPotentiallyObscuredPrincipalCount
            GraphRowsDuplicatePrincipal = [int]$graphRowsDuplicatePrincipalCount
            GraphPopulated         = [int]$graphMatchedMailboxCount
            ActiveMailboxes        = [int]$activeMailboxes.Count
            ActiveMailboxMatchedByGraph = [int]$activeMailboxMatchedByGraphCount
            ActiveMailboxesWithoutLookupKey = [int]$activeMailboxWithoutLookupKeyCount
            ActiveMailboxLookupMiss = [int]$activeMailboxLookupMissCount
            ActiveMailboxUnmatchedAfterGraph = [int]$activeMailboxUnmatchedAfterGraphCount
            ExoFallbackRequested   = [int]$exoFallbackRequestedCount
            ExoFallbackRequestedActive = [int]$exoFallbackActiveCandidateCount
            ExoFallbackRequestedInactive = [int]$exoFallbackInactiveCandidateCount
            ExoFallbackReturned    = [int]$exoFallbackReturnedCount
            ExoFallbackPopulated   = [int]$exoFilledCount
            DeferredUnifiedGroupFallback = [int]$deferredUnifiedGroupMailboxCount
            ProfileSkippedFallback = [int]$profileSkippedForFallbackCount
            GraphPhaseSeconds      = [double]$graphPhaseSeconds
            ExoFallbackPhaseSeconds = [double]$exoFallbackPhaseSeconds
            Populated              = [int]$finalStatsCount
            Missing                = [int]$missingMailboxStatsCount
            TotalMailboxes         = [int]$totalMailboxes
            GraphCoverageWarning   = $graphCoverageWarning
        }
        Write-Host ("  Primary mailbox stats source breakdown: Graph={0}, EXO fallback={1}, Missing={2}, ActiveUnmatchedAfterGraph={3}" -f $graphMatchedMailboxCount, $exoFilledCount, $missingMailboxStatsCount, $activeMailboxUnmatchedAfterGraphCount) -ForegroundColor DarkGray
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Mailbox stats source breakdown: GraphReportRows=$graphReportRowCount; GraphUsablePrincipalRows=$graphUsablePrincipalRowCount; GraphRowsWithoutPrincipal=$graphRowsWithoutPrincipalCount; GraphRowsPotentiallyObscuredPrincipal=$graphRowsPotentiallyObscuredPrincipalCount; GraphRowsDuplicatePrincipal=$graphRowsDuplicatePrincipalCount; GraphPopulated=$graphMatchedMailboxCount; ActiveMailboxCount=$($activeMailboxes.Count); ActiveMailboxMatchedByGraph=$activeMailboxMatchedByGraphCount; ActiveMailboxWithoutLookupKey=$activeMailboxWithoutLookupKeyCount; ActiveMailboxLookupMiss=$activeMailboxLookupMissCount; ActiveMailboxUnmatchedAfterGraph=$activeMailboxUnmatchedAfterGraphCount; EXOFallbackRequested=$exoFallbackRequestedCount; EXOFallbackRequestedActive=$exoFallbackActiveCandidateCount; EXOFallbackRequestedInactive=$exoFallbackInactiveCandidateCount; EXOFallbackReturned=$exoFallbackReturnedCount; EXOFallbackPopulated=$exoFilledCount; DeferredUnifiedGroupFallback=$deferredUnifiedGroupMailboxCount; ProfileSkippedFallback=$profileSkippedForFallbackCount; Populated=$finalStatsCount; Missing=$missingMailboxStatsCount; TotalMailboxes=$totalMailboxes." -ExportFileLocation $ExportDetails

        # Pre-cache unified group mailbox stats so later unified-group collection can reuse this data.
        if ($shouldPreCacheUnifiedGroupStats) {
            $unifiedPreCacheStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                Write-Host "    Step 3/3 Unified group pre-cache: starting..." -ForegroundColor DarkGray
                Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Pre-caching unified group mailbox stats into PrimaryMailboxStats." -ExportFileLocation $ExportDetails
                $groupActivityLookup = Get-Office365GroupsActivityMailboxLookup
                $groupActivityReportRows = if ($groupActivityLookup -and $groupActivityLookup.PSObject.Properties['Rows']) { [int]$groupActivityLookup.Rows } else { 0 }
                $unifiedGroupsForStats = @()
                $groupMailboxCandidateSource = 'AllMailboxes(GroupMailbox)'
                $groupMailboxCandidates = @($script:tenantStatsHash['AllMailboxes'].Values | Where-Object {
                    $_.PSObject.Properties['RecipientTypeDetails'] -and [string]$_.RecipientTypeDetails -eq 'GroupMailbox'
                })
                if ($groupMailboxCandidates.Count -eq 0) {
                    $groupMailboxCandidateSource = 'UnifiedGroupsInventory'
                    $unifiedGroupsForStats = if ($script:UnifiedGroupsInventoryCache) { @($script:UnifiedGroupsInventoryCache) } else { @() }
                    if ($unifiedGroupsForStats.Count -eq 0) {
                        switch ($detailLevel) {
                            {$_ -in "minimum", "combined", "all"} {
                                $desiredUnifiedGroupProperties = @(
                                    "PrimarySmtpAddress", "DisplayName", "AccessType", "RecipientTypeDetails",
                                    "ExternalDirectoryObjectId",
                                    "ExchangeGuid", @{Name="ManagedByDetails"; Expression={$_.ManagedByDetails -join ','}}, "Notes",
                                    "SharePointSiteUrl", "ContentMailboxName", "GroupMemberCount",
                                    "AllowAddGuests", "WhenSoftDeleted", "HiddenFromExchangeClientsEnabled",
                                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, @{Name="ModeratedBy"; Expression={$_.ModeratedBy -join ','}}, "FolderPath",
                                    @{Name="Description"; Expression={$_.Description -join ','}}, "WhenCreated"
                                )
                                $unifiedGroupsForStats = @(
                                    Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue } |
                                        Select-Object $desiredUnifiedGroupProperties
                                )
                            }
                            default {
                                $unifiedGroupsForStats = @(
                                    Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue }
                                )
                            }
                        }
                    }
                    $groupMailboxCandidates = @($unifiedGroupsForStats)
                }
                $groupMailboxCandidateCount = $groupMailboxCandidates.Count
                $groupsMissingStats = New-Object System.Collections.Generic.List[object]
                $cachedUnifiedGroupStats = 0
                $graphActivityUnifiedGroupStats = 0
                $groupsWithoutGuidForPreCacheCount = 0

                foreach ($groupMailbox in $groupMailboxCandidates) {
                    $groupGuidKey = $null
                    if ($groupMailbox.PSObject.Properties['ExchangeGuid'] -and $groupMailbox.ExchangeGuid) {
                        $groupGuidKey = Convert-ToMailboxGuidKey -GuidValue $groupMailbox.ExchangeGuid
                    }
                    elseif ($groupMailbox.PSObject.Properties['Guid'] -and $groupMailbox.Guid) {
                        $groupGuidKey = Convert-ToMailboxGuidKey -GuidValue $groupMailbox.Guid
                    }

                    if ([string]::IsNullOrWhiteSpace($groupGuidKey)) {
                        $groupsWithoutGuidForPreCacheCount++
                        continue
                    }

                    if ($script:tenantStatsHash["PrimaryMailboxStats"].ContainsKey($groupGuidKey)) {
                        $cachedUnifiedGroupStats++
                        continue
                    }

                    $groupActivityRow = Get-GroupMailboxActivityRowForUnifiedGroup -GroupRecord $groupMailbox -ActivityLookup $groupActivityLookup
                    if ($groupActivityRow) {
                        $groupActivityStat = New-GroupMailboxStatFromActivityRow -GroupRecord $groupMailbox -ActivityRow $groupActivityRow
                        if ($groupActivityStat) {
                            $script:tenantStatsHash["PrimaryMailboxStats"][$groupGuidKey] = $groupActivityStat
                            $graphActivityUnifiedGroupStats++
                            continue
                        }
                    }

                    [void]$groupsMissingStats.Add($groupMailbox)
                }

                $fetchedUnifiedGroupStats = 0
                if ($groupsMissingStats.Count -gt 0) {
                    $unifiedGroupStatsResult = Get-ExoMailboxStatisticsSafe -MailboxObjects $groupsMissingStats.ToArray() -ProgressActivity "Pre-caching Unified Group Mailbox Statistics" -ProgressId $primaryStatsProgressId
                    foreach ($groupStat in $unifiedGroupStatsResult.Results) {
                        $key = Convert-ToMailboxGuidKey -GuidValue $groupStat.MailboxGuid
                        if (-not $key) { continue }
                        $script:tenantStatsHash["PrimaryMailboxStats"][$key] = $groupStat
                    }
                    $fetchedUnifiedGroupStats = @($unifiedGroupStatsResult.Results).Count
                    Write-ExoStatisticsFailureSummary -OperationName 'Get-AllExchangeMailboxDetails unified group mailbox statistics pre-cache' -Failures $unifiedGroupStatsResult.Failures
                }

                # Pre-cache unified group inventory for later collectors.
                if (-not $unifiedGroupsForStats) {
                    $unifiedGroupsForStats = @()
                }
                if ($unifiedGroupsForStats.Count -eq 0) {
                    $unifiedGroupsForStats = if ($script:UnifiedGroupsInventoryCache) { @($script:UnifiedGroupsInventoryCache) } else { @() }
                }
                if ($unifiedGroupsForStats.Count -eq 0) {
                    switch ($detailLevel) {
                        {$_ -in "minimum", "combined", "all"} {
                            $desiredUnifiedGroupProperties = @(
                                "PrimarySmtpAddress", "DisplayName", "AccessType", "RecipientTypeDetails",
                                "ExternalDirectoryObjectId",
                                "ExchangeGuid", @{Name="ManagedByDetails"; Expression={$_.ManagedByDetails -join ','}}, "Notes",
                                "SharePointSiteUrl", "ContentMailboxName", "GroupMemberCount",
                                "AllowAddGuests", "WhenSoftDeleted", "HiddenFromExchangeClientsEnabled",
                                @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, @{Name="ModeratedBy"; Expression={$_.ModeratedBy -join ','}}, "FolderPath",
                                @{Name="Description"; Expression={$_.Description -join ','}}, "WhenCreated"
                            )
                            $unifiedGroupsForStats = @(
                                Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue } |
                                    Select-Object $desiredUnifiedGroupProperties
                            )
                        }
                        default {
                            $unifiedGroupsForStats = @(
                                Invoke-QuietCommand -ScriptBlock { Get-UnifiedGroup -ResultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue }
                            )
                        }
                    }
                }
                $script:UnifiedGroupsInventoryCache = @($unifiedGroupsForStats)
                $unifiedGroupInventoryCount = $unifiedGroupsForStats.Count

                Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Unified group stats pre-cache summary: source=$groupMailboxCandidateSource; groupMailboxCandidates=$groupMailboxCandidateCount; reused=$cachedUnifiedGroupStats; graphActivityPopulated=$graphActivityUnifiedGroupStats; fetched=$fetchedUnifiedGroupStats; unresolved=$($groupsMissingStats.Count); withoutGuid=$groupsWithoutGuidForPreCacheCount; groupActivityReportRows=$groupActivityReportRows; unifiedGroupInventory=$unifiedGroupInventoryCount." -ExportFileLocation $ExportDetails
                if ($groupMailboxCandidateCount -gt 0) {
                    if ($unifiedPreCacheStopwatch -and $unifiedPreCacheStopwatch.IsRunning) {
                        $unifiedPreCacheStopwatch.Stop()
                    }
                    if ($unifiedPreCacheStopwatch) {
                        $unifiedPreCachePhaseSeconds = [math]::Round($unifiedPreCacheStopwatch.Elapsed.TotalSeconds, 2)
                    }
                    Write-Host ("    Step 3/3 Unified group pre-cache: {0}s | source={1}, candidates={2}, reused={3}, graph={4}, fetched={5}, unresolved={6}" -f $unifiedPreCachePhaseSeconds, $groupMailboxCandidateSource, $groupMailboxCandidateCount, $cachedUnifiedGroupStats, $graphActivityUnifiedGroupStats, $fetchedUnifiedGroupStats, $groupsMissingStats.Count) -ForegroundColor DarkGray
                    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 3/3 (unified group pre-cache) completed in $unifiedPreCachePhaseSeconds sec; source=$groupMailboxCandidateSource; groupMailboxCandidates=$groupMailboxCandidateCount; reused=$cachedUnifiedGroupStats; graphActivityPopulated=$graphActivityUnifiedGroupStats; fetched=$fetchedUnifiedGroupStats; unresolved=$($groupsMissingStats.Count); withoutGuid=$groupsWithoutGuidForPreCacheCount; groupActivityReportRows=$groupActivityReportRows; unifiedGroupInventory=$unifiedGroupInventoryCount." -ExportFileLocation $ExportDetails
                }
                else {
                    if ($unifiedPreCacheStopwatch -and $unifiedPreCacheStopwatch.IsRunning) {
                        $unifiedPreCacheStopwatch.Stop()
                    }
                    if ($unifiedPreCacheStopwatch) {
                        $unifiedPreCachePhaseSeconds = [math]::Round($unifiedPreCacheStopwatch.Elapsed.TotalSeconds, 2)
                    }
                    Write-Host ("    Step 3/3 Unified group pre-cache: {0}s | no group mailboxes found" -f $unifiedPreCachePhaseSeconds) -ForegroundColor DarkGray
                    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Primary mailbox stats step 3/3 (unified group pre-cache) completed in $unifiedPreCachePhaseSeconds sec; no group mailboxes found." -ExportFileLocation $ExportDetails
                }
            }
            catch {
                if ($unifiedPreCacheStopwatch -and $unifiedPreCacheStopwatch.IsRunning) {
                    $unifiedPreCacheStopwatch.Stop()
                }
                if ($unifiedPreCacheStopwatch) {
                    $unifiedPreCachePhaseSeconds = [math]::Round($unifiedPreCacheStopwatch.Elapsed.TotalSeconds, 2)
                }
                Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] Unified group mailbox statistics pre-cache failed: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                Write-Host ("    Step 3/3 Unified group pre-cache: failed after {0}s" -f $unifiedPreCachePhaseSeconds) -ForegroundColor Yellow
            }
            finally {
                Write-Progress -Id $primaryStatsProgressId -Activity "Pre-caching Unified Group Mailbox Statistics" -Completed
            }
        }
        else {
            Write-Host "    Step 3/3 Unified group pre-cache: skipped by profile/depth" -ForegroundColor DarkGray
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Unified group mailbox statistics pre-cache is disabled for this profile/depth." -ExportFileLocation $ExportDetails
        }
        if ($script:tenantStatsHash.ContainsKey("PrimaryMailboxStatsCollectionSummary") -and $script:tenantStatsHash["PrimaryMailboxStatsCollectionSummary"]) {
            $script:tenantStatsHash["PrimaryMailboxStatsCollectionSummary"] | Add-Member -NotePropertyName 'UnifiedPreCachePhaseSeconds' -NotePropertyValue ([double]$unifiedPreCachePhaseSeconds) -Force
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Mailbox Satistics and adding to Hash Table. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-Progress -Id $primaryStatsProgressId -Activity "Gathering All Primary Mailbox Statistics" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] COMPLETED: Gathering All Primary Mailbox Statistics in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }
    
    ## Archive Mailbox Stats to Hash Table
    if ($script:tenantStatsHash['AllMailboxes'].Values | Where-Object {$_.ArchiveStatus -ne "None"}) {
        try {
            $start = Get-Date
            Write-Host "  Getting archive mailbox stats..." -ForegroundColor Cyan -nonewline
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Archive mailbox size/item metrics are not exposed in Graph mailbox usage reports; using EXO statistics for archive mailboxes." -ExportFileLocation $ExportDetails
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering All Archive Mailbox Statistics. Including Group and Inactive Mailboxes" -ExportFileLocation $ExportDetails
            Write-Progress -Id $archiveStatsProgressId -Activity "Gathering All Archive Mailbox Statistics" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
            $archiveMailboxCandidates = @($script:tenantStatsHash['AllMailboxes'].Values | Where-Object {$_.ArchiveStatus -ne "None"})
            $archiveMailboxStatsResult = Get-ExoMailboxStatisticsSafe -MailboxObjects $archiveMailboxCandidates -Archive -ProgressActivity "Gathering All Archive Mailbox Statistics" -ProgressId $archiveStatsProgressId
            $archiveMailboxStats = @($archiveMailboxStatsResult.Results)
            if ($archiveMailboxStats) {
                $script:tenantStatsHash["ArchiveMailboxStats"] = @{}
                
                #Add to Tenant Stats Hash
                $archiveMailboxStats | ForEach-Object {
                    # Using MailboxGUID from the Mailbox Statistics as the key; matches against the ExchangeGUID from the Mailbox
                    $key = Convert-ToMailboxGuidKey -GuidValue $_.MailboxGuid
                    if (-not $key) { continue }
                    $value = $_
                    $script:tenantStatsHash["ArchiveMailboxStats"][$key] = $value
                }
            }
            Write-ExoStatisticsFailureSummary -OperationName 'Get-AllExchangeMailboxDetails archive mailbox statistics' -Failures $archiveMailboxStatsResult.Failures
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Adding Archive Mailbox Statistics to Tenant Stats Hash." -ExportFileLocation $ExportDetails
        }
        catch {
            if ($_.Exception.Message -like "You cannot call a method on a null-valued expression") {
                Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Archive Mailbox Satistics and adding to Hash Table. There are no Archive mailboxes found. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
            else {
                Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Archive Mailbox Satistics and adding to Hash Table. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
        }
        finally { 
            Write-Progress -Id $archiveStatsProgressId -Activity "Gathering All Archive Mailbox Statistics" -Completed
            $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
            Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails]COMPLETED: Gathering All Archive Mailbox Statistics in $($CompletedTime)" -ExportFileLocation $ExportDetails
        }
    } else {
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] No Archive Mailboxes found." -ExportFileLocation $ExportDetails
        Write-Host "  No Archive Mailboxes found." -ForegroundColor Yellow
    }
}

function Get-AllRecipientDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    $recipientProgressId = 33
    try {
        $start = Get-Date
        # Ensure global hash table structure
        if (-not $script:tenantStatsHash) {
            $script:tenantStatsHash = @{}
        }
        $script:tenantStatsHash["AllRecipients"] = @{}
        Write-Host "Getting all Exchange Online Recipients $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] START: Gathering all Exchange Online Recipients $($detailLevel) details" -ExportFileLocation $ExportDetails
        Write-Progress -Id $recipientProgressId -Activity "Gathering All Exchange Online Recipients" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))

        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $Properties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Identity", "RecipientTypeDetails", "PrimarySMTPAddress"
                    "EmailAddresses", "HiddenFromAddressListsEnabled", "AddressBookPolicy"
                    "ManagedBy", "SKUAssigned", "WhenCreated", "WhenSoftDeleted", "GUID"
                    "alias", "Notes"
                )
                # Properties for Select-Object
                $DesiredProperties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Identity", "RecipientTypeDetails", "PrimarySMTPAddress",
                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, 
                    "HiddenFromAddressListsEnabled", "AddressBookPolicy",
                    @{Name="ManagedBy"; Expression={$_.ManagedBy -join ","}}, 
                    "SKUAssigned", "WhenCreated", "WhenSoftDeleted", "GUID",
                    "alias", "Notes"
                )
                $allRecipients = Invoke-QuietCommand -ScriptBlock {
                    Get-EXORecipient -Properties $Properties -ResultSize Unlimited -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox' -and RecipientTypeDetails -ne 'MailContact' -and RecipientTypeDetails -ne 'GuestMailUser' -and RecipientTypeDetails -ne 'MailUser'" -ErrorAction Stop | select $DesiredProperties
                }
            }           
            geek {
                $allRecipients = Invoke-QuietCommand -ScriptBlock {
                    Get-EXORecipient -PropertySets All -ResultSize Unlimited -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox' -and RecipientTypeDetails -ne 'MailContact' -and RecipientTypeDetails -ne 'GuestMailUser' -and RecipientTypeDetails -ne 'MailUser'" -ErrorAction Stop
                }
            }
        }
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] FOUND $($allRecipients.count) Exchange Online Recipients $($detailLevel) details" -ExportFileLocation $ExportDetails
        #Add to hash table
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] Adding Exchange Online Recipients to Tenant Stats Hash" -ExportFileLocation $ExportDetails
        foreach ($recipient in $allRecipients) {
            $script:tenantStatsHash["AllRecipients"][$recipient.PrimarySmtpAddress] = $recipient
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllRecipientDetails] An error occurred in running Get-AllRecipientDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-Progress -Id $recipientProgressId -Activity "Gathering All Exchange Online Recipients" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] COMPLETED: Gathering all Exchange Online Recipients" -ExportFileLocation $ExportDetails
    }
}

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

    try {
        $periodDuration = if ($detailLevel -eq 'minimum') { 'D90' } else { 'D180' }
        $topLimit = if ($detailLevel -eq 'minimum') { 10 } else { 25 }
        $emailActivityRows = @()
        $emailActivitySource = 'Export-ArrayaGraphReportCsv'

        $graphActivityCommand = Get-Command -Name 'Office365Custom\Get-GraphAPIActivityReport' -ErrorAction SilentlyContinue
        if ($graphActivityCommand) {
            try {
                $savedProgressPreference = $ProgressPreference
                try {
                    $ProgressPreference = 'SilentlyContinue'
                    $emailActivityRows = @(
                        Office365Custom\Get-GraphAPIActivityReport -ServiceName EmailActivity -PeriodDuration $periodDuration -ErrorAction Stop
                    )
                }
                finally {
                    $ProgressPreference = $savedProgressPreference
                }

                if ($emailActivityRows.Count -gt 0) {
                    $emailActivitySource = 'Office365Custom.Get-GraphAPIActivityReport'
                }
            }
            catch {
                Write-Log -Type WARNING -Message "[Get-EmailActivityInsights] Office365Custom\\Get-GraphAPIActivityReport failed: $($_.Exception.Message). Falling back to direct Graph report URI." -ExportFileLocation $ExportDetails
                $emailActivityRows = @()
            }
        }

        if ($emailActivityRows.Count -eq 0) {
            $emailActivityUri = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='$periodDuration')"
            $emailActivityRows = @(
                Export-ArrayaGraphReportCsv -Uri $emailActivityUri -Activity "Email activity user detail report ($periodDuration)" -Headers $global:GraphHeaders
            )
            $emailActivitySource = 'Export-ArrayaGraphReportCsv'
        }

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
            }
            Write-Log -Type INFO -Message "[Get-EmailActivityInsights] No rows returned for email activity report." -ExportFileLocation $ExportDetails
            return
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
        }

        Write-Host ("Top senders={0}, top receivers={1}" -f $topSenders.Count, $topReceivers.Count) -ForegroundColor DarkGray -NoNewline
        Write-Log -Type INFO -Message "[Get-EmailActivityInsights] Email activity summary: source=$emailActivitySource; period=$periodDuration; reportRows=$($emailActivityRows.Count); usersNormalized=$($normalizedUsers.Count); activeUsers=$activeUsers; topSenders=$($topSenders.Count); topReceivers=$($topReceivers.Count)." -ExportFileLocation $ExportDetails
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

# Exchange Group Details
function Get-ExchangeGroupDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    $exchangeGroupProgressId = 36
    $exchangeGroupProgressTotal = 1
    $depthPolicy = if ($script:CollectionDepthPolicy) {
        $script:CollectionDepthPolicy
    } else {
        Get-ArrayaCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($detailLevel.ToLowerInvariant()))
    }
    $isCombinedDepth = [bool]($depthPolicy.PSObject.Properties['IsCombined'] -and $depthPolicy.IsCombined)
    $collectExchangeGroupMetadataDetails = ($depthPolicy.IsAll -or $depthPolicy.IsGeek)
    $collectExchangeGroupMemberExpansion = ($depthPolicy.IsAll -or $depthPolicy.IsGeek)
    try {
        $start = Get-Date
        # Ensure global hash table structure
        if (-not $script:tenantStatsHash) {
            $script:tenantStatsHash = @{}
        }
        $script:tenantStatsHash["AllExchangeGroups"] = @{}
        $cachedUnifiedGroupsByAddress = @{}
        if (
            $script:tenantStatsHash.ContainsKey('UnifiedGroups') -and
            $script:tenantStatsHash['UnifiedGroups'] -is [System.Collections.IDictionary]
        ) {
            foreach ($entry in $script:tenantStatsHash['UnifiedGroups'].GetEnumerator()) {
                $smtp = [string]$entry.Key
                if ([string]::IsNullOrWhiteSpace($smtp)) {
                    continue
                }
                $cachedUnifiedGroupsByAddress[$smtp.ToLowerInvariant()] = $entry.Value
            }
        }
        if ($script:UnifiedGroupsInventoryCache) {
            foreach ($group in @($script:UnifiedGroupsInventoryCache)) {
                if (-not $group -or -not $group.PSObject.Properties['PrimarySmtpAddress']) {
                    continue
                }
                $smtp = [string]$group.PrimarySmtpAddress
                if ([string]::IsNullOrWhiteSpace($smtp)) {
                    continue
                }
                $lookupKey = $smtp.ToLowerInvariant()
                if (-not $cachedUnifiedGroupsByAddress.ContainsKey($lookupKey)) {
                    $cachedUnifiedGroupsByAddress[$lookupKey] = $group
                }
            }
        }
        $metadataReusedFromRecipientCount = 0
        $metadataReusedFromUnifiedCacheCount = 0
        $metadataFetchedFromExoCount = 0
        
        Write-Host "Getting all Exchange Online Groups ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] START: Gathering all Exchange Online Groups with $($detailLevel) details" -ExportFileLocation $ExportDetails
        if (-not $collectExchangeGroupMetadataDetails) {
            if ($depthPolicy.IsMinimum) {
                Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Minimum mode optimization active. Reusing recipient inventory and skipping EXO group metadata/member expansion." -ExportFileLocation $ExportDetails
            }
            elseif ($isCombinedDepth) {
                Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Combined mode optimization active. Reusing recipient inventory and unified-group cache for metadata; skipping per-group EXO metadata/member expansion." -ExportFileLocation $ExportDetails
            }
        }
        elseif (-not $collectExchangeGroupMemberExpansion) {
            Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Combined mode optimization active. Gathering group metadata while skipping deep member expansion." -ExportFileLocation $ExportDetails
        }

        # gather All Exchange Online Groups
        $allMailGroups = $script:tenantStatsHash['AllRecipients'].Values | Where-Object { $_.RecipientTypeDetails -like "*group" } | Sort-Object DisplayName
        
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Gathering all Exchange Online Groups Details" -ExportFileLocation $ExportDetails
        $totalCount = $allMailGroups.count
        $exchangeGroupProgressTotal = [Math]::Max($totalCount, 1)
        foreach ($object in $allMailGroups) {
            try {
                $identity = $object.identity.tostring()
                $PrimarySMTPAddress = if ($object.PrimarySMTPAddress) { $object.PrimarySMTPAddress.ToString() } else { $identity }
                $primarySmtpLookupKey = if ([string]::IsNullOrWhiteSpace($PrimarySMTPAddress)) { $null } else { $PrimarySMTPAddress.ToLowerInvariant() }
                Write-ProgressHelper -Total $exchangeGroupProgressTotal -Id $exchangeGroupProgressId -Activity "Gathering All Exchange Online Group Details" -Operation "Gathering Group Details for $($PrimarySMTPAddress)"
                Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] Gathering '{0}' '{1}' Group Details" -f $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails
    
                # Clear details
                $attributesToClear = @('groupDetails', 'groupOwners','groupMembers')
                foreach ($attribute in $attributesToClear) {
                    Set-Variable -Name $attribute -Value @()
                }
                $groupMembersCount = 0
                $cachedUnifiedGroup = $null
                $usedUnifiedGroupCache = $false

                if (
                    $primarySmtpLookupKey -and
                    $cachedUnifiedGroupsByAddress.ContainsKey($primarySmtpLookupKey)
                ) {
                    $cachedUnifiedGroup = $cachedUnifiedGroupsByAddress[$primarySmtpLookupKey]
                }
                
                # Conditional logic for different recipient types
                switch ($object.RecipientTypeDetails) {
                    "DynamicDistributionGroup" {
                        if ($collectExchangeGroupMetadataDetails) {
                            $groupDetails = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails details for $PrimarySMTPAddress" -DefaultValue $null -ExportFileLocation $ExportDetails -ScriptBlock {
                                Get-DynamicDistributionGroup $identity -ErrorAction Stop
                            }
                            if ($groupDetails) {
                                $metadataFetchedFromExoCount++
                            }
                        }
                        else {
                            $groupDetails = $object
                            $metadataReusedFromRecipientCount++
                        }

                        if ($collectExchangeGroupMemberExpansion) {
                            $groupMembers = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails members for $PrimarySMTPAddress" -DefaultValue @() -ExportFileLocation $ExportDetails -ScriptBlock {
                                @(Get-DynamicDistributionGroupMember $identity -ErrorAction Stop -ResultSize unlimited -WarningAction SilentlyContinue)
                            }
                        }
                        else {
                            $groupMembers = @()
                        }
                    }
                    {$_ -in 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', "MailNonUniversalGroup"} {
                        if ($collectExchangeGroupMetadataDetails) {
                            $groupDetails = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails details for $PrimarySMTPAddress" -DefaultValue $null -ExportFileLocation $ExportDetails -ScriptBlock {
                                Get-DistributionGroup $identity -ErrorAction Stop
                            }
                            if ($groupDetails) {
                                $metadataFetchedFromExoCount++
                            }
                        }
                        else {
                            $groupDetails = $object
                            $metadataReusedFromRecipientCount++
                        }

                        if ($collectExchangeGroupMemberExpansion) {
                            $groupMembers = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails members for $PrimarySMTPAddress" -DefaultValue @() -ExportFileLocation $ExportDetails -ScriptBlock {
                                @(Get-DistributionGroupMember $identity -ResultSize unlimited -ErrorAction Stop)
                            }
                        }
                        else {
                            $groupMembers = @()
                        }
                    }
                    "GroupMailbox" {
                        if ($cachedUnifiedGroup) {
                            $groupDetails = $cachedUnifiedGroup
                            $usedUnifiedGroupCache = $true
                            $metadataReusedFromUnifiedCacheCount++
                            if ($cachedUnifiedGroup.PSObject.Properties['GroupMemberCount'] -and $cachedUnifiedGroup.GroupMemberCount -ne $null -and $cachedUnifiedGroup.GroupMemberCount -ne '') {
                                try { $groupMembersCount = [int]$cachedUnifiedGroup.GroupMemberCount } catch { $groupMembersCount = 0 }
                            }
                            $groupMembers = @()
                        }

                        if ($collectExchangeGroupMetadataDetails -and -not $usedUnifiedGroupCache) {
                            $groupDetails = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails details for $PrimarySMTPAddress" -DefaultValue $null -ExportFileLocation $ExportDetails -ScriptBlock {
                                Get-UnifiedGroup $identity -ErrorAction Stop
                            }
                            if ($groupDetails) {
                                $metadataFetchedFromExoCount++
                            }
                        }
                        elseif (-not $groupDetails) {
                            $groupDetails = $object
                            $metadataReusedFromRecipientCount++
                        }

                        if ($collectExchangeGroupMemberExpansion -and -not $usedUnifiedGroupCache) {
                            $groupMembers = Invoke-ArrayaCollectionStepSafe -OperationName "Get-ExchangeGroupDetails members for $PrimarySMTPAddress" -DefaultValue @() -ExportFileLocation $ExportDetails -ScriptBlock {
                                @(Get-UnifiedGroupLinks -Identity $identity -LinkType Member -ResultSize unlimited -ErrorAction Stop)
                            }
                        }
                        else {
                            $groupMembers = @()
                        }
                    }
                }

                if (-not $groupDetails) {
                    $groupDetails = $object
                    $metadataReusedFromRecipientCount++
                }
    
                #Check Group Owners Size and Get Owners Addresses
                if ($object.ManagedBy.count -ge 1) {
                    $groupOwners = $object.ManagedBy
                    Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] '{0}' Owners Found for '{1}' '{2}'" -f $object.ManagedBy.count, $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails
    
                }
                #Check Group Members Size and Get Group Addresses
                if ($groupMembers.count -ge 1) {
                    Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] '{0}' Members Found for '{1}' '{2}'" -f $groupMembers.count, $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails
                }
                if ($groupMembersCount -le 0) {
                    $groupMembersCount = ($groupMembers | Measure-Object).Count
                }
                Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] Create Group Output Details for '{0}' '{1}'" -f $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails
    
                #Output Group Details
                $currentobject = [PSCustomObject]@{
                    DisplayName                              = $object.DisplayName
                    Identity                                 = $identity
                    Alias                                    = $object.alias
                    Notes                                    = $object.Notes
                    IsDirSynced                              = $groupDetails.IsDirSynced
                    HiddenFromAddressListsEnabled            = $object.HiddenFromAddressListsEnabled
                    PrimarySMTPAddress                       = $object.PrimarySMTPAddress
                    RecipientTypeDetails                     = $object.RecipientTypeDetails
                    ResourceProvisioningOptions              = ($groupDetails.ResourceProvisioningOptions -join ",")
                    IsMailboxConfigured                      = $groupDetails.IsMailboxConfigured
                    EmailAddresses                           = $object.EmailAddresses
                    OwnersCount                              = ($groupOwners | measure-object).count
                    MembersCount                             = $groupMembersCount
                    HiddenGroupMembershipEnabled             = ($groupDetails.HiddenGroupMembershipEnabled -join ",")
                    ModeratedBy                              = ($ModeratedByRecipients -join ",")
                    RequireSenderAuthenticationEnabled       = $groupDetails.RequireSenderAuthenticationEnabled
                    AcceptMessagesOnlyFrom                   = ($groupDetails.AcceptMessagesOnlyFrom -join ",")
                    AcceptMessagesOnlyFromDLMembers          = ($groupDetails.AcceptMessagesOnlyFromDLMembers -join ",")
                    AcceptMessagesOnlyFromSendersOrMembers   = ($groupDetails.AcceptMessagesOnlyFromSendersOrMembers -join ",")
                    RejectMessagesFrom                       = ($groupDetails.RejectMessagesFrom -join ",")
                    RejectMessagesFromDLMembers              = ($groupDetails.RejectMessagesFromDLMembers -join ",")
                    RejectMessagesFromSendersOrMembers       = ($groupDetails.RejectMessagesFromSendersOrMembers -join ",")
                    AccessType                               = $groupDetails.AccessType
                    AllowAddGuests                           = $groupDetails.AllowAddGuests
                    SharePointSiteUrl                        = $groupDetails.SharePointSiteUrl
                }
    
                Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] Add '{0}' '{1}' Group Details to Tenant Stats Hash" -f $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails
                $script:tenantStatsHash["AllExchangeGroups"][$object.identity] = $currentobject
            }
            catch {
                Write-Log -Type ERROR -Message "[Get-ExchangeGroupDetails] An error occurred in running Get-ExchangeGroupDetails function. Exception: $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
        }
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Group metadata source breakdown: RecipientCache=$metadataReusedFromRecipientCount; UnifiedCache=$metadataReusedFromUnifiedCacheCount; EXOMetadataCalls=$metadataFetchedFromExoCount; TotalGroups=$totalCount." -ExportFileLocation $ExportDetails
        Write-Host ("  Group metadata source breakdown: Recipient cache={0}, Unified cache={1}, EXO metadata calls={2}" -f $metadataReusedFromRecipientCount, $metadataReusedFromUnifiedCacheCount, $metadataFetchedFromExoCount) -ForegroundColor DarkGray
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-ExchangeGroupDetails] An error occurred in running Get-ExchangeGroupDetails function. Exception: $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-ProgressHelper -Total $exchangeGroupProgressTotal -Id $exchangeGroupProgressId -Activity "Gathering All Exchange Online Group Details" -Completed
    }
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] COMPLETED: Gathering all Exchange Online Groups in $($CompletedTime)" -ExportFileLocation $ExportDetails
}

# Get Mail Flow Rules and Connectors
function Get-MailFlowRulesandConnectors {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    

    $start = Get-Date
    $mailFlowProgressId = 34
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    # Create Hash Tables for Mail Flow Rules and Connectors
    $script:tenantStatsHash["MailFlowRules"] = @{}
    $script:tenantStatsHash["MailFlowConnectors"] = @{}

    Write-Host "Getting all Mail Flow Rules and Connectors ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] START: Gathering all Mail Flow Rules and Connectors with $($detailLevel) details" -ExportFileLocation $ExportDetails
    Write-Progress -Id $mailFlowProgressId -Activity "Getting all Mail Flow Rules details" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))

    ### Get Mail Flow Rules ### - START
    try {
        Write-Log -Type INFO -Message "Gathering all Mail Flow Rules" -ExportFileLocation $ExportDetails
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "Name", "State", "Mode", "Priority", "Description"
                )
                try { $mailFlowRules = Get-TransportRule -IncludeTestModeConnectors -ErrorAction Continue | Select-Object $DesiredProperties }
                catch {  $mailFlowRules = Get-TransportRule -ErrorAction Continue | Select-Object $DesiredProperties  }
            }
            geek {
                try {  $mailFlowRules = Get-TransportRule -IncludeTestModeConnectors -ErrorAction Continue  }
                catch {  $mailFlowRules = Get-TransportRule -ErrorAction Continue  }
            }
        }

        #convert Mail Flow Rules to Hash Table
        $progresscounter = 0
        $totalCount = ($mailFlowRules | Measure-Object).count
        foreach ($rule in $mailFlowRules) {
            $progresscounter++
            Write-Log -Type DEBUG -Message "[Get-MailFlowRulesandConnectors] Gathering Mail Flow Details for $($rule.Name): $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails
            $script:tenantStatsHash["MailFlowRules"][$rule.Priority] = $rule
        }
        Write-Progress -Id $mailFlowProgressId -Activity "Getting all Mail Flow Rules details" -Completed
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-MailFlowRulesandConnectors] An error occurred in gathering MailFlow Rules. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    ### Get Mail Flow Rules ### - END
    
    ### Get Mail Flow Connectors  ### - START
    function Add-ConnectorToHash {
        param (
            [Parameter(Mandatory=$true)]
            $connectorList,
            [Parameter(Mandatory=$true)]
            [ValidateSet('Inbound', 'Outbound')]
            $direction
        )
    
        $progressCounter = 0
        $totalCount = ($connectorList | Measure-Object).Count
        foreach ($connector in $connectorList) {
            $progressCounter++
            Write-Log -Type DEBUG -Message ("[Add-ConnectorToHash] ({0}/{1}) Gathering '{2}' Mail Connector Details for {3}" -f $progressCounter, $totalCount, $direction, $connector.ID) -ExportFileLocation $ExportDetails

            $currentConnector = $connector | Select-Object *
            $currentConnector | Add-Member -MemberType NoteProperty -Name "ConnectorDirection" -Value $direction -Force
    
            if ($detailLevel -ne "geek") {
                $propertiesToAdd = @("RecipientDomains", "SmartHosts", "ValidationRecipients", "SenderDomains", "SenderIPAddresses", "TrustedOrganizations", "EFSkipIPs", "EFSkipMailGateway", "EFUsers")
                foreach ($prop in $propertiesToAdd) {
                    if ($connector.$prop) {
                        $currentConnector | Add-Member -MemberType NoteProperty -Name $prop -Value ($connector.$prop -join ",") -Force
                    }
                }
            }    
            $script:tenantStatsHash["MailFlowConnectors"][$connector.Id] = $currentConnector
        }
    }

    try {
        # Gather Connector Details - Inbound and Outbound Connector
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Gathering all Inbound Mail Connectors" -ExportFileLocation $ExportDetails
        $mailFlowInboundConnectors = Get-InboundConnector -ErrorAction Stop
        #$mailFlowInboundConnectors | foreach { $_ | Add-Member -MemberType NoteProperty -Name "ConnectorDirection" -Value "Inbound" -Force }
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Found $(($mailFlowInboundConnectors| Measure-Object).count) Inbound Mail Connectors" -ExportFileLocation $ExportDetails

        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Gathering all Outbound Mail Connectors" -ExportFileLocation $ExportDetails
        $mailFlowOutboundConnectors = Get-OutboundConnector -IncludeTestModeConnectors $true -ErrorAction Stop
        #$mailFlowOutboundConnectors | foreach { $_ | Add-Member -MemberType NoteProperty -Name "ConnectorDirection" -Value "Outbound" -Force }
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Found $(($mailFlowOutboundConnectors | Measure-Object).count) Outbound Mail Connectors" -ExportFileLocation $ExportDetails

        # Filter properties if detail level is minimum, combined, or all
        if ($detailLevel -in @("minimum", "combined", "all")) {
            $DesiredProperties = @(
                "Id", "ConnectorDirection", "Comment", "Enabled", "TestMode", "ConnectorType",
                "UseMXRecord", "IsTransportRuleScoped", "RecipientDomains",
                @{Name="SmartHosts"; Expression={ if ($_.SmartHosts -is [array]) { $_.SmartHosts -join "," } else { $_.SmartHosts } }},
                "AllAcceptedDomains", "SenderRewritingEnabled",
                "RouteAllMessagesViaOnPremises", "CloudServicesMailEnabled",
                "ValidationRecipients", "Description", "IsValidated",
                "LastValidationTimestamp",
                @{Name="SenderDomains"; Expression={ if ($_.SenderDomains -is [array]) { $_.SenderDomains -join "," } else { $_.SenderDomains } }},
                @{Name="SenderIPAddresses"; Expression={ if ($_.SenderIPAddresses -is [array]) { $_.SenderIPAddresses -join "," } else { $_.SenderIPAddresses } }},
                @{Name="TrustedOrganizations"; Expression={ if ($_.TrustedOrganizations -is [array]) { $_.TrustedOrganizations -join "," } else { $_.TrustedOrganizations } }},
                "RequireTls", "TlsSettings", "TlsDomain",
                "TreatMessagesAsInternal", "EFTestMode", "EFSkipLastIP",
                @{Name="EFSkipIPs"; Expression={ if ($_.EFSkipIPs -is [array]) { $_.EFSkipIPs -join "," } else { $_.EFSkipIPs } }},
                @{Name="EFSkipMailGateway"; Expression={ if ($_.EFSkipMailGateway -is [array]) { $_.EFSkipMailGateway -join "," } else { $_.EFSkipMailGateway } }},
                @{Name="EFUsers"; Expression={ if ($_.EFUsers -is [array]) { $_.EFUsers -join "," } else { $_.EFUsers } }}
            )
            $mailFlowInboundConnectors = $mailFlowInboundConnectors | Select-Object $DesiredProperties
            $mailFlowOutboundConnectors = $mailFlowOutboundConnectors | Select-Object $DesiredProperties
        }

        # Ensure MailFlowConnectors is initialized as a hashtable
        $script:tenantStatsHash["MailFlowConnectors"] = @{}

        #convert Mail Flow Connectors to Hash Table
        if ($mailFlowInboundConnectors) {
            Add-ConnectorToHash -connectorList $mailFlowInboundConnectors -direction "Inbound"
        }
        if ($mailFlowOutboundConnectors) {
            Add-ConnectorToHash -connectorList $mailFlowOutboundConnectors -direction "Outbound"
        }
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] Add Mail Connectors Details to Tenant Stats Hash" -ExportFileLocation $ExportDetails
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-MailFlowRulesandConnectors] An error occurred in MailFlow Connectors. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-Progress -Id $mailFlowProgressId -Activity "Getting all Mail Flow Rules details" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] COMPLETED: Gathering All Mail Flow Rules and Connectors in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }
}

function Get-SMTPRelayConfiguration {
    param ()
    
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["SMTPRelayConfig"] = @{}
    
    Write-Host "Checking SMTP Relay Configuration ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-SMTPRelayConfiguration] START: Checking SMTP Relay Configuration" -ExportFileLocation $ExportDetails
    
    try {
        $smtpConfig = [PSCustomObject]@{
            SMTPAuthEnabled = $false
            ConnectorBasedRelay = $false
            DirectSendEnabled = $false
            RelayConnectors = @()
            SMTPAuthUsers = 0
        }
        
        # Check for SMTP AUTH enabled on mailboxes
        try {
            $smtpAuthUsers = $script:tenantStatsHash["AllMailboxes"].Values | 
                Where-Object {$_.SmtpClientAuthenticationDisabled -eq $false}
            
            if ($smtpAuthUsers) {
                $smtpConfig.SMTPAuthEnabled = $true
                $smtpConfig.SMTPAuthUsers = $smtpAuthUsers.Count
                Write-Log -Type INFO -Message "[Get-SMTPRelayConfiguration] Found $($smtpAuthUsers.Count) mailboxes with SMTP AUTH enabled" -ExportFileLocation $ExportDetails
            }
        } catch {
            Write-Log -Type WARNING -Message "[Get-SMTPRelayConfiguration] Unable to check SMTP AUTH on mailboxes" -ExportFileLocation $ExportDetails
        }
        
        # Check outbound connectors for relay configuration
        $outboundConnectors = $script:tenantStatsHash["MailFlowConnectors"].Values | 
            Where-Object {$_.ConnectorDirection -eq "Outbound" -and $_.Enabled -eq $true}
        
        if ($outboundConnectors) {
            foreach ($connector in $outboundConnectors) {
                # Check if connector is used for relaying (has SmartHosts or CloudServicesMailEnabled)
                if ($connector.SmartHosts -or $connector.CloudServicesMailEnabled) {
                    $smtpConfig.ConnectorBasedRelay = $true
                    
                    $relayInfo = [PSCustomObject]@{
                        ConnectorName = $connector.Id
                        SmartHosts = $connector.SmartHosts
                        UseMXRecord = $connector.UseMXRecord
                        CloudServicesMailEnabled = $connector.CloudServicesMailEnabled
                        TlsSettings = $connector.TlsSettings
                        ConnectorType = $connector.ConnectorType
                    }
                    $smtpConfig.RelayConnectors += $relayInfo
                }
            }
            Write-Log -Type INFO -Message "[Get-SMTPRelayConfiguration] Found $($smtpConfig.RelayConnectors.Count) relay connectors" -ExportFileLocation $ExportDetails
        }
        
        # Check for direct send capability (no authentication required for certain scenarios)
        $acceptedDomains = Get-AcceptedDomain -ErrorAction SilentlyContinue
        if ($acceptedDomains) {
            $authoritative = $acceptedDomains | Where-Object {$_.DomainType -eq "Authoritative"}
            if ($authoritative.Count -gt 0) {
                $smtpConfig.DirectSendEnabled = $true
                $smtpConfig | Add-Member -MemberType NoteProperty -Name "AuthoritativeDomains" -Value ($authoritative.DomainName -join ",")
            }
        }
        
        # Check organization config for SMTP client authentication
        try {
            $orgConfig = Get-TransportConfig -ErrorAction SilentlyContinue
            if ($orgConfig) {
                $smtpConfig | Add-Member -MemberType NoteProperty -Name "SmtpClientAuthenticationDisabled" -Value $orgConfig.SmtpClientAuthenticationDisabled
            }
        } catch {
            Write-Log -Type WARNING -Message "[Get-SMTPRelayConfiguration] Unable to check transport config" -ExportFileLocation $ExportDetails
        }
        
        $script:tenantStatsHash["SMTPRelayConfig"]["Configuration"] = $smtpConfig
        
    } catch {
        Write-Log -Type ERROR -Message "[Get-SMTPRelayConfiguration] Error checking SMTP relay configuration: $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-SMTPRelayConfiguration] COMPLETED: Checking SMTP Relay Configuration in $($CompletedTime)" -ExportFileLocation $ExportDetails
}

# Function to analyze Mail Flow Connectors for 3rd Party Spam Filtering
function Get-ThirdPartySpamFilteringConfig {
    param ()
    
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["SpamFilteringConfig"] = @{}
    
    Write-Host "Analyzing Mail Flow for 3rd Party Spam Filtering ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-ThirdPartySpamFilteringConfig] START: Analyzing mail flow connectors" -ExportFileLocation $ExportDetails
    
    try {
        $spamFilterConfig = [PSCustomObject]@{
            Uses3rdPartyFiltering = $false
            InboundConnectorCount = 0
            OutboundConnectorCount = 0
            InboundConnectors = @()
            OutboundConnectors = @()
            TransportRuleCount = 0
            TransportRulesWithTrustedIPs = 0
            TransportRuleIndicators = @()
            PotentialSpamFilters = @()
        }
        
        # Analyze inbound connectors
        $inboundConnectors = $script:tenantStatsHash["MailFlowConnectors"].Values | Where-Object {$_.ConnectorDirection -eq "Inbound"}
        $spamFilterConfig.InboundConnectorCount = ($inboundConnectors | Measure-Object).Count
        
        foreach ($connector in $inboundConnectors) {
            $connectorInfo = [PSCustomObject]@{
                Name = $connector.Id
                Enabled = $connector.Enabled
                SenderDomains = $connector.SenderDomains
                SenderIPAddresses = $connector.SenderIPAddresses
                RequireTls = $connector.RequireTls
                TreatMessagesAsInternal = $connector.TreatMessagesAsInternal
            }
            $spamFilterConfig.InboundConnectors += $connectorInfo
            
            # Check for common spam filter patterns
            if ($connector.SenderIPAddresses -or $connector.TlsSenderCertificateName) {
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $spamFilterConfig.PotentialSpamFilters += "Inbound: $($connector.Id)"
            }
        }
        
        # Analyze outbound connectors for relay
        $outboundConnectors = $script:tenantStatsHash["MailFlowConnectors"].Values | Where-Object {$_.ConnectorDirection -eq "Outbound"}
        $spamFilterConfig.OutboundConnectorCount = ($outboundConnectors | Measure-Object).Count
        
        foreach ($connector in $outboundConnectors) {
            $connectorInfo = [PSCustomObject]@{
                Name = $connector.Id
                Enabled = $connector.Enabled
                SmartHosts = $connector.SmartHosts
                UseMXRecord = $connector.UseMXRecord
                CloudServicesMailEnabled = $connector.CloudServicesMailEnabled
                TlsSettings = $connector.TlsSettings
            }
            $spamFilterConfig.OutboundConnectors += $connectorInfo
            
            # Check if routing through smart host (potential spam filter)
            if ($connector.SmartHosts -and $connector.Enabled) {
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $spamFilterConfig.PotentialSpamFilters += "Outbound: $($connector.Id) -> $($connector.SmartHosts)"
            }

            if ($connector.SmartHosts -match '(?i)proofpoint|ppe-hosted|pphosted|ironport|cisco|barracuda|mimecast|forcepoint|sophos|trendmicro|spamtitan') {
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $spamFilterConfig.PotentialSpamFilters += "Outbound SmartHost match: $($connector.Id) -> $($connector.SmartHosts)"
            }
        }

        # Analyze transport rules for trusted IPs / bypass patterns
        $transportRules = $script:tenantStatsHash["MailFlowRules"].Values
        $spamFilterConfig.TransportRuleCount = ($transportRules | Measure-Object).Count

        foreach ($rule in $transportRules) {
            $indicatorHits = @()
            $trustedIpValue = $null

            if ($rule.PSObject.Properties['SenderIPRanges'] -and $rule.SenderIPRanges) {
                $trustedIpValue = $rule.SenderIPRanges
                $indicatorHits += "SenderIPRanges"
            } elseif ($rule.PSObject.Properties['SenderIPAddresses'] -and $rule.SenderIPAddresses) {
                $trustedIpValue = $rule.SenderIPAddresses
                $indicatorHits += "SenderIPAddresses"
            }

            if ($rule.PSObject.Properties['SetSCL'] -and $rule.SetSCL -eq -1) {
                $indicatorHits += "SetSCL=-1"
            }
            if ($rule.PSObject.Properties['BypassSpamFiltering'] -and $rule.BypassSpamFiltering -eq $true) {
                $indicatorHits += "BypassSpamFiltering"
            }
            if ($rule.PSObject.Properties['HeaderContainsWords'] -and $rule.HeaderContainsWords -match '(?i)proofpoint|ironport|mimecast|barracuda|sophos|spamtitan') {
                $indicatorHits += "HeaderContainsWords"
            }
            if ($rule.PSObject.Properties['HeaderMatchesMessageHeader'] -and $rule.HeaderMatchesMessageHeader -match '(?i)proofpoint|ironport|mimecast|barracuda|sophos|spamtitan') {
                $indicatorHits += "HeaderMatchesMessageHeader"
            }

            if ($indicatorHits.Count -gt 0) {
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $summary = "Rule: $($rule.Name) ($($indicatorHits -join ', '))"
                if ($trustedIpValue) {
                    $summary += " | Trusted IPs: $trustedIpValue"
                    $spamFilterConfig.TransportRulesWithTrustedIPs++
                }
                $spamFilterConfig.TransportRuleIndicators += $summary
            }
        }
        
        # Check hosted filtering configuration
        try {
            $hostedFilter = Get-HostedConnectionFilterPolicy -ErrorAction SilentlyContinue
            if ($hostedFilter.IPAllowList.Count -gt 0) {
                $spamFilterConfig | Add-Member -MemberType NoteProperty -Name "IPAllowList" -Value ($hostedFilter.IPAllowList -join ",")
                $spamFilterConfig.Uses3rdPartyFiltering = $true
                $spamFilterConfig.TransportRulesWithTrustedIPs += $hostedFilter.IPAllowList.Count
                Write-Log -Type INFO -Message "[Get-ThirdPartySpamFilteringConfig] IP Allow List configured with $($hostedFilter.IPAllowList.Count) entries" -ExportFileLocation $ExportDetails
            }
        } catch {
            Write-Log -Type WARNING -Message "[Get-ThirdPartySpamFilteringConfig] Unable to check hosted connection filter policy" -ExportFileLocation $ExportDetails
        }
        
        $script:tenantStatsHash["SpamFilteringConfig"]["Configuration"] = $spamFilterConfig
        
    } catch {
        Write-Log -Type ERROR -Message "[Get-ThirdPartySpamFilteringConfig] Error analyzing spam filtering config: $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-ThirdPartySpamFilteringConfig] COMPLETED: Analyzing mail flow in $($CompletedTime)" -ExportFileLocation $ExportDetails
}

function Get-ExchangeHybridConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('minimum','combined','all','geek')]
        [string]$detailLevel
    )

    $start = Get-Date
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["HybridConfiguration"] = @{}

    Write-Host "Checking for Exchange Hybrid Configuration ..." -ForegroundColor Cyan -NoNewline
    Write-Log -Type INFO -Message "[Get-ExchangeHybridConfiguration] START" -ExportFileLocation $ExportDetails

    try {
        # --- Signals created by HCW / hybrid ---
        #$ioc  = Get-IntraOrganizationConnector -ErrorAction SilentlyContinue | Where-Object { $_.Enabled }
        $orgR = Get-OrganizationRelationship -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.TargetApplicationUri -or $_.TargetAutodiscoverEpr -or $_.DomainNames
                }
        $inb  = Get-InboundConnector  -ErrorAction SilentlyContinue | Where-Object { $_.ConnectorType -eq 'OnPremises' }
        $outb = Get-OutboundConnector -ErrorAction SilentlyContinue | Where-Object { $_.ConnectorType -eq 'OnPremises' }
        $mig  = Get-MigrationEndpoint -ErrorAction SilentlyContinue |
                Where-Object { $_.EndpointType -in @('ExchangeRemote','ExchangeRemoteMove') }

        # --- Additional hybrid hints (mail flow + org config) ---
        $hybridConfig = $null
        try {
            $hybridConfig = Get-HybridConfiguration -ErrorAction SilentlyContinue
        } catch {}

        $orgConfig = $null
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
        } catch {}

        $mailFlowConnectors = @()
        if ($script:tenantStatsHash -and $script:tenantStatsHash.ContainsKey("MailFlowConnectors")) {
            $mailFlowConnectors = $script:tenantStatsHash["MailFlowConnectors"].Values
        } else {
            try {
                $mailFlowConnectors = @(
                    (Get-InboundConnector -ErrorAction SilentlyContinue),
                    (Get-OutboundConnector -ErrorAction SilentlyContinue)
                ) | Where-Object { $_ }
            } catch {}
        }

        $onPremFlowConnectors = $mailFlowConnectors | Where-Object {
            $_.ConnectorType -eq 'OnPremises' -or $_.RouteAllMessagesViaOnPremises -eq $true
        }

        $signals = [ordered]@{
            #IntraOrganizationConnectors   = ($ioc | Select-Object -ExpandProperty Name)
            OrganizationRelationships     = ($orgR | Select-Object -ExpandProperty Name)
            InboundConnectorsOnPrem       = ($inb  | Select-Object -ExpandProperty Name)
            OutboundConnectorsOnPrem      = ($outb | Select-Object -ExpandProperty Name)
            MigrationEndpoints            = ($mig  | Select-Object -ExpandProperty RemoteServer)
            HybridConfigurationObject     = if ($hybridConfig) { $hybridConfig.Name } else { $null }
            HybridDomains                 = if ($orgConfig -and $orgConfig.PSObject.Properties['HybridDomains']) { ($orgConfig.HybridDomains -join ', ') } else { $null }
            MailFlowOnPremConnectors      = ($onPremFlowConnectors | Select-Object -ExpandProperty ID)
        }

        $evidence = New-Object System.Collections.Generic.List[string]
        if ($orgConfig -and $orgConfig.PSObject.Properties['HybridDomains'] -and $orgConfig.HybridDomains.Count -gt 0) {
            $evidence.Add("HybridDomains configured in Get-OrganizationConfig") | Out-Null
        }
        if ($ioc.Count -gt 0)  { $evidence.Add("IntraOrganizationConnector enabled ($($ioc.Count))") | Out-Null }
        if ($orgR.Count -gt 0) { $evidence.Add("OrganizationRelationship configured ($($orgR.Count))") | Out-Null }
        if ($inb.Count -gt 0)  { $evidence.Add("Inbound OnPremises connector(s) ($($inb.Count))") | Out-Null }
        if ($outb.Count -gt 0) { $evidence.Add("Outbound OnPremises connector(s) ($($outb.Count))") | Out-Null }
        if ($mig.Count -gt 0)  { $evidence.Add("Migration endpoint(s) ($($mig.Count))") | Out-Null }
        if ($onPremFlowConnectors.Count -gt 0) { $evidence.Add("Mail flow On-Premises connector(s) ($($onPremFlowConnectors.Count))") | Out-Null }

        $hasConnectorPair = (($inb.Count -gt 0) -and ($outb.Count -gt 0))
        $hasSingleConnector = (($inb.Count -gt 0) -xor ($outb.Count -gt 0))
        $hasMigrationEndpoint = ($mig.Count -gt 0)
        $hasHybridDomains = ($orgConfig -and $orgConfig.PSObject.Properties['HybridDomains'] -and $orgConfig.HybridDomains.Count -gt 0)
        $hasFederationOnly = ($ioc.Count -gt 0 -or $orgR.Count -gt 0)

        $isHybrid =
            $hasMigrationEndpoint -or
            $hasConnectorPair -or
            $hasHybridDomains

        $hybridStatus = if ($isHybrid) {
            "Hybrid"
        } elseif ($hasSingleConnector) {
            "Possible Hybrid"
        } elseif ($hasFederationOnly) {
            "Federation/Relationship"
        } else {
            "None"
        }

        $hybridType = if ($hasMigrationEndpoint) {
            "Migration/Remote Move"
        } elseif ($hasConnectorPair) {
            "Mail Flow (On-Premises connectors)"
        } elseif ($hasHybridDomains) {
            "Hybrid Domains"
        } elseif ($hasSingleConnector) {
            "Mail Flow (Single On-Premises connector)"
        } elseif ($hasFederationOnly) {
            "Free/Busy or OAuth Federation"
        } else {
            "None"
        }

        $details = [pscustomobject]@{
            IsHybridConfigured           = [bool]$isHybrid
            HybridStatus                 = $hybridStatus
            HybridType                   = $hybridType
            EvidenceCount                = $evidence.Count
            Evidence                     = ($evidence -join '; ')
            IntraOrgConnectorCount       = $ioc.Count
            OrgRelationshipCount         = $orgR.Count
            InboundOnPremConnectorCount  = $inb.Count
            OutboundOnPremConnectorCount = $outb.Count
            MigrationEndpointCount       = $mig.Count
            MailFlowOnPremConnectorCount = $onPremFlowConnectors.Count
        }

        # add richer details based on requested level
        if ($detailLevel -in @('combined','all','geek')) {
            #$details | Add-Member NoteProperty IntraOrgConnectors ($signals.IntraOrganizationConnectors -join ', ')
            $details | Add-Member NoteProperty OrganizationRelationships ($signals.OrganizationRelationships -join ', ')
            $details | Add-Member NoteProperty InboundOnPremConnectors ($signals.InboundConnectorsOnPrem -join ', ')
            $details | Add-Member NoteProperty OutboundOnPremConnectors ($signals.OutboundConnectorsOnPrem -join ', ')
            $details | Add-Member NoteProperty MigrationEndpoints ($signals.MigrationEndpoints -join ', ')
            $details | Add-Member NoteProperty HybridDomains ($signals.HybridDomains)
            $details | Add-Member NoteProperty MailFlowOnPremConnectors ($signals.MailFlowOnPremConnectors -join ', ')
        }

        $script:tenantStatsHash["HybridConfiguration"]["ExchangeHybrid"] = $details
        Write-Log -Type INFO -Message "[Get-ExchangeHybridConfiguration] Hybrid=$($details.IsHybridConfigured) Type=$($details.HybridType) Evidence=$($details.EvidenceCount)" -ExportFileLocation $ExportDetails

    } catch {
        Write-Log -Type WARNING -Message "[Get-ExchangeHybridConfiguration] Error: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        $script:tenantStatsHash["HybridConfiguration"]["ExchangeHybrid"] = [pscustomobject]@{
            IsHybridConfigured = $false
            Message            = "Unable to determine hybrid configuration"
            Error              = $_.Exception.Message
        }
    }

    $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
    Write-Host " Completed in $elapsed" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-ExchangeHybridConfiguration] COMPLETED in $elapsed" -ExportFileLocation $ExportDetails
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

#Public Folder Data; Statistics; Permissions Convert to Hash Tables
function Get-AllPublicFolderDetails {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$false, HelpMessage='Specify Exchange Environment')]
        [ValidateSet('On-Premises', 'Office365')]
        [string]$ExchangeEnvironment = 'Office365'     
    )
    $start = Get-Date
    $depthPolicy = if ($script:CollectionDepthPolicy) {
        $script:CollectionDepthPolicy
    } else {
        Get-ArrayaCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($detailLevel.ToLowerInvariant()))
    }
    $collectPublicFolderPermissions = $true
    if (
        ($detailLevel -eq 'combined') -or
        ($depthPolicy -and $depthPolicy.PSObject.Properties['IsCombined'] -and ($depthPolicy.IsCombined -eq $true))
    ) {
        $collectPublicFolderPermissions = $false
    }
    # Ensure global hash table structure
    if (-not $script:tenantStatsHash) {
        $script:tenantStatsHash = @{}
    }
    $script:tenantStatsHash["PublicFolderDetails"] = @{}
    Write-Host "Getting public folders, Stats and Perms ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] START: Gathering all public folder details with $($detailLevel) details" -ExportFileLocation $ExportDetails
    
    try {
        #Get Public Folder Data, Statistics, and Permissions
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "Identity", "Name", "MailEnabled"
                    "MailRecipientGuid", "ParentPath", "ContentMailboxName"
                    "EntryId", "FolderSize", "HasSubfolders"
                    "FolderClass", "FolderPath", "ExtendedFolderFlags"
                )
                if ($ExchangeEnvironment -eq 'On-Premises') {
                    $allPublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object { $_.Name -NE "IPM_SUBTREE" } | Select-Object $DesiredProperties
                } else {
                    $allPublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object { $_.Name -NE "IPM_SUBTREE" } | Select-Object $DesiredProperties
                }
            }
            geek {
                if ($ExchangeEnvironment -eq 'On-Premises') {
                    $allPublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object { $_.Name -NE "IPM_SUBTREE" }
                } else {
                    $allPublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object { $_.Name -NE "IPM_SUBTREE" }
                }
            }
        }
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $(($allPublicFolders | Measure-Object).count) Public Folders in Exchange" -ExportFileLocation $ExportDetails
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllPublicFolderDetails] An error occurred in running Get-AllPublicFolderDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_

    }
    
    # Public Folder Statistics
    #**************************
    Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Gathering all public folder statistics" -ExportFileLocation $ExportDetails
    try {
        if ($ExchangeEnvironment -eq 'On-Premises') {
            if (-not $PublicFolderDatabase) {
                $PublicFolderDatabase = (Get-PublicFolderDatabase -ErrorAction SilentlyContinue | Select-Object -First 1).Identity
                Write-Host "Using Public Folder Database: $PublicFolderDatabase" -ForegroundColor Yellow
            }
            $PublicFolderStatistics = $allPublicFolders | Get-PublicFolderStatistics -Server $PublicFolderDatabase -ErrorAction SilentlyContinue
        } else {
            $PublicFolderStatistics = $allPublicFolders | Get-PublicFolderStatistics -ErrorAction SilentlyContinue
        }
        $PublicFolderStatsHash = @{}
        foreach($publicFolderStat in $PublicFolderStatistics) {
            $key = $PublicFolderStat.EntryId
            $PublicFolderStatsHash[$key] = $PublicFolderStat
        }
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $($($PublicFolderStatistics | Measure-Object).count) Public Folder Statistics in Exchange" -ExportFileLocation $ExportDetails
    }
    catch {
        Write-Log -Type ERROR -Message "An error occurred in running Get-AllPublicFolderStatistics function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }

    # Public Folder Permissions
    #***************************
    $script:tenantStatsHash["PublicFolderPerms"] = @{}
    $publicFolderPermProgressId = 37
    $publicFolderPermProgressTotal = 1
    if ($collectPublicFolderPermissions) {
        try {
            Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Gathering all public folder permissions" -ExportFileLocation $ExportDetails
            $PublicFolderPermissions = $allPublicFolders | get-publicfolderclientpermission -ErrorAction SilentlyContinue
            Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $($PublicFolderPermissions.count) public folder permissions" -ExportFileLocation $ExportDetails
            
            Write-Host "Processing Public Folder Permissions..." -ForegroundColor Cyan -nonewline
            Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Processing all public folder permissions" -ExportFileLocation $ExportDetails
            $totalCount = ($PublicFolderPermissions | Measure-Object).count
            $publicFolderPermProgressTotal = [Math]::Max($totalCount, 1)
            foreach($publicFolderPermission in $PublicFolderPermissions) {
                Write-ProgressHelper -Total $publicFolderPermProgressTotal -Id $publicFolderPermProgressId -Activity "Processing all public folder permissions" -Operation "Gathering Public Folder Permissions for $($publicFolderPermission.Identity)"
                Write-Log -Type DEBUG -Message "[Get-AllPublicFolderDetails] Gathering Public Folder Permissions for $($publicFolderPermission.Identity)" -ExportFileLocation $ExportDetails

                $key = "$($publicFolderPermission.Identity)-$($publicFolderPermission.User.Displayname)"
                $permissionObject = @(
                    [PSCustomObject]@{
                        FolderName = $publicFolderPermission.FolderName
                        FolderPath = $publicFolderPermission.Identity
                        Displayname = $publicFolderPermission.User.Displayname
                        PrimarySMTPAddress = $publicFolderPermission.User.RecipientPrincipal.PrimarySmtpAddress
                        AccessRights = ($publicFolderPermission.AccessRights -join ",")
                    }
                )

                if($script:tenantStatsHash["PublicFolderPerms"].ContainsKey($key)) {
                    $script:tenantStatsHash["PublicFolderPerms"][$key] += $permissionObject
                }
                else {
                    $script:tenantStatsHash["PublicFolderPerms"][$key] = @($permissionObject)
                }
            }
        }
        catch {
            Write-Log -Type ERROR -Message "[Get-AllPublicFolderDetails] An error occurred in running Get-AllPublicFolderPermissions function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }
        finally {
            Write-ProgressHelper -Total $publicFolderPermProgressTotal -Id $publicFolderPermProgressId -Activity "Processing all public folder permissions" -Completed
        }
    }
    else {
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Combined mode optimization active. Skipping public folder permission expansion and exporting an empty PublicFolderPerms table." -ExportFileLocation $ExportDetails
        Write-Host "Skipping Public Folder Permissions in combined mode..." -ForegroundColor DarkGray -nonewline
    }
    
    #Combine Stats with Details
    $script:tenantStatsHash["PublicFolderDetails"] = @{}
    $totalCount = ($allPublicFolders | Measure-Object).count
    foreach($pf in $allPublicFolders) {
        $pfStatsCheck = $PublicFolderStatsHash[$pf.EntryId]
        Write-Log -Type DEBUG -Message "Combine Public Folder Stats for $($pf.FolderPath)" -ExportFileLocation $ExportDetails
        $pf | Add-Member -MemberType NoteProperty -Name "FolderPath" -Value ($pf.FolderPath -join ",") -Force
        $pf | Add-Member -MemberType NoteProperty -Name "ItemCount" -Value $pfStatsCheck.ItemCount -Force
        $pf | Add-Member -MemberType NoteProperty -Name "LastModificationTime" -Value $pfStatsCheck.LastModificationTime -Force
        $pf | Add-Member -MemberType NoteProperty -Name "OwnerCount" -Value $pfStatsCheck.OwnerCount -Force
        $pf | Add-Member -MemberType NoteProperty -Name "TotalAssociatedItemSize" -Value $pfStatsCheck.TotalAssociatedItemSize -Force
        $pf | Add-Member -MemberType NoteProperty -Name "TotalDeletedItemSize" -Value $pfStatsCheck.TotalDeletedItemSize -Force
        $pf | Add-Member -MemberType NoteProperty -Name "TotalItemSize" -Value $pfStatsCheck.TotalItemSize -Force
        $pf | Add-Member -MemberType NoteProperty -Name "MailboxOwnerId" -Value $pfStatsCheck.MailboxOwnerId -Force
        $script:tenantStatsHash["PublicFolderDetails"][$pf.Identity] = $pf
    }

    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] COMPLETED: Gathering all public folder details in $($CompletedTime)" -ExportFileLocation $ExportDetails
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
                foreach ($site in (Get-MgSite -All -Property $graphSiteProperties -ErrorAction Stop)) {
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
                    $DesiredProperties = @("Template", "IsHubSite", "Title", "LastContentModifiedDate", "Status", "ArchiveStatus", "StorageUsageCurrent", "LockState", "Url", "Owner", "StorageQuota", "GroupId", "IsTeamsConnected", "IsTeamsChannelConnected")
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
                    return Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputType PSObject -ErrorAction Stop
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

<#Get Teams Details
function Get-TeamsDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$false,HelpMessage='Provide the service name')]
        [ValidateSet('MGGraph', 'Teams')]
        [string]$ServiceName = 'MGGraph'
    )

    try {
        $start = Get-Date
        $script:tenantStatsHash["AllTeams"] = @{}

        Write-Host "Getting all Microsoft Teams details ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-TeamsDetails] START: Gathering all Microsoft Teams with $($detailLevel) details" -ExportFileLocation $ExportDetails

        $allTeams = @()
        switch ($ServiceName) {
            MGGraph {
                $allTeams = Get-MgTeam -All -ErrorAction Stop
            }
            Teams {
                # Fetch all Teams
                $ProgressPreference = "SilentlyContinue"
                $allTeams = Get-Team -ErrorAction Stop
                $ProgressPreference = "Continue"
            }
        }

        # Check for error of user not licensed to run Get-Team
        if ($allTeams -like "*ErrorMessage: Failed to get license information for the user. Ensure user has a valid Office365 license assigned to them*") {
            Write-Log -Type ERROR -Message "[Get-TeamsDetails] An error occurred in running Get-TeamsDetails function. Exception: Failed to get license information for the user. Ensure user has a valid Office365 license assigned to them" -ExportFileLocation $ExportDetails
            throw "Failed to get Teams Details. Ensure user has a valid Office365 license assigned to them"
        } else {
            $totalCount = $allTeams.count
            foreach ($team in $allTeams) {
                try {
                    Write-ProgressHelper -Total $totalCount -Activity "Gathering all Microsoft Teams with $($detailLevel) details" -Operation "Gathering Team Details for $($team.DisplayName)"
                    Write-Log -Type DEBUG -Message ("[Get-TeamsDetails] Gathering Team Details for {0}" -f $team.DisplayName) -ExportFileLocation $ExportDetails

                    # Fetch SharePoint site size for each Team
                    Write-Log -Type DEBUG -Message ("[Get-TeamsDetails] Gathering SharePoint Online Details for {0}" -f $team.DisplayName) -ExportFileLocation $ExportDetails
                    $SPOSiteDetails = $null
                    if ($script:tenantStatsHash.ContainsKey('SharePoint') -and $script:tenantStatsHash['SharePoint'] -is [System.Collections.IDictionary]) {
                        $sharePointRows = @($script:tenantStatsHash['SharePoint'].Values)
                        if ($team.PSObject.Properties['Id'] -and $team.Id) {
                            $SPOSiteDetails = $sharePointRows | Where-Object { $_.GroupId -and ([string]$_.GroupId -eq [string]$team.Id) } | Select-Object -First 1
                        }
                        if (-not $SPOSiteDetails -and $team.PSObject.Properties['DisplayName'] -and $team.DisplayName) {
                            $SPOSiteDetails = $sharePointRows | Where-Object { $_.Title -eq $team.DisplayName } | Select-Object -First 1
                        }
                    }
                    if ($SPOSiteDetails.Template -eq "TEAMCHANNEL#0" -or $SPOSiteDetails.Template -eq "GROUP#0") {
                        $siteSize = $SPOSiteDetails.StorageUsageCurrent
                        $siteSizeGB = [math]::Round($SPOSiteDetails.StorageUsageCurrent / 1024, 3)
                    } else {
                        $siteSize = 0
                        $siteSizeGB = 0
                    }

                    # Fetch Channels for each Team
                    Write-Log -Type DEBUG -Message ("[Get-TeamsDetails] Gathering Channels {0}" -f $team.DisplayName) -ExportFileLocation $ExportDetails
                    $channels = @()
                    switch ($ServiceName) {
                        MGGraph {
                            $teamId = $team.Id
                            $channels = Get-MgTeamChannel -TeamId $teamId -ErrorAction SilentlyContinue
                        }
                        Teams {
                            $teamId = $team.GroupID
                            $channels = Get-TeamChannel -GroupId $teamId -ErrorAction SilentlyContinue
                        }
                    }

                    # Initialize separate arrays for each channel type
                    $publicChannels = @()
                    $privateChannels = @()
                    $sharedChannels = @()
                    foreach ($channel in $channels) {
                        Write-Log -Type DEBUG -Message ("[Get-TeamsDetails] Gathering Channels Details for {0} - {1}" -f $team.DisplayName, $channel.DisplayName) -ExportFileLocation $ExportDetails
                        $channelType = if ($channel.MembershipType -eq "Private") { "private" } elseif ($channel.MembershipType -eq "Standard") { "public" } else { "shared" }

                        # Add the channel names to the respective arrays based on their type
                        switch ($channelType) {
                            "public" { $publicChannels += $channel.DisplayName }
                            "private" { $privateChannels += $channel.DisplayName }
                            "shared" { $sharedChannels += $channel.DisplayName }
                        }
                    }
                    $TotalNumberOfChannels = $publicChannels.Count + $privateChannels.Count + $sharedChannels.Count

                    # Create output object
                    $currentTeam = [ordered]@{
                        DisplayName       = $team.DisplayName
                        Description       = $team.Description
                        Visibility        = $team.Visibility
                        SharePointSiteUrl = $SPOSiteDetails.Url
                        "SiteSize-GB"     = $siteSizeGB
                        SiteSize          = $siteSize
                        TotalChannels     = $TotalNumberOfChannels
                        PublicChannels    = $publicChannels -join ','
                        PrivateChannels   = $privateChannels -join ','
                        SharedChannels    = $sharedChannels -join ','
                    }
                    $teamKeyBase = if ($team.PSObject.Properties['Id'] -and $team.Id) { [string]$team.Id } elseif ($team.PSObject.Properties['GroupId'] -and $team.GroupId) { [string]$team.GroupId } else { [string]$team.DisplayName }
                    $teamKey = $teamKeyBase
                    $teamDuplicateSuffix = 2
                    while ($script:tenantStatsHash["AllTeams"].ContainsKey($teamKey)) {
                        $teamKey = "{0}#{1}" -f $teamKeyBase, $teamDuplicateSuffix
                        $teamDuplicateSuffix++
                    }
                    $script:tenantStatsHash["AllTeams"][$teamKey] = $currentTeam
                }
                catch {
                    Write-Log -Type ERROR -Message "[Get-TeamsDetails] An error occurred in Gathering Teams Details for $($team.DisplayName). $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                }
            }
        }
    }
    catch {
        if ($_.Exception.Message -like "*Ensure user has a valid Office365 license assigned to them*") {
            Write-Log -Type ERROR -Message "[Get-TeamsDetails] An error occurred in running Get-TeamsDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }
        else {
            Write-Log -Type ERROR -Message "[Get-TeamsDetails] An error occurred in running Get-TeamsDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-ProgressHelper -Total 1 -Activity "Gathering all Microsoft Teams with $($detailLevel) details" -Completed
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-TeamsDetails] COMPLETED: Gathering all Microsoft Teams in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }
}
#>

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
function Get-GraphData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [int]$PageSize,

        [Parameter(Mandatory = $false)]
        [string]$Activity = "Fetching Data from Microsoft Graph",

        [Parameter(Mandatory = $false)]
        [string]$Operation = "Retrieving Data",

        [Parameter(Mandatory = $false)]
        [int]$Id = 1,

        [Parameter(Mandatory = $false)]
        [int]$ParentId,

        [Parameter(Mandatory = $false)]
        [switch]$UseRestMethod,

        [Parameter(Mandatory = $false)]
        [string]$AccessToken,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 5
    )

    begin {
        $headers = $null
        if ($global:GraphHeaders) {
            $headers = $global:GraphHeaders
        }
        elseif ($global:GraphToken) {
            $headers = @{
                'Content-Type'     = 'application/json'
                'Authorization'    = "Bearer $global:GraphToken"
                'ConsistencyLevel' = 'eventual'
            }
        }
        elseif ($AccessToken) {
            $headers = @{
                'Content-Type'     = 'application/json'
                'Authorization'    = "Bearer $AccessToken"
                'ConsistencyLevel' = 'eventual'
            }
        }
    }

    process {
        $result = Get-ArrayaGraphResource -Uri $Uri -PageSize $PageSize -Activity $Activity -PreferRest:$UseRestMethod -Headers $headers -MaxRetries $MaxRetries
        Write-ProgressHelper -Total 1 -Activity $Activity -Operation $Operation -Id $Id -Completed

        if ($null -eq $result) {
            return @()
        }

        if ($Uri -match '/\$count(\?|$)') {
            return $result
        }

        if ($result -is [array]) {
            return $result
        }
        if ($result -is [System.Collections.IEnumerable] -and -not ($result -is [string]) -and -not ($result -is [System.Collections.IDictionary])) {
            return @($result)
        }

        return @($result)
    }
}

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


#Verify all required modules are installed and Connect to Office 365 Services
function Connect-Office365 {
    <#
    .SYNOPSIS
        Flexible connection utility for Microsoft 365 services: Graph, ExchangeOnline, SharePointOnline, Teams.
    .DESCRIPTION
        Connects to one or more specified Microsoft 365 services: Microsoft Graph, Exchange Online, SharePoint Online Admin, or Teams.
        Handles PS7 module expectations, optional re-authentication, and supports application or certificate-based authentication (Graph/ExchangeOnline).
        Use -ServiceName to specify one, several, or ALL services to connect.
    .PARAMETER ServiceName
        One or more services to connect: 'Graph', 'ExchangeOnline', 'SharePointOnline', 'Teams', or 'ALL'.
    .PARAMETER WriteGraph
        Include write/elevated Graph scopes (delegate only).
    .PARAMETER Force
        Forces re-authentication (disconnects existing sessions where applicable).
    .PARAMETER TenantId
        Tenant ID (GUID). Required for Application/Certificate auth.
    .PARAMETER ClientSecretCredential
        Application client secret in PSCredential form (username = ClientId, password = client secret as SecureString).
        If not provided and required, will prompt.
    .PARAMETER CertificateThumbprint
        Thumbprint of the certificate (required for Certificate auth).
    .PARAMETER ClientId
        Application (Client) ID (required for Application/Certificate auth).
    .EXAMPLE
        Connect-Office365 -ServiceName ALL
    .EXAMPLE
        Connect-Office365 -ServiceName Graph,ExchangeOnline
    .EXAMPLE
        Connect-Office365 -ServiceName SharePointOnline
    .EXAMPLE
        Connect-Office365 -ServiceName Graph -WriteGraph
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('ExchangeOnline', 'EXO', 'MSOnline', 'MSO', 'Graph', 'MSGraph','MGGraph', 'SharePointOnline', 'SPO', 'ALL', 'Teams')]
        [string[]]$ServiceName = @('ALL'),
        [Parameter()]
        [switch]$WriteGraph,
        [Parameter()]
        [switch]$Force,
        [Parameter()]
        [string]$TenantId,
        [Parameter(Mandatory = $false, HelpMessage = "Enter a PSCredential where the username is your ClientId and the password is the app client secret as a SecureString. You can also run:`$ClientSecretCredential = Get-Credential -Message 'Please enter ClientId as username; password is the app client secret.'")]
        [ValidateNotNull()]
        [System.Management.Automation.Credential()]
        [System.Management.Automation.PSCredential]$ClientSecretCredential,
        [Parameter()]
        [string]$CertificateThumbprint,
        [Parameter()]
        [string]$ClientId
    )

    begin {
        function Resolve-AuthCertificate {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$Thumbprint
            )

            $normalizedThumbprint = $Thumbprint.Replace(' ', '').ToUpperInvariant()
            foreach ($storePath in @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')) {
                $certificate = Get-ChildItem -Path $storePath -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.Thumbprint -eq $normalizedThumbprint -and
                        $_.HasPrivateKey
                    } |
                    Sort-Object -Property NotAfter -Descending |
                    Select-Object -First 1

                if ($certificate) {
                    return $certificate
                }
            }

            throw "Certificate thumbprint '$Thumbprint' was not found in CurrentUser\\My or LocalMachine\\My with an accessible private key."
        }

        function Get-PlainTextSecretFromCredential {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.Management.Automation.PSCredential]$Credential
            )

            $networkCredential = $Credential.GetNetworkCredential()
            if (-not $networkCredential -or [string]::IsNullOrWhiteSpace($networkCredential.Password)) {
                throw "Client secret credential does not contain a usable secret."
            }

            return $networkCredential.Password
        }

        function Get-ClientSecretAccessToken {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$TenantId,
                [Parameter(Mandatory = $true)]
                [string]$ClientId,
                [Parameter(Mandatory = $true)]
                [System.Management.Automation.PSCredential]$ClientSecretCredential,
                [Parameter(Mandatory = $true)]
                [string]$Resource
            )

            $clientSecretPlainText = Get-PlainTextSecretFromCredential -Credential $ClientSecretCredential
            $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
            $scope = if ($Resource.EndsWith('/')) { "$Resource.default" } else { "$Resource/.default" }
            $tokenResponse = Invoke-QuietRestMethod -Parameters @{
                Method      = 'POST'
                Uri         = $tokenEndpoint
                ContentType = 'application/x-www-form-urlencoded'
                Body        = @{
                    client_id     = $ClientId
                    client_secret = $clientSecretPlainText
                    scope         = $scope
                    grant_type    = 'client_credentials'
                }
                ErrorAction = 'Stop'
            }

            if (-not $tokenResponse.access_token) {
                throw "No access token was returned for resource '$Resource'."
            }

            return $tokenResponse.access_token
        }

        function Initialize-GraphRestHeaders {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [ValidateSet('Delegate', 'Certificate', 'ClientSecret')]
                [string]$AuthenticationType,
                [Parameter(Mandatory = $false)]
                [string]$TenantId,
                [Parameter(Mandatory = $false)]
                [string]$ClientId,
                [Parameter(Mandatory = $false)]
                [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
                [Parameter(Mandatory = $false)]
                [System.Management.Automation.PSCredential]$ClientSecretCredential
            )

            $global:GraphToken = $null
            $global:GraphHeaders = $null

            if ($AuthenticationType -eq 'Delegate') {
                return
            }

            try {
                $graphAccessToken = $null
                if ($AuthenticationType -eq 'Certificate') {
                    # Avoid loading MSAL.PS in-session to prevent assembly version conflicts with ExchangeOnlineManagement.
                    Write-Verbose "Skipping certificate-based Graph REST header initialization. SDK context will be used for Graph requests."
                    return
                }
                elseif ($AuthenticationType -eq 'ClientSecret') {
                    $graphAccessToken = Get-ClientSecretAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecretCredential $ClientSecretCredential -Resource 'https://graph.microsoft.com'
                }

                if ([string]::IsNullOrWhiteSpace($graphAccessToken)) {
                    throw "Graph access token was empty."
                }

                $global:GraphToken = $graphAccessToken
                $global:GraphHeaders = @{
                    'Content-Type'     = 'application/json'
                    'Authorization'    = "Bearer $graphAccessToken"
                    'ConsistencyLevel' = 'eventual'
                }
                Write-Verbose "Initialized Graph REST headers for application authentication."
            }
            catch {
                Write-Log -Type WARNING -Message "[Connect-Office365] Unable to initialize Graph REST headers: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }
        }

        $result = [ordered]@{
            Graph              = $false
            TenantName         = $null
            OnPremisesSyncEnabled = $false
            OnPremisesLastSyncDateTime = $null
            ExchangeOnline     = $false
            SharePointOnline   = $false
            Teams              = $false
            Tenant             = $null
            SharePointAdmin    = $null
            AuthenticationType = $null
            InitialDomain      = $null
        }

        # Normalize ServiceName: treat 'ALL' as all services, translate aliases
        $selectedServices = @()
        $serviceLookup = @{
            "Graph"             = "Graph"
            "MSGraph"           = "Graph"
            "MGGraph"           = "Graph"
            "ExchangeOnline"    = "ExchangeOnline"
            "EXO"               = "ExchangeOnline"
            "MSOnline"          = "ExchangeOnline"
            "MSO"               = "ExchangeOnline"
            "SharePointOnline"  = "SharePointOnline"
            "SPO"               = "SharePointOnline"
            "Teams"             = "Teams"
        }
        if ($ServiceName -contains 'ALL') {
            $selectedServices = @('Graph','SharePointOnline','ExchangeOnline','Teams')
        } else {
            foreach ($svc in $ServiceName) {
                if ($serviceLookup.ContainsKey($svc)) {
                    $realService = $serviceLookup[$svc]
                    if (-not $selectedServices.Contains($realService)) {
                        $selectedServices += $realService
                    }
                }
            }
        }
        if (-not $selectedServices) {
            Throw "You must specify at least one valid -ServiceName (Graph, ExchangeOnline, SharePointOnline, Teams, or ALL)."
        }

        $usingApplicationAuth = $false
        $authCertificate = $null
        $exchangeAccessToken = $null
        $AuthenticationType = if ($CertificateThumbprint) { 
            'Certificate' 
        } elseif ($ClientSecretCredential) { 
            'ClientSecret' 
        } else { 
            'Delegate' 
        }
        $result.AuthenticationType = $AuthenticationType

        if ($AuthenticationType -eq 'ClientSecret' -or $AuthenticationType -eq 'Certificate') {
            $usingApplicationAuth = $true

            if ($AuthenticationType -eq 'Certificate') {
                $authCertificate = Resolve-AuthCertificate -Thumbprint $CertificateThumbprint
            }
            elseif ($AuthenticationType -eq 'ClientSecret') {
                # Ensure PSCredential is present, prompt if absent/invalid
                if (
                    ($null -eq $ClientSecretCredential) -or
                    -not ($ClientSecretCredential -is [System.Management.Automation.PSCredential]) -or
                    ([string]::IsNullOrEmpty($ClientSecretCredential.GetNetworkCredential().UserName)) -or
                    ([string]::IsNullOrEmpty($ClientSecretCredential.GetNetworkCredential().Password))
                ) {
                    $ClientIdPrompt = Read-Host "Please enter ClientId (username for the App Registration)"
                    $ClientSecretCredential = Get-Credential -UserName $ClientIdPrompt -Message "Please enter the client secret as the password"
                }
            }
        }
        Write-Verbose "Authentication Type: $AuthenticationType"
        Write-Verbose "Selected services for connect: $($selectedServices -join ', ')"
    }

    process {
        # Disconnect previous sessions if -Force specified
        if ($Force) {
            Write-Host "Disconnecting existing sessions ..." -ForegroundColor Cyan -NoNewline
            try { 
                Write-Verbose "Disconnecting from Microsoft Graph..."
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null 
            } catch {}
            try { 
                Write-Verbose "Disconnecting from Exchange Online..."
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 
            } catch {}
            try { 
                Write-Verbose "Disconnecting from Microsoft Teams..."
                Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue 
            } catch {}
            Write-Host "✓ Disconnected" -ForegroundColor Green
        }

        # ===== PRE-LOAD EXCHANGEONLINE MANAGEMENT MODULE: only if needed =====
        if ($selectedServices -contains "ExchangeOnline") {
            if (-not (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
                Write-Host "✗ Exchange Online: The module 'ExchangeOnlineManagement' is not installed. Please install it before proceeding." -ForegroundColor Red
                return
            }
            else { 
                if (-not (Get-Module -Name 'ExchangeOnlineManagement')) {
                    Write-Host "Importing ExchangeOnlineManagement module..." -ForegroundColor Cyan
                    Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop -WarningAction SilentlyContinue
                    Write-Host "ExchangeOnlineManagement module imported successfully." -ForegroundColor Green
                }
            }
        }

        # ===== CONNECT TO GRAPH FIRST, if selected =====
        $org = $null
        $TenantName = $null
        if ($selectedServices -contains "Graph") {
            $CommonScopes = @(
                "Directory.Read.All","Group.Read.All","GroupMember.Read.All",
                "User.Read.All","Sites.Read.All","Files.Read.All",
                "AuditLog.Read.All","Policy.Read.All",
                "Team.ReadBasic.All","TeamSettings.Read.All","TeamsTab.Read.All",
                "LicenseAssignment.ReadWrite.All","User.EnableDisableAccount.All","User.Export.All",
                "Device.Read.All","SecurityEvents.Read.All","SharePointTenantSettings.Read.All","Organization.Read.All",
                "Reports.Read.All","MailboxSettings.Read","Domain.Read.All","RoleManagement.Read.All",
                "Application.Read.All","DirectoryRecommendations.Read.All","CrossTenantInformation.ReadBasic.All",
                "Policy.ReadWrite.CrossTenantAccess"
            )
            $WriteScopes = @(
                "Directory.ReadWrite.All","Synchronization.ReadWrite.All","Organization.ReadWrite.All",
                "IdentityRiskyUser.ReadWrite.All","IdentityUserFlow.ReadWrite.All","User.ReadWrite.All",
                "UserAuthenticationMethod.ReadWrite.All","User-LifeCycleInfo.ReadWrite.All",
                "DeviceManagementManagedDevices.ReadWrite.All","Domain.ReadWrite.All",
                "Group.ReadWrite.All","GroupMember.ReadWrite.All","Sites.ReadWrite.All"
            )
            $GraphScopes = if ($AuthenticationType -eq 'Delegate' -and $WriteGraph) { 
                $CommonScopes + $WriteScopes 
            } elseif ($AuthenticationType -eq 'Delegate') { 
                $CommonScopes 
            } else { 
                @() 
            }

            #region Microsoft Graph PowerShell Module Import
            try {
                Write-Verbose "Checking for Microsoft.Graph module and dependencies (handling PS 7 compatibility)..."
                if ($PSVersionTable.PSVersion.Major -ge 7) {
                    # Skipping Graph import for PS7; best effort (let module autoload)
                } else {
                    $neededModules = @(
                        "Microsoft.Graph.Authentication","Microsoft.Graph.Users","Microsoft.Graph.Groups",
                        "Microsoft.Graph.Sites","Microsoft.Graph.Files","Microsoft.Graph.DirectoryObjects","Microsoft.Graph.Reports"
                    )
                    foreach ($mod in $neededModules) {
                        if (Get-Module -ListAvailable -Name $mod) {
                            if (-not (Get-Module -Name $mod)) {
                                Write-Verbose "Importing module $mod..."
                                Import-Module $mod -ErrorAction Stop
                            } else {
                                Write-Verbose "Module $mod is already imported."
                            }
                        } else {
                            throw "Module $mod is not installed."
                        }
                    }
                }
            #endregion Microsoft Graph PowerShell Module Import

            #region Check if already connected to Microsoft Graph
                $existing = $null
                try { 
                    Write-Verbose "Checking if already connected to Microsoft Graph..."
                    $existing = Get-MgContext -ErrorAction Stop 
                } catch {}
                
                if ($existing -and -not $Force) {
                    Write-Host "✓ Graph (already connected)" -ForegroundColor Green
                    Write-Verbose "Existing Microsoft Graph session detected via '$($existing.AuthType)', skipping reconnect."
                } else {
                    try {
                        switch ($AuthenticationType) {
                            'Certificate' {
                                Write-Verbose "Using certificate-based authentication for Graph."
                                if (-not $TenantId) { throw "Graph Certificate auth requires -TenantId." }
                                if (-not $ClientId) { throw "Graph Certificate auth requires -ClientId." }
                                Write-Host "Connecting to Graph (certificate)..." -ForegroundColor Cyan
                                Write-Verbose "Running Connect-MgGraph with -TenantId $TenantId -ClientId $ClientId -Certificate <resolved certificate>"
                                Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Certificate $authCertificate -NoWelcome -ErrorAction Stop | Out-Null
                            }
                            'ClientSecret' {
                                Write-Verbose "Using application secret authentication for Graph."
                                if (-not $TenantId) { throw "Graph Application auth requires -TenantId." }
                                Write-Host "Connecting to Graph (application)..." -ForegroundColor Cyan
                                Write-Verbose "Running Connect-MgGraph with -TenantId $TenantId -ClientSecretCredential <PSCredential>"
                                Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome -ErrorAction Stop | Out-Null
                            }
                            Default {
                                Write-Verbose "Using delegated authentication for Graph. Scopes: $($GraphScopes -join ', ')"
                                Write-Host "Connecting to Graph (delegate)..." -ForegroundColor Cyan
                                Write-Verbose "Running Connect-MgGraph with $($GraphScopes.Count) scopes"
                                Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop | Out-Null
                            }
                        }
                        Write-Host "✓ Graph connected" -ForegroundColor Green
                    }
                    catch {
                        if ($AuthenticationType -eq 'Certificate') {
                            Write-Host "✗ Graph: $($_.Exception.Message)" -ForegroundColor Red
                            Write-Warning "Verify the certificate thumbprint is installed in CurrentUser\\My or LocalMachine\\My on this machine, includes the private key, and that the app registration allows certificate auth for ClientId $ClientId."
                        } else {
                            Write-Host "✗ Graph: $($_.Exception.Message)" -ForegroundColor Red
                        }
                        $result.Graph = $false
                        return
                    }
                }

                Initialize-GraphRestHeaders -AuthenticationType $AuthenticationType -TenantId $TenantId -ClientId $ClientId -Certificate $authCertificate -ClientSecretCredential $ClientSecretCredential
                
                Write-Verbose "Retrieving organization info from Graph..."
                $org = Get-AssessmentTenantOrganization
                if (-not $org) {
                    throw "Unable to resolve tenant organization details from Microsoft Graph."
                }
                $result.OnPremisesSyncEnabled = $org.OnPremisesSyncEnabled
                $result.OnPremisesLastSyncDateTime = $org.OnPremisesLastSyncDateTime
                $result.Graph = $true
                $result.TenantName = $org.DisplayName
                Write-Verbose "Connected to tenant: '$($result.TenantName)'"
            }
            catch {
                Write-Host "✗ Graph: $($_.Exception.Message)" -ForegroundColor Red
                return
            }
            #endregion Check if already connected to Microsoft Graph

            #region Get Tenant Name from Graph if needed
            try {
                Write-Verbose "Extracting tenant name from organization domains..."
                $verifiedDomains = $org.VerifiedDomains
                $defaultDomain = $verifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1
                if (-not $defaultDomain) { $defaultDomain = $verifiedDomains | Select-Object -First 1 }
                if (-not $defaultDomain) { 
                    Write-Host "✗ Could not determine SharePoint tenant name: Unable to determine tenant default domain from Graph Organization object." -ForegroundColor Red
                    return
                }
                $domainPart = $defaultDomain.Name.Split('.')[0]
                $result.InitialDomain = $defaultDomain.Name
                $TenantName = $domainPart
                Write-Verbose "Using tenant name '$TenantName' derived from domain '$($defaultDomain.Name)'"
            }
            catch {
                Write-Host "✗ Could not determine SharePoint tenant name: $($_.Exception.Message)" -ForegroundColor Red
                return
            }
            #endregion Get Tenant Name from Graph if needed
        }

        # ===== CONNECT TO EXCHANGE ONLINE BEFORE SPO CERT AUTH =====
        if ($selectedServices -contains "ExchangeOnline" -and -not $result.ExchangeOnline) {
            try {
                $existingEXOOrg = $null
                try {
                    Write-Verbose "Checking if already connected to Exchange Online..."
                    $existingEXOOrg = (Get-OrganizationConfig -ErrorAction Stop).Name
                } catch {}

                if ($existingEXOOrg -and -not $Force) {
                    Write-Host "✓ Exchange Online (already connected)" -ForegroundColor Green
                    Write-Verbose "Existing Exchange Online session detected for '$existingEXOOrg', skipping reconnect."
                    $result.ExchangeOnline = $true
                } else {
                    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
                    if ($authenticationType -eq 'Certificate') {
                        Write-Verbose "Using certificate-based authentication for Exchange Online."
                        if (-not $ClientId) {
                            Write-Host "✗ Exchange Online: ExchangeOnline certificate auth requires -ClientId (AppId)." -ForegroundColor Red
                            return
                        }
                        if (-not $result.InitialDomain) {
                            Write-Host "✗ Exchange Online: ExchangeOnline certificate auth requires Organization (initial domain)." -ForegroundColor Red
                            return
                        }
                        Write-Verbose "Running Connect-ExchangeOnline -AppId $ClientId -Organization $($result.InitialDomain) -CertificateThumbprint $CertificateThumbprint"
                        Connect-ExchangeOnline -AppId $ClientId -Organization $result.InitialDomain -CertificateThumbprint $CertificateThumbprint -ShowBanner:$false -ErrorAction Stop | Out-Null
                    }
                    elseif ($authenticationType -eq 'ClientSecret') {
                        Write-Verbose "Using client-secret app authentication for Exchange Online."
                        if (-not $ClientId) {
                            Write-Host "✗ Exchange Online: Exchange Online client secret auth requires -ClientId (AppId)." -ForegroundColor Red
                            return
                        }
                        if (-not $TenantId) {
                            Write-Host "✗ Exchange Online: Exchange Online client secret auth requires -TenantId." -ForegroundColor Red
                            return
                        }
                        if (-not $result.InitialDomain) {
                            Write-Host "✗ Exchange Online: Exchange Online client secret auth requires Organization (initial domain)." -ForegroundColor Red
                            return
                        }
                        $exchangeAccessToken = Get-ClientSecretAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecretCredential $ClientSecretCredential -Resource 'https://outlook.office365.com'
                        Connect-ExchangeOnline -AccessToken $exchangeAccessToken -Organization $result.InitialDomain -ShowBanner:$false -ErrorAction Stop | Out-Null
                    }
                    else {
                        Write-Verbose "Using delegated authentication for Exchange Online."
                        Write-Verbose "Running Connect-ExchangeOnline with ShowBanner disabled"
                        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
                    }
                    $result.ExchangeOnline = $true
                    Write-Host "✓ Exchange Online connected" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "✗ Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
                return
            }
        }

        # ===== CONNECT TO SHAREPOINT ONLINE (ADMIN) =====
        if ($selectedServices -contains "SharePointOnline") {
            # Graph is required for tenant name/initial domain for SharePointOnline
            if (-not $TenantName) {
                Throw "Unable to proceed: TenantName could not be determined (Graph connection required)."
            }
            if ($AuthenticationType -eq 'Certificate' -and $PSVersionTable.PSVersion.Major -ge 7) {
                Write-Warning "Skipping SharePoint Online PowerShell certificate connection in PowerShell 7. Microsoft Graph will be used for site discovery."
                $result.SharePointOnline = $false
                $result.SharePointAdmin = $null
            }
            else {
            try {
                Write-Verbose "Checking for SharePoint module 'Microsoft.Online.SharePoint.PowerShell'..."
                if (-not (Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell')) {
                    Write-Host "✗ SharePoint Online: The module 'Microsoft.Online.SharePoint.PowerShell' is not installed. Please install it before proceeding." -ForegroundColor Red
                    return
                }
                if (-not (Get-Module -Name 'Microsoft.Online.SharePoint.PowerShell')) {
                    Write-Host "Importing SharePoint module..." -ForegroundColor Cyan
                    if ($PSVersionTable.PSVersion.Major -ge 7 -and $AuthenticationType -ne 'Certificate') {
                        Write-Verbose "Importing SharePoint module using Windows PowerShell context (for PS7+ compatibility)..."
                        Import-Module 'Microsoft.Online.SharePoint.PowerShell' -UseWindowsPowerShell -ErrorAction Stop -WarningAction SilentlyContinue
                    } else {
                        Write-Verbose "Importing SharePoint module..."
                        Import-Module 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction Stop -WarningAction SilentlyContinue
                    }
                    Write-Host "SharePoint module imported successfully." -ForegroundColor Green
                } else {
                    Write-Verbose "SharePoint module already imported."
                }
                $spoAdminUrl = "https://$TenantName-admin.sharepoint.com"
                $existingAdminSite = $null
                try {
                    Write-Verbose "Checking if already connected to SharePoint Online..."
                    $tenant = Get-SPOTenant -ErrorAction Stop
                    if ($tenant) {
                        $existingAdminSite = $spoAdminUrl
                    }
                } catch {}
                if ($existingAdminSite -and -not $Force) {
                    Write-Host "✓ SharePoint Online (already connected)" -ForegroundColor Green
                    Write-Verbose "Existing SharePoint Online session detected at '$existingAdminSite', skipping reconnect."
                    $result.SharePointOnline = $true
                    $result.SharePointAdmin = $existingAdminSite
                }
                else {
                    try {
                        switch ($authenticationType) {
                            'Certificate' {
                                Write-Verbose "Using certificate-based authentication for SharePoint Online ($spoAdminUrl)..."
                                if (-not $TenantId) { 
                                    Write-Host "✗ SharePoint Online: SharePoint Online Certificate auth requires -TenantId." -ForegroundColor Red
                                    return
                                }
                                if (-not $ClientId) { 
                                    Write-Host "✗ SharePoint Online: SharePoint Online Certificate auth requires -ClientId." -ForegroundColor Red
                                    return
                                }
                                Write-Host "Connecting to SharePoint Online (certificate)..." -ForegroundColor Cyan
                                Write-Verbose "Running Connect-SPOService with -ClientId $ClientId -TenantId $TenantId -Certificate <resolved certificate>"
                                Connect-SPOService -Url $spoAdminUrl -ClientId $ClientId -TenantId $TenantId -Certificate $authCertificate -ErrorAction Stop
                                $result.SharePointOnline = $true
                            }
                            'ClientSecret' {
                                Write-Warning "SharePoint Online does not support client secret authentication via Connect-SPOService. Will Rely on Microsoft Graph API instead."
                                $result.SharePointOnline = $false
                                continue
                            }
                            Default {
                                Write-Verbose "Connecting to SharePoint Online ($spoAdminUrl) with Delegate authentication..."
                                Write-Host "Connecting to SharePoint Online (delegate)..." -ForegroundColor Cyan
                                Connect-SPOService -Url $spoAdminUrl -ErrorAction Stop
                                $result.SharePointOnline = $true
                            }
                        }
                        $result.SharePointAdmin = $spoAdminUrl
                        Write-Host "✓ SharePoint Online connected" -ForegroundColor Green
                    } catch {
                        Write-Host "✗ SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
                        if ($AuthenticationType -eq 'Certificate' -and $result.Graph) {
                            Write-Warning "SharePoint Online certificate connection failed. Continuing with Microsoft Graph for site discovery."
                            $result.SharePointOnline = $false
                        } else {
                            return
                        }
                    }
                }
            }
            catch { 
                Write-Host "✗ SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
                if ($AuthenticationType -eq 'Certificate' -and $result.Graph) {
                    Write-Warning "SharePoint Online certificate connection failed. Continuing with Microsoft Graph for site discovery."
                    $result.SharePointOnline = $false
                } else {
                    return 
                }
            }
            }
        }

        # ===== CONNECT TO EXCHANGE ONLINE =====
        if ($selectedServices -contains "ExchangeOnline" -and -not $result.ExchangeOnline) {
            try {
                $existingEXOOrg = $null
                try {
                    Write-Verbose "Checking if already connected to Exchange Online..."
                    $existingEXOOrg = (Get-OrganizationConfig -ErrorAction Stop).Name
                } catch {}
                
                if ($existingEXOOrg -and -not $Force) {
                    Write-Host "✓ Exchange Online (already connected)" -ForegroundColor Green
                    Write-Verbose "Existing Exchange Online session detected for '$existingEXOOrg', skipping reconnect."
                    $result.ExchangeOnline = $true
                } else {
                    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
                    if ($authenticationType -eq 'Certificate') {
                        Write-Verbose "Using certificate-based authentication for Exchange Online."
                        if (-not $ClientId) { 
                            Write-Host "✗ Exchange Online: ExchangeOnline certificate auth requires -ClientId (AppId)." -ForegroundColor Red
                            return
                        }
                        if (-not $result.InitialDomain) { 
                            Write-Host "✗ Exchange Online: ExchangeOnline certificate auth requires Organization (initial domain)." -ForegroundColor Red
                            return
                        }
                        Write-Verbose "Running Connect-ExchangeOnline -AppId $ClientId -Organization $($result.InitialDomain) -CertificateThumbprint $CertificateThumbprint"
                        Connect-ExchangeOnline -AppId $ClientId -Organization $result.InitialDomain -CertificateThumbprint $CertificateThumbprint -ShowBanner:$false -ErrorAction Stop | Out-Null
                    }
                    elseif ($authenticationType -eq 'ClientSecret') {
                        Write-Verbose "Using client-secret app authentication for Exchange Online."
                        if (-not $ClientId) { 
                            Write-Host "✗ Exchange Online: Exchange Online client secret auth requires -ClientId (AppId)." -ForegroundColor Red
                            return
                        }
                        if (-not $TenantId) { 
                            Write-Host "✗ Exchange Online: Exchange Online client secret auth requires -TenantId." -ForegroundColor Red
                            return
                        }
                        if (-not $result.InitialDomain) { 
                            Write-Host "✗ Exchange Online: Exchange Online client secret auth requires Organization (initial domain)." -ForegroundColor Red
                            return
                        }
                        $exchangeAccessToken = Get-ClientSecretAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecretCredential $ClientSecretCredential -Resource 'https://outlook.office365.com'
                        Connect-ExchangeOnline -AccessToken $exchangeAccessToken -Organization $result.InitialDomain -ShowBanner:$false -ErrorAction Stop | Out-Null
                    }
                    else {
                        Write-Verbose "Using delegated authentication for Exchange Online."
                        Write-Verbose "Running Connect-ExchangeOnline with ShowBanner disabled"
                        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
                    }
                    $result.ExchangeOnline = $true
                    Write-Host "✓ Exchange Online connected" -ForegroundColor Green
                }
            }
            catch { 
                Write-Host "✗ Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
                return 
            }
        }

        # ===== CONNECT TO TEAMS =====
        if ($selectedServices -contains "Teams") {
            if ($AuthenticationType -eq 'Certificate' -and $PSVersionTable.PSVersion.Major -ge 7) {
                Write-Warning "Skipping Microsoft Teams PowerShell certificate connection in PowerShell 7. Teams PowerShell data will be unavailable."
                $result.Teams = $false
            }
            elseif ($AuthenticationType -eq 'ClientSecret') {
                Write-Warning "Skipping Microsoft Teams PowerShell client secret connection. Teams PowerShell data will be unavailable in app-secret mode."
                $result.Teams = $false
            }
            else {
            try {
                if (-not (Get-Module -ListAvailable -Name 'MicrosoftTeams')) {
                    throw "The module 'MicrosoftTeams' is not installed. Please install it before proceeding."
                }
                if (-not (Get-Module -Name 'MicrosoftTeams')) {
                    if ($PSVersionTable.PSVersion.Major -ge 7 -and $AuthenticationType -ne 'Certificate') {
                        Write-Verbose "Importing MicrosoftTeams module using Windows PowerShell context (for PS7+ compatibility)..."
                        Import-Module 'MicrosoftTeams' -UseWindowsPowerShell -ErrorAction Stop
                    } else {
                        Write-Verbose "Importing MicrosoftTeams module..."
                        Import-Module 'MicrosoftTeams' -ErrorAction Stop
                    }
                    Write-Host "MicrosoftTeams module imported successfully." -ForegroundColor Green
                } else {
                    Write-Verbose "MicrosoftTeams module already imported."
                }

                if ($AuthenticationType -eq 'Certificate') {
                    if (-not $TenantId) {
                        Write-Host "✗ Microsoft Teams: Teams certificate auth requires -TenantId." -ForegroundColor Red
                        return
                    }
                    if (-not $ClientId) {
                        Write-Host "✗ Microsoft Teams: Teams certificate auth requires -ClientId (ApplicationId)." -ForegroundColor Red
                        return
                    }

                    Write-Verbose "Connecting to Microsoft Teams with certificate authentication..."
                    Write-Host "Connecting to Microsoft Teams (certificate)..." -ForegroundColor Cyan
                    Connect-MicrosoftTeams -TenantId $TenantId -ApplicationId $ClientId -Certificate $authCertificate -ErrorAction Stop | Out-Null
                    $result.Teams = $true
                    Write-Host "✓ Microsoft Teams connected" -ForegroundColor Green
                } else {
                    $existingTeamsOrg = $null
                    try {
                        Write-Verbose "Checking if already connected to Microsoft Teams..."
                        $existingTeamsOrg = (Get-CsTenant -ErrorAction Stop).DisplayName
                    } catch {}
                    
                    if ($existingTeamsOrg -and -not $Force) {
                        Write-Host "✓ Microsoft Teams (already connected)" -ForegroundColor Green
                        Write-Verbose "Existing Microsoft Teams session detected for '$existingTeamsOrg', skipping reconnect."
                        $result.Teams = $true
                    } else {
                        Write-Verbose "Connecting to Microsoft Teams (delegate authentication only)..."
                        Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
                        if ($CertificateThumbprint -or $usingApplicationAuth) {
                            Write-Warning "Microsoft Teams connection currently supports delegate authentication only. Attempting connection..."
                        }
                        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
                        $result.Teams = $true
                        Write-Host "✓ Microsoft Teams connected" -ForegroundColor Green
                    }
                }
            }
            catch { 
                Write-Host "✗ Microsoft Teams: $($_.Exception.Message)" -ForegroundColor Red
                if ($AuthenticationType -eq 'Certificate') {
                    Write-Warning "Microsoft Teams certificate connection failed. Continuing without Teams PowerShell data."
                    $result.Teams = $false
                } else {
                    return 
                }
            }
            }
        }
    }

    end {
        # Build summary based on connected subset
        $expected = @{}
        foreach ($svc in @('Graph','SharePointOnline','ExchangeOnline','Teams')) {
            if ($selectedServices -contains $svc) { $expected[$svc] = $true }
        }
        if (
            $AuthenticationType -eq 'Certificate' -and
            $selectedServices -contains 'SharePointOnline' -and
            -not $result.SharePointOnline -and
            $result.Graph
        ) {
            $expected.Remove('SharePointOnline')
        }
        if (
            $AuthenticationType -eq 'Certificate' -and
            $selectedServices -contains 'Teams' -and
            -not $result.Teams
        ) {
            $expected.Remove('Teams')
        }
        if (
            $AuthenticationType -eq 'ClientSecret' -and
            $selectedServices -contains 'SharePointOnline' -and
            -not $result.SharePointOnline -and
            $result.Graph
        ) {
            $expected.Remove('SharePointOnline')
        }
        if (
            $AuthenticationType -eq 'ClientSecret' -and
            $selectedServices -contains 'Teams' -and
            -not $result.Teams
        ) {
            $expected.Remove('Teams')
        }
        $allGood = $true
        foreach ($svc in $expected.Keys) {
            if (-not $result[$svc]) { $allGood = $false }
        }
        if (-not $allGood) {
            Throw "Failed to connect to one or more requested services. Check output above for connection problems encountered."
        } else {
            Write-Verbose "Returning connection result object."
            Write-Host "`nConnection Summary:" -ForegroundColor Cyan
            if ($selectedServices -contains "Graph") {
                Write-Host "  Tenant: $($result.TenantName)" -ForegroundColor White
                Write-Host "  Auth Type: $($result.AuthenticationType)" -ForegroundColor White
                Write-Host "  Initial Domain: $($result.InitialDomain)" -ForegroundColor White
                Write-Host "  Tenant DirSync Enabled: $($result.OnPremisesSyncEnabled)" -ForegroundColor White
                Write-Host "  Tenant DirSync Last Successful Sync: $($result.OnPremisesLastSyncDateTime)" -ForegroundColor White
            }
            if ($selectedServices -contains "SharePointOnline") {
                Write-Host "  SharePoint Admin: $($result.SharePointAdmin)" -ForegroundColor White
                Write-Host "  SharePoint Online: $($result.SharePointOnline)" -ForegroundColor White
            }
            if ($selectedServices -contains "ExchangeOnline") {
                Write-Host "  Exchange Online: $($result.ExchangeOnline)" -ForegroundColor White
            }
            if ($selectedServices -contains "Teams") {
                Write-Host "  Teams: $($result.Teams)" -ForegroundColor White
            }
        }
        return [pscustomobject]$result
    }
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
    $skus = Get-MgSubscribedSku -ErrorAction Continue | ? {$_.AppliesTo}
    $script:SkuLookupById = @{}
    $script:ServicePlanLookupById = @{}

    # Get subscription metadata to identify trials
    $subscriptions = @()
    try {
        $subscriptions = Get-MgDirectorySubscription -All -ErrorAction SilentlyContinue
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
        $adminRoles = Get-MgDirectoryRole | Select DisplayName, ID, Description | ? {$null -ne $_.DisplayName}
        $totalCount = $adminRoles.count
        Write-Log -Type INFO -Message "[Get-AllOffice365Admins] Admin Roles Found: $($adminRoles.count)" -ExportFileLocation $ExportDetails
        foreach ($role in $adminRoles) {
            $roleName = $role.DisplayName
            Write-Log -Type DEBUG -Message "[Get-AllOffice365Admins] $($roleName): Gathering Admins in Role" -ExportFileLocation $ExportDetails
            Write-ProgressHelper -Total $totalCount -Id 1 -Activity "Gathering Admins in Roles" -Operation "Checking Role: $($roleName)" 
            $roleMemberList = Invoke-ArrayaCollectionStepSafe -OperationName "Get-AllOffice365Admins role membership for $roleName" -DefaultValue @() -ExportFileLocation $ExportDetails -ScriptBlock {
                @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction Stop | Where-Object { $null -ne $_.Id })
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
        $domains = Get-MgDomain | ? {$null -ne $_.ID}
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
                $AuthenticationType = $domain.AuthenticationType
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

                #Check for Exchange Online Dependencies
                if ($exchangeOnline -eq $true) {
                    Write-Log -Type DEBUG -Message "[Get-AllOffice365Domains] Gathering Exchange Online Domain Details for '$($domainName)'" -ExportFileLocation $ExportDetails
                    $DomainType = $acceptedDomains | ?{$_.DomainName -eq $domainName} | Select-Object -ExpandProperty "DomainType"
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
        $userCount = $TenantStatsStore["Users"].Keys.Count

        foreach ($userKey in $TenantStatsStore["Users"].Keys) {
            $user = $TenantStatsStore["Users"][$userKey]
            if ($logPerRecordDebug) {
                Write-Log -Type DEBUG -Message ("[Combine-UserAndMailboxStats] Combining User '{0}' Details" -f $user.DisplayName) -ExportFileLocation $ExportDetails
            }

            Write-ProgressHelper -Total ([Math]::Max($userCount, 1)) -Id $combinedUserProgressId -Activity "Processing User Data" -Operation "Processing user: $($user.DisplayName)"

            $userDetails = Populate-Details -entity $user
            $TenantStatsStore["UserFullDetails"][$user.UserPrincipalName] = $userDetails
        }
        Write-ProgressHelper -Total ([Math]::Max($userCount, 1)) -Id $combinedUserProgressId -Activity "Processing User Data" -Completed

        # Process Mailboxes
        $allMailboxes = @($TenantStatsStore['AllRecipients'].Values | Where-Object { $_.RecipientTypeDetails -like "*Mailbox" })
        $mailboxCount = $allMailboxes.Count
        Write-Log -Type INFO -Message "[Combine-UserAndMailboxStats] Combining User and Mailbox Details - Processing Mailboxes" -ExportFileLocation $ExportDetails
        foreach ($mailbox in $allMailboxes) {
            #$mailbox = $TenantStatsStore["AllMailboxes"][$mailbox.PrimarySMTPAddress]
            if ($logPerRecordDebug) {
                Write-Log -Type DEBUG -Message ("[Combine-UserAndMailboxStats] Combining Mailbox '{0}' Details" -f $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails
            }
            Write-ProgressHelper -Total ([Math]::Max($mailboxCount, 1)) -Id $combinedMailboxProgressId -Activity "Processing Mailbox Data" -Operation "Processing mailbox: $($mailbox.PrimarySMTPAddress)"

            $mailboxDetails = Populate-Details -entity $mailbox -IsMailbox
            $TenantStatsStore["MailboxFullDetails"][$mailbox.PrimarySMTPAddress] = $mailboxDetails
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
        $inactiveMailboxes = $TenantStatsStore['InactiveMailboxes'].Values
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

# Combined Graph Report Details - Still in Progress
function Get-GraphUserStats {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet('D90', 'D180', 'D60', 'D30')]
        [String]
        $PeriodDuration = "D90"
    )

    # Initialize the tenant statistics hash table
    $script:tenantStatsHash["AllGraphUserStats"] = @{}

    $StartTime1 = Get-Date
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Gathering all User Combined Summary Details from Graph" -ExportFileLocation $ExportDetails

    Write-Host ""

    # Define the URIs for different data
    $dataUris = @{
        TeamsUserReportsURI     = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='$($PeriodDuration)')"
        OneDriveUsageUri        = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='$($PeriodDuration)')"
        EmailReportsUri         = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='$($PeriodDuration)')"
        MailboxUsageReportsUri  = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$($PeriodDuration)')"
        SPOUsageReportsUri      = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='$($PeriodDuration)')"
        YammerUsageReportsUri   = "https://graph.microsoft.com/v1.0/reports/getYammerActivityUserDetail(period='$($PeriodDuration)')"
        SignInUri               = "https://graph.microsoft.com/V1.0/users?`$select=displayName,userPrincipalName, mail, id, CreatedDateTime,signInActivity,UserType"
    }

    # Function to fetch data with error handling
    function Get-GraphDataWithLogging {
        param (
            [string]$Uri,
            [string]$DataName
        )
        try {
            Write-Host "Fetching $DataName from the Graph..." -NoNewline -ForegroundColor Cyan
            if ($DataName -eq "User Sign-In Data") {
                [array]$data = Get-GraphData -Uri $Uri -PageSize 999
            } else {
                $data = (Get-GraphData -Uri $Uri) -Replace "...Report Refresh Date", "Report Refresh Date" | ConvertFrom-Csv
            }
            Write-Host "Completed" -ForegroundColor Green
            return $data
        } catch {
            Write-Log -Type ERROR -Message "[Get-GraphUserStats] Unable to fetch $DataName with Graph REST API. Exception: $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            return @() # Return an empty array on error
        }
    }

    # Fetch Graph Data and store in $script:tenantStatsHash
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching Teams User Activity from Graph" -ExportFileLocation $ExportDetails 
    $script:tenantStatsHash["TeamsUserData"] = Get-GraphDataWithLogging -Uri $dataUris.TeamsUserReportsURI -DataName "Teams User Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching OneDrive User Activity from Graph" -ExportFileLocation $ExportDetails 
    $script:tenantStatsHash["OneDriveData"] = Get-GraphDataWithLogging -Uri $dataUris.OneDriveUsageUri -DataName "OneDrive Usage Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching Exchange User Activity from Graph" -ExportFileLocation $ExportDetails 
    $script:tenantStatsHash["EmailData"] = Get-GraphDataWithLogging -Uri $dataUris.EmailReportsUri -DataName "Exchange Activity Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching Exchange Mailbox Usage Report from Graph" -ExportFileLocation $ExportDetails 
    $script:tenantStatsHash["MailboxUsage"] = Get-GraphDataWithLogging -Uri $dataUris.MailboxUsageReportsUri -DataName "Mailbox Usage Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching SharePoint Online Usage Report from Graph" -ExportFileLocation $ExportDetails 
    $script:tenantStatsHash["SPOUsage"] = Get-GraphDataWithLogging -Uri $dataUris.SPOUsageReportsUri -DataName "SharePoint Activity Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching Yammer Usage Report from Graph" -ExportFileLocation $ExportDetails 
    $script:tenantStatsHash["YammerUsage"] = Get-GraphDataWithLogging -Uri $dataUris.YammerUsageReportsUri -DataName "Yammer Activity Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching User Sign In Report from Graph" -ExportFileLocation $ExportDetails 
    $script:tenantStatsHash["SignInData"] = Get-GraphDataWithLogging -Uri $dataUris.SignInUri -DataName "User Sign-In Data"

    # Create hash table for user sign-in data within $script:tenantStatsHash
    $script:tenantStatsHash["UserSignIns"] = @{}
    
    # Process User sign-in data and store in $script:tenantStatsHash["UserSignIns"]
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Processing User Sign-In Data fetched from Graph" -ExportFileLocation $ExportDetails
    [array]$UserSignInData = $script:tenantStatsHash["SignInData"] | Where-Object { $_.UserType -eq "Member" } | Sort-Object UserPrincipalName -Unique
    ForEach ($U in $UserSignInData) {
        If ($U.SignInActivity.LastSignInDateTime) {
            $LastSignInDate = Get-Date($U.SignInActivity.LastSignInDateTime) -format g
            $script:tenantStatsHash["UserSignIns"].Add([String]$U.UserPrincipalName, $LastSignInDate)
        } Else {
            $script:tenantStatsHash["UserSignIns"].Add([String]$U.UserPrincipalName, $Null)
        }
    }

    $StartTime2 = Get-Date
    Write-Host "Processing activity data fetched from the Graph..."
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Processing activity data fetched from the Graph" -ExportFileLocation $ExportDetails

    # Initialize the user data hash table within $script:tenantStatsHash
    $DataTable = @{}

    # Process Teams Data
    ForEach ($T in $script:tenantStatsHash["TeamsUserData"]) {
        If ([string]::IsNullOrEmpty($T."Last Activity Date")) { 
            $TeamsLastActivity = "No activity"
            $TeamsDaysSinceActive = "N/A" 
        } Else {
            $TeamsLastActivity = Get-Date($T."Last Activity Date") -format "dd-MMM-yyyy" 
            $TeamsDaysSinceActive = (New-TimeSpan($TeamsLastActivity)).Days 
        }
        $ReportLine  = [PSCustomObject]@{          
            TeamsUPN               = $T."User Principal Name"
            TeamsLastActive        = $TeamsLastActivity  
            TeamsDaysSinceActive   = $TeamsDaysSinceActive      
            TeamsReportDate        = Get-Date($T."Report Refresh Date") -format "dd-MMM-yyyy"  
            TeamsLicense           = $T."Assigned Products"
            TeamsChannelChats      = $T."Team Chat Message Count"
            TeamsPrivateChats      = $T."Private Chat Message Count"
            TeamsCalls             = $T."Call Count"
            TeamsMeetings          = $T."Meeting Count"
            TeamsRecordType        = "Teams"
        }
        $DataTable[$T."User Principal Name"] = $ReportLine
    } 

    # Process Exchange Data
    ForEach ($E in $script:tenantStatsHash["EmailData"]) {
        $ExoDaysSinceActive = $Null
        If ([string]::IsNullOrEmpty($E."Last Activity Date")) { 
            $ExoLastActivity = "No activity"
            $ExoDaysSinceActive = "N/A" 
        } Else {
            $ExoLastActivity = Get-Date($E."Last Activity Date") -format "dd-MMM-yyyy"
            $ExoDaysSinceActive = (New-TimeSpan($ExoLastActivity)).Days 
        }
        $ReportLine  = [PSCustomObject]@{          
            ExoUPN                = $E."User Principal Name"
            ExoDisplayName        = $E."Display Name"
            ExoLastActive         = $ExoLastActivity   
            ExoDaysSinceActive    = $ExoDaysSinceActive    
            ExoReportDate         = Get-Date($E."Report Refresh Date") -format "dd-MMM-yyyy"  
            ExoSendCount          = [int]$E."Send Count"
            ExoReadCount          = [int]$E."Read Count"
            ExoReceiveCount       = [int]$E."Receive Count"
            ExoIsDeleted          = $E."Is Deleted"
            ExoRecordType         = "Exchange Activity"
        }
        [Array]$ExistingData = $DataTable[$E."User Principal Name"] 
        [Array]$NewData = $ExistingData + $ReportLine
        $DataTable[$E."User Principal Name"] = $NewData
    } 

    # Process Mailbox Usage Data
    ForEach ($M in $script:tenantStatsHash["MailboxUsage"]) {
        If ([string]::IsNullOrEmpty($M."Last Activity Date")) { 
            $ExoLastActivity = "No activity" 
        } Else {
            $ExoLastActivity = Get-Date($M."Last Activity Date") -format "dd-MMM-yyyy"
            $ExoDaysSinceActive = (New-TimeSpan($ExoLastActivity)).Days 
        }
        $ReportLine  = [PSCustomObject]@{          
            MbxUPN                = $M."User Principal Name"
            MbxDisplayName        = $M."Display Name"
            MbxLastActive         = $ExoLastActivity 
            MbxDaysSinceActive    = $ExoDaysSinceActive          
            MbxReportDate         = Get-Date($M."Report Refresh Date") -format "dd-MMM-yyyy"  
            MbxQuotaUsed          = [Math]::Round($M."Storage Used (Byte)"/1GB,2) 
            MbxItems              = [int]$M."Item Count"
            MbxRecordType         = "Exchange Storage"
        }
        [Array]$ExistingData = $DataTable[$M."User Principal Name"] 
        [Array]$NewData = $ExistingData + $ReportLine
        $DataTable[$M."User Principal Name"] = $NewData
    } 

    # Process SharePoint Data
    ForEach ($S in $script:tenantStatsHash["SPOUsage"]) {
        If ([string]::IsNullOrEmpty($S."Last Activity Date")) { 
            $SPOLastActivity = "No activity"
            $SPODaysSinceActive = "N/A" 
        } Else {
            $SPOLastActivity = Get-Date($S."Last Activity Date") -format "dd-MMM-yyyy"
            $SPODaysSinceActive = (New-TimeSpan ($SPOLastActivity)).Days 
        }
        $ReportLine  = [PSCustomObject]@{          
            SPOUPN              = $S."User Principal Name"
            SPOLastActive       = $SPOLastActivity    
            SPODaysSinceActive  = $SPODaysSinceActive 
            SPOViewedEdited     = [int]$S."Viewed or Edited File Count"     
            SPOSyncedFileCount  = [int]$S."Synced File Count"
            SPOSharedExt        = [int]$S."Shared Externally File Count"
            SPOSharedInt        = [int]$S."Shared Internally File Count"
            SPOVisitedPages     = [int]$S."Visited Page Count" 
            SPORecordType       = "SharePoint Usage"
        }
        [Array]$ExistingData = $DataTable[$S."User Principal Name"] 
        [Array]$NewData = $ExistingData + $ReportLine
        $DataTable[$S."User Principal Name"] = $NewData
    }  

    # Process OneDrive Data
    ForEach ($O in $script:tenantStatsHash["OneDriveData"]) {
        $OneDriveLastActivity = $Null
        If ([string]::IsNullOrEmpty($O."Last Activity Date")) { 
            $OneDriveLastActivity = "No activity"
            $OneDriveDaysSinceActive = "N/A" 
        } Else {
            $OneDriveLastActivity = Get-Date($O."Last Activity Date") -format "dd-MMM-yyyy" 
            $OneDriveDaysSinceActive = (New-TimeSpan($OneDriveLastActivity)).Days 
        }
        $ReportLine  = [PSCustomObject]@{          
            ODUPN               = $O."Owner Principal Name"
            ODDisplayName       = $O."Owner Display Name"
            ODLastActive        = $OneDriveLastActivity    
            ODDaysSinceActive   = $OneDriveDaysSinceActive    
            ODSite              = $O."Site URL"
            ODFileCount         = [int]$O."File Count"
            ODStorageUsed       = [Math]::Round($O."Storage Used (Byte)"/1GB,4) 
            ODQuota             = [Math]::Round($O."Storage Allocated (Byte)"/1GB,2) 
            ODRecordType        = "OneDrive Storage"
        }
        [Array]$ExistingData = $DataTable[$O."Owner Principal Name"] 
        [Array]$NewData = $ExistingData + $ReportLine
        $DataTable[$O."Owner Principal Name"] = $NewData
    }  

    # Process Yammer Data
    ForEach ($Y in $script:tenantStatsHash["YammerUsage"]) {  
        If ([string]::IsNullOrEmpty($Y."Last Activity Date")) { 
            $YammerLastActivity = "No activity" 
            $YammerDaysSinceActive = "N/A" 
        } Else {
            $YammerLastActivity = Get-Date($Y."Last Activity Date") -format "dd-MMM-yyyy" 
            $YammerDaysSinceActive = (New-TimeSpan ($YammerLastActivity)).Days 
        }
        $ReportLine  = [PSCustomObject]@{          
            YUPN             = $Y."User Principal Name"
            YDisplayName     = $Y."Display Name"
            YLastActive      = $YammerLastActivity      
            YDaysSinceActive = $YammerDaysSinceActive   
            YPostedCount     = [int]$Y."Posted Count"
            YReadCount       = [int]$Y."Read Count"
            YLikedCount      = [int]$Y."Liked Count"
            YRecordType      = "Yammer Usage"
        }
        [Array]$ExistingData = $DataTable[$Y."User Principal Name"] 
        [Array]$NewData = $ExistingData + $ReportLine
        $DataTable[$Y."User Principal Name"] = $NewData
    }


    # Create set of users that we've collected data for
    [System.Collections.ArrayList]$Users = @()
    ForEach ($UserPrincipalName in $DataTable.Keys) { 
        If ($DataTable[$UserPrincipalName]) {
            $obj = [PSCustomObject]@{ 
                UPN  = $UserPrincipalName
            }
            $Users.add($obj) | Out-Null
        }
    }

    
    # Set up progress bar
    $StartTime3 = Get-Date
    $UserNumber = 0
    $TotalUsers = $Users.Count
    $graphUserExtractionProgressId = 84

    # Process each user to extract Exchange, Teams, OneDrive, SharePoint, and Yammer statistics for their activity
    ForEach ($UserPrincipalName in $Users) {
        $U = $UserPrincipalName.UPN
        $UserNumber++
        Write-ProgressHelper -Id $graphUserExtractionProgressId -Activity "Extracting User Data" -Operation "Processing user: $U" -Index $UserNumber -Total ([Math]::Max($TotalUsers, 1))

        Write-Log -Type INFO -Message "[Get-GraphUserStats] Processing User Data for $U" -ExportFileLocation $ExportDetails
        $UserData = $DataTable[$U]

        # Process Exchange Data
        [string]$ExoUPN = (Out-String -InputObject $UserData.ExoUPN).Trim()
        [string]$ExoLastActive = (Out-String -InputObject $UserData.ExoLastActive).Trim()
        If ([string]::IsNullOrEmpty($ExoUPN) -or $ExoLastActive -eq "No Activity") {
            $ExoDaysSinceActive  = "N/A"
            $ExoLastActive = "No Activity" 
        } Else {
            [string]$ExoLastActive = (Out-String -InputObject $UserData.ExoLastActive).Trim()
            [string]$ExoDaysSinceActive = (Out-String -InputObject $UserData.ExoDaysSinceActive).Trim() 
        }

        # Parse OneDrive for Business usage data 
        [string]$ODLastActive = (Out-String -InputObject $UserData.ODLastActive).Trim()
        If (($ODLastActive -Like "*No Activity*") -or ([string]::IsNullOrEmpty($ODLastActive))) {
            $ODLastActive = "No Activity"
            $ODDaysSinceActive  = "N/A"
            $ODFiles            = 0
            $ODStorage          = 0
            $ODQuota            = 1024 
        } Else {
            [string]$ODDaysSinceActive = (Out-String -InputObject $UserData.ODDaysSinceActive).Trim()
            [string]$ODLastActive = (Out-String -InputObject $UserData.ODLastActive).Trim()
            [string]$ODFiles = (Out-String -InputObject $UserData.ODFileCount).Trim()
            [string]$ODStorage = (Out-String -InputObject $UserData.ODStorageUsed).Trim()
            [string]$ODQuota = (Out-String -InputObject $UserData.ODQuota).Trim()  
        }

        # Parse Yammer usage data
        [string]$YUPN = (Out-String -InputObject $UserData.YUPN).Trim()
        [string]$YammerLastActive = (Out-String -InputObject $UserData.YLastActive).Trim()
        If (([string]::IsNullOrEmpty($YUPN) -or ($YammerLastActive -eq "No Activity"))) { 
            [string]$YammerLastActive = "No Activity"  
            [string]$YammerDaysSinceActive  = "N/A" 
            $YammerPosts             = 0
            $YammerReads             = 0
            $YammerLikes             = 0 
        } Else {
            $YammerDaysSinceActive = (Out-String -InputObject $UserData.YDaysSinceActive).Trim()
            $YammerPosts = (Out-String -InputObject $UserData.YPostedCount).Trim()
            $YammerReads = (Out-String -InputObject $UserData.YReadCount).Trim()
            $YammerLikes = (Out-String -InputObject $UserData.YLikedCount).Trim() 
        }

        # Parse Teams usage data
        If ($UserData.TeamsDaysSinceActive -gt 0) {
            [string]$TeamsDaysSinceActive = (Out-String -InputObject $UserData.TeamsDaysSinceActive).Trim()
            [string]$TeamsLastActive = (Out-String -InputObject $UserData.TeamsLastActive).Trim()
        } Else { 
            [string]$TeamsDaysSinceActive = "N/A"
            [string]$TeamsLastActive = "No Activity"
        }
        # Teams Usage Status
        if ($TeamsDaysSinceActive -le 30 -and $UserData.TeamsPrivateChats -gt 50) {
            [string]$TeamsUsage = "Active"
        } elseif ($TeamsDaysSinceActive -ge 31 -and $TeamsDaysSinceActive -le 60 -and $UserData.TeamsPrivateChats -ge 100) {
            [string]$TeamsUsage = "Low"
        } elseif ($TeamsDaysSinceActive -ge 31 -and $TeamsDaysSinceActive -le 90 -and $UserData.TeamsPrivateChats -le 100) {
            [string]$TeamsUsage = "Minimal"
        } else {
            [string]$TeamsUsage = "No Activity"
        }

        # Parse SharePoint Online usage data
        If ($UserData.SPODaysSinceActive -gt 0) {
            [string]$SPODaysSinceActive = (Out-String -InputObject $UserData.SPODaysSinceActive).Trim()
            [string]$SPOLastActive = (Out-String -InputObject $UserData.SPOLastActive).Trim() 
        } Else { 
            [string]$SPODaysSinceActive = "N/A"
            [string]$SPOLastActive = "No Activity" 
        }

        # Fetch the sign-in data if available
        $LastAccountSignIn = $Null; $DaysSinceSignIn = 0
        $LastAccountSignIn = $script:tenantStatsHash["UserSignIns"].Item($U)
        If ($null -eq $LastAccountSignIn) { 
            $LastAccountSignIn = "No sign in data found"; $DaysSinceSignIn = "N/A"
        } Else { 
            $DaysSinceSignIn = (New-TimeSpan($LastAccountSignIn)).Days 
        }
   
        # Determine if the account is in use
        [int]$ExoDays = 365; [int]$TeamsDays = 365; [int]$SPODays = 365; [int]$ODDays = 365; [int]$YammerDays = 365
        If ($ExoDaysSinceActive -ne "N/A") {$ExoDays = $ExoDaysSinceActive -as [int]}
        If ($TeamsDaysSinceActive -eq "N/A") {$TeamsDays = 365} Else {$TeamsDays = $TeamsDaysSinceActive -as [int]}
        If ($SPODaysSinceActive -eq "N/A") {$SPODays = 365} Else {$SPODays = $SPODaysSinceActive -as [int]}  
        If ($ODDaysSinceActive -eq "N/A") {$ODDays = 365} Else {$ODDays = $ODDaysSinceActive -as [int]} 
        If ($YammerDaysSinceActive -eq "N/A") {$YammerDays = 365} Else {$YammerDays = $YammerDaysSinceActive -as [int]}

        # Average days per workload used
        $AverageDaysSinceUse = [Math]::Round((($ExoDays + $TeamsDays + $SPODays + $ODDays)/4),2)

        Switch ($AverageDaysSinceUse) { 
            ({$PSItem -le 8})                          { $AccountStatus = "Heavy usage" }
            ({$PSItem -ge 9 -and $PSItem -le 50} )     { $AccountStatus = "Moderate usage" }   
            ({$PSItem -ge 51 -and $PSItem -le 120} )   { $AccountStatus = "Poor usage" }
            ({$PSItem -ge 121 -and $PSItem -le 300 } ) { $AccountStatus = "Review account"  }
            default                                    { $AccountStatus = "Account unused" }
        }

        # Override if someone has been active in any workload in the last 14 days
        [int]$DaysCheck = 14 
        If (($ExoDays -le $DaysCheck) -or ($TeamsDays -le $DaysCheck) -or ($SPODays -le $DaysCheck) -or ($ODDays -le $DaysCheck) -or ($YammerDays -le $DaysCheck)) {
            $AccountStatus = "Account in use"
        }

        # Build a line for the report file with the collected data for all workloads and write it to the list
        If ((![string]::IsNullOrEmpty($ExoUPN))) {
            $OutLine  = [PSCustomObject] @{          
                UPN                     = $U
                DisplayName             = (Out-String -InputObject $UserData.ExoDisplayName).Trim()
                Status                  = $AccountStatus
                LastSignIn              = $LastAccountSignIn
                DaysSinceSignIn         = $DaysSinceSignIn 
                EXOLastActive           = $ExoLastActive  
                EXODaysSinceActive      = $ExoDaysSinceActive  
                EXOQuotaUsed            = (Out-String -InputObject $UserData.MbxQuotaUsed).Trim()
                EXOItems                = (Out-String -InputObject $UserData.MbxItems).Trim()
                EXOSendCount            = (Out-String -InputObject $UserData.ExoSendCount).Trim()
                EXOReadCount            = (Out-String -InputObject $UserData.ExoReadCount).Trim()
                EXOReceiveCount         = (Out-String -InputObject $UserData.ExoReceiveCount).Trim()
                TeamsLastActive         = $TeamsLastActive
                TeamsDaysSinceActive    = $TeamsDays
                TeamsUsage              = $TeamsUsage
                TeamsChannelChat        = (Out-String -InputObject $UserData.TeamsChannelChats).Trim()
                TeamsPrivateChat        = (Out-String -InputObject $UserData.TeamsPrivateChats).Trim()
                TeamsMeetings           = (Out-String -InputObject $UserData.TeamsMeetings).Trim()
                TeamsCalls              = (Out-String -InputObject $UserData.TeamsCalls).Trim()
                SPOLastActive           = $SPOLastActive
                SPODaysSinceActive      = $SPODays 
                SPOViewedEditedFiles    = (Out-String -InputObject $UserData.SPOViewedEdited).Trim()
                SPOSyncedFiles          = (Out-String -InputObject $UserData.SPOSyncedFileCount).Trim()
                SPOSharedExtFiles       = (Out-String -InputObject $UserData.SPOSharedExt).Trim()
                SPOSharedIntFiles       = (Out-String -InputObject $UserData.SPOSharedInt).Trim()
                SPOVisitedPages         = (Out-String -InputObject $UserData.SPOVisitedPages).Trim()
                OneDriveLastActive      = $ODLastActive
                OneDriveDaysSinceActive = $ODDaysSinceActive
                OneDriveFiles           = $ODFiles
                OneDriveStorage         = $ODStorage
                OneDriveQuota           = $ODQuota
                YammerLastActive        = $YammerLastActive  
                YammerDaysSinceActive   = $YammerDaysSinceActive
                YammerPosts             = $YammerPosts
                YammerReads             = $YammerReads
                YammerLikes             = $YammerLikes
                License                 = (Out-String -InputObject $UserData.TeamsLicense).Trim()
                OneDriveSite            = (Out-String -InputObject $UserData.ODSite).Trim()
                IsDeleted               = (Out-String -InputObject $UserData.ExoIsDeleted).Trim()
                EXOReportDate           = (Out-String -InputObject $UserData.ExoReportDate).Trim()
                TeamsReportDate         = (Out-String -InputObject $UserData.TeamsReportDate).Trim()
                "AllServices-AverageDaysSinceUse"             = $AverageDaysSinceUse
            }
            $script:tenantStatsHash["AllGraphUserStats"][$U] = $OutLine
        } 
    }

    Write-ProgressHelper -Id $graphUserExtractionProgressId -Total ([Math]::Max($TotalUsers, 1)) -Activity "Extracting User Data" -Completed

    $StartTime4 = Get-Date
    $GraphTime = $StartTime2 - $StartTime1
    $PrepTime = $StartTime3 - $StartTime2
    $ReportTime = $StartTime4 - $StartTime3
    $ScriptTime = $StartTime4 - $StartTime1
    $AccountsPerMinute = [math]::Round(($script:tenantStatsHash["AllGraphUserStats"].Values.count/($ScriptTime.TotalSeconds/60)),2)
    $GraphElapsed = $GraphTime.Minutes.ToString() + ":" + $GraphTime.Seconds.ToString()
    $PrepElapsed = $PrepTime.Minutes.ToString() + ":" + $PrepTime.Seconds.ToString()
    $ReportElapsed = $ReportTime.Minutes.ToString() + ":" + $ReportTime.Seconds.ToString()
    $ScriptElapsed = $ScriptTime.Minutes.ToString() + ":" + $ScriptTime.Seconds.ToString()

    Write-Verbose ""
    Write-Verbose "Statistics for Graph Report Script V2.0"
    Write-Verbose "---------------------------------------"
    Write-Verbose "Time to fetch data from Microsoft Graph: $GraphElapsed"
    Write-Verbose "Time to prepare data for processing:     $PrepElapsed"
    Write-Verbose "Time to create report from data:         $ReportElapsed"
    Write-Verbose "Total time for script:                   $ScriptElapsed"
    Write-Verbose "Total accounts processed:                $($script:tenantStatsHash["AllGraphUserStats"].Values.count)"
    Write-Verbose "Accounts processed per minute:           $AccountsPerMinute"
    Write-Verbose ""
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
            $devices = Get-MgDevice -All -ErrorAction Stop | ? {$null -ne $_.ID}
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
            $conditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
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

        foreach ($policy in $conditionalAccessPolicies) {
            Write-ProgressHelper -Total $conditionalAccessProgressTotal -Id $conditionalAccessProgressId -Activity "Processing all Conditional Access Policies" -Operation "Expanding Policy Details for $($policy.DisplayName)"
            Write-Log -Type INFO -Message "[Get-ConditionalAccessPoliciesReport] Expanding Conditional Access Policies for $($policy.DisplayName)" -ExportFileLocation $ExportDetails

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
            }

            $script:tenantStatsHash["ConditionalAccessPolicies"][$policy.DisplayName] = $policyDetailsHash
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
                    $secureScoreResponse = Invoke-MgGraphRequest -Uri $secureScoreUri -Method GET -OutputType PSObject -ErrorAction Stop
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

function Get-EntraIDGroups {
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$false,HelpMessage='Provide the Graph Authentication Type')]
        [ValidateSet('SDK','REST')]
        [string[]]$GraphAuthType
    )
    $groupFetchProgressId = 61
    $groupLoopProgressId = 62
    $groupDetailProgressId = 63
    $depthPolicy = if ($script:CollectionDepthPolicy) {
        $script:CollectionDepthPolicy
    } else {
        Get-ArrayaCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($detailLevel.ToLowerInvariant()))
    }
    $collectDeepGroupDetails = ($depthPolicy.CollectEntraGroupDeepDetails -eq $true)
    $collectGroupLicenseChecks = ($depthPolicy.CollectEntraGroupLicenseChecks -eq $true)
    $collectGroupMemberCounts = ($depthPolicy.CollectEntraGroupMemberCounts -eq $true)
    $collectGroupOwnerCounts = ($depthPolicy.CollectEntraGroupOwnerCounts -eq $true)
    $groupMemberCountLookup = @{}
    $groupOwnerCountLookup = @{}
    $groupSelectProperties = @(
        'id',
        'displayName',
        'description',
        'visibility',
        'createdDateTime',
        'groupTypes',
        'mailEnabled',
        'securityEnabled',
        'onPremisesSyncEnabled',
        'onPremisesLastSyncDateTime',
        'isAssignableToRole',
        'mail',
        'membershipRule',
        'assignedLicenses'
    )

    function Convert-GraphBatchCountValue {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Body
        )

        if ($null -eq $Body) {
            return $null
        }

        if ($Body -is [int] -or $Body -is [long] -or $Body -is [double] -or $Body -is [decimal]) {
            try { return [int]$Body } catch { return $null }
        }

        if ($Body -is [string]) {
            $parsedValue = 0
            if ([int]::TryParse($Body, [ref]$parsedValue)) {
                return $parsedValue
            }
        }

        foreach ($propertyName in @('value', '@odata.count')) {
            if ($Body.PSObject.Properties[$propertyName]) {
                try { return [int]$Body.PSObject.Properties[$propertyName].Value } catch {}
            }
        }

        return $null
    }

    function Invoke-GraphBatchRequests {
        param(
            [Parameter(Mandatory = $true)]
            [array]$Requests,
            [Parameter(Mandatory = $false)]
            [string]$Activity = 'Graph batch request'
        )

        $result = @{}
        if (-not $Requests -or $Requests.Count -eq 0) {
            return $result
        }

        $batchEndpoint = 'https://graph.microsoft.com/v1.0/$batch'
        $chunkSize = 20

        for ($offset = 0; $offset -lt $Requests.Count; $offset += $chunkSize) {
            $chunk = @($Requests | Select-Object -Skip $offset -First $chunkSize)
            if ($chunk.Count -eq 0) {
                continue
            }

            $payload = @{ requests = $chunk } | ConvertTo-Json -Depth 10 -Compress
            $batchResponse = $null

            try {
                $canUseSdkBatch = (
                    ($GraphAuthType -contains 'SDK') -and
                    (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue) -and
                    (Get-MgContext -ErrorAction SilentlyContinue)
                )
                if ($canUseSdkBatch) {
                    $batchResponse = Invoke-QuietCommand -ScriptBlock {
                        Invoke-MgGraphRequest -Method POST -Uri $batchEndpoint -Body $payload -OutputType PSObject -ErrorAction Stop
                    }
                }
                else {
                    $headers = @{}
                    if ($global:GraphHeaders) {
                        $headers = $global:GraphHeaders.Clone()
                    }
                    elseif ($global:GraphToken) {
                        $headers = @{
                            'Content-Type'     = 'application/json'
                            'Authorization'    = "Bearer $global:GraphToken"
                            'ConsistencyLevel' = 'eventual'
                        }
                    }
                    if (-not $headers.ContainsKey('Content-Type')) {
                        $headers['Content-Type'] = 'application/json'
                    }
                    if (-not $headers.ContainsKey('ConsistencyLevel')) {
                        $headers['ConsistencyLevel'] = 'eventual'
                    }

                    $batchResponse = Invoke-QuietRestMethod -Parameters @{
                        Uri         = $batchEndpoint
                        Headers     = $headers
                        Method      = 'POST'
                        ContentType = 'application/json'
                        Body        = $payload
                        ErrorAction = 'Stop'
                    }
                }
            }
            catch {
                Write-Log -Type WARNING -Message "[Get-EntraIDGroups] $Activity failed for request chunk starting at index $offset. $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                continue
            }

            foreach ($response in @($batchResponse.responses)) {
                if ($response -and $response.id) {
                    $result[[string]$response.id] = $response
                }
            }
        }

        return $result
    }

    function Get-EntraGroupCountLookups {
        param(
            [Parameter(Mandatory = $true)]
            [array]$Groups,
            [Parameter(Mandatory = $false)]
            [switch]$IncludeMemberCounts,
            [Parameter(Mandatory = $false)]
            [switch]$IncludeOwnerCounts
        )

        $lookups = @{
            Members = @{}
            Owners  = @{}
        }

        if (-not $Groups -or $Groups.Count -eq 0) {
            return $lookups
        }
        if (-not $IncludeMemberCounts -and -not $IncludeOwnerCounts) {
            return $lookups
        }

        $requests = New-Object System.Collections.Generic.List[object]
        foreach ($group in $Groups) {
            $groupId = if ($group -and $group.PSObject.Properties['id']) { [string]$group.id } else { $null }
            if ([string]::IsNullOrWhiteSpace($groupId)) {
                continue
            }

            $isDynamicDistributionGroup = ($group.groupTypes -contains "DynamicMembership") -and ($group.mailEnabled -eq $true) -and ($group.securityEnabled -eq $false) -and (-not ($group.groupTypes -contains "Unified"))
            if ($isDynamicDistributionGroup) {
                continue
            }

            if ($IncludeMemberCounts) {
                $requests.Add([PSCustomObject]@{
                    id      = "m:$groupId"
                    method  = 'GET'
                    url     = "/groups/$groupId/members/`$count"
                    headers = @{ ConsistencyLevel = 'eventual' }
                }) | Out-Null
            }

            if ($IncludeOwnerCounts) {
                $requests.Add([PSCustomObject]@{
                    id      = "o:$groupId"
                    method  = 'GET'
                    url     = "/groups/$groupId/owners/`$count"
                    headers = @{ ConsistencyLevel = 'eventual' }
                }) | Out-Null
            }
        }

        if ($requests.Count -eq 0) {
            return $lookups
        }

        Write-Log -Type INFO -Message "[Get-EntraIDGroups] Prefetching group member/owner counts via Graph batch for $($Groups.Count) groups ($($requests.Count) count request(s))." -ExportFileLocation $ExportDetails
        $responses = Invoke-GraphBatchRequests -Requests $requests.ToArray() -Activity 'Group member/owner count prefetch'

        foreach ($response in @($responses.Values)) {
            if (-not $response -or -not $response.id) {
                continue
            }

            $responseId = [string]$response.id
            $statusCode = 0
            try { $statusCode = [int]$response.status } catch { $statusCode = 0 }

            if ($statusCode -lt 200 -or $statusCode -ge 300) {
                continue
            }

            $countValue = Convert-GraphBatchCountValue -Body $response.body
            if ($null -eq $countValue) {
                continue
            }

            if ($responseId.StartsWith('m:')) {
                $groupId = $responseId.Substring(2)
                $lookups.Members[$groupId] = [int]$countValue
            }
            elseif ($responseId.StartsWith('o:')) {
                $groupId = $responseId.Substring(2)
                $lookups.Owners[$groupId] = [int]$countValue
            }
        }

        Write-Log -Type INFO -Message "[Get-EntraIDGroups] Group count prefetch complete. MemberCounts=$($lookups.Members.Count) OwnerCounts=$($lookups.Owners.Count)" -ExportFileLocation $ExportDetails
        return $lookups
    }

    # Function to get group details by ID or DisplayName
    function Get-EntraGroupDetails {
        param (
            [Parameter(Mandatory = $true)]
            [object]$GroupIdentifier,  # Can be Group object, Group ID, or DisplayName
            [Parameter(Mandatory=$false,HelpMessage='Provide the Graph Authentication Type')]
            [ValidateSet('SDK','REST')]
            [string[]]$GraphAuthType
        )
        switch ($GraphAuthType) {
            REST { 
                $groupLookupKey = [string]$GroupIdentifier
                if ($GroupIdentifier -isnot [string]) {
                    $GroupDetails = $GroupIdentifier
                    if ($GroupDetails.PSObject.Properties['id'] -and $GroupDetails.id) {
                        $groupLookupKey = [string]$GroupDetails.id
                    }
                }
                else {
                    # Construct the API URL to fetch group details by ID or DisplayName
                    if ($groupLookupKey -match "^[0-9a-fA-F-]{36}$") {
                        $uri = "https://graph.microsoft.com/v1.0/groups/$groupLookupKey"
                    } else {
                        $uri = "https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '$groupLookupKey'"
                    }

                    # Fetch the group details
                    try {
                        Write-Log -Type INFO -Message "Fetching group details for $groupLookupKey" -ExportFileLocation $ExportDetails
                        $GroupDetails = Get-GraphData -PageSize 999 -URI $uri -ID $groupDetailProgressId -Activity "Gathering Group Details"
                        Write-Log -Type INFO -Message "Group details fetched successfully for $($GroupDetails.displayName)" -ExportFileLocation $ExportDetails
                    }
                    catch {
                        Write-Log -Type ERROR -Message "Error fetching group details for $($GroupDetails.displayName): $($_)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                        return
                    }
                }

                $isManagingLicenses = $false
                if ($collectGroupLicenseChecks) {
                    try {
                        $assignedLicenses = @()
                        if ($GroupDetails.PSObject.Properties['assignedLicenses'] -and $GroupDetails.assignedLicenses) {
                            $assignedLicenses = @($GroupDetails.assignedLicenses)
                        }
                        elseif ($GroupDetails.id) {
                            Write-Log -Type INFO -Message "Checking license details for group $($GroupDetails.displayName)" -ExportFileLocation $ExportDetails
                            $licenseUri = "https://graph.microsoft.com/v1.0/groups/$($GroupDetails.id)?`$select=assignedLicenses"
                            $licenseDetails = Get-GraphData -PageSize 999 -ID $groupDetailProgressId -URI $licenseUri -Activity "Gathering License Details"
                            if ($licenseDetails -and $licenseDetails.PSObject.Properties['assignedLicenses']) {
                                $assignedLicenses = @($licenseDetails.assignedLicenses)
                            }
                        }

                        if ($assignedLicenses.Count -gt 0) {
                            Write-Log -Type INFO -Message "Group $($GroupDetails.displayName) is managing licenses" -ExportFileLocation $ExportDetails
                            $isManagingLicenses = $true
                        }
                    } catch {
                        Write-Log -Type ERROR -Message "Error checking license details for group $($GroupDetails.displayName): $($_)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                    }
                }

                <# Not Working - Other Group Details - REQUIRES WRITE PERMISSIONS
                https://learn.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0#properties
                $otherDetailsURI = "https://graph.microsoft.com/v1.0/groups/$($GroupDetails.id)?`$select=hideFromAddressLists"
                $otherDetails = Invoke-RestMethod -Uri $otherDetailsURI -Headers $global:GraphHeaders -Method Get -ContentType "application/json"
                #>
                
                # If search was by DisplayName, get the first result
                if (($GroupIdentifier -is [string]) -and $groupLookupKey -notmatch "^[0-9a-fA-F-]{36}$") {
                    $GroupDetails = $GroupDetails | Select-Object -First 1
                    #$GroupDetails = $GroupDetails.value | Select-Object -First 1
                }
                $classification = Get-ArrayaEntraGroupClassification -GroupDetails $GroupDetails
                $MembershipType = $classification.MembershipType
                $MembershipRule = $classification.MembershipRule
                Write-Log -Type INFO -Message "Group Type: $MembershipType" -ExportFileLocation $ExportDetails
                $GroupType = $classification.GroupType
                Write-Log -Type INFO -Message "Group Type: $GroupType" -ExportFileLocation $ExportDetails

                # Count members and owners using Graph $count endpoints
                $MemberCount = 0
                $OwnerCount = 0
                $hasNestedMembers = "Skipped"
                $groupId = if ($GroupDetails.PSObject.Properties['id']) { [string]$GroupDetails.id } else { $null }
                $isDynamicDistributionGroup = ($GroupDetails.groupTypes -contains "DynamicMembership") -and ($GroupDetails.mailEnabled -eq $true) -and ($GroupDetails.securityEnabled -eq $false) -and (-not ($GroupDetails.groupTypes -contains "Unified"))
                if ($isDynamicDistributionGroup) {
                    $MemberCount = "Skipped (Dynamic Distribution Group)"
                    $OwnerCount = "Skipped (Dynamic Distribution Group)"
                    Write-Log -Type INFO -Message "Skipping member/owner counts for dynamic distribution group $($GroupDetails.displayName)" -ExportFileLocation $ExportDetails
                } elseif ($collectGroupMemberCounts -or $collectGroupOwnerCounts) {
                    try {
                        Write-Log -Type INFO -Message "Checking member and owner count for group $($GroupDetails.displayName)" -ExportFileLocation $ExportDetails
                        if ($collectGroupMemberCounts) {
                            if ($groupId -and $groupMemberCountLookup.ContainsKey($groupId)) {
                                $MemberCount = [int]$groupMemberCountLookup[$groupId]
                            }
                            else {
                                $memberUri = "https://graph.microsoft.com/v1.0/groups/$($GroupDetails.id)/members/\$count"
                                $memberCountResult = Get-GraphData -URI $memberUri -ID $groupDetailProgressId -Activity "Counting Members"
                                $MemberCount = if ($memberCountResult -is [array]) { [int]($memberCountResult | Select-Object -First 1) } else { [int]$memberCountResult }
                            }
                        }
                        else {
                            $MemberCount = 'NotCollected (minimum mode)'
                        }

                        if ($collectGroupOwnerCounts) {
                            if ($groupId -and $groupOwnerCountLookup.ContainsKey($groupId)) {
                                $OwnerCount = [int]$groupOwnerCountLookup[$groupId]
                            }
                            else {
                                $ownerUri = "https://graph.microsoft.com/v1.0/groups/$($GroupDetails.id)/owners/\$count"
                                $ownerCountResult = Get-GraphData -URI $ownerUri -ID $groupDetailProgressId -Activity "Counting Owners"
                                $OwnerCount = if ($ownerCountResult -is [array]) { [int]($ownerCountResult | Select-Object -First 1) } else { [int]$ownerCountResult }
                            }
                        }
                        else {
                            $OwnerCount = 'NotCollected (minimum mode)'
                        }
                    } catch {
                        Write-Log -Type ERROR -Message "Error counting members/owners for group $($GroupDetails.displayName): $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                    }
                } else {
                    $MemberCount = 'NotCollected (minimum mode)'
                    $OwnerCount = 'NotCollected (minimum mode)'
                }
                Write-Log -Type INFO -Message "Member Count: $MemberCount" -ExportFileLocation $ExportDetails
                Write-Log -Type INFO -Message "Owner Count: $OwnerCount" -ExportFileLocation $ExportDetails

                # Source of the group
                $Source = $classification.Source
                Write-Log -Type INFO -Message "Group Source: $Source" -ExportFileLocation $ExportDetails
             }
            SDK { # TBD
                if ($GroupIdentifier -is [string]) {
                    $savedProgressPreference = $ProgressPreference
                    try {
                        $ProgressPreference = 'SilentlyContinue'
                        $GroupDetails = Get-MgGroup -GroupId $GroupIdentifier -ErrorAction Stop
                    }
                    finally {
                        $ProgressPreference = $savedProgressPreference
                    }
                } else {
                    $GroupDetails = $GroupIdentifier
                }
                $classification = Get-ArrayaEntraGroupClassification -GroupDetails $GroupDetails
                $MembershipType = $classification.MembershipType
                $MembershipRule = $classification.MembershipRule
                $GroupType = $classification.GroupType
                $Source = $classification.Source
                $assignedLicenses = @()
                if ($GroupDetails.PSObject.Properties['assignedLicenses'] -and $GroupDetails.assignedLicenses) {
                    $assignedLicenses = @($GroupDetails.assignedLicenses)
                }
                $isManagingLicenses = if ($collectGroupLicenseChecks) { ($assignedLicenses.Count -gt 0) } else { 'NotCollected (minimum mode)' }
                $groupId = if ($GroupDetails.PSObject.Properties['id']) { [string]$GroupDetails.id } else { $null }
                $isDynamicDistributionGroup = ($GroupDetails.groupTypes -contains "DynamicMembership") -and ($GroupDetails.mailEnabled -eq $true) -and ($GroupDetails.securityEnabled -eq $false) -and (-not ($GroupDetails.groupTypes -contains "Unified"))
                if ($isDynamicDistributionGroup) {
                    $MemberCount = "Skipped (Dynamic Distribution Group)"
                    $OwnerCount = "Skipped (Dynamic Distribution Group)"
                }
                else {
                    $MemberCount = if ($collectGroupMemberCounts) {
                        if ($groupId -and $groupMemberCountLookup.ContainsKey($groupId)) { [int]$groupMemberCountLookup[$groupId] } else { 0 }
                    } else { 'NotCollected (minimum mode)' }
                    $OwnerCount = if ($collectGroupOwnerCounts) {
                        if ($groupId -and $groupOwnerCountLookup.ContainsKey($groupId)) { [int]$groupOwnerCountLookup[$groupId] } else { 0 }
                    } else { 'NotCollected (minimum mode)' }
                }
                $hasNestedMembers = $false
            }
        
        }

        # Output the group details in a structured object
        $result = [PSCustomObject]@{
            # ===== Identity & Description =====
            ID              = $GroupDetails.id
            DisplayName     = $GroupDetails.displayName
            Description     = $GroupDetails.description
            Source          = $Source
            Visibility      = $GroupDetails.visibility
            CreatedDateTime = $GroupDetails.createdDateTime
    
            # ===== Group Type & Membership =====
            GroupType        = $GroupType
            MembershipType   = $MembershipType
            MembershipRule   = $MembershipRule
            HasNestedMembers = $hasNestedMembers
    
            # ===== Directory Sync & Licensing =====
            OnPremisesSyncEnabled      = $GroupDetails.onPremisesSyncEnabled
            OnPremisesLastSyncDateTime = $GroupDetails.onPremisesLastSyncDateTime
            IsManagingLicenses         = $isManagingLicenses
            IsAssignableToRole         = $GroupDetails.isAssignableToRole
    
            # ===== Mail & Security Settings =====
            Mail            = $GroupDetails.mail
            MailEnabled     = $GroupDetails.mailEnabled
            SecurityEnabled = $GroupDetails.securityEnabled
    
            # ===== Group Membership Stats =====
            MemberCount = $MemberCount
            OwnerCount  = $OwnerCount
        }
        # Return the group details
        return $result
    }

    # Initialize hash table for storing group details
    $script:tenantStatsHash["EntraIDGroups"] = @{}
    $totalGroups = 0

    try {
        if ($GraphAuthType -contains 'SDK') {
            $savedProgressPreference = $ProgressPreference
            try {
                $ProgressPreference = 'SilentlyContinue'
                $Groups = @(
                    Invoke-QuietCommand -ScriptBlock {
                        Get-MgGroup -All -Property $groupSelectProperties -ErrorAction Stop
                    }
                )
                $sdkGroupTotal = $Groups.Count
                if ($collectDeepGroupDetails -and ($collectGroupMemberCounts -or $collectGroupOwnerCounts) -and $sdkGroupTotal -gt 0) {
                    $countLookups = Get-EntraGroupCountLookups -Groups $Groups -IncludeMemberCounts:$collectGroupMemberCounts -IncludeOwnerCounts:$collectGroupOwnerCounts
                    $groupMemberCountLookup = $countLookups.Members
                    $groupOwnerCountLookup = $countLookups.Owners
                }

                $processedGroups = 0
                foreach ($Group in $Groups) {
                    $processedGroups++
                    Write-Progress -Id $groupLoopProgressId -Activity "Getting Group Details (SDK)" -Status "Processed $processedGroups group(s): $($Group.DisplayName)"

                    Write-Log -Type INFO -Message "Checking group $($Group.displayName)" -ExportFileLocation $ExportDetails
                    $GroupDetails = Invoke-ArrayaCollectionStepSafe -OperationName "Get-EntraIDGroups details for $($Group.displayName)" -DefaultValue $null -ExportFileLocation $ExportDetails -ScriptBlock {
                        Get-EntraGroupDetails -GroupIdentifier $Group -GraphAuthType $GraphAuthType
                    }
                    
                    if ($GroupDetails) {
                        $script:tenantStatsHash['EntraIDGroups'][$GroupDetails.ID] = $GroupDetails
                    } else {
                        Write-Log -Type WARNING -Message "[Get-EntraIDGroups] Could not retrieve details for group $($Group.displayName). Continuing." -ExportFileLocation $ExportDetails
                    }
                }
            }
            finally {
                $ProgressPreference = $savedProgressPreference
            }
        }
        else {
            $Groups = @()

            # Endpoint and initial URL for group data
            $GroupsEndpoint = "https://graph.microsoft.com/v1.0/groups?`$select=$($groupSelectProperties -join ',')"

            # Loop to handle paging through all group data
            try {
                Write-Log -Type INFO -Message "Fetching initial Entra Groups" -ExportFileLocation $ExportDetails
                $Groups = Get-GraphData -PageSize 999 -URI $GroupsEndpoint -ID $groupFetchProgressId -Activity "Gathering Group Details"
                Write-Progress -Activity "Getting Group Details" -Completed -Id $groupFetchProgressId
            }
            catch {
                Write-Log -Type ERROR -Message "Error fetching initial Entra Groups: $($_)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                return @()
            }

            # Process the groups data
            $totalGroups = $Groups.Count
            if ($collectDeepGroupDetails -and ($collectGroupMemberCounts -or $collectGroupOwnerCounts) -and $totalGroups -gt 0) {
                $countLookups = Get-EntraGroupCountLookups -Groups $Groups -IncludeMemberCounts:$collectGroupMemberCounts -IncludeOwnerCounts:$collectGroupOwnerCounts
                $groupMemberCountLookup = $countLookups.Members
                $groupOwnerCountLookup = $countLookups.Owners
            }
            foreach ($Group in $Groups) {
                Write-ProgressHelper -Total $totalGroups -Id $groupLoopProgressId -Activity "Getting Group Details" -Operation "Processing Group: $($Group.displayName)"

                Write-Log -Type INFO -Message "Checking group $($Group.displayName)" -ExportFileLocation $ExportDetails
                $GroupDetails = Invoke-ArrayaCollectionStepSafe -OperationName "Get-EntraIDGroups details for $($Group.displayName)" -DefaultValue $null -ExportFileLocation $ExportDetails -ScriptBlock {
                    Get-EntraGroupDetails -GroupIdentifier $Group -GraphAuthType $GraphAuthType
                }
                
                if ($GroupDetails) {
                    $script:tenantStatsHash['EntraIDGroups'][$GroupDetails.ID] = $GroupDetails
                } else {
                    Write-Log -Type WARNING -Message "[Get-EntraIDGroups] Could not retrieve details for group $($Group.displayName). Continuing." -ExportFileLocation $ExportDetails
                }
            }
        }
    }
    finally {
        Write-Progress -Activity "Getting Group Details" -Completed -Id $groupFetchProgressId
        Write-Progress -Activity "Getting Group Details" -Completed -Id $groupLoopProgressId
        Write-ProgressHelper -Total 1 -Id $groupDetailProgressId -Activity "Gathering Group Details" -Completed
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
    
    Write-Host "Checking Authentication and SSO Configuration ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] START: Checking Authentication Configuration" -ExportFileLocation $ExportDetails
    
    try {
        $depthPolicy = if ($script:CollectionDepthPolicy) {
            $script:CollectionDepthPolicy
        } else {
            Get-ArrayaCollectionDepthPolicy -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($detailLevel.ToLowerInvariant()))
        }
        $collectSsoAppDetails = ($depthPolicy.CollectSsoApplicationDetails -eq $true)

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
                foreach ($app in (Get-MgServicePrincipal -All -Filter "tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')" -ErrorAction SilentlyContinue)) {
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
            $authPolicyData = Get-GraphData -Uri $authMethodsPolicyUri -Activity "Fetching Authentication Methods Policy"
            
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
            $authorizationPolicyResponse = Get-GraphData -Uri "https://graph.microsoft.com/v1.0/policies/authorizationPolicy" -Activity "Fetching authorization policy"
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
            $adminConsentPolicyResponse = Get-GraphData -Uri "https://graph.microsoft.com/v1.0/policies/adminConsentRequestPolicy" -Activity "Fetching admin consent request policy"
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
                $regData = Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop
            }
        } catch {}
        
        if (-not $regData -or $regData.Count -eq 0) {
            try {
                $uri = "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails"
                $regData = Get-GraphData -Uri $uri -Activity "Fetching MFA Registration Details"
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
            $crossTenantPolicy = Get-MgPolicyCrossTenantAccessPolicy -ErrorAction Stop
        } catch {
            Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] Unable to retrieve CrossTenantAccessPolicy: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        try {
            $crossTenantPartners = Get-MgPolicyCrossTenantAccessPolicyPartner -All -ErrorAction Stop
        } catch {
            Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] Unable to retrieve CrossTenantAccessPolicy partners: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        try {
            $b2bPolicy = Get-MgPolicyB2BManagementPolicy -ErrorAction Stop
        } catch {
            if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
                try {
                    $b2bPolicy = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/b2bManagementPolicy' -ErrorAction Stop
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
                    $lookup = Find-MgTenantRelationshipTenantInformationByTenantId -TenantId $TenantId -ErrorAction Stop
                    if ($lookup.DisplayName) { return $lookup.DisplayName }
                }
                if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
                    $lookup = Get-GraphData -Uri $tenantLookupUri -Activity "Tenant lookup (v1.0)"
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
                    $lookup = Get-GraphData -Uri $tenantLookupUri -Activity "Tenant lookup (beta)"
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

    $snapshot = Import-ArrayaTenantSnapshot -Path $Path
    $script:LoadedTenantSnapshot = $snapshot

    $validation = Test-ArrayaTenantSnapshot -Snapshot $snapshot -Purpose Export
    if (-not $validation.Valid) {
        throw ("Tenant snapshot is invalid: {0}" -f ($validation.Errors -join '; '))
    }
    if ($validation.Warnings.Count -gt 0) {
        Write-Warning ("Tenant snapshot imported with warnings: {0}" -f ($validation.Warnings -join '; '))
    }

    return Convert-ArrayaSnapshotToLegacyTenantStatsHash -Snapshot $snapshot
}

#region HTML Report Helpers
if ($true) {
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

#region Helper Functions

function Get-FromHash {
    <#
    .SYNOPSIS
        Safely retrieves data from hashtable with fallback to empty array.
    #>
    param(
        [Parameter(Mandatory)]
        [hashtable]$Hashtable,
        
        [Parameter(Mandatory)]
        [string]$Key
    )
    
    if ($Hashtable.ContainsKey($Key)) {
        $Hashtable[$Key]
    } else {
        Write-Warning "Key '$Key' not found in hashtable. Section may be incomplete."
        @()
    }
}

function ConvertTo-Array {
    <#
    .SYNOPSIS
        Ensures input is converted to array format. Handles hashtables by extracting values.
    #>
    param($InputObject)

    if (Get-Command -Name Convert-ArrayaObjectToArray -ErrorAction SilentlyContinue) {
        return Convert-ArrayaObjectToArray -InputObject $InputObject
    }

    if ($null -eq $InputObject) {
        return @()
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        return @($InputObject.Values)
    }

    if ($InputObject -is [string]) {
        return @($InputObject)
    }

    if ($InputObject -is [System.Collections.IEnumerable]) {
        return @($InputObject)
    }

    return @($InputObject)
}

function Test-IsBlankDisplayValue {
    <#
    .SYNOPSIS
        Checks whether a value should be treated as blank for HTML display.
    #>
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return $true
    }

    if ($Value -is [string]) {
        return [string]::IsNullOrWhiteSpace($Value)
    }

    return $false
}

function Format-Number {
    <#
    .SYNOPSIS
        Formats numbers with thousand separators and optional decimal places.
    #>
    param(
        [AllowNull()]
        $Number,
        
        [int]$DecimalPlaces = 0
    )
    
    if (Test-IsBlankDisplayValue -Value $Number) { return 'N/A' }
    
    try {
        $num = [double]$Number
        if ($DecimalPlaces -eq 0) {
            return $num.ToString('N0')
        } else {
            return $num.ToString("N$DecimalPlaces")
        }
    } catch {
        return $Number.ToString()
    }
}

function Format-DataSize {
    param(
        [AllowNull()]
        [double]$SizeInGB = 0
    )

    if ($null -eq $SizeInGB) {
        return 'N/A'
    }

    if ($SizeInGB -lt 0) {
        return 'N/A'
    }

    if ($SizeInGB -ge 1048576) {
        $sizeInPB = $SizeInGB / 1048576
        return "$(Format-Number $sizeInPB -DecimalPlaces 2) PB"
    }

    if ($SizeInGB -ge 1024) {
        $sizeInTB = $SizeInGB / 1024
        return "$(Format-Number $sizeInTB -DecimalPlaces 2) TB"
    }

    if ($SizeInGB -ge 1) {
        return "$(Format-Number $SizeInGB -DecimalPlaces 2) GB"
    }

    return "$(Format-Number ($SizeInGB * 1024) -DecimalPlaces 2) MB"
}

function Format-Percentage {
    <#
    .SYNOPSIS
        Formats a value as percentage.
    #>
    param(
        [AllowNull()]
        $Value,
        
        [int]$DecimalPlaces = 1
    )
    
    if (Test-IsBlankDisplayValue -Value $Value) { return 'N/A' }
    
    try {
        $num = [double]$Value
        return $num.ToString("P$DecimalPlaces")
    } catch {
        return $Value.ToString()
    }
}

function Get-RiskBadge {
    <#
    .SYNOPSIS
        Returns HTML for risk badge based on severity.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Info', 'Warning', 'Risk', 'Success')]
        [string]$Type,
        
        [Parameter(Mandatory)]
        [string]$Text
    )
    
    $iconMap = @{
        'Info' = 'ℹ️'
        'Warning' = '⚠️'
        'Risk' = '🔴'
        'Success' = '✅'
    }
    
    $icon = $iconMap[$Type]
    $class = $Type.ToLower()
    
    return "<span class='badge badge-$class'>$icon $Text</span>"
}

function New-HtmlTable {
    <#
    .SYNOPSIS
        Generates HTML table from array of objects with optional risk highlighting.
    #>
    param(
        [AllowNull()]
        [array]$Data,
        
        [Parameter(Mandatory)]
        [string[]]$Columns,
        
        [hashtable]$ColumnHeaders,
        
        [hashtable]$RiskColumns,

        [hashtable]$ValueFormatters,
        
        [string]$EmptyMessage = 'No data available',
        
        [string]$CssClass = 'data-table'
    )
    
    if ($null -eq $Data -or $Data.Count -eq 0) {
        return "<div class='empty-state'>$EmptyMessage</div>"
    }
    
    $html = "<table class='$CssClass'>`n<thead><tr>"
    
    # Headers
    foreach ($col in $Columns) {
        $header = if ($ColumnHeaders -and $ColumnHeaders.ContainsKey($col)) {
            $ColumnHeaders[$col]
        } else {
            $col
        }
        $html += "<th>$header</th>"
    }
    $html += "</tr></thead>`n<tbody>"
    
    # Rows
    foreach ($row in $Data) {
        $html += "<tr>"
        foreach ($col in $Columns) {
            $value = $row.$col
            
            # Format value
            if ($ValueFormatters -and $ValueFormatters.ContainsKey($col)) {
                $formatter = $ValueFormatters[$col]
                if ($formatter -is [scriptblock]) {
                    $formattedValue = & $formatter $value $row
                    $displayValue = [System.Web.HttpUtility]::HtmlEncode([string]$formattedValue)
                }
                else {
                    $displayValue = [System.Web.HttpUtility]::HtmlEncode([string]$value)
                }
            } elseif (Test-IsBlankDisplayValue -Value $value) {
                $displayValue = 'N/A'
            } elseif ($value -is [datetime]) {
                $displayValue = $value.ToString('yyyy-MM-dd')
            } elseif ($value -is [bool]) {
                $displayValue = if ($value) { '✓' } else { '✗' }
            } elseif ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                $listValues = @($value | ForEach-Object {
                    if (-not (Test-IsBlankDisplayValue -Value $_)) {
                        [string]$_
                    }
                })
                $displayValue = if ($listValues.Count -gt 0) {
                    [System.Web.HttpUtility]::HtmlEncode(($listValues -join ', '))
                } else {
                    'N/A'
                }
            } else {
                $displayValue = [System.Web.HttpUtility]::HtmlEncode($value.ToString())
            }
            
            # Check for risk highlighting
            $cellClass = ''
            if ($RiskColumns -and $RiskColumns.ContainsKey($col)) {
                $riskCheck = $RiskColumns[$col]
                if ($riskCheck -is [scriptblock]) {
                    if (& $riskCheck $value $row) {
                        $cellClass = ' class="risk-cell"'
                    }
                }
            }
            
            $html += "<td$cellClass>$displayValue</td>"
        }
        $html += "</tr>`n"
    }
    
    $html += "</tbody></table>"
    return $html
}

function New-KpiCard {
    <#
    .SYNOPSIS
        Creates a KPI card with value and optional change indicator.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        
        [Parameter(Mandatory)]
        [object]$Value,
        
        [string]$Subtitle,
        
        [ValidateSet('', 'up', 'down', 'neutral')]
        [string]$Trend = '',
        
        [ValidateSet('default', 'success', 'warning', 'danger')]
        [string]$Theme = 'default'
    )
    
    $trendIcon = switch ($Trend) {
        'up' { '↗️' }
        'down' { '↘️' }
        'neutral' { '→' }
        default { '' }
    }
    
    $subtitleHtml = if ($Subtitle) {
        "<div class='kpi-subtitle'>$Subtitle</div>"
    } else { '' }

    $displayValue = if ($null -eq $Value) {
        'N/A'
    } elseif ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        (($Value | ForEach-Object { $_.ToString() }) -join ', ')
    } else {
        $Value.ToString()
    }
    
    return @"
<div class="kpi-card kpi-$Theme">
    <div class="kpi-title">$Title</div>
    <div class="kpi-value">$displayValue $trendIcon</div>
    $subtitleHtml
</div>
"@
}

function New-CalloutBox {
    <#
    .SYNOPSIS
        Creates a callout/alert box.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('info', 'warning', 'danger', 'success')]
        [string]$Type,
        
        [Parameter(Mandatory)]
        [string]$Title,
        
        [Parameter(Mandatory)]
        [string]$Content
    )
    
    $iconMap = @{
        'info' = 'ℹ️'
        'warning' = '⚠️'
        'danger' = '🔴'
        'success' = '✅'
    }
    
    return @"
<div class="callout callout-$Type">
    <div class="callout-header">$($iconMap[$Type]) $Title</div>
    <div class="callout-body">$Content</div>
</div>
"@
}

#endregion

#region Analysis Functions

function Get-LicenseAnalysis {
    <#
    .SYNOPSIS
        Analyzes license data with refined severity levels
        
    .DESCRIPTION
        Critical: Over capacity (negative remaining)
        Warning: At capacity (0 remaining) or high utilization (>85% with licenses available)
        Info: Overall utilization summary
    #>
    param(
        [array]$Licenses,
        [int]$UserCount = 0
    )
    
    $findings = @()
    $criticalFindings = @()
    $warningFindings = @()
    $infoFindings = @()
    
    # Process licenses with consistent paid/free classification
    $processedLicenses = $Licenses | ForEach-Object {
        Get-LicenseInventoryRecord -License $_ -UserCount $UserCount
    }
    
    # Filter paid licenses
    $paidLicenses = $processedLicenses | Where-Object { $_.IsPaid -eq $true }
    
    # Analyze findings with NEW logic
    foreach ($lic in $paidLicenses) {
        
        # 🔴 CRITICAL: Over capacity (negative remaining)
        $licenseName = if ($lic.SkuFriendlyName) { $lic.SkuFriendlyName } else { $lic.SkuPartNumber }
        if ($lic.RemainingUnits -lt 0 -and $lic.PurchasedUnits -gt 0) {
            $criticalFindings += @{
                Type = 'Risk'
                Category = 'Critical'
                Message = "License '$licenseName' is OVER capacity by $([math]::Abs($lic.RemainingUnits)) licenses"
                Anchor = 'licenses'
                Priority = 1
            }
        }
        # ⚠️ WARNING: At capacity (0 remaining)
        elseif ($lic.RemainingUnits -eq 0 -and $lic.PurchasedUnits -gt 0) {
            $warningFindings += @{
                Type = 'Warning'
                Category = 'At Capacity'
                Message = "License '$licenseName' has 0 remaining (fully allocated)"
                Anchor = 'licenses'
                Priority = 2
            }
        }
        # ⚠️ WARNING: High utilization but has some remaining
        elseif ($lic.Utilization -ge 85 -and $lic.RemainingUnits -gt 0) {
            $warningFindings += @{
                Type = 'Warning'
                Category = 'High Utilization'
                Message = "License '$licenseName' is $([math]::Round($lic.Utilization,1))% utilized ($($lic.RemainingUnits) remaining)"
                Anchor = 'licenses'
                Priority = 3
            }
        }
    }
    
    # Overall utilization (info level)
    $totalPurchased = ($paidLicenses | Measure-Object -Property PurchasedUnits -Sum).Sum
    $totalConsumed = ($paidLicenses | Measure-Object -Property ConsumedUnits -Sum).Sum
    
    if ($totalPurchased -gt 0) {
        $overallUtil = ($totalConsumed / $totalPurchased) * 100
        if ($overallUtil -ge 85) {
            $infoFindings += @{
                Type = 'Warning'
                Category = 'Overall Utilization'
                Message = "Overall license utilization is $([math]::Round($overallUtil,1))% across all paid licenses"
                Anchor = 'licenses'
                Priority = 4
            }
        }
    }
    
    # Combine findings in priority order
    $findings = $criticalFindings + $warningFindings + $infoFindings
    
    return @{
        Findings = $findings
        TotalPurchased = $totalPurchased
        TotalConsumed = $totalConsumed
        PaidLicenses = $paidLicenses
        CriticalCount = $criticalFindings.Count
        WarningCount = $warningFindings.Count
    }
}

function Get-MailboxAnalysis {
    <#
    .SYNOPSIS
        Analyzes mailbox data and returns findings
    #>
    param([array]$Mailboxes)
    
    $findings = @()
    $largeMailboxes = @()
    $largeArchiveCount = 0
    $largestArchiveMailbox = $null
    
    foreach ($mbx in $Mailboxes) {
        # Handle N/A values safely
        $mbxSize = 0
        $archiveSize = 0
        
        if ($mbx.MBXSizeGB -ne 'N/A' -and $null -ne $mbx.MBXSizeGB -and $mbx.MBXSizeGB -ne '') {
            try { $mbxSize = [double]$mbx.MBXSizeGB } catch { $mbxSize = 0 }
        }
        
        if ($mbx.ArchiveSizeGB -ne 'N/A' -and $null -ne $mbx.ArchiveSizeGB -and $mbx.ArchiveSizeGB -ne '') {
            try { $archiveSize = [double]$mbx.ArchiveSizeGB } catch { $archiveSize = 0 }
        }
        
        # Large mailbox warnings
        if ($mbxSize -gt $script:DefaultThresholds.MailboxSizeGB) {
            $largeMailboxes += $mbx
        }
        
        # Large archive warnings
        if ($archiveSize -gt $script:DefaultThresholds.ArchiveSizeGB) {
            $largeArchiveCount++
            if (-not $largestArchiveMailbox -or $archiveSize -gt $largestArchiveMailbox.SizeGB) {
                $largestArchiveMailbox = [PSCustomObject]@{
                    DisplayName = [string]$mbx.DisplayName
                    SizeGB = [math]::Round($archiveSize, 1)
                }
            }
        }
    }
    
    # Summary finding for large mailboxes
    if ($largeMailboxes.Count -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large Mailboxes'
            Message = "$($largeMailboxes.Count) mailboxes exceed $($script:DefaultThresholds.MailboxSizeGB) GB"
            Anchor = 'mailboxes'
            Priority = 2
        }
    }

    if ($largeArchiveCount -gt 0) {
        $largestArchiveLabel = if ($largestArchiveMailbox -and -not [string]::IsNullOrWhiteSpace($largestArchiveMailbox.DisplayName)) {
            "'$($largestArchiveMailbox.DisplayName)' at $($largestArchiveMailbox.SizeGB) GB"
        } else {
            'unavailable'
        }
        $findings += @{
            Type = 'Warning'
            Category = 'Large Archives'
            Message = "$largeArchiveCount archive mailbox(es) exceed $($script:DefaultThresholds.ArchiveSizeGB) GB (largest: $largestArchiveLabel)"
            Anchor = 'mailboxes'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
        LargeMailboxes = $largeMailboxes
        LargeArchiveCount = $largeArchiveCount
        LargestArchiveMailbox = $largestArchiveMailbox
    }
}

function Get-DomainAnalysis {
    <#
    .SYNOPSIS
        Analyzes domain configuration and returns findings
    #>
    param(
        [array]$Domains,
        [object]$SpamFilteringSummary,
        [object]$SMTPRelaySummary
    )
    
    $findings = @()
    $mailEnabledCustomDomainCount = 0
    $spfPassingDomainCount = 0
    $dmarcConfiguredDomainCount = 0
    $dmarcEnforcedDomainCount = 0
    $dkimCompleteDomainCount = 0
    
    foreach ($domain in $Domains) {
        $domainName = [string]$domain.Domain
        $isTenantServiceDomain = $domainName -match '(?i)\.onmicrosoft\.com$'
        $isCustomDomain = -not $isTenantServiceDomain
        $hasRecipientUsage = $false
        try {
            $hasRecipientUsage = ([int]$domain.TotalDomainRecipients -gt 0)
        } catch {}

        # Check verification
        if (-not $domain.Verified -and $script:DefaultThresholds.DomainVerificationRequired) {
            $findings += @{
                Type = 'Risk'
                Category = 'Domain Verification'
                Message = "Domain '$($domain.Domain)' is not verified"
                Anchor = 'domains'
                Priority = 1
            }
        }
        
        # Check MX records
        if ($script:DefaultThresholds.MXRecordValidation) {
            if ($domain.Office365MailExchanger -eq $false) {
                $findings += @{
                    Type = 'Info'
                    Category = 'DNS Configuration'
                    Message = "Domain '$($domain.Domain)' MX record does not point to Microsoft 365"
                    Anchor = 'domains-dns'
                    Priority = 3
                }
            }
        }

        if ($isCustomDomain -and $domain.Verified -and ($domain.Office365MailExchanger -eq $true -or $hasRecipientUsage)) {
            $mailEnabledCustomDomainCount++

            $spfIncludesM365 = $false
            if ($domain.PSObject.Properties['SpfIncludesM365']) {
                $spfIncludesM365 = ($domain.SpfIncludesM365 -eq $true)
            }
            if ($spfIncludesM365) { $spfPassingDomainCount++ }
            if ($domain.PSObject.Properties['SpfIncludesM365'] -and $domain.SpfIncludesM365 -ne $true) {
                $findings += @{
                    Type = 'Warning'
                    Category = 'SPF'
                    Message = "Domain '$($domain.Domain)' does not show an SPF record including spf.protection.outlook.com"
                    Anchor = 'domains-dns'
                    Priority = 2
                }
            }
            if ($domain.PSObject.Properties['SpfPolicyMode'] -and $domain.SpfPolicyMode) {
                if ([string]$domain.SpfPolicyMode -eq 'AllowAll (+all)') {
                    $findings += @{
                        Type = 'Risk'
                        Category = 'SPF'
                        Message = "Domain '$($domain.Domain)' SPF record is configured as +all (allow all), which weakens spoof protection"
                        Anchor = 'domains-dns'
                        Priority = 1
                    }
                }
                elseif ([string]$domain.SpfPolicyMode -eq 'SoftFail (~all)') {
                    $findings += @{
                        Type = 'Info'
                        Category = 'SPF'
                        Message = "Domain '$($domain.Domain)' uses SPF soft-fail (~all); consider hard-fail (-all) after validation"
                        Anchor = 'domains-dns'
                        Priority = 3
                    }
                }
            }

            $dmarcConfigured = $false
            if ($domain.PSObject.Properties['DmarcConfigured']) {
                $dmarcConfigured = ($domain.DmarcConfigured -eq $true)
            }
            if ($dmarcConfigured) { $dmarcConfiguredDomainCount++ }
            if ($domain.PSObject.Properties['DmarcConfigured'] -and $domain.DmarcConfigured -ne $true) {
                $findings += @{
                    Type = 'Warning'
                    Category = 'DMARC'
                    Message = "Domain '$($domain.Domain)' does not show a DMARC policy record"
                    Anchor = 'domains-dns'
                    Priority = 2
                }
            }
            if ($dmarcConfigured -and $domain.PSObject.Properties['DmarcPolicy']) {
                $dmarcPolicy = [string]$domain.DmarcPolicy
                if ($dmarcPolicy -eq 'none') {
                    $findings += @{
                        Type = 'Warning'
                        Category = 'DMARC'
                        Message = "Domain '$($domain.Domain)' DMARC policy is p=none (monitor only); move toward quarantine/reject for anti-spoofing enforcement"
                        Anchor = 'domains-dns'
                        Priority = 2
                    }
                } elseif ($dmarcPolicy -in @('quarantine', 'reject')) {
                    $dmarcEnforcedDomainCount++
                }

                $dmarcPercent = 100
                if ($domain.PSObject.Properties['DmarcPercent']) {
                    try { $dmarcPercent = [int]$domain.DmarcPercent } catch { $dmarcPercent = 100 }
                }
                if ($dmarcPercent -lt 100) {
                    $findings += @{
                        Type = 'Info'
                        Category = 'DMARC'
                        Message = "Domain '$($domain.Domain)' DMARC enforcement scope is pct=$dmarcPercent; increase toward 100 for full anti-spoofing coverage"
                        Anchor = 'domains-dns'
                        Priority = 3
                    }
                }
            }

            if ($domain.PSObject.Properties['DkimSelectorsConfigured']) {
                $dkimSelectorCount = 0
                try { $dkimSelectorCount = [int]$domain.DkimSelectorsConfigured } catch { $dkimSelectorCount = 0 }
                if ($dkimSelectorCount -ge 2) { $dkimCompleteDomainCount++ }
                if ($dkimSelectorCount -lt 2) {
                    $findings += @{
                        Type = 'Warning'
                        Category = 'DKIM'
                        Message = "Domain '$($domain.Domain)' has $dkimSelectorCount DKIM selector CNAME record(s) detected (expected: 2)"
                        Anchor = 'domains-dns'
                        Priority = 2
                    }
                }
            }
        }
    }

    if ($mailEnabledCustomDomainCount -gt 0) {
        $spfPct = [math]::Round((($spfPassingDomainCount / $mailEnabledCustomDomainCount) * 100), 1)
        $dmarcPct = [math]::Round((($dmarcConfiguredDomainCount / $mailEnabledCustomDomainCount) * 100), 1)
        $dmarcEnforcedPct = [math]::Round((($dmarcEnforcedDomainCount / $mailEnabledCustomDomainCount) * 100), 1)
        $dkimPct = [math]::Round((($dkimCompleteDomainCount / $mailEnabledCustomDomainCount) * 100), 1)

        $coverageType = if ($dmarcEnforcedPct -lt 60 -or $dkimPct -lt 60 -or $spfPct -lt 80) { 'Warning' } else { 'Info' }
        $coveragePriority = if ($coverageType -eq 'Warning') { 2 } else { 3 }
        $findings += @{
            Type = $coverageType
            Category = 'Email Authentication Coverage'
            Message = "Mail-auth coverage across $mailEnabledCustomDomainCount custom mail domain(s): SPF $spfPct%, DKIM $dkimPct%, DMARC configured $dmarcPct%, DMARC enforcement (quarantine/reject) $dmarcEnforcedPct%"
            Anchor = 'domains-dns'
            Priority = $coveragePriority
        }
    }

    if ($SpamFilteringSummary) {
        $trustedBypassCount = 0
        if ($SpamFilteringSummary.PSObject.Properties['TransportRulesWithTrustedIPs']) {
            try { $trustedBypassCount = [int]$SpamFilteringSummary.TransportRulesWithTrustedIPs } catch { $trustedBypassCount = 0 }
        }
        $uses3rdPartyFiltering = $null
        if ($SpamFilteringSummary.PSObject.Properties['Uses3rdPartyFiltering']) {
            $rawThirdParty = $SpamFilteringSummary.Uses3rdPartyFiltering
            if ($rawThirdParty -is [bool]) {
                $uses3rdPartyFiltering = $rawThirdParty
            } elseif ($null -ne $rawThirdParty) {
                $uses3rdPartyFiltering = ([string]$rawThirdParty -match '(?i)^(yes|true|enabled)$')
            }
        }
        if ($trustedBypassCount -gt 0) {
            $findings += @{
                Type = 'Warning'
                Category = 'Anti-Spoofing Bypass'
                Message = "$trustedBypassCount transport rule or connection-filter trusted IP bypass indicator(s) detected; validate spoof protection exceptions"
                Anchor = 'domains-dns'
                Priority = 2
            }
        }
        $findings += @{
            Type = 'Info'
            Category = 'Anti-Spoofing Controls'
            Message = "Mail-flow anti-spoof controls: trusted bypass indicators=$trustedBypassCount; third-party filtering detected=$(if ($null -eq $uses3rdPartyFiltering) { 'Unknown' } elseif ($uses3rdPartyFiltering) { 'Yes' } else { 'No' })"
            Anchor = 'domains-dns'
            Priority = 3
        }
    }

    if ($SMTPRelaySummary) {
        $smtpAuthEnabled = $false
        if ($SMTPRelaySummary.PSObject.Properties['SMTPAuthEnabled']) {
            $rawSmtpAuth = $SMTPRelaySummary.SMTPAuthEnabled
            if ($rawSmtpAuth -is [bool]) {
                $smtpAuthEnabled = $rawSmtpAuth
            } elseif ($null -ne $rawSmtpAuth) {
                $smtpAuthEnabled = ([string]$rawSmtpAuth -match '(?i)^(yes|true|enabled)$')
            }
        }

        $smtpAuthUsers = 0
        if ($SMTPRelaySummary.PSObject.Properties['SMTPAuthUsers']) {
            try { $smtpAuthUsers = [int]$SMTPRelaySummary.SMTPAuthUsers } catch { $smtpAuthUsers = 0 }
        }

        if ($smtpAuthEnabled -and $smtpAuthUsers -gt 0) {
            $findings += @{
                Type = 'Warning'
                Category = 'SMTP AUTH Exposure'
                Message = "$smtpAuthUsers mailbox(es) still have SMTP AUTH enabled; this can increase impersonation and brute-force attack surface"
                Anchor = 'domains-dns'
                Priority = 2
            }
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-DeviceAnalysis {
    <#
    .SYNOPSIS
        Analyzes device data and returns findings
    #>
    param([array]$Devices)
    
    $findings = @()
    
    if ($Devices.Count -eq 0) {
        return @{ Findings = @() }
    }
    
    # Stale devices
    $staleDevices = $Devices | Where-Object {
        $_.DeviceStale -eq $true
    }
    
    if ($staleDevices.Count -gt 0) {
        $stalePercent = ($staleDevices.Count / $Devices.Count) * 100
        $findings += @{
            Type = 'Warning'
            Category = 'Stale Devices'
            Message = "$($staleDevices.Count) devices ($([math]::Round($stalePercent,1))%) are stale (inactive > $($script:DefaultThresholds.DeviceStaleMonths) months)"
            Anchor = 'devices'
            Priority = 2
        }
    }
    
    # Device compliance
    $compliantDevices = $Devices | Where-Object { $_.IsCompliant -eq $true }
    $compliancePercent = ($compliantDevices.Count / $Devices.Count) * 100
    
    if ($compliancePercent -lt $script:DefaultThresholds.DeviceCompliancePercent) {
        $findings += @{
            Type = 'Risk'
            Category = 'Device Compliance'
            Message = "Device compliance is only $([math]::Round($compliancePercent,1))% (target: $($script:DefaultThresholds.DeviceCompliancePercent)%)"
            Anchor = 'devices'
            Priority = 1
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Convert-ToOwnershipAssessmentDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value }

    try {
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text) -or $text -eq 'NotCollected (minimum mode)' -or $text -eq 'N/A') {
            return $null
        }
        return [datetime]$text
    }
    catch {
        return $null
    }
}

function Get-OneDriveDefaultOwnerFromUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Url,
        [Parameter(Mandatory = $false)]
        [hashtable]$SegmentToUpnMap,
        [Parameter(Mandatory = $false)]
        [string[]]$KnownDomains
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $null
    }

    try {
        $pathSegment = ($Url.TrimEnd('/') -split '/')[-1]
    }
    catch {
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($pathSegment)) {
        return $null
    }

    $normalizedSegment = $pathSegment.Trim().ToLowerInvariant()

    if ($SegmentToUpnMap -and $SegmentToUpnMap.ContainsKey($normalizedSegment)) {
        $mappedOwner = [string]$SegmentToUpnMap[$normalizedSegment]
        if (-not [string]::IsNullOrWhiteSpace($mappedOwner)) {
            return $mappedOwner.ToLowerInvariant()
        }
    }

    if ($KnownDomains -and $KnownDomains.Count -gt 0) {
        $orderedDomains = @(
            $KnownDomains |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { $_.Trim().ToLowerInvariant() } |
                Select-Object -Unique |
                Sort-Object Length -Descending
        )

        foreach ($domain in $orderedDomains) {
            $encodedDomain = ($domain -replace '\.', '_')
            if ([string]::IsNullOrWhiteSpace($encodedDomain)) { continue }

            $suffix = "_$encodedDomain"
            if (-not $normalizedSegment.EndsWith($suffix)) { continue }

            $encodedLocal = $normalizedSegment.Substring(0, $normalizedSegment.Length - $suffix.Length)
            if ([string]::IsNullOrWhiteSpace($encodedLocal)) { continue }

            $localPart = ($encodedLocal -replace '_', '.')
            return ("{0}@{1}" -f $localPart, $domain).ToLowerInvariant()
        }
    }

    $firstSeparatorIndex = $pathSegment.IndexOf('_')
    if ($firstSeparatorIndex -lt 1) {
        return $pathSegment.ToLowerInvariant()
    }

    $localPart = $pathSegment.Substring(0, $firstSeparatorIndex)
    $domainPart = $pathSegment.Substring($firstSeparatorIndex + 1) -replace '_', '.'
    return ("{0}@{1}" -f $localPart, $domainPart).ToLowerInvariant()
}

function Update-OwnershipGovernanceTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    if (-not $TenantStatsHash) {
        return
    }

    $context = Get-TenantAssessmentContext -TenantStatsHash $TenantStatsHash
    $TenantStatsHash['UnmanagedObjects'] = @{}
    $TenantStatsHash['OneDriveOwnerMismatches'] = @{}
    $TenantStatsHash['OwnershipGovernanceSummary'] = @{}

    $staleOwnerDays = 180
    $staleCutoff = (Get-Date).AddDays(-1 * $staleOwnerDays)
    $ownersByUpn = @{}

    function Add-OwnerProfile {
        param([object]$Record)

        if (-not $Record) { return }
        $upn = [string]$Record.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($upn)) { return }

        $normalizedUpn = $upn.Trim().ToLowerInvariant()
        $accountEnabled = $true
        if ($Record.PSObject.Properties['AccountEnabled']) {
            try { $accountEnabled = [bool]$Record.AccountEnabled } catch { $accountEnabled = $true }
        }
        $lastSignIn = $null
        if ($Record.PSObject.Properties['LastSignInDateTime']) {
            $lastSignIn = Convert-ToOwnershipAssessmentDate -Value $Record.LastSignInDateTime
        }

        if (-not $ownersByUpn.ContainsKey($normalizedUpn)) {
            $ownersByUpn[$normalizedUpn] = [PSCustomObject]@{
                UserPrincipalName = $upn
                AccountEnabled = $accountEnabled
                LastSignInDateTime = $lastSignIn
            }
            return
        }

        $existing = $ownersByUpn[$normalizedUpn]
        if ($accountEnabled -eq $false) {
            $existing | Add-Member -MemberType NoteProperty -Name AccountEnabled -Value $false -Force
        }
        if ($lastSignIn -and (-not $existing.LastSignInDateTime -or $lastSignIn -gt $existing.LastSignInDateTime)) {
            $existing | Add-Member -MemberType NoteProperty -Name LastSignInDateTime -Value $lastSignIn -Force
        }
    }

    foreach ($user in $context.Users) { Add-OwnerProfile -Record $user }
    foreach ($admin in $context.Admins) { Add-OwnerProfile -Record $admin }

    $oneDriveSegmentToUpnMap = @{}
    $knownOwnerDomains = @(
        $ownersByUpn.Keys |
            Where-Object { $_ -and $_ -like '*@*' } |
            ForEach-Object { (($_ -split '@', 2)[1]).Trim().ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )
    foreach ($upn in $ownersByUpn.Keys) {
        if ([string]::IsNullOrWhiteSpace($upn) -or $upn -notlike '*@*') { continue }
        $normalizedUpn = $upn.Trim().ToLowerInvariant()
        $encodedSegment = ($normalizedUpn -replace '@', '_' -replace '\.', '_')
        if ([string]::IsNullOrWhiteSpace($encodedSegment)) { continue }
        if (-not $oneDriveSegmentToUpnMap.ContainsKey($encodedSegment)) {
            $oneDriveSegmentToUpnMap[$encodedSegment] = $normalizedUpn
        }
    }

    function Get-OwnerTokens {
        param($OwnerValue)

        $rawTokens = @()
        if ($null -eq $OwnerValue) {
            return @()
        }

        if ($OwnerValue -is [System.Collections.IEnumerable] -and -not ($OwnerValue -is [string])) {
            foreach ($item in $OwnerValue) {
                if ($null -ne $item) { $rawTokens += [string]$item }
            }
        }
        else {
            $rawTokens += ([string]$OwnerValue -split '[,;]')
        }

        $normalized = New-Object 'System.Collections.Generic.List[string]'
        foreach ($token in $rawTokens) {
            $trimmed = ([string]$token).Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
            $trimmed = ($trimmed -replace '(?i)^smtp:', '').Trim()
            $emailMatch = [regex]::Match($trimmed, '(?i)[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}')
            if ($emailMatch.Success) {
                $trimmed = $emailMatch.Value
            }
            $normalized.Add($trimmed.ToLowerInvariant())
        }

        return @($normalized | Select-Object -Unique)
    }

    function Get-OwnerState {
        param($OwnerValue)

        $owners = @(Get-OwnerTokens -OwnerValue $OwnerValue)
        if ($owners.Count -eq 0) {
            return [PSCustomObject]@{
                Owners = @()
                PrimaryOwner = $null
                State = 'Missing'
                Reason = 'No owner assigned'
                ResolvedOwnerCount = 0
                HealthyOwnerCount = 0
                DisabledOwnerCount = 0
                StaleOwnerCount = 0
                UnknownOwnerCount = 0
            }
        }

        $resolved = 0
        $healthy = 0
        $disabled = 0
        $stale = 0
        $unknown = 0

        foreach ($owner in $owners) {
            if (-not $ownersByUpn.ContainsKey($owner)) {
                $unknown++
                continue
            }

            $resolved++
            $ownerProfile = $ownersByUpn[$owner]
            if ($ownerProfile.AccountEnabled -eq $false) {
                $disabled++
                continue
            }

            $lastSignIn = $ownerProfile.LastSignInDateTime
            if ($lastSignIn -and $lastSignIn -lt $staleCutoff) {
                $stale++
            }
            else {
                $healthy++
            }
        }

        $state = 'Healthy'
        $reason = 'Owner account is active'

        if ($resolved -eq 0) {
            $state = 'Unknown'
            $reason = 'Owner account could not be resolved to collected user data'
        }
        elseif ($disabled -gt 0 -and $healthy -eq 0 -and $stale -eq 0) {
            $state = 'Disabled'
            $reason = 'Owner account is disabled'
        }
        elseif ($stale -gt 0 -and $healthy -eq 0 -and $disabled -eq 0) {
            $state = 'Stale'
            $reason = "Owner account has no sign-in within $staleOwnerDays days"
        }
        elseif (($disabled + $stale) -gt 0) {
            $state = 'Mixed'
            $reason = "At least one resolved owner account is disabled or stale ($staleOwnerDays+ days)"
        }

        return [PSCustomObject]@{
            Owners = $owners
            PrimaryOwner = if ($owners.Count -gt 0) { $owners[0] } else { $null }
            State = $state
            Reason = $reason
            ResolvedOwnerCount = $resolved
            HealthyOwnerCount = $healthy
            DisabledOwnerCount = $disabled
            StaleOwnerCount = $stale
            UnknownOwnerCount = $unknown
        }
    }

    function Convert-ToNullableInt {
        param($Value)
        if ($null -eq $Value) { return $null }
        if ($Value -is [int]) { return [int]$Value }
        if ($Value -is [long]) { return [int]$Value }

        try {
            $text = [string]$Value
            if ([string]::IsNullOrWhiteSpace($text)) { return $null }
            $trimmed = $text.Trim()
            if ($trimmed -match '^\d+$') {
                return [int]$trimmed
            }
        }
        catch {
            return $null
        }

        return $null
    }

    $teamsNameSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($team in $context.Teams) {
        $teamName = [string]$team.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($teamName)) {
            $null = $teamsNameSet.Add($teamName.Trim())
        }
    }

    $unmanagedRows = New-Object System.Collections.Generic.List[object]
    $mismatchRows = New-Object System.Collections.Generic.List[object]
    $unknownOwnerStateCount = 0

    function Add-UnmanagedRow {
        param(
            [string]$Workload,
            [string]$ObjectType,
            [string]$DisplayName,
            [string]$Identifier,
            [string]$CurrentOwner,
            [string]$OwnerState,
            [string]$Reason,
            [string]$Severity,
            [string]$SourceWorksheet
        )

        $unmanagedRows.Add([PSCustomObject]@{
            Workload = $Workload
            ObjectType = $ObjectType
            DisplayName = $DisplayName
            Identifier = $Identifier
            CurrentOwner = $CurrentOwner
            OwnerState = $OwnerState
            Reason = $Reason
            Severity = $Severity
            SourceWorksheet = $SourceWorksheet
        }) | Out-Null
    }

    foreach ($oneDrive in $context.OneDrive) {
        $displayName = if ($oneDrive.Title) { [string]$oneDrive.Title } else { [string]$oneDrive.Url }
        $identifier = if ($oneDrive.Url) { [string]$oneDrive.Url } elseif ($oneDrive.SiteId) { [string]$oneDrive.SiteId } else { $displayName }
        $ownerState = Get-OwnerState -OwnerValue $oneDrive.Owner
        if ($ownerState.State -eq 'Unknown') { $unknownOwnerStateCount++ }

        $expectedOwner = Get-OneDriveDefaultOwnerFromUrl -Url $oneDrive.Url -SegmentToUpnMap $oneDriveSegmentToUpnMap -KnownDomains $knownOwnerDomains
        if ($expectedOwner -and $ownerState.PrimaryOwner -and $expectedOwner -ne $ownerState.PrimaryOwner) {
            $mismatchRows.Add([PSCustomObject]@{
                Workload = 'OneDrive'
                DisplayName = $displayName
                SiteUrl = [string]$oneDrive.Url
                SiteId = [string]$oneDrive.SiteId
                CurrentOwner = [string]$ownerState.PrimaryOwner
                ExpectedDefaultOwner = [string]$expectedOwner
                Assessment = 'Review'
                SourceWorksheet = 'OneDrive'
            }) | Out-Null
        }

        if ($ownerState.State -eq 'Missing') {
            Add-UnmanagedRow -Workload 'OneDrive' -ObjectType 'OneDrive Site' -DisplayName $displayName -Identifier $identifier -CurrentOwner '' -OwnerState $ownerState.State -Reason $ownerState.Reason -Severity 'Risk' -SourceWorksheet 'OneDrive'
        }
        elseif ($ownerState.State -in @('Disabled', 'Stale', 'Mixed')) {
            Add-UnmanagedRow -Workload 'OneDrive' -ObjectType 'OneDrive Site' -DisplayName $displayName -Identifier $identifier -CurrentOwner ([string]$oneDrive.Owner) -OwnerState $ownerState.State -Reason $ownerState.Reason -Severity 'Warning' -SourceWorksheet 'OneDrive'
        }
    }

    foreach ($site in $context.SharePoint) {
        $isTeamsSite = ($site.Template -eq 'TEAMCHANNEL#0' -or $site.IsTeamsChannelConnected -eq $true)
        $workload = if ($isTeamsSite) { 'Teams' } else { 'SharePoint' }
        $objectType = if ($isTeamsSite) { 'Teams Site' } else { 'SharePoint Site' }
        $displayName = if ($site.Title) { [string]$site.Title } else { [string]$site.Url }
        $identifier = if ($site.Url) { [string]$site.Url } elseif ($site.SiteId) { [string]$site.SiteId } else { $displayName }
        $ownerState = Get-OwnerState -OwnerValue $site.Owner
        if ($ownerState.State -eq 'Unknown') { $unknownOwnerStateCount++ }

        if ($ownerState.State -eq 'Missing') {
            Add-UnmanagedRow -Workload $workload -ObjectType $objectType -DisplayName $displayName -Identifier $identifier -CurrentOwner '' -OwnerState $ownerState.State -Reason $ownerState.Reason -Severity 'Risk' -SourceWorksheet 'SharePoint'
        }
        elseif ($ownerState.State -in @('Disabled', 'Stale', 'Mixed')) {
            Add-UnmanagedRow -Workload $workload -ObjectType $objectType -DisplayName $displayName -Identifier $identifier -CurrentOwner ([string]$site.Owner) -OwnerState $ownerState.State -Reason $ownerState.Reason -Severity 'Warning' -SourceWorksheet 'SharePoint'
        }

        if ($isTeamsSite -and -not [string]::IsNullOrWhiteSpace($displayName)) {
            $null = $teamsNameSet.Add($displayName.Trim())
        }
    }

    foreach ($group in $context.Groups) {
        $ownerCount = Convert-ToNullableInt -Value $group.OwnerCount
        if ($null -eq $ownerCount) {
            $unknownOwnerStateCount++
            continue
        }

        if ($ownerCount -eq 0) {
            $groupDisplayName = [string]$group.DisplayName
            $isTeamGroup = (-not [string]::IsNullOrWhiteSpace($groupDisplayName) -and $teamsNameSet.Contains($groupDisplayName.Trim()))
            Add-UnmanagedRow `
                -Workload $(if ($isTeamGroup) { 'Teams' } else { 'Entra ID' }) `
                -ObjectType $(if ($isTeamGroup) { 'Team (M365 Group)' } else { 'Entra Group' }) `
                -DisplayName $groupDisplayName `
                -Identifier ([string]$group.ID) `
                -CurrentOwner 'None (OwnerCount=0)' `
                -OwnerState 'Missing' `
                -Reason 'No owner assigned (OwnerCount=0)' `
                -Severity 'Risk' `
                -SourceWorksheet 'EntraIDGroups'
        }
    }

    foreach ($group in $context.ExchangeGroups) {
        $ownerCount = Convert-ToNullableInt -Value $group.OwnersCount
        if ($null -eq $ownerCount) {
            $unknownOwnerStateCount++
            continue
        }

        if ($ownerCount -eq 0) {
            $groupDisplayName = [string]$group.DisplayName
            $identifier = if ($group.PrimarySMTPAddress) { [string]$group.PrimarySMTPAddress } else { [string]$group.Identity }
            Add-UnmanagedRow `
                -Workload 'Exchange' `
                -ObjectType $(if ($group.RecipientTypeDetails) { [string]$group.RecipientTypeDetails } else { 'Exchange Group' }) `
                -DisplayName $groupDisplayName `
                -Identifier $identifier `
                -CurrentOwner 'None (OwnersCount=0)' `
                -OwnerState 'Missing' `
                -Reason 'No owner assigned (OwnersCount=0)' `
                -Severity 'Risk' `
                -SourceWorksheet 'AllExchangeGroups'
        }
    }

    $sortedUnmanagedRows = @(
        $unmanagedRows |
            Sort-Object @{ Expression = {
                switch ($_.Severity) {
                    'Risk' { 1 }
                    'Warning' { 2 }
                    default { 3 }
                }
            } }, Workload, ObjectType, DisplayName
    )
    $sortedMismatchRows = @($mismatchRows | Sort-Object DisplayName, SiteUrl)

    $unmanagedIndex = 0
    foreach ($row in $sortedUnmanagedRows) {
        $unmanagedIndex++
        $key = "{0:D4}-{1}" -f $unmanagedIndex, ($row.Identifier -replace '[^a-zA-Z0-9@._-]', '_')
        $TenantStatsHash['UnmanagedObjects'][$key] = $row
    }

    $mismatchIndex = 0
    foreach ($row in $sortedMismatchRows) {
        $mismatchIndex++
        $key = "{0:D4}-{1}" -f $mismatchIndex, (($row.SiteUrl) -replace '[^a-zA-Z0-9@._/-]', '_')
        $TenantStatsHash['OneDriveOwnerMismatches'][$key] = $row
    }

    $missingOwnerCount = @($sortedUnmanagedRows | Where-Object { $_.OwnerState -eq 'Missing' }).Count
    $ownerHealthRiskCount = @($sortedUnmanagedRows | Where-Object { $_.OwnerState -in @('Disabled', 'Stale', 'Mixed') }).Count

    $summary = [PSCustomObject]@{
        TotalObjectsReviewed = @($context.OneDrive).Count + @($context.SharePoint).Count + @($context.Groups).Count + @($context.ExchangeGroups).Count
        UnmanagedObjectCount = $sortedUnmanagedRows.Count
        MissingOwnerCount = $missingOwnerCount
        OwnerHealthRiskCount = $ownerHealthRiskCount
        OneDriveOwnerMismatchCount = $sortedMismatchRows.Count
        UnknownOwnerStateCount = $unknownOwnerStateCount
        OneDriveSiteCount = @($context.OneDrive).Count
        SharePointSiteCount = @($context.SharePoint).Count
        TeamsObjectCount = @($context.SharePoint | Where-Object { $_.Template -eq 'TEAMCHANNEL#0' -or $_.IsTeamsChannelConnected -eq $true }).Count
        EntraGroupCount = @($context.Groups).Count
        ExchangeGroupCount = @($context.ExchangeGroups).Count
        StaleOwnerThresholdDays = $staleOwnerDays
    }

    $TenantStatsHash['OwnershipGovernanceSummary']['Summary'] = $summary
}

function Get-OwnershipGovernanceAnalysis {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [array]$UnmanagedObjects,
        [Parameter(Mandatory = $false)]
        [array]$OneDriveOwnerMismatches,
        [Parameter(Mandatory = $false)]
        [object]$OwnershipGovernanceSummary
    )

    $findings = @()
    $unmanagedCount = @($UnmanagedObjects).Count
    $mismatchCount = @($OneDriveOwnerMismatches).Count

    $missingOwnerCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['MissingOwnerCount']) {
        [int]$OwnershipGovernanceSummary.MissingOwnerCount
    } else {
        @($UnmanagedObjects | Where-Object { $_.OwnerState -eq 'Missing' }).Count
    }

    $ownerHealthRiskCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['OwnerHealthRiskCount']) {
        [int]$OwnershipGovernanceSummary.OwnerHealthRiskCount
    } else {
        @($UnmanagedObjects | Where-Object { $_.OwnerState -in @('Disabled', 'Stale', 'Mixed') }).Count
    }

    $unknownOwnerStateCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['UnknownOwnerStateCount']) {
        [int]$OwnershipGovernanceSummary.UnknownOwnerStateCount
    } else {
        0
    }

    if ($missingOwnerCount -gt 0) {
        $findings += @{
            Type = 'Risk'
            Category = 'Unowned Objects'
            Message = "$missingOwnerCount object(s) are missing owners and require ownership assignment"
            Anchor = 'ownership-governance'
            Priority = 1
        }
    }

    if ($ownerHealthRiskCount -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Owner Health'
            Message = "$ownerHealthRiskCount object(s) are owned by disabled or stale owner accounts"
            Anchor = 'ownership-governance'
            Priority = 2
        }
    }

    if ($mismatchCount -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'OneDrive Ownership Mismatch'
            Message = "$mismatchCount OneDrive site(s) have a current owner different from the URL-derived default owner"
            Anchor = 'ownership-governance'
            Priority = 2
        }
    }

    if ($unmanagedCount -eq 0 -and $mismatchCount -eq 0 -and $unknownOwnerStateCount -gt 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Owner Telemetry'
            Message = "$unknownOwnerStateCount object(s) had unresolved ownership state in collected data; review deeper collection modes for complete ownership validation"
            Anchor = 'ownership-governance'
            Priority = 3
        }
    }

    return @{
        Findings = $findings
    }
}

function Get-SharePointOneDriveAnalysis {
    <#
    .SYNOPSIS
        Analyzes SharePoint and OneDrive usage
    #>
    param(
        [array]$SharePointSites,
        [array]$OneDriveSites
    )
    
    $findings = @()
    
    # Large SharePoint sites
    $largeSPSites = $SharePointSites | Where-Object {
        $_.StorageUsedGB -gt $script:DefaultThresholds.SharePointSiteGB
    }
    
    if ($largeSPSites.Count -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large SharePoint Sites'
            Message = "$($largeSPSites.Count) SharePoint sites exceed $($script:DefaultThresholds.SharePointSiteGB) GB"
            Anchor = 'sharepoint-onedrive'
            Priority = 2
        }
    }

    # Large Teams / M365 Groups (over 1 TB) - include names
    $largeTeamsSites = $SharePointSites | Where-Object {
        $_.Template -eq 'TEAMCHANNEL#0' -and $_.StorageUsedGB -gt $script:DefaultThresholds.SharePointSiteGB
    }
    foreach ($site in $largeTeamsSites) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large Teams Site'
            Message = "Teams site '$($site.Title)' exceeds $($script:DefaultThresholds.SharePointSiteGB) GB"
            Anchor = 'sharepoint-onedrive'
            Priority = 2
        }
    }

    $largeGroupSites = $SharePointSites | Where-Object {
        $_.IsOffice365GroupsConnected -and $_.Template -ne 'TEAMCHANNEL#0' -and $_.StorageUsedGB -gt $script:DefaultThresholds.SharePointSiteGB
    }
    foreach ($site in $largeGroupSites) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large M365 Group Site'
            Message = "Microsoft 365 group site '$($site.Title)' exceeds $($script:DefaultThresholds.SharePointSiteGB) GB"
            Anchor = 'sharepoint-onedrive'
            Priority = 2
        }
    }
    
    # Large OneDrive sites
    $largeODSites = $OneDriveSites | Where-Object {
        $_.StorageUsedGB -gt $script:DefaultThresholds.OneDriveSiteGB
    }
    
    if ($largeODSites.Count -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large OneDrive Sites'
            Message = "$($largeODSites.Count) OneDrive sites exceed $($script:DefaultThresholds.OneDriveSiteGB) GB"
            Anchor = 'sharepoint-onedrive'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-InactiveMailboxAnalysis {
    <#
    .SYNOPSIS
        Analyzes inactive mailboxes
    #>
    param([array]$InactiveMailboxes)
    
    $findings = @()
    
    if ($InactiveMailboxes.Count -eq 0) {
        return @{ Findings = @() }
    }
    
    # Count of inactive mailboxes
    $findings += @{
        Type = 'Info'
        Category = 'Inactive Mailboxes'
        Message = "$($InactiveMailboxes.Count) inactive mailboxes are consuming storage (not counted in license utilization)"
        Anchor = 'inactive-mailboxes'
        Priority = 3
    }
    
    # Large inactive mailboxes
    $largeInactive = $InactiveMailboxes | Where-Object {
        try {
            $size = [double]$_.MBXSizeGB
            $size -gt 50
        } catch {
            $false
        }
    }
    
    if ($largeInactive.Count -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large Inactive Mailboxes'
            Message = "$($largeInactive.Count) inactive mailboxes exceed 50 GB - consider if retention is still required"
            Anchor = 'inactive-mailboxes'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-ExchangeHybridAnalysis {
    <#
    .SYNOPSIS
        Analyzes Exchange hybrid configuration and returns findings
    #>
    param([object]$HybridInfo)
    
    $findings = @()
    
    if (-not $HybridInfo) {
        return @{ Findings = @() }
    }
    
    $partialIndicators = (
        ($HybridInfo.InboundOnPremConnectorCount -gt 0) -or
        ($HybridInfo.OutboundOnPremConnectorCount -gt 0) -or
        ($HybridInfo.MigrationEndpointCount -gt 0) -or
        ($HybridInfo.EvidenceCount -gt 0)
    )
    
    if ($HybridInfo.IsHybridConfigured -eq $true) {
        $findings += @{
            Type = 'Info'
            Category = 'Exchange Hybrid'
            Message = "Exchange hybrid detected: $($HybridInfo.HybridType) (evidence: $($HybridInfo.EvidenceCount))"
            Anchor = 'exchange-hybrid'
            Priority = 3
        }
    } elseif ($HybridInfo.HybridStatus -eq 'Possible Hybrid' -or $partialIndicators) {
        $findings += @{
            Type = 'Warning'
            Category = 'Exchange Hybrid'
            Message = "Partial hybrid indicators detected without full configuration (review mail flow connectors and org settings)"
            Anchor = 'exchange-hybrid'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-IdentityAdminAnalysis {
    <#
    .SYNOPSIS
        Analyzes identity and admin data for findings
    #>
    param(
        [array]$Users,
        [array]$Admins,
        [array]$Groups,
        [array]$ConditionalAccessPolicies
    )
    
    $findings = @()
    
    if ($Admins.Count -eq 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Admins'
            Message = "No admin role assignments were found in the report"
            Anchor = 'identity-admins'
            Priority = 2
        }
    }
    
    if ($Groups.Count -eq 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Groups'
            Message = "No Entra ID groups found (or not collected)"
            Anchor = 'identity-admins'
            Priority = 3
        }
    }

    function Convert-ToAssessmentDate {
        param([Parameter(Mandatory = $false)][AllowNull()]$Value)
        if ($null -eq $Value) { return $null }
        if ($Value -is [datetime]) { return $Value }
        try {
            $text = [string]$Value
            if ([string]::IsNullOrWhiteSpace($text) -or $text -eq 'NotCollected (minimum mode)') {
                return $null
            }
            return [datetime]$text
        }
        catch {
            return $null
        }
    }

    $inactiveDaysThreshold = 180
    $inactiveCutoff = (Get-Date).AddDays(-1 * $inactiveDaysThreshold)
    $emergencyAccessRecentSignInThresholdDays = 30
    $emergencyAccessRecentSignInCutoff = (Get-Date).AddDays(-1 * $emergencyAccessRecentSignInThresholdDays)

    $memberUsers = @(
        $Users | Where-Object {
            $null -ne $_ -and
            ($_.UserType -ne 'Guest') -and
            ($_.UserType -ne 'GuestUser') -and
            ($_.UserPrincipalName -notlike '*#EXT#*')
        }
    )
    $enabledMemberUsers = @(
        $memberUsers | Where-Object {
            if ($null -eq $_) { return $false }
            $accountEnabled = $true
            if ($_.PSObject.Properties['AccountEnabled']) {
                try { $accountEnabled = [bool]$_.AccountEnabled } catch { $accountEnabled = $true }
            }
            $accountEnabled
        }
    )

    $usersWithSignin = @()
    foreach ($user in $enabledMemberUsers) {
        $lastSignIn = Convert-ToAssessmentDate -Value $user.LastSignInDateTime
        if ($lastSignIn) {
            $usersWithSignin += [PSCustomObject]@{
                User = $user
                LastSignIn = $lastSignIn
            }
        }
    }

    $inactiveUsers = @($usersWithSignin | Where-Object { $_.LastSignIn -lt $inactiveCutoff })
    if ($inactiveUsers.Count -gt 0) {
        $inactivePct = if ($enabledMemberUsers.Count -gt 0) {
            [math]::Round((($inactiveUsers.Count / $enabledMemberUsers.Count) * 100), 1)
        }
        else { 0 }
        $inactiveType = if ($inactivePct -ge 20) { 'Risk' } else { 'Warning' }
        $inactivePriority = if ($inactivePct -ge 20) { 1 } else { 2 }
        $findings += @{
            Type = $inactiveType
            Category = 'Inactive Users'
            Message = "$($inactiveUsers.Count) enabled member account(s) ($inactivePct%) have no sign-in within the last $inactiveDaysThreshold days"
            Anchor = 'identity-admins'
            Priority = $inactivePriority
        }
    }

    if ($enabledMemberUsers.Count -gt 0) {
        $signInCoveragePct = [math]::Round((($usersWithSignin.Count / $enabledMemberUsers.Count) * 100), 1)
        if ($signInCoveragePct -lt 60) {
            $findings += @{
                Type = 'Info'
                Category = 'Admin Sign-in Telemetry'
                Message = "User sign-in telemetry coverage is $signInCoveragePct% for enabled member users; inactivity findings may be under-reported"
                Anchor = 'identity-admins'
                Priority = 3
            }
        }
    }

    $guestUsers = @(
        $Users | Where-Object {
            ($_.UserType -eq 'Guest') -or
            ($_.UserType -eq 'GuestUser') -or
            ($_.UserPrincipalName -like '*#EXT#*')
        }
    )
    $enabledGuestUsers = @(
        $guestUsers | Where-Object {
            $accountEnabled = $true
            if ($_.PSObject.Properties['AccountEnabled']) {
                try { $accountEnabled = [bool]$_.AccountEnabled } catch { $accountEnabled = $true }
            }
            $accountEnabled
        }
    )
    $inactiveGuests = @(
        $enabledGuestUsers | Where-Object {
            $lastSignIn = Convert-ToAssessmentDate -Value $_.LastSignInDateTime
            $lastSignIn -and $lastSignIn -lt $inactiveCutoff
        }
    )
    if ($inactiveGuests.Count -gt 0) {
        $inactiveGuestPct = if ($enabledGuestUsers.Count -gt 0) {
            [math]::Round((($inactiveGuests.Count / $enabledGuestUsers.Count) * 100), 1)
        } else { 0 }
        $findings += @{
            Type = 'Warning'
            Category = 'Inactive Guest Users'
            Message = "$($inactiveGuests.Count) enabled guest account(s) ($inactiveGuestPct%) have no sign-in within the last $inactiveDaysThreshold days"
            Anchor = 'identity-admins'
            Priority = 2
        }
    }

    $normalizedAdmins = @(
        $Admins | Where-Object { $_ -ne $null }
    )

    $userByUpn = @{}
    foreach ($user in $Users) {
        $userUpn = [string]$user.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($userUpn)) { continue }
        $userByUpn[$userUpn.ToLowerInvariant()] = $user
    }

    if ($normalizedAdmins.Count -gt 0) {
        $globalAdmins = @(
            $normalizedAdmins | Where-Object {
                $roleText = [string]$_.Role
                $roleText -match 'Global Administrator|Company Administrator'
            }
        )
        if ($globalAdmins.Count -gt 5) {
            $findings += @{
                Type = 'Risk'
                Category = 'Global Admin Count'
                Message = "$($globalAdmins.Count) Global Administrator account(s) detected; Microsoft least-privilege guidance recommends reducing standing Global Admins"
                Anchor = 'identity-admins'
                Priority = 1
            }
        }
        elseif ($globalAdmins.Count -gt 4) {
            $findings += @{
                Type = 'Warning'
                Category = 'Global Admin Count'
                Message = "$($globalAdmins.Count) Global Administrator account(s) detected; review whether all are required as standing access"
                Anchor = 'identity-admins'
                Priority = 2
            }
        }

        $staleAdmins = @(
            $normalizedAdmins | Where-Object {
                $lastSignIn = Convert-ToAssessmentDate -Value $_.LastSignInDateTime
                $lastSignIn -and $lastSignIn -lt $inactiveCutoff
            }
        )
        if ($staleAdmins.Count -gt 0) {
            $findings += @{
                Type = 'Risk'
                Category = 'Inactive Admin Accounts'
                Message = "$($staleAdmins.Count) admin account(s) have not signed in within $inactiveDaysThreshold days and should be reviewed or removed from privileged roles"
                Anchor = 'identity-admins'
                Priority = 1
            }
        }

        $globalAdminUsers = @()
        $seenGlobalAdminUpns = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($admin in $globalAdmins) {
            $adminUpn = [string]$admin.UserPrincipalName
            if ([string]::IsNullOrWhiteSpace($adminUpn)) { continue }
            $normalizedUpn = $adminUpn.ToLowerInvariant()
            if ($seenGlobalAdminUpns.Contains($normalizedUpn)) { continue }
            [void]$seenGlobalAdminUpns.Add($normalizedUpn)

            $matchedUser = $null
            if ($userByUpn.ContainsKey($normalizedUpn)) {
                $matchedUser = $userByUpn[$normalizedUpn]
            }

            $accountEnabled = $true
            if ($matchedUser -and $matchedUser.PSObject.Properties['AccountEnabled']) {
                try { $accountEnabled = [bool]$matchedUser.AccountEnabled } catch { $accountEnabled = $true }
            }
            elseif ($admin.PSObject.Properties['AccountEnabled']) {
                try { $accountEnabled = [bool]$admin.AccountEnabled } catch { $accountEnabled = $true }
            }

            $lastSignIn = $null
            if ($matchedUser) {
                $lastSignIn = Convert-ToAssessmentDate -Value $matchedUser.LastSignInDateTime
            }
            if (-not $lastSignIn) {
                $lastSignIn = Convert-ToAssessmentDate -Value $admin.LastSignInDateTime
            }

            $isCloudOnly = $true
            if ($matchedUser -and $matchedUser.PSObject.Properties['OnPremisesSyncEnabled'] -and $matchedUser.OnPremisesSyncEnabled -eq $true) {
                $isCloudOnly = $false
            }

            $objectId = $null
            if ($matchedUser -and $matchedUser.PSObject.Properties['Id'] -and $matchedUser.Id) {
                $objectId = [string]$matchedUser.Id
            }

            $globalAdminUsers += [PSCustomObject]@{
                UserPrincipalName = $adminUpn
                AccountEnabled = $accountEnabled
                LastSignInDateTime = $lastSignIn
                IsCloudOnly = $isCloudOnly
                ObjectId = $objectId
            }
        }

        $enabledCloudOnlyGlobalAdmins = @(
            $globalAdminUsers | Where-Object { $_.AccountEnabled -eq $true -and $_.IsCloudOnly -eq $true }
        )
        $emergencyAccessCandidates = @(
            $enabledCloudOnlyGlobalAdmins | Where-Object {
                (-not $_.LastSignInDateTime) -or ($_.LastSignInDateTime -lt $emergencyAccessRecentSignInCutoff)
            }
        )

        if ($enabledCloudOnlyGlobalAdmins.Count -eq 0) {
            $findings += @{
                Type = 'Warning'
                Category = 'Emergency Access Accounts'
                Message = "No enabled cloud-only Global Administrator accounts were detected; maintain dedicated emergency access accounts per Microsoft guidance"
                Anchor = 'identity-admins'
                Priority = 2
            }
        }
        elseif ($emergencyAccessCandidates.Count -lt 2) {
            $findings += @{
                Type = 'Warning'
                Category = 'Emergency Access Accounts'
                Message = "Only $($emergencyAccessCandidates.Count) cloud-only Global Administrator account(s) appear to fit emergency-access profile (enabled and no recent sign-in > $emergencyAccessRecentSignInThresholdDays days); Microsoft recommends at least two"
                Anchor = 'identity-admins'
                Priority = 2
            }
        }

        $enabledSyncedGlobalAdmins = @(
            $globalAdminUsers | Where-Object { $_.AccountEnabled -eq $true -and $_.IsCloudOnly -eq $false }
        )
        if ($enabledSyncedGlobalAdmins.Count -gt 0) {
            $findings += @{
                Type = 'Info'
                Category = 'Emergency Access Accounts'
                Message = "$($enabledSyncedGlobalAdmins.Count) enabled Global Administrator account(s) are synchronized from on-premises; emergency access accounts should be cloud-only"
                Anchor = 'identity-admins'
                Priority = 3
            }
        }

        if ($emergencyAccessCandidates.Count -gt 0 -and $ConditionalAccessPolicies) {
            $enabledPolicies = @($ConditionalAccessPolicies | Where-Object { $_.State -eq 'enabled' })
            $excludedUserValues = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($policy in $enabledPolicies) {
                $excludedUsersText = [string]$policy.ExcludedUsers
                if ([string]::IsNullOrWhiteSpace($excludedUsersText)) { continue }
                foreach ($token in ($excludedUsersText -split ',')) {
                    $trimmedToken = $token.Trim().ToLowerInvariant()
                    if (-not [string]::IsNullOrWhiteSpace($trimmedToken)) {
                        [void]$excludedUserValues.Add($trimmedToken)
                    }
                }
            }

            $excludedCandidateCount = 0
            foreach ($candidate in $emergencyAccessCandidates) {
                $candidateMatched = $false
                if ($candidate.ObjectId -and $excludedUserValues.Contains($candidate.ObjectId.ToLowerInvariant())) {
                    $candidateMatched = $true
                }
                elseif ($candidate.UserPrincipalName -and $excludedUserValues.Contains($candidate.UserPrincipalName.ToLowerInvariant())) {
                    $candidateMatched = $true
                }
                if ($candidateMatched) { $excludedCandidateCount++ }
            }

            if ($excludedCandidateCount -lt 1) {
                $findings += @{
                    Type = 'Warning'
                    Category = 'Emergency Access CA Exclusions'
                    Message = "No emergency-access candidate appears in enabled Conditional Access exclusion lists; validate lockout-safe emergency account design"
                    Anchor = 'identity-admins'
                    Priority = 2
                }
            }
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-ConditionalAccessMfaAnalysis {
    <#
    .SYNOPSIS
        Analyzes conditional access and MFA configuration
    #>
    param(
        [array]$ConditionalAccessPolicies,
        [object]$AuthConfig,
        [int]$TotalUsers
    )
    
    $findings = @()
    $enabledPolicies = $ConditionalAccessPolicies | Where-Object { $_.State -eq 'enabled' }
    $mfaPolicies = $enabledPolicies | Where-Object { $_.GrantControls_BuiltInControls -match '(?i)mfa' }
    $mfaEnabled = ($mfaPolicies.Count -gt 0) -or ($AuthConfig -and $AuthConfig.MFAEnabled -eq $true)
    
    if ($enabledPolicies.Count -eq 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Conditional Access'
            Message = "No enabled Conditional Access policies detected"
            Anchor = 'conditional-access-mfa'
            Priority = 2
        }
    }
    
    if (-not $mfaEnabled) {
        $findings += @{
            Type = 'Warning'
            Category = 'MFA Enforcement'
            Message = "No active MFA enforcement detected (no enabled MFA CA policies)"
            Anchor = 'conditional-access-mfa'
            Priority = 2
        }
    }

    $legacyAuthBlockPolicies = @(
        $enabledPolicies | Where-Object {
            ([string]$_.ClientAppTypes -match '(?i)exchangeActiveSync|other') -and
            ([string]$_.GrantControls_BuiltInControls -match '(?i)block')
        }
    )
    if ($enabledPolicies.Count -gt 0 -and $legacyAuthBlockPolicies.Count -eq 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Legacy Authentication'
            Message = "No enabled Conditional Access policy appears to explicitly block legacy authentication client app types"
            Anchor = 'conditional-access-mfa'
            Priority = 2
        }
    }

    $riskPolicies = @(
        $enabledPolicies | Where-Object {
            (-not [string]::IsNullOrWhiteSpace([string]$_.SignInRiskLevels_IncludeLevels)) -or
            (-not [string]::IsNullOrWhiteSpace([string]$_.ServicePrincipalRiskLevels_IncludeLevels))
        }
    )
    if ($enabledPolicies.Count -gt 0 -and $riskPolicies.Count -eq 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Risk-based Conditional Access'
            Message = "No enabled risk-based Conditional Access policies detected (sign-in risk / user risk)"
            Anchor = 'conditional-access-mfa'
            Priority = 3
        }
    }

    if ($AuthConfig) {
        $adminConsentWorkflowEnabled = $null
        if ($AuthConfig.PSObject.Properties['AdminConsentWorkflowEnabled']) {
            $rawWorkflow = $AuthConfig.AdminConsentWorkflowEnabled
            if ($rawWorkflow -is [bool]) {
                $adminConsentWorkflowEnabled = [bool]$rawWorkflow
            } elseif ($null -ne $rawWorkflow) {
                $workflowText = [string]$rawWorkflow
                if ($workflowText -match '(?i)^(enabled|true|yes)$') { $adminConsentWorkflowEnabled = $true }
                elseif ($workflowText -match '(?i)^(disabled|false|no)$') { $adminConsentWorkflowEnabled = $false }
            }
        }
        if ($adminConsentWorkflowEnabled -eq $false) {
            $findings += @{
                Type = 'Warning'
                Category = 'Admin Consent Workflow'
                Message = "Admin consent request workflow is disabled; enable it to govern end-user app consent escalation"
                Anchor = 'conditional-access-mfa'
                Priority = 2
            }
        }

        $defaultUserCanCreateApps = $null
        if ($AuthConfig.PSObject.Properties['DefaultUserCanCreateApps']) {
            $rawCreateApps = $AuthConfig.DefaultUserCanCreateApps
            if ($rawCreateApps -is [bool]) {
                $defaultUserCanCreateApps = [bool]$rawCreateApps
            } elseif ($null -ne $rawCreateApps) {
                $createAppsText = [string]$rawCreateApps
                if ($createAppsText -match '(?i)^(yes|true|enabled)$') { $defaultUserCanCreateApps = $true }
                elseif ($createAppsText -match '(?i)^(no|false|disabled)$') { $defaultUserCanCreateApps = $false }
            }
        }
        if ($defaultUserCanCreateApps -eq $true) {
            $findings += @{
                Type = 'Info'
                Category = 'App Consent Governance'
                Message = "Default users are allowed to create app registrations; verify enterprise governance requirements for app creation"
                Anchor = 'conditional-access-mfa'
                Priority = 3
            }
        }

        $permissionGrantPolicyText = $null
        if ($AuthConfig.PSObject.Properties['PermissionGrantPoliciesAssigned']) {
            $permissionGrantPolicyText = (@($AuthConfig.PermissionGrantPoliciesAssigned) -join ',')
        } elseif ($AuthConfig.PSObject.Properties['PermissionGrantPolicies']) {
            $permissionGrantPolicyText = [string]$AuthConfig.PermissionGrantPolicies
        }
        if (-not [string]::IsNullOrWhiteSpace($permissionGrantPolicyText) -and $permissionGrantPolicyText -match '(?i)legacy') {
            $findings += @{
                Type = 'Warning'
                Category = 'App Consent Governance'
                Message = "Permission grant policy assignment includes legacy/default consent behavior; review least-privilege user consent posture"
                Anchor = 'conditional-access-mfa'
                Priority = 2
            }
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-AdConnectAnalysis {
    <#
    .SYNOPSIS
        Analyzes AD Connect sync data
    #>
    param([object]$AdConnect)
    
    $findings = @()
    if ($AdConnect -and $AdConnect.Summary) {
        if ($AdConnect.Summary.OnPremisesSyncEnabled -ne $true) {
            $findings += @{
                Type = 'Info'
                Category = 'AD Connect'
                Message = "Directory sync is not enabled for this tenant"
                Anchor = 'ad-connect'
                Priority = 3
            }
        }
        if ($AdConnect.ErrorCount -gt 0) {
            $findings += @{
                Type = 'Warning'
                Category = 'AD Connect'
                Message = "$($AdConnect.ErrorCount) recent AD Connect sync errors detected"
                Anchor = 'ad-connect'
                Priority = 2
            }
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-FederationAnalysis {
    <#
    .SYNOPSIS
        Analyzes federation and cross-tenant settings
    #>
    param(
        [object]$ExchangeFederation,
        [object]$CrossTenantAccess,
        [object]$ExternalIdentities
    )
    
    $findings = @()
    
    if ($ExchangeFederation -and ($ExchangeFederation.OrganizationRelationshipCount -gt 0 -or $ExchangeFederation.IntraOrgConnectorCount -gt 0)) {
        $findings += @{
            Type = 'Info'
            Category = 'Federation/Relationship'
            Message = "Exchange federation relationships detected (OrgRel: $($ExchangeFederation.OrganizationRelationshipCount), IOC: $($ExchangeFederation.IntraOrgConnectorCount))"
            Anchor = 'exchange-hybrid'
            Priority = 3
        }
    }
    
    if ($CrossTenantAccess -and $CrossTenantAccess.PartnerCount -gt 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Cross-Tenant Access'
            Message = "Cross-tenant access partners configured: $($CrossTenantAccess.PartnerCount)"
            Anchor = 'cross-tenant-access'
            Priority = 3
        }
    }
    
    if ($ExternalIdentities -and $ExternalIdentities.B2BManagementPolicyPresent -eq $false) {
        $findings += @{
            Type = 'Warning'
            Category = 'External Identities'
            Message = "B2B management policy not found; external identity settings may be default/unconfigured"
            Anchor = 'cross-tenant-access'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Convert-MailboxSizeToGB {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $SizeValue
    )

    if (Test-IsBlankDisplayValue -Value $SizeValue) {
        return 0
    }

    try {
        if ($SizeValue -is [ValueType] -and -not ($SizeValue -is [bool]) -and -not ($SizeValue -is [datetime])) {
            $numericValue = [double]$SizeValue
            if ($numericValue -gt 1GB) {
                return [math]::Round(($numericValue / 1GB), 3)
            }
            return [math]::Round($numericValue, 3)
        }
    } catch {}

    try {
        if ($SizeValue.PSObject -and $SizeValue.PSObject.Methods['ToBytes']) {
            return [math]::Round(($SizeValue.ToBytes() / 1GB), 3)
        }
    } catch {}

    $sizeText = [string]$SizeValue
    if ([string]::IsNullOrWhiteSpace($sizeText)) {
        return 0
    }

    $bytesMatch = [regex]::Match($sizeText, '\((?<bytes>[\d,\.]+)\s*bytes\)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($bytesMatch.Success) {
        $bytesText = ($bytesMatch.Groups['bytes'].Value -replace ',', '')
        $bytes = 0.0
        if ([double]::TryParse($bytesText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$bytes)) {
            return [math]::Round(($bytes / 1GB), 3)
        }
    }

    $unitMatch = [regex]::Match($sizeText, '(?<value>[\d,\.]+)\s*(?<unit>TB|GB|MB|KB|B)\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($unitMatch.Success) {
        $valueText = ($unitMatch.Groups['value'].Value -replace ',', '')
        $value = 0.0
        if ([double]::TryParse($valueText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value)) {
            switch ($unitMatch.Groups['unit'].Value.ToUpperInvariant()) {
                'TB' { return [math]::Round(($value * 1024), 3) }
                'GB' { return [math]::Round($value, 3) }
                'MB' { return [math]::Round(($value / 1024), 3) }
                'KB' { return [math]::Round(($value / 1MB), 3) }
                'B'  { return [math]::Round(($value / 1GB), 3) }
            }
        }
    }

    $fallbackText = ($sizeText -replace ',', '')
    $fallback = 0.0
    if ([double]::TryParse($fallbackText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$fallback)) {
        if ($fallback -gt 1GB) {
            return [math]::Round(($fallback / 1GB), 3)
        }
        return [math]::Round($fallback, 3)
    }

    return 0
}

function Get-RecordValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Record,
        [Parameter(Mandatory)]
        [string]$Key
    )

    if ($null -eq $Record) {
        return $null
    }

    if ($Record -is [System.Collections.IDictionary] -and $Record.Contains($Key)) {
        return $Record[$Key]
    }
    if ($Record -is [System.Collections.Specialized.OrderedDictionary] -and $Record.Contains($Key)) {
        return $Record[$Key]
    }
    if ($Record.PSObject -and $Record.PSObject.Properties[$Key]) {
        return $Record.$Key
    }

    return $null
}

function Set-RecordValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Record,
        [Parameter(Mandatory)]
        [string]$Key,
        [AllowNull()]
        $Value
    )

    if ($Record -is [System.Collections.IDictionary]) {
        $Record[$Key] = $Value
        return
    }
    if ($Record -is [System.Collections.Specialized.OrderedDictionary]) {
        $Record[$Key] = $Value
        return
    }

    $Record | Add-Member -MemberType NoteProperty -Name $Key -Value $Value -Force
}

function Get-MailboxStatForRecord {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $MailboxRecord,
        [AllowNull()]
        $StatsHash,
        [switch]$Archive
    )

    if ($null -eq $MailboxRecord -or $null -eq $StatsHash) {
        return $null
    }

    $candidateKeys = Get-MailboxGuidCandidateKeys -MailboxRecord $MailboxRecord -IncludeArchiveGuid:$Archive

    foreach ($key in $candidateKeys) {
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }
        if ($StatsHash.ContainsKey($key)) {
            return $StatsHash[$key]
        }
    }

    return $null
}

function Get-NormalizedMailboxRecords {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [array]$MailboxRecords,
        [AllowNull()]
        $PrimaryMailboxStats,
        [AllowNull()]
        $ArchiveMailboxStats
    )

    if (-not $MailboxRecords -or $MailboxRecords.Count -eq 0) {
        return @()
    }

    $normalized = foreach ($record in $MailboxRecords) {
        if ($null -eq $record) {
            continue
        }

        # Update in place to avoid duplicating large mailbox objects in memory.
        $targetRecord = $record

        $existingMbxSize = Convert-MailboxSizeToGB -SizeValue (Get-RecordValue -Record $targetRecord -Key 'MBXSizeGB')
        $existingArchiveSize = Convert-MailboxSizeToGB -SizeValue (Get-RecordValue -Record $targetRecord -Key 'ArchiveSizeGB')
        $existingMbxItemCount = Get-RecordValue -Record $targetRecord -Key 'MBXItemCount'
        $existingArchiveItemCount = Get-RecordValue -Record $targetRecord -Key 'ArchiveItemCount'
        if (Test-IsBlankDisplayValue -Value $existingMbxItemCount) { $existingMbxItemCount = 0 }
        if (Test-IsBlankDisplayValue -Value $existingArchiveItemCount) { $existingArchiveItemCount = 0 }

        $primaryStat = Get-MailboxStatForRecord -MailboxRecord $targetRecord -StatsHash $PrimaryMailboxStats
        $archiveStat = Get-MailboxStatForRecord -MailboxRecord $targetRecord -StatsHash $ArchiveMailboxStats -Archive

        $primaryTotalItemSize = Get-RecordValue -Record $primaryStat -Key 'TotalItemSize'
        $archiveTotalItemSize = Get-RecordValue -Record $archiveStat -Key 'TotalItemSize'

        $resolvedMbxSize = if ($existingMbxSize -gt 0) { $existingMbxSize } elseif ($primaryStat) { Convert-MailboxSizeToGB -SizeValue $primaryTotalItemSize } else { 0 }
        $resolvedArchiveSize = if ($existingArchiveSize -gt 0) { $existingArchiveSize } elseif ($archiveStat) { Convert-MailboxSizeToGB -SizeValue $archiveTotalItemSize } else { 0 }

        $resolvedMbxItemCount = if (-not (Test-IsBlankDisplayValue -Value $existingMbxItemCount) -and [string]$existingMbxItemCount -ne '0') {
            $existingMbxItemCount
        } elseif ($primaryStat -and $null -ne (Get-RecordValue -Record $primaryStat -Key 'ItemCount')) {
            Get-RecordValue -Record $primaryStat -Key 'ItemCount'
        } else {
            0
        }

        $resolvedArchiveItemCount = if (-not (Test-IsBlankDisplayValue -Value $existingArchiveItemCount) -and [string]$existingArchiveItemCount -ne '0') {
            $existingArchiveItemCount
        } elseif ($archiveStat -and $null -ne (Get-RecordValue -Record $archiveStat -Key 'ItemCount')) {
            Get-RecordValue -Record $archiveStat -Key 'ItemCount'
        } else {
            0
        }

        Set-RecordValue -Record $targetRecord -Key 'MBXSizeGB' -Value ([math]::Round($resolvedMbxSize, 3))
        Set-RecordValue -Record $targetRecord -Key 'ArchiveSizeGB' -Value ([math]::Round($resolvedArchiveSize, 3))
        Set-RecordValue -Record $targetRecord -Key 'MBXItemCount' -Value $resolvedMbxItemCount
        Set-RecordValue -Record $targetRecord -Key 'ArchiveItemCount' -Value $resolvedArchiveItemCount

        $targetRecord
    }

    return @($normalized)
}

function Get-TenantAssessmentContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    function Get-ContextArray {
        param([string]$Key)

        if (-not $TenantStatsHash.ContainsKey($Key)) {
            return @()
        }

        $value = $TenantStatsHash[$Key]
        if ($null -eq $value) {
            return @()
        }

        if ($value -is [hashtable] -or $value -is [System.Collections.Specialized.OrderedDictionary]) {
            return @($value.Values)
        }

        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
            return @($value)
        }

        return @($value)
    }

    function Resolve-ContextSummaryRecord {
        param(
            [AllowNull()]
            $Container
        )

        if ($null -eq $Container) {
            return $null
        }

        $summaryValue = $null
        if ($Container -is [System.Collections.IDictionary]) {
            if ($Container.Contains('Summary')) {
                $summaryValue = $Container['Summary']
            }
            else {
                $summaryValue = $Container
            }
        }
        elseif ($Container -is [System.Collections.Specialized.OrderedDictionary]) {
            if ($Container.Contains('Summary')) {
                $summaryValue = $Container['Summary']
            }
            else {
                $summaryValue = $Container
            }
        }
        elseif ($Container -is [array]) {
            $summaryValue = @($Container | Select-Object -First 1)
            if ($summaryValue.Count -gt 0) {
                $summaryValue = $summaryValue[0]
            }
            else {
                $summaryValue = $null
            }
        }
        else {
            $summaryValue = $Container
        }

        if ($summaryValue -is [System.Collections.IDictionary]) {
            return [PSCustomObject]$summaryValue
        }
        if ($summaryValue -is [System.Collections.Specialized.OrderedDictionary]) {
            return [PSCustomObject]$summaryValue
        }
        return $summaryValue
    }

    $authConfig = $null
    if ($TenantStatsHash.ContainsKey('AuthenticationConfig')) {
        $authContainer = $TenantStatsHash['AuthenticationConfig']
        if ($authContainer -is [hashtable] -and $authContainer.ContainsKey('Configuration')) {
            $authConfig = $authContainer['Configuration']
        }
    }

    $adConnect = $null
    if ($TenantStatsHash.ContainsKey('AdConnectConfiguration')) {
        $adConnect = $TenantStatsHash['AdConnectConfiguration']
    }

    $mfaRegistrationSummary = $null
    if ($TenantStatsHash.ContainsKey('MfaRegistrationSummary')) {
        $mfaRegistrationSummary = $TenantStatsHash['MfaRegistrationSummary']
    }

    $hybridInfo = $null
    if ($TenantStatsHash.ContainsKey('HybridConfiguration')) {
        $hybridContainer = $TenantStatsHash['HybridConfiguration']
        if ($hybridContainer -is [hashtable] -and $hybridContainer.ContainsKey('ExchangeHybrid')) {
            $hybridInfo = $hybridContainer['ExchangeHybrid']
        }
    }

    $federationExchange = $null
    $federationCrossTenant = $null
    $federationExternal = $null
    if ($TenantStatsHash.ContainsKey('FederationConfiguration')) {
        $fedContainer = $TenantStatsHash['FederationConfiguration']
        if ($fedContainer -is [hashtable]) {
            if ($fedContainer.ContainsKey('ExchangeFederation')) { $federationExchange = $fedContainer['ExchangeFederation'] }
            if ($fedContainer.ContainsKey('CrossTenantAccess')) { $federationCrossTenant = $fedContainer['CrossTenantAccess'] }
            if ($fedContainer.ContainsKey('ExternalIdentities')) { $federationExternal = $fedContainer['ExternalIdentities'] }
        }
    }

    $teamsVoice = $null
    if ($TenantStatsHash.ContainsKey('TeamsVoice')) {
        $teamsVoice = $TenantStatsHash['TeamsVoice']
    }

    $spamFilteringSummary = $null
    if ($TenantStatsHash.ContainsKey('SpamFilteringSummary')) {
        $spamSummaryContainer = $TenantStatsHash['SpamFilteringSummary']
        if ($spamSummaryContainer -is [hashtable] -and $spamSummaryContainer.ContainsKey('Summary')) {
            $spamFilteringSummary = $spamSummaryContainer['Summary']
        }
    }
    if (-not $spamFilteringSummary -and $TenantStatsHash.ContainsKey('SpamFilteringConfig')) {
        $spamConfigContainer = $TenantStatsHash['SpamFilteringConfig']
        if ($spamConfigContainer -is [hashtable] -and $spamConfigContainer.ContainsKey('Configuration')) {
            $spamFilteringSummary = $spamConfigContainer['Configuration']
        }
    }

    $smtpRelaySummary = $null
    if ($TenantStatsHash.ContainsKey('SMTPRelaySummary')) {
        $smtpSummaryContainer = $TenantStatsHash['SMTPRelaySummary']
        if ($smtpSummaryContainer -is [hashtable] -and $smtpSummaryContainer.ContainsKey('Summary')) {
            $smtpRelaySummary = $smtpSummaryContainer['Summary']
        }
    }
    if (-not $smtpRelaySummary -and $TenantStatsHash.ContainsKey('SMTPRelayConfig')) {
        $smtpConfigContainer = $TenantStatsHash['SMTPRelayConfig']
        if ($smtpConfigContainer -is [hashtable] -and $smtpConfigContainer.ContainsKey('Configuration')) {
            $smtpRelaySummary = $smtpConfigContainer['Configuration']
        }
    }

    $emailActivitySummary = $null
    if ($TenantStatsHash.ContainsKey('EmailActivitySummary')) {
        $emailActivitySummary = Resolve-ContextSummaryRecord -Container $TenantStatsHash['EmailActivitySummary']
    }

    $ownershipGovernanceSummary = $null
    if ($TenantStatsHash.ContainsKey('OwnershipGovernanceSummary')) {
        $ownershipGovernanceSummary = Resolve-ContextSummaryRecord -Container $TenantStatsHash['OwnershipGovernanceSummary']
    }

    $mailboxSourceKey = if ($TenantStatsHash.ContainsKey('MailboxFullDetails')) {
        'MailboxFullDetails'
    } elseif ($TenantStatsHash.ContainsKey('AllMailboxes')) {
        'AllMailboxes'
    } else {
        $null
    }
    $rawMailboxes = if ($mailboxSourceKey) { Get-ContextArray -Key $mailboxSourceKey } else { @() }
    $primaryMailboxStats = if ($TenantStatsHash.ContainsKey('PrimaryMailboxStats')) { $TenantStatsHash['PrimaryMailboxStats'] } else { $null }
    $archiveMailboxStats = if ($TenantStatsHash.ContainsKey('ArchiveMailboxStats')) { $TenantStatsHash['ArchiveMailboxStats'] } else { $null }
    $normalizedMailboxes = Get-NormalizedMailboxRecords -MailboxRecords $rawMailboxes -PrimaryMailboxStats $primaryMailboxStats -ArchiveMailboxStats $archiveMailboxStats

    $rawInactiveMailboxes = if ($TenantStatsHash.ContainsKey('InactiveMailboxDetails')) {
        Get-ContextArray -Key 'InactiveMailboxDetails'
    } else {
        @($normalizedMailboxes | Where-Object { $_.PSObject.Properties['IsInactiveMailbox'] -and $_.IsInactiveMailbox -eq $true })
    }
    $normalizedInactiveMailboxes = Get-NormalizedMailboxRecords -MailboxRecords $rawInactiveMailboxes -PrimaryMailboxStats $primaryMailboxStats -ArchiveMailboxStats $archiveMailboxStats

    return [PSCustomObject]@{
        Licenses               = Get-ContextArray -Key 'LicenseSKUs'
        Recipients             = Get-ContextArray -Key 'AllRecipients'
        Mailboxes              = $normalizedMailboxes
        InactiveMailboxes      = $normalizedInactiveMailboxes
        PublicFolders          = Get-ContextArray -Key 'PublicFolderDetails'
        SharePoint             = Get-ContextArray -Key 'SharePoint'
        OneDrive               = Get-ContextArray -Key 'OneDrive'
        Domains                = Get-ContextArray -Key 'Domains'
        Devices                = Get-ContextArray -Key 'DeviceDetails'
        SecureScore            = Get-ContextArray -Key 'SecuritySecureScore'
        SecureScoreActions     = Get-ContextArray -Key 'SecureScoreActions'
        Teams                  = Get-ContextArray -Key 'AllTeams'
        Users                  = Get-ContextArray -Key 'Users'
        Admins                 = Get-ContextArray -Key 'Admins'
        Groups                 = Get-ContextArray -Key 'EntraIDGroups'
        ExchangeGroups         = Get-ContextArray -Key 'AllExchangeGroups'
        ConditionalAccess      = Get-ContextArray -Key 'ConditionalAccessPolicies'
        MailFlowConnectors     = Get-ContextArray -Key 'MailFlowConnectors'
        RemoteDomains          = Get-ContextArray -Key 'RemoteDomains'
        AuthConfig             = $authConfig
        MfaRegistrationSummary = $mfaRegistrationSummary
        AdConnect              = $adConnect
        HybridInfo             = $hybridInfo
        FederationExchange     = $federationExchange
        FederationCrossTenant  = $federationCrossTenant
        FederationExternal     = $federationExternal
        TeamsVoice             = $teamsVoice
        SpamFilteringSummary   = $spamFilteringSummary
        SMTPRelaySummary       = $smtpRelaySummary
        EmailActivitySummary   = $emailActivitySummary
        EmailActivityTopSenders = Get-ContextArray -Key 'EmailActivityTopSenders'
        EmailActivityTopReceivers = Get-ContextArray -Key 'EmailActivityTopReceivers'
        OwnershipGovernanceSummary = $ownershipGovernanceSummary
        UnmanagedObjects       = Get-ContextArray -Key 'UnmanagedObjects'
        OneDriveOwnerMismatches = Get-ContextArray -Key 'OneDriveOwnerMismatches'
    }
}

function Get-AssessmentWorksheetName {
    [CmdletBinding()]
    param([string]$Anchor)

    switch ($Anchor) {
        'licenses' { return 'LicenseSKUs' }
        'domains' { return 'Domains' }
        'domains-dns' { return 'Domains' }
        'identity-admins' { return 'Users' }
        'mailboxes' { return 'MailboxFullDetails' }
        'inactive-mailboxes' { return 'InactiveMailboxDetails' }
        'sharepoint-onedrive' { return 'SharePoint / OneDrive' }
        'devices' { return 'DeviceDetails' }
        'ad-connect' { return 'AdConnectConfiguration' }
        'conditional-access-mfa' { return 'ConditionalAccessPolicies' }
        'exchange-hybrid' { return 'HybridConfiguration' }
        'cross-tenant-access' { return 'FederationConfiguration' }
        'secure-score' { return 'SecureScoreActions' }
        'ownership-governance' { return 'UnmanagedObjects' }
        default { return $null }
    }
}

function Get-AssessmentRecommendationText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Finding
    )

    switch ($Finding.Anchor) {
        'licenses' { return 'Review SKU capacity, reclaim unused assignments, and align target-tenant licensing before cutover.' }
        'domains' { return 'Verify all domains and confirm authoritative routing before migration sequencing.' }
        'domains-dns' {
            switch ([string]$Finding.Category) {
                'DMARC' { return 'Publish and enforce DMARC for custom email domains to improve spoofing protection and align with Microsoft email security guidance.' }
                'SPF' { return 'Ensure SPF includes Microsoft 365 mail protection endpoints and remains within DNS lookup limits.' }
                'DKIM' { return 'Configure both DKIM selectors and enable DKIM signing for custom domains used for mail flow.' }
                'Email Authentication Coverage' { return 'Raise SPF/DKIM/DMARC coverage on all active custom mail domains and prioritize DMARC enforcement (quarantine/reject).' }
                'Anti-Spoofing Bypass' { return 'Review trusted-IP and bypass rules to ensure anti-spoofing controls are not unintentionally bypassed.' }
                'Anti-Spoofing Controls' { return 'Track anti-spoofing controls over time and keep trusted bypasses and relay exceptions tightly scoped.' }
                'SMTP AUTH Exposure' { return 'Disable SMTP AUTH where possible and use modern authentication or scoped relay alternatives for legacy apps/devices.' }
                default { return 'Review MX, autodiscover, and mail-routing records to plan coexistence and cutover.' }
            }
        }
        'identity-admins' {
            switch ([string]$Finding.Category) {
                'Inactive Users' { return 'Follow Microsoft identity hygiene guidance: disable or investigate member accounts inactive for 180+ days and keep only justified exceptions.' }
                'Inactive Guest Users' { return 'Review stale guest accounts and use Entra access reviews/lifecycle governance to remove unneeded external identities.' }
                'Inactive Admin Accounts' { return 'Remove or time-bound stale privileged assignments and use Entra PIM eligible roles instead of standing admin access.' }
                'Global Admin Count' { return 'Follow least-privilege guidance: keep only a small set of Global Administrators (typically 2-4) and delegate other tasks to scoped roles.' }
                'Admin Sign-in Telemetry' { return 'Ensure sign-in telemetry is available for privileged accounts (app permissions/log retention) so stale-admin monitoring is reliable.' }
                'Emergency Access Accounts' { return 'Maintain at least two cloud-only emergency access accounts, monitor them, and keep them excluded from daily operational use.' }
                'Emergency Access CA Exclusions' { return 'Validate Conditional Access emergency-access exclusions so break-glass accounts can sign in during policy or identity outages.' }
                default { return 'Validate admin access, guest usage, and group ownership before identity migration activities.' }
            }
        }
        'mailboxes' { return 'Identify oversized or specialized mailboxes early to plan batching, archives, and exception handling.' }
        'inactive-mailboxes' { return 'Decide whether inactive mailboxes need retention, restore, or exclusion from scope.' }
        'sharepoint-onedrive' { return 'Use site inventory, ownership, and storage metrics to prioritize high-risk collaboration workloads.' }
        'devices' { return 'Review stale and non-compliant devices before identity and endpoint cutover.' }
        'ad-connect' { return 'Document synchronization dependencies and plan cloud identity cutover or staged decommissioning.' }
        'conditional-access-mfa' {
            switch ([string]$Finding.Category) {
                'Legacy Authentication' { return 'Block legacy authentication with Conditional Access and verify modern-auth readiness before enforcement.' }
                'Risk-based Conditional Access' { return 'Implement sign-in risk and user risk Conditional Access policies to align with Microsoft identity protection practices.' }
                'Admin Consent Workflow' { return 'Enable and tune the admin consent request workflow so app consent escalations follow governance controls.' }
                'App Consent Governance' { return 'Harden app consent and app registration settings to reduce over-privileged or unmanaged enterprise app risk.' }
                default { return 'Review CA and MFA design to avoid post-migration lockouts or authentication regressions.' }
            }
        }
        'exchange-hybrid' { return 'Validate hybrid, connectors, and migration endpoints because they affect tenant-to-tenant messaging strategy.' }
        'cross-tenant-access' { return 'Review cross-tenant and B2B settings for coexistence, external collaboration, and post-migration cleanup.' }
        'secure-score' { return 'Use the mapped Microsoft Secure Score action to prioritize remediation with the highest security impact.' }
        'ownership-governance' {
            switch ([string]$Finding.Category) {
                'Unowned Objects' { return 'Assign at least one accountable owner to each collaboration object and validate ownership handoff before migration or governance workflows.' }
                'Owner Health' { return 'Reassign ownership from disabled or stale accounts to active custodians and formalize backup ownership coverage.' }
                'OneDrive Ownership Mismatch' { return 'Review OneDrive sites where the current owner differs from the URL-derived default user and confirm documented stewardship.' }
                default { return 'Review ownership governance tables and assign healthy owners for all unmanaged or mismatched objects.' }
            }
        }
        default { return 'Review the related worksheet and validate whether remediation is required for your migration or security objectives.' }
    }
}

function Update-AssessmentReportTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    if (-not $TenantStatsHash) {
        return
    }

    $context = Get-TenantAssessmentContext -TenantStatsHash $TenantStatsHash

    $TenantStatsHash['BestPractices'] = @{}
    $TenantStatsHash['BestPracticeFindings'] = @{}
    $TenantStatsHash['MigrationReadiness'] = @{}

    $summaryRows = New-Object System.Collections.Generic.List[object]
    $findingRows = New-Object System.Collections.Generic.List[object]
    $migrationRows = New-Object System.Collections.Generic.List[object]
    $summaryIndex = 0
    $findingIndex = 0
    $migrationIndex = 0

    function Add-AreaSummary {
        param(
            [string]$Area,
            [array]$AreaFindings,
            [string]$AssessmentType,
            [string]$RelatedWorksheet,
            [string]$Notes
        )

        $criticalCount = @($AreaFindings | Where-Object { $_.Type -eq 'Risk' }).Count
        $warningCount = @($AreaFindings | Where-Object { $_.Type -eq 'Warning' }).Count
        $infoCount = @($AreaFindings | Where-Object { $_.Type -eq 'Info' }).Count

        $status = if ($criticalCount -gt 0) {
            'Critical'
        } elseif ($warningCount -gt 0) {
            'Warning'
        } elseif ($AreaFindings.Count -gt 0) {
            'Informational'
        } else {
            'Healthy'
        }

        $topFinding = @(
            $AreaFindings |
                Sort-Object @{ Expression = {
                    switch ([string]$_.Type) {
                        'Risk' { 1 }
                        'Warning' { 2 }
                        'Info' { 3 }
                        default { 9 }
                    }
                } }, Priority |
                Select-Object -First 1
        )
        $topCategoryRollups = @(
            $AreaFindings |
                Group-Object Category |
                Sort-Object @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Descending = $false } |
                Select-Object -First 3 |
                ForEach-Object {
                    $categoryName = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Uncategorized' } else { $_.Name }
                    "{0}: {1}" -f $categoryName, $_.Count
                }
        )
        $primaryFindingText = if ($AreaFindings.Count -gt 0) {
            $severitySummary = "{0} critical, {1} warning, {2} informational finding(s)" -f $criticalCount, $warningCount, $infoCount
            if ($topCategoryRollups.Count -gt 0) {
                "$severitySummary. Top signals: $([string]::Join('; ', $topCategoryRollups))."
            } else {
                "$severitySummary."
            }
        } else {
            'No automated findings detected for this assessment area.'
        }
        $recommendedAction = if ($topFinding.Count -gt 0) { Get-AssessmentRecommendationText -Finding $topFinding[0] } else { 'Use the detailed workload worksheets for validation and migration planning.' }

        $summaryRows.Add([PSCustomObject]@{
            Area               = $Area
            Status             = $status
            CriticalFindings   = $criticalCount
            WarningFindings    = $warningCount
            InfoFindings       = $infoCount
            TotalFindings      = $AreaFindings.Count
            PrimaryFinding     = $primaryFindingText
            RecommendedAction  = $recommendedAction
            RelatedWorksheet   = $RelatedWorksheet
            AssessmentType     = $AssessmentType
            Notes              = $Notes
        }) | Out-Null

        foreach ($finding in $AreaFindings) {
            $findingRows.Add([PSCustomObject]@{
                Area               = $Area
                Severity           = $finding.Type
                Category           = $finding.Category
                Message            = $finding.Message
                Priority           = $finding.Priority
                RelatedSection     = $finding.Anchor
                RelatedWorksheet   = Get-AssessmentWorksheetName -Anchor $finding.Anchor
                RecommendedAction  = Get-AssessmentRecommendationText -Finding $finding
                SourceType         = $AssessmentType
            }) | Out-Null
        }
    }

    function Add-MigrationRow {
        param(
            [string]$Category,
            [string]$Item,
            [string]$Status,
            [string]$Value,
            [string]$Notes,
            [string]$MigrationAction,
            [string]$SourceWorksheet
        )

        $migrationRows.Add([PSCustomObject]@{
            Category         = $Category
            Item             = $Item
            Status           = $Status
            Value            = $Value
            Notes            = $Notes
            MigrationAction  = $MigrationAction
            SourceWorksheet  = $SourceWorksheet
        }) | Out-Null
    }

    if ($context.Licenses.Count -gt 0) {
        $licAnalysis = Get-LicenseAnalysis -Licenses $context.Licenses -UserCount $context.Users.Count
        Add-AreaSummary -Area 'Licensing' -AreaFindings $licAnalysis.Findings -AssessmentType 'Assessment heuristic using Microsoft 365 license data' -RelatedWorksheet 'LicenseSKUs' -Notes 'Evaluates capacity, at-capacity SKUs, and high utilization.'
    }

    if ($context.Domains.Count -gt 0) {
        $domainAnalysis = Get-DomainAnalysis -Domains $context.Domains -SpamFilteringSummary $context.SpamFilteringSummary -SMTPRelaySummary $context.SMTPRelaySummary
        Add-AreaSummary -Area 'Domains' -AreaFindings $domainAnalysis.Findings -AssessmentType 'Assessment heuristic using Microsoft 365 domain state' -RelatedWorksheet 'Domains' -Notes 'Highlights verification and mail-routing concerns relevant to migration cutover.'
    }

    if ($context.Users.Count -gt 0 -or $context.Admins.Count -gt 0 -or $context.Groups.Count -gt 0) {
        $identityAnalysis = Get-IdentityAdminAnalysis -Users $context.Users -Admins $context.Admins -Groups $context.Groups -ConditionalAccessPolicies $context.ConditionalAccess
        Add-AreaSummary -Area 'Identity & Admins' -AreaFindings $identityAnalysis.Findings -AssessmentType 'Assessment heuristic using Entra users, groups, and admin assignments' -RelatedWorksheet 'Users' -Notes 'Surfaces admin and group inventory signals that affect migration readiness.'
    }

    if ($context.Mailboxes.Count -gt 0 -or $context.PublicFolders.Count -gt 0) {
        $mailboxAnalysis = Get-MailboxAnalysis -Mailboxes $context.Mailboxes
        Add-AreaSummary -Area 'Mailboxes' -AreaFindings $mailboxAnalysis.Findings -AssessmentType 'Assessment heuristic using Exchange mailbox inventory' -RelatedWorksheet 'MailboxFullDetails' -Notes 'Flags mailbox sizing and archive patterns that influence batch and exception planning.'
    }

    if ($context.InactiveMailboxes.Count -gt 0) {
        $inactiveAnalysis = Get-InactiveMailboxAnalysis -InactiveMailboxes $context.InactiveMailboxes
        Add-AreaSummary -Area 'Inactive Mailboxes' -AreaFindings $inactiveAnalysis.Findings -AssessmentType 'Assessment heuristic using inactive mailbox inventory' -RelatedWorksheet 'InactiveMailboxDetails' -Notes 'Supports retention and scope decisions for mailbox migration.'
    }

    if ($context.SharePoint.Count -gt 0 -or $context.OneDrive.Count -gt 0) {
        $spodAnalysis = Get-SharePointOneDriveAnalysis -SharePointSites $context.SharePoint -OneDriveSites $context.OneDrive
        Add-AreaSummary -Area 'SharePoint & OneDrive' -AreaFindings $spodAnalysis.Findings -AssessmentType 'Assessment heuristic using collaboration site inventory' -RelatedWorksheet 'SharePoint / OneDrive' -Notes 'Flags oversized sites and owner-linked OneDrive inventory for migration planning.'
    }

    if (
        $context.OwnershipGovernanceSummary -or
        $context.UnmanagedObjects.Count -gt 0 -or
        $context.OneDriveOwnerMismatches.Count -gt 0
    ) {
        $ownershipAnalysis = Get-OwnershipGovernanceAnalysis `
            -UnmanagedObjects $context.UnmanagedObjects `
            -OneDriveOwnerMismatches $context.OneDriveOwnerMismatches `
            -OwnershipGovernanceSummary $context.OwnershipGovernanceSummary
        Add-AreaSummary -Area 'Ownership & Stewardship' -AreaFindings $ownershipAnalysis.Findings -AssessmentType 'Assessment heuristic using ownership governance signals across collaboration and group workloads' -RelatedWorksheet 'UnmanagedObjects' -Notes 'Highlights unowned objects, unhealthy owners, and OneDrive owner-mismatch governance review items.'
    }

    if ($context.Devices.Count -gt 0) {
        $deviceAnalysis = Get-DeviceAnalysis -Devices $context.Devices
        Add-AreaSummary -Area 'Devices' -AreaFindings $deviceAnalysis.Findings -AssessmentType 'Assessment heuristic using Entra device inventory' -RelatedWorksheet 'DeviceDetails' -Notes 'Highlights stale and non-compliant devices that can affect user cutover readiness.'
    }

    if ($context.AdConnect) {
        $adAnalysis = Get-AdConnectAnalysis -AdConnect $context.AdConnect
        Add-AreaSummary -Area 'AD Connect / Sync' -AreaFindings $adAnalysis.Findings -AssessmentType 'Assessment heuristic using directory synchronization signals' -RelatedWorksheet 'AdConnectConfiguration' -Notes 'Identifies synchronization dependencies and recent sync issues.'
    }

    if ($context.ConditionalAccess.Count -gt 0 -or $context.AuthConfig) {
        $caAnalysis = Get-ConditionalAccessMfaAnalysis -ConditionalAccessPolicies $context.ConditionalAccess -AuthConfig $context.AuthConfig -TotalUsers $context.Users.Count
        Add-AreaSummary -Area 'Conditional Access & MFA' -AreaFindings $caAnalysis.Findings -AssessmentType 'Assessment heuristic using Entra protection settings' -RelatedWorksheet 'ConditionalAccessPolicies' -Notes 'Evaluates MFA and Conditional Access coverage with a Microsoft best-practice orientation.'
    }

    if ($context.HybridInfo) {
        $hybridAnalysis = Get-ExchangeHybridAnalysis -HybridInfo $context.HybridInfo
        $fedAnalysis = Get-FederationAnalysis -ExchangeFederation $context.FederationExchange -CrossTenantAccess $context.FederationCrossTenant -ExternalIdentities $context.FederationExternal
        $hybridFindings = @($hybridAnalysis.Findings + $fedAnalysis.Findings)
        Add-AreaSummary -Area 'Hybrid / Federation' -AreaFindings $hybridFindings -AssessmentType 'Assessment heuristic using Exchange hybrid and federation signals' -RelatedWorksheet 'HybridConfiguration' -Notes 'Highlights hybrid dependencies, federation, and cross-tenant settings relevant to coexistence and migration.'
    }

    if ($context.SecureScore.Count -gt 0) {
        $latestScore = @($context.SecureScore | Sort-Object CreatedDateTime -Descending | Select-Object -First 1)
        $secureScoreFindings = @()
        if ($latestScore.Count -gt 0) {
            $scorePct = 0
            try { $scorePct = [double]$latestScore[0].SecurityScorePercentage } catch { $scorePct = 0 }
            if ($scorePct -lt 60) {
                $secureScoreFindings += @{
                    Type = 'Risk'
                    Category = 'Secure Score'
                    Message = "Microsoft Secure Score is $([math]::Round($scorePct,1))%, which is below the target range for a mature tenant baseline."
                    Anchor = 'secure-score'
                    Priority = 1
                }
            } elseif ($scorePct -lt 80) {
                $secureScoreFindings += @{
                    Type = 'Warning'
                    Category = 'Secure Score'
                    Message = "Microsoft Secure Score is $([math]::Round($scorePct,1))%; prioritize high-rank actions to raise baseline security."
                    Anchor = 'secure-score'
                    Priority = 2
                }
            } else {
                $secureScoreFindings += @{
                    Type = 'Info'
                    Category = 'Secure Score'
                    Message = "Microsoft Secure Score is $([math]::Round($scorePct,1))%, indicating a relatively strong security baseline."
                    Anchor = 'secure-score'
                    Priority = 3
                }
            }
        }
        Add-AreaSummary -Area 'Secure Score' -AreaFindings $secureScoreFindings -AssessmentType 'Microsoft Secure Score recommendation mapping' -RelatedWorksheet 'SecureScoreActions' -Notes 'Uses Microsoft Secure Score snapshots and control profile metadata, including Microsoft Learn action URLs.'
    }

    foreach ($summaryRow in $summaryRows) {
        $summaryIndex++
        $TenantStatsHash['BestPractices'][("{0:D3}-{1}" -f $summaryIndex, $summaryRow.Area)] = $summaryRow
    }

    foreach ($findingRow in $findingRows) {
        $findingIndex++
        $TenantStatsHash['BestPracticeFindings'][("{0:D3}-{1}" -f $findingIndex, $findingRow.Area)] = $findingRow
    }

    $verifiedDomains = @($context.Domains | Where-Object { $_.Verified -eq $true }).Count
    $unverifiedDomains = @($context.Domains | Where-Object { $_.Verified -ne $true }).Count
    $nonM365MxDomains = @($context.Domains | Where-Object { $_.Office365MailExchanger -eq $false }).Count
    $dirSyncEnabled = [bool]($context.AdConnect -and $context.AdConnect.Summary -and $context.AdConnect.Summary.OnPremisesSyncEnabled -eq $true)
    $guestCount = @($context.Users | Where-Object { $_.UserType -match 'Guest' -or $_.UserPrincipalName -like '*#EXT#*' }).Count
    $licensedUsers = @($context.Users | Where-Object { $_.AssignedLicenses }).Count
    $archiveMailboxCount = @($context.Mailboxes | Where-Object { $_.ArchiveStatus -and $_.ArchiveStatus -ne 'None' }).Count
    $publicFolderCount = @($context.PublicFolders).Count
    $connectorCount = @($context.MailFlowConnectors).Count
    $hybridDetected = [bool]($context.HybridInfo -and (($context.HybridInfo.IsHybridConfigured -eq $true) -or ($context.HybridInfo.MigrationEndpointCount -gt 0) -or ($context.HybridInfo.EvidenceCount -gt 0)))
    $crossTenantPartnerCount = if ($context.FederationCrossTenant) { [int]$context.FederationCrossTenant.PartnerCount } else { 0 }
    $teamsCollected = $TenantStatsHash.ContainsKey('AllTeams')
    $paidLicenseAnalysis = if ($context.Licenses.Count -gt 0) { Get-LicenseAnalysis -Licenses $context.Licenses -UserCount $context.Users.Count } else { $null }
    $overallLicenseUtilization = if ($paidLicenseAnalysis -and $paidLicenseAnalysis.TotalPurchased -gt 0) {
        [math]::Round((($paidLicenseAnalysis.TotalConsumed / $paidLicenseAnalysis.TotalPurchased) * 100), 1)
    } else {
        $null
    }
    $voiceSummary = if ($context.TeamsVoice -and $context.TeamsVoice.ContainsKey('Summary')) { $context.TeamsVoice['Summary'] } else { $null }

    Add-MigrationRow -Category 'Domains' -Item 'Verified custom domains' -Status $(if ($unverifiedDomains -gt 0) { 'Blocker' } else { 'Ready' }) -Value "$verifiedDomains verified / $unverifiedDomains unverified" -Notes 'All accepted domains should be validated and sequenced for migration and cutover.' -MigrationAction 'Confirm domain ownership, cutover timing, and accepted domain strategy in the target tenant.' -SourceWorksheet 'Domains'
    Add-MigrationRow -Category 'Domains' -Item 'Mail routing' -Status $(if ($nonM365MxDomains -gt 0) { 'Review' } else { 'Ready' }) -Value "$nonM365MxDomains domain(s) with non-Microsoft 365 MX" -Notes 'Non-M365 MX routing can indicate third-party filtering, staged coexistence, or non-standard cutover requirements.' -MigrationAction 'Document current MX and transport path before migration planning.' -SourceWorksheet 'Domains'
    Add-MigrationRow -Category 'Identity' -Item 'Directory synchronization' -Status $(if ($dirSyncEnabled) { 'Review' } else { 'Ready' }) -Value $(if ($dirSyncEnabled) { 'On-prem sync enabled' } else { 'Cloud-only identity model' }) -Notes 'Hybrid identity affects object authority and user cutover sequencing.' -MigrationAction 'Plan whether identities stay synced during migration or transition to cloud-managed.' -SourceWorksheet 'AdConnectConfiguration'
    Add-MigrationRow -Category 'Identity' -Item 'Guests and external identities' -Status $(if ($guestCount -gt 0) { 'Review' } else { 'Info' }) -Value "$guestCount guest/external user(s)" -Notes 'Guest access usually requires separate planning from member user migration.' -MigrationAction 'Decide whether guest objects are recreated, invited, or excluded from scope.' -SourceWorksheet 'Users'
    Add-MigrationRow -Category 'Messaging' -Item 'Mailbox inventory' -Status 'Info' -Value "$($context.Mailboxes.Count) mailbox(es), $archiveMailboxCount archive-enabled" -Notes 'Mailbox and archive counts drive migration batch sizing and exception planning.' -MigrationAction 'Use mailbox detail sheets to segment batches and identify oversized or special-case mailboxes.' -SourceWorksheet 'MailboxFullDetails'
    Add-MigrationRow -Category 'Messaging' -Item 'Inactive mailboxes' -Status $(if ($context.InactiveMailboxes.Count -gt 0) { 'Review' } else { 'Ready' }) -Value "$($context.InactiveMailboxes.Count) inactive mailbox(es)" -Notes 'Inactive mailboxes may be retained for compliance rather than migrated.' -MigrationAction 'Confirm retention, restore, or exclusion decisions before migration scope is finalized.' -SourceWorksheet 'InactiveMailboxDetails'
    Add-MigrationRow -Category 'Messaging' -Item 'Public folders' -Status $(if ($publicFolderCount -gt 0) { 'Review' } else { 'Ready' }) -Value "$publicFolderCount public folder object(s)" -Notes 'Public folders frequently require separate migration tooling or remediation.' -MigrationAction 'Validate whether public folders remain in scope and determine their target-state strategy.' -SourceWorksheet 'PublicFolderDetails'
    Add-MigrationRow -Category 'Messaging' -Item 'Mail flow dependencies' -Status $(if ($connectorCount -gt 0) { 'Review' } else { 'Ready' }) -Value "$connectorCount connector(s), $($context.RemoteDomains.Count) remote domain(s)" -Notes 'Connectors and remote domains can indicate coexistence, partner routing, or relay dependencies.' -MigrationAction 'Inventory connectors, relay paths, and remote domains before cutover design.' -SourceWorksheet 'MailFlowConnectors'
    Add-MigrationRow -Category 'Collaboration' -Item 'SharePoint and OneDrive' -Status 'Info' -Value "$($context.SharePoint.Count) SharePoint site(s), $($context.OneDrive.Count) OneDrive site(s)" -Notes 'Collaboration workload size and ownership patterns influence tooling and wave planning.' -MigrationAction 'Use site inventory, size, and owner data to prioritize migration waves.' -SourceWorksheet 'SharePoint / OneDrive'
    Add-MigrationRow -Category 'Collaboration' -Item 'Teams workload data' -Status $(if ($teamsCollected) { 'Info' } else { 'Needs Data' }) -Value $(if ($teamsCollected) { "$($context.Teams.Count) team(s) collected" } else { 'Teams inventory not collected in current run' }) -Notes 'Teams topology may require delegated or expanded app permissions beyond this cert-auth path.' -MigrationAction 'Collect Teams team/channel inventory before finalizing collaboration migration planning.' -SourceWorksheet $(if ($teamsCollected) { 'AllTeams' } else { 'N/A' })
    Add-MigrationRow -Category 'Security' -Item 'Conditional Access and MFA' -Status $(if ($context.ConditionalAccess.Count -gt 0 -or $context.AuthConfig) { 'Info' } else { 'Review' }) -Value "$($context.ConditionalAccess.Count) CA policy/policies; MFA summary collected=$(if ($null -ne $context.MfaRegistrationSummary) { 'Yes' } else { 'No' })" -Notes 'Security controls need parity planning to avoid cutover lockouts.' -MigrationAction 'Map CA, MFA, and authentication controls between source and target tenant.' -SourceWorksheet 'ConditionalAccessPolicies'
    Add-MigrationRow -Category 'Hybrid' -Item 'Hybrid or coexistence indicators' -Status $(if ($hybridDetected) { 'Review' } else { 'Ready' }) -Value $(if ($hybridDetected) { "Hybrid signals detected; migration endpoints=$($context.HybridInfo.MigrationEndpointCount)" } else { 'No hybrid indicators detected' }) -Notes 'Hybrid configuration affects mailbox authority, routing, and migration tooling choices.' -MigrationAction 'Validate whether hybrid remains required during migration or can be removed from scope.' -SourceWorksheet 'HybridConfiguration'
    Add-MigrationRow -Category 'External Access' -Item 'Cross-tenant and B2B settings' -Status $(if ($crossTenantPartnerCount -gt 0) { 'Review' } else { 'Info' }) -Value "$crossTenantPartnerCount partner relationship(s)" -Notes 'Cross-tenant policies may affect coexistence and post-migration collaboration behavior.' -MigrationAction 'Review B2B and cross-tenant access settings as part of coexistence planning.' -SourceWorksheet 'FederationConfiguration'
    Add-MigrationRow -Category 'Licensing' -Item 'Target licensing readiness' -Status $(if ($null -ne $overallLicenseUtilization -and $overallLicenseUtilization -ge 85) { 'Review' } else { 'Info' }) -Value $(if ($null -ne $overallLicenseUtilization) { "$overallLicenseUtilization% utilized; $licensedUsers licensed user(s)" } else { 'License utilization unavailable' }) -Notes 'Target tenant licensing should be aligned before user and workload onboarding.' -MigrationAction 'Review paid SKU utilization and confirm target-tenant licensing for migration scope.' -SourceWorksheet 'LicenseSKUs'
    Add-MigrationRow -Category 'Teams Voice' -Item 'Voice workload readiness' -Status $(if ($voiceSummary -and $voiceSummary.PSObject.Properties['DataSource'] -and $voiceSummary.DataSource -eq 'GraphLicenseInference') { 'Needs Data' } else { 'Info' }) -Value $(if ($voiceSummary) { "Voice users=$($voiceSummary.VoiceUserCount); source=$($voiceSummary.DataSource)" } else { 'No Teams voice summary collected' }) -Notes 'Current app-auth path infers voice licensing but does not capture full PSTN or number-assignment state.' -MigrationAction 'Add Teams voice/call record permissions or collect delegated Teams PowerShell data before final voice migration planning.' -SourceWorksheet 'TeamsVoice'

    foreach ($migrationRow in $migrationRows) {
        $migrationIndex++
        $TenantStatsHash['MigrationReadiness'][("{0:D3}-{1}" -f $migrationIndex, $migrationRow.Item)] = $migrationRow
    }

    Write-Log -Type INFO -Message "[Update-AssessmentReportTables] Created $($TenantStatsHash['BestPractices'].Count) best-practice summary rows, $($TenantStatsHash['BestPracticeFindings'].Count) detailed findings, and $($TenantStatsHash['MigrationReadiness'].Count) migration readiness rows" -ExportFileLocation $ExportDetails
}

function Update-ConfigurationSummaryTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    if (-not $TenantStatsHash) {
        return
    }

    $TenantStatsHash['TenantInfoSummary'] = @{}
    $TenantStatsHash['SpamFilteringSummary'] = @{}
    $TenantStatsHash['SMTPRelaySummary'] = @{}
    $TenantStatsHash['FederationSummary'] = @{}
    $TenantStatsHash['TeamsVoiceSummary'] = @{}

    if ($TenantStatsHash.ContainsKey('TenantInfo') -and $TenantStatsHash['TenantInfo']) {
        $tenantInfo = $TenantStatsHash['TenantInfo']
        $selfService = $tenantInfo.SelfServicePurchase
        $azureUsage = $tenantInfo.AzureResourceUsage
        $TenantStatsHash['TenantInfoSummary']['Summary'] = [PSCustomObject]@{
            DisplayName             = $tenantInfo.DisplayName
            TenantId                = $tenantInfo.TenantId
            InitialDomain           = $tenantInfo.InitialDomain
            DefaultDomain           = $tenantInfo.DefaultDomain
            Country                 = $tenantInfo.Country
            CountryLetterCode       = $tenantInfo.CountryLetterCode
            PreferredDataLocation   = $tenantInfo.PreferredDataLocation
            MultiGeoEnabled         = $tenantInfo.MultiGeoEnabled
            MultiGeoAllowed         = $(if ($tenantInfo.MultiGeoAllowed) { $tenantInfo.MultiGeoAllowed -join ', ' } else { $null })
            MultiGeoCentral         = $tenantInfo.MultiGeoCentral
            SelfServicePurchase     = $(if ($selfService) { $selfService.Answer } else { $null })
            SelfServicePurchaseNotes = $(if ($selfService) { $selfService.Notes } else { $null })
            AzureResourceUsage      = $(if ($azureUsage) { $azureUsage.Answer } else { $null })
            AzureResourceUsageNotes = $(if ($azureUsage) { $azureUsage.Notes } else { $null })
        }
    }

    if ($TenantStatsHash.ContainsKey('SpamFilteringConfig') -and $TenantStatsHash['SpamFilteringConfig'].ContainsKey('Configuration')) {
        $spamConfig = $TenantStatsHash['SpamFilteringConfig']['Configuration']
        $TenantStatsHash['SpamFilteringSummary']['Summary'] = [PSCustomObject]@{
            Uses3rdPartyFiltering      = $spamConfig.Uses3rdPartyFiltering
            InboundConnectorCount      = $spamConfig.InboundConnectorCount
            OutboundConnectorCount     = $spamConfig.OutboundConnectorCount
            TransportRuleCount         = $spamConfig.TransportRuleCount
            TransportRulesWithTrustedIPs = $spamConfig.TransportRulesWithTrustedIPs
            PotentialSpamFilters       = $(if ($spamConfig.PotentialSpamFilters) { $spamConfig.PotentialSpamFilters -join '; ' } else { $null })
            TransportRuleIndicators    = $(if ($spamConfig.TransportRuleIndicators) { $spamConfig.TransportRuleIndicators -join '; ' } else { $null })
        }
    }

    if ($TenantStatsHash.ContainsKey('SMTPRelayConfig') -and $TenantStatsHash['SMTPRelayConfig'].ContainsKey('Configuration')) {
        $smtpConfig = $TenantStatsHash['SMTPRelayConfig']['Configuration']
        $TenantStatsHash['SMTPRelaySummary']['Summary'] = [PSCustomObject]@{
            SMTPAuthEnabled                  = $smtpConfig.SMTPAuthEnabled
            SMTPAuthUsers                    = $smtpConfig.SMTPAuthUsers
            ConnectorBasedRelay              = $smtpConfig.ConnectorBasedRelay
            RelayConnectorCount              = @($smtpConfig.RelayConnectors).Count
            DirectSendEnabled                = $smtpConfig.DirectSendEnabled
            AuthoritativeDomains             = $(if ($smtpConfig.PSObject.Properties['AuthoritativeDomains']) { $smtpConfig.AuthoritativeDomains } else { $null })
            SmtpClientAuthenticationDisabled = $(if ($smtpConfig.PSObject.Properties['SmtpClientAuthenticationDisabled']) { $smtpConfig.SmtpClientAuthenticationDisabled } else { $null })
        }
    }

    if ($TenantStatsHash.ContainsKey('FederationConfiguration') -and $TenantStatsHash['FederationConfiguration']) {
        $fedConfig = $TenantStatsHash['FederationConfiguration']
        $exchangeFed = if ($fedConfig.ContainsKey('ExchangeFederation')) { $fedConfig['ExchangeFederation'] } else { $null }
        $crossTenant = if ($fedConfig.ContainsKey('CrossTenantAccess')) { $fedConfig['CrossTenantAccess'] } else { $null }
        $externalIds = if ($fedConfig.ContainsKey('ExternalIdentities')) { $fedConfig['ExternalIdentities'] } else { $null }

        $TenantStatsHash['FederationSummary']['Summary'] = [PSCustomObject]@{
            OrganizationRelationshipCount = $(if ($exchangeFed) { $exchangeFed.OrganizationRelationshipCount } else { 0 })
            IntraOrgConnectorCount        = $(if ($exchangeFed) { $exchangeFed.IntraOrgConnectorCount } else { 0 })
            CrossTenantPartnerCount       = $(if ($crossTenant) { $crossTenant.PartnerCount } else { 0 })
            CrossTenantPartners           = $(if ($crossTenant) { $crossTenant.PartnerTenantNames } else { $null })
            B2BManagementPolicyPresent    = $(if ($externalIds) { $externalIds.B2BManagementPolicyPresent } else { $null })
            GuestUserRole                 = $(if ($externalIds) { $externalIds.GuestUserRole } else { $null })
            InvitationsAllowed            = $(if ($externalIds) { $externalIds.InvitationsAllowed } else { $null })
        }
    }

    if ($TenantStatsHash.ContainsKey('TeamsVoice') -and $TenantStatsHash['TeamsVoice'].ContainsKey('Summary')) {
        $TenantStatsHash['TeamsVoiceSummary']['Summary'] = $TenantStatsHash['TeamsVoice']['Summary']
    }
}

function Update-LicenseClassificationMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    if (-not $TenantStatsHash.ContainsKey('LicenseSKUs') -or -not $TenantStatsHash['LicenseSKUs']) {
        return
    }

    $userCount = 0
    if ($TenantStatsHash.ContainsKey('Users') -and $TenantStatsHash['Users']) {
        $userCount = @($TenantStatsHash['Users'].Values).Count
    }

    foreach ($licenseKey in @($TenantStatsHash['LicenseSKUs'].Keys)) {
        $license = $TenantStatsHash['LicenseSKUs'][$licenseKey]
        if (-not $license) { continue }

        $classification = Get-LicenseClassification -License $license -UserCount $userCount
        $license | Add-Member -MemberType NoteProperty -Name IsPaid -Value $classification.IsPaid -Force
        $license | Add-Member -MemberType NoteProperty -Name LicenseClass -Value $classification.LicenseClass -Force
        $license | Add-Member -MemberType NoteProperty -Name LicenseClassificationReason -Value $classification.Reason -Force
    }
}

#endregion

#region HTML Generation Functions

function Get-HtmlStyle {
    <#
    .SYNOPSIS
        Returns complete CSS styling for the report.
    #>
    return @'
<style>
    :root {
        --primary-color: #0078d4;
        --success-color: #107c10;
        --warning-color: #f7630c;
        --danger-color: #d13438;
        --info-color: #00bcf2;
        --bg-light: #faf9f8;
        --bg-white: #ffffff;
        --text-primary: #323130;
        --text-secondary: #605e5c;
        --border-color: #edebe9;
        --shadow: 0 2px 4px rgba(0,0,0,0.1);
        --anchor-offset: 120px;
    }
    
    * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
    }
    
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        line-height: 1.6;
        color: var(--text-primary);
        background: var(--bg-light);
        padding: 0;
        margin: 0;
    }
    
    .container {
        max-width: 1400px;
        margin: 0 auto;
        padding: 20px;
    }
    
    /* Header */
    .report-header {
        background: linear-gradient(135deg, #0078d4 0%, #004e8c 100%);
        color: white;
        padding: 30px;
        border-radius: 8px;
        margin-bottom: 30px;
        box-shadow: var(--shadow);
    }
    
    .report-header h1 {
        font-size: 2em;
        margin-bottom: 10px;
        font-weight: 600;
    }
    
    .report-meta {
        display: flex;
        gap: 20px;
        flex-wrap: wrap;
        margin-top: 15px;
        font-size: 0.9em;
        opacity: 0.95;
    }
    
    .report-meta-item {
        display: flex;
        align-items: center;
        gap: 5px;
    }
    
    /* Navigation */
    .nav-container {
        position: sticky;
        top: 0;
        background: var(--bg-white);
        border-bottom: 2px solid var(--border-color);
        z-index: 100;
        margin: 0 -20px 30px -20px;
        padding: 0 20px 10px 20px;
    }
    
    .nav {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        overflow-x: auto;
        padding: 10px 0 6px 0;
    }

    .nav-link {
        padding: 8px 16px;
        text-decoration: none;
        color: var(--text-primary);
        border-radius: 4px;
        white-space: nowrap;
        transition: all 0.2s;
        font-size: 0.9em;
        background: #f8f9fb;
        border: 1px solid #e5e7eb;
    }

    .nav-link-primary {
        background: #e8f3ff;
        border-color: #b9dbff;
        font-weight: 600;
    }

    .nav-link:hover {
        background: #eef6ff;
        color: var(--primary-color);
    }

    .nav-group {
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 6px;
        padding: 6px 8px;
        border: 1px solid var(--border-color);
        border-radius: 6px;
        background: #ffffff;
    }

    .nav-group-label {
        font-size: 0.72em;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: var(--text-secondary);
        font-weight: 600;
        padding: 0 4px 0 2px;
    }
    
    /* KPI Cards */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 20px;
        margin-bottom: 30px;
    }
    
    .kpi-card {
        background: var(--bg-white);
        padding: 20px;
        border-radius: 8px;
        box-shadow: var(--shadow);
        border-left: 4px solid var(--primary-color);
    }
    
    .kpi-card.kpi-success { border-left-color: var(--success-color); }
    .kpi-card.kpi-warning { border-left-color: var(--warning-color); }
    .kpi-card.kpi-danger { border-left-color: var(--danger-color); }
    
    .kpi-title {
        font-size: 0.85em;
        color: var(--text-secondary);
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .kpi-value {
        font-size: clamp(1.3rem, 2vw, 2em);
        font-weight: 600;
        color: var(--text-primary);
        margin-bottom: 5px;
        line-height: 1.25;
        overflow-wrap: anywhere;
        font-variant-numeric: tabular-nums;
    }
    
    .kpi-subtitle {
        font-size: 0.85em;
        color: var(--text-secondary);
    }
    
    /* Findings Panel */
    .findings-panel {
        background: var(--bg-white);
        padding: 25px;
        border-radius: 8px;
        box-shadow: var(--shadow);
        margin-bottom: 30px;
    }
    
    .findings-panel h2 {
        font-size: 1.3em;
        margin-bottom: 20px;
        color: var(--text-primary);
    }
    
    .findings-list {
        list-style: none;
    }
    
    .findings-list li {
        padding: 12px 15px;
        margin-bottom: 10px;
        border-radius: 4px;
        border-left: 4px solid;
        background: var(--bg-light);
    }
    
    .findings-list li.finding-risk {
        border-left-color: var(--danger-color);
        background: #fff4f4;
    }
    
    .findings-list li.finding-warning {
        border-left-color: var(--warning-color);
        background: #fffaf4;
    }
    
    .findings-list li.finding-info {
        border-left-color: var(--info-color);
        background: #f4fcff;
    }
    
    .findings-list li a {
        color: inherit;
        text-decoration: none;
        font-weight: 500;
    }
    
    .findings-list li a:hover {
        text-decoration: underline;
    }
    
    /* Sections */
    .section {
        background: var(--bg-white);
        padding: 30px;
        border-radius: 8px;
        box-shadow: var(--shadow);
        margin-bottom: 30px;
        scroll-margin-top: var(--anchor-offset);
    }

    #highlights {
        scroll-margin-top: var(--anchor-offset);
    }
    
    .section h2 {
        font-size: 1.5em;
        color: var(--text-primary);
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 2px solid var(--border-color);
    }

    .section-workload {
        display: inline-block;
        margin-bottom: 10px;
        padding: 4px 10px;
        font-size: 0.74em;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        border-radius: 999px;
        background: #edf7ff;
        color: #0f5a99;
        border: 1px solid #c7e4ff;
        font-weight: 700;
    }
    
    .section-footer {
        margin-top: 25px;
        padding-top: 20px;
        border-top: 1px solid var(--border-color);
    }
    
    .section-footer h3 {
        font-size: 1em;
        color: var(--text-primary);
        margin-bottom: 10px;
    }
    
    .section-footer p {
        color: var(--text-secondary);
        font-size: 0.9em;
        line-height: 1.6;
    }
    
    /* Tables */
    .data-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        font-size: 0.9em;
    }
    
    .data-table thead {
        position: sticky;
        top: 50px;
        background: var(--bg-light);
        z-index: 10;
    }
    
    .data-table th {
        padding: 12px;
        text-align: left;
        font-weight: 600;
        color: var(--text-primary);
        border-bottom: 2px solid var(--border-color);
        white-space: nowrap;
    }
    
    .data-table td {
        padding: 10px 12px;
        border-bottom: 1px solid var(--border-color);
    }

    .wrap-headers th {
        white-space: normal;
        word-break: break-word;
    }

    .wrap-cells td {
        white-space: normal;
        word-break: break-word;
    }
    
    .data-table tbody tr:nth-child(even) {
        background: var(--bg-light);
    }
    
    .data-table tbody tr:hover {
        background: #e1f5fe;
        cursor: pointer;
    }
    
    .data-table td.risk-cell {
        background: #fff4f4;
        color: var(--danger-color);
        font-weight: 600;
    }
    
    /* Badges */
    .badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 0.85em;
        font-weight: 500;
        white-space: nowrap;
    }
    
    .badge-info {
        background: #e1f5fe;
        color: #0277bd;
    }
    
    .badge-warning {
        background: #fff3e0;
        color: #e65100;
    }
    
    .badge-risk {
        background: #ffebee;
        color: #c62828;
    }
    
    .badge-success {
        background: #e8f5e9;
        color: #2e7d32;
    }
    
    /* Callouts */
    .callout {
        padding: 20px;
        border-radius: 6px;
        margin: 20px 0;
        border-left: 4px solid;
    }
    
    .callout-info {
        background: #f4fcff;
        border-left-color: var(--info-color);
    }
    
    .callout-warning {
        background: #fffaf4;
        border-left-color: var(--warning-color);
    }
    
    .callout-danger {
        background: #fff4f4;
        border-left-color: var(--danger-color);
    }
    
    .callout-success {
        background: #f4fff4;
        border-left-color: var(--success-color);
    }
    
    .callout-header {
        font-weight: 600;
        font-size: 1.1em;
        margin-bottom: 10px;
    }
    
    .callout-body {
        color: var(--text-secondary);
    }
    
    /* Charts */
    .chart-container {
        position: relative;
        margin: 30px 0;
        padding: 20px;
        background: var(--bg-light);
        border-radius: 6px;
    }
    
    .chart-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
        gap: 30px;
        margin: 20px 0;
    }
    
    /* Empty State */
    .empty-state {
        text-align: center;
        padding: 60px 20px;
        color: var(--text-secondary);
        font-style: italic;
    }
    
    /* Responsive */
    @media (max-width: 768px) {
        .container {
            padding: 10px;
        }
        
        .kpi-grid {
            grid-template-columns: 1fr;
        }
        
        .report-header h1 {
            font-size: 1.5em;
        }
        
        .nav {
            flex-direction: column;
        }

        .nav-group {
            width: 100%;
        }
        
        .data-table {
            font-size: 0.8em;
        }
        
        .data-table th,
        .data-table td {
            padding: 8px;
        }
    }
    
    /* Print Styles */
    @media print {
        .nav-container,
        .chart-container {
            display: none;
        }
        
        .section {
            page-break-inside: avoid;
        }
    }
</style>
'@
}

function Get-HtmlScript {
    <#
    .SYNOPSIS
        Returns JavaScript for charts and interactivity - MUST BE PLACED AFTER CHART.JS CDN
    #>
    return @'
<script>
    // Chart color palette
    const colors = {
        primary: '#0078d4',
        success: '#107c10',
        warning: '#f7630c',
        danger: '#d13438',
        info: '#00bcf2',
        palette: [
            '#0078d4', '#107c10', '#f7630c', '#d13438', '#00bcf2',
            '#8764b8', '#00b7c3', '#bad80a', '#ff8c00', '#e3008c'
        ]
    };
    
    // Default chart options
    Chart.defaults.font.family = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
    Chart.defaults.plugins.legend.position = 'bottom';
    Chart.defaults.plugins.legend.labels.padding = 15;
    Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(0,0,0,0.8)';
    Chart.defaults.plugins.tooltip.padding = 12;
    Chart.defaults.plugins.tooltip.cornerRadius = 4;
    
    // FUNCTION 1: Create pie chart
    function createPieChart(canvasId, data, labels, title) {
        const ctx = document.getElementById(canvasId);
        if (!ctx) {
            console.error('Canvas not found:', canvasId);
            return;
        }
        
        new Chart(ctx, {
            type: 'pie',
            data: {
                labels: labels,
                datasets: [{
                    data: data,
                    backgroundColor: colors.palette,
                    borderWidth: 2,
                    borderColor: '#ffffff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    title: {
                        display: true,
                        text: title,
                        font: { size: 16, weight: 'bold' },
                        padding: 20
                    },
                    legend: {
                        display: true,
                        position: 'bottom'
                    }
                }
            }
        });
    }
    
    // FUNCTION 2: Create bar chart
    function createBarChart(canvasId, data, labels, title) {
        const ctx = document.getElementById(canvasId);
        if (!ctx) {
            console.error('Canvas not found:', canvasId);
            return;
        }
        
        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: title,
                    data: data,
                    backgroundColor: colors.primary,
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    title: {
                        display: true,
                        text: title,
                        font: { size: 16, weight: 'bold' },
                        padding: 20
                    },
                    legend: {
                        display: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: 'rgba(0,0,0,0.05)'
                        }
                    },
                    x: {
                        grid: {
                            display: false
                        }
                    }
                }
            }
        });
    }
    
    // Smooth scroll for anchor links with sticky-nav offset awareness
    document.addEventListener('DOMContentLoaded', function() {
        function updateAnchorOffset() {
            const navContainer = document.querySelector('.nav-container');
            const navHeight = navContainer ? Math.ceil(navContainer.getBoundingClientRect().height) : 0;
            const offset = Math.max(navHeight + 12, 80);
            document.documentElement.style.setProperty('--anchor-offset', `${offset}px`);
            return offset;
        }

        function getAnchorOffset() {
            const cssValue = getComputedStyle(document.documentElement).getPropertyValue('--anchor-offset').trim();
            const parsed = parseInt(cssValue.replace('px', ''), 10);
            return Number.isFinite(parsed) ? parsed : 120;
        }

        function scrollToHashTarget(hash, behavior) {
            if (!hash || hash.length < 2) {
                return;
            }
            const target = document.querySelector(hash);
            if (!target) {
                return;
            }

            const offset = getAnchorOffset();
            const targetTop = target.getBoundingClientRect().top + window.pageYOffset - offset;
            window.scrollTo({
                top: Math.max(targetTop, 0),
                behavior: behavior || 'smooth'
            });
        }

        updateAnchorOffset();

        document.querySelectorAll('a[href^="#"]').forEach(anchor => {
            anchor.addEventListener('click', function(e) {
                const hash = this.getAttribute('href');
                if (!hash || hash === '#') {
                    return;
                }
                e.preventDefault();
                updateAnchorOffset();
                scrollToHashTarget(hash, 'smooth');
                if (window.history && window.history.pushState) {
                    window.history.pushState(null, '', hash);
                }
            });
        });

        window.addEventListener('resize', function() {
            updateAnchorOffset();
        });

        if (window.location.hash) {
            setTimeout(function() {
                updateAnchorOffset();
                scrollToHashTarget(window.location.hash, 'auto');
            }, 0);
        }
    });
</script>
'@
}

#region Section Builders
# Your proven function (included for reference)
function Convert-ArrayToPieChart {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Array,
        
        [Parameter(Mandatory=$true)]
        [string]$LabelProperty,

        [Parameter(Mandatory=$true)]
        [string]$ValueProperty,

        [string]$ChartTitle = "Pie Chart",
        
        [int]$Width = 300,
        [int]$Height = 300
    )

    $labels = @()
    $data = @()
    
    foreach ($item in $Array) {
        $labels += $item.$LabelProperty
        $data += $item.$ValueProperty
    }

    $labelsJSON = $labels | ConvertTo-Json -Compress
    $dataJSON = $data | ConvertTo-Json -Compress
    $safeTitle = if ([string]::IsNullOrWhiteSpace($ChartTitle)) { "PieChart" } else { ($ChartTitle -replace '[^a-zA-Z0-9_-]', '-') }
    $chartId = "myPieChart-$safeTitle-$([guid]::NewGuid().ToString('N').Substring(0,8))"

    $html = @"
<canvas id="$chartId" style="width:100%; height:${Height}px;"></canvas>
<script>
    var ctx = document.getElementById('$chartId').getContext('2d');
    var myPieChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: $labelsJSON,
            datasets: [{
                data: $dataJSON,
                backgroundColor: [
                    'rgba(255, 99, 132, 0.2)',
                    'rgba(54, 162, 235, 0.2)',
                    'rgba(255, 206, 86, 0.2)',
                    'rgba(75, 192, 192, 0.2)',
                    'rgba(153, 102, 255, 0.2)',
                    'rgba(255, 159, 64, 0.2)'
                ],
                borderColor: [
                    'rgba(255, 99, 132, 1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(255, 206, 86, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(153, 102, 255, 1)',
                    'rgba(255, 159, 64, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                },
                title: {
                    display: true,
                    text: '$ChartTitle'
                }
            }
        },
    });
</script>
"@

    return $html
}

# Updated Build-LicenseSection using your function
function Build-LicenseSection {
    param(
        [array]$Licenses,
        [int]$UserCount = 0
    )
    
    if ($Licenses.Count -eq 0) {
        return "<div class='empty-state'>No license data available</div>"
    }
    
    #Write-Host "  Building License section..." -ForegroundColor Gray
    
    # Get analysis with processed licenses
    $analysis = Get-LicenseAnalysis -Licenses $Licenses -UserCount $UserCount
    $paidLicenses = $analysis.PaidLicenses
    
    # Calculate totals
    $totalPurchased = $analysis.TotalPurchased
    $totalConsumed = $analysis.TotalConsumed
    $totalRemaining = $totalPurchased - $totalConsumed
    
    # Build summary KPIs
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Paid Licenses" -Value (Format-Number $totalPurchased) -Subtitle "Used for utilization only" -Theme 'default'
    $kpiHtml += New-KpiCard -Title "Consumed" -Value (Format-Number $totalConsumed) -Theme 'default'
    $kpiHtml += New-KpiCard -Title "Available" -Value (Format-Number $totalRemaining) -Theme 'default'
    
    $utilPct = if ($totalPurchased -gt 0) {
        ($totalConsumed / $totalPurchased) * 100
    } else { 0 }
    $utilTheme = if ($utilPct -ge 85) { 'warning' } elseif ($utilPct -ge 70) { 'warning' } else { 'success' }
    $kpiHtml += New-KpiCard -Title "Utilization" -Value (Format-Percentage ($utilPct/100) -DecimalPlaces 1) -Theme $utilTheme
    $kpiHtml += "</div>"
    
    # Updated summary message with new categories
    $overCapacityCount = ($paidLicenses | Where-Object { $_.RemainingUnits -lt 0 }).Count
    $atCapacityCount = ($paidLicenses | Where-Object { $_.RemainingUnits -eq 0 }).Count
    $highUtilCount = ($paidLicenses | Where-Object { $_.Utilization -ge 85 -and $_.RemainingUnits -gt 0 }).Count
    
    if ($overCapacityCount -gt 0 -or $atCapacityCount -gt 0 -or $highUtilCount -gt 0) {
        $summaryHtml = "<div style='margin: 20px 0; padding: 15px; background: "
        $summaryHtml += if ($overCapacityCount -gt 0) { '#fff4f4' } elseif ($atCapacityCount -gt 0) { '#fffaf4' } else { '#fff8e1' }
        $summaryHtml += "; border-left: 4px solid "
        $summaryHtml += if ($overCapacityCount -gt 0) { '#d13438' } elseif ($atCapacityCount -gt 0) { '#f7630c' } else { '#ffa726' }
        $summaryHtml += "; border-radius: 4px;'>"
        
        if ($overCapacityCount -gt 0) {
            $summaryHtml += "<strong>🔴 CRITICAL: $overCapacityCount license(s) OVER capacity</strong> - You are consuming more licenses than purchased!<br>"
        }
        if ($atCapacityCount -gt 0) {
            $summaryHtml += "<strong>⚠️ WARNING: $atCapacityCount license(s) at 0 remaining</strong> - Fully allocated, no licenses available for new users<br>"
        }
        if ($highUtilCount -gt 0) {
            $summaryHtml += "<strong>⚠️ WARNING: $highUtilCount license(s) with high utilization (&gt;85%)</strong> - Running low, plan purchases soon<br>"
        }
        $summaryHtml += "See <a href='#highlights'>Findings of Note</a> for details."
        $summaryHtml += "</div>"
        $kpiHtml += $summaryHtml
    }
    
    # Prepare table data
    $tableData = $paidLicenses | ForEach-Object {
        $licenseName = if ($_.PSObject.Properties['SkuFriendlyName'] -and $_.SkuFriendlyName) {
            $_.SkuFriendlyName
        } else {
            Get-FriendlyProductName -SkuPartNumber $_.SkuPartNumber
        }
        [PSCustomObject]@{
            LicenseName = $licenseName
            AppliesTo = 'User'
            PurchasedUnits = $_.PurchasedUnits
            ConsumedUnits = $_.ConsumedUnits
            RemainingUnits = $_.RemainingUnits
            Utilization = [math]::Round($_.Utilization, 1)
        }
    }
    
    # Build table
    $tableColumns = @('LicenseName', 'AppliesTo', 'PurchasedUnits', 'ConsumedUnits', 'RemainingUnits', 'Utilization')
    $tableHeaders = @{
        'LicenseName' = 'License Type'
        'AppliesTo' = 'Applies To'
        'PurchasedUnits' = 'Purchased'
        'ConsumedUnits' = 'Consumed'
        'RemainingUnits' = 'Remaining'
        'Utilization' = 'Utilization %'
    }
    
    # Updated risk highlighting - differentiate critical vs warning
    $riskColumns = @{
        'RemainingUnits' = { param($val, $row) 
            # Critical (red) if negative, Warning (orange) if 0
            [int]$val -lt 0
        }
        'Utilization' = { param($val, $row)
            [double]$val -ge 85
        }
    }
    
    # Sort by remaining (problems first)
    $sortedData = $tableData | Sort-Object RemainingUnits, @{Expression={$_.ConsumedUnits}; Descending=$true}
    
    $tableHtml = New-HtmlTable -Data $sortedData -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns
    
    # Add custom CSS for warning-level highlighting (0 remaining)
    $tableHtml = @"
<style>
    .warning-cell {
        background: #fffaf4 !important;
        color: #e65100 !important;
        font-weight: 600;
    }
</style>
<script>
    // Highlight cells with 0 remaining in orange (warning)
    document.addEventListener('DOMContentLoaded', function() {
        const table = document.querySelector('.data-table');
        if (table) {
            const rows = table.querySelectorAll('tbody tr');
            rows.forEach(row => {
                const cells = row.querySelectorAll('td');
                const remainingCell = cells[4]; // RemainingUnits column (0-indexed)
                if (remainingCell && remainingCell.textContent.trim() === '0') {
                    remainingCell.classList.add('warning-cell');
                }
            });
        }
    });
</script>
$tableHtml
"@
    
    # Add note about trial/viral
    $allProcessed = $Licenses | ForEach-Object {
        [PSCustomObject]@{
            SKU = if ($_.PSObject.Properties['SkuFriendlyName'] -and $_.SkuFriendlyName) {
                $_.SkuFriendlyName
            } else {
                Get-FriendlyProductName -SkuPartNumber $_.SkuPartNumber
            }
            IsPaid = (Test-IsPaidLicenseSku -License $_ -UserCount $UserCount)
        }
    }
    $trialViralCount = ($allProcessed | Where-Object { -not $_.IsPaid }).Count
    
    if ($trialViralCount -gt 0) {
        $tableHtml += @"
<div style='margin-top: 15px; padding: 10px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;'>
    <strong>ℹ️ Note:</strong> $trialViralCount free, trial, preview, viral, or benefit licenses exist but are <strong>excluded from this table and utilization calculations</strong>. Very large user-based seat pools that greatly exceed tenant user count are also treated as likely freemium/benefit inventory.
</div>
"@
    }
    
    # Charts
    $chartHtml = "<div class='chart-grid'>"
    
    $top10 = $paidLicenses | Where-Object { 
        $_.ConsumedUnits -gt 0 
    } | Sort-Object ConsumedUnits -Descending | Select-Object -First 10
    
    if ($top10.Count -gt 0) {
        $chartHtml += "<div class='chart-container'>"
        $chartHtml += Convert-ArrayToPieChart `
            -Array $top10 `
            -LabelProperty 'SkuFriendlyName' `
            -ValueProperty 'ConsumedUnits' `
            -ChartTitle 'Top 10 Consumed Licenses' `
            -Width 400 `
            -Height 300
        $chartHtml += "</div>"
    }
    
    if ($totalPurchased -gt 0 -and $totalRemaining -ge 0) {
        $utilizationData = @(
            [PSCustomObject]@{ Label = 'Consumed'; Value = $totalConsumed }
            [PSCustomObject]@{ Label = 'Available'; Value = $totalRemaining }
        )
        
        $chartHtml += "<div class='chart-container'>"
        $chartHtml += Convert-ArrayToPieChart `
            -Array $utilizationData `
            -LabelProperty 'Label' `
            -ValueProperty 'Value' `
            -ChartTitle 'License Utilization' `
            -Width 400 `
            -Height 300
        $chartHtml += "</div>"
    }
    
    $chartHtml += "</div>"
    
    # Footer
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>License inventory for <strong>paid licenses only</strong>. Severity levels:</p>
    <ul>
        <li><strong>🔴 Critical (red):</strong> Over capacity - consuming more than purchased (negative remaining)</li>
        <li><strong>⚠️ Warning (orange):</strong> At capacity - 0 remaining, no licenses for new users</li>
        <li><strong>⚠️ Warning (orange):</strong> High utilization - &gt;85% used but some available</li>
    </ul>
    <h3>Recommended Next Steps</h3>
    <p><strong>Critical:</strong> Purchase additional licenses IMMEDIATELY - you're over allocated.<br>
    <strong>At 0 Remaining:</strong> Purchase licenses soon - next user assignment will fail.<br>
    <strong>High Utilization:</strong> Plan purchases for next month.</p>
</div>
"@
    
    return $kpiHtml + $tableHtml + $chartHtml + $footerHtml
}

# Updated Build-RecipientsSection
function Build-RecipientsSection {
    param([array]$Recipients)
    
    if ($Recipients.Count -eq 0) {
        return "<div class='empty-state'>No recipient data available</div>"
    }
    
    # Group by type
    $grouped = $Recipients | Group-Object RecipientTypeDetails | 
        Select-Object @{N='Type';E={$_.Name}}, Count |
        Sort-Object Count -Descending
    
    $tableHtml = New-HtmlTable -Data $grouped -Columns @('Type','Count') -ColumnHeaders @{'Type'='Recipient Type';'Count'='Quantity'}
    
    # Chart using your function
    $chartHtml = ""
    if ($grouped.Count -gt 0) {
        $chartHtml = "<div class='chart-container'>"
        $chartHtml += Convert-ArrayToPieChart `
            -Array $grouped `
            -LabelProperty 'Type' `
            -ValueProperty 'Count' `
            -ChartTitle 'Recipients by Type' `
            -Width 500 `
            -Height 400
        $chartHtml += "</div>"
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows the distribution of all mail-enabled recipient objects in your tenant, including mailboxes, groups, contacts, and resources.</p>
</div>
"@
    
    return $tableHtml + $chartHtml + $footerHtml
}

# Updated Build-MailboxesSection
function Build-MailboxesSection {
    param(
        [array]$Mailboxes,
        [array]$InactiveMailboxes,
        [array]$PublicFolders
    )
    
    $publicFolderCount = if ($PublicFolders) { $PublicFolders.Count } else { 0 }
    $publicFolderMailEnabled = if ($PublicFolders) { ($PublicFolders | Where-Object { $_.MailEnabled -eq $true }).Count } else { 0 }
    $publicFolderHasSubfolders = if ($PublicFolders) { ($PublicFolders | Where-Object { $_.HasSubfolders -eq $true }).Count } else { 0 }

    if ($Mailboxes.Count -eq 0 -and $publicFolderCount -eq 0) {
        return "<div class='empty-state'>No mailbox data available</div>"
    }
    
    $analysis = Get-MailboxAnalysis -Mailboxes $Mailboxes
    
    # Group by type and calculate stats
    $summary = $Mailboxes | Group-Object RecipientTypeDetails | ForEach-Object {
        $mbxs = $_.Group
        
        # Filter out N/A values and convert to double
        $validMbxSizes = $mbxs | ForEach-Object { 
            $size = $_.MBXSizeGB
            if ($size -ne 'N/A' -and $null -ne $size -and $size -ne '') {
                try { [double]$size } catch { $null }
            }
        } | Where-Object { $_ -ne $null }
        
        $validArchiveSizes = $mbxs | ForEach-Object { 
            $size = $_.ArchiveSizeGB
            if ($size -ne 'N/A' -and $null -ne $size -and $size -ne '') {
                try { [double]$size } catch { $null }
            }
        } | Where-Object { $_ -ne $null }
        
        $totalSize = ($validMbxSizes | Measure-Object -Sum).Sum
        $totalArchive = ($validArchiveSizes | Measure-Object -Sum).Sum
        $avgSize = if ($validMbxSizes.Count -gt 0) { $totalSize / $validMbxSizes.Count } else { 0 }
        $avgArchive = if ($validArchiveSizes.Count -gt 0) { $totalArchive / $validArchiveSizes.Count } else { 0 }
        $over50 = ($validMbxSizes | Where-Object { $_ -gt 50 }).Count
        $largest = ($validMbxSizes | Measure-Object -Maximum).Maximum
        if ($null -eq $largest) { $largest = 0 }
        
        [PSCustomObject]@{
            Type = $_.Name
            Quantity = $mbxs.Count
            TotalMBXSizeGB = [math]::Round($totalSize, 2)
            TotalArchiveSizeGB = [math]::Round($totalArchive, 2)
            AverageMBXSizeGB = [math]::Round($avgSize, 2)
            AverageArchiveSizeGB = [math]::Round($avgArchive, 2)
            Over50GB = $over50
            LargestGB = [math]::Round($largest, 2)
        }
    } | Sort-Object Quantity -Descending
    
    # KPIs
    $kpiHtml = "<div class='kpi-grid'>"
    $totalMbx = ($summary | Measure-Object Quantity -Sum).Sum
    $totalSize = ($summary | Measure-Object TotalMBXSizeGB -Sum).Sum
    $totalArchive = ($summary | Measure-Object TotalArchiveSizeGB -Sum).Sum
    $totalLarge = ($summary | Measure-Object Over50GB -Sum).Sum
    
    $kpiHtml += New-KpiCard -Title "Total Mailboxes" -Value (Format-Number $totalMbx)
    $kpiHtml += New-KpiCard -Title "Total Storage" -Value (Format-DataSize $totalSize)
    $kpiHtml += New-KpiCard -Title "Total Archive Storage" -Value (Format-DataSize $totalArchive)
    
    $avgMbxSize = if ($totalMbx -gt 0) { $totalSize / $totalMbx } else { 0 }
    $kpiHtml += New-KpiCard -Title "Average Size" -Value (Format-DataSize $avgMbxSize)
    
    $largeTheme = if ($totalLarge -gt ($totalMbx * 0.1)) { 'warning' } else { 'success' }
    $kpiHtml += New-KpiCard -Title "Over 50 GB" -Value (Format-Number $totalLarge) -Theme $largeTheme
    
    $inactiveCount = if ($InactiveMailboxes) { $InactiveMailboxes.Count } else { 0 }
    $inactiveSize = if ($InactiveMailboxes) {
        ($InactiveMailboxes | Measure-Object MBXSizeGB -Sum).Sum
    } else { 0 }
    $kpiHtml += New-KpiCard -Title "Inactive Mailboxes" -Value (Format-Number $inactiveCount)
    $kpiHtml += New-KpiCard -Title "Inactive Storage" -Value (Format-DataSize $inactiveSize)
    
    $publicFolderTheme = if ($publicFolderCount -gt 0) { 'warning' } else { 'success' }
    $kpiHtml += New-KpiCard -Title "Public Folders" -Value (Format-Number $publicFolderCount) -Theme $publicFolderTheme
    $kpiHtml += New-KpiCard -Title "Mail-Enabled Public Folders" -Value (Format-Number $publicFolderMailEnabled)
    $kpiHtml += "</div>"
    
    # Table
    $tableColumns = @('Type','Quantity','TotalMBXSizeGB','TotalArchiveSizeGB','AverageMBXSizeGB','AverageArchiveSizeGB','Over50GB','LargestGB')
    $tableHeaders = @{
        'Type' = 'Mailbox Type'
        'Quantity' = 'Count'
        'TotalMBXSizeGB' = 'Total Size (GB)'
        'TotalArchiveSizeGB' = 'Total Archive (GB)'
        'AverageMBXSizeGB' = 'Avg Size (GB)'
        'AverageArchiveSizeGB' = 'Avg Archive (GB)'
        'Over50GB' = 'Over 50 GB'
        'LargestGB' = 'Largest (GB)'
    }
    
    $riskColumns = @{
        'Over50GB' = { param($val) [int]$val -gt 0 }
        'LargestGB' = { param($val) [double]$val -gt 50 }
    }
    
    $tableHtml = New-HtmlTable -Data $summary -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns
    
    # Inactive mailbox summary table
    $inactiveSummaryHtml = ""
    if ($InactiveMailboxes -and $InactiveMailboxes.Count -gt 0) {
        $inactiveSummary = $InactiveMailboxes | Group-Object RecipientTypeDetails | ForEach-Object {
            $mbxs = $_.Group
            $totalSize = ($mbxs | Measure-Object MBXSizeGB -Sum).Sum
            $avgSize = if ($mbxs.Count -gt 0) { $totalSize / $mbxs.Count } else { 0 }
            $over50 = ($mbxs | Where-Object { [double]$_.MBXSizeGB -gt 50 }).Count
            
            [PSCustomObject]@{
                Type = $_.Name
                Quantity = $mbxs.Count
                TotalSizeGB = [math]::Round($totalSize, 2)
                AverageSizeGB = [math]::Round($avgSize, 2)
                Over50GB = $over50
            }
        }
        $inactiveTableColumns = @('Type','Quantity','TotalSizeGB','AverageSizeGB','Over50GB')
        $inactiveTableHeaders = @{
            'Type' = 'Type'
            'Quantity' = 'Count'
            'TotalSizeGB' = 'Total Size (GB)'
            'AverageSizeGB' = 'Avg Size (GB)'
            'Over50GB' = 'Over 50 GB'
        }
        $inactiveSummaryHtml = "<h3 style='margin-top:30px;'>Inactive Mailboxes Summary</h3>" +
            (New-HtmlTable -Data $inactiveSummary -Columns $inactiveTableColumns -ColumnHeaders $inactiveTableHeaders)
    }

    # Public folder summary table
    $publicFolderSummaryHtml = ""
    if ($publicFolderCount -gt 0) {
        $publicFolderSummary = @(
            [PSCustomObject]@{
                TotalPublicFolders = $publicFolderCount
                MailEnabled = $publicFolderMailEnabled
                WithSubfolders = $publicFolderHasSubfolders
            }
        )
        $publicFolderTableColumns = @('TotalPublicFolders','MailEnabled','WithSubfolders')
        $publicFolderTableHeaders = @{
            'TotalPublicFolders' = 'Total Public Folders'
            'MailEnabled' = 'Mail-Enabled'
            'WithSubfolders' = 'With Subfolders'
        }
        $publicFolderSummaryHtml = "<h3 style='margin-top:30px;'>Public Folders</h3>" +
            (New-HtmlTable -Data $publicFolderSummary -Columns $publicFolderTableColumns -ColumnHeaders $publicFolderTableHeaders)
    }
    
    # Chart using your function
    $chartHtml = ""
    $chartData = $summary | Where-Object { $_.Quantity -gt 0 }
    if ($chartData.Count -gt 0) {
        $chartHtml = "<div class='chart-container'>"
        $chartHtml += Convert-ArrayToPieChart `
            -Array $chartData `
            -LabelProperty 'Type' `
            -ValueProperty 'Quantity' `
            -ChartTitle 'Mailboxes by Type' `
            -Width 500 `
            -Height 400
        $chartHtml += "</div>"
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows mailbox distribution, archive usage, and inactive mailbox storage. Mailboxes over 50 GB may experience performance issues and should be archived or cleaned.</p>
    <h3>Recommended Next Steps</h3>
    <p>Enable archive mailboxes for users over 50 GB. Review retention policies. Consider implementing auto-expanding archives for heavy users.</p>
</div>
"@
    
    return $kpiHtml + $tableHtml + $inactiveSummaryHtml + $publicFolderSummaryHtml + $chartHtml + $footerHtml
}

function Build-EmailActivitySection {
    param(
        [array]$TopSenders,
        [array]$TopReceivers,
        [object]$EmailActivitySummary
    )

    $summaryRecord = $EmailActivitySummary
    if ($summaryRecord -is [array]) {
        $summaryRecord = @($summaryRecord | Select-Object -First 1)
        if ($summaryRecord.Count -gt 0) {
            $summaryRecord = $summaryRecord[0]
        }
        else {
            $summaryRecord = $null
        }
    }

    if ((@($TopSenders).Count -eq 0) -and (@($TopReceivers).Count -eq 0) -and (-not $summaryRecord)) {
        return "<div class='empty-state'>No email activity data available</div>"
    }

    $reportRows = 0
    $activeUsers = 0
    $totalSendCount = 0
    $totalReceiveCount = 0
    $reportPeriod = 'N/A'
    $reportRefreshDate = 'N/A'
    function Get-SummaryValue {
        param(
            [AllowNull()]
            $Record,
            [string]$Key,
            [AllowNull()]
            $Default = $null
        )

        if ($null -eq $Record) {
            return $Default
        }
        if ($Record -is [System.Collections.IDictionary] -and $Record.Contains($Key)) {
            return $Record[$Key]
        }
        if ($Record -is [System.Collections.Specialized.OrderedDictionary] -and $Record.Contains($Key)) {
            return $Record[$Key]
        }
        if ($Record.PSObject -and $Record.PSObject.Properties[$Key]) {
            return $Record.$Key
        }
        return $Default
    }

    if ($summaryRecord) {
        try { $reportRows = [int](Get-SummaryValue -Record $summaryRecord -Key 'ReportRows' -Default 0) } catch { $reportRows = 0 }
        try { $activeUsers = [int](Get-SummaryValue -Record $summaryRecord -Key 'ActiveUsers' -Default 0) } catch { $activeUsers = 0 }
        try { $totalSendCount = [int64](Get-SummaryValue -Record $summaryRecord -Key 'TotalSendCount' -Default 0) } catch { $totalSendCount = 0 }
        try { $totalReceiveCount = [int64](Get-SummaryValue -Record $summaryRecord -Key 'TotalReceiveCount' -Default 0) } catch { $totalReceiveCount = 0 }
        $periodValue = Get-SummaryValue -Record $summaryRecord -Key 'PeriodDuration'
        if (-not [string]::IsNullOrWhiteSpace([string]$periodValue)) { $reportPeriod = [string]$periodValue }
        $refreshValue = Get-SummaryValue -Record $summaryRecord -Key 'ReportRefreshDate'
        if (-not [string]::IsNullOrWhiteSpace([string]$refreshValue)) { $reportRefreshDate = [string]$refreshValue }
    }

    if ($reportRows -eq 0 -and ((@($TopSenders).Count -gt 0) -or (@($TopReceivers).Count -gt 0))) {
        $reportRows = [math]::Max(@($TopSenders).Count, @($TopReceivers).Count)
    }
    if ($activeUsers -eq 0 -and @($TopSenders).Count -gt 0) {
        $activeUsers = @($TopSenders | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.UserPrincipalName) } | Select-Object -ExpandProperty UserPrincipalName -Unique).Count
    }

    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title 'Report Rows' -Value (Format-Number $reportRows)
    $kpiHtml += New-KpiCard -Title 'Active Users' -Value (Format-Number $activeUsers)
    $kpiHtml += New-KpiCard -Title 'Total Sends' -Value (Format-Number $totalSendCount)
    $kpiHtml += New-KpiCard -Title 'Total Receives' -Value (Format-Number $totalReceiveCount)
    $kpiHtml += New-KpiCard -Title 'Report Period' -Value $reportPeriod
    $kpiHtml += New-KpiCard -Title 'Report Refresh' -Value $reportRefreshDate
    $kpiHtml += "</div>"

    $sendersColumns = @('Rank', 'DisplayName', 'UserPrincipalName', 'SendCount', 'ReceiveCount', 'LastActivityDate')
    $sendersHeaders = @{
        Rank              = 'Rank'
        DisplayName       = 'Display Name'
        UserPrincipalName = 'User Principal Name'
        SendCount         = 'Send Count'
        ReceiveCount      = 'Receive Count'
        LastActivityDate  = 'Last Activity'
    }
    $receiversColumns = @('Rank', 'DisplayName', 'UserPrincipalName', 'ReceiveCount', 'SendCount', 'LastActivityDate')
    $receiversHeaders = @{
        Rank              = 'Rank'
        DisplayName       = 'Display Name'
        UserPrincipalName = 'User Principal Name'
        ReceiveCount      = 'Receive Count'
        SendCount         = 'Send Count'
        LastActivityDate  = 'Last Activity'
    }

    $topSendersHtml = "<h3 style='margin-top:30px;'>Top Senders</h3>"
    if (@($TopSenders).Count -gt 0) {
        $topSendersHtml += New-HtmlTable -Data $TopSenders -Columns $sendersColumns -ColumnHeaders $sendersHeaders
    }
    else {
        $topSendersHtml += "<div class='empty-state'>No sender activity rows were returned.</div>"
    }

    $topReceiversHtml = "<h3 style='margin-top:30px;'>Top Receivers</h3>"
    if (@($TopReceivers).Count -gt 0) {
        $topReceiversHtml += New-HtmlTable -Data $TopReceivers -Columns $receiversColumns -ColumnHeaders $receiversHeaders
    }
    else {
        $topReceiversHtml += "<div class='empty-state'>No receiver activity rows were returned.</div>"
    }

    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This section summarizes Microsoft Graph email activity telemetry and highlights the highest-volume senders and receivers in the organization.</p>
</div>
"@

    return $kpiHtml + $topSendersHtml + $topReceiversHtml + $footerHtml
}

function Build-InactiveMailboxesSection {
    param([array]$InactiveMailboxes)
    
    if ($InactiveMailboxes.Count -eq 0) {
        $callout = New-CalloutBox -Type 'success' -Title 'No Inactive Mailboxes' -Content 'Your tenant has no inactive mailboxes, which is good for license optimization.'
        return $callout
    }
    
    # Summary stats
    $summary = $InactiveMailboxes | Group-Object RecipientTypeDetails | ForEach-Object {
        $mbxs = $_.Group
        $totalSize = ($mbxs | Measure-Object MBXSizeGB -Sum).Sum
        $avgSize = if ($mbxs.Count -gt 0) { $totalSize / $mbxs.Count } else { 0 }
        $over50 = ($mbxs | Where-Object { [double]$_.MBXSizeGB -gt 50 }).Count
        
        [PSCustomObject]@{
            Type = $_.Name
            Quantity = $mbxs.Count
            TotalSizeGB = [math]::Round($totalSize, 2)
            AverageSizeGB = [math]::Round($avgSize, 2)
            Over50GB = $over50
        }
    }
    
    # Warning callout
    $calloutHtml = New-CalloutBox -Type 'warning' -Title 'Inactive Mailboxes Detected' -Content "You have $($InactiveMailboxes.Count) inactive mailboxes. These do not consume licenses but do use storage. Consider if they need to be retained for compliance."
    
    # Table
    $tableColumns = @('Type','Quantity','TotalSizeGB','AverageSizeGB','Over50GB')
    $tableHeaders = @{
        'Type' = 'Type'
        'Quantity' = 'Count'
        'TotalSizeGB' = 'Total Size (GB)'
        'AverageSizeGB' = 'Avg Size (GB)'
        'Over50GB' = 'Over 50 GB'
    }
    
    $tableHtml = New-HtmlTable -Data $summary -Columns $tableColumns -ColumnHeaders $tableHeaders
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>Inactive mailboxes are former user mailboxes that have been preserved for compliance. They consume storage but not licenses.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review retention requirements. If mailboxes are no longer needed for legal hold or eDiscovery, consider permanent deletion to free storage.</p>
</div>
"@
    
    return $calloutHtml + $tableHtml + $footerHtml
}

function Build-SharePointOneDriveSection {
    param(
        [array]$SharePointSites,
        [array]$OneDriveSites
    )
    
    # Calculate summaries
    $summaries = @()
    
    # SharePoint (non-O365 connected)
    $spRegular = $SharePointSites | Where-Object { -not $_.IsOffice365GroupsConnected }
    if ($spRegular.Count -gt 0) {
        $totalGB = ($spRegular | Measure-Object StorageUsedGB -Sum).Sum
        $avgGB = $totalGB / $spRegular.Count
        $over1TB = ($spRegular | Where-Object { $_.StorageUsedGB -gt 1024 }).Count
        $largest = ($spRegular | Measure-Object StorageUsedGB -Maximum).Maximum
        
        $summaries += [PSCustomObject]@{
            Category = 'SharePoint Sites'
            Quantity = $spRegular.Count
            TotalStorageGB = [math]::Round($totalGB, 2)
            AverageStorageGB = [math]::Round($avgGB, 2)
            Over1TB = $over1TB
            LargestSizeGB = [math]::Round($largest, 2)
        }
    }
    
    # OneDrive
    if ($OneDriveSites.Count -gt 0) {
        $totalGB = ($OneDriveSites | Measure-Object StorageUsedGB -Sum).Sum
        $avgGB = $totalGB / $OneDriveSites.Count
        $over1TB = ($OneDriveSites | Where-Object { $_.StorageUsedGB -gt 1024 }).Count
        $largest = ($OneDriveSites | Measure-Object StorageUsedGB -Maximum).Maximum
        
        $summaries += [PSCustomObject]@{
            Category = 'OneDrive Sites'
            Quantity = $OneDriveSites.Count
            TotalStorageGB = [math]::Round($totalGB, 2)
            AverageStorageGB = [math]::Round($avgGB, 2)
            Over1TB = $over1TB
            LargestSizeGB = [math]::Round($largest, 2)
        }
    }
    
    # O365 Groups
    $spGroups = $SharePointSites | Where-Object { $_.IsOffice365GroupsConnected -and $_.Template -ne 'TEAMCHANNEL#0' }
    if ($spGroups.Count -gt 0) {
        $totalGB = ($spGroups | Measure-Object StorageUsedGB -Sum).Sum
        $avgGB = $totalGB / $spGroups.Count
        $over1TB = ($spGroups | Where-Object { $_.StorageUsedGB -gt 1024 }).Count
        $largest = ($spGroups | Measure-Object StorageUsedGB -Maximum).Maximum
        
        $summaries += [PSCustomObject]@{
            Category = 'Office 365 Groups'
            Quantity = $spGroups.Count
            TotalStorageGB = [math]::Round($totalGB, 2)
            AverageStorageGB = [math]::Round($avgGB, 2)
            Over1TB = $over1TB
            LargestSizeGB = [math]::Round($largest, 2)
        }
    }
    
    # Teams
    $teams = $SharePointSites | Where-Object { $_.Template -eq 'TEAMCHANNEL#0' }
    if ($teams.Count -gt 0) {
        $totalGB = ($teams | Measure-Object StorageUsedGB -Sum).Sum
        $avgGB = $totalGB / $teams.Count
        $over1TB = ($teams | Where-Object { $_.StorageUsedGB -gt 1024 }).Count
        $largest = ($teams | Measure-Object StorageUsedGB -Maximum).Maximum
        
        $summaries += [PSCustomObject]@{
            Category = 'Teams Sites'
            Quantity = $teams.Count
            TotalStorageGB = [math]::Round($totalGB, 2)
            AverageStorageGB = [math]::Round($avgGB, 2)
            Over1TB = $over1TB
            LargestSizeGB = [math]::Round($largest, 2)
        }
    }
    
    if ($summaries.Count -eq 0) {
        return "<div class='empty-state'>No SharePoint or OneDrive data available</div>"
    }
    
    # KPIs
    $kpiHtml = "<div class='kpi-grid'>"
    $totalSites = ($summaries | Measure-Object Quantity -Sum).Sum
    $totalStorage = ($summaries | Measure-Object TotalStorageGB -Sum).Sum
    
    $kpiHtml += New-KpiCard -Title "Total Sites" -Value (Format-Number $totalSites)
    $kpiHtml += New-KpiCard -Title "Total Storage" -Value (Format-DataSize $totalStorage)
    
    $avgStorage = if ($totalSites -gt 0) { $totalStorage / $totalSites } else { 0 }
    $kpiHtml += New-KpiCard -Title "Average Size" -Value (Format-DataSize $avgStorage)
    $kpiHtml += "</div>"
    
    # Table
    $tableColumns = @('Category','Quantity','TotalStorageGB','AverageStorageGB','Over1TB','LargestSizeGB')
    $tableHeaders = @{
        'Category' = 'Site Type'
        'Quantity' = 'Count'
        'TotalStorageGB' = 'Total Size'
        'AverageStorageGB' = 'Avg Size'
        'Over1TB' = 'Over 1 TB'
        'LargestSizeGB' = 'Largest Site'
    }
    
    $riskColumns = @{
        'Over1TB' = { param($val) [int]$val -gt 0 }
    }

    $valueFormatters = @{
        'Quantity' = { param($val) Format-Number $val }
        'TotalStorageGB' = { param($val) Format-DataSize $val }
        'AverageStorageGB' = { param($val) Format-DataSize $val }
        'Over1TB' = { param($val) Format-Number $val }
        'LargestSizeGB' = { param($val) Format-DataSize $val }
    }
    
    $tableHtml = New-HtmlTable -Data $summaries -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns -ValueFormatters $valueFormatters
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows storage consumption across SharePoint, OneDrive, and Teams. Sites over 1 TB may have performance impacts.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review large sites for archived content that can be moved to cold storage. Implement retention policies and educate users on storage best practices.</p>
</div>
"@
    
    return $kpiHtml + $tableHtml + $footerHtml
}

function Build-OwnershipGovernanceSection {
    param(
        [array]$UnmanagedObjects,
        [array]$OneDriveOwnerMismatches,
        [object]$OwnershipGovernanceSummary
    )

    $unmanagedCount = @($UnmanagedObjects).Count
    $mismatchCount = @($OneDriveOwnerMismatches).Count

    if ($unmanagedCount -eq 0 -and $mismatchCount -eq 0 -and -not $OwnershipGovernanceSummary) {
        return "<div class='empty-state'>No ownership governance data available</div>"
    }

    $missingOwnerCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['MissingOwnerCount']) {
        [int]$OwnershipGovernanceSummary.MissingOwnerCount
    } else {
        @($UnmanagedObjects | Where-Object { $_.OwnerState -eq 'Missing' }).Count
    }
    $ownerHealthRiskCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['OwnerHealthRiskCount']) {
        [int]$OwnershipGovernanceSummary.OwnerHealthRiskCount
    } else {
        @($UnmanagedObjects | Where-Object { $_.OwnerState -in @('Disabled', 'Stale', 'Mixed') }).Count
    }

    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Unmanaged Objects" -Value (Format-Number $unmanagedCount) -Theme $(if ($unmanagedCount -gt 0) { 'warning' } else { 'success' })
    $kpiHtml += New-KpiCard -Title "Missing Owner" -Value (Format-Number $missingOwnerCount) -Theme $(if ($missingOwnerCount -gt 0) { 'danger' } else { 'success' })
    $kpiHtml += New-KpiCard -Title "Owner Health Risks" -Value (Format-Number $ownerHealthRiskCount) -Theme $(if ($ownerHealthRiskCount -gt 0) { 'warning' } else { 'success' })
    $kpiHtml += New-KpiCard -Title "OneDrive Mismatches" -Value (Format-Number $mismatchCount) -Theme $(if ($mismatchCount -gt 0) { 'warning' } else { 'success' })
    $kpiHtml += "</div>"

    $topUnmanaged = @(
        $UnmanagedObjects |
            Sort-Object @{ Expression = {
                switch ($_.Severity) {
                    'Risk' { 1 }
                    'Warning' { 2 }
                    default { 3 }
                }
            } }, Workload, DisplayName |
            Select-Object -First 12 Workload, ObjectType, DisplayName, CurrentOwner, OwnerState, Reason
    )
    $unmanagedHtml = "<h3>Unmanaged Object Review</h3>" +
        (New-HtmlTable -Data $topUnmanaged -Columns @('Workload','ObjectType','DisplayName','CurrentOwner','OwnerState','Reason') -ColumnHeaders @{
            Workload='Workload'; ObjectType='Object Type'; DisplayName='Object'; CurrentOwner='Current Owner'; OwnerState='Owner State'; Reason='Finding'
        } -RiskColumns @{
            OwnerState = { param($val) [string]$val -in @('Missing','Disabled','Stale','Mixed') }
        })

    $topMismatches = @(
        $OneDriveOwnerMismatches |
            Sort-Object DisplayName, SiteUrl |
            Select-Object -First 12 DisplayName, SiteUrl, CurrentOwner, ExpectedDefaultOwner
    )
    $mismatchHtml = "<h3 style='margin-top:20px;'>OneDrive Owner Mismatch Review</h3>" +
        (New-HtmlTable -Data $topMismatches -Columns @('DisplayName','SiteUrl','CurrentOwner','ExpectedDefaultOwner') -ColumnHeaders @{
            DisplayName='OneDrive'; SiteUrl='Site Url'; CurrentOwner='Current Owner'; ExpectedDefaultOwner='Expected Default Owner'
        })

    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>Objects without a healthy owner increase governance risk and slow migration/workload handoff decisions.</p>
    <h3>Recommended Next Steps</h3>
    <p>Assign accountable owners, reassign from disabled or stale identities, and validate OneDrive owner mismatches with business data stewards.</p>
</div>
"@

    return $kpiHtml + $unmanagedHtml + $mismatchHtml + $footerHtml
}

function Build-TeamsSection {
    param(
        [array]$Teams,
        [array]$Licenses,
        [object]$TeamsVoice,
        [int]$UserCount = 0
    )
    
    if ($Teams.Count -eq 0) {
        return "<div class='empty-state'>No Teams data available</div>"
    }
    
    # Determine Teams Voice availability from service plans
    $voicePlans = @('MCOEV','MCOPSTN1','MCOPSTN2','MCOEV_VIRTUALUSER','MCOEV_DOD','MCOPSTNC')
    $hasVoice = $false
    $voicePlanHits = @()
    $paidLicenses = $Licenses | Where-Object { Test-IsPaidLicenseSku -License $_ -UserCount $UserCount }

    foreach ($lic in $paidLicenses) {
        $plans = @()
        if ($lic.PSObject.Properties['ServicePlans'] -and $lic.ServicePlans) {
            $plans = $lic.ServicePlans -split ','
        }
        foreach ($plan in $plans) {
            $trimmed = $plan.Trim()
            if ($voicePlans -contains $trimmed) {
                $hasVoice = $true
                $voicePlanHits += $trimmed
            }
        }
    }
    $voicePlanHits = $voicePlanHits | Select-Object -Unique
    
    $teamRows = $Teams | ForEach-Object {
        $publicCount = if ($_.PublicChannels) { ((@($_.PublicChannels -split ',')) | Where-Object { $_ -and $_.Trim() -ne '' }).Count } else { 0 }
        $privateCount = if ($_.PrivateChannels) { ((@($_.PrivateChannels -split ',')) | Where-Object { $_ -and $_.Trim() -ne '' }).Count } else { 0 }
        $sharedCount = if ($_.SharedChannels) { ((@($_.SharedChannels -split ',')) | Where-Object { $_ -and $_.Trim() -ne '' }).Count } else { 0 }
        [PSCustomObject]@{
            DisplayName = $_.DisplayName
            Visibility = $_.Visibility
            SiteSizeGB = $_.'SiteSize-GB'
            TotalChannels = $_.TotalChannels
            PublicChannels = $publicCount
            PrivateChannels = $privateCount
            SharedChannels = $sharedCount
        }
    }
    
    $totalTeams = $Teams.Count
    $totalChannels = ($Teams | Measure-Object -Property TotalChannels -Sum).Sum
    
    $voiceValue = if ($hasVoice) { 'Detected' } else { 'Not Detected' }
    $voiceTheme = if ($hasVoice) { 'success' } else { 'warning' }
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Total Teams" -Value (Format-Number $totalTeams)
    $kpiHtml += New-KpiCard -Title "Total Channels" -Value (Format-Number $totalChannels)
    $kpiHtml += New-KpiCard -Title "Teams Voice" -Value $voiceValue -Theme $voiceTheme
    $kpiHtml += "</div>"
    
    if ($voicePlanHits.Count -gt 0) {
        $kpiHtml += "<div style='margin-top:10px; color:#605e5c; font-size:0.9em;'>Voice-related plans detected: $([string]::Join(', ', $voicePlanHits))</div>"
    }

    $voiceSummaryTable = ""
    if ($TeamsVoice -and $TeamsVoice.Summary) {
        $voiceSummary = $TeamsVoice.Summary
        $voiceRows = @(
            [PSCustomObject]@{ Metric = 'PSTN Total Minutes (Last 30 Days)'; Value = $voiceSummary.PstnTotalMinutes },
            [PSCustomObject]@{ Metric = 'PSTN Total Calls (Last 30 Days)'; Value = $voiceSummary.PstnTotalCalls },
            [PSCustomObject]@{ Metric = 'Calling Policies'; Value = $voiceSummary.CallingPolicyCount },
            [PSCustomObject]@{ Metric = 'Assigned Phone Numbers'; Value = $voiceSummary.PhoneNumberCount },
            [PSCustomObject]@{ Metric = 'Voice-Enabled Users'; Value = $voiceSummary.VoiceUserCount }
        )
        if ($voiceSummary.PSObject.Properties['DataSource'] -and $voiceSummary.DataSource) {
            $voiceRows += [PSCustomObject]@{ Metric = 'Data Source'; Value = $voiceSummary.DataSource }
        }
        if ($voiceSummary.PSObject.Properties['Notes'] -and $voiceSummary.Notes) {
            $voiceRows += [PSCustomObject]@{ Metric = 'Notes'; Value = $voiceSummary.Notes }
        }
        $voiceSummaryTable = "<h3 style='margin-top:20px;'>Teams Voice Summary</h3>" +
            (New-HtmlTable -Data $voiceRows -Columns @('Metric','Value') -ColumnHeaders @{ Metric='Metric'; Value='Value' })
    }
    
    $tableColumns = @('DisplayName','Visibility','SiteSizeGB','TotalChannels','PublicChannels','PrivateChannels','SharedChannels')
    $tableHeaders = @{
        DisplayName = 'Team'
        Visibility = 'Visibility'
        SiteSizeGB = 'Site Size (GB)'
        TotalChannels = 'Channels'
        PublicChannels = 'Public'
        PrivateChannels = 'Private'
        SharedChannels = 'Shared'
    }
    $tableHtml = New-HtmlTable -Data $teamRows -Columns $tableColumns -ColumnHeaders $tableHeaders -CssClass 'data-table wrap-cells'
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This summarizes Teams inventory, channel breakdowns, and Teams Voice license indicators.</p>
    <h3>Recommended Next Steps</h3>
    <p>If Teams Voice is required, confirm Phone System and PSTN add-ons are correctly licensed and configured.</p>
</div>
"@
    
    return $kpiHtml + $voiceSummaryTable + $tableHtml + $footerHtml
}

function Build-DomainsSection {
    param(
        [array]$Domains,
        [object]$SpamFilteringSummary,
        [object]$SMTPRelaySummary,
        [array]$MailFlowConnectors = @()
    )
    
    if ($Domains.Count -eq 0) {
        return "<div class='empty-state'>No domain data available</div>"
    }
    
    $analysis = Get-DomainAnalysis -Domains $Domains -SpamFilteringSummary $SpamFilteringSummary -SMTPRelaySummary $SMTPRelaySummary

    function Convert-MailFlowValueList {
        param($Value)

        if ($null -eq $Value) {
            return @()
        }
        if ($Value -is [string]) {
            return @(
                $Value -split '[,;]' |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }
        if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
            $items = @()
            foreach ($item in $Value) {
                if ($null -eq $item) { continue }
                $itemText = [string]$item
                if (-not [string]::IsNullOrWhiteSpace($itemText)) {
                    $items += $itemText.Trim()
                }
            }
            return @($items)
        }

        $singleText = [string]$Value
        if ([string]::IsNullOrWhiteSpace($singleText)) {
            return @()
        }
        return @($singleText.Trim())
    }

    function Test-ConnectorAppliesToDomain {
        param(
            [object]$Connector,
            [string]$DomainName
        )

        if (-not $Connector -or [string]::IsNullOrWhiteSpace($DomainName)) {
            return $false
        }

        $recipientDomains = Convert-MailFlowValueList -Value $Connector.RecipientDomains
        if ($recipientDomains.Count -eq 0) {
            return $true
        }

        $domainLower = $DomainName.ToLowerInvariant()
        foreach ($recipientDomain in $recipientDomains) {
            $candidate = ([string]$recipientDomain).Trim().ToLowerInvariant()
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            if ($candidate -eq '*') { return $true }
            if ($candidate -eq $domainLower) { return $true }
            if ($candidate -like "*.$domainLower") { return $true }
            if ($candidate -eq "*.$domainLower") { return $true }
        }

        return $false
    }

    function Get-ConnectorEndpointEvidence {
        param([object]$Connector)

        $endpointSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($propertyName in @('SmartHosts', 'SenderIPAddresses', 'EFSkipIPs', 'EFSkipMailGateway', 'TlsDomain')) {
            if (-not $Connector.PSObject.Properties[$propertyName]) { continue }
            foreach ($value in (Convert-MailFlowValueList -Value $Connector.$propertyName)) {
                $null = $endpointSet.Add($value)
            }
        }
        return @($endpointSet)
    }
    
    # Basic domain info
    $tableColumns = @('Domain','Verified','AuthenticationType','DomainType','IsDefault')
    $tableHeaders = @{
        'Domain' = 'Domain Name'
        'Verified' = 'Verified'
        'AuthenticationType' = 'Auth Type'
        'DomainType' = 'Domain Type'
        'IsDefault' = 'Default'
    }
    
    $riskColumns = @{
        'Verified' = { param($val) $val -eq $false }
    }
    
    $tableHtml = "<h3>Domain Configuration</h3>"
    $tableHtml += New-HtmlTable -Data $Domains -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns

    # Mail authentication summary
    $mailEnabledCustomDomains = @(
        $Domains | Where-Object {
            ([string]$_.Domain -notmatch '(?i)\.onmicrosoft\.com$') -and
            ($_.Verified -eq $true) -and
            ($_.Office365MailExchanger -eq $true -or ([int]$_.TotalDomainRecipients -gt 0))
        }
    )

    $mailAuthKpiHtml = ""
    if ($mailEnabledCustomDomains.Count -gt 0) {
        $spfPassing = @($mailEnabledCustomDomains | Where-Object { $_.SpfIncludesM365 -eq $true }).Count
        $dmarcConfigured = @($mailEnabledCustomDomains | Where-Object { $_.DmarcConfigured -eq $true }).Count
        $dmarcEnforced = @($mailEnabledCustomDomains | Where-Object { $_.DmarcConfigured -eq $true -and $_.DmarcPolicy -in @('quarantine', 'reject') }).Count
        $dkimComplete = @($mailEnabledCustomDomains | Where-Object { [int]$_.DkimSelectorsConfigured -ge 2 }).Count

        $mailAuthKpiHtml = "<h3 id='domains-mailauth' style='margin-top:30px;'>Mail Authentication & Anti-Spoofing</h3>"
        $mailAuthKpiHtml += "<div class='kpi-grid'>"
        $mailAuthKpiHtml += New-KpiCard -Title "Mail Domains" -Value (Format-Number $mailEnabledCustomDomains.Count)
        $mailAuthKpiHtml += New-KpiCard -Title "SPF Coverage" -Value ("{0}%" -f [math]::Round((($spfPassing / $mailEnabledCustomDomains.Count) * 100), 1))
        $mailAuthKpiHtml += New-KpiCard -Title "DKIM (2 Selectors)" -Value ("{0}%" -f [math]::Round((($dkimComplete / $mailEnabledCustomDomains.Count) * 100), 1))
        $mailAuthKpiHtml += New-KpiCard -Title "DMARC Enforced" -Value ("{0}%" -f [math]::Round((($dmarcEnforced / $mailEnabledCustomDomains.Count) * 100), 1))
        $mailAuthKpiHtml += "</div>"

        $mailAuthRows = @(
            $mailEnabledCustomDomains | ForEach-Object {
                $spfStatus = if ($_.SpfIncludesM365 -eq $true) { 'Pass' } else { 'Needs Review' }
                $dmarcStatus = if ($_.DmarcConfigured -eq $true) { 'Present' } else { 'Missing' }
                $dkimStatus = if ([int]$_.DkimSelectorsConfigured -ge 2) { 'Pass' } else { 'Needs Review' }
                $dmarcPolicyDisplay = if ($_.DmarcPolicy) { [string]$_.DmarcPolicy } else { 'N/A' }
                $dmarcPctDisplay = if ($_.DmarcPercent) { "$($_.DmarcPercent)%" } else { 'N/A' }

                [PSCustomObject]@{
                    Domain = $_.Domain
                    M365MX = $(if ($_.Office365MailExchanger -eq $true) { 'Yes' } else { 'No' })
                    SPF = $spfStatus
                    SPFMode = $(if ($_.SpfPolicyMode) { $_.SpfPolicyMode } else { 'N/A' })
                    DMARC = $dmarcStatus
                    DMARCPolicy = $dmarcPolicyDisplay
                    DMARCPct = $dmarcPctDisplay
                    DKIM = $dkimStatus
                }
            }
        )

        $mailAuthColumns = @('Domain','M365MX','SPF','SPFMode','DMARC','DMARCPolicy','DMARCPct','DKIM')
        $mailAuthHeaders = @{
            'Domain' = 'Domain'
            'M365MX' = 'M365 MX'
            'SPF' = 'SPF'
            'SPFMode' = 'SPF Mode'
            'DMARC' = 'DMARC'
            'DMARCPolicy' = 'DMARC Policy'
            'DMARCPct' = 'DMARC %'
            'DKIM' = 'DKIM'
        }
        $mailAuthRiskColumns = @{
            'SPF' = { param($val) [string]$val -eq 'Needs Review' }
            'DMARC' = { param($val) [string]$val -eq 'Missing' }
            'DKIM' = { param($val) [string]$val -eq 'Needs Review' }
        }
        $mailAuthKpiHtml += New-HtmlTable -Data $mailAuthRows -Columns $mailAuthColumns -ColumnHeaders $mailAuthHeaders -RiskColumns $mailAuthRiskColumns -CssClass 'data-table wrap-cells'
    }
    
    # DNS info
    $dnsColumns = @('Domain','NSRecords','ARecords','MXRecords','Office365MailExchanger')
    $dnsHeaders = @{
        'Domain' = 'Domain'
        'NSRecords' = 'Name Servers'
        'ARecords' = 'A Records'
        'MXRecords' = 'MX Records'
        'Office365MailExchanger' = 'M365 MX'
    }
    
    $tableHtml += "<h3 id='domains-dns' style='margin-top:30px;'>DNS Configuration</h3>"
    $tableHtml += New-HtmlTable -Data $Domains -Columns $dnsColumns -ColumnHeaders $dnsHeaders -CssClass 'data-table wrap-cells'

    # Anti-spoofing control summary
    $antiSpoofRows = @()
    if ($SpamFilteringSummary -or $SMTPRelaySummary) {
        $trustedBypassCount = 0
        if ($SpamFilteringSummary -and $SpamFilteringSummary.PSObject.Properties['TransportRulesWithTrustedIPs']) {
            try { $trustedBypassCount = [int]$SpamFilteringSummary.TransportRulesWithTrustedIPs } catch { $trustedBypassCount = 0 }
        }
        $uses3rdPartyFiltering = $null
        if ($SpamFilteringSummary -and $SpamFilteringSummary.PSObject.Properties['Uses3rdPartyFiltering']) {
            $rawThirdParty = $SpamFilteringSummary.Uses3rdPartyFiltering
            if ($rawThirdParty -is [bool]) {
                $uses3rdPartyFiltering = $rawThirdParty
            } else {
                $uses3rdPartyFiltering = ([string]$rawThirdParty -match '(?i)^(yes|true|enabled)$')
            }
        }

        $smtpAuthEnabled = $null
        $smtpAuthUsers = 0
        $smtpClientAuthDisabled = $null
        if ($SMTPRelaySummary) {
            if ($SMTPRelaySummary.PSObject.Properties['SMTPAuthEnabled']) {
                $rawSmtpAuth = $SMTPRelaySummary.SMTPAuthEnabled
                if ($rawSmtpAuth -is [bool]) {
                    $smtpAuthEnabled = $rawSmtpAuth
                } else {
                    $smtpAuthEnabled = ([string]$rawSmtpAuth -match '(?i)^(yes|true|enabled)$')
                }
            }
            if ($SMTPRelaySummary.PSObject.Properties['SMTPAuthUsers']) {
                try { $smtpAuthUsers = [int]$SMTPRelaySummary.SMTPAuthUsers } catch { $smtpAuthUsers = 0 }
            }
            if ($SMTPRelaySummary.PSObject.Properties['SmtpClientAuthenticationDisabled']) {
                $smtpClientRaw = $SMTPRelaySummary.SmtpClientAuthenticationDisabled
                if ($smtpClientRaw -is [bool]) {
                    $smtpClientAuthDisabled = $smtpClientRaw
                } else {
                    $smtpClientAuthDisabled = ([string]$smtpClientRaw -match '(?i)^(yes|true)$')
                }
            }
        }

        $antiSpoofRows += [PSCustomObject]@{
            Signal = 'Trusted IP / bypass indicators'
            Value = $trustedBypassCount
            Assessment = $(if ($trustedBypassCount -gt 0) { 'Review' } else { 'Good' })
        }
        $antiSpoofRows += [PSCustomObject]@{
            Signal = 'SMTP AUTH enabled mailboxes'
            Value = $smtpAuthUsers
            Assessment = $(if ($smtpAuthEnabled -eq $true -and $smtpAuthUsers -gt 0) { 'Review' } else { 'Good' })
        }
        $antiSpoofRows += [PSCustomObject]@{
            Signal = 'Org-wide SMTP client auth disabled'
            Value = $(if ($null -eq $smtpClientAuthDisabled) { 'Unknown' } elseif ($smtpClientAuthDisabled) { 'Yes' } else { 'No' })
            Assessment = $(if ($smtpClientAuthDisabled -eq $false) { 'Review' } else { 'Good' })
        }
        $antiSpoofRows += [PSCustomObject]@{
            Signal = 'Third-party mail filtering detected'
            Value = $(if ($null -eq $uses3rdPartyFiltering) { 'Unknown' } elseif ($uses3rdPartyFiltering) { 'Yes' } else { 'No' })
            Assessment = $(if ($uses3rdPartyFiltering -eq $true) { 'Review' } else { 'Good' })
        }

        $tableHtml += "<h3 id='domains-antispoof' style='margin-top:30px;'>Anti-Spoofing Control Signals</h3>"
        $tableHtml += New-HtmlTable -Data $antiSpoofRows -Columns @('Signal','Value','Assessment') -ColumnHeaders @{ Signal='Control'; Value='Value'; Assessment='Assessment' } -RiskColumns @{ Assessment = { param($val) [string]$val -eq 'Review' } }
    }

    # Estimated inbound/outbound mail paths per domain (best-effort)
    $mailFlowEstimateRows = @()
    $mailRoutingDomains = @(
        $Domains | Where-Object {
            ([string]$_.Domain -notmatch '(?i)\.onmicrosoft\.com$') -and
            ($_.Verified -eq $true) -and
            ($_.Office365MailExchanger -eq $true -or ([int]$_.TotalDomainRecipients -gt 0))
        }
    )
    $allConnectors = @($MailFlowConnectors)
    $inboundConnectors = @($allConnectors | Where-Object { [string]$_.ConnectorDirection -eq 'Inbound' -and $_.Enabled -ne $false })
    $outboundConnectors = @($allConnectors | Where-Object { [string]$_.ConnectorDirection -eq 'Outbound' -and $_.Enabled -ne $false })

    foreach ($domain in $mailRoutingDomains) {
        $domainName = [string]$domain.Domain
        if ([string]::IsNullOrWhiteSpace($domainName)) { continue }

        $domainInboundConnectors = @($inboundConnectors | Where-Object { Test-ConnectorAppliesToDomain -Connector $_ -DomainName $domainName })
        $domainOutboundConnectors = @($outboundConnectors | Where-Object { Test-ConnectorAppliesToDomain -Connector $_ -DomainName $domainName })

        $mxEvidence = Convert-MailFlowValueList -Value $domain.MXRecords
        $inboundEndpointSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $outboundEndpointSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($mx in $mxEvidence) { $null = $inboundEndpointSet.Add($mx) }
        foreach ($connector in $domainInboundConnectors) {
            foreach ($endpoint in (Get-ConnectorEndpointEvidence -Connector $connector)) {
                $null = $inboundEndpointSet.Add($endpoint)
            }
        }

        $usesSmartHostOutbound = $false
        foreach ($connector in $domainOutboundConnectors) {
            $smartHosts = Convert-MailFlowValueList -Value $connector.SmartHosts
            if ($smartHosts.Count -gt 0 -and $connector.UseMXRecord -ne $true) {
                $usesSmartHostOutbound = $true
            }
            foreach ($endpoint in (Get-ConnectorEndpointEvidence -Connector $connector)) {
                $null = $outboundEndpointSet.Add($endpoint)
            }
        }

        if ($outboundEndpointSet.Count -eq 0) {
            if ($domainOutboundConnectors.Count -gt 0) {
                $null = $outboundEndpointSet.Add('Recipient MX (connector-managed)')
            } else {
                $null = $outboundEndpointSet.Add('Recipient MX (Exchange Online default)')
            }
        }
        if ($inboundEndpointSet.Count -eq 0 -and $domain.Office365MailExchanger -eq $true) {
            $null = $inboundEndpointSet.Add('Microsoft 365 MX / EOP')
        }

        $inboundPath = if ($domain.Office365MailExchanger -eq $true) {
            if ($domainInboundConnectors.Count -gt 0) {
                'Internet -> Microsoft 365 MX/EOP -> Inbound connector path(s) -> Exchange Online'
            } else {
                'Internet -> Microsoft 365 MX/EOP -> Exchange Online'
            }
        } elseif ($domainInboundConnectors.Count -gt 0) {
            'Internet/Partner -> Inbound connector path(s) -> Exchange Online'
        } else {
            'External MX or custom path -> Review routing manually'
        }

        $outboundPath = if ($domainOutboundConnectors.Count -gt 0) {
            if ($usesSmartHostOutbound) {
                'Exchange Online -> Outbound connector smarthost(s) -> Destination'
            } else {
                'Exchange Online -> Outbound connector(s) -> Recipient MX'
            }
        } else {
            'Exchange Online -> Recipient MX (default)'
        }

        $confidence = if ($mxEvidence.Count -gt 0 -and ($domainInboundConnectors.Count -gt 0 -or $domainOutboundConnectors.Count -gt 0)) {
            'High'
        } elseif ($mxEvidence.Count -gt 0 -or $domain.Office365MailExchanger -eq $true -or $domainInboundConnectors.Count -gt 0 -or $domainOutboundConnectors.Count -gt 0) {
            'Medium'
        } else {
            'Low'
        }

        $mailFlowEstimateRows += [PSCustomObject]@{
            Domain = $domainName
            InboundPath = $inboundPath
            InboundEndpoints = [string]::Join(', ', @($inboundEndpointSet | Select-Object -First 8))
            OutboundPath = $outboundPath
            OutboundEndpoints = [string]::Join(', ', @($outboundEndpointSet | Select-Object -First 8))
            Confidence = $confidence
        }
    }

    if ($mailFlowEstimateRows.Count -gt 0) {
        $tableHtml += "<h3 id='domains-mailflow' style='margin-top:30px;'>Estimated Domain Mail Flow Paths</h3>"
        $tableHtml += "<p style='margin-top:8px;color:#5f6b74;font-size:13px;'>Best-effort estimate from domain MX metadata plus Exchange connector configuration. Validate with production transport design before cutover.</p>"
        $tableHtml += New-HtmlTable `
            -Data $mailFlowEstimateRows `
            -Columns @('Domain','InboundPath','InboundEndpoints','OutboundPath','OutboundEndpoints','Confidence') `
            -ColumnHeaders @{
                Domain = 'Domain'
                InboundPath = 'Inbound Path'
                InboundEndpoints = 'Inbound Endpoints (Estimated)'
                OutboundPath = 'Outbound Path'
                OutboundEndpoints = 'Outbound Endpoints (Estimated)'
                Confidence = 'Confidence'
            } `
            -RiskColumns @{
                Confidence = { param($val) [string]$val -eq 'Low' }
            } `
            -CssClass 'data-table wrap-cells'
    }
    
    # Recipient counts
    $recipColumns = @('Domain','PrimarySMTPRecipients','AliasOnlyRecipients','TotalDomainRecipients')
    $recipHeaders = @{
        'Domain' = 'Domain'
        'PrimarySMTPRecipients' = 'Primary SMTP'
        'AliasOnlyRecipients' = 'Alias Only'
        'TotalDomainRecipients' = 'Total Domain Recipients'
    }
    
    $tableHtml += "<h3 id='domains-recipients' style='margin-top:30px;'>Recipient Distribution</h3>"
    $tableHtml += New-HtmlTable -Data $Domains -Columns $recipColumns -ColumnHeaders $recipHeaders
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows all domains in your tenant, DNS/mail-auth posture (SPF, DKIM, DMARC), recipient usage, and anti-spoofing control signals.</p>
    <h3>Recommended Next Steps</h3>
    <p>For mail-enabled custom domains, target SPF alignment, two DKIM selectors, and DMARC enforcement (quarantine/reject). Review trusted-IP bypasses and SMTP AUTH exceptions regularly.</p>
</div>
"@
    
    return $tableHtml + $mailAuthKpiHtml + $footerHtml
}

function Build-DevicesSection {
    param([array]$Devices)
    
    if ($Devices.Count -eq 0) {
        return "<div class='empty-state'>No device data available</div>"
    }
    
    $analysis = Get-DeviceAnalysis -Devices $Devices
    
    # Group by OS
    $osSummary = $Devices | Group-Object OperatingSystem | ForEach-Object {
        $osDevices = $_.Group
        $total = $osDevices.Count
        $stale = ($osDevices | Where-Object { $_.DeviceStale -eq $true }).Count
        $managed = ($osDevices | Where-Object { $_.IsManaged -eq $true }).Count
        $compliant = ($osDevices | Where-Object { $_.IsCompliant -eq $true }).Count
        
        [PSCustomObject]@{
            OS = $_.Name
            Total = $total
            AllDevicePercentage = [math]::Round(($total / $Devices.Count) * 100, 1)
            StaleDevicePercentage = [math]::Round(($stale / $total) * 100, 1)
            MDMManagedPercentage = [math]::Round(($managed / $total) * 100, 1)
            IsCompliantPercentage = [math]::Round(($compliant / $total) * 100, 1)
        }
    } | Sort-Object Total -Descending
    
    # KPIs
    $kpiHtml = "<div class='kpi-grid'>"
    $totalDevices = $Devices.Count
    $managedCount = ($Devices | Where-Object { $_.IsManaged -eq $true }).Count
    $compliantCount = ($Devices | Where-Object { $_.IsCompliant -eq $true }).Count
    
    $kpiHtml += New-KpiCard -Title "Total Devices" -Value (Format-Number $totalDevices)
    
    $managedPct = if ($totalDevices -gt 0) { ($managedCount / $totalDevices) * 100 } else { 0 }
    $managedTheme = if ($managedPct -ge 80) { 'success' } elseif ($managedPct -ge 60) { 'warning' } else { 'danger' }
    $managedValue = "{0:N1}%" -f [double]$managedPct
    $kpiHtml += New-KpiCard -Title "MDM Managed" -Value $managedValue -Theme $managedTheme

    $compliantPct = if ($totalDevices -gt 0) { ($compliantCount / $totalDevices) * 100 } else { 0 }
    $compTheme = if ($compliantPct -ge 80) { 'success' } elseif ($compliantPct -ge 60) { 'warning' } else { 'danger' }
    $compliantValue = "{0:N1}%" -f [double]$compliantPct
    $kpiHtml += New-KpiCard -Title "Compliant" -Value $compliantValue -Theme $compTheme
    $kpiHtml += "</div>"
    
    # OS Summary table
    $tableColumns = @('OS','Total','AllDevicePercentage','StaleDevicePercentage','MDMManagedPercentage','IsCompliantPercentage')
    $tableHeaders = @{
        'OS' = 'Operating System'
        'Total' = 'Count'
        'AllDevicePercentage' = '% of All'
        'StaleDevicePercentage' = '% Stale'
        'MDMManagedPercentage' = '% Managed'
        'IsCompliantPercentage' = '% Compliant'
    }
    
    $riskColumns = @{
        'StaleDevicePercentage' = { param($val) [double]$val -gt 20 }
        'IsCompliantPercentage' = { param($val) [double]$val -lt 80 }
    }
    
    $tableHtml = New-HtmlTable -Data $osSummary -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns
    
    # Top OS versions
    $versionHtml = "<h3 style='margin-top:30px;'>Top 5 OS Versions</h3>"
    foreach ($os in ($osSummary | Select-Object -First 3)) {
        $osDevices = $Devices | Where-Object { $_.OperatingSystem -eq $os.OS }
        $topVersions = $osDevices | Group-Object OperatingSystemVersion | 
            Select-Object @{N='Version';E={$_.Name}}, Count |
            Sort-Object Count -Descending |
            Select-Object -First 5
        
        if ($topVersions.Count -gt 0) {
            $versionHtml += "<h4>$($os.OS)</h4>"
            $versionHtml += New-HtmlTable -Data $topVersions -Columns @('Version','Count') -CssClass 'data-table'
        }
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows your device inventory, management status, and compliance. Stale devices (inactive > 6 months) should be reviewed for removal.</p>
    <h3>Recommended Next Steps</h3>
    <p>Enroll unmanaged devices in Intune. Investigate non-compliant devices. Remove or disable stale devices to reduce security risk.</p>
</div>
"@
    
    return $kpiHtml + $tableHtml + $versionHtml + $footerHtml
}

function Build-SecureScoreSection {
    param([array]$SecureScore)
    
    if ($SecureScore.Count -eq 0) {
        return "<div class='empty-state'>No Secure Score data available</div>"
    }
    
    # Use most recent score
    $latest = $SecureScore | Sort-Object CreatedDateTime -Descending | Select-Object -First 1
    
    $currentScore = [int]$latest.CurrentScore
    $maxScore = [int]$latest.MaxScore
    $scorePct = if ($maxScore -gt 0) { ($currentScore / $maxScore) * 100 } else { 0 }
    
    # KPI
    $scoreTheme = if ($scorePct -ge 80) { 'success' } elseif ($scorePct -ge 60) { 'warning' } else { 'danger' }
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Current Score" -Value "$currentScore / $maxScore" -Subtitle "$(Format-Number $scorePct -DecimalPlaces 1)%" -Theme $scoreTheme
    $kpiHtml += New-KpiCard -Title "Enabled Services" -Value $latest.EnabledServicesCount
    $kpiHtml += "</div>"
    
    # Comparison
    $compHtml = "<h3>Score Comparison</h3><div class='kpi-grid'>"
    
    if ($latest.SimilarSizeOrg_ComparitiveScore) {
        $compScore = [int]$latest.SimilarSizeOrg_ComparitiveScore
        $compTheme = if ($currentScore -ge $compScore) { 'success' } else { 'warning' }
        $compHtml += New-KpiCard -Title "Similar Size Orgs" -Value $compScore -Subtitle "Your score: $currentScore" -Theme $compTheme
    }
    
    if ($latest.AllTenants_ComparitiveScore) {
        $allScore = [int]$latest.AllTenants_ComparitiveScore
        $allTheme = if ($currentScore -ge $allScore) { 'success' } else { 'warning' }
        $compHtml += New-KpiCard -Title "All Tenants Average" -Value $allScore -Subtitle "Your score: $currentScore" -Theme $allTheme
    }
    
    $compHtml += "</div>"
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>Microsoft Secure Score measures your security posture. Higher scores indicate better protection against threats. Compare your score to similar organizations.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review top improvement actions in the Microsoft 365 Security Center. Prioritize actions with the highest score impact. Aim for gradual improvement each quarter.</p>
</div>
"@
    
    return $kpiHtml + $compHtml + $footerHtml
}

# Exchange Hybrid
function Build-ExchangeHybridSection {
    param(
        [object]$HybridInfo,
        [object]$ExchangeFederation
    )
    
    if (-not $HybridInfo) {
        return "<div class='empty-state'>No Exchange hybrid data available</div>"
    }
    
    $isHybrid = $HybridInfo.IsHybridConfigured -eq $true
    $partial = (-not $isHybrid) -and (
        ($HybridInfo.InboundOnPremConnectorCount -gt 0) -or
        ($HybridInfo.OutboundOnPremConnectorCount -gt 0) -or
        ($HybridInfo.MigrationEndpointCount -gt 0) -or
        ($HybridInfo.EvidenceCount -gt 0)
    )
    if ($HybridInfo.HybridStatus) {
        $statusText = $HybridInfo.HybridStatus
    } elseif ($isHybrid) {
        $statusText = 'Hybrid'
    } elseif ($partial) {
        $statusText = 'Possible Hybrid'
    } else {
        $statusText = 'None'
    }
    
    $callout = if ($isHybrid) {
        New-CalloutBox -Type 'info' -Title 'Hybrid Configuration Detected' -Content "Exchange hybrid indicators were found. Hybrid type: <strong>$($HybridInfo.HybridType)</strong>."
    } elseif ($partial) {
        New-CalloutBox -Type 'warning' -Title 'Partial Hybrid Indicators' -Content "Some hybrid indicators were found, but a full hybrid configuration could not be confirmed."
    } else {
        New-CalloutBox -Type 'success' -Title 'No Hybrid Detected' -Content "No Exchange hybrid indicators were detected in this tenant."
    }
    
    $statusTextDisplay = $statusText -replace '/', ' / '
    $kpiHtml = "<div class='kpi-grid'>"
    $statusTheme = if ($isHybrid -or $statusText -eq 'Possible Hybrid') { 'warning' } else { 'success' }
    $hybridTypeValue = if ($HybridInfo.HybridType) { $HybridInfo.HybridType } else { "Unknown" }
    $kpiHtml += New-KpiCard -Title "Hybrid Status" -Value $statusTextDisplay -Theme $statusTheme
    $kpiHtml += New-KpiCard -Title "Hybrid Type" -Value $hybridTypeValue
    $kpiHtml += New-KpiCard -Title "Evidence Signals" -Value (Format-Number $HybridInfo.EvidenceCount)
    $kpiHtml += New-KpiCard -Title "On-Prem Connectors" -Value (Format-Number ($HybridInfo.InboundOnPremConnectorCount + $HybridInfo.OutboundOnPremConnectorCount))
    $kpiHtml += "</div>"
    
    $tableData = @(
        [PSCustomObject]@{ Signal = 'IntraOrg Connectors'; Count = $HybridInfo.IntraOrgConnectorCount },
        [PSCustomObject]@{ Signal = 'Organization Relationships'; Count = $HybridInfo.OrgRelationshipCount },
        [PSCustomObject]@{ Signal = 'Inbound On-Prem Connectors'; Count = $HybridInfo.InboundOnPremConnectorCount },
        [PSCustomObject]@{ Signal = 'Outbound On-Prem Connectors'; Count = $HybridInfo.OutboundOnPremConnectorCount },
        [PSCustomObject]@{ Signal = 'Migration Endpoints'; Count = $HybridInfo.MigrationEndpointCount },
        [PSCustomObject]@{ Signal = 'Mail Flow On-Prem Connectors'; Count = $HybridInfo.MailFlowOnPremConnectorCount }
    )
    
    $tableHtml = New-HtmlTable -Data $tableData -Columns @('Signal','Count') -ColumnHeaders @{'Signal'='Signal';'Count'='Count'}

    $endpointRows = @()
    if ($HybridInfo.MigrationEndpoints) {
        $endpointRows += [PSCustomObject]@{ Category = 'Migration Endpoints'; Values = $HybridInfo.MigrationEndpoints }
    }
    if ($HybridInfo.InboundOnPremConnectors) {
        $endpointRows += [PSCustomObject]@{ Category = 'Inbound On-Prem Connectors'; Values = $HybridInfo.InboundOnPremConnectors }
    }
    if ($HybridInfo.OutboundOnPremConnectors) {
        $endpointRows += [PSCustomObject]@{ Category = 'Outbound On-Prem Connectors'; Values = $HybridInfo.OutboundOnPremConnectors }
    }
    # Federation details shown in subsection below to avoid duplication

    $endpointsHtml = ""
    if ($endpointRows.Count -gt 0) {
        $endpointsHtml = "<h3 style='margin-top:20px;'>Endpoints / Objects</h3>"
        $endpointsHtml += New-HtmlTable -Data $endpointRows -Columns @('Category','Values') -ColumnHeaders @{'Category'='Category';'Values'='Values'}
    }
    
    $evidenceHtml = ""
    if ($HybridInfo.Evidence) {
        $evidenceItems = $HybridInfo.Evidence -split '; ' | Where-Object { $_ -and $_.Trim().Length -gt 0 }
        if ($evidenceItems.Count -gt 0) {
            $evidenceHtml = "<h3 style='margin-top:20px;'>Evidence</h3><ul>"
            foreach ($item in $evidenceItems) {
                $encoded = [System.Web.HttpUtility]::HtmlEncode($item)
                $evidenceHtml += "<li>$encoded</li>"
            }
            $evidenceHtml += "</ul>"
        }
    }
    
    $federationHtml = ""
    if ($ExchangeFederation) {
        $fedRows = @(
            [PSCustomObject]@{ Setting = 'Organization Relationships'; Value = $ExchangeFederation.OrganizationRelationships },
            [PSCustomObject]@{ Setting = 'IntraOrg Connectors'; Value = $ExchangeFederation.IntraOrgConnectors }
        )
        $federationHtml = "<h3 style='margin-top:20px;'>Exchange Federation</h3>" +
            (New-HtmlTable -Data $fedRows -Columns @('Setting','Value') -ColumnHeaders @{'Setting'='Setting';'Value'='Value'})
    }

    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>Hybrid indicators show whether Exchange Online is integrated with an on-premises Exchange environment for mail flow, free/busy, or migrations.</p>
    <h3>Recommended Next Steps</h3>
    <p>If hybrid is detected, validate hybrid configuration in the Exchange admin center and confirm connectors align with your intended mail routing. If partial indicators exist, review connectors and organization relationships for stale or incomplete setup.</p>
</div>
"@
    
    return $callout + $kpiHtml + $tableHtml + $endpointsHtml + $evidenceHtml + $federationHtml + $footerHtml
}

function Build-IdentityAdminSection {
    param(
        [array]$Users,
        [array]$Admins,
        [array]$Groups,
        [array]$ExchangeGroups
    )
    
    $totalUsers = $Users.Count
    $memberUsers = ($Users | Where-Object { $_.UserType -eq 'Member' }).Count
    $guestUsers = ($Users | Where-Object { $_.UserType -eq 'Guest' }).Count
    $externalUsers = ($Users | Where-Object { $_.UserPrincipalName -like '*#EXT#*' }).Count
    $externalMembers = ($Users | Where-Object { $_.UserType -eq 'Member' -and $_.UserPrincipalName -like '*#EXT#*' }).Count
    $internalMembers = ($Users | Where-Object { $_.UserType -eq 'Member' -and $_.UserPrincipalName -notlike '*#EXT#*' }).Count
    $groupsCount = $Groups.Count
    $adminCount = $Admins.Count
    
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Total Users" -Value (Format-Number $totalUsers)
    $kpiHtml += New-KpiCard -Title "Members (Internal)" -Value (Format-Number $internalMembers)
    $kpiHtml += New-KpiCard -Title "Guests" -Value (Format-Number $guestUsers)
    $kpiHtml += New-KpiCard -Title "External/B2B" -Value (Format-Number $externalUsers)
    $kpiHtml += New-KpiCard -Title "Groups" -Value (Format-Number $groupsCount)
    $kpiHtml += New-KpiCard -Title "Admins" -Value (Format-Number $adminCount)
    $kpiHtml += "</div>"
    
    $userSummary = @(
        [PSCustomObject]@{ Category = 'Members (Internal)'; Count = $internalMembers },
        [PSCustomObject]@{ Category = 'Guests'; Count = $guestUsers },
        [PSCustomObject]@{ Category = 'Members (External)'; Count = $externalMembers }
    )
    $userTable = "<h3>User Type Summary</h3>" + (New-HtmlTable -Data $userSummary -Columns @('Category','Count') -ColumnHeaders @{'Category'='Category';'Count'='Count'})
    
    $dynamicGroups = ($Groups | Where-Object { $_.MembershipType -eq 'Dynamic Group' }).Count
    $dynamicDistributionGroups = ($ExchangeGroups | Where-Object { $_.RecipientTypeDetails -eq 'DynamicDistributionGroup' }).Count
    $m365Groups = ($Groups | Where-Object { $_.GroupType -eq 'Microsoft 365' }).Count
    $securityGroups = ($Groups | Where-Object { $_.SecurityEnabled -eq $true }).Count
    $onPremGroups = ($Groups | Where-Object { $_.OnPremisesSyncEnabled -eq $true -or $_.Source -eq 'On-Premises' }).Count
    $cloudGroups = ($Groups | Where-Object { $_.OnPremisesSyncEnabled -ne $true -and $_.Source -ne 'On-Premises' }).Count
    $groupSummary = @(
        [PSCustomObject]@{ Category = 'Total groups'; Count = $groupsCount },
        [PSCustomObject]@{ Category = 'Dynamic groups'; Count = $dynamicGroups },
        [PSCustomObject]@{ Category = 'Dynamic distribution groups (mail enabled)'; Count = $dynamicDistributionGroups },
        [PSCustomObject]@{ Category = 'M365 groups'; Count = $m365Groups },
        [PSCustomObject]@{ Category = 'Security groups'; Count = $securityGroups },
        [PSCustomObject]@{ Category = 'Cloud groups'; Count = $cloudGroups },
        [PSCustomObject]@{ Category = 'On-premises groups'; Count = $onPremGroups }
    )
    $groupTable = "<h3 style='margin-top:20px;'>Group Overview</h3>" +
        (New-HtmlTable -Data $groupSummary -Columns @('Category','Count') -ColumnHeaders @{'Category'='Category';'Count'='Count'})
    
    $roleCounts = @{}
    foreach ($admin in $Admins) {
        $roles = @()
        if ($admin.Role) {
            $roles = $admin.Role -split ',\s*'
        }
        foreach ($role in $roles) {
            if (-not $roleCounts.ContainsKey($role)) { $roleCounts[$role] = 0 }
            $roleCounts[$role]++
        }
    }
    $roleSummary = $roleCounts.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{ Role = $_.Key; AdminCount = $_.Value }
    } | Sort-Object AdminCount -Descending
    
    $roleTable = ""
    if ($roleSummary.Count -gt 0) {
        $roleTable = "<h3 style='margin-top:20px;'>Admin Roles Summary</h3>" +
            (New-HtmlTable -Data $roleSummary -Columns @('Role','AdminCount') -ColumnHeaders @{'Role'='Role';'AdminCount'='Admins'})
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This section summarizes user types, external/B2B presence, group inventory, and admin role coverage.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review guest and external users for necessity. Ensure admin roles follow least privilege and are distributed appropriately.</p>
</div>
"@
    
    return $kpiHtml + $userTable + $groupTable + $roleTable + $footerHtml
}

function Build-TenantOverviewSection {
    param(
        [object]$AuthConfig,
        [object]$AdConnect,
        [array]$ConditionalAccess
    )

    $federatedDomains = if ($AuthConfig -and $AuthConfig.FederatedDomains) { $AuthConfig.FederatedDomains } else { @() }
    $signOnProvider = if ($federatedDomains.Count -gt 0) { "Federated (ADFS/3rd-party IdP)" } else { "Entra ID (cloud)" }
    $dirSyncEnabled = if ($AdConnect -and $AdConnect.Summary) { if ($AdConnect.Summary.OnPremisesSyncEnabled -eq $true) { "Yes" } else { "No" } } else { "Unknown" }
    $ssprWriteback = "Not collected"
    $mfaProvider = if ($AuthConfig -and $AuthConfig.MFAEnabled -eq $true) {
        $methods = if ($AuthConfig.MFAMethods -and $AuthConfig.MFAMethods.Count -gt 0) { $AuthConfig.MFAMethods -join ", " } else { "Methods not listed" }
        "Entra ID MFA ($methods)"
    } else { "Not detected" }
    $ssoAppsCount = if ($AuthConfig -and $AuthConfig.SSOApplications) { $AuthConfig.SSOApplications.Count } else { 0 }
    $enterpriseSso = if ($AuthConfig -and $AuthConfig.SSOEnabled -eq $true) { "Yes ($ssoAppsCount apps)" } else { "No" }
    $caCount = if ($ConditionalAccess) { $ConditionalAccess.Count } else { 0 }
    $conditionalAccessAnswer = if ($caCount -gt 0) { "Yes ($caCount policies)" } else { "No" }
    $appProxy = "Not collected"
    $privateAccess = "Not collected"
    $pim = "Not collected"

    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Sign-On Provider" -Value $signOnProvider
    $kpiHtml += New-KpiCard -Title "DirSync Enabled" -Value $dirSyncEnabled -Theme ($(if ($dirSyncEnabled -eq 'Yes') { 'success' } elseif ($dirSyncEnabled -eq 'No') { 'warning' } else { 'default' }))
    $kpiHtml += New-KpiCard -Title "MFA" -Value $(if ($AuthConfig -and $AuthConfig.MFAEnabled) { 'Enabled' } else { 'Not Detected' }) -Theme ($(if ($AuthConfig -and $AuthConfig.MFAEnabled) { 'success' } else { 'warning' }))
    $kpiHtml += New-KpiCard -Title "Conditional Access" -Value $conditionalAccessAnswer -Theme ($(if ($caCount -gt 0) { 'success' } else { 'warning' }))
    $kpiHtml += "</div>"

    $rows = @(
        [PSCustomObject]@{ Topic = "Sign-on provider"; Answer = $signOnProvider; Notes = if ($federatedDomains.Count -gt 0) { "Federated domains: $($federatedDomains -join ', ')" } else { "No federated domains detected" } }
        [PSCustomObject]@{ Topic = "Directory sync from on-prem AD"; Answer = $dirSyncEnabled; Notes = if ($AdConnect -and $AdConnect.Summary) { "Last sync: $($AdConnect.Summary.OnPremisesLastSyncDateTime)" } else { "No sync data available" } }
        [PSCustomObject]@{ Topic = "SSPR / Password writeback"; Answer = $ssprWriteback; Notes = "Not collected in this report" }
        [PSCustomObject]@{ Topic = "MFA provider"; Answer = $mfaProvider; Notes = if ($AuthConfig -and $AuthConfig.MFAMethods) { "Methods: $($AuthConfig.MFAMethods -join ', ')" } else { "No MFA method data found" } }
        [PSCustomObject]@{ Topic = "Enterprise SSO to SaaS apps"; Answer = $enterpriseSso; Notes = if ($ssoAppsCount -gt 0) { "Provide app list spreadsheet (separately)" } else { "No SSO apps detected" } }
        [PSCustomObject]@{ Topic = "Conditional Access"; Answer = $conditionalAccessAnswer; Notes = if ($caCount -gt 0) { "Policies detected in tenant" } else { "No policies detected" } }
        [PSCustomObject]@{ Topic = "Entra App Proxy"; Answer = $appProxy; Notes = "Not collected in this report" }
        [PSCustomObject]@{ Topic = "Entra Private Access"; Answer = $privateAccess; Notes = "Not collected in this report" }
        [PSCustomObject]@{ Topic = "Privileged Identity Management (PIM)"; Answer = $pim; Notes = "Not collected in this report" }
    )

    $tableHtml = New-HtmlTable -Data $rows -Columns @('Topic','Answer','Notes') -ColumnHeaders @{ 'Topic'='Topic'; 'Answer'='Answer'; 'Notes'='Notes' }

    $footerHtml = @"
<div class='section-footer'>
    <h3>Notes</h3>
    <p>This section is intended for Solutions Architect review. Items marked "Not collected" require manual validation.</p>
</div>
"@

    return $kpiHtml + $tableHtml + $footerHtml
}

function Build-ConditionalAccessMfaSection {
    param(
        [array]$ConditionalAccessPolicies,
        [object]$AuthConfig,
        [array]$Users,
        [object]$MfaRegistrationSummary
    )
    
    $totalUsers = $Users.Count
    $enabledPolicies = @($ConditionalAccessPolicies | Where-Object { $_.State -eq 'enabled' })
    $reportOnlyPolicies = @($ConditionalAccessPolicies | Where-Object { $_.State -eq 'reportOnly' })
    $disabledPolicies = @($ConditionalAccessPolicies | Where-Object { $_.State -eq 'disabled' })
    $mfaPolicies = @($enabledPolicies | Where-Object { $_.GrantControls_BuiltInControls -match '(?i)mfa' })
    $mfaEnabled = ($mfaPolicies.Count -gt 0) -or ($AuthConfig -and $AuthConfig.MFAEnabled -eq $true)
    
    $mfaMethods = if ($AuthConfig -and $AuthConfig.MFAMethods) { ($AuthConfig.MFAMethods -join ', ') } else { 'Not available' }
    $passwordlessMethods = if ($AuthConfig -and $AuthConfig.PasswordlessMethods) { ($AuthConfig.PasswordlessMethods -join ', ') } else { 'Not available' }
    
    $mfaAllUsers = $false
    foreach ($policy in $enabledPolicies) {
        $includesAll = ($policy.IncludedUsersCount -eq 'All') -or ($policy.IncludedUsers -eq 'All')
        $excludesNone = ($policy.ExcludedUsersCount -eq 0 -or [string]::IsNullOrWhiteSpace($policy.ExcludedUsers))
        $requiresMfa = ($policy.GrantControls_BuiltInControls -match '(?i)mfa')
        if ($includesAll -and $excludesNone -and $requiresMfa) {
            $mfaAllUsers = $true
            break
        }
    }
    
    $caAllUsers = $false
    foreach ($policy in $enabledPolicies) {
        $includesAll = ($policy.IncludedUsersCount -eq 'All') -or ($policy.IncludedUsers -eq 'All')
        $excludesNone = ($policy.ExcludedUsersCount -eq 0 -or [string]::IsNullOrWhiteSpace($policy.ExcludedUsers))
        if ($includesAll -and $excludesNone) {
            $caAllUsers = $true
            break
        }
    }
    
    $pctNotMfa = if ($mfaAllUsers) { '0%' } else { 'Unknown' }
    $pctNotCA = if ($caAllUsers) { '0%' } else { 'Unknown' }
    
    $mfaEnforcedValue = if ($mfaEnabled) { 'Yes' } else { 'No' }
    $mfaEnforcedTheme = if ($mfaEnabled) { 'success' } else { 'warning' }
    
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "CA Policies" -Value (Format-Number $ConditionalAccessPolicies.Count)
    $kpiHtml += New-KpiCard -Title "Enabled" -Value (Format-Number $enabledPolicies.Count)
    $kpiHtml += New-KpiCard -Title "Report-Only" -Value (Format-Number $reportOnlyPolicies.Count)
    $kpiHtml += New-KpiCard -Title "MFA Enforced" -Value $mfaEnforcedValue -Theme $mfaEnforcedTheme
    $kpiHtml += "</div>"
    
    $summaryRows = @(
        [PSCustomObject]@{ Metric = 'MFA CA Policies Enabled'; Value = $mfaPolicies.Count },
        [PSCustomObject]@{ Metric = 'MFA Methods (Policy)'; Value = $mfaMethods },
        [PSCustomObject]@{ Metric = 'Passwordless Methods'; Value = $passwordlessMethods },
        [PSCustomObject]@{ Metric = '% Users Not Covered by CA'; Value = $pctNotCA },
        [PSCustomObject]@{ Metric = '% Users Not Enforced with MFA'; Value = $pctNotMfa },
        [PSCustomObject]@{ Metric = 'Default User Can Create Apps'; Value = $(if ($AuthConfig -and $AuthConfig.PSObject.Properties['DefaultUserCanCreateApps']) { $AuthConfig.DefaultUserCanCreateApps } else { 'Not available' }) },
        [PSCustomObject]@{ Metric = 'Permission Grant Policies'; Value = $(if ($AuthConfig -and $AuthConfig.PSObject.Properties['PermissionGrantPoliciesAssigned']) { (@($AuthConfig.PermissionGrantPoliciesAssigned) -join ', ') } elseif ($AuthConfig -and $AuthConfig.PSObject.Properties['PermissionGrantPolicies']) { [string]$AuthConfig.PermissionGrantPolicies } else { 'Not available' }) },
        [PSCustomObject]@{ Metric = 'Admin Consent Workflow'; Value = $(if ($AuthConfig -and $AuthConfig.PSObject.Properties['AdminConsentWorkflowEnabled']) { $AuthConfig.AdminConsentWorkflowEnabled } else { 'Not available' }) }
    )

    if ($MfaRegistrationSummary) {
        $summaryRows += [PSCustomObject]@{ Metric = 'MFA Registered Users'; Value = $MfaRegistrationSummary.RegisteredUsers }
        $summaryRows += [PSCustomObject]@{ Metric = 'MFA Not Registered Users'; Value = $MfaRegistrationSummary.NotRegisteredUsers }
        $summaryRows += [PSCustomObject]@{ Metric = 'MFA Registration %'; Value = "$($MfaRegistrationSummary.RegistrationPercent)%" }
    }
    $summaryTable = "<h3>Conditional Access & MFA Summary</h3>" +
        (New-HtmlTable -Data $summaryRows -Columns @('Metric','Value') -ColumnHeaders @{'Metric'='Metric';'Value'='Value'})

    $methodTable = ""
    if ($MfaRegistrationSummary -and $MfaRegistrationSummary.MethodCounts) {
        $methodCounts = $MfaRegistrationSummary.MethodCounts
        $methodRows = @()
        if ($methodCounts -is [hashtable]) {
            $methodRows = $methodCounts.GetEnumerator() | ForEach-Object {
                [PSCustomObject]@{ Method = $_.Key; Count = $_.Value }
            }
        } elseif ($methodCounts -is [pscustomobject]) {
            $methodRows = $methodCounts.PSObject.Properties | ForEach-Object {
                [PSCustomObject]@{ Method = $_.Name; Count = $_.Value }
            }
        }
        $methodRows = $methodRows | Sort-Object Count -Descending
        if ($methodRows.Count -gt 0) {
            $methodTable = "<h3 style='margin-top:20px;'>Registered MFA Methods</h3>" +
                (New-HtmlTable -Data $methodRows -Columns @('Method','Count') -ColumnHeaders @{'Method'='Method';'Count'='Count'})
        }
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This summarizes conditional access policies and MFA enforcement signals. Per-user registration/enforcement data is not collected in this report.</p>
    <h3>Recommended Next Steps</h3>
    <p>Ensure MFA is enforced for all users via Conditional Access, review report-only policies, and validate registration methods in your authentication policy.</p>
</div>
"@
    
    return $kpiHtml + $summaryTable + $methodTable + $footerHtml
}

function Build-AdConnectSection {
    param([object]$AdConnect)
    
    if (-not $AdConnect -or -not $AdConnect.Summary) {
        return "<div class='empty-state'>No AD Connect/Sync data available</div>"
    }
    
    $summary = $AdConnect.Summary
    $syncEnabledFlag = ($summary.OnPremisesSyncEnabled -eq $true)
    $syncEnabled = if ($syncEnabledFlag) { 'Yes' } else { 'No' }
    $lastSyncValue = $summary.OnPremisesLastSyncDateTime
    $lastSyncDate = $null
    if ($lastSyncValue) {
        try {
            $lastSyncDate = [datetime]$lastSyncValue
        }
        catch {
            $lastSyncDate = $null
        }
    }
    $lastSync = if ($lastSyncDate) { $lastSyncDate.ToString('yyyy-MM-dd HH:mm:ss') } elseif ($lastSyncValue) { [string]$lastSyncValue } else { 'N/A' }
    $lastSyncSubtitle = $null
    $lastSyncTheme = 'default'
    if ($syncEnabledFlag -and $lastSyncDate) {
        $ageDays = [math]::Round(((Get-Date).ToUniversalTime() - $lastSyncDate.ToUniversalTime()).TotalDays, 1)
        $lastSyncSubtitle = "$ageDays day(s) ago"
        if ($ageDays -le 1) {
            $lastSyncTheme = 'success'
        }
        elseif ($ageDays -le 7) {
            $lastSyncTheme = 'warning'
        }
        else {
            $lastSyncTheme = 'danger'
        }
    }
    elseif ($syncEnabledFlag) {
        $lastSyncSubtitle = 'Last sync timestamp unavailable'
        $lastSyncTheme = 'warning'
    }
    $errorCount = if ($AdConnect.ErrorCount -gt 0) { $AdConnect.ErrorCount } else { 0 }
    $syncTheme = if ($syncEnabledFlag) { 'success' } else { 'warning' }
    $errorTheme = if ($errorCount -gt 0) { 'warning' } else { 'success' }
    
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "DirSync Enabled" -Value $syncEnabled -Theme $syncTheme
    $kpiHtml += New-KpiCard -Title "Last Sync" -Value $lastSync -Subtitle $lastSyncSubtitle -Theme $lastSyncTheme
    $kpiHtml += New-KpiCard -Title "Sync Errors" -Value (Format-Number $errorCount) -Theme $errorTheme
    $kpiHtml += "</div>"
    
    $serviceTable = ""
    if ($AdConnect.SyncServices -and $AdConnect.SyncServices.Count -gt 0) {
        $serviceTable = "<h3>Sync Service(s)</h3>" +
            (New-HtmlTable -Data $AdConnect.SyncServices -Columns @('ServiceName','ServerName','LastSyncTime') -ColumnHeaders @{'ServiceName'='Service';'ServerName'='Server';'LastSyncTime'='Last Sync'})
    }
    
    $errorTable = ""
    if ($AdConnect.RecentErrors -and $AdConnect.RecentErrors.Count -gt 0) {
        $errorRows = $AdConnect.RecentErrors | ForEach-Object {
            [PSCustomObject]@{
                TimeGenerated = $_.TimeGenerated
                Error = if ($_.Error) { $_.Error } elseif ($_.Reason) { $_.Reason } else { $_.ErrorCode }
                Description = if ($_.Description) { $_.Description } elseif ($_.Message) { $_.Message } else { $null }
            }
        }
        $errorTable = "<h3 style='margin-top:20px;'>Recent Sync Errors</h3>" +
            (New-HtmlTable -Data $errorRows -Columns @('TimeGenerated','Error','Description') -ColumnHeaders @{'TimeGenerated'='Time';'Error'='Error';'Description'='Description'})
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>AD Connect/Sync status indicates whether on-prem directory synchronization is active and healthy.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review sync errors and verify the AD Connect server is healthy. Confirm last sync is recent.</p>
</div>
"@
    
    return $kpiHtml + $serviceTable + $errorTable + $footerHtml
}

function Build-CrossTenantAccessSection {
    param(
        [object]$CrossTenantAccess,
        [object]$ExternalIdentities
    )
    
    if (-not $CrossTenantAccess -and -not $ExternalIdentities) {
        return "<div class='empty-state'>No cross-tenant or external identity data available</div>"
    }
    
    $hasCrossTenant = $CrossTenantAccess -and ($CrossTenantAccess.PartnerCount -gt 0)
    
    $callout = if ($hasCrossTenant) {
        New-CalloutBox -Type 'info' -Title 'Cross-Tenant Access Detected' -Content 'This tenant has cross-tenant access partner settings configured.'
    } else {
        New-CalloutBox -Type 'success' -Title 'No Cross-Tenant Partners Detected' -Content 'No cross-tenant access partners were detected.'
    }
    
    $kpiHtml = "<div class='kpi-grid'>"
    $partnerCount = if ($CrossTenantAccess) { $CrossTenantAccess.PartnerCount } else { 0 }
    $b2bPresent = if ($ExternalIdentities -and $ExternalIdentities.B2BManagementPolicyPresent) { $true } else { $false }
    $b2bValue = if ($b2bPresent) { "Present" } else { "Not Found" }
    $b2bTheme = if ($b2bPresent) { 'success' } else { 'warning' }
    $kpiHtml += New-KpiCard -Title "Cross-Tenant Partners" -Value (Format-Number $partnerCount)
    $kpiHtml += New-KpiCard -Title "B2B Policy" -Value $b2bValue -Theme $b2bTheme
    $kpiHtml += "</div>"
    
    $crossRows = @()
    if ($CrossTenantAccess) {
        $crossRows += [PSCustomObject]@{ Setting = 'Cross-Tenant Access Policy'; Value = $CrossTenantAccess.HasCrossTenantAccessPolicy }
        $crossRows += [PSCustomObject]@{ Setting = 'Partner Count'; Value = $CrossTenantAccess.PartnerCount }
        $crossRows += [PSCustomObject]@{ Setting = 'Default Inbound MFA Accepted'; Value = $CrossTenantAccess.DefaultInboundAccess }
        $crossRows += [PSCustomObject]@{ Setting = 'Default Outbound MFA Accepted'; Value = $CrossTenantAccess.DefaultOutboundAccess }
        $crossRows += [PSCustomObject]@{ Setting = 'Default B2B Direct Connect (Inbound)'; Value = $CrossTenantAccess.DefaultB2BDirectConnectInbound }
        $crossRows += [PSCustomObject]@{ Setting = 'Default B2B Direct Connect (Outbound)'; Value = $CrossTenantAccess.DefaultB2BDirectConnectOutbound }
    }
    $crossHtml = "<h3>Cross-Tenant Access</h3>" + (New-HtmlTable -Data $crossRows -Columns @('Setting','Value') -ColumnHeaders @{'Setting'='Setting';'Value'='Value'})
    
    $partnerDetailRows = @()
    if ($CrossTenantAccess -and $CrossTenantAccess.PartnerTenantDetails) {
        foreach ($tenant in $CrossTenantAccess.PartnerTenantDetails) {
            $partnerDetailRows += [PSCustomObject]@{
                TenantId = $tenant.TenantId
                DisplayName = $tenant.DisplayName
                B2BDirectConnectInbound = $tenant.B2BDirectConnectInbound
                B2BDirectConnectOutbound = $tenant.B2BDirectConnectOutbound
                TrustMfa = $tenant.TrustMfa
                TrustCompliantDevices = $tenant.TrustCompliantDevices
                TrustHybridJoinedDevices = $tenant.TrustHybridJoinedDevices
            }
        }
    }
    $partnerDetailsHtml = ""
    if ($partnerDetailRows.Count -gt 0) {
        $partnerDetailsHtml = "<h3 style='margin-top:20px;'>Partner Tenant Details</h3>" +
            (New-HtmlTable -Data $partnerDetailRows -Columns @('TenantId','DisplayName','B2BDirectConnectInbound','B2BDirectConnectOutbound','TrustMfa','TrustCompliantDevices','TrustHybridJoinedDevices') -ColumnHeaders @{
                'TenantId'='Tenant ID'
                'DisplayName'='Display Name'
                'B2BDirectConnectInbound'='B2B Direct Connect Inbound'
                'B2BDirectConnectOutbound'='B2B Direct Connect Outbound'
                'TrustMfa'='Trust MFA'
                'TrustCompliantDevices'='Trust Compliant Devices'
                'TrustHybridJoinedDevices'='Trust Hybrid Joined Devices'
            } -CssClass 'data-table wrap-headers')
    }
    
    $externalRows = @()
    if ($ExternalIdentities) {
        $externalRows += [PSCustomObject]@{ Setting = 'B2B Management Policy Present'; Value = $ExternalIdentities.B2BManagementPolicyPresent }
        $externalRows += [PSCustomObject]@{ Setting = 'Guest User Role'; Value = $ExternalIdentities.GuestUserRole }
        $externalRows += [PSCustomObject]@{ Setting = 'Invitations Allowed'; Value = $ExternalIdentities.InvitationsAllowed }
    }
    $externalHtml = "<h3 style='margin-top:20px;'>External Identities</h3>" + (New-HtmlTable -Data $externalRows -Columns @('Setting','Value') -ColumnHeaders @{'Setting'='Setting';'Value'='Value'})
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This section summarizes cross-tenant access policies and external identity settings for B2B collaboration.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review partner access rules for least-privilege. Ensure B2B invitation policies align with your security requirements.</p>
</div>
"@
    
    return $callout + $kpiHtml + $crossHtml + $partnerDetailsHtml + $externalHtml + $footerHtml
}

#endregion

#region Main Report Builder

function Build-FindingsPanel {
    param([array]$AllFindings)
    
    if ($AllFindings.Count -eq 0) {
        $content = New-CalloutBox -Type 'success' -Title 'No Issues Found' -Content 'Your tenant configuration looks good! No automatic risk flags were detected.'
        return $content
    }
    
    # Group findings by category
    $grouped = $AllFindings | Group-Object { 
        if ($_.Category) { $_.Category } else { 'Other' }
    }
    
    # Sort groups by priority
    $sortedGroups = $grouped | Sort-Object {
        switch ($_.Name) {
            'Critical' { 1 }
            'At Capacity' { 2 }
            'High Utilization' { 3 }
            'Overall Utilization' { 4 }
            default { 5 }
        }
    }
    
    $html = "<div class='findings-panel'><h2>🔍 Findings of Note</h2>"
    
    # Count by severity
    $riskCount = ($AllFindings | Where-Object { $_.Type -eq 'Risk' }).Count
    $warningCount = ($AllFindings | Where-Object { $_.Type -eq 'Warning' }).Count
    $infoCount = ($AllFindings | Where-Object { $_.Type -eq 'Info' }).Count
    
    # Summary badges
    $html += "<div style='margin-bottom: 20px; display: flex; gap: 10px; flex-wrap: wrap;'>"
    if ($riskCount -gt 0) {
        $html += "<span class='badge badge-risk' style='font-size: 0.9em; padding: 8px 12px;'>🔴 $riskCount Critical</span>"
    }
    if ($warningCount -gt 0) {
        $html += "<span class='badge badge-warning' style='font-size: 0.9em; padding: 8px 12px;'>⚠️ $warningCount Warnings</span>"
    }
    if ($infoCount -gt 0) {
        $html += "<span class='badge badge-info' style='font-size: 0.9em; padding: 8px 12px;'>ℹ️ $infoCount Info</span>"
    }
    $html += "</div>"
    
    # Render findings by category
    foreach ($group in $sortedGroups) {
        $categoryName = $group.Name
        $categoryFindings = $group.Group
        
        # Category header with appropriate icon
        $categoryIcon = switch ($categoryName) {
            'Critical' { '🔴' }
            'At Capacity' { '⚠️' }
            'High Utilization' { '⚠️' }
            'Overall Utilization' { 'ℹ️' }
            default { '📋' }
        }
        
        # Add description for each category
        $categoryDesc = switch ($categoryName) {
            'Critical' { 'Over allocated - consuming more licenses than purchased' }
            'At Capacity' { 'Fully allocated - no licenses available for new users' }
            'High Utilization' { 'Running low - consider purchasing more soon' }
            'Overall Utilization' { 'Tenant-wide license usage summary' }
            default { '' }
        }
        
        $html += "<h3 style='margin-top: 20px; margin-bottom: 5px; color: #333; font-size: 1.1em;'>$categoryIcon $categoryName</h3>"
        if ($categoryDesc) {
            $html += "<p style='margin: 0 0 10px 0; color: #666; font-size: 0.9em; font-style: italic;'>$categoryDesc</p>"
        }
        $html += "<ul class='findings-list'>"
        
        foreach ($finding in $categoryFindings) {
            $class = "finding-$($finding.Type.ToLower())"
            $anchor = if ($finding.Anchor) { "#$($finding.Anchor)" } else { "#" }
            $html += "<li class='$class'><a href='$anchor'>$($finding.Message)</a></li>"
        }
        
        $html += "</ul>"
    }
    
    $html += "</div>"
    return $html
}

function Get-AssessmentSectionWorkload {
    [CmdletBinding()]
    param([string]$SectionId)

    switch ($SectionId) {
        'tenant-overview' { return 'Foundation' }
        'licenses' { return 'Commercial' }
        'domains' { return 'Messaging' }
        'recipients' { return 'Messaging' }
        'mailboxes' { return 'Messaging' }
        'email-activity' { return 'Messaging' }
        'inactive-mailboxes' { return 'Messaging' }
        'exchange-hybrid' { return 'Messaging' }
        'cross-tenant-access' { return 'Messaging' }
        'identity-admins' { return 'Identity & Security' }
        'conditional-access-mfa' { return 'Identity & Security' }
        'devices' { return 'Identity & Security' }
        'secure-score' { return 'Identity & Security' }
        'ad-connect' { return 'Platform' }
        'sharepoint-onedrive' { return 'Collaboration' }
        'ownership-governance' { return 'Collaboration' }
        'teams' { return 'Collaboration' }
        default { return 'Other' }
    }
}

function Build-Navigation {
    param([array]$Sections)

    $groupOrder = @('Foundation', 'Commercial', 'Messaging', 'Identity & Security', 'Collaboration', 'Platform', 'Other')
    $groups = [ordered]@{}
    foreach ($name in $groupOrder) {
        $groups[$name] = @()
    }

    foreach ($section in $Sections) {
        $workload = Get-AssessmentSectionWorkload -SectionId $section.Id
        if (-not $groups.Contains($workload)) {
            $groups[$workload] = @()
        }
        $groups[$workload] += $section
    }

    $html = "<div class='nav-container'><nav class='nav'>"
    $html += "<a href='#highlights' class='nav-link nav-link-primary'>Highlights</a>"
    foreach ($groupName in $groups.Keys) {
        $items = @($groups[$groupName])
        if ($items.Count -eq 0) {
            continue
        }
        $html += "<div class='nav-group'>"
        $html += "<span class='nav-group-label'>$groupName</span>"
        foreach ($section in $items) {
            $html += "<a href='#$($section.Id)' class='nav-link'>$($section.Name)</a>"
        }
        $html += "</div>"
    }

    $html += "</nav></div>"
    return $html
}

function Sort-AssessmentSections {
    [CmdletBinding()]
    param([array]$Sections)

    if (-not $Sections -or $Sections.Count -eq 0) {
        return @()
    }

    $orderedIds = @(
        'tenant-overview',
        'licenses',
        'domains',
        'recipients',
        'mailboxes',
        'email-activity',
        'inactive-mailboxes',
        'exchange-hybrid',
        'cross-tenant-access',
        'identity-admins',
        'conditional-access-mfa',
        'devices',
        'secure-score',
        'sharepoint-onedrive',
        'ownership-governance',
        'teams',
        'ad-connect'
    )

    $rank = @{}
    $idx = 0
    foreach ($id in $orderedIds) {
        $rank[$id] = $idx
        $idx++
    }

    return @(
        $Sections | Sort-Object @{
            Expression = {
                if ($rank.ContainsKey($_.Id)) { $rank[$_.Id] } else { 9999 }
            }
        }, @{
            Expression = { [string]$_.Name }
        }
    )
}

function New-TenantHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [hashtable]$TenantStatsHash = $script:tenantStatsHash,
        
        [Parameter()]
        [string]$OutputPath,
        
        [Parameter()]
        [hashtable]$Thresholds,
        
        [Parameter()]
        [string[]]$IncludeSections,

        [Parameter()]
        [string]$JsonPath,

        [Parameter()]
        [switch]$UseJsonCache
    )
    
    #region Validation
    
    if ($JsonPath -and ($UseJsonCache -or -not $TenantStatsHash -or $TenantStatsHash.Count -eq 0)) {
        $loaded = Import-TenantStatsJson -Path $JsonPath
        if ($loaded) {
            $TenantStatsHash = $loaded
        }
    }

    if (-not $TenantStatsHash -or $TenantStatsHash.Count -eq 0) {
        throw "TenantStatsHash is null or empty. Please ensure data collection has completed successfully or provide -JsonPath."
    }
    
    # Merge thresholds
    if ($Thresholds) {
        foreach ($key in $Thresholds.Keys) {
            $script:DefaultThresholds[$key] = $Thresholds[$key]
        }
    }
    
    #endregion
    
    #region Gather Data
    
    #Write-Host "Building HTML report..." -ForegroundColor Cyan
    #Write-Host "  Validating hashtable structure..." -ForegroundColor Gray
    
    # Extract and normalize data through shared assessment context helper.
    $context = Get-TenantAssessmentContext -TenantStatsHash $TenantStatsHash

    $licenses = $context.Licenses
    Write-Verbose "Licenses: Found $($licenses.Count) items"

    $recipients = $context.Recipients
    Write-Verbose "Recipients: Found $($recipients.Count) items"

    $mailboxes = $context.Mailboxes
    Write-Verbose "Mailboxes: Found $($mailboxes.Count) items"

    $inactiveMailboxes = $context.InactiveMailboxes
    Write-Verbose "Inactive Mailboxes: Found $($inactiveMailboxes.Count) items"

    $publicFolders = $context.PublicFolders
    Write-Verbose "Public Folders: Found $($publicFolders.Count) items"

    $sharepoint = $context.SharePoint
    Write-Verbose "SharePoint: Found $($sharepoint.Count) items"

    $onedrive = $context.OneDrive
    Write-Verbose "OneDrive: Found $($onedrive.Count) items"

    $unmanagedObjects = $context.UnmanagedObjects
    Write-Verbose "Unmanaged objects: Found $($unmanagedObjects.Count) items"

    $oneDriveOwnerMismatches = $context.OneDriveOwnerMismatches
    Write-Verbose "OneDrive owner mismatches: Found $($oneDriveOwnerMismatches.Count) items"

    $ownershipGovernanceSummary = $context.OwnershipGovernanceSummary

    $domains = $context.Domains
    Write-Verbose "Domains: Found $($domains.Count) items"

    $devices = $context.Devices
    Write-Verbose "Devices: Found $($devices.Count) items"

    $secureScore = $context.SecureScore
    Write-Verbose "Secure Score: Found $($secureScore.Count) items"

    $teams = $context.Teams
    Write-Verbose "Teams: Found $($teams.Count) items"
    $teamsVoice = $context.TeamsVoice

    $users = $context.Users
    Write-Verbose "Users: Found $($users.Count) items"

    $admins = $context.Admins
    Write-Verbose "Admins: Found $($admins.Count) items"

    $groups = $context.Groups
    Write-Verbose "Groups: Found $($groups.Count) items"

    $exchangeGroups = $context.ExchangeGroups
    Write-Verbose "Exchange Groups: Found $($exchangeGroups.Count) items"

    $conditionalAccess = $context.ConditionalAccess
    Write-Verbose "Conditional Access: Found $($conditionalAccess.Count) items"

    $authConfig = $context.AuthConfig
    $mfaRegistrationSummary = $context.MfaRegistrationSummary
    $adConnect = $context.AdConnect
    $hybridInfo = $context.HybridInfo
    $federationExchange = $context.FederationExchange
    $federationCrossTenant = $context.FederationCrossTenant
    $federationExternal = $context.FederationExternal
    $spamFilteringSummary = $context.SpamFilteringSummary
    $smtpRelaySummary = $context.SMTPRelaySummary
    $mailFlowConnectors = $context.MailFlowConnectors
    $emailActivitySummary = $context.EmailActivitySummary
    $emailActivityTopSenders = $context.EmailActivityTopSenders
    $emailActivityTopReceivers = $context.EmailActivityTopReceivers
    Write-Verbose "Email activity: TopSenders=$($emailActivityTopSenders.Count), TopReceivers=$($emailActivityTopReceivers.Count)"
    
    Write-Host "  Data extraction complete" -ForegroundColor Green
    
    #endregion
    
    #region Build Sections and Collect ALL Findings
    
    $allFindings = @()
    $sectionContents = @()
    
    # Tenant Overview
    $sectionContents += @{
        Id = 'tenant-overview'
        Name = 'Tenant Overview'
        Content = Build-TenantOverviewSection -AuthConfig $authConfig -AdConnect $adConnect -ConditionalAccess $conditionalAccess
    }

    # Licenses - ALWAYS analyze first for findings
    if ($licenses.Count -gt 0) {
        #Write-Host "  Building Licenses section..." -ForegroundColor Gray
        $licAnalysis = Get-LicenseAnalysis -Licenses $licenses -UserCount $users.Count
        $allFindings += $licAnalysis.Findings
        $sectionContents += @{
            Id = 'licenses'
            Name = 'License Overview'
            Content = Build-LicenseSection -Licenses $licenses -UserCount $users.Count
        }
    }

    # Domains
    if ($domains.Count -gt 0) {
        #Write-Host "  Building Domains section..." -ForegroundColor Gray
        $domainAnalysis = Get-DomainAnalysis -Domains $domains -SpamFilteringSummary $spamFilteringSummary -SMTPRelaySummary $smtpRelaySummary
        $allFindings += $domainAnalysis.Findings
        $sectionContents += @{
            Id = 'domains'
            Name = 'Domain Details'
            Content = Build-DomainsSection -Domains $domains -SpamFilteringSummary $spamFilteringSummary -SMTPRelaySummary $smtpRelaySummary -MailFlowConnectors $mailFlowConnectors
        }
    }
    
    # Recipients
    if ($recipients.Count -gt 0) {
        #Write-Host "  Building Recipients section..." -ForegroundColor Gray
        $sectionContents += @{
            Id = 'recipients'
            Name = 'Recipient Objects'
            Content = Build-RecipientsSection -Recipients $recipients
        }
    }
    
    # Identity & Admins
    if ($users.Count -gt 0 -or $admins.Count -gt 0 -or $groups.Count -gt 0) {
        #Write-Host "  Building Identity & Admins section..." -ForegroundColor Gray
        $identityAnalysis = Get-IdentityAdminAnalysis -Users $users -Admins $admins -Groups $groups -ConditionalAccessPolicies $conditionalAccess
        $allFindings += $identityAnalysis.Findings
        $sectionContents += @{
            Id = 'identity-admins'
            Name = 'Identity & Admins'
            Content = Build-IdentityAdminSection -Users $users -Admins $admins -Groups $groups -ExchangeGroups $exchangeGroups
        }
    }
    
    # Mailboxes
    if ($mailboxes.Count -gt 0 -or $publicFolders.Count -gt 0) {
        #Write-Host "  Building Mailboxes section..." -ForegroundColor Gray
        $mbxAnalysis = Get-MailboxAnalysis -Mailboxes $mailboxes
        $allFindings += $mbxAnalysis.Findings
        $sectionContents += @{
            Id = 'mailboxes'
            Name = 'Mailbox Overview'
            Content = Build-MailboxesSection -Mailboxes $mailboxes -InactiveMailboxes $inactiveMailboxes -PublicFolders $publicFolders
        }
    }

    if ($emailActivitySummary -or $emailActivityTopSenders.Count -gt 0 -or $emailActivityTopReceivers.Count -gt 0) {
        $sectionContents += @{
            Id = 'email-activity'
            Name = 'Email Activity'
            Content = Build-EmailActivitySection -TopSenders $emailActivityTopSenders -TopReceivers $emailActivityTopReceivers -EmailActivitySummary $emailActivitySummary
        }
    }
    
    # Inactive Mailboxes
    if ($inactiveMailboxes.Count -gt 0) {
        #Write-Host "  Building Inactive Mailboxes section..." -ForegroundColor Gray
        $inactiveAnalysis = Get-InactiveMailboxAnalysis -InactiveMailboxes $inactiveMailboxes
        $allFindings += $inactiveAnalysis.Findings
        $sectionContents += @{
            Id = 'inactive-mailboxes'
            Name = 'Inactive Mailboxes'
            Content = Build-InactiveMailboxesSection -InactiveMailboxes $inactiveMailboxes
        }
    }
    
    # SharePoint & OneDrive
    if ($sharepoint.Count -gt 0 -or $onedrive.Count -gt 0) {
        #Write-Host "  Building SharePoint/OneDrive section..." -ForegroundColor Gray
        $spodAnalysis = Get-SharePointOneDriveAnalysis -SharePointSites $sharepoint -OneDriveSites $onedrive
        $allFindings += $spodAnalysis.Findings
        $sectionContents += @{
            Id = 'sharepoint-onedrive'
            Name = 'SharePoint & OneDrive'
            Content = Build-SharePointOneDriveSection -SharePointSites $sharepoint -OneDriveSites $onedrive
        }
    }

    if ($ownershipGovernanceSummary -or $unmanagedObjects.Count -gt 0 -or $oneDriveOwnerMismatches.Count -gt 0) {
        $ownershipAnalysis = Get-OwnershipGovernanceAnalysis `
            -UnmanagedObjects $unmanagedObjects `
            -OneDriveOwnerMismatches $oneDriveOwnerMismatches `
            -OwnershipGovernanceSummary $ownershipGovernanceSummary
        $allFindings += $ownershipAnalysis.Findings
        $sectionContents += @{
            Id = 'ownership-governance'
            Name = 'Ownership Governance'
            Content = Build-OwnershipGovernanceSection -UnmanagedObjects $unmanagedObjects -OneDriveOwnerMismatches $oneDriveOwnerMismatches -OwnershipGovernanceSummary $ownershipGovernanceSummary
        }
    }

    # Teams
    if ($teams.Count -gt 0) {
        #Write-Host "  Building Teams section..." -ForegroundColor Gray
        $sectionContents += @{
            Id = 'teams'
            Name = 'Teams Overview'
            Content = Build-TeamsSection -Teams $teams -Licenses $licenses -TeamsVoice $teamsVoice -UserCount $users.Count
        }
    }
    
    # Devices
    if ($devices.Count -gt 0) {
        #Write-Host "  Building Devices section..." -ForegroundColor Gray
        $deviceAnalysis = Get-DeviceAnalysis -Devices $devices
        $allFindings += $deviceAnalysis.Findings
        $sectionContents += @{
            Id = 'devices'
            Name = 'Device Overview'
            Content = Build-DevicesSection -Devices $devices
        }
    }
    
    # AD Connect / Sync
    if ($adConnect) {
        #Write-Host "  Building AD Connect section..." -ForegroundColor Gray
        $adAnalysis = Get-AdConnectAnalysis -AdConnect $adConnect
        $allFindings += $adAnalysis.Findings
        $sectionContents += @{
            Id = 'ad-connect'
            Name = 'AD Connect / Sync'
            Content = Build-AdConnectSection -AdConnect $adConnect
        }
    }
    
    # Conditional Access & MFA
    if ($conditionalAccess.Count -gt 0 -or $authConfig) {
        #Write-Host "  Building Conditional Access & MFA section..." -ForegroundColor Gray
        $caAnalysis = Get-ConditionalAccessMfaAnalysis -ConditionalAccessPolicies $conditionalAccess -AuthConfig $authConfig -TotalUsers $users.Count
        $allFindings += $caAnalysis.Findings
        $sectionContents += @{
            Id = 'conditional-access-mfa'
            Name = 'Conditional Access & MFA'
            Content = Build-ConditionalAccessMfaSection -ConditionalAccessPolicies $conditionalAccess -AuthConfig $authConfig -Users $users -MfaRegistrationSummary $mfaRegistrationSummary
        }
    }
    
    # Exchange Hybrid (includes federation subsection)
    if ($hybridInfo) {
        #Write-Host "  Building Exchange Hybrid section..." -ForegroundColor Gray
        $hybridAnalysis = Get-ExchangeHybridAnalysis -HybridInfo $hybridInfo
        $fedAnalysis = Get-FederationAnalysis -ExchangeFederation $federationExchange -CrossTenantAccess $federationCrossTenant -ExternalIdentities $federationExternal
        $allFindings += $hybridAnalysis.Findings
        $allFindings += $fedAnalysis.Findings
        $sectionContents += @{
            Id = 'exchange-hybrid'
            Name = 'Exchange Hybrid'
            Content = Build-ExchangeHybridSection -HybridInfo $hybridInfo -ExchangeFederation $federationExchange
        }
    }
    
    if ($federationCrossTenant -or $federationExternal) {
        #Write-Host "  Building Cross-Tenant Access section..." -ForegroundColor Gray
        $sectionContents += @{
            Id = 'cross-tenant-access'
            Name = 'Cross-Tenant Access'
            Content = Build-CrossTenantAccessSection -CrossTenantAccess $federationCrossTenant -ExternalIdentities $federationExternal
        }
    }
    
    # Secure Score
    if ($secureScore.Count -gt 0) {
        #Write-Host "  Building Secure Score section..." -ForegroundColor Gray
        $sectionContents += @{
            Id = 'secure-score'
            Name = 'Secure Score'
            Content = Build-SecureScoreSection -SecureScore $secureScore
        }
    }
    
    #Write-Host "  Total findings collected: $($allFindings.Count)" -ForegroundColor Green
    $sectionContents = Sort-AssessmentSections -Sections $sectionContents
    
    #endregion
    
    #region Build HTML
    
    # Get tenant name
    $tenantName = "Microsoft 365 Tenant"
    try {
        $org = Get-AssessmentTenantOrganization
        if ($org) {
            $tenantName = $org.DisplayName
        }
    } catch {}

    # Determine output path (include tenant name)
    if (-not $OutputPath) {
        $safeName = ($tenantName -replace '[^\w\-. ]', '').Trim()
        if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "TenantReport" }
        if ($global:ExportDetails) {
            $exportDir = Split-Path -Path $global:ExportDetails -Parent
            $OutputPath = Join-Path $exportDir "$safeName-Report.html"
        } else {
            $OutputPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "$safeName-Report.html"
        }
    }
    
    # Ensure .html extension
    if ($OutputPath -notmatch '\.html$') {
        $OutputPath += '.html'
    }
    
    $reportDate = Get-Date -Format "MMMM dd, yyyy h:mm tt"
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$tenantName - Tenant Discovery Report</title>
    
    <!-- Load Chart.js FIRST -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    
    <!-- Load CSS -->
    $(Get-HtmlStyle)
</head>
<body>
    <div class="container">
        <!-- Header -->
        <div class="report-header">
            <h1>$tenantName</h1>
            <p style="font-size:1.1em;">Tenant Discovery Report</p>
            <div class="report-meta">
                <div class="report-meta-item">📅 Generated: $reportDate</div>
                <div class="report-meta-item">💾 Data Source: In-memory hashtable</div>
                <div class="report-meta-item">✅ Data Freshness: Current session</div>
            </div>
        </div>
        
        <!-- Navigation -->
        $(Build-Navigation -Sections $sectionContents)
        
        <!-- Findings Panel - Uses ALL collected findings -->
        <div id="highlights">
            $(Build-FindingsPanel -AllFindings $allFindings)
        </div>
        
        <!-- Sections -->
"@
    
    foreach ($section in $sectionContents) {
        $sectionWorkload = Get-AssessmentSectionWorkload -SectionId $section.Id
        $htmlContent += @"
        <div class="section" id="$($section.Id)">
            <div class="section-workload">$sectionWorkload</div>
            <h2>$($section.Name)</h2>
            $($section.Content)
        </div>
"@
    }
    
    $htmlContent += @"
    </div>
    
    <!-- Chart.js function definitions -->
    $(Get-HtmlScript)
</body>
</html>
"@
    
    #endregion
    
    #region Write File
    
    try {
        $htmlContent | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
        #Write-Host "✅ HTML report saved: $OutputPath" -ForegroundColor Green
        
        # Display summary
        Write-Host "`n📊 Report Summary:" -ForegroundColor Cyan
        Write-Host "   Sections: $($sectionContents.Count)" -ForegroundColor White
        Write-Host "   Findings: $($allFindings.Count)" -ForegroundColor White
        
        $criticalCount = ($allFindings | Where-Object { $_.Type -eq 'Risk' }).Count
        $warningCount = ($allFindings | Where-Object { $_.Type -eq 'Warning' }).Count
        $infoCount = ($allFindings | Where-Object { $_.Type -eq 'Info' }).Count
        
        if ($criticalCount -gt 0) {
            Write-Host "   🔴 Critical: $criticalCount" -ForegroundColor Red
        }
        if ($warningCount -gt 0) {
            Write-Host "   ⚠️  Warnings: $warningCount" -ForegroundColor Yellow
        }
        if ($infoCount -gt 0) {
            Write-Host "   ℹ️  Info: $infoCount" -ForegroundColor Cyan
        }
        
        # Open in browser
        try {
            Start-Process $OutputPath
        } catch {
            Write-Verbose "Could not auto-open report in browser: $_"
        }
        
        return [PSCustomObject]@{
            Success = $true
            OutputPath = $OutputPath
            SectionCount = $sectionContents.Count
            FindingsCount = $allFindings.Count
            CriticalCount = $criticalCount
            WarningCount = $warningCount
            InfoCount = $infoCount
        }
    } catch {
        Write-Error "Failed to write HTML report: $_"
        return [PSCustomObject]@{
            Success = $false
            OutputPath = $null
            Error = $_.Exception.Message
        }
    }
    #endregion
}

#endregion
#endregion HTML Report Helpers
}

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
    # Connect to Microsoft Office 365 Services
    $connectOffice365Params = @{}
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

    # Get default tenant display name from live connection
    $defaultOrganization = $null
    try {
        $defaultOrganization = Get-AssessmentTenantOrganization
    } catch {}
    $defaultTenantDisplayName = if ($defaultOrganization -and $defaultOrganization.DisplayName) { $defaultOrganization.DisplayName } else { 'Tenant' }
}

#Get Export Path
$profileFileTagSource = if ([string]::IsNullOrWhiteSpace($effectiveOutputProfileLabel)) { $OutputProfile } else { $effectiveOutputProfileLabel }
$profileFileTag = if ([string]::IsNullOrWhiteSpace($profileFileTagSource)) { 'Profile' } else { ($profileFileTagSource -replace '[^A-Za-z0-9_-]', '') }
$defaultReportFileName = ("{0} Tenant Discovery Report-{1}" -f $defaultTenantDisplayName, $profileFileTag)
if ([string]::IsNullOrWhiteSpace($ExportPath)) {
    $ExportDetails = Get-ExportPath -FileName $defaultReportFileName
} else {
    $ExportDetails = Get-ExportPath -FileName $defaultReportFileName -UserInputPath $ExportPath
}

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
            $script:ProfileCollectionPlan.CollectEmailActivityDetails = $false
            $script:ProfileCollectionPlan.CollectExchangeGroups = $false
            $script:ProfileCollectionPlan.CollectMailFlowRulesConnectors = $false
            $script:ProfileCollectionPlan.CollectPublicFolders = $false
            $script:ProfileCollectionPlan.CollectThirdPartySpamFiltering = $false
            $script:ProfileCollectionPlan.CollectSmtpRelayConfiguration = $false
            $script:ProfileCollectionPlan.CollectTeamsVoiceDetails = $false
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
            $script:ProfileCollectionPlan.CollectTeamsVoiceDetails
        ) { 'Collected' } else { 'Skipped' }
        Collectors = [ordered]@{
            UnifiedGroups  = [bool]$script:ProfileCollectionPlan.CollectUnifiedGroups
            TeamsVoice     = [bool]$script:ProfileCollectionPlan.CollectTeamsVoiceDetails
            SharePointSite = $true
            Teams          = $true
        }
        NotCollectedReason = if (
            $script:ProfileCollectionPlan.CollectUnifiedGroups -or
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
Write-Log -Type INFO -Message ("Profile collection plan ({0}): ExchangeRecipients={1}; EmailActivity={2}; ExchangeGroups={3}; MailFlow={4}; PublicFolders={5}; SpamFiltering={6}; SMTPRelay={7}; TeamsVoice={8}; UnifiedGroups={9}; OwnershipTables={10}; AssessmentTables={11}; ConfigSummaryTables={12}; LicenseMetadata={13}" -f $effectiveOutputProfileLabel, $script:ProfileCollectionPlan.CollectExchangeRecipients, $script:ProfileCollectionPlan.CollectEmailActivityDetails, $script:ProfileCollectionPlan.CollectExchangeGroups, $script:ProfileCollectionPlan.CollectMailFlowRulesConnectors, $script:ProfileCollectionPlan.CollectPublicFolders, $script:ProfileCollectionPlan.CollectThirdPartySpamFiltering, $script:ProfileCollectionPlan.CollectSmtpRelayConfiguration, $script:ProfileCollectionPlan.CollectTeamsVoiceDetails, $script:ProfileCollectionPlan.CollectUnifiedGroups, $script:ProfileCollectionPlan.BuildOwnershipGovernanceTables, $script:ProfileCollectionPlan.BuildAssessmentReportTables, $script:ProfileCollectionPlan.BuildConfigurationSummaryTables, $script:ProfileCollectionPlan.BuildLicenseClassificationMetadata) -ExportFileLocation $ExportDetails

#Global Start Time for Script
$global:InitialStart = Get-Date

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
    $baseCollectionSteps = 12 # Exchange(6) + Hybrid(4) + Collaboration(2)
    $identitySteps = if ($GraphTest -eq 'REST') { 2 } else { 13 }
    $combineSteps = if ($reportingMode -eq "combined" -or $reportingMode -eq "all") { 1 } else { 0 }
    $postProcessingSteps = 4
    $overallCollectionSteps = $baseCollectionSteps + $identitySteps + $combineSteps + $postProcessingSteps
    Initialize-AssessmentProgress -TotalSteps $overallCollectionSteps

    Write-ConsoleSection -Step '1/5' -Title 'Exchange inventory'
    Invoke-ProfileAwareAssessmentStep -Name 'Exchange recipients' -Enabled $script:ProfileCollectionPlan.CollectExchangeRecipients -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-AllRecipientDetails -detailLevel $reportingMode }
    Invoke-AssessmentProgressStep -Name 'Exchange mailboxes' -ScriptBlock { Get-AllExchangeMailboxDetails -detailLevel $reportingMode }
    Invoke-ProfileAwareAssessmentStep -Name 'Email activity insights' -Enabled $script:ProfileCollectionPlan.CollectEmailActivityDetails -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-EmailActivityInsights -detailLevel $reportingMode }
    Invoke-ProfileAwareAssessmentStep -Name 'Exchange groups' -Enabled $script:ProfileCollectionPlan.CollectExchangeGroups -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-ExchangeGroupDetails -detailLevel $reportingMode }
    Invoke-ProfileAwareAssessmentStep -Name 'Mail flow rules/connectors' -Enabled $script:ProfileCollectionPlan.CollectMailFlowRulesConnectors -SkipReason 'Skipped in best-practices-only profile to reduce runtime.' -ScriptBlock { Get-MailFlowRulesandConnectors -detailLevel $reportingMode }
    Invoke-ProfileAwareAssessmentStep -Name 'Public folders' -Enabled $script:ProfileCollectionPlan.CollectPublicFolders -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-AllPublicFolderDetails -detailLevel $reportingMode }

    Write-ConsoleSection -Step '2/5' -Title 'Hybrid and configuration'
    Invoke-AssessmentProgressStep -Name 'Exchange hybrid configuration' -ScriptBlock { Get-ExchangeHybridConfiguration -detailLevel $reportingMode }
    Invoke-AssessmentProgressStep -Name 'Federation/cross-tenant configuration' -ScriptBlock { Get-FederationAndCrossTenantConfiguration }
    Invoke-ProfileAwareAssessmentStep -Name 'Third-party spam filtering configuration' -Enabled $script:ProfileCollectionPlan.CollectThirdPartySpamFiltering -SkipReason 'Requires mail flow connector/rule collection, which is disabled for this profile.' -ScriptBlock { Get-ThirdPartySpamFilteringConfig }
    Invoke-ProfileAwareAssessmentStep -Name 'SMTP relay configuration' -Enabled $script:ProfileCollectionPlan.CollectSmtpRelayConfiguration -SkipReason 'Requires mail flow connector collection, which is disabled for this profile.' -ScriptBlock { Get-SMTPRelayConfiguration }

    Write-ConsoleSection -Step '3/5' -Title 'Identity, devices, and licensing'
    # Determine if using REST or SDK Graph API
    switch ($GraphTest) {
        "REST" {
            Write-Verbose "Attempting to use Microsoft Graph REST API for Tenant Object and License details"
            Invoke-AssessmentProgressStep -Name 'Graph user statistics' -ScriptBlock { Get-GraphUserStats }
            Invoke-AssessmentProgressStep -Name 'Entra groups (REST)' -ScriptBlock { Get-EntraIDGroups -detailLevel $reportingMode -GraphAuthType REST }
         }
        "SDK" {
            Write-Verbose "Attempting to use Microsoft Graph SDK for Tenant Object and License details"
            Invoke-AssessmentProgressStep -Name 'Conditional Access policies' -ScriptBlock { Get-ConditionalAccessPoliciesReport -detailLevel $reportingMode }
            Invoke-AssessmentProgressStep -Name 'License SKUs' -ScriptBlock { Get-AllLicenseSKUs }
            Invoke-AssessmentProgressStep -Name 'Users' -ScriptBlock { Get-AllUserDetails -detailLevel $reportingMode }
            Invoke-ProfileAwareAssessmentStep -Name 'Teams voice details' -Enabled $script:ProfileCollectionPlan.CollectTeamsVoiceDetails -SkipReason 'Not required for this profile output.' -ScriptBlock { Get-TeamsVoiceDetails }
            Invoke-AssessmentProgressStep -Name 'Entra groups (SDK)' -ScriptBlock { Get-EntraIDGroups -detailLevel $reportingMode -GraphAuthType SDK }
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
    $sharePointDiscoveryService = if ($connectionResult -and $connectionResult.SharePointOnline) {
        'SPO'
    }
    elseif ($script:CollectionDepthPolicy -and $script:CollectionDepthPolicy.IsMinimum) {
        'API'
    }
    else {
        'MGGraph'
    }
    Invoke-AssessmentProgressStep -Name "SharePoint/OneDrive sites ($sharePointDiscoveryService)" -ScriptBlock { Get-SharePointAndOneDriveSites -detailLevel $reportingMode -ServiceName $sharePointDiscoveryService }
    Write-Host


    #Combine Reporting - Optional
    if ($reportingMode -eq "combined" -or $reportingMode -eq "all") {
        Write-Host
        Write-Host "Consolidating Discovery Report data for each user / object into one file" -ForegroundColor Black -BackgroundColor Green
        #Combine Reports
        Invoke-AssessmentProgressStep -Name 'Combined user/mailbox reporting' -ScriptBlock { Report-UserAndMailboxStats }
    }

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

Write-Host ""
if ($global:AllDiscoveryErrors.Count -gt 0) {
    try {
        $errorReportSummary = Export-ErrorReports -ExportFileLocation $ExportDetails -ErrorData $global:AllDiscoveryErrors -logReportDirectory $ExportDetails
        if ($errorReportSummary -and -not [string]::IsNullOrWhiteSpace([string]$errorReportSummary.FolderPath)) {
            $generatedArtifacts['Error Reports'] = $errorReportSummary.FolderPath
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

########################################################
### End of HTML Report Integration ###
########################################################

Write-Log -Type INFO -Message "COMPLETED: Gathered Tenant Details. Completed Time: $($timeString)" -ExportFileLocation $ExportDetails

# Encourage GC after large export/report generation to reduce retained working set in long-lived shells.
$ExportTenantStatsHash = $null
$script:AssessmentStepMetrics = $null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
