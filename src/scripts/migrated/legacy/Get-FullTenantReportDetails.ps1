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
5. Follow the prompts to provide necessary input such as ReportingMode, ExportPath, Authentication, and others as prompted.
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
    [ValidateSet('Minimum', 'Combined', 'All', 'Geek')]
    [string]$ReportingMode,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Lean', 'Standard', 'Full')]
    [string]$OutputProfile = 'Standard',
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
    [string]$ClientId
)

$effectiveSkipHtmlReport = $SkipHtmlReport.IsPresent
$effectiveSkipPdfReport = $SkipPdfReport.IsPresent
$effectiveSkipJsonReport = $SkipJsonReport.IsPresent

switch ($OutputProfile) {
    'Lean' {
        if (-not $PSBoundParameters.ContainsKey('SkipHtmlReport')) { $effectiveSkipHtmlReport = $true }
        if (-not $PSBoundParameters.ContainsKey('SkipPdfReport')) { $effectiveSkipPdfReport = $true }
        if (-not $PSBoundParameters.ContainsKey('SkipJsonReport')) { $effectiveSkipJsonReport = $true }
    }
    'Standard' {
        if (-not $PSBoundParameters.ContainsKey('SkipPdfReport')) { $effectiveSkipPdfReport = $true }
        if (-not $PSBoundParameters.ContainsKey('SkipJsonReport')) { $effectiveSkipJsonReport = $true }
    }
    'Full' {
    }
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

$tenantHtmlReportPath = Join-Path -Path $PSScriptRoot -ChildPath 'New-TenantHtmlReport.ps1'
if (Test-Path $tenantHtmlReportPath) {
    . $tenantHtmlReportPath
} else {
    Write-Warning "Optional HTML helper script not found: $tenantHtmlReportPath. Built-in HTML generation remains available, but PDF export helpers will be unavailable."
}

$tenantQuestionnairePath = Join-Path -Path $PSScriptRoot -ChildPath 'Export-TenantToTenantQuestionnaireMarkdown.ps1'
if (Test-Path $tenantQuestionnairePath) {
    . $tenantQuestionnairePath
} else {
    Write-Warning "Optional questionnaire helper script not found: $tenantQuestionnairePath. Questionnaire export will be skipped."
}


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
#Level of Detail Reporting
function Set-ReportMode {

    Write-Host "Reporting Mode Explanations:" -ForegroundColor Yellow
    Write-Host "Minimum   - Provides basic details for a quick overview."
    Write-Host "Combined  - Combines details from similar reports. E.g., combining a user's mailbox and SharePoint data."
    Write-Host "All       - Provides a comprehensive, detailed report that includes all possible details and combined reports."
    Write-Host "Geek      - Provides every available detail from reports, does not include Combined reports"
    Write-Host ""
    
    $selectedMode = $null
    do {
        $selectedMode = Read-Host "Please select a reporting mode (Minimum, Combined, All, Geek) or type 'help' for explanations"
        
        # Check if the user wants to see the explanations again
        if ($selectedMode -eq 'help') {
            ShowModeExplanations
            $selectedMode = $null  # Reset to null to continue the loop
        }
    } while ($selectedMode -notin @('Minimum', 'Combined', 'All', 'Geek'))

    $selectedMode = $selectedMode.ToLower()
    Write-Host "You selected: $selectedMode reporting mode" -ForegroundColor Green
    return $selectedMode
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

    # If user input is empty, default to Desktop
    if ([string]::IsNullOrEmpty($userInput)) {
        $userInput = [Environment]::GetFolderPath("Desktop")
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
            # If no folder path (i.e., just a file name), use Desktop
            if ([string]::IsNullOrWhiteSpace($folderPath)) {
                $folderPath = [Environment]::GetFolderPath("Desktop")
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

    # Using List to improve performance on adding items
    $totalCount = $HashToConvert.Keys.Count
    $customObject = [PSCustomObject]@{}

    foreach ($nestedKey in $HashToConvert.Keys) {
        Write-ProgressHelper -Total $totalCount -Activity "Converting Hash Table" -Operation "Converting $($nestedKey)"
        
        # Define the attributes
        $attributes = $HashToConvert[$nestedKey]

        # Initialize a new custom object for each item
        #$customObject = [PSCustomObject]@{}

        # If the attributes are a hashtable, convert them to a custom object
        if ($attributes -is [hashtable] -or $attributes -is [System.Collections.Specialized.OrderedDictionary]) {
            Write-Verbose "Converting $($nestedKey) Hash Table to Array"
           # $customObject = New-Object -TypeName PSObject

            # Add the tenant name to the attribute name
            if ($tenant) {
                foreach ($attribute in $attributes.keys) {
                    Write-Verbose "Adding $($attribute)_$($tenant) to the custom object"
                    $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)_$($tenant)" -Value ($attributes[$attribute] -join ';')
                }
            } else {
                foreach ($attribute in $attributes.keys) {
                    Write-Verbose "Adding $($attribute) to the custom object"
                    $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)" -Value ($attributes[$attribute] -join ';')
                }
            }

            #$ExportArray.Add($customObject)
        } 
        # If the attributes are an array or a custom object, add them directly
        elseif ($attributes -is [array] -or $attributes -is [PSCustomObject]) {
            Write-Verbose "Adding $($nestedKey) Array to the export array"
            $ExportArray.Add($attributes)
        }
        # If the attributes are a string, add it to the export array
        else {
            Write-Verbose "Adding '$($nestedKey)' String to the export array with attribute '$($attributes)'"
            $customObject | Add-Member -MemberType NoteProperty -Name "$($nestedKey)" -Value ($attributes)
        }
    }

    Write-ProgressHelper -Total $totalCount -Activity "Converting Hash Table" -Completed
    return $customObject  # Convert the List to an array if needed outside this function
}

function Filter-TenantStatsHash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$tenantStatsHash,
        [Parameter(Mandatory = $true)]
        [string]$reportingMode,
        [Parameter(Mandatory = $false)]
        $GraphTest
    )

    # Common tables to remove for both 'combined' and 'minimum' detail levels
    $commonTablesToRemove = @(
        'AllMailboxes-MailIdentity'
    )

    # Clone the original hash to avoid modifying the input directly
    $filteredStatsHash = $tenantStatsHash.Clone()

    # Determine additional tables to remove based on the reporting mode
    $additionalTables = switch ($reportingMode) {
        "combined" {
            @(
                'Users', 'ArchiveMailboxes', 'ArchiveMailboxStats', 'NonUserMailboxes',
                'PrimaryMailboxStats', 'AllMailboxes', 'InActiveMailboxes', 
                'LitigationHoldMailboxes', 'UnifiedGroups'
            )
        }
        "minimum" {
            @(
                'InActiveMailboxes','ArchiveMailboxStats', 'NonUserMailboxes',
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

    # Remove the specified tables from the cloned hash
    foreach ($key in $tablesToRemove) {
        if ($filteredStatsHash[$key]) {
            $filteredStatsHash.Remove($key)
        }
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
        'SpamFilteringSummary',
        'SMTPRelaySummary',
        'FederationSummary',
        'TeamsVoiceSummary',
        'MfaRegistrationSummary'
    )
    
    # === Sheet ordering ===
    $desiredOrder = @(
        # Assessment Outputs
        "BestPractices", "BestPracticeFindings", "MigrationReadiness", "SecureScoreActions",

        # Licensing & Tenant Info
        "TenantInfoSummary", "LicenseSKUs", "Domains", "AuthenticationConfigSummary", "AuthenticationMethods", "AuthenticationSSOApplications", "AuthenticationConfig", "Admins",

        # Users
        "Users", "UserFullDetails", "DeviceDetails",

        # Mailboxes
        "AllMailboxes", "PrimaryMailboxStats", "MailboxFullDetails", "ArchiveMailboxes", "ArchiveMailboxStats", "LitigationHoldMailboxes", "InactiveMailboxes", "InactiveMailboxDetails", "NonUserMailboxes", "AllRecipients",

        # Groups
        "AllExchangeGroups", "UnifiedGroups", "EntraIDGroups",

        # Public Folders
        "PublicFolderDetails", "PublicFolderPerms",

        # Mail Flow
        "MailFlowRules", "MailFlowConnectors", "RemoteDomains", "SMTPRelayConfig",

        # Security & Compliance
        "SecuritySecureScore", "ConditionalAccessPolicies", "SpamFilteringSummary", "SMTPRelaySummary", "FederationSummary", "TeamsVoiceSummary", "SpamFilteringConfig",

        # Cloud Services
        "OneDrive",
        "SharePoint",

        # Hybrid / Infra
        "HybridConfiguration", "TenantInfo"
    )

    $orderedTables = @()
    foreach ($name in $desiredOrder) {
        if ($hashtable.ContainsKey($name)) {
            $orderedTables += $name
        }
    }
    $orderedTables += ($hashtable.Keys | Where-Object { $orderedTables -notcontains $_ } | Sort-Object)

    $totalCount = ($orderedTables | Measure-Object).Count
    foreach ($table in $orderedTables) {
        try {
            Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash To Excel"
            Write-Log -Type DEBUG -Message ("Exporting '{0}' Hash Table to '{1}'" -f $table, $ExportDetails) -ExportFileLocation $ExportDetails

            $tableValue = $hashtable[$table]
            $exportData = @()
            if ($tableValue -is [hashtable] -or $tableValue -is [System.Collections.Specialized.OrderedDictionary]) {
                $exportData = @($tableValue.Values)
            } elseif ($tableValue -is [System.Collections.IEnumerable] -and -not ($tableValue -is [string])) {
                $exportData = @($tableValue)
            } else {
                $exportData = @($tableValue)
            }

            if ($exportData.Count -gt 0) {
                $exportData = @($exportData | ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ })
                $attempt = 0
                $maxAttempts = 3
                $saved = $false
                while (-not $saved -and $attempt -lt $maxAttempts) {
                    $attempt++
                    if (Test-FileLocked -Path $ExportDetails) {
                        Write-Log -Type WARNING -Message "Excel file is locked (attempt $attempt/$maxAttempts). Close the file and retrying in 5 seconds..." -ExportFileLocation $ExportDetails
                        Start-Sleep -Seconds 5
                        continue
                    }
                    try {
                        $exportData | Export-Excel -Path $ExportDetails -WorksheetName $table -ClearSheet -AutoSize -BoldTopRow
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
                $logType = if ($optionalEmptySheets -contains $table) { 'INFO' } else { 'WARNING' }
                Write-Log -Type $logType -Message "No data found for $($table) to export to Excel" -ExportFileLocation $ExportDetails
            }
        }
        catch {
            Write-Log -Type Error -Message "An error occurred in Exporting Hash To Excel for $($table) to '$($ExportDetails)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }
    }
    Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash Table to Excel" -Completed
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

    # Validate ErrorData
    if ($ErrorData.Count -eq 0) {
        Write-Log -Type WARNING -Message "No error data provided to export. Exiting function." -ExportFileLocation $ExportDetails
        return
    }

    Write-Log -Type INFO -Message "START: Export all Errors" -ExportFileLocation $ExportDetails
    # Handle quotes in input
    $ExportFileLocation = $ExportFileLocation -replace '"', ''

    if ($ExportFileLocation) {
        $directory = [System.IO.Path]::GetDirectoryName($ExportFileLocation)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ExportFileLocation)
        $errorReportFolderName = $baseName + " Error Reporting"
        $errorReportFolderDirectory = Join-Path -Path $directory -ChildPath $errorReportFolderName
    } else {
        if ($ErrorReportFolderDirectory) {
            $directory = $ErrorReportFolderDirectory
        } else {
            $directory = $env:TEMP
        }
        
        if ($ErrorReportFolderName) {
            $errorReportFolderName = $ErrorReportFolderName
        } else {
            $errorReportFolderName = "$BaseName Error Reporting"
        }
        $errorReportFolderDirectory = Join-Path -Path $directory -ChildPath $errorReportFolderName
    }

    Write-Log -Type INFO -Message "INFO: Exporting Error Logs to directory $($errorReportFolderDirectory)" -ExportFileLocation $ExportDetails

    try {
        if (-not (Test-Path $errorReportFolderDirectory)) {
            $result = New-Item -Path $errorReportFolderDirectory -ItemType Directory
            Write-Log -Type INFO -Message "INFO: Error Report Directory '$($errorReportFolderDirectory)' does not exist. Created Folder Directory" -ExportFileLocation $ExportDetails
        }

        $newBaseName = "$BaseName-ErrorLog"
        
        $paths = @{
            'json' = Join-Path -Path $errorReportFolderDirectory -ChildPath "$newBaseName.json"
            'txt'  = Join-Path -Path $errorReportFolderDirectory -ChildPath "$newBaseName.log"
            'csv'  = Join-Path -Path $errorReportFolderDirectory -ChildPath "$newBaseName.csv"
        }

        $ErrorData | ConvertTo-Json -Depth 1 | Set-Content -Path $paths['json']
        $ErrorData | Out-File $paths['txt']
        $ErrorData | Export-Csv -Path $paths['csv'] -NoTypeInformation -Encoding UTF8

        $paths.GetEnumerator() | ForEach-Object {
            Write-Log -Type INFO -Message "INFO: Exported $($_.Key) Error Logs to directory $($_.Value)" -ExportFileLocation $ExportDetails
        }
            #Display Error Details
            Write-Host "Error Reporting Details" -ForegroundColor Black -BackgroundColor Yellow
            Write-Host "Check '$($errorReportFolderDirectory)' for error logs " -ForegroundColor Cyan
            Write-Host "$($global:AllDiscoveryErrors.count) " -ForegroundColor Red -NoNewline
            Write-Host "Error(s) encountered. "

    } catch {
        Write-Error "Failed to export error reports: $_"
    }
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
        if ($MailObjectHash['AllRecipients'].ContainsKey($recipientName)) {
            $matchingRecipient = $MailObjectHash['AllRecipients'][$recipientName]
        } elseif ($MailObjectHash['AllMailboxes'].ContainsKey($recipientName)) {
            $matchingRecipient = $MailObjectHash['AllMailboxes'][$recipientName]
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

#Gather all Mailboxes, Group Mailboxes, Unified Groups, and Public Folders
function Get-AllExchangeMailboxDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    Write-Host "Getting all mailboxes and inactive mailboxes with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] START: Getting all mailboxes with $($detailLevel) details" -ExportFileLocation $ExportDetails
    try {
        # Gather Mailboxes - Include InActive Mailboxes
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
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

                $exoMailboxes = Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -Properties $Properties -IncludeInactiveMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Select-Object $DesiredProperties
            }
            geek {
                $exoMailboxes = Get-EXOMailbox -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox'" -IncludeInactiveMailbox -PropertySets All -ResultSize Unlimited -ErrorAction SilentlyContinue 
            }
        }
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering all mailboxes (Get-EXOMailbox) including Inactive Mailboxes" -ExportFileLocation $ExportDetails
        
        #Create Hash Table to store mailboxes
        # Ensure global hash table structure
        if (-not $global:tenantStatsHash) {
            $global:tenantStatsHash = @{}
        }
        $global:tenantStatsHash["AllMailboxes"] = @{}
        $Global:tenantStatsHash["AllMailboxes-MailIdentity"] = @{}
        $global:tenantStatsHash["NonUserMailboxes"] = @{}
        
        $global:tenantStatsHash["ArchiveMailboxes"] = @{}
        $global:tenantStatsHash["InactiveMailboxes"] = @{}
        $global:tenantStatsHash["LitigationHoldMailboxes"] = @{}

        # Insert individual mailboxes into the hashtable
        $totalCount = $exoMailboxes.Count
        foreach ($mailbox in $exoMailboxes) {
            $key = $mailbox.UserPrincipalName
            Write-ProgressHelper -Total $totalCount -Activity "Gathering All Exchange Mailbox Details" -Operation "Gathering Mailbox Details for $($key)"
            
            # Set the key based on the mailbox type
            #$MailboxTypeKey = $mailbox.RecipientTypeDetails.tostring() # Use RecipientTypeDetails as the key
            $Global:tenantStatsHash["AllMailboxes"][$key] = $mailbox
            $Global:tenantStatsHash["AllMailboxes-MailIdentity"][$mailbox.Identity] = $mailbox

            if ($mailbox.RecipientTypeDetails -ne "UserMailbox" -and $mailbox.RecipientTypeDetails -ne "GroupMailbox") {
                #$MailboxTypeKey = "NonUserMailboxes"
                $global:tenantStatsHash["NonUserMailboxes"][$key] = $mailbox
            }
            if ($mailbox.ArchiveStatus -eq "Active") {
                #$MailboxTypeKey = "ArchiveMailboxes"
                $global:tenantStatsHash["ArchiveMailboxes"][$key] = $mailbox
            }
            if ($mailbox.IsInactiveMailbox -eq $true) {
                #$MailboxTypeKey = "InactiveMailboxes"
                $global:tenantStatsHash["InactiveMailboxes"][$key] = $mailbox
            }
            if ($mailbox.LitigationHoldEnabled -eq $true) {
                #$MailboxTypeKey = "LitigationHoldMailboxes"
                $global:tenantStatsHash["LitigationHoldMailboxes"][$key] = $mailbox
            }
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Mailbox Details. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    } finally {
        Write-ProgressHelper -Total $totalCount -Activity "Gathering All Exchange Mailbox Details" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] COMPLETED: Gathering All Mailbox Details in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }

    #Mailbox Statistics to Hash Table
    ###########################################################################################################################################
    ## Primary Mailbox Stats
    try {
        $start = Get-Date
        Write-Progress -Activity "Gathering All Primary Mailbox Statistics" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        Write-Host "  Getting primary mailbox stats..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering All Primary Mailbox Statistics including Group Mailboxes and Inactive Mailboxes" -ExportFileLocation $ExportDetails

        $global:tenantStatsHash["PrimaryMailboxStats"] = @{}
        $activeMailboxes = $global:tenantStatsHash['AllMailboxes'].Values | Where-Object { $_.IsInactiveMailbox -ne $true }
        $graphStatsCount = 0

        # Try Graph mailbox usage report (fast) for active mailboxes
        try {
            if (Get-MgContext -ErrorAction SilentlyContinue) {
                $mailboxUsageUri = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D180')"
                $tempCsvFile = Join-Path $env:TEMP "MailboxUsageReport-$(Get-Date -Format 'yyyyMMddHHmmss').csv"
                Invoke-MgGraphRequest -Uri $mailboxUsageUri -Method GET -OutputFilePath $tempCsvFile -ErrorAction Stop | Out-Null
                if (Test-Path $tempCsvFile) {
                    $graphReportData = Import-Csv -Path $tempCsvFile -ErrorAction Stop
                    Remove-Item -Path $tempCsvFile -Force -ErrorAction SilentlyContinue
                    if ($graphReportData -and $graphReportData.Count -gt 0) {
                        $graphMailboxHash = @{}
                        foreach ($item in $graphReportData) {
                            $upn = $item.'User Principal Name'
                            if (-not [string]::IsNullOrWhiteSpace($upn)) {
                                $graphMailboxHash[$upn] = $item
                            }
                        }
                        foreach ($mailbox in $activeMailboxes) {
                            if (-not $graphMailboxHash.ContainsKey($mailbox.UserPrincipalName)) { continue }
                            $graphData = $graphMailboxHash[$mailbox.UserPrincipalName]
                            $storageBytes = 0
                            $deletedBytes = 0
                            $itemCount = "0"
                            if ($graphData.'Storage Used (Byte)') { $storageBytes = [double]$graphData.'Storage Used (Byte)' }
                            if ($graphData.'Deleted Item Size (Byte)') { $deletedBytes = [double]$graphData.'Deleted Item Size (Byte)' }
                            if ($graphData.'Item Count') { $itemCount = $graphData.'Item Count' }
                            $guidKey = if ($mailbox.ExchangeGuid) { $mailbox.ExchangeGuid.ToString() } elseif ($mailbox.Guid) { $mailbox.Guid.ToString() } else { $null }
                            if (-not $guidKey) { continue }
                            $stats = [PSCustomObject]@{
                                DisplayName = $mailbox.DisplayName
                                TotalItemSize = "$([math]::Round($storageBytes / 1GB, 4)) GB ($storageBytes bytes)"
                                ItemCount = $itemCount
                                TotalDeletedItemSize = "$([math]::Round($deletedBytes / 1GB, 4)) GB ($deletedBytes bytes)"
                                MailboxType = $mailbox.RecipientTypeDetails
                                MailboxGuid = $mailbox.ExchangeGuid
                            }
                            $global:tenantStatsHash["PrimaryMailboxStats"][$guidKey] = $stats
                        }
                    }
                }
            }
            $graphStatsCount = $global:tenantStatsHash["PrimaryMailboxStats"].Count
        } catch {
            Write-Log -Type WARNING -Message "[Get-AllExchangeMailboxDetails] Graph mailbox usage report failed: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        Write-Progress -Completed

        # Fallback to EXO stats for missing mailboxes only (includes inactive)
        $mailboxesNeedingStats = $global:tenantStatsHash['AllMailboxes'].Values | Where-Object {
            if ($_.ExchangeGuid) {
                -not $global:tenantStatsHash["PrimaryMailboxStats"].ContainsKey($_.ExchangeGuid.ToString())
            } else {
                $true
            }
        }
        if ($mailboxesNeedingStats.Count -gt 0) {
            #$inactiveMBXTest = ($global:tenantStatsHash['AllMailboxes'].Values | Where-Object { $_.IsInactiveMailbox -eq $true }).Count -gt 0
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Graph report covered $($global:tenantStatsHash['PrimaryMailboxStats'].Count) mailboxes; fetching EXO stats for $($mailboxesNeedingStats.Count) missing/inactive." -ExportFileLocation $ExportDetails

            <#
            $validMailboxStatsInput = $mailboxesNeedingStats | Where-Object {
                $_.UserPrincipalName -and $_.UserPrincipalName -match ".+@.+"
            }
            $invalidMailboxStatsInput = $mailboxesNeedingStats | Where-Object {
                -not $_.UserPrincipalName -or $_.UserPrincipalName -notmatch ".+@.+"
            }
            #>

            $allRemainingMBXStats = $mailboxesNeedingStats | Get-EXOMailboxStatistics -Identity $_.DistinguishedName -IncludeSoftDeletedRecipient
            $allRemainingMBXStats | ForEach-Object {
                $key = $_.MailboxGuid.ToString()
                if (-not $global:tenantStatsHash["PrimaryMailboxStats"].ContainsKey($key)) {
                    $global:tenantStatsHash["PrimaryMailboxStats"][$key] = $_
                }
            }
        }

        $finalStatsCount = $global:tenantStatsHash["PrimaryMailboxStats"].Count
        $exoStatsFilledCount = $finalStatsCount - $graphStatsCount
        Write-Verbose "exoStatsFilledCount: $exoStatsFilledCount"
        Write-Verbose "graphStatsCount: $graphStatsCount"
        Write-Verbose "finalStatsCount: $finalStatsCount"

        if ($exoFilledCount -lt 0) { $exoFilledCount = 0 }
        $totalMailboxes = $global:tenantStatsHash['AllMailboxes'].Values.Count
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Mailbox stats summary: Graph=$graphStatsCount; EXO filled=$exoFilledCount; Total mailboxes=$totalMailboxes." -ExportFileLocation $ExportDetails
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllExchangeMailboxDetails] An error occurred in Gathering Mailbox Satistics and adding to Hash Table. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        $ProgressPreference = "SilentlyContinue"
        Write-Progress -Activity "Gathering All Primary Mailbox Statistics" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] COMPLETED: Gathering All Primary Mailbox Statistics in $($CompletedTime)" -ExportFileLocation $ExportDetails
        $ProgressPreference = "Continue"
    }
    
    ## Archive Mailbox Stats to Hash Table
    if ($global:tenantStatsHash['AllMailboxes'].Values | Where-Object {$_.ArchiveStatus -ne "None"}) {
        try {
            $start = Get-Date
            Write-Host "  Getting archive mailbox stats..." -ForegroundColor Cyan -nonewline
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails] Gathering All Archive Mailbox Statistics. Including Group and Inactive Mailboxes" -ExportFileLocation $ExportDetails
            Write-Progress -Activity "Gathering All Archive Mailbox Statistics" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
            $archiveMailboxStats = $global:tenantStatsHash['AllMailboxes'].Values | Where-Object {$_.ArchiveStatus -ne "None"} | Get-EXOMailboxStatistics -Archive -Properties MailboxGuid -IncludeSoftDeletedRecipients -ErrorAction SilentlyContinue
            if ($archiveMailboxStats) {
                $global:tenantStatsHash["ArchiveMailboxStats"] = @{}
                
                #Add to Tenant Stats Hash
                $archiveMailboxStats | ForEach-Object {
                    # Using MailboxGUID from the Mailbox Statistics as the key; matches against the ExchangeGUID from the Mailbox
                    $key = $_.MailboxGuid.ToString()
                    $value = $_
                    $global:tenantStatsHash["ArchiveMailboxStats"][$key] = $value
                }
            }
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
            $ProgressPreference = "SilentlyContinue"
            Write-Progress -Activity "Gathering All Archive Mailbox Statistics" -Completed
            $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
            Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
            Write-Log -Type INFO -Message "[Get-AllExchangeMailboxDetails]COMPLETED: Gathering All Archive Mailbox Statistics in $($CompletedTime)" -ExportFileLocation $ExportDetails
            $ProgressPreference = "Continue"
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
    try {
        $start = Get-Date
        # Ensure global hash table structure
        if (-not $global:tenantStatsHash) {
            $global:tenantStatsHash = @{}
        }
        $global:tenantStatsHash["AllRecipients"] = @{}
        Write-Host "Getting all Exchange Online Recipients $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] START: Gathering all Exchange Online Recipients $($detailLevel) details" -ExportFileLocation $ExportDetails
        Write-Progress -Activity "Gathering All Exchange Online Recipients" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))

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
                $allRecipients = Get-EXORecipient -Properties $Properties -ResultSize Unlimited -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox' -and RecipientTypeDetails -ne 'MailContact' -and RecipientTypeDetails -ne 'GuestMailUser' -and RecipientTypeDetails -ne 'MailUser'" -ErrorAction Stop | select $DesiredProperties
            }           
            geek {
                $allRecipients = Get-EXORecipient -PropertySets All -ResultSize Unlimited -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox' -and RecipientTypeDetails -ne 'MailContact' -and RecipientTypeDetails -ne 'GuestMailUser' -and RecipientTypeDetails -ne 'MailUser'" -ErrorAction Stop
            }
        }
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] FOUND $($allRecipients.count) Exchange Online Recipients $($detailLevel) details" -ExportFileLocation $ExportDetails
        #Add to hash table
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] Adding Exchange Online Recipients to Tenant Stats Hash" -ExportFileLocation $ExportDetails
        foreach ($recipient in $allRecipients) {
            $global:tenantStatsHash["AllRecipients"][$recipient.PrimarySmtpAddress] = $recipient
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllRecipientDetails] An error occurred in running Get-AllRecipientDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-Progress -Activity "Gathering All Exchange Online Recipients" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] COMPLETED: Gathering all Exchange Online Recipients" -ExportFileLocation $ExportDetails
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

    try {
        $start = Get-Date
        # Ensure global hash table structure
        if (-not $global:tenantStatsHash) {
            $global:tenantStatsHash = @{}
        }
        $global:tenantStatsHash["AllExchangeGroups"] = @{}
        
        Write-Host "Getting all Exchange Online Groups ..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] START: Gathering all Exchange Online Groups with $($detailLevel) details" -ExportFileLocation $ExportDetails

        # gather All Exchange Online Groups
        $allMailGroups = $global:tenantStatsHash['AllRecipients'].Values | Where-Object { $_.RecipientTypeDetails -like "*group" } | Sort-Object DisplayName
        
        Write-Log -Type INFO -Message "[Get-ExchangeGroupDetails] Gathering all Exchange Online Groups Details" -ExportFileLocation $ExportDetails
        $totalCount = $allMailGroups.count
        foreach ($object in $allMailGroups) {
            try {
                $identity = $object.identity.tostring()
                $PrimarySMTPAddress = $object.PrimarySMTPAddress.ToString()
                Write-ProgressHelper -Total $totalCount -Activity "Gathering All Exchange Online Group Details" -Operation "Gathering Group Details for $($PrimarySMTPAddress)"
                Write-Log -Type DEBUG -Message ("[Get-ExchangeGroupDetails] Gathering '{0}' '{1}' Group Details" -f $object.RecipientTypeDetails, $PrimarySMTPAddress) -ExportFileLocation $ExportDetails
    
                # Clear details
                $attributesToClear = @('groupDetails', 'groupOwners','groupMembers')
                foreach ($attribute in $attributesToClear) {
                    Set-Variable -Name $attribute -Value @()
                }
                
                # Conditional logic for different recipient types
                switch ($object.RecipientTypeDetails) {
                    "DynamicDistributionGroup" {
                        $groupDetails = Get-DynamicDistributionGroup $identity -ErrorAction SilentlyContinue
                        $groupMembers = Get-DynamicDistributionGroupMember $identity -ErrorAction SilentlyContinue -ResultSize unlimited -warningaction silentlycontinue
                    }
                    {$_ -in 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', "MailNonUniversalGroup"} {
                        $groupDetails = Get-DistributionGroup $identity -ErrorAction SilentlyContinue
                        $groupMembers = Get-DistributionGroupMember $identity -ResultSize unlimited -ErrorAction SilentlyContinue
                    }
                    "GroupMailbox" {
                        $groupDetails = Get-UnifiedGroup $identity -ErrorAction SilentlyContinue
                        $groupMembers = Get-UnifiedGroupLinks -Identity $identity -LinkType Member -ResultSize unlimited -ErrorAction SilentlyContinue
                    }
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
                    MembersCount                             = ($groupMembers | measure-object).count
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
                $global:tenantStatsHash["AllExchangeGroups"][$object.identity] = $currentobject
            }
            catch {
                Write-Log -Type ERROR -Message "[Get-ExchangeGroupDetails] An error occurred in running Get-ExchangeGroupDetails function. Exception: $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
        }
        Write-ProgressHelper -Total $totalCount -Activity "Gathering All Exchange Online Group Details" -Completed
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-ExchangeGroupDetails] An error occurred in running Get-ExchangeGroupDetails function. Exception: $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
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
    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    # Create Hash Tables for Mail Flow Rules and Connectors
    $global:tenantStatsHash["MailFlowRules"] = @{}
    $global:tenantStatsHash["MailFlowConnectors"] = @{}

    Write-Host "Getting all Mail Flow Rules and Connectors ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] START: Gathering all Mail Flow Rules and Connectors with $($detailLevel) details" -ExportFileLocation $ExportDetails
    Write-Progress -Activity "Getting all Mail Flow Rules details" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))

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
            $global:tenantStatsHash["MailFlowRules"][$rule.Priority] = $rule
        }
        Write-Progress -Activity "Getting all Mail Flow Rules details" -Completed
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
            $global:tenantStatsHash["MailFlowConnectors"][$connector.Id] = $currentConnector
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
        $global:tenantStatsHash["MailFlowConnectors"] = @{}

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
        Write-Progress -Activity "Getting all Mail Flow Rules details" -Completed
        Write-Progress -Activity "Adding Mail Connectors Details to Tenant Stats Hash" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-MailFlowRulesandConnectors] COMPLETED: Gathering All Mail Flow Rules and Connectors in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }
}

function Get-SMTPRelayConfiguration {
    param ()
    
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["SMTPRelayConfig"] = @{}
    
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
            $smtpAuthUsers = $global:tenantStatsHash["AllMailboxes"].Values | 
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
        $outboundConnectors = $global:tenantStatsHash["MailFlowConnectors"].Values | 
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
        
        $global:tenantStatsHash["SMTPRelayConfig"]["Configuration"] = $smtpConfig
        
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
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["SpamFilteringConfig"] = @{}
    
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
        $inboundConnectors = $global:tenantStatsHash["MailFlowConnectors"].Values | Where-Object {$_.ConnectorDirection -eq "Inbound"}
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
        $outboundConnectors = $global:tenantStatsHash["MailFlowConnectors"].Values | Where-Object {$_.ConnectorDirection -eq "Outbound"}
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
        $transportRules = $global:tenantStatsHash["MailFlowRules"].Values
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
        
        $global:tenantStatsHash["SpamFilteringConfig"]["Configuration"] = $spamFilterConfig
        
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
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["HybridConfiguration"] = @{}

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
        if ($global:tenantStatsHash -and $global:tenantStatsHash.ContainsKey("MailFlowConnectors")) {
            $mailFlowConnectors = $global:tenantStatsHash["MailFlowConnectors"].Values
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

        $global:tenantStatsHash["HybridConfiguration"]["ExchangeHybrid"] = $details
        Write-Log -Type INFO -Message "[Get-ExchangeHybridConfiguration] Hybrid=$($details.IsHybridConfigured) Type=$($details.HybridType) Evidence=$($details.EvidenceCount)" -ExportFileLocation $ExportDetails

    } catch {
        Write-Log -Type WARNING -Message "[Get-ExchangeHybridConfiguration] Error: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        $global:tenantStatsHash["HybridConfiguration"]["ExchangeHybrid"] = [pscustomobject]@{
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
    try {
        $start = Get-Date
        # Ensure global hash table structure
        if (-not $global:tenantStatsHash) {
            $global:tenantStatsHash = @{}
        }
        $global:tenantStatsHash["UnifiedGroups"] = @{}
        Write-Host "Getting all unified groups (including soft deleted)..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] START: Gathering all Unified with $($detailLevel) details" -ExportFileLocation $ExportDetails
        Write-Progress -Activity "Getting all Unified Group data" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        switch ($detailLevel) {
            {$_ -in "minimum", "combined", "all"} { 
                $DesiredProperties = @(
                    "PrimarySmtpAddress", "DisplayName", "AccessType", "RecipientTypeDetails",
                    "ExchangeGuid", @{Name="ManagedByDetails"; Expression={$_.ManagedByDetails -join ','}}, "Notes",
                    "SharePointSiteUrl", "ContentMailboxName", "GroupMemberCount",
                    "AllowAddGuests", "WhenSoftDeleted", "HiddenFromExchangeClientsEnabled",
                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, @{Name="ModeratedBy"; Expression={$_.ModeratedBy -join ','}}, "FolderPath",
                    @{Name="Description"; Expression={$_.Description -join ','}}, "WhenCreated"
                )

                $allUnifiedGroups = Get-UnifiedGroup -resultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue| Select $DesiredProperties
            }
            geek {$allUnifiedGroups = Get-UnifiedGroup -resultSize unlimited -IncludeSoftDeletedGroups -ErrorAction SilentlyContinue}
        }
        
        Write-Progress -Activity "Adding Unified Group data to Hash" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Adding Unified Group data to Hash" -ExportFileLocation $ExportDetails
        foreach ($group in $allUnifiedGroups) {
            #$key = $group.ExchangeGuid.ToString()
            $global:tenantStatsHash["UnifiedGroups"][$group.PrimarySmtpAddress] = $group
        }
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))

        # Get Unified Group Statistics
        Write-Log -Type INFO -Message "[Get-AllUnifiedGroups] Gathering all Unified Group Statistics" -ExportFileLocation $ExportDetails
        $allUnifiedGroupStatistics = $allUnifiedGroups | Get-EXOMailboxStatistics -IncludeSoftDeletedRecipients
        foreach ($groupStat in $allUnifiedGroupStatistics) {
            $key = $groupStat.MailboxGuid.ToString()
            $global:tenantStatsHash["PrimaryMailboxStats"][$key] = $groupStat
        } 
    }    
    catch {
        Write-Log -Type ERROR -Message "[Get-AllUnifiedGroups] An error occurred in running Get-AllUnifiedGroups function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-Progress -Activity "Getting all Unified Group data" -Completed
        Write-Progress -Activity "Adding Unified Group data to Hash" -Completed
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
    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["PublicFolderDetails"] = @{}
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
    try {
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Gathering all public folder permissions" -ExportFileLocation $ExportDetails
        $PublicFolderPermissions = $allPublicFolders | get-publicfolderclientpermission -ErrorAction SilentlyContinue
        #Progress Bar Parameters Reset
        $start = Get-Date
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $($PublicFolderPermissions.count) public folder permissions" -ExportFileLocation $ExportDetails

        $global:tenantStatsHash["PublicFolderPerms"] = @{}
        
        Write-Host "Processing Public Folder Permissions..." -ForegroundColor Cyan -nonewline
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Processing all public folder permissions" -ExportFileLocation $ExportDetails
        $totalCount = ($PublicFolderPermissions | Measure-Object).count
        foreach($publicFolderPermission in $PublicFolderPermissions) {
            Write-ProgressHelper -Total $totalCount -Activity "Processing all public folder permissions" -Operation "Gathering Public Folder Permissions for $($publicFolderPermission.Identity)"
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

            if($global:tenantStatsHash["PublicFolderPerms"].ContainsKey($key)) {
                $global:tenantStatsHash["PublicFolderPerms"][$key] += $permissionObject
            }
            else {
                $global:tenantStatsHash["PublicFolderPerms"][$key] = @($permissionObject)
            }
        }
        Write-ProgressHelper -Total $totalCount -Activity "Processing all public folder permissions" -Completed
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllPublicFolderDetails] An error occurred in running Get-AllPublicFolderPermissions function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }   
    
    #Combine Stats with Details
    $global:tenantStatsHash["PublicFolderDetails"] = @{}
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
        $global:tenantStatsHash["PublicFolderDetails"][$pf.Identity] = $pf
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
    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash['SharePoint'] = @{}
    $global:tenantStatsHash['OneDrive'] = @{}
    Write-Host "Getting all $($ServiceName) SharePoint Online and OneDrive Sites with $($detailLevel) ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type Info -Message "[Get-SharePointAndOneDriveSites] START: Getting all SharePoint Online and OneDrive $($detailLevel) details ($($ServiceName))" -ExportFileLocation $ExportDetails

    function Get-GraphCsvReportLookup {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Uri,
            [Parameter(Mandatory = $true)]
            [string]$LookupName
        )

        $tempFilePath = Join-Path $env:TEMP ("{0}-{1}.csv" -f $LookupName, [guid]::NewGuid().ToString('N'))
        try {
            Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputFilePath $tempFilePath -ErrorAction Stop | Out-Null
            $rows = @(Import-Csv -Path $tempFilePath -ErrorAction Stop)
            $lookup = @{}
            foreach ($row in $rows) {
                $siteId = $row.'Site Id'
                if (-not [string]::IsNullOrWhiteSpace($siteId) -and -not $lookup.ContainsKey($siteId)) {
                    $lookup[$siteId] = $row
                }
            }
            return $lookup
        } catch {
            Write-Log -Type WARNING -Message "[Get-SharePointAndOneDriveSites] Unable to download $LookupName report: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            return @{}
        } finally {
            if (Test-Path -Path $tempFilePath) {
                Remove-Item -Path $tempFilePath -Force -ErrorAction SilentlyContinue
            }
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
            $owner = (($url -split '/')[-1] -replace '_', '@')
        }

        return [PSCustomObject]@{
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
            StorageUsedGB             = [math]::Round(($storageUsageCurrent / 1024), 3)
            IsOffice365GroupsConnected = ($groupId -and $groupId -ne '00000000-0000-0000-0000-000000000000')
            IsOneDrive                = $IsOneDrive
        }
    }

    $sharePointUsageBySiteId = @{}
    $oneDriveUsageBySiteId = @{}
    if ($ServiceName -eq 'MGGraph') {
        $sharePointUsageBySiteId = Get-GraphCsvReportLookup -Uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D7')" -LookupName 'SharePointSiteUsageDetail'
        $oneDriveUsageBySiteId = Get-GraphCsvReportLookup -Uri "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D7')" -LookupName 'OneDriveUsageAccountDetail'
    }

    # Option 1: Use Microsoft Graph PowerShell SDK
    function Get-SharePointAndOneDriveSitesFromGraphSdk {
        try {
            Write-Verbose "Fetching SharePoint and OneDrive sites using Microsoft Graph SDK"
            $sites = @(Get-MgSite -All -ErrorAction Stop)
            $totalCount = $sites.Count
            $progressCounter = 0
            foreach ($site in $sites) {
                $progressCounter++
                $isOneDrive = ($site.WebUrl -like "*-my.sharepoint.com*")
                $reportSiteId = Get-GraphSiteReportId -CompositeSiteId $site.Id
                $usageReport = if ($isOneDrive) { $oneDriveUsageBySiteId[$reportSiteId] } else { $sharePointUsageBySiteId[$reportSiteId] }
                Write-ProgressHelper -Total $totalCount -Index $progressCounter -Activity "Gather Additional Site Details" -Operation "Gathering Site Details for $($site.DisplayName)"
                $siteData = ConvertTo-NormalizedSiteData -Site $site -IsOneDrive:$isOneDrive -Source MGGraph -UsageReport $usageReport

                if ($isOneDrive) {
                    $oneDriveKey = ConvertTo-OneDriveOwnerKey -Owner $siteData.Owner -Url $siteData.Url
                    if ($oneDriveKey) {
                        $global:tenantStatsHash['OneDrive'][$oneDriveKey] = $siteData
                    }
                } else {
                    $global:tenantStatsHash['SharePoint'][$siteData.Url] = $siteData
                }
            }
            Write-ProgressHelper -Total $totalCount -Activity "Gather Additional Site Details" -Completed
        } catch {
            Write-Error "Error fetching SharePoint and OneDrive sites with Graph SDK: $($_.Exception.Message)"
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
        try {
            Write-Log -Type INFO -Message "[Get-SharePointAndOneDriveSitesFromSPO] START: Gathering all SharePoint Online Sites with OneDrives with $($detailLevel) details" -ExportFileLocation $ExportDetails
            Write-Progress -Activity "Gather all SharePoint Online Sites with OneDrives" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
            switch ($detailLevel) {
                geek { $sites = Get-SPOSite -IncludePersonalSite $True -Limit All }
                Default {  
                    $DesiredProperties = @("Template", "IsHubSite", "Title", "LastContentModifiedDate", "Status", "ArchiveStatus", "StorageUsageCurrent", "LockState", "Url", "Owner", "StorageQuota", "GroupId", "IsTeamsConnected", "IsTeamsChannelConnected")
                    $sites = Get-SPOSite -IncludePersonalSite $True -Limit All | Select-Object -Property $DesiredProperties
                }
            }
            
            $start = Get-Date
            $totalCount = ($sites | Measure-Object).count
            foreach ($site in $sites) {
                Write-ProgressHelper -Total $totalCount -Activity "Gather Additional Site Details" -Operation "Gathering Site Details for $($site.Title)"
                # Determine if the site is a OneDrive or a standard SharePoint site
                $isOneDrive = ($site.Url -like "*-my.sharepoint.com*")
                $siteData = ConvertTo-NormalizedSiteData -Site $site -IsOneDrive:$isOneDrive -Source SPO

                # Store data in appropriate hashtable
                if ($isOneDrive) {
                    $oneDriveKey = ConvertTo-OneDriveOwnerKey -Owner $siteData.Owner -Url $siteData.Url
                    if ($oneDriveKey) {
                        $global:tenantStatsHash['OneDrive'][$oneDriveKey] = $siteData
                    }
                } else {
                    $global:tenantStatsHash['SharePoint'][$siteData.Url] = $siteData
                }
            }
            Write-Progress -Activity "Gather Additional Site Details" -Completed
        } catch {
            Write-Log -Type ERROR -Message "[Get-SharePointAndOneDriveSitesFromSPO] An error occurred in running Get-SharePointAndOneDriveSitesFromSPO function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }
    }

    # Option 3: Use REST API
    function Get-SharePointAndOneDriveSitesFromRESTAPI {
        $allSitesUri = "https://graph.microsoft.com/v1.0/sites?search=*"
        $pageCount = 0
        
        try {
            Write-Verbose "Fetching initial SharePoint and OneDrive sites via REST API"
            $pageCount++
            $initialResponse = Invoke-RestMethod -Uri $allSitesUri -Headers $global:GraphHeaders -Method Get -ContentType "application/json" -ErrorAction Stop
            $sites = $initialResponse.value
            Write-Progress -Activity "Fetching Sites" -ID 1 -Status "Processing page $pageCount"

            $allSitesUri = $initialResponse.'@odata.nextLink'
            while ($allSitesUri) {
                $pageCount++
                Write-Verbose "Fetching next page of sites via REST API. Page $pageCount"
                $response = Invoke-RestMethod -Uri $allSitesUri -Headers $global:GraphHeaders -Method Get -ContentType "application/json" -ErrorAction Stop
                $sites += $response.value
                $allSitesUri = $response.'@odata.nextLink'
                Write-Progress -Activity "Fetching Sites" -ID 1 -Status "Processing page $pageCount"
            }

            # Filter and store relevant sites in the global hashtable
            foreach ($site in $sites) {
                # Determine if the site is a OneDrive or a standard SharePoint site
                $isOneDrive = ($site.webUrl -like "*-my.sharepoint.com*")
                $siteData = ConvertTo-NormalizedSiteData -Site $site -IsOneDrive:$isOneDrive -Source API

                # Store data in appropriate hashtable
                if ($isOneDrive) {
                    $oneDriveKey = ConvertTo-OneDriveOwnerKey -Owner $siteData.Owner -Url $siteData.Url
                    if ($oneDriveKey) {
                        $global:tenantStatsHash['OneDrive'][$oneDriveKey] = $siteData
                    }
                } else {
                    $global:tenantStatsHash['SharePoint'][$siteData.Url] = $siteData
                }
            }
        } catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                Write-Host "Throttling detected. Please try again later." -ForegroundColor Yellow
            } else {
                Write-Host "Error fetching sites via REST API: $($_.Exception.Message)" -ForegroundColor Red
            }
        } finally {
            Write-Progress -Activity "Fetching Sites" -ID 1 -Completed
        }
    }

# Determine which method to use based on ServiceName
    switch ($ServiceName) {
        'MGGraph' { Get-SharePointAndOneDriveSitesFromGraphSdk }
        'SPO'     { Get-SharePointAndOneDriveSitesFromSPO -detailLevel $detailLevel }
        'API'     { Get-SharePointAndOneDriveSitesFromRESTAPI }
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
        $global:tenantStatsHash["AllTeams"] = @{}

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
                    $SPOSiteDetails = $global:tenantStatsHash["SharePointSites"][$team.displayName]
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
                    $global:tenantStatsHash["AllTeams"][$team.DisplayName] = $currentTeam
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
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["TeamsVoice"] = @{}

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
        if ($global:tenantStatsHash.ContainsKey('Users')) {
            $users = @($global:tenantStatsHash['Users'].Values)
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

        $global:tenantStatsHash["TeamsVoice"]["Summary"] = $summary
        $global:tenantStatsHash["TeamsVoice"]["PstnUsage"] = $pstnUsage
        $global:tenantStatsHash["TeamsVoice"]["CallingPolicies"] = $callingPolicies
        $global:tenantStatsHash["TeamsVoice"]["PhoneNumbers"] = $phoneAssignments
        $global:tenantStatsHash["TeamsVoice"]["VoiceUsers"] = $voiceUsers

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

    $global:tenantStatsHash["TeamsVoice"]["Summary"] = $summary
    $global:tenantStatsHash["TeamsVoice"]["PstnUsage"] = $pstnUsage
    $global:tenantStatsHash["TeamsVoice"]["CallingPolicies"] = $callingPolicies
    $global:tenantStatsHash["TeamsVoice"]["PhoneNumbers"] = $phoneAssignments
    $global:tenantStatsHash["TeamsVoice"]["VoiceUsers"] = $voiceUsers

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
        # Initialize progress tracking
        $pageCount = 0
        $totalRecordsRetrieved = 0

        # Prepare authentication headers for REST fallback
        $Headers = $null
        if ($global:GraphHeaders) {
            $Headers = $global:GraphHeaders
        } elseif ($global:GraphToken) {
            $Headers = @{
                'Content-Type'     = "application/json"
                'Authorization'    = "Bearer $global:GraphToken"
                'ConsistencyLevel' = "eventual"
            }
        } elseif ($AccessToken) {
            $Headers = @{
                'Content-Type'     = "application/json"
                'Authorization'    = "Bearer $AccessToken"
                'ConsistencyLevel' = "eventual"
            }
        }

        # Build query URI with pagination
        $QueryUri = $Uri
        if ($PageSize -and $QueryUri -notmatch "\`$top=") {
            $separator = if ($QueryUri.Contains('?')) { '&' } else { '?' }
            $QueryUri += "${separator}`$top=$PageSize"
        }

        # Initialize results collection
        $QueryResults = [System.Collections.Generic.List[PSObject]]::new()
        
        Write-Verbose "Starting Get-GraphData for URI: $QueryUri"
    }

    process {
        $CurrentUri = $QueryUri
        $MorePages = $true

        do {
            $pageCount++
            $Results = $null
            $RetryCount = 0
            $Success = $false

            # Retry loop for handling throttling and transient errors
            while (-not $Success -and $RetryCount -lt $MaxRetries) {
                try {
                    # Primary method: Use Microsoft Graph PowerShell SDK
                    if (-not $UseRestMethod) {
                        Write-Verbose "Page $pageCount : Invoke-MgGraphRequest => $CurrentUri"
                        $Results = Invoke-MgGraphRequest -Uri $CurrentUri -Method GET -OutputType PSObject -ErrorAction Stop
                        $Success = $true
                    } else {
                        # Fallback: Use REST method
                        Write-Verbose "Page $pageCount : Invoke-RestMethod => $CurrentUri"
                        if (-not $Headers) {
                            throw "No authentication headers available for Invoke-RestMethod. Please ensure you're connected to Microsoft Graph."
                        }
                        $Results = Invoke-RestMethod -Uri $CurrentUri -Headers $Headers -Method GET -ContentType "application/json" -UseBasicParsing -ErrorAction Stop
                        $Success = $true
                    }
                }
                catch {
                    $statusCode = $null
                    $retryAfter = 5

                    # Extract status code if available
                    if ($_.Exception.Response) {
                        $statusCode = $_.Exception.Response.StatusCode.value__
                        # Check for Retry-After header
                        if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers['Retry-After']) {
                            $retryAfter = [int]$_.Exception.Response.Headers['Retry-After']
                        }
                    }

                    # Handle specific error codes
                    switch ($statusCode) {
                        429 {
                            # Throttling - exponential backoff
                            $waitTime = [math]::Min(($retryAfter * [math]::Pow(2, $RetryCount)), 300) # Max 5 minutes
                            Write-Warning "Throttled (429). Waiting $waitTime seconds before retry $($RetryCount + 1)/$MaxRetries..."
                            Start-Sleep -Seconds $waitTime
                            $RetryCount++
                        }
                        401 {
                            # Unauthorized - token might be expired
                            Write-Warning "Authentication failed (401). Please reconnect to Microsoft Graph."
                            throw $_
                        }
                        403 {
                            # Forbidden - insufficient permissions
                            Write-Warning "Access denied (403). Insufficient permissions for: $CurrentUri"
                            throw $_
                        }
                        404 {
                            # Not found - return empty
                            Write-Warning "Resource not found (404): $CurrentUri"
                            return @()
                        }
                        503 {
                            # Service unavailable - retry with backoff
                            $waitTime = [math]::Min((5 * [math]::Pow(2, $RetryCount)), 60)
                            Write-Warning "Service unavailable (503). Waiting $waitTime seconds before retry $($RetryCount + 1)/$MaxRetries..."
                            Start-Sleep -Seconds $waitTime
                            $RetryCount++
                        }
                        default {
                            # For SDK failure, try REST fallback once
                            if (-not $UseRestMethod -and $Headers) {
                                Write-Warning "Invoke-MgGraphRequest failed: $($_.Exception.Message). Falling back to Invoke-RestMethod..."
                                $UseRestMethod = $true
                                $RetryCount++
                            } else {
                                Write-Error "Graph API request failed: $($_.Exception.Message)"
                                throw $_
                            }
                        }
                    }
                }
            }

            # Check if max retries exceeded
            if (-not $Success) {
                Write-Error "Failed to retrieve data after $MaxRetries attempts from: $CurrentUri"
                break
            }

            # Process results
            if ($Results.value) {
                # Standard Graph response with value property
                foreach ($item in $Results.value) {
                    $QueryResults.Add([PSObject]$item)
                }
                $totalRecordsRetrieved += $Results.value.Count
            } elseif ($Results -is [System.Collections.IEnumerable] -and -not $Results.PSObject.Properties['value']) {
                # Collection without value property (some reports)
                foreach ($item in $Results) {
                    $QueryResults.Add([PSObject]$item)
                }
                $totalRecordsRetrieved += $Results.Count
            } elseif ($Results) {
                # Single object response
                $QueryResults.Add([PSObject]$Results)
                $totalRecordsRetrieved++
            }

            # Update progress
            $progressParams = @{
                Total     = [math]::Max($totalRecordsRetrieved, 1)  # Prevent divide by zero
                Activity  = $Activity
                Operation = "$Operation (Page $pageCount, $totalRecordsRetrieved records)"
                Id        = $Id
            }
            if ($PSBoundParameters.ContainsKey('ParentId')) {
                $progressParams.ParentId = $ParentId
            }
            Write-ProgressHelper @progressParams

            # Check for next page
            $NextLink = $null
            if ($Results.'@odata.nextLink') {
                $NextLink = $Results.'@odata.nextLink'
            } elseif ($Results.PSObject.Properties['nextLink']) {
                $NextLink = $Results.nextLink
            }

            if ($NextLink) {
                $CurrentUri = $NextLink
                Write-Verbose "Next page available: $NextLink"
            } else {
                $MorePages = $false
                Write-Verbose "No more pages. Total records retrieved: $totalRecordsRetrieved"
            }

        } while ($MorePages)
    }

    end {
        # Complete the progress bar
        Write-ProgressHelper -Total 1 -Activity $Activity -Id $Id -Completed
        $ProgressPreference = "SilentlyContinue"
        $ProgressPreference = "Continue"

        # Return results
        Write-Verbose "Returning $($QueryResults.Count) total results"
        if ($QueryResults.Count -eq 0) {
            return @()
        } else {
            return $QueryResults.ToArray()
        }
    }
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
                                Write-Verbose "Running Connect-MgGraph with -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint"
                                Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop | Out-Null
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
                        Write-Host "✗ Graph: $($_.Exception.Message)" -ForegroundColor Red
                        $result.Graph = $false
                        return
                    }
                }
                
                Write-Verbose "Retrieving organization info from Graph..."
                $org = Get-MgOrganization -ErrorAction Stop
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
                        Write-Warning "Exchange Online does not support ClientSecretCredential (App/Secret) authentication. Falling back to delegated authentication."
                        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
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
                        Write-Warning "Exchange Online does not support ClientSecretCredential (App/Secret) authentication. Falling back to delegated authentication."
                        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
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

function Get-FriendlyProductName {
    param(
        [string]$SkuPartNumber
    )
    if ([string]::IsNullOrWhiteSpace($SkuPartNumber)) {
        return $SkuPartNumber
    }
    if ($script:CommonProductNameMapStatic -and $script:CommonProductNameMapStatic.ContainsKey($SkuPartNumber)) {
        return $script:CommonProductNameMapStatic[$SkuPartNumber]
    }
    return $SkuPartNumber
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
    }

    # Build a hashtable for license sku. Create start time of the function
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["LicenseSKUs"] = @{}
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

        Write-Log -Type DEBUG -Message "[Get-AllLicenseSKUs] Gathering License details for $($AccountSkuId)" -ExportFileLocation $ExportDetails

        #Create Hash Table for License SKUs
        $global:tenantStatsHash["LicenseSKUs"][$sku.SkuId.tostring()] = $skuDetails

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
    #Get the start time of the function
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["Users"] = @{} #Hash table to store all user details

    # Gather all Microsoft Graph User Details
    try {
        Write-Host "Getting all Microsoft Graph $($detailLevel) User data..." -ForegroundColor Cyan -nonewline
        Write-Log -Type Info -Message "[Get-allUserDetails] START: Getting all Microsoft Graph $($detailLevel) User data" -ExportFileLocation $ExportDetails
        Write-Progress -Activity "Getting all Microsoft Graph User Data" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))

        $maxRetries = 3
        $success = $false
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                switch ($detailLevel) {
                    #Not Sure this will pull all the details for a user. Need to review Property variable. Might need to include all properties that expand further
                    geek { $allTenantUsers = Get-MgUser -all -ErrorAction Stop }
                    default { 
                        $DesiredProperties = @(
                            "DisplayName", "AssignedLicenses", "UserPrincipalName"
                            "UserType", "Id", "AccountEnabled"
                            "CreatedDateTime", "Mail", "JobTitle"
                            "Department", "CompanyName", "OfficeLocation"
                            "City", "State", "Country"
                            "OnPremisesSyncEnabled", "OnPremisesDistinguishedName", "OnPremisesLastSyncDateTime"
                            "UsageLocation", "SignInActivity", "ProxyAddresses"
                        )
                        $allTenantUsers = Get-MgUser -all -Property $DesiredProperties -ErrorAction Stop | Select-Object $DesiredProperties | Where-Object { $null -ne $_.ID }
                    }
                }
                $success = $true
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
        Write-Progress -Activity "Getting all Microsoft Graph User Data" -Completed
    }
    catch {
        # Using the Capture-ErrorHelper function to capture and log the error.
        if ($_.Exception.Message -like "*Neither tenant is B2C or tenant doesn't have premium license*") {
            #Run if Error received is Get-MgUser : Neither tenant is B2C or tenant doesn't have premium license
            #Status: 403 (Forbidden)
            #ErrorCode: Authentication_RequestFromNonPremiumTenantOrB2CTenant
            Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            
            Write-Host
            Write-Host "Caught a tenant license exception. Getting all Microsoft Graph User data without licenses and sign in activity..." -ForegroundColor Yellow -nonewline    
            try {
                #Fallback has bug that doesn't return all properties to help build licenses and sign in activity
                Write-Log -Type Info -Message "[Get-allUserDetails] Attempt 2. Getting all Microsoft Graph $($detailLevel) with limited User Details" -ExportFileLocation $ExportDetails
                $allTenantUsers = Get-MgUser -all -ErrorAction Stop
                $BasicMGDetails = $true
            }
            catch {
                Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                
                }
        }
        else {
            Write-Log -Type Error -Message "[Get-allUserDetails] An error occurred in running Get-allMGUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            Write-Log -Type WARNING -Message "[Get-allUserDetails] Continuing without user details due to Graph request failure." -ExportFileLocation $ExportDetails
            $allTenantUsers = @()
            return
            # Handle other exceptions
        }
    }
    #Add Additional Properties to Hash Table - Licensing, MailboxStats, OneDriveStats, ArchiveStats
    Write-Log -Type Info -Message "[Get-allUserDetails] Adding Additional Properties for $($allTenantUsers.count) Users" -ExportFileLocation $ExportDetails
    try {
        $totalCount = $allTenantUsers.count
        $licensedUserCount = 0
        $unlicensedUserCount = 0
        $licenseFallbackUserCount = 0
        $unresolvedSkuCount = 0
        foreach ($user in $allTenantUsers) {
            try {
                # Create Hash Table for each user
                Write-ProgressHelper -Total $totalCount -Activity "Gathering Tenant User Details" -Operation "Gathering Tenant User Details for $($user.DisplayName)"
                Write-Log -Type DEBUG -Message ("[Get-allUserDetails] Creating Hash for '{0}'" -f $user.UserPrincipalName) -ExportFileLocation $ExportDetails
                # New hashtable to build upon the existing properties
                $global:tenantStatsHash["Users"][$user.UserPrincipalName] = [PSCustomObject]@{}
                $customUserDetails = [PSCustomObject]@{}

                # Populate the existing properties
                Write-Log -Type DEBUG -Message ("[Get-allUserDetails] Create '{0}' Hash with Existing Properties" -f $user.UserPrincipalName) -ExportFileLocation $ExportDetails

                foreach ($property in $user.PSObject.Properties) {
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
                }

                # Add specific additional properties - MGGraph
                Write-Log -Type DEBUG -Message ("[Get-allUserDetails] Adding '{0}' Specific Additional Properties (MGGraph)" -f $user.UserPrincipalName) -ExportFileLocation $ExportDetails

                #Combine ProxyAddresses
                $combinedProxyAddresses = ($user.ProxyAddresses -replace '^[sS][mM][tT][pP]:') -join ';'
                $customUserDetails | Add-Member -MemberType NoteProperty -Name "ProxyAddresses" -Value $combinedProxyAddresses -Force

                if ($BasicMGDetails) {
                    Write-Log -Type DEBUG -Message ("[Get-allUserDetails] Updating '{0}' UserType to HashTable if Basic Details" -f $user.UserPrincipalName) -ExportFileLocation $ExportDetails
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name "UserType" -Value (if ($user.UserPrincipalName -like "*#EXT#*") { "GuestUser" } else { "User" })
                }
                else {
                    Write-Log -Type DEBUG -Message ("[Get-allUserDetails] Gather '{0}' License Friendly Names" -f $user.UserPrincipalName) -ExportFileLocation $ExportDetails
                    $assignedLicenses = $null
                    $assignedLicensesFriendly = $null
                    $disabledPlans = $null
                    $enabledServicePlans = $null
                    $assignedLicenseEntries = @($user.AssignedLicenses)
                    if ($assignedLicenseEntries.Count -gt 0) {
                        $licensedUserCount++
                        $resolvedSkuParts = New-Object System.Collections.Generic.List[string]
                        $resolvedFriendlyNames = New-Object System.Collections.Generic.List[string]
                        $resolvedDisabledPlans = New-Object System.Collections.Generic.List[string]
                        $resolvedEnabledPlans = New-Object System.Collections.Generic.List[string]

                        foreach ($assignedLicenseEntry in $assignedLicenseEntries) {
                            $skuId = $null
                            if ($assignedLicenseEntry.PSObject.Properties['SkuId']) {
                                $skuId = $assignedLicenseEntry.SkuId
                            } elseif ($assignedLicenseEntry -is [guid]) {
                                $skuId = $assignedLicenseEntry
                            }

                            $skuIdText = if ($skuId) { $skuId.ToString() } else { $null }
                            $skuLookup = if ($skuIdText -and $script:SkuLookupById.ContainsKey($skuIdText)) { $script:SkuLookupById[$skuIdText] } else { $null }
                            if ($skuLookup) {
                                $resolvedSkuParts.Add($skuLookup.SkuPartNumber)
                                $resolvedFriendlyNames.Add($skuLookup.FriendlyName)

                                $disabledPlanIds = @()
                                if ($assignedLicenseEntry.PSObject.Properties['DisabledPlans'] -and $assignedLicenseEntry.DisabledPlans) {
                                    $disabledPlanIds = @($assignedLicenseEntry.DisabledPlans | ForEach-Object { $_.ToString() })
                                }

                                $disabledPlanNamesForSku = @()
                                foreach ($disabledPlanId in $disabledPlanIds) {
                                    if ($script:ServicePlanLookupById.ContainsKey($disabledPlanId)) {
                                        $disabledPlanName = $script:ServicePlanLookupById[$disabledPlanId]
                                        $disabledPlanNamesForSku += $disabledPlanName
                                        $resolvedDisabledPlans.Add($disabledPlanName)
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
                                    $resolvedEnabledPlans.Add($enabledPlanName)
                                }
                            } else {
                                $unresolvedSkuCount++
                                if ($skuIdText) {
                                    $resolvedSkuParts.Add($skuIdText)
                                    $resolvedFriendlyNames.Add($skuIdText)
                                }
                            }
                        }

                        $licenseFallbackUserCount++
                        $AssignedLicenses = ($resolvedSkuParts | Select-Object -Unique) -join ","
                        $AssignedLicensesFriendly = ($resolvedFriendlyNames | Select-Object -Unique) -join ","
                        $disabledPlans = $resolvedDisabledPlans | Select-Object -Unique
                        $EnabledServicePlans = ($resolvedEnabledPlans | Select-Object -Unique) -join ","
                    }
                    else {
                        $unlicensedUserCount++
                    }
                    # Create Current User Object - Additional Attributes custom object
                    #$customUserDetails | Add-Member -MemberType NoteProperty -Name "UserType" -Value $user.UserType
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name "AssignedLicenses" -Value $AssignedLicenses -Force
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name "AssignedLicensesFriendly" -Value $AssignedLicensesFriendly -Force
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name "License-DisabledArray" -Value $disabledPlans
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name "EnabledServicePlans" -Value $EnabledServicePlans
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name "LastNonInteractiveSignInDateTime" -Value $user.SignInActivity.LastNonInteractiveSignInDateTime
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name "LastNonInteractiveSignInRequestId" -Value $user.SignInActivity.LastNonInteractiveSignInRequestId
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name "LastSignInDateTime" -Value $user.SignInActivity.LastSignInDateTime
                    $customUserDetails | Add-Member -MemberType NoteProperty -Name "LastSignInRequestId" -Value $user.SignInActivity.LastSignInRequestId
                }

                $global:tenantStatsHash["Users"][$user.UserPrincipalName] = $customUserDetails
            }
            catch {
                Write-Log -Type ERROR -Message ("[Get-allUserDetails] An error occurred in Creating User Hash for user '{0}'. $($_.Exception.Message)" -f $user.UserPrincipalName) -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
        }
        Write-ProgressHelper -Total $totalCount -Activity "Gathering Tenant User Details" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-allUserDetails] Licensing summary: LicensedUsers=$licensedUserCount UnlicensedUsers=$unlicensedUserCount LicenseLookupUsers=$licenseFallbackUserCount UnresolvedSkuReferences=$unresolvedSkuCount" -ExportFileLocation $ExportDetails
        Write-Log -Type Info -Message "[Get-allUserDetails] COMPLETED: Gathering all  User Details in $($CompletedTime)" -ExportFileLocation $ExportDetails
    }     
    catch {
        Write-Log -Type ERROR -Message "[Get-allUserDetails] An error occurred in running Get-allUserDetails function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
}

#Gather all Office 365 Admins
function Get-AllOffice365Admins {
    param ()
    $start = Get-Date
    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["Admins"] = @{}
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
            $roleMemberList = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id | ? {$null -ne $_.Id}
            if ($roleMemberList) {
                $totalCount2 = $roleMemberList.count
                Write-Log -Type INFO -Message "[Get-AllOffice365Admins] $($roleName) Users Found: $($roleMemberList.count)" -ExportFileLocation $ExportDetails
                foreach ($roleMember in $roleMemberList) {
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

                        # Check if user is in the tenantStatsHash
                        if ($global:tenantStatsHash['users'].ContainsKey($UPN)) {
                            $userMatch = $global:tenantStatsHash['users'][$UPN]
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
                        GroupMailEnabled = if ($GroupMailEnabled) { $mailEnabled } else { $null }
                        GroupMailNickname = if ($GroupMailNickname) { $mailNickname } else { $null }
                        GroupType = if ($GroupType) { $GroupType } else { $null }
                    })
                        
                    Write-ProgressHelper -Total $totalCount2 -Id 2 -ParentId 1 -Activity "Gathering Admin Details" -Operation "Gathering Admin Role Details: $($roleMember.DisplayName)"
    
                    $adminResults.Add($currentAdmin)
                }
            }
            
        }
        
        Write-Log -Type INFO -Message "[Get-AllOffice365Admins] Combined Group Roles for $($groupedResults)" -ExportFileLocation $ExportDetails
        #Group by DisplayName or UserPrincipalName and combine roles into a comma-separated list
        $groupedResults = $adminResults | Group-Object -Property DisplayName,UserPrincipalName
        $finalResults = $groupedResults | ForEach-Object {
            $group = $_.Group
            $roles = ($group.Role -join ', ')
            # Use the properties of the first user in each group, but replace the Role with the combined roles
            $group[0] | Add-Member -MemberType NoteProperty -Name 'Role' -Value $roles -Force
            $group[0] | Add-Member -MemberType NoteProperty -Name 'RolesAssigned' -Value (($group.Role | Measure-Object).count) -Force
            $group[0]
        }

        foreach ($result in $finalResults) {
            $global:tenantStatsHash["Admins"][$result.DisplayName] = $result
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllOffice365Admins] An error occurred in running Get-AllOffice365Admins function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-ProgressHelper -Total $totalCount -Id 1 -Activity "Gathering Admins in Roles" -Completed
        Write-ProgressHelper -Total 1 -Id 2 -Activity "Checking Admin Role Details" -Completed
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
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["Domains"] = @{}
    $global:tenantStatsHash["RemoteDomains"] = @{}
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
            if ($global:tenantStatsHash['AllRecipients']) {
                $recipients = $global:tenantStatsHash['AllRecipients'].values
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
            
                    # ===== Recipient Counts =====
                    PrimarySMTPRecipients = $RecipientCounts.PrimarySMTPCount
                    AliasOnlyRecipients   = $RecipientCounts.AliasOnlyCount
                    TotalDomainRecipients = $RecipientCounts.TotalDomainRecipientsCount
                }
                $global:tenantStatsHash["Domains"][$domainName] = $currentDomain
            }
            catch {
                Write-Log -Type ERROR -Message "An error occurred in running Get-AllOffice365Domains function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
        }
        foreach ($domain in $remoteDomains) {
            # Add to the results Hash Table
            $global:tenantStatsHash["RemoteDomains"][$domain.Identity] = $domain
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

    # Helper function to populate details
    function Populate-Details {
        param (
            [Parameter(Mandatory = $true)]
            [hashtable]$tenantStatsHash,
            [Parameter(Mandatory = $true)]
            [object]$entity,  # This can be a mailbox or user object
            [Parameter(Mandatory = $false)]
            [switch]$IsMailbox  # Differentiates whether the entity is a mailbox or user
        )
        $details = [PSCustomObject]@{}
        try {
            # Populate existing entity-specific properties
            foreach ($property in $entity.PSObject.Properties) {
                $details | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
            }
    
            # Determine mailbox-specific details
            $mailboxDetails = if ($IsMailbox -and $entity.RecipientTypeDetails -eq "GroupMailbox" -and $tenantStatsHash["UnifiedGroups"]) {
                # Check for Group Mailbox Details
                if ($tenantStatsHash["UnifiedGroups"].ContainsKey($entity.PrimarySMTPAddress)) {
                    $tenantStatsHash["UnifiedGroups"][$entity.PrimarySMTPAddress]
                    Write-Verbose "Unified Group Details Found: $($entity.PrimarySMTPAddress)"
                } else { $null }
            } elseif ($IsMailbox -and $tenantStatsHash["AllMailboxes-MailIdentity"]) {
                # Check against Mailbox Identity
                if ($tenantStatsHash["AllMailboxes-MailIdentity"].ContainsKey($entity.Identity)) {
                    $tenantStatsHash["AllMailboxes-MailIdentity"][$entity.Identity]
                    Write-Verbose "Mailbox Details Found: $($entity.Identity)"
                } else { $null }
            } elseif ($tenantStatsHash["AllMailboxes"] -and $entity.UserPrincipalName) {
                # Check against UserPrincipalName
                if ($tenantStatsHash["AllMailboxes"].ContainsKey($entity.UserPrincipalName)) {
                    $tenantStatsHash["AllMailboxes"][$entity.UserPrincipalName]
                    Write-Verbose "Mailbox Details Found (UPN): $($entity.UserPrincipalName)"
                } else { $null }
            } else { 
                Write-Verbose "Mailbox Details Not Found"
                $null }
    
            $mailboxStats = if ($mailboxDetails -and $tenantStatsHash["PrimaryMailboxStats"]) {
                if ($tenantStatsHash["PrimaryMailboxStats"].ContainsKey($mailboxDetails.ExchangeGuid.ToString())) {
                    $tenantStatsHash["PrimaryMailboxStats"][$mailboxDetails.ExchangeGuid.ToString()]
                    Write-Verbose "Mailbox Stats Found: $($mailboxDetails.ExchangeGuid.ToString())"
                } else { $null }
            } else { 
                Write-Verbose "Mailbox Stats Not Found"
                $null }
    
            $archiveStats = if ($mailboxDetails -and $tenantStatsHash["ArchiveMailboxStats"] -and $mailboxDetails.ArchiveGuid) {
                if ($tenantStatsHash["ArchiveMailboxStats"].ContainsKey($mailboxDetails.ArchiveGuid.ToString())) {
                    $tenantStatsHash["ArchiveMailboxStats"][$mailboxDetails.ArchiveGuid.ToString()]
                    Write-Verbose "Archive Stats Found: $($mailboxDetails.ArchiveGuid.ToString())"
                } else { $null }
            } else { 
                Write-Verbose "Archive Stats Not Found"
                $null }
    
            # Determine drive-specific details (OneDrive or GroupMailbox SharePoint)
            $DriveData = if ($IsMailbox -and $entity.RecipientTypeDetails -eq "GroupMailbox" -and $tenantStatsHash["SharePoint"] -and $mailboxDetails.SharePointSiteUrl) {
                if ($tenantStatsHash["SharePoint"].ContainsKey($mailboxDetails.SharePointSiteUrl)) {
                    $tenantStatsHash["SharePoint"][$mailboxDetails.SharePointSiteUrl]
                    Write-Verbose "SharePoint Details Found: $($mailboxDetails.SharePointSiteUrl)"
                } else { $null }
            } elseif ($IsMailbox -and $tenantStatsHash["OneDrive"] -and $mailboxDetails.UserPrincipalName -and $tenantStatsHash["OneDrive"].ContainsKey($mailboxDetails.UserPrincipalName)) {
                $tenantStatsHash["OneDrive"][$mailboxDetails.UserPrincipalName]
                Write-Verbose "OneDrive Details Found: $($mailboxDetails.UserPrincipalName)"
            } elseif ($tenantStatsHash["OneDrive"] -and $entity.UserPrincipalName -and $tenantStatsHash["OneDrive"].ContainsKey($entity.UserPrincipalName)) {
                $tenantStatsHash["OneDrive"][$entity.UserPrincipalName]
                Write-Verbose "OneDrive Details Found: $($entity.UserPrincipalName)"
            }  else { $null }
    
            # Add mailbox stats, ensuring null safety
            $MBXSizeGB = if ($mailboxStats -and $mailboxStats.TotalItemSize) {
                [math]::Round(($mailboxStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)", "") / 1GB, 3)
            } else { 0 }
    
            $MBXItemCount = if ($mailboxStats -and $mailboxStats.ItemCount) {
                $mailboxStats.ItemCount
            } else { 0 }
    
            # Add archive stats, ensuring null safety
            $ArchiveSizeGB = if ($archiveStats -and $archiveStats.TotalItemSize) {
                [math]::Round(($archiveStats.TotalItemSize.ToString() -replace "(.*\()|,| [a-z]*\)", "") / 1GB, 3)
            } else { 0 }
    
            $ArchiveItemCount = if ($archiveStats -and $archiveStats.ItemCount) {
                $archiveStats.ItemCount
            } else { 0 }
    
            # Add drive stats, ensuring null safety
            $DriveURL = if ($DriveData -and $DriveData.URL) {
                $DriveData.URL
            } else { $null }
    
            $DriveStorageGB = if ($DriveData -and $DriveData.StorageUsageCurrent) {
                [math]::Round($DriveData.StorageUsageCurrent / 1024, 3)
            } else { 0 }
    
            # Add the computed values to the details object
            $details | Add-Member -MemberType NoteProperty -Name "MBXSizeGB" -Value $MBXSizeGB
            $details | Add-Member -MemberType NoteProperty -Name "MBXItemCount" -Value $MBXItemCount
            $details | Add-Member -MemberType NoteProperty -Name "ArchiveSizeGB" -Value $ArchiveSizeGB
            $details | Add-Member -MemberType NoteProperty -Name "ArchiveItemCount" -Value $ArchiveItemCount
            $details | Add-Member -MemberType NoteProperty -Name "DriveURL" -Value $DriveURL
            $details | Add-Member -MemberType NoteProperty -Name "DriveStorageGB" -Value $DriveStorageGB
        } catch {
            Write-Log -Type ERROR -Message "[Populate-Details] An error occurred in Populating Details for $($entity.PrimarySMTPAddress). $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        }
    
        return $details
    }
    
    # Combine User and Mailbox Stats with Progress
    function Combine-UserAndMailboxStats {
        param (
            [hashtable]$tenantStatsHash
        )

        # Initialize variables. Include the start time for the script and logging
        $start = Get-Date
        # Ensure global hash table structure
        if (-not $global:tenantStatsHash) {
            $global:tenantStatsHash = @{}
        }
        # Create hash tables for user and mailbox details
        $global:tenantStatsHash["UserFullDetails"] = @{}
        $global:tenantStatsHash["MailboxFullDetails"] = @{}
        
        # Process Users
        $userCount = $tenantStatsHash["Users"].Keys.Count

        foreach ($userKey in $tenantStatsHash["Users"].Keys) {
            $user = $tenantStatsHash["Users"][$userKey]
            Write-Log -Type DEBUG -Message ("[Combine-UserAndMailboxStats] Combining User '{0}' Details" -f $user.DisplayName) -ExportFileLocation $ExportDetails

            Write-ProgressHelper -Total $userCount -Activity "Processing User Data" -Operation "Processing user: $($user.DisplayName)"

            $userDetails = Populate-Details -tenantStatsHash $tenantStatsHash -entity $user
            $global:tenantStatsHash["UserFullDetails"][$user.UserPrincipalName] = $userDetails
        }

        # Process Mailboxes
        $allMailboxes = $tenantStatsHash['AllRecipients'].Values | Where-Object { $_.RecipientTypeDetails -like "*Mailbox" }
        $mailboxCount = $allMailboxes.Count
        Write-Log -Type INFO -Message "[Combine-UserAndMailboxStats] Combining User and Mailbox Details - Processing Mailboxes" -ExportFileLocation $ExportDetails
        foreach ($mailbox in $allMailboxes) {
            #$mailbox = $tenantStatsHash["AllMailboxes"][$mailbox.PrimarySMTPAddress]
            Write-Log -Type DEBUG -Message ("[Combine-UserAndMailboxStats] Combining Mailbox '{0}' Details" -f $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails
            Write-ProgressHelper -Total $mailboxCount -Activity "Processing Mailbox Data" -Operation "Processing mailbox: $($mailbox.PrimarySMTPAddress)"

            $mailboxDetails = Populate-Details -tenantStatsHash $tenantStatsHash -entity $mailbox -IsMailbox
            $global:tenantStatsHash["MailboxFullDetails"][$mailbox.PrimarySMTPAddress] = $mailboxDetails
        }

        Write-ProgressHelper -Total $mailboxCount -Activity "Processing Complete" -Completed
    }
    # New function to generate Inactive Mailboxes report
    function Report-InactiveMailboxes {
        [CmdletBinding()]
        param (
            [hashtable]$tenantStatsHash
        )

        # Initialize variables. Include the start time for the script and logging
        $start = Get-Date
        # Ensure global hash table structure
        if (-not $global:tenantStatsHash) {
            $global:tenantStatsHash = @{}
        }

        # Create hash table for inactive mailbox details
        $global:tenantStatsHash["InactiveMailboxDetails"] = @{}

        # Process Inactive Mailboxes
        $inactiveMailboxes = $tenantStatsHash['InactiveMailboxes'].Values
        $inactiveMailboxCount = $inactiveMailboxes.Count
        Write-Log -Type INFO -Message "[Report-InactiveMailboxes] Processing Inactive Mailboxes" -ExportFileLocation $ExportDetails
        foreach ($mailbox in $inactiveMailboxes) {
            Write-Log -Type DEBUG -Message ("[Report-InactiveMailboxes] Processing Inactive Mailbox '{0}'" -f $mailbox.PrimarySMTPAddress) -ExportFileLocation $ExportDetails
            Write-ProgressHelper -Total $inactiveMailboxCount -Activity "Processing Inactive Mailbox Data" -Operation "Processing inactive mailbox: $($mailbox.PrimarySMTPAddress)"

            $mailboxDetails = Populate-Details -tenantStatsHash $tenantStatsHash -entity $mailbox -IsMailbox
            $global:tenantStatsHash["InactiveMailboxDetails"][$mailbox.PrimarySMTPAddress] = $mailboxDetails
        }

        Write-ProgressHelper -Total $inactiveMailboxCount -Activity "Processing Complete" -Completed

        $completedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed Inactive Mailbox Report in $completedTime" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Report-InactiveMailboxes] COMPLETED: Processed Inactive Mailboxes in $completedTime" -ExportFileLocation $ExportDetails
    }

    # Combine user and mailbox stats
    Write-Log -Type INFO -Message "[Combine-UserAndMailboxStats] Combining User and Mailbox Details" -ExportFileLocation $ExportDetails
    try {
        $start = Get-Date
        Write-Host "Combining User and Mailbox Details..." -ForegroundColor Cyan -NoNewline
        Combine-UserAndMailboxStats -tenantStatsHash $global:tenantStatsHash
        Write-Host "Generating Inactive Mailbox Report..." -ForegroundColor Cyan -NoNewline
        Report-InactiveMailboxes -tenantStatsHash $global:tenantStatsHash
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
    $global:tenantStatsHash["AllGraphUserStats"] = @{}

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

    # Fetch Graph Data and store in $global:tenantStatsHash
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching Teams User Activity from Graph" -ExportFileLocation $ExportDetails 
    $global:tenantStatsHash["TeamsUserData"] = Get-GraphDataWithLogging -Uri $dataUris.TeamsUserReportsURI -DataName "Teams User Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching OneDrive User Activity from Graph" -ExportFileLocation $ExportDetails 
    $global:tenantStatsHash["OneDriveData"] = Get-GraphDataWithLogging -Uri $dataUris.OneDriveUsageUri -DataName "OneDrive Usage Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching Exchange User Activity from Graph" -ExportFileLocation $ExportDetails 
    $global:tenantStatsHash["EmailData"] = Get-GraphDataWithLogging -Uri $dataUris.EmailReportsUri -DataName "Exchange Activity Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching Exchange Mailbox Usage Report from Graph" -ExportFileLocation $ExportDetails 
    $global:tenantStatsHash["MailboxUsage"] = Get-GraphDataWithLogging -Uri $dataUris.MailboxUsageReportsUri -DataName "Mailbox Usage Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching SharePoint Online Usage Report from Graph" -ExportFileLocation $ExportDetails 
    $global:tenantStatsHash["SPOUsage"] = Get-GraphDataWithLogging -Uri $dataUris.SPOUsageReportsUri -DataName "SharePoint Activity Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching Yammer Usage Report from Graph" -ExportFileLocation $ExportDetails 
    $global:tenantStatsHash["YammerUsage"] = Get-GraphDataWithLogging -Uri $dataUris.YammerUsageReportsUri -DataName "Yammer Activity Report"
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Fetching User Sign In Report from Graph" -ExportFileLocation $ExportDetails 
    $global:tenantStatsHash["SignInData"] = Get-GraphDataWithLogging -Uri $dataUris.SignInUri -DataName "User Sign-In Data"

    # Create hash table for user sign-in data within $global:tenantStatsHash
    $global:tenantStatsHash["UserSignIns"] = @{}
    
    # Process User sign-in data and store in $global:tenantStatsHash["UserSignIns"]
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Processing User Sign-In Data fetched from Graph" -ExportFileLocation $ExportDetails
    [array]$UserSignInData = $global:tenantStatsHash["SignInData"] | Where-Object { $_.UserType -eq "Member" } | Sort-Object UserPrincipalName -Unique
    ForEach ($U in $UserSignInData) {
        If ($U.SignInActivity.LastSignInDateTime) {
            $LastSignInDate = Get-Date($U.SignInActivity.LastSignInDateTime) -format g
            $global:tenantStatsHash["UserSignIns"].Add([String]$U.UserPrincipalName, $LastSignInDate)
        } Else {
            $global:tenantStatsHash["UserSignIns"].Add([String]$U.UserPrincipalName, $Null)
        }
    }

    $StartTime2 = Get-Date
    Write-Host "Processing activity data fetched from the Graph..."
    Write-Log -Type INFO -Message "[Get-GraphUserStats] Processing activity data fetched from the Graph" -ExportFileLocation $ExportDetails

    # Initialize the user data hash table within $global:tenantStatsHash
    $DataTable = @{}

    # Process Teams Data
    ForEach ($T in $global:tenantStatsHash["TeamsUserData"]) {
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
    ForEach ($E in $global:tenantStatsHash["EmailData"]) {
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
    ForEach ($M in $global:tenantStatsHash["MailboxUsage"]) {
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
    ForEach ($S in $global:tenantStatsHash["SPOUsage"]) {
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
    ForEach ($O in $global:tenantStatsHash["OneDriveData"]) {
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
    ForEach ($Y in $global:tenantStatsHash["YammerUsage"]) {  
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

    # Process each user to extract Exchange, Teams, OneDrive, SharePoint, and Yammer statistics for their activity
    ForEach ($UserPrincipalName in $Users) {
        $U = $UserPrincipalName.UPN
        $UserNumber++
        Write-ProgressHelper -Activity "Extracting User Data" `
            -CurrentOperation "Processing user: $U" `
            -ProgressCounter $UserNumber `
            -TotalCount $TotalUsers `
            -StartTime $StartTime3

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
        $LastAccountSignIn = $global:tenantStatsHash["UserSignIns"].Item($U)
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
            $global:tenantStatsHash["AllGraphUserStats"][$U] = $OutLine
        } 
    }

    Write-ProgressHelper -Total 1 -Activity "Extracting information for user" -Completed

    $StartTime4 = Get-Date
    $GraphTime = $StartTime2 - $StartTime1
    $PrepTime = $StartTime3 - $StartTime2
    $ReportTime = $StartTime4 - $StartTime3
    $ScriptTime = $StartTime4 - $StartTime1
    $AccountsPerMinute = [math]::Round(($global:tenantStatsHash["AllGraphUserStats"].Values.count/($ScriptTime.TotalSeconds/60)),2)
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
    Write-Verbose "Total accounts processed:                $($global:tenantStatsHash["AllGraphUserStats"].Values.count)"
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
        # Ensure global hash table structure
        if (-not $global:tenantStatsHash) {
            $global:tenantStatsHash = @{}
        }
    $global:tenantStatsHash["DeviceDetails"] = @{}

    Write-Host "Getting Device Details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllDevicesReport] START: Gathering all Device with $($detailLevel) details" -ExportFileLocation $ExportDetails
    
    try {
        $devices = Get-MgDevice -All -ErrorAction Stop | ? {$null -ne $_.ID}
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
        foreach ($device in $devices) {
            Write-ProgressHelper -Total $totalCount -Activity "Processing all Found Devices" -Operation "Updating Device Details for $($device.DisplayName)"
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
                $global:tenantStatsHash["DeviceDetails"][$device.ObjectID] = $device
            }
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllDevicesReport] An error occurred in running Get-AllDeviceReport function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-ProgressHelper -Total $totalCount -Activity "Processing all Entra Found Devices" -Completed
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

    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["ConditionalAccessPolicies"] = @{}

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
        $conditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop

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

        foreach ($policy in $conditionalAccessPolicies) {
            Write-ProgressHelper -Total $totalCount -Activity "Processing all Conditional Access Policies" -Operation "Expanding Policy Details for $($policy.DisplayName)"
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

            $global:tenantStatsHash["ConditionalAccessPolicies"][$policy.DisplayName] = $policyDetailsHash
        }
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-ConditionalAccessPoliciesReport] An error occurred in running Get-ConditionalAccessPoliciesReport function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-ProgressHelper -Total $totalCount -Activity "Processing all Conditional Access Policies" -Completed
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
    # Ensure global hash table structure
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["SecuritySecureScore"] = @{}
    $global:tenantStatsHash["SecureScoreActions"] = @{}

    Write-Host "Getting Security Score Details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] START: Gathering all Security Score details" -ExportFileLocation $ExportDetails
    try {
        $controlProfileLookup = @{}
        try {
            $controlProfiles = @(Get-MgSecuritySecureScoreControlProfile -All -ErrorAction Stop)
            foreach ($profile in $controlProfiles) {
                if ($profile.Id) {
                    $controlProfileLookup[$profile.Id] = $profile
                }
            }
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Loaded $($controlProfileLookup.Count) Secure Score control profile mappings" -ExportFileLocation $ExportDetails
        } catch {
            Write-Log -Type WARNING -Message "[Get-SecuritySecureScoreReport] Unable to load Secure Score control profiles for source mapping: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        }

        if ($MostRecent) {
            $secureScore = Get-MgSecuritySecureScore -Top 1 -ErrorAction Stop | Where-Object {$null -ne $_.ID}
        } else {
            $secureScore = Get-MgSecuritySecureScore -All -ErrorAction Stop | Where-Object {$null -ne $_.ID}
        }

        # Add Additional Properties
        $totalCount = $secureScore.count
        foreach ($score in $secureScore) {
            Write-ProgressHelper -Total $totalCount -Activity "Processing all Security Score Details" -Operation "Gathering Score Details for $($score.ID)"
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
                    foreach ($individualScore in $ComparisonScore.AdditionalProperties.keys) {
                        Write-Log -Type DEBUG -Message "[Get-SecuritySecureScoreReport] Gathering $($ComparisonScore.Basis):$($individualScore) Score Details for $($score.ID): Add Comparitive Scores" -ExportFileLocation $ExportDetails
                        $currentSecurityScores | Add-Member -MemberType NoteProperty -Name "$($ComparisonBasisName)_$($individualScore)" -Value $ComparisonScore.AdditionalProperties[$individualScore]
                    }
                }
            }

            
            #Add to Hash Table
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Gathering Score Details for $($score.ID): Add Score to Tenant Stats Hash Table" -ExportFileLocation $ExportDetails
            $global:tenantStatsHash["SecuritySecureScore"][$score.ID] = $currentSecurityScores
        }

        $latestScore = @($secureScore | Sort-Object CreatedDateTime -Descending | Select-Object -First 1)
        if ($latestScore.Count -gt 0) {
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

                $global:tenantStatsHash["SecureScoreActions"][("{0:D3}-{1}" -f $actionIndex, $controlName)] = $actionRow
            }
            Write-Log -Type INFO -Message "[Get-SecuritySecureScoreReport] Added $($global:tenantStatsHash['SecureScoreActions'].Count) Secure Score recommendation mappings" -ExportFileLocation $ExportDetails
        }

    }
    catch {
        Write-Log -Type ERROR -Message "[Get-SecuritySecureScoreReport] An error occurred in running Get-SecuritySecureScoreReport function. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-ProgressHelper -Total $totalCount -Activity "Processing all Security Score Details" -Completed
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

    # Function to get group details by ID or DisplayName
    function Get-EntraGroupDetails {
        param (
            [Parameter(Mandatory = $true)]
            [string]$GroupIdentifier,  # Can be Group ID or DisplayName
            [Parameter(Mandatory=$false,HelpMessage='Provide the Graph Authentication Type')]
            [ValidateSet('SDK','REST')]
            [string[]]$GraphAuthType
        )
        switch ($GraphAuthType) {
            REST { 
                # Construct the API URL to fetch group details by ID or DisplayName
                if ($GroupIdentifier -match "^[0-9a-fA-F-]{36}$") {
                    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupIdentifier"
                } else {
                    $uri = "https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '$GroupIdentifier'"
                }

                # Fetch the group details
                try {
                    Write-Log -Type INFO -Message "Fetching group details for $GroupIdentifier" -ExportFileLocation $ExportDetails
                    $GroupDetails = Get-GraphData -PageSize 999 -URI $uri -ID 3 -Activity "Gathering Group Details"
                    #$GroupDetails2 = Invoke-RestMethod -Uri $uri -Headers $global:GraphHeaders -Method Get -ContentType "application/json"
                    Write-Log -Type INFO -Message "Group details fetched successfully for $($GroupDetails.displayName)" -ExportFileLocation $ExportDetails
                }
                catch {
                    Write-Log -Type ERROR -Message "Error fetching group details for $($GroupDetails.displayName): $($_)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                    return
                }

                # Check for group-based licensing
                $isManagingLicenses = $false
                try {
                    Write-Log -Type INFO -Message "Checking license details for group $($GroupDetails.displayName)" -ExportFileLocation $ExportDetails
                    $licenseUri = "https://graph.microsoft.com/v1.0/groups/$($GroupDetails.id)?`$select=assignedLicenses"
                    # Fetch the license details for the group
                    $licenseDetails = Get-GraphData -PageSize 999 -ID 3 -URI $licenseUri -Activity "Gathering License Details"
                    #$licenseDetails = Invoke-RestMethod -Uri $licenseUri -Headers $global:GraphHeaders -Method Get -ContentType "application/json"

                    if ($licenseDetails.assignedLicenses.count -gt 0) {
                        Write-Log -Type INFO -Message "Group $($GroupDetails.displayName) is managing licenses" -ExportFileLocation $ExportDetails
                        $isManagingLicenses = $true
                    }
                } catch {
                    Write-Log -Type ERROR -Message "Error checking license details for group $($GroupDetails.displayName): $($_)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                }

                <# Not Working - Other Group Details - REQUIRES WRITE PERMISSIONS
                https://learn.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0#properties
                $otherDetailsURI = "https://graph.microsoft.com/v1.0/groups/$($GroupDetails.id)?`$select=hideFromAddressLists"
                $otherDetails = Invoke-RestMethod -Uri $otherDetailsURI -Headers $global:GraphHeaders -Method Get -ContentType "application/json"
                #>
                
                # If search was by DisplayName, get the first result
                if ($GroupIdentifier -notmatch "^[0-9a-fA-F-]{36}$") {
                    $GroupDetails = $GroupDetails | Select-Object -First 1
                    #$GroupDetails = $GroupDetails.value | Select-Object -First 1
                }
                # Determine group type with added check for Dynamic Distribution Groups
                if ($GroupDetails.groupTypes -contains "DynamicMembership") {
                    $MembershipType = "Dynamic Group"
                    $MembershipRule = $GroupDetails.membershipRule
                } else {
                    $MembershipType = "Assigned"
                    $MembershipRule = $null
                }
                Write-Log -Type INFO -Message "Group Type: $MembershipType" -ExportFileLocation $ExportDetails
                # Determine group type based on mailEnabled and securityEnabled properties
                if ($GroupDetails.groupTypes -contains "Unified") {
                    $GroupType = "Microsoft 365"

                } elseif ($GroupDetails.mailEnabled -eq $true -and $GroupDetails.securityEnabled -eq $true) {
                    $GroupType = "Mail-enabled Security Group"

                } elseif ($GroupDetails.mailEnabled -eq $false -and $GroupDetails.securityEnabled -eq $true) {
                    $GroupType = "Security Group"

                } elseif ($GroupDetails.mailEnabled -eq $true -and $GroupDetails.securityEnabled -eq $false) {
                    $GroupType = "Distribution Group"

                } else {
                    $GroupType = "Unknown"
                }
                Write-Log -Type INFO -Message "Group Type: $GroupType" -ExportFileLocation $ExportDetails

                # Count members and owners using Graph $count endpoints
                $MemberCount = 0
                $OwnerCount = 0
                $hasNestedMembers = "Skipped"
                $isDynamicDistributionGroup = ($GroupDetails.groupTypes -contains "DynamicMembership") -and ($GroupDetails.mailEnabled -eq $true) -and ($GroupDetails.securityEnabled -eq $false) -and (-not ($GroupDetails.groupTypes -contains "Unified"))
                if ($isDynamicDistributionGroup) {
                    $MemberCount = "Skipped (Dynamic Distribution Group)"
                    $OwnerCount = "Skipped (Dynamic Distribution Group)"
                    Write-Log -Type INFO -Message "Skipping member/owner counts for dynamic distribution group $($GroupDetails.displayName)" -ExportFileLocation $ExportDetails
                } else {
                    try {
                        Write-Log -Type INFO -Message "Checking member and owner count for group $($GroupDetails.displayName)" -ExportFileLocation $ExportDetails
                        $memberUri = "https://graph.microsoft.com/v1.0/groups/$($GroupDetails.id)/members/\$count"
                        $memberCountResult = Get-GraphData -URI $memberUri -ID 3 -Activity "Counting Members"
                        $MemberCount = if ($memberCountResult -is [array]) { [int]($memberCountResult | Select-Object -First 1) } else { [int]$memberCountResult }

                        $ownerUri = "https://graph.microsoft.com/v1.0/groups/$($GroupDetails.id)/owners/\$count"
                        $ownerCountResult = Get-GraphData -URI $ownerUri -ID 3 -Activity "Counting Owners"
                        $OwnerCount = if ($ownerCountResult -is [array]) { [int]($ownerCountResult | Select-Object -First 1) } else { [int]$ownerCountResult }
                    } catch {
                        Write-Log -Type ERROR -Message "Error counting members/owners for group $($GroupDetails.displayName): $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                    }
                }
                Write-Log -Type INFO -Message "Member Count: $MemberCount" -ExportFileLocation $ExportDetails
                Write-Log -Type INFO -Message "Owner Count: $OwnerCount" -ExportFileLocation $ExportDetails

                # Source of the group
                $Source = if ($GroupDetails.onPremisesSyncEnabled) { "On-Premises" } else { "Cloud" }
                Write-Log -Type INFO -Message "Group Source: $Source" -ExportFileLocation $ExportDetails
             }
            SDK { # TBD
                $GroupDetails = Get-MgGroup -GroupId $GroupIdentifier -ErrorAction Stop
                if ($GroupDetails.groupTypes -contains "DynamicMembership") {
                    $MembershipType = "Dynamic Group"
                    $MembershipRule = $GroupDetails.membershipRule
                } else {
                    $MembershipType = "Assigned"
                    $MembershipRule = $null
                }
                if ($GroupDetails.groupTypes -contains "Unified") {
                    $GroupType = "Microsoft 365"
                } elseif ($GroupDetails.mailEnabled -eq $true -and $GroupDetails.securityEnabled -eq $true) {
                    $GroupType = "Mail-enabled Security Group"
                } elseif ($GroupDetails.mailEnabled -eq $false -and $GroupDetails.securityEnabled -eq $true) {
                    $GroupType = "Security Group"
                } elseif ($GroupDetails.mailEnabled -eq $true -and $GroupDetails.securityEnabled -eq $false) {
                    $GroupType = "Distribution Group"
                } else {
                    $GroupType = "Unknown"
                }
                $Source = if ($GroupDetails.onPremisesSyncEnabled) { "On-Premises" } else { "Cloud" }
                $isManagingLicenses = $false
                $MemberCount = 0
                $OwnerCount = 0
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
    $Global:tenantStatsHash["EntraIDGroups"] = @{}

    switch ($GraphAuthType) {
        SDK { 
            $Groups = Get-MgGroup -All -ErrorAction Stop
            $totalGroups = $Groups.Count
            }
        REST { 
            $Groups = @()

            # Endpoint and initial URL for group data
            $GroupsEndpoint = "https://graph.microsoft.com/v1.0/groups"
        
            # Loop to handle paging through all group data
            try {
                Write-Log -Type INFO -Message "Fetching initial Entra Groups" -ExportFileLocation $ExportDetails
                $Groups = Get-GraphData -PageSize 999 -URI $GroupsEndpoint -ID 1 -Activity "Gathering Group Details"   
                Write-Progress -Activity "Getting Group Details" -Completed -Id 1
            }
            catch {
                Write-Log -Type ERROR -Message "Error fetching initial Entra Groups: $($_)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
                return @()
            }
        }
    }

    #Process the groups data
    $totalGroups = $Groups.Count
    foreach ($Group in $Groups) {
        Write-ProgressHelper -Total $totalGroups -Id 2 -Activity "Getting Group Details" -Operation "Processing Group: $($Group.displayName)"

        # Fetch detailed group information using Get-EntraGroupDetails
        #Write-Host "Checking group $($Group.displayName)"
        Write-Log -Type INFO -Message "Checking group $($Group.displayName)" -ExportFileLocation $ExportDetails
        $GroupDetails = Get-EntraGroupDetails -GroupIdentifier $Group.id -GraphAuthType $GraphAuthType
        
        if ($GroupDetails) {
            # Store the detailed group information in the hash table
            $global:tenantStatsHash['EntraIDGroups'][$GroupDetails.ID] = $GroupDetails
        } else { Write-Warning "Could not retrieve details for group $($Group.displayName)" }
    }

    Write-ProgressHelper -Total $totalGroups -Id 2 -Activity "Getting Group Details" -Completed
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
    if (-not $global:tenantStatsHash) {
        $global:tenantStatsHash = @{}
    }
    $global:tenantStatsHash["AuthenticationConfig"] = @{}
    $global:tenantStatsHash["AuthenticationConfigSummary"] = @{}
    $global:tenantStatsHash["AuthenticationMethods"] = @{}
    $global:tenantStatsHash["AuthenticationSSOApplications"] = @{}
    
    Write-Host "Checking Authentication and SSO Configuration ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] START: Checking Authentication Configuration" -ExportFileLocation $ExportDetails
    
    try {
        # Get authentication methods policy
        $authMethodsPolicy = [PSCustomObject]@{
            MFAEnabled = $false
            MFAMethods = @()
            SSOEnabled = $false
            SSOApplications = @()
            FederatedDomains = @()
            PasswordlessMethods = @()
        }
        
        # Check for federated domains (indicates SSO)
        $domains = $global:tenantStatsHash["Domains"].Values
        $federatedDomains = $domains | Where-Object {$_.AuthenticationType -eq "Federated"}
        
        if ($federatedDomains) {
            $authMethodsPolicy.SSOEnabled = $true
            $authMethodsPolicy.FederatedDomains = $federatedDomains.Domain
            Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Found $($federatedDomains.Count) federated domains" -ExportFileLocation $ExportDetails
        }
        
        # Get Enterprise Applications configured for SSO
        try {
            Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Checking Enterprise Applications for SSO" -ExportFileLocation $ExportDetails
            
            $ssoApps = Get-MgServicePrincipal -All -Filter "tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')" -ErrorAction SilentlyContinue | 
                Where-Object {$null -ne $_.PreferredSingleSignOnMode -and $_.PreferredSingleSignOnMode -ne "notSupported"}
            
            if ($ssoApps) {
                $authMethodsPolicy.SSOEnabled = $true
                $ssoAppDetails = @()
                
                foreach ($app in $ssoApps) {
                    $appDetail = [PSCustomObject]@{
                        DisplayName = $app.DisplayName
                        AppId = $app.AppId
                        SSOMode = $app.PreferredSingleSignOnMode
                        ServicePrincipalType = $app.ServicePrincipalType
                        AccountEnabled = $app.AccountEnabled
                    }
                    $ssoAppDetails += $appDetail
                }
                
                $authMethodsPolicy.SSOApplications = $ssoAppDetails
                Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Found $($ssoApps.Count) SSO-enabled applications" -ExportFileLocation $ExportDetails
            }
            
        } catch {
            Write-Log -Type WARNING -Message "[Get-AuthenticationConfiguration] Unable to retrieve SSO applications: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
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
        if ($global:tenantStatsHash["ConditionalAccessPolicies"]) {
            $mfaPolicies = $global:tenantStatsHash["ConditionalAccessPolicies"].Values | 
                Where-Object {$_.GrantControls -like "*mfa*"}
            
            if ($mfaPolicies) {
                $authMethodsPolicy.MFAEnabled = $true
                $authMethodsPolicy | Add-Member -MemberType NoteProperty -Name "MFAConditionalAccessPolicies" -Value $mfaPolicies.Count
                Write-Log -Type INFO -Message "[Get-AuthenticationConfiguration] Found $($mfaPolicies.Count) CA policies requiring MFA" -ExportFileLocation $ExportDetails
            }
        }
        
        $global:tenantStatsHash["AuthenticationConfig"]["Configuration"] = $authMethodsPolicy

        $ssoAppNames = @($authMethodsPolicy.SSOApplications | ForEach-Object { $_.DisplayName } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $global:tenantStatsHash["AuthenticationConfigSummary"]["Summary"] = [PSCustomObject]@{
            MFAEnabled                   = $authMethodsPolicy.MFAEnabled
            MFAMethods                   = ($authMethodsPolicy.MFAMethods -join ', ')
            SSOEnabled                   = $authMethodsPolicy.SSOEnabled
            SSOApplicationsCount         = @($authMethodsPolicy.SSOApplications).Count
            SSOApplications              = ($ssoAppNames -join ', ')
            FederatedDomains             = ($authMethodsPolicy.FederatedDomains -join ', ')
            PasswordlessMethods          = ($authMethodsPolicy.PasswordlessMethods -join ', ')
            MFAConditionalAccessPolicies = $(if ($authMethodsPolicy.PSObject.Properties['MFAConditionalAccessPolicies']) { $authMethodsPolicy.MFAConditionalAccessPolicies } else { 0 })
        }

        $methodIndex = 0
        foreach ($method in @($authMethodsPolicy.MFAMethods)) {
            $methodIndex++
            $global:tenantStatsHash["AuthenticationMethods"][("{0:D3}-MFA" -f $methodIndex)] = [PSCustomObject]@{
                Category = 'MFA Method'
                Value    = $method
            }
        }

        foreach ($method in @($authMethodsPolicy.PasswordlessMethods)) {
            $methodIndex++
            $global:tenantStatsHash["AuthenticationMethods"][("{0:D3}-Passwordless" -f $methodIndex)] = [PSCustomObject]@{
                Category = 'Passwordless Method'
                Value    = $method
            }
        }

        foreach ($domainName in @($authMethodsPolicy.FederatedDomains)) {
            $methodIndex++
            $global:tenantStatsHash["AuthenticationMethods"][("{0:D3}-Federated" -f $methodIndex)] = [PSCustomObject]@{
                Category = 'Federated Domain'
                Value    = $domainName
            }
        }

        $ssoIndex = 0
        foreach ($app in @($authMethodsPolicy.SSOApplications)) {
            $ssoIndex++
            $global:tenantStatsHash["AuthenticationSSOApplications"][("{0:D3}-{1}" -f $ssoIndex, $app.DisplayName)] = $app
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
    if (-not $global:tenantStatsHash) { $global:tenantStatsHash = @{} }
    $global:tenantStatsHash["TenantInfo"] = @{}

    Write-Host "Gathering Tenant Overview Info ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-TenantOverviewInfo] START: Gathering tenant overview information" -ExportFileLocation $ExportDetails

    try {
        $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1

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

        $global:tenantStatsHash["TenantInfo"] = [PSCustomObject]@{
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
    if (-not $global:tenantStatsHash) { $global:tenantStatsHash = @{} }
    $global:tenantStatsHash["AdConnectConfiguration"] = @{}
    
    Write-Host "Checking AD Connect/Sync status ..." -ForegroundColor Cyan -NoNewline
    Write-Log -Type INFO -Message "[Get-AdConnectSyncDetails] START" -ExportFileLocation $ExportDetails
    
    try {
        $org = $null
        try { $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1 } catch {}
        
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
        
        $global:tenantStatsHash["AdConnectConfiguration"]["Summary"] = $summary
        $global:tenantStatsHash["AdConnectConfiguration"]["SyncServices"] = $serviceDetails
        $global:tenantStatsHash["AdConnectConfiguration"]["RecentErrors"] = $syncErrors | Select-Object -First 10
        $global:tenantStatsHash["AdConnectConfiguration"]["ErrorCount"] = ($syncErrors | Measure-Object).Count
        
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
    if (-not $global:tenantStatsHash) { $global:tenantStatsHash = @{} }
    $global:tenantStatsHash["MfaRegistrationDetails"] = @{}
    $global:tenantStatsHash["MfaRegistrationSummary"] = $null
    
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
                $global:tenantStatsHash["MfaRegistrationDetails"][$upn] = [pscustomobject]@{
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
            $global:tenantStatsHash["MfaRegistrationSummary"] = $summary
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
    if (-not $global:tenantStatsHash) { $global:tenantStatsHash = @{} }
    $global:tenantStatsHash["FederationConfiguration"] = @{}

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
                    $lookup = Invoke-RestMethod -Uri $tenantLookupUri -Headers $global:GraphHeaders -Method GET -ContentType "application/json" -ErrorAction Stop
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
                        $response = Invoke-WebRequest -Uri $uri -UseBasicParsing -ErrorAction Stop
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
                    $lookup = Invoke-RestMethod -Uri $tenantLookupUri -Headers $global:GraphHeaders -Method GET -ContentType "application/json" -ErrorAction Stop
                    if ($lookup.displayName) { return $lookup.displayName }
                }
            } catch {
                Write-Log -Type WARNING -Message "[Get-FederationAndCrossTenantConfiguration] Tenant name lookup (beta) failed for $($TenantId): $($_.Exception.Message)" -ExportFileLocation $ExportDetails
            }
            
            # 5) Last resort - this often fails cross-tenant
            try {
                $orgInfo = Get-MgOrganization -OrganizationId $TenantId -ErrorAction Stop
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

        $global:tenantStatsHash["FederationConfiguration"]["ExchangeFederation"] = $exchangeFed
        $global:tenantStatsHash["FederationConfiguration"]["CrossTenantAccess"] = $crossTenantSummary
        $global:tenantStatsHash["FederationConfiguration"]["ExternalIdentities"] = $externalIdentities

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

    function ConvertTo-JsonFriendlyValue {
        param(
            [Parameter(Mandatory = $false)]
            $Value,
            [Parameter(Mandatory = $false)]
            [int]$Depth = 0,
            [Parameter(Mandatory = $false)]
            [System.Collections.Generic.HashSet[int]]$Visited
        )

        if ($null -eq $Value) {
            return $null
        }

        if ($null -eq $Visited) {
            $Visited = [System.Collections.Generic.HashSet[int]]::new()
        }

        if ($Depth -ge 20) {
            return '[MaxDepthExceeded]'
        }

        $valueType = $Value.GetType()
        $isReferenceType = -not $valueType.IsValueType -and $Value -isnot [string]
        $referenceId = $null
        if ($isReferenceType) {
            $referenceId = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Value)
            if (-not $Visited.Add($referenceId)) {
                return '[CircularReference]'
            }
        }

        try {
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

            if ($Value -is [datetime]) {
                return $Value.ToString('o')
            }

            if ($Value -is [datetimeoffset]) {
                return $Value.ToString('o')
            }

            if ($Value -is [timespan] -or $Value -is [guid] -or $Value -is [uri] -or $Value -is [version]) {
                return $Value.ToString()
            }

            if ($Value -is [enum]) {
                return $Value.ToString()
            }

            if ($Value -is [securestring]) {
                return '[SecureString]'
            }

            if ($Value -is [System.Management.Automation.SwitchParameter]) {
                return [bool]$Value
            }

            if ($Value -is [System.Collections.IDictionary]) {
                $result = [ordered]@{}
                foreach ($key in $Value.Keys) {
                    $result[[string]$key] = ConvertTo-JsonFriendlyValue -Value $Value[$key] -Depth ($Depth + 1) -Visited $Visited
                }
                return $result
            }

            if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
                $items = New-Object System.Collections.Generic.List[object]
                foreach ($item in $Value) {
                    $items.Add((ConvertTo-JsonFriendlyValue -Value $item -Depth ($Depth + 1) -Visited $Visited))
                }
                return $items.ToArray()
            }

            $serializableProperties = @(
                $Value.PSObject.Properties |
                    Where-Object {
                        $_.MemberType -in @('NoteProperty', 'AliasProperty') -and
                        $_.Name -ne 'SyncRoot'
                    }
            )

            if ($serializableProperties.Count -gt 0) {
                $result = [ordered]@{}
                foreach ($property in $serializableProperties) {
                    try {
                        $result[$property.Name] = ConvertTo-JsonFriendlyValue -Value $property.Value -Depth ($Depth + 1) -Visited $Visited
                    } catch {
                        $result[$property.Name] = "[PropertyReadError] $($_.Exception.Message)"
                    }
                }
                return $result
            }

            return $Value.ToString()
        } finally {
            if ($isReferenceType -and $null -ne $referenceId) {
                $Visited.Remove($referenceId) | Out-Null
            }
        }
    }

    $visited = [System.Collections.Generic.HashSet[int]]::new()
    $payload = [ordered]@{
        SchemaVersion = 1
        GeneratedAt   = (Get-Date).ToString("o")
        Data          = ConvertTo-JsonFriendlyValue -Value $TenantStatsHash -Visited $visited
    }

    $jsonOptions = [System.Text.Json.JsonSerializerOptions]::new()
    $jsonOptions.WriteIndented = $true
    $jsonOptions.ReferenceHandler = [System.Text.Json.Serialization.ReferenceHandler]::IgnoreCycles
    $json = [System.Text.Json.JsonSerializer]::Serialize($payload, $jsonOptions)
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.UTF8Encoding]::new($false))
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

    $json = Get-Content -Raw -Path $Path | ConvertFrom-Json
    return $json.Data
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
if ($Thresholds) {
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
    
    if ($null -eq $InputObject) {
        return @()
    }
    
    # If it's a hashtable, extract the values as an array
    if ($InputObject -is [hashtable] -or $InputObject -is [System.Collections.Specialized.OrderedDictionary]) {
        Write-Verbose "Converting hashtable with $($InputObject.Count) items to array"
        return @($InputObject.Values)
    }
    
    # If it's already enumerable (array, list, etc), ensure it's an array
    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        return @($InputObject)
    }
    
    # Single object - wrap in array
    return @($InputObject)
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
    
    if ($null -eq $Number -or $Number -eq '') { return 'N/A' }
    
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
    
    # If over 1000 GB (1 TB), show in TB
    if ($SizeInGB -ge 1000) {
        $sizeInTB = [math]::Round($SizeInGB / 1024, 2)
        return "$sizeInTB TB"
    }
    else {
        return "$([math]::Round($SizeInGB, 2)) GB"
    }
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
    
    if ($null -eq $Value -or $Value -eq '') { return 'N/A' }
    
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
            if ($null -eq $value -or $value -eq '') {
                $displayValue = 'N/A'
            } elseif ($value -is [datetime]) {
                $displayValue = $value.ToString('yyyy-MM-dd')
            } elseif ($value -is [bool]) {
                $displayValue = if ($value) { '✓' } else { '✗' }
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
    param([array]$Licenses)
    
    $findings = @()
    $criticalFindings = @()
    $warningFindings = @()
    $infoFindings = @()
    
    # Process licenses with forced recalculation
    $processedLicenses = $Licenses | ForEach-Object {
        $lic = $_
        
        $purchased = 0
        $consumed = 0
        
        if ($null -ne $lic.PurchasedUnits -and $lic.PurchasedUnits -ne 'N/A' -and $lic.PurchasedUnits -ne '') {
            try { $purchased = [int]$lic.PurchasedUnits } catch { $purchased = 0 }
        }
        
        if ($null -ne $lic.ConsumedUnits -and $lic.ConsumedUnits -ne 'N/A' -and $lic.ConsumedUnits -ne '') {
            try { $consumed = [int]$lic.ConsumedUnits } catch { $consumed = 0 }
        }
        
        $remaining = $purchased - $consumed
        $utilization = if ($purchased -gt 0) { ($consumed / $purchased) * 100 } else { 0 }
        $friendlyName = $null
        if ($lic.PSObject.Properties['SkuFriendlyName'] -and $lic.SkuFriendlyName) {
            $friendlyName = $lic.SkuFriendlyName
        } else {
            $friendlyName = Get-FriendlyProductName -SkuPartNumber $lic.SkuPartNumber
        }
        $isTrialFlag = $false
        if ($lic.PSObject.Properties['IsTrial'] -and $lic.IsTrial -eq $true) {
            $isTrialFlag = $true
        }
        $isFreeOrTrialFlag = $false
        if ($lic.PSObject.Properties['IsFreeOrTrial'] -and $lic.IsFreeOrTrial -eq $true) {
            $isFreeOrTrialFlag = $true
        }
        $ignoreLifecycle = $false
        if ($lic.PSObject.Properties['IgnoreLifecycle'] -and $lic.IgnoreLifecycle -eq $true) {
            $ignoreLifecycle = $true
        }
        $isPaid = (-not $isTrialFlag) -and (-not $isFreeOrTrialFlag) -and (-not $ignoreLifecycle) -and ($lic.SkuPartNumber -notmatch 'TRIAL|EXPLORATORY|Windows_Store|FREE|VIRAL|_FACULTY|_STUDENT')
        
        [PSCustomObject]@{
            SkuPartNumber = $lic.SkuPartNumber
            SkuFriendlyName = $friendlyName
            PurchasedUnits = $purchased
            ConsumedUnits = $consumed
            RemainingUnits = $remaining
            Utilization = $utilization
            IsPaid = $isPaid
        }
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
            $findings += @{
                Type = 'Warning'
                Category = 'Large Archives'
                Message = "Archive for '$($mbx.DisplayName)' is $([math]::Round($archiveSize,1)) GB (threshold: $($script:DefaultThresholds.ArchiveSizeGB) GB)"
                Anchor = 'mailboxes'
                Priority = 2
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
    
    return @{
        Findings = $findings
        LargeMailboxes = $largeMailboxes
    }
}

function Get-DomainAnalysis {
    <#
    .SYNOPSIS
        Analyzes domain configuration and returns findings
    #>
    param([array]$Domains)
    
    $findings = @()
    
    foreach ($domain in $Domains) {
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
        [array]$Groups
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

    return [PSCustomObject]@{
        Licenses               = Get-ContextArray -Key 'LicenseSKUs'
        Recipients             = Get-ContextArray -Key 'AllRecipients'
        Mailboxes              = Get-ContextArray -Key 'MailboxFullDetails'
        InactiveMailboxes      = Get-ContextArray -Key 'InactiveMailboxDetails'
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
        'domains-dns' { return 'Review MX, autodiscover, and mail-routing records to plan coexistence and cutover.' }
        'identity-admins' { return 'Validate admin access, guest usage, and group ownership before identity migration activities.' }
        'mailboxes' { return 'Identify oversized or specialized mailboxes early to plan batching, archives, and exception handling.' }
        'inactive-mailboxes' { return 'Decide whether inactive mailboxes need retention, restore, or exclusion from scope.' }
        'sharepoint-onedrive' { return 'Use site inventory, ownership, and storage metrics to prioritize high-risk collaboration workloads.' }
        'devices' { return 'Review stale and non-compliant devices before identity and endpoint cutover.' }
        'ad-connect' { return 'Document synchronization dependencies and plan cloud identity cutover or staged decommissioning.' }
        'conditional-access-mfa' { return 'Review CA and MFA design to avoid post-migration lockouts or authentication regressions.' }
        'exchange-hybrid' { return 'Validate hybrid, connectors, and migration endpoints because they affect tenant-to-tenant messaging strategy.' }
        'cross-tenant-access' { return 'Review cross-tenant and B2B settings for coexistence, external collaboration, and post-migration cleanup.' }
        'secure-score' { return 'Use the mapped Microsoft Secure Score action to prioritize remediation with the highest security impact.' }
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

        $topFinding = @($AreaFindings | Sort-Object Priority | Select-Object -First 1)
        $primaryFindingText = if ($topFinding.Count -gt 0) { $topFinding[0].Message } else { 'No automated findings detected for this assessment area.' }
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
        $licAnalysis = Get-LicenseAnalysis -Licenses $context.Licenses
        Add-AreaSummary -Area 'Licensing' -AreaFindings $licAnalysis.Findings -AssessmentType 'Assessment heuristic using Microsoft 365 license data' -RelatedWorksheet 'LicenseSKUs' -Notes 'Evaluates capacity, at-capacity SKUs, and high utilization.'
    }

    if ($context.Domains.Count -gt 0) {
        $domainAnalysis = Get-DomainAnalysis -Domains $context.Domains
        Add-AreaSummary -Area 'Domains' -AreaFindings $domainAnalysis.Findings -AssessmentType 'Assessment heuristic using Microsoft 365 domain state' -RelatedWorksheet 'Domains' -Notes 'Highlights verification and mail-routing concerns relevant to migration cutover.'
    }

    if ($context.Users.Count -gt 0 -or $context.Admins.Count -gt 0 -or $context.Groups.Count -gt 0) {
        $identityAnalysis = Get-IdentityAdminAnalysis -Users $context.Users -Admins $context.Admins -Groups $context.Groups
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
    $paidLicenseAnalysis = if ($context.Licenses.Count -gt 0) { Get-LicenseAnalysis -Licenses $context.Licenses } else { $null }
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
        padding: 0 20px;
    }
    
    .nav {
        display: flex;
        gap: 5px;
        overflow-x: auto;
        padding: 10px 0;
    }
    
    .nav a {
        padding: 8px 16px;
        text-decoration: none;
        color: var(--text-primary);
        border-radius: 4px;
        white-space: nowrap;
        transition: all 0.2s;
        font-size: 0.9em;
    }
    
    .nav a:hover {
        background: var(--bg-light);
        color: var(--primary-color);
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
        font-size: 2em;
        font-weight: 600;
        color: var(--text-primary);
        margin-bottom: 5px;
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
    }
    
    .section h2 {
        font-size: 1.5em;
        color: var(--text-primary);
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 2px solid var(--border-color);
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
    
    // Smooth scroll for anchor links
    document.addEventListener('DOMContentLoaded', function() {
        document.querySelectorAll('a[href^="#"]').forEach(anchor => {
            anchor.addEventListener('click', function(e) {
                e.preventDefault();
                const target = document.querySelector(this.getAttribute('href'));
                if (target) {
                    target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            });
        });
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

    $html = @"
<canvas id="myPieChart$ChartTitle" style="width:${Width}px; height:${Height}px;"></canvas>
<script>
    var ctx = document.getElementById('myPieChart$ChartTitle').getContext('2d');
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
            responsive: false,
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
    param([array]$Licenses)
    
    if ($Licenses.Count -eq 0) {
        return "<div class='empty-state'>No license data available</div>"
    }
    
    #Write-Host "  Building License section..." -ForegroundColor Gray
    
    # Get analysis with processed licenses
    $analysis = Get-LicenseAnalysis -Licenses $Licenses
    $paidLicenses = $analysis.PaidLicenses
    
    # Calculate totals
    $totalPurchased = $analysis.TotalPurchased
    $totalConsumed = $analysis.TotalConsumed
    $totalRemaining = $totalPurchased - $totalConsumed
    
    # Build summary KPIs
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Total Licenses" -Value (Format-Number $totalPurchased) -Subtitle "Paid licenses only" -Theme 'default'
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
        $licenseName = if ($_.PSObject.Properties['SkuFriendlyName'] -and $_.SkuFriendlyName) { $_.SkuFriendlyName } else { $_.SkuPartNumber }
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
        $trialFlag = $false
        if ($_.PSObject.Properties['IsTrial'] -and $_.IsTrial -eq $true) {
            $trialFlag = $true
        }
        $freeOrTrialFlag = $false
        if ($_.PSObject.Properties['IsFreeOrTrial'] -and $_.IsFreeOrTrial -eq $true) {
            $freeOrTrialFlag = $true
        }
        $ignoreLifecycle = $false
        if ($_.PSObject.Properties['IgnoreLifecycle'] -and $_.IgnoreLifecycle -eq $true) {
            $ignoreLifecycle = $true
        }
        $isPaid = (-not $trialFlag) -and (-not $freeOrTrialFlag) -and (-not $ignoreLifecycle) -and ($_.SkuPartNumber -notmatch 'TRIAL|EXPLORATORY|Windows_Store|FREE|VIRAL|_FACULTY|_STUDENT')
        [PSCustomObject]@{
            SKU = $_.SkuPartNumber
            IsPaid = $isPaid
        }
    }
    $trialViralCount = ($allProcessed | Where-Object { -not $_.IsPaid }).Count
    
    if ($trialViralCount -gt 0) {
        $tableHtml += @"
<div style='margin-top: 15px; padding: 10px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;'>
    <strong>ℹ️ Note:</strong> $trialViralCount trial, viral, or free licenses exist but are <strong>excluded from this table and calculations</strong>.
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
        'TotalStorageGB' = 'Total (GB)'
        'AverageStorageGB' = 'Avg (GB)'
        'Over1TB' = 'Over 1 TB'
        'LargestSizeGB' = 'Largest (GB)'
    }
    
    $riskColumns = @{
        'Over1TB' = { param($val) [int]$val -gt 0 }
    }
    
    $tableHtml = New-HtmlTable -Data $summaries -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns
    
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

function Build-TeamsSection {
    param(
        [array]$Teams,
        [array]$Licenses,
        [object]$TeamsVoice
    )
    
    if ($Teams.Count -eq 0) {
        return "<div class='empty-state'>No Teams data available</div>"
    }
    
    # Determine Teams Voice availability from service plans
    $voicePlans = @('MCOEV','MCOPSTN1','MCOPSTN2','MCOEV_VIRTUALUSER','MCOEV_DOD','MCOPSTNC')
    $hasVoice = $false
    $voicePlanHits = @()
    $paidLicenses = $Licenses | Where-Object {
        $trialFlag = $false
        if ($_.PSObject.Properties['IsTrial'] -and $_.IsTrial -eq $true) { $trialFlag = $true }
        $freeOrTrialFlag = $false
        if ($_.PSObject.Properties['IsFreeOrTrial'] -and $_.IsFreeOrTrial -eq $true) { $freeOrTrialFlag = $true }
        $ignoreLifecycle = $false
        if ($_.PSObject.Properties['IgnoreLifecycle'] -and $_.IgnoreLifecycle -eq $true) { $ignoreLifecycle = $true }
        (-not $trialFlag) -and (-not $freeOrTrialFlag) -and (-not $ignoreLifecycle) -and ($_.SkuPartNumber -notmatch 'TRIAL|EXPLORATORY|Windows_Store|FREE|VIRAL|_FACULTY|_STUDENT')
    }

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
    param([array]$Domains)
    
    if ($Domains.Count -eq 0) {
        return "<div class='empty-state'>No domain data available</div>"
    }
    
    $analysis = Get-DomainAnalysis -Domains $Domains
    
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
    <p>This shows all domains in your tenant, their verification status, DNS configuration, and how many recipients use each domain.</p>
    <h3>Recommended Next Steps</h3>
    <p>Ensure all domains are verified. Confirm MX records point to Microsoft for mail delivery. Review unused domains for removal.</p>
</div>
"@
    
    return $tableHtml + $footerHtml
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
    $kpiHtml += New-KpiCard -Title "MDM Managed" -Value "$(Format-Number $managedPct -DecimalPlaces 1)%" -Theme $managedTheme
    
    $compliantPct = if ($totalDevices -gt 0) { ($compliantCount / $totalDevices) * 100 } else { 0 }
    $compTheme = if ($compliantPct -ge 80) { 'success' } elseif ($compliantPct -ge 60) { 'warning' } else { 'danger' }
    $kpiHtml += New-KpiCard -Title "Compliant" -Value "$(Format-Number $compliantPct -DecimalPlaces 1)%" -Theme $compTheme
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
    $kpiHtml += New-KpiCard -Title "Licensed Users" -Value (Format-Number $latest.LicensedUserCount)
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
    $conditionalAccess = if ($caCount -gt 0) { "Yes ($caCount policies)" } else { "No" }
    $appProxy = "Not collected"
    $privateAccess = "Not collected"
    $pim = "Not collected"

    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Sign-On Provider" -Value $signOnProvider
    $kpiHtml += New-KpiCard -Title "DirSync Enabled" -Value $dirSyncEnabled -Theme ($(if ($dirSyncEnabled -eq 'Yes') { 'success' } elseif ($dirSyncEnabled -eq 'No') { 'warning' } else { 'default' }))
    $kpiHtml += New-KpiCard -Title "MFA" -Value $(if ($AuthConfig -and $AuthConfig.MFAEnabled) { 'Enabled' } else { 'Not Detected' }) -Theme ($(if ($AuthConfig -and $AuthConfig.MFAEnabled) { 'success' } else { 'warning' }))
    $kpiHtml += New-KpiCard -Title "Conditional Access" -Value $conditionalAccess -Theme ($(if ($caCount -gt 0) { 'success' } else { 'warning' }))
    $kpiHtml += "</div>"

    $rows = @(
        [PSCustomObject]@{ Topic = "Sign-on provider"; Answer = $signOnProvider; Notes = if ($federatedDomains.Count -gt 0) { "Federated domains: $($federatedDomains -join ', ')" } else { "No federated domains detected" } }
        [PSCustomObject]@{ Topic = "Directory sync from on-prem AD"; Answer = $dirSyncEnabled; Notes = if ($AdConnect -and $AdConnect.Summary) { "Last sync: $($AdConnect.Summary.OnPremisesLastSyncDateTime)" } else { "No sync data available" } }
        [PSCustomObject]@{ Topic = "SSPR / Password writeback"; Answer = $ssprWriteback; Notes = "Not collected in this report" }
        [PSCustomObject]@{ Topic = "MFA provider"; Answer = $mfaProvider; Notes = if ($AuthConfig -and $AuthConfig.MFAMethods) { "Methods: $($AuthConfig.MFAMethods -join ', ')" } else { "No MFA method data found" } }
        [PSCustomObject]@{ Topic = "Enterprise SSO to SaaS apps"; Answer = $enterpriseSso; Notes = if ($ssoAppsCount -gt 0) { "Provide app list spreadsheet (separately)" } else { "No SSO apps detected" } }
        [PSCustomObject]@{ Topic = "Conditional Access"; Answer = $conditionalAccess; Notes = if ($caCount -gt 0) { "Policies detected in tenant" } else { "No policies detected" } }
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
    $enabledPolicies = $ConditionalAccessPolicies | Where-Object { $_.State -eq 'enabled' }
    $reportOnlyPolicies = $ConditionalAccessPolicies | Where-Object { $_.State -eq 'reportOnly' }
    $disabledPolicies = $ConditionalAccessPolicies | Where-Object { $_.State -eq 'disabled' }
    $mfaPolicies = $enabledPolicies | Where-Object { $_.GrantControls_BuiltInControls -match '(?i)mfa' }
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
        [PSCustomObject]@{ Metric = '% Users Not Enforced with MFA'; Value = $pctNotMfa }
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
    $syncEnabled = if ($summary.OnPremisesSyncEnabled -eq $true) { 'Yes' } else { 'No' }
    $lastSync = if ($summary.OnPremisesLastSyncDateTime) { $summary.OnPremisesLastSyncDateTime } else { 'N/A' }
    $errorCount = if ($AdConnect.ErrorCount -gt 0) { $AdConnect.ErrorCount } else { 0 }
    $syncTheme = if ($summary.OnPremisesSyncEnabled -eq $true) { 'success' } else { 'warning' }
    $errorTheme = if ($errorCount -gt 0) { 'warning' } else { 'success' }
    
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "DirSync Enabled" -Value $syncEnabled -Theme $syncTheme
    $kpiHtml += New-KpiCard -Title "Last Sync" -Value $lastSync
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

function Build-Navigation {
    param([array]$Sections)
    
    $html = "<div class='nav-container'><nav class='nav'>"
    $html += "<a href='#highlights'>Highlights</a>"
    
    foreach ($section in $Sections) {
        $html += "<a href='#$($section.Id)'>$($section.Name)</a>"
    }
    
    $html += "</nav></div>"
    return $html
}

function New-TenantHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [hashtable]$TenantStatsHash = $global:tenantStatsHash,
        
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
    
    # Extract data safely
    $licenses = ConvertTo-Array (Get-FromHash $TenantStatsHash 'LicenseSKUs')
    Write-Verbose "Licenses: Found $($licenses.Count) items"
    
    $recipients = ConvertTo-Array (Get-FromHash $TenantStatsHash 'AllRecipients')
    Write-Verbose "Recipients: Found $($recipients.Count) items"
    
    $mailboxSourceKey = if ($TenantStatsHash.ContainsKey('MailboxFullDetails')) {
        'MailboxFullDetails'
    } elseif ($TenantStatsHash.ContainsKey('AllMailboxes')) {
        'AllMailboxes'
    } else {
        $null
    }
    $mailboxes = if ($mailboxSourceKey) { ConvertTo-Array (Get-FromHash $TenantStatsHash $mailboxSourceKey) } else { @() }
    Write-Verbose "Mailboxes: Found $($mailboxes.Count) items"
    
    if ($TenantStatsHash.ContainsKey('InactiveMailboxDetails')) {
        $inactiveMailboxes = ConvertTo-Array (Get-FromHash $TenantStatsHash 'InactiveMailboxDetails')
    } else {
        $inactiveMailboxes = @($mailboxes | Where-Object { $_.PSObject.Properties['IsInactiveMailbox'] -and $_.IsInactiveMailbox -eq $true })
    }
    Write-Verbose "Inactive Mailboxes: Found $($inactiveMailboxes.Count) items"

    $publicFolders = ConvertTo-Array (Get-FromHash $TenantStatsHash 'PublicFolderDetails')
    Write-Verbose "Public Folders: Found $($publicFolders.Count) items"
    
    $sharepoint = ConvertTo-Array (Get-FromHash $TenantStatsHash 'SharePoint')
    Write-Verbose "SharePoint: Found $($sharepoint.Count) items"
    
    $onedrive = ConvertTo-Array (Get-FromHash $TenantStatsHash 'OneDrive')
    Write-Verbose "OneDrive: Found $($onedrive.Count) items"
    
    $domains = ConvertTo-Array (Get-FromHash $TenantStatsHash 'Domains')
    Write-Verbose "Domains: Found $($domains.Count) items"
    
    $devices = ConvertTo-Array (Get-FromHash $TenantStatsHash 'DeviceDetails')
    Write-Verbose "Devices: Found $($devices.Count) items"
    
    $secureScore = ConvertTo-Array (Get-FromHash $TenantStatsHash 'SecuritySecureScore')
    Write-Verbose "Secure Score: Found $($secureScore.Count) items"

    $teams = if ($TenantStatsHash.ContainsKey('AllTeams')) { ConvertTo-Array (Get-FromHash $TenantStatsHash 'AllTeams') } else { @() }
    Write-Verbose "Teams: Found $($teams.Count) items"
    $teamsVoice = $null
    if ($TenantStatsHash.ContainsKey('TeamsVoice')) {
        $teamsVoice = $TenantStatsHash['TeamsVoice']
    }

    $users = ConvertTo-Array (Get-FromHash $TenantStatsHash 'Users')
    Write-Verbose "Users: Found $($users.Count) items"
    
    $admins = ConvertTo-Array (Get-FromHash $TenantStatsHash 'Admins')
    Write-Verbose "Admins: Found $($admins.Count) items"
    
    $groups = ConvertTo-Array (Get-FromHash $TenantStatsHash 'EntraIDGroups')
    Write-Verbose "Groups: Found $($groups.Count) items"
    
    $exchangeGroups = ConvertTo-Array (Get-FromHash $TenantStatsHash 'AllExchangeGroups')
    Write-Verbose "Exchange Groups: Found $($exchangeGroups.Count) items"
    
    $conditionalAccess = ConvertTo-Array (Get-FromHash $TenantStatsHash 'ConditionalAccessPolicies')
    Write-Verbose "Conditional Access: Found $($conditionalAccess.Count) items"
    
    $authConfig = $null
    if ($TenantStatsHash.ContainsKey('AuthenticationConfig')) {
        $authContainer = $TenantStatsHash['AuthenticationConfig']
        if ($authContainer -is [hashtable] -and $authContainer.ContainsKey('Configuration')) {
            $authConfig = $authContainer['Configuration']
        }
    }
    
    $mfaRegistrationSummary = $null
    if ($TenantStatsHash.ContainsKey('MfaRegistrationSummary')) {
        $mfaRegistrationSummary = $TenantStatsHash['MfaRegistrationSummary']
    }

    $adConnect = $null
    if ($TenantStatsHash.ContainsKey('AdConnectConfiguration')) {
        $adConnect = $TenantStatsHash['AdConnectConfiguration']
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
        $licAnalysis = Get-LicenseAnalysis -Licenses $licenses
        $allFindings += $licAnalysis.Findings
        $sectionContents += @{
            Id = 'licenses'
            Name = 'License Overview'
            Content = Build-LicenseSection -Licenses $licenses
        }
    }

    # Domains
    if ($domains.Count -gt 0) {
        #Write-Host "  Building Domains section..." -ForegroundColor Gray
        $domainAnalysis = Get-DomainAnalysis -Domains $domains
        $allFindings += $domainAnalysis.Findings
        $sectionContents += @{
            Id = 'domains'
            Name = 'Domain Details'
            Content = Build-DomainsSection -Domains $domains
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
        $identityAnalysis = Get-IdentityAdminAnalysis -Users $users -Admins $admins -Groups $groups
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
    
    # Teams
    if ($teams.Count -gt 0) {
        #Write-Host "  Building Teams section..." -ForegroundColor Gray
        $sectionContents += @{
            Id = 'teams'
            Name = 'Teams Overview'
            Content = Build-TeamsSection -Teams $teams -Licenses $licenses -TeamsVoice $teamsVoice
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
    
    #endregion
    
    #region Build HTML
    
    # Get tenant name
    $tenantName = "Microsoft 365 Tenant"
    try {
        $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
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
        $htmlContent += @"
        <div class="section" id="$($section.Id)">
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

# Check for and import ImportExcel module, installing if needed
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not install ImportExcel module. Defaulting to CSV output only."
        throw
    }
}
try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Warning "Could not import ImportExcel module. Defaulting to CSV output only."
    throw $_
}

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

# Connect to Microsoft Office 365 Services
$connectOffice365Params = @{}
if ($PSBoundParameters.ContainsKey('TenantId')) { $connectOffice365Params.TenantId = $TenantId }
if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $connectOffice365Params.CertificateThumbprint = $CertificateThumbprint }
if ($PSBoundParameters.ContainsKey('ClientId')) { $connectOffice365Params.ClientId = $ClientId }
$connectionResult = Connect-Office365 @connectOffice365Params -ErrorAction Stop

#Get Export Path
$defaultReportFileName = ((Get-MgOrganization).DisplayName + " Tenant Discovery Report")
if ([string]::IsNullOrWhiteSpace($ExportPath)) {
    $ExportDetails = Get-ExportPath -FileName $defaultReportFileName
} else {
    $ExportDetails = Get-ExportPath -FileName $defaultReportFileName -UserInputPath $ExportPath
}

# Initialize list to store all discovery errors
$global:AllDiscoveryErrors = New-Object System.Collections.Generic.List[pscustomobject]

# Prompt for level of detail
if ([string]::IsNullOrWhiteSpace($ReportingMode)) {
    $reportingMode = Set-ReportMode
} else {
    $reportingMode = $ReportingMode.ToLower()
    Write-Host "You selected: $reportingMode reporting mode" -ForegroundColor Green
}
Write-Host "Output profile: $OutputProfile" -ForegroundColor Green

#Hash Table to hold final report data
$global:tenantStatsHash = @{}

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
Write-Host "Starting Office 365 Discovery Script" -ForegroundColor Black -BackgroundColor Yellow
Write-Host
Write-Host "Gathering Exchange Online Objects and data" -ForegroundColor Black -BackgroundColor Yellow
Get-AllRecipientDetails -detailLevel $reportingMode
Get-AllExchangeMailboxDetails -detailLevel $reportingMode
Get-ExchangeGroupDetails -detailLevel $reportingMode

Get-MailFlowRulesandConnectors -detailLevel $reportingMode
Get-AllPublicFolderDetails -detailLevel $reportingMode
Write-Host

Write-Host "Gathering Hybrid and Configuration Details" -ForegroundColor Black -BackgroundColor Yellow
Get-ExchangeHybridConfiguration -detailLevel $reportingMode
Get-FederationAndCrossTenantConfiguration
Get-ThirdPartySpamFilteringConfig
Get-SMTPRelayConfiguration
Write-Host

Write-Host "Gathering Tenant Objects and License details" -ForegroundColor Black -BackgroundColor Yellow
$GraphTest = if (Get-MgContext -ErrorAction SilentlyContinue) { "SDK" } elseif ($global:GraphHeaders) { "API" }
# Determine if using REST or SDK Graph API
switch ($GraphTest) {
    "REST" {
        Write-Verbose "Attempting to use Microsoft Graph REST API for Tenant Object and License details"
        Get-GraphUserStats
        #Get-SharePointAndOneDriveSites -detailLevel $reportingMode -ServiceName API
        Get-EntraIDGroups -detailLevel $reportingMode -GraphAuthType REST
     }
    "SDK" {
        Write-Verbose "Attempting to use Microsoft Graph SDK for Tenant Object and License details"
        #Get-TeamsDetails -detailLevel $reportingMode
        Get-ConditionalAccessPoliciesReport -detailLevel $reportingMode
        Get-AllLicenseSKUs
        Get-AllUserDetails -detailLevel $reportingMode
        Get-TeamsVoiceDetails
        Get-EntraIDGroups -detailLevel $reportingMode -GraphAuthType SDK
        Get-AllOffice365Domains
        Get-AllOffice365Admins
        Get-AllDevicesReport -detailLevel $reportingMode
        Get-TenantOverviewInfo
        Get-AuthenticationConfiguration -detailLevel $reportingMode
        Get-AdConnectSyncDetails
        Get-MfaRegistrationDetails
        Get-SecuritySecureScoreReport -detailLevel $reportingMode -MostRecent
     }
}
Write-Host

Write-Host "Gathering Collaboration/SharePoint Objects and data" -ForegroundColor Black -BackgroundColor Yellow
Get-AllUnifiedGroups -detailLevel $reportingMode
$sharePointDiscoveryService = if ($connectionResult -and $connectionResult.SharePointOnline) { 'SPO' } else { 'MGGraph' }
Get-SharePointAndOneDriveSites -detailLevel $reportingMode -ServiceName $sharePointDiscoveryService
Write-Host


#Combine Reporting - Optional
if ($reportingMode -eq "combined" -or $reportingMode -eq "all") {
    Write-Host
    Write-Host "Consolidating Discovery Report data for each user / object into one file" -ForegroundColor Black -BackgroundColor Green
    #Combine Reports
    Report-UserAndMailboxStats
}

Update-AssessmentReportTables -TenantStatsHash $global:tenantStatsHash
Update-ConfigurationSummaryTables -TenantStatsHash $global:tenantStatsHash


########################################################
### Export Reports ###
########################################################

#Exclude specific reports from Export
Write-Host
Write-Host "Exporting Tenant Statistics" -ForegroundColor Black -BackgroundColor Yellow
$ExportTenantStatsHash = Filter-TenantStatsHash -tenantStatsHash $global:tenantStatsHash -reportingMode $reportingMode -GraphTest $GraphTest

#Export Reports: Exports each individual hashtable to own CSV file and then combines into Excel file
Write-Log -Type INFO -Message "Exporting the Tenant Statistics to $($ExportDetails)." -ExportFileLocation $ExportDetails
try {
    Export-HashTableToExcel -hashtable $ExportTenantStatsHash -ExportDetails $ExportDetails
}
catch {
    Write-Log -Type ERROR -Message "An error occurred in Exporting the Tenant Statistics to $($ExportDetails). Please re-run the script and verify the location is valid and the file is not open in another application. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
    Write-Log -Type ERROR -Message "An error occurred in Exporting the Tenant Statistics to $($ExportDetails). Please re-run the script and verify the location is valid and the file is not open in another application. $($_.Exception.Message)" -ExportFileLocation $ExportDetails
}

# Export JSON snapshot for reuse
if ($effectiveSkipJsonReport) {
    Write-Log -Type INFO -Message "Skipping JSON report generation because -SkipJsonReport was provided." -ExportFileLocation $ExportDetails
} else {
    try {
        $jsonExportPath = $ExportDetails -replace '\.xlsx$', '.json'
        Export-TenantStatsJson -TenantStatsHash $ExportTenantStatsHash -Path $jsonExportPath
        Write-Log -Type INFO -Message "Exported Tenant Statistics JSON to $jsonExportPath" -ExportFileLocation $ExportDetails
    } catch {
        Write-Log -Type WARNING -Message "Unable to export Tenant Statistics JSON: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }
}

try {
    if (Get-Command -Name Export-TenantToTenantQuestionnaireMarkdown -ErrorAction SilentlyContinue) {
        $questionnaireTemplateCandidates = @(
            [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\..\docs\Microsoft 365 Tenant to Tenant Questionnaire.md')),
            [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\..\docs\templates\Microsoft 365 Tenant to Tenant Questionnaire.md'))
        )
        $questionnaireTemplatePath = $questionnaireTemplateCandidates | Where-Object { Test-Path -Path $_ } | Select-Object -First 1
        if (-not $questionnaireTemplatePath) {
            throw "Questionnaire template not found in expected locations: $($questionnaireTemplateCandidates -join '; ')"
        }
        $questionnaireExportPath = $ExportDetails -replace '\.xlsx$', '-TenantToTenantQuestionnaire.md'
        Export-TenantToTenantQuestionnaireMarkdown -TenantStatsHash $global:tenantStatsHash -TemplatePath $questionnaireTemplatePath -Path $questionnaireExportPath
        Write-Log -Type INFO -Message "Exported Tenant to Tenant Questionnaire to $questionnaireExportPath" -ExportFileLocation $ExportDetails
    } else {
        Write-Log -Type WARNING -Message "Skipping questionnaire export because Export-TenantToTenantQuestionnaireMarkdown is unavailable." -ExportFileLocation $ExportDetails
    }
} catch {
    Write-Log -Type WARNING -Message "Unable to export Tenant to Tenant Questionnaire: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
}

Write-Host ""

#Export Errors
if ($global:AllDiscoveryErrors.Count -gt 0) {
    #Write-Host "Exporting Error Reports" -ForegroundColor Black -BackgroundColor Yellow
    try {
        Export-ErrorReports -ExportFileLocation $ExportDetails -ErrorData $global:AllDiscoveryErrors -logReportDirectory $ExportDetails
        }
    catch {
        Write-Log -Type ERROR -Message "An error occurred in Exporting the Error Reports. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
        Write-Log -Type ERROR -Message "An error occurred in Exporting the Error Reports. $($_.Exception.Message)" -ExportFileLocation $ExportDetails
    }
}

########################################################
### HTML Report Generation (Commented Out) ###
########################################################
Write-Host ""
Write-Host "Generating HTML Report..." -ForegroundColor Black -BackgroundColor Yellow

try {
    if (Get-Command New-TenantAssessmentHtmlReport -ErrorAction SilentlyContinue) {
        $assessmentHtmlPath = $ExportDetails -replace '\.xlsx$', '-Assessment.html'
        $assessmentHtmlResult = New-TenantAssessmentHtmlReport -TenantStatsHash $global:tenantStatsHash -OutputPath $assessmentHtmlPath
        if ($assessmentHtmlResult.Success) {
            Write-Log -Type INFO -Message "Assessment HTML report generated: $($assessmentHtmlResult.OutputPath)" -ExportFileLocation $ExportDetails
        } else {
            Write-Log -Type WARNING -Message "Assessment HTML report generation failed: $($assessmentHtmlResult.Error)" -ExportFileLocation $ExportDetails
        }
    } else {
        Write-Log -Type WARNING -Message "Skipping assessment HTML report generation because New-TenantAssessmentHtmlReport is unavailable." -ExportFileLocation $ExportDetails
    }
} catch {
    Write-Log -Type WARNING -Message "Error generating assessment HTML report: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
}

if ($effectiveSkipHtmlReport) {
    Write-Host "Skipping full HTML report generation because -SkipHtmlReport was provided or the Lean output profile is active." -ForegroundColor Yellow
    Write-Log -Type INFO -Message "Skipping full HTML report generation because -SkipHtmlReport was provided or the Lean output profile is active." -ExportFileLocation $ExportDetails
}
elseif (-not (Get-Command New-TenantHtmlReport -ErrorAction SilentlyContinue)) {
    Write-Warning "New-TenantHtmlReport function is not available. Skipping full HTML report generation."
    Write-Log -Type WARNING -Message "Skipping full HTML report generation because New-TenantHtmlReport is unavailable." -ExportFileLocation $ExportDetails
}
else {
    try {
        # Define custom thresholds (optional - remove this block to use defaults)
        $reportThresholds = @{
            LicenseUtilization = 85
            MailboxSizeGB = 50
            ArchiveSizeGB = 50
            SharePointSiteGB = 1024
            OneDriveSiteGB = 1024
            DeviceStaleMonths = 6
            DeviceCompliancePercent = 80
        }
        
        # Generate the HTML report
        Write-Log -Type INFO -Message "Generating HTML report from hashtable data" -ExportFileLocation $ExportDetails
        
        # Enable verbose output for debugging (comment out after testing)
        #$VerbosePreference = 'SilentlyContinue'
        $HTMLExportPath = $ExportDetails -replace '\.xlsx$', '.html'
        
        $htmlResult = New-TenantHtmlReport `
            -TenantStatsHash $global:tenantStatsHash `
            -Thresholds $reportThresholds `
            -OutputPath $HTMLExportPath
        
        # Reset verbose preference
        #$VerbosePreference = 'SilentlyContinue'
        
        if ($htmlResult.Success) {
            ##Write-Host "   Location: $($htmlResult.OutputPath)" -ForegroundColor Cyan
            #Write-Host "   Sections: $($htmlResult.SectionCounts.TotalSections)" -ForegroundColor Gray
            #Write-Host "   Findings: $($htmlResult.SectionCounts.TotalFindings)" -ForegroundColor Gray
            
            Write-Log -Type INFO -Message "HTML report generated: $($htmlResult.OutputPath)" -ExportFileLocation $ExportDetails
            #Write-Log -Type INFO -Message "HTML report sections: $($htmlResult.SectionCounts.TotalSections)" -ExportFileLocation $ExportDetails
            #Write-Log -Type INFO -Message "HTML report findings: $($htmlResult.SectionCounts.TotalFindings)" -ExportFileLocation $ExportDetails

            if ($effectiveSkipPdfReport) {
                Write-Host "Skipping PDF report generation because -SkipPdfReport was provided or the selected output profile disables it." -ForegroundColor Yellow
                Write-Log -Type INFO -Message "Skipping PDF report generation because -SkipPdfReport was provided or the selected output profile disables it." -ExportFileLocation $ExportDetails
            }
            elseif (-not (Get-Command Export-TenantHtmlReportPdf -ErrorAction SilentlyContinue)) {
                Write-Warning "PDF report helper is unavailable. HTML report was generated, but PDF export was skipped."
                Write-Log -Type WARNING -Message "Skipping PDF report generation because Export-TenantHtmlReportPdf is unavailable." -ExportFileLocation $ExportDetails
            }
            else {
                try {
                    $pdfExportPath = $htmlResult.OutputPath -replace '\.html$', '.pdf'
                    $pdfResult = Export-TenantHtmlReportPdf -HtmlPath $htmlResult.OutputPath -PdfPath $pdfExportPath

                    if ($pdfResult.Success) {
                        Write-Log -Type INFO -Message "PDF report generated: $($pdfResult.PdfPath) using $($pdfResult.Renderer)" -ExportFileLocation $ExportDetails
                    } else {
                        Write-Warning "PDF report generation failed: $($pdfResult.Error)"
                        Write-Log -Type WARNING -Message "PDF report generation failed: $($pdfResult.Error)" -ExportFileLocation $ExportDetails
                    }
                } catch {
                    Write-Warning "Error generating PDF report: $($_.Exception.Message)"
                    Write-Log -Type WARNING -Message "Error generating PDF report: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                }
            }
        } else {
            Write-Warning "HTML report generation failed: $($htmlResult.Error)"
            Write-Log -Type ERROR -Message "HTML report generation failed: $($htmlResult.Error)" -ExportFileLocation $ExportDetails
        }
        
    } catch {
        Write-Warning "Error generating HTML report: $($_.Exception.Message)"
        Write-Warning "Excel report is still available at: $ExportDetails"
        Write-Log -Type ERROR -Message "HTML report generation error: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
        
        # Capture error but don't stop script
        Write-Log -Type ERROR -Message "An error occurred generating HTML report. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
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

Write-Host ""
Write-Host "Report Generation Complete in $($timeString)" -ForegroundColor Black -BackgroundColor Green
#Write-Host "Excel Report: $ExportDetails" -ForegroundColor Cyan

########################################################
### End of HTML Report Integration ###
########################################################

Write-Log -Type INFO -Message "COMPLETED: Gathered Tenant Details. Completed Time: $($timeString)" -ExportFileLocation $ExportDetails
