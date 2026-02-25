########################################################
#region - Intial Variables and Functions
########################################################

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

#region - Default Script Helper Functions
# ----------------------------------

function Capture-ErrorHelper {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecordVar,
        [Parameter(Mandatory=$true)]
        [string]$errorMessage
    )
    Office365Custom\Capture-ErrorHelper -ErrorRecordVar $ErrorRecordVar -errorMessage $errorMessage
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
        [string]$LogPath,

        [Parameter(Mandatory=$false)]
        [string]$ExportFileLocation

    )
    if ($Type -ne 'ERROR') {
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

        Office365Custom\Write-Log @logSplat
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
            $LogFile = $LogPath + "\FullActiveDirectoryReportLog.txt"
            $logMessage | Out-File -Append -FilePath $LogFile
        } catch {
            Write-Error "Failed to write to log file at '$LogFile': $_"
        }
    }
    elseif ($ExportFileLocation) {
        $directory = [System.IO.Path]::GetDirectoryName($ExportFileLocation)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ExportFileLocation)
        $txtFileName = $baseName + "-FullReportLog.txt"
        $newLogFolder = $baseName + " Discovery Log Reporting"
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
}

function Install-ImportExcelModule {
    # Check if ImportExcel module is installed
    if (!(Get-Module -ListAvailable -Name ImportExcel)) {
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not install ImportExcel module. Defaulting to CSV output only."
            return $false
        }
    }

    # Import ImportExcel module
    try {
        Import-Module ImportExcel -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not import ImportExcel module. Defaulting to CSV output only."
        return $false
    }

    return $true
}

function Set-ReportMode {
    param (
        [switch]$ShowExplanationsOnStart = $true
    )

    function ShowModeExplanations {
        Write-Host "Reporting Mode Explanations:" -ForegroundColor Yellow
        Write-Host "Minimum   - Provides basic details for a quick overview."
        Write-Host "Combined  - Combines details from similar reports. E.g., combining a user's mailbox and SharePoint data."
        Write-Host "All       - Provides a comprehensive, detailed report that includes all possible details and combined reports."
        Write-Host "Geek      - Provides every available detail from reports, does not include Combined reports"

    }

    if ($ShowExplanationsOnStart) {
        ShowModeExplanations
        Write-Host ""
    }

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
#endregion - Default Script Helper Functions

# ----------------------------------
#region - Export Script Functions
# ----------------------------------

#Convert Hash Table to Custom Object Array for Export
function Convert-HashToArray {
    [CmdletBinding()]
    param ( [Parameter(Mandatory=$true)]
        [Hashtable]$HashToConvert,
        [Parameter(Mandatory=$False)]
        [String]$table        
    )
    $ExportArray = @()
    $start = Get-Date
    $totalCount = ($HashToConvert.keys | measure).count    
    $progresscounter = 0

    foreach ($nestedKey in $HashToConvert.Keys) {
        $progresscounter++
        Write-ProgressHelper -Total $totalCount -Index $progresscounter -Activity "Converting Hash Table" -Operation "Converting $($nestedKey)"
        #Define the attributes
        $attributes = $HashToConvert[$nestedKey]
        # If the attributes are a hashtable, convert them to a custom object
        if ($attributes -is [hashtable] -or $attributes -is [System.Collections.Specialized.OrderedDictionary]) {
            #Write-Verbose "Attributes are a hashtable"
            $customObject = New-Object -TypeName PSObject

            # Add the key as a property
            foreach ($attribute in $attributes.keys) {
                $customObject | Add-Member -MemberType NoteProperty -Name "$($attribute)" -Value ($attributes[$attribute] -join ';')
            }
            
            $ExportArray += $customObject
        } 
        # If the attributes are an array, add them to the export array
        elseif ($attributes -is [array] -or $attributes -is [PSCustomObject] ) {
            $ExportArray += $attributes
        }
        # If the attributes are a string, add them to the export array
        else {
            $ExportArray += $attributes
        }
    }
    Write-ProgressHelper -Total $totalCount -Activity "Converting Hash Table" -Completed
    Return $ExportArray
}

#Export Hash Table to Excel; export each key in hash to csv and then combine into excel
function Export-HashTables {
    [CmdletBinding()]
    param ( [Parameter(Mandatory=$True)] [Hashtable]$hashtable,
        [Parameter(Mandatory=$True)] [Array]$ExportDetails
	)

    #Combine temporary csv files created from hash key from temp folder
    function Combine-TempCSVFiles {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory=$false)]
            [Hashtable]$HashToConvert,
            [Parameter(Mandatory=$False)]
            [String]$table
        )
        #Combine files with global tnenant value from temp folder
        Write-Progress -Activity "Combine All CSV Files into One Excel" -Status (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
        $ExportedCSVFiles = Get-ChildItem -Path $env:TEMP -Filter "*$global:ActiveDirectory*.csv"
    
        if (-not $ExportedCSVFiles) {
            $ExportedCSVFiles = Get-ChildItem -Path $env:TEMP -Filter "*ArrayaDiscoveryReport*.csv"
        }
        
        # Check again if no files are found
        if (-not $ExportedCSVFiles) {
            Write-Warning "No CSV files found in $env:TEMP matching the provided patterns."
        }
        
        
        $totalCount = ($ExportedCSVFiles | Measure).count
        $progresscounter = 0
        $start = Get-Date
        foreach ($file in $ExportedCSVFiles) {
            $progresscounter++
            $worksheetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    
            # Check if the filename contains $global:ActiveDirectory and replace if it exists
            if ($file.Name -like "*$global:ActiveDirectory*") {
                $worksheetName = $worksheetName.Replace("_$global:ActiveDirectory","")
            }
            # Check if the filename contains "_ArrayaDiscoveryReport" and replace if it exists
            elseif ($file.Name -like "*_ArrayaDiscoveryReport*") {
                $worksheetName = $worksheetName.Replace("_ArrayaDiscoveryReport","")
            }
            
            Write-ProgressHelper -Total $totalCount -Index $progresscounter -Activity "Adding Worksheets to Excel File" -Operation "Adding Worksheet $($worksheetName) to $($ExportDetails[0])"
            Import-Csv -Path $file.FullName | Export-Excel -Path $ExportDetails[0] -WorksheetName $worksheetName -ClearSheet
            Write-ProgressHelper -Total $totalCount -Activity "Adding Worksheets to Excel File" -Completed
    
        }
        #Delete the temporary CSV files
        Get-ChildItem -Path $env:TEMP -Filter "*$global:ActiveDirectory*.csv" | Remove-Item
    }

    # Progress Bar
    $totalCount = ($hashtable.keys | measure).count
    $progresscounter = 0
    $start = Get-Date

    # Check if the Excel module is available
    $excelModuleAvailable = Get-Module -ListAvailable -Name "ImportExcel" -ErrorAction SilentlyContinue
  
    if (-not $excelModuleAvailable) {
        Write-Warning "Excel module 'ImportExcel' is not available. Exporting to CSV."
        # Export each hashtable to a separate CSV file in the $ExportDetails path
        $directory = [System.IO.Path]::GetDirectoryName($ExportDetails)        
        foreach ($table in $hashtable.Keys) {
            $ExportStatsArray = Convert-HashToArray -table $table -HashToConvert $hashtable[$table]
            Write-Log -Type DEBUG -Message ("({0}/{1}) Export '{2}' Array to CSV in temp folder {3}" -f $progressCounter, $totalCount, $table, $tempPath) -ExportFileLocation $ExportDetails[0]
            $csvPath = Join-Path -Path $directory -ChildPath ("$table.csv")
            $ExportStatsArray | Export-Csv -Path $csvPath -NoTypeInformation -Force
        }
        Write-Host "ActiveDirectory Report located at: $($directory)" -ForegroundColor Green
        return
    }
    
    ## Combine all CSV files into a single Excel file
    # Export each hashtable to a separate CSV file
    foreach ($table in $hashtable.Keys){
        try {
            $progresscounter++
            #Export All provided Tables
            Write-ProgressHelper -Id 2 -Total $totalCount -Index $progresscounter -Activity "Exporting '$($table)' Hash To Excel"
            Write-Log -Type DEBUG -Message ("({0}/{1}) Converting '{2}' Hash Table to Array" -f $progressCounter, $totalCount, $table) -ExportFileLocation $ExportDetails[0]            

            $ExportStatsArray = Convert-HashToArray -table $table -HashToConvert $hashtable[$table]
                        
            # Use the table name to create a unique temporary file
            $tempPath = Join-Path -Path $env:TEMP -ChildPath ("{0}_{1}.csv" -f $table, "ArrayaDiscoveryReport")
            Write-Log -Type DEBUG -Message ("({0}/{1}) Export '{2}' Array to CSV in temp folder {3}" -f $progressCounter, $totalCount, $table, $tempPath) -ExportFileLocation $ExportDetails[0]            
            $ExportStatsArray | Export-Csv -Path $tempPath -NoTypeInformation -Encoding UTF8
            Write-ProgressHelper -Id 2 -Total $totalCount -Activity "Exporting $($table) Hash To Temp Folder" -Completed
        }
        catch {
            Write-Log -Type Error -Message "An error occurred in converting Hash to Array for $($table). $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]            
            Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in converting Hash To Array for $($table). $($_.Exception.Message)"`r`n            throw $_
        }
    }

    try {
        #Combine files with global tenant value from temp folder
        Combine-TempCSVFiles -errorAction Stop
        Write-Host "ActiveDirectory Report located at: $($ExportDetails[0])" -ForegroundColor Green

        #Delete the temporary CSV files
        Get-ChildItem -Path $env:TEMP -Filter "*ArrayaDiscoveryReport*.csv" | Remove-Item
    }
    catch {
        Write-Log -Type Error -Message "An error occurred in Exporting the ActiveDirectory Statistics to $($ExportDetails[0]). $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]            
        Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Exporting the ActiveDirectory Statistics to $($ExportDetails[0]). Please check if the location is valid or if the file is open in another application. $($_.Exception.Message)"`r`n        try {
            Write-Log -Type DEBUG -Message "Attempt Number 2 to Export ActiveDirectory Statistics to $($ExportDetails[0])" -ExportFileLocation $ExportDetails[0]            
            $ExportDetails = Get-ExportPath
            #Combine files with global tnenant value from temp folder
            Combine-TempCSVFiles -errorAction Stop
            Write-Host "ActiveDirectory Report located at: $($ExportDetails[0])" -ForegroundColor Green

            #Delete the temporary CSV files
            Get-ChildItem -Path $env:TEMP -Filter "*ArrayaDiscoveryReport*.csv" | Remove-Item
        }
        catch {
            Write-Log -Type Error -Message "Second Attempt to Export ActiveDirectory Statistics failed to $($ExportDetails[0]). $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]            
            Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "Second Attempt to Export ActiveDirectory Statistics failed to $($ExportDetails[0]). $($_.Exception.Message)"`r`n            throw $_
        } 
    }  
}

#Function to get Export Path
function Get-ExportPath {
    $fileName = "$global:ActiveDirectory-ActiveDirectoryStats"
    $fullPath = Office365Custom\Get-ExportPath -FileName $fileName -DefaultExtension ".xlsx"
    $fileNameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension((Split-Path -Path $fullPath -Leaf))

    Write-Host ""
    return $fullPath, $fileNameWithoutExtension
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
        Write-Log -Type WARNING -Message "No error data provided to export. Exiting function." -ExportFileLocation $ExportDetails[0]
        return
    }

    Write-Log -Type INFO -Message "START: Export all Errors" -ExportFileLocation $ExportDetails[0]

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

    Write-Log -Type INFO -Message "INFO: Exporting Error Logs to directory $($errorReportFolderDirectory)" -ExportFileLocation $ExportDetails[0]

    try {
        if (-not (Test-Path $errorReportFolderDirectory)) {
            $result = New-Item -Path $errorReportFolderDirectory -ItemType Directory
            Write-Log -Type INFO -Message "INFO: Error Report Directory '$($errorReportFolderDirectory)' does not exist. Created Folder Directory" -ExportFileLocation $ExportDetails[0]
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
            Write-Log -Type INFO -Message "INFO: Exported $($_.Key) Error Logs to directory $($_.Value)" -ExportFileLocation $ExportDetails[0]
        }

    } catch {
        Write-Error "Failed to export error reports: $_"
    }
}

#endregion - Export Script Functions

# ----------------------------------
#region - Active Directory Specific Functions
# ----------------------------------

function Get-AllADForests {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    
    # Instantiate the hash table
    $ADReportHash["ADForests"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting All AD Forests with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-ADForests] START: Getting All AD Forests with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    # AD Recycle BIN Status
    $ADRecycleBIN = Get-ADOptionalFeature -filter {Name -eq 'Recycle Bin Feature'} | Select-Object -ExpandProperty EnabledScopes
    Write-Log -Type INFO -Message "[Get-AllADForests] Getting AD Recycle Bin details" -ExportFileLocation $ExportDetails[0]
    
    If (!$ADRecycleBIN){
        $ADRecycleBIN = 'Disabled'
    } else {
        $ADRecycleBIN = 'Enabled'
    }

    # Get the Domain Functional Level 
    $DomainMode = ($global:DNSRoot | foreach { Get-ADDomain -Identity $_ }  | Select-Object -ExpandProperty DomainMode) -join ' , '
    Write-Log -Type INFO -Message "[Get-AllADForests] Getting Domain Mode details" -ExportFileLocation $ExportDetails[0]

    # Get the Domain NetBIOS Name
    $NetBIOSName = $global:domainInfo.netBIOSName
    Write-Log -Type INFO -Message "[Get-AllADForests] Getting Net BIOS Name details" -ExportFileLocation $ExportDetails[0]

    # Export all Forest data
    $ADForests = Get-ADForest

    foreach ($forest in $ADForests) {
        Write-Log -Type INFO -Message "[Get-ADForests] Getting AD Forests details" -ExportFileLocation $ExportDetails[0]

        # Domain and Forest Information Output Object
        $DomainAndForestOutputObj  = New-Object -TypeName PSObject
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "Name" -Value $forest.Name -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "RootDomain" -Value $forest.RootDomain -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "Domain-NetBIOSName" -Value $NetBIOSName -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "Domain-FunctionalLevel" -Value $DomainMode -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "Forest-FunctionalLevel" -Value $forest.ForestMode -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "ADRecycleBIN" -Value $ADRecycleBIN -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "SchemaMaster" -Value $forest.SchemaMaster -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "DomainNamingMaster" -Value $forest.DomainNamingMaster -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "SchemaMaster" -Value $forest.SchemaMaster -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "SchemaMaster" -Value $forest.SchemaMaster -Force

        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "Domains" -Value ($forest.Domains -join ",") -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "GlobalCatalogs" -Value ($forest.GlobalCatalogs -join ",") -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "Sites" -Value ($forest.Sites -join ",") -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "ApplicationPartitions" -Value ($forest.ApplicationPartitions -join ",") -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "SPNSuffixes" -Value ($forest.SPNSuffixes -join ",") -Force
        $DomainAndForestOutputObj | Add-Member -MemberType NoteProperty -Name "UPNSuffixes" -Value ($forest.UPNSuffixes -join ",") -Force

        $ADReportHash["ADForests"][$forest.Name] = $DomainAndForestOutputObj
    }

    # Completed Details - Output and Log
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-DomainAndForestInfo] COMPLETED: Gathering Forests Details in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

# Get Domain Information
function Get-AllADDomains {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    # Instantiate the hash table
    $ADReportHash["DomainSites"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting All Domains with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllADDomains] START: Getting All Domains with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    $allDomains = $global:forestInfo.Domains
    $allDomains | foreach {
        $domainInfo = Get-ADDomain -Identity $_
        $ADReportHash["DomainSites"][$domainInfo.DNSRootame] = $domainInfo
    }
    Write-Log -Type INFO -Message "[Get-AllADDomains] Added All Domains to Hash" -ExportFileLocation $ExportDetails[0]

    # Completed Details - Output and Log
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AllADDomains] COMPLETED: Gathering All Domains in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

# Get List of Domain Controllers
function Get-AllDomainControllers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    # Instantiate the hash table
    $ADReportHash["DomainControllers"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting All Domain Controllers with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllDomainControllers] START: Getting All Domain Controllers with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    #Get Domain Controllers
    $allDomainControllers = Get-ADDomainController -Filter * | Sort Name
    Write-Log -Type INFO -Message "[Get-AllDomainControllers] Gathered All Domain Controllers" -ExportFileLocation $ExportDetails[0]
    
    #Add to Hash
    foreach ($DC in $allDomainControllers) {
        $allDCComputerProperties = Get-ADComputer -Identity $DC.Name -Properties *
        $DC | Add-Member -MemberType NoteProperty -Name "Site" -Value (Get-ADReplicationSite -Identity $DC.Site) -Force
        $DC | Add-Member -MemberType NoteProperty -Name "OperationMasterRoles" -Value ($DC.OperationMasterRoles -join ', ') -Force
        $DC | Add-Member -MemberType NoteProperty -Name "LastChanged" -Value $allDCComputerProperties.whenChanged -Force
        $DC | Add-Member -MemberType NoteProperty -Name "LastLogon" -Value $allDCComputerProperties.LastLogon -Force
        $DC | Add-Member -MemberType NoteProperty -Name "Partitions" -Value ($DC.Partitions -join ';') -Force
        $ADReportHash["DomainControllers"][$DC.Name] = $DC
    }

    # Completed Details - Output and Log
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AllDomainControllers] COMPLETED: Gathering All Domain Controllers in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

#Get FSMO Roles
function Get-ADFSMORoles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    # Instantiate the hash table
    $ADReportHash["FSMORoles"] = @{}
    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting All FSMO Roles with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-ADFSMORoles] START: Getting All FSMO Roles with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    $forest = Get-ADForest
    $schemaMaster = $global:forestInfo.SchemaMaster
    $domainNamingMaster = $global:forestInfo.DomainNamingMaster

    $domains = Get-ADDomain
    foreach ($domain in $domains) {
        Write-Log -Type DEBUG -Message "[Get-ADFSMORoles] Getting All FSMO Roles for domain '$($domain.Name)'" -ExportFileLocation $ExportDetails[0]
        $domainRoleHolders = [PSCustomObject]@{
            Domain = $domain.Name
            SchemaMaster = $schemaMaster
            DomainNamingMaster = $domainNamingMaster
            InfrastructureMaster = $domain.InfrastructureMaster
            RIDMaster = $domain.RIDMaster
            PDCEmulator = $domain.PDCEmulator
        }
        $ADReportHash["FSMORoles"][$domain.Name] = $domainRoleHolders
    }

    # Completed Details - Output and Log
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-ADFSMORoles] COMPLETED: Gathering All FSMO Roles in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

#Get DNS Info
function Get-DNSInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )
    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting All DNS with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-DNSInfo] START: Getting All DNS with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    # Ensure the DNS Server module is available
    try {
        Import-Module DNSServer -ErrorAction Stop
        # Continue with your script if the module is successfully imported
    }
    catch {
        Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in gathering DNS details. DNS Server module is not available. $($_.Exception.Message)"`r`n        Write-Log -Type ERROR -Message "An error occurred in gathering DNS details. DNS Server module is not available. $($_.Exception.Message)"-ExportFileLocation $ExportDetails[0]
        return
    }
    if (-not (Get-Module -ListAvailable -Name DNSServer)) {
        Write-Host "An error occurred in gathering. DNS Server module is not available." -ForegroundColor Red
        return
    }

    # Instantiate the hash table
    $ADReportHash["DNS"] = @{}

    # Define the DNS server to query
    $dnsServer = $global:PDCEmulator

    # DNS Forward Lookup Zones
    $ForwardLookupZones = (Get-DnsServerZone -ComputerName $global:PDCEmulator | Where-Object {$_.IsReverseLookupZone -eq $False} | Select-Object -ExpandProperty ZoneName) -join ', '
    $ADReportHash["DNS"]["ForwardLookupZones"] = $ForwardLookupZones

    # DNS Reverse Lookup Zones
    $ReverseLookupZones = (Get-DnsServerZone -ComputerName $global:PDCEmulator | Where-Object {$_.IsReverseLookupZone -eq $True} | Select-Object -ExpandProperty ZoneName) -join ', '
    $ADReportHash["DNS"]["ReverseLookupZones"] = $ReverseLookupZones

    # NS records
    $NSRecords = (Resolve-DnsName -Name $global:DNSRoot -Type NS | Where-Object {$_.QueryType -eq 'NS'} | Select-Object -ExpandProperty NameHost) -join ', '
    $ADReportHash["DNS"]["NSRecords"] = $NSRecords

    # MX Records
    $MXRecords = (Resolve-DnsName -Name $global:DNSRoot -Type MX | Where-Object {$_.QueryType -eq 'MX'} | Select-Object -ExpandProperty NameExchange) -join ', '
    $ADReportHash["DNS"]["MXRecords"] = $MXRecords

    # Forwarders
    $DNSForwarders = (Get-DnsServerForwarder -ComputerName $global:PDCEmulator | Select-Object -ExpandProperty IPAddress) -join ', '
    $ADReportHash["DNS"]["DNSForwarders"] = $DNSForwarders

    # Scavenging (Returns True or False)
    $DNSScavenging = (Get-DnsServerScavenging -ComputerName $global:PDCEmulator).ScavengingState
    $ADReportHash["DNS"]["DNSScavenging"] = $DNSScavenging

    # Aging (Returns True or False)
    $DNSAging = (Get-DnsServerZoneAging -Name $global:DNSRoot -ComputerName $global:PDCEmulator).AgingEnabled
    $ADReportHash["DNS"]["DNSAging"] = $DNSAging
    
    # Completed Details - Output and Log
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-DNSInfo] COMPLETED: Gathering All DNS in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

#Get DHCP Info
function Get-DHCPInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel    
    )

     # Suppress progress bar
    $originalProgressPreference = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'

    # Instantiate the hash table
    $ADReportHash["DHCP"] = @{}
    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting DHCP Server with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-DHCPInfo] START: Getting DHCP Server with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    try {
        $DHCP = Get-WindowsFeature -name DHCP -ea Stop | Where-Object {$_.Installed -eq $True}
        $DHCPServers = Get-DhcpServerInDC -ea Stop

        # Create a custom object from the values above
        foreach ($DHCPServer in $DHCPServers) {
            #Cleare variables
            $DHCPStatus = $null
            $DHCPScope = @()
            $DHCPScopeStat = @()
            $DHCPSettings = @()
            try {
                $DHCPStatus = "Success"
                Write-Log -Type DEBUG -Message "[Get-DHCPInfo] Gathered DHCP Server details for '$($dhcp.Name)" -ExportFileLocation $ExportDetails[0]
                $DHCPScope = Get-dhcpserverv4scope -ComputerName $DHCPServer.DnsName
                $DHCPScopeStat = Get-dhcpserverv4scopestatistics -ComputerName $DHCPServer.DnsName
                $DHCPSettings = Get-dhcpserversetting -ComputerName $DHCPServer.DnsName
            }
            catch {
                $DHCPStatus = "Failed"
                Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in gathering DHCP Server Details. Failed to get DHCP Server information: DHCP Server may not be installed. $($_.Exception.Message)"`r`n                Write-Log -Type ERROR -Message "[Get-DHCPInfo] Unable to find DHCP. DHCP Server may not be installed. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
            }
            # DHCP Server Report Details
            $DHCPObject = [PSCustomObject]@{
                DHCPStatus = $DHCPStatus
                Name = $dhcp.Name
                FQDN = $DHCPServer.DNSName
                IsDomainJoined = $DHCPSettings.IsdomainJoined
                NAPEnabled = $DHCPSettings.NapEnabled
                IPAddress = $DHCPServer.IPAddress
                SubnetMask = $DHCPScope.SubnetMask
                StartRange = $DHCPScope.StartRange
                EndRange = $DHCPScope.EndRange
                AddressesFree = $DHCPScopeStat.AddressesFree
                AddressesInUse = $DHCPScopeStat.AddressesInUse
                PercentageinInuse = $DHCPScopeStat.PercentageInUse
                ReservedAddress = $DHCPScopeStat.Reserved
                InstallState = $dhcp.InstallState
                
            }
            #Collate report details
            $ADReportHash["DHCP"][$dhcp.Name] = $DHCPObject          
        }

        # Completed Details - Output and Log
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-DHCPInfo] COMPLETED: Gathering DHCP Server in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
    catch {
        if ($_.Exception.Message -eq "Invalid class ") {
            Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in gathering DHCP Server Details. Failed to get DHCP Server information: DHCP Server may not be installed."`r`n            Write-Log -Type ERROR -Message "[Get-DHCPInfo] Unable to find DHCP. DHCP Server may not be installed." -ExportFileLocation $ExportDetails[0]
        }
        else {
            Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in gathering DHCP Server Details. Failed to get DHCP Server information: DHCP Server may not be installed. $($_.Exception.Message)"`r`n            Write-Log -Type ERROR -Message "[Get-DHCPInfo] Unable to find DHCP. DHCP Server may not be installed. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        }
    }
    finally {
        # Reset progress preference
        $ProgressPreference = $originalProgressPreferen
    }
}

# Get GPO Information 
function Get-GPOInformation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    # Instantiate the hash table
    $ADReportHash["GPO"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting Group Policy Object with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-GPOInformation] START: Getting Group Policy Object with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
    
    $GPOs = Get-GPO -all
    foreach ($gpo in $GPOS) {
        $GPOObject = [PSCustomObject]@{
            Name = $gpo.DisplayName
            ID = $gpo.Id
            Status = $gpo.GpoStatus
            CreationTime = $gpo.CreationTime
            ModificationTime = $gpo.ModificationTime
            Owner = $gpo.Owner
            DomainName = $gpo.DomainName
            UserVersion = ("AD Version: $($gpo.User.DSversion) , SysVol Version: $($gpo.User.sysvolversion)")  
            ComputerVersion = ("AD Version: $($gpo.Computer.DSversion) , SysVol Version: $($gpo.Computer.sysvolversion)"             )
        }
        $ADReportHash["GPO"][$gpo.DisplayName] = $GPOObject
    }
    Write-Log -Type INFO -Message "[Get-GPOInformation] Getting All Group Policy Objects" -ExportFileLocation $ExportDetails[0]

    # Completed Details - Output and Log    
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-GPOInformation] COMPLETED: Gathering Group Policy Object in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

#This function will retrieve information about the Sites and Services of the Active Directory
function Get-ADSiteInventory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    # Instantiate the hash table
    $ADReportHash["SiteInventory"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting Active Directory Site Inventory with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-ADSiteInventory] START: Getting Active Directory Site Inventory with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
    
    Try {
        $Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
        $SiteInfo = $Forest.Sites
        Write-Log -Type INFO -Message "[Get-ADSiteInventory] Getting All Sites Objects" -ExportFileLocation $ExportDetails[0]

        $ForestType = [System.DirectoryServices.ActiveDirectory.DirectoryContexttype]"forest"
        $ForestContext = New-Object -TypeName System.DirectoryServices.ActiveDirectory.DirectoryContext -ArgumentList $ForestType, $Forest
        Write-Log -Type INFO -Message "[Get-ADSiteInventory] Getting All Forests" -ExportFileLocation $ExportDetails[0]

        $Configuration = ([ADSI]"LDAP://RootDSE").configurationNamingContext
        $SubnetsContainer = [ADSI]"LDAP://CN=Subnets,CN=Sites,$Configuration"

        # Gather Site Link Details
        $progresscounter = 0
        $totalCount = ($SiteInfo | measure).count
        foreach ($site in $SiteInfo) {
            $progresscounter++
            Write-Log -Type DEBUG -Message "[Get-ADSiteInventory] Gathering Site Details for $($Site.Name): $($progresscounter2)/$($totalCount2)" -ExportFileLocation $ExportDetails[0]
            $LinksInfo = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::FindByName($ForestContext, $($site.name))).SiteLinks
            $progresscounter2 = 0
            $totalCount2 = ($LinksInfo | measure).count
            foreach ($link in $LinksInfo) {
                #Write-Host "Processing Site Link: $($link.Name) in Site: $($site.Name)"
                $progresscounter2++
                Write-Log -Type DEBUG -Message "[Get-ADSiteInventory] Gathering Site Link Details for '$($link.Name)' in Site '$($site.Name)': $($progresscounter2)/$($totalCount2)" -ExportFileLocation $ExportDetails[0]
                $siteLinkName = $site.Name + "-" + $link.name
                $siteObject = [PSCustomObject]@{
                    Name                         = $site.Name
                    SiteLink                     = $link.Name
                    Servers                      = $site.Servers -join ","
                    Domains                      = $site.Domains -join ","
                    Options                      = $site.options
                    AdjacentSites                = $site.AdjacentSites -join ','
                    InterSiteTopologyGenerator   = $site.InterSiteTopologyGenerator
                    Location                     = $site.location
                    Subnets                      = ($site.Subnets.name | ForEach-Object {
                        $SubnetAdditionalInfo = $SubnetsContainer.Children | Where-Object {$_.name -like "*$_*"}
                        "$_ -- $($SubnetAdditionalInfo.Description)"
                    }) -join ","
                    SiteLinkCost                 = $link.Cost
                    ReplicationInterval          = $link.ReplicationInterval
                    ReciprocalReplicationEnabled = $link.ReciprocalReplicationEnabled
                    NotificationEnabled          = $link.NotificationEnabled
                    TransportType                = $link.TransportType
                    InterSiteReplicationSchedule = $link.InterSiteReplicationSchedule
                    DataCompressionEnabled       = $link.DataCompressionEnabled
                }
                $ADReportHash["SiteInventory"][($siteLinkName)] = $siteObject
            }
        }
    }
    Catch {
        Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in gathering Site Inventory. $($_.Exception.Message)"`r`n        Write-Log -Type Error -Message "[Get-ADSiteInventory] An error occurred in gathering Site Inventory. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
    finally {
        # Completed Details - Output and Log
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
        Write-Log -Type INFO -Message "[Get-ADSiteInventory] COMPLETED: Gathering Active Directory Site Inventory in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
    }
}

#Perform DC Diag on all Domain Controllers
function Invoke-DcDiagCheck {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$False,HelpMessage='Provide the parent directory where you want to save the output files')]
        [string]$parentOutputDirectory
    )
    # Instantiate the hash table
    $ADReportHash["DCDiag"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Performing DC Diag Check with $($detailLevel) details ..." -ForegroundColor Cyan
    Write-Log -Type INFO -Message "[Invoke-DcDiagCheck] START: Getting Active Directory Site Inventory with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    $DCDiagTests = @(
        "DNS"
        "Netlogons"
        "Replications"
        "Services"
        "Advertising"
        "CheckSecurityError"
        "CutoffServers"
        "FrsEvent"
        "DFSREvent"
        "SysVolCheck"
        "KccEvent"
        "KnowsOfRoleHolders"
        "MachineAccount"
        "NCSecDesc"
        "NetLogons"
        "ObjectsReplicated"
        "OutboundSecureChannels"
        "Replications"
        "RidManager"
        "Services"
        "SystemLog"
        "Topology"
        "VerifyEnterpriseReferences"
        "VerifyReferences"
        "VerifyReplicas"
        "CheckSDRefDom"
        "CrossRefValidation"
        "LocatorCheck"
        "Intersite"
    )
    $AllDomainControllers = Get-ADDomainController -Filter *
    Write-Log -Type INFO -Message "[Invoke-DcDiagCheck] Getting All Domain Controllers" -ExportFileLocation $ExportDetails[0]
    

    # Prompt for the parent output directory
    if (!($parentOutputDirectory)) {
        $parentOutputDirectory = Read-Host "Enter the parent directory where you want to save the output files"
    }

    $progresscounter = 0
    $TotalCount = ($AllDomainControllers | Measure-Object).Count
    foreach ($DC in $AllDomainControllers) {
        #DC Name
        $dcName = $DC.Name
        $DCDiagFolderName = $DcName + "_DCDiag"
        # Creates a hash of the DC Results for each DC
        $dcResult = @{
            ConnectionError = $null;
            CommandStatus   = $null;
            Tests           = @{}
            TestSummary     = @()
        }
        Write-Host
        Write-Host "Processing DC: $dcName..." -ForegroundColor Cyan
        $progresscounter++
        Write-ProgressHelper -Id 1 -Total $TotalCount -Index $progresscounter -Activity "Running DC Diag for DC: $dcName"
        Write-Log -Type DEBUG -Message "[Invoke-DcDiagCheck] Running DC Diag for DC: $dcName. $($progresscounter)/$($TotalCount)" -ExportFileLocation $ExportDetails[0]

        # Create a folder for each domain controller
        $dcOutputDirectory = Join-Path -Path $parentOutputDirectory -ChildPath $DCDiagFolderName
        if (-not (Test-Path -Path $dcOutputDirectory)) {
            try {
                $folderResult = New-Item -Path $dcOutputDirectory -ItemType Directory -Force
                Write-Log -Type DEBUG -Message "[Invoke-DcDiagCheck] Created DC Diag Folder for $($dcName)."
                Write-Output "Created DC Diag Folder for '$($dcName)' at '$($dcOutputDirectory)'."
            }
            catch {
                <#Do this if a terminating exception happens#>
            }
        }

        try {
            #Checks if services are running the DC
            $AllServices = Get-WMIObject Win32_Service -ComputerName $dcName

            # Run DCDiag Tests - Loop through all Tests
            $DCtestResults = @()

            #Progress Bar 2
            $start2 = Get-Date
            $progresscounter2 = 0
            $TotalCount2 = $DCDiagTests.Count
            foreach ($testName in $DCDiagTests) {
                $progressCounter2++
                Write-ProgressHelper -Id 2 -Total $TotalCount2 -Index $progresscounter2 -Activity "Running Test: $testName"
 
                # Run the test
                Write-Host "Running Test: $testName  "  -NoNewLine # Debugging output
                Write-Log -Type DEBUG -Message "[Invoke-DcDiagCheck] Running Test: $testName. $($progresscounter2)/$($TotalCount2)" -ExportFileLocation $ExportDetails[0]
                $ThisTestToRun = $testName

                $DCDiagResult = dcdiag /test:$ThisTestToRun /v -s:$dcName
                Write-Log -Type DEBUG -Message "[Invoke-DcDiagCheck] Test: $testName Completed. $($progresscounter2)/$($TotalCount2)" -ExportFileLocation $ExportDetails[0]

                # Check if the specific success string is present in the output
                if ($DCDiagResult -match "$DCname passed test $ThisTestToRun") {
                    $testResult = 'Passed'
                    Write-Host "Passed" -foregroundcolor Green
                    Write-Log -Type DEBUG -Message "[Invoke-DcDiagCheck] Test: $testName Passed. $($progresscounter2)/$($TotalCount2)" -ExportFileLocation $ExportDetails[0]
                } else {
                # Add the DC results to the hash
                    $testResult = 'Failed'
                    Write-Host "Failed" -foregroundcolor Red
                    Write-Log -Type DEBUG -Message "[Invoke-DcDiagCheck] Test: $testName Failed. $($progresscounter2)/$($TotalCount2)" -ExportFileLocation $ExportDetails[0]
                }

                # Create the current test object
                $currentTest = New-Object PSObject -Property ([ordered]@{
                    "DCName" = $dcName
                    "TestName" = $testname
                    "TestResult" = $testresult
                })

                #Add current test to DC's list of all test results
                $DCtestResults += $currentTest
               
                # Export test results to text files in the domain controller's folder
                $outputFileName = "$testName.txt"
                $outputFilePath = Join-Path -Path $dcOutputDirectory -ChildPath $outputFileName
                $DCDiagResult | Out-File -FilePath $outputFilePath -Force
                Write-Log -Type DEBUG -Message "[Invoke-DcDiagCheck] Exported Test Results to $outputFilePath. $($progresscounter2)/$($TotalCount2)" -ExportFileLocation $ExportDetails[0]

            }
            # Add the DC results to the hash
            $adReportHash["DCDiag"][$DcName] = $DCtestResults

            Write-ProgressHelper -Id 2 -Total $TotalCount2 -Activity "Running Test: $testName" -Completed
        }
        Catch {
            Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in running DC Diagnostics for $($dcName). $($_.Exception.Message)"`r`n            Write-Log -Type Error -Message "[Invoke-DcDiagCheck] An error occurred in running DC Diagnostics for $($dcName). $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
        }
    }

    # Completed Details - Output and Log
    Write-ProgressHelper -Id 1 -Total $TotalCount -Activity "Running DC Diag for DC: $dcName" -Completed
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Invoke-DcDiagCheck] COMPLETED: Gathering DC Diagnostics Info in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]

}

#endregion - Active Directory Specific Functions
# ----------------------------------
#region - Active Directory Specific Functions - Objects
# ----------------------------------

# Get Privleged Accounts Information
function Get-AllADAdmins {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    function Add-AdminGroupMembers {
        param (
            [string]$GroupName,
            [string]$LogMessage,
            [string]$ExportFileLocation
        )

        try {
            $members = Get-ADGroupMember -Identity $GroupName -ea SilentlyContinue | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name "Group" -Value $GroupName -PassThru -Force
            }
            Write-Log -Type INFO -Message $LogMessage -ExportFileLocation $ExportFileLocation
            return $members
        } catch {
            if ($groupName -like 'Exchange Servers' -or $groupName -like 'Organization Management') {
                Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "Failed to get members of $($GroupName). Exchange Server may not be installed: $($_.Exception.Message)"
                Write-Log -Type ERROR -Message "[Get-AllADAdmins] Failed to get members of '$($GroupName)'. Exchange Server may not be installed. $($_.Exception.Message)" -ExportFileLocation $ExportFileLocation
            }
            else {
                Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "Failed to get members of '$($GroupName)': $($_.Exception.Message)"
                Write-Log -Type ERROR -Message "[Get-AllADAdmins] An error occurred in gathering '$($GroupName)': $($_.Exception.Message)" -ExportFileLocation $ExportFileLocation
            }
            return @()
        }
    }

    # Instantiate the hash table
    $ADReportHash["ADAdmins"] = @{}
    $ADAdmins = @()

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting Admin Accounts with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllADAdmins] START: Getting Admin Accounts with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
    
    # Admin groups to process
    $adminGroups = @(
        @{ Name = 'Domain Admins'; LogMessage = '[Get-AllADAdmins] Getting all Domain Admins' },
        @{ Name = 'Enterprise Admins'; LogMessage = '[Get-AllADAdmins] Getting all Enterprise Admins' },
        @{ Name = 'Schema Admins'; LogMessage = '[Get-AllADAdmins] Getting all Schema Admins' },
        @{ Name = 'Organization Management'; LogMessage = '[Get-AllADAdmins] Getting all Organization Management Users' },
        @{ Name = 'Exchange Servers'; LogMessage = '[Get-AllADAdmins] Getting all Exchange Servers' }
    )

    # Gather Special Admin Groups - Domain Admins, Enterprise Admins, Schema Admins, Org Management, Exchange Servers
    foreach ($group in $adminGroups) {
        $ADAdmins += Add-AdminGroupMembers -GroupName $group.Name -LogMessage $group.LogMessage -ExportFileLocation $ExportDetails[0]
    }

    # Gather Special Admin Groups -  Never Expire
    $ADAdmins += Get-ADUser -Filter {PasswordNeverExpires -eq $true}  | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name "Group" -Value 'Never Expire' -PassThru -Force
    }
    Write-Log -Type INFO -Message "[Get-AllADAdmins] Getting all NeverExpire Users" -ExportFileLocation $ExportDetails[0]

    #Combine All AD ADmins
    $progresscounter = 0
    $totalCount = ($ADAdmins | measure).count
    foreach ($user in $ADAdmins) {
        $progresscounter++
        Write-Log -Type DEBUG -Message "[Get-AllADAdmins] Processing user $($progresscounter)/$($totalCount): '$($user.Name)' in the group '$($user.Group)' for Active Directory Admin Group analysis." -ExportFileLocation $ExportDetails[0]
        $ADReportHash["ADAdmins"][$user.SamAccountName] = $user
    }

    # Completed Details - Output and Log
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AllADAdmins] COMPLETED: Gathering Admin Accounts in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

# Retrieve All AD Objects - Users
function Get-AllADUsers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    # Instantiate the hash table
    $ADReportHash["ADUsers"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting All User Accounts with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllADUsers] START: Getting All User Accounts with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
    
    # Get Password Expiry Info
    function Get-PasswordExpiryInfo {
        param (
            [Parameter(Mandatory=$True)]
            [string]$sAMAccountName
        )
        $PSO= Get-ADUserResultantPasswordPolicy -Identity $sAMAccountName  # store ad user's current password policy
        $defaultDomainPolicy = Get-ADDefaultDomainPasswordPolicy   # get and store the domain's current default password policy        
        $passwordexpirydefaultdomainpolicy = $defaultDomainPolicy.MaxPasswordAge.Days -ne 0  # make sure the default policy is anything but zero        
        # Collect and store the default domain policy, in case no PSO was applied to the user
        if($passwordexpirydefaultdomainpolicy)            
        {            
            $defaultdomainpolicyMaxPasswordAge = $defaultDomainPolicy.MaxPasswordAge.Days  # set max password age      
        }  
        # if resulting password policy is anything but null
        if ($PSO) {  
            $PSOMaxPasswordAge = $PSO.MaxPasswordAge.days  # store password age
            $pwdlastset = [datetime]::FromFileTime((Get-ADUser -LDAPFilter "(&(samaccountname=$sAMAccountName))" -properties pwdLastSet).pwdLastSet) # get password last set time
            $expirydate = ($pwdlastset).AddDays($PSOMaxPasswordAge) # get expiry date    
            $delta = ($expirydate - (Get-Date)).Days  # store the delta between today and the expiry date to find how many days are remaining
            $ResultingPasswordObject = New-object -TypeName PSObject
            $ResultingPasswordObject | Add-Member -MemberType NoteProperty -Name MaxPasswordAge -Value $PSOMaxPasswordAge -Force
            $ResultingPasswordObject | Add-Member -MemberType NoteProperty -Name PwdExpiryDate -Value $pwdlastset -Force
            $ResultingPasswordObject | Add-Member -MemberType NoteProperty -Name PwdExpiryDate -Value $expirydate -Force
            $ResultingPasswordObject | Add-Member -MemberType NoteProperty -Name PwdsRemaining -Value $delta -Force
        }        
        # if the password policy was not defined, use the default policy    
        else {
            if($passwordexpirydefaultdomainpolicy) {            
                $pwdlastset = [datetime]::FromFileTime((Get-ADUser -LDAPFilter "(&(samaccountname=$samaccountname))" -properties pwdLastSet).pwdLastSet)            
                $expirydate = ($pwdlastset).AddDays($defaultdomainpolicyMaxPasswordAge)            
                $delta = ($expirydate - (Get-Date)).Days      
                $ResultingPasswordObject = New-object -TypeName PSObject
                $ResultingPasswordObject | Add-Member -MemberType NoteProperty -Name MaxPasswordAge -Value $defaultdomainpolicyMaxPasswordAge -Force
                $ResultingPasswordObject | Add-Member -MemberType NoteProperty -Name PwdExpiryDate -Value $pwdlastset -Force
                $ResultingPasswordObject | Add-Member -MemberType NoteProperty -Name PwdExpiryDate -Value $expirydate -Force
                $ResultingPasswordObject | Add-Member -MemberType NoteProperty -Name PwdsRemaining -Value $delta -Force
                return $ResultingPasswordObject
            }            
        }  
        return $ResultingPasswordObject
    }
    
    # This command could take some time, depending on the size of the AD
    
    switch ($detailLevel) {
        {$_ -in "minimum", "combined","all"} { 
            $DesiredProperties = @(
                "DisplayName", "AccountStatus", "PasswordChangeAtNextLogOn", "MaxPasswordAge"
                "ExpiryDate", "DaysRemaining"
                "AccountExpirationDate", "accountExpires", "AccountLockoutTime"
                "AccountNotDelegated", "BadLogonCount", "badPasswordTime", "badPwdCount"
                "CannotChangePassword", "CanonicalName", "City", "CN"
                "Company", "Country", "countryCode"  
                "Deleted", "Department", "Description", "DistinguishedName", "EmailAddress"
                "EmployeeID", "EmployeeNumber", "Enabled", "Fax", "GivenName"
                "HomeDirectory", "HomeDrive", "HomePhone", "Initials",
                "isDeleted", "LastBadPasswordAttempt", "LastKnownParent", "LastLogoff"
                "LastLogon", "LastLogonDate", "LastLogonTimestamp", "LockedOut", "logonCount"
                "logonHours", "logonWorkstation", "Manager", "MemberOf", "MiddleName"
                "MobilePhone", "Modified", "Name", "Notes", "Office"
                "OfficePhone", "Organization", "OtherName", "PasswordExpired", "PasswordLastSet"
                "PasswordNeverExpires", "PasswordNotRequired", "POBox", "PostalCode", "PrimaryGroup"
                "proxyAddresses", "ProfilePath", "ProtectedFromAccidentalDeletion", "pwdLastSet", "SamAccountName"
                "SamAccountType", "ServicePrincipalNames", "SID"
                "SIDHistory", "SmartcardLogonRequired", "sn", "State", "StreetAddress"
                "Surname", "Title", "TrustedForDelegation", "UserPrincipalName"
                "whenChanged", "whenCreated", "ObjectCategory", "ObjectGUID"
            )
            $AllADUsers = Get-ADUser -Filter * -Properties * | Select $DesiredProperties
        }
        {$_ -in "geek"} {
            $AllADUsers = Get-ADUser -Filter * -Properties *
        }
    }

    $AllADUsers = Get-ADUser -Filter * -Properties *
    Write-Log -Type INFO -Message "[Get-AllADUsers] Gathered All Users with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    $progressCounter = 0
    $totalCount = ($AllADUsers | measure).count
    foreach ($User in $AllADUsers) {
        $progressCounter++
        Write-ProgressHelper -Total $totalCount -Index $progresscounter -Activity "Gathering All Active Directory Users" -Operation "Gathering User Details for $($User.SamAccountName)"

        # Get the user's manager
        $managerDisplayName = if ($User.Manager) { (Get-ADUser -Identity $User.Manager).DisplayName } else { $null }
        $user | Add-Member -MemberType NoteProperty -Name Manager -Value $managerDisplayName -Force
        Write-Log -Type DEBUG -Message "[Get-AllADUsers] Gathered Manager Details for $($user.SamAccountName) : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]

        # Determine account status
        $accountStatus = if ($User.Enabled) { 'Enabled' } else { 'Disabled' }
        $user | Add-Member -MemberType NoteProperty -Name AccountStatus -Value $accountStatus -Force
        Write-Log -Type DEBUG -Message "[Get-AllADUsers] Gathered Account Status for $($user.SamAccountName) : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]

        # Proxy Addresses Expansion
        $user | Add-Member -MemberType NoteProperty -Name ProxyAddresses -Value ($user.ProxyAddresses -join ";") -Force
        Write-Log -Type DEBUG -Message "[Get-AllADUsers] Gathered Proxy Addresses for $($user.SamAccountName) : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]

        # Service Prinicpal names Expansion
        $user | Add-Member -MemberType NoteProperty -Name ServicePrincipalNames -Value ($user.ServicePrincipalNames -join ";") -Force
        Write-Log -Type DEBUG -Message "[Get-AllADUsers] Gathered Proxy Addresses for $($user.SamAccountName) : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]


        # Check if password change is required at next logon
        $passwordChangeAtNextLogOn = if ($User.pwdLastSet -eq 0) { 'TRUE' } else { 'FALSE' }
        $user | Add-Member -MemberType NoteProperty -Name PasswordChangeAtNextLogOn -Value $passwordChangeAtNextLogOn -Force
        Write-Log -Type DEBUG -Message "[Get-AllADUsers] Gathered Password Change Info for $($user.SamAccountName) : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]
 
        #Check Password Expiry Info
        $passwordExpiryInfo = Get-PasswordExpiryInfo -sAMAccountName $user.SamAccountName
        $user | Add-Member -MemberType NoteProperty -Name MaxPasswordAge -Value $passwordExpiryInfo.MaxPasswordAge -Force
        $user | Add-Member -MemberType NoteProperty -Name pwdLastSet -Value $passwordExpiryInfo.pwdLastSet -Force
        $user | Add-Member -MemberType NoteProperty -Name ExpiryDate -Value $passwordExpiryInfo.ExpiryDate -Force
        $user | Add-Member -MemberType NoteProperty -Name DaysRemaining -Value $passwordExpiryInfo.DaysRemaining -Force
        Write-Log -Type DEBUG -Message "[Get-AllADUsers] Gathered Password Expiration Info for $($user.SamAccountName) : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]
 
        # Get group memberships
        $groupMemberships = (Get-ADPrincipalGroupMembership -Identity $User.SamAccountName -ea SilentlyContinue | Sort-Object | Select-Object -ExpandProperty Name) -join ', ' 
        $user | Add-Member -MemberType NoteProperty -Name GroupMemberships -Value $groupMemberships -Force
        Write-Log -Type DEBUG -Message "[Get-AllADUsers] Gathered Group Membership for $($user.SamAccountName) : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]

        # Add user to the hash table
        $ADReportHash["ADUsers"][$User.SamAccountName] = $user
    }

    # Completed Details - Output and Log
    Write-ProgressHelper -Total $totalCount -Activity "Gathering All Active Directory Users" -Completed
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AllADUsers] COMPLETED: Gathering All User Accounts in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

# Retrieve All AD Objects - Groups
function Get-AllADGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    # Instantiate the hash table
    $ADReportHash["ADGroups"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting All AD Groups with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllADGroups] START: Getting All AD Groups with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]
    
    function Test-PrivilegedGroup {
        param(
            [Parameter(Mandatory=$True)]
            [string]$domain,
            [Parameter(Mandatory=$True)]
            [string]$groupSID
        )
        $DomainObject = Get-ADDomain $domain
        $DomainSID = $DomainObject.DomainSID
        $PrivilegedGroups = "$($DomainSid)-512", "$($DomainSid)-518",
                            "$($DomainSid)-519", "$($DomainSid)-520",
                            "S-1-5-32-544", "S-1-5-32-548", "S-1-5-32-549",
                            "S-1-5-32-550", "S-1-5-32-551", "S-1-5-32-552",
                            "S-1-5-32-556", "S-1-5-32-557", "S-1-5-32-573",
                            "S-1-5-32-578", "S-1-5-32-580"
        $IsPrivileged = "Standard"  
        ForEach($PrivilegedGroup in $PrivilegedGroups) {
            if ($PrivilegedGroup -eq $groupSID)
            {
                $IsPrivileged = "Privileged"
            }
        }
        return $IsPrivileged
    }
    
    # Get all Group data with all extended properties
    $AllADGroups = Get-ADGroup -Filter *
    Write-Log -Type INFO -Message "[Get-AllADGroups] Gathered All AD Groups." -ExportFileLocation $ExportDetails[0]

    $progressCounter = 0
    $totalCount = ($AllADGroups | measure).count
    foreach ($group in $AllADGroups) {
        $progressCounter++
        Write-ProgressHelper -Total $totalCount -Index $progresscounter -Activity "Gathering All Active Directory Groups" -Operation "Gathering Group Details for '$($group.Name)'"
        # Test Group for Privileged
        $groupSID = $group.SID.ToString()
        $domain = (Get-ADDomain).forest
        $group | Add-Member -MemberType NoteProperty -Name Privileged -Value (Test-PrivilegedGroup -domain $domain -groupSID $groupSID) -Force
        Write-Log -Type DEBUG -Message "[Get-AllADGroups] Test Group for Privileged Permisssions for '$($group.Name)' : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]

        # Add Member Count
        try {
            $GroupMemberCount = (Get-ADGroupMember -Identity $group.DistinguishedName -ea Stop | measure).Count
        }
        catch {
            if ($_.Exception.Message -eq "The size limit for this request was exceeded") {
                $GroupMemberCount = "ExceededLimit"
            }
            else {$GroupMemberCount = "Uknown"}
        }
        $group | Add-Member -MemberType NoteProperty -Name MemberCount -Value $GroupMemberCount -Force
        Write-Log -Type DEBUG -Message "[Get-AllADGroups] Gathered Group Membership for '$($group.Name)' : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]
        
        # Add Restricted Access
        if ($user.Enabled -eq 'TRUE') {$RestrictedAccess = 'Enabled'} Else {$RestrictedAccess = 'Disabled'}
        $group | Add-Member -MemberType NoteProperty -Name RestrictedAccess -Value $RestrictedAccess -Force # the 'if statement replaces $_.Enabled output with a user-friendly readout
        Write-Log -Type DEBUG -Message "[Get-AllADGroups] Gathered Restricted Access for '$($group.Name)' : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]

        $ADReportHash["ADGroups"][$group.Name] = $group
    }

    # Completed Details - Output and Log
    Write-ProgressHelper -Total $totalCount -Activity "Gathering All Active Directory Groups" -Completed
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AllADGroups] COMPLETED: Gathering All User Accounts in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

function Get-AllADComputers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    # Instantiate the hash table
    $ADReportHash["ADComputers"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting All AD Computers with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllADComputers] START: Getting All AD Computers with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    # Get all Computer data with all extended properties
    $AllADComputers = Get-ADComputer -Filter *
    Write-Log -Type INFO -Message "[Get-AllADComputers] Gathered All AD Computers." -ExportFileLocation $ExportDetails[0]

    $progressCounter = 0
    $totalCount = ($AllADComputers | measure).count
    foreach ($computer in $AllADComputers) {
        $progressCounter++
        Write-ProgressHelper -Total $totalCount -Index $progresscounter -Activity "Gathering All Active Directory Computers" -Operation "Gathering Computers Details for $($group.Name)"
        Write-Log -Type DEBUG -Message "[Get-AllADComputers] Gather Computer Details for $($computer.Name) : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]

        # Get Install Date
        $computer | Add-Member -MemberType NoteProperty -Name ManufacturerInstallDate -Value ({ForEach-Object{(([WMI] "").ConvertToDateTime((Get-WmiObject Win32_OperatingSystem -ComputerName $_.Name).InstallDate))}}) -Force
        $ADReportHash["ADComputers"][$computer.Name] = $computer
    }
    # Completed Details - Output and Log
    Write-ProgressHelper -Total $totalCount -Activity "Gathering All Active Directory Computers" -Completed
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AllADComputers] COMPLETED: Gathering All Computers in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]
}

# Retrieve All AD Objects - Organizational Units
function Get-AllADOUs {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'combined', 'all', 'geek')]
        [string]$detailLevel     
    )

    # Instantiate the hash table
    $ADReportHash["ADOUs"] = @{}

    #Start - Start Time, Screen Output, and Log
    $start = Get-Date
    Write-Host "Getting All AD Organizational Units with $($detailLevel) details ..." -ForegroundColor Cyan -nonewline
    Write-Log -Type INFO -Message "[Get-AllADOUs] START: Getting All AD Organizational Units with $($detailLevel) details" -ExportFileLocation $ExportDetails[0]

    #Export all Organizational Unit Information
    $ADOUs = Get-ADOrganizationalUnit -Properties * -Filter *
    Write-Log -Type INFO -Message "[Get-AllADOUs] Gathered All AD Organizational Units." -ExportFileLocation $ExportDetails[0]

    $progressCounter = 0
    $totalCount = ($ADOUs | measure).count
    foreach ($OU in $ADOUs) {
        $progressCounter++
        Write-ProgressHelper -Total $totalCount -Index $progresscounter -Activity "Gathering All Active Directory OrganizationalUnits" -Operation "Gathering OU Details for $($OU.Name)"
        $OU | Add-Member -MemberType NoteProperty -Name "LinkedGroupPolicyObjects" -Value ($OU.LinkedGroupPolicyObjects -join "; ") -Force
        $OU | Add-Member -MemberType NoteProperty -Name "LinkedGroupPolicyObjectCount" -Value (($OU.LinkedGroupPolicyObjects | measure).count) -Force
        $OU | Add-Member -MemberType NoteProperty -Name "OU" -Value ($OU.OU -join "; ") -Force
        $GPOInheritance = Get-GPInheritance -target $OU.DistinguishedName
        $OU | Add-Member -MemberType NoteProperty -Name "Inheritance Blocked" -Value $GPOInheritance.GpoInheritanceBlocked -Force
        $OU | Add-Member -MemberType NoteProperty -Name "Linked GPOs" -Value ($GPOInheritance.GPOLinks.DisplayName -join ";") -Force
        $OU | Add-Member -MemberType NoteProperty -Name "Linked GPO Count" -Value (($GPOInheritance.GPOLinks |measure).count) -Force
        $OU | Add-Member -MemberType NoteProperty -Name "Inherited GPO Links" -Value ($GPOInheritance.InheritedGpoLinks.DisplayName -join "; ") -Force
        $OU | Add-Member -MemberType NoteProperty -Name "Inherited GPO Count" -Value (($GPOInheritance.InheritedGpoLinks |measure).count) -Force
        Write-Log -Type DEBUG -Message "[Get-AllADOUs] Gather Organization Unit Details for $($OU.Name) : $($progresscounter)/$($totalCount)" -ExportFileLocation $ExportDetails[0]
        $ADReportHash["ADOUs"][$OU.Name] = $OU
    }
    # Completed Details - Output and Log
    Write-ProgressHelper -Total $totalCount -Activity "Gathering All Active Directory OrganizationalUnits" -Completed
    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-Host "Completed in $($CompletedTime)" -ForegroundColor Green
    Write-Log -Type INFO -Message "[Get-AllADOUs] COMPLETED: Gathering All Organizational Units in $($CompletedTime)" -ExportFileLocation $ExportDetails[0]

}

#endregion - Active Directory Specific Functions - Objects

# ----------------------------------
#region - Global Variables
# ----------------------------------

# The AD Module must be imported before AD-specific commands can be run
Import-Module ActiveDirectory
Install-ImportExcelModule | Out-Null

$global:forestInfo = Get-ADForest
$global:domainInfo = Get-ADDomain
$global:PDCEmulator = (Get-ADDomain).PDCEmulator
$global:ADsiteLinks = Get-ADReplicationSiteLink -Filter *
$global:ActiveDirectory = $global:domainInfo.dnsroot
$global:DNSRoot = $global:ActiveDirectory

#Active Directory Stats Initialization
$ADReportHash = @{}

#Get Export Path

$ExportDetails = Get-ExportPath
$directory = [System.IO.Path]::GetDirectoryName($ExportDetails)

$resultSize = 'unlimited'
$global:AllDiscoveryErrors = New-Object System.Collections.Generic.List[pscustomobject]

#Global Start Time for Script
$global:InitialStart = Get-Date

# Prompt for level of detail
$reportingMode = Set-ReportMode

#endregion - Global Variables

#endregion - Intial Variables and Functions
########################################################


#start of script
########################################################

#region - Active Directory Discovery
Write-Host "Gathering Active Directory Details" -ForegroundColor Black -BackgroundColor Yellow
Get-ADSiteInventory -detailLevel $reportingMode
Get-ADFSMORoles -detailLevel $reportingMode
Get-AllADForests -detailLevel $reportingMode
Get-AllDomainControllers -detailLevel $reportingMode
Get-DNSInfo -detailLevel $reportingMode
Get-DHCPInfo -detailLevel $reportingMode
Get-GPOInformation -detailLevel $reportingMode

Write-Host 
Write-Host "Running DC Diagnostic Check" -ForegroundColor Black -BackgroundColor Yellow
Invoke-DcDiagCheck -parentOutputDirectory $directory

#endregion - Active Directory Discovery

#region - Active Directory Objects
Write-Host "Gathering Active Directory Objects" -ForegroundColor Black -BackgroundColor Yellow
Get-AllADAdmins -detailLevel $reportingMode
Get-AllADComputers -detailLevel $reportingMode
Get-AllADGroups -detailLevel $reportingMode
Get-AllADUsers -detailLevel $reportingMode
Get-AllADOUs -detailLevel $reportingMode

#endregion - end of script
########################################################
### Export Reports ###
########################################################

#Exclude specific reports from Export
# Common tables to remove for both 'combined' and 'minimum' detail levels
$commonTablesToRemove = @(

)
#Exclude specific reports from Export
switch ($reportingMode) {
    "combined" {
        # Additional tables for combined detail level
        $additionalTables = @()
        $tablesToRemove = $commonTablesToRemove + $additionalTables
    }
    "minimum" {
        # Additional tables for minimum detail level
        $additionalTables = @('UserMailboxes')
        $tablesToRemove = $commonTablesToRemove + $additionalTables
    }
    default {
        # Default case to handle other detail levels where nothing should be removed
        $additionalTables = @()
        $tablesToRemove = $commonTablesToRemove + $additionalTables
    }
}
$ExportADStatsHash = @{}
$ExportADStatsHash = $ADReportHash

#Exclude specific reports from Export
foreach ($key in $tablesToRemove) {
    if ($ExportADStatsHash[$key]) {
        $ExportADStatsHash.Remove($key)
    }
}

#Export Reports: Exports each individual hashtable to own CSV file and then combines into Excel file
Write-Host
Write-Host "Exporting the Active Directory Statistics" -ForegroundColor Black -BackgroundColor Yellow
Write-Log -Type INFO -Message "Exporting the Active Directory Statistics to $($ExportDetails[0])." -ExportFileLocation $ExportDetails[0]
try {
    Export-HashTables -hashtable $ExportADStatsHash -ExportDetails $ExportDetails
}
catch {
    Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Exporting the Active Directory Statistics to $($ExportDetails[0]). Please re-run the script and verify the location is valid and the file is not open in another application. $($_.Exception.Message)"`r`n    Write-Log -Type ERROR -Message "An error occurred in Exporting the Active Directory Statistics to $($ExportDetails[0]). Please re-run the script and verify the location is valid and the file is not open in another application. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
}

Write-Host ""

#Export Errors
try {
    #silence error reporting locations
    Export-ErrorReports -ExportFileLocation $ExportDetails[0] -ErrorData $global:AllDiscoveryErrors -logReportDirectory $ExportDetails[0]

    #Display Error Details
    Write-Host "Error Reporting Details" -ForegroundColor Black -BackgroundColor Yellow
    Write-Host "Check '$($errorReportFolderDirectory)' for error logs " -ForegroundColor Cyan
    Write-Host "$($global:AllDiscoveryErrors.count) " -ForegroundColor Red -NoNewline
    Write-Host "Error(s) encountered. "
    }
catch {
    if ($_.Exception.Message -like '*because it is an empty collection*') {
        Write-Warning "No Errors Found!"
    }
    else {
        Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "An error occurred in Exporting the Error Reports. Inndividual CSV files can be found in local temp folder. $($_.Exception.Message)"`r`n        Write-Log -Type ERROR -Message "An error occurred in Exporting the Error Reports. $($_.Exception.Message)" -ExportFileLocation $ExportDetails[0]
    }
}

Write-Host ""
Write-Host "Object Count Table" -ForegroundColor Black -BackgroundColor Green
$CompletedTime = (((Get-Date) - $global:initialStart).ToString('hh\:mm\:ss'))
Write-Host "COMPLETED: Gathered Active Directory Details. Completed Time: $($CompletedTime)" -ForegroundColor Cyan
Write-Log -Type INFO -Message "COMPLETED: Gathered Active Directory Details. Completed Time: $($CompletedTime)" -ExportFileLocation $ExportDetails[0]


#Final Output of Active Directory Counts
$ADStatsOutput = @()

foreach ($key in $ADReportHash.Keys) {
    $count = $ADReportHash[$key].Count
    # Create a custom object for the current key-value pair
    $object = New-Object -TypeName PSCustomObject -Property @{
        "Key" = $key
        "Count" = $count
    }
    # Add the custom object to the array
    $ADStatsOutput += $object
}
$ADStatsOutput | ft

