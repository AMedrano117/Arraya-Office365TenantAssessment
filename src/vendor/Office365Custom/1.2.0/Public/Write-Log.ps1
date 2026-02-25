<#
.SYNOPSIS
    Logs messages to a file or outputs verbose logs.
.DESCRIPTION
    This function writes log entries (INFO, WARNING, ERROR, DEBUG) to a specified file location or displays them in the host.
.PARAMETER Type
    The type of log message.
.PARAMETER Message
    The log message.
.PARAMETER LogPath
    Optional directory for saving the log file.
.PARAMETER ExportFileLocation
    Alternative path information for detailed logs.
.EXAMPLE
    Write-Log -Type "INFO" -Message "Operation completed successfully" -LogPath "C:\Logs"
#>
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
    # If the log file path is provided, append the log message to the file
    if ($LogPath) {
        # Get the directory, filename without extension, and the extension
        # Create the 'Log Reporting' directory if it doesn't exist
        if (-not (Test-Path $LogPath)) {
            $newfolder = New-Item -Path $LogPath -ItemType Directory
        }
        try {
            # Prepare the log message with a timestamp
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "$timestamp : [$Type] $Message"
            $LogFile = $LogPath + "\FullTenantReportLog.txt"
            $logMessage | Out-File -Append -FilePath $LogFile
        } catch {
            Write-Error "Failed to write to log file at '$LogFile': $_"
        }
    }
    elseif ($ExportFileLocation) {
        # Get the directory, filename without extension, and the extension
        $directory = [System.IO.Path]::GetDirectoryName($ExportFileLocation)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ExportFileLocation)
        $txtFileName = $baseName + "-FullReportLog.txt"
        $NewLogFolder = $baseName + " Reporting"
        $Newdirectory = Join-Path -Path $directory -ChildPath $NewLogFolder
        $LogFile = Join-Path -Path $Newdirectory -ChildPath $txtFileName

        # Create the 'Log Reporting' directory if it doesn't exist
        if (-not (Test-Path $Newdirectory)) {
            $newfolder = New-Item -Path $Newdirectory -ItemType Directory
        }
        try {
            # Prepare the log message with a timestamp
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "$timestamp : [$Type] $Message"
            $logMessage | Out-File -Append -FilePath $LogFile
        } catch {
            Write-Error "Failed to write to log file at '$LogFile': $_"
        }
    }

    # If the Verbose switch is used, also display the log message on the screen
    if ($VerbosePreference -eq 'Continue') {
        Write-Verbose $Message
    }

    if (-not $CaptureError -and $Type -eq 'ERROR') {
        $CaptureError = $true
    }
    if ($CaptureError) {
        Capture-ErrorHelper -ErrorRecordVar $ErrorRecordVar -errorMessage $Message
    }
}