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

    function Resolve-DebugLogDirectory {
        param(
            [Parameter(Mandatory=$true)]
            [string]$BaseDirectory
        )

        if ([string]::IsNullOrWhiteSpace($BaseDirectory)) {
            $BaseDirectory = (Get-Location).Path
        }

        if ([string]::Equals((Split-Path -Path $BaseDirectory -Leaf), 'Debugging', [System.StringComparison]::OrdinalIgnoreCase)) {
            return $BaseDirectory
        }

        if ([string]::Equals((Split-Path -Path $BaseDirectory -Leaf), 'Deliverables', [System.StringComparison]::OrdinalIgnoreCase)) {
            $runRootDirectory = Split-Path -Path $BaseDirectory -Parent
            if (-not [string]::IsNullOrWhiteSpace($runRootDirectory)) {
                return (Join-Path -Path (Join-Path -Path $runRootDirectory -ChildPath 'Support') -ChildPath 'Debugging')
            }
        }

        return (Join-Path -Path $BaseDirectory -ChildPath 'Debugging')
    }

    # If the log file path is provided, append the log message to the file
    if ($LogPath) {
        # Get the directory, filename without extension, and the extension
        # Create the 'Log Reporting' directory if it doesn't exist
        $resolvedLogDirectory = Resolve-DebugLogDirectory -BaseDirectory $LogPath
        if (-not (Test-Path $resolvedLogDirectory)) {
            $newfolder = New-Item -Path $resolvedLogDirectory -ItemType Directory -Force
        }
        try {
            # Prepare the log message with a timestamp
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "$timestamp : [$Type] $Message"
            $LogFile = Join-Path -Path $resolvedLogDirectory -ChildPath 'FullTenantReportLog.txt'
            $logMessage | Out-File -Append -FilePath $LogFile
        } catch {
            Write-Error "Failed to write to log file at '$LogFile': $_"
        }
    }
    elseif ($ExportFileLocation) {
        # Get the directory, filename without extension, and the extension
        $directory = [System.IO.Path]::GetDirectoryName($ExportFileLocation)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ExportFileLocation)
        if ([string]::IsNullOrWhiteSpace($directory)) {
            $directory = (Get-Location).Path
        }
        if ([string]::IsNullOrWhiteSpace($baseName)) {
            $baseName = 'Assessment'
        }

        $directory = Resolve-DebugLogDirectory -BaseDirectory $directory
        $txtFileName = $baseName + "-FullReportLog.txt"
        $LogFile = Join-Path -Path $directory -ChildPath $txtFileName

        if (-not (Test-Path $directory)) {
            $newfolder = New-Item -Path $directory -ItemType Directory -Force
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
        if ($PSBoundParameters.ContainsKey('ErrorRecordVar') -and $null -ne $ErrorRecordVar) {
            Capture-ErrorHelper -ErrorRecordVar $ErrorRecordVar -errorMessage $Message
        }
        elseif ($VerbosePreference -eq 'Continue') {
            Write-Verbose "CaptureError was requested for Write-Log, but no ErrorRecordVar was supplied. Skipping structured error capture."
        }
    }
}
