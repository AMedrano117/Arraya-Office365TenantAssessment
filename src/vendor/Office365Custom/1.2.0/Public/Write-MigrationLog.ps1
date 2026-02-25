# Track Migration Logging for Domain Cutover
function Write-MigrationLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$tenantName,

        [Parameter(Mandatory=$true)]
        [string]$LogPath,

        [Parameter(Mandatory=$true)]
        [string]$OriginalValue,

        [Parameter(Mandatory=$true)]
        [string]$NewValue,

        [Parameter(Mandatory=$true)]
        [string]$ExecutedCommand,

        [Parameter(Mandatory=$false)]
        [string]$CustomEntry = ""
    )

    # Create the log directory if it doesn't exist
    $logDirectory = Split-Path -Path $LogPath -Parent
    if (-not (Test-Path $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory | Out-Null
    }

    # Prepare the log entry as a custom object
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = [PSCustomObject]@{
        Timestamp       = $timestamp
        tenantName      = $tenantName
        OriginalValue   = $OriginalValue
        NewValue        = $NewValue
        ExecutedCommand = $ExecutedCommand
        CustomEntry     = $CustomEntry
    }

    # Export to Excel
    if (-not (Test-Path $LogPath)) {
        # Export the first entry and create the file if it does not exist
        $logEntry | Export-Excel -Path $LogPath -AutoSize -TableName "MigrationLog" -TableStyle Medium9
    } else {
        # Append to the existing Excel file
        $logEntry | Export-Excel -Path $LogPath -AutoSize -Append -TableName "MigrationLog" -TableStyle Medium9
    }

    # Output to console if running in a verbose mode
    Write-Verbose "Logged change: `"$ExecutedCommand`" applied from `"$OriginalValue`" to `"$NewValue`""
}