<#
.SYNOPSIS
Determines the full path and file name for exporting files.
.DESCRIPTION
Prompts the user for a file or directory path, processes the input to determine the export folder and filename,
and defaults to known locations if the input is empty or invalid.
.PARAMETER FileName
The base name for the file (e.g. "Migration Report", "Cutover Log"). Required.
.PARAMETER DefaultExtension
The default file extension (e.g. ".xlsx" or ".csv"). Default is ".xlsx".
.EXAMPLE
PS C:\> Get-ExportPath -FileName "Migration Report" -DefaultExtension ".csv"
.EXAMPLE
PS C:\> Get-ExportPath -FileName "Cutover Log" 
#>
function Get-ExportPath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$FileName,
        [string]$DefaultExtension = ".xlsx"
    )

    # Ask user for Export location
    Write-Host "Gather Export Path and/or File Name" -ForegroundColor Cyan
    $userInput = Read-Host -Prompt "Enter the file path (with .xlsx or .csv extension) or folder path to save the file"

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
        # If the user input is a folder, assign the folder path
        $folderPath = $userInput
    } else {
        # If it's a full path, split the file name and folder path
        $folderPath = Split-Path -Path $userInput -Parent
        $inputFileName = Split-Path -Path $userInput -Leaf
    }

    # If folderPath is empty or invalid, default to current script location
    if ([string]::IsNullOrEmpty($folderPath) -or !(Test-Path $folderPath)) {
        $folderPath = $PSScriptRoot
    }

    # Add timestamp to filename for uniqueness
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

    # Determine the filename to use
    if (-not [string]::IsNullOrEmpty($inputFileName)) {
        # If user provided a file name in the path
        $fileName = $inputFileName
    } else {
        # Use provided FileName with timestamp
        $fileName = "$FileName`_$timestamp"
    }

    # Ensure the file has the correct extension
    $extension = [IO.Path]::GetExtension($fileName)
    if ([string]::IsNullOrEmpty($extension)) {
        $fileName += $DefaultExtension
    }

    # Full path
    $fullPath = Join-Path -Path $folderPath -ChildPath $fileName
    
    Write-Host "The file will be saved to: $fullPath" -ForegroundColor Green
    return $fullPath
}
