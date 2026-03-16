function Export-ArrayaErrorReports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$BaseName = 'ErrorReport',
        [Parameter(Mandatory = $false)]
        [string]$ExportFileLocation,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$ErrorData,
        [Parameter(Mandatory = $false)]
        [string]$ErrorReportFolderName,
        [Parameter(Mandatory = $false)]
        [string]$ErrorReportFolderDirectory,
        [Parameter(Mandatory = $false)]
        [string]$LogReportDirectory,
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 20)]
        [int]$JsonDepth = 5
    )

    function Write-ArrayaErrorExportLog {
        param(
            [Parameter(Mandatory = $true)]
            [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG')]
            [string]$Type,
            [Parameter(Mandatory = $true)]
            [string]$Message
        )

        $writeLogCommand = Get-Command -Name 'Write-Log' -ErrorAction SilentlyContinue
        if ($writeLogCommand) {
            $logSplat = @{
                Type    = $Type
                Message = $Message
            }
            if (-not [string]::IsNullOrWhiteSpace($ExportFileLocation)) {
                $logSplat.ExportFileLocation = $ExportFileLocation
            }
            elseif (-not [string]::IsNullOrWhiteSpace($LogReportDirectory)) {
                $logSplat.ExportFileLocation = $LogReportDirectory
            }

            try {
                & $writeLogCommand @logSplat
                return
            }
            catch {
            }
        }

        if ($Type -in @('WARNING', 'ERROR')) {
            Write-Warning $Message
        }
        else {
            Write-Verbose $Message
        }
    }

    $normalizedErrors = @($ErrorData | Where-Object { $null -ne $_ })
    if ($normalizedErrors.Count -eq 0) {
        Write-ArrayaErrorExportLog -Type 'WARNING' -Message 'No error data provided to export. Skipping error report export.'
        return $null
    }

    Write-ArrayaErrorExportLog -Type 'INFO' -Message 'START: Export all Errors'

    $resolvedBaseName = $BaseName
    $resolvedDirectory = $null
    $resolvedFolderName = $ErrorReportFolderName
    $resolvedFolderPath = $null
    $useDedicatedFolder = $false

    if (-not [string]::IsNullOrWhiteSpace($ExportFileLocation)) {
        $cleanExportFileLocation = $ExportFileLocation -replace '"', ''
        $resolvedDirectory = [System.IO.Path]::GetDirectoryName($cleanExportFileLocation)
        $resolvedBaseName = [System.IO.Path]::GetFileNameWithoutExtension($cleanExportFileLocation)
        $resolvedFolderPath = $resolvedDirectory
    }
    else {
        if (-not [string]::IsNullOrWhiteSpace($ErrorReportFolderDirectory)) {
            $resolvedDirectory = $ErrorReportFolderDirectory
        }
        else {
            $resolvedDirectory = $env:TEMP
        }

        if ([string]::IsNullOrWhiteSpace($resolvedFolderName)) {
            $resolvedFolderName = "$BaseName Error Reporting"
        }

        $useDedicatedFolder = $true
    }

    if ([string]::IsNullOrWhiteSpace($resolvedDirectory)) {
        $resolvedDirectory = (Get-Location).Path
    }

    if ($useDedicatedFolder) {
        $resolvedFolderPath = Join-Path -Path $resolvedDirectory -ChildPath $resolvedFolderName
    }
    elseif ([string]::IsNullOrWhiteSpace($resolvedFolderPath)) {
        $resolvedFolderPath = $resolvedDirectory
    }

    Write-ArrayaErrorExportLog -Type 'INFO' -Message "INFO: Exporting Error Logs to directory $resolvedFolderPath"

    if (-not (Test-Path -Path $resolvedFolderPath)) {
        $null = New-Item -Path $resolvedFolderPath -ItemType Directory -Force
        Write-ArrayaErrorExportLog -Type 'INFO' -Message "INFO: Created error report directory '$resolvedFolderPath'"
    }

    $fileBaseName = "$resolvedBaseName-ErrorLog"
    $jsonPath = Join-Path -Path $resolvedFolderPath -ChildPath "$fileBaseName.json"
    $logPath = Join-Path -Path $resolvedFolderPath -ChildPath "$fileBaseName.log"
    $csvPath = Join-Path -Path $resolvedFolderPath -ChildPath "$fileBaseName.csv"

    $normalizedErrors | ConvertTo-Json -Depth $JsonDepth | Set-Content -Path $jsonPath -Encoding UTF8
    $normalizedErrors | Out-File -FilePath $logPath -Encoding UTF8
    $normalizedErrors | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-ArrayaErrorExportLog -Type 'INFO' -Message "INFO: Exported json Error Logs to $jsonPath"
    Write-ArrayaErrorExportLog -Type 'INFO' -Message "INFO: Exported txt Error Logs to $logPath"
    Write-ArrayaErrorExportLog -Type 'INFO' -Message "INFO: Exported csv Error Logs to $csvPath"

    return [PSCustomObject]@{
        ErrorCount = $normalizedErrors.Count
        FolderPath = $resolvedFolderPath
        JsonPath   = $jsonPath
        LogPath    = $logPath
        CsvPath    = $csvPath
    }
}
