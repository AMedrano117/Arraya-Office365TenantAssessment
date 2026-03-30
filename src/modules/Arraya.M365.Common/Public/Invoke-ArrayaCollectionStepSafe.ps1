function Invoke-ArrayaCollectionStepSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OperationName,
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $DefaultValue = $null,
        [Parameter(Mandatory = $false)]
        [string]$ExportFileLocation
    )

    try {
        return (& $ScriptBlock)
    }
    catch {
        $warningMessage = "[$OperationName] $($_.Exception.Message). Continuing."
        $suppressConsoleWarning = $warningMessage -match '(?i)conditional access|inbox rule'
        $writeLogCommand = Get-Command -Name 'Write-Log' -ErrorAction SilentlyContinue
        if ($writeLogCommand) {
            $logSplat = @{
                Type    = 'WARNING'
                Message = $warningMessage
            }
            if (-not [string]::IsNullOrWhiteSpace($ExportFileLocation)) {
                $logSplat.ExportFileLocation = $ExportFileLocation
            }

            try {
                & $writeLogCommand @logSplat
            }
            catch {
                if ($suppressConsoleWarning) {
                    Write-Verbose $warningMessage
                }
                else {
                    Write-Warning $warningMessage
                }
            }
        }
        else {
            if ($suppressConsoleWarning) {
                Write-Verbose $warningMessage
            }
            else {
                Write-Warning $warningMessage
            }
        }

        return $DefaultValue
    }
}
