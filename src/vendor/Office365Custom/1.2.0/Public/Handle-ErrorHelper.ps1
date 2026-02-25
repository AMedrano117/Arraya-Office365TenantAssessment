<#
.SYNOPSIS
    Handles errors by outputting a structured error object.
.DESCRIPTION
    This function takes an ErrorRecord and an error message, then processes and returns a custom error object.
.PARAMETER ErrorRecordVar
    The error record to process.
.PARAMETER errorMessage
    A descriptive error message.
.EXAMPLE
    Handle-ErrorHelper -ErrorRecordVar $ErrorRecord -errorMessage "An error occurred"
#>
function Handle-ErrorHelper {
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecordVar,
        [Parameter(Mandatory=$true)]
        [string]$errorMessage
    )

    if($ErrorRecordVar) {
        Write-Host $errorMessage -ForegroundColor Red
        foreach ($errorCheck in $ErrorRecordVar) {
            if ($errorCheck.Exception.Message -match "'[^']*/[^']*'") {
                $recipient = $matches[0].Trim("'")
            } else {
                $recipient = $null
            }

            $CurrentError = New-Object PSObject
            $CurrentError | Add-Member -Type NoteProperty -Name "TimeStamp" -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Force
            $CurrentError | Add-Member -Type NoteProperty -Name "ErrorMessage" -Value $errorMessage -Force
            $CurrentError | Add-Member -Type NoteProperty -Name "Commandlet" -Value $errorCheck.CategoryInfo.Activity
            $CurrentError | Add-Member -Type NoteProperty -Name "Reason" -Value $errorCheck.CategoryInfo.Reason -Force
            $CurrentError | Add-Member -Type NoteProperty -Name "Exception" -Value $errorCheck.Exception.Message -Force
            $CurrentError | Add-Member -Type NoteProperty -Name "Exception-Message" -Value $errorCheck.Exception -Force
            $CurrentError | Add-Member -Type NoteProperty -Name "Recipient" -Value $recipient -Force
            $CurrentError | Add-Member -Type NoteProperty -Name "TargetObject" -Value $errorCheck.TargetObject -Force
            #$CurrentError | Add-Member -Type NoteProperty -Name "FullError" -Value $errorCheck -Force

            return $CurrentError
        }
    }
}
