<#
.SYNOPSIS
    Processes one or more error records into a structured error object.
.DESCRIPTION
    This function accepts an ErrorRecord and an associated error message. It iterates through 
    each error record, extracts relevant details (including the exception message and related information),
    creates a custom object for each error, and optionally stores them in a global list.
.PARAMETER ErrorRecordVar
    One or more error records that will be processed.
.PARAMETER errorMessage
    A descriptive message to display and include with the error details.
.EXAMPLE
    Capture-ErrorHelper -ErrorRecordVar $ErrorRecord -errorMessage "An error occurred while processing."
#>
function Capture-ErrorHelper {
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecordVar,
        [Parameter(Mandatory=$true)]
        [string]$errorMessage
    )
    # Initialize a List to store error details
    $currentErrors = New-Object System.Collections.Generic.List[pscustomobject]

    Write-Host $errorMessage -ForegroundColor Red

    if ($ErrorRecordVar) {
        foreach ($errorCheck in $ErrorRecordVar) {
            if ($errorCheck.Exception.Message -match "'[^']*/[^']*'") {
                $recipient = $matches[0].Trim("'")
            } else {
                $recipient = $null
            }

            # Create a custom object for each error
            $currentError = [PSCustomObject]@{
                TimeStamp           = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                ErrorMessage        = $errorMessage
                Commandlet          = $errorCheck.CategoryInfo.Activity
                Reason              = $errorCheck.CategoryInfo.Reason
                "Exception-Message" = $errorCheck.Exception.Message
                Exception           = $errorCheck.Exception
                Recipient           = $recipient
                TargetObject        = $errorCheck.TargetObject
            }

            # Add the custom object to the List
            $currentErrors.Add($currentError)
        }

        # Optionally, you might want to store these errors in a global variable
        if (!$global:AllDiscoveryErrors) {
            $global:AllDiscoveryErrors = New-Object System.Collections.Generic.List[pscustomobject]
        }
        foreach ($error in $currentErrors) {
            $global:AllDiscoveryErrors.Add($error)
        }
    }
}
