<#
.SYNOPSIS
    Displays available headers from imported data and prompts user to select one.
.DESCRIPTION
    This helper function lists the headers (property names) from the provided data and allows the user to select one.
.PARAMETER ImportedFile
    The source data (object, CSV, or Excel) to process.
.PARAMETER CustomPromptMessage
    An optional custom prompt message.
.EXAMPLE
    Select-ImportedHeader -ImportedFile $data -CustomPromptMessage "Choose a header:"
#>
function Select-ImportedHeader {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Specifies the input source containing the data to be processed.")]
        [object]$ImportedFile,

        [Parameter(Mandatory=$false, HelpMessage="Custom prompt message for selecting the header.")]
        [string]$CustomPromptMessage = "Please enter the header you want to use"
    )
    
    # Validate and load the data
    $importedData = Validate-ImportedData -ImportedFile $ImportedFile
    $headers = $importedData[0].PSObject.Properties.Name

    # Display headers found in the file
    Write-Verbose "The following headers were found in the import data: $($headers -join ', ')"
    Write-Host "The following headers were found in the import data:" -ForegroundColor Cyan
    $headers | ForEach-Object { Write-Host " - $_" }

    # Allow custom prompt modification
    $selectedHeader = Read-Host $CustomPromptMessage
    if ($headers -contains $selectedHeader.Trim()) { # Trim the input to avoid whitespace issues
        Write-Verbose "Selected header: $($selectedHeader.Trim())"
        Write-Host "You've selected to work with the header: $($selectedHeader.Trim())" -ForegroundColor Green
        return $selectedHeader.Trim()
    } else {
        throw "The selected header does not exist in the import file. Please check the header name and try again."
    }
}
