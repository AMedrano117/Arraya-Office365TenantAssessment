# ----------------------------------
# Data Helper Functions
# ----------------------------------

<#
.SYNOPSIS
    Validates and imports data from various sources.
.DESCRIPTION
    This function accepts an object, CSV file path, or Excel file path and returns an array of imported objects.
.PARAMETER ImportedFile
    The object, CSV file path, or Excel file path to import.
.EXAMPLE
    Validate-ImportedData -ImportedFile "C:\data.csv"
#>
function Validate-ImportedData {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Accepts either an array of objects, a CSV file path, or an Excel (XLSX) file path.")]
        [object]$ImportedFile
    )

    Write-Verbose "Validating and importing data from the input source..."
    Write-Verbose "Input type: $($ImportedFile.GetType().Name)"

    # Determine the type of the input and handle accordingly
    switch ($ImportedFile.GetType().Name) {
        "Object[]" { # Array handling
            return $ImportedFile
        }
        "Hashtable" { # Hashtable handling
            return $ImportedFile.Values
        }
        "System.Object" { # Array handling
            return $ImportedFile
        }
        "String" { # Assuming file path for CSV or Excel
            if ($ImportedFile -like "*.csv") {
                # Handle quotes in input
                $ImportedFile = $ImportedFile -replace '"', ''
                return Import-Csv -Path $ImportedFile
            } elseif ($ImportedFile -like "*.xlsx") {
                # Handle quotes in input
                $ImportedFile = $ImportedFile -replace '"', ''

                # Load the Excel file and list the worksheets
                $sheets = Get-ExcelSheetInfo -Path $ImportedFile

                # If there are multiple worksheets, prompt the user to select one
                if ($sheets.Count -gt 1) {
                    Write-Host "The following worksheets were found in the Excel file:" -ForegroundColor Cyan
                    Write-Verbose "Worksheets found: $($sheets.Name -join ', ')"
                    $sheets | ForEach-Object { Write-Host " - $($_.Name)" }

                    $selectedSheet = Read-Host "Please enter the worksheet name you want to use"
                    if ($sheets.Name -contains $selectedSheet) {
                        return Import-Excel -Path $ImportedFile -WorksheetName $selectedSheet
                    } else {
                        throw "The selected worksheet does not exist in the Excel file. Please check the worksheet name and try again."
                    }
                } else {
                    # Only one worksheet, use it automatically
                    Write-Verbose "Only one worksheet found in the Excel file. Using the first worksheet."
                    return Import-Excel -Path $ImportedFile -WorksheetName $sheets[0].Name
                }
            } else {
                throw "Unsupported file type. Only array of objects, CSV, and XLSX are supported."
            }
        }
        default {
            throw "Unsupported input type. The input must be an array of objects, a CSV file path, or an Excel file path."
        }
    }
}
