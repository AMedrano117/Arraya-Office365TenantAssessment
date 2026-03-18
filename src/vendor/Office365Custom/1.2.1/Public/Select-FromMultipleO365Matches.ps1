# Performs specified command and helps provide info if multiple matches are found
function Select-FromMultipleO365Matches {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]$RecipientCommand
    )

    try {
        # Execute the recipient command, which should return all matches
        $recipientsFound = & $RecipientCommand
        $recipientsFoundCount = ($recipientsFound | Measure-Object).Count

        if ($recipientsFoundCount -eq 1) {
            $recipientCheck = $recipientsFound
            #Write-Host "Unique recipient found: $($recipientCheck.DisplayName)" -ForegroundColor Green
            return $recipientCheck
        }
        elseif ($recipientsFoundCount -gt 1) {
            Write-Host "Multiple matches found. Please select one to proceed or enter 'Skip' to skip the selection:" -ForegroundColor Yellow
            
            # Display all matches for the user to choose
            $index = 0
            $recipientsFound | ForEach-Object {
                Write-Host "$index`: $($_.DisplayName) - $($_.PrimarySmtpAddress)"
                $index++
            }
            Write-Host "Skip: None of the above"

            # User selection
            $selectedIndex = $null
            while (($selectedIndex -notmatch '^\d+$' -or $selectedIndex -lt 0 -or $selectedIndex -ge $recipientsFoundCount) -and $selectedIndex -ne 'Skip') {
                $selectedIndex = Read-Host "Enter the number corresponding to the correct recipient or 'Skip' to skip"
                if ($selectedIndex -eq 'Skip') {
                    Write-Host "Skipping selection..." -ForegroundColor Cyan
                    return $null
                }
            }
            $recipientCheck = $recipientsFound[$selectedIndex]
            Write-Host "You selected: $($recipientCheck.DisplayName)" -ForegroundColor Green
            return $recipientCheck
        }
        else {
            throw "No object found"
        }
    }
    catch {
        Write-Host "Error fetching object: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}