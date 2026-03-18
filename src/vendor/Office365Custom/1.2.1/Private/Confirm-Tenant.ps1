# Helper function to prompt user for tenant confirmation
function Confirm-Tenant {
    param (
        [string]$TenantName
    )
    Write-Host "Already Connected: $TenantName" -ForegroundColor Green
    $userInput = Read-Host -Prompt "Connected to the correct tenant? (Yes/Y/No/N). Default (blank) is Yes"
    if ($userInput -in "Yes", "Y","") {
        return $true
    } else {
        return $false
    }
}
