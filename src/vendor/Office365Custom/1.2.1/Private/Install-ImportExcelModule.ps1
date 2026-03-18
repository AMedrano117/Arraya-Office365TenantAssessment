# ----------------------------------
# Import Other Module Helper Functions
# ----------------------------------
<#
.SYNOPSIS
    Ensures the ImportExcel module is installed and imported.
.DESCRIPTION
    This function checks if the ImportExcel module is available on the system. If it is not installed, 
    it attempts to install the module. Then, it imports the module into the current session. 
    Returns $true if the module is ready to use; otherwise, returns $false.
.EXAMPLE
    if (Install-ImportExcelModule) { Write-Host "ImportExcel module is ready!" }
#>
function Install-ImportExcelModule {
    # Check if ImportExcel module is installed
    if (!(Get-Module -ListAvailable -Name ImportExcel)) {
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not install ImportExcel module. Defaulting to CSV output only."
            return $false
        }
    }

    # Import ImportExcel module
    try {
        Import-Module ImportExcel -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not import ImportExcel module. Defaulting to CSV output only."
        return $false
    }

    return $true
}
