# ----------------------------------
# Connect to Services Helper Functions
# ----------------------------------

# Helper function to check and import modules
function Import-RequiredModule {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,

        [Parameter(Mandatory=$false)]
        [string]$InstallMessage = "Please install or verify the $ModuleName module before continuing.",

        [Switch]$UseWindowsPowerShell = $false
    )

    Write-Verbose "Checking if '$ModuleName' is installed."
    $installedModule = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue
    if (-not $installedModule) {
        Write-Error $InstallMessage
        throw "$ModuleName module not found."
    }

    Write-Verbose "'$ModuleName' is installed. Checking if it is already loaded."
    if (Get-Module -Name $ModuleName) {
        Write-Verbose "'$ModuleName' is already loaded in this session."
        return
    }

    Write-Verbose "Attempting to import '$ModuleName' module..."
    try {
        if ($UseWindowsPowerShell) {
            Write-Verbose "Using Windows PowerShell import switch."
            Import-Module $ModuleName -UseWindowsPowerShell -ErrorAction Stop
        }
        else {
            Write-Verbose "Using '-Force' parameter."
            Import-Module $ModuleName -Force -ErrorAction Stop
        }
        Write-Host "$ModuleName module imported successfully." -ForegroundColor Green
    }
    catch {
        Write-Verbose "Initial import attempt failed. Trying '-Force' import as a fallback."
        try {
            Import-Module $ModuleName -Force -ErrorAction Stop
            Write-Host "$ModuleName module imported successfully (fallback)." -ForegroundColor Green
        }
        catch {
            Write-Error "Error importing '$ModuleName': $($_.Exception.Message)"
            throw
        }
    }
}
