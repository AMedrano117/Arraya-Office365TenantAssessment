    # Function to connect to Exchange Online
    function Connect-ExchangeOnlineModuleAndService {
        Write-Host "Exchange Online: Checking for Existing Connections and Required Modules" -ForegroundColor Cyan
        try {
            $EXOOrgCheck = Get-OrganizationConfig -ErrorAction Stop
            if (Confirm-Tenant $EXOOrgCheck.Name) {
                $global:tenant = $EXOOrgCheck.Name
                return
            }
            else {
                Disconnect-ExchangeOnline -Confirm:$false
            }
        } catch {
            Import-RequiredModule -ModuleName "ExchangeOnlineManagement" -InstallMessage "Run 'Install-Module ExchangeOnlineManagement' as an Administrator."
        }

        Write-Host "Connecting to ExchangeOnline..." -ForegroundColor Yellow
        try {
            Connect-ExchangeOnline -ErrorAction Stop *> $null
            $EXOOrgCheck = Get-OrganizationConfig -ErrorAction Stop
            Update-TitleBar $EXOOrgCheck.Name
            Write-Host "Connected: $($EXOOrgCheck.Name)" -ForegroundColor Green
            $global:tenant = $EXOOrgCheck.Name
        } catch {
            Write-Error "Error connecting to ExchangeOnline: $($_.Exception.Message)"
        }
    }
