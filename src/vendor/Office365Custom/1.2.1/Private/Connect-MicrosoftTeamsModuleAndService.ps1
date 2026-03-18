    # Function to connect to Microsoft Teams
    function Connect-MicrosoftTeamsModuleAndService {
        Write-Host "Microsoft Teams: Checking for Existing Connections and Required Modules" -ForegroundColor Cyan
        try {
            $MicrosoftTeamsTenant = (Get-CsTenant).DisplayName
            if (Confirm-Tenant $MicrosoftTeamsTenant) {
                return
            } else {
                Disconnect-MicrosoftTeams
            }
        } catch {
            Import-RequiredModule -ModuleName "MicrosoftTeams" -InstallMessage "Run 'Install-Module MicrosoftTeams' as an Administrator."
        }

        Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Yellow
        try {
            Connect-MicrosoftTeams -ErrorAction Stop
            $MicrosoftTeamsTenant = (Get-CsTenant).DisplayName
            Update-TitleBar $MicrosoftTeamsTenant
            Write-Host "Connected: $($MicrosoftTeamsTenant)" -ForegroundColor Green
        } catch {
            Write-Error "Error connecting to Microsoft Teams: $($_.Exception.Message)"
        }
    }
