    
    # Function to connect to SharePoint Online
    function Connect-SharePointOnlineModuleAndService {
        Write-Host "SharePoint Online: Checking for Existing Connections and Required Modules" -ForegroundColor Cyan
        try {
            $rootSiteURL = Get-SPOSite -Limit 1 -ErrorAction Stop -WarningAction SilentlyContinue
            $rootURL = $rootSiteURL.Url -replace '/sites.*', ''
            if (Confirm-Tenant $rootURL) {
                return
            }
        } catch {
            if ($isPowerShell7) {
                Import-RequiredModule -UseWindowsPowerShell -ModuleName "Microsoft.Online.SharePoint.PowerShell" -InstallMessage "Run 'Install-Module Microsoft.Online.SharePoint.PowerShell' as an Administrator."
            } else {
                Import-RequiredModule -ModuleName "Microsoft.Online.SharePoint.PowerShell" -InstallMessage "Run 'Install-Module Microsoft.Online.SharePoint.PowerShell' as an Administrator."
            }
        }

        Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow
        try {
            $SPOAdminURL = Read-Host -Prompt "Provide the SharePoint Online Admin URL. Name is usually formatted 'https://<yourtenant>-admin.sharepoint.com'"
            Connect-SPOService -Url $SPOAdminURL -ErrorAction Stop
            $rootSiteURL = Get-SPOSite -Limit 1 -ErrorAction Stop -WarningAction SilentlyContinue
            $rootURL = $rootSiteURL.Url -replace '/sites.*', ''
            Update-TitleBar $rootURL
            Write-Host "Connected: $($rootURL)" -ForegroundColor Green
        } catch {
            Write-Error "Error connecting to SharePoint Online: $($_.Exception.Message)"
        }
    }
