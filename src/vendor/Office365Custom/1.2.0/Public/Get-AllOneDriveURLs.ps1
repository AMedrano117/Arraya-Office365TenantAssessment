# ----------------------------------
# Office 365 Specific Helper Functions
# ----------------------------------
<#
.SYNOPSIS
    Retrieves OneDrive URLs for all users in a tenant.
.DESCRIPTION
    This function connects to SharePoint Online using the specified admin URL, retrieves all personal site URLs 
    matching the OneDrive pattern, and processes them to build a hashtable mapping the site owner to 
    the OneDrive URL. It also identifies duplicate owners and displays overall statistics.
.PARAMETER AdminUrl
    The SharePoint Online Admin URL used to connect to the tenant.
.PARAMETER OneDriveHash
    A hashtable that will be populated with key/value pairs where the key is the site owner and the value is the OneDrive URL.
.EXAMPLE
    $oneDriveSites = @{}
    Get-OneDriveURLs -AdminUrl "https://yourtenant-admin.sharepoint.com" -OneDriveHash $oneDriveSites
#>
# Function to gather OneDrive URLs for a tenant - NEW
function Get-AllOneDriveURLs {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AdminUrl,
        [Parameter(Mandatory = $true)]
        [hashtable]$OneDriveHash
    )

    # Connect to SharePoint Online
    Connect-SharePointOnline -SPOAdminURL $AdminUrl

    # Create a temporary hash table for processing OneDrive URLs
    $OneDriveHashTmp = @{}
    $global:duplicateOwners = @()  # Global variable to store duplicate owners

    Write-Host "Gathering all OneDrive URLs for tenant: $AdminUrl" -ForegroundColor Cyan
    Write-Host "This may take a while depending on the number of OneDrive URLs to process." -ForegroundColor Yellow

    $starttime = Get-Date
    $SPOSites = Get-SPOSite -IncludePersonalSite $true -Limit All -ErrorAction Stop
    $oneDriveURLs = $SPOSites | Where-Object { $_.Url -like "*-my.sharepoint.com/personal/*" }

    $I = 0
    foreach ($site in $oneDriveURLs) {
        $I++
        Write-Progress -Activity "Processing OneDrive URLs" `
                       -Status "Processing $($site.Owner) - ($I of $($oneDriveURLs.Count))" `
                       -PercentComplete (($I / $oneDriveURLs.Count) * 100)
        if ($OneDriveHashTmp.ContainsKey($site.Owner.ToString())) {
            $global:duplicateOwners += $site
            Write-Verbose "$($site.Owner) already exists. Verifying default OneDrive URL."
            $OneDriveURL = ((Get-GraphData -Uri "https://graph.microsoft.com/v1.0/users/$($site.Owner)/drive" -ErrorAction SilentlyContinue).webUrl) -replace "/Documents$", ""
        }
        else {
            $OneDriveURL = $site.Url
        }
        $OneDriveHashTmp[$site.Owner.ToString()] = $OneDriveURL
    }

    Write-Verbose "OneDrive URLs gathered for $($OneDriveHashTmp.Count) users."

    # Merge the temporary hash table into the provided destination hash table.
    foreach ($key in $OneDriveHashTmp.Keys) {
        $OneDriveHash[$key] = $OneDriveHashTmp[$key]
    }

    $CompletedTime = (((Get-Date) - $starttime).ToString('hh\:mm\:ss'))
    Write-Host "Number of OneDrive URLs gathered: $($OneDriveHash.Count)" -ForegroundColor Green
    Write-Host "Time taken to retrieve OneDrive URLs: $CompletedTime"
    Write-Host "Duplicate Owners Found: $($global:duplicateOwners.count)"
}
