#Public Folder Data; Statistics; Permissions Convert to Hash Tables
function Get-AllPublicFolderDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'operator', 'combined', 'automation', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory=$false, HelpMessage='Specify Exchange Environment')]
        [ValidateSet('On-Premises', 'Office365')]
        [string]$ExchangeEnvironment = 'Office365',
        [Parameter(Mandatory = $false)]
        $Context
    )
    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context -DetailLevel $detailLevel
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $start = Get-Date
    $depthPolicy = $Context.Policies['CollectionDepth']
    # Collect permissions in all detail modes so combined-mode output keeps full public-folder governance visibility.
    $collectPublicFolderPermissions = $true
    $tenantStatsHash["PublicFolderDetails"] = @{}
    Write-ArrayaExchangeCollectorBanner -Message ("[Get-AllPublicFolderDetails] START: Gathering all public folder details with {0} details" -f $detailLevel) -ExportFileLocation $exportDetails
    Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] START: Gathering all public folder details with $($detailLevel) details" -ExportFileLocation $exportDetails
    
    try {
        #Get Public Folder Data, Statistics, and Permissions
        switch ($detailLevel) {
            {$_ -in "minimum", "operator", "combined", "automation", "all"} { 
                $DesiredProperties = @(
                    "Identity", "Name", "MailEnabled"
                    "MailRecipientGuid", "ParentPath", "ContentMailboxName"
                    "EntryId", "FolderSize", "HasSubfolders"
                    "FolderClass", "FolderPath", "ExtendedFolderFlags"
                )
                if ($ExchangeEnvironment -eq 'On-Premises') {
                    $allPublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object { $_.Name -NE "IPM_SUBTREE" } | Select-Object $DesiredProperties
                } else {
                    $allPublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object { $_.Name -NE "IPM_SUBTREE" } | Select-Object $DesiredProperties
                }
            }
            geek {
                if ($ExchangeEnvironment -eq 'On-Premises') {
                    $allPublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object { $_.Name -NE "IPM_SUBTREE" }
                } else {
                    $allPublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object { $_.Name -NE "IPM_SUBTREE" }
                }
            }
        }
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $(($allPublicFolders | Measure-Object).count) Public Folders in Exchange" -ExportFileLocation $exportDetails
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllPublicFolderDetails] An error occurred in running Get-AllPublicFolderDetails function. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_

    }
    
    # Public Folder Statistics
    #**************************
    Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Gathering all public folder statistics" -ExportFileLocation $exportDetails
    try {
        if ($ExchangeEnvironment -eq 'On-Premises') {
            if (-not $PublicFolderDatabase) {
                $PublicFolderDatabase = (Get-PublicFolderDatabase -ErrorAction SilentlyContinue | Select-Object -First 1).Identity
                Write-Host "Using Public Folder Database: $PublicFolderDatabase" -ForegroundColor Yellow
            }
            $PublicFolderStatistics = $allPublicFolders | Get-PublicFolderStatistics -Server $PublicFolderDatabase -ErrorAction SilentlyContinue
        } else {
            $PublicFolderStatistics = $allPublicFolders | Get-PublicFolderStatistics -ErrorAction SilentlyContinue
        }
        $PublicFolderStatsHash = @{}
        foreach($publicFolderStat in $PublicFolderStatistics) {
            $key = $PublicFolderStat.EntryId
            $PublicFolderStatsHash[$key] = $PublicFolderStat
        }
        Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $($($PublicFolderStatistics | Measure-Object).count) Public Folder Statistics in Exchange" -ExportFileLocation $exportDetails
    }
    catch {
        Write-Log -Type ERROR -Message "An error occurred in running Get-AllPublicFolderStatistics function. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }

    # Public Folder Permissions
    #***************************
    $tenantStatsHash["PublicFolderPerms"] = @{}
    $publicFolderPermProgressId = 37
    $publicFolderPermProgressTotal = 1
    if ($collectPublicFolderPermissions) {
        try {
            Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Gathering all public folder permissions" -ExportFileLocation $exportDetails
            $PublicFolderPermissions = $allPublicFolders | get-publicfolderclientpermission -ErrorAction SilentlyContinue
            Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Found $($PublicFolderPermissions.count) public folder permissions" -ExportFileLocation $exportDetails
            
            Write-ArrayaExchangeCollectorSubstep -Message 'Public folders: permissions processing'
            Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] Processing all public folder permissions" -ExportFileLocation $exportDetails
            $totalCount = ($PublicFolderPermissions | Measure-Object).count
            $publicFolderPermProgressTotal = [Math]::Max($totalCount, 1)
            foreach($publicFolderPermission in $PublicFolderPermissions) {
                Write-ProgressHelper -Total $publicFolderPermProgressTotal -Id $publicFolderPermProgressId -Activity "Processing all public folder permissions" -Operation "Gathering Public Folder Permissions for $($publicFolderPermission.Identity)"
                Write-Log -Type DEBUG -Message "[Get-AllPublicFolderDetails] Gathering Public Folder Permissions for $($publicFolderPermission.Identity)" -ExportFileLocation $exportDetails

                $key = "$($publicFolderPermission.Identity)-$($publicFolderPermission.User.Displayname)"
                $permissionObject = @(
                    [PSCustomObject]@{
                        FolderName = $publicFolderPermission.FolderName
                        FolderPath = $publicFolderPermission.Identity
                        Displayname = $publicFolderPermission.User.Displayname
                        PrimarySMTPAddress = $publicFolderPermission.User.RecipientPrincipal.PrimarySmtpAddress
                        AccessRights = ($publicFolderPermission.AccessRights -join ",")
                    }
                )

                if($tenantStatsHash["PublicFolderPerms"].ContainsKey($key)) {
                    $tenantStatsHash["PublicFolderPerms"][$key] += $permissionObject
                }
                else {
                    $tenantStatsHash["PublicFolderPerms"][$key] = @($permissionObject)
                }
            }
        }
        catch {
            Write-Log -Type ERROR -Message "[Get-AllPublicFolderDetails] An error occurred in running Get-AllPublicFolderPermissions function. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
        }
        finally {
            Write-ProgressHelper -Total $publicFolderPermProgressTotal -Id $publicFolderPermProgressId -Activity "Processing all public folder permissions" -Completed
        }
    }
    
    #Combine Stats with Details
    $tenantStatsHash["PublicFolderDetails"] = @{}
    $totalCount = ($allPublicFolders | Measure-Object).count
    foreach($pf in $allPublicFolders) {
        $pfStatsCheck = $PublicFolderStatsHash[$pf.EntryId]
        Write-Log -Type DEBUG -Message "Combine Public Folder Stats for $($pf.FolderPath)" -ExportFileLocation $exportDetails
        $pf | Add-Member -MemberType NoteProperty -Name "FolderPath" -Value ($pf.FolderPath -join ",") -Force
        $pf | Add-Member -MemberType NoteProperty -Name "ItemCount" -Value $pfStatsCheck.ItemCount -Force
        $pf | Add-Member -MemberType NoteProperty -Name "LastModificationTime" -Value $pfStatsCheck.LastModificationTime -Force
        $pf | Add-Member -MemberType NoteProperty -Name "OwnerCount" -Value $pfStatsCheck.OwnerCount -Force
        $pf | Add-Member -MemberType NoteProperty -Name "TotalAssociatedItemSize" -Value $pfStatsCheck.TotalAssociatedItemSize -Force
        $pf | Add-Member -MemberType NoteProperty -Name "TotalDeletedItemSize" -Value $pfStatsCheck.TotalDeletedItemSize -Force
        $pf | Add-Member -MemberType NoteProperty -Name "TotalItemSize" -Value $pfStatsCheck.TotalItemSize -Force
        $pf | Add-Member -MemberType NoteProperty -Name "MailboxOwnerId" -Value $pfStatsCheck.MailboxOwnerId -Force
        $tenantStatsHash["PublicFolderDetails"][$pf.Identity] = $pf
    }

    $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
    Write-ArrayaExchangeCollectorCompletionBanner -Message "[Get-AllPublicFolderDetails] COMPLETED: Gathering all public folder details in $($CompletedTime)" -ExportFileLocation $exportDetails
    Write-Log -Type INFO -Message "[Get-AllPublicFolderDetails] COMPLETED: Gathering all public folder details in $($CompletedTime)" -ExportFileLocation $exportDetails
}
