# Function to update UPN using Microsoft Graph
function Update-UserUPN {
    param (
        [Parameter(Mandatory=$true)]
        [object]$User,

        [Parameter(Mandatory=$true)]
        [string]$NewUPN,

        [Parameter(Mandatory=$true)]
        [string]$LogFilePath
    )

    try {
        Write-Host "[$($User.displayname)] [$($User.UserPrincipalName)]" -ForegroundColor Magenta
        Write-Host "UPNUpdate.." -NoNewline -foregroundcolor DarkCyan
        Write-host "Set to $($NewUPN)... " -NoNewline

        # Update UPN using Microsoft Graph
        Update-MgUser -UserId $User.Id -UserPrincipalName $NewUPN -ErrorAction Stop
        
        Write-MigrationLog -tenantName $NewUPN.Split('@')[1] -LogPath $LogFilePath -OriginalValue $User.UserPrincipalName -NewValue $NewUPN -ExecutedCommand "Update-MgUser"
        Write-host "Completed" -foregroundcolor Green
    }
    catch {
        $ErrorMessage = "[$($User.displayname)] [$($User.UserPrincipalName)] FAILED to Update UPN. $($_.Exception.Message)"
        Write-Host $ErrorMessage -ForegroundColor Red
        Write-Log -Type ERROR -Message $ErrorMessage -ExportFileLocation $LogFilePath
        $global:AllErrors += $_
        return $false
    }
    return $true
}