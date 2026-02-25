# Function to export recipients with migrating domain
function Export-MigratingDomainRecipients {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$MigratingDomain,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ExportPath
    )

    Write-Host "Exporting recipients with migrating domain '$($MigratingDomain)'..." -ForegroundColor Cyan
    
    # Create timestamp for filename and use parent folder
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $parentFolder = Split-Path -Path $ExportPath -Parent
    $exportFile = Join-Path $parentFolder "MigratingDomainRecipients_$($MigratingDomain)_$timestamp.xlsx"
    
    # Get all recipients with migrating domain
    $migratingDomainEmail = "@" + $MigratingDomain
    $excludedTypes = "DiscoveryMailbox", "MonitoringMailbox"
    
    try {
        $recipients = Get-EXORecipient -ResultSize Unlimited | Where-Object {
            $_.EmailAddresses -like "*$migratingDomainEmail" -and $_.RecipientTypeDetails -notin $excludedTypes
        } | Select-Object DisplayName, 
            PrimarySmtpAddress, 
            RecipientTypeDetails,
            EmailAddresses,
            UserPrincipalName,
            @{Name='MigratingDomainAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -like "*$migratingDomainEmail"}) -join ';'}},
            @{Name='OtherAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -notlike "*$migratingDomainEmail"}) -join ';'}}

        if ($recipients) {
            # Export to Excel
            $recipients | Export-Excel -Path $exportFile -WorkSheetName "MigratingRecipients"
            
            Write-Host "Successfully exported $($recipients.Count) recipients to: $exportFile" -ForegroundColor Green
            Write-Host "Recipient Types Found:" -ForegroundColor Yellow
            $recipients | Group-Object RecipientTypeDetails | ForEach-Object {
                Write-Host "  $($_.Name): $($_.Count)"
            }
        } else {
            Write-Host "No recipients found with migrating domain '$($MigratingDomain)'" -ForegroundColor Yellow
        }
    }
    catch {
        $ErrorMessage = "Error exporting recipients: $($_.Exception.Message)"
        Write-Host $ErrorMessage -ForegroundColor Red
        $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage $ErrorMessage
        $global:AllErrors += $ErrorObject
    }
}