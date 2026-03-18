#Manages domain aliases for recipient objects in an Exchange environment
function Manage-DomainAlias {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Recipient,

        [Parameter(Mandatory = $false)]
        [string]$AliasAddress,

        [Parameter(Mandatory = $false)]
        [string]$DomainAlias,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Add", "Remove")]
        [string]$Operation
    )

    if ($DomainAlias) {
        $AliasAddress = ($Recipient.PrimarySmtpAddress -split "@")[0] + "@$DomainAlias"
    }

    try { 
        $existingAlias = ($Recipient.EmailAddresses | Where-Object {$_ -like "*$AliasAddress" -and $_ -notlike "*sip:*"} | Select-Object -First 1).Replace("SMTP:", "").Replace("smtp:", "") 
    }
    catch { 
        $existingAlias = $null  
    }

    # Perform the operation based on the parameter
    switch ($Operation) {
        "Add" {
            if ($existingAlias) {
                Write-Host "Alias address '$AliasAddress' already exists." -ForegroundColor Yellow
                return $existingAlias
            } else {
                Write-Host "Adding Alias address '$AliasAddress'..." -ForegroundColor DarkGray -NoNewline
            }
        }
        "Remove" {
            if ($existingAlias) {
                Write-Host "Removing Alias address '$AliasAddress'..." -ForegroundColor DarkGray -NoNewline
            } else {
                Write-Host "Alias address '$AliasAddress' does not exist." -ForegroundColor Yellow
                return $null
            }
        }
    }

    # Create a new domain address based on the recipient's primary SMTP address
    switch ($Recipient.RecipientTypeDetails) {
        DynamicDistributionGroup {
            Set-DynamicDistributionGroup -Identity $recipient.Identity -EmailAddresses @{$Operation="$AliasAddress"} -ErrorAction Stop -WarningAction SilentlyContinue
        }
        {$_ -in 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', "MailNonUniversalGroup"} {
            Set-DistributionGroup -Identity $Recipient.Identity -EmailAddresses @{$Operation="$AliasAddress"} -ErrorAction Stop -WarningAction SilentlyContinue
        }
        {$_ -in 'GroupMailbox', 'Team'} {
            Set-UnifiedGroup -Identity $Recipient.Identity -EmailAddresses @{$Operation="$AliasAddress"} -ErrorAction Stop -WarningAction SilentlyContinue
        }
        {$_ -in 'UserMailbox', 'SharedMailbox', "RoomMailbox", "SchedulingMailbox", "EquipmentMailbox"} {
            Set-Mailbox -Identity $Recipient.Identity -EmailAddresses @{$Operation="$AliasAddress"} -ErrorAction Stop -WarningAction SilentlyContinue
        }
        default {
            Write-Host "Recipient Type '$($Recipient.RecipientTypeDetails)' is not supported for this operation." -ForegroundColor Red
            return $null
        }
    }
    Write-Host "successfully completed." -ForegroundColor Green
    return $AliasAddress 
}