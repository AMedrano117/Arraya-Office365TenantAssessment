# Manages the Primary SMTP Address of the recipient object in an Exchange Environment
function Set-PrimarySMTPAddress {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Recipient,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$NewSMTPAddress,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$NewDomain,

        [Parameter(Mandatory=$true)]
        [string]$LogFilePath
    )

    #Determine New SMTP Address
    if ($NewSMTPAddress){
        $NewSMTPAddress = $NewSMTPAddress
    }
    elseif ($NewDomain) {
        $NewSMTPAddress = ($Recipient.PrimarySmtpAddress -split "@")[0] + "@$NewDomain"
    }
    else {
        throw "Missing Required Parameters for Primary SMTP Address Update. Please specify either the New SMTP Address or the New Domain to apply in this function"
    }

    #check if smtp address is set
    if ($Recipient.PrimarySmtpAddress -eq "$NewSMTPAddress") {
        Write-Host "Primary SMTP Address already set to '$NewSMTPAddress'. Skipping..." -ForegroundColor Yellow
        return
    } else {
        Write-Host "Updating Primary SMTP Address to '$NewSMTPAddress'..." -ForegroundColor Cyan -NoNewline
    }

    try {
        switch ($recipient.RecipientTypeDetails) {
            DynamicDistributionGroup {
                Set-DynamicDistributionGroup -Identity $recipient.Identity -PrimarySmtpAddress $NewSMTPAddress -ErrorAction Stop
                Write-Verbose "Completed updating Primary SMTP Address for '$($Recipient.DisplayName)' to '$NewSMTPAddress'."
            }
            {$_ -in 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', "MailNonUniversalGroup"} {
                Set-DistributionGroup -Identity $recipient.Identity -PrimarySmtpAddress $NewSMTPAddress -ErrorAction Stop
                Write-Verbose "Completed updating Primary SMTP Address for '$($Recipient.DisplayName)' to '$NewSMTPAddress'."
            }
            {$_ -in 'GroupMailbox', 'Team'} {
                Set-UnifiedGroup -Identity $recipient.Identity -PrimarySmtpAddress $NewSMTPAddress -ErrorAction Stop
                Write-Verbose "Completed updating Primary SMTP Address for '$($Recipient.DisplayName)' to '$NewSMTPAddress'."
            }
            {$_ -in 'UserMailbox', 'SharedMailbox', "RoomMailbox", "SchedulingMailbox", "EquipmentMailbox"} {
                Set-Mailbox -Identity $recipient.Identity -WindowsEmailAddress $NewSMTPAddress -ErrorAction Stop
                Write-Verbose "Completed updating Primary SMTP Address for '$($Recipient.DisplayName)' to '$NewSMTPAddress'."
            }
            default {
                Write-Verbose "Recipient Type '$($Recipient.RecipientTypeDetails)' is not supported for this operation." -ForegroundColor Red
                return $null
            }
        }
        Write-Host "Update successful." -ForegroundColor Green
        Write-Log -Type INFO -Message "Primary SMTP Address for '$($Recipient.DisplayName)' updated to '$NewSMTPAddress'." -ExportFileLocation $LogFilePath
    } catch {
        $ErrorMessage = "Failed to update SMTP Address for '$($Recipient.DisplayName)': $($_.Exception.Message)"
        Write-Host $ErrorMessage -ForegroundColor Red
        Write-Log -Type ERROR -Message $ErrorMessage -ExportFileLocation $LogFilePath
        $global:AllErrors += $_
        continue
    }
}