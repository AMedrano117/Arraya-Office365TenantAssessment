# Function to perform actions to Tenant
function Start-T2TDomainCutoverMigration {
    param (
        [Parameter(Mandatory = $True, HelpMessage="What is the location of the Tenant? Source or Destination")]
        [ValidateSet('Source', 'Destination')]
        [string]$TenantLocation,
        [Parameter(Mandatory = $True, HelpMessage='Specifies the domain that is being migrated such as example.org')]
        [string]$MigratingDomain,
        [Parameter(Mandatory=$false, HelpMessage="Use this switch to perform this for all objects in specified location. Use parameter 'ImportedMatchedUsers' for specified users.")]
        [switch]$All,
        [Parameter(Mandatory=$false, HelpMessage="Specifies the input source containing the Matched Users file to be processed. This parameter can accept: - An array of objects, each representing an item with properties. - A string specifying the path to a CSV file with the first row as headers. - A string specifying the path to an Excel file (XLSX) with the first row as headers. Ensure the input has valid headers for data column identification.")]
        [object]$ImportedMatchedData,
        [Parameter(Mandatory=$false,HelpMessage='Do you wish to update the primary SMTP address from migrating domain to .onmicrosoft?')]
        [switch] $PrimarySMTPAddressUpdate,
        [Parameter(Mandatory=$false,HelpMessage='Do you want to remove the migrating Domain alias from recipient objects?')]
        [switch] $RemoveAliases,
        [Parameter(Mandatory=$false,HelpMessage='Do you want to Add the migrating Domain alias from recipient objects?')]
        [switch] $AddAliases,
        [Parameter(Mandatory=$false,HelpMessage='Do you want to update the UPN of users from migrating Domain to .onmicrosoft?')] 
        [switch] $UPNSwitch,
        [Parameter(Mandatory=$false,HelpMessage="Do you want to Test?")] [switch] $TestMode
    )
    ##Specifiy Variables Migrating Domain
    $onMicrosoftDomain = (Get-AcceptedDomain | ? {$_.name -like "*.onmicrosoft.com" -and $_.name -notlike "*.mail.onmicrosoft.com"}).Name

    # Helps specify email addresses with the domain when using @ symbol
    $MigratingDomainEmail = "@" + $MigratingDomain

    switch ($TenantLocation) {
        Source {
            Write-Verbose "Working on Source Tenant"
            # MSOL Users Update UserPrincipalName for Migrating Domain Users Only
            if ($UPNSwitch) {
                Write-Host "Performing User Principal Name Update for Migrating Users" -ForegroundColor Black -BackgroundColor Yellow
                # Get All MSOL Users with migrating domain
                try {
                    Write-Host "Looking for UPN users with '$($MigratingDomain)' domain" -ForegroundColor Cyan
                    $allMigratingMSOLUsers = Get-MgUser -All | Where-Object {$_.UserPrincipalName -like "*$MigratingDomainEmail"}
                    $migratingMSOLUsers = Manage-UserBatches -ProcessAll -UserArray $allMigratingMSOLUsers
                }
                catch {
                    Write-Host "No UPNs with '$($MigratingDomain)' found to update" -ForegroundColor Yellow
                    Continue
                }
        
                $start = Get-Date
                $progresscounter = 0
                $totalCount = $migratingMSOLUsers.count
                foreach ($user in $migratingMSOLUsers) {
                    Write-ProgressHelper -Activity "UPN Domain Cutover for '$($user.displayname)' to '$($onMicrosoftDomain)'" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
                    $newOnMicrosoftUPN = ($user.UserPrincipalName -split "@")[0] + "@$onMicrosoftDomain"
                    Update-UserUPN -User $user -NewUPN $newOnMicrosoftUPN -LogFilePath $ExportDetails
                }
            }

            # Exchange Online Recipients Update for Migrating Domain Users only - Source Tenant Removal
            if ($RemoveAliases -or $PrimarySMTPAddressUpdate) {
                Write-Host "Performing Email Address Changes for Migrating Users" -ForegroundColor Black -BackgroundColor Yellow
                ##Specifiy Variables Migrating Domain
                #$excludedTypes = "MailContact", "GuestMailUser", "MailUser", "DiscoveryMailbox", "MonitoringMailbox"
                $excludedTypes = "DiscoveryMailbox", "MonitoringMailbox"
                
                try {
                    $allMigratingRecipients = Get-EXORecipient -ResultSize Unlimited | ? {
                        $_.EmailAddresses -like "*$MigratingDomainEmail" -and $_.RecipientTypeDetails -notin $excludedTypes
                    } | sort DisplayName
    
                    # Set Batch Jobs
                    $recipientGroup = Manage-UserBatches -ProcessAll -UserArray $allMigratingRecipients
                    #Write-Host "$(($recipientGroup | measure).count) Users to Update"
                }
                catch {
                    Write-Host "No Aliases with '$($MigratingDomain)' found to update" -ForegroundColor Yellow
                    Continue
                }
                
                #Progress Bar 1A
                $start = get-date
                $totalCount = ($recipientGroup).count
                $progresscounter = 0
                foreach ($recipient in $recipientGroup) {
                    #Progress Bar 1B
                    $progresscounter += 1
                    Write-ProgressHelper -Activity "Email Address Domain Cutover for '$($recipient.displayname)' to '$($onMicrosoftDomain)'" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start

                    # Set User Specific Variables
                    $oldPrimarySMTPAddress = $recipient.PrimarySmtpAddress
                    [array]$MigratingEmailAddresses = $recipient.EmailAddresses | ? {$_ -like "*$MigratingDomainEmail*" -and $_ -notlike "*onmicrosoft.com" -and $_ -notlike "*sip:*"}

                    Write-Log -Type INFO -Message "[$($recipient.RecipientTypeDetails)] [$($oldPrimarySMTPAddress)] Start Domain Removal Process" -ExportFileLocation $ExportDetails
                    Write-Log -Type INFO -Message "[$($recipient.RecipientTypeDetails)] [$($oldPrimarySMTPAddress)] Aliases Found - $(($EmailAddresses | measure).Count)" -ExportFileLocation $ExportDetails

                    Write-Host "[$($recipient.RecipientTypeDetails)] [$($oldPrimarySMTPAddress)] " -ForegroundColor Magenta

                    #Check for OnMicrosoft Domain Address
                    $newOnMicrosoftAddress = Manage-DomainAlias -Operation Add -Recipient $recipient -DomainAlias $onMicrosoftDomain
                    
                    #Update PrimarySMTPAddress
                    if ($PrimarySMTPAddressUpdate) {
                        #Write-Host "...PrimarySMTPAddressUpdate.." -NoNewline -foregroundcolor DarkCyan  
                        Set-PrimarySMTPAddress -Recipient $recipient -NewDomain $onMicrosoftDomain -LogFilePath $ExportDetails
                        Write-MigrationLog -tenantName $onMicrosoftDomain -LogPath $ExportDetails -OriginalValue $recipient.PrimarySmtpAddress -NewValue $newOnMicrosoftAddress -ExecutedCommand "SMTPAddressUpdate"
                    }
                    #Remove Migrating Domain Aliases
                    if ($RemoveAliases) {
                        $MigratingAliasCount = ($MigratingEmailAddresses | measure).count
                        Write-Host "Removing " -ForegroundColor Cyan -NoNewline
                        Write-Host "$($MigratingAliasCount) " -ForegroundColor Yellow -NoNewline
                        Write-Host "Aliases for '$($recipient.displayname)'.. " -ForegroundColor Cyan -NoNewline
                        if ($MigratingEmailAddresses) {
                            foreach ($alias in $MigratingEmailAddresses) {
                                try {
                                    $aliasRemove = Manage-DomainAlias -Recipient $recipient -Operation Remove -AliasAddress $alias
                                    Write-MigrationLog -tenantName $onMicrosoftDomain -LogPath $ExportDetails -OriginalValue $recipient.PrimarySmtpAddress -NewValue $alias -ExecutedCommand "RemoveAlias"

                                }
                                catch {
                                    $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "[$($recipient.RecipientTypeDetails)] [$($recipient.PrimarySmtpAddress)] FAILED to Remove Alias '$($alias)' Address'. $($_.Exception.Message)"
                                    $global:AllErrors += $ErrorObject
                                    Write-Log -Type ERROR -Message "[$($recipient.RecipientTypeDetails)] [$($recipient.PrimarySmtpAddress)] FAILED to Remove Alias '$($alias)' Address'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                                    continue
                                }
                            }
                        }
                    }
                }
            }
        }
        Destination {
            Write-Verbose "Working on Destination Tenant"
            #Check if Imported or All users
            # Check if neither Imported nor All users are specified
            if (-not $All -and $null -eq $ImportedMatchedData) {
                Write-Host "Unable to perform operation: Unable to validate data in. Please specify either all users with `-All` or review imported data through `-ImportedMatchedUsers`." -ForegroundColor Red
                return
            }
            elseif ($ImportedMatchedData) {
                # Declare User Specific Attributes to use for Source and Destination Email Address
                [string]$SourceAddressAttribute = Select-ImportedHeader -ImportedFile $ImportedMatchedData -CustomPromptMessage "Select the column that contains the source email addresses"
                [string]$EmailAddressesAttribute = Select-ImportedHeader -ImportedFile $ImportedMatchedData -CustomPromptMessage "Select the column that contains the alias email addresses"
                [string]$TargetAddressAttribute = Select-ImportedHeader -ImportedFile $ImportedMatchedData -CustomPromptMessage "Select the column that contains the target email addresses"
            }

            # MSOL Users Update UserPrincipalName for Migrating Domain Users Only
            if ($UPNSwitch) {
                # Specify Array to perform action
                Write-Host "Performing UPN Updates to Users" -ForegroundColor Black -BackgroundColor Yellow
                if ($All){
                    Write-Host "Gathering All MSOL users" -ForegroundColor Cyan
                    $excludedUPNs = @('amedrano@a-a-ron.org')
                    $excludedNames = @('On-Premises Directory Synchronization Service Account')

                    $usersToUpdate = Get-MgUser -All | Where-Object {
                        $_.UserType -eq "Member" -and $ExcludedUPNs -notcontains $_.UserPrincipalName -and $ExcludedNames -notcontains $_.DisplayName
                    }
                    $migratingMSOLUsers = Manage-UserBatches -ProcessAll -UserArray $usersToUpdate
                }
                ## For Specified Users from imported data
                elseif ($ImportedMatchedData) {
                    $usersToUpdate = Validate-ImportedData -ImportedFile $ImportedMatchedData
                    Write-Verbose "Imported Matched Data"
                }
        
                $start = Get-Date
                $progresscounter = 0
                $totalCount = $usersToUpdate.count
                foreach ($user in $usersToUpdate) {
                    $progresscounter += 1

                    # Specify Recipient to Perform Action and validate address exists
                    $userCheck = @()
                    if ($All) {
                        $userCheck = $user
                        Write-Host "[$($userCheck.displayname)] [$($userCheck.UserPrincipalName)] " -ForegroundColor Magenta
                    }
                    elseif ($ImportedMatchedData) {
                        Write-Host "[$($user.$TargetAddressAttribute)] " -ForegroundColor Magenta -NoNewline
                        $userCheck = Select-FromMultipleO365Matches -RecipientCommand {Get-MgUser -UserId $user.$TargetAddressAttribute -ErrorAction SilentlyContinue}
                        if (-not $userCheck) {
                            Write-Host "No valid user found for $($user.$TargetAddressAttribute). Skipping..." -ForegroundColor Red
                            continue
                        }
                        else {
                            Write-Host "[$($userCheck.displayname)] [$($userCheck.UserPrincipalName)]" -ForegroundColor Magenta
                        }
                    }
                    Write-Verbose "User Check: $($userCheck.UserPrincipalName)"

                    Write-ProgressHelper -Activity "UPN Domain Cutover for '$($userCheck.displayname)' to '$($MigratingDomain)'" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
                    $newUPN = ($userCheck.UserPrincipalName -split "@")[0] + "@$MigratingDomain"
                    Update-UserUPN -User $userCheck -NewUPN $newUPN -LogFilePath $ExportDetails
                }
            }
            # Exchange Online Recipients Update for Migrating Domain Users only - Source Tenant Removal
            if ($AddAliases -or $PrimarySMTPAddressUpdate) {
                # Specify Array to perform action
                Write-Host "Performing Email Address Changes for Migrating Users" -ForegroundColor Black -BackgroundColor Yellow
                if ($All){
                    ##Specifiy Variables Migrating Domain
                    $excludedTypes = "MailContact", "GuestMailUser", "MailUser", "DiscoveryMailbox", "MonitoringMailbox"
                    $allMigratingRecipients = Get-EXORecipient -ResultSize Unlimited | ? {
                        $_.RecipientTypeDetails -notin $excludedTypes
                    } | sort DisplayName
                    $recipientsToUpdate = $allMigratingRecipients
                }
                ## For Specified Users from imported data
                elseif ($ImportedMatchedData) {
                    $recipientsToUpdate = Validate-ImportedData -ImportedFile $ImportedMatchedData
                }
                Write-Verbose "Recipients to Update: $($recipientsToUpdate.count)"
          
                #Progress Bar 1A
                $start = get-date
                $totalCount = ($recipientsToUpdate).count
                $progresscounter = 0
                foreach ($recipient in $recipientsToUpdate) {
                    $progresscounter += 1

                    # Skip if Target Address is not in import list
                    if ($recipient.$TargetAddressAttribute -eq $null) {
                        Write-Verbose "[$($recipient.$TargetAddressAttribute)] is not a valid address for migration. Skipping..."
                        continue
                    }
                    # Specify Recipient to Perform Action and validate address exists
                    $recipientCheck = @()
                    if ($All) {
                        $recipientCheck = $recipient
                        Write-Host "[$($recipientCheck.RecipientTypeDetails)] [$($recipientCheck.displayname)] " -ForegroundColor Magenta
                    }
                    elseif ($ImportedMatchedData) {
                        Write-Host "[$($recipient.$TargetAddressAttribute)] " -ForegroundColor Magenta -NoNewline
                        $recipientCheck = Select-FromMultipleO365Matches -RecipientCommand {Get-ExoRecipient $recipient.$TargetAddressAttribute -ErrorAction SilentlyContinue}
                        if (-not $recipientCheck) {
                            Write-Host "No valid recipient found for $($recipient.$TargetAddressAttribute). Skipping..." -ForegroundColor Red
                            continue
                        }
                        else {
                            Write-Host "[$($recipientCheck.RecipientTypeDetails)] [$($recipientCheck.displayname)] " -ForegroundColor Magenta
                        }
                    }
                    #Progress Bar 1B
                    
                    Write-ProgressHelper -Activity "Email Address Domain Cutover for '$($recipientCheck.displayname)' to '$($MigratingDomain)'" -ProgressCounter ($progresscounter) -TotalCount $totalCount -StartTime $start
                    
                    #Update PrimarySMTPAddress to migrating Domain
                    if ($PrimarySMTPAddressUpdate) {
                        #Write-Host "...PrimarySMTPAddressUpdate.." -NoNewline -foregroundcolor DarkCyan
                        if ($All) {
                            Set-PrimarySMTPAddress -Recipient $recipientCheck -NewDomain $MigratingDomain -LogFilePath $ExportDetails
                            Write-MigrationLog -tenantName $onMicrosoftDomain -LogPath $ExportDetails -OriginalValue $recipientCheck.PrimarySmtpAddress -NewValue $MigratingDomain -ExecutedCommand "SMTPAddressUpdate"
                        }
                        elseif ($ImportedMatchedData) {
                            Set-PrimarySMTPAddress -Recipient $recipientCheck -NewSMTPAddress $recipient.$SourceAddressAttribute -LogFilePath $ExportDetails
                            Write-MigrationLog -tenantName $onMicrosoftDomain -LogPath $ExportDetails -OriginalValue $recipientCheck.PrimarySmtpAddress -NewValue $recipient.$SourceAddressAttribute -ExecutedCommand "SMTPAddressUpdate"
                        }
                    }

                    # Add Migrating Domain Aliases
                    if ($AddAliases) {
                        Write-Host "Adding Aliases to Destination Tenant" -ForegroundColor Black -BackgroundColor Yellow
                        try {
                            $aliasAdd = Manage-DomainAlias -Recipient $recipientCheck -Operation Add -DomainAlias $MigratingDomain
                        }
                        catch {
                            $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "[$($recipientCheck.RecipientTypeDetails)] [$($recipientCheck.PrimarySmtpAddress)] FAILED to Add Alias Address for '$($migratingDomain)'. $($_.Exception.Message)"
                            $global:AllErrors += $ErrorObject
                            Write-Log -Type ERROR -Message "[$($recipientCheck.RecipientTypeDetails)] [$($recipientCheck.PrimarySmtpAddress)] FAILED to Add Alias Address for '$($migratingdomain)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                        }

                        #Gather all aliases to matched users
                        [array]$MigratingEmailAddresses = ($recipient.$EmailAddressesAttribute -split ",") | ? {$_ -like "*$MigratingDomainEmail*" -and $_ -notlike "*onmicrosoft.com" -and $_ -notlike "*sip:*"}
                        if ($MigratingEmailAddresses) {
                            $MigratingAliasCount = ($MigratingEmailAddresses | measure).count
                            Write-Host "Adding " -ForegroundColor Cyan -NoNewline
                            Write-Host "$($MigratingAliasCount) " -ForegroundColor Yellow -NoNewline
                            Write-Host "Additional Aliases for '$($recipientCheck.displayname)'.. " -ForegroundColor Cyan -NoNewline

                            foreach ($alias in $MigratingEmailAddresses) {
                                try {
                                    $alias = $alias.Replace("SMTP:", "").Replace("smtp:", "")
                                    $aliasAdd = Manage-DomainAlias -Recipient $recipientCheck -Operation Add -AliasAddress $alias
                                    Write-MigrationLog -tenantName $onMicrosoftDomain -LogPath $ExportDetails -OriginalValue $recipientCheck.PrimarySmtpAddress -NewValue $alias -ExecutedCommand "AddAlias"
                                }
                                catch {
                                    $ErrorObject = Capture-ErrorHelper -ErrorRecordVar $_ -errorMessage "[$($recipientCheck.RecipientTypeDetails)] [$($recipient.PrimarySmtpAddress)] FAILED to Add Alias Address for '$($migratingdomain)'. $($_.Exception.Message)"
                                    $global:AllErrors += $ErrorObject
                                    Write-Log -Type ERROR -Message "[$($recipientCheck.RecipientTypeDetails)] [$($recipientCheck.PrimarySmtpAddress)] FAILED to Add Alias Address for '$($migratingdomain)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                                }
                            }
                        }
                        else {
                            Write-Host "No Alternate Email Aliases to Add" -ForegroundColor Gray
                        }
                    }
                }    
            }
        }
    }
}