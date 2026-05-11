function Get-AllRecipientDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,HelpMessage='Provide the level of detail')]
        [ValidateSet('minimum', 'operator', 'combined', 'automation', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory = $false)]
        $Context
    )
    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context -DetailLevel $detailLevel
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $initialStart = Get-ArrayaExchangeCollectorStartTime -Context $Context
    $recipientProgressId = 33
    $cacheKey = "Exchange:AllRecipients:$detailLevel"
    $cachedRecipients = Get-ArrayaCollectorCacheValue -Context $Context -Key $cacheKey
    if ($null -ne $cachedRecipients) {
        $tenantStatsHash['AllRecipients'] = $cachedRecipients
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] Reused cached Exchange recipient inventory for $detailLevel details." -ExportFileLocation $exportDetails
        return $tenantStatsHash['AllRecipients']
    }
    try {
        $start = Get-Date
        $tenantStatsHash["AllRecipients"] = @{}
        Write-ArrayaExchangeCollectorBanner -Message ("[Get-AllRecipientDetails] START: Gathering all Exchange Online Recipients {0} details" -f $detailLevel) -ExportFileLocation $exportDetails
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] START: Gathering all Exchange Online Recipients $($detailLevel) details" -ExportFileLocation $exportDetails
        Write-Progress -Id $recipientProgressId -Activity "Gathering All Exchange Online Recipients" -Status (((Get-Date) - $initialStart).ToString('hh\:mm\:ss'))

        switch ($detailLevel) {
            {$_ -in "minimum", "operator", "combined", "automation", "all"} { 
                $Properties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Identity", "RecipientTypeDetails", "PrimarySMTPAddress"
                    "EmailAddresses", "HiddenFromAddressListsEnabled", "AddressBookPolicy"
                    "ManagedBy", "SKUAssigned", "WhenCreated", "WhenSoftDeleted", "GUID"
                    "alias", "Notes"
                )
                # Properties for Select-Object
                $DesiredProperties = @(
                    "ExternalDirectoryObjectId", "DisplayName", "Identity", "RecipientTypeDetails", "PrimarySMTPAddress",
                    @{Name="EmailAddresses"; Expression={$_.EmailAddresses -join ","}}, 
                    "HiddenFromAddressListsEnabled", "AddressBookPolicy",
                    @{Name="ManagedBy"; Expression={$_.ManagedBy -join ","}}, 
                    "SKUAssigned", "WhenCreated", "WhenSoftDeleted", "GUID",
                    "alias", "Notes"
                )
                $allRecipients = Invoke-QuietCommand -ScriptBlock {
                    Get-EXORecipient -Properties $Properties -ResultSize Unlimited -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox' -and RecipientTypeDetails -ne 'MailContact' -and RecipientTypeDetails -ne 'GuestMailUser' -and RecipientTypeDetails -ne 'MailUser'" -ErrorAction Stop | select $DesiredProperties
                }
            }           
            geek {
                $allRecipients = Invoke-QuietCommand -ScriptBlock {
                    Get-EXORecipient -PropertySets All -ResultSize Unlimited -Filter "RecipientTypeDetails -ne 'DiscoveryMailbox' -and RecipientTypeDetails -ne 'MailContact' -and RecipientTypeDetails -ne 'GuestMailUser' -and RecipientTypeDetails -ne 'MailUser'" -ErrorAction Stop
                }
            }
        }
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] FOUND $($allRecipients.count) Exchange Online Recipients $($detailLevel) details" -ExportFileLocation $exportDetails
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] Adding Exchange Online Recipients to Tenant Stats Hash" -ExportFileLocation $exportDetails
        foreach ($recipient in $allRecipients) {
            $tenantStatsHash["AllRecipients"][$recipient.PrimarySmtpAddress] = $recipient
        }
        Set-ArrayaCollectorCacheValue -Context $Context -Key $cacheKey -Value $tenantStatsHash['AllRecipients'] | Out-Null
    }
    catch {
        Write-Log -Type ERROR -Message "[Get-AllRecipientDetails] An error occurred in running Get-AllRecipientDetails function. $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        Write-Progress -Id $recipientProgressId -Activity "Gathering All Exchange Online Recipients" -Completed
        $CompletedTime = (((Get-Date) - $start).ToString('hh\:mm\:ss'))
        Write-ArrayaExchangeCollectorCompletionBanner -Message "[Get-AllRecipientDetails] COMPLETED: Gathering all Exchange Online Recipients in $($CompletedTime)" -ExportFileLocation $exportDetails
        Write-Log -Type INFO -Message "[Get-AllRecipientDetails] COMPLETED: Gathering all Exchange Online Recipients" -ExportFileLocation $exportDetails
    }
}
