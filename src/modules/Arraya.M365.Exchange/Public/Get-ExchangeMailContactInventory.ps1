function Get-ExchangeMailContactInventory {
    <#
    .SYNOPSIS
        Collects Exchange mail contacts and mail users into the MailContacts dataset.

    .DESCRIPTION
        The recipient inventory (Get-AllRecipientDetails) excludes mail contacts and mail
        users, so they never reached any output. Neither object type migrates with a mailbox
        move; each has to be recreated in the destination, and contacts are often members of
        distribution lists that are recreated too. Guest mail users are not included: guests
        are inventoried from Entra ID.

        Read-only. Uses Get-EXORecipient, and falls back to Get-Recipient when the REST
        cmdlet rejects the property list.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = 'Provide the level of detail')]
        [ValidateSet('minimum', 'operator', 'combined', 'automation', 'all', 'geek')]
        [string]$detailLevel,
        [Parameter(Mandatory = $false)]
        $Context
    )

    $Context = Resolve-ArrayaExchangeCollectorContext -Context $Context -DetailLevel $detailLevel
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $rows = [ordered]@{}
    $evidenceState = 'CompleteWithRecords'
    $evidenceReason = $null
    $start = Get-Date

    Write-ArrayaExchangeCollectorBanner -Message '[Get-ExchangeMailContactInventory] START: Gathering mail contacts and mail users' -ExportFileLocation $exportDetails
    try {
        $properties = @(
            'ExternalDirectoryObjectId', 'DisplayName', 'Alias', 'RecipientTypeDetails', 'PrimarySmtpAddress',
            'ExternalEmailAddress', 'EmailAddresses', 'HiddenFromAddressListsEnabled', 'IsDirSynced', 'WhenCreated'
        )
        $recipients = $null
        try {
            $recipients = @(Invoke-QuietCommand -ScriptBlock {
                    Get-EXORecipient -RecipientTypeDetails MailContact, MailUser -Properties $properties -ResultSize Unlimited -ErrorAction Stop
                })
        }
        catch {
            Write-Log -Type WARNING -Message "[Get-ExchangeMailContactInventory] Get-EXORecipient failed ($($_.Exception.Message)); retrying with Get-Recipient." -ExportFileLocation $exportDetails
            $recipients = @(Invoke-QuietCommand -ScriptBlock {
                    Get-Recipient -RecipientTypeDetails MailContact, MailUser -ResultSize Unlimited -ErrorAction Stop
                })
        }

        foreach ($recipient in @($recipients | Where-Object { $null -ne $_ })) {
            $externalAddress = [string](Get-ArrayaObjectValue -Object $recipient -Names @('ExternalEmailAddress'))
            $externalAddress = $externalAddress -replace '^(?i)smtp:', ''
            $primarySmtpAddress = [string](Get-ArrayaObjectValue -Object $recipient -Names @('PrimarySmtpAddress'))
            $key = @($primarySmtpAddress, $externalAddress, [string](Get-ArrayaObjectValue -Object $recipient -Names @('ExternalDirectoryObjectId'))) |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -First 1
            if ([string]::IsNullOrWhiteSpace([string]$key)) {
                $key = 'contact-{0:D5}' -f ($rows.Count + 1)
            }
            if ($rows.Contains([string]$key)) {
                continue
            }

            $rows[[string]$key] = [pscustomobject]@{
                DisplayName                   = [string](Get-ArrayaObjectValue -Object $recipient -Names @('DisplayName'))
                Alias                         = [string](Get-ArrayaObjectValue -Object $recipient -Names @('Alias'))
                RecipientTypeDetails          = [string](Get-ArrayaObjectValue -Object $recipient -Names @('RecipientTypeDetails'))
                PrimarySmtpAddress            = $primarySmtpAddress
                ExternalEmailAddress          = $externalAddress
                EmailAddresses                = (@(Get-ArrayaObjectValue -Object $recipient -Names @('EmailAddresses')) | ForEach-Object { [string]$_ }) -join ','
                HiddenFromAddressListsEnabled = Get-ArrayaObjectValue -Object $recipient -Names @('HiddenFromAddressListsEnabled')
                IsDirSynced                   = Get-ArrayaObjectValue -Object $recipient -Names @('IsDirSynced')
                ExternalDirectoryObjectId     = [string](Get-ArrayaObjectValue -Object $recipient -Names @('ExternalDirectoryObjectId'))
                WhenCreated                   = Get-ArrayaObjectValue -Object $recipient -Names @('WhenCreated')
            }
        }

        if ($rows.Count -eq 0) {
            $evidenceState = 'CompleteEmpty'
        }
        Write-Log -Type INFO -Message "[Get-ExchangeMailContactInventory] Found $($rows.Count) mail contact(s) and mail user(s)." -ExportFileLocation $exportDetails
    }
    catch {
        $evidenceState = 'Failed'
        $evidenceReason = $_.Exception.Message
        Write-Log -Type ERROR -Message "[Get-ExchangeMailContactInventory] Mail contact collection failed: $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
    }
    finally {
        $tenantStatsHash['MailContacts'] = $rows
        if (Get-Command -Name Set-ArrayaCollectorEvidenceState -ErrorAction SilentlyContinue) {
            Set-ArrayaCollectorEvidenceState -TenantStatsHash $tenantStatsHash -DatasetKey 'MailContacts' -State $evidenceState -Reason $evidenceReason -Collector 'Get-ExchangeMailContactInventory' | Out-Null
        }
        $completedTime = ((Get-Date) - $start).ToString('hh\:mm\:ss')
        Write-ArrayaExchangeCollectorCompletionBanner -Message "[Get-ExchangeMailContactInventory] COMPLETED: Gathering mail contacts and mail users in $completedTime" -ExportFileLocation $exportDetails
    }

    return $tenantStatsHash['MailContacts']
}
