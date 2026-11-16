function New-ArrayaTenantToTenantIdentityInventory {
    <#
    .SYNOPSIS
        Builds the group and guest inventories a tenant-to-tenant migration recreates.

    .DESCRIPTION
        Groups and guests do not move with mailbox or content migration. Every group has to
        be recreated in the destination, and every guest has to be re-invited there. This
        function shapes the collected evidence into two worksheets for that work.

        GroupInventory joins the Entra group inventory (EntraIDGroups) to the Exchange group
        inventory (AllExchangeGroups) so each group appears once. The join uses the Exchange
        ExternalDirectoryObjectId, then the primary SMTP address. Exchange dynamic
        distribution lists exist only in Exchange, so they come from AllExchangeGroups alone.
        Each row states the group type, whether it is mail-enabled, how membership is decided
        (assigned, an Entra membership rule, or an Exchange recipient filter), where the group
        has to be recreated, and what makes recreation harder.

        GroupMembers lists members and owners. Graph supplies security and Microsoft 365
        groups (EntraGroupMembers); Exchange supplies mail-enabled groups
        (ExchangeGroupMembers). A group listed by Graph is not repeated from Exchange.

        GuestUsers lists Entra guest accounts. A guest is a user whose UserType is Guest, or,
        when UserType was not collected, whose UPN carries the #EXT# marker.

        ContactInventory lists Exchange mail contacts and mail users (MailContacts), the
        groups each one belongs to, and whether its external address sits in a domain the
        source tenant owns.

        Each inventory reports an evidence status. A count of zero is reported as Ready only
        when the collectors behind it ran; otherwise it is Needs Data, because an empty table
        from a skipped collector is not evidence that no groups or guests exist.

    .PARAMETER TenantStatsHash
        The assessment store.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$TenantStatsHash
    )

    function Get-InventoryRowSet {
        param([string]$Key)
        if (-not $TenantStatsHash.Contains($Key) -or $null -eq $TenantStatsHash[$Key]) {
            return @()
        }
        $value = $TenantStatsHash[$Key]
        if ($value -is [System.Collections.IDictionary]) {
            return @($value.Values | Where-Object { $null -ne $_ })
        }
        if (($value -is [System.Collections.IEnumerable]) -and -not ($value -is [string])) {
            return @($value | Where-Object { $null -ne $_ })
        }
        return @($value)
    }

    function Get-InventoryText {
        param($Record, [string[]]$Names)
        $value = Get-ArrayaObjectValue -Object $Record -Names $Names
        if ($null -eq $value) {
            return $null
        }
        if (($value -is [System.Collections.IEnumerable]) -and -not ($value -is [string])) {
            $parts = @($value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($parts.Count -eq 0) { return $null }
            return ($parts -join '; ')
        }
        $text = ([string]$value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        return $text
    }

    function Get-InventoryCount {
        param($Value)
        # Collectors write text such as 'NotCollected (minimum mode)' where a count was not
        # gathered. Only a real number is a count.
        if ($null -eq $Value) { return $null }
        if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
            return [int]$Value
        }
        $parsed = 0
        if ([int]::TryParse(([string]$Value).Trim(), [ref]$parsed)) {
            return $parsed
        }
        return $null
    }

    function Get-InventoryKey {
        param($Value)
        if ($null -eq $Value) { return $null }
        $text = ([string]$Value).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        return $text
    }

    function Resolve-EntraGroupType {
        param($EntraGroup)
        $groupType = [string](Get-ArrayaObjectValue -Object $EntraGroup -Names @('GroupType'))
        switch -Regex ($groupType) {
            '^(?i)microsoft 365' { return 'Microsoft 365 group' }
            '^(?i)mail-enabled security' { return 'Mail-enabled security group' }
            '^(?i)security' { return 'Security group' }
            '^(?i)distribution' { return 'Distribution list' }
        }

        $mailEnabled = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $EntraGroup -Names @('MailEnabled'))
        $securityEnabled = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $EntraGroup -Names @('SecurityEnabled'))
        if ($mailEnabled -eq $true -and $securityEnabled -eq $true) { return 'Mail-enabled security group' }
        if ($mailEnabled -eq $false -and $securityEnabled -eq $true) { return 'Security group' }
        if ($mailEnabled -eq $true -and $securityEnabled -eq $false) { return 'Distribution list' }
        return 'Unknown'
    }

    function Resolve-ExchangeGroupType {
        param([string]$RecipientTypeDetails)
        switch ($RecipientTypeDetails) {
            'GroupMailbox' { return 'Microsoft 365 group' }
            'MailUniversalSecurityGroup' { return 'Mail-enabled security group' }
            'DynamicDistributionGroup' { return 'Dynamic distribution list' }
            'MailUniversalDistributionGroup' { return 'Distribution list' }
            'MailNonUniversalGroup' { return 'Distribution list' }
        }
        if ([string]::IsNullOrWhiteSpace($RecipientTypeDetails)) { return 'Unknown' }
        return $RecipientTypeDetails
    }

    $entraGroups = @(Get-InventoryRowSet -Key 'EntraIDGroups')
    $exchangeGroups = @(Get-InventoryRowSet -Key 'AllExchangeGroups')
    $entraEvidence = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'EntraIDGroups' -CollectorStepName 'Entra groups'
    $exchangeEvidence = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'AllExchangeGroups' -CollectorStepName 'Exchange groups'

    # Lookups hold positions in $exchangeGroups so matching does not depend on how
    # PowerShell compares two PSCustomObjects.
    $exchangeByObjectId = @{}
    $exchangeBySmtp = @{}
    for ($index = 0; $index -lt $exchangeGroups.Count; $index++) {
        $exchangeGroup = $exchangeGroups[$index]
        $objectIdKey = Get-InventoryKey (Get-ArrayaObjectValue -Object $exchangeGroup -Names @('ExternalDirectoryObjectId'))
        if ($objectIdKey -and -not $exchangeByObjectId.ContainsKey($objectIdKey)) { $exchangeByObjectId[$objectIdKey] = $index }
        $smtpKey = Get-InventoryKey (Get-ArrayaObjectValue -Object $exchangeGroup -Names @('PrimarySmtpAddress', 'PrimarySMTPAddress'))
        if ($smtpKey -and -not $exchangeBySmtp.ContainsKey($smtpKey)) { $exchangeBySmtp[$smtpKey] = $index }
    }
    $matchedExchangeIndexes = New-Object 'System.Collections.Generic.HashSet[int]'

    $teamSiteUrls = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($teamRow in @(Get-InventoryRowSet -Key 'AllTeams')) {
        $teamSiteKey = Get-InventoryKey (Get-ArrayaObjectValue -Object $teamRow -Names @('SharePointSiteUrl'))
        if ($teamSiteKey) { [void]$teamSiteUrls.Add($teamSiteKey.TrimEnd('/')) }
    }

    $newGroupRow = {
        param($EntraGroup, $ExchangeGroup)

        $exchangeType = [string](Get-ArrayaObjectValue -Object $ExchangeGroup -Names @('RecipientTypeDetails'))
        $groupType = if ($EntraGroup) { Resolve-EntraGroupType -EntraGroup $EntraGroup } else { Resolve-ExchangeGroupType -RecipientTypeDetails $exchangeType }
        if ($groupType -eq 'Unknown' -and $ExchangeGroup) { $groupType = Resolve-ExchangeGroupType -RecipientTypeDetails $exchangeType }
        $isDynamicDistribution = ($exchangeType -eq 'DynamicDistributionGroup')

        $entraMembershipType = [string](Get-ArrayaObjectValue -Object $EntraGroup -Names @('MembershipType'))
        $isEntraDynamic = ($entraMembershipType -match '(?i)dynamic')
        $membershipType = if ($isDynamicDistribution) { 'Dynamic (Exchange recipient filter)' } elseif ($isEntraDynamic) { 'Dynamic (Entra membership rule)' } else { 'Assigned' }
        $membershipRule = if ($isDynamicDistribution) {
            Get-InventoryText -Record $ExchangeGroup -Names @('RecipientFilter')
        }
        elseif ($isEntraDynamic) {
            Get-InventoryText -Record $EntraGroup -Names @('MembershipRule')
        }
        else {
            $null
        }

        $mailEnabled = if ($ExchangeGroup) { $true } else { Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $EntraGroup -Names @('MailEnabled')) }
        $securityEnabled = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $EntraGroup -Names @('SecurityEnabled'))
        if ($null -eq $securityEnabled -and $ExchangeGroup) {
            $securityEnabled = ($exchangeType -eq 'MailUniversalSecurityGroup')
        }

        $entraSource = [string](Get-ArrayaObjectValue -Object $EntraGroup -Names @('Source'))
        $isSynced = (Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $EntraGroup -Names @('OnPremisesSyncEnabled'))) -eq $true -or
            $entraSource -match '(?i)on-prem' -or
            (Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $ExchangeGroup -Names @('IsDirSynced'))) -eq $true

        $isRoleAssignable = (Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $EntraGroup -Names @('IsAssignableToRole'))) -eq $true
        $isManagingLicenses = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $EntraGroup -Names @('IsManagingLicenses'))
        $assignedLicenses = if ($isManagingLicenses -eq $true) { Get-InventoryText -Record $EntraGroup -Names @('AssignedLicenseFriendlyNames', 'AssignedLicenseSkuPartNumbers') } else { $null }
        $resourceOptions = [string](Get-ArrayaObjectValue -Object $ExchangeGroup -Names @('ResourceProvisioningOptions'))
        $siteUrl = Get-InventoryText -Record $ExchangeGroup -Names @('SharePointSiteUrl')
        $isTeam = $null
        if ($groupType -eq 'Microsoft 365 group') {
            if (-not [string]::IsNullOrWhiteSpace($resourceOptions)) {
                $isTeam = $resourceOptions -match '(?i)team'
            }
            elseif ($teamSiteUrls.Count -gt 0) {
                # The reused unified group inventory carries no provisioning options; a Team
                # is recognized by its site in the Teams inventory instead.
                $siteKey = Get-InventoryKey $siteUrl
                $isTeam = [bool]($siteKey -and $teamSiteUrls.Contains($siteKey.TrimEnd('/')))
            }
        }

        $entraMemberCount = Get-InventoryCount (Get-ArrayaObjectValue -Object $EntraGroup -Names @('MemberCount'))
        $exchangeMemberStatus = [string](Get-ArrayaObjectValue -Object $ExchangeGroup -Names @('MemberInventoryStatus'))
        $entraListState = [string](Get-ArrayaObjectValue -Object $EntraGroup -Names @('MemberListState'))
        $exchangeMemberCount = if ($exchangeMemberStatus -in @('Collected', 'CountOnly')) { Get-InventoryCount (Get-ArrayaObjectValue -Object $ExchangeGroup -Names @('MembersCount')) } else { $null }
        $memberCount = if ($null -ne $entraMemberCount) { $entraMemberCount } else { $exchangeMemberCount }
        $memberListStatus = if ($exchangeMemberStatus -eq 'Collected' -or $entraListState -eq 'Collected') {
            if ($isDynamicDistribution) { 'Listed in GroupMembers (computed from the filter at collection time)' }
            elseif ($isEntraDynamic) { 'Listed in GroupMembers (computed from the rule at collection time)' }
            else { 'Listed in GroupMembers' }
        }
        elseif ($entraListState -in @('LookupFailed', 'Partial') -or $exchangeMemberStatus -eq 'LookupFailed') {
            'Lookup failed'
        }
        elseif ($null -ne $memberCount) {
            'Count only'
        }
        else {
            'Not collected'
        }

        $entraOwnerCount = Get-InventoryCount (Get-ArrayaObjectValue -Object $EntraGroup -Names @('OwnerCount'))
        $exchangeOwnerCount = Get-InventoryCount (Get-ArrayaObjectValue -Object $ExchangeGroup -Names @('OwnersCount'))
        $ownerCount = if ($null -ne $entraOwnerCount) { $entraOwnerCount } else { $exchangeOwnerCount }

        $recreateIn = if ($isSynced) {
            'On-premises AD (synced)'
        }
        elseif ($groupType -eq 'Microsoft 365 group') {
            'Microsoft 365 (Entra ID or Exchange Online)'
        }
        elseif ($groupType -eq 'Security group') {
            'Entra ID'
        }
        elseif ($groupType -in @('Mail-enabled security group', 'Distribution list', 'Dynamic distribution list')) {
            'Exchange Online'
        }
        else {
            'Review'
        }

        $considerations = New-Object System.Collections.Generic.List[string]
        if ($isSynced) { $considerations.Add('Synced from on-premises AD; recreate through the destination directory sync design, not in the cloud.') | Out-Null }
        if ($isDynamicDistribution) { $considerations.Add('Recreate with the recipient filter; rewrite any condition that references source-only attributes, domains, or containers.') | Out-Null }
        if ($isEntraDynamic) { $considerations.Add('Recreate the membership rule; confirm every attribute it uses is populated in the destination.') | Out-Null }
        if ($isRoleAssignable) { $considerations.Add('Role-assignable; this can only be set when the group is created.') | Out-Null }
        if ($isManagingLicenses -eq $true) { $considerations.Add('Assigns licenses; the destination SKUs must exist before members are added.') | Out-Null }
        if ($groupType -eq 'Mail-enabled security group') { $considerations.Add('Mail-enabled security groups are created in Exchange Online, not in Entra ID.') | Out-Null }
        if ($groupType -eq 'Microsoft 365 group') { $considerations.Add('Group mailbox, site, and Team content migrate separately from the group object.') | Out-Null }
        if ((Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $ExchangeGroup -Names @('RequireSenderAuthenticationEnabled'))) -eq $false) { $considerations.Add('Accepts mail from external senders.') | Out-Null }
        if ($ExchangeGroup -and -not [string]::IsNullOrWhiteSpace((Get-InventoryText -Record $ExchangeGroup -Names @('AcceptMessagesOnlyFrom', 'AcceptMessagesOnlyFromSendersOrMembers', 'AcceptMessagesOnlyFromDLMembers')))) { $considerations.Add('Restricts who can send to it; re-point the restriction at destination objects.') | Out-Null }
        if ($null -ne $ownerCount -and $ownerCount -eq 0 -and $groupType -ne 'Dynamic distribution list') { $considerations.Add('No owner recorded.') | Out-Null }

        $inventorySource = if ($EntraGroup -and $ExchangeGroup) { 'Entra ID and Exchange Online' } elseif ($EntraGroup) { 'Entra ID' } else { 'Exchange Online' }

        [pscustomobject][ordered]@{
            DisplayName                        = if ($EntraGroup) { Get-InventoryText -Record $EntraGroup -Names @('DisplayName') } else { Get-InventoryText -Record $ExchangeGroup -Names @('DisplayName') }
            GroupType                          = $groupType
            MailEnabled                        = $mailEnabled
            SecurityEnabled                    = $securityEnabled
            MembershipType                     = $membershipType
            MembershipRule                     = $membershipRule
            Source                             = if ($isSynced) { 'Synced from on-premises' } else { 'Cloud' }
            RecreateIn                         = $recreateIn
            PrimarySmtpAddress                 = if ($ExchangeGroup) { Get-InventoryText -Record $ExchangeGroup -Names @('PrimarySmtpAddress', 'PrimarySMTPAddress') } else { Get-InventoryText -Record $EntraGroup -Names @('Mail') }
            Alias                              = Get-InventoryText -Record $ExchangeGroup -Names @('Alias')
            EmailAddresses                     = Get-InventoryText -Record $ExchangeGroup -Names @('EmailAddresses')
            HiddenFromAddressListsEnabled      = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $ExchangeGroup -Names @('HiddenFromAddressListsEnabled'))
            RequireSenderAuthenticationEnabled = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $ExchangeGroup -Names @('RequireSenderAuthenticationEnabled'))
            Visibility                         = if ($EntraGroup -and -not [string]::IsNullOrWhiteSpace((Get-InventoryText -Record $EntraGroup -Names @('Visibility')))) { Get-InventoryText -Record $EntraGroup -Names @('Visibility') } else { Get-InventoryText -Record $ExchangeGroup -Names @('AccessType') }
            IsTeam                             = $isTeam
            SharePointSiteUrl                  = $siteUrl
            Owners                             = $(
                $exchangeOwners = Get-InventoryText -Record $ExchangeGroup -Names @('Owners')
                if ($exchangeOwners) { $exchangeOwners } else { Get-InventoryText -Record $EntraGroup -Names @('Owners') }
            )
            OwnerCount                         = $ownerCount
            MemberCount                        = $memberCount
            MemberList                         = $memberListStatus
            IsAssignableToRole                 = $isRoleAssignable
            AssignsLicenses                    = $isManagingLicenses
            AssignedLicenses                   = $assignedLicenses
            RecreationConsiderations           = if ($considerations.Count -gt 0) { $considerations.ToArray() -join ' ' } else { $null }
            InventorySource                    = $inventorySource
            EntraObjectId                      = if ($EntraGroup) { Get-InventoryText -Record $EntraGroup -Names @('ID', 'Id') } else { Get-InventoryText -Record $ExchangeGroup -Names @('ExternalDirectoryObjectId') }
            TargetDisplayName                  = $null
            TargetPrimarySmtpAddress           = $null
            Wave                               = $null
            Notes                              = $null
        }
    }

    $groupRows = New-Object System.Collections.Generic.List[object]
    foreach ($entraGroup in $entraGroups) {
        $matchIndex = -1
        $entraIdKey = Get-InventoryKey (Get-ArrayaObjectValue -Object $entraGroup -Names @('ID', 'Id'))
        if ($entraIdKey -and $exchangeByObjectId.ContainsKey($entraIdKey)) {
            $matchIndex = [int]$exchangeByObjectId[$entraIdKey]
        }
        if ($matchIndex -lt 0) {
            $mailKey = Get-InventoryKey (Get-ArrayaObjectValue -Object $entraGroup -Names @('Mail'))
            if ($mailKey -and $exchangeBySmtp.ContainsKey($mailKey)) {
                $matchIndex = [int]$exchangeBySmtp[$mailKey]
            }
        }
        $exchangeMatch = $null
        if ($matchIndex -ge 0 -and $matchedExchangeIndexes.Add($matchIndex)) {
            $exchangeMatch = $exchangeGroups[$matchIndex]
        }
        $groupRows.Add((& $newGroupRow $entraGroup $exchangeMatch)) | Out-Null
    }
    for ($index = 0; $index -lt $exchangeGroups.Count; $index++) {
        if ($matchedExchangeIndexes.Contains($index)) {
            continue
        }
        $groupRows.Add((& $newGroupRow $null $exchangeGroups[$index])) | Out-Null
    }
    $sortedGroupRows = @($groupRows.ToArray() | Sort-Object -Property @{ Expression = { [string]$_.GroupType } }, @{ Expression = { [string]$_.DisplayName } })

    # ---- Group membership ---------------------------------------------------------------------
    # Graph lists security and Microsoft 365 groups (members and owners); Exchange lists
    # mail-enabled groups. A group listed by Graph is not repeated from Exchange.
    $groupRowByObjectId = @{}
    $groupRowBySmtp = @{}
    foreach ($groupRow in $sortedGroupRows) {
        $objectIdKey = Get-InventoryKey $groupRow.EntraObjectId
        if ($objectIdKey -and -not $groupRowByObjectId.ContainsKey($objectIdKey)) { $groupRowByObjectId[$objectIdKey] = $groupRow }
        $smtpKey = Get-InventoryKey $groupRow.PrimarySmtpAddress
        if ($smtpKey -and -not $groupRowBySmtp.ContainsKey($smtpKey)) { $groupRowBySmtp[$smtpKey] = $groupRow }
    }

    $groupMemberRows = New-Object System.Collections.Generic.List[object]
    $graphListedGroupIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($entraMember in @(Get-InventoryRowSet -Key 'EntraGroupMembers')) {
        $groupIdKey = Get-InventoryKey (Get-ArrayaObjectValue -Object $entraMember -Names @('GroupId'))
        if (-not $groupIdKey) { continue }
        [void]$graphListedGroupIds.Add($groupIdKey)
        $parentGroup = if ($groupRowByObjectId.ContainsKey($groupIdKey)) { $groupRowByObjectId[$groupIdKey] } else { $null }
        $groupMemberRows.Add([pscustomobject][ordered]@{
            GroupDisplayName        = Get-InventoryText -Record $entraMember -Names @('GroupDisplayName')
            GroupType               = if ($parentGroup) { $parentGroup.GroupType } else { $null }
            GroupPrimarySmtpAddress = if ($parentGroup -and $parentGroup.PrimarySmtpAddress) { $parentGroup.PrimarySmtpAddress } else { Get-InventoryText -Record $entraMember -Names @('GroupMail') }
            GroupObjectId           = Get-InventoryText -Record $entraMember -Names @('GroupId')
            Relationship            = Get-InventoryText -Record $entraMember -Names @('Relationship')
            MemberDisplayName       = Get-InventoryText -Record $entraMember -Names @('MemberDisplayName')
            MemberAddress           = Get-InventoryText -Record $entraMember -Names @('MemberAddress')
            MemberType              = Get-InventoryText -Record $entraMember -Names @('MemberType')
            MembershipSource        = Get-InventoryText -Record $entraMember -Names @('MembershipSource')
            Source                  = 'Entra ID'
        }) | Out-Null
    }
    foreach ($exchangeMember in @(Get-InventoryRowSet -Key 'ExchangeGroupMembers')) {
        $groupSmtpKey = Get-InventoryKey (Get-ArrayaObjectValue -Object $exchangeMember -Names @('GroupPrimarySmtpAddress'))
        $parentGroup = if ($groupSmtpKey -and $groupRowBySmtp.ContainsKey($groupSmtpKey)) { $groupRowBySmtp[$groupSmtpKey] } else { $null }
        $parentObjectIdKey = if ($parentGroup) { Get-InventoryKey $parentGroup.EntraObjectId } else { $null }
        if ($parentObjectIdKey -and $graphListedGroupIds.Contains($parentObjectIdKey)) {
            continue
        }
        $groupMemberRows.Add([pscustomobject][ordered]@{
            GroupDisplayName        = Get-InventoryText -Record $exchangeMember -Names @('GroupDisplayName')
            GroupType               = if ($parentGroup) { $parentGroup.GroupType } else { Resolve-ExchangeGroupType -RecipientTypeDetails ([string](Get-ArrayaObjectValue -Object $exchangeMember -Names @('GroupRecipientTypeDetails'))) }
            GroupPrimarySmtpAddress = Get-InventoryText -Record $exchangeMember -Names @('GroupPrimarySmtpAddress')
            GroupObjectId           = if ($parentGroup) { $parentGroup.EntraObjectId } else { $null }
            Relationship            = 'Member'
            MemberDisplayName       = Get-InventoryText -Record $exchangeMember -Names @('MemberDisplayName')
            MemberAddress           = Get-InventoryText -Record $exchangeMember -Names @('MemberPrimarySmtpAddress')
            MemberType              = Get-InventoryText -Record $exchangeMember -Names @('MemberRecipientTypeDetails')
            MembershipSource        = Get-InventoryText -Record $exchangeMember -Names @('MembershipSource')
            Source                  = 'Exchange Online'
        }) | Out-Null
    }
    $sortedGroupMemberRows = @($groupMemberRows.ToArray() | Sort-Object -Property @{ Expression = { [string]$_.GroupDisplayName } }, @{ Expression = { [string]$_.Relationship }; Descending = $true }, @{ Expression = { [string]$_.MemberDisplayName } })

    # ---- Mail contacts and mail users ---------------------------------------------------------
    $groupsByMemberAddress = @{}
    foreach ($memberRow in $sortedGroupMemberRows) {
        if ($memberRow.Relationship -ne 'Member') { continue }
        $addressKey = Get-InventoryKey $memberRow.MemberAddress
        if (-not $addressKey) { continue }
        if (-not $groupsByMemberAddress.ContainsKey($addressKey)) {
            $groupsByMemberAddress[$addressKey] = New-Object System.Collections.Generic.List[string]
        }
        if (-not $groupsByMemberAddress[$addressKey].Contains([string]$memberRow.GroupDisplayName)) {
            $groupsByMemberAddress[$addressKey].Add([string]$memberRow.GroupDisplayName) | Out-Null
        }
    }
    $ownedDomains = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($domainRow in @(Get-InventoryRowSet -Key 'Domains')) {
        $domainKey = Get-InventoryKey (Get-ArrayaObjectValue -Object $domainRow -Names @('Name', 'Domain', 'Id'))
        if ($domainKey) { [void]$ownedDomains.Add($domainKey) }
    }

    $contactRows = New-Object System.Collections.Generic.List[object]
    foreach ($contact in @(Get-InventoryRowSet -Key 'MailContacts')) {
        $recipientType = [string](Get-ArrayaObjectValue -Object $contact -Names @('RecipientTypeDetails'))
        $contactType = switch ($recipientType) {
            'MailContact' { 'Mail contact' }
            'MailUser' { 'Mail user' }
            default { if ([string]::IsNullOrWhiteSpace($recipientType)) { 'Unknown' } else { $recipientType } }
        }
        $externalAddress = Get-InventoryText -Record $contact -Names @('ExternalEmailAddress')
        $primaryAddress = Get-InventoryText -Record $contact -Names @('PrimarySmtpAddress')
        $isSynced = (Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $contact -Names @('IsDirSynced'))) -eq $true

        $memberOf = New-Object System.Collections.Generic.List[string]
        foreach ($candidate in @($primaryAddress, $externalAddress)) {
            $candidateKey = Get-InventoryKey $candidate
            if ($candidateKey -and $groupsByMemberAddress.ContainsKey($candidateKey)) {
                foreach ($groupName in $groupsByMemberAddress[$candidateKey]) {
                    if (-not $memberOf.Contains($groupName)) { $memberOf.Add($groupName) | Out-Null }
                }
            }
        }

        $externalDomain = if ($externalAddress -and $externalAddress -match '@(.+)$') { $Matches[1].ToLowerInvariant() } else { $null }
        $considerations = New-Object System.Collections.Generic.List[string]
        if ($contactType -eq 'Mail user') { $considerations.Add('A mail user is a sign-in account with an external address; recreate it as a user or mail user, not as a contact.') | Out-Null }
        if ($isSynced) { $considerations.Add('Synced from on-premises AD; recreate through the destination directory sync design.') | Out-Null }
        if ($externalDomain -and $ownedDomains.Contains($externalDomain)) { $considerations.Add(('The external address is in {0}, a domain this tenant owns; it will point somewhere else once that domain moves.' -f $externalDomain)) | Out-Null }
        if ($memberOf.Count -gt 0) { $considerations.Add(('Member of {0} group(s); create it before those groups are populated.' -f $memberOf.Count)) | Out-Null }

        $contactRows.Add([pscustomobject][ordered]@{
            DisplayName                   = Get-InventoryText -Record $contact -Names @('DisplayName')
            ContactType                   = $contactType
            ExternalEmailAddress          = $externalAddress
            PrimarySmtpAddress            = $primaryAddress
            Alias                         = Get-InventoryText -Record $contact -Names @('Alias')
            EmailAddresses                = Get-InventoryText -Record $contact -Names @('EmailAddresses')
            HiddenFromAddressListsEnabled = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $contact -Names @('HiddenFromAddressListsEnabled'))
            Source                        = if ($isSynced) { 'Synced from on-premises' } else { 'Cloud' }
            RecreateIn                    = if ($isSynced) { 'On-premises AD (synced)' } else { 'Exchange Online' }
            MemberOfGroups                = if ($memberOf.Count -gt 0) { $memberOf.ToArray() -join '; ' } else { $null }
            RecreationConsiderations      = if ($considerations.Count -gt 0) { $considerations.ToArray() -join ' ' } else { $null }
            TargetDisplayName             = $null
            TargetPrimarySmtpAddress      = $null
            Wave                          = $null
            Notes                         = $null
        }) | Out-Null
    }
    $sortedContactRows = @($contactRows.ToArray() | Sort-Object -Property @{ Expression = { [string]$_.ContactType } }, @{ Expression = { [string]$_.DisplayName } })
    $contactEvidence = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'MailContacts' -CollectorStepName 'Exchange mail contacts'
    $contactStatus = if ($sortedContactRows.Count -gt 0) { 'Review' } elseif ($contactEvidence.State -eq 'Collected') { 'Ready' } else { 'Needs Data' }

    # ---- Guests ---------------------------------------------------------------------------
    $users = @(Get-InventoryRowSet -Key 'Users')
    $usersEvidence = Get-ArrayaDatasetEvidenceState -TenantStatsHash $TenantStatsHash -DatasetKey 'Users' -CollectorStepName 'Users'
    $guestRows = New-Object System.Collections.Generic.List[object]
    foreach ($user in $users) {
        $userType = [string](Get-ArrayaObjectValue -Object $user -Names @('UserType'))
        $userPrincipalName = [string](Get-ArrayaObjectValue -Object $user -Names @('UserPrincipalName'))
        $isGuest = if (-not [string]::IsNullOrWhiteSpace($userType)) { $userType -match '^(?i)guest' } else { $userPrincipalName -like '*#EXT#*' }
        if (-not $isGuest) {
            continue
        }

        # The export shaping falls back to the UPN when Mail is empty; a guest UPN is a
        # resource-tenant address, not the guest's own mailbox.
        $mail = Get-InventoryText -Record $user -Names @('Mail')
        if ($mail -and $mail -like '*#EXT#*') { $mail = $null }
        $homeDomain = $null
        if ($mail -and $mail -match '@(.+)$') {
            $homeDomain = $Matches[1].ToLowerInvariant()
        }
        elseif ($userPrincipalName -match '^(.+)#EXT#') {
            $localPart = $Matches[1]
            $separator = $localPart.LastIndexOf('_')
            if ($separator -ge 0 -and $separator -lt ($localPart.Length - 1)) {
                $homeDomain = $localPart.Substring($separator + 1).ToLowerInvariant()
            }
        }

        $lastSignIn = Get-ArrayaObjectValue -Object $user -Names @('LastSignInDateTime')
        if ($null -eq $lastSignIn) {
            $signInActivity = Get-ArrayaObjectValue -Object $user -Names @('SignInActivity')
            $lastSignIn = Get-ArrayaObjectValue -Object $signInActivity -Names @('LastSignInDateTime', 'lastSignInDateTime')
        }

        $guestRows.Add([pscustomobject][ordered]@{
            DisplayName        = Get-InventoryText -Record $user -Names @('DisplayName')
            Mail               = $mail
            UserPrincipalName  = $userPrincipalName
            HomeDomain         = $homeDomain
            AccountEnabled     = Convert-ArrayaToBoolean -Value (Get-ArrayaObjectValue -Object $user -Names @('AccountEnabled'))
            ExternalUserState  = Get-InventoryText -Record $user -Names @('ExternalUserState')
            CreationType       = Get-InventoryText -Record $user -Names @('CreationType')
            CreatedDateTime    = Get-ArrayaObjectValue -Object $user -Names @('CreatedDateTime')
            LastSignInDateTime = $lastSignIn
            CompanyName        = Get-InventoryText -Record $user -Names @('CompanyName')
            ReinviteDecision   = $null
            Notes              = $null
        }) | Out-Null
    }
    $sortedGuestRows = @($guestRows.ToArray() | Sort-Object -Property @{ Expression = { [string]$_.HomeDomain } }, @{ Expression = { [string]$_.DisplayName } })

    # ---- Evidence ---------------------------------------------------------------------------
    $groupGaps = New-Object System.Collections.Generic.List[string]
    if ($entraEvidence.State -ne 'Collected') { $groupGaps.Add(('Entra groups: {0}' -f $entraEvidence.Reason)) | Out-Null }
    if ($exchangeEvidence.State -ne 'Collected') { $groupGaps.Add(('Exchange groups: {0}' -f $exchangeEvidence.Reason)) | Out-Null }
    $groupStatus = if ($sortedGroupRows.Count -gt 0) { 'Review' } elseif ($groupGaps.Count -eq 0) { 'Ready' } else { 'Needs Data' }
    $guestStatus = if ($sortedGuestRows.Count -gt 0) { 'Review' } elseif ($usersEvidence.State -eq 'Collected') { 'Ready' } else { 'Needs Data' }

    $countByType = [ordered]@{}
    foreach ($typeName in @('Security group', 'Mail-enabled security group', 'Distribution list', 'Dynamic distribution list', 'Microsoft 365 group')) {
        $countByType[$typeName] = @($sortedGroupRows | Where-Object { $_.GroupType -eq $typeName }).Count
    }
    $otherTypeCount = @($sortedGroupRows | Where-Object { -not $countByType.Contains([string]$_.GroupType) }).Count
    if ($otherTypeCount -gt 0) { $countByType['Other'] = $otherTypeCount }

    return [pscustomobject][ordered]@{
        GroupInventory          = $sortedGroupRows
        GroupMembers            = $sortedGroupMemberRows
        GuestUsers              = $sortedGuestRows
        ContactInventory        = $sortedContactRows
        MailContactCount        = @($sortedContactRows | Where-Object { $_.ContactType -eq 'Mail contact' }).Count
        MailUserCount           = @($sortedContactRows | Where-Object { $_.ContactType -eq 'Mail user' }).Count
        ContactStatus           = $contactStatus
        ContactEvidenceGap      = if ($contactEvidence.State -ne 'Collected') { $contactEvidence.Reason } else { $null }
        GroupCountByType        = $countByType
        MailEnabledGroupCount   = @($sortedGroupRows | Where-Object { $_.MailEnabled -eq $true }).Count
        DynamicGroupCount       = @($sortedGroupRows | Where-Object { $_.MembershipType -ne 'Assigned' }).Count
        SyncedGroupCount        = @($sortedGroupRows | Where-Object { $_.Source -ne 'Cloud' }).Count
        RoleAssignableCount     = @($sortedGroupRows | Where-Object { $_.IsAssignableToRole -eq $true }).Count
        LicenseGroupCount       = @($sortedGroupRows | Where-Object { $_.AssignsLicenses -eq $true }).Count
        GroupStatus             = $groupStatus
        GroupEvidenceGap        = if ($groupGaps.Count -gt 0) { $groupGaps.ToArray() -join ' ' } else { $null }
        GuestStatus             = $guestStatus
        GuestEvidenceGap        = if ($usersEvidence.State -ne 'Collected') { $usersEvidence.Reason } else { $null }
        EnabledGuestCount       = @($sortedGuestRows | Where-Object { $_.AccountEnabled -ne $false }).Count
        PendingGuestCount       = @($sortedGuestRows | Where-Object { [string]$_.ExternalUserState -match '(?i)pending' }).Count
    }
}
