function Export-ArrayaTenantToTenantCutoverPack {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory = $true)]
        [string]$BaseExportPath
    )

    function Get-CutoverPackContainerValue {
        [CmdletBinding()]
        param(
            [AllowNull()]$Container,
            [Parameter(Mandatory = $true)]
            [string]$Name
        )

        if ($null -eq $Container -or [string]::IsNullOrWhiteSpace($Name)) {
            return $null
        }

        if ($Container -is [System.Collections.IDictionary] -and $Container.Contains($Name)) {
            return $Container[$Name]
        }

        if ($Container.PSObject -and $Container.PSObject.Properties[$Name]) {
            return $Container.PSObject.Properties[$Name].Value
        }

        return $null
    }

    function ConvertTo-CutoverPackArray {
        [CmdletBinding()]
        param(
            [AllowNull()]$Value
        )

        if ($null -eq $Value) {
            return @()
        }

        if ($Value -is [System.Collections.IDictionary]) {
            return @($Value.Values)
        }

        if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
            return @($Value)
        }

        return @($Value)
    }

    function Get-CutoverPackArtifactPrefix {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$ResolvedBasePath
        )

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ResolvedBasePath)
        if ([string]::IsNullOrWhiteSpace($baseName)) {
            return $null
        }

        foreach ($suffix in @('-Assess', ' - Tenant Details', '-Tenant Details')) {
            if ($baseName.EndsWith($suffix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $baseName = $baseName.Substring(0, $baseName.Length - $suffix.Length)
                break
            }
        }

        $baseName = ($baseName -replace '\s+', ' ').Trim()
        if (
            [string]::IsNullOrWhiteSpace($baseName) -or
            [string]::Equals($baseName, 'Assess', [System.StringComparison]::OrdinalIgnoreCase) -or
            [string]::Equals($baseName, 'Tenant Details', [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            return $null
        }

        return $baseName
    }

    function ConvertTo-CutoverPackKey {
        [CmdletBinding()]
        param(
            [AllowNull()]$Value
        )

        if ($null -eq $Value) {
            return $null
        }

        $text = ([string]$Value).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        return $text
    }

    function ConvertTo-OneDriveLookupKey {
        [CmdletBinding()]
        param(
            [AllowNull()]$Value
        )

        if ($null -eq $Value) {
            return $null
        }

        $text = ([string]$Value).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        return (($text -replace '[^a-z0-9]', '_') -replace '_{2,}', '_').Trim('_')
    }

    function ConvertTo-BoolValue {
        [CmdletBinding()]
        param(
            [AllowNull()]$Value
        )

        if ($null -eq $Value) {
            return $false
        }

        if ($Value -is [bool]) {
            return [bool]$Value
        }

        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $false
        }

        return $text -match '^(?i:true|1|yes|enabled)$'
    }

    function Get-CutoverPackFieldValue {
        [CmdletBinding()]
        param(
            [AllowNull()]
            [object]$Record,
            [Parameter(Mandatory = $true)]
            [string[]]$Names
        )

        if ($null -eq $Record) {
            return $null
        }

        return Get-ArrayaObjectValue -Object $Record -Names $Names
    }

    function Get-CutoverPackTextValue {
        [CmdletBinding()]
        param(
            [AllowNull()]
            [object]$Record,
            [Parameter(Mandatory = $true)]
            [string[]]$Names
        )

        $value = Get-CutoverPackFieldValue -Record $Record -Names $Names
        if ($null -eq $value) {
            return $null
        }

        $text = ([string]$value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        return $text
    }

    function ConvertTo-CutoverPackProxyArray {
        [CmdletBinding()]
        param(
            [AllowNull()]$Value
        )

        if ($null -eq $Value) {
            return @()
        }

        if ($Value -is [string]) {
            return @(
                $Value -split '[,;]' |
                    ForEach-Object { ([string]$_).Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }

        if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
            return @(
                $Value |
                    ForEach-Object { ([string]$_).Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }

        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return @()
        }

        return @($text)
    }

    function Get-CutoverPackProxySummary {
        [CmdletBinding()]
        param(
            [AllowNull()]
            [object]$Record
        )

        if ($null -eq $Record) {
            return [pscustomobject]@{
                OnMicrosoftAlias   = $null
                OnMicrosoftAliases = $null
            }
        }

        $resolvedOnMicrosoftAlias = Get-CutoverPackTextValue -Record $Record -Names @('OnMicrosoftAlias')
        $resolvedOnMicrosoftAliases = Get-CutoverPackTextValue -Record $Record -Names @('OnMicrosoftAliases')

        $proxyAddresses = ConvertTo-CutoverPackProxyArray -Value (Get-CutoverPackFieldValue -Record $Record -Names @('EmailAddresses'))
        $onMicrosoftAliases = New-Object System.Collections.Generic.List[string]
        foreach ($proxyAddress in $proxyAddresses) {
            $proxyText = ([string]$proxyAddress).Trim()
            if ($proxyText -match '^(?i)smtp:(?<address>[^@]+@[^@]+\.onmicrosoft\.com)$') {
                $address = [string]$Matches['address']
                if (-not $onMicrosoftAliases.Contains($address)) {
                    $onMicrosoftAliases.Add($address) | Out-Null
                }
            }
        }

        if ([string]::IsNullOrWhiteSpace($resolvedOnMicrosoftAliases) -and $onMicrosoftAliases.Count -gt 0) {
            $resolvedOnMicrosoftAliases = @($onMicrosoftAliases.ToArray()) -join ';'
        }

        if ([string]::IsNullOrWhiteSpace($resolvedOnMicrosoftAlias) -and -not [string]::IsNullOrWhiteSpace($resolvedOnMicrosoftAliases)) {
            $resolvedOnMicrosoftAlias = @(
                $resolvedOnMicrosoftAliases -split ';' |
                    ForEach-Object { ([string]$_).Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Select-Object -First 1
            ) | Select-Object -First 1
        }

        return [pscustomobject]@{
            OnMicrosoftAlias   = $resolvedOnMicrosoftAlias
            OnMicrosoftAliases = $resolvedOnMicrosoftAliases
        }
    }

    function New-CutoverPackPlaceholderRow {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string[]]$Columns
        )

        $row = [ordered]@{}
        foreach ($column in $Columns) {
            $row[$column] = $null
        }

        return [pscustomobject]$row
    }

    function Resolve-CutoverPackOneDriveUrl {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.IDictionary]$Lookup,
            [Parameter(Mandatory = $true)]
            [AllowNull()]
            [AllowEmptyString()]
            [string[]]$CandidateKeys
        )

        foreach ($candidateKey in @($CandidateKeys)) {
            if ([string]::IsNullOrWhiteSpace($candidateKey)) {
                continue
            }

            if ($Lookup.Contains($candidateKey)) {
                return [string]$Lookup[$candidateKey]
            }
        }

        return $null
    }

    function Get-CutoverPackColumnWidth {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$ColumnName
        )

        if ($ColumnName -match '^(MailboxSizeGB|DeletedItemsGB|ArchiveSizeGB|ArchiveDeletedItemsGB|TotalDataToMigrateGB|BitTitanLicenseCount|Count|Wave|HasMailbox|IsDirSynced|HasOneDrive|DeliverToMailboxAndForward|IsInactiveMailbox|LitigationHoldEnabled|ArchiveStatus)$') {
            return 14
        }

        if ($ColumnName -match '^(DisplayName|UserPrincipalName|PrimarySmtpAddress|MailboxPrimarySmtpAddress|MailboxUserPrincipalName|SourceAccount|SourceUPN|SourcePrimarySmtpAddress|TargetAccount|TargetUPN|TargetPrimarySmtpAddress|TargetMail|SourceEmail|SourceOneDriveUrl|TargetOneDriveUrl|TargetFolderName|Direction|ApprovalStatus|MailboxDisplayName|ForwardingAddress|ForwardingSmtpAddress)$') {
            return 26
        }

        if ($ColumnName -match '^(OnMicrosoftAliases|LegacyExchangeDn|LegacyExchangeDnX500|X500Addresses|X400Addresses|FullAccessDelegates|SendAsDelegates|CalendarDelegates|GrantSendOnBehalfTo|PermissionTarget|PermissionTargetDisplayName|AccessRights|SharingPermissionFlags|Notes)$') {
            return 34
        }

        return 20
    }

    function Set-CutoverPackWorksheetColumnWidths {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            $ExcelPackage,
            [Parameter(Mandatory = $true)]
            [string]$WorksheetName,
            [Parameter(Mandatory = $true)]
            [string[]]$Columns
        )

        $worksheet = $ExcelPackage.Workbook.Worksheets[$WorksheetName]
        if ($null -eq $worksheet) {
            return
        }

        for ($columnIndex = 1; $columnIndex -le $Columns.Count; $columnIndex++) {
            $worksheet.Column($columnIndex).Width = Get-CutoverPackColumnWidth -ColumnName $Columns[$columnIndex - 1]
            $worksheet.Column($columnIndex).BestFit = $false
        }
    }

    if (-not (Get-Command -Name Open-ExcelPackage -ErrorAction SilentlyContinue)) {
        throw 'ImportExcel is required to export the tenant-to-tenant cutover pack.'
    }

    $resolvedBasePath = [System.IO.Path]::GetFullPath($BaseExportPath)
    $baseDirectory = Split-Path -Path $resolvedBasePath -Parent
    if ([string]::IsNullOrWhiteSpace($baseDirectory)) {
        $baseDirectory = (Get-Location).Path
    }

    $supportDirectory = Join-Path -Path $baseDirectory -ChildPath 'Support'
    if (-not (Test-Path -Path $supportDirectory)) {
        $null = New-Item -ItemType Directory -Path $supportDirectory -Force
    }

    $cutoverPackDirectory = Join-Path -Path $supportDirectory -ChildPath 'T2T-CutoverPack'
    if (-not (Test-Path -Path $cutoverPackDirectory)) {
        $null = New-Item -ItemType Directory -Path $cutoverPackDirectory -Force
    }

    $artifactPrefix = Get-CutoverPackArtifactPrefix -ResolvedBasePath $resolvedBasePath
    $workbookFileName = if ([string]::IsNullOrWhiteSpace($artifactPrefix)) {
        'T2T-CutoverPack.xlsx'
    }
    else {
        '{0}-T2T-CutoverPack.xlsx' -f $artifactPrefix
    }
    $workbookPath = Join-Path -Path $cutoverPackDirectory -ChildPath $workbookFileName

    $tenantInfo = Get-CutoverPackContainerValue -Container $TenantStatsHash -Name 'TenantInfo'
    $tenantDisplayName = Get-CutoverPackTextValue -Record $tenantInfo -Names @('DisplayName')
    if ([string]::IsNullOrWhiteSpace($tenantDisplayName)) {
        $tenantDisplayName = 'Tenant'
    }

    $users = ConvertTo-CutoverPackArray -Value (Get-CutoverPackContainerValue -Container $TenantStatsHash -Name 'Users')
    $recipients = ConvertTo-CutoverPackArray -Value (Get-CutoverPackContainerValue -Container $TenantStatsHash -Name 'AllRecipients')
    $mailboxes = ConvertTo-CutoverPackArray -Value (Get-CutoverPackContainerValue -Container $TenantStatsHash -Name 'MailboxFullDetails')
    if ($mailboxes.Count -eq 0) {
        $mailboxes = ConvertTo-CutoverPackArray -Value (Get-CutoverPackContainerValue -Container $TenantStatsHash -Name 'AllMailboxes')
    }
    $calendarDelegatePermissions = ConvertTo-CutoverPackArray -Value (Get-CutoverPackContainerValue -Container $TenantStatsHash -Name 'MailboxCalendarDelegatePermissions')
    $oneDriveSites = ConvertTo-CutoverPackArray -Value (Get-CutoverPackContainerValue -Container $TenantStatsHash -Name 'OneDrive')

    $mailboxLookup = @{}
    foreach ($mailbox in @($mailboxes)) {
        foreach ($candidateKey in @(
                (ConvertTo-CutoverPackKey (Get-CutoverPackFieldValue -Record $mailbox -Names @('UserPrincipalName'))),
                (ConvertTo-CutoverPackKey (Get-CutoverPackFieldValue -Record $mailbox -Names @('PrimarySmtpAddress'))),
                (ConvertTo-CutoverPackKey (Get-CutoverPackFieldValue -Record $mailbox -Names @('WindowsEmailAddress'))),
                (ConvertTo-CutoverPackKey (Get-CutoverPackFieldValue -Record $mailbox -Names @('ExternalDirectoryObjectId')))
            )) {
            if ([string]::IsNullOrWhiteSpace($candidateKey)) {
                continue
            }
            if (-not $mailboxLookup.ContainsKey($candidateKey)) {
                $mailboxLookup[$candidateKey] = $mailbox
            }
        }
    }

    $oneDriveLookup = @{}
    foreach ($oneDriveSite in @($oneDriveSites)) {
        $siteUrl = Get-CutoverPackTextValue -Record $oneDriveSite -Names @('Url', 'WebUrl', 'SiteUrl')
        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            continue
        }

        foreach ($candidateKey in @(
                (ConvertTo-CutoverPackKey (Get-CutoverPackFieldValue -Record $oneDriveSite -Names @('Owner', 'OwnerEmail', 'OwnerPrincipalName', 'UserPrincipalName', 'Email'))),
                (ConvertTo-OneDriveLookupKey (Get-CutoverPackFieldValue -Record $oneDriveSite -Names @('Owner', 'OwnerEmail', 'OwnerPrincipalName', 'UserPrincipalName', 'Email')))
            )) {
            if ([string]::IsNullOrWhiteSpace($candidateKey)) {
                continue
            }
            if (-not $oneDriveLookup.ContainsKey($candidateKey)) {
                $oneDriveLookup[$candidateKey] = $siteUrl
            }
        }

        try {
            $uri = [System.Uri]$siteUrl
            $lastSegment = ($uri.AbsolutePath.Trim('/') -split '/')[-1]
            $segmentKey = ConvertTo-OneDriveLookupKey -Value $lastSegment
            if (-not [string]::IsNullOrWhiteSpace($segmentKey) -and -not $oneDriveLookup.ContainsKey($segmentKey)) {
                $oneDriveLookup[$segmentKey] = $siteUrl
            }
        }
        catch {}
    }

    $userWaveColumns = @('SourceAccount', 'SourceUPN', 'SourcePrimarySmtpAddress', 'MailboxPrimarySmtpAddress', 'RecipientTypeDetails', 'HasMailbox', 'IsDirSynced', 'TargetAccount', 'TargetUPN', 'TargetPrimarySmtpAddress', 'TargetMail', 'Wave', 'MigrationState', 'CutoverDate', 'Notes')
    $mailboxReferenceColumns = @('DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'ExchangeGuid', 'ArchiveGuid', 'OnMicrosoftAlias', 'OnMicrosoftAliases', 'LegacyExchangeDn', 'LegacyExchangeDnX500', 'X500Addresses', 'X400Addresses', 'MailboxSizeGB', 'DeletedItemsGB', 'ArchiveStatus', 'ArchiveSizeGB', 'ArchiveDeletedItemsGB', 'TotalDataToMigrateGB', 'BitTitanLicenseType', 'BitTitanLicenseCount', 'FullAccessDelegates', 'FullAccessDelegateCount', 'FullAccessDelegateState', 'SendAsDelegates', 'SendAsDelegateCount', 'SendAsDelegateState', 'CalendarDelegates', 'CalendarDelegateCount', 'CalendarPermissionEntryCount', 'CalendarDelegateState', 'GrantSendOnBehalfTo', 'GrantSendOnBehalfToCount', 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward', 'IsInactiveMailbox', 'LitigationHoldEnabled', 'RetentionPolicy', 'Wave', 'Notes')
    $calendarDelegateColumns = @('MailboxDisplayName', 'MailboxPrimarySmtpAddress', 'MailboxUserPrincipalName', 'RecipientTypeDetails', 'CalendarName', 'CalendarPath', 'PermissionTarget', 'PermissionTargetDisplayName', 'PermissionTargetType', 'AccessRights', 'SharingPermissionFlags', 'Wave', 'Notes')
    $recipientReferenceColumns = @('DisplayName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'UserPrincipalName', 'WindowsEmailAddress', 'EmailAddresses', 'OnMicrosoftAlias', 'OnMicrosoftAliases', 'HiddenFromAddressListsEnabled', 'Wave', 'Notes')
    $collaborationColumns = @('SourceAccount', 'TargetAccount', 'SourceUPN', 'TargetUPN', 'SourcePrimarySmtpAddress', 'Direction', 'IncludeEntraGroups', 'IncludeTeams', 'IncludeSharePoint', 'ApprovalStatus', 'ApprovalNotes', 'Wave', 'Notes')
    $oneDriveColumns = @('SourceUPN', 'SourceEmail', 'SourceOneDriveUrl', 'HasOneDrive', 'TargetMail', 'TargetUPN', 'TargetOneDriveUrl', 'TargetFolderName', 'Wave', 'Notes')

    $userWaveRows = New-Object System.Collections.Generic.List[object]
    foreach ($user in @($users)) {
        $userPrincipalName = Get-CutoverPackTextValue -Record $user -Names @('UserPrincipalName')
        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            continue
        }

        $userType = Get-CutoverPackTextValue -Record $user -Names @('UserType')
        if (-not [string]::IsNullOrWhiteSpace($userType) -and $userType -match '^(?i)guest$') {
            continue
        }

        $accountEnabled = Get-CutoverPackFieldValue -Record $user -Names @('AccountEnabled')
        if ($null -ne $accountEnabled -and -not (ConvertTo-BoolValue -Value $accountEnabled)) {
            continue
        }

        $mailbox = $null
        foreach ($candidateKey in @(
                (ConvertTo-CutoverPackKey $userPrincipalName),
                (ConvertTo-CutoverPackKey (Get-CutoverPackFieldValue -Record $user -Names @('Mail'))),
                (ConvertTo-CutoverPackKey (Get-CutoverPackFieldValue -Record $user -Names @('PrimarySmtpAddress'))),
                (ConvertTo-CutoverPackKey (Get-CutoverPackFieldValue -Record $user -Names @('Id', 'ExternalDirectoryObjectId')))
            )) {
            if ([string]::IsNullOrWhiteSpace($candidateKey)) {
                continue
            }
            if ($mailboxLookup.ContainsKey($candidateKey)) {
                $mailbox = $mailboxLookup[$candidateKey]
                break
            }
        }

        $userMail = Get-CutoverPackTextValue -Record $user -Names @('Mail', 'PrimarySmtpAddress')
        $mailboxPrimarySmtp = Get-CutoverPackTextValue -Record $mailbox -Names @('PrimarySmtpAddress', 'WindowsEmailAddress')
        $recipientTypeDetails = Get-CutoverPackTextValue -Record $mailbox -Names @('RecipientTypeDetails')
        $hasMailbox = $null -ne $mailbox
        $isDirSynced = Get-CutoverPackFieldValue -Record $user -Names @('OnPremisesSyncEnabled', 'IsDirSynced')
        if ($null -eq $isDirSynced -and $mailbox) {
            $isDirSynced = Get-CutoverPackFieldValue -Record $mailbox -Names @('IsDirSynced')
        }

        $userWaveRows.Add([pscustomobject]@{
            SourceAccount           = $userPrincipalName
            SourceUPN               = $userPrincipalName
            SourcePrimarySmtpAddress = $userMail
            MailboxPrimarySmtpAddress = $mailboxPrimarySmtp
            RecipientTypeDetails    = $recipientTypeDetails
            HasMailbox              = $hasMailbox
            IsDirSynced             = if ($null -eq $isDirSynced) { $null } else { [bool](ConvertTo-BoolValue -Value $isDirSynced) }
            TargetAccount           = $null
            TargetUPN               = $null
            TargetPrimarySmtpAddress = $null
            TargetMail              = $null
            Wave                    = $null
            MigrationState          = $null
            CutoverDate             = $null
            Notes                   = $null
        }) | Out-Null
    }

    $mailboxReferenceRows = New-Object System.Collections.Generic.List[object]
    foreach ($mailbox in @($mailboxes)) {
        $recipientTypeDetails = Get-CutoverPackTextValue -Record $mailbox -Names @('RecipientTypeDetails')
        if ([string]::Equals($recipientTypeDetails, 'GroupMailbox', [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $proxySummary = Get-CutoverPackProxySummary -Record $mailbox
        $mailboxReferenceRows.Add([pscustomobject]@{
            DisplayName                = Get-CutoverPackTextValue -Record $mailbox -Names @('DisplayName')
            UserPrincipalName          = Get-CutoverPackTextValue -Record $mailbox -Names @('UserPrincipalName')
            PrimarySmtpAddress         = Get-CutoverPackTextValue -Record $mailbox -Names @('PrimarySmtpAddress')
            RecipientTypeDetails       = $recipientTypeDetails
            ExchangeGuid               = Get-CutoverPackTextValue -Record $mailbox -Names @('ExchangeGuid')
            ArchiveGuid                = Get-CutoverPackTextValue -Record $mailbox -Names @('ArchiveGuid')
            OnMicrosoftAlias           = $proxySummary.OnMicrosoftAlias
            OnMicrosoftAliases         = $proxySummary.OnMicrosoftAliases
            LegacyExchangeDn           = Get-CutoverPackTextValue -Record $mailbox -Names @('LegacyExchangeDn', 'LegacyExchangeDN')
            LegacyExchangeDnX500       = Get-CutoverPackTextValue -Record $mailbox -Names @('LegacyExchangeDnX500')
            X500Addresses              = Get-CutoverPackTextValue -Record $mailbox -Names @('X500Addresses')
            X400Addresses              = Get-CutoverPackTextValue -Record $mailbox -Names @('X400Addresses')
            MailboxSizeGB              = Get-CutoverPackFieldValue -Record $mailbox -Names @('MailboxSizeGB', 'MBXSizeGB')
            DeletedItemsGB             = Get-CutoverPackFieldValue -Record $mailbox -Names @('DeletedItemsGB')
            FullAccessDelegates        = Get-CutoverPackTextValue -Record $mailbox -Names @('FullAccessDelegates')
            FullAccessDelegateCount    = Get-CutoverPackFieldValue -Record $mailbox -Names @('FullAccessDelegateCount')
            FullAccessDelegateState    = Get-CutoverPackTextValue -Record $mailbox -Names @('FullAccessDelegateState')
            SendAsDelegates            = Get-CutoverPackTextValue -Record $mailbox -Names @('SendAsDelegates')
            SendAsDelegateCount        = Get-CutoverPackFieldValue -Record $mailbox -Names @('SendAsDelegateCount')
            SendAsDelegateState        = Get-CutoverPackTextValue -Record $mailbox -Names @('SendAsDelegateState')
            CalendarDelegates          = Get-CutoverPackTextValue -Record $mailbox -Names @('CalendarDelegates')
            CalendarDelegateCount      = Get-CutoverPackFieldValue -Record $mailbox -Names @('CalendarDelegateCount')
            CalendarPermissionEntryCount = Get-CutoverPackFieldValue -Record $mailbox -Names @('CalendarPermissionEntryCount')
            CalendarDelegateState      = Get-CutoverPackTextValue -Record $mailbox -Names @('CalendarDelegateState')
            ArchiveSizeGB              = Get-CutoverPackFieldValue -Record $mailbox -Names @('ArchiveSizeGB')
            ArchiveDeletedItemsGB      = Get-CutoverPackFieldValue -Record $mailbox -Names @('ArchiveDeletedItemsGB')
            TotalDataToMigrateGB       = Get-CutoverPackFieldValue -Record $mailbox -Names @('TotalDataToMigrateGB')
            BitTitanLicenseType        = Get-CutoverPackTextValue -Record $mailbox -Names @('BitTitanLicenseType')
            BitTitanLicenseCount       = Get-CutoverPackFieldValue -Record $mailbox -Names @('BitTitanLicenseCount')
            GrantSendOnBehalfTo        = Get-CutoverPackTextValue -Record $mailbox -Names @('GrantSendOnBehalfTo')
            GrantSendOnBehalfToCount   = Get-CutoverPackFieldValue -Record $mailbox -Names @('GrantSendOnBehalfToCount')
            ForwardingAddress          = Get-CutoverPackTextValue -Record $mailbox -Names @('ForwardingAddress')
            ForwardingSmtpAddress      = Get-CutoverPackTextValue -Record $mailbox -Names @('ForwardingSmtpAddress')
            DeliverToMailboxAndForward = Get-CutoverPackFieldValue -Record $mailbox -Names @('DeliverToMailboxAndForward')
            IsInactiveMailbox          = Get-CutoverPackFieldValue -Record $mailbox -Names @('IsInactiveMailbox')
            ArchiveStatus              = Get-CutoverPackTextValue -Record $mailbox -Names @('ArchiveStatus')
            LitigationHoldEnabled      = Get-CutoverPackFieldValue -Record $mailbox -Names @('LitigationHoldEnabled')
            RetentionPolicy            = Get-CutoverPackTextValue -Record $mailbox -Names @('RetentionPolicy')
            Wave                       = $null
            Notes                      = $null
        }) | Out-Null
    }

    $calendarDelegateRows = New-Object System.Collections.Generic.List[object]
    foreach ($calendarDelegatePermission in @($calendarDelegatePermissions)) {
        $calendarDelegateRows.Add([pscustomobject]@{
            MailboxDisplayName         = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('MailboxDisplayName')
            MailboxPrimarySmtpAddress  = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('MailboxPrimarySmtpAddress')
            MailboxUserPrincipalName   = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('MailboxUserPrincipalName')
            RecipientTypeDetails       = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('RecipientTypeDetails')
            CalendarName               = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('CalendarName')
            CalendarPath               = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('CalendarPath')
            PermissionTarget           = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('PermissionTarget')
            PermissionTargetDisplayName = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('PermissionTargetDisplayName')
            PermissionTargetType       = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('PermissionTargetType')
            AccessRights               = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('AccessRights')
            SharingPermissionFlags     = Get-CutoverPackTextValue -Record $calendarDelegatePermission -Names @('SharingPermissionFlags')
            Wave                       = $null
            Notes                      = $null
        }) | Out-Null
    }

    $recipientReferenceRows = New-Object System.Collections.Generic.List[object]
    foreach ($recipient in @($recipients)) {
        $proxySummary = Get-CutoverPackProxySummary -Record $recipient
        $recipientReferenceRows.Add([pscustomobject]@{
            DisplayName                  = Get-CutoverPackTextValue -Record $recipient -Names @('DisplayName')
            PrimarySmtpAddress           = Get-CutoverPackTextValue -Record $recipient -Names @('PrimarySmtpAddress')
            RecipientTypeDetails         = Get-CutoverPackTextValue -Record $recipient -Names @('RecipientTypeDetails')
            UserPrincipalName            = Get-CutoverPackTextValue -Record $recipient -Names @('UserPrincipalName')
            WindowsEmailAddress          = Get-CutoverPackTextValue -Record $recipient -Names @('WindowsEmailAddress')
            EmailAddresses               = Get-CutoverPackTextValue -Record $recipient -Names @('EmailAddresses')
            OnMicrosoftAlias             = $proxySummary.OnMicrosoftAlias
            OnMicrosoftAliases           = $proxySummary.OnMicrosoftAliases
            HiddenFromAddressListsEnabled = Get-CutoverPackFieldValue -Record $recipient -Names @('HiddenFromAddressListsEnabled')
            Wave                         = $null
            Notes                        = $null
        }) | Out-Null
    }

    $collaborationRows = New-Object System.Collections.Generic.List[object]
    $oneDriveTemplateRows = New-Object System.Collections.Generic.List[object]
    foreach ($userWaveRow in @($userWaveRows.ToArray())) {
        $candidateOneDriveUrl = Resolve-CutoverPackOneDriveUrl -Lookup $oneDriveLookup -CandidateKeys @(
            (ConvertTo-CutoverPackKey $userWaveRow.SourceUPN),
            (ConvertTo-CutoverPackKey $userWaveRow.SourcePrimarySmtpAddress),
            (ConvertTo-OneDriveLookupKey $userWaveRow.SourceUPN),
            (ConvertTo-OneDriveLookupKey $userWaveRow.SourcePrimarySmtpAddress)
        )

        $collaborationRows.Add([pscustomobject]@{
            SourceAccount        = $userWaveRow.SourceAccount
            TargetAccount        = $null
            SourceUPN            = $userWaveRow.SourceUPN
            TargetUPN            = $null
            SourcePrimarySmtpAddress = $userWaveRow.SourcePrimarySmtpAddress
            Direction            = 'MemberToGuest'
            IncludeEntraGroups   = $true
            IncludeTeams         = $true
            IncludeSharePoint    = $true
            ApprovalStatus       = $null
            ApprovalNotes        = $null
            Wave                 = $null
            Notes                = $null
        }) | Out-Null

        $oneDriveTemplateRows.Add([pscustomobject]@{
            SourceUPN         = $userWaveRow.SourceUPN
            SourceEmail       = $userWaveRow.SourcePrimarySmtpAddress
            SourceOneDriveUrl = $candidateOneDriveUrl
            HasOneDrive       = -not [string]::IsNullOrWhiteSpace($candidateOneDriveUrl)
            TargetMail        = $null
            TargetUPN         = $null
            TargetOneDriveUrl = $null
            TargetFolderName  = $null
            Wave              = $null
            Notes             = $null
        }) | Out-Null
    }

    $sheetRowCounts = [ordered]@{
        UserWaveTemplate          = $userWaveRows.Count
        MailboxCutoverReference   = $mailboxReferenceRows.Count
        CalendarDelegateReference = $calendarDelegateRows.Count
        RecipientDomainReference  = $recipientReferenceRows.Count
        CollaborationAccessTemplate = $collaborationRows.Count
        OneDriveMigrationTemplate = $oneDriveTemplateRows.Count
    }

    $packSummaryRows = @(
        [pscustomobject]@{ Section = 'Source'; Metric = 'TenantDisplayName'; Value = $tenantDisplayName; Notes = 'Tenant display name from the assessment snapshot.' },
        [pscustomobject]@{ Section = 'Source'; Metric = 'BaseExportPath'; Value = $resolvedBasePath; Notes = 'Assessment artifact path used to derive the cutover pack.' },
        [pscustomobject]@{ Section = 'Source'; Metric = 'GeneratedAt'; Value = (Get-Date).ToString('o'); Notes = 'Cutover pack generation timestamp.' },
        [pscustomobject]@{ Section = 'Counts'; Metric = 'UserWaveTemplate'; Value = $sheetRowCounts['UserWaveTemplate']; Notes = 'Enabled member users with a populated source UPN.' },
        [pscustomobject]@{ Section = 'Counts'; Metric = 'MailboxCutoverReference'; Value = $sheetRowCounts['MailboxCutoverReference']; Notes = 'Mailbox rows excluding group mailboxes.' },
        [pscustomobject]@{ Section = 'Counts'; Metric = 'CalendarDelegateReference'; Value = $sheetRowCounts['CalendarDelegateReference']; Notes = 'Calendar permission rows that may need to be rebuilt after cutover.' },
        [pscustomobject]@{ Section = 'Counts'; Metric = 'RecipientDomainReference'; Value = $sheetRowCounts['RecipientDomainReference']; Notes = 'Recipient inventory rows collected for domain and proxy planning.' },
        [pscustomobject]@{ Section = 'Counts'; Metric = 'CollaborationAccessTemplate'; Value = $sheetRowCounts['CollaborationAccessTemplate']; Notes = 'Wave-ready source account rows for B2B access planning.' },
        [pscustomobject]@{ Section = 'Counts'; Metric = 'OneDriveMigrationTemplate'; Value = $sheetRowCounts['OneDriveMigrationTemplate']; Notes = 'User rows ready for ShareGate or OneDrive pre-stage mapping.' },
        [pscustomobject]@{ Section = 'Compatibility'; Metric = 'Start-TenantMailCutover'; Value = 'SourcePrimarySmtpAddress, TargetPrimarySmtpAddress, SourceUPN'; Notes = 'Map SourcePrimarySmtpAddress to SourceEmailColumn, TargetPrimarySmtpAddress to TargetEmailColumn, and SourceUPN to SourceUPNColumn.' },
        [pscustomobject]@{ Section = 'Compatibility'; Metric = 'Start-TenantDomainCutover'; Value = 'RecipientDomainReference + MailboxCutoverReference'; Notes = 'Use recipient and mailbox routing/proxy columns when planning accepted-domain cutover batches and validation.' },
        [pscustomobject]@{ Section = 'Compatibility'; Metric = 'CalendarDelegateReference'; Value = 'MailboxPrimarySmtpAddress, CalendarPath, PermissionTarget, AccessRights'; Notes = 'Use the calendar delegate reference to rebuild folder-level delegate permissions after cutover because mailbox calendar rights do not migrate automatically.' },
        [pscustomobject]@{ Section = 'Compatibility'; Metric = 'Re-Apply Migration Access to B2B Accounts.ps1'; Value = 'SourceAccount, TargetAccount, SourceUPN'; Notes = 'CollaborationAccessTemplate is shaped for access-report stamping and later target-account mapping.' },
        [pscustomobject]@{ Section = 'Compatibility'; Metric = 'Start-BulkShareGateOneDriveMigration'; Value = 'SourceUPN, TargetMail, TargetUPN, SourceOneDriveUrl'; Notes = 'OneDriveMigrationTemplate aligns with the common SourceUPN and TargetMail workflow used by the ShareGate migration helpers.' }
    )

    $sheetDefinitions = [ordered]@{
        PackSummary                = [pscustomobject]@{ Rows = @($packSummaryRows); Columns = @('Section', 'Metric', 'Value', 'Notes') }
        UserWaveTemplate           = [pscustomobject]@{ Rows = @($userWaveRows.ToArray()); Columns = $userWaveColumns }
        MailboxCutoverReference    = [pscustomobject]@{ Rows = @($mailboxReferenceRows.ToArray()); Columns = $mailboxReferenceColumns }
        CalendarDelegateReference  = [pscustomobject]@{ Rows = @($calendarDelegateRows.ToArray()); Columns = $calendarDelegateColumns }
        RecipientDomainReference   = [pscustomobject]@{ Rows = @($recipientReferenceRows.ToArray()); Columns = $recipientReferenceColumns }
        CollaborationAccessTemplate = [pscustomobject]@{ Rows = @($collaborationRows.ToArray()); Columns = $collaborationColumns }
        OneDriveMigrationTemplate  = [pscustomobject]@{ Rows = @($oneDriveTemplateRows.ToArray()); Columns = $oneDriveColumns }
    }

    if (Test-Path -Path $workbookPath) {
        Remove-Item -LiteralPath $workbookPath -Force -ErrorAction Stop
    }

    $excelPackage = $null
    $csvArtifacts = [ordered]@{}
    try {
        $excelPackage = Open-ExcelPackage -Path $workbookPath -Create
        foreach ($sheetName in $sheetDefinitions.Keys) {
            $sheetDefinition = $sheetDefinitions[$sheetName]
            $sheetRows = @($sheetDefinition.Rows)
            $sheetColumns = @($sheetDefinition.Columns)
            if ($sheetRows.Count -eq 0) {
                $sheetRows = @(New-CutoverPackPlaceholderRow -Columns $sheetColumns)
            }

            $exportRows = @(
                $sheetRows |
                    Select-Object $sheetColumns |
                    ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ }
            )

            $exportRows |
                Export-Excel -ExcelPackage $excelPackage -WorksheetName $sheetName -ClearSheet -BoldTopRow -PassThru |
                ForEach-Object { $excelPackage = $_ }
            Set-CutoverPackWorksheetColumnWidths -ExcelPackage $excelPackage -WorksheetName $sheetName -Columns $sheetColumns

            $csvFileName = if ([string]::IsNullOrWhiteSpace($artifactPrefix)) {
                '{0}.csv' -f $sheetName
            }
            else {
                '{0}-{1}.csv' -f $artifactPrefix, $sheetName
            }
            $csvPath = Join-Path -Path $cutoverPackDirectory -ChildPath $csvFileName
            $exportRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            $csvArtifacts[$sheetName] = $csvPath
        }
    }
    finally {
        if ($null -ne $excelPackage) {
            try {
                Close-ExcelPackage -ExcelPackage $excelPackage
            }
            catch {
                $excelPackage.Save()
                $excelPackage.Dispose()
            }
        }
    }

    return [pscustomobject]@{
        WorkbookPath  = $workbookPath
        CsvArtifacts  = $csvArtifacts
        SheetRowCounts = $sheetRowCounts
        OutputFolder  = $cutoverPackDirectory
    }
}
