## Export Hash Table to Excel
function Export-HashTableToExcel {
    [CmdletBinding()]
    param ( 
        [Parameter(Mandatory=$True)] 
        [Hashtable]$hashtable,
        [Parameter(Mandatory=$True)] 
        [string]$ExportDetails,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Default', 'TenantToTenantCutover')]
        [string]$WorkbookExportPolicy = 'Default'
        #[Parameter(Mandatory=$false)] [switch]$tenant
    )

    # Ensure ExportDetails has .xlsx in the path name
    if ($ExportDetails -notmatch '\.xlsx$') {
        $ExportDetails += ".xlsx"
    }

    function Test-FileLocked {
        param([string]$Path)
        if (-not (Test-Path $Path)) { return $false }
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $stream.Close()
            return $false
        } catch {
            return $true
        }
    }

    # Central worksheet alias map to avoid Excel auto-truncation warnings.
    $worksheetAliases = @{
        'Office365GroupsActivityTopGroups' = 'O365GroupsActivityTop'
        'EmployeeExperienceInsightsSummary' = 'EmployeeExpInsights'
        'MailboxCalendarDelegatePermissions' = 'MailboxCalendarDelegatePerms'
        'PrivilegedAccessRemediationSummary' = 'PrivilegedAccessRemediation'
    }

    $worksheetCleanupAliases = @{
        'Office365GroupsActivityTopGroups' = @('Office365GroupsActivityTopGroups', 'Office365GroupsActivityTopGroup', 'O365GroupsActivityTop')
        'EmployeeExperienceInsightsSummary' = @('EmployeeExperienceInsightsSummary', 'EmployeeExperienceInsightsSumma', 'EmployeeExpInsights')
        'MailboxCalendarDelegatePermissions' = @('MailboxCalendarDelegatePermissions', 'MailboxCalendarDelegatePermissi', 'MailboxCalendarDelegatePerms')
        'PrivilegedAccessRemediationSummary' = @('PrivilegedAccessRemediationSummary', 'PrivilegedAccessRemediationSumm', 'PrivilegedAccessRemediation')
    }

    function Resolve-WorksheetExportName {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$LogicalName
        )

        if ($worksheetAliases.ContainsKey($LogicalName)) {
            return [string]$worksheetAliases[$LogicalName]
        }
        return $LogicalName
    }

    function Get-WorksheetCleanupCandidates {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$LogicalName
        )

        if ($worksheetCleanupAliases.ContainsKey($LogicalName)) {
            return @($worksheetCleanupAliases[$LogicalName] | Select-Object -Unique)
        }

        return @($LogicalName)
    }

    $optionalEmptySheets = @(
        'AuthenticationSSOApplications',
        'EnterpriseApplications',
        'SMTPRelaySummary',
        'TeamsVoiceSummary',
        'SharePointActivityUserDetail',
        'OneDriveActivityUserDetail',
        'UnmanagedObjects',
        'OneDriveOwnerMismatches',
        'TeamsActivityTopUsers',
        'Office365GroupsActivityTopGroups',
        'EmployeeExperienceInsightsSummary',
        'InboxRulesExternalForwarding',
        'ExternalSharingSiteOverrides',
        'MailboxCalendarDelegatePermissions',
        'MfaEnforcementGapUsers',
        'MfaEnforcementScopeReview',
        'ConditionalAccessOptimization',
        'MfaMethodPostureSummary',
        'PrivilegedAccessRemediationSummary',
        'TeamsGroupsCleanupCandidates',
        'GroupLicensingSummary',
        'LicenseOptimizationCandidates'
    )

    $defaultExcludedWorksheets = @(
        'OwnershipGovernanceSummary',
        'TenantInfoSummary',
        'AuthenticationConfigSummary',
        'SpamFilteringSummary',
        'FederationSummary',
        'MfaRegistrationSummary',
        'EmailActivitySummary',
        'PrimaryMailboxStatsCollectionSummary',
        'UnifiedGroupMailboxStatsCollectionSummary',
        'PrimaryMailboxStatsCollectionSu',
        'UnifiedGroupMailboxStatsCollect'
    )

    function Wait-ForExportFileAvailability {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,
            [Parameter(Mandatory = $false)]
            [int]$MaxAttempts = 3
        )

        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            if (-not (Test-FileLocked -Path $Path)) {
                return
            }

            if ($attempt -lt $MaxAttempts) {
                Write-Log -Type WARNING -Message "Excel file is locked (attempt $attempt/$MaxAttempts). Close the file and retrying in 5 seconds..." -ExportFileLocation $ExportDetails
                Start-Sleep -Seconds 5
                continue
            }

            throw "Excel file '$Path' is locked and could not be opened for export."
        }
    }

    function Get-WorksheetExportSourceInfo {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            $TableValue
        )

        $sourceEnumerable = $null
        $sourceCount = 0
        $singleValue = $null

        if ($TableValue -is [hashtable] -or $TableValue -is [System.Collections.Specialized.OrderedDictionary]) {
            $sourceCount = $TableValue.Count
            $sourceEnumerable = $TableValue.Values
        }
        elseif ($TableValue -is [System.Collections.IEnumerable] -and -not ($TableValue -is [string])) {
            if ($TableValue.PSObject.Properties['Count']) {
                try { $sourceCount = [int]$TableValue.Count } catch { $sourceCount = -1 }
            }
            else {
                $sourceCount = -1
            }
            $sourceEnumerable = $TableValue
        }
        else {
            if ($null -ne $TableValue) {
                $singleValue = $TableValue
                $sourceCount = 1
            }
        }

        if ($sourceCount -lt 0 -and $null -ne $sourceEnumerable) {
            $sourceEnumerable = @($sourceEnumerable)
            $sourceCount = $sourceEnumerable.Count
        }

        return [pscustomobject]@{
            Count  = $sourceCount
            Source = if ($null -ne $singleValue) { @($singleValue) } else { $sourceEnumerable }
        }
    }

    function Convert-WorkbookExportTableRows {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            $TableValue
        )

        if ($null -eq $TableValue) {
            return @()
        }

        if ($TableValue -is [hashtable] -or $TableValue -is [System.Collections.Specialized.OrderedDictionary]) {
            return @($TableValue.Values)
        }

        if (($TableValue -is [System.Collections.IEnumerable]) -and -not ($TableValue -is [string])) {
            return @($TableValue)
        }

        return @($TableValue)
    }

    function Get-WorkbookExportObjectValue {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $InputObject,
            [Parameter(Mandatory = $true)]
            [string[]]$Names
        )

        if ($null -eq $InputObject) {
            return $null
        }

        if ($InputObject -is [System.Collections.IDictionary]) {
            foreach ($name in $Names) {
                foreach ($key in @($InputObject.Keys)) {
                    if ([string]::Equals([string]$key, $name, [System.StringComparison]::OrdinalIgnoreCase)) {
                        return $InputObject[$key]
                    }
                }
            }
        }

        foreach ($name in $Names) {
            $property = $InputObject.PSObject.Properties[$name]
            if ($null -ne $property) {
                return $property.Value
            }
        }

        return $null
    }

    function Convert-WorkbookExportNumber {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Value
        )

        if ($null -eq $Value) {
            return $null
        }

        if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
            return [double]$Value
        }

        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        $number = 0.0
        if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
            return $number
        }

        return $null
    }

    function Test-WorkbookExportMeaningfulLicenseText {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Value
        )

        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $false
        }

        return ($text -notmatch '^(?i)notcollected|not surfaced|unknown|unavailable$')
    }

    function Test-WorkbookExportLicensingGroupRow {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Record
        )

        if ($null -eq $Record) {
            return $false
        }

        $assignedLicenseCount = Convert-WorkbookExportNumber -Value (Get-WorkbookExportObjectValue -InputObject $Record -Names @('AssignedLicenseCount'))
        if ($null -ne $assignedLicenseCount) {
            return ($assignedLicenseCount -gt 0)
        }

        foreach ($fieldName in @('AssignedLicenseSkuIds', 'AssignedLicenseSkuPartNumbers', 'AssignedLicenseFriendlyNames', 'SkuNames')) {
            if (Test-WorkbookExportMeaningfulLicenseText -Value (Get-WorkbookExportObjectValue -InputObject $Record -Names @($fieldName))) {
                return $true
            }
        }

        $isManagingLicenses = Get-WorkbookExportObjectValue -InputObject $Record -Names @('IsManagingLicenses')
        if ($isManagingLicenses -is [bool]) {
            return [bool]$isManagingLicenses
        }

        return (([string]$isManagingLicenses).Trim() -match '^(?i)true$')
    }

    # === Container worksheet shaping ===
    # Several collectors store a worksheet as a "container": a Summary record plus sibling
    # arrays (TeamsVoice, AdConnectConfiguration), or a single Configuration record whose
    # properties are themselves arrays (AuthenticationConfig). Emitting those through
    # $hashtable.Values collapses each array into one cell, or worse, reflects the array's
    # own .NET members as worksheet columns. These helpers flatten them into
    # Section / Item / Value / Notes rows instead, one row per underlying record.

    $containerWorksheetSections = @{
        'AdConnectConfiguration' = 'Directory Sync'
        'TeamsVoice'             = 'Voice Summary'
        'AuthenticationConfig'   = 'Authentication'
    }

    function Convert-WorkbookExportLabel {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$Text
        )

        if ([string]::IsNullOrWhiteSpace($Text)) {
            return ''
        }

        # PascalCase to spaced words, preserving acronyms: SyncServices -> Sync Services,
        # SSOApplications -> SSO Applications, MFAMethods -> MFA Methods.
        $spaced = [regex]::Replace($Text, '(?<=[a-z0-9])([A-Z])', ' $1')
        $spaced = [regex]::Replace($spaced, '(?<=[A-Z])([A-Z][a-z])', ' $1')
        return $spaced.Trim()
    }

    function Convert-WorkbookExportDisplayText {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Value,
            [Parameter(Mandatory = $false)]
            [string]$Default = ''
        )

        $text = [string](ConvertTo-ExportFriendlyValue -Value $Value)
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $Default
        }

        return $text.Trim()
    }

    function Test-WorkbookExportCollection {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Value
        )

        return (
            ($Value -is [System.Collections.IEnumerable]) -and
            -not ($Value -is [string]) -and
            -not ($Value -is [System.Collections.IDictionary])
        )
    }

    function New-WorkbookExportConfigurationRow {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Section,
            [Parameter(Mandatory = $true)]
            [string]$Item,
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Value,
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Notes
        )

        return [pscustomobject][ordered]@{
            Section = $Section
            Item    = $Item
            Value   = ConvertTo-ExportFriendlyValue -Value $Value
            Notes   = ConvertTo-ExportFriendlyValue -Value $Notes
        }
    }

    function Get-WorkbookExportRecordFields {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Record
        )

        $fields = [ordered]@{}
        if ($null -eq $Record) {
            return $fields
        }

        if ($Record -is [System.Collections.IDictionary]) {
            foreach ($key in @($Record.Keys)) {
                $fields[[string]$key] = $Record[$key]
            }
            return $fields
        }

        foreach ($property in @($Record.PSObject.Properties | Where-Object { $_.MemberType -in @('NoteProperty', 'AliasProperty') })) {
            $fields[$property.Name] = $property.Value
        }

        return $fields
    }

    function Add-WorkbookExportCollectionRows {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            $Rows,
            [Parameter(Mandatory = $true)]
            [string]$Section,
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Collection
        )

        $added = 0
        foreach ($item in @($Collection)) {
            if ($null -eq $item) {
                continue
            }

            $added++

            if (($item -is [string]) -or ($item -is [ValueType])) {
                $Rows.Add((New-WorkbookExportConfigurationRow -Section $Section -Item (Convert-WorkbookExportDisplayText -Value $item -Default "Entry $added") -Value 'Configured')) | Out-Null
                continue
            }

            $fields = Get-WorkbookExportRecordFields -Record $item
            $label = ''
            foreach ($labelName in @('DisplayName', 'Name', 'Title', 'ServiceName', 'PolicyName', 'ErrorType', 'Category', 'Identity', 'UserPrincipalName', 'Mail', 'TelephoneNumber', 'PhoneNumber', 'Domain', 'AppId', 'Id')) {
                if ($fields.Contains($labelName)) {
                    $label = Convert-WorkbookExportDisplayText -Value $fields[$labelName]
                    if (-not [string]::IsNullOrWhiteSpace($label)) { break }
                }
            }
            if ([string]::IsNullOrWhiteSpace($label)) {
                $label = "Entry $added"
            }

            # Remaining fields become the Notes column so no collected detail is dropped.
            $detail = @()
            foreach ($fieldName in @($fields.Keys)) {
                $fieldText = Convert-WorkbookExportDisplayText -Value $fields[$fieldName]
                if ([string]::IsNullOrWhiteSpace($fieldText) -or $fieldText -eq $label) {
                    continue
                }
                $detail += ('{0}: {1}' -f (Convert-WorkbookExportLabel -Text $fieldName), $fieldText)
            }

            $Rows.Add((New-WorkbookExportConfigurationRow -Section $Section -Item $label -Value 'Detected' -Notes ($detail -join '; '))) | Out-Null
        }

        if ($added -eq 0) {
            $Rows.Add((New-WorkbookExportConfigurationRow -Section $Section -Item 'None detected' -Value 0 -Notes "No $Section records were collected in this run.")) | Out-Null
        }
    }

    function Add-WorkbookExportRecordRows {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            $Rows,
            [Parameter(Mandatory = $true)]
            [string]$Section,
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $Record
        )

        $fields = Get-WorkbookExportRecordFields -Record $Record
        foreach ($fieldName in @($fields.Keys)) {
            $value = $fields[$fieldName]

            # An array-valued property is the issue-2 case: expand it per row instead of
            # letting ConvertTo-ExportFriendlyValue join it into a single cell.
            if (Test-WorkbookExportCollection -Value $value) {
                Add-WorkbookExportCollectionRows -Rows $Rows -Section (Convert-WorkbookExportLabel -Text $fieldName) -Collection $value
                continue
            }

            $Rows.Add((New-WorkbookExportConfigurationRow -Section $Section -Item (Convert-WorkbookExportLabel -Text $fieldName) -Value $value)) | Out-Null
        }
    }

    function ConvertTo-WorkbookExportContainerRows {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TableName,
            [Parameter(Mandatory = $true)]
            $TableValue
        )

        $defaultSection = if ($containerWorksheetSections.ContainsKey($TableName)) {
            [string]$containerWorksheetSections[$TableName]
        }
        else {
            'Summary'
        }

        $rows = New-Object System.Collections.Generic.List[object]

        # Hashtable key order is not stable across runs. Emit the Summary/Configuration
        # record first, then the remaining keys alphabetically, so the worksheet is
        # reproducible and reads top-down.
        $summaryKeys = @(@($TableValue.Keys) | Where-Object { [string]$_ -in @('Summary', 'Configuration') })
        $otherKeys = @(@($TableValue.Keys) | Where-Object { [string]$_ -notin @('Summary', 'Configuration') } | Sort-Object { [string]$_ })

        foreach ($key in @($summaryKeys + $otherKeys)) {
            $value = $TableValue[$key]
            if ($null -eq $value) {
                continue
            }

            $keyName = [string]$key

            if (Test-WorkbookExportCollection -Value $value) {
                Add-WorkbookExportCollectionRows -Rows $rows -Section (Convert-WorkbookExportLabel -Text $keyName) -Collection $value
                continue
            }

            # Summary / Configuration hold the record whose properties are the real settings.
            if ($keyName -in @('Summary', 'Configuration')) {
                Add-WorkbookExportRecordRows -Rows $rows -Section $defaultSection -Record $value
                continue
            }

            if (($value -is [System.Collections.IDictionary]) -or
                (@($value.PSObject.Properties | Where-Object { $_.MemberType -in @('NoteProperty', 'AliasProperty') }).Count -gt 0)) {
                Add-WorkbookExportRecordRows -Rows $rows -Section (Convert-WorkbookExportLabel -Text $keyName) -Record $value
                continue
            }

            $rows.Add((New-WorkbookExportConfigurationRow -Section 'General' -Item (Convert-WorkbookExportLabel -Text $keyName) -Value $value)) | Out-Null
        }

        return @($rows.ToArray())
    }

    function Test-WorkbookExportMixedContainer {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            $TableValue
        )

        if (-not ($TableValue -is [System.Collections.IDictionary]) -or $TableValue.Count -eq 0) {
            return $false
        }

        $hasCollection = $false
        $hasRecord = $false
        foreach ($value in $TableValue.Values) {
            if ($null -eq $value) { continue }
            if (Test-WorkbookExportCollection -Value $value) { $hasCollection = $true } else { $hasRecord = $true }
        }

        # Uniform lookup tables (Users, SharePoint) and single-Summary containers
        # (CollaborationActivitySummary) keep their existing one-record-per-value behaviour.
        return ($hasCollection -and $hasRecord)
    }

    function ConvertTo-WorkbookExportMfaMethodRows {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            $TableValue,
            [Parameter(Mandatory = $true)]
            [hashtable]$ExportTables
        )

        $summary = $TableValue
        if ($summary -is [System.Collections.IDictionary]) {
            $summary = Get-WorkbookExportObjectValue -InputObject $summary -Names @('Summary', 'Configuration')
        }
        if ($null -eq $summary) {
            return $null
        }

        # The pre-joined *Breakdown strings on MfaEnrollmentSummary are lossy. The raw
        # per-method count maps survive on MfaRegistrationSummary, which is excluded from
        # the workbook, so read the counts back from there and emit one row per method.
        $registration = $null
        if ($ExportTables.ContainsKey('MfaRegistrationSummary')) {
            $registration = $ExportTables['MfaRegistrationSummary']
            if ($registration -is [System.Collections.IDictionary]) {
                $registration = Get-WorkbookExportObjectValue -InputObject $registration -Names @('Summary', 'Configuration')
            }
        }
        if ($null -eq $registration) {
            return $null
        }

        $totalUsers = Convert-WorkbookExportNumber -Value (Get-WorkbookExportObjectValue -InputObject $summary -Names @('TotalUsers'))
        $rows = New-Object System.Collections.Generic.List[object]

        foreach ($totalName in @('TotalUsers', 'RegisteredUsers', 'NotRegisteredUsers', 'RegistrationPercent', 'UsersWithWeakMethodsOnly', 'UsersWithWeakDefaultMethod', 'UsersWithStrongMethods', 'UsersWithPhishingResistantMethods')) {
            $totalValue = Get-WorkbookExportObjectValue -InputObject $summary -Names @($totalName)
            if ($null -eq $totalValue) { continue }
            $rows.Add([pscustomobject][ordered]@{
                Category       = 'Totals'
                Method         = Convert-WorkbookExportLabel -Text $totalName
                UserCount      = $totalValue
                PercentOfUsers = $null
                Notes          = 'Tenant-wide MFA enrollment total.'
            }) | Out-Null
        }

        $categoryMaps = @(
            @{ Category = 'Registered';         Names = @('MethodCounts') ;                  Notes = 'Users with this authentication method registered.' }
            @{ Category = 'Weak';               Names = @('WeakMethodCounts') ;              Notes = 'SMS, voice, and email OTP methods are phishable.' }
            @{ Category = 'Strong';             Names = @('StrongMethodCounts') ;            Notes = 'Authenticator app and comparable strong methods.' }
            @{ Category = 'PhishingResistant';  Names = @('PhishingResistantMethodCounts') ; Notes = 'FIDO2, passkeys, and certificate-based methods.' }
            @{ Category = 'Default';            Names = @('DefaultMethodCounts') ;           Notes = 'Method currently set as the user default.' }
        )

        foreach ($categoryMap in $categoryMaps) {
            $countMap = Get-WorkbookExportObjectValue -InputObject $registration -Names $categoryMap.Names
            if ($null -eq $countMap) { continue }

            $countFields = Get-WorkbookExportRecordFields -Record $countMap
            foreach ($methodName in @($countFields.Keys)) {
                $count = Convert-WorkbookExportNumber -Value $countFields[$methodName]
                $percent = if ($null -ne $count -and $null -ne $totalUsers -and $totalUsers -gt 0) {
                    [math]::Round(($count / $totalUsers) * 100, 1)
                }
                else {
                    $null
                }

                $rows.Add([pscustomobject][ordered]@{
                    Category       = [string]$categoryMap.Category
                    # Method names are Graph enum values (microsoftAuthenticator, fido2).
                    # Keep them verbatim so they match what operators see in Entra.
                    Method         = [string]$methodName
                    UserCount      = $countFields[$methodName]
                    PercentOfUsers = $percent
                    Notes          = [string]$categoryMap.Notes
                }) | Out-Null
            }
        }

        if ($rows.Count -eq 0) {
            return $null
        }

        return @($rows.ToArray())
    }

    function Select-WorkbookExportRows {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TableName,
            [Parameter(Mandatory = $false)]
            $TableValue,
            [Parameter(Mandatory = $true)]
            [hashtable]$ExportTables
        )

        if ($TableName -eq 'MfaEnrollmentSummary') {
            $methodRows = ConvertTo-WorkbookExportMfaMethodRows -TableValue $TableValue -ExportTables $ExportTables
            if ($null -ne $methodRows -and @($methodRows).Count -gt 0) {
                return $methodRows
            }

            return $TableValue
        }

        # Only dictionaries are reshaped. Under the TenantToTenantCutover policy some of
        # these names already resolve to pre-flattened row arrays, which must pass through.
        if ($TableValue -is [System.Collections.IDictionary] -and $TableValue.Count -gt 0) {
            if ($containerWorksheetSections.ContainsKey($TableName) -or (Test-WorkbookExportMixedContainer -TableValue $TableValue)) {
                return (ConvertTo-WorkbookExportContainerRows -TableName $TableName -TableValue $TableValue)
            }
        }

        if ($TableName -eq 'GroupLicensingSummary') {
            return @(Convert-WorkbookExportTableRows -TableValue $TableValue | Where-Object { Test-WorkbookExportLicensingGroupRow -Record $_ })
        }

        if ($TableName -eq 'LicenseOptimizationCandidates') {
            $validLicensingGroupNames = @{}
            if ($ExportTables.ContainsKey('GroupLicensingSummary')) {
                foreach ($groupRow in @(Convert-WorkbookExportTableRows -TableValue $ExportTables['GroupLicensingSummary'] | Where-Object { Test-WorkbookExportLicensingGroupRow -Record $_ })) {
                    $groupName = [string](Get-WorkbookExportObjectValue -InputObject $groupRow -Names @('GroupName', 'DisplayName', 'Name'))
                    if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                        $validLicensingGroupNames[$groupName.ToLowerInvariant()] = $true
                    }
                }
            }

            return @(
                Convert-WorkbookExportTableRows -TableValue $TableValue | Where-Object {
                    $objectType = [string](Get-WorkbookExportObjectValue -InputObject $_ -Names @('ObjectType'))
                    if ($objectType -notmatch '^(?i)group$') {
                        return $true
                    }

                    if (Test-WorkbookExportMeaningfulLicenseText -Value (Get-WorkbookExportObjectValue -InputObject $_ -Names @('SkuNames', 'AssignedLicenseFriendlyNames', 'AssignedLicenseSkuPartNumbers'))) {
                        return $true
                    }

                    $displayName = [string](Get-WorkbookExportObjectValue -InputObject $_ -Names @('DisplayName', 'GroupName', 'Name'))
                    return (-not [string]::IsNullOrWhiteSpace($displayName) -and $validLicensingGroupNames.ContainsKey($displayName.ToLowerInvariant()))
                }
            )
        }

        return $TableValue
    }

    function New-ExplicitWorksheetPlaceholderRow {
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

    function Get-TenantToTenantWorksheetColumns {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$LogicalName
        )

        $columnMap = @{
            'TenantInfo'                      = @('DisplayName', 'TenantId', 'InitialDomain', 'DefaultDomain', 'Country', 'CountryLetterCode', 'PreferredDataLocation', 'MultiGeoEnabled', 'MultiGeoAllowed', 'MultiGeoCentral', 'SelfServicePurchase', 'SelfServicePurchaseNotes', 'AzureResourceUsage', 'AzureResourceUsageNotes')
            'MigrationExecutiveSummary'       = @('Section', 'Metric', 'Value', 'Notes')
            'RecipientDomainSummary'          = @('Domain', 'RecipientCount', 'UserMailboxCount', 'SharedMailboxCount', 'GroupRecipientCount', 'HiddenFromAddressListsCount', 'Notes')
            'MailboxMigrationSummary'         = @('RecipientTypeDetails', 'MailboxCount', 'ActiveMailboxCount', 'InactiveMailboxCount', 'ArchiveEnabledCount', 'ForwardingCount', 'MailboxesWithDelegateDependencies', 'TotalDataToMigrateGB')
            'BitTitanLicenseSummary'          = @('Section', 'Metric', 'Value', 'Notes')
            'DelegateSummary'                 = @('PermissionType', 'AffectedMailboxCount', 'AssignmentCount', 'CollectionStates', 'Notes')
            'CollaborationSummary'            = @('Workload', 'TotalCount', 'TotalStorageGB', 'UnknownStorageCount', 'LargestObjectName', 'LargestObjectSizeGB', 'Notes')
            'CutoverPrepSummary'              = @('Category', 'Item', 'Status', 'Value', 'Notes')
            'AdConnectConfiguration'          = @('Section', 'Item', 'Value', 'Notes')
            'HybridConfiguration'             = @('Section', 'Item', 'Value', 'Notes')
            'Domains'                         = @('Name', 'IsDefault', 'IsInitial', 'Verified', 'AuthenticationType', 'SupportedServices', 'MXRecords', 'Office365MailExchanger', 'ThirdPartySpamFilterReview', 'HybridRoutingReview', 'RecipientCount', 'Notes')
            'Users'                           = @('DisplayName', 'UserPrincipalName', 'Mail', 'AccountEnabled', 'OnPremisesSyncEnabled', 'AssignedLicensesFriendly', 'HasMailbox', 'MailboxPrimarySmtpAddress', 'RecipientTypeDetails', 'TargetUPN', 'TargetPrimarySmtpAddress', 'Wave', 'MigrationState', 'CutoverDate', 'Notes')
            'AllRecipients'                   = @('DisplayName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'UserPrincipalName', 'WindowsEmailAddress', 'PrimaryDomain', 'EmailAddresses', 'OnMicrosoftAlias', 'OnMicrosoftAliases', 'HiddenFromAddressListsEnabled', 'Notes')
            'AllMailboxes'                    = @('DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'IsInactiveMailbox', 'OnPremisesSyncEnabled', 'ExchangeGuid', 'ArchiveGuid', 'OnMicrosoftAlias', 'OnMicrosoftAliases', 'LegacyExchangeDn', 'LegacyExchangeDnX500', 'X500Addresses', 'X400Addresses', 'MailboxSizeGB', 'DeletedItemsGB', 'ArchiveStatus', 'ArchiveSizeGB', 'ArchiveDeletedItemsGB', 'TotalDataToMigrateGB', 'BitTitanLicenseType', 'BitTitanLicenseCount', 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward', 'LitigationHoldEnabled', 'RetentionPolicy', 'FullAccessDelegateCount', 'SendAsDelegateCount', 'GrantSendOnBehalfToCount', 'CalendarDelegateCount', 'Wave', 'MigrationState', 'CutoverDate', 'Notes')
            'MailboxDelegateAssignments'      = @('MailboxDisplayName', 'MailboxPrimarySmtpAddress', 'MailboxUserPrincipalName', 'RecipientTypeDetails', 'PermissionType', 'Delegate', 'DelegateCountSource', 'CollectionState', 'Wave', 'Notes')
            'MailboxCalendarDelegatePermissions' = @('MailboxDisplayName', 'MailboxPrimarySmtpAddress', 'MailboxUserPrincipalName', 'RecipientTypeDetails', 'CalendarName', 'CalendarPath', 'PermissionTarget', 'PermissionTargetDisplayName', 'PermissionTargetType', 'AccessRights', 'SharingPermissionFlags', 'Wave', 'Notes')
            'InactiveMailboxDetails'          = @('DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'IsInactiveMailbox', 'OnPremisesSyncEnabled', 'ExchangeGuid', 'ArchiveGuid', 'OnMicrosoftAlias', 'OnMicrosoftAliases', 'LegacyExchangeDn', 'LegacyExchangeDnX500', 'X500Addresses', 'X400Addresses', 'MailboxSizeGB', 'DeletedItemsGB', 'ArchiveStatus', 'ArchiveSizeGB', 'ArchiveDeletedItemsGB', 'TotalDataToMigrateGB', 'BitTitanLicenseType', 'BitTitanLicenseCount', 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward', 'LitigationHoldEnabled', 'RetentionPolicy', 'FullAccessDelegateCount', 'SendAsDelegateCount', 'GrantSendOnBehalfToCount', 'CalendarDelegateCount', 'Wave', 'MigrationState', 'CutoverDate', 'Notes')
            'PublicFolderDetails'             = @('Name', 'Identity', 'Path', 'MailEnabled', 'PrimarySmtpAddress', 'ItemCount', 'FolderCount', 'EntryId', 'Notes')
            'MailFlowConnectors'              = @('Name', 'Enabled', 'ConnectorType', 'ConnectorSource', 'SenderDomains', 'RecipientDomains', 'SmartHosts', 'TlsSettings', 'Comment', 'Notes')
            'RemoteDomains'                   = @('Name', 'DomainName', 'AutoForwardEnabled', 'AllowedOOFType', 'TNEFEnabled', 'TrustedMailOutboundEnabled', 'Notes')
            'SMTPRelayServiceAccounts'        = @('DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'IsDirSynced', 'AssignedLicensesFriendly', 'Notes')
            'AllTeams'                        = @('DisplayName', 'Visibility', 'IsArchived', 'SharePointSiteUrl', 'SiteSize-GB', 'TotalChannels', 'SharedChannelCount', 'SharedChannels', 'OwnerCount', 'MemberCount', 'GuestCount', 'LastActivityDate', 'Notes')
            'SharePoint'                      = @('Title', 'Url', 'Template', 'Owner', 'StorageUsedGB', 'StorageQuota', 'LastContentModifiedDate', 'LockState', 'ArchiveStatus', 'SharingCapability', 'IsTeamsConnected', 'Notes')
            'OneDrive'                        = @('Title', 'Url', 'Template', 'Owner', 'StorageUsedGB', 'StorageQuota', 'LastContentModifiedDate', 'LockState', 'ArchiveStatus', 'SharingCapability', 'IsTeamsConnected', 'Notes')
        }

        if ($columnMap.ContainsKey($LogicalName)) {
            return @($columnMap[$LogicalName])
        }

        return @()
    }

    function Get-TenantToTenantWorksheetSourceName {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$LogicalName,
            [Parameter(Mandatory = $true)]
            [hashtable]$ExportTables
        )

        switch ($LogicalName) {
            'TenantInfo' {
                if (
                    $ExportTables.ContainsKey('TenantInfoSummary') -and
                    (Get-WorksheetExportSourceInfo -TableValue $ExportTables['TenantInfoSummary']).Count -gt 0
                ) {
                    return 'TenantInfoSummary'
                }
            }
            'AdConnectConfiguration' {
                return 'TenantToTenantAdConnectConfiguration'
            }
            'HybridConfiguration' {
                return 'TenantToTenantHybridConfiguration'
            }
            'AllMailboxes' {
                if (
                    $ExportTables.ContainsKey('MailboxFullDetails') -and
                    (Get-WorksheetExportSourceInfo -TableValue $ExportTables['MailboxFullDetails']).Count -gt 0
                ) {
                    return 'MailboxFullDetails'
                }
            }
        }

        return $LogicalName
    }

    function Get-TenantToTenantColumnWidth {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$ColumnName
        )

        if ($ColumnName -match '^(MailboxSizeGB|DeletedItemsGB|ArchiveSizeGB|ArchiveDeletedItemsGB|TotalDataToMigrateGB|BitTitanLicenseCount|RecipientCount|UserMailboxCount|SharedMailboxCount|GroupRecipientCount|HiddenFromAddressListsCount|MailboxCount|ActiveMailboxCount|InactiveMailboxCount|ArchiveEnabledCount|ForwardingCount|MailboxesWithDelegateDependencies|AffectedMailboxCount|AssignmentCount|TotalCount|TotalStorageGB|UnknownStorageCount|LargestObjectSizeGB|OwnerCount|MemberCount|GuestCount|TotalChannels|SharedChannelCount|Count|Status|AccountEnabled|HasMailbox|IsInactiveMailbox|OnPremisesSyncEnabled|IsArchived|Verified|IsDefault|IsInitial|Office365MailExchanger|LitigationHoldEnabled|DeliverToMailboxAndForward)$') {
            return 14
        }

        if ($ColumnName -match '^(DisplayName|UserPrincipalName|PrimarySmtpAddress|WindowsEmailAddress|MailboxPrimarySmtpAddress|MailboxUserPrincipalName|Mail|Owner|Title|Url|SharePointSiteUrl|Source|Target|TargetUPN|TargetPrimarySmtpAddress|TargetMail|SourceUPN|SourceEmail|SourceOneDriveUrl|TargetOneDriveUrl|LargestObjectName|ForwardingAddress|ForwardingSmtpAddress)$') {
            return 26
        }

        if ($ColumnName -match '^(EmailAddresses|OnMicrosoftAliases|LegacyExchangeDn|LegacyExchangeDnX500|X500Addresses|X400Addresses|Delegate|FullAccessDelegates|SendAsDelegates|GrantSendOnBehalfTo|SmartHosts|SenderDomains|RecipientDomains|Comment|Notes|CollectionStates|SupportedServices|TlsSettings|MXRecords|SharedChannels)$') {
            return 34
        }

        return 20
    }

    function Set-TenantToTenantWorksheetColumnWidths {
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
            $worksheet.Column($columnIndex).Width = Get-TenantToTenantColumnWidth -ColumnName $Columns[$columnIndex - 1]
            $worksheet.Column($columnIndex).BestFit = $false
        }
    }
    function ConvertTo-WorkbookExportUniformRecords {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [object[]]$Records
        )

        if ($null -eq $Records -or $Records.Count -le 1) {
            return $Records
        }

        # Export-Excel derives its columns from the first record only, so any property a
        # later record adds is silently dropped. Build the union and pad only when the
        # records are actually ragged, to avoid copying large uniform sheets.
        # Single pass over each record's properties: large worksheets run this for every
        # row, so the per-record property count is tallied inline rather than by
        # re-enumerating PSObject.Properties.
        $union = [ordered]@{}
        $ragged = $false
        foreach ($record in $Records) {
            if ($null -eq $record) { continue }
            $propertyCount = 0
            foreach ($property in $record.PSObject.Properties) {
                $propertyCount++
                if (-not $union.Contains($property.Name)) {
                    if ($union.Count -gt 0) { $ragged = $true }
                    $union[$property.Name] = $true
                }
            }
            if ($propertyCount -ne $union.Count) {
                $ragged = $true
            }
        }

        if (-not $ragged) {
            return $Records
        }

        $columns = @($union.Keys)
        return @(
            foreach ($record in $Records) {
                $normalized = [ordered]@{}
                foreach ($column in $columns) {
                    $property = if ($null -ne $record) { $record.PSObject.Properties[$column] } else { $null }
                    $normalized[$column] = if ($property) { $property.Value } else { $null }
                }
                [pscustomobject]$normalized
            }
        )
    }

    # === Sheet ordering ===
    $defaultDesiredOrder = @(
        # Tenant Overview
        "TenantInfo", "LicenseSKUs", "GroupLicensingSummary", "LicenseOptimizationCandidates", "AdConnectConfiguration",

        # Identity
        "Users", "UserFullDetails", "Admins", "EntraIDGroups", "Domains",
        "AuthenticationMethods", "AuthenticationSSOApplications", "EnterpriseApplications", "AuthenticationConfig",
        "MfaEnrollmentSummary", "MfaMethodPostureSummary", "MfaEnforcementSummary", "MfaEnforcementGapUsers", "MfaEnforcementScopeReview",
        "ConditionalAccessPolicySummary", "ConditionalAccessOptimization", "ConditionalAccessPolicies", "SecurityDefaultsPolicy", "GuestSignInSummary",
        "GuestAccessConfiguration", "ExternalIdentityRestrictions", "PrivilegedAccessSummary", "PrivilegedAccessRemediationSummary",

        # Exchange
        "HybridConfiguration",
        "AllRecipients", "AllMailboxes", "PrimaryMailboxStats", "MailboxFullDetails", "MailboxCalendarDelegatePermissions", "NonUserMailboxes",
        "ArchiveMailboxes", "ArchiveMailboxStats", "LitigationHoldMailboxes", "InactiveMailboxes", "InactiveMailboxDetails",
        "AllExchangeGroups", "PublicFolderDetails", "PublicFolderPerms",
        "MailFlowRules", "MailFlowConnectors", "RemoteDomains", "EmailActivityTopSenders", "EmailActivityTopReceivers",
        "SMTPRelayConfig", "SMTPRelayServiceAccounts", "SpamFilteringConfig", "SharedMailboxGovernanceSummary",
        "ForwardingPolicySummary", "InboxRuleForwardingSummary", "InboxRulesExternalForwarding",

        # Collaboration
        "UnifiedGroups", "AllTeams", "TeamsGroupsCleanupCandidates", "TeamsVoiceSummary", "TeamsVoice", "SharePoint", "OneDrive",
        "SharePointSharingSummary", "SharePointActivityUserDetail", "OneDriveActivityUserDetail",
        "TeamsActivityTopUsers", "Office365GroupsActivityTopGroups",
        "EmployeeExperienceInsightsSummary", "CollaborationActivitySummary",

        # Endpoint
        "DeviceDetails", "DeviceManagementSummary",

        # Governance
        "SecuritySecureScore", "SecureScoreActions", "SMTPRelaySummary", "RetentionPolicies", "DlpPolicies",
        "PasswordLifecycleSummary", "ExternalSharingSummary", "ExternalSharingSiteOverrides", "ExternalExposureFindings",
        "UnmanagedObjects", "OneDriveOwnerMismatches",

        # Assessment Outputs
        "BestPractices", "BestPracticeFindings", "MigrationReadiness"
    )

    $tenantToTenantCutoverDesiredOrder = @(
        'AllMailboxes',
        'TenantInfo',
        'MigrationExecutiveSummary',
        'RecipientDomainSummary',
        'MailboxMigrationSummary',
        'BitTitanLicenseSummary',
        'DelegateSummary',
        'CollaborationSummary',
        'CutoverPrepSummary',
        'Domains',
        'AdConnectConfiguration',
        'HybridConfiguration',
        'Users',
        'AllRecipients',
        'MailboxDelegateAssignments',
        'MailboxCalendarDelegatePermissions',
        'InactiveMailboxDetails',
        'PublicFolderDetails',
        'MailFlowConnectors',
        'RemoteDomains',
        'SMTPRelayServiceAccounts',
        'AllTeams',
        'SharePoint',
        'OneDrive'
    )
    $tenantToTenantCutoverExcludedWorksheets = @(
        'Admins', 'AuthenticationConfig', 'AuthenticationSSOApplications', 'MfaEnrollmentSummary', 'MfaEnforcementSummary',
        'ConditionalAccessPolicySummary', 'ConditionalAccessPolicies', 'SecurityDefaultsPolicy', 'GuestAccessConfiguration',
        'ExternalIdentityRestrictions', 'MailboxFullDetails', 'ArchiveMailboxes', 'ArchiveMailboxStats', 'UnifiedGroups',
        'SpamFilteringConfig', 'SMTPRelayConfig', 'SharedMailboxGovernanceSummary', 'ForwardingPolicySummary', 'BestPractices',
        'BestPracticeFindings', 'UserFullDetails', 'EntraIDGroups', 'EnterpriseApplications', 'SecuritySecureScore',
        'SecureScoreActions', 'EmailActivityTopSenders', 'EmailActivityTopReceivers', 'TeamsVoiceSummary',
        'CollaborationActivitySummary', 'TeamsActivityTopUsers', 'Office365GroupsActivityTopGroups',
        'EmployeeExperienceInsightsSummary', 'DeviceDetails', 'DeviceManagementSummary', 'LitigationHoldMailboxes',
        'NonUserMailboxes', 'InactiveMailboxes', 'PublicFolderPerms', 'RetentionPolicies', 'DlpPolicies',
        'UnmanagedObjects', 'OneDriveOwnerMismatches', 'ExternalSharingSummary', 'ExternalSharingSiteOverrides',
        'SharePointSharingSummary', 'AllExchangeGroups', 'MigrationReadiness', 'ConditionalAccessOptimization',
        'MfaMethodPostureSummary', 'PrivilegedAccessRemediationSummary', 'TeamsGroupsCleanupCandidates',
        'GroupLicensingSummary', 'LicenseOptimizationCandidates'
    )

    $desiredOrder = $defaultDesiredOrder
    $excludedWorksheets = $defaultExcludedWorksheets
    $appendRemainingWorksheets = $true
    switch ($WorkbookExportPolicy) {
        'TenantToTenantCutover' {
            $desiredOrder = $tenantToTenantCutoverDesiredOrder
            $excludedWorksheets = @($defaultExcludedWorksheets + $tenantToTenantCutoverExcludedWorksheets | Select-Object -Unique)
            $appendRemainingWorksheets = $false
        }
    }

    $orderedTables = @()
    foreach ($name in $desiredOrder) {
        if (($excludedWorksheets -notcontains $name) -and $hashtable.ContainsKey($name)) {
            $orderedTables += $name
        }
    }
    if ($appendRemainingWorksheets) {
        $orderedTables += ($hashtable.Keys | Where-Object { ($orderedTables -notcontains $_) -and ($excludedWorksheets -notcontains $_) } | Sort-Object)
    }

    foreach ($excludedSheet in $excludedWorksheets) {
        if ($hashtable.ContainsKey($excludedSheet)) {
            Write-Log -Type INFO -Message "Skipping worksheet '$excludedSheet' by export policy '$WorkbookExportPolicy'." -ExportFileLocation $ExportDetails
        }
    }

    $totalCount = ($orderedTables | Measure-Object).Count
    $excelPackage = $null
    $autoSizeRowLimit = 1000
    $singleRecordWorksheets = @('TenantInfo')

    try {
        Wait-ForExportFileAvailability -Path $ExportDetails
        if (Test-Path -Path $ExportDetails) {
            Remove-Item -LiteralPath $ExportDetails -Force -ErrorAction Stop
            Write-Log -Type INFO -Message "Removed existing workbook prior to export so the workbook can be rebuilt in a single optimized pass." -ExportFileLocation $ExportDetails
        }

        $excelPackage = Open-ExcelPackage -Path $ExportDetails -Create

        foreach ($table in $orderedTables) {
            try {
                $worksheetName = Resolve-WorksheetExportName -LogicalName $table
                Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash To Excel"
                Write-Log -Type DEBUG -Message ("Exporting '{0}' Hash Table to '{1}' as worksheet '{2}'" -f $table, $ExportDetails, $worksheetName) -ExportFileLocation $ExportDetails

                $sourceTableName = $table
                $tableValue = if ($WorkbookExportPolicy -eq 'TenantToTenantCutover') {
                    $sourceTableName = Get-TenantToTenantWorksheetSourceName -LogicalName $table -ExportTables $hashtable
                    if ($hashtable.ContainsKey($sourceTableName)) { $hashtable[$sourceTableName] } else { $null }
                }
                else {
                    $hashtable[$table]
                }
                $tableValue = Select-WorkbookExportRows -TableName $table -TableValue $tableValue -ExportTables $hashtable
                $sourceInfo = if (($singleRecordWorksheets -contains $table) -and ($sourceTableName -eq $table) -and ($tableValue -is [System.Collections.IDictionary])) {
                    [pscustomobject]@{
                        Count  = 1
                        Source = @($tableValue)
                    }
                }
                else {
                    Get-WorksheetExportSourceInfo -TableValue $tableValue
                }
                $sourceCount = [int]$sourceInfo.Count
                $exportSource = $sourceInfo.Source
                $explicitColumns = if ($WorkbookExportPolicy -eq 'TenantToTenantCutover') { @(Get-TenantToTenantWorksheetColumns -LogicalName $table) } else { @() }

                if ($sourceCount -gt 0) {
                    $autoSizeSheet = ($sourceCount -le $autoSizeRowLimit -and $WorkbookExportPolicy -ne 'TenantToTenantCutover')
                    if (-not $autoSizeSheet) {
                        Write-Log -Type INFO -Message "Skipping AutoSize for worksheet '$worksheetName' due to row count ($sourceCount) to reduce export runtime/memory pressure." -ExportFileLocation $ExportDetails
                    }

                    $excelSplat = @{
                        ExcelPackage = $excelPackage
                        WorksheetName = $worksheetName
                        ClearSheet    = $true
                        BoldTopRow    = $true
                    }
                    if ($autoSizeSheet) {
                        $excelSplat.AutoSize = $true
                    }

                    $rowsToExport = if ($explicitColumns.Count -gt 0) {
                        @($exportSource | Select-Object $explicitColumns)
                    }
                    else {
                        @($exportSource)
                    }

                    $friendlyRows = @($rowsToExport | ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ })
                    $friendlyRows = @(ConvertTo-WorkbookExportUniformRecords -Records $friendlyRows)

                    # On a flattened configuration sheet the Item column holds labels such as
                    # phone numbers ('+15555550100'). Excel would otherwise coerce those to
                    # numbers and drop the leading '+'.
                    if ($friendlyRows.Count -gt 0 -and @($friendlyRows[0].PSObject.Properties.Name) -join ',' -eq 'Section,Item,Value,Notes') {
                        $excelSplat.NoNumberConversion = @('Item')
                    }

                    $friendlyRows |
                        Export-Excel @excelSplat -PassThru |
                        ForEach-Object { $excelPackage = $_ }

                    if ($WorkbookExportPolicy -eq 'TenantToTenantCutover' -and $explicitColumns.Count -gt 0) {
                        Set-TenantToTenantWorksheetColumnWidths -ExcelPackage $excelPackage -WorksheetName $worksheetName -Columns $explicitColumns
                    }
                    continue
                }

                if ($WorkbookExportPolicy -eq 'TenantToTenantCutover' -and $explicitColumns.Count -gt 0) {
                    @(New-ExplicitWorksheetPlaceholderRow -Columns $explicitColumns) |
                        ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                        Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheetName -ClearSheet -BoldTopRow -PassThru |
                        ForEach-Object { $excelPackage = $_ }
                    Set-TenantToTenantWorksheetColumnWidths -ExcelPackage $excelPackage -WorksheetName $worksheetName -Columns $explicitColumns
                    Write-Log -Type INFO -Message "No data found for '$table'; exported an empty-schema worksheet for the tenant-to-tenant workbook contract." -ExportFileLocation $ExportDetails
                    continue
                }

                if ($table -eq 'OneDriveOwnerMismatches') {
                    $placeholderRows = @(
                        [PSCustomObject]@{
                            DisplayName          = 'N/A'
                            SiteUrl              = 'N/A'
                            CurrentOwner         = 'N/A'
                            ExpectedDefaultOwner = 'N/A'
                            Notes                = 'No owner mismatches detected in this run.'
                        }
                    )

                    $placeholderRows |
                        ForEach-Object { ConvertTo-ExportFriendlyRecord -InputObject $_ } |
                        Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheetName -ClearSheet -BoldTopRow -AutoSize -PassThru |
                        ForEach-Object { $excelPackage = $_ }
                    Write-Log -Type INFO -Message "No owner mismatch rows found; exported placeholder row to worksheet '$worksheetName'." -ExportFileLocation $ExportDetails
                    continue
                }

                $logType = if ($optionalEmptySheets -contains $table) { 'INFO' } else { 'WARNING' }
                Write-Log -Type $logType -Message "No data found for $table to export to Excel" -ExportFileLocation $ExportDetails
            }
            catch {
                Write-Log -Type Error -Message "An error occurred in Exporting Hash To Excel for $($table) to '$($ExportDetails)'. $($_.Exception.Message)" -ExportFileLocation $ExportDetails -CaptureError -ErrorRecordVar $_
            }
        }
    }
    finally {
        Write-ProgressHelper -Total $totalCount -Id 2 -Activity "Exporting Hash To Excel" -Completed
        if ($null -ne $excelPackage) {
            try {
                Close-ExcelPackage -ExcelPackage $excelPackage
            }
            catch {
                try {
                    $excelPackage.Save()
                    $excelPackage.Dispose()
                }
                catch {
                    Write-Log -Type WARNING -Message "Unable to close Excel package cleanly: $($_.Exception.Message)" -ExportFileLocation $ExportDetails
                }
            }
        }
    }

    Write-Host "The report has been exported to: $($ExportDetails)" -ForegroundColor Green
}
