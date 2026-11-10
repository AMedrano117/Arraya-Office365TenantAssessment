# User activity lookups and last-activity date resolution shared by reporting surfaces.
# Shared by the HTML assessment report and the tenant improvement plan so both
# deliverables describe the same tenant from one implementation.

function Get-ArrayaNormalizedPrincipalKey {
    [CmdletBinding()]
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    return $Value.Trim().ToLowerInvariant()
}

function New-ArrayaActivityLookup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $false)][string[]]$PrincipalNames = @('UserPrincipalName', 'UPN', 'Mail', 'PrimarySmtpAddress')
    )

    $lookup = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row) {
            continue
        }

        foreach ($principalName in @($PrincipalNames)) {
            $principal = Get-ArrayaNormalizedPrincipalKey -Value ([string](Get-ArrayaObjectValue -Object $row -Names @($principalName)))
            if ([string]::IsNullOrWhiteSpace($principal)) {
                continue
            }

            if (-not $lookup.ContainsKey($principal)) {
                $lookup[$principal] = New-Object System.Collections.Generic.List[object]
            }

            $lookup[$principal].Add($row) | Out-Null
        }
    }

    return $lookup
}

function Get-ArrayaLatestDateFromRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $false)][string[]]$DateNames = @('LastActivityDate', 'LastActivityDateTime', 'LastSignInDateTime', 'LastSuccessfulSignInDateTime')
    )

    $dates = @(
        foreach ($row in @($Rows)) {
            $date = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $row -Names $DateNames)
            if ($null -ne $date) {
                $date
            }
        }
    )

    if ($dates.Count -eq 0) {
        return $null
    }

    return ($dates | Sort-Object -Descending | Select-Object -First 1)
}

function Get-ArrayaDateText {
    [CmdletBinding()]
    param([AllowNull()]$Date)

    if ($null -eq $Date) {
        return 'Not surfaced in current source'
    }

    try {
        return ([datetime]$Date).ToString('yyyy-MM-dd')
    }
    catch {
        return [string]$Date
    }
}

function Get-ArrayaUserLicenseActivityState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$User,
        [Parameter(Mandatory = $false)]$EmailActivityLookup = @{},
        [Parameter(Mandatory = $false)]$TeamsActivityLookup = @{},
        [Parameter(Mandatory = $false)]$Microsoft365ActivityLookup = @{}
    )

    if (-not ($EmailActivityLookup -is [System.Collections.IDictionary])) {
        $EmailActivityLookup = @{}
    }
    if (-not ($TeamsActivityLookup -is [System.Collections.IDictionary])) {
        $TeamsActivityLookup = @{}
    }
    if (-not ($Microsoft365ActivityLookup -is [System.Collections.IDictionary])) {
        $Microsoft365ActivityLookup = @{}
    }

    $principalKeys = @(
        Get-ArrayaNormalizedPrincipalKey -Value ([string](Get-ArrayaObjectValue -Object $User -Names @('UserPrincipalName')))
        Get-ArrayaNormalizedPrincipalKey -Value ([string](Get-ArrayaObjectValue -Object $User -Names @('Mail', 'PrimarySmtpAddress')))
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    $emailRows = New-Object System.Collections.Generic.List[object]
    $teamsRows = New-Object System.Collections.Generic.List[object]
    $m365Rows = New-Object System.Collections.Generic.List[object]
    foreach ($principalKey in @($principalKeys)) {
        if ($EmailActivityLookup.ContainsKey($principalKey)) {
            foreach ($row in @($EmailActivityLookup[$principalKey])) { $emailRows.Add($row) | Out-Null }
        }
        if ($TeamsActivityLookup.ContainsKey($principalKey)) {
            foreach ($row in @($TeamsActivityLookup[$principalKey])) { $teamsRows.Add($row) | Out-Null }
        }
        if ($Microsoft365ActivityLookup.ContainsKey($principalKey)) {
            foreach ($row in @($Microsoft365ActivityLookup[$principalKey])) { $m365Rows.Add($row) | Out-Null }
        }
    }

    $lastSignInDate = Get-ArrayaLatestDateFromRows -Rows @($User) -DateNames @(
        'LastSuccessfulSignInDateTime',
        'LastSignInDateTime',
        'LastNonInteractiveSignInDateTime',
        'SignInActivityLastSignInDateTime',
        'SignInActivityLastSuccessfulSignInDateTime'
    )
    $lastExchangeActivityDate = Get-ArrayaLatestDateFromRows -Rows @(@($emailRows.ToArray()) + @($User)) -DateNames @('LastExchangeActivityDate', 'ExchangeLastActivityDate', 'LastActivityDate', 'LastActivityDateTime')
    $lastTeamsActivityDate = Get-ArrayaLatestDateFromRows -Rows @(@($teamsRows.ToArray()) + @($User)) -DateNames @('LastTeamsActivityDate', 'TeamsLastActivityDate', 'LastActivityDate', 'LastActivityDateTime')
    $lastMicrosoft365ActivityDate = Get-ArrayaLatestDateFromRows -Rows @(@($m365Rows.ToArray()) + @($User)) -DateNames @(
        'LastMicrosoft365ActivityDate',
        'Microsoft365LastActivityDate',
        'LastActivityDate',
        'LastActivityDateTime'
    )
    $availableDates = @(
        $lastSignInDate
        $lastExchangeActivityDate
        $lastTeamsActivityDate
        $lastMicrosoft365ActivityDate
    ) | Where-Object { $null -ne $_ }
    $lastActivityDate = if ($availableDates.Count -gt 0) {
        $availableDates | Sort-Object -Descending | Select-Object -First 1
    }
    else {
        $null
    }

    return [pscustomobject]@{
        LastSignInDateTime          = $lastSignInDate
        LastExchangeActivityDate    = $lastExchangeActivityDate
        LastTeamsActivityDate       = $lastTeamsActivityDate
        LastMicrosoft365ActivityDate = $lastMicrosoft365ActivityDate
        LastActivityDate            = $lastActivityDate
        AvailableSignalCount        = $availableDates.Count
    }
}

