function Get-CustomerAssessmentTemplatePath {
    [CmdletBinding()]
    param()

    $templatePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\..\templates\customer\Microsoft 365 Tenant Best Practices Assessment Template.docx'))
    if (-not (Test-Path -Path $templatePath -PathType Leaf)) {
        throw "Customer assessment template was not found: $templatePath"
    }

    return $templatePath
}

function New-CustomerWordParagraphBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $false)][string]$Style = 'Normal'
    )

    return [pscustomobject]@{
        Type  = 'Paragraph'
        Text  = $Text
        Style = $Style
    }
}

function New-CustomerWordTableBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string[]]$Headers,
        [Parameter(Mandatory = $true)]$Rows,
        [Parameter(Mandatory = $false)][string]$Style
    )

    return [pscustomobject]@{
        Type    = 'Table'
        Headers = @($Headers)
        Rows    = $Rows
        Style   = $Style
    }
}

function New-CustomerWordTableRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Cells
    )

    return [pscustomobject]@{
        Cells = @($Cells)
    }
}

function New-CustomerWordBlankCell {
    [CmdletBinding()]
    param()

    return '__ARRAYA_BLANK__'
}

function Get-CustomerDistinctRoadmapActions {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object[]]$RoadmapActions = @())

    $priorityRanks = @{
        'Critical' = 0
        'High'     = 1
        'Medium'   = 2
        'Low'      = 3
        'Info'     = 4
    }
    $phaseRanks = @{
        'Near Term' = 0
        'Planned'   = 1
        'Monitor'   = 2
    }

    $groupedActions = @(
        @($RoadmapActions | Where-Object { $null -ne $_ }) |
            Group-Object {
                $actionTitle = ([string]$_.ActionTitle).Trim().ToLowerInvariant()
                if ([string]::IsNullOrWhiteSpace($actionTitle)) {
                    $actionTitle = ([string]$_.Theme).Trim().ToLowerInvariant()
                }

                if ([string]::IsNullOrWhiteSpace($actionTitle)) {
                    $actionTitle = ([string]$_.Recommendation).Trim().ToLowerInvariant()
                }

                $actionTitle
            }
    )

    $distinct = foreach ($group in $groupedActions) {
        if ([string]::IsNullOrWhiteSpace([string]$group.Name)) {
            continue
        }

        $bestAction = @($group.Group | Sort-Object `
            @{ Expression = { if ($priorityRanks.ContainsKey([string]$_.Priority)) { $priorityRanks[[string]$_.Priority] } else { 99 } } }, `
            @{ Expression = { if ($phaseRanks.ContainsKey([string]$_.RoadmapPhase)) { $phaseRanks[[string]$_.RoadmapPhase] } else { 99 } } }, `
            @{ Expression = { [string]$_.ActionTitle } }) | Select-Object -First 1

        if ($null -ne $bestAction) {
            $bestAction
        }
    }

    return @($distinct | Sort-Object `
        @{ Expression = { if ($phaseRanks.ContainsKey([string]$_.RoadmapPhase)) { $phaseRanks[[string]$_.RoadmapPhase] } else { 99 } } }, `
        @{ Expression = { if ($priorityRanks.ContainsKey([string]$_.Priority)) { $priorityRanks[[string]$_.Priority] } else { 99 } } }, `
        @{ Expression = { [string]$_.ActionTitle } })
}

function Convert-ToCustomerAssessmentMarkdownHeadingPrefix {
    [CmdletBinding()]
    param([AllowNull()][string]$Style)

    switch ([string]$Style) {
        'Title' { return '# ' }
        'Heading1' { return '## ' }
        'Heading2' { return '### ' }
        'Heading3' { return '#### ' }
        'Heading4' { return '##### ' }
        default { return $null }
    }
}

function Convert-ToCustomerAssessmentMarkdownCellText {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$Content)

    if ($null -eq $Content) {
        return 'Not surfaced in current source'
    }

    if (($Content -is [string]) -and $Content -eq '__ARRAYA_BLANK__') {
        return ''
    }

    if (($Content -is [string]) -and $Content.Length -eq 0) {
        return ''
    }

    $segments = New-Object System.Collections.Generic.List[string]
    $contentQueue = New-Object System.Collections.Queue
    $contentQueue.Enqueue($Content)

    while ($contentQueue.Count -gt 0) {
        $current = $contentQueue.Dequeue()
        if ($null -eq $current) {
            continue
        }

        if (($current -is [System.Collections.IDictionary])) {
            foreach ($dictionaryValue in @($current.Values)) {
                $contentQueue.Enqueue($dictionaryValue)
            }
            continue
        }

        if (($current -is [System.Array] -or $current -is [System.Collections.IList]) -and -not ($current -is [string])) {
            foreach ($item in @($current)) {
                $contentQueue.Enqueue($item)
            }
            continue
        }

        $text = Convert-ToCustomerAssessmentDisplayText -Value $current -Default ''
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $segments.Add(((Convert-ToArrayaMarkdownText $text) -replace '\r?\n', '<br/>')) | Out-Null
        }
    }

    if ($segments.Count -eq 0) {
        return 'Not surfaced in current source'
    }

    return ($segments.ToArray() -join '<br/>')
}

function Convert-CustomerAssessmentBlocksToMarkdown {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object[]]$Blocks)

    $lines = New-Object System.Collections.Generic.List[string]

    foreach ($block in @($Blocks)) {
        if ($null -eq $block) {
            continue
        }

        switch ([string]$block.Type) {
            'Paragraph' {
                $headingPrefix = Convert-ToCustomerAssessmentMarkdownHeadingPrefix -Style ([string]$block.Style)
                $text = Convert-ToArrayaDisplayText -Value $block.Text -Default ''
                if ([string]::IsNullOrWhiteSpace($text)) {
                    continue
                }

                if (-not [string]::IsNullOrWhiteSpace($headingPrefix)) {
                    $lines.Add($headingPrefix + $text) | Out-Null
                }
                elseif ([string]$block.Style -eq 'Subtitle') {
                    $lines.Add("**$text**") | Out-Null
                }
                else {
                    $lines.Add($text) | Out-Null
                }
                $lines.Add('') | Out-Null
            }
            'Table' {
                $headers = @($block.Headers | ForEach-Object { Convert-ToCustomerAssessmentMarkdownCellText -Content $_ })
                if ($headers.Count -eq 0) {
                    continue
                }

                $lines.Add('| ' + ($headers -join ' | ') + ' |') | Out-Null
                $lines.Add('| ' + (@($headers | ForEach-Object { '---' }) -join ' | ') + ' |') | Out-Null

                foreach ($row in @($block.Rows)) {
                    $cells = if ($null -eq $row) {
                        @()
                    }
                    elseif ($row.PSObject.Properties.Name -contains 'Cells') {
                        @($row.Cells)
                    }
                    elseif (($row -is [System.Collections.IEnumerable]) -and -not ($row -is [string])) {
                        @($row)
                    }
                    else {
                        @($row)
                    }

                    while ($cells.Count -lt $headers.Count) {
                        $cells += ''
                    }

                    $cellTexts = @(
                        $cells |
                            Select-Object -First $headers.Count |
                            ForEach-Object { Convert-ToCustomerAssessmentMarkdownCellText -Content $_ }
                    )
                    $lines.Add('| ' + ($cellTexts -join ' | ') + ' |') | Out-Null
                }

                $lines.Add('') | Out-Null
            }
        }
    }

    return ($lines -join [Environment]::NewLine).Trim() + [Environment]::NewLine
}

function Write-CustomerAssessmentMarkdownFromBlocks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][object[]]$Blocks
    )

    $markdown = Convert-CustomerAssessmentBlocksToMarkdown -Blocks $Blocks
    [System.IO.File]::WriteAllText($OutputPath, $markdown, [System.Text.UTF8Encoding]::new($false))
}

function Convert-ToCustomerAssessmentDisplayText {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory = $false)][string]$Default = 'Not validated from the reviewed data'
    )

    $normalized = Convert-ToArrayaDisplayText -Value $Value -Default $Default
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $Default
    }

    switch -Regex ($normalized) {
        '^(?i)not surfaced in current source$' { return 'Not validated from the reviewed data' }
        '^(?i)not collected$' { return 'Not validated from the reviewed data' }
        '^(?i)not available in current auth mode$' { return 'Not validated in current auth mode' }
        default { return $normalized }
    }
}

function Convert-ToCustomerAssessmentMfaMethodBreakdownText {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory = $false)][string]$Default = 'Not validated from the reviewed data'
    )

    $normalized = Convert-ToCustomerAssessmentDisplayText -Value $Value -Default $Default
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $Default
    }

    return (
        $normalized `
            -replace '(?i)\bs\s+of\s+tw\s+ar\s+eO\s+ne\s+Ti\s+me\s+Pa\s+ss\s+co\s+de\b', 'Software one-time passcode' `
            -replace '(?i)\bsoftwareOneTimePasscode\b', 'Software one-time passcode' `
            -replace '(?i)\bsoftwareOathAuthenticationMethodConfiguration\b', 'Software OATH token'
    )
}

function Convert-ToCustomerAssessmentBooleanLabel {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory = $false)][string]$TrueText = 'Yes',
        [Parameter(Mandatory = $false)][string]$FalseText = 'No',
        [Parameter(Mandatory = $false)][string]$Default = 'Not validated from the reviewed data'
    )

    $booleanValue = Convert-ToArrayaBoolean $Value
    if ($null -ne $booleanValue) {
        return $(if ($booleanValue) { $TrueText } else { $FalseText })
    }

    return Convert-ToCustomerAssessmentDisplayText -Value $Value -Default $Default
}

function Convert-ToCustomerAssessmentNarrativeText {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    $normalized = Convert-ToCustomerAssessmentDisplayText -Value $Text -Default ''
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $normalized = $normalized -replace '\bcurrent source\b', 'tenant review'
    $normalized = $normalized -replace '(?i)\bnot surfaced in current source\b', 'not validated from the reviewed data'
    $normalized = $normalized -replace '(?i)\bnot available in current auth mode\b', 'not validated in current auth mode'
    $normalized = $normalized -replace '\bThis matters in this tenant because\b', 'In this tenant, this matters because'
    $normalized = $normalized -replace '\bWhat is working well is that\b', 'A positive signal in the current state is that'
    return $normalized
}

function Get-CustomerEnterpriseApplicationSignalState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [string]$Default = 'Unavailable'
    )

    $state = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @($Name)) -Default $Default
    if ([string]::IsNullOrWhiteSpace($state)) {
        return $Default
    }

    return $state
}

function Test-CustomerEnterpriseApplicationSsoEnabled {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    $preferredSingleSignOnMode = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('SSOMode', 'PreferredSingleSignOnMode')) -Default ''
    $explicitSsoEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('SsoEnabled'))
    return (
        ($explicitSsoEnabled -eq $true) -or
        (-not [string]::IsNullOrWhiteSpace($preferredSingleSignOnMode) -and $preferredSingleSignOnMode -ne 'notSupported')
    )
}

function Get-CustomerEnterpriseApplicationCoverageState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$TotalCount,
        [Parameter(Mandatory = $false)]
        [int]$CollectedCount = 0,
        [Parameter(Mandatory = $false)]
        [int]$PartialCount = 0,
        [Parameter(Mandatory = $false)]
        [int]$UnavailableCount = 0,
        [Parameter(Mandatory = $false)]
        [int]$NotApplicableCount = 0,
        [Parameter(Mandatory = $false)]
        [switch]$SupportsNotApplicable
    )

    if ($TotalCount -eq 0) {
        return 'Unavailable'
    }

    if ($SupportsNotApplicable -and (($CollectedCount + $PartialCount + $UnavailableCount) -eq 0) -and $NotApplicableCount -gt 0) {
        return 'NotApplicable'
    }

    if (($PartialCount + $UnavailableCount) -eq 0) {
        return 'Collected'
    }

    if (($CollectedCount + $PartialCount) -gt 0) {
        return 'Partial'
    }

    return 'Unavailable'
}

function Get-CustomerEnterpriseApplicationSummaryFromRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object[]]$ApplicationRows = @()
    )

    $rows = @($ApplicationRows)
    $applicationSourceCollectedCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ApplicationSourceState') -eq 'Collected' }).Count
    $applicationSourceUnavailableCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ApplicationSourceState') -eq 'Unavailable' }).Count
    $ownerSignalCollectedCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'Collected' }).Count
    $ownerSignalUnavailableCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'Unavailable' }).Count
    $ownerSignalNotApplicableCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'NotApplicable' }).Count
    $redirectUriSignalCollectedCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState') -eq 'Collected' }).Count
    $redirectUriSignalPartialCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState') -eq 'Partial' }).Count
    $redirectUriSignalUnavailableCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState') -eq 'Unavailable' }).Count
    $activitySignalCollectedCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Collected' }).Count
    $activitySignalPartialCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Partial' }).Count
    $activitySignalUnavailableCount = @($rows | Where-Object { (Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Unavailable' }).Count

    return [pscustomobject]@{
        TotalEnterpriseApplications      = $rows.Count
        FirstPartyApplications           = @($rows | Where-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'First Party' }).Count
        ThirdPartyApplications           = @($rows | Where-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'Third Party' }).Count
        UnknownSourceApplications        = @($rows | Where-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'Unknown' }).Count
        ApplicationsWithHighPrivilege    = @($rows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissionCount'))) -gt 0 }).Count
        ApplicationsWithDelegatedGrants  = @($rows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('DelegatedPermissionGrantCount'))) -gt 0 }).Count
        ApplicationsWithApplicationPerms = @($rows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount'))) -gt 0 }).Count
        FirstPartyAppsWithoutOwners      = @($rows | Where-Object {
            ([string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'First Party') -and
            ((Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'Collected') -and
            ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))) -le 0)
        }).Count
        ThirdPartyAppsWithApplicationPerms = @($rows | Where-Object {
            ([string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'Third Party') -and
            ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount'))) -gt 0)
        }).Count
        ApplicationsWithNoRecentActivity = @($rows | Where-Object {
            ((Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Collected') -and
            ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasRecentActivity'))) -eq $false)
        }).Count
        ApplicationsWithInsecureRedirectUris = @($rows | Where-Object {
            $redirectSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState'
            (($redirectSignalState -eq 'Collected') -or ($redirectSignalState -eq 'Partial')) -and
            ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasInsecureRedirectUris'))) -eq $true)
        }).Count
        SsoEnabledApplications           = @($rows | Where-Object { Test-CustomerEnterpriseApplicationSsoEnabled -ApplicationRow $_ }).Count
        ApplicationSourceCoverageState   = Get-CustomerEnterpriseApplicationCoverageState -TotalCount $rows.Count -CollectedCount $applicationSourceCollectedCount -UnavailableCount $applicationSourceUnavailableCount
        ApplicationSourceCollectedCount  = $applicationSourceCollectedCount
        ApplicationSourceUnavailableCount = $applicationSourceUnavailableCount
        OwnerSignalCoverageState         = Get-CustomerEnterpriseApplicationCoverageState -TotalCount $rows.Count -CollectedCount $ownerSignalCollectedCount -UnavailableCount $ownerSignalUnavailableCount -NotApplicableCount $ownerSignalNotApplicableCount -SupportsNotApplicable
        OwnerSignalCollectedCount        = $ownerSignalCollectedCount
        OwnerSignalUnavailableCount      = $ownerSignalUnavailableCount
        OwnerSignalNotApplicableCount    = $ownerSignalNotApplicableCount
        RedirectUriSignalCoverageState   = Get-CustomerEnterpriseApplicationCoverageState -TotalCount $rows.Count -CollectedCount $redirectUriSignalCollectedCount -PartialCount $redirectUriSignalPartialCount -UnavailableCount $redirectUriSignalUnavailableCount
        RedirectUriSignalCollectedCount  = $redirectUriSignalCollectedCount
        RedirectUriSignalPartialCount    = $redirectUriSignalPartialCount
        RedirectUriSignalUnavailableCount = $redirectUriSignalUnavailableCount
        ActivitySignalCoverageState      = Get-CustomerEnterpriseApplicationCoverageState -TotalCount $rows.Count -CollectedCount $activitySignalCollectedCount -PartialCount $activitySignalPartialCount -UnavailableCount $activitySignalUnavailableCount
        ActivitySignalCollectedCount     = $activitySignalCollectedCount
        ActivitySignalPartialCount       = $activitySignalPartialCount
        ActivitySignalUnavailableCount   = $activitySignalUnavailableCount
    }
}

function Test-CustomerEnterpriseApplicationNoteworthy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    $highPrivilegeCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HighPrivilegePermissionCount'))
    $applicationPermissionCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('ApplicationPermissionCount'))
    $delegatedPermissionGrantCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('DelegatedPermissionGrantCount'))
    $applicationSource = [string](Get-ArrayaObjectValue -Object $ApplicationRow -Names @('ApplicationSource'))
    $ownerSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'OwnerSignalState'
    $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('OwnerCount'))
    $redirectSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'RedirectUriSignalState'
    $hasInsecureRedirectUris = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HasInsecureRedirectUris'))
    $activitySignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'ActivitySignalState'
    $hasRecentActivity = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HasRecentActivity'))
    $appRoleAssignmentRequired = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('AppRoleAssignmentRequired'))

    return (
        (($null -ne $highPrivilegeCount) -and ($highPrivilegeCount -gt 0)) -or
        (($null -ne $applicationPermissionCount) -and ($applicationPermissionCount -gt 0)) -or
        (($null -ne $delegatedPermissionGrantCount) -and ($delegatedPermissionGrantCount -gt 0)) -or
        (Test-CustomerEnterpriseApplicationSsoEnabled -ApplicationRow $ApplicationRow) -or
        (($applicationSource -eq 'First Party') -and ($ownerSignalState -eq 'Collected') -and ($null -ne $ownerCount) -and ($ownerCount -le 0)) -or
        ((($redirectSignalState -eq 'Collected') -or ($redirectSignalState -eq 'Partial')) -and ($hasInsecureRedirectUris -eq $true)) -or
        (($activitySignalState -eq 'Collected') -and ($hasRecentActivity -eq $false)) -or
        ($appRoleAssignmentRequired -eq $true)
    )
}

function Get-CustomerNoteworthyEnterpriseApplications {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object[]]$ApplicationRows = @(),
        [Parameter(Mandatory = $false)]
        [int]$Top = 8
    )

    $rows = @($ApplicationRows)
    if ($rows.Count -eq 0) {
        return @()
    }

    $rankedNoteworthyRows = @(
        $rows |
            Where-Object { Test-CustomerEnterpriseApplicationNoteworthy -ApplicationRow $_ } |
            Sort-Object `
                @{ Expression = { $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissionCount')); if ($null -eq $value) { 0 } else { $value } }; Descending = $true }, `
                @{ Expression = { $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount')); if ($null -eq $value) { 0 } else { $value } }; Descending = $true }, `
                @{ Expression = { $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('DelegatedPermissionGrantCount')); if ($null -eq $value) { 0 } else { $value } }; Descending = $true }, `
                @{ Expression = { if (Test-CustomerEnterpriseApplicationSsoEnabled -ApplicationRow $_) { 1 } else { 0 } }; Descending = $true }, `
                @{ Expression = {
                    $ownerState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState'
                    $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))
                    if (([string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'First Party') -and ($ownerState -eq 'Collected') -and ($null -ne $ownerCount) -and ($ownerCount -le 0)) { 1 } else { 0 }
                }; Descending = $true }, `
                @{ Expression = {
                    $redirectState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState'
                    $hasRisk = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasInsecureRedirectUris'))
                    if ((($redirectState -eq 'Collected') -or ($redirectState -eq 'Partial')) -and ($hasRisk -eq $true)) { 1 } else { 0 }
                }; Descending = $true }, `
                @{ Expression = {
                    $activityState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState'
                    $hasRecentActivity = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasRecentActivity'))
                    if (($activityState -eq 'Collected') -and ($hasRecentActivity -eq $false)) { 1 } else { 0 }
                }; Descending = $true }, `
                @{ Expression = { if ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AppRoleAssignmentRequired'))) -eq $true) { 1 } else { 0 } }; Descending = $true }, `
                @{ Expression = { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed application' } }
    )

    if ($rankedNoteworthyRows.Count -ge $Top) {
        return @($rankedNoteworthyRows | Select-Object -First $Top)
    }

    $remainingRows = @(
        $rows |
            Where-Object { $rankedNoteworthyRows -notcontains $_ } |
            Sort-Object @{ Expression = { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed application' } }
    )

    return @(@($rankedNoteworthyRows) + @($remainingRows | Select-Object -First ($Top - $rankedNoteworthyRows.Count)))
}

function Get-CustomerTechnicalObservationByTitle {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$TechnicalObservations = @(),
        [Parameter(Mandatory = $true)][string]$Title
    )

    return @($TechnicalObservations | Where-Object { [string]$_.SectionTitle -eq $Title } | Select-Object -First 1)
}

function Get-CustomerObservationState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Observation,
        [Parameter(Mandatory = $true)][string]$Signal
    )

    if ($null -eq $Observation) {
        return 'Not validated from the reviewed data'
    }

    $row = @($Observation.ConfigurationRows | Where-Object { [string]$_.Signal -eq $Signal } | Select-Object -First 1)
    if ($null -eq $row -or $row.Count -eq 0) {
        return 'Not validated from the reviewed data'
    }

    return Convert-ToCustomerAssessmentDisplayText -Value $row[0].State -Default 'Not validated from the reviewed data'
}

function Convert-CustomerConfigurationRowsToWordTableRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $false)][string]$DefaultText = 'Not validated from the reviewed data'
    )

    if (@($Rows).Count -eq 0) {
        return @((New-CustomerWordTableRow -Cells @($DefaultText, $DefaultText)))
    }

    return @(
        $Rows | ForEach-Object {
            New-CustomerWordTableRow -Cells @(
                (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('Signal', 'Term')) -Default $DefaultText),
                (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('State', 'Meaning')) -Default $DefaultText)
            )
        }
    )
}

function Convert-CustomerThreeColumnRowsToWordTableRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $true)][string[]]$PropertyNames,
        [Parameter(Mandatory = $false)][string]$DefaultText = 'Not validated from the reviewed data'
    )

    if (@($Rows).Count -eq 0) {
        return @((New-CustomerWordTableRow -Cells @($DefaultText, $DefaultText, $DefaultText)))
    }

    return @(
        $Rows | ForEach-Object {
            New-CustomerWordTableRow -Cells @(
                (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @($PropertyNames[0])) -Default $DefaultText),
                (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @($PropertyNames[1])) -Default $DefaultText),
                (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @($PropertyNames[2])) -Default $DefaultText)
            )
        }
    )
}

function Convert-ToCustomerAssessmentCollectionRows {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [string[]]$MarkerNames = @()
    )

    if ($null -eq $Value) {
        return @()
    }

    $looksLikeRow = $false
    if ($MarkerNames.Count -gt 0) {
        $looksLikeRow = $null -ne (Get-ArrayaObjectValue -Object $Value -Names $MarkerNames)
    }

    if ($looksLikeRow) {
        if ($Value -is [System.Collections.IDictionary]) {
            return @([pscustomobject]$Value)
        }

        return @($Value)
    }

    if ($Value -is [System.Collections.IDictionary]) {
        return @(
            @($Value.Values) |
                Where-Object { $null -ne $_ } |
                ForEach-Object {
                    if ($MarkerNames.Count -gt 0 -and $null -eq (Get-ArrayaObjectValue -Object $_ -Names $MarkerNames)) {
                        return
                    }

                    if ($_ -is [System.Collections.IDictionary]) {
                        [pscustomobject]$_
                    }
                    else {
                        $_
                    }
                }
        )
    }

    return @(
        (Convert-ArrayaObjectToArray $Value) |
            Where-Object {
                if ($null -eq $_) {
                    return $false
                }

                if ($MarkerNames.Count -gt 0) {
                    return $null -ne (Get-ArrayaObjectValue -Object $_ -Names $MarkerNames)
                }

                return $true
            } |
            ForEach-Object {
                if ($_ -is [System.Collections.IDictionary]) {
                    [pscustomobject]$_
                }
                else {
                    $_
                }
            }
    )
}

function Get-CustomerMfaGapCategoryLabel {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $text = Convert-ToCustomerAssessmentDisplayText -Value $Value -Default 'Not validated from the reviewed data'
    switch ($text.ToLowerInvariant()) {
        'excluded from enabled mfa ca policy' { return 'Explicitly excluded' }
        'outside enabled mfa ca include scope' { return 'Outside include scope' }
        default { return $text }
    }
}

function Get-CustomerActionImpactLabel {
    [CmdletBinding()]
    param([AllowNull()][string]$Severity)

    switch (([string]$Severity).ToLowerInvariant()) {
        'critical' { return 'High' }
        'high' { return 'High' }
        'medium' { return 'Medium' }
        'low' { return 'Low' }
        'info' { return 'Low' }
        default { return 'Not validated from the reviewed data' }
    }
}

function Get-CustomerSeverityRank {
    [CmdletBinding()]
    param([AllowNull()][string]$Severity)

    switch (([string]$Severity).ToLowerInvariant()) {
        'critical' { return 5 }
        'high' { return 4 }
        'medium' { return 3 }
        'low' { return 2 }
        'info' { return 1 }
        default { return 0 }
    }
}

function Get-CustomerPriorityRank {
    [CmdletBinding()]
    param([AllowNull()][string]$Priority)

    switch (([string]$Priority).ToLowerInvariant()) {
        'immediate' { return 1 }
        'near term' { return 2 }
        'planned' { return 3 }
        'monitor' { return 4 }
        default { return 5 }
    }
}

function Get-CustomerGuestUserRoleLabel {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $text = Convert-ToCustomerAssessmentDisplayText -Value $Value -Default 'Not validated from the reviewed data'
    if ([string]::IsNullOrWhiteSpace($text)) {
        return 'Not validated from the reviewed data'
    }

    if ($text -match '^[0-9a-fA-F-]{36}$') {
        return 'Custom guest role configured'
    }

    return $text
}

function Convert-ToCustomerInvitationControlLabel {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $text = Convert-ToCustomerAssessmentDisplayText -Value $Value -Default 'Not validated from the reviewed data'
    switch ($text.ToLowerInvariant()) {
        'everyone' { return 'Anyone in the organization can invite guests' }
        'adminsandguestinviters' { return 'Admins and approved guest inviters can invite guests' }
        'adminsguestinvitersandallmembers' { return 'Admins, approved guest inviters, and members can invite guests' }
        'adminsonly' { return 'Only admins can invite guests' }
        default { return $text }
    }
}

function Convert-ToCustomerSharingCapabilityLabel {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $text = Convert-ToCustomerAssessmentDisplayText -Value $Value -Default 'Not validated from the reviewed data'
    switch ($text.ToLowerInvariant()) {
        'externaluserandguestsharing' { return 'New and existing guests can be used for external sharing' }
        'externalusersharingonly' { return 'Only existing guests can be used for external sharing' }
        'disabled' { return 'Only people in the organization can access shared content' }
        'existingexternalusersharingonly' { return 'Only existing external users can be used for external sharing' }
        default { return $text }
    }
}

function Convert-ToCustomerSharingRestrictionModeLabel {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $text = Convert-ToCustomerAssessmentDisplayText -Value $Value -Default 'Not validated from the reviewed data'
    switch ($text.ToLowerInvariant()) {
        'none' { return 'No domain allow/block list is applied' }
        'allowlist' { return 'Only explicitly allowed domains may be shared externally' }
        'blocklist' { return 'Specific blocked domains are prevented from external sharing' }
        default { return $text }
    }
}

function Convert-ToCustomerSharingLinkTypeLabel {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $text = Convert-ToCustomerAssessmentDisplayText -Value $Value -Default 'Not validated from the reviewed data'
    switch ($text.ToLowerInvariant()) {
        'anonymousaccess' { return 'Anyone link' }
        'internal' { return 'People in your organization' }
        'direct' { return 'Specific people' }
        'organization' { return 'People in your organization' }
        default { return $text }
    }
}

function Convert-ToCustomerFindingSummaryText {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    $normalized = Convert-ToCustomerAssessmentDisplayText -Value $Text -Default ''
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $matchText = $normalized.ToLowerInvariant()
    switch -Regex ($matchText) {
        'license sku|skus are near or at capacity|at capacity' { return 'Multiple Microsoft 365 SKUs are fully consumed or near capacity and should be reviewed against current assignment and growth.' }
        'global administrator count exceeds' { return 'Standing Global Administrator access remains above the recommended operating threshold and should be reduced.' }
        'break-glass|emergency access' { return 'Emergency access coverage should be confirmed because the expected break-glass design is not yet fully visible in the reviewed state.' }
        'inactive guest accounts' { return 'Inactive guest accounts remain present and should be reviewed against current sponsorship and business need.' }
        'guest lifecycle hygiene' { return 'Guest lifecycle hygiene needs review so dormant external access does not remain in place by default.' }
        'risk-based conditional access' { return 'Risk-based Conditional Access is not clearly visible in the current baseline and should be confirmed against the intended access model.' }
        'report-only conditional access|none appear enabled' { return 'Conditional Access is still carrying report-only or non-enforced patterns that should be progressed into the supported baseline.' }
        'sharepoint sharing is configured to allow external sharing' { return 'Tenant-wide external sharing is enabled and should be validated against the intended collaboration baseline.' }
        'anonymous links appear to remain the default' { return 'Anonymous-link style sharing defaults should be reviewed against named-user collaboration requirements.' }
        'externally exposed.*stale|ownership drift' { return 'Externally exposed collaboration locations show stale activity or ownership drift and should be reviewed before they remain accessible.' }
        'guest-heavy|dormant guest-enabled' { return 'Guest-enabled Teams or groups show ownership or activity patterns that deserve cleanup review.' }
        'inbox rules with external forwarding' { return 'External forwarding paths are present and should be validated against the approved mail-flow baseline.' }
        'shared mailbox governance' { return 'Shared mailboxes show ownership or growth gaps that should be reconciled before they continue to sprawl.' }
        'unverified domains were detected' { return 'Unverified domains remain in the tenant and should be confirmed or removed.' }
        'spf|dkim|dmarc|spoof' { return 'Mail-authentication and anti-spoofing controls still need follow-through across the accepted domain set.' }
        'device stale rate is above|stale devices' { return 'A meaningful portion of the device estate appears stale, which weakens lifecycle and compliance discipline.' }
        'non-compliant device|device compliance' { return 'Device compliance results are below the expected baseline and should be reviewed with endpoint operations.' }
        'secure score' { return 'Baseline security improvement opportunities remain open and should be sequenced into the agreed operating plan.' }
        default {
            $sentence = $normalized.Trim()
            if (-not $sentence.EndsWith('.')) {
                $sentence += '.'
            }

            return $sentence
        }
    }
}

function Get-CustomerOverallFindingsSummaryRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$WorkstreamSummaries = @(),
        [Parameter(Mandatory = $false)][object[]]$Findings = @()
    )

    if (@($WorkstreamSummaries).Count -eq 0) {
        return @((New-CustomerWordTableRow -Cells @('Not validated from the reviewed data', 'Not validated from the reviewed data', '0', 'No workstream summary rows were available in the reviewed data.')))
    }

    $rows = New-Object System.Collections.Generic.List[object]
    $groupedSummaries = @(
        $WorkstreamSummaries |
            Group-Object { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('Workstream', 'OwnerTeam')) -Default 'Not validated from the reviewed data' } |
            Sort-Object `
                @{ Expression = { ($_.Group | ForEach-Object { Get-CustomerSeverityRank -Severity (Get-ArrayaObjectValue -Object $_ -Names @('Severity')) } | Measure-Object -Maximum).Maximum }; Descending = $true }, `
                @{ Expression = { ($_.Group | ForEach-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OpenFindings')) } | Measure-Object -Sum).Sum }; Descending = $true }, `
                Name
    )

    foreach ($group in $groupedSummaries) {
        $workstream = [string]$group.Name
        $maxSeverity = @(
            $group.Group |
                ForEach-Object { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('Severity')) -Default 'Not validated from the reviewed data' } |
                Sort-Object { Get-CustomerSeverityRank -Severity $_ } -Descending
        ) | Select-Object -First 1

        $openFindings = @(
            $group.Group |
                ForEach-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OpenFindings')) } |
                Where-Object { $null -ne $_ } |
                Measure-Object -Sum
        ).Sum

        if ($null -eq $openFindings) {
            $openFindings = @(
                $Findings | Where-Object { ([string]$_.OwnerTeam) -eq $workstream }
            ).Count
        }

        $relatedFindings = @(
            $Findings |
                Where-Object { ([string]$_.OwnerTeam) -eq $workstream } |
                Sort-Object `
                    @{ Expression = { Get-CustomerSeverityRank -Severity ([string]$_.Severity) }; Descending = $true }, `
                    @{ Expression = { Get-CustomerPriorityRank -Priority ([string]$_.PriorityBand) } }, `
                    RuleId
        )

        $standoutItems = @(
            $relatedFindings |
                ForEach-Object { Convert-ToCustomerFindingSummaryText -Text ([string]$_.Finding) } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique |
                Select-Object -First 2
        )

        if ($standoutItems.Count -eq 0) {
            $standoutItems = @(
                $group.Group |
                    ForEach-Object { Convert-ToCustomerFindingSummaryText -Text ([string](Get-ArrayaObjectValue -Object $_ -Names @('TopSignals'))) } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Select-Object -Unique |
                    Select-Object -First 2
            )
        }

        $whatStandsOut = if ($standoutItems.Count -gt 1) {
            ($standoutItems -join ' ')
        }
        elseif ($standoutItems.Count -eq 1) {
            $standoutItems[0]
        }
        else {
            'The reviewed signals in this workstream did not surface one concise standout summary, so use 15.10 Full Findings Inventory for the detailed item list.'
        }

        $rows.Add((New-CustomerWordTableRow -Cells @(
            $workstream,
            $maxSeverity,
            $(if ($null -eq $openFindings) { '0' } else { [string]$openFindings }),
            $whatStandsOut
        ))) | Out-Null
    }

    return @($rows.ToArray())
}

function Get-CustomerLeadershipDecisionText {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Action)

    switch ([string]$Action.ActionTitle) {
        'Reduce privileged access and strengthen identity controls' { return 'Approve the reduction of standing privileged access and the move from observation-only identity controls into the supported baseline.' }
        'Establish accountable ownership for collaboration spaces' { return 'Approve an ownership model for collaboration spaces, including when stale workspaces should be retained, transferred, or retired.' }
        'Strengthen domain and anti-spoofing controls' { return 'Approve the domain and mail-authentication baseline so verification, SPF, DKIM, and DMARC cleanup can be completed consistently.' }
        'Improve device compliance and managed endpoint coverage' { return 'Approve the expected endpoint baseline and the exception path for devices that should not retain access while non-compliant.' }
        'Review external forwarding and mail flow exposure' { return 'Approve the external forwarding and mail-flow standard so exceptions can be formally validated or removed.' }
        'Reconcile license capacity and tenant governance gaps' { return 'Approve the capacity and governance cleanup path so constrained licensing and cross-tenant settings can be reconciled together.' }
        'Strengthen baseline security and access protections' { return 'Approve the security baseline changes required to close the highest-value protection gaps first.' }
        default { return 'Approve the operating model, accountable owner, and remediation sequence for this work item.' }
    }
}

function Get-CustomerLeadershipDecisionRows {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object[]]$RoadmapActions = @())

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($action in @(Get-CustomerDistinctRoadmapActions -RoadmapActions $RoadmapActions | Select-Object -First 3)) {
        $rows.Add((New-CustomerWordTableRow -Cells @(
            (Convert-ToCustomerAssessmentDisplayText -Value $action.ActionTitle -Default 'Priority work item'),
            (Get-CustomerLeadershipDecisionText -Action $action),
            (Convert-ToCustomerAssessmentNarrativeText -Text $action.WhyItMatters)
        ))) | Out-Null
    }

    if ($rows.Count -eq 0) {
        $rows.Add((New-CustomerWordTableRow -Cells @(
            'Priority work item not clearly surfaced',
            'Approve the remediation path that best matches the reviewed evidence.',
            'The current source did not include a clear roadmap action ordering, so 15.10 Full Findings Inventory should be used for the detailed crosswalk.'
        ))) | Out-Null
    }

    return @($rows.ToArray())
}

function Get-CustomerRelevantRoadmapActions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$RoadmapActions = @(),
        [Parameter(Mandatory = $false)][string[]]$Workstreams = @(),
        [Parameter(Mandatory = $false)][string[]]$Keywords = @(),
        [Parameter(Mandatory = $false)][int]$Max = 3
    )

    $filtered = @(Get-CustomerDistinctRoadmapActions -RoadmapActions $RoadmapActions)
    if (@($Workstreams).Count -gt 0) {
        $filtered = @($filtered | Where-Object { $Workstreams -contains [string]$_.Workstream })
    }

    if (@($Keywords).Count -gt 0) {
        $keywordRegex = ($Keywords | ForEach-Object { [regex]::Escape($_) }) -join '|'
        $filtered = @(
            $filtered |
                Where-Object {
                    ([string]$_.ActionTitle -match $keywordRegex) -or
                    ([string]$_.Theme -match $keywordRegex) -or
                    ([string]$_.WhatThisAddresses -match $keywordRegex)
                }
        )
    }

    return @($filtered | Select-Object -First $Max)
}

function Get-CustomerRelevantFindings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Findings = @(),
        [Parameter(Mandatory = $false)][string[]]$Keywords = @(),
        [Parameter(Mandatory = $false)][int]$Max = 3
    )

    $filtered = @($Findings)
    if (@($Keywords).Count -gt 0) {
        $keywordRegex = ($Keywords | ForEach-Object { [regex]::Escape($_) }) -join '|'
        $filtered = @(
            $filtered |
                Where-Object {
                    ([string]$_.Area -match $keywordRegex) -or
                    ([string]$_.Category -match $keywordRegex) -or
                    ([string]$_.RuleId -match $keywordRegex) -or
                    ([string]$_.Finding -match $keywordRegex) -or
                    ([string]$_.Recommendation -match $keywordRegex)
                }
        )
    }

    return @($filtered | Select-Object -First $Max)
}

function Get-CustomerAssessmentAppendixSections {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject]@{
            Title      = '15.1 Device Management'
            Intro      = 'These Microsoft references support the device management, compliance, and managed access observations documented in this assessment.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Get started with device compliance policies in Microsoft Intune' -Url 'https://learn.microsoft.com/en-us/intune/intune-service/protect/device-compliance-get-started' -WhyItIsRelevant 'Supports the compliance baseline and managed-device observations in the endpoint review.'),
                (New-CustomerDocumentationReference -Title 'Require compliant or hybrid Microsoft Entra joined device' -Url 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-all-users-device-compliance' -WhyItIsRelevant 'Provides Microsoft guidance for tying device state to access-control enforcement.')
            )
        },
        [pscustomobject]@{
            Title      = '15.2 Entra Guest Access Best Practices'
            Intro      = 'These references support the observations related to guest lifecycle, external collaboration, and guest governance.'
            References = @(
                (New-CustomerDocumentationReference -Title 'B2B collaboration fundamentals' -Url 'https://learn.microsoft.com/en-us/entra/external-id/b2b-fundamentals' -WhyItIsRelevant 'Supports guest access governance and external collaboration design.'),
                (New-CustomerDocumentationReference -Title 'Overview of external sharing in SharePoint and OneDrive' -Url 'https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview' -WhyItIsRelevant 'Provides Microsoft guidance for the collaboration-sharing observations in this report.')
            )
        },
        [pscustomobject]@{
            Title      = '15.3 Application Consent and Authentication Methods'
            Intro      = 'These Microsoft references support the application governance, consent, and authentication-method observations in the tenant review.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Configure the admin consent workflow' -Url 'https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-admin-consent-workflow' -WhyItIsRelevant 'Relevant to application governance and approval workflow maturity.'),
                (New-CustomerDocumentationReference -Title 'How to manage authentication methods' -Url 'https://learn.microsoft.com/en-us/entra/identity/authentication/how-to-authentication-methods-manage' -WhyItIsRelevant 'Supports recommendations tied to MFA registration and authentication-method governance.')
            )
        },
        [pscustomobject]@{
            Title      = '15.4 Access and Privileged Identity Management'
            Intro      = 'These references support the Conditional Access, privileged-access, and standing-admin observations identified in the assessment.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Conditional Access overview' -Url 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview' -WhyItIsRelevant 'Supports the policy coverage, exclusions, and enforcement observations.'),
                (New-CustomerDocumentationReference -Title 'Privileged Identity Management overview' -Url 'https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure' -WhyItIsRelevant 'Supports the recommendations related to privileged role hygiene and reducing standing access.')
            )
        },
        [pscustomobject]@{
            Title      = '15.5 Exchange Online Archives and SMTP Relay'
            Intro      = 'These references support the messaging, forwarding, transport, archive, and relay observations documented in the Exchange review.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Control automatic external email forwarding in Microsoft 365' -Url 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-policies-external-email-forwarding' -WhyItIsRelevant 'Supports the forwarding-control observations in the messaging section.'),
                (New-CustomerDocumentationReference -Title 'How to set up a multifunction device or application to send email using Microsoft 365 or Office 365' -Url 'https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-microsoft-365-or-office-365' -WhyItIsRelevant 'Relevant to SMTP relay, connector, and mail-flow exception handling.')
            )
        },
        [pscustomobject]@{
            Title      = '15.6 Microsoft Teams and SharePoint Online'
            Intro      = 'These references support the collaboration observations related to ownership, lifecycle, and sharing controls.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Manage who can create Microsoft 365 Groups' -Url 'https://learn.microsoft.com/en-us/microsoft-365/solutions/manage-creation-of-groups' -WhyItIsRelevant 'Supports governance of Teams-connected groups and workspace sprawl.'),
                (New-CustomerDocumentationReference -Title 'Set expiration for Microsoft 365 groups' -Url 'https://learn.microsoft.com/en-us/entra/identity/users/groups-lifecycle' -WhyItIsRelevant 'Relevant to dormant collaboration spaces and lifecycle control.')
            )
        },
        [pscustomobject]@{
            Title      = '15.8 DNS DMARC and OneDrive'
            Intro      = 'These references support the domain-authentication and OneDrive lifecycle observations in this assessment.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Set up SPF in Microsoft 365 to help prevent spoofing' -Url 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/set-up-spf-in-office-365-to-help-prevent-spoofing' -WhyItIsRelevant 'Supports SPF and anti-spoofing guidance for the reviewed domains.'),
                (New-CustomerDocumentationReference -Title 'Use DKIM to validate outbound email sent from your custom domain' -Url 'https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dkim-configure' -WhyItIsRelevant 'Supports the observed mail-authentication posture for custom domains.'),
                (New-CustomerDocumentationReference -Title 'Retention and deletion in OneDrive and SharePoint' -Url 'https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion' -WhyItIsRelevant 'Relevant to stale OneDrive and SharePoint lifecycle handling.')
            )
        },
        [pscustomobject]@{
            Title      = '15.9 Retention Policies and Data Loss Prevention'
            Intro      = 'These references support the current-state observations around retention visibility, data lifecycle governance, and DLP maturity.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Learn about retention policies and retention labels' -Url 'https://learn.microsoft.com/en-us/purview/retention' -WhyItIsRelevant 'Supports the retention-policy observations and the need for documented lifecycle controls.'),
                (New-CustomerDocumentationReference -Title 'Learn about data loss prevention' -Url 'https://learn.microsoft.com/en-us/purview/dlp-learn-about-dlp' -WhyItIsRelevant 'Provides Microsoft guidance for DLP and data-protection governance.')
            )
        },
        [pscustomobject]@{
            Title      = '15.9 Pass-Through Authentication and Password Writeback'
            Intro      = 'These references support the password writeback, SSPR, and hybrid identity posture observations described in the report.'
            References = @(
                (New-CustomerDocumentationReference -Title 'User self-service password reset deep dive' -Url 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks' -WhyItIsRelevant 'Supports the password reset and self-service password reset observations.'),
                (New-CustomerDocumentationReference -Title 'How password writeback works in Microsoft Entra Connect' -Url 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-writeback' -WhyItIsRelevant 'Relevant to password writeback and hybrid credential-management considerations.')
            )
        }
    )
}

function Convert-ToCustomerWordCellParagraphs {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$Content)

    if ($null -eq $Content) {
        return @('Not validated from the reviewed data')
    }

    if (($Content -is [string]) -and $Content -eq '__ARRAYA_BLANK__') {
        return @(' ')
    }

    if (($Content -is [string]) -and $Content.Length -eq 0) {
        return @(' ')
    }

    if (($Content -is [System.Collections.IEnumerable]) -and -not ($Content -is [string])) {
        $items = @()
        foreach ($item in @($Content)) {
            $items += @(Convert-ToCustomerWordCellParagraphs -Content $item)
        }
        return @($items | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    $text = Convert-ToCustomerAssessmentDisplayText -Value $Content -Default 'Not validated from the reviewed data'
    $paragraphs = @(
        $text -split "(`r`n|`n|`r)" |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($paragraphs.Count -eq 0) {
        return @('Not validated from the reviewed data')
    }

    return $paragraphs
}

function Convert-ToCustomerWordXmlText {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    $normalized = Convert-ToArrayaDisplayText -Value $Text -Default ''
    return [System.Security.SecurityElement]::Escape($normalized)
}

function New-CustomerWordParagraphXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $false)][string]$Style = 'Normal',
        [Parameter(Mandatory = $false)][switch]$Bold
    )

    $styleXml = if ([string]::IsNullOrWhiteSpace($Style)) { '' } else { "<w:pPr><w:pStyle w:val=""$Style""/></w:pPr>" }
    $runProperties = if ($Bold) { '<w:rPr><w:b/></w:rPr>' } else { '' }
    $escapedText = Convert-ToCustomerWordXmlText -Text $Text
    return "<w:p>$styleXml<w:r>$runProperties<w:t xml:space=""preserve"">$escapedText</w:t></w:r></w:p>"
}

function New-CustomerWordTableCellXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Content,
        [Parameter(Mandatory = $false)][switch]$Header
    )

    $paragraphXml = @()
    foreach ($paragraph in (Convert-ToCustomerWordCellParagraphs -Content $Content)) {
        $paragraphXml += New-CustomerWordParagraphXml -Text $paragraph -Style 'Normal' -Bold:$Header
    }

    $cellProperties = if ($Header) {
        '<w:tcPr><w:tcW w:w="0" w:type="auto"/><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr>'
    }
    else {
        '<w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>'
    }

    return "<w:tc>$cellProperties$($paragraphXml -join '')</w:tc>"
}

function New-CustomerWordTableXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string[]]$Headers,
        [Parameter(Mandatory = $true)]$Rows,
        [Parameter(Mandatory = $false)][string]$Style
    )

    $headerCells = foreach ($header in $Headers) {
        New-CustomerWordTableCellXml -Content $header -Header
    }

    $columnCount = @($Headers).Count
    $tblGridColumns = @()
    if ($columnCount -gt 0) {
        $gridWidth = [math]::Floor(9000 / $columnCount)
        for ($index = 0; $index -lt $columnCount; $index++) {
            $tblGridColumns += "<w:gridCol w:w=""$gridWidth""/>"
        }
    }

    $styleXml = if ([string]::IsNullOrWhiteSpace($Style)) { '' } else { "<w:tblStyle w:val=""$Style""/>" }

    $rowXml = New-Object System.Collections.Generic.List[string]
    $rowXml.Add("<w:tr>$($headerCells -join '')</w:tr>") | Out-Null
    foreach ($row in @($Rows)) {
        $cells = @()
        if ($null -ne $row -and $row.PSObject.Properties['Cells']) {
            $cells = @($row.Cells)
        }
        elseif (($row -is [System.Collections.IEnumerable]) -and -not ($row -is [string]) -and -not ($row -is [System.Collections.IDictionary])) {
            $cells = @($row)
        }
        else {
            $cells = @($row)
        }

        if ($columnCount -gt 0 -and $cells.Count -lt $columnCount) {
            while ($cells.Count -lt $columnCount) {
                $cells += $null
            }
        }

        if ($columnCount -gt 0 -and $cells.Count -gt $columnCount) {
            $leadingCells = @()
            if ($columnCount -gt 1) {
                $leadingCells = @($cells[0..($columnCount - 2)])
            }
            $trailingCells = @($cells[($columnCount - 1)..($cells.Count - 1)])
            $cells = @($leadingCells + ,$trailingCells)
        }

        $cellXml = foreach ($cell in $cells) {
            New-CustomerWordTableCellXml -Content $cell
        }
        $rowXml.Add("<w:tr>$($cellXml -join '')</w:tr>") | Out-Null
    }

    return @"
<w:tbl>
  <w:tblPr>
    $styleXml
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="12" w:space="0" w:color="5C667A"/>
      <w:left w:val="single" w:sz="12" w:space="0" w:color="5C667A"/>
      <w:bottom w:val="single" w:sz="12" w:space="0" w:color="5C667A"/>
      <w:right w:val="single" w:sz="12" w:space="0" w:color="5C667A"/>
      <w:insideH w:val="single" w:sz="8" w:space="0" w:color="7E879A"/>
      <w:insideV w:val="single" w:sz="8" w:space="0" w:color="7E879A"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="80" w:type="dxa"/>
      <w:left w:w="90" w:type="dxa"/>
      <w:bottom w:w="80" w:type="dxa"/>
      <w:right w:w="90" w:type="dxa"/>
    </w:tblCellMar>
    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
  </w:tblPr>
  <w:tblGrid>
    $($tblGridColumns -join '')
  </w:tblGrid>
  $($rowXml -join [Environment]::NewLine)
</w:tbl>
"@
}

function Get-CustomerTemplateSectionPropertiesXml {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$TemplatePath)

    Add-Type -AssemblyName System.IO.Compression
    $templateStream = [System.IO.File]::Open($TemplatePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $archive = New-Object System.IO.Compression.ZipArchive($templateStream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
    try {
        $entry = $archive.GetEntry('word/document.xml')
        if ($null -eq $entry) {
            return '<w:sectPr/>'
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            [xml]$documentXml = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }

        $ns = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $sectPr = $documentXml.SelectSingleNode('//w:body/w:sectPr', $ns)
        if ($null -eq $sectPr) {
            return '<w:sectPr/>'
        }

        return $sectPr.OuterXml
    }
    finally {
        $archive.Dispose()
        $templateStream.Dispose()
    }
}

function Copy-CustomerAssessmentTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TemplatePath,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $inputStream = [System.IO.File]::Open($TemplatePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $outputStream = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $inputStream.CopyTo($outputStream)
        }
        finally {
            $outputStream.Dispose()
        }
    }
    finally {
        $inputStream.Dispose()
    }
}

function Set-CustomerZipTextEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Archive,
        [Parameter(Mandatory = $true)][string]$EntryName,
        [Parameter(Mandatory = $true)][string]$Content
    )

    $existingEntry = $Archive.GetEntry($EntryName)
    if ($null -ne $existingEntry) {
        $existingEntry.Delete()
    }

    $entry = $Archive.CreateEntry($EntryName)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), [System.Text.UTF8Encoding]::new($false))
    try {
        $writer.Write($Content)
    }
    finally {
        $writer.Dispose()
    }
}

function Set-CustomerAssessmentCoreProperties {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Archive,
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt
    )

    $timestamp = $GeneratedAt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $title = Convert-ToCustomerWordXmlText -Text "$TenantName Microsoft 365 Tenant Best Practices Assessment"
    $creator = Convert-ToCustomerWordXmlText -Text 'Arraya Solutions'
    $coreXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>$title</dc:title>
  <dc:creator>$creator</dc:creator>
  <cp:lastModifiedBy>$creator</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$timestamp</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$timestamp</dcterms:modified>
</cp:coreProperties>
"@

    Set-CustomerZipTextEntry -Archive $Archive -EntryName 'docProps/core.xml' -Content $coreXml
}

function New-CustomerAssessmentDocumentXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Blocks,
        [Parameter(Mandatory = $true)][string]$TemplatePath
    )

    $bodyElements = New-Object System.Collections.Generic.List[string]
    foreach ($block in @($Blocks)) {
        if ($null -eq $block) { continue }
        switch ([string]$block.Type) {
            'Paragraph' {
                $bodyElements.Add((New-CustomerWordParagraphXml -Text ([string]$block.Text) -Style ([string]$block.Style))) | Out-Null
            }
            'Table' {
                $bodyElements.Add((New-CustomerWordTableXml -Headers @($block.Headers) -Rows @($block.Rows) -Style ([string]$block.Style))) | Out-Null
            }
        }
    }

    $sectPrXml = Get-CustomerTemplateSectionPropertiesXml -TemplatePath $TemplatePath
    $bodyElements.Add($sectPrXml) | Out-Null

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    $($bodyElements -join [Environment]::NewLine)
  </w:body>
</w:document>
"@
}

function Write-CustomerAssessmentDocxFromModel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TemplatePath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt,
        [Parameter(Mandatory = $true)][object[]]$Blocks
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath -Force

    $documentXml = New-CustomerAssessmentDocumentXml -Blocks $Blocks -TemplatePath $OutputPath
    $archive = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        Set-CustomerZipTextEntry -Archive $Archive -EntryName 'word/document.xml' -Content $documentXml
        Set-CustomerAssessmentCoreProperties -Archive $Archive -TenantName $TenantName -GeneratedAt $GeneratedAt
    }
    finally {
        $archive.Dispose()
    }
}

function Get-CustomerAssessmentDocumentXml {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry('word/document.xml')
        if ($null -eq $entry) {
            throw "Customer assessment DOCX did not contain word/document.xml: $Path"
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            return [xml]$reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Get-CustomerAssessmentXmlNamespaceManager {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][xml]$Document)

    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
    $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    return (, $namespaceManager)
}

function Get-CustomerAssessmentParagraphStyle {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlNode]$ParagraphNode,
        [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $styleNode = $ParagraphNode.SelectSingleNode('./w:pPr/w:pStyle', $NamespaceManager)
    if ($null -eq $styleNode -or $null -eq $styleNode.Attributes) {
        return $null
    }

    $styleAttribute = @($styleNode.Attributes | Where-Object { $_.LocalName -eq 'val' } | Select-Object -First 1)
    if ($styleAttribute.Count -eq 0) {
        return $null
    }

    return [string]$styleAttribute[0].Value
}

function Get-CustomerAssessmentParagraphText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlNode]$ParagraphNode,
        [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($child in @($ParagraphNode.SelectNodes('.//w:r/*', $NamespaceManager))) {
        switch ($child.LocalName) {
            't' {
                $segments.Add([string]$child.InnerText) | Out-Null
            }
            'tab' {
                $segments.Add("`t") | Out-Null
            }
            'br' {
                $segments.Add([Environment]::NewLine) | Out-Null
            }
        }
    }

    return (($segments.ToArray() -join '') -replace '\s+$', '').Trim()
}

function Convert-CustomerAssessmentParagraphNodeToMarkdown {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlNode]$ParagraphNode,
        [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $text = Get-CustomerAssessmentParagraphText -ParagraphNode $ParagraphNode -NamespaceManager $NamespaceManager
    if ([string]::IsNullOrWhiteSpace($text)) {
        return @()
    }

    $style = Get-CustomerAssessmentParagraphStyle -ParagraphNode $ParagraphNode -NamespaceManager $NamespaceManager
    switch ([string]$style) {
        'Title' { return @("# $text", '') }
        'Heading1' { return @("## $text", '') }
        'Heading2' { return @("### $text", '') }
        'Heading3' { return @("#### $text", '') }
        'Heading4' { return @("##### $text", '') }
        'Subtitle' { return @("**$text**", '') }
        default { return @($text, '') }
    }
}

function Convert-CustomerAssessmentTableNodeToMarkdown {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlNode]$TableNode,
        [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $rows = @($TableNode.SelectNodes('./w:tr', $NamespaceManager))
    if ($rows.Count -eq 0) {
        return @()
    }

    $renderedRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($rows)) {
        $cells = New-Object System.Collections.Generic.List[string]
        foreach ($cell in @($row.SelectNodes('./w:tc', $NamespaceManager))) {
            $cellParagraphs = New-Object System.Collections.Generic.List[string]
            foreach ($paragraph in @($cell.SelectNodes('./w:p', $NamespaceManager))) {
                $paragraphText = Get-CustomerAssessmentParagraphText -ParagraphNode $paragraph -NamespaceManager $NamespaceManager
                if (-not [string]::IsNullOrWhiteSpace($paragraphText)) {
                    $cellParagraphs.Add(((Convert-ToArrayaMarkdownText $paragraphText) -replace '\r?\n', '<br/>')) | Out-Null
                }
            }

            if ($cellParagraphs.Count -eq 0) {
                $cells.Add('') | Out-Null
            }
            else {
                $cells.Add(($cellParagraphs.ToArray() -join '<br/>')) | Out-Null
            }
        }
        $renderedRows.Add(@($cells.ToArray())) | Out-Null
    }

    $headerRow = @($renderedRows[0])
    $columnCount = $headerRow.Count
    if ($columnCount -eq 0) {
        return @()
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('| ' + ($headerRow -join ' | ') + ' |') | Out-Null
    $lines.Add('| ' + (@(1..$columnCount | ForEach-Object { '---' }) -join ' | ') + ' |') | Out-Null

    if ($renderedRows.Count -gt 1) {
        for ($rowIndex = 1; $rowIndex -lt $renderedRows.Count; $rowIndex++) {
            $rowCells = @($renderedRows[$rowIndex])
            while ($rowCells.Count -lt $columnCount) {
                $rowCells += ''
            }
            $lines.Add('| ' + (@($rowCells | Select-Object -First $columnCount) -join ' | ') + ' |') | Out-Null
        }
    }
    $lines.Add('') | Out-Null

    return @($lines.ToArray())
}

function Write-CustomerAssessmentMarkdownFromDocx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $document = Get-CustomerAssessmentDocumentXml -Path $InputPath
    $namespaceManager = Get-CustomerAssessmentXmlNamespaceManager -Document $document
    $lines = New-Object System.Collections.Generic.List[string]

    foreach ($node in @($document.DocumentElement.SelectNodes('./w:body/*', $namespaceManager))) {
        switch ($node.LocalName) {
            'p' {
                foreach ($line in @(Convert-CustomerAssessmentParagraphNodeToMarkdown -ParagraphNode $node -NamespaceManager $namespaceManager)) {
                    $lines.Add($line) | Out-Null
                }
            }
            'tbl' {
                foreach ($line in @(Convert-CustomerAssessmentTableNodeToMarkdown -TableNode $node -NamespaceManager $namespaceManager)) {
                    $lines.Add($line) | Out-Null
                }
            }
        }
    }

    $markdown = ($lines.ToArray() -join [Environment]::NewLine).Trim() + [Environment]::NewLine
    [System.IO.File]::WriteAllText($OutputPath, $markdown, [System.Text.UTF8Encoding]::new($false))
}

function New-CustomerAssessmentDocumentBlocks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$SourceModel,
        [Parameter(Mandatory = $true)]$Signals,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt
    )

    $tenantName = [string]$SourceModel.TenantName
    $assessmentVersion = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $SourceModel -Names @('AssessmentVersion')) -Default '3.x'
    $findings = @($SourceModel.Findings)
    $roadmapActions = @(Get-CustomerDistinctRoadmapActions -RoadmapActions $SourceModel.RoadmapActions)
    $workstreamSummaries = @($SourceModel.WorkstreamSummaries)
    $executiveThemes = @($SourceModel.ExecutiveThemes)
    $sourceSummaryRows = @($SourceModel.SummaryRows)
    $findingsLegendRows = @($SourceModel.FindingsLegendRows)
    $executiveDecisionSummary = $SourceModel.ExecutiveDecisionSummary
    $guestMfaExperienceSummary = $SourceModel.GuestMfaExperienceSummary
    $identityConsultativeSummary = $SourceModel.IdentityConsultativeSummary
    $messagingConsultativeSummary = $SourceModel.MessagingConsultativeSummary
    $collaborationConsultativeSummary = $SourceModel.CollaborationConsultativeSummary
    $governanceConsultativeSummary = $SourceModel.GovernanceConsultativeSummary
    $lifecycleConsultativeSummary = $SourceModel.LifecycleConsultativeSummary
    $leadershipDecisionRows = if ($null -ne $executiveDecisionSummary -and @($executiveDecisionSummary.DecisionRows).Count -gt 0) {
        Convert-CustomerThreeColumnRowsToWordTableRows -Rows @($executiveDecisionSummary.DecisionRows) -PropertyNames @('DecisionFocus', 'WhatShouldHappenNext', 'WhyNow')
    }
    else {
        Get-CustomerLeadershipDecisionRows -RoadmapActions $roadmapActions
    }
    $executiveRiskRows = if ($null -ne $executiveDecisionSummary -and @($executiveDecisionSummary.RiskRows).Count -gt 0) {
        Convert-CustomerThreeColumnRowsToWordTableRows -Rows @($executiveDecisionSummary.RiskRows) -PropertyNames @('RiskCluster', 'WhatStandsOut', 'WhyLeadershipShouldCare')
    }
    else {
        @()
    }
    $identityObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Identity & Access (Entra ID)'
    $deviceObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Devices & Endpoint Management'
    $messagingObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Messaging (Exchange Online)'
    $collaborationObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Collaboration (Teams, SharePoint, OneDrive)'
    $governanceObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Data Protection & Governance'
    $lifecycleObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Offboarding & Lifecycle Management'

    $userRows = Convert-ArrayaObjectToArray $Signals.Users
    $deviceRows = Convert-ArrayaObjectToArray $Signals.DeviceDetails
    $domainRows = Convert-ArrayaObjectToArray $Signals.Domains
    $enterpriseApplications = Convert-ArrayaObjectToArray $Signals.EnterpriseApplications
    $adminRows = Convert-ArrayaObjectToArray $Signals.Admins
    $recipientRows = Convert-ArrayaObjectToArray $Signals.AllRecipients
    $mailboxRows = Convert-ArrayaObjectToArray $Signals.AllMailboxes
    $primaryMailboxStatsRows = Convert-ArrayaObjectToArray $Signals.PrimaryMailboxStats
    $archiveMailboxRows = Convert-ArrayaObjectToArray $Signals.ArchiveMailboxes
    $archiveMailboxStatsRows = Convert-ArrayaObjectToArray $Signals.ArchiveMailboxStats
    $inactiveMailboxRows = Convert-ArrayaObjectToArray $Signals.InactiveMailboxes
    $nonUserMailboxRows = Convert-ArrayaObjectToArray $Signals.NonUserMailboxes
    $emailActivitySummaryRecord = if ($Signals.EmailActivitySummary) { Get-ArrayaObjectValue -Object $Signals.EmailActivitySummary -Names @('Summary') } else { $null }
    $emailActivityTopSenders = Convert-ArrayaObjectToArray $Signals.EmailActivityTopSenders
    $smtpRelayConfigRecord = if ($Signals.SMTPRelayConfig) { Get-ArrayaObjectValue -Object $Signals.SMTPRelayConfig -Names @('Configuration', 'Summary') } else { $null }
    $smtpRelayServiceAccounts = Convert-ArrayaObjectToArray $Signals.SMTPRelayServiceAccounts
    $teamRows = Convert-ArrayaObjectToArray $Signals.AllTeams
    $sharePointRows = Convert-ArrayaObjectToArray $Signals.SharePoint
    $oneDriveRows = Convert-ArrayaObjectToArray $Signals.OneDrive
    $unifiedGroupRows = Convert-ArrayaObjectToArray $Signals.UnifiedGroups
    $teamsVoiceSummaryRecord = if ($Signals.TeamsVoiceSummary) { Get-ArrayaObjectValue -Object $Signals.TeamsVoiceSummary -Names @('Summary') } else { $null }
    $adConnectSummaryRecord = if ($Signals.AdConnectConfiguration) { Get-ArrayaObjectValue -Object $Signals.AdConnectConfiguration -Names @('Summary') } else { $null }
    $spamFilteringSummaryRecord = if ($Signals.SpamFilteringSummary) { Get-ArrayaObjectValue -Object $Signals.SpamFilteringSummary -Names @('Summary') } else { $null }
    $sharePointSharingSummaryRecord = if ($Signals.SharePointSharingSummary) { Get-ArrayaObjectValue -Object $Signals.SharePointSharingSummary -Names @('Summary') } else { $null }
    $externalSharingSummaryRecord = if ($Signals.ExternalSharingSummary) { Get-ArrayaObjectValue -Object $Signals.ExternalSharingSummary -Names @('Summary') } else { $null }
    $externalIdentityRestrictionsRecord = if ($Signals.ExternalIdentityRestrictions) { Get-ArrayaObjectValue -Object $Signals.ExternalIdentityRestrictions -Names @('Summary') } else { $null }
    $guestAccessConfigurationRecord = if ($Signals.GuestAccessConfiguration) { Get-ArrayaObjectValue -Object $Signals.GuestAccessConfiguration -Names @('Summary') } else { $null }
    $authenticationConfigRecord = if ($Signals.AuthenticationConfig) { Get-ArrayaObjectValue -Object $Signals.AuthenticationConfig -Names @('Configuration') } else { $null }
    $mfaEnrollmentSummaryRecord = if ($Signals.MfaEnrollmentSummary) { $Signals.MfaEnrollmentSummary } elseif ($Signals.MfaRegistrationSummary) { $Signals.MfaRegistrationSummary } else { $null }
    $mfaEnforcementSummaryRecord = if ($Signals.MfaEnforcementSummary) { $Signals.MfaEnforcementSummary } else { $null }
    $mfaRegistrationDetailRows = @(
        Convert-ToCustomerAssessmentCollectionRows `
            -Value $Signals.MfaRegistrationDetails `
            -MarkerNames @('UserPrincipalName', 'IsMfaRegistered')
    )
    $mfaEnforcementGapUserRows = @(
        Convert-ToCustomerAssessmentCollectionRows `
            -Value $Signals.MfaEnforcementGapUsers `
            -MarkerNames @('DisplayName', 'UserPrincipalName', 'GapCategory')
    )
    $mfaEnforcementScopeReviewRows = @(
        Convert-ToCustomerAssessmentCollectionRows `
            -Value $Signals.MfaEnforcementScopeReview `
            -MarkerNames @('PolicyName', 'ScopeType', 'DisplayName')
    )
    $adminMfaSummaryRecord = if ($Signals.AdminMfaSummary) { $Signals.AdminMfaSummary } else { $null }
    $adminMfaRegistrationGapRows = @(
        Convert-ToCustomerAssessmentCollectionRows `
            -Value $Signals.AdminMfaRegistrationGaps `
            -MarkerNames @('DisplayName', 'UserPrincipalName', 'MfaRegistrationState')
    )
    $adminMfaEnforcementGapRows = @(
        Convert-ToCustomerAssessmentCollectionRows `
            -Value $Signals.AdminMfaEnforcementGaps `
            -MarkerNames @('DisplayName', 'UserPrincipalName', 'MfaEnforcementState')
    )
    $securityDefaultsPolicyRecord = if ($Signals.SecurityDefaultsPolicy) { Get-ArrayaObjectValue -Object $Signals.SecurityDefaultsPolicy -Names @('Configuration', 'Summary') } else { $null }
    $passwordLifecycleSummaryRecord = $Signals.PasswordLifecycleSummary
    $externalSharingSiteOverrides = Convert-ArrayaObjectToArray $Signals.ExternalSharingSiteOverrides
    $externalExposureFindings = Convert-ArrayaObjectToArray $Signals.ExternalExposureFindings
    $retentionPolicyRows = Convert-ArrayaObjectToArray $Signals.RetentionPolicies
    $dlpPolicyRows = Convert-ArrayaObjectToArray $Signals.DlpPolicies

    $guestUsers = @($userRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('UserType'))).ToLowerInvariant() -eq 'guest' })
    $memberUsers = @($userRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('UserType'))).ToLowerInvariant() -ne 'guest' })
    $enabledUsers = @($userRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))) -ne $false })
    $enabledGuestUsers = @($guestUsers | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))) -ne $false })
    $enabledMemberUsers = @($memberUsers | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))) -ne $false })
    $enabledAdminRows = @($adminRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))) -ne $false })
    $inactiveMemberUsers = @(
        $enabledMemberUsers | Where-Object {
            $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime', 'SignInActivityLastSignInDateTime'))
            $null -eq $lastSignIn -or $lastSignIn -lt (Get-Date).AddDays(-180)
        }
    )
    $inactiveGuestUsers = @(
        $guestUsers | Where-Object {
            $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime', 'SignInActivityLastSignInDateTime'))
            $enabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))
            ($enabled -ne $false) -and ($null -eq $lastSignIn -or $lastSignIn -lt (Get-Date).AddDays(-90))
        }
    )

    $platformGroups = @(
        $deviceRows |
            Group-Object {
                $raw = Get-ArrayaObjectValue -Object $_ -Names @('OperatingSystem', 'OS', 'DeviceOSType')
                Convert-ToArrayaDisplayText -Value $raw -Default 'Not surfaced in current source'
            } |
            Sort-Object @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name } }
    )
    $dominantPlatform = if ($platformGroups.Count -gt 0) {
        $top = $platformGroups[0]
        $pct = if ($deviceRows.Count -gt 0) { [math]::Round(($top.Count / $deviceRows.Count) * 100, 2) } else { $null }
        if ($null -ne $pct) { '{0} ({1}% of reviewed devices)' -f $top.Name, $pct } else { $top.Name }
    }
    else {
        'Not surfaced in current source'
    }

    $applicationSummaryRecord = if ($Signals.EnterpriseApplicationSummary) { Get-ArrayaObjectValue -Object $Signals.EnterpriseApplicationSummary -Names @('Summary') } else { $null }
    $enterpriseApplicationRowSummaryRecord = if ($enterpriseApplications.Count -gt 0) { Get-CustomerEnterpriseApplicationSummaryFromRows -ApplicationRows $enterpriseApplications } else { $null }
    if ($null -eq $applicationSummaryRecord -and $null -ne $enterpriseApplicationRowSummaryRecord) {
        $applicationSummaryRecord = $enterpriseApplicationRowSummaryRecord
    }
    elseif ($null -ne $applicationSummaryRecord -and $null -ne $enterpriseApplicationRowSummaryRecord) {
        $mergedApplicationSummaryRecord = [ordered]@{}
        foreach ($property in @($applicationSummaryRecord.PSObject.Properties)) {
            $mergedApplicationSummaryRecord[$property.Name] = $property.Value
        }
        foreach ($property in @($enterpriseApplicationRowSummaryRecord.PSObject.Properties)) {
            $hasCurrentValue = $mergedApplicationSummaryRecord.Contains($property.Name)
            $currentValue = if ($hasCurrentValue) { $mergedApplicationSummaryRecord[$property.Name] } else { $null }
            if (-not $hasCurrentValue -or $null -eq $currentValue -or [string]::IsNullOrWhiteSpace([string]$currentValue)) {
                $mergedApplicationSummaryRecord[$property.Name] = $property.Value
            }
        }
        $applicationSummaryRecord = [pscustomobject]$mergedApplicationSummaryRecord
    }

    $totalEnterpriseApps = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('TotalEnterpriseApplications', 'EnterpriseApplicationCount'))
    $highPrivilegeApps = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('ApplicationsWithHighPrivilege', 'HighPrivilegeApplicationCount'))
    $ssoEnabledApps = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('SsoEnabledApplications', 'SsoEnabledApplicationCount'))
    $firstPartyAppsWithoutOwners = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('FirstPartyAppsWithoutOwners'))
    $appsWithRedirectUriRisk = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('ApplicationsWithInsecureRedirectUris'))
    $appsWithNoRecentActivity = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('ApplicationsWithNoRecentActivity'))
    $thirdPartyAppsWithApplicationPerms = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('ThirdPartyAppsWithApplicationPerms'))
    $ownerSignalCoverageState = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('OwnerSignalCoverageState')) -Default $(if ($null -ne $enterpriseApplicationRowSummaryRecord) { [string]$enterpriseApplicationRowSummaryRecord.OwnerSignalCoverageState } else { 'Unavailable' })
    $redirectUriSignalCoverageState = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('RedirectUriSignalCoverageState')) -Default $(if ($null -ne $enterpriseApplicationRowSummaryRecord) { [string]$enterpriseApplicationRowSummaryRecord.RedirectUriSignalCoverageState } else { 'Unavailable' })
    $activitySignalCoverageState = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('ActivitySignalCoverageState')) -Default $(if ($null -ne $enterpriseApplicationRowSummaryRecord) { [string]$enterpriseApplicationRowSummaryRecord.ActivitySignalCoverageState } else { 'Unavailable' })
    if ($null -eq $totalEnterpriseApps -or (($totalEnterpriseApps -eq 0) -and ($enterpriseApplications.Count -gt 0))) { $totalEnterpriseApps = $enterpriseApplications.Count }
    if ($null -eq $ssoEnabledApps) {
        $ssoEnabledApps = @(
            $enterpriseApplications |
                Where-Object { Test-CustomerEnterpriseApplicationSsoEnabled -ApplicationRow $_ }
        ).Count
    }
    if ($null -eq $firstPartyAppsWithoutOwners) {
        $firstPartyAppsWithoutOwners = @(
            $enterpriseApplications |
                Where-Object {
                    ([string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'First Party') -and
                    ((Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'Collected') -and
                    ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))) -le 0)
                }
        ).Count
    }
    if ($null -eq $appsWithRedirectUriRisk) {
        $appsWithRedirectUriRisk = @(
            $enterpriseApplications |
                Where-Object {
                    $redirectSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState'
                    (($redirectSignalState -eq 'Collected') -or ($redirectSignalState -eq 'Partial')) -and
                    ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasInsecureRedirectUris'))) -eq $true)
                }
        ).Count
    }
    if ($null -eq $appsWithNoRecentActivity) {
        $appsWithNoRecentActivity = @(
            $enterpriseApplications |
                Where-Object {
                    ((Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Collected') -and
                    ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasRecentActivity'))) -eq $false)
                }
        ).Count
    }
    if ($null -eq $thirdPartyAppsWithApplicationPerms) {
        $thirdPartyAppsWithApplicationPerms = @(
            $enterpriseApplications |
                Where-Object {
                    ([string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'Third Party') -and
                    ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount'))) -gt 0)
                }
        ).Count
    }
    $ownerSignalFullyValidated = ($ownerSignalCoverageState -eq 'Collected') -or ($ownerSignalCoverageState -eq 'NotApplicable')
    $redirectUriSignalFullyValidated = ($redirectUriSignalCoverageState -eq 'Collected')
    $activitySignalFullyValidated = ($activitySignalCoverageState -eq 'Collected')
    $applicationValidationNotes = New-Object System.Collections.Generic.List[string]
    if (-not $ownerSignalFullyValidated) {
        $applicationValidationNotes.Add('Owner validation was not fully validated in this run.') | Out-Null
    }
    if (-not $redirectUriSignalFullyValidated) {
        $applicationValidationNotes.Add('Redirect URI review was not fully validated in this run.') | Out-Null
    }
    if (-not $activitySignalFullyValidated) {
        $applicationValidationNotes.Add('Recent activity validation was not fully validated in this run.') | Out-Null
    }
    $applicationValidationNoteText = @($applicationValidationNotes.ToArray()) -join ' '
    $applicationInventoryValidated = ($null -ne $applicationSummaryRecord) -or ($enterpriseApplications.Count -gt 0)
    $applicationInventoryClearlyEmpty = $applicationInventoryValidated -and ($enterpriseApplications.Count -eq 0) -and ($totalEnterpriseApps -eq 0)

    $passwordWritebackState = Convert-ToCustomerAssessmentDisplayText `
        -Value (Get-ArrayaObjectValue -Object $passwordLifecycleSummaryRecord -Names @('PasswordWriteback', 'PasswordWritebackEnabled')) `
        -Default (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $adConnectSummaryRecord -Names @('PasswordWriteback', 'PasswordWritebackEnabled')) -Default 'Not validated from the reviewed data')
    $selfServicePasswordResetState = Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $passwordLifecycleSummaryRecord -Names @('SelfServicePasswordReset', 'SelfServicePasswordResetEnabled')) -TrueText 'Enabled' -FalseText 'Not enabled' -Default 'Not validated from the reviewed data'
    $passThroughAuthenticationState = Convert-ToCustomerAssessmentDisplayText `
        -Value (Get-ArrayaObjectValue -Object $passwordLifecycleSummaryRecord -Names @('PassThroughAuthentication', 'PassThroughAuthenticationEnabled')) `
        -Default (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $adConnectSummaryRecord -Names @('PassThroughAuthentication', 'PassThroughAuthenticationEnabled')) -Default 'Not validated from the reviewed data')
    $directorySynchronizationState = Convert-ToCustomerAssessmentBooleanLabel `
        -Value (Get-ArrayaObjectValue -Object $passwordLifecycleSummaryRecord -Names @('OnPremisesSyncEnabled')) `
        -TrueText 'Enabled' `
        -FalseText 'Not enabled' `
        -Default (Get-CustomerObservationState -Observation $governanceObservation -Signal 'Directory synchronization')
    $onPremisesLastSyncState = Convert-ToCustomerAssessmentDisplayText `
        -Value (Get-ArrayaObjectValue -Object $passwordLifecycleSummaryRecord -Names @('OnPremisesLastSyncDateTime')) `
        -Default (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $adConnectSummaryRecord -Names @('OnPremisesLastSyncDateTime')) -Default 'Not validated from the reviewed data')
    $authenticationVisibilityState = Get-CustomerObservationState -Observation $governanceObservation -Signal 'Admin consent workflow'

    $gaRows = @(
        $adminRows | Where-Object {
            [string](Get-ArrayaObjectValue -Object $_ -Names @('Role', 'DirectoryRole', 'AdminRole', 'RoleName')) -match 'Global Administrator'
        }
    )
    $verifiedDomains = @(
        $domainRows | Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsVerified', 'Verified'))) -eq $true
        }
    )

    $formatSizeGb = {
        param($Bytes)

        $sizeBytes = Convert-ArrayaToNumber $Bytes
        if ($null -eq $sizeBytes) {
            return 'Not surfaced in current source'
        }

        return ('{0:N2}' -f [math]::Round(($sizeBytes / 1GB), 2))
    }
    $formatSizeGbOrZero = {
        param($Bytes)

        $sizeBytes = Convert-ArrayaToNumber $Bytes
        if ($null -eq $sizeBytes) {
            return '0.00'
        }

        return ('{0:N2}' -f [math]::Round(($sizeBytes / 1GB), 2))
    }

    $mailboxGuidLookup = @{}
    foreach ($mailboxRow in $mailboxRows) {
        foreach ($candidateKey in @(
            (Get-ArrayaObjectValue -Object $mailboxRow -Names @('ExchangeGuid')),
            (Get-ArrayaObjectValue -Object $mailboxRow -Names @('Guid')),
            (Get-ArrayaObjectValue -Object $mailboxRow -Names @('ExternalDirectoryObjectId'))
        )) {
            $lookupKey = Convert-ToArrayaDisplayText -Value $candidateKey -Default ''
            if (-not [string]::IsNullOrWhiteSpace($lookupKey) -and -not $mailboxGuidLookup.ContainsKey($lookupKey)) {
                $mailboxGuidLookup[$lookupKey] = $mailboxRow
            }
        }
    }

    $mailboxTypeStatistics = @{}
    foreach ($primaryMailboxStatsRow in $primaryMailboxStatsRows) {
        $mailboxType = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $primaryMailboxStatsRow -Names @('MailboxType', 'RecipientTypeDetails')) -Default 'Unknown'
        if (-not $mailboxTypeStatistics.ContainsKey($mailboxType)) {
            $mailboxTypeStatistics[$mailboxType] = [ordered]@{
                MailboxType       = $mailboxType
                Quantity          = 0
                TotalSizeBytes    = 0
                TotalArchiveBytes = 0
            }
        }

        $mailboxTypeStatistics[$mailboxType].Quantity++
        $mailboxTypeStatistics[$mailboxType].TotalSizeBytes += [double](Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $primaryMailboxStatsRow -Names @('TotalItemSizeBytes')) | ForEach-Object { if ($null -eq $_) { 0 } else { $_ } })
    }
    foreach ($archiveMailboxStatsRow in $archiveMailboxStatsRows) {
        $archiveMailboxGuid = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $archiveMailboxStatsRow -Names @('MailboxGuid')) -Default ''
        $archiveSourceMailbox = if (-not [string]::IsNullOrWhiteSpace($archiveMailboxGuid) -and $mailboxGuidLookup.ContainsKey($archiveMailboxGuid)) { $mailboxGuidLookup[$archiveMailboxGuid] } else { $null }
        $archiveMailboxType = if ($archiveSourceMailbox) {
            Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $archiveSourceMailbox -Names @('RecipientTypeDetails')) -Default 'Unknown'
        }
        else {
            'Unknown'
        }

        if (-not $mailboxTypeStatistics.ContainsKey($archiveMailboxType)) {
            $mailboxTypeStatistics[$archiveMailboxType] = [ordered]@{
                MailboxType       = $archiveMailboxType
                Quantity          = 0
                TotalSizeBytes    = 0
                TotalArchiveBytes = 0
            }
        }

        $mailboxTypeStatistics[$archiveMailboxType].TotalArchiveBytes += [double](Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $archiveMailboxStatsRow -Names @('TotalItemSizeBytes')) | ForEach-Object { if ($null -eq $_) { 0 } else { $_ } })
    }

    $mailboxStatisticsRows = if ($mailboxTypeStatistics.Count -gt 0) {
        @(
            foreach ($mailboxTypeStatistic in @($mailboxTypeStatistics.Values | Sort-Object @{ Expression = { $_.Quantity }; Descending = $true }, MailboxType)) {
                $averagePrimaryBytes = if ($mailboxTypeStatistic.Quantity -gt 0) { $mailboxTypeStatistic.TotalSizeBytes / $mailboxTypeStatistic.Quantity } else { $null }
                $averageArchiveBytes = if ($mailboxTypeStatistic.Quantity -gt 0) { $mailboxTypeStatistic.TotalArchiveBytes / $mailboxTypeStatistic.Quantity } else { $null }
                New-CustomerWordTableRow -Cells @(
                    $mailboxTypeStatistic.MailboxType,
                    $mailboxTypeStatistic.Quantity,
                    (& $formatSizeGbOrZero $mailboxTypeStatistic.TotalSizeBytes),
                    (& $formatSizeGbOrZero $mailboxTypeStatistic.TotalArchiveBytes),
                    (& $formatSizeGbOrZero $averagePrimaryBytes),
                    (& $formatSizeGbOrZero $averageArchiveBytes)
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }

    $userMailboxStatsRows = @($primaryMailboxStatsRows | Where-Object { (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('MailboxType')) -Default '') -eq 'UserMailbox' })
    $sharedMailboxStatsRows = @($primaryMailboxStatsRows | Where-Object { (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('MailboxType')) -Default '') -eq 'SharedMailbox' })
    $groupMailboxStatsRows = @($primaryMailboxStatsRows | Where-Object { (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('MailboxType')) -Default '') -eq 'GroupMailbox' })
    $largestUserMailbox = @($userMailboxStatsRows | Sort-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes')) } -Descending | Select-Object -First 1)
    $largestSharedMailbox = @($sharedMailboxStatsRows | Sort-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes')) } -Descending | Select-Object -First 1)
    $largestGroupMailbox = @($groupMailboxStatsRows | Sort-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes')) } -Descending | Select-Object -First 1)
    $userMailboxesUnderOneGb = @($userMailboxStatsRows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes'))) -lt 1GB }).Count
    $sharedMailboxesUnderOneGb = @($sharedMailboxStatsRows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes'))) -lt 1GB }).Count
    $archiveMailboxesOverFiftyGb = @($archiveMailboxStatsRows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes'))) -ge 50GB }).Count
    $archiveMailboxesWithHolds = @(
        $archiveMailboxRows | Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('LitigationHoldEnabled'))) -eq $true -or
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('RetentionHoldEnabled'))) -eq $true -or
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DelayHoldApplied'))) -eq $true
        }
    ).Count

    $smtpAuthMailboxRows = @(
        $mailboxRows | Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('SmtpClientAuthenticationDisabled'))) -eq $false
        }
    )
    $smtpAuthMailboxLookup = @{}
    foreach ($smtpAuthMailbox in $smtpAuthMailboxRows) {
        foreach ($smtpKeyCandidate in @(
            (Get-ArrayaObjectValue -Object $smtpAuthMailbox -Names @('UserPrincipalName')),
            (Get-ArrayaObjectValue -Object $smtpAuthMailbox -Names @('PrimarySmtpAddress'))
        )) {
            $smtpLookupKey = Convert-ToArrayaDisplayText -Value $smtpKeyCandidate -Default ''
            if (-not [string]::IsNullOrWhiteSpace($smtpLookupKey)) {
                $smtpAuthMailboxLookup[$smtpLookupKey.ToLowerInvariant()] = $smtpAuthMailbox
            }
        }
    }

    $relayConnectorRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $smtpRelayConfigRecord -Names @('RelayConnectors'))
    $relayServiceAccountRows = if ($smtpRelayServiceAccounts.Count -gt 0) {
        @($smtpRelayServiceAccounts)
    }
    elseif ($smtpAuthMailboxRows.Count -gt 0) {
        @(
            foreach ($topSender in $emailActivityTopSenders) {
                $senderKey = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $topSender -Names @('UserPrincipalName')) -Default ''
                if (-not [string]::IsNullOrWhiteSpace($senderKey) -and $smtpAuthMailboxLookup.ContainsKey($senderKey.ToLowerInvariant())) {
                    [pscustomobject]@{
                        UserPrincipalName = $senderKey
                        DisplayName       = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $topSender -Names @('DisplayName')) -Default (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $smtpAuthMailboxLookup[$senderKey.ToLowerInvariant()] -Names @('DisplayName')) -Default $senderKey)
                        SendCount         = Get-ArrayaObjectValue -Object $topSender -Names @('SendCount')
                        LastActivityDate  = Get-ArrayaObjectValue -Object $topSender -Names @('LastActivityDate')
                    }
                }
            }
        )
    }
    else {
        @()
    }

    $dormantTeams = @(
        $teamRows | Where-Object {
            $lastActivity = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastActivityDate'))
            $lastActivity -and $lastActivity -lt (Get-Date).AddDays(-90)
        }
    ).Count
    $archivedTeams = @($teamRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsArchived'))) -eq $true }).Count
    $largestSharePointSites = @($sharePointRows | Sort-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('StorageUsedGB', 'StorageUsageCurrent')) } -Descending | Select-Object -First 5)
    $sharePointSitesExternalSharingEnabled = @(
        $sharePointRows | Where-Object {
            $sharingCapability = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('SharingCapability')) -Default ''
            -not [string]::IsNullOrWhiteSpace($sharingCapability) -and $sharingCapability -notmatch 'Disabled|Internal'
        }
    ).Count
    $teamConnectedSites = @($sharePointRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsTeamsConnected'))) -eq $true }).Count
    $guestInvitationControlText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $guestAccessConfigurationRecord -Names @('GuestInvitationControl')) -Default ''
    if ([string]::IsNullOrWhiteSpace($guestInvitationControlText) -or $guestInvitationControlText -eq 'Not validated from the reviewed data') {
        $guestInvitationControlText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $authenticationConfigRecord -Names @('AllowInvitesFrom')) -Default ''
    }
    if ([string]::IsNullOrWhiteSpace($guestInvitationControlText) -or $guestInvitationControlText -eq 'Not validated from the reviewed data') {
        $guestInvitationControlText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('AllowInvitesFrom')) -Default 'Not validated from the reviewed data'
    }
    $guestInvitationControlText = Convert-ToCustomerInvitationControlLabel -Value $guestInvitationControlText
    $tenantSharingCapabilityText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('TenantSharingCapability')) -Default ''
    if ([string]::IsNullOrWhiteSpace($tenantSharingCapabilityText) -or $tenantSharingCapabilityText -eq 'Not validated from the reviewed data') {
        $tenantSharingCapabilityText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('TenantSharingCapability')) -Default 'Not validated from the reviewed data'
    }
    $tenantSharingCapabilityText = Convert-ToCustomerSharingCapabilityLabel -Value $tenantSharingCapabilityText
    $defaultSharingLinkTypeText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('DefaultSharingLinkType')) -Default ''
    if ([string]::IsNullOrWhiteSpace($defaultSharingLinkTypeText) -or $defaultSharingLinkTypeText -eq 'Not validated from the reviewed data') {
        $defaultSharingLinkTypeText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $guestAccessConfigurationRecord -Names @('DefaultSharingLinkType')) -Default ''
    }
    if ([string]::IsNullOrWhiteSpace($defaultSharingLinkTypeText) -or $defaultSharingLinkTypeText -eq 'Not validated from the reviewed data') {
        $defaultSharingLinkTypeText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('DefaultSharingLinkType')) -Default 'Not validated from the reviewed data'
    }
    $defaultSharingLinkTypeText = Convert-ToCustomerSharingLinkTypeLabel -Value $defaultSharingLinkTypeText
    $crossTenantPartnerCountText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('CrossTenantPartnerCount')) -Default 'Not validated from the reviewed data'
    $defaultInboundMfaTrustText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('DefaultInboundMfaTrust')) -Default 'Not validated from the reviewed data'
    $sharingDomainRestrictionModeText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SharingDomainRestrictionMode')) -Default 'Not validated from the reviewed data'
    $sharingDomainRestrictionModeText = Convert-ToCustomerSharingRestrictionModeLabel -Value $sharingDomainRestrictionModeText
    $siteOverrideCountText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SiteOverrideCount')) -Default $(if ($externalSharingSiteOverrides.Count -gt 0) { $externalSharingSiteOverrides.Count } else { 'Not validated from the reviewed data' })
    $guestMfaWhyThisMattersText = Convert-ToCustomerAssessmentNarrativeText -Text (Get-ArrayaObjectValue -Object $guestMfaExperienceSummary -Names @('WhyThisMatters'))
    $guestMfaCurrentStateSummaryText = Convert-ToCustomerAssessmentNarrativeText -Text (Get-ArrayaObjectValue -Object $guestMfaExperienceSummary -Names @('CurrentStateSummary'))
    $guestMfaUserExperienceText = Convert-ToCustomerAssessmentNarrativeText -Text (Get-ArrayaObjectValue -Object $guestMfaExperienceSummary -Names @('UserExperience'))
    $guestMfaDesiredStateText = Convert-ToCustomerAssessmentNarrativeText -Text (Get-ArrayaObjectValue -Object $guestMfaExperienceSummary -Names @('DesiredState'))
    $guestMfaRecommendationStrategyText = Convert-ToCustomerAssessmentNarrativeText -Text (Get-ArrayaObjectValue -Object $guestMfaExperienceSummary -Names @('RecommendationStrategy'))
    $largestDomainsByRecipients = @(
        $domainRows |
            Sort-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalDomainRecipients')) } -Descending |
            Select-Object -First 5
    )
    $dmarcEnabledDomains = @($domainRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DmarcConfigured'))) -eq $true }).Count
    $dkimEnabledDomains = @($domainRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DkimConfigured'))) -eq $true }).Count
    $domainsWithoutDmarc = @(
        $domainRows | Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DmarcConfigured'))) -ne $true
        }
    )
    $domainsWithoutDkim = @(
        $domainRows | Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DkimConfigured'))) -ne $true
        }
    )
    $domainsWithoutSpfSignal = @(
        $domainRows | Where-Object {
            [string]::IsNullOrWhiteSpace((Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('SpfPolicyMode', 'SpfRecord')) -Default ''))
        }
    )
    $mailboxesWithExplicitRetentionPolicy = @(
        $mailboxRows | Where-Object {
            -not [string]::IsNullOrWhiteSpace((Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('RetentionPolicy')) -Default ''))
        }
    ).Count
    $mailboxesWithHoldSignals = @(
        $mailboxRows | Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('LitigationHoldEnabled'))) -eq $true -or
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('RetentionHoldEnabled'))) -eq $true -or
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DelayHoldApplied'))) -eq $true
        }
    ).Count
    $uniqueRetentionPolicyNames = @(
        $mailboxRows |
            ForEach-Object { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('RetentionPolicy')) -Default '' } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )

    $recommendationRows = @(
        foreach ($action in @($roadmapActions)) {
            New-CustomerWordTableRow -Cells @(
                [string]$action.ActionTitle,
                @(
                    (Convert-ToCustomerAssessmentDisplayText -Value ([string]$action.RoadmapPhase) -Default 'Not validated from the reviewed data'),
                    ('Impact: {0}' -f (Get-CustomerActionImpactLabel -Severity ([string]$action.HighestSeverity)))
                ),
                (Convert-ToCustomerAssessmentDisplayText -Value ([string]$action.FirstValidationStep) -Default 'Validate the current state behind this work item before scheduling remediation.'),
                (Convert-ToCustomerAssessmentDisplayText -Value ([string]$action.SuccessCheck) -Default 'The agreed target state should be validated in the relevant detailed section.'),
                (New-CustomerWordBlankCell)
            )
        }
    )
    if ($recommendationRows.Count -eq 0) {
        $recommendationRows = @(
            New-CustomerWordTableRow -Cells @(
                'Priority work item not clearly surfaced',
                @('Monitor', 'Impact: Info'),
                'Validate the current state behind this work item before scheduling remediation.',
                'The agreed target state should be validated in the relevant detailed section.',
                (New-CustomerWordBlankCell)
            )
        )
    }

    $recommendationImpactRows = @(
        foreach ($action in @($roadmapActions)) {
            New-CustomerWordTableRow -Cells @(
                (Convert-ToCustomerAssessmentDisplayText -Value ([string]$action.ActionTitle) -Default 'Priority work item'),
                (Convert-ToCustomerAssessmentNarrativeText -Text ([string]$action.UserImpactExperience))
            )
        }
    )
    if ($recommendationImpactRows.Count -eq 0) {
        $recommendationImpactRows = @(
            New-CustomerWordTableRow -Cells @(
                'Priority work item not clearly surfaced',
                'User impact should be confirmed from the detailed findings before remediation is scheduled.'
            )
        )
    }

    $overallFindingsSummaryRows = @(Get-CustomerOverallFindingsSummaryRows -WorkstreamSummaries $workstreamSummaries -Findings $findings)
    if ($overallFindingsSummaryRows.Count -eq 0) {
        $overallFindingsSummaryRows = @(
            New-CustomerWordTableRow -Cells @(
                'Assessment summary not clearly surfaced',
                'Info',
                '0',
                'The current source did not include workstream summary rows, so the detailed findings inventory should be used as the primary crosswalk.'
            )
        )
    }
    $fullFindingsInventoryRows = if ($findings.Count -gt 0) {
        @(
            foreach ($finding in @($findings)) {
                $ruleId = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $finding -Names @('RuleId')) -Default 'Not surfaced in current source'
                $findingText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $finding -Names @('Finding')) -Default 'Not surfaced in current source'
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $finding -Names @('Severity')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $finding -Names @('PriorityBand')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $finding -Names @('OwnerTeam', 'Area')) -Default 'Not surfaced in current source'),
                    ("{0}: {1}" -f $ruleId, $findingText),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $finding -Names @('WhyFlagged')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $finding -Names @('Recommendation')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $finding -Names @('TargetValue')) -Default 'Not surfaced in current source')
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }

    $userSummaryRows = @(
        @('Total users reviewed', $(if ($userRows.Count -gt 0) { $userRows.Count } else { 'Not surfaced in current source' })),
        @('Guest accounts', $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { $guestUsers.Count } else { 'Not surfaced in current source' })),
        @('Internal members', $(if ($memberUsers.Count -gt 0 -or $userRows.Count -gt 0) { $memberUsers.Count } else { 'Not surfaced in current source' })),
        @('Enabled internal members with no sign-in within 180 days', $(if ($enabledMemberUsers.Count -gt 0 -or $userRows.Count -gt 0) { $inactiveMemberUsers.Count } else { 'Not surfaced in current source' })),
        @('Inactive guest accounts (>90 days)', $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { $inactiveGuestUsers.Count } else { (Get-CustomerObservationState -Observation $lifecycleObservation -Signal 'Inactive guest accounts (>90 days)') }))
    )
    $userComparisonRows = @(
        @('Internal members', $(if ($memberUsers.Count -gt 0 -or $userRows.Count -gt 0) { $memberUsers.Count } else { 'Not surfaced in current source' }), $(if ($enabledMemberUsers.Count -gt 0 -or $userRows.Count -gt 0) { $enabledMemberUsers.Count } else { 'Not surfaced in current source' }), $(if ($enabledMemberUsers.Count -gt 0 -or $userRows.Count -gt 0) { $inactiveMemberUsers.Count } else { 'Not surfaced in current source' })),
        @('Guest accounts', $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { $guestUsers.Count } else { 'Not surfaced in current source' }), $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { @($guestUsers | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))) -ne $false }).Count } else { 'Not surfaced in current source' }), $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { $inactiveGuestUsers.Count } else { 'Not surfaced in current source' }))
    )
    $deviceSummaryRows = @(
        @('Total devices reviewed', (Get-CustomerObservationState -Observation $deviceObservation -Signal 'Devices represented in current source')),
        @('Dominant platform', $dominantPlatform),
        @('Non-compliant devices', (Get-CustomerObservationState -Observation $deviceObservation -Signal 'Non-compliant devices')),
        @('Unmanaged devices', (Get-CustomerObservationState -Observation $deviceObservation -Signal 'Unmanaged devices')),
        @('Stale devices (>180 days)', (Get-CustomerObservationState -Observation $deviceObservation -Signal 'Stale devices (>180 days)')),
        @('Intune management coverage', (Get-CustomerObservationState -Observation $deviceObservation -Signal 'Intune management coverage'))
    )
    $platformDistributionRows = if ($platformGroups.Count -gt 0) {
        @(
            foreach ($platformGroup in $platformGroups) {
                $share = if ($deviceRows.Count -gt 0) { '{0}%' -f [math]::Round(($platformGroup.Count / $deviceRows.Count) * 100, 2) } else { 'Not surfaced in current source' }
                New-CustomerWordTableRow -Cells @($platformGroup.Name, $platformGroup.Count, $share)
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }
    $noteworthyEnterpriseApplications = if ($enterpriseApplications.Count -gt 0) {
        @(Get-CustomerNoteworthyEnterpriseApplications -ApplicationRows $enterpriseApplications -Top 8)
    }
    else {
        @()
    }
    $applicationInventoryRows = if ($enterpriseApplications.Count -gt 0) {
        @(
            foreach ($enterpriseApplication in @($noteworthyEnterpriseApplications)) {
                $displayName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('DisplayName')) -Default 'Unnamed application'
                $preferredSingleSignOnMode = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('SSOMode', 'PreferredSingleSignOnMode')) -Default 'Not detected'
                $ssoEnabled = Test-CustomerEnterpriseApplicationSsoEnabled -ApplicationRow $enterpriseApplication
                $ssoEnabledText = Convert-ToCustomerAssessmentBooleanLabel -Value $ssoEnabled -TrueText 'Enabled' -FalseText 'Not detected' -Default 'Not validated from the reviewed data'
                $effectiveSsoMode = if ($ssoEnabled -eq $true -and $preferredSingleSignOnMode -ne 'Not detected') { $preferredSingleSignOnMode } elseif ($ssoEnabled -eq $true) { 'Configured' } else { 'Not detected' }
                $highPrivilegeCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('HighPrivilegePermissionCount'))
                $delegatedCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('DelegatedPermissionGrantCount'))
                $applicationPermissionCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('ApplicationPermissionCount'))
                $appRoleAssignmentRequired = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('AppRoleAssignmentRequired'))
                $applicationSource = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('ApplicationSource')) -Default 'Source not classified'
                $ownerSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $enterpriseApplication -Name 'OwnerSignalState'
                $redirectUriSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $enterpriseApplication -Name 'RedirectUriSignalState'
                $activitySignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $enterpriseApplication -Name 'ActivitySignalState'
                $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('OwnerCount'))
                $appCredentials = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('AppCredentials')) -Default ''
                $delegatedLastSignIn = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('DelegatedLastSignIn')) -Default ''
                $applicationLastSignIn = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('ApplicationLastSignIn')) -Default ''
                $insecureRedirectUriCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('InsecureRedirectUriCount'))
                $hasRecentActivity = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('HasRecentActivity'))
                $lastSignInDateTime = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('LastSignInDateTime')) -Default ''
                $lastSignInUserDisplayName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('LastSignInUserDisplayName')) -Default ''
                $lastSignInUserPrincipalName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('LastSignInUserPrincipalName')) -Default ''
                $lastConditionalAccessStatus = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('LastConditionalAccessStatus')) -Default ''
                $lastClientAppUsed = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('LastClientAppUsed')) -Default ''

                $observationParts = New-Object System.Collections.Generic.List[string]
                if (-not [string]::IsNullOrWhiteSpace($applicationSource) -and $applicationSource -ne 'Source not classified') {
                    $observationParts.Add("Application source is classified as $applicationSource.") | Out-Null
                }
                if ($ssoEnabled -eq $true) {
                    $observationParts.Add(("SSO is configured{0}." -f $(if ($effectiveSsoMode -ne 'Configured') { " via $effectiveSsoMode" } else { '' }))) | Out-Null
                }
                if ($null -ne $highPrivilegeCount -and $highPrivilegeCount -gt 0) {
                    $observationParts.Add('High-privilege permissions were surfaced in the reviewed data.') | Out-Null
                } elseif (($null -ne $applicationPermissionCount -and $applicationPermissionCount -gt 0) -or ($null -ne $delegatedCount -and $delegatedCount -gt 0)) {
                    $observationParts.Add('Permission grants were surfaced in the reviewed data.') | Out-Null
                }
                if ($appRoleAssignmentRequired -eq $true) {
                    $observationParts.Add('User assignment is required before access is granted.') | Out-Null
                }
                if ($ownerSignalState -eq 'Collected' -and $null -ne $ownerCount) {
                    if ($ownerCount -le 0) {
                        $observationParts.Add('No owner signal was surfaced for this application registration.') | Out-Null
                    }
                    else {
                        $observationParts.Add(("Owner coverage shows {0} owner signal(s)." -f $ownerCount)) | Out-Null
                    }
                }
                elseif ($ownerSignalState -eq 'Unavailable') {
                    $observationParts.Add('Owner validation was not fully completed in the reviewed data.') | Out-Null
                }
                if (-not [string]::IsNullOrWhiteSpace($appCredentials)) {
                    $observationParts.Add(("Application credentials surfaced: {0}." -f $appCredentials)) | Out-Null
                }
                if ($null -ne $insecureRedirectUriCount -and $insecureRedirectUriCount -gt 0) {
                    $observationParts.Add(("Redirect URI review flagged {0} potentially insecure or overly broad endpoint(s)." -f $insecureRedirectUriCount)) | Out-Null
                }
                elseif ($redirectUriSignalState -eq 'Partial') {
                    $observationParts.Add('Redirect URI review was only partially validated in the reviewed data.') | Out-Null
                }
                if (-not [string]::IsNullOrWhiteSpace($delegatedLastSignIn) -and $delegatedLastSignIn -ne 'Not validated from the reviewed data') {
                    $observationParts.Add(("Latest delegated app sign-in: {0}." -f $delegatedLastSignIn)) | Out-Null
                }
                if (-not [string]::IsNullOrWhiteSpace($applicationLastSignIn) -and $applicationLastSignIn -ne 'Not validated from the reviewed data') {
                    $observationParts.Add(("Latest application credential sign-in: {0}." -f $applicationLastSignIn)) | Out-Null
                }
                if (-not [string]::IsNullOrWhiteSpace($lastSignInDateTime)) {
                    if ($lastSignInDateTime -eq 'No sign-ins found') {
                        $observationParts.Add('No recent sign-in was surfaced for this application in the reviewed sign-in logs.') | Out-Null
                    }
                    else {
                        $lastSignInObservation = "Latest sign-in seen on $lastSignInDateTime"
                        if (-not [string]::IsNullOrWhiteSpace($lastSignInUserDisplayName) -and $lastSignInUserDisplayName -ne 'Not validated from the reviewed data') {
                            $lastSignInObservation += " by $lastSignInUserDisplayName"
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace($lastSignInUserPrincipalName) -and $lastSignInUserPrincipalName -ne 'Not validated from the reviewed data') {
                            $lastSignInObservation += " by $lastSignInUserPrincipalName"
                        }
                        $lastSignInObservation += '.'
                        $observationParts.Add($lastSignInObservation) | Out-Null
                    }
                }
                if (-not [string]::IsNullOrWhiteSpace($lastConditionalAccessStatus) -and $lastConditionalAccessStatus -ne 'Not validated from the reviewed data') {
                    $observationParts.Add("Latest sign-in Conditional Access status: $lastConditionalAccessStatus.") | Out-Null
                }
                if (-not [string]::IsNullOrWhiteSpace($lastClientAppUsed) -and $lastClientAppUsed -ne 'Not validated from the reviewed data') {
                    $observationParts.Add("Latest client app used: $lastClientAppUsed.") | Out-Null
                }
                if ($activitySignalState -eq 'Collected' -and $hasRecentActivity -eq $false) {
                    $observationParts.Add('No recent activity signal was surfaced across the reviewed application sign-in telemetry.') | Out-Null
                }
                elseif ($activitySignalState -eq 'Partial') {
                    $observationParts.Add('Recent activity validation was only partially completed in the reviewed data.') | Out-Null
                }
                if ($observationParts.Count -eq 0) {
                    $observationParts.Add('Base application inventory was surfaced, but deeper permission usage was not validated from the reviewed data.') | Out-Null
                }

                New-CustomerWordTableRow -Cells @($displayName, $ssoEnabledText, $effectiveSsoMode, ($observationParts -join ' '))
            }
        )
    }
    elseif ($applicationInventoryClearlyEmpty) {
        @((New-CustomerWordTableRow -Cells @('No enterprise applications surfaced in the reviewed data', 'Not detected', 'Not detected', 'Treat this as a current-state result and validate it if the tenant expects third-party or line-of-business applications.')))
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Enterprise application inventory not validated', 'Validation note', 'Validation note', 'The current review did not surface a usable enterprise application inventory, so this section should not be read as proof that no enterprise applications exist.')))
    }
    $globalAdminTableRows = if ($gaRows.Count -gt 0) {
        @(
            foreach ($globalAdmin in @($gaRows | Select-Object -First 15)) {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $globalAdmin -Names @('DisplayName')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $globalAdmin -Names @('CreatedDateTime', 'CreatedDate', 'Created')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $globalAdmin -Names @('UserPrincipalName', 'UPN')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $globalAdmin -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime')) -Default 'Not surfaced in current source')
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }
    $domainOverviewRows = if ($domainRows.Count -gt 0) {
        @(
            foreach ($domainRow in $domainRows) {
                $dmarcValue = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('DmarcPolicy', 'DmarcRecord')) -Default ''
                if ([string]::IsNullOrWhiteSpace($dmarcValue)) {
                    $dmarcConfigured = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $domainRow -Names @('DmarcConfigured'))
                    if ($null -eq $dmarcConfigured) {
                        $dmarcValue = 'Not surfaced in current source'
                    }
                    else {
                        $dmarcValue = if ($dmarcConfigured) { 'Configured' } else { 'Not configured' }
                    }
                }
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('Domain', 'Id')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('IsVerified', 'Verified')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('AuthenticationType')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('DomainType')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('IsDefault', 'Default')) -Default 'Not surfaced in current source'),
                    $dmarcValue
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }
    $dnsRows = if ($domainRows.Count -gt 0) {
        @(
            foreach ($domainRow in $domainRows) {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('Domain', 'Id')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('SpfPolicyMode', 'SpfRecord')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('DkimConfigured')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('DmarcPolicy', 'DmarcConfigured')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('AuthenticationType')) -Default 'Not surfaced in current source')
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }
    $averageUserMailboxSizeBytes = if ($userMailboxStatsRows.Count -gt 0) { (($userMailboxStatsRows | ForEach-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes')) } | Measure-Object -Sum).Sum / $userMailboxStatsRows.Count) } else { $null }
    $averageSharedMailboxSizeBytes = if ($sharedMailboxStatsRows.Count -gt 0) { (($sharedMailboxStatsRows | ForEach-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes')) } | Measure-Object -Sum).Sum / $sharedMailboxStatsRows.Count) } else { $null }
    $averageGroupMailboxSizeBytes = if ($groupMailboxStatsRows.Count -gt 0) { (($groupMailboxStatsRows | ForEach-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes')) } | Measure-Object -Sum).Sum / $groupMailboxStatsRows.Count) } else { $null }
    $smtpAuthUsersCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $smtpRelayConfigRecord -Names @('SMTPAuthUsers'))
    if ($null -eq $smtpAuthUsersCount) { $smtpAuthUsersCount = $smtpAuthMailboxRows.Count }
    $relayConnectorCount = if ($relayConnectorRows.Count -gt 0) { $relayConnectorRows.Count } else { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $smtpRelayConfigRecord -Names @('RelayConnectorCount')) }
    $topSendCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $emailActivitySummaryRecord -Names @('TotalSendCount'))
    $activeSenderCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $emailActivitySummaryRecord -Names @('ActiveUsers'))
    $teamsVoiceUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $teamsVoiceSummaryRecord -Names @('VoiceUserCount'))
    $teamsVoiceSource = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $teamsVoiceSummaryRecord -Names @('Source')) -Default 'Not surfaced in current source'
    $dmarcRows = if ($domainRows.Count -gt 0) {
        @(
            foreach ($domainRow in $domainRows) {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('Domain', 'Id')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('DmarcConfigured')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('DmarcPolicy')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('DmarcRecord')) -Default 'Not surfaced in current source')
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }
    $userMailboxGrowthRows = @(
        @('User mailboxes represented in statistics', $userMailboxStatsRows.Count),
        @('User mailboxes under 1 GB', $userMailboxesUnderOneGb),
        @('Largest user mailbox', $(if ($largestUserMailbox.Count -gt 0) { '{0} ({1} GB)' -f (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $largestUserMailbox[0] -Names @('DisplayName')) -Default 'Not surfaced in current source'), (& $formatSizeGb (Get-ArrayaObjectValue -Object $largestUserMailbox[0] -Names @('TotalItemSizeBytes'))) } else { 'Not surfaced in current source' })),
        @('Average user mailbox size', $(if ($null -ne $averageUserMailboxSizeBytes) { (& $formatSizeGb $averageUserMailboxSizeBytes) + ' GB' } else { 'Not surfaced in current source' })),
        @('Active archive mailboxes', $archiveMailboxRows.Count)
    )
    $sharedMailboxReviewRows = @(
        @('Shared mailboxes represented in statistics', $sharedMailboxStatsRows.Count),
        @('Shared mailboxes under 1 GB', $sharedMailboxesUnderOneGb),
        @('Largest shared mailbox', $(if ($largestSharedMailbox.Count -gt 0) { '{0} ({1} GB)' -f (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $largestSharedMailbox[0] -Names @('DisplayName')) -Default 'Not surfaced in current source'), (& $formatSizeGb (Get-ArrayaObjectValue -Object $largestSharedMailbox[0] -Names @('TotalItemSizeBytes'))) } else { 'Not surfaced in current source' })),
        @('Average shared mailbox size', $(if ($null -ne $averageSharedMailboxSizeBytes) { (& $formatSizeGb $averageSharedMailboxSizeBytes) + ' GB' } else { 'Not surfaced in current source' })),
        @('Shared mailboxes without ownership signal', (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Shared mailboxes without ownership signal'))
    )
    $inactiveMailboxSummaryRows = @(
        @('Inactive mailboxes', $inactiveMailboxRows.Count),
        @('Inactive mailbox types represented', $(if ($inactiveMailboxRows.Count -gt 0) { ((@($inactiveMailboxRows | Group-Object { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('RecipientTypeDetails')) -Default 'Unknown' }) | Sort-Object Count -Descending | ForEach-Object { '{0}: {1}' -f $_.Name, $_.Count }) -join '; ') } else { 'Not surfaced in current source' })),
        @('Example inactive mailboxes', $(if ($inactiveMailboxRows.Count -gt 0) { ((@($inactiveMailboxRows | Select-Object -First 2 | ForEach-Object { '{0} ({1})' -f (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Not surfaced in current source'), (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('PrimarySmtpAddress', 'UserPrincipalName')) -Default 'address not surfaced') })) -join '; ') } else { 'Not surfaced in current source' }))
    )
    $groupMailboxUtilizationRows = @(
        @('Group mailboxes represented in statistics', $groupMailboxStatsRows.Count),
        @('Largest group mailbox', $(if ($largestGroupMailbox.Count -gt 0) { '{0} ({1} GB)' -f (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $largestGroupMailbox[0] -Names @('DisplayName')) -Default 'Not surfaced in current source'), (& $formatSizeGb (Get-ArrayaObjectValue -Object $largestGroupMailbox[0] -Names @('TotalItemSizeBytes'))) } else { 'Not surfaced in current source' })),
        @('Average group mailbox size', $(if ($null -ne $averageGroupMailboxSizeBytes) { (& $formatSizeGb $averageGroupMailboxSizeBytes) + ' GB' } else { 'Not surfaced in current source' })),
        @('Dormant Teams connected to review scope', $dormantTeams)
    )
    $archiveMailboxUsageRows = @(
        @('Archive mailboxes', $archiveMailboxRows.Count),
        @('Archive statistics represented', $archiveMailboxStatsRows.Count),
        @('Archives over 50 GB', $archiveMailboxesOverFiftyGb),
        @('Archive mailboxes with hold signals', $archiveMailboxesWithHolds),
        @('Non-user mailboxes in scope', $nonUserMailboxRows.Count)
    )
    $topRecipientTypesText = Join-ArrayaReadableList -Items @(
        @(
            Get-CustomerObservationState -Observation $messagingObservation -Signal 'Top recipient type 1'
            Get-CustomerObservationState -Observation $messagingObservation -Signal 'Top recipient type 2'
            Get-CustomerObservationState -Observation $messagingObservation -Signal 'Top recipient type 3'
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '^Not validated' }
    )
    if ([string]::IsNullOrWhiteSpace($topRecipientTypesText)) {
        $topRecipientTypesText = 'Not validated from the reviewed data'
    }

    $topRecipientDomainsText = Join-ArrayaReadableList -Items @(
        @(
            Get-CustomerObservationState -Observation $messagingObservation -Signal 'Top recipient domain 1'
            Get-CustomerObservationState -Observation $messagingObservation -Signal 'Top recipient domain 2'
            Get-CustomerObservationState -Observation $messagingObservation -Signal 'Top recipient domain 3'
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '^Not validated' -and $_ -notmatch ': 0 total' } | Select-Object -First 2
    )
    if ([string]::IsNullOrWhiteSpace($topRecipientDomainsText)) {
        $topRecipientDomainsText = 'Not validated from the reviewed data'
    }

    $messagingSnapshotRows = if ($null -ne $messagingConsultativeSummary -and @($messagingConsultativeSummary.SnapshotRows).Count -gt 0) {
        Convert-CustomerConfigurationRowsToWordTableRows -Rows @($messagingConsultativeSummary.SnapshotRows)
    }
    else {
        @(
            New-CustomerWordTableRow -Cells @('Recipients in current source', (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Recipients in current source'))
            New-CustomerWordTableRow -Cells @('Largest recipient mix', $topRecipientTypesText)
            New-CustomerWordTableRow -Cells @('Most concentrated recipient domains', $topRecipientDomainsText)
            New-CustomerWordTableRow -Cells @('Mailboxes with forwarding configured', (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Mailboxes with forwarding configured'))
            New-CustomerWordTableRow -Cells @('Shared mailboxes without ownership signal', (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Shared mailboxes without ownership signal'))
            New-CustomerWordTableRow -Cells @('Public folders still present', (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Public folders still present'))
        )
    }
    $mailboxLifecycleRows = if ($null -ne $messagingConsultativeSummary -and @($messagingConsultativeSummary.SecondaryRows).Count -gt 0) {
        Convert-CustomerConfigurationRowsToWordTableRows -Rows @($messagingConsultativeSummary.SecondaryRows)
    }
    else {
        @(
            New-CustomerWordTableRow -Cells @('Shared mailboxes without owner signal', (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Shared mailboxes without ownership signal'))
            New-CustomerWordTableRow -Cells @('Oversized shared mailboxes', 'Not validated from the reviewed data')
            New-CustomerWordTableRow -Cells @('Inactive mailboxes', $inactiveMailboxRows.Count)
            New-CustomerWordTableRow -Cells @('Mailboxes with hold signals', $mailboxesWithHoldSignals)
            New-CustomerWordTableRow -Cells @('Archive-enabled mailboxes', $archiveMailboxRows.Count)
            New-CustomerWordTableRow -Cells @('Archive mailboxes over 50 GB', $archiveMailboxesOverFiftyGb)
        )
    }
    $transportExposureRows = if ($null -ne $messagingConsultativeSummary -and @($messagingConsultativeSummary.TertiaryRows).Count -gt 0) {
        Convert-CustomerConfigurationRowsToWordTableRows -Rows @($messagingConsultativeSummary.TertiaryRows)
    }
    else {
        @(
            New-CustomerWordTableRow -Cells @('Outbound transport posture', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $smtpRelayConfigRecord -Names @('ConnectorBasedRelay')) -Default 'Not surfaced in current source'))
            New-CustomerWordTableRow -Cells @('Relay connectors', $(if ($null -ne $relayConnectorCount) { $relayConnectorCount } else { 'Not surfaced in current source' }))
            New-CustomerWordTableRow -Cells @('SMTP-authenticated accounts', $smtpAuthUsersCount)
            New-CustomerWordTableRow -Cells @('Public folder objects', $publicFolderRows.Count)
            New-CustomerWordTableRow -Cells @('External forwarding exposure', (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Mailboxes with forwarding configured'))
        )
    }
    $smtpRelayUsageRows = @(
        @('SMTP AUTH enabled', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $smtpRelayConfigRecord -Names @('SMTPAuthEnabled')) -Default 'Not surfaced in current source')),
        @('SMTP-authenticated accounts', $smtpAuthUsersCount),
        @('Connector-based relay', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $smtpRelayConfigRecord -Names @('ConnectorBasedRelay')) -Default 'Not surfaced in current source')),
        @('Relay connectors', $(if ($null -ne $relayConnectorCount) { $relayConnectorCount } else { 'Not surfaced in current source' })),
        @('Direct send enabled', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $smtpRelayConfigRecord -Names @('DirectSendEnabled')) -Default 'Not surfaced in current source')),
        @('Total send count in activity summary', $(if ($null -ne $topSendCount) { $topSendCount } else { 'Not surfaced in current source' })),
        @('Active sender count in activity summary', $(if ($null -ne $activeSenderCount) { $activeSenderCount } else { 'Not surfaced in current source' }))
    )
    $smtpRelayServiceAccountRows = if ($relayServiceAccountRows.Count -gt 0) {
        @(
            foreach ($relayServiceAccount in @($relayServiceAccountRows | Select-Object -First 10)) {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $relayServiceAccount -Names @('DisplayName', 'UserPrincipalName')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $relayServiceAccount -Names @('UserPrincipalName', 'PrimarySmtpAddress')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $relayServiceAccount -Names @('SendCount')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $relayServiceAccount -Names @('LastActivityDate')) -Default 'Not surfaced in current source')
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }
    $teamsGovernanceRows = @(
        @('Teams reviewed', $teamRows.Count),
        @('Ownerless Teams', (Get-CustomerObservationState -Observation $collaborationObservation -Signal 'Ownerless Teams')),
        @('Dormant Teams (>90 days)', $dormantTeams),
        @('Archived Teams', $archivedTeams),
        @('Guest-heavy Teams', (Get-CustomerObservationState -Observation $collaborationObservation -Signal 'Guest-heavy Teams')),
        @('Voice-user signal', $(if ($null -ne $teamsVoiceUsers) { "$teamsVoiceUsers user(s); source=$teamsVoiceSource" } else { 'Not surfaced in current source' }))
    )
    $sharePointStorageRows = if ($largestSharePointSites.Count -gt 0) {
        @(
            foreach ($sharePointSite in $largestSharePointSites) {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSite -Names @('Title')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSite -Names @('Owner')) -Default 'Not surfaced in current source'),
                    $(if ($null -ne (Get-ArrayaObjectValue -Object $sharePointSite -Names @('StorageUsedGB'))) { ('{0:N2}' -f [double](Get-ArrayaObjectValue -Object $sharePointSite -Names @('StorageUsedGB'))) } else { 'Not surfaced in current source' }),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSite -Names @('LastContentModifiedDate')) -Default 'Not surfaced in current source')
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }
    $retentionRows = @(
        @('Retention policies surfaced', (Get-CustomerObservationState -Observation $governanceObservation -Signal 'Retention policies surfaced')),
        @('Mailbox retention policies assigned', $mailboxesWithExplicitRetentionPolicy),
        @('Mailboxes with hold signals', $mailboxesWithHoldSignals),
        @('Unique retention policies named', $(if ($uniqueRetentionPolicyNames.Count -gt 0) { ($uniqueRetentionPolicyNames | Select-Object -First 5) -join '; ' } else { 'Not surfaced in current source' })),
        @('Data loss prevention controls', $(if ($dlpPolicyRows.Count -gt 0) { $dlpPolicyRows.Count } else { 'Not surfaced in current source' })),
        @('Secure Score posture', (Get-CustomerObservationState -Observation $governanceObservation -Signal 'Secure Score')),
        @('Highest license utilization', (Get-CustomerObservationState -Observation $governanceObservation -Signal 'Highest license utilization'))
    )
    $offboardingSupportRows = @(
        @('Inactive guest accounts (>90 days)', (Get-CustomerObservationState -Observation $lifecycleObservation -Signal 'Inactive guest accounts (>90 days)')),
        @('Stale privileged accounts (>90 days)', (Get-CustomerObservationState -Observation $lifecycleObservation -Signal 'Stale privileged accounts (>90 days)')),
        @('Inactive or disabled licensed users', (Get-CustomerObservationState -Observation $lifecycleObservation -Signal 'Inactive or disabled licensed users')),
        @('Stale SharePoint / OneDrive locations', (Get-CustomerObservationState -Observation $lifecycleObservation -Signal 'Stale SharePoint / OneDrive locations')),
        @('Shared mailboxes without ownership signal', (Get-CustomerObservationState -Observation $lifecycleObservation -Signal 'Shared mailboxes without ownership signal'))
    )
    $externalSharingSnapshotRows = @(
        @('Tenant sharing capability', $tenantSharingCapabilityText),
        @('OneDrive sharing capability', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('OneDriveSharingCapability')) -Default 'Not validated from the reviewed data')),
        @('Default sharing link type', $defaultSharingLinkTypeText),
        @('Default link permission', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('DefaultLinkPermission')) -Default 'Not validated from the reviewed data')),
        @('Sharing domain restriction mode', $sharingDomainRestrictionModeText),
        @('Allowed sharing domains', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SharingAllowedDomainList')) -Default 'Not validated from the reviewed data')),
        @('Blocked sharing domains', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SharingBlockedDomainList')) -Default 'Not validated from the reviewed data')),
        @('Anonymous link expiration (days)', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('AnonymousLinkExpirationInDays')) -Default 'Not validated from the reviewed data')),
        @('Require invited-user match', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('RequireInvitedUserMatch')) -TrueText 'Required' -FalseText 'Not required' -Default 'Not validated from the reviewed data')),
        @('Prevent external users from resharing', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('PreventExternalUsersFromResharing')) -TrueText 'Prevented' -FalseText 'Allowed' -Default 'Not validated from the reviewed data')),
        @('Guest access expiration (days)', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('GuestExpirationInDays')) -Default 'Not validated from the reviewed data')),
        @('Sites with external sharing enabled', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SitesWithExternalSharingEnabled')) -Default 'Not validated from the reviewed data')),
        @('Site override count', $siteOverrideCountText)
    )
    $sharePointTenantControlRows = @(
        @('Tenant settings collection source', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('CollectionSource')) -Default 'Not validated from the reviewed data')),
        @('Tenant settings API version', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('SettingsApiVersion')) -Default 'Not validated from the reviewed data')),
        @('Deleted user personal site retention (days)', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('DeletedUserPersonalSiteRetentionPeriodInDays')) -Default 'Not validated from the reviewed data')),
        @('OneDrive default storage quota (MB)', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('OneDriveStorageQuotaMB')) -Default 'Not validated from the reviewed data')),
        @('Site creation default storage quota (MB)', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('SiteCreationDefaultStorageQuotaMB')) -Default 'Not validated from the reviewed data')),
        @('Automatic site storage limits', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsSitesStorageLimitAutomatic')) -TrueText 'Automatic' -FalseText 'Manual' -Default 'Not validated from the reviewed data')),
        @('Legacy auth protocols enabled', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsLegacyAuthProtocolsEnabled')) -TrueText 'Enabled' -FalseText 'Disabled' -Default 'Not validated from the reviewed data')),
        @('Custom app authentication disabled', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('DisableCustomAppAuthentication')) -TrueText 'Disabled' -FalseText 'Allowed' -Default 'Not validated from the reviewed data')),
        @('Unmanaged sync app restricted', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsUnmanagedSyncAppForTenantRestricted')) -TrueText 'Restricted' -FalseText 'Not restricted' -Default 'Not validated from the reviewed data')),
        @('Sync button hidden on personal sites', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsSyncButtonHiddenOnPersonalSite')) -TrueText 'Hidden' -FalseText 'Visible' -Default 'Not validated from the reviewed data')),
        @('Site creation enabled', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsSiteCreationEnabled')) -TrueText 'Enabled' -FalseText 'Disabled' -Default 'Not validated from the reviewed data')),
        @('Site creation UI enabled', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsSiteCreationUIEnabled')) -TrueText 'Enabled' -FalseText 'Disabled' -Default 'Not validated from the reviewed data')),
        @('Loop enabled', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsLoopEnabled')) -TrueText 'Enabled' -FalseText 'Disabled' -Default 'Not validated from the reviewed data'))
    )
    $externalAccessSnapshotRows = @(
        @('Guest invitation control', $guestInvitationControlText),
        @('Allow email-verified users to join organization', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('AllowEmailVerifiedUsersToJoinOrganization')) -TrueText 'Allowed' -FalseText 'Not allowed' -Default 'Not validated from the reviewed data')),
        @('Guest user role', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('GuestUserRoleLabel', 'GuestUserRoleId')) -Default 'Not validated from the reviewed data')),
        @('Conditional Access guest coverage', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $guestAccessConfigurationRecord -Names @('ConditionalAccessGuestCoverage')) -TrueText 'Detected' -FalseText 'Not detected' -Default 'Not validated from the reviewed data')),
        @('Has cross-tenant access policy', (Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('HasCrossTenantAccessPolicy')) -TrueText 'Present' -FalseText 'Not present' -Default 'Not validated from the reviewed data')),
        @('Cross-tenant partner count', $crossTenantPartnerCountText),
        @('Default inbound MFA trust', $defaultInboundMfaTrustText),
        @('Default outbound MFA trust', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('DefaultOutboundMfaTrust')) -Default 'Not validated from the reviewed data')),
        @('Default B2B direct connect inbound', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('DefaultB2BDirectConnectInbound')) -Default 'Not validated from the reviewed data')),
        @('Default B2B direct connect outbound', (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('DefaultB2BDirectConnectOutbound')) -Default 'Not validated from the reviewed data'))
    )
    $externalAccessSnapshotTableRows = @(
        (New-CustomerWordTableRow -Cells @('Inactive guest accounts (>90 days)', $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { $inactiveGuestUsers.Count } else { (Get-CustomerObservationState -Observation $lifecycleObservation -Signal 'Inactive guest accounts (>90 days)') })))
    ) + @(
        $externalAccessSnapshotRows |
            ForEach-Object { New-CustomerWordTableRow -Cells @($_[0], $_[1]) }
    )
    $externalExposureReviewRows = if ($externalExposureFindings.Count -gt 0) {
        @(
            foreach ($externalExposureFinding in @($externalExposureFindings | Select-Object -First 10)) {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalExposureFinding -Names @('Workload')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalExposureFinding -Names @('AssetType')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalExposureFinding -Names @('Title')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalExposureFinding -Names @('ExposureCategory')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalExposureFinding -Names @('GapReason')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalExposureFinding -Names @('ReviewPriority')) -Default 'Not surfaced in current source')
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'No externally exposed assets or tenant-level posture gaps were flagged from the reviewed data.', 'Not surfaced in current source', 'The reviewed source did not surface an actionable external exposure row.', 'Not surfaced in current source')))
    }
    $staleExternalExposureCount = @($externalExposureFindings | Where-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('ExposureCategory')) -eq 'Stale externally shared content' }).Count
    $ownerDriftExternalExposureCount = @($externalExposureFindings | Where-Object {
        ([string](Get-ArrayaObjectValue -Object $_ -Names @('ExposureCategory')) -eq 'Externally sharable OneDrive with ownership mismatch') -or
        ((Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('OwnerSignal')) -Default '') -match '0 owners|ownership mismatch')
    }).Count
    $guestHeavyExposureCount = @($externalExposureFindings | Where-Object {
        ([string](Get-ArrayaObjectValue -Object $_ -Names @('ExposureCategory'))) -match 'Guest-heavy|Dormant guest-enabled|Guest-enabled ownerless'
    }).Count
    $siteOverrideExposureCount = @($externalExposureFindings | Where-Object {
        ([string](Get-ArrayaObjectValue -Object $_ -Names @('ExposureCategory'))) -match 'site sharing|default sharing link type|default link permission'
    }).Count
    $identityActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Identity') -Max 3
    $endpointActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Endpoint') -Max 3
    $messagingActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Messaging') -Max 3
    $collaborationActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Collaboration') -Max 3
    $governanceActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Governance') -Max 3
    $domainRegistrationActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Keywords @('unverified domain', 'domain verification', 'accepted domain') -Max 2
    $dnsActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Keywords @('dns', 'spf', 'dkim', 'dmarc', 'spoof', 'anti-spoof') -Max 2
    $dnsFindings = Get-CustomerRelevantFindings -Findings $findings -Keywords @('spf', 'dkim', 'dmarc', 'dns', 'spoof', 'domain') -Max 3

    $blocks = New-Object System.Collections.Generic.List[object]
    $blocks.Add((New-CustomerWordParagraphBlock -Text "$tenantName Microsoft 365 Tenant Best Practices Assessment" -Style 'Title')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Prepared by: Arraya Solutions' -Style 'Subtitle')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Generated: {0}" -f $GeneratedAt.ToString('yyyy-MM-dd')) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Assessment version: {0}" -f $assessmentVersion) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Version History' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This section tracks the issued version of the assessment report so the customer-facing document has a clear revision record.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Date', 'Revision', 'Author', 'Description', 'Reviewers') -Rows @(
        (New-CustomerWordTableRow -Cells @($GeneratedAt.ToString('MM/dd/yyyy'), $assessmentVersion, 'Arraya Solutions', 'Current assessment report release', 'Not surfaced in current source'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '1.0 Introduction' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text "This assessment documents the Microsoft 365 state observed in $tenantName and is organized to separate orientation, risk, action, and supporting evidence." -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Assessment Snapshot At A Glance' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Signal', 'Current State') -Rows @($sourceSummaryRows | ForEach-Object { New-CustomerWordTableRow -Cells @($_.Signal, $_.State) }))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($SourceModel.SummaryText) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '2.0 Project Scope' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The scope of this report is limited to the Microsoft 365 signals surfaced in the tenant review across identity, devices, messaging, collaboration, governance, and lifecycle controls. The document focuses on the conditions that were visible in the tenant data and maps those conditions into the operating areas most likely to affect security, administration, and day-to-day support.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Where a template field expects a detail that was not visible in the tenant review, the report states that clearly rather than inferring a value that the current source did not support.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '3.0 Executive Summary' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text "This section is intended to give leadership a decision-ready view of where risk is clustering and why those patterns matter now." -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $executiveDecisionSummary) { $executiveDecisionSummary.Narrative } else { $SourceModel.ExecutiveNarrative })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Risk Clusters' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Risk Cluster', 'What Stands Out', 'Why Leadership Should Care') -Rows $executiveRiskRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Leadership Decision Brief' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Decision Focus', 'What Should Happen Next', 'Why Now') -Rows $leadershipDecisionRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This brief highlights the short list of leadership approvals that would remove the biggest blockers to a cleaner operating baseline.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Overall Findings Summary' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Workstream', 'Severity / Impact', 'Open Findings', 'What Stands Out') -Rows $overallFindingsSummaryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This table shows where findings are clustering before the report moves into the detailed sections.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Findings Legend' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Term', 'What It Means In This Report') -Rows @($findingsLegendRows | ForEach-Object { New-CustomerWordTableRow -Cells @($_.Term, $_.Meaning) }))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The full normalized finding inventory is included later in 15.10 Full Findings Inventory so each grouped issue can be traced back to its supporting evidence.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '4.0 Modern Workplace Recommendations' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This action matrix is the main execution view for the report. Use it to sequence the work, then use the section pages and 15.10 Full Findings Inventory to validate the supporting evidence.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Recommendation', 'Priority / Impact', 'First Step', 'Success Check', 'Level of Effort') -Rows $recommendationRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'User Impact / Expected Experience' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Recommendation', 'User Impact / Expected Experience') -Rows $recommendationImpactRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Level of Effort is intentionally left blank where the current source does not support a reliable time or complexity estimate. That field can be completed during delivery planning once the team confirms ownership, dependencies, and remediation scope.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '5.0 Entra ID Review: User and Device Inventory' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Review of the tenant configuration indicated that identity hygiene and device governance need to be read together in this environment. The same parts of the tenant that are carrying stale privileged access are also the parts of the environment where compliance-driven access control is least mature.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '5.1 Entra User' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'User Account Summary' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $userSummaryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The user inventory shows {0} internal member account(s) and {1} guest account(s) in the reviewed data. What stands out is not just the count itself, but the number of enabled member accounts that no longer show recent sign-in activity. That pattern usually points to lifecycle drift rather than a single isolated exception." -f $(if ($memberUsers.Count -gt 0 -or $userRows.Count -gt 0) { $memberUsers.Count } else { 'an unconfirmed number of' }), $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { $guestUsers.Count } else { 'an unconfirmed number of' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'User Account Status Comparison' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Account Type', 'Total', 'Enabled', 'Inactive') -Rows $userComparisonRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'When viewed collectively, the identity inventory suggests that access cleanup is not limited to one user population. Internal members and guest identities both show evidence of stale access, which increases the chance that a dormant identity still retains a path into collaboration, messaging, or privileged workflows.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '5.2 Entra Device' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Registration Summary' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $deviceSummaryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.ObservedNarrative) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device Platform Distribution' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Platform', 'Count', 'Share') -Rows $platformDistributionRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device Registration and Compliance Gaps' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The reviewed device inventory makes it possible to see where endpoint registration exists without the same level of follow-through in management and compliance. In this tenant, that gap is most visible where stale, unmanaged, or non-compliant devices remain part of the active footprint.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.WhyItMatters) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.PositiveNarrative) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '5.3 Entra Guest Access Configuration' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'External Access Snapshot' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $externalAccessSnapshotTableRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The guest access review showed that external collaboration is active enough to require ongoing governance rather than periodic cleanup. Inactive guest accounts and externally permissive sharing signals rarely create visible pain day to day, but they expand the tenant surface area if ownership decisions are delayed. The current guest invitation control is shown as {0}, and the tenant currently has {1} configured cross-tenant partner relationship(s) in the reviewed data." -f $guestInvitationControlText, $crossTenantPartnerCountText) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if (-not [string]::IsNullOrWhiteSpace($guestMfaWhyThisMattersText)) { $guestMfaWhyThisMattersText } else { ("This matters because guest access combines identity risk with collaboration exposure. Once guest lifecycle, sharing defaults, and sponsor accountability drift at the same time, it becomes harder to validate which external access is still justified. Default inbound MFA trust is currently shown as {0}, which means cross-tenant trust decisions should be reviewed alongside guest invitation settings rather than as a separate design concern." -f $defaultInboundMfaTrustText) })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if (-not [string]::IsNullOrWhiteSpace($guestMfaCurrentStateSummaryText)) { $guestMfaCurrentStateSummaryText } else { ("Default inbound MFA trust is currently shown as {0}, so guest MFA design should be reviewed alongside guest invitation settings rather than as a separate design concern." -f $defaultInboundMfaTrustText) })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if (-not [string]::IsNullOrWhiteSpace($guestMfaUserExperienceText)) { $guestMfaUserExperienceText } else { 'When guest MFA is enforced, the desired experience is to require strong authentication and trust the guest home-tenant MFA where that design is supported and approved.' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if (-not [string]::IsNullOrWhiteSpace($guestMfaDesiredStateText)) { $guestMfaDesiredStateText } else { 'The desired baseline is strong guest authentication with trusted home-tenant MFA where supported and approved, not a blanket requirement for every guest to register separately in the resource tenant.' })) -Style 'Normal')) | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($guestMfaRecommendationStrategyText)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text $guestMfaRecommendationStrategyText -Style 'Normal')) | Out-Null
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text '5.4 Entra Applications and Access Review' -Style 'Heading2')) | Out-Null
    if ($applicationInventoryClearlyEmpty) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The application review looked for enterprise application inventory, consent-related controls, and elevated permission signals. In the reviewed data, no enterprise applications were surfaced. That should be read as the current reported result for this source, with a follow-up validation only if the tenant expects line-of-business or third-party enterprise applications to appear here.' -Style 'Normal')) | Out-Null
    }
    elseif ($applicationInventoryValidated) {
        $validatedApplicationNarrativeSegments = New-Object System.Collections.Generic.List[string]
        if ($null -ne $ssoEnabledApps) {
            $validatedApplicationNarrativeSegments.Add("$ssoEnabledApps with SSO enabled") | Out-Null
        }
        if ($null -ne $highPrivilegeApps) {
            $validatedApplicationNarrativeSegments.Add("$highPrivilegeApps with elevated permissions") | Out-Null
        }
        if ($ownerSignalFullyValidated -and $null -ne $firstPartyAppsWithoutOwners -and $firstPartyAppsWithoutOwners -gt 0) {
            $validatedApplicationNarrativeSegments.Add("$firstPartyAppsWithoutOwners first-party app(s) without owner coverage") | Out-Null
        }
        if ($null -ne $thirdPartyAppsWithApplicationPerms -and $thirdPartyAppsWithApplicationPerms -gt 0) {
            $validatedApplicationNarrativeSegments.Add("$thirdPartyAppsWithApplicationPerms third-party app(s) with application permissions") | Out-Null
        }
        if ($redirectUriSignalFullyValidated -and $null -ne $appsWithRedirectUriRisk -and $appsWithRedirectUriRisk -gt 0) {
            $validatedApplicationNarrativeSegments.Add("$appsWithRedirectUriRisk redirect URI review flag(s)") | Out-Null
        }
        if ($activitySignalFullyValidated -and $null -ne $appsWithNoRecentActivity -and $appsWithNoRecentActivity -gt 0) {
            $validatedApplicationNarrativeSegments.Add("$appsWithNoRecentActivity with no recent activity signal") | Out-Null
        }

        $validatedApplicationNarrativeText = Join-ArrayaReadableList -Items @($validatedApplicationNarrativeSegments.ToArray())
        $applicationNarrative = "The application review looked at enterprise application inventory, sign-in posture, and the degree to which privileged permissions are visible in the current tenant data. The reviewed source surfaced {0} enterprise application(s)" -f $(if ($null -ne $totalEnterpriseApps) { $totalEnterpriseApps } else { 'an unconfirmed number of' })
        if (-not [string]::IsNullOrWhiteSpace($validatedApplicationNarrativeText)) {
            $applicationNarrative += ", including $validatedApplicationNarrativeText"
        }
        $applicationNarrative += '.'
        if (-not [string]::IsNullOrWhiteSpace($applicationValidationNoteText)) {
            $applicationNarrative += " $applicationValidationNoteText"
        }
        $blocks.Add((New-CustomerWordParagraphBlock -Text $applicationNarrative -Style 'Normal')) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current review did not surface a usable enterprise application inventory. Validation note: this section should be treated as incomplete rather than as proof that no enterprise applications exist in the tenant.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application Inventory' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Application', 'SSO Enabled', 'SSO Mode', 'Observation') -Rows $applicationInventoryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application Sign-In Activity and Security Analysis' -Style 'Heading3')) | Out-Null
    if ($applicationInventoryClearlyEmpty) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Because the current review did not surface enterprise application objects, this section does not point to a present app-governance concentration on its own. The next decision here is simply whether the observed zero-app result matches the tenant operating model or whether the inventory should be validated again.' -Style 'Normal')) | Out-Null
    }
    elseif (-not $applicationInventoryValidated) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Validation note: the current review did not surface enough enterprise application detail to make a confident governance call. Before using this section for a remediation decision, confirm whether the app inventory was intentionally out of scope or simply not visible in the reviewed data.' -Style 'Normal')) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Where privileged permissions are present and approval workflow maturity is not clearly surfaced, the tenant can accumulate enterprise applications that are difficult to rationalize quickly during a security review. The application inventory should therefore be reviewed in the same operating rhythm as privileged identities and guest access rather than as a separate governance track.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recommendations for Application Governance' -Style 'Heading3')) | Out-Null
    if ($applicationInventoryValidated -or $applicationInventoryClearlyEmpty) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current recommendation set points back to identity governance and least-privilege cleanup. If enterprise applications are expected in this tenant, confirm that the inventory and ownership model are complete before treating the current state as stable.' -Style 'Normal')) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current recommendation here is to validate the inventory first. Once the tenant surfaces a usable enterprise application list, the same identity-governance and least-privilege review rhythm can be applied to app ownership, consent, and permission scope.' -Style 'Normal')) | Out-Null
    }

    $mfaEnrollmentRate = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('RegistrationPercent'))
    $mfaRegisteredUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('RegisteredUsers'))
    $mfaNotRegisteredUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('NotRegisteredUsers'))
    $mfaTotalUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('TotalUsers'))
    $weakMethodsOnlyUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('UsersWithWeakMethodsOnly'))
    $weakDefaultMethodUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('UsersWithWeakDefaultMethod'))
    $phishingResistantMethodUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('UsersWithPhishingResistantMethods'))
    $mfaConditionalAccessPoliciesReviewed = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('ConditionalAccessPoliciesReviewed'))
    $enabledMfaEnforcementPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('EnabledPoliciesRequiringMfa'))
    $reportOnlyMfaEnforcementPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('ReportOnlyPoliciesRequiringMfa'))
    $mfaPoliciesWithExclusions = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('PoliciesWithExclusions'))
    $mfaEnabledUsersReviewed = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('EnabledUsersReviewed'))
    $mfaUsersCoveredByEnabledPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('UsersCoveredByEnabledMfaPolicies'))
    $mfaUserCoveragePercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('UserCoveragePercent'))
    $mfaEnabledMemberUsersReviewed = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('EnabledMemberUsersReviewed'))
    $mfaMemberUsersCoveredByEnabledPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('MemberUsersCoveredByEnabledMfaPolicies'))
    $mfaMemberUserCoveragePercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('MemberUserCoveragePercent'))
    $mfaEnabledGuestUsersReviewed = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('EnabledGuestUsersReviewed'))
    $mfaGuestUsersCoveredByEnabledPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('GuestUsersCoveredByEnabledMfaPolicies'))
    $mfaGuestUserCoveragePercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('GuestUserCoveragePercent'))
    $guestUserEnforcementStateText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('GuestUserEnforcementState')) -Default 'Not validated from the reviewed data'
    $mfaUsersNotCoveredByEnabledPolicies = if ($null -ne $mfaEnabledUsersReviewed -and $null -ne $mfaUsersCoveredByEnabledPolicies) { [math]::Max(($mfaEnabledUsersReviewed - $mfaUsersCoveredByEnabledPolicies), 0) } else { $null }
    $mfaCoverageCalculationNoteText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('CoverageCalculationNote')) -Default 'Not validated from the reviewed data'
    $caSummaryTotalPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('TotalPolicies'))
    $caSummaryPoliciesWithExclusions = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('PoliciesWithExclusions'))
    if ($null -ne $caSummaryTotalPolicies) { $mfaConditionalAccessPoliciesReviewed = $caSummaryTotalPolicies }
    if ($null -ne $caSummaryPoliciesWithExclusions) { $mfaPoliciesWithExclusions = $caSummaryPoliciesWithExclusions }
    $mfaEnforcementStateText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('EnforcementState')) -Default 'Not validated from the reviewed data'
    $registeredMethodMixText = Convert-ToCustomerAssessmentMfaMethodBreakdownText -Value (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('RegisteredMethodBreakdown')) -Default 'Not validated from the reviewed data'
    $weakMethodMixText = Convert-ToCustomerAssessmentMfaMethodBreakdownText -Value (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('WeakMethodBreakdown')) -Default 'Not validated from the reviewed data'
    $phishingResistantMethodMixText = Convert-ToCustomerAssessmentMfaMethodBreakdownText -Value (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('PhishingResistantMethodBreakdown')) -Default 'Not validated from the reviewed data'
    $securityDefaultsStateText = Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $securityDefaultsPolicyRecord -Names @('IsEnabled', 'Enabled')) -TrueText 'Enabled' -FalseText 'Disabled' -Default 'Not validated from the reviewed data'
    $guestCoverageStateText = Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('GuestOrExternalCoverage')) -TrueText 'Detected' -FalseText 'Not detected' -Default 'Not validated from the reviewed data'
    $privilegedCoverageStateText = Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('PrivilegedRoleCoverage')) -TrueText 'Detected' -FalseText 'Not detected' -Default 'Not validated from the reviewed data'
    $riskCoverageStateText = Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('RiskBasedCoverage')) -TrueText 'Detected' -FalseText 'Not detected' -Default 'Not validated from the reviewed data'
    $compliantDeviceRequirementStateText = Convert-ToCustomerAssessmentBooleanLabel -Value (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('CompliantDeviceRequirement')) -TrueText 'Detected' -FalseText 'Not detected' -Default 'Not validated from the reviewed data'
    $enabledAdminUsersReviewed = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $adminMfaSummaryRecord -Names @('EnabledAdminUsersReviewed'))
    $adminUsersRegisteredForMfa = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $adminMfaSummaryRecord -Names @('AdminUsersRegisteredForMfa'))
    $adminUsersNotRegisteredForMfa = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $adminMfaSummaryRecord -Names @('AdminUsersNotRegisteredForMfa'))
    $adminUsersCoveredByMfaEnforcement = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $adminMfaSummaryRecord -Names @('AdminUsersCoveredByMfaEnforcement'))
    $adminUsersNotCoveredByMfaEnforcement = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $adminMfaSummaryRecord -Names @('AdminUsersNotCoveredByMfaEnforcement'))
    $adminUserCoveragePercent = if ($null -ne $enabledAdminUsersReviewed -and $enabledAdminUsersReviewed -gt 0 -and $null -ne $adminUsersCoveredByMfaEnforcement) { [math]::Round(($adminUsersCoveredByMfaEnforcement / $enabledAdminUsersReviewed) * 100, 1) } else { $null }
    $mfaGapRowsSorted = @(
        $mfaEnforcementGapUserRows |
            Sort-Object `
                @{ Expression = {
                    switch ((Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory')) -Default '').ToLowerInvariant()) {
                        'excluded from enabled mfa ca policy' { 0 }
                        'outside enabled mfa ca include scope' { 1 }
                        default { 2 }
                    }
                } }, `
                @{ Expression = { (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed user').ToLowerInvariant() } }, `
                @{ Expression = { (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default '').ToLowerInvariant() } }
    )
    $mfaGapInternalMemberRowsSorted = @(
        $mfaGapRowsSorted |
            Where-Object {
                $userTypeText = (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserType')) -Default '').ToLowerInvariant()
                $userPrincipalName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default ''
                $isGuestIdentity = ($userTypeText -eq 'guest' -or $userPrincipalName -like '*#EXT#*')
                -not $isGuestIdentity
            }
    )
    $mfaScopeRowsSorted = @(
        $mfaEnforcementScopeReviewRows |
            Sort-Object `
                @{ Expression = { (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('PolicyName')) -Default '').ToLowerInvariant() } }, `
                @{ Expression = {
                    switch ((Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ScopeType')) -Default '').ToLowerInvariant()) {
                        'exclude' { 0 }
                        'include' { 1 }
                        default { 2 }
                    }
                } }, `
                @{ Expression = { (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ObjectType')) -Default '').ToLowerInvariant() } }, `
                @{ Expression = { (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName', 'Identifier')) -Default '').ToLowerInvariant() } }
    )
    $mfaExcludedUserCount = @($mfaGapRowsSorted | Where-Object {
        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory')) -Default '').ToLowerInvariant() -eq 'excluded from enabled mfa ca policy'
    }).Count
    $mfaOutsideIncludeUserCount = @($mfaGapRowsSorted | Where-Object {
        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory')) -Default '').ToLowerInvariant() -eq 'outside enabled mfa ca include scope'
    }).Count
    $mfaUncoveredMemberCount = @($mfaGapRowsSorted | Where-Object {
        $userTypeText = (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserType')) -Default '').ToLowerInvariant()
        $userPrincipalName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default ''
        $isGuestIdentity = ($userTypeText -eq 'guest' -or $userPrincipalName -like '*#EXT#*')
        -not $isGuestIdentity
    }).Count
    $mfaUncoveredGuestCount = @($mfaGapRowsSorted | Where-Object {
        $userTypeText = (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserType')) -Default '').ToLowerInvariant()
        $userPrincipalName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default ''
        ($userTypeText -eq 'guest' -or $userPrincipalName -like '*#EXT#*')
    }).Count

    $mfaReviewedPopulationWasDerived = $false
    $mfaCoverageWasEstimated = $false
    $adminCoverageWasEstimated = $false
    $adminRegistrationWasEstimated = $false

    if ($null -eq $mfaEnabledUsersReviewed -and $enabledUsers.Count -gt 0) {
        $mfaEnabledUsersReviewed = $enabledUsers.Count
        $mfaReviewedPopulationWasDerived = $true
    }
    if ($null -eq $mfaEnabledMemberUsersReviewed -and $enabledMemberUsers.Count -gt 0) {
        $mfaEnabledMemberUsersReviewed = $enabledMemberUsers.Count
        $mfaReviewedPopulationWasDerived = $true
    }
    if ($null -eq $mfaEnabledGuestUsersReviewed -and $enabledGuestUsers.Count -gt 0) {
        $mfaEnabledGuestUsersReviewed = $enabledGuestUsers.Count
        $mfaReviewedPopulationWasDerived = $true
    }

    if ($mfaGapRowsSorted.Count -gt 0) {
        if ($null -eq $mfaUsersCoveredByEnabledPolicies -and $null -ne $mfaEnabledUsersReviewed) {
            $mfaUsersCoveredByEnabledPolicies = [math]::Max(($mfaEnabledUsersReviewed - $mfaGapRowsSorted.Count), 0)
            $mfaCoverageWasEstimated = $true
        }
        if ($null -eq $mfaMemberUsersCoveredByEnabledPolicies -and $null -ne $mfaEnabledMemberUsersReviewed) {
            $mfaMemberUsersCoveredByEnabledPolicies = [math]::Max(($mfaEnabledMemberUsersReviewed - $mfaUncoveredMemberCount), 0)
            $mfaCoverageWasEstimated = $true
        }
        if ($null -eq $mfaGuestUsersCoveredByEnabledPolicies -and $null -ne $mfaEnabledGuestUsersReviewed) {
            $mfaGuestUsersCoveredByEnabledPolicies = [math]::Max(($mfaEnabledGuestUsersReviewed - $mfaUncoveredGuestCount), 0)
            $mfaCoverageWasEstimated = $true
        }
    }

    if ($null -eq $mfaUsersNotCoveredByEnabledPolicies -and $null -ne $mfaEnabledUsersReviewed -and $null -ne $mfaUsersCoveredByEnabledPolicies) {
        $mfaUsersNotCoveredByEnabledPolicies = [math]::Max(($mfaEnabledUsersReviewed - $mfaUsersCoveredByEnabledPolicies), 0)
    }
    if ($null -eq $mfaUserCoveragePercent -and $null -ne $mfaEnabledUsersReviewed -and $mfaEnabledUsersReviewed -gt 0 -and $null -ne $mfaUsersCoveredByEnabledPolicies) {
        $mfaUserCoveragePercent = [math]::Round(($mfaUsersCoveredByEnabledPolicies / $mfaEnabledUsersReviewed) * 100, 1)
        $mfaCoverageWasEstimated = $true
    }
    if ($null -eq $mfaMemberUserCoveragePercent -and $null -ne $mfaEnabledMemberUsersReviewed -and $mfaEnabledMemberUsersReviewed -gt 0 -and $null -ne $mfaMemberUsersCoveredByEnabledPolicies) {
        $mfaMemberUserCoveragePercent = [math]::Round(($mfaMemberUsersCoveredByEnabledPolicies / $mfaEnabledMemberUsersReviewed) * 100, 1)
        $mfaCoverageWasEstimated = $true
    }
    if ($null -eq $mfaGuestUserCoveragePercent -and $null -ne $mfaEnabledGuestUsersReviewed -and $mfaEnabledGuestUsersReviewed -gt 0 -and $null -ne $mfaGuestUsersCoveredByEnabledPolicies) {
        $mfaGuestUserCoveragePercent = [math]::Round(($mfaGuestUsersCoveredByEnabledPolicies / $mfaEnabledGuestUsersReviewed) * 100, 1)
        $mfaCoverageWasEstimated = $true
    }

    if ($null -eq $enabledAdminUsersReviewed -and $enabledAdminRows.Count -gt 0) {
        $enabledAdminUsersReviewed = $enabledAdminRows.Count
        $adminCoverageWasEstimated = $true
    }
    if ($null -eq $adminUsersNotCoveredByMfaEnforcement -and $adminMfaEnforcementGapRows.Count -gt 0) {
        $adminUsersNotCoveredByMfaEnforcement = $adminMfaEnforcementGapRows.Count
        $adminCoverageWasEstimated = $true
    }
    if ($null -eq $adminUsersCoveredByMfaEnforcement -and $null -ne $enabledAdminUsersReviewed -and $null -ne $adminUsersNotCoveredByMfaEnforcement) {
        $adminUsersCoveredByMfaEnforcement = [math]::Max(($enabledAdminUsersReviewed - $adminUsersNotCoveredByMfaEnforcement), 0)
        $adminCoverageWasEstimated = $true
    }
    if ($null -eq $adminUserCoveragePercent -and $null -ne $enabledAdminUsersReviewed -and $enabledAdminUsersReviewed -gt 0 -and $null -ne $adminUsersCoveredByMfaEnforcement) {
        $adminUserCoveragePercent = [math]::Round(($adminUsersCoveredByMfaEnforcement / $enabledAdminUsersReviewed) * 100, 1)
        $adminCoverageWasEstimated = $true
    }

    $mfaRegistrationByUpn = @{}
    foreach ($registrationRow in @($mfaRegistrationDetailRows)) {
        $upn = (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $registrationRow -Names @('UserPrincipalName')) -Default '').Trim().ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($upn) -and -not $mfaRegistrationByUpn.ContainsKey($upn)) {
            $mfaRegistrationByUpn[$upn] = $registrationRow
        }
    }
    if (($null -eq $adminUsersRegisteredForMfa -or $null -eq $adminUsersNotRegisteredForMfa) -and $enabledAdminRows.Count -gt 0 -and $mfaRegistrationByUpn.Count -gt 0) {
        $matchedAdminRegistrationRows = @(
            $enabledAdminRows |
                ForEach-Object {
                    $adminUpn = (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default '').Trim().ToLowerInvariant()
                    if (-not [string]::IsNullOrWhiteSpace($adminUpn) -and $mfaRegistrationByUpn.ContainsKey($adminUpn)) {
                        $mfaRegistrationByUpn[$adminUpn]
                    }
                } |
                Where-Object { $null -ne $_ }
        )
        if ($matchedAdminRegistrationRows.Count -eq $enabledAdminRows.Count) {
            if ($null -eq $adminUsersRegisteredForMfa) {
                $adminUsersRegisteredForMfa = @($matchedAdminRegistrationRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsMfaRegistered'))) -eq $true }).Count
                $adminRegistrationWasEstimated = $true
            }
            if ($null -eq $adminUsersNotRegisteredForMfa) {
                $adminUsersNotRegisteredForMfa = @($matchedAdminRegistrationRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsMfaRegistered'))) -ne $true }).Count
                $adminRegistrationWasEstimated = $true
            }
        }
    }
    if ($null -eq $adminUsersNotRegisteredForMfa -and $adminMfaRegistrationGapRows.Count -gt 0) {
        $adminUsersNotRegisteredForMfa = $adminMfaRegistrationGapRows.Count
        $adminRegistrationWasEstimated = $true
    }
    if ($null -eq $adminUsersRegisteredForMfa -and $null -ne $enabledAdminUsersReviewed -and $null -ne $adminUsersNotRegisteredForMfa) {
        $adminUsersRegisteredForMfa = [math]::Max(($enabledAdminUsersReviewed - $adminUsersNotRegisteredForMfa), 0)
        $adminRegistrationWasEstimated = $true
    }

    if ($mfaCoverageWasEstimated) {
        $mfaCoverageCalculationNoteText = 'Coverage counts are estimated from the enabled account inventory and surfaced uncovered-user gap rows in the reviewed data.'
    }
    elseif ($mfaReviewedPopulationWasDerived -and [string]::IsNullOrWhiteSpace(($mfaCoverageCalculationNoteText -replace 'Not validated from the reviewed data', '').Trim())) {
        $mfaCoverageCalculationNoteText = 'Enabled-user review counts are derived from the enabled account inventory. The current source did not surface a complete MFA coverage count.'
    }
    $mfaGapSummaryRows = if ($mfaGapRowsSorted.Count -gt 0) {
        @(
            New-CustomerWordTableRow -Cells @('Users outside enabled MFA CA include scope', $mfaOutsideIncludeUserCount)
            New-CustomerWordTableRow -Cells @('Users explicitly excluded from enabled MFA CA policies', $mfaExcludedUserCount)
            New-CustomerWordTableRow -Cells @('Member users not covered', $mfaUncoveredMemberCount)
            New-CustomerWordTableRow -Cells @('Guest users not covered', $mfaUncoveredGuestCount)
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not validated from the reviewed data', 'Not validated from the reviewed data')))
    }
    $mfaGapDetailRows = if ($mfaGapInternalMemberRowsSorted.Count -gt 0) {
        @(
            $mfaGapInternalMemberRowsSorted |
                Select-Object -First 15 |
                ForEach-Object {
                    $relatedPolicyText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('RelatedPolicies')) -Default ''
                    if ([string]::IsNullOrWhiteSpace($relatedPolicyText) -or $relatedPolicyText -eq 'Not validated from the reviewed data') {
                        $relatedPolicyText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('GapReason')) -Default 'Not validated from the reviewed data'
                    }
                    $userPrincipalName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default 'Not validated from the reviewed data'
                    $userTypeText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserType')) -Default ''
                    $displayUserType = if ($userTypeText.ToLowerInvariant() -eq 'guest' -or $userPrincipalName -like '*#EXT#*') { 'Guest' } else { $(if ([string]::IsNullOrWhiteSpace($userTypeText)) { 'Member' } else { $userTypeText }) }

                    New-CustomerWordTableRow -Cells @(
                        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed user'),
                        $userPrincipalName,
                        $displayUserType,
                        (Get-CustomerMfaGapCategoryLabel -Value (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory'))),
                        $relatedPolicyText
                    )
                }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('No uncovered internal member users surfaced in the reviewed data', 'N/A', 'Member', 'N/A', 'Guest and external-user gaps remain summarized separately in this report and in the workbook tabs.')))
    }
    $guestScopeKeywordPattern = '(?i)guest|external|b2b|partner'
    $guestRelevantMfaScopeRows = @(
        $mfaScopeRowsSorted | Where-Object {
            $scopeCompositeText = @(
                (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('PolicyName')) -Default ''),
                (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ObjectType')) -Default ''),
                (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName', 'Identifier')) -Default ''),
                (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('Notes')) -Default '')
            ) -join ' '
            $scopeCompositeText -match $guestScopeKeywordPattern
        }
    )
    $guestRelevantExcludeCount = @($guestRelevantMfaScopeRows | Where-Object {
        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ScopeType')) -Default '').ToLowerInvariant() -eq 'exclude'
    }).Count
    $guestRelevantIncludeCount = @($guestRelevantMfaScopeRows | Where-Object {
        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ScopeType')) -Default '').ToLowerInvariant() -eq 'include'
    }).Count
    $mfaScopeDriverText = if ($mfaUncoveredGuestCount -le 0 -and $mfaGapRowsSorted.Count -gt 0) {
        'The reviewed enabled guest population appears covered by the active MFA baseline, so the current guest story is more about maintaining that design than closing a large enforcement gap.'
    }
    elseif ($mfaUncoveredGuestCount -le 0) {
        'The reviewed data did not surface a large guest-specific MFA gap, but guest coverage should still be validated alongside cross-tenant trust and invitation controls.'
    }
    elseif ($guestRelevantMfaScopeRows.Count -eq 0 -or $guestCoverageStateText -match '(?i)^not clearly detected|^not validated') {
        'The guest MFA gap appears to be driven mainly by the absence of a clearly guest-specific enforcement pattern in the reviewed baseline rather than by one isolated exclusion.'
    }
    elseif ($guestRelevantExcludeCount -gt 0) {
        'The guest MFA gap appears to be driven by a mix of explicit guest or external exclusions and guests falling outside the active include scope.'
    }
    else {
        'The guest MFA gap appears to be driven mostly by guests sitting outside the active include scope rather than by a large number of explicit guest exclusions.'
    }
    $mfaScopeInterpretationText = 'Some Conditional Access policies naturally target employee, admin, or workload-specific populations, so a long raw include and exclude list is not the best way to understand guest risk. The summary below focuses on what is actually shaping guest coverage, while the workbook tab MfaEnforcementScopeReview keeps the full policy detail for implementation planning.'
    $mfaScopeExampleSourceRows = @(
        $mfaScopeRowsSorted |
            Sort-Object `
                @{ Expression = {
                    $scopeTypeText = (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ScopeType')) -Default '').ToLowerInvariant()
                    $scopeCompositeText = @(
                        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('PolicyName')) -Default ''),
                        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ObjectType')) -Default ''),
                        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName', 'Identifier')) -Default ''),
                        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('Notes')) -Default '')
                    ) -join ' '
                    if ($scopeCompositeText -match $guestScopeKeywordPattern -and $scopeTypeText -eq 'exclude') { 0 }
                    elseif ($scopeCompositeText -match $guestScopeKeywordPattern) { 1 }
                    elseif ($scopeTypeText -eq 'exclude') { 2 }
                    else { 3 }
                } }, `
                @{ Expression = {
                    $affectedUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('AffectedEnabledUsers'))
                    if ($null -eq $affectedUsers) { -1 } else { -1 * $affectedUsers }
                } }, `
                @{ Expression = { (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('PolicyName')) -Default '').ToLowerInvariant() } }, `
                @{ Expression = { (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName', 'Identifier')) -Default '').ToLowerInvariant() } }
    )
    $mfaScopeExampleRows = if ($mfaScopeExampleSourceRows.Count -gt 0) {
        @(
            $mfaScopeExampleSourceRows |
                Select-Object -First 6 |
                ForEach-Object {
                    $policyNameText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('PolicyName')) -Default 'Not validated from the reviewed data'
                    $scopeTypeText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ScopeType')) -Default 'Not validated from the reviewed data'
                    $objectTypeText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ObjectType')) -Default ''
                    $displayNameText = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName', 'Identifier')) -Default 'Not validated from the reviewed data'
                    $scopeSignalText = if (-not [string]::IsNullOrWhiteSpace($objectTypeText) -and $objectTypeText -ne 'Not validated from the reviewed data') {
                        '{0}: {1}' -f $objectTypeText, $displayNameText
                    }
                    else {
                        $displayNameText
                    }
                    $scopeCompositeText = @($policyNameText, $objectTypeText, $displayNameText, (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('Notes')) -Default '')) -join ' '
                    $observationText = if ($scopeCompositeText -match $guestScopeKeywordPattern -and $scopeTypeText.ToLowerInvariant() -eq 'exclude') {
                        'This example shows an explicit guest or external exclusion path that can reduce guest MFA coverage.'
                    }
                    elseif ($scopeCompositeText -match $guestScopeKeywordPattern) {
                        'This example shows guest or external access being targeted directly in the MFA policy scope.'
                    }
                    elseif ($scopeTypeText.ToLowerInvariant() -eq 'exclude') {
                        'This example shows an exclusion path that matters if that excluded group overlaps guest access or guest sponsors.'
                    }
                    else {
                        'This example shows one of the groups shaping the active MFA include scope for the current baseline.'
                    }

                    New-CustomerWordTableRow -Cells @(
                        $policyNameText,
                        $scopeTypeText,
                        $scopeSignalText,
                        $observationText
                    )
                }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not validated from the reviewed data', 'Not validated from the reviewed data', 'Not validated from the reviewed data', 'Not validated from the reviewed data')))
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text '6.0 Authentication Methods, MFA Enrollment, and MFA Enforcement' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This section separates MFA enrollment from MFA enforcement. Enrollment shows which authentication methods users have registered and whether weaker methods remain in use. Enforcement shows whether active access controls are actually requiring MFA during sign-in.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'MFA Enrollment' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows @(
        @('Users reviewed', $(if ($null -ne $mfaTotalUsers) { $mfaTotalUsers } else { 'Not validated from the reviewed data' })),
        @('Registered for MFA', $(if ($null -ne $mfaRegisteredUsers) { $mfaRegisteredUsers } else { 'Not validated from the reviewed data' })),
        @('Not registered for MFA', $(if ($null -ne $mfaNotRegisteredUsers) { $mfaNotRegisteredUsers } else { 'Not validated from the reviewed data' })),
        @('MFA enrollment rate', $(if ($null -ne $mfaEnrollmentRate) { "$mfaEnrollmentRate%" } else { 'Not validated from the reviewed data' })),
        @('Registered method mix', $registeredMethodMixText),
        @('Weak methods observed', $weakMethodMixText),
        @('Users with weak MFA methods only', $(if ($null -ne $weakMethodsOnlyUsers) { $weakMethodsOnlyUsers } else { 'Not validated from the reviewed data' })),
        @('Users with weak default MFA method', $(if ($null -ne $weakDefaultMethodUsers) { $weakDefaultMethodUsers } else { 'Not validated from the reviewed data' })),
        @('Users with phishing-resistant methods', $(if ($null -ne $phishingResistantMethodUsers) { $phishingResistantMethodUsers } else { 'Not validated from the reviewed data' })),
        @('Phishing-resistant method mix', $phishingResistantMethodMixText)
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Enrollment shows readiness, not enforcement. In the reviewed data, MFA enrollment is {0}, and {1} registered user(s) currently rely only on weaker methods while {2} default to a weaker MFA method. That means the tenant should evaluate method quality alongside raw registration coverage before treating MFA enrollment as mature." -f $(if ($null -ne $mfaEnrollmentRate) { "$mfaEnrollmentRate%" } else { 'not clearly validated' }), $(if ($null -eq $weakMethodsOnlyUsers) { 'an unconfirmed number of' } else { $weakMethodsOnlyUsers }), $(if ($null -eq $weakDefaultMethodUsers) { 'an unconfirmed number of users' } else { $weakDefaultMethodUsers })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'MFA Enforcement' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows @(
        @('Conditional Access policies reviewed', $(if ($null -ne $mfaConditionalAccessPoliciesReviewed) { $mfaConditionalAccessPoliciesReviewed } else { (Get-CustomerObservationState -Observation $identityObservation -Signal 'Conditional Access policies') })),
        @('Enabled MFA enforcement policies', $(if ($null -ne $enabledMfaEnforcementPolicies) { $enabledMfaEnforcementPolicies } else { 'Not validated from the reviewed data' })),
        @('Report-only MFA enforcement policies', $(if ($null -ne $reportOnlyMfaEnforcementPolicies) { $reportOnlyMfaEnforcementPolicies } else { 'Not validated from the reviewed data' })),
        @('Policies with exclusions', $(if ($null -ne $mfaPoliciesWithExclusions) { $mfaPoliciesWithExclusions } else { (Get-CustomerObservationState -Observation $identityObservation -Signal 'Policies with exclusions') })),
        @('Enabled users reviewed', $(if ($null -ne $mfaEnabledUsersReviewed) { $mfaEnabledUsersReviewed } else { 'Not validated from the reviewed data' })),
        @('Users covered by enabled MFA enforcement policies', $(if ($null -ne $mfaUsersCoveredByEnabledPolicies) { $mfaUsersCoveredByEnabledPolicies } else { 'Not validated from the reviewed data' })),
        @('Users not covered by enabled MFA enforcement policies', $(if ($null -ne $mfaUsersNotCoveredByEnabledPolicies) { $mfaUsersNotCoveredByEnabledPolicies } else { 'Not validated from the reviewed data' })),
        @('Estimated enabled-user CA MFA coverage', $(if ($null -ne $mfaUserCoveragePercent) { "$mfaUserCoveragePercent%" } else { 'Not validated from the reviewed data' })),
        @('Enabled member users reviewed', $(if ($null -ne $mfaEnabledMemberUsersReviewed) { $mfaEnabledMemberUsersReviewed } else { 'Not validated from the reviewed data' })),
        @('Member users covered by enabled MFA enforcement policies', $(if ($null -ne $mfaMemberUsersCoveredByEnabledPolicies) { $mfaMemberUsersCoveredByEnabledPolicies } else { 'Not validated from the reviewed data' })),
        @('Estimated member-user CA MFA coverage', $(if ($null -ne $mfaMemberUserCoveragePercent) { "$mfaMemberUserCoveragePercent%" } else { 'Not validated from the reviewed data' })),
        @('Enabled guest users reviewed', $(if ($null -ne $mfaEnabledGuestUsersReviewed) { $mfaEnabledGuestUsersReviewed } else { 'Not validated from the reviewed data' })),
        @('Guest users covered by enabled MFA enforcement policies', $(if ($null -ne $mfaGuestUsersCoveredByEnabledPolicies) { $mfaGuestUsersCoveredByEnabledPolicies } else { 'Not validated from the reviewed data' })),
        @('Estimated guest-user CA MFA coverage', $(if ($null -ne $mfaGuestUserCoveragePercent) { "$mfaGuestUserCoveragePercent%" } else { 'Not validated from the reviewed data' })),
        @('Guest-user MFA enforcement summary', $guestUserEnforcementStateText),
        @('Enabled admin users reviewed', $(if ($null -ne $enabledAdminUsersReviewed) { $enabledAdminUsersReviewed } else { 'Not validated from the reviewed data' })),
        @('Admin users registered for MFA', $(if ($null -ne $adminUsersRegisteredForMfa) { $adminUsersRegisteredForMfa } else { 'Not validated from the reviewed data' })),
        @('Admin users not registered for MFA', $(if ($null -ne $adminUsersNotRegisteredForMfa) { $adminUsersNotRegisteredForMfa } else { 'Not validated from the reviewed data' })),
        @('Admin users covered by enabled MFA enforcement policies', $(if ($null -ne $adminUsersCoveredByMfaEnforcement) { $adminUsersCoveredByMfaEnforcement } else { 'Not validated from the reviewed data' })),
        @('Admin users not covered by enabled MFA enforcement policies', $(if ($null -ne $adminUsersNotCoveredByMfaEnforcement) { $adminUsersNotCoveredByMfaEnforcement } else { 'Not validated from the reviewed data' })),
        @('Estimated admin-user CA MFA coverage', $(if ($null -ne $adminUserCoveragePercent) { "$adminUserCoveragePercent%" } else { 'Not validated from the reviewed data' })),
        @('Guest / external-user Conditional Access coverage', $guestCoverageStateText),
        @('Privileged-role Conditional Access coverage', $privilegedCoverageStateText),
        @('Risk-based Conditional Access coverage', $riskCoverageStateText),
        @('Compliant-device requirement in summary', $compliantDeviceRequirementStateText),
        @('Security Defaults policy', $securityDefaultsStateText),
        @('MFA enforcement state', $mfaEnforcementStateText),
        @('Coverage calculation note', $mfaCoverageCalculationNoteText)
    ))) | Out-Null
    $mfaCoverageNarrativeText = if ($null -ne $mfaUsersCoveredByEnabledPolicies -and $null -ne $mfaEnabledUsersReviewed -and $null -ne $mfaUserCoveragePercent) {
        "Based on the reviewed enabled-user inventory, those enabled MFA policies appear to cover $mfaUsersCoveredByEnabledPolicies of $mfaEnabledUsersReviewed enabled reviewed user(s), or $mfaUserCoveragePercent%."
    }
    elseif ($null -ne $mfaEnabledUsersReviewed) {
        "The reviewed enabled-user inventory included $mfaEnabledUsersReviewed enabled account(s), but the current source did not surface a complete covered-user total."
    }
    else {
        'The current source did not surface enough enabled-user detail to calculate a reliable MFA coverage total.'
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Enforcement shows whether users are actually being required to perform MFA, not just whether they have registered methods. In this review, the policy baseline shows {0} enabled Conditional Access policy/policies that require MFA and {1} still in report-only mode. {2} Report-only policies do not count as enforced coverage. Detailed uncovered-user and policy-scope review rows are available in the workbook tabs MfaEnforcementGapUsers and MfaEnforcementScopeReview." -f $(if ($null -eq $enabledMfaEnforcementPolicies) { 'an unconfirmed number of' } else { $enabledMfaEnforcementPolicies }), $(if ($null -eq $reportOnlyMfaEnforcementPolicies) { 'an unconfirmed number of policies' } else { $reportOnlyMfaEnforcementPolicies }), $mfaCoverageNarrativeText) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Guest MFA should be read as its own design question, not as a side effect of employee-only policies. {0}" -f $guestUserEnforcementStateText) -Style 'Normal')) | Out-Null
    $adminRegistrationNarrativeText = if ($null -ne $adminUsersNotRegisteredForMfa) {
        "$adminUsersNotRegisteredForMfa admin account(s) are not registered for MFA"
    }
    else {
        'the current source did not confirm how many admin accounts are not registered for MFA'
    }
    $adminCoverageNarrativeText = if ($null -ne $adminUsersNotCoveredByMfaEnforcement) {
        "$adminUsersNotCoveredByMfaEnforcement are not covered by the active MFA enforcement baseline"
    }
    else {
        'the current source did not confirm how many admin accounts are outside the active MFA enforcement baseline'
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Privileged access should also be reviewed separately from the general population. In the reviewed enabled-admin set, {0}, and {1}. Detailed admin gap rows are available in the workbook tabs AdminMfaRegistrationGaps and AdminMfaEnforcementGaps." -f $adminRegistrationNarrativeText, $adminCoverageNarrativeText) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if (-not [string]::IsNullOrWhiteSpace($guestMfaUserExperienceText)) { $guestMfaUserExperienceText } else { 'Guest-user experience should be reviewed separately from enrollment counts because strong guest authentication can often rely on the guest home-tenant MFA rather than a separate registration in this tenant.' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if (-not [string]::IsNullOrWhiteSpace($guestMfaDesiredStateText)) { $guestMfaDesiredStateText } else { 'The desired baseline is strong guest authentication with trusted home-tenant MFA where supported and approved, not a blanket requirement for every guest to register separately in the resource tenant.' })) -Style 'Normal')) | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($guestMfaRecommendationStrategyText)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text $guestMfaRecommendationStrategyText -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'MFA Enforcement Gap Summary' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Coverage Gap Signal', 'Current State') -Rows $mfaGapSummaryRows)) | Out-Null
    $mfaGapNarrativeText = if ($null -ne $mfaUsersNotCoveredByEnabledPolicies) {
        "Of the $mfaUsersNotCoveredByEnabledPolicies enabled reviewed user(s) not currently covered by active MFA enforcement, $mfaExcludedUserCount appear explicitly excluded while $mfaOutsideIncludeUserCount fall outside the active include scope."
    }
    elseif ($mfaGapRowsSorted.Count -gt 0) {
        "The current source surfaced $($mfaGapRowsSorted.Count) uncovered reviewed user(s); $mfaExcludedUserCount appear explicitly excluded while $mfaOutsideIncludeUserCount fall outside the active include scope."
    }
    else {
        'The current source did not surface uncovered-user totals, so the table below should be treated as representative evidence rather than a complete uncovered-user census.'
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0} The table below highlights the first {1} uncovered internal member identities so the team can see whether the current employee-focused gap is being driven by exclusions, narrow targeting, or both." -f $mfaGapNarrativeText, [math]::Min($mfaGapInternalMemberRowsSorted.Count, 15)) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Top Internal Member Users Not Covered by Enabled MFA Enforcement' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Display Name', 'User Principal Name', 'User Type', 'Gap Category', 'Related Policy / Scope') -Rows $mfaGapDetailRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Guest MFA Coverage Drivers' -Style 'Heading3')) | Out-Null
    $guestCoverageInterpretationText = if ($null -ne $mfaGuestUsersCoveredByEnabledPolicies) {
        "Guest users currently show $mfaGuestUsersCoveredByEnabledPolicies covered and $mfaUncoveredGuestCount uncovered account(s) in the reviewed enabled-user inventory."
    }
    elseif ($null -ne $mfaEnabledGuestUsersReviewed) {
        "The reviewed data included $mfaEnabledGuestUsersReviewed enabled guest account(s), but a complete guest covered-count total was not surfaced."
    }
    else {
        'The reviewed data did not surface a complete guest MFA coverage count.'
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0} {1} There are {2} guest or external-focused include row(s) and {3} guest or external-focused exclusion row(s) surfaced in the reviewed policy scope. {4}" -f $guestCoverageInterpretationText, $mfaScopeDriverText, $guestRelevantIncludeCount, $guestRelevantExcludeCount, $mfaScopeInterpretationText) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Representative MFA Scope Examples' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Policy', 'Scope Type', 'Scope Signal', 'Why It Matters') -Rows $mfaScopeExampleRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $identityConsultativeSummary) { $identityConsultativeSummary.RecommendationSupport } else { 'This section supports the identity and access recommendations in 4.0.' })) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '7.0 Password Writeback and Self-Service Password Reset' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows @(
        @('Password writeback', $passwordWritebackState),
        @('Self-service password reset', $selfServicePasswordResetState),
        @('Pass-through authentication', $passThroughAuthenticationState),
        @('Directory synchronization', $directorySynchronizationState),
        @('On-premises last sync', $onPremisesLastSyncState),
        @('Authentication configuration visibility', $authenticationVisibilityState)
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The hybrid identity portion of the review shows self-service password reset as {0} and directory synchronization as {1}. Password writeback is currently {2}, while pass-through authentication is shown as {3}. This makes it possible to separate settings that are clearly visible from settings that still need additional validation before they are treated as design decisions." -f $selfServicePasswordResetState, $directorySynchronizationState, $passwordWritebackState, $passThroughAuthenticationState) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The last surfaced on-premises sync timestamp is {0}. Validation note: where a password-lifecycle field is shown as not validated in current auth mode or not validated from the reviewed data, that should be read as a collection limitation rather than as proof that the tenant is misconfigured." -f $onPremisesLastSyncState) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '8.0 Authorization: Admin Access and Role Assignments' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Administrative access stood out in this tenant because the privileged footprint is large enough that stale or excess assignments create more than a theoretical risk. The review showed both the number of Global Administrators and the age of some privileged identities, which is usually a sign that role cleanup has not kept pace with operational change.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Summary of Findings' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows @(
        @('Total admin accounts reviewed', $(if ($adminRows.Count -gt 0) { $adminRows.Count } else { 'Not surfaced in current source' })),
        @('Global Administrators', (Get-CustomerObservationState -Observation $identityObservation -Signal 'Global Administrator count')),
        @('Stale privileged admins (>180 days)', (Get-CustomerObservationState -Observation $identityObservation -Signal 'Stale privileged admins (>180 days)')),
        @('Admin users not registered for MFA', $(if ($null -ne $adminUsersNotRegisteredForMfa) { $adminUsersNotRegisteredForMfa } else { 'Not validated from the reviewed data' })),
        @('Admin users not covered by enabled MFA enforcement policies', $(if ($null -ne $adminUsersNotCoveredByMfaEnforcement) { $adminUsersNotCoveredByMfaEnforcement } else { 'Not validated from the reviewed data' })),
        @('Example stale privileged identities', (Get-CustomerObservationState -Observation $identityObservation -Signal 'Example stale privileged identities'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Why This Matters' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $identityConsultativeSummary) { $identityConsultativeSummary.Narrative } else { 'Privileged access review matters most where stale administrative access and broad standing privileges begin to accumulate together.' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $identityConsultativeSummary) { $identityConsultativeSummary.RecommendationSupport } else { 'This section supports the identity and access recommendations in 4.0.' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Service Accounts' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current review did not surface a dedicated service-account inventory for privileged roles, but mailbox-based SMTP activity and application-related send patterns still provide useful clues about long-lived non-user access. Where service identities remain active, they should be reviewed with the same ownership and lifecycle discipline applied to privileged users.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('DisplayName', 'UserPrincipalName', 'Send Count', 'Last Activity') -Rows $smtpRelayServiceAccountRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Global Administrator Accounts' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('DisplayName', 'Created', 'UserPrincipalName', 'LastSignIn') -Rows $globalAdminTableRows)) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '9.0 Exchange Online: Mailboxes and Storage Overview' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'A review was performed on mailbox usage and the distribution of recipient objects. The primary goals of this assessment were to evaluate current storage patterns, identify resources that are no longer active, and document where mailbox lifecycle governance is becoming difficult to manage cleanly.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '9.1 High-Level Observations and Recommendations' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Messaging Snapshot At A Glance' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Signal', 'Current State') -Rows $messagingSnapshotRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Mailbox Lifecycle Summary' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Signal', 'Current State') -Rows $mailboxLifecycleRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Transport Exposure Summary' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Signal', 'Current State') -Rows $transportExposureRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $(if ($null -ne $messagingConsultativeSummary) { $messagingConsultativeSummary.Narrative } else { $messagingObservation.WhyItMatters })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $messagingConsultativeSummary) { $messagingConsultativeSummary.RecommendationSupport } else { 'This section supports the messaging and anti-spoofing recommendations in 4.0.' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'User Mailbox Growth' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $userMailboxGrowthRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("User mailbox growth does not stand out here because of a single oversized mailbox. It stands out because the tenant has {0} user mailbox statistic row(s) in scope, with the largest mailbox currently at {1}. When the mailbox footprint is reviewed alongside archive adoption, it becomes easier to see whether storage growth is being managed intentionally or simply carried forward by default." -f $userMailboxStatsRows.Count, $(if ($largestUserMailbox.Count -gt 0) { (& $formatSizeGb (Get-ArrayaObjectValue -Object $largestUserMailbox[0] -Names @('TotalItemSizeBytes'))) + ' GB' } else { 'Not surfaced in current source' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Shared Mailbox Review' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $sharedMailboxReviewRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Shared mailbox governance is a more visible pressure point in this tenant. The current source shows {0} shared mailbox statistic row(s), and the governance summary still reflects shared mailboxes without a clear ownership signal. That combination is notable because shared mailboxes tend to accumulate business-critical content long after day-to-day accountability becomes less obvious." -f $sharedMailboxStatsRows.Count) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Inactive Mailboxes' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $inactiveMailboxSummaryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Inactive mailboxes remain part of the tenant storage and lifecycle picture even when they are no longer part of routine operations. In this tenant, the inactive mailbox count is meaningful because it shows that mailbox retention and deprovisioning decisions continue to affect storage, compliance, and post-departure data handling.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Group Mailbox Utilization' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $groupMailboxUtilizationRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Group mailboxes are not carrying the same volume as the larger shared-mailbox footprint, but they are still a useful signal for collaboration hygiene. Where group mailboxes remain active while Teams and Microsoft 365 groups show ownership or dormancy drift, the collaboration story becomes harder to manage consistently.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Archive Mailbox Usage and Licensing' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $archiveMailboxUsageRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Archive usage does not appear uniformly mature across the mailbox population. The archive count, over-50GB archive indicator, and hold coverage together help show whether long-term email retention is being managed through a deliberate policy model or through mailbox-by-mailbox exceptions.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Mailbox Statistics Overview' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Mailbox Type', 'Quantity', 'Total Size (GB)', 'Total Archive Size (GB)', 'Average Size (GB)', 'Average Archive Size (GB)') -Rows $mailboxStatisticsRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The tenant review counted {0} recipients in scope, with the heaviest concentration in {1}. Domain concentration is visible in {2}, which means mailbox lifecycle and transport decisions affect a concentrated namespace rather than a thin edge population." -f (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Recipients in current source'), (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Top recipient type 1'), (Get-CustomerObservationState -Observation $messagingObservation -Signal 'Top recipient domain 1')) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '9.2 SMTP Relay Usage' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Outbound email configuration was reviewed to understand whether system-generated mail is still relying on mailbox-based SMTP authentication, connector-based relay, or direct send patterns. This matters because automated sending paths often remain in place long after the original application or device owner changes.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $smtpRelayUsageRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Benefits of Transitioning' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('DisplayName', 'UserPrincipalName', 'Send Count', 'Last Activity') -Rows $smtpRelayServiceAccountRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current source can show outbound activity and SMTP posture, but it does not always attribute every sender cleanly to a dedicated relay service account. Even so, the overlap between SMTP-authenticated accounts and message activity is enough to show whether mail relay is being handled as a defined service pattern or through mailbox accounts that now require ownership review.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '10. Microsoft Teams Governance and Cleanup' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'An analysis of the current Microsoft Teams environment shows how collaboration governance is being carried in practice. Team ownership, dormancy, guest presence, and voice workload signals together provide a clearer picture than a raw Team count alone.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Key Findings' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $teamsGovernanceRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $(if ($null -ne $collaborationConsultativeSummary) { $collaborationConsultativeSummary.Narrative } else { $collaborationObservation.ObservedNarrative })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Opportunities for Cleanup' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $collaborationObservation.WhyItMatters) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $collaborationConsultativeSummary) { $collaborationConsultativeSummary.RecommendationSupport } else { 'This section supports the collaboration ownership, lifecycle, and external-sharing recommendations in 4.0.' })) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '11.0 SharePoint Online Storage and External Sharing' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("SharePoint and OneDrive storage were reviewed together with external-sharing posture because those signals show whether collaboration growth is still being matched by ownership, lifecycle, and sharing control. The current source includes {0} SharePoint sites and {1} OneDrive locations, which is enough to see where content has continued to accumulate." -f $sharePointRows.Count, $oneDriveRows.Count) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Tenant External Sharing Snapshot' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $externalSharingSnapshotRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'SharePoint Tenant Controls Snapshot' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $sharePointTenantControlRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Observations and Recommendations' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Site Title', 'Owner', 'Storage Used (GB)', 'Last Content Modified') -Rows $sharePointStorageRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("External sharing remains an important part of the collaboration posture in this tenant. Tenant-level sharing is currently shown as {0}, the default sharing link type is {1}, and {2} reviewed SharePoint site(s) surfaced an external-sharing capability in the source data. Team-connected sites account for {3} of the reviewed SharePoint locations, which reinforces how closely SharePoint governance is tied to broader Teams and group ownership patterns." -f $tenantSharingCapabilityText, $defaultSharingLinkTypeText, $sharePointSitesExternalSharingEnabled, $teamConnectedSites) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The external-sharing summary also shows sharing domain restriction mode as {0}, with {1} site-level sharing override(s) identified in the reviewed site inventory. That combination is important because it shows whether external exposure is being controlled only at the tenant level or is also being shaped materially by site-level exceptions." -f $sharingDomainRestrictionModeText, $siteOverrideCountText) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The broader SharePoint tenant settings also show whether operational controls are reinforcing the sharing baseline or leaving it to stand on its own. The current source records legacy auth protocols as {0}, custom app authentication disabled as {1}, unmanaged sync restriction as {2}, and deleted-user personal site retention as {3} day(s). Those settings matter because external collaboration risk is shaped not only by link defaults, but also by sync controls, client behavior, and how long stale personal content remains in the tenant." -f (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsLegacyAuthProtocolsEnabled')) -Default 'Not surfaced in current source'), (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('DisableCustomAppAuthentication')) -Default 'Not surfaced in current source'), (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsUnmanagedSyncAppForTenantRestricted')) -Default 'Not surfaced in current source'), (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('DeletedUserPersonalSiteRetentionPeriodInDays')) -Default 'Not surfaced in current source')) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'External Exposure Review' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Workload', 'Asset Type', 'Title', 'Exposure Category', 'Gap Reason', 'Review Priority') -Rows $externalExposureReviewRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The external exposure review currently surfaces {0} actionable row(s). The strongest concentrations are {1} stale externally exposed site or OneDrive row(s), {2} owner-drift or ownership-mismatch row(s), {3} guest-heavy or dormant collaboration row(s), and {4} baseline or override mismatch row(s). That mix is useful because it separates broad tenant posture from the specific workspaces that now deserve cleanup." -f $externalExposureFindings.Count, $staleExternalExposureCount, $ownerDriftExternalExposureCount, $guestHeavyExposureCount, $siteOverrideExposureCount) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This view is intentionally limited to app-only signals. It is not a file-permission crawl, but it is enough to show where tenant sharing settings, site-level exceptions, stale externally exposed content, and guest collaboration patterns have started to diverge from a tighter external-access baseline.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '12.0 Retention Policies and Data Loss Prevention' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $retentionRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The retention view is stronger than a simple yes-or-no check, but it still shows a gap between mailbox-level lifecycle controls and a clearly surfaced cross-workload retention or DLP program. In the current source, explicit retention signals are visible on individual mailboxes, while DLP detail remains limited or absent.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This matters because retention and DLP are the point where messaging, collaboration, and compliance expectations converge. If those controls are partially implemented or insufficiently visible, legal, operational, and security outcomes become harder to validate with confidence.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '13.0 Domain Configuration and DNS Overview' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The tenant review surfaced {0} verified domain(s) out of {1} domain record(s) reviewed. This section keeps domain registration state and DNS trust controls together because accepted-domain hygiene, spoofing resistance, and message trust are closely related in the current environment." -f $verifiedDomains.Count, $(if ($domainRows.Count -gt 0) { $domainRows.Count } else { 'an unconfirmed number of' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Domain', 'Verified', 'Authentication Type', 'Domain Type', 'Default', 'DMARC') -Rows $domainOverviewRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text $(if ($null -ne $governanceConsultativeSummary) { $governanceConsultativeSummary.Narrative } else { 'The reviewed domain inventory did not surface a separate domain-registration remediation item beyond the authentication and DNS posture shown below. Where accepted domains remain active, verification state and default-domain usage should still be confirmed as part of normal namespace governance.' }) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'DNS Configuration' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Domain', 'SPF', 'DKIM', 'DMARC', 'Notes') -Rows $dnsRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("DMARC is currently configured on {0} domain(s), while DKIM is configured on {1} domain(s). The largest recipient concentrations remain in {2}. That mix is notable because mail-authentication posture and recipient concentration are closely linked in a tenant where most messaging volume stays centered on a small set of accepted domains." -f $dmarcEnabledDomains, $dkimEnabledDomains, $(if ($largestDomainsByRecipients.Count -gt 0) { ((@($largestDomainsByRecipients | Select-Object -First 2 | ForEach-Object { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('Domain', 'Id')) -Default 'Not surfaced in current source' })) -join '; ') } else { 'domains not surfaced clearly enough for comparison' })) -Style 'Normal')) | Out-Null
    if ($domainsWithoutDmarc.Count -gt 0 -or $domainsWithoutDkim.Count -gt 0 -or $domainsWithoutSpfSignal.Count -gt 0) {
        $dnsFocusParts = New-Object System.Collections.Generic.List[string]
        if ($domainsWithoutDmarc.Count -gt 0) { $dnsFocusParts.Add(("{0} domain(s) do not currently show DMARC" -f $domainsWithoutDmarc.Count)) | Out-Null }
        if ($domainsWithoutDkim.Count -gt 0) { $dnsFocusParts.Add(("{0} domain(s) do not currently show DKIM" -f $domainsWithoutDkim.Count)) | Out-Null }
        if ($domainsWithoutSpfSignal.Count -gt 0) { $dnsFocusParts.Add(("{0} domain(s) do not currently show an SPF signal in the reviewed data" -f $domainsWithoutSpfSignal.Count)) | Out-Null }
        $blocks.Add((New-CustomerWordParagraphBlock -Text ("What stands out here is the concentration of mail-authentication gaps across active namespaces. The current source shows {0}." -f (($dnsFocusParts.ToArray()) -join '; ')) -Style 'Normal')) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current source does not show a separate DNS hardening issue beyond the per-domain SPF, DKIM, and DMARC posture documented above.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $governanceConsultativeSummary) { $governanceConsultativeSummary.RecommendationSupport } else { 'This section supports the governance, domain, and security-baseline recommendations in 4.0.' })) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '14.0 Offboarding Recommendation' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Lifecycle signals in this tenant suggest that access, content, and shared workloads are aging out of active ownership at different rates. That is a common sign that offboarding is being handled tactically across workloads rather than through one consistently governed process.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Key Offboarding Objectives' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current evidence suggests three objectives should remain central: remove stale access promptly, preserve business data only where ownership is clear, and reclaim licenses and shared workloads that no longer have an active sponsor.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recommended Offboarding Workflow' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'A prescriptive end-to-end offboarding workflow was not surfaced in the current source. The assessment evidence does, however, support the lifecycle actions already prioritized in the recommendation set, especially for stale privileged access, inactive guests, stale collaboration locations, and shared mailboxes without ownership signals.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Supporting Observations from Environment Review' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Lifecycle Signal', 'Current State') -Rows $offboardingSupportRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $(if ($null -ne $lifecycleConsultativeSummary) { $lifecycleConsultativeSummary.Narrative } else { $lifecycleObservation.ObservedNarrative })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $lifecycleObservation.WhyItMatters) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $lifecycleConsultativeSummary) { $lifecycleConsultativeSummary.RecommendationSupport } else { 'This section supports the lifecycle and ownership-governance recommendations in 4.0.' })) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.0 Appendix' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The appendix sections that follow provide Microsoft documentation references and a full finding reference table so the assessment can be reviewed both as a leadership document and as a working technical reference.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.1 Device Management' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device management is central to maintaining security and compliance in a tenant where unmanaged and stale devices remain part of the current footprint. The references below support the endpoint observations documented in this assessment.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Get started with device compliance policies in Microsoft Intune - https://learn.microsoft.com/en-us/intune/intune-service/protect/device-compliance-get-started', 'Supports the compliance baseline and managed-device observations in the endpoint review.')),
        (New-CustomerWordTableRow -Cells @('Require compliant or hybrid Microsoft Entra joined device - https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-all-users-device-compliance', 'Provides Microsoft guidance for tying device state to access-control enforcement.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.2 Entra Guest Access Best Practices' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Guest access best practices matter in this tenant because external collaboration, inactive guests, and sharing posture are all part of the current risk picture. These references support the guest lifecycle and external-access observations in the report.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('B2B collaboration fundamentals - https://learn.microsoft.com/en-us/entra/external-id/b2b-fundamentals', 'Supports guest access governance and external collaboration design.')),
        (New-CustomerWordTableRow -Cells @('Overview of external sharing in SharePoint and OneDrive - https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview', 'Provides Microsoft guidance for the collaboration-sharing observations in this report.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.3 Application Consent and Authentication Methods' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application governance and authentication-method controls both affect how quickly identity exposure can grow in a tenant. The references in this appendix support the application-consent, MFA, and authentication-method observations documented earlier in the assessment.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application User Consent Management' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application user consent should be managed deliberately wherever users can authorize apps to access organizational data. In a tenant where privileged app permissions and consent workflow maturity are already part of the review, consent governance becomes an important control boundary.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recommended Resources' -Style 'Heading4')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Configure the admin consent workflow - https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-admin-consent-workflow', 'Supports the governance observations around application consent, approval workflow, and control maturity.')),
        (New-CustomerWordTableRow -Cells @('Manage authentication methods - https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-authentication-methods-manage', 'Relevant where app access, MFA, and modern authentication governance are being reviewed together.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Authentication Methods Migration' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Migrating from legacy MFA and SSPR controls to the Authentication Methods policy becomes especially important when the tenant already shows mixed enforcement and uneven registration. The resources below support that transition.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Key Resources for Authentication Migration' -Style 'Heading4')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Manage authentication methods - https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-authentication-methods-manage', 'Supports migration away from legacy MFA and SSPR policy management.')),
        (New-CustomerWordTableRow -Cells @('Self-service password reset deep dive - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks', 'Provides Microsoft guidance for SSPR design and operational implications.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.4 Access and Privileged Identity Management' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Conditional Access and privileged-role governance are two of the strongest levers available for reducing identity risk in this tenant. These references support the policy-state, exclusions, and standing-admin observations documented in the report.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Conditional Access overview - https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview', 'Supports the policy coverage, exclusions, and enforcement observations.')),
        (New-CustomerWordTableRow -Cells @('Privileged Identity Management overview - https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure', 'Supports the recommendations related to privileged role hygiene and reducing standing access.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.5 Exchange Online Archives and SMTP Relay' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The Exchange appendix supports the mailbox-growth, archive, forwarding, and relay observations documented in the assessment. These references are useful where mailbox lifecycle and transport controls are both part of the same remediation path.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Exchange Online Archives and Retention Policies' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Archive usage and mailbox retention are part of the same long-term lifecycle story. The references below support archive enablement, mailbox retention, and storage planning.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Learn about retention policies and retention labels - https://learn.microsoft.com/en-us/purview/retention', 'Supports retention-policy observations and mailbox lifecycle planning.')),
        (New-CustomerWordTableRow -Cells @('On-premises password writeback with self-service password reset - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-writeback', 'Relevant when archive, retention, and lifecycle handling intersect with hybrid identity operations.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Third-Party SMTP Relay Configuration' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Relay and application-based sending paths should be documented as service patterns rather than left attached to mailbox accounts by default. These references support the SMTP relay observations in the messaging review.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Control automatic external email forwarding in Microsoft 365 - https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-policies-external-email-forwarding', 'Supports the forwarding-control observations in the messaging section.')),
        (New-CustomerWordTableRow -Cells @('How to set up a multifunction device or application to send email using Microsoft 365 or Office 365 - https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-microsoft-365-or-office-365', 'Relevant to SMTP relay, connector, and mail-flow exception handling.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.6 Microsoft Teams and SharePoint Online' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Teams and SharePoint governance are closely linked in this tenant because collaboration growth, ownership, and sharing posture are moving together. The following Microsoft guidance supports the observations documented in the collaboration sections.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Microsoft Teams Governance' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Teams governance references are most relevant where ownerless workspaces, dormant collaboration spaces, and group-creation controls are part of the current-state review.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Manage who can create Microsoft 365 Groups - https://learn.microsoft.com/en-us/microsoft-365/solutions/manage-creation-of-groups', 'Supports governance of Teams-connected groups and workspace sprawl.')),
        (New-CustomerWordTableRow -Cells @('Set expiration for Microsoft 365 groups - https://learn.microsoft.com/en-us/entra/identity/users/groups-lifecycle', 'Relevant to dormant collaboration spaces and lifecycle control.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.7 SharePoint Online Collaboration' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'SharePoint and OneDrive guidance is particularly relevant where storage growth, stale content, and external sharing need to be evaluated together rather than as separate issues.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Overview of external sharing in SharePoint and OneDrive - https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview', 'Supports the collaboration-sharing observations in this report.')),
        (New-CustomerWordTableRow -Cells @('Retention and deletion in OneDrive and SharePoint - https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion', 'Relevant to stale OneDrive and SharePoint lifecycle handling.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.8 DNS DMARC and OneDrive' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This appendix supports the domain-authentication, DMARC, and OneDrive lifecycle observations that surfaced during the review of accepted domains and collaboration services.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'DMARC Records' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'DMARC guidance is especially relevant where recipient volume is concentrated on a small set of accepted domains and mail-authentication posture is uneven across those namespaces.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Set up SPF in Microsoft 365 to help prevent spoofing - https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/set-up-spf-in-office-365-to-help-prevent-spoofing', 'Supports SPF and anti-spoofing guidance for the reviewed domains.')),
        (New-CustomerWordTableRow -Cells @('Use DKIM to validate outbound email sent from your custom domain - https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dkim-configure', 'Supports the observed DKIM posture for custom domains.')),
        (New-CustomerWordTableRow -Cells @('Use DMARC to validate email, setup steps - https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dmarc-configure', 'Supports DMARC record design, reporting, and policy progression.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'OneDrive' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'OneDrive guidance is relevant where user lifecycle, delegated stewardship, and stale personal content repositories are part of the current-state review.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Retention and deletion in OneDrive and SharePoint - https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion', 'Relevant to stale OneDrive lifecycle handling and post-departure content management.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.9 Retention Policies and Data Loss Prevention' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Retention and DLP guidance becomes especially important where the tenant shows partial retention visibility, but not enough evidence to confirm a mature cross-workload lifecycle and data-protection program.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Additional Resources' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Learn about retention policies and retention labels - https://learn.microsoft.com/en-us/purview/retention', 'Supports retention-policy observations and the need for documented lifecycle controls.')),
        (New-CustomerWordTableRow -Cells @('Learn about data loss prevention - https://learn.microsoft.com/en-us/purview/dlp-learn-about-dlp', 'Provides Microsoft guidance for DLP and data-protection governance.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.9 Pass-Through Authentication and Password Writeback' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Hybrid identity references are included here because directory synchronization, PTA, password writeback, and SSPR visibility all influence how identity operations can be supported safely and consistently.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Pass-Through Authentication (PTA)' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'PTA guidance is relevant anywhere the tenant depends on on-premises credential validation or is still deciding between hybrid sign-in approaches. These references support the hybrid identity observations surfaced in the report.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Microsoft Entra Connect: Pass-through Authentication - https://learn.microsoft.com/en-us/azure/active-directory/hybrid/how-to-connect-pta', 'Supports pass-through authentication design, agent requirements, and operational considerations.')),
        (New-CustomerWordTableRow -Cells @('Microsoft Entra Connect: User sign-in - https://learn.microsoft.com/en-us/azure/active-directory/hybrid/plan-connect-user-signin', 'Provides comparison guidance for hybrid sign-in options and role selection.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Password Writeback with AD Sync' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Password writeback and SSPR references are included because they directly affect password recovery, hybrid lifecycle handling, and end-user support workflows.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Self-service password reset deep dive - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks', 'Supports the password reset and SSPR observations in the report.')),
        (New-CustomerWordTableRow -Cells @('On-premises password writeback with self-service password reset - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-writeback', 'Relevant to password writeback and hybrid credential-management considerations.')),
        (New-CustomerWordTableRow -Cells @('Enable Microsoft Entra password writeback - https://learn.microsoft.com/en-us/azure/active-directory/authentication/tutorial-enable-sspr-writeback', 'Provides implementation guidance where password writeback is part of the target operating model.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.10 Full Findings Inventory' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This appendix table contains the full normalized findings inventory that supports the recommendations and workstream summaries in the main body of the report. The entries remain in the same order and preserve the same severity, priority, and recommendation text used throughout the assessment output.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Severity', 'Priority', 'Workstream', 'Rule / Finding', 'Why Flagged', 'Recommended Action', 'Success Criteria') -Rows $fullFindingsInventoryRows)) | Out-Null

    return @($blocks.ToArray())
}
