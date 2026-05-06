function Get-CustomerAssessmentTemplatePath {
    [CmdletBinding()]
    param()

    $templatePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\..\templates\customer\Microsoft 365 Tenant Best Practices Assessment Template.docx'))
    if (-not (Test-Path -Path $templatePath -PathType Leaf)) {
        throw "Customer assessment template was not found: $templatePath"
    }

    return $templatePath
}

function Get-RoadmapRemediationTemplatePath {
    [CmdletBinding()]
    param()

    $templatePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\..\..\docs\templates\M365 - Roadmap Remediation Plan_05.01.26 - Empty Copy.docx'))
    if (-not (Test-Path -Path $templatePath -PathType Leaf)) {
        throw "Roadmap remediation template was not found: $templatePath"
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

function New-CustomerWordListBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string[]]$Items,
        [Parameter(Mandatory = $false)][string]$Style = 'ListBullet'
    )

    return [pscustomobject]@{
        Type  = 'List'
        Items = @($Items | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        Style = $Style
    }
}

function New-CustomerWordImageBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][byte[]]$ImageBytes,
        [Parameter(Mandatory = $true)][string]$AltText,
        [Parameter(Mandatory = $false)][int]$WidthPx = 640,
        [Parameter(Mandatory = $false)][int]$HeightPx = 240,
        [Parameter(Mandatory = $false)][string]$ContentType = 'image/png'
    )

    return [pscustomobject]@{
        Type        = 'Image'
        ImageBytes  = $ImageBytes
        AltText     = $AltText
        WidthPx     = $WidthPx
        HeightPx    = $HeightPx
        ContentType = $ContentType
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

function New-CustomerWordTableCellContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Content,
        [Parameter(Mandatory = $false)][string]$FillColor
    )

    return [pscustomobject]@{
        CellType  = 'Formatted'
        Content   = $Content
        FillColor = $FillColor
    }
}

function Get-CustomerWordTableCellMetadata {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$Content)

    if ($null -ne $Content -and $Content.PSObject -and $Content.PSObject.Properties['CellType'] -and [string]$Content.CellType -eq 'Formatted') {
        return [pscustomobject]@{
            Content   = $Content.Content
            FillColor = [string]$Content.FillColor
        }
    }

    return [pscustomobject]@{
        Content   = $Content
        FillColor = $null
    }
}

function Get-CustomerDistinctRoadmapActions {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object[]]$RoadmapActions = @())

    $priorityRanks = @{
        'Immediate' = 0
        'Critical'  = 0
        'Near Term' = 1
        'High'      = 1
        'Planned'   = 2
        'Medium'    = 2
        'Monitor'   = 3
        'Low'       = 3
        'Info'      = 4
    }
    $phaseRanks = @{
        'Immediate' = 0
        'Near Term' = 1
        'Planned'   = 2
        'Monitor'   = 3
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
            @{ Expression = { $priority = [string](Get-ArrayaObjectValue -Object $_ -Names @('PriorityBand', 'Priority')); if ($priorityRanks.ContainsKey($priority)) { $priorityRanks[$priority] } else { 99 } } }, `
            @{ Expression = { if ($phaseRanks.ContainsKey([string]$_.RoadmapPhase)) { $phaseRanks[[string]$_.RoadmapPhase] } else { 99 } } }, `
            @{ Expression = { [string]$_.ActionTitle } }) | Select-Object -First 1

        if ($null -ne $bestAction) {
            $bestAction
        }
    }

    return @($distinct | Sort-Object `
        @{ Expression = { if ($phaseRanks.ContainsKey([string]$_.RoadmapPhase)) { $phaseRanks[[string]$_.RoadmapPhase] } else { 99 } } }, `
        @{ Expression = { $priority = [string](Get-ArrayaObjectValue -Object $_ -Names @('PriorityBand', 'Priority')); if ($priorityRanks.ContainsKey($priority)) { $priorityRanks[$priority] } else { 99 } } }, `
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

    $cellMetadata = Get-CustomerWordTableCellMetadata -Content $Content
    $Content = $cellMetadata.Content

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
            'List' {
                foreach ($item in @($block.Items)) {
                    $text = Convert-ToArrayaDisplayText -Value $item -Default ''
                    if ([string]::IsNullOrWhiteSpace($text)) {
                        continue
                    }

                    $lines.Add('- ' + $text) | Out-Null
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
            'Image' {
                continue
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

function Get-CustomerCompactNarrativeSentences {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Text,
        [Parameter(Mandatory = $false)][int]$MaxSentences = 2,
        [Parameter(Mandatory = $false)][int]$MaxLength = 180
    )

    $normalized = Convert-ToCustomerAssessmentNarrativeText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @()
    }

    $sentences = @(
        ($normalized -split '(?<=[.!?])\s+') |
            ForEach-Object { ($_ -replace '\s+', ' ').Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($sentences.Count -eq 0) {
        $sentences = @((($normalized -replace '\s+', ' ').Trim()))
    }

    return @(
        $sentences |
            Select-Object -First ([Math]::Max(1, $MaxSentences)) |
            ForEach-Object {
                if ($_.Length -gt $MaxLength) {
                    ($_.Substring(0, [Math]::Max(1, $MaxLength - 3)).TrimEnd() + '...')
                }
                else {
                    $_
                }
            }
    )
}

function Get-CustomerCompactNarrativeLine {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Text,
        [Parameter(Mandatory = $false)][int]$MaxSentences = 1,
        [Parameter(Mandatory = $false)][int]$MaxLength = 160
    )

    $segments = @(Get-CustomerCompactNarrativeSentences -Text $Text -MaxSentences $MaxSentences -MaxLength $MaxLength)
    if ($segments.Count -eq 0) {
        return $null
    }

    return ($segments -join ' ')
}

function Get-CustomerLabeledNarrativeBulletItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [AllowNull()][string]$Text,
        [Parameter(Mandatory = $false)][string]$ContinuationLabel,
        [Parameter(Mandatory = $false)][int]$MaxSentences = 1,
        [Parameter(Mandatory = $false)][int]$MaxLength = 180
    )

    $segments = @(Get-CustomerCompactNarrativeSentences -Text $Text -MaxSentences $MaxSentences -MaxLength $MaxLength)
    if ($segments.Count -eq 0) {
        return @()
    }

    return @(
        for ($index = 0; $index -lt $segments.Count; $index++) {
            $currentLabel = if ($index -eq 0 -or [string]::IsNullOrWhiteSpace($ContinuationLabel)) {
                $Label
            }
            else {
                $ContinuationLabel
            }

            '{0}: {1}' -f $currentLabel, $segments[$index]
        }
    )
}

function Get-CustomerRoadmapNarrativeSegments {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Text,
        [Parameter(Mandatory = $false)][int]$MaxSentences = 3,
        [Parameter(Mandatory = $false)][int]$MaxLength = 155
    )

    $normalized = Convert-ToCustomerAssessmentNarrativeText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @()
    }

    $sentences = @(
        ($normalized -split '(?<=[.!?])\s+') |
            ForEach-Object { ($_ -replace '\s+', ' ').Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First ([Math]::Max(1, $MaxSentences))
    )

    if ($sentences.Count -eq 0) {
        $sentences = @((($normalized -replace '\s+', ' ').Trim()))
    }

    function Format-CustomerRoadmapNarrativeSegment {
        param([AllowNull()][string]$Segment)

        $clean = (($Segment -replace '\s+', ' ').Trim() -replace '^(?i)(and|but|so)\s+', '')
        if ([string]::IsNullOrWhiteSpace($clean)) {
            return $null
        }

        if ($clean.Length -gt 1 -and [char]::IsLower($clean[0])) {
            return ([char]::ToUpperInvariant($clean[0]) + $clean.Substring(1))
        }

        return $clean
    }

    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($sentence in $sentences) {
        if ($sentence.Length -le $MaxLength) {
            $formattedSentence = Format-CustomerRoadmapNarrativeSegment -Segment $sentence
            if (-not [string]::IsNullOrWhiteSpace($formattedSentence)) {
                $segments.Add($formattedSentence) | Out-Null
            }
            continue
        }

        $pieces = @(
            ($sentence -split ';\s+') |
                ForEach-Object { ($_ -replace '\s+', ' ').Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ($pieces.Count -le 1) {
            $segments.Add(($sentence.Substring(0, [Math]::Max(1, $MaxLength - 3)).TrimEnd() + '...')) | Out-Null
            continue
        }

        $current = ''
        foreach ($piece in $pieces) {
            $candidate = if ([string]::IsNullOrWhiteSpace($current)) { $piece } else { "$current, $piece" }
            if ($candidate.Length -le $MaxLength) {
                $current = $candidate
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($current)) {
                $formattedCurrent = Format-CustomerRoadmapNarrativeSegment -Segment $current
                if (-not [string]::IsNullOrWhiteSpace($formattedCurrent)) {
                    $segments.Add($formattedCurrent) | Out-Null
                }
            }

            $current = if ($piece.Length -gt $MaxLength) {
                $piece.Substring(0, [Math]::Max(1, $MaxLength - 3)).TrimEnd() + '...'
            }
            else {
                $piece
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($current)) {
            $formattedCurrent = Format-CustomerRoadmapNarrativeSegment -Segment $current
            if (-not [string]::IsNullOrWhiteSpace($formattedCurrent)) {
                $segments.Add($formattedCurrent) | Out-Null
            }
        }
    }

    return @($segments.ToArray() | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
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

function Get-CustomerEnterpriseApplicationSsoModeLabel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow,
        [Parameter(Mandatory = $false)]
        [string]$Default = 'Not detected'
    )

    $mode = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('SSOMode', 'PreferredSingleSignOnMode')) -Default ''
    if ([string]::IsNullOrWhiteSpace($mode)) {
        return $Default
    }

    switch -Regex ($mode.ToLowerInvariant()) {
        '^saml$' { return 'SAML' }
        '^(oidc|openidconnect)$' { return 'OIDC' }
        '^wsfed$' { return 'WS-Fed' }
        '^password$' { return 'Password-based SSO' }
        '^external$' { return 'External' }
        default { return $mode }
    }
}

function Get-CustomerEnterpriseApplicationCredentialReviewState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    $state = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('CredentialReviewState')) -Default ''
    if (-not [string]::IsNullOrWhiteSpace($state)) {
        return $state
    }

    foreach ($signalName in @('ExpiredCredentialCount', 'ExpiringCredentialCount', 'KeyCredentialCount', 'PasswordCredentialCount')) {
        if ($null -ne (Get-ArrayaObjectValue -Object $ApplicationRow -Names @($signalName))) {
            return 'Collected'
        }
    }

    $appCredentials = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('AppCredentials')) -Default ''
    if (-not [string]::IsNullOrWhiteSpace($appCredentials)) {
        return 'Collected'
    }

    return 'Unavailable'
}

function Get-CustomerEnterpriseApplicationLatestActivityDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    $activityDates = New-Object System.Collections.Generic.List[datetime]
    foreach ($propertyName in @(
        'LastSignInDateTime',
        'DelegatedLastSignIn',
        'ApplicationLastSignIn',
        'LastServicePrincipalSignInDateTime'
    )) {
        $activityDate = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $ApplicationRow -Names @($propertyName))
        if ($null -ne $activityDate) {
            $activityDates.Add($activityDate) | Out-Null
        }
    }

    if ($activityDates.Count -eq 0) {
        return $null
    }

    return @($activityDates.ToArray() | Sort-Object -Descending | Select-Object -First 1)[0]
}

function Get-CustomerEnterpriseApplicationActivityText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    $activitySignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'ActivitySignalState'
    $hasRecentActivity = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HasRecentActivity'))
    $latestActivityDate = Get-CustomerEnterpriseApplicationLatestActivityDate -ApplicationRow $ApplicationRow

    if ($null -ne $latestActivityDate) {
        return ('Latest activity {0}' -f $latestActivityDate.ToString('yyyy-MM-dd'))
    }

    if ($activitySignalState -eq 'Collected' -and $hasRecentActivity -eq $false) {
        return 'No recent activity surfaced'
    }

    if ($activitySignalState -eq 'Partial') {
        return 'Activity review only partially validated'
    }

    if ($activitySignalState -eq 'Unavailable') {
        return 'Activity not validated from the reviewed data'
    }

    return 'Activity not clearly surfaced'
}

function Test-CustomerEnterpriseApplicationHasExpiredCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    return (
        (Get-CustomerEnterpriseApplicationCredentialReviewState -ApplicationRow $ApplicationRow) -eq 'Collected' -and
        ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HasExpiredCredentials'))) -eq $true -or
        (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('ExpiredCredentialCount'))) -gt 0)
    )
}

function Test-CustomerEnterpriseApplicationHasCredentialsExpiringSoon {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    return (
        (Get-CustomerEnterpriseApplicationCredentialReviewState -ApplicationRow $ApplicationRow) -eq 'Collected' -and
        ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HasCredentialsExpiringSoon'))) -eq $true -or
        (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('ExpiringCredentialCount'))) -gt 0)
    )
}

function Get-CustomerEnterpriseApplicationCredentialStateText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    $credentialReviewState = Get-CustomerEnterpriseApplicationCredentialReviewState -ApplicationRow $ApplicationRow
    $credentialIssueSummary = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('CredentialIssueSummary')) -Default ''
    if (-not [string]::IsNullOrWhiteSpace($credentialIssueSummary)) {
        return $credentialIssueSummary
    }

    $appCredentials = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('AppCredentials')) -Default ''
    if ($credentialReviewState -ne 'Collected') {
        if (-not [string]::IsNullOrWhiteSpace($appCredentials)) {
            return "Credential types surfaced ($appCredentials), but expiry review was not fully validated"
        }

        return 'Credential lifecycle not validated from the reviewed data'
    }

    if (-not [string]::IsNullOrWhiteSpace($appCredentials)) {
        $nextExpiry = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('NextCredentialExpiryDateTime'))
        if ($null -ne $nextExpiry) {
            return '{0}; next expiry {1}' -f $appCredentials, $nextExpiry.ToString('yyyy-MM-dd')
        }

        return $appCredentials
    }

    return 'No app credentials surfaced'
}

function Get-CustomerEnterpriseApplicationPermissionSummaryText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    $segments = New-Object System.Collections.Generic.List[string]
    $highPrivilegeCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HighPrivilegePermissionCount'))
    $applicationPermissionCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('ApplicationPermissionCount'))
    $delegatedPermissionGrantCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('DelegatedPermissionGrantCount'))

    if ($null -ne $highPrivilegeCount -and $highPrivilegeCount -gt 0) {
        $segments.Add(('{0} high-privilege permission signal(s)' -f $highPrivilegeCount)) | Out-Null
    }
    if ($null -ne $applicationPermissionCount -and $applicationPermissionCount -gt 0) {
        $segments.Add(('{0} application permission(s)' -f $applicationPermissionCount)) | Out-Null
    }
    if ($null -ne $delegatedPermissionGrantCount -and $delegatedPermissionGrantCount -gt 0) {
        $segments.Add(('{0} delegated grant(s)' -f $delegatedPermissionGrantCount)) | Out-Null
    }

    if ($segments.Count -eq 0) {
        return 'No broad permission signal surfaced'
    }

    return Join-ArrayaReadableList -Items @($segments.ToArray())
}

function Get-CustomerEnterpriseApplicationObservationText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    $observationParts = New-Object System.Collections.Generic.List[string]
    $displayName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('DisplayName')) -Default 'This application'
    $ssoEnabled = Test-CustomerEnterpriseApplicationSsoEnabled -ApplicationRow $ApplicationRow
    $ssoMode = Get-CustomerEnterpriseApplicationSsoModeLabel -ApplicationRow $ApplicationRow -Default 'Configured'
    $permissionSummary = Get-CustomerEnterpriseApplicationPermissionSummaryText -ApplicationRow $ApplicationRow
    $applicationSource = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('ApplicationSource')) -Default 'Source not classified'
    $ownerSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'OwnerSignalState'
    $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('OwnerCount'))
    $redirectSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'RedirectUriSignalState'
    $insecureRedirectUriCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('InsecureRedirectUriCount'))
    $activitySignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'ActivitySignalState'
    $hasRecentActivity = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HasRecentActivity'))
    $appRoleAssignmentRequired = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('AppRoleAssignmentRequired'))

    if ($ssoEnabled -eq $true) {
        $observationParts.Add("$displayName uses SSO via $ssoMode.") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($permissionSummary) -and $permissionSummary -ne 'No broad permission signal surfaced') {
        $observationParts.Add("Permission posture shows $permissionSummary.") | Out-Null
    }
    if ($appRoleAssignmentRequired -eq $true) {
        $observationParts.Add('User assignment is required before access is granted.') | Out-Null
    }
    if ($ownerSignalState -eq 'Collected' -and $null -ne $ownerCount -and $ownerCount -le 0) {
        $observationParts.Add('No owner signal was surfaced for the associated app registration.') | Out-Null
    }
    elseif ($ownerSignalState -eq 'Unavailable') {
        $observationParts.Add('Owner validation was not fully completed in the reviewed data.') | Out-Null
    }
    if (Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $ApplicationRow) {
        $observationParts.Add(("Credential review surfaced expired material: {0}." -f (Get-CustomerEnterpriseApplicationCredentialStateText -ApplicationRow $ApplicationRow))) | Out-Null
    }
    elseif (Test-CustomerEnterpriseApplicationHasCredentialsExpiringSoon -ApplicationRow $ApplicationRow) {
        $observationParts.Add(("Credential lifecycle needs near-term review: {0}." -f (Get-CustomerEnterpriseApplicationCredentialStateText -ApplicationRow $ApplicationRow))) | Out-Null
    }
    elseif ((Get-CustomerEnterpriseApplicationCredentialReviewState -ApplicationRow $ApplicationRow) -eq 'Collected') {
        $appCredentials = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('AppCredentials')) -Default ''
        if (-not [string]::IsNullOrWhiteSpace($appCredentials)) {
            $observationParts.Add(("Application credentials surfaced: {0}." -f $appCredentials)) | Out-Null
        }
    }
    if ($null -ne $insecureRedirectUriCount -and $insecureRedirectUriCount -gt 0) {
        $observationParts.Add(("Redirect URI review flagged {0} endpoint(s) for follow-up." -f $insecureRedirectUriCount)) | Out-Null
    }
    elseif ($redirectSignalState -eq 'Partial') {
        $observationParts.Add('Redirect URI review was only partially validated in the reviewed data.') | Out-Null
    }
    if ($activitySignalState -eq 'Collected' -and $hasRecentActivity -eq $false) {
        $observationParts.Add('No recent activity signal was surfaced across the reviewed app sign-in telemetry.') | Out-Null
    }
    elseif ($activitySignalState -eq 'Partial') {
        $observationParts.Add('Recent activity validation was only partially completed in the reviewed data.') | Out-Null
    }

    $latestActivityText = Get-CustomerEnterpriseApplicationActivityText -ApplicationRow $ApplicationRow
    if (-not [string]::IsNullOrWhiteSpace($latestActivityText) -and $latestActivityText -notmatch '^Activity not ') {
        $observationParts.Add("$latestActivityText.") | Out-Null
    }

    if ($observationParts.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($applicationSource) -and $applicationSource -ne 'Source not classified') {
        $observationParts.Add("Application source is classified as $applicationSource.") | Out-Null
    }
    if ($observationParts.Count -eq 0) {
        $observationParts.Add('Application inventory was surfaced, but the current tenant data did not expose a stronger permission, activity, or credential cleanup signal for this app.') | Out-Null
    }

    return ($observationParts -join ' ')
}

function Test-CustomerEnterpriseApplicationRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    if ($null -eq $ApplicationRow -or -not $ApplicationRow.PSObject -or $ApplicationRow.PSObject.Properties.Count -eq 0) {
        return $false
    }

    foreach ($signal in @(
        (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('ServicePrincipalId', 'Id')),
        (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('AppId')),
        (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('DisplayName')),
        (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('SSOMode', 'PreferredSingleSignOnMode'))
    )) {
        $signalText = Convert-ToCustomerAssessmentDisplayText -Value $signal -Default ''
        if (-not [string]::IsNullOrWhiteSpace($signalText)) {
            return $true
        }
    }

    return (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('SsoEnabled'))) -eq $true
}

function Merge-CustomerEnterpriseApplicationRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object[]]$EnterpriseApplicationRows = @(),
        [Parameter(Mandatory = $false)]
        [object[]]$AuthenticationSsoApplicationRows = @()
    )

    $enterpriseApplicationByIdentity = @{}
    foreach ($applicationRow in @(@($EnterpriseApplicationRows) + @($AuthenticationSsoApplicationRows))) {
        if (-not (Test-CustomerEnterpriseApplicationRow -ApplicationRow $applicationRow)) {
            continue
        }

        $identityParts = @(
            Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $applicationRow -Names @('ServicePrincipalId', 'Id')) -Default ''
            Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $applicationRow -Names @('AppId')) -Default ''
            Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $applicationRow -Names @('DisplayName')) -Default ''
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        if ($identityParts.Count -eq 0) {
            $identityParts = @([guid]::NewGuid().Guid)
        }

        $applicationIdentity = ($identityParts -join '|').ToLowerInvariant()
        if ($enterpriseApplicationByIdentity.ContainsKey($applicationIdentity)) {
            $existing = $enterpriseApplicationByIdentity[$applicationIdentity]
            $mergedProperties = [ordered]@{}
            foreach ($property in @($existing.PSObject.Properties)) {
                $mergedProperties[$property.Name] = $property.Value
            }
            foreach ($property in @($applicationRow.PSObject.Properties)) {
                $hasCurrentValue = $mergedProperties.Contains($property.Name)
                $currentValue = if ($hasCurrentValue) { $mergedProperties[$property.Name] } else { $null }
                if (-not $hasCurrentValue -or $null -eq $currentValue -or [string]::IsNullOrWhiteSpace([string]$currentValue)) {
                    $mergedProperties[$property.Name] = $property.Value
                }
            }
            $enterpriseApplicationByIdentity[$applicationIdentity] = [pscustomobject]$mergedProperties
            continue
        }

        $enterpriseApplicationByIdentity[$applicationIdentity] = $applicationRow
    }

    return @(
        foreach ($applicationIdentity in @($enterpriseApplicationByIdentity.Keys | Sort-Object)) {
            $enterpriseApplicationByIdentity[$applicationIdentity]
        }
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
        ApplicationsWithExpiredCredentials = @($rows | Where-Object { Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $_ }).Count
        ApplicationsWithCredentialsExpiringSoon = @($rows | Where-Object { Test-CustomerEnterpriseApplicationHasCredentialsExpiringSoon -ApplicationRow $_ }).Count
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
    $hasExpiredCredentials = Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $ApplicationRow
    $hasCredentialsExpiringSoon = Test-CustomerEnterpriseApplicationHasCredentialsExpiringSoon -ApplicationRow $ApplicationRow

    return (
        (($null -ne $highPrivilegeCount) -and ($highPrivilegeCount -gt 0)) -or
        (($null -ne $applicationPermissionCount) -and ($applicationPermissionCount -gt 0)) -or
        (($null -ne $delegatedPermissionGrantCount) -and ($delegatedPermissionGrantCount -gt 0)) -or
        $hasExpiredCredentials -or
        $hasCredentialsExpiringSoon -or
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
                @{ Expression = { if (Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $_) { 1 } else { 0 } }; Descending = $true }, `
                @{ Expression = { if (Test-CustomerEnterpriseApplicationHasCredentialsExpiringSoon -ApplicationRow $_) { 1 } else { 0 } }; Descending = $true }, `
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
                @{ Expression = {
                    $latestActivity = Get-CustomerEnterpriseApplicationLatestActivityDate -ApplicationRow $_
                    if ($null -eq $latestActivity) { [datetime]::MinValue } else { $latestActivity }
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
    switch -Regex ($text.ToLowerInvariant()) {
        'excluded' { return 'Explicitly excluded' }
        'outside.*include scope' { return 'Outside include scope' }
        default { return $text }
    }
}

function Get-CustomerCompactMfaGapDriverText {
    [CmdletBinding()]
    param(
        [AllowNull()]$GapCategory,
        [AllowNull()]$GapReason,
        [AllowNull()]$RelatedPolicies
    )

    $policyText = Convert-ToCustomerAssessmentDisplayText -Value $RelatedPolicies -Default ''
    if (-not [string]::IsNullOrWhiteSpace($policyText) -and $policyText -ne 'Not validated from the reviewed data') {
        return "Policy: $policyText"
    }

    $reasonText = Convert-ToCustomerAssessmentDisplayText -Value $GapReason -Default ''
    if (-not [string]::IsNullOrWhiteSpace($reasonText) -and $reasonText -ne 'Not validated from the reviewed data') {
        if ($reasonText -match '(?i)outside the include scope') {
            return 'Outside include scope'
        }

        if ($reasonText -match '(?i)excluded through group') {
            return 'Excluded through group or policy scope'
        }

        if ($reasonText -match '(?i)\bexcluded\b') {
            return 'Explicitly excluded from enabled MFA scope'
        }

        $firstSentence = (($reasonText -split '(?<=[.!?])\s+', 2)[0]).Trim()
        if (-not [string]::IsNullOrWhiteSpace($firstSentence)) {
            if ($firstSentence.Length -gt 96) {
                return ($firstSentence.Substring(0, 93).TrimEnd() + '...')
            }

            return $firstSentence
        }
    }

    return (Get-CustomerMfaGapCategoryLabel -Value $GapCategory)
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

function Get-CustomerActionCriticalityLabel {
    [CmdletBinding()]
    param([AllowNull()][string]$Severity)

    switch (([string]$Severity).ToLowerInvariant()) {
        'critical' { return 'Critical' }
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

function Get-CustomerCriticalityFillColor {
    [CmdletBinding()]
    param([AllowNull()][string]$Criticality)

    switch (([string]$Criticality).ToLowerInvariant()) {
        'critical' { return 'FDE9E7' }
        'high' { return 'FCE5CD' }
        'medium' { return 'FFF2CC' }
        'low' { return 'E7E6E6' }
        'info' { return 'E7E6E6' }
        default { return $null }
    }
}

function Get-CustomerActionEffortLabel {
    [CmdletBinding()]
    param([AllowNull()]$Action)

    $effortTier = [string](Get-ArrayaObjectValue -Object $Action -Names @('EffortTier'))
    if (-not [string]::IsNullOrWhiteSpace($effortTier)) {
        return $effortTier
    }

    $actionTitle = [string](Get-ArrayaObjectValue -Object $Action -Names @('ActionTitle'))
    switch ($actionTitle) {
        'Improve device compliance and managed endpoint coverage' { return 'High' }
        'Establish accountable ownership for collaboration spaces' { return 'Medium' }
        'Reduce privileged access and strengthen identity controls' { return 'Medium' }
        'Strengthen domain and anti-spoofing controls' { return 'Medium' }
        'Reconcile license capacity and tenant governance gaps' { return 'Medium' }
        'Review external forwarding and mail flow exposure' { return 'Medium' }
        'Strengthen baseline security and access protections' { return 'Medium' }
        default { return 'Medium' }
    }
}

function Get-CustomerActionEstimatedPsHoursLabel {
    [CmdletBinding()]
    param([AllowNull()]$Action)

    $estimatedPsHours = [string](Get-ArrayaObjectValue -Object $Action -Names @('EstimatedPsHours'))
    if (-not [string]::IsNullOrWhiteSpace($estimatedPsHours)) {
        return $estimatedPsHours
    }

    return '14-24 hours'
}

function Get-CustomerTableRowCells {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$Row)

    if ($null -ne $Row -and $Row.PSObject -and $Row.PSObject.Properties['Cells']) {
        return @($Row.Cells)
    }

    if (($Row -is [System.Collections.IEnumerable]) -and -not ($Row -is [string]) -and -not ($Row -is [System.Collections.IDictionary])) {
        return @($Row)
    }

    return @($Row)
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
        'global administrator count exceeds' { return 'Standing Global Administrator access remains above the recommended operating threshold and should be reduced by removing stale assignments, moving infrequent admins to lower-privilege roles, and using eligible activation for full tenant-wide access.' }
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
            'The reviewed signals in this workstream did not surface one concise standout summary, so use the Engineer Pack for the detailed item list behind this grouped workstream.'
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
            'The current review did not surface a clear roadmap ordering, so use the Engineer Pack and grouped recommendations for the detailed crosswalk.'
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

function Convert-CustomerDocumentationReferencesToListItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$References = @()
    )

    return @(
        foreach ($reference in @($References)) {
            if ($null -eq $reference) {
                continue
            }

            $title = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $reference -Names @('Title')) -Default ''
            $url = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $reference -Names @('Url')) -Default ''
            $why = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $reference -Names @('WhyItIsRelevant')) -Default ''
            $segments = @()
            if (-not [string]::IsNullOrWhiteSpace($title)) {
                $segments += $title
            }
            if (-not [string]::IsNullOrWhiteSpace($url)) {
                $segments += $url
            }
            $line = $segments -join ' - '
            if (-not [string]::IsNullOrWhiteSpace($why)) {
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    $line += '. '
                }
                $line += ('Why it is relevant: {0}' -f $why)
            }

            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $line
            }
        }
    )
}

function Get-CustomerChartPaletteColors {
    [CmdletBinding()]
    param()

    Add-Type -AssemblyName System.Drawing
    return @(
        [System.Drawing.Color]::FromArgb(36, 93, 165),
        [System.Drawing.Color]::FromArgb(0, 146, 153),
        [System.Drawing.Color]::FromArgb(219, 120, 44),
        [System.Drawing.Color]::FromArgb(114, 81, 181),
        [System.Drawing.Color]::FromArgb(74, 125, 63),
        [System.Drawing.Color]::FromArgb(194, 67, 112)
    )
}

function Get-CustomerChartRenderableRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @()
    )

    return @(
        foreach ($row in @($Rows)) {
            if ($null -eq $row) {
                continue
            }

            $label = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $row -Names @('Label', 'Name', 'Workstream')) -Default ''
            $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $row -Names @('Value', 'Count', 'OpenFindings'))
            if ([string]::IsNullOrWhiteSpace($label) -or $null -eq $value -or $value -le 0) {
                continue
            }

            [pscustomobject]@{
                Label = $label
                Value = [double]$value
            }
        }
    )
}

function New-CustomerChartPngBytesFromBitmap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Drawing.Bitmap]$Bitmap
    )

    $stream = New-Object System.IO.MemoryStream
    try {
        $Bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        return $stream.ToArray()
    }
    finally {
        $stream.Dispose()
    }
}

function New-CustomerHorizontalBarChartPngBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $false)][int]$WidthPx = 720,
        [Parameter(Mandatory = $false)][int]$HeightPx = 260
    )

    $renderableRows = @(Get-CustomerChartRenderableRows -Rows $Rows)
    if ($renderableRows.Count -lt 2) {
        return $null
    }

    Add-Type -AssemblyName System.Drawing
    $colors = Get-CustomerChartPaletteColors
    $bitmap = [System.Drawing.Bitmap]::new($WidthPx, $HeightPx)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $labelFont = [System.Drawing.Font]::new('Segoe UI', 10)
    $valueFont = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $labelBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(52, 58, 64))
    $axisPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(197, 205, 214), 1)
    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::White)

        $topPadding = 20
        $leftPadding = 210
        $rightPadding = 56
        $rowSpacing = [Math]::Max([Math]::Floor(($HeightPx - ($topPadding * 2)) / $renderableRows.Count), 26)
        $barHeight = [Math]::Max($rowSpacing - 10, 14)
        $maxValue = ($renderableRows | Measure-Object -Property Value -Maximum).Maximum
        if ($null -eq $maxValue -or $maxValue -le 0) {
            return $null
        }

        $graphics.DrawLine($axisPen, $leftPadding, $topPadding - 4, $leftPadding, $HeightPx - $topPadding + 2)
        for ($index = 0; $index -lt $renderableRows.Count; $index++) {
            $row = $renderableRows[$index]
            $y = $topPadding + ($index * $rowSpacing)
            $barWidth = [Math]::Max([Math]::Round((($WidthPx - $leftPadding - $rightPadding) * ($row.Value / $maxValue))), 4)
            $barRect = [System.Drawing.RectangleF]::new($leftPadding, $y, $barWidth, $barHeight)
            $barBrush = [System.Drawing.SolidBrush]::new($colors[$index % $colors.Count])
            try {
                $graphics.DrawString($row.Label, $labelFont, $labelBrush, 8, $y - 1)
                $graphics.FillRectangle($barBrush, $barRect)
                $graphics.DrawString(([string][int][Math]::Round($row.Value, 0)), $valueFont, $labelBrush, ($leftPadding + $barWidth + 8), ($y - 1))
            }
            finally {
                $barBrush.Dispose()
            }
        }

        return New-CustomerChartPngBytesFromBitmap -Bitmap $bitmap
    }
    catch {
        return $null
    }
    finally {
        $axisPen.Dispose()
        $labelBrush.Dispose()
        $labelFont.Dispose()
        $valueFont.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function New-CustomerDonutChartPngBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $false)][int]$WidthPx = 720,
        [Parameter(Mandatory = $false)][int]$HeightPx = 250
    )

    $renderableRows = @(Get-CustomerChartRenderableRows -Rows $Rows)
    if ($renderableRows.Count -lt 2) {
        return $null
    }

    Add-Type -AssemblyName System.Drawing
    $colors = Get-CustomerChartPaletteColors
    $bitmap = [System.Drawing.Bitmap]::new($WidthPx, $HeightPx)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $labelFont = [System.Drawing.Font]::new('Segoe UI', 10)
    $smallFont = [System.Drawing.Font]::new('Segoe UI', 9)
    $boldFont = [System.Drawing.Font]::new('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $labelBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(52, 58, 64))
    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::White)

        $chartSize = [Math]::Min($HeightPx - 32, 190)
        $chartRect = [System.Drawing.Rectangle]::new(18, [Math]::Max([Math]::Floor(($HeightPx - $chartSize) / 2), 16), $chartSize, $chartSize)
        $innerSize = [Math]::Floor($chartSize * 0.56)
        $innerOffset = [Math]::Floor(($chartSize - $innerSize) / 2)
        $total = ($renderableRows | Measure-Object -Property Value -Sum).Sum
        if ($null -eq $total -or $total -le 0) {
            return $null
        }

        $startAngle = -90.0
        for ($index = 0; $index -lt $renderableRows.Count; $index++) {
            $row = $renderableRows[$index]
            $sweepAngle = [double](360.0 * ($row.Value / $total))
            $sliceBrush = [System.Drawing.SolidBrush]::new($colors[$index % $colors.Count])
            try {
                $graphics.FillPie($sliceBrush, $chartRect, [single]$startAngle, [single]$sweepAngle)
            }
            finally {
                $sliceBrush.Dispose()
            }
            $startAngle += $sweepAngle
        }

        $graphics.FillEllipse([System.Drawing.Brushes]::White, ($chartRect.X + $innerOffset), ($chartRect.Y + $innerOffset), $innerSize, $innerSize)
        $totalText = [string][int][Math]::Round($total, 0)
        $totalLabelSize = $graphics.MeasureString($totalText, $boldFont)
        $totalCaptionSize = $graphics.MeasureString('Total', $smallFont)
        $centerX = $chartRect.X + ($chartRect.Width / 2)
        $centerY = $chartRect.Y + ($chartRect.Height / 2)
        $graphics.DrawString($totalText, $boldFont, $labelBrush, ($centerX - ($totalLabelSize.Width / 2)), ($centerY - 16))
        $graphics.DrawString('Total', $smallFont, $labelBrush, ($centerX - ($totalCaptionSize.Width / 2)), ($centerY + 6))

        $legendX = $chartRect.Right + 28
        $legendY = $chartRect.Y + 2
        for ($index = 0; $index -lt $renderableRows.Count; $index++) {
            $row = $renderableRows[$index]
            $legendItemY = $legendY + ($index * 32)
            $legendBrush = [System.Drawing.SolidBrush]::new($colors[$index % $colors.Count])
            try {
                $graphics.FillRectangle($legendBrush, $legendX, $legendItemY + 4, 12, 12)
            }
            finally {
                $legendBrush.Dispose()
            }

            $percentageText = '{0:N1}%' -f (($row.Value / $total) * 100)
            $graphics.DrawString($row.Label, $labelFont, $labelBrush, ($legendX + 20), $legendItemY)
            $graphics.DrawString(($percentageText + ' (' + [int][Math]::Round($row.Value, 0) + ')'), $smallFont, $labelBrush, ($legendX + 20), ($legendItemY + 14))
        }

        return New-CustomerChartPngBytesFromBitmap -Bitmap $bitmap
    }
    catch {
        return $null
    }
    finally {
        $labelBrush.Dispose()
        $labelFont.Dispose()
        $smallFont.Dispose()
        $boldFont.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function New-CustomerChartImageBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('HorizontalBar', 'Donut')][string]$ChartType,
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $true)][string]$AltText,
        [Parameter(Mandatory = $false)][int]$WidthPx = 720,
        [Parameter(Mandatory = $false)][int]$HeightPx = 260
    )

    $imageBytes = switch ($ChartType) {
        'HorizontalBar' { New-CustomerHorizontalBarChartPngBytes -Rows $Rows -WidthPx $WidthPx -HeightPx $HeightPx }
        'Donut' { New-CustomerDonutChartPngBytes -Rows $Rows -WidthPx $WidthPx -HeightPx $HeightPx }
    }

    if ($null -eq $imageBytes -or $imageBytes.Length -eq 0) {
        return $null
    }

    return New-CustomerWordImageBlock -ImageBytes $imageBytes -AltText $AltText -WidthPx $WidthPx -HeightPx $HeightPx
}

function Get-CustomerOverallFindingsChartRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$WorkstreamSummaries = @(),
        [Parameter(Mandatory = $false)][object[]]$Findings = @()
    )

    return @(
        foreach ($group in @(
                $WorkstreamSummaries |
                    Where-Object { $null -ne $_ } |
                    Group-Object { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('Workstream', 'OwnerTeam')) -Default 'Not validated from the reviewed data' } |
                    Sort-Object `
                        @{ Expression = { ($_.Group | ForEach-Object { Get-CustomerSeverityRank -Severity (Get-ArrayaObjectValue -Object $_ -Names @('Severity')) } | Measure-Object -Maximum).Maximum }; Descending = $true }, `
                        @{ Expression = { ($_.Group | ForEach-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OpenFindings')) } | Where-Object { $null -ne $_ } | Measure-Object -Sum).Sum }; Descending = $true }, `
                        Name
            )) {
            $workstream = [string]$group.Name
            $openFindings = @(
                $group.Group |
                    ForEach-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OpenFindings')) } |
                    Where-Object { $null -ne $_ } |
                    Measure-Object -Sum
            ).Sum

            if ($null -eq $openFindings) {
                $openFindings = @($Findings | Where-Object { ([string]$_.OwnerTeam) -eq $workstream }).Count
            }

            if ([string]::IsNullOrWhiteSpace($workstream) -or $null -eq $openFindings -or $openFindings -le 0) {
                continue
            }

            [pscustomobject]@{
                Label = $workstream
                Value = $openFindings
            }
        }
    )
}

function Get-CustomerRecipientTypeChartRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$RecipientRows = @()
    )

    return @(
        @(
            $RecipientRows |
                Group-Object {
                    Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('RecipientTypeDetails')) -Default 'Unknown'
                } |
                Sort-Object Count -Descending
        ) |
            Select-Object -First 5 |
            ForEach-Object {
                [pscustomobject]@{
                    Label = $_.Name
                    Value = $_.Count
                }
            }
    )
}

function Get-CustomerRecipientDomainSummaryRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$RecipientRows = @(),
        [Parameter(Mandatory = $false)][int]$Top = 5
    )

    $domainSummary = @{}
    foreach ($recipientRow in @($RecipientRows)) {
        if ($null -eq $recipientRow) {
            continue
        }

        $primarySmtp = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $recipientRow -Names @('PrimarySmtpAddress', 'WindowsEmailAddress', 'UserPrincipalName')) -Default ''
        if (-not [string]::IsNullOrWhiteSpace($primarySmtp) -and $primarySmtp -match '@(?<domain>[^@\s>]+)$') {
            $domainName = $Matches['domain'].ToLowerInvariant()
            if (-not $domainSummary.ContainsKey($domainName)) {
                $domainSummary[$domainName] = [ordered]@{ Domain = $domainName; Primary = 0; AliasOnly = 0 }
            }
            $domainSummary[$domainName].Primary++
        }

        $emailAddresses = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $recipientRow -Names @('EmailAddresses'))
        foreach ($emailAddress in @($emailAddresses)) {
            $addressText = Convert-ToArrayaDisplayText -Value $emailAddress -Default ''
            if ([string]::IsNullOrWhiteSpace($addressText)) {
                continue
            }

            if ($addressText -match '^[A-Za-z]+:(?<value>.+)$') {
                $addressText = $Matches['value']
            }
            if ($addressText -notmatch '@(?<domain>[^@\s>]+)$') {
                continue
            }

            $domainName = $Matches['domain'].ToLowerInvariant()
            if (-not $domainSummary.ContainsKey($domainName)) {
                $domainSummary[$domainName] = [ordered]@{ Domain = $domainName; Primary = 0; AliasOnly = 0 }
            }

            if (-not [string]::IsNullOrWhiteSpace($primarySmtp) -and [string]::Equals($addressText, $primarySmtp, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            $domainSummary[$domainName].AliasOnly++
        }
    }

    return @(
        @($domainSummary.Values) |
            ForEach-Object {
                [pscustomobject]@{
                    Domain    = $_.Domain
                    Total     = ($_.Primary + $_.AliasOnly)
                    Primary   = $_.Primary
                    AliasOnly = $_.AliasOnly
                }
            } |
            Sort-Object Total, Primary -Descending |
            Select-Object -First $Top
    )
}

function Get-CustomerRecipientDomainChartRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$RecipientRows = @(),
        [Parameter(Mandatory = $false)][int]$Top = 6
    )

    return @(
        Get-CustomerRecipientDomainSummaryRows -RecipientRows $RecipientRows -Top $Top |
            ForEach-Object {
                [pscustomobject]@{
                    Label = $_.Domain
                    Value = [double]$_.Total
                }
            }
    )
}

function Get-CustomerExternalExposureCategoryChartRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$ExternalExposureFindings = @(),
        [Parameter(Mandatory = $false)][int]$Top = 6
    )

    return @(
        @(
            $ExternalExposureFindings |
                Where-Object {
                    -not [string]::IsNullOrWhiteSpace((Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ExposureCategory')) -Default ''))
                } |
                Group-Object {
                    Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ExposureCategory')) -Default 'Unknown'
                } |
                Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Descending = $false }
        ) |
            Select-Object -First $Top |
            ForEach-Object {
                [pscustomobject]@{
                    Label = $_.Name
                    Value = [double]$_.Count
                }
            }
    )
}

function Get-CustomerMfaMethodChartRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$RegistrationSummaryRecord,
        [Parameter(Mandatory = $false)]$EnrollmentSummaryRecord,
        [Parameter(Mandatory = $false)][int]$Top = 6
    )

    $methodCounts = New-Object 'System.Collections.Generic.List[object]'
    $methodCountSource = Get-ArrayaObjectValue -Object $RegistrationSummaryRecord -Names @('MethodCounts')

    if ($methodCountSource -is [System.Collections.IDictionary]) {
        foreach ($key in $methodCountSource.Keys) {
            $label = Convert-ToCustomerAssessmentMfaMethodBreakdownText -Value $key -Default ''
            $value = Convert-ArrayaToNumber $methodCountSource[$key]
            if (-not [string]::IsNullOrWhiteSpace($label) -and $null -ne $value -and $value -gt 0) {
                $methodCounts.Add([pscustomobject]@{
                        Label = $label
                        Value = [double]$value
                    }) | Out-Null
            }
        }
    }
    elseif ($null -ne $methodCountSource -and $methodCountSource.PSObject) {
        foreach ($property in $methodCountSource.PSObject.Properties) {
            $label = Convert-ToCustomerAssessmentMfaMethodBreakdownText -Value $property.Name -Default ''
            $value = Convert-ArrayaToNumber $property.Value
            if (-not [string]::IsNullOrWhiteSpace($label) -and $null -ne $value -and $value -gt 0) {
                $methodCounts.Add([pscustomobject]@{
                        Label = $label
                        Value = [double]$value
                    }) | Out-Null
            }
        }
    }

    if ($methodCounts.Count -eq 0) {
        $registeredMethodBreakdown = Convert-ToCustomerAssessmentMfaMethodBreakdownText -Value (Get-ArrayaObjectValue -Object $EnrollmentSummaryRecord -Names @('RegisteredMethodBreakdown')) -Default ''
        foreach ($segment in @($registeredMethodBreakdown -split ';')) {
            $trimmedSegment = ([string]$segment).Trim()
            if ([string]::IsNullOrWhiteSpace($trimmedSegment) -or $trimmedSegment -notmatch '^(?<label>.+?)\s*=\s*(?<value>-?\d+(?:\.\d+)?)$') {
                continue
            }

            $label = Convert-ToCustomerAssessmentMfaMethodBreakdownText -Value $Matches['label'] -Default ''
            $value = Convert-ArrayaToNumber $Matches['value']
            if (-not [string]::IsNullOrWhiteSpace($label) -and $null -ne $value -and $value -gt 0) {
                $methodCounts.Add([pscustomobject]@{
                        Label = $label
                        Value = [double]$value
                    }) | Out-Null
            }
        }
    }

    return @(
        @($methodCounts.ToArray()) |
            Group-Object Label |
            ForEach-Object {
                [pscustomobject]@{
                    Label = $_.Name
                    Value = [double](($_.Group | Measure-Object -Property Value -Sum).Sum)
                }
            } |
            Sort-Object Value, Label -Descending |
            Select-Object -First $Top
    )
}

function Test-CustomerMfaGapHasExclusionHit {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$GapRow)

    $gapCategory = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $GapRow -Names @('GapCategory')) -Default ''
    $gapReason = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $GapRow -Names @('GapReason')) -Default ''

    return ($gapCategory -match '(?i)\bexcluded\b' -or $gapReason -match '(?i)\bexcluded\b')
}

function Get-CustomerMessageActivityTableRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $true)][string]$CountPropertyName,
        [Parameter(Mandatory = $false)][int]$Top = 5
    )

    return @(
        @(
            $Rows |
                Where-Object {
                    $countValue = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @($CountPropertyName))
                    $null -ne $countValue -and $countValue -gt 0
                } |
                Sort-Object `
                    @{ Expression = { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @($CountPropertyName)) }; Descending = $true }, `
                    @{ Expression = { (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName', 'DisplayName')) -Default '').ToLowerInvariant() } }
        ) |
            Select-Object -First $Top |
            ForEach-Object {
                $activityCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @($CountPropertyName))
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Not surfaced in current source'),
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default 'Not surfaced in current source'),
                    $(if ($null -ne $activityCount) { [int64][Math]::Round([double]$activityCount, 0) } else { 'Not surfaced in current source' }),
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('LastActivityDate')) -Default 'Not surfaced in current source')
                )
            }
    )
}

function Get-CustomerAdminActivityChartRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$EnabledAdminRows = @(),
        [Parameter(Mandatory = $false)][int]$StaleThresholdDays = 180
    )

    if (@($EnabledAdminRows).Count -eq 0) {
        return @()
    }

    $staleAdmins = @(
        $EnabledAdminRows | Where-Object {
            $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime', 'SignInActivityLastSignInDateTime'))
            $null -eq $lastSignIn -or $lastSignIn -lt (Get-Date).AddDays(-1 * [Math]::Abs($StaleThresholdDays))
        }
    ).Count
    $recentAdmins = [Math]::Max((@($EnabledAdminRows).Count - $staleAdmins), 0)

    return @(
        [pscustomobject]@{ Label = 'Recent admins'; Value = [double]$recentAdmins }
        [pscustomobject]@{ Label = "Stale admins (>$StaleThresholdDays days or no sign-in)"; Value = [double]$staleAdmins }
    )
}

function Get-CustomerDataFootprintChartRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$PrimaryMailboxStatsRows = @(),
        [Parameter(Mandatory = $false)][object[]]$ArchiveMailboxStatsRows = @(),
        [Parameter(Mandatory = $false)][object[]]$SharePointRows = @(),
        [Parameter(Mandatory = $false)][object[]]$OneDriveRows = @()
    )

    $sumBytes = {
        param([object[]]$Rows)

        $total = 0.0
        foreach ($row in @($Rows)) {
            $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $row -Names @('TotalItemSizeBytes'))
            if ($null -ne $value -and $value -gt 0) {
                $total += [double]$value
            }
        }

        return ($total / 1GB)
    }

    $sumSiteGb = {
        param([object[]]$Rows)

        $total = 0.0
        foreach ($row in @($Rows)) {
            $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $row -Names @('StorageUsedGB', 'StorageUsageCurrent'))
            if ($null -ne $value -and $value -gt 0) {
                $total += [double]$value
            }
        }

        return $total
    }

    $exchangeMailboxRows = @(
        $PrimaryMailboxStatsRows | Where-Object {
            (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('MailboxType')) -Default '') -ne 'GroupMailbox'
        }
    )
    $groupMailboxRows = @(
        $PrimaryMailboxStatsRows | Where-Object {
            (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('MailboxType')) -Default '') -eq 'GroupMailbox'
        }
    )
    $teamConnectedSharePointRows = @(
        $SharePointRows | Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsTeamsConnected'))) -eq $true
        }
    )
    $standaloneSharePointRows = @(
        $SharePointRows | Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsTeamsConnected'))) -ne $true
        }
    )

    return @(
        [pscustomobject]@{ Label = 'Exchange Mailboxes'; Value = [double](& $sumBytes $exchangeMailboxRows) }
        [pscustomobject]@{ Label = 'Archive Mailboxes'; Value = [double](& $sumBytes $ArchiveMailboxStatsRows) }
        [pscustomobject]@{ Label = 'Unified Group Mailboxes'; Value = [double](& $sumBytes $groupMailboxRows) }
        [pscustomobject]@{ Label = 'Team-Connected SharePoint Sites'; Value = [double](& $sumSiteGb $teamConnectedSharePointRows) }
        [pscustomobject]@{ Label = 'Standalone SharePoint Sites'; Value = [double](& $sumSiteGb $standaloneSharePointRows) }
        [pscustomobject]@{ Label = 'OneDrive'; Value = [double](& $sumSiteGb $OneDriveRows) }
    )
}

function Get-CustomerExecutiveRiskBulletItems {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object[]]$Rows = @())

    return @(
        foreach ($row in @($Rows)) {
            $cells = @(Get-CustomerTableRowCells -Row $row)
            if ($cells.Count -lt 3) {
                continue
            }

            $cluster = Convert-ToCustomerAssessmentDisplayText -Value $cells[0] -Default ''
            $standout = Convert-ToCustomerAssessmentDisplayText -Value $cells[1] -Default ''
            $care = Convert-ToCustomerAssessmentDisplayText -Value $cells[2] -Default ''
            if ([string]::IsNullOrWhiteSpace($cluster)) {
                continue
            }

            $standoutLine = Get-CustomerCompactNarrativeLine -Text $standout -MaxSentences 1 -MaxLength 120
            $careLine = Get-CustomerCompactNarrativeLine -Text $care -MaxSentences 1 -MaxLength 110

            if (-not [string]::IsNullOrWhiteSpace($standoutLine) -and -not [string]::IsNullOrWhiteSpace($careLine)) {
                '{0}: {1} Leadership impact: {2}' -f $cluster, $standoutLine, $careLine
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($standoutLine)) {
                '{0}: {1}' -f $cluster, $standoutLine
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($careLine)) {
                '{0}: Leadership impact: {1}' -f $cluster, $careLine
            }
        }
    )
}

function Get-CustomerLeadershipDecisionBulletItems {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object[]]$Rows = @())

    return @(
        foreach ($row in @($Rows)) {
            $cells = @(Get-CustomerTableRowCells -Row $row)
            if ($cells.Count -lt 3) {
                continue
            }

            $focus = Convert-ToCustomerAssessmentDisplayText -Value $cells[0] -Default ''
            $next = Convert-ToCustomerAssessmentDisplayText -Value $cells[1] -Default ''
            $whyNow = Convert-ToCustomerAssessmentDisplayText -Value $cells[2] -Default ''
            if ([string]::IsNullOrWhiteSpace($focus)) {
                continue
            }

            $nextLine = Get-CustomerCompactNarrativeLine -Text $next -MaxSentences 1 -MaxLength 120
            $whyNowLine = Get-CustomerCompactNarrativeLine -Text $whyNow -MaxSentences 1 -MaxLength 110

            if (-not [string]::IsNullOrWhiteSpace($nextLine) -and -not [string]::IsNullOrWhiteSpace($whyNowLine)) {
                '{0}: {1} Why now: {2}' -f $focus, $nextLine, $whyNowLine
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($nextLine)) {
                '{0}: {1}' -f $focus, $nextLine
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($whyNowLine)) {
                '{0}: Why now: {1}' -f $focus, $whyNowLine
            }
        }
    )
}

function New-CustomerExecutiveRiskBlocks {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object[]]$Rows = @())

    $blocks = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($Rows)) {
        $cells = @(Get-CustomerTableRowCells -Row $row)
        if ($cells.Count -lt 3) {
            continue
        }

        $cluster = Convert-ToCustomerAssessmentDisplayText -Value $cells[0] -Default ''
        $standout = Convert-ToCustomerAssessmentDisplayText -Value $cells[1] -Default ''
        $care = Convert-ToCustomerAssessmentDisplayText -Value $cells[2] -Default ''
        if ([string]::IsNullOrWhiteSpace($cluster)) {
            continue
        }

        $blocks.Add((New-CustomerWordParagraphBlock -Text $cluster -Style 'Heading3')) | Out-Null
        $standoutItems = @(Get-CustomerRoadmapNarrativeSegments -Text $standout -MaxSentences 2 -MaxLength 155)
        if ($standoutItems.Count -gt 0) {
            $blocks.Add((New-CustomerWordParagraphBlock -Text 'What stands out' -Style 'Heading4')) | Out-Null
            $blocks.Add((New-CustomerWordListBlock -Items $standoutItems)) | Out-Null
        }

        $careItems = @(Get-CustomerRoadmapNarrativeSegments -Text $care -MaxSentences 2 -MaxLength 155)
        if ($careItems.Count -gt 0) {
            $blocks.Add((New-CustomerWordParagraphBlock -Text 'Leadership impact' -Style 'Heading4')) | Out-Null
            $blocks.Add((New-CustomerWordListBlock -Items $careItems)) | Out-Null
        }
    }

    return @($blocks.ToArray())
}

function New-CustomerLeadershipDecisionBlocks {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object[]]$Rows = @())

    $blocks = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($Rows)) {
        $cells = @(Get-CustomerTableRowCells -Row $row)
        if ($cells.Count -lt 3) {
            continue
        }

        $focus = Convert-ToCustomerAssessmentDisplayText -Value $cells[0] -Default ''
        $next = Convert-ToCustomerAssessmentDisplayText -Value $cells[1] -Default ''
        $whyNow = Convert-ToCustomerAssessmentDisplayText -Value $cells[2] -Default ''
        if ([string]::IsNullOrWhiteSpace($focus)) {
            continue
        }

        $blocks.Add((New-CustomerWordParagraphBlock -Text $focus -Style 'Heading3')) | Out-Null
        $nextItems = @(Get-CustomerRoadmapNarrativeSegments -Text $next -MaxSentences 2 -MaxLength 155)
        if ($nextItems.Count -gt 0) {
            $blocks.Add((New-CustomerWordParagraphBlock -Text 'Decision needed' -Style 'Heading4')) | Out-Null
            $blocks.Add((New-CustomerWordListBlock -Items $nextItems)) | Out-Null
        }

        $whyNowItems = @(Get-CustomerRoadmapNarrativeSegments -Text $whyNow -MaxSentences 2 -MaxLength 155)
        if ($whyNowItems.Count -gt 0) {
            $blocks.Add((New-CustomerWordParagraphBlock -Text 'Why now' -Style 'Heading4')) | Out-Null
            $blocks.Add((New-CustomerWordListBlock -Items $whyNowItems)) | Out-Null
        }
    }

    return @($blocks.ToArray())
}

function Get-CustomerRoadmapBucketHeading {
    [CmdletBinding()]
    param([AllowNull()][string]$RoadmapPhase)

    switch (([string]$RoadmapPhase).Trim()) {
        'Immediate' { return '0-30 Days (Foundation)' }
        'Near Term' { return '31-60 Days (Enforcement & Cleanup)' }
        'Planned' { return '61-90 Days (Stabilization)' }
        default { return 'Operational Model' }
    }
}

function Get-CustomerDocumentRevisionLabel {
    [CmdletBinding()]
    param()

    return '1.0'
}

function Get-CustomerRoadmapDocumentInfoRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt,
        [Parameter(Mandatory = $false)][string]$DocumentRevision
    )

    return @(
        New-CustomerWordTableRow -Cells @('Client', $TenantName)
        New-CustomerWordTableRow -Cells @('Document Type', 'Microsoft 365 Remediation Roadmap')
        New-CustomerWordTableRow -Cells @('Document Version', $(if ([string]::IsNullOrWhiteSpace($DocumentRevision)) { (Get-CustomerDocumentRevisionLabel) } else { $DocumentRevision }))
        New-CustomerWordTableRow -Cells @('Prepared By', 'Arraya Solutions')
        New-CustomerWordTableRow -Cells @('Generated Date', $GeneratedAt.ToString('yyyy-MM-dd'))
        New-CustomerWordTableRow -Cells @('Source of Detail', 'Use the Engineer Pack and support JSON for detailed technical evidence and validation paths.')
    )
}

function Get-CustomerRoadmapEnvironmentReviewText {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Signals)

    $userCount = @(Convert-ArrayaObjectToArray $Signals.Users).Count
    $adminCount = @(Convert-ArrayaObjectToArray $Signals.Admins).Count
    $deviceCount = @(Convert-ArrayaObjectToArray $Signals.DeviceDetails).Count
    $mailboxCount = @(Convert-ArrayaObjectToArray $Signals.AllMailboxes).Count
    $teamCount = @(Convert-ArrayaObjectToArray $Signals.AllTeams).Count
    $sharePointCount = @(Convert-ArrayaObjectToArray $Signals.SharePoint).Count
    $oneDriveCount = @(Convert-ArrayaObjectToArray $Signals.OneDrive).Count
    $domainCount = @(Convert-ArrayaObjectToArray $Signals.Domains).Count
    $licenseCount = @(Convert-ArrayaObjectToArray $Signals.LicenseSKUs).Count

    return ('The reviewed tenant snapshot surfaced {0} user account(s), {1} admin account(s), {2} device record(s), {3} mailbox(es), {4} Team(s), {5} SharePoint site(s), {6} OneDrive location(s), {7} accepted domain(s), and {8} license SKU record(s). The roadmap below focuses on the areas where the current review showed repeated control drift, concentrated exposure, or operational cleanup that now needs an accountable execution path.' -f $userCount, $adminCount, $deviceCount, $mailboxCount, $teamCount, $sharePointCount, $oneDriveCount, $domainCount, $licenseCount)
}

function Get-CustomerRoadmapEnvironmentReviewBulletItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Signals,
        [Parameter(Mandatory = $false)]$ExecutiveDecisionSummary
    )

    $userCount = @(Convert-ArrayaObjectToArray $Signals.Users).Count
    $adminCount = @(Convert-ArrayaObjectToArray $Signals.Admins).Count
    $deviceCount = @(Convert-ArrayaObjectToArray $Signals.DeviceDetails).Count
    $mailboxCount = @(Convert-ArrayaObjectToArray $Signals.AllMailboxes).Count
    $teamCount = @(Convert-ArrayaObjectToArray $Signals.AllTeams).Count
    $sharePointCount = @(Convert-ArrayaObjectToArray $Signals.SharePoint).Count
    $oneDriveCount = @(Convert-ArrayaObjectToArray $Signals.OneDrive).Count
    $domainCount = @(Convert-ArrayaObjectToArray $Signals.Domains).Count
    $licenseCount = @(Convert-ArrayaObjectToArray $Signals.LicenseSKUs).Count

    $items = New-Object System.Collections.Generic.List[string]
    $items.Add(('Reviewed footprint: {0} user account(s), {1} admin account(s), and {2} device record(s) were included in the reviewed snapshot.' -f $userCount, $adminCount, $deviceCount)) | Out-Null
    $items.Add(('Reviewed workloads: {0} mailbox(es), {1} Team(s), {2} SharePoint site(s), and {3} OneDrive location(s) were assessed for this roadmap.' -f $mailboxCount, $teamCount, $sharePointCount, $oneDriveCount)) | Out-Null
    $items.Add(('Tenant context: {0} accepted domain(s) and {1} license SKU record(s) were part of the reviewed baseline.' -f $domainCount, $licenseCount)) | Out-Null

    if ($null -ne $ExecutiveDecisionSummary -and -not [string]::IsNullOrWhiteSpace([string]$ExecutiveDecisionSummary.Narrative)) {
        $leadershipFraming = Get-CustomerCompactNarrativeLine -Text ([string]$ExecutiveDecisionSummary.Narrative) -MaxSentences 2 -MaxLength 150
        if (-not [string]::IsNullOrWhiteSpace($leadershipFraming)) {
            $items.Add(('Leadership framing: {0}' -f $leadershipFraming)) | Out-Null
        }
    }

    $items.Add('How to use this roadmap: use it for executive sequencing first, then use the Engineer Pack for the technical evidence and validation path behind each grouped action.') | Out-Null
    return @($items.ToArray() | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function New-CustomerRoadmapEnvironmentReviewBlocks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Signals,
        [Parameter(Mandatory = $false)]$ExecutiveDecisionSummary
    )

    $userCount = @(Convert-ArrayaObjectToArray $Signals.Users).Count
    $adminCount = @(Convert-ArrayaObjectToArray $Signals.Admins).Count
    $deviceCount = @(Convert-ArrayaObjectToArray $Signals.DeviceDetails).Count
    $mailboxCount = @(Convert-ArrayaObjectToArray $Signals.AllMailboxes).Count
    $teamCount = @(Convert-ArrayaObjectToArray $Signals.AllTeams).Count
    $sharePointCount = @(Convert-ArrayaObjectToArray $Signals.SharePoint).Count
    $oneDriveCount = @(Convert-ArrayaObjectToArray $Signals.OneDrive).Count
    $domainCount = @(Convert-ArrayaObjectToArray $Signals.Domains).Count
    $licenseCount = @(Convert-ArrayaObjectToArray $Signals.LicenseSKUs).Count

    $blocks = New-Object System.Collections.Generic.List[object]
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Reviewed footprint' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(
                ('Identity: {0} user account(s) and {1} admin account(s).' -f $userCount, $adminCount),
                ('Endpoint: {0} device record(s).' -f $deviceCount),
                ('Messaging: {0} mailbox(es).' -f $mailboxCount),
                ('Collaboration: {0} Team(s), {1} SharePoint site(s), and {2} OneDrive location(s).' -f $teamCount, $sharePointCount, $oneDriveCount),
                ('Tenant context: {0} accepted domain(s) and {1} license SKU record(s).' -f $domainCount, $licenseCount)
            ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Roadmap focus' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(
                'Prioritize repeated control drift, concentrated exposure, and operational cleanup with clear owners.',
                'Use this document for executive sequencing before stepping into technical validation.'
            ))) | Out-Null

    if ($null -ne $ExecutiveDecisionSummary -and -not [string]::IsNullOrWhiteSpace([string]$ExecutiveDecisionSummary.Narrative)) {
        $leadershipItems = @(Get-CustomerRoadmapNarrativeSegments -Text ([string]$ExecutiveDecisionSummary.Narrative) -MaxSentences 3 -MaxLength 155)
        if ($leadershipItems.Count -gt 0) {
            $blocks.Add((New-CustomerWordParagraphBlock -Text 'Leadership framing' -Style 'Heading3')) | Out-Null
            $blocks.Add((New-CustomerWordListBlock -Items $leadershipItems)) | Out-Null
        }
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'How to use this roadmap' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(
                'Confirm sequencing and ownership in this roadmap.',
                'Use the Engineer Pack for evidence, validation paths, and implementation detail behind each grouped action.'
            ))) | Out-Null

    return @($blocks.ToArray())
}

function Get-CustomerRoadmapSectionNarrative {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Observation,
        [Parameter(Mandatory = $false)]$ConsultativeSummary,
        [Parameter(Mandatory = $true)][string]$Fallback
    )

    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in @(
            $(if ($null -ne $Observation) { Convert-ToCustomerAssessmentNarrativeText -Text ([string]$Observation.ObservedNarrative) } else { $null }),
            $(if ($null -ne $ConsultativeSummary) { Convert-ToCustomerAssessmentNarrativeText -Text ([string]$ConsultativeSummary.Narrative) } else { $null }),
            $(if ($null -ne $Observation) { Convert-ToCustomerAssessmentNarrativeText -Text ([string]$Observation.WhyItMatters) } else { $null }),
            $(if ($null -ne $ConsultativeSummary) { Convert-ToCustomerAssessmentNarrativeText -Text ([string]$ConsultativeSummary.RecommendationSupport) } else { $null })
        )) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        if ($segments -notcontains $candidate) {
            $segments.Add($candidate) | Out-Null
        }
    }

    if ($segments.Count -eq 0) {
        return $Fallback
    }

    return (($segments.ToArray() | Select-Object -First 3) -join ' ')
}

function Get-CustomerRoadmapSectionBulletItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Observation,
        [Parameter(Mandatory = $false)]$ConsultativeSummary,
        [Parameter(Mandatory = $true)][string]$Fallback
    )

    $items = New-Object System.Collections.Generic.List[string]

    $currentState = if ($null -ne $Observation -and -not [string]::IsNullOrWhiteSpace([string]$Observation.ObservedNarrative)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$Observation.ObservedNarrative)
    }
    elseif ($null -ne $ConsultativeSummary -and -not [string]::IsNullOrWhiteSpace([string]$ConsultativeSummary.Narrative)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$ConsultativeSummary.Narrative)
    }
    else {
        $null
    }

    $whyItMatters = if ($null -ne $Observation -and -not [string]::IsNullOrWhiteSpace([string]$Observation.WhyItMatters)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$Observation.WhyItMatters)
    }
    else {
        $null
    }

    $positiveSignal = if ($null -ne $Observation -and -not [string]::IsNullOrWhiteSpace([string]$Observation.PositiveNarrative)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$Observation.PositiveNarrative)
    }
    else {
        $null
    }

    $recommendationFocus = if ($null -ne $ConsultativeSummary -and -not [string]::IsNullOrWhiteSpace([string]$ConsultativeSummary.RecommendationSupport)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$ConsultativeSummary.RecommendationSupport)
    }
    else {
        $null
    }

    foreach ($item in @(Get-CustomerLabeledNarrativeBulletItems -Label 'Current state' -ContinuationLabel 'Current state detail' -Text $currentState -MaxSentences 2 -MaxLength 180)) {
        $items.Add($item) | Out-Null
    }
    foreach ($item in @(Get-CustomerLabeledNarrativeBulletItems -Label 'Why it matters' -Text $whyItMatters -MaxSentences 1 -MaxLength 170)) {
        $items.Add($item) | Out-Null
    }
    foreach ($item in @(Get-CustomerLabeledNarrativeBulletItems -Label 'Positive signal' -Text $positiveSignal -MaxSentences 1 -MaxLength 170)) {
        $items.Add($item) | Out-Null
    }
    foreach ($item in @(Get-CustomerLabeledNarrativeBulletItems -Label 'Recommendation focus' -Text $recommendationFocus -MaxSentences 1 -MaxLength 170)) {
        if ($items -notcontains $item) {
            $items.Add($item) | Out-Null
        }
    }

    if ($items.Count -eq 0) {
        $fallbackLine = Get-CustomerCompactNarrativeLine -Text $Fallback -MaxSentences 1 -MaxLength 180
        $items.Add($(if ([string]::IsNullOrWhiteSpace($fallbackLine)) { $Fallback } else { $fallbackLine })) | Out-Null
    }

    return @($items.ToArray())
}

function New-CustomerRoadmapNarrativeBlocks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Heading,
        [AllowNull()][string]$Text,
        [Parameter(Mandatory = $false)][int]$MaxSentences = 3,
        [Parameter(Mandatory = $false)][int]$MaxLength = 155
    )

    $items = @(Get-CustomerRoadmapNarrativeSegments -Text $Text -MaxSentences $MaxSentences -MaxLength $MaxLength)
    if ($items.Count -eq 0) {
        return @()
    }

    return @(
        New-CustomerWordParagraphBlock -Text $Heading -Style 'Heading3'
        New-CustomerWordListBlock -Items $items
    )
}

function New-CustomerRoadmapSectionBlocks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Observation,
        [Parameter(Mandatory = $false)]$ConsultativeSummary,
        [Parameter(Mandatory = $true)][string]$Fallback
    )

    $currentState = if ($null -ne $Observation -and -not [string]::IsNullOrWhiteSpace([string]$Observation.ObservedNarrative)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$Observation.ObservedNarrative)
    }
    elseif ($null -ne $ConsultativeSummary -and -not [string]::IsNullOrWhiteSpace([string]$ConsultativeSummary.Narrative)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$ConsultativeSummary.Narrative)
    }
    else {
        $null
    }

    $whyItMatters = if ($null -ne $Observation -and -not [string]::IsNullOrWhiteSpace([string]$Observation.WhyItMatters)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$Observation.WhyItMatters)
    }
    else {
        $null
    }

    $positiveSignal = if ($null -ne $Observation -and -not [string]::IsNullOrWhiteSpace([string]$Observation.PositiveNarrative)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$Observation.PositiveNarrative)
    }
    else {
        $null
    }

    $recommendationFocus = if ($null -ne $ConsultativeSummary -and -not [string]::IsNullOrWhiteSpace([string]$ConsultativeSummary.RecommendationSupport)) {
        Convert-ToCustomerAssessmentNarrativeText -Text ([string]$ConsultativeSummary.RecommendationSupport)
    }
    else {
        $null
    }

    $blocks = New-Object System.Collections.Generic.List[object]
    foreach ($block in @(New-CustomerRoadmapNarrativeBlocks -Heading 'Current state' -Text $currentState -MaxSentences 3 -MaxLength 155)) {
        $blocks.Add($block) | Out-Null
    }
    foreach ($block in @(New-CustomerRoadmapNarrativeBlocks -Heading 'Why it matters' -Text $whyItMatters -MaxSentences 2 -MaxLength 155)) {
        $blocks.Add($block) | Out-Null
    }
    foreach ($block in @(New-CustomerRoadmapNarrativeBlocks -Heading 'Positive signal' -Text $positiveSignal -MaxSentences 2 -MaxLength 155)) {
        $blocks.Add($block) | Out-Null
    }
    foreach ($block in @(New-CustomerRoadmapNarrativeBlocks -Heading 'Recommendation focus' -Text $recommendationFocus -MaxSentences 2 -MaxLength 155)) {
        $blocks.Add($block) | Out-Null
    }

    if ($blocks.Count -eq 0) {
        $fallbackItems = @(Get-CustomerRoadmapNarrativeSegments -Text $Fallback -MaxSentences 2 -MaxLength 155)
        if ($fallbackItems.Count -gt 0) {
            $blocks.Add((New-CustomerWordListBlock -Items $fallbackItems)) | Out-Null
        }
    }

    return @($blocks.ToArray())
}

function New-CustomerRoadmapActionBlocks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$RoadmapActions = @(),
        [Parameter(Mandatory = $true)][string]$BucketHeading
    )

    $blocks = New-Object System.Collections.Generic.List[object]
    foreach ($action in @($RoadmapActions | Where-Object { (Get-CustomerRoadmapBucketHeading -RoadmapPhase ([string]$_.RoadmapPhase)) -eq $BucketHeading })) {
            $title = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $action -Names @('ActionTitle')) -Default 'Priority work item'
            $phase = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $action -Names @('RoadmapPhase')) -Default 'Monitor'
            $criticality = Get-CustomerActionCriticalityLabel -Severity ([string](Get-ArrayaObjectValue -Object $action -Names @('HighestSeverity')))
            $estimatedPsHours = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $action -Names @('EstimatedPsHours')) -Default 'Not validated from the reviewed data'
            $owner = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $action -Names @('PrimaryOwner')) -Default 'Shared operational owner'
            $nextStep = [string](Get-ArrayaObjectValue -Object $action -Names @('RecommendedNextStep'))
            $whyItMatters = [string](Get-ArrayaObjectValue -Object $action -Names @('WhyItMatters'))
            $successCheck = [string](Get-ArrayaObjectValue -Object $action -Names @('SuccessCheck'))

            $detailItems = New-Object System.Collections.Generic.List[string]
            $detailItems.Add(('Phase label: {0}' -f $phase)) | Out-Null
            $detailItems.Add(('Criticality: {0}' -f $criticality)) | Out-Null
            $detailItems.Add(('Rough PS Hours: {0}' -f $estimatedPsHours)) | Out-Null
            $detailItems.Add(('Owner: {0}' -f $owner)) | Out-Null

            foreach ($item in @(Get-CustomerLabeledNarrativeBulletItems -Label 'Next step' -Text $nextStep -MaxSentences 1 -MaxLength 170)) {
                $detailItems.Add($item) | Out-Null
            }
            foreach ($item in @(Get-CustomerLabeledNarrativeBulletItems -Label 'Why it matters' -Text $whyItMatters -MaxSentences 1 -MaxLength 170)) {
                $detailItems.Add($item) | Out-Null
            }
            foreach ($item in @(Get-CustomerLabeledNarrativeBulletItems -Label 'Success signal' -Text $successCheck -MaxSentences 1 -MaxLength 170)) {
                $detailItems.Add($item) | Out-Null
            }

            $blocks.Add((New-CustomerWordParagraphBlock -Text $title -Style 'Heading3')) | Out-Null
            $blocks.Add((New-CustomerWordListBlock -Items @($detailItems.ToArray()))) | Out-Null
        }

    return @($blocks.ToArray())
}

function New-RoadmapRemediationDocumentBlocks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$SourceModel,
        [Parameter(Mandatory = $true)]$Signals,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt
    )

    $tenantName = [string]$SourceModel.TenantName
    $documentRevision = Get-CustomerDocumentRevisionLabel
    $roadmapActions = @(Get-CustomerDistinctRoadmapActions -RoadmapActions $SourceModel.RoadmapActions)
    $executiveDecisionSummary = $SourceModel.ExecutiveDecisionSummary
    $riskRows = if ($null -ne $executiveDecisionSummary -and @($executiveDecisionSummary.RiskRows).Count -gt 0) {
        Convert-CustomerThreeColumnRowsToWordTableRows -Rows @($executiveDecisionSummary.RiskRows) -PropertyNames @('RiskCluster', 'WhatStandsOut', 'WhyLeadershipShouldCare')
    }
    else {
        @()
    }
    $decisionRows = if ($null -ne $executiveDecisionSummary -and @($executiveDecisionSummary.DecisionRows).Count -gt 0) {
        Convert-CustomerThreeColumnRowsToWordTableRows -Rows @($executiveDecisionSummary.DecisionRows) -PropertyNames @('DecisionFocus', 'WhatShouldHappenNext', 'WhyNow')
    }
    else {
        @()
    }

    $identityObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Identity & Access (Entra ID)'
    $collaborationObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Collaboration (Teams, SharePoint, OneDrive)'
    $endpointObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Devices & Endpoint Management'
    $messagingObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Messaging (Exchange Online)'
    $governanceObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Data Protection & Governance'
    $lifecycleObservation = Get-CustomerTechnicalObservationByTitle -TechnicalObservations $SourceModel.TechnicalObservations -Title 'Offboarding & Lifecycle Management'

    $keyHighlightBlocks = @(New-CustomerExecutiveRiskBlocks -Rows $riskRows)
    $leadershipDecisionBlocks = @(New-CustomerLeadershipDecisionBlocks -Rows $decisionRows)

    $blocks = New-Object System.Collections.Generic.List[object]
    $blocks.Add((New-CustomerWordParagraphBlock -Text "$tenantName Microsoft 365 Remediation Roadmap" -Style 'Title')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Prepared by: Arraya Solutions' -Style 'Subtitle')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Generated: {0}" -f $GeneratedAt.ToString('yyyy-MM-dd')) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Document revision: {0}" -f $documentRevision) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Version Control' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Date', 'Version', 'Prepared By', 'Description') -Rows @(
        (New-CustomerWordTableRow -Cells @($GeneratedAt.ToString('yyyy-MM-dd'), $documentRevision, 'Arraya Solutions', 'Initial companion remediation roadmap generated from the reviewed tenant data.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Document Information' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Field', 'Value') -Rows (Get-CustomerRoadmapDocumentInfoRows -TenantName $tenantName -GeneratedAt $GeneratedAt -DocumentRevision $documentRevision))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Executive Summary' -Style 'Heading1')) | Out-Null
    $executiveSummaryNarrative = if ($null -ne $executiveDecisionSummary -and -not [string]::IsNullOrWhiteSpace([string]$executiveDecisionSummary.Narrative)) {
        [string]$executiveDecisionSummary.Narrative
    }
    else {
        [string]$SourceModel.ExecutiveNarrative
    }
    $compactExecutiveSummary = Get-CustomerCompactNarrativeLine -Text $executiveSummaryNarrative -MaxSentences 2 -MaxLength 180
    if ([string]::IsNullOrWhiteSpace($compactExecutiveSummary)) {
        $compactExecutiveSummary = 'This roadmap summarizes the highest-value Microsoft 365 remediation themes surfaced from the reviewed tenant data.'
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text $compactExecutiveSummary -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Key Highlights' -Style 'Heading2')) | Out-Null
    if ($keyHighlightBlocks.Count -gt 0) {
        foreach ($block in $keyHighlightBlocks) {
            $blocks.Add($block) | Out-Null
        }
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The reviewed data did not surface distinct executive highlights beyond the grouped recommendations.' -Style 'Normal')) | Out-Null
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Current State Analysis' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Environment Review' -Style 'Heading2')) | Out-Null
    foreach ($block in @(New-CustomerRoadmapEnvironmentReviewBlocks -Signals $Signals -ExecutiveDecisionSummary $executiveDecisionSummary)) {
        $blocks.Add($block) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Identity & Access' -Style 'Heading2')) | Out-Null
    foreach ($block in @(New-CustomerRoadmapSectionBlocks -Observation $identityObservation -ConsultativeSummary $SourceModel.IdentityConsultativeSummary -Fallback 'Identity and access observations were not validated clearly enough in the reviewed data to generate a stronger roadmap narrative.')) {
        $blocks.Add($block) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Collaboration & Lifecycle' -Style 'Heading2')) | Out-Null
    foreach ($block in @(New-CustomerRoadmapSectionBlocks -Observation $collaborationObservation -ConsultativeSummary $SourceModel.CollaborationConsultativeSummary -Fallback 'Collaboration and lifecycle observations were not validated clearly enough in the reviewed data to generate a stronger roadmap narrative.')) {
        $blocks.Add($block) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Endpoint & Device Management' -Style 'Heading2')) | Out-Null
    foreach ($block in @(New-CustomerRoadmapSectionBlocks -Observation $endpointObservation -ConsultativeSummary $null -Fallback 'Endpoint and device-management observations were not validated clearly enough in the reviewed data to generate a stronger roadmap narrative.')) {
        $blocks.Add($block) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Messaging & Security' -Style 'Heading2')) | Out-Null
    foreach ($block in @(New-CustomerRoadmapSectionBlocks -Observation $messagingObservation -ConsultativeSummary $SourceModel.MessagingConsultativeSummary -Fallback 'Messaging and security observations were not validated clearly enough in the reviewed data to generate a stronger roadmap narrative.')) {
        $blocks.Add($block) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Governance & Licensing' -Style 'Heading2')) | Out-Null
    foreach ($block in @(New-CustomerRoadmapSectionBlocks -Observation $governanceObservation -ConsultativeSummary $SourceModel.GovernanceConsultativeSummary -Fallback 'Governance and licensing observations were not validated clearly enough in the reviewed data to generate a stronger roadmap narrative.')) {
        $blocks.Add($block) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'What This Means' -Style 'Heading2')) | Out-Null
    foreach ($block in @(New-CustomerRoadmapSectionBlocks -Observation $lifecycleObservation -ConsultativeSummary $SourceModel.LifecycleConsultativeSummary -Fallback 'The reviewed data shows repeated control drift across the same operating areas that already carry the most day-to-day support load, so the roadmap emphasizes ownership, enforcement, and cleanup sequencing rather than isolated one-off fixes.')) {
        $blocks.Add($block) | Out-Null
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Solution Approach' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(
        'Sequence the highest-value controls first and assign accountable owners early so remediation does not stall between workstreams.',
        'Start with the actions carrying the highest current impact or the broadest control drift.',
        'Use the grouped recommendations to confirm delivery sequence, then validate implementation detail from the Engineer Pack before execution.',
        'Treat cross-workload governance and lifecycle actions as part of the same operating model rather than as separate cleanup tracks.'
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Implementation Approach' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(
        '0-30 Days (Foundation): establish the baseline owner decisions and control changes that unblock the rest of the plan.',
        '31-60 Days (Enforcement & Cleanup): move the approved controls into enforcement and work through the highest-value cleanup items.',
        '61-90 Days (Stabilization): close the remaining medium-priority gaps and confirm the new baseline is operating as intended.',
        'Operational Model: keep the ongoing governance, monitoring, and lifecycle work visible after the initial remediation window.'
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Remediation Roadmap' -Style 'Heading1')) | Out-Null
    foreach ($bucketHeading in @('0-30 Days (Foundation)', '31-60 Days (Enforcement & Cleanup)', '61-90 Days (Stabilization)', 'Operational Model')) {
        $bucketBlocks = @(New-CustomerRoadmapActionBlocks -RoadmapActions $roadmapActions -BucketHeading $bucketHeading)
        $blocks.Add((New-CustomerWordParagraphBlock -Text $bucketHeading -Style 'Heading2')) | Out-Null
        if ($bucketBlocks.Count -gt 0) {
            foreach ($block in $bucketBlocks) {
                $blocks.Add($block) | Out-Null
            }
        }
        else {
            $blocks.Add((New-CustomerWordParagraphBlock -Text 'No roadmap action was placed in this bucket from the reviewed data.' -Style 'Normal')) | Out-Null
        }
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Executive Decision Required' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'These are the leadership approvals or owner decisions that would remove the biggest blockers to execution sequencing in the current roadmap.' -Style 'Normal')) | Out-Null
    if ($leadershipDecisionBlocks.Count -gt 0) {
        foreach ($block in $leadershipDecisionBlocks) {
            $blocks.Add($block) | Out-Null
        }
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Leadership decisions were not clearly separated in the reviewed data, so the grouped recommendations should be used to confirm execution order.' -Style 'Normal')) | Out-Null
    }

    return @($blocks.ToArray())
}

function Get-CustomerApplicationPrimaryRiskSignalText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$ApplicationRow
    )

    $signals = New-Object System.Collections.Generic.List[string]
    $highPrivilegeCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HighPrivilegePermissionCount'))
    $applicationPermissionCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('ApplicationPermissionCount'))
    $activitySignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'ActivitySignalState'
    $hasRecentActivity = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('HasRecentActivity'))
    $redirectSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'RedirectUriSignalState'
    $insecureRedirectUriCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('InsecureRedirectUriCount'))
    $ownerSignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $ApplicationRow -Name 'OwnerSignalState'
    $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('OwnerCount'))

    if (Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $ApplicationRow) {
        $signals.Add('Expired credentials') | Out-Null
    }
    elseif (Test-CustomerEnterpriseApplicationHasCredentialsExpiringSoon -ApplicationRow $ApplicationRow) {
        $signals.Add('Credentials expiring soon') | Out-Null
    }

    if ($null -ne $highPrivilegeCount -and $highPrivilegeCount -gt 0) {
        $signals.Add(('{0} high-privilege permission signal(s)' -f $highPrivilegeCount)) | Out-Null
    }
    elseif ($null -ne $applicationPermissionCount -and $applicationPermissionCount -gt 0) {
        $signals.Add(('{0} application permission(s)' -f $applicationPermissionCount)) | Out-Null
    }

    if ($activitySignalState -eq 'Collected' -and $hasRecentActivity -eq $false) {
        $signals.Add('No recent activity') | Out-Null
    }

    if (($redirectSignalState -eq 'Collected' -or $redirectSignalState -eq 'Partial') -and $null -ne $insecureRedirectUriCount -and $insecureRedirectUriCount -gt 0) {
        $signals.Add(('{0} redirect URI flag(s)' -f $insecureRedirectUriCount)) | Out-Null
    }

    if ($ownerSignalState -eq 'Collected' -and $null -ne $ownerCount -and $ownerCount -le 0) {
        $signals.Add('No owner coverage') | Out-Null
    }

    if (Test-CustomerEnterpriseApplicationSsoEnabled -ApplicationRow $ApplicationRow) {
        $signals.Add(('SSO: {0}' -f (Get-CustomerEnterpriseApplicationSsoModeLabel -ApplicationRow $ApplicationRow -Default 'Configured'))) | Out-Null
    }

    if ($signals.Count -eq 0) {
        return 'Review current owner, permission, and activity posture'
    }

    return Join-ArrayaReadableList -Items @($signals.ToArray())
}

function Get-CustomerEnterpriseApplicationBulletItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$ApplicationRows = @(),
        [Parameter(Mandatory = $false)][string]$Mode = 'General',
        [Parameter(Mandatory = $false)][int]$Max = 8
    )

    return @(
        foreach ($applicationRow in @($ApplicationRows | Select-Object -First $Max)) {
            $displayName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $applicationRow -Names @('DisplayName')) -Default 'Unnamed application'
            $activityText = Get-CustomerEnterpriseApplicationActivityText -ApplicationRow $applicationRow
            $observationText = Get-CustomerEnterpriseApplicationObservationText -ApplicationRow $applicationRow
            $line = switch ($Mode) {
                'Sso' {
                    '{0} - {1}; {2}; {3}' -f $displayName, (Get-CustomerEnterpriseApplicationSsoModeLabel -ApplicationRow $applicationRow -Default 'Configured'), (Get-CustomerApplicationPrimaryRiskSignalText -ApplicationRow $applicationRow), $observationText
                }
                'Cleanup' {
                    '{0} - {1}; {2}; {3}' -f $displayName, (Get-CustomerApplicationPrimaryRiskSignalText -ApplicationRow $applicationRow), $activityText, $observationText
                }
                'Credential' {
                    '{0} - {1}; {2}; {3}' -f $displayName, (Get-CustomerEnterpriseApplicationCredentialStateText -ApplicationRow $applicationRow), $activityText, $observationText
                }
                default {
                    '{0} - {1}; {2}' -f $displayName, (Get-CustomerApplicationPrimaryRiskSignalText -ApplicationRow $applicationRow), $observationText
                }
            }

            $line
        }
    )
}

function Get-CustomerReadableItemSampleText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string[]]$Items = @(),
        [Parameter(Mandatory = $false)][int]$Max = 4
    )

    $cleanItems = @(
        $Items |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            ForEach-Object { [string]$_ } |
            Select-Object -Unique
    )

    if ($cleanItems.Count -le 0) {
        return 'No named items surfaced'
    }

    if ($cleanItems.Count -le $Max) {
        return Join-ArrayaReadableList -Items $cleanItems
    }

    $sampleItems = @($cleanItems | Select-Object -First $Max)
    return '{0} (and {1} more)' -f (Join-ArrayaReadableList -Items $sampleItems), ($cleanItems.Count - $Max)
}

function Get-CustomerEnterpriseApplicationSpotlightListItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$ApplicationRows = @(),
        [Parameter(Mandatory = $false)][int]$Max = 8
    )

    $selectedApplicationRows = @($ApplicationRows | Select-Object -First $Max)
    if ($selectedApplicationRows.Count -le 0) {
        return @()
    }

    $groupedSpotlightRows = New-Object System.Collections.Generic.List[object]
    foreach ($applicationRow in $selectedApplicationRows) {
        $displayName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $applicationRow -Names @('DisplayName')) -Default 'Unnamed application'
        $ssoMode = Get-CustomerEnterpriseApplicationSsoModeLabel -ApplicationRow $applicationRow -Default 'Not detected'
        $riskSignal = Get-CustomerApplicationPrimaryRiskSignalText -ApplicationRow $applicationRow
        $activityText = Get-CustomerEnterpriseApplicationActivityText -ApplicationRow $applicationRow
        $groupKey = '{0}||{1}||{2}' -f $ssoMode, $riskSignal, $activityText

        $groupedSpotlightRows.Add([pscustomobject]@{
            DisplayName = $displayName
            SsoMode     = $ssoMode
            RiskSignal  = $riskSignal
            Activity    = $activityText
            GroupKey    = $groupKey
        }) | Out-Null
    }

    return @(
        foreach ($spotlightGroup in @($groupedSpotlightRows | Group-Object -Property GroupKey)) {
            if ($spotlightGroup.Count -le 0) {
                continue
            }

            $sampleRow = $spotlightGroup.Group | Select-Object -First 1
            $applicationNames = @($spotlightGroup.Group | ForEach-Object { [string]$_.DisplayName })
            $applicationNameText = Get-CustomerReadableItemSampleText -Items $applicationNames -Max 4
            'Applications: {0}; SSO: {1}; Primary risk signal: {2}; Latest activity: {3}' -f $applicationNameText, $sampleRow.SsoMode, $sampleRow.RiskSignal, $sampleRow.Activity
        }
    )
}

function Get-CustomerLargestSiteReviewListItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$SiteRows = @(),
        [Parameter(Mandatory = $false)][int]$Max = 5
    )

    return @(
        foreach ($siteRow in @($SiteRows | Select-Object -First $Max)) {
            $title = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $siteRow -Names @('Title')) -Default 'Unnamed site'
            $owner = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $siteRow -Names @('Owner')) -Default 'Owner not surfaced'
            $storage = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $siteRow -Names @('StorageUsedGB', 'StorageUsageCurrent')) -Default 'Not surfaced in current source'
            if ($storage -ne 'Not surfaced in current source') {
                $storage = ('{0} GB' -f $storage)
            }
            $lastModified = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $siteRow -Names @('LastContentModifiedDate')) -Default 'Not surfaced in current source'
            '{0} - Owner: {1}; Storage used: {2}; Last content modified: {3}' -f $title, $owner, $storage, $lastModified
        }
    )
}

function Convert-ToCustomerWordCellParagraphs {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$Content)

    $cellMetadata = Get-CustomerWordTableCellMetadata -Content $Content
    $Content = $cellMetadata.Content

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

function New-CustomerWordListXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string[]]$Items,
        [Parameter(Mandatory = $false)][string]$Style = 'ListBullet'
    )

    return @(
        foreach ($item in @($Items)) {
            $text = Convert-ToArrayaDisplayText -Value $item -Default ''
            if ([string]::IsNullOrWhiteSpace($text)) {
                continue
            }

            New-CustomerWordParagraphXml -Text $text -Style $Style
        }
    )
}

function New-CustomerWordTableCellXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$Content,
        [Parameter(Mandatory = $false)][switch]$Header
    )

    $cellMetadata = Get-CustomerWordTableCellMetadata -Content $Content
    $paragraphXml = @()
    foreach ($paragraph in (Convert-ToCustomerWordCellParagraphs -Content $cellMetadata.Content)) {
        $paragraphXml += New-CustomerWordParagraphXml -Text $paragraph -Style 'Normal' -Bold:$Header
    }

    $cellProperties = if ($Header) {
        '<w:tcPr><w:tcW w:w="0" w:type="auto"/><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr>'
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$cellMetadata.FillColor)) {
        '<w:tcPr><w:tcW w:w="0" w:type="auto"/><w:shd w:val="clear" w:color="auto" w:fill="' + ([string]$cellMetadata.FillColor).TrimStart('#') + '"/></w:tcPr>'
    }
    else {
        '<w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>'
    }

    return "<w:tc>$cellProperties$($paragraphXml -join '')</w:tc>"
}

function Get-CustomerImageContentTypeExtension {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$ContentType)

    switch ($ContentType.ToLowerInvariant()) {
        'image/png' { return 'png' }
        default { throw "Unsupported customer report image content type: $ContentType" }
    }
}

function Initialize-CustomerImageBlocks {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object[]]$Blocks)

    $imageIndex = 0
    foreach ($block in @($Blocks)) {
        if ($null -eq $block -or [string]$block.Type -ne 'Image') {
            continue
        }

        $extension = Get-CustomerImageContentTypeExtension -ContentType ([string]$block.ContentType)
        $mediaPath = 'word/media/customer-chart-{0:D3}.{1}' -f ($imageIndex + 1), $extension
        $relationshipId = 'rIdCustomerChart{0:D3}' -f ($imageIndex + 1)
        $docPrId = 5000 + $imageIndex
        $cx = [int]$block.WidthPx * 9525
        $cy = [int]$block.HeightPx * 9525

        $block | Add-Member -NotePropertyName MediaPath -NotePropertyValue $mediaPath -Force
        $block | Add-Member -NotePropertyName RelationshipId -NotePropertyValue $relationshipId -Force
        $block | Add-Member -NotePropertyName DocPrId -NotePropertyValue $docPrId -Force
        $block | Add-Member -NotePropertyName WidthEmu -NotePropertyValue $cx -Force
        $block | Add-Member -NotePropertyName HeightEmu -NotePropertyValue $cy -Force

        $imageIndex++
    }

    return @($Blocks)
}

function New-CustomerWordImageXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Block
    )

    $altText = Convert-ToCustomerWordXmlText -Text ([string]$Block.AltText)
    $rId = [string]$Block.RelationshipId
    $docPrId = [string]$Block.DocPrId
    $cx = [string]$Block.WidthEmu
    $cy = [string]$Block.HeightEmu

    return @"
<w:p>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="$cx" cy="$cy"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="$docPrId" name="Customer Chart $docPrId" descr="$altText"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="Customer Chart $docPrId" descr="$altText"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="$rId"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="$cx" cy="$cy"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
"@
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

function Set-CustomerZipBinaryEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Archive,
        [Parameter(Mandatory = $true)][string]$EntryName,
        [Parameter(Mandatory = $true)][byte[]]$Content
    )

    $existingEntry = $Archive.GetEntry($EntryName)
    if ($null -ne $existingEntry) {
        $existingEntry.Delete()
    }

    $entry = $Archive.CreateEntry($EntryName)
    $stream = $entry.Open()
    try {
        $stream.Write($Content, 0, $Content.Length)
    }
    finally {
        $stream.Dispose()
    }
}

function Set-CustomerAssessmentCoreProperties {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Archive,
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt,
        [Parameter(Mandatory = $false)][string]$DocumentTitle
    )

    $timestamp = $GeneratedAt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    if ([string]::IsNullOrWhiteSpace($DocumentTitle)) {
        $DocumentTitle = "$TenantName Microsoft 365 Tenant Best Practices Assessment"
    }

    $title = Convert-ToCustomerWordXmlText -Text $DocumentTitle
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

function Set-CustomerAssessmentContentTypes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Archive,
        [Parameter(Mandatory = $false)][object[]]$ImageBlocks = @()
    )

    if (@($ImageBlocks).Count -eq 0) {
        return
    }

    $entry = $Archive.GetEntry('[Content_Types].xml')
    if ($null -eq $entry) {
        return
    }

    $reader = New-Object System.IO.StreamReader($entry.Open())
    try {
        [xml]$contentTypesXml = $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }

    $typesNode = $contentTypesXml.Types
    foreach ($imageBlock in @($ImageBlocks)) {
        $extension = Get-CustomerImageContentTypeExtension -ContentType ([string]$imageBlock.ContentType)
        $hasDefault = @($typesNode.Default | Where-Object { [string]$_.Extension -eq $extension }).Count -gt 0
        if (-not $hasDefault) {
            $defaultNode = $contentTypesXml.CreateElement('Default', $typesNode.NamespaceURI)
            $null = $defaultNode.SetAttribute('Extension', $extension)
            $null = $defaultNode.SetAttribute('ContentType', [string]$imageBlock.ContentType)
            $null = $typesNode.AppendChild($defaultNode)
        }
    }

    Set-CustomerZipTextEntry -Archive $Archive -EntryName '[Content_Types].xml' -Content $contentTypesXml.OuterXml
}

function Set-CustomerAssessmentDocumentRelationships {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Archive,
        [Parameter(Mandatory = $false)][object[]]$ImageBlocks = @()
    )

    if (@($ImageBlocks).Count -eq 0) {
        return
    }

    $entryName = 'word/_rels/document.xml.rels'
    $entry = $Archive.GetEntry($entryName)
    if ($null -eq $entry) {
        return
    }

    $reader = New-Object System.IO.StreamReader($entry.Open())
    try {
        [xml]$relationshipsXml = $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }

    $relationshipsNode = $relationshipsXml.Relationships
    foreach ($imageBlock in @($ImageBlocks)) {
        $existing = @($relationshipsNode.Relationship | Where-Object { [string]$_.Id -eq [string]$imageBlock.RelationshipId })
        if ($existing.Count -gt 0) {
            continue
        }

        $relationshipNode = $relationshipsXml.CreateElement('Relationship', $relationshipsNode.NamespaceURI)
        $null = $relationshipNode.SetAttribute('Id', [string]$imageBlock.RelationshipId)
        $null = $relationshipNode.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        $null = $relationshipNode.SetAttribute('Target', ([string]$imageBlock.MediaPath -replace '^word/', ''))
        $null = $relationshipsNode.AppendChild($relationshipNode)
    }

    Set-CustomerZipTextEntry -Archive $Archive -EntryName $entryName -Content $relationshipsXml.OuterXml
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
            'List' {
                foreach ($paragraphXml in @(New-CustomerWordListXml -Items @($block.Items) -Style ([string]$block.Style))) {
                    $bodyElements.Add($paragraphXml) | Out-Null
                }
            }
            'Table' {
                $bodyElements.Add((New-CustomerWordTableXml -Headers @($block.Headers) -Rows @($block.Rows) -Style ([string]$block.Style))) | Out-Null
            }
            'Image' {
                $bodyElements.Add((New-CustomerWordImageXml -Block $block)) | Out-Null
            }
        }
    }

    $sectPrXml = Get-CustomerTemplateSectionPropertiesXml -TemplatePath $TemplatePath
    $bodyElements.Add($sectPrXml) | Out-Null

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
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
        [Parameter(Mandatory = $true)][object[]]$Blocks,
        [Parameter(Mandatory = $false)][string]$DocumentTitle
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath -Force

    $preparedBlocks = Initialize-CustomerImageBlocks -Blocks $Blocks
    $imageBlocks = @($preparedBlocks | Where-Object { $null -ne $_ -and [string]$_.Type -eq 'Image' })
    $documentXml = New-CustomerAssessmentDocumentXml -Blocks $preparedBlocks -TemplatePath $OutputPath
    $archive = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        Set-CustomerZipTextEntry -Archive $Archive -EntryName 'word/document.xml' -Content $documentXml
        Set-CustomerAssessmentCoreProperties -Archive $Archive -TenantName $TenantName -GeneratedAt $GeneratedAt -DocumentTitle $DocumentTitle
        Set-CustomerAssessmentContentTypes -Archive $Archive -ImageBlocks $imageBlocks
        Set-CustomerAssessmentDocumentRelationships -Archive $Archive -ImageBlocks $imageBlocks
        foreach ($imageBlock in @($imageBlocks)) {
            Set-CustomerZipBinaryEntry -Archive $Archive -EntryName ([string]$imageBlock.MediaPath) -Content ([byte[]]$imageBlock.ImageBytes)
        }
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
        'ListBullet' { return @("- $text") }
        'ListBullet2' { return @("- $text") }
        'ListNumber' { return @("- $text") }
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
    $documentRevision = Get-CustomerDocumentRevisionLabel
    $findings = @($SourceModel.Findings)
    $roadmapActions = @(Get-CustomerDistinctRoadmapActions -RoadmapActions $SourceModel.RoadmapActions)
    $workstreamSummaries = @($SourceModel.WorkstreamSummaries)
    $executiveThemes = @($SourceModel.ExecutiveThemes)
    $sourceSummaryRows = @($SourceModel.SummaryRows)
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
    $rawEnterpriseApplications = Convert-ArrayaObjectToArray $Signals.EnterpriseApplications
    $authenticationSsoApplications = Convert-ArrayaObjectToArray $Signals.AuthenticationSSOApplications
    $enterpriseApplications = Merge-CustomerEnterpriseApplicationRows -EnterpriseApplicationRows $rawEnterpriseApplications -AuthenticationSsoApplicationRows $authenticationSsoApplications
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
    $emailActivityTopReceivers = Convert-ArrayaObjectToArray $Signals.EmailActivityTopReceivers
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
    $appsWithExpiredCredentials = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('ApplicationsWithExpiredCredentials'))
    $appsWithCredentialsExpiringSoon = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('ApplicationsWithCredentialsExpiringSoon'))
    if ($null -eq $appsWithExpiredCredentials) {
        $appsWithExpiredCredentials = @($enterpriseApplications | Where-Object { Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $_ }).Count
    }
    if ($null -eq $appsWithCredentialsExpiringSoon) {
        $appsWithCredentialsExpiringSoon = @($enterpriseApplications | Where-Object { Test-CustomerEnterpriseApplicationHasCredentialsExpiringSoon -ApplicationRow $_ }).Count
    }

    $ssoApplicationRows = @(
        $enterpriseApplications |
            Where-Object { Test-CustomerEnterpriseApplicationSsoEnabled -ApplicationRow $_ } |
            Sort-Object `
                @{ Expression = { $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissionCount')); if ($null -eq $value) { 0 } else { $value } }; Descending = $true }, `
                @{ Expression = { $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount')); if ($null -eq $value) { 0 } else { $value } }; Descending = $true }, `
                @{ Expression = {
                    $latestActivity = Get-CustomerEnterpriseApplicationLatestActivityDate -ApplicationRow $_
                    if ($null -eq $latestActivity) { [datetime]::MinValue } else { $latestActivity }
                }; Descending = $true }, `
                @{ Expression = { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed application' } }
    )
    $mostActiveEnterpriseApplications = @(
        $enterpriseApplications |
            Where-Object { $null -ne (Get-CustomerEnterpriseApplicationLatestActivityDate -ApplicationRow $_) } |
            Sort-Object `
                @{ Expression = {
                    $latestActivity = Get-CustomerEnterpriseApplicationLatestActivityDate -ApplicationRow $_
                    if ($null -eq $latestActivity) { [datetime]::MinValue } else { $latestActivity }
                }; Descending = $true }, `
                @{ Expression = { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed application' } }
    )
    $cleanupCandidateApplications = @(
        $enterpriseApplications |
            Where-Object {
                ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissionCount'))) -gt 0) -or
                ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount'))) -gt 0) -or
                (((Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Collected') -and ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasRecentActivity'))) -eq $false)) -or
                (Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $_)
            } |
            Sort-Object `
                @{ Expression = {
                    $activityState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState'
                    $hasRecentActivity = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasRecentActivity'))
                    if (($activityState -eq 'Collected') -and ($hasRecentActivity -eq $false)) { 1 } else { 0 }
                }; Descending = $true }, `
                @{ Expression = { if (Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $_) { 1 } else { 0 } }; Descending = $true }, `
                @{ Expression = { $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissionCount')); if ($null -eq $value) { 0 } else { $value } }; Descending = $true }, `
                @{ Expression = { $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount')); if ($null -eq $value) { 0 } else { $value } }; Descending = $true }, `
                @{ Expression = {
                    $latestActivity = Get-CustomerEnterpriseApplicationLatestActivityDate -ApplicationRow $_
                    if ($null -eq $latestActivity) { [datetime]::MinValue } else { $latestActivity }
                } }, `
                @{ Expression = { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed application' } }
    )
    $credentialCleanupApplications = @(
        $enterpriseApplications |
            Where-Object {
                (Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $_) -or
                (Test-CustomerEnterpriseApplicationHasCredentialsExpiringSoon -ApplicationRow $_) -or
                (
                    -not [string]::IsNullOrWhiteSpace((Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('AppCredentials')) -Default '')) -and
                    ((Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Collected') -and
                    ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasRecentActivity'))) -eq $false)
                )
            } |
            Sort-Object `
                @{ Expression = { if (Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $_) { 1 } else { 0 } }; Descending = $true }, `
                @{ Expression = { if (Test-CustomerEnterpriseApplicationHasCredentialsExpiringSoon -ApplicationRow $_) { 1 } else { 0 } }; Descending = $true }, `
                @{ Expression = {
                    $activityState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState'
                    $hasRecentActivity = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasRecentActivity'))
                    if (($activityState -eq 'Collected') -and ($hasRecentActivity -eq $false)) { 1 } else { 0 }
                }; Descending = $true }, `
                @{ Expression = { $value = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissionCount')); if ($null -eq $value) { 0 } else { $value } }; Descending = $true }, `
                @{ Expression = { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed application' } }
    )
    $ssoExampleNames = @(
        $ssoApplicationRows |
            Select-Object -First 3 |
            ForEach-Object { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default '' } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
    $activeExampleNames = @(
        $mostActiveEnterpriseApplications |
            Select-Object -First 3 |
            ForEach-Object { Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default '' } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

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

    $indexedRoadmapActions = @(
        for ($index = 0; $index -lt @($roadmapActions).Count; $index++) {
            [pscustomobject]@{
                OriginalIndex = $index
                Action        = $roadmapActions[$index]
            }
        }
    )
    $sortedRoadmapActions = @(
        $indexedRoadmapActions |
            Sort-Object -Property `
                @{ Expression = { Get-CustomerSeverityRank -Severity ([string]$_.Action.HighestSeverity) }; Descending = $true }, `
                @{ Expression = 'OriginalIndex'; Descending = $false } |
            ForEach-Object {
                $_.Action
            }
    )

    $recommendationRows = @(
        foreach ($action in @($sortedRoadmapActions)) {
            $criticality = Get-CustomerActionCriticalityLabel -Severity ([string]$action.HighestSeverity)
            $effort = Get-CustomerActionEffortLabel -Action $action
            $roughPsHours = Get-CustomerActionEstimatedPsHoursLabel -Action $action
            New-CustomerWordTableRow -Cells @(
                [string]$action.ActionTitle,
                (New-CustomerWordTableCellContent -Content $criticality -FillColor (Get-CustomerCriticalityFillColor -Criticality $criticality)),
                $effort,
                $roughPsHours
            )
        }
    )
    if ($recommendationRows.Count -eq 0) {
        $recommendationRows = @(
            New-CustomerWordTableRow -Cells @(
                'Priority work item not clearly surfaced',
                (New-CustomerWordTableCellContent -Content 'Low' -FillColor (Get-CustomerCriticalityFillColor -Criticality 'Low')),
                'Standard',
                '14-24 hours'
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
    $devicePlatformChartRows = @(
        $platformGroups |
            Sort-Object Count -Descending |
            Select-Object -First 5 |
            ForEach-Object {
                [pscustomobject]@{
                    Label = $_.Name
                    Value = $_.Count
                }
            }
    )
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
                $ssoEnabled = Test-CustomerEnterpriseApplicationSsoEnabled -ApplicationRow $enterpriseApplication
                $ssoEnabledText = Convert-ToCustomerAssessmentBooleanLabel -Value $ssoEnabled -TrueText 'Enabled' -FalseText 'Not detected' -Default 'Not validated from the reviewed data'
                $effectiveSsoMode = if ($ssoEnabled -eq $true) { Get-CustomerEnterpriseApplicationSsoModeLabel -ApplicationRow $enterpriseApplication -Default 'Configured' } else { 'Not detected' }
                New-CustomerWordTableRow -Cells @(
                    $displayName,
                    $ssoEnabledText,
                    $effectiveSsoMode,
                    (Get-CustomerEnterpriseApplicationObservationText -ApplicationRow $enterpriseApplication)
                )
            }
        )
    }
    elseif ($applicationInventoryClearlyEmpty) {
        @((New-CustomerWordTableRow -Cells @('No enterprise applications surfaced in the reviewed data', 'Not detected', 'Not detected', 'Treat this as a current-state result and validate it if the tenant expects third-party or line-of-business applications.')))
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Enterprise application inventory not validated', 'Validation note', 'Validation note', 'The current review did not surface a usable enterprise application inventory, so this section should not be read as proof that no enterprise applications exist.')))
    }
    $ssoFocusedApplicationRows = if ($ssoApplicationRows.Count -gt 0) {
        @(
            foreach ($enterpriseApplication in @($ssoApplicationRows | Select-Object -First 8)) {
                $permissionSummary = Get-CustomerEnterpriseApplicationPermissionSummaryText -ApplicationRow $enterpriseApplication
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('DisplayName')) -Default 'Unnamed application'),
                    (Get-CustomerEnterpriseApplicationSsoModeLabel -ApplicationRow $enterpriseApplication -Default 'Configured'),
                    ('{0}; {1}' -f $permissionSummary, (Get-CustomerEnterpriseApplicationActivityText -ApplicationRow $enterpriseApplication)),
                    (Get-CustomerEnterpriseApplicationObservationText -ApplicationRow $enterpriseApplication)
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('No SSO-enabled enterprise applications were surfaced in the reviewed data', 'Not detected', 'Not validated from the reviewed data', 'If SSO-enabled SaaS or line-of-business apps are expected in this tenant, validate that the enterprise application inventory is complete.')))
    }
    $cleanupOpportunityRows = if ($cleanupCandidateApplications.Count -gt 0) {
        @(
            foreach ($enterpriseApplication in @($cleanupCandidateApplications | Select-Object -First 8)) {
                $highPrivilegeCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('HighPrivilegePermissionCount'))
                $applicationPermissionCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('ApplicationPermissionCount'))
                $activitySignalState = Get-CustomerEnterpriseApplicationSignalState -ApplicationRow $enterpriseApplication -Name 'ActivitySignalState'
                $hasRecentActivity = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('HasRecentActivity'))
                $cleanupSignals = New-Object System.Collections.Generic.List[string]
                if ($activitySignalState -eq 'Collected' -and $hasRecentActivity -eq $false) {
                    $cleanupSignals.Add('No recent activity') | Out-Null
                }
                if ($null -ne $highPrivilegeCount -and $highPrivilegeCount -gt 0) {
                    $cleanupSignals.Add(('{0} high-privilege permission signal(s)' -f $highPrivilegeCount)) | Out-Null
                }
                if ($null -ne $applicationPermissionCount -and $applicationPermissionCount -gt 0) {
                    $cleanupSignals.Add(('{0} application permission(s)' -f $applicationPermissionCount)) | Out-Null
                }
                if (Test-CustomerEnterpriseApplicationHasExpiredCredentials -ApplicationRow $enterpriseApplication) {
                    $cleanupSignals.Add('Expired credentials') | Out-Null
                }
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('DisplayName')) -Default 'Unnamed application'),
                    (Join-ArrayaReadableList -Items @($cleanupSignals.ToArray())),
                    (Get-CustomerEnterpriseApplicationActivityText -ApplicationRow $enterpriseApplication),
                    (Get-CustomerEnterpriseApplicationObservationText -ApplicationRow $enterpriseApplication)
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('No inactive or high-privilege cleanup candidates were surfaced in the reviewed data', 'No immediate cleanup signal', 'Not validated from the reviewed data', 'Continue to validate that enterprise applications with broad access remain governed through owner review and sign-in monitoring.')))
    }
    $credentialLifecycleRows = if ($credentialCleanupApplications.Count -gt 0) {
        @(
            foreach ($enterpriseApplication in @($credentialCleanupApplications | Select-Object -First 8)) {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('DisplayName')) -Default 'Unnamed application'),
                    (Get-CustomerEnterpriseApplicationCredentialStateText -ApplicationRow $enterpriseApplication),
                    (Get-CustomerEnterpriseApplicationActivityText -ApplicationRow $enterpriseApplication),
                    (Get-CustomerEnterpriseApplicationObservationText -ApplicationRow $enterpriseApplication)
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('No credential lifecycle cleanup candidates were surfaced in the reviewed data', 'Credential lifecycle did not show an immediate issue in the reviewed app set', 'Not validated from the reviewed data', 'If app credentials are expected in this tenant, continue to review expiry posture alongside enterprise-app activity and least-privilege cleanup.')))
    }
    $applicationSpotlightListItems = @(Get-CustomerEnterpriseApplicationSpotlightListItems -ApplicationRows $noteworthyEnterpriseApplications -Max 8)
    $ssoApplicationListItems = @(Get-CustomerEnterpriseApplicationBulletItems -ApplicationRows $ssoApplicationRows -Mode 'Sso' -Max 8)
    $cleanupCandidateListItems = @(Get-CustomerEnterpriseApplicationBulletItems -ApplicationRows $cleanupCandidateApplications -Mode 'Cleanup' -Max 8)
    $credentialCleanupListItems = @(Get-CustomerEnterpriseApplicationBulletItems -ApplicationRows $credentialCleanupApplications -Mode 'Credential' -Max 8)
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
    $globalAdminRecommendationItems = @(
        'Reduce the number of active Global Administrator assignments and remove the role entirely from inactive or stale privileged accounts.',
        'Move administrators who do not routinely need full tenant-wide change authority to lower-privilege roles such as Global Reader.',
        'Retain Global Administrator as an eligible role that requires approval before activation when elevated access is only occasionally needed.'
    )
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
    $inactiveMailboxExampleItems = @(
        foreach ($inactiveMailbox in @($inactiveMailboxRows | Select-Object -First 5)) {
            $displayName = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $inactiveMailbox -Names @('DisplayName')) -Default 'Unnamed inactive mailbox'
            $address = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $inactiveMailbox -Names @('PrimarySmtpAddress', 'UserPrincipalName')) -Default 'address not surfaced'
            $recipientType = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $inactiveMailbox -Names @('RecipientTypeDetails')) -Default 'type not surfaced'
            '{0} - {1}; {2}' -f $displayName, $address, $recipientType
        }
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
    $recipientMixChartRows = @(Get-CustomerRecipientTypeChartRows -RecipientRows $recipientRows)
    $recipientDomainChartRows = @(Get-CustomerRecipientDomainChartRows -RecipientRows $recipientRows -Top 6)
    $topRecipientDomainSummaryRows = @(
        Get-CustomerRecipientDomainSummaryRows -RecipientRows $recipientRows -Top 6 |
            ForEach-Object {
                New-CustomerWordTableRow -Cells @($_.Domain, $_.Total, $_.Primary, $_.AliasOnly)
            }
    )
    $externalExposureCategoryChartRows = @(Get-CustomerExternalExposureCategoryChartRows -ExternalExposureFindings $externalExposureFindings -Top 6)
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
    $largestCollaborationSiteListItems = @(Get-CustomerLargestSiteReviewListItems -SiteRows $largestSharePointSites -Max 5)
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
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Document revision: {0}" -f $documentRevision) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Version History' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This section tracks the issued version of the assessment report so the customer-facing document has a clear revision record.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Date', 'Revision', 'Author', 'Description', 'Reviewers') -Rows @(
        (New-CustomerWordTableRow -Cells @($GeneratedAt.ToString('MM/dd/yyyy'), $documentRevision, 'Arraya Solutions', 'Current assessment report release', 'Not surfaced in current source'))
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
    $blocks.Add((New-CustomerWordListBlock -Items (Get-CustomerExecutiveRiskBulletItems -Rows $executiveRiskRows))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Leadership Decision Brief' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items (Get-CustomerLeadershipDecisionBulletItems -Rows $leadershipDecisionRows))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This brief highlights the short list of leadership approvals that would remove the biggest blockers to a cleaner operating baseline.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Overall Findings Summary' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Workstream', 'Severity / Impact', 'Open Findings', 'What Stands Out') -Rows $overallFindingsSummaryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This table shows where findings are clustering before the report moves into the detailed sections.' -Style 'Normal')) | Out-Null
    $overallFindingsChartRows = Get-CustomerOverallFindingsChartRows -WorkstreamSummaries $workstreamSummaries -Findings $findings
    $overallFindingsChartHeight = [Math]::Max(220, (55 + (28 * @($overallFindingsChartRows).Count)))
    $overallFindingsChartBlock = New-CustomerChartImageBlock -ChartType 'HorizontalBar' -Rows $overallFindingsChartRows -AltText 'Overall Findings by Workstream' -WidthPx 720 -HeightPx $overallFindingsChartHeight
    if ($null -ne $overallFindingsChartBlock) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Overall Findings by Workstream' -Style 'Heading3')) | Out-Null
        $blocks.Add($overallFindingsChartBlock) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Use 4.0 Modern Workplace Recommendations for the prioritized execution view, the companion remediation roadmap for leadership sequencing, and the Engineer Pack for detailed evidence behind each grouped issue.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '4.0 Modern Workplace Recommendations' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This section is the streamlined execution view for the report. Use it to prioritize the work, then use the detailed sections, companion roadmap, and Engineer Pack to validate the evidence behind each recommendation.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Recommendation', 'Criticality', 'Level of Effort', 'Rough PS Hours') -Rows $recommendationRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Level of Effort is an initial delivery-planning estimate intended to help sequence work at a glance. Rough PS Hours is a combined engineering and project-management estimate covering prep, review, presentation, implementation, QA, and finalization. It does not include customer-side wait states or third-party execution.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '5.0 Entra ID Review: User and Device Inventory' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Review of the tenant configuration indicated that identity hygiene and device governance need to be read together in this environment. The same parts of the tenant that are carrying stale privileged access are also the parts of the environment where compliance-driven access control is least mature.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '5.1 Entra User' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'User Account Summary' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $userSummaryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'User Account Status Comparison' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Account Type', 'Total', 'Enabled', 'Inactive') -Rows $userComparisonRows)) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(
        ("The user inventory shows {0} internal member account(s) and {1} guest account(s) in the reviewed data." -f $(if ($memberUsers.Count -gt 0 -or $userRows.Count -gt 0) { $memberUsers.Count } else { 'an unconfirmed number of' }), $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { $guestUsers.Count } else { 'an unconfirmed number of' })),
        'Enabled member accounts with no recent sign-in activity usually point to lifecycle drift rather than a single isolated exception.',
        'When viewed collectively, the identity inventory suggests that access cleanup is not limited to one user population.',
        'Internal members and guest identities both show evidence of stale access, which increases the chance that a dormant identity still retains a path into collaboration, messaging, or privileged workflows.'
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '5.2 Entra Device' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Registration Summary' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $deviceSummaryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device Platform Distribution' -Style 'Heading3')) | Out-Null
    $devicePlatformChartBlock = New-CustomerChartImageBlock -ChartType 'Donut' -Rows $devicePlatformChartRows -AltText 'Device Platform Distribution' -WidthPx 540 -HeightPx 280
    if ($null -ne $devicePlatformChartBlock) {
        $blocks.Add($devicePlatformChartBlock) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device platform distribution was not surfaced clearly enough to render a chart from the reviewed data.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordListBlock -Items @(
        @(
            Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.ObservedNarrative
            ('Dominant platform: {0}' -f $dominantPlatform)
            (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.WhyItMatters)
            (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.PositiveNarrative)
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device Registration and Compliance Gaps' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(
        'The reviewed device inventory makes it possible to see where endpoint registration exists without the same level of follow-through in management and compliance.',
        'In this tenant, that gap is most visible where stale, unmanaged, or non-compliant devices remain part of the active footprint.',
        (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.WhyItMatters),
        (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.PositiveNarrative)
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '5.3 Entra Guest Access Configuration' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'External Access Snapshot' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $externalAccessSnapshotTableRows)) | Out-Null
    $guestObservedSegments = @(
        ('Guest invitation is {0}.' -f $guestInvitationControlText),
        ('The tenant currently has {0} configured cross-tenant partner relationship(s).' -f $crossTenantPartnerCountText),
        ('Default inbound MFA trust is shown as {0}.' -f $defaultInboundMfaTrustText),
        $(if (-not [string]::IsNullOrWhiteSpace($guestMfaCurrentStateSummaryText)) { $guestMfaCurrentStateSummaryText } else { $null })
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $guestObservedText = (@($guestObservedSegments) -join ' ')
    $guestWhyItMattersText = if (-not [string]::IsNullOrWhiteSpace($guestMfaWhyThisMattersText)) {
        $guestMfaWhyThisMattersText
    }
    else {
        'Guest access combines identity risk with collaboration exposure, so invitation scope and inbound trust should be reviewed as one control story rather than as separate settings.'
    }
    $guestRecommendedNextStepParts = @(
        $(if (-not [string]::IsNullOrWhiteSpace($guestMfaDesiredStateText)) { $guestMfaDesiredStateText } else { 'Require strong guest authentication through a guest-specific Conditional Access policy and review cross-tenant trust deliberately rather than broadly trusting every external tenant by default.' }),
        $(if (-not [string]::IsNullOrWhiteSpace($guestMfaRecommendationStrategyText)) { $guestMfaRecommendationStrategyText } elseif (-not [string]::IsNullOrWhiteSpace($guestMfaUserExperienceText)) { $guestMfaUserExperienceText } else { $null })
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Observed Guest and B2B Posture' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text $guestObservedText -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Why It Matters' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text $guestWhyItMattersText -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recommended Next Step' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ((@($guestRecommendedNextStepParts) -join ' ')) -Style 'Normal')) | Out-Null

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
        if ($null -ne $appsWithExpiredCredentials -and $appsWithExpiredCredentials -gt 0) {
            $validatedApplicationNarrativeSegments.Add("$appsWithExpiredCredentials with expired app credentials") | Out-Null
        }
        if ($null -ne $appsWithCredentialsExpiringSoon -and $appsWithCredentialsExpiringSoon -gt 0) {
            $validatedApplicationNarrativeSegments.Add("$appsWithCredentialsExpiringSoon with credentials expiring soon") | Out-Null
        }

        $validatedApplicationNarrativeText = Join-ArrayaReadableList -Items @($validatedApplicationNarrativeSegments.ToArray())
        $applicationNarrative = "The application review looked at enterprise application inventory, sign-in posture, and the degree to which privileged permissions are visible in the current tenant data. The reviewed source surfaced {0} enterprise application(s)" -f $(if ($null -ne $totalEnterpriseApps) { $totalEnterpriseApps } else { 'an unconfirmed number of' })
        if (-not [string]::IsNullOrWhiteSpace($validatedApplicationNarrativeText)) {
            $applicationNarrative += ", including $validatedApplicationNarrativeText"
        }
        $applicationNarrative += '.'
        if ($ssoExampleNames.Count -gt 0) {
            $applicationNarrative += " The stronger SSO examples in the reviewed data were {0}." -f (Join-ArrayaReadableList -Items $ssoExampleNames)
        }
        if ($activeExampleNames.Count -gt 0) {
            $applicationNarrative += " The most recently active applications surfaced in the reviewed sign-in history were {0}." -f (Join-ArrayaReadableList -Items $activeExampleNames)
        }
        if (-not [string]::IsNullOrWhiteSpace($applicationValidationNoteText)) {
            $applicationNarrative += " $applicationValidationNoteText"
        }
        $blocks.Add((New-CustomerWordParagraphBlock -Text $applicationNarrative -Style 'Normal')) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current review did not surface a usable enterprise application inventory. Validation note: this section should be treated as incomplete rather than as proof that no enterprise applications exist in the tenant.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application Spotlight' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This spotlight list highlights the enterprise applications that surfaced the strongest SSO, privilege, credential, and activity cleanup signals from the current tenant data.' -Style 'Normal')) | Out-Null
    if ($applicationSpotlightListItems.Count -gt 0) {
        $blocks.Add((New-CustomerWordListBlock -Items $applicationSpotlightListItems)) | Out-Null
    }
    elseif ($applicationInventoryClearlyEmpty) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'No enterprise applications were surfaced in the reviewed data. Treat that as a current-state result and validate it if the tenant expects third-party or line-of-business applications.' -Style 'Normal')) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current review did not surface enough enterprise application detail to build a confident spotlight list, so this section should be treated as an inventory validation note rather than a clean bill of health.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'SSO-Enabled Applications' -Style 'Heading3')) | Out-Null
    if ($ssoApplicationListItems.Count -gt 0) {
        $blocks.Add((New-CustomerWordListBlock -Items $ssoApplicationListItems)) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'No SSO-enabled applications surfaced an actionable SSO review signal from the reviewed data.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Inactive or High-Privilege Applications' -Style 'Heading3')) | Out-Null
    if ($cleanupCandidateListItems.Count -gt 0) {
        $blocks.Add((New-CustomerWordListBlock -Items $cleanupCandidateListItems)) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'No inactive or high-privilege applications surfaced as immediate cleanup candidates from the reviewed data.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Credential Cleanup Opportunities' -Style 'Heading3')) | Out-Null
    if ($credentialCleanupListItems.Count -gt 0) {
        $blocks.Add((New-CustomerWordListBlock -Items $credentialCleanupListItems)) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'No expired or near-expiry credential cleanup candidates surfaced from the reviewed application set.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application Sign-In Activity and Security Analysis' -Style 'Heading3')) | Out-Null
    if ($applicationInventoryClearlyEmpty) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Because the current review did not surface enterprise application objects, this section does not point to a present app-governance concentration on its own. The next decision here is simply whether the observed zero-app result matches the tenant operating model or whether the inventory should be validated again.' -Style 'Normal')) | Out-Null
    }
    elseif (-not $applicationInventoryValidated) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Validation note: the current review did not surface enough enterprise application detail to make a confident governance call. Before using this section for a remediation decision, confirm whether the app inventory was intentionally out of scope or simply not visible in the reviewed data.' -Style 'Normal')) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The most useful cleanup signals in this section are SSO configuration type, elevated permission posture, recent or absent app activity, and whether credential lifecycle hygiene is still being maintained. Applications that combine high privilege with stale activity or expired secrets/certificates are usually the fastest rationalization opportunities because they can represent both standing access and operational drift.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recommendations for Application Governance' -Style 'Heading3')) | Out-Null
    if ($applicationInventoryValidated -or $applicationInventoryClearlyEmpty) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Use the prioritized application sample together with the SSO, inactivity, and credential lifecycle tables to identify cleanup candidates first. The most important follow-up actions are to confirm owners, retire apps that no longer show recent use, reduce broad permissions where business need is unclear, and rotate or remove expired client secrets and certificates before treating the current app estate as stable.' -Style 'Normal')) | Out-Null
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
                    $gapCategoryText = (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory')) -Default '').ToLowerInvariant()
                    switch -Regex ($gapCategoryText) {
                        'excluded' { 0 }
                        'outside.*include scope' { 1 }
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
        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory')) -Default '').ToLowerInvariant() -match 'excluded'
    }).Count
    $mfaOutsideIncludeUserCount = @($mfaGapRowsSorted | Where-Object {
        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory')) -Default '').ToLowerInvariant() -match 'outside.*include scope'
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
                    $coverageDriverText = Get-CustomerCompactMfaGapDriverText `
                        -GapCategory (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory')) `
                        -GapReason (Get-ArrayaObjectValue -Object $_ -Names @('GapReason')) `
                        -RelatedPolicies (Get-ArrayaObjectValue -Object $_ -Names @('RelatedPolicies'))
                    $userPrincipalName = Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default 'Not validated from the reviewed data'

                    New-CustomerWordTableRow -Cells @(
                        (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed user'),
                        $userPrincipalName,
                        (Get-CustomerMfaGapCategoryLabel -Value (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory'))),
                        $coverageDriverText
                    )
                }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('No uncovered internal member users surfaced in the reviewed data', 'N/A', 'N/A', 'Guest and external-user gaps remain summarized separately in this report and in the workbook tabs.')))
    }
    $mfaGapDriverSummaryItems = if ($mfaGapInternalMemberRowsSorted.Count -gt 0) {
        @(
            $mfaGapInternalMemberRowsSorted |
                ForEach-Object {
                    Get-CustomerCompactMfaGapDriverText `
                        -GapCategory (Get-ArrayaObjectValue -Object $_ -Names @('GapCategory')) `
                        -GapReason (Get-ArrayaObjectValue -Object $_ -Names @('GapReason')) `
                        -RelatedPolicies (Get-ArrayaObjectValue -Object $_ -Names @('RelatedPolicies'))
                } |
                Group-Object |
                Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Descending = $false } |
                Select-Object -First 4 |
                ForEach-Object {
                    "{0}: {1} uncovered internal member user(s)." -f $_.Name, $_.Count
                }
        )
    }
    else {
        @('No repeated internal-member MFA coverage driver surfaced in the reviewed data.')
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

    $mfaOutsideIncludeInternalMemberRows = @(
        $mfaGapInternalMemberRowsSorted | Where-Object { -not (Test-CustomerMfaGapHasExclusionHit -GapRow $_) }
    )
    $mfaExcludedInternalMemberRows = @(
        $mfaGapInternalMemberRowsSorted | Where-Object { Test-CustomerMfaGapHasExclusionHit -GapRow $_ }
    )
    $mfaOutsideIncludeDetailRows = @(
        $mfaOutsideIncludeInternalMemberRows |
            Select-Object -First 15 |
            ForEach-Object {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed user'),
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default 'Not validated from the reviewed data')
                )
            }
    )
    $mfaExcludedDetailRows = @(
        $mfaExcludedInternalMemberRows |
            Select-Object -First 15 |
            ForEach-Object {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default 'Unnamed user'),
                    (Convert-ToCustomerAssessmentDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('UserPrincipalName')) -Default 'Not validated from the reviewed data')
                )
            }
    )

    $mfaEnrollmentStatusChartRows = @(
        [pscustomobject]@{ Label = 'Registered'; Value = $(if ($null -ne $mfaRegisteredUsers) { [double]$mfaRegisteredUsers } else { 0.0 }) }
        [pscustomobject]@{ Label = 'Not registered'; Value = $(if ($null -ne $mfaNotRegisteredUsers) { [double]$mfaNotRegisteredUsers } else { 0.0 }) }
    )
    $mfaEnrollmentMethodChartRows = Get-CustomerMfaMethodChartRows -RegistrationSummaryRecord $mfaRegistrationSummaryRecord -EnrollmentSummaryRecord $mfaEnrollmentSummaryRecord -Top 6
    $mfaEnforcementCoverageChartRows = @(
        [pscustomobject]@{ Label = 'Covered by active MFA enforcement'; Value = $(if ($null -ne $mfaUsersCoveredByEnabledPolicies) { [double]$mfaUsersCoveredByEnabledPolicies } else { 0.0 }) }
        [pscustomobject]@{ Label = 'Not covered by active MFA enforcement'; Value = $(if ($null -ne $mfaUsersNotCoveredByEnabledPolicies) { [double]$mfaUsersNotCoveredByEnabledPolicies } else { 0.0 }) }
    )
    $mfaEnforcementDriverChartRows = @(
        [pscustomobject]@{ Label = 'Outside active MFA include scope'; Value = [double]$mfaOutsideIncludeUserCount }
        [pscustomobject]@{ Label = 'Excluded from active MFA policies'; Value = [double]$mfaExcludedUserCount }
    )
    $adminActivityChartRows = Get-CustomerAdminActivityChartRows -EnabledAdminRows $enabledAdminRows -StaleThresholdDays 180
    $adminMfaCoverageChartRows = @(
        [pscustomobject]@{ Label = 'Covered by active MFA enforcement'; Value = $(if ($null -ne $adminUsersCoveredByMfaEnforcement) { [double]$adminUsersCoveredByMfaEnforcement } else { 0.0 }) }
        [pscustomobject]@{ Label = 'Not covered by active MFA enforcement'; Value = $(if ($null -ne $adminUsersNotCoveredByMfaEnforcement) { [double]$adminUsersNotCoveredByMfaEnforcement } else { 0.0 }) }
    )
    $dataFootprintChartRows = Get-CustomerDataFootprintChartRows -PrimaryMailboxStatsRows $primaryMailboxStatsRows -ArchiveMailboxStatsRows $archiveMailboxStatsRows -SharePointRows $sharePointRows -OneDriveRows $oneDriveRows
    $topSenderTableRows = Get-CustomerMessageActivityTableRows -Rows $emailActivityTopSenders -CountPropertyName 'SendCount' -Top 5
    $topReceiverTableRows = Get-CustomerMessageActivityTableRows -Rows $emailActivityTopReceivers -CountPropertyName 'ReceiveCount' -Top 5

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
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'MFA Enrollment Status' -Style 'Heading3')) | Out-Null
    $mfaEnrollmentStatusChartBlock = New-CustomerChartImageBlock -ChartType 'Donut' -Rows $mfaEnrollmentStatusChartRows -AltText 'MFA Enrollment Status' -WidthPx 540 -HeightPx 280
    if ($null -ne $mfaEnrollmentStatusChartBlock) {
        $blocks.Add($mfaEnrollmentStatusChartBlock) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'MFA enrollment status was not surfaced clearly enough to render a chart from the reviewed data.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Registered MFA Method Mix' -Style 'Heading3')) | Out-Null
    $mfaEnrollmentMethodChartBlock = New-CustomerChartImageBlock -ChartType 'HorizontalBar' -Rows $mfaEnrollmentMethodChartRows -AltText 'Registered MFA Method Mix' -WidthPx 720 -HeightPx 240
    if ($null -ne $mfaEnrollmentMethodChartBlock) {
        $blocks.Add($mfaEnrollmentMethodChartBlock) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Registered MFA method mix was not surfaced clearly enough to render a chart from the reviewed data.' -Style 'Normal')) | Out-Null
    }
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
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Enrollment and enforcement are intentionally reported as separate views in this report. Users who are not enrolled in MFA remain an important readiness gap, but they are not added into the Conditional Access enforcement counts below.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Enforcement shows whether users are actually being required to perform MFA, not just whether they have registered methods. In this review, the policy baseline shows {0} enabled Conditional Access policy/policies that require MFA and {1} still in report-only mode. {2} Report-only policies do not count as enforced coverage. Detailed uncovered-user and policy-scope review rows are available in the workbook tabs MfaEnforcementGapUsers and MfaEnforcementScopeReview." -f $(if ($null -eq $enabledMfaEnforcementPolicies) { 'an unconfirmed number of' } else { $enabledMfaEnforcementPolicies }), $(if ($null -eq $reportOnlyMfaEnforcementPolicies) { 'an unconfirmed number of policies' } else { $reportOnlyMfaEnforcementPolicies }), $mfaCoverageNarrativeText) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Covered vs Not Covered by Active MFA Enforcement' -Style 'Heading3')) | Out-Null
    $mfaEnforcementCoverageChartBlock = New-CustomerChartImageBlock -ChartType 'Donut' -Rows $mfaEnforcementCoverageChartRows -AltText 'Covered vs Not Covered by Active MFA Enforcement' -WidthPx 540 -HeightPx 280
    if ($null -ne $mfaEnforcementCoverageChartBlock) {
        $blocks.Add($mfaEnforcementCoverageChartBlock) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Active MFA enforcement coverage was not surfaced clearly enough to render a chart from the reviewed data.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'MFA Enforcement Driver Breakdown' -Style 'Heading3')) | Out-Null
    $mfaEnforcementDriverChartBlock = New-CustomerChartImageBlock -ChartType 'HorizontalBar' -Rows $mfaEnforcementDriverChartRows -AltText 'MFA Enforcement Driver Breakdown' -WidthPx 720 -HeightPx 240
    if ($null -ne $mfaEnforcementDriverChartBlock) {
        $blocks.Add($mfaEnforcementDriverChartBlock) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The reviewed MFA enforcement gap data did not surface enough distinct driver categories to render a chart.' -Style 'Normal')) | Out-Null
    }
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
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0} The summary below shows the most repeated coverage drivers, and the two concise user tables then separate internal members who fall outside the active MFA include scope from those who are explicitly excluded by enabled policies." -f $mfaGapNarrativeText) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Common Coverage Drivers' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items $mfaGapDriverSummaryItems)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Internal Member Users Outside Active MFA Include Scope' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'These users are enabled internal members who were not surfaced inside the active MFA include scope and did not show an explicit exclusion-policy hit in the reviewed data. This bucket is usually the fastest way to find populations that were never targeted cleanly enough by the baseline.' -Style 'Normal')) | Out-Null
    if ($mfaOutsideIncludeDetailRows.Count -gt 0) {
        $blocks.Add((New-CustomerWordTableBlock -Headers @('Display Name', 'User Principal Name') -Rows $mfaOutsideIncludeDetailRows)) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'No internal member users outside the active MFA include scope were surfaced in the reviewed data.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Internal Member Users Explicitly Excluded from Active MFA Policies' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'These users are enabled internal members who surfaced one or more explicit exclusion-policy hits in the reviewed MFA baseline. This bucket matters because exclusions can be intentional break-glass design, but they should be few, documented, and reviewed deliberately.' -Style 'Normal')) | Out-Null
    if ($mfaExcludedDetailRows.Count -gt 0) {
        $blocks.Add((New-CustomerWordTableBlock -Headers @('Display Name', 'User Principal Name') -Rows $mfaExcludedDetailRows)) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'No internal member users explicitly excluded from active MFA policies were surfaced in the reviewed data.' -Style 'Normal')) | Out-Null
    }
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
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recent vs Stale Admins' -Style 'Heading3')) | Out-Null
    $adminActivityChartBlock = New-CustomerChartImageBlock -ChartType 'Donut' -Rows $adminActivityChartRows -AltText 'Recent vs Stale Admins' -WidthPx 540 -HeightPx 280
    if ($null -ne $adminActivityChartBlock) {
        $blocks.Add($adminActivityChartBlock) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Admin activity timing was not surfaced clearly enough to render a chart from the reviewed data.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Admins Covered vs Not Covered by Active MFA Enforcement' -Style 'Heading3')) | Out-Null
    $adminMfaCoverageChartBlock = New-CustomerChartImageBlock -ChartType 'Donut' -Rows $adminMfaCoverageChartRows -AltText 'Admins Covered vs Not Covered by Active MFA Enforcement' -WidthPx 540 -HeightPx 280
    if ($null -ne $adminMfaCoverageChartBlock) {
        $blocks.Add($adminMfaCoverageChartBlock) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Admin MFA enforcement coverage was not surfaced clearly enough to render a chart from the reviewed data.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Admin registration gaps remain separate from admin enforcement gaps in this report. An admin can be enrolled but still sit outside the active enforcement baseline, and an unenrolled admin remains a readiness issue even where policy targeting is stronger.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Why This Matters' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $identityConsultativeSummary) { $identityConsultativeSummary.Narrative } else { 'Privileged access review matters most where stale administrative access and broad standing privileges begin to accumulate together.' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($(if ($null -ne $identityConsultativeSummary) { $identityConsultativeSummary.RecommendationSupport } else { 'This section supports the identity and access recommendations in 4.0.' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Service Accounts' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current review did not surface a dedicated service-account inventory for privileged roles, but mailbox-based SMTP activity and application-related send patterns still provide useful clues about long-lived non-user access. Where service identities remain active, they should be reviewed with the same ownership and lifecycle discipline applied to privileged users.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('DisplayName', 'UserPrincipalName', 'Send Count', 'Last Activity') -Rows $smtpRelayServiceAccountRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Global Administrator Accounts' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items $globalAdminRecommendationItems)) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('DisplayName', 'Created', 'UserPrincipalName', 'LastSignIn') -Rows $globalAdminTableRows)) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '9.0 Exchange Online: Mailboxes and Storage Overview' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'A review was performed on mailbox usage and the distribution of recipient objects. The primary goals of this assessment were to evaluate current storage patterns, identify resources that are no longer active, and document where mailbox lifecycle governance is becoming difficult to manage cleanly.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Data Footprint by Workload' -Style 'Heading2')) | Out-Null
    $dataFootprintChartBlock = New-CustomerChartImageBlock -ChartType 'HorizontalBar' -Rows $dataFootprintChartRows -AltText 'Data Footprint by Workload' -WidthPx 720 -HeightPx 260
    if ($null -ne $dataFootprintChartBlock) {
        $blocks.Add($dataFootprintChartBlock) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'The reviewed storage signals were not complete enough to render a consolidated workload-footprint chart.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Teams storage is represented here through team-connected SharePoint site storage because separate Teams message-storage size is not collected in this report path.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '9.1 Recipient and Mailbox Footprint' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recipient and Mailbox Footprint' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Signal', 'Current State') -Rows $messagingSnapshotRows)) | Out-Null
    $recipientMixChartBlock = New-CustomerChartImageBlock -ChartType 'Donut' -Rows $recipientMixChartRows -AltText 'Recipient Mix by Type' -WidthPx 540 -HeightPx 280
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recipient Mix by Type' -Style 'Heading3')) | Out-Null
    if ($null -ne $recipientMixChartBlock) {
        $blocks.Add($recipientMixChartBlock) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recipient mix was not surfaced clearly enough to render a chart from the reviewed data.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Top Recipient Domains' -Style 'Heading3')) | Out-Null
    $recipientDomainChartBlock = New-CustomerChartImageBlock -ChartType 'HorizontalBar' -Rows $recipientDomainChartRows -AltText 'Top Recipient Domains' -WidthPx 720 -HeightPx 240
    if ($null -ne $recipientDomainChartBlock) {
        $blocks.Add($recipientDomainChartBlock) | Out-Null
    }
    if ($topRecipientDomainSummaryRows.Count -gt 0) {
        $blocks.Add((New-CustomerWordTableBlock -Headers @('Domain', 'Total', 'Primary', 'Alias-only') -Rows $topRecipientDomainSummaryRows)) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Top recipient domain concentration was not surfaced clearly enough for a summary table.' -Style 'Normal')) | Out-Null
    }
    if ($topSenderTableRows.Count -gt 0) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Top Senders' -Style 'Heading3')) | Out-Null
        $blocks.Add((New-CustomerWordTableBlock -Headers @('Display Name', 'User Principal Name', 'Send Count', 'Last Activity') -Rows $topSenderTableRows)) | Out-Null
    }
    if ($topReceiverTableRows.Count -gt 0) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Top Receivers' -Style 'Heading3')) | Out-Null
        $blocks.Add((New-CustomerWordTableBlock -Headers @('Display Name', 'User Principal Name', 'Receive Count', 'Last Activity') -Rows $topReceiverTableRows)) | Out-Null
    }
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
    if ($inactiveMailboxExampleItems.Count -gt 0) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'Inactive Mailbox Examples' -Style 'Heading4')) | Out-Null
        $blocks.Add((New-CustomerWordListBlock -Items $inactiveMailboxExampleItems)) | Out-Null
    }
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
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Largest Collaboration Sites to Review' -Style 'Heading2')) | Out-Null
    if ($largestCollaborationSiteListItems.Count -gt 0) {
        $blocks.Add((New-CustomerWordListBlock -Items $largestCollaborationSiteListItems)) | Out-Null
    }
    else {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'No large SharePoint sites were surfaced clearly enough in the reviewed data for a focused review list.' -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("External sharing remains an important part of the collaboration posture in this tenant. Tenant-level sharing is currently shown as {0}, the default sharing link type is {1}, and {2} reviewed SharePoint site(s) surfaced an external-sharing capability in the source data. Team-connected sites account for {3} of the reviewed SharePoint locations, which reinforces how closely SharePoint governance is tied to broader Teams and group ownership patterns." -f $tenantSharingCapabilityText, $defaultSharingLinkTypeText, $sharePointSitesExternalSharingEnabled, $teamConnectedSites) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The external-sharing summary also shows sharing domain restriction mode as {0}, with {1} site-level sharing override(s) identified in the reviewed site inventory. That combination is important because it shows whether external exposure is being controlled only at the tenant level or is also being shaped materially by site-level exceptions." -f $sharingDomainRestrictionModeText, $siteOverrideCountText) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The broader SharePoint tenant settings also show whether operational controls are reinforcing the sharing baseline or leaving it to stand on its own. The current source records legacy auth protocols as {0}, custom app authentication disabled as {1}, unmanaged sync restriction as {2}, and deleted-user personal site retention as {3} day(s). Those settings matter because external collaboration risk is shaped not only by link defaults, but also by sync controls, client behavior, and how long stale personal content remains in the tenant." -f (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsLegacyAuthProtocolsEnabled')) -Default 'Not surfaced in current source'), (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('DisableCustomAppAuthentication')) -Default 'Not surfaced in current source'), (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('IsUnmanagedSyncAppForTenantRestricted')) -Default 'Not surfaced in current source'), (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('DeletedUserPersonalSiteRetentionPeriodInDays')) -Default 'Not surfaced in current source')) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'External Exposure Review' -Style 'Heading2')) | Out-Null
    $externalExposureCategoryChartBlock = New-CustomerChartImageBlock -ChartType 'Donut' -Rows $externalExposureCategoryChartRows -AltText 'External Exposure by Category' -WidthPx 540 -HeightPx 280
    if ($null -ne $externalExposureCategoryChartBlock) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text 'External Exposure by Category' -Style 'Heading3')) | Out-Null
        $blocks.Add($externalExposureCategoryChartBlock) | Out-Null
    }
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
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Get started with device compliance policies in Microsoft Intune' -Url 'https://learn.microsoft.com/en-us/intune/intune-service/protect/device-compliance-get-started' -WhyItIsRelevant 'Supports the compliance baseline and managed-device observations in the endpoint review.'),
        (New-CustomerDocumentationReference -Title 'Require compliant or hybrid Microsoft Entra joined device' -Url 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-all-users-device-compliance' -WhyItIsRelevant 'Provides Microsoft guidance for tying device state to access-control enforcement.')
    )))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.2 Entra Guest Access Best Practices' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Guest access best practices matter in this tenant because external collaboration, inactive guests, and sharing posture are all part of the current risk picture. These references support the guest lifecycle and external-access observations in the report.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'B2B collaboration fundamentals' -Url 'https://learn.microsoft.com/en-us/entra/external-id/b2b-fundamentals' -WhyItIsRelevant 'Supports guest access governance and external collaboration design.'),
        (New-CustomerDocumentationReference -Title 'Overview of external sharing in SharePoint and OneDrive' -Url 'https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview' -WhyItIsRelevant 'Provides Microsoft guidance for the collaboration-sharing observations in this report.')
    )))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.3 Application Consent and Authentication Methods' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application governance and authentication-method controls both affect how quickly identity exposure can grow in a tenant. The references in this appendix support the application-consent, MFA, and authentication-method observations documented earlier in the assessment.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application User Consent Management' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application user consent should be managed deliberately wherever users can authorize apps to access organizational data. In a tenant where privileged app permissions and consent workflow maturity are already part of the review, consent governance becomes an important control boundary.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recommended Resources' -Style 'Heading4')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Configure the admin consent workflow' -Url 'https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-admin-consent-workflow' -WhyItIsRelevant 'Supports the governance observations around application consent, approval workflow, and control maturity.'),
        (New-CustomerDocumentationReference -Title 'Manage authentication methods' -Url 'https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-authentication-methods-manage' -WhyItIsRelevant 'Relevant where app access, MFA, and modern authentication governance are being reviewed together.')
    )))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Authentication Methods Migration' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Migrating from legacy MFA and SSPR controls to the Authentication Methods policy becomes especially important when the tenant already shows mixed enforcement and uneven registration. The resources below support that transition.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Key Resources for Authentication Migration' -Style 'Heading4')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Manage authentication methods' -Url 'https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-authentication-methods-manage' -WhyItIsRelevant 'Supports migration away from legacy MFA and SSPR policy management.'),
        (New-CustomerDocumentationReference -Title 'Self-service password reset deep dive' -Url 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks' -WhyItIsRelevant 'Provides Microsoft guidance for SSPR design and operational implications.')
    )))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.4 Access and Privileged Identity Management' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Conditional Access and privileged-role governance are two of the strongest levers available for reducing identity risk in this tenant. These references support the policy-state, exclusions, and standing-admin observations documented in the report.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Conditional Access overview' -Url 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview' -WhyItIsRelevant 'Supports the policy coverage, exclusions, and enforcement observations.'),
        (New-CustomerDocumentationReference -Title 'Privileged Identity Management overview' -Url 'https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure' -WhyItIsRelevant 'Supports the recommendations related to privileged role hygiene and reducing standing access.')
    )))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.5 Exchange Online Archives and SMTP Relay' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The Exchange appendix supports the mailbox-growth, archive, forwarding, and relay observations documented in the assessment. These references are useful where mailbox lifecycle and transport controls are both part of the same remediation path.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Exchange Online Archives and Retention Policies' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Archive usage and mailbox retention are part of the same long-term lifecycle story. The references below support archive enablement, mailbox retention, and storage planning.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Learn about retention policies and retention labels' -Url 'https://learn.microsoft.com/en-us/purview/retention' -WhyItIsRelevant 'Supports retention-policy observations and mailbox lifecycle planning.'),
        (New-CustomerDocumentationReference -Title 'On-premises password writeback with self-service password reset' -Url 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-writeback' -WhyItIsRelevant 'Relevant when archive, retention, and lifecycle handling intersect with hybrid identity operations.')
    )))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Third-Party SMTP Relay Configuration' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Relay and application-based sending paths should be documented as service patterns rather than left attached to mailbox accounts by default. These references support the SMTP relay observations in the messaging review.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Control automatic external email forwarding in Microsoft 365' -Url 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-policies-external-email-forwarding' -WhyItIsRelevant 'Supports the forwarding-control observations in the messaging section.'),
        (New-CustomerDocumentationReference -Title 'How to set up a multifunction device or application to send email using Microsoft 365 or Office 365' -Url 'https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-microsoft-365-or-office-365' -WhyItIsRelevant 'Relevant to SMTP relay, connector, and mail-flow exception handling.')
    )))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.6 Microsoft Teams and SharePoint Online' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Teams and SharePoint governance are closely linked in this tenant because collaboration growth, ownership, and sharing posture are moving together. The following Microsoft guidance supports the observations documented in the collaboration sections.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Microsoft Teams Governance' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Teams governance references are most relevant where ownerless workspaces, dormant collaboration spaces, and group-creation controls are part of the current-state review.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Manage who can create Microsoft 365 Groups' -Url 'https://learn.microsoft.com/en-us/microsoft-365/solutions/manage-creation-of-groups' -WhyItIsRelevant 'Supports governance of Teams-connected groups and workspace sprawl.'),
        (New-CustomerDocumentationReference -Title 'Set expiration for Microsoft 365 groups' -Url 'https://learn.microsoft.com/en-us/entra/identity/users/groups-lifecycle' -WhyItIsRelevant 'Relevant to dormant collaboration spaces and lifecycle control.')
    )))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.7 SharePoint Online Collaboration' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'SharePoint and OneDrive guidance is particularly relevant where storage growth, stale content, and external sharing need to be evaluated together rather than as separate issues.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Overview of external sharing in SharePoint and OneDrive' -Url 'https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview' -WhyItIsRelevant 'Supports the collaboration-sharing observations in this report.'),
        (New-CustomerDocumentationReference -Title 'Retention and deletion in OneDrive and SharePoint' -Url 'https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion' -WhyItIsRelevant 'Relevant to stale OneDrive and SharePoint lifecycle handling.')
    )))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.8 DNS DMARC and OneDrive' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This appendix supports the domain-authentication, DMARC, and OneDrive lifecycle observations that surfaced during the review of accepted domains and collaboration services.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'DMARC Records' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'DMARC guidance is especially relevant where recipient volume is concentrated on a small set of accepted domains and mail-authentication posture is uneven across those namespaces.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Set up SPF in Microsoft 365 to help prevent spoofing' -Url 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/set-up-spf-in-office-365-to-help-prevent-spoofing' -WhyItIsRelevant 'Supports SPF and anti-spoofing guidance for the reviewed domains.'),
        (New-CustomerDocumentationReference -Title 'Use DKIM to validate outbound email sent from your custom domain' -Url 'https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dkim-configure' -WhyItIsRelevant 'Supports the observed DKIM posture for custom domains.'),
        (New-CustomerDocumentationReference -Title 'Use DMARC to validate email, setup steps' -Url 'https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dmarc-configure' -WhyItIsRelevant 'Supports DMARC record design, reporting, and policy progression.')
    )))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'OneDrive' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'OneDrive guidance is relevant where user lifecycle, delegated stewardship, and stale personal content repositories are part of the current-state review.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Retention and deletion in OneDrive and SharePoint' -Url 'https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion' -WhyItIsRelevant 'Relevant to stale OneDrive lifecycle handling and post-departure content management.')
    )))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.9 Retention Policies and Data Loss Prevention' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Retention and DLP guidance becomes especially important where the tenant shows partial retention visibility, but not enough evidence to confirm a mature cross-workload lifecycle and data-protection program.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Additional Resources' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Learn about retention policies and retention labels' -Url 'https://learn.microsoft.com/en-us/purview/retention' -WhyItIsRelevant 'Supports retention-policy observations and the need for documented lifecycle controls.'),
        (New-CustomerDocumentationReference -Title 'Learn about data loss prevention' -Url 'https://learn.microsoft.com/en-us/purview/dlp-learn-about-dlp' -WhyItIsRelevant 'Provides Microsoft guidance for DLP and data-protection governance.')
    )))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text '15.9 Pass-Through Authentication and Password Writeback' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Hybrid identity references are included here because directory synchronization, PTA, password writeback, and SSPR visibility all influence how identity operations can be supported safely and consistently.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Pass-Through Authentication (PTA)' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'PTA guidance is relevant anywhere the tenant depends on on-premises credential validation or is still deciding between hybrid sign-in approaches. These references support the hybrid identity observations surfaced in the report.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Microsoft Entra Connect: Pass-through Authentication' -Url 'https://learn.microsoft.com/en-us/azure/active-directory/hybrid/how-to-connect-pta' -WhyItIsRelevant 'Supports pass-through authentication design, agent requirements, and operational considerations.'),
        (New-CustomerDocumentationReference -Title 'Microsoft Entra Connect: User sign-in' -Url 'https://learn.microsoft.com/en-us/azure/active-directory/hybrid/plan-connect-user-signin' -WhyItIsRelevant 'Provides comparison guidance for hybrid sign-in options and role selection.')
    )))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Password Writeback with AD Sync' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Password writeback and SSPR references are included because they directly affect password recovery, hybrid lifecycle handling, and end-user support workflows.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordListBlock -Items @(Convert-CustomerDocumentationReferencesToListItems -References @(
        (New-CustomerDocumentationReference -Title 'Self-service password reset deep dive' -Url 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks' -WhyItIsRelevant 'Supports the password reset and SSPR observations in the report.'),
        (New-CustomerDocumentationReference -Title 'On-premises password writeback with self-service password reset' -Url 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-writeback' -WhyItIsRelevant 'Relevant to password writeback and hybrid credential-management considerations.'),
        (New-CustomerDocumentationReference -Title 'Enable Microsoft Entra password writeback' -Url 'https://learn.microsoft.com/en-us/azure/active-directory/authentication/tutorial-enable-sspr-writeback' -WhyItIsRelevant 'Provides implementation guidance where password writeback is part of the target operating model.')
    )))) | Out-Null
    return @($blocks.ToArray())
}
