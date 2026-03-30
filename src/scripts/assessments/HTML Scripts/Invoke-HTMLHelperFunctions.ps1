#region Helper Functions

function Convert-AssessmentHtmlArray {
    <#
    .SYNOPSIS
        Ensures input is converted to array format. Handles hashtables by extracting values.
    #>
    param($InputObject)

    function Test-IsObjectMap {
        param([AllowNull()]$Value)

        if ($null -eq $Value) {
            return $false
        }

        if ($Value -is [System.Collections.IDictionary]) {
            return $true
        }

        if (-not ($Value -is [pscustomobject])) {
            return $false
        }

        $properties = @($Value.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' })
        if ($properties.Count -eq 0) {
            return $false
        }

        if ($properties.Count -eq 1 -and $properties[0].Name -in @('Summary', 'Configuration', 'Data', 'Value')) {
            return $false
        }

        $nameKeyHitCount = 0
        $complexValueCount = 0
        foreach ($property in $properties) {
            if ([string]$property.Name -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$' -or
                [string]$property.Name -match '@' -or
                [string]$property.Name -match '^[A-Za-z0-9._%-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') {
                $nameKeyHitCount++
            }

            $propertyValue = $property.Value
            if ($null -eq $propertyValue) {
                continue
            }
            if ($propertyValue -is [System.Collections.IDictionary]) {
                $complexValueCount++
                continue
            }
            if ($propertyValue -is [pscustomobject]) {
                $complexValueCount++
                continue
            }
            if ($propertyValue -is [System.Collections.IEnumerable] -and -not ($propertyValue -is [string])) {
                $complexValueCount++
                continue
            }
        }

        if ($nameKeyHitCount -ge [Math]::Min(2, $properties.Count)) {
            return $true
        }

        if ($complexValueCount -ge [Math]::Min(2, $properties.Count)) {
            return $true
        }

        return $false
    }

    function Convert-ToArrayLocal {
        param([AllowNull()]$Value)

        if ($null -eq $Value) {
            return @()
        }

        if ($Value -is [System.Collections.IDictionary]) {
            return @($Value.Values)
        }

        if ($Value -is [string]) {
            return @($Value)
        }

        if ($Value -is [pscustomobject]) {
            $properties = @($Value.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' })

            if ($properties.Count -eq 1 -and $properties[0].Name -in @('Summary', 'Configuration', 'Data', 'Value')) {
                return Convert-ToArrayLocal -Value $properties[0].Value
            }

            if (Test-IsObjectMap -Value $Value) {
                $rows = foreach ($property in $properties) {
                    if ($null -ne $property.Value) {
                        $property.Value
                    }
                }
                return @($rows)
            }
        }

        if ($Value -is [System.Collections.IEnumerable]) {
            return @($Value)
        }

        return @($Value)
    }

    if (Get-Command -Name Convert-ArrayaObjectToArray -ErrorAction SilentlyContinue) {
        try {
            $converted = @(Convert-ArrayaObjectToArray -InputObject $InputObject)
            if (-not (Test-IsObjectMap -Value $InputObject)) {
                return $converted
            }
            if ($converted.Count -gt 1) {
                return $converted
            }
            if ($converted.Count -eq 1 -and $converted[0] -ne $InputObject) {
                return $converted
            }
        }
        catch {
            # Fall through to local conversion rules when the shared converter is unavailable or incompatible.
        }
    }

    return Convert-ToArrayLocal -Value $InputObject
}

function Test-AssessmentHtmlBlankValue {
    <#
    .SYNOPSIS
        Checks whether a value should be treated as blank for HTML display.
    #>
    param(
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return $true
    }

    if ($Value -is [string]) {
        return [string]::IsNullOrWhiteSpace($Value)
    }

    return $false
}

function Format-AssessmentHtmlNumber {
    <#
    .SYNOPSIS
        Formats numbers with thousand separators and optional decimal places.
    #>
    param(
        [AllowNull()]
        $Number,
        
        [int]$DecimalPlaces = 0
    )
    
    if (Test-AssessmentHtmlBlankValue -Value $Number) { return 'N/A' }
    
    try {
        $num = [double]$Number
        if ($DecimalPlaces -eq 0) {
            return $num.ToString('N0')
        } else {
            return $num.ToString("N$DecimalPlaces")
        }
    } catch {
        return $Number.ToString()
    }
}

function Format-AssessmentHtmlDataSize {
    param(
        [AllowNull()]
        [double]$SizeInGB = 0
    )

    if ($null -eq $SizeInGB) {
        return 'N/A'
    }

    if ($SizeInGB -lt 0) {
        return 'N/A'
    }

    if ($SizeInGB -ge 1048576) {
        $sizeInPB = $SizeInGB / 1048576
        return "$(Format-AssessmentHtmlNumber $sizeInPB -DecimalPlaces 2) PB"
    }

    if ($SizeInGB -ge 1024) {
        $sizeInTB = $SizeInGB / 1024
        return "$(Format-AssessmentHtmlNumber $sizeInTB -DecimalPlaces 2) TB"
    }

    if ($SizeInGB -ge 1) {
        return "$(Format-AssessmentHtmlNumber $SizeInGB -DecimalPlaces 2) GB"
    }

    return "$(Format-AssessmentHtmlNumber ($SizeInGB * 1024) -DecimalPlaces 2) MB"
}

function Format-AssessmentHtmlPercentage {
    <#
    .SYNOPSIS
        Formats a value as percentage.
    #>
    param(
        [AllowNull()]
        $Value,
        
        [int]$DecimalPlaces = 1
    )
    
    if (Test-AssessmentHtmlBlankValue -Value $Value) { return 'N/A' }
    
    try {
        $num = [double]$Value
        return $num.ToString("P$DecimalPlaces")
    } catch {
        return $Value.ToString()
    }
}

function Resolve-LicenseInventoryRecord {
    <#
    .SYNOPSIS
        Returns a normalized license inventory record, using shared helper when available.
    #>
    param(
        [AllowNull()]
        $License,
        [int]$UserCount = 0
    )

    if (Get-Command -Name Get-LicenseInventoryRecord -ErrorAction SilentlyContinue) {
        return Get-LicenseInventoryRecord -License $License -UserCount $UserCount
    }

    if ($null -eq $License) {
        return [PSCustomObject]@{
            SkuPartNumber   = $null
            SkuFriendlyName = $null
            PurchasedUnits  = 0
            ConsumedUnits   = 0
            RemainingUnits  = 0
            Utilization     = 0
            IsPaid          = $false
        }
    }

    function Get-LicenseField {
        param(
            [AllowNull()]$Record,
            [Parameter(Mandatory)][string[]]$Names
        )

        foreach ($name in $Names) {
            if ($Record -is [System.Collections.IDictionary] -and $Record.Contains($name)) {
                return $Record[$name]
            }
            if ($Record.PSObject -and $Record.PSObject.Properties[$name]) {
                return $Record.PSObject.Properties[$name].Value
            }
        }
        return $null
    }

    function Convert-ToInt {
        param([AllowNull()]$Value)

        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
            return 0
        }

        $parsed = 0
        if ([int]::TryParse(([string]$Value), [ref]$parsed)) {
            return $parsed
        }

        try { return [int][double]$Value } catch { return 0 }
    }

    $skuPartNumber = [string](Get-LicenseField -Record $License -Names @('SkuPartNumber', 'SkuPart', 'Sku', 'SkuId'))
    $skuFriendlyName = [string](Get-LicenseField -Record $License -Names @('SkuFriendlyName', 'SkuDisplayName', 'ProductName', 'DisplayName'))
    if ([string]::IsNullOrWhiteSpace($skuFriendlyName)) {
        $skuFriendlyName = $skuPartNumber
    }

    $purchasedUnits = Convert-ToInt (Get-LicenseField -Record $License -Names @('PurchasedUnits', 'PrepaidUnitsEnabled', 'TotalUnits', 'ActiveUnits'))
    $consumedUnits = Convert-ToInt (Get-LicenseField -Record $License -Names @('ConsumedUnits', 'AssignedUnits', 'Consumed'))
    $remainingUnits = Convert-ToInt (Get-LicenseField -Record $License -Names @('RemainingUnits'))
    if ($remainingUnits -eq 0 -and $purchasedUnits -gt 0) {
        $remainingUnits = $purchasedUnits - $consumedUnits
    }

    $utilization = 0.0
    if ($purchasedUnits -gt 0) {
        $utilization = [math]::Round((($consumedUnits / $purchasedUnits) * 100), 2)
    }

    $isPaid = $false
    $licenseIsPaidField = Get-LicenseField -Record $License -Names @('IsPaid')
    if ($null -ne $licenseIsPaidField) {
        $isPaid = [bool]$licenseIsPaidField
    }
    elseif ($purchasedUnits -gt 0) {
        $isLikelyFreeFromScale = ($UserCount -gt 0 -and $purchasedUnits -gt ($UserCount * 5))
        $isPaid = -not $isLikelyFreeFromScale
    }

    return [PSCustomObject]@{
        SkuPartNumber   = $skuPartNumber
        SkuFriendlyName = $skuFriendlyName
        PurchasedUnits  = $purchasedUnits
        ConsumedUnits   = $consumedUnits
        RemainingUnits  = $remainingUnits
        Utilization     = $utilization
        IsPaid          = $isPaid
    }
}

function New-HtmlTable {
    <#
    .SYNOPSIS
        Generates HTML table from array of objects with optional risk highlighting.
    #>
    param(
        [AllowNull()]
        [array]$Data,
        
        [Parameter(Mandatory)]
        [string[]]$Columns,
        
        [hashtable]$ColumnHeaders,
        
        [hashtable]$RiskColumns,

        [hashtable]$ValueFormatters,
        
        [string]$EmptyMessage = 'No data available',
        
        [string]$CssClass = 'data-table'
    )
    
    if ($null -eq $Data -or $Data.Count -eq 0) {
        return "<div class='empty-state'>$EmptyMessage</div>"
    }
    
    $html = "<table class='$CssClass'>`n<thead><tr>"
    
    # Headers
    foreach ($col in $Columns) {
        $header = if ($ColumnHeaders -and $ColumnHeaders.ContainsKey($col)) {
            $ColumnHeaders[$col]
        } else {
            $col
        }
        $html += "<th>$header</th>"
    }
    $html += "</tr></thead>`n<tbody>"
    
    # Rows
    foreach ($row in $Data) {
        $html += "<tr>"
        foreach ($col in $Columns) {
            $value = $row.$col
            
            # Format value
            if ($ValueFormatters -and $ValueFormatters.ContainsKey($col)) {
                $formatter = $ValueFormatters[$col]
                if ($formatter -is [scriptblock]) {
                    $formattedValue = & $formatter $value $row
                    $displayValue = [System.Web.HttpUtility]::HtmlEncode([string]$formattedValue)
                }
                else {
                    $displayValue = [System.Web.HttpUtility]::HtmlEncode([string]$value)
                }
            } elseif (Test-AssessmentHtmlBlankValue -Value $value) {
                $displayValue = 'N/A'
            } elseif ($value -is [datetime]) {
                $displayValue = $value.ToString('yyyy-MM-dd')
            } elseif ($value -is [bool]) {
                $displayValue = if ($value) { '✓' } else { '✗' }
            } elseif ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                $listValues = @($value | ForEach-Object {
                    if (-not (Test-AssessmentHtmlBlankValue -Value $_)) {
                        [string]$_
                    }
                })
                $displayValue = if ($listValues.Count -gt 0) {
                    [System.Web.HttpUtility]::HtmlEncode(($listValues -join ', '))
                } else {
                    'N/A'
                }
            } else {
                $displayValue = [System.Web.HttpUtility]::HtmlEncode($value.ToString())
            }
            
            # Check for risk highlighting
            $cellClass = ''
            if ($RiskColumns -and $RiskColumns.ContainsKey($col)) {
                $riskCheck = $RiskColumns[$col]
                if ($riskCheck -is [scriptblock]) {
                    if (& $riskCheck $value $row) {
                        $cellClass = ' class="risk-cell"'
                    }
                }
            }
            
            $html += "<td$cellClass>$displayValue</td>"
        }
        $html += "</tr>`n"
    }
    
    $html += "</tbody></table>"
    return $html
}

function New-KpiCard {
    <#
    .SYNOPSIS
        Creates a KPI card with value and optional change indicator.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        
        [Parameter(Mandatory)]
        [object]$Value,
        
        [string]$Subtitle,
        
        [ValidateSet('', 'up', 'down', 'neutral')]
        [string]$Trend = '',
        
        [ValidateSet('default', 'success', 'warning', 'danger')]
        [string]$Theme = 'default'
    )
    
    $trendIcon = switch ($Trend) {
        'up' { '↗️' }
        'down' { '↘️' }
        'neutral' { '→' }
        default { '' }
    }
    
    $subtitleHtml = if ($Subtitle) {
        "<div class='kpi-subtitle'>$Subtitle</div>"
    } else { '' }

    $displayValue = if ($null -eq $Value) {
        'N/A'
    } elseif ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        (($Value | ForEach-Object { $_.ToString() }) -join ', ')
    } else {
        $Value.ToString()
    }
    
    return @"
<div class="kpi-card kpi-$Theme">
    <div class="kpi-title">$Title</div>
    <div class="kpi-value">$displayValue $trendIcon</div>
    $subtitleHtml
</div>
"@
}

function New-CalloutBox {
    <#
    .SYNOPSIS
        Creates a callout/alert box.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('info', 'warning', 'danger', 'success')]
        [string]$Type,
        
        [Parameter(Mandatory)]
        [string]$Title,
        
        [Parameter(Mandatory)]
        [string]$Content
    )
    
    $iconMap = @{
        'info' = 'ℹ️'
        'warning' = '⚠️'
        'danger' = '🔴'
        'success' = '✅'
    }
    
    return @"
<div class="callout callout-$Type">
    <div class="callout-header">$($iconMap[$Type]) $Title</div>
    <div class="callout-body">$Content</div>
</div>
"@
}

#endregion

#region Analysis Functions

function Get-LicenseAnalysis {
    <#
    .SYNOPSIS
        Analyzes license data with refined severity levels
        
    .DESCRIPTION
        Critical: Over capacity (negative remaining)
        Warning: At capacity (0 remaining) or high utilization (>85% with licenses available)
        Info: Overall utilization summary
    #>
    param(
        [array]$Licenses,
        [int]$UserCount = 0
    )
    
    $findings = @()
    $criticalFindings = @()
    $warningFindings = @()
    $infoFindings = @()
    
    # Process licenses with consistent paid/free classification
    $processedLicenses = $Licenses | ForEach-Object {
        Resolve-LicenseInventoryRecord -License $_ -UserCount $UserCount
    }
    
    # Filter paid licenses
    $paidLicenses = $processedLicenses | Where-Object { $_.IsPaid -eq $true }
    
    # Analyze findings with NEW logic
    foreach ($lic in $paidLicenses) {
        
        # 🔴 CRITICAL: Over capacity (negative remaining)
        $licenseName = if ($lic.SkuFriendlyName) { $lic.SkuFriendlyName } else { $lic.SkuPartNumber }
        if ($lic.RemainingUnits -lt 0 -and $lic.PurchasedUnits -gt 0) {
            $criticalFindings += @{
                Type = 'Risk'
                Category = 'Critical'
                Message = "License '$licenseName' is OVER capacity by $([math]::Abs($lic.RemainingUnits)) licenses"
                Anchor = 'licenses'
                Priority = 1
            }
        }
        # ⚠️ WARNING: At capacity (0 remaining)
        elseif ($lic.RemainingUnits -eq 0 -and $lic.PurchasedUnits -gt 0) {
            $warningFindings += @{
                Type = 'Warning'
                Category = 'At Capacity'
                Message = "License '$licenseName' has 0 remaining (fully allocated)"
                Anchor = 'licenses'
                Priority = 2
            }
        }
        # ⚠️ WARNING: High utilization but has some remaining
        elseif ($lic.Utilization -ge 85 -and $lic.RemainingUnits -gt 0) {
            $warningFindings += @{
                Type = 'Warning'
                Category = 'High Utilization'
                Message = "License '$licenseName' is $([math]::Round($lic.Utilization,1))% utilized ($($lic.RemainingUnits) remaining)"
                Anchor = 'licenses'
                Priority = 3
            }
        }
    }
    
    # Overall utilization (info level)
    $totalPurchased = ($paidLicenses | Measure-Object -Property PurchasedUnits -Sum).Sum
    $totalConsumed = ($paidLicenses | Measure-Object -Property ConsumedUnits -Sum).Sum
    
    if ($totalPurchased -gt 0) {
        $overallUtil = ($totalConsumed / $totalPurchased) * 100
        if ($overallUtil -ge 85) {
            $infoFindings += @{
                Type = 'Warning'
                Category = 'Overall Utilization'
                Message = "Overall license utilization is $([math]::Round($overallUtil,1))% across all paid licenses"
                Anchor = 'licenses'
                Priority = 4
            }
        }
    }
    
    # Combine findings in priority order
    $findings = $criticalFindings + $warningFindings + $infoFindings
    
    return @{
        Findings = $findings
        TotalPurchased = $totalPurchased
        TotalConsumed = $totalConsumed
        PaidLicenses = $paidLicenses
        CriticalCount = $criticalFindings.Count
        WarningCount = $warningFindings.Count
    }
}

function Get-MailboxAnalysis {
    <#
    .SYNOPSIS
        Analyzes mailbox data and returns findings
    #>
    param([array]$Mailboxes)
    
    $findings = @()
    $largeMailboxes = @()
    $largeArchiveCount = 0
    $largestArchiveMailbox = $null
    
    foreach ($mbx in $Mailboxes) {
        # Handle N/A values safely
        $mbxSize = 0
        $archiveSize = 0
        
        if ($mbx.MBXSizeGB -ne 'N/A' -and $null -ne $mbx.MBXSizeGB -and $mbx.MBXSizeGB -ne '') {
            try { $mbxSize = [double]$mbx.MBXSizeGB } catch { $mbxSize = 0 }
        }
        
        if ($mbx.ArchiveSizeGB -ne 'N/A' -and $null -ne $mbx.ArchiveSizeGB -and $mbx.ArchiveSizeGB -ne '') {
            try { $archiveSize = [double]$mbx.ArchiveSizeGB } catch { $archiveSize = 0 }
        }
        
        # Large mailbox warnings
        if ($mbxSize -gt $script:DefaultThresholds.MailboxSizeGB) {
            $largeMailboxes += $mbx
        }
        
        # Large archive warnings
        if ($archiveSize -gt $script:DefaultThresholds.ArchiveSizeGB) {
            $largeArchiveCount++
            if (-not $largestArchiveMailbox -or $archiveSize -gt $largestArchiveMailbox.SizeGB) {
                $largestArchiveMailbox = [PSCustomObject]@{
                    DisplayName = [string]$mbx.DisplayName
                    SizeGB = [math]::Round($archiveSize, 1)
                }
            }
        }
    }
    
    # Summary finding for large mailboxes
    if ($largeMailboxes.Count -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large Mailboxes'
            Message = "$($largeMailboxes.Count) mailboxes exceed $($script:DefaultThresholds.MailboxSizeGB) GB"
            Anchor = 'mailboxes'
            Priority = 2
        }
    }

    if ($largeArchiveCount -gt 0) {
        $largestArchiveLabel = if ($largestArchiveMailbox -and -not [string]::IsNullOrWhiteSpace($largestArchiveMailbox.DisplayName)) {
            "'$($largestArchiveMailbox.DisplayName)' at $($largestArchiveMailbox.SizeGB) GB"
        } else {
            'unavailable'
        }
        $findings += @{
            Type = 'Warning'
            Category = 'Large Archives'
            Message = "$largeArchiveCount archive mailbox(es) exceed $($script:DefaultThresholds.ArchiveSizeGB) GB (largest: $largestArchiveLabel)"
            Anchor = 'mailboxes'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
        LargeMailboxes = $largeMailboxes
        LargeArchiveCount = $largeArchiveCount
        LargestArchiveMailbox = $largestArchiveMailbox
    }
}

function Get-ComplianceRetentionAnalysis {
    <#
    .SYNOPSIS
        Summarizes oversized mailbox/archive compliance posture against litigation and retention holds
    #>
    param([array]$Mailboxes)

    function Convert-ToMailboxBoolean {
        param([AllowNull()]$Value)
        if ($null -eq $Value) { return $false }
        if ($Value -is [bool]) { return $Value }
        $valueText = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($valueText)) { return $false }
        switch -Regex ($valueText.ToLowerInvariant()) {
            '^(true|1|yes|y)$' { return $true }
            default {
                try { return [bool]$Value } catch { return $false }
            }
        }
    }

    function Get-MailboxNumericSize {
        param([AllowNull()]$Value)
        if ($null -eq $Value) { return 0 }
        $valueText = [string]$Value
        if ([string]::IsNullOrWhiteSpace($valueText) -or $valueText -eq 'N/A') { return 0 }
        try { return [double]$valueText } catch { return 0 }
    }

    function Get-MailboxHoldCoverageSummary {
        param([array]$Records)

        $summary = [ordered]@{
            TotalOversized          = 0
            LitigationHoldCount     = 0
            RetentionHoldCount      = 0
            RetentionControlCount   = 0
            BothCount               = 0
            UncoveredCount          = 0
            UnknownCoverageCount    = 0
        }

        if (-not $Records -or $Records.Count -eq 0) {
            return [PSCustomObject]$summary
        }

        foreach ($record in $Records) {
            if ($null -eq $record) {
                continue
            }
            $summary.TotalOversized++

            $hasLitProp = ($record.PSObject -and $record.PSObject.Properties['LitigationHoldEnabled'])
            $hasRetentionProp = ($record.PSObject -and $record.PSObject.Properties['RetentionHoldEnabled'])
            $hasDelayHoldProp = ($record.PSObject -and $record.PSObject.Properties['DelayHoldApplied'])
            $hasRetentionPolicyProp = ($record.PSObject -and $record.PSObject.Properties['RetentionPolicy'])
            $hasInPlaceHoldsProp = ($record.PSObject -and $record.PSObject.Properties['InPlaceHolds'])

            $litigationHold = $hasLitProp -and (Convert-ToMailboxBoolean -Value $record.LitigationHoldEnabled)
            $retentionHold = $hasRetentionProp -and (Convert-ToMailboxBoolean -Value $record.RetentionHoldEnabled)
            $delayHold = $hasDelayHoldProp -and (Convert-ToMailboxBoolean -Value $record.DelayHoldApplied)
            $hasRetentionPolicy = $hasRetentionPolicyProp -and -not [string]::IsNullOrWhiteSpace([string]$record.RetentionPolicy)
            $hasInPlaceHolds = $hasInPlaceHoldsProp -and -not [string]::IsNullOrWhiteSpace([string]$record.InPlaceHolds)
            $retentionControl = ($retentionHold -or $delayHold -or $hasRetentionPolicy -or $hasInPlaceHolds)

            if ($litigationHold) { $summary.LitigationHoldCount++ }
            if ($retentionHold) { $summary.RetentionHoldCount++ }
            if ($retentionControl) { $summary.RetentionControlCount++ }
            if ($litigationHold -and $retentionControl) { $summary.BothCount++ }

            $coverageSignalsKnown = ($hasLitProp -or $hasRetentionProp -or $hasDelayHoldProp -or $hasRetentionPolicyProp -or $hasInPlaceHoldsProp)
            if (-not $coverageSignalsKnown) {
                $summary.UnknownCoverageCount++
            }
            elseif (-not $litigationHold -and -not $retentionControl) {
                $summary.UncoveredCount++
            }
        }

        return [PSCustomObject]$summary
    }

    $findings = @()
    if (-not $Mailboxes -or $Mailboxes.Count -eq 0) {
        return @{
            Findings = @()
            PrimaryOversizedSummary = [PSCustomObject]@{}
            ArchiveOversizedSummary = [PSCustomObject]@{}
        }
    }

    $mailboxThreshold = [double]$script:DefaultThresholds.MailboxSizeGB
    $archiveThreshold = [double]$script:DefaultThresholds.ArchiveSizeGB

    $oversizedPrimary = @(
        $Mailboxes | Where-Object {
            (Get-MailboxNumericSize -Value $_.MBXSizeGB) -gt $mailboxThreshold
        }
    )
    $oversizedArchive = @(
        $Mailboxes | Where-Object {
            (Get-MailboxNumericSize -Value $_.ArchiveSizeGB) -gt $archiveThreshold
        }
    )

    $primarySummary = Get-MailboxHoldCoverageSummary -Records $oversizedPrimary
    $archiveSummary = Get-MailboxHoldCoverageSummary -Records $oversizedArchive

    $findings += @{
        Type = 'Info'
        Category = 'Oversized Mailbox Compliance Coverage'
        Message = "Primary > $mailboxThreshold GB: $($primarySummary.TotalOversized) (litigation hold=$($primarySummary.LitigationHoldCount), retention hold=$($primarySummary.RetentionHoldCount), any retention control=$($primarySummary.RetentionControlCount), both=$($primarySummary.BothCount), uncovered=$($primarySummary.UncoveredCount), unknown=$($primarySummary.UnknownCoverageCount)). Archive > $archiveThreshold GB: $($archiveSummary.TotalOversized) (litigation hold=$($archiveSummary.LitigationHoldCount), retention hold=$($archiveSummary.RetentionHoldCount), any retention control=$($archiveSummary.RetentionControlCount), both=$($archiveSummary.BothCount), uncovered=$($archiveSummary.UncoveredCount), unknown=$($archiveSummary.UnknownCoverageCount))."
        Anchor = 'compliance-retention'
        Priority = 2
    }

    $totalUncovered = [int]$primarySummary.UncoveredCount + [int]$archiveSummary.UncoveredCount
    if ($totalUncovered -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Oversized Mailboxes Without Hold Controls'
            Message = "$totalUncovered oversized mailbox/archive object(s) have neither litigation hold nor retention controls."
            Anchor = 'compliance-retention'
            Priority = 2
        }
    }
    elseif (($primarySummary.TotalOversized + $archiveSummary.TotalOversized) -gt 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Oversized Mailbox Hold Coverage'
            Message = 'All oversized mailbox/archive objects have at least one hold signal (litigation hold and/or retention control).'
            Anchor = 'compliance-retention'
            Priority = 3
        }
    }

    $totalUnknown = [int]$primarySummary.UnknownCoverageCount + [int]$archiveSummary.UnknownCoverageCount
    if ($totalUnknown -gt 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Hold Signal Coverage'
            Message = "$totalUnknown oversized object(s) did not expose hold properties in this profile; rerun with a deeper profile for full hold comparison fidelity."
            Anchor = 'compliance-retention'
            Priority = 3
        }
    }

    return @{
        Findings = @($findings)
        PrimaryOversizedSummary = $primarySummary
        ArchiveOversizedSummary = $archiveSummary
    }
}

function Get-TeamsCollaborationAnalysis {
    <#
    .SYNOPSIS
        Analyzes Teams inventory and collaboration posture signals
    #>
    param(
        [array]$Teams,
        [array]$TeamsTopUsers = @(),
        [object]$TeamsVoice = $null
    )

    function Get-TeamCountFromCsv {
        param([AllowNull()]$Value)
        if ($null -eq $Value) { return 0 }
        $raw = [string]$Value
        if ([string]::IsNullOrWhiteSpace($raw)) { return 0 }
        return @(
            $raw -split ',' |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        ).Count
    }

    $findings = @()
    $teams = @($Teams | Where-Object { $null -ne $_ })
    $teamCount = $teams.Count

    if ($teamCount -eq 0) {
        if ($TeamsTopUsers.Count -gt 0) {
            $findings += @{
                Type = 'Info'
                Category = 'Teams Inventory Coverage'
                Message = 'Teams activity telemetry was collected, but Teams inventory details were not available in this run.'
                Anchor = 'teams-collaboration'
                Priority = 3
            }
        }

        return @{
            Findings = @($findings)
        }
    }

    $archivedTeams = @($teams | Where-Object { $_.PSObject.Properties['IsArchived'] -and $_.IsArchived -eq $true })
    $privateChannelTeams = @($teams | Where-Object {
        $privateCount = if ($_.PSObject.Properties['PrivateChannelCount']) { [int]$_.PrivateChannelCount } else { Get-TeamCountFromCsv -Value $_.PrivateChannels }
        $privateCount -gt 0
    })
    $totalChannels = @($teams | Measure-Object -Property TotalChannels -Sum).Sum
    $unownedTeams = @($teams | Where-Object { $_.PSObject.Properties['OwnershipState'] -and [string]$_.OwnershipState -eq 'Unowned' })
    $unknownOwnerTeams = @($teams | Where-Object { $_.PSObject.Properties['OwnershipState'] -and [string]$_.OwnershipState -eq 'Unknown' })

    $findings += @{
        Type = 'Info'
        Category = 'Teams Inventory Baseline'
        Message = "$teamCount team(s) discovered with $totalChannels total channel(s); archived=$($archivedTeams.Count); teams with private channels=$($privateChannelTeams.Count)."
        Anchor = 'teams-collaboration'
        Priority = 3
    }

    if ($unownedTeams.Count -gt 0) {
        $severity = if ($unownedTeams.Count -ge 5) { 'Risk' } else { 'Warning' }
        $findings += @{
            Type = $severity
            Category = 'Teams Ownership'
            Message = "$($unownedTeams.Count) team(s) appear unowned based on linked Microsoft 365 group ownership."
            Anchor = 'teams-collaboration'
            Priority = 2
        }
    }
    elseif ($unknownOwnerTeams.Count -gt 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Teams Ownership'
            Message = "Ownership could not be confirmed for $($unknownOwnerTeams.Count) team(s) in this profile/run."
            Anchor = 'teams-collaboration'
            Priority = 3
        }
    }

    if ($teamCount -gt 0) {
        $privateRatio = [math]::Round((($privateChannelTeams.Count / [double]$teamCount) * 100), 1)
        if ($privateRatio -ge 40) {
            $findings += @{
                Type = 'Warning'
                Category = 'Private Channel Footprint'
                Message = "$privateRatio% of teams use private channels ($($privateChannelTeams.Count)/$teamCount). Review lifecycle and eDiscovery governance coverage."
                Anchor = 'teams-collaboration'
                Priority = 2
            }
        }
    }

    if ($TeamsTopUsers.Count -eq 0 -and $teamCount -gt 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Usage Telemetry Coverage'
            Message = 'Teams inventory exists, but no top Teams activity users were captured in this period.'
            Anchor = 'teams-collaboration'
            Priority = 3
        }
    }

    if ($TeamsVoice -and $TeamsVoice.PSObject -and $TeamsVoice.PSObject.Properties['Summary'] -and $TeamsVoice.Summary) {
        $voiceSummary = $TeamsVoice.Summary
        if (
            $voiceSummary.PSObject.Properties['DataSource'] -and
            [string]$voiceSummary.DataSource -eq 'GraphLicenseInference'
        ) {
            $findings += @{
                Type = 'Info'
                Category = 'Teams Voice Data Source'
                Message = "Teams voice metrics are inferred from licensing in this run (source=$($voiceSummary.DataSource)); PSTN/calling detail requires Teams PowerShell connectivity."
                Anchor = 'teams-collaboration'
                Priority = 3
            }
        }
    }

    return @{
        Findings = @($findings)
    }
}

function Get-DomainAnalysis {
    <#
    .SYNOPSIS
        Analyzes domain configuration and returns findings
    #>
    param(
        [array]$Domains,
        [object]$SpamFilteringSummary,
        [object]$SMTPRelaySummary
    )
    
    $findings = @()
    $mailEnabledCustomDomainCount = 0
    $spfPassingDomainCount = 0
    $dmarcConfiguredDomainCount = 0
    $dmarcEnforcedDomainCount = 0
    $dkimCompleteDomainCount = 0
    
    foreach ($domain in $Domains) {
        $domainName = [string]$domain.Domain
        $isTenantServiceDomain = $domainName -match '(?i)\.onmicrosoft\.com$'
        $isCustomDomain = -not $isTenantServiceDomain
        $hasRecipientUsage = $false
        try {
            $hasRecipientUsage = ([int]$domain.TotalDomainRecipients -gt 0)
        } catch {}

        # Check verification
        if (-not $domain.Verified -and $script:DefaultThresholds.DomainVerificationRequired) {
            $findings += @{
                Type = 'Risk'
                Category = 'Domain Verification'
                Message = "Domain '$($domain.Domain)' is not verified"
                Anchor = 'domains'
                Priority = 1
            }
        }
        
        # Check MX records
        if ($script:DefaultThresholds.MXRecordValidation) {
            if ($domain.Office365MailExchanger -eq $false) {
                $findings += @{
                    Type = 'Info'
                    Category = 'DNS Configuration'
                    Message = "Domain '$($domain.Domain)' MX record does not point to Microsoft 365"
                    Anchor = 'domains-dns'
                    Priority = 3
                }
            }
        }

        if ($isCustomDomain -and $domain.Verified -and ($domain.Office365MailExchanger -eq $true -or $hasRecipientUsage)) {
            $mailEnabledCustomDomainCount++

            $spfIncludesM365 = $false
            if ($domain.PSObject.Properties['SpfIncludesM365']) {
                $spfIncludesM365 = ($domain.SpfIncludesM365 -eq $true)
            }
            if ($spfIncludesM365) { $spfPassingDomainCount++ }
            if ($domain.PSObject.Properties['SpfIncludesM365'] -and $domain.SpfIncludesM365 -ne $true) {
                $findings += @{
                    Type = 'Warning'
                    Category = 'SPF'
                    Message = "Domain '$($domain.Domain)' does not show an SPF record including spf.protection.outlook.com"
                    Anchor = 'domains-dns'
                    Priority = 2
                }
            }
            if ($domain.PSObject.Properties['SpfPolicyMode'] -and $domain.SpfPolicyMode) {
                if ([string]$domain.SpfPolicyMode -eq 'AllowAll (+all)') {
                    $findings += @{
                        Type = 'Risk'
                        Category = 'SPF'
                        Message = "Domain '$($domain.Domain)' SPF record is configured as +all (allow all), which weakens spoof protection"
                        Anchor = 'domains-dns'
                        Priority = 1
                    }
                }
                elseif ([string]$domain.SpfPolicyMode -eq 'SoftFail (~all)') {
                    $findings += @{
                        Type = 'Info'
                        Category = 'SPF'
                        Message = "Domain '$($domain.Domain)' uses SPF soft-fail (~all); consider hard-fail (-all) after validation"
                        Anchor = 'domains-dns'
                        Priority = 3
                    }
                }
            }

            $dmarcConfigured = $false
            if ($domain.PSObject.Properties['DmarcConfigured']) {
                $dmarcConfigured = ($domain.DmarcConfigured -eq $true)
            }
            if ($dmarcConfigured) { $dmarcConfiguredDomainCount++ }
            if ($domain.PSObject.Properties['DmarcConfigured'] -and $domain.DmarcConfigured -ne $true) {
                $findings += @{
                    Type = 'Warning'
                    Category = 'DMARC'
                    Message = "Domain '$($domain.Domain)' does not show a DMARC policy record"
                    Anchor = 'domains-dns'
                    Priority = 2
                }
            }
            if ($dmarcConfigured -and $domain.PSObject.Properties['DmarcPolicy']) {
                $dmarcPolicy = [string]$domain.DmarcPolicy
                if ($dmarcPolicy -eq 'none') {
                    $findings += @{
                        Type = 'Warning'
                        Category = 'DMARC'
                        Message = "Domain '$($domain.Domain)' DMARC policy is p=none (monitor only); move toward quarantine/reject for anti-spoofing enforcement"
                        Anchor = 'domains-dns'
                        Priority = 2
                    }
                } elseif ($dmarcPolicy -in @('quarantine', 'reject')) {
                    $dmarcEnforcedDomainCount++
                }

                $dmarcPercent = 100
                if ($domain.PSObject.Properties['DmarcPercent']) {
                    try { $dmarcPercent = [int]$domain.DmarcPercent } catch { $dmarcPercent = 100 }
                }
                if ($dmarcPercent -lt 100) {
                    $findings += @{
                        Type = 'Info'
                        Category = 'DMARC'
                        Message = "Domain '$($domain.Domain)' DMARC enforcement scope is pct=$dmarcPercent; increase toward 100 for full anti-spoofing coverage"
                        Anchor = 'domains-dns'
                        Priority = 3
                    }
                }
            }

            if ($domain.PSObject.Properties['DkimSelectorsConfigured']) {
                $dkimSelectorCount = 0
                try { $dkimSelectorCount = [int]$domain.DkimSelectorsConfigured } catch { $dkimSelectorCount = 0 }
                if ($dkimSelectorCount -ge 2) { $dkimCompleteDomainCount++ }
                if ($dkimSelectorCount -lt 2) {
                    $findings += @{
                        Type = 'Warning'
                        Category = 'DKIM'
                        Message = "Domain '$($domain.Domain)' has $dkimSelectorCount DKIM selector CNAME record(s) detected (expected: 2)"
                        Anchor = 'domains-dns'
                        Priority = 2
                    }
                }
            }
        }
    }

    if ($mailEnabledCustomDomainCount -gt 0) {
        $spfPct = [math]::Round((($spfPassingDomainCount / $mailEnabledCustomDomainCount) * 100), 1)
        $dmarcPct = [math]::Round((($dmarcConfiguredDomainCount / $mailEnabledCustomDomainCount) * 100), 1)
        $dmarcEnforcedPct = [math]::Round((($dmarcEnforcedDomainCount / $mailEnabledCustomDomainCount) * 100), 1)
        $dkimPct = [math]::Round((($dkimCompleteDomainCount / $mailEnabledCustomDomainCount) * 100), 1)

        $coverageType = if ($dmarcEnforcedPct -lt 60 -or $dkimPct -lt 60 -or $spfPct -lt 80) { 'Warning' } else { 'Info' }
        $coveragePriority = if ($coverageType -eq 'Warning') { 2 } else { 3 }
        $findings += @{
            Type = $coverageType
            Category = 'Email Authentication Coverage'
            Message = "Mail-auth coverage across $mailEnabledCustomDomainCount custom mail domain(s): SPF $spfPct%, DKIM $dkimPct%, DMARC configured $dmarcPct%, DMARC enforcement (quarantine/reject) $dmarcEnforcedPct%"
            Anchor = 'domains-dns'
            Priority = $coveragePriority
        }
    }

    if ($SpamFilteringSummary) {
        $trustedBypassCount = 0
        if ($SpamFilteringSummary.PSObject.Properties['TransportRulesWithTrustedIPs']) {
            try { $trustedBypassCount = [int]$SpamFilteringSummary.TransportRulesWithTrustedIPs } catch { $trustedBypassCount = 0 }
        }
        $uses3rdPartyFiltering = $null
        if ($SpamFilteringSummary.PSObject.Properties['Uses3rdPartyFiltering']) {
            $rawThirdParty = $SpamFilteringSummary.Uses3rdPartyFiltering
            if ($rawThirdParty -is [bool]) {
                $uses3rdPartyFiltering = $rawThirdParty
            } elseif ($null -ne $rawThirdParty) {
                $uses3rdPartyFiltering = ([string]$rawThirdParty -match '(?i)^(yes|true|enabled)$')
            }
        }
        if ($trustedBypassCount -gt 0) {
            $findings += @{
                Type = 'Warning'
                Category = 'Anti-Spoofing Bypass'
                Message = "$trustedBypassCount transport rule or connection-filter trusted IP bypass indicator(s) detected; validate spoof protection exceptions"
                Anchor = 'domains-dns'
                Priority = 2
            }
        }
        $findings += @{
            Type = 'Info'
            Category = 'Anti-Spoofing Controls'
            Message = "Mail-flow anti-spoof controls: trusted bypass indicators=$trustedBypassCount; third-party filtering detected=$(if ($null -eq $uses3rdPartyFiltering) { 'Unknown' } elseif ($uses3rdPartyFiltering) { 'Yes' } else { 'No' })"
            Anchor = 'domains-dns'
            Priority = 3
        }
    }

    if ($SMTPRelaySummary) {
        $smtpAuthEnabled = $false
        if ($SMTPRelaySummary.PSObject.Properties['SMTPAuthEnabled']) {
            $rawSmtpAuth = $SMTPRelaySummary.SMTPAuthEnabled
            if ($rawSmtpAuth -is [bool]) {
                $smtpAuthEnabled = $rawSmtpAuth
            } elseif ($null -ne $rawSmtpAuth) {
                $smtpAuthEnabled = ([string]$rawSmtpAuth -match '(?i)^(yes|true|enabled)$')
            }
        }

        $smtpAuthUsers = 0
        if ($SMTPRelaySummary.PSObject.Properties['SMTPAuthUsers']) {
            try { $smtpAuthUsers = [int]$SMTPRelaySummary.SMTPAuthUsers } catch { $smtpAuthUsers = 0 }
        }

        if ($smtpAuthEnabled -and $smtpAuthUsers -gt 0) {
            $findings += @{
                Type = 'Warning'
                Category = 'SMTP AUTH Exposure'
                Message = "$smtpAuthUsers mailbox(es) still have SMTP AUTH enabled; this can increase impersonation and brute-force attack surface"
                Anchor = 'domains-dns'
                Priority = 2
            }
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-DeviceAnalysis {
    <#
    .SYNOPSIS
        Analyzes device data and returns findings
    #>
    param([array]$Devices)
    
    $findings = @()
    
    if ($Devices.Count -eq 0) {
        return @{ Findings = @() }
    }
    
    # Stale devices
    $staleDevices = $Devices | Where-Object {
        $_.DeviceStale -eq $true
    }
    
    if ($staleDevices.Count -gt 0) {
        $stalePercent = ($staleDevices.Count / $Devices.Count) * 100
        $findings += @{
            Type = 'Warning'
            Category = 'Stale Devices'
            Message = "$($staleDevices.Count) devices ($([math]::Round($stalePercent,1))%) are stale (inactive > $($script:DefaultThresholds.DeviceStaleMonths) months)"
            Anchor = 'devices'
            Priority = 2
        }
    }
    
    # Device compliance
    $compliantDevices = $Devices | Where-Object { $_.IsCompliant -eq $true }
    $compliancePercent = ($compliantDevices.Count / $Devices.Count) * 100
    
    if ($compliancePercent -lt $script:DefaultThresholds.DeviceCompliancePercent) {
        $findings += @{
            Type = 'Risk'
            Category = 'Device Compliance'
            Message = "Device compliance is only $([math]::Round($compliancePercent,1))% (target: $($script:DefaultThresholds.DeviceCompliancePercent)%)"
            Anchor = 'devices'
            Priority = 1
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Convert-ToOwnershipAssessmentDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value }

    try {
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text) -or $text -eq 'NotCollected (minimum mode)' -or $text -eq 'N/A') {
            return $null
        }
        return [datetime]$text
    }
    catch {
        return $null
    }
}

function Get-OneDriveDefaultOwnerFromUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Url,
        [Parameter(Mandatory = $false)]
        [hashtable]$SegmentToUpnMap,
        [Parameter(Mandatory = $false)]
        [string[]]$KnownDomains
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $null
    }

    try {
        $pathSegment = ($Url.TrimEnd('/') -split '/')[-1]
    }
    catch {
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($pathSegment)) {
        return $null
    }

    $normalizedSegment = $pathSegment.Trim().ToLowerInvariant()

    if ($SegmentToUpnMap -and $SegmentToUpnMap.ContainsKey($normalizedSegment)) {
        $mappedOwner = [string]$SegmentToUpnMap[$normalizedSegment]
        if (-not [string]::IsNullOrWhiteSpace($mappedOwner)) {
            return $mappedOwner.ToLowerInvariant()
        }
    }

    if ($KnownDomains -and $KnownDomains.Count -gt 0) {
        $orderedDomains = @(
            $KnownDomains |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { $_.Trim().ToLowerInvariant() } |
                Select-Object -Unique |
                Sort-Object Length -Descending
        )

        foreach ($domain in $orderedDomains) {
            $encodedDomain = ($domain -replace '\.', '_')
            if ([string]::IsNullOrWhiteSpace($encodedDomain)) { continue }

            $suffix = "_$encodedDomain"
            if (-not $normalizedSegment.EndsWith($suffix)) { continue }

            $encodedLocal = $normalizedSegment.Substring(0, $normalizedSegment.Length - $suffix.Length)
            if ([string]::IsNullOrWhiteSpace($encodedLocal)) { continue }

            $localPart = ($encodedLocal -replace '_', '.')
            return ("{0}@{1}" -f $localPart, $domain).ToLowerInvariant()
        }
    }

    $firstSeparatorIndex = $pathSegment.IndexOf('_')
    if ($firstSeparatorIndex -lt 1) {
        return $pathSegment.ToLowerInvariant()
    }

    $localPart = $pathSegment.Substring(0, $firstSeparatorIndex)
    $domainPart = $pathSegment.Substring($firstSeparatorIndex + 1) -replace '_', '.'
    return ("{0}@{1}" -f $localPart, $domainPart).ToLowerInvariant()
}

function Update-OwnershipGovernanceTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    if (-not $TenantStatsHash) {
        return
    }

    $context = Get-TenantAssessmentContext -TenantStatsHash $TenantStatsHash
    $TenantStatsHash['UnmanagedObjects'] = @{}
    $TenantStatsHash['OneDriveOwnerMismatches'] = @{}
    $TenantStatsHash['OwnershipGovernanceSummary'] = @{}

    $staleOwnerDays = 180
    $staleCutoff = (Get-Date).AddDays(-1 * $staleOwnerDays)
    $ownersByUpn = @{}

    function Add-OwnerProfile {
        param([object]$Record)

        if (-not $Record) { return }
        $upn = [string]$Record.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($upn)) { return }

        $normalizedUpn = $upn.Trim().ToLowerInvariant()
        $accountEnabled = $true
        if ($Record.PSObject.Properties['AccountEnabled']) {
            try { $accountEnabled = [bool]$Record.AccountEnabled } catch { $accountEnabled = $true }
        }
        $lastSignIn = $null
        if ($Record.PSObject.Properties['LastSignInDateTime']) {
            $lastSignIn = Convert-ToOwnershipAssessmentDate -Value $Record.LastSignInDateTime
        }

        if (-not $ownersByUpn.ContainsKey($normalizedUpn)) {
            $ownersByUpn[$normalizedUpn] = [PSCustomObject]@{
                UserPrincipalName = $upn
                AccountEnabled = $accountEnabled
                LastSignInDateTime = $lastSignIn
            }
            return
        }

        $existing = $ownersByUpn[$normalizedUpn]
        if ($accountEnabled -eq $false) {
            $existing | Add-Member -MemberType NoteProperty -Name AccountEnabled -Value $false -Force
        }
        if ($lastSignIn -and (-not $existing.LastSignInDateTime -or $lastSignIn -gt $existing.LastSignInDateTime)) {
            $existing | Add-Member -MemberType NoteProperty -Name LastSignInDateTime -Value $lastSignIn -Force
        }
    }

    foreach ($user in $context.Users) { Add-OwnerProfile -Record $user }
    foreach ($admin in $context.Admins) { Add-OwnerProfile -Record $admin }

    $oneDriveSegmentToUpnMap = @{}
    $knownOwnerDomains = @(
        $ownersByUpn.Keys |
            Where-Object { $_ -and $_ -like '*@*' } |
            ForEach-Object { (($_ -split '@', 2)[1]).Trim().ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )
    foreach ($upn in $ownersByUpn.Keys) {
        if ([string]::IsNullOrWhiteSpace($upn) -or $upn -notlike '*@*') { continue }
        $normalizedUpn = $upn.Trim().ToLowerInvariant()
        $encodedSegment = ($normalizedUpn -replace '@', '_' -replace '\.', '_')
        if ([string]::IsNullOrWhiteSpace($encodedSegment)) { continue }
        if (-not $oneDriveSegmentToUpnMap.ContainsKey($encodedSegment)) {
            $oneDriveSegmentToUpnMap[$encodedSegment] = $normalizedUpn
        }
    }

    function Get-OwnerTokens {
        param($OwnerValue)

        $rawTokens = @()
        if ($null -eq $OwnerValue) {
            return @()
        }

        if ($OwnerValue -is [System.Collections.IEnumerable] -and -not ($OwnerValue -is [string])) {
            foreach ($item in $OwnerValue) {
                if ($null -ne $item) { $rawTokens += [string]$item }
            }
        }
        else {
            $rawTokens += ([string]$OwnerValue -split '[,;]')
        }

        $normalized = New-Object 'System.Collections.Generic.List[string]'
        foreach ($token in $rawTokens) {
            $trimmed = ([string]$token).Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
            $trimmed = ($trimmed -replace '(?i)^smtp:', '').Trim()
            $emailMatch = [regex]::Match($trimmed, '(?i)[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}')
            if ($emailMatch.Success) {
                $trimmed = $emailMatch.Value
            }
            $normalized.Add($trimmed.ToLowerInvariant())
        }

        return @($normalized | Select-Object -Unique)
    }

    function Get-OwnerState {
        param($OwnerValue)

        $owners = @(Get-OwnerTokens -OwnerValue $OwnerValue)
        if ($owners.Count -eq 0) {
            return [PSCustomObject]@{
                Owners = @()
                PrimaryOwner = $null
                State = 'Missing'
                Reason = 'No owner assigned'
                ResolvedOwnerCount = 0
                HealthyOwnerCount = 0
                DisabledOwnerCount = 0
                StaleOwnerCount = 0
                UnknownOwnerCount = 0
            }
        }

        $resolved = 0
        $healthy = 0
        $disabled = 0
        $stale = 0
        $unknown = 0

        foreach ($owner in $owners) {
            if (-not $ownersByUpn.ContainsKey($owner)) {
                $unknown++
                continue
            }

            $resolved++
            $ownerProfile = $ownersByUpn[$owner]
            if ($ownerProfile.AccountEnabled -eq $false) {
                $disabled++
                continue
            }

            $lastSignIn = $ownerProfile.LastSignInDateTime
            if ($lastSignIn -and $lastSignIn -lt $staleCutoff) {
                $stale++
            }
            else {
                $healthy++
            }
        }

        $state = 'Healthy'
        $reason = 'Owner account is active'

        if ($resolved -eq 0) {
            $state = 'Unknown'
            $reason = 'Owner account could not be resolved to collected user data'
        }
        elseif ($disabled -gt 0 -and $healthy -eq 0 -and $stale -eq 0) {
            $state = 'Disabled'
            $reason = 'Owner account is disabled'
        }
        elseif ($stale -gt 0 -and $healthy -eq 0 -and $disabled -eq 0) {
            $state = 'Stale'
            $reason = "Owner account has no sign-in within $staleOwnerDays days"
        }
        elseif (($disabled + $stale) -gt 0) {
            $state = 'Mixed'
            $reason = "At least one resolved owner account is disabled or stale ($staleOwnerDays+ days)"
        }

        return [PSCustomObject]@{
            Owners = $owners
            PrimaryOwner = if ($owners.Count -gt 0) { $owners[0] } else { $null }
            State = $state
            Reason = $reason
            ResolvedOwnerCount = $resolved
            HealthyOwnerCount = $healthy
            DisabledOwnerCount = $disabled
            StaleOwnerCount = $stale
            UnknownOwnerCount = $unknown
        }
    }

    function Convert-ToNullableInt {
        param($Value)
        if ($null -eq $Value) { return $null }
        if ($Value -is [int]) { return [int]$Value }
        if ($Value -is [long]) { return [int]$Value }

        try {
            $text = [string]$Value
            if ([string]::IsNullOrWhiteSpace($text)) { return $null }
            $trimmed = $text.Trim()
            if ($trimmed -match '^\d+$') {
                return [int]$trimmed
            }
        }
        catch {
            return $null
        }

        return $null
    }

    $teamsNameSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($team in $context.Teams) {
        $teamName = [string]$team.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($teamName)) {
            $null = $teamsNameSet.Add($teamName.Trim())
        }
    }

    $unmanagedRows = New-Object System.Collections.Generic.List[object]
    $mismatchRows = New-Object System.Collections.Generic.List[object]
    $unknownOwnerStateCount = 0

    function Add-UnmanagedRow {
        param(
            [string]$Workload,
            [string]$ObjectType,
            [string]$DisplayName,
            [string]$Identifier,
            [string]$CurrentOwner,
            [string]$OwnerState,
            [string]$Reason,
            [string]$Severity,
            [string]$SourceWorksheet
        )

        $unmanagedRows.Add([PSCustomObject]@{
            Workload = $Workload
            ObjectType = $ObjectType
            DisplayName = $DisplayName
            Identifier = $Identifier
            CurrentOwner = $CurrentOwner
            OwnerState = $OwnerState
            Reason = $Reason
            Severity = $Severity
            SourceWorksheet = $SourceWorksheet
        }) | Out-Null
    }

    foreach ($oneDrive in $context.OneDrive) {
        $displayName = if ($oneDrive.Title) { [string]$oneDrive.Title } else { [string]$oneDrive.Url }
        $identifier = if ($oneDrive.Url) { [string]$oneDrive.Url } elseif ($oneDrive.SiteId) { [string]$oneDrive.SiteId } else { $displayName }
        $ownerState = Get-OwnerState -OwnerValue $oneDrive.Owner
        if ($ownerState.State -eq 'Unknown') { $unknownOwnerStateCount++ }

        $expectedOwner = Get-OneDriveDefaultOwnerFromUrl -Url $oneDrive.Url -SegmentToUpnMap $oneDriveSegmentToUpnMap -KnownDomains $knownOwnerDomains
        if ($expectedOwner -and $ownerState.PrimaryOwner -and $expectedOwner -ne $ownerState.PrimaryOwner) {
            $mismatchRows.Add([PSCustomObject]@{
                Workload = 'OneDrive'
                DisplayName = $displayName
                SiteUrl = [string]$oneDrive.Url
                SiteId = [string]$oneDrive.SiteId
                CurrentOwner = [string]$ownerState.PrimaryOwner
                ExpectedDefaultOwner = [string]$expectedOwner
                Assessment = 'Review'
                SourceWorksheet = 'OneDrive'
            }) | Out-Null
        }

        if ($ownerState.State -eq 'Missing') {
            Add-UnmanagedRow -Workload 'OneDrive' -ObjectType 'OneDrive Site' -DisplayName $displayName -Identifier $identifier -CurrentOwner '' -OwnerState $ownerState.State -Reason $ownerState.Reason -Severity 'Risk' -SourceWorksheet 'OneDrive'
        }
        elseif ($ownerState.State -in @('Disabled', 'Stale', 'Mixed')) {
            Add-UnmanagedRow -Workload 'OneDrive' -ObjectType 'OneDrive Site' -DisplayName $displayName -Identifier $identifier -CurrentOwner ([string]$oneDrive.Owner) -OwnerState $ownerState.State -Reason $ownerState.Reason -Severity 'Warning' -SourceWorksheet 'OneDrive'
        }
    }

    foreach ($site in $context.SharePoint) {
        $isTeamsSite = ($site.Template -eq 'TEAMCHANNEL#0' -or $site.IsTeamsChannelConnected -eq $true)
        $workload = if ($isTeamsSite) { 'Teams' } else { 'SharePoint' }
        $objectType = if ($isTeamsSite) { 'Teams Site' } else { 'SharePoint Site' }
        $displayName = if ($site.Title) { [string]$site.Title } else { [string]$site.Url }
        $identifier = if ($site.Url) { [string]$site.Url } elseif ($site.SiteId) { [string]$site.SiteId } else { $displayName }
        $ownerState = Get-OwnerState -OwnerValue $site.Owner
        if ($ownerState.State -eq 'Unknown') { $unknownOwnerStateCount++ }

        if ($ownerState.State -eq 'Missing') {
            Add-UnmanagedRow -Workload $workload -ObjectType $objectType -DisplayName $displayName -Identifier $identifier -CurrentOwner '' -OwnerState $ownerState.State -Reason $ownerState.Reason -Severity 'Risk' -SourceWorksheet 'SharePoint'
        }
        elseif ($ownerState.State -in @('Disabled', 'Stale', 'Mixed')) {
            Add-UnmanagedRow -Workload $workload -ObjectType $objectType -DisplayName $displayName -Identifier $identifier -CurrentOwner ([string]$site.Owner) -OwnerState $ownerState.State -Reason $ownerState.Reason -Severity 'Warning' -SourceWorksheet 'SharePoint'
        }

        if ($isTeamsSite -and -not [string]::IsNullOrWhiteSpace($displayName)) {
            $null = $teamsNameSet.Add($displayName.Trim())
        }
    }

    foreach ($group in $context.Groups) {
        $ownerCount = Convert-ToNullableInt -Value $group.OwnerCount
        if ($null -eq $ownerCount) {
            $unknownOwnerStateCount++
            continue
        }

        if ($ownerCount -eq 0) {
            $groupDisplayName = [string]$group.DisplayName
            $isTeamGroup = (-not [string]::IsNullOrWhiteSpace($groupDisplayName) -and $teamsNameSet.Contains($groupDisplayName.Trim()))
            Add-UnmanagedRow `
                -Workload $(if ($isTeamGroup) { 'Teams' } else { 'Entra ID' }) `
                -ObjectType $(if ($isTeamGroup) { 'Team (M365 Group)' } else { 'Entra Group' }) `
                -DisplayName $groupDisplayName `
                -Identifier ([string]$group.ID) `
                -CurrentOwner 'None (OwnerCount=0)' `
                -OwnerState 'Missing' `
                -Reason 'No owner assigned (OwnerCount=0)' `
                -Severity 'Risk' `
                -SourceWorksheet 'EntraIDGroups'
        }
    }

    foreach ($group in $context.ExchangeGroups) {
        $ownerCount = Convert-ToNullableInt -Value $group.OwnersCount
        if ($null -eq $ownerCount) {
            $unknownOwnerStateCount++
            continue
        }

        if ($ownerCount -eq 0) {
            $groupDisplayName = [string]$group.DisplayName
            $identifier = if ($group.PrimarySMTPAddress) { [string]$group.PrimarySMTPAddress } else { [string]$group.Identity }
            Add-UnmanagedRow `
                -Workload 'Exchange' `
                -ObjectType $(if ($group.RecipientTypeDetails) { [string]$group.RecipientTypeDetails } else { 'Exchange Group' }) `
                -DisplayName $groupDisplayName `
                -Identifier $identifier `
                -CurrentOwner 'None (OwnersCount=0)' `
                -OwnerState 'Missing' `
                -Reason 'No owner assigned (OwnersCount=0)' `
                -Severity 'Risk' `
                -SourceWorksheet 'AllExchangeGroups'
        }
    }

    $sortedUnmanagedRows = @(
        $unmanagedRows |
            Sort-Object @{ Expression = {
                switch ($_.Severity) {
                    'Risk' { 1 }
                    'Warning' { 2 }
                    default { 3 }
                }
            } }, Workload, ObjectType, DisplayName
    )
    $sortedMismatchRows = @($mismatchRows | Sort-Object DisplayName, SiteUrl)

    $unmanagedIndex = 0
    foreach ($row in $sortedUnmanagedRows) {
        $unmanagedIndex++
        $key = "{0:D4}-{1}" -f $unmanagedIndex, ($row.Identifier -replace '[^a-zA-Z0-9@._-]', '_')
        $TenantStatsHash['UnmanagedObjects'][$key] = $row
    }

    $mismatchIndex = 0
    foreach ($row in $sortedMismatchRows) {
        $mismatchIndex++
        $key = "{0:D4}-{1}" -f $mismatchIndex, (($row.SiteUrl) -replace '[^a-zA-Z0-9@._/-]', '_')
        $TenantStatsHash['OneDriveOwnerMismatches'][$key] = $row
    }

    $missingOwnerCount = @($sortedUnmanagedRows | Where-Object { $_.OwnerState -eq 'Missing' }).Count
    $ownerHealthRiskCount = @($sortedUnmanagedRows | Where-Object { $_.OwnerState -in @('Disabled', 'Stale', 'Mixed') }).Count

    $summary = [PSCustomObject]@{
        TotalObjectsReviewed = @($context.OneDrive).Count + @($context.SharePoint).Count + @($context.Groups).Count + @($context.ExchangeGroups).Count
        UnmanagedObjectCount = $sortedUnmanagedRows.Count
        MissingOwnerCount = $missingOwnerCount
        OwnerHealthRiskCount = $ownerHealthRiskCount
        OneDriveOwnerMismatchCount = $sortedMismatchRows.Count
        UnknownOwnerStateCount = $unknownOwnerStateCount
        OneDriveSiteCount = @($context.OneDrive).Count
        SharePointSiteCount = @($context.SharePoint).Count
        TeamsObjectCount = @($context.SharePoint | Where-Object { $_.Template -eq 'TEAMCHANNEL#0' -or $_.IsTeamsChannelConnected -eq $true }).Count
        EntraGroupCount = @($context.Groups).Count
        ExchangeGroupCount = @($context.ExchangeGroups).Count
        StaleOwnerThresholdDays = $staleOwnerDays
    }

    $TenantStatsHash['OwnershipGovernanceSummary']['Summary'] = $summary
}

function Get-OwnershipGovernanceAnalysis {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [array]$UnmanagedObjects,
        [Parameter(Mandatory = $false)]
        [array]$OneDriveOwnerMismatches,
        [Parameter(Mandatory = $false)]
        [object]$OwnershipGovernanceSummary
    )

    $findings = @()
    $unmanagedCount = @($UnmanagedObjects).Count
    $mismatchCount = @($OneDriveOwnerMismatches).Count

    $missingOwnerCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['MissingOwnerCount']) {
        [int]$OwnershipGovernanceSummary.MissingOwnerCount
    } else {
        @($UnmanagedObjects | Where-Object { $_.OwnerState -eq 'Missing' }).Count
    }

    $ownerHealthRiskCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['OwnerHealthRiskCount']) {
        [int]$OwnershipGovernanceSummary.OwnerHealthRiskCount
    } else {
        @($UnmanagedObjects | Where-Object { $_.OwnerState -in @('Disabled', 'Stale', 'Mixed') }).Count
    }

    $unknownOwnerStateCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['UnknownOwnerStateCount']) {
        [int]$OwnershipGovernanceSummary.UnknownOwnerStateCount
    } else {
        0
    }

    if ($missingOwnerCount -gt 0) {
        $findings += @{
            Type = 'Risk'
            Category = 'Unowned Objects'
            Message = "$missingOwnerCount object(s) are missing owners and require ownership assignment"
            Anchor = 'ownership-governance'
            Priority = 1
        }
    }

    if ($ownerHealthRiskCount -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Owner Health'
            Message = "$ownerHealthRiskCount object(s) are owned by disabled or stale owner accounts"
            Anchor = 'ownership-governance'
            Priority = 2
        }
    }

    if ($mismatchCount -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'OneDrive Ownership Mismatch'
            Message = "$mismatchCount OneDrive site(s) have a current owner different from the URL-derived default owner"
            Anchor = 'ownership-governance'
            Priority = 2
        }
    }

    if ($unmanagedCount -eq 0 -and $mismatchCount -eq 0 -and $unknownOwnerStateCount -gt 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Owner Telemetry'
            Message = "$unknownOwnerStateCount object(s) had unresolved ownership state in collected data; review deeper collection modes for complete ownership validation"
            Anchor = 'ownership-governance'
            Priority = 3
        }
    }

    return @{
        Findings = $findings
    }
}

function Get-SharePointOneDriveAnalysis {
    <#
    .SYNOPSIS
        Analyzes SharePoint and OneDrive usage
    #>
    param(
        [array]$SharePointSites,
        [array]$OneDriveSites
    )
    
    $findings = @()
    
    # Large SharePoint sites
    $largeSPSites = $SharePointSites | Where-Object {
        $_.StorageUsedGB -gt $script:DefaultThresholds.SharePointSiteGB
    }
    
    if ($largeSPSites.Count -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large SharePoint Sites'
            Message = "$($largeSPSites.Count) SharePoint sites exceed $($script:DefaultThresholds.SharePointSiteGB) GB"
            Anchor = 'sharepoint-onedrive'
            Priority = 2
        }
    }

    # Large Teams / M365 Groups (over 1 TB) - include names
    $largeTeamsSites = $SharePointSites | Where-Object {
        $_.Template -eq 'TEAMCHANNEL#0' -and $_.StorageUsedGB -gt $script:DefaultThresholds.SharePointSiteGB
    }
    foreach ($site in $largeTeamsSites) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large Teams Site'
            Message = "Teams site '$($site.Title)' exceeds $($script:DefaultThresholds.SharePointSiteGB) GB"
            Anchor = 'sharepoint-onedrive'
            Priority = 2
        }
    }

    $largeGroupSites = $SharePointSites | Where-Object {
        $_.IsOffice365GroupsConnected -and $_.Template -ne 'TEAMCHANNEL#0' -and $_.StorageUsedGB -gt $script:DefaultThresholds.SharePointSiteGB
    }
    foreach ($site in $largeGroupSites) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large M365 Group Site'
            Message = "Microsoft 365 group site '$($site.Title)' exceeds $($script:DefaultThresholds.SharePointSiteGB) GB"
            Anchor = 'sharepoint-onedrive'
            Priority = 2
        }
    }
    
    # Large OneDrive sites
    $largeODSites = $OneDriveSites | Where-Object {
        $_.StorageUsedGB -gt $script:DefaultThresholds.OneDriveSiteGB
    }
    
    if ($largeODSites.Count -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large OneDrive Sites'
            Message = "$($largeODSites.Count) OneDrive sites exceed $($script:DefaultThresholds.OneDriveSiteGB) GB"
            Anchor = 'sharepoint-onedrive'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-InactiveMailboxAnalysis {
    <#
    .SYNOPSIS
        Analyzes inactive mailboxes
    #>
    param([array]$InactiveMailboxes)
    
    $findings = @()
    
    if ($InactiveMailboxes.Count -eq 0) {
        return @{ Findings = @() }
    }
    
    # Count of inactive mailboxes
    $findings += @{
        Type = 'Info'
        Category = 'Inactive Mailboxes'
        Message = "$($InactiveMailboxes.Count) inactive mailboxes are consuming storage (not counted in license utilization)"
        Anchor = 'inactive-mailboxes'
        Priority = 3
    }
    
    # Large inactive mailboxes
    $largeInactive = $InactiveMailboxes | Where-Object {
        try {
            $size = [double]$_.MBXSizeGB
            $size -gt 50
        } catch {
            $false
        }
    }
    
    if ($largeInactive.Count -gt 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Large Inactive Mailboxes'
            Message = "$($largeInactive.Count) inactive mailboxes exceed 50 GB - consider if retention is still required"
            Anchor = 'inactive-mailboxes'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-ExchangeHybridAnalysis {
    <#
    .SYNOPSIS
        Analyzes Exchange hybrid configuration and returns findings
    #>
    param([object]$HybridInfo)
    
    $findings = @()
    
    if (-not $HybridInfo) {
        return @{ Findings = @() }
    }
    
    $partialIndicators = (
        ($HybridInfo.InboundOnPremConnectorCount -gt 0) -or
        ($HybridInfo.OutboundOnPremConnectorCount -gt 0) -or
        ($HybridInfo.MigrationEndpointCount -gt 0) -or
        ($HybridInfo.EvidenceCount -gt 0)
    )
    
    if ($HybridInfo.IsHybridConfigured -eq $true) {
        $findings += @{
            Type = 'Info'
            Category = 'Exchange Hybrid'
            Message = "Exchange hybrid detected: $($HybridInfo.HybridType) (evidence: $($HybridInfo.EvidenceCount))"
            Anchor = 'exchange-hybrid'
            Priority = 3
        }
    } elseif ($HybridInfo.HybridStatus -eq 'Possible Hybrid' -or $partialIndicators) {
        $findings += @{
            Type = 'Warning'
            Category = 'Exchange Hybrid'
            Message = "Partial hybrid indicators detected without full configuration (review mail flow connectors and org settings)"
            Anchor = 'exchange-hybrid'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-IdentityAdminAnalysis {
    <#
    .SYNOPSIS
        Analyzes identity and admin data for findings
    #>
    param(
        [array]$Users,
        [array]$Admins,
        [array]$Groups,
        [array]$ConditionalAccessPolicies
    )
    
    $findings = @()
    
    if ($Admins.Count -eq 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Admins'
            Message = "No admin role assignments were found in the report"
            Anchor = 'identity-admins'
            Priority = 2
        }
    }
    
    if ($Groups.Count -eq 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Groups'
            Message = "No Entra ID groups found (or not collected)"
            Anchor = 'identity-admins'
            Priority = 3
        }
    }

    function Convert-ToAssessmentDate {
        param([Parameter(Mandatory = $false)][AllowNull()]$Value)
        if ($null -eq $Value) { return $null }
        if ($Value -is [datetime]) { return $Value }
        try {
            $text = [string]$Value
            if ([string]::IsNullOrWhiteSpace($text) -or $text -eq 'NotCollected (minimum mode)') {
                return $null
            }
            return [datetime]$text
        }
        catch {
            return $null
        }
    }

    $inactiveDaysThreshold = 180
    $inactiveCutoff = (Get-Date).AddDays(-1 * $inactiveDaysThreshold)
    $emergencyAccessRecentSignInThresholdDays = 30
    $emergencyAccessRecentSignInCutoff = (Get-Date).AddDays(-1 * $emergencyAccessRecentSignInThresholdDays)

    $memberUsers = @(
        $Users | Where-Object {
            $null -ne $_ -and
            ($_.UserType -ne 'Guest') -and
            ($_.UserType -ne 'GuestUser') -and
            ($_.UserPrincipalName -notlike '*#EXT#*')
        }
    )
    $enabledMemberUsers = @(
        $memberUsers | Where-Object {
            if ($null -eq $_) { return $false }
            $accountEnabled = $true
            if ($_.PSObject.Properties['AccountEnabled']) {
                try { $accountEnabled = [bool]$_.AccountEnabled } catch { $accountEnabled = $true }
            }
            $accountEnabled
        }
    )

    $usersWithSignin = @()
    foreach ($user in $enabledMemberUsers) {
        $lastSignIn = Convert-ToAssessmentDate -Value $user.LastSignInDateTime
        if ($lastSignIn) {
            $usersWithSignin += [PSCustomObject]@{
                User = $user
                LastSignIn = $lastSignIn
            }
        }
    }

    $inactiveUsers = @($usersWithSignin | Where-Object { $_.LastSignIn -lt $inactiveCutoff })
    if ($inactiveUsers.Count -gt 0) {
        $inactivePct = if ($enabledMemberUsers.Count -gt 0) {
            [math]::Round((($inactiveUsers.Count / $enabledMemberUsers.Count) * 100), 1)
        }
        else { 0 }
        $inactiveType = if ($inactivePct -ge 20) { 'Risk' } else { 'Warning' }
        $inactivePriority = if ($inactivePct -ge 20) { 1 } else { 2 }
        $findings += @{
            Type = $inactiveType
            Category = 'Inactive Users'
            Message = "$($inactiveUsers.Count) enabled member account(s) ($inactivePct%) have no sign-in within the last $inactiveDaysThreshold days"
            Anchor = 'identity-admins'
            Priority = $inactivePriority
        }
    }

    if ($enabledMemberUsers.Count -gt 0) {
        $signInCoveragePct = [math]::Round((($usersWithSignin.Count / $enabledMemberUsers.Count) * 100), 1)
        if ($signInCoveragePct -lt 60) {
            $findings += @{
                Type = 'Info'
                Category = 'Admin Sign-in Telemetry'
                Message = "User sign-in telemetry coverage is $signInCoveragePct% for enabled member users; inactivity findings may be under-reported"
                Anchor = 'identity-admins'
                Priority = 3
            }
        }
    }

    $guestUsers = @(
        $Users | Where-Object {
            ($_.UserType -eq 'Guest') -or
            ($_.UserType -eq 'GuestUser') -or
            ($_.UserPrincipalName -like '*#EXT#*')
        }
    )
    $enabledGuestUsers = @(
        $guestUsers | Where-Object {
            $accountEnabled = $true
            if ($_.PSObject.Properties['AccountEnabled']) {
                try { $accountEnabled = [bool]$_.AccountEnabled } catch { $accountEnabled = $true }
            }
            $accountEnabled
        }
    )
    $inactiveGuests = @(
        $enabledGuestUsers | Where-Object {
            $lastSignIn = Convert-ToAssessmentDate -Value $_.LastSignInDateTime
            $lastSignIn -and $lastSignIn -lt $inactiveCutoff
        }
    )
    if ($inactiveGuests.Count -gt 0) {
        $inactiveGuestPct = if ($enabledGuestUsers.Count -gt 0) {
            [math]::Round((($inactiveGuests.Count / $enabledGuestUsers.Count) * 100), 1)
        } else { 0 }
        $findings += @{
            Type = 'Warning'
            Category = 'Inactive Guest Users'
            Message = "$($inactiveGuests.Count) enabled guest account(s) ($inactiveGuestPct%) have no sign-in within the last $inactiveDaysThreshold days"
            Anchor = 'identity-admins'
            Priority = 2
        }
    }

    $normalizedAdmins = @(
        $Admins | Where-Object { $_ -ne $null }
    )

    $userByUpn = @{}
    foreach ($user in $Users) {
        $userUpn = [string]$user.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($userUpn)) { continue }
        $userByUpn[$userUpn.ToLowerInvariant()] = $user
    }

    if ($normalizedAdmins.Count -gt 0) {
        $globalAdmins = @(
            $normalizedAdmins | Where-Object {
                $roleText = [string]$_.Role
                $roleText -match 'Global Administrator|Company Administrator'
            }
        )
        if ($globalAdmins.Count -gt 5) {
            $findings += @{
                Type = 'Risk'
                Category = 'Global Admin Count'
                Message = "$($globalAdmins.Count) Global Administrator account(s) detected; Microsoft least-privilege guidance recommends reducing standing Global Admins"
                Anchor = 'identity-admins'
                Priority = 1
            }
        }
        elseif ($globalAdmins.Count -gt 4) {
            $findings += @{
                Type = 'Warning'
                Category = 'Global Admin Count'
                Message = "$($globalAdmins.Count) Global Administrator account(s) detected; review whether all are required as standing access"
                Anchor = 'identity-admins'
                Priority = 2
            }
        }

        $staleAdmins = @(
            $normalizedAdmins | Where-Object {
                $lastSignIn = Convert-ToAssessmentDate -Value $_.LastSignInDateTime
                $lastSignIn -and $lastSignIn -lt $inactiveCutoff
            }
        )
        if ($staleAdmins.Count -gt 0) {
            $findings += @{
                Type = 'Risk'
                Category = 'Inactive Admin Accounts'
                Message = "$($staleAdmins.Count) admin account(s) have not signed in within $inactiveDaysThreshold days and should be reviewed or removed from privileged roles"
                Anchor = 'identity-admins'
                Priority = 1
            }
        }

        $globalAdminUsers = @()
        $seenGlobalAdminUpns = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($admin in $globalAdmins) {
            $adminUpn = [string]$admin.UserPrincipalName
            if ([string]::IsNullOrWhiteSpace($adminUpn)) { continue }
            $normalizedUpn = $adminUpn.ToLowerInvariant()
            if ($seenGlobalAdminUpns.Contains($normalizedUpn)) { continue }
            [void]$seenGlobalAdminUpns.Add($normalizedUpn)

            $matchedUser = $null
            if ($userByUpn.ContainsKey($normalizedUpn)) {
                $matchedUser = $userByUpn[$normalizedUpn]
            }

            $accountEnabled = $true
            if ($matchedUser -and $matchedUser.PSObject.Properties['AccountEnabled']) {
                try { $accountEnabled = [bool]$matchedUser.AccountEnabled } catch { $accountEnabled = $true }
            }
            elseif ($admin.PSObject.Properties['AccountEnabled']) {
                try { $accountEnabled = [bool]$admin.AccountEnabled } catch { $accountEnabled = $true }
            }

            $lastSignIn = $null
            if ($matchedUser) {
                $lastSignIn = Convert-ToAssessmentDate -Value $matchedUser.LastSignInDateTime
            }
            if (-not $lastSignIn) {
                $lastSignIn = Convert-ToAssessmentDate -Value $admin.LastSignInDateTime
            }

            $isCloudOnly = $true
            if ($matchedUser -and $matchedUser.PSObject.Properties['OnPremisesSyncEnabled'] -and $matchedUser.OnPremisesSyncEnabled -eq $true) {
                $isCloudOnly = $false
            }

            $objectId = $null
            if ($matchedUser -and $matchedUser.PSObject.Properties['Id'] -and $matchedUser.Id) {
                $objectId = [string]$matchedUser.Id
            }

            $globalAdminUsers += [PSCustomObject]@{
                UserPrincipalName = $adminUpn
                AccountEnabled = $accountEnabled
                LastSignInDateTime = $lastSignIn
                IsCloudOnly = $isCloudOnly
                ObjectId = $objectId
            }
        }

        $enabledCloudOnlyGlobalAdmins = @(
            $globalAdminUsers | Where-Object { $_.AccountEnabled -eq $true -and $_.IsCloudOnly -eq $true }
        )
        $emergencyAccessCandidates = @(
            $enabledCloudOnlyGlobalAdmins | Where-Object {
                (-not $_.LastSignInDateTime) -or ($_.LastSignInDateTime -lt $emergencyAccessRecentSignInCutoff)
            }
        )

        if ($enabledCloudOnlyGlobalAdmins.Count -eq 0) {
            $findings += @{
                Type = 'Warning'
                Category = 'Emergency Access Accounts'
                Message = "No enabled cloud-only Global Administrator accounts were detected; maintain dedicated emergency access accounts per Microsoft guidance"
                Anchor = 'identity-admins'
                Priority = 2
            }
        }
        elseif ($emergencyAccessCandidates.Count -lt 2) {
            $findings += @{
                Type = 'Warning'
                Category = 'Emergency Access Accounts'
                Message = "Only $($emergencyAccessCandidates.Count) cloud-only Global Administrator account(s) appear to fit emergency-access profile (enabled and no recent sign-in > $emergencyAccessRecentSignInThresholdDays days); Microsoft recommends at least two"
                Anchor = 'identity-admins'
                Priority = 2
            }
        }

        $enabledSyncedGlobalAdmins = @(
            $globalAdminUsers | Where-Object { $_.AccountEnabled -eq $true -and $_.IsCloudOnly -eq $false }
        )
        if ($enabledSyncedGlobalAdmins.Count -gt 0) {
            $findings += @{
                Type = 'Info'
                Category = 'Emergency Access Accounts'
                Message = "$($enabledSyncedGlobalAdmins.Count) enabled Global Administrator account(s) are synchronized from on-premises; emergency access accounts should be cloud-only"
                Anchor = 'identity-admins'
                Priority = 3
            }
        }

        if ($emergencyAccessCandidates.Count -gt 0 -and $ConditionalAccessPolicies) {
            $enabledPolicies = @($ConditionalAccessPolicies | Where-Object { $_.State -eq 'enabled' })
            $excludedUserValues = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($policy in $enabledPolicies) {
                $excludedUsersText = [string]$policy.ExcludedUsers
                if ([string]::IsNullOrWhiteSpace($excludedUsersText)) { continue }
                foreach ($token in ($excludedUsersText -split ',')) {
                    $trimmedToken = $token.Trim().ToLowerInvariant()
                    if (-not [string]::IsNullOrWhiteSpace($trimmedToken)) {
                        [void]$excludedUserValues.Add($trimmedToken)
                    }
                }
            }

            $excludedCandidateCount = 0
            foreach ($candidate in $emergencyAccessCandidates) {
                $candidateMatched = $false
                if ($candidate.ObjectId -and $excludedUserValues.Contains($candidate.ObjectId.ToLowerInvariant())) {
                    $candidateMatched = $true
                }
                elseif ($candidate.UserPrincipalName -and $excludedUserValues.Contains($candidate.UserPrincipalName.ToLowerInvariant())) {
                    $candidateMatched = $true
                }
                if ($candidateMatched) { $excludedCandidateCount++ }
            }

            if ($excludedCandidateCount -lt 1) {
                $findings += @{
                    Type = 'Warning'
                    Category = 'Emergency Access CA Exclusions'
                    Message = "No emergency-access candidate appears in enabled Conditional Access exclusion lists; validate lockout-safe emergency account design"
                    Anchor = 'identity-admins'
                    Priority = 2
                }
            }
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-ConditionalAccessMfaAnalysis {
    <#
    .SYNOPSIS
        Analyzes conditional access and MFA configuration
    #>
    param(
        [array]$ConditionalAccessPolicies,
        [object]$AuthConfig,
        [int]$TotalUsers
    )
    
    $findings = @()
    $enabledPolicies = $ConditionalAccessPolicies | Where-Object { $_.State -eq 'enabled' }
    $mfaPolicies = $enabledPolicies | Where-Object { $_.GrantControls_BuiltInControls -match '(?i)mfa' }
    $mfaEnabled = ($mfaPolicies.Count -gt 0) -or ($AuthConfig -and $AuthConfig.MFAEnabled -eq $true)
    
    if ($enabledPolicies.Count -eq 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Conditional Access'
            Message = "No enabled Conditional Access policies detected"
            Anchor = 'conditional-access-mfa'
            Priority = 2
        }
    }
    
    if (-not $mfaEnabled) {
        $findings += @{
            Type = 'Warning'
            Category = 'MFA Enforcement'
            Message = "No active MFA enforcement detected (no enabled MFA CA policies)"
            Anchor = 'conditional-access-mfa'
            Priority = 2
        }
    }

    $legacyAuthBlockPolicies = @(
        $enabledPolicies | Where-Object {
            ([string]$_.ClientAppTypes -match '(?i)exchangeActiveSync|other') -and
            ([string]$_.GrantControls_BuiltInControls -match '(?i)block')
        }
    )
    if ($enabledPolicies.Count -gt 0 -and $legacyAuthBlockPolicies.Count -eq 0) {
        $findings += @{
            Type = 'Warning'
            Category = 'Legacy Authentication'
            Message = "No enabled Conditional Access policy appears to explicitly block legacy authentication client app types"
            Anchor = 'conditional-access-mfa'
            Priority = 2
        }
    }

    $riskPolicies = @(
        $enabledPolicies | Where-Object {
            (-not [string]::IsNullOrWhiteSpace([string]$_.SignInRiskLevels_IncludeLevels)) -or
            (-not [string]::IsNullOrWhiteSpace([string]$_.ServicePrincipalRiskLevels_IncludeLevels))
        }
    )
    if ($enabledPolicies.Count -gt 0 -and $riskPolicies.Count -eq 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Risk-based Conditional Access'
            Message = "No enabled risk-based Conditional Access policies detected (sign-in risk / user risk)"
            Anchor = 'conditional-access-mfa'
            Priority = 3
        }
    }

    if ($AuthConfig) {
        $adminConsentWorkflowEnabled = $null
        if ($AuthConfig.PSObject.Properties['AdminConsentWorkflowEnabled']) {
            $rawWorkflow = $AuthConfig.AdminConsentWorkflowEnabled
            if ($rawWorkflow -is [bool]) {
                $adminConsentWorkflowEnabled = [bool]$rawWorkflow
            } elseif ($null -ne $rawWorkflow) {
                $workflowText = [string]$rawWorkflow
                if ($workflowText -match '(?i)^(enabled|true|yes)$') { $adminConsentWorkflowEnabled = $true }
                elseif ($workflowText -match '(?i)^(disabled|false|no)$') { $adminConsentWorkflowEnabled = $false }
            }
        }
        if ($adminConsentWorkflowEnabled -eq $false) {
            $findings += @{
                Type = 'Warning'
                Category = 'Admin Consent Workflow'
                Message = "Admin consent request workflow is disabled; enable it to govern end-user app consent escalation"
                Anchor = 'conditional-access-mfa'
                Priority = 2
            }
        }

        $defaultUserCanCreateApps = $null
        if ($AuthConfig.PSObject.Properties['DefaultUserCanCreateApps']) {
            $rawCreateApps = $AuthConfig.DefaultUserCanCreateApps
            if ($rawCreateApps -is [bool]) {
                $defaultUserCanCreateApps = [bool]$rawCreateApps
            } elseif ($null -ne $rawCreateApps) {
                $createAppsText = [string]$rawCreateApps
                if ($createAppsText -match '(?i)^(yes|true|enabled)$') { $defaultUserCanCreateApps = $true }
                elseif ($createAppsText -match '(?i)^(no|false|disabled)$') { $defaultUserCanCreateApps = $false }
            }
        }
        if ($defaultUserCanCreateApps -eq $true) {
            $findings += @{
                Type = 'Info'
                Category = 'App Consent Governance'
                Message = "Default users are allowed to create app registrations; verify enterprise governance requirements for app creation"
                Anchor = 'conditional-access-mfa'
                Priority = 3
            }
        }

        $permissionGrantPolicyText = $null
        if ($AuthConfig.PSObject.Properties['PermissionGrantPoliciesAssigned']) {
            $permissionGrantPolicyText = (@($AuthConfig.PermissionGrantPoliciesAssigned) -join ',')
        } elseif ($AuthConfig.PSObject.Properties['PermissionGrantPolicies']) {
            $permissionGrantPolicyText = [string]$AuthConfig.PermissionGrantPolicies
        }
        if (-not [string]::IsNullOrWhiteSpace($permissionGrantPolicyText) -and $permissionGrantPolicyText -match '(?i)legacy') {
            $findings += @{
                Type = 'Warning'
                Category = 'App Consent Governance'
                Message = "Permission grant policy assignment includes legacy/default consent behavior; review least-privilege user consent posture"
                Anchor = 'conditional-access-mfa'
                Priority = 2
            }
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-AdConnectAnalysis {
    <#
    .SYNOPSIS
        Analyzes AD Connect sync data
    #>
    param([object]$AdConnect)
    
    $findings = @()
    if ($AdConnect -and $AdConnect.Summary) {
        if ($AdConnect.Summary.OnPremisesSyncEnabled -ne $true) {
            $findings += @{
                Type = 'Info'
                Category = 'AD Connect'
                Message = "Directory sync is not enabled for this tenant"
                Anchor = 'ad-connect'
                Priority = 3
            }
        }
        if ($AdConnect.ErrorCount -gt 0) {
            $findings += @{
                Type = 'Warning'
                Category = 'AD Connect'
                Message = "$($AdConnect.ErrorCount) recent AD Connect sync errors detected"
                Anchor = 'ad-connect'
                Priority = 2
            }
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Get-FederationAnalysis {
    <#
    .SYNOPSIS
        Analyzes federation and cross-tenant settings
    #>
    param(
        [object]$ExchangeFederation,
        [object]$CrossTenantAccess,
        [object]$ExternalIdentities
    )
    
    $findings = @()
    
    if ($ExchangeFederation -and ($ExchangeFederation.OrganizationRelationshipCount -gt 0 -or $ExchangeFederation.IntraOrgConnectorCount -gt 0)) {
        $findings += @{
            Type = 'Info'
            Category = 'Federation/Relationship'
            Message = "Exchange federation relationships detected (OrgRel: $($ExchangeFederation.OrganizationRelationshipCount), IOC: $($ExchangeFederation.IntraOrgConnectorCount))"
            Anchor = 'exchange-hybrid'
            Priority = 3
        }
    }
    
    if ($CrossTenantAccess -and $CrossTenantAccess.PartnerCount -gt 0) {
        $findings += @{
            Type = 'Info'
            Category = 'Cross-Tenant Access'
            Message = "Cross-tenant access partners configured: $($CrossTenantAccess.PartnerCount)"
            Anchor = 'cross-tenant-access'
            Priority = 3
        }
    }
    
    if ($ExternalIdentities -and $ExternalIdentities.B2BManagementPolicyPresent -eq $false) {
        $findings += @{
            Type = 'Warning'
            Category = 'External Identities'
            Message = "B2B management policy not found; external identity settings may be default/unconfigured"
            Anchor = 'cross-tenant-access'
            Priority = 2
        }
    }
    
    return @{
        Findings = $findings
    }
}

function Convert-MailboxSizeToGB {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $SizeValue
    )

    if (Test-AssessmentHtmlBlankValue -Value $SizeValue) {
        return 0
    }

    try {
        if ($SizeValue -is [ValueType] -and -not ($SizeValue -is [bool]) -and -not ($SizeValue -is [datetime])) {
            $numericValue = [double]$SizeValue
            if ($numericValue -gt 1GB) {
                return [math]::Round(($numericValue / 1GB), 3)
            }
            return [math]::Round($numericValue, 3)
        }
    } catch {}

    try {
        if ($SizeValue.PSObject -and $SizeValue.PSObject.Methods['ToBytes']) {
            return [math]::Round(($SizeValue.ToBytes() / 1GB), 3)
        }
    } catch {}

    $sizeText = [string]$SizeValue
    if ([string]::IsNullOrWhiteSpace($sizeText)) {
        return 0
    }

    $bytesMatch = [regex]::Match($sizeText, '\((?<bytes>[\d,\.]+)\s*bytes\)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($bytesMatch.Success) {
        $bytesText = ($bytesMatch.Groups['bytes'].Value -replace ',', '')
        $bytes = 0.0
        if ([double]::TryParse($bytesText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$bytes)) {
            return [math]::Round(($bytes / 1GB), 3)
        }
    }

    $unitMatch = [regex]::Match($sizeText, '(?<value>[\d,\.]+)\s*(?<unit>TB|GB|MB|KB|B)\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($unitMatch.Success) {
        $valueText = ($unitMatch.Groups['value'].Value -replace ',', '')
        $value = 0.0
        if ([double]::TryParse($valueText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value)) {
            switch ($unitMatch.Groups['unit'].Value.ToUpperInvariant()) {
                'TB' { return [math]::Round(($value * 1024), 3) }
                'GB' { return [math]::Round($value, 3) }
                'MB' { return [math]::Round(($value / 1024), 3) }
                'KB' { return [math]::Round(($value / 1MB), 3) }
                'B'  { return [math]::Round(($value / 1GB), 3) }
            }
        }
    }

    $fallbackText = ($sizeText -replace ',', '')
    $fallback = 0.0
    if ([double]::TryParse($fallbackText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$fallback)) {
        if ($fallback -gt 1GB) {
            return [math]::Round(($fallback / 1GB), 3)
        }
        return [math]::Round($fallback, 3)
    }

    return 0
}

function Get-RecordValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Record,
        [Parameter(Mandatory)]
        [string]$Key
    )

    if ($null -eq $Record) {
        return $null
    }

    if ($Record -is [System.Collections.IDictionary] -and $Record.Contains($Key)) {
        return $Record[$Key]
    }
    if ($Record -is [System.Collections.Specialized.OrderedDictionary] -and $Record.Contains($Key)) {
        return $Record[$Key]
    }
    if ($Record.PSObject -and $Record.PSObject.Properties[$Key]) {
        return $Record.$Key
    }

    return $null
}

function Set-RecordValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Record,
        [Parameter(Mandatory)]
        [string]$Key,
        [AllowNull()]
        $Value
    )

    if ($Record -is [System.Collections.IDictionary]) {
        $Record[$Key] = $Value
        return
    }
    if ($Record -is [System.Collections.Specialized.OrderedDictionary]) {
        $Record[$Key] = $Value
        return
    }

    $Record | Add-Member -MemberType NoteProperty -Name $Key -Value $Value -Force
}

function Get-MailboxStatForRecord {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $MailboxRecord,
        [AllowNull()]
        $StatsHash,
        [switch]$Archive
    )

    if ($null -eq $MailboxRecord -or $null -eq $StatsHash) {
        return $null
    }

    function Convert-ToMailboxLookupKey {
        param([AllowNull()]$Value)

        if ($null -eq $Value) {
            return $null
        }

        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        $text = $text.Trim('{', '}')
        $guidValue = [guid]::Empty
        if ([guid]::TryParse($text, [ref]$guidValue)) {
            return $guidValue.ToString('D').ToLowerInvariant()
        }

        return $text.ToLowerInvariant()
    }

    function Get-ContainerValueByKey {
        param(
            [AllowNull()]$Container,
            [string]$Key
        )

        if ($null -eq $Container -or [string]::IsNullOrWhiteSpace($Key)) {
            return $null
        }

        if ($Container -is [System.Collections.IDictionary] -and $Container.Contains($Key)) {
            return $Container[$Key]
        }

        if ($Container.PSObject -and $Container.PSObject.Properties[$Key]) {
            return $Container.PSObject.Properties[$Key].Value
        }

        return $null
    }

    function Get-MailboxLookupCandidates {
        param(
            [AllowNull()]$Record,
            [switch]$UseArchiveGuid
        )

        if ($null -eq $Record) {
            return @()
        }

        $propertyOrder = @(
            'MailboxStatsLookupKey',
            'ExchangeGuid',
            'MailboxGuid',
            'Guid',
            'ExternalDirectoryObjectId',
            'Id',
            'Identity',
            'PrimarySmtpAddress',
            'WindowsEmailAddress',
            'UserPrincipalName'
        )
        if ($UseArchiveGuid) {
            $propertyOrder = @('ArchiveGuid') + $propertyOrder
        }

        $candidateSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($propertyName in $propertyOrder) {
            $rawValue = Get-RecordValue -Record $Record -Key $propertyName
            if ($null -eq $rawValue -or [string]::IsNullOrWhiteSpace([string]$rawValue)) {
                continue
            }

            $rawText = ([string]$rawValue).Trim()
            if (-not [string]::IsNullOrWhiteSpace($rawText)) {
                [void]$candidateSet.Add($rawText)
            }

            $normalized = Convert-ToMailboxLookupKey -Value $rawText
            if (-not [string]::IsNullOrWhiteSpace($normalized)) {
                [void]$candidateSet.Add($normalized)
            }
        }

        return @($candidateSet)
    }

    $candidateKeys = Get-MailboxLookupCandidates -Record $MailboxRecord -UseArchiveGuid:$Archive

    foreach ($key in $candidateKeys) {
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }

        $rawMatch = Get-ContainerValueByKey -Container $StatsHash -Key $key
        if ($null -ne $rawMatch) {
            return $rawMatch
        }

        $normalizedKey = Convert-ToMailboxLookupKey -Value $key
        if (-not [string]::IsNullOrWhiteSpace($normalizedKey) -and $normalizedKey -ne $key) {
            $normalizedMatch = Get-ContainerValueByKey -Container $StatsHash -Key $normalizedKey
            if ($null -ne $normalizedMatch) {
                return $normalizedMatch
            }
        }
    }

    if ($StatsHash -is [System.Collections.IEnumerable] -and -not ($StatsHash -is [string])) {
        foreach ($statRow in $StatsHash) {
            if ($null -eq $statRow) {
                continue
            }

            $rowCandidates = Get-MailboxLookupCandidates -Record $statRow -UseArchiveGuid:$Archive
            foreach ($rowCandidate in $rowCandidates) {
                if ([string]::IsNullOrWhiteSpace($rowCandidate)) {
                    continue
                }
                if ($candidateKeys -contains $rowCandidate) {
                    return $statRow
                }
                $normalizedRowCandidate = Convert-ToMailboxLookupKey -Value $rowCandidate
                if (-not [string]::IsNullOrWhiteSpace($normalizedRowCandidate) -and $candidateKeys -contains $normalizedRowCandidate) {
                    return $statRow
                }
            }
        }
    }

    return $null
}

function Get-NormalizedMailboxRecords {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [array]$MailboxRecords,
        [AllowNull()]
        $PrimaryMailboxStats,
        [AllowNull()]
        $ArchiveMailboxStats
    )

    if (-not $MailboxRecords -or $MailboxRecords.Count -eq 0) {
        return @()
    }

    $normalized = foreach ($record in $MailboxRecords) {
        if ($null -eq $record) {
            continue
        }

        # Update in place to avoid duplicating large mailbox objects in memory.
        $targetRecord = $record

        $existingMbxSize = Convert-MailboxSizeToGB -SizeValue (Get-RecordValue -Record $targetRecord -Key 'MBXSizeGB')
        $existingArchiveSize = Convert-MailboxSizeToGB -SizeValue (Get-RecordValue -Record $targetRecord -Key 'ArchiveSizeGB')
        $existingMbxItemCount = Get-RecordValue -Record $targetRecord -Key 'MBXItemCount'
        $existingArchiveItemCount = Get-RecordValue -Record $targetRecord -Key 'ArchiveItemCount'
        if (Test-AssessmentHtmlBlankValue -Value $existingMbxItemCount) { $existingMbxItemCount = 0 }
        if (Test-AssessmentHtmlBlankValue -Value $existingArchiveItemCount) { $existingArchiveItemCount = 0 }

        $primaryStat = Get-MailboxStatForRecord -MailboxRecord $targetRecord -StatsHash $PrimaryMailboxStats
        $archiveStat = Get-MailboxStatForRecord -MailboxRecord $targetRecord -StatsHash $ArchiveMailboxStats -Archive

        $primaryTotalItemSize = Get-RecordValue -Record $primaryStat -Key 'TotalItemSize'
        $archiveTotalItemSize = Get-RecordValue -Record $archiveStat -Key 'TotalItemSize'

        $resolvedMbxSize = if ($existingMbxSize -gt 0) { $existingMbxSize } elseif ($primaryStat) { Convert-MailboxSizeToGB -SizeValue $primaryTotalItemSize } else { 0 }
        $resolvedArchiveSize = if ($existingArchiveSize -gt 0) { $existingArchiveSize } elseif ($archiveStat) { Convert-MailboxSizeToGB -SizeValue $archiveTotalItemSize } else { 0 }

        $resolvedMbxItemCount = if (-not (Test-AssessmentHtmlBlankValue -Value $existingMbxItemCount) -and [string]$existingMbxItemCount -ne '0') {
            $existingMbxItemCount
        } elseif ($primaryStat -and $null -ne (Get-RecordValue -Record $primaryStat -Key 'ItemCount')) {
            Get-RecordValue -Record $primaryStat -Key 'ItemCount'
        } else {
            0
        }

        $resolvedArchiveItemCount = if (-not (Test-AssessmentHtmlBlankValue -Value $existingArchiveItemCount) -and [string]$existingArchiveItemCount -ne '0') {
            $existingArchiveItemCount
        } elseif ($archiveStat -and $null -ne (Get-RecordValue -Record $archiveStat -Key 'ItemCount')) {
            Get-RecordValue -Record $archiveStat -Key 'ItemCount'
        } else {
            0
        }

        Set-RecordValue -Record $targetRecord -Key 'MBXSizeGB' -Value ([math]::Round($resolvedMbxSize, 3))
        Set-RecordValue -Record $targetRecord -Key 'ArchiveSizeGB' -Value ([math]::Round($resolvedArchiveSize, 3))
        Set-RecordValue -Record $targetRecord -Key 'MBXItemCount' -Value $resolvedMbxItemCount
        Set-RecordValue -Record $targetRecord -Key 'ArchiveItemCount' -Value $resolvedArchiveItemCount

        $targetRecord
    }

    return @($normalized)
}

function Get-TenantAssessmentContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    $contextKeyDomains = @{
        'AllRecipients'                       = 'Exchange'
        'AllMailboxes'                        = 'Exchange'
        'MailboxFullDetails'                  = 'Exchange'
        'PrimaryMailboxStats'                 = 'Exchange'
        'ArchiveMailboxStats'                 = 'Exchange'
        'ArchiveMailboxes'                    = 'Exchange'
        'InactiveMailboxDetails'              = 'Exchange'
        'PublicFolderDetails'                 = 'Exchange'
        'AllExchangeGroups'                   = 'Exchange'
        'MailFlowConnectors'                  = 'Exchange'
        'EmailActivitySummary'                = 'Exchange'
        'EmailActivityTopSenders'             = 'Exchange'
        'EmailActivityTopReceivers'           = 'Exchange'
        'LicenseSKUs'                         = 'Identity'
        'Users'                               = 'Identity'
        'Admins'                              = 'Identity'
        'EntraIDGroups'                       = 'Identity'
        'DeviceDetails'                       = 'Identity'
        'ConditionalAccessPolicies'           = 'Identity'
        'AuthenticationConfig'                = 'Identity'
        'MfaRegistrationSummary'              = 'Identity'
        'AuthenticationMethods'               = 'Identity'
        'AuthenticationSSOApplications'       = 'Identity'
        'SharePoint'                          = 'Collaboration'
        'OneDrive'                            = 'Collaboration'
        'AllTeams'                            = 'Collaboration'
        'TeamsVoice'                          = 'Collaboration'
        'TeamsVoiceSummary'                   = 'Collaboration'
        'UnmanagedObjects'                    = 'Collaboration'
        'OneDriveOwnerMismatches'             = 'Collaboration'
        'OwnershipGovernanceSummary'          = 'Collaboration'
        'SecuritySecureScore'                 = 'Security'
        'SecureScoreActions'                  = 'Security'
        'SMTPRelaySummary'                    = 'Security'
        'SMTPRelayConfig'                     = 'Security'
        'SpamFilteringSummary'                = 'Security'
        'SpamFilteringConfig'                 = 'Security'
        'Domains'                             = 'Tenant'
        'AdConnectConfiguration'              = 'Tenant'
        'HybridConfiguration'                 = 'Tenant'
        'FederationConfiguration'             = 'Tenant'
        'TenantInfo'                          = 'Tenant'
        'TenantInfoSummary'                   = 'Tenant'
    }

    function Get-ContainerValue {
        param(
            [AllowNull()]$Container,
            [Parameter(Mandatory)][string]$Name
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

    function Resolve-ContextValue {
        param([Parameter(Mandatory)][string]$Key)

        $directValue = Get-ContainerValue -Container $TenantStatsHash -Name $Key
        if ($null -ne $directValue) {
            return $directValue
        }

        $derivedContainer = Get-ContainerValue -Container $TenantStatsHash -Name 'Derived'
        $derivedValue = Get-ContainerValue -Container $derivedContainer -Name $Key
        if ($null -ne $derivedValue) {
            return $derivedValue
        }

        $topLevelCollectorStats = Get-ContainerValue -Container $TenantStatsHash -Name 'CollectorStats'
        $topLevelCollectorStatValue = Get-ContainerValue -Container $topLevelCollectorStats -Name $Key
        if ($null -ne $topLevelCollectorStatValue) {
            return $topLevelCollectorStatValue
        }

        $topLevelSourceCoverage = Get-ContainerValue -Container $TenantStatsHash -Name 'SourceCoverage'
        $topLevelSourceCoverageValue = Get-ContainerValue -Container $topLevelSourceCoverage -Name $Key
        if ($null -ne $topLevelSourceCoverageValue) {
            return $topLevelSourceCoverageValue
        }

        $diagnosticsContainer = Get-ContainerValue -Container $TenantStatsHash -Name 'Diagnostics'
        $collectorStatsContainer = Get-ContainerValue -Container $diagnosticsContainer -Name 'CollectorStats'
        $collectorStatValue = Get-ContainerValue -Container $collectorStatsContainer -Name $Key
        if ($null -ne $collectorStatValue) {
            return $collectorStatValue
        }
        $sourceCoverageContainer = Get-ContainerValue -Container $diagnosticsContainer -Name 'SourceCoverage'
        $sourceCoverageValue = Get-ContainerValue -Container $sourceCoverageContainer -Name $Key
        if ($null -ne $sourceCoverageValue) {
            return $sourceCoverageValue
        }

        $dataContainer = Get-ContainerValue -Container $TenantStatsHash -Name 'Data'
        if ($null -eq $dataContainer) {
            return $null
        }

        if ($contextKeyDomains.ContainsKey($Key)) {
            $domainPath = [string]$contextKeyDomains[$Key]
            $currentContainer = $dataContainer
            foreach ($segment in $domainPath.Split('.', [System.StringSplitOptions]::RemoveEmptyEntries)) {
                $currentContainer = Get-ContainerValue -Container $currentContainer -Name $segment
                if ($null -eq $currentContainer) {
                    break
                }
            }
            $domainMappedValue = Get-ContainerValue -Container $currentContainer -Name $Key
            if ($null -ne $domainMappedValue) {
                return $domainMappedValue
            }
        }

        foreach ($domainName in @('Exchange', 'Identity', 'Collaboration', 'Security', 'Tenant', 'Other')) {
            $domainContainer = Get-ContainerValue -Container $dataContainer -Name $domainName
            if ($null -eq $domainContainer) {
                continue
            }
            $domainValue = Get-ContainerValue -Container $domainContainer -Name $Key
            if ($null -ne $domainValue) {
                return $domainValue
            }
        }

        return $null
    }

    function Get-ContextArray {
        param([string]$Key)

        $value = Resolve-ContextValue -Key $Key
        return @(Convert-AssessmentHtmlArray -InputObject $value)
    }

    function Resolve-ContextSummaryRecord {
        param(
            [AllowNull()]
            $Container
        )

        if ($null -eq $Container) {
            return $null
        }

        $summaryValue = $null
        if ($Container -is [System.Collections.IDictionary]) {
            if ($Container.Contains('Summary')) {
                $summaryValue = $Container['Summary']
            }
            else {
                $summaryValue = $Container
            }
        }
        elseif ($Container -is [System.Collections.Specialized.OrderedDictionary]) {
            if ($Container.Contains('Summary')) {
                $summaryValue = $Container['Summary']
            }
            else {
                $summaryValue = $Container
            }
        }
        elseif ($Container -is [array]) {
            $summaryValue = @($Container | Select-Object -First 1)
            if ($summaryValue.Count -gt 0) {
                $summaryValue = $summaryValue[0]
            }
            else {
                $summaryValue = $null
            }
        }
        else {
            $summaryValue = $Container
        }

        if ($summaryValue -is [System.Collections.IDictionary]) {
            return [PSCustomObject]$summaryValue
        }
        if ($summaryValue -is [System.Collections.Specialized.OrderedDictionary]) {
            return [PSCustomObject]$summaryValue
        }
        if ($summaryValue -is [pscustomobject]) {
            if ($summaryValue.PSObject.Properties['Summary']) {
                return Resolve-ContextSummaryRecord -Container $summaryValue.Summary
            }
            if ($summaryValue.PSObject.Properties['Configuration']) {
                return Resolve-ContextSummaryRecord -Container $summaryValue.Configuration
            }
        }
        return $summaryValue
    }

    $authConfig = $null
    $authContainer = Resolve-ContextValue -Key 'AuthenticationConfig'
    if ($authContainer) {
        $authConfig = (Resolve-ContextSummaryRecord -Container $authContainer)
    }

    $adConnect = Resolve-ContextValue -Key 'AdConnectConfiguration'

    $mfaRegistrationSummary = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'MfaRegistrationSummary')

    $hybridInfo = $null
    $hybridContainer = Resolve-ContextValue -Key 'HybridConfiguration'
    if ($hybridContainer) {
        $hybridInfo = Get-ContainerValue -Container $hybridContainer -Name 'ExchangeHybrid'
    }

    $federationExchange = $null
    $federationCrossTenant = $null
    $federationExternal = $null
    $fedContainer = Resolve-ContextValue -Key 'FederationConfiguration'
    if ($fedContainer) {
        $federationExchange = Get-ContainerValue -Container $fedContainer -Name 'ExchangeFederation'
        $federationCrossTenant = Get-ContainerValue -Container $fedContainer -Name 'CrossTenantAccess'
        $federationExternal = Get-ContainerValue -Container $fedContainer -Name 'ExternalIdentities'
    }

    $teamsVoice = Resolve-ContextValue -Key 'TeamsVoice'

    $spamFilteringSummary = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'SpamFilteringSummary')
    if (-not $spamFilteringSummary) {
        $spamFilteringSummary = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'SpamFilteringConfig')
    }

    $smtpRelaySummary = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'SMTPRelaySummary')
    if (-not $smtpRelaySummary) {
        $smtpRelaySummary = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'SMTPRelayConfig')
    }

    $emailActivitySummary = $null
    $adminReportSettings = $null
    $employeeExperienceInsightsSummary = $null
    $emailActivityContainer = Resolve-ContextValue -Key 'EmailActivitySummary'
    if ($emailActivityContainer) {
        $emailActivitySummary = Resolve-ContextSummaryRecord -Container $emailActivityContainer
        if ($emailActivityContainer -is [System.Collections.IDictionary] -and $emailActivityContainer.Contains('AdminReportSettings')) {
            $adminReportSettings = Resolve-ContextSummaryRecord -Container $emailActivityContainer['AdminReportSettings']
        }
        elseif ($emailActivitySummary -is [System.Collections.IDictionary] -and $emailActivitySummary.Contains('AdminReportSettings')) {
            $adminReportSettings = Resolve-ContextSummaryRecord -Container $emailActivitySummary['AdminReportSettings']
        }
        elseif ($emailActivitySummary -and $emailActivitySummary.PSObject.Properties['AdminReportSettings']) {
            $adminReportSettings = Resolve-ContextSummaryRecord -Container $emailActivitySummary.AdminReportSettings
        }
    }
    if (-not $adminReportSettings) {
        $adminReportSettings = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'AdminReportSettings')
    }
    $employeeExperienceInsightsSummary = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'EmployeeExperienceInsightsSummary')

    $ownershipGovernanceSummary = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'OwnershipGovernanceSummary')

    $primaryMailboxStatsCollectionSummary = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'PrimaryMailboxStatsCollectionSummary')
    if (-not $primaryMailboxStatsCollectionSummary) {
        $primaryMailboxSourceCoverage = Resolve-ContextSummaryRecord -Container (Resolve-ContextValue -Key 'PrimaryMailboxStats')
        $hasSummaryShape = $false
        if ($primaryMailboxSourceCoverage -is [System.Collections.IDictionary] -and $primaryMailboxSourceCoverage.Contains('TotalMailboxes')) {
            $hasSummaryShape = $true
        }
        elseif ($primaryMailboxSourceCoverage.PSObject -and $primaryMailboxSourceCoverage.PSObject.Properties['TotalMailboxes']) {
            $hasSummaryShape = $true
        }
        if ($hasSummaryShape) {
            $primaryMailboxStatsCollectionSummary = $primaryMailboxSourceCoverage
        }
    }

    $mailboxSourceKey = $null
    foreach ($candidateKey in @('MailboxFullDetails', 'AllMailboxes', 'ArchiveMailboxes')) {
        $candidateRows = @(Get-ContextArray -Key $candidateKey)
        if ($candidateRows.Count -gt 0) {
            $mailboxSourceKey = $candidateKey
            break
        }
    }
    $rawMailboxes = if ($mailboxSourceKey) { Get-ContextArray -Key $mailboxSourceKey } else { @() }
    $primaryMailboxStats = Resolve-ContextValue -Key 'PrimaryMailboxStats'
    $archiveMailboxStats = Resolve-ContextValue -Key 'ArchiveMailboxStats'
    $normalizedMailboxes = Get-NormalizedMailboxRecords -MailboxRecords $rawMailboxes -PrimaryMailboxStats $primaryMailboxStats -ArchiveMailboxStats $archiveMailboxStats

    $rawInactiveMailboxes = if ($null -ne (Resolve-ContextValue -Key 'InactiveMailboxDetails')) {
        Get-ContextArray -Key 'InactiveMailboxDetails'
    } else {
        @($normalizedMailboxes | Where-Object { $_.PSObject.Properties['IsInactiveMailbox'] -and $_.IsInactiveMailbox -eq $true })
    }
    $normalizedInactiveMailboxes = Get-NormalizedMailboxRecords -MailboxRecords $rawInactiveMailboxes -PrimaryMailboxStats $primaryMailboxStats -ArchiveMailboxStats $archiveMailboxStats

    if (-not $primaryMailboxStatsCollectionSummary) {
        $primaryStatsCount = 0
        if ($primaryMailboxStats -is [System.Collections.IDictionary]) {
            $primaryStatsCount = @($primaryMailboxStats.Keys).Count
        }
        elseif ($primaryMailboxStats -is [System.Collections.IEnumerable] -and -not ($primaryMailboxStats -is [string])) {
            $primaryStatsCount = @($primaryMailboxStats).Count
        }

        if ($primaryStatsCount -gt 0 -or $normalizedMailboxes.Count -gt 0) {
            $derivedTotalMailboxCount = if ($normalizedMailboxes.Count -gt 0) { $normalizedMailboxes.Count } else { $primaryStatsCount }
            $derivedInactiveMailboxCount = @(
                $normalizedMailboxes | Where-Object {
                    $_ -and $_.PSObject -and $_.PSObject.Properties['IsInactiveMailbox'] -and $_.IsInactiveMailbox -eq $true
                }
            ).Count
            $derivedActiveMailboxCount = [Math]::Max(($derivedTotalMailboxCount - $derivedInactiveMailboxCount), 0)
            $derivedMissingMailboxCount = [Math]::Max(($derivedTotalMailboxCount - $primaryStatsCount), 0)

            $primaryMailboxStatsCollectionSummary = [PSCustomObject]@{
                TotalMailboxes       = [int]$derivedTotalMailboxCount
                ActiveMailboxes      = [int]$derivedActiveMailboxCount
                GraphPopulated       = 0
                ExoFallbackPopulated = 0
                Populated            = [int]$primaryStatsCount
                Missing              = [int]$derivedMissingMailboxCount
            }
        }
    }

    return [PSCustomObject]@{
        Licenses               = Get-ContextArray -Key 'LicenseSKUs'
        Recipients             = Get-ContextArray -Key 'AllRecipients'
        Mailboxes              = $normalizedMailboxes
        InactiveMailboxes      = $normalizedInactiveMailboxes
        PublicFolders          = Get-ContextArray -Key 'PublicFolderDetails'
        SharePoint             = Get-ContextArray -Key 'SharePoint'
        OneDrive               = Get-ContextArray -Key 'OneDrive'
        Domains                = Get-ContextArray -Key 'Domains'
        Devices                = Get-ContextArray -Key 'DeviceDetails'
        SecureScore            = Get-ContextArray -Key 'SecuritySecureScore'
        SecureScoreActions     = Get-ContextArray -Key 'SecureScoreActions'
        Teams                  = Get-ContextArray -Key 'AllTeams'
        Users                  = Get-ContextArray -Key 'Users'
        Admins                 = Get-ContextArray -Key 'Admins'
        Groups                 = Get-ContextArray -Key 'EntraIDGroups'
        ExchangeGroups         = Get-ContextArray -Key 'AllExchangeGroups'
        ConditionalAccess      = Get-ContextArray -Key 'ConditionalAccessPolicies'
        MailFlowConnectors     = Get-ContextArray -Key 'MailFlowConnectors'
        RemoteDomains          = Get-ContextArray -Key 'RemoteDomains'
        AuthConfig             = $authConfig
        MfaRegistrationSummary = $mfaRegistrationSummary
        AdConnect              = $adConnect
        HybridInfo             = $hybridInfo
        FederationExchange     = $federationExchange
        FederationCrossTenant  = $federationCrossTenant
        FederationExternal     = $federationExternal
        TeamsVoice             = $teamsVoice
        SpamFilteringSummary   = $spamFilteringSummary
        SMTPRelaySummary       = $smtpRelaySummary
        EmailActivitySummary   = $emailActivitySummary
        EmailActivityTopSenders = Get-ContextArray -Key 'EmailActivityTopSenders'
        EmailActivityTopReceivers = Get-ContextArray -Key 'EmailActivityTopReceivers'
        TeamsActivityTopUsers  = Get-ContextArray -Key 'TeamsActivityTopUsers'
        Office365GroupsActivityTopGroups = Get-ContextArray -Key 'Office365GroupsActivityTopGroups'
        EmployeeExperienceInsightsSummary = $employeeExperienceInsightsSummary
        AdminReportSettings    = $adminReportSettings
        PrimaryMailboxStatsCollectionSummary = $primaryMailboxStatsCollectionSummary
        OwnershipGovernanceSummary = $ownershipGovernanceSummary
        UnmanagedObjects       = Get-ContextArray -Key 'UnmanagedObjects'
        OneDriveOwnerMismatches = Get-ContextArray -Key 'OneDriveOwnerMismatches'
    }
}

function Get-TenantIdentityDetails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    function Get-IdentityValue {
        param(
            [AllowNull()]
            $Record,
            [Parameter(Mandatory)]
            [string[]]$Names
        )

        if ($null -eq $Record) {
            return $null
        }

        foreach ($name in $Names) {
            if ($Record -is [System.Collections.IDictionary] -and $Record.Contains($name)) {
                $value = $Record[$name]
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    return [string]$value
                }
            }
            elseif ($Record.PSObject -and $Record.PSObject.Properties[$name]) {
                $value = $Record.PSObject.Properties[$name].Value
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    return [string]$value
                }
            }
        }

        return $null
    }

    $tenantInfoRecord = $null
    if ($TenantStatsHash.ContainsKey('TenantInfo') -and $TenantStatsHash['TenantInfo']) {
        $tenantInfoRecord = $TenantStatsHash['TenantInfo']
    }

    $tenantInfoSummaryRecord = $null
    if ($TenantStatsHash.ContainsKey('TenantInfoSummary') -and $TenantStatsHash['TenantInfoSummary']) {
        $summaryContainer = $TenantStatsHash['TenantInfoSummary']
        if ($summaryContainer -is [System.Collections.IDictionary] -and $summaryContainer.Contains('Summary')) {
            $tenantInfoSummaryRecord = $summaryContainer['Summary']
        }
        elseif ($summaryContainer.PSObject -and $summaryContainer.PSObject.Properties['Summary']) {
            $tenantInfoSummaryRecord = $summaryContainer.Summary
        }
        else {
            $tenantInfoSummaryRecord = $summaryContainer
        }
    }

    $displayName = Get-IdentityValue -Record $tenantInfoRecord -Names @('DisplayName')
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = Get-IdentityValue -Record $tenantInfoSummaryRecord -Names @('DisplayName', 'TenantName', 'OrganizationName')
    }

    $tenantId = Get-IdentityValue -Record $tenantInfoRecord -Names @('TenantId', 'Id')
    if ([string]::IsNullOrWhiteSpace($tenantId)) {
        $tenantId = Get-IdentityValue -Record $tenantInfoSummaryRecord -Names @('TenantId', 'Id')
    }

    $defaultDomain = Get-IdentityValue -Record $tenantInfoRecord -Names @('DefaultDomain')
    if ([string]::IsNullOrWhiteSpace($defaultDomain)) {
        $defaultDomain = Get-IdentityValue -Record $tenantInfoSummaryRecord -Names @('DefaultDomain')
    }

    $initialDomain = Get-IdentityValue -Record $tenantInfoRecord -Names @('InitialDomainName', 'InitialDomain')
    if ([string]::IsNullOrWhiteSpace($initialDomain)) {
        $initialDomain = Get-IdentityValue -Record $tenantInfoSummaryRecord -Names @('InitialDomain', 'InitialDomainName')
    }

    $country = Get-IdentityValue -Record $tenantInfoRecord -Names @('Country', 'CountryLetterCode')
    if ([string]::IsNullOrWhiteSpace($country)) {
        $country = Get-IdentityValue -Record $tenantInfoSummaryRecord -Names @('Country', 'CountryLetterCode')
    }

    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = 'Microsoft 365 Tenant'
    }

    return [PSCustomObject]@{
        DisplayName   = $displayName
        TenantId      = $tenantId
        DefaultDomain = $defaultDomain
        InitialDomain = $initialDomain
        Country       = $country
    }
}

function Update-AssessmentReportTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    if (-not $TenantStatsHash) {
        return
    }

    $context = Get-TenantAssessmentContext -TenantStatsHash $TenantStatsHash

    $TenantStatsHash['BestPractices'] = @{}
    $TenantStatsHash['BestPracticeFindings'] = @{}
    $TenantStatsHash['MigrationReadiness'] = @{}

    $summaryRows = New-Object System.Collections.Generic.List[object]
    $findingRows = New-Object System.Collections.Generic.List[object]
    $migrationRows = New-Object System.Collections.Generic.List[object]
    $summaryIndex = 0
    $findingIndex = 0
    $migrationIndex = 0

    function Add-AreaSummary {
        param(
            [string]$Area,
            [array]$AreaFindings,
            [string]$AssessmentType,
            [string]$RelatedWorksheet,
            [string]$Notes
        )

        $criticalCount = @($AreaFindings | Where-Object { $_.Type -eq 'Risk' }).Count
        $warningCount = @($AreaFindings | Where-Object { $_.Type -eq 'Warning' }).Count
        $infoCount = @($AreaFindings | Where-Object { $_.Type -eq 'Info' }).Count

        $status = if ($criticalCount -gt 0) {
            'Critical'
        } elseif ($warningCount -gt 0) {
            'Warning'
        } elseif ($AreaFindings.Count -gt 0) {
            'Informational'
        } else {
            'Healthy'
        }

        $topFinding = @(
            $AreaFindings |
                Sort-Object @{ Expression = {
                    switch ([string]$_.Type) {
                        'Risk' { 1 }
                        'Warning' { 2 }
                        'Info' { 3 }
                        default { 9 }
                    }
                } }, Priority |
                Select-Object -First 1
        )
        $topCategoryRollups = @(
            $AreaFindings |
                Group-Object Category |
                Sort-Object @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Descending = $false } |
                Select-Object -First 3 |
                ForEach-Object {
                    $categoryName = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Uncategorized' } else { $_.Name }
                    "{0}: {1}" -f $categoryName, $_.Count
                }
        )
        $primaryFindingText = if ($AreaFindings.Count -gt 0) {
            $severitySummary = "{0} critical, {1} warning, {2} informational finding(s)" -f $criticalCount, $warningCount, $infoCount
            if ($topCategoryRollups.Count -gt 0) {
                "$severitySummary. Top signals: $([string]::Join('; ', $topCategoryRollups))."
            } else {
                "$severitySummary."
            }
        } else {
            'No automated findings detected for this assessment area.'
        }
        $recommendedAction = if ($topFinding.Count -gt 0) { Get-ArrayaAssessmentRecommendationText -Finding $topFinding[0] } else { 'Use the detailed workload worksheets for validation and migration planning.' }

        $summaryRows.Add([PSCustomObject]@{
            Area               = $Area
            Status             = $status
            CriticalFindings   = $criticalCount
            WarningFindings    = $warningCount
            InfoFindings       = $infoCount
            TotalFindings      = $AreaFindings.Count
            PrimaryFinding     = $primaryFindingText
            RecommendedAction  = $recommendedAction
            RelatedWorksheet   = $RelatedWorksheet
            AssessmentType     = $AssessmentType
            Notes              = $Notes
        }) | Out-Null

        foreach ($finding in $AreaFindings) {
            $findingRows.Add([PSCustomObject]@{
                Area               = $Area
                Severity           = $finding.Type
                Category           = $finding.Category
                Message            = $finding.Message
                Priority           = $finding.Priority
                RelatedSection     = $finding.Anchor
                RelatedWorksheet   = Get-ArrayaAssessmentWorksheetName -Anchor $finding.Anchor
                RecommendedAction  = Get-ArrayaAssessmentRecommendationText -Finding $finding
                SourceType         = $AssessmentType
            }) | Out-Null
        }
    }

    function Add-MigrationRow {
        param(
            [string]$Category,
            [string]$Item,
            [string]$Status,
            [string]$Value,
            [string]$Notes,
            [string]$MigrationAction,
            [string]$SourceWorksheet
        )

        $migrationRows.Add([PSCustomObject]@{
            Category         = $Category
            Item             = $Item
            Status           = $Status
            Value            = $Value
            Notes            = $Notes
            MigrationAction  = $MigrationAction
            SourceWorksheet  = $SourceWorksheet
        }) | Out-Null
    }

    if ($context.Licenses.Count -gt 0) {
        $licAnalysis = Get-LicenseAnalysis -Licenses $context.Licenses -UserCount $context.Users.Count
        Add-AreaSummary -Area 'Licensing' -AreaFindings $licAnalysis.Findings -AssessmentType 'Assessment heuristic using Microsoft 365 license data' -RelatedWorksheet 'LicenseSKUs' -Notes 'Evaluates capacity, at-capacity SKUs, and high utilization.'
    }

    if ($context.Domains.Count -gt 0) {
        $domainAnalysis = Get-DomainAnalysis -Domains $context.Domains -SpamFilteringSummary $context.SpamFilteringSummary -SMTPRelaySummary $context.SMTPRelaySummary
        Add-AreaSummary -Area 'Domains' -AreaFindings $domainAnalysis.Findings -AssessmentType 'Assessment heuristic using Microsoft 365 domain state' -RelatedWorksheet 'Domains' -Notes 'Highlights verification and mail-routing concerns relevant to migration cutover.'
    }

    if ($context.Users.Count -gt 0 -or $context.Admins.Count -gt 0 -or $context.Groups.Count -gt 0) {
        $identityAnalysis = Get-IdentityAdminAnalysis -Users $context.Users -Admins $context.Admins -Groups $context.Groups -ConditionalAccessPolicies $context.ConditionalAccess
        Add-AreaSummary -Area 'Identity & Admins' -AreaFindings $identityAnalysis.Findings -AssessmentType 'Assessment heuristic using Entra users, groups, and admin assignments' -RelatedWorksheet 'Users' -Notes 'Surfaces admin and group inventory signals that affect migration readiness.'
    }

    if ($context.Mailboxes.Count -gt 0 -or $context.PublicFolders.Count -gt 0) {
        $mailboxAnalysis = Get-MailboxAnalysis -Mailboxes $context.Mailboxes
        Add-AreaSummary -Area 'Mailboxes' -AreaFindings $mailboxAnalysis.Findings -AssessmentType 'Assessment heuristic using Exchange mailbox inventory' -RelatedWorksheet 'MailboxFullDetails' -Notes 'Flags mailbox sizing and archive patterns that influence batch and exception planning.'
    }

    if ($context.Mailboxes.Count -gt 0) {
        $complianceRetentionAnalysis = Get-ComplianceRetentionAnalysis -Mailboxes $context.Mailboxes
        Add-AreaSummary -Area 'Compliance & Retention' -AreaFindings $complianceRetentionAnalysis.Findings -AssessmentType 'Assessment heuristic using oversized Exchange mailbox/archive inventory and hold coverage' -RelatedWorksheet 'MailboxFullDetails' -Notes 'Compares mailboxes and archives above size thresholds to litigation hold and retention-hold/control signals.'
    }

    if ($context.Teams.Count -gt 0 -or $context.TeamsActivityTopUsers.Count -gt 0) {
        $teamsAnalysis = Get-TeamsCollaborationAnalysis -Teams $context.Teams -TeamsTopUsers $context.TeamsActivityTopUsers -TeamsVoice $context.TeamsVoice
        Add-AreaSummary -Area 'Teams Collaboration' -AreaFindings $teamsAnalysis.Findings -AssessmentType 'Assessment heuristic using Teams inventory, ownership, and activity telemetry' -RelatedWorksheet 'AllTeams' -Notes 'Highlights Teams ownership posture, channel footprint, and activity coverage signals.'
    }

    if (
        $context.EmailActivitySummary -or
        $context.EmailActivityTopSenders.Count -gt 0 -or
        $context.EmailActivityTopReceivers.Count -gt 0 -or
        $context.EmployeeExperienceInsightsSummary -or
        $context.TeamsActivityTopUsers.Count -gt 0 -or
        $context.Office365GroupsActivityTopGroups.Count -gt 0
    ) {
        $employeeExperienceAnalysis = Get-ArrayaEmployeeExperienceInsightsAnalysis `
            -EmailActivitySummary $context.EmailActivitySummary `
            -EmployeeExperienceInsightsSummary $context.EmployeeExperienceInsightsSummary `
            -TopSenders $context.EmailActivityTopSenders `
            -TopReceivers $context.EmailActivityTopReceivers `
            -TeamsTopUsers $context.TeamsActivityTopUsers `
            -GroupsTopGroups $context.Office365GroupsActivityTopGroups `
            -AdminReportSettings $context.AdminReportSettings `
            -AuthConfig $context.AuthConfig `
            -ConditionalAccessPolicies $context.ConditionalAccess `
            -TotalUsers $context.Users.Count
        Add-AreaSummary -Area 'Zero Trust Signals' -AreaFindings $employeeExperienceAnalysis.Findings -AssessmentType 'Assessment heuristic using security telemetry quality, Conditional Access posture, and authentication controls' -RelatedWorksheet 'EmployeeExpInsights' -Notes 'Uses Microsoft 365 activity telemetry visibility, Conditional Access signals, MFA/passwordless posture, and high-volume account patterns to guide zero-trust improvements.'
    }

    if ($context.InactiveMailboxes.Count -gt 0) {
        $inactiveAnalysis = Get-InactiveMailboxAnalysis -InactiveMailboxes $context.InactiveMailboxes
        Add-AreaSummary -Area 'Inactive Mailboxes' -AreaFindings $inactiveAnalysis.Findings -AssessmentType 'Assessment heuristic using inactive mailbox inventory' -RelatedWorksheet 'InactiveMailboxDetails' -Notes 'Supports retention and scope decisions for mailbox migration.'
    }

    if ($context.SharePoint.Count -gt 0 -or $context.OneDrive.Count -gt 0) {
        $spodAnalysis = Get-SharePointOneDriveAnalysis -SharePointSites $context.SharePoint -OneDriveSites $context.OneDrive
        Add-AreaSummary -Area 'SharePoint & OneDrive' -AreaFindings $spodAnalysis.Findings -AssessmentType 'Assessment heuristic using collaboration site inventory' -RelatedWorksheet 'SharePoint / OneDrive' -Notes 'Flags oversized sites and owner-linked OneDrive inventory for migration planning.'
    }

    if (
        $context.OwnershipGovernanceSummary -or
        $context.UnmanagedObjects.Count -gt 0 -or
        $context.OneDriveOwnerMismatches.Count -gt 0
    ) {
        $ownershipAnalysis = Get-OwnershipGovernanceAnalysis `
            -UnmanagedObjects $context.UnmanagedObjects `
            -OneDriveOwnerMismatches $context.OneDriveOwnerMismatches `
            -OwnershipGovernanceSummary $context.OwnershipGovernanceSummary
        Add-AreaSummary -Area 'Ownership & Stewardship' -AreaFindings $ownershipAnalysis.Findings -AssessmentType 'Assessment heuristic using ownership governance signals across collaboration and group workloads' -RelatedWorksheet 'UnmanagedObjects' -Notes 'Highlights unowned objects, unhealthy owners, and OneDrive owner-mismatch governance review items.'
    }

    if ($context.Devices.Count -gt 0) {
        $deviceAnalysis = Get-DeviceAnalysis -Devices $context.Devices
        Add-AreaSummary -Area 'Devices' -AreaFindings $deviceAnalysis.Findings -AssessmentType 'Assessment heuristic using Entra device inventory' -RelatedWorksheet 'DeviceDetails' -Notes 'Highlights stale and non-compliant devices that can affect user cutover readiness.'
    }

    if ($context.AdConnect) {
        $adAnalysis = Get-AdConnectAnalysis -AdConnect $context.AdConnect
        Add-AreaSummary -Area 'AD Connect / Sync' -AreaFindings $adAnalysis.Findings -AssessmentType 'Assessment heuristic using directory synchronization signals' -RelatedWorksheet 'AdConnectConfiguration' -Notes 'Identifies synchronization dependencies and recent sync issues.'
    }

    if ($context.ConditionalAccess.Count -gt 0 -or $context.AuthConfig) {
        $caAnalysis = Get-ConditionalAccessMfaAnalysis -ConditionalAccessPolicies $context.ConditionalAccess -AuthConfig $context.AuthConfig -TotalUsers $context.Users.Count
        Add-AreaSummary -Area 'Conditional Access & MFA' -AreaFindings $caAnalysis.Findings -AssessmentType 'Assessment heuristic using Entra protection settings' -RelatedWorksheet 'ConditionalAccessPolicies' -Notes 'Evaluates MFA and Conditional Access coverage with a Microsoft best-practice orientation.'
    }

    if ($context.HybridInfo) {
        $hybridAnalysis = Get-ExchangeHybridAnalysis -HybridInfo $context.HybridInfo
        $fedAnalysis = Get-FederationAnalysis -ExchangeFederation $context.FederationExchange -CrossTenantAccess $context.FederationCrossTenant -ExternalIdentities $context.FederationExternal
        $hybridFindings = @($hybridAnalysis.Findings + $fedAnalysis.Findings)
        Add-AreaSummary -Area 'Hybrid / Federation' -AreaFindings $hybridFindings -AssessmentType 'Assessment heuristic using Exchange hybrid and federation signals' -RelatedWorksheet 'HybridConfiguration' -Notes 'Highlights hybrid dependencies, federation, and cross-tenant settings relevant to coexistence and migration.'
    }

    if ($context.SecureScore.Count -gt 0) {
        $latestScore = @($context.SecureScore | Sort-Object CreatedDateTime -Descending | Select-Object -First 1)
        $secureScoreFindings = @()
        if ($latestScore.Count -gt 0) {
            $scorePct = 0
            try { $scorePct = [double]$latestScore[0].SecurityScorePercentage } catch { $scorePct = 0 }
            if ($scorePct -lt 60) {
                $secureScoreFindings += @{
                    Type = 'Risk'
                    Category = 'Secure Score'
                    Message = "Microsoft Secure Score is $([math]::Round($scorePct,1))%, which is below the target range for a mature tenant baseline."
                    Anchor = 'secure-score'
                    Priority = 1
                }
            } elseif ($scorePct -lt 80) {
                $secureScoreFindings += @{
                    Type = 'Warning'
                    Category = 'Secure Score'
                    Message = "Microsoft Secure Score is $([math]::Round($scorePct,1))%; prioritize high-rank actions to raise baseline security."
                    Anchor = 'secure-score'
                    Priority = 2
                }
            } else {
                $secureScoreFindings += @{
                    Type = 'Info'
                    Category = 'Secure Score'
                    Message = "Microsoft Secure Score is $([math]::Round($scorePct,1))%, indicating a relatively strong security baseline."
                    Anchor = 'secure-score'
                    Priority = 3
                }
            }
        }
        Add-AreaSummary -Area 'Secure Score' -AreaFindings $secureScoreFindings -AssessmentType 'Microsoft Secure Score recommendation mapping' -RelatedWorksheet 'SecureScoreActions' -Notes 'Uses Microsoft Secure Score snapshots and control profile metadata, including Microsoft Learn action URLs.'
    }

    foreach ($summaryRow in $summaryRows) {
        $summaryIndex++
        $TenantStatsHash['BestPractices'][("{0:D3}-{1}" -f $summaryIndex, $summaryRow.Area)] = $summaryRow
    }

    foreach ($findingRow in $findingRows) {
        $findingIndex++
        $TenantStatsHash['BestPracticeFindings'][("{0:D3}-{1}" -f $findingIndex, $findingRow.Area)] = $findingRow
    }

    $verifiedDomains = @($context.Domains | Where-Object { $_.Verified -eq $true }).Count
    $unverifiedDomains = @($context.Domains | Where-Object { $_.Verified -ne $true }).Count
    $nonM365MxDomains = @($context.Domains | Where-Object { $_.Office365MailExchanger -eq $false }).Count
    $dirSyncEnabled = [bool]($context.AdConnect -and $context.AdConnect.Summary -and $context.AdConnect.Summary.OnPremisesSyncEnabled -eq $true)
    $guestCount = @($context.Users | Where-Object { $_.UserType -match 'Guest' -or $_.UserPrincipalName -like '*#EXT#*' }).Count
    $licensedUsers = @($context.Users | Where-Object { $_.AssignedLicenses }).Count
    $archiveMailboxCount = @($context.Mailboxes | Where-Object { $_.ArchiveStatus -and $_.ArchiveStatus -ne 'None' }).Count
    $publicFolderCount = @($context.PublicFolders).Count
    $connectorCount = @($context.MailFlowConnectors).Count
    $hybridDetected = [bool]($context.HybridInfo -and (($context.HybridInfo.IsHybridConfigured -eq $true) -or ($context.HybridInfo.MigrationEndpointCount -gt 0) -or ($context.HybridInfo.EvidenceCount -gt 0)))
    $crossTenantPartnerCount = if ($context.FederationCrossTenant) { [int]$context.FederationCrossTenant.PartnerCount } else { 0 }
    $teamsCollected = $TenantStatsHash.ContainsKey('AllTeams')
    $paidLicenseAnalysis = if ($context.Licenses.Count -gt 0) { Get-LicenseAnalysis -Licenses $context.Licenses -UserCount $context.Users.Count } else { $null }
    $overallLicenseUtilization = if ($paidLicenseAnalysis -and $paidLicenseAnalysis.TotalPurchased -gt 0) {
        [math]::Round((($paidLicenseAnalysis.TotalConsumed / $paidLicenseAnalysis.TotalPurchased) * 100), 1)
    } else {
        $null
    }
    $voiceSummary = if ($context.TeamsVoice -and $context.TeamsVoice.ContainsKey('Summary')) { $context.TeamsVoice['Summary'] } else { $null }

    Add-MigrationRow -Category 'Domains' -Item 'Verified custom domains' -Status $(if ($unverifiedDomains -gt 0) { 'Blocker' } else { 'Ready' }) -Value "$verifiedDomains verified / $unverifiedDomains unverified" -Notes 'All accepted domains should be validated and sequenced for migration and cutover.' -MigrationAction 'Confirm domain ownership, cutover timing, and accepted domain strategy in the target tenant.' -SourceWorksheet 'Domains'
    Add-MigrationRow -Category 'Domains' -Item 'Mail routing' -Status $(if ($nonM365MxDomains -gt 0) { 'Review' } else { 'Ready' }) -Value "$nonM365MxDomains domain(s) with non-Microsoft 365 MX" -Notes 'Non-M365 MX routing can indicate third-party filtering, staged coexistence, or non-standard cutover requirements.' -MigrationAction 'Document current MX and transport path before migration planning.' -SourceWorksheet 'Domains'
    Add-MigrationRow -Category 'Identity' -Item 'Directory synchronization' -Status $(if ($dirSyncEnabled) { 'Review' } else { 'Ready' }) -Value $(if ($dirSyncEnabled) { 'On-prem sync enabled' } else { 'Cloud-only identity model' }) -Notes 'Hybrid identity affects object authority and user cutover sequencing.' -MigrationAction 'Plan whether identities stay synced during migration or transition to cloud-managed.' -SourceWorksheet 'AdConnectConfiguration'
    Add-MigrationRow -Category 'Identity' -Item 'Guests and external identities' -Status $(if ($guestCount -gt 0) { 'Review' } else { 'Info' }) -Value "$guestCount guest/external user(s)" -Notes 'Guest access usually requires separate planning from member user migration.' -MigrationAction 'Decide whether guest objects are recreated, invited, or excluded from scope.' -SourceWorksheet 'Users'
    Add-MigrationRow -Category 'Messaging' -Item 'Mailbox inventory' -Status 'Info' -Value "$($context.Mailboxes.Count) mailbox(es), $archiveMailboxCount archive-enabled" -Notes 'Mailbox and archive counts drive migration batch sizing and exception planning.' -MigrationAction 'Use mailbox detail sheets to segment batches and identify oversized or special-case mailboxes.' -SourceWorksheet 'MailboxFullDetails'
    Add-MigrationRow -Category 'Messaging' -Item 'Inactive mailboxes' -Status $(if ($context.InactiveMailboxes.Count -gt 0) { 'Review' } else { 'Ready' }) -Value "$($context.InactiveMailboxes.Count) inactive mailbox(es)" -Notes 'Inactive mailboxes may be retained for compliance rather than migrated.' -MigrationAction 'Confirm retention, restore, or exclusion decisions before migration scope is finalized.' -SourceWorksheet 'InactiveMailboxDetails'
    Add-MigrationRow -Category 'Messaging' -Item 'Public folders' -Status $(if ($publicFolderCount -gt 0) { 'Review' } else { 'Ready' }) -Value "$publicFolderCount public folder object(s)" -Notes 'Public folders frequently require separate migration tooling or remediation.' -MigrationAction 'Validate whether public folders remain in scope and determine their target-state strategy.' -SourceWorksheet 'PublicFolderDetails'
    Add-MigrationRow -Category 'Messaging' -Item 'Mail flow dependencies' -Status $(if ($connectorCount -gt 0) { 'Review' } else { 'Ready' }) -Value "$connectorCount connector(s), $($context.RemoteDomains.Count) remote domain(s)" -Notes 'Connectors and remote domains can indicate coexistence, partner routing, or relay dependencies.' -MigrationAction 'Inventory connectors, relay paths, and remote domains before cutover design.' -SourceWorksheet 'MailFlowConnectors'
    Add-MigrationRow -Category 'Collaboration' -Item 'SharePoint and OneDrive' -Status 'Info' -Value "$($context.SharePoint.Count) SharePoint site(s), $($context.OneDrive.Count) OneDrive site(s)" -Notes 'Collaboration workload size and ownership patterns influence tooling and wave planning.' -MigrationAction 'Use site inventory, size, and owner data to prioritize migration waves.' -SourceWorksheet 'SharePoint / OneDrive'
    Add-MigrationRow -Category 'Collaboration' -Item 'Teams workload data' -Status $(if ($teamsCollected) { 'Info' } else { 'Needs Data' }) -Value $(if ($teamsCollected) { "$($context.Teams.Count) team(s) collected" } else { 'Teams inventory not collected in current run' }) -Notes 'Teams topology may require delegated or expanded app permissions beyond this cert-auth path.' -MigrationAction 'Collect Teams team/channel inventory before finalizing collaboration migration planning.' -SourceWorksheet $(if ($teamsCollected) { 'AllTeams' } else { 'N/A' })
    Add-MigrationRow -Category 'Security' -Item 'Conditional Access and MFA' -Status $(if ($context.ConditionalAccess.Count -gt 0 -or $context.AuthConfig) { 'Info' } else { 'Review' }) -Value "$($context.ConditionalAccess.Count) CA policy/policies; MFA summary collected=$(if ($null -ne $context.MfaRegistrationSummary) { 'Yes' } else { 'No' })" -Notes 'Security controls need parity planning to avoid cutover lockouts.' -MigrationAction 'Map CA, MFA, and authentication controls between source and target tenant.' -SourceWorksheet 'ConditionalAccessPolicies'
    Add-MigrationRow -Category 'Hybrid' -Item 'Hybrid or coexistence indicators' -Status $(if ($hybridDetected) { 'Review' } else { 'Ready' }) -Value $(if ($hybridDetected) { "Hybrid signals detected; migration endpoints=$($context.HybridInfo.MigrationEndpointCount)" } else { 'No hybrid indicators detected' }) -Notes 'Hybrid configuration affects mailbox authority, routing, and migration tooling choices.' -MigrationAction 'Validate whether hybrid remains required during migration or can be removed from scope.' -SourceWorksheet 'HybridConfiguration'
    Add-MigrationRow -Category 'External Access' -Item 'Cross-tenant and B2B settings' -Status $(if ($crossTenantPartnerCount -gt 0) { 'Review' } else { 'Info' }) -Value "$crossTenantPartnerCount partner relationship(s)" -Notes 'Cross-tenant policies may affect coexistence and post-migration collaboration behavior.' -MigrationAction 'Review B2B and cross-tenant access settings as part of coexistence planning.' -SourceWorksheet 'FederationConfiguration'
    Add-MigrationRow -Category 'Licensing' -Item 'Target licensing readiness' -Status $(if ($null -ne $overallLicenseUtilization -and $overallLicenseUtilization -ge 85) { 'Review' } else { 'Info' }) -Value $(if ($null -ne $overallLicenseUtilization) { "$overallLicenseUtilization% utilized; $licensedUsers licensed user(s)" } else { 'License utilization unavailable' }) -Notes 'Target tenant licensing should be aligned before user and workload onboarding.' -MigrationAction 'Review paid SKU utilization and confirm target-tenant licensing for migration scope.' -SourceWorksheet 'LicenseSKUs'
    Add-MigrationRow -Category 'Teams Voice' -Item 'Voice workload readiness' -Status $(if ($voiceSummary -and $voiceSummary.PSObject.Properties['DataSource'] -and $voiceSummary.DataSource -eq 'GraphLicenseInference') { 'Needs Data' } else { 'Info' }) -Value $(if ($voiceSummary) { "Voice users=$($voiceSummary.VoiceUserCount); source=$($voiceSummary.DataSource)" } else { 'No Teams voice summary collected' }) -Notes 'Current app-auth path infers voice licensing but does not capture full PSTN or number-assignment state.' -MigrationAction 'Add Teams voice/call record permissions or collect delegated Teams PowerShell data before final voice migration planning.' -SourceWorksheet 'TeamsVoice'

    foreach ($migrationRow in $migrationRows) {
        $migrationIndex++
        $TenantStatsHash['MigrationReadiness'][("{0:D3}-{1}" -f $migrationIndex, $migrationRow.Item)] = $migrationRow
    }

    Write-Log -Type INFO -Message "[Update-AssessmentReportTables] Created $($TenantStatsHash['BestPractices'].Count) best-practice summary rows, $($TenantStatsHash['BestPracticeFindings'].Count) detailed findings, and $($TenantStatsHash['MigrationReadiness'].Count) migration readiness rows" -ExportFileLocation $ExportDetails
}

function Update-ConfigurationSummaryTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    if (-not $TenantStatsHash) {
        return
    }

    foreach ($summaryKey in @('TenantInfoSummary', 'SpamFilteringSummary', 'SMTPRelaySummary', 'FederationSummary', 'TeamsVoiceSummary')) {
        if (-not $TenantStatsHash.ContainsKey($summaryKey) -or -not $TenantStatsHash[$summaryKey]) {
            $TenantStatsHash[$summaryKey] = @{}
            continue
        }

        if ($TenantStatsHash[$summaryKey] -is [System.Collections.IDictionary]) {
            continue
        }

        $existingSummary = $null
        if ($TenantStatsHash[$summaryKey].PSObject -and $TenantStatsHash[$summaryKey].PSObject.Properties['Summary']) {
            $existingSummary = $TenantStatsHash[$summaryKey].Summary
        }
        $TenantStatsHash[$summaryKey] = @{}
        if ($existingSummary) {
            $TenantStatsHash[$summaryKey]['Summary'] = $existingSummary
        }
    }

    if ($TenantStatsHash.ContainsKey('TenantInfo') -and $TenantStatsHash['TenantInfo']) {
        $tenantInfo = $TenantStatsHash['TenantInfo']
        $selfService = $tenantInfo.SelfServicePurchase
        $azureUsage = $tenantInfo.AzureResourceUsage
        $TenantStatsHash['TenantInfoSummary']['Summary'] = [PSCustomObject]@{
            DisplayName             = $tenantInfo.DisplayName
            TenantId                = $tenantInfo.TenantId
            InitialDomain           = $tenantInfo.InitialDomain
            DefaultDomain           = $tenantInfo.DefaultDomain
            Country                 = $tenantInfo.Country
            CountryLetterCode       = $tenantInfo.CountryLetterCode
            PreferredDataLocation   = $tenantInfo.PreferredDataLocation
            MultiGeoEnabled         = $tenantInfo.MultiGeoEnabled
            MultiGeoAllowed         = $(if ($tenantInfo.MultiGeoAllowed) { $tenantInfo.MultiGeoAllowed -join ', ' } else { $null })
            MultiGeoCentral         = $tenantInfo.MultiGeoCentral
            SelfServicePurchase     = $(if ($selfService) { $selfService.Answer } else { $null })
            SelfServicePurchaseNotes = $(if ($selfService) { $selfService.Notes } else { $null })
            AzureResourceUsage      = $(if ($azureUsage) { $azureUsage.Answer } else { $null })
            AzureResourceUsageNotes = $(if ($azureUsage) { $azureUsage.Notes } else { $null })
        }
    }

    if ($TenantStatsHash.ContainsKey('SpamFilteringConfig') -and $TenantStatsHash['SpamFilteringConfig'].ContainsKey('Configuration')) {
        $spamConfig = $TenantStatsHash['SpamFilteringConfig']['Configuration']
        $TenantStatsHash['SpamFilteringSummary']['Summary'] = [PSCustomObject]@{
            Uses3rdPartyFiltering      = $spamConfig.Uses3rdPartyFiltering
            InboundConnectorCount      = $spamConfig.InboundConnectorCount
            OutboundConnectorCount     = $spamConfig.OutboundConnectorCount
            TransportRuleCount         = $spamConfig.TransportRuleCount
            TransportRulesWithTrustedIPs = $spamConfig.TransportRulesWithTrustedIPs
            PotentialSpamFilters       = $(if ($spamConfig.PotentialSpamFilters) { $spamConfig.PotentialSpamFilters -join '; ' } else { $null })
            TransportRuleIndicators    = $(if ($spamConfig.TransportRuleIndicators) { $spamConfig.TransportRuleIndicators -join '; ' } else { $null })
        }
    }

    if ($TenantStatsHash.ContainsKey('SMTPRelayConfig') -and $TenantStatsHash['SMTPRelayConfig'].ContainsKey('Configuration')) {
        $smtpConfig = $TenantStatsHash['SMTPRelayConfig']['Configuration']
        $TenantStatsHash['SMTPRelaySummary']['Summary'] = [PSCustomObject]@{
            SMTPAuthEnabled                  = $smtpConfig.SMTPAuthEnabled
            SMTPAuthUsers                    = $smtpConfig.SMTPAuthUsers
            ConnectorBasedRelay              = $smtpConfig.ConnectorBasedRelay
            RelayConnectorCount              = @($smtpConfig.RelayConnectors).Count
            DirectSendEnabled                = $smtpConfig.DirectSendEnabled
            AuthoritativeDomains             = $(if ($smtpConfig.PSObject.Properties['AuthoritativeDomains']) { $smtpConfig.AuthoritativeDomains } else { $null })
            SmtpClientAuthenticationDisabled = $(if ($smtpConfig.PSObject.Properties['SmtpClientAuthenticationDisabled']) { $smtpConfig.SmtpClientAuthenticationDisabled } else { $null })
        }
    }

    if ($TenantStatsHash.ContainsKey('FederationConfiguration') -and $TenantStatsHash['FederationConfiguration']) {
        $fedConfig = $TenantStatsHash['FederationConfiguration']
        $exchangeFed = if ($fedConfig.ContainsKey('ExchangeFederation')) { $fedConfig['ExchangeFederation'] } else { $null }
        $crossTenant = if ($fedConfig.ContainsKey('CrossTenantAccess')) { $fedConfig['CrossTenantAccess'] } else { $null }
        $externalIds = if ($fedConfig.ContainsKey('ExternalIdentities')) { $fedConfig['ExternalIdentities'] } else { $null }

        $TenantStatsHash['FederationSummary']['Summary'] = [PSCustomObject]@{
            OrganizationRelationshipCount = $(if ($exchangeFed) { $exchangeFed.OrganizationRelationshipCount } else { 0 })
            IntraOrgConnectorCount        = $(if ($exchangeFed) { $exchangeFed.IntraOrgConnectorCount } else { 0 })
            CrossTenantPartnerCount       = $(if ($crossTenant) { $crossTenant.PartnerCount } else { 0 })
            CrossTenantPartners           = $(if ($crossTenant) { $crossTenant.PartnerTenantNames } else { $null })
            B2BManagementPolicyPresent    = $(if ($externalIds) { $externalIds.B2BManagementPolicyPresent } else { $null })
            GuestUserRole                 = $(if ($externalIds) { $externalIds.GuestUserRole } else { $null })
            InvitationsAllowed            = $(if ($externalIds) { $externalIds.InvitationsAllowed } else { $null })
        }
    }

    if ($TenantStatsHash.ContainsKey('TeamsVoice') -and $TenantStatsHash['TeamsVoice'].ContainsKey('Summary')) {
        $TenantStatsHash['TeamsVoiceSummary']['Summary'] = $TenantStatsHash['TeamsVoice']['Summary']
    }
}

function Update-LicenseClassificationMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash
    )

    if (-not $TenantStatsHash.ContainsKey('LicenseSKUs') -or -not $TenantStatsHash['LicenseSKUs']) {
        return
    }

    $userCount = 0
    if ($TenantStatsHash.ContainsKey('Users') -and $TenantStatsHash['Users']) {
        $userCount = @($TenantStatsHash['Users'].Values).Count
    }

    foreach ($licenseKey in @($TenantStatsHash['LicenseSKUs'].Keys)) {
        $license = $TenantStatsHash['LicenseSKUs'][$licenseKey]
        if (-not $license) { continue }

        $classification = Get-LicenseClassification -License $license -UserCount $userCount
        $license | Add-Member -MemberType NoteProperty -Name IsPaid -Value $classification.IsPaid -Force
        $license | Add-Member -MemberType NoteProperty -Name LicenseClass -Value $classification.LicenseClass -Force
        $license | Add-Member -MemberType NoteProperty -Name LicenseClassificationReason -Value $classification.Reason -Force
    }
}

#endregion

#region HTML Generation Functions

function Get-HtmlStyle {
    <#
    .SYNOPSIS
        Returns complete CSS styling for the report.
    #>
    return @'
<style>
    :root {
        --primary-color: #0f4c5c;
        --success-color: #107c10;
        --warning-color: #c77d2b;
        --danger-color: #d13438;
        --info-color: #00bcf2;
        --bg-light: #faf9f8;
        --bg-white: #ffffff;
        --text-primary: #323130;
        --text-secondary: #605e5c;
        --border-color: #edebe9;
        --shadow: 0 2px 4px rgba(0,0,0,0.1);
        --anchor-offset: 120px;
    }
    
    * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
    }
    
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        line-height: 1.6;
        color: var(--text-primary);
        background: var(--bg-light);
        padding: 0;
        margin: 0;
    }
    
    .container {
        max-width: 1400px;
        margin: 0 auto;
        padding: 20px;
    }
    
    /* Header */
    .report-header {
        background: linear-gradient(145deg, #12343b 0%, #1e5160 55%, #c77d2b 140%);
        color: white;
        padding: 30px;
        border-radius: 8px;
        margin-bottom: 30px;
        box-shadow: var(--shadow);
    }

    .report-brand {
        font-size: 0.78em;
        text-transform: uppercase;
        letter-spacing: 0.16em;
        font-weight: 700;
        opacity: 0.88;
        margin-bottom: 10px;
    }
    
    .report-header h1 {
        font-size: 2em;
        margin-bottom: 10px;
        font-weight: 600;
    }
    
    .report-meta {
        display: flex;
        gap: 20px;
        flex-wrap: wrap;
        margin-top: 15px;
        font-size: 0.9em;
        opacity: 0.95;
    }
    
    .report-meta-item {
        display: flex;
        align-items: center;
        gap: 5px;
    }
    
    /* Navigation */
    .nav-container {
        position: sticky;
        top: 0;
        background: var(--bg-white);
        border-bottom: 2px solid var(--border-color);
        z-index: 100;
        margin: 0 -20px 30px -20px;
        padding: 0 20px 10px 20px;
    }
    
    .nav {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        overflow-x: auto;
        padding: 10px 0 6px 0;
    }

    .nav-link {
        padding: 8px 16px;
        text-decoration: none;
        color: var(--text-primary);
        border-radius: 4px;
        white-space: nowrap;
        transition: all 0.2s;
        font-size: 0.9em;
        background: #f8f9fb;
        border: 1px solid #e5e7eb;
    }

    .nav-link-primary {
        background: #e8f3ff;
        border-color: #b9dbff;
        font-weight: 600;
    }

    .nav-link:hover {
        background: #eef6ff;
        color: var(--primary-color);
    }

    .nav-group {
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 6px;
        padding: 6px 8px;
        border: 1px solid var(--border-color);
        border-radius: 6px;
        background: #ffffff;
    }

    .nav-group-label {
        font-size: 0.72em;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: var(--text-secondary);
        font-weight: 600;
        padding: 0 4px 0 2px;
    }
    
    /* KPI Cards */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 20px;
        margin-bottom: 30px;
    }
    
    .kpi-card {
        background: var(--bg-white);
        padding: 20px;
        border-radius: 8px;
        box-shadow: var(--shadow);
        border-left: 4px solid var(--primary-color);
    }
    
    .kpi-card.kpi-success { border-left-color: var(--success-color); }
    .kpi-card.kpi-warning { border-left-color: var(--warning-color); }
    .kpi-card.kpi-danger { border-left-color: var(--danger-color); }
    
    .kpi-title {
        font-size: 0.85em;
        color: var(--text-secondary);
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .kpi-value {
        font-size: clamp(1.3rem, 2vw, 2em);
        font-weight: 600;
        color: var(--text-primary);
        margin-bottom: 5px;
        line-height: 1.25;
        overflow-wrap: anywhere;
        font-variant-numeric: tabular-nums;
    }
    
    .kpi-subtitle {
        font-size: 0.85em;
        color: var(--text-secondary);
    }
    
    /* Findings Panel */
    .findings-panel {
        background: var(--bg-white);
        padding: 25px;
        border-radius: 8px;
        box-shadow: var(--shadow);
        margin-bottom: 30px;
    }
    
    .findings-panel h2 {
        font-size: 1.3em;
        margin-bottom: 20px;
        color: var(--text-primary);
    }
    
    .findings-list {
        list-style: none;
    }
    
    .findings-list li {
        padding: 12px 15px;
        margin-bottom: 10px;
        border-radius: 4px;
        border-left: 4px solid;
        background: var(--bg-light);
    }
    
    .findings-list li.finding-risk {
        border-left-color: var(--danger-color);
        background: #fff4f4;
    }
    
    .findings-list li.finding-warning {
        border-left-color: var(--warning-color);
        background: #fffaf4;
    }
    
    .findings-list li.finding-info {
        border-left-color: var(--info-color);
        background: #f4fcff;
    }
    
    .findings-list li a {
        color: inherit;
        text-decoration: none;
        font-weight: 500;
    }
    
    .findings-list li a:hover {
        text-decoration: underline;
    }
    
    /* Sections */
    .section {
        background: var(--bg-white);
        padding: 30px;
        border-radius: 8px;
        box-shadow: var(--shadow);
        margin-bottom: 30px;
        scroll-margin-top: var(--anchor-offset);
    }

    #highlights {
        scroll-margin-top: var(--anchor-offset);
    }
    
    .section h2 {
        font-size: 1.5em;
        color: var(--text-primary);
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 2px solid var(--border-color);
    }

    .section-workload {
        display: inline-block;
        margin-bottom: 10px;
        padding: 4px 10px;
        font-size: 0.74em;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        border-radius: 999px;
        background: #edf7ff;
        color: #0f5a99;
        border: 1px solid #c7e4ff;
        font-weight: 700;
    }
    
    .section-footer {
        margin-top: 25px;
        padding-top: 20px;
        border-top: 1px solid var(--border-color);
    }
    
    .section-footer h3 {
        font-size: 1em;
        color: var(--text-primary);
        margin-bottom: 10px;
    }
    
    .section-footer p {
        color: var(--text-secondary);
        font-size: 0.9em;
        line-height: 1.6;
    }
    
    /* Tables */
    .data-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        font-size: 0.9em;
    }
    
    .data-table thead {
        position: sticky;
        top: 50px;
        background: var(--bg-light);
        z-index: 10;
    }
    
    .data-table th {
        padding: 12px;
        text-align: left;
        font-weight: 600;
        color: var(--text-primary);
        border-bottom: 2px solid var(--border-color);
        white-space: nowrap;
    }
    
    .data-table td {
        padding: 10px 12px;
        border-bottom: 1px solid var(--border-color);
    }

    .wrap-headers th {
        white-space: normal;
        word-break: break-word;
    }

    .wrap-cells td {
        white-space: normal;
        word-break: break-word;
    }
    
    .data-table tbody tr:nth-child(even) {
        background: var(--bg-light);
    }
    
    .data-table tbody tr:hover {
        background: #e1f5fe;
        cursor: pointer;
    }
    
    .data-table td.risk-cell {
        background: #fff4f4;
        color: var(--danger-color);
        font-weight: 600;
    }
    
    /* Badges */
    .badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 0.85em;
        font-weight: 500;
        white-space: nowrap;
    }
    
    .badge-info {
        background: #e1f5fe;
        color: #0277bd;
    }
    
    .badge-warning {
        background: #fff3e0;
        color: #e65100;
    }
    
    .badge-risk {
        background: #ffebee;
        color: #c62828;
    }
    
    .badge-success {
        background: #e8f5e9;
        color: #2e7d32;
    }
    
    /* Callouts */
    .callout {
        padding: 20px;
        border-radius: 6px;
        margin: 20px 0;
        border-left: 4px solid;
    }
    
    .callout-info {
        background: #f4fcff;
        border-left-color: var(--info-color);
    }
    
    .callout-warning {
        background: #fffaf4;
        border-left-color: var(--warning-color);
    }
    
    .callout-danger {
        background: #fff4f4;
        border-left-color: var(--danger-color);
    }
    
    .callout-success {
        background: #f4fff4;
        border-left-color: var(--success-color);
    }
    
    .callout-header {
        font-weight: 600;
        font-size: 1.1em;
        margin-bottom: 10px;
    }
    
    .callout-body {
        color: var(--text-secondary);
    }
    
    /* Charts */
    .chart-container {
        position: relative;
        margin: 30px 0;
        padding: 20px;
        background: var(--bg-light);
        border-radius: 6px;
    }
    
    .chart-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
        gap: 30px;
        margin: 20px 0;
    }
    
    /* Empty State */
    .empty-state {
        text-align: center;
        padding: 60px 20px;
        color: var(--text-secondary);
        font-style: italic;
    }
    
    /* Responsive */
    @media (max-width: 768px) {
        .container {
            padding: 10px;
        }
        
        .kpi-grid {
            grid-template-columns: 1fr;
        }
        
        .report-header h1 {
            font-size: 1.5em;
        }
        
        .nav {
            flex-direction: column;
        }

        .nav-group {
            width: 100%;
        }
        
        .data-table {
            font-size: 0.8em;
        }
        
        .data-table th,
        .data-table td {
            padding: 8px;
        }
    }
    
    /* Print Styles */
    @media print {
        .nav-container,
        .chart-container {
            display: none;
        }
        
        .section {
            page-break-inside: avoid;
        }
    }
</style>
'@
}

function Get-HtmlScript {
    <#
    .SYNOPSIS
        Returns JavaScript for charts and interactivity - MUST BE PLACED AFTER CHART.JS CDN
    #>
    return @'
<script>
    // Chart color palette
    const colors = {
        primary: '#0078d4',
        success: '#107c10',
        warning: '#f7630c',
        danger: '#d13438',
        info: '#00bcf2',
        palette: [
            '#0078d4', '#107c10', '#f7630c', '#d13438', '#00bcf2',
            '#8764b8', '#00b7c3', '#bad80a', '#ff8c00', '#e3008c'
        ]
    };
    
    // Default chart options
    Chart.defaults.font.family = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
    Chart.defaults.plugins.legend.position = 'bottom';
    Chart.defaults.plugins.legend.labels.padding = 15;
    Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(0,0,0,0.8)';
    Chart.defaults.plugins.tooltip.padding = 12;
    Chart.defaults.plugins.tooltip.cornerRadius = 4;
    
    // FUNCTION 1: Create pie chart
    function createPieChart(canvasId, data, labels, title) {
        const ctx = document.getElementById(canvasId);
        if (!ctx) {
            console.error('Canvas not found:', canvasId);
            return;
        }
        
        new Chart(ctx, {
            type: 'pie',
            data: {
                labels: labels,
                datasets: [{
                    data: data,
                    backgroundColor: colors.palette,
                    borderWidth: 2,
                    borderColor: '#ffffff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    title: {
                        display: true,
                        text: title,
                        font: { size: 16, weight: 'bold' },
                        padding: 20
                    },
                    legend: {
                        display: true,
                        position: 'bottom'
                    }
                }
            }
        });
    }
    
    // FUNCTION 2: Create bar chart
    function createBarChart(canvasId, data, labels, title) {
        const ctx = document.getElementById(canvasId);
        if (!ctx) {
            console.error('Canvas not found:', canvasId);
            return;
        }
        
        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: title,
                    data: data,
                    backgroundColor: colors.primary,
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    title: {
                        display: true,
                        text: title,
                        font: { size: 16, weight: 'bold' },
                        padding: 20
                    },
                    legend: {
                        display: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: 'rgba(0,0,0,0.05)'
                        }
                    },
                    x: {
                        grid: {
                            display: false
                        }
                    }
                }
            }
        });
    }
    
    // Smooth scroll for anchor links with sticky-nav offset awareness
    document.addEventListener('DOMContentLoaded', function() {
        function updateAnchorOffset() {
            const navContainer = document.querySelector('.nav-container');
            const navHeight = navContainer ? Math.ceil(navContainer.getBoundingClientRect().height) : 0;
            const offset = Math.max(navHeight + 12, 80);
            document.documentElement.style.setProperty('--anchor-offset', `${offset}px`);
            return offset;
        }

        function getAnchorOffset() {
            const cssValue = getComputedStyle(document.documentElement).getPropertyValue('--anchor-offset').trim();
            const parsed = parseInt(cssValue.replace('px', ''), 10);
            return Number.isFinite(parsed) ? parsed : 120;
        }

        function scrollToHashTarget(hash, behavior) {
            if (!hash || hash.length < 2) {
                return;
            }
            const target = document.querySelector(hash);
            if (!target) {
                return;
            }

            const offset = getAnchorOffset();
            const targetTop = target.getBoundingClientRect().top + window.pageYOffset - offset;
            window.scrollTo({
                top: Math.max(targetTop, 0),
                behavior: behavior || 'smooth'
            });
        }

        updateAnchorOffset();

        document.querySelectorAll('a[href^="#"]').forEach(anchor => {
            anchor.addEventListener('click', function(e) {
                const hash = this.getAttribute('href');
                if (!hash || hash === '#') {
                    return;
                }
                e.preventDefault();
                updateAnchorOffset();
                scrollToHashTarget(hash, 'smooth');
                if (window.history && window.history.pushState) {
                    window.history.pushState(null, '', hash);
                }
            });
        });

        window.addEventListener('resize', function() {
            updateAnchorOffset();
        });

        if (window.location.hash) {
            setTimeout(function() {
                updateAnchorOffset();
                scrollToHashTarget(window.location.hash, 'auto');
            }, 0);
        }
    });
</script>
'@
}

#region Section Builders
# Your proven function (included for reference)
function Convert-ArrayToPieChart {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Array,
        
        [Parameter(Mandatory=$true)]
        [string]$LabelProperty,

        [Parameter(Mandatory=$true)]
        [string]$ValueProperty,

        [string]$ChartTitle = "Pie Chart",
        
        [int]$Width = 300,
        [int]$Height = 300
    )

    $labels = @()
    $data = @()
    
    foreach ($item in $Array) {
        $labels += $item.$LabelProperty
        $data += $item.$ValueProperty
    }

    $labelsJSON = $labels | ConvertTo-Json -Compress
    $dataJSON = $data | ConvertTo-Json -Compress
    $safeTitle = if ([string]::IsNullOrWhiteSpace($ChartTitle)) { "PieChart" } else { ($ChartTitle -replace '[^a-zA-Z0-9_-]', '-') }
    $chartId = "myPieChart-$safeTitle-$([guid]::NewGuid().ToString('N').Substring(0,8))"

    $html = @"
<canvas id="$chartId" style="width:100%; height:${Height}px;"></canvas>
<script>
    var ctx = document.getElementById('$chartId').getContext('2d');
    var myPieChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: $labelsJSON,
            datasets: [{
                data: $dataJSON,
                backgroundColor: [
                    'rgba(255, 99, 132, 0.2)',
                    'rgba(54, 162, 235, 0.2)',
                    'rgba(255, 206, 86, 0.2)',
                    'rgba(75, 192, 192, 0.2)',
                    'rgba(153, 102, 255, 0.2)',
                    'rgba(255, 159, 64, 0.2)'
                ],
                borderColor: [
                    'rgba(255, 99, 132, 1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(255, 206, 86, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(153, 102, 255, 1)',
                    'rgba(255, 159, 64, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                },
                title: {
                    display: true,
                    text: '$ChartTitle'
                }
            }
        },
    });
</script>
"@

    return $html
}

# Updated Build-LicenseSection using your function
function Build-LicenseSection {
    param(
        [array]$Licenses,
        [int]$UserCount = 0
    )
    
    if ($Licenses.Count -eq 0) {
        return "<div class='empty-state'>No license data available</div>"
    }
    
    #Write-Host "  Building License section..." -ForegroundColor Gray
    
    # Get analysis with processed licenses
    $analysis = Get-LicenseAnalysis -Licenses $Licenses -UserCount $UserCount
    $paidLicenses = $analysis.PaidLicenses
    
    # Calculate totals
    $totalPurchased = $analysis.TotalPurchased
    $totalConsumed = $analysis.TotalConsumed
    $totalRemaining = $totalPurchased - $totalConsumed
    
    # Build summary KPIs
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Paid Licenses" -Value (Format-AssessmentHtmlNumber $totalPurchased) -Subtitle "Used for utilization only" -Theme 'default'
    $kpiHtml += New-KpiCard -Title "Consumed" -Value (Format-AssessmentHtmlNumber $totalConsumed) -Theme 'default'
    $kpiHtml += New-KpiCard -Title "Available" -Value (Format-AssessmentHtmlNumber $totalRemaining) -Theme 'default'
    
    $utilPct = if ($totalPurchased -gt 0) {
        ($totalConsumed / $totalPurchased) * 100
    } else { 0 }
    $utilTheme = if ($utilPct -ge 85) { 'warning' } elseif ($utilPct -ge 70) { 'warning' } else { 'success' }
    $kpiHtml += New-KpiCard -Title "Utilization" -Value (Format-AssessmentHtmlPercentage ($utilPct/100) -DecimalPlaces 1) -Theme $utilTheme
    $kpiHtml += "</div>"
    
    # Updated summary message with new categories
    $overCapacityCount = ($paidLicenses | Where-Object { $_.RemainingUnits -lt 0 }).Count
    $atCapacityCount = ($paidLicenses | Where-Object { $_.RemainingUnits -eq 0 }).Count
    $highUtilCount = ($paidLicenses | Where-Object { $_.Utilization -ge 85 -and $_.RemainingUnits -gt 0 }).Count
    
    if ($overCapacityCount -gt 0 -or $atCapacityCount -gt 0 -or $highUtilCount -gt 0) {
        $summaryHtml = "<div style='margin: 20px 0; padding: 15px; background: "
        $summaryHtml += if ($overCapacityCount -gt 0) { '#fff4f4' } elseif ($atCapacityCount -gt 0) { '#fffaf4' } else { '#fff8e1' }
        $summaryHtml += "; border-left: 4px solid "
        $summaryHtml += if ($overCapacityCount -gt 0) { '#d13438' } elseif ($atCapacityCount -gt 0) { '#f7630c' } else { '#ffa726' }
        $summaryHtml += "; border-radius: 4px;'>"
        
        if ($overCapacityCount -gt 0) {
            $summaryHtml += "<strong>🔴 CRITICAL: $overCapacityCount license(s) OVER capacity</strong> - You are consuming more licenses than purchased!<br>"
        }
        if ($atCapacityCount -gt 0) {
            $summaryHtml += "<strong>⚠️ WARNING: $atCapacityCount license(s) at 0 remaining</strong> - Fully allocated, no licenses available for new users<br>"
        }
        if ($highUtilCount -gt 0) {
            $summaryHtml += "<strong>⚠️ WARNING: $highUtilCount license(s) with high utilization (&gt;85%)</strong> - Running low, plan purchases soon<br>"
        }
        $summaryHtml += "See <a href='#highlights'>Findings of Note</a> for details."
        $summaryHtml += "</div>"
        $kpiHtml += $summaryHtml
    }
    
    # Prepare table data
    $tableData = $paidLicenses | ForEach-Object {
        $licenseName = if ($_.PSObject.Properties['SkuFriendlyName'] -and $_.SkuFriendlyName) {
            $_.SkuFriendlyName
        } else {
            Get-FriendlyProductName -SkuPartNumber $_.SkuPartNumber
        }
        [PSCustomObject]@{
            LicenseName = $licenseName
            AppliesTo = 'User'
            PurchasedUnits = $_.PurchasedUnits
            ConsumedUnits = $_.ConsumedUnits
            RemainingUnits = $_.RemainingUnits
            Utilization = [math]::Round($_.Utilization, 1)
        }
    }
    
    # Build table
    $tableColumns = @('LicenseName', 'AppliesTo', 'PurchasedUnits', 'ConsumedUnits', 'RemainingUnits', 'Utilization')
    $tableHeaders = @{
        'LicenseName' = 'License Type'
        'AppliesTo' = 'Applies To'
        'PurchasedUnits' = 'Purchased'
        'ConsumedUnits' = 'Consumed'
        'RemainingUnits' = 'Remaining'
        'Utilization' = 'Utilization %'
    }
    
    # Updated risk highlighting - differentiate critical vs warning
    $riskColumns = @{
        'RemainingUnits' = { param($val, $row) 
            # Critical (red) if negative, Warning (orange) if 0
            [int]$val -lt 0
        }
        'Utilization' = { param($val, $row)
            [double]$val -ge 85
        }
    }
    
    # Sort by remaining (problems first)
    $sortedData = $tableData | Sort-Object RemainingUnits, @{Expression={$_.ConsumedUnits}; Descending=$true}
    
    $tableHtml = New-HtmlTable -Data $sortedData -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns
    
    # Add custom CSS for warning-level highlighting (0 remaining)
    $tableHtml = @"
<style>
    .warning-cell {
        background: #fffaf4 !important;
        color: #e65100 !important;
        font-weight: 600;
    }
</style>
<script>
    // Highlight cells with 0 remaining in orange (warning)
    document.addEventListener('DOMContentLoaded', function() {
        const table = document.querySelector('.data-table');
        if (table) {
            const rows = table.querySelectorAll('tbody tr');
            rows.forEach(row => {
                const cells = row.querySelectorAll('td');
                const remainingCell = cells[4]; // RemainingUnits column (0-indexed)
                if (remainingCell && remainingCell.textContent.trim() === '0') {
                    remainingCell.classList.add('warning-cell');
                }
            });
        }
    });
</script>
$tableHtml
"@
    
    # Add note about trial/viral
    $allProcessed = $Licenses | ForEach-Object {
        [PSCustomObject]@{
            SKU = if ($_.PSObject.Properties['SkuFriendlyName'] -and $_.SkuFriendlyName) {
                $_.SkuFriendlyName
            } else {
                Get-FriendlyProductName -SkuPartNumber $_.SkuPartNumber
            }
            IsPaid = (Test-IsPaidLicenseSku -License $_ -UserCount $UserCount)
        }
    }
    $trialViralCount = ($allProcessed | Where-Object { -not $_.IsPaid }).Count
    
    if ($trialViralCount -gt 0) {
        $tableHtml += @"
<div style='margin-top: 15px; padding: 10px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;'>
    <strong>ℹ️ Note:</strong> $trialViralCount free, trial, preview, viral, or benefit licenses exist but are <strong>excluded from this table and utilization calculations</strong>. Very large user-based seat pools that greatly exceed tenant user count are also treated as likely freemium/benefit inventory.
</div>
"@
    }
    
    # Charts
    $chartHtml = "<div class='chart-grid'>"
    
    $top10 = $paidLicenses | Where-Object { 
        $_.ConsumedUnits -gt 0 
    } | Sort-Object ConsumedUnits -Descending | Select-Object -First 10
    
    if ($top10.Count -gt 0) {
        $chartHtml += "<div class='chart-container'>"
        $chartHtml += Convert-ArrayToPieChart `
            -Array $top10 `
            -LabelProperty 'SkuFriendlyName' `
            -ValueProperty 'ConsumedUnits' `
            -ChartTitle 'Top 10 Consumed Licenses' `
            -Width 400 `
            -Height 300
        $chartHtml += "</div>"
    }
    
    if ($totalPurchased -gt 0 -and $totalRemaining -ge 0) {
        $utilizationData = @(
            [PSCustomObject]@{ Label = 'Consumed'; Value = $totalConsumed }
            [PSCustomObject]@{ Label = 'Available'; Value = $totalRemaining }
        )
        
        $chartHtml += "<div class='chart-container'>"
        $chartHtml += Convert-ArrayToPieChart `
            -Array $utilizationData `
            -LabelProperty 'Label' `
            -ValueProperty 'Value' `
            -ChartTitle 'License Utilization' `
            -Width 400 `
            -Height 300
        $chartHtml += "</div>"
    }
    
    $chartHtml += "</div>"
    
    # Footer
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>License inventory for <strong>paid licenses only</strong>. Severity levels:</p>
    <ul>
        <li><strong>🔴 Critical (red):</strong> Over capacity - consuming more than purchased (negative remaining)</li>
        <li><strong>⚠️ Warning (orange):</strong> At capacity - 0 remaining, no licenses for new users</li>
        <li><strong>⚠️ Warning (orange):</strong> High utilization - &gt;85% used but some available</li>
    </ul>
    <h3>Recommended Next Steps</h3>
    <p><strong>Critical:</strong> Purchase additional licenses IMMEDIATELY - you're over allocated.<br>
    <strong>At 0 Remaining:</strong> Purchase licenses soon - next user assignment will fail.<br>
    <strong>High Utilization:</strong> Plan purchases for next month.</p>
</div>
"@
    
    return $kpiHtml + $tableHtml + $chartHtml + $footerHtml
}

# Updated Build-RecipientsSection
function Build-RecipientsSection {
    param([array]$Recipients)
    
    if ($Recipients.Count -eq 0) {
        return "<div class='empty-state'>No recipient data available</div>"
    }
    
    # Group by type
    $grouped = $Recipients | Group-Object RecipientTypeDetails | 
        Select-Object @{N='Type';E={$_.Name}}, Count |
        Sort-Object Count -Descending
    
    $tableHtml = New-HtmlTable -Data $grouped -Columns @('Type','Count') -ColumnHeaders @{'Type'='Recipient Type';'Count'='Quantity'}
    
    # Chart using your function
    $chartHtml = ""
    if ($grouped.Count -gt 0) {
        $chartHtml = "<div class='chart-container'>"
        $chartHtml += Convert-ArrayToPieChart `
            -Array $grouped `
            -LabelProperty 'Type' `
            -ValueProperty 'Count' `
            -ChartTitle 'Recipients by Type' `
            -Width 500 `
            -Height 400
        $chartHtml += "</div>"
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows the distribution of all mail-enabled recipient objects in your tenant, including mailboxes, groups, contacts, and resources.</p>
</div>
"@
    
    return $tableHtml + $chartHtml + $footerHtml
}

# Helper for isolated mailbox overview testing without a full assessment run.
function New-SampleMailboxOverviewData {
    [CmdletBinding()]
    param()

    $mailboxes = @(
        [PSCustomObject]@{
            ExternalDirectoryObjectId     = '11111111-1111-1111-1111-111111111111'
            DisplayName                   = 'Adele Vance'
            Identity                      = 'Adele Vance'
            RecipientTypeDetails          = 'UserMailbox'
            PrimarySmtpAddress            = 'adele.vance@contoso.com'
            EmailAddresses                = @('SMTP:adele.vance@contoso.com', 'smtp:adele@contoso.com')
            HiddenFromAddressListsEnabled = $false
            AddressBookPolicy             = $null
            ManagedBy                     = $null
            SKUAssigned                   = $true
            WhenCreated                   = [datetime]'2022-10-11T10:31:39'
            WhenSoftDeleted               = $null
            Guid                          = 'aaaaaaaa-1111-2222-3333-bbbbbbbbbbbb'
            Alias                         = 'adele.vance'
            Notes                         = $null
            DriveURL                      = 'https://contoso-my.sharepoint.com/personal/adele_vance_contoso_com'
            DriveStorageGB                = 12.4
            MBXSizeGB                     = 0.034
            ArchiveSizeGB                 = 0.034
            MBXItemCount                  = 4935
            ArchiveItemCount              = 94
            IsInactiveMailbox             = $false
        }
        [PSCustomObject]@{
            ExternalDirectoryObjectId     = '22222222-2222-2222-2222-222222222222'
            DisplayName                   = 'Megan Bowen'
            Identity                      = 'Megan Bowen'
            RecipientTypeDetails          = 'UserMailbox'
            PrimarySmtpAddress            = 'megan.bowen@contoso.com'
            EmailAddresses                = @('SMTP:megan.bowen@contoso.com')
            HiddenFromAddressListsEnabled = $false
            AddressBookPolicy             = $null
            ManagedBy                     = $null
            SKUAssigned                   = $true
            WhenCreated                   = [datetime]'2021-06-02T08:14:00'
            WhenSoftDeleted               = $null
            Guid                          = 'cccccccc-1111-2222-3333-dddddddddddd'
            Alias                         = 'megan.bowen'
            Notes                         = $null
            DriveURL                      = 'https://contoso-my.sharepoint.com/personal/megan_bowen_contoso_com'
            DriveStorageGB                = 51.2
            MBXSizeGB                     = 24.781
            ArchiveSizeGB                 = 3.551
            MBXItemCount                  = 18245
            ArchiveItemCount              = 882
            IsInactiveMailbox             = $false
        }
        [PSCustomObject]@{
            ExternalDirectoryObjectId     = '33333333-3333-3333-3333-333333333333'
            DisplayName                   = 'Alex Wilber'
            Identity                      = 'Alex Wilber'
            RecipientTypeDetails          = 'SharedMailbox'
            PrimarySmtpAddress            = 'sales@contoso.com'
            EmailAddresses                = @('SMTP:sales@contoso.com')
            HiddenFromAddressListsEnabled = $false
            AddressBookPolicy             = $null
            ManagedBy                     = $null
            SKUAssigned                   = $false
            WhenCreated                   = [datetime]'2020-03-15T15:20:00'
            WhenSoftDeleted               = $null
            Guid                          = 'eeeeeeee-1111-2222-3333-ffffffffffff'
            Alias                         = 'sales'
            Notes                         = 'Shared Sales mailbox'
            DriveURL                      = $null
            DriveStorageGB                = 0
            MBXSizeGB                     = 63.442
            ArchiveSizeGB                 = 18.005
            MBXItemCount                  = 64120
            ArchiveItemCount              = 6421
            IsInactiveMailbox             = $false
        }
        [PSCustomObject]@{
            ExternalDirectoryObjectId     = '44444444-4444-4444-4444-444444444444'
            DisplayName                   = 'Old Project Mailbox'
            Identity                      = 'Old Project Mailbox'
            RecipientTypeDetails          = 'UserMailbox'
            PrimarySmtpAddress            = 'old.project@contoso.com'
            EmailAddresses                = @('SMTP:old.project@contoso.com')
            HiddenFromAddressListsEnabled = $true
            AddressBookPolicy             = $null
            ManagedBy                     = $null
            SKUAssigned                   = $false
            WhenCreated                   = [datetime]'2019-01-08T09:00:00'
            WhenSoftDeleted               = $null
            Guid                          = '12121212-3434-5656-7878-909090909090'
            Alias                         = 'old.project'
            Notes                         = 'Inactive test mailbox'
            DriveURL                      = $null
            DriveStorageGB                = 0
            MBXSizeGB                     = 8.117
            ArchiveSizeGB                 = 0.512
            MBXItemCount                  = 7211
            ArchiveItemCount              = 141
            IsInactiveMailbox             = $true
        }
    )

    $inactiveMailboxes = @($mailboxes | Where-Object { $_.IsInactiveMailbox -eq $true })

    $publicFolders = @(
        [PSCustomObject]@{
            MailEnabled   = $true
            HasSubfolders = $true
        }
    )

    $mailboxStatsSummary = [PSCustomObject]@{
        TotalMailboxes       = $mailboxes.Count
        ActiveMailboxes      = (@($mailboxes | Where-Object { $_.IsInactiveMailbox -ne $true })).Count
        GraphPopulated       = 2
        ExoFallbackPopulated = 2
        Populated            = $mailboxes.Count
        Missing              = 0
    }

    return [PSCustomObject]@{
        Mailboxes           = $mailboxes
        InactiveMailboxes   = $inactiveMailboxes
        PublicFolders       = $publicFolders
        MailboxStatsSummary = $mailboxStatsSummary
    }
}

function New-SampleMailboxOverviewHtml {
    [CmdletBinding()]
    param()

    $sampleData = New-SampleMailboxOverviewData
    return Build-MailboxesSection `
        -Mailboxes $sampleData.Mailboxes `
        -InactiveMailboxes $sampleData.InactiveMailboxes `
        -PublicFolders $sampleData.PublicFolders `
        -MailboxStatsSummary $sampleData.MailboxStatsSummary
}

# Updated Build-MailboxesSection
function Build-MailboxesSection {
    param(
        [array]$Mailboxes,
        [array]$InactiveMailboxes,
        [array]$PublicFolders,
        [AllowNull()]
        $MailboxStatsSummary
    )

    function Get-MailboxSummaryValue {
        param(
            [AllowNull()]
            $Record,
            [string]$Key,
            [double]$Default = 0
        )

        if ($null -eq $Record) { return $Default }

        $value = $null
        if ($Record -is [System.Collections.IDictionary] -and $Record.Contains($Key)) {
            $value = $Record[$Key]
        }
        elseif ($Record.PSObject -and $Record.PSObject.Properties[$Key]) {
            $value = $Record.PSObject.Properties[$Key].Value
        }

        if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
            return $Default
        }

        try {
            return [double]$value
        }
        catch {
            return $Default
        }
    }

    $mailboxRows = @()
    if ($null -ne $Mailboxes) {
        $mailboxRows = @($Mailboxes | Where-Object { $null -ne $_ })
    }

    $inactiveMailboxRows = @()
    if ($null -ne $InactiveMailboxes) {
        $inactiveMailboxRows = @($InactiveMailboxes | Where-Object { $null -ne $_ })
    }

    $publicFolderRows = @()
    if ($null -ne $PublicFolders) {
        $publicFolderRows = @($PublicFolders | Where-Object { $null -ne $_ })
    }

    $publicFolderCount = $publicFolderRows.Count
    $publicFolderMailEnabled = ($publicFolderRows | Where-Object { $_.MailEnabled -eq $true }).Count
    $publicFolderHasSubfolders = ($publicFolderRows | Where-Object { $_.HasSubfolders -eq $true }).Count

    $summaryTotalMailboxes = [int](Get-MailboxSummaryValue -Record $MailboxStatsSummary -Key 'TotalMailboxes' -Default 0)
    $summaryActiveMailboxes = [int](Get-MailboxSummaryValue -Record $MailboxStatsSummary -Key 'ActiveMailboxes' -Default 0)
    $summaryInactiveMailboxes = [Math]::Max(($summaryTotalMailboxes - $summaryActiveMailboxes), 0)
    $summaryGraphPopulated = [int](Get-MailboxSummaryValue -Record $MailboxStatsSummary -Key 'GraphPopulated' -Default 0)
    $summaryExoPopulated = [int](Get-MailboxSummaryValue -Record $MailboxStatsSummary -Key 'ExoFallbackPopulated' -Default 0)
    $summaryMissing = [int](Get-MailboxSummaryValue -Record $MailboxStatsSummary -Key 'Missing' -Default 0)

    if ($mailboxRows.Count -eq 0 -and $publicFolderCount -eq 0 -and $summaryTotalMailboxes -eq 0) {
        return "<div class='empty-state'>No mailbox data available</div>"
    }

    $hasDetailedMailboxRows = $mailboxRows.Count -gt 0

    # Group by type and calculate stats when detailed mailbox rows are available.
    $summary = @()
    if ($hasDetailedMailboxRows) {
        $summary = $mailboxRows | Group-Object RecipientTypeDetails | ForEach-Object {
            $mbxs = $_.Group

            $validMbxSizes = $mbxs | ForEach-Object {
                $size = $_.MBXSizeGB
                if ($size -ne 'N/A' -and $null -ne $size -and $size -ne '') {
                    try { [double]$size } catch { $null }
                }
            } | Where-Object { $_ -ne $null }

            $validArchiveSizes = $mbxs | ForEach-Object {
                $size = $_.ArchiveSizeGB
                if ($size -ne 'N/A' -and $null -ne $size -and $size -ne '') {
                    try { [double]$size } catch { $null }
                }
            } | Where-Object { $_ -ne $null }

            $totalSize = ($validMbxSizes | Measure-Object -Sum).Sum
            $totalArchive = ($validArchiveSizes | Measure-Object -Sum).Sum
            $avgSize = if ($validMbxSizes.Count -gt 0) { $totalSize / $validMbxSizes.Count } else { 0 }
            $avgArchive = if ($validArchiveSizes.Count -gt 0) { $totalArchive / $validArchiveSizes.Count } else { 0 }
            $over50 = ($validMbxSizes | Where-Object { $_ -gt 50 }).Count
            $largest = ($validMbxSizes | Measure-Object -Maximum).Maximum
            if ($null -eq $largest) { $largest = 0 }

            [PSCustomObject]@{
                Type = $_.Name
                Quantity = $mbxs.Count
                TotalMBXSizeGB = [math]::Round($totalSize, 2)
                TotalArchiveSizeGB = [math]::Round($totalArchive, 2)
                AverageMBXSizeGB = [math]::Round($avgSize, 2)
                AverageArchiveSizeGB = [math]::Round($avgArchive, 2)
                Over50GB = $over50
                LargestGB = [math]::Round($largest, 2)
            }
        } | Sort-Object Quantity -Descending
    }
    
    # KPIs
    $kpiHtml = "<div class='kpi-grid'>"
    $totalMbx = if ($hasDetailedMailboxRows) {
        [int](($summary | Measure-Object Quantity -Sum).Sum)
    }
    else {
        $summaryTotalMailboxes
    }
    $totalSize = if ($hasDetailedMailboxRows) { ($summary | Measure-Object TotalMBXSizeGB -Sum).Sum } else { $null }
    $totalArchive = if ($hasDetailedMailboxRows) { ($summary | Measure-Object TotalArchiveSizeGB -Sum).Sum } else { $null }
    $totalLarge = if ($hasDetailedMailboxRows) { ($summary | Measure-Object Over50GB -Sum).Sum } else { $null }

    $kpiHtml += New-KpiCard -Title "Total Mailboxes" -Value (Format-AssessmentHtmlNumber $totalMbx)
    $kpiHtml += New-KpiCard -Title "Total Storage" -Value $(if ($null -ne $totalSize) { Format-AssessmentHtmlDataSize $totalSize } else { 'Not Collected' })
    $kpiHtml += New-KpiCard -Title "Total Archive Storage" -Value $(if ($null -ne $totalArchive) { Format-AssessmentHtmlDataSize $totalArchive } else { 'Not Collected' })

    $avgMbxSize = if ($null -ne $totalSize -and $totalMbx -gt 0) { $totalSize / $totalMbx } else { $null }
    $kpiHtml += New-KpiCard -Title "Average Size" -Value $(if ($null -ne $avgMbxSize) { Format-AssessmentHtmlDataSize $avgMbxSize } else { 'Not Collected' })

    $largeTheme = if ($null -eq $totalLarge) { 'default' } elseif ($totalLarge -gt ($totalMbx * 0.1)) { 'warning' } else { 'success' }
    $kpiHtml += New-KpiCard -Title "Over 50 GB" -Value $(if ($null -ne $totalLarge) { Format-AssessmentHtmlNumber $totalLarge } else { 'Not Collected' }) -Theme $largeTheme

    $inactiveCount = if ($inactiveMailboxRows.Count -gt 0) { $inactiveMailboxRows.Count } elseif ($summaryTotalMailboxes -gt 0) { $summaryInactiveMailboxes } else { 0 }
    $inactiveSize = if ($inactiveMailboxRows.Count -gt 0) {
        ($inactiveMailboxRows | Measure-Object MBXSizeGB -Sum).Sum
    } else { $null }
    $kpiHtml += New-KpiCard -Title "Inactive Mailboxes" -Value (Format-AssessmentHtmlNumber $inactiveCount)
    $kpiHtml += New-KpiCard -Title "Inactive Storage" -Value $(if ($null -ne $inactiveSize) { Format-AssessmentHtmlDataSize $inactiveSize } else { 'Not Collected' })

    $publicFolderTheme = if ($publicFolderCount -gt 0) { 'warning' } else { 'success' }
    $kpiHtml += New-KpiCard -Title "Public Folders" -Value (Format-AssessmentHtmlNumber $publicFolderCount) -Theme $publicFolderTheme
    $kpiHtml += New-KpiCard -Title "Mail-Enabled Public Folders" -Value (Format-AssessmentHtmlNumber $publicFolderMailEnabled)
    $kpiHtml += "</div>"
    
    # Table
    $tableHtml = ""
    if ($hasDetailedMailboxRows) {
        $tableColumns = @('Type','Quantity','TotalMBXSizeGB','TotalArchiveSizeGB','AverageMBXSizeGB','AverageArchiveSizeGB','Over50GB','LargestGB')
        $tableHeaders = @{
            'Type' = 'Mailbox Type'
            'Quantity' = 'Count'
            'TotalMBXSizeGB' = 'Total Size (GB)'
            'TotalArchiveSizeGB' = 'Total Archive (GB)'
            'AverageMBXSizeGB' = 'Avg Size (GB)'
            'AverageArchiveSizeGB' = 'Avg Archive (GB)'
            'Over50GB' = 'Over 50 GB'
            'LargestGB' = 'Largest (GB)'
        }

        $riskColumns = @{
            'Over50GB' = { param($val) [int]$val -gt 0 }
            'LargestGB' = { param($val) [double]$val -gt 50 }
        }

        $tableHtml = New-HtmlTable -Data $summary -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns
    }
    elseif ($summaryTotalMailboxes -gt 0) {
        $coverageRows = @(
            [PSCustomObject]@{ Metric = 'Total mailboxes discovered'; Value = (Format-AssessmentHtmlNumber $summaryTotalMailboxes) }
            [PSCustomObject]@{ Metric = 'Active mailboxes'; Value = (Format-AssessmentHtmlNumber $summaryActiveMailboxes) }
            [PSCustomObject]@{ Metric = 'Inactive mailboxes (derived)'; Value = (Format-AssessmentHtmlNumber $summaryInactiveMailboxes) }
            [PSCustomObject]@{ Metric = 'Mailbox stats populated via Graph'; Value = (Format-AssessmentHtmlNumber $summaryGraphPopulated) }
            [PSCustomObject]@{ Metric = 'Mailbox stats populated via EXO fallback'; Value = (Format-AssessmentHtmlNumber $summaryExoPopulated) }
            [PSCustomObject]@{ Metric = 'Mailboxes still missing stats'; Value = (Format-AssessmentHtmlNumber $summaryMissing) }
        )
        $tableHtml = New-CalloutBox -Type 'info' -Title 'Mailbox Size Detail Not Collected' -Content 'This snapshot does not include per-mailbox size rows, so size and >50 GB breakdowns are unavailable. Run a profile that collects detailed mailbox records for full mailbox analytics.'
        $tableHtml += New-HtmlTable -Data $coverageRows -Columns @('Metric','Value') -ColumnHeaders @{ Metric = 'Collection Summary'; Value = 'Value' }
    }
    
    # Inactive mailbox summary table
    $inactiveSummaryHtml = ""
    if ($inactiveMailboxRows.Count -gt 0) {
        $inactiveSummary = $inactiveMailboxRows | Group-Object RecipientTypeDetails | ForEach-Object {
            $mbxs = $_.Group
            $totalSize = ($mbxs | Measure-Object MBXSizeGB -Sum).Sum
            $avgSize = if ($mbxs.Count -gt 0) { $totalSize / $mbxs.Count } else { 0 }
            $over50 = ($mbxs | Where-Object { [double]$_.MBXSizeGB -gt 50 }).Count
            
            [PSCustomObject]@{
                Type = $_.Name
                Quantity = $mbxs.Count
                TotalSizeGB = [math]::Round($totalSize, 2)
                AverageSizeGB = [math]::Round($avgSize, 2)
                Over50GB = $over50
            }
        }
        $inactiveTableColumns = @('Type','Quantity','TotalSizeGB','AverageSizeGB','Over50GB')
        $inactiveTableHeaders = @{
            'Type' = 'Type'
            'Quantity' = 'Count'
            'TotalSizeGB' = 'Total Size (GB)'
            'AverageSizeGB' = 'Avg Size (GB)'
            'Over50GB' = 'Over 50 GB'
        }
        $inactiveSummaryHtml = "<h3 style='margin-top:30px;'>Inactive Mailboxes Summary</h3>" +
            (New-HtmlTable -Data $inactiveSummary -Columns $inactiveTableColumns -ColumnHeaders $inactiveTableHeaders)
    }

    # Public folder summary table
    $publicFolderSummaryHtml = ""
    if ($publicFolderCount -gt 0) {
        $publicFolderSummary = @(
            [PSCustomObject]@{
                TotalPublicFolders = $publicFolderCount
                MailEnabled = $publicFolderMailEnabled
                WithSubfolders = $publicFolderHasSubfolders
            }
        )
        $publicFolderTableColumns = @('TotalPublicFolders','MailEnabled','WithSubfolders')
        $publicFolderTableHeaders = @{
            'TotalPublicFolders' = 'Total Public Folders'
            'MailEnabled' = 'Mail-Enabled'
            'WithSubfolders' = 'With Subfolders'
        }
        $publicFolderSummaryHtml = "<h3 style='margin-top:30px;'>Public Folders</h3>" +
            (New-HtmlTable -Data $publicFolderSummary -Columns $publicFolderTableColumns -ColumnHeaders $publicFolderTableHeaders)
    }
    
    # Chart using your function
    $chartHtml = ""
    $chartData = if ($hasDetailedMailboxRows) { $summary | Where-Object { $_.Quantity -gt 0 } } else { @() }
    if (@($chartData).Count -gt 0) {
        $chartHtml = "<div class='chart-container'>"
        $chartHtml += Convert-ArrayToPieChart `
            -Array $chartData `
            -LabelProperty 'Type' `
            -ValueProperty 'Quantity' `
            -ChartTitle 'Mailboxes by Type' `
            -Width 500 `
            -Height 400
        $chartHtml += "</div>"
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows mailbox distribution, archive usage, and inactive mailbox storage. Mailboxes over 50 GB may experience performance issues and should be archived or cleaned.</p>
    <h3>Recommended Next Steps</h3>
    <p>Enable archive mailboxes for users over 50 GB. Review retention policies. Consider implementing auto-expanding archives for heavy users.</p>
</div>
"@
    
    return $kpiHtml + $tableHtml + $inactiveSummaryHtml + $publicFolderSummaryHtml + $chartHtml + $footerHtml
}

function Build-EmailActivitySection {
    param(
        [array]$TopSenders,
        [array]$TopReceivers,
        [object]$EmailActivitySummary
    )

    $summaryRecord = $EmailActivitySummary
    if ($summaryRecord -is [array]) {
        $summaryRecord = @($summaryRecord | Select-Object -First 1)
        if ($summaryRecord.Count -gt 0) {
            $summaryRecord = $summaryRecord[0]
        }
        else {
            $summaryRecord = $null
        }
    }

    if ((@($TopSenders).Count -eq 0) -and (@($TopReceivers).Count -eq 0) -and (-not $summaryRecord)) {
        return "<div class='empty-state'>No email activity data available</div>"
    }

    $reportRows = 0
    $activeUsers = 0
    $totalSendCount = 0
    $totalReceiveCount = 0
    $reportPeriod = 'N/A'
    $reportRefreshDate = 'N/A'
    function Get-SummaryValue {
        param(
            [AllowNull()]
            $Record,
            [string]$Key,
            [AllowNull()]
            $Default = $null
        )

        if ($null -eq $Record) {
            return $Default
        }
        if ($Record -is [System.Collections.IDictionary] -and $Record.Contains($Key)) {
            return $Record[$Key]
        }
        if ($Record -is [System.Collections.Specialized.OrderedDictionary] -and $Record.Contains($Key)) {
            return $Record[$Key]
        }
        if ($Record.PSObject -and $Record.PSObject.Properties[$Key]) {
            return $Record.$Key
        }
        return $Default
    }

    if ($summaryRecord) {
        try { $reportRows = [int](Get-SummaryValue -Record $summaryRecord -Key 'ReportRows' -Default 0) } catch { $reportRows = 0 }
        try { $activeUsers = [int](Get-SummaryValue -Record $summaryRecord -Key 'ActiveUsers' -Default 0) } catch { $activeUsers = 0 }
        try { $totalSendCount = [int64](Get-SummaryValue -Record $summaryRecord -Key 'TotalSendCount' -Default 0) } catch { $totalSendCount = 0 }
        try { $totalReceiveCount = [int64](Get-SummaryValue -Record $summaryRecord -Key 'TotalReceiveCount' -Default 0) } catch { $totalReceiveCount = 0 }
        $periodValue = Get-SummaryValue -Record $summaryRecord -Key 'PeriodDuration'
        if (-not [string]::IsNullOrWhiteSpace([string]$periodValue)) { $reportPeriod = [string]$periodValue }
        $refreshValue = Get-SummaryValue -Record $summaryRecord -Key 'ReportRefreshDate'
        if (-not [string]::IsNullOrWhiteSpace([string]$refreshValue)) { $reportRefreshDate = [string]$refreshValue }
    }

    if ($reportRows -eq 0 -and ((@($TopSenders).Count -gt 0) -or (@($TopReceivers).Count -gt 0))) {
        $reportRows = [math]::Max(@($TopSenders).Count, @($TopReceivers).Count)
    }
    if ($activeUsers -eq 0 -and @($TopSenders).Count -gt 0) {
        $activeUsers = @($TopSenders | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.UserPrincipalName) } | Select-Object -ExpandProperty UserPrincipalName -Unique).Count
    }

    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title 'Report Rows' -Value (Format-AssessmentHtmlNumber $reportRows)
    $kpiHtml += New-KpiCard -Title 'Active Users' -Value (Format-AssessmentHtmlNumber $activeUsers)
    $kpiHtml += New-KpiCard -Title 'Total Sends' -Value (Format-AssessmentHtmlNumber $totalSendCount)
    $kpiHtml += New-KpiCard -Title 'Total Receives' -Value (Format-AssessmentHtmlNumber $totalReceiveCount)
    $kpiHtml += New-KpiCard -Title 'Report Period' -Value $reportPeriod
    $kpiHtml += New-KpiCard -Title 'Report Refresh' -Value $reportRefreshDate
    $kpiHtml += "</div>"

    $sendersColumns = @('Rank', 'DisplayName', 'UserPrincipalName', 'SendCount', 'ReceiveCount', 'LastActivityDate')
    $sendersHeaders = @{
        Rank              = 'Rank'
        DisplayName       = 'Display Name'
        UserPrincipalName = 'User Principal Name'
        SendCount         = 'Send Count'
        ReceiveCount      = 'Receive Count'
        LastActivityDate  = 'Last Activity'
    }
    $receiversColumns = @('Rank', 'DisplayName', 'UserPrincipalName', 'ReceiveCount', 'SendCount', 'LastActivityDate')
    $receiversHeaders = @{
        Rank              = 'Rank'
        DisplayName       = 'Display Name'
        UserPrincipalName = 'User Principal Name'
        ReceiveCount      = 'Receive Count'
        SendCount         = 'Send Count'
        LastActivityDate  = 'Last Activity'
    }

    $topSendersHtml = "<h3 style='margin-top:30px;'>Top Senders</h3>"
    if (@($TopSenders).Count -gt 0) {
        $topSendersHtml += New-HtmlTable -Data $TopSenders -Columns $sendersColumns -ColumnHeaders $sendersHeaders
    }
    else {
        $topSendersHtml += "<div class='empty-state'>No sender activity rows were returned.</div>"
    }

    $topReceiversHtml = "<h3 style='margin-top:30px;'>Top Receivers</h3>"
    if (@($TopReceivers).Count -gt 0) {
        $topReceiversHtml += New-HtmlTable -Data $TopReceivers -Columns $receiversColumns -ColumnHeaders $receiversHeaders
    }
    else {
        $topReceiversHtml += "<div class='empty-state'>No receiver activity rows were returned.</div>"
    }

    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This section summarizes Microsoft Graph email activity telemetry and highlights the highest-volume senders and receivers in the organization.</p>
</div>
"@

    return $kpiHtml + $topSendersHtml + $topReceiversHtml + $footerHtml
}

function Build-InactiveMailboxesSection {
    param([array]$InactiveMailboxes)
    
    if ($InactiveMailboxes.Count -eq 0) {
        $callout = New-CalloutBox -Type 'success' -Title 'No Inactive Mailboxes' -Content 'Your tenant has no inactive mailboxes, which is good for license optimization.'
        return $callout
    }
    
    # Summary stats
    $summary = $InactiveMailboxes | Group-Object RecipientTypeDetails | ForEach-Object {
        $mbxs = $_.Group
        $totalSize = ($mbxs | Measure-Object MBXSizeGB -Sum).Sum
        $avgSize = if ($mbxs.Count -gt 0) { $totalSize / $mbxs.Count } else { 0 }
        $over50 = ($mbxs | Where-Object { [double]$_.MBXSizeGB -gt 50 }).Count
        
        [PSCustomObject]@{
            Type = $_.Name
            Quantity = $mbxs.Count
            TotalSizeGB = [math]::Round($totalSize, 2)
            AverageSizeGB = [math]::Round($avgSize, 2)
            Over50GB = $over50
        }
    }
    
    # Warning callout
    $calloutHtml = New-CalloutBox -Type 'warning' -Title 'Inactive Mailboxes Detected' -Content "You have $($InactiveMailboxes.Count) inactive mailboxes. These do not consume licenses but do use storage. Consider if they need to be retained for compliance."
    
    # Table
    $tableColumns = @('Type','Quantity','TotalSizeGB','AverageSizeGB','Over50GB')
    $tableHeaders = @{
        'Type' = 'Type'
        'Quantity' = 'Count'
        'TotalSizeGB' = 'Total Size (GB)'
        'AverageSizeGB' = 'Avg Size (GB)'
        'Over50GB' = 'Over 50 GB'
    }
    
    $tableHtml = New-HtmlTable -Data $summary -Columns $tableColumns -ColumnHeaders $tableHeaders
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>Inactive mailboxes are former user mailboxes that have been preserved for compliance. They consume storage but not licenses.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review retention requirements. If mailboxes are no longer needed for legal hold or eDiscovery, consider permanent deletion to free storage.</p>
</div>
"@
    
    return $calloutHtml + $tableHtml + $footerHtml
}

function Build-SharePointOneDriveSection {
    param(
        [array]$SharePointSites,
        [array]$OneDriveSites
    )
    
    # Calculate summaries
    $summaries = @()
    
    # SharePoint (non-O365 connected)
    $spRegular = $SharePointSites | Where-Object { -not $_.IsOffice365GroupsConnected }
    if ($spRegular.Count -gt 0) {
        $totalGB = ($spRegular | Measure-Object StorageUsedGB -Sum).Sum
        $avgGB = $totalGB / $spRegular.Count
        $over1TB = ($spRegular | Where-Object { $_.StorageUsedGB -gt 1024 }).Count
        $largest = ($spRegular | Measure-Object StorageUsedGB -Maximum).Maximum
        
        $summaries += [PSCustomObject]@{
            Category = 'SharePoint Sites'
            Quantity = $spRegular.Count
            TotalStorageGB = [math]::Round($totalGB, 2)
            AverageStorageGB = [math]::Round($avgGB, 2)
            Over1TB = $over1TB
            LargestSizeGB = [math]::Round($largest, 2)
        }
    }
    
    # OneDrive
    if ($OneDriveSites.Count -gt 0) {
        $totalGB = ($OneDriveSites | Measure-Object StorageUsedGB -Sum).Sum
        $avgGB = $totalGB / $OneDriveSites.Count
        $over1TB = ($OneDriveSites | Where-Object { $_.StorageUsedGB -gt 1024 }).Count
        $largest = ($OneDriveSites | Measure-Object StorageUsedGB -Maximum).Maximum
        
        $summaries += [PSCustomObject]@{
            Category = 'OneDrive Sites'
            Quantity = $OneDriveSites.Count
            TotalStorageGB = [math]::Round($totalGB, 2)
            AverageStorageGB = [math]::Round($avgGB, 2)
            Over1TB = $over1TB
            LargestSizeGB = [math]::Round($largest, 2)
        }
    }
    
    # O365 Groups
    $spGroups = $SharePointSites | Where-Object { $_.IsOffice365GroupsConnected -and $_.Template -ne 'TEAMCHANNEL#0' }
    if ($spGroups.Count -gt 0) {
        $totalGB = ($spGroups | Measure-Object StorageUsedGB -Sum).Sum
        $avgGB = $totalGB / $spGroups.Count
        $over1TB = ($spGroups | Where-Object { $_.StorageUsedGB -gt 1024 }).Count
        $largest = ($spGroups | Measure-Object StorageUsedGB -Maximum).Maximum
        
        $summaries += [PSCustomObject]@{
            Category = 'Office 365 Groups'
            Quantity = $spGroups.Count
            TotalStorageGB = [math]::Round($totalGB, 2)
            AverageStorageGB = [math]::Round($avgGB, 2)
            Over1TB = $over1TB
            LargestSizeGB = [math]::Round($largest, 2)
        }
    }
    
    # Teams
    $teams = $SharePointSites | Where-Object { $_.Template -eq 'TEAMCHANNEL#0' }
    if ($teams.Count -gt 0) {
        $totalGB = ($teams | Measure-Object StorageUsedGB -Sum).Sum
        $avgGB = $totalGB / $teams.Count
        $over1TB = ($teams | Where-Object { $_.StorageUsedGB -gt 1024 }).Count
        $largest = ($teams | Measure-Object StorageUsedGB -Maximum).Maximum
        
        $summaries += [PSCustomObject]@{
            Category = 'Teams Sites'
            Quantity = $teams.Count
            TotalStorageGB = [math]::Round($totalGB, 2)
            AverageStorageGB = [math]::Round($avgGB, 2)
            Over1TB = $over1TB
            LargestSizeGB = [math]::Round($largest, 2)
        }
    }
    
    if ($summaries.Count -eq 0) {
        return "<div class='empty-state'>No SharePoint or OneDrive data available</div>"
    }
    
    # KPIs
    $kpiHtml = "<div class='kpi-grid'>"
    $totalSites = ($summaries | Measure-Object Quantity -Sum).Sum
    $totalStorage = ($summaries | Measure-Object TotalStorageGB -Sum).Sum
    
    $kpiHtml += New-KpiCard -Title "Total Sites" -Value (Format-AssessmentHtmlNumber $totalSites)
    $kpiHtml += New-KpiCard -Title "Total Storage" -Value (Format-AssessmentHtmlDataSize $totalStorage)
    
    $avgStorage = if ($totalSites -gt 0) { $totalStorage / $totalSites } else { 0 }
    $kpiHtml += New-KpiCard -Title "Average Size" -Value (Format-AssessmentHtmlDataSize $avgStorage)
    $kpiHtml += "</div>"
    
    # Table
    $tableColumns = @('Category','Quantity','TotalStorageGB','AverageStorageGB','Over1TB','LargestSizeGB')
    $tableHeaders = @{
        'Category' = 'Site Type'
        'Quantity' = 'Count'
        'TotalStorageGB' = 'Total Size'
        'AverageStorageGB' = 'Avg Size'
        'Over1TB' = 'Over 1 TB'
        'LargestSizeGB' = 'Largest Site'
    }
    
    $riskColumns = @{
        'Over1TB' = { param($val) [int]$val -gt 0 }
    }

    $valueFormatters = @{
        'Quantity' = { param($val) Format-AssessmentHtmlNumber $val }
        'TotalStorageGB' = { param($val) Format-AssessmentHtmlDataSize $val }
        'AverageStorageGB' = { param($val) Format-AssessmentHtmlDataSize $val }
        'Over1TB' = { param($val) Format-AssessmentHtmlNumber $val }
        'LargestSizeGB' = { param($val) Format-AssessmentHtmlDataSize $val }
    }
    
    $tableHtml = New-HtmlTable -Data $summaries -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns -ValueFormatters $valueFormatters
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows storage consumption across SharePoint, OneDrive, and Teams. Sites over 1 TB may have performance impacts.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review large sites for archived content that can be moved to cold storage. Implement retention policies and educate users on storage best practices.</p>
</div>
"@
    
    return $kpiHtml + $tableHtml + $footerHtml
}

function Build-OwnershipGovernanceSection {
    param(
        [array]$UnmanagedObjects,
        [array]$OneDriveOwnerMismatches,
        [object]$OwnershipGovernanceSummary
    )

    $unmanagedCount = @($UnmanagedObjects).Count
    $mismatchCount = @($OneDriveOwnerMismatches).Count

    if ($unmanagedCount -eq 0 -and $mismatchCount -eq 0 -and -not $OwnershipGovernanceSummary) {
        return "<div class='empty-state'>No ownership governance data available</div>"
    }

    $missingOwnerCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['MissingOwnerCount']) {
        [int]$OwnershipGovernanceSummary.MissingOwnerCount
    } else {
        @($UnmanagedObjects | Where-Object { $_.OwnerState -eq 'Missing' }).Count
    }
    $ownerHealthRiskCount = if ($OwnershipGovernanceSummary -and $OwnershipGovernanceSummary.PSObject.Properties['OwnerHealthRiskCount']) {
        [int]$OwnershipGovernanceSummary.OwnerHealthRiskCount
    } else {
        @($UnmanagedObjects | Where-Object { $_.OwnerState -in @('Disabled', 'Stale', 'Mixed') }).Count
    }

    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Unmanaged Objects" -Value (Format-AssessmentHtmlNumber $unmanagedCount) -Theme $(if ($unmanagedCount -gt 0) { 'warning' } else { 'success' })
    $kpiHtml += New-KpiCard -Title "Missing Owner" -Value (Format-AssessmentHtmlNumber $missingOwnerCount) -Theme $(if ($missingOwnerCount -gt 0) { 'danger' } else { 'success' })
    $kpiHtml += New-KpiCard -Title "Owner Health Risks" -Value (Format-AssessmentHtmlNumber $ownerHealthRiskCount) -Theme $(if ($ownerHealthRiskCount -gt 0) { 'warning' } else { 'success' })
    $kpiHtml += New-KpiCard -Title "OneDrive Mismatches" -Value (Format-AssessmentHtmlNumber $mismatchCount) -Theme $(if ($mismatchCount -gt 0) { 'warning' } else { 'success' })
    $kpiHtml += "</div>"

    $topUnmanaged = @(
        $UnmanagedObjects |
            Sort-Object @{ Expression = {
                switch ($_.Severity) {
                    'Risk' { 1 }
                    'Warning' { 2 }
                    default { 3 }
                }
            } }, Workload, DisplayName |
            Select-Object -First 12 Workload, ObjectType, DisplayName, CurrentOwner, OwnerState, Reason
    )
    $unmanagedHtml = "<h3>Unmanaged Object Review</h3>" +
        (New-HtmlTable -Data $topUnmanaged -Columns @('Workload','ObjectType','DisplayName','CurrentOwner','OwnerState','Reason') -ColumnHeaders @{
            Workload='Workload'; ObjectType='Object Type'; DisplayName='Object'; CurrentOwner='Current Owner'; OwnerState='Owner State'; Reason='Finding'
        } -RiskColumns @{
            OwnerState = { param($val) [string]$val -in @('Missing','Disabled','Stale','Mixed') }
        })

    $topMismatches = @(
        $OneDriveOwnerMismatches |
            Sort-Object DisplayName, SiteUrl |
            Select-Object -First 12 DisplayName, SiteUrl, CurrentOwner, ExpectedDefaultOwner
    )
    $mismatchHtml = "<h3 style='margin-top:20px;'>OneDrive Owner Mismatch Review</h3>" +
        (New-HtmlTable -Data $topMismatches -Columns @('DisplayName','SiteUrl','CurrentOwner','ExpectedDefaultOwner') -ColumnHeaders @{
            DisplayName='OneDrive'; SiteUrl='Site Url'; CurrentOwner='Current Owner'; ExpectedDefaultOwner='Expected Default Owner'
        })

    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>Objects without a healthy owner increase governance risk and slow migration/workload handoff decisions.</p>
    <h3>Recommended Next Steps</h3>
    <p>Assign accountable owners, reassign from disabled or stale identities, and validate OneDrive owner mismatches with business data stewards.</p>
</div>
"@

    return $kpiHtml + $unmanagedHtml + $mismatchHtml + $footerHtml
}

function Build-TeamsSection {
    param(
        [array]$Teams,
        [array]$Licenses,
        [object]$TeamsVoice,
        [int]$UserCount = 0
    )
    
    if ($Teams.Count -eq 0) {
        return "<div class='empty-state'>No Teams data available</div>"
    }
    
    # Determine Teams Voice availability from service plans
    $voicePlans = @('MCOEV','MCOPSTN1','MCOPSTN2','MCOEV_VIRTUALUSER','MCOEV_DOD','MCOPSTNC')
    $hasVoice = $false
    $voicePlanHits = @()
    $paidLicenses = $Licenses | Where-Object { Test-IsPaidLicenseSku -License $_ -UserCount $UserCount }

    foreach ($lic in $paidLicenses) {
        $plans = @()
        if ($lic.PSObject.Properties['ServicePlans'] -and $lic.ServicePlans) {
            $plans = $lic.ServicePlans -split ','
        }
        foreach ($plan in $plans) {
            $trimmed = $plan.Trim()
            if ($voicePlans -contains $trimmed) {
                $hasVoice = $true
                $voicePlanHits += $trimmed
            }
        }
    }
    $voicePlanHits = $voicePlanHits | Select-Object -Unique
    
    $teamRows = $Teams | ForEach-Object {
        $publicCount = if ($_.PublicChannels) { ((@($_.PublicChannels -split ',')) | Where-Object { $_ -and $_.Trim() -ne '' }).Count } else { 0 }
        $privateCount = if ($_.PrivateChannels) { ((@($_.PrivateChannels -split ',')) | Where-Object { $_ -and $_.Trim() -ne '' }).Count } else { 0 }
        $sharedCount = if ($_.SharedChannels) { ((@($_.SharedChannels -split ',')) | Where-Object { $_ -and $_.Trim() -ne '' }).Count } else { 0 }
        [PSCustomObject]@{
            DisplayName = $_.DisplayName
            Visibility = $_.Visibility
            SiteSizeGB = $_.'SiteSize-GB'
            TotalChannels = $_.TotalChannels
            PublicChannels = $publicCount
            PrivateChannels = $privateCount
            SharedChannels = $sharedCount
        }
    }
    
    $totalTeams = $Teams.Count
    $totalChannels = ($Teams | Measure-Object -Property TotalChannels -Sum).Sum
    
    $voiceValue = if ($hasVoice) { 'Detected' } else { 'Not Detected' }
    $voiceTheme = if ($hasVoice) { 'success' } else { 'warning' }
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Total Teams" -Value (Format-AssessmentHtmlNumber $totalTeams)
    $kpiHtml += New-KpiCard -Title "Total Channels" -Value (Format-AssessmentHtmlNumber $totalChannels)
    $kpiHtml += New-KpiCard -Title "Teams Voice" -Value $voiceValue -Theme $voiceTheme
    $kpiHtml += "</div>"
    
    if ($voicePlanHits.Count -gt 0) {
        $kpiHtml += "<div style='margin-top:10px; color:#605e5c; font-size:0.9em;'>Voice-related plans detected: $([string]::Join(', ', $voicePlanHits))</div>"
    }

    $voiceSummaryTable = ""
    if ($TeamsVoice -and $TeamsVoice.Summary) {
        $voiceSummary = $TeamsVoice.Summary
        $voiceRows = @(
            [PSCustomObject]@{ Metric = 'PSTN Total Minutes (Last 30 Days)'; Value = $voiceSummary.PstnTotalMinutes },
            [PSCustomObject]@{ Metric = 'PSTN Total Calls (Last 30 Days)'; Value = $voiceSummary.PstnTotalCalls },
            [PSCustomObject]@{ Metric = 'Calling Policies'; Value = $voiceSummary.CallingPolicyCount },
            [PSCustomObject]@{ Metric = 'Assigned Phone Numbers'; Value = $voiceSummary.PhoneNumberCount },
            [PSCustomObject]@{ Metric = 'Voice-Enabled Users'; Value = $voiceSummary.VoiceUserCount }
        )
        if ($voiceSummary.PSObject.Properties['DataSource'] -and $voiceSummary.DataSource) {
            $voiceRows += [PSCustomObject]@{ Metric = 'Data Source'; Value = $voiceSummary.DataSource }
        }
        if ($voiceSummary.PSObject.Properties['Notes'] -and $voiceSummary.Notes) {
            $voiceRows += [PSCustomObject]@{ Metric = 'Notes'; Value = $voiceSummary.Notes }
        }
        $voiceSummaryTable = "<h3 style='margin-top:20px;'>Teams Voice Summary</h3>" +
            (New-HtmlTable -Data $voiceRows -Columns @('Metric','Value') -ColumnHeaders @{ Metric='Metric'; Value='Value' })
    }
    
    $tableColumns = @('DisplayName','Visibility','SiteSizeGB','TotalChannels','PublicChannels','PrivateChannels','SharedChannels')
    $tableHeaders = @{
        DisplayName = 'Team'
        Visibility = 'Visibility'
        SiteSizeGB = 'Site Size (GB)'
        TotalChannels = 'Channels'
        PublicChannels = 'Public'
        PrivateChannels = 'Private'
        SharedChannels = 'Shared'
    }
    $tableHtml = New-HtmlTable -Data $teamRows -Columns $tableColumns -ColumnHeaders $tableHeaders -CssClass 'data-table wrap-cells'
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This summarizes Teams inventory, channel breakdowns, and Teams Voice license indicators.</p>
    <h3>Recommended Next Steps</h3>
    <p>If Teams Voice is required, confirm Phone System and PSTN add-ons are correctly licensed and configured.</p>
</div>
"@
    
    return $kpiHtml + $voiceSummaryTable + $tableHtml + $footerHtml
}

function Build-DomainsSection {
    param(
        [array]$Domains,
        [object]$SpamFilteringSummary,
        [object]$SMTPRelaySummary,
        [array]$MailFlowConnectors = @()
    )
    
    if ($Domains.Count -eq 0) {
        return "<div class='empty-state'>No domain data available</div>"
    }
    
    $analysis = Get-DomainAnalysis -Domains $Domains -SpamFilteringSummary $SpamFilteringSummary -SMTPRelaySummary $SMTPRelaySummary

    function Convert-MailFlowValueList {
        param($Value)

        if ($null -eq $Value) {
            return @()
        }
        if ($Value -is [string]) {
            return @(
                $Value -split '[,;]' |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }
        if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
            $items = @()
            foreach ($item in $Value) {
                if ($null -eq $item) { continue }
                $itemText = [string]$item
                if (-not [string]::IsNullOrWhiteSpace($itemText)) {
                    $items += $itemText.Trim()
                }
            }
            return @($items)
        }

        $singleText = [string]$Value
        if ([string]::IsNullOrWhiteSpace($singleText)) {
            return @()
        }
        return @($singleText.Trim())
    }

    function Test-ConnectorAppliesToDomain {
        param(
            [object]$Connector,
            [string]$DomainName
        )

        if (-not $Connector -or [string]::IsNullOrWhiteSpace($DomainName)) {
            return $false
        }

        $recipientDomains = Convert-MailFlowValueList -Value $Connector.RecipientDomains
        if ($recipientDomains.Count -eq 0) {
            return $true
        }

        $domainLower = $DomainName.ToLowerInvariant()
        foreach ($recipientDomain in $recipientDomains) {
            $candidate = ([string]$recipientDomain).Trim().ToLowerInvariant()
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            if ($candidate -eq '*') { return $true }
            if ($candidate -eq $domainLower) { return $true }
            if ($candidate -like "*.$domainLower") { return $true }
            if ($candidate -eq "*.$domainLower") { return $true }
        }

        return $false
    }

    function Get-ConnectorEndpointEvidence {
        param([object]$Connector)

        $endpointSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($propertyName in @('SmartHosts', 'SenderIPAddresses', 'EFSkipIPs', 'EFSkipMailGateway', 'TlsDomain')) {
            if (-not $Connector.PSObject.Properties[$propertyName]) { continue }
            foreach ($value in (Convert-MailFlowValueList -Value $Connector.$propertyName)) {
                $null = $endpointSet.Add($value)
            }
        }
        return @($endpointSet)
    }
    
    # Basic domain info
    $tableColumns = @('Domain','Verified','AuthenticationType','DomainType','IsDefault')
    $tableHeaders = @{
        'Domain' = 'Domain Name'
        'Verified' = 'Verified'
        'AuthenticationType' = 'Auth Type'
        'DomainType' = 'Domain Type'
        'IsDefault' = 'Default'
    }
    
    $riskColumns = @{
        'Verified' = { param($val) $val -eq $false }
    }
    
    $tableHtml = "<h3>Domain Configuration</h3>"
    $tableHtml += New-HtmlTable -Data $Domains -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns

    # Mail authentication summary
    $mailEnabledCustomDomains = @(
        $Domains | Where-Object {
            ([string]$_.Domain -notmatch '(?i)\.onmicrosoft\.com$') -and
            ($_.Verified -eq $true) -and
            ($_.Office365MailExchanger -eq $true -or ([int]$_.TotalDomainRecipients -gt 0))
        }
    )

    $mailAuthKpiHtml = ""
    if ($mailEnabledCustomDomains.Count -gt 0) {
        $spfPassing = @($mailEnabledCustomDomains | Where-Object { $_.SpfIncludesM365 -eq $true }).Count
        $dmarcConfigured = @($mailEnabledCustomDomains | Where-Object { $_.DmarcConfigured -eq $true }).Count
        $dmarcEnforced = @($mailEnabledCustomDomains | Where-Object { $_.DmarcConfigured -eq $true -and $_.DmarcPolicy -in @('quarantine', 'reject') }).Count
        $dkimComplete = @($mailEnabledCustomDomains | Where-Object { [int]$_.DkimSelectorsConfigured -ge 2 }).Count

        $mailAuthKpiHtml = "<h3 id='domains-mailauth' style='margin-top:30px;'>Mail Authentication & Anti-Spoofing</h3>"
        $mailAuthKpiHtml += "<div class='kpi-grid'>"
        $mailAuthKpiHtml += New-KpiCard -Title "Mail Domains" -Value (Format-AssessmentHtmlNumber $mailEnabledCustomDomains.Count)
        $mailAuthKpiHtml += New-KpiCard -Title "SPF Coverage" -Value ("{0}%" -f [math]::Round((($spfPassing / $mailEnabledCustomDomains.Count) * 100), 1))
        $mailAuthKpiHtml += New-KpiCard -Title "DKIM (2 Selectors)" -Value ("{0}%" -f [math]::Round((($dkimComplete / $mailEnabledCustomDomains.Count) * 100), 1))
        $mailAuthKpiHtml += New-KpiCard -Title "DMARC Enforced" -Value ("{0}%" -f [math]::Round((($dmarcEnforced / $mailEnabledCustomDomains.Count) * 100), 1))
        $mailAuthKpiHtml += "</div>"

        $mailAuthRows = @(
            $mailEnabledCustomDomains | ForEach-Object {
                $spfStatus = if ($_.SpfIncludesM365 -eq $true) { 'Pass' } else { 'Needs Review' }
                $dmarcStatus = if ($_.DmarcConfigured -eq $true) { 'Present' } else { 'Missing' }
                $dkimStatus = if ([int]$_.DkimSelectorsConfigured -ge 2) { 'Pass' } else { 'Needs Review' }
                $dmarcPolicyDisplay = if ($_.DmarcPolicy) { [string]$_.DmarcPolicy } else { 'N/A' }
                $dmarcPctDisplay = if ($_.DmarcPercent) { "$($_.DmarcPercent)%" } else { 'N/A' }

                [PSCustomObject]@{
                    Domain = $_.Domain
                    M365MX = $(if ($_.Office365MailExchanger -eq $true) { 'Yes' } else { 'No' })
                    SPF = $spfStatus
                    SPFMode = $(if ($_.SpfPolicyMode) { $_.SpfPolicyMode } else { 'N/A' })
                    DMARC = $dmarcStatus
                    DMARCPolicy = $dmarcPolicyDisplay
                    DMARCPct = $dmarcPctDisplay
                    DKIM = $dkimStatus
                }
            }
        )

        $mailAuthColumns = @('Domain','M365MX','SPF','SPFMode','DMARC','DMARCPolicy','DMARCPct','DKIM')
        $mailAuthHeaders = @{
            'Domain' = 'Domain'
            'M365MX' = 'M365 MX'
            'SPF' = 'SPF'
            'SPFMode' = 'SPF Mode'
            'DMARC' = 'DMARC'
            'DMARCPolicy' = 'DMARC Policy'
            'DMARCPct' = 'DMARC %'
            'DKIM' = 'DKIM'
        }
        $mailAuthRiskColumns = @{
            'SPF' = { param($val) [string]$val -eq 'Needs Review' }
            'DMARC' = { param($val) [string]$val -eq 'Missing' }
            'DKIM' = { param($val) [string]$val -eq 'Needs Review' }
        }
        $mailAuthKpiHtml += New-HtmlTable -Data $mailAuthRows -Columns $mailAuthColumns -ColumnHeaders $mailAuthHeaders -RiskColumns $mailAuthRiskColumns -CssClass 'data-table wrap-cells'
    }
    
    # DNS info
    $dnsColumns = @('Domain','NSRecords','ARecords','MXRecords','Office365MailExchanger')
    $dnsHeaders = @{
        'Domain' = 'Domain'
        'NSRecords' = 'Name Servers'
        'ARecords' = 'A Records'
        'MXRecords' = 'MX Records'
        'Office365MailExchanger' = 'M365 MX'
    }
    
    $tableHtml += "<h3 id='domains-dns' style='margin-top:30px;'>DNS Configuration</h3>"
    $tableHtml += New-HtmlTable -Data $Domains -Columns $dnsColumns -ColumnHeaders $dnsHeaders -CssClass 'data-table wrap-cells'

    # Anti-spoofing control summary
    $antiSpoofRows = @()
    if ($SpamFilteringSummary -or $SMTPRelaySummary) {
        $trustedBypassCount = 0
        if ($SpamFilteringSummary -and $SpamFilteringSummary.PSObject.Properties['TransportRulesWithTrustedIPs']) {
            try { $trustedBypassCount = [int]$SpamFilteringSummary.TransportRulesWithTrustedIPs } catch { $trustedBypassCount = 0 }
        }
        $uses3rdPartyFiltering = $null
        if ($SpamFilteringSummary -and $SpamFilteringSummary.PSObject.Properties['Uses3rdPartyFiltering']) {
            $rawThirdParty = $SpamFilteringSummary.Uses3rdPartyFiltering
            if ($rawThirdParty -is [bool]) {
                $uses3rdPartyFiltering = $rawThirdParty
            } else {
                $uses3rdPartyFiltering = ([string]$rawThirdParty -match '(?i)^(yes|true|enabled)$')
            }
        }

        $smtpAuthEnabled = $null
        $smtpAuthUsers = 0
        $smtpClientAuthDisabled = $null
        if ($SMTPRelaySummary) {
            if ($SMTPRelaySummary.PSObject.Properties['SMTPAuthEnabled']) {
                $rawSmtpAuth = $SMTPRelaySummary.SMTPAuthEnabled
                if ($rawSmtpAuth -is [bool]) {
                    $smtpAuthEnabled = $rawSmtpAuth
                } else {
                    $smtpAuthEnabled = ([string]$rawSmtpAuth -match '(?i)^(yes|true|enabled)$')
                }
            }
            if ($SMTPRelaySummary.PSObject.Properties['SMTPAuthUsers']) {
                try { $smtpAuthUsers = [int]$SMTPRelaySummary.SMTPAuthUsers } catch { $smtpAuthUsers = 0 }
            }
            if ($SMTPRelaySummary.PSObject.Properties['SmtpClientAuthenticationDisabled']) {
                $smtpClientRaw = $SMTPRelaySummary.SmtpClientAuthenticationDisabled
                if ($smtpClientRaw -is [bool]) {
                    $smtpClientAuthDisabled = $smtpClientRaw
                } else {
                    $smtpClientAuthDisabled = ([string]$smtpClientRaw -match '(?i)^(yes|true)$')
                }
            }
        }

        $antiSpoofRows += [PSCustomObject]@{
            Signal = 'Trusted IP / bypass indicators'
            Value = $trustedBypassCount
            Assessment = $(if ($trustedBypassCount -gt 0) { 'Review' } else { 'Good' })
        }
        $antiSpoofRows += [PSCustomObject]@{
            Signal = 'SMTP AUTH enabled mailboxes'
            Value = $smtpAuthUsers
            Assessment = $(if ($smtpAuthEnabled -eq $true -and $smtpAuthUsers -gt 0) { 'Review' } else { 'Good' })
        }
        $antiSpoofRows += [PSCustomObject]@{
            Signal = 'Org-wide SMTP client auth disabled'
            Value = $(if ($null -eq $smtpClientAuthDisabled) { 'Unknown' } elseif ($smtpClientAuthDisabled) { 'Yes' } else { 'No' })
            Assessment = $(if ($smtpClientAuthDisabled -eq $false) { 'Review' } else { 'Good' })
        }
        $antiSpoofRows += [PSCustomObject]@{
            Signal = 'Third-party mail filtering detected'
            Value = $(if ($null -eq $uses3rdPartyFiltering) { 'Unknown' } elseif ($uses3rdPartyFiltering) { 'Yes' } else { 'No' })
            Assessment = $(if ($uses3rdPartyFiltering -eq $true) { 'Review' } else { 'Good' })
        }

        $tableHtml += "<h3 id='domains-antispoof' style='margin-top:30px;'>Anti-Spoofing Control Signals</h3>"
        $tableHtml += New-HtmlTable -Data $antiSpoofRows -Columns @('Signal','Value','Assessment') -ColumnHeaders @{ Signal='Control'; Value='Value'; Assessment='Assessment' } -RiskColumns @{ Assessment = { param($val) [string]$val -eq 'Review' } }
    }

    # Estimated inbound/outbound mail paths per domain (best-effort)
    $mailFlowEstimateRows = @()
    $mailRoutingDomains = @(
        $Domains | Where-Object {
            ([string]$_.Domain -notmatch '(?i)\.onmicrosoft\.com$') -and
            ($_.Verified -eq $true) -and
            ($_.Office365MailExchanger -eq $true -or ([int]$_.TotalDomainRecipients -gt 0))
        }
    )
    $allConnectors = @($MailFlowConnectors)
    $inboundConnectors = @($allConnectors | Where-Object { [string]$_.ConnectorDirection -eq 'Inbound' -and $_.Enabled -ne $false })
    $outboundConnectors = @($allConnectors | Where-Object { [string]$_.ConnectorDirection -eq 'Outbound' -and $_.Enabled -ne $false })

    foreach ($domain in $mailRoutingDomains) {
        $domainName = [string]$domain.Domain
        if ([string]::IsNullOrWhiteSpace($domainName)) { continue }

        $domainInboundConnectors = @($inboundConnectors | Where-Object { Test-ConnectorAppliesToDomain -Connector $_ -DomainName $domainName })
        $domainOutboundConnectors = @($outboundConnectors | Where-Object { Test-ConnectorAppliesToDomain -Connector $_ -DomainName $domainName })

        $mxEvidence = Convert-MailFlowValueList -Value $domain.MXRecords
        $inboundEndpointSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $outboundEndpointSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($mx in $mxEvidence) { $null = $inboundEndpointSet.Add($mx) }
        foreach ($connector in $domainInboundConnectors) {
            foreach ($endpoint in (Get-ConnectorEndpointEvidence -Connector $connector)) {
                $null = $inboundEndpointSet.Add($endpoint)
            }
        }

        $usesSmartHostOutbound = $false
        foreach ($connector in $domainOutboundConnectors) {
            $smartHosts = Convert-MailFlowValueList -Value $connector.SmartHosts
            if ($smartHosts.Count -gt 0 -and $connector.UseMXRecord -ne $true) {
                $usesSmartHostOutbound = $true
            }
            foreach ($endpoint in (Get-ConnectorEndpointEvidence -Connector $connector)) {
                $null = $outboundEndpointSet.Add($endpoint)
            }
        }

        if ($outboundEndpointSet.Count -eq 0) {
            if ($domainOutboundConnectors.Count -gt 0) {
                $null = $outboundEndpointSet.Add('Recipient MX (connector-managed)')
            } else {
                $null = $outboundEndpointSet.Add('Recipient MX (Exchange Online default)')
            }
        }
        if ($inboundEndpointSet.Count -eq 0 -and $domain.Office365MailExchanger -eq $true) {
            $null = $inboundEndpointSet.Add('Microsoft 365 MX / EOP')
        }

        $inboundPath = if ($domain.Office365MailExchanger -eq $true) {
            if ($domainInboundConnectors.Count -gt 0) {
                'Internet -> Microsoft 365 MX/EOP -> Inbound connector path(s) -> Exchange Online'
            } else {
                'Internet -> Microsoft 365 MX/EOP -> Exchange Online'
            }
        } elseif ($domainInboundConnectors.Count -gt 0) {
            'Internet/Partner -> Inbound connector path(s) -> Exchange Online'
        } else {
            'External MX or custom path -> Review routing manually'
        }

        $outboundPath = if ($domainOutboundConnectors.Count -gt 0) {
            if ($usesSmartHostOutbound) {
                'Exchange Online -> Outbound connector smarthost(s) -> Destination'
            } else {
                'Exchange Online -> Outbound connector(s) -> Recipient MX'
            }
        } else {
            'Exchange Online -> Recipient MX (default)'
        }

        $confidence = if ($mxEvidence.Count -gt 0 -and ($domainInboundConnectors.Count -gt 0 -or $domainOutboundConnectors.Count -gt 0)) {
            'High'
        } elseif ($mxEvidence.Count -gt 0 -or $domain.Office365MailExchanger -eq $true -or $domainInboundConnectors.Count -gt 0 -or $domainOutboundConnectors.Count -gt 0) {
            'Medium'
        } else {
            'Low'
        }

        $mailFlowEstimateRows += [PSCustomObject]@{
            Domain = $domainName
            InboundPath = $inboundPath
            InboundEndpoints = [string]::Join(', ', @($inboundEndpointSet | Select-Object -First 8))
            OutboundPath = $outboundPath
            OutboundEndpoints = [string]::Join(', ', @($outboundEndpointSet | Select-Object -First 8))
            Confidence = $confidence
        }
    }

    if ($mailFlowEstimateRows.Count -gt 0) {
        $tableHtml += "<h3 id='domains-mailflow' style='margin-top:30px;'>Estimated Domain Mail Flow Paths</h3>"
        $tableHtml += "<p style='margin-top:8px;color:#5f6b74;font-size:13px;'>Best-effort estimate from domain MX metadata plus Exchange connector configuration. Validate with production transport design before cutover.</p>"
        $tableHtml += New-HtmlTable `
            -Data $mailFlowEstimateRows `
            -Columns @('Domain','InboundPath','InboundEndpoints','OutboundPath','OutboundEndpoints','Confidence') `
            -ColumnHeaders @{
                Domain = 'Domain'
                InboundPath = 'Inbound Path'
                InboundEndpoints = 'Inbound Endpoints (Estimated)'
                OutboundPath = 'Outbound Path'
                OutboundEndpoints = 'Outbound Endpoints (Estimated)'
                Confidence = 'Confidence'
            } `
            -RiskColumns @{
                Confidence = { param($val) [string]$val -eq 'Low' }
            } `
            -CssClass 'data-table wrap-cells'
    }
    
    # Recipient counts
    $recipColumns = @('Domain','PrimarySMTPRecipients','AliasOnlyRecipients','TotalDomainRecipients')
    $recipHeaders = @{
        'Domain' = 'Domain'
        'PrimarySMTPRecipients' = 'Primary SMTP'
        'AliasOnlyRecipients' = 'Alias Only'
        'TotalDomainRecipients' = 'Total Domain Recipients'
    }
    
    $tableHtml += "<h3 id='domains-recipients' style='margin-top:30px;'>Recipient Distribution</h3>"
    $tableHtml += New-HtmlTable -Data $Domains -Columns $recipColumns -ColumnHeaders $recipHeaders
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows all domains in your tenant, DNS/mail-auth posture (SPF, DKIM, DMARC), recipient usage, and anti-spoofing control signals.</p>
    <h3>Recommended Next Steps</h3>
    <p>For mail-enabled custom domains, target SPF alignment, two DKIM selectors, and DMARC enforcement (quarantine/reject). Review trusted-IP bypasses and SMTP AUTH exceptions regularly.</p>
</div>
"@
    
    return $tableHtml + $mailAuthKpiHtml + $footerHtml
}

function Build-DevicesSection {
    param([array]$Devices)
    
    if ($Devices.Count -eq 0) {
        return "<div class='empty-state'>No device data available</div>"
    }
    
    $analysis = Get-DeviceAnalysis -Devices $Devices
    
    # Group by OS
    $osSummary = $Devices | Group-Object OperatingSystem | ForEach-Object {
        $osDevices = $_.Group
        $total = $osDevices.Count
        $stale = ($osDevices | Where-Object { $_.DeviceStale -eq $true }).Count
        $managed = ($osDevices | Where-Object { $_.IsManaged -eq $true }).Count
        $compliant = ($osDevices | Where-Object { $_.IsCompliant -eq $true }).Count
        
        [PSCustomObject]@{
            OS = $_.Name
            Total = $total
            AllDevicePercentage = [math]::Round(($total / $Devices.Count) * 100, 1)
            StaleDevicePercentage = [math]::Round(($stale / $total) * 100, 1)
            MDMManagedPercentage = [math]::Round(($managed / $total) * 100, 1)
            IsCompliantPercentage = [math]::Round(($compliant / $total) * 100, 1)
        }
    } | Sort-Object Total -Descending
    
    # KPIs
    $kpiHtml = "<div class='kpi-grid'>"
    $totalDevices = $Devices.Count
    $managedCount = ($Devices | Where-Object { $_.IsManaged -eq $true }).Count
    $compliantCount = ($Devices | Where-Object { $_.IsCompliant -eq $true }).Count
    
    $kpiHtml += New-KpiCard -Title "Total Devices" -Value (Format-AssessmentHtmlNumber $totalDevices)
    
    $managedPct = if ($totalDevices -gt 0) { ($managedCount / $totalDevices) * 100 } else { 0 }
    $managedTheme = if ($managedPct -ge 80) { 'success' } elseif ($managedPct -ge 60) { 'warning' } else { 'danger' }
    $managedValue = "{0:N1}%" -f [double]$managedPct
    $kpiHtml += New-KpiCard -Title "MDM Managed" -Value $managedValue -Theme $managedTheme

    $compliantPct = if ($totalDevices -gt 0) { ($compliantCount / $totalDevices) * 100 } else { 0 }
    $compTheme = if ($compliantPct -ge 80) { 'success' } elseif ($compliantPct -ge 60) { 'warning' } else { 'danger' }
    $compliantValue = "{0:N1}%" -f [double]$compliantPct
    $kpiHtml += New-KpiCard -Title "Compliant" -Value $compliantValue -Theme $compTheme
    $kpiHtml += "</div>"
    
    # OS Summary table
    $tableColumns = @('OS','Total','AllDevicePercentage','StaleDevicePercentage','MDMManagedPercentage','IsCompliantPercentage')
    $tableHeaders = @{
        'OS' = 'Operating System'
        'Total' = 'Count'
        'AllDevicePercentage' = '% of All'
        'StaleDevicePercentage' = '% Stale'
        'MDMManagedPercentage' = '% Managed'
        'IsCompliantPercentage' = '% Compliant'
    }
    
    $riskColumns = @{
        'StaleDevicePercentage' = { param($val) [double]$val -gt 20 }
        'IsCompliantPercentage' = { param($val) [double]$val -lt 80 }
    }
    
    $tableHtml = New-HtmlTable -Data $osSummary -Columns $tableColumns -ColumnHeaders $tableHeaders -RiskColumns $riskColumns
    
    # Top OS versions
    $versionHtml = "<h3 style='margin-top:30px;'>Top 5 OS Versions</h3>"
    foreach ($os in ($osSummary | Select-Object -First 3)) {
        $osDevices = $Devices | Where-Object { $_.OperatingSystem -eq $os.OS }
        $topVersions = $osDevices | Group-Object OperatingSystemVersion | 
            Select-Object @{N='Version';E={$_.Name}}, Count |
            Sort-Object Count -Descending |
            Select-Object -First 5
        
        if ($topVersions.Count -gt 0) {
            $versionHtml += "<h4>$($os.OS)</h4>"
            $versionHtml += New-HtmlTable -Data $topVersions -Columns @('Version','Count') -CssClass 'data-table'
        }
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This shows your device inventory, management status, and compliance. Stale devices (inactive > 6 months) should be reviewed for removal.</p>
    <h3>Recommended Next Steps</h3>
    <p>Enroll unmanaged devices in Intune. Investigate non-compliant devices. Remove or disable stale devices to reduce security risk.</p>
</div>
"@
    
    return $kpiHtml + $tableHtml + $versionHtml + $footerHtml
}

function Build-SecureScoreSection {
    param([array]$SecureScore)
    
    if ($SecureScore.Count -eq 0) {
        return "<div class='empty-state'>No Secure Score data available</div>"
    }
    
    # Use most recent score
    $latest = $SecureScore | Sort-Object CreatedDateTime -Descending | Select-Object -First 1
    
    $currentScore = [int]$latest.CurrentScore
    $maxScore = [int]$latest.MaxScore
    $scorePct = if ($maxScore -gt 0) { ($currentScore / $maxScore) * 100 } else { 0 }
    
    # KPI
    $scoreTheme = if ($scorePct -ge 80) { 'success' } elseif ($scorePct -ge 60) { 'warning' } else { 'danger' }
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Current Score" -Value "$currentScore / $maxScore" -Subtitle "$(Format-AssessmentHtmlNumber $scorePct -DecimalPlaces 1)%" -Theme $scoreTheme
    $kpiHtml += New-KpiCard -Title "Enabled Services" -Value $latest.EnabledServicesCount
    $kpiHtml += "</div>"
    
    # Comparison
    $compHtml = "<h3>Score Comparison</h3><div class='kpi-grid'>"
    
    if ($latest.SimilarSizeOrg_ComparitiveScore) {
        $compScore = [int]$latest.SimilarSizeOrg_ComparitiveScore
        $compTheme = if ($currentScore -ge $compScore) { 'success' } else { 'warning' }
        $compHtml += New-KpiCard -Title "Similar Size Orgs" -Value $compScore -Subtitle "Your score: $currentScore" -Theme $compTheme
    }
    
    if ($latest.AllTenants_ComparitiveScore) {
        $allScore = [int]$latest.AllTenants_ComparitiveScore
        $allTheme = if ($currentScore -ge $allScore) { 'success' } else { 'warning' }
        $compHtml += New-KpiCard -Title "All Tenants Average" -Value $allScore -Subtitle "Your score: $currentScore" -Theme $allTheme
    }
    
    $compHtml += "</div>"
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>Microsoft Secure Score measures your security posture. Higher scores indicate better protection against threats. Compare your score to similar organizations.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review top improvement actions in the Microsoft 365 Security Center. Prioritize actions with the highest score impact. Aim for gradual improvement each quarter.</p>
</div>
"@
    
    return $kpiHtml + $compHtml + $footerHtml
}

# Exchange Hybrid
function Build-ExchangeHybridSection {
    param(
        [object]$HybridInfo,
        [object]$ExchangeFederation
    )
    
    if (-not $HybridInfo) {
        return "<div class='empty-state'>No Exchange hybrid data available</div>"
    }
    
    $isHybrid = $HybridInfo.IsHybridConfigured -eq $true
    $partial = (-not $isHybrid) -and (
        ($HybridInfo.InboundOnPremConnectorCount -gt 0) -or
        ($HybridInfo.OutboundOnPremConnectorCount -gt 0) -or
        ($HybridInfo.MigrationEndpointCount -gt 0) -or
        ($HybridInfo.EvidenceCount -gt 0)
    )
    if ($HybridInfo.HybridStatus) {
        $statusText = $HybridInfo.HybridStatus
    } elseif ($isHybrid) {
        $statusText = 'Hybrid'
    } elseif ($partial) {
        $statusText = 'Possible Hybrid'
    } else {
        $statusText = 'None'
    }
    
    $callout = if ($isHybrid) {
        New-CalloutBox -Type 'info' -Title 'Hybrid Configuration Detected' -Content "Exchange hybrid indicators were found. Hybrid type: <strong>$($HybridInfo.HybridType)</strong>."
    } elseif ($partial) {
        New-CalloutBox -Type 'warning' -Title 'Partial Hybrid Indicators' -Content "Some hybrid indicators were found, but a full hybrid configuration could not be confirmed."
    } else {
        New-CalloutBox -Type 'success' -Title 'No Hybrid Detected' -Content "No Exchange hybrid indicators were detected in this tenant."
    }
    
    $statusTextDisplay = $statusText -replace '/', ' / '
    $kpiHtml = "<div class='kpi-grid'>"
    $statusTheme = if ($isHybrid -or $statusText -eq 'Possible Hybrid') { 'warning' } else { 'success' }
    $hybridTypeValue = if ($HybridInfo.HybridType) { $HybridInfo.HybridType } else { "Unknown" }
    $kpiHtml += New-KpiCard -Title "Hybrid Status" -Value $statusTextDisplay -Theme $statusTheme
    $kpiHtml += New-KpiCard -Title "Hybrid Type" -Value $hybridTypeValue
    $kpiHtml += New-KpiCard -Title "Evidence Signals" -Value (Format-AssessmentHtmlNumber $HybridInfo.EvidenceCount)
    $kpiHtml += New-KpiCard -Title "On-Prem Connectors" -Value (Format-AssessmentHtmlNumber ($HybridInfo.InboundOnPremConnectorCount + $HybridInfo.OutboundOnPremConnectorCount))
    $kpiHtml += "</div>"
    
    $tableData = @(
        [PSCustomObject]@{ Signal = 'IntraOrg Connectors'; Count = $HybridInfo.IntraOrgConnectorCount },
        [PSCustomObject]@{ Signal = 'Organization Relationships'; Count = $HybridInfo.OrgRelationshipCount },
        [PSCustomObject]@{ Signal = 'Inbound On-Prem Connectors'; Count = $HybridInfo.InboundOnPremConnectorCount },
        [PSCustomObject]@{ Signal = 'Outbound On-Prem Connectors'; Count = $HybridInfo.OutboundOnPremConnectorCount },
        [PSCustomObject]@{ Signal = 'Migration Endpoints'; Count = $HybridInfo.MigrationEndpointCount },
        [PSCustomObject]@{ Signal = 'Mail Flow On-Prem Connectors'; Count = $HybridInfo.MailFlowOnPremConnectorCount }
    )
    
    $tableHtml = New-HtmlTable -Data $tableData -Columns @('Signal','Count') -ColumnHeaders @{'Signal'='Signal';'Count'='Count'}

    $endpointRows = @()
    if ($HybridInfo.MigrationEndpoints) {
        $endpointRows += [PSCustomObject]@{ Category = 'Migration Endpoints'; Values = $HybridInfo.MigrationEndpoints }
    }
    if ($HybridInfo.InboundOnPremConnectors) {
        $endpointRows += [PSCustomObject]@{ Category = 'Inbound On-Prem Connectors'; Values = $HybridInfo.InboundOnPremConnectors }
    }
    if ($HybridInfo.OutboundOnPremConnectors) {
        $endpointRows += [PSCustomObject]@{ Category = 'Outbound On-Prem Connectors'; Values = $HybridInfo.OutboundOnPremConnectors }
    }
    # Federation details shown in subsection below to avoid duplication

    $endpointsHtml = ""
    if ($endpointRows.Count -gt 0) {
        $endpointsHtml = "<h3 style='margin-top:20px;'>Endpoints / Objects</h3>"
        $endpointsHtml += New-HtmlTable -Data $endpointRows -Columns @('Category','Values') -ColumnHeaders @{'Category'='Category';'Values'='Values'}
    }
    
    $evidenceHtml = ""
    if ($HybridInfo.Evidence) {
        $evidenceItems = $HybridInfo.Evidence -split '; ' | Where-Object { $_ -and $_.Trim().Length -gt 0 }
        if ($evidenceItems.Count -gt 0) {
            $evidenceHtml = "<h3 style='margin-top:20px;'>Evidence</h3><ul>"
            foreach ($item in $evidenceItems) {
                $encoded = [System.Web.HttpUtility]::HtmlEncode($item)
                $evidenceHtml += "<li>$encoded</li>"
            }
            $evidenceHtml += "</ul>"
        }
    }
    
    $federationHtml = ""
    if ($ExchangeFederation) {
        $fedRows = @(
            [PSCustomObject]@{ Setting = 'Organization Relationships'; Value = $ExchangeFederation.OrganizationRelationships },
            [PSCustomObject]@{ Setting = 'IntraOrg Connectors'; Value = $ExchangeFederation.IntraOrgConnectors }
        )
        $federationHtml = "<h3 style='margin-top:20px;'>Exchange Federation</h3>" +
            (New-HtmlTable -Data $fedRows -Columns @('Setting','Value') -ColumnHeaders @{'Setting'='Setting';'Value'='Value'})
    }

    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>Hybrid indicators show whether Exchange Online is integrated with an on-premises Exchange environment for mail flow, free/busy, or migrations.</p>
    <h3>Recommended Next Steps</h3>
    <p>If hybrid is detected, validate hybrid configuration in the Exchange admin center and confirm connectors align with your intended mail routing. If partial indicators exist, review connectors and organization relationships for stale or incomplete setup.</p>
</div>
"@
    
    return $callout + $kpiHtml + $tableHtml + $endpointsHtml + $evidenceHtml + $federationHtml + $footerHtml
}

function Build-IdentityAdminSection {
    param(
        [array]$Users,
        [array]$Admins,
        [array]$Groups,
        [array]$ExchangeGroups
    )
    
    $totalUsers = $Users.Count
    $memberUsers = ($Users | Where-Object { $_.UserType -eq 'Member' }).Count
    $guestUsers = ($Users | Where-Object { $_.UserType -eq 'Guest' }).Count
    $externalUsers = ($Users | Where-Object { $_.UserPrincipalName -like '*#EXT#*' }).Count
    $externalMembers = ($Users | Where-Object { $_.UserType -eq 'Member' -and $_.UserPrincipalName -like '*#EXT#*' }).Count
    $internalMembers = ($Users | Where-Object { $_.UserType -eq 'Member' -and $_.UserPrincipalName -notlike '*#EXT#*' }).Count
    $groupsCount = $Groups.Count
    $adminCount = $Admins.Count
    
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Total Users" -Value (Format-AssessmentHtmlNumber $totalUsers)
    $kpiHtml += New-KpiCard -Title "Members (Internal)" -Value (Format-AssessmentHtmlNumber $internalMembers)
    $kpiHtml += New-KpiCard -Title "Guests" -Value (Format-AssessmentHtmlNumber $guestUsers)
    $kpiHtml += New-KpiCard -Title "External/B2B" -Value (Format-AssessmentHtmlNumber $externalUsers)
    $kpiHtml += New-KpiCard -Title "Groups" -Value (Format-AssessmentHtmlNumber $groupsCount)
    $kpiHtml += New-KpiCard -Title "Admins" -Value (Format-AssessmentHtmlNumber $adminCount)
    $kpiHtml += "</div>"
    
    $userSummary = @(
        [PSCustomObject]@{ Category = 'Members (Internal)'; Count = $internalMembers },
        [PSCustomObject]@{ Category = 'Guests'; Count = $guestUsers },
        [PSCustomObject]@{ Category = 'Members (External)'; Count = $externalMembers }
    )
    $userTable = "<h3>User Type Summary</h3>" + (New-HtmlTable -Data $userSummary -Columns @('Category','Count') -ColumnHeaders @{'Category'='Category';'Count'='Count'})
    
    $dynamicGroups = ($Groups | Where-Object { $_.MembershipType -eq 'Dynamic Group' }).Count
    $dynamicDistributionGroups = ($ExchangeGroups | Where-Object { $_.RecipientTypeDetails -eq 'DynamicDistributionGroup' }).Count
    $m365Groups = ($Groups | Where-Object { $_.GroupType -eq 'Microsoft 365' }).Count
    $securityGroups = ($Groups | Where-Object { $_.SecurityEnabled -eq $true }).Count
    $onPremGroups = ($Groups | Where-Object { $_.OnPremisesSyncEnabled -eq $true -or $_.Source -eq 'On-Premises' }).Count
    $cloudGroups = ($Groups | Where-Object { $_.OnPremisesSyncEnabled -ne $true -and $_.Source -ne 'On-Premises' }).Count
    $groupSummary = @(
        [PSCustomObject]@{ Category = 'Total groups'; Count = $groupsCount },
        [PSCustomObject]@{ Category = 'Dynamic groups'; Count = $dynamicGroups },
        [PSCustomObject]@{ Category = 'Dynamic distribution groups (mail enabled)'; Count = $dynamicDistributionGroups },
        [PSCustomObject]@{ Category = 'M365 groups'; Count = $m365Groups },
        [PSCustomObject]@{ Category = 'Security groups'; Count = $securityGroups },
        [PSCustomObject]@{ Category = 'Cloud groups'; Count = $cloudGroups },
        [PSCustomObject]@{ Category = 'On-premises groups'; Count = $onPremGroups }
    )
    $groupTable = "<h3 style='margin-top:20px;'>Group Overview</h3>" +
        (New-HtmlTable -Data $groupSummary -Columns @('Category','Count') -ColumnHeaders @{'Category'='Category';'Count'='Count'})
    
    $roleCounts = @{}
    foreach ($admin in $Admins) {
        $roles = @()
        if ($admin.Role) {
            $roles = $admin.Role -split ',\s*'
        }
        foreach ($role in $roles) {
            if (-not $roleCounts.ContainsKey($role)) { $roleCounts[$role] = 0 }
            $roleCounts[$role]++
        }
    }
    $roleSummary = $roleCounts.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{ Role = $_.Key; AdminCount = $_.Value }
    } | Sort-Object AdminCount -Descending
    
    $roleTable = ""
    if ($roleSummary.Count -gt 0) {
        $roleTable = "<h3 style='margin-top:20px;'>Admin Roles Summary</h3>" +
            (New-HtmlTable -Data $roleSummary -Columns @('Role','AdminCount') -ColumnHeaders @{'Role'='Role';'AdminCount'='Admins'})
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This section summarizes user types, external/B2B presence, group inventory, and admin role coverage.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review guest and external users for necessity. Ensure admin roles follow least privilege and are distributed appropriately.</p>
</div>
"@
    
    return $kpiHtml + $userTable + $groupTable + $roleTable + $footerHtml
}

function Build-TenantOverviewSection {
    param(
        [object]$AuthConfig,
        [object]$AdConnect,
        [array]$ConditionalAccess
    )

    $federatedDomains = if ($AuthConfig -and $AuthConfig.FederatedDomains) { $AuthConfig.FederatedDomains } else { @() }
    $signOnProvider = if ($federatedDomains.Count -gt 0) { "Federated (ADFS/3rd-party IdP)" } else { "Entra ID (cloud)" }
    $dirSyncEnabled = if ($AdConnect -and $AdConnect.Summary) { if ($AdConnect.Summary.OnPremisesSyncEnabled -eq $true) { "Yes" } else { "No" } } else { "Unknown" }
    $ssprWriteback = "Not collected"
    $mfaProvider = if ($AuthConfig -and $AuthConfig.MFAEnabled -eq $true) {
        $methods = if ($AuthConfig.MFAMethods -and $AuthConfig.MFAMethods.Count -gt 0) { $AuthConfig.MFAMethods -join ", " } else { "Methods not listed" }
        "Entra ID MFA ($methods)"
    } else { "Not detected" }
    $ssoAppsCount = if ($AuthConfig -and $AuthConfig.SSOApplications) { $AuthConfig.SSOApplications.Count } else { 0 }
    $enterpriseSso = if ($AuthConfig -and $AuthConfig.SSOEnabled -eq $true) { "Yes ($ssoAppsCount apps)" } else { "No" }
    $caCount = if ($ConditionalAccess) { $ConditionalAccess.Count } else { 0 }
    $conditionalAccessAnswer = if ($caCount -gt 0) { "Yes ($caCount policies)" } else { "No" }
    $appProxy = "Not collected"
    $privateAccess = "Not collected"
    $pim = "Not collected"

    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "Sign-On Provider" -Value $signOnProvider
    $kpiHtml += New-KpiCard -Title "DirSync Enabled" -Value $dirSyncEnabled -Theme ($(if ($dirSyncEnabled -eq 'Yes') { 'success' } elseif ($dirSyncEnabled -eq 'No') { 'warning' } else { 'default' }))
    $kpiHtml += New-KpiCard -Title "MFA" -Value $(if ($AuthConfig -and $AuthConfig.MFAEnabled) { 'Enabled' } else { 'Not Detected' }) -Theme ($(if ($AuthConfig -and $AuthConfig.MFAEnabled) { 'success' } else { 'warning' }))
    $kpiHtml += New-KpiCard -Title "Conditional Access" -Value $conditionalAccessAnswer -Theme ($(if ($caCount -gt 0) { 'success' } else { 'warning' }))
    $kpiHtml += "</div>"

    $rows = @(
        [PSCustomObject]@{ Topic = "Sign-on provider"; Answer = $signOnProvider; Notes = if ($federatedDomains.Count -gt 0) { "Federated domains: $($federatedDomains -join ', ')" } else { "No federated domains detected" } }
        [PSCustomObject]@{ Topic = "Directory sync from on-prem AD"; Answer = $dirSyncEnabled; Notes = if ($AdConnect -and $AdConnect.Summary) { "Last sync: $($AdConnect.Summary.OnPremisesLastSyncDateTime)" } else { "No sync data available" } }
        [PSCustomObject]@{ Topic = "SSPR / Password writeback"; Answer = $ssprWriteback; Notes = "Not collected in this report" }
        [PSCustomObject]@{ Topic = "MFA provider"; Answer = $mfaProvider; Notes = if ($AuthConfig -and $AuthConfig.MFAMethods) { "Methods: $($AuthConfig.MFAMethods -join ', ')" } else { "No MFA method data found" } }
        [PSCustomObject]@{ Topic = "Enterprise SSO to SaaS apps"; Answer = $enterpriseSso; Notes = if ($ssoAppsCount -gt 0) { "Provide app list spreadsheet (separately)" } else { "No SSO apps detected" } }
        [PSCustomObject]@{ Topic = "Conditional Access"; Answer = $conditionalAccessAnswer; Notes = if ($caCount -gt 0) { "Policies detected in tenant" } else { "No policies detected" } }
        [PSCustomObject]@{ Topic = "Entra App Proxy"; Answer = $appProxy; Notes = "Not collected in this report" }
        [PSCustomObject]@{ Topic = "Entra Private Access"; Answer = $privateAccess; Notes = "Not collected in this report" }
        [PSCustomObject]@{ Topic = "Privileged Identity Management (PIM)"; Answer = $pim; Notes = "Not collected in this report" }
    )

    $tableHtml = New-HtmlTable -Data $rows -Columns @('Topic','Answer','Notes') -ColumnHeaders @{ 'Topic'='Topic'; 'Answer'='Answer'; 'Notes'='Notes' }

    $footerHtml = @"
<div class='section-footer'>
    <h3>Notes</h3>
    <p>This section is intended for Solutions Architect review. Items marked "Not collected" require manual validation.</p>
</div>
"@

    return $kpiHtml + $tableHtml + $footerHtml
}

function Build-ConditionalAccessMfaSection {
    param(
        [array]$ConditionalAccessPolicies,
        [object]$AuthConfig,
        [array]$Users,
        [object]$MfaRegistrationSummary
    )
    
    $totalUsers = $Users.Count
    $enabledPolicies = @($ConditionalAccessPolicies | Where-Object { $_.State -eq 'enabled' })
    $reportOnlyPolicies = @($ConditionalAccessPolicies | Where-Object { $_.State -eq 'reportOnly' })
    $disabledPolicies = @($ConditionalAccessPolicies | Where-Object { $_.State -eq 'disabled' })
    $mfaPolicies = @($enabledPolicies | Where-Object { $_.GrantControls_BuiltInControls -match '(?i)mfa' })
    $mfaEnabled = ($mfaPolicies.Count -gt 0) -or ($AuthConfig -and $AuthConfig.MFAEnabled -eq $true)
    
    $mfaMethods = if ($AuthConfig -and $AuthConfig.MFAMethods) { ($AuthConfig.MFAMethods -join ', ') } else { 'Not available' }
    $passwordlessMethods = if ($AuthConfig -and $AuthConfig.PasswordlessMethods) { ($AuthConfig.PasswordlessMethods -join ', ') } else { 'Not available' }
    
    $mfaAllUsers = $false
    foreach ($policy in $enabledPolicies) {
        $includesAll = ($policy.IncludedUsersCount -eq 'All') -or ($policy.IncludedUsers -eq 'All')
        $excludesNone = ($policy.ExcludedUsersCount -eq 0 -or [string]::IsNullOrWhiteSpace($policy.ExcludedUsers))
        $requiresMfa = ($policy.GrantControls_BuiltInControls -match '(?i)mfa')
        if ($includesAll -and $excludesNone -and $requiresMfa) {
            $mfaAllUsers = $true
            break
        }
    }
    
    $caAllUsers = $false
    foreach ($policy in $enabledPolicies) {
        $includesAll = ($policy.IncludedUsersCount -eq 'All') -or ($policy.IncludedUsers -eq 'All')
        $excludesNone = ($policy.ExcludedUsersCount -eq 0 -or [string]::IsNullOrWhiteSpace($policy.ExcludedUsers))
        if ($includesAll -and $excludesNone) {
            $caAllUsers = $true
            break
        }
    }
    
    $pctNotMfa = if ($mfaAllUsers) { '0%' } else { 'Unknown' }
    $pctNotCA = if ($caAllUsers) { '0%' } else { 'Unknown' }
    
    $mfaEnforcedValue = if ($mfaEnabled) { 'Yes' } else { 'No' }
    $mfaEnforcedTheme = if ($mfaEnabled) { 'success' } else { 'warning' }
    
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "CA Policies" -Value (Format-AssessmentHtmlNumber $ConditionalAccessPolicies.Count)
    $kpiHtml += New-KpiCard -Title "Enabled" -Value (Format-AssessmentHtmlNumber $enabledPolicies.Count)
    $kpiHtml += New-KpiCard -Title "Report-Only" -Value (Format-AssessmentHtmlNumber $reportOnlyPolicies.Count)
    $kpiHtml += New-KpiCard -Title "MFA Enforced" -Value $mfaEnforcedValue -Theme $mfaEnforcedTheme
    $kpiHtml += "</div>"
    
    $summaryRows = @(
        [PSCustomObject]@{ Metric = 'MFA CA Policies Enabled'; Value = $mfaPolicies.Count },
        [PSCustomObject]@{ Metric = 'MFA Methods (Policy)'; Value = $mfaMethods },
        [PSCustomObject]@{ Metric = 'Passwordless Methods'; Value = $passwordlessMethods },
        [PSCustomObject]@{ Metric = '% Users Not Covered by CA'; Value = $pctNotCA },
        [PSCustomObject]@{ Metric = '% Users Not Enforced with MFA'; Value = $pctNotMfa },
        [PSCustomObject]@{ Metric = 'Default User Can Create Apps'; Value = $(if ($AuthConfig -and $AuthConfig.PSObject.Properties['DefaultUserCanCreateApps']) { $AuthConfig.DefaultUserCanCreateApps } else { 'Not available' }) },
        [PSCustomObject]@{ Metric = 'Permission Grant Policies'; Value = $(if ($AuthConfig -and $AuthConfig.PSObject.Properties['PermissionGrantPoliciesAssigned']) { (@($AuthConfig.PermissionGrantPoliciesAssigned) -join ', ') } elseif ($AuthConfig -and $AuthConfig.PSObject.Properties['PermissionGrantPolicies']) { [string]$AuthConfig.PermissionGrantPolicies } else { 'Not available' }) },
        [PSCustomObject]@{ Metric = 'Admin Consent Workflow'; Value = $(if ($AuthConfig -and $AuthConfig.PSObject.Properties['AdminConsentWorkflowEnabled']) { $AuthConfig.AdminConsentWorkflowEnabled } else { 'Not available' }) }
    )

    if ($MfaRegistrationSummary) {
        $summaryRows += [PSCustomObject]@{ Metric = 'MFA Registered Users'; Value = $MfaRegistrationSummary.RegisteredUsers }
        $summaryRows += [PSCustomObject]@{ Metric = 'MFA Not Registered Users'; Value = $MfaRegistrationSummary.NotRegisteredUsers }
        $summaryRows += [PSCustomObject]@{ Metric = 'MFA Registration %'; Value = "$($MfaRegistrationSummary.RegistrationPercent)%" }
    }
    $summaryTable = "<h3>Conditional Access & MFA Summary</h3>" +
        (New-HtmlTable -Data $summaryRows -Columns @('Metric','Value') -ColumnHeaders @{'Metric'='Metric';'Value'='Value'})

    $methodTable = ""
    if ($MfaRegistrationSummary -and $MfaRegistrationSummary.MethodCounts) {
        $methodCounts = $MfaRegistrationSummary.MethodCounts
        $methodRows = @()
        if ($methodCounts -is [hashtable]) {
            $methodRows = $methodCounts.GetEnumerator() | ForEach-Object {
                [PSCustomObject]@{ Method = $_.Key; Count = $_.Value }
            }
        } elseif ($methodCounts -is [pscustomobject]) {
            $methodRows = $methodCounts.PSObject.Properties | ForEach-Object {
                [PSCustomObject]@{ Method = $_.Name; Count = $_.Value }
            }
        }
        $methodRows = $methodRows | Sort-Object Count -Descending
        if ($methodRows.Count -gt 0) {
            $methodTable = "<h3 style='margin-top:20px;'>Registered MFA Methods</h3>" +
                (New-HtmlTable -Data $methodRows -Columns @('Method','Count') -ColumnHeaders @{'Method'='Method';'Count'='Count'})
        }
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This summarizes conditional access policies and MFA enforcement signals. Per-user registration/enforcement data is not collected in this report.</p>
    <h3>Recommended Next Steps</h3>
    <p>Ensure MFA is enforced for all users via Conditional Access, review report-only policies, and validate registration methods in your authentication policy.</p>
</div>
"@
    
    return $kpiHtml + $summaryTable + $methodTable + $footerHtml
}

function Build-AdConnectSection {
    param([object]$AdConnect)
    
    if (-not $AdConnect -or -not $AdConnect.Summary) {
        return "<div class='empty-state'>No AD Connect/Sync data available</div>"
    }
    
    $summary = $AdConnect.Summary
    $syncEnabledFlag = ($summary.OnPremisesSyncEnabled -eq $true)
    $syncEnabled = if ($syncEnabledFlag) { 'Yes' } else { 'No' }
    $lastSyncValue = $summary.OnPremisesLastSyncDateTime
    $lastSyncDate = $null
    if ($lastSyncValue) {
        try {
            $lastSyncDate = [datetime]$lastSyncValue
        }
        catch {
            $lastSyncDate = $null
        }
    }
    $lastSync = if ($lastSyncDate) { $lastSyncDate.ToString('yyyy-MM-dd HH:mm:ss') } elseif ($lastSyncValue) { [string]$lastSyncValue } else { 'N/A' }
    $lastSyncSubtitle = $null
    $lastSyncTheme = 'default'
    if ($syncEnabledFlag -and $lastSyncDate) {
        $ageDays = [math]::Round(((Get-Date).ToUniversalTime() - $lastSyncDate.ToUniversalTime()).TotalDays, 1)
        $lastSyncSubtitle = "$ageDays day(s) ago"
        if ($ageDays -le 1) {
            $lastSyncTheme = 'success'
        }
        elseif ($ageDays -le 7) {
            $lastSyncTheme = 'warning'
        }
        else {
            $lastSyncTheme = 'danger'
        }
    }
    elseif ($syncEnabledFlag) {
        $lastSyncSubtitle = 'Last sync timestamp unavailable'
        $lastSyncTheme = 'warning'
    }
    $errorCount = if ($AdConnect.ErrorCount -gt 0) { $AdConnect.ErrorCount } else { 0 }
    $syncTheme = if ($syncEnabledFlag) { 'success' } else { 'warning' }
    $errorTheme = if ($errorCount -gt 0) { 'warning' } else { 'success' }
    
    $kpiHtml = "<div class='kpi-grid'>"
    $kpiHtml += New-KpiCard -Title "DirSync Enabled" -Value $syncEnabled -Theme $syncTheme
    $kpiHtml += New-KpiCard -Title "Last Sync" -Value $lastSync -Subtitle $lastSyncSubtitle -Theme $lastSyncTheme
    $kpiHtml += New-KpiCard -Title "Sync Errors" -Value (Format-AssessmentHtmlNumber $errorCount) -Theme $errorTheme
    $kpiHtml += "</div>"
    
    $serviceTable = ""
    if ($AdConnect.SyncServices -and $AdConnect.SyncServices.Count -gt 0) {
        $serviceTable = "<h3>Sync Service(s)</h3>" +
            (New-HtmlTable -Data $AdConnect.SyncServices -Columns @('ServiceName','ServerName','LastSyncTime') -ColumnHeaders @{'ServiceName'='Service';'ServerName'='Server';'LastSyncTime'='Last Sync'})
    }
    
    $errorTable = ""
    if ($AdConnect.RecentErrors -and $AdConnect.RecentErrors.Count -gt 0) {
        $errorRows = $AdConnect.RecentErrors | ForEach-Object {
            [PSCustomObject]@{
                TimeGenerated = $_.TimeGenerated
                Error = if ($_.Error) { $_.Error } elseif ($_.Reason) { $_.Reason } else { $_.ErrorCode }
                Description = if ($_.Description) { $_.Description } elseif ($_.Message) { $_.Message } else { $null }
            }
        }
        $errorTable = "<h3 style='margin-top:20px;'>Recent Sync Errors</h3>" +
            (New-HtmlTable -Data $errorRows -Columns @('TimeGenerated','Error','Description') -ColumnHeaders @{'TimeGenerated'='Time';'Error'='Error';'Description'='Description'})
    }
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>AD Connect/Sync status indicates whether on-prem directory synchronization is active and healthy.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review sync errors and verify the AD Connect server is healthy. Confirm last sync is recent.</p>
</div>
"@
    
    return $kpiHtml + $serviceTable + $errorTable + $footerHtml
}

function Build-CrossTenantAccessSection {
    param(
        [object]$CrossTenantAccess,
        [object]$ExternalIdentities
    )
    
    if (-not $CrossTenantAccess -and -not $ExternalIdentities) {
        return "<div class='empty-state'>No cross-tenant or external identity data available</div>"
    }
    
    $hasCrossTenant = $CrossTenantAccess -and ($CrossTenantAccess.PartnerCount -gt 0)
    
    $callout = if ($hasCrossTenant) {
        New-CalloutBox -Type 'info' -Title 'Cross-Tenant Access Detected' -Content 'This tenant has cross-tenant access partner settings configured.'
    } else {
        New-CalloutBox -Type 'success' -Title 'No Cross-Tenant Partners Detected' -Content 'No cross-tenant access partners were detected.'
    }
    
    $kpiHtml = "<div class='kpi-grid'>"
    $partnerCount = if ($CrossTenantAccess) { $CrossTenantAccess.PartnerCount } else { 0 }
    $b2bPresent = if ($ExternalIdentities -and $ExternalIdentities.B2BManagementPolicyPresent) { $true } else { $false }
    $b2bValue = if ($b2bPresent) { "Present" } else { "Not Found" }
    $b2bTheme = if ($b2bPresent) { 'success' } else { 'warning' }
    $kpiHtml += New-KpiCard -Title "Cross-Tenant Partners" -Value (Format-AssessmentHtmlNumber $partnerCount)
    $kpiHtml += New-KpiCard -Title "B2B Policy" -Value $b2bValue -Theme $b2bTheme
    $kpiHtml += "</div>"
    
    $crossRows = @()
    if ($CrossTenantAccess) {
        $crossRows += [PSCustomObject]@{ Setting = 'Cross-Tenant Access Policy'; Value = $CrossTenantAccess.HasCrossTenantAccessPolicy }
        $crossRows += [PSCustomObject]@{ Setting = 'Partner Count'; Value = $CrossTenantAccess.PartnerCount }
        $crossRows += [PSCustomObject]@{ Setting = 'Default Inbound MFA Accepted'; Value = $CrossTenantAccess.DefaultInboundAccess }
        $crossRows += [PSCustomObject]@{ Setting = 'Default Outbound MFA Accepted'; Value = $CrossTenantAccess.DefaultOutboundAccess }
        $crossRows += [PSCustomObject]@{ Setting = 'Default B2B Direct Connect (Inbound)'; Value = $CrossTenantAccess.DefaultB2BDirectConnectInbound }
        $crossRows += [PSCustomObject]@{ Setting = 'Default B2B Direct Connect (Outbound)'; Value = $CrossTenantAccess.DefaultB2BDirectConnectOutbound }
    }
    $crossHtml = "<h3>Cross-Tenant Access</h3>" + (New-HtmlTable -Data $crossRows -Columns @('Setting','Value') -ColumnHeaders @{'Setting'='Setting';'Value'='Value'})
    
    $partnerDetailRows = @()
    if ($CrossTenantAccess -and $CrossTenantAccess.PartnerTenantDetails) {
        foreach ($tenant in $CrossTenantAccess.PartnerTenantDetails) {
            $partnerDetailRows += [PSCustomObject]@{
                TenantId = $tenant.TenantId
                DisplayName = $tenant.DisplayName
                B2BDirectConnectInbound = $tenant.B2BDirectConnectInbound
                B2BDirectConnectOutbound = $tenant.B2BDirectConnectOutbound
                TrustMfa = $tenant.TrustMfa
                TrustCompliantDevices = $tenant.TrustCompliantDevices
                TrustHybridJoinedDevices = $tenant.TrustHybridJoinedDevices
            }
        }
    }
    $partnerDetailsHtml = ""
    if ($partnerDetailRows.Count -gt 0) {
        $partnerDetailsHtml = "<h3 style='margin-top:20px;'>Partner Tenant Details</h3>" +
            (New-HtmlTable -Data $partnerDetailRows -Columns @('TenantId','DisplayName','B2BDirectConnectInbound','B2BDirectConnectOutbound','TrustMfa','TrustCompliantDevices','TrustHybridJoinedDevices') -ColumnHeaders @{
                'TenantId'='Tenant ID'
                'DisplayName'='Display Name'
                'B2BDirectConnectInbound'='B2B Direct Connect Inbound'
                'B2BDirectConnectOutbound'='B2B Direct Connect Outbound'
                'TrustMfa'='Trust MFA'
                'TrustCompliantDevices'='Trust Compliant Devices'
                'TrustHybridJoinedDevices'='Trust Hybrid Joined Devices'
            } -CssClass 'data-table wrap-headers')
    }
    
    $externalRows = @()
    if ($ExternalIdentities) {
        $externalRows += [PSCustomObject]@{ Setting = 'B2B Management Policy Present'; Value = $ExternalIdentities.B2BManagementPolicyPresent }
        $externalRows += [PSCustomObject]@{ Setting = 'Guest User Role'; Value = $ExternalIdentities.GuestUserRole }
        $externalRows += [PSCustomObject]@{ Setting = 'Invitations Allowed'; Value = $ExternalIdentities.InvitationsAllowed }
    }
    $externalHtml = "<h3 style='margin-top:20px;'>External Identities</h3>" + (New-HtmlTable -Data $externalRows -Columns @('Setting','Value') -ColumnHeaders @{'Setting'='Setting';'Value'='Value'})
    
    $footerHtml = @"
<div class='section-footer'>
    <h3>What This Means</h3>
    <p>This section summarizes cross-tenant access policies and external identity settings for B2B collaboration.</p>
    <h3>Recommended Next Steps</h3>
    <p>Review partner access rules for least-privilege. Ensure B2B invitation policies align with your security requirements.</p>
</div>
"@
    
    return $callout + $kpiHtml + $crossHtml + $partnerDetailsHtml + $externalHtml + $footerHtml
}

#endregion

#region Main Report Builder

function Build-FindingsPanel {
    param([array]$AllFindings)
    
    if ($AllFindings.Count -eq 0) {
        $content = New-CalloutBox -Type 'success' -Title 'No Issues Found' -Content 'Your tenant configuration looks good! No automatic risk flags were detected.'
        return $content
    }
    
    # Group findings by category
    $grouped = $AllFindings | Group-Object { 
        if ($_.Category) { $_.Category } else { 'Other' }
    }
    
    # Sort groups by priority
    $sortedGroups = $grouped | Sort-Object {
        switch ($_.Name) {
            'Critical' { 1 }
            'At Capacity' { 2 }
            'High Utilization' { 3 }
            'Overall Utilization' { 4 }
            default { 5 }
        }
    }
    
    $html = "<div class='findings-panel'><h2>🔍 Findings of Note</h2>"
    
    # Count by severity
    $riskCount = ($AllFindings | Where-Object { $_.Type -eq 'Risk' }).Count
    $warningCount = ($AllFindings | Where-Object { $_.Type -eq 'Warning' }).Count
    $infoCount = ($AllFindings | Where-Object { $_.Type -eq 'Info' }).Count
    
    # Summary badges
    $html += "<div style='margin-bottom: 20px; display: flex; gap: 10px; flex-wrap: wrap;'>"
    if ($riskCount -gt 0) {
        $html += "<span class='badge badge-risk' style='font-size: 0.9em; padding: 8px 12px;'>🔴 $riskCount Critical</span>"
    }
    if ($warningCount -gt 0) {
        $html += "<span class='badge badge-warning' style='font-size: 0.9em; padding: 8px 12px;'>⚠️ $warningCount Warnings</span>"
    }
    if ($infoCount -gt 0) {
        $html += "<span class='badge badge-info' style='font-size: 0.9em; padding: 8px 12px;'>ℹ️ $infoCount Info</span>"
    }
    $html += "</div>"
    
    # Render findings by category
    foreach ($group in $sortedGroups) {
        $categoryName = $group.Name
        $categoryFindings = $group.Group
        
        # Category header with appropriate icon
        $categoryIcon = switch ($categoryName) {
            'Critical' { '🔴' }
            'At Capacity' { '⚠️' }
            'High Utilization' { '⚠️' }
            'Overall Utilization' { 'ℹ️' }
            default { '📋' }
        }
        
        # Add description for each category
        $categoryDesc = switch ($categoryName) {
            'Critical' { 'Over allocated - consuming more licenses than purchased' }
            'At Capacity' { 'Fully allocated - no licenses available for new users' }
            'High Utilization' { 'Running low - consider purchasing more soon' }
            'Overall Utilization' { 'Tenant-wide license usage summary' }
            default { '' }
        }
        
        $html += "<h3 style='margin-top: 20px; margin-bottom: 5px; color: #333; font-size: 1.1em;'>$categoryIcon $categoryName</h3>"
        if ($categoryDesc) {
            $html += "<p style='margin: 0 0 10px 0; color: #666; font-size: 0.9em; font-style: italic;'>$categoryDesc</p>"
        }
        $html += "<ul class='findings-list'>"
        
        foreach ($finding in $categoryFindings) {
            $class = "finding-$($finding.Type.ToLower())"
            $anchor = if ($finding.Anchor) { "#$($finding.Anchor)" } else { "#" }
            $html += "<li class='$class'><a href='$anchor'>$($finding.Message)</a></li>"
        }
        
        $html += "</ul>"
    }
    
    $html += "</div>"
    return $html
}

function Get-AssessmentSectionWorkload {
    [CmdletBinding()]
    param([string]$SectionId)

    switch ($SectionId) {
        'tenant-overview' { return 'Foundation' }
        'licenses' { return 'Commercial' }
        'domains' { return 'Messaging' }
        'recipients' { return 'Messaging' }
        'mailboxes' { return 'Messaging' }
        'email-activity' { return 'Messaging' }
        'inactive-mailboxes' { return 'Messaging' }
        'exchange-hybrid' { return 'Messaging' }
        'cross-tenant-access' { return 'Messaging' }
        'identity-admins' { return 'Identity & Security' }
        'conditional-access-mfa' { return 'Identity & Security' }
        'devices' { return 'Identity & Security' }
        'secure-score' { return 'Identity & Security' }
        'ad-connect' { return 'Platform' }
        'sharepoint-onedrive' { return 'Collaboration' }
        'ownership-governance' { return 'Collaboration' }
        'teams' { return 'Collaboration' }
        default { return 'Other' }
    }
}

function Build-Navigation {
    param([array]$Sections)

    $groupOrder = @('Foundation', 'Commercial', 'Messaging', 'Identity & Security', 'Collaboration', 'Platform', 'Other')
    $groups = [ordered]@{}
    foreach ($name in $groupOrder) {
        $groups[$name] = @()
    }

    foreach ($section in $Sections) {
        $workload = Get-AssessmentSectionWorkload -SectionId $section.Id
        if (-not $groups.Contains($workload)) {
            $groups[$workload] = @()
        }
        $groups[$workload] += $section
    }

    $html = "<div class='nav-container'><nav class='nav'>"
    $html += "<a href='#highlights' class='nav-link nav-link-primary'>Highlights</a>"
    foreach ($groupName in $groups.Keys) {
        $items = @($groups[$groupName])
        if ($items.Count -eq 0) {
            continue
        }
        $html += "<div class='nav-group'>"
        $html += "<span class='nav-group-label'>$groupName</span>"
        foreach ($section in $items) {
            $html += "<a href='#$($section.Id)' class='nav-link'>$($section.Name)</a>"
        }
        $html += "</div>"
    }

    $html += "</nav></div>"
    return $html
}

function Sort-AssessmentSections {
    [CmdletBinding()]
    param([array]$Sections)

    if (-not $Sections -or $Sections.Count -eq 0) {
        return @()
    }

    $orderedIds = @(
        'tenant-overview',
        'licenses',
        'domains',
        'recipients',
        'mailboxes',
        'email-activity',
        'inactive-mailboxes',
        'exchange-hybrid',
        'cross-tenant-access',
        'identity-admins',
        'conditional-access-mfa',
        'devices',
        'secure-score',
        'sharepoint-onedrive',
        'ownership-governance',
        'teams',
        'ad-connect'
    )

    $rank = @{}
    $idx = 0
    foreach ($id in $orderedIds) {
        $rank[$id] = $idx
        $idx++
    }

    return @(
        $Sections | Sort-Object @{
            Expression = {
                if ($rank.ContainsKey($_.Id)) { $rank[$_.Id] } else { 9999 }
            }
        }, @{
            Expression = { [string]$_.Name }
        }
    )
}

function New-TenantHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [hashtable]$TenantStatsHash = $script:tenantStatsHash,
        
        [Parameter()]
        [string]$OutputPath,
        
        [Parameter()]
        [hashtable]$Thresholds,
        
        [Parameter()]
        [string[]]$IncludeSections,

        [Parameter()]
        [string]$JsonPath,

        [Parameter()]
        [switch]$UseJsonCache
    )
    
    #region Validation
    
    if ($JsonPath -and ($UseJsonCache -or -not $TenantStatsHash -or $TenantStatsHash.Count -eq 0)) {
        $loaded = Import-TenantStatsJson -Path $JsonPath
        if ($loaded) {
            $TenantStatsHash = $loaded
        }
    }

    if (-not $TenantStatsHash -or $TenantStatsHash.Count -eq 0) {
        throw "TenantStatsHash is null or empty. Please ensure data collection has completed successfully or provide -JsonPath."
    }
    
    # Merge thresholds
    if ($Thresholds) {
        foreach ($key in $Thresholds.Keys) {
            $script:DefaultThresholds[$key] = $Thresholds[$key]
        }
    }
    
    #endregion
    
    #region Gather Data
    
    #Write-Host "Building HTML report..." -ForegroundColor Cyan
    #Write-Host "  Validating hashtable structure..." -ForegroundColor Gray
    
    # Extract and normalize data through shared assessment context helper.
    $context = Get-TenantAssessmentContext -TenantStatsHash $TenantStatsHash

    $licenses = $context.Licenses
    Write-Verbose "Licenses: Found $($licenses.Count) items"

    $recipients = $context.Recipients
    Write-Verbose "Recipients: Found $($recipients.Count) items"

    $mailboxes = $context.Mailboxes
    Write-Verbose "Mailboxes: Found $($mailboxes.Count) items"

    $inactiveMailboxes = $context.InactiveMailboxes
    Write-Verbose "Inactive Mailboxes: Found $($inactiveMailboxes.Count) items"

    $publicFolders = $context.PublicFolders
    Write-Verbose "Public Folders: Found $($publicFolders.Count) items"

    $sharepoint = $context.SharePoint
    Write-Verbose "SharePoint: Found $($sharepoint.Count) items"

    $onedrive = $context.OneDrive
    Write-Verbose "OneDrive: Found $($onedrive.Count) items"

    $unmanagedObjects = $context.UnmanagedObjects
    Write-Verbose "Unmanaged objects: Found $($unmanagedObjects.Count) items"

    $oneDriveOwnerMismatches = $context.OneDriveOwnerMismatches
    Write-Verbose "OneDrive owner mismatches: Found $($oneDriveOwnerMismatches.Count) items"

    $ownershipGovernanceSummary = $context.OwnershipGovernanceSummary

    $domains = $context.Domains
    Write-Verbose "Domains: Found $($domains.Count) items"

    $devices = $context.Devices
    Write-Verbose "Devices: Found $($devices.Count) items"

    $secureScore = $context.SecureScore
    Write-Verbose "Secure Score: Found $($secureScore.Count) items"

    $teams = $context.Teams
    Write-Verbose "Teams: Found $($teams.Count) items"
    $teamsVoice = $context.TeamsVoice

    $users = $context.Users
    Write-Verbose "Users: Found $($users.Count) items"

    $admins = $context.Admins
    Write-Verbose "Admins: Found $($admins.Count) items"

    $groups = $context.Groups
    Write-Verbose "Groups: Found $($groups.Count) items"

    $exchangeGroups = $context.ExchangeGroups
    Write-Verbose "Exchange Groups: Found $($exchangeGroups.Count) items"

    $conditionalAccess = $context.ConditionalAccess
    Write-Verbose "Conditional Access: Found $($conditionalAccess.Count) items"

    $authConfig = $context.AuthConfig
    $mfaRegistrationSummary = $context.MfaRegistrationSummary
    $adConnect = $context.AdConnect
    $hybridInfo = $context.HybridInfo
    $federationExchange = $context.FederationExchange
    $federationCrossTenant = $context.FederationCrossTenant
    $federationExternal = $context.FederationExternal
    $spamFilteringSummary = $context.SpamFilteringSummary
    $smtpRelaySummary = $context.SMTPRelaySummary
    $mailFlowConnectors = $context.MailFlowConnectors
    $primaryMailboxStatsCollectionSummary = $context.PrimaryMailboxStatsCollectionSummary
    $emailActivitySummary = $context.EmailActivitySummary
    $emailActivityTopSenders = $context.EmailActivityTopSenders
    $emailActivityTopReceivers = $context.EmailActivityTopReceivers
    Write-Verbose "Email activity: TopSenders=$($emailActivityTopSenders.Count), TopReceivers=$($emailActivityTopReceivers.Count)"
    
    Write-Host "  Data extraction complete" -ForegroundColor Green
    
    #endregion
    
    #region Build Sections and Collect ALL Findings
    
    $allFindings = @()
    $sectionContents = @()
    
    # Tenant Overview
    $sectionContents += @{
        Id = 'tenant-overview'
        Name = 'Tenant Overview'
        Content = Build-TenantOverviewSection -AuthConfig $authConfig -AdConnect $adConnect -ConditionalAccess $conditionalAccess
    }

    # Licenses - ALWAYS analyze first for findings
    if ($licenses.Count -gt 0) {
        #Write-Host "  Building Licenses section..." -ForegroundColor Gray
        $licAnalysis = Get-LicenseAnalysis -Licenses $licenses -UserCount $users.Count
        $allFindings += $licAnalysis.Findings
        $sectionContents += @{
            Id = 'licenses'
            Name = 'License Overview'
            Content = Build-LicenseSection -Licenses $licenses -UserCount $users.Count
        }
    }

    # Domains
    if ($domains.Count -gt 0) {
        #Write-Host "  Building Domains section..." -ForegroundColor Gray
        $domainAnalysis = Get-DomainAnalysis -Domains $domains -SpamFilteringSummary $spamFilteringSummary -SMTPRelaySummary $smtpRelaySummary
        $allFindings += $domainAnalysis.Findings
        $sectionContents += @{
            Id = 'domains'
            Name = 'Domain Details'
            Content = Build-DomainsSection -Domains $domains -SpamFilteringSummary $spamFilteringSummary -SMTPRelaySummary $smtpRelaySummary -MailFlowConnectors $mailFlowConnectors
        }
    }
    
    # Recipients
    if ($recipients.Count -gt 0) {
        #Write-Host "  Building Recipients section..." -ForegroundColor Gray
        $sectionContents += @{
            Id = 'recipients'
            Name = 'Recipient Objects'
            Content = Build-RecipientsSection -Recipients $recipients
        }
    }
    
    # Identity & Admins
    if ($users.Count -gt 0 -or $admins.Count -gt 0 -or $groups.Count -gt 0) {
        #Write-Host "  Building Identity & Admins section..." -ForegroundColor Gray
        $identityAnalysis = Get-IdentityAdminAnalysis -Users $users -Admins $admins -Groups $groups -ConditionalAccessPolicies $conditionalAccess
        $allFindings += $identityAnalysis.Findings
        $sectionContents += @{
            Id = 'identity-admins'
            Name = 'Identity & Admins'
            Content = Build-IdentityAdminSection -Users $users -Admins $admins -Groups $groups -ExchangeGroups $exchangeGroups
        }
    }
    
    # Mailboxes
    if ($mailboxes.Count -gt 0 -or $publicFolders.Count -gt 0 -or $primaryMailboxStatsCollectionSummary) {
        #Write-Host "  Building Mailboxes section..." -ForegroundColor Gray
        $mbxAnalysis = Get-MailboxAnalysis -Mailboxes $mailboxes
        $allFindings += $mbxAnalysis.Findings
        $sectionContents += @{
            Id = 'mailboxes'
            Name = 'Mailbox Overview'
            Content = Build-MailboxesSection -Mailboxes $mailboxes -InactiveMailboxes $inactiveMailboxes -PublicFolders $publicFolders -MailboxStatsSummary $primaryMailboxStatsCollectionSummary
        }
    }

    if ($emailActivitySummary -or $emailActivityTopSenders.Count -gt 0 -or $emailActivityTopReceivers.Count -gt 0) {
        $sectionContents += @{
            Id = 'email-activity'
            Name = 'Email Activity'
            Content = Build-EmailActivitySection -TopSenders $emailActivityTopSenders -TopReceivers $emailActivityTopReceivers -EmailActivitySummary $emailActivitySummary
        }
    }
    
    # Inactive Mailboxes
    if ($inactiveMailboxes.Count -gt 0) {
        #Write-Host "  Building Inactive Mailboxes section..." -ForegroundColor Gray
        $inactiveAnalysis = Get-InactiveMailboxAnalysis -InactiveMailboxes $inactiveMailboxes
        $allFindings += $inactiveAnalysis.Findings
        $sectionContents += @{
            Id = 'inactive-mailboxes'
            Name = 'Inactive Mailboxes'
            Content = Build-InactiveMailboxesSection -InactiveMailboxes $inactiveMailboxes
        }
    }
    
    # SharePoint & OneDrive
    if ($sharepoint.Count -gt 0 -or $onedrive.Count -gt 0) {
        #Write-Host "  Building SharePoint/OneDrive section..." -ForegroundColor Gray
        $spodAnalysis = Get-SharePointOneDriveAnalysis -SharePointSites $sharepoint -OneDriveSites $onedrive
        $allFindings += $spodAnalysis.Findings
        $sectionContents += @{
            Id = 'sharepoint-onedrive'
            Name = 'SharePoint & OneDrive'
            Content = Build-SharePointOneDriveSection -SharePointSites $sharepoint -OneDriveSites $onedrive
        }
    }

    if ($ownershipGovernanceSummary -or $unmanagedObjects.Count -gt 0 -or $oneDriveOwnerMismatches.Count -gt 0) {
        $ownershipAnalysis = Get-OwnershipGovernanceAnalysis `
            -UnmanagedObjects $unmanagedObjects `
            -OneDriveOwnerMismatches $oneDriveOwnerMismatches `
            -OwnershipGovernanceSummary $ownershipGovernanceSummary
        $allFindings += $ownershipAnalysis.Findings
        $sectionContents += @{
            Id = 'ownership-governance'
            Name = 'Ownership Governance'
            Content = Build-OwnershipGovernanceSection -UnmanagedObjects $unmanagedObjects -OneDriveOwnerMismatches $oneDriveOwnerMismatches -OwnershipGovernanceSummary $ownershipGovernanceSummary
        }
    }

    # Teams
    if ($teams.Count -gt 0) {
        #Write-Host "  Building Teams section..." -ForegroundColor Gray
        $sectionContents += @{
            Id = 'teams'
            Name = 'Teams Overview'
            Content = Build-TeamsSection -Teams $teams -Licenses $licenses -TeamsVoice $teamsVoice -UserCount $users.Count
        }
    }
    
    # Devices
    if ($devices.Count -gt 0) {
        #Write-Host "  Building Devices section..." -ForegroundColor Gray
        $deviceAnalysis = Get-DeviceAnalysis -Devices $devices
        $allFindings += $deviceAnalysis.Findings
        $sectionContents += @{
            Id = 'devices'
            Name = 'Device Overview'
            Content = Build-DevicesSection -Devices $devices
        }
    }
    
    # AD Connect / Sync
    if ($adConnect) {
        #Write-Host "  Building AD Connect section..." -ForegroundColor Gray
        $adAnalysis = Get-AdConnectAnalysis -AdConnect $adConnect
        $allFindings += $adAnalysis.Findings
        $sectionContents += @{
            Id = 'ad-connect'
            Name = 'AD Connect / Sync'
            Content = Build-AdConnectSection -AdConnect $adConnect
        }
    }
    
    # Conditional Access & MFA
    if ($conditionalAccess.Count -gt 0 -or $authConfig) {
        #Write-Host "  Building Conditional Access & MFA section..." -ForegroundColor Gray
        $caAnalysis = Get-ConditionalAccessMfaAnalysis -ConditionalAccessPolicies $conditionalAccess -AuthConfig $authConfig -TotalUsers $users.Count
        $allFindings += $caAnalysis.Findings
        $sectionContents += @{
            Id = 'conditional-access-mfa'
            Name = 'Conditional Access & MFA'
            Content = Build-ConditionalAccessMfaSection -ConditionalAccessPolicies $conditionalAccess -AuthConfig $authConfig -Users $users -MfaRegistrationSummary $mfaRegistrationSummary
        }
    }
    
    # Exchange Hybrid (includes federation subsection)
    if ($hybridInfo) {
        #Write-Host "  Building Exchange Hybrid section..." -ForegroundColor Gray
        $hybridAnalysis = Get-ExchangeHybridAnalysis -HybridInfo $hybridInfo
        $fedAnalysis = Get-FederationAnalysis -ExchangeFederation $federationExchange -CrossTenantAccess $federationCrossTenant -ExternalIdentities $federationExternal
        $allFindings += $hybridAnalysis.Findings
        $allFindings += $fedAnalysis.Findings
        $sectionContents += @{
            Id = 'exchange-hybrid'
            Name = 'Exchange Hybrid'
            Content = Build-ExchangeHybridSection -HybridInfo $hybridInfo -ExchangeFederation $federationExchange
        }
    }
    
    if ($federationCrossTenant -or $federationExternal) {
        #Write-Host "  Building Cross-Tenant Access section..." -ForegroundColor Gray
        $sectionContents += @{
            Id = 'cross-tenant-access'
            Name = 'Cross-Tenant Access'
            Content = Build-CrossTenantAccessSection -CrossTenantAccess $federationCrossTenant -ExternalIdentities $federationExternal
        }
    }
    
    # Secure Score
    if ($secureScore.Count -gt 0) {
        #Write-Host "  Building Secure Score section..." -ForegroundColor Gray
        $sectionContents += @{
            Id = 'secure-score'
            Name = 'Secure Score'
            Content = Build-SecureScoreSection -SecureScore $secureScore
        }
    }
    
    #Write-Host "  Total findings collected: $($allFindings.Count)" -ForegroundColor Green
    $sectionContents = Sort-AssessmentSections -Sections $sectionContents
    
    #endregion
    
    #region Build HTML
    
    $tenantIdentity = Get-TenantIdentityDetails -TenantStatsHash $TenantStatsHash
    $tenantName = [string]$tenantIdentity.DisplayName
    $tenantIdText = if (-not [string]::IsNullOrWhiteSpace([string]$tenantIdentity.TenantId)) { [string]$tenantIdentity.TenantId } else { 'Unavailable' }
    $defaultDomainText = if (-not [string]::IsNullOrWhiteSpace([string]$tenantIdentity.DefaultDomain)) { [string]$tenantIdentity.DefaultDomain } else { 'Unavailable' }
    $initialDomainText = if (-not [string]::IsNullOrWhiteSpace([string]$tenantIdentity.InitialDomain)) { [string]$tenantIdentity.InitialDomain } else { 'Unavailable' }
    $countryText = if (-not [string]::IsNullOrWhiteSpace([string]$tenantIdentity.Country)) { [string]$tenantIdentity.Country } else { 'Unavailable' }

    # Determine output path (include tenant name)
    if (-not $OutputPath) {
        $safeName = ($tenantName -replace '[^\w\-. ]', '').Trim()
        if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "TenantReport" }
        if ($global:ExportDetails) {
            $exportDir = Split-Path -Path $global:ExportDetails -Parent
            $OutputPath = Join-Path $exportDir "$safeName-Report.html"
        } else {
            $OutputPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "$safeName-Report.html"
        }
    }
    
    # Ensure .html extension
    if ($OutputPath -notmatch '\.html$') {
        $OutputPath += '.html'
    }
    
    $reportDate = Get-Date -Format "MMMM dd, yyyy h:mm tt"
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$tenantName - Tenant Snapshot | Arraya Solutions</title>
    
    <!-- Load Chart.js FIRST -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    
    <!-- Load CSS -->
    $(Get-HtmlStyle)
</head>
<body>
    <div class="container">
        <!-- Header -->
        <div class="report-header">
            <div class="report-brand">Prepared by Arraya Solutions</div>
            <h1>$tenantName</h1>
            <p style="font-size:1.1em;">Tenant Snapshot</p>
            <div class="report-meta">
                <div class="report-meta-item">📅 Generated: $reportDate</div>
                <div class="report-meta-item">🆔 Tenant ID: $tenantIdText</div>
                <div class="report-meta-item">🌐 Default Domain: $defaultDomainText</div>
                <div class="report-meta-item">📛 Initial Domain: $initialDomainText</div>
                <div class="report-meta-item">🌍 Country: $countryText</div>
                <div class="report-meta-item">✅ Data Freshness: Current session</div>
                <div class="report-meta-item">🏢 Prepared by: Arraya Solutions</div>
            </div>
        </div>
        
        <!-- Navigation -->
        $(Build-Navigation -Sections $sectionContents)
        
        <!-- Findings Panel - Uses ALL collected findings -->
        <div id="highlights">
            $(Build-FindingsPanel -AllFindings $allFindings)
        </div>
        
        <!-- Sections -->
"@
    
    foreach ($section in $sectionContents) {
        $sectionWorkload = Get-AssessmentSectionWorkload -SectionId $section.Id
        $htmlContent += @"
        <div class="section" id="$($section.Id)">
            <div class="section-workload">$sectionWorkload</div>
            <h2>$($section.Name)</h2>
            $($section.Content)
        </div>
"@
    }
    
    $htmlContent += @"
    </div>
    
    <!-- Chart.js function definitions -->
    $(Get-HtmlScript)
</body>
</html>
"@
    
    #endregion
    
    #region Write File
    
    try {
        $htmlContent | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
        #Write-Host "✅ HTML report saved: $OutputPath" -ForegroundColor Green
        
        # Display summary
        Write-Host "`n📊 Report Summary:" -ForegroundColor Cyan
        Write-Host "   Sections: $($sectionContents.Count)" -ForegroundColor White
        Write-Host "   Findings: $($allFindings.Count)" -ForegroundColor White
        
        $criticalCount = ($allFindings | Where-Object { $_.Type -eq 'Risk' }).Count
        $warningCount = ($allFindings | Where-Object { $_.Type -eq 'Warning' }).Count
        $infoCount = ($allFindings | Where-Object { $_.Type -eq 'Info' }).Count
        
        if ($criticalCount -gt 0) {
            Write-Host "   🔴 Critical: $criticalCount" -ForegroundColor Red
        }
        if ($warningCount -gt 0) {
            Write-Host "   ⚠️  Warnings: $warningCount" -ForegroundColor Yellow
        }
        if ($infoCount -gt 0) {
            Write-Host "   ℹ️  Info: $infoCount" -ForegroundColor Cyan
        }
        
        return [PSCustomObject]@{
            Success = $true
            OutputPath = $OutputPath
            SectionCount = $sectionContents.Count
            FindingsCount = $allFindings.Count
            CriticalCount = $criticalCount
            WarningCount = $warningCount
            InfoCount = $infoCount
        }
    } catch {
        Write-Error "Failed to write HTML report: $_"
        return [PSCustomObject]@{
            Success = $false
            OutputPath = $null
            Error = $_.Exception.Message
        }
    }
    #endregion
}

#endregion

function Get-TenantPdfRendererPath {
    [CmdletBinding()]
    param()

    $candidatePaths = @(
        'C:\Program Files\Google\Chrome\Application\chrome.exe',
        'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
        'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
        'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
    )

    foreach ($path in $candidatePaths) {
        if (Test-Path -Path $path) {
            return [PSCustomObject]@{
                Path   = $path
                Engine = [System.IO.Path]::GetFileNameWithoutExtension($path)
            }
        }
    }

    foreach ($commandName in @('chrome', 'chromium', 'msedge')) {
        $command = Get-Command -Name $commandName -ErrorAction SilentlyContinue
        if ($command) {
            return [PSCustomObject]@{
                Path   = $command.Source
                Engine = $commandName
            }
        }
    }

    return $null
}

function Export-TenantHtmlReportPdf {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$HtmlPath,
        [Parameter(Mandatory)]
        [string]$PdfPath,
        [int]$VirtualTimeBudgetMs = 15000,
        [int]$WaitTimeoutSeconds = 45
    )

    if (-not (Test-Path -Path $HtmlPath)) {
        throw "HTML report not found: $HtmlPath"
    }

    $renderer = Get-TenantPdfRendererPath
    if (-not $renderer) {
        return [PSCustomObject]@{
            Success = $false
            PdfPath  = $null
            Renderer = $null
            Error    = 'No supported Chromium-based PDF renderer was found.'
        }
    }

    $resolvedHtmlPath = (Resolve-Path -Path $HtmlPath).Path
    $resolvedPdfPath = [System.IO.Path]::GetFullPath($PdfPath)
    $pdfDirectory = Split-Path -Path $resolvedPdfPath -Parent
    if (-not (Test-Path -Path $pdfDirectory)) {
        New-Item -Path $pdfDirectory -ItemType Directory -Force | Out-Null
    }

    $htmlUri = [System.Uri]::new($resolvedHtmlPath).AbsoluteUri
    $tempProfilePath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("tenant-html-pdf-" + [guid]::NewGuid().ToString('N'))
    New-Item -Path $tempProfilePath -ItemType Directory -Force | Out-Null

    try {
        if (Test-Path -Path $resolvedPdfPath) {
            Remove-Item -Path $resolvedPdfPath -Force -ErrorAction SilentlyContinue
        }

        $arguments = @(
            '--headless=new',
            '--disable-gpu',
            '--hide-scrollbars',
            '--allow-file-access-from-files',
            "--virtual-time-budget=$VirtualTimeBudgetMs",
            "--print-to-pdf=$resolvedPdfPath",
            $htmlUri
        )
        if ($renderer.Engine -eq 'msedge') {
            $arguments = @(
                '--headless=new',
                '--disable-gpu',
                '--hide-scrollbars',
                '--allow-file-access-from-files',
                "--user-data-dir=$tempProfilePath",
                "--virtual-time-budget=$VirtualTimeBudgetMs",
                "--print-to-pdf=$resolvedPdfPath",
                $htmlUri
            )
        }

        & $renderer.Path @arguments *> $null
        $exitCode = $LASTEXITCODE

        $deadline = (Get-Date).AddSeconds($WaitTimeoutSeconds)
        while ((Get-Date) -lt $deadline) {
            if (Test-Path -Path $resolvedPdfPath) {
                $fileInfo = Get-Item -Path $resolvedPdfPath -ErrorAction SilentlyContinue
                if ($fileInfo -and $fileInfo.Length -gt 0) {
                    return [PSCustomObject]@{
                        Success = $true
                        PdfPath  = $resolvedPdfPath
                        Renderer = $renderer.Engine
                        Error    = $null
                    }
                }
            }

            Start-Sleep -Milliseconds 500
        }

        return [PSCustomObject]@{
            Success = $false
            PdfPath  = $null
            Renderer = $renderer.Engine
            Error    = "Renderer exited with code $exitCode, but no PDF was produced."
        }
    }
    finally {
        Remove-Item -Path $tempProfilePath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function New-TenantAssessmentHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$TenantStatsHash,
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    function ConvertTo-AssessmentArray {
        param([string]$Key)

        function Get-AssessmentContainerValue {
            param(
                [AllowNull()]$Container,
                [Parameter(Mandatory)][string]$Name
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

        $value = Get-AssessmentContainerValue -Container $TenantStatsHash -Name $Key
        if ($null -eq $value) {
            $derived = Get-AssessmentContainerValue -Container $TenantStatsHash -Name 'Derived'
            $value = Get-AssessmentContainerValue -Container $derived -Name $Key
        }

        return @(Convert-AssessmentHtmlArray -InputObject $value)
    }

    function Encode-AssessmentHtml {
        param($Value)
        if ($null -eq $Value) { return '' }
        return [System.Web.HttpUtility]::HtmlEncode([string]$Value)
    }

    function Get-AssessmentBadgeClass {
        param($Value)

        switch -Regex ([string]$Value) {
            '^(Critical|Risk|Blocker)$' { return 'pill pill-critical' }
            '^(Warning|Review|Needs Data)$' { return 'pill pill-warning' }
            '^(Ready|Healthy|Pass|Enabled|Complete|Covered)$' { return 'pill pill-good' }
            default { return 'pill pill-neutral' }
        }
    }

    function Format-AssessmentCell {
        param(
            [string]$Column,
            $Value
        )

        if ($null -eq $Value) {
            return ''
        }

        $encodedValue = Encode-AssessmentHtml $Value
        if ($Column -in @('Status', 'Severity')) {
            return "<span class='$(Get-AssessmentBadgeClass -Value $Value)'>$encodedValue</span>"
        }

        if ($Column -eq 'ActionUrl' -and -not [string]::IsNullOrWhiteSpace([string]$Value)) {
            $encodedUrl = [System.Web.HttpUtility]::HtmlAttributeEncode([string]$Value)
            return "<a class='action-link' href='$encodedUrl'>Open recommendation</a>"
        }

        return $encodedValue
    }

    function New-AssessmentTable {
        param(
            [array]$Rows,
            [string[]]$Columns
        )

        if (-not $Rows -or $Rows.Count -eq 0) {
            return "<div class='empty-state'>No data available</div>"
        }

        $html = "<table><thead><tr>"
        foreach ($column in $Columns) {
            $html += "<th>$(Encode-AssessmentHtml $column)</th>"
        }
        $html += "</tr></thead><tbody>"

        foreach ($row in $Rows) {
            $html += '<tr>'
            foreach ($column in $Columns) {
                $property = $row.PSObject.Properties[$column]
                $value = if ($property) { $property.Value } else { $null }
                $html += "<td>$(Format-AssessmentCell -Column $column -Value $value)</td>"
            }
            $html += '</tr>'
        }

        $html += '</tbody></table>'
        return $html
    }

    function New-AssessmentList {
        param(
            [array]$Rows,
            [scriptblock]$ContentScript,
            [string]$EmptyMessage
        )

        if (-not $Rows -or $Rows.Count -eq 0) {
            return "<div class='empty-state'>$([System.Web.HttpUtility]::HtmlEncode($EmptyMessage))</div>"
        }

        $items = foreach ($row in $Rows) {
            & $ContentScript $row
        }

        return "<ul class='takeaway-list'>$($items -join '')</ul>"
    }

    function Convert-ToAreaAnchor {
        param([string]$AreaName)

        if ([string]::IsNullOrWhiteSpace($AreaName)) {
            return 'findings-uncategorized'
        }

        $normalized = ($AreaName.ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-')
        if ([string]::IsNullOrWhiteSpace($normalized)) {
            return 'findings-uncategorized'
        }
        return "findings-$normalized"
    }

    $bestPractices = ConvertTo-AssessmentArray -Key 'BestPractices'
    $findings = ConvertTo-AssessmentArray -Key 'BestPracticeFindings'
    $secureScoreActions = ConvertTo-AssessmentArray -Key 'SecureScoreActions'
    $ownershipSummaryRows = ConvertTo-AssessmentArray -Key 'OwnershipGovernanceSummary'
    $ownershipSummary = if ($ownershipSummaryRows.Count -gt 0) { $ownershipSummaryRows | Select-Object -First 1 } else { $null }
    $unmanagedObjects = ConvertTo-AssessmentArray -Key 'UnmanagedObjects'
    $oneDriveOwnerMismatches = ConvertTo-AssessmentArray -Key 'OneDriveOwnerMismatches'

    if ($bestPractices.Count -eq 0 -and $findings.Count -eq 0) {
        return [PSCustomObject]@{
            Success = $false
            OutputPath = $null
            Error = 'Assessment tables were not present in the tenant stats hash.'
        }
    }

    $criticalAreas = @($bestPractices | Where-Object { $_.Status -eq 'Critical' }).Count
    $warningAreas = @($bestPractices | Where-Object { $_.Status -eq 'Warning' }).Count
    $riskFindings = @($findings | Where-Object { $_.Severity -eq 'Risk' }).Count
    $warningFindings = @($findings | Where-Object { $_.Severity -eq 'Warning' }).Count
    $openFindings = $riskFindings + $warningFindings

    $priorityFindings = @(
        $findings |
            Sort-Object @{ Expression = {
                switch ($_.Severity) {
                    'Risk' { 1 }
                    'Warning' { 2 }
                    'Info' { 3 }
                    default { 9 }
                }
            } }, Priority |
            Select-Object -First 6
    )

    $bestPracticeRows = @(
        $bestPractices |
            Select-Object Area, Status, CriticalFindings, WarningFindings, InfoFindings, TotalFindings, PrimaryFinding, RecommendedAction
    )
    $findingRows = @(
        $findings |
            Sort-Object @{ Expression = {
                switch ($_.Severity) {
                    'Risk' { 1 }
                    'Warning' { 2 }
                    'Info' { 3 }
                    default { 9 }
                }
            } }, Priority |
            Select-Object Severity, Area, Category, Message, RecommendedAction, Priority
    )
    $secureScoreRows = @(
        $secureScoreActions |
            Sort-Object Rank, ScoreGap |
            Select-Object -First 8 RecommendationTitle, Status, ScoreGap, ActionUrl
    )
    $ownershipSnapshotRows = @(
        [PSCustomObject]@{ Metric = 'Unmanaged Objects'; Value = if ($ownershipSummary -and $ownershipSummary.PSObject.Properties['UnmanagedObjectCount']) { $ownershipSummary.UnmanagedObjectCount } else { $unmanagedObjects.Count } },
        [PSCustomObject]@{ Metric = 'Missing Owner'; Value = if ($ownershipSummary -and $ownershipSummary.PSObject.Properties['MissingOwnerCount']) { $ownershipSummary.MissingOwnerCount } else { @($unmanagedObjects | Where-Object { $_.OwnerState -eq 'Missing' }).Count } },
        [PSCustomObject]@{ Metric = 'Owner Health Risks'; Value = if ($ownershipSummary -and $ownershipSummary.PSObject.Properties['OwnerHealthRiskCount']) { $ownershipSummary.OwnerHealthRiskCount } else { @($unmanagedObjects | Where-Object { $_.OwnerState -in @('Disabled', 'Stale', 'Mixed') }).Count } },
        [PSCustomObject]@{ Metric = 'OneDrive Owner Mismatches'; Value = if ($ownershipSummary -and $ownershipSummary.PSObject.Properties['OneDriveOwnerMismatchCount']) { $ownershipSummary.OneDriveOwnerMismatchCount } else { $oneDriveOwnerMismatches.Count } },
        [PSCustomObject]@{ Metric = 'Unknown Owner State'; Value = if ($ownershipSummary -and $ownershipSummary.PSObject.Properties['UnknownOwnerStateCount']) { $ownershipSummary.UnknownOwnerStateCount } else { 0 } }
    )
    $unmanagedPreviewRows = @(
        $unmanagedObjects |
            Sort-Object @{ Expression = {
                switch ($_.Severity) {
                    'Risk' { 1 }
                    'Warning' { 2 }
                    default { 3 }
                }
            } }, Workload, DisplayName |
            Select-Object -First 10 Workload, ObjectType, DisplayName, OwnerState, Reason
    )
    $oneDriveMismatchPreviewRows = @(
        $oneDriveOwnerMismatches |
            Sort-Object DisplayName, SiteUrl |
            Select-Object -First 10 DisplayName, SiteUrl, CurrentOwner, ExpectedDefaultOwner
    )

    $displayFindingRows = @($findingRows)
    $groupedFindingAreas = @(
        $displayFindingRows |
            Group-Object Area |
            Sort-Object Name |
            ForEach-Object {
                $areaRows = @(
                    $_.Group |
                        Sort-Object @{ Expression = {
                            switch ([string]$_.Severity) {
                                'Risk' { 1 }
                                'Warning' { 2 }
                                'Info' { 3 }
                                default { 9 }
                            }
                        } }, Priority
                )
                $areaName = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Uncategorized' } else { $_.Name }
                [PSCustomObject]@{
                    Area = $areaName
                    Anchor = Convert-ToAreaAnchor -AreaName $areaName
                    RiskCount = @($areaRows | Where-Object { $_.Severity -eq 'Risk' }).Count
                    WarningCount = @($areaRows | Where-Object { $_.Severity -eq 'Warning' }).Count
                    InfoCount = @($areaRows | Where-Object { $_.Severity -eq 'Info' }).Count
                    TotalCount = $areaRows.Count
                    Rows = $areaRows
                }
            }
    )

    $groupedFindingNavHtml = if ($groupedFindingAreas.Count -gt 0) {
        $items = foreach ($areaGroup in $groupedFindingAreas) {
            $areaText = Encode-AssessmentHtml $areaGroup.Area
            "<a class='area-chip' href='#$($areaGroup.Anchor)'><span class='area-chip-name'>$areaText</span><span class='area-chip-count'>$($areaGroup.TotalCount)</span></a>"
        }
        "<div class='area-chip-grid'>$($items -join '')</div>"
    } else {
        "<div class='empty-state'>No findings were generated for this run.</div>"
    }

    $groupedFindingDetailsHtml = if ($groupedFindingAreas.Count -gt 0) {
        $blocks = foreach ($areaGroup in $groupedFindingAreas) {
            $areaText = Encode-AssessmentHtml $areaGroup.Area
            $areaTable = New-AssessmentTable -Rows $areaGroup.Rows -Columns @('Severity','Category','Message','RecommendedAction')
            @"
<div class='finding-area-block' id='$($areaGroup.Anchor)'>
    <div class='finding-area-head'>
        <h3>$areaText</h3>
        <div class='finding-area-counts'>
            <span class='pill pill-critical'>Risk: $($areaGroup.RiskCount)</span>
            <span class='pill pill-warning'>Warning: $($areaGroup.WarningCount)</span>
            <span class='pill pill-neutral'>Info: $($areaGroup.InfoCount)</span>
        </div>
    </div>
    $areaTable
</div>
"@
        }
        $blocks -join ''
    } else {
        ''
    }

    function Get-RoadmapPhaseFromSeverity {
        param([string]$Severity)
        switch ($Severity) {
            'Risk' { '0-30 Days' }
            'Warning' { '31-60 Days' }
            default { '61-90 Days' }
        }
    }

    function Get-TargetStateFromStatus {
        param([string]$Status)
        switch ($Status) {
            'Critical' { 'Healthy / Controlled' }
            'Warning' { 'Healthy / Controlled' }
            default { 'Maintain / Monitor' }
        }
    }

    function Get-GapNarrativeFromStatus {
        param([string]$Status)
        switch ($Status) {
            'Critical' { 'Material gap from Microsoft best practice baseline' }
            'Warning' { 'Partial coverage or inconsistent control implementation' }
            default { 'No immediate gap requiring escalation' }
        }
    }

    $realityRows = @(
        $bestPractices | ForEach-Object {
            $phase = switch ([string]$_.Status) {
                'Critical' { '0-30 Days' }
                'Warning' { '31-60 Days' }
                default { '61-90 Days' }
            }

            [PSCustomObject]@{
                Area = $_.Area
                CurrentState = $_.Status
                BestPracticeTarget = Get-TargetStateFromStatus -Status $_.Status
                Gap = Get-GapNarrativeFromStatus -Status $_.Status
                RoadmapPhase = $phase
                NextAction = $_.RecommendedAction
            }
        }
    )

    $roadmapRows = @(
        $findings |
            Sort-Object @{ Expression = {
                switch ($_.Severity) {
                    'Risk' { 1 }
                    'Warning' { 2 }
                    default { 3 }
                }
            } }, Priority |
            Select-Object -First 9 |
            ForEach-Object {
                [PSCustomObject]@{
                    Phase = Get-RoadmapPhaseFromSeverity -Severity $_.Severity
                    Area = $_.Area
                    Focus = $_.Message
                    Outcome = $_.RecommendedAction
                }
            }
    )

    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $tenantIdentity = Get-TenantIdentityDetails -TenantStatsHash $TenantStatsHash
    $tenantName = [string]$tenantIdentity.DisplayName
    if ([string]::IsNullOrWhiteSpace($tenantName)) {
        $tenantName = 'Tenant Assessment'
    }
    $tenantId = [string]$tenantIdentity.TenantId
    $tenantDefaultDomain = [string]$tenantIdentity.DefaultDomain

    $actionableAreas = $criticalAreas + $warningAreas
    $executiveHeadline = if ($criticalAreas -gt 0) {
        "$criticalAreas assessment area(s) need immediate attention."
    } elseif ($warningAreas -gt 0) {
        "No critical areas were detected, but $warningAreas area(s) still need follow-up."
    } else {
        "No critical or warning-level assessment areas were detected in this run."
    }

    $priorityFindingList = New-AssessmentList -Rows $priorityFindings -EmptyMessage 'No high-priority findings were identified.' -ContentScript {
        param($row)
        $area = Encode-AssessmentHtml $row.Area
        $message = Encode-AssessmentHtml $row.Message
        $action = Encode-AssessmentHtml $row.RecommendedAction
        $badge = Format-AssessmentCell -Column 'Severity' -Value $row.Severity
        "<li><div class='takeaway-head'>$badge <span>$area</span></div><div class='takeaway-body'>$message</div><div class='takeaway-action'>$action</div></li>"
    }

    $roadmapList = New-AssessmentList -Rows $roadmapRows -EmptyMessage 'No roadmap actions were identified from findings.' -ContentScript {
        param($row)
        $phase = Encode-AssessmentHtml $row.Phase
        $focus = Encode-AssessmentHtml $row.Focus
        $outcome = Encode-AssessmentHtml $row.Outcome
        $area = Encode-AssessmentHtml $row.Area
        "<li><div class='takeaway-head'><span class='pill pill-neutral'>$phase</span> <span>$area</span></div><div class='takeaway-body'>$focus</div><div class='takeaway-action'>$outcome</div></li>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$tenantName - Best Practices Snapshot | Arraya Solutions</title>
    <style>
        :root {
            --bg: #f3f0e8;
            --paper: #fffdf8;
            --ink: #1f2933;
            --muted: #5b6872;
            --line: #ddd5c7;
            --accent: #0f4c5c;
            --accent-2: #d97706;
            --critical: #b42318;
            --warning: #b54708;
            --good: #067647;
            --neutral: #475467;
        }
        body { font-family: Georgia, 'Times New Roman', serif; margin: 0; background: radial-gradient(circle at top, #faf8f3 0%, var(--bg) 60%, #ece7db 100%); color: var(--ink); }
        .wrap { max-width: 1320px; margin: 0 auto; padding: 34px 28px 40px; }
        .hero { background: linear-gradient(145deg, #12343b 0%, #1e5160 55%, #c77d2b 140%); color: white; padding: 34px; border-radius: 24px; box-shadow: 0 20px 40px rgba(22, 34, 42, 0.18); }
        .hero-top { display: flex; justify-content: space-between; gap: 24px; align-items: flex-start; }
        .eyebrow { text-transform: uppercase; letter-spacing: 0.16em; font: 600 11px/1.4 Segoe UI, Arial, sans-serif; opacity: 0.85; margin-bottom: 10px; }
        .hero h1 { margin: 0; font-size: 38px; line-height: 1.05; }
        .hero-meta { font: 500 12px/1.5 Segoe UI, Arial, sans-serif; text-align: right; opacity: 0.9; }
        .hero-subtitle { font: 500 15px/1.6 Segoe UI, Arial, sans-serif; margin: 14px 0 0; max-width: 840px; opacity: 0.96; }
        .headline { margin-top: 18px; padding: 16px 18px; background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.16); border-radius: 18px; font: 600 18px/1.4 Segoe UI, Arial, sans-serif; }
        .grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 16px; margin: 24px 0; }
        .card { background: var(--paper); border: 1px solid rgba(18, 52, 59, 0.08); border-radius: 18px; padding: 20px; box-shadow: 0 10px 24px rgba(29, 41, 57, 0.06); }
        .card .label { font: 600 11px/1.4 Segoe UI, Arial, sans-serif; text-transform: uppercase; letter-spacing: 0.12em; color: var(--muted); margin-bottom: 10px; }
        .card .value { font-size: 34px; font-weight: 700; color: #111827; font-family: 'Segoe UI', Arial, sans-serif; }
        .card .caption { margin-top: 6px; color: var(--muted); font: 500 12px/1.5 Segoe UI, Arial, sans-serif; }
        .summary-grid { display: grid; grid-template-columns: 1.1fr 0.9fr; gap: 18px; margin-top: 10px; }
        .section { background: var(--paper); border: 1px solid rgba(18, 52, 59, 0.08); border-radius: 22px; padding: 24px; margin-top: 22px; box-shadow: 0 10px 24px rgba(29, 41, 57, 0.06); }
        .section h2 { margin: 0 0 8px; font-size: 24px; }
        .section p.note { color: var(--muted); margin: 0; font: 500 13px/1.6 Segoe UI, Arial, sans-serif; }
        .section-header { display: flex; justify-content: space-between; gap: 16px; align-items: baseline; margin-bottom: 14px; }
        .section-kicker { color: var(--accent); font: 700 11px/1.4 Segoe UI, Arial, sans-serif; text-transform: uppercase; letter-spacing: 0.14em; }
        table { width: 100%; border-collapse: collapse; margin-top: 16px; font-family: Segoe UI, Arial, sans-serif; }
        th, td { text-align: left; padding: 12px; border-bottom: 1px solid var(--line); vertical-align: top; font-size: 13px; }
        th { background: #f4efe5; color: #214654; font-weight: 700; }
        tr:nth-child(even) td { background: #fffaf0; }
        .empty-state { padding: 14px; border: 1px dashed #c7bfb0; border-radius: 14px; color: var(--muted); background: #fbf8f1; font: 500 13px/1.6 Segoe UI, Arial, sans-serif; }
        .pill { display: inline-flex; align-items: center; border-radius: 999px; padding: 4px 10px; font: 700 11px/1 Segoe UI, Arial, sans-serif; text-transform: uppercase; letter-spacing: 0.06em; white-space: nowrap; }
        .pill-critical { background: #fef3f2; color: var(--critical); border: 1px solid #fecdca; }
        .pill-warning { background: #fffaeb; color: var(--warning); border: 1px solid #fedf89; }
        .pill-good { background: #ecfdf3; color: var(--good); border: 1px solid #abefc6; }
        .pill-neutral { background: #f2f4f7; color: var(--neutral); border: 1px solid #d0d5dd; }
        .takeaway-list { list-style: none; padding: 0; margin: 16px 0 0; display: grid; gap: 12px; }
        .takeaway-list li { border: 1px solid var(--line); background: #fffaf2; border-radius: 16px; padding: 14px 16px; }
        .takeaway-head { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; font: 700 14px/1.5 Segoe UI, Arial, sans-serif; }
        .takeaway-body { font: 500 13px/1.6 Segoe UI, Arial, sans-serif; color: var(--ink); }
        .takeaway-action { margin-top: 8px; font: 600 12px/1.6 Segoe UI, Arial, sans-serif; color: var(--accent); }
        .area-chip-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 10px; margin-top: 16px; }
        .area-chip { display: flex; justify-content: space-between; align-items: center; gap: 12px; border: 1px solid var(--line); border-radius: 14px; padding: 10px 12px; text-decoration: none; background: #fffaf2; color: var(--ink); font: 600 12px/1.4 Segoe UI, Arial, sans-serif; }
        .area-chip:hover { border-color: #c9bba4; background: #fff6e9; }
        .area-chip-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .area-chip-count { display: inline-flex; align-items: center; justify-content: center; min-width: 26px; height: 22px; border-radius: 999px; background: #f2e7d4; color: #7a4f10; font-weight: 700; padding: 0 8px; }
        .finding-area-block { margin-top: 18px; border: 1px solid var(--line); border-radius: 16px; padding: 14px; background: #fffdf8; }
        .finding-area-head { display: flex; justify-content: space-between; gap: 10px; align-items: center; margin-bottom: 4px; flex-wrap: wrap; }
        .finding-area-head h3 { margin: 0; font-size: 18px; }
        .finding-area-counts { display: flex; gap: 8px; flex-wrap: wrap; }
        .action-link { color: var(--accent); text-decoration: none; font-weight: 600; }
        .action-link:hover { text-decoration: underline; }
        .methodology-note { margin-top: 16px; padding: 16px 18px; border-radius: 16px; background: #f8f3e8; border: 1px solid var(--line); font: 500 13px/1.7 Segoe UI, Arial, sans-serif; color: var(--muted); }
        .footer-note { margin-top: 18px; color: var(--muted); font: 500 12px/1.6 Segoe UI, Arial, sans-serif; }
        @media (max-width: 1100px) { .grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } .summary-grid { grid-template-columns: 1fr; } .hero-top { flex-direction: column; } .hero-meta { text-align: left; } }
        @media (max-width: 720px) { .wrap { padding: 18px; } .grid { grid-template-columns: 1fr; } .hero h1 { font-size: 30px; } .section-header { display: block; } }
        @media print {
            body { background: white; }
            .wrap { max-width: none; padding: 0; }
            .hero, .card, .section { box-shadow: none; border: 1px solid #d7d1c4; }
            .section, .card { break-inside: avoid; page-break-inside: avoid; }
            th { position: static; }
            a { color: inherit; text-decoration: none; }
        }
    </style>
</head>
<body>
    <div class="wrap">
        <div class="hero">
            <div class="hero-top">
                <div>
                    <div class="eyebrow">Prepared by Arraya Solutions</div>
                    <h1>$tenantName</h1>
                    <p class="hero-subtitle">Microsoft 365 Best Practices Snapshot for leadership review. Use the workbook for raw inventory and full worksheet detail.</p>
                </div>
                <div class="hero-meta">
                    <div>Generated: $reportDate</div>
                    <div>Tenant ID: $(if ($tenantId) { Encode-AssessmentHtml $tenantId } else { 'Unavailable' })</div>
                    <div>Default Domain: $(if (-not [string]::IsNullOrWhiteSpace($tenantDefaultDomain)) { Encode-AssessmentHtml $tenantDefaultDomain } else { 'Unavailable' })</div>
                    <div>Artifact: Best Practices Snapshot</div>
                    <div>Prepared by: Arraya Solutions</div>
                </div>
            </div>
            <div class="headline">$executiveHeadline</div>
        </div>

        <div class="grid">
            <div class="card"><div class="label">Assessment Areas</div><div class="value">$($bestPractices.Count)</div><div class="caption">Area-level rollups evaluated</div></div>
            <div class="card"><div class="label">Critical Areas</div><div class="value">$criticalAreas</div><div class="caption">Immediate remediation candidates</div></div>
            <div class="card"><div class="label">Actionable Areas</div><div class="value">$actionableAreas</div><div class="caption">Critical or warning areas</div></div>
            <div class="card"><div class="label">Open Findings</div><div class="value">$openFindings</div><div class="caption">Risk and warning findings requiring action</div></div>
        </div>

        <div class="summary-grid">
            <div class="section">
                <div class="section-header">
                    <div>
                        <div class="section-kicker">Executive Summary</div>
                        <h2>Top Priorities</h2>
                    </div>
                </div>
                <p class="note">Highest-priority findings drawn from the detailed assessment table.</p>
                $priorityFindingList
            </div>

            <div class="section">
                <div class="section-header">
                    <div>
                        <div class="section-kicker">Roadmap</div>
                        <h2>90-Day Action Plan</h2>
                    </div>
                </div>
                <p class="note">Prioritized roadmap actions from highest-severity best-practice findings.</p>
                $roadmapList
            </div>
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Governance</div>
                    <h2>Ownership Governance Snapshot</h2>
                </div>
            </div>
            <p class="note">Unmanaged object coverage across OneDrive, SharePoint, Teams, Entra groups, and Exchange groups. OneDrive owner mismatch is tracked separately as a review signal.</p>
            $(New-AssessmentTable -Rows $ownershipSnapshotRows -Columns @('Metric','Value'))
            <h3 style="margin-top:20px;">Top Unmanaged Objects</h3>
            $(New-AssessmentTable -Rows $unmanagedPreviewRows -Columns @('Workload','ObjectType','DisplayName','OwnerState','Reason'))
            <h3 style="margin-top:20px;">Top OneDrive Owner Mismatches</h3>
            $(New-AssessmentTable -Rows $oneDriveMismatchPreviewRows -Columns @('DisplayName','SiteUrl','CurrentOwner','ExpectedDefaultOwner'))
        </div>

        <div class="section">
                <div class="section-header">
                    <div>
                        <div class="section-kicker">Rollup</div>
                        <h2>Best Practices</h2>
                    </div>
                </div>
            <p class="note">One row per assessment area with rolled-up severity counts and top signal categories for overall tenant posture.</p>
            $(New-AssessmentTable -Rows $bestPracticeRows -Columns @('Area','Status','CriticalFindings','WarningFindings','InfoFindings','TotalFindings','PrimaryFinding','RecommendedAction'))
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Reality vs Target</div>
                    <h2>Reality vs Best Practices</h2>
                </div>
            </div>
            <p class="note">Current-state gaps aligned to recommended target state and phased action timing.</p>
            $(New-AssessmentTable -Rows $realityRows -Columns @('Area','CurrentState','BestPracticeTarget','Gap','RoadmapPhase','NextAction'))
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Detail</div>
                    <h2>Best Practice Findings</h2>
                </div>
            </div>
            <p class="note">Showing $($displayFindingRows.Count) detailed findings across $($groupedFindingAreas.Count) assessment area(s), grouped for faster review.</p>
            $groupedFindingNavHtml
            $groupedFindingDetailsHtml
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Execution Plan</div>
                    <h2>Roadmap Details</h2>
                </div>
            </div>
            <p class="note">Phased action plan built directly from priority findings.</p>
            $(New-AssessmentTable -Rows $roadmapRows -Columns @('Phase','Area','Focus','Outcome'))
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Microsoft Guidance</div>
                    <h2>Secure Score Actions</h2>
                </div>
            </div>
            <p class="note">Top mapped Microsoft Secure Score actions for remediation planning.</p>
            $(New-AssessmentTable -Rows $secureScoreRows -Columns @('RecommendationTitle','Status','ScoreGap','ActionUrl'))
            <div class="footer-note">Use the workbook tabs for complete findings, raw inventory, and supporting worksheets.</div>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force

    return [PSCustomObject]@{
        Success = $true
        OutputPath = $OutputPath
        BestPracticesCount = $bestPractices.Count
        FindingsCount = $findings.Count
        MigrationReadinessCount = 0
    }
}

