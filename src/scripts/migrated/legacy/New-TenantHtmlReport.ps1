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

        if (-not $TenantStatsHash.ContainsKey($Key) -or $null -eq $TenantStatsHash[$Key]) {
            return @()
        }

        $value = $TenantStatsHash[$Key]
        if ($value -is [hashtable] -or $value -is [System.Collections.Specialized.OrderedDictionary]) {
            return @($value.Values)
        }
        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
            return @($value)
        }
        return @($value)
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

    $bestPractices = ConvertTo-AssessmentArray -Key 'BestPractices'
    $findings = ConvertTo-AssessmentArray -Key 'BestPracticeFindings'
    $migrationReadiness = ConvertTo-AssessmentArray -Key 'MigrationReadiness'
    $secureScoreActions = ConvertTo-AssessmentArray -Key 'SecureScoreActions'

    if ($bestPractices.Count -eq 0 -and $findings.Count -eq 0 -and $migrationReadiness.Count -eq 0) {
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
    $migrationReviews = @($migrationReadiness | Where-Object { $_.Status -in @('Blocker', 'Review', 'Needs Data') }).Count

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
            Select-Object -First 5
    )
    $priorityMigrationItems = @(
        $migrationReadiness |
            Where-Object { $_.Status -in @('Blocker', 'Review', 'Needs Data') } |
            Select-Object -First 5
    )

    $bestPracticeRows = @(
        $bestPractices |
            Select-Object Area, Status, CriticalFindings, WarningFindings, TotalFindings, PrimaryFinding, RecommendedAction
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
            Select-Object Severity, Area, Category, Message, RecommendedAction
    )
    $migrationRows = @(
        $migrationReadiness |
            Select-Object Category, Item, Status, Value, MigrationAction
    )
    $secureScoreRows = @(
        $secureScoreActions |
            Sort-Object Rank, ScoreGap |
            Select-Object -First 20 RecommendationTitle, Status, Rank, ScoreGap, ActionUrl
    )

    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $tenantName = 'Tenant Assessment'
    $tenantId = $null
    if ($TenantStatsHash.ContainsKey('TenantInfo')) {
        $tenantInfo = $TenantStatsHash['TenantInfo']
        $displayName = $tenantInfo.PSObject.Properties['DisplayName']
        if ($displayName -and $displayName.Value) {
            $tenantName = $displayName.Value
        }
        $tenantIdProperty = $tenantInfo.PSObject.Properties['TenantId']
        if ($tenantIdProperty -and $tenantIdProperty.Value) {
            $tenantId = $tenantIdProperty.Value
        }
    }

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

    $priorityMigrationList = New-AssessmentList -Rows $priorityMigrationItems -EmptyMessage 'No migration blockers or review items were identified.' -ContentScript {
        param($row)
        $item = Encode-AssessmentHtml $row.Item
        $value = Encode-AssessmentHtml $row.Value
        $action = Encode-AssessmentHtml $row.MigrationAction
        $badge = Format-AssessmentCell -Column 'Status' -Value $row.Status
        "<li><div class='takeaway-head'>$badge <span>$item</span></div><div class='takeaway-body'>$value</div><div class='takeaway-action'>$action</div></li>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$tenantName - Best Practices Analysis</title>
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
        .grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 16px; margin: 24px 0; }
        .card { background: var(--paper); border: 1px solid rgba(18, 52, 59, 0.08); border-radius: 18px; padding: 20px; box-shadow: 0 10px 24px rgba(29, 41, 57, 0.06); }
        .card .label { font: 600 11px/1.4 Segoe UI, Arial, sans-serif; text-transform: uppercase; letter-spacing: 0.12em; color: var(--muted); margin-bottom: 10px; }
        .card .value { font-size: 34px; font-weight: 700; color: #111827; font-family: 'Segoe UI', Arial, sans-serif; }
        .card .caption { margin-top: 6px; color: var(--muted); font: 500 12px/1.5 Segoe UI, Arial, sans-serif; }
        .summary-grid { display: grid; grid-template-columns: 1.2fr 1fr; gap: 18px; margin-top: 10px; }
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
        .action-link { color: var(--accent); text-decoration: none; font-weight: 600; }
        .action-link:hover { text-decoration: underline; }
        .methodology-note { margin-top: 16px; padding: 16px 18px; border-radius: 16px; background: #f8f3e8; border: 1px solid var(--line); font: 500 13px/1.7 Segoe UI, Arial, sans-serif; color: var(--muted); }
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
                    <div class="eyebrow">Microsoft 365 Best Practices Analysis</div>
                    <h1>$tenantName</h1>
                    <p class="hero-subtitle"><strong>BestPractices</strong> is the area-level rollup. <strong>BestPracticeFindings</strong> is the detailed issue list behind that rollup. This summary is optimized for handoff and decision-making, not raw inventory review.</p>
                </div>
                <div class="hero-meta">
                    <div>Generated: $reportDate</div>
                    <div>Tenant ID: $(if ($tenantId) { Encode-AssessmentHtml $tenantId } else { 'Unavailable' })</div>
                    <div>Artifact: Best Practices Analysis</div>
                </div>
            </div>
            <div class="headline">$executiveHeadline</div>
        </div>

        <div class="grid">
            <div class="card"><div class="label">Assessment Areas</div><div class="value">$($bestPractices.Count)</div><div class="caption">Area-level rollups evaluated</div></div>
            <div class="card"><div class="label">Critical Areas</div><div class="value">$criticalAreas</div><div class="caption">Immediate remediation candidates</div></div>
            <div class="card"><div class="label">Warning Areas</div><div class="value">$warningAreas</div><div class="caption">Needs planned follow-up</div></div>
            <div class="card"><div class="label">Risk Findings</div><div class="value">$riskFindings</div><div class="caption">Detailed high-risk issues</div></div>
            <div class="card"><div class="label">Migration Reviews</div><div class="value">$migrationReviews</div><div class="caption">Migration blockers or review items</div></div>
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
                        <div class="section-kicker">Migration Focus</div>
                        <h2>Readiness Watchlist</h2>
                    </div>
                </div>
                <p class="note">Migration items that still need design review, data collection, or remediation.</p>
                $priorityMigrationList
            </div>
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Methodology</div>
                    <h2>License Utilization Scope</h2>
                </div>
            </div>
            <div class="methodology-note">
                Paid-license utilization excludes free, trial, preview, viral, and known benefit SKUs. It also excludes very large user-based seat pools that greatly exceed tenant user count, because those are typically freemium or bundled benefit inventories rather than true purchased capacity.
            </div>
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Rollup</div>
                    <h2>Best Practices</h2>
                </div>
            </div>
            <p class="note">One row per assessment area. This is the rollup view used to understand overall tenant posture at a glance.</p>
            $(New-AssessmentTable -Rows $bestPracticeRows -Columns @('Area','Status','CriticalFindings','WarningFindings','TotalFindings','PrimaryFinding','RecommendedAction'))
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Detail</div>
                    <h2>Best Practice Findings</h2>
                </div>
            </div>
            <p class="note">These are the individual findings that drive the rollup rows above.</p>
            $(New-AssessmentTable -Rows $findingRows -Columns @('Severity','Area','Category','Message','RecommendedAction'))
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Migration</div>
                    <h2>Migration Readiness</h2>
                </div>
            </div>
            <p class="note">Migration-specific blockers, reviews, and readiness notes derived from the tenant inventory.</p>
            $(New-AssessmentTable -Rows $migrationRows -Columns @('Category','Item','Status','Value','MigrationAction'))
        </div>

        <div class="section">
            <div class="section-header">
                <div>
                    <div class="section-kicker">Microsoft Guidance</div>
                    <h2>Secure Score Actions</h2>
                </div>
            </div>
            <p class="note">Top mapped Microsoft Secure Score actions included for security remediation prioritization.</p>
            $(New-AssessmentTable -Rows $secureScoreRows -Columns @('RecommendationTitle','Status','Rank','ScoreGap','ActionUrl'))
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
        MigrationReadinessCount = $migrationReadiness.Count
    }
}
