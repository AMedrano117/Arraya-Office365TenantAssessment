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
                $value = $row.PSObject.Properties[$column].Value
                $html += "<td>$(Encode-AssessmentHtml $value)</td>"
            }
            $html += '</tr>'
        }

        $html += '</tbody></table>'
        return $html
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
    if ($TenantStatsHash.ContainsKey('TenantInfo')) {
        $tenantInfo = $TenantStatsHash['TenantInfo']
        $displayName = $tenantInfo.PSObject.Properties['DisplayName']
        if ($displayName -and $displayName.Value) {
            $tenantName = $displayName.Value
        }
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$tenantName - Assessment Summary</title>
    <style>
        body { font-family: Segoe UI, Arial, sans-serif; margin: 0; background: #f5f7fa; color: #1f2937; }
        .wrap { max-width: 1360px; margin: 0 auto; padding: 32px; }
        .hero { background: linear-gradient(135deg, #0f4c81, #2563eb); color: white; padding: 28px 32px; border-radius: 18px; box-shadow: 0 10px 30px rgba(15,76,129,0.2); }
        .hero h1 { margin: 0 0 6px 0; font-size: 32px; }
        .hero p { margin: 4px 0; opacity: 0.95; }
        .grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 16px; margin: 24px 0; }
        .card { background: white; border-radius: 16px; padding: 18px; box-shadow: 0 6px 20px rgba(15, 23, 42, 0.08); }
        .card .label { font-size: 12px; text-transform: uppercase; letter-spacing: 0.08em; color: #6b7280; margin-bottom: 8px; }
        .card .value { font-size: 28px; font-weight: 700; color: #111827; }
        .section { background: white; border-radius: 18px; padding: 24px; margin-top: 22px; box-shadow: 0 6px 20px rgba(15, 23, 42, 0.08); }
        .section h2 { margin-top: 0; font-size: 22px; }
        .section p.note { color: #4b5563; margin-top: 0; }
        table { width: 100%; border-collapse: collapse; margin-top: 14px; }
        th, td { text-align: left; padding: 10px 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; font-size: 13px; }
        th { background: #eff6ff; color: #1d4ed8; position: sticky; top: 0; }
        .empty-state { padding: 12px; border: 1px dashed #cbd5e1; border-radius: 12px; color: #64748b; background: #f8fafc; }
        @media (max-width: 1100px) { .grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
        @media (max-width: 720px) { .wrap { padding: 18px; } .grid { grid-template-columns: 1fr; } }
    </style>
</head>
<body>
    <div class="wrap">
        <div class="hero">
            <h1>$tenantName Assessment Summary</h1>
            <p>Generated: $reportDate</p>
            <p><strong>BestPractices</strong> is the area-level rollup. <strong>BestPracticeFindings</strong> is the detailed issue list behind that rollup.</p>
        </div>

        <div class="grid">
            <div class="card"><div class="label">Assessment Areas</div><div class="value">$($bestPractices.Count)</div></div>
            <div class="card"><div class="label">Critical Areas</div><div class="value">$criticalAreas</div></div>
            <div class="card"><div class="label">Warning Areas</div><div class="value">$warningAreas</div></div>
            <div class="card"><div class="label">Risk Findings</div><div class="value">$riskFindings</div></div>
            <div class="card"><div class="label">Migration Reviews</div><div class="value">$migrationReviews</div></div>
        </div>

        <div class="section">
            <h2>Best Practices</h2>
            <p class="note">One row per assessment area. This is the rollup view used to understand overall tenant posture at a glance.</p>
            $(New-AssessmentTable -Rows $bestPracticeRows -Columns @('Area','Status','CriticalFindings','WarningFindings','TotalFindings','PrimaryFinding','RecommendedAction'))
        </div>

        <div class="section">
            <h2>Best Practice Findings</h2>
            <p class="note">These are the individual findings that drive the rollup rows above.</p>
            $(New-AssessmentTable -Rows $findingRows -Columns @('Severity','Area','Category','Message','RecommendedAction'))
        </div>

        <div class="section">
            <h2>Migration Readiness</h2>
            <p class="note">Migration-specific blockers, reviews, and readiness notes derived from the tenant inventory.</p>
            $(New-AssessmentTable -Rows $migrationRows -Columns @('Category','Item','Status','Value','MigrationAction'))
        </div>

        <div class="section">
            <h2>Secure Score Actions</h2>
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
