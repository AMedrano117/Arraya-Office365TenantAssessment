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

function Convert-ToCustomerAssessmentMarkdownHeadingPrefix {
    [CmdletBinding()]
    param([AllowNull()][string]$Style)

    switch ([string]$Style) {
        'Title' { return '# ' }
        'Heading1' { return '## ' }
        'Heading2' { return '### ' }
        'Heading3' { return '#### ' }
        default { return $null }
    }
}

function Convert-ToCustomerAssessmentMarkdownCellText {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$Content)

    if ($null -eq $Content) {
        return 'Not surfaced in current source'
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

        $text = Convert-ToArrayaDisplayText -Value $current -Default ''
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

function Convert-ToCustomerAssessmentNarrativeText {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    $normalized = Convert-ToArrayaDisplayText -Value $Text -Default ''
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $normalized = $normalized -replace '\bcurrent source\b', 'tenant review'
    $normalized = $normalized -replace '\bThis matters in this tenant because\b', 'In this tenant, this matters because'
    $normalized = $normalized -replace '\bWhat is working well is that\b', 'A positive signal in the current state is that'
    return $normalized
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
        return 'Not surfaced in current source'
    }

    $row = @($Observation.ConfigurationRows | Where-Object { [string]$_.Signal -eq $Signal } | Select-Object -First 1)
    if ($null -eq $row -or $row.Count -eq 0) {
        return 'Not surfaced in current source'
    }

    return Convert-ToArrayaDisplayText -Value $row[0].State -Default 'Not surfaced in current source'
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
        default { return 'Not surfaced in current source' }
    }
}

function Get-CustomerRelevantRoadmapActions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$RoadmapActions = @(),
        [Parameter(Mandatory = $false)][string[]]$Workstreams = @(),
        [Parameter(Mandatory = $false)][string[]]$Keywords = @(),
        [Parameter(Mandatory = $false)][int]$Max = 3
    )

    $filtered = @($RoadmapActions)
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

function Get-CustomerAssessmentAppendixSections {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject]@{
            Title      = 'Appendix: Device Management'
            Intro      = 'These Microsoft references support the device management, compliance, and managed access observations documented in this assessment.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Get started with device compliance policies in Microsoft Intune' -Url 'https://learn.microsoft.com/en-us/intune/intune-service/protect/device-compliance-get-started' -WhyItIsRelevant 'Supports the compliance baseline and managed-device observations in the endpoint review.'),
                (New-CustomerDocumentationReference -Title 'Require compliant or hybrid Microsoft Entra joined device' -Url 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-all-users-device-compliance' -WhyItIsRelevant 'Provides Microsoft guidance for tying device state to access-control enforcement.')
            )
        },
        [pscustomobject]@{
            Title      = 'Appendix: Entra Guest Access Best Practices'
            Intro      = 'These references support the observations related to guest lifecycle, external collaboration, and guest governance.'
            References = @(
                (New-CustomerDocumentationReference -Title 'B2B collaboration fundamentals' -Url 'https://learn.microsoft.com/en-us/entra/external-id/b2b-fundamentals' -WhyItIsRelevant 'Supports guest access governance and external collaboration design.'),
                (New-CustomerDocumentationReference -Title 'Overview of external sharing in SharePoint and OneDrive' -Url 'https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview' -WhyItIsRelevant 'Provides Microsoft guidance for the collaboration-sharing observations in this report.')
            )
        },
        [pscustomobject]@{
            Title      = 'Appendix: Application Consent and Authentication Methods'
            Intro      = 'These Microsoft references support the application governance, consent, and authentication-method observations in the tenant review.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Configure the admin consent workflow' -Url 'https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-admin-consent-workflow' -WhyItIsRelevant 'Relevant to application governance and approval workflow maturity.'),
                (New-CustomerDocumentationReference -Title 'How to manage authentication methods' -Url 'https://learn.microsoft.com/en-us/entra/identity/authentication/how-to-authentication-methods-manage' -WhyItIsRelevant 'Supports recommendations tied to MFA registration and authentication-method governance.')
            )
        },
        [pscustomobject]@{
            Title      = 'Appendix: Conditional Access and Privileged Identity Management'
            Intro      = 'These references support the Conditional Access, privileged-access, and standing-admin observations identified in the assessment.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Conditional Access overview' -Url 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview' -WhyItIsRelevant 'Supports the policy coverage, exclusions, and enforcement observations.'),
                (New-CustomerDocumentationReference -Title 'Privileged Identity Management overview' -Url 'https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure' -WhyItIsRelevant 'Supports the recommendations related to privileged role hygiene and reducing standing access.')
            )
        },
        [pscustomobject]@{
            Title      = 'Appendix: Exchange Online Archives and SMTP Relay'
            Intro      = 'These references support the messaging, forwarding, transport, archive, and relay observations documented in the Exchange review.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Control automatic external email forwarding in Microsoft 365' -Url 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-policies-external-email-forwarding' -WhyItIsRelevant 'Supports the forwarding-control observations in the messaging section.'),
                (New-CustomerDocumentationReference -Title 'How to set up a multifunction device or application to send email using Microsoft 365 or Office 365' -Url 'https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-microsoft-365-or-office-365' -WhyItIsRelevant 'Relevant to SMTP relay, connector, and mail-flow exception handling.')
            )
        },
        [pscustomobject]@{
            Title      = 'Appendix: Microsoft Teams and SharePoint Online'
            Intro      = 'These references support the collaboration observations related to ownership, lifecycle, and sharing controls.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Manage who can create Microsoft 365 Groups' -Url 'https://learn.microsoft.com/en-us/microsoft-365/solutions/manage-creation-of-groups' -WhyItIsRelevant 'Supports governance of Teams-connected groups and workspace sprawl.'),
                (New-CustomerDocumentationReference -Title 'Set expiration for Microsoft 365 groups' -Url 'https://learn.microsoft.com/en-us/entra/identity/users/groups-lifecycle' -WhyItIsRelevant 'Relevant to dormant collaboration spaces and lifecycle control.')
            )
        },
        [pscustomobject]@{
            Title      = 'Appendix: DNS DMARC and OneDrive'
            Intro      = 'These references support the domain-authentication and OneDrive lifecycle observations in this assessment.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Set up SPF in Microsoft 365 to help prevent spoofing' -Url 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/set-up-spf-in-office-365-to-help-prevent-spoofing' -WhyItIsRelevant 'Supports SPF and anti-spoofing guidance for the reviewed domains.'),
                (New-CustomerDocumentationReference -Title 'Use DKIM to validate outbound email sent from your custom domain' -Url 'https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dkim-configure' -WhyItIsRelevant 'Supports the observed mail-authentication posture for custom domains.'),
                (New-CustomerDocumentationReference -Title 'Retention and deletion in OneDrive and SharePoint' -Url 'https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion' -WhyItIsRelevant 'Relevant to stale OneDrive and SharePoint lifecycle handling.')
            )
        },
        [pscustomobject]@{
            Title      = 'Appendix: Retention Policies and Data Loss Prevention'
            Intro      = 'These references support the current-state observations around retention visibility, data lifecycle governance, and DLP maturity.'
            References = @(
                (New-CustomerDocumentationReference -Title 'Learn about retention policies and retention labels' -Url 'https://learn.microsoft.com/en-us/purview/retention' -WhyItIsRelevant 'Supports the retention-policy observations and the need for documented lifecycle controls.'),
                (New-CustomerDocumentationReference -Title 'Learn about data loss prevention' -Url 'https://learn.microsoft.com/en-us/purview/dlp-learn-about-dlp' -WhyItIsRelevant 'Provides Microsoft guidance for DLP and data-protection governance.')
            )
        },
        [pscustomobject]@{
            Title      = 'Appendix: Pass-Through Authentication and Password Writeback'
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
        return @('Not surfaced in current source')
    }

    if (($Content -is [System.Collections.IEnumerable]) -and -not ($Content -is [string])) {
        $items = @()
        foreach ($item in @($Content)) {
            $items += @(Convert-ToCustomerWordCellParagraphs -Content $item)
        }
        return @($items | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    $text = Convert-ToArrayaDisplayText -Value $Content -Default 'Not surfaced in current source'
    $paragraphs = @(
        $text -split "(`r`n|`n|`r)" |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($paragraphs.Count -eq 0) {
        return @('Not surfaced in current source')
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

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($TemplatePath)
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

    $documentXml = New-CustomerAssessmentDocumentXml -Blocks $Blocks -TemplatePath $TemplatePath
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
                $cells.Add('Not surfaced in current source') | Out-Null
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
    $roadmapActions = @($SourceModel.RoadmapActions)
    $executiveThemes = @($SourceModel.ExecutiveThemes)
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
    $externalSharingSummaryRecord = if ($Signals.ExternalSharingSummary) { Get-ArrayaObjectValue -Object $Signals.ExternalSharingSummary -Names @('Summary') } else { $null }
    $externalIdentityRestrictionsRecord = if ($Signals.ExternalIdentityRestrictions) { Get-ArrayaObjectValue -Object $Signals.ExternalIdentityRestrictions -Names @('Summary') } else { $null }
    $guestAccessConfigurationRecord = if ($Signals.GuestAccessConfiguration) { Get-ArrayaObjectValue -Object $Signals.GuestAccessConfiguration -Names @('Summary') } else { $null }
    $authenticationConfigRecord = if ($Signals.AuthenticationConfig) { Get-ArrayaObjectValue -Object $Signals.AuthenticationConfig -Names @('Configuration') } else { $null }
    $externalSharingSiteOverrides = Convert-ArrayaObjectToArray $Signals.ExternalSharingSiteOverrides
    $retentionPolicyRows = Convert-ArrayaObjectToArray $Signals.RetentionPolicies
    $dlpPolicyRows = Convert-ArrayaObjectToArray $Signals.DlpPolicies

    $guestUsers = @($userRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('UserType'))).ToLowerInvariant() -eq 'guest' })
    $memberUsers = @($userRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('UserType'))).ToLowerInvariant() -ne 'guest' })
    $enabledMemberUsers = @($memberUsers | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))) -ne $false })
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
    $totalEnterpriseApps = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('TotalEnterpriseApplications', 'EnterpriseApplicationCount'))
    $highPrivilegeApps = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $applicationSummaryRecord -Names @('ApplicationsWithHighPrivilege', 'HighPrivilegeApplicationCount'))
    if ($null -eq $totalEnterpriseApps) { $totalEnterpriseApps = $enterpriseApplications.Count }

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
    $guestInvitationControlText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $guestAccessConfigurationRecord -Names @('GuestInvitationControl')) -Default ''
    if ([string]::IsNullOrWhiteSpace($guestInvitationControlText) -or $guestInvitationControlText -eq 'Not surfaced in current source') {
        $guestInvitationControlText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $authenticationConfigRecord -Names @('AllowInvitesFrom')) -Default ''
    }
    if ([string]::IsNullOrWhiteSpace($guestInvitationControlText) -or $guestInvitationControlText -eq 'Not surfaced in current source') {
        $guestInvitationControlText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('AllowInvitesFrom')) -Default 'Not surfaced in current source'
    }
    $crossTenantPartnerCountText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('CrossTenantPartnerCount')) -Default 'Not surfaced in current source'
    $defaultInboundMfaTrustText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('DefaultInboundMfaTrust')) -Default 'Not surfaced in current source'
    $sharingDomainRestrictionModeText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SharingDomainRestrictionMode')) -Default 'Not surfaced in current source'
    $siteOverrideCountText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SiteOverrideCount')) -Default $(if ($externalSharingSiteOverrides.Count -gt 0) { $externalSharingSiteOverrides.Count } else { 'Not surfaced in current source' })
    $largestDomainsByRecipients = @(
        $domainRows |
            Sort-Object { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalDomainRecipients')) } -Descending |
            Select-Object -First 5
    )
    $dmarcEnabledDomains = @($domainRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DmarcConfigured'))) -eq $true }).Count
    $dkimEnabledDomains = @($domainRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DkimConfigured'))) -eq $true }).Count
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
                @(
                    [string]$action.ActionTitle,
                    ('Priority: {0}' -f [string]$action.RoadmapPhase),
                    ('Action: {0}' -f [string]$action.RecommendedNextStep)
                ),
                (Get-CustomerActionImpactLabel -Severity ([string]$action.HighestSeverity)),
                'Not surfaced in current source'
            )
        }
    )

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
    $applicationInventoryRows = if ($enterpriseApplications.Count -gt 0) {
        @(
            foreach ($enterpriseApplication in @($enterpriseApplications | Select-Object -First 8)) {
                $displayName = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('DisplayName')) -Default 'Unnamed application'
                $apiPermissions = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('HighPrivilegePermissions', 'ApiPermissions', 'Permissions')) -Default 'Not surfaced in current source'
                $highPrivilegeCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('HighPrivilegePermissionCount'))
                $delegatedCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $enterpriseApplication -Names @('DelegatedPermissionGrantCount'))
                $observation = if ($null -ne $highPrivilegeCount -and $highPrivilegeCount -gt 0) { 'The review surfaced high-privilege permissions on this application.' } elseif ($null -ne $delegatedCount -and $delegatedCount -gt 0) { 'Delegated permission grants were surfaced in the current review.' } else { 'Detailed usage telemetry was not surfaced in the current source.' }
                New-CustomerWordTableRow -Cells @($displayName, $apiPermissions, $observation)
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
    }
    $globalAdminTableRows = if ($gaRows.Count -gt 0) {
        @(
            foreach ($globalAdmin in @($gaRows | Select-Object -First 15)) {
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $globalAdmin -Names @('DisplayName')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $globalAdmin -Names @('CreatedDateTime', 'Created')) -Default 'Not surfaced in current source'),
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
                New-CustomerWordTableRow -Cells @(
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('Domain', 'Id')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('IsVerified', 'Verified')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('AuthenticationType')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('DomainType')) -Default 'Not surfaced in current source'),
                    (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $domainRow -Names @('IsDefault', 'Default')) -Default 'Not surfaced in current source')
                )
            }
        )
    }
    else {
        @((New-CustomerWordTableRow -Cells @('Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source', 'Not surfaced in current source')))
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
    $identityActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Identity') -Max 3
    $endpointActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Endpoint') -Max 3
    $messagingActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Messaging') -Max 3
    $collaborationActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Collaboration') -Max 3
    $governanceActions = Get-CustomerRelevantRoadmapActions -RoadmapActions $roadmapActions -Workstreams @('Governance') -Max 3

    $blocks = New-Object System.Collections.Generic.List[object]
    $blocks.Add((New-CustomerWordParagraphBlock -Text "$tenantName Microsoft 365 Tenant Best Practices Assessment" -Style 'Title')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Prepared by: Arraya Solutions' -Style 'Subtitle')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("Generated: {0}" -f $GeneratedAt.ToString('yyyy-MM-dd')) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Version History' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This section tracks the issued version of the assessment report so the customer-facing document has a clear revision record.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Date', 'Revision', 'Author', 'Description', 'Reviewers') -Rows @(
        (New-CustomerWordTableRow -Cells @($GeneratedAt.ToString('MM/dd/yyyy'), '1.0', 'Arraya Solutions', 'Initial assessment report', 'Not surfaced in current source'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Introduction' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text "This assessment documents the current Microsoft 365 state observed in $tenantName and is intended to show where control implementation, administrative discipline, and ownership governance are holding together and where they are beginning to drift." -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($SourceModel.SummaryText) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Project Scope' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The scope of this report is limited to the Microsoft 365 signals surfaced in the tenant review across identity, devices, messaging, collaboration, governance, and lifecycle controls. The document focuses on the conditions that were visible in the tenant data and maps those conditions into the operating areas most likely to affect security, administration, and day-to-day support.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Where a template field expects a detail that was not visible in the tenant review, the report states that clearly rather than inferring a value that the current source did not support.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Executive Summary' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text "This assessment provides a leadership view of the current Microsoft 365 environment for $tenantName, with specific attention to where findings cluster and what those clusters imply for operational and security risk." -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ($SourceModel.ExecutiveNarrative) -Style 'Normal')) | Out-Null
    foreach ($theme in @($executiveThemes | Select-Object -First 3)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0}. {1} {2}" -f $theme.Theme, (Convert-ToCustomerAssessmentNarrativeText -Text $theme.WhatThisMeans), (Convert-ToCustomerAssessmentNarrativeText -Text $theme.WhyItMatters)) -Style 'Normal')) | Out-Null
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Modern Workplace Recommendations' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The recommendations below retain the existing remediation titles, priorities, and action lines. They are presented in the same order produced by the current assessment model so the customer can see which actions are rising first and why.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Recommendation', 'Criticality / Impact', 'Level of Effort') -Rows $recommendationRows)) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Entra ID Review: User and Device Inventory' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Review of the tenant configuration indicated that identity hygiene and device governance need to be read together in this environment. The same parts of the tenant that are carrying stale privileged access are also the parts of the environment where compliance-driven access control is least mature.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'User Account Summary' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $userSummaryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The user inventory shows {0} internal member account(s) and {1} guest account(s) in the reviewed data. What stands out is not just the count itself, but the number of enabled member accounts that no longer show recent sign-in activity. That pattern usually points to lifecycle drift rather than a single isolated exception." -f $(if ($memberUsers.Count -gt 0 -or $userRows.Count -gt 0) { $memberUsers.Count } else { 'an unconfirmed number of' }), $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { $guestUsers.Count } else { 'an unconfirmed number of' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'User Account Status Comparison' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Account Type', 'Total', 'Enabled', 'Inactive') -Rows $userComparisonRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'When viewed collectively, the identity inventory suggests that access cleanup is not limited to one user population. Internal members and guest identities both show evidence of stale access, which increases the chance that a dormant identity still retains a path into collaboration, messaging, or privileged workflows.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device Registration Summary' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows $deviceSummaryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.ObservedNarrative) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device Platform Distribution' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Platform', 'Count', 'Share') -Rows $platformDistributionRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device Registration and Compliance Gaps' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The reviewed device inventory makes it possible to see where endpoint registration exists without the same level of follow-through in management and compliance. In this tenant, that gap is most visible where stale, unmanaged, or non-compliant devices remain part of the active footprint.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.WhyItMatters) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $deviceObservation.PositiveNarrative) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Entra Guest Access Configuration' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows @(
        @('Inactive guest accounts (>90 days)', $(if ($guestUsers.Count -gt 0 -or $userRows.Count -gt 0) { $inactiveGuestUsers.Count } else { (Get-CustomerObservationState -Observation $lifecycleObservation -Signal 'Inactive guest accounts (>90 days)') })),
        @('Conditional Access guest coverage', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $Signals.ConditionalAccessSummary -Names @('Summary')) -Names @('HasGuestCoverage')) -Default 'Not surfaced in current source')),
        @('Guest invitation control', $guestInvitationControlText),
        @('Cross-tenant partner count', $crossTenantPartnerCountText),
        @('Default inbound MFA trust', $defaultInboundMfaTrustText),
        @('Tenant sharing capability', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $Signals.SharePointSharingSummary -Names @('Summary')) -Names @('TenantSharingCapability')) -Default 'Not surfaced in current source')),
        @('Default sharing link type', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $Signals.SharePointSharingSummary -Names @('Summary')) -Names @('DefaultSharingLinkType')) -Default 'Not surfaced in current source'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The guest access review showed that external collaboration is active enough to require ongoing governance rather than periodic cleanup. Inactive guest accounts and externally permissive sharing signals rarely create visible pain day to day, but they expand the tenant surface area if ownership decisions are delayed. The current guest invitation control is shown as {0}, and the tenant currently has {1} configured cross-tenant partner relationship(s) in the reviewed data." -f $guestInvitationControlText, $crossTenantPartnerCountText) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("This matters because guest access combines identity risk with collaboration exposure. Once guest lifecycle, sharing defaults, and sponsor accountability drift at the same time, it becomes harder to validate which external access is still justified. Default inbound MFA trust is currently shown as {0}, which means cross-tenant trust decisions should be reviewed alongside guest invitation settings rather than as a separate design concern." -f $defaultInboundMfaTrustText) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Entra Applications and Access Review' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The application review looked at enterprise application inventory, consent-related controls, and the degree to which privileged permissions are visible in the current tenant data. The reviewed source surfaced {0} enterprise application(s){1}." -f $(if ($null -ne $totalEnterpriseApps) { $totalEnterpriseApps } else { 'an unconfirmed number of' }), $(if ($null -ne $highPrivilegeApps) { ", including $highPrivilegeApps application(s) with elevated permissions" } else { '' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application Inventory' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('DisplayName', 'API Permissions', 'Observation') -Rows $applicationInventoryRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Sign-In Activity and Security Analysis' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'A notable pattern observed was that application governance pressure in this tenant is showing up alongside broader identity strain. Where privileged permissions are present and approval workflow maturity is not clearly surfaced, the tenant can accumulate enterprise applications that are difficult to rationalize quickly during a security review.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recommendations for Application Governance' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current recommendation set points back to identity governance and least-privilege cleanup. The application inventory should be reviewed in the same operating rhythm as privileged identities and guest access so that stale approval paths and broad app permissions do not remain outside normal ownership review.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Modernizing Authentication: Duo Integration, Conditional Access, and MFA Coverage' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Authentication controls in this tenant are not absent, but they are uneven. The review shows a mix of existing Conditional Access coverage, report-only policies, and exclusions that indicate the baseline has started to form without yet being enforced consistently.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Current State Analysis' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows @(
        @('Conditional Access policies', (Get-CustomerObservationState -Observation $identityObservation -Signal 'Conditional Access policies')),
        @('Report-only Conditional Access policies', (Get-CustomerObservationState -Observation $identityObservation -Signal 'Report-only Conditional Access policies')),
        @('Policies with exclusions', (Get-CustomerObservationState -Observation $identityObservation -Signal 'Policies with exclusions')),
        @('MFA registration rate', (Get-CustomerObservationState -Observation $identityObservation -Signal 'MFA registration rate')),
        @('Risk-based Conditional Access coverage', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $Signals.ConditionalAccessSummary -Names @('Summary')) -Names @('HasRiskBasedCoverage')) -Default 'Not surfaced in current source')),
        @('Compliant-device requirement in summary', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $Signals.ConditionalAccessSummary -Names @('Summary')) -Names @('HasCompliantDeviceRequirement')) -Default 'Not surfaced in current source')),
        @('Security Defaults policy', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $Signals.SecurityDefaultsPolicy -Names @('Summary')) -Names @('IsEnabled')) -Default 'Not surfaced in current source'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $identityObservation.ObservedNarrative) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Key Recommendations' -Style 'Heading2')) | Out-Null
    foreach ($action in @($identityActions)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0}. Action: {1}" -f $action.ActionTitle, $action.RecommendedNextStep) -Style 'Normal')) | Out-Null
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Password Writeback and Self-Service Password Reset' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows @(
        @('Password writeback', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $adConnectSummaryRecord -Names @('PasswordWriteback', 'PasswordWritebackEnabled')) -Default 'Not surfaced in current source')),
        @('Self-service password reset', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $Signals.PasswordLifecycleSummary -Names @('SelfServicePasswordReset', 'SelfServicePasswordResetEnabled')) -Default 'Not surfaced in current source')),
        @('Pass-through authentication', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $adConnectSummaryRecord -Names @('PassThroughAuthentication', 'PassThroughAuthenticationEnabled')) -Default 'Not surfaced in current source')),
        @('Directory synchronization', (Get-CustomerObservationState -Observation $governanceObservation -Signal 'Directory synchronization')),
        @('On-premises last sync', (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $adConnectSummaryRecord -Names @('OnPremisesLastSyncDateTime')) -Default 'Not surfaced in current source')),
        @('Authentication configuration visibility', (Get-CustomerObservationState -Observation $governanceObservation -Signal 'Admin consent workflow'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The tenant review did not surface enough direct configuration detail to confirm the current password writeback or self-service password reset posture. That absence is notable because hybrid identity operations, onboarding, and password recovery workflows are difficult to assess cleanly when these controls are not visible in the reviewed data.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Where password writeback and SSPR are part of the operating model, they should be documented and reviewed alongside authentication controls so the customer can distinguish between policy design gaps and visibility gaps.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Authorization: Admin Access and Role Assignments' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Administrative access stood out in this tenant because the privileged footprint is large enough that stale or excess assignments create more than a theoretical risk. The review showed both the number of Global Administrators and the age of some privileged identities, which is usually a sign that role cleanup has not kept pace with operational change.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Summary of Findings' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Metric', 'Current State') -Rows @(
        @('Total admin accounts reviewed', $(if ($adminRows.Count -gt 0) { $adminRows.Count } else { 'Not surfaced in current source' })),
        @('Global Administrators', (Get-CustomerObservationState -Observation $identityObservation -Signal 'Global Administrator count')),
        @('Stale privileged admins (>180 days)', (Get-CustomerObservationState -Observation $identityObservation -Signal 'Stale privileged admins (>180 days)')),
        @('Example stale privileged identities', (Get-CustomerObservationState -Observation $identityObservation -Signal 'Example stale privileged identities'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'High-Level Recommendations' -Style 'Heading2')) | Out-Null
    foreach ($action in @($identityActions)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0}. Action: {1}" -f $action.ActionTitle, $action.RecommendedNextStep) -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Global Administrator Accounts' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('DisplayName', 'Created', 'UserPrincipalName', 'LastSignIn') -Rows $globalAdminTableRows)) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Exchange Online: Mailboxes and Storage Overview' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'A review was performed on mailbox usage and the distribution of recipient objects. The primary goals of this assessment were to evaluate current storage patterns, identify resources that are no longer active, and document where mailbox lifecycle governance is becoming difficult to manage cleanly.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'High-Level Observations and Recommendations' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $messagingObservation.ObservedNarrative) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $messagingObservation.WhyItMatters) -Style 'Normal')) | Out-Null
    foreach ($action in @($messagingActions)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0}. Action: {1}" -f $action.ActionTitle, $action.RecommendedNextStep) -Style 'Normal')) | Out-Null
    }
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

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'SMTP Relay Usage and Service Accounts' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Outbound email configuration was reviewed to understand whether system-generated mail is still relying on mailbox-based SMTP authentication, connector-based relay, or direct send patterns. This matters because automated sending paths often remain in place long after the original application or device owner changes.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $smtpRelayUsageRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Service Account Activity' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('DisplayName', 'UserPrincipalName', 'Send Count', 'Last Activity') -Rows $smtpRelayServiceAccountRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current source can show outbound activity and SMTP posture, but it does not always attribute every sender cleanly to a dedicated relay service account. Even so, the overlap between SMTP-authenticated accounts and message activity is enough to show whether mail relay is being handled as a defined service pattern or through mailbox accounts that now require ownership review.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Microsoft Teams Governance and Cleanup' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'An analysis of the current Microsoft Teams environment shows how collaboration governance is being carried in practice. Team ownership, dormancy, guest presence, and voice workload signals together provide a clearer picture than a raw Team count alone.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $teamsGovernanceRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $collaborationObservation.ObservedNarrative) -Style 'Normal')) | Out-Null
    foreach ($action in @($collaborationActions | Select-Object -First 3)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0}. Action: {1}" -f $action.ActionTitle, $action.RecommendedNextStep) -Style 'Normal')) | Out-Null
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'SharePoint Online Storage and External Sharing' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("SharePoint and OneDrive storage were reviewed together with external-sharing posture because those signals show whether collaboration growth is still being matched by ownership, lifecycle, and sharing control. The current source includes {0} SharePoint sites and {1} OneDrive locations, which is enough to see where content has continued to accumulate." -f $sharePointRows.Count, $oneDriveRows.Count) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Site Title', 'Owner', 'Storage Used (GB)', 'Last Content Modified') -Rows $sharePointStorageRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("External sharing remains an important part of the collaboration posture in this tenant. Tenant-level sharing is currently shown as {0}, the default sharing link type is {1}, and {2} reviewed SharePoint site(s) surfaced an external-sharing capability in the source data. Team-connected sites account for {3} of the reviewed SharePoint locations, which reinforces how closely SharePoint governance is tied to broader Teams and group ownership patterns." -f (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $Signals.SharePointSharingSummary -Names @('Summary')) -Names @('TenantSharingCapability')) -Default 'Not surfaced in current source'), (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $Signals.SharePointSharingSummary -Names @('Summary')) -Names @('DefaultSharingLinkType')) -Default 'Not surfaced in current source'), $sharePointSitesExternalSharingEnabled, $teamConnectedSites) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The external-sharing summary also shows sharing domain restriction mode as {0}, with {1} site-level sharing override(s) identified in the reviewed site inventory. That combination is important because it shows whether external exposure is being controlled only at the tenant level or is also being shaped materially by site-level exceptions." -f $sharingDomainRestrictionModeText, $siteOverrideCountText) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Retention Policies and Data Loss Prevention' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Configuration Signal', 'Current State') -Rows $retentionRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The retention view is stronger than a simple yes-or-no check, but it still shows a gap between mailbox-level lifecycle controls and a clearly surfaced cross-workload retention or DLP program. In the current source, explicit retention signals are visible on individual mailboxes, while DLP detail remains limited or absent.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This matters because retention and DLP are the point where messaging, collaboration, and compliance expectations converge. If those controls are partially implemented or insufficiently visible, legal, operational, and security outcomes become harder to validate with confidence.' -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Domain Configuration and DNS Overview' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("The tenant review surfaced {0} verified domain(s) out of {1} domain record(s) reviewed. Domain hygiene, DNS-based trust controls, and licensing headroom appear together in this section because domain posture is not just a mail-flow issue in this tenant; it is also a governance signal." -f $verifiedDomains.Count, $(if ($domainRows.Count -gt 0) { $domainRows.Count } else { 'an unconfirmed number of' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Domain', 'Verified', 'Authentication Type', 'Domain Type', 'Default') -Rows $domainOverviewRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Domain Recommendations' -Style 'Heading2')) | Out-Null
    foreach ($action in @($governanceActions | Where-Object { $_.ActionTitle -match 'license|tenant governance' } | Select-Object -First 2)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0}. Action: {1}" -f $action.ActionTitle, $action.RecommendedNextStep) -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'DNS Configuration' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Domain', 'SPF', 'DKIM', 'DMARC', 'Notes') -Rows $dnsRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text ("DMARC is currently configured on {0} domain(s), while DKIM is configured on {1} domain(s). The largest recipient concentrations remain in {2}. That mix is notable because mail-authentication posture and recipient concentration are closely linked in a tenant where most messaging volume stays centered on a small set of accepted domains." -f $dmarcEnabledDomains, $dkimEnabledDomains, $(if ($largestDomainsByRecipients.Count -gt 0) { ((@($largestDomainsByRecipients | Select-Object -First 2 | ForEach-Object { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('Domain', 'Id')) -Default 'Not surfaced in current source' })) -join '; ') } else { 'domains not surfaced clearly enough for comparison' })) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'DNS Recommendations' -Style 'Heading2')) | Out-Null
    foreach ($action in @($messagingActions | Select-Object -First 2)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0}. Action: {1}" -f $action.ActionTitle, $action.RecommendedNextStep) -Style 'Normal')) | Out-Null
    }

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Offboarding Recommendation' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Lifecycle signals in this tenant suggest that access, content, and shared workloads are aging out of active ownership at different rates. That is a common sign that offboarding is being handled tactically across workloads rather than through one consistently governed process.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Key Offboarding Objectives' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The current evidence suggests three objectives should remain central: remove stale access promptly, preserve business data only where ownership is clear, and reclaim licenses and shared workloads that no longer have an active sponsor.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recommended Offboarding Workflow' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'A prescriptive end-to-end offboarding workflow was not surfaced in the current source. The assessment evidence does, however, support the lifecycle actions already prioritized in the recommendation set, especially for stale privileged access, inactive guests, stale collaboration locations, and shared mailboxes without ownership signals.' -Style 'Normal')) | Out-Null
    foreach ($action in @($collaborationActions + $identityActions + $endpointActions | Select-Object -First 4)) {
        $blocks.Add((New-CustomerWordParagraphBlock -Text ("{0}. Action: {1}" -f $action.ActionTitle, $action.RecommendedNextStep) -Style 'Normal')) | Out-Null
    }
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Supporting Observations from Environment Review' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Lifecycle Signal', 'Current State') -Rows $offboardingSupportRows)) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $lifecycleObservation.ObservedNarrative) -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text (Convert-ToCustomerAssessmentNarrativeText -Text $lifecycleObservation.WhyItMatters) -Style 'Normal')) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Appendix: Device Management' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Device management is central to maintaining security and compliance in a tenant where unmanaged and stale devices remain part of the current footprint. The references below support the endpoint observations documented in this assessment.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Get started with device compliance policies in Microsoft Intune - https://learn.microsoft.com/en-us/intune/intune-service/protect/device-compliance-get-started', 'Supports the compliance baseline and managed-device observations in the endpoint review.')),
        (New-CustomerWordTableRow -Cells @('Require compliant or hybrid Microsoft Entra joined device - https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-all-users-device-compliance', 'Provides Microsoft guidance for tying device state to access-control enforcement.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Appendix: Entra Guest Access Best Practices' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Guest access best practices matter in this tenant because external collaboration, inactive guests, and sharing posture are all part of the current risk picture. These references support the guest lifecycle and external-access observations in the report.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('B2B collaboration fundamentals - https://learn.microsoft.com/en-us/entra/external-id/b2b-fundamentals', 'Supports guest access governance and external collaboration design.')),
        (New-CustomerWordTableRow -Cells @('Overview of external sharing in SharePoint and OneDrive - https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview', 'Provides Microsoft guidance for the collaboration-sharing observations in this report.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Appendix: Application Consent and Authentication Methods' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application governance and authentication-method controls both affect how quickly identity exposure can grow in a tenant. The references in this appendix support the application-consent, MFA, and authentication-method observations documented earlier in the assessment.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application User Consent Management' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Application user consent should be managed deliberately wherever users can authorize apps to access organizational data. In a tenant where privileged app permissions and consent workflow maturity are already part of the review, consent governance becomes an important control boundary.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Recommended Resources' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Configure the admin consent workflow - https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-admin-consent-workflow', 'Supports the governance observations around application consent, approval workflow, and control maturity.')),
        (New-CustomerWordTableRow -Cells @('Manage authentication methods - https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-authentication-methods-manage', 'Relevant where app access, MFA, and modern authentication governance are being reviewed together.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Authentication Methods Migration' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Migrating from legacy MFA and SSPR controls to the Authentication Methods policy becomes especially important when the tenant already shows mixed enforcement and uneven registration. The resources below support that transition.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Key Resources for Authentication Migration' -Style 'Heading3')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Manage authentication methods - https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-authentication-methods-manage', 'Supports migration away from legacy MFA and SSPR policy management.')),
        (New-CustomerWordTableRow -Cells @('Self-service password reset deep dive - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks', 'Provides Microsoft guidance for SSPR design and operational implications.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Appendix: Conditional Access and Privileged Identity Management' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Conditional Access and privileged-role governance are two of the strongest levers available for reducing identity risk in this tenant. These references support the policy-state, exclusions, and standing-admin observations documented in the report.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Conditional Access overview - https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview', 'Supports the policy coverage, exclusions, and enforcement observations.')),
        (New-CustomerWordTableRow -Cells @('Privileged Identity Management overview - https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure', 'Supports the recommendations related to privileged role hygiene and reducing standing access.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Appendix: Exchange Online Archives and SMTP Relay' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'The Exchange appendix supports the mailbox-growth, archive, forwarding, and relay observations documented in the assessment. These references are useful where mailbox lifecycle and transport controls are both part of the same remediation path.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Exchange Online Archives and Retention Policies' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Archive usage and mailbox retention are part of the same long-term lifecycle story. The references below support archive enablement, mailbox retention, and storage planning.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Learn about retention policies and retention labels - https://learn.microsoft.com/en-us/purview/retention', 'Supports retention-policy observations and mailbox lifecycle planning.')),
        (New-CustomerWordTableRow -Cells @('On-premises password writeback with self-service password reset - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-writeback', 'Relevant when archive, retention, and lifecycle handling intersect with hybrid identity operations.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Third-Party SMTP Relay Configuration' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Relay and application-based sending paths should be documented as service patterns rather than left attached to mailbox accounts by default. These references support the SMTP relay observations in the messaging review.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Control automatic external email forwarding in Microsoft 365 - https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-policies-external-email-forwarding', 'Supports the forwarding-control observations in the messaging section.')),
        (New-CustomerWordTableRow -Cells @('How to set up a multifunction device or application to send email using Microsoft 365 or Office 365 - https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-microsoft-365-or-office-365', 'Relevant to SMTP relay, connector, and mail-flow exception handling.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Appendix: Microsoft Teams and SharePoint Online' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Teams and SharePoint governance are closely linked in this tenant because collaboration growth, ownership, and sharing posture are moving together. The following Microsoft guidance supports the observations documented in the collaboration sections.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Microsoft Teams Governance' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Teams governance references are most relevant where ownerless workspaces, dormant collaboration spaces, and group-creation controls are part of the current-state review.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Manage who can create Microsoft 365 Groups - https://learn.microsoft.com/en-us/microsoft-365/solutions/manage-creation-of-groups', 'Supports governance of Teams-connected groups and workspace sprawl.')),
        (New-CustomerWordTableRow -Cells @('Set expiration for Microsoft 365 groups - https://learn.microsoft.com/en-us/entra/identity/users/groups-lifecycle', 'Relevant to dormant collaboration spaces and lifecycle control.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'SharePoint Online Collaboration' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'SharePoint and OneDrive guidance is particularly relevant where storage growth, stale content, and external sharing need to be evaluated together rather than as separate issues.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Overview of external sharing in SharePoint and OneDrive - https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview', 'Supports the collaboration-sharing observations in this report.')),
        (New-CustomerWordTableRow -Cells @('Retention and deletion in OneDrive and SharePoint - https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion', 'Relevant to stale OneDrive and SharePoint lifecycle handling.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Appendix: DNS DMARC and OneDrive' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'This appendix supports the domain-authentication, DMARC, and OneDrive lifecycle observations that surfaced during the review of accepted domains and collaboration services.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'DMARC Records' -Style 'Heading2')) | Out-Null
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

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Appendix: Retention Policies and Data Loss Prevention' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Retention and DLP guidance becomes especially important where the tenant shows partial retention visibility, but not enough evidence to confirm a mature cross-workload lifecycle and data-protection program.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Additional Resources' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Learn about retention policies and retention labels - https://learn.microsoft.com/en-us/purview/retention', 'Supports retention-policy observations and the need for documented lifecycle controls.')),
        (New-CustomerWordTableRow -Cells @('Learn about data loss prevention - https://learn.microsoft.com/en-us/purview/dlp-learn-about-dlp', 'Provides Microsoft guidance for DLP and data-protection governance.'))
    ))) | Out-Null

    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Appendix: Pass-Through Authentication and Password Writeback' -Style 'Heading1')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Hybrid identity references are included here because directory synchronization, PTA, password writeback, and SSPR visibility all influence how identity operations can be supported safely and consistently.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Pass-Through Authentication (PTA)' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'PTA guidance is relevant anywhere the tenant depends on on-premises credential validation or is still deciding between hybrid sign-in approaches. These references support the hybrid identity observations surfaced in the report.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Microsoft Entra Connect: Pass-through Authentication - https://learn.microsoft.com/en-us/azure/active-directory/hybrid/how-to-connect-pta', 'Supports pass-through authentication design, agent requirements, and operational considerations.')),
        (New-CustomerWordTableRow -Cells @('Microsoft Entra Connect: User sign-in - https://learn.microsoft.com/en-us/azure/active-directory/hybrid/plan-connect-user-signin', 'Provides comparison guidance for hybrid sign-in options and role selection.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Password Writeback with AD Sync' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Password writeback and SSPR references are included because they directly affect password recovery, hybrid lifecycle handling, and end-user support workflows.' -Style 'Normal')) | Out-Null
    $blocks.Add((New-CustomerWordTableBlock -Headers @('Microsoft Guidance', 'Why It Is Relevant') -Rows @(
        (New-CustomerWordTableRow -Cells @('Self-service password reset deep dive - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-howitworks', 'Supports the password reset and SSPR observations in the report.')),
        (New-CustomerWordTableRow -Cells @('On-premises password writeback with self-service password reset - https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sspr-writeback', 'Relevant to password writeback and hybrid credential-management considerations.')),
        (New-CustomerWordTableRow -Cells @('Enable Microsoft Entra password writeback - https://learn.microsoft.com/en-us/azure/active-directory/authentication/tutorial-enable-sspr-writeback', 'Provides implementation guidance where password writeback is part of the target operating model.'))
    ))) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'References' -Style 'Heading2')) | Out-Null
    $blocks.Add((New-CustomerWordParagraphBlock -Text 'Microsoft documentation references are included within each appendix section so the customer can trace the recommendations and current-state observations back to the relevant Microsoft guidance.' -Style 'Normal')) | Out-Null

    return @($blocks.ToArray())
}
