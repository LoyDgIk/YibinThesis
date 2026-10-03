#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [string]$PdfOutput,

    [switch]$RefreshOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$documentPath = [System.IO.Path]::GetFullPath($InputPath)
if (-not (Test-Path -LiteralPath $documentPath -PathType Leaf)) {
    throw "DOCX not found: $documentPath"
}
if (-not [string]::IsNullOrWhiteSpace($PdfOutput)) {
    $pdfPath = [System.IO.Path]::GetFullPath($PdfOutput)
    New-Item -ItemType Directory -Force (Split-Path -Parent $pdfPath) | Out-Null
}

function Assert-Near {
    param(
        [string]$Label,
        [double]$Actual,
        [double]$Expected,
        [double]$Tolerance = 0.75
    )
    if ([math]::Abs($Actual - $Expected) -gt $Tolerance) {
        throw "$Label expected $Expected, got $Actual"
    }
}

function Assert-Font {
    param(
        [string]$Label,
        [string]$Actual,
        [string[]]$Allowed
    )
    if ($Allowed -notcontains $Actual) {
        throw "$Label expected one of '$($Allowed -join ', ')', got '$Actual'"
    }
}

function Assert-Text {
    param(
        [string]$Label,
        [string]$Actual,
        [string]$Expected
    )
    if ($Actual -cne $Expected) {
        throw "$Label expected '$Expected', got '$Actual'"
    }
}

function Get-ParagraphText {
    param([Parameter(Mandatory = $true)]$Paragraph)
    return $Paragraph.Range.Text.Trim([char[]]@([char]13, [char]7, [char]12))
}

function Get-StyleName {
    param([Parameter(Mandatory = $true)]$Paragraph)
    try {
        # Range.Style may report the final character style when the paragraph
        # also owns a section break (the last cover field is such a paragraph).
        return [string]$Paragraph.Range.ParagraphStyle.NameLocal
    }
    catch {
        try { return [string]$Paragraph.Range.Style.NameLocal } catch { return '' }
    }
}

function Get-ParagraphFontSize {
    param([Parameter(Mandatory = $true)]$Paragraph)

    $size = [double]$Paragraph.Range.Font.Size
    if ([math]::Abs($size) -gt 1000 -or $size -le 0) {
        $size = [double]$Paragraph.Range.ParagraphStyle.Font.Size
    }
    return $size
}

function Get-DocxCoreProperties {
    param([Parameter(Mandatory = $true)][string]$Path)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry('docProps/core.xml')
        if ($null -eq $entry) {
            return [pscustomobject]@{ Keywords = ''; Subject = '' }
        }

        $stream = $entry.Open()
        try {
            $reader = New-Object System.IO.StreamReader(
                $stream,
                [System.Text.Encoding]::UTF8,
                $true
            )
            try {
                $xml = New-Object System.Xml.XmlDocument
                $xml.PreserveWhitespace = $true
                $xml.LoadXml($reader.ReadToEnd())
            }
            finally {
                $reader.Dispose()
            }
        }
        finally {
            $stream.Dispose()
        }

        $namespaces = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $namespaces.AddNamespace(
            'cp',
            'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        )
        $namespaces.AddNamespace('dc', 'http://purl.org/dc/elements/1.1/')
        $keywords = $xml.SelectSingleNode('/cp:coreProperties/cp:keywords', $namespaces)
        $subject = $xml.SelectSingleNode('/cp:coreProperties/dc:subject', $namespaces)
        return [pscustomobject]@{
            Keywords = if ($null -eq $keywords) { '' } else { [string]$keywords.InnerText }
            Subject = if ($null -eq $subject) { '' } else { [string]$subject.InnerText }
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Resolve-YibinDocumentType {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)]$CoreProperties
    )

    $keywordMatches = [regex]::Matches(
        [string]$CoreProperties.Keywords,
        '(?i)(?:^|;)\s*document-type\s*=\s*([^;]+)'
    )
    if ($keywordMatches.Count -gt 0) {
        $declaredTypes = @(
            $keywordMatches |
                ForEach-Object { $_.Groups[1].Value.Trim().ToLowerInvariant() } |
                Select-Object -Unique
        )
        if ($declaredTypes.Count -ne 1) {
            throw "DOCX CoreProperties contains conflicting document-type values: $($declaredTypes -join ', ')"
        }
        if ($declaredTypes[0] -notin @('thesis', 'proposal', 'literature-review')) {
            throw "Unsupported DOCX document-type: $($declaredTypes[0])"
        }
        return $declaredTypes[0]
    }

    $subject = [string]$CoreProperties.Subject
    if ($subject -match '开题报告') {
        return 'proposal'
    }
    if ($subject -match '文献综述') {
        return 'literature-review'
    }

    $signals = @{
        thesis = 0
        proposal = 0
        'literature-review' = 0
    }
    foreach ($paragraph in @($Document.Paragraphs)) {
        $styleName = Get-StyleName -Paragraph $paragraph
        if ($styleName -match '^宜宾开题-') {
            $signals.proposal++
        }
        elseif ($styleName -match '^宜宾综述-') {
            $signals['literature-review']++
        }
        elseif ($styleName -match '^(?:CoverLogo|CoverThesisType|DeclarationTitle|DeclarationBody|宜宾论文-目录标题)$') {
            $signals.thesis++
        }
    }

    $detected = @(
        $signals.GetEnumerator() |
            Where-Object { $_.Value -gt 0 } |
            ForEach-Object { [string]$_.Key }
    )
    if ($detected.Count -eq 1) {
        return $detected[0]
    }
    if ($detected.Count -gt 1) {
        throw "DOCX applied styles indicate conflicting document types: $($detected -join ', ')"
    }

    Write-Warning 'DOCX has no document-type metadata or profile-specific applied styles; treating it as a legacy thesis document.'
    return 'thesis'
}

function Get-AppliedStyleParagraphs {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][string]$StyleName
    )

    return @(
        foreach ($paragraph in @($Document.Paragraphs)) {
            if ((Get-StyleName -Paragraph $paragraph) -eq $StyleName) {
                $paragraph
            }
        }
    )
}

function Assert-PrimaryFooterPageNumbering {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][int]$SectionIndex,
        [Nullable[int]]$ExpectedStart = $null
    )

    $section = $null
    $footer = $null
    try {
        $section = $Document.Sections.Item($SectionIndex)
        $footer = $section.Footers.Item(1)
        $pageFieldCount = 0
        foreach ($field in @($footer.Range.Fields)) {
            try {
                if ([string]$field.Code.Text -match '(?i)^\s*PAGE(?:\s|$)') {
                    $pageFieldCount++
                }
            }
            finally {
                if ($null -ne $field) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($field)
                }
            }
        }

        if ($null -eq $ExpectedStart) {
            if ($pageFieldCount -ne 0) {
                throw "Section $SectionIndex must not display a PAGE field."
            }
            return
        }
        if ($pageFieldCount -ne 1) {
            throw "Section $SectionIndex PAGE field count expected 1, got $pageFieldCount"
        }
        if (-not [bool]$footer.PageNumbers.RestartNumberingAtSection) {
            throw "Section $SectionIndex must restart page numbering."
        }
        $expectedPageNumber = [int]$ExpectedStart
        if ([int]$footer.PageNumbers.StartingNumber -ne $expectedPageNumber) {
            throw (
                "Section $SectionIndex page number expected to start at " +
                "$expectedPageNumber, got $($footer.PageNumbers.StartingNumber)"
            )
        }
    }
    finally {
        foreach ($comObject in @($footer, $section)) {
            if ($null -ne $comObject) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
            }
        }
    }
}

function Assert-ProposalProfile {
    param([Parameter(Mandatory = $true)]$Document)

    if ($Document.Sections.Count -ne 1) {
        throw "Proposal section count expected 1, got $($Document.Sections.Count)"
    }
    if ($Document.Tables.Count -lt 1) {
        throw "Proposal form table was not found."
    }

    $titles = @(Get-AppliedStyleParagraphs -Document $Document -StyleName '宜宾开题-标题')
    $subtitles = @(Get-AppliedStyleParagraphs -Document $Document -StyleName '宜宾开题-副标题')
    $labels = @(Get-AppliedStyleParagraphs -Document $Document -StyleName '宜宾开题-栏目')
    $bodies = @(Get-AppliedStyleParagraphs -Document $Document -StyleName '宜宾开题-正文')
    $signatures = @(Get-AppliedStyleParagraphs -Document $Document -StyleName '宜宾开题-签名')
    if ($titles.Count -ne 1) { throw "Proposal title style count expected 1, got $($titles.Count)" }
    if ($subtitles.Count -ne 1) { throw "Proposal subtitle style count expected 1, got $($subtitles.Count)" }
    if ($labels.Count -ne 7) { throw "Proposal field-label style count expected 7, got $($labels.Count)" }
    if ($bodies.Count -lt 7) { throw "Proposal body style count expected at least 7, got $($bodies.Count)" }
    if ($signatures.Count -ne 1) { throw "Proposal signature style count expected 1, got $($signatures.Count)" }

    Assert-Text 'Proposal title text' `
        (Get-ParagraphText -Paragraph $titles[0]) `
        '宜宾学院本科毕业论文（设计）开题报告'
    Assert-Font 'Proposal title font' $titles[0].Range.Font.NameFarEast @('楷体', 'KaiTi')
    Assert-Near 'Proposal title size' $titles[0].Range.Font.Size 18
    if ($titles[0].Range.Font.Bold -eq 0) { throw 'Proposal title must be bold.' }
    if ($titles[0].Format.Alignment -ne 1) { throw 'Proposal title must be centered.' }

    Assert-Text 'Proposal subtitle text' `
        (Get-ParagraphText -Paragraph $subtitles[0]) `
        '（学生填写）'
    Assert-Font 'Proposal subtitle font' $subtitles[0].Range.Font.NameFarEast @('宋体', 'SimSun')
    Assert-Near 'Proposal subtitle size' $subtitles[0].Range.Font.Size 12
    Assert-PrimaryFooterPageNumbering -Document $Document -SectionIndex 1 -ExpectedStart 4
}

function Assert-LiteratureReviewProfile {
    param([Parameter(Mandatory = $true)]$Document)

    if ($Document.Sections.Count -ne 2) {
        throw "Literature-review section count expected 2, got $($Document.Sections.Count)"
    }
    if ($Document.Tables.Count -lt 1) {
        throw 'Literature-review information page table was not found.'
    }

    $titles = @(Get-AppliedStyleParagraphs -Document $Document -StyleName '宜宾综述-文档标题')
    $thesisTitles = @(Get-AppliedStyleParagraphs -Document $Document -StyleName '宜宾综述-论文题目')
    $labels = @(Get-AppliedStyleParagraphs -Document $Document -StyleName '宜宾综述-信息标签')
    $values = @(Get-AppliedStyleParagraphs -Document $Document -StyleName '宜宾综述-信息值')
    if ($titles.Count -ne 1) { throw "Literature-review title style count expected 1, got $($titles.Count)" }
    if ($thesisTitles.Count -ne 1) {
        throw "Literature-review thesis-title style count expected 1, got $($thesisTitles.Count)"
    }
    if ($labels.Count -lt 7) {
        throw "Literature-review information-label style count expected at least 7, got $($labels.Count)"
    }
    if ($values.Count -lt 6) {
        throw "Literature-review information-value style count expected at least 6, got $($values.Count)"
    }

    Assert-Text 'Literature-review title text' `
        (Get-ParagraphText -Paragraph $titles[0]) `
        '文献综述'
    Assert-Font 'Literature-review title font' $titles[0].Range.Font.NameFarEast @('黑体', 'SimHei')
    Assert-Near 'Literature-review title size' $titles[0].Range.Font.Size 28
    if ($titles[0].Range.Font.Bold -eq 0) { throw 'Literature-review title must be bold.' }
    if ($titles[0].Format.Alignment -ne 1) { throw 'Literature-review title must be centered.' }

    Assert-Font `
        'Literature-review thesis-title font' `
        $thesisTitles[0].Range.Font.NameFarEast `
        @('黑体', 'SimHei')
    Assert-Near 'Literature-review thesis-title size' $thesisTitles[0].Range.Font.Size 18
    Assert-PrimaryFooterPageNumbering -Document $Document -SectionIndex 1
    Assert-PrimaryFooterPageNumbering -Document $Document -SectionIndex 2 -ExpectedStart 1
}

function Set-StyleFont {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$EastAsia,
        [Parameter(Mandatory = $true)][string]$Latin,
        [Parameter(Mandatory = $true)][double]$Size,
        [Nullable[bool]]$Bold = $null
    )
    try {
        $style = $Document.Styles.Item($Name)
        $style.Font.NameFarEast = $EastAsia
        $style.Font.Name = $Latin
        $style.Font.Size = $Size
        if ($null -ne $Bold) {
            $style.Font.Bold = if ($Bold) { -1 } else { 0 }
        }
    }
    catch {
        # Pandoc does not materialize every optional style in every document.
    }
}

function Set-OfficialTocOoxml {
    param([Parameter(Mandatory = $true)][string]$Path)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $archive = [System.IO.Compression.ZipFile]::Open(
        $Path,
        [System.IO.Compression.ZipArchiveMode]::Update
    )
    try {
        $parts = @('word/styles.xml', 'word/document.xml', 'word/settings.xml')
        foreach ($partName in $parts) {
            $entry = $archive.GetEntry($partName)
            if ($null -eq $entry) {
                throw "DOCX part not found: $partName"
            }

            $stream = $entry.Open()
            try {
                $reader = New-Object System.IO.StreamReader(
                    $stream,
                    [System.Text.Encoding]::UTF8,
                    $true
                )
                try {
                    $xmlText = $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $stream.Dispose()
            }

            $xml = New-Object System.Xml.XmlDocument
            $xml.PreserveWhitespace = $true
            $xml.LoadXml($xmlText)
            $namespaces = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $namespaces.AddNamespace('w', $wordNamespace)

            function New-WElement {
                param(
                    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Owner,
                    [Parameter(Mandatory = $true)][string]$LocalName
                )
                return $Owner.CreateElement('w', $LocalName, $wordNamespace)
            }

            function Ensure-WChild {
                param(
                    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Owner,
                    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Parent,
                    [Parameter(Mandatory = $true)][string]$LocalName
                )
                $node = $Parent.SelectSingleNode("w:$LocalName", $namespaces)
                if ($null -eq $node) {
                    $node = New-WElement -Owner $Owner -LocalName $LocalName
                    [void]$Parent.AppendChild($node)
                }
                return $node
            }

            function Set-WAttribute {
                param(
                    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Owner,
                    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Node,
                    [Parameter(Mandatory = $true)][string]$Name,
                    [Parameter(Mandatory = $true)][string]$Value
                )
                $attribute = $Owner.CreateAttribute('w', $Name, $wordNamespace)
                $attribute.Value = $Value
                [void]$Node.Attributes.SetNamedItem($attribute)
            }

            function Set-TocParagraphProperties {
                param(
                    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Owner,
                    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Properties,
                    [Parameter(Mandatory = $true)][int]$Level,
                    [switch]$IncludeIndent
                )
                $left = @(0, 420, 840)[$Level - 1]
                $leftChars = @(0, 200, 400)[$Level - 1]
                $line = if ($Level -lt 3) { '360' } else { '240' }

                $spacing = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'spacing'
                foreach ($pair in @(
                    @('before', '0'),
                    @('after', '0'),
                    @('line', $line),
                    @('lineRule', 'auto')
                )) {
                    Set-WAttribute -Owner $Owner -Node $spacing -Name $pair[0] -Value $pair[1]
                }

                $alignment = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'jc'
                Set-WAttribute -Owner $Owner -Node $alignment -Name 'val' -Value 'both'

                $tabs = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'tabs'
                $tabs.RemoveAll()
                $tab = New-WElement -Owner $Owner -LocalName 'tab'
                Set-WAttribute -Owner $Owner -Node $tab -Name 'val' -Value 'right'
                Set-WAttribute -Owner $Owner -Node $tab -Name 'leader' -Value 'dot'
                Set-WAttribute -Owner $Owner -Node $tab -Name 'pos' -Value '8777'
                [void]$tabs.AppendChild($tab)

                if ($IncludeIndent) {
                    $indent = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'ind'
                    foreach ($pair in @(
                        @('left', [string]$left),
                        @('leftChars', [string]$leftChars),
                        @('right', '0'),
                        @('rightChars', '0'),
                        @('firstLine', '0'),
                        @('firstLineChars', '0'),
                        @('hanging', '0'),
                        @('hangingChars', '0')
                    )) {
                        Set-WAttribute -Owner $Owner -Node $indent -Name $pair[0] -Value $pair[1]
                    }
                }
            }

            function Set-TocRunProperties {
                param(
                    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Owner,
                    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Properties
                )
                $fonts = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'rFonts'
                foreach ($fontAttribute in @('ascii', 'hAnsi', 'eastAsia', 'cs')) {
                    Set-WAttribute -Owner $Owner -Node $fonts -Name $fontAttribute -Value 'SimSun'
                }
                foreach ($themeAttribute in @(
                    'asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'
                )) {
                    [void]$fonts.RemoveAttribute($themeAttribute, $wordNamespace)
                }
                $size = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'sz'
                Set-WAttribute -Owner $Owner -Node $size -Name 'val' -Value '24'
                $sizeComplex = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'szCs'
                Set-WAttribute -Owner $Owner -Node $sizeComplex -Name 'val' -Value '24'
                $bold = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'b'
                Set-WAttribute -Owner $Owner -Node $bold -Name 'val' -Value '0'
                $boldComplex = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'bCs'
                Set-WAttribute -Owner $Owner -Node $boldComplex -Name 'val' -Value '0'
                $color = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'color'
                Set-WAttribute -Owner $Owner -Node $color -Name 'val' -Value '000000'
                $underline = Ensure-WChild -Owner $Owner -Parent $Properties -LocalName 'u'
                Set-WAttribute -Owner $Owner -Node $underline -Name 'val' -Value 'none'
            }

            function Set-QuickStyleVisible {
                param(
                    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Owner,
                    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Style,
                    [Parameter(Mandatory = $true)][int]$Priority
                )
                foreach ($localName in @('hidden', 'semiHidden', 'unhideWhenUsed', 'autoRedefine')) {
                    $node = $Style.SelectSingleNode("w:$localName", $namespaces)
                    if ($null -ne $node) {
                        [void]$Style.RemoveChild($node)
                    }
                }
                $priorityNode = Ensure-WChild -Owner $Owner -Parent $Style -LocalName 'uiPriority'
                Set-WAttribute -Owner $Owner -Node $priorityNode -Name 'val' -Value ([string]$Priority)
                [void](Ensure-WChild -Owner $Owner -Parent $Style -LocalName 'qFormat')
            }

            function Set-QuickStyleHidden {
                param(
                    [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Owner,
                    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Style
                )
                $quick = $Style.SelectSingleNode('w:qFormat', $namespaces)
                if ($null -ne $quick) {
                    [void]$Style.RemoveChild($quick)
                }
                [void](Ensure-WChild -Owner $Owner -Parent $Style -LocalName 'semiHidden')
            }

            function Clear-DirectoryParagraphOverrides {
                param(
                    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph
                )

                $paragraphProperties = Ensure-WChild `
                    -Owner $xml -Parent $Paragraph -LocalName 'pPr'
                # Directory styles own all visible formatting. Strip Word's
                # regenerated direct overrides so a style edit remains global.
                foreach ($localName in @('spacing', 'ind', 'jc', 'tabs', 'rPr')) {
                    $direct = $paragraphProperties.SelectSingleNode(
                        "w:$localName",
                        $namespaces
                    )
                    if ($null -ne $direct) {
                        [void]$paragraphProperties.RemoveChild($direct)
                    }
                }
                foreach ($run in @($Paragraph.SelectNodes('.//w:r', $namespaces))) {
                    $runProperties = $run.SelectSingleNode('w:rPr', $namespaces)
                    if ($null -eq $runProperties) {
                        continue
                    }
                    foreach ($localName in @('rFonts', 'sz', 'szCs', 'b', 'bCs', 'color', 'u')) {
                        $direct = $runProperties.SelectSingleNode(
                            "w:$localName",
                            $namespaces
                        )
                        if ($null -ne $direct) {
                            [void]$runProperties.RemoveChild($direct)
                        }
                    }
                    if ($runProperties.ChildNodes.Count -eq 0) {
                        [void]$run.RemoveChild($runProperties)
                    }
                }
            }

            if ($partName -eq 'word/styles.xml') {
                $styleIds = @{}
                $captionDirectoryStyleIds = @()
                foreach ($style in @($xml.SelectNodes('//w:style', $namespaces))) {
                    $nameNode = $style.SelectSingleNode('w:name', $namespaces)
                    if ($null -eq $nameNode) {
                        continue
                    }
                    $styleName = $nameNode.GetAttribute('val', $wordNamespace)
                    if ($styleName -match '(?i)^toc\s+([1-3])$') {
                        $level = [int]$Matches[1]
                        $styleId = $style.GetAttribute('styleId', $wordNamespace)
                        $styleIds[$level] = $styleId
                        $paragraphProperties = Ensure-WChild `
                            -Owner $xml -Parent $style -LocalName 'pPr'
                        Set-TocParagraphProperties `
                            -Owner $xml `
                            -Properties $paragraphProperties `
                            -Level $level `
                            -IncludeIndent
                        $runProperties = Ensure-WChild `
                            -Owner $xml -Parent $style -LocalName 'rPr'
                        Set-TocRunProperties -Owner $xml -Properties $runProperties
                        Set-QuickStyleVisible -Owner $xml -Style $style -Priority (38 + $level)
                    }
                    elseif ($styleName -match '(?i)^(?:table of figures|图表目录)$') {
                        $styleId = $style.GetAttribute('styleId', $wordNamespace)
                        $captionDirectoryStyleIds += $styleId
                        $paragraphProperties = Ensure-WChild `
                            -Owner $xml -Parent $style -LocalName 'pPr'
                        Set-TocParagraphProperties `
                            -Owner $xml `
                            -Properties $paragraphProperties `
                            -Level 1 `
                            -IncludeIndent
                        $runProperties = Ensure-WChild `
                            -Owner $xml -Parent $style -LocalName 'rPr'
                        Set-TocRunProperties -Owner $xml -Properties $runProperties
                        Set-QuickStyleVisible -Owner $xml -Style $style -Priority 42
                    }
                    elseif ($styleName -match '(?i)^toc\s+[1-3]\d+$') {
                        Set-QuickStyleHidden -Owner $xml -Style $style
                    }
                    elseif ($styleName -match '(?i)^(?:caption|题注)$') {
                        Set-QuickStyleVisible -Owner $xml -Style $style -Priority 35
                    }
                }
                if ($styleIds.Count -ne 3) {
                    throw 'Word TOC 1/2/3 styles were not found in styles.xml.'
                }
                if ($captionDirectoryStyleIds.Count -eq 0) {
                    throw 'Word Table of Figures style was not found in styles.xml.'
                }
                $script:TocStyleIds = $styleIds
                $script:CaptionDirectoryStyleIds = $captionDirectoryStyleIds
            }
            elseif ($partName -eq 'word/document.xml') {
                if ($null -eq $script:TocStyleIds -or $script:TocStyleIds.Count -ne 3) {
                    throw 'Internal TOC style map was not initialized.'
                }
                foreach ($level in 1..3) {
                    $styleId = [string]$script:TocStyleIds[$level]
                    $paragraphs = @($xml.SelectNodes(
                        "//w:p[w:pPr/w:pStyle[@w:val='$styleId']]",
                        $namespaces
                    ))
                    foreach ($paragraph in $paragraphs) {
                        Clear-DirectoryParagraphOverrides -Paragraph $paragraph
                    }
                }
                if (
                    $null -eq $script:CaptionDirectoryStyleIds -or
                    $script:CaptionDirectoryStyleIds.Count -eq 0
                ) {
                    throw 'Internal Table of Figures style map was not initialized.'
                }
                foreach ($styleId in @($script:CaptionDirectoryStyleIds)) {
                    $paragraphs = @($xml.SelectNodes(
                        "//w:p[w:pPr/w:pStyle[@w:val='$styleId']]",
                        $namespaces
                    ))
                    foreach ($paragraph in $paragraphs) {
                        Clear-DirectoryParagraphOverrides -Paragraph $paragraph
                    }
                }
            }
            elseif ($partName -eq 'word/settings.xml') {
                $settingsRoot = $xml.SelectSingleNode('/w:settings', $namespaces)
                if ($null -eq $settingsRoot) {
                    throw 'Word settings root was not found in settings.xml.'
                }
                $updateFields = Ensure-WChild `
                    -Owner $xml -Parent $settingsRoot -LocalName 'updateFields'
                Set-WAttribute `
                    -Owner $xml -Node $updateFields -Name 'val' -Value 'true'
            }
            else {
                throw "Unsupported DOCX part: $partName"
            }

            $entry.Delete()
            $newEntry = $archive.CreateEntry(
                $partName,
                [System.IO.Compression.CompressionLevel]::Optimal
            )
            $output = $newEntry.Open()
            try {
                $settings = New-Object System.Xml.XmlWriterSettings
                $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
                $settings.Indent = $false
                $writer = [System.Xml.XmlWriter]::Create($output, $settings)
                try {
                    $xml.Save($writer)
                    $writer.Flush()
                }
                finally {
                    $writer.Dispose()
                }
            }
            finally {
                $output.Dispose()
            }
        }
    }
    finally {
        $archive.Dispose()
        Remove-Variable -Name TocStyleIds -Scope Script -ErrorAction SilentlyContinue
        Remove-Variable -Name CaptionDirectoryStyleIds -Scope Script -ErrorAction SilentlyContinue
    }
}

function Close-WordSession {
    param($Document, $Word, [bool]$SaveDocument = $false)
    if ($null -ne $Document) {
        try {
            if ($SaveDocument) {
                $Document.Save()
            }
            $Document.Close($false)
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Document)
        }
    }
    if ($null -ne $Word) {
        try {
            $Word.Quit()
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Word)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Test-LockedCitationField {
    param([Parameter(Mandatory = $true)]$Field)

    try {
        if (-not [bool]$Field.Locked) {
            return $false
        }
    }
    catch {
        return $false
    }

    try {
        # 96 = wdFieldCitation.  The code-text fallback also covers Word
        # versions that do not expose a stable Type value for a new field.
        if ([int]$Field.Type -eq 96) {
            return $true
        }
    }
    catch {}

    try {
        return ([string]$Field.Code.Text) -match '^\s*CITATION(?:\s|$)'
    }
    catch {
        return $false
    }
}

function Update-WordRangeFields {
    param([Parameter(Mandatory = $true)]$Range)

    # Update fields individually so a locked native CITATION field is never
    # included in a bulk Fields.Update call.  Other unlocked fields (TOC,
    # SEQ, REF, PAGEREF, PAGE and NUMPAGES) remain refreshable.
    for ($index = $Range.Fields.Count; $index -ge 1; $index--) {
        $field = $null
        try {
            $field = $Range.Fields.Item($index)
            if (Test-LockedCitationField -Field $field) {
                continue
            }
            [void]$field.Update()
        }
        finally {
            if ($null -ne $field) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($field)
            }
        }
    }
}

function Update-WordDocumentFields {
    param([Parameter(Mandatory = $true)]$Document)

    foreach ($storyType in 1..17) {
        try {
            $range = $Document.StoryRanges.Item($storyType)
            while ($null -ne $range) {
                Update-WordRangeFields -Range $range
                $range = $range.NextStoryRange
            }
        }
        catch {
            # A document does not necessarily contain every Word story type.
        }
    }
    foreach ($toc in @($Document.TablesOfContents)) {
        $toc.Update() | Out-Null
    }
}

function Set-CitationFieldStyles {
    param([Parameter(Mandatory = $true)]$Document)

    try {
        $citationStyle = $Document.Styles.Item('宜宾论文-文献上标')
    }
    catch {
        throw 'Reusable citation style 宜宾论文-文献上标 is missing.'
    }
    foreach ($storyType in 1..17) {
        try {
            $range = $Document.StoryRanges.Item($storyType)
            while ($null -ne $range) {
                foreach ($field in @($range.Fields)) {
                    try {
                        $code = [string]$field.Code.Text
                        if ($code -match '(?i)^\s*(?:REF\s+YibinCitation_|CITATION\s+)') {
                            $field.Code.Style = $citationStyle
                            $field.Result.Style = $citationStyle
                        }
                    }
                    finally {
                        if ($null -ne $field) {
                            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($field)
                        }
                    }
                }
                $range = $range.NextStoryRange
            }
        }
        catch {
            # Not every story type is present in every document.
        }
    }
}

function Format-WordFieldIdentifier {
    param([Parameter(Mandatory = $true)][string]$Name)

    $value = $Name.Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw 'Word CaptionLabel name cannot be empty.'
    }
    if ($value -notmatch '[\s"\\]') {
        return $value
    }
    return '"' + $value.Replace('"', '""') + '"'
}

function Normalize-NativeCaptionSequences {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)]$Word
    )

    $registeredLabels = New-Object System.Collections.Generic.List[object]
    try {
        $desired = @{
            Figure = Format-WordFieldIdentifier -Name '图'
            Table = Format-WordFieldIdentifier -Name '表'
            Equation = Format-WordFieldIdentifier -Name '公式'
        }

        $missing = New-Object System.Collections.Generic.List[string]
        foreach ($name in @('图', '表', '公式')) {
            try {
                $label = $Word.CaptionLabels.Item($name)
                $registeredLabels.Add($label)
            }
            catch {
                $missing.Add($name)
            }
        }
        if ($missing.Count -gt 0) {
            $installer = Join-Path $PSScriptRoot 'install_word_caption_labels.ps1'
            throw (
                "Word 缺少宜宾论文自定义题注标签：$($missing -join '、')。" +
                "请先运行 '$installer'，重启 Word 后重新构建。"
            )
        }

        $updated = 0
        foreach ($field in @($Document.Fields)) {
            $paragraph = $null
            try {
                $code = [string]$field.Code.Text
                if ($code -notmatch '(?i)\bSEQ\s+(?:"[^"]+"|\S+)') {
                    continue
                }
                $paragraph = $field.Result.Paragraphs.Item(1)
                $text = Get-ParagraphText -Paragraph $paragraph
                $kind = $null
                if ($text.StartsWith('图')) {
                    $kind = 'Figure'
                }
                elseif ($text.StartsWith('表')) {
                    $kind = 'Table'
                }
                elseif ($paragraph.Range.OMaths.Count -gt 0) {
                    $kind = 'Equation'
                }
                if ($null -eq $kind) {
                    continue
                }
                $replacement = 'SEQ ' + $desired[$kind]
                $newCode = [regex]::Replace(
                    $code,
                    '(?i)\bSEQ\s+(?:"[^"]+"|\S+)',
                    $replacement,
                    1
                )
                if ($newCode -cne $code) {
                    $field.Code.Text = $newCode
                    $updated++
                }
            }
            finally {
                if ($null -ne $paragraph) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($paragraph)
                }
                if ($null -ne $field) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($field)
                }
            }
        }
        Write-Host (
            "Word captions: normalized $updated SEQ field(s) to registered labels " +
            "Figure='图', Table='表', Equation='公式'."
        )
    }
    finally {
        foreach ($label in $registeredLabels) {
            if ($null -ne $label) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($label)
            }
        }
    }
}

function Repair-OrphanedTableHeaders {
    param([Parameter(Mandatory = $true)]$Document)

    # wdActiveEndPageNumber=3, wdCollapseStart=1, wdCollapseEnd=0.
    $Document.Repaginate()
    $moved = 0
    foreach ($table in @($Document.Tables)) {
        if ($table.Rows.Count -lt 2) {
            continue
        }

        $headerRange = $null
        $dataRange = $null
        $beforeRange = $null
        $caption = $null
        try {
            $headerRange = $table.Rows.Item(1).Range.Duplicate
            $headerRange.Collapse(0)
            $headerPage = [int]$headerRange.Information(3)

            $dataRange = $table.Rows.Item(2).Range.Duplicate
            $dataRange.Collapse(1)
            $dataPage = [int]$dataRange.Information(3)
            if ($headerPage -eq $dataPage) {
                continue
            }

            $beforeEnd = [math]::Max(0, [int]$table.Range.Start - 1)
            $beforeRange = $Document.Range(0, $beforeEnd)
            if ($beforeRange.Paragraphs.Count -lt 1) {
                continue
            }
            $caption = $beforeRange.Paragraphs.Item($beforeRange.Paragraphs.Count)
            $captionStyle = Get-StyleName -Paragraph $caption
            if ($captionStyle -notmatch '(?i)Caption|题注|宜宾论文-表题') {
                Write-Warning "Table $($table.Index) has an orphaned header but no preceding Caption paragraph."
                continue
            }
            if ($caption.Format.PageBreakBefore -eq 0) {
                $caption.Format.PageBreakBefore = -1
                $moved++
            }
        }
        finally {
            foreach ($comObject in @($caption, $beforeRange, $dataRange, $headerRange)) {
                if ($null -ne $comObject) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
                }
            }
        }
    }

    if ($moved -gt 0) {
        $Document.Repaginate()
    }
    Write-Host "Word table pagination: moved $moved table(s) to avoid an orphaned repeating header."
}

function Get-TableCaptionInfo {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)]$Table
    )

    $beforeRange = $null
    $paragraph = $null
    try {
        $beforeEnd = [math]::Max(0, [int]$Table.Range.Start - 1)
        $beforeRange = $Document.Range(0, $beforeEnd)
        for ($index = $beforeRange.Paragraphs.Count; $index -ge 1; $index--) {
            $paragraph = $beforeRange.Paragraphs.Item($index)
            $text = Get-ParagraphText -Paragraph $paragraph
            if ([string]::IsNullOrWhiteSpace($text)) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($paragraph)
                $paragraph = $null
                continue
            }

            $styleName = Get-StyleName -Paragraph $paragraph
            $isCaption = $styleName -match '(?i)Caption|题注|宜宾论文-表题'
            $isContinuation = $styleName -in @('Yibin Table Continuation', '宜宾论文-续表题')
            if (-not ($isCaption -or $isContinuation)) {
                return $null
            }

            $number = ''
            if ($text -match '^(?:表|续表)\s*([0-9]+(?:\.[0-9]+)?)') {
                $number = $Matches[1]
            }
            if ([string]::IsNullOrWhiteSpace($number)) {
                return $null
            }

            $bookmarkName = ''
            foreach ($bookmark in @($paragraph.Range.Bookmarks)) {
                try {
                    # The generated number-only target is hidden, just like
                    # Word's native _Ref bookmarks.  Continuation captions use
                    # it so "续表" is followed by only 2.1, not "表2.1".
                    if ([string]$bookmark.Name -match '^_RefNum[0-9]+$') {
                        $bookmarkName = [string]$bookmark.Name
                        break
                    }
                }
                finally {
                    if ($null -ne $bookmark) {
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($bookmark)
                    }
                }
            }
            if ([string]::IsNullOrWhiteSpace($bookmarkName)) {
                foreach ($field in @($paragraph.Range.Fields)) {
                    try {
                        if ([string]$field.Code.Text -match '(?i)\bREF\s+([A-Za-z0-9_]+)') {
                            $bookmarkName = $Matches[1]
                            break
                        }
                    }
                    finally {
                        if ($null -ne $field) {
                            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($field)
                        }
                    }
                }
            }
            return [pscustomobject]@{
                Number = $number
                Bookmark = $bookmarkName
                Style = $styleName
                IsContinuation = $isContinuation
            }
        }
        return $null
    }
    finally {
        foreach ($comObject in @($paragraph, $beforeRange)) {
            if ($null -ne $comObject) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
            }
        }
    }
}

function Stabilize-TableRowsForContinuation {
    param([Parameter(Mandatory = $true)]$Document)

    # A continuation label can only be inserted between rows.  Keep a row
    # atomic when Word reports that it currently crosses a page and its line
    # count shows that it fits on a fresh A4 text page.  This is narrower than
    # the old blanket cantSplit policy and therefore avoids pinning every long
    # row to the next page.
    $changedTotal = 0
    for ($pass = 1; $pass -le 5; $pass++) {
        $Document.Repaginate()
        $changed = 0
        for ($tableIndex = 1; $tableIndex -le $Document.Tables.Count; $tableIndex++) {
            $table = $Document.Tables.Item($tableIndex)
            try {
                if ($table.Rows.Count -lt 2 -or $null -eq (Get-TableCaptionInfo -Document $Document -Table $table)) {
                    continue
                }
                for ($rowIndex = 2; $rowIndex -le $table.Rows.Count; $rowIndex++) {
                    $row = $table.Rows.Item($rowIndex)
                    $startRange = $null
                    $endRange = $null
                    try {
                        $startRange = $row.Range.Duplicate
                        $endRange = $row.Range.Duplicate
                        $startRange.Collapse(1)
                        $endRange.Collapse(0)
                        $startPage = [int]$startRange.Information(3)
                        $endPage = [int]$endRange.Information(3)
                        if ($startPage -eq $endPage) {
                            continue
                        }
                        $lineCount = [int]$row.Range.ComputeStatistics(1)
                        if ($lineCount -le 40 -and $row.AllowBreakAcrossPages -ne 0) {
                            $row.AllowBreakAcrossPages = 0
                            $changed++
                            $changedTotal++
                        }
                        elseif ($lineCount -gt 40) {
                            Write-Warning "Table $tableIndex row $rowIndex spans pages and is too tall for safe continuation splitting ($lineCount lines)."
                        }
                    }
                    finally {
                        foreach ($comObject in @($endRange, $startRange, $row)) {
                            if ($null -ne $comObject) {
                                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
                            }
                        }
                    }
                }
            }
            finally {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($table) } catch {}
            }
        }
        if ($changed -eq 0) {
            break
        }
    }
    if ($changedTotal -gt 0) {
        $Document.Repaginate()
    }
    Write-Host "Word table pagination: kept $changedTotal page-crossing row(s) together for continuation labels."
}

function Copy-TableHeaderRow {
    param(
        [Parameter(Mandatory = $true)]$SourceRow,
        [Parameter(Mandatory = $true)]$TargetTable
    )

    $targetRow = $TargetTable.Rows.Add($TargetTable.Rows.Item(1))
    try {
        if ($targetRow.Cells.Count -ne $SourceRow.Cells.Count) {
            throw 'Cannot copy a continuation header across incompatible table grids.'
        }
        for ($cellIndex = 1; $cellIndex -le $SourceRow.Cells.Count; $cellIndex++) {
            $sourceCell = $SourceRow.Cells.Item($cellIndex)
            $targetCell = $targetRow.Cells.Item($cellIndex)
            $sourceRange = $null
            $targetRange = $null
            try {
                $sourceRange = $sourceCell.Range.Duplicate
                $targetRange = $targetCell.Range.Duplicate
                [void]$sourceRange.MoveEnd(1, -1)
                [void]$targetRange.MoveEnd(1, -1)
                $targetRange.FormattedText = $sourceRange.FormattedText
                if ([int]$sourceCell.VerticalAlignment -ne 9999999) {
                    $targetCell.VerticalAlignment = $sourceCell.VerticalAlignment
                }
                if ([int]$sourceCell.Shading.BackgroundPatternColor -ne 9999999) {
                    $targetCell.Shading.BackgroundPatternColor = $sourceCell.Shading.BackgroundPatternColor
                }
                foreach ($borderType in @(-1, -2, -3, -4)) {
                    $sourceBorder = $sourceCell.Borders.Item($borderType)
                    $targetBorder = $targetCell.Borders.Item($borderType)
                    try {
                        if ([int]$sourceBorder.LineStyle -ne 9999999) {
                            $targetBorder.LineStyle = $sourceBorder.LineStyle
                        }
                        # Word reports 9999999 for an undefined width/color.
                        # Reassigning that sentinel raises "value out of range".
                        if ([int]$sourceBorder.LineStyle -notin @(0, 9999999) -and
                            [int]$sourceBorder.LineWidth -gt 0 -and
                            [int]$sourceBorder.LineWidth -ne 9999999) {
                            try { $targetBorder.LineWidth = $sourceBorder.LineWidth } catch {}
                        }
                        if ([int]$sourceBorder.Color -ne 9999999) {
                            try { $targetBorder.Color = $sourceBorder.Color } catch {}
                        }
                    }
                    finally {
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($targetBorder)
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sourceBorder)
                    }
                }
            }
            finally {
                foreach ($comObject in @($targetRange, $sourceRange, $targetCell, $sourceCell)) {
                    if ($null -ne $comObject) {
                        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
                    }
                }
            }
        }
        $targetRow.HeadingFormat = -1
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($targetRow)
    }
}

function Set-ContinuationParagraph {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)]$PreviousTable,
        [Parameter(Mandatory = $true)]$ContinuationTable,
        [Parameter(Mandatory = $true)]$CaptionInfo
    )

    $gap = $null
    $paragraph = $null
    $textRange = $null
    $fieldRange = $null
    $field = $null
    try {
        $gap = $Document.Range($PreviousTable.Range.End, $ContinuationTable.Range.Start)
        $paragraph = $gap.Paragraphs.Item(1)
        $textRange = $paragraph.Range.Duplicate
        if ($textRange.Characters.Count -gt 0) {
            [void]$textRange.MoveEnd(1, -1)
        }
        $textRange.Text = '续表 '

        $fieldRange = $paragraph.Range.Duplicate
        if ($fieldRange.Characters.Count -gt 0) {
            [void]$fieldRange.MoveEnd(1, -1)
        }
        $fieldRange.Collapse(0)
        if (-not [string]::IsNullOrWhiteSpace([string]$CaptionInfo.Bookmark)) {
            $field = $Document.Fields.Add(
                $fieldRange,
                -1,
                " REF $($CaptionInfo.Bookmark) \h ",
                $true
            )
            [void]$field.Update()
        }
        else {
            $fieldRange.Text = [string]$CaptionInfo.Number
        }

        try {
            $paragraph.Range.Style = $Document.Styles.Item('宜宾论文-续表题')
        }
        catch {}
    }
    finally {
        foreach ($comObject in @($field, $fieldRange, $textRange, $paragraph, $gap)) {
            if ($null -ne $comObject) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
            }
        }
    }
}

function Split-TableSegmentsForContinuation {
    param([Parameter(Mandatory = $true)]$Document)

    $splitCount = 0
    $initialTableCount = $Document.Tables.Count
    for ($tableIndex = $initialTableCount; $tableIndex -ge 1; $tableIndex--) {
        $table = $Document.Tables.Item($tableIndex)
        $sourceHeader = $null
        try {
            if ($table.Rows.Count -lt 3) {
                continue
            }
            $captionInfo = Get-TableCaptionInfo -Document $Document -Table $table
            if ($null -eq $captionInfo) {
                continue
            }

            # Explicit continuation splitting inserts a copy of the first row
            # into every later segment.  That is only safe when every row uses
            # the same cell grid as the header.  Questionnaire matrices often
            # merge header cells across a finer body grid (for example 3 header
            # cells over 19 response columns).  Word can repeat that native
            # header itself, but a copied row cannot be inserted into the
            # incompatible grid without corrupting the table.  Keep the native
            # repeating header and skip explicit ``续表`` segmentation there.
            $sourceHeader = $table.Rows.Item(1)
            $headerCellCount = $sourceHeader.Cells.Count
            $compatibleGrid = $true
            for ($rowIndex = 2; $rowIndex -le $table.Rows.Count; $rowIndex++) {
                $gridRow = $table.Rows.Item($rowIndex)
                try {
                    if ($gridRow.Cells.Count -ne $headerCellCount) {
                        $compatibleGrid = $false
                        break
                    }
                }
                finally {
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($gridRow) } catch {}
                }
            }
            if (-not $compatibleGrid) {
                Write-Warning (
                    "Table $tableIndex uses merged/incompatible row grids; " +
                    "kept Word's native repeating header and skipped explicit continuation splitting."
                )
                continue
            }

            $boundaries = New-Object System.Collections.Generic.List[int]
            $previousPage = 0
            for ($rowIndex = 2; $rowIndex -le $table.Rows.Count; $rowIndex++) {
                $rowRange = $table.Rows.Item($rowIndex).Range.Duplicate
                try {
                    $rowRange.Collapse(1)
                    $page = [int]$rowRange.Information(3)
                }
                finally {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($rowRange)
                }
                if ($previousPage -ne 0 -and $page -ne $previousPage) {
                    $boundaries.Add($rowIndex)
                }
                $previousPage = $page
            }
            if ($boundaries.Count -eq 0) {
                continue
            }

            $continuationTables = New-Object System.Collections.Generic.List[object]
            for ($boundaryIndex = $boundaries.Count - 1; $boundaryIndex -ge 0; $boundaryIndex--) {
                $row = $table.Rows.Item($boundaries[$boundaryIndex])
                try {
                    $continuationTables.Add($table.Split($row))
                }
                finally {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($row)
                }
            }
            $orderedContinuations = @($continuationTables | Sort-Object { $_.Range.Start })
            $previousTable = $table
            foreach ($continuationTable in $orderedContinuations) {
                Copy-TableHeaderRow -SourceRow $sourceHeader -TargetTable $continuationTable
                Set-ContinuationParagraph `
                    -Document $Document `
                    -PreviousTable $previousTable `
                    -ContinuationTable $continuationTable `
                    -CaptionInfo $captionInfo
                if ($previousTable -ne $table) {
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($previousTable) } catch {}
                }
                $previousTable = $continuationTable
                $splitCount++
            }
            if ($previousTable -ne $table) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($previousTable) } catch {}
            }
        }
        finally {
            foreach ($comObject in @($sourceHeader, $table)) {
                if ($null -ne $comObject) {
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
                }
            }
        }
    }
    if ($splitCount -gt 0) {
        $Document.Repaginate()
    }
    return $splitCount
}

function Add-TableContinuations {
    param([Parameter(Mandatory = $true)]$Document)

    $total = 0
    for ($pass = 1; $pass -le 4; $pass++) {
        Stabilize-TableRowsForContinuation -Document $Document
        $added = Split-TableSegmentsForContinuation -Document $Document
        $total += $added
        if ($added -eq 0) {
            break
        }
    }
    Write-Host "Word table continuation: inserted $total continuation segment(s)."
}

function Compact-TableContinuations {
    param([Parameter(Mandatory = $true)]$Document)

    # A split boundary is discovered before its continuation paragraph forces
    # the following segment to the top of a page.  That move can create enough
    # room to reunite two adjacent segments.  Test the layout without the later
    # page break and merge only when both complete segments fit on one page.
    $mergedTotal = 0
    for ($pass = 1; $pass -le 12; $pass++) {
        $Document.Repaginate()
        $merged = $false
        for ($tableIndex = 1; $tableIndex -lt $Document.Tables.Count; $tableIndex++) {
            $leftTable = $null
            $rightTable = $null
            $captionRange = $null
            $captionParagraph = $null
            $leftStart = $null
            $rightEnd = $null
            try {
                $leftTable = $Document.Tables.Item($tableIndex)
                $rightTable = $Document.Tables.Item($tableIndex + 1)
                $leftInfo = Get-TableCaptionInfo -Document $Document -Table $leftTable
                $rightInfo = Get-TableCaptionInfo -Document $Document -Table $rightTable
                if ($null -eq $leftInfo -or $null -eq $rightInfo -or
                    -not [bool]$rightInfo.IsContinuation -or
                    [string]$leftInfo.Bookmark -ne [string]$rightInfo.Bookmark) {
                    continue
                }

                $captionRange = $Document.Range($leftTable.Range.End, $rightTable.Range.Start)
                $captionParagraph = $captionRange.Paragraphs.Item(1)
                if ((Get-StyleName -Paragraph $captionParagraph) -ne '宜宾论文-续表题') {
                    continue
                }
                $captionParagraph.Format.PageBreakBefore = 0
                $Document.Repaginate()

                $leftStart = $leftTable.Rows.Item(1).Range.Duplicate
                $leftStart.Collapse(1)
                $rightEnd = $rightTable.Rows.Item($rightTable.Rows.Count).Range.Duplicate
                $rightEnd.Collapse(0)
                $startPage = [int]$leftStart.Information(3)
                $endPage = [int]$rightEnd.Information(3)
                if ($startPage -ne $endPage) {
                    $captionParagraph.Format.PageBreakBefore = -1
                    continue
                }

                # The left segment already owns the repeated header.  Removing
                # the later header and separator paragraph makes Word join the
                # compatible table grids into one semantic table.
                $rightTable.Rows.Item(1).Delete()
                $separator = $Document.Range($leftTable.Range.End, $rightTable.Range.Start)
                try {
                    $separator.Delete()
                }
                finally {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($separator)
                }
                $Document.Repaginate()
                $mergedTotal++
                $merged = $true
                break
            }
            finally {
                foreach ($comObject in @(
                    $rightEnd,
                    $leftStart,
                    $captionParagraph,
                    $captionRange,
                    $rightTable,
                    $leftTable
                )) {
                    if ($null -ne $comObject) {
                        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
                    }
                }
            }
        }
        if (-not $merged) {
            break
        }
    }
    Write-Host "Word table continuation: compacted $mergedTotal adjacent segment(s)."
}

# First pass: update fields in Word and apply visible formatting.  The official
# document uses a 10.5 pt Normal style as the character-indent measuring base,
# while all visible TOC text is directly formatted as 12 pt Song.  Keeping that
# distinction is what makes 2/4 character indents compute to 21/42 pt.
$coreProperties = Get-DocxCoreProperties -Path $documentPath
$documentType = $null
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($documentPath, $false, $false)
    $documentType = Resolve-YibinDocumentType `
        -Document $document `
        -CoreProperties $coreProperties
    Write-Host "Word document profile: $documentType"

    try { $document.Bookmarks.ShowHidden = $true } catch {}
    Normalize-NativeCaptionSequences -Document $document -Word $word

    Update-WordDocumentFields -Document $document

    # -1 is wdStyleNormal and is locale-independent.
    $normalStyle = $document.Styles.Item(-1)
    $normalStyle.Font.NameFarEast = '宋体'
    $normalStyle.Font.Name = 'Times New Roman'
    $normalStyle.Font.Size = 10.5

    $thesisOnlyStyleNames = @(
        'ChineseAbstract',
        'EnglishAbstract',
        'Keywords',
        'FrontTitle',
        '宜宾论文-中文页标题',
        '宜宾论文-英文摘要标题',
        '宜宾论文-目录标题',
        '宜宾论文-附录标题',
        '宜宾论文-中文摘要正文',
        '宜宾论文-英文摘要正文',
        '宜宾论文-关键词',
        'DeclarationBody',
        'DeclarationSignature',
        'DeclarationDate',
        'DeclarationRegulationLead',
        'DeclarationRegulationClause',
        'DeclarationRegulationItem'
    )

    foreach ($contract in @(
        @('First Paragraph', '宋体', 'Times New Roman', 12, $false),
        @('Body Text', '宋体', 'Times New Roman', 12, $false),
        @('ChineseAbstract', '宋体', 'Times New Roman', 12, $false),
        @('EnglishAbstract', 'Times New Roman', 'Times New Roman', 12, $false),
        @('Keywords', '宋体', 'Times New Roman', 12, $false),
        @('FrontTitle', '黑体', 'Times New Roman', 16, $true),
        @('Heading 1', '黑体', 'Times New Roman', 16, $true),
        @('Heading 2', '楷体', 'Times New Roman', 15, $true),
        @('Heading 3', '宋体', 'Times New Roman', 14, $true),
        @('Heading 4', '宋体', 'Times New Roman', 14, $true),
        @('Yibin Heading 1', '黑体', 'Times New Roman', 16, $true),
        @('Yibin Heading 2', '楷体', 'Times New Roman', 15, $true),
        @('Yibin Heading 3', '宋体', 'Times New Roman', 14, $true),
        @('Yibin Heading 4', '宋体', 'Times New Roman', 14, $true),
        @('Unnumbered Heading 1', '黑体', 'Times New Roman', 16, $true),
        @('宜宾论文-正文', '宋体', 'Times New Roman', 12, $false),
        @('宜宾论文-首段', '宋体', 'Times New Roman', 12, $false),
        @('宜宾论文-一级标题', '黑体', 'Times New Roman', 16, $true),
        @('宜宾论文-二级标题', '楷体', 'Times New Roman', 15, $true),
        @('宜宾论文-三级标题', '宋体', 'Times New Roman', 14, $true),
        @('宜宾论文-四级标题', '宋体', 'Times New Roman', 14, $true),
        @('宜宾论文-无编号标题', '黑体', 'Times New Roman', 16, $true),
        @('宜宾论文-中文页标题', '黑体', 'Times New Roman', 16, $true),
        @('宜宾论文-英文摘要标题', 'Times New Roman', 'Times New Roman', 16, $true),
        @('宜宾论文-目录标题', '黑体', 'Times New Roman', 16, $true),
        @('宜宾论文-附录标题', '黑体', 'Times New Roman', 16, $true),
        @('宜宾论文-中文摘要正文', '宋体', 'Times New Roman', 12, $false),
        @('宜宾论文-英文摘要正文', 'Times New Roman', 'Times New Roman', 12, $false),
        @('宜宾论文-关键词', '宋体', 'Times New Roman', 12, $false),
        @('宜宾论文-插图', '宋体', 'Times New Roman', 12, $false),
        @('宜宾论文-公式', '宋体', 'Times New Roman', 12, $false),
        @('DeclarationBody', '宋体', 'Times New Roman', 12, $false),
        @('DeclarationSignature', '宋体', 'Times New Roman', 12, $false),
        @('DeclarationDate', '宋体', 'Times New Roman', 12, $false),
        @('DeclarationRegulationLead', '宋体', '宋体', 14, $true),
        @('DeclarationRegulationClause', '宋体', '宋体', 14, $true),
        @('DeclarationRegulationItem', '宋体', '宋体', 14, $true),
        @('caption', '宋体', 'Times New Roman', 10.5, $false),
        @('Caption', '宋体', 'Times New Roman', 10.5, $false),
        @('Table Caption', '宋体', 'Times New Roman', 10.5, $false),
        @('Image Caption', '宋体', 'Times New Roman', 10.5, $false),
        @('Figure Caption', '宋体', 'Times New Roman', 10.5, $false),
        @('Yibin Table Continuation', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-图题', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-表题', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-续表题', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-表格正文', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-表格居中', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-表格右对齐', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-表头', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-表头左对齐', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-表头右对齐', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-参考文献', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-注释', '宋体', 'Times New Roman', 10.5, $false),
        @('宜宾论文-文献上标', '宋体', 'Times New Roman', 9, $false),
        @('宜宾论文-三线表', '宋体', 'Times New Roman', 10.5, $false),
        @('Bibliography', '宋体', 'Times New Roman', 10.5, $false),
        @('Notes', '宋体', 'Times New Roman', 10.5, $false),
        @('Header', '宋体', 'Times New Roman', 9, $false),
        @('Footer', '宋体', 'Times New Roman', 9, $false)
    )) {
        if ($documentType -ne 'thesis' -and $contract[0] -in $thesisOnlyStyleNames) {
            continue
        }
        Set-StyleFont `
            -Document $document `
            -Name $contract[0] `
            -EastAsia $contract[1] `
            -Latin $contract[2] `
            -Size ([double]$contract[3]) `
            -Bold ([bool]$contract[4])
    }

    $profileStyleContracts = switch ($documentType) {
        'proposal' {
            @(
                @('宜宾开题-标题', '楷体', 'Times New Roman', 18, $true),
                @('宜宾开题-副标题', '宋体', 'Times New Roman', 12, $true),
                @('宜宾开题-栏目', '黑体', 'Times New Roman', 12, $true),
                @('宜宾开题-正文', '宋体', 'Times New Roman', 12, $false),
                @('宜宾开题-提示', '宋体', 'Times New Roman', 12, $false),
                @('宜宾开题-无编号标题', '楷体', 'Times New Roman', 15, $true),
                @('宜宾开题-签名', '宋体', 'Times New Roman', 12, $false)
            )
        }
        'literature-review' {
            @(
                @('宜宾综述-文档标题', '黑体', 'Times New Roman', 28, $true),
                @('宜宾综述-论文题目', '黑体', 'Times New Roman', 18, $true),
                @('宜宾综述-信息标签', '黑体', 'Times New Roman', 18, $true),
                @('宜宾综述-信息值', '宋体', 'Times New Roman', 16, $true)
            )
        }
        default { @() }
    }
    foreach ($contract in $profileStyleContracts) {
        Set-StyleFont `
            -Document $document `
            -Name $contract[0] `
            -EastAsia $contract[1] `
            -Latin $contract[2] `
            -Size ([double]$contract[3]) `
            -Bold ([bool]$contract[4])
    }

    if ($documentType -eq 'thesis') {
        foreach ($level in 1..3) {
            try {
                $tocStyle = $document.Styles.Item("TOC $level")
                $leftIndent = @(0, 21, 42)[$level - 1]
                $characterIndent = @(0, 2, 4)[$level - 1]
                $tocStyle.Font.NameFarEast = '宋体'
                $tocStyle.Font.Name = '宋体'
                $tocStyle.Font.Size = 12
                $tocStyle.Font.Bold = 0
                $tocStyle.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                $tocStyle.ParagraphFormat.FirstLineIndent = 0
                $tocStyle.ParagraphFormat.CharacterUnitLeftIndent = $characterIndent
                $tocStyle.ParagraphFormat.LeftIndent = $leftIndent
                $tocStyle.ParagraphFormat.Alignment = 3
                $tocStyle.ParagraphFormat.SpaceBefore = 0
                $tocStyle.ParagraphFormat.SpaceAfter = 0
                $tocStyle.ParagraphFormat.LineSpacingRule = if ($level -lt 3) { 1 } else { 0 }
                $tocStyle.ParagraphFormat.LineSpacing = if ($level -lt 3) { 18 } else { 12 }
                $tocStyle.ParagraphFormat.TabStops.ClearAll()
                [void]$tocStyle.ParagraphFormat.TabStops.Add(438.85, 2, 1)
            }
            catch {
                # A short thesis may not materialize every TOC level.
            }
        }

        # -36 is wdStyleTableOfFigures and is locale-independent. Word applies
        # it to entries generated by both the figure- and table-directory fields.
        $captionDirectoryStyle = $document.Styles.Item(-36)
        $captionDirectoryStyle.Font.NameFarEast = '宋体'
        $captionDirectoryStyle.Font.Name = '宋体'
        $captionDirectoryStyle.Font.Size = 12
        $captionDirectoryStyle.Font.Bold = 0
        $captionDirectoryStyle.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        $captionDirectoryStyle.ParagraphFormat.FirstLineIndent = 0
        $captionDirectoryStyle.ParagraphFormat.CharacterUnitLeftIndent = 0
        $captionDirectoryStyle.ParagraphFormat.LeftIndent = 0
        $captionDirectoryStyle.ParagraphFormat.Alignment = 3
        $captionDirectoryStyle.ParagraphFormat.SpaceBefore = 0
        $captionDirectoryStyle.ParagraphFormat.SpaceAfter = 0
        $captionDirectoryStyle.ParagraphFormat.LineSpacingRule = 1
        $captionDirectoryStyle.ParagraphFormat.LineSpacing = 18
        $captionDirectoryStyle.ParagraphFormat.TabStops.ClearAll()
        [void]$captionDirectoryStyle.ParagraphFormat.TabStops.Add(438.85, 2, 1)

        try {
            $tocHeadingStyle = $document.Styles.Item('宜宾论文-目录标题')
            $tocHeadingStyle.ParagraphFormat.Alignment = 1
            $tocHeadingStyle.ParagraphFormat.SpaceBefore = 0
            $tocHeadingStyle.ParagraphFormat.SpaceAfter = 0
            $tocHeadingStyle.ParagraphFormat.LineSpacingRule = 0
        }
        catch {
            # The generated thesis is still checked paragraph-by-paragraph below.
        }
    }

    Stabilize-TableRowsForContinuation -Document $document
    Repair-OrphanedTableHeaders -Document $document
    Add-TableContinuations -Document $document
    Set-CitationFieldStyles -Document $document
    $document.Repaginate()
    # Continuation captions and row stabilization can change physical page
    # positions.  Refresh the native TOC/figure/table fields only after those
    # layout changes have settled so their cached page numbers are current.
    Update-WordDocumentFields -Document $document
    $document.Repaginate()
    $document.Save()
}
finally {
    Close-WordSession -Document $document -Word $word
}

# COM cannot persist w:left and w:leftChars simultaneously on every Word
# version. Patch the two OOXML parts after the field update, then reopen the
# result read-only for computed-format and PDF verification.
if ($documentType -eq 'thesis') {
    Set-OfficialTocOoxml -Path $documentPath
}

if ($RefreshOnly) {
    Write-Host "Word fields refreshed and document repaginated ($documentType): $documentPath"
    return
}

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($documentPath, $false, $true)
    $document.Repaginate()

    $normalStyle = $document.Styles.Item(-1)
    Assert-Font 'Normal East Asian font' $normalStyle.Font.NameFarEast @('宋体', 'SimSun')
    Assert-Near 'Normal measuring size' $normalStyle.Font.Size 10.5

    if ($documentType -eq 'proposal') {
        Assert-ProposalProfile -Document $document
    }
    elseif ($documentType -eq 'literature-review') {
        Assert-LiteratureReviewProfile -Document $document
    }

    if ($documentType -ne 'thesis') {
        if (-not [string]::IsNullOrWhiteSpace($PdfOutput)) {
            # 17 = wdExportFormatPDF
            $document.ExportAsFixedFormat($pdfPath, 17)
            Write-Host "PDF exported: $pdfPath"
        }
        Write-Host "Word fields refreshed and profile audited ($documentType): $documentPath"
        return
    }

    $coverLogo = @()
    $coverThesisType = @()
    $coverTitle = @()
    $coverVersions = @()
    $coverFields = @()
    $coverValueLines = @()
    $coverDates = @()
    $declarationTitles = @()
    $declarationBodies = @()
    $declarationSignatures = @()
    $regulationLeads = @()
    $regulationClauses = @()
    $regulationItems = @()
    $tocHeadings = @()
    $tocParagraphs = @{}
    $captionDirectoryParagraphs = @()
    foreach ($level in 1..3) { $tocParagraphs[$level] = @() }

    foreach ($paragraph in @($document.Paragraphs)) {
        $styleName = Get-StyleName -Paragraph $paragraph
        switch -Regex ($styleName) {
            '^CoverLogo$' { $coverLogo += $paragraph; continue }
            '^CoverThesisType$' { $coverThesisType += $paragraph; continue }
            '^CoverTitle(?:WithVersion)?$' { $coverTitle += $paragraph; continue }
            '^CoverVersion$' { $coverVersions += $paragraph; continue }
            '^CoverField$' { $coverFields += $paragraph; continue }
            '^CoverValueLine$' { $coverValueLines += $paragraph; continue }
            '^CoverDate$' { $coverDates += $paragraph; continue }
            '^DeclarationTitle$' { $declarationTitles += $paragraph; continue }
            '^DeclarationBody$' { $declarationBodies += $paragraph; continue }
            '^DeclarationSignature$' { $declarationSignatures += $paragraph; continue }
            '^DeclarationRegulationLead$' { $regulationLeads += $paragraph; continue }
            '^DeclarationRegulationClause$' { $regulationClauses += $paragraph; continue }
            '^DeclarationRegulationItem$' { $regulationItems += $paragraph; continue }
            '^(?:Yibin TOC Heading|宜宾论文-目录标题)$' { $tocHeadings += $paragraph; continue }
            '(?i)^TOC\s+([1-3])$' {
                $tocParagraphs[[int]$Matches[1]] += $paragraph
                continue
            }
            '(?i)^(?:Table of Figures|图表目录)$' {
                $captionDirectoryParagraphs += $paragraph
                continue
            }
        }
    }

    if ($coverLogo.Count -ne 1) { throw "CoverLogo paragraph count expected 1, got $($coverLogo.Count)" }
    if ($coverLogo[0].Range.InlineShapes.Count -ne 1) {
        throw 'CoverLogo must contain exactly one inline official logo.'
    }
    $logo = $coverLogo[0].Range.InlineShapes.Item(1)
    Assert-Near 'Cover logo width' $logo.Width 372.75 0.8
    Assert-Near 'Cover logo height' $logo.Height 103.5 0.8

    if ($coverThesisType.Count -ne 1) {
        throw "CoverThesisType paragraph count expected 1, got $($coverThesisType.Count)"
    }
    Assert-Text 'Cover thesis type text' `
        (Get-ParagraphText -Paragraph $coverThesisType[0]) `
        '本科生毕业论文（设计）'
    Assert-Font 'Cover thesis type font' `
        $coverThesisType[0].Range.Font.NameFarEast @('黑体', 'SimHei')
    Assert-Near 'Cover thesis type size' $coverThesisType[0].Range.Font.Size 28
    if ($coverThesisType[0].Range.Font.Bold -eq 0) { throw 'Cover thesis type must be bold.' }
    if ($coverThesisType[0].Format.Alignment -ne 1) { throw 'Cover thesis type must be centered.' }

    if ($coverTitle.Count -ne 1) { throw "CoverTitle paragraph count expected 1, got $($coverTitle.Count)" }
    Assert-Font 'Cover title font' $coverTitle[0].Range.Font.NameFarEast @('黑体', 'SimHei')
    Assert-Near 'Cover title size' $coverTitle[0].Range.Font.Size 18
    if ($coverTitle[0].Range.Font.Bold -eq 0) { throw 'Cover title must be bold.' }
    if ($coverTitle[0].Format.Alignment -ne 1) { throw 'Cover title must be centered.' }
    if ($coverTitle[0].Range.Font.Underline -eq 0) { throw 'Cover title must use a single underline.' }

    if ($coverVersions.Count -gt 1) {
        throw "CoverVersion paragraph count expected at most 1, got $($coverVersions.Count)"
    }
    $coverTitleStyle = Get-StyleName -Paragraph $coverTitle[0]
    if ($coverVersions.Count -eq 1) {
        if ($coverTitleStyle -ne 'CoverTitleWithVersion') {
            throw "Cover title with a version must use CoverTitleWithVersion, got '$coverTitleStyle'"
        }
        Assert-Font 'Cover version font' $coverVersions[0].Range.Font.NameFarEast @('黑体', 'SimHei')
        Assert-Near 'Cover version size' $coverVersions[0].Range.Font.Size 14
        if ($coverVersions[0].Range.Font.Bold -ne 0) { throw 'Cover version must not be bold.' }
        if ($coverVersions[0].Format.Alignment -ne 1) { throw 'Cover version must be centered.' }
    }
    elseif ($coverTitleStyle -ne 'CoverTitle') {
        throw "Cover title without a version must use CoverTitle, got '$coverTitleStyle'"
    }

    if ($coverFields.Count -ne 6) { throw "CoverField paragraph count expected 6, got $($coverFields.Count)" }
    if ($coverDates.Count -gt 1) {
        throw "CoverDate paragraph count expected at most 1, got $($coverDates.Count)"
    }
    if ($coverDates.Count -eq 1) {
        Assert-Font 'Cover date font' $coverDates[0].Range.Font.NameFarEast @('宋体', 'SimSun')
        Assert-Near 'Cover date size' $coverDates[0].Range.Font.Size 15
        if ($coverDates[0].Range.Font.Bold -eq 0) { throw 'Cover date must be bold.' }
        if ($coverDates[0].Format.Alignment -ne 1) { throw 'Cover date must be centered.' }
    }
    $fieldTokens = @(
        @{ Tokens = @('学院（部）') },
        @{ Tokens = @('专业') },
        @{ Tokens = @('学生姓名') },
        @{ Tokens = @('学号') },
        @{ Tokens = @('指导教师（校内）') },
        @{ Tokens = @('指导教师（校外）') }
    )
    for ($index = 0; $index -lt $coverFields.Count; $index++) {
        $normalized = (Get-ParagraphText -Paragraph $coverFields[$index]) -replace '\s', ''
        foreach ($token in $fieldTokens[$index].Tokens) {
            if ($normalized -notlike "*$($token -replace '\s', '')*") {
                throw "Cover field $($index + 1) is missing '$token': $normalized"
            }
        }
    }

    $coverValueParagraphs = @()
    foreach ($paragraph in $coverValueLines) {
        $hasCoverValue = $false
        foreach ($character in @($paragraph.Range.Characters)) {
            $characterStyle = ''
            try { $characterStyle = [string]$character.Style.NameLocal } catch {}
            if ($characterStyle -eq 'CoverValue' -and $character.Text -ne [string][char]13) {
                $hasCoverValue = $true
                if ($character.Font.Underline -ne 0) {
                    throw 'CoverValue must not use character underlining; the CoverValueLine paragraph border is the only fill line.'
                }
                if ($character.Text -eq [string][char]0x00A0) {
                    throw 'CoverValue must not contain NBSP padding.'
                }
            }
        }
        if ($hasCoverValue) {
            if ($paragraph.Format.Alignment -ne 1) {
                throw 'CoverValue paragraph must be centered inside its fill slot.'
            }
            $coverValueParagraphs += $paragraph
        }
    }
    # Empty fill slots deliberately contain no padding characters, so Word COM
    # cannot enumerate them by character style.  Exact nine-slot geometry is
    # enforced by the deterministic OOXML audit in lib/audit_format.py.

    $originalityTitle = $declarationTitles | Where-Object {
        (Get-ParagraphText -Paragraph $_) -eq '原创性声明'
    } | Select-Object -First 1
    if ($null -eq $originalityTitle) { throw 'Originality declaration title was not found.' }
    Assert-Font 'Declaration title font' $originalityTitle.Range.Font.NameFarEast @('黑体', 'SimHei')
    Assert-Near 'Declaration title size' $originalityTitle.Range.Font.Size 16
    if ($originalityTitle.Range.Font.Bold -eq 0) { throw 'Declaration title must be bold.' }
    if ($originalityTitle.Format.Alignment -ne 1) { throw 'Declaration title must be centered.' }

    $mainDeclaration = $declarationBodies | Where-Object {
        (Get-ParagraphText -Paragraph $_).StartsWith('本人呈交的学位论文')
    } | Select-Object -First 1
    if ($null -eq $mainDeclaration) { throw 'Originality declaration body was not found.' }
    Assert-Font 'Declaration body font' $mainDeclaration.Range.Font.NameFarEast @('宋体', 'SimSun')
    Assert-Near 'Declaration body size' $mainDeclaration.Range.Font.Size 12
    Assert-Near 'Declaration body first-line indent' $mainDeclaration.Format.FirstLineIndent 24 1.0
    if ($mainDeclaration.Format.LineSpacingRule -ne 2) { throw 'Declaration body must use double line spacing.' }
    Assert-Near 'Declaration body line spacing' $mainDeclaration.Format.LineSpacing 24 1.0

    $officialAuthorSignature = $declarationSignatures | Where-Object {
        (Get-ParagraphText -Paragraph $_).StartsWith('学位论文作者：')
    } | Select-Object -First 1
    if ($null -eq $officialAuthorSignature) {
        throw 'The official signature line must start with 学位论文作者：.'
    }

    if ($regulationLeads.Count -ne 1 -or $regulationClauses.Count -ne 1 -or $regulationItems.Count -ne 1) {
        throw 'The originality declaration must contain exactly three regulation paragraphs.'
    }
    Assert-Text 'Regulation lead text' `
        (Get-ParagraphText -Paragraph $regulationLeads[0]) `
        '附：《普通高等学校学生管理规定》（中华人民共和国教育部令第41号）'
    Assert-Text 'Regulation clause text' `
        (Get-ParagraphText -Paragraph $regulationClauses[0]) `
        "第五十二条$([char]0x00A0)学生有下列情形之一，学校可以给予开除学籍处分："
    Assert-Text 'Regulation item text' `
        (Get-ParagraphText -Paragraph $regulationItems[0]) `
        '（五）学位论文、公开发表的研究成果存在抄袭、篡改、伪造等学术不端行为，情节严重的，或者代写论文、买卖论文的；'

    $leadLabel = $regulationLeads[0].Range.Duplicate
    $leadLabel.End = $leadLabel.Start + 2
    Assert-Font 'Regulation 附： font' $leadLabel.Font.NameFarEast @('宋体', 'SimSun')
    Assert-Near 'Regulation 附： size' $leadLabel.Font.Size 16
    if ($leadLabel.Font.Bold -eq 0) { throw 'Regulation 附： must be bold.' }
    $leadText = $regulationLeads[0].Range.Duplicate
    $leadText.Start = $leadText.Start + 2
    $leadText.End = $leadText.End - 1
    Assert-Font 'Regulation lead font' $leadText.Font.NameFarEast @('宋体', 'SimSun')
    Assert-Near 'Regulation lead size' $leadText.Font.Size 14
    if ($leadText.Font.Bold -eq 0) { throw 'Regulation lead must be bold.' }

    foreach ($pair in @(
        @('Regulation clause', $regulationClauses[0]),
        @('Regulation item', $regulationItems[0])
    )) {
        $regulationParagraph = $pair[1]
        Assert-Font "$($pair[0]) font" $regulationParagraph.Range.Font.NameFarEast @('宋体', 'SimSun')
        Assert-Near "$($pair[0]) size" $regulationParagraph.Range.Font.Size 14
        if ($regulationParagraph.Range.Font.Bold -eq 0) { throw "$($pair[0]) must be bold." }
        if ($regulationParagraph.Format.LineSpacingRule -ne 1) { throw "$($pair[0]) must use 1.5-line spacing." }
        Assert-Near "$($pair[0]) line spacing" $regulationParagraph.Format.LineSpacing 18 1.0
    }
    Assert-Near 'Regulation clause first-line indent' $regulationClauses[0].Format.FirstLineIndent 35 1.0
    Assert-Near 'Regulation item left indent' $regulationItems[0].Format.LeftIndent 7.15 1.0

    $expectedDirectoryHeadings = @('目录', '图目录', '表目录')
    if ($tocHeadings.Count -ne $expectedDirectoryHeadings.Count) {
        throw "Directory heading paragraph count expected $($expectedDirectoryHeadings.Count), got $($tocHeadings.Count)"
    }
    $actualDirectoryHeadings = @(
        $tocHeadings | ForEach-Object { Get-ParagraphText -Paragraph $_ }
    )
    foreach ($expectedHeading in $expectedDirectoryHeadings) {
        if ($actualDirectoryHeadings -notcontains $expectedHeading) {
            throw "Directory heading '$expectedHeading' was not found."
        }
    }
    foreach ($tocHeading in $tocHeadings) {
        Assert-Font 'Directory heading font' $tocHeading.Range.Font.NameFarEast @('黑体', 'SimHei')
        Assert-Near 'Directory heading size' $tocHeading.Range.Font.Size 16
        if ($tocHeading.Format.Alignment -ne 1) { throw 'Directory heading must be centered.' }
        if ($tocHeading.Format.LineSpacingRule -ne 0) { throw 'Directory heading must use single line spacing.' }
        Assert-Near 'Directory heading before spacing' $tocHeading.Format.SpaceBefore 0
        Assert-Near 'Directory heading after spacing' $tocHeading.Format.SpaceAfter 0
    }

    foreach ($level in 1..3) {
        if ($tocParagraphs[$level].Count -eq 0) { throw "TOC $level has no materialized paragraphs." }
        $expectedLeft = @(0, 21, 42)[$level - 1]
        $expectedCharacters = @(0, 2, 4)[$level - 1]
        $expectedLine = if ($level -lt 3) { 18 } else { 12 }
        $expectedRule = if ($level -lt 3) { 1 } else { 0 }
        foreach ($paragraph in $tocParagraphs[$level]) {
            $font = $paragraph.Range.Font
            $format = $paragraph.Format
            Assert-Font "TOC $level East Asian font" $font.NameFarEast @('宋体', 'SimSun')
            Assert-Font "TOC $level Latin font" $font.Name @('宋体', 'SimSun')
            Assert-Near "TOC $level size" $font.Size 12
            Assert-Near "TOC $level first-line indent" $format.FirstLineIndent 0
            Assert-Near "TOC $level character first-line indent" $format.CharacterUnitFirstLineIndent 0 0.05
            Assert-Near "TOC $level left indent" $format.LeftIndent $expectedLeft 1.0
            Assert-Near "TOC $level character left indent" $format.CharacterUnitLeftIndent $expectedCharacters 0.05
            if ($format.Alignment -ne 3) { throw "TOC $level must be justified (Alignment=3)." }
            if ($format.LineSpacingRule -ne $expectedRule) {
                throw "TOC $level line-spacing rule expected $expectedRule, got $($format.LineSpacingRule)"
            }
            Assert-Near "TOC $level line spacing" $format.LineSpacing $expectedLine 1.0
            Assert-Near "TOC $level before spacing" $format.SpaceBefore 0
            Assert-Near "TOC $level after spacing" $format.SpaceAfter 0
            $officialTab = @($format.TabStops) | Where-Object {
                [math]::Abs($_.Position - 438.85) -le 0.75 -and
                $_.Alignment -eq 2 -and $_.Leader -eq 1
            } | Select-Object -First 1
            if ($null -eq $officialTab) {
                throw "TOC $level is missing the 438.85 pt right-aligned dot-leader tab."
            }
        }
    }

    if ($captionDirectoryParagraphs.Count -eq 0) {
        throw 'Figure/table directory fields have no materialized Table of Figures paragraphs.'
    }
    foreach ($paragraph in $captionDirectoryParagraphs) {
        $font = $paragraph.Range.Font
        $format = $paragraph.Format
        Assert-Font 'Figure/table directory East Asian font' $font.NameFarEast @('宋体', 'SimSun')
        Assert-Font 'Figure/table directory Latin font' $font.Name @('宋体', 'SimSun')
        Assert-Near 'Figure/table directory size' $font.Size 12
        Assert-Near 'Figure/table directory first-line indent' $format.FirstLineIndent 0
        Assert-Near 'Figure/table directory character first-line indent' $format.CharacterUnitFirstLineIndent 0 0.05
        Assert-Near 'Figure/table directory left indent' $format.LeftIndent 0 1.0
        Assert-Near 'Figure/table directory character left indent' $format.CharacterUnitLeftIndent 0 0.05
        if ($format.Alignment -ne 3) {
            throw 'Figure/table directory entries must be justified (Alignment=3).'
        }
        if ($format.LineSpacingRule -ne 1) {
            throw 'Figure/table directory entries must use 1.5-line spacing.'
        }
        Assert-Near 'Figure/table directory line spacing' $format.LineSpacing 18 1.0
        Assert-Near 'Figure/table directory before spacing' $format.SpaceBefore 0
        Assert-Near 'Figure/table directory after spacing' $format.SpaceAfter 0
        $officialTab = @($format.TabStops) | Where-Object {
            [math]::Abs($_.Position - 438.85) -le 0.75 -and
            $_.Alignment -eq 2 -and $_.Leader -eq 1
        } | Select-Object -First 1
        if ($null -eq $officialTab) {
            throw 'Figure/table directory entry is missing the 438.85 pt right-aligned dot-leader tab.'
        }
    }

    # Existing body checks guard against the official 10.5 pt Normal measuring
    # base leaking into visible 12 pt body and abstract styles.
    foreach ($paragraph in @($document.Paragraphs)) {
        $styleName = Get-StyleName -Paragraph $paragraph
        $paragraphText = Get-ParagraphText -Paragraph $paragraph
        if ([string]::IsNullOrWhiteSpace($paragraphText) -and
            $styleName -in @('First Paragraph', '宜宾论文-首段', 'Body Text', '宜宾论文-正文')) {
            continue
        }
        $font = $paragraph.Range.Font
        $format = $paragraph.Format
        switch ($styleName) {
            { $_ -in @('First Paragraph', '宜宾论文-首段') } {
                Assert-Font 'Body first paragraph font' $font.NameFarEast @('宋体', 'SimSun')
                Assert-Near 'Body first paragraph size' (Get-ParagraphFontSize -Paragraph $paragraph) 12
                Assert-Near 'Body first paragraph first-line indent' $format.FirstLineIndent 24 1.0
            }
            { $_ -in @('Body Text', '宜宾论文-正文') } {
                Assert-Font 'Body text font' $font.NameFarEast @('宋体', 'SimSun')
                Assert-Near 'Body text size' (Get-ParagraphFontSize -Paragraph $paragraph) 12
            }
            { $_ -in @('Heading 1', '标题 1', 'Yibin Heading 1', '宜宾论文-一级标题') } {
                Assert-Font 'Heading 1 font' $font.NameFarEast @('黑体', 'SimHei')
                Assert-Near 'Heading 1 size' (Get-ParagraphFontSize -Paragraph $paragraph) 16
                if ($font.Bold -eq 0) { throw 'Heading 1 must be bold.' }
                if ($format.Alignment -eq 1) {
                    Assert-Near 'Science Heading 1 first-line indent' $format.FirstLineIndent 0
                }
                else {
                    Assert-Near 'Humanities Heading 1 first-line indent' $format.FirstLineIndent 32 1.0
                }
            }
            { $_ -in @('Heading 2', '标题 2', 'Yibin Heading 2', '宜宾论文-二级标题') } {
                Assert-Font 'Heading 2 font' $font.NameFarEast @('楷体', 'KaiTi')
                Assert-Near 'Heading 2 size' (Get-ParagraphFontSize -Paragraph $paragraph) 15
                if ($font.Bold -eq 0) { throw 'Heading 2 must be bold.' }
                Assert-Near 'Heading 2 first-line indent' $format.FirstLineIndent 30 1.0
            }
            { $_ -in @('Heading 3', '标题 3', 'Yibin Heading 3', '宜宾论文-三级标题') } {
                Assert-Font 'Heading 3 font' $font.NameFarEast @('宋体', 'SimSun')
                Assert-Near 'Heading 3 size' (Get-ParagraphFontSize -Paragraph $paragraph) 14
                if ($font.Bold -eq 0) { throw 'Heading 3 must be bold.' }
                Assert-Near 'Heading 3 first-line indent' $format.FirstLineIndent 28 1.0
            }
            { $_ -in @('Heading 4', '标题 4', 'Yibin Heading 4', '宜宾论文-四级标题') } {
                Assert-Font 'Heading 4 font' $font.NameFarEast @('宋体', 'SimSun')
                Assert-Near 'Heading 4 size' (Get-ParagraphFontSize -Paragraph $paragraph) 14
                if ($font.Bold -eq 0) { throw 'Heading 4 must be bold.' }
                Assert-Near 'Heading 4 first-line indent' $format.FirstLineIndent 28 1.0
            }
            { $_ -in @('ChineseAbstract', '宜宾论文-中文摘要正文') } {
                Assert-Font 'Chinese abstract font' $font.NameFarEast @('宋体', 'SimSun')
                Assert-Near 'Chinese abstract size' (Get-ParagraphFontSize -Paragraph $paragraph) 12
                Assert-Near 'Chinese abstract first-line indent' $format.FirstLineIndent 24 1.0
            }
            { $_ -in @('EnglishAbstract', '宜宾论文-英文摘要正文') } {
                Assert-Font 'English abstract font' $font.Name @('Times New Roman')
                Assert-Near 'English abstract size' (Get-ParagraphFontSize -Paragraph $paragraph) 12
                Assert-Near 'English abstract first-line indent' $format.FirstLineIndent 0
            }
        }
    }

    $bodyHeader = $document.Sections.Item($document.Sections.Count).Headers.Item(1).Range
    Assert-Font 'Body header font' $bodyHeader.Font.NameFarEast @('宋体', 'SimSun')
    Assert-Near 'Body header size' $bodyHeader.Font.Size 9

    if (-not [string]::IsNullOrWhiteSpace($PdfOutput)) {
        # 17 = wdExportFormatPDF
        $document.ExportAsFixedFormat($pdfPath, 17)
        Write-Host "PDF exported: $pdfPath"
    }
    Write-Host "Word fields refreshed and audited: $documentPath"
}
finally {
    Close-WordSession -Document $document -Word $word
}
