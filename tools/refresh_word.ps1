#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [string]$PdfOutput
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

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($documentPath, $false, $false)

    foreach ($storyType in 1..17) {
        try {
            $range = $document.StoryRanges.Item($storyType)
            while ($null -ne $range) {
                $range.Fields.Update() | Out-Null
                $range = $range.NextStoryRange
            }
        }
        catch {
            # A document does not necessarily contain every Word story type.
        }
    }
    foreach ($toc in @($document.TablesOfContents)) {
        $toc.Update() | Out-Null
    }

    # Word regenerates TOC paragraphs with its built-in styles.  Reapply the
    # school contract after the field update so Normal's two-character indent
    # is not inherited by directory entries.
    foreach ($level in 1..3) {
        try {
            $tocStyle = $document.Styles.Item("TOC $level")
            $leftIndent = 24 * ($level - 1)
            $tocStyle.Font.NameFarEast = '宋体'
            $tocStyle.Font.Name = 'Times New Roman'
            $tocStyle.Font.Size = 12
            $tocStyle.Font.Bold = 0
            $tocStyle.ParagraphFormat.CharacterUnitFirstLineIndent = 0
            $tocStyle.ParagraphFormat.FirstLineIndent = 0
            $tocStyle.ParagraphFormat.CharacterUnitLeftIndent = 0
            $tocStyle.ParagraphFormat.LeftIndent = $leftIndent
            $tocStyle.ParagraphFormat.SpaceBefore = 0
            $tocStyle.ParagraphFormat.SpaceAfter = 0
            $tocStyle.ParagraphFormat.LineSpacingRule = 1
            $tocStyle.ParagraphFormat.LineSpacing = 18
        }
        catch {
            # A short document may not materialize every TOC level.
        }
    }
    foreach ($paragraph in @($document.Paragraphs)) {
        $styleName = ''
        try {
            $styleName = $paragraph.Range.Style.NameLocal
        }
        catch {
            continue
        }
        $level = 0
        if ($styleName -match '^TOC\s+([1-3])$') {
            $level = [int]$Matches[1]
        }
        if ($level -gt 0) {
            $leftIndent = 24 * ($level - 1)
            $paragraph.Range.Font.NameFarEast = '宋体'
            $paragraph.Range.Font.Name = 'Times New Roman'
            $paragraph.Range.Font.Size = 12
            $paragraph.Range.Font.Bold = 0
            $paragraph.Format.FirstLineIndent = 0
            try { $paragraph.Format.CharacterUnitFirstLineIndent = 0 } catch {}
            $paragraph.Format.LeftIndent = $leftIndent
            try { $paragraph.Format.CharacterUnitLeftIndent = 0 } catch {}
            $paragraph.Format.SpaceBefore = 0
            $paragraph.Format.SpaceAfter = 0
            $paragraph.Format.LineSpacingRule = 1
            $paragraph.Format.LineSpacing = 18

            try {
                $style = $paragraph.Range.Style
                $style.Font.NameFarEast = '宋体'
                $style.Font.Name = 'Times New Roman'
                $style.Font.Size = 12
                $style.Font.Bold = 0
                $style.ParagraphFormat.FirstLineIndent = 0
                $style.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                $style.ParagraphFormat.LeftIndent = $leftIndent
                $style.ParagraphFormat.CharacterUnitLeftIndent = 0
                $style.ParagraphFormat.SpaceBefore = 0
                $style.ParagraphFormat.SpaceAfter = 0
                $style.ParagraphFormat.LineSpacingRule = 1
                $style.ParagraphFormat.LineSpacing = 18
            }
            catch {
                # Direct paragraph formatting above remains authoritative.
            }
        }
    }
    $document.Repaginate()

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

    # Audit Word's computed formatting, not only OOXML declarations. This
    # catches theme-font substitution and inherited first-line indents after
    # Word has regenerated the TOC.
    foreach ($paragraph in @($document.Paragraphs)) {
        $styleName = ''
        try { $styleName = $paragraph.Range.Style.NameLocal } catch { continue }
        $font = $paragraph.Range.Font
        $format = $paragraph.Format
        switch ($styleName) {
            'First Paragraph' {
                Assert-Font 'Body first paragraph font' $font.NameFarEast @('宋体', 'SimSun')
                Assert-Near 'Body first paragraph size' $font.Size 12
                Assert-Near 'Body first paragraph first-line indent' $format.FirstLineIndent 24 1.0
            }
            'Yibin Heading 1' {
                Assert-Font 'Heading 1 font' $font.NameFarEast @('黑体', 'SimHei')
                Assert-Near 'Heading 1 size' $font.Size 16
                if ($font.Bold -eq 0) { throw 'Heading 1 must be bold.' }
                if ($format.Alignment -eq 1) {
                    Assert-Near 'Science Heading 1 first-line indent' $format.FirstLineIndent 0
                }
                else {
                    Assert-Near 'Humanities Heading 1 first-line indent' $format.FirstLineIndent 32 1.0
                }
            }
            'Yibin Heading 2' {
                Assert-Font 'Heading 2 font' $font.NameFarEast @('楷体', 'KaiTi')
                Assert-Near 'Heading 2 size' $font.Size 15
                if ($font.Bold -eq 0) { throw 'Heading 2 must be bold.' }
                Assert-Near 'Heading 2 first-line indent' $format.FirstLineIndent 30 1.0
            }
            'Yibin Heading 3' {
                Assert-Font 'Heading 3 font' $font.NameFarEast @('宋体', 'SimSun')
                Assert-Near 'Heading 3 size' $font.Size 14
                if ($font.Bold -eq 0) { throw 'Heading 3 must be bold.' }
                Assert-Near 'Heading 3 first-line indent' $format.FirstLineIndent 28 1.0
            }
            'Yibin Heading 4' {
                Assert-Font 'Heading 4 font' $font.NameFarEast @('宋体', 'SimSun')
                Assert-Near 'Heading 4 size' $font.Size 14
                if ($font.Bold -eq 0) { throw 'Heading 4 must be bold.' }
                Assert-Near 'Heading 4 first-line indent' $format.FirstLineIndent 28 1.0
            }
            'ChineseAbstract' {
                Assert-Font 'Chinese abstract font' $font.NameFarEast @('宋体', 'SimSun')
                Assert-Near 'Chinese abstract size' $font.Size 12
                Assert-Near 'Chinese abstract first-line indent' $format.FirstLineIndent 24 1.0
            }
            'EnglishAbstract' {
                Assert-Font 'English abstract font' $font.Name @('Times New Roman')
                Assert-Near 'English abstract size' $font.Size 12
                Assert-Near 'English abstract first-line indent' $format.FirstLineIndent 0
            }
            'Yibin TOC Heading' {
                Assert-Font 'TOC heading font' $font.NameFarEast @('黑体', 'SimHei')
                Assert-Near 'TOC heading size' $font.Size 16
                Assert-Near 'TOC heading first-line indent' $format.FirstLineIndent 0
                if ($format.Alignment -ne 1) { throw 'TOC heading must be centered.' }
            }
        }

        if ($styleName -match '^TOC\s+([1-3])$') {
            $level = [int]$Matches[1]
            Assert-Font "TOC $level font" $font.NameFarEast @('宋体', 'SimSun')
            Assert-Near "TOC $level size" $font.Size 12
            Assert-Near "TOC $level first-line indent" $format.FirstLineIndent 0
            Assert-Near "TOC $level left indent" $format.LeftIndent (24 * ($level - 1)) 1.0
        }
    }

    $bodyHeader = $document.Sections.Item($document.Sections.Count).Headers.Item(1).Range
    Assert-Font 'Body header font' $bodyHeader.Font.NameFarEast @('宋体', 'SimSun')
    Assert-Near 'Body header size' $bodyHeader.Font.Size 9

    $document.Save()

    if (-not [string]::IsNullOrWhiteSpace($PdfOutput)) {
        # 17 = wdExportFormatPDF
        $document.ExportAsFixedFormat($pdfPath, 17)
        Write-Host "PDF exported: $pdfPath"
    }
    Write-Host "Word fields refreshed: $documentPath"
}
finally {
    if ($null -ne $document) {
        $document.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($null -ne $word) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
