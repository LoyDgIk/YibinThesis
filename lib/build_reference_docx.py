#!/usr/bin/env python3
"""Generate the semantic Pandoc reference document for YibinThesis.

The output is a derived, editable Word style package.  It does not copy the
content or package structure of the University's 2024 Word example.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


SIMSUN = "SimSun"
SIMHEI = "SimHei"
KAITI = "KaiTi"
TIMES = "Times New Roman"

STYLE_BODY = "宜宾论文-正文"
STYLE_FIRST_PARAGRAPH = "宜宾论文-首段"
STYLE_LIST_BODY = "宜宾论文-列表正文"
STYLE_HEADING_1 = "宜宾论文-一级标题"
STYLE_HEADING_2 = "宜宾论文-二级标题"
STYLE_HEADING_3 = "宜宾论文-三级标题"
STYLE_HEADING_4 = "宜宾论文-四级标题"
STYLE_UNNUMBERED_HEADING = "宜宾论文-无编号标题"
STYLE_FRONT_TITLE = "宜宾论文-中文页标题"
STYLE_ENGLISH_ABSTRACT_TITLE = "宜宾论文-英文摘要标题"
STYLE_TOC_TITLE = "宜宾论文-目录标题"
STYLE_TABLE_OF_FIGURES = "Table of Figures"
STYLE_APPENDIX_HEADING = "宜宾论文-附录标题"
STYLE_APPENDIX_SECTION = "宜宾论文-附录二级标题"
STYLE_CHINESE_ABSTRACT = "宜宾论文-中文摘要正文"
STYLE_ENGLISH_ABSTRACT = "宜宾论文-英文摘要正文"
STYLE_KEYWORDS = "宜宾论文-关键词"
STYLE_FIGURE = "宜宾论文-插图"
STYLE_EQUATION = "宜宾论文-公式"
STYLE_FIGURE_CAPTION = "宜宾论文-图题"
STYLE_TABLE_CAPTION = "宜宾论文-表题"
STYLE_TABLE_CONTINUATION = "宜宾论文-续表题"
STYLE_TABLE_TEXT = "宜宾论文-表格正文"
STYLE_TABLE_CENTER = "宜宾论文-表格居中"
STYLE_TABLE_RIGHT = "宜宾论文-表格右对齐"
STYLE_TABLE_HEADER = "宜宾论文-表头"
STYLE_TABLE_HEADER_LEFT = "宜宾论文-表头左对齐"
STYLE_TABLE_HEADER_RIGHT = "宜宾论文-表头右对齐"
STYLE_BIBLIOGRAPHY = "宜宾论文-参考文献"
STYLE_NOTES = "宜宾论文-注释"
STYLE_CITATION = "宜宾论文-文献上标"
STYLE_THREE_LINE_TABLE = "宜宾论文-三线表"
STYLE_PROPOSAL_TITLE = "宜宾开题-标题"
STYLE_PROPOSAL_SUBTITLE = "宜宾开题-副标题"
STYLE_PROPOSAL_LABEL = "宜宾开题-栏目"
STYLE_PROPOSAL_BODY = "宜宾开题-正文"
STYLE_PROPOSAL_PROMPT = "宜宾开题-提示"
STYLE_PROPOSAL_UNNUMBERED_HEADING = "宜宾开题-无编号标题"
STYLE_PROPOSAL_SIGNATURE = "宜宾开题-签名"
STYLE_REVIEW_DOCUMENT_TITLE = "宜宾综述-文档标题"
STYLE_REVIEW_THESIS_TITLE = "宜宾综述-论文题目"
STYLE_REVIEW_INFO_LABEL = "宜宾综述-信息标签"
STYLE_REVIEW_INFO_VALUE = "宜宾综述-信息值"
STYLE_REVIEW_DATE = "宜宾综述-日期"
STYLE_COVER_LAYOUT_TABLE = "YibinCoverLayout"
STYLE_FRONT_LAYOUT_TABLE = "YibinFrontLayout"
STYLE_PROPOSAL_FORM_TABLE = "YibinProposalForm"
STYLE_REVIEW_INFO_TABLE = "YibinReviewInfoLayout"


def _get_or_add_style(document: Document, name: str, style_type: WD_STYLE_TYPE):
    try:
        return document.styles[name]
    except KeyError:
        return document.styles.add_style(name, style_type)


def _set_fonts(
    style,
    *,
    east_asia: str,
    latin: str = TIMES,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
) -> None:
    style.font.name = latin
    style.font.bold = bold
    style.font.italic = italic
    if size is not None:
        style.font.size = Pt(size)

    r_pr = style.element.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), latin)
    r_fonts.set(qn("w:hAnsi"), latin)
    r_fonts.set(qn("w:cs"), latin)
    r_fonts.set(qn("w:eastAsia"), east_asia)
    r_fonts.set(qn("w:hint"), "eastAsia")

    color = r_pr.find(qn("w:color"))
    if color is None:
        color = OxmlElement("w:color")
        r_pr.append(color)
    color.set(qn("w:val"), "000000")
    for attribute in ("w:themeColor", "w:themeTint", "w:themeShade"):
        color.attrib.pop(qn(attribute), None)

    if bold is not None:
        for tag in ("w:b", "w:bCs"):
            node = r_pr.find(qn(tag))
            if node is None:
                node = OxmlElement(tag)
                r_pr.append(node)
            node.set(qn("w:val"), "1" if bold else "0")


def _set_first_line_chars(style, chars: int = 200) -> None:
    p_pr = style.element.get_or_add_pPr()
    ind = p_pr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        p_pr.append(ind)
    ind.set(qn("w:firstLineChars"), str(chars))


def _set_outline_level(style, level: int) -> None:
    p_pr = style.element.get_or_add_pPr()
    outline = p_pr.find(qn("w:outlineLvl"))
    if outline is None:
        outline = OxmlElement("w:outlineLvl")
        p_pr.append(outline)
    outline.set(qn("w:val"), str(level))


def _set_indent_xml(
    style,
    *,
    left: int = 0,
    left_chars: int = 0,
    right: int = 0,
) -> None:
    """Write deterministic Word indent values in twips/character units."""

    p_pr = style.element.get_or_add_pPr()
    ind = p_pr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        p_pr.append(ind)
    for attribute, value in (
        ("w:left", left),
        ("w:leftChars", left_chars),
        ("w:right", right),
        ("w:rightChars", 0),
        ("w:firstLine", 0),
        ("w:firstLineChars", 0),
        ("w:hanging", 0),
        ("w:hangingChars", 0),
    ):
        ind.set(qn(attribute), str(value))


def _set_style_tab_stop(
    style,
    *,
    position_twips: int,
    alignment: str = "right",
    leader: str = "dot",
) -> None:
    p_pr = style.element.get_or_add_pPr()
    existing = p_pr.find(qn("w:tabs"))
    if existing is not None:
        p_pr.remove(existing)
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), alignment)
    tab.set(qn("w:leader"), leader)
    tab.set(qn("w:pos"), str(position_twips))
    tabs.append(tab)
    p_pr.append(tabs)


def _set_style_bottom_border(style) -> None:
    """Add the single reusable underline rule used by cover value slots."""

    p_pr = style.element.get_or_add_pPr()
    borders = p_pr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        p_pr.append(borders)
    bottom = borders.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        borders.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), "000000")


def _set_keep(style, *, next_paragraph: bool = False, lines: bool = False) -> None:
    p_pr = style.element.get_or_add_pPr()
    for tag, enabled in (("w:keepNext", next_paragraph), ("w:keepLines", lines)):
        node = p_pr.find(qn(tag))
        if enabled and node is None:
            p_pr.append(OxmlElement(tag))
        elif not enabled and node is not None:
            p_pr.remove(node)


def _set_quick_style(style, priority: int) -> None:
    """Expose a reusable thesis style in Word's Quick Style gallery."""

    style.hidden = False
    style.quick_style = True
    style.priority = priority
    style.unhide_when_used = False
    element = style.element
    for tag in ("w:semiHidden", "w:unhideWhenUsed"):
        node = element.find(qn(tag))
        if node is not None:
            element.remove(node)


def _set_style_tab_stops(
    style,
    stops: tuple[tuple[int, str, str | None], ...],
) -> None:
    p_pr = style.element.get_or_add_pPr()
    existing = p_pr.find(qn("w:tabs"))
    if existing is not None:
        p_pr.remove(existing)
    tabs = OxmlElement("w:tabs")
    for position_twips, alignment, leader in stops:
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), alignment)
        tab.set(qn("w:pos"), str(position_twips))
        if leader:
            tab.set(qn("w:leader"), leader)
        tabs.append(tab)
    p_pr.append(tabs)


def _set_doc_defaults(document: Document) -> None:
    styles = document.styles.element
    defaults = styles.find(qn("w:docDefaults"))
    if defaults is None:
        defaults = OxmlElement("w:docDefaults")
        styles.insert(0, defaults)

    r_pr_default = defaults.find(qn("w:rPrDefault"))
    if r_pr_default is None:
        r_pr_default = OxmlElement("w:rPrDefault")
        defaults.append(r_pr_default)
    r_pr = r_pr_default.find(qn("w:rPr"))
    if r_pr is None:
        r_pr = OxmlElement("w:rPr")
        r_pr_default.append(r_pr)
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr, value in (
        ("w:ascii", TIMES),
        ("w:hAnsi", TIMES),
        ("w:cs", TIMES),
        ("w:eastAsia", SIMSUN),
        ("w:hint", "eastAsia"),
    ):
        r_fonts.set(qn(attr), value)
    size = r_pr.find(qn("w:sz"))
    if size is None:
        size = OxmlElement("w:sz")
        r_pr.append(size)
    size.set(qn("w:val"), "24")
    size_cs = r_pr.find(qn("w:szCs"))
    if size_cs is None:
        size_cs = OxmlElement("w:szCs")
        r_pr.append(size_cs)
    size_cs.set(qn("w:val"), "24")

    p_pr_default = defaults.find(qn("w:pPrDefault"))
    if p_pr_default is None:
        p_pr_default = OxmlElement("w:pPrDefault")
        defaults.append(p_pr_default)
    p_pr = p_pr_default.find(qn("w:pPr"))
    if p_pr is None:
        p_pr = OxmlElement("w:pPr")
        p_pr_default.append(p_pr)
    spacing = p_pr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        p_pr.append(spacing)
    spacing.set(qn("w:line"), "360")
    spacing.set(qn("w:lineRule"), "auto")


def _configure_paragraph_style(
    document: Document,
    name: str,
    *,
    east_asia: str = SIMSUN,
    latin: str = TIMES,
    size: float = 12,
    bold: bool = False,
    alignment: WD_ALIGN_PARAGRAPH | None = None,
    first_line: bool = False,
    line_spacing: float = 1.5,
    before: float = 0,
    after: float = 0,
    keep_next: bool = False,
    keep_lines: bool = False,
    page_break_before: bool = False,
    base: str | None = None,
    outline_level: int | None = None,
):
    style = _get_or_add_style(document, name, WD_STYLE_TYPE.PARAGRAPH)
    if base:
        style.base_style = document.styles[base]
    _set_fonts(
        style,
        east_asia=east_asia,
        latin=latin,
        size=size,
        bold=bold,
    )
    fmt = style.paragraph_format
    fmt.alignment = alignment
    fmt.line_spacing = line_spacing
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.first_line_indent = Pt(24) if first_line else Pt(0)
    fmt.keep_with_next = keep_next
    fmt.keep_together = keep_lines
    fmt.page_break_before = page_break_before
    fmt.widow_control = True
    if first_line:
        _set_first_line_chars(style)
    else:
        # Block a two-character indent inherited from Normal/Heading styles.
        _set_first_line_chars(style, 0)
    if outline_level is not None:
        _set_outline_level(style, outline_level)
    _set_keep(style, next_paragraph=keep_next, lines=keep_lines)
    return style


def _configure_character_style(
    document: Document,
    name: str,
    *,
    east_asia: str,
    latin: str = TIMES,
    size: float,
    bold: bool = False,
    superscript: bool = False,
    quick_priority: int | None = None,
):
    style = _get_or_add_style(document, name, WD_STYLE_TYPE.CHARACTER)
    _set_fonts(
        style,
        east_asia=east_asia,
        latin=latin,
        size=size,
        bold=bold,
    )
    style.font.superscript = superscript
    if superscript:
        r_pr = style.element.get_or_add_rPr()
        vertical = r_pr.find(qn("w:vertAlign"))
        if vertical is None:
            vertical = OxmlElement("w:vertAlign")
            r_pr.append(vertical)
        vertical.set(qn("w:val"), "superscript")
    if quick_priority is not None:
        _set_quick_style(style, quick_priority)
    return style


def _configure_three_line_table_style(document: Document):
    style = _get_or_add_style(document, STYLE_THREE_LINE_TABLE, WD_STYLE_TYPE.TABLE)
    _set_fonts(style, east_asia=SIMSUN, latin=TIMES, size=10.5, bold=False)
    table_properties = style.element.find(qn("w:tblPr"))
    if table_properties is None:
        table_properties = OxmlElement("w:tblPr")
        style.element.append(table_properties)
    borders = table_properties.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        table_properties.append(borders)
    for edge, value, size in (
        ("top", "single", "12"),
        ("bottom", "single", "12"),
        ("left", "nil", "0"),
        ("right", "nil", "0"),
        ("insideH", "nil", "0"),
        ("insideV", "nil", "0"),
    ):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), value)
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), "000000")
    _set_quick_style(style, 23)
    return style


def _configure_layout_table_style(
    document: Document,
    name: str,
    *,
    priority: int,
    grid: bool,
):
    """Register reusable table roles used by deterministic layout geometry."""

    style = _get_or_add_style(document, name, WD_STYLE_TYPE.TABLE)
    _set_fonts(style, east_asia=SIMSUN, latin=TIMES, size=10.5, bold=False)
    table_properties = style.element.find(qn("w:tblPr"))
    if table_properties is None:
        table_properties = OxmlElement("w:tblPr")
        style.element.append(table_properties)
    layout = table_properties.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        table_properties.append(layout)
    layout.set(qn("w:type"), "fixed")
    borders = table_properties.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        table_properties.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single" if grid else "nil")
        node.set(qn("w:sz"), "4" if grid else "0")
        node.set(qn("w:color"), "000000")
    _set_quick_style(style, priority)
    return style


def build_reference_docx(output: Path) -> Path:
    document = Document()
    _set_doc_defaults(document)

    section = document.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.5)
    section.header_distance = Cm(1.5)
    section.footer_distance = Cm(1.5)

    _configure_paragraph_style(
        document,
        "Normal",
        size=10.5,
        first_line=True,
        line_spacing=1.5,
    )
    for name in ("Body Text", "First Paragraph"):
        _configure_paragraph_style(
            document,
            name,
            first_line=True,
            line_spacing=1.5,
            base="Normal",
        )
    _configure_paragraph_style(
        document,
        "Compact",
        first_line=True,
        line_spacing=1.5,
        base="Normal",
    )

    # Keep Word's built-in Heading 1-4 styles untouched.  Thesis headings use
    # the independent, reusable 宜宾论文-* styles registered below; Pandoc's
    # built-in heading output is mapped to them during OOXML post-processing.

    _configure_paragraph_style(
        document,
        "Title",
        east_asia=SIMHEI,
        size=18,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        line_spacing=1.5,
        after=12,
    )
    _configure_paragraph_style(
        document,
        "Subtitle",
        east_asia=SIMSUN,
        size=16,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        line_spacing=1.5,
    )
    _configure_paragraph_style(
        document,
        "Header",
        east_asia=SIMSUN,
        size=9,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        line_spacing=1.0,
    )
    _configure_paragraph_style(
        document,
        "Footer",
        east_asia=SIMSUN,
        size=9,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        line_spacing=1.0,
    )
    _configure_paragraph_style(
        document,
        "Caption",
        east_asia=SIMSUN,
        size=10.5,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        line_spacing=1.0,
        before=3,
        after=3,
        keep_next=False,
        keep_lines=True,
    )
    for name in ("Table Caption", "Image Caption", "Figure Caption"):
        _configure_paragraph_style(
            document,
            name,
            east_asia=SIMSUN,
            size=10.5,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            line_spacing=1.0,
            before=3,
            after=3,
            keep_next=name == "Table Caption",
            keep_lines=True,
            base="Caption",
        )
    _configure_paragraph_style(
        document,
        "Bibliography",
        east_asia=SIMSUN,
        size=10.5,
        line_spacing=1.5,
        before=0,
        after=0,
    )
    bibliography = document.styles["Bibliography"]
    bibliography.paragraph_format.left_indent = Cm(0.74)
    bibliography.paragraph_format.first_line_indent = Cm(-0.74)
    _configure_paragraph_style(
        document,
        "Footnote Text",
        east_asia=SIMSUN,
        size=9,
        line_spacing=1.0,
    )

    _configure_paragraph_style(
        document,
        "TOC Heading",
        east_asia=SIMHEI,
        latin=SIMHEI,
        size=16,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        line_spacing=1.0,
        before=0,
        after=0,
        keep_next=True,
    )
    for level in range(1, 4):
        toc = _configure_paragraph_style(
            document,
            f"TOC {level}",
            east_asia=SIMSUN,
            latin=SIMSUN,
            size=12,
            line_spacing=1.0,
            alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
            base="Normal",
        )
        toc.paragraph_format.first_line_indent = Pt(0)
        toc.paragraph_format.left_indent = Pt((level - 1) * 21)
        toc.paragraph_format.right_indent = Pt(0)
        # Word's official template uses auto line spacing (1.5 lines), not an
        # exact 18 pt line height.  The visible result is 18 pt at 12 pt text.
        toc.paragraph_format.line_spacing = 1.5 if level < 3 else 1.0
        toc.paragraph_format.space_before = Pt(0)
        toc.paragraph_format.space_after = Pt(0)
        _set_indent_xml(
            toc,
            left=(level - 1) * 420,
            left_chars=(level - 1) * 200,
        )
        _set_style_tab_stop(toc, position_twips=8777)

    # Word uses one built-in Table of Figures style for both the figure and
    # table directories produced by TOC \c fields. Match the first-level TOC
    # contract so every generated directory entry is Song small-four.
    caption_directory = _configure_paragraph_style(
        document,
        STYLE_TABLE_OF_FIGURES,
        east_asia=SIMSUN,
        latin=SIMSUN,
        size=12,
        line_spacing=1.5,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        base="Normal",
    )
    # ``add_style`` marks newly materialized latent styles as custom. Removing
    # that flag lets Word bind the OOXML style to built-in id -36 instead of
    # creating a second, default-formatted 图表目录 style when fields update.
    caption_directory.element.attrib.pop(qn("w:customStyle"), None)
    caption_directory.paragraph_format.first_line_indent = Pt(0)
    caption_directory.paragraph_format.left_indent = Pt(0)
    caption_directory.paragraph_format.right_indent = Pt(0)
    caption_directory.paragraph_format.space_before = Pt(0)
    caption_directory.paragraph_format.space_after = Pt(0)
    _set_indent_xml(caption_directory)
    _set_style_tab_stop(caption_directory, position_twips=8777)

    custom_paragraphs = (
        ("CoverLogo", SIMHEI, 12, False, WD_ALIGN_PARAGRAPH.CENTER, 1.0, 2.0, 0, False, False, None),
        ("CoverSchool", SIMHEI, 26, True, WD_ALIGN_PARAGRAPH.CENTER, 1.0, 0, 12, False, False, None),
        ("CoverThesisType", SIMHEI, 28, True, WD_ALIGN_PARAGRAPH.CENTER, 1.0, 24.35, 52.35, False, False, None),
        ("CoverVersion", SIMHEI, 14, False, WD_ALIGN_PARAGRAPH.CENTER, 1.0, 0, 23, False, False, None),
        ("CoverTitle", SIMHEI, 18, True, WD_ALIGN_PARAGRAPH.CENTER, 2.0, 0, 47.3, False, False, None),
        ("CoverTitleWithVersion", SIMHEI, 18, True, WD_ALIGN_PARAGRAPH.CENTER, 2.0, 0, 10, False, False, None),
        ("CoverField", SIMSUN, 16, True, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.0, 0, 0, False, False, None),
        ("CoverValueLine", SIMSUN, 16, True, WD_ALIGN_PARAGRAPH.CENTER, 1.0, 0, 0, False, False, None),
        ("CoverDate", SIMSUN, 15, True, WD_ALIGN_PARAGRAPH.CENTER, 1.0, 24, 0, False, False, None),
        ("Yibin Table Continuation", SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.RIGHT, 1.0, 0, 0, True, True, None),
        ("DeclarationTitle", SIMHEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, 1.5, 0, 31.2, True, False, None),
        ("DeclarationBody", SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.LEFT, 2.0, 0, 0, False, False, None),
        ("DeclarationSignature", SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.LEFT, 1.5, 36, 0, False, False, None),
        ("DeclarationDate", SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.LEFT, 1.5, 18, 0, False, False, None),
        ("DeclarationRegulationLead", SIMSUN, 14, True, WD_ALIGN_PARAGRAPH.LEFT, 1.0, 187, 0, False, False, None),
        ("DeclarationRegulationClause", SIMSUN, 14, True, WD_ALIGN_PARAGRAPH.LEFT, 1.5, 0, 0, False, False, None),
        ("DeclarationRegulationItem", SIMSUN, 14, True, WD_ALIGN_PARAGRAPH.LEFT, 1.5, 0, 0, False, False, None),
        ("FrontTitle", SIMHEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, 1.5, 0, 12, True, False, 0),
        ("Yibin TOC Heading", SIMHEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, 1.0, 0, 0, True, False, None),
        ("ChineseAbstract", SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.5, 0, 0, False, False, None),
        ("EnglishAbstract", TIMES, 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.5, 0, 0, False, False, None),
        ("Keywords", SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.LEFT, 1.5, 6, 0, False, False, None),
        ("Unnumbered Heading 1", SIMHEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, 1.5, 0, 12, True, True, 0),
        ("Notes", SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.5, 0, 0, False, False, None),
        ("YibinSectionMarker", SIMSUN, 1, False, WD_ALIGN_PARAGRAPH.LEFT, 1.0, 0, 0, False, False, None),
    )
    for (
        name,
        font,
        size,
        bold,
        alignment,
        line_spacing,
        before,
        after,
        keep_next,
        page_break,
        outline,
    ) in custom_paragraphs:
        _configure_paragraph_style(
            document,
            name,
            east_asia=font,
            latin=TIMES,
            size=size,
            bold=bold,
            alignment=alignment,
            line_spacing=line_spacing,
            before=before,
            after=after,
            keep_next=keep_next,
            keep_lines=keep_next,
            page_break_before=page_break,
            outline_level=outline,
        )
    document.styles["CoverTitle"].font.underline = True
    document.styles["CoverTitleWithVersion"].font.underline = True
    cover_value_line = document.styles["CoverValueLine"]
    cover_value_line.base_style = document.styles["Normal"]
    cover_value_line.paragraph_format.line_spacing = Pt(20)
    _set_style_bottom_border(cover_value_line)
    for regulation_style_name in (
        "DeclarationRegulationLead",
        "DeclarationRegulationClause",
        "DeclarationRegulationItem",
    ):
        _set_fonts(
            document.styles[regulation_style_name],
            east_asia=SIMSUN,
            latin=SIMSUN,
            size=14,
            bold=True,
        )
    for declaration_style_name in (
        "DeclarationBody",
        "DeclarationSignature",
        "DeclarationDate",
    ):
        declaration_style = document.styles[declaration_style_name]
        declaration_style.paragraph_format.first_line_indent = Pt(24)
        _set_first_line_chars(declaration_style)
    _set_first_line_chars(document.styles["ChineseAbstract"])

    regulation_clause = document.styles["DeclarationRegulationClause"]
    regulation_clause.paragraph_format.first_line_indent = Pt(35)
    clause_ppr = regulation_clause.element.get_or_add_pPr()
    clause_ind = clause_ppr.find(qn("w:ind"))
    if clause_ind is None:
        clause_ind = OxmlElement("w:ind")
        clause_ppr.append(clause_ind)
    clause_ind.set(qn("w:firstLine"), "700")
    clause_ind.set(qn("w:firstLineChars"), "249")

    regulation_item = document.styles["DeclarationRegulationItem"]
    regulation_item.paragraph_format.left_indent = Pt(7.15)
    item_ppr = regulation_item.element.get_or_add_pPr()
    item_ind = item_ppr.find(qn("w:ind"))
    if item_ind is None:
        item_ind = OxmlElement("w:ind")
        item_ppr.append(item_ind)
    item_ind.set(qn("w:left"), "143")
    item_ind.set(qn("w:firstLine"), "0")
    item_ind.set(qn("w:firstLineChars"), "0")

    _configure_character_style(
        document,
        "CoverLabel",
        east_asia=SIMHEI,
        size=18,
        bold=True,
    )
    _configure_character_style(
        document,
        "CoverValue",
        east_asia=SIMSUN,
        size=16,
        bold=True,
    )
    document.styles["CoverValue"].font.underline = False
    _configure_character_style(
        document,
        "DeclarationRegulationLeadLabel",
        east_asia=SIMSUN,
        latin=SIMSUN,
        size=16,
        bold=True,
    )
    _configure_character_style(
        document,
        "DeclarationRegulationText",
        east_asia=SIMSUN,
        latin=SIMSUN,
        size=14,
        bold=True,
    )
    _configure_character_style(
        document,
        "KeywordLabel",
        east_asia=SIMHEI,
        size=12,
        bold=True,
    )

    # User-facing thesis styles.  The 2024 humanities DOCX relies heavily on
    # direct formatting, so these reusable styles distill the observable
    # template values together with the explicit school-level requirements.
    visible_paragraph_styles = (
        (STYLE_BODY, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, True, 1.5, 0, 0, False, False, 0, None),
        (STYLE_FIRST_PARAGRAPH, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, True, 1.5, 0, 0, False, False, 1, None),
        (STYLE_LIST_BODY, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, False, 1.5, 0, 0, False, False, 11, None),
        (STYLE_HEADING_1, SIMHEI, 16, True, WD_ALIGN_PARAGRAPH.LEFT, True, 1.5, 12, 6, True, True, 2, 0),
        (STYLE_HEADING_2, KAITI, 15, True, WD_ALIGN_PARAGRAPH.LEFT, True, 1.5, 9, 6, True, False, 3, 1),
        (STYLE_HEADING_3, SIMSUN, 14, True, WD_ALIGN_PARAGRAPH.LEFT, True, 1.5, 6, 3, True, False, 4, 2),
        (STYLE_HEADING_4, SIMSUN, 14, True, WD_ALIGN_PARAGRAPH.LEFT, True, 1.5, 6, 3, True, False, 5, 3),
        (STYLE_UNNUMBERED_HEADING, SIMHEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.5, 0, 12, True, True, 6, 0),
        (STYLE_FRONT_TITLE, SIMHEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.5, 0, 12, True, True, 7, 0),
        (STYLE_ENGLISH_ABSTRACT_TITLE, TIMES, 16, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.5, 0, 12, True, True, 8, 0),
        (STYLE_TOC_TITLE, SIMHEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 0, 0, True, False, 9, None),
        (STYLE_APPENDIX_HEADING, SIMHEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.5, 0, 12, True, True, 10, 0),
        (STYLE_APPENDIX_SECTION, KAITI, 15, True, WD_ALIGN_PARAGRAPH.LEFT, True, 1.5, 9, 6, True, False, 62, 1),
        (STYLE_CHINESE_ABSTRACT, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, True, 1.5, 0, 0, False, False, 12, None),
        (STYLE_ENGLISH_ABSTRACT, TIMES, 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, False, 1.5, 0, 0, False, False, 13, None),
        (STYLE_KEYWORDS, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.LEFT, False, 1.5, 6, 0, False, False, 14, None),
        (STYLE_FIGURE, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 0, 0, True, False, 19, None),
        (STYLE_EQUATION, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.LEFT, False, 1.0, 3, 3, False, False, 32, None),
        (STYLE_FIGURE_CAPTION, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 3, 3, False, False, 20, None),
        (STYLE_TABLE_CAPTION, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 3, 3, True, False, 21, None),
        (STYLE_TABLE_CONTINUATION, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.RIGHT, False, 1.0, 0, 0, True, True, 22, None),
        (STYLE_TABLE_TEXT, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.LEFT, False, 1.0, 0, 0, False, False, 23, None),
        (STYLE_TABLE_CENTER, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 0, 0, False, False, 24, None),
        (STYLE_TABLE_RIGHT, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.RIGHT, False, 1.0, 0, 0, False, False, 25, None),
        (STYLE_TABLE_HEADER, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 0, 0, True, False, 26, None),
        (STYLE_TABLE_HEADER_LEFT, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.LEFT, False, 1.0, 0, 0, True, False, 27, None),
        (STYLE_TABLE_HEADER_RIGHT, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.RIGHT, False, 1.0, 0, 0, True, False, 28, None),
        (STYLE_BIBLIOGRAPHY, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.JUSTIFY, False, 1.5, 0, 0, False, False, 29, None),
        (STYLE_NOTES, SIMSUN, 10.5, False, WD_ALIGN_PARAGRAPH.JUSTIFY, False, 1.5, 0, 0, False, False, 30, None),
        (STYLE_PROPOSAL_TITLE, KAITI, 18, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 0, 6, True, False, 50, None),
        (STYLE_PROPOSAL_SUBTITLE, SIMSUN, 12, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 0, 6, True, False, 51, None),
        (STYLE_PROPOSAL_LABEL, SIMHEI, 12, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 0, 0, False, False, 52, None),
        (STYLE_PROPOSAL_BODY, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, True, 1.5, 0, 0, False, False, 53, None),
        (STYLE_PROPOSAL_PROMPT, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.LEFT, False, 1.5, 0, 0, True, False, 54, None),
        (STYLE_PROPOSAL_UNNUMBERED_HEADING, KAITI, 15, True, WD_ALIGN_PARAGRAPH.LEFT, True, 1.5, 9, 6, True, False, 55, None),
        (STYLE_PROPOSAL_SIGNATURE, SIMSUN, 12, False, WD_ALIGN_PARAGRAPH.RIGHT, False, 1.5, 0, 0, True, False, 56, None),
        (STYLE_REVIEW_DOCUMENT_TITLE, SIMHEI, 28, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 42, 72, True, False, 57, None),
        (STYLE_REVIEW_THESIS_TITLE, SIMHEI, 18, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.5, 0, 78, True, False, 58, None),
        (STYLE_REVIEW_INFO_LABEL, SIMHEI, 18, True, WD_ALIGN_PARAGRAPH.LEFT, False, 1.0, 0, 0, True, False, 59, None),
        (STYLE_REVIEW_INFO_VALUE, SIMSUN, 16, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 0, 0, True, False, 60, None),
        (STYLE_REVIEW_DATE, SIMSUN, 16, True, WD_ALIGN_PARAGRAPH.CENTER, False, 1.0, 0, 0, True, False, 61, None),
    )
    quick_paragraph_styles = {
        STYLE_BODY,
        STYLE_LIST_BODY,
        STYLE_HEADING_1,
        STYLE_HEADING_2,
        STYLE_HEADING_3,
        STYLE_HEADING_4,
        STYLE_UNNUMBERED_HEADING,
        STYLE_FRONT_TITLE,
        STYLE_ENGLISH_ABSTRACT_TITLE,
        STYLE_TOC_TITLE,
        STYLE_APPENDIX_HEADING,
        STYLE_APPENDIX_SECTION,
        STYLE_CHINESE_ABSTRACT,
        STYLE_ENGLISH_ABSTRACT,
        STYLE_KEYWORDS,
        STYLE_FIGURE,
        STYLE_EQUATION,
        STYLE_FIGURE_CAPTION,
        STYLE_TABLE_CAPTION,
        STYLE_TABLE_CONTINUATION,
        STYLE_TABLE_TEXT,
        STYLE_TABLE_CENTER,
        STYLE_TABLE_RIGHT,
        STYLE_TABLE_HEADER,
        STYLE_TABLE_HEADER_LEFT,
        STYLE_TABLE_HEADER_RIGHT,
        STYLE_BIBLIOGRAPHY,
        STYLE_NOTES,
        STYLE_PROPOSAL_TITLE,
        STYLE_PROPOSAL_SUBTITLE,
        STYLE_PROPOSAL_LABEL,
        STYLE_PROPOSAL_BODY,
        STYLE_PROPOSAL_PROMPT,
        STYLE_PROPOSAL_UNNUMBERED_HEADING,
        STYLE_PROPOSAL_SIGNATURE,
        STYLE_REVIEW_DOCUMENT_TITLE,
        STYLE_REVIEW_THESIS_TITLE,
        STYLE_REVIEW_INFO_LABEL,
        STYLE_REVIEW_INFO_VALUE,
        STYLE_REVIEW_DATE,
    }
    for (
        name,
        font,
        size,
        bold,
        alignment,
        first_line,
        line_spacing,
        before,
        after,
        keep_next,
        page_break,
        priority,
        outline,
    ) in visible_paragraph_styles:
        style = _configure_paragraph_style(
            document,
            name,
            east_asia=font,
            latin=TIMES,
            size=size,
            bold=bold,
            alignment=alignment,
            first_line=first_line,
            line_spacing=line_spacing,
            before=before,
            after=after,
            keep_next=keep_next,
            keep_lines=keep_next,
            page_break_before=page_break,
            base="Caption" if name in {STYLE_FIGURE_CAPTION, STYLE_TABLE_CAPTION} else "Normal",
            outline_level=outline,
        )
        if name in quick_paragraph_styles:
            _set_quick_style(style, priority)

    _set_style_tab_stops(
        document.styles[STYLE_EQUATION],
        (
            (4394, "center", None),
            (8787, "right", None),
        ),
    )

    bibliography_visible = document.styles[STYLE_BIBLIOGRAPHY]
    bibliography_visible.paragraph_format.left_indent = Cm(0.74)
    bibliography_visible.paragraph_format.first_line_indent = Cm(-0.74)
    table_continuation_visible = document.styles[STYLE_TABLE_CONTINUATION]
    table_continuation_visible.base_style = document.styles["Caption"]
    _set_style_bottom_border(document.styles[STYLE_REVIEW_INFO_VALUE])

    citation_style = _configure_character_style(
        document,
        STYLE_CITATION,
        east_asia=SIMSUN,
        latin=TIMES,
        size=9,
        superscript=True,
        quick_priority=31,
    )
    citation_style.base_style = document.styles["Default Paragraph Font"]
    _configure_three_line_table_style(document)
    _configure_layout_table_style(
        document,
        STYLE_COVER_LAYOUT_TABLE,
        priority=61,
        grid=False,
    )

    _configure_layout_table_style(
        document,
        STYLE_FRONT_LAYOUT_TABLE,
        priority=62,
        grid=False,
    )
    _configure_layout_table_style(
        document,
        STYLE_PROPOSAL_FORM_TABLE,
        priority=63,
        grid=True,
    )
    _configure_layout_table_style(
        document,
        STYLE_REVIEW_INFO_TABLE,
        priority=64,
        grid=False,
    )

    # Word-generated directory and Caption paragraphs keep their semantic
    # style IDs.  Expose those styles in the gallery so users can modify the
    # official directory hierarchy and standard Caption base directly.
    for priority, style_name in enumerate(
        ("Caption", "TOC 1", "TOC 2", "TOC 3", STYLE_TABLE_OF_FIGURES),
        start=40,
    ):
        _set_quick_style(document.styles[style_name], priority)

    for style_name in (
        STYLE_HEADING_1,
        STYLE_HEADING_2,
        STYLE_HEADING_3,
        STYLE_HEADING_4,
        STYLE_UNNUMBERED_HEADING,
        STYLE_FRONT_TITLE,
        STYLE_ENGLISH_ABSTRACT_TITLE,
        STYLE_APPENDIX_HEADING,
        STYLE_APPENDIX_SECTION,
        STYLE_EQUATION,
        STYLE_PROPOSAL_TITLE,
        STYLE_PROPOSAL_SUBTITLE,
        STYLE_PROPOSAL_UNNUMBERED_HEADING,
        STYLE_REVIEW_DOCUMENT_TITLE,
        STYLE_REVIEW_THESIS_TITLE,
    ):
        document.styles[style_name].next_paragraph_style = document.styles[STYLE_BODY]
    document.styles[STYLE_BODY].next_paragraph_style = document.styles[STYLE_BODY]
    document.styles[STYLE_FIRST_PARAGRAPH].next_paragraph_style = document.styles[STYLE_BODY]
    document.styles[STYLE_LIST_BODY].next_paragraph_style = document.styles[STYLE_LIST_BODY]
    document.styles[STYLE_FIGURE_CAPTION].next_paragraph_style = document.styles[STYLE_BODY]
    document.styles[STYLE_FIGURE].next_paragraph_style = document.styles[STYLE_FIGURE_CAPTION]
    document.styles[STYLE_TABLE_CAPTION].next_paragraph_style = document.styles[STYLE_TABLE_TEXT]
    document.styles[STYLE_TABLE_CONTINUATION].next_paragraph_style = document.styles[STYLE_TABLE_HEADER]
    document.styles[STYLE_PROPOSAL_BODY].next_paragraph_style = document.styles[STYLE_PROPOSAL_BODY]
    document.styles[STYLE_PROPOSAL_PROMPT].next_paragraph_style = document.styles[STYLE_PROPOSAL_BODY]
    document.styles[STYLE_PROPOSAL_UNNUMBERED_HEADING].next_paragraph_style = document.styles[STYLE_PROPOSAL_BODY]

    marker = document.styles["YibinSectionMarker"]
    marker.font.hidden = True

    settings = document.settings.element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")

    document.core_properties.title = "YibinThesis Pandoc reference document"
    document.core_properties.subject = "宜宾学院本科毕业论文非官方 LaTeX 模板的 Word 样式"
    document.core_properties.author = "YibinThesis contributors"
    document.core_properties.comments = (
        "Semantic reference.docx derived from public school formatting evidence; "
        "not an official University template."
    )

    # Pandoc imports styles and section settings, not the reference body content.
    if document.paragraphs:
        document.paragraphs[0].text = ""
    else:
        document.add_paragraph("")
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    return output


def main() -> int:
    project_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        type=Path,
        default=project_root / "word" / "reference.docx",
        help="Output DOCX path (default: word/reference.docx).",
    )
    args = parser.parse_args()
    output = args.output.expanduser().resolve()
    build_reference_docx(output)
    print(f"Generated: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
