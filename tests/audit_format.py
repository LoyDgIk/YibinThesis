#!/usr/bin/env python3
"""Audit the reusable YibinThesis font and paragraph-format contract."""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
ROOT = Path(__file__).resolve().parents[1]


class Audit:
    def __init__(self) -> None:
        self.errors: list[str] = []

    def equal(self, label: str, actual: object, expected: object) -> None:
        if actual != expected:
            self.errors.append(f"{label}: expected {expected!r}, got {actual!r}")

    def true(self, label: str, condition: bool) -> None:
        if not condition:
            self.errors.append(label)


def wattr(name: str) -> str:
    return W + name


def child(element: ET.Element | None, name: str) -> ET.Element | None:
    return None if element is None else element.find(W + name)


def attr(element: ET.Element | None, name: str) -> str | None:
    return None if element is None else element.get(wattr(name))


def enabled(element: ET.Element | None) -> bool:
    if element is None:
        return False
    return attr(element, "val") not in {"0", "false", "off"}


def load_part(docx: Path, member: str) -> ET.Element:
    with zipfile.ZipFile(docx) as archive:
        return ET.fromstring(archive.read(member))


def style_map(styles_root: ET.Element) -> dict[str, ET.Element]:
    result: dict[str, ET.Element] = {}
    for style in styles_root.findall(W + "style"):
        name = child(style, "name")
        value = attr(name, "val")
        if value:
            result[value.casefold()] = style
    return result


def audit_style(
    audit: Audit,
    styles: dict[str, ET.Element],
    name: str,
    *,
    style_type: str = "paragraph",
    east_asia: str | None = None,
    latin: str | None = None,
    size: str | None = None,
    bold: bool | None = None,
    line: str | None = None,
    line_rule: str | None = None,
    first_chars: str | None | object = ...,  # ``...`` means do not check.
    first_line: str | None | object = ...,
    left: str | None | object = ...,
    hanging: str | None | object = ...,
    alignment: str | None = None,
    page_break: bool | None = None,
) -> None:
    style = styles.get(name.casefold())
    audit.true(f"missing Word style: {name}", style is not None)
    if style is None:
        return
    audit.equal(f"{name}.type", attr(style, "type"), style_type)
    p_pr = child(style, "pPr")
    r_pr = child(style, "rPr")
    fonts = child(r_pr, "rFonts")
    spacing = child(p_pr, "spacing")
    indent = child(p_pr, "ind")
    if east_asia is not None:
        audit.equal(f"{name}.font.eastAsia", attr(fonts, "eastAsia"), east_asia)
    if latin is not None:
        audit.equal(f"{name}.font.ascii", attr(fonts, "ascii"), latin)
        audit.equal(f"{name}.font.hAnsi", attr(fonts, "hAnsi"), latin)
    if size is not None:
        audit.equal(f"{name}.size", attr(child(r_pr, "sz"), "val"), size)
    if bold is not None:
        audit.equal(f"{name}.bold", enabled(child(r_pr, "b")), bold)
    if line is not None:
        audit.equal(f"{name}.line", attr(spacing, "line"), line)
    if line_rule is not None:
        audit.equal(f"{name}.lineRule", attr(spacing, "lineRule"), line_rule)
    if first_chars is not ...:
        audit.equal(f"{name}.firstLineChars", attr(indent, "firstLineChars"), first_chars)
    if first_line is not ...:
        audit.equal(f"{name}.firstLine", attr(indent, "firstLine"), first_line)
    if left is not ...:
        audit.equal(f"{name}.left", attr(indent, "left"), left)
    if hanging is not ...:
        audit.equal(f"{name}.hanging", attr(indent, "hanging"), hanging)
    if alignment is not None:
        audit.equal(f"{name}.alignment", attr(child(p_pr, "jc"), "val"), alignment)
    if page_break is not None:
        audit.equal(
            f"{name}.pageBreakBefore",
            enabled(child(p_pr, "pageBreakBefore")),
            page_break,
        )


def audit_geometry(audit: Audit, docx: Path) -> None:
    document = load_part(docx, "word/document.xml")
    sections = document.findall(f".//{W}sectPr")
    audit.true(f"{docx.name}: no section properties", bool(sections))
    for index, section in enumerate(sections, start=1):
        size = child(section, "pgSz")
        margin = child(section, "pgMar")
        prefix = f"{docx.name}.section[{index}]"
        audit.equal(prefix + ".page.width", attr(size, "w"), "11906")
        audit.equal(prefix + ".page.height", attr(size, "h"), "16838")
        for key, expected in {
            "top": "1417",
            "right": "1417",
            "bottom": "1417",
            "left": "1701",
            "header": "850",
            "footer": "850",
            "gutter": "0",
        }.items():
            audit.equal(prefix + f".margin.{key}", attr(margin, key), expected)


def audit_docx(audit: Audit, docx: Path, profile: str | None = None) -> None:
    audit.true(f"DOCX not found: {docx}", docx.is_file())
    if not docx.is_file():
        return
    try:
        styles = style_map(load_part(docx, "word/styles.xml"))
        audit_geometry(audit, docx)
    except (KeyError, OSError, ET.ParseError, zipfile.BadZipFile) as exc:
        audit.errors.append(f"cannot inspect {docx}: {exc}")
        return

    common = (
        ("Normal", "SimSun", "Times New Roman", "24", False, "360", "200"),
        ("Body Text", "SimSun", "Times New Roman", "24", False, "360", "200"),
        ("First Paragraph", "SimSun", "Times New Roman", "24", False, "360", "200"),
        ("Heading 2", "KaiTi", "Times New Roman", "30", True, "360", "200"),
        ("Heading 3", "SimSun", "Times New Roman", "28", True, "360", "200"),
        ("Heading 4", "SimSun", "Times New Roman", "28", True, "360", "200"),
        ("ChineseAbstract", "SimSun", "Times New Roman", "24", False, "360", "200"),
    )
    for name, east_asia, latin, size, bold, line, first_chars in common:
        audit_style(
            audit,
            styles,
            name,
            east_asia=east_asia,
            latin=latin,
            size=size,
            bold=bold,
            line=line,
            line_rule="auto",
            first_chars=first_chars,
        )

    for name, east_asia, size in (
        ("Yibin Heading 1", "SimHei", "32"),
        ("Yibin Heading 2", "KaiTi", "30"),
        ("Yibin Heading 3", "SimSun", "28"),
        ("Yibin Heading 4", "SimSun", "28"),
    ):
        audit_style(
            audit,
            styles,
            name,
            east_asia=east_asia,
            latin="Times New Roman",
            size=size,
            bold=True,
            line="360",
            line_rule="auto",
            first_chars="200",
        )

    audit_style(
        audit,
        styles,
        "Heading 1",
        east_asia="SimHei",
        latin="Times New Roman",
        size="32",
        bold=True,
        line="360",
        line_rule="auto",
        first_chars="200" if profile != "science" else None,
        first_line="0" if profile == "science" else "480",
        alignment="center" if profile == "science" else ("left" if profile == "humanities" else None),
        page_break=True if profile == "science" else False,
    )
    audit_style(
        audit,
        styles,
        "EnglishAbstract",
        east_asia="Times New Roman",
        latin="Times New Roman",
        size="24",
        bold=False,
        line="360",
        line_rule="auto",
        first_chars="0",
        first_line="0",
        alignment="both",
    )
    audit_style(
        audit,
        styles,
        "FrontTitle",
        east_asia="SimHei",
        latin="Times New Roman",
        size="32",
        bold=True,
        line="360",
        alignment="center",
    )
    audit_style(
        audit,
        styles,
        "Unnumbered Heading 1",
        east_asia="SimHei",
        latin="Times New Roman",
        size="32",
        bold=True,
        line="360",
        alignment="center",
        page_break=True,
    )
    audit_style(
        audit,
        styles,
        "Keywords",
        east_asia="SimSun",
        latin="Times New Roman",
        size="24",
        bold=False,
        line="360",
        first_chars="0",
        first_line="0",
        alignment="left",
    )
    audit_style(
        audit,
        styles,
        "KeywordLabel",
        style_type="character",
        east_asia="SimHei",
        latin="Times New Roman",
        size="24",
        bold=True,
    )
    audit_style(
        audit,
        styles,
        "TOC Heading",
        east_asia="SimHei",
        latin="Times New Roman",
        size="32",
        bold=True,
        line="360",
        alignment="center",
    )
    audit_style(
        audit,
        styles,
        "Yibin TOC Heading",
        east_asia="SimHei",
        latin="Times New Roman",
        size="32",
        bold=True,
        line="360",
        first_chars="0",
        first_line="0",
        alignment="center",
    )
    for level, left in ((1, "0"), (2, "480"), (3, "960")):
        audit_style(
            audit,
            styles,
            f"TOC {level}",
            east_asia="SimSun",
            latin="Times New Roman",
            size="24",
            bold=False,
            line="360",
            first_chars="0",
            first_line="0",
            left=left,
        )
    audit_style(
        audit,
        styles,
        "Bibliography",
        east_asia="SimSun",
        latin="Times New Roman",
        size="21",
        bold=False,
        line="360",
        left="420",
        hanging="420",
    )
    audit_style(
        audit,
        styles,
        "Notes",
        east_asia="SimSun",
        latin="Times New Roman",
        size="18",
        bold=False,
        line="360",
        first_line="0",
    )
    for name, size, line in (("Header", "18", "240"), ("Footer", "18", "240"), ("Caption", "21", "240")):
        audit_style(
            audit,
            styles,
            name,
            east_asia="SimSun",
            latin="Times New Roman",
            size=size,
            line=line,
            alignment="center",
        )


def audit_latex(audit: Audit, class_file: Path) -> None:
    audit.true(f"class file not found: {class_file}", class_file.is_file())
    if not class_file.is_file():
        return
    text = class_file.read_text(encoding="utf-8")
    for label, token in {
        "body size is small-four": "zihao=-4",
        "body line spacing is 1.5": r"\setstretch{1.5}",
        "body first-line indent is two em": r"\setlength{\parindent}{2em}",
        "SimSun is the CJK main font": r"\setCJKmainfont{SimSun}",
        "SimSun backs the Song family": r"\setCJKfamilyfont{yibin-song}{SimSun}",
        "SimHei is the CJK sans font": r"\setCJKsansfont{SimHei}",
        "SimHei backs the Hei family": r"\setCJKfamilyfont{yibin-hei}{SimHei}",
        "KaiTi backs the Kai family": r"\setCJKfamilyfont{yibin-kai}{KaiTi}",
        "Times New Roman is the Latin main font": r"\setmainfont{Times New Roman}",
        "Times New Roman backs the English family": r"\newfontfamily\yibinenglishfont{Times New Roman}",
        "level-one heading is SimHei three-size": r"\yibinhei\bfseries\zihao{3}",
        "level-two heading is KaiTi small-three": r"\yibinkai\bfseries\zihao{-3}",
        "level-three heading is SimSun four-size": r"\yibinsong\bfseries\zihao{4}",
        "header uses small-five": r"\yibinsong\zihao{-5}\yibintitle",
        "TOC title keeps symmetric fill": r"\renewcommand{\cftaftertoctitle}{\hfill\mbox{}}",
    }.items():
        audit.true(f"LaTeX contract missing: {label}", token in text)
    for label, token in {
        "science chapter spacing preserves following indent": r"\titlespacing{\chapter}{0pt}{0pt}{24pt}",
        "humanities chapter spacing preserves following indent": r"\titlespacing{\chapter}{2em}{18pt}{12pt}",
        "numberless chapter spacing preserves following indent": r"\titlespacing{name=\chapter,numberless}{0pt}{0pt}{24pt}",
        "section spacing preserves following indent": r"\titlespacing{\section}{2em}{18pt}{12pt}",
        "subsection spacing preserves following indent": r"\titlespacing{\subsection}{2em}{15pt}{9pt}",
        "subsubsection spacing preserves following indent": r"\titlespacing{\subsubsection}{2em}{12pt}{6pt}",
    }.items():
        audit.true(f"LaTeX heading spacing missing: {label}", token in text)
    starred_heading_spacing = re.search(
        r"\\titlespacing\*\s*\{(?:name=\\chapter,numberless|\\(?:chapter|section|subsection|subsubsection))\}",
        text,
    )
    audit.true(
        "LaTeX headings must not suppress the following paragraph indent with \\titlespacing*",
        starred_heading_spacing is None,
    )
    english_abstract = re.search(
        r"\\NewDocumentEnvironment\{enabstract\}.*?"
        r"\\setlength\{\\parindent\}\{0pt\}",
        text,
        flags=re.DOTALL,
    )
    audit.true("LaTeX English abstract must be flush left (no first-line indent)", english_abstract is not None)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--class-file", type=Path, default=ROOT / "yibinthesis.cls")
    parser.add_argument("--reference", type=Path, default=ROOT / "word" / "reference.docx")
    args = parser.parse_args()

    audit = Audit()
    audit_latex(audit, args.class_file)
    audit_docx(audit, args.reference)

    if audit.errors:
        print("FORMAT AUDIT: FAIL", file=sys.stderr)
        for error in audit.errors:
            print(f"- {error}", file=sys.stderr)
        return 1
    print("FORMAT AUDIT: PASS")
    print("- SimSun/SimHei/KaiTi/Times New Roman style mapping")
    print("- 12 pt body, 1.5-line spacing, two-character Chinese first-line indent after headings")
    print("- flush-left English abstract, heading ladder, captions, notes, references")
    print("- A4 geometry and school margins/header/footer distances")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
