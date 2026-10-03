#!/usr/bin/env python3
"""Audit the reusable YibinThesis font and paragraph-format contract."""

from __future__ import annotations

import argparse
import hashlib
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R = f"{{{R_NS}}}"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WP = f"{{{WP_NS}}}"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
A = f"{{{A_NS}}}"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PR = f"{{{PR_NS}}}"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
CP = f"{{{CP_NS}}}"
ROOT = Path(__file__).resolve().parents[1]
OFFICIAL_LOGO_SHA256 = "5E3A203868F8EE661D7185D35404C555399EEC14CF7C467DAAC12DA54F516AFA"


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


def enabled_value(value: str | None) -> bool:
    return value is not None and value.casefold() not in {"0", "false", "off"}


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


def styles_by_id(styles_root: ET.Element) -> dict[str, ET.Element]:
    return {
        style_id: style
        for style in styles_root.findall(W + "style")
        if (style_id := attr(style, "styleId"))
    }


def style_name(style: ET.Element | None) -> str:
    return attr(child(style, "name"), "val") or ""


def paragraph_text(paragraph: ET.Element) -> str:
    pieces: list[str] = []
    for node in paragraph.iter():
        if node.tag == W + "t":
            pieces.append(node.text or "")
        elif node.tag == W + "tab":
            pieces.append("\t")
    return "".join(pieces)


def paragraph_style(
    paragraph: ET.Element,
    by_id: dict[str, ET.Element],
) -> tuple[str, ET.Element | None]:
    style_id = attr(child(child(paragraph, "pPr"), "pStyle"), "val")
    style = by_id.get(style_id or "")
    return style_name(style), style


def style_chain(
    style: ET.Element | None,
    by_id: dict[str, ET.Element],
) -> list[ET.Element]:
    chain: list[ET.Element] = []
    visited: set[str] = set()
    current = style
    while current is not None:
        style_id = attr(current, "styleId") or str(id(current))
        if style_id in visited:
            break
        visited.add(style_id)
        chain.append(current)
        current = by_id.get(attr(child(current, "basedOn"), "val") or "")
    return chain


def effective_p_element(
    paragraph: ET.Element,
    style: ET.Element | None,
    by_id: dict[str, ET.Element],
    name: str,
) -> ET.Element | None:
    direct = child(child(paragraph, "pPr"), name)
    if direct is not None:
        return direct
    for candidate in style_chain(style, by_id):
        value = child(child(candidate, "pPr"), name)
        if value is not None:
            return value
    return None


def effective_run_element(
    run: ET.Element,
    paragraph_style_element: ET.Element | None,
    by_id: dict[str, ET.Element],
    name: str,
) -> ET.Element | None:
    run_pr = child(run, "rPr")
    direct = child(run_pr, name)
    if direct is not None:
        return direct
    character_style = by_id.get(attr(child(run_pr, "rStyle"), "val") or "")
    for candidate in style_chain(character_style, by_id):
        value = child(child(candidate, "rPr"), name)
        if value is not None:
            return value
    for candidate in style_chain(paragraph_style_element, by_id):
        value = child(child(candidate, "rPr"), name)
        if value is not None:
            return value
    return None


def effective_run_style_name(run: ET.Element, by_id: dict[str, ET.Element]) -> str:
    style_id = attr(child(child(run, "rPr"), "rStyle"), "val") or ""
    return style_name(by_id.get(style_id))


def effective_bold(
    run: ET.Element,
    paragraph_style_element: ET.Element | None,
    by_id: dict[str, ET.Element],
) -> bool:
    return enabled(effective_run_element(run, paragraph_style_element, by_id, "b"))


def effective_size(
    run: ET.Element,
    paragraph_style_element: ET.Element | None,
    by_id: dict[str, ET.Element],
) -> str | None:
    return attr(effective_run_element(run, paragraph_style_element, by_id, "sz"), "val")


def effective_font(
    run: ET.Element,
    paragraph_style_element: ET.Element | None,
    by_id: dict[str, ET.Element],
    field: str,
) -> str | None:
    return attr(
        effective_run_element(run, paragraph_style_element, by_id, "rFonts"),
        field,
    )


def normalized_field_text(text: str) -> str:
    return re.sub(r"\s+", "", text)


def paragraph_field_instructions(paragraph: ET.Element) -> list[str]:
    """Return normalized field instructions contained in one paragraph."""
    return field_instructions_from_nodes([paragraph])


def field_instructions_from_nodes(nodes: list[ET.Element]) -> list[str]:
    """Parse both fldSimple and Word-normalized nested complex fields."""
    return [instruction for instruction, _locked in field_records_from_nodes(nodes)]


def field_records_from_nodes(nodes: list[ET.Element]) -> list[tuple[str, bool]]:
    records: list[tuple[str, bool]] = []
    stack: list[dict[str, object]] = []

    def visit(element: ET.Element) -> None:
        if element.tag == W + "fldSimple":
            value = element.get(W + "instr", "")
            if value.strip():
                records.append((value, enabled_value(element.get(W + "fldLock"))))
        elif element.tag == W + "fldChar":
            field_type = element.get(W + "fldCharType", "")
            if field_type == "begin":
                stack.append(
                    {
                        "parts": [],
                        "locked": enabled_value(element.get(W + "fldLock")),
                    }
                )
            elif field_type == "end" and stack:
                current = stack.pop()
                value = "".join(current["parts"])
                if value.strip():
                    records.append((value, bool(current["locked"])))
        elif element.tag == W + "instrText" and stack:
            stack[-1]["parts"].append(element.text or "")
        for child_node in element:
            visit(child_node)

    for node in nodes:
        visit(node)
    return [
        (re.sub(r"\s+", " ", value).strip(), locked)
        for value, locked in records
        if value.strip()
    ]


def detailed_field_records(
    nodes: list[ET.Element],
) -> list[tuple[str, bool, list[ET.Element]]]:
    """Return instruction, lock state and cached-result runs for each field."""

    records: list[tuple[str, bool, list[ET.Element]]] = []
    stack: list[dict[str, object]] = []

    def visit(element: ET.Element, current_run: ET.Element | None = None) -> None:
        if element.tag == W + "r":
            current_run = element
        if element.tag == W + "fldSimple":
            value = element.get(W + "instr", "")
            result_runs = element.findall(f".//{W}r")
            if value.strip():
                records.append(
                    (
                        re.sub(r"\s+", " ", value).strip(),
                        enabled_value(element.get(W + "fldLock")),
                        result_runs,
                    )
                )
            return
        if element.tag == W + "fldChar":
            field_type = element.get(W + "fldCharType", "")
            if field_type == "begin":
                stack.append(
                    {
                        "parts": [],
                        "locked": enabled_value(element.get(W + "fldLock")),
                        "separated": False,
                        "runs": [],
                    }
                )
            elif field_type == "separate" and stack:
                stack[-1]["separated"] = True
            elif field_type == "end" and stack:
                current = stack.pop()
                value = "".join(current["parts"])
                if value.strip():
                    records.append(
                        (
                            re.sub(r"\s+", " ", value).strip(),
                            bool(current["locked"]),
                            list(current["runs"]),
                        )
                    )
        elif element.tag == W + "instrText" and stack:
            stack[-1]["parts"].append(element.text or "")
        elif element.tag == W + "t" and stack and bool(stack[-1]["separated"]):
            runs = stack[-1]["runs"]
            if current_run is not None and current_run not in runs:
                runs.append(current_run)
        for child_node in element:
            visit(child_node, current_run)

    for node in nodes:
        visit(node)
    return records


def active_underline(element: ET.Element | None) -> bool:
    return (
        element is not None
        and (attr(element, "val") or "single")
        not in {"0", "none", "false", "off", "nil"}
    )


def active_bottom_borders(cell: ET.Element) -> list[ET.Element]:
    borders = cell.findall(f"{W}tcPr/{W}tcBorders/{W}bottom")
    return [
        border
        for border in borders
        if (attr(border, "val") or "single")
        not in {"0", "none", "false", "off", "nil"}
    ]


def effective_paragraph_bottom_border(
    paragraph: ET.Element,
    style: ET.Element | None,
    by_id: dict[str, ET.Element],
) -> ET.Element | None:
    direct = paragraph.find(f"{W}pPr/{W}pBdr/{W}bottom")
    if direct is not None:
        return direct
    for candidate in style_chain(style, by_id):
        border = candidate.find(f"{W}pPr/{W}pBdr/{W}bottom")
        if border is not None:
            return border
    return None


def twips_value(element: ET.Element | None, name: str) -> int:
    value = attr(element, name)
    return int(value) if value not in (None, "") else 0


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
    underline: bool | None = None,
    line: str | None = None,
    line_rule: str | None = None,
    first_chars: str | None | object = ...,  # ``...`` means do not check.
    first_line: str | None | object = ...,
    left: str | None | object = ...,
    left_chars: str | None | object = ...,
    hanging: str | None | object = ...,
    alignment: str | None = None,
    before: str | None | object = ...,
    after: str | None | object = ...,
    tab_pos: str | None = None,
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
    if underline is not None:
        underline_element = child(r_pr, "u")
        underline_enabled = (
            underline_element is not None
            and (attr(underline_element, "val") or "single")
            not in {"0", "none", "false", "off"}
        )
        audit.equal(f"{name}.underline", underline_enabled, underline)
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
    if left_chars is not ...:
        audit.equal(f"{name}.leftChars", attr(indent, "leftChars"), left_chars)
    if hanging is not ...:
        audit.equal(f"{name}.hanging", attr(indent, "hanging"), hanging)
    if alignment is not None:
        audit.equal(f"{name}.alignment", attr(child(p_pr, "jc"), "val"), alignment)
    if before is not ...:
        audit.equal(f"{name}.before", attr(spacing, "before"), before)
    if after is not ...:
        audit.equal(f"{name}.after", attr(spacing, "after"), after)
    if tab_pos is not None:
        tabs = [] if p_pr is None else p_pr.findall(f"{W}tabs/{W}tab")
        audit.true(
            f"{name}.tab must be right/dot at {tab_pos}",
            any(
                attr(tab, "val") == "right"
                and attr(tab, "leader") == "dot"
                and attr(tab, "pos") in {tab_pos, str(int(tab_pos) + 1)}
                for tab in tabs
            ),
        )
    if page_break is not None:
        audit.equal(
            f"{name}.pageBreakBefore",
            enabled(child(p_pr, "pageBreakBefore")),
            page_break,
        )


def audit_table_continuation_style(
    audit: Audit,
    styles: dict[str, ET.Element],
    prefix: str,
) -> None:
    """Audit the named style used by Word-generated continuation captions."""
    name = "宜宾论文-续表题"
    style = styles.get(name.casefold())
    audit.true(f"{prefix}.{name} style is missing", style is not None)
    if style is None:
        return

    audit.equal(f"{prefix}.{name}.type", attr(style, "type"), "paragraph")
    by_id = {
        attr(candidate, "styleId") or "": candidate
        for candidate in styles.values()
    }
    chain = style_chain(style, by_id)

    def effective(container: str, property_name: str) -> ET.Element | None:
        for candidate in chain:
            value = child(child(candidate, container), property_name)
            if value is not None:
                return value
        return None

    def effective_attr(
        container: str,
        property_name: str,
        attribute_name: str,
    ) -> str | None:
        for candidate in chain:
            value = child(child(candidate, container), property_name)
            inherited = attr(value, attribute_name)
            if inherited is not None:
                return inherited
        return None

    p_pr = child(style, "pPr")
    fonts = effective("rPr", "rFonts")
    size = effective("rPr", "sz")
    spacing = effective("pPr", "spacing")
    audit.true(
        f"{prefix}.{name}.font.eastAsia must be SimSun",
        attr(fonts, "eastAsia") in {"SimSun", "宋体"},
    )
    audit.equal(
        f"{prefix}.{name}.font.ascii",
        attr(fonts, "ascii"),
        "Times New Roman",
    )
    audit.equal(
        f"{prefix}.{name}.font.hAnsi",
        attr(fonts, "hAnsi"),
        "Times New Roman",
    )
    audit.equal(f"{prefix}.{name}.size", attr(size, "val"), "21")
    audit.equal(
        f"{prefix}.{name}.alignment",
        attr(child(p_pr, "jc"), "val"),
        "right",
    )
    audit.equal(
        f"{prefix}.{name}.line",
        effective_attr("pPr", "spacing", "line"),
        "240",
    )
    audit.equal(
        f"{prefix}.{name}.lineRule",
        effective_attr("pPr", "spacing", "lineRule"),
        "auto",
    )
    # Word removes an explicit zero before-spacing when it saves the file;
    # an absent value is therefore the OOXML-equivalent of zero here.
    audit.equal(f"{prefix}.{name}.before", twips_value(spacing, "before"), 0)
    audit.equal(f"{prefix}.{name}.after", twips_value(spacing, "after"), 0)
    audit.true(
        f"{prefix}.{name}.pageBreakBefore must be enabled",
        enabled(child(p_pr, "pageBreakBefore")),
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


def audit_logo_asset(audit: Audit, logo: Path) -> None:
    audit.true(f"official logo asset not found: {logo}", logo.is_file())
    if not logo.is_file():
        return
    digest = hashlib.sha256(logo.read_bytes()).hexdigest().upper()
    audit.equal("official logo SHA-256", digest, OFFICIAL_LOGO_SHA256)


def audit_generated_docx(audit: Audit, docx: Path, logo: Path) -> None:
    audit.true(f"generated DOCX not found: {docx}", docx.is_file())
    if not docx.is_file():
        return
    try:
        with zipfile.ZipFile(docx) as archive:
            document = ET.fromstring(archive.read("word/document.xml"))
            styles_root = ET.fromstring(archive.read("word/styles.xml"))
            numbering_root = ET.fromstring(archive.read("word/numbering.xml"))
            relationships = ET.fromstring(
                archive.read("word/_rels/document.xml.rels")
            )
            members = set(archive.namelist())
            media_bytes = {
                member: archive.read(member)
                for member in members
                if member.startswith("word/media/")
            }
            custom_xml = {
                member: archive.read(member)
                for member in members
                if member.startswith("customXml/") and member.endswith(".xml")
            }
            core_properties = ET.fromstring(archive.read("docProps/core.xml"))
    except (KeyError, OSError, ET.ParseError, zipfile.BadZipFile) as exc:
        audit.errors.append(f"cannot inspect generated {docx}: {exc}")
        return

    prefix = docx.name
    keywords_node = core_properties.find(CP + "keywords")
    keywords = "" if keywords_node is None else (keywords_node.text or "")
    type_match = re.search(
        r"(?:^|;)\s*document-type\s*=\s*(thesis|proposal|literature-review)",
        keywords,
        re.IGNORECASE,
    )
    document_type = type_match.group(1).casefold() if type_match else "thesis"
    is_thesis = document_type == "thesis"
    by_id = styles_by_id(styles_root)
    by_name = style_map(styles_root)
    heading_num_ids: list[str] = []
    heading_style_names = (
        "宜宾论文-一级标题",
        "宜宾论文-二级标题",
        "宜宾论文-三级标题",
        "宜宾论文-四级标题",
    )
    for level, style_key in enumerate(heading_style_names):
        style = by_name.get(style_key.casefold())
        audit.true(f"{prefix}.{style_key} style is missing", style is not None)
        if style is None:
            continue
        num_pr = child(child(style, "pPr"), "numPr")
        audit.true(f"{prefix}.{style_key} native numPr is missing", num_pr is not None)
        if num_pr is not None:
            actual_level = attr(child(num_pr, "ilvl"), "val")
            if level == 0:
                audit.true(
                    f"{prefix}.{style_key}.ilvl",
                    actual_level in {None, "0"},
                )
            else:
                audit.equal(
                    f"{prefix}.{style_key}.ilvl",
                    actual_level,
                    str(level),
                )
            heading_num_ids.append(attr(child(num_pr, "numId"), "val") or "")
    audit.equal(
        f"{prefix}.Heading native numId count",
        len({value for value in heading_num_ids if value}),
        1,
    )
    if heading_num_ids and heading_num_ids[0]:
        num = next(
            (
                node
                for node in numbering_root.findall(W + "num")
                if attr(node, "numId") == heading_num_ids[0]
            ),
            None,
        )
        audit.true(f"{prefix}.Heading numbering instance is missing", num is not None)
        if num is not None:
            abstract_id = attr(child(num, "abstractNumId"), "val")
            abstract = next(
                (
                    node
                    for node in numbering_root.findall(W + "abstractNum")
                    if attr(node, "abstractNumId") == abstract_id
                ),
                None,
            )
            audit.true(f"{prefix}.Heading abstract numbering is missing", abstract is not None)
            if abstract is not None:
                levels = abstract.findall(W + "lvl")
                formats = [
                    attr(child(level, "numFmt"), "val") for level in levels[:4]
                ]
                texts = [
                    attr(child(level, "lvlText"), "val") for level in levels[:4]
                ]
                audit.true(
                    f"{prefix}.Heading numbering profile is not native humanities/science",
                    (tuple(formats), tuple(texts))
                    in {
                        (
                            ("chineseCounting", "chineseCounting", "decimal", "decimal"),
                            ("%1、", "（%2）", "%3.", "（%4）"),
                        ),
                        (
                            ("decimal", "decimal", "decimal", "decimal"),
                            ("%1", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4"),
                        ),
                    },
                )
    audit_table_continuation_style(audit, style_map(styles_root), prefix)
    paragraphs = document.findall(f".//{W}p")
    parents = {child_node: parent for parent in document.iter() for child_node in parent}
    field_code = "".join(
        node.text or "" for node in document.findall(f".//{W}instrText")
    ) + " ".join(
        node.get(W + "instr", "") for node in document.findall(f".//{W}fldSimple")
    )
    audit.true(
        f"{prefix}.TOC field is missing",
        not is_thesis or ("TOC" in field_code and '1-3' in field_code),
    )
    named: dict[str, list[ET.Element]] = {}
    for paragraph in paragraphs:
        name, _ = paragraph_style(paragraph, by_id)
        named.setdefault(name.casefold(), []).append(paragraph)

    def only(style: str) -> ET.Element | None:
        matches = named.get(style.casefold(), [])
        if not is_thesis:
            return None
        audit.equal(f"{prefix}.{style}.count", len(matches), 1)
        return matches[0] if len(matches) == 1 else None

    cover_logo = only("CoverLogo")
    cover_type = only("CoverThesisType")
    cover_titles = named.get("covertitle", []) + named.get("covertitlewithversion", [])
    audit.equal(
        f"{prefix}.CoverTitle.variant.count",
        len(cover_titles),
        1 if is_thesis else 0,
    )
    cover_title = cover_titles[0] if len(cover_titles) == 1 else None
    cover_fields = named.get("coverfield", [])
    audit.equal(f"{prefix}.CoverField.count", len(cover_fields), 6 if is_thesis else 0)
    audit.true(
        f"{prefix}.CoverVersion.count must be optional and unique",
        len(named.get("coverversion", [])) <= (1 if is_thesis else 0),
    )
    audit.true(
        f"{prefix}.CoverDate.count must be optional and unique",
        len(named.get("coverdate", [])) <= (1 if is_thesis else 0),
    )

    if cover_logo is not None:
        extents = cover_logo.findall(f".//{WP}extent")
        audit.equal(f"{prefix}.cover.logo.extent.count", len(extents), 1)
        if len(extents) == 1:
            audit.true(
                f"{prefix}.cover.logo.width must be 13.15 cm",
                abs(int(extents[0].get("cx", "0")) - 4_734_000) <= 10_000,
            )
            audit.true(
                f"{prefix}.cover.logo.height must be 3.65 cm",
                abs(int(extents[0].get("cy", "0")) - 1_314_000) <= 10_000,
            )
        blips = cover_logo.findall(f".//{A}blip")
        audit.equal(f"{prefix}.cover.logo.blip.count", len(blips), 1)
    else:
        blips = []

    rel_map = {
        relationship.get("Id", ""): relationship
        for relationship in relationships.findall(PR + "Relationship")
    }
    for index, blip in enumerate(document.findall(f".//{A}blip"), start=1):
        embed = blip.get(R + "embed")
        link = blip.get(R + "link")
        audit.true(f"{prefix}.image[{index}] must use r:embed", bool(embed))
        audit.true(f"{prefix}.image[{index}] must not use r:link", not link)
        relationship = rel_map.get(embed or "")
        audit.true(
            f"{prefix}.image[{index}] relationship is missing",
            relationship is not None,
        )
        if relationship is None:
            continue
        audit.true(
            f"{prefix}.image[{index}] must not be external",
            relationship.get("TargetMode") != "External",
        )
        target = relationship.get("Target", "")
        member = posixpath.normpath(posixpath.join("word", target))
        audit.true(
            f"{prefix}.image[{index}] media part is missing: {member}",
            member in media_bytes,
        )

    if blips and logo.is_file():
        logo_embed = blips[0].get(R + "embed")
        relationship = rel_map.get(logo_embed or "")
        if relationship is not None:
            member = posixpath.normpath(
                posixpath.join("word", relationship.get("Target", ""))
            )
            if member in media_bytes:
                digest = hashlib.sha256(media_bytes[member]).hexdigest().upper()
                audit.equal(
                    f"{prefix}.cover.logo.media.SHA-256",
                    digest,
                    OFFICIAL_LOGO_SHA256,
                )

    if cover_type is not None:
        audit.equal(
            f"{prefix}.cover.type.text",
            paragraph_text(cover_type),
            "本科生毕业论文（设计）",
        )
    if cover_title is not None:
        _, paragraph_style_element = paragraph_style(cover_title, by_id)
        runs = [run for run in cover_title.findall(f".//{W}r") if paragraph_text(run)]
        audit.true(f"{prefix}.cover.title has no text run", bool(runs))
        if runs:
            audit.equal(
                f"{prefix}.cover.title.size",
                effective_size(runs[0], paragraph_style_element, by_id),
                "36",
            )

    expected_field_tokens = (
        ("学院（部）",),
        ("专业",),
        ("学生姓名",),
        ("学号",),
        ("指导教师（校内）",),
        ("指导教师（校外）",),
    )
    for index, (paragraph, tokens) in enumerate(
        zip(cover_fields, expected_field_tokens, strict=False),
        start=1,
    ):
        text = normalized_field_text(paragraph_text(paragraph))
        for token in tokens:
            audit.true(
                f"{prefix}.cover.field[{index}] missing {token}",
                normalized_field_text(token) in text,
            )

    cover_layout_tables: list[ET.Element] = []
    seen_cover_tables: set[int] = set()
    for paragraph in cover_fields:
        table = parents.get(paragraph)
        while table is not None and table.tag != W + "tbl":
            table = parents.get(table)
        if table is not None and id(table) not in seen_cover_tables:
            seen_cover_tables.add(id(table))
            cover_layout_tables.append(table)
    expected_cover_geometry = (
        ([1980, 6583], "926"),
        ([1980, 6763], "925"),
        ([1800, 6763], "952"),
        ([1800, 3214, 900, 2572], "925"),
        ([2880, 2906, 874, 1774], "926"),
        ([2880, 2906, 874, 1774], "926"),
    ) if is_thesis else ()
    audit.equal(
        f"{prefix}.cover.geometry-table.count",
        len(cover_layout_tables),
        len(expected_cover_geometry),
    )
    for index, (table, (expected_grid, expected_height)) in enumerate(
        zip(cover_layout_tables, expected_cover_geometry, strict=False),
        start=1,
    ):
        actual_grid = [
            int(attr(column, "w") or "0")
            for column in table.findall(f"{W}tblGrid/{W}gridCol")
        ]
        audit.equal(
            f"{prefix}.cover.geometry-table[{index}].grid",
            actual_grid,
            expected_grid,
        )
        row_height = table.find(f"{W}tr/{W}trPr/{W}trHeight")
        audit.equal(
            f"{prefix}.cover.geometry-table[{index}].height",
            attr(row_height, "val"),
            expected_height,
        )
        audit.equal(
            f"{prefix}.cover.geometry-table[{index}].height-rule",
            attr(row_height, "hRule"),
            "exact",
        )
        alignment = attr(table.find(f"{W}tblPr/{W}jc"), "val")
        audit.true(
            f"{prefix}.cover.geometry-table[{index}] must be left aligned",
            alignment in {None, "left", "start"},
        )

    # There are nine independently fillable slots on the six official cover
    # rows (college, major, author, student id, grade, two advisers and their
    # titles).  Each slot must be a centred paragraph in one cell.  Its named
    # paragraph style supplies the only bottom rule, keeping the line at the
    # official text underline height rather than at the bottom of a tall row.
    cover_value_runs = [
        run
        for run in document.findall(f".//{W}r")
        if effective_run_style_name(run, by_id).casefold() == "covervalue"
    ]
    value_cells: set[int] = set()
    value_tables: set[int] = set()
    for slot_index, run in enumerate(cover_value_runs, start=1):
        paragraph = parents.get(run)
        while paragraph is not None and paragraph.tag != W + "p":
            paragraph = parents.get(paragraph)
        cell = paragraph
        while cell is not None and cell.tag != W + "tc":
            cell = parents.get(cell)
        audit.true(f"{prefix}.cover.value-slot[{slot_index}] must be in a table cell", cell is not None)
        if paragraph is not None:
            value_style_name, value_style = paragraph_style(paragraph, by_id)
            audit.equal(
                f"{prefix}.cover.value-slot[{slot_index}].paragraph-style",
                value_style_name,
                "CoverValueLine",
            )
            underline = effective_run_element(run, value_style, by_id, "u")
            audit.true(
                f"{prefix}.cover.value-slot[{slot_index}] must not use character underline",
                not active_underline(underline),
            )
            audit.true(
                f"{prefix}.cover.value-slot[{slot_index}] must not contain NBSP padding",
                "\u00a0" not in paragraph_text(run),
            )
            alignment = effective_p_element(paragraph, value_style, by_id, "jc")
            audit.equal(
                f"{prefix}.cover.value-slot[{slot_index}].alignment",
                attr(alignment, "val"),
                "center",
            )
            underscore_leaders = [
                tab
                for tab in paragraph.findall(f"{W}pPr/{W}tabs/{W}tab")
                if attr(tab, "leader") in {"underscore", "heavy"}
            ]
            audit.equal(
                f"{prefix}.cover.value-slot[{slot_index}].underscore-leader.count",
                len(underscore_leaders),
                0,
            )
            paragraph_border = effective_paragraph_bottom_border(
                paragraph,
                value_style,
                by_id,
            )
            audit.true(
                f"{prefix}.cover.value-slot[{slot_index}] must have one paragraph bottom border",
                paragraph_border is not None
                and (attr(paragraph_border, "val") or "single")
                not in {"0", "none", "false", "off", "nil"},
            )
        if cell is not None:
            value_cells.add(id(cell))
            table = cell
            while table is not None and table.tag != W + "tbl":
                table = parents.get(table)
            if table is not None:
                value_tables.add(id(table))
            audit.equal(
                f"{prefix}.cover.value-slot[{slot_index}].cell-bottom-border.count",
                len(active_bottom_borders(cell)),
                0,
            )
    audit.equal(f"{prefix}.cover.value-cell.unique-count", len(value_cells), 9 if is_thesis else 0)
    audit.equal(f"{prefix}.cover.layout-table.count", len(value_tables), 6 if is_thesis else 0)

    declaration_title = next(
        (
            paragraph
            for paragraph in named.get("declarationtitle", [])
            if paragraph_text(paragraph) == "原创性声明"
        ),
        None,
    )
    audit.true(f"{prefix}.originality title missing", not is_thesis or declaration_title is not None)
    declaration_body = next(
        (
            paragraph
            for paragraph in named.get("declarationbody", [])
            if paragraph_text(paragraph).startswith("本人呈交的学位论文")
        ),
        None,
    )
    audit.true(f"{prefix}.originality body missing", not is_thesis or declaration_body is not None)
    if declaration_body is not None:
        _, paragraph_style_element = paragraph_style(declaration_body, by_id)
        indent = effective_p_element(
            declaration_body, paragraph_style_element, by_id, "ind"
        )
        spacing = effective_p_element(
            declaration_body, paragraph_style_element, by_id, "spacing"
        )
        audit.equal(f"{prefix}.originality.firstLine", attr(indent, "firstLine"), "480")
        audit.equal(f"{prefix}.originality.line", attr(spacing, "line"), "480")
        audit.equal(f"{prefix}.originality.lineRule", attr(spacing, "lineRule"), "auto")

    originality_signatures = [
        paragraph
        for paragraph in named.get("declarationsignature", [])
        if paragraph_text(paragraph).startswith("学位论文作者：")
    ]
    audit.equal(
        f"{prefix}.originality DeclarationSignature.count",
        len(originality_signatures),
        1 if is_thesis else 0,
    )

    regulation_contract = (
        (
            "DeclarationRegulationLead",
            "附：《普通高等学校学生管理规定》（中华人民共和国教育部令第41号）",
        ),
        (
            "DeclarationRegulationClause",
            "第五十二条\u00a0学生有下列情形之一，学校可以给予开除学籍处分：",
        ),
        (
            "DeclarationRegulationItem",
            "（五）学位论文、公开发表的研究成果存在抄袭、篡改、伪造等学术不端行为，情节严重的，或者代写论文、买卖论文的；",
        ),
    )
    regulation_paragraphs: list[ET.Element] = []
    for style, expected in regulation_contract:
        paragraph = only(style)
        if paragraph is not None:
            regulation_paragraphs.append(paragraph)
            audit.equal(f"{prefix}.{style}.text", paragraph_text(paragraph), expected)

    if len(regulation_paragraphs) == 3:
        for index, paragraph in enumerate(regulation_paragraphs):
            _, paragraph_style_element = paragraph_style(paragraph, by_id)
            text_runs = [
                run for run in paragraph.findall(f".//{W}r") if paragraph_text(run)
            ]
            audit.true(f"{prefix}.regulation[{index + 1}] has no runs", bool(text_runs))
            for run_index, run in enumerate(text_runs):
                expected_size = "32" if index == 0 and run_index == 0 else "28"
                audit.equal(
                    f"{prefix}.regulation[{index + 1}].run[{run_index + 1}].size",
                    effective_size(run, paragraph_style_element, by_id),
                    expected_size,
                )
                audit.true(
                    f"{prefix}.regulation[{index + 1}].run[{run_index + 1}] must be bold",
                    effective_bold(run, paragraph_style_element, by_id),
                )
                audit.true(
                    f"{prefix}.regulation[{index + 1}].run[{run_index + 1}] must use SimSun",
                    effective_font(run, paragraph_style_element, by_id, "eastAsia")
                    in {"SimSun", "宋体"},
                )
        clause_style = paragraph_style(regulation_paragraphs[1], by_id)[1]
        clause_indent = effective_p_element(
            regulation_paragraphs[1], clause_style, by_id, "ind"
        )
        clause_spacing = effective_p_element(
            regulation_paragraphs[1], clause_style, by_id, "spacing"
        )
        item_style = paragraph_style(regulation_paragraphs[2], by_id)[1]
        item_indent = effective_p_element(
            regulation_paragraphs[2], item_style, by_id, "ind"
        )
        item_spacing = effective_p_element(
            regulation_paragraphs[2], item_style, by_id, "spacing"
        )
        audit.true(
            f"{prefix}.regulation.clause firstLine must be about 35 pt",
            690 <= twips_value(clause_indent, "firstLine") <= 710,
        )
        default_spacing = styles_root.find(f"{W}docDefaults/{W}pPrDefault/{W}pPr/{W}spacing")
        audit.equal(
            f"{prefix}.regulation.clause.line",
            attr(clause_spacing, "line") or attr(default_spacing, "line"),
            "360",
        )
        audit.true(
            f"{prefix}.regulation.item left must be about 7.15 pt",
            140 <= twips_value(item_indent, "left") <= 146,
        )
        audit.equal(
            f"{prefix}.regulation.item.line",
            attr(item_spacing, "line") or attr(default_spacing, "line"),
            "360",
        )

    toc_headings = named.get("宜宾论文-目录标题".casefold(), [])
    audit.true(
        f"{prefix}.宜宾论文-目录标题.count must be 1 to 3",
        (not is_thesis and not toc_headings)
        or (is_thesis and 1 <= len(toc_headings) <= 3),
    )
    directory_titles = {paragraph_text(paragraph) for paragraph in toc_headings}
    if "图目录" in directory_titles:
        audit.true(
            f"{prefix}.figure directory field is missing",
            '\\c "图"' in field_code,
        )
    if "表目录" in directory_titles:
        audit.true(
            f"{prefix}.table directory field is missing",
            '\\c "表"' in field_code,
        )
    for heading_index, toc_heading in enumerate(toc_headings, start=1):
        _, toc_heading_style = paragraph_style(toc_heading, by_id)
        spacing = effective_p_element(toc_heading, toc_heading_style, by_id, "spacing")
        audit.equal(
            f"{prefix}.toc.heading[{heading_index}].before",
            twips_value(spacing, "before"),
            0,
        )
        audit.equal(
            f"{prefix}.toc.heading[{heading_index}].after",
            twips_value(spacing, "after"),
            0,
        )
        audit.equal(
            f"{prefix}.toc.heading[{heading_index}].line",
            attr(spacing, "line"),
            "240",
        )

    for level, expected_left, expected_chars, expected_line in (
        (1, 0, 0, 360),
        (2, 420, 200, 360),
        (3, 840, 400, 240),
    ):
        # Before Microsoft Word refreshes the TOC field, a valid generated
        # DOCX may contain no materialized TOC paragraphs.  Audit entries when
        # present and rely on the field-code assertion above otherwise.
        toc_entries = named.get(f"toc {level}", [])
        for entry_index, paragraph in enumerate(toc_entries, start=1):
            _, paragraph_style_element = paragraph_style(paragraph, by_id)
            indent = effective_p_element(paragraph, paragraph_style_element, by_id, "ind")
            spacing = effective_p_element(
                paragraph, paragraph_style_element, by_id, "spacing"
            )
            alignment = effective_p_element(
                paragraph, paragraph_style_element, by_id, "jc"
            )
            tabs = effective_p_element(paragraph, paragraph_style_element, by_id, "tabs")
            audit.equal(
                f"{prefix}.TOC {level}[{entry_index}].left",
                twips_value(indent, "left"),
                expected_left,
            )
            audit.equal(
                f"{prefix}.TOC {level}[{entry_index}].leftChars",
                twips_value(indent, "leftChars"),
                expected_chars,
            )
            audit.equal(
                f"{prefix}.TOC {level}[{entry_index}].firstLine",
                twips_value(indent, "firstLine"),
                0,
            )
            audit.equal(
                f"{prefix}.TOC {level}[{entry_index}].alignment",
                attr(alignment, "val"),
                "both",
            )
            audit.equal(
                f"{prefix}.TOC {level}[{entry_index}].line",
                int(attr(spacing, "line") or 0),
                expected_line,
            )
            tab_nodes = [] if tabs is None else tabs.findall(W + "tab")
            audit.true(
                f"{prefix}.TOC {level}[{entry_index}] missing official tab",
                any(
                    attr(tab, "val") == "right"
                    and attr(tab, "leader") == "dot"
                    and attr(tab, "pos") in {"8777", "8778"}
                    for tab in tab_nodes
                ),
            )

    caption_names = {
        "caption",
        "image caption",
        "figure caption",
        "table caption",
        "宜宾论文-图题",
        "宜宾论文-表题",
    }
    captions = [
        paragraph
        for paragraph in paragraphs
        if paragraph_style(paragraph, by_id)[0].casefold() in caption_names
    ]
    # Literature reviews and proposal forms can legitimately contain no figures
    # or tables.  Require at least one caption only for the thesis profile; when
    # a non-thesis document does contain captions, the checks below still audit
    # every caption's style, field, and bookmark structure.
    audit.true(
        f"{prefix}.Caption paragraph is missing",
        bool(captions) or not is_thesis,
    )
    for caption_index, paragraph in enumerate(captions, start=1):
        paragraph_style_name, paragraph_style_element = paragraph_style(paragraph, by_id)
        audit.true(
            f"{prefix}.Caption[{caption_index}].style",
            paragraph_style_name.casefold() in {"宜宾论文-图题", "宜宾论文-表题"},
        )
        based_on = by_id.get(attr(child(paragraph_style_element, "basedOn"), "val") or "")
        audit.equal(
            f"{prefix}.Caption[{caption_index}].basedOn",
            style_name(based_on).casefold(),
            "caption",
        )

    bookmark_starts = document.findall(f".//{W}bookmarkStart")
    bookmark_ends = document.findall(f".//{W}bookmarkEnd")
    bookmark_name_list = [attr(node, "name") or "" for node in bookmark_starts]
    bookmark_names = set(bookmark_name_list)
    folded_bookmark_names = [name.casefold() for name in bookmark_name_list if name]
    audit.equal(
        f"{prefix}.bookmark.name.case-insensitive-unique-count",
        len(set(folded_bookmark_names)),
        len(folded_bookmark_names),
    )
    for name in bookmark_name_list:
        audit.true(f"{prefix}.bookmark name exceeds 40 characters: {name}", len(name) <= 40)
    end_ids = {attr(node, "id") or "" for node in bookmark_ends}
    for node in bookmark_starts:
        name = attr(node, "name") or ""
        bookmark_id = attr(node, "id") or ""
        audit.true(f"{prefix}.bookmark {name!r} has no matching end", bookmark_id in end_ids)

    continuation_paragraphs = named.get("宜宾论文-续表题", [])
    continuation_ref_targets: set[str] = set()
    for continuation_index, paragraph in enumerate(continuation_paragraphs, start=1):
        label = f"{prefix}.TableContinuation[{continuation_index}]"
        visible_text = paragraph_text(paragraph).strip()
        audit.true(
            f"{label}.text must be '续表' followed by a table number",
            re.fullmatch(r"续表\s*\d+(?:\.\d+)?", visible_text) is not None,
        )
        instructions = paragraph_field_instructions(paragraph)
        ref_matches = [
            match
            for instruction in instructions
            if (match := re.match(r"REF\s+([^ \\]+)", instruction, re.IGNORECASE))
        ]
        audit.equal(f"{label}.REF.count", len(ref_matches), 1)
        audit.equal(
            f"{label}.SEQ.count",
            sum(
                1
                for instruction in instructions
                if re.search(r"(?:^| )SEQ(?: |$)", instruction, re.IGNORECASE)
            ),
            0,
        )
        if len(ref_matches) == 1:
            target = ref_matches[0].group(1)
            continuation_ref_targets.add(target)
            audit.true(
                f"{label}.REF target must be a table-caption bookmark: {target}",
                target.startswith("_RefNum") and target in bookmark_names,
            )

        parent = parents.get(paragraph)
        siblings = [] if parent is None else list(parent)
        try:
            paragraph_index = siblings.index(paragraph)
        except ValueError:
            paragraph_index = -1
        next_element = (
            siblings[paragraph_index + 1]
            if paragraph_index >= 0 and paragraph_index + 1 < len(siblings)
            else None
        )
        audit.true(f"{label} must be immediately followed by a table", next_element is not None and next_element.tag == W + "tbl")
        if next_element is not None and next_element.tag == W + "tbl":
            rows = next_element.findall(W + "tr")
            audit.true(f"{label}.table must contain a header row", bool(rows))
            if rows:
                audit.true(
                    f"{label}.table header must repeat",
                    enabled(child(child(rows[0], "trPr"), "tblHeader")),
                )

    caption_bookmarks: set[str] = set()
    caption_number_bookmarks: set[str] = set()
    caption_seq_count = 0
    registered_sequences = {
        "图": {"图"},
        "表": {"表"},
    }
    for caption_index, paragraph in enumerate(captions, start=1):
        text = paragraph_text(paragraph).strip()
        expected_kind = "表" if text.startswith("表") else "图"
        seq_fields: list[tuple[str, str]] = []
        for instruction in paragraph_field_instructions(paragraph):
            match = re.search(
                r'(?:^| )SEQ\s+(?P<name>"[^"]+"|\S+)',
                instruction,
                re.IGNORECASE,
            )
            if match:
                seq_fields.append(
                    (instruction, match.group("name").strip('"'))
                )
        audit.equal(f"{prefix}.Caption[{caption_index}].SEQ.count", len(seq_fields), 1)
        if len(seq_fields) == 1:
            audit.true(
                f"{prefix}.Caption[{caption_index}].SEQ must use the registered {expected_kind} label",
                seq_fields[0][1] in registered_sequences[expected_kind],
            )
        caption_seq_count += len(seq_fields)
        local_bookmarks = {
            attr(node, "name") or ""
            for node in paragraph.findall(f".//{W}bookmarkStart")
            if re.fullmatch(r"_Ref\d+", attr(node, "name") or "")
        }
        local_number_bookmarks = {
            attr(node, "name") or ""
            for node in paragraph.findall(f".//{W}bookmarkStart")
            if re.fullmatch(r"_RefNum\d+", attr(node, "name") or "")
        }
        audit.equal(f"{prefix}.Caption[{caption_index}].bookmark.full-count", len(local_bookmarks), 1)
        audit.equal(f"{prefix}.Caption[{caption_index}].bookmark.number-count", len(local_number_bookmarks), 1)
        caption_bookmarks.update(local_bookmarks)
        caption_number_bookmarks.update(local_number_bookmarks)
        if len(local_bookmarks) == 1:
            bookmark_name = next(iter(local_bookmarks))
            direct_children = list(paragraph)
            starts = [
                child for child in direct_children
                if child.tag == W + "bookmarkStart" and attr(child, "name") == bookmark_name
            ]
            if len(starts) == 1:
                start = starts[0]
                bookmark_id = attr(start, "id")
                ends = [
                    child for child in direct_children
                    if child.tag == W + "bookmarkEnd" and attr(child, "id") == bookmark_id
                ]
                audit.equal(
                    f"{prefix}.Caption[{caption_index}].bookmark.end-count",
                    len(ends),
                    1,
                )
                if len(ends) == 1:
                    start_index = direct_children.index(start)
                    end_index = direct_children.index(ends[0])
                    inside = direct_children[start_index + 1 : end_index]
                    inside_fields = field_instructions_from_nodes(inside)
                    audit.equal(
                        f"{prefix}.Caption[{caption_index}].bookmark.SEQ-count",
                        sum(
                            1
                            for instruction in inside_fields
                            if re.search(r"(?:^| )SEQ(?: |$)", instruction, re.IGNORECASE)
                        ),
                        1,
                    )
                    audit.true(
                        f"{prefix}.Caption[{caption_index}].full bookmark must include visible label",
                        "".join(
                            node.text or ""
                            for child_node in inside
                            for node in child_node.iter(W + "t")
                        ).startswith(expected_kind),
                    )
    audit.equal(f"{prefix}.Caption.SEQ.total", caption_seq_count, len(captions))
    audit.equal(f"{prefix}.Caption.bookmark.total", len(caption_bookmarks), len(captions))
    audit.equal(
        f"{prefix}.Caption.number-bookmark.total",
        len(caption_number_bookmarks),
        len(captions),
    )

    equation_paragraphs: list[ET.Element] = []
    for paragraph in paragraphs:
        equation_sequences = []
        for instruction in paragraph_field_instructions(paragraph):
            match = re.search(
                r'(?:^| )SEQ\s+(?P<name>"[^"]+"|\S+)',
                instruction,
                re.IGNORECASE,
            )
            if match and match.group("name").strip('"') == "公式":
                equation_sequences.append(instruction)
        if not equation_sequences:
            continue
        equation_paragraphs.append(paragraph)
        audit.equal(
            f"{prefix}.Equation[{len(equation_paragraphs)}].SEQ.count",
            len(equation_sequences),
            1,
        )
        audit.equal(
            f"{prefix}.Equation[{len(equation_paragraphs)}].style",
            paragraph_style(paragraph, by_id)[0].casefold(),
            "宜宾论文-公式".casefold(),
        )
        full_names = {
            attr(node, "name") or ""
            for node in paragraph.findall(f".//{W}bookmarkStart")
            if re.fullmatch(r"_Ref\d+", attr(node, "name") or "")
        }
        number_names = {
            attr(node, "name") or ""
            for node in paragraph.findall(f".//{W}bookmarkStart")
            if re.fullmatch(r"_RefNum\d+", attr(node, "name") or "")
        }
        audit.equal(
            f"{prefix}.Equation[{len(equation_paragraphs)}].bookmark.full-count",
            len(full_names),
            1,
        )
        audit.equal(
            f"{prefix}.Equation[{len(equation_paragraphs)}].bookmark.number-count",
            len(number_names),
            1,
        )

    field_instructions = [
        instruction
        for paragraph in paragraphs
        for instruction in paragraph_field_instructions(paragraph)
    ]
    ref_targets = [
        match.group(1)
        for instruction in field_instructions
        if (match := re.match(r"REF ([^ \\]+)", instruction, re.IGNORECASE))
    ]
    for target in ref_targets:
        audit.true(f"{prefix}.REF target is missing: {target}", target in bookmark_names)
    unresolved_markers = ("YIBIN_INTERNAL_REF", "[[REF:", "错误！未找到引用源", "Error! Reference source not found")
    visible_text = "\n".join(paragraph_text(paragraph) for paragraph in paragraphs)
    for marker in unresolved_markers:
        audit.true(f"{prefix}.unresolved cross-reference marker remains: {marker}", marker not in visible_text)

    citation_bookmarks = {
        name for name in bookmark_names if name.startswith("YibinCitation_")
    }
    hyperlink_targets = {
        node.get(W + "anchor", "") for node in document.findall(f".//{W}hyperlink")
    }
    for instruction in field_instructions:
        match = re.match(r'HYPERLINK \\l "([^"]+)"', instruction, re.IGNORECASE)
        if match:
            hyperlink_targets.add(match.group(1))
    for instruction in field_instructions:
        match = re.match(r"REF\s+([^ \\]+)(?P<switches>.*)", instruction, re.IGNORECASE)
        if not match:
            continue
        target = match.group(1)
        if re.fullmatch(r"_Ref(?:Num)?\d+", target):
            audit.true(
                f"{prefix}.native REF must include hyperlink switch: {target}",
                re.search(r"(?:^|\s)\\h(?:\s|$)", match.group("switches")) is not None,
            )
            audit.true(
                f"{prefix}.native REF must not be wrapped in explicit w:hyperlink: {target}",
                target not in hyperlink_targets,
            )
    citation_ref_targets = set(ref_targets) | hyperlink_targets
    citation_fields = [
        instruction for instruction in field_instructions if re.match(r"CITATION(?: |$)", instruction, re.IGNORECASE)
    ]
    field_records = field_records_from_nodes([document])
    citation_records = [
        (instruction, locked)
        for instruction, locked in field_records
        if re.match(r"CITATION(?: |$)", instruction, re.IGNORECASE)
    ]
    citation_details = [
        (instruction, locked, result_runs)
        for instruction, locked, result_runs in detailed_field_records([document])
        if re.match(
            r"(?:CITATION(?: |$)|REF\s+YibinCitation_)",
            instruction,
            re.IGNORECASE,
        )
    ]
    audit.true(f"{prefix}.citation field is missing", bool(citation_details))
    for citation_index, (instruction, _locked, result_runs) in enumerate(
        citation_details,
        start=1,
    ):
        audit.true(
            f"{prefix}.citation[{citation_index}] must preserve its character style",
            "CHARFORMAT" in instruction.upper(),
        )
        audit.true(
            f"{prefix}.citation[{citation_index}] cached result is missing",
            bool(result_runs),
        )
        for run in result_runs:
            audit.equal(
                f"{prefix}.citation[{citation_index}].result-style",
                effective_run_style_name(run, by_id),
                "宜宾论文-文献上标",
            )

    citation_marker_runs = [
        run
        for run in document.findall(f".//{W}r")
        if effective_run_style_name(run, by_id) == "宜宾论文-文献上标"
        and "".join(node.text or "" for node in run.findall(W + "t"))
        in {"[", "]", ","}
    ]
    marker_text = [
        "".join(node.text or "" for node in run.findall(W + "t"))
        for run in citation_marker_runs
    ]
    audit.true(f"{prefix}.citation superscript brackets are missing", bool(marker_text))
    audit.equal(f"{prefix}.citation bracket balance", marker_text.count("["), marker_text.count("]"))
    bibliography_sources: list[ET.Element] = []
    bibliography_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/bibliography}"
    for member, payload in custom_xml.items():
        try:
            custom_root = ET.fromstring(payload)
        except ET.ParseError:
            continue
        if custom_root.tag == bibliography_ns + "Sources":
            bibliography_sources.extend(custom_root.findall(bibliography_ns + "Source"))

    citation_mode = "native" if citation_fields or bibliography_sources else "linked"
    bibliography_paragraphs = [
        *named.get("bibliography", []),
        *named.get("宜宾论文-参考文献", []),
        *named.get("书目1", []),
        *named.get("bibliography 1", []),
    ]
    if bibliography_paragraphs:
        if citation_mode == "native":
            audit.true(f"{prefix}.native CITATION field is missing", bool(citation_fields))
            audit.true(f"{prefix}.native bibliography customXml Source is missing", bool(bibliography_sources))
            for index, (_instruction, locked) in enumerate(citation_records, start=1):
                audit.true(f"{prefix}.native CITATION[{index}] must be locked", locked)
            citation_keys = {
                match.group(1)
                for instruction in citation_fields
                if (match := re.match(r"CITATION ([^ ]+)", instruction, re.IGNORECASE))
            }
            source_keys = {
                (source.findtext(bibliography_ns + "Tag") or "").strip()
                for source in bibliography_sources
            }
            audit.equal(f"{prefix}.native bibliography source keys", source_keys, citation_keys)
            for key in sorted(source_keys):
                source = next(
                    source
                    for source in bibliography_sources
                    if (source.findtext(bibliography_ns + "Tag") or "").strip() == key
                )
                for field_name in ("SourceType", "Guid", "RefOrder", "Title"):
                    audit.true(
                        f"{prefix}.native source {key!r} missing {field_name}",
                        bool((source.findtext(bibliography_ns + field_name) or "").strip()),
                    )
            custom_targets = {
                posixpath.normpath(posixpath.join("word", relationship.get("Target", "")))
                for relationship in relationships
                if relationship.get("Type", "").endswith("/customXml")
            }
            bibliography_members = {
                member
                for member, payload in custom_xml.items()
                if ET.fromstring(payload).tag == bibliography_ns + "Sources"
            }
            for member in bibliography_members:
                audit.true(
                    f"{prefix}.native bibliography customXml part is not package-related: {member}",
                    member in custom_targets,
                )
                item_match = re.search(r"item(\d+)\.xml$", member)
                if item_match:
                    item_index = item_match.group(1)
                    audit.true(
                        f"{prefix}.native bibliography itemProps missing",
                        f"customXml/itemProps{item_index}.xml" in members,
                    )
                    audit.true(
                        f"{prefix}.native bibliography item relationship missing",
                        f"customXml/_rels/item{item_index}.xml.rels" in members,
                    )
        else:
            audit.true(f"{prefix}.linked citation bookmark is missing", bool(citation_bookmarks))
            for name in citation_bookmarks:
                audit.true(
                    f"{prefix}.linked citation bookmark is not referenced: {name}",
                    name in citation_ref_targets,
                )

    # Audit only semantic tables: layout tables have no repeating-header flag
    # and no adjacent Table Caption.  This still catches a missing tblHeader
    # when the caption identifies the table.
    body = child(document, "body")
    body_children = [] if body is None else list(body)
    semantic_tables: list[ET.Element] = []
    for child_index, element in enumerate(body_children):
        if element.tag != W + "tbl":
            continue
        nearby = body_children[max(0, child_index - 2) : child_index]
        has_caption = any(
            node.tag == W + "p"
            and paragraph_style(node, by_id)[0].casefold()
            in {"table caption", "宜宾论文-表题"}
            for node in nearby
        )
        has_header = bool(element.findall(f"{W}tr/{W}trPr/{W}tblHeader"))
        if has_caption or has_header:
            semantic_tables.append(element)
    for table_index, table in enumerate(semantic_tables, start=1):
        rows = table.findall(W + "tr")
        audit.true(f"{prefix}.table[{table_index}] must contain a header row", bool(rows))
        if not rows:
            continue
        audit.true(
            f"{prefix}.table[{table_index}] header must repeat",
            enabled(child(child(rows[0], "trPr"), "tblHeader")),
        )


def audit_docx(audit: Audit, docx: Path, profile: str | None = None) -> None:
    audit.true(f"DOCX not found: {docx}", docx.is_file())
    if not docx.is_file():
        return
    try:
        styles_root = load_part(docx, "word/styles.xml")
        styles = style_map(styles_root)
        by_id = styles_by_id(styles_root)
        audit_geometry(audit, docx)
    except (KeyError, OSError, ET.ParseError, zipfile.BadZipFile) as exc:
        audit.errors.append(f"cannot inspect {docx}: {exc}")
        return

    audit_table_continuation_style(audit, styles, docx.name)

    audit_style(
        audit,
        styles,
        "Normal",
        east_asia="SimSun",
        latin="Times New Roman",
        size="21",
        bold=False,
        line="360",
        line_rule="auto",
        first_chars="200",
    )
    common = (
        ("Body Text", "SimSun", "Times New Roman", "24", False, "360", "200"),
        ("First Paragraph", "SimSun", "Times New Roman", "24", False, "360", "200"),
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

    audit_style(
        audit,
        styles,
        "宜宾论文-列表正文",
        east_asia="SimSun",
        latin="Times New Roman",
        size="24",
        bold=False,
        line="360",
        line_rule="auto",
        first_chars="0",
        first_line="0",
    )

    audit_style(
        audit,
        styles,
        "宜宾论文-附录二级标题",
        east_asia="KaiTi",
        latin="Times New Roman",
        size="30",
        bold=True,
        line="360",
        line_rule="auto",
        first_chars="200",
    )

    for name, east_asia, size in (
        ("宜宾论文-一级标题", "SimHei", "32"),
        ("宜宾论文-二级标题", "KaiTi", "30"),
        ("宜宾论文-三级标题", "SimSun", "28"),
        ("宜宾论文-四级标题", "SimSun", "28"),
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
            page_break=name == "宜宾论文-一级标题",
        )

    custom_heading_ids: set[str] = set()
    for level, custom_name in enumerate(
        (
            "宜宾论文-一级标题",
            "宜宾论文-二级标题",
            "宜宾论文-三级标题",
            "宜宾论文-四级标题",
        ),
        start=1,
    ):
        builtin_id = f"Heading{level}"
        builtin = by_id.get(builtin_id)
        audit.true(f"{docx.name}.{builtin_id} built-in style is missing", builtin is not None)
        if builtin is not None:
            audit.equal(
                f"{docx.name}.{builtin_id} built-in name",
                style_name(builtin).casefold(),
                f"heading {level}",
            )
        custom = styles.get(custom_name.casefold())
        if custom is not None:
            custom_id = attr(custom, "styleId") or ""
            custom_heading_ids.add(custom_id)
            audit.true(
                f"{docx.name}.{custom_name} must not reuse {builtin_id}",
                custom_id != builtin_id,
            )
    audit.equal(
        f"{docx.name}.custom heading styleId count",
        len(custom_heading_ids),
        4,
    )
    audit_style(
        audit,
        styles,
        "宜宾论文-公式",
        east_asia="SimSun",
        latin="Times New Roman",
        size="24",
        bold=False,
        line="240",
        line_rule="auto",
        first_chars="0",
        first_line="0",
        alignment="left",
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
        latin="SimHei",
        size="32",
        bold=True,
        line="240",
        alignment="center",
        before="0",
        after="0",
    )
    audit_style(
        audit,
        styles,
        "Yibin TOC Heading",
        east_asia="SimHei",
        latin="Times New Roman",
        size="32",
        bold=True,
        line="240",
        first_chars="0",
        first_line="0",
        alignment="center",
        before="0",
        after="0",
    )
    for level, left, left_chars, line in (
        (1, "0", "0", "360"),
        (2, "420", "200", "360"),
        (3, "840", "400", "240"),
    ):
        audit_style(
            audit,
            styles,
            f"TOC {level}",
            east_asia="SimSun",
            latin="SimSun",
            size="24",
            bold=False,
            line=line,
            line_rule="auto",
            first_chars="0",
            first_line="0",
            left=left,
            left_chars=left_chars,
            alignment="both",
            before="0",
            after="0",
            tab_pos="8777",
        )

    audit_style(
        audit,
        styles,
        "Table of Figures",
        east_asia="SimSun",
        latin="SimSun",
        size="24",
        bold=False,
        line="360",
        line_rule="auto",
        first_chars="0",
        first_line="0",
        left="0",
        left_chars="0",
        alignment="both",
        before="0",
        after="0",
        tab_pos="8777",
    )

    audit_style(
        audit,
        styles,
        "CoverThesisType",
        east_asia="SimHei",
        latin="Times New Roman",
        size="56",
        bold=True,
        before="487",
        after="1047",
        alignment="center",
    )
    audit_style(
        audit,
        styles,
        "CoverVersion",
        east_asia="SimHei",
        latin="Times New Roman",
        size="28",
        bold=False,
        before="0",
        after="460",
        alignment="center",
    )
    audit_style(
        audit,
        styles,
        "CoverTitle",
        east_asia="SimHei",
        latin="Times New Roman",
        size="36",
        bold=True,
        underline=True,
        alignment="center",
    )
    audit_style(
        audit,
        styles,
        "CoverTitleWithVersion",
        east_asia="SimHei",
        latin="Times New Roman",
        size="36",
        bold=True,
        underline=True,
        before="0",
        after="200",
        alignment="center",
    )
    audit_style(
        audit,
        styles,
        "CoverField",
        east_asia="SimSun",
        latin="Times New Roman",
        size="32",
        bold=True,
        line="240",
        line_rule="auto",
        alignment="both",
    )
    audit_style(
        audit,
        styles,
        "CoverValueLine",
        east_asia="SimSun",
        latin="Times New Roman",
        size="32",
        bold=True,
        line="400",
        line_rule="exact",
        alignment="center",
        before="0",
        after="0",
    )
    audit_style(
        audit,
        styles,
        "CoverDate",
        east_asia="SimSun",
        latin="Times New Roman",
        size="30",
        bold=True,
        before="480",
        after="0",
        alignment="center",
    )
    audit_style(
        audit,
        styles,
        "CoverLabel",
        style_type="character",
        east_asia="SimHei",
        latin="Times New Roman",
        size="36",
        bold=True,
    )
    audit_style(
        audit,
        styles,
        "CoverValue",
        style_type="character",
        east_asia="SimSun",
        latin="Times New Roman",
        size="32",
        bold=True,
    )
    audit_style(
        audit,
        styles,
        "DeclarationBody",
        east_asia="SimSun",
        latin="Times New Roman",
        size="24",
        bold=False,
        line="480",
        line_rule="auto",
        first_line="480",
    )
    audit_style(
        audit,
        styles,
        "DeclarationRegulationLead",
        east_asia="SimSun",
        latin="SimSun",
        size="28",
        bold=True,
    )
    audit_style(
        audit,
        styles,
        "DeclarationRegulationClause",
        east_asia="SimSun",
        latin="SimSun",
        size="28",
        bold=True,
        line="360",
        line_rule="auto",
    )
    audit_style(
        audit,
        styles,
        "DeclarationRegulationItem",
        east_asia="SimSun",
        latin="SimSun",
        size="28",
        bold=True,
        line="360",
        line_rule="auto",
    )
    audit_style(
        audit,
        styles,
        "DeclarationRegulationLeadLabel",
        style_type="character",
        east_asia="SimSun",
        latin="SimSun",
        size="32",
        bold=True,
    )
    audit_style(
        audit,
        styles,
        "DeclarationRegulationText",
        style_type="character",
        east_asia="SimSun",
        latin="SimSun",
        size="28",
        bold=True,
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
        size="21",
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
        "unnumbered chapters are written to the TOC": r"\addcontentsline{toc}{chapter}{#1}",
        "science equations use point-separated chapter numbering": (
            r"\renewcommand{\theequation}{\arabic{chapter}.\arabic{equation}}"
        ),
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
    closing_chapter = (
        r"\NewDocumentCommand{\yibinclosingchapter}{O{结论}}{%"
        "\n  "
        r"\yibinunnumberedchapter{#1}"
        "\n}"
    )
    audit.true(
        "LaTeX science/humanities closing chapter must be unnumbered and remain in the TOC",
        closing_chapter in text,
    )


def audit_word_builder(audit: Audit, builder: Path) -> None:
    audit.true(f"Word builder not found: {builder}", builder.is_file())
    if not builder.is_file():
        return
    text = builder.read_text(encoding="utf-8")
    audit.true(
        "Word builder must normalize yibinclosingchapter as an unnumbered chapter",
        'lambda match: unnumbered_thematic_chapter(match, "结论")' in text,
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--class-file", type=Path, default=ROOT / "yibinthesis.cls")
    parser.add_argument("--reference", type=Path, default=ROOT / "word" / "reference.docx")
    parser.add_argument(
        "--generated",
        type=Path,
        action="append",
        default=None,
        help=(
            "Generated DOCX to audit; may be repeated. If omitted, existing "
            "build/word/main.docx and full-featured.docx are audited."
        ),
    )
    parser.add_argument(
        "--manuscript-root",
        type=Path,
        action="append",
        default=None,
        help="Additional manuscript source root to audit; may be repeated.",
    )
    parser.add_argument(
        "--logo",
        type=Path,
        default=ROOT / "assets" / "yibin-university-logo.png",
    )
    args = parser.parse_args()

    audit = Audit()
    audit_latex(audit, args.class_file)
    audit_word_builder(audit, ROOT / "lib" / "word_core.py")
    audit_docx(audit, args.reference)
    audit_logo_asset(audit, args.logo)
    generated = args.generated
    if generated is None:
        generated = [
            candidate
            for candidate in (
                ROOT / "build" / "word" / "main.docx",
                ROOT / "build" / "word" / "full-featured.docx",
            )
            if candidate.is_file()
        ]
    for docx in generated:
        audit_generated_docx(audit, docx, args.logo)

    if audit.errors:
        print("FORMAT AUDIT: FAIL", file=sys.stderr)
        for error in audit.errors:
            print(f"- {error}", file=sys.stderr)
        return 1
    print("FORMAT AUDIT: PASS")
    print("- SimSun/SimHei/KaiTi/Times New Roman style mapping")
    print("- 12 pt body, 1.5-line spacing, two-character Chinese first-line indent after headings")
    print("- flush-left English abstract, heading ladder, captions, notes, references")
    print("- point-separated science counters and unnumbered conclusion retained in the TOC")
    print("- A4 geometry and school margins/header/footer distances")
    print("- official cover/logo, originality regulation block, TOC indents and embedded images")
    print("- single-border centred cover value slots without underlined padding")
    print("- editable Caption/SEQ/bookmark/REF fields and linked/native citations")
    print("- structured continuation captions and repeating table headers")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
