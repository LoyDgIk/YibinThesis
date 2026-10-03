"""Read-only structural audit for school-supplied DOCX templates.

This module deliberately does not edit the source documents. It extracts the
layout contract that the LaTeX and Word builders must implement: sections,
page geometry, recurring text styles, table grids, cell borders and page
number fields. Use ``--json`` to save a machine-readable audit report.
"""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import asdict, dataclass
import hashlib
import json
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _attr(node: Any, name: str) -> str | None:
    if node is None:
        return None
    return node.get(qn(f"w:{name}"))


def _child(node: Any, name: str) -> Any:
    return None if node is None else node.find(qn(f"w:{name}"))


def _twips(value: str | None) -> float | None:
    return None if value is None else round(int(value) / 20, 2)


def _run_signature(run: Any) -> dict[str, Any]:
    rpr = run._r.rPr
    fonts = _child(rpr, "rFonts")
    size = _child(rpr, "sz")
    return {
        "eastAsia": _attr(fonts, "eastAsia"),
        "ascii": _attr(fonts, "ascii"),
        "hAnsi": _attr(fonts, "hAnsi"),
        "sizePt": None if size is None else int(_attr(size, "val")) / 2,
        "bold": _child(rpr, "b") is not None,
        "underline": None if _child(rpr, "u") is None else _attr(_child(rpr, "u"), "val"),
    }


def _paragraph_summary(paragraph: Any) -> dict[str, Any]:
    ppr = paragraph._p.pPr
    spacing = _child(ppr, "spacing")
    indent = _child(ppr, "ind")
    return {
        "style": paragraph.style.name,
        "text": paragraph.text,
        "alignment": None if paragraph.alignment is None else str(paragraph.alignment),
        "indentPt": {
            "left": _twips(_attr(indent, "left")),
            "right": _twips(_attr(indent, "right")),
            "firstLine": _twips(_attr(indent, "firstLine")),
            "hanging": _twips(_attr(indent, "hanging")),
        },
        "spacing": {
            "beforePt": _twips(_attr(spacing, "before")),
            "afterPt": _twips(_attr(spacing, "after")),
            "line": _attr(spacing, "line"),
            "lineRule": _attr(spacing, "lineRule"),
        },
        "runs": [_run_signature(run) for run in paragraph.runs if run.text],
    }


def _table_summary(table: Any) -> dict[str, Any]:
    grid = table._tbl.tblGrid
    grid_widths = [] if grid is None else [int(_attr(col, "w")) for col in grid.findall(qn("w:gridCol"))]
    rows: list[dict[str, Any]] = []
    for index, row in enumerate(table.rows):
        cells = []
        for cell in row.cells:
            tcpr = cell._tc.tcPr
            width = _child(tcpr, "tcW")
            shading = _child(tcpr, "shd")
            borders = _child(tcpr, "tcBorders")
            cells.append({
                "widthTwips": None if width is None else int(_attr(width, "w")),
                "text": " / ".join(p.text for p in cell.paragraphs),
                "fill": _attr(shading, "fill"),
                "borders": None if borders is None else {
                    edge.tag.rsplit("}", 1)[-1]: {
                        "val": _attr(edge, "val"), "size": _attr(edge, "sz"), "color": _attr(edge, "color")
                    } for edge in borders
                },
                "paragraphs": [_paragraph_summary(p) for p in cell.paragraphs],
            })
        trpr = _child(row._tr.trPr, "trHeight")
        rows.append({
            "index": index,
            "heightTwips": None if trpr is None else int(_attr(trpr, "val")),
            "heightRule": None if trpr is None else _attr(trpr, "hRule"),
            "repeatHeader": _child(row._tr.trPr, "tblHeader") is not None,
            "cells": cells,
        })
    return {"rows": len(table.rows), "columns": len(table.columns), "gridTwips": grid_widths, "rowsDetail": rows}


def audit_docx(path: Path) -> dict[str, Any]:
    document = Document(path)
    fonts = Counter()
    sizes = Counter()
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            for key in ("eastAsia", "ascii", "hAnsi"):
                value = _run_signature(run).get(key)
                if value:
                    fonts[value] += len(run.text or "")
            size = _run_signature(run).get("sizePt")
            if size:
                sizes[size] += len(run.text or "")
    sections = []
    for section in document.sections:
        sect = section._sectPr
        page_size = _child(sect, "pgSz")
        margins = _child(sect, "pgMar")
        sections.append({
            "pageWidthTwips": _attr(page_size, "w"), "pageHeightTwips": _attr(page_size, "h"),
            "orientation": _attr(page_size, "orient"),
            "marginsTwips": {name: _attr(margins, name) for name in ("top", "right", "bottom", "left", "header", "footer", "gutter")},
            "startType": _attr(_child(sect, "type"), "val"),
            "differentFirstPage": _child(sect, "titlePg") is not None,
            "headerLinked": section.header.is_linked_to_previous,
            "footerLinked": section.footer.is_linked_to_previous,
        })
    with ZipFile(path) as archive:
        package_parts = archive.namelist()
    return {
        "path": str(path), "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "sections": sections, "paragraphCount": len(document.paragraphs), "tableCount": len(document.tables),
        "fontsByCharacterCount": dict(fonts), "sizesByCharacterCount": {str(k): v for k, v in sizes.items()},
        "nonEmptyParagraphs": [_paragraph_summary(p) for p in document.paragraphs if p.text.strip()],
        "tables": [_table_summary(table) for table in document.tables],
        "packageParts": package_parts,
        "headers": [part for part in package_parts if part.startswith("word/header")],
        "footers": [part for part in package_parts if part.startswith("word/footer")],
        "comments": [part for part in package_parts if "comments" in part],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path, nargs="+")
    parser.add_argument("--json", type=Path, help="写入 JSON 报告；多个输入时写入目录")
    args = parser.parse_args()
    reports = [audit_docx(path.expanduser().resolve()) for path in args.docx]
    if args.json:
        if len(reports) == 1 and args.json.suffix.lower() == ".json":
            args.json.parent.mkdir(parents=True, exist_ok=True)
            args.json.write_text(json.dumps(reports[0], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        else:
            args.json.mkdir(parents=True, exist_ok=True)
            for report in reports:
                (args.json / (Path(report["path"]).stem + ".json")).write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    for report in reports:
        print(f"{report['path']}: sections={len(report['sections'])}, paragraphs={report['paragraphCount']}, tables={report['tableCount']}")
        print(f"  fonts={report['fontsByCharacterCount']}")
        print(f"  section_geometry={report['sections']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
