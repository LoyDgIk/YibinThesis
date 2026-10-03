#!/usr/bin/env python3
"""Build an editable YibinThesis DOCX from the template's LaTeX sources.

This converter intentionally supports the controlled source subset used by this
project: metadata in ``metadata.tex``, the ordered ``main.tex`` input list,
standard sectioning, ordinary paragraphs, Pandoc-readable tables/figures/math,
citations, and the YibinThesis abstract/note/unnumbered-chapter commands.

PDF remains the layout-primary output.  Complex TikZ, arbitrary custom macros,
manual page geometry, and unusually nested floats cannot be guaranteed to
round-trip to Word and are reported when detected.
"""

from __future__ import annotations

import argparse
import copy
import html as html_lib
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree as ET

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import (
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_ROW_HEIGHT_RULE,
    WD_TABLE_ALIGNMENT,
)
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


SECTION_COVER_END = "YIBIN_INTERNAL_SECTION_COVER_END_5B8D"
SECTION_ABSTRACT_END = "YIBIN_INTERNAL_SECTION_ABSTRACT_END_6458"
SECTION_FRONT_END = "YIBIN_INTERNAL_SECTION_FRONT_END_8F91"
TOC_MARKER = "YIBIN_INTERNAL_TOC_76EE"
LOF_MARKER = "YIBIN_INTERNAL_LOF_56FE"
LOT_MARKER = "YIBIN_INTERNAL_LOT_8868"
SECTION_REVIEW_COVER_END = "YIBIN_INTERNAL_SECTION_REVIEW_COVER_END_A2C4"
SECTION_REVIEW_REFERENCES = "YIBIN_INTERNAL_SECTION_REVIEW_REFERENCES_7B19"
SECTION_REVIEW_TAIL = "YIBIN_INTERNAL_SECTION_REVIEW_TAIL_91E2"
PROPOSAL_FIELD_MARKER_PREFIX = "YIBIN_INTERNAL_PROPOSAL_FIELD_"
STYLE_BODY = "宜宾论文-正文"
STYLE_FIRST_PARAGRAPH = "宜宾论文-首段"
STYLE_LIST_BODY = "宜宾论文-列表正文"
STYLE_HEADING_1 = "宜宾论文-一级标题"
STYLE_HEADING_2 = "宜宾论文-二级标题"
STYLE_HEADING_3 = "宜宾论文-三级标题"
STYLE_HEADING_4 = "宜宾论文-四级标题"
BUILTIN_HEADING_STYLES = (
    "Heading 1",
    "Heading 2",
    "Heading 3",
    "Heading 4",
)
BUILTIN_HEADING_STYLE_IDS = (
    "Heading1",
    "Heading2",
    "Heading3",
    "Heading4",
)
STYLE_UNNUMBERED_HEADING = "宜宾论文-无编号标题"
STYLE_FRONT_TITLE = "宜宾论文-中文页标题"
STYLE_ENGLISH_ABSTRACT_TITLE = "宜宾论文-英文摘要标题"
STYLE_TOC_TITLE = "宜宾论文-目录标题"
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
STYLE_PROPOSAL_FORM_TABLE = "YibinProposalForm"
STYLE_REVIEW_INFO_TABLE = "YibinReviewInfoLayout"
FIGURE_FILTER = (
    Path(__file__).resolve().parents[1]
    / "word"
    / "filters"
    / "latex-figure-to-image.lua"
)

PAGE_BREAK = """```{=openxml}
<w:p><w:r><w:br w:type="page"/></w:r></w:p>
```"""

UNSUPPORTED_PATTERNS = {
    r"\\begin\s*\{tikzpicture\}": "TikZ 图形",
    r"\\begin\s*\{pspicture\}": "PSTricks 图形",
    r"\\begin\s*\{minted\}": "minted 代码环境",
    r"\\begin\s*\{algorithm2e\}": "algorithm2e 环境",
    r"\\write18\b": "shell-escape 命令",
}


class BuildError(RuntimeError):
    """A user-actionable build failure."""


@dataclass
class MainEvent:
    phase: str
    kind: str
    value: str | None = None


@dataclass(frozen=True)
class DocumentProfile:
    document_type: str
    discipline: str
    template_year: str


@dataclass
class NoteRegistry:
    entries: list[str] = field(default_factory=list)

    def add(self, value: str) -> str:
        self.entries.append(value.strip())
        number = len(self.entries)
        marker = chr(0x245F + number) if number <= 20 else f"[{number}]"
        return rf"\textsuperscript{{{marker}}}"


@dataclass
class HeadingCounters:
    chapter: int = 0
    section: int = 0
    subsection: int = 0
    subsubsection: int = 0


@dataclass
class FloatCounters:
    chapter: int | None = None
    figure: int = 0
    table: int = 0
    equation: int = 0

    def synchronize(self, chapter: int, discipline: str) -> None:
        if discipline == "science" and self.chapter != chapter:
            self.figure = self.table = 0
        if self.chapter != chapter:
            self.equation = 0
        self.chapter = chapter


@dataclass(frozen=True)
class LatexTableColumn:
    """Column semantics distilled from a LaTeX table specification."""

    horizontal: str
    vertical: str = "center"
    width_fraction: float | None = None
    flexible: bool = False
    first_line_indent_pt: float | None = None


@dataclass(frozen=True)
class LatexTableLayout:
    environment: str
    columns: tuple[LatexTableColumn, ...]
    label: str | None = None
    table_width_fraction: float | None = None
    tabcolsep_pt: float | None = None


@dataclass(frozen=True)
class _LatexTablePreamble:
    environment: str
    begin_start: int
    specification_start: int
    specification_end: int
    specification: str
    table_width_fraction: float | None
    tabcolsep_pt: float | None


@dataclass(frozen=True)
class LabelTarget:
    number: str
    kind: str


@dataclass
class LabelRegistry:
    """Track LaTeX labels and deferred references across all source files."""

    targets: dict[str, LabelTarget] = field(default_factory=dict)
    references: dict[str, tuple[str, bool]] = field(default_factory=dict)
    next_reference: int = 0

    def placeholder(self, label: str, *, parenthesized: bool) -> str:
        label = label.strip()
        if not label:
            raise BuildError("检测到空的 LaTeX 交叉引用标签。")
        self.next_reference += 1
        token = f"YIBINXREF{self.next_reference:08d}"
        self.references[token] = (label, parenthesized)
        return token

    def register(self, label: str, number: str, kind: str) -> None:
        label = html_lib.unescape(label.strip())
        if not label:
            return
        if label in self.targets:
            previous = self.targets[label]
            raise BuildError(
                f"LaTeX 标签重复：{label}（{previous.kind} {previous.number} / "
                f"{kind} {number}）。"
            )
        self.targets[label] = LabelTarget(number=number, kind=kind)

    def resolve(self, text: str) -> str:
        missing: set[str] = set()
        for token, (label, parenthesized) in self.references.items():
            target = self.targets.get(label)
            if target is None:
                missing.add(label)
                continue
            # Keep opaque markers until the OOXML stage.  Replacing them here
            # would permanently flatten Word cross-references into plain text.
        if missing:
            raise BuildError("未找到交叉引用标签：" + "、".join(sorted(missing)))
        return text


@dataclass
class CitationRegistry:
    clusters: dict[str, list[str]] = field(default_factory=dict)
    order: list[str] = field(default_factory=list)
    next_cluster: int = 0

    def placeholder(self, keys: str) -> str:
        parsed = [item.strip() for item in split_top_level(keys) if item.strip()]
        if not parsed:
            raise BuildError("检测到空的文献引用键。")
        self.next_cluster += 1
        token = f"YIBINCITE{self.next_cluster:08d}"
        self.clusters[token] = parsed
        for key in parsed:
            if key not in self.order:
                self.order.append(key)
        return token

    def add_nocite(self, keys: str) -> str:
        parsed = [item.strip() for item in split_top_level(keys) if item.strip()]
        if not parsed:
            raise BuildError("检测到空的 \\nocite 文献键。")
        if "*" in parsed:
            raise BuildError("Word 转换暂不支持 \\nocite{*}；请显式列出文献键。")
        for key in parsed:
            if key not in self.order:
                self.order.append(key)
        return ""


def strip_tex_comments(text: str) -> str:
    """Remove unescaped LaTeX comments while retaining line structure."""
    output: list[str] = []
    for line in text.splitlines():
        cut = len(line)
        for index, char in enumerate(line):
            if char != "%":
                continue
            backslashes = 0
            cursor = index - 1
            while cursor >= 0 and line[cursor] == "\\":
                backslashes += 1
                cursor -= 1
            if backslashes % 2 == 0:
                cut = index
                break
        output.append(line[:cut])
    return "\n".join(output)


def extract_balanced(text: str, opening: int) -> tuple[str, int]:
    if opening >= len(text) or text[opening] != "{":
        raise BuildError("内部解析错误：预期花括号参数。")
    depth = 0
    escaped = False
    for index in range(opening, len(text)):
        char = text[index]
        if escaped:
            escaped = False
            continue
        if char == "\\":
            escaped = True
            continue
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[opening + 1 : index], index + 1
    raise BuildError("LaTeX 参数花括号未闭合。")


def find_command_argument(text: str, command: str) -> str | None:
    match = re.search(rf"\\{re.escape(command)}\s*\{{", text)
    if not match:
        return None
    opening = text.find("{", match.start())
    value, _ = extract_balanced(text, opening)
    return value


def replace_braced_command(
    text: str,
    command: str,
    replacement: Callable[[str], str],
) -> str:
    pattern = re.compile(rf"\\{re.escape(command)}\s*\{{")
    cursor = 0
    parts: list[str] = []
    while match := pattern.search(text, cursor):
        opening = text.find("{", match.start())
        value, end = extract_balanced(text, opening)
        parts.append(text[cursor : match.start()])
        parts.append(replacement(value))
        cursor = end
    parts.append(text[cursor:])
    return "".join(parts)


PROPOSAL_FIELD_ORDER = (
    "significance",
    "research-status",
    "research-content",
    "research-approach",
    "schedule",
    "references",
    "advisor-opinion",
)


def extract_proposal_fields(text: str) -> list[tuple[str, str]]:
    r"""Extract balanced ``\yibinproposalfield{key}{content}`` calls."""

    source = strip_tex_comments(text)
    pattern = re.compile(r"\\yibinproposalfield\s*\{")
    cursor = 0
    fields: list[tuple[str, str]] = []
    while match := pattern.search(source, cursor):
        key_open = source.find("{", match.start())
        key, key_end = extract_balanced(source, key_open)
        content_open = key_end
        while content_open < len(source) and source[content_open].isspace():
            content_open += 1
        if content_open >= len(source) or source[content_open] != "{":
            raise BuildError(f"开题报告字段 {key.strip()} 缺少内容参数。")
        content, cursor = extract_balanced(source, content_open)
        fields.append((key.strip(), content.strip()))

    keys = [key for key, _ in fields]
    duplicates = sorted({key for key in keys if keys.count(key) > 1})
    if duplicates:
        raise BuildError("开题报告字段重复：" + "、".join(duplicates))
    unknown = sorted(set(keys) - set(PROPOSAL_FIELD_ORDER))
    if unknown:
        raise BuildError("未知开题报告字段：" + "、".join(unknown))
    missing = [key for key in PROPOSAL_FIELD_ORDER if key not in keys]
    if missing:
        raise BuildError("开题报告字段缺失：" + "、".join(missing))
    if tuple(keys) != PROPOSAL_FIELD_ORDER:
        raise BuildError("开题报告字段顺序必须为：" + "、".join(PROPOSAL_FIELD_ORDER))
    return fields


def _replace_proposal_figure_commands(text: str) -> str:
    pattern = re.compile(r"\\yibinproposalfigure")
    cursor = 0
    parts: list[str] = []
    while match := pattern.search(text, cursor):
        position = match.end()
        while position < len(text) and text[position].isspace():
            position += 1
        width = r"0.82\linewidth"
        if position < len(text) and text[position] == "[":
            close = text.find("]", position + 1)
            if close < 0:
                raise BuildError("开题报告图片宽度参数未闭合。")
            width = text[position + 1 : close].strip() or width
            position = close + 1
        arguments: list[str] = []
        for label in ("路径", "题名", "标签"):
            while position < len(text) and text[position].isspace():
                position += 1
            if position >= len(text) or text[position] != "{":
                raise BuildError(f"开题报告图片缺少{label}参数。")
            value, position = extract_balanced(text, position)
            arguments.append(value.strip())
        path, caption, label = arguments
        parts.append(text[cursor : match.start()])
        parts.append(
            "\\begin{figure}\n\\centering\n"
            f"\\includegraphics[width={width}]{{{path}}}\n"
            f"\\caption{{{caption}}}\n\\label{{{label}}}\n"
            "\\end{figure}"
        )
        cursor = position
    parts.append(text[cursor:])
    return "".join(parts)


def normalize_proposal_latex(text: str) -> str:
    text = _replace_proposal_figure_commands(text)
    table_pattern = re.compile(
        r"\\begin\s*\{yibinproposaltable\}\s*\{(?P<title>[^{}]*)\}\s*"
        r"\{(?P<label>[^{}]*)\}(?P<body>.*?)"
        r"\\end\s*\{yibinproposaltable\}",
        flags=re.DOTALL,
    )

    def replace_table(match: re.Match[str]) -> str:
        return (
            "\\begin{table}\n\\centering\n"
            f"\\caption{{{match.group('title').strip()}}}\n"
            f"\\label{{{match.group('label').strip()}}}\n"
            f"{match.group('body').strip()}\n\\end{{table}}"
        )

    return table_pattern.sub(replace_table, text)


def split_top_level(value: str, delimiter: str = ",") -> list[str]:
    depth = 0
    escaped = False
    start = 0
    items: list[str] = []
    for index, char in enumerate(value):
        if escaped:
            escaped = False
            continue
        if char == "\\":
            escaped = True
            continue
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
        elif char == delimiter and depth == 0:
            items.append(value[start:index])
            start = index + 1
    items.append(value[start:])
    return items


def clean_tex_scalar(value: str) -> str:
    value = value.strip()
    while value.startswith("{") and value.endswith("}"):
        try:
            inner, end = extract_balanced(value, 0)
        except BuildError:
            break
        if end != len(value):
            break
        value = inner.strip()
    replacements = {
        r"\&": "&",
        r"\%": "%",
        r"\#": "#",
        r"\_": "_",
        r"\{": "{",
        r"\}": "}",
        "~": " ",
        r"\\": " ",
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    value = re.sub(r"\\(?:textbf|textit|emph|textrm|textsf|texttt)\s*\{([^{}]*)\}", r"\1", value)
    return re.sub(r"\s+", " ", value).strip()


def parse_metadata(path: Path) -> dict[str, str]:
    text = strip_tex_comments(path.read_text(encoding="utf-8"))
    payload = find_command_argument(text, "yibinsetup")
    if payload is None:
        raise BuildError(f"未在 {path} 中找到 \\yibinsetup{{...}}。")
    metadata: dict[str, str] = {}
    for item in split_top_level(payload):
        if not item.strip():
            continue
        key_value = split_top_level(item, delimiter="=")
        if len(key_value) < 2:
            raise BuildError(f"无法解析元数据项：{item.strip()}")
        key = key_value[0].strip()
        raw_value = "=".join(key_value[1:]).strip()
        metadata[key] = clean_tex_scalar(raw_value)
    return metadata


def format_grade_class(metadata: dict[str, str]) -> str:
    """Return the cover value while preserving legacy numeric grade metadata."""
    grade = metadata.get("grade", "").strip()
    class_name = metadata.get("class-name", "").strip()
    if grade and "级" not in grade:
        grade += "级"
    return grade + class_name


def parse_main_events(main_text: str) -> list[MainEvent]:
    text = strip_tex_comments(main_text)
    pattern = re.compile(
        r"\\(?P<kind>frontmatter|mainmatter|backmatter|"
        r"makeyibincover|makeyibindeclarations|makeyibinproposal|"
        r"makeyibinliteraturereviewcover|tableofcontents|"
        r"listoffigures|listoftables|"
        r"printyibinnotes|printyibinbibliography|nocite|"
        r"printyibinproposalbibliography|input|include)"
        r"(?:\s*\{(?P<value>[^{}]+)\})?"
    )
    phase = "pre"
    events: list[MainEvent] = []
    for match in pattern.finditer(text):
        kind = match.group("kind")
        value = match.group("value")
        if kind == "frontmatter":
            phase = "front"
            events.append(MainEvent(phase, kind))
        elif kind == "mainmatter":
            phase = "main"
            events.append(MainEvent(phase, kind))
        elif kind == "backmatter":
            phase = "back"
            events.append(MainEvent(phase, kind))
        else:
            events.append(MainEvent(phase, kind, value.strip() if value else None))
    return events


def parse_discipline(main_text: str) -> str:
    match = re.search(
        r"\\documentclass(?:\[(?P<options>[^]]*)\])?\s*\{yibinthesis\}",
        strip_tex_comments(main_text),
    )
    if not match:
        return "humanities"
    options = {item.strip().lower() for item in (match.group("options") or "").split(",")}
    return "science" if "science" in options else "humanities"


def parse_document_type(main_text: str) -> str:
    match = re.search(
        r"\\documentclass(?:\[(?P<options>[^]]*)\])?\s*\{yibinthesis\}",
        strip_tex_comments(main_text),
    )
    if not match:
        return "thesis"
    options = {
        item.strip().lower() for item in (match.group("options") or "").split(",")
    }
    selected = options & {"thesis", "proposal", "literature-review"}
    if len(selected) > 1:
        raise BuildError("文档类型类选项互斥：" + "、".join(sorted(selected)))
    return next(iter(selected), "thesis")


def resolve_profile(main_text: str, metadata: dict[str, str]) -> DocumentProfile:
    document_type = parse_document_type(main_text)
    expected_year = "2022" if document_type == "proposal" else "2024"
    template_year = metadata.get("template-year", "").strip() or expected_year
    if template_year != expected_year:
        raise BuildError(
            f"文档类型 {document_type} 仅支持 template-year={expected_year}，"
            f"收到 template-year={template_year}。"
        )
    return DocumentProfile(
        document_type=document_type,
        discipline=parse_discipline(main_text),
        template_year=template_year,
    )


def path_is_within(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
    except ValueError:
        return False
    return True


def resolve_source(
    name: str,
    main_dir: Path,
    project_root: Path,
    *,
    allow_project_fallback: bool = True,
) -> Path:
    raw = Path(name)
    names = [raw] if raw.suffix else [raw.with_suffix(".tex"), raw]
    bases = [main_dir]
    if allow_project_fallback and main_dir.resolve() != project_root.resolve():
        bases.append(project_root)
    for base in bases:
        for candidate_name in names:
            candidate = (base / candidate_name).resolve()
            if candidate.is_file():
                return candidate
    scope = "入口文件目录" if not allow_project_fallback else "入口文件或模板目录"
    raise BuildError(f"找不到主文件引用的源文件：{name}（已检查{scope}）")


def expand_nested_inputs(
    path: Path,
    project_root: Path,
    seen: set[Path] | None = None,
    *,
    allow_project_fallback: bool = True,
) -> str:
    seen = seen or set()
    resolved = path.resolve()
    if resolved in seen:
        raise BuildError(f"检测到循环 \\input：{resolved}")
    seen.add(resolved)
    text = path.read_text(encoding="utf-8")
    pattern = re.compile(r"\\(?:input|include)\s*\{([^{}]+)\}")

    def include(match: re.Match[str]) -> str:
        child = resolve_source(
            match.group(1).strip(),
            path.parent,
            project_root,
            allow_project_fallback=allow_project_fallback,
        )
        return expand_nested_inputs(
            child,
            project_root,
            seen.copy(),
            allow_project_fallback=allow_project_fallback,
        )

    return pattern.sub(include, text)


def extract_environment(text: str, environment: str) -> str | None:
    match = re.search(
        rf"\\begin\s*\{{{re.escape(environment)}\}}(?P<body>.*?)"
        rf"\\end\s*\{{{re.escape(environment)}\}}",
        text,
        flags=re.DOTALL,
    )
    return match.group("body").strip() if match else None


def detect_unsupported(text: str, source: Path, warnings: list[str]) -> None:
    for pattern, label in UNSUPPORTED_PATTERNS.items():
        if re.search(pattern, text, flags=re.IGNORECASE):
            warnings.append(f"{source}: 检测到{label}；Word 转换不保证版式或内容完整。")


def _latex_dimension_fraction(value: str) -> float | None:
    """Convert a controlled LaTeX width to a fraction of the 15.5 cm text area."""

    compact = re.sub(r"\s+", "", value.strip())
    relative = re.fullmatch(
        r"(?P<factor>(?:\d+(?:\.\d*)?|\.\d+)?)"
        r"\\(?:textwidth|linewidth|columnwidth)",
        compact,
    )
    if relative:
        return float(relative.group("factor") or "1")

    absolute = re.fullmatch(
        r"(?P<value>\d+(?:\.\d*)?|\.\d+)"
        r"(?P<unit>cm|mm|in|pt|bp|pc)",
        compact,
        flags=re.IGNORECASE,
    )
    if not absolute:
        return None
    amount = float(absolute.group("value"))
    unit = absolute.group("unit").casefold()
    centimetres = {
        "cm": amount,
        "mm": amount / 10,
        "in": amount * 2.54,
        "pt": amount * 2.54 / 72.27,
        "bp": amount * 2.54 / 72,
        "pc": amount * 12 * 2.54 / 72.27,
    }[unit]
    return centimetres / 15.5


def _latex_length_points(value: str) -> float | None:
    compact = re.sub(r"\s+", "", value.strip())
    match = re.fullmatch(
        r"(?P<value>-?(?:\d+(?:\.\d*)?|\.\d+))"
        r"(?P<unit>pt|bp|pc|cm|mm|in|em)",
        compact,
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    amount = float(match.group("value"))
    return {
        "pt": amount,
        "bp": amount * 72.27 / 72,
        "pc": amount * 12,
        "cm": amount * 72.27 / 2.54,
        "mm": amount * 72.27 / 25.4,
        "in": amount * 72.27,
        # Table text is five-size (10.5 pt), so source em indents remain
        # proportional to the style rather than to Word's Normal style.
        "em": amount * 10.5,
    }[match.group("unit").casefold()]


def _declaration_alignment(value: str) -> str | None:
    if re.search(r"\\centering\b", value):
        return "center"
    if re.search(r"\\raggedleft\b", value):
        return "right"
    if re.search(r"\\raggedright\b", value):
        return "left"
    return None


def _declaration_first_line_indent(value: str) -> float | None:
    match = re.search(
        r"\\setlength\s*\{\s*\\parindent\s*\}\s*\{([^{}]+)\}",
        value,
    )
    return _latex_length_points(match.group(1)) if match else None


def parse_latex_table_columns(specification: str) -> list[LatexTableColumn]:
    """Parse widths and alignment from the table-column subset used by the class."""

    columns: list[LatexTableColumn] = []
    pending_alignment: str | None = None
    pending_first_line_indent: float | None = None
    index = 0
    length = len(specification)

    def skip_space(cursor: int) -> int:
        while cursor < length and specification[cursor].isspace():
            cursor += 1
        return cursor

    while index < length:
        token = specification[index]
        if token in "><@!":
            cursor = skip_space(index + 1)
            if cursor < length and specification[cursor] == "{":
                declaration, index = extract_balanced(specification, cursor)
                if token == ">":
                    pending_alignment = (
                        _declaration_alignment(declaration) or pending_alignment
                    )
                    declared_indent = _declaration_first_line_indent(declaration)
                    if declared_indent is not None:
                        pending_first_line_indent = declared_indent
                continue

        if token == "*":
            cursor = skip_space(index + 1)
            if cursor < length and specification[cursor] == "{":
                repeats_text, cursor = extract_balanced(specification, cursor)
                cursor = skip_space(cursor)
                if cursor < length and specification[cursor] == "{":
                    repeated_spec, index = extract_balanced(specification, cursor)
                    try:
                        repeats = max(0, int(repeats_text.strip()))
                    except ValueError as error:
                        raise BuildError(
                            f"Word 转换不支持非整数表格重复列次数：{repeats_text.strip()}"
                        ) from error
                    repeated_columns = parse_latex_table_columns(repeated_spec)
                    for _ in range(repeats):
                        columns.extend(repeated_columns)
                    pending_alignment = None
                    pending_first_line_indent = None
                    continue

        if token in "LCRpmb":
            cursor = skip_space(index + 1)
            if cursor < length and specification[cursor] == "{":
                width, index = extract_balanced(specification, cursor)
                if token in "LCR":
                    horizontal = {
                        "L": "left",
                        "C": "center",
                        "R": "right",
                    }[token]
                    vertical = "center"
                else:
                    horizontal = pending_alignment or "left"
                    vertical = {"p": "top", "m": "center", "b": "bottom"}[token]
                columns.append(
                    LatexTableColumn(
                        horizontal=pending_alignment or horizontal,
                        vertical=vertical,
                        width_fraction=_latex_dimension_fraction(width),
                        first_line_indent_pt=pending_first_line_indent,
                    )
                )
                pending_alignment = None
                pending_first_line_indent = None
                continue

        if token in "lcrXS":
            if token in "lcr":
                horizontal = {"l": "left", "c": "center", "r": "right"}[token]
            elif token == "S":
                horizontal = "right"
            else:
                horizontal = "left"
            columns.append(
                LatexTableColumn(
                    horizontal=pending_alignment or horizontal,
                    vertical="center",
                    flexible=token == "X",
                    first_line_indent_pt=pending_first_line_indent,
                )
            )
            pending_alignment = None
            pending_first_line_indent = None
            index += 1
            if token == "S":
                index = skip_space(index)
                if index < length and specification[index] == "[":
                    closing = specification.find("]", index + 1)
                    if closing < 0:
                        raise BuildError("S 表格列选项缺少右方括号。")
                    index = closing + 1
            continue

        if token == "\\":
            command = re.match(r"\\[A-Za-z@]+|\\.", specification[index:])
            index += len(command.group(0)) if command else 1
            index = skip_space(index)
            while index < length and specification[index] == "{":
                _, index = extract_balanced(specification, index)
                index = skip_space(index)
            continue
        if token.isalpha():
            raise BuildError(f"Word 转换不支持表格列类型：{token}")
        index += 1
    return columns


def count_latex_table_columns(specification: str) -> int:
    """Count columns in the controlled LaTeX table specification subset."""

    return len(parse_latex_table_columns(specification))


def _table_local_length_points(
    text: str,
    begin_start: int,
    command: str,
) -> float | None:
    r"""Return the nearest table-local ``\setlength`` value before a preamble.

    Table spacing declarations in the supplied sources live either inside a
    ``table`` float or a ``\begingroup`` block immediately surrounding a
    ``longtable``.  Stop at the nearest opening/closing table boundary so a
    value from an earlier table cannot leak into a later one.
    """

    boundary_markers = (
        r"\begingroup",
        r"\endgroup",
        r"\begin{table}",
        r"\begin{table*}",
        r"\end{table}",
        r"\end{table*}",
        r"\end{longtable}",
    )
    scope_start = max(
        (text.rfind(marker, 0, begin_start) for marker in boundary_markers),
        default=-1,
    )
    segment = text[scope_start if scope_start >= 0 else 0 : begin_start]
    matches = list(
        re.finditer(
            rf"\\setlength\s*\{{\s*\\{re.escape(command)}\s*\}}"
            r"\s*\{([^{}]+)\}",
            segment,
        )
    )
    return _latex_length_points(matches[-1].group(1)) if matches else None


def _find_latex_table_preambles(text: str) -> list[_LatexTablePreamble]:
    preambles: list[_LatexTablePreamble] = []
    pattern = re.compile(
        r"\\begin\s*\{(?P<environment>longtable|tabularx|tabular\*|tabular)\}"
    )

    def skip_space(cursor: int) -> int:
        while cursor < len(text) and text[cursor].isspace():
            cursor += 1
        return cursor

    def skip_optional_position(cursor: int) -> int:
        cursor = skip_space(cursor)
        if cursor >= len(text) or text[cursor] != "[":
            return cursor
        closing = text.find("]", cursor + 1)
        return skip_space(closing + 1) if closing >= 0 else len(text)

    for match in pattern.finditer(text):
        environment = match.group("environment")
        cursor = skip_space(match.end())
        table_width_fraction: float | None = None
        try:
            if environment in {"tabularx", "tabular*"}:
                if cursor >= len(text) or text[cursor] != "{":
                    continue
                width, cursor = extract_balanced(text, cursor)
                table_width_fraction = _latex_dimension_fraction(width)
                cursor = skip_optional_position(cursor)
            else:
                cursor = skip_optional_position(cursor)
            cursor = skip_space(cursor)
            if cursor >= len(text) or text[cursor] != "{":
                continue
            specification_start = cursor + 1
            specification, cursor = extract_balanced(text, cursor)
            specification_end = cursor - 1
        except BuildError:
            continue
        preambles.append(
            _LatexTablePreamble(
                environment=environment,
                begin_start=match.start(),
                specification_start=specification_start,
                specification_end=specification_end,
                specification=specification,
                table_width_fraction=(
                    1.0 if environment == "longtable" else table_width_fraction
                ),
                tabcolsep_pt=_table_local_length_points(
                    text,
                    match.start(),
                    "tabcolsep",
                ),
            )
        )
    return preambles


def _table_layout_label(text: str, preamble: _LatexTablePreamble) -> str | None:
    if preamble.environment == "longtable":
        end = text.find(r"\end{longtable}", preamble.specification_end)
        segment = text[
            preamble.begin_start : end if end >= 0 else len(text)
        ]
    else:
        prefix = text[: preamble.begin_start]
        table_start = prefix.rfind(r"\begin{table}")
        if table_start < 0:
            table_start = prefix.rfind(r"\begin{table*}")
        last_table_end = max(prefix.rfind(r"\end{table}"), prefix.rfind(r"\end{table*}"))
        if table_start > last_table_end:
            candidates = [
                position
                for marker in (r"\end{table}", r"\end{table*}")
                if (position := text.find(marker, preamble.specification_end)) >= 0
            ]
            end = min(candidates) if candidates else len(text)
            segment = text[table_start:end]
        else:
            end_marker = rf"\end{{{preamble.environment}}}"
            end = text.find(end_marker, preamble.specification_end)
            segment = text[
                preamble.begin_start : end if end >= 0 else len(text)
            ]
    labels = re.findall(r"\\label\s*\{([^{}]+)\}", segment)
    return next((label for label in labels if label.strip().startswith("tab:")), labels[0] if labels else None)


def extract_latex_table_layouts(text: str) -> list[LatexTableLayout]:
    """Collect semantic table layouts in document order before Pandoc runs."""

    layouts: list[LatexTableLayout] = []
    for preamble in _find_latex_table_preambles(text):
        columns = parse_latex_table_columns(preamble.specification)
        if columns:
            layouts.append(
                LatexTableLayout(
                    environment=preamble.environment,
                    columns=tuple(columns),
                    label=_table_layout_label(text, preamble),
                    table_width_fraction=preamble.table_width_fraction,
                    tabcolsep_pt=preamble.tabcolsep_pt,
                )
            )
    return layouts


def normalize_latex_table_preambles(text: str) -> str:
    """Make only table preambles Pandoc-readable; leave body prose untouched."""

    replacements: list[tuple[int, int, str]] = []
    for preamble in _find_latex_table_preambles(text):
        specification = re.sub(
            r"(?<![A-Za-z@\\])(?P<column>[LCR])\s*\{[^{}]+\}",
            lambda match: {"L": "l", "C": "c", "R": "r"}[
                match.group("column")
            ],
            preamble.specification,
        )
        # Pandoc retains the ``X`` column shell but can discard its cell
        # contents when the column is preceded by an array declaration such
        # as ``>{\centering\arraybackslash}``. The original alignment has
        # already been recorded for the DOCX layout pass, so remove only this
        # Pandoc-incompatible prefix from the conversion copy.
        specification = re.sub(
            r">\s*\{[^{}]*\}\s*(?=X)",
            "",
            specification,
        )
        if specification != preamble.specification:
            replacements.append(
                (
                    preamble.specification_start,
                    preamble.specification_end,
                    specification,
                )
            )
    for start, end, replacement in reversed(replacements):
        text = text[:start] + replacement + text[end:]
    return text


def normalize_longtable_headers(text: str) -> str:
    """Remove page-only header and footer definitions before Pandoc reads longtable.

    Pandoc already turns the first longtable header into a Markdown table
    header.  Keeping the block between ``\\endfirsthead`` and ``\\endhead``
    adds a second header plus a blank row.  Likewise, rows declared before
    ``\\endfoot`` (for example ``续下页``) are page furniture in LaTeX but
    become ordinary Word data rows unless the entire page-definition block is
    removed before conversion.
    """

    replacements: list[tuple[int, int, str]] = []
    for preamble in _find_latex_table_preambles(text):
        if preamble.environment != "longtable":
            continue
        end_match = re.search(
            r"\\end\s*\{longtable\}",
            text[preamble.specification_end + 1 :],
        )
        if end_match is None:
            continue
        end_start = preamble.specification_end + 1 + end_match.start()
        end = preamble.specification_end + 1 + end_match.end()
        body = text[preamble.specification_end + 1 : end_start]
        first_header_end = re.search(r"\\endfirsthead\b", body)
        if first_header_end is not None:
            page_definitions = body[first_header_end.end() :]
            last_footer_end = re.search(r"\\endlastfoot\b", page_definitions)
            ordinary_footer_end = re.search(r"\\endfoot\b", page_definitions)
            repeated_header_end = re.search(r"\\endhead\b", page_definitions)
            definition_end = (
                last_footer_end or ordinary_footer_end or repeated_header_end
            )
            if definition_end is not None:
                body = (
                    body[: first_header_end.start()]
                    + "\n"
                    + page_definitions[definition_end.end() :]
                )
        body = re.sub(r"\\end(?:firsthead|head|foot|lastfoot)\b", "", body)
        columns = parse_latex_table_columns(preamble.specification)
        simplified_spec = "".join(
            {"left": "l", "center": "c", "right": "r"}[column.horizontal]
            for column in columns
        ) or "l"
        replacements.append(
            (
                preamble.begin_start,
                end,
                text[preamble.begin_start : preamble.specification_start]
                + simplified_spec
                + text[preamble.specification_end : preamble.specification_end + 1]
                + body
                + text[end_start:end],
            )
        )
    for start, end, replacement in reversed(replacements):
        text = text[:start] + replacement + text[end:]
    return text


def unwrap_table_resizeboxes(text: str) -> str:
    """Remove ``\\resizebox`` wrappers whose body is a LaTeX table.

    Pandoc does not convert a ``tabular`` environment when it remains nested
    inside the graphics-oriented ``\\resizebox`` command.  PDF output may be
    correct while the corresponding Word table, caption label, and cross-
    references disappear.  Keep non-table resize boxes untouched and unwrap
    only bodies that contain a supported table environment.
    """

    pattern = re.compile(r"\\resizebox\*?\s*\{")
    cursor = 0
    parts: list[str] = []
    while match := pattern.search(text, cursor):
        width_open = text.find("{", match.start())
        _, width_end = extract_balanced(text, width_open)

        height_open = width_end
        while height_open < len(text) and text[height_open].isspace():
            height_open += 1
        if height_open >= len(text) or text[height_open] != "{":
            cursor = match.end()
            continue
        _, height_end = extract_balanced(text, height_open)

        body_open = height_end
        while body_open < len(text) and text[body_open].isspace():
            body_open += 1
        if body_open >= len(text) or text[body_open] != "{":
            cursor = match.end()
            continue
        body, body_end = extract_balanced(text, body_open)

        parts.append(text[cursor : match.start()])
        if re.search(
            r"\\begin\s*\{(?:longtable|tabularx|tabular\*|tabular)\}",
            body,
        ):
            parts.append(body)
        else:
            parts.append(text[match.start() : body_end])
        cursor = body_end
    parts.append(text[cursor:])
    return "".join(parts)


def resolve_word_image_target(target: str, resource_roots: Iterable[Path]) -> str:
    """Prefer a Word-compatible raster sibling for PDF/SVG figure targets."""

    value = html_lib.unescape(target).strip()
    if not value or re.match(r"^(?:https?|data):", value, flags=re.IGNORECASE):
        return value

    relative = Path(value.replace("/", os.sep))
    if relative.suffix.casefold() not in {".pdf", ".svg", ".eps"}:
        return value

    candidates = [relative] if relative.is_absolute() else [root / relative for root in resource_roots]
    for candidate in candidates:
        for suffix in (".png", ".jpg", ".jpeg"):
            raster = candidate.with_suffix(suffix)
            if not raster.is_file():
                continue
            if relative.is_absolute():
                return raster.resolve().as_posix()
            return relative.with_suffix(suffix).as_posix()
    return value


def normalize_latex(
    text: str,
    notes: NoteRegistry,
    discipline: str = "humanities",
    labels: LabelRegistry | None = None,
    citations: CitationRegistry | None = None,
    table_layouts: list[LatexTableLayout] | None = None,
) -> str:
    text = strip_tex_comments(text)
    text = unwrap_table_resizeboxes(text)
    if table_layouts is not None:
        table_layouts.extend(extract_latex_table_layouts(text))
    text = normalize_longtable_headers(text)
    # Pandoc understands native l/c/r columns but discards the alignment on
    # array's paragraph columns.  Width/vertical semantics are retained in the
    # table layout registry and reapplied to Word after conversion.
    text = normalize_latex_table_preambles(text)
    labels = labels if labels is not None else LabelRegistry()
    citations = citations if citations is not None else CitationRegistry()
    chapter_command = "chapter" if discipline == "science" else "chapter*"

    def thematic_chapter(match: re.Match[str], default: str) -> str:
        title = match.group("title") if match.group("title") is not None else default
        return rf"\{chapter_command}{{{title}}}"

    def unnumbered_thematic_chapter(match: re.Match[str], default: str) -> str:
        title = match.group("title") if match.group("title") is not None else default
        return rf"\chapter*{{{title}}}"

    text = re.sub(
        r"\\yibinopeningchapter(?![A-Za-z@])(?:\s*\[(?P<title>[^]]*)\])?",
        lambda match: thematic_chapter(match, "绪论"),
        text,
    )
    text = re.sub(
        r"\\yibinclosingchapter(?![A-Za-z@])(?:\s*\[(?P<title>[^]]*)\])?",
        lambda match: unnumbered_thematic_chapter(match, "结论"),
        text,
    )
    text = replace_braced_command(
        text,
        "yibinunnumberedchapter",
        lambda title: rf"\chapter*{{{title}}}",
    )
    text = re.sub(
        r"\\begin\s*\{yibinappendices\}",
        "",
        text,
    )
    text = re.sub(r"\\end\s*\{yibinappendices\}", "", text)
    appendix_index = 0

    def appendix_heading(title: str) -> str:
        nonlocal appendix_index
        appendix_index += 1
        marker = chr(ord("A") + appendix_index - 1) if appendix_index <= 26 else str(appendix_index)
        return rf"\chapter*{{附录{marker} {title}}}"

    text = replace_braced_command(text, "yibinappendix", appendix_heading)

    appendix_marker: str | None = None

    def appendix_section_heading(match: re.Match[str]) -> str:
        nonlocal appendix_marker
        if match.group("marker") is not None:
            appendix_marker = match.group("marker")
            return match.group(0)
        if appendix_marker is None:
            raise BuildError("\\yibinappendixsection 必须写在 \\yibinappendix 之后。")
        title = match.group("title")
        return rf"\section*{{{title}}}"

    text = re.sub(
        r"\\chapter\*\{附录(?P<marker>[A-Z]|\d+)\s+[^{}]*\}"
        r"|\\yibinappendixsection\s*\{(?P<title>[^{}]*)\}",
        appendix_section_heading,
        text,
    )

    def register_appendix_label(match: re.Match[str]) -> str:
        labels.register(match.group("label"), match.group("marker"), "附录")
        return match.group("heading")

    text = re.sub(
        r"(?P<heading>\\chapter\*\{附录(?P<marker>[A-Z]|\d+)\s+[^{}]*\})"
        r"\s*\\label\s*\{(?P<label>[^{}]+)\}",
        register_appendix_label,
        text,
    )
    text = replace_braced_command(
        text,
        "yibincite",
        citations.placeholder,
    )
    text = replace_braced_command(text, "cite", citations.placeholder)
    text = replace_braced_command(text, "nocite", citations.add_nocite)
    text = replace_braced_command(text, "yibinnote", notes.add)
    # Pandoc resolves each standalone LaTeX fragment with a fresh counter, so
    # its native \ref text is stale for cross-file targets.  Preserve the
    # semantic reference as an opaque token and resolve it after all labels
    # have received their template-specific global number.
    text = replace_braced_command(
        text,
        "eqref",
        lambda label: labels.placeholder(label, parenthesized=True),
    )
    text = replace_braced_command(
        text,
        "ref",
        lambda label: labels.placeholder(label, parenthesized=False),
    )
    text = re.sub(r"\\FloatBarrier\b", "", text)
    # The example's framed rule is a LaTeX-only visual placeholder.  Preserve
    # its semantic presence instead of silently dropping the whole figure.
    text = re.sub(
        r"\\fbox\s*\{\s*\\rule\s*\{[^{}]*\}\s*\{[^{}]*\}\s*"
        r"\\rule\s*\{[^{}]*\}\s*\{[^{}]*\}\s*\}",
        r"\\textit{[插图占位框]}",
        text,
    )
    return text


def run_checked(
    command: list[str],
    *,
    cwd: Path,
    capture_output: bool = False,
) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        command,
        cwd=cwd,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE if capture_output else None,
        stderr=subprocess.PIPE,
        check=False,
    )
    if result.returncode != 0:
        rendered = subprocess.list2cmdline(command)
        detail = (result.stderr or "").strip()
        raise BuildError(f"命令执行失败（{result.returncode}）：{rendered}\n{detail}")
    return result


def collect_graphic_paths(
    main_text: str,
    *,
    main_dir: Path,
    project_root: Path,
    allow_project_fallback: bool = True,
) -> list[Path]:
    """Return existing directories declared through ``\\graphicspath``.

    LaTeX resolves graphics against its process/search path, while the
    Pandoc DOCX writer resolves image targets through ``--resource-path``.
    Entry-file-relative directories take priority. Project-root-relative
    fallback is retained only for entry files that live inside the template
    repository, where bundled examples intentionally use root-relative paths.
    """
    paths: list[Path] = []
    cleaned = strip_tex_comments(main_text)
    for outer in re.finditer(
        r"\\graphicspath\s*\{(?P<body>(?:\s*\{[^{}]*\}\s*)+)\}",
        cleaned,
        flags=re.DOTALL,
    ):
        for item in re.findall(r"\{([^{}]+)\}", outer.group("body")):
            raw = item.strip().replace("/", os.sep)
            if not raw:
                continue
            candidate = Path(raw)
            if candidate.is_absolute():
                options = [candidate]
            else:
                options = [main_dir / candidate]
                if allow_project_fallback and main_dir.resolve() != project_root.resolve():
                    options.append(project_root / candidate)
            for option in options:
                resolved = option.resolve()
                if resolved.is_dir() and resolved not in paths:
                    paths.append(resolved)
    return paths


def latex_to_markdown(
    text: str,
    *,
    pandoc: Path,
    cwd: Path,
    temp_dir: Path,
    name: str,
) -> str:
    source = temp_dir / f"{name}.tex"
    source.write_text(text, encoding="utf-8", newline="\n")
    command = [
        str(pandoc),
        "--from=latex-smart",
        "--to=markdown+raw_attribute+fenced_divs+bracketed_spans+tex_math_dollars",
        "--wrap=none",
    ]
    if FIGURE_FILTER.is_file():
        command.append(f"--lua-filter={FIGURE_FILTER}")
    command.append(str(source))
    result = run_checked(
        command,
        cwd=cwd,
        capture_output=True,
    )
    markdown = (result.stdout or "").strip()

    def restore_chinese_quotes(match: re.Match[str]) -> str:
        content = match.group(1)
        if re.search(r"[\u3400-\u9fff]", content):
            return f"“{content}”"
        return match.group(0)

    return re.sub(r'"([^"\n]+)"', restore_chinese_quotes, markdown)


def chinese_number(value: int) -> str:
    digits = "零一二三四五六七八九"
    if value < 10:
        return digits[value]
    if value == 10:
        return "十"
    if value < 20:
        return "十" + digits[value % 10]
    if value < 100:
        suffix = "" if value % 10 == 0 else digits[value % 10]
        return digits[value // 10] + "十" + suffix
    return str(value)


def apply_heading_numbering(
    markdown: str,
    discipline: str,
    counters: HeadingCounters,
    labels: LabelRegistry | None = None,
) -> str:
    labels = labels or LabelRegistry()
    output: list[str] = []
    heading = re.compile(r"^(#{1,4})\s+(.+?)\s*$")
    attrs = re.compile(r"^(?P<title>.*?)(?P<attrs>\s+\{[^{}]*\})?$")
    for line in markdown.splitlines():
        match = heading.match(line)
        if not match:
            output.append(line)
            continue
        level = len(match.group(1))
        parsed = attrs.match(match.group(2))
        assert parsed is not None
        title = parsed.group("title").strip()
        attribute_text = parsed.group("attrs") or ""
        unnumbered = ".unnumbered" in attribute_text or re.search(r"(?:^|\s)-(?:\s|})", attribute_text)

        if unnumbered:
            if level == 1:
                counters.section = counters.subsection = counters.subsubsection = 0
            output.append(f"{'#' * level} {title}{attribute_text}")
            continue

        if level == 1:
            counters.chapter += 1
            counters.section = counters.subsection = counters.subsubsection = 0
        elif level == 2:
            counters.section += 1
            counters.subsection = counters.subsubsection = 0
        elif level == 3:
            counters.subsection += 1
            counters.subsubsection = 0
        else:
            counters.subsubsection += 1

        if discipline == "science":
            numbers = [
                counters.chapter,
                counters.section,
                counters.subsection,
                counters.subsubsection,
            ][:level]
            prefix = ".".join(str(number) for number in numbers)
            reference_number = prefix
        elif level == 1:
            reference_number = f"{chinese_number(counters.chapter)}、"
        elif level == 2:
            reference_number = f"（{chinese_number(counters.section)}）"
        elif level == 3:
            reference_number = f"{counters.subsection}."
        else:
            reference_number = f"（{counters.subsubsection}）"
        identifier = re.search(r"(?:^|[\s{])#([^\s}]+)", attribute_text)
        if identifier:
            labels.register(identifier.group(1), reference_number, "标题")
        if level == 1:
            output.append(f"<!-- YIBIN_INTERNAL_CHAPTER_{counters.chapter} -->")
        # Numbering is attached to the reusable Yibin heading styles during
        # OOXML post-processing.  Keep the paragraph text itself unnumbered so
        # Word users can renumber, reorder, and cross-reference headings.
        output.append(f"{'#' * level} {title}{attribute_text}")
    return "\n".join(output)


def demote_proposal_headings(markdown: str) -> str:
    """Map standalone Pandoc section levels back to the book-class hierarchy."""

    heading = re.compile(r"^(?P<marks>#{1,3})\s+(?P<body>.+?)\s*$")
    attributes = re.compile(r"^(?P<title>.*?)(?:\s+\{(?P<attrs>[^{}]*)\})?$")
    output: list[str] = []
    for line in markdown.splitlines():
        match = heading.match(line)
        if not match:
            output.append(line)
            continue
        parsed = attributes.match(match.group("body"))
        assert parsed is not None
        attrs = parsed.group("attrs") or ""
        if ".unnumbered" in attrs.split() or re.search(r"(?:^|\s)-(?:\s|$)", attrs):
            if output and output[-1].strip():
                output.append("")
            output.extend(
                custom_block(
                    STYLE_PROPOSAL_UNNUMBERED_HEADING,
                    parsed.group("title").strip(),
                ).splitlines()
            )
            output.append("")
            continue
        if output and output[-1].strip():
            output.append("")
        output.append("#" * (len(match.group("marks")) + 1) + " " + match.group("body"))
        output.append("")
    return "\n".join(output)


def apply_float_numbering(
    markdown: str,
    discipline: str,
    chapter: int,
    counters: FloatCounters,
    labels: LabelRegistry | None = None,
    image_target_resolver: Callable[[str], str] | None = None,
) -> str:
    """Add global caption/equation numbers and register their LaTeX labels.

    ``chapter`` is the chapter active before this Markdown fragment.  Internal
    markers emitted by :func:`apply_heading_numbering` let one source file
    contain multiple chapters without assigning every float to the final one.
    """

    labels = labels or LabelRegistry()

    def markdown_identifier(attributes: str | None) -> str | None:
        if not attributes:
            return None
        attributes = attributes.strip().strip("{}")
        match = re.search(r"(?:^|\s)#([^\s}]+)", attributes)
        return match.group(1) if match else None

    def process_chunk(chunk: str, current_chapter: int) -> str:
        counters.synchronize(current_chapter, discipline)

        def next_number(kind: str) -> str:
            value = getattr(counters, kind) + 1
            setattr(counters, kind, value)
            if kind == "equation" or discipline == "science":
                return f"{current_chapter}.{value}"
            return str(value)

        # Pandoc wraps labelled tables in a fenced div and emits the caption as
        # an indented ``: caption`` line.  Track the surrounding div id while
        # numbering so the label and caption receive the same global number.
        table_lines: list[str] = []
        div_stack: list[str | None] = []
        for line in chunk.splitlines():
            opening = re.match(r"^\s*:::\s+\{(?P<attrs>[^{}]*)\}\s*$", line)
            closing = re.match(r"^\s*:::\s*$", line)
            if opening:
                div_stack.append(markdown_identifier(opening.group("attrs")))
                table_lines.append(line)
                continue
            if closing:
                if div_stack:
                    div_stack.pop()
                table_lines.append(line)
                continue
            caption = re.match(
                r"^(?P<indent>[ \t]*):\s+(?P<caption>[^\n]+)$",
                line,
            )
            if caption:
                number = next_number("table")
                identifier = next(
                    (item for item in reversed(div_stack) if item),
                    None,
                )
                if identifier:
                    labels.register(identifier, number, "表")
                line = (
                    f"{caption.group('indent')}: 表{number} "
                    f"{caption.group('caption').strip()}"
                )
            table_lines.append(line)
        chunk = "\n".join(table_lines)

        def markdown_figure(match: re.Match[str]) -> str:
            number = next_number("figure")
            identifier = markdown_identifier(match.group("attrs"))
            if identifier:
                labels.register(identifier, number, "图")
            return (
                f"![图{number} {match.group('caption').strip()}]"
                f"({match.group('target')}){match.group('attrs') or ''}"
            )

        chunk = re.sub(
            r"(?m)^!\[(?P<caption>[^\]\n]*)\]"
            r"\((?P<target><[^>]+>|[^)\n]+)\)"
            r"(?P<attrs>\{[^{}]*\})?[ \t]*$",
            markdown_figure,
            chunk,
        )

        def image(match: re.Match[str]) -> str:
            attributes = match.group("attributes")
            src_match = re.search(
                r'\bsrc="([^"]+)"',
                attributes,
                flags=re.IGNORECASE,
            )
            if not src_match:
                return match.group(0)
            src = html_lib.unescape(src_match.group(1))
            if image_target_resolver is not None:
                src = image_target_resolver(src)
            alt_match = re.search(
                r'\balt="([^"]*)"',
                attributes,
                flags=re.IGNORECASE,
            )
            alt = html_lib.unescape(alt_match.group(1)) if alt_match else ""
            style_match = re.search(
                r'\bstyle="([^"]*)"',
                attributes,
                flags=re.IGNORECASE,
            )
            options: list[str] = []
            if style_match:
                style = style_match.group(1)
                for dimension in ("width", "height"):
                    dimension_value = re.search(
                        rf"(?:^|;)\s*{dimension}\s*:\s*([^;]+)",
                        style,
                        flags=re.IGNORECASE,
                    )
                    if dimension_value:
                        options.append(f"{dimension}={dimension_value.group(1).strip()}")
            target_value = f"<{src}>" if re.search(r"[\s()]", src) else src
            attributes_md = "{" + " ".join(options) + "}" if options else ""
            return f"![{alt}]({target_value}){attributes_md}"

        def figure(match: re.Match[str]) -> str:
            number = next_number("figure")
            attrs = match.group("attrs")
            identifier = re.search(r'\bid="([^"]+)"', attrs)
            if identifier:
                labels.register(identifier.group(1), number, "图")
            caption = re.sub(r"<[^>]+>", "", match.group("caption")).strip()
            body = match.group("body").strip()
            body = re.sub(r"^<p>(.*?)</p>$", r"\1", body, flags=re.DOTALL)

            body = re.sub(
                r"<(?:img|embed)\s+(?P<attributes>[^>]*?)/?>",
                image,
                body,
                flags=re.IGNORECASE,
            )
            opening = f"::: {{#{identifier.group(1)}}}" if identifier else "::: {}"
            return (
                f"{opening}\n{body}\n:::\n\n"
                + custom_block("Caption", f"图{number} {caption}")
            )

        chunk = re.sub(
            r"<figure(?P<attrs>[^>]*)>\s*(?P<body>.*?)\s*"
            r"<figcaption>(?P<caption>.*?)</figcaption>\s*</figure>",
            figure,
            chunk,
            flags=re.DOTALL | re.IGNORECASE,
        )

        chunk = re.sub(
            r"<(?:img|embed)\s+(?P<attributes>[^>]*?)/?>",
            image,
            chunk,
            flags=re.IGNORECASE,
        )

        def equation(match: re.Match[str]) -> str:
            number = next_number("equation")
            body = match.group("body")
            equation_labels = re.findall(r"\\label\s*\{([^{}]+)\}", body)
            for label in equation_labels:
                labels.register(label, number, "公式")
            body = re.sub(r"\\label\s*\{[^{}]+\}", "", body).strip()
            marker = custom_block(
                "YibinSectionMarker",
                f"YIBIN_INTERNAL_EQUATION_{number.replace('.', '_')}",
            )
            # Keep a blank line after the fenced marker.  Without it, the
            # following prose can be parsed into the same Word paragraph as
            # the marker, preventing equation-number post-processing.
            return f"$$\n{body}\n$$\n\n{marker}\n\n"

        return re.sub(
            r"\$\$\s*\\begin\s*\{equation\}(?P<body>.*?)"
            r"\\end\s*\{equation\}\s*\$\$",
            equation,
            chunk,
            flags=re.DOTALL,
        )

    marker = re.compile(r"(?m)^<!-- YIBIN_INTERNAL_CHAPTER_(\d+) -->\s*$")
    output: list[str] = []
    cursor = 0
    current_chapter = chapter
    for match in marker.finditer(markdown):
        output.append(process_chunk(markdown[cursor : match.start()], current_chapter))
        current_chapter = int(match.group(1))
        cursor = match.end()
    output.append(process_chunk(markdown[cursor:], current_chapter))
    return "".join(output)


def prepare_bibliography(
    bibliography: Path,
    *,
    pandoc: Path,
    cwd: Path,
    temp_dir: Path,
) -> Path:
    """Convert BibLaTeX to CSL JSON and preserve @standard as CSL standard.

    Pandoc maps BibLaTeX ``@standard`` to ``legislation`` by default, which the
    bundled GB/T style correctly renders as archival material [A].  The source
    semantics are unambiguous here, so restore only those entries to the CSL
    ``standard`` type before citeproc runs.
    """
    if bibliography.suffix.lower() not in {".bib", ".biblatex"}:
        return bibliography
    result = run_checked(
        [str(pandoc), "--from=biblatex", "--to=csljson", str(bibliography)],
        cwd=cwd,
        capture_output=True,
    )
    try:
        items = json.loads(result.stdout or "[]")
    except json.JSONDecodeError as error:
        raise BuildError(f"Pandoc 无法解析参考文献 JSON：{error}") from error
    source = strip_tex_comments(bibliography.read_text(encoding="utf-8"))
    entry_types = {
        key.strip(): entry_type.lower()
        for entry_type, key in re.findall(
            r"@([A-Za-z]+)\s*\{\s*([^,\s]+)",
            source,
        )
    }
    for item in items:
        if entry_types.get(str(item.get("id", ""))) == "standard":
            item["type"] = "standard"
    output = temp_dir / "references.csl.json"
    output.write_text(
        json.dumps(items, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    return output


def custom_block(style: str, content: str) -> str:
    return f'::: {{custom-style="{style}"}}\n{content.strip()}\n:::'


def styled_marker(value: str) -> str:
    return custom_block("YibinSectionMarker", value)


def escape_markdown(value: str) -> str:
    for char in ("\\", "[", "]", "*", "_"):
        value = value.replace(char, "\\" + char)
    return value


def abstract_markdown(
    source_text: str,
    *,
    environment: str,
    pandoc: Path,
    cwd: Path,
    temp_dir: Path,
    notes: NoteRegistry,
    discipline: str,
    labels: LabelRegistry,
    citations: CitationRegistry,
    table_layouts: list[LatexTableLayout],
) -> str:
    body = extract_environment(source_text, environment)
    if body is None:
        raise BuildError(f"摘要文件缺少 {environment} 环境。")
    keyword_command = "cnkeywords" if environment == "cnabstract" else "enkeywords"
    keywords = find_command_argument(body, keyword_command) or ""
    body = replace_braced_command(body, keyword_command, lambda _: "")
    body_md = latex_to_markdown(
        normalize_latex(
            body,
            notes,
            discipline,
            labels,
            citations,
            table_layouts,
        ),
        pandoc=pandoc,
        cwd=cwd,
        temp_dir=temp_dir,
        name=environment,
    )
    if environment == "cnabstract":
        heading = custom_block("FrontTitle", "摘要")
        body_block = custom_block("ChineseAbstract", body_md)
        label = "关键词："
        separator = ""
    else:
        heading = custom_block("FrontTitle", "Abstract")
        body_block = custom_block("EnglishAbstract", body_md)
        label = "Keywords:"
        separator = " "
    keyword_line = (
        f'[{label}]{{custom-style="KeywordLabel"}}{separator}'
        f'{escape_markdown(clean_tex_scalar(keywords))}'
    )
    return "\n\n".join([heading, body_block, custom_block("Keywords", keyword_line)])


def bibliography_markdown() -> str:
    return "# 参考文献 {.unnumbered}\n\n::: {#refs .references}\n:::"


def citation_seed_markdown(citations: CitationRegistry) -> str:
    """Feed citeproc every cited key without leaving a visible citation.

    Inline citations are replaced by opaque markers so they can become Word
    fields later.  Citeproc still needs semantic citation nodes to construct
    the bibliography; a removable seed paragraph supplies them in first-use
    order.
    """
    if not citations.order:
        return ""
    items = "; ".join(f"@{key}" for key in citations.order)
    return f"YIBIN_INTERNAL_CITATION_SEED [{items}]"


def notes_markdown(notes: NoteRegistry) -> str:
    if not notes.entries:
        return ""
    lines = ["# 注释 {.unnumbered}"]
    for index, entry in enumerate(notes.entries, start=1):
        marker = chr(0x245F + index) if index <= 20 else f"[{index}]"
        lines.append(custom_block("Notes", f"{marker} {entry}"))
    return "\n\n".join(lines)


def find_pandoc(explicit: str | None, project_root: Path) -> Path:
    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit).expanduser())
    if os.environ.get("PANDOC"):
        candidates.append(Path(os.environ["PANDOC"]).expanduser())
    candidates.extend(
        [
            project_root / ".tools" / "pandoc-3.9.0.2" / "pandoc.exe",
            project_root / ".tools" / "pandoc" / "pandoc.exe",
            project_root / "tools" / "pandoc.exe",
            project_root / "tools" / "pandoc" / "pandoc.exe",
            project_root / "vendor" / "pandoc" / "pandoc.exe",
        ]
    )
    on_path = shutil.which("pandoc")
    if on_path:
        candidates.append(Path(on_path))
    for candidate in candidates:
        candidate = candidate.resolve()
        if candidate.is_file():
            return candidate
    raise BuildError(
        "未找到 Pandoc。请安装 Pandoc 并加入 PATH，或使用 --pandoc 指定可执行文件。"
    )


def _set_page_number_format(section, format_name: str | None, start: int | None = None) -> None:
    sect_pr = section._sectPr
    existing = sect_pr.find(qn("w:pgNumType"))
    if format_name is None:
        if existing is not None:
            sect_pr.remove(existing)
        return
    if existing is None:
        existing = OxmlElement("w:pgNumType")
        sect_pr.append(existing)
    existing.set(qn("w:fmt"), format_name)
    if start is not None:
        existing.set(qn("w:start"), str(start))


def _strip_section_links(sect_pr) -> None:
    for child in list(sect_pr):
        if child.tag in {
            qn("w:headerReference"),
            qn("w:footerReference"),
            qn("w:pgNumType"),
            qn("w:titlePg"),
            qn("w:type"),
        }:
            sect_pr.remove(child)


def _add_section_break(paragraph, base_sect_pr, page_format: str | None) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    for existing in p_pr.findall(qn("w:sectPr")):
        p_pr.remove(existing)
    sect_pr = copy.deepcopy(base_sect_pr)
    _strip_section_links(sect_pr)
    section_type = OxmlElement("w:type")
    section_type.set(qn("w:val"), "nextPage")
    sect_pr.insert(0, section_type)
    if page_format:
        page_num = OxmlElement("w:pgNumType")
        page_num.set(qn("w:fmt"), page_format)
        page_num.set(qn("w:start"), "1")
        sect_pr.append(page_num)
    p_pr.append(sect_pr)
    for child in list(paragraph._p):
        if child is not p_pr:
            paragraph._p.remove(child)


def _insert_field_after(paragraph, instruction: str) -> None:
    toc_paragraph = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    p_style = OxmlElement("w:pStyle")
    p_style.set(qn("w:val"), "Normal")
    p_pr.append(p_style)
    toc_paragraph.append(p_pr)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    field.set(qn("w:dirty"), "true")
    toc_paragraph.append(field)
    paragraph._p.addnext(toc_paragraph)


def _insert_toc_after(paragraph) -> None:
    _insert_field_after(paragraph, 'TOC \\o "1-3" \\h \\z \\u')


def _insert_caption_list_after(paragraph, sequence_name: str) -> None:
    _insert_field_after(paragraph, f'TOC \\h \\z \\c "{sequence_name}"')


def _format_equation_number(
    paragraph,
    number: str,
    labels: LabelRegistry | None = None,
    *,
    discipline: str = "science",
    equation_sequence: str = "Equation",
) -> None:
    """Center an OMML equation and add a native Word ``SEQ 公式`` number.

    ``equation_sequence`` is detected from Word's built-in CaptionLabels when
    the PowerShell build entry is used.  The visible form remains ``(2.1)``.
    """
    if paragraph._p.find(qn("m:oMathPara")) is None and paragraph._p.find(qn("m:oMath")) is None:
        raise BuildError(f"公式编号 {number} 前未找到 Word 公式对象。")
    parsed_number = re.fullmatch(r"(?P<chapter>\d+)\.(?P<sequence>\d+)", number)
    if parsed_number is None:
        raise BuildError(f"Word 公式编号必须为章号.章内序号：{number}")
    chapter_number = parsed_number.group("chapter")
    sequence_number = parsed_number.group("sequence")
    p_pr = paragraph._p.get_or_add_pPr()
    # Tab positions, spacing, alignment, and fonts live in the reusable
    # 宜宾论文-公式 paragraph style.  Keep only the two tab characters and the
    # fields as paragraph content.
    for name in ("tabs", "jc", "ind", "spacing"):
        node = p_pr.find(qn(f"w:{name}"))
        if node is not None:
            p_pr.remove(node)

    leading_run = OxmlElement("w:r")
    leading_run.append(OxmlElement("w:tab"))
    paragraph._p.insert(1 if paragraph._p[0] is p_pr else 0, leading_run)
    trailing_tab = OxmlElement("w:r")
    trailing_tab.append(OxmlElement("w:tab"))
    paragraph._p.append(trailing_tab)

    matching_labels = [
        label
        for label, target in (labels.targets.items() if labels else [])
        if target.kind == "公式" and target.number == number
    ]
    document = paragraph.part.document
    bookmark_id = _next_bookmark_id(document)
    full_pairs = []
    number_pairs = []
    bookmark_seeds: list[str | None] = matching_labels or [None]
    for label in bookmark_seeds:
        seed = label or f"anonymous:equation:{number}:{bookmark_id}"
        full_start = OxmlElement("w:bookmarkStart")
        full_start.set(qn("w:id"), str(bookmark_id))
        full_start.set(
            qn("w:name"),
            _label_bookmark_name(label)
            if label
            else _hidden_ref_bookmark_name(f"full:{seed}"),
        )
        full_end = OxmlElement("w:bookmarkEnd")
        full_end.set(qn("w:id"), str(bookmark_id))
        bookmark_id += 1
        number_start = OxmlElement("w:bookmarkStart")
        number_start.set(qn("w:id"), str(bookmark_id))
        number_start.set(
            qn("w:name"),
            _label_number_bookmark_name(label)
            if label
            else _hidden_ref_bookmark_name(
                f"number:{seed}",
                prefix="_RefNum",
            ),
        )
        number_end = OxmlElement("w:bookmarkEnd")
        number_end.set(qn("w:id"), str(bookmark_id))
        bookmark_id += 1
        full_pairs.append((full_start, full_end))
        number_pairs.append((number_start, number_end))

    for start, _ in full_pairs:
        paragraph._p.append(start)
    open_run = OxmlElement("w:r")
    open_text = OxmlElement("w:t")
    open_text.text = "("
    open_run.append(open_text)
    paragraph._p.append(open_run)
    for start, _ in number_pairs:
        paragraph._p.append(start)

    if discipline == "science":
        # The science profile's thesis Heading 1 style owns an Arabic native
        # multilevel number, so STYLEREF keeps the chapter prefix synchronized.
        for node in _complex_field_runs(
            f' STYLEREF "{STYLE_HEADING_1}" \\n ',
            chapter_number,
        ):
            paragraph._p.append(node)
    else:
        # Humanities headings use Chinese counters (一、/（一）), while the
        # school equation form remains Arabic (2.1).  The Arabic chapter value
        # therefore remains source-derived; the equation sequence is native.
        chapter_run = _plain_run(chapter_number)
        paragraph._p.append(chapter_run)

    separator_run = _plain_run(".")
    paragraph._p.append(separator_run)
    equation_instruction = f" SEQ {_word_field_identifier(equation_sequence)} "
    if sequence_number == "1":
        equation_instruction += "\\r 1 "
    equation_instruction += "\\* ARABIC "
    for node in _complex_field_runs(equation_instruction, sequence_number):
        paragraph._p.append(node)
    for _, end in reversed(number_pairs):
        paragraph._p.append(end)
    close_run = OxmlElement("w:r")
    close_text = OxmlElement("w:t")
    close_text.text = ")"
    close_run.append(close_text)
    paragraph._p.append(close_run)
    for _, end in reversed(full_pairs):
        paragraph._p.append(end)


def _clear_story(story) -> None:
    paragraphs = list(story.paragraphs)
    if not paragraphs:
        story.add_paragraph()
        return
    first = paragraphs[0]
    first.clear()
    for paragraph in paragraphs[1:]:
        paragraph._element.getparent().remove(paragraph._element)


def _add_page_field(paragraph) -> None:
    paragraph.clear()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph._p.append(_field_run(" PAGE ", "1"))


def _add_header_border(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    borders = p_pr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        p_pr.append(borders)
    bottom = borders.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        borders.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")


def _set_run_fonts(
    run,
    east_asia: str,
    latin: str,
    size: float,
    *,
    bold: bool | None = None,
    underline: bool | None = None,
) -> None:
    run.font.name = latin
    run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if underline is not None:
        run.font.underline = underline
    r_pr = run._r.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), latin)
    r_fonts.set(qn("w:hAnsi"), latin)
    r_fonts.set(qn("w:cs"), latin)
    r_fonts.set(qn("w:eastAsia"), east_asia)
    color = r_pr.find(qn("w:color"))
    if color is None:
        color = OxmlElement("w:color")
        r_pr.append(color)
    color.set(qn("w:val"), "000000")
    for attribute in ("w:themeColor", "w:themeTint", "w:themeShade"):
        color.attrib.pop(qn(attribute), None)


def _png_has_transparency(data: bytes) -> bool:
    """Return whether a PNG uses an alpha channel or palette transparency."""

    if len(data) < 33 or data[:8] != b"\x89PNG\r\n\x1a\n":
        return False
    color_type = data[25]
    if color_type in {4, 6}:
        return True
    if color_type != 3:
        return False
    offset = 8
    while offset + 12 <= len(data):
        length = int.from_bytes(data[offset : offset + 4], "big")
        chunk_type = data[offset + 4 : offset + 8]
        if chunk_type == b"tRNS":
            return True
        offset += 12 + length
        if chunk_type == b"IEND":
            break
    return False


def _flatten_transparent_png(data: bytes, media_name: str) -> bytes:
    try:
        from PIL import Image
    except ImportError as error:
        raise BuildError(
            "检测到带透明通道的 PNG（"
            + media_name
            + "），Word 兼容转换需要 Pillow；请运行 python -m pip install Pillow。"
        ) from error

    with Image.open(io.BytesIO(data)) as source:
        source.load()
        rgba = source.convert("RGBA")
        white = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        white.alpha_composite(rgba)
        rgb = white.convert("RGB")
        output = io.BytesIO()
        rgb.save(output, format="PNG", compress_level=6)
        return output.getvalue()


def whiten_signature_image(source_path: Path, output_path: Path | None = None) -> io.BytesIO:
    """Map a photographed signature's paper background to white without touching the source."""

    try:
        from PIL import Image
    except ImportError as error:
        raise BuildError("签名背景漂白需要 Pillow；请安装 requirements-word.txt。") from error

    with Image.open(source_path) as source:
        source.load()
        grayscale = source.convert("L")
        histogram = grayscale.histogram()
        pixel_count = sum(histogram)

        def percentile(fraction: float) -> int:
            threshold = pixel_count * fraction
            cumulative = 0
            for value, count in enumerate(histogram):
                cumulative += count
                if cumulative >= threshold:
                    return value
            return 255

        black_point = percentile(0.02)
        white_point = percentile(0.20)
        if white_point - black_point < 20:
            whitened = Image.new("RGB", grayscale.size, (255, 255, 255))
        else:
            scale = 255.0 / (white_point - black_point)
            lookup = [
                0
                if value <= black_point
                else 255
                if value >= white_point
                else round((value - black_point) * scale)
                for value in range(256)
            ]
            whitened = grayscale.point(lookup).convert("RGB")
        stream = io.BytesIO()
        whitened.save(stream, format="PNG", compress_level=6, dpi=(300, 300))
        stream.seek(0)
        if output_path is not None:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_bytes(stream.getvalue())
            stream.seek(0)
        return stream


def _normalize_embedded_pngs(docx_path: Path) -> int:
    """Flatten transparent embedded PNG media without changing relationships."""

    replacements: dict[str, bytes] = {}
    with zipfile.ZipFile(docx_path, "r") as package:
        for item in package.infolist():
            media_name = item.filename
            if not media_name.startswith("word/media/") or not media_name.lower().endswith(".png"):
                continue
            data = package.read(item)
            if _png_has_transparency(data):
                replacements[media_name] = _flatten_transparent_png(data, media_name)
    if not replacements:
        return 0

    temporary = docx_path.with_name(docx_path.stem + ".rgb-normalized.docx")
    with zipfile.ZipFile(docx_path, "r") as source, zipfile.ZipFile(temporary, "w") as target:
        for item in source.infolist():
            data = replacements.get(item.filename, source.read(item))
            target.writestr(item, data)
    os.replace(temporary, docx_path)
    return len(replacements)



def _bind_oxml_module():
    import word_oxml
    module = word_oxml.bind_core(sys.modules[__name__])
    for name in dir(module):
        if name.startswith("_"):
            globals()[name] = getattr(module, name)
    return module


_bind_oxml_module()

def _apply_common_body_formatting(
    document: Document,
    *,
    discipline: str,
    labels: LabelRegistry,
    citations: CitationRegistry,
    citation_mode: str,
    table_layouts: list[LatexTableLayout],
    word_figure_sequence: str,
    word_table_sequence: str,
) -> int:
    _require_independent_heading_styles(document)
    reusable_style_map = {
        "Normal": STYLE_BODY,
        "Body Text": STYLE_BODY,
        "Compact": STYLE_BODY,
        "First Paragraph": STYLE_FIRST_PARAGRAPH,
        "Heading 1": STYLE_HEADING_1,
        "Heading 2": STYLE_HEADING_2,
        "Heading 3": STYLE_HEADING_3,
        "Heading 4": STYLE_HEADING_4,
        "Yibin Heading 1": STYLE_HEADING_1,
        "Yibin Heading 2": STYLE_HEADING_2,
        "Yibin Heading 3": STYLE_HEADING_3,
        "Yibin Heading 4": STYLE_HEADING_4,
        "ChineseAbstract": STYLE_CHINESE_ABSTRACT,
        "EnglishAbstract": STYLE_ENGLISH_ABSTRACT,
        "Keywords": STYLE_KEYWORDS,
        "Bibliography": STYLE_BIBLIOGRAPHY,
        "Notes": STYLE_NOTES,
    }
    for paragraph in document.paragraphs:
        source_style = paragraph.style.name
        target_style = reusable_style_map.get(source_style)
        if target_style:
            paragraph.style = document.styles[target_style]
        if source_style == "FrontTitle":
            paragraph.style = document.styles[
                STYLE_ENGLISH_ABSTRACT_TITLE
                if paragraph.text.strip() == "Abstract"
                else STYLE_FRONT_TITLE
            ]
        for run in paragraph.runs:
            if paragraph.style.name in reusable_style_map.values():
                _clear_direct_run_typography(run)

    _apply_list_paragraph_styles(document)

    heading_one = document.styles[STYLE_HEADING_1]
    if discipline == "science":
        heading_one.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading_one.paragraph_format.first_line_indent = Pt(0)
        heading_one.paragraph_format.page_break_before = True
        first_line = "0"
        first_chars = "0"
    else:
        heading_one.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        heading_one.paragraph_format.first_line_indent = Pt(24)
        heading_one.paragraph_format.page_break_before = True
        first_line = "480"
        first_chars = "200"
    heading_ppr = heading_one.element.get_or_add_pPr()
    heading_indent = heading_ppr.find(qn("w:ind"))
    if heading_indent is None:
        heading_indent = OxmlElement("w:ind")
        heading_ppr.append(heading_indent)
    heading_indent.set(qn("w:firstLine"), first_line)
    heading_indent.set(qn("w:firstLineChars"), first_chars)

    unnumbered_titles = {"结论", "注释", "参考文献", "附录", "致谢"}
    if discipline == "humanities":
        unnumbered_titles.add("绪论")
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if paragraph.style.name == STYLE_HEADING_1 and text in unnumbered_titles:
            paragraph.style = document.styles[STYLE_UNNUMBERED_HEADING]
            _clear_paragraph_numbering(paragraph)
        if re.match(r"^附录[A-ZＡ-Ｚ0-9一二三四五六七八九十]+(?:\s|　)", text):
            paragraph.style = document.styles[STYLE_APPENDIX_HEADING]

    heading_abstract_id = _configure_heading_numbering(document, discipline)
    if discipline == "humanities":
        _isolate_humanities_introduction_numbering(
            document,
            heading_abstract_id,
        )
    _configure_appendix_section_numbering(document, discipline)
    title_styles = {
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
    }
    for paragraph in document.paragraphs:
        if paragraph.style.name not in title_styles:
            continue
        _clear_direct_paragraph_format(
            paragraph,
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "spacing",
            "ind",
            "jc",
        )
        for run in paragraph.runs:
            _clear_direct_run_typography(run)

    _format_tables(document, table_layouts, labels)
    _replace_pandoc_checkbox_math(document)
    _clamp_images(document)
    _format_images_and_captions(
        document,
        labels,
        discipline,
        figure_sequence=word_figure_sequence,
        table_sequence=word_table_sequence,
    )
    _replace_reference_markers(document, labels, citations, citation_mode)
    return heading_abstract_id


def postprocess_docx(
    input_docx: Path,
    output_docx: Path,
    *,
    metadata: dict[str, str],
    project_root: Path,
    main_dir: Path,
    metadata_dir: Path,
    discipline: str,
    allow_project_fallback: bool,
    include_cover: bool,
    include_declarations: bool,
    citation_mode: str = "linked",
    labels: LabelRegistry | None = None,
    citations: CitationRegistry | None = None,
    bibliography_items: list[dict[str, object]] | None = None,
    table_layouts: list[LatexTableLayout] | None = None,
    word_figure_sequence: str = "图",
    word_table_sequence: str = "表",
    word_equation_sequence: str = "公式",
) -> None:
    document = Document(input_docx)
    _require_independent_heading_styles(document)
    labels = labels or LabelRegistry()
    citations = citations or CitationRegistry()
    bibliography_items = bibliography_items or []
    table_layouts = table_layouts or []
    if "CoverValue" in [style.name for style in document.styles]:
        document.styles["CoverValue"].font.underline = False
    title = metadata.get("title", "")
    english_title = metadata.get("english-title", "")
    author = metadata.get("author", "")
    body_sect_pr = document._element.body.sectPr
    if body_sect_pr is None:
        raise BuildError("Pandoc 输出缺少正文分节属性。")

    for paragraph in list(document.paragraphs):
        if paragraph.text.strip().startswith("YIBIN_INTERNAL_CITATION_SEED"):
            paragraph._p.getparent().remove(paragraph._p)

    paragraphs = list(document.paragraphs)
    for index, paragraph in enumerate(paragraphs):
        match = re.fullmatch(
            r"YIBIN_INTERNAL_EQUATION_(\d+)_(\d+)",
            paragraph.text.strip(),
        )
        if not match:
            continue
        equation_paragraph = next(
            (
                candidate
                for candidate in reversed(paragraphs[:index])
                if candidate._p.find(qn("m:oMathPara")) is not None
                or candidate._p.find(qn("m:oMath")) is not None
            ),
            None,
        )
        if equation_paragraph is None:
            raise BuildError("公式编号标记前未找到公式。")
        equation_paragraph.style = document.styles[STYLE_EQUATION]
        _format_equation_number(
            equation_paragraph,
            f"{match.group(1)}.{match.group(2)}",
            labels,
            discipline=discipline,
            equation_sequence=word_equation_sequence,
        )
        paragraph._p.getparent().remove(paragraph._p)

    cover_marker = None
    abstract_marker = None
    front_marker = None
    toc_paragraph = None
    lof_paragraph = None
    lot_paragraph = None
    for paragraph in document.paragraphs:
        value = paragraph.text.strip()
        if value == SECTION_COVER_END:
            cover_marker = paragraph
        elif value == SECTION_ABSTRACT_END:
            abstract_marker = paragraph
        elif value == SECTION_FRONT_END:
            front_marker = paragraph
        elif value == TOC_MARKER:
            toc_paragraph = paragraph
        elif value == LOF_MARKER:
            lof_paragraph = paragraph
        elif value == LOT_MARKER:
            lot_paragraph = paragraph
    if (
        cover_marker is None
        or abstract_marker is None
        or front_marker is None
        or toc_paragraph is None
    ):
        missing = [
            name
            for value, name in (
                (cover_marker, "封面分节"),
                (abstract_marker, "摘要分节"),
                (front_marker, "前置分节"),
                (toc_paragraph, "目录"),
            )
            if value is None
        ]
        raise BuildError("Word 后处理标记缺失：" + "、".join(missing))

    _build_front_matter(
        document,
        cover_marker,
        metadata,
        project_root,
        main_dir,
        metadata_dir,
        allow_project_fallback=allow_project_fallback,
        include_cover=include_cover,
        include_declarations=include_declarations,
    )

    toc_paragraph.text = "目录"
    toc_paragraph.style = document.styles[STYLE_TOC_TITLE]
    _insert_toc_after(toc_paragraph)
    if lof_paragraph is not None:
        lof_paragraph.text = "图目录"
        lof_paragraph.style = document.styles[STYLE_TOC_TITLE]
        _insert_caption_list_after(lof_paragraph, word_figure_sequence)
    if lot_paragraph is not None:
        lot_paragraph.text = "表目录"
        lot_paragraph.style = document.styles[STYLE_TOC_TITLE]
        _insert_caption_list_after(lot_paragraph, word_table_sequence)
    _add_section_break(cover_marker, body_sect_pr, None)
    _add_section_break(abstract_marker, body_sect_pr, "upperRoman")
    _add_section_break(front_marker, body_sect_pr, None)
    _set_page_number_format(document.sections[-1], "decimal", 1)

    stage = output_docx.with_suffix(".sections.docx")
    document.save(stage)
    document = Document(stage)
    if len(document.sections) != 4:
        stage.unlink(missing_ok=True)
        raise BuildError(f"预期生成 4 个 Word 分节，实际为 {len(document.sections)}。")

    for section in document.sections:
        section.start_type = WD_SECTION_START.NEW_PAGE
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.5)
        section.header_distance = Cm(1.5)
        section.footer_distance = Cm(1.5)
        section.different_first_page_header_footer = False

    cover, abstracts, toc, body = document.sections
    for section in (cover, abstracts, toc, body):
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        _clear_story(section.header)
        _clear_story(section.footer)

    _set_page_number_format(cover, None)
    _set_page_number_format(abstracts, "upperRoman", 1)
    _set_page_number_format(toc, None)
    _set_page_number_format(body, "decimal", 1)

    front_footer = abstracts.footer.paragraphs[0]
    front_footer.style = document.styles["Footer"]
    _add_page_field(front_footer)
    for run in front_footer.runs:
        _set_run_fonts(run, "SimSun", "Times New Roman", 9)

    body_header = body.header.paragraphs[0]
    body_header.style = document.styles["Header"]
    body_header.text = title
    body_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_header_border(body_header)
    for run in body_header.runs:
        _set_run_fonts(run, "SimSun", "Times New Roman", 9)

    body_footer = body.footer.paragraphs[0]
    body_footer.style = document.styles["Footer"]
    _add_page_field(body_footer)
    for run in body_footer.runs:
        _set_run_fonts(run, "SimSun", "Times New Roman", 9)

    reusable_style_map = {
        "Normal": STYLE_BODY,
        "Body Text": STYLE_BODY,
        "Compact": STYLE_BODY,
        "First Paragraph": STYLE_FIRST_PARAGRAPH,
        "Heading 1": STYLE_HEADING_1,
        "Heading 2": STYLE_HEADING_2,
        "Heading 3": STYLE_HEADING_3,
        "Heading 4": STYLE_HEADING_4,
        "Yibin Heading 1": STYLE_HEADING_1,
        "Yibin Heading 2": STYLE_HEADING_2,
        "Yibin Heading 3": STYLE_HEADING_3,
        "Yibin Heading 4": STYLE_HEADING_4,
        "ChineseAbstract": STYLE_CHINESE_ABSTRACT,
        "EnglishAbstract": STYLE_ENGLISH_ABSTRACT,
        "Keywords": STYLE_KEYWORDS,
        "Bibliography": STYLE_BIBLIOGRAPHY,
        "Notes": STYLE_NOTES,
    }
    for paragraph in document.paragraphs:
        style_name = paragraph.style.name
        target_style = reusable_style_map.get(style_name)
        if target_style:
            paragraph.style = document.styles[target_style]
        if style_name == "FrontTitle":
            paragraph.style = document.styles[
                STYLE_ENGLISH_ABSTRACT_TITLE
                if paragraph.text.strip() == "Abstract"
                else STYLE_FRONT_TITLE
            ]
        for run in paragraph.runs:
            if paragraph.style.name in reusable_style_map.values():
                _clear_direct_run_typography(run)

    _apply_list_paragraph_styles(document)

    heading_one = document.styles[STYLE_HEADING_1]
    if discipline == "science":
        heading_one.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading_one.paragraph_format.first_line_indent = Pt(0)
        heading_one.paragraph_format.page_break_before = True
        heading_ppr = heading_one.element.get_or_add_pPr()
        heading_indent = heading_ppr.find(qn("w:ind"))
        if heading_indent is not None:
            heading_indent.set(qn("w:firstLineChars"), "0")
            heading_indent.set(qn("w:firstLine"), "0")
    else:
        heading_one.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        heading_one.paragraph_format.first_line_indent = Pt(24)
        heading_one.paragraph_format.page_break_before = True
        heading_ppr = heading_one.element.get_or_add_pPr()
        heading_indent = heading_ppr.find(qn("w:ind"))
        if heading_indent is None:
            heading_indent = OxmlElement("w:ind")
            heading_ppr.append(heading_indent)
        heading_indent.set(qn("w:firstLine"), "480")
        heading_indent.set(qn("w:firstLineChars"), "200")

    unnumbered_titles = {"结论", "注释", "参考文献", "附录", "致谢"}
    if discipline == "humanities":
        unnumbered_titles.add("绪论")
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if paragraph.style.name == STYLE_HEADING_1 and text in unnumbered_titles:
            paragraph.style = document.styles[STYLE_UNNUMBERED_HEADING]
            _clear_paragraph_numbering(paragraph)
        if re.match(r"^附录[A-ZＡ-Ｚ0-9一二三四五六七八九十]+(?:\s|　)", text):
            paragraph.style = document.styles[STYLE_APPENDIX_HEADING]

    heading_abstract_id = _configure_heading_numbering(document, discipline)
    if discipline == "humanities":
        _isolate_humanities_introduction_numbering(
            document,
            heading_abstract_id,
        )
    _configure_appendix_section_numbering(document, discipline)

    # All semantic section titles are style-driven.  Pandoc may leave direct
    # paragraph/run formatting on headings even after assigning a named style;
    # remove only those redundant overrides so a user's later style edit
    # changes every matching title consistently.
    title_styles = {
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
    }
    for paragraph in document.paragraphs:
        if paragraph.style.name not in title_styles:
            continue
        _clear_direct_paragraph_format(
            paragraph,
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "spacing",
            "ind",
            "jc",
        )
        for run in paragraph.runs:
            _clear_direct_run_typography(run)

    _format_tables(document, table_layouts, labels)
    _replace_pandoc_checkbox_math(document)
    _clamp_images(document)
    _format_images_and_captions(
        document,
        labels,
        discipline,
        figure_sequence=word_figure_sequence,
        table_sequence=word_table_sequence,
    )
    _replace_reference_markers(document, labels, citations, citation_mode)
    settings = document.settings.element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")

    document.core_properties.title = title
    document.core_properties.author = author
    subject = "宜宾学院本科毕业论文（非官方技术模板）"
    if english_title.strip():
        subject += " / " + english_title.strip()
    document.core_properties.subject = subject
    document.core_properties.keywords = (
        "YibinThesis;document-type=thesis;template-year=2024"
    )
    if citation_mode not in {"linked", "native"}:
        raise BuildError(f"不支持的 Word citation mode：{citation_mode}")
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_docx)
    if citation_mode == "native":
        _inject_bibliography_sources(output_docx, citations, bibliography_items)
    final_normalized_pngs = _normalize_embedded_pngs(output_docx)
    if final_normalized_pngs:
        print(
            f"Word 图片兼容处理：最终文档中另有 {final_normalized_pngs} 个透明 PNG "
            "已转为白底 RGB PNG。"
        )
    stage.unlink(missing_ok=True)

    # Re-open once so corrupt OOXML fails in the build rather than on delivery.
    check = Document(output_docx)
    residual = {
        paragraph.text.strip()
        for paragraph in check.paragraphs
        if paragraph.text.strip().startswith("YIBIN_INTERNAL_")
    }
    if residual:
        raise BuildError("Word 后处理标记未清理：" + ", ".join(sorted(residual)))


def _profiles_module():
    import word_profiles
    return word_profiles.bind_core(sys.modules[__name__])


def _postprocess_proposal_docx(*args, **kwargs):
    return _profiles_module()._postprocess_proposal_docx(*args, **kwargs)


def _postprocess_review_docx(*args, **kwargs):
    return _profiles_module()._postprocess_review_docx(*args, **kwargs)


def _build_non_thesis(*args, **kwargs):
    return _profiles_module()._build_non_thesis(*args, **kwargs)


def _resolve_argument_path(value: str | Path, project_root: Path) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        cwd_candidate = (Path.cwd() / path).resolve()
        project_candidate = (project_root / path).resolve()
        path = cwd_candidate if cwd_candidate.exists() else project_candidate
    return path.resolve()


def build(args: argparse.Namespace) -> Path:
    project_root = Path(__file__).resolve().parents[1]
    main_path = _resolve_argument_path(args.main, project_root)
    if not main_path.is_file():
        raise BuildError(f"主文件不存在：{main_path}")
    main_dir = main_path.parent
    allow_project_fallback = path_is_within(main_path, project_root)
    main_text = main_path.read_text(encoding="utf-8")
    events = parse_main_events(main_text)

    metadata_path: Path | None = None
    if args.metadata:
        metadata_path = _resolve_argument_path(args.metadata, project_root)
    else:
        for event in events:
            if event.kind in {"input", "include"} and event.phase == "pre" and event.value:
                candidate = resolve_source(
                    event.value,
                    main_dir,
                    project_root,
                    allow_project_fallback=allow_project_fallback,
                )
                if "\\yibinsetup" in candidate.read_text(encoding="utf-8"):
                    metadata_path = candidate
                    break
        if metadata_path is None:
            candidates = [main_dir / "metadata.tex"]
            if allow_project_fallback:
                candidates.append(project_root / "metadata.tex")
            metadata_path = next((path.resolve() for path in candidates if path.is_file()), None)
    if metadata_path is None or not metadata_path.is_file():
        raise BuildError("未找到 metadata.tex；可用 --metadata 显式指定。")
    metadata = parse_metadata(metadata_path)
    profile = resolve_profile(main_text, metadata)
    discipline = profile.discipline

    pandoc = find_pandoc(args.pandoc, project_root)
    reference_doc = _resolve_argument_path(args.reference_doc, project_root)
    if not reference_doc.is_file():
        generator = project_root / "lib" / "build_reference_docx.py"
        run_checked(
            [sys.executable, str(generator), "--output", str(reference_doc)],
            cwd=project_root,
        )
    csl = _resolve_argument_path(args.csl, project_root)
    if not csl.is_file():
        raise BuildError(f"CSL 文件不存在：{csl}")

    bibliography: Path | None
    if args.bibliography:
        bibliography = _resolve_argument_path(args.bibliography, project_root)
    else:
        bib_match = re.search(
            r"\\addbibresource(?:\[[^]]*\])?\s*\{([^{}]+)\}",
            strip_tex_comments(main_text),
        )
        bib_name = bib_match.group(1).strip() if bib_match else "references.bib"
        bibliography = next(
            (
                candidate.resolve()
                for candidate in (
                    [main_dir / bib_name]
                    + ([project_root / bib_name] if allow_project_fallback else [])
                )
                if candidate.is_file()
            ),
            None,
        )
    if bibliography is None:
        raise BuildError("未找到 references.bib；可用 --bibliography 显式指定。")

    output = (
        _resolve_argument_path(args.output, project_root)
        if args.output
        else (main_dir / "build" / f"{main_path.stem}.docx").resolve()
    )
    if output.resolve() == reference_doc.resolve():
        raise BuildError("输出文件不能覆盖 word/reference.docx。")

    if profile.document_type != "thesis":
        return _build_non_thesis(
            args,
            profile=profile,
            project_root=project_root,
            main_path=main_path,
            main_text=main_text,
            main_dir=main_dir,
            metadata_path=metadata_path,
            metadata=metadata,
            pandoc=pandoc,
            reference_doc=reference_doc,
            csl=csl,
            bibliography=bibliography,
            output=output,
            allow_project_fallback=allow_project_fallback,
        )

    warnings: list[str] = []
    notes = NoteRegistry()
    counters = HeadingCounters()
    float_counters = FloatCounters()
    labels = LabelRegistry()
    citations = CitationRegistry()
    table_layouts: list[LatexTableLayout] = []
    markdown_parts: list[str] = []

    graphic_paths = collect_graphic_paths(
        main_text,
        main_dir=main_dir,
        project_root=project_root,
        allow_project_fallback=allow_project_fallback,
    )
    resource_candidates = [
        main_dir.resolve(),
        (main_dir / "assets").resolve(),
        (main_dir / "chapters").resolve(),
        *graphic_paths,
    ]
    if allow_project_fallback:
        resource_candidates.extend(
            [project_root.resolve(), (project_root / "assets").resolve()]
        )
    resource_candidates = list(dict.fromkeys(resource_candidates))

    def image_target_resolver(target: str) -> str:
        return resolve_word_image_target(target, resource_candidates)

    pre_events = [event for event in events if event.phase == "pre"]
    include_cover = any(event.kind == "makeyibincover" for event in pre_events)
    include_declarations = any(
        event.kind == "makeyibindeclarations" for event in pre_events
    )
    if not include_cover and not include_declarations:
        raise BuildError("主文件未调用 \\makeyibincover 或 \\makeyibindeclarations。")
    markdown_parts.append(styled_marker(SECTION_COVER_END))

    # Keep conversion intermediates beside the requested output.  Windows
    # installations commonly place TEMP on a small system drive; a thesis with
    # many raster figures can otherwise fail even when the workspace drive has
    # ample free space.
    temporary_root = output.parent / ".tmp"
    temporary_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(
        prefix="yibinthesis-word-",
        dir=temporary_root,
    ) as temporary:
        temp_dir = Path(temporary)
        front_has_content = False
        abstract_section_closed = False
        front_events = [event for event in events if event.phase == "front"]
        for index, event in enumerate(front_events):
            if event.kind in {"frontmatter", "mainmatter", "backmatter"}:
                continue
            if event.kind in {"input", "include"} and event.value:
                source = resolve_source(
                    event.value,
                    main_dir,
                    project_root,
                    allow_project_fallback=allow_project_fallback,
                )
                if source.resolve() == metadata_path.resolve():
                    continue
                text = expand_nested_inputs(
                    source,
                    project_root,
                    allow_project_fallback=allow_project_fallback,
                )
                detect_unsupported(text, source, warnings)
                environment = (
                    "cnabstract"
                    if extract_environment(text, "cnabstract") is not None
                    else "enabstract"
                    if extract_environment(text, "enabstract") is not None
                    else None
                )
                if front_has_content:
                    markdown_parts.append(PAGE_BREAK)
                if environment:
                    markdown_parts.append(
                        abstract_markdown(
                            text,
                            environment=environment,
                            pandoc=pandoc,
                            cwd=main_dir,
                            temp_dir=temp_dir,
                            notes=notes,
                            discipline=discipline,
                            labels=labels,
                            citations=citations,
                            table_layouts=table_layouts,
                        )
                    )
                else:
                    chapter_md = latex_to_markdown(
                        normalize_latex(
                            text,
                            notes,
                            discipline,
                            labels,
                            citations,
                            table_layouts,
                        ),
                        pandoc=pandoc,
                        cwd=main_dir,
                        temp_dir=temp_dir,
                        name=f"front-{index}",
                    )
                    markdown_parts.append(chapter_md)
                front_has_content = True
            elif event.kind == "tableofcontents":
                markdown_parts.append(styled_marker(SECTION_ABSTRACT_END))
                markdown_parts.append(styled_marker(TOC_MARKER))
                abstract_section_closed = True
                front_has_content = True
            elif event.kind == "listoffigures":
                markdown_parts.append(PAGE_BREAK)
                markdown_parts.append(styled_marker(LOF_MARKER))
                front_has_content = True
            elif event.kind == "listoftables":
                markdown_parts.append(PAGE_BREAK)
                markdown_parts.append(styled_marker(LOT_MARKER))
                front_has_content = True
        if not any(event.kind == "tableofcontents" for event in front_events):
            markdown_parts.append(styled_marker(SECTION_ABSTRACT_END))
            markdown_parts.append(styled_marker(TOC_MARKER))
            abstract_section_closed = True
        if not abstract_section_closed:
            markdown_parts.append(styled_marker(SECTION_ABSTRACT_END))
        markdown_parts.append(styled_marker(SECTION_FRONT_END))

        content_events = [event for event in events if event.phase in {"main", "back"}]
        bibliography_inserted = False
        for index, event in enumerate(content_events):
            if event.kind in {"mainmatter", "backmatter", "frontmatter"}:
                continue
            if event.kind in {"input", "include"} and event.value:
                source = resolve_source(
                    event.value,
                    main_dir,
                    project_root,
                    allow_project_fallback=allow_project_fallback,
                )
                text = expand_nested_inputs(
                    source,
                    project_root,
                    allow_project_fallback=allow_project_fallback,
                )
                detect_unsupported(text, source, warnings)
                chapter_before = counters.chapter
                converted = latex_to_markdown(
                    normalize_latex(
                        text,
                        notes,
                        discipline,
                        labels,
                        citations,
                        table_layouts,
                    ),
                    pandoc=pandoc,
                    cwd=main_dir,
                    temp_dir=temp_dir,
                    name=f"chapter-{index}",
                )
                converted = apply_heading_numbering(
                    converted,
                    discipline,
                    counters,
                    labels,
                )
                converted = apply_float_numbering(
                    converted,
                    discipline,
                    chapter_before,
                    float_counters,
                    labels,
                    image_target_resolver,
                )
                markdown_parts.append(converted)
            elif event.kind == "printyibinnotes":
                rendered_notes = notes_markdown(notes)
                if rendered_notes:
                    markdown_parts.append(rendered_notes)
            elif event.kind == "nocite" and event.value:
                citations.add_nocite(event.value)
            elif event.kind == "printyibinbibliography":
                seed = citation_seed_markdown(citations)
                if seed:
                    markdown_parts.append(seed)
                markdown_parts.append(bibliography_markdown())
                bibliography_inserted = True
        if not bibliography_inserted:
            seed = citation_seed_markdown(citations)
            if seed:
                markdown_parts.append(seed)
            markdown_parts.append(bibliography_markdown())

        assembled = "\n\n".join(part for part in markdown_parts if part.strip()) + "\n"
        assembled = labels.resolve(assembled)
        markdown_file = temp_dir / "assembled.md"
        markdown_file.write_text(assembled, encoding="utf-8", newline="\n")
        if args.keep_intermediate:
            kept = output.with_suffix(".pandoc.md")
            kept.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(markdown_file, kept)

        pandoc_docx = temp_dir / "pandoc.docx"
        resource_paths = os.pathsep.join(
            str(path) for path in resource_candidates
        )
        citeproc_bibliography = prepare_bibliography(
            bibliography,
            pandoc=pandoc,
            cwd=main_dir,
            temp_dir=temp_dir,
        )
        bibliography_items: list[dict[str, object]] = []
        if citeproc_bibliography.suffix.lower() == ".json":
            parsed_items = json.loads(citeproc_bibliography.read_text(encoding="utf-8"))
            if isinstance(parsed_items, list):
                bibliography_items = [item for item in parsed_items if isinstance(item, dict)]
        command = [
            str(pandoc),
            "--from=markdown+smart+raw_attribute+fenced_divs+bracketed_spans+citations+tex_math_dollars",
            "--to=docx",
            "--standalone",
            "--citeproc",
            f"--reference-doc={reference_doc}",
            f"--bibliography={citeproc_bibliography}",
            f"--csl={csl}",
            f"--resource-path={resource_paths}",
            "--metadata=lang:zh-CN",
            f"--output={pandoc_docx}",
            str(markdown_file),
        ]
        run_checked(command, cwd=main_dir)
        normalized_pngs = _normalize_embedded_pngs(pandoc_docx)
        if normalized_pngs:
            print(
                f"Word 图片兼容处理：已将 {normalized_pngs} 个透明 PNG 转为白底 RGB PNG。"
            )
        postprocess_docx(
            pandoc_docx,
            output,
            metadata=metadata,
            project_root=project_root,
            main_dir=main_dir,
            metadata_dir=metadata_path.parent,
            discipline=discipline,
            allow_project_fallback=allow_project_fallback,
            include_cover=include_cover,
            include_declarations=include_declarations,
            citation_mode=args.citation_mode,
            labels=labels,
            citations=citations,
            bibliography_items=bibliography_items,
            table_layouts=table_layouts,
            word_figure_sequence=args.word_figure_sequence,
            word_table_sequence=args.word_table_sequence,
            word_equation_sequence=args.word_equation_sequence,
        )

    print(f"Generated: {output}")
    print("Word 输出可编辑；目录和页码域将在 Microsoft Word 打开时自动更新。")
    print("说明：复杂 TikZ、任意自定义宏、复杂浮动体不保证与 PDF 同页或等版式转换。")
    for warning in dict.fromkeys(warnings):
        print(f"WARNING: {warning}", file=sys.stderr)
    return output


def make_parser() -> argparse.ArgumentParser:
    project_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--main",
        required=True,
        help="LaTeX entry point from an external consumer project.",
    )
    parser.add_argument("--metadata", help="Override metadata.tex path.")
    parser.add_argument("--bibliography", help="Override BibTeX database path.")
    parser.add_argument("--output", help="Output DOCX path.")
    parser.add_argument("--pandoc", help="Pandoc executable path.")
    parser.add_argument(
        "--reference-doc",
        default=str(project_root / "word" / "reference.docx"),
        help="Pandoc semantic reference DOCX.",
    )
    parser.add_argument(
        "--csl",
        default=str(project_root / "word" / "china-national-standard-gb-t-7714-2015-numeric.csl"),
        help="Citation Style Language file.",
    )
    parser.add_argument(
        "--keep-intermediate",
        action="store_true",
        help="Keep the assembled Pandoc Markdown beside the output DOCX.",
    )
    parser.add_argument(
        "--citation-mode",
        choices=("linked", "native"),
        default="linked",
        help=(
            "Word 引用模式；linked 保持 citeproc 编号并生成可跳转的参考文献书签，"
            "native 另写入 Word CITATION 域和 bibliography customXml。"
        ),
    )
    parser.add_argument(
        "--word-figure-sequence",
        default="图",
        help="Registered Word figure CaptionLabel name.",
    )
    parser.add_argument(
        "--word-table-sequence",
        default="表",
        help="Registered Word table CaptionLabel name.",
    )
    parser.add_argument(
        "--word-equation-sequence",
        default="公式",
        help="Registered Word equation CaptionLabel name.",
    )
    return parser


def main() -> int:
    try:
        build(make_parser().parse_args())
    except BuildError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("ERROR: 构建已中止。", file=sys.stderr)
        return 130
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
