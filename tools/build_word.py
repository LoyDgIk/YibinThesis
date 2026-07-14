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
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


SECTION_COVER_END = "YIBIN_INTERNAL_SECTION_COVER_END_5B8D"
SECTION_ABSTRACT_END = "YIBIN_INTERNAL_SECTION_ABSTRACT_END_6458"
SECTION_FRONT_END = "YIBIN_INTERNAL_SECTION_FRONT_END_8F91"
TOC_MARKER = "YIBIN_INTERNAL_TOC_76EE"
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
            rendered = f"({target.number})" if parenthesized else target.number
            text = text.replace(token, rendered)
        if missing:
            raise BuildError("未找到交叉引用标签：" + "、".join(sorted(missing)))
        residual = sorted(set(re.findall(r"YIBINXREF\d{8}", text)))
        if residual:
            raise BuildError("交叉引用占位符未清理：" + "、".join(residual))
        return text


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


def parse_main_events(main_text: str) -> list[MainEvent]:
    text = strip_tex_comments(main_text)
    pattern = re.compile(
        r"\\(?P<kind>frontmatter|mainmatter|backmatter|"
        r"makeyibincover|makeyibindeclarations|tableofcontents|"
        r"printyibinnotes|printyibinbibliography|input|include)"
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


def resolve_source(name: str, main_dir: Path, project_root: Path) -> Path:
    raw = Path(name)
    names = [raw] if raw.suffix else [raw.with_suffix(".tex"), raw]
    for base in (main_dir, project_root):
        for candidate_name in names:
            candidate = (base / candidate_name).resolve()
            if candidate.is_file():
                return candidate
    raise BuildError(f"找不到主文件引用的源文件：{name}")


def expand_nested_inputs(
    path: Path,
    project_root: Path,
    seen: set[Path] | None = None,
) -> str:
    seen = seen or set()
    resolved = path.resolve()
    if resolved in seen:
        raise BuildError(f"检测到循环 \\input：{resolved}")
    seen.add(resolved)
    text = path.read_text(encoding="utf-8")
    pattern = re.compile(r"\\(?:input|include)\s*\{([^{}]+)\}")

    def include(match: re.Match[str]) -> str:
        child = resolve_source(match.group(1).strip(), path.parent, project_root)
        return expand_nested_inputs(child, project_root, seen.copy())

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


def normalize_latex(
    text: str,
    notes: NoteRegistry,
    discipline: str = "humanities",
    labels: LabelRegistry | None = None,
) -> str:
    text = strip_tex_comments(text)
    labels = labels if labels is not None else LabelRegistry()
    chapter_command = "chapter" if discipline == "science" else "chapter*"

    def thematic_chapter(match: re.Match[str], default: str) -> str:
        title = match.group("title") if match.group("title") is not None else default
        return rf"\{chapter_command}{{{title}}}"

    text = re.sub(
        r"\\yibinopeningchapter(?![A-Za-z@])(?:\s*\[(?P<title>[^]]*)\])?",
        lambda match: thematic_chapter(match, "绪论"),
        text,
    )
    text = re.sub(
        r"\\yibinclosingchapter(?![A-Za-z@])(?:\s*\[(?P<title>[^]]*)\])?",
        lambda match: thematic_chapter(match, "结论"),
        text,
    )
    text = replace_braced_command(
        text,
        "yibinunnumberedchapter",
        lambda title: rf"\chapter*{{{title}}}",
    )
    text = replace_braced_command(
        text,
        "yibincite",
        lambda keys: rf"\cite{{{keys}}}",
    )
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
) -> list[Path]:
    """Return existing directories declared through ``\\graphicspath``.

    LaTeX resolves graphics against its process/search path, while the
    Pandoc DOCX writer resolves image targets through ``--resource-path``.
    Include both project-root-relative and entry-file-relative candidates so
    reusable examples can keep their images in a local ``assets`` folder.
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
                options = [project_root / candidate, main_dir / candidate]
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
            rendered = f"{prefix} {title}"
        elif level == 1:
            reference_number = f"{chinese_number(counters.chapter)}、"
            rendered = f"{chinese_number(counters.chapter)}、{title}"
        elif level == 2:
            reference_number = f"（{chinese_number(counters.section)}）"
            rendered = f"（{chinese_number(counters.section)}）{title}"
        elif level == 3:
            reference_number = f"{counters.subsection}."
            rendered = f"{counters.subsection}.{title}"
        else:
            reference_number = f"（{counters.subsubsection}）"
            rendered = f"（{counters.subsubsection}）{title}"
        identifier = re.search(r"(?:^|\s)#([^\s}]+)", attribute_text)
        if identifier:
            labels.register(identifier.group(1), reference_number, "标题")
        if level == 1:
            output.append(f"<!-- YIBIN_INTERNAL_CHAPTER_{counters.chapter} -->")
        output.append(f"{'#' * level} {rendered}{attribute_text}")
    return "\n".join(output)


def apply_float_numbering(
    markdown: str,
    discipline: str,
    chapter: int,
    counters: FloatCounters,
    labels: LabelRegistry | None = None,
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
            r"(?P<attrs>\{[^{}]*\})?\s*$",
            markdown_figure,
            chunk,
        )

        def figure(match: re.Match[str]) -> str:
            number = next_number("figure")
            attrs = match.group("attrs")
            identifier = re.search(r'\bid="([^"]+)"', attrs)
            if identifier:
                labels.register(identifier.group(1), number, "图")
            caption = re.sub(r"<[^>]+>", "", match.group("caption")).strip()
            body = match.group("body").strip()
            body = re.sub(r"^<p>(.*?)</p>$", r"\1", body, flags=re.DOTALL)

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
                        value = re.search(
                            rf"(?:^|;)\s*{dimension}\s*:\s*([^;]+)",
                            style,
                            flags=re.IGNORECASE,
                        )
                        if value:
                            options.append(f"{dimension}={value.group(1).strip()}")
                target = f"<{src}>" if re.search(r"[\s()]", src) else src
                attributes_md = "{" + " ".join(options) + "}" if options else ""
                return f"![{alt}]({target}){attributes_md}"

            body = re.sub(
                r"<img\s+(?P<attributes>[^>]*?)/?>",
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
            return f"$$\n{body}\n$$\n\n{marker}"

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


def cover_markdown(metadata: dict[str, str], project_root: Path) -> str:
    title = escape_markdown(metadata.get("title", "（填写中文题目）"))
    blocks: list[str] = []
    logo_value = metadata.get("logo", "assets/yibin-university-logo.png")
    logo_candidates = [
        project_root / logo_value,
        project_root / "assets" / "yibin-university-logo.png",
        project_root / "assets" / "yibin-logo.png",
    ]
    logo = next((path for path in logo_candidates if path.is_file()), None)
    if logo:
        try:
            resource = logo.relative_to(project_root).as_posix()
        except ValueError:
            resource = logo.as_posix()
        # A non-empty Markdown image description becomes a visible figure
        # caption in Pandoc DOCX output.  The logo is decorative on the cover,
        # so keep the description empty and avoid a duplicate blue school name.
        blocks.append(custom_block("CoverLogo", f"![]({resource}){{width=10cm}}"))
    else:
        blocks.append(custom_block("CoverSchool", "宜宾学院"))
    blocks.append(custom_block("CoverThesisType", "本科生毕业论文（设计）"))
    blocks.append(custom_block("CoverTitle", title))

    fields = [
        ("学院（部）", metadata.get("college", "")),
        ("专　　业", metadata.get("major", "")),
        ("学生姓名", metadata.get("author", "")),
        ("学　　号", metadata.get("student-id", "")),
        ("年　　级", metadata.get("grade", "")),
        (
            "指导教师（校内）",
            "　".join(
                item
                for item in (metadata.get("advisor", ""), metadata.get("advisor-title", ""))
                if item
            ),
        ),
    ]
    external = "　".join(
        item
        for item in (
            metadata.get("external-advisor", ""),
            metadata.get("external-advisor-title", ""),
        )
        if item
    )
    if external:
        fields.append(("指导教师（校外）", external))
    for label, value in fields:
        line = (
            f'[{escape_markdown(label)}]{{custom-style="CoverLabel"}}　'
            f'[{escape_markdown(value or "（待填写）")}]{{custom-style="CoverValue"}}'
        )
        blocks.append(custom_block("CoverField", line))
    blocks.append(custom_block("CoverDate", escape_markdown(metadata.get("date", ""))))
    return "\n\n".join(blocks)


def declarations_markdown(metadata: dict[str, str]) -> str:
    title = metadata.get("title", "（填写中文题目）")
    original = "\n\n".join(
        [
            custom_block("DeclarationTitle", "原创性声明"),
            custom_block(
                "DeclarationBody",
                "本人呈交的学位论文（设计）《"
                + escape_markdown(title)
                + "》，是在导师的指导下，独立进行研究取得的成果。除文中已经注明引用的内容外，"
                "本论文（设计）不包括其他个人或集体已经发表或撰写过的作品成果。对本文（设计）"
                "做出贡献的个人和集体，均已在文中以明确方式标明。本人完全意识到本声明的法律"
                "后果，因本声明而产生的法律后果由本人承担。",
            ),
            custom_block("DeclarationBody", "学位论文作者（签名）：＿＿＿＿＿＿＿＿"),
            custom_block("DeclarationBody", "日期：＿＿＿＿年＿＿月＿＿日"),
            custom_block(
                "DeclarationBody",
                "附：《普通高等学校学生管理规定》（中华人民共和国教育部令第 41 号）第五十二条："
                "学生有下列情形之一，学校可以给予开除学籍处分：（五）学位论文、公开发表的研究"
                "成果存在抄袭、篡改、伪造等学术不端行为，情节严重的，或者代写论文、买卖论文的。",
            ),
        ]
    )

    confidential = metadata.get("secrecy", "public").strip().lower() == "confidential"
    declassify = metadata.get("declassify-year", "") or "____"
    authorization = "\n\n".join(
        [
            custom_block("DeclarationTitle", "学位论文（设计）版权使用授权书"),
            custom_block(
                "DeclarationBody",
                "本学位论文（设计）作者完全了解学校有关保留、使用学位论文（设计）的规定，同意"
                "学校保留并向国家有关部门或机构送交论文（设计）的复印件和电子版，允许论文"
                "（设计）被查阅和借阅。本人授权宜宾学院将本学位论文（设计）的全部或部分内容"
                "编入有关数据库进行检索，可以采用影印、缩印或扫描等复制手段保存和汇编。",
            ),
            custom_block(
                "DeclarationBody",
                f"{'■' if confidential else '□'} 保密，在 {escape_markdown(declassify)} 年解密后适用本授权书。",
            ),
            custom_block("DeclarationBody", f"{'□' if confidential else '■'} 不保密。"),
            custom_block(
                "DeclarationBody",
                "作者（签名）：＿＿＿＿＿＿＿＿　　指导教师（签名）：＿＿＿＿＿＿＿＿",
            ),
            custom_block(
                "DeclarationBody",
                "日期：＿＿＿＿年＿＿月＿＿日　　　　日期：＿＿＿＿年＿＿月＿＿日",
            ),
        ]
    )
    return f"{original}\n\n{PAGE_BREAK}\n\n{authorization}"


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
) -> str:
    body = extract_environment(source_text, environment)
    if body is None:
        raise BuildError(f"摘要文件缺少 {environment} 环境。")
    keyword_command = "cnkeywords" if environment == "cnabstract" else "enkeywords"
    keywords = find_command_argument(body, keyword_command) or ""
    body = replace_braced_command(body, keyword_command, lambda _: "")
    body_md = latex_to_markdown(
        normalize_latex(body, notes, discipline, labels),
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


def _insert_toc_after(paragraph) -> None:
    toc_paragraph = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    p_style = OxmlElement("w:pStyle")
    p_style.set(qn("w:val"), "Normal")
    p_pr.append(p_style)
    toc_paragraph.append(p_pr)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), 'TOC \\o "1-3" \\h \\z \\u')
    field.set(qn("w:dirty"), "true")
    toc_paragraph.append(field)
    paragraph._p.addnext(toc_paragraph)


def _format_equation_number(paragraph, number: str) -> None:
    """Center an OMML equation and place its number at the right margin."""
    if paragraph._p.find(qn("m:oMathPara")) is None and paragraph._p.find(qn("m:oMath")) is None:
        raise BuildError(f"公式编号 {number} 前未找到 Word 公式对象。")
    p_pr = paragraph._p.get_or_add_pPr()
    tabs = p_pr.find(qn("w:tabs"))
    if tabs is not None:
        p_pr.remove(tabs)
    tabs = OxmlElement("w:tabs")
    center = OxmlElement("w:tab")
    center.set(qn("w:val"), "center")
    center.set(qn("w:pos"), "4394")
    right = OxmlElement("w:tab")
    right.set(qn("w:val"), "right")
    right.set(qn("w:pos"), "8787")
    tabs.extend([center, right])
    insertion = 0
    for index, child in enumerate(p_pr):
        if child.tag in {qn("w:pStyle"), qn("w:keepNext"), qn("w:keepLines"), qn("w:pageBreakBefore")}:
            insertion = index + 1
    p_pr.insert(insertion, tabs)

    justification = p_pr.find(qn("w:jc"))
    if justification is None:
        justification = OxmlElement("w:jc")
        p_pr.append(justification)
    justification.set(qn("w:val"), "left")

    leading_run = OxmlElement("w:r")
    leading_run.append(OxmlElement("w:tab"))
    paragraph._p.insert(1 if paragraph._p[0] is p_pr else 0, leading_run)
    trailing_run = OxmlElement("w:r")
    trailing_run.append(OxmlElement("w:tab"))
    text = OxmlElement("w:t")
    text.text = f"({number})"
    trailing_run.append(text)
    paragraph._p.append(trailing_run)


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
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    result = OxmlElement("w:t")
    result.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, separate, result, end])


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


def _set_run_fonts(run, east_asia: str, latin: str, size: float) -> None:
    run.font.name = latin
    run.font.size = Pt(size)
    r_pr = run._r.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), latin)
    r_fonts.set(qn("w:hAnsi"), latin)
    r_fonts.set(qn("w:cs"), latin)
    r_fonts.set(qn("w:eastAsia"), east_asia)


def _format_tables(document: Document) -> None:
    target_width = int(Cm(15.5).twips)
    for table in document.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        tbl_pr = table._tbl.tblPr

        table_width = tbl_pr.find(qn("w:tblW"))
        if table_width is None:
            table_width = OxmlElement("w:tblW")
            tbl_pr.append(table_width)
        table_width.set(qn("w:type"), "dxa")
        table_width.set(qn("w:w"), str(target_width))

        table_indent = tbl_pr.find(qn("w:tblInd"))
        if table_indent is None:
            table_indent = OxmlElement("w:tblInd")
            tbl_pr.append(table_indent)
        table_indent.set(qn("w:type"), "dxa")
        table_indent.set(qn("w:w"), "0")

        layout = tbl_pr.find(qn("w:tblLayout"))
        if layout is None:
            layout = OxmlElement("w:tblLayout")
            tbl_pr.append(layout)
        layout.set(qn("w:type"), "fixed")

        cell_margins = tbl_pr.find(qn("w:tblCellMar"))
        if cell_margins is None:
            cell_margins = OxmlElement("w:tblCellMar")
            tbl_pr.append(cell_margins)
        for edge, width in (("top", 80), ("left", 120), ("bottom", 80), ("right", 120)):
            margin = cell_margins.find(qn(f"w:{edge}"))
            if margin is None:
                margin = OxmlElement(f"w:{edge}")
                cell_margins.append(margin)
            margin.set(qn("w:type"), "dxa")
            margin.set(qn("w:w"), str(width))

        grid = table._tbl.tblGrid
        grid_columns = list(grid.iterchildren(qn("w:gridCol")))
        column_count = max(1, len(table.columns))
        raw_widths = [int(column.get(qn("w:w"), "0")) for column in grid_columns]
        if len(raw_widths) != column_count or sum(raw_widths) <= 0:
            raw_widths = [1] * column_count
            while len(grid_columns) < column_count:
                column = OxmlElement("w:gridCol")
                grid.append(column)
                grid_columns.append(column)
        scale = target_width / sum(raw_widths)
        widths = [max(1, round(width * scale)) for width in raw_widths]
        widths[-1] += target_width - sum(widths)
        for column, width in zip(grid_columns, widths):
            column.set(qn("w:w"), str(width))

        borders = tbl_pr.find(qn("w:tblBorders"))
        if borders is None:
            borders = OxmlElement("w:tblBorders")
            tbl_pr.append(borders)
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
        for row_index, row in enumerate(table.rows):
            tr_pr = row._tr.get_or_add_trPr()
            cant_split = tr_pr.find(qn("w:cantSplit"))
            if cant_split is None:
                tr_pr.append(OxmlElement("w:cantSplit"))
            if row_index == 0:
                header = tr_pr.find(qn("w:tblHeader"))
                if header is None:
                    header = OxmlElement("w:tblHeader")
                    header.set(qn("w:val"), "true")
                    tr_pr.append(header)
            grid_index = 0
            for cell_xml in row._tr.findall(qn("w:tc")):
                tc_pr = cell_xml.find(qn("w:tcPr"))
                if tc_pr is None:
                    tc_pr = OxmlElement("w:tcPr")
                    cell_xml.insert(0, tc_pr)
                grid_span = tc_pr.find(qn("w:gridSpan"))
                span = int(grid_span.get(qn("w:val"), "1")) if grid_span is not None else 1
                cell_width = sum(widths[grid_index : grid_index + span])
                grid_index += span
                tc_width = tc_pr.find(qn("w:tcW"))
                if tc_width is None:
                    tc_width = OxmlElement("w:tcW")
                    tc_pr.append(tc_width)
                tc_width.set(qn("w:type"), "dxa")
                tc_width.set(qn("w:w"), str(cell_width))

            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if row_index == 0:
                    tc_pr = cell._tc.get_or_add_tcPr()
                    cell_borders = tc_pr.find(qn("w:tcBorders"))
                    if cell_borders is None:
                        cell_borders = OxmlElement("w:tcBorders")
                        tc_pr.append(cell_borders)
                    bottom = cell_borders.find(qn("w:bottom"))
                    if bottom is None:
                        bottom = OxmlElement("w:bottom")
                        cell_borders.append(bottom)
                    bottom.set(qn("w:val"), "single")
                    bottom.set(qn("w:sz"), "8")
                    bottom.set(qn("w:color"), "000000")
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.first_line_indent = Pt(0)
                    if row_index == 0:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        _set_run_fonts(run, "SimSun", "Times New Roman", 10.5)


def _clamp_images(document: Document) -> None:
    maximum_width = int(15.5 * 360000)  # 15.5 cm usable width in EMU.
    for inline in document.element.iter(qn("wp:inline")):
        extent = inline.find(qn("wp:extent"))
        if extent is None:
            continue
        width = int(extent.get("cx", "0"))
        height = int(extent.get("cy", "0"))
        if width <= maximum_width or width <= 0:
            continue
        scale = maximum_width / width
        new_width = maximum_width
        new_height = max(1, int(height * scale))
        extent.set("cx", str(new_width))
        extent.set("cy", str(new_height))
        for drawing_extent in inline.iter(qn("a:ext")):
            drawing_extent.set("cx", str(new_width))
            drawing_extent.set("cy", str(new_height))


def postprocess_docx(
    input_docx: Path,
    output_docx: Path,
    *,
    title: str,
    english_title: str,
    author: str,
    discipline: str,
) -> None:
    document = Document(input_docx)
    body_sect_pr = document._element.body.sectPr
    if body_sect_pr is None:
        raise BuildError("Pandoc 输出缺少正文分节属性。")

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
        _format_equation_number(
            equation_paragraph,
            f"{match.group(1)}.{match.group(2)}",
        )
        paragraph._p.getparent().remove(paragraph._p)

    cover_marker = None
    abstract_marker = None
    front_marker = None
    toc_paragraph = None
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

    toc_paragraph.text = "目录"
    toc_paragraph.style = document.styles["Yibin TOC Heading"]
    _insert_toc_after(toc_paragraph)
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

    for paragraph in document.paragraphs:
        style_name = paragraph.style.name
        if style_name in {"Heading 1", "Heading 2", "Heading 3", "Heading 4"}:
            paragraph.style = document.styles[f"Yibin {style_name}"]

    heading_one = document.styles["Yibin Heading 1"]
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
        heading_one.paragraph_format.page_break_before = False
        heading_ppr = heading_one.element.get_or_add_pPr()
        heading_indent = heading_ppr.find(qn("w:ind"))
        if heading_indent is None:
            heading_indent = OxmlElement("w:ind")
            heading_ppr.append(heading_indent)
        heading_indent.set(qn("w:firstLine"), "480")
        heading_indent.set(qn("w:firstLineChars"), "200")

    unnumbered_titles = {"绪论", "结论", "注释", "参考文献", "附录", "致谢"}
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if (
            paragraph.style.name == "Yibin Heading 1"
            and (text in unnumbered_titles or text.startswith("附录 "))
        ):
            paragraph.style = document.styles["Unnumbered Heading 1"]

    _format_tables(document)
    _clamp_images(document)
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
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_docx)
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
    main_text = main_path.read_text(encoding="utf-8")
    events = parse_main_events(main_text)
    discipline = parse_discipline(main_text)

    metadata_path: Path | None = None
    if args.metadata:
        metadata_path = _resolve_argument_path(args.metadata, project_root)
    else:
        for event in events:
            if event.kind in {"input", "include"} and event.phase == "pre" and event.value:
                candidate = resolve_source(event.value, main_dir, project_root)
                if "\\yibinsetup" in candidate.read_text(encoding="utf-8"):
                    metadata_path = candidate
                    break
        if metadata_path is None:
            candidates = [main_dir / "metadata.tex", project_root / "metadata.tex"]
            metadata_path = next((path.resolve() for path in candidates if path.is_file()), None)
    if metadata_path is None or not metadata_path.is_file():
        raise BuildError("未找到 metadata.tex；可用 --metadata 显式指定。")
    metadata = parse_metadata(metadata_path)

    pandoc = find_pandoc(args.pandoc, project_root)
    reference_doc = _resolve_argument_path(args.reference_doc, project_root)
    if not reference_doc.is_file():
        generator = project_root / "tools" / "build_reference_docx.py"
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
                for candidate in (main_dir / bib_name, project_root / bib_name)
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

    warnings: list[str] = []
    notes = NoteRegistry()
    counters = HeadingCounters()
    float_counters = FloatCounters()
    labels = LabelRegistry()
    markdown_parts: list[str] = []

    pre_events = [event for event in events if event.phase == "pre"]
    if any(event.kind == "makeyibincover" for event in pre_events):
        markdown_parts.append(cover_markdown(metadata, project_root))
    if any(event.kind == "makeyibindeclarations" for event in pre_events):
        if markdown_parts:
            markdown_parts.append(PAGE_BREAK)
        markdown_parts.append(declarations_markdown(metadata))
    if not markdown_parts:
        raise BuildError("主文件未调用 \\makeyibincover 或 \\makeyibindeclarations。")
    markdown_parts.append(styled_marker(SECTION_COVER_END))

    with tempfile.TemporaryDirectory(prefix="yibinthesis-word-") as temporary:
        temp_dir = Path(temporary)
        front_has_content = False
        abstract_section_closed = False
        front_events = [event for event in events if event.phase == "front"]
        for index, event in enumerate(front_events):
            if event.kind in {"frontmatter", "mainmatter", "backmatter"}:
                continue
            if event.kind in {"input", "include"} and event.value:
                source = resolve_source(event.value, main_dir, project_root)
                if source.resolve() == metadata_path.resolve():
                    continue
                text = expand_nested_inputs(source, project_root)
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
                        )
                    )
                else:
                    chapter_md = latex_to_markdown(
                        normalize_latex(text, notes, discipline, labels),
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
                source = resolve_source(event.value, main_dir, project_root)
                text = expand_nested_inputs(source, project_root)
                detect_unsupported(text, source, warnings)
                chapter_before = counters.chapter
                converted = latex_to_markdown(
                    normalize_latex(text, notes, discipline, labels),
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
                )
                markdown_parts.append(converted)
            elif event.kind == "printyibinnotes":
                rendered_notes = notes_markdown(notes)
                if rendered_notes:
                    markdown_parts.append(rendered_notes)
            elif event.kind == "printyibinbibliography":
                markdown_parts.append(bibliography_markdown())
                bibliography_inserted = True
        if not bibliography_inserted:
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
        graphic_paths = collect_graphic_paths(
            main_text,
            main_dir=main_dir,
            project_root=project_root,
        )
        resource_paths = os.pathsep.join(
            str(path)
            for path in dict.fromkeys(
                [
                    project_root.resolve(),
                    main_dir.resolve(),
                    (project_root / "assets").resolve(),
                    (main_dir / "assets").resolve(),
                    (main_dir / "chapters").resolve(),
                    *graphic_paths,
                ]
            )
        )
        citeproc_bibliography = prepare_bibliography(
            bibliography,
            pandoc=pandoc,
            cwd=main_dir,
            temp_dir=temp_dir,
        )
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
        postprocess_docx(
            pandoc_docx,
            output,
            title=metadata.get("title", ""),
            english_title=metadata.get("english-title", ""),
            author=metadata.get("author", ""),
            discipline=discipline,
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
        default=str(project_root / "main.tex"),
        help="LaTeX entry point; examples may provide their own main.tex.",
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
