from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "yibinthesis_build_word", ROOT / "tools" / "build_word.py"
)
assert SPEC is not None and SPEC.loader is not None
BUILD_WORD = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = BUILD_WORD
SPEC.loader.exec_module(BUILD_WORD)


class ExternalEntryResolutionTests(unittest.TestCase):
    def test_external_source_does_not_fall_back_to_template(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            base = Path(temporary)
            template = base / "template"
            external = base / "external"
            template.mkdir()
            external.mkdir()
            (template / "metadata.tex").write_text("template", encoding="utf-8")

            with self.assertRaises(BUILD_WORD.BuildError):
                BUILD_WORD.resolve_source(
                    "metadata",
                    external,
                    template,
                    allow_project_fallback=False,
                )

            self.assertEqual(
                BUILD_WORD.resolve_source(
                    "metadata",
                    external,
                    template,
                    allow_project_fallback=True,
                ),
                (template / "metadata.tex").resolve(),
            )

    def test_external_source_wins_when_names_collide(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            base = Path(temporary)
            template = base / "template"
            external = base / "external"
            template.mkdir()
            external.mkdir()
            (template / "chapter.tex").write_text("template", encoding="utf-8")
            (external / "chapter.tex").write_text("external", encoding="utf-8")

            selected = BUILD_WORD.resolve_source(
                "chapter",
                external,
                template,
                allow_project_fallback=False,
            )
            self.assertEqual(selected, (external / "chapter.tex").resolve())

    def test_external_table_layout_semantics_win_when_names_collide(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            base = Path(temporary)
            template = base / "template"
            external = base / "external"
            template.mkdir()
            external.mkdir()
            (template / "chapter.tex").write_text(
                r"\begin{tabular}{rr}a&b\\\end{tabular}",
                encoding="utf-8",
            )
            (external / "chapter.tex").write_text(
                r"\begin{tabular}{lc}a&b\\\end{tabular}",
                encoding="utf-8",
            )

            selected = BUILD_WORD.resolve_source(
                "chapter",
                external,
                template,
                allow_project_fallback=False,
            )
            layouts = BUILD_WORD.extract_latex_table_layouts(
                selected.read_text(encoding="utf-8")
            )
            self.assertEqual(
                [column.horizontal for column in layouts[0].columns],
                ["left", "center"],
            )

    def test_external_graphics_do_not_use_template_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            base = Path(temporary)
            template = base / "template"
            external = base / "external"
            (template / "assets").mkdir(parents=True)
            external.mkdir()
            main_text = r"\graphicspath{{assets/}}"

            self.assertEqual(
                BUILD_WORD.collect_graphic_paths(
                    main_text,
                    main_dir=external,
                    project_root=template,
                    allow_project_fallback=False,
                ),
                [],
            )

            (external / "assets").mkdir()
            self.assertEqual(
                BUILD_WORD.collect_graphic_paths(
                    main_text,
                    main_dir=external,
                    project_root=template,
                    allow_project_fallback=False,
                ),
                [(external / "assets").resolve()],
            )

    def test_external_logo_missing_is_an_error(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            base = Path(temporary)
            template = base / "template"
            external = base / "external"
            template.mkdir()
            external.mkdir()
            (template / "assets").mkdir()
            (template / "assets" / "yibin-university-logo.png").write_bytes(b"png")

            with self.assertRaises(BUILD_WORD.BuildError):
                BUILD_WORD._resolve_cover_logo(
                    {"logo": "assets/yibin-university-logo.png"},
                    template,
                    external,
                    external,
                    allow_project_fallback=False,
                )


class WordConversionNormalizationTests(unittest.TestCase):
    @staticmethod
    def _field_instructions(paragraph) -> list[str]:
        instructions = [
            node.text or ""
            for node in paragraph._p.iter(BUILD_WORD.qn("w:instrText"))
        ]
        instructions.extend(
            node.get(BUILD_WORD.qn("w:instr"), "")
            for node in paragraph._p.iter(BUILD_WORD.qn("w:fldSimple"))
        )
        return instructions

    def test_heading_numbering_is_not_written_into_heading_text(self) -> None:
        labels = BUILD_WORD.LabelRegistry()
        counters = BUILD_WORD.HeadingCounters()
        converted = BUILD_WORD.apply_heading_numbering(
            "# 理论与框架 {#sec:theory}\n## 标准需求\n### 对象边界",
            "humanities",
            counters,
            labels,
        )
        self.assertIn("# 理论与框架 {#sec:theory}", converted)
        self.assertIn("## 标准需求", converted)
        self.assertIn("### 对象边界", converted)
        self.assertNotIn("一、理论与框架", converted)
        self.assertNotIn("（一）标准需求", converted)
        self.assertEqual(labels.targets["sec:theory"].number, "一、")

    def test_heading_styles_use_one_native_multilevel_list(self) -> None:
        document = BUILD_WORD.Document()
        for style_name in (
            BUILD_WORD.STYLE_HEADING_1,
            BUILD_WORD.STYLE_HEADING_2,
            BUILD_WORD.STYLE_HEADING_3,
            BUILD_WORD.STYLE_HEADING_4,
        ):
            document.styles.add_style(
                style_name,
                BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH,
            )
        BUILD_WORD._configure_heading_numbering(document, "humanities")
        num_ids = []
        for level, style_name in enumerate(
            (
                BUILD_WORD.STYLE_HEADING_1,
                BUILD_WORD.STYLE_HEADING_2,
                BUILD_WORD.STYLE_HEADING_3,
                BUILD_WORD.STYLE_HEADING_4,
            )
        ):
            num_pr = document.styles[style_name]._element.find(
                BUILD_WORD.qn("w:pPr") + "/" + BUILD_WORD.qn("w:numPr")
            )
            self.assertIsNotNone(num_pr)
            self.assertEqual(
                num_pr.find(BUILD_WORD.qn("w:ilvl")).get(
                    BUILD_WORD.qn("w:val")
                ),
                str(level),
            )
            num_ids.append(
                num_pr.find(BUILD_WORD.qn("w:numId")).get(
                    BUILD_WORD.qn("w:val")
                )
            )
        self.assertEqual(len(set(num_ids)), 1)
        abstract = document.part.numbering_part.element.findall(
            BUILD_WORD.qn("w:abstractNum")
        )[-1]
        self.assertEqual(
            [
                level.find(BUILD_WORD.qn("w:numFmt")).get(
                    BUILD_WORD.qn("w:val")
                )
                for level in abstract.findall(BUILD_WORD.qn("w:lvl"))
            ],
            ["chineseCounting", "chineseCounting", "decimal", "decimal"],
        )

    def test_longtable_continuation_header_is_removed(self) -> None:
        source = r"""
\begin{longtable}{C{0.20\textwidth}L{0.30\textwidth}R{0.40\textwidth}}
名称 & 数量 & 说明 \\
\midrule
\endfirsthead
\multicolumn{3}{r}{续表~\thetable} \\
名称 & 数量 & 说明 \\
\midrule
\endhead
对象A & 1 & 正文 \\
\end{longtable}
"""
        normalized = BUILD_WORD.normalize_longtable_headers(source)
        self.assertEqual(normalized.count("名称 & 数量 & 说明"), 1)
        self.assertNotIn("续表", normalized)
        self.assertNotIn(r"\endfirsthead", normalized)
        self.assertNotIn(r"\endhead", normalized)
        self.assertIn("对象A & 1 & 正文", normalized)
        self.assertIn(r"\begin{longtable}{clr}", normalized)

    def test_complex_column_spec_is_counted_before_simplification(self) -> None:
        specification = (
            r">{\centering\arraybackslash}m{0.05\textwidth}"
            r">{\raggedright\arraybackslash}m{0.35\textwidth}"
            r">{\centering\arraybackslash}m{0.14\textwidth}"
            r">{\centering\arraybackslash}m{0.10\textwidth}"
            r">{\centering\arraybackslash}m{0.07\textwidth}"
            r">{\centering\arraybackslash}m{0.14\textwidth}"
        )
        self.assertEqual(BUILD_WORD.count_latex_table_columns(specification), 6)

    def test_table_layout_preserves_latex_width_and_alignment(self) -> None:
        specification = (
            r"C{0.13\textwidth}"
            r"L{0.24\textwidth}"
            r"L{0.20\textwidth}"
            r">{\raggedright\arraybackslash}X"
        )
        columns = BUILD_WORD.parse_latex_table_columns(specification)

        self.assertEqual(
            [column.horizontal for column in columns],
            ["center", "left", "left", "left"],
        )
        self.assertEqual(
            [column.vertical for column in columns],
            ["center", "center", "center", "center"],
        )
        self.assertEqual(
            [column.width_fraction for column in columns],
            [0.13, 0.24, 0.20, None],
        )
        self.assertTrue(columns[-1].flexible)

        widths = BUILD_WORD._resolve_table_column_widths(
            BUILD_WORD.LatexTableLayout("tabularx", tuple(columns)),
            [1, 1, 1, 1],
            10_000,
        )
        self.assertEqual(sum(widths), 10_000)
        self.assertEqual(widths, [1300, 2400, 2000, 4300])

    def test_adjacent_custom_columns_remain_pandoc_readable(self) -> None:
        normalized = BUILD_WORD.normalize_latex(
            r"\begin{tabularx}{\textwidth}{L{.2\textwidth}C{.3\textwidth}R{.4\textwidth}}",
            BUILD_WORD.NoteRegistry(),
        )
        self.assertIn(r"\begin{tabularx}{\textwidth}{lcr}", normalized)

    def test_table_preamble_normalization_does_not_rewrite_prose(self) -> None:
        normalized = BUILD_WORD.normalize_latex(
            "集合 C{60} 与 R{2}。\n"
            r"\begin{tabular}{C{.4\textwidth}L{.5\textwidth}}a&b\\\end{tabular}",
            BUILD_WORD.NoteRegistry(),
        )
        self.assertIn("集合 C{60} 与 R{2}。", normalized)
        self.assertIn(r"\begin{tabular}{cl}", normalized)

    def test_tabularx_optional_position_and_label_are_preserved(self) -> None:
        layouts = BUILD_WORD.extract_latex_table_layouts(
            r"""
\begin{table}[htbp]
\caption{测试}\label{tab:optional-position}
\begin{tabularx}{.8\textwidth}[t]{lX}
a&b\\
\end{tabularx}
\end{table}
"""
        )
        self.assertEqual(len(layouts), 1)
        self.assertEqual(layouts[0].label, "tab:optional-position")
        self.assertEqual(layouts[0].table_width_fraction, 0.8)
        self.assertEqual(len(layouts[0].columns), 2)

    def test_siunitx_options_do_not_create_phantom_columns(self) -> None:
        columns = BUILD_WORD.parse_latex_table_columns("S[table-format=3.2]")
        self.assertEqual(len(columns), 1)
        self.assertEqual(columns[0].horizontal, "right")

    def test_unknown_column_type_fails_instead_of_guessing(self) -> None:
        with self.assertRaises(BUILD_WORD.BuildError):
            BUILD_WORD.parse_latex_table_columns(r"M{2cm}")

    def test_multiline_longtable_preamble_is_simplified(self) -> None:
        normalized = BUILD_WORD.normalize_longtable_headers(
            r"""
\begin{longtable}{
  C{.2\textwidth}
  L{.7\textwidth}
}
A&B\\
\endfirsthead
A&B\\
\endhead
x&y\\
\end{longtable}
"""
        )
        self.assertRegex(normalized, r"\\begin\{longtable\}\{\s*cl\}")
        self.assertNotIn(r"\endfirsthead", normalized)

    def test_explicit_table_parindent_is_source_derived(self) -> None:
        columns = BUILD_WORD.parse_latex_table_columns(
            r">{\setlength{\parindent}{2em}}p{.5\textwidth}"
        )
        self.assertEqual(columns[0].first_line_indent_pt, 21.0)

    def test_longtable_layout_is_captured_before_normalization(self) -> None:
        layouts: list[BUILD_WORD.LatexTableLayout] = []
        normalized = BUILD_WORD.normalize_latex(
            r"""
\begin{longtable}{L{.3\textwidth}>{\centering\arraybackslash\setlength{\parindent}{2em}}m{.6\textwidth}}
A&B\\
\endfirsthead
A&B\\
\endhead
x&y\\
\end{longtable}
""",
            BUILD_WORD.NoteRegistry(),
            table_layouts=layouts,
        )

        self.assertEqual(len(layouts), 1)
        self.assertEqual(
            [column.horizontal for column in layouts[0].columns],
            ["left", "center"],
        )
        self.assertEqual(
            [column.first_line_indent_pt for column in layouts[0].columns],
            [None, 21.0],
        )
        self.assertNotIn(r"\endfirsthead", normalized)
        self.assertNotIn(r"\endhead", normalized)

    def test_table_tabcolsep_is_source_derived_and_applied(self) -> None:
        layouts = BUILD_WORD.extract_latex_table_layouts(
            r"""
\begin{table}
\setlength{\tabcolsep}{2.5pt}
\begin{tabular}{lc}
a&b\\
\end{tabular}
\end{table}
"""
        )
        self.assertEqual(len(layouts), 1)
        self.assertEqual(layouts[0].tabcolsep_pt, 2.5)

        document = BUILD_WORD.Document()
        for name in (
            BUILD_WORD.STYLE_TABLE_TEXT,
            BUILD_WORD.STYLE_TABLE_CENTER,
            BUILD_WORD.STYLE_TABLE_RIGHT,
            BUILD_WORD.STYLE_TABLE_HEADER,
            BUILD_WORD.STYLE_TABLE_HEADER_LEFT,
            BUILD_WORD.STYLE_TABLE_HEADER_RIGHT,
        ):
            document.styles.add_style(name, BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH)
        document.styles.add_style(
            BUILD_WORD.STYLE_THREE_LINE_TABLE,
            BUILD_WORD.WD_STYLE_TYPE.TABLE,
        )
        table = document.add_table(rows=2, cols=2)
        BUILD_WORD._format_tables(document, layouts)

        margins = table._tbl.tblPr.find(BUILD_WORD.qn("w:tblCellMar"))
        self.assertIsNotNone(margins)
        self.assertEqual(
            margins.find(BUILD_WORD.qn("w:left")).get(BUILD_WORD.qn("w:w")),
            "50",
        )
        self.assertEqual(
            margins.find(BUILD_WORD.qn("w:right")).get(BUILD_WORD.qn("w:w")),
            "50",
        )

    def test_native_paragraph_columns_control_vertical_alignment(self) -> None:
        columns = BUILD_WORD.parse_latex_table_columns(
            r">{\centering\arraybackslash}p{2cm}m{3cm}b{4cm}"
        )
        self.assertEqual(
            [(column.horizontal, column.vertical) for column in columns],
            [("center", "top"), ("left", "center"), ("left", "bottom")],
        )

    def test_all_supported_table_environments_are_collected_in_order(self) -> None:
        layouts = BUILD_WORD.extract_latex_table_layouts(
            r"""
\begin{tabular}{lrr}a&b&c\\\end{tabular}
\begin{tabularx}{\textwidth}{C{.2\textwidth}X}a&b\\\end{tabularx}
\begin{longtable}{L{.3\textwidth}R{.6\textwidth}}a&b\\\end{longtable}
"""
        )
        self.assertEqual(
            [layout.environment for layout in layouts],
            ["tabular", "tabularx", "longtable"],
        )
        self.assertEqual(
            [[column.horizontal for column in layout.columns] for layout in layouts],
            [["left", "right", "right"], ["center", "left"], ["left", "right"]],
        )

    def test_word_table_styles_follow_registered_latex_columns(self) -> None:
        document = BUILD_WORD.Document()
        paragraph_styles = (
            BUILD_WORD.STYLE_TABLE_TEXT,
            BUILD_WORD.STYLE_TABLE_CENTER,
            BUILD_WORD.STYLE_TABLE_RIGHT,
            BUILD_WORD.STYLE_TABLE_HEADER,
            BUILD_WORD.STYLE_TABLE_HEADER_LEFT,
            BUILD_WORD.STYLE_TABLE_HEADER_RIGHT,
        )
        for name in paragraph_styles:
            document.styles.add_style(name, BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH)
        document.styles.add_style(
            BUILD_WORD.STYLE_THREE_LINE_TABLE,
            BUILD_WORD.WD_STYLE_TYPE.TABLE,
        )
        table = document.add_table(rows=2, cols=3)
        for row in table.rows:
            for cell in row.cells:
                cell.text = "值"

        columns = BUILD_WORD.parse_latex_table_columns(
            r">{\centering\arraybackslash}p{.2\textwidth}m{.3\textwidth}b{.5\textwidth}"
        )
        BUILD_WORD._format_tables(
            document,
            [BUILD_WORD.LatexTableLayout("tabular", tuple(columns))],
        )

        self.assertEqual(
            [cell.paragraphs[0].style.name for cell in table.rows[0].cells],
            [
                BUILD_WORD.STYLE_TABLE_HEADER,
                BUILD_WORD.STYLE_TABLE_HEADER_LEFT,
                BUILD_WORD.STYLE_TABLE_HEADER_LEFT,
            ],
        )
        self.assertEqual(
            [cell.paragraphs[0].style.name for cell in table.rows[1].cells],
            [
                BUILD_WORD.STYLE_TABLE_CENTER,
                BUILD_WORD.STYLE_TABLE_TEXT,
                BUILD_WORD.STYLE_TABLE_TEXT,
            ],
        )
        self.assertEqual(
            [cell.vertical_alignment for cell in table.rows[1].cells],
            [
                BUILD_WORD.WD_CELL_VERTICAL_ALIGNMENT.TOP,
                BUILD_WORD.WD_CELL_VERTICAL_ALIGNMENT.CENTER,
                BUILD_WORD.WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
            ],
        )
        widths = [
            int(column.get(BUILD_WORD.qn("w:w")))
            for column in table._tbl.tblGrid.findall(BUILD_WORD.qn("w:gridCol"))
        ]
        self.assertEqual(widths, [1757, 2636, 4394])

    def test_table_cells_are_style_driven_but_keep_explicit_latex_indent(self) -> None:
        document = BUILD_WORD.Document()
        for name in (
            BUILD_WORD.STYLE_TABLE_TEXT,
            BUILD_WORD.STYLE_TABLE_CENTER,
            BUILD_WORD.STYLE_TABLE_RIGHT,
            BUILD_WORD.STYLE_TABLE_HEADER,
            BUILD_WORD.STYLE_TABLE_HEADER_LEFT,
            BUILD_WORD.STYLE_TABLE_HEADER_RIGHT,
        ):
            document.styles.add_style(name, BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH)
        document.styles.add_style(
            BUILD_WORD.STYLE_THREE_LINE_TABLE,
            BUILD_WORD.WD_STYLE_TYPE.TABLE,
        )
        table = document.add_table(rows=2, cols=2)
        for row in table.rows:
            for cell in row.cells:
                cell.text = "值"

        default_paragraph = table.rows[1].cells[0].paragraphs[0]
        explicit_paragraph = table.rows[1].cells[1].paragraphs[0]
        default_paragraph.paragraph_format.first_line_indent = BUILD_WORD.Pt(24)
        default_paragraph.alignment = BUILD_WORD.WD_ALIGN_PARAGRAPH.RIGHT
        explicit_paragraph.paragraph_format.first_line_indent = BUILD_WORD.Pt(24)
        explicit_paragraph.alignment = BUILD_WORD.WD_ALIGN_PARAGRAPH.RIGHT

        columns = BUILD_WORD.parse_latex_table_columns(
            r"l>{\centering\arraybackslash\setlength{\parindent}{2em}}p{.5\textwidth}"
        )
        BUILD_WORD._format_tables(
            document,
            [BUILD_WORD.LatexTableLayout("tabular", tuple(columns))],
        )

        self.assertEqual(
            default_paragraph.style.name,
            BUILD_WORD.STYLE_TABLE_TEXT,
        )
        self.assertIsNone(
            default_paragraph._p.get_or_add_pPr().find(BUILD_WORD.qn("w:ind"))
        )
        self.assertIsNone(
            default_paragraph._p.get_or_add_pPr().find(BUILD_WORD.qn("w:jc"))
        )
        self.assertEqual(
            explicit_paragraph.style.name,
            BUILD_WORD.STYLE_TABLE_CENTER,
        )
        self.assertAlmostEqual(
            explicit_paragraph.paragraph_format.first_line_indent.pt,
            21.0,
        )
        self.assertIsNone(
            explicit_paragraph._p.get_or_add_pPr().find(BUILD_WORD.qn("w:jc"))
        )

    def test_merged_cell_alignment_overrides_column_default(self) -> None:
        document = BUILD_WORD.Document()
        for name in (
            BUILD_WORD.STYLE_TABLE_TEXT,
            BUILD_WORD.STYLE_TABLE_CENTER,
            BUILD_WORD.STYLE_TABLE_RIGHT,
            BUILD_WORD.STYLE_TABLE_HEADER,
            BUILD_WORD.STYLE_TABLE_HEADER_LEFT,
            BUILD_WORD.STYLE_TABLE_HEADER_RIGHT,
        ):
            document.styles.add_style(name, BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH)
        document.styles.add_style(
            BUILD_WORD.STYLE_THREE_LINE_TABLE,
            BUILD_WORD.WD_STYLE_TYPE.TABLE,
        )
        table = document.add_table(rows=2, cols=2)
        merged = table.rows[1].cells[0].merge(table.rows[1].cells[1])
        merged.paragraphs[0].alignment = BUILD_WORD.WD_ALIGN_PARAGRAPH.CENTER
        layout = BUILD_WORD.LatexTableLayout(
            "tabular",
            tuple(BUILD_WORD.parse_latex_table_columns("ll")),
        )
        BUILD_WORD._format_tables(document, [layout])
        self.assertEqual(
            merged.paragraphs[0].style.name,
            BUILD_WORD.STYLE_TABLE_CENTER,
        )

    def test_table_labels_prevent_order_based_layout_drift(self) -> None:
        document = BUILD_WORD.Document()
        for name in (
            BUILD_WORD.STYLE_TABLE_TEXT,
            BUILD_WORD.STYLE_TABLE_CENTER,
            BUILD_WORD.STYLE_TABLE_RIGHT,
            BUILD_WORD.STYLE_TABLE_HEADER,
            BUILD_WORD.STYLE_TABLE_HEADER_LEFT,
            BUILD_WORD.STYLE_TABLE_HEADER_RIGHT,
        ):
            document.styles.add_style(name, BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH)
        document.styles.add_style(
            BUILD_WORD.STYLE_THREE_LINE_TABLE,
            BUILD_WORD.WD_STYLE_TYPE.TABLE,
        )
        document.add_paragraph("表1 左表")
        left_table = document.add_table(rows=2, cols=2)
        document.add_paragraph("表2 中表")
        center_table = document.add_table(rows=2, cols=2)
        labels = BUILD_WORD.LabelRegistry()
        labels.targets["tab:left"] = BUILD_WORD.LabelTarget("1", "表")
        labels.targets["tab:center"] = BUILD_WORD.LabelTarget("2", "表")
        layouts = [
            BUILD_WORD.LatexTableLayout(
                "tabular",
                tuple(BUILD_WORD.parse_latex_table_columns("cc")),
                label="tab:center",
            ),
            BUILD_WORD.LatexTableLayout(
                "tabular",
                tuple(BUILD_WORD.parse_latex_table_columns("ll")),
                label="tab:left",
            ),
        ]
        BUILD_WORD._format_tables(document, layouts, labels)
        self.assertEqual(
            left_table.rows[1].cells[0].paragraphs[0].style.name,
            BUILD_WORD.STYLE_TABLE_TEXT,
        )
        self.assertEqual(
            center_table.rows[1].cells[0].paragraphs[0].style.name,
            BUILD_WORD.STYLE_TABLE_CENTER,
        )

    def test_citation_cluster_is_sorted_by_bibliography_number(self) -> None:
        document = BUILD_WORD.Document()
        document.styles.add_style(
            BUILD_WORD.STYLE_CITATION,
            BUILD_WORD.WD_STYLE_TYPE.CHARACTER,
        )
        paragraph = document.add_paragraph("YIBINCITE00000001")
        citations = BUILD_WORD.CitationRegistry(
            clusters={"YIBINCITE00000001": ["second", "first"]},
            order=["first", "second"],
            next_cluster=1,
        )
        BUILD_WORD._replace_reference_markers(
            document,
            BUILD_WORD.LabelRegistry(),
            citations,
            "linked",
        )
        self.assertEqual(paragraph.text, "[1,2]")

    def test_science_equation_uses_native_word_sequence_and_full_bookmark(
        self,
    ) -> None:
        document = BUILD_WORD.Document()
        paragraph = document.add_paragraph()
        paragraph._p.append(BUILD_WORD.OxmlElement("m:oMath"))
        labels = BUILD_WORD.LabelRegistry()
        labels.register("eq:native", "3.1", "公式")

        BUILD_WORD._format_equation_number(
            paragraph,
            "3.1",
            labels,
            discipline="science",
            equation_sequence="公式",
        )

        instructions = self._field_instructions(paragraph)
        self.assertIn(
            f' STYLEREF "{BUILD_WORD.STYLE_HEADING_1}" \\n ',
            instructions,
        )
        self.assertIn(" SEQ 公式 \\r 1 \\* ARABIC ", instructions)

        bookmark_name = BUILD_WORD._label_bookmark_name("eq:native")
        children = list(paragraph._p)
        start_index = next(
            index
            for index, child in enumerate(children)
            if child.tag == BUILD_WORD.qn("w:bookmarkStart")
            and child.get(BUILD_WORD.qn("w:name")) == bookmark_name
        )
        bookmark_id = children[start_index].get(BUILD_WORD.qn("w:id"))
        end_index = next(
            index
            for index, child in enumerate(children)
            if child.tag == BUILD_WORD.qn("w:bookmarkEnd")
            and child.get(BUILD_WORD.qn("w:id")) == bookmark_id
        )
        bookmarked_text = "".join(
            node.text or ""
            for child in children[start_index + 1 : end_index]
            for node in child.iter(BUILD_WORD.qn("w:t"))
        )
        self.assertEqual(bookmarked_text, "(3.1)")

    def test_humanities_equation_keeps_arabic_chapter_with_native_sequence(
        self,
    ) -> None:
        document = BUILD_WORD.Document()
        paragraph = document.add_paragraph()
        paragraph._p.append(BUILD_WORD.OxmlElement("m:oMath"))

        BUILD_WORD._format_equation_number(
            paragraph,
            "2.4",
            discipline="humanities",
            equation_sequence="公式",
        )

        instructions = self._field_instructions(paragraph)
        self.assertNotIn("STYLEREF", " ".join(instructions))
        self.assertIn(" SEQ 公式 \\* ARABIC ", instructions)
        rendered = "".join(
            node.text or "" for node in paragraph._p.iter(BUILD_WORD.qn("w:t"))
        )
        self.assertEqual(rendered, "(2.4)")

    def test_caption_uses_registered_sequence_and_hidden_native_bookmarks(self) -> None:
        document = BUILD_WORD.Document()
        document.styles.add_style(
            BUILD_WORD.STYLE_FIGURE_CAPTION,
            BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH,
        )
        document.styles.add_style(
            BUILD_WORD.STYLE_TABLE_CAPTION,
            BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH,
        )
        figure_paragraph = document.add_paragraph(
            "图1 测试图",
            style=BUILD_WORD.STYLE_FIGURE_CAPTION,
        )
        table_paragraph = document.add_paragraph(
            "表1 测试表",
            style=BUILD_WORD.STYLE_TABLE_CAPTION,
        )
        labels = BUILD_WORD.LabelRegistry()
        labels.register("fig:test", "1", "图")
        labels.register("tab:test", "1", "表")

        BUILD_WORD._promote_caption_fields(
            document,
            labels,
            "humanities",
            figure_sequence="图",
            table_sequence="表",
        )

        self.assertIn(
            " SEQ 图 \\r 1 \\* ARABIC ",
            self._field_instructions(figure_paragraph),
        )
        self.assertIn(
            " SEQ 表 \\r 1 \\* ARABIC ",
            self._field_instructions(table_paragraph),
        )
        figure_names = [
            node.get(BUILD_WORD.qn("w:name"))
            for node in figure_paragraph._p.iter(BUILD_WORD.qn("w:bookmarkStart"))
        ]
        table_names = [
            node.get(BUILD_WORD.qn("w:name"))
            for node in table_paragraph._p.iter(BUILD_WORD.qn("w:bookmarkStart"))
        ]
        self.assertIn(BUILD_WORD._label_bookmark_name("fig:test"), figure_names)
        self.assertIn(
            BUILD_WORD._label_number_bookmark_name("fig:test"), figure_names
        )
        self.assertIn(BUILD_WORD._label_bookmark_name("tab:test"), table_names)
        self.assertIn(
            BUILD_WORD._label_number_bookmark_name("tab:test"), table_names
        )
        self.assertTrue(
            all(name.startswith("_") for name in figure_names + table_names)
        )

    def test_caption_reference_is_plain_ref_field_without_outer_hyperlink(self) -> None:
        document = BUILD_WORD.Document()
        document.styles.add_style(
            BUILD_WORD.STYLE_CITATION,
            BUILD_WORD.WD_STYLE_TYPE.CHARACTER,
        )
        paragraph = document.add_paragraph("见图YIBINXREF00000001。")
        labels = BUILD_WORD.LabelRegistry(
            targets={"fig:test": BUILD_WORD.LabelTarget("1", "图")},
            references={"YIBINXREF00000001": ("fig:test", False)},
        )
        BUILD_WORD._replace_reference_markers(
            document,
            labels,
            BUILD_WORD.CitationRegistry(),
            "linked",
        )

        self.assertEqual(paragraph.text, "见图1。")
        self.assertEqual(
            list(paragraph._p.iter(BUILD_WORD.qn("w:hyperlink"))),
            [],
        )
        self.assertIn(
            f" REF {BUILD_WORD._label_bookmark_name('fig:test')} \\h ",
            self._field_instructions(paragraph),
        )

    def test_pdf_figure_prefers_existing_png_sibling(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            assets = root / "assets"
            assets.mkdir()
            (assets / "figure.pdf").write_bytes(b"pdf")
            (assets / "figure.png").write_bytes(b"png")

            self.assertEqual(
                BUILD_WORD.resolve_word_image_target(
                    "assets/figure.pdf",
                    [root],
                ),
                "assets/figure.png",
            )

    def test_embed_is_converted_to_numbered_markdown_image(self) -> None:
        labels = BUILD_WORD.LabelRegistry()
        converted = BUILD_WORD.apply_float_numbering(
            """<figure id="fig:test">
<embed src="assets/figure.pdf" style="width:80%" />
<figcaption>测试图</figcaption>
</figure>""",
            "humanities",
            1,
            BUILD_WORD.FloatCounters(),
            labels,
            lambda target: target.replace(".pdf", ".png"),
        )
        self.assertIn("![](assets/figure.png){width=80%}", converted)
        self.assertIn("图1 测试图", converted)
        self.assertEqual(labels.targets["fig:test"].number, "1")


class BackmatterOrderTests(unittest.TestCase):
    def test_template_and_example_follow_school_backmatter_order(self) -> None:
        cases = (
            (
                ROOT / "main.tex",
                (
                    r"\input{chapters/99-acknowledgements}",
                    r"\printyibinnotes",
                    r"\printyibinbibliography",
                    r"\input{chapters/90-appendix}",
                ),
            ),
            (
                ROOT / "examples/full-featured/main.tex",
                (
                    r"\input{examples/full-featured/chapters/90-acknowledgements}",
                    r"\printyibinnotes",
                    r"\printyibinbibliography",
                    r"\input{examples/full-featured/chapters/99-appendix}",
                ),
            ),
        )
        for path, markers in cases:
            with self.subTest(path=path):
                source = path.read_text(encoding="utf-8")
                positions = [source.index(marker) for marker in markers]
                self.assertEqual(positions, sorted(positions))


class GradeClassFormattingTests(unittest.TestCase):
    def test_legacy_numeric_grade_keeps_implicit_suffix(self) -> None:
        self.assertEqual(BUILD_WORD.format_grade_class({"grade": "2022"}), "2022级")

    def test_class_name_is_appended_after_implicit_grade_suffix(self) -> None:
        self.assertEqual(
            BUILD_WORD.format_grade_class({"grade": "2023", "class-name": "7班"}),
            "2023级7班",
        )

    def test_complete_grade_text_is_not_duplicated(self) -> None:
        self.assertEqual(
            BUILD_WORD.format_grade_class({"grade": "2023级7班"}),
            "2023级7班",
        )

    def test_latex_cover_uses_shared_grade_class_field(self) -> None:
        source = (ROOT / "yibinthesis.cls").read_text(encoding="utf-8")
        self.assertIn("class-name", source)
        self.assertIn(r"{\yibingradeclass}", source)


if __name__ == "__main__":
    unittest.main()
