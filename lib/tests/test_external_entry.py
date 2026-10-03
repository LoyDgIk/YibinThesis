from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location(
    "yibinthesis_build_word", ROOT / "lib" / "build_word.py"
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

    @staticmethod
    def _add_table_styles(document) -> None:
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

    @staticmethod
    def _grid_widths(table) -> list[int]:
        return [
            int(column.get(BUILD_WORD.qn("w:w")))
            for column in table._tbl.tblGrid.findall(BUILD_WORD.qn("w:gridCol"))
        ]

    @staticmethod
    def _make_ahp_table(document, left: str = "C42废弃物处理与循环利用", right: str = "C44清洁能源与节能资源"):
        table = document.add_table(rows=3, cols=19)
        header = table.rows[0].cells
        header[0].text = "左侧指标"
        header[18].text = "右侧指标"
        header[1].merge(header[17]).text = "评价尺度"
        for cell, value in zip(table.rows[1].cells[1:18], BUILD_WORD._AHP_SCALE_VALUES):
            cell.text = value
        table.rows[2].cells[0].text = left
        table.rows[2].cells[18].text = right
        return table

    @staticmethod
    def _make_rating_meaning_table(document):
        table = document.add_table(rows=2, cols=3)
        for cell, value in zip(table.rows[0].cells, ("等级", "分值", "一般含义")):
            cell.text = value
        for cell, value in zip(table.rows[1].cells, ("优", "5", "表现突出，能够持续保持")):
            cell.text = value
        return table

    @staticmethod
    def _make_fce_table(document):
        table = document.add_table(rows=2, cols=7)
        header = ("编码", "评价指标", "优（5）", "良（4）", "中（3）", "较差（2）", "差（1）")
        for cell, value in zip(table.rows[0].cells, header):
            cell.text = value
        table.rows[1].cells[0].text = "C42"
        table.rows[1].cells[1].text = "农业废弃物规范处理与循环利用程度"
        return table

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

    def test_native_list_paragraph_uses_reusable_style_without_losing_numbering(self) -> None:
        document = BUILD_WORD.Document()
        document.styles.add_style(
            BUILD_WORD.STYLE_BODY,
            BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH,
        )
        document.styles.add_style(
            BUILD_WORD.STYLE_LIST_BODY,
            BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH,
        )

        listed = document.add_paragraph("研究问题一", style=BUILD_WORD.STYLE_BODY)
        listed_ppr = listed._p.get_or_add_pPr()
        listed_indent = BUILD_WORD.OxmlElement("w:ind")
        listed_indent.set(BUILD_WORD.qn("w:left"), "720")
        listed_indent.set(BUILD_WORD.qn("w:hanging"), "360")
        listed_indent.set(BUILD_WORD.qn("w:firstLine"), "480")
        listed_indent.set(BUILD_WORD.qn("w:firstLineChars"), "200")
        listed_ppr.append(listed_indent)
        BUILD_WORD._set_paragraph_numbering(listed, num_id=7, level=0)

        ordinary = document.add_paragraph("普通正文", style=BUILD_WORD.STYLE_BODY)

        BUILD_WORD._apply_list_paragraph_styles(document)

        self.assertEqual(listed.style.name, BUILD_WORD.STYLE_LIST_BODY)
        numbered_ppr = listed._p.get_or_add_pPr()
        num_pr = numbered_ppr.find(BUILD_WORD.qn("w:numPr"))
        self.assertIsNotNone(num_pr)
        self.assertEqual(
            num_pr.find(BUILD_WORD.qn("w:numId")).get(BUILD_WORD.qn("w:val")),
            "7",
        )
        self.assertEqual(
            num_pr.find(BUILD_WORD.qn("w:ilvl")).get(BUILD_WORD.qn("w:val")),
            "0",
        )
        preserved_indent = numbered_ppr.find(BUILD_WORD.qn("w:ind"))
        self.assertEqual(preserved_indent.get(BUILD_WORD.qn("w:left")), "720")
        self.assertEqual(preserved_indent.get(BUILD_WORD.qn("w:hanging")), "360")
        self.assertIsNone(preserved_indent.get(BUILD_WORD.qn("w:firstLine")))
        self.assertIsNone(preserved_indent.get(BUILD_WORD.qn("w:firstLineChars")))
        self.assertEqual(ordinary.style.name, BUILD_WORD.STYLE_BODY)
        self.assertIsNone(
            ordinary._p.get_or_add_pPr().find(BUILD_WORD.qn("w:numPr"))
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

    def test_longtable_page_footers_are_removed(self) -> None:
        source = r"""
\begin{longtable}{C{0.20\textwidth}L{0.30\textwidth}R{0.40\textwidth}}
名称 & 数量 & 说明 \\
\midrule
\endfirsthead
\multicolumn{3}{c}{续表\thetable} \\
名称 & 数量 & 说明 \\
\midrule
\endhead
\midrule
\multicolumn{3}{r}{续下页} \\
\endfoot
\bottomrule
\endlastfoot
对象A & 1 & 正文 \\
对象B & 2 & 正文 \\
\end{longtable}
"""
        normalized = BUILD_WORD.normalize_longtable_headers(source)
        self.assertEqual(normalized.count("名称 & 数量 & 说明"), 1)
        self.assertNotIn("续表", normalized)
        self.assertNotIn("续下页", normalized)
        self.assertNotIn(r"\endfoot", normalized)
        self.assertNotIn(r"\endlastfoot", normalized)
        self.assertIn("对象A & 1 & 正文", normalized)
        self.assertIn("对象B & 2 & 正文", normalized)

    def test_longtable_endfoot_without_endlastfoot_is_removed(self) -> None:
        source = r"""
\begin{longtable}{cc}
名称 & 数量 \\
\endfirsthead
名称 & 数量 \\
\endhead
\multicolumn{2}{r}{续下页} \\
\endfoot
对象A & 1 \\
\end{longtable}
"""
        normalized = BUILD_WORD.normalize_longtable_headers(source)
        self.assertEqual(normalized.count("名称 & 数量"), 1)
        self.assertNotIn("续下页", normalized)
        self.assertNotIn(r"\endfoot", normalized)
        self.assertIn("对象A & 1", normalized)

    def test_resizebox_table_is_unwrapped_for_pandoc(self) -> None:
        source = r"""
\begin{table}[htbp]
\caption{测试表}
\label{tab:resizebox}
\resizebox{\textwidth}{!}{%
  \begin{tabular}{clrr}
  编码 & 名称 & 权重 & 排序 \\
  B1 & 资源共享盘活 & 0.2 & 1 \\
  \end{tabular}%
}
\end{table}
\resizebox{2cm}{!}{\includegraphics{logo.png}}
"""
        layouts: list[BUILD_WORD.LatexTableLayout] = []
        normalized = BUILD_WORD.normalize_latex(
            source,
            BUILD_WORD.NoteRegistry(),
            table_layouts=layouts,
        )

        self.assertNotIn(r"\resizebox{\textwidth}{!}", normalized)
        self.assertIn(r"\begin{tabular}{clrr}", normalized)
        self.assertIn(r"\label{tab:resizebox}", normalized)
        self.assertIn(r"\resizebox{2cm}{!}{\includegraphics{logo.png}}", normalized)
        self.assertEqual(len(layouts), 1)
        self.assertEqual(layouts[0].label, "tab:resizebox")

    def test_appendix_label_is_registered_before_pandoc(self) -> None:
        labels = BUILD_WORD.LabelRegistry()
        normalized = BUILD_WORD.normalize_latex(
            r"""
\begin{yibinappendices}
\yibinappendix{共享农庄综合评价专家问卷}
\label{app:expert-questionnaire}
问卷正文。
\end{yibinappendices}
""",
            BUILD_WORD.NoteRegistry(),
            labels=labels,
        )

        self.assertIn(r"\chapter*{附录A 共享农庄综合评价专家问卷}", normalized)
        self.assertNotIn(r"\label{app:expert-questionnaire}", normalized)
        self.assertEqual(labels.targets["app:expert-questionnaire"].number, "A")
        self.assertEqual(labels.targets["app:expert-questionnaire"].kind, "附录")

    def test_appendix_sections_leave_numbering_to_native_word_lists(self) -> None:
        source = r"""
\begin{yibinappendices}
\yibinappendix{第一组}
\yibinappendixsection{边界说明}
\yibinappendixsection{参数说明}
\yibinappendix{第二组}
\yibinappendixsection{代码说明}
\end{yibinappendices}
"""
        humanities = BUILD_WORD.normalize_latex(
            source,
            BUILD_WORD.NoteRegistry(),
            discipline="humanities",
        )
        science = BUILD_WORD.normalize_latex(
            source,
            BUILD_WORD.NoteRegistry(),
            discipline="science",
        )

        for normalized in (humanities, science):
            self.assertIn(r"\section*{边界说明}", normalized)
            self.assertIn(r"\section*{参数说明}", normalized)
            self.assertIn(r"\section*{代码说明}", normalized)
            self.assertNotIn(r"\section*{（一）", normalized)
            self.assertNotIn(r"\section*{A.1", normalized)

    def test_appendix_sections_use_native_numbering_and_restart(self) -> None:
        for discipline, expected_labels in (
            ("humanities", ("（%1）", "（%1）")),
            ("science", ("A.%1", "B.%1")),
        ):
            with self.subTest(discipline=discipline):
                document = BUILD_WORD.Document()
                for style_name in (
                    BUILD_WORD.STYLE_HEADING_2,
                    BUILD_WORD.STYLE_APPENDIX_HEADING,
                    BUILD_WORD.STYLE_APPENDIX_SECTION,
                ):
                    document.styles.add_style(
                        style_name,
                        BUILD_WORD.WD_STYLE_TYPE.PARAGRAPH,
                    )
                ordinary = document.add_paragraph(
                    "正文二级标题",
                    style=BUILD_WORD.STYLE_HEADING_2,
                )
                document.add_paragraph(
                    "附录A 第一组",
                    style=BUILD_WORD.STYLE_APPENDIX_HEADING,
                )
                first = document.add_paragraph(
                    "边界说明",
                    style=BUILD_WORD.STYLE_HEADING_2,
                )
                second = document.add_paragraph(
                    "参数说明",
                    style=BUILD_WORD.STYLE_HEADING_2,
                )
                document.add_paragraph(
                    "附录B 第二组",
                    style=BUILD_WORD.STYLE_APPENDIX_HEADING,
                )
                restarted = document.add_paragraph(
                    "代码说明",
                    style=BUILD_WORD.STYLE_HEADING_2,
                )

                BUILD_WORD._configure_appendix_section_numbering(
                    document,
                    discipline,
                )

                self.assertEqual(ordinary.style.name, BUILD_WORD.STYLE_HEADING_2)
                self.assertIsNone(
                    ordinary._p.get_or_add_pPr().find(BUILD_WORD.qn("w:numPr"))
                )
                for paragraph in (first, second, restarted):
                    self.assertEqual(
                        paragraph.style.name,
                        BUILD_WORD.STYLE_APPENDIX_SECTION,
                    )
                    self.assertIsNotNone(
                        paragraph._p.get_or_add_pPr().find(
                            BUILD_WORD.qn("w:numPr")
                        )
                    )
                num_ids = []
                for paragraph in (first, second, restarted):
                    num_pr = paragraph._p.get_or_add_pPr().find(
                        BUILD_WORD.qn("w:numPr")
                    )
                    num_ids.append(
                        num_pr.find(BUILD_WORD.qn("w:numId")).get(
                            BUILD_WORD.qn("w:val")
                        )
                    )
                self.assertEqual(num_ids[0], num_ids[1])
                self.assertNotEqual(num_ids[1], num_ids[2])

                numbering = document.part.numbering_part.element
                level_texts = []
                for num_id in (num_ids[0], num_ids[2]):
                    num = next(
                        node
                        for node in numbering.findall(BUILD_WORD.qn("w:num"))
                        if node.get(BUILD_WORD.qn("w:numId")) == num_id
                    )
                    abstract_id = num.find(
                        BUILD_WORD.qn("w:abstractNumId")
                    ).get(BUILD_WORD.qn("w:val"))
                    abstract = next(
                        node
                        for node in numbering.findall(
                            BUILD_WORD.qn("w:abstractNum")
                        )
                        if node.get(BUILD_WORD.qn("w:abstractNumId"))
                        == abstract_id
                    )
                    level_texts.append(
                        abstract.find(
                            BUILD_WORD.qn("w:lvl")
                            + "/"
                            + BUILD_WORD.qn("w:lvlText")
                        ).get(BUILD_WORD.qn("w:val"))
                    )
                self.assertEqual(tuple(level_texts), expected_labels)

    def test_equation_marker_is_separated_from_following_prose(self) -> None:
        labels = BUILD_WORD.LabelRegistry()
        converted = BUILD_WORD.apply_float_numbering(
            "$$\n\\begin{equation}x=1\\label{eq:test}\\end{equation}\n$$\n"
            "式中，x为变量。",
            "humanities",
            2,
            BUILD_WORD.FloatCounters(),
            labels,
        )

        self.assertRegex(
            converted,
            r"YIBIN_INTERNAL_EQUATION_2_1\n:::\n{2,}式中，x为变量。",
        )
        self.assertEqual(labels.targets["eq:test"].number, "2.1")

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

    def test_table_header_stays_with_first_data_row(self) -> None:
        document = BUILD_WORD.Document()
        self._add_table_styles(document)
        table = document.add_table(rows=2, cols=2)
        layout = BUILD_WORD.LatexTableLayout(
            "longtable",
            tuple(BUILD_WORD.parse_latex_table_columns("ll")),
        )

        BUILD_WORD._format_tables(document, [layout])

        self.assertIsNotNone(
            table.rows[0]._tr.get_or_add_trPr().find(BUILD_WORD.qn("w:tblHeader"))
        )
        for cell in table.rows[0].cells:
            for paragraph in cell.paragraphs:
                self.assertTrue(paragraph.paragraph_format.keep_with_next)
                self.assertIsNotNone(
                    paragraph._p.get_or_add_pPr().find(BUILD_WORD.qn("w:keepNext"))
                )
        for cell in table.rows[1].cells:
            for paragraph in cell.paragraphs:
                self.assertIsNone(paragraph.paragraph_format.keep_with_next)

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

    def test_questionnaire_tables_require_semantic_headers(self) -> None:
        document = BUILD_WORD.Document()
        ahp = self._make_ahp_table(document)
        meaning = self._make_rating_meaning_table(document)
        fce = self._make_fce_table(document)

        self.assertEqual(BUILD_WORD._questionnaire_table_kind(ahp), "ahp")
        self.assertEqual(
            BUILD_WORD._questionnaire_table_kind(meaning),
            "rating_meaning",
        )
        self.assertEqual(BUILD_WORD._questionnaire_table_kind(fce), "fce")

        ahp.rows[1].cells[1].text = "0"
        self.assertIsNone(BUILD_WORD._questionnaire_table_kind(ahp))
        fce.rows[0].cells[1].text = "普通字段"
        self.assertIsNone(BUILD_WORD._questionnaire_table_kind(fce))

    def test_questionnaire_tables_use_fixed_printable_geometry(self) -> None:
        document = BUILD_WORD.Document()
        self._add_table_styles(document)
        ahp_tables = [self._make_ahp_table(document) for _ in range(6)]
        meaning_tables = [self._make_rating_meaning_table(document) for _ in range(2)]
        fce_table = self._make_fce_table(document)
        layouts = [
            BUILD_WORD.LatexTableLayout(
                "tabular",
                tuple(BUILD_WORD.parse_latex_table_columns("l" + "c" * 17 + "l")),
            ),
            BUILD_WORD.LatexTableLayout(
                "tabular",
                tuple(BUILD_WORD.parse_latex_table_columns("ccl")),
            ),
            BUILD_WORD.LatexTableLayout(
                "longtable",
                tuple(BUILD_WORD.parse_latex_table_columns("clccccc")),
            ),
        ]

        BUILD_WORD._format_tables(document, layouts)

        total_width = int(BUILD_WORD.Cm(15.5).twips)
        outer_width = int(BUILD_WORD.Cm(3.0).twips)
        expected_ahp = [
            outer_width,
            *BUILD_WORD._distribute_twips(total_width - 2 * outer_width, 17),
            outer_width,
        ]
        for table in ahp_tables:
            widths = self._grid_widths(table)
            self.assertEqual(widths, expected_ahp)
            self.assertEqual(sum(widths), total_width)
            self.assertEqual(widths[0], outer_width)
            self.assertEqual(widths[-1], outer_width)
            self.assertLessEqual(max(widths[1:18]) - min(widths[1:18]), 1)
            header_widths = [
                int(
                    cell.find(BUILD_WORD.qn("w:tcPr"))
                    .find(BUILD_WORD.qn("w:tcW"))
                    .get(BUILD_WORD.qn("w:w"))
                )
                for cell in table.rows[0]._tr.findall(BUILD_WORD.qn("w:tc"))
            ]
            self.assertEqual(
                header_widths,
                [outer_width, sum(widths[1:18]), outer_width],
            )
            height = table.rows[2]._tr.get_or_add_trPr().find(
                BUILD_WORD.qn("w:trHeight")
            )
            self.assertIsNotNone(height)
            self.assertEqual(height.get(BUILD_WORD.qn("w:val")), "560")
            self.assertEqual(height.get(BUILD_WORD.qn("w:hRule")), "atLeast")
            scale_margins = table.rows[2].cells[1]._tc.get_or_add_tcPr().find(
                BUILD_WORD.qn("w:tcMar")
            )
            self.assertEqual(
                scale_margins.find(BUILD_WORD.qn("w:left")).get(
                    BUILD_WORD.qn("w:w")
                ),
                "20",
            )
            self.assertEqual(
                scale_margins.find(BUILD_WORD.qn("w:right")).get(
                    BUILD_WORD.qn("w:w")
                ),
                "20",
            )
            self.assertEqual(
                table.rows[2].cells[0].paragraphs[0].style.name,
                BUILD_WORD.STYLE_TABLE_TEXT,
            )
            self.assertEqual(
                table.rows[2].cells[1].paragraphs[0].style.name,
                BUILD_WORD.STYLE_TABLE_CENTER,
            )
            self.assertIn("\n", table.rows[2].cells[0].text)
            self.assertEqual(
                table.rows[2].cells[0].text.replace("\n", ""),
                "C42废弃物处理与循环利用",
            )

        expected_meaning = [
            int(BUILD_WORD.Cm(2.0).twips),
            int(BUILD_WORD.Cm(2.0).twips),
            total_width - 2 * int(BUILD_WORD.Cm(2.0).twips),
        ]
        for table in meaning_tables:
            self.assertEqual(self._grid_widths(table), expected_meaning)
            self.assertEqual(
                table.rows[1].cells[2].paragraphs[0].style.name,
                BUILD_WORD.STYLE_TABLE_TEXT,
            )

        code_width = int(BUILD_WORD.Cm(1.3).twips)
        indicator_width = int(BUILD_WORD.Cm(7.2).twips)
        expected_fce = [
            code_width,
            indicator_width,
            *BUILD_WORD._distribute_twips(
                total_width - code_width - indicator_width,
                5,
            ),
        ]
        self.assertEqual(self._grid_widths(fce_table), expected_fce)
        self.assertEqual(
            fce_table.rows[1].cells[1].paragraphs[0].style.name,
            BUILD_WORD.STYLE_TABLE_TEXT,
        )
        for cell in fce_table.rows[1].cells[2:]:
            self.assertEqual(
                cell.paragraphs[0].style.name,
                BUILD_WORD.STYLE_TABLE_CENTER,
            )

    def test_pandoc_math_checkboxes_become_native_printable_squares(self) -> None:
        document = BUILD_WORD.Document()
        self._add_table_styles(document)
        table = document.add_table(rows=1, cols=3)

        def math_node(value: str):
            math = BUILD_WORD.OxmlElement("m:oMath")
            run = BUILD_WORD.OxmlElement("m:r")
            text = BUILD_WORD.OxmlElement("m:t")
            text.text = value
            run.append(text)
            math.append(run)
            return math

        table.rows[0].cells[0].paragraphs[0]._p.append(math_node("\u25ab"))
        wrapper = BUILD_WORD.OxmlElement("m:oMathPara")
        wrapper.append(math_node("\u25ab"))
        table.rows[0].cells[1].paragraphs[0]._p.append(wrapper)
        table.rows[0].cells[2].paragraphs[0]._p.append(math_node("x+1"))

        BUILD_WORD._format_tables(
            document,
            [
                BUILD_WORD.LatexTableLayout(
                    "tabular",
                    tuple(BUILD_WORD.parse_latex_table_columns("ccc")),
                )
            ],
        )
        replaced = BUILD_WORD._replace_pandoc_checkbox_math(document)
        self.assertEqual(replaced, 2)

        remaining = [
            "".join(node.text or "" for node in math.iter(BUILD_WORD.qn("m:t")))
            for math in document.element.iter(BUILD_WORD.qn("m:oMath"))
        ]
        self.assertEqual(remaining, ["x+1"])
        for math_wrapper in document.element.iter(BUILD_WORD.qn("m:oMathPara")):
            self.assertEqual(math_wrapper.findall(BUILD_WORD.qn("w:r")), [])

        square_runs = [
            run
            for run in document.element.iter(BUILD_WORD.qn("w:r"))
            if "".join(
                node.text or "" for node in run.iter(BUILD_WORD.qn("w:t"))
            )
            == "\u25a1"
        ]
        self.assertEqual(len(square_runs), 2)
        for run in square_runs:
            properties = run.find(BUILD_WORD.qn("w:rPr"))
            fonts = properties.find(BUILD_WORD.qn("w:rFonts"))
            for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
                self.assertEqual(
                    fonts.get(BUILD_WORD.qn(f"w:{attribute}")),
                    "SimSun",
                )
            self.assertEqual(
                properties.find(BUILD_WORD.qn("w:sz")).get(
                    BUILD_WORD.qn("w:val")
                ),
                "24",
            )
            self.assertEqual(
                properties.find(BUILD_WORD.qn("w:szCs")).get(
                    BUILD_WORD.qn("w:val")
                ),
                "24",
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


class OptionalCoverAndDirectoryTests(unittest.TestCase):
    def test_main_parser_recognizes_figure_and_table_directories(self) -> None:
        events = BUILD_WORD.parse_main_events(
            r"""\frontmatter
\tableofcontents
\listoffigures
\listoftables
\mainmatter"""
        )
        self.assertEqual(
            [event.kind for event in events],
            [
                "frontmatter",
                "tableofcontents",
                "listoffigures",
                "listoftables",
                "mainmatter",
            ],
        )

    def test_caption_directory_uses_native_word_caption_field(self) -> None:
        document = BUILD_WORD.Document()
        title = document.add_paragraph("图目录")
        BUILD_WORD._insert_caption_list_after(title, "图")
        inserted = title._p.getnext()
        field = inserted.find(BUILD_WORD.qn("w:fldSimple"))
        self.assertIsNotNone(field)
        self.assertEqual(
            field.get(BUILD_WORD.qn("w:instr")),
            'TOC \\h \\z \\c "图"',
        )
        self.assertEqual(field.get(BUILD_WORD.qn("w:dirty")), "true")

    def test_toc_uses_native_word_field(self) -> None:
        document = BUILD_WORD.Document()
        title = document.add_paragraph("目录")
        BUILD_WORD._insert_toc_after(title)
        inserted = title._p.getnext()
        field = inserted.find(BUILD_WORD.qn("w:fldSimple"))
        self.assertIsNotNone(field)
        self.assertEqual(
            field.get(BUILD_WORD.qn("w:instr")),
            'TOC \\o "1-3" \\h \\z \\u',
        )
        self.assertEqual(field.get(BUILD_WORD.qn("w:dirty")), "true")

    def test_word_refresh_reupdates_directories_after_layout_changes(self) -> None:
        source = (ROOT / "tools" / "refresh_word.ps1").read_text(encoding="utf-8")
        layout_start = source.index("Stabilize-TableRowsForContinuation -Document $document")
        save = source.index("$document.Save()", layout_start)
        final_refresh = source.index(
            "Update-WordDocumentFields -Document $document",
            layout_start,
        )
        self.assertLess(final_refresh, save)
        self.assertEqual(
            source.count("Update-WordDocumentFields -Document $document"),
            2,
        )
        ooxml_patch = source.index("Set-OfficialTocOoxml -Path $documentPath", save)
        self.assertIn(
            "$parts = @('word/styles.xml', 'word/document.xml', 'word/settings.xml')",
            source,
        )
        self.assertIn("$partName -eq 'word/settings.xml'", source[:ooxml_patch])
        self.assertIn("-LocalName 'updateFields'", source[:ooxml_patch])
        self.assertIn("-Name 'val' -Value 'true'", source[:ooxml_patch])

    def test_word_refresh_formats_caption_directories_like_toc_level_one(self) -> None:
        source = (ROOT / "tools" / "refresh_word.ps1").read_text(encoding="utf-8")
        self.assertIn("$captionDirectoryStyle = $document.Styles.Item(-36)", source)
        self.assertIn("$captionDirectoryStyle.Font.NameFarEast = '宋体'", source)
        self.assertIn("$captionDirectoryStyle.Font.Size = 12", source)
        self.assertIn("$captionDirectoryStyle.ParagraphFormat.LineSpacing = 18", source)
        self.assertIn("(?:table of figures|图表目录)", source)
        self.assertIn("$captionDirectoryParagraphs", source)

    def test_latex_cover_keeps_version_and_date_optional(self) -> None:
        source = (ROOT / "yibinthesis.cls").read_text(encoding="utf-8")
        self.assertIn(r"version                 = {}", source)
        self.assertIn(r"date                    = {}", source)
        self.assertIn(r"\yibinifversion", source)
        self.assertIn(r"\yibinifdate", source)
        cover = source[source.index(r"\NewDocumentCommand{\makeyibincover}") :]
        self.assertLess(
            cover.index(r"\centering\yibin@covertitle"),
            cover.index(r"\centering\yibinhei\zihao{4}\yibinversion"),
        )
        self.assertIn(r"(0cm,12.55cm)", cover)
        self.assertIn(r"\renewcommand{\listfigurename}{图目录}", source)
        self.assertIn(r"\renewcommand{\listtablename}{表目录}", source)

    def test_word_cover_spacing_is_owned_by_conditional_title_styles(self) -> None:
        source = (ROOT / "lib" / "word_core.py").read_text(encoding="utf-8")
        self.assertIn(
            'title_style = "CoverTitleWithVersion" if version_value else "CoverTitle"',
            source,
        )
        self.assertNotIn("title_paragraph.paragraph_format.space_after", source)

        reference_source = (ROOT / "lib" / "build_reference_docx.py").read_text(
            encoding="utf-8"
        )
        self.assertIn('("CoverTitleWithVersion", SIMHEI, 18, True', reference_source)

    def test_word_cover_audit_accepts_and_validates_optional_date(self) -> None:
        source = (ROOT / "tools" / "refresh_word.ps1").read_text(encoding="utf-8")
        self.assertNotIn("must not contain a CoverDate paragraph", source)
        self.assertIn("CoverDate paragraph count expected at most 1", source)
        self.assertIn("Assert-Font 'Cover date font'", source)
        self.assertIn("Assert-Near 'Cover date size'", source)

    def test_word_cover_value_audit_is_scoped_to_cover_value_lines(self) -> None:
        source = (ROOT / "tools" / "refresh_word.ps1").read_text(encoding="utf-8")
        self.assertIn("'^CoverValueLine$' { $coverValueLines += $paragraph", source)
        self.assertIn("foreach ($paragraph in $coverValueLines)", source)

    def test_word_declaration_audit_searches_signature_slots_by_text(self) -> None:
        source = (ROOT / "tools" / "refresh_word.ps1").read_text(encoding="utf-8")
        self.assertIn("$officialAuthorSignature = $declarationSignatures | Where-Object", source)
        self.assertNotIn("$declarationSignatures[0]).StartsWith", source)

    def test_word_directory_audit_accepts_main_figure_and_table_headings(self) -> None:
        source = (ROOT / "tools" / "refresh_word.ps1").read_text(encoding="utf-8")
        self.assertIn("@('目录', '图目录', '表目录')", source)
        self.assertNotIn("TOC heading paragraph count expected 1", source)

    def test_word_body_audit_ignores_only_empty_structural_paragraphs(self) -> None:
        source = (ROOT / "tools" / "refresh_word.ps1").read_text(encoding="utf-8")
        self.assertIn("[string]::IsNullOrWhiteSpace($paragraphText)", source)
        self.assertIn("'Body Text', '宜宾论文-正文'", source)

    def test_word_body_audit_handles_mixed_citation_font_sizes(self) -> None:
        source = (ROOT / "tools" / "refresh_word.ps1").read_text(encoding="utf-8")
        self.assertIn("function Get-ParagraphFontSize", source)
        self.assertIn("[math]::Abs($size) -gt 1000", source)
        self.assertIn("$Paragraph.Range.ParagraphStyle.Font.Size", source)


class DocumentModulePackagingTests(unittest.TestCase):
    def test_document_specific_layouts_stay_out_of_core_class(self) -> None:
        source = (ROOT / "yibinthesis.cls").read_text(encoding="utf-8")
        self.assertNotIn(r"\newtcolorbox{yibin@proposalbox}", source)
        self.assertNotIn(r"\NewDocumentCommand{\makeyibinproposal}", source)
        self.assertNotIn(
            r"\NewDocumentCommand{\makeyibinliteraturereviewcover}",
            source,
        )
        self.assertIn(r"\RequirePackage{yibinthesis-proposal}", source)
        self.assertIn(r"\RequirePackage{yibinthesis-literature-review}", source)

    def test_external_runtime_packages_every_document_module(self) -> None:
        source = (ROOT / "build.ps1").read_text(encoding="utf-8")
        for name in (
            "yibinthesis.cls",
            "yibinthesis-proposal.sty",
            "yibinthesis-literature-review.sty",
        ):
            with self.subTest(name=name):
                self.assertIn(f"'{name}'", source)


if __name__ == "__main__":
    unittest.main()
