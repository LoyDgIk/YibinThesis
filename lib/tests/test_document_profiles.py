from __future__ import annotations

import importlib.util
import re
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[2]


def _load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


BUILD_WORD = _load_module(
    "yibinthesis_build_word_profiles",
    ROOT / "lib" / "build_word.py",
)
BUILD_REFERENCE = _load_module(
    "yibinthesis_build_reference_profiles",
    ROOT / "lib" / "build_reference_docx.py",
)


class DocumentProfileParsingTests(unittest.TestCase):
    def test_document_type_defaults_and_explicit_options(self) -> None:
        cases = (
            (r"\documentclass{yibinthesis}", "thesis", "humanities", "2024"),
            (
                r"\documentclass[humanities,proposal]{yibinthesis}",
                "proposal",
                "humanities",
                "2022",
            ),
            (
                r"\documentclass[literature-review,science]{yibinthesis}",
                "literature-review",
                "science",
                "2024",
            ),
        )
        for source, document_type, discipline, template_year in cases:
            with self.subTest(source=source):
                profile = BUILD_WORD.resolve_profile(
                    source,
                    {"template-year": template_year},
                )
                self.assertEqual(profile.document_type, document_type)
                self.assertEqual(profile.discipline, discipline)
                self.assertEqual(profile.template_year, template_year)

    def test_document_type_options_are_mutually_exclusive(self) -> None:
        for options in (
            "thesis,proposal",
            "proposal,literature-review",
            "thesis,proposal,literature-review",
        ):
            with self.subTest(options=options):
                with self.assertRaisesRegex(BUILD_WORD.BuildError, "文档类型类选项互斥"):
                    BUILD_WORD.parse_document_type(
                        rf"\documentclass[{options}]{{yibinthesis}}"
                    )

    def test_template_year_is_bound_to_document_type(self) -> None:
        cases = (
            (r"\documentclass{yibinthesis}", "2024"),
            (r"\documentclass[proposal]{yibinthesis}", "2022"),
            (r"\documentclass[literature-review]{yibinthesis}", "2024"),
        )
        for main, expected in cases:
            with self.subTest(main=main):
                self.assertEqual(BUILD_WORD.resolve_profile(main, {}).template_year, expected)
                self.assertEqual(
                    BUILD_WORD.resolve_profile(main, {"template-year": "  "}).template_year,
                    expected,
                )
                self.assertEqual(
                    BUILD_WORD.resolve_profile(main, {"template-year": expected}).template_year,
                    expected,
                )

    def test_mismatched_template_year_is_rejected(self) -> None:
        cases = (
            (r"\documentclass{yibinthesis}", "2022", "thesis", "2024"),
            (r"\documentclass[proposal]{yibinthesis}", "2024", "proposal", "2022"),
            (
                r"\documentclass[literature-review]{yibinthesis}",
                "2022",
                "literature-review",
                "2024",
            ),
        )
        for main, actual, document_type, expected in cases:
            with self.subTest(main=main, actual=actual):
                with self.assertRaisesRegex(
                    BUILD_WORD.BuildError,
                    rf"文档类型 {re.escape(document_type)} 仅支持 template-year={expected}",
                ):
                    BUILD_WORD.resolve_profile(main, {"template-year": actual})


class ProposalFieldContractTests(unittest.TestCase):
    @staticmethod
    def _source(keys: tuple[str, ...]) -> str:
        return "\n".join(
            rf"\yibinproposalfield{{{key}}}{{{key} 的内容 {{含嵌套参数}}}}"
            for key in keys
        )

    def test_seven_fields_are_extracted_in_the_official_order(self) -> None:
        fields = BUILD_WORD.extract_proposal_fields(
            self._source(BUILD_WORD.PROPOSAL_FIELD_ORDER)
        )
        self.assertEqual(
            tuple(key for key, _ in fields),
            BUILD_WORD.PROPOSAL_FIELD_ORDER,
        )
        self.assertEqual(len(fields), 7)
        self.assertIn("{含嵌套参数}", fields[0][1])

    def test_missing_field_is_rejected(self) -> None:
        keys = BUILD_WORD.PROPOSAL_FIELD_ORDER[:-1]
        with self.assertRaisesRegex(
            BUILD_WORD.BuildError,
            "开题报告字段缺失：advisor-opinion",
        ):
            BUILD_WORD.extract_proposal_fields(self._source(keys))

    def test_duplicate_field_is_rejected(self) -> None:
        keys = BUILD_WORD.PROPOSAL_FIELD_ORDER + ("schedule",)
        with self.assertRaisesRegex(
            BUILD_WORD.BuildError,
            "开题报告字段重复：schedule",
        ):
            BUILD_WORD.extract_proposal_fields(self._source(keys))

    def test_out_of_order_fields_are_rejected(self) -> None:
        keys = list(BUILD_WORD.PROPOSAL_FIELD_ORDER)
        keys[1], keys[2] = keys[2], keys[1]
        with self.assertRaisesRegex(
            BUILD_WORD.BuildError,
            "开题报告字段顺序必须为",
        ):
            BUILD_WORD.extract_proposal_fields(self._source(tuple(keys)))

    def test_unnumbered_heading_uses_reusable_proposal_style(self) -> None:
        converted = BUILD_WORD.demote_proposal_headings(
            "正文\n# 编号标题\n内容\n# 可行性分析 {.unnumbered}\n后文"
        )
        self.assertIn("## 编号标题", converted)
        self.assertIn(
            f'custom-style="{BUILD_WORD.STYLE_PROPOSAL_UNNUMBERED_HEADING}"',
            converted,
        )
        self.assertNotIn("## 可行性分析", converted)

    def test_figure_numbering_preserves_following_custom_style_block(self) -> None:
        source = (
            "![研究路线](route.png){#fig:route}\n\n"
            f'::: {{custom-style="{BUILD_WORD.STYLE_PROPOSAL_UNNUMBERED_HEADING}"}}\n'
            "可行性分析\n:::"
        )
        converted = BUILD_WORD.apply_float_numbering(
            source,
            "humanities",
            0,
            BUILD_WORD.FloatCounters(),
            BUILD_WORD.LabelRegistry(),
        )
        self.assertIn(
            f'}}\n\n::: {{custom-style="{BUILD_WORD.STYLE_PROPOSAL_UNNUMBERED_HEADING}"}}',
            converted,
        )


class WordProfileStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._temporary = tempfile.TemporaryDirectory()
        cls.root = Path(cls._temporary.name)
        cls.reference_docx = cls.root / "reference.docx"
        BUILD_REFERENCE.build_reference_docx(cls.reference_docx)

    @classmethod
    def tearDownClass(cls) -> None:
        cls._temporary.cleanup()

    @staticmethod
    def _metadata() -> dict[str, str]:
        return {
            "template-year": "2024",
            "title": "公共服务数据治理方法研究",
            "author": "张同学",
            "student-id": "202400000",
            "college": "示例学部",
            "major": "示例专业",
            "grade": "2024",
            "class-name": "1班",
            "advisor": "李老师",
            "advisor-title": "讲师",
            "date": "2026年7月",
        }

    @staticmethod
    def _field_instructions(element) -> list[str]:
        instructions = [
            node.text or ""
            for node in element.iter(BUILD_WORD.qn("w:instrText"))
        ]
        instructions.extend(
            node.get(BUILD_WORD.qn("w:instr"), "")
            for node in element.iter(BUILD_WORD.qn("w:fldSimple"))
        )
        return instructions

    @staticmethod
    def _page_number_attributes(section) -> dict[str, str] | None:
        node = section._sectPr.find(BUILD_WORD.qn("w:pgNumType"))
        if node is None:
            return None
        return {
            "format": node.get(BUILD_WORD.qn("w:fmt"), ""),
            "start": node.get(BUILD_WORD.qn("w:start"), ""),
        }

    def test_reference_document_defines_reusable_profile_styles(self) -> None:
        document = Document(self.reference_docx)
        expected = {
            BUILD_WORD.STYLE_PROPOSAL_TITLE,
            BUILD_WORD.STYLE_PROPOSAL_SUBTITLE,
            BUILD_WORD.STYLE_PROPOSAL_LABEL,
            BUILD_WORD.STYLE_PROPOSAL_BODY,
            BUILD_WORD.STYLE_PROPOSAL_PROMPT,
            BUILD_WORD.STYLE_PROPOSAL_UNNUMBERED_HEADING,
            BUILD_WORD.STYLE_PROPOSAL_SIGNATURE,
            BUILD_WORD.STYLE_REVIEW_DOCUMENT_TITLE,
            BUILD_WORD.STYLE_REVIEW_THESIS_TITLE,
            BUILD_WORD.STYLE_REVIEW_INFO_LABEL,
            BUILD_WORD.STYLE_REVIEW_INFO_VALUE,
            BUILD_WORD.STYLE_REVIEW_DATE,
            BUILD_WORD.STYLE_FIGURE_CAPTION,
            BUILD_WORD.STYLE_TABLE_CAPTION,
            BUILD_WORD.STYLE_CITATION,
            BUILD_WORD.STYLE_LIST_BODY,
            BUILD_WORD.STYLE_APPENDIX_SECTION,
        }
        self.assertTrue(expected.issubset({style.name for style in document.styles}))

    def test_reference_caption_directory_matches_first_level_toc(self) -> None:
        document = Document(self.reference_docx)
        style = document.styles[BUILD_REFERENCE.STYLE_TABLE_OF_FIGURES]
        paragraph = style.paragraph_format
        properties = style.element.get_or_add_pPr()
        indent = properties.find(BUILD_REFERENCE.qn("w:ind"))
        tabs = properties.find(BUILD_REFERENCE.qn("w:tabs"))
        tab = tabs.find(BUILD_REFERENCE.qn("w:tab"))
        fonts = style.element.get_or_add_rPr().find(BUILD_REFERENCE.qn("w:rFonts"))

        self.assertEqual(style.font.name, "SimSun")
        self.assertEqual(style.font.size.pt, 12)
        self.assertFalse(style.font.bold)
        self.assertIsNone(style.element.get(BUILD_REFERENCE.qn("w:customStyle")))
        self.assertEqual(
            fonts.get(BUILD_REFERENCE.qn("w:eastAsia")),
            "SimSun",
        )
        self.assertEqual(paragraph.alignment, BUILD_REFERENCE.WD_ALIGN_PARAGRAPH.JUSTIFY)
        self.assertEqual(paragraph.line_spacing, 1.5)
        self.assertEqual(paragraph.space_before.pt, 0)
        self.assertEqual(paragraph.space_after.pt, 0)
        self.assertEqual(indent.get(BUILD_REFERENCE.qn("w:left")), "0")
        self.assertEqual(indent.get(BUILD_REFERENCE.qn("w:leftChars")), "0")
        self.assertEqual(indent.get(BUILD_REFERENCE.qn("w:firstLine")), "0")
        self.assertEqual(tab.get(BUILD_REFERENCE.qn("w:val")), "right")
        self.assertEqual(tab.get(BUILD_REFERENCE.qn("w:leader")), "dot")
        self.assertEqual(tab.get(BUILD_REFERENCE.qn("w:pos")), "8777")

    def test_existing_pandoc_caption_is_not_duplicated(self) -> None:
        document = Document(self.reference_docx)
        drawing = document.paragraphs[0]
        drawing.text = ""
        drawing.add_run().add_picture(
            str(ROOT / "examples" / "full-featured" / "assets" / "document-pipeline.png")
        )
        for properties in drawing._p.iter(BUILD_WORD.qn("wp:docPr")):
            properties.set("descr", "图1 测试图题")
        document.add_paragraph("图1 测试图题", style=BUILD_WORD.STYLE_FIGURE_CAPTION)

        BUILD_WORD._split_pandoc_captioned_figures(document)

        captions = [paragraph for paragraph in document.paragraphs if paragraph.text == "图1 测试图题"]
        self.assertEqual(len(captions), 1)
        self.assertEqual(captions[0].style.name, BUILD_WORD.STYLE_FIGURE_CAPTION)

    def test_custom_thesis_headings_do_not_replace_builtin_headings(self) -> None:
        document = Document(self.reference_docx)
        custom_names = (
            BUILD_WORD.STYLE_HEADING_1,
            BUILD_WORD.STYLE_HEADING_2,
            BUILD_WORD.STYLE_HEADING_3,
            BUILD_WORD.STYLE_HEADING_4,
        )
        for level, custom_name in enumerate(custom_names, start=1):
            with self.subTest(level=level):
                builtin = document.styles[f"Heading {level}"]
                custom = document.styles[custom_name]
                self.assertEqual(builtin.style_id, f"Heading{level}")
                self.assertEqual(builtin.name, f"Heading {level}")
                self.assertNotEqual(custom.style_id, builtin.style_id)
                self.assertEqual(custom.name, custom_name)

        self.assertTrue(
            document.styles[BUILD_WORD.STYLE_HEADING_1]
            .paragraph_format.page_break_before
        )

    def test_common_cover_optional_styles_follow_school_type_scale(self) -> None:
        document = Document(self.reference_docx)
        version = document.styles["CoverVersion"].font
        date = document.styles["CoverDate"].font

        self.assertEqual(version.size.pt, 14)
        self.assertFalse(version.bold)
        self.assertEqual(
            document.styles["CoverVersion"].paragraph_format.space_after.pt,
            23,
        )
        self.assertEqual(date.size.pt, 15)
        self.assertTrue(date.bold)

    def test_proposal_has_cover_and_form_sections_starting_at_page_four(self) -> None:
        source = Document(self.reference_docx)
        source.paragraphs[0].text = ""
        for key in BUILD_WORD.PROPOSAL_FIELD_ORDER:
            source.add_paragraph(BUILD_WORD.PROPOSAL_FIELD_MARKER_PREFIX + key.upper())
            source.add_paragraph(f"{key} 内容")

        input_docx = self.root / "proposal-input.docx"
        output_docx = self.root / "proposal-output.docx"
        source.save(input_docx)
        BUILD_WORD._postprocess_proposal_docx(
            input_docx,
            output_docx,
            metadata={**self._metadata(), "template-year": "2022"},
            main_dir=self.root,
            metadata_dir=self.root,
            discipline="humanities",
            citation_mode="linked",
            labels=BUILD_WORD.LabelRegistry(),
            citations=BUILD_WORD.CitationRegistry(),
            bibliography_items=[],
            table_layouts=[],
            word_figure_sequence="图",
            word_table_sequence="表",
        )

        document = Document(output_docx)
        self.assertIn("document-type=proposal", document.core_properties.keywords)
        self.assertIn("template-year=2022", document.core_properties.keywords)
        self.assertEqual(len(document.sections), 1)
        self.assertEqual(
            self._page_number_attributes(document.sections[0]),
            {"format": "decimal", "start": "4"},
        )
        self.assertEqual([len(table.rows) for table in document.tables], [7])
        self.assertEqual(
            [
                "".join(row.cells[0].paragraphs[0].text.split())
                for table in document.tables
                for row in table.rows
            ],
            [
                "学院（部）",
                "专    业",
                "学生姓名",
                "学    号",
                "年    级",
                "指导教师",
            ],
        )
        self.assertTrue(
            all(
                row.cells[0].paragraphs[0].style.name
                == BUILD_WORD.STYLE_PROPOSAL_LABEL
                for table in document.tables[:1]
                for row in table.rows
            )
        )
        page_breaks = [
            node
            for node in document.element.body.iter(BUILD_WORD.qn("w:br"))
            if node.get(BUILD_WORD.qn("w:type")) == "page"
        ]
        self.assertEqual(len(page_breaks), 0)
        self.assertIn(
            " PAGE ",
            self._field_instructions(document.sections[0].footer._element),
        )

    def test_literature_review_has_two_sections_and_native_word_fields(self) -> None:
        source = Document(self.reference_docx)
        source.paragraphs[0].text = BUILD_WORD.SECTION_REVIEW_COVER_END
        source.add_paragraph("研究背景", style=BUILD_WORD.STYLE_HEADING_1)
        source.add_paragraph("图1 技术路线", style="Caption")

        labels = BUILD_WORD.LabelRegistry()
        labels.register("fig:review", "1", "图")
        reference = labels.placeholder("fig:review", parenthesized=False)
        citations = BUILD_WORD.CitationRegistry()
        citation = citations.placeholder("source1")
        source.add_paragraph(f"见图{reference}，参考文献{citation}。")
        source.add_paragraph("[1] 示例文献", style="Bibliography")

        input_docx = self.root / "review-input.docx"
        output_docx = self.root / "review-output.docx"
        source.save(input_docx)
        BUILD_WORD._postprocess_review_docx(
            input_docx,
            output_docx,
            metadata=self._metadata(),
            main_dir=self.root,
            metadata_dir=self.root,
            discipline="humanities",
            citation_mode="linked",
            labels=labels,
            citations=citations,
            bibliography_items=[],
            table_layouts=[],
            word_figure_sequence="图",
            word_table_sequence="表",
        )

        document = Document(output_docx)
        self.assertEqual(len(document.sections), 2)
        self.assertIsNone(self._page_number_attributes(document.sections[0]))
        self.assertEqual(
            self._page_number_attributes(document.sections[1]),
            {"format": "decimal", "start": "1"},
        )
        self.assertEqual(
            next(p for p in document.paragraphs if p.text == "研究背景").style.name,
            BUILD_WORD.STYLE_HEADING_1,
        )

        caption = next(p for p in document.paragraphs if p.text == "图1 技术路线")
        self.assertEqual(caption.style.name, BUILD_WORD.STYLE_FIGURE_CAPTION)
        self.assertIn(
            " SEQ 图 \\r 1 \\* ARABIC ",
            self._field_instructions(caption._p),
        )
        bookmark_names = {
            node.get(BUILD_WORD.qn("w:name"))
            for node in caption._p.iter(BUILD_WORD.qn("w:bookmarkStart"))
        }
        self.assertIn(BUILD_WORD._label_bookmark_name("fig:review"), bookmark_names)
        self.assertIn(
            BUILD_WORD._label_number_bookmark_name("fig:review"),
            bookmark_names,
        )

        reference_paragraph = next(
            p for p in document.paragraphs if p.text == "见图1，参考文献[1]。"
        )
        instructions = self._field_instructions(reference_paragraph._p)
        self.assertIn(
            f" REF {BUILD_WORD._label_bookmark_name('fig:review')} \\h ",
            instructions,
        )
        citation_bookmark = BUILD_WORD._citation_bookmark_name("source1")
        self.assertIn(
            f" REF {citation_bookmark} \\h \\* CHARFORMAT ",
            instructions,
        )
        all_bookmarks = {
            node.get(BUILD_WORD.qn("w:name"))
            for node in document.element.iter(BUILD_WORD.qn("w:bookmarkStart"))
        }
        ref_targets = {
            match.group(1)
            for instruction in instructions
            if (match := re.search(r"\bREF\s+(\S+)", instruction))
        }
        self.assertTrue(ref_targets.issubset(all_bookmarks))

        citation_style = document.styles[BUILD_WORD.STYLE_CITATION]
        vertical = citation_style.element.find(
            BUILD_WORD.qn("w:rPr") + "/" + BUILD_WORD.qn("w:vertAlign")
        )
        self.assertIsNotNone(vertical)
        self.assertEqual(vertical.get(BUILD_WORD.qn("w:val")), "superscript")
        styled_runs = []
        for run in reference_paragraph._p.iter(BUILD_WORD.qn("w:r")):
            style = run.find(
                BUILD_WORD.qn("w:rPr") + "/" + BUILD_WORD.qn("w:rStyle")
            )
            if style is not None and style.get(BUILD_WORD.qn("w:val")) == citation_style.style_id:
                styled_runs.append(run)
        self.assertGreaterEqual(len(styled_runs), 7)
        self.assertTrue(
            all(
                run.find(
                    BUILD_WORD.qn("w:rPr") + "/" + BUILD_WORD.qn("w:vertAlign")
                )
                is None
                for run in styled_runs
            ),
            "引用上标应由字符样式驱动，而不是逐个 run 直接编码。",
        )


if __name__ == "__main__":
    unittest.main()
