"""Proposal and literature-review Word profile assembly."""

from __future__ import annotations

import sys


def bind_core(core):
    """Expose core helpers to profile functions without duplicating them."""
    module = sys.modules[__name__]
    profile_functions = {
        "_build_non_thesis",
        "_postprocess_proposal_docx",
        "_postprocess_review_docx",
    }
    for name in dir(core):
        if name not in {"bind_core", "sys"} and name not in profile_functions:
            setattr(module, name, getattr(core, name))
    return module


def _set_form_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "000000")


def _set_form_cell_margins(cell, *, top: int = 90, left: int = 110, bottom: int = 90, right: int = 110) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for edge, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = margins.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _move_nodes_to_cell(cell, nodes: list[object], document: Document) -> None:
    anchor = cell.paragraphs[0]
    for node in nodes:
        anchor._p.addprevious(node)
    anchor.style = document.styles[STYLE_PROPOSAL_BODY]
    anchor.paragraph_format.space_before = Pt(0)
    anchor.paragraph_format.space_after = Pt(0)
    anchor.paragraph_format.line_spacing = Pt(1)
    for paragraph in cell.paragraphs:
        if paragraph is anchor:
            continue
        if paragraph.style.name in {STYLE_BODY, STYLE_FIRST_PARAGRAPH, "Normal", "Body Text"}:
            paragraph.style = document.styles[STYLE_PROPOSAL_BODY]


def _extract_marker_ranges(document: Document) -> dict[str, list[object]]:
    body = document._element.body
    children = list(body)
    markers: list[tuple[str, int]] = []
    for index, child in enumerate(children):
        if child.tag != qn("w:p"):
            continue
        text = "".join(node.text or "" for node in child.iter(qn("w:t"))).strip()
        if text.startswith(PROPOSAL_FIELD_MARKER_PREFIX):
            markers.append((text.removeprefix(PROPOSAL_FIELD_MARKER_PREFIX).lower(), index))
    expected = list(PROPOSAL_FIELD_ORDER)
    if [key for key, _ in markers] != expected:
        raise BuildError("Word 开题报告字段标记缺失或顺序错误。")
    result: dict[str, list[object]] = {}
    section_index = children.index(body.sectPr) if body.sectPr is not None else len(children)
    for marker_index, (key, start) in enumerate(markers):
        end = markers[marker_index + 1][1] if marker_index + 1 < len(markers) else section_index
        nodes = children[start + 1 : end]
        if key == "references":
            nodes = [
                node
                for node in nodes
                if not (
                    node.tag == qn("w:p")
                    and "".join(part.text or "" for part in node.iter(qn("w:t"))).strip()
                    == "参考文献"
                )
            ]
        result[key] = nodes
    for child in list(body):
        if child is not body.sectPr:
            body.remove(child)
    return result


def _add_profile_page_field(section, document: Document) -> None:
    section.footer.is_linked_to_previous = False
    _clear_story(section.footer)
    paragraph = section.footer.paragraphs[0]
    paragraph.style = document.styles["Footer"]
    _add_page_field(paragraph)
    for run in paragraph.runs:
        _set_run_fonts(run, "SimSun", "Times New Roman", 9)


def _insert_proposal_cover(
    document: Document,
    metadata: dict[str, str],
    *,
    project_root: Path,
    main_dir: Path,
    metadata_dir: Path,
) -> None:
    """Insert the filled sample's cover block before the page-4 form."""

    try:
        logo = _resolve_cover_logo(
            metadata,
            project_root,
            main_dir,
            metadata_dir,
            allow_project_fallback=True,
        )
    except BuildError:
        # Unit fixtures and external profile callers may intentionally omit
        # the optional logo; keep the textual cover usable in that case.
        logo = None
    logo_paragraph = document.add_paragraph(style="CoverLogo")
    logo_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo is not None:
        logo_paragraph.add_run().add_picture(str(logo), width=Cm(13.15), height=Cm(3.65))
    else:
        run = logo_paragraph.add_run("宜宾学院")
        _set_run_fonts(run, "SimHei", "Times New Roman", 26, bold=True)
    title = document.add_paragraph(style=STYLE_PROPOSAL_TITLE)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("本科生毕业论文")
    paper_title = document.add_paragraph(style=STYLE_PROPOSAL_TITLE)
    paper_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paper_title.add_run(metadata.get("title", "（填写中文题目）"))
    info = document.add_table(rows=6, cols=2)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info.autofit = False
    _set_fixed_table_geometry(info, (3.2, 13.8), style_name=STYLE_PROPOSAL_FORM_TABLE)
    _set_table_borders_none(info)
    rows = (
        ("学院（部）", metadata.get("college", "")),
        ("专    业", metadata.get("major", "")),
        ("学生姓名", metadata.get("author", "")),
        ("学    号", metadata.get("student-id", "")),
        ("年    级", format_grade_class(metadata)),
        ("指导教师", metadata.get("advisor", "")),
    )
    for row, (label, value) in zip(info.rows, rows):
        for cell in row.cells:
            _set_form_cell_margins(cell, top=0, bottom=0, left=40, right=40)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        row.cells[0].paragraphs[0].style = document.styles[STYLE_PROPOSAL_LABEL]
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        row.cells[0].paragraphs[0].add_run(label)
        row.cells[1].paragraphs[0].style = document.styles[STYLE_PROPOSAL_BODY]
        row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[1].paragraphs[0].add_run(value)
    date = document.add_paragraph(style=STYLE_REVIEW_DATE)
    date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date.add_run(metadata.get("date", ""))


def _append_proposal_table(
    document: Document,
    fields: tuple[str, ...],
    field_nodes: dict[str, list[object]],
    *,
    heading_abstract_id: int,
    advisor_signature: Path | None = None,
    signature_background: str = "preserve",
) -> None:
    labels = {
        "significance": "选题意义",
        "research-status": "国内外研究现状概述",
        "research-content": "主要研究内容",
        "research-approach": "拟采用的研究思路",
        "schedule": "研究工作安排及进度",
        "references": "参考文献目录",
        "advisor-opinion": "指导教师意见",
    }
    prompts = {
        "research-status": "国内、国外研究现状应分开概述；概述时语言简练、重点突出。",
        "research-approach": "主要描述研究方法、技术路线、可行性论证等。",
        "schedule": "进度（应与二级学部（院）的进度大体一致）",
        "references": "根据《信息与文献—参考文献著录规则》（GB/T 7714-2015），规范表述。",
    }
    minimum_heights = {
        "significance": 165.15,
        "research-status": 297.15,
        "research-content": 183.15,
        "research-approach": 411.6,
        "schedule": 140.1,
        "references": 383.6,
        "advisor-opinion": 99.9,
    }
    table = document.add_table(rows=len(fields), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_fixed_table_geometry(
        table,
        (1.52, 14.23),
        style_name=STYLE_PROPOSAL_FORM_TABLE,
    )
    _set_form_table_borders(table)
    for row_index, key in enumerate(fields):
        row = table.rows[row_index]
        row.height = Pt(minimum_heights[key])
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row_properties = row._tr.get_or_add_trPr()
        cant_split = row_properties.find(qn("w:cantSplit"))
        if cant_split is not None:
            row_properties.remove(cant_split)
        label_cell, body_cell = row.cells
        label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        body_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        _set_form_cell_margins(label_cell, left=40, right=40)
        _set_form_cell_margins(body_cell)
        label_paragraph = label_cell.paragraphs[0]
        label_paragraph.style = document.styles[STYLE_PROPOSAL_LABEL]
        label_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        label_paragraph.add_run("\n".join(labels[key]))
        if key in prompts:
            prompt = body_cell.paragraphs[0]
            prompt.style = document.styles[STYLE_PROPOSAL_PROMPT]
            prompt.text = prompts[key]
            anchor = body_cell.add_paragraph(style=STYLE_PROPOSAL_BODY)
            for node in field_nodes[key]:
                anchor._p.addprevious(node)
            for paragraph in body_cell.paragraphs:
                if paragraph.style.name in {STYLE_BODY, STYLE_FIRST_PARAGRAPH, "Normal", "Body Text"}:
                    paragraph.style = document.styles[STYLE_PROPOSAL_BODY]
        else:
            _move_nodes_to_cell(body_cell, field_nodes[key], document)
        _fit_nested_tables_to_cell(body_cell)
        _restart_cell_heading_numbering(document, body_cell, heading_abstract_id)
        if key == "advisor-opinion":
            signature = body_cell.add_paragraph(style=STYLE_PROPOSAL_SIGNATURE)
            signature.add_run("签名： ")
            if advisor_signature is None:
                signature.add_run("        ")
            else:
                picture_source: str | io.BytesIO = str(advisor_signature)
                if signature_background == "whiten":
                    picture_source = whiten_signature_image(advisor_signature)
                picture_run = signature.add_run()
                picture_run.add_picture(picture_source, width=Cm(2.6))
                for doc_property in picture_run._r.iter(qn("wp:docPr")):
                    doc_property.attrib.pop("descr", None)
            signature.add_run("    年    月    日")


def _postprocess_proposal_docx(
    input_docx: Path,
    output_docx: Path,
    *,
    metadata: dict[str, str],
    project_root: Path,
    main_dir: Path,
    metadata_dir: Path,
    discipline: str,
    citation_mode: str,
    labels: LabelRegistry,
    citations: CitationRegistry,
    bibliography_items: list[dict[str, object]],
    table_layouts: list[LatexTableLayout],
    word_figure_sequence: str,
    word_table_sequence: str,
) -> None:
    document = Document(input_docx)
    for paragraph in list(document.paragraphs):
        if paragraph.text.strip().startswith("YIBIN_INTERNAL_CITATION_SEED"):
            paragraph._p.getparent().remove(paragraph._p)
    heading_abstract_id = _apply_common_body_formatting(
        document,
        discipline="humanities",
        labels=labels,
        citations=citations,
        citation_mode=citation_mode,
        table_layouts=table_layouts,
        word_figure_sequence=word_figure_sequence,
        word_table_sequence=word_table_sequence,
    )
    field_nodes = _extract_marker_ranges(document)
    signature_background = metadata.get("signature-background", "preserve").strip().casefold()
    if signature_background not in {"preserve", "whiten"}:
        raise BuildError(
            "元数据 signature-background 只能是 preserve 或 whiten："
            + signature_background
        )
    advisor_signature = _resolve_signature_asset(
        metadata,
        "advisor-signature",
        main_dir,
        metadata_dir,
    )

    _insert_proposal_cover(
        document,
        metadata,
        project_root=project_root,
        main_dir=main_dir,
        metadata_dir=metadata_dir,
    )
    # The supplied sample has a separate cover section and starts the form on
    # the page numbered 4. Keep the form's existing marker content after the
    # break so long fields can still flow naturally.
    document.add_section(WD_SECTION_START.NEW_PAGE)

    title = document.add_paragraph(style=STYLE_PROPOSAL_TITLE)
    title.text = "宜宾学院本科毕业论文（设计）开题报告"
    subtitle = document.add_paragraph(style=STYLE_PROPOSAL_SUBTITLE)
    subtitle.text = "（学生填写）"
    _append_proposal_table(
        document,
        PROPOSAL_FIELD_ORDER,
        field_nodes,
        heading_abstract_id=heading_abstract_id,
        advisor_signature=advisor_signature,
        signature_background=signature_background,
    )

    if len(document.sections) != 2:
        raise BuildError(f"开题报告应生成 2 个 Word 分节，实际为 {len(document.sections)}。")
    cover, section = document.sections
    cover.start_type = WD_SECTION_START.NEW_PAGE
    cover.page_width = Cm(21)
    cover.page_height = Cm(29.7)
    cover.top_margin = Cm(2.5)
    cover.bottom_margin = Cm(2.5)
    cover.left_margin = Cm(3.0)
    cover.right_margin = Cm(2.5)
    cover.header_distance = Cm(1.5)
    cover.footer_distance = Cm(1.5)
    cover.header.is_linked_to_previous = False
    cover.footer.is_linked_to_previous = False
    _clear_story(cover.header)
    _clear_story(cover.footer)
    _set_page_number_format(cover, None)
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Pt(72)
    section.bottom_margin = Pt(72)
    section.left_margin = Pt(85.05)
    section.right_margin = Pt(62.35)
    section.header_distance = Pt(0)
    section.footer_distance = Pt(49.6)
    section.header.is_linked_to_previous = False
    _clear_story(section.header)
    _set_page_number_format(section, "decimal", 4)
    _add_profile_page_field(section, document)

    document.core_properties.title = metadata.get("title", "")
    document.core_properties.author = metadata.get("author", "")
    document.core_properties.subject = "宜宾学院本科毕业论文（设计）开题报告"
    document.core_properties.keywords = "YibinThesis;document-type=proposal;template-year=2022"
    settings = document.settings.element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_docx)
    if citation_mode == "native":
        _inject_bibliography_sources(output_docx, citations, bibliography_items)
    _normalize_embedded_pngs(output_docx)
    Document(output_docx)


def _set_single_cell_paragraph(cell, document: Document, style_name: str, value: str):
    """Replace merge residue with one style-driven paragraph."""

    paragraph = cell.paragraphs[0]
    for extra in list(cell.paragraphs[1:]):
        cell._tc.remove(extra._p)
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)
    paragraph.style = document.styles[style_name]
    if value:
        paragraph.add_run(value)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return paragraph


def _insert_review_cover(
    document: Document,
    anchor,
    metadata: dict[str, str],
    *,
    project_root: Path,
    main_dir: Path,
    metadata_dir: Path,
    discipline: str,
) -> None:
    try:
        logo = _resolve_cover_logo(
            metadata,
            project_root,
            main_dir,
            metadata_dir,
            allow_project_fallback=True,
        )
    except BuildError:
        logo = None
    logo_paragraph = _insert_paragraph_before(document, anchor, "CoverLogo")
    logo_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo is not None:
        logo_paragraph.add_run().add_picture(str(logo), width=Cm(13.15), height=Cm(3.65))
    heading = _insert_paragraph_before(document, anchor, STYLE_REVIEW_DOCUMENT_TITLE)
    heading.text = (
        "本科生毕业论文（设计）文献综述"
        if discipline == "science"
        else "本科生毕业论文文献综述"
    )
    thesis_title = _insert_paragraph_before(document, anchor, STYLE_REVIEW_THESIS_TITLE)
    thesis_title.text = metadata.get("title", "")

    table = document.add_table(rows=5, cols=4)
    anchor._p.addprevious(table._tbl)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_fixed_table_geometry(
        table,
        (3.2, 5.7, 1.6, 5.0),
        style_name=STYLE_REVIEW_INFO_TABLE,
    )
    _set_table_borders_none(table)
    rows = (
        ("学院（部）", metadata.get("college", ""), "", ""),
        ("专    业", metadata.get("major", ""), "", ""),
        ("学生姓名", metadata.get("author", ""), "", ""),
        ("学    号", metadata.get("student-id", ""), "年级", format_grade_class(metadata)),
        ("指导教师", metadata.get("advisor", ""), "职称", metadata.get("advisor-title", "")),
    )
    for row_index in range(3):
        merged = table.cell(row_index, 1).merge(table.cell(row_index, 3))
        _set_single_cell_paragraph(
            table.cell(row_index, 0),
            document,
            STYLE_REVIEW_INFO_LABEL,
            rows[row_index][0],
        )
        _set_single_cell_paragraph(
            merged,
            document,
            STYLE_REVIEW_INFO_VALUE,
            rows[row_index][1],
        )

    for row_index in range(3, 5):
        for column_index, value in enumerate(rows[row_index]):
            _set_single_cell_paragraph(
                table.cell(row_index, column_index),
                document,
                (
                    STYLE_REVIEW_INFO_LABEL
                    if column_index % 2 == 0
                    else STYLE_REVIEW_INFO_VALUE
                ),
                value,
            )

    for row in table.rows:
        row.height = Pt(34)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    date = _insert_paragraph_before(document, anchor, STYLE_REVIEW_DATE)
    date.text = metadata.get("date", "")
    date.paragraph_format.space_before = Pt(68)


def _postprocess_review_docx(
    input_docx: Path,
    output_docx: Path,
    *,
    metadata: dict[str, str],
    project_root: Path,
    main_dir: Path,
    metadata_dir: Path,
    discipline: str,
    citation_mode: str,
    labels: LabelRegistry,
    citations: CitationRegistry,
    bibliography_items: list[dict[str, object]],
    table_layouts: list[LatexTableLayout],
    word_figure_sequence: str,
    word_table_sequence: str,
) -> None:
    document = Document(input_docx)
    for paragraph in list(document.paragraphs):
        if paragraph.text.strip().startswith("YIBIN_INTERNAL_CITATION_SEED"):
            paragraph._p.getparent().remove(paragraph._p)
    marker = next(
        (p for p in document.paragraphs if p.text.strip() == SECTION_REVIEW_COVER_END),
        None,
    )
    if marker is None:
        raise BuildError("Word 文献综述信息页分节标记缺失。")
    body_sect_pr = document._element.body.sectPr
    if body_sect_pr is None:
        raise BuildError("Pandoc 输出缺少正文分节属性。")
    references_marker = next(
        (p for p in document.paragraphs if p.text.strip() == SECTION_REVIEW_REFERENCES),
        None,
    )
    tail_marker = next(
        (p for p in document.paragraphs if p.text.strip() == SECTION_REVIEW_TAIL),
        None,
    )
    if references_marker is None or tail_marker is None:
        raise BuildError("文献综述参考文献或尾部分节标记缺失。")
    _insert_review_cover(
        document,
        marker,
        metadata,
        project_root=project_root,
        main_dir=main_dir,
        metadata_dir=metadata_dir,
        discipline=discipline,
    )
    _add_section_break(marker, body_sect_pr, None)
    _add_section_break(references_marker, body_sect_pr, None)
    _add_section_break(tail_marker, body_sect_pr, None)
    _apply_common_body_formatting(
        document,
        discipline=discipline,
        labels=labels,
        citations=citations,
        citation_mode=citation_mode,
        table_layouts=table_layouts,
        word_figure_sequence=word_figure_sequence,
        word_table_sequence=word_table_sequence,
    )
    stage = output_docx.with_suffix(".review-sections.docx")
    document.save(stage)
    document = Document(stage)
    if len(document.sections) != 4:
        stage.unlink(missing_ok=True)
        raise BuildError(f"文献综述应生成 4 个 Word 分节，实际为 {len(document.sections)}。")
    cover, body, references, tail = document.sections
    for section in (cover, body, references, tail):
        section.start_type = WD_SECTION_START.NEW_PAGE
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.5)
        section.header_distance = Cm(1.5)
        section.footer_distance = Cm(1.5)
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        _clear_story(section.header)
        _clear_story(section.footer)
    _set_page_number_format(cover, None)
    _set_page_number_format(body, "decimal", 1)
    _set_page_number_format(references, "decimal")
    _set_page_number_format(tail, "decimal")
    body_header = body.header.paragraphs[0]
    body_header.style = document.styles["Header"]
    body_header.text = metadata.get("title", "")
    body_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_header_border(body_header)
    for run in body_header.runs:
        _set_run_fonts(run, "SimSun", "Times New Roman", 9)
    _add_profile_page_field(body, document)
    _add_profile_page_field(references, document)
    _add_profile_page_field(tail, document)

    document.core_properties.title = metadata.get("title", "")
    document.core_properties.author = metadata.get("author", "")
    document.core_properties.subject = "宜宾学院本科毕业论文（设计）文献综述"
    document.core_properties.keywords = "YibinThesis;document-type=literature-review;template-year=2024"
    settings = document.settings.element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_docx)
    if citation_mode == "native":
        _inject_bibliography_sources(output_docx, citations, bibliography_items)
    _normalize_embedded_pngs(output_docx)
    Document(output_docx)
    stage.unlink(missing_ok=True)


def _build_non_thesis(
    args: argparse.Namespace,
    *,
    profile: DocumentProfile,
    project_root: Path,
    main_path: Path,
    main_text: str,
    main_dir: Path,
    metadata_path: Path,
    metadata: dict[str, str],
    pandoc: Path,
    reference_doc: Path,
    csl: Path,
    bibliography: Path,
    output: Path,
    allow_project_fallback: bool,
) -> Path:
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

    expanded_main = expand_nested_inputs(
        main_path,
        project_root,
        allow_project_fallback=allow_project_fallback,
    )
    temporary_root = output.parent / ".tmp"
    temporary_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(
        prefix=f"yibinthesis-{profile.document_type}-word-",
        dir=temporary_root,
    ) as temporary:
        temp_dir = Path(temporary)
        if profile.document_type == "proposal":
            if "\\makeyibinproposal" not in strip_tex_comments(main_text):
                raise BuildError("开题报告入口未调用 \\makeyibinproposal。")
            fields = extract_proposal_fields(expanded_main)
            main_nocites = [
                event.value
                for event in parse_main_events(main_text)
                if event.kind == "nocite" and event.value
            ]
            bibliography_requested = False
            for index, (key, field_text) in enumerate(fields):
                bibliography_requested = bibliography_requested or (
                    "\\printyibinproposalbibliography" in field_text
                )
                field_text = field_text.replace("\\printyibinproposalbibliography", "")
                field_text = normalize_proposal_latex(field_text)
                detect_unsupported(field_text, main_path, warnings)
                # Each official form field is an independent outline.  A
                # ``\section`` inside every cell must therefore start at
                # （一） instead of continuing from the previous field.
                field_counters = HeadingCounters()
                chapter_before = field_counters.chapter
                converted = latex_to_markdown(
                    normalize_latex(
                        field_text,
                        notes,
                        profile.discipline,
                        labels,
                        citations,
                        table_layouts,
                    ),
                    pandoc=pandoc,
                    cwd=main_dir,
                    temp_dir=temp_dir,
                    name=f"proposal-{index}-{key}",
                )
                converted = demote_proposal_headings(converted)
                converted = apply_heading_numbering(
                    converted,
                    "humanities",
                    field_counters,
                    labels,
                )
                converted = apply_float_numbering(
                    converted,
                    "humanities",
                    chapter_before,
                    float_counters,
                    labels,
                    image_target_resolver,
                )
                markdown_parts.append(
                    styled_marker(PROPOSAL_FIELD_MARKER_PREFIX + key.upper())
                )
                markdown_parts.append(converted)
                if key == "references" and bibliography_requested:
                    for keys in main_nocites:
                        citations.add_nocite(keys)
                    seed = citation_seed_markdown(citations)
                    if seed:
                        markdown_parts.append(seed)
                    markdown_parts.append(bibliography_markdown())
        else:
            if "\\makeyibinliteraturereviewcover" not in strip_tex_comments(main_text):
                raise BuildError(
                    "文献综述入口未调用 \\makeyibinliteraturereviewcover。"
                )
            markdown_parts.append(styled_marker(SECTION_REVIEW_COVER_END))
            events = parse_main_events(main_text)
            bibliography_inserted = False
            for index, event in enumerate(events):
                if event.kind in {
                    "frontmatter",
                    "mainmatter",
                    "backmatter",
                    "makeyibinliteraturereviewcover",
                    "makeyibinproposal",
                }:
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
                    chapter_before = counters.chapter
                    converted = latex_to_markdown(
                        normalize_latex(
                            text,
                            notes,
                            profile.discipline,
                            labels,
                            citations,
                            table_layouts,
                        ),
                        pandoc=pandoc,
                        cwd=main_dir,
                        temp_dir=temp_dir,
                        name=f"review-{index}",
                    )
                    converted = apply_heading_numbering(
                        converted,
                        profile.discipline,
                        counters,
                        labels,
                    )
                    converted = apply_float_numbering(
                        converted,
                        profile.discipline,
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
                    markdown_parts.append(styled_marker(SECTION_REVIEW_REFERENCES))
                    markdown_parts.append(bibliography_markdown())
                    bibliography_inserted = True
            if not bibliography_inserted:
                seed = citation_seed_markdown(citations)
                if seed:
                    markdown_parts.append(seed)
                markdown_parts.append(styled_marker(SECTION_REVIEW_REFERENCES))
                markdown_parts.append(bibliography_markdown())
            markdown_parts.append(styled_marker(SECTION_REVIEW_TAIL))

        assembled = "\n\n".join(part for part in markdown_parts if part.strip()) + "\n"
        assembled = labels.resolve(assembled)
        markdown_file = temp_dir / "assembled.md"
        markdown_file.write_text(assembled, encoding="utf-8", newline="\n")
        if args.keep_intermediate:
            kept = output.with_suffix(".pandoc.md")
            kept.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(markdown_file, kept)

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
                bibliography_items = [
                    item for item in parsed_items if isinstance(item, dict)
                ]
        pandoc_docx = temp_dir / "pandoc.docx"
        resource_paths = os.pathsep.join(str(path) for path in resource_candidates)
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
        _normalize_embedded_pngs(pandoc_docx)
        if profile.document_type == "proposal":
            _postprocess_proposal_docx(
                pandoc_docx,
                output,
                metadata=metadata,
                project_root=project_root,
                main_dir=main_dir,
                metadata_dir=metadata_path.parent,
                discipline=profile.discipline,
                citation_mode=args.citation_mode,
                labels=labels,
                citations=citations,
                bibliography_items=bibliography_items,
                table_layouts=table_layouts,
                word_figure_sequence=args.word_figure_sequence,
                word_table_sequence=args.word_table_sequence,
            )
        else:
            _postprocess_review_docx(
                pandoc_docx,
                output,
                metadata=metadata,
                project_root=project_root,
                main_dir=main_dir,
                metadata_dir=metadata_path.parent,
                discipline=profile.discipline,
                citation_mode=args.citation_mode,
                labels=labels,
                citations=citations,
                bibliography_items=bibliography_items,
                table_layouts=table_layouts,
                word_figure_sequence=args.word_figure_sequence,
                word_table_sequence=args.word_table_sequence,
            )

    print(f"Generated: {output}")
    print("Word 输出可编辑；字段将在 Microsoft Word 打开或刷新时更新。")
    for warning in dict.fromkeys(warnings):
        print(f"WARNING: {warning}", file=sys.stderr)
    return output
