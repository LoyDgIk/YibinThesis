"""Shared OOXML, cover, table, field, and citation operations."""

from __future__ import annotations

import sys


def bind_core(core):
    module = sys.modules[__name__]
    for name in dir(core):
        if name not in {"bind_core", "sys"}:
            setattr(module, name, getattr(core, name))
    return module


def _insert_paragraph_before(document: Document, anchor, style_name: str):
    paragraph = document.add_paragraph(style=style_name)
    anchor._p.addprevious(paragraph._p)
    return paragraph


def _set_tabs(paragraph, stops: Iterable[tuple[int, str, str | None]]) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    existing = p_pr.find(qn("w:tabs"))
    if existing is not None:
        p_pr.remove(existing)
    tabs = OxmlElement("w:tabs")
    for position, alignment, leader in stops:
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), alignment)
        tab.set(qn("w:pos"), str(position))
        if leader:
            tab.set(qn("w:leader"), leader)
        tabs.append(tab)
    p_pr.append(tabs)


def _add_tab(paragraph) -> None:
    paragraph.add_run().add_tab()


def _add_cover_label(paragraph, value: str) -> None:
    run = paragraph.add_run(value)
    run.style = "CoverLabel"
    _set_run_fonts(run, "SimHei", "SimHei", 18, bold=True, underline=False)
    if value.startswith("指导教师"):
        # Eight 18 pt CJK glyphs occupy the official 144 pt label slot exactly.
        # A tiny style-neutral character condensation prevents Word's table end
        # mark from wrapping the final two glyphs without changing the font size.
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:val"), "-4")
        run._r.get_or_add_rPr().append(spacing)


def _add_cover_value(paragraph, value: str, *, suffix: str = "") -> None:
    # A zero-width character keeps an intentionally empty slot alive when Word
    # opens and saves the layout table.  It is invisible and, unlike NBSP or
    # ordinary spaces, cannot create a second underline.
    visible = value.strip()
    run = paragraph.add_run((visible if visible else "\u200b") + suffix)
    run.style = "CoverValue"
    _set_run_fonts(run, "SimSun", "SimSun", 16, bold=True, underline=False)


def _prepare_cover_field(paragraph, stops: Iterable[tuple[int, str, str | None]]) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.first_line_indent = Pt(0)
    paragraph.paragraph_format.left_indent = Pt(0)
    paragraph.paragraph_format.right_indent = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.keep_together = True
    _set_tabs(paragraph, stops)


def _set_paragraph_bottom_border(target) -> None:
    """Give a paragraph or paragraph style one Word-native bottom border."""

    p_pr = target._element.get_or_add_pPr()
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


def _hide_run(run) -> None:
    r_pr = run._r.get_or_add_rPr(); vanish = OxmlElement("w:vanish"); r_pr.append(vanish)


def _set_fixed_table_geometry(table, widths_cm: Iterable[float], *, style_name: str) -> None:
    """Give Word a complete fixed-width grid instead of width hints.

    Setting only ``cell.width`` leaves the original equal-column ``tblGrid``
    created by python-docx in place.  Word then trusts that grid on save and
    can collapse or expand the cover columns.  Keep tblW, tblGrid and every
    tcW in exact agreement.
    """
    widths = [int(Cm(width).twips) for width in widths_cm]
    tbl_pr = table._tbl.tblPr
    style = tbl_pr.find(qn("w:tblStyle"))
    if style is None:
        style = OxmlElement("w:tblStyle")
        tbl_pr.insert(0, style)
    style.set(qn("w:val"), style_name)

    table_width = tbl_pr.find(qn("w:tblW"))
    if table_width is None:
        table_width = OxmlElement("w:tblW")
        tbl_pr.append(table_width)
    table_width.set(qn("w:type"), "dxa")
    table_width.set(qn("w:w"), str(sum(widths)))

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
    for edge in ("top", "left", "bottom", "right"):
        margin = cell_margins.find(qn(f"w:{edge}"))
        if margin is None:
            margin = OxmlElement(f"w:{edge}")
            cell_margins.append(margin)
        margin.set(qn("w:type"), "dxa")
        margin.set(qn("w:w"), "0")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(width))
        grid.append(column)

    for row in table.rows:
        cant_split = row._tr.get_or_add_trPr().find(qn("w:cantSplit"))
        if cant_split is None:
            row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for cell, width in zip(row.cells, widths, strict=True):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:type"), "dxa")
            tc_w.set(qn("w:w"), str(width))


def _set_table_grid_widths_twips(table, widths: list[int]) -> None:
    """Set a table's total/grid/cell widths to one internally consistent grid."""

    tbl_pr = table._tbl.tblPr
    table_width = tbl_pr.find(qn("w:tblW"))
    if table_width is None:
        table_width = OxmlElement("w:tblW")
        tbl_pr.append(table_width)
    table_width.set(qn("w:type"), "dxa")
    table_width.set(qn("w:w"), str(sum(widths)))
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    indent = tbl_pr.find(qn("w:tblInd"))
    if indent is None:
        indent = OxmlElement("w:tblInd")
        tbl_pr.append(indent)
    indent.set(qn("w:type"), "dxa")
    indent.set(qn("w:w"), "0")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(width))
        grid.append(column)
    for row in table.rows:
        for cell, width in zip(row.cells, widths, strict=True):
            tc_w = cell._tc.get_or_add_tcPr().find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                cell._tc.get_or_add_tcPr().append(tc_w)
            tc_w.set(qn("w:type"), "dxa")
            tc_w.set(qn("w:w"), str(width))


def _fit_nested_tables_to_cell(cell) -> None:
    """Scale moved Pandoc tables to the proposal response slot width."""

    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        return
    available = int(tc_w.get(qn("w:w"), "0"))
    margins = tc_pr.find(qn("w:tcMar"))
    if margins is not None:
        for edge in ("left", "right"):
            node = margins.find(qn(f"w:{edge}"))
            if node is not None:
                available -= int(node.get(qn("w:w"), "0"))
    if available <= 0:
        return

    for nested in cell.tables:
        raw = [
            int(column.get(qn("w:w"), "0"))
            for column in nested._tbl.tblGrid.findall(qn("w:gridCol"))
        ]
        total = sum(raw)
        if not raw or total <= available:
            continue
        widths = [max(1, round(width * available / total)) for width in raw]
        widths[-1] += available - sum(widths)
        _set_table_grid_widths_twips(nested, widths)


def _set_table_borders_none(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    else:
        for child in list(borders):
            borders.remove(child)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "nil")
        borders.append(node)


def _insert_layout_separator(anchor) -> None:
    """Prevent Word from merging adjacent layout tables on open/save."""
    paragraph = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "1")
    spacing.set(qn("w:lineRule"), "exact")
    p_pr.append(spacing)
    mark_properties = OxmlElement("w:rPr")
    mark_properties.append(OxmlElement("w:vanish"))
    mark_size = OxmlElement("w:sz")
    mark_size.set(qn("w:val"), "2")
    mark_properties.append(mark_size)
    mark_size_cs = OxmlElement("w:szCs")
    mark_size_cs.set(qn("w:val"), "2")
    mark_properties.append(mark_size_cs)
    p_pr.append(mark_properties)
    paragraph.append(p_pr)
    anchor._p.addprevious(paragraph)


def _insert_cover_table(
    document: Document,
    anchor,
    cells: list[tuple[str, str, float]],
    *,
    row_height_pt: float,
    hidden_tokens: str = "",
):
    """Insert one official cover row as a borderless layout table."""
    table = document.add_table(rows=1, cols=len(cells))
    anchor._p.addprevious(table._tbl)
    # The official cover starts every field row at the 3 cm text margin, while
    # the underline length differs from row to row.  Left alignment preserves
    # that common origin; centring narrower rows shifts every underline.
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    _set_fixed_table_geometry(table, (cell[2] for cell in cells), style_name="YibinCoverLayout")
    _set_table_borders_none(table)
    row = table.rows[0]
    row.height = Pt(row_height_pt)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    for index, (kind, value, width_cm) in enumerate(cells):
        cell = table.cell(0, index)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0); p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.keep_together = True
        p_pr = p._p.get_or_add_pPr()
        indent = p_pr.find(qn("w:ind"))
        if indent is None:
            indent = OxmlElement("w:ind")
            p_pr.append(indent)
        for attribute in ("left", "right", "firstLine", "leftChars", "rightChars", "firstLineChars"):
            indent.set(qn(f"w:{attribute}"), "0")
        if kind == "label":
            if index == 0:
                p.style = document.styles["CoverField"]
                if value.startswith("指导教师"):
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.line_spacing = 1.0
            _add_cover_label(p, value)
        else:
            p.style = document.styles["CoverValueLine"]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _add_cover_value(p, value)
    _insert_layout_separator(anchor)
    return table


def _calibrate_cover_styles(document: Document) -> None:
    """Apply the official first-page rhythm through reusable Word styles."""

    logo = document.styles["CoverLogo"].paragraph_format
    logo.space_before = Pt(2.0)
    logo.space_after = Pt(0)

    thesis_type = document.styles["CoverThesisType"].paragraph_format
    thesis_type.space_before = Pt(24.35)
    thesis_type.space_after = Pt(52.35)

    version = document.styles["CoverVersion"].paragraph_format
    version.space_before = Pt(0)
    version.space_after = Pt(23)

    title = document.styles["CoverTitle"].paragraph_format
    title.space_before = Pt(0)
    title.space_after = Pt(47.3)
    title.line_spacing = 2.0

    title_with_version = document.styles["CoverTitleWithVersion"].paragraph_format
    title_with_version.space_before = Pt(0)
    title_with_version.space_after = Pt(10)
    title_with_version.line_spacing = 2.0

    field = document.styles["CoverField"].paragraph_format
    field.space_before = Pt(0)
    field.space_after = Pt(0)
    field.line_spacing = 1.0

    try:
        value_line_style = document.styles["CoverValueLine"]
    except KeyError:
        value_line_style = document.styles.add_style(
            "CoverValueLine", WD_STYLE_TYPE.PARAGRAPH
        )
    value_line_style.base_style = document.styles["Normal"]
    value_line = value_line_style.paragraph_format
    value_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    value_line.first_line_indent = Pt(0)
    value_line.left_indent = Pt(0)
    value_line.right_indent = Pt(0)
    value_line.space_before = Pt(0)
    value_line.space_after = Pt(0)
    # An exact 20 pt line box keeps the rule at the official underline height
    # even when the slot is empty and contains only the zero-width keeper.
    value_line.line_spacing = Pt(20.0)
    value_line.keep_together = True
    _set_paragraph_bottom_border(value_line_style)


def _insert_declaration_table(
    document: Document,
    anchor,
    cells: list[tuple[str, str | Path | None, float]],
    *,
    space_before: float = 0,
    signature_background: str = "preserve",
):
    """Insert fixed signature/date slots that survive a Word save."""
    table = document.add_table(rows=1, cols=len(cells))
    anchor._p.addprevious(table._tbl)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_fixed_table_geometry(table, (cell[2] for cell in cells), style_name="YibinFrontLayout")
    _set_table_borders_none(table)
    for index, (kind, value, width_cm) in enumerate(cells):
        cell = table.cell(0, index)
        paragraph = cell.paragraphs[0]
        paragraph.style = document.styles[
            "DeclarationDate" if kind == "date" else "DeclarationSignature"
        ]
        paragraph.paragraph_format.first_line_indent = Pt(0)
        paragraph.paragraph_format.space_before = Pt(space_before)
        paragraph.paragraph_format.space_after = Pt(0)
        p_pr = paragraph._p.get_or_add_pPr()
        indent = p_pr.find(qn("w:ind"))
        if indent is None:
            indent = OxmlElement("w:ind")
            p_pr.append(indent)
        for attribute in ("left", "right", "firstLine", "leftChars", "rightChars", "firstLineChars"):
            indent.set(qn(f"w:{attribute}"), "0")
        if kind == "label":
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if kind == "signature":
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Keep the rule attached to the signature paragraph.  A cell
            # border sits at the bottom of the tallest image row and therefore
            # drifts visibly downward compared with Word's official line.
            _set_paragraph_bottom_border(paragraph)
            if isinstance(value, Path):
                picture_run = paragraph.add_run()
                picture_source: str | io.BytesIO = str(value)
                if signature_background == "whiten":
                    picture_source = whiten_signature_image(value)
                picture_run.add_picture(
                    picture_source,
                    width=Cm(max(0.5, min(width_cm - 0.2, 2.6))),
                )
                for doc_property in picture_run._r.iter(qn("wp:docPr")):
                    doc_property.attrib.pop("descr", None)
            else:
                paragraph.add_run("\u200b")
        else:
            run = paragraph.add_run(str(value) if value else "\u200b")
            _set_run_fonts(run, "SimSun", "SimSun", 12, bold=False, underline=False)
    _insert_layout_separator(anchor)
    return table


def _resolve_cover_logo(
    metadata: dict[str, str],
    project_root: Path,
    main_dir: Path,
    metadata_dir: Path,
    *,
    allow_project_fallback: bool,
) -> Path | None:
    configured = metadata.get("logo", "builtin").strip()
    normalized = configured.casefold()
    if normalized in {"builtin", "built-in", "default"}:
        logo = (project_root / "assets" / "yibin-university-logo.png").resolve()
        if not logo.is_file():
            raise BuildError(f"模板内置封面校徽不存在：{logo}")
        return logo
    if not configured or normalized == "none":
        return None
    configured_path = Path(configured).expanduser()
    if configured_path.is_absolute():
        candidates = [configured_path]
    else:
        candidates = [
            main_dir / configured_path,
            metadata_dir / configured_path,
        ]
    if allow_project_fallback:
        candidates.extend(
            [
                project_root / "assets" / "yibin-university-logo.png",
                project_root / "assets" / "yibin-logo.png",
            ]
        )
    logo = next(
        (
            candidate.resolve()
            for candidate in dict.fromkeys(candidates)
            if candidate.is_file()
        ),
        None,
    )
    if logo is None and configured and not allow_project_fallback:
        raise BuildError(
            f"外部入口声明的封面校徽不存在：{configured}；"
            "不会回退到模板目录中的同名资源。"
        )
    return logo


def _resolve_signature_asset(
    metadata: dict[str, str],
    key: str,
    main_dir: Path,
    metadata_dir: Path,
) -> Path | None:
    configured = metadata.get(key, "").strip()
    if not configured or configured.casefold() in {"none", "null"}:
        return None
    configured_path = Path(configured).expanduser()
    candidates = (
        [configured_path]
        if configured_path.is_absolute()
        else [metadata_dir / configured_path, main_dir / configured_path]
    )
    asset = next(
        (
            candidate.resolve()
            for candidate in dict.fromkeys(candidates)
            if candidate.is_file()
        ),
        None,
    )
    if asset is None:
        raise BuildError(f"元数据 {key} 声明的签名图片不存在：{configured}")
    return asset


def _build_cover_page(
    document: Document,
    anchor,
    metadata: dict[str, str],
    project_root: Path,
    main_dir: Path,
    metadata_dir: Path,
    *,
    allow_project_fallback: bool,
    page_break_after: bool,
) -> None:
    _calibrate_cover_styles(document)
    logo = _resolve_cover_logo(
        metadata,
        project_root,
        main_dir,
        metadata_dir,
        allow_project_fallback=allow_project_fallback,
    )
    logo_paragraph = _insert_paragraph_before(document, anchor, "CoverLogo")
    logo_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo_paragraph.paragraph_format.first_line_indent = Pt(0)
    if logo is not None:
        picture_run = logo_paragraph.add_run()
        picture_run.add_picture(str(logo), width=Cm(13.15), height=Cm(3.65))
        for doc_property in picture_run._r.iter(qn("wp:docPr")):
            doc_property.attrib.pop("descr", None)
    else:
        run = logo_paragraph.add_run("宜宾学院")
        _set_run_fonts(run, "SimHei", "Times New Roman", 26, bold=True)

    thesis_type = _insert_paragraph_before(document, anchor, "CoverThesisType")
    thesis_type.alignment = WD_ALIGN_PARAGRAPH.CENTER
    thesis_type.paragraph_format.first_line_indent = Pt(0)
    type_run = thesis_type.add_run("本科生毕业论文（设计）")
    _set_run_fonts(type_run, "SimHei", "SimHei", 28, bold=True, underline=False)

    version_value = metadata.get("version", "").strip()
    title_style = "CoverTitleWithVersion" if version_value else "CoverTitle"
    title_paragraph = _insert_paragraph_before(document, anchor, title_style)
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.first_line_indent = Pt(0)
    title_run = title_paragraph.add_run(metadata.get("title", "").strip() or "（填写中文题目）")
    _set_run_fonts(title_run, "SimHei", "SimHei", 18, bold=True, underline=True)

    if version_value:
        version_paragraph = _insert_paragraph_before(document, anchor, "CoverVersion")
        version_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        version_paragraph.paragraph_format.first_line_indent = Pt(0)
        version_run = version_paragraph.add_run(version_value)
        _set_run_fonts(version_run, "SimHei", "SimHei", 14, bold=False, underline=False)

    # Widths are distilled from the official Word cover at page coordinates
    # (points): college 84.85|183.85|513.00, major 84.85|183.85|522.00,
    # student 84.85|174.85|513.00.  Each value cell owns the only visible rule.
    _insert_cover_table(
        document,
        anchor,
        [("label", "学院（部）", 3.4925), ("value", metadata.get("college", ""), 11.6121)],
        row_height_pt=46.30,
    )
    _insert_cover_table(
        document,
        anchor,
        [("label", "专    业", 3.4925), ("value", metadata.get("major", ""), 11.9296)],
        row_height_pt=46.25,
    )
    _insert_cover_table(
        document,
        anchor,
        [("label", "学生姓名", 3.1750), ("value", metadata.get("author", ""), 11.9296)],
        row_height_pt=47.60,
    )

    _insert_cover_table(
        document, anchor,
        [("label", "学    号", 3.1750), ("value", metadata.get("student-id", ""), 5.6686),
         ("label", "年级", 1.5875), ("value", format_grade_class(metadata), 4.5367)],
        row_height_pt=46.25,
    )

    advisor_rows = (
        (
            "指导教师（校内）",
            metadata.get("advisor", ""),
            metadata.get("advisor-title", ""),
        ),
        (
            "指导教师（校外）",
            metadata.get("external-advisor", ""),
            metadata.get("external-advisor-title", ""),
        ),
    )
    last_table = None
    for label, person, professional_title in advisor_rows:
        last_table = _insert_cover_table(
            document, anchor,
            [("label", label, 5.0800), ("value", person, 5.1259),
             ("label", "职称", 1.5416), ("value", professional_title, 3.1289)],
            row_height_pt=46.30,
        )

    date_value = metadata.get("date", "").strip()
    if date_value:
        date_paragraph = _insert_paragraph_before(document, anchor, "CoverDate")
        date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_paragraph.paragraph_format.first_line_indent = Pt(0)
        date_run = date_paragraph.add_run(date_value)
        _set_run_fonts(date_run, "SimSun", "Times New Roman", 15, bold=True, underline=False)


def _estimate_declaration_lines(text: str) -> int:
    units = sum(0.5 if ord(character) < 128 else 1.0 for character in text)
    if units <= 28:
        return 1
    return 1 + int((units - 28 + 30.999) // 31)


def _add_declaration_text(document: Document, anchor, style_name: str, text: str):
    paragraph = _insert_paragraph_before(document, anchor, style_name)
    run = paragraph.add_run(text)
    _set_run_fonts(run, "SimSun", "SimSun", 12, bold=False, underline=False)
    return paragraph


def _build_originality_page(
    document: Document,
    anchor,
    metadata: dict[str, str],
    author_signature: Path | None,
    signature_background: str,
) -> None:
    heading = _insert_paragraph_before(document, anchor, "DeclarationTitle")
    heading.paragraph_format.page_break_before = True
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    heading_run = heading.add_run("原创性声明")
    _set_run_fonts(heading_run, "SimHei", "SimHei", 16, bold=True, underline=False)

    title = metadata.get("title", "").strip() or "（填写中文题目）"
    statement = (
        f"本人呈交的学位论文（设计）《{title}》，是在导师的指导下，独立进行研究取得的成果。"
        "除文中已经注明引用的内容外，本论文（设计）不包括其他个人或集体已经发表或撰写过的作品成果。"
        "对本文（设计）做出贡献的个人和集体，均已在文中以明确方式标明。"
        "本人完全意识到本声明的法律后果，因本声明而产生的法律后果由本人承担。"
    )
    body = _add_declaration_text(document, anchor, "DeclarationBody", statement)
    body.alignment = WD_ALIGN_PARAGRAPH.LEFT
    body.paragraph_format.first_line_indent = Pt(24)

    extra_lines = max(0, _estimate_declaration_lines(statement) - 5)
    _insert_declaration_table(
        document,
        anchor,
        [
            ("spacer", "", 0.85),
            ("label", "学位论文作者：", 3.6),
            ("signature", author_signature, 3.0),
            ("spacer", "", 8.05),
        ],
        space_before=max(0, 36 - extra_lines * 31.2),
        signature_background=signature_background,
    )

    date = _add_declaration_text(
        document,
        anchor,
        "DeclarationDate",
        "日期：    年   月   日",
    )
    date.paragraph_format.first_line_indent = Pt(24)

    lead = _insert_paragraph_before(document, anchor, "DeclarationRegulationLead")
    lead_run = lead.add_run("附：")
    lead_run.style = "DeclarationRegulationLeadLabel"
    _set_run_fonts(lead_run, "SimSun", "SimSun", 16, bold=True, underline=False)
    regulation_title = lead.add_run(
        "《普通高等学校学生管理规定》（中华人民共和国教育部令第41号）"
    )
    regulation_title.style = "DeclarationRegulationText"
    _set_run_fonts(regulation_title, "SimSun", "SimSun", 14, bold=True, underline=False)

    clause = _insert_paragraph_before(document, anchor, "DeclarationRegulationClause")
    clause_run = clause.add_run(
        "第五十二条\u00a0学生有下列情形之一，学校可以给予开除学籍处分："
    )
    clause_run.style = "DeclarationRegulationText"
    _set_run_fonts(clause_run, "SimSun", "SimSun", 14, bold=True, underline=False)

    item = _insert_paragraph_before(document, anchor, "DeclarationRegulationItem")
    item_run = item.add_run(
        "（五）学位论文、公开发表的研究成果存在抄袭、篡改、伪造等学术不端行为，"
        "情节严重的，或者代写论文、买卖论文的；"
    )
    item_run.style = "DeclarationRegulationText"
    _set_run_fonts(item_run, "SimSun", "SimSun", 14, bold=True, underline=False)
    item.add_run().add_break(WD_BREAK.PAGE)


def _build_authorization_page(
    document: Document,
    anchor,
    metadata: dict[str, str],
    author_signature: Path | None,
    advisor_signature: Path | None,
    signature_background: str,
) -> None:
    heading = _insert_paragraph_before(document, anchor, "DeclarationTitle")
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    heading_run = heading.add_run("学位论文（设计）版权使用授权书")
    _set_run_fonts(heading_run, "SimHei", "SimHei", 16, bold=True, underline=False)

    authorization = (
        "本学位论文（设计）作者完全了解学校有关保留、使用学位论文（设计）的规定，"
        "同意学校保留并向国家有关部门或机构送交论文（设计）的复印件和电子版，允许论文（设计）"
        "被查阅和借阅。本人授权宜宾学院将本学位论文（设计）的全部或部分内容编入有关数据库进行检索，"
        "可以采用影印、缩印或扫描等复制手段保存和汇编。"
    )
    body = _add_declaration_text(document, anchor, "DeclarationBody", authorization)
    body.alignment = WD_ALIGN_PARAGRAPH.LEFT
    body.paragraph_format.first_line_indent = Pt(24)

    choice_intro = _insert_paragraph_before(document, anchor, "DeclarationBody")
    choice_intro.paragraph_format.first_line_indent = Pt(24)
    choice_intro.paragraph_format.line_spacing = 1.5
    choice_intro.paragraph_format.space_before = Pt(36)
    intro_run = choice_intro.add_run("本学位论文（设计）属于")
    _set_run_fonts(intro_run, "SimSun", "SimSun", 12, bold=False, underline=False)
    instruction = choice_intro.add_run("（请在以下相应方框内打“√”）作品")
    _set_run_fonts(instruction, "SimSun", "SimSun", 12, bold=False, underline=True)

    confidential = metadata.get("secrecy", "public").strip().lower() == "confidential"
    declassify = metadata.get("declassify-year", "").strip() or "____"
    for text in (
        f"保  密{'■' if confidential else '□'}，在 {declassify} 年解密后适用本授权书。",
        f"不保密{'□' if confidential else '■'}。",
    ):
        paragraph = _add_declaration_text(document, anchor, "DeclarationBody", text)
        paragraph.paragraph_format.first_line_indent = Pt(24)
        paragraph.paragraph_format.line_spacing = 1.5

    _insert_declaration_table(
        document,
        anchor,
        [
            ("label", "作者（签名）：", 3.5),
            ("signature", author_signature, 2.8),
            ("spacer", "", 0.2),
            ("label", "指导教师（签名）：", 4.2),
            ("signature", advisor_signature, 2.8),
        ],
        space_before=72,
        signature_background=signature_background,
    )
    _insert_declaration_table(
        document,
        anchor,
        [
            ("date", "日期：    年   月   日", 6.2),
            ("spacer", "", 1.3),
            ("date", "日期：    年   月   日", 6.0),
        ],
    )


def _build_front_matter(
    document: Document,
    anchor,
    metadata: dict[str, str],
    project_root: Path,
    main_dir: Path,
    metadata_dir: Path,
    *,
    allow_project_fallback: bool,
    include_cover: bool,
    include_declarations: bool,
) -> None:
    if include_cover:
        _build_cover_page(
            document,
            anchor,
            metadata,
            project_root,
            main_dir,
            metadata_dir,
            allow_project_fallback=allow_project_fallback,
            page_break_after=include_declarations,
        )
    if include_declarations:
        signature_background = metadata.get("signature-background", "preserve").strip().casefold()
        if signature_background not in {"preserve", "whiten"}:
            raise BuildError(
                "元数据 signature-background 只能是 preserve 或 whiten："
                + signature_background
            )
        author_signature = _resolve_signature_asset(
            metadata,
            "author-signature",
            main_dir,
            metadata_dir,
        )
        advisor_signature = _resolve_signature_asset(
            metadata,
            "advisor-signature",
            main_dir,
            metadata_dir,
        )
        _build_originality_page(
            document,
            anchor,
            metadata,
            author_signature,
            signature_background,
        )
        _build_authorization_page(
            document,
            anchor,
            metadata,
            author_signature,
            advisor_signature,
            signature_background,
        )


def _clear_direct_paragraph_format(paragraph, *names: str) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    for name in names:
        node = p_pr.find(qn(f"w:{name}"))
        if node is not None:
            p_pr.remove(node)


def _clear_direct_run_typography(run) -> None:
    r_pr = run._r.find(qn("w:rPr"))
    if r_pr is None:
        return
    for name in ("rFonts", "sz", "szCs"):
        node = r_pr.find(qn(f"w:{name}"))
        if node is not None:
            r_pr.remove(node)
    if len(r_pr) == 0:
        run._r.remove(r_pr)


def _clear_title_run_typography(run) -> None:
    r_pr = run._r.find(qn("w:rPr"))
    if r_pr is None:
        return
    for name in (
        "rFonts",
        "b",
        "bCs",
        "i",
        "iCs",
        "color",
        "u",
        "sz",
        "szCs",
        "highlight",
        "caps",
        "smallCaps",
        "strike",
        "dstrike",
        "vertAlign",
        "spacing",
        "position",
        "kern",
    ):
        node = r_pr.find(qn(f"w:{name}"))
        if node is not None:
            r_pr.remove(node)
    if len(r_pr) == 0:
        run._r.remove(r_pr)


def _next_numbering_id(numbering, tag: str, attribute: str) -> int:
    values = [
        int(value)
        for element in numbering.findall(qn(f"w:{tag}"))
        if (value := element.get(qn(f"w:{attribute}"))) is not None
        and value.isdigit()
    ]
    return max(values, default=0) + 1


def _set_style_numbering(style, num_id: int, level: int) -> None:
    p_pr = style._element.get_or_add_pPr()
    existing = p_pr.find(qn("w:numPr"))
    if existing is not None:
        p_pr.remove(existing)
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(level))
    num = OxmlElement("w:numId")
    num.set(qn("w:val"), str(num_id))
    num_pr.extend([ilvl, num])
    insertion = 0
    for index, child_element in enumerate(p_pr):
        if child_element.tag == qn("w:pStyle"):
            insertion = index + 1
    p_pr.insert(insertion, num_pr)


def _set_paragraph_numbering(paragraph, num_id: int, level: int) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    existing = p_pr.find(qn("w:numPr"))
    if existing is not None:
        p_pr.remove(existing)
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(level))
    num = OxmlElement("w:numId")
    num.set(qn("w:val"), str(num_id))
    num_pr.extend([ilvl, num])
    p_pr.append(num_pr)


def _clear_paragraph_numbering(paragraph) -> None:
    """Remove a direct list binding when a title becomes unnumbered."""

    p_pr = paragraph._p.get_or_add_pPr()
    num_pr = p_pr.find(qn("w:numPr"))
    if num_pr is not None:
        p_pr.remove(num_pr)


def _apply_list_paragraph_styles(document: Document) -> None:
    """Let one reusable style govern list typography and first-line indent."""

    eligible_styles = {
        STYLE_BODY,
        STYLE_FIRST_PARAGRAPH,
        STYLE_LIST_BODY,
        "Normal",
        "Body Text",
        "Compact",
        "First Paragraph",
        "List Paragraph",
        "List Number",
        "List Number 2",
        "List Number 3",
        "List Bullet",
        "List Bullet 2",
        "List Bullet 3",
    }
    for paragraph in document.paragraphs:
        p_pr = paragraph._p.pPr
        if (
            p_pr is None
            or p_pr.find(qn("w:numPr")) is None
            or paragraph.style.name not in eligible_styles
        ):
            continue
        paragraph.style = document.styles[STYLE_LIST_BODY]
        indent = p_pr.find(qn("w:ind"))
        if indent is None:
            continue
        for attribute in ("w:firstLine", "w:firstLineChars"):
            indent.attrib.pop(qn(attribute), None)
        if not indent.attrib and not list(indent):
            p_pr.remove(indent)


def _new_numbering_instance(document: Document, abstract_id: int) -> int:
    numbering = document.part.numbering_part.element
    num_id = _next_numbering_id(numbering, "num", "numId")
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    numbering.append(num)
    return num_id


def _configure_heading_numbering(document: Document, discipline: str) -> int:
    """Attach a real four-level Word list to the reusable heading styles."""

    numbering = document.part.numbering_part.element
    abstract_id = _next_numbering_id(numbering, "abstractNum", "abstractNumId")
    num_id = _next_numbering_id(numbering, "num", "numId")
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    nsid = OxmlElement("w:nsid")
    nsid.set(qn("w:val"), hashlib.sha1(f"YibinHeading:{discipline}".encode()).hexdigest()[:8].upper())
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.extend([nsid, multi])

    if discipline == "science":
        formats = ["decimal"] * 4
        level_texts = ["%1", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4"]
        suffix = "space"
    else:
        formats = [
            "chineseCounting",
            "chineseCounting",
            "decimal",
            "decimal",
        ]
        level_texts = ["%1、", "（%2）", "%3.", "（%4）"]
        suffix = "nothing"

    style_names = [
        STYLE_HEADING_1,
        STYLE_HEADING_2,
        STYLE_HEADING_3,
        STYLE_HEADING_4,
    ]
    for level, (style_name, number_format, level_text) in enumerate(
        zip(style_names, formats, level_texts)
    ):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(level))
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        num_fmt = OxmlElement("w:numFmt")
        num_fmt.set(qn("w:val"), number_format)
        p_style = OxmlElement("w:pStyle")
        p_style.set(qn("w:val"), document.styles[style_name].style_id)
        suff = OxmlElement("w:suff")
        suff.set(qn("w:val"), suffix)
        text = OxmlElement("w:lvlText")
        text.set(qn("w:val"), level_text)
        justification = OxmlElement("w:lvlJc")
        justification.set(qn("w:val"), "left")
        # Word's default for a multilevel list is to restart each lower level
        # when its immediately higher level advances.  Explicit lvlRestart
        # values generated by hand are easy to get off by one, so retain the
        # native default verified through Word COM.
        lvl.extend([start, num_fmt, p_style, suff, text, justification])
        abstract.append(lvl)
        _set_style_numbering(document.styles[style_name], num_id, level)

    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract)
    else:
        numbering.insert(numbering.index(first_num), abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    numbering.append(num)
    return abstract_id


def _new_appendix_section_numbering(
    document: Document,
    discipline: str,
    marker: str,
) -> int:
    """Create one native list instance that restarts within an appendix."""

    numbering = document.part.numbering_part.element
    abstract_id = _next_numbering_id(numbering, "abstractNum", "abstractNumId")
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    nsid = OxmlElement("w:nsid")
    nsid.set(
        qn("w:val"),
        hashlib.sha1(
            f"YibinAppendixSection:{discipline}:{marker}".encode()
        ).hexdigest()[:8].upper(),
    )
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.extend([nsid, multi])

    level = OxmlElement("w:lvl")
    level.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    number_format = OxmlElement("w:numFmt")
    number_format.set(
        qn("w:val"),
        "decimal" if discipline == "science" else "chineseCounting",
    )
    paragraph_style = OxmlElement("w:pStyle")
    paragraph_style.set(
        qn("w:val"),
        document.styles[STYLE_APPENDIX_SECTION].style_id,
    )
    suffix = OxmlElement("w:suff")
    suffix.set(qn("w:val"), "space")
    level_text = OxmlElement("w:lvlText")
    level_text.set(
        qn("w:val"),
        f"{marker}.%1" if discipline == "science" else "（%1）",
    )
    justification = OxmlElement("w:lvlJc")
    justification.set(qn("w:val"), "left")
    level.extend(
        [
            start,
            number_format,
            paragraph_style,
            suffix,
            level_text,
            justification,
        ]
    )
    abstract.append(level)

    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract)
    else:
        numbering.insert(numbering.index(first_num), abstract)
    return _new_numbering_instance(document, abstract_id)


def _configure_appendix_section_numbering(
    document: Document,
    discipline: str,
) -> None:
    """Apply style-driven, automatically restarting appendix subsection numbers."""

    current_num_id: int | None = None
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if paragraph.style.name == STYLE_APPENDIX_HEADING:
            match = re.match(
                r"^附录(?P<marker>[A-ZＡ-Ｚ0-9一二三四五六七八九十]+)(?:\s|　)",
                text,
            )
            if match is None:
                raise BuildError(f"无法识别附录编号：{text}")
            marker = match.group("marker")
            current_num_id = _new_appendix_section_numbering(
                document,
                discipline,
                marker,
            )
            continue
        if current_num_id is None or paragraph.style.name != STYLE_HEADING_2:
            continue
        paragraph.style = document.styles[STYLE_APPENDIX_SECTION]
        _clear_paragraph_numbering(paragraph)
        _set_paragraph_numbering(paragraph, current_num_id, 0)


def _require_independent_heading_styles(document: Document) -> None:
    """Reject reference documents that alias thesis styles to built-in headings."""

    thesis_styles = (
        STYLE_HEADING_1,
        STYLE_HEADING_2,
        STYLE_HEADING_3,
        STYLE_HEADING_4,
        STYLE_APPENDIX_SECTION,
    )
    missing: list[str] = []
    for style_name in thesis_styles:
        try:
            style = document.styles[style_name]
        except KeyError:
            missing.append(style_name)
            continue
        if style.style_id in BUILTIN_HEADING_STYLE_IDS:
            raise BuildError(
                "reference.docx 将自定义标题样式错误地绑定到 Word 内置样式 "
                f"{style.style_id}：{style_name}；请重新生成参考模板。"
            )
    if missing:
        raise BuildError(
            "reference.docx 缺少独立的论文标题样式：" + "、".join(missing)
        )

    for style_name, style_id in zip(
        BUILTIN_HEADING_STYLES,
        BUILTIN_HEADING_STYLE_IDS,
        strict=True,
    ):
        try:
            style = document.styles[style_name]
        except KeyError as error:
            raise BuildError(
                f"reference.docx 缺少 Word 内置标题样式：{style_name}"
            ) from error
        if style.style_id != style_id:
            raise BuildError(
                "reference.docx 修改了 Word 内置标题样式标识："
                f"{style_name}={style.style_id}，应为 {style_id}。"
            )


def _restart_cell_heading_numbering(
    document: Document,
    cell,
    abstract_id: int,
) -> None:
    """Give one proposal field an independent native heading list instance."""

    levels = {
        STYLE_HEADING_1: 0,
        STYLE_HEADING_2: 1,
        STYLE_HEADING_3: 2,
        STYLE_HEADING_4: 3,
        **{name: index for index, name in enumerate(BUILTIN_HEADING_STYLES)},
    }
    headings = [p for p in cell.paragraphs if p.style.name in levels]
    if not headings:
        return
    num_id = _new_numbering_instance(document, abstract_id)
    for paragraph in headings:
        _set_paragraph_numbering(paragraph, num_id, levels[paragraph.style.name])


def _isolate_humanities_introduction_numbering(
    document: Document,
    abstract_id: int,
) -> None:
    """Keep 绪论 subsections from consuming the first body chapter number."""

    levels = {
        STYLE_HEADING_1: 0,
        STYLE_HEADING_2: 1,
        STYLE_HEADING_3: 2,
        STYLE_HEADING_4: 3,
        **{name: index for index, name in enumerate(BUILTIN_HEADING_STYLES)},
    }
    introduction_num_id: int | None = None
    inside_introduction = False
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if paragraph.style.name == STYLE_UNNUMBERED_HEADING and text == "绪论":
            introduction_num_id = _new_numbering_instance(document, abstract_id)
            inside_introduction = True
            continue
        if not inside_introduction:
            continue
        level = levels.get(paragraph.style.name)
        if level == 0:
            break
        if level is not None and introduction_num_id is not None:
            _set_paragraph_numbering(paragraph, introduction_num_id, level)


def _resolve_table_column_widths(
    layout: LatexTableLayout | None,
    raw_widths: list[int],
    target_width: int,
) -> list[int]:
    if layout is not None and len(layout.columns) == len(raw_widths):
        fractions = [column.width_fraction for column in layout.columns]
        if any(value is not None and value > 0 for value in fractions):
            known_sum = sum(value or 0 for value in fractions)
            unknown = [index for index, value in enumerate(fractions) if value is None]
            if unknown:
                outer_fraction = layout.table_width_fraction or 1.0
                if 0 < known_sum < outer_fraction:
                    fallback = (outer_fraction - known_sum) / len(unknown)
                else:
                    known = [value for value in fractions if value is not None and value > 0]
                    fallback = (sum(known) / len(known)) if known else 1
                weights = [fallback if value is None else value for value in fractions]
            else:
                weights = [value or 0 for value in fractions]
            if sum(weights) > 0:
                scale = target_width / sum(weights)
                widths = [max(1, round(weight * scale)) for weight in weights]
                widths[-1] += target_width - sum(widths)
                return widths

    usable = raw_widths if len(raw_widths) > 0 and sum(raw_widths) > 0 else [1]
    scale = target_width / sum(usable)
    widths = [max(1, round(width * scale)) for width in usable]
    widths[-1] += target_width - sum(widths)
    return widths


def _table_style_for_alignment(row_index: int, horizontal: str) -> str:
    if row_index == 0:
        return {
            "left": STYLE_TABLE_HEADER_LEFT,
            "center": STYLE_TABLE_HEADER,
            "right": STYLE_TABLE_HEADER_RIGHT,
        }.get(horizontal, STYLE_TABLE_HEADER_LEFT)
    return {
        "left": STYLE_TABLE_TEXT,
        "center": STYLE_TABLE_CENTER,
        "right": STYLE_TABLE_RIGHT,
    }.get(horizontal, STYLE_TABLE_TEXT)


def _caption_number_by_table(document: Document) -> dict[object, str]:
    result: dict[object, str] = {}
    body = document._element.body
    children = list(body)
    for index, element in enumerate(children):
        if element.tag != qn("w:tbl"):
            continue
        for previous in reversed(children[max(0, index - 3) : index]):
            if previous.tag == qn("w:tbl"):
                break
            if previous.tag != qn("w:p"):
                continue
            value = "".join(node.text or "" for node in previous.iter(qn("w:t"))).strip()
            match = re.match(r"^表\s*([0-9]+(?:\.[0-9]+)?)\b", value)
            if match:
                result[element] = match.group(1)
                break
    return result


_AHP_SCALE_VALUES = (
    "9",
    "8",
    "7",
    "6",
    "5",
    "4",
    "3",
    "2",
    "1",
    "2",
    "3",
    "4",
    "5",
    "6",
    "7",
    "8",
    "9",
)


def _compact_table_text(value: str) -> str:
    return re.sub(r"\s+", "", value)


def _unique_row_cells(row) -> list[tuple[int, object]]:
    result: list[tuple[int, object]] = []
    seen: set[int] = set()
    for column_index, cell in enumerate(row.cells):
        key = id(cell._tc)
        if key in seen:
            continue
        seen.add(key)
        result.append((column_index, cell))
    return result


def _questionnaire_table_kind(table) -> str | None:
    """Identify questionnaire tables by their semantic headers."""

    if not table.rows:
        return None
    column_count = len(table.columns)
    header_cells = _unique_row_cells(table.rows[0])
    header = [_compact_table_text(cell.text) for _, cell in header_cells]

    if column_count == 19 and header == ["左侧指标", "评价尺度", "右侧指标"]:
        physical_cells = table.rows[0]._tr.findall(qn("w:tc"))
        spans = []
        for cell in physical_cells:
            properties = cell.find(qn("w:tcPr"))
            span = properties.find(qn("w:gridSpan")) if properties is not None else None
            spans.append(int(span.get(qn("w:val"), "1")) if span is not None else 1)
        if spans != [1, 17, 1] or len(table.rows) < 2:
            return None
        scale = tuple(
            _compact_table_text(cell.text)
            for cell in table.rows[1].cells[1:18]
        )
        return "ahp" if scale == _AHP_SCALE_VALUES else None

    if column_count == 7 and header == [
        "编码",
        "评价指标",
        "优（5）",
        "良（4）",
        "中（3）",
        "较差（2）",
        "差（1）",
    ]:
        return "fce"

    if column_count == 3 and header == ["等级", "分值", "一般含义"]:
        return "rating_meaning"
    return None


def _distribute_twips(total: int, count: int) -> list[int]:
    base, remainder = divmod(total, count)
    return [base + 1] * remainder + [base] * (count - remainder)


def _questionnaire_table_widths(kind: str, total_width: int) -> list[int]:
    if kind == "ahp":
        outer = int(Cm(3.0).twips)
        return [outer, *_distribute_twips(total_width - 2 * outer, 17), outer]
    if kind == "fce":
        code = int(Cm(1.3).twips)
        indicator = int(Cm(7.2).twips)
        return [code, indicator, *_distribute_twips(total_width - code - indicator, 5)]
    if kind == "rating_meaning":
        short = int(Cm(2.0).twips)
        return [short, short, total_width - 2 * short]
    raise ValueError(f"Unsupported questionnaire table kind: {kind}")


def _questionnaire_column_alignment(kind: str, column_index: int) -> str:
    if kind == "ahp":
        return "left" if column_index in {0, 18} else "center"
    if kind == "fce":
        return "left" if column_index == 1 else "center"
    if kind == "rating_meaning":
        return "left" if column_index == 2 else "center"
    return "left"


def _set_questionnaire_cell_margins(cell, *, left: int, right: int) -> None:
    properties = cell._tc.get_or_add_tcPr()
    margins = properties.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        properties.append(margins)
    for edge, width in (("top", 60), ("left", left), ("bottom", 60), ("right", right)):
        node = margins.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            margins.append(node)
        node.set(qn("w:type"), "dxa")
        node.set(qn("w:w"), str(width))


def _balanced_ahp_label_lines(value: str) -> tuple[str, str] | None:
    compact = _compact_table_text(value)
    match = re.fullmatch(r"([A-Z]\d{1,2})(.+)", compact)
    if match is None:
        return None
    code, label = match.groups()
    if len(label) <= 6:
        return None

    midpoint = len(label) / 2
    semantic = {
        index + 1
        for index, character in enumerate(label)
        if character in "与及和、，；"
        and 3 <= index + 1 <= len(label) - 3
    }
    nearby = [index for index in semantic if abs(index - midpoint) <= 2.5]
    if nearby:
        split_at = min(nearby, key=lambda index: (abs(index - midpoint), -index))
    else:
        balanced = [index for index in range(3, len(label) - 2) if index % 2 == 0]
        if not balanced:
            balanced = list(range(3, len(label) - 2))
        split_at = min(balanced, key=lambda index: (abs(index - midpoint), -index))
    return code + label[:split_at], label[split_at:]


def _balance_ahp_label_cell(cell) -> None:
    if not cell.paragraphs:
        return
    lines = _balanced_ahp_label_lines(cell.text)
    if lines is None:
        return
    cell.paragraphs[0].text = lines[0] + "\n" + lines[1]


def _format_tables(
    document: Document,
    table_layouts: Iterable[LatexTableLayout] = (),
    labels: LabelRegistry | None = None,
) -> None:
    page_target_width = int(Cm(15.5).twips)
    layouts = list(table_layouts)
    caption_numbers = _caption_number_by_table(document)
    numbered_layouts: dict[str, LatexTableLayout] = {}
    if labels is not None:
        for layout in layouts:
            target = labels.targets.get(layout.label or "")
            if target is not None and target.kind == "表":
                numbered_layouts[target.number] = layout
    unused_layouts = list(layouts)
    used_layout_ids: set[int] = set()
    questionnaire_layouts: dict[str, LatexTableLayout | None] = {}
    semantic_table_count = 0
    for table in document.tables:
        style = table._tbl.tblPr.find(qn("w:tblStyle"))
        if style is not None and style.get(qn("w:val")) in {
            "YibinCoverLayout",
            "YibinFrontLayout",
            STYLE_REVIEW_INFO_TABLE,
            STYLE_PROPOSAL_FORM_TABLE,
        }:
            continue
        semantic_table_count += 1
        caption_number = caption_numbers.get(table._tbl)
        questionnaire_kind = _questionnaire_table_kind(table)
        if questionnaire_kind is not None:
            if questionnaire_kind not in questionnaire_layouts:
                questionnaire_layouts[questionnaire_kind] = next(
                    (
                        layout
                        for layout in unused_layouts
                        if id(layout) not in used_layout_ids
                        and len(layout.columns) == len(table.columns)
                    ),
                    None,
                )
            source_layout = questionnaire_layouts[questionnaire_kind]
        else:
            source_layout = numbered_layouts.get(caption_number or "")
            if source_layout is None:
                source_layout = next(
                    (layout for layout in unused_layouts if id(layout) not in used_layout_ids),
                    None,
                )
        if source_layout is not None:
            used_layout_ids.add(id(source_layout))
        target_width = page_target_width
        if questionnaire_kind is None and (
            source_layout is not None
            and source_layout.table_width_fraction is not None
            and source_layout.table_width_fraction > 0
        ):
            effective_fraction = source_layout.table_width_fraction
            known_widths = [
                column.width_fraction
                for column in source_layout.columns
                if column.width_fraction is not None
                and column.width_fraction > 0
            ]
            if len(known_widths) == len(source_layout.columns):
                effective_fraction = min(
                    effective_fraction,
                    sum(known_widths),
                )
            target_width = max(
                1,
                round(
                    page_target_width
                    * min(1.0, effective_fraction)
                ),
            )
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = document.styles[STYLE_THREE_LINE_TABLE]
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
        horizontal_margin = (
            20
            if questionnaire_kind is not None
            else max(
                0,
                round(
                    20
                    * (
                        source_layout.tabcolsep_pt
                        if source_layout is not None
                        and source_layout.tabcolsep_pt is not None
                        else 6.0
                    )
                ),
            )
        )
        for edge, width in (
            ("top", 60 if questionnaire_kind is not None else 80),
            ("left", horizontal_margin),
            ("bottom", 60 if questionnaire_kind is not None else 80),
            ("right", horizontal_margin),
        ):
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
        if source_layout is not None and len(source_layout.columns) != column_count:
            print(
                "WARNING: LaTeX 表格列数与 Word 表格不一致，已回退到 Pandoc 列宽/对齐："
                f"LaTeX={len(source_layout.columns)}, Word={column_count}",
                file=sys.stderr,
            )
            source_layout = None
        widths = (
            _questionnaire_table_widths(questionnaire_kind, target_width)
            if questionnaire_kind is not None
            else _resolve_table_column_widths(source_layout, raw_widths, target_width)
        )
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
            if questionnaire_kind == "ahp" and row_index >= 2:
                row.height = Pt(28)
                row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            if row_index == 0:
                header = tr_pr.find(qn("w:tblHeader"))
                if header is None:
                    header = OxmlElement("w:tblHeader")
                    header.set(qn("w:val"), "true")
                    tr_pr.append(header)
            else:
                # Short rows remain atomic; rows containing substantial text
                # are deliberately left splittable to avoid large white gaps.
                text_length = sum(len(p.text) for c in row.cells for p in c.paragraphs)
                if text_length < 160:
                    cant_split = tr_pr.find(qn("w:cantSplit"))
                    if cant_split is None:
                        tr_pr.append(OxmlElement("w:cantSplit"))
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

            seen_cells: set[int] = set()
            for column_index, cell in enumerate(row.cells):
                cell_key = id(cell._tc)
                if cell_key in seen_cells:
                    continue
                seen_cells.add(cell_key)
                source_column = (
                    source_layout.columns[column_index]
                    if source_layout is not None
                    and column_index < len(source_layout.columns)
                    else None
                )
                if questionnaire_kind == "ahp" and row_index >= 2 and column_index in {0, 18}:
                    _balance_ahp_label_cell(cell)
                if questionnaire_kind is not None:
                    margin = 20
                    if (
                        (questionnaire_kind == "ahp" and column_index in {0, 18})
                        or (questionnaire_kind == "fce" and column_index in {0, 1})
                        or questionnaire_kind == "rating_meaning"
                    ):
                        margin = 80
                    _set_questionnaire_cell_margins(cell, left=margin, right=margin)
                    vertical = "center"
                else:
                    vertical = source_column.vertical if source_column is not None else "center"
                cell.vertical_alignment = {
                    "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
                    "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
                    "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
                }.get(vertical, WD_CELL_VERTICAL_ALIGNMENT.CENTER)
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
                    source_alignment = paragraph.alignment
                    tc_pr = cell._tc.get_or_add_tcPr()
                    grid_span = tc_pr.find(qn("w:gridSpan"))
                    is_merged_cell = (
                        grid_span is not None
                        and int(grid_span.get(qn("w:val"), "1")) > 1
                    )
                    if questionnaire_kind is not None:
                        horizontal = _questionnaire_column_alignment(
                            questionnaire_kind,
                            column_index,
                        )
                    elif source_column is not None and not is_merged_cell:
                        # For ordinary cells the LaTeX preamble is the source
                        # of truth.  Pandoc may leave direct paragraph
                        # alignment behind while converting the simplified
                        # l/c/r preamble; do not let that incidental OOXML
                        # override L/C/R/X or p/m/b column semantics.
                        horizontal = source_column.horizontal
                    elif source_alignment == WD_ALIGN_PARAGRAPH.CENTER:
                        horizontal = "center"
                    elif source_alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                        horizontal = "right"
                    elif source_alignment == WD_ALIGN_PARAGRAPH.LEFT:
                        horizontal = "left"
                    else:
                        horizontal = "left"
                    paragraph.style = document.styles[
                        _table_style_for_alignment(row_index, horizontal)
                    ]
                    _clear_direct_paragraph_format(
                        paragraph,
                        "ind",
                        "spacing",
                        "jc",
                    )
                    if row_index == 0:
                        # Keep the repeated header with the first real data row.
                        # Otherwise Word may leave the caption/header at the
                        # bottom of a page when the first body row is tall.
                        paragraph.paragraph_format.keep_with_next = True
                    if (
                        source_column is not None
                        and source_column.first_line_indent_pt is not None
                        and source_column.first_line_indent_pt != 0
                    ):
                        paragraph.paragraph_format.first_line_indent = Pt(
                            source_column.first_line_indent_pt
                        )
                    for run in paragraph.runs:
                        _clear_direct_run_typography(run)

    if layouts and (
        semantic_table_count != len(layouts) or len(used_layout_ids) != len(layouts)
    ):
        print(
            "WARNING: LaTeX 表格布局数量与 Word 语义表格数量不一致，"
            f"LaTeX={len(layouts)}, Word={semantic_table_count}, "
            f"matched={len(used_layout_ids)}；未匹配表格已使用 Pandoc 回退。",
            file=sys.stderr,
        )


def _replace_pandoc_checkbox_math(document: Document) -> int:
    """Replace simple Pandoc math checkboxes with printable native Word runs."""

    replaced = 0
    for math in list(document.element.iter(qn("m:oMath"))):
        children = list(math)
        texts = list(math.iter(qn("m:t")))
        if (
            len(children) != 1
            or children[0].tag != qn("m:r")
            or len(texts) != 1
            or (texts[0].text or "") != "\u25ab"
        ):
            continue

        parent = math.getparent()
        if parent is None:
            continue
        replacement_parent = parent
        replacement_target = math
        if parent.tag == qn("m:oMathPara"):
            paragraph = parent.getparent()
            if (
                len(parent) != 1
                or paragraph is None
                or paragraph.tag != qn("w:p")
            ):
                continue
            replacement_parent = paragraph
            replacement_target = parent
        elif parent.tag != qn("w:p"):
            continue

        run = OxmlElement("w:r")
        properties = OxmlElement("w:rPr")
        fonts = OxmlElement("w:rFonts")
        for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn(f"w:{attribute}"), "SimSun")
        properties.append(fonts)
        for name in ("sz", "szCs"):
            size = OxmlElement(f"w:{name}")
            size.set(qn("w:val"), "24")
            properties.append(size)
        text = OxmlElement("w:t")
        text.text = "\u25a1"
        run.append(properties)
        run.append(text)
        replacement_parent.replace(replacement_target, run)
        replaced += 1
    return replaced


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


def _split_pandoc_captioned_figures(document: Document) -> None:
    """Split Pandoc's drawing, alt-caption and following prose into paragraphs.

    Pandoc can serialize a standalone Markdown figure and the next source
    paragraph into one Word paragraph.  The visible caption then survives only
    as ``wp:docPr/@descr``, so a style-only pass cannot create Caption/SEQ/REF
    semantics.  Reconstruct the three semantic paragraphs before promoting
    caption fields.
    """

    for paragraph in list(document.paragraphs):
        if paragraph.style.name == "CoverLogo" or not paragraph._p.xpath(
            ".//w:drawing|.//w:pict"
        ):
            continue
        descriptions = [
            (node.get("descr") or "").strip()
            for node in paragraph._p.xpath(".//wp:docPr")
        ]
        caption_text = next(
            (
                value
                for value in descriptions
                if re.match(r"^(?:图|表)[0-9]+(?:\.[0-9]+)?(?:\s|$)", value)
            ),
            "",
        )
        if not caption_text:
            continue

        following = paragraph._p.getnext()
        following_text = ""
        if following is not None and following.tag == qn("w:p"):
            following_text = "".join(
                node.text or "" for node in following.iter(qn("w:t"))
            ).strip()

        figure = paragraph.insert_paragraph_before(style=STYLE_FIGURE)
        if following_text != caption_text:
            paragraph.insert_paragraph_before(
                caption_text,
                style=STYLE_FIGURE_CAPTION,
            )
        for child in list(paragraph._p):
            if child.tag == qn("w:pPr"):
                continue
            if any(True for _ in child.iter(qn("w:drawing"))) or any(
                True for _ in child.iter(qn("w:pict"))
            ):
                paragraph._p.remove(child)
                figure._p.append(child)

        # Pandoc's source label bookmark encloses only the image and may use a
        # colon, which Word rewrites on save.  Native full/number bookmarks are
        # generated deterministically around the promoted caption instead.
        for bookmark in list(paragraph._p.xpath("./w:bookmarkStart|./w:bookmarkEnd")):
            paragraph._p.remove(bookmark)
        for run in list(paragraph.runs):
            if (run.text or "").strip():
                break
            run._r.getparent().remove(run._r)

        if not paragraph.text.strip() and not paragraph._p.xpath(
            ".//w:drawing|.//w:pict|.//m:oMath|.//m:oMathPara"
        ):
            paragraph._p.getparent().remove(paragraph._p)
        else:
            paragraph.style = document.styles[STYLE_BODY]


def _format_images_and_captions(
    document: Document,
    labels: LabelRegistry | None = None,
    discipline: str = "science",
    *,
    figure_sequence: str = "图",
    table_sequence: str = "表",
) -> None:
    """Apply Word-style-driven layout to embedded figures and captions."""

    _split_pandoc_captioned_figures(document)
    for paragraph in document.paragraphs:
        has_drawing = bool(paragraph._p.xpath(".//w:drawing|.//w:pict"))
        if has_drawing and paragraph.style.name not in {"CoverLogo"}:
            paragraph.style = document.styles[STYLE_FIGURE]
            _clear_direct_paragraph_format(
                paragraph,
                "keepNext",
                "keepLines",
                "spacing",
                "ind",
                "jc",
            )

        style_name = paragraph.style.name
        text = paragraph.text.strip()
        if style_name not in {
            "Caption",
            "Image Caption",
            "Figure Caption",
            "Table Caption",
            STYLE_FIGURE_CAPTION,
            STYLE_TABLE_CAPTION,
        }:
            continue
        is_table_caption = text.startswith("表")
        # The reusable thesis styles inherit Word's built-in Caption style,
        # while keeping table-vs-figure pagination explicit and editable.
        paragraph.style = document.styles[
            STYLE_TABLE_CAPTION if is_table_caption else STYLE_FIGURE_CAPTION
        ]
        _clear_direct_paragraph_format(
            paragraph,
            "keepNext",
            "keepLines",
            "spacing",
            "ind",
            "jc",
        )
        for run in paragraph.runs:
            _clear_title_run_typography(run)

    _promote_caption_fields(
        document,
        labels,
        discipline,
        figure_sequence=figure_sequence,
        table_sequence=table_sequence,
    )

    for blip in document.element.iter(qn("a:blip")):
        if blip.get(qn("r:link")):
            raise BuildError("检测到 Word 图片外链；模板只允许内嵌图片关系。")


def _word_field_identifier(value: str) -> str:
    value = value.strip()
    if not value:
        raise BuildError("Word 题注序列名不能为空。")
    if re.fullmatch(r'[^\s"\\]+', value):
        return value
    return '"' + value.replace('"', '""') + '"'


def _field_run(
    instruction: str,
    result: str = "",
    *,
    locked: bool = False,
):
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    field.set(qn("w:dirty"), "true")
    if locked:
        field.set(qn("w:fldLock"), "true")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = result
    run.append(text)
    field.append(run)
    return field


def _promote_caption_fields(
    document: Document,
    labels: LabelRegistry | None = None,
    discipline: str = "science",
    *,
    figure_sequence: str = "图",
    table_sequence: str = "表",
) -> None:
    """Turn Pandoc caption text into native Word caption sequences.

    The visible Chinese label and the registered Word CaptionLabel identifier
    intentionally match.  This keeps generated captions, Word's Insert Caption
    dialog and the Cross-reference dialog on the same ``图``/``表`` sequence.
    """

    bookmark_id = _next_bookmark_id(document)
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        match = re.match(r"^(图|表)([0-9]+(?:\.[0-9]+)?)\s*(.*)$", text)
        if not match or paragraph.style.name not in {
            "Figure Caption",
            "Table Caption",
            "Caption",
            "Image Caption",
            STYLE_FIGURE_CAPTION,
            STYLE_TABLE_CAPTION,
        }:
            continue
        kind, number, rest = match.groups()
        # Rebuild paragraph while retaining its caption style and formatting.
        for child in list(paragraph._p):
            if child.tag != qn("w:pPr"):
                paragraph._p.remove(child)
        chapter_number = ""
        sequence_number = number
        sequence_name = figure_sequence if kind == "图" else table_sequence
        if discipline == "science" and "." in number:
            chapter, sequence_number = number.split(".", 1)
            chapter_number = chapter

        matched = next(
            (
                label
                for label, target in (labels.targets.items() if labels else [])
                if target.kind == kind and target.number == number
            ),
            None,
        )
        seed = matched or f"anonymous:{kind}:{number}:{bookmark_id}"
        full_name = (
            _label_bookmark_name(matched)
            if matched
            else _hidden_ref_bookmark_name(f"full:{seed}")
        )
        number_name = (
            _label_number_bookmark_name(matched)
            if matched
            else _hidden_ref_bookmark_name(f"number:{seed}", prefix="_RefNum")
        )
        full_id = bookmark_id
        number_id = bookmark_id + 1
        bookmark_id += 2

        full_start = OxmlElement("w:bookmarkStart")
        full_start.set(qn("w:id"), str(full_id))
        full_start.set(qn("w:name"), full_name)
        full_end = OxmlElement("w:bookmarkEnd")
        full_end.set(qn("w:id"), str(full_id))
        number_start = OxmlElement("w:bookmarkStart")
        number_start.set(qn("w:id"), str(number_id))
        number_start.set(qn("w:name"), number_name)
        number_end = OxmlElement("w:bookmarkEnd")
        number_end.set(qn("w:id"), str(number_id))

        paragraph._p.append(full_start)
        paragraph._p.append(_plain_run(kind))
        paragraph._p.append(number_start)
        if chapter_number:
            for node in _complex_field_runs(
                f' STYLEREF "{STYLE_HEADING_1}" \\n ',
                chapter_number,
            ):
                paragraph._p.append(node)
            paragraph._p.append(_plain_run("."))
        sequence_instruction = f" SEQ {_word_field_identifier(sequence_name)} "
        if sequence_number == "1":
            sequence_instruction += "\\r 1 "
        sequence_instruction += "\\* ARABIC "
        for node in _complex_field_runs(sequence_instruction, sequence_number):
            paragraph._p.append(node)
        paragraph._p.append(number_end)
        paragraph._p.append(full_end)
        if rest:
            paragraph._p.append(_plain_run(" " + rest))


def _label_bookmark_name(label: str) -> str:
    return _hidden_ref_bookmark_name(f"full:{label}")


def _label_number_bookmark_name(label: str) -> str:
    return _hidden_ref_bookmark_name(f"number:{label}", prefix="_RefNum")


def _hidden_ref_bookmark_name(seed: str, *, prefix: str = "_Ref") -> str:
    """Return a deterministic hidden bookmark that resembles Word's own.

    Leading underscores keep the target out of Word's normal bookmark list.
    A decimal digest is used because native cross-reference bookmarks are
    conventionally named ``_Ref#########``.
    """

    digest = int(hashlib.sha1(seed.encode("utf-8")).hexdigest()[:15], 16)
    return f"{prefix}{digest % 10**15:015d}"


def _next_bookmark_id(document: Document) -> int:
    values = [
        int(value)
        for node in document.element.iter(qn("w:bookmarkStart"))
        if (value := node.get(qn("w:id"))) is not None and value.isdigit()
    ]
    return max(values, default=0) + 1


def _citation_bookmark_name(key: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_]", "_", key or "")[:17]
    digest = hashlib.sha1((key or "").encode("utf-8")).hexdigest()[:8]
    return f"YibinCitation_{safe}_{digest}"


def _run_style_properties(character_style_id: str | None):
    if not character_style_id:
        return None
    r_pr = OxmlElement("w:rPr")
    r_style = OxmlElement("w:rStyle")
    r_style.set(qn("w:val"), character_style_id)
    r_pr.append(r_style)
    return r_pr


def _plain_run(text: str, *, character_style_id: str | None = None):
    run = OxmlElement("w:r")
    r_pr = _run_style_properties(character_style_id)
    if r_pr is not None:
        run.append(r_pr)
    node = OxmlElement("w:t")
    if text.startswith(" ") or text.endswith(" "):
        node.set(qn("xml:space"), "preserve")
    node.text = text
    run.append(node)
    return run


def _complex_field_runs(
    instruction: str,
    result: str,
    *,
    character_style_id: str | None = None,
    locked: bool = False,
) -> list:
    """Return a schema-valid complex field whose result keeps its style.

    Word recreates an unlocked REF result during F9.  CHARFORMAT reapplies the
    character style carried by the field-code runs, so the citation number
    remains a 9 pt superscript after refresh instead of dropping to baseline.
    """

    def styled_run(child):
        run = OxmlElement("w:r")
        r_pr = _run_style_properties(character_style_id)
        if r_pr is not None:
            run.append(r_pr)
        run.append(child)
        return run

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    if locked:
        begin.set(qn("w:fldLock"), "true")
    code = OxmlElement("w:instrText")
    code.set(qn("xml:space"), "preserve")
    code.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = result
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    return [
        styled_run(begin),
        styled_run(code),
        styled_run(separate),
        styled_run(text),
        styled_run(end),
    ]


def _replace_reference_markers(document: Document, labels: LabelRegistry, citations: CitationRegistry, mode: str) -> None:
    """Replace opaque LaTeX reference/citation markers with Word fields."""
    label_names = {label: _label_bookmark_name(label) for label in labels.targets}
    label_number_names = {
        label: _label_number_bookmark_name(label) for label in labels.targets
    }
    citation_names = {key: _citation_bookmark_name(key) for key in citations.order}
    citation_numbers = {
        key: index for index, key in enumerate(citations.order, start=1)
    }
    generated_names = [
        *label_names.values(),
        *label_number_names.values(),
        *citation_names.values(),
    ]
    if any(len(name) > 40 for name in generated_names):
        raise BuildError("Word 书签名超过 40 字符限制。")
    folded_names = [name.casefold() for name in generated_names]
    if len(folded_names) != len(set(folded_names)):
        raise BuildError("Word 书签名发生大小写不敏感碰撞。")
    citation_style_id = document.styles[STYLE_CITATION].style_id
    for paragraph in document.paragraphs:
        for run in list(paragraph.runs):
            value = run.text or ""
            markers = list(re.finditer(r"YIBINXREF\d{8}|YIBINCITE\d{8}", value))
            if not markers:
                continue
            parent = run._r.getparent(); index = parent.index(run._r)
            parent.remove(run._r)
            cursor = 0
            for match in markers:
                token = match.group(0)
                if token.startswith("YIBINXREF"):
                    label, paren = labels.references.get(token, ("", False)); target = labels.targets.get(label)
                    if target is None: raise BuildError(f"未找到交叉引用标签：{label}")
                    prefix = value[cursor:match.start()]
                    use_full_target = bool(paren and target.kind == "公式")
                    if not paren and target.kind in {"图", "表"} and prefix.endswith(target.kind):
                        prefix = prefix[: -len(target.kind)]
                        use_full_target = True
                    if prefix:
                        parent.insert(index, _plain_run(prefix)); index += 1
                    bookmark = (
                        label_names[label]
                        if use_full_target
                        else label_number_names[label]
                    )
                    if use_full_target and target.kind in {"图", "表"}:
                        cached_result = f"{target.kind}{target.number}"
                    elif use_full_target and target.kind == "公式":
                        cached_result = f"({target.number})"
                    else:
                        cached_result = (
                            f"({target.number})" if paren else target.number
                        )
                    for field_run in _complex_field_runs(
                        f" REF {bookmark} \\h ",
                        cached_result,
                    ):
                        parent.insert(index, field_run)
                        index += 1
                else:
                    prefix = value[cursor:match.start()]
                    if prefix:
                        parent.insert(index, _plain_run(prefix)); index += 1
                    keys = sorted(
                        citations.clusters.get(token, []),
                        key=lambda key: citation_numbers[key],
                    )
                    open_run = _plain_run("[", character_style_id=citation_style_id)
                    parent.insert(index, open_run); index += 1
                    for key_index, key in enumerate(keys):
                        if key_index:
                            sep = copy.deepcopy(open_run); sep.find(qn("w:t")).text = ","; parent.insert(index, sep); index += 1
                        number = citation_numbers[key]
                        instruction = (
                            f" CITATION {key} \\l 2052 \\* CHARFORMAT "
                            if mode == "native"
                            else f" REF {citation_names[key]} \\h \\* CHARFORMAT "
                        )
                        for field_run in _complex_field_runs(
                            instruction,
                            str(number),
                            character_style_id=citation_style_id,
                            locked=mode == "native",
                        ):
                            parent.insert(index, field_run)
                            index += 1
                    close_run = copy.deepcopy(open_run); close_run.find(qn("w:t")).text = "]"; parent.insert(index, close_run); index += 1
                cursor = match.end()
            if cursor < len(value):
                parent.insert(index, _plain_run(value[cursor:]))
    # Bookmark bibliography entries by citation order where possible.
    for paragraph in document.paragraphs:
        if paragraph.style.name.casefold() not in {
            "bibliography",
            STYLE_BIBLIOGRAPHY.casefold(),
        }:
            continue
        m = re.match(r"^\s*\[?(\d+)\]?\s*", paragraph.text or "")
        if not m: continue
        number = int(m.group(1))
        if 1 <= number <= len(citations.order):
            key = citations.order[number - 1]; bid = 3000 + number
            # Isolate the displayed number so REF returns only the number,
            # never the complete bibliography entry.
            first = paragraph.runs[0] if paragraph.runs else None
            number_run = None
            first_match = (
                re.match(r"^(\s*)\[?\d+\]?", first.text)
                if first is not None
                else None
            )
            if first is not None and first_match is not None:
                leading = first_match.group(1)
                remainder = first.text[first_match.end():]
                parent = first._r.getparent()
                position = parent.index(first._r)
                r_pr = first._r.find(qn("w:rPr"))
                parent.remove(first._r)
                pieces = (leading + "[", str(number), "]" + remainder)
                created = []
                for piece in pieces:
                    node = OxmlElement("w:r")
                    if r_pr is not None:
                        node.append(copy.deepcopy(r_pr))
                    text_node = OxmlElement("w:t")
                    if piece.startswith(" ") or piece.endswith(" "):
                        text_node.set(qn("xml:space"), "preserve")
                    text_node.text = piece
                    node.append(text_node)
                    parent.insert(position, node)
                    position += 1
                    created.append(node)
                number_run = created[1]
            start = OxmlElement("w:bookmarkStart"); start.set(qn("w:id"), str(bid)); start.set(qn("w:name"), citation_names[key])
            end = OxmlElement("w:bookmarkEnd"); end.set(qn("w:id"), str(bid))
            if number_run is not None:
                pos = paragraph._p.index(number_run)
                paragraph._p.insert(pos, start)
                paragraph._p.insert(pos + 2, end)
            else:
                raise BuildError(f"无法隔离参考文献编号书签：{paragraph.text[:80]}")


def _word_source_type(csl_type: str) -> str:
    return {
        "article-journal": "JournalArticle",
        "article-magazine": "ArticleInAPeriodical",
        "article-newspaper": "ArticleInAPeriodical",
        "book": "Book",
        "chapter": "BookSection",
        "paper-conference": "ConferenceProceedings",
        "report": "Report",
        "thesis": "Report",
        "webpage": "DocumentFromInternetSite",
        "post-weblog": "InternetSite",
        "patent": "Patent",
    }.get(csl_type, "Misc")


def _append_word_source_text(source: ET.Element, name: str, value: object) -> None:
    text = str(value or "").strip()
    if not text:
        return
    ET.SubElement(source, f"{{{BIBLIOGRAPHY_NS}}}{name}").text = text


BIBLIOGRAPHY_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
CUSTOM_XML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _bibliography_sources_xml(
    citations: CitationRegistry,
    bibliography_items: list[dict[str, object]],
) -> bytes:
    ET.register_namespace("b", BIBLIOGRAPHY_NS)
    by_key = {str(item.get("id", "")): item for item in bibliography_items}
    root = ET.Element(
        f"{{{BIBLIOGRAPHY_NS}}}Sources",
        {
            "SelectedStyle": "\\Yibin-GB-T-7714-Numeric.xsl",
            "StyleName": "GB/T 7714 数字引用（YibinThesis）",
            "Version": "6",
        },
    )
    for number, key in enumerate(citations.order, start=1):
        item = by_key.get(key, {})
        source = ET.SubElement(root, f"{{{BIBLIOGRAPHY_NS}}}Source")
        _append_word_source_text(source, "Tag", key)
        _append_word_source_text(source, "SourceType", _word_source_type(str(item.get("type", ""))))
        _append_word_source_text(
            source,
            "Guid",
            "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"yibinthesis:{key}")).upper() + "}",
        )
        _append_word_source_text(source, "LCID", "2052")
        _append_word_source_text(source, "RefOrder", number)

        authors = item.get("author")
        if isinstance(authors, list) and authors:
            author_node = ET.SubElement(source, f"{{{BIBLIOGRAPHY_NS}}}Author")
            author_role = ET.SubElement(author_node, f"{{{BIBLIOGRAPHY_NS}}}Author")
            people = [author for author in authors if isinstance(author, dict) and not author.get("literal")]
            corporate = [str(author.get("literal", "")).strip() for author in authors if isinstance(author, dict) and author.get("literal")]
            if people:
                name_list = ET.SubElement(author_role, f"{{{BIBLIOGRAPHY_NS}}}NameList")
                for author in people:
                    person = ET.SubElement(name_list, f"{{{BIBLIOGRAPHY_NS}}}Person")
                    _append_word_source_text(person, "Last", author.get("family"))
                    _append_word_source_text(person, "First", author.get("given"))
            elif corporate:
                _append_word_source_text(author_role, "Corporate", "；".join(corporate))

        title = item.get("title") or key
        _append_word_source_text(source, "Title", title)
        container = item.get("container-title")
        if isinstance(container, list):
            container = "; ".join(str(value) for value in container)
        _append_word_source_text(source, "JournalName", container)
        issued = item.get("issued")
        if isinstance(issued, dict):
            date_parts = issued.get("date-parts")
            if isinstance(date_parts, list) and date_parts and isinstance(date_parts[0], list):
                parts = date_parts[0]
                if parts:
                    _append_word_source_text(source, "Year", parts[0])
                if len(parts) > 1:
                    _append_word_source_text(source, "Month", parts[1])
                if len(parts) > 2:
                    _append_word_source_text(source, "Day", parts[2])
        for csl_name, word_name in (
            ("publisher", "Publisher"),
            ("publisher-place", "City"),
            ("volume", "Volume"),
            ("issue", "Issue"),
            ("page", "Pages"),
            ("DOI", "DOI"),
            ("URL", "URL"),
            ("number", "StandardNumber"),
        ):
            _append_word_source_text(source, word_name, item.get(csl_name))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _inject_bibliography_sources(
    docx_path: Path,
    citations: CitationRegistry,
    bibliography_items: list[dict[str, object]],
) -> None:
    """Add a fully related Word bibliography custom XML data store."""
    if not citations.order:
        return
    temp = docx_path.with_suffix(".sources.docx")
    with zipfile.ZipFile(docx_path, "r") as src:
        names = set(src.namelist())
        index = 1
        while f"customXml/item{index}.xml" in names:
            index += 1
        item_name = f"customXml/item{index}.xml"
        props_name = f"customXml/itemProps{index}.xml"
        item_rels_name = f"customXml/_rels/item{index}.xml.rels"

        document_rels = ET.fromstring(src.read("word/_rels/document.xml.rels"))
        relationship_ids = {
            relationship.get("Id", "") for relationship in document_rels
        }
        rel_index = 1
        while f"rIdYibinBibliography{rel_index}" in relationship_ids:
            rel_index += 1
        ET.SubElement(
            document_rels,
            f"{{{PACKAGE_REL_NS}}}Relationship",
            {
                "Id": f"rIdYibinBibliography{rel_index}",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
                "Target": f"../{item_name}",
            },
        )

        content_types = ET.fromstring(src.read("[Content_Types].xml"))
        ET.SubElement(
            content_types,
            f"{{{CONTENT_TYPES_NS}}}Override",
            {
                "PartName": "/" + props_name,
                "ContentType": "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
            },
        )

        store_id = "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, "yibinthesis:bibliography-store")).upper() + "}"
        ET.register_namespace("ds", CUSTOM_XML_NS)
        props = ET.Element(f"{{{CUSTOM_XML_NS}}}datastoreItem", {f"{{{CUSTOM_XML_NS}}}itemID": store_id})
        schema_refs = ET.SubElement(props, f"{{{CUSTOM_XML_NS}}}schemaRefs")
        ET.SubElement(schema_refs, f"{{{CUSTOM_XML_NS}}}schemaRef", {f"{{{CUSTOM_XML_NS}}}uri": BIBLIOGRAPHY_NS})

        ET.register_namespace("", PACKAGE_REL_NS)
        item_rels = ET.Element(f"{{{PACKAGE_REL_NS}}}Relationships")
        ET.SubElement(
            item_rels,
            f"{{{PACKAGE_REL_NS}}}Relationship",
            {
                "Id": "rId1",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
                "Target": f"itemProps{index}.xml",
            },
        )

        replacements = {
            "word/_rels/document.xml.rels": ET.tostring(document_rels, encoding="utf-8", xml_declaration=True),
            "[Content_Types].xml": ET.tostring(content_types, encoding="utf-8", xml_declaration=True),
        }
        with zipfile.ZipFile(temp, "w") as dst:
            for member in src.infolist():
                dst.writestr(member, replacements.get(member.filename, src.read(member)))
            dst.writestr(item_name, _bibliography_sources_xml(citations, bibliography_items))
            dst.writestr(props_name, ET.tostring(props, encoding="utf-8", xml_declaration=True))
            dst.writestr(item_rels_name, ET.tostring(item_rels, encoding="utf-8", xml_declaration=True))
    os.replace(temp, docx_path)


