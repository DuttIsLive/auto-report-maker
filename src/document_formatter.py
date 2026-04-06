"""
document_formatter.py
~~~~~~~~~~~~~~~~~~~~~
Builds a professionally formatted .docx report from structured content.

All text uses Times New Roman as required.  Visual differentiation is
achieved exclusively through size, weight, colour, and spacing.
"""

from __future__ import annotations

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from src.color_schemes import ColorScheme, pick_scheme

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

FONT_NAME = "Times New Roman"
BODY_SIZE = Pt(12)
H1_SIZE = Pt(18)
H2_SIZE = Pt(15)
H3_SIZE = Pt(13)
CAPTION_SIZE = Pt(10)
SMALL_SIZE = Pt(10)
FOOTER_SIZE = Pt(10)


# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------

def _hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_cell_shading(cell, hex_color: str) -> None:
    """Apply a solid background fill to a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#"))
    tcPr.append(shd)


def _set_cell_border(cell, *, left_color: str = None, width_pt: int = 6) -> None:
    """Apply a thick coloured left border to a table cell."""
    if not left_color:
        return
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(width_pt * 8))  # eighths of a point
    left.set(qn("w:space"), "0")
    left.set(qn("w:color"), left_color.lstrip("#"))
    tcBorders.append(left)
    tcPr.append(tcBorders)


def _set_table_border(table, hex_color: str) -> None:
    """Apply uniform outside/inside borders to a table."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    color = hex_color.lstrip("#")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tblBorders.append(el)
    tblPr.append(tblBorders)


def _add_page_break(doc: Document) -> None:
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)


def _paragraph_page_break(doc: Document) -> None:
    """Insert a page break as a paragraph-level run."""
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)


def _add_horizontal_rule(doc: Document, hex_color: str, thickness_pt: int = 2) -> None:
    """Add a thin coloured paragraph border acting as a section divider."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(thickness_pt * 8))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), hex_color.lstrip("#"))
    pBdr.append(bottom)
    pPr.append(pBdr)


def _apply_run_font(run, size: Pt = BODY_SIZE, bold: bool = False,
                    italic: bool = False, color: str = None) -> None:
    run.font.name = FONT_NAME
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = _hex_to_rgb(color)


def _set_paragraph_spacing(para, space_before: Pt = Pt(0),
                            space_after: Pt = Pt(6),
                            line_spacing: float = 1.15) -> None:
    fmt = para.paragraph_format
    fmt.space_before = space_before
    fmt.space_after = space_after
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    fmt.line_spacing = line_spacing


# ---------------------------------------------------------------------------
# Style setup
# ---------------------------------------------------------------------------

def _setup_document_styles(doc: Document, scheme: ColorScheme) -> None:
    """Override built-in Word styles to use Times New Roman + scheme colours."""
    styles_map = {
        "Normal": (BODY_SIZE, False, False, scheme.neutral_dark),
        "Heading 1": (H1_SIZE, True, False, scheme.primary),
        "Heading 2": (H2_SIZE, True, False, scheme.secondary),
        "Heading 3": (H3_SIZE, True, True, scheme.secondary),
        "Caption": (CAPTION_SIZE, False, True, scheme.neutral_dark),
        "Footer": (FOOTER_SIZE, False, False, scheme.neutral_dark),
        "Header": (FOOTER_SIZE, False, False, scheme.neutral_dark),
    }
    for style_name, (size, bold, italic, color) in styles_map.items():
        try:
            style = doc.styles[style_name]
        except KeyError:
            continue
        style.font.name = FONT_NAME
        style.font.size = size
        style.font.bold = bold
        style.font.italic = italic
        style.font.color.rgb = _hex_to_rgb(color)
        if style_name == "Normal":
            style.paragraph_format.space_after = Pt(6)
            style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            style.paragraph_format.line_spacing = 1.15
        if style_name.startswith("Heading"):
            style.paragraph_format.space_before = Pt(12)
            style.paragraph_format.space_after = Pt(6)
            style.paragraph_format.keep_with_next = True


def _setup_page_layout(doc: Document) -> None:
    """Set standard academic margins (1 inch all round)."""
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.25)
    section.right_margin = Inches(1.25)


# ---------------------------------------------------------------------------
# Cover page
# ---------------------------------------------------------------------------

def _build_cover_page(doc: Document, content: dict, scheme: ColorScheme) -> None:
    """Construct a designed cover page as the first section."""

    # --- top colour band (full-width shaded table) --------------------------
    band_table = doc.add_table(rows=1, cols=1)
    band_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    band_cell = band_table.cell(0, 0)
    _set_cell_shading(band_cell, scheme.primary)
    band_cell.width = Inches(6.5)

    band_para = band_cell.paragraphs[0]
    band_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    band_para.paragraph_format.space_before = Pt(36)
    band_para.paragraph_format.space_after = Pt(8)

    title_run = band_para.add_run(content.get("title", "Report"))
    _apply_run_font(title_run, size=Pt(28), bold=True, color=scheme.cover_text)

    # Subtitle inside the same band
    if content.get("subtitle"):
        sub_para = band_cell.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_para.paragraph_format.space_before = Pt(6)
        sub_para.paragraph_format.space_after = Pt(36)
        sub_run = sub_para.add_run(content["subtitle"])
        _apply_run_font(sub_run, size=Pt(16), italic=True, color=scheme.cover_text)
    else:
        # extra bottom padding
        pad_para = band_cell.add_paragraph()
        pad_para.paragraph_format.space_before = Pt(0)
        pad_para.paragraph_format.space_after = Pt(24)

    # --- decorative rule ----------------------------------------------------
    _add_horizontal_rule(doc, scheme.accent, thickness_pt=3)

    # --- meta block (author / institution / date) ---------------------------
    for label, key in [("Prepared by", "author"), ("Institution", "institution"),
                       ("Date", "date")]:
        value = content.get(key, "")
        if value:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_paragraph_spacing(p, space_before=Pt(6), space_after=Pt(4))
            label_run = p.add_run(f"{label}:  ")
            _apply_run_font(label_run, bold=True, color=scheme.primary)
            val_run = p.add_run(value)
            _apply_run_font(val_run)

    # --- spacer then page break ---------------------------------------------
    spacer = doc.add_paragraph()
    _set_paragraph_spacing(spacer, space_before=Pt(24), space_after=Pt(0))
    _paragraph_page_break(doc)


# ---------------------------------------------------------------------------
# Abstract / Executive Summary
# ---------------------------------------------------------------------------

def _build_abstract(doc: Document, text: str, scheme: ColorScheme) -> None:
    h = doc.add_heading("Abstract", level=1)
    _set_paragraph_spacing(h, space_before=Pt(0), space_after=Pt(6))

    # Wrap abstract in a shaded box
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    _set_cell_shading(cell, scheme.neutral_light)
    _set_cell_border(cell, left_color=scheme.secondary, width_pt=5)
    cell.width = Inches(6.0)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_paragraph_spacing(p, space_before=Pt(6), space_after=Pt(6), line_spacing=1.15)
    run = p.add_run(text)
    _apply_run_font(run, italic=True)

    doc.add_paragraph()  # breathing room


# ---------------------------------------------------------------------------
# Table of Contents (manual field-based TOC)
# ---------------------------------------------------------------------------

def _build_toc(doc: Document, scheme: ColorScheme) -> None:
    """Insert a TOC heading and a Word TOC field."""
    h = doc.add_heading("Table of Contents", level=1)
    _set_paragraph_spacing(h, space_before=Pt(0), space_after=Pt(6))

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    fldChar_begin = OxmlElement("w:fldChar")
    fldChar_begin.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = ' TOC \\o "1-3" \\h \\z \\u '
    fldChar_separate = OxmlElement("w:fldChar")
    fldChar_separate.set(qn("w:fldCharType"), "separate")
    fldChar_end = OxmlElement("w:fldChar")
    fldChar_end.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_separate)
    run._r.append(fldChar_end)

    note_p = doc.add_paragraph()
    note_run = note_p.add_run(
        "(Right-click this field in Word and choose 'Update Field' to populate the Table of Contents.)"
    )
    _apply_run_font(note_run, size=Pt(9), italic=True, color="888888")
    _set_paragraph_spacing(note_p, space_before=Pt(4), space_after=Pt(12))

    _paragraph_page_break(doc)


# ---------------------------------------------------------------------------
# Headers & footers
# ---------------------------------------------------------------------------

def _build_header_footer(doc: Document, title: str, scheme: ColorScheme) -> None:
    section = doc.sections[0]

    # Header
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run(title[:60])
    _apply_run_font(hr, size=FOOTER_SIZE, color=scheme.primary)
    _add_horizontal_rule_to_para(hp, scheme.accent)

    # Footer – left: institution, right: page number
    footer = section.footer
    footer.is_linked_to_previous = False
    ft = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()

    # Add top rule
    _add_top_rule_to_para(ft, scheme.neutral_dark)

    ft.clear()
    ft.alignment = WD_ALIGN_PARAGRAPH.CENTER

    page_run = ft.add_run()
    _apply_run_font(page_run, size=FOOTER_SIZE, color=scheme.neutral_dark)
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")
    page_run._r.append(fldChar1)
    page_run._r.append(instrText)
    page_run._r.append(fldChar2)
    page_run._r.append(fldChar3)


def _add_horizontal_rule_to_para(para, hex_color: str) -> None:
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), hex_color.lstrip("#"))
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_top_rule_to_para(para, hex_color: str) -> None:
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "8")
    top.set(qn("w:space"), "1")
    top.set(qn("w:color"), hex_color.lstrip("#"))
    pBdr.append(top)
    pPr.append(pBdr)


# ---------------------------------------------------------------------------
# Content block renderers
# ---------------------------------------------------------------------------

def _render_paragraph(doc: Document, text: str, scheme: ColorScheme) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_paragraph_spacing(p, space_before=Pt(0), space_after=Pt(6), line_spacing=1.15)
    run = p.add_run(text)
    _apply_run_font(run, color=scheme.neutral_dark)


def _render_bullet_list(doc: Document, items: list[str], scheme: ColorScheme) -> None:
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        _set_paragraph_spacing(p, space_before=Pt(2), space_after=Pt(2), line_spacing=1.15)
        run = p.add_run(item)
        _apply_run_font(run, color=scheme.neutral_dark)


def _render_numbered_list(doc: Document, items: list[str], scheme: ColorScheme) -> None:
    for item in items:
        p = doc.add_paragraph(style="List Number")
        _set_paragraph_spacing(p, space_before=Pt(2), space_after=Pt(2), line_spacing=1.15)
        run = p.add_run(item)
        _apply_run_font(run, color=scheme.neutral_dark)


def _render_highlight_box(doc: Document, block: dict, scheme: ColorScheme) -> None:
    """Render a shaded highlight/callout box with an optional title."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    _set_cell_shading(cell, scheme.neutral_light)
    _set_cell_border(cell, left_color=scheme.accent, width_pt=6)

    if block.get("title"):
        tp = cell.paragraphs[0]
        _set_paragraph_spacing(tp, space_before=Pt(6), space_after=Pt(2))
        title_run = tp.add_run(f"▶  {block['title']}")
        _apply_run_font(title_run, bold=True, color=scheme.primary)
        body_p = cell.add_paragraph()
    else:
        body_p = cell.paragraphs[0]

    body_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_paragraph_spacing(body_p, space_before=Pt(2), space_after=Pt(6), line_spacing=1.15)
    body_run = body_p.add_run(block.get("text", ""))
    _apply_run_font(body_run, color=scheme.neutral_dark)

    doc.add_paragraph()  # spacing after box


def _render_quote_block(doc: Document, block: dict, scheme: ColorScheme) -> None:
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    _set_cell_shading(cell, "F9F9F9")
    _set_cell_border(cell, left_color=scheme.secondary, width_pt=8)

    qp = cell.paragraphs[0]
    qp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_paragraph_spacing(qp, space_before=Pt(6), space_after=Pt(4), line_spacing=1.2)
    qr = qp.add_run(f'"{block.get("text", "")}"')
    _apply_run_font(qr, italic=True, size=Pt(12), color=scheme.neutral_dark)

    if block.get("source"):
        sp = cell.add_paragraph()
        _set_paragraph_spacing(sp, space_before=Pt(2), space_after=Pt(6))
        sp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sr = sp.add_run(f"— {block['source']}")
        _apply_run_font(sr, size=Pt(10), bold=True, color=scheme.secondary)

    doc.add_paragraph()


def _render_table(doc: Document, block: dict, scheme: ColorScheme) -> None:
    """Render a data table with a coloured header row and alternating row shading."""
    headers = block.get("headers", [])
    rows = block.get("rows", [])
    if not headers and not rows:
        return

    num_cols = len(headers) if headers else (len(rows[0]) if rows else 1)
    tbl = doc.add_table(rows=1 + len(rows), cols=num_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_border(tbl, scheme.secondary)

    # Header row
    hdr_row = tbl.rows[0]
    for j, hdr_text in enumerate(headers):
        cell = hdr_row.cells[j]
        _set_cell_shading(cell, scheme.primary)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_paragraph_spacing(p, space_before=Pt(4), space_after=Pt(4))
        run = p.add_run(str(hdr_text))
        _apply_run_font(run, bold=True, color="FFFFFF")

    # Data rows
    for i, row_data in enumerate(rows):
        row_obj = tbl.rows[i + 1]
        bg = scheme.neutral_light if i % 2 == 0 else "FFFFFF"
        for j, cell_text in enumerate(row_data):
            cell = row_obj.cells[j]
            _set_cell_shading(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _set_paragraph_spacing(p, space_before=Pt(3), space_after=Pt(3))
            run = p.add_run(str(cell_text))
            _apply_run_font(run, color=scheme.neutral_dark)

    caption_p = doc.add_paragraph()
    _set_paragraph_spacing(caption_p, space_before=Pt(4), space_after=Pt(10))
    if block.get("caption"):
        cr = caption_p.add_run(f"Table: {block['caption']}")
        _apply_run_font(cr, size=CAPTION_SIZE, italic=True, color=scheme.neutral_dark)


def _render_figure_placeholder(doc: Document, block: dict, scheme: ColorScheme) -> None:
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    _set_cell_shading(cell, "EEEEEE")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_paragraph_spacing(p, space_before=Pt(30), space_after=Pt(30))
    run = p.add_run(f"[ Figure: {block.get('caption', 'Insert figure here')} ]")
    _apply_run_font(run, size=Pt(11), italic=True, color="888888")

    caption_p = doc.add_paragraph()
    _set_paragraph_spacing(caption_p, space_before=Pt(4), space_after=Pt(10))
    cr = caption_p.add_run(f"Figure: {block.get('caption', '')}")
    _apply_run_font(cr, size=CAPTION_SIZE, italic=True, color=scheme.neutral_dark)


def _render_section(doc: Document, section: dict, scheme: ColorScheme) -> None:
    """Render a section heading and all of its content blocks."""
    level = section.get("level", 1)
    heading_text = section.get("heading", "")

    h = doc.add_heading(heading_text, level=level)
    if level == 1:
        _set_paragraph_spacing(h, space_before=Pt(18), space_after=Pt(8))
        _add_horizontal_rule(doc, scheme.accent, thickness_pt=1)
    elif level == 2:
        _set_paragraph_spacing(h, space_before=Pt(12), space_after=Pt(6))
    else:
        _set_paragraph_spacing(h, space_before=Pt(8), space_after=Pt(4))

    for block in section.get("blocks", []):
        _render_block(doc, block, scheme)


def _render_block(doc: Document, block: dict, scheme: ColorScheme) -> None:
    """Dispatch a single content block to the appropriate renderer."""
    btype = block.get("type", "paragraph")
    if btype == "paragraph":
        _render_paragraph(doc, block.get("text", ""), scheme)
    elif btype == "bullet_list":
        _render_bullet_list(doc, block.get("items", []), scheme)
    elif btype == "numbered_list":
        _render_numbered_list(doc, block.get("items", []), scheme)
    elif btype == "highlight_box":
        _render_highlight_box(doc, block, scheme)
    elif btype == "quote_block":
        _render_quote_block(doc, block, scheme)
    elif btype == "table":
        _render_table(doc, block, scheme)
    elif btype == "figure":
        _render_figure_placeholder(doc, block, scheme)


def _render_references(doc: Document, references: list[str], scheme: ColorScheme) -> None:
    h = doc.add_heading("References", level=1)
    _set_paragraph_spacing(h, space_before=Pt(18), space_after=Pt(8))
    _add_horizontal_rule(doc, scheme.accent, thickness_pt=1)
    for ref in references:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_paragraph_spacing(p, space_before=Pt(2), space_after=Pt(4), line_spacing=1.15)
        fmt = p.paragraph_format
        fmt.left_indent = Inches(0.5)
        fmt.first_line_indent = Inches(-0.5)
        run = p.add_run(ref)
        _apply_run_font(run, size=Pt(11), color=scheme.neutral_dark)


def _render_appendix(doc: Document, text: str, scheme: ColorScheme) -> None:
    _paragraph_page_break(doc)
    h = doc.add_heading("Appendix", level=1)
    _set_paragraph_spacing(h, space_before=Pt(0), space_after=Pt(8))
    _add_horizontal_rule(doc, scheme.accent, thickness_pt=1)
    _render_paragraph(doc, text, scheme)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_document(content: dict, output_path: str) -> str:
    """
    Build a fully-formatted .docx report from *content* and save it to
    *output_path*.

    Parameters
    ----------
    content : dict
        Structured report content as produced by
        :func:`src.content_generator.generate_content`.
    output_path : str
        Destination file path (must end in ``.docx``).

    Returns
    -------
    str
        The resolved output path.
    """
    topic = content.get("title", "report")
    scheme = pick_scheme(topic)

    doc = Document()
    _setup_page_layout(doc)
    _setup_document_styles(doc, scheme)

    # Cover page
    _build_cover_page(doc, content, scheme)

    # Header / footer (applies to all pages after cover in a real Word doc)
    _build_header_footer(doc, content.get("title", ""), scheme)

    # Table of contents
    _build_toc(doc, scheme)

    # Abstract
    if content.get("abstract"):
        _build_abstract(doc, content["abstract"], scheme)
        doc.add_paragraph()

    # Body sections
    for section in content.get("sections", []):
        _render_section(doc, section, scheme)

    # References
    if content.get("references"):
        _paragraph_page_break(doc)
        _render_references(doc, content["references"], scheme)

    # Appendix
    if content.get("appendix"):
        _render_appendix(doc, content["appendix"], scheme)

    doc.save(output_path)
    return output_path
