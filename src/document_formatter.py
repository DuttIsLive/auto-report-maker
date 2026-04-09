"""
document_formatter.py
---------------------
Core Word-document builder for auto-report-maker.

All text uses Times New Roman throughout.  The cover page is built from a
borderless "cover band" table so that the coloured accent strip and the title
block can sit side-by-side cleanly.  The abstract lives inside its own
bordered table cell (the "abstract box") beneath the cover.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------

ACCENT_COLOR = RGBColor(0x1F, 0x45, 0x78)   # deep navy
LIGHT_BAND   = RGBColor(0xE8, 0xF0, 0xFE)   # pale blue
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT    = RGBColor(0x1A, 0x1A, 0x2E)
SUBTEXT      = RGBColor(0x4A, 0x4A, 0x6A)
HEADING_COLOR = RGBColor(0x1F, 0x45, 0x78)

# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class Section:
    """A single content section with a heading and body text."""
    heading: str
    body: str

@dataclass
class ReportData:
    """All content needed to produce one designer report."""
    title: str
    subtitle: str = ""
    author: str = ""
    date: str = ""
    abstract: str = ""
    sections: List[Section] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------

def _rgb_hex(rgb: RGBColor) -> str:
    """Return the 6-character uppercase hex string for an RGBColor."""
    return str(rgb).upper()


def _set_cell_bg(cell, rgb: RGBColor) -> None:
    """Fill a table cell with a solid background colour."""
    hex_color = _rgb_hex(rgb)
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    existing = tc_pr.find(qn("w:shd"))
    if existing is not None:
        tc_pr.remove(existing)
    tc_pr.append(shd)


def _set_cell_borders(cell, border_spec: dict) -> None:
    """
    Apply borders to a table cell.
    border_spec keys: top, bottom, left, right
    Each value is a dict with keys: val, sz, color
    """
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for side, attrs in border_spec.items():
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   attrs.get("val",   "single"))
        el.set(qn("w:sz"),    str(attrs.get("sz", 6)))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), attrs.get("color", "auto"))
        tc_borders.append(el)
    existing = tc_pr.find(qn("w:tcBorders"))
    if existing is not None:
        tc_pr.remove(existing)
    tc_pr.append(tc_borders)


def _remove_table_borders(table) -> None:
    """Remove all visible borders from a table (make it invisible)."""
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tbl_pr)
    tbl_borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "none")
        el.set(qn("w:sz"),    "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        tbl_borders.append(el)
    existing = tbl_pr.find(qn("w:tblBorders"))
    if existing is not None:
        tbl_pr.remove(existing)
    tbl_pr.append(tbl_borders)


def _set_cell_vertical_alignment(cell, alignment: str = "center") -> None:
    """Set vertical alignment of cell content (top / center / bottom)."""
    tc_pr = cell._tc.get_or_add_tcPr()
    v_align = OxmlElement("w:vAlign")
    v_align.set(qn("w:val"), alignment)
    existing = tc_pr.find(qn("w:vAlign"))
    if existing is not None:
        tc_pr.remove(existing)
    tc_pr.append(v_align)


def _set_row_height(row, height_cm: float) -> None:
    """Fix the height of a table row."""
    tr_pr = row._tr.get_or_add_trPr()
    tr_height = OxmlElement("w:trHeight")
    tr_height.set(qn("w:val"), str(int(height_cm * 567)))  # 567 EMU per cm approx
    tr_height.set(qn("w:hRule"), "exact")
    existing = tr_pr.find(qn("w:trHeight"))
    if existing is not None:
        tr_pr.remove(existing)
    tr_pr.append(tr_height)


def _add_run(paragraph, text: str, font_size: int, bold: bool = False,
             italic: bool = False, color: Optional[RGBColor] = None) -> None:
    """Append a styled run to an existing paragraph."""
    run = paragraph.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color


def _add_paragraph(cell_or_doc, text: str, font_size: int,
                   bold: bool = False, italic: bool = False,
                   color: Optional[RGBColor] = None,
                   alignment: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.LEFT,
                   space_before: int = 0, space_after: int = 0) -> None:
    """Add a paragraph with a single run to a cell or document."""
    if hasattr(cell_or_doc, "paragraphs") and hasattr(cell_or_doc, "add_paragraph"):
        para = cell_or_doc.add_paragraph()
    else:
        para = cell_or_doc.add_paragraph()
    para.alignment = alignment
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after  = Pt(space_after)
    _add_run(para, text, font_size, bold=bold, italic=italic, color=color)
    return para


# ---------------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------------

def _set_page_margins(doc: Document,
                      top: float = 2.54, bottom: float = 2.54,
                      left: float = 3.17, right: float = 3.17) -> None:
    """Set page margins in centimetres."""
    section = doc.sections[0]
    section.top_margin    = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin   = Cm(left)
    section.right_margin  = Cm(right)


# ---------------------------------------------------------------------------
# Cover page
# ---------------------------------------------------------------------------

def _build_cover(doc: Document, data: ReportData) -> None:
    """
    Build the designer cover page.

    Layout (2-column borderless table):
      Col 0 (narrow) – solid ACCENT_COLOR accent strip
      Col 1 (wide)   – title, subtitle, author, date on LIGHT_BAND background
    """
    # Full page width minus margins ≈ 21 – 3.17 – 3.17 = 14.66 cm
    cover_table = doc.add_table(rows=1, cols=2)
    cover_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    _remove_table_borders(cover_table)

    row = cover_table.rows[0]
    _set_row_height(row, 10)

    # Column widths
    cover_table.columns[0].width = Cm(1.2)
    cover_table.columns[1].width = Cm(13.46)

    accent_cell = row.cells[0]
    title_cell  = row.cells[1]

    _set_cell_bg(accent_cell, ACCENT_COLOR)
    _set_cell_bg(title_cell,  LIGHT_BAND)
    _set_cell_vertical_alignment(title_cell, "center")

    # Clear default empty paragraph from title cell
    for para in title_cell.paragraphs:
        para.clear()

    # Title
    title_para = title_cell.paragraphs[0]
    title_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title_para.paragraph_format.space_before = Pt(30)
    title_para.paragraph_format.space_after  = Pt(6)
    title_para.paragraph_format.left_indent  = Cm(0.6)
    _add_run(title_para, data.title, font_size=28, bold=True, color=ACCENT_COLOR)

    # Subtitle
    if data.subtitle:
        sub_para = title_cell.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sub_para.paragraph_format.space_before = Pt(0)
        sub_para.paragraph_format.space_after  = Pt(18)
        sub_para.paragraph_format.left_indent  = Cm(0.6)
        _add_run(sub_para, data.subtitle, font_size=16, italic=True, color=SUBTEXT)

    # Author / Date metadata
    if data.author or data.date:
        meta_parts = []
        if data.author:
            meta_parts.append(f"Author: {data.author}")
        if data.date:
            meta_parts.append(f"Date: {data.date}")
        meta_para = title_cell.add_paragraph()
        meta_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        meta_para.paragraph_format.space_before = Pt(8)
        meta_para.paragraph_format.space_after  = Pt(30)
        meta_para.paragraph_format.left_indent  = Cm(0.6)
        _add_run(meta_para, "   |   ".join(meta_parts), font_size=11, color=SUBTEXT)

    # Bottom accent line below cover table
    doc.add_paragraph()


# ---------------------------------------------------------------------------
# Abstract box
# ---------------------------------------------------------------------------

def _build_abstract(doc: Document, abstract: str) -> None:
    """
    Render the abstract inside a bordered single-cell table (abstract box).
    Left border is a thick accent bar; other borders are thin grey.
    """
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    _remove_table_borders(table)

    cell = table.rows[0].cells[0]
    _set_cell_bg(cell, LIGHT_BAND)

    border_spec = {
        "top":    {"val": "single", "sz": 4,  "color": "AAAAAA"},
        "bottom": {"val": "single", "sz": 4,  "color": "AAAAAA"},
        "right":  {"val": "single", "sz": 4,  "color": "AAAAAA"},
        "left":   {"val": "single", "sz": 18, "color": _rgb_hex(ACCENT_COLOR)},
    }
    _set_cell_borders(cell, border_spec)

    # Clear placeholder paragraph
    for para in cell.paragraphs:
        para.clear()

    heading_para = cell.paragraphs[0]
    heading_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    heading_para.paragraph_format.space_before = Pt(8)
    heading_para.paragraph_format.space_after  = Pt(4)
    heading_para.paragraph_format.left_indent  = Cm(0.3)
    _add_run(heading_para, "Abstract", font_size=11, bold=True, color=ACCENT_COLOR)

    body_para = cell.add_paragraph()
    body_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body_para.paragraph_format.space_before = Pt(0)
    body_para.paragraph_format.space_after  = Pt(8)
    body_para.paragraph_format.left_indent  = Cm(0.3)
    body_para.paragraph_format.right_indent = Cm(0.3)
    _add_run(body_para, abstract, font_size=11, italic=True, color=DARK_TEXT)

    doc.add_paragraph()


# ---------------------------------------------------------------------------
# Section builder
# ---------------------------------------------------------------------------

def _build_section(doc: Document, section: Section, index: int) -> None:
    """Render one content section with a styled heading and body text."""
    # Section number prefix
    number = f"{index}.  "

    heading_para = doc.add_paragraph()
    heading_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    heading_para.paragraph_format.space_before = Pt(14)
    heading_para.paragraph_format.space_after  = Pt(4)

    _add_run(heading_para, number,          font_size=14, bold=True,  color=ACCENT_COLOR)
    _add_run(heading_para, section.heading, font_size=14, bold=True,  color=HEADING_COLOR)

    # Thin underline rule
    rule_para = doc.add_paragraph()
    rule_para.paragraph_format.space_before = Pt(0)
    rule_para.paragraph_format.space_after  = Pt(6)
    # Build a bottom-border on this paragraph to mimic a horizontal rule
    p_pr = rule_para._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), _rgb_hex(ACCENT_COLOR))
    p_bdr.append(bottom)
    p_pr.append(p_bdr)

    body_para = doc.add_paragraph()
    body_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body_para.paragraph_format.space_before = Pt(4)
    body_para.paragraph_format.space_after  = Pt(10)
    _add_run(body_para, section.body, font_size=12, color=DARK_TEXT)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_report(data: ReportData, output_path: str) -> str:
    """
    Generate a designer Word report (.docx) from *data* and save it to
    *output_path*.  Returns the output path.
    """
    doc = Document()
    _set_page_margins(doc)

    _build_cover(doc, data)

    if data.abstract:
        _build_abstract(doc, data.abstract)

    for idx, section in enumerate(data.sections, start=1):
        _build_section(doc, section, idx)

    doc.save(output_path)
    return output_path
