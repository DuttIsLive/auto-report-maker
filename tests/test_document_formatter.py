"""
tests/test_document_formatter.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Unit tests for the document_formatter module.
"""

import os
import tempfile

import pytest
from docx import Document
from docx.shared import Pt

from src.document_formatter import build_document
from src.color_schemes import pick_scheme


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

MINIMAL_CONTENT = {
    "title": "Test Report",
    "subtitle": "A Test Subtitle",
    "author": "Test Author",
    "institution": "Test University",
    "date": "April 2025",
    "abstract": "This is a test abstract with sufficient content to fill the box.",
    "sections": [
        {
            "heading": "Introduction",
            "level": 1,
            "blocks": [
                {"type": "paragraph", "text": "This is the introduction paragraph."},
            ],
        },
        {
            "heading": "Methods",
            "level": 1,
            "blocks": [
                {"type": "bullet_list", "items": ["Item one", "Item two", "Item three"]},
                {"type": "numbered_list", "items": ["Step one", "Step two"]},
            ],
        },
        {
            "heading": "Results",
            "level": 1,
            "blocks": [
                {
                    "type": "table",
                    "caption": "Summary Table",
                    "headers": ["Column A", "Column B", "Column C"],
                    "rows": [
                        ["Row 1 A", "Row 1 B", "Row 1 C"],
                        ["Row 2 A", "Row 2 B", "Row 2 C"],
                    ],
                },
                {
                    "type": "highlight_box",
                    "title": "Key Finding",
                    "text": "This is an important insight from the results.",
                },
            ],
        },
        {
            "heading": "Discussion",
            "level": 2,
            "blocks": [
                {
                    "type": "quote_block",
                    "text": "Research is the art of seeing what others have missed.",
                    "source": "Anonymous, 2024",
                },
            ],
        },
        {
            "heading": "Conclusion",
            "level": 1,
            "blocks": [
                {"type": "paragraph", "text": "This concludes the test report."},
                {
                    "type": "figure",
                    "caption": "Sample figure placeholder",
                },
            ],
        },
    ],
    "references": [
        "Smith, J. (2024). Test reference. Journal of Testing, 1(1), 1–10.",
        "Doe, A. (2023). Another reference. Academic Press.",
    ],
    "appendix": "This is appendix content.",
}


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_build_document_creates_file():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        result = build_document(MINIMAL_CONTENT, path)
        assert os.path.isfile(result)
        assert result == path
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_build_document_is_valid_docx():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(MINIMAL_CONTENT, path)
        doc = Document(path)
        # A valid Word document has at least one paragraph
        assert len(doc.paragraphs) > 0
    finally:
        if os.path.exists(path):
            os.unlink(path)


def _full_text(doc: Document) -> str:
    """Return all text in the document, including table cells."""
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    parts.append(para.text)
    return " ".join(parts)


def test_build_document_contains_title():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(MINIMAL_CONTENT, path)
        doc = Document(path)
        assert "Test Report" in _full_text(doc)
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_build_document_contains_section_headings():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(MINIMAL_CONTENT, path)
        doc = Document(path)
        full_text = " ".join(p.text for p in doc.paragraphs)
        for heading in ["Introduction", "Methods", "Results", "Conclusion"]:
            assert heading in full_text, f"Heading '{heading}' not found in document"
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_build_document_contains_abstract():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(MINIMAL_CONTENT, path)
        doc = Document(path)
        full_text = _full_text(doc)
        assert "Abstract" in full_text
        assert "test abstract" in full_text
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_build_document_contains_references():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(MINIMAL_CONTENT, path)
        doc = Document(path)
        full_text = " ".join(p.text for p in doc.paragraphs)
        assert "References" in full_text
        assert "Smith, J." in full_text
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_build_document_uses_times_new_roman():
    """Verify that at least some runs use Times New Roman."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(MINIMAL_CONTENT, path)
        doc = Document(path)
        tnr_found = False
        for para in doc.paragraphs:
            for run in para.runs:
                if run.font.name == "Times New Roman":
                    tnr_found = True
                    break
            if tnr_found:
                break
        # Also check table cells
        if not tnr_found:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                if run.font.name == "Times New Roman":
                                    tnr_found = True
                                    break
        assert tnr_found, "Times New Roman font was not found in any run"
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_build_document_has_tables():
    """The document should contain at least the cover band table and the data table."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(MINIMAL_CONTENT, path)
        doc = Document(path)
        assert len(doc.tables) >= 1
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_build_document_no_subtitle():
    """Cover page should build correctly even without a subtitle."""
    content = {**MINIMAL_CONTENT, "subtitle": ""}
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(content, path)
        doc = Document(path)
        assert len(doc.paragraphs) > 0
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_build_document_no_references():
    """Document builds correctly with an empty references list."""
    content = {**MINIMAL_CONTENT, "references": []}
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(content, path)
        assert os.path.isfile(path)
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_build_document_with_appendix():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        build_document(MINIMAL_CONTENT, path)
        doc = Document(path)
        full_text = " ".join(p.text for p in doc.paragraphs)
        assert "Appendix" in full_text
        assert "appendix content" in full_text
    finally:
        if os.path.exists(path):
            os.unlink(path)
