"""
tests/test_content_generator.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Unit tests for the content_generator module.
"""

import pytest

from src.content_generator import generate_content, _placeholder_content
from src.file_analyzer import FileAnalysisResult


# ---------------------------------------------------------------------------
# Tests — _placeholder_content (always available, no API key needed)
# ---------------------------------------------------------------------------

def test_placeholder_content_returns_dict():
    result = _placeholder_content("Test Topic", num_pages=None)
    assert isinstance(result, dict)


def test_placeholder_content_required_keys():
    result = _placeholder_content("Test Topic", num_pages=None)
    for key in ("title", "subtitle", "author", "institution", "date",
                "abstract", "sections", "references"):
        assert key in result, f"Missing key: {key}"


def test_placeholder_content_title_matches():
    result = _placeholder_content("Renewable Energy", num_pages=None)
    assert result["title"] == "Renewable Energy"


def test_placeholder_content_sections_is_list():
    result = _placeholder_content("Climate Change", num_pages=None)
    assert isinstance(result["sections"], list)
    assert len(result["sections"]) >= 4


def test_placeholder_content_sections_have_required_fields():
    result = _placeholder_content("Healthcare AI", num_pages=None)
    for section in result["sections"]:
        assert "heading" in section
        assert "level" in section
        assert "blocks" in section
        assert isinstance(section["blocks"], list)


def test_placeholder_content_levels_are_valid():
    result = _placeholder_content("Supply Chain", num_pages=None)
    for section in result["sections"]:
        assert section["level"] in (1, 2, 3)


def test_placeholder_content_references_not_empty():
    result = _placeholder_content("Cybersecurity", num_pages=None)
    assert isinstance(result["references"], list)
    assert len(result["references"]) >= 4


def test_placeholder_content_has_table_block():
    result = _placeholder_content("Finance", num_pages=None)
    block_types = [
        b["type"]
        for s in result["sections"]
        for b in s["blocks"]
    ]
    assert "table" in block_types


def test_placeholder_content_has_highlight_box():
    result = _placeholder_content("Education", num_pages=None)
    block_types = [
        b["type"]
        for s in result["sections"]
        for b in s["blocks"]
    ]
    assert "highlight_box" in block_types


def test_placeholder_content_includes_page_note_when_pages_specified():
    result = _placeholder_content("Test", num_pages=10)
    assert "10 page" in result["abstract"]


def test_placeholder_content_no_page_note_when_none():
    result = _placeholder_content("Test", num_pages=None)
    assert "page" not in result["abstract"][:30]


# ---------------------------------------------------------------------------
# Tests — generate_content (placeholder path, no API key)
# ---------------------------------------------------------------------------

def test_generate_content_without_api_key(monkeypatch):
    """Without OPENAI_API_KEY set, should fall back to placeholder."""
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    result = generate_content("Machine Learning", num_pages=5, use_ai=True)
    assert isinstance(result, dict)
    assert result["title"] == "Machine Learning"


def test_generate_content_use_ai_false(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "fake-key")
    result = generate_content("Deep Learning", num_pages=None, use_ai=False)
    assert isinstance(result, dict)
    assert result["title"] == "Deep Learning"


def test_generate_content_with_analysis(monkeypatch):
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    analysis = FileAnalysisResult(
        detected_tone="academic",
        citation_style="APA",
        detected_sections=["Introduction", "Literature Review", "Conclusion"],
    )
    result = generate_content("Blockchain", analysis=analysis, use_ai=False)
    assert isinstance(result, dict)


def test_generate_content_all_blocks_have_type(monkeypatch):
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    result = generate_content("Data Science", use_ai=False)
    for section in result["sections"]:
        for block in section["blocks"]:
            assert "type" in block, f"Block missing 'type': {block}"
