"""
file_analyzer.py
~~~~~~~~~~~~~~~~
Extracts formatting rules, structural patterns, and style guidelines from
uploaded .docx template or guide files.  The results are merged into a
``FileAnalysisResult`` that the content generator can use to adjust the
generated report.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Optional

from docx import Document
from docx.shared import Pt


@dataclass
class FileAnalysisResult:
    """Merged guidelines extracted from one or more uploaded files."""

    # Structural
    detected_sections: list[str] = field(default_factory=list)
    heading_hierarchy: list[str] = field(default_factory=list)

    # Typography
    body_font: Optional[str] = None
    heading_font: Optional[str] = None
    body_font_size_pt: Optional[float] = None

    # Spacing
    line_spacing: Optional[float] = None
    space_after_pt: Optional[float] = None

    # Margins (in cm)
    margin_top_cm: Optional[float] = None
    margin_bottom_cm: Optional[float] = None
    margin_left_cm: Optional[float] = None
    margin_right_cm: Optional[float] = None

    # Style notes (free-form strings collected from headings / titles)
    style_notes: list[str] = field(default_factory=list)

    # Writing tone detected from content
    detected_tone: str = "professional"

    # Citation style hints
    citation_style: Optional[str] = None

    # Raw text extracted (used to pass context to the content generator)
    raw_text_excerpt: str = ""

    def summary(self) -> str:
        """Return a human-readable summary of the extracted guidelines."""
        lines = ["Extracted formatting guidelines:"]
        if self.body_font:
            lines.append(f"  Body font: {self.body_font}")
        if self.heading_font:
            lines.append(f"  Heading font: {self.heading_font}")
        if self.body_font_size_pt:
            lines.append(f"  Body size: {self.body_font_size_pt}pt")
        if self.line_spacing:
            lines.append(f"  Line spacing: {self.line_spacing}")
        if self.margin_left_cm:
            lines.append(
                f"  Margins: top={self.margin_top_cm}cm, "
                f"bottom={self.margin_bottom_cm}cm, "
                f"left={self.margin_left_cm}cm, "
                f"right={self.margin_right_cm}cm"
            )
        if self.detected_sections:
            lines.append(f"  Detected sections: {', '.join(self.detected_sections)}")
        if self.style_notes:
            lines.append(f"  Style notes: {'; '.join(self.style_notes[:5])}")
        lines.append(f"  Detected tone: {self.detected_tone}")
        if self.citation_style:
            lines.append(f"  Citation style: {self.citation_style}")
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_TONE_KEYWORDS = {
    "academic": ["research", "study", "analysis", "methodology", "hypothesis",
                 "literature", "findings", "abstract"],
    "technical": ["system", "architecture", "implementation", "algorithm",
                  "configuration", "deployment", "specification"],
    "business": ["revenue", "stakeholder", "roi", "market", "strategy",
                 "executive", "quarterly", "forecast"],
}

_CITATION_HINTS = {
    "APA": ["doi:", "retrieved from", "journal of", "(20", "(19"],
    "IEEE": ["[1]", "[2]", "[3]", "IEEE", "proc."],
    "Harvard": ["et al.,", "ibid", "op. cit."],
    "Chicago": ["footnote", "endnote", "ibid.", "chicago"],
}


def _detect_tone(text: str) -> str:
    text_lower = text.lower()
    scores = {tone: sum(1 for kw in kws if kw in text_lower)
              for tone, kws in _TONE_KEYWORDS.items()}
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "professional"


def _detect_citation_style(text: str) -> Optional[str]:
    for style, hints in _CITATION_HINTS.items():
        if sum(1 for h in hints if h.lower() in text.lower()) >= 2:
            return style
    return None


def _pt_from_emu(emu: Optional[int]) -> Optional[float]:
    """Convert EMU (English Metric Unit) to points."""
    if emu is None:
        return None
    return round(emu / 12700, 1)


def _cm_from_emu(emu: Optional[int]) -> Optional[float]:
    if emu is None:
        return None
    return round(emu / 360000, 2)


# ---------------------------------------------------------------------------
# Single-file analyser
# ---------------------------------------------------------------------------

def _analyze_single_file(path: str) -> FileAnalysisResult:
    result = FileAnalysisResult()

    try:
        doc = Document(path)
    except Exception as exc:
        result.style_notes.append(f"Could not open {os.path.basename(path)}: {exc}")
        return result

    # --- margins -----------------------------------------------------------
    if doc.sections:
        sec = doc.sections[0]
        result.margin_top_cm = _cm_from_emu(sec.top_margin)
        result.margin_bottom_cm = _cm_from_emu(sec.bottom_margin)
        result.margin_left_cm = _cm_from_emu(sec.left_margin)
        result.margin_right_cm = _cm_from_emu(sec.right_margin)

    full_text_parts: list[str] = []
    heading_texts: list[str] = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        full_text_parts.append(text)

        style_name = para.style.name if para.style else ""

        # Collect headings
        if "Heading" in style_name and text:
            heading_texts.append(text)
            result.detected_sections.append(text)

        # Typography from runs
        for run in para.runs:
            if run.font.name and not result.body_font:
                if "Heading" not in style_name:
                    result.body_font = run.font.name
                else:
                    result.heading_font = run.font.name
            if run.font.size and not result.body_font_size_pt:
                if "Heading" not in style_name:
                    result.body_font_size_pt = _pt_from_emu(run.font.size)

        # Spacing from paragraph format
        fmt = para.paragraph_format
        if fmt.line_spacing and not result.line_spacing:
            result.line_spacing = round(float(fmt.line_spacing), 2)
        if fmt.space_after and not result.space_after_pt:
            result.space_after_pt = _pt_from_emu(fmt.space_after)

    # Styles from document style definitions
    for style in doc.styles:
        if "Heading 1" in style.name and style.font.name:
            result.heading_font = style.font.name
        if style.name == "Normal" and style.font.name:
            if not result.body_font:
                result.body_font = style.font.name

    full_text = " ".join(full_text_parts)
    result.raw_text_excerpt = full_text[:2000]
    result.detected_tone = _detect_tone(full_text)
    result.citation_style = _detect_citation_style(full_text)
    result.heading_hierarchy = heading_texts[:10]

    return result


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def analyze_files(file_paths: list[str]) -> FileAnalysisResult:
    """
    Analyze one or more uploaded files and return merged guidelines.

    Parameters
    ----------
    file_paths : list[str]
        Paths to ``.docx`` files to analyse.

    Returns
    -------
    FileAnalysisResult
        Merged formatting and style guidelines extracted from all files.
    """
    if not file_paths:
        return FileAnalysisResult()

    results = [_analyze_single_file(p) for p in file_paths if p.endswith(".docx")]

    if not results:
        return FileAnalysisResult()

    # Merge: first non-None value wins; lists are concatenated (de-duped)
    merged = FileAnalysisResult()
    for r in results:
        if not merged.body_font and r.body_font:
            merged.body_font = r.body_font
        if not merged.heading_font and r.heading_font:
            merged.heading_font = r.heading_font
        if not merged.body_font_size_pt and r.body_font_size_pt:
            merged.body_font_size_pt = r.body_font_size_pt
        if not merged.line_spacing and r.line_spacing:
            merged.line_spacing = r.line_spacing
        if not merged.space_after_pt and r.space_after_pt:
            merged.space_after_pt = r.space_after_pt
        if not merged.margin_top_cm and r.margin_top_cm:
            merged.margin_top_cm = r.margin_top_cm
            merged.margin_bottom_cm = r.margin_bottom_cm
            merged.margin_left_cm = r.margin_left_cm
            merged.margin_right_cm = r.margin_right_cm
        if not merged.citation_style and r.citation_style:
            merged.citation_style = r.citation_style

        for section in r.detected_sections:
            if section not in merged.detected_sections:
                merged.detected_sections.append(section)
        merged.style_notes.extend(r.style_notes)
        if r.raw_text_excerpt:
            merged.raw_text_excerpt = (merged.raw_text_excerpt + " " + r.raw_text_excerpt).strip()[:4000]

    # Tone: take the most common
    tone_counts: dict[str, int] = {}
    for r in results:
        tone_counts[r.detected_tone] = tone_counts.get(r.detected_tone, 0) + 1
    merged.detected_tone = max(tone_counts, key=tone_counts.get)

    return merged
