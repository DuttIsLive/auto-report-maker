"""
report_generator.py
~~~~~~~~~~~~~~~~~~~
Orchestrates the full pipeline:

  1. Verify inputs and analyse any uploaded files.
  2. Summarise extracted guidelines (Verification Step).
  3. Generate structured report content (AI or placeholder).
  4. Build and save the .docx document.

This module is the single public entry-point for the auto-report-maker.
"""

from __future__ import annotations

import datetime
import os
from typing import Optional

from src.file_analyzer import analyze_files, FileAnalysisResult
from src.content_generator import generate_content
from src.document_formatter import build_document


# ---------------------------------------------------------------------------
# Verification step
# ---------------------------------------------------------------------------

def _verification_summary(
    topic: str,
    num_pages: Optional[int],
    file_paths: list[str],
    analysis: FileAnalysisResult,
    use_ai: bool,
) -> str:
    """Print a pre-generation checklist that satisfies the Verification Step."""
    lines = [
        "=" * 60,
        "  AUTO-REPORT-MAKER — PRE-GENERATION VERIFICATION",
        "=" * 60,
        "",
        f"  Topic         : {topic}",
        f"  Target pages  : {num_pages if num_pages else 'Not specified'}",
        f"  AI content    : {'Yes (OpenAI GPT-4o)' if use_ai and os.environ.get('OPENAI_API_KEY') else 'No (rich placeholder)'}",
        f"  Uploaded files: {len(file_paths)}",
        "",
    ]

    if file_paths:
        lines.append("  Files analysed:")
        for fp in file_paths:
            lines.append(f"    • {os.path.basename(fp)}")
        lines.append("")
        lines.append("  " + analysis.summary().replace("\n", "\n  "))
        lines.append("")

    lines += [
        "  Report will include:",
        "    ✓ Designed cover page",
        "    ✓ Table of contents (auto-update field)",
        "    ✓ Abstract / executive summary",
        "    ✓ Structured sections (H1 → H2 → H3)",
        "    ✓ Data tables with styled headers",
        "    ✓ Highlight boxes and callout blocks",
        "    ✓ Conclusion, recommendations, and future outlook",
        "    ✓ References",
        "    ✓ Times New Roman throughout",
        "    ✓ Professional colour scheme",
        "    ✓ Headers and footers with page numbers",
        "",
        f"  Minimum page target ({num_pages or 'N/A'}) will "
        + ("be met." if num_pages else "not be enforced (not specified)."),
        "",
        "=" * 60,
        "",
    ]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_report(
    topic: str,
    output_path: Optional[str] = None,
    num_pages: Optional[int] = None,
    file_paths: Optional[list[str]] = None,
    use_ai: bool = True,
    verbose: bool = True,
) -> str:
    """
    Generate a professional Word report and save it as a ``.docx`` file.

    Parameters
    ----------
    topic : str
        The report topic or title (required).
    output_path : str, optional
        Destination file path.  Defaults to ``<sanitised_topic>_report.docx``
        in the current working directory.
    num_pages : int, optional
        Minimum target page count.
    file_paths : list[str], optional
        Paths to uploaded template / guideline ``.docx`` files to analyse.
    use_ai : bool
        Whether to call the OpenAI API for content generation.
    verbose : bool
        If ``True``, print the verification summary and progress messages.

    Returns
    -------
    str
        Absolute path of the saved ``.docx`` file.

    Raises
    ------
    ValueError
        If *topic* is empty.
    """
    topic = topic.strip()
    if not topic:
        raise ValueError("Report topic must not be empty.")

    file_paths = [fp for fp in (file_paths or []) if os.path.isfile(fp)]

    # Step 1 – analyse uploaded files
    analysis: FileAnalysisResult = analyze_files(file_paths)

    # Step 2 – verification summary
    actual_use_ai = use_ai and bool(os.environ.get("OPENAI_API_KEY"))
    if verbose:
        print(_verification_summary(topic, num_pages, file_paths, analysis, use_ai))

    # Step 3 – generate content
    if verbose:
        print("  Generating report content …")
    content = generate_content(topic, num_pages=num_pages, analysis=analysis, use_ai=use_ai)

    # Ensure date is populated
    if not content.get("date"):
        content["date"] = datetime.date.today().strftime("%B %Y")

    # Step 4 – build document
    if output_path is None:
        safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in topic)
        safe_name = safe_name.strip().replace(" ", "_")[:50]
        output_path = f"{safe_name}_report.docx"

    if verbose:
        print(f"  Building document → {output_path} …")

    saved_path = build_document(content, output_path)

    if verbose:
        print(f"  ✓ Report saved: {saved_path}\n")

    return os.path.abspath(saved_path)
