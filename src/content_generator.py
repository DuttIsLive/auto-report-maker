"""
content_generator.py
~~~~~~~~~~~~~~~~~~~~
Generates structured report content using the OpenAI Chat API.

If no API key is available, a richly-populated placeholder report is returned
so that the document formatter can still produce a valid .docx output.

The structured content schema returned by this module is a plain Python dict
that the document_formatter can consume directly.
"""

from __future__ import annotations

import json
import os
import textwrap
from typing import Optional

from src.file_analyzer import FileAnalysisResult

# ---------------------------------------------------------------------------
# Schema helpers
# ---------------------------------------------------------------------------

CONTENT_SCHEMA_DESCRIPTION = textwrap.dedent("""
Return a JSON object that strictly follows this schema:

{
  "title": "<report title>",
  "subtitle": "<optional subtitle or empty string>",
  "author": "<Author Name or 'Author Name'>",
  "institution": "<Institution or 'Institution Name'>",
  "date": "<Month Year>",
  "abstract": "<200-300 word executive summary / abstract>",
  "sections": [
    {
      "heading": "<section heading>",
      "level": <1|2|3>,
      "blocks": [
        {"type": "paragraph", "text": "<text>"},
        {"type": "bullet_list", "items": ["<item1>", "<item2>"]},
        {"type": "numbered_list", "items": ["<step1>", "<step2>"]},
        {
          "type": "highlight_box",
          "title": "<box title or empty string>",
          "text": "<box body text>"
        },
        {
          "type": "table",
          "caption": "<optional caption>",
          "headers": ["<col1>", "<col2>"],
          "rows": [["<val>", "<val>"], ["<val>", "<val>"]]
        },
        {
          "type": "quote_block",
          "text": "<quotation text>",
          "source": "<source attribution>"
        },
        {
          "type": "figure",
          "caption": "<figure caption>"
        }
      ]
    }
  ],
  "references": ["<formatted reference 1>", "<formatted reference 2>"],
  "appendix": "<optional appendix text or null>"
}

Rules:
- Sections at level 1 are major sections. Levels 2 and 3 are subsections.
- Every major section (level 1) must have a heading.
- Include a mix of paragraphs, bullet lists, at least one table, and at
  least two highlight_box blocks across the whole report.
- References must use the specified citation style.
- All strings must be plain text — no markdown, no asterisks, no pound signs.
- Return ONLY the JSON object and nothing else.
""").strip()


def _system_prompt(tone: str, citation_style: Optional[str]) -> str:
    cite = citation_style or "APA"
    return (
        "You are an expert professional report writer and researcher. "
        f"Write in a {tone} tone. Use {cite} citation format for references. "
        "Produce comprehensive, well-researched, logically structured content. "
        "Avoid filler. Every paragraph must add meaningful information."
    )


def _user_prompt(
    topic: str,
    num_pages: Optional[int],
    analysis: Optional[FileAnalysisResult],
    tone: str,
) -> str:
    page_instruction = ""
    if num_pages:
        page_instruction = (
            f"\nThe report must be detailed enough to fill at least {num_pages} "
            "pages in a standard Word document (12pt Times New Roman, 1-inch margins, "
            "1.15 line spacing). Generate extensive content accordingly."
        )

    template_instruction = ""
    if analysis and analysis.detected_sections:
        sections_list = ", ".join(analysis.detected_sections[:8])
        template_instruction = (
            f"\nThe uploaded template suggests these sections: {sections_list}. "
            "Follow this structure where appropriate."
        )
    if analysis and analysis.raw_text_excerpt:
        template_instruction += (
            "\nAdditional context from uploaded files (use for style reference only):\n"
            + analysis.raw_text_excerpt[:1000]
        )

    return (
        f"Generate a complete, professional report on the following topic:\n\n"
        f"Topic: {topic}\n"
        f"{page_instruction}"
        f"{template_instruction}\n\n"
        f"The report must include:\n"
        "- A descriptive title and subtitle\n"
        "- An abstract (200–300 words)\n"
        "- Introduction with background and context\n"
        "- At least 4 major sections (level 1 headings) with subsections\n"
        "- At least one data table\n"
        "- At least two highlight_box blocks for key insights\n"
        "- A conclusion with recommendations\n"
        "- A future outlook section\n"
        "- A references list (at least 6 credible sources)\n\n"
        f"{CONTENT_SCHEMA_DESCRIPTION}"
    )


# ---------------------------------------------------------------------------
# OpenAI backend
# ---------------------------------------------------------------------------

def _call_openai(
    topic: str,
    num_pages: Optional[int],
    analysis: Optional[FileAnalysisResult],
    tone: str,
    citation_style: Optional[str],
) -> dict:
    try:
        from openai import OpenAI  # lazy import to avoid hard dependency at module load
    except ImportError as exc:
        raise RuntimeError(
            "openai package is not installed. Run: pip install openai"
        ) from exc

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise ValueError(
            "OPENAI_API_KEY environment variable is not set. "
            "Please set it or run without an API key to get a placeholder report."
        )

    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": _system_prompt(tone, citation_style)},
            {"role": "user", "content": _user_prompt(topic, num_pages, analysis, tone)},
        ],
        response_format={"type": "json_object"},
        temperature=0.7,
        max_tokens=8000,
    )
    raw = response.choices[0].message.content
    return json.loads(raw)


# ---------------------------------------------------------------------------
# Placeholder content (used when no API key is present)
# ---------------------------------------------------------------------------

def _placeholder_content(topic: str, num_pages: Optional[int]) -> dict:
    """Return a well-structured placeholder report for *topic*."""
    pages_note = (
        f"This report is designed to span approximately {num_pages} page(s). "
        if num_pages
        else ""
    )
    return {
        "title": topic,
        "subtitle": "A Comprehensive Professional Report",
        "author": "Author Name",
        "institution": "Institution Name",
        "date": "2025",
        "abstract": (
            f"{pages_note}"
            f"This report presents a thorough examination of {topic}. "
            "It covers the foundational concepts, current state of knowledge, "
            "key challenges, and emerging opportunities. Through systematic analysis "
            "of available evidence and established frameworks, this document provides "
            "actionable insights and evidence-based recommendations for practitioners "
            "and decision-makers. The report is structured to guide the reader from "
            "a contextual overview through detailed analysis to forward-looking "
            "conclusions, ensuring both breadth and depth of coverage."
        ),
        "sections": [
            {
                "heading": "Introduction",
                "level": 1,
                "blocks": [
                    {
                        "type": "paragraph",
                        "text": (
                            f"The subject of {topic} has attracted significant attention "
                            "across academic, professional, and policy communities in "
                            "recent years. Rapid developments in technology, shifting "
                            "societal priorities, and evolving global dynamics have "
                            "combined to elevate the importance of this field. "
                            "This report offers a structured analysis grounded in "
                            "current research and professional best practices."
                        ),
                    },
                    {
                        "type": "highlight_box",
                        "title": "Purpose of This Report",
                        "text": (
                            f"This document aims to provide a comprehensive overview "
                            f"of {topic}, identifying key trends, critical challenges, "
                            "and practical recommendations to inform strategic decision-making."
                        ),
                    },
                ],
            },
            {
                "heading": "Background and Context",
                "level": 2,
                "blocks": [
                    {
                        "type": "paragraph",
                        "text": (
                            "Understanding the historical context and foundational "
                            "principles is essential before engaging with contemporary "
                            "developments. This section traces the evolution of the "
                            f"field and establishes the conceptual framework for "
                            f"the analysis of {topic}."
                        ),
                    },
                    {
                        "type": "bullet_list",
                        "items": [
                            "Historical development and milestone events",
                            "Key stakeholders and their roles",
                            "Regulatory and institutional landscape",
                            "Foundational theoretical frameworks",
                        ],
                    },
                ],
            },
            {
                "heading": "Analysis and Findings",
                "level": 1,
                "blocks": [
                    {
                        "type": "paragraph",
                        "text": (
                            "This section presents the core analytical findings "
                            "derived from a review of current literature, empirical "
                            "data, and expert commentary. The analysis is structured "
                            "around three primary dimensions: quantitative performance "
                            "indicators, qualitative assessments, and comparative "
                            "benchmarking against established standards."
                        ),
                    },
                    {
                        "type": "table",
                        "caption": "Summary of Key Performance Indicators",
                        "headers": ["Indicator", "Current Value", "Benchmark", "Status"],
                        "rows": [
                            ["Metric A", "78%", "75%", "Above Target"],
                            ["Metric B", "62%", "70%", "Below Target"],
                            ["Metric C", "91%", "85%", "Exceeds Target"],
                            ["Metric D", "55%", "60%", "Needs Improvement"],
                        ],
                    },
                    {
                        "type": "highlight_box",
                        "title": "Key Insight",
                        "text": (
                            "Three of the four primary indicators show performance "
                            "at or above benchmark levels. However, targeted "
                            "intervention is recommended for Metric B and Metric D "
                            "to ensure overall programme objectives are met."
                        ),
                    },
                ],
            },
            {
                "heading": "Quantitative Analysis",
                "level": 2,
                "blocks": [
                    {
                        "type": "paragraph",
                        "text": (
                            "Quantitative analysis reveals statistically significant "
                            "patterns in the dataset. Using descriptive and inferential "
                            "statistical methods, the study identified positive "
                            "correlations between investment levels and outcome quality, "
                            "suggesting that resource allocation is a key determinant "
                            "of success in this domain."
                        ),
                    },
                    {
                        "type": "numbered_list",
                        "items": [
                            "Data collection and preparation methodology",
                            "Statistical analysis framework applied",
                            "Primary findings and significance levels",
                            "Limitations of the quantitative approach",
                        ],
                    },
                ],
            },
            {
                "heading": "Qualitative Assessment",
                "level": 2,
                "blocks": [
                    {
                        "type": "paragraph",
                        "text": (
                            "Qualitative research methods, including structured "
                            "interviews and thematic analysis of documentary sources, "
                            "were employed to complement the quantitative findings. "
                            "Themes emerging from this analysis include capacity "
                            "constraints, communication gaps, and the need for "
                            "stronger cross-sector collaboration."
                        ),
                    },
                    {
                        "type": "quote_block",
                        "text": (
                            "Effective implementation requires not just adequate "
                            "resources, but a coherent strategy that aligns "
                            "organisational goals with operational realities."
                        ),
                        "source": "Domain Expert, 2024",
                    },
                ],
            },
            {
                "heading": "Challenges and Risk Factors",
                "level": 1,
                "blocks": [
                    {
                        "type": "paragraph",
                        "text": (
                            f"Several challenges impede optimal progress in {topic}. "
                            "These range from structural barriers embedded in "
                            "institutional frameworks to more immediate operational "
                            "constraints. Understanding these challenges is a "
                            "prerequisite for developing effective mitigation strategies."
                        ),
                    },
                    {
                        "type": "table",
                        "caption": "Risk Register",
                        "headers": ["Risk", "Likelihood", "Impact", "Mitigation Strategy"],
                        "rows": [
                            ["Resource shortfall", "High", "High", "Phased budgeting"],
                            ["Stakeholder misalignment", "Medium", "High", "Engagement programme"],
                            ["Technology obsolescence", "Low", "Medium", "Regular review cycle"],
                            ["Regulatory change", "Medium", "Medium", "Compliance monitoring"],
                        ],
                    },
                ],
            },
            {
                "heading": "Opportunities and Strategic Directions",
                "level": 1,
                "blocks": [
                    {
                        "type": "paragraph",
                        "text": (
                            "Despite the challenges identified, the landscape presents "
                            "substantial opportunities for growth and improvement. "
                            "Strategic investment in innovation, talent development, "
                            "and partnerships can unlock significant value and advance "
                            f"the objectives associated with {topic}."
                        ),
                    },
                    {
                        "type": "bullet_list",
                        "items": [
                            "Technology-enabled efficiency gains",
                            "Expanded stakeholder collaboration",
                            "Policy reform to enable innovation",
                            "Capacity building and skills development",
                            "Data-driven decision-making frameworks",
                        ],
                    },
                ],
            },
            {
                "heading": "Conclusion",
                "level": 1,
                "blocks": [
                    {
                        "type": "paragraph",
                        "text": (
                            f"This report has provided a comprehensive analysis of "
                            f"{topic}, examining the current state, key challenges, "
                            "and strategic opportunities. The findings indicate that "
                            "while progress has been made, a more coordinated and "
                            "resource-adequate approach is required to fully realise "
                            "the potential of this domain."
                        ),
                    },
                ],
            },
            {
                "heading": "Recommendations",
                "level": 2,
                "blocks": [
                    {
                        "type": "numbered_list",
                        "items": [
                            "Establish a dedicated steering committee to oversee implementation.",
                            "Allocate ring-fenced funding for critical capability development.",
                            "Develop a monitoring and evaluation framework with clear KPIs.",
                            "Invest in stakeholder communication and engagement strategies.",
                            "Commission a follow-up review within 12 months to assess progress.",
                        ],
                    },
                ],
            },
            {
                "heading": "Future Outlook",
                "level": 2,
                "blocks": [
                    {
                        "type": "paragraph",
                        "text": (
                            f"The trajectory of {topic} over the next five to ten years "
                            "will be shaped by technological advancement, shifting policy "
                            "environments, and evolving stakeholder expectations. "
                            "Organisations that invest proactively in adaptive capacity "
                            "and evidence-based governance will be best positioned to "
                            "thrive in this dynamic landscape."
                        ),
                    },
                ],
            },
        ],
        "references": [
            f"Smith, J., & Jones, A. (2023). Advances in {topic}. Journal of Professional Studies, 14(2), 45–67.",
            f"Brown, K. (2022). Strategic frameworks for {topic}. Oxford University Press.",
            "Williams, R., & Taylor, S. (2021). Quantitative methods in applied research. Sage Publications.",
            f"Davis, M. (2023). {topic}: Challenges and opportunities. Policy Brief, 7(1), 1–12.",
            "International Standards Organisation. (2022). Guidelines for professional report writing. ISO Press.",
            f"Johnson, P., et al. (2024). A systematic review of {topic}. Review Journal, 8(3), 100–134.",
        ],
        "appendix": None,
    }


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_content(
    topic: str,
    num_pages: Optional[int] = None,
    analysis: Optional[FileAnalysisResult] = None,
    use_ai: bool = True,
) -> dict:
    """
    Generate structured report content for *topic*.

    Parameters
    ----------
    topic : str
        The report topic or title.
    num_pages : int, optional
        Minimum target page count for the generated report.
    analysis : FileAnalysisResult, optional
        Merged formatting guidelines from any uploaded template files.
    use_ai : bool
        If ``True`` (default) and ``OPENAI_API_KEY`` is set, the OpenAI API
        is called to generate content.  Otherwise a rich placeholder is used.

    Returns
    -------
    dict
        Structured report content ready for :func:`src.document_formatter.build_document`.
    """
    tone = analysis.detected_tone if analysis else "professional"
    citation_style = analysis.citation_style if analysis else None

    if use_ai and os.environ.get("OPENAI_API_KEY"):
        return _call_openai(topic, num_pages, analysis, tone, citation_style)

    return _placeholder_content(topic, num_pages)
