"""
report_maker.py
---------------
High-level API and CLI entry point for auto-report-maker.

Usage (CLI):
    python -m src.report_maker \
        --title "My Report" \
        --subtitle "A deep dive" \
        --author "Jane Doe" \
        --date "April 2026" \
        --abstract "This report covers …" \
        --output my_report.docx

    Content sections can be supplied via --section-heading / --section-body
    (repeated as a pair for each section):
        --section-heading "Introduction" --section-body "Body text …"
"""

from __future__ import annotations

import argparse
import sys
from typing import List, Optional

from src.document_formatter import ReportData, Section, build_report


def make_report(
    title: str,
    output_path: str,
    subtitle: str = "",
    author: str = "",
    date: str = "",
    abstract: str = "",
    sections: Optional[List[Section]] = None,
) -> str:
    """
    Convenience wrapper: assemble a :class:`ReportData` and call
    :func:`build_report`.

    Returns the path to the generated .docx file.
    """
    data = ReportData(
        title=title,
        subtitle=subtitle,
        author=author,
        date=date,
        abstract=abstract,
        sections=sections or [],
    )
    return build_report(data, output_path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="report-maker",
        description="Generate a designer Word report.",
    )
    parser.add_argument("--title",    required=True,  help="Report title")
    parser.add_argument("--subtitle", default="",     help="Optional subtitle")
    parser.add_argument("--author",   default="",     help="Author name")
    parser.add_argument("--date",     default="",     help="Date string")
    parser.add_argument("--abstract", default="",     help="Abstract paragraph")
    parser.add_argument(
        "--section-heading",
        dest="section_headings",
        action="append",
        default=[],
        metavar="HEADING",
        help="Section heading (repeat for multiple sections)",
    )
    parser.add_argument(
        "--section-body",
        dest="section_bodies",
        action="append",
        default=[],
        metavar="BODY",
        help="Section body text (must match --section-heading count)",
    )
    parser.add_argument(
        "--output", "-o",
        default="report.docx",
        help="Output .docx file path (default: report.docx)",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> None:
    args = _parse_args(sys.argv[1:] if argv is None else argv)

    if len(args.section_headings) != len(args.section_bodies):
        print(
            "error: --section-heading and --section-body must be supplied in "
            "matching pairs.",
            file=sys.stderr,
        )
        sys.exit(1)

    sections = [
        Section(heading=h, body=b)
        for h, b in zip(args.section_headings, args.section_bodies)
    ]

    output = make_report(
        title=args.title,
        output_path=args.output,
        subtitle=args.subtitle,
        author=args.author,
        date=args.date,
        abstract=args.abstract,
        sections=sections,
    )
    print(f"Report saved to: {output}")


if __name__ == "__main__":
    main()
