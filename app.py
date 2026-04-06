#!/usr/bin/env python3
"""
app.py — CLI entry-point for the auto-report-maker.

Usage
-----
  python app.py "Machine Learning in Healthcare"
  python app.py "Climate Change" --pages 10 --output climate_report.docx
  python app.py "Supply Chain" --pages 8 --files template.docx rubric.docx
  python app.py "AI Ethics" --no-ai   # use placeholder content (no API key needed)
"""

from __future__ import annotations

import argparse
import sys
import os

from dotenv import load_dotenv

load_dotenv()  # load .env file if present


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="auto-report-maker",
        description="Generate a professional .docx report with Times New Roman formatting.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "topic",
        help="Report topic or title (required).",
    )
    parser.add_argument(
        "--pages",
        type=int,
        default=None,
        metavar="N",
        help="Minimum target number of pages.",
    )
    parser.add_argument(
        "--output",
        default=None,
        metavar="FILE",
        help="Output .docx file path (default: <topic>_report.docx).",
    )
    parser.add_argument(
        "--files",
        nargs="*",
        default=[],
        metavar="FILE",
        help="Uploaded template / guideline .docx files to analyse.",
    )
    parser.add_argument(
        "--no-ai",
        action="store_true",
        help="Skip the OpenAI API call and use rich placeholder content instead.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress progress output.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

    # Validate attached files
    missing = [f for f in args.files if not os.path.isfile(f)]
    if missing:
        print(f"Error: The following files were not found: {', '.join(missing)}", file=sys.stderr)
        return 1

    from src.report_generator import generate_report

    try:
        saved = generate_report(
            topic=args.topic,
            output_path=args.output,
            num_pages=args.pages,
            file_paths=args.files or [],
            use_ai=not args.no_ai,
            verbose=not args.quiet,
        )
        if args.quiet:
            print(saved)
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        print(f"Unexpected error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
