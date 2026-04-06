# auto-report-maker

An AI-powered Microsoft Word (`.docx`) report designer and content generator that creates highly professional, visually striking, and structurally sophisticated reports.

---

## Features

- **Times New Roman throughout** — every element (headings, body, tables, captions, headers/footers) uses Times New Roman exclusively
- **AI content generation** — integrates with OpenAI GPT-4o to produce well-researched, professionally written reports; falls back to a richly structured placeholder when no API key is set
- **Template file analysis** — upload existing `.docx` templates or style guides and the tool will extract formatting rules, section structure, writing tone, and citation style before generating the report
- **Advanced document design** — topic-driven deterministic colour schemes (6 palettes), designed cover page, auto-updating Table of Contents, styled data tables, highlight/callout boxes, quote blocks, figure placeholders, and styled headers/footers with page numbers
- **Professional structure** — cover page → TOC → abstract → introduction → body sections (H1/H2/H3) → conclusion → recommendations → future outlook → references → appendix
- **Minimum page target** — specify a required number of pages; the content generator scales accordingly

---

## Installation

```bash
pip install -r requirements.txt
```

### Optional: OpenAI API key

Set your key to enable AI-generated content:

```bash
export OPENAI_API_KEY="sk-..."
```

Or create a `.env` file in the project root:

```
OPENAI_API_KEY=sk-...
```

---

## Usage

```bash
# Basic usage (AI content if key set, else rich placeholder)
python app.py "Machine Learning in Healthcare"

# Specify minimum page count
python app.py "Climate Change Policy" --pages 10

# Save to a specific file
python app.py "Supply Chain Resilience" --output supply_chain.docx

# Analyse uploaded template / rubric files
python app.py "AI Ethics" --files template.docx rubric.docx

# Force placeholder content (no API key needed)
python app.py "Quantum Computing" --no-ai

# Suppress progress output (prints only the output path)
python app.py "Cybersecurity Trends" --quiet
```

### Full options

| Flag | Description |
|------|-------------|
| `topic` | Report topic or title (required) |
| `--pages N` | Minimum target page count |
| `--output FILE` | Output `.docx` path (default: `<topic>_report.docx`) |
| `--files FILE …` | Uploaded `.docx` template / guideline files to analyse |
| `--no-ai` | Skip OpenAI API; use rich placeholder content |
| `--quiet` | Suppress verbose output |

---

## Project Structure

```
auto-report-maker/
├── app.py                    # CLI entry-point
├── requirements.txt
├── src/
│   ├── color_schemes.py      # 6 deterministic colour palettes
│   ├── content_generator.py  # OpenAI GPT-4o content generation + placeholder
│   ├── document_formatter.py # python-docx Word document builder
│   ├── file_analyzer.py      # .docx template / guide analyser
│   └── report_generator.py   # Orchestration pipeline
└── tests/
    ├── test_color_schemes.py
    ├── test_content_generator.py
    ├── test_document_formatter.py
    ├── test_file_analyzer.py
    └── test_report_generator.py
```

---

## Running Tests

```bash
pip install pytest
python -m pytest tests/ -v
```

---

## Document Design System

The formatter applies a distinct visual identity per topic:

| Element | Implementation |
|---------|---------------|
| Cover page | Full-width coloured band (primary colour), large bold title, italic subtitle, author/institution/date block |
| Colour system | 6 palettes; topic MD5 hash selects palette deterministically |
| Headings | H1 16–18pt bold + coloured rule; H2 14–15pt bold; H3 13pt bold italic |
| Abstract | Shaded box with coloured left border |
| Tables | Coloured header row, alternating row shading, full border |
| Highlight boxes | Shaded background, thick accent left border, bold title prefix |
| Quote blocks | Italic text, bold attribution, coloured left border |
| Figure placeholders | Grey shaded box with italic caption |
| TOC | Word auto-update field (`Ctrl+A`, `F9` to refresh in Word) |
| Headers | Right-aligned title with coloured underline rule |
| Footers | Centred page number field with top rule |
