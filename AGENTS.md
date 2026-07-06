# AGENTS.md — AI Agent Guide for OfficeLens

## What This Project Does
OfficeLens is a Python CLI tool and library for analyzing, extracting, and converting Microsoft Office documents (DOCX, XLSX, PPTX). It provides:
- Multi-format parsing (DOCX, XLSX, PPTX) into a unified document model
- Text and table extraction
- Format conversion (DOCX→Markdown, XLSX→CSV/JSON, PPTX→Markdown)
- Document analysis (readability scoring, vocabulary analysis, statistics)
- Full-text search across any Office document
- Batch processing

## Architecture
- `src/officelens/parsers.py` — Core parsers that convert Office formats into `Document` data model
- `src/officelens/converters.py` — Format conversion functions (Markdown, CSV, JSON, Text)
- `src/officelens/analyzers.py` — Document analysis (stats, readability, search)
- `src/officelens/cli.py` — Click-based CLI with 7 commands
- `tests/` — pytest test suite with fixtures that create test documents

## Key Data Model
The central type is `Document` (in parsers.py) which contains:
- `paragraphs: list[Paragraph]` — for DOCX
- `worksheets: list[Worksheet]` — for XLSX
- `slides: list[Slide]` — for PPTX
- `tables: list[Table]` — extracted tables
- Metadata, word counts, etc.

## Dependencies
- `python-docx` — DOCX parsing
- `openpyxl` — XLSX parsing
- `python-pptx` — PPTX parsing
- `click` — CLI framework
- `rich` — Terminal output formatting

## Running Tests
```bash
pip install -e ".[dev]"
pytest tests/ -v
```

## Common Tasks
- Adding a new output format: Add converter function in `converters.py`, add CLI option in `cli.py`
- Adding a new analysis metric: Add to `analyzers.py`, update `DocStats` dataclass
- Supporting new document types: Add parser in `parsers.py`, extend `DocType` enum, update `parse()` dispatcher
