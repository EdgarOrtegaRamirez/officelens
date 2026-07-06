# OfficeLens

**Office Document Analysis & Conversion Toolkit** — parse, extract, convert, and analyze DOCX, XLSX, and PPTX files from the command line.

## Features

- **Multi-format support**: DOCX (Word), XLSX (Excel), PPTX (PowerPoint)
- **Text extraction**: Pull all text content from any Office document
- **Table extraction**: Extract tables to structured format
- **Format conversion**: DOCX→Markdown, XLSX→CSV/JSON, PPTX→Markdown
- **Document analysis**: Word count, readability scoring, vocabulary analysis
- **Content search**: Full-text search across any Office document
- **Batch processing**: Process multiple files at once
- **Metadata extraction**: Author, dates, custom properties
- **Sheet navigation**: Browse XLSX worksheets and structure

## Installation

```bash
pip install officelens
```

Or from source:

```bash
git clone https://github.com/EdgarOrtegaRamirez/officelens.git
cd officelens
pip install -e .
```

## Quick Start

```bash
# Extract text from a Word document
officelens extract document.docx

# Convert Word to Markdown
officelens convert document.docx -f markdown -o output.md

# Get document statistics
officelens info document.docx

# Search for content
officelens search document.docx "search term"

# List Excel sheets
officelens sheets workbook.xlsx

# Extract tables
officelens tables document.docx

# Batch process multiple files
officelens batch *.docx *.xlsx -f markdown -o output/
```

## CLI Commands

| Command | Description |
|---------|-------------|
| `extract` | Extract text content from an Office document |
| `convert` | Convert between formats (Markdown, CSV, JSON, Text) |
| `info` | Show document statistics and metadata |
| `search` | Full-text search within a document |
| `tables` | Extract and display tables |
| `sheets` | List XLSX worksheets and structure |
| `batch` | Process multiple documents at once |

## Python API

```python
from officelens.parsers import parse
from officelens.converters import to_markdown, to_csv, to_json
from officelens.analyzers import analyze, search_content

# Parse any Office document
doc = parse("document.docx")

# Get plain text
print(doc.plain_text())

# Convert to Markdown
md = to_markdown(doc)

# Analyze document
stats = analyze(doc)
print(f"Words: {stats.word_count}, Readability: {stats.readability_score}")

# Search content
matches = search_content(doc, "keyword")
```

## Supported Formats

| Input | Output Formats |
|-------|---------------|
| `.docx` | Text, Markdown, JSON |
| `.xlsx` | Text, CSV, JSON, Markdown |
| `.pptx` | Text, Markdown, JSON |

## Dependencies

- `python-docx` — DOCX parsing
- `openpyxl` — XLSX parsing
- `python-pptx` — PPTX parsing
- `click` — CLI framework
- `rich` — Terminal output formatting

## License

MIT
