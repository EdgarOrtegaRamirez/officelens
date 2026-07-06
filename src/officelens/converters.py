"""Format converters — DOCX→Markdown, XLSX→CSV/JSON, PPTX→Markdown."""

from __future__ import annotations

import csv
import io
import json
from pathlib import Path

from .parsers import DocType, Document, Table


def docx_to_markdown(doc: Document) -> str:
    """Convert DOCX document to Markdown."""
    lines: list[str] = []

    if doc.title:
        lines.append(f"# {doc.title}")
        lines.append("")

    for para in doc.paragraphs:
        if para.level > 0:
            prefix = "#" * min(para.level, 6)
            lines.append(f"{prefix} {para.text}")
        elif para.style and "List" in para.style:
            lines.append(f"- {para.text}")
        else:
            lines.append(para.text)
        lines.append("")

    for i, table in enumerate(doc.tables):
        if i > 0:
            lines.append("")
        lines.append(_table_to_markdown(table))

    return "\n".join(lines).strip() + "\n"


def _table_to_markdown(table: Table) -> str:
    """Convert a Table to Markdown table syntax."""
    grid = table.to_dict()
    if not grid:
        return ""

    lines: list[str] = []
    # Header
    header = grid[0]
    lines.append("| " + " | ".join(str(c) for c in header) + " |")
    lines.append("| " + " | ".join("---" for _ in header) + " |")
    # Data rows
    for row in grid[1:]:
        lines.append("| " + " | ".join(str(c) for c in row) + " |")
    return "\n".join(lines)


def xlsx_to_csv(doc: Document, sheet_name: str | None = None, delimiter: str = ",") -> str:
    """Convert XLSX to CSV string. If sheet_name is None, first sheet is used."""
    output = io.StringIO()
    writer = csv.writer(output, delimiter=delimiter)

    worksheets = doc.worksheets
    if sheet_name:
        worksheets = [ws for ws in worksheets if ws.name == sheet_name]
        if not worksheets:
            raise ValueError(f"Sheet '{sheet_name}' not found")

    for ws in worksheets:
        for row in ws.data:
            writer.writerow([str(c) if c is not None else "" for c in row])

    return output.getvalue()


def xlsx_to_json(doc: Document, sheet_name: str | None = None, use_headers: bool = True) -> str:
    """Convert XLSX to JSON. If use_headers, first row is used as keys."""
    worksheets = doc.worksheets
    if sheet_name:
        worksheets = [ws for ws in worksheets if ws.name == sheet_name]
        if not worksheets:
            raise ValueError(f"Sheet '{sheet_name}' not found")

    result: dict[str, Any] = {}
    for ws in worksheets:
        if use_headers and ws.headers:
            records: list[dict[str, Any]] = []
            for row in ws.data[1:]:
                record: dict[str, Any] = {}
                for i, header in enumerate(ws.headers):
                    if i < len(row):
                        record[header] = row[i]
                records.append(record)
            result[ws.name] = records
        else:
            result[ws.name] = ws.data

    return json.dumps(result, indent=2, default=str)


def pptx_to_markdown(doc: Document, include_notes: bool = True) -> str:
    """Convert PPTX to Markdown."""
    lines: list[str] = []

    if doc.title:
        lines.append(f"# {doc.title}")
        lines.append("")

    for slide in doc.slides:
        lines.append(f"## Slide {slide.number}")
        lines.append("")
        if slide.title:
            lines.append(f"### {slide.title}")
            lines.append("")

        for para in slide.paragraphs:
            lines.append(para.text)
        lines.append("")

        for table in slide.tables:
            lines.append(_table_to_markdown(table))
            lines.append("")

        if include_notes and slide.notes:
            lines.append(f"<!-- Notes: {slide.notes} -->")
            lines.append("")

    return "\n".join(lines).strip() + "\n"


def to_markdown(doc: Document, **kwargs) -> str:
    """Convert any document to Markdown."""
    if doc.doc_type == DocType.DOCX:
        return docx_to_markdown(doc)
    elif doc.doc_type == DocType.PPTX:
        return pptx_to_markdown(doc, **kwargs)
    elif doc.doc_type == DocType.XLSX:
        # XLSX → Markdown table
        lines: list[str] = []
        if doc.title:
            lines.append(f"# {doc.title}")
            lines.append("")
        for ws in doc.worksheets:
            lines.append(f"## {ws.name}")
            lines.append("")
            if ws.data:
                lines.append(_data_to_markdown_table(ws.data))
            lines.append("")
        return "\n".join(lines).strip() + "\n"
    else:
        raise ValueError(f"Cannot convert {doc.doc_type.value} to Markdown")


def _data_to_markdown_table(data: list[list]) -> str:
    """Convert 2D data list to markdown table."""
    if not data:
        return ""
    max_cols = max(len(row) for row in data)
    normalized = [list(row) + [""] * (max_cols - len(row)) for row in data]

    lines: list[str] = []
    header = normalized[0]
    lines.append("| " + " | ".join(str(c) for c in header) + " |")
    lines.append("| " + " | ".join("---" for _ in header) + " |")
    for row in normalized[1:]:
        lines.append("| " + " | ".join(str(c) for c in row) + " |")
    return "\n".join(lines)


def to_csv(doc: Document, **kwargs) -> str:
    """Convert document to CSV."""
    if doc.doc_type == DocType.XLSX:
        return xlsx_to_csv(doc, **kwargs)
    else:
        raise ValueError(f"Cannot convert {doc.doc_type.value} to CSV (only XLSX supported)")


def to_json(doc: Document, **kwargs) -> str:
    """Convert document to JSON."""
    if doc.doc_type == DocType.XLSX:
        return xlsx_to_json(doc, **kwargs)
    else:
        # Generic JSON from document model
        import dataclasses

        data = dataclasses.asdict(doc)
        # Remove non-serializable items
        for key in list(data.keys()):
            if isinstance(data[key], DocType):
                data[key] = data[key].value
        return json.dumps(data, indent=2, default=str)


def convert_file(input_path: str | Path, output_path: str | Path, fmt: str) -> str:
    """Convert a file and write to output_path. Returns the output path."""
    from .parsers import parse

    doc = parse(input_path)
    input_path = Path(input_path)
    output_path = Path(output_path)

    if fmt == "markdown" or fmt == "md":
        content = to_markdown(doc)
        if not output_path.suffix:
            output_path = output_path.with_suffix(".md")
    elif fmt == "csv":
        content = to_csv(doc)
        if not output_path.suffix:
            output_path = output_path.with_suffix(".csv")
    elif fmt == "json":
        content = to_json(doc)
        if not output_path.suffix:
            output_path = output_path.with_suffix(".json")
    elif fmt == "text" or fmt == "txt":
        content = doc.plain_text()
        if not output_path.suffix:
            output_path = output_path.with_suffix(".txt")
    else:
        raise ValueError(f"Unsupported output format: {fmt}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    return str(output_path)


# Need Any for xlsx_to_json
from typing import Any  # noqa: E402
