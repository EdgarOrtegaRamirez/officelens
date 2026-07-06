"""Tests for converters."""

import json
from pathlib import Path

import pytest

from officelens.converters import (
    convert_file,
    docx_to_markdown,
    pptx_to_markdown,
    to_markdown,
    xlsx_to_csv,
    xlsx_to_json,
)
from officelens.parsers import parse


class TestDocxToMarkdown:
    def test_basic_conversion(self, sample_docx):
        doc = parse(sample_docx)
        md = docx_to_markdown(doc)
        assert "# Test Document" in md
        assert "# Introduction" in md
        assert "test document" in md.lower()

    def test_heading_levels(self, sample_docx):
        doc = parse(sample_docx)
        md = docx_to_markdown(doc)
        assert "# Introduction" in md  # level 1 → #
        assert "## Section 1" in md  # level 2 → ##

    def test_table_in_markdown(self, sample_docx):
        doc = parse(sample_docx)
        md = docx_to_markdown(doc)
        assert "| Name |" in md
        assert "| Alice |" in md
        assert "| --- |" in md

    def test_empty_paragraphs_filtered(self, sample_docx):
        doc = parse(sample_docx)
        md = docx_to_markdown(doc)
        # Should not have excessive blank lines
        assert "\n\n\n" not in md


class TestXlsxToCsv:
    def test_first_sheet(self, sample_xlsx):
        doc = parse(sample_xlsx)
        csv_str = xlsx_to_csv(doc)
        assert "Name,Age,City,Department" in csv_str
        assert "Alice" in csv_str
        assert "Bob" in csv_str

    def test_specific_sheet(self, sample_xlsx):
        doc = parse(sample_xlsx)
        csv_str = xlsx_to_csv(doc, sheet_name="Summary")
        assert "Metric,Value" in csv_str
        assert "Total Employees" in csv_str

    def test_invalid_sheet(self, sample_xlsx):
        doc = parse(sample_xlsx)
        with pytest.raises(ValueError, match="not found"):
            xlsx_to_csv(doc, sheet_name="Nonexistent")


class TestXlsxToJson:
    def test_first_sheet(self, sample_xlsx):
        doc = parse(sample_xlsx)
        json_str = xlsx_to_json(doc)
        data = json.loads(json_str)
        assert "Employees" in data
        assert len(data["Employees"]) == 3
        assert data["Employees"][0]["Name"] == "Alice"

    def test_with_headers(self, sample_xlsx):
        doc = parse(sample_xlsx)
        json_str = xlsx_to_json(doc, use_headers=True)
        data = json.loads(json_str)
        assert isinstance(data["Employees"], list)
        assert isinstance(data["Employees"][0], dict)

    def test_without_headers(self, sample_xlsx):
        doc = parse(sample_xlsx)
        json_str = xlsx_to_json(doc, use_headers=False)
        data = json.loads(json_str)
        assert isinstance(data["Employees"], list)
        assert isinstance(data["Employees"][0], list)


class TestPptxToMarkdown:
    def test_basic_conversion(self, sample_pptx):
        doc = parse(sample_pptx)
        md = pptx_to_markdown(doc)
        assert "## Slide 1" in md
        assert "Presentation Title" in md

    def test_multiple_slides(self, sample_pptx):
        doc = parse(sample_pptx)
        md = pptx_to_markdown(doc)
        assert "## Slide 1" in md
        assert "## Slide 2" in md
        assert "## Slide 3" in md

    def test_table_in_pptx(self, sample_pptx):
        doc = parse(sample_pptx)
        md = pptx_to_markdown(doc)
        assert "| Product |" in md
        assert "| Widget A |" in md

    def test_include_notes(self, sample_pptx):
        doc = parse(sample_pptx)
        md_with = pptx_to_markdown(doc, include_notes=True)
        md_without = pptx_to_markdown(doc, include_notes=False)
        # Notes don't exist in our test, but the option should work
        assert isinstance(md_with, str)
        assert isinstance(md_without, str)


class TestToMarkdown:
    def test_docx(self, sample_docx):
        doc = parse(sample_docx)
        md = to_markdown(doc)
        assert "# Test Document" in md

    def test_pptx(self, sample_pptx):
        doc = parse(sample_pptx)
        md = to_markdown(doc)
        assert "## Slide 1" in md

    def test_xlsx(self, sample_xlsx):
        doc = parse(sample_xlsx)
        md = to_markdown(doc)
        assert "| Name |" in md or "Employees" in md


class TestConvertFile:
    def test_docx_to_markdown_file(self, sample_docx, tmp_dir):
        out = tmp_dir / "output.md"
        result = convert_file(sample_docx, out, "markdown")
        assert Path(result).exists()
        content = Path(result).read_text()
        assert "# Test Document" in content

    def test_xlsx_to_csv_file(self, sample_xlsx, tmp_dir):
        out = tmp_dir / "output.csv"
        result = convert_file(sample_xlsx, out, "csv")
        assert Path(result).exists()
        content = Path(result).read_text()
        assert "Alice" in content

    def test_xlsx_to_json_file(self, sample_xlsx, tmp_dir):
        out = tmp_dir / "output.json"
        result = convert_file(sample_xlsx, out, "json")
        assert Path(result).exists()
        data = json.loads(Path(result).read_text())
        assert "Employees" in data

    def test_pptx_to_markdown_file(self, sample_pptx, tmp_dir):
        out = tmp_dir / "output.md"
        result = convert_file(sample_pptx, out, "markdown")
        assert Path(result).exists()
        content = Path(result).read_text()
        assert "Slide 1" in content

    def test_auto_extension(self, sample_docx, tmp_dir):
        out = tmp_dir / "output"  # No extension
        result = convert_file(sample_docx, out, "markdown")
        assert result.endswith(".md")
