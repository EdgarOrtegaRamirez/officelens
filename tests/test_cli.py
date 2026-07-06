"""Tests for CLI commands."""

import json

import pytest
from click.testing import CliRunner

from officelens.cli import main


@pytest.fixture
def runner():
    return CliRunner()


class TestCli:
    def test_version(self, runner):
        result = runner.invoke(main, ["--version"])
        assert result.exit_code == 0
        assert "officelens" in result.output

    def test_extract_docx(self, runner, sample_docx):
        result = runner.invoke(main, ["extract", str(sample_docx)])
        assert result.exit_code == 0
        assert "test document" in result.output.lower()

    def test_extract_json(self, runner, sample_docx):
        result = runner.invoke(main, ["extract", str(sample_docx), "-f", "json"])
        assert result.exit_code == 0
        data = json.loads(result.output)
        assert data["type"] == "docx"
        assert "text" in data

    def test_extract_to_file(self, runner, sample_docx, tmp_dir):
        out = tmp_dir / "extracted.txt"
        result = runner.invoke(main, ["extract", str(sample_docx), "-o", str(out)])
        assert result.exit_code == 0
        assert out.exists()

    def test_info_docx(self, runner, sample_docx):
        result = runner.invoke(main, ["info", str(sample_docx)])
        assert result.exit_code == 0
        assert "Test Document" in result.output

    def test_info_json(self, runner, sample_docx):
        result = runner.invoke(main, ["info", str(sample_docx), "-j"])
        assert result.exit_code == 0
        data = json.loads(result.output)
        assert data["doc_type"] == "docx"
        assert "word_count" in data

    def test_convert_docx_to_md(self, runner, sample_docx, tmp_dir):
        out = tmp_dir / "out.md"
        result = runner.invoke(main, ["convert", str(sample_docx), "-f", "markdown", "-o", str(out)])
        assert result.exit_code == 0
        assert out.exists()
        assert "# Test Document" in out.read_text()

    def test_convert_xlsx_to_csv(self, runner, sample_xlsx, tmp_dir):
        out = tmp_dir / "out.csv"
        result = runner.invoke(main, ["convert", str(sample_xlsx), "-f", "csv", "-o", str(out)])
        assert result.exit_code == 0
        assert out.exists()
        assert "Alice" in out.read_text()

    def test_convert_pptx_to_md(self, runner, sample_pptx, tmp_dir):
        out = tmp_dir / "out.md"
        result = runner.invoke(main, ["convert", str(sample_pptx), "-f", "markdown", "-o", str(out)])
        assert result.exit_code == 0
        assert out.exists()
        assert "Slide 1" in out.read_text()

    def test_search_found(self, runner, sample_docx):
        result = runner.invoke(main, ["search", str(sample_docx), "Alice"])
        assert result.exit_code == 0
        assert "Alice" in result.output

    def test_search_not_found(self, runner, sample_docx):
        result = runner.invoke(main, ["search", str(sample_docx), "Nonexistent123"])
        assert result.exit_code == 0
        assert "No matches" in result.output

    def test_search_json(self, runner, sample_docx):
        result = runner.invoke(main, ["search", str(sample_docx), "Alice", "-j"])
        assert result.exit_code == 0
        data = json.loads(result.output)
        assert len(data) > 0

    def test_tables_docx(self, runner, sample_docx):
        result = runner.invoke(main, ["tables", str(sample_docx)])
        assert result.exit_code == 0
        assert "Table 1" in result.output
        assert "Alice" in result.output

    def test_sheets_xlsx(self, runner, sample_xlsx):
        result = runner.invoke(main, ["sheets", str(sample_xlsx)])
        assert result.exit_code == 0
        assert "Employees" in result.output
        assert "Summary" in result.output

    def test_sheets_non_xlsx(self, runner, sample_docx):
        result = runner.invoke(main, ["sheets", str(sample_docx)])
        assert result.exit_code == 0
        assert "only works with XLSX" in result.output

    def test_batch_extract(self, runner, sample_docx, sample_xlsx, sample_pptx):
        result = runner.invoke(main, ["batch", str(sample_docx), str(sample_xlsx), str(sample_pptx)])
        assert result.exit_code == 0
        assert "===" in result.output

    def test_batch_to_dir(self, runner, sample_docx, sample_xlsx, tmp_dir):
        out_dir = tmp_dir / "batch_out"
        result = runner.invoke(
            main, ["batch", str(sample_docx), str(sample_xlsx), "-f", "markdown", "-o", str(out_dir)]
        )
        assert result.exit_code == 0
        assert out_dir.exists()
        assert len(list(out_dir.glob("*.md"))) >= 2
