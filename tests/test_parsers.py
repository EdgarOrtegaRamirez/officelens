"""Tests for DOCX parser."""

from officelens.parsers import DocType, parse


class TestDocxParser:
    def test_parse_returns_document(self, sample_docx):
        doc = parse(sample_docx)
        assert doc is not None
        assert doc.doc_type == DocType.DOCX

    def test_metadata(self, sample_docx):
        doc = parse(sample_docx)
        assert doc.title == "Test Document"
        assert doc.author == "Test Author"
        assert doc.subject == "Testing"

    def test_paragraphs_extracted(self, sample_docx):
        doc = parse(sample_docx)
        assert len(doc.paragraphs) > 0
        texts = [p.text for p in doc.paragraphs]
        assert any("test document" in t.lower() for t in texts)

    def test_headings_detected(self, sample_docx):
        doc = parse(sample_docx)
        headings = [p for p in doc.paragraphs if p.level > 0]
        assert len(headings) >= 3
        heading_texts = [h.text for h in headings]
        assert "Introduction" in heading_texts
        assert "Section 1" in heading_texts
        assert "Section 2" in heading_texts

    def test_heading_levels(self, sample_docx):
        doc = parse(sample_docx)
        intro = [p for p in doc.paragraphs if p.text == "Introduction"][0]
        assert intro.level == 1

    def test_tables_extracted(self, sample_docx):
        doc = parse(sample_docx)
        assert len(doc.tables) == 1
        table = doc.tables[0]
        assert table.rows == 3
        assert table.cols == 3

    def test_table_data(self, sample_docx):
        doc = parse(sample_docx)
        table = doc.tables[0]
        grid = table.to_dict()
        assert grid[0][0] == "Name"
        assert grid[1][0] == "Alice"
        assert grid[1][1] == "30"
        assert grid[2][2] == "London"

    def test_word_count(self, sample_docx):
        doc = parse(sample_docx)
        assert doc.word_count > 0

    def test_char_count(self, sample_docx):
        doc = parse(sample_docx)
        assert doc.char_count > 0
        assert doc.char_count >= doc.word_count

    def test_plain_text(self, sample_docx):
        doc = parse(sample_docx)
        text = doc.plain_text()
        assert "test document" in text.lower()
        assert "Alice" in text
        assert "Introduction" in text

    def test_image_count(self, sample_docx):
        doc = parse(sample_docx)
        assert doc.image_count == 0  # No images in test doc

    def test_doc_type_property(self, sample_docx):
        assert DocType.from_path("test.docx") == DocType.DOCX
        assert DocType.from_path("test.xlsx") == DocType.XLSX
        assert DocType.from_path("test.pptx") == DocType.PPTX
        assert DocType.from_path("test.txt") == DocType.UNKNOWN
