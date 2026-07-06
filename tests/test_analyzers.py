"""Tests for analyzers."""

from officelens.analyzers import analyze, search_content
from officelens.parsers import parse


class TestAnalyze:
    def test_docx_stats(self, sample_docx):
        doc = parse(sample_docx)
        stats = analyze(doc)
        assert stats.doc_type == "docx"
        assert stats.title == "Test Document"
        assert stats.word_count > 0
        assert stats.char_count > 0
        assert stats.paragraph_count > 0

    def test_docx_headings(self, sample_docx):
        doc = parse(sample_docx)
        stats = analyze(doc)
        assert stats.heading_count >= 3
        assert len(stats.heading_structure) >= 3

    def test_docx_tables(self, sample_docx):
        doc = parse(sample_docx)
        stats = analyze(doc)
        assert stats.table_count == 1

    def test_readability_score(self, sample_docx):
        doc = parse(sample_docx)
        stats = analyze(doc)
        assert 0 <= stats.readability_score <= 100
        assert stats.readability_grade != ""

    def test_vocabulary_richness(self, sample_docx):
        doc = parse(sample_docx)
        stats = analyze(doc)
        assert 0 < stats.vocabulary_richness <= 1.0

    def test_top_words(self, sample_docx):
        doc = parse(sample_docx)
        stats = analyze(doc)
        assert len(stats.top_words) > 0
        assert isinstance(stats.top_words[0], tuple)
        assert len(stats.top_words[0]) == 2

    def test_xlsx_stats(self, sample_xlsx):
        doc = parse(sample_xlsx)
        stats = analyze(doc)
        assert stats.doc_type == "xlsx"
        assert stats.word_count > 0

    def test_pptx_stats(self, sample_pptx):
        doc = parse(sample_pptx)
        stats = analyze(doc)
        assert stats.doc_type == "pptx"
        assert stats.word_count > 0

    def test_page_estimate(self, sample_docx):
        doc = parse(sample_docx)
        stats = analyze(doc)
        assert stats.page_estimate >= 1

    def test_avg_words_per_paragraph(self, sample_docx):
        doc = parse(sample_docx)
        stats = analyze(doc)
        assert stats.avg_words_per_paragraph > 0


class TestSearchContent:
    def test_search_found(self, sample_docx):
        doc = parse(sample_docx)
        matches = search_content(doc, "Alice")
        assert len(matches) > 0
        assert any("Alice" in m["match"] for m in matches)

    def test_search_not_found(self, sample_docx):
        doc = parse(sample_docx)
        matches = search_content(doc, "NonexistentWord123")
        assert len(matches) == 0

    def test_search_case_insensitive(self, sample_docx):
        doc = parse(sample_docx)
        matches = search_content(doc, "alice", case_sensitive=False)
        assert len(matches) > 0

    def test_search_case_sensitive(self, sample_docx):
        doc = parse(sample_docx)
        matches_upper = search_content(doc, "Alice", case_sensitive=True)
        matches_lower = search_content(doc, "alice", case_sensitive=True)
        # "Alice" should find results, "alice" should not
        assert len(matches_upper) > 0
        assert len(matches_lower) == 0

    def test_search_xlsx(self, sample_xlsx):
        doc = parse(sample_xlsx)
        matches = search_content(doc, "Engineering")
        assert len(matches) > 0
        assert any(m.get("sheet") == "Employees" for m in matches)

    def test_search_pptx(self, sample_pptx):
        doc = parse(sample_pptx)
        matches = search_content(doc, "Widget")
        assert len(matches) > 0

    def test_search_result_structure(self, sample_docx):
        doc = parse(sample_docx)
        matches = search_content(doc, "test")
        if matches:
            m = matches[0]
            assert "match" in m
            assert "start" in m
            assert "end" in m
            assert "snippet" in m
            assert "type" in m
