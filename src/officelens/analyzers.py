"""Document analyzers — statistics, content analysis, readability scoring."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

from .parsers import DocType, Document


@dataclass
class DocStats:
    """Comprehensive document statistics."""

    doc_type: str
    title: str
    word_count: int
    char_count: int
    char_count_no_spaces: int
    paragraph_count: int
    sentence_count: int
    heading_count: int
    table_count: int
    image_count: int
    page_estimate: int
    avg_words_per_paragraph: float
    avg_chars_per_word: float
    readability_score: float  # Flesch Reading Ease (0-100)
    readability_grade: str  # Flesch-Kincaid grade level
    unique_words: int
    vocabulary_richness: float  # type-token ratio
    top_words: list[tuple[str, int]]
    heading_structure: list[tuple[int, str]]


def _count_sentences(text: str) -> int:
    """Count sentences in text."""
    if not text.strip():
        return 0
    sentences = re.split(r"[.!?]+", text)
    return max(1, len([s for s in sentences if s.strip()]))


def _flesch_reading_ease(words: list[str], sentences: int, syllables: int) -> float:
    """Calculate Flesch Reading Ease score."""
    if not words or sentences == 0:
        return 0.0
    avg_words_per_sentence = len(words) / sentences
    avg_syllables_per_word = syllables / len(words)
    score = 206.835 - 1.015 * avg_words_per_sentence - 84.6 * avg_syllables_per_word
    return round(max(0, min(100, score)), 1)


def _flesch_kincaid_grade(words: list[str], sentences: int, syllables: int) -> str:
    """Calculate Flesch-Kincaid Grade Level."""
    if not words or sentences == 0:
        return "N/A"
    avg_words_per_sentence = len(words) / sentences
    avg_syllables_per_word = syllables / len(words)
    grade = 0.39 * avg_words_per_sentence + 11.8 * avg_syllables_per_word - 15.59
    grade = max(0, grade)
    if grade < 1:
        return "Kindergarten"
    elif grade <= 12:
        return f"Grade {int(round(grade))}"
    else:
        return f"College (Grade {int(round(grade))})"


def _count_syllables(word: str) -> int:
    """Estimate syllable count for English word."""
    word = word.lower().strip()
    if not word:
        return 0
    if len(word) <= 3:
        return 1

    # Remove trailing e
    if word.endswith("e"):
        word = word[:-1]
    if not word:
        return 1

    # Count vowel groups
    vowels = "aeiouy"
    count = 0
    prev_vowel = False
    for char in word:
        is_vowel = char in vowels
        if is_vowel and not prev_vowel:
            count += 1
        prev_vowel = is_vowel

    return max(1, count)


def analyze(doc: Document) -> DocStats:
    """Perform comprehensive document analysis."""
    # Collect all text
    all_paragraphs: list[str] = []
    heading_count = 0
    heading_structure: list[tuple[int, str]] = []

    if doc.doc_type == DocType.DOCX:
        for p in doc.paragraphs:
            if p.text.strip():
                all_paragraphs.append(p.text)
                if p.level > 0:
                    heading_count += 1
                    heading_structure.append((p.level, p.text))
    elif doc.doc_type == DocType.PPTX:
        for s in doc.slides:
            for p in s.paragraphs:
                if p.text.strip():
                    all_paragraphs.append(p.text)
    elif doc.doc_type == DocType.XLSX:
        for ws in doc.worksheets:
            for row in ws.data:
                for cell in row:
                    if cell is not None:
                        all_paragraphs.append(str(cell))

    full_text = " ".join(all_paragraphs)
    words = full_text.split()
    word_count = len(words)
    char_count = len(full_text)
    char_count_no_spaces = len(full_text.replace(" ", ""))

    # Sentences
    sentences = _count_sentences(full_text)

    # Tables
    if doc.doc_type == DocType.DOCX:
        table_count = len(doc.tables)
    elif doc.doc_type == DocType.PPTX:
        table_count = sum(len(s.tables) for s in doc.slides)
    else:
        table_count = 0

    # Syllables
    total_syllables = sum(_count_syllables(w) for w in words)

    # Readability
    readability_ease = _flesch_reading_ease(words, sentences, total_syllables)
    readability_grade = _flesch_kincaid_grade(words, sentences, total_syllables)

    # Vocabulary
    word_freq = Counter(w.lower().strip(".,!?;:\"'()-") for w in words if len(w) > 1)
    unique_words = len(word_freq)
    vocabulary_richness = unique_words / max(1, word_count)
    top_words = word_freq.most_common(20)

    # Averages
    para_count = max(1, len(all_paragraphs))
    avg_words_per_para = round(word_count / para_count, 1)
    avg_chars_per_word = round(char_count / max(1, word_count), 1)

    return DocStats(
        doc_type=doc.doc_type.value,
        title=doc.title or "(untitled)",
        word_count=word_count,
        char_count=char_count,
        char_count_no_spaces=char_count_no_spaces,
        paragraph_count=len(all_paragraphs),
        sentence_count=sentences,
        heading_count=heading_count,
        table_count=table_count,
        image_count=doc.image_count,
        page_estimate=doc.page_estimate if doc.page_estimate else max(1, (word_count + 249) // 250),
        avg_words_per_paragraph=avg_words_per_para,
        avg_chars_per_word=avg_chars_per_word,
        readability_score=readability_ease,
        readability_grade=readability_grade,
        unique_words=unique_words,
        vocabulary_richness=round(vocabulary_richness, 3),
        top_words=top_words,
        heading_structure=heading_structure,
    )


def search_content(doc: Document, query: str, case_sensitive: bool = False) -> list[dict]:
    """Search for content in a document. Returns list of matches."""
    matches: list[dict] = []
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(re.escape(query), flags)

    def _search_text(text: str, context: dict):
        for match in pattern.finditer(text):
            start = max(0, match.start() - 40)
            end = min(len(text), match.end() + 40)
            snippet = text[start:end]
            matches.append(
                {
                    "match": match.group(),
                    "start": match.start(),
                    "end": match.end(),
                    "snippet": f"...{snippet}...",
                    **context,
                }
            )

    if doc.doc_type == DocType.DOCX:
        for i, p in enumerate(doc.paragraphs):
            if p.text.strip():
                _search_text(p.text, {"type": "paragraph", "index": i, "style": p.style})
        # Also search table cells
        for ti, table in enumerate(doc.tables):
            for cell in table.cells:
                if cell.text.strip():
                    _search_text(cell.text, {"type": "cell", "table": ti + 1, "row": cell.row + 1, "col": cell.col + 1})

    elif doc.doc_type == DocType.XLSX:
        for ws in doc.worksheets:
            for ri, row in enumerate(ws.data):
                for ci, cell in enumerate(row):
                    if cell is not None:
                        _search_text(
                            str(cell),
                            {
                                "type": "cell",
                                "sheet": ws.name,
                                "row": ri + 1,
                                "col": ci + 1,
                            },
                        )

    elif doc.doc_type == DocType.PPTX:
        for s in doc.slides:
            for i, p in enumerate(s.paragraphs):
                _search_text(p.text, {"type": "paragraph", "slide": s.number, "index": i})
            # Also search table cells
            for ti, table in enumerate(s.tables):
                for cell in table.cells:
                    if cell.text.strip():
                        cell_meta = {
                            "type": "cell",
                            "slide": s.number,
                            "table": ti + 1,
                            "row": cell.row + 1,
                            "col": cell.col + 1,
                        }
                        _search_text(cell.text, cell_meta)

    return matches
