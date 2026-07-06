"""OfficeLens CLI — command-line interface for Office document operations."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import click
from rich.console import Console
from rich.table import Table as RichTable

from . import __version__
from .analyzers import analyze, search_content
from .converters import convert_file, to_csv, to_markdown
from .parsers import DocType, parse

console = Console()


@click.group()
@click.version_option(version=__version__, prog_name="officelens")
def main():
    """OfficeLens — Office Document Analysis & Conversion Toolkit.

    Parse, extract, convert, and analyze DOCX, XLSX, and PPTX files.
    """


@main.command()
@click.argument("file", type=click.Path(exists=True))
@click.option(
    "--format", "-f", "fmt", type=click.Choice(["text", "markdown", "json"]), default="text", help="Output format"
)
@click.option("--output", "-o", type=click.Path(), default=None, help="Write output to file")
def extract(file: str, fmt: str, output: str | None):
    """Extract text content from an Office document."""
    try:
        doc = parse(file)
        if fmt == "text":
            content = doc.plain_text()
        elif fmt == "markdown":
            content = to_markdown(doc)
        elif fmt == "json":
            content = json.dumps(
                {
                    "path": doc.path,
                    "type": doc.doc_type.value,
                    "title": doc.title,
                    "text": doc.plain_text(),
                    "word_count": doc.word_count,
                },
                indent=2,
            )
        else:
            content = doc.plain_text()

        if output:
            Path(output).write_text(content, encoding="utf-8")
            console.print(f"[green]✓[/] Written to {output}")
        else:
            click.echo(content)
    except Exception as e:
        console.print(f"[red]Error:[/] {e}")
        sys.exit(1)


@main.command()
@click.argument("file", type=click.Path(exists=True))
@click.option(
    "--format", "-f", "fmt", type=click.Choice(["markdown", "csv", "json", "text"]), required=True, help="Output format"
)
@click.option("--output", "-o", type=click.Path(), required=True, help="Output file path")
@click.option("--sheet", "-s", default=None, help="Sheet name (XLSX only)")
def convert(file: str, fmt: str, output: str, sheet: str | None):
    """Convert an Office document to another format."""
    try:
        result = convert_file(file, output, fmt)
        console.print(f"[green]✓[/] Converted to {result}")
    except Exception as e:
        console.print(f"[red]Error:[/] {e}")
        sys.exit(1)


@main.command()
@click.argument("file", type=click.Path(exists=True))
@click.option("--json-output", "-j", is_flag=True, help="Output as JSON")
def info(file: str, json_output: bool):
    """Show detailed document statistics and information."""
    try:
        doc = parse(file)
        stats = analyze(doc)

        if json_output:
            data = {
                "path": doc.path,
                "doc_type": stats.doc_type,
                "title": stats.title,
                "author": doc.author,
                "word_count": stats.word_count,
                "char_count": stats.char_count,
                "char_count_no_spaces": stats.char_count_no_spaces,
                "paragraph_count": stats.paragraph_count,
                "sentence_count": stats.sentence_count,
                "heading_count": stats.heading_count,
                "table_count": stats.table_count,
                "image_count": stats.image_count,
                "page_estimate": stats.page_estimate,
                "avg_words_per_paragraph": stats.avg_words_per_paragraph,
                "avg_chars_per_word": stats.avg_chars_per_word,
                "readability_score": stats.readability_score,
                "readability_grade": stats.readability_grade,
                "unique_words": stats.unique_words,
                "vocabulary_richness": stats.vocabulary_richness,
                "top_words": stats.top_words,
                "metadata": doc.metadata,
            }
            click.echo(json.dumps(data, indent=2))
            return

        # Rich output
        table = RichTable(title=f"Document Info: {stats.title}", show_header=False, show_lines=True)
        table.add_column("Property", style="cyan", width=30)
        table.add_column("Value", style="white")
        table.add_row("File", doc.path)
        table.add_row("Type", stats.doc_type.upper())
        table.add_row("Author", doc.author or "(unknown)")
        table.add_row("Words", f"{stats.word_count:,}")
        table.add_row("Characters", f"{stats.char_count:,}")
        table.add_row("Characters (no spaces)", f"{stats.char_count_no_spaces:,}")
        table.add_row("Paragraphs", str(stats.paragraph_count))
        table.add_row("Sentences", str(stats.sentence_count))
        table.add_row("Headings", str(stats.heading_count))
        table.add_row("Tables", str(stats.table_count))
        table.add_row("Images", str(stats.image_count))
        table.add_row("Estimated Pages", str(stats.page_estimate))
        table.add_row("Avg Words/Paragraph", str(stats.avg_words_per_paragraph))
        table.add_row("Avg Chars/Word", str(stats.avg_chars_per_word))
        table.add_row("Readability Score", str(stats.readability_score))
        table.add_row("Readability Grade", stats.readability_grade)
        table.add_row("Unique Words", str(stats.unique_words))
        table.add_row("Vocabulary Richness", f"{stats.vocabulary_richness:.3f}")
        console.print(table)

        # Top words
        if stats.top_words:
            console.print("\n[bold]Top 10 Words:[/]")
            for word, count in stats.top_words[:10]:
                console.print(f"  {word}: {count}")

        # Headings
        if stats.heading_structure:
            console.print("\n[bold]Document Structure:[/]")
            for level, text in stats.heading_structure[:20]:
                indent = "  " * (level - 1)
                console.print(f"  {indent}{'#' * level} {text[:80]}")

        # Metadata
        if doc.metadata:
            console.print("\n[bold]Metadata:[/]")
            for key, val in doc.metadata.items():
                console.print(f"  {key}: {val}")

    except Exception as e:
        console.print(f"[red]Error:[/] {e}")
        sys.exit(1)


@main.command()
@click.argument("file", type=click.Path(exists=True))
@click.argument("query")
@click.option("--case-sensitive", "-c", is_flag=True, help="Case-sensitive search")
@click.option("--json-output", "-j", is_flag=True, help="Output as JSON")
def search(file: str, query: str, case_sensitive: bool, json_output: bool):
    """Search for content within an Office document."""
    try:
        doc = parse(file)
        matches = search_content(doc, query, case_sensitive=case_sensitive)

        if not matches:
            console.print(f"[yellow]No matches found for '{query}'[/]")
            return

        if json_output:
            click.echo(json.dumps(matches, indent=2))
            return

        console.print(f"[green]Found {len(matches)} match(es) for '{query}':[/]\n")
        for i, m in enumerate(matches, 1):
            loc = ""
            if "sheet" in m:
                loc = f"Sheet '{m['sheet']}' R{m['row']}C{m['col']}"
            elif "slide" in m:
                loc = f"Slide {m['slide']}"
            elif "index" in m:
                loc = f"Paragraph {m['index']}"

            console.print(f"  [cyan]{i}.[/] [{m.get('type', '?')}] {loc}")
            console.print(f"     {m['snippet']}")
            console.print()

    except Exception as e:
        console.print(f"[red]Error:[/] {e}")
        sys.exit(1)


@main.command()
@click.argument("files", nargs=-1, type=click.Path(exists=True), required=True)
@click.option(
    "--format",
    "-f",
    "fmt",
    type=click.Choice(["text", "markdown", "json", "csv"]),
    default="text",
    help="Output format",
)
@click.option("--output", "-o", type=click.Path(), default=None, help="Output directory")
def batch(files: tuple[str, ...], fmt: str, output: str | None):
    """Process multiple Office documents at once."""
    output_dir = Path(output) if output else None
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict] = []
    errors: list[dict] = []

    for file in files:
        try:
            doc = parse(file)
            if fmt == "text":
                content = doc.plain_text()
            elif fmt == "markdown":
                content = to_markdown(doc)
            elif fmt == "csv":
                content = to_csv(doc)
            elif fmt == "json":
                stats = analyze(doc)
                content = json.dumps(
                    {
                        "path": doc.path,
                        "type": doc.doc_type.value,
                        "title": stats.title,
                        "word_count": stats.word_count,
                        "table_count": stats.table_count,
                        "readability_score": stats.readability_score,
                    },
                    indent=2,
                )
            else:
                content = doc.plain_text()

            if output_dir:
                ext_map = {"text": ".txt", "markdown": ".md", "csv": ".csv", "json": ".json"}
                out_path = output_dir / (Path(file).stem + ext_map.get(fmt, ".txt"))
                out_path.write_text(content, encoding="utf-8")
                results.append({"file": file, "output": str(out_path), "status": "ok"})
            else:
                console.print(f"\n[bold cyan]=== {file} ===[/]")
                console.print(content[:2000])
                if len(content) > 2000:
                    console.print(f"\n[yellow]... ({len(content) - 2000} chars truncated)[/]")

        except Exception as e:
            errors.append({"file": file, "error": str(e)})
            console.print(f"[red]Error processing {file}:[/] {e}")

    if output_dir:
        console.print(f"\n[green]✓[/] Processed {len(results)} file(s) to {output_dir}")
        if errors:
            console.print(f"[red]✗[/] {len(errors)} file(s) failed")


@main.command()
@click.argument("file", type=click.Path(exists=True))
def tables(file: str):
    """Extract tables from an Office document."""
    try:
        doc = parse(file)
        table_list = []
        if doc.doc_type == DocType.DOCX:
            table_list = doc.tables
        elif doc.doc_type == DocType.PPTX:
            for slide in doc.slides:
                table_list.extend(slide.tables)

        if not table_list:
            console.print("[yellow]No tables found in document[/]")
            return

        for i, table in enumerate(table_list, 1):
            console.print(f"\n[bold cyan]Table {i}:[/] ({table.rows} rows × {table.cols} cols)")
            grid = table.to_dict()
            if grid:
                rich_table = RichTable(show_header=True, show_lines=True)
                if grid[0]:
                    for cell in grid[0]:
                        rich_table.add_column(str(cell), style="cyan")
                    for row in grid[1:]:
                        rich_table.add_row(*[str(c) for c in row])
                console.print(rich_table)

    except Exception as e:
        console.print(f"[red]Error:[/] {e}")
        sys.exit(1)


@main.command()
@click.argument("file", type=click.Path(exists=True))
def sheets(file: str):
    """List sheets and structure of an XLSX workbook."""
    try:
        doc = parse(file)
        if doc.doc_type != DocType.XLSX:
            console.print("[yellow]This command only works with XLSX files[/]")
            return

        if not doc.worksheets:
            console.print("[yellow]No worksheets found[/]")
            return

        for ws in doc.worksheets:
            console.print(f"\n[bold cyan]Sheet: {ws.name}[/]")
            console.print(f"  Rows: {ws.rows}, Columns: {ws.cols}")
            if ws.headers:
                console.print(f"  Headers: {', '.join(ws.headers[:10])}")
                if len(ws.headers) > 10:
                    console.print(f"           ... and {len(ws.headers) - 10} more")
            console.print(f"  Data rows: {len(ws.data) - 1 if ws.data else 0}")

    except Exception as e:
        console.print(f"[red]Error:[/] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
