"""Command-line interface for python-docx-redline.

Provides commands for editing Word documents with tracked changes from the terminal.
"""

from pathlib import Path
from typing import Annotated

import typer

from . import Document, __version__

app = typer.Typer(
    name="docx-redline",
    help="Edit Word documents with tracked changes from the command line.",
    no_args_is_help=True,
)


def version_callback(value: bool) -> None:
    """Print version and exit."""
    if value:
        typer.echo(f"docx-redline version {__version__}")
        raise typer.Exit()


@app.callback()
def main(
    version: Annotated[
        bool | None,
        typer.Option(
            "--version",
            "-v",
            help="Show version and exit.",
            callback=version_callback,
            is_eager=True,
        ),
    ] = None,
) -> None:
    """Edit Word documents with tracked changes from the command line."""
    pass


@app.command()
def insert(
    file: Annotated[Path, typer.Argument(help="Path to the .docx file")],
    text: Annotated[str, typer.Option("--text", "-t", help="Text to insert")],
    after: Annotated[
        str | None, typer.Option("--after", "-a", help="Insert after this text")
    ] = None,
    before: Annotated[
        str | None, typer.Option("--before", "-b", help="Insert before this text")
    ] = None,
    author: Annotated[
        str | None, typer.Option("--author", help="Author name for tracked change")
    ] = None,
    output: Annotated[Path | None, typer.Option("--output", "-o", help="Output file path")] = None,
    scope: Annotated[
        str | None,
        typer.Option("--scope", "-s", help="Limit to scope (e.g., 'section:Name')"),
    ] = None,
) -> None:
    """Insert text with tracked changes."""
    if not after and not before:
        typer.echo("Error: Must specify either --after or --before", err=True)
        raise typer.Exit(1)
    if after and before:
        typer.echo("Error: Cannot specify both --after and --before", err=True)
        raise typer.Exit(1)

    try:
        doc = Document(str(file), author=author or "CLI User")
        doc.insert_tracked(text, after=after, before=before, scope=scope)
        output_path = output or file
        doc.save(str(output_path))
        typer.echo(f"Inserted text and saved to {output_path}")
    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(1)


@app.command()
def replace(
    file: Annotated[Path, typer.Argument(help="Path to the .docx file")],
    find: Annotated[str, typer.Option("--find", "-f", help="Text to find")],
    replacement: Annotated[str, typer.Option("--replace", "-r", help="Replacement text")],
    author: Annotated[
        str | None, typer.Option("--author", help="Author name for tracked change")
    ] = None,
    output: Annotated[Path | None, typer.Option("--output", "-o", help="Output file path")] = None,
    scope: Annotated[
        str | None,
        typer.Option("--scope", "-s", help="Limit to scope (e.g., 'section:Name')"),
    ] = None,
    regex: Annotated[bool, typer.Option("--regex", help="Treat find as regex pattern")] = False,
) -> None:
    """Replace text with tracked changes."""
    try:
        doc = Document(str(file), author=author or "CLI User")
        doc.replace_tracked(find, replacement, scope=scope, regex=regex)
        output_path = output or file
        doc.save(str(output_path))
        typer.echo(f"Replaced text and saved to {output_path}")
    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(1)


@app.command()
def delete(
    file: Annotated[Path, typer.Argument(help="Path to the .docx file")],
    text: Annotated[str, typer.Option("--text", "-t", help="Text to delete")],
    author: Annotated[
        str | None, typer.Option("--author", help="Author name for tracked change")
    ] = None,
    output: Annotated[Path | None, typer.Option("--output", "-o", help="Output file path")] = None,
    scope: Annotated[
        str | None,
        typer.Option("--scope", "-s", help="Limit to scope (e.g., 'section:Name')"),
    ] = None,
    regex: Annotated[bool, typer.Option("--regex", help="Treat text as regex pattern")] = False,
) -> None:
    """Delete text with tracked changes."""
    try:
        doc = Document(str(file), author=author or "CLI User")
        doc.delete_tracked(text, scope=scope, regex=regex)
        output_path = output or file
        doc.save(str(output_path))
        typer.echo(f"Deleted text and saved to {output_path}")
    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(1)


@app.command()
def move(
    file: Annotated[Path, typer.Argument(help="Path to the .docx file")],
    text: Annotated[str, typer.Option("--text", "-t", help="Text to move")],
    after: Annotated[
        str | None, typer.Option("--after", "-a", help="Move to after this text")
    ] = None,
    before: Annotated[
        str | None, typer.Option("--before", "-b", help="Move to before this text")
    ] = None,
    author: Annotated[
        str | None, typer.Option("--author", help="Author name for tracked change")
    ] = None,
    output: Annotated[Path | None, typer.Option("--output", "-o", help="Output file path")] = None,
) -> None:
    """Move text to a new location with tracked changes."""
    if not after and not before:
        typer.echo("Error: Must specify either --after or --before", err=True)
        raise typer.Exit(1)
    if after and before:
        typer.echo("Error: Cannot specify both --after and --before", err=True)
        raise typer.Exit(1)

    try:
        doc = Document(str(file), author=author or "CLI User")
        doc.move_tracked(text, after=after, before=before)
        output_path = output or file
        doc.save(str(output_path))
        typer.echo(f"Moved text and saved to {output_path}")
    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(1)


@app.command("accept-all")
def accept_all(
    file: Annotated[Path, typer.Argument(help="Path to the .docx file")],
    output: Annotated[Path | None, typer.Option("--output", "-o", help="Output file path")] = None,
) -> None:
    """Accept all tracked changes in the document."""
    try:
        doc = Document(str(file))
        doc.accept_all_changes()
        output_path = output or file
        doc.save(str(output_path))
        typer.echo(f"Accepted all changes and saved to {output_path}")
    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(1)


@app.command()
def apply(
    file: Annotated[Path, typer.Argument(help="Path to the .docx file")],
    edits: Annotated[Path, typer.Argument(help="Path to YAML/JSON edits file")],
    output: Annotated[Path | None, typer.Option("--output", "-o", help="Output file path")] = None,
    author: Annotated[
        str | None, typer.Option("--author", help="Default author for changes")
    ] = None,
) -> None:
    """Apply edits from a YAML or JSON file."""
    try:
        doc = Document(str(file), author=author or "CLI User")
        results = doc.apply_edit_file(str(edits))
        output_path = output or file
        doc.save(str(output_path))

        success_count = sum(1 for r in results if r.success)
        fail_count = len(results) - success_count
        typer.echo(f"Applied {success_count} edits ({fail_count} failed), saved to {output_path}")

        if fail_count > 0:
            for r in results:
                if not r.success:
                    typer.echo(f"  Failed: {r.message}", err=True)
    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(1)


@app.command()
def info(
    file: Annotated[Path, typer.Argument(help="Path to the .docx file")],
) -> None:
    """Show document information."""
    try:
        doc = Document(str(file))
        typer.echo(f"File: {file}")
        typer.echo(f"Paragraphs: {len(doc.paragraphs)}")
        typer.echo(f"Sections: {len(doc.sections)}")
        typer.echo(f"Tables: {len(doc.tables)}")
        typer.echo(f"Comments: {len(doc.get_comments())}")
        typer.echo(f"Has tracked changes: {doc.has_tracked_changes}")
    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    app()
