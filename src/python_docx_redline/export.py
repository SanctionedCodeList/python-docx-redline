"""
Export functionality for tracked changes in Word documents.

This module provides functions to export tracked changes to various formats
including HTML, Markdown, and JSON for visualization and integration with
external tools.
"""

from __future__ import annotations

import html
import json
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING, Any, Literal

from lxml import etree

from .constants import WORD_NAMESPACE

if TYPE_CHECKING:
    from .document import Document
    from .models.tracked_change import TrackedChange


@dataclass
class ChangeContext:
    """Context information for a tracked change.

    Provides surrounding text to help understand where a change was made.
    """

    before: str
    """Text appearing before the change."""

    after: str
    """Text appearing after the change."""

    paragraph_text: str
    """Full text of the paragraph containing the change."""

    paragraph_index: int
    """Zero-based index of the paragraph in the document."""


@dataclass
class ExportedChange:
    """A tracked change with metadata and optional context.

    This is a serializable representation of a TrackedChange suitable for export.
    """

    id: str
    """Unique change identifier."""

    change_type: str
    """Type of change: insertion, deletion, move_from, move_to, format_run, format_paragraph."""

    author: str
    """Author who made the change."""

    date: str | None
    """ISO 8601 formatted date string, or None if no date."""

    text: str
    """The text content of the change."""

    context: ChangeContext | None = None
    """Optional surrounding context."""


@dataclass
class ChangeReport:
    """A complete change report with summary statistics and grouped changes.

    Attributes:
        title: Report title
        generated_at: When the report was generated (ISO format)
        total_changes: Total number of changes
        insertions: Number of insertions
        deletions: Number of deletions
        moves: Number of move operations
        format_changes: Number of formatting changes
        changes: List of exported changes
        by_author: Changes grouped by author
        by_type: Changes grouped by type
    """

    title: str
    generated_at: str
    total_changes: int
    insertions: int
    deletions: int
    moves: int
    format_changes: int
    changes: list[ExportedChange] = field(default_factory=list)
    by_author: dict[str, list[ExportedChange]] = field(default_factory=dict)
    by_type: dict[str, list[ExportedChange]] = field(default_factory=dict)


def _find_paragraph_for_element(
    doc: Document, element: etree._Element
) -> tuple[etree._Element | None, int]:
    """Find the paragraph containing an element.

    Args:
        doc: The document
        element: The XML element to find

    Returns:
        Tuple of (paragraph element, paragraph index) or (None, -1) if not found
    """
    paragraphs = list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

    # Walk up from element to find its containing paragraph
    current = element
    while current is not None:
        if current.tag == f"{{{WORD_NAMESPACE}}}p":
            # Found the paragraph, get its index
            try:
                idx = paragraphs.index(current)
                return current, idx
            except ValueError:
                return None, -1
        current = current.getparent()

    return None, -1


def _extract_paragraph_text(para_element: etree._Element) -> str:
    """Extract text from a paragraph element.

    Includes both regular text (w:t) and deleted text (w:delText).
    """
    text_elements = para_element.findall(f".//{{{WORD_NAMESPACE}}}t")
    deltext_elements = para_element.findall(f".//{{{WORD_NAMESPACE}}}delText")
    return "".join(elem.text or "" for elem in text_elements + deltext_elements)


def _get_change_context(
    doc: Document,
    change: TrackedChange,
    context_chars: int = 50,
) -> ChangeContext | None:
    """Get the context around a tracked change.

    Args:
        doc: The document containing the change
        change: The tracked change
        context_chars: Number of characters of context to include

    Returns:
        ChangeContext with surrounding text, or None if context unavailable
    """
    para_element, para_idx = _find_paragraph_for_element(doc, change.element)
    if para_element is None:
        return None

    para_text = _extract_paragraph_text(para_element)
    change_text = change.text

    # Try to find the change text in the paragraph
    if change_text and change_text in para_text:
        pos = para_text.find(change_text)
        before_start = max(0, pos - context_chars)
        after_end = min(len(para_text), pos + len(change_text) + context_chars)

        before = para_text[before_start:pos]
        after = para_text[pos + len(change_text) : after_end]
    else:
        # Can't find exact text, just return paragraph info
        before = para_text[:context_chars] if len(para_text) > context_chars else para_text
        after = ""

    return ChangeContext(
        before=before,
        after=after,
        paragraph_text=para_text,
        paragraph_index=para_idx,
    )


def _export_change(
    doc: Document,
    change: TrackedChange,
    include_context: bool = False,
    context_chars: int = 50,
) -> ExportedChange:
    """Convert a TrackedChange to an ExportedChange.

    Args:
        doc: The document
        change: The tracked change to export
        include_context: Whether to include surrounding context
        context_chars: Number of context characters to include

    Returns:
        ExportedChange instance
    """
    context = None
    if include_context:
        context = _get_change_context(doc, change, context_chars)

    return ExportedChange(
        id=change.id,
        change_type=change.change_type.value,
        author=change.author,
        date=change.date.isoformat() if change.date else None,
        text=change.text,
        context=context,
    )


def export_changes_json(
    doc: Document,
    include_context: bool = True,
    context_chars: int = 50,
    indent: int | None = 2,
) -> str:
    """Export all tracked changes to JSON format.

    Args:
        doc: The document to export changes from
        include_context: Whether to include surrounding text context
        context_chars: Number of context characters to include on each side
        indent: JSON indentation level, or None for compact output

    Returns:
        JSON string containing all tracked changes

    Example:
        >>> json_data = doc.export_changes_json()
        >>> changes = json.loads(json_data)
        >>> print(f"Found {len(changes['changes'])} changes")
    """
    changes = doc.get_tracked_changes()
    exported = [_export_change(doc, c, include_context, context_chars) for c in changes]

    # Convert to serializable dicts
    result = {
        "total_changes": len(exported),
        "changes": [_change_to_dict(c) for c in exported],
    }

    return json.dumps(result, indent=indent, ensure_ascii=False)


def _change_to_dict(change: ExportedChange) -> dict[str, Any]:
    """Convert an ExportedChange to a dictionary for JSON serialization."""
    result: dict[str, Any] = {
        "id": change.id,
        "change_type": change.change_type,
        "author": change.author,
        "date": change.date,
        "text": change.text,
    }
    if change.context:
        result["context"] = {
            "before": change.context.before,
            "after": change.context.after,
            "paragraph_text": change.context.paragraph_text,
            "paragraph_index": change.context.paragraph_index,
        }
    return result


def export_changes_markdown(
    doc: Document,
    include_context: bool = True,
    context_chars: int = 50,
    group_by: Literal["none", "author", "type"] | None = None,
) -> str:
    """Export tracked changes to Markdown format.

    Creates a human-readable Markdown document showing all tracked changes
    with optional context and grouping.

    Args:
        doc: The document to export changes from
        include_context: Whether to include surrounding text context
        context_chars: Number of context characters to include on each side
        group_by: How to group changes: "author", "type", or "none"/None

    Returns:
        Markdown formatted string with all tracked changes

    Example:
        >>> md = doc.export_changes_markdown(group_by="author")
        >>> print(md)
    """
    changes = doc.get_tracked_changes()
    exported = [_export_change(doc, c, include_context, context_chars) for c in changes]

    lines = ["# Tracked Changes Report", ""]

    # Summary
    stats = doc.comparison_stats
    lines.extend(
        [
            "## Summary",
            "",
            f"- **Total changes**: {stats.total}",
            f"- **Insertions**: {stats.insertions}",
            f"- **Deletions**: {stats.deletions}",
        ]
    )
    if stats.moves:
        lines.append(f"- **Moves**: {stats.moves}")
    if stats.format_changes:
        lines.append(f"- **Format changes**: {stats.format_changes}")
    lines.append("")

    if not exported:
        lines.append("*No tracked changes found.*")
        return "\n".join(lines)

    # Group changes if requested
    if group_by == "author":
        lines.append("## Changes by Author")
        lines.append("")
        by_author: dict[str, list[ExportedChange]] = defaultdict(list)
        for c in exported:
            by_author[c.author or "Unknown"].append(c)

        for author in sorted(by_author.keys()):
            lines.append(f"### {author}")
            lines.append("")
            for c in by_author[author]:
                lines.extend(_format_change_markdown(c, include_context))
            lines.append("")

    elif group_by == "type":
        lines.append("## Changes by Type")
        lines.append("")
        by_type: dict[str, list[ExportedChange]] = defaultdict(list)
        for c in exported:
            by_type[c.change_type].append(c)

        type_order = [
            "insertion",
            "deletion",
            "move_from",
            "move_to",
            "format_run",
            "format_paragraph",
        ]
        for change_type in type_order:
            if change_type in by_type:
                lines.append(f"### {_format_type_header(change_type)}")
                lines.append("")
                for c in by_type[change_type]:
                    lines.extend(_format_change_markdown(c, include_context))
                lines.append("")

    else:
        lines.append("## All Changes")
        lines.append("")
        for c in exported:
            lines.extend(_format_change_markdown(c, include_context))

    return "\n".join(lines)


def _format_type_header(change_type: str) -> str:
    """Format a change type as a readable header."""
    return {
        "insertion": "Insertions",
        "deletion": "Deletions",
        "move_from": "Move Sources",
        "move_to": "Move Destinations",
        "format_run": "Run Format Changes",
        "format_paragraph": "Paragraph Format Changes",
    }.get(change_type, change_type.replace("_", " ").title())


def _format_change_markdown(
    change: ExportedChange,
    include_context: bool,
) -> list[str]:
    """Format a single change as Markdown lines."""
    lines = []

    # Change icon based on type
    icon = {
        "insertion": "+",
        "deletion": "-",
        "move_from": "<<",
        "move_to": ">>",
        "format_run": "~",
        "format_paragraph": "~",
    }.get(change.change_type, "*")

    # Header line with metadata
    date_str = f" ({change.date[:10]})" if change.date else ""
    author_str = f" by {change.author}" if change.author else ""
    change_type_str = change.change_type.replace("_", " ").title()
    lines.append(f"- **[{icon}]** {change_type_str}{author_str}{date_str}")

    # Text content
    if change.text:
        text_preview = change.text[:100]
        if len(change.text) > 100:
            text_preview += "..."
        # Use inline code for the changed text
        lines.append(f"  - Text: `{text_preview}`")

    # Context if included
    if include_context and change.context:
        ctx = change.context
        if ctx.before or ctx.after:
            context_line = f"  - Context: ...{ctx.before}**[change]**{ctx.after}..."
            lines.append(context_line)

    lines.append("")
    return lines


def export_changes_html(
    doc: Document,
    include_context: bool = True,
    context_chars: int = 50,
    group_by: Literal["none", "author", "type"] | None = None,
    inline_styles: bool = True,
) -> str:
    """Export tracked changes to HTML format.

    Creates an HTML document with a code-review style visualization of
    tracked changes, similar to diff views in version control systems.

    Args:
        doc: The document to export changes from
        include_context: Whether to include surrounding text context
        context_chars: Number of context characters to include on each side
        group_by: How to group changes: "author", "type", or "none"/None
        inline_styles: Whether to include inline CSS styles (True) or just classes (False)

    Returns:
        HTML formatted string with all tracked changes

    Example:
        >>> html_content = doc.export_changes_html(group_by="author")
        >>> with open("changes.html", "w") as f:
        ...     f.write(html_content)
    """
    changes = doc.get_tracked_changes()
    exported = [_export_change(doc, c, include_context, context_chars) for c in changes]

    parts = []

    # HTML header with styles
    parts.append(_get_html_header(inline_styles))

    parts.append('<div class="change-report">')
    parts.append("<h1>Tracked Changes Report</h1>")

    # Summary section
    stats = doc.comparison_stats
    parts.append('<div class="summary">')
    parts.append("<h2>Summary</h2>")
    parts.append("<ul>")
    parts.append(f"<li><strong>Total changes:</strong> {stats.total}</li>")
    parts.append(f"<li><strong>Insertions:</strong> {stats.insertions}</li>")
    parts.append(f"<li><strong>Deletions:</strong> {stats.deletions}</li>")
    if stats.moves:
        parts.append(f"<li><strong>Moves:</strong> {stats.moves}</li>")
    if stats.format_changes:
        parts.append(f"<li><strong>Format changes:</strong> {stats.format_changes}</li>")
    parts.append("</ul>")
    parts.append("</div>")

    if not exported:
        parts.append('<p class="no-changes">No tracked changes found.</p>')
    elif group_by == "author":
        parts.append('<div class="changes-by-author">')
        parts.append("<h2>Changes by Author</h2>")

        by_author: dict[str, list[ExportedChange]] = defaultdict(list)
        for c in exported:
            by_author[c.author or "Unknown"].append(c)

        for author in sorted(by_author.keys()):
            parts.append('<div class="author-section">')
            parts.append(f"<h3>{html.escape(author)}</h3>")
            parts.append('<div class="change-list">')
            for c in by_author[author]:
                parts.append(_format_change_html(c, include_context, inline_styles))
            parts.append("</div>")
            parts.append("</div>")

        parts.append("</div>")

    elif group_by == "type":
        parts.append('<div class="changes-by-type">')
        parts.append("<h2>Changes by Type</h2>")

        by_type: dict[str, list[ExportedChange]] = defaultdict(list)
        for c in exported:
            by_type[c.change_type].append(c)

        type_order = [
            "insertion",
            "deletion",
            "move_from",
            "move_to",
            "format_run",
            "format_paragraph",
        ]
        for change_type in type_order:
            if change_type in by_type:
                parts.append(f'<div class="type-section type-{change_type}">')
                parts.append(f"<h3>{html.escape(_format_type_header(change_type))}</h3>")
                parts.append('<div class="change-list">')
                for c in by_type[change_type]:
                    parts.append(_format_change_html(c, include_context, inline_styles))
                parts.append("</div>")
                parts.append("</div>")

        parts.append("</div>")

    else:
        parts.append('<div class="all-changes">')
        parts.append("<h2>All Changes</h2>")
        parts.append('<div class="change-list">')
        for c in exported:
            parts.append(_format_change_html(c, include_context, inline_styles))
        parts.append("</div>")
        parts.append("</div>")

    parts.append("</div>")
    parts.append("</body>")
    parts.append("</html>")

    return "\n".join(parts)


def _get_html_header(inline_styles: bool) -> str:
    """Get the HTML header with optional inline styles."""
    # fmt: off
    styles = """
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto,
                         Oxygen, Ubuntu, sans-serif;
            line-height: 1.6;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background: #f5f5f5;
        }
        .change-report {
            background: white;
            border-radius: 8px;
            padding: 24px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 { color: #333; border-bottom: 2px solid #e1e4e8; padding-bottom: 10px; }
        h2 { color: #444; margin-top: 24px; }
        h3 { color: #555; margin-top: 16px; }
        .summary {
            background: #f6f8fa; padding: 16px;
            border-radius: 6px; margin: 16px 0;
        }
        .summary ul { margin: 0; }
        .change-item {
            border: 1px solid #e1e4e8;
            border-radius: 6px;
            margin: 12px 0;
            overflow: hidden;
        }
        .change-header {
            padding: 12px 16px;
            font-weight: 500;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .change-body {
            padding: 12px 16px;
            background: #fafbfc;
            border-top: 1px solid #e1e4e8;
        }
        .change-text {
            font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', monospace;
            font-size: 13px;
            background: #f6f8fa;
            padding: 8px 12px;
            border-radius: 4px;
            white-space: pre-wrap;
            word-break: break-word;
        }
        .change-context { color: #666; font-size: 13px; margin-top: 8px; }
        .context-before, .context-after { color: #666; }
        .context-change { font-weight: bold; }
        .badge {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
        }
        .badge-insertion { background: #dcffe4; color: #22863a; }
        .badge-deletion { background: #ffeef0; color: #cb2431; }
        .badge-move_from, .badge-move_to { background: #fff5b1; color: #735c0f; }
        .badge-format_run, .badge-format_paragraph { background: #f1f8ff; color: #0366d6; }
        .insertion .change-header {
            background: #f0fff4; border-left: 4px solid #22863a;
        }
        .deletion .change-header {
            background: #ffeef0; border-left: 4px solid #cb2431;
        }
        .move_from .change-header, .move_to .change-header {
            background: #fffbdd; border-left: 4px solid #735c0f;
        }
        .format_run .change-header, .format_paragraph .change-header {
            background: #f1f8ff; border-left: 4px solid #0366d6;
        }
        .author { color: #586069; }
        .date { color: #959da5; font-size: 12px; }
        .no-changes { color: #666; font-style: italic; }
    </style>
    """ if inline_styles else ""
    # fmt: on

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tracked Changes Report</title>
    {styles}
</head>
<body>"""


def _format_change_html(
    change: ExportedChange,
    include_context: bool,
    inline_styles: bool,
) -> str:
    """Format a single change as HTML."""
    parts = []

    type_class = change.change_type
    parts.append(f'<div class="change-item {type_class}">')

    # Header
    parts.append('<div class="change-header">')
    badge_text = change.change_type.replace("_", " ").upper()
    parts.append(f'<span class="badge badge-{type_class}">{badge_text}</span>')
    if change.author:
        parts.append(f'<span class="author">by {html.escape(change.author)}</span>')
    if change.date:
        parts.append(f'<span class="date">{html.escape(change.date[:10])}</span>')
    parts.append("</div>")

    # Body
    parts.append('<div class="change-body">')
    if change.text:
        text_preview = change.text[:200]
        if len(change.text) > 200:
            text_preview += "..."
        parts.append(f'<div class="change-text">{html.escape(text_preview)}</div>')

    if include_context and change.context:
        ctx = change.context
        if ctx.before or ctx.after:
            parts.append('<div class="change-context">')
            parts.append('<span class="context-before">...' + html.escape(ctx.before) + "</span>")
            parts.append('<span class="context-change">[change]</span>')
            parts.append('<span class="context-after">' + html.escape(ctx.after) + "...</span>")
            parts.append("</div>")

    parts.append("</div>")
    parts.append("</div>")

    return "\n".join(parts)


def generate_change_report(
    doc: Document,
    format: Literal["html", "markdown", "json"] = "html",
    include_context: bool = True,
    context_chars: int = 50,
    group_by: Literal["none", "author", "type"] | None = "author",
    title: str | None = None,
) -> str:
    """Generate a comprehensive change report in the specified format.

    This is a convenience function that wraps the individual export functions
    with sensible defaults for generating reports.

    Args:
        doc: The document to generate a report for
        format: Output format: "html", "markdown", or "json"
        include_context: Whether to include surrounding text context
        context_chars: Number of context characters to include on each side
        group_by: How to group changes: "author", "type", or "none"/None
        title: Optional custom title for the report

    Returns:
        Formatted report string in the specified format

    Example:
        >>> report = doc.generate_change_report(format="html", group_by="author")
        >>> with open("report.html", "w") as f:
        ...     f.write(report)
    """
    if format == "json":
        # For JSON, return a more comprehensive report structure
        changes = doc.get_tracked_changes()
        exported = [_export_change(doc, c, include_context, context_chars) for c in changes]

        stats = doc.comparison_stats

        # Group changes
        by_author: dict[str, list[dict]] = defaultdict(list)
        by_type: dict[str, list[dict]] = defaultdict(list)
        changes_list = []

        for c in exported:
            change_dict = _change_to_dict(c)
            changes_list.append(change_dict)
            by_author[c.author or "Unknown"].append(change_dict)
            by_type[c.change_type].append(change_dict)

        report = {
            "title": title or "Tracked Changes Report",
            "generated_at": datetime.now().isoformat(),
            "summary": {
                "total_changes": stats.total,
                "insertions": stats.insertions,
                "deletions": stats.deletions,
                "moves": stats.moves,
                "format_changes": stats.format_changes,
            },
            "changes": changes_list,
            "by_author": dict(by_author),
            "by_type": dict(by_type),
        }

        return json.dumps(report, indent=2, ensure_ascii=False)

    elif format == "markdown":
        return export_changes_markdown(
            doc,
            include_context=include_context,
            context_chars=context_chars,
            group_by=group_by,
        )

    else:  # html
        return export_changes_html(
            doc,
            include_context=include_context,
            context_chars=context_chars,
            group_by=group_by,
        )
