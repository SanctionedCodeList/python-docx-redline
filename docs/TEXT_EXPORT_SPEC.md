# Text and Markdown Export Specification

> Native text/markdown export from AccessibilityTree, replacing pandoc dependency

**Version**: 1.0.0
**Date**: 2025-12-29
**Status**: Ready for Implementation

---

## Executive Summary

This specification defines `to_text()` and `to_markdown()` methods for AccessibilityTree that provide complete document export without external dependencies. This replaces pandoc for DOCX reading workflows while offering features pandoc lacks (comments, better tracked changes handling).

---

## 1. API Design

### 1.1 Simple Usage

```python
from python_docx_redline import Document
from python_docx_redline.accessibility import AccessibilityTree

doc = Document("contract.docx")
tree = AccessibilityTree.from_document(doc)

# Plain text (tracked changes accepted by default)
text = tree.to_text()

# Markdown format
md = tree.to_markdown()
```

### 1.2 Configuration Object

```python
@dataclass
class TextExportConfig:
    """Configuration for text/markdown export."""

    # === Content Inclusion ===
    include_headers: bool = True
    include_footers: bool = True
    include_footnotes: bool = True
    include_endnotes: bool = True
    include_comments: bool = False
    include_images: bool = True  # As placeholders

    # === Tracked Changes ===
    # "accept" - Show final state (insertions visible, deletions hidden)
    # "reject" - Show original state (deletions visible, insertions hidden)
    # "all"    - Show both with CriticMarkup syntax {++ins++} {--del--}
    tracked_changes: Literal["accept", "reject", "all"] = "accept"

    # === Formatting ===
    # Table format for markdown output
    table_format: Literal["markdown", "simple", "grid"] = "markdown"

    # Line width for text wrapping (0 = no wrap)
    line_width: int = 0

    # Heading style for plain text
    heading_style: Literal["underline", "prefix"] = "underline"
    # "underline": HEADING\n=======
    # "prefix":    # HEADING
```

### 1.3 Full Configuration Example

```python
config = TextExportConfig(
    include_headers=True,
    include_footers=True,
    include_footnotes=True,
    include_endnotes=True,
    include_comments=True,  # Include as [Comment: text]
    include_images=True,
    tracked_changes="all",  # CriticMarkup syntax
    table_format="markdown",
    line_width=80,
)

md = tree.to_markdown(config)
text = tree.to_text(config)
```

---

## 2. Output Formats

### 2.1 Markdown Output

```markdown
# SERVICES AGREEMENT

This Agreement ("Agreement") is entered into as of {++January 15++}{--January 1--}, 2024.

## 1. Definitions

The following terms have the meanings set forth below:

| Term | Definition | Section |
|------|------------|---------|
| Services | The consulting services described in Exhibit A | 2.1 |
| Client | ACME Corporation | 1.1 |
| Fees | $150,000 annually | 4.1 |

## 2. Services

The Contractor shall provide the Services[^1] in accordance with the terms herein.

> **Note**: All services are subject to the limitations in Section 5.

[Comment by Legal Team: Review indemnification language]

Payment is due within {--30--}{++45++} days of invoice.

[image: Figure 1 - Service Architecture]

See [Payment Terms](#payment-terms) for details.

---

**Header**: ACME Corp - Confidential

**Footer**: Page {PAGE} of {NUMPAGES}

---

[^1]: As defined in Section 1.
[^2]: Subject to annual adjustment per Section 4.3.
```

### 2.2 Plain Text Output

```
SERVICES AGREEMENT
==================

This Agreement ("Agreement") is entered into as of January 15, 2024.

1. Definitions
--------------

The following terms have the meanings set forth below:

Term       | Definition                                    | Section
-----------|-----------------------------------------------|--------
Services   | The consulting services described in Exhibit A | 2.1
Client     | ACME Corporation                              | 1.1
Fees       | $150,000 annually                             | 4.1

2. Services
-----------

The Contractor shall provide the Services [1] in accordance with the terms herein.

    Note: All services are subject to the limitations in Section 5.

[Comment by Legal Team: Review indemnification language]

Payment is due within 45 days of invoice.

[image: Figure 1 - Service Architecture]

See "Payment Terms" for details.

---

Header: ACME Corp - Confidential

Footer: Page {PAGE} of {NUMPAGES}

---

Footnotes:
[1] As defined in Section 1.
[2] Subject to annual adjustment per Section 4.3.
```

---

## 3. Element Rendering

### 3.1 Headings

**Markdown:**
```markdown
# Level 1 Heading
## Level 2 Heading
### Level 3 Heading
```

**Plain Text (underline style):**
```
Level 1 Heading
===============

Level 2 Heading
---------------

Level 3 Heading
~~~~~~~~~~~~~~~
```

**Plain Text (prefix style):**
```
# Level 1 Heading

## Level 2 Heading

### Level 3 Heading
```

### 3.2 Paragraphs

- Standard paragraphs separated by blank lines
- Preserve inline formatting cues where meaningful
- Handle smart quotes (convert to straight quotes in plain text)

### 3.3 Lists

**Bullet Lists (Markdown):**
```markdown
- First item
- Second item
  - Nested item
  - Another nested
- Third item
```

**Numbered Lists (Markdown):**
```markdown
1. First item
2. Second item
   1. Nested item
   2. Another nested
3. Third item
```

**Plain Text:** Same format, works in both.

### 3.4 Tables

**Markdown format:**
```markdown
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |
```

**Simple format:**
```
Header 1   | Header 2   | Header 3
-----------|------------|----------
Cell 1     | Cell 2     | Cell 3
Cell 4     | Cell 5     | Cell 6
```

**Grid format:**
```
+----------+----------+----------+
| Header 1 | Header 2 | Header 3 |
+==========+==========+==========+
| Cell 1   | Cell 2   | Cell 3   |
+----------+----------+----------+
| Cell 4   | Cell 5   | Cell 6   |
+----------+----------+----------+
```

### 3.5 Block Quotes

**Markdown:**
```markdown
> This is a block quote.
> It can span multiple lines.
```

**Plain Text:**
```
    This is a block quote.
    It can span multiple lines.
```

### 3.6 Hyperlinks

**Markdown:**
```markdown
See [Payment Terms](https://example.com/terms) for details.
Internal: See [Section 2](#section-2) below.
```

**Plain Text:**
```
See "Payment Terms" (https://example.com/terms) for details.
Internal: See "Section 2" below.
```

### 3.7 Images

**Markdown:**
```markdown
![Figure 1 - Architecture Diagram](image1.png)
```

Or as placeholder if image data not available:
```markdown
[image: Figure 1 - Architecture Diagram]
```

**Plain Text:**
```
[image: Figure 1 - Architecture Diagram]
```

### 3.8 Comments

**Markdown:**
```markdown
[Comment by John Smith: Please review this section]
```

**Plain Text:**
```
[Comment by John Smith: Please review this section]
```

Comments appear inline at their anchor point.

---

## 4. Tracked Changes Handling

### 4.1 Accept Mode (Default)

Shows the final state - insertions visible, deletions hidden.

**Input:** "Payment in ~~30~~ **45** days"
**Output:** "Payment in 45 days"

### 4.2 Reject Mode

Shows the original state - deletions visible, insertions hidden.

**Input:** "Payment in ~~30~~ **45** days"
**Output:** "Payment in 30 days"

### 4.3 All Mode (CriticMarkup)

Shows both states using CriticMarkup syntax, matching our existing `to_criticmarkup()` support.

**Output:**
```
Payment in {--30--}{++45++} days
```

CriticMarkup syntax reference:
- `{++insertion++}` - Added text
- `{--deletion--}` - Removed text
- `{~~old~>new~~}` - Substitution (optional, can also use del+ins)
- `{>>comment<<}` - Comments (if tracked_changes="all" and include_comments=True)
- `{==highlight==}` - Highlights (future)

---

## 5. Document Parts

### 5.1 Body Content

Always included. This is the main document content.

### 5.2 Headers and Footers

When `include_headers=True` or `include_footers=True`:

**Markdown:**
```markdown
---

**Header**: [header content]

**Footer**: [footer content]

---
```

**Plain Text:**
```
---

Header: [header content]

Footer: [footer content]

---
```

Headers/footers appear at the end of the document, after a horizontal rule.

### 5.3 Footnotes and Endnotes

**Markdown:** Uses standard markdown footnote syntax.

```markdown
This has a footnote[^1] and another[^2].

[^1]: First footnote content.
[^2]: Second footnote content.
```

**Plain Text:** Uses bracketed numbers.

```
This has a footnote [1] and another [2].

---
Footnotes:
[1] First footnote content.
[2] Second footnote content.
```

Endnotes are rendered the same way but labeled "Endnotes:".

---

## 6. Implementation Architecture

### 6.1 Class Structure

```python
# src/python_docx_redline/accessibility/export.py

@dataclass
class TextExportConfig:
    """Configuration for text/markdown export."""
    include_headers: bool = True
    include_footers: bool = True
    include_footnotes: bool = True
    include_endnotes: bool = True
    include_comments: bool = False
    include_images: bool = True
    tracked_changes: Literal["accept", "reject", "all"] = "accept"
    table_format: Literal["markdown", "simple", "grid"] = "markdown"
    line_width: int = 0
    heading_style: Literal["underline", "prefix"] = "underline"


class TextExporter:
    """Base class for text export."""

    def __init__(self, tree: AccessibilityTree, config: TextExportConfig):
        self.tree = tree
        self.config = config
        self.footnotes: list[tuple[int, str]] = []
        self.endnotes: list[tuple[int, str]] = []

    def export(self) -> str:
        """Export the tree to text format."""
        ...

    def _render_node(self, node: AccessibilityNode) -> str:
        """Render a single node to text."""
        ...

    def _render_paragraph(self, node: AccessibilityNode) -> str:
        ...

    def _render_heading(self, node: AccessibilityNode) -> str:
        ...

    def _render_table(self, node: AccessibilityNode) -> str:
        ...

    def _render_list(self, node: AccessibilityNode) -> str:
        ...

    def _apply_tracked_changes(self, text: str, changes: list) -> str:
        """Apply tracked changes based on config."""
        ...


class MarkdownExporter(TextExporter):
    """Markdown-specific export."""

    def _render_heading(self, node: AccessibilityNode) -> str:
        level = node.heading_level or 1
        return f"{'#' * level} {node.text}\n"

    def _render_table(self, node: AccessibilityNode) -> str:
        # Markdown table format
        ...


class PlainTextExporter(TextExporter):
    """Plain text export."""

    def _render_heading(self, node: AccessibilityNode) -> str:
        if self.config.heading_style == "underline":
            underline = "=" if node.heading_level == 1 else "-"
            return f"{node.text}\n{underline * len(node.text)}\n"
        else:
            return f"{'#' * node.heading_level} {node.text}\n"
```

### 6.2 Integration with AccessibilityTree

```python
# In tree.py

class AccessibilityTree:
    ...

    def to_text(self, config: TextExportConfig | None = None) -> str:
        """Export as plain text.

        Args:
            config: Export configuration. If None, uses defaults.

        Returns:
            Plain text representation of the document.
        """
        from .export import PlainTextExporter, TextExportConfig
        cfg = config or TextExportConfig()
        exporter = PlainTextExporter(self, cfg)
        return exporter.export()

    def to_markdown(self, config: TextExportConfig | None = None) -> str:
        """Export as markdown.

        Args:
            config: Export configuration. If None, uses defaults.

        Returns:
            Markdown representation of the document.
        """
        from .export import MarkdownExporter, TextExportConfig
        cfg = config or TextExportConfig()
        exporter = MarkdownExporter(self, cfg)
        return exporter.export()
```

---

## 7. Comparison with Pandoc

| Feature | Pandoc | Our Implementation |
|---------|--------|-------------------|
| External dependency | Yes (CLI) | No (pure Python) |
| Headings | ✓ | ✓ |
| Tables | ✓ (markdown) | ✓ (markdown, simple, grid) |
| Lists | ✓ | ✓ |
| Footnotes | ✓ | ✓ |
| Headers/footers | ✓ | ✓ |
| Tracked changes | accept/reject/all | accept/reject/all (CriticMarkup) |
| Comments | ✗ | ✓ |
| Images | ✓ | ✓ (placeholders) |
| Hyperlinks | ✓ | ✓ |
| Block quotes | ✓ | ✓ |
| Line wrapping | ✓ | ✓ |
| Refs for editing | ✗ | ✓ (via AccessibilityTree) |

**Key advantages over pandoc:**
1. No external dependency
2. Comments support
3. Same tree used for reading AND editing
4. CriticMarkup output matches our existing support
5. Pure Python, works everywhere

---

## 8. Test Plan

### 8.1 Unit Tests

1. **Basic rendering**: Each element type renders correctly
2. **Tracked changes modes**: Accept, reject, all produce correct output
3. **Table formats**: Markdown, simple, grid all work
4. **Document parts**: Headers, footers, footnotes, endnotes
5. **Configuration**: All config options work
6. **Edge cases**: Empty documents, nested lists, complex tables

### 8.2 Integration Tests

1. **Round-trip**: Read document, export, compare to pandoc output
2. **Large documents**: Performance on 100+ page documents
3. **Complex documents**: Mix of all element types

### 8.3 Comparison Tests

1. Compare output to pandoc for same document
2. Verify feature parity (or superiority)

---

## 9. Migration Guide

### Before (pandoc):
```bash
pandoc --track-changes=all document.docx -o output.md
```

### After (python-docx-redline):
```python
from python_docx_redline import Document
from python_docx_redline.accessibility import AccessibilityTree, TextExportConfig

doc = Document("document.docx")
tree = AccessibilityTree.from_document(doc)

# Equivalent to pandoc --track-changes=all
config = TextExportConfig(tracked_changes="all")
md = tree.to_markdown(config)

with open("output.md", "w") as f:
    f.write(md)
```

Or simpler:
```python
# Default: tracked_changes="accept" (like pandoc default)
md = tree.to_markdown()
```

---

## 10. Implementation Checklist

- [ ] Create `TextExportConfig` dataclass
- [ ] Create base `TextExporter` class
- [ ] Implement `PlainTextExporter`
- [ ] Implement `MarkdownExporter`
- [ ] Add `to_text()` method to AccessibilityTree
- [ ] Add `to_markdown()` method to AccessibilityTree
- [ ] Handle tracked changes (accept/reject/all)
- [ ] Handle tables (markdown/simple/grid)
- [ ] Handle footnotes and endnotes
- [ ] Handle headers and footers
- [ ] Handle comments
- [ ] Handle images (placeholders)
- [ ] Handle hyperlinks
- [ ] Handle block quotes
- [ ] Handle lists (bullet and numbered)
- [ ] Write unit tests
- [ ] Write integration tests
- [ ] Update documentation
- [ ] Update skill to recommend over pandoc
