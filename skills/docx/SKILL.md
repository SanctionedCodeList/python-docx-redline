---
name: docx
description: "Document creation, editing, and analysis with tracked changes. Use for .docx files: creating documents, editing with or without tracked changes, adding comments, text extraction, template population, Table of Contents (insert/update/remove TOC), or CriticMarkup round-trip workflows (export/import tracked changes as markdown). python-docx-redline is the recommended tool for ALL editing tasks - it handles run fragmentation that breaks python-docx find/replace, with optional tracked changes via track=True."
---

# DOCX Skill

This skill creates professional, persuasive documents that drive decisions. It combines design intelligence from elite business contexts (consulting, banking, legal) with robust technical workflows for DOCX manipulation.

## Design Thinking

Before creating or editing any document, understand the context:

- **Purpose**: What decision or action should this document drive?
- **Audience**: Who are they? What do they need? How much time do they have?
- **Key Message**: If the reader remembers only ONE thing, what must it be?
- **Tone**: Match the context — analytical (consulting), precise (legal), data-driven (banking)

### The Non-Negotiables

**1. Lead with the Answer**
State your conclusion or recommendation in the first paragraph of the document and the first sentence of each section. Don't make executives hunt for your point.

**2. Action Headings**
Every section heading states the **takeaway**, not the topic.

| Topic Heading (Weak) | Action Heading (Strong) |
|----------------------|-------------------------|
| "Q3 Results" | "Q3 revenue beat targets by 12%, driven by enterprise" |
| "Market Analysis" | "Market consolidating—three players now control 70%" |

See [references/action-headings.md](references/action-headings.md) for details.

**3. Pyramid Structure**
Organize content in a pyramid: main point at top, supporting arguments below, evidence below that.

See [references/document-structure.md](references/document-structure.md) for frameworks.

**4. Professional Formatting**
Typography, spacing, and layout signal credibility before anyone reads a word.

See [references/design-principles.md](references/design-principles.md) for guidelines.

### Document Anti-Patterns

**NEVER** produce generic documents. Avoid:

- **Topic headings**: "Background", "Analysis" — say nothing. Use action headings
- **Burying the lead**: Recommendation on page 20 means failure. Lead with conclusions
- **Wall of text**: No headings, no structure, no scanability
- **Hedge-everything language**: "We might consider potentially..." — take a position
- **Inconsistent formatting**: Mixed fonts, random spacing signals carelessness

---

## Technical Workflows

### Installation

```bash
pip install python-docx                # Creating new documents
pip install python-docx-redline        # Editing with tracked changes (recommended)
brew install pandoc                    # Text extraction (macOS)
```

### Library Source & Issues

**python-docx-redline** source: [parkerhancock/python-docx-redline](https://github.com/parkerhancock/python-docx-redline)

Report bugs, unexpected behavior, or feature requests as issues on the repository. Include the python-docx-redline version, minimal reproduction code, and the problematic .docx file if possible.

## Decision Tree

| Task | Tool | Guide |
|------|------|-------|
| **Read/extract text** | pandoc or python-docx-redline | [reading.md](./reading.md) |
| **Structured document view (YAML)** | AccessibilityTree | [accessibility.md](./accessibility.md) |
| **Large document navigation** | OutlineTree | [accessibility.md](./accessibility.md) |
| **Ref-based precise editing** | python-docx-redline refs | [accessibility.md](./accessibility.md) |
| **Create new document** | python-docx | [creation.md](./creation.md) |
| **Generate from data/template** | DocxBuilder | [templating.md](./templating.md) |
| **Edit existing document** | python-docx-redline | [editing.md](./editing.md) |
| **Edit with tracked changes** | python-docx-redline (track=True) | [tracked-changes.md](./tracked-changes.md) |
| **Delete entire section** | python-docx-redline delete_section() | [editing.md](./editing.md#section-operations) |
| **Add comments** | python-docx-redline | [comments.md](./comments.md) |
| **Footnotes/endnotes** | python-docx-redline | [footnotes.md](./footnotes.md) |
| **Insert or edit hyperlinks** | python-docx-redline | [hyperlinks.md](./hyperlinks.md) |
| **Table of Contents** | python-docx-redline | [toc.md](./toc.md) |
| **Cross-references** | python-docx-redline | [cross-references.md](./cross-references.md) |
| **Bookmarks** | python-docx-redline | [cross-references.md](./cross-references.md) |
| **Create or manage styles** | python-docx-redline StyleManager | [styles.md](./styles.md) |
| **CriticMarkup workflow** | python-docx-redline | [criticmarkup.md](./criticmarkup.md) |
| **Use both libraries together** | from_python_docx / to_python_docx | [integration.md](./integration.md) |
| **Complex XML manipulation** | Raw OOXML | [ooxml.md](./ooxml.md) |

**Note:** python-docx-redline is recommended for ALL editing (not just tracked changes) because it handles run fragmentation that breaks python-docx find/replace.

**For LLM/Agent workflows:** Use the AccessibilityTree for structured YAML output that fits in context windows, with stable refs for unambiguous element targeting. See [accessibility.md](./accessibility.md).

## Quick Examples

### Extract Text
```bash
pandoc --track-changes=all document.docx -o output.md
```

### Create New Document
```python
from docx import Document

# Use a style template for custom styles (see creation.md)
doc = Document("styles/corporate.docx")
doc.add_heading("Title", 0)
doc.add_paragraph("Content here.")
doc.save("new.docx")
```

### Generate from Data (DocxBuilder)
```python
from python_docx_redline import DocxBuilder

doc = DocxBuilder()
doc.heading("Sales Report")
doc.markdown("Revenue **exceeded** targets by 12%.")
doc.table_from(items, ["product", "revenue", "growth"])
doc.save("report.docx")
```

### Edit Existing Document (Silent)
```python
from python_docx_redline import Document

doc = Document("existing.docx")
doc.replace("OLD_VALUE", "new_value")         # Handles run boundaries
doc.replace("{{NAME}}", "John Doe")           # Template population
doc.insert(" Inc.", after="Acme Corp")        # Append text
doc.delete("DRAFT - ")                        # Remove text
doc.save("modified.docx")
```

### Edit with Tracked Changes
```python
from python_docx_redline import Document

doc = Document("contract.docx")
doc.replace("30 days", "45 days", track=True)     # With track parameter
doc.insert(" (amended)", after="Section 2.1", track=True)
doc.delete("subject to approval", track=True)
# Or use explicit *_tracked methods:
doc.replace_tracked("Contractor", "Service Provider")
doc.save("contract_redlined.docx")
```

### Add Comments
```python
doc.add_comment("Please review", on="Section 2.1")
doc.add_comment("Check all", on="TODO", occurrence="all")  # Multiple occurrences
```

### Find Text Before Editing
```python
matches = doc.find_all("payment")
for m in matches:
    print(f"{m.index}: {m.context}")

# Then target specific occurrence
doc.replace_tracked("payment", "Payment", occurrence=2)
```

### Scoped Edits
```python
doc.replace_tracked("Client", "Customer", scope="section:Payment Terms")
```

### Footnotes and Endnotes
```python
# Insert footnotes
doc.insert_footnote("See Smith (2020) for details", at="original study")
doc.insert_footnote(["First paragraph.", "Second with **bold**."], at="citation")

# Get, edit, delete footnotes
footnote = doc.get_footnote(1)
footnote.edit("Updated citation")
footnote.delete()

# Tracked changes inside footnotes
doc.insert_tracked_in_footnote(1, " [revised]", after="citation")
doc.replace_tracked_in_footnote(1, "2020", "2024")

# Search in footnotes
matches = doc.find_all("reference", scope="footnotes")
matches = doc.find_all("citation", include_footnotes=True)
```

### Table of Contents
```python
# Insert TOC at start of document
doc.insert_toc(levels=(1, 3), title="Contents")

# Inspect existing TOC
toc = doc.get_toc()
if toc:
    print(f"Levels: {toc.levels}, Dirty: {toc.is_dirty}")

# Update TOC settings in place
doc.update_toc(levels=(1, 5), hyperlinks=False)

# Remove TOC
doc.remove_toc()
```

## Common Patterns

### Handle Ambiguous Text
```python
# If text appears multiple times, use occurrence parameter
doc.replace_tracked("Section", "Article", occurrence=1)      # First match
doc.replace_tracked("Section", "Article", occurrence="all")  # All matches
```

### Smart Quote Handling
```python
# Curly quotes in documents match straight quotes in code automatically
doc.replace_tracked("Defendant's motion", "party's motion")  # Just works
```

### Batch Operations
```python
edits = [
    {"type": "replace", "find": "{{NAME}}", "replace": "John"},  # Untracked
    {"type": "replace", "find": "old", "replace": "new", "track": True},  # Tracked
    {"type": "delete", "text": "DRAFT", "track": False},  # Explicit untracked
]
doc.apply_edits(edits, default_track=False)  # Set default for edits without track field
```

### CriticMarkup Round-Trip
```python
# Export tracked changes to markdown
doc = Document("contract.docx")
markdown = doc.to_criticmarkup()
# Output: "Payment in {--30--}{++45++} days"

# Apply CriticMarkup changes back to DOCX
doc.apply_criticmarkup("{++new clause++}", author="Reviewer")
doc.save("updated.docx")
```

---

## Design References

Guidance on creating effective, professional documents:

- **[references/action-headings.md](references/action-headings.md)** — Writing insight-driven headings
- **[references/document-structure.md](references/document-structure.md)** — Pyramid Principle, SCQA, IRAC frameworks
- **[references/executive-summaries.md](references/executive-summaries.md)** — Crafting standalone summaries
- **[references/industry-styles.md](references/industry-styles.md)** — Consulting, banking, legal, VC conventions
- **[references/design-principles.md](references/design-principles.md)** — Typography, layout, visual hierarchy

## Technical Guides

Detailed workflows for document manipulation:

- **[accessibility.md](./accessibility.md)** — DocTree accessibility layer: YAML output, refs, OutlineTree for large docs
- **[templating.md](./templating.md)** — DocxBuilder: generate documents from data with markdown support
- **[creation.md](./creation.md)** — Creating new documents with style templates
- **[reading.md](./reading.md)** — Text extraction, find_all(), document structure, tables
- **[editing.md](./editing.md)** — All editing with python-docx-redline (both tracked and untracked)
- **[tracked-changes.md](./tracked-changes.md)** — Tracked changes details: insert/delete/replace, regex, scopes, batch ops
- **[comments.md](./comments.md)** — Adding comments, occurrence parameter, replies, resolution
- **[footnotes.md](./footnotes.md)** — Footnotes/endnotes: CRUD, tracked changes, rich content, search
- **[hyperlinks.md](./hyperlinks.md)** — Hyperlink operations: insert, edit, remove in body, headers, footers, footnotes
- **[toc.md](./toc.md)** — Table of Contents: insert, inspect, update, remove TOC
- **[cross-references.md](./cross-references.md)** — Cross-references and bookmarks: reference headings, figures, tables, notes
- **[styles.md](./styles.md)** — Style management: reading, creating, ensuring styles exist, formatting options
- **[criticmarkup.md](./criticmarkup.md)** — Export/import with CriticMarkup, round-trip workflows
- **[integration.md](./integration.md)** — python-docx integration: from_python_docx, to_python_docx, workflows
- **[ooxml.md](./ooxml.md)** — Raw XML manipulation for complex scenarios

---

Remember: Claude is capable of creating documents that rival top-tier consulting and legal firms. Lead with your answer, use action headings, and execute every detail with intention. The goal isn't a "good enough" document—it's one that drives decisions.
