---
name: docx
description: "Word document creation, editing, and manipulation with Python. Use for .docx files: creating documents, editing with or without tracked changes, comments, footnotes, Table of Contents, cross-references, text extraction, template population, CriticMarkup workflows. Two sub-skills: design/ for professional document writing, python/ for python-docx/python-docx-redline operations."
---

# DOCX Skill

Professional Word document creation and manipulation.

## Installation

```bash
./install.sh
```

### Library Source & Issues

**python-docx-redline** source: [SanctionedCodeList/python-docx-redline](https://github.com/SanctionedCodeList/python-docx-redline)

Report bugs, unexpected behavior, or feature requests as issues on the repository. Include the python-docx-redline version, minimal reproduction code, and the problematic .docx file if possible.

## Decision Tree

| What do you need? | Go to |
|-------------------|-------|
| **Write a professional document** (memo, report, proposal) | [design/](./design/SKILL.md) |
| **Create or edit .docx files with Python** | [python/](./python/SKILL.md) |
| **Edit live in Microsoft Word** (via add-in) | Install [office-bridge](https://github.com/SanctionedCodeList/office-bridge) plugin |

### Python Library Tasks

| Task | Tool | Guide |
|------|------|-------|
| **Read/extract text** | pandoc or python-docx-redline | [python/reading.md](./python/reading.md) |
| **Structured document view (YAML)** | AccessibilityTree | [python/accessibility.md](./python/accessibility.md) |
| **Large document navigation** | OutlineTree | [python/accessibility.md](./python/accessibility.md) |
| **Ref-based precise editing** | python-docx-redline refs | [python/accessibility.md](./python/accessibility.md) |
| **Create new document** | python-docx | [python/creation.md](./python/creation.md) |
| **Generate from data/template** | DocxBuilder | [python/templating.md](./python/templating.md) |
| **Edit existing document** | python-docx-redline | [python/editing.md](./python/editing.md) |
| **Edit with tracked changes** | python-docx-redline (track=True) | [python/tracked-changes.md](./python/tracked-changes.md) |
| **Delete entire section** | python-docx-redline delete_section() | [python/editing.md](./python/editing.md#section-operations) |
| **Add comments** | python-docx-redline | [python/comments.md](./python/comments.md) |
| **Footnotes/endnotes** | python-docx-redline | [python/footnotes.md](./python/footnotes.md) |
| **Insert or edit hyperlinks** | python-docx-redline | [python/hyperlinks.md](./python/hyperlinks.md) |
| **Table of Contents** | python-docx-redline | [toc.md](./toc.md) |
| **Cross-references** | python-docx-redline | [cross-references.md](./cross-references.md) |
| **Bookmarks** | python-docx-redline | [cross-references.md](./cross-references.md) |
| **Create or manage styles** | python-docx-redline StyleManager | [python/styles.md](./python/styles.md) |
| **CriticMarkup workflow** | python-docx-redline | [python/criticmarkup.md](./python/criticmarkup.md) |
| **Use both libraries together** | from_python_docx / to_python_docx | [python/integration.md](./python/integration.md) |
| **Complex XML manipulation** | Raw OOXML | [python/ooxml.md](./python/ooxml.md) |

## Quick Reference

### Design Principles (Always Apply)

- **Lead with the answer** — Conclusion first, evidence after
- **Action headings** — "Revenue grew 12%" not "Q3 Results"
- **Pyramid structure** — Main point → supporting arguments → evidence

### Python Quick Start

```python
from python_docx_redline import Document

doc = Document("contract.docx")
doc.replace("30 days", "45 days", track=True)  # Tracked change
doc.save("redlined.docx")
```

## Sub-Skills

| Folder | Purpose |
|--------|---------|
| [design/](./design/SKILL.md) | Document design: action headings, pyramid structure, industry styles, AI antipatterns |
| [python/](./python/SKILL.md) | Python libraries: python-docx (creation), python-docx-redline (editing, tracked changes) |

## Code Examples

### Targeted Edits
```python
# Find all occurrences first
matches = doc.find_all("payment")
for m in matches:
    print(f"[{m.index}] {m.context}")

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

- **[python/accessibility.md](./python/accessibility.md)** — DocTree accessibility layer: YAML output, refs, OutlineTree for large docs
- **[python/templating.md](./python/templating.md)** — DocxBuilder: generate documents from data with markdown support
- **[python/creation.md](./python/creation.md)** — Creating new documents with style templates
- **[python/reading.md](./python/reading.md)** — Text extraction, find_all(), document structure, tables
- **[python/editing.md](./python/editing.md)** — All editing with python-docx-redline (both tracked and untracked)
- **[python/tracked-changes.md](./python/tracked-changes.md)** — Tracked changes details: insert/delete/replace, regex, scopes, batch ops
- **[python/comments.md](./python/comments.md)** — Adding comments, occurrence parameter, replies, resolution
- **[python/footnotes.md](./python/footnotes.md)** — Footnotes/endnotes: CRUD, tracked changes, rich content, search
- **[python/hyperlinks.md](./python/hyperlinks.md)** — Hyperlink operations: insert, edit, remove in body, headers, footers, footnotes
- **[toc.md](./toc.md)** — Table of Contents: insert, inspect, update, remove TOC
- **[cross-references.md](./cross-references.md)** — Cross-references and bookmarks: reference headings, figures, tables, notes
- **[python/styles.md](./python/styles.md)** — Style management: reading, creating, ensuring styles exist, formatting options
- **[python/criticmarkup.md](./python/criticmarkup.md)** — Export/import with CriticMarkup, round-trip workflows
- **[python/integration.md](./python/integration.md)** — python-docx integration: from_python_docx, to_python_docx, workflows
- **[python/ooxml.md](./python/ooxml.md)** — Raw XML manipulation for complex scenarios

---

Remember: Claude is capable of creating documents that rival top-tier consulting and legal firms. Lead with your answer, use action headings, and execute every detail with intention. The goal isn't a "good enough" document—it's one that drives decisions.
