# python-docx-redline

A high-level Python API for editing Word documents, with optional tracked changes support.

## The Problem

Editing Word documents programmatically is frustrating. Text is fragmented across XML runs, making even simple find-and-replace unreliable. Adding tracked changes requires 30+ lines of OOXML namespace handling.

**Before** (raw python-docx):
```python
from docx import Document

doc = Document("contract.docx")
# This often fails - "Contract" might be split across runs as "Con" + "tract"
for para in doc.paragraphs:
    if "Contract" in para.text:
        para.text = para.text.replace("Contract", "Agreement")  # Loses all formatting!
```

**After** (python-docx-redline):
```python
from python_docx_redline import Document

doc = Document("contract.docx")
doc.replace("Contract", "Agreement")  # Handles run boundaries, preserves formatting
doc.save("contract_edited.docx")
```

## Installation

```bash
pip install python-docx-redline
```

Requires Python 3.10+

## Quick Start

### Untracked (Silent) Editing

Make edits without revision marks - the document appears as if it was always that way:

```python
from python_docx_redline import Document

doc = Document("template.docx")
doc.replace("{{NAME}}", "John Doe")       # Silent replacement
doc.replace("{{DATE}}", "2024-12-28")     # No revision marks
doc.insert(" Inc.", after="Acme Corp")    # Append text silently
doc.delete("DRAFT - ")                    # Remove text silently
doc.save("output.docx")
```

### Tracked Editing

Show changes as tracked revisions for review:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Use track=True for tracked changes
doc.replace("30 days", "45 days", track=True)
doc.insert(" (revised)", after="Exhibit A", track=True)
doc.delete("unless otherwise agreed", track=True)

# Or use the explicit *_tracked methods
doc.replace_tracked("the Contractor", "the Service Provider")

doc.save("contract_redlined.docx")
```

## Features

### Text Operations
- **Insert, delete, replace, move** text with or without tracked changes
- **Smart text search** handles text fragmented across XML runs
- **Regex support** with capture groups for pattern matching
- **Fuzzy matching** for OCR'd or inconsistently formatted documents
- **Quote normalization** matches curly quotes with straight quotes

### Structural Operations
- **Insert paragraphs** with styles and formatting
- **Delete sections** by heading with tracked changes
- **Section/Paragraph wrappers** for document navigation

### Document Viewing
- **Read paragraphs** with text, style, and heading info
- **Parse sections** via heading structure
- **Find all occurrences** with context preview
- **Extract full text** for analysis

### Formatting
- **Markdown syntax** in insertions (`**bold**`, `*italic*`, `++underline++`)
- **Format-only tracked changes** (bold, italic without text changes)
- **Minimal editing mode** for clean legal redlines

### Advanced
- **Batch operations** from YAML/JSON configuration
- **Scope filtering** to limit operations to specific sections
- **MS365 identity integration** for enterprise environments
- **Context-aware editing** with fragment detection
- **Document rendering** to PNG images
- **Image insertion** with tracked changes

### Integration
- **python-docx compatibility** via `from_python_docx()` and `to_python_docx()`
- **In-memory workflows** with `save_to_bytes()`
- **Full type hints** for IDE support

## Examples

### Batch Operations from YAML

```yaml
# edits.yaml
default_track: false  # or true for all tracked

edits:
  - type: replace
    find: "{{COMPANY}}"
    replace: "Acme Inc."
    # Uses default_track (false = untracked)

  - type: replace
    find: "30 days"
    replace: "45 days"
    track: true  # Override: this one is tracked

  - type: insert
    text: " (as amended)"
    after: "Agreement dated"
```

```python
doc.apply_edit_file("edits.yaml")
```

### Scoped Replacements

```python
# Only replace in the Payment Terms section
doc.replace_tracked("net 30", "net 45", scope="section:Payment Terms")
```

### Reading Before Editing

```python
# Understand document structure before making changes
for section in doc.sections:
    print(f"{section.heading_text}: {len(section.paragraphs)} paragraphs")

# Find all occurrences with context
matches = doc.find_all("confidential", context_chars=50)
for m in matches:
    print(f"  ...{m.context_before}[{m.matched_text}]{m.context_after}...")
```

## Documentation

- **[Documentation Index](docs/index.md)** — All documentation
- **[API Reference](docs/PROPOSED_API.md)** — Complete API documentation
- **[Quick Reference](docs/QUICK_REFERENCE.md)** — Cheat sheet
- **[OOXML Internals](docs/ERIC_WHITE_ALGORITHM.md)** — How text search works

## Claude Code Plugin

For AI-assisted document editing, install as a Claude Code plugin:

```
/plugin install python-docx-redline@SanctionedCodeList/python-docx-redline
```

## Development

```bash
git clone https://github.com/parkerhancock/python_docx_redline.git
cd python_docx_redline
pip install -e ".[dev]"
pytest
```

## License

MIT License — see LICENSE file for details.

## Links

- [GitHub](https://github.com/parkerhancock/python_docx_redline)
- [PyPI](https://pypi.org/project/python-docx-redline/)
- [Documentation](docs/)
