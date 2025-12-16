# python-docx-redline

A high-level Python API for editing Word documents with tracked changes.

## The Problem

Making surgical edits to Word documents with tracked changes requires writing complex OOXML XML manipulation code—typically 30+ lines of namespace handling, element construction, and tree traversal.

**Before** (raw OOXML):
```python
from lxml import etree
from datetime import datetime, timezone

# Find the paragraph, handle namespaces, generate IDs, construct elements...
paragraphs = root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
for p in paragraphs:
    if 'Section 2.1' in ''.join(p.itertext()):
        # 25+ more lines of XML construction...
```

**After** (python-docx-redline):
```python
from python_docx_redline import Document

doc = Document("contract.docx")
doc.insert_tracked(" (as amended)", after="Section 2.1")
doc.save("contract_edited.docx")
```

## Installation

```bash
pip install python-docx-redline
```

Requires Python 3.10+

## Quick Start

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Insert text with tracked changes
doc.insert_tracked(" (revised)", after="Exhibit A")

# Replace text with tracked changes
doc.replace_tracked("the Contractor", "the Service Provider")

# Delete text with tracked changes
doc.delete_tracked("unless otherwise agreed")

doc.save("contract_edited.docx")
```

## Features

### Text Operations
- **Insert, delete, replace** text with tracked changes
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
edits:
  - type: replace_tracked
    find: "30 days"
    replace: "45 days"
  - type: insert_tracked
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

- **[User Guide](docs/user-guide/)** — Detailed usage documentation
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
