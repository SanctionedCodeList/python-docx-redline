# python-docx-redline

**High-level Word document editing with tracked changes support.**

A Python library that makes programmatic Word editing reliable by handling XML run fragmentation and providing simple tracked changes APIs.

---

## 🚀 Installation

### As Python Library

```bash
pip install python-docx-redline
```

Requires Python 3.10+

### As Claude Code Plugin

```bash
claude plugins add SanctionedCodeList/python-docx-redline
```

Or via the SCL marketplace:

```bash
claude plugins add SanctionedCodeList/SCL_marketplace
```

---

## The Problem

Editing Word documents programmatically is frustrating. Text is fragmented across XML runs, making even simple find-and-replace unreliable.

**Before** (raw python-docx):
```python
from docx import Document

doc = Document("contract.docx")
# This often fails - "Contract" might be split as "Con" + "tract"
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

---

## Quick Start

### Silent Editing (No Revision Marks)

```python
from python_docx_redline import Document

doc = Document("template.docx")
doc.replace("{{NAME}}", "John Doe")
doc.replace("{{DATE}}", "2024-12-28")
doc.insert(" Inc.", after="Acme Corp")
doc.delete("DRAFT - ")
doc.save("output.docx")
```

### Tracked Changes (With Revision Marks)

```python
from python_docx_redline import Document

doc = Document("contract.docx")
doc.replace("30 days", "45 days", track=True)
doc.insert(" (revised)", after="Exhibit A", track=True)
doc.delete("unless otherwise agreed", track=True)
doc.save("contract_redlined.docx")
```

---

## Features

### Text Operations
| Feature | Description |
|---------|-------------|
| **Smart search** | Handles text fragmented across XML runs |
| **Replace** | Find and replace with formatting preservation |
| **Insert/Delete** | Add or remove text at any position |
| **Regex support** | Pattern matching with capture groups |
| **Fuzzy matching** | For OCR'd or inconsistent documents |

### Tracked Changes
| Feature | Description |
|---------|-------------|
| **Insertions** | Show added text in revision marks |
| **Deletions** | Show removed text with strikethrough |
| **Replacements** | Combined delete + insert |
| **Format changes** | Bold, italic changes as revisions |

### Structural Operations
| Feature | Description |
|---------|-------------|
| **Section navigation** | Find and edit by heading |
| **Paragraph insertion** | Add new paragraphs with styles |
| **Scoped edits** | Limit changes to specific sections |

### Footnotes & Endnotes
| Feature | Description |
|---------|-------------|
| **CRUD operations** | Insert, read, edit, delete notes |
| **Tracked changes** | Revisions inside footnotes |
| **Auto-renumbering** | Notes renumber when deleted |

---

## Usage Examples

### Batch Operations from YAML

```yaml
# edits.yaml
default_track: false

edits:
  - type: replace
    find: "{{COMPANY}}"
    replace: "Acme Inc."

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

### Reading Document Structure

```python
# Understand structure before editing
for section in doc.sections:
    print(f"{section.heading_text}: {len(section.paragraphs)} paragraphs")

# Find all occurrences with context
matches = doc.find_all("confidential", context_chars=50)
for m in matches:
    print(f"  ...{m.context_before}[{m.matched_text}]{m.context_after}...")
```

---

## With Claude Code

After installing the plugin, describe what you need:

```
> Open contract.docx and replace "Contractor" with "Service Provider" using tracked changes

> Find all instances of "confidential" in the document and show me the context

> Create a redlined version replacing the payment terms from 30 to 45 days
```

---

## Documentation

| Document | Description |
|----------|-------------|
| [Documentation Index](docs/index.md) | All documentation |
| [API Reference](docs/PROPOSED_API.md) | Complete API documentation |
| [Quick Reference](docs/QUICK_REFERENCE.md) | Cheat sheet |
| [OOXML Internals](docs/ERIC_WHITE_ALGORITHM.md) | How text search works |

---

## Development

```bash
git clone https://github.com/SanctionedCodeList/python-docx-redline.git
cd python-docx-redline
pip install -e ".[dev]"
pytest
```

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Text not found | May be fragmented across runs; library handles this automatically |
| Formatting lost | Use `replace()` instead of modifying `para.text` directly |
| Tracked changes not showing | Ensure `track=True` parameter is set |
| Import errors | Ensure Python 3.10+ is installed |

---

## Links

- [GitHub](https://github.com/SanctionedCodeList/python-docx-redline)
- [PyPI](https://pypi.org/project/python-docx-redline/)
- [SCL Marketplace](https://github.com/SanctionedCodeList/SCL_marketplace)
