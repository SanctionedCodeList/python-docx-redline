---
name: docx
description: "Document creation, editing, and analysis with tracked changes and comments. Use for .docx files: creating documents, editing with tracked changes, adding comments, or text extraction."
---

# DOCX Skill - Quick Reference

## Installation

```bash
pip install python-docx                # Creating new documents
pip install python-docx-redline        # Editing with tracked changes (recommended)
brew install pandoc                    # Text extraction (macOS)
```

## Decision Tree

| Task | Tool | Guide |
|------|------|-------|
| **Read/extract text** | pandoc or python-docx-redline | [reading.md](./reading.md) |
| **Create new document** | python-docx | See below |
| **Edit with tracked changes** | python-docx-redline | [tracked-changes.md](./tracked-changes.md) |
| **Add comments** | python-docx-redline | [comments.md](./comments.md) |
| **Complex XML manipulation** | Raw OOXML | [ooxml.md](./ooxml.md) |

## Quick Examples

### Extract Text
```bash
pandoc --track-changes=all document.docx -o output.md
```

### Create New Document
```python
from docx import Document
doc = Document()
doc.add_heading("Title", 0)
doc.add_paragraph("Content here.")
doc.save("new.docx")
```

### Edit with Tracked Changes
```python
from python_docx_redline import Document

doc = Document("contract.docx")
doc.replace_tracked("30 days", "45 days")
doc.insert_tracked(" (amended)", after="Section 2.1")
doc.delete_tracked("subject to approval")
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
    {"type": "replace_tracked", "find": "old", "replace": "new"},
    {"type": "delete_tracked", "text": "remove this"},
]
doc.apply_edits(edits)
```

## Detailed Guides

- **[reading.md](./reading.md)** - Text extraction, find_all(), document structure, tables
- **[tracked-changes.md](./tracked-changes.md)** - Insert/delete/replace, regex, scopes, batch ops, formatting
- **[comments.md](./comments.md)** - Adding comments, occurrence parameter, replies, resolution
- **[ooxml.md](./ooxml.md)** - Raw XML manipulation for complex scenarios
