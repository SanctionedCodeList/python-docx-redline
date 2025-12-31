---
name: docx-python
description: "Python library for Word document manipulation. Use python-docx for creating new documents, python-docx-redline for ALL editing tasks (handles run fragmentation, optional tracked changes). Covers text extraction, find/replace, comments, footnotes, hyperlinks, styles, CriticMarkup, and raw OOXML."
---

# Python DOCX Libraries

Two libraries for Word document manipulation:

- **python-docx**: Creating new documents from scratch
- **python-docx-redline**: Editing existing documents (recommended for ALL editing)

## Why python-docx-redline for Editing?

python-docx-redline handles **run fragmentation** that breaks python-docx find/replace. Word splits text across XML runs unpredictably—"Hello World" might be `<w:r>Hel</w:r><w:r>lo Wor</w:r><w:r>ld</w:r>`. python-docx-redline finds and edits text regardless of fragmentation.

## Quick Start

```python
# Creating new documents
from docx import Document
doc = Document()
doc.add_heading("Title", 0)
doc.add_paragraph("Content here.")
doc.save("new.docx")

# Editing existing documents (silent or tracked)
from python_docx_redline import Document
doc = Document("existing.docx")
doc.replace("OLD", "new")                    # Silent edit
doc.replace("30 days", "45 days", track=True)  # Tracked change
doc.save("modified.docx")
```

## Decision Tree

| Task | Guide |
|------|-------|
| **Create new document** | [creation.md](./creation.md) |
| **Generate from data/template** | [templating.md](./templating.md) |
| **Extract text** | [reading.md](./reading.md) |
| **Edit existing document** | [editing.md](./editing.md) |
| **Edit with tracked changes** | [tracked-changes.md](./tracked-changes.md) |
| **Add comments** | [comments.md](./comments.md) |
| **Footnotes/endnotes** | [footnotes.md](./footnotes.md) |
| **Hyperlinks** | [hyperlinks.md](./hyperlinks.md) |
| **Style management** | [styles.md](./styles.md) |
| **CriticMarkup round-trip** | [criticmarkup.md](./criticmarkup.md) |
| **Use both libraries together** | [integration.md](./integration.md) |
| **Structured YAML view (refs)** | [accessibility.md](./accessibility.md) |
| **Raw OOXML manipulation** | [ooxml.md](./ooxml.md) |

## Common Patterns

### Silent Editing (No Tracking)

```python
doc = Document("contract.docx")
doc.replace("{{NAME}}", "John Doe")      # Template population
doc.insert(" Inc.", after="Acme Corp")   # Append text
doc.delete("DRAFT - ")                   # Remove text
doc.save("filled.docx")
```

### Tracked Changes

```python
doc = Document("contract.docx")
doc.replace("30 days", "45 days", track=True)
doc.insert(" (amended)", after="Section 2.1", track=True)
doc.delete("subject to approval", track=True)
doc.save("redlined.docx")
```

### Find Before Editing

```python
matches = doc.find_all("payment")
for m in matches:
    print(f"{m.index}: {m.context}")

# Target specific occurrence
doc.replace("payment", "Payment", occurrence=2, track=True)
```

### Scoped Edits

```python
doc.replace("Client", "Customer", scope="section:Payment Terms", track=True)
```

### Batch Operations

```python
edits = [
    {"type": "replace", "find": "{{NAME}}", "replace": "John"},
    {"type": "replace", "find": "old", "replace": "new", "track": True},
    {"type": "delete", "text": "DRAFT"},
]
doc.apply_edits(edits, default_track=False)
```

## For LLM/Agent Workflows

Use the **AccessibilityTree** for structured YAML output with stable refs:

```python
from python_docx_redline import Document
doc = Document("contract.docx")
tree = doc.accessibility_tree()
print(tree.to_yaml())  # Structured view with refs like p:5, tbl:0/row:1/cell:2
```

Then edit by ref for precision:

```python
doc.replace_by_ref("p:5", "New paragraph text", track=True)
doc.delete_by_ref("p:10", track=True)
```

See [accessibility.md](./accessibility.md) for full details.

## Reference Files

| File | Purpose |
|------|---------|
| [creation.md](./creation.md) | Creating documents with style templates |
| [templating.md](./templating.md) | DocxBuilder: generate from data with markdown |
| [reading.md](./reading.md) | Text extraction, find_all(), document structure |
| [editing.md](./editing.md) | All editing operations (tracked and untracked) |
| [tracked-changes.md](./tracked-changes.md) | Tracked changes: insert/delete/replace, scopes, batch |
| [comments.md](./comments.md) | Adding comments, replies, resolution |
| [footnotes.md](./footnotes.md) | Footnotes/endnotes: CRUD, tracked changes, search |
| [hyperlinks.md](./hyperlinks.md) | Insert, edit, remove hyperlinks |
| [styles.md](./styles.md) | Style management and formatting |
| [criticmarkup.md](./criticmarkup.md) | Export/import tracked changes as markdown |
| [integration.md](./integration.md) | python-docx integration workflows |
| [accessibility.md](./accessibility.md) | AccessibilityTree for agent workflows |
| [ooxml.md](./ooxml.md) | Raw XML manipulation for complex scenarios |
