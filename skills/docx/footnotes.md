# Footnotes and Endnotes

Use **python-docx-redline** for comprehensive footnote and endnote operations including CRUD, tracked changes inside notes, and rich content.

## Overview

| Operation | Method | Tracked Support |
|-----------|--------|-----------------|
| **Insert** footnote | `insert_footnote()` | No (insertion point only) |
| **Get** footnote | `get_footnote(id)` | N/A |
| **Edit** footnote | `edit_footnote(id, text)` or `footnote.edit()` | Yes |
| **Delete** footnote | `delete_footnote(id)` or `footnote.delete()` | No |
| **Search** in footnotes | `find_all(scope="footnotes")` | N/A |

All methods work identically for endnotes by substituting "footnote" with "endnote".

## Basic Operations

### Insert a Footnote

```python
from python_docx_redline import Document

doc = Document("paper.docx")

# Insert a simple footnote
doc.insert_footnote("See Smith (2020) for details.", at="original study")

# Insert at specific occurrence
doc.insert_footnote("Citation needed.", at="claim", occurrence=2)

# Insert endnote
doc.insert_endnote("Additional context here.", at="methodology")

doc.save("paper_with_notes.docx")
```

### Get a Footnote

```python
# Get by ID (1-indexed, matching Word's display)
footnote = doc.get_footnote(1)
print(footnote.text)     # Plain text content
print(footnote.id)       # Footnote ID

# Get all footnotes
for fn in doc.footnotes:
    print(f"[{fn.id}] {fn.text}")

# Get all endnotes
for en in doc.endnotes:
    print(f"[{en.id}] {en.text}")
```

### Edit a Footnote

```python
# Edit via Document facade
doc.edit_footnote(1, "Updated citation: Smith (2024)")

# Or via model method
footnote = doc.get_footnote(1)
footnote.edit("Updated citation: Smith (2024)")

# Edit endnote
doc.edit_endnote(1, "Revised endnote content")
```

### Delete a Footnote

```python
# Delete via Document facade
doc.delete_footnote(1)

# Or via model method
footnote = doc.get_footnote(2)
footnote.delete()

# Remaining footnotes are automatically renumbered to match Word behavior
```

## Tracked Changes Inside Notes

Make tracked edits within footnote/endnote content:

```python
# Insert tracked text inside a footnote
doc.insert_tracked_in_footnote(1, " [revised]", after="citation")

# Delete tracked text inside a footnote
doc.delete_tracked_in_footnote(1, "preliminary")

# Replace tracked text inside a footnote
doc.replace_tracked_in_footnote(1, "2020", "2024")

# Same operations on endnotes
doc.insert_tracked_in_endnote(1, " (updated)", after="reference")
doc.delete_tracked_in_endnote(1, "old text")
doc.replace_tracked_in_endnote(1, "original", "revised")
```

### Using Model Methods

```python
footnote = doc.get_footnote(1)

# Tracked changes via model
footnote.insert_tracked(" [updated]", after="see")
footnote.delete_tracked("preliminary")
footnote.replace_tracked("2020", "2024")
```

## Rich Content

### Markdown in Footnotes

```python
# Single paragraph with markdown
doc.insert_footnote("See **Smith (2020)** for *detailed* analysis.", at="study")

# Multiple paragraphs
doc.insert_footnote([
    "First paragraph with **bold** text.",
    "Second paragraph with *italic* and ++underline++.",
    "Third with ~~strikethrough~~."
], at="complex citation")
```

Supported markdown:
- `**bold**` → bold text
- `*italic*` → italic text
- `++underline++` → underlined text
- `~~strikethrough~~` → strikethrough text

### Reading Rich Content

```python
footnote = doc.get_footnote(1)

# Get formatted representation
print(footnote.formatted_text)  # Markdown-style: "See **Smith** for details"
print(footnote.html)            # HTML: "See <b>Smith</b> for details"
```

## Searching in Footnotes

### Find All in Footnotes

```python
# Search only in footnotes
matches = doc.find_all("citation", scope="footnotes")

# Search only in endnotes
matches = doc.find_all("reference", scope="endnotes")

# Search in all notes
matches = doc.find_all("Smith", scope="notes")

# Search in specific footnote
matches = doc.find_all("2020", scope="footnote:1")

# Include footnotes in document-wide search
matches = doc.find_all("payment", include_footnotes=True)
matches = doc.find_all("appendix", include_endnotes=True)
```

### Convenience Methods

```python
# Find in all footnotes
matches = doc.find_in_footnotes("citation")

# Find in all endnotes
matches = doc.find_in_endnotes("reference")
```

## Finding Reference Locations

Get the location of a footnote reference in the main document:

```python
# Via Document method
location = doc.get_footnote_reference_location(1)
print(f"Footnote 1 is in paragraph {location.paragraph_index}")
print(f"Context: {location.context}")

# Via model property
footnote = doc.get_footnote(1)
location = footnote.reference_location
```

## Scope Parameters

| Scope | Description |
|-------|-------------|
| `"footnotes"` | All footnotes |
| `"endnotes"` | All endnotes |
| `"notes"` | All footnotes and endnotes |
| `"footnote:N"` | Specific footnote by ID |
| `"endnote:N"` | Specific endnote by ID |

## Error Handling

```python
from python_docx_redline import Document
from python_docx_redline.errors import NoteNotFoundError, TextNotFoundError

doc = Document("paper.docx")

try:
    footnote = doc.get_footnote(99)
except NoteNotFoundError as e:
    print(f"Footnote not found: {e}")
    print(f"Available IDs: {e.available_ids}")

try:
    doc.insert_footnote("Note", at="nonexistent text")
except TextNotFoundError as e:
    print(f"Anchor text not found: {e}")
```

## Complete Example

```python
from python_docx_redline import Document

doc = Document("research_paper.docx")

# Add footnotes with rich content
doc.insert_footnote(
    "See **Smith & Jones (2024)** for the *original* methodology.",
    at="novel approach"
)

doc.insert_footnote([
    "This finding contradicts earlier work.",
    "However, the sample size was ++significantly larger++."
], at="surprising result")

# Edit existing footnote with tracked changes
doc.replace_tracked_in_footnote(1, "2020", "2024")
doc.insert_tracked_in_footnote(1, " (revised)", after="methodology")

# Search across footnotes
for match in doc.find_all("Smith", scope="footnotes"):
    print(f"Found in footnote: {match.context}")

# Get reference location
fn = doc.get_footnote(1)
loc = fn.reference_location
print(f"Footnote 1 appears at: {loc.context}")

doc.save("annotated_paper.docx")
```

## API Reference

### Document Methods

| Method | Description |
|--------|-------------|
| `insert_footnote(text, at, occurrence=1)` | Insert footnote at anchor |
| `insert_endnote(text, at, occurrence=1)` | Insert endnote at anchor |
| `get_footnote(id)` | Get footnote by ID |
| `get_endnote(id)` | Get endnote by ID |
| `edit_footnote(id, text)` | Edit footnote content |
| `edit_endnote(id, text)` | Edit endnote content |
| `delete_footnote(id)` | Delete footnote (renumbers remaining) |
| `delete_endnote(id)` | Delete endnote (renumbers remaining) |
| `insert_tracked_in_footnote(id, text, after)` | Insert tracked text in footnote |
| `delete_tracked_in_footnote(id, text)` | Delete tracked text in footnote |
| `replace_tracked_in_footnote(id, find, replace)` | Replace tracked text in footnote |
| `get_footnote_reference_location(id)` | Get reference location in document |

### Footnote Model Methods

| Property/Method | Description |
|-----------------|-------------|
| `.id` | Footnote ID |
| `.text` | Plain text content |
| `.formatted_text` | Markdown-formatted content |
| `.html` | HTML-formatted content |
| `.reference_location` | Location in main document |
| `.edit(text)` | Edit content |
| `.delete()` | Delete footnote |
| `.insert_tracked(text, after)` | Insert tracked text |
| `.delete_tracked(text)` | Delete tracked text |
| `.replace_tracked(find, replace)` | Replace tracked text |

### Search Parameters

| Parameter | Description |
|-----------|-------------|
| `scope="footnotes"` | Search in all footnotes |
| `scope="footnote:N"` | Search in specific footnote |
| `include_footnotes=True` | Include footnotes in document search |
| `include_endnotes=True` | Include endnotes in document search |
