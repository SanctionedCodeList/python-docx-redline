# Cross-References and Bookmarks

Use **python-docx-redline** to create cross-references to headings, figures, tables, footnotes, endnotes, and bookmarks. Cross-references are field codes that Word calculates when the document opens.

## Quick Reference

| Operation | Method | Returns |
|-----------|--------|---------|
| Insert cross-reference | `insert_cross_reference()` | bookmark name |
| Insert page reference | `insert_page_reference()` | bookmark name |
| Insert note reference | `insert_note_reference()` | bookmark name |
| Create bookmark | `create_bookmark()` | bookmark name |
| Create heading bookmark | `create_heading_bookmark()` | bookmark name |
| List bookmarks | `list_bookmarks()` | `list[BookmarkInfo]` |
| Get bookmark | `get_bookmark()` | `BookmarkInfo` or `None` |
| List cross-references | `get_cross_references()` | `list[CrossReference]` |
| List available targets | `get_cross_reference_targets()` | `list[CrossReferenceTarget]` |
| Mark fields dirty | `mark_cross_references_dirty()` | count |

## Inserting Cross-References

```python
from python_docx_redline import Document

doc = Document("report.docx")

# Reference a heading
doc.insert_cross_reference("heading:Introduction", after="See ")

# Reference a figure by number
doc.insert_cross_reference("figure:1", display="label_number", after="as shown in ")

# Reference a table by caption text
doc.insert_cross_reference("table:Revenue Data", display="page", after="on page ")

# Reference a footnote
doc.insert_cross_reference("footnote:1", after="see note ")

doc.save("report_with_refs.docx")
```

### Target Formats

| Format | Example | Description |
|--------|---------|-------------|
| Bookmark | `"my_bookmark"` | Direct bookmark reference |
| Heading | `"heading:Introduction"` | Find heading by text (partial match) |
| Figure | `"figure:1"` or `"figure:Architecture"` | By number or caption text |
| Table | `"table:2"` or `"table:Revenue"` | By number or caption text |
| Footnote | `"footnote:1"` | By note number |
| Endnote | `"endnote:2"` | By note number |

### Display Options

| Option | Description | Field Type |
|--------|-------------|------------|
| `"text"` | Target content (default) | REF |
| `"page"` | Page number | PAGEREF |
| `"above_below"` | "above" or "below" relative position | REF with `\p` |
| `"number"` | Heading/caption number only | REF with `\n` |
| `"full_number"` | Full number including chapter | REF with `\w` |
| `"relative_number"` | Number relative to context | REF with `\r` |
| `"label_number"` | "Figure 1" or "Table 2" | REF |
| `"number_only"` | Just "1" or "2" | REF |

### insert_cross_reference() Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `target` | `str` | required | What to reference (see formats above) |
| `display` | `str` | `"text"` | What to display (see options above) |
| `after` | `str` | `None` | Insert after this text |
| `before` | `str` | `None` | Insert before this text |
| `scope` | `str \| dict` | `None` | Limit text search scope |
| `hyperlink` | `bool` | `True` | Make reference clickable |
| `track` | `bool` | `False` | Track as change |
| `author` | `str` | `None` | Author for tracked change |

## Convenience Methods

### Page References

```python
# Insert "see page 5" style reference
doc.insert_page_reference("heading:Conclusion", after="see page ")

# With relative position ("above" or "below" if on same page)
doc.insert_page_reference("figure:1", after="on page ", show_position=True)
```

### Note References

```python
# Reference a footnote
doc.insert_note_reference("footnote", 1, after="see note ")

# Reference an endnote with note formatting
doc.insert_note_reference("endnote", 2, after="refer to ", use_note_style=True)
```

## Bookmark Management

### Creating Bookmarks

```python
# Create a bookmark at specific text
doc.create_bookmark("intro_section", at="1. Introduction")

# Later, reference it
doc.insert_cross_reference("intro_section", after="See ")
```

### Creating Heading Bookmarks

```python
# Auto-generate hidden _Ref bookmark
bookmark_name = doc.create_heading_bookmark("Introduction")
# Returns something like "_Ref12345"

# Or specify custom name
doc.create_heading_bookmark("Conclusion", bookmark_name="my_conclusion")
```

### Listing Bookmarks

```python
# List visible bookmarks only
for bm in doc.list_bookmarks():
    print(f"{bm.name}: {bm.text_preview}")

# Include hidden _Ref bookmarks
for bm in doc.list_bookmarks(include_hidden=True):
    print(f"{bm.name} (hidden: {bm.is_hidden})")
```

### Getting Bookmark Info

```python
bm = doc.get_bookmark("my_bookmark")
if bm:
    print(f"Name: {bm.name}")
    print(f"Location: {bm.location}")
    print(f"Text: {bm.text_preview}")
    print(f"Hidden: {bm.is_hidden}")
```

## Inspection

### List All Cross-References

```python
for xref in doc.get_cross_references():
    print(f"{xref.field_type} -> {xref.target_bookmark}")
    print(f"  Display: {xref.display_value}")
    print(f"  Dirty: {xref.is_dirty}")
    print(f"  Hyperlink: {xref.is_hyperlink}")
```

### List Available Targets

```python
for target in doc.get_cross_reference_targets():
    print(f"{target.type}: {target.display_name}")
    if target.number:
        print(f"  Number: {target.number}")
```

### Mark Fields for Update

```python
# After modifying document content
count = doc.mark_cross_references_dirty()
print(f"Marked {count} cross-references for update")
doc.save("updated.docx")
# Word will recalculate when opening
```

## How Cross-References Work

Cross-references are stored as **field codes** in Word:

```xml
<w:fldChar w:fldCharType="begin" w:dirty="true"/>
<w:instrText> REF _Ref12345 \h </w:instrText>
<w:fldChar w:fldCharType="separate"/>
<w:t>Introduction</w:t>
<w:fldChar w:fldCharType="end"/>
```

### Field Types

| Field | Purpose |
|-------|---------|
| `REF` | Reference bookmark content |
| `PAGEREF` | Reference page number |
| `NOTEREF` | Reference footnote/endnote number |

### Common Switches

| Switch | Meaning |
|--------|---------|
| `\h` | Create hyperlink |
| `\p` | Show "above" or "below" position |
| `\n` | Number only (no context) |
| `\w` | Full number with context |
| `\r` | Relative number |
| `\f` | Footnote reference formatting |

### Why Word Calculates Values

Cross-reference display values depend on:
- Page layout (for page numbers)
- Document structure (for heading numbers)
- Relative position (for above/below)

Python cannot replicate Word's layout engine, so we:
1. Mark fields as dirty (`w:dirty="true"`)
2. Insert placeholder text
3. Word recalculates when opening the document

## Common Patterns

### Reference Multiple Items

```python
doc = Document("report.docx")

# Reference all figures in a summary
doc.insert_cross_reference("figure:1", display="label_number", after="See ")
doc.insert(" and ", after="Figure 1")
doc.insert_cross_reference("figure:2", display="label_number", after=" and ")

doc.save("report.docx")
```

### Create Reference Before Target Exists

```python
# Create bookmark first
doc.create_bookmark("future_section", at="TBD")

# Reference it
doc.insert_cross_reference("future_section", after="See ")

# Later, the bookmark text can be updated
```

### Scoped References

```python
# Only search within a specific section
doc.insert_cross_reference(
    "heading:Results",
    after="See ",
    scope="section:Chapter 2"
)
```
