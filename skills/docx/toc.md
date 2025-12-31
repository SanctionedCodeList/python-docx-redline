# Table of Contents

Use **python-docx-redline** to insert, inspect, update, and remove Tables of Contents. The TOC uses Word's field codes, so page numbers are calculated by Word when the document is opened.

## Quick Reference

| Operation | Method | Returns |
|-----------|--------|---------|
| Create TOC | `insert_toc()` | None |
| Inspect TOC | `get_toc()` | `TOC` object or `None` |
| Modify TOC | `update_toc()` | `bool` |
| Remove TOC | `remove_toc()` | `bool` |
| Flag for update | `mark_toc_dirty()` | `bool` |

## Creating a TOC

```python
from python_docx_redline import Document

doc = Document("report.docx")

# Basic TOC with defaults (levels 1-3, with title)
doc.insert_toc()

# Customized TOC
doc.insert_toc(
    levels=(1, 5),              # Include Heading 1 through Heading 5
    title="Contents",           # Title above TOC (None for no title)
    hyperlinks=True,            # Clickable links in Word
    show_page_numbers=True,     # Show page numbers
    position="start",           # "start", "end", or paragraph index
)

doc.save("report_with_toc.docx")
```

### insert_toc() Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `position` | `int \| str` | `0` | Where to insert: `"start"`, `"end"`, or paragraph index |
| `levels` | `tuple[int, int]` | `(1, 3)` | Heading levels to include (min, max) |
| `title` | `str \| None` | `"Table of Contents"` | Title text, or `None` for no title |
| `hyperlinks` | `bool` | `True` | Create clickable links to headings |
| `show_page_numbers` | `bool` | `True` | Display page numbers |
| `use_outline_levels` | `bool` | `True` | Use outline levels for hierarchy |
| `update_on_open` | `bool` | `True` | Auto-update TOC when document opens |

## Inspecting an Existing TOC

```python
doc = Document("report.docx")
toc = doc.get_toc()

if toc:
    print(f"Position: paragraph {toc.position}")
    print(f"Levels: {toc.levels}")          # e.g., (1, 3)
    print(f"Is dirty: {toc.is_dirty}")      # True = needs update
    print(f"Switches: {toc.switches}")      # Raw field instruction

    # Query specific switches
    print(f"Hyperlinks: {toc.get_switch('h') is not None}")
    print(f"Outline levels: {toc.get_switch('o')}")  # e.g., "1-3"

    # Read cached entries (may be stale if document was modified)
    for entry in toc.entries:
        print(f"  L{entry.level}: {entry.text} ... p.{entry.page_number}")
else:
    print("No TOC found")
```

### TOC Object Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `position` | `int` | Paragraph index in document body |
| `levels` | `tuple[int, int]` | (min_level, max_level) from `\o` switch |
| `switches` | `str` | Raw field instruction string |
| `is_dirty` | `bool` | Whether TOC is flagged for update |
| `entries` | `list[TOCEntry]` | Cached entries (text, level, page, bookmark) |

### TOCEntry Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `text` | `str` | Heading text |
| `level` | `int` | Heading level (1-9) |
| `page_number` | `str \| None` | Cached page number |
| `bookmark` | `str \| None` | Target bookmark name |
| `style` | `str` | Applied style (e.g., "TOC1") |

### get_switch() Method

Query specific TOC field switches:

```python
toc = doc.get_toc()

# Switches with values return the value
toc.get_switch("o")   # "1-3" (outline levels)
toc.get_switch("t")   # "CustomStyle,1" (custom styles)

# Switches without values return empty string
toc.get_switch("h")   # "" (hyperlinks enabled)
toc.get_switch("z")   # "" (hide in web view)

# Missing switches return None
toc.get_switch("b")   # None (bookmark switch not used)
```

## Updating an Existing TOC

Modify TOC settings without removing and re-inserting:

```python
doc = Document("report.docx")

# Change heading levels
doc.update_toc(levels=(1, 5))

# Toggle hyperlinks off
doc.update_toc(hyperlinks=False)

# Change multiple settings
doc.update_toc(
    levels=(1, 4),
    show_page_numbers=True,
    title="Table of Contents",
)

# Change the title
doc.update_toc(title="Contents")

# Remove the title entirely
doc.update_toc(title=None)

doc.save("updated.docx")
```

### update_toc() Parameters

All parameters are optional. Only provided values are updated; others preserve existing settings.

| Parameter | Type | Description |
|-----------|------|-------------|
| `levels` | `tuple[int, int] \| None` | Change heading levels |
| `hyperlinks` | `bool \| None` | Toggle `\h` switch |
| `show_page_numbers` | `bool \| None` | Toggle `\n` switch |
| `use_outline_levels` | `bool \| None` | Toggle `\u` switch |
| `title` | `str \| None` | Change title, or `None` to remove |

**Note:** If `title` parameter is not provided, the existing title is preserved. Use `title=None` explicitly to remove the title.

## Removing a TOC

```python
doc = Document("report.docx")

if doc.remove_toc():
    print("TOC removed")
else:
    print("No TOC found")

doc.save("report_no_toc.docx")
```

The `remove_toc()` method removes both the TOC and its title paragraph (if present).

## Marking TOC for Update

After modifying document headings, mark the TOC as "dirty" so Word recalculates it:

```python
doc = Document("report.docx")

# Make changes to headings
doc.replace("Introduction", "Executive Summary")

# Mark TOC for update
doc.mark_toc_dirty()

doc.save("report.docx")
# Word will recalculate page numbers when opening the document
```

## How TOC Works in Word

The TOC is stored as a **field** wrapped in a **Structured Document Tag (SDT)**:

1. **Field instruction**: Contains switches like `TOC \o "1-3" \h \z \u`
2. **Cached content**: Text and page numbers from last update
3. **Dirty flag**: When `true`, Word recalculates on open

### Field Switches

| Switch | Meaning |
|--------|---------|
| `\o "1-3"` | Include heading levels 1-3 |
| `\h` | Create hyperlinks |
| `\z` | Hide tab leaders and page numbers in web view |
| `\u` | Use outline levels |
| `\n` | Omit page numbers |

### Why Page Numbers Require Word

Page numbers depend on Word's layout engine, which calculates:
- Font metrics and line breaks
- Page breaks and section breaks
- Headers, footers, and margins

Python cannot replicate this, so we:
1. Set `update_on_open=True` (adds `<w:updateFields>` to settings.xml)
2. Mark the TOC as dirty (`w:dirty="true"`)
3. Let Word calculate page numbers when the document opens

## Common Patterns

### Check if Document Has TOC

```python
if doc.get_toc() is not None:
    print("Document has a TOC")
```

### Replace TOC with Different Settings

```python
# Remove existing and insert new
doc.remove_toc()
doc.insert_toc(levels=(1, 5), title="Contents")
```

### Add TOC to Template

```python
doc = Document("template.docx")

# Populate template
doc.replace("{{TITLE}}", "Annual Report 2024")
doc.replace("{{AUTHOR}}", "Finance Team")

# Add TOC at start
doc.insert_toc(position="start", title="Contents")

doc.save("report.docx")
```

### Ensure TOC is Up to Date

```python
doc = Document("report.docx")

toc = doc.get_toc()
if toc and not toc.is_dirty:
    # TOC exists but not flagged for update
    doc.mark_toc_dirty()

doc.save("report.docx")
```
