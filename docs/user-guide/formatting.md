# Formatting

Apply text formatting in tracked insertions and track formatting-only changes.

## Markdown Syntax in Insertions

Use markdown syntax for inline formatting in `insert_tracked()`:

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Bold and italic
doc.insert_tracked(
    "**Important:** See *Smith v. Jones* for precedent",
    after="Section 2.1"
)

# Underline and strikethrough
doc.insert_tracked(
    "This clause is ++mandatory++ and ~~optional~~ required",
    after="Terms"
)

# Combined formatting
doc.insert_tracked(
    "***Critical Notice:*** Review ++immediately++",
    after="Exhibit A"
)

doc.save("formatted.docx")
```

### Supported Markdown

| Syntax | Result | OOXML Element |
|--------|--------|---------------|
| `**text**` | **bold** | `<w:b/>` |
| `*text*` | *italic* | `<w:i/>` |
| `++text++` | underline | `<w:u/>` |
| `~~text~~` | ~~strikethrough~~ | `<w:strike/>` |
| `***text***` | ***bold italic*** | `<w:b/><w:i/>` |

### Line Breaks

Use two spaces followed by newline for line breaks:

```python
doc.insert_tracked(
    "First line  \nSecond line",  # Two spaces before \n
    after="Introduction"
)
```

## Format-Only Tracked Changes

Track formatting changes without modifying text:

```python
# Apply bold formatting with tracking
result = doc.format_tracked(
    find="Important Notice",
    bold=True
)
print(f"Applied bold to '{result.text_matched}'")

# Apply multiple formats
doc.format_tracked(
    find="CONFIDENTIAL",
    bold=True,
    italic=True,
    underline=True
)

# Remove formatting (explicit False)
doc.format_tracked(
    find="previously bold text",
    bold=False  # Removes bold with tracked change
)
```

### FormatResult Details

```python
result = doc.format_tracked(find="text", bold=True)

print(result.success)              # True if operation succeeded
print(result.changed)              # True if formatting actually changed
print(result.text_matched)         # The text that was formatted
print(result.changes_applied)      # {'bold': True}
print(result.previous_formatting)  # Previous state per run
print(result.change_id)            # OOXML change ID for tracking
```

## Paragraph Formatting

Track paragraph-level formatting changes:

```python
doc.format_paragraph_tracked(
    paragraph_index=0,
    alignment="center",
    style="Heading1"
)
```

## Accept/Reject Formatting Changes

Formatting changes are included in accept/reject operations:

```python
# Accept all changes by an author (text AND formatting)
doc.accept_by_author("Claude")

# Reject all changes by an author
doc.reject_by_author("Claude")
```

## Next Steps

- [Basic Operations](basic-operations.md) — Text editing operations
- [Advanced Features](advanced.md) — Scopes, MS365 identity, rendering
- [API Reference](../PROPOSED_API.md) — Complete method documentation
