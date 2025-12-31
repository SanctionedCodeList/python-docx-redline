# Style Management

Use **python-docx-redline** to programmatically create, read, and manage Word document styles. The StyleManager API lets you ensure required styles exist and create custom styles without OOXML manipulation.

## Overview

| Operation | Method | Use Case |
|-----------|--------|----------|
| **List** styles | `styles.list()` | View all or filtered styles |
| **Get** style | `styles.get(style_id)` | Retrieve existing style |
| **Check** style exists | `if style_id in styles` | Conditional creation |
| **Create/ensure** style | `styles.ensure_style()` | Ensure style exists, creating if needed |
| **Add** new style | `styles.add(style)` | Add a new style |
| **Update** style | `styles.update(style)` | Modify existing style |
| **Remove** style | `styles.remove(style_id)` | Delete a style |

## Basic Operations

### Access Styles

```python
from python_docx_redline import Document

doc = Document("document.docx")
styles = doc.styles
```

### List All Styles

```python
# List all styles
for style in styles.list():
    print(f"{style.style_id}: {style.name} ({style.style_type.value})")

# Filter by type
from python_docx_redline.models.style import StyleType

paragraph_styles = styles.list(style_type=StyleType.PARAGRAPH)
character_styles = styles.list(style_type=StyleType.CHARACTER)
```

### Get a Specific Style

```python
# Get by ID
normal = styles.get("Normal")

# Get by name
footnote_ref = styles.get_by_name("footnote reference")

# Check if style exists
if "FootnoteReference" in styles:
    print("Style exists")
```

## Creating Styles

### Basic Style Creation

```python
from python_docx_redline.models.style import (
    Style, StyleType, RunFormatting, ParagraphFormatting
)

# Create a character style
my_style = Style(
    style_id="MyHighlight",
    name="My Highlight",
    style_type=StyleType.CHARACTER,
    run_formatting=RunFormatting(
        bold=True,
        color="FF0000",
        highlight="yellow"
    )
)
styles.add(my_style)
styles.save()
```

### Paragraph Style with Formatting

```python
# Create a custom paragraph style
custom_para = Style(
    style_id="CustomParagraph",
    name="Custom Paragraph",
    style_type=StyleType.PARAGRAPH,
    based_on="Normal",
    paragraph_formatting=ParagraphFormatting(
        alignment="justify",
        spacing_after=12,
        line_spacing=1.5
    ),
    run_formatting=RunFormatting(
        font_name="Arial",
        font_size=11
    )
)
styles.add(custom_para)
styles.save()
```

## Ensure Style Exists

The `ensure_style()` method is the primary tool for ensuring required styles exist. It returns the existing style if present, or creates and returns a new one.

```python
# Ensure a style exists (creates if missing)
style = styles.ensure_style(
    style_id="FootnoteReference",
    name="footnote reference",
    style_type=StyleType.CHARACTER,
    based_on="DefaultParagraphFont",
    run_formatting=RunFormatting(superscript=True),
    ui_priority=99,
    unhide_when_used=True,
)
```

### Common Use Case: Footnote Styles

```python
# Ensure footnote reference style (superscript character style)
styles.ensure_style(
    style_id="FootnoteReference",
    name="footnote reference",
    style_type=StyleType.CHARACTER,
    based_on="DefaultParagraphFont",
    run_formatting=RunFormatting(superscript=True),
    ui_priority=99,
    unhide_when_used=True,
)

# Ensure footnote text style (paragraph style with smaller font)
styles.ensure_style(
    style_id="FootnoteText",
    name="footnote text",
    style_type=StyleType.PARAGRAPH,
    based_on="Normal",
    linked_style="FootnoteTextChar",
    paragraph_formatting=ParagraphFormatting(
        spacing_after=0,
        line_spacing=1.0,
    ),
    run_formatting=RunFormatting(font_size=10),
    ui_priority=99,
    unhide_when_used=True,
)

styles.save()
```

## Formatting Options

### RunFormatting (Character Formatting)

```python
run_fmt = RunFormatting(
    bold=True,                      # bool or None
    italic=True,                    # bool or None
    underline=True,                 # bool, str, or None
    strikethrough=False,            # bool or None
    font_name="Calibri",            # str or None
    font_size=12.0,                 # float (in points) or None
    color="FF0000",                 # Hex "#RRGGBB" or None
    highlight="yellow",             # str or None
    superscript=True,               # bool or None
    subscript=False,                # bool or None
    small_caps=False,               # bool or None
    all_caps=False,                 # bool or None
)
```

### ParagraphFormatting

```python
para_fmt = ParagraphFormatting(
    alignment="justify",            # "left", "center", "right", "justify" or None
    spacing_before=6.0,             # float (in points) or None
    spacing_after=12.0,             # float (in points) or None
    line_spacing=1.5,               # float (multiplier) or None
    indent_left=0.5,                # float (in inches) or None
    indent_right=0.25,              # float (in inches) or None
    indent_first_line=0.5,          # float (in inches) or None
    keep_next=False,                # bool or None
    keep_lines=False,               # bool or None
    outline_level=0,                # int (0-8 for headings) or None
)
```

## Modifying Styles

### Update Existing Style

```python
# Get and modify
style = styles.get("Normal")
style.run_formatting.font_size = 12
style.paragraph_formatting.line_spacing = 1.15
styles.update(style)
styles.save()
```

### Remove a Style

```python
styles.remove("MyCustomStyle")
styles.save()
```

## Style Types

```python
from python_docx_redline.models.style import StyleType

StyleType.PARAGRAPH      # Applied to whole paragraphs
StyleType.CHARACTER      # Applied to runs of text
StyleType.TABLE          # Applied to tables
StyleType.NUMBERING      # Applied to numbered/bulleted lists
```

## Complete Example

```python
from python_docx_redline import Document
from python_docx_redline.models.style import (
    Style, StyleType, RunFormatting, ParagraphFormatting
)

# Open document
doc = Document("contract.docx")
styles = doc.styles

# Ensure footnote styles exist (for footnote features)
styles.ensure_style(
    style_id="FootnoteReference",
    name="footnote reference",
    style_type=StyleType.CHARACTER,
    based_on="DefaultParagraphFont",
    run_formatting=RunFormatting(superscript=True),
)

styles.ensure_style(
    style_id="FootnoteText",
    name="footnote text",
    style_type=StyleType.PARAGRAPH,
    based_on="Normal",
    paragraph_formatting=ParagraphFormatting(font_size=10),
)

# Create custom style for emphasis
emphasis = Style(
    style_id="CustomEmphasis",
    name="Custom Emphasis",
    style_type=StyleType.CHARACTER,
    run_formatting=RunFormatting(
        bold=True,
        italic=True,
        color="0070C0"
    )
)
styles.add(emphasis)

# Save all changes
styles.save()

# Use the custom style via Word's UI for new content
doc.save("styled_contract.docx")
```

## Style Properties Reference

### Style Dataclass

```python
@dataclass
class Style:
    style_id: str                                    # Unique identifier
    name: str                                        # Display name
    style_type: StyleType                            # PARAGRAPH, CHARACTER, TABLE, NUMBERING
    based_on: str | None = None                      # Parent style
    next_style: str | None = None                    # Style for next paragraph
    linked_style: str | None = None                  # Linked style (character/paragraph pair)
    run_formatting: RunFormatting = default_factory  # Character formatting
    paragraph_formatting: ParagraphFormatting = ...  # Paragraph formatting
    ui_priority: int | None = None                   # UI sort order (0-999)
    quick_format: bool = False                       # Show in styles gallery
    semi_hidden: bool = False                        # Hidden from normal UI
    unhide_when_used: bool = False                   # Appear after first use
```

## When to Use StyleManager

Use StyleManager when you need to:

- **Ensure required styles exist** before using features like footnotes
- **Create custom styles** for specialized formatting
- **Query available styles** in a document
- **Programmatically manage styles** without manual Word editing
- **Prepare documents** for features that require specific styles

StyleManager handles all OOXML complexity—you just work with Python dataclasses.

## Error Handling

```python
from python_docx_redline.errors import StyleNotFoundError

try:
    style = styles.get("NonexistentStyle")
    if style is None:
        print("Style not found")
except Exception as e:
    print(f"Error retrieving style: {e}")
```

## API Reference

### StyleManager Methods

| Method | Description |
|--------|-------------|
| `list(style_type=None)` | List all styles, optionally filtered by type |
| `get(style_id)` | Get style by ID, returns None if not found |
| `get_by_name(name)` | Get style by display name |
| `__contains__(style_id)` | Check if style exists with `if style_id in styles` |
| `__iter__()` | Iterate all styles with `for style in styles` |
| `add(style)` | Add new style |
| `update(style)` | Update existing style |
| `remove(style_id)` | Remove style by ID |
| `ensure_style(**kwargs)` | Ensure style exists, creating if necessary |
| `save()` | Persist all style changes to document |

### Style Dataclass

| Property | Type | Description |
|----------|------|-------------|
| `style_id` | str | Unique identifier |
| `name` | str | Display name in Word UI |
| `style_type` | StyleType | Style category |
| `based_on` | str \| None | Parent style ID |
| `next_style` | str \| None | Style for next paragraph |
| `linked_style` | str \| None | Linked character/paragraph style |
| `run_formatting` | RunFormatting | Character-level formatting |
| `paragraph_formatting` | ParagraphFormatting | Paragraph-level formatting |
| `ui_priority` | int \| None | Sorting in styles UI (0-999) |
| `quick_format` | bool | Include in quick style gallery |
| `semi_hidden` | bool | Hide from normal UI |
| `unhide_when_used` | bool | Appear after first use |
