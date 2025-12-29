# Style Management for python-docx-redline

## Overview and Motivation

python-docx-redline currently lacks a proper Style Management API. This gap is visible in `operations/notes.py` where `_ensure_footnote_styles()` and `_ensure_endnote_styles()` are called but not implemented. These methods need to ensure that required styles like `FootnoteReference` (a character style with superscript) and `FootnoteText` (a paragraph style) exist in `word/styles.xml`.

**Current Pain Points:**
1. No way to programmatically read or create styles
2. Footnote/endnote features are blocked by missing style infrastructure
3. Manual XML manipulation required for any style work
4. python-docx has limited style creation capabilities

**Design Goals:**
1. Follow existing patterns in the library (similar to `RelationshipManager`, `ContentTypeManager`)
2. Provide high-level API that hides OOXML complexity
3. Support both reading and writing styles
4. Enable "ensure exists" pattern for required styles

## Word styles.xml Structure

Word documents store styles in `word/styles.xml`:

```xml
<w:styles xmlns:w="...">
  <!-- Document defaults -->
  <w:docDefaults>
    <w:rPrDefault>...</w:rPrDefault>
    <w:pPrDefault>...</w:pPrDefault>
  </w:docDefaults>

  <!-- Style definitions -->
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>...</w:pPr>
    <w:rPr>...</w:rPr>
  </w:style>

  <w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:rPr>
      <w:vertAlign w:val="superscript"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="FootnoteText">
    <w:name w:val="footnote text"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="FootnoteTextChar"/>
    <w:pPr>...</w:pPr>
    <w:rPr>...</w:rPr>
  </w:style>
</w:styles>
```

**Style Types:**
- `paragraph` - Applied to whole paragraphs
- `character` - Applied to runs of text
- `table` - Applied to tables
- `numbering` - Applied to numbered/bulleted lists

## API Design

### Model Classes

```python
from python_docx_redline.models.style import (
    Style, StyleType, RunFormatting, ParagraphFormatting
)

# StyleType enum
class StyleType(Enum):
    PARAGRAPH = "paragraph"
    CHARACTER = "character"
    TABLE = "table"
    NUMBERING = "numbering"

# RunFormatting (character formatting)
@dataclass
class RunFormatting:
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | str | None = None
    strikethrough: bool | None = None
    font_name: str | None = None
    font_size: float | None = None  # In points
    color: str | None = None  # Hex "#RRGGBB"
    highlight: str | None = None
    superscript: bool | None = None
    subscript: bool | None = None
    small_caps: bool | None = None
    all_caps: bool | None = None

# ParagraphFormatting
@dataclass
class ParagraphFormatting:
    alignment: str | None = None  # "left", "center", "right", "justify"
    spacing_before: float | None = None  # Points
    spacing_after: float | None = None  # Points
    line_spacing: float | None = None  # Multiplier (1.0, 1.5, 2.0)
    indent_left: float | None = None  # Inches
    indent_right: float | None = None  # Inches
    indent_first_line: float | None = None  # Inches
    keep_next: bool | None = None
    keep_lines: bool | None = None
    outline_level: int | None = None  # 0-8 for headings

# Style
@dataclass
class Style:
    style_id: str
    name: str
    style_type: StyleType
    based_on: str | None = None
    next_style: str | None = None
    linked_style: str | None = None
    run_formatting: RunFormatting = field(default_factory=RunFormatting)
    paragraph_formatting: ParagraphFormatting = field(default_factory=ParagraphFormatting)
    ui_priority: int | None = None
    quick_format: bool = False
    semi_hidden: bool = False
    unhide_when_used: bool = False
```

### StyleManager Class

```python
from python_docx_redline import Document
from python_docx_redline.styles import StyleManager
from python_docx_redline.models.style import Style, StyleType, RunFormatting

# Access via Document
doc = Document("contract.docx")
styles = doc.styles

# List all styles
for style in styles:
    print(f"{style.style_id}: {style.name} ({style.style_type.value})")

# List by type
paragraph_styles = styles.list(style_type=StyleType.PARAGRAPH)
character_styles = styles.list(style_type=StyleType.CHARACTER)

# Get specific style
normal = styles.get("Normal")
heading1 = styles.get("Heading1")

# Get by display name
footnote_ref = styles.get_by_name("footnote reference")

# Check if style exists
if "FootnoteReference" in styles:
    print("Style exists")
```

### Creating Styles

```python
# Create a new character style
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

# Create a new paragraph style
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

### Ensuring Styles Exist

The `ensure_style()` method is the primary method for features that require specific styles:

```python
# Ensure a style exists, creating it if necessary
style = styles.ensure_style(
    style_id="FootnoteReference",
    name="footnote reference",
    style_type=StyleType.CHARACTER,
    based_on="DefaultParagraphFont",
    run_formatting=RunFormatting(superscript=True),
    ui_priority=99,
    unhide_when_used=True,
)

# Returns existing style if it exists, or creates and returns new one
```

### Modifying and Removing Styles

```python
# Update an existing style
style = styles.get("Normal")
style.run_formatting.font_size = 12
style.paragraph_formatting.line_spacing = 1.15
styles.update(style)
styles.save()

# Remove a style
styles.remove("MyCustomStyle")
styles.save()
```

## Integration with Footnotes

The `_ensure_footnote_styles()` method in `operations/notes.py`:

```python
def _ensure_footnote_styles(self) -> None:
    """Ensure required footnote styles exist in the document."""
    styles = self._document.styles

    # Ensure FootnoteReference character style (superscript)
    styles.ensure_style(
        style_id="FootnoteReference",
        name="footnote reference",
        style_type=StyleType.CHARACTER,
        based_on="DefaultParagraphFont",
        run_formatting=RunFormatting(superscript=True),
        ui_priority=99,
        unhide_when_used=True,
    )

    # Ensure FootnoteText paragraph style
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

## Predefined Style Templates

Common styles are available as templates:

```python
from python_docx_redline.style_templates import (
    ensure_standard_styles,
    STANDARD_STYLES
)

# Ensure multiple standard styles at once
ensure_standard_styles(
    doc.styles,
    "FootnoteReference",
    "FootnoteText",
    "FootnoteTextChar",
    "EndnoteReference",
    "EndnoteText"
)

# Available standard styles:
# - FootnoteReference (character, superscript)
# - FootnoteText (paragraph)
# - FootnoteTextChar (character, linked to FootnoteText)
# - EndnoteReference (character, superscript)
# - EndnoteText (paragraph)
# - EndnoteTextChar (character, linked to EndnoteText)
# - Hyperlink (character, blue underline)
```

## Implementation Phases

### Phase 1: Core StyleManager (Priority: High)
- Create `models/style.py` with `Style`, `StyleType`, `RunFormatting`, `ParagraphFormatting`
- Create `styles.py` with `StyleManager` class
- Implement `get()`, `list()`, `__contains__()`, `__iter__()`
- Implement `_load()`, `_parse_styles()`, `_element_to_style()`
- Add unit tests for reading styles

### Phase 2: Style Creation (Priority: High)
- Implement `add()`, `_style_to_element()`
- Implement `ensure_style()`
- Implement `save()`
- Add unit tests for creating styles

### Phase 3: Integration (Priority: High)
- Add `styles` property to Document class
- Implement `_ensure_footnote_styles()` and `_ensure_endnote_styles()` in notes.py
- Add `style_templates.py` with predefined styles
- Add integration tests

### Phase 4: Advanced Features (Priority: Medium)
- Implement `update()` and `remove()`
- Add style modification tracking (for tracked changes)
- Add style inheritance resolution
- Add theme color support

## File Structure

```
src/python_docx_redline/
    models/
        style.py           # Style, StyleType, RunFormatting, ParagraphFormatting
    styles.py              # StyleManager class
    style_templates.py     # Predefined standard styles
    __init__.py            # Export Style, StyleType, StyleManager

docs/
    STYLE_MANAGEMENT.md    # This design document
```

## Testing Strategy

```python
def test_list_styles():
    """Test listing styles from a document."""
    doc = Document("tests/fixtures/simple_document.docx")
    styles = doc.styles.list()
    assert len(styles) > 0
    assert any(s.style_id == "Normal" for s in styles)

def test_get_style_by_id():
    """Test getting a style by ID."""
    doc = Document("tests/fixtures/simple_document.docx")
    normal = doc.styles.get("Normal")
    assert normal is not None
    assert normal.style_type == StyleType.PARAGRAPH

def test_ensure_style_creates_if_missing():
    """Test that ensure_style creates missing styles."""
    doc = Document("tests/fixtures/simple_document.docx")

    # Ensure style doesn't exist yet
    assert doc.styles.get("FootnoteReference") is None

    # Ensure it
    style = doc.styles.ensure_style(
        style_id="FootnoteReference",
        name="footnote reference",
        style_type=StyleType.CHARACTER,
        run_formatting=RunFormatting(superscript=True),
    )

    # Should now exist
    assert doc.styles.get("FootnoteReference") is not None
    assert style.run_formatting.superscript is True

def test_footnote_styles_integration():
    """Test that footnote insertion works with style creation."""
    doc = Document("tests/fixtures/simple_document.docx")
    doc.insert_footnote("Test footnote", at="some text")

    # Verify styles were created
    assert doc.styles.get("FootnoteReference") is not None
    assert doc.styles.get("FootnoteText") is not None
```
