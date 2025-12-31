# Table of Contents API Specification

## Overview

This document specifies the API for Table of Contents (TOC) support in python_docx_redline.

## Design Principles

1. **Honest about limitations** - Page numbers require Word's layout engine
2. **Pythonic API** - `levels=(1, 3)` instead of raw switches like `\o "1-3"`
3. **Progressive disclosure** - Simple defaults, full control available
4. **Leverage existing infrastructure** - StyleManager, BookmarkRegistry, HyperlinkOperations

## Core API

### Inserting a TOC

```python
from python_docx_redline import Document

doc = Document("document.docx")

# Simple: Insert TOC with defaults
doc.insert_toc()

# Full options
doc.insert_toc(
    # Position
    position="start",              # "start", "end", int index, or "after:bookmark_name"

    # Content
    levels=(1, 3),                 # Heading levels to include (maps to \o switch)
    use_outline_levels=True,       # Include paragraphs with outline levels (\u switch)
    custom_styles=None,            # Dict mapping style names to levels (\t switch)
                                   # e.g., {"ClauseHeading": 2, "SubClause": 3}

    # Formatting
    title="Table of Contents",     # Title text (None for no title)
    title_style="TOCHeading",      # Style for title paragraph
    hyperlinks=True,               # Make entries clickable (\h switch)
    show_page_numbers=True,        # Show page numbers (False = \n switch)
    tab_leader="dot",              # "dot", "hyphen", "underscore", "none"
    right_align_page_numbers=True, # False for inline page numbers

    # Scope
    bookmark=None,                 # Limit TOC to bookmarked section (\b switch)

    # Behavior
    update_on_open=True,           # Set w:updateFields in settings.xml
)
```

### TOC Styles

```python
# Get/modify TOC entry styles
toc_style = doc.styles.get("TOC1")  # Uses existing StyleManager
toc_style.paragraph.indent_left = 0.0
toc_style.paragraph.tab_stops = [
    TabStop(position=6.5, alignment="right", leader="dot")
]
toc_style.run.bold = True
doc.styles.update(toc_style)

# Ensure all TOC styles exist with defaults
doc.ensure_toc_styles(levels=3)  # Creates TOC1, TOC2, TOC3 if missing
```

### Inspecting Existing TOC

```python
# Get existing TOC (returns None if not found)
toc = doc.get_toc()

if toc:
    print(f"TOC found at paragraph {toc.position}")
    print(f"Levels: {toc.levels}")        # e.g., (1, 3)
    print(f"Switches: {toc.switches}")    # Raw switch string
    print(f"Is dirty: {toc.is_dirty}")    # Needs update?

    # Read cached entries (may be stale)
    for entry in toc.entries:
        print(f"L{entry.level}: {entry.text} ... {entry.page_number}")
```

### Modifying TOC

```python
# Mark TOC as needing update
doc.get_toc().mark_dirty()

# Remove TOC entirely
doc.remove_toc()

# Replace TOC with new one
doc.remove_toc()
doc.insert_toc(levels=(1, 4))
```

### Table of Figures / Tables

```python
# Insert Table of Figures
doc.insert_tof(
    caption_label="Figure",        # SEQ field identifier
    position="after:toc",
    title="List of Figures",
    include_label=True,            # Include "Figure 1:" prefix
)

# Insert Table of Tables
doc.insert_tof(
    caption_label="Table",
    title="List of Tables",
)
```

### TC Field Entries (Manual TOC Entries)

```python
# Add custom entry that appears in TOC
doc.insert_tc_entry(
    text="Appendix A: Glossary",
    level=1,
    at="after:glossary_heading",
    identifier=None,               # For filtering with \f switch
)
```

## Data Models

### TOC

```python
@dataclass
class TOC:
    """Represents an existing TOC in the document."""
    position: int                  # Paragraph index
    levels: tuple[int, int]        # (min, max) heading levels
    switches: str                  # Raw field instruction
    is_dirty: bool                 # Marked for update
    entries: list[TOCEntry]        # Cached entries (may be stale)

    def mark_dirty(self) -> None:
        """Mark TOC as needing update."""

    def get_switch(self, name: str) -> str | None:
        """Get value of specific switch (e.g., 'o' returns '1-3')."""
```

### TOCEntry

```python
@dataclass
class TOCEntry:
    """A single entry in a TOC."""
    text: str                      # Entry text
    level: int                     # Hierarchy level (1-9)
    page_number: str | None        # Cached page number (may be stale)
    bookmark: str | None           # Target bookmark name (e.g., "_Toc123456")
    style: str                     # Applied style (e.g., "TOC1")
```

### TabStop

```python
@dataclass
class TabStop:
    """A tab stop definition."""
    position: float                # Position in inches from left margin
    alignment: str                 # "left", "right", "center", "decimal"
    leader: str                    # "dot", "hyphen", "underscore", "none"
```

### TOCStyle (Extension to existing Style)

```python
# Extend ParagraphFormatting to include tab_stops
@dataclass
class ParagraphFormatting:
    # ... existing fields ...
    tab_stops: list[TabStop] | None = None
```

## XML Generation

### TOC Field Structure

```xml
<!-- Title paragraph -->
<w:p>
  <w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>
  <w:r><w:t>Table of Contents</w:t></w:r>
</w:p>

<!-- TOC Field -->
<w:sdt>
  <w:sdtPr>
    <w:docPartGallery w:val="Table of Contents"/>
    <w:docPartUnique/>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
      <w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>Update this table of contents</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>
```

### Settings.xml Addition

```xml
<w:settings>
  <w:updateFields w:val="true"/>
</w:settings>
```

### TOC Style Definition

```xml
<w:style w:type="paragraph" w:styleId="TOC1">
  <w:name w:val="toc 1"/>
  <w:basedOn w:val="Normal"/>
  <w:uiPriority w:val="39"/>
  <w:semiHidden/>
  <w:unhideWhenUsed/>
  <w:pPr>
    <w:tabs>
      <w:tab w:val="right" w:leader="dot" w:pos="9360"/>
    </w:tabs>
    <w:spacing w:after="100"/>
  </w:pPr>
  <w:rPr>
    <w:b/>
  </w:rPr>
</w:style>
```

## Implementation Phases

### Phase 1: Basic Insertion (MVP)
- [ ] `insert_toc()` with core options (levels, title, hyperlinks)
- [ ] Generate valid OOXML field structure with SDT wrapper
- [ ] Set `w:dirty="true"` and `w:updateFields`
- [ ] Ensure TOCHeading style exists
- [ ] Basic tests

### Phase 2: Style Management
- [ ] Add `tab_stops` to `ParagraphFormatting`
- [ ] Create TOC1-TOC9 style templates in `style_templates.py`
- [ ] `ensure_toc_styles()` helper
- [ ] Tab leader support

### Phase 3: TOC Inspection
- [ ] `get_toc()` - parse existing TOC field
- [ ] `TOC` and `TOCEntry` data models
- [ ] Extract cached entries
- [ ] Parse switches

### Phase 4: TOC Manipulation
- [ ] `remove_toc()`
- [ ] `mark_dirty()`
- [ ] Position options ("after:bookmark")

### Phase 5: Extended Features
- [ ] `insert_tof()` for figures/tables
- [ ] `insert_tc_entry()` for custom entries
- [ ] Multiple TOCs support
- [ ] Bookmark scoping (`\b` switch)

### Phase 6: Pre-population (Optional)
- [ ] Scan document for headings
- [ ] Generate bookmarks at headings
- [ ] Create TOC entries with hyperlinks
- [ ] Placeholder page numbers

## Switch Mapping

| Python Parameter | TOC Switch | Default |
|------------------|------------|---------|
| `levels=(1, 3)` | `\o "1-3"` | `(1, 3)` |
| `use_outline_levels=True` | `\u` | `True` |
| `hyperlinks=True` | `\h` | `True` |
| `show_page_numbers=False` | `\n` | `True` |
| `custom_styles={"Style": 1}` | `\t "Style,1"` | `None` |
| `bookmark="name"` | `\b "name"` | `None` |
| (web view setting) | `\z` | Always included |

## Error Handling

```python
class TOCError(Exception):
    """Base exception for TOC operations."""

class TOCNotFoundError(TOCError):
    """No TOC found in document."""

class TOCAlreadyExistsError(TOCError):
    """TOC already exists (use replace=True to overwrite)."""

class InvalidTOCPositionError(TOCError):
    """Invalid position for TOC insertion."""
```

## Examples

### Legal Document with Custom Styles

```python
doc = Document("contract.docx")

# Map contract-specific styles to TOC levels
doc.insert_toc(
    title="Table of Contents",
    levels=(1, 4),
    custom_styles={
        "ArticleHeading": 1,
        "SectionHeading": 2,
        "ClauseHeading": 3,
        "SubClauseHeading": 4,
    },
    tab_leader="dot",
)
```

### Academic Paper with Multiple TOCs

```python
doc = Document("thesis.docx")

# Main TOC
doc.insert_toc(
    position="after:title_page",
    title="Contents",
    levels=(1, 3),
)

# List of Figures
doc.insert_tof(
    position="after:toc",
    caption_label="Figure",
    title="List of Figures",
)

# List of Tables
doc.insert_tof(
    position="after:lof",
    caption_label="Table",
    title="List of Tables",
)
```

### Minimal TOC (No Styling)

```python
doc = Document("simple.docx")
doc.insert_toc(
    title=None,
    show_page_numbers=False,
    tab_leader="none",
)
```

## Dependencies

### Existing Infrastructure to Leverage
- `StyleManager` - Style creation and modification
- `BookmarkRegistry` - Bookmark creation for hyperlinks
- `HyperlinkOperations` - Internal hyperlink generation
- `ParagraphFormatting` - Paragraph properties
- `format_builder.py` - XML generation utilities

### New Infrastructure Needed
- `TabStop` model and XML generation
- `settings.xml` access (for `updateFields`)
- SDT (Structured Document Tag) generation
- Field code parsing utilities

## Testing Strategy

1. **Unit tests**: XML generation, switch building, style creation
2. **Integration tests**: Full TOC insertion, save/reload, Word compatibility
3. **Roundtrip tests**: Insert TOC, save, open in Word, verify structure
4. **Edge cases**: Empty document, no headings, deeply nested, RTL text

## References

- [OOXML TOC Specification](http://www.officeopenxml.com/WPtableOfContents.php)
- [Eric White's TOC Blog Series](http://www.ericwhite.com/blog/exploring-tables-of-contents-in-open-xml-wordprocessingml-documents-part-2/)
- [Aspose.Words TOC API](https://docs.aspose.com/words/net/working-with-table-of-contents/)
- [docx4j TocGenerator](https://javadoc.io/static/org.docx4j/docx4j-core/11.4.7/org/docx4j/toc/TocGenerator.html)
