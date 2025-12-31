# Layout API Design Document

## Overview

### Problem Statement

python-docx provides basic layout features (margins, orientation, page size) but lacks:
- **Multi-column layouts** - Creating newspaper-style or academic paper layouts
- **Page borders** - Adding decorative or functional borders around pages
- **Line numbering** - Essential for legal documents and collaborative editing
- **Tracked change support** - No way to track layout modifications for review

python-docx-redline currently inherits python-docx capabilities but doesn't add tracked change support for layout modifications. Users who need to:
1. Programmatically create multi-column layouts
2. Add page borders to documents
3. Enable line numbering for legal or academic documents
4. Track these layout changes for review workflows

...must currently manipulate raw OOXML XML directly, managing complex `<w:sectPr>` (section properties) elements and the `<w:sectPrChange>` tracked change markup.

### Solution

Provide a high-level `LayoutOperations` class that:
- Wraps python-docx layout features with tracked change support
- Adds new capabilities for multi-column layouts, page borders, and line numbering
- Follows existing patterns from `NoteOperations`, `HyperlinkOperations`
- Uses `<w:sectPrChange>` for proper OOXML tracked changes

### Impact

```python
# BEFORE (manual XML construction)
from lxml import etree
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
sect_pr = doc.xml_root.find(".//w:sectPr", ns)
cols = etree.SubElement(sect_pr, "{http://...}cols")
cols.set("{http://...}num", "2")
cols.set("{http://...}space", "720")
# ... manually construct sectPrChange for tracking ...

# AFTER (proposed API)
doc.set_columns(count=2, spacing=0.5, track=True)
```

**Time savings**: 5-10x faster, more readable, built-in validation.

---

## API Design

### 1. Section Layout Operations (Enhanced python-docx with Tracked Changes)

#### set_margins

```python
def set_margins(
    self,
    top: float | None = None,
    bottom: float | None = None,
    left: float | None = None,
    right: float | None = None,
    gutter: float | None = None,
    header: float | None = None,
    footer: float | None = None,
    section_index: int | None = None,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Set page margins for a section.

    Margins are specified in inches. Only provided values are changed;
    unspecified margins retain their current values.

    Args:
        top: Top margin in inches
        bottom: Bottom margin in inches
        left: Left margin in inches
        right: Right margin in inches
        gutter: Gutter margin in inches (for binding)
        header: Distance from top of page to header in inches
        footer: Distance from bottom of page to footer in inches
        section_index: Section to modify (None = last/default section)
        track: If True, record the change as a tracked revision
        author: Optional author override for tracked changes

    Raises:
        IndexError: If section_index is out of range

    Example:
        >>> # Set 1-inch margins all around
        >>> doc.set_margins(top=1.0, bottom=1.0, left=1.0, right=1.0)

        >>> # Set only left/right for wider text area
        >>> doc.set_margins(left=0.75, right=0.75)

        >>> # Track the change for review
        >>> doc.set_margins(left=1.5, right=1.5, track=True)
    """
```

#### set_orientation

```python
def set_orientation(
    self,
    orientation: str,
    section_index: int | None = None,
    preserve_margins: bool = True,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Set page orientation.

    Args:
        orientation: "portrait" or "landscape"
        section_index: Section to modify (None = last/default section)
        preserve_margins: If True, swap left/right with top/bottom when
            changing orientation to maintain relative spacing
        track: If True, record the change as a tracked revision
        author: Optional author override for tracked changes

    Raises:
        ValueError: If orientation is not "portrait" or "landscape"
        IndexError: If section_index is out of range

    Example:
        >>> doc.set_orientation("landscape")
        >>> doc.set_orientation("portrait", track=True)
    """
```

#### set_page_size

```python
def set_page_size(
    self,
    width: float | None = None,
    height: float | None = None,
    preset: str | None = None,
    section_index: int | None = None,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Set page size.

    Specify either width/height in inches OR use a preset name.

    Args:
        width: Page width in inches
        height: Page height in inches
        preset: Preset page size name. Options:
            - "letter" (8.5 x 11 inches, US standard)
            - "legal" (8.5 x 14 inches, US legal)
            - "a4" (210 x 297 mm, ISO standard)
            - "a3" (297 x 420 mm)
            - "a5" (148 x 210 mm)
            - "executive" (7.25 x 10.5 inches)
            - "tabloid" (11 x 17 inches)
        section_index: Section to modify (None = last/default section)
        track: If True, record the change as a tracked revision
        author: Optional author override for tracked changes

    Raises:
        ValueError: If both width/height and preset specified, or neither
        ValueError: If preset name is not recognized
        IndexError: If section_index is out of range

    Example:
        >>> doc.set_page_size(preset="letter")
        >>> doc.set_page_size(preset="a4", track=True)
        >>> doc.set_page_size(width=8.5, height=14)  # Custom legal size
    """
```

#### set_header_footer_distance

```python
def set_header_footer_distance(
    self,
    header_distance: float | None = None,
    footer_distance: float | None = None,
    section_index: int | None = None,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Set header and footer distance from page edge.

    Args:
        header_distance: Distance from top of page to header in inches
        footer_distance: Distance from bottom of page to footer in inches
        section_index: Section to modify (None = last/default section)
        track: If True, record the change as a tracked revision
        author: Optional author override for tracked changes

    Raises:
        IndexError: If section_index is out of range

    Example:
        >>> doc.set_header_footer_distance(header_distance=0.5)
        >>> doc.set_header_footer_distance(footer_distance=0.5, track=True)
    """
```

---

### 2. Multi-Column Layout (New Capability)

#### set_columns

```python
def set_columns(
    self,
    count: int,
    spacing: float = 0.5,
    equal_width: bool = True,
    separator: bool = False,
    column_widths: list[float] | None = None,
    section_index: int | None = None,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Set multi-column layout for a section.

    Creates a multi-column layout. By default, columns are equal width.
    For unequal columns, set equal_width=False and provide column_widths.

    Args:
        count: Number of columns (1-45)
        spacing: Space between columns in inches (default 0.5")
        equal_width: If True, all columns have equal width (default)
        separator: If True, draw a vertical line between columns
        column_widths: List of column widths in inches (required if
            equal_width=False). Must have exactly `count` elements.
        section_index: Section to modify (None = last/default section)
        track: If True, record the change as a tracked revision
        author: Optional author override for tracked changes

    Raises:
        ValueError: If count < 1 or count > 45
        ValueError: If equal_width=False but column_widths not provided
        ValueError: If len(column_widths) != count
        IndexError: If section_index is out of range

    Example:
        >>> # Two equal columns
        >>> doc.set_columns(count=2)

        >>> # Three columns with separator lines
        >>> doc.set_columns(count=3, separator=True)

        >>> # Newspaper-style: wide left, narrow right
        >>> doc.set_columns(
        ...     count=2,
        ...     equal_width=False,
        ...     column_widths=[4.5, 2.5],
        ...     spacing=0.25
        ... )

        >>> # Track the layout change
        >>> doc.set_columns(count=2, track=True)
    """
```

#### get_columns

```python
def get_columns(
    self,
    section_index: int | None = None,
) -> ColumnInfo:
    """Get column layout information for a section.

    Args:
        section_index: Section to query (None = last/default section)

    Returns:
        ColumnInfo dataclass with column configuration

    Raises:
        IndexError: If section_index is out of range

    Example:
        >>> info = doc.get_columns()
        >>> print(f"Document has {info.count} columns")
        >>> if info.count > 1:
        ...     print(f"Column spacing: {info.spacing} inches")
    """
```

#### add_column_break

```python
def add_column_break(
    self,
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Any | None = None,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Insert a column break at a specific location.

    Forces text to continue in the next column. In single-column
    layouts, this behaves like a page break.

    Args:
        after: Insert column break after this text
        before: Insert column break before this text
        scope: Limit search scope for anchor text
        track: If True, wrap in tracked insertion
        author: Optional author override for tracked changes

    Raises:
        ValueError: If neither after nor before specified, or both
        TextNotFoundError: If anchor text not found
        AmbiguousTextError: If anchor text found multiple times

    Example:
        >>> doc.add_column_break(after="End of first column content.")
        >>> doc.add_column_break(before="Start of new column", track=True)
    """
```

---

### 3. Page Borders (New Capability)

#### set_page_border

```python
def set_page_border(
    self,
    style: str = "single",
    color: str = "000000",
    width: float = 0.5,
    sides: str | list[str] = "all",
    offset_from: str = "text",
    art: str | None = None,
    section_index: int | None = None,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Add a border around the page.

    Args:
        style: Border style. Options:
            - "single" (default), "double", "triple"
            - "dashed", "dotted", "dotDash", "dotDotDash"
            - "wave", "doubleWave"
            - "thick", "thickThinSmallGap", "thinThickSmallGap"
            - "3d", "inset", "outset"
            - See OOXML ST_Border for full list
        color: Border color as hex string (e.g., "FF0000" for red)
        width: Border width in points (1/72 inch)
        sides: Which sides to apply border. Options:
            - "all" (default) - all four sides
            - "box" - same as "all"
            - List of sides: ["top", "bottom", "left", "right"]
        offset_from: Where to measure border offset from:
            - "text" (default) - offset from text area
            - "page" - offset from page edge
        art: Art border type (e.g., "apples", "stars"). Overrides style.
        section_index: Section to modify (None = last/default section)
        track: If True, record the change as a tracked revision
        author: Optional author override for tracked changes

    Raises:
        ValueError: If style is not recognized
        ValueError: If color is not valid hex
        IndexError: If section_index is out of range

    Example:
        >>> # Simple black single-line border
        >>> doc.set_page_border()

        >>> # Double red border
        >>> doc.set_page_border(style="double", color="FF0000")

        >>> # Top and bottom only
        >>> doc.set_page_border(sides=["top", "bottom"])

        >>> # Art border for certificates
        >>> doc.set_page_border(art="stars")

        >>> # Track the change
        >>> doc.set_page_border(style="single", color="0000FF", track=True)
    """
```

#### remove_page_border

```python
def remove_page_border(
    self,
    sides: str | list[str] = "all",
    section_index: int | None = None,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Remove page border.

    Args:
        sides: Which sides to remove border from:
            - "all" (default) - remove all borders
            - List of sides: ["top", "bottom", "left", "right"]
        section_index: Section to modify (None = last/default section)
        track: If True, record the change as a tracked revision
        author: Optional author override for tracked changes

    Raises:
        IndexError: If section_index is out of range

    Example:
        >>> doc.remove_page_border()
        >>> doc.remove_page_border(sides=["top", "bottom"], track=True)
    """
```

#### get_page_border

```python
def get_page_border(
    self,
    section_index: int | None = None,
) -> BorderInfo | None:
    """Get page border information for a section.

    Args:
        section_index: Section to query (None = last/default section)

    Returns:
        BorderInfo dataclass with border configuration, or None if no border

    Raises:
        IndexError: If section_index is out of range

    Example:
        >>> info = doc.get_page_border()
        >>> if info:
        ...     print(f"Border style: {info.style}, color: {info.color}")
    """
```

---

### 4. Line Numbering (New Capability)

#### set_line_numbering

```python
def set_line_numbering(
    self,
    start: int = 1,
    count_by: int = 1,
    restart: str = "page",
    distance: float | None = None,
    section_index: int | None = None,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Enable line numbering for a section.

    Line numbering is essential for legal documents and collaborative
    editing where specific lines need to be referenced.

    Args:
        start: Starting line number (default 1)
        count_by: Only show numbers for every Nth line (default 1 = every line)
        restart: When to restart numbering:
            - "page" (default) - restart on each page
            - "section" - restart on each section
            - "continuous" - continue throughout document
        distance: Distance from text to line numbers in inches
            (default: use document default)
        section_index: Section to modify (None = last/default section)
        track: If True, record the change as a tracked revision
        author: Optional author override for tracked changes

    Raises:
        ValueError: If restart is not "page", "section", or "continuous"
        IndexError: If section_index is out of range

    Example:
        >>> # Standard line numbering (every line, restart per page)
        >>> doc.set_line_numbering()

        >>> # Legal style: every 5th line, restart per page
        >>> doc.set_line_numbering(count_by=5)

        >>> # Continuous numbering throughout document
        >>> doc.set_line_numbering(restart="continuous")

        >>> # Track the change
        >>> doc.set_line_numbering(count_by=1, track=True)
    """
```

#### remove_line_numbering

```python
def remove_line_numbering(
    self,
    section_index: int | None = None,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Remove line numbering from a section.

    Args:
        section_index: Section to modify (None = last/default section)
        track: If True, record the change as a tracked revision
        author: Optional author override for tracked changes

    Raises:
        IndexError: If section_index is out of range

    Example:
        >>> doc.remove_line_numbering()
        >>> doc.remove_line_numbering(track=True)
    """
```

#### get_line_numbering

```python
def get_line_numbering(
    self,
    section_index: int | None = None,
) -> LineNumberingInfo | None:
    """Get line numbering configuration for a section.

    Args:
        section_index: Section to query (None = last/default section)

    Returns:
        LineNumberingInfo dataclass, or None if line numbering is disabled

    Raises:
        IndexError: If section_index is out of range

    Example:
        >>> info = doc.get_line_numbering()
        >>> if info:
        ...     print(f"Numbering every {info.count_by} lines")
    """
```

---

### 5. Section Breaks

#### add_section_break

```python
def add_section_break(
    self,
    break_type: str = "next_page",
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Any | None = None,
    track: bool = False,
    author: str | None = None,
) -> int:
    """Insert a section break at a specific location.

    Section breaks allow different parts of a document to have
    different layouts (margins, orientation, columns, etc.).

    Args:
        break_type: Type of section break:
            - "next_page" (default) - new section starts on next page
            - "continuous" - new section starts immediately (same page)
            - "even_page" - new section starts on next even page
            - "odd_page" - new section starts on next odd page
        after: Insert section break after this text
        before: Insert section break before this text
        scope: Limit search scope for anchor text
        track: If True, wrap in tracked insertion
        author: Optional author override for tracked changes

    Returns:
        Index of the new section (0-based)

    Raises:
        ValueError: If neither after nor before specified, or both
        ValueError: If break_type is not recognized
        TextNotFoundError: If anchor text not found
        AmbiguousTextError: If anchor text found multiple times

    Example:
        >>> # Start new section on next page
        >>> section_idx = doc.add_section_break(after="End of Chapter 1")

        >>> # Continuous section break for column layout change
        >>> section_idx = doc.add_section_break(
        ...     break_type="continuous",
        ...     before="Start of two-column content"
        ... )
        >>> doc.set_columns(count=2, section_index=section_idx)

        >>> # Track the break insertion
        >>> doc.add_section_break(after="Chapter end", track=True)
    """
```

#### get_sections

```python
@property
def sections(self) -> list[SectionLayout]:
    """Get all document sections with their layout properties.

    Returns:
        List of SectionLayout objects

    Example:
        >>> for i, section in enumerate(doc.sections):
        ...     print(f"Section {i}: {section.page_size}, {section.orientation}")
    """
```

---

## Data Models

```python
from dataclasses import dataclass, field
from typing import Literal


@dataclass
class ColumnInfo:
    """Information about column layout in a section.

    Attributes:
        count: Number of columns
        spacing: Space between columns in inches
        equal_width: Whether all columns have equal width
        separator: Whether separator lines are shown between columns
        column_widths: List of individual column widths (empty if equal_width=True)
    """
    count: int
    spacing: float
    equal_width: bool
    separator: bool
    column_widths: list[float] = field(default_factory=list)


@dataclass
class BorderSide:
    """Border configuration for one side of the page.

    Attributes:
        style: Border style (e.g., "single", "double", "dashed")
        color: Border color as hex string
        width: Border width in points
        space: Space between border and content in points
    """
    style: str
    color: str
    width: float
    space: float = 0


@dataclass
class BorderInfo:
    """Information about page borders in a section.

    Attributes:
        top: Top border configuration (None if no border)
        bottom: Bottom border configuration (None if no border)
        left: Left border configuration (None if no border)
        right: Right border configuration (None if no border)
        offset_from: Where border offset is measured from ("text" or "page")
        z_order: Whether border is in front of or behind text ("front" or "back")
    """
    top: BorderSide | None = None
    bottom: BorderSide | None = None
    left: BorderSide | None = None
    right: BorderSide | None = None
    offset_from: Literal["text", "page"] = "text"
    z_order: Literal["front", "back"] = "front"


@dataclass
class LineNumberingInfo:
    """Information about line numbering in a section.

    Attributes:
        start: Starting line number
        count_by: Interval for displaying numbers (1 = every line)
        restart: When numbering restarts ("page", "section", "continuous")
        distance: Distance from text to line numbers in inches
    """
    start: int
    count_by: int
    restart: Literal["page", "section", "continuous"]
    distance: float | None = None


@dataclass
class PageSize:
    """Page dimensions.

    Attributes:
        width: Page width in inches
        height: Page height in inches
        name: Optional preset name (e.g., "letter", "a4")
    """
    width: float
    height: float
    name: str | None = None


@dataclass
class PageMargins:
    """Page margin settings.

    Attributes:
        top: Top margin in inches
        bottom: Bottom margin in inches
        left: Left margin in inches
        right: Right margin in inches
        gutter: Gutter margin in inches
        header: Header distance from page edge in inches
        footer: Footer distance from page edge in inches
    """
    top: float
    bottom: float
    left: float
    right: float
    gutter: float = 0
    header: float = 0.5
    footer: float = 0.5


@dataclass
class SectionLayout:
    """Complete layout information for a document section.

    Attributes:
        index: Section index (0-based)
        page_size: Page size configuration
        orientation: "portrait" or "landscape"
        margins: Page margin settings
        columns: Column layout information
        line_numbering: Line numbering configuration (None if disabled)
        border: Page border configuration (None if no border)
        break_type: Type of break before this section
            ("next_page", "continuous", "even_page", "odd_page")
    """
    index: int
    page_size: PageSize
    orientation: Literal["portrait", "landscape"]
    margins: PageMargins
    columns: ColumnInfo
    line_numbering: LineNumberingInfo | None = None
    border: BorderInfo | None = None
    break_type: str = "next_page"
```

---

## OOXML Structure

### Section Properties (`<w:sectPr>`)

Section properties are stored in `<w:sectPr>` elements. The document body has one final `<w:sectPr>` for the last section, and additional sections are defined by `<w:sectPr>` elements within `<w:pPr>` of paragraphs that end sections.

```xml
<w:sectPr>
    <!-- Page size -->
    <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>

    <!-- Page margins (all values in twips: 1 inch = 1440 twips) -->
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
             w:header="720" w:footer="720" w:gutter="0"/>

    <!-- Columns -->
    <w:cols w:num="2" w:space="720" w:equalWidth="true" w:sep="false">
        <!-- Optional: individual column definitions when equalWidth="false" -->
        <w:col w:w="3600" w:space="720"/>
        <w:col w:w="3600"/>
    </w:cols>

    <!-- Page borders -->
    <w:pgBorders w:offsetFrom="text" w:zOrder="front">
        <w:top w:val="single" w:sz="4" w:space="24" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="24" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="24" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="24" w:color="000000"/>
    </w:pgBorders>

    <!-- Line numbering -->
    <w:lnNumType w:countBy="1" w:start="1" w:restart="newPage" w:distance="360"/>

    <!-- Section type (break type) -->
    <w:type w:val="nextPage"/>
</w:sectPr>
```

### Tracked Changes for Section Properties (`<w:sectPrChange>`)

When layout changes are tracked, OOXML uses `<w:sectPrChange>` to store the original values:

```xml
<w:sectPr>
    <!-- Current (new) values -->
    <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"
             w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:num="2" w:space="720"/>

    <!-- Tracked change: stores the ORIGINAL values before modification -->
    <w:sectPrChange w:id="1" w:author="Jane Doe" w:date="2025-12-29T10:30:00Z">
        <w:sectPr>
            <!-- Original values (single column, 1" margins) -->
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                     w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:num="1"/>
        </w:sectPr>
    </w:sectPrChange>
</w:sectPr>
```

### Column Break

```xml
<w:r>
    <w:br w:type="column"/>
</w:r>
```

### Section Break (within paragraph)

Section breaks are implemented by placing `<w:sectPr>` inside the paragraph properties of the last paragraph before the break:

```xml
<w:p>
    <w:pPr>
        <w:sectPr>
            <w:type w:val="continuous"/>
            <!-- Other section properties for the PREVIOUS section -->
        </w:sectPr>
    </w:pPr>
    <w:r><w:t>Last paragraph of previous section.</w:t></w:r>
</w:p>
```

---

## Implementation Phases

### Phase 1: Core Infrastructure and Reading (Priority: High)

**Goal**: Establish the LayoutOperations class and ability to read layout properties.

**Tasks**:
1. Create `src/python_docx_redline/operations/layout.py` with `LayoutOperations` class
2. Create `src/python_docx_redline/models/layout.py` with dataclasses
3. Implement `get_columns()`, `get_page_border()`, `get_line_numbering()`
4. Implement `sections` property to list all sections with their layout info
5. Add `_get_section_by_index()` helper method
6. Add `sections` property to Document class
7. Write unit tests for reading layout properties

**Estimated Effort**: 2-3 days

### Phase 2: Basic Layout Modification (Priority: High)

**Goal**: Enable setting margins, orientation, and page size WITHOUT tracked changes.

**Tasks**:
1. Implement `set_margins()` (track=False only initially)
2. Implement `set_orientation()` (track=False only initially)
3. Implement `set_page_size()` with preset support
4. Implement `set_header_footer_distance()`
5. Add unit conversion helpers (inches to twips, etc.)
6. Write unit tests for layout modification

**Estimated Effort**: 2-3 days

### Phase 3: Multi-Column Layout (Priority: High)

**Goal**: Add full column layout support.

**Tasks**:
1. Implement `set_columns()` with equal/unequal column support
2. Implement `add_column_break()`
3. Write unit tests for column operations

**Estimated Effort**: 2 days

### Phase 4: Page Borders (Priority: Medium)

**Goal**: Add page border support.

**Tasks**:
1. Implement `set_page_border()` with all style options
2. Implement `remove_page_border()`
3. Add art border support
4. Write unit tests for border operations

**Estimated Effort**: 2 days

### Phase 5: Line Numbering (Priority: Medium)

**Goal**: Add line numbering support.

**Tasks**:
1. Implement `set_line_numbering()`
2. Implement `remove_line_numbering()`
3. Write unit tests for line numbering

**Estimated Effort**: 1 day

### Phase 6: Section Breaks (Priority: Medium)

**Goal**: Enable creating new sections in documents.

**Tasks**:
1. Implement `add_section_break()` with all break types
2. Handle section property inheritance for new sections
3. Write unit tests for section break operations

**Estimated Effort**: 2 days

### Phase 7: Tracked Changes Support (Priority: Medium)

**Goal**: Add `track=True` support to all layout operations.

**Tasks**:
1. Create `_create_sect_pr_change()` helper method
2. Add tracking to `set_margins()`
3. Add tracking to `set_orientation()` and `set_page_size()`
4. Add tracking to `set_columns()`
5. Add tracking to `set_page_border()` and `remove_page_border()`
6. Add tracking to `set_line_numbering()` and `remove_line_numbering()`
7. Add tracking to `add_section_break()`
8. Write unit tests for tracked changes

**Estimated Effort**: 3-4 days

### Phase 8: Documentation and Integration (Priority: Low)

**Goal**: Complete documentation and polish.

**Tasks**:
1. Add skill guide `skills/docx/layout.md`
2. Update SKILL.md with layout capabilities
3. Add examples to documentation
4. Write integration tests with real documents

**Estimated Effort**: 1-2 days

---

## Key Implementation Details

### 1. Unit Conversions

OOXML uses twips (twentieths of a point) for measurements. Key conversions:

```python
# Conversion constants
TWIPS_PER_INCH = 1440
TWIPS_PER_POINT = 20
POINTS_PER_INCH = 72

def inches_to_twips(inches: float) -> int:
    """Convert inches to twips."""
    return int(round(inches * TWIPS_PER_INCH))

def twips_to_inches(twips: int) -> float:
    """Convert twips to inches."""
    return twips / TWIPS_PER_INCH

def points_to_twips(points: float) -> int:
    """Convert points to twips."""
    return int(round(points * TWIPS_PER_POINT))
```

### 2. Page Size Presets

```python
PAGE_SIZE_PRESETS = {
    "letter": (8.5, 11.0),      # US Letter
    "legal": (8.5, 14.0),       # US Legal
    "a4": (8.27, 11.69),        # ISO A4 (210mm x 297mm)
    "a3": (11.69, 16.54),       # ISO A3
    "a5": (5.83, 8.27),         # ISO A5
    "executive": (7.25, 10.5),  # Executive
    "tabloid": (11.0, 17.0),    # Tabloid / Ledger
}
```

### 3. Section Property Change Tracking

When `track=True`:

```python
def _create_sect_pr_change(
    self,
    current_sect_pr: etree._Element,
    author: str | None = None,
) -> etree._Element:
    """Create sectPrChange element for tracked changes.

    Args:
        current_sect_pr: The current section properties element
        author: Author name for the change

    Returns:
        The sectPrChange element containing the original values
    """
    from copy import deepcopy
    from datetime import datetime, timezone

    author_name = author or self._document.author
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    change_id = self._document._xml_generator.next_change_id
    self._document._xml_generator.next_change_id += 1

    # Create sectPrChange element
    sect_pr_change = etree.Element(f"{{{WORD_NAMESPACE}}}sectPrChange")
    sect_pr_change.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
    sect_pr_change.set(f"{{{WORD_NAMESPACE}}}author", author_name)
    sect_pr_change.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

    # Copy current (original) section properties
    original_sect_pr = deepcopy(current_sect_pr)
    # Remove any existing sectPrChange from the copy
    for existing_change in original_sect_pr.findall(f"{{{WORD_NAMESPACE}}}sectPrChange"):
        original_sect_pr.remove(existing_change)

    sect_pr_change.append(original_sect_pr)
    return sect_pr_change
```

### 4. Finding Section Properties

```python
def _get_section_properties(
    self,
    section_index: int | None = None
) -> tuple[etree._Element, int]:
    """Get section properties element and its actual index.

    Args:
        section_index: Requested section index (None = last section)

    Returns:
        Tuple of (sectPr element, actual section index)

    Raises:
        IndexError: If section_index is out of range
    """
    # Find all section properties
    # Body sectPr is for the last/only section
    body = self._document.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
    body_sect_pr = body.find(f"{{{WORD_NAMESPACE}}}sectPr")

    # Paragraph sectPr elements define section breaks
    para_sect_prs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}sectPr"))
    # Filter to only those inside pPr (not the body one)
    para_sect_prs = [
        sp for sp in para_sect_prs
        if sp.getparent().tag == f"{{{WORD_NAMESPACE}}}pPr"
    ]

    # Total sections = para_sect_prs + 1 (body section)
    total_sections = len(para_sect_prs) + 1

    if section_index is None:
        # Return last section (body sectPr)
        return body_sect_pr, total_sections - 1

    if section_index < 0 or section_index >= total_sections:
        raise IndexError(
            f"Section index {section_index} out of range "
            f"(document has {total_sections} sections)"
        )

    if section_index == total_sections - 1:
        return body_sect_pr, section_index
    else:
        return para_sect_prs[section_index], section_index
```

### 5. Border Style Constants

```python
# Common border styles (ST_Border enumeration)
BORDER_STYLES = {
    "none", "nil",
    "single", "thick", "double", "triple",
    "dotted", "dashed", "dotDash", "dotDotDash", "dashSmallGap",
    "wave", "doubleWave",
    "inset", "outset", "3d",
    "thinThickSmallGap", "thickThinSmallGap",
    "thinThickMediumGap", "thickThinMediumGap",
    "thinThickLargeGap", "thickThinLargeGap",
    "thinThickThinSmallGap", "thinThickThinMediumGap", "thinThickThinLargeGap",
    # Art borders (examples)
    "apples", "stars", "hearts", "flowers", "holly",
    # ... many more art borders defined in OOXML spec
}
```

### 6. Line Number Restart Values

```python
# ST_LineNumberRestart enumeration
LINE_NUMBER_RESTART = {
    "page": "newPage",       # Restart on each page
    "section": "newSection", # Restart on each section
    "continuous": "continuous",  # Continue throughout
}
```

---

## File Structure

```
src/python_docx_redline/
    models/
        layout.py          # ColumnInfo, BorderInfo, LineNumberingInfo, etc.
    operations/
        layout.py          # LayoutOperations class
    __init__.py            # Export layout models and operations

docs/
    LAYOUT_API.md          # This design document

skills/docx/
    layout.md              # Skill guide for layout operations (Phase 8)
```

---

## Testing Strategy

### Unit Tests

```python
def test_get_columns_single_column():
    """Test reading single-column layout (default)."""
    doc = Document("tests/fixtures/single_column.docx")
    info = doc.get_columns()
    assert info.count == 1


def test_get_columns_multi_column():
    """Test reading multi-column layout."""
    doc = Document("tests/fixtures/two_columns.docx")
    info = doc.get_columns()
    assert info.count == 2
    assert info.equal_width is True


def test_set_columns():
    """Test setting column layout."""
    doc = Document("tests/fixtures/simple.docx")
    doc.set_columns(count=2, spacing=0.5)

    info = doc.get_columns()
    assert info.count == 2
    assert info.spacing == 0.5


def test_set_columns_tracked():
    """Test tracked column layout change."""
    doc = Document("tests/fixtures/simple.docx")
    doc.set_columns(count=2, track=True)

    # Verify sectPrChange exists
    sect_pr = doc._layout_ops._get_section_properties()[0]
    sect_pr_change = sect_pr.find(f"{{{WORD_NAMESPACE}}}sectPrChange")
    assert sect_pr_change is not None


def test_set_margins():
    """Test setting page margins."""
    doc = Document("tests/fixtures/simple.docx")
    doc.set_margins(top=1.5, bottom=1.5, left=1.25, right=1.25)

    section = doc.sections[0]
    assert abs(section.margins.top - 1.5) < 0.01
    assert abs(section.margins.left - 1.25) < 0.01


def test_set_page_border():
    """Test setting page border."""
    doc = Document("tests/fixtures/simple.docx")
    doc.set_page_border(style="double", color="FF0000")

    info = doc.get_page_border()
    assert info is not None
    assert info.top.style == "double"
    assert info.top.color == "FF0000"


def test_set_line_numbering():
    """Test enabling line numbering."""
    doc = Document("tests/fixtures/simple.docx")
    doc.set_line_numbering(count_by=5, restart="page")

    info = doc.get_line_numbering()
    assert info is not None
    assert info.count_by == 5
    assert info.restart == "page"


def test_add_section_break():
    """Test adding a section break."""
    doc = Document("tests/fixtures/simple.docx")
    initial_sections = len(doc.sections)

    section_idx = doc.add_section_break(
        break_type="continuous",
        after="some text in the document"
    )

    assert len(doc.sections) == initial_sections + 1
    assert section_idx == initial_sections
```

### Integration Tests

```python
def test_multi_column_layout_roundtrip():
    """Test creating and saving multi-column document."""
    doc = Document("tests/fixtures/simple.docx")

    # Add section break and change to 2 columns
    doc.add_section_break(after="Introduction", break_type="continuous")
    doc.set_columns(count=2, section_index=-1)

    # Add another break and return to single column
    doc.add_section_break(after="End of two-column content", break_type="continuous")
    doc.set_columns(count=1, section_index=-1)

    # Save and reload
    doc.save("tests/output/multi_column.docx")

    # Verify
    doc2 = Document("tests/output/multi_column.docx")
    assert len(doc2.sections) == 3
    assert doc2.get_columns(section_index=1).count == 2


def test_legal_document_layout():
    """Test typical legal document layout with line numbers."""
    doc = Document("tests/fixtures/simple.docx")

    # Set margins (1.5" left for line numbers)
    doc.set_margins(left=1.5, right=1.0)

    # Enable line numbering
    doc.set_line_numbering(count_by=1, restart="page")

    # Add page border
    doc.set_page_border(style="single", color="000000")

    # Save and verify opens in Word without errors
    doc.save("tests/output/legal_format.docx")
```

---

## Questions for Dev Team

1. **Section index convention**: Should we use 0-based (Python standard) or 1-based (Word UI) indexing for sections?

2. **Default section behavior**: When `section_index=None`, should we always target the last section, or the "current" section if there's concept of cursor position?

3. **Margin validation**: Should we validate that margins don't exceed page dimensions?

4. **Art border support**: Full art border support requires embedding images. Should this be Phase 1 or a future enhancement?

5. **python-docx compatibility**: Should we expose layout operations through the python-docx Section object interface for familiarity?

---

## Success Criteria

Implementation is complete when:

1. All documented methods work as specified
2. Tracked changes appear correctly in Word's Track Changes view
3. Documents created/modified by the API open without corruption in Word
4. 80%+ test coverage on layout operations
5. Integration tests pass with real documents
6. Documentation includes working examples
7. Skill guide enables AI agents to use layout features effectively

---

**End of Specification**
