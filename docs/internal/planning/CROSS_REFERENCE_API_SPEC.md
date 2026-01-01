# Cross-Reference API Specification

## Executive Summary

This document specifies the API for cross-reference support in python_docx_redline. Cross-references are dynamic links that point to other locations within a document, such as headings, figures, tables, bookmarks, and footnotes. Unlike static text, cross-references automatically update when the target content changes.

### Key Capabilities

1. **Insert cross-references** to bookmarks, headings, figures, tables, and footnotes/endnotes
2. **Display options**: text content, page number, paragraph number, or "above/below"
3. **Hyperlink support**: optionally make references clickable
4. **Inspection**: find and analyze existing cross-references
5. **Management**: mark dirty for update, list targets

### Design Principles

1. **Honest about limitations** - Cross-reference display values (page numbers, paragraph numbers) require Word's field calculation engine. We insert the field code and mark it dirty; Word populates the value when opened.
2. **Pythonic API** - `reference_type="heading"` instead of raw field switches like `\r \h`
3. **Progressive disclosure** - Simple defaults for common cases, full control available
4. **Leverage existing infrastructure** - BookmarkRegistry, HyperlinkOperations, StyleManager, field code utilities from TOC

---

## OOXML Background

### Field Code Architecture

Cross-references in Word are implemented as field codes. The three primary field types are:

| Field | Purpose | Example |
|-------|---------|---------|
| `REF` | Reference to bookmark content, paragraph numbers | `{ REF _Ref123456 \h }` |
| `PAGEREF` | Page number of bookmarked location | `{ PAGEREF _Ref123456 \p }` |
| `NOTEREF` | Footnote/endnote mark number | `{ NOTEREF FootnoteBookmark \h }` |

### Field Structure in XML

Cross-references use the complex field code pattern (same as TOC):

```xml
<w:p>
  <!-- Field begin -->
  <w:r>
    <w:fldChar w:fldCharType="begin" w:dirty="true"/>
  </w:r>
  <!-- Field instruction -->
  <w:r>
    <w:instrText xml:space="preserve"> REF _Ref123456789 \h </w:instrText>
  </w:r>
  <!-- Field separator -->
  <w:r>
    <w:fldChar w:fldCharType="separate"/>
  </w:r>
  <!-- Field result (placeholder until Word calculates) -->
  <w:r>
    <w:t>Section 2.1</w:t>
  </w:r>
  <!-- Field end -->
  <w:r>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
</w:p>
```

### Bookmark-Based Targeting

All cross-references work through bookmarks. When you insert a cross-reference to a heading in Word, it:

1. Creates a hidden bookmark at the heading (e.g., `_Ref116788778`)
2. Inserts a REF field pointing to that bookmark
3. Calculates and displays the result

Hidden bookmarks follow the naming pattern `_Ref` followed by 9+ digits. They are invisible in Word's UI unless explicitly shown.

### REF Field Switches

| Switch | Description | Example Result |
|--------|-------------|----------------|
| `\h` | Create hyperlink to target | Clickable reference |
| `\p` | Insert "above" or "below" based on relative position | "see Figure 1 below" |
| `\n` | Paragraph number without trailing periods | "1" (for "1.2.3") |
| `\r` | Relative paragraph number | "2.3" (relative to current) |
| `\w` | Full paragraph number | "1.2.3" |
| `\d` | Suppress non-numeric text (with \n, \r, \w) | "1.2.3" without "Section" |
| `\f` | Increment footnote/endnote number | Next note number |

### PAGEREF Field Switches

| Switch | Description | Example Result |
|--------|-------------|----------------|
| `\h` | Create hyperlink to target | Clickable page number |
| `\p` | Insert "above" or "below" | "on page 5 above" |

### NOTEREF Field Switches

| Switch | Description | Example Result |
|--------|-------------|----------------|
| `\h` | Create hyperlink to note | Clickable note marker |
| `\p` | Insert "above" or "below" | "see note 3 below" |
| `\f` | Use note's formatting style | Superscript mark |

### SEQ Field for Captions

Captions (Figure 1, Table 2, etc.) use the SEQ field for automatic numbering:

```xml
<w:p>
  <!-- Caption paragraph -->
  <w:r><w:t>Figure </w:t></w:r>
  <w:r>
    <w:fldChar w:fldCharType="begin"/>
  </w:r>
  <w:r>
    <w:instrText> SEQ Figure \* ARABIC </w:instrText>
  </w:r>
  <w:r>
    <w:fldChar w:fldCharType="separate"/>
  </w:r>
  <w:r>
    <w:t>1</w:t>
  </w:r>
  <w:r>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
  <w:r><w:t>: System Architecture</w:t></w:r>
</w:p>
```

Cross-references to captions point to a bookmark surrounding the SEQ field, displaying "Figure 1" or just "1" depending on the reference type.

---

## Proposed API

### Core Methods

#### `insert_cross_reference()`

```python
def insert_cross_reference(
    self,
    # Target specification
    target: str,                           # Bookmark name, "heading:Section 2", "figure:1", etc.

    # Display options
    display: str = "text",                 # What to show: "text", "page", "number", "above_below", "label_number"

    # Position specification
    after: str | None = None,              # Text to insert after
    before: str | None = None,             # Text to insert before
    scope: str | dict | Any | None = None, # Limit search scope

    # Formatting options
    hyperlink: bool = True,                # Make the reference clickable
    include_separator: str | None = None,  # Separator text if multiple values (e.g., ", ")

    # Tracked changes
    track: bool = False,
    author: str | None = None,
) -> str:
    """Insert a cross-reference to a target location in the document.

    The cross-reference is inserted as a field code that Word will calculate
    when the document is opened. The field is marked dirty to ensure Word
    updates it.

    Args:
        target: What to reference. Supports several formats:
            - Bookmark name: "MyBookmark" (direct bookmark reference)
            - Heading: "heading:Section 2.1" (finds heading by text)
            - Figure: "figure:1" or "figure:Architecture Diagram" (by number or caption)
            - Table: "table:2" or "table:Sales Data" (by number or caption)
            - Footnote: "footnote:3" (by note ID)
            - Endnote: "endnote:1" (by note ID)

        display: What text to display for the reference:
            - "text": The bookmarked text content
            - "page": Page number where target appears
            - "number": Paragraph/heading number (e.g., "2.1")
            - "full_number": Full paragraph number (e.g., "1.2.3")
            - "relative_number": Relative paragraph number
            - "above_below": "above" or "below" based on position
            - "label_number": For captions, "Figure 1" or "Table 2"
            - "label_only": Just "Figure" or "Table"
            - "number_only": Just "1" for Figure 1
            - "caption_text": Caption text without label/number

        after: Insert after this text (mutually exclusive with before)
        before: Insert before this text (mutually exclusive with after)
        scope: Limit text search to specific scope

        hyperlink: If True, the reference will be a clickable link to the target.
            Corresponds to the \h switch in REF/PAGEREF/NOTEREF fields.

        include_separator: When combining display values (future), separator text.

        track: If True, wrap the insertion in tracked change markup.
        author: Author for tracked changes.

    Returns:
        The bookmark name used for the cross-reference (may be auto-generated
        for headings, figures, etc.)

    Raises:
        ValueError: If both after and before specified, or neither specified
        ValueError: If target format is invalid
        TextNotFoundError: If anchor text not found
        AmbiguousTextError: If anchor text found multiple times
        CrossReferenceTargetNotFoundError: If the specified target doesn't exist

    Example:
        >>> # Reference a bookmark
        >>> doc.insert_cross_reference(
        ...     target="DefinitionsSection",
        ...     display="text",
        ...     after="as defined in "
        ... )
        'DefinitionsSection'

        >>> # Reference to a heading showing page number
        >>> doc.insert_cross_reference(
        ...     target="heading:Introduction",
        ...     display="page",
        ...     after="(see page ",
        ... )
        '_Ref123456789'

        >>> # Reference to Figure 1 showing "Figure 1"
        >>> doc.insert_cross_reference(
        ...     target="figure:1",
        ...     display="label_number",
        ...     after="illustrated in "
        ... )
        '_Ref987654321'

        >>> # Reference to footnote with "above/below"
        >>> doc.insert_cross_reference(
        ...     target="footnote:3",
        ...     display="above_below",
        ...     after="see footnote 3 "
        ... )
        '_Ref112233445'
    """
```

#### `insert_page_reference()`

Convenience method for the common "see page X" pattern:

```python
def insert_page_reference(
    self,
    target: str,
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Any | None = None,
    show_position: bool = False,           # Add "above" or "below"
    hyperlink: bool = True,
    track: bool = False,
    author: str | None = None,
) -> str:
    """Insert a page number cross-reference.

    This is a convenience wrapper around insert_cross_reference with
    display="page". Use this when you want to show the page number
    where a target appears.

    Args:
        target: What to reference (bookmark name, "heading:...", etc.)
        after: Insert after this text
        before: Insert before this text
        scope: Limit text search scope
        show_position: If True, append "above" or "below" based on position
        hyperlink: Make the page number clickable
        track: Track the insertion as a change
        author: Author for tracked change

    Returns:
        The bookmark name used for the reference

    Example:
        >>> # "see page 5"
        >>> doc.insert_page_reference(
        ...     target="heading:Methodology",
        ...     after="see page "
        ... )

        >>> # "on page 12 above"
        >>> doc.insert_page_reference(
        ...     target="TableOfResults",
        ...     after="on page ",
        ...     show_position=True
        ... )
    """
```

#### `insert_note_reference()`

Convenience method for footnote/endnote references:

```python
def insert_note_reference(
    self,
    note_type: str,                        # "footnote" or "endnote"
    note_id: int | str,
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Any | None = None,
    show_position: bool = False,           # Add "above" or "below"
    use_note_style: bool = True,           # Format like note marker (superscript)
    hyperlink: bool = True,
    track: bool = False,
    author: str | None = None,
) -> str:
    """Insert a cross-reference to a footnote or endnote.

    Creates a NOTEREF field that displays the note's marker number.
    Useful for referring to a note from elsewhere in the document.

    Args:
        note_type: Either "footnote" or "endnote"
        note_id: The note's ID number
        after: Insert after this text
        before: Insert before this text
        scope: Limit text search scope
        show_position: If True, append "above" or "below"
        use_note_style: If True, format the marker like the note reference style
        hyperlink: Make the reference clickable
        track: Track the insertion
        author: Author for tracked change

    Returns:
        The bookmark name for the note reference

    Example:
        >>> # "see note 3"
        >>> doc.insert_note_reference(
        ...     note_type="footnote",
        ...     note_id=3,
        ...     after="see note "
        ... )
    """
```

### Inspection Methods

#### `get_cross_references()`

```python
def get_cross_references(self) -> list[CrossReference]:
    """Get all cross-references in the document.

    Scans the document for REF, PAGEREF, and NOTEREF fields and returns
    information about each cross-reference found.

    Returns:
        List of CrossReference objects with field details

    Example:
        >>> for xref in doc.get_cross_references():
        ...     print(f"{xref.field_type}: {xref.target_bookmark} -> {xref.display_value}")
    """
```

#### `get_cross_reference_targets()`

```python
def get_cross_reference_targets(self) -> list[CrossReferenceTarget]:
    """Get all potential cross-reference targets in the document.

    Identifies:
    - All bookmarks (both user-created and auto-generated _Ref bookmarks)
    - All numbered headings
    - All captioned figures, tables, equations
    - All footnotes and endnotes

    Returns:
        List of CrossReferenceTarget objects describing available targets

    Example:
        >>> targets = doc.get_cross_reference_targets()
        >>> for target in targets:
        ...     print(f"{target.type}: {target.display_name}")
        heading: 1. Introduction
        heading: 1.1 Background
        figure: Figure 1: Architecture
        bookmark: DefinitionsSection
    """
```

### Bookmark Management

Since cross-references depend on bookmarks, we need robust bookmark support:

#### `create_bookmark()`

```python
def create_bookmark(
    self,
    name: str,
    at: str,                               # Text to bookmark
    scope: str | dict | Any | None = None,
) -> str:
    """Create a named bookmark at the specified text.

    The bookmark will span the matched text, allowing cross-references
    to display that text content.

    Args:
        name: Bookmark name (must be unique, no spaces, max 40 chars)
        at: The text to bookmark
        scope: Limit text search scope

    Returns:
        The bookmark name

    Raises:
        ValueError: If bookmark name is invalid or already exists
        TextNotFoundError: If text not found
        AmbiguousTextError: If text found multiple times

    Example:
        >>> doc.create_bookmark("ImportantClause", at="Force Majeure")
        'ImportantClause'
    """
```

#### `create_heading_bookmark()`

```python
def create_heading_bookmark(
    self,
    heading_text: str,
    bookmark_name: str | None = None,
) -> str:
    """Create a bookmark at a heading for cross-reference purposes.

    If bookmark_name is not provided, generates a hidden bookmark
    using the _Ref naming convention.

    Args:
        heading_text: Text of the heading to bookmark
        bookmark_name: Custom bookmark name (or None for auto-generated)

    Returns:
        The bookmark name (auto-generated if not provided)

    Example:
        >>> bk = doc.create_heading_bookmark("Section 2.1 Methodology")
        >>> doc.insert_cross_reference(target=bk, display="text", after="see ")
    """
```

#### `get_bookmark()`

```python
def get_bookmark(self, name: str) -> BookmarkInfo | None:
    """Get information about a specific bookmark.

    Args:
        name: The bookmark name

    Returns:
        BookmarkInfo if found, None otherwise
    """
```

#### `list_bookmarks()`

```python
def list_bookmarks(
    self,
    include_hidden: bool = False,
) -> list[BookmarkInfo]:
    """List all bookmarks in the document.

    Args:
        include_hidden: If True, include _Ref and other hidden bookmarks

    Returns:
        List of BookmarkInfo objects
    """
```

### Field Update Management

#### `mark_cross_references_dirty()`

```python
def mark_cross_references_dirty(self) -> int:
    """Mark all cross-reference fields as needing update.

    Sets the dirty flag on all REF, PAGEREF, and NOTEREF fields so
    Word will recalculate them when the document is opened.

    Returns:
        Number of fields marked dirty

    Example:
        >>> doc.edit("1.1", "1.2", at="Section 1.1")
        >>> doc.mark_cross_references_dirty()  # References to this section will update
        5
    """
```

---

## Data Models

### CrossReference

```python
@dataclass
class CrossReference:
    """Information about a cross-reference in the document.

    Represents a REF, PAGEREF, or NOTEREF field that references
    another location in the document.
    """
    ref: str                               # Unique reference ID (e.g., "xref:5")
    field_type: str                        # "REF", "PAGEREF", or "NOTEREF"
    target_bookmark: str                   # The bookmark being referenced
    switches: str                          # Raw field switches (e.g., "\h \r")
    display_value: str | None              # Current cached display value (may be stale)
    is_dirty: bool                         # Whether field is marked for update
    is_hyperlink: bool                     # Has \h switch
    position: str                          # Location in document (e.g., "p:15")

    # Parsed switch information
    show_position: bool                    # Has \p switch
    number_format: str | None              # "full" (\w), "relative" (\r), "no_context" (\n)
    suppress_non_numeric: bool             # Has \d switch
```

### CrossReferenceTarget

```python
@dataclass
class CrossReferenceTarget:
    """A potential target for a cross-reference.

    Represents something that can be referenced: a bookmark, heading,
    caption, or note.
    """
    type: str                              # "bookmark", "heading", "figure", "table", "footnote", "endnote"
    bookmark_name: str                     # The bookmark name (may be auto-generated)
    display_name: str                      # Human-readable name
    text_preview: str                      # First ~100 chars of target content
    position: str                          # Location in document
    is_hidden: bool                        # Is this a hidden _Ref bookmark?

    # For numbered items
    number: str | None                     # "1", "2.1", "Figure 3", etc.
    level: int | None                      # Heading level (1-9)
    sequence_id: str | None                # SEQ field identifier ("Figure", "Table")
```

### BookmarkInfo

The existing `BookmarkInfo` from `accessibility/types.py` should be used:

```python
@dataclass
class BookmarkInfo:
    """Information about a bookmark."""
    name: str
    ref: str
    location: str
    bookmark_id: str
    text_preview: str
    span_end_location: str | None
    referenced_by: list[str]               # List of cross-reference refs
```

---

## Implementation Notes

### Field Code Generation

Reuse the field code generation pattern from TOCOperations:

```python
def _create_cross_reference_field(
    self,
    field_type: str,                       # "REF", "PAGEREF", "NOTEREF"
    bookmark_name: str,
    switches: list[str],
    placeholder_text: str = "",
) -> etree._Element:
    """Create the XML structure for a cross-reference field.

    Uses the complex field pattern: begin -> instruction -> separate -> result -> end
    """
    para = etree.Element(w("p"))

    # Field begin with dirty flag
    run_begin = etree.SubElement(para, w("r"))
    fld_char_begin = etree.SubElement(run_begin, w("fldChar"))
    fld_char_begin.set(w("fldCharType"), "begin")
    fld_char_begin.set(w("dirty"), "true")

    # Field instruction
    run_instr = etree.SubElement(para, w("r"))
    instr_text = etree.SubElement(run_instr, w("instrText"))
    instr_text.set(f"{{{XML_NAMESPACE}}}space", "preserve")

    # Build instruction string
    switch_str = " ".join(switches) if switches else ""
    instr_text.text = f" {field_type} {bookmark_name} {switch_str} "

    # Field separator
    run_sep = etree.SubElement(para, w("r"))
    fld_char_sep = etree.SubElement(run_sep, w("fldChar"))
    fld_char_sep.set(w("fldCharType"), "separate")

    # Placeholder result
    run_result = etree.SubElement(para, w("r"))
    text_result = etree.SubElement(run_result, w("t"))
    text_result.text = placeholder_text or "[Update field]"

    # Field end
    run_end = etree.SubElement(para, w("r"))
    fld_char_end = etree.SubElement(run_end, w("fldChar"))
    fld_char_end.set(w("fldCharType"), "end")

    return para
```

### Target Resolution

The `_resolve_target()` method handles different target formats:

```python
def _resolve_target(self, target: str) -> tuple[str, str | None]:
    """Resolve a target specification to a bookmark name.

    Args:
        target: Target specification (bookmark name, "heading:...", etc.)

    Returns:
        Tuple of (bookmark_name, created_bookmark_name or None)

    Raises:
        CrossReferenceTargetNotFoundError: If target not found
    """
    # Direct bookmark reference
    if not ":" in target or target.startswith("_Ref"):
        bookmark = self._document.get_bookmark(target)
        if bookmark:
            return target, None
        raise CrossReferenceTargetNotFoundError(f"Bookmark '{target}' not found")

    prefix, value = target.split(":", 1)

    if prefix == "heading":
        return self._resolve_heading_target(value)
    elif prefix == "figure":
        return self._resolve_caption_target("Figure", value)
    elif prefix == "table":
        return self._resolve_caption_target("Table", value)
    elif prefix == "footnote":
        return self._resolve_note_target("footnote", value)
    elif prefix == "endnote":
        return self._resolve_note_target("endnote", value)
    else:
        raise ValueError(f"Unknown target type: {prefix}")
```

### Heading Resolution

```python
def _resolve_heading_target(self, heading_text: str) -> tuple[str, str]:
    """Find or create a bookmark for a heading.

    Searches for headings matching the text and either returns an
    existing _Ref bookmark or creates a new one.
    """
    # Find paragraph with heading style matching the text
    for para in self._document.paragraphs:
        if para.style and para.style.name.startswith("Heading"):
            if heading_text.lower() in para.text.lower():
                # Check for existing _Ref bookmark
                existing = self._find_ref_bookmark_at(para)
                if existing:
                    return existing, None

                # Create new hidden bookmark
                new_bookmark = self._generate_ref_bookmark_name()
                self._create_bookmark_at_paragraph(new_bookmark, para)
                return new_bookmark, new_bookmark

    raise CrossReferenceTargetNotFoundError(f"Heading '{heading_text}' not found")
```

### Switch Mapping

Map display options to field switches:

```python
DISPLAY_TO_SWITCHES = {
    "text": [],                            # REF with no special switches
    "page": [],                            # Use PAGEREF instead of REF
    "number": ["\\n"],                     # Paragraph number without context
    "full_number": ["\\w"],                # Full paragraph number
    "relative_number": ["\\r"],            # Relative paragraph number
    "above_below": ["\\p"],                # Position relative to reference
    "label_number": [],                    # For captions, display full "Figure 1"
    "number_only": ["\\#"],                # Just the number from caption
    "label_only": [],                      # Just "Figure" or "Table" (requires text parsing)
    "caption_text": [],                    # Caption text only (requires bookmark adjustment)
}

FIELD_TYPE_BY_DISPLAY = {
    "page": "PAGEREF",
    # All others use REF by default
}
```

### Integration with Existing Code

The implementation should:

1. **Use BookmarkRegistry** for bookmark extraction and validation
2. **Follow HyperlinkOperations patterns** for text search and insertion
3. **Use TOCOperations field code patterns** for XML generation
4. **Integrate with NoteOperations** for footnote/endnote handling
5. **Leverage existing RelationshipManager** if hyperlinks need relationships

---

## Examples

### Legal Document with Clause References

```python
doc = Document("contract.docx")

# Create bookmarks at key sections
doc.create_bookmark("DefSection", at="1. DEFINITIONS")
doc.create_bookmark("LimitationClause", at="4.3 Limitation of Liability")

# Insert cross-references
doc.insert_cross_reference(
    target="DefSection",
    display="text",
    after="As defined in "
)

doc.insert_cross_reference(
    target="LimitationClause",
    display="page",
    after="(see page ",
)
# Results in: "(see page 5)" when opened in Word
```

### Academic Paper with Figure References

```python
doc = Document("thesis.docx")

# Reference to a figure by number
doc.insert_cross_reference(
    target="figure:1",
    display="label_number",
    after="As shown in ",
)
# Results in: "As shown in Figure 1"

# Reference to the same figure with page
doc.insert_cross_reference(
    target="figure:1",
    display="page",
    after=" (page ",
)
# Combine: "As shown in Figure 1 (page 12)"

# Reference using figure caption text
doc.insert_cross_reference(
    target="figure:System Architecture",
    display="number_only",
    after="see Figure ",
)
# Results in: "see Figure 1"
```

### Report with Heading References

```python
doc = Document("report.docx")

# Reference by heading text
doc.insert_cross_reference(
    target="heading:Methodology",
    display="text",
    after="For details, see "
)
# Results in: "For details, see Section 2.1 Methodology"

# Show paragraph number only
doc.insert_cross_reference(
    target="heading:Results and Discussion",
    display="number",
    after="discussed in section ",
)
# Results in: "discussed in section 3"

# Page reference with above/below
doc.insert_page_reference(
    target="heading:Conclusion",
    after="concluding remarks on page ",
    show_position=True,
)
# Results in: "concluding remarks on page 15 below"
```

### Footnote References

```python
doc = Document("article.docx")

# Reference to an existing footnote
doc.insert_note_reference(
    note_type="footnote",
    note_id=3,
    after="(see note ",
)
# Results in: "(see note 3)"

# With position indicator
doc.insert_note_reference(
    note_type="footnote",
    note_id=3,
    after="see note ",
    show_position=True,
)
# Results in: "see note 3 above"
```

---

## Error Handling

### Exception Hierarchy

```python
class CrossReferenceError(Exception):
    """Base exception for cross-reference operations."""

class CrossReferenceTargetNotFoundError(CrossReferenceError):
    """The specified target for a cross-reference could not be found."""
    def __init__(self, target: str, available_targets: list[str] | None = None):
        self.target = target
        self.available_targets = available_targets or []
        super().__init__(f"Cross-reference target not found: {target}")

class InvalidBookmarkNameError(CrossReferenceError):
    """The bookmark name is invalid."""
    def __init__(self, name: str, reason: str):
        self.name = name
        self.reason = reason
        super().__init__(f"Invalid bookmark name '{name}': {reason}")

class BookmarkAlreadyExistsError(CrossReferenceError):
    """A bookmark with this name already exists."""
    def __init__(self, name: str):
        self.name = name
        super().__init__(f"Bookmark '{name}' already exists")
```

### Validation Rules

1. **Bookmark names**: Must start with a letter, contain only alphanumeric and underscores, max 40 characters
2. **Target format**: Must be valid format or existing bookmark name
3. **Note IDs**: Must reference existing footnotes/endnotes
4. **Heading text**: Must match at least one heading paragraph

---

## Test Scenarios

### Unit Tests

1. **Field code generation**
   - REF field with various switch combinations
   - PAGEREF field with/without \p switch
   - NOTEREF field with formatting switch
   - Proper dirty flag setting

2. **Target resolution**
   - Direct bookmark reference
   - Heading by exact text match
   - Heading by partial text match
   - Figure by number
   - Figure by caption text
   - Footnote/endnote by ID

3. **Bookmark creation**
   - Valid name creation
   - Invalid name rejection
   - Duplicate name rejection
   - Auto-generated _Ref names

4. **Switch mapping**
   - Display option to switch conversion
   - Field type selection based on display

### Integration Tests

1. **Round-trip**: Insert cross-reference, save, reload, verify field structure
2. **Word compatibility**: Open in Word, update fields, verify calculated values
3. **Multiple references**: Create several references to same target
4. **Mixed types**: REF, PAGEREF, NOTEREF in same document

### Edge Cases

1. **No headings**: Handle documents without heading styles
2. **No captions**: Handle documents without figure/table captions
3. **Empty bookmarks**: Handle bookmarks with no text content
4. **Nested bookmarks**: Handle overlapping bookmark ranges
5. **Large documents**: Performance with many references

---

## Implementation Phases

### Phase 1: Core Cross-Reference Insertion (MVP)

- [ ] `insert_cross_reference()` with bookmark targets
- [ ] REF field generation with \h switch
- [ ] Basic display options: "text", "page"
- [ ] Field dirty flag handling
- [ ] Basic unit tests

### Phase 2: Heading References

- [ ] Target resolution for "heading:..."
- [ ] Hidden _Ref bookmark creation
- [ ] Heading text search
- [ ] Numbered heading support (\n, \r, \w switches)

### Phase 3: Caption References

- [ ] Target resolution for "figure:..." and "table:..."
- [ ] SEQ field detection
- [ ] Caption bookmark handling
- [ ] Label/number display options

### Phase 4: Note References

- [ ] `insert_note_reference()` method
- [ ] NOTEREF field generation
- [ ] \f formatting switch
- [ ] Integration with NoteOperations

### Phase 5: Convenience Methods

- [ ] `insert_page_reference()` shorthand
- [ ] `create_bookmark()` method
- [ ] `create_heading_bookmark()` method

### Phase 6: Inspection and Management

- [ ] `get_cross_references()` - list all references
- [ ] `get_cross_reference_targets()` - list available targets
- [ ] `mark_cross_references_dirty()` - bulk update
- [ ] `list_bookmarks()` with hidden filter

### Phase 7: Advanced Features (Future)

- [ ] Cross-references in headers/footers
- [ ] Cross-references in footnotes/endnotes
- [ ] Tracked change support for insertions
- [ ] Bookmark editing (rename, delete)
- [ ] Broken reference detection

---

## Dependencies

### Existing Infrastructure to Leverage

- `BookmarkRegistry` - Bookmark extraction and lookup
- `HyperlinkOperations` - Text search, insertion patterns, hyperlink creation
- `TOCOperations` - Field code structure, dirty flag handling
- `NoteOperations` - Footnote/endnote access
- `TextSearch` - Finding text spans for insertion
- `ScopeEvaluator` - Paragraph filtering
- `RelationshipManager` - If hyperlinks need relationship entries

### New Infrastructure Needed

- `CrossReferenceOperations` class following existing patterns
- `CrossReference` and `CrossReferenceTarget` dataclasses
- Target resolution utilities
- `_Ref` bookmark generation

---

## References

- [OOXML REF Field Specification](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_REFREF_topic_ID0ESRL1.html)
- [OOXML PAGEREF Field](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_PAGEREFPAGEREF_topic_ID0EHXK1.html)
- [OOXML NOTEREF Field](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_NOTEREFNOTEREF_topic_ID0E5EK1.html)
- [OOXML SEQ Field](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_SEQSEQ_topic_ID0ETJM1.html)
- [OOXML Bookmark Elements](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_Bookmarks_topic_ID0EFMWW.html)
- [Cross-reference Fields in Word](https://wordaddins.com/support/cross-reference-fields-in-word/)
- [Microsoft Create a Cross-Reference](https://support.microsoft.com/en-us/office/create-a-cross-reference-300b208c-e45a-487a-880b-a02767d9774b)
- [python-docx Cross-Reference Feature Request](https://github.com/python-openxml/python-docx/issues/97)
