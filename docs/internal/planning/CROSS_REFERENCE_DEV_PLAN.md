# Cross-Reference Development Plan

This document outlines the implementation plan for CRUD operations on cross-references in python-docx-redline.

## Overview

Cross-references are dynamic links within a Word document that point to other locations (headings, figures, bookmarks, etc.). They are implemented as field codes (REF, PAGEREF, NOTEREF) that Word calculates when the document is opened.

### Key Design Decisions

1. **Field code approach**: Cross-references use complex field codes (begin/instrText/separate/result/end pattern), similar to TOC.
2. **Bookmark-based targeting**: All cross-references work through bookmarks. For headings and captions, we create hidden `_Ref` bookmarks.
3. **Honest about limitations**: Display values require Word's calculation engine. We insert dirty fields with placeholder text.
4. **Leverage existing infrastructure**: Reuse BookmarkRegistry, text search, relationship management, and field code patterns from TOC.

---

## File Structure

### New Files to Create

| File | Purpose |
|------|---------|
| `src/python_docx_redline/operations/cross_references.py` | Main CrossReferenceOperations class |
| `tests/test_cross_references.py` | Unit and integration tests |

### Existing Files to Modify

| File | Changes |
|------|---------|
| `src/python_docx_redline/operations/__init__.py` | Export CrossReferenceOperations, CrossReference, CrossReferenceTarget |
| `src/python_docx_redline/document.py` | Add cross_references property and delegate methods |
| `src/python_docx_redline/errors.py` | Add CrossReferenceError, CrossReferenceTargetNotFoundError, InvalidBookmarkNameError, BookmarkAlreadyExistsError |
| `src/python_docx_redline/accessibility/bookmarks.py` | Enhance _Ref bookmark handling, add create_bookmark_at_element |
| `src/python_docx_redline/accessibility/types.py` | Add CrossReference and CrossReferenceTarget dataclasses (if not using local definitions) |

---

## Implementation Phases

### Phase 1: Core Infrastructure and Data Models

**Goal**: Establish foundation for cross-reference operations with data models and basic field code generation.

**Scope**:
- Create `CrossReferenceOperations` class skeleton
- Define `CrossReference` and `CrossReferenceTarget` dataclasses
- Add cross-reference exception classes to `errors.py`
- Implement `_create_field_code()` helper for REF/PAGEREF/NOTEREF fields
- Implement switch mapping from display options to field switches

**Methods to Implement**:
```python
class CrossReferenceOperations:
    def __init__(self, document: Document) -> None
    def _create_field_code(
        self,
        field_type: str,           # "REF", "PAGEREF", "NOTEREF"
        bookmark_name: str,
        switches: list[str],
        placeholder_text: str = ""
    ) -> list[etree._Element]
    def _get_switches_for_display(self, display: str, hyperlink: bool) -> tuple[str, list[str]]
```

**Test Scenarios**:
- Unit test field code XML structure matches expected OOXML
- Unit test switch mapping for all display options
- Unit test dirty flag is set on field begin
- Verify field code contains correct instruction text

**Dependencies**: None (foundation phase)

**Complexity**: Simple

---

### Phase 2: Bookmark Management (Create/Read)

**Goal**: Enable creating bookmarks at text locations and listing existing bookmarks.

**Scope**:
- Implement `create_bookmark()` for creating named bookmarks at text
- Implement `list_bookmarks()` for retrieving all bookmarks
- Implement `get_bookmark()` for retrieving a specific bookmark
- Add bookmark name validation (alphanumeric, underscore, max 40 chars)
- Handle hidden `_Ref` bookmarks vs user-visible bookmarks

**Methods to Implement**:
```python
def create_bookmark(
    self,
    name: str,
    at: str,
    scope: str | dict | Any | None = None
) -> str

def list_bookmarks(
    self,
    include_hidden: bool = False
) -> list[BookmarkInfo]

def get_bookmark(self, name: str) -> BookmarkInfo | None

def _validate_bookmark_name(self, name: str) -> None
def _generate_ref_bookmark_name(self) -> str
```

**Test Scenarios**:
- Create bookmark at single text location
- Reject invalid bookmark names (spaces, too long, special chars)
- Reject duplicate bookmark names
- List bookmarks filters hidden _Ref bookmarks by default
- Get specific bookmark by name
- Generate unique _Ref bookmark names

**Dependencies**: Phase 1

**Complexity**: Medium

---

### Phase 3: Basic Cross-Reference Insertion (Bookmark Targets)

**Goal**: Insert cross-references to existing bookmarks with basic display options.

**Scope**:
- Implement `insert_cross_reference()` for bookmark targets
- Support display options: "text", "page", "above_below"
- Support hyperlink option (adds `\h` switch)
- Use text search for insertion position (after/before)
- Mark fields dirty for Word calculation

**Methods to Implement**:
```python
def insert_cross_reference(
    self,
    target: str,
    display: str = "text",
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Any | None = None,
    hyperlink: bool = True,
    track: bool = False,
    author: str | None = None
) -> str

def _resolve_target(self, target: str) -> tuple[str, bool]  # (bookmark_name, was_created)
def _insert_field_at_position(self, field_elements: list, match: TextSpan, after: bool) -> None
```

**Test Scenarios**:
- Insert REF field pointing to existing bookmark
- Insert PAGEREF field (display="page")
- Verify hyperlink switch presence/absence
- Error when bookmark target doesn't exist
- Error when anchor text not found
- Round-trip: insert, save, reload, verify field structure

**Dependencies**: Phase 1, Phase 2

**Complexity**: Medium

---

### Phase 4: Heading References

**Goal**: Support cross-references to headings by text, with auto-generated bookmarks.

**Scope**:
- Extend target resolution for "heading:..." format
- Find headings by text (partial match)
- Create hidden `_Ref` bookmarks at heading paragraphs
- Support numbered heading display options: "number", "full_number", "relative_number"
- Implement `create_heading_bookmark()` convenience method

**Methods to Implement**:
```python
def create_heading_bookmark(
    self,
    heading_text: str,
    bookmark_name: str | None = None
) -> str

def _resolve_heading_target(self, heading_text: str) -> tuple[str, str | None]
def _find_heading_paragraph(self, heading_text: str) -> etree._Element | None
def _create_bookmark_at_paragraph(self, name: str, paragraph: etree._Element) -> None
def _find_existing_ref_bookmark(self, paragraph: etree._Element) -> str | None
```

**Test Scenarios**:
- Reference heading by exact text match
- Reference heading by partial text match
- Auto-create _Ref bookmark if none exists
- Reuse existing _Ref bookmark if present
- Display heading number (with \n, \r, \w switches)
- Error when heading not found

**Dependencies**: Phase 3

**Complexity**: Medium

---

### Phase 5: Caption References (Figures and Tables)

**Goal**: Support cross-references to captioned figures and tables.

**Scope**:
- Extend target resolution for "figure:N", "figure:caption text", "table:N", "table:caption text"
- Detect SEQ fields in captions
- Create bookmarks around caption content
- Support display options: "label_number", "number_only", "label_only", "caption_text"

**Methods to Implement**:
```python
def _resolve_caption_target(
    self,
    seq_id: str,           # "Figure" or "Table"
    identifier: str        # Number or caption text
) -> tuple[str, str | None]

def _find_caption_paragraph(
    self,
    seq_id: str,
    identifier: str
) -> etree._Element | None

def _parse_caption_number(self, paragraph: etree._Element, seq_id: str) -> str | None
def _get_caption_text(self, paragraph: etree._Element, seq_id: str) -> str
```

**Test Scenarios**:
- Reference figure by number (figure:1)
- Reference figure by caption text (figure:Architecture Diagram)
- Reference table by number and text
- Display "Figure 1" vs just "1" vs just "Figure"
- Error when figure/table not found
- Handle captions without SEQ fields gracefully

**Dependencies**: Phase 4

**Complexity**: Medium

---

### Phase 6: Note References and Convenience Methods

**Goal**: Support footnote/endnote references and add convenience APIs.

**Scope**:
- Extend target resolution for "footnote:N" and "endnote:N"
- Implement NOTEREF field generation with appropriate switches
- Implement `insert_page_reference()` convenience method
- Implement `insert_note_reference()` convenience method
- Support `\f` switch for note formatting style

**Methods to Implement**:
```python
def insert_page_reference(
    self,
    target: str,
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Any | None = None,
    show_position: bool = False,
    hyperlink: bool = True,
    track: bool = False,
    author: str | None = None
) -> str

def insert_note_reference(
    self,
    note_type: str,
    note_id: int | str,
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Any | None = None,
    show_position: bool = False,
    use_note_style: bool = True,
    hyperlink: bool = True,
    track: bool = False,
    author: str | None = None
) -> str

def _resolve_note_target(
    self,
    note_type: str,
    note_id: str
) -> tuple[str, str | None]

def _find_note_bookmark(self, note_type: str, note_id: str) -> str | None
def _create_note_bookmark(self, note_type: str, note_id: str) -> str
```

**Test Scenarios**:
- Insert NOTEREF to footnote
- Insert NOTEREF to endnote
- Apply note formatting style (\f switch)
- insert_page_reference convenience method works
- insert_note_reference convenience method works
- Error when note ID doesn't exist

**Dependencies**: Phase 5, existing NoteOperations

**Complexity**: Medium

---

### Phase 7: Inspection and Field Management

**Goal**: Enable reading existing cross-references and managing field updates.

**Scope**:
- Implement `get_cross_references()` to list all cross-references in document
- Implement `get_cross_reference_targets()` to list available targets
- Implement `mark_cross_references_dirty()` for bulk field update
- Parse existing field codes to extract target, switches, display value

**Methods to Implement**:
```python
def get_cross_references(self) -> list[CrossReference]

def get_cross_reference_targets(self) -> list[CrossReferenceTarget]

def mark_cross_references_dirty(self) -> int

def _parse_field_instruction(self, instruction: str) -> dict
def _extract_fields_from_body(self, field_types: list[str]) -> list[tuple[str, etree._Element]]
def _build_cross_reference_from_field(self, field_type: str, field_elem: etree._Element) -> CrossReference
```

**Test Scenarios**:
- List all REF, PAGEREF, NOTEREF fields in document
- Extract target bookmark from field instruction
- Parse switches from field instruction
- Get cached display value from field result
- List all potential targets (bookmarks, headings, captions, notes)
- Mark all cross-reference fields dirty
- Verify dirty flag is set on marked fields

**Dependencies**: Phase 6

**Complexity**: Medium

---

## Complexity Summary

| Phase | Description | Complexity | Estimated Effort |
|-------|-------------|------------|------------------|
| 1 | Core Infrastructure | Simple | 1 session |
| 2 | Bookmark Management | Medium | 1 session |
| 3 | Basic Cross-Reference Insertion | Medium | 1-2 sessions |
| 4 | Heading References | Medium | 1 session |
| 5 | Caption References | Medium | 1-2 sessions |
| 6 | Note References & Convenience | Medium | 1 session |
| 7 | Inspection & Field Management | Medium | 1 session |

**Total Estimated Effort**: 7-9 sessions

---

## Patterns to Follow

### From TOCOperations

1. **Field code structure**: Use the same begin/instrText/separate/result/end pattern
2. **Dirty flag handling**: Set `w:dirty="true"` on fldChar begin element
3. **Settings management**: Optionally set `updateFields` in settings.xml
4. **Placeholder text**: Provide meaningful placeholder (e.g., "[Update field]")

### From HyperlinkOperations

1. **Text search for insertion**: Use `_document._text_search.find_text()` with scope filtering
2. **After/before positioning**: Use `_insert_after_match()` and `_insert_before_match()` patterns
3. **Style management**: Ensure required styles exist via `_ensure_*_style()` pattern
4. **Validation**: Validate mutually exclusive parameters (after vs before)

### From BookmarkRegistry

1. **Bookmark extraction**: Use existing `_extract_bookmarks()` logic
2. **Hidden bookmark detection**: Check for `_Ref` prefix to identify hidden bookmarks
3. **Reference tracking**: Maintain bidirectional references between bookmarks and links

---

## Data Models

### CrossReference (in cross_references.py)

```python
@dataclass
class CrossReference:
    """Information about a cross-reference in the document."""
    ref: str                    # Unique reference ID (e.g., "xref:5")
    field_type: str             # "REF", "PAGEREF", or "NOTEREF"
    target_bookmark: str        # The bookmark being referenced
    switches: str               # Raw field switches (e.g., "\\h \\r")
    display_value: str | None   # Current cached display value
    is_dirty: bool              # Whether field is marked for update
    is_hyperlink: bool          # Has \h switch
    position: str               # Location in document (e.g., "p:15")

    # Parsed switch information
    show_position: bool         # Has \p switch
    number_format: str | None   # "full" (\w), "relative" (\r), "no_context" (\n)
    suppress_non_numeric: bool  # Has \d switch
```

### CrossReferenceTarget (in cross_references.py)

```python
@dataclass
class CrossReferenceTarget:
    """A potential target for a cross-reference."""
    type: str                   # "bookmark", "heading", "figure", "table", "footnote", "endnote"
    bookmark_name: str          # The bookmark name (may be auto-generated)
    display_name: str           # Human-readable name
    text_preview: str           # First ~100 chars of target content
    position: str               # Location in document
    is_hidden: bool             # Is this a hidden _Ref bookmark?

    # For numbered items
    number: str | None          # "1", "2.1", "Figure 3", etc.
    level: int | None           # Heading level (1-9)
    sequence_id: str | None     # SEQ field identifier ("Figure", "Table")
```

---

## Error Classes (add to errors.py)

```python
class CrossReferenceError(DocxRedlineError):
    """Base exception for cross-reference operations."""
    pass

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

---

## Testing Strategy

### Unit Tests (per phase)

Each phase should include unit tests for:
- Happy path functionality
- Error cases and validation
- Edge cases specific to that phase

### Integration Tests

After Phase 7, add integration tests:
1. **Round-trip test**: Insert cross-reference, save, reload, verify structure
2. **Word compatibility**: Open in Word, update fields, verify values (manual verification)
3. **Multiple references**: Create several references to same target
4. **Mixed types**: REF, PAGEREF, NOTEREF in same document

### Test Fixtures

Create test documents with:
- Multiple headings at different levels
- Captioned figures and tables with SEQ fields
- Existing bookmarks (both visible and hidden)
- Footnotes and endnotes
- Existing cross-references for inspection tests

---

## Future Enhancements (Out of Scope)

The following are documented for future work but not included in this plan:

1. **Cross-references in headers/footers**: Requires relationship management for header/footer parts
2. **Cross-references in footnotes/endnotes**: Similar to headers/footers
3. **Tracked change support**: Wrap insertions in tracked change markup
4. **Bookmark deletion**: Remove bookmarks and handle orphaned references
5. **Bookmark renaming**: Rename bookmark and update all references
6. **Broken reference detection**: Find references pointing to missing bookmarks
7. **Cross-document references**: References to content in other files

---

## Success Criteria

Each phase is complete when:

1. All specified methods are implemented
2. Unit tests pass with >80% coverage for new code
3. Integration tests pass (where applicable)
4. Code follows project patterns (type hints, docstrings, error handling)
5. No regressions in existing test suite
