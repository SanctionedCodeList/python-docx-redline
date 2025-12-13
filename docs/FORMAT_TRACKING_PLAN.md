# Format-Only Tracked Changes — Product Spec & Implementation Plan

## Summary

Add support for tracking formatting changes (bold, italic, font size, color, etc.) as proper OOXML revision elements (`<w:rPrChange>`, `<w:pPrChange>`), so that formatting modifications appear in Word's tracked changes panel alongside text insertions and deletions.

This completes the tracked changes story by enabling the library to track **all** types of document modifications, not just text content changes.

## Problem

Current behavior:

- Text insertions are tracked via `<w:ins>`
- Text deletions are tracked via `<w:del>`
- Text moves are tracked via `<w:moveFrom>` / `<w:moveTo>`
- **Formatting changes are NOT tracked** — applying bold/italic/etc. modifies the document silently

This gap means:

1. Users cannot see formatting changes in Word's revision pane
2. `accept_all_changes()` and `reject_all_changes()` ignore formatting modifications
3. Legal and compliance workflows that require complete audit trails are incomplete

## Goals

### G1 — Track run-level formatting changes

Generate proper `<w:rPrChange>` elements when modifying character-level properties:

- Font styling: bold, italic, underline, strikethrough
- Font attributes: name, size, color, highlight
- Text effects: superscript, subscript, small caps, all caps

### G2 — Track paragraph-level formatting changes

Generate proper `<w:pPrChange>` elements when modifying paragraph properties:

- Alignment: left, center, right, justify
- Spacing: before, after, line spacing
- Indentation: left, right, first line, hanging

### G3 — Integrate with existing change management

- `accept_all_changes()` should accept formatting changes (apply current, remove `<w:rPrChange>`)
- `reject_all_changes()` should reject formatting changes (restore previous, remove `<w:rPrChange>`)
- Change ID sequencing should include formatting changes

### G4 — Preserve existing formatting

When applying new formatting to text that already has formatting, preserve unmodified properties. For example, applying bold to italic text should result in bold+italic, not just bold.

## Non-Goals (MVP)

- Tracking style changes (applying a named style like "Heading 1")
- Tracking section-level formatting (`<w:sectPrChange>`)
- Tracking table formatting (`<w:tblPrChange>`, `<w:trPrChange>`, `<w:tcPrChange>`)
- Tracking changes to headers/footers
- Format comparison in `compare_to()` (detecting formatting differences between documents)

## OOXML Background

### Run Properties Change (`<w:rPrChange>`)

Tracks changes to character-level formatting within a run:

```xml
<w:r>
  <w:rPr>
    <w:b/>                    <!-- Current state: bold ON -->
    <w:i/>                    <!-- Current state: italic ON -->
    <w:rPrChange w:id="5" w:author="Jane" w:date="2025-12-12T10:00:00Z">
      <w:rPr>
        <w:i/>                <!-- Previous state: only italic (bold was OFF) -->
      </w:rPr>
    </w:rPrChange>
  </w:rPr>
  <w:t>Important text</w:t>
</w:r>
```

Key points:

- `<w:rPrChange>` is a **child** of the current `<w:rPr>`
- It contains a nested `<w:rPr>` with the **previous** properties
- Required attributes: `w:id`, `w:author`, `w:date`
- Optional: MS365 identity attributes (`w15:userId`, `w15:providerId`)

### Paragraph Properties Change (`<w:pPrChange>`)

Tracks changes to paragraph-level formatting:

```xml
<w:pPr>
  <w:jc w:val="center"/>      <!-- Current state: centered -->
  <w:pPrChange w:id="6" w:author="Jane" w:date="2025-12-12T10:00:00Z">
    <w:pPr>
      <w:jc w:val="left"/>    <!-- Previous state: left-aligned -->
    </w:pPr>
  </w:pPrChange>
</w:pPr>
```

### Common Run Properties (w:rPr children)

| Element | Property | Python Parameter |
|---------|----------|------------------|
| `<w:b/>` | Bold | `bold=True` |
| `<w:i/>` | Italic | `italic=True` |
| `<w:u w:val="single"/>` | Underline | `underline=True` or `underline="single"` |
| `<w:strike/>` | Strikethrough | `strikethrough=True` |
| `<w:sz w:val="24"/>` | Font size (half-points) | `font_size=12` (points) |
| `<w:szCs w:val="24"/>` | Complex script size | (auto-set with font_size) |
| `<w:rFonts w:ascii="Arial"/>` | Font name | `font_name="Arial"` |
| `<w:color w:val="FF0000"/>` | Text color | `color="#FF0000"` |
| `<w:highlight w:val="yellow"/>` | Highlight color | `highlight="yellow"` |
| `<w:vertAlign w:val="superscript"/>` | Superscript | `superscript=True` |
| `<w:vertAlign w:val="subscript"/>` | Subscript | `subscript=True` |
| `<w:smallCaps/>` | Small caps | `small_caps=True` |
| `<w:caps/>` | All caps | `all_caps=True` |

### Common Paragraph Properties (w:pPr children)

| Element | Property | Python Parameter |
|---------|----------|------------------|
| `<w:jc w:val="center"/>` | Alignment | `alignment="center"` |
| `<w:spacing w:before="240"/>` | Space before (twips) | `spacing_before=12` (points) |
| `<w:spacing w:after="240"/>` | Space after (twips) | `spacing_after=12` (points) |
| `<w:spacing w:line="360"/>` | Line spacing (twips) | `line_spacing=1.5` (multiplier) |
| `<w:ind w:left="720"/>` | Left indent (twips) | `indent_left=0.5` (inches) |
| `<w:ind w:right="720"/>` | Right indent (twips) | `indent_right=0.5` (inches) |
| `<w:ind w:firstLine="720"/>` | First line indent | `indent_first_line=0.5` (inches) |
| `<w:ind w:hanging="720"/>` | Hanging indent | `indent_hanging=0.5` (inches) |

Note: 1 point = 20 twips, 1 inch = 1440 twips

## Public API Spec

### 1) `format_tracked()` — Character Formatting

Apply character-level formatting to specific text with tracking:

```python
def format_tracked(
    self,
    text: str,
    *,
    # Font styling (None = don't change, True/False = set)
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | str | None = None,  # True, False, or style name
    strikethrough: bool | None = None,

    # Font attributes (None = don't change)
    font_name: str | None = None,
    font_size: float | None = None,       # Points (e.g., 12, 14.5)
    color: str | None = None,             # Hex color "#RRGGBB" or "auto"
    highlight: str | None = None,         # Color name: "yellow", "green", etc.

    # Text effects (None = don't change)
    superscript: bool | None = None,
    subscript: bool | None = None,
    small_caps: bool | None = None,
    all_caps: bool | None = None,

    # Targeting
    scope: str | dict | Callable | None = None,
    occurrence: int | str = "first",      # 1, 2, ..., "first", "last", "all"

    # Attribution
    author: str | None = None,
) -> FormatResult:
    """Apply character formatting to text with tracked changes.

    Args:
        text: The text to format (found via text search)
        bold: Set bold on (True), off (False), or leave unchanged (None)
        italic: Set italic on/off/unchanged
        underline: Set underline on/off/unchanged, or underline style
        strikethrough: Set strikethrough on/off/unchanged
        font_name: Set font family name
        font_size: Set font size in points
        color: Set text color as hex "#RRGGBB" or "auto"
        highlight: Set highlight color name
        superscript: Set superscript on/off/unchanged
        subscript: Set subscript on/off/unchanged
        small_caps: Set small caps on/off/unchanged
        all_caps: Set all caps on/off/unchanged
        scope: Limit search to specific paragraphs/sections
        occurrence: Which occurrence(s) to format
        author: Override default author for this change

    Returns:
        FormatResult with details of the formatting applied

    Raises:
        TextNotFoundError: If text is not found
        AmbiguousTextError: If multiple matches and occurrence not specified

    Example:
        >>> doc.format_tracked("IMPORTANT", bold=True, color="#FF0000")
        >>> doc.format_tracked("Section 2.1", italic=True, scope="section:Introduction")
    """
```

### 2) `format_paragraph_tracked()` — Paragraph Formatting

Apply paragraph-level formatting with tracking:

```python
def format_paragraph_tracked(
    self,
    *,
    # Paragraph targeting (at least one required)
    containing: str | None = None,
    starting_with: str | None = None,
    ending_with: str | None = None,
    index: int | None = None,             # 0-based paragraph index

    # Alignment (None = don't change)
    alignment: str | None = None,         # "left", "center", "right", "justify"

    # Spacing in points (None = don't change)
    spacing_before: float | None = None,
    spacing_after: float | None = None,
    line_spacing: float | None = None,    # Multiplier: 1.0, 1.5, 2.0

    # Indentation in inches (None = don't change)
    indent_left: float | None = None,
    indent_right: float | None = None,
    indent_first_line: float | None = None,
    indent_hanging: float | None = None,

    # Targeting
    scope: str | dict | Callable | None = None,

    # Attribution
    author: str | None = None,
) -> FormatResult:
    """Apply paragraph formatting with tracked changes.

    Args:
        containing: Find paragraph containing this text
        starting_with: Find paragraph starting with this text
        ending_with: Find paragraph ending with this text
        index: Target paragraph by index (0-based)
        alignment: Set paragraph alignment
        spacing_before: Set space before paragraph (points)
        spacing_after: Set space after paragraph (points)
        line_spacing: Set line spacing multiplier
        indent_left: Set left indent (inches)
        indent_right: Set right indent (inches)
        indent_first_line: Set first line indent (inches)
        indent_hanging: Set hanging indent (inches)
        scope: Limit search to specific sections
        author: Override default author for this change

    Returns:
        FormatResult with details of the formatting applied

    Raises:
        TextNotFoundError: If no matching paragraph found
        AmbiguousTextError: If multiple matches without disambiguation

    Example:
        >>> doc.format_paragraph_tracked(containing="WHEREAS", alignment="center")
        >>> doc.format_paragraph_tracked(index=0, spacing_after=12)
    """
```

### 3) `FormatResult` — Return Type

```python
@dataclass
class FormatResult:
    """Result of a format operation."""

    success: bool
    text_matched: str                     # The text that was formatted
    paragraph_index: int                  # Index of affected paragraph
    changes_applied: dict[str, Any]       # {"bold": True, "color": "#FF0000"}
    previous_formatting: dict[str, Any]   # {"bold": False, "color": "auto"}
    change_id: int                        # The w:id assigned to this change
```

### 4) Batch Operations

Support format operations in `apply_edits()`:

```python
doc.apply_edits([
    {
        "type": "format_tracked",
        "text": "IMPORTANT",
        "bold": True,
        "color": "#FF0000"
    },
    {
        "type": "format_paragraph_tracked",
        "containing": "Section 1",
        "alignment": "center"
    }
])
```

And in YAML files:

```yaml
edits:
  - type: format_tracked
    text: "IMPORTANT"
    bold: true
    color: "#FF0000"

  - type: format_paragraph_tracked
    containing: "Section 1"
    alignment: center
```

### 5) Accept/Reject Integration

Extend existing methods to handle formatting changes:

```python
# Accept formatting changes (keep current formatting, remove tracking)
doc.accept_all_changes()  # Already exists, extend to handle rPrChange/pPrChange

# Reject formatting changes (restore previous formatting, remove tracking)
doc.reject_all_changes()  # Already exists, extend to handle rPrChange/pPrChange
```

## Functional Requirements (Acceptance Criteria)

### A) Character Formatting

1. **Single property change**
   - Apply `bold=True` to "Section 2.1"
   - Must produce `<w:rPrChange>` with previous state (no bold)
   - Word must show the change in tracked changes pane

2. **Multiple property changes**
   - Apply `bold=True, italic=True, color="#FF0000"` to same text
   - Must produce single `<w:rPrChange>` capturing all previous values
   - Not three separate changes

3. **Preserve existing formatting**
   - Text already has `<w:i/>` (italic)
   - Apply `bold=True`
   - Result must have `<w:b/><w:i/>` with `<w:rPrChange>` showing only `<w:i/>` as previous

4. **Toggle formatting off**
   - Text has `<w:b/>` (bold)
   - Apply `bold=False`
   - Result must remove `<w:b/>` with `<w:rPrChange>` showing `<w:b/>` as previous

5. **Partial run formatting**
   - Paragraph: "This is important text"
   - Format only "important"
   - Must split run and apply formatting only to "important" portion

6. **Cross-run text formatting**
   - Target text spans multiple runs with different existing formatting
   - Must handle run boundaries correctly
   - Each affected run gets its own `<w:rPrChange>` with its previous state

### B) Paragraph Formatting

7. **Alignment change**
   - Change paragraph from left-aligned to centered
   - Must produce `<w:pPrChange>` with `<w:jc w:val="left"/>` as previous

8. **Spacing change**
   - Set `spacing_before=12, spacing_after=12`
   - Must produce correct twip values (240) in XML

9. **Indentation change**
   - Set `indent_left=0.5` (inches)
   - Must produce correct twip value (720) in XML

### C) Integration

10. **Accept formatting changes**
    - Document has `<w:rPrChange>` elements
    - Call `accept_all_changes()`
    - Must remove `<w:rPrChange>` elements, keep current formatting

11. **Reject formatting changes**
    - Document has `<w:rPrChange>` elements
    - Call `reject_all_changes()`
    - Must restore previous formatting from `<w:rPrChange>`, remove element

12. **Change ID sequencing**
    - Apply text insertion, then formatting change, then deletion
    - All three must have unique, sequential `w:id` values

13. **MS365 identity support**
    - Document has `AuthorIdentity` configured
    - Formatting changes must include `w15:userId` and `w15:providerId`

### D) Edge Cases

14. **No-op formatting**
    - Apply `bold=True` to text that is already bold
    - Should not create a tracked change (no actual change)

15. **Empty previous state**
    - Apply formatting to text with no existing `<w:rPr>`
    - Must create `<w:rPrChange>` with empty `<w:rPr/>` as previous

16. **Nested existing changes**
    - Text is inside `<w:ins>` (tracked insertion)
    - Apply formatting
    - Must correctly nest `<w:rPrChange>` within the inserted run

## Proposed Technical Design

### 1) TrackedXMLGenerator Extensions

Add new methods to `TrackedXMLGenerator`:

```python
def create_run_property_change(
    self,
    current_rpr: etree.Element | None,
    previous_rpr: etree.Element | None,
    author: str | None = None,
) -> etree.Element:
    """Generate <w:rPrChange> element.

    Returns an element to insert as last child of <w:rPr>.
    """

def create_paragraph_property_change(
    self,
    current_ppr: etree.Element | None,
    previous_ppr: etree.Element | None,
    author: str | None = None,
) -> etree.Element:
    """Generate <w:pPrChange> element.

    Returns an element to insert as last child of <w:pPr>.
    """
```

### 2) Run Property Builder

New utility class for constructing `<w:rPr>` elements:

```python
class RunPropertyBuilder:
    """Build <w:rPr> elements from Python parameters."""

    PROPERTY_MAP = {
        "bold": ("w:b", None),           # (element_name, value_attr)
        "italic": ("w:i", None),
        "underline": ("w:u", "w:val"),
        "strikethrough": ("w:strike", None),
        "font_size": ("w:sz", "w:val"),  # Needs twip conversion
        "font_name": ("w:rFonts", "w:ascii"),
        "color": ("w:color", "w:val"),
        "highlight": ("w:highlight", "w:val"),
        # ... etc
    }

    @classmethod
    def build(cls, **kwargs) -> etree.Element:
        """Build <w:rPr> from keyword arguments."""

    @classmethod
    def merge(cls, base: etree.Element, updates: dict) -> etree.Element:
        """Merge updates into existing <w:rPr>, return new element."""

    @classmethod
    def diff(cls, old: etree.Element, new: etree.Element) -> dict:
        """Return dict of properties that differ between old and new."""

    @classmethod
    def extract(cls, rpr: etree.Element) -> dict:
        """Extract Python dict from <w:rPr> element."""
```

### 3) Paragraph Property Builder

Similar utility for `<w:pPr>`:

```python
class ParagraphPropertyBuilder:
    """Build <w:pPr> elements from Python parameters."""

    PROPERTY_MAP = {
        "alignment": ("w:jc", "w:val"),
        "spacing_before": ("w:spacing", "w:before"),  # Needs twip conversion
        "spacing_after": ("w:spacing", "w:after"),
        "line_spacing": ("w:spacing", "w:line"),
        "indent_left": ("w:ind", "w:left"),           # Needs twip conversion
        # ... etc
    }
```

### 4) Document.format_tracked() Implementation

```python
def format_tracked(self, text: str, *, scope=None, occurrence="first",
                   author=None, **format_kwargs) -> FormatResult:
    # 1. Find target text using existing TextSearch
    spans = self._find_text_spans(text, scope=scope)

    # 2. Handle occurrence selection
    target_spans = self._select_occurrences(spans, occurrence)

    # 3. For each span:
    for span in target_spans:
        # a. Split runs if text doesn't align with run boundaries
        affected_runs = self._split_runs_for_span(span)

        # b. For each affected run:
        for run in affected_runs:
            # i. Get current <w:rPr> (or create empty one)
            current_rpr = self._get_or_create_rpr(run)

            # ii. Deep copy as previous state
            previous_rpr = copy.deepcopy(current_rpr)

            # iii. Apply format changes to current
            RunPropertyBuilder.merge(current_rpr, format_kwargs)

            # iv. Check if anything actually changed
            if not RunPropertyBuilder.diff(previous_rpr, current_rpr):
                continue  # No-op, skip

            # v. Create and append <w:rPrChange>
            rpr_change = self._xml_generator.create_run_property_change(
                current_rpr, previous_rpr, author
            )
            current_rpr.append(rpr_change)

    # 4. Return result
    return FormatResult(...)
```

### 5) Accept/Reject Extensions

Extend `accept_all_changes()`:

```python
def accept_all_changes(self) -> AcceptResult:
    # ... existing ins/del handling ...

    # Handle rPrChange: remove element, keep parent rPr as-is
    for rpr_change in self._root.xpath("//w:rPrChange", namespaces=NS):
        rpr_change.getparent().remove(rpr_change)

    # Handle pPrChange: remove element, keep parent pPr as-is
    for ppr_change in self._root.xpath("//w:pPrChange", namespaces=NS):
        ppr_change.getparent().remove(ppr_change)
```

Extend `reject_all_changes()`:

```python
def reject_all_changes(self) -> RejectResult:
    # ... existing ins/del handling ...

    # Handle rPrChange: replace parent rPr with previous rPr from change
    for rpr_change in self._root.xpath("//w:rPrChange", namespaces=NS):
        parent_rpr = rpr_change.getparent()
        previous_rpr = rpr_change.find("w:rPr", namespaces=NS)

        # Replace parent's children with previous state
        for child in list(parent_rpr):
            parent_rpr.remove(child)
        for child in previous_rpr:
            parent_rpr.append(copy.deepcopy(child))
```

### 6) Unit Conversion Utilities

```python
def points_to_twips(points: float) -> int:
    """Convert points to twips (1 point = 20 twips)."""
    return int(points * 20)

def inches_to_twips(inches: float) -> int:
    """Convert inches to twips (1 inch = 1440 twips)."""
    return int(inches * 1440)

def points_to_half_points(points: float) -> int:
    """Convert points to half-points for font size (w:sz uses half-points)."""
    return int(points * 2)
```

## File Changes

### New Files

| File | Purpose | Est. Lines |
|------|---------|------------|
| `src/python_docx_redline/format_builder.py` | RunPropertyBuilder, ParagraphPropertyBuilder | ~250 |
| `tests/test_format_tracked.py` | Unit tests for format_tracked() | ~400 |
| `tests/test_format_paragraph_tracked.py` | Unit tests for paragraph formatting | ~200 |
| `tests/test_format_accept_reject.py` | Tests for accept/reject with formatting | ~150 |

### Modified Files

| File | Changes | Est. Lines |
|------|---------|------------|
| `tracked_xml.py` | Add create_run_property_change(), create_paragraph_property_change() | +80 |
| `document.py` | Add format_tracked(), format_paragraph_tracked() | +250 |
| `document.py` | Extend accept_all_changes(), reject_all_changes() | +50 |
| `results.py` | Add FormatResult dataclass | +20 |
| `errors.py` | Add FormatError if needed | +10 |

**Total estimated: ~1,400 lines**

## Test Plan

### Unit Tests

1. **RunPropertyBuilder tests**
   - Build rPr from kwargs
   - Merge updates into existing rPr
   - Diff two rPr elements
   - Extract dict from rPr element
   - Handle all supported properties

2. **ParagraphPropertyBuilder tests**
   - Similar coverage for paragraph properties
   - Unit conversion accuracy (twips)

3. **TrackedXMLGenerator tests**
   - create_run_property_change() produces valid XML
   - Proper ID sequencing
   - MS365 identity attributes included

### Integration Tests

4. **format_tracked() happy path**
   - Format found text
   - Verify XML structure
   - Verify Word opens document correctly

5. **format_tracked() edge cases**
   - Text not found
   - Multiple occurrences
   - Cross-run text
   - Nested in existing tracked changes

6. **format_paragraph_tracked() tests**
   - Alignment changes
   - Spacing changes
   - Indentation changes

7. **Accept/reject tests**
   - Accept removes rPrChange, keeps current
   - Reject restores previous from rPrChange
   - Mixed text and format changes

### Round-Trip Tests

8. **Create → Save → Open in Word → Verify**
   - Tracked format changes appear in revision pane
   - Accept in Word produces expected result
   - Reject in Word produces expected result

## Developer Task Breakdown

### Phase 1: Core Infrastructure (~3 tasks)

1. **RunPropertyBuilder class**
   - Property mapping for all supported run properties
   - build(), merge(), diff(), extract() methods
   - Unit conversion utilities
   - Unit tests

2. **ParagraphPropertyBuilder class**
   - Property mapping for paragraph properties
   - Similar methods
   - Unit tests

3. **TrackedXMLGenerator extensions**
   - create_run_property_change()
   - create_paragraph_property_change()
   - Unit tests

### Phase 2: Document Methods (~2 tasks)

4. **format_tracked() implementation**
   - Text search integration
   - Run splitting for partial formatting
   - FormatResult return type
   - Integration tests

5. **format_paragraph_tracked() implementation**
   - Paragraph finding logic
   - Property application
   - Integration tests

### Phase 3: Change Management (~2 tasks)

6. **accept_all_changes() extension**
   - Handle rPrChange elements
   - Handle pPrChange elements
   - Tests

7. **reject_all_changes() extension**
   - Restore previous formatting
   - Tests

### Phase 4: Batch Operations (~1 task)

8. **apply_edits() and YAML support**
   - Add format_tracked type
   - Add format_paragraph_tracked type
   - Documentation and examples

## Rollout Plan

1. **MVP Release**: format_tracked() for character formatting only
2. **Follow-up**: format_paragraph_tracked() for paragraph formatting
3. **Documentation**: Add to PROPOSED_API.md, create examples
4. **Future**: Consider format comparison in compare_to()

## Open Questions

1. **Style application**: Should we support applying named styles (e.g., "Heading 1") with tracking? This is more complex as it affects both rPr and pPr.

2. **Partial undo**: Should we support accepting/rejecting individual format changes by ID?

3. **Format diffing**: Should compare_to() detect and track formatting differences between documents?

## References

- [OOXML w:rPrChange Specification](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_rPrChange_topic_ID0EC5SW.html)
- [OOXML w:pPrChange Specification](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_pPrChange_topic_ID0EXXRW.html)
- [Eric White - Detecting Tracked Revisions](http://www.ericwhite.com/blog/using-xml-dom-to-detect-tracked-revisions-in-an-open-xml-wordprocessingml-document/)
- [doNotTrackFormatting Setting](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_doNotTrackFormatting_topic_ID0ESWYX.html)
