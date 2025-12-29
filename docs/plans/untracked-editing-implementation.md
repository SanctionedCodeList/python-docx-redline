# Implementation Plan: Untracked Editing Support

**Status**: Planning
**Date**: 2024-12-28
**Goal**: Extend python-docx-redline to support both tracked and untracked editing with a unified API

---

## Executive Summary

Add untracked editing capability across the entire library, making python-docx-redline the go-to solution for all Word document editing—not just tracked changes.

### API Design

```python
# New generic API (track=False by default)
doc.replace("old", "new")                    # Untracked (silent edit)
doc.replace("old", "new", track=True)        # Tracked
doc.insert("text", after="anchor")           # Untracked
doc.delete("text to remove")                 # Untracked

# Backwards-compatible aliases (always tracked)
doc.replace_tracked("old", "new")            # Same as replace(..., track=True)
doc.insert_tracked("text", after="anchor")   # Same as insert(..., track=True)
doc.delete_tracked("text")                   # Same as delete(..., track=True)
```

---

## Phase 1: Core Infrastructure

### 1.1 Add Plain Run Generation to TrackedXMLGenerator

**File**: `src/python_docx_redline/tracked_xml.py`

Add methods to generate plain `<w:r>` elements without tracked change wrappers:

```python
def create_plain_run(self, text: str, source_run: Element | None = None) -> Element:
    """Generate a plain <w:r> element without tracked change wrapper.

    Args:
        text: The text content
        source_run: Optional run to copy formatting from

    Returns:
        lxml Element for the run
    """

def create_plain_runs(self, text: str, source_run: Element | None = None) -> list[Element]:
    """Generate plain runs, handling markdown formatting.

    Supports same markdown as create_insertion:
    - **bold** -> <w:b/>
    - *italic* -> <w:i/>
    - ++underline++ -> <w:u/>
    - ~~strikethrough~~ -> <w:strike/>
    """
```

**Implementation notes**:
- Reuse `_generate_run()` and `_generate_runs()` but without the `<w:ins>` wrapper
- Preserve formatting from source_run's `<w:rPr>` when provided
- Handle `xml:space="preserve"` for whitespace

### 1.2 Update TextSearch to Handle Deleted Text

**File**: `src/python_docx_redline/text_search.py`

Add parameter to control whether deleted text is searchable:

```python
def find_text(
    self,
    text: str,
    paragraphs: list[Element],
    regex: bool = False,
    normalize_special_chars: bool = True,
    fuzzy: FuzzyConfig | None = None,
    include_deleted: bool = False,  # NEW
) -> list[TextSpan]:
```

**Implementation**:
- When `include_deleted=False`, skip text inside `<w:del>` and `<w:delText>` elements
- Add helper method `_is_deleted_content(element) -> bool`
- Modify text extraction to filter based on this flag

### 1.3 Propagate `include_deleted` to Document Methods

Update `find_all()` signature:

```python
def find_all(
    self,
    text: str,
    scope: str | dict | None = None,
    regex: bool = False,
    normalize_special_chars: bool = True,
    fuzzy: float | dict | None = None,
    include_deleted: bool = False,  # NEW
) -> list[Match]:
```

---

## Phase 2: Core Edit Operations

### 2.1 Refactor TrackedChangeOperations

**File**: `src/python_docx_redline/operations/tracked_changes.py`

**Option A**: Rename class to `EditOperations` (breaking change for internal API)
**Option B**: Keep class name, add `track` parameter to all methods

**Recommendation**: Option B - less churn, clearer internal naming

Add `track: bool = False` parameter to:
- `insert()`
- `delete()`
- `replace()`
- `move()`

**Implementation pattern for each method**:

```python
def replace(
    self,
    find: str,
    replace: str,
    author: str | None = None,
    scope: str | dict | Any | None = None,
    occurrence: int | list[int] | str = "first",
    regex: bool = False,
    normalize_special_chars: bool = True,
    track: bool = False,  # NEW - default untracked
    # ... other params
) -> None:
    # ... find matches (existing code) ...

    for match in reversed(target_matches):
        matched_text = match.text
        replacement_text = self._expand_replacement(match, replace, regex)

        if track:
            # Existing tracked logic
            deletion_xml = self._document._xml_generator.create_deletion(matched_text, author)
            insertion_xml = self._document._xml_generator.create_insertion(replacement_text, author)
            elements = self._parse_xml_elements(f"{deletion_xml}\n    {insertion_xml}")
            self._replace_match_with_elements(match, elements)
        else:
            # New untracked logic
            new_run = self._document._xml_generator.create_plain_run(
                replacement_text,
                source_run=match.runs[0]
            )
            self._replace_match_with_element(match, new_run)
```

### 2.2 Add Helper for Untracked Deletion

For `delete()` with `track=False`, we simply remove runs without replacement:

```python
def _remove_match(self, match: TextSpan) -> None:
    """Remove matched text without any replacement.

    Handles:
    - Single run matches (remove run or split if partial)
    - Multi-run matches (remove all, preserve before/after text)
    - Runs inside tracked change wrappers
    """
```

**Implementation notes**:
- Similar to `_replace_match_with_element` but with no replacement
- Need to handle partial matches (split run, keep before/after text)
- Clean up empty wrappers after removal

---

## Phase 3: Document Class API

### 3.1 Add Generic Methods

**File**: `src/python_docx_redline/document.py`

Add new generic methods that delegate to `TrackedChangeOperations`:

```python
def insert(
    self,
    text: str,
    after: str | None = None,
    before: str | None = None,
    author: str | None = None,
    scope: str | dict | Any | None = None,
    occurrence: int | list[int] | str = "first",
    regex: bool = False,
    normalize_special_chars: bool = True,
    track: bool = False,
    fuzzy: float | dict[str, Any] | None = None,
) -> None:
    """Insert text after or before a specific location.

    Args:
        text: The text to insert
        after: Insert after this text
        before: Insert before this text
        track: If True, insert as tracked change (default: False)
        # ... other params same as insert_tracked
    """
    self._tracked_ops.insert(
        text, after=after, before=before, author=author,
        scope=scope, occurrence=occurrence, regex=regex,
        normalize_special_chars=normalize_special_chars,
        track=track, fuzzy=fuzzy,
    )

def delete(
    self,
    text: str,
    author: str | None = None,
    scope: str | dict | Any | None = None,
    occurrence: int | list[int] | str = "first",
    regex: bool = False,
    normalize_special_chars: bool = True,
    track: bool = False,
    fuzzy: float | dict[str, Any] | None = None,
) -> None:
    """Delete text from the document.

    Args:
        text: The text to delete
        track: If True, delete as tracked change (default: False)
        # ... other params
    """

def replace(
    self,
    find: str,
    replace: str,
    author: str | None = None,
    scope: str | dict | Any | None = None,
    occurrence: int | list[int] | str = "first",
    regex: bool = False,
    normalize_special_chars: bool = True,
    track: bool = False,
    show_context: bool = False,
    check_continuity: bool = False,
    context_chars: int = 50,
    fuzzy: float | dict[str, Any] | None = None,
) -> None:
    """Find and replace text in the document.

    Args:
        find: Text to find
        replace: Replacement text
        track: If True, show as tracked change (default: False)
        # ... other params
    """

def move(
    self,
    text: str,
    after: str | None = None,
    before: str | None = None,
    author: str | None = None,
    source_scope: str | dict | Any | None = None,
    dest_scope: str | dict | Any | None = None,
    regex: bool = False,
    normalize_special_chars: bool = True,
    track: bool = False,
) -> None:
    """Move text to a new location.

    Args:
        text: The text to move
        after: Move to after this text
        before: Move to before this text
        track: If True, show as tracked move (default: False)
        # ... other params
    """
```

### 3.2 Convert Existing Methods to Aliases

Update existing `*_tracked` methods to be aliases:

```python
def insert_tracked(
    self,
    text: str,
    after: str | None = None,
    before: str | None = None,
    author: str | None = None,
    scope: str | dict | Any | None = None,
    occurrence: int | list[int] | str = "first",
    regex: bool = False,
    normalize_special_chars: bool = True,
    fuzzy: float | dict[str, Any] | None = None,
) -> None:
    """Insert text with tracked changes.

    This is an alias for insert(..., track=True).
    See insert() for full documentation.
    """
    self.insert(
        text, after=after, before=before, author=author,
        scope=scope, occurrence=occurrence, regex=regex,
        normalize_special_chars=normalize_special_chars,
        track=True, fuzzy=fuzzy,
    )

# Similar for delete_tracked, replace_tracked, move_tracked
```

---

## Phase 4: Extended Operations

### 4.1 Table Operations

**File**: `src/python_docx_redline/operations/tables.py`

Add `track` parameter to:
- `replace_in_table()`
- `replace_in_cell()`
- Any other table edit methods

### 4.2 Header/Footer Operations

**File**: `src/python_docx_redline/operations/header_footer.py`

Add `track` parameter to:
- `replace_in_header()`
- `replace_in_footer()`

### 4.3 Formatting Operations

**File**: `src/python_docx_redline/operations/formatting.py`

Review if formatting changes should support `track` parameter:
- `apply_bold()`, `apply_italic()`, etc.
- These create `<w:rPrChange>` for tracking - should also support untracked

### 4.4 Notes Operations

**File**: `src/python_docx_redline/operations/notes.py`

Add `track` parameter to footnote/endnote editing if applicable.

---

## Phase 5: Batch Operations

### 5.1 Update apply_edits()

**File**: `src/python_docx_redline/operations/batch.py`

Support `track` field in edit dictionaries:

```python
edits = [
    {
        "type": "replace",
        "find": "old text",
        "replace": "new text",
        "track": False,  # Optional, defaults to False
    },
    {
        "type": "insert",
        "text": "inserted content",
        "after": "anchor",
        "track": True,  # This one is tracked
    },
]

doc.apply_edits(edits)
```

Also add a global default:

```python
def apply_edits(
    self,
    edits: list[dict],
    stop_on_error: bool = False,
    default_track: bool = False,  # NEW
) -> list[EditResult]:
    """Apply multiple edits.

    Args:
        edits: List of edit specifications
        stop_on_error: Stop on first error
        default_track: Default value for 'track' if not specified per-edit
    """
```

### 5.2 Update apply_edit_file()

YAML/JSON files should support `track` field:

```yaml
document: input.docx
output: output.docx
default_track: false  # Global default

edits:
  - type: replace
    find: "old"
    replace: "new"
    # uses default_track: false

  - type: insert
    text: "tracked insertion"
    after: "anchor"
    track: true  # Override for this edit
```

---

## Phase 6: Testing

### 6.1 Unit Tests for Untracked Operations

**File**: `tests/test_untracked_editing.py`

```python
class TestUntrackedInsert:
    def test_insert_untracked_basic(self):
        """Insert text without tracking."""

    def test_insert_untracked_preserves_formatting(self):
        """Inserted text inherits formatting from context."""

    def test_insert_untracked_with_markdown(self):
        """Markdown formatting works in untracked inserts."""

    def test_insert_untracked_inside_tracked_wrapper(self):
        """Untracked insert inside <w:ins> works correctly."""


class TestUntrackedDelete:
    def test_delete_untracked_basic(self):
        """Delete text without tracking."""

    def test_delete_untracked_partial_run(self):
        """Delete partial run content."""

    def test_delete_untracked_multi_run(self):
        """Delete text spanning multiple runs."""

    def test_delete_untracked_preserves_surrounding(self):
        """Text before/after deletion is preserved."""


class TestUntrackedReplace:
    def test_replace_untracked_basic(self):
        """Replace text without tracking."""

    def test_replace_untracked_preserves_formatting(self):
        """Replacement inherits formatting from original."""

    def test_replace_untracked_with_regex(self):
        """Regex replacement works untracked."""

    def test_replace_untracked_occurrence_all(self):
        """Replace all occurrences untracked."""


class TestUntrackedMove:
    def test_move_untracked_basic(self):
        """Move text without tracking markers."""


class TestMixedTracking:
    def test_mixed_tracked_and_untracked(self):
        """Combine tracked and untracked edits in sequence."""

    def test_batch_mixed_tracking(self):
        """apply_edits with mixed track values."""
```

### 6.2 Tests for include_deleted Parameter

**File**: `tests/test_text_search.py`

```python
class TestIncludeDeleted:
    def test_find_excludes_deleted_by_default(self):
        """find_all() skips text inside <w:del>."""

    def test_find_includes_deleted_when_requested(self):
        """find_all(include_deleted=True) finds deleted text."""

    def test_replace_on_deleted_text(self):
        """Can replace deleted text when include_deleted=True."""
```

### 6.3 Integration Tests

**File**: `tests/test_integration_untracked.py`

```python
def test_full_workflow_untracked():
    """Complete workflow using only untracked edits."""
    doc = Document("template.docx")

    # Populate template
    doc.replace("{{NAME}}", "John Doe")
    doc.replace("{{DATE}}", "2024-12-28")
    doc.delete("{{OPTIONAL_SECTION}}")

    doc.save("output.docx")

    # Verify no tracked changes in output
    result = Document("output.docx")
    assert not result.has_tracked_changes()


def test_round_trip_mixed():
    """Document with both tracked and untracked edits."""
    doc = Document("input.docx")

    # Silent fixes
    doc.replace("teh", "the")  # Untracked typo fix

    # Substantive changes (tracked)
    doc.replace("30 days", "45 days", track=True)

    doc.save("output.docx")

    # Verify only substantive change is tracked
    result = Document("output.docx")
    changes = result.get_tracked_changes()
    assert len(changes) == 2  # One deletion + one insertion
    assert "30 days" in [c.text for c in changes if c.type == "delete"]
```

---

## Phase 7: Documentation

### 7.1 Update Docstrings

All modified methods need updated docstrings explaining the `track` parameter.

### 7.2 Update SKILL.md

Update the docx skill to reflect new capabilities:

```markdown
## Decision Tree

| Task | Tool | Guide |
|------|------|-------|
| **Read/extract text** | pandoc or python-docx-redline | [reading.md](./reading.md) |
| **Create new document** | python-docx | See below |
| **Edit (any mode)** | python-docx-redline | [editing.md](./editing.md) |
| **Add comments** | python-docx-redline | [comments.md](./comments.md) |
```

### 7.3 Update editing.md

Replace the python-docx content with python-docx-redline examples for untracked editing.

### 7.4 Add Migration Guide

Document for users upgrading:
- New `insert()`, `delete()`, `replace()`, `move()` methods
- `*_tracked` methods still work, now aliases
- New `track` parameter
- New `include_deleted` parameter for searches

---

## Phase 8: Package Updates

### 8.1 Update __init__.py Exports

Export new methods (they're already on Document, but verify all needed).

### 8.2 Update Package Description

Update pyproject.toml description to reflect broader scope:
```toml
description = "High-level Python API for editing Word documents, with optional tracked changes support"
```

### 8.3 Consider Package Rename?

**Decision needed**: Should the package be renamed to reflect broader scope?
- Current: `python-docx-redline`
- Options: Keep as-is (redlining is still a key feature), or rename

**Recommendation**: Keep current name. The "redline" capability is the differentiator, and untracked editing is an extension of the same API.

---

## Implementation Order

### Sprint 1: Core Infrastructure (Days 1-2)
1. [ ] Add `create_plain_run()` to TrackedXMLGenerator
2. [ ] Add `include_deleted` to TextSearch
3. [ ] Add `_remove_match()` helper for untracked deletion
4. [ ] Unit tests for new helpers

### Sprint 2: Core Operations (Days 2-3)
1. [ ] Add `track` param to TrackedChangeOperations.insert()
2. [ ] Add `track` param to TrackedChangeOperations.delete()
3. [ ] Add `track` param to TrackedChangeOperations.replace()
4. [ ] Add `track` param to TrackedChangeOperations.move()
5. [ ] Unit tests for each operation

### Sprint 3: Document API (Day 3)
1. [ ] Add generic methods to Document class (insert, delete, replace, move)
2. [ ] Convert *_tracked methods to aliases
3. [ ] Update find_all() signature
4. [ ] Integration tests

### Sprint 4: Extended Operations (Day 4)
1. [ ] Update table operations
2. [ ] Update header/footer operations
3. [ ] Update formatting operations
4. [ ] Update notes operations
5. [ ] Tests for extended operations

### Sprint 5: Batch & Polish (Day 4-5)
1. [ ] Update apply_edits() with track support
2. [ ] Update apply_edit_file() with track support
3. [ ] Full integration test suite
4. [ ] Documentation updates
5. [ ] Skill updates

---

## Success Criteria

1. [ ] All existing tests pass (backwards compatibility)
2. [ ] New untracked operations work correctly
3. [ ] Mixed tracked/untracked workflows work
4. [ ] `find_all()` correctly excludes deleted text by default
5. [ ] Formatting is preserved in untracked edits
6. [ ] Batch operations support per-edit tracking
7. [ ] Documentation is complete
8. [ ] Skill is updated

---

## Open Questions

1. **Author parameter for untracked edits**: Should we still accept `author` param for untracked edits (ignored) or remove it from signature?
   - **Recommendation**: Keep it, ignore silently. Easier for users switching between modes.

2. **Return values**: Current `*_tracked` methods return None. Should new generic methods return something useful (e.g., the modified TextSpan)?
   - **Recommendation**: Keep returning None for now, consistent with existing API. Can add in future.

3. **Formatting preservation strategy**: When doing untracked replace, how to handle formatting?
   - **Recommendation**: Copy `<w:rPr>` from first matched run. Document this behavior.
