# python_docx_redline Refactoring Plan

**Based on:** CODE_REVIEW_REPORT.md (2025-12-09), Audit (2025-12-13)
**Goal:** Transform the codebase from prototype to production-grade library
**Last Updated:** 2025-12-13

---

## Current State (2025-12-13 Audit)

| Metric | Value | Target |
|--------|-------|--------|
| `document.py` lines | 7,522 | <500 |
| Document class methods | 131 | <30 (facade) |
| Longest method | 314 lines (`_apply_single_edit`) | <50 |
| Test count | 877 | Maintain |
| Coverage | 80%+ | 90%+ |
| Mypy error codes disabled | 9 | 0 |
| Print statements | 39 | 0 (use logging) |
| Namespace duplications | 6 files | 1 (constants.py) |

---

## Overview

This plan is organized into 6 phases, ordered by:
1. **Risk reduction** - Fix bugs that could corrupt documents first
2. **Foundation** - Establish patterns that make later work easier
3. **Incremental extraction** - Break up the God class piece by piece
4. **Polish** - Add production essentials like logging

Each phase can be completed independently with passing tests between phases.

---

## Phase 1: Critical Bug Fixes (High Priority)

**Goal:** Fix issues that could corrupt documents or leak resources.

### 1.1 Implement Proper Change ID Tracking

**File:** `tracked_xml.py`
**Issue:** `_get_max_change_id()` always returns 0, causing ID collisions

**Tasks:**
- [ ] Implement `_get_max_change_id()` to scan for existing `w:id` attributes on:
  - `w:ins` elements
  - `w:del` elements
  - `w:moveFrom` elements
  - `w:moveTo` elements
  - `w:moveFromRangeStart` / `w:moveToRangeStart`
- [ ] Add tests with documents containing existing tracked changes
- [ ] Verify IDs are unique after multiple edit sessions

**Estimated scope:** ~50 lines of code, ~5 tests

### 1.2 Fix Author Name XML Escaping

**File:** `tracked_xml.py`
**Issue:** Author names are interpolated without escaping

**Tasks:**
- [ ] Apply `_escape_xml()` to `author` parameter in all XML generation methods
- [ ] Add tests with author names containing `"`, `<`, `>`, `&`, `'`
- [ ] Consider creating elements with lxml instead of string interpolation (see Phase 3)

**Estimated scope:** ~10 lines of code, ~3 tests

### 1.3 Implement Proper Resource Cleanup

**File:** `document.py`
**Issue:** Temp directories may not be cleaned up

**Tasks:**
- [ ] Add `__del__` method to clean up `_temp_dir`
- [ ] Verify context manager (`__enter__`/`__exit__`) properly cleans up
- [ ] Add `close()` method for explicit cleanup
- [ ] Add tests verifying temp dirs are removed in all scenarios:
  - Normal save
  - Exception during processing
  - Context manager exit
  - Explicit close()

**Estimated scope:** ~30 lines of code, ~6 tests

### 1.4 Fix lxml FutureWarning Deprecations

**File:** `document.py` (lines ~4921, ~4945)
**Issue:** `if not end_run:` will break in future lxml versions

**Tasks:**
- [ ] Find all occurrences of boolean element checks
- [ ] Replace `if not elem:` with `if elem is None:`
- [ ] Replace `if elem:` with `if elem is not None:`
- [ ] Run tests with `-W error::FutureWarning` to catch any remaining issues

**Estimated scope:** ~10 lines of code, ~0 tests (existing tests verify)

---

## Phase 2: Foundation - Constants & Types (Medium Priority)

**Goal:** Establish shared infrastructure that makes later refactoring easier.

### 2.1 Create Centralized Constants Module

**New file:** `src/python_docx_redline/constants.py`

**Tasks:**
- [ ] Create `constants.py` with all namespace definitions:
  ```python
  # OOXML Namespaces
  W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
  W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
  W16DU_NS = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"
  REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
  CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

  # Common namespace maps
  NSMAP = {"w": W_NS}
  NSMAP_FULL = {"w": W_NS, "w14": W14_NS, "w15": W15_NS, "w16du": W16DU_NS}

  # Relationship types
  REL_TYPE_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  REL_TYPE_COMMENTS_EX = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
  # ... etc

  # Content types
  CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
  # ... etc

  # Magic numbers with names
  CONTEXT_CHARS_DEFAULT = 40
  MAX_BATCH_ITERATIONS = 100
  ```
- [ ] Update all imports across the codebase
- [ ] Remove duplicate namespace definitions

**Estimated scope:** ~100 lines new, ~50 lines modified across 6 files

### 2.2 Create Type Definitions Module

**New file:** `src/python_docx_redline/types.py`

**Tasks:**
- [ ] Define type aliases for lxml elements:
  ```python
  from typing import TypeAlias
  from lxml import etree

  Element: TypeAlias = etree._Element
  ElementTree: TypeAlias = etree._ElementTree
  ```
- [ ] Create Protocol for TextSpan-like objects
- [ ] Create Protocol for scope evaluators
- [ ] Update type hints across codebase to use these types
- [ ] Run mypy and fix any new type errors

**Estimated scope:** ~50 lines new, ~100 lines modified

### 2.3 Add Logging Infrastructure

**New file:** `src/python_docx_redline/logging.py` (or add to `__init__.py`)

**Tasks:**
- [ ] Create module-level logger: `logger = logging.getLogger("python_docx_redline")`
- [ ] Replace all `print()` statements with appropriate log levels
- [ ] Add debug logging for:
  - Document load/save operations
  - Text search matches
  - XML modifications
- [ ] Document how users can configure logging

**Estimated scope:** ~20 lines new, ~30 lines modified

---

## Phase 3: Extract Package Management (Medium Priority)

**Goal:** Extract OOXML package handling into dedicated classes.

### 3.1 Create Package Manager

**New file:** `src/python_docx_redline/package.py`

**Tasks:**
- [ ] Create `OOXMLPackage` class:
  ```python
  class OOXMLPackage:
      """Manages the OOXML ZIP package structure."""

      def __init__(self, path: Path):
          self.path = path
          self._temp_dir: Path | None = None

      def extract(self) -> Path:
          """Extract to temp directory, return path."""

      def get_part(self, part_name: str) -> Element | None:
          """Load and parse an XML part."""

      def set_part(self, part_name: str, element: Element) -> None:
          """Write an XML part."""

      def save(self, output_path: Path) -> None:
          """Repack the ZIP file."""

      def close(self) -> None:
          """Clean up temp directory."""

      def __enter__(self) -> "OOXMLPackage": ...
      def __exit__(self, ...): ...
  ```
- [ ] Move extraction logic from `Document._extract_docx()`
- [ ] Move save/repack logic from `Document.save()`
- [ ] Update `Document` to use `OOXMLPackage`
- [ ] Add comprehensive tests

**Estimated scope:** ~200 lines new, ~100 lines removed from document.py

### 3.2 Create Relationship Manager

**New file:** `src/python_docx_redline/relationships.py`

**Tasks:**
- [ ] Create `RelationshipManager` class:
  ```python
  class RelationshipManager:
      """Manages .rels files in OOXML packages."""

      def __init__(self, package: OOXMLPackage, part_name: str):
          """Initialize for a specific part's relationships."""

      def get_relationship(self, rel_type: str) -> str | None:
          """Get target for a relationship type."""

      def add_relationship(self, rel_type: str, target: str) -> str:
          """Add relationship, return assigned rId."""

      def remove_relationship(self, rel_type: str) -> bool:
          """Remove relationship by type."""

      def save(self) -> None:
          """Write changes back to .rels file."""
  ```
- [ ] Extract all `_ensure_*_relationship()` methods
- [ ] Extract relationship removal logic from `delete_all_comments()`
- [ ] Update `Document` to use `RelationshipManager`

**Estimated scope:** ~150 lines new, ~200 lines removed from document.py

### 3.3 Create Content Type Manager

**New file:** `src/python_docx_redline/content_types.py`

**Tasks:**
- [ ] Create `ContentTypeManager` class:
  ```python
  class ContentTypeManager:
      """Manages [Content_Types].xml."""

      def __init__(self, package: OOXMLPackage):
          """Initialize from package."""

      def get_content_type(self, part_name: str) -> str | None:
          """Get content type for a part."""

      def add_override(self, part_name: str, content_type: str) -> None:
          """Add content type override."""

      def remove_override(self, part_name: str) -> bool:
          """Remove content type override."""

      def save(self) -> None:
          """Write changes back."""
  ```
- [ ] Extract all `_ensure_*_content_type()` methods
- [ ] Update `Document` to use `ContentTypeManager`

**Estimated scope:** ~100 lines new, ~150 lines removed from document.py

---

## Phase 4: Extract Domain Operations (High Priority)

**Goal:** Break the God class into focused domain classes.

### 4.1 Create Tracked Change Operations Class

**New file:** `src/python_docx_redline/operations/tracked_changes.py`

**Tasks:**
- [ ] Create `TrackedChangeOperations` class:
  ```python
  class TrackedChangeOperations:
      """Handles insert, delete, replace, move with tracking."""

      def __init__(self, document: "Document"):
          self._doc = document
          self._xml_gen = TrackedXMLGenerator(document)
          self._search = TextSearch()

      def insert(self, text: str, after: str | None = None,
                 before: str | None = None, **kwargs) -> None:
          """Insert text with tracked changes."""

      def delete(self, text: str, **kwargs) -> None:
          """Delete text with tracked changes."""

      def replace(self, find: str, replace: str, **kwargs) -> None:
          """Replace text with tracked changes."""

      def move(self, text: str, after: str | None = None,
               before: str | None = None, **kwargs) -> None:
          """Move text with tracked changes."""
  ```
- [ ] Move `insert_tracked()`, `delete_tracked()`, `replace_tracked()`, `move_tracked()`
- [ ] Move `_insert_after_match()`, `_insert_before_match()`
- [ ] Move `_replace_match_with_element()`, `_replace_match_with_elements()`
- [ ] Move `_split_and_replace_in_run()`, `_split_and_replace_in_run_multiple()`
- [ ] Document class delegates to this class
- [ ] Maintain backward compatibility with existing API

**Estimated scope:** ~500 lines moved, ~50 lines new glue code

### 4.2 Create Change Management Class

**New file:** `src/python_docx_redline/operations/change_management.py`

**Tasks:**
- [ ] Create `ChangeManagement` class:
  ```python
  class ChangeManagement:
      """Handles accepting/rejecting tracked changes."""

      def accept_all(self) -> None: ...
      def reject_all(self) -> None: ...
      def accept_insertions(self) -> int: ...
      def reject_insertions(self) -> int: ...
      def accept_deletions(self) -> int: ...
      def reject_deletions(self) -> int: ...
      def accept_change(self, change_id: str | int) -> None: ...
      def reject_change(self, change_id: str | int) -> None: ...
      def accept_by_author(self, author: str) -> int: ...
      def reject_by_author(self, author: str) -> int: ...
  ```
- [ ] Move all accept/reject methods from Document
- [ ] Move `_unwrap_element()`, `_unwrap_deletion()`, `_remove_element()`

**Estimated scope:** ~250 lines moved

### 4.3 Create Comment Operations Class

**New file:** `src/python_docx_redline/operations/comments.py`

**Tasks:**
- [ ] Create `CommentOperations` class:
  ```python
  class CommentOperations:
      """Handles comment reading, adding, deleting."""

      @property
      def all(self) -> list[Comment]: ...
      def get(self, author: str | None = None,
              scope: Any = None) -> list[Comment]: ...
      def add(self, text: str, on: str | None = None,
              reply_to: Comment | None = None, **kwargs) -> Comment: ...
      def delete(self, comment: Comment | str | int) -> None: ...
      def delete_all(self) -> None: ...
      def resolve(self, comment: Comment | str | int) -> None: ...
      def unresolve(self, comment: Comment | str | int) -> None: ...
  ```
- [ ] Move `comments` property, `get_comments()`, `add_comment()`
- [ ] Move `delete_all_comments()`, `_delete_comment()`
- [ ] Move all comment helper methods (`_load_comments_xml()`, etc.)

**Estimated scope:** ~600 lines moved

### 4.4 Create Footnote/Endnote Operations Class

**New file:** `src/python_docx_redline/operations/notes.py`

**Tasks:**
- [ ] Create `NoteOperations` class for footnotes and endnotes
- [ ] Move `footnotes`, `endnotes` properties
- [ ] Move `add_footnote()`, `add_endnote()`
- [ ] Move all footnote/endnote helper methods

**Estimated scope:** ~400 lines moved

### 4.5 Create Pattern Helpers Class

**New file:** `src/python_docx_redline/operations/patterns.py`

**Tasks:**
- [ ] Create `PatternHelpers` class:
  ```python
  class PatternHelpers:
      """Common document editing patterns."""

      def normalize_currency(self, **kwargs) -> int: ...
      def normalize_dates(self, **kwargs) -> int: ...
      def update_section_references(self, **kwargs) -> int: ...
  ```
- [ ] Move `normalize_currency()`, `normalize_dates()`, `update_section_references()`

**Estimated scope:** ~300 lines moved

### 4.6 Create Style Operations Class

**New file:** `src/python_docx_redline/operations/styles.py`

**Tasks:**
- [ ] Create `StyleOperations` class:
  ```python
  class StyleOperations:
      """Handles paragraph and text styling."""

      def apply_style(self, find: str, style: str, **kwargs) -> int: ...
      def format_text(self, find: str, bold: bool | None = None,
                      italic: bool | None = None, **kwargs) -> int: ...
      def copy_format(self, from_text: str, to_text: str, **kwargs) -> int: ...
  ```
- [ ] Move `apply_style()`, `format_text()`, `copy_format()`

**Estimated scope:** ~200 lines moved

---

## Phase 5: Refactor XML Generation (Medium Priority)

**Goal:** Replace string-based XML generation with proper lxml element construction.

### 5.1 Refactor TrackedXMLGenerator to Use lxml

**File:** `tracked_xml.py`

**Tasks:**
- [ ] Create helper method to build elements:
  ```python
  def _create_element(self, tag: str, attrib: dict[str, str],
                      nsmap: dict[str, str] | None = None) -> Element:
      """Create an element with proper namespace handling."""
  ```
- [ ] Refactor `create_insertion()` to build elements directly
- [ ] Refactor `create_deletion()` to build elements directly
- [ ] Refactor `create_move_from()` and `create_move_to()`
- [ ] Remove string interpolation entirely
- [ ] Update Document to receive elements instead of strings

**Estimated scope:** ~200 lines modified

### 5.2 Remove XML String Wrapping Pattern

**File:** `document.py` (multiple locations)

**Tasks:**
- [ ] Create helper in package.py or a new xml_utils.py:
  ```python
  def parse_xml_fragment(xml: str, nsmap: dict[str, str]) -> Element:
      """Parse an XML fragment with namespace context."""
  ```
- [ ] Replace all instances of the wrapped_xml pattern
- [ ] Eventually eliminate once 5.1 is complete

**Estimated scope:** ~100 lines modified

---

## Phase 6: Production Hardening (Lower Priority)

**Goal:** Add production-grade features.

### 6.1 Integrate Validation

**Tasks:**
- [ ] Add `validate_on_save` option to Document
- [ ] Run validation automatically before save (optional)
- [ ] Improve validation coverage to 80%+
- [ ] Add clear error messages for validation failures

### 6.2 Add Transaction Support

**Tasks:**
- [ ] Create `DocumentTransaction` context manager:
  ```python
  with doc.transaction() as tx:
      doc.insert_tracked(...)
      doc.delete_tracked(...)
      # If exception, rollback to pre-transaction state
  ```
- [ ] Implement snapshot/restore mechanism
- [ ] Add tests for rollback scenarios

### 6.3 Add Integration Tests

**Tasks:**
- [ ] Create `tests/integration/` directory
- [ ] Add test documents created by Microsoft Word
- [ ] Test opening real documents
- [ ] Test round-trip (open → edit → save → open)
- [ ] Test documents with existing tracked changes

### 6.4 Add Thread Safety Documentation

**Tasks:**
- [ ] Document that Document objects are not thread-safe
- [ ] Consider adding optional locking for multi-threaded access
- [ ] Add warnings to documentation

---

## Implementation Order Recommendation

**Epic:** docx_redline-6jg

### Foundation (No Dependencies)
| Phase | Task | Beads ID |
|-------|------|----------|
| 1 | Centralize namespace constants | docx_redline-7ds |
| 2 | Replace print with logging | docx_redline-4zk |

### Infrastructure Extraction (Sequential)
| Phase | Task | Beads ID | Depends On |
|-------|------|----------|------------|
| 3 | Extract OOXMLPackage | docx_redline-7g1 | Phase 1 |
| 4 | Extract RelationshipManager | docx_redline-j11 | Phase 3 |
| 5 | Extract ContentTypeManager | docx_redline-1jg | Phase 3 |

### Domain Operations (Parallel after Infrastructure)
| Phase | Task | Beads ID | Depends On |
|-------|------|----------|------------|
| 6 | Extract TrackedChangeOperations | docx_redline-8no | Phase 4 |
| 7 | Extract CommentOperations | docx_redline-j42 | Phase 5 |
| 8 | Extract ChangeManagement | docx_redline-de0 | Phase 6 |
| 9 | Extract FormatOperations | docx_redline-hxd | Phase 6 |
| 10 | Extract TableOperations | docx_redline-xr8 | Phase 4 |
| 11 | Extract NoteOperations | docx_redline-9qc | Phase 5 |

### Cleanup (After Extraction)
| Phase | Task | Beads ID | Depends On |
|-------|------|----------|------------|
| 12 | Enable stricter mypy | docx_redline-hfx | - |
| 13 | Break up large methods | docx_redline-7af | Phase 6, 9 |
| 14 | Clean up TODO comments | docx_redline-b94 | - |
| 15 | Validation test coverage | docx_redline-j6i | - |

### Commands
```bash
bd ready              # Show tasks with no blockers
bd blocked            # Show blocked tasks
bd show <id>          # View task details
bd update <id> -s in_progress  # Start work
bd close <id>         # Complete task
```

---

## Success Criteria

After refactoring, the codebase should have:

| Criteria | Target | Verification |
|----------|--------|--------------|
| document.py size | <500 lines | `wc -l src/python_docx_redline/document.py` |
| Document class methods | <30 | Facade only, delegates to operations |
| Maximum method size | <50 lines | No method exceeds 50 lines |
| Test count | ≥877 | All existing tests pass |
| Coverage | ≥90% | `pytest --cov-fail-under=90` |
| Mypy strictness | 0 disabled codes | Enable all error codes |
| Print statements | 0 | Use logging module |
| FutureWarnings | 0 | `pytest -W error::FutureWarning` |
| Resource cleanup | Verified | Tests confirm temp dir cleanup |
| Backward compatibility | 100% | Public API unchanged |

---

## Migration Guide for Users

The refactoring maintains backward compatibility. The `Document` class API remains unchanged - it simply delegates to the new classes internally.

For users who want to use the new classes directly:

```python
# Old way (still works)
doc = Document("file.docx")
doc.insert_tracked("text", after="anchor")

# New way (more explicit)
from python_docx_redline.operations import TrackedChangeOperations
doc = Document("file.docx")
ops = TrackedChangeOperations(doc)
ops.insert("text", after="anchor")
```

Both approaches will be supported.
