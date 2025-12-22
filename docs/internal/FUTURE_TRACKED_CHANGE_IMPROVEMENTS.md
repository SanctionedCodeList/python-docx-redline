# Future Improvements: Tracked Change Editing

**Created:** 2025-12-19
**Updated:** 2025-12-22
**Status:** Most work complete; one optional enhancement remains
**Related:** `docs/internal/issues/ISSUE_CHAINED_EDITS_ON_TRACKED_CHANGES.md`

---

## What Was Implemented (2025-12-19)

Fixed the core issue: `replace_tracked()` and `delete_tracked()` now work on text inside `<w:ins>` elements.

**Key changes in `src/python_docx_redline/operations/tracked_changes.py`:**
- Added helper functions: `_is_tracked_change_wrapper()`, `_is_insertion_wrapper()`, `_is_deletion_wrapper()`, `_get_wrapper_author()`, `_is_same_author()`, `_get_run_parent_info()`
- Updated `_replace_run_with_elements()` to use `getparent()`
- Updated `_replace_match_with_element()` for parent detection
- Updated `_replace_match_with_elements()` for parent detection

**Tests:** `tests/test_chained_tracked_edits.py` (7 tests)

---

## What Was Implemented (2025-12-22)

### 1. Spanning Matches - IMPLEMENTED ✅

**Problem (now fixed):** Matches that span from regular text into `<w:ins>` (or vice versa) now work correctly.

**Example:**
```
Document: "The quick [ins:red] fox"
Attempted: replace_tracked("quick red", "slow blue")
Result: Works! " fox" remains in w:ins with original author attribution
```

**Implementation:**
- Added `_clone_wrapper()` helper to copy wrapper attributes with new ID
- Added `_extract_remaining_content_from_wrapper()` for wrapper splitting
- Added `_replace_multirun_match_with_elements()` for complex multi-parent cases
- Added `_find_paragraph_insertion_index()` for proper insertion positioning
- When before_text or after_text comes from a run inside a wrapper, the text is wrapped in a cloned wrapper preserving original author attribution

**New tests:** `tests/test_chained_tracked_edits.py` now has 13 tests including:
- `TestSpanningMatches::test_match_spanning_into_insertion`
- `TestSpanningMatches::test_match_spanning_out_of_insertion`
- `TestSpanningMatches::test_match_spanning_deletion_and_insertion`

---

## Remaining Optional Improvement

### Perfect `<w:ins>` Splitting (Low Priority)

**Status:** Optional polish - current implementation is functionally correct

**Current behavior (single-run partial edit inside w:ins):**
```xml
<!-- After replacing "brown" with "red" inside an insertion: -->
<w:ins author="A">
  <w:r><w:t>the quick </w:t></w:r>
  <w:del author="B"><w:r><w:delText>brown</w:delText></w:r></w:del>
  <w:ins author="B"><w:r><w:t>red</w:t></w:r></w:ins>
  <w:r><w:t> fox</w:t></w:r>
</w:ins>
```

**Word's behavior (splits wrapper):**
```xml
<w:ins author="A"><w:r><w:t>the quick </w:t></w:r></w:ins>
<w:del author="B"><w:r><w:delText>brown</w:delText></w:r></w:del>
<w:ins author="B"><w:r><w:t>red</w:t></w:r></w:ins>
<w:ins author="A"><w:r><w:t> fox</w:t></w:r></w:ins>
```

**Why it's optional:**
- Current behavior is functionally correct - attribution is preserved
- Nested structure is valid OOXML and renders correctly in Word
- The difference is cosmetic/structural, not functional

**If needed later:** Would require modifying `_split_and_replace_in_run_multiple()` to:
1. Detect if run is inside a wrapper
2. Remove the run from the wrapper
3. Create cloned wrappers for before/after text
4. Insert all elements at paragraph level instead of inside wrapper

**Estimate:** 1-2 hours

---

## Not Implementing

### Editing Inside `<w:del>`

**Recommendation:** Skip this. It's semantically questionable - if text is deleted, editing it makes no sense. Users should accept/reject the deletion first.

---

## Test Files for Reference

Word test files used for empirical testing are in `~/Downloads/`:
- `phase1.docx`, `phase2.docx`, `phase3.docx` - Author change scenarios
- `base.docx`, `T1_starting_point.docx`, `T4_into_ins.docx`, `T5_out_of_ins.docx`, `T10_del_and_ins.docx` - Edge cases

These show exactly how Word handles each scenario. Extract with:
```bash
unzip -d extracted phase1.docx && cat extracted/word/document.xml | xmllint --format -
```

---

## Key Code Locations

| Function | File | Purpose |
|----------|------|---------|
| `_replace_match_with_element` | tracked_changes.py | Single-element replacement |
| `_replace_match_with_elements` | tracked_changes.py | Multi-element replacement (del+ins) |
| `_replace_multirun_match_with_elements` | tracked_changes.py | Multi-run replacement with wrapper handling |
| `_replace_run_with_elements` | tracked_changes.py | Low-level run replacement |
| `_clone_wrapper` | tracked_changes.py | Clone w:ins/w:del with new ID |
| `_is_tracked_change_wrapper` | tracked_changes.py | Detect w:ins/w:del |
| `_find_paragraph_insertion_index` | tracked_changes.py | Find insertion point at paragraph level |
