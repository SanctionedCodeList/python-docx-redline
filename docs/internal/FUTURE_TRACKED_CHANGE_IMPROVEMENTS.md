# Future Improvements: Tracked Change Editing

**Created:** 2025-12-19
**Status:** Documentation for future work
**Related:** `docs/internal/issues/ISSUE_CHAINED_EDITS_ON_TRACKED_CHANGES.md`

---

## What Was Implemented (2025-12-19)

Fixed the core issue: `replace_tracked()` and `delete_tracked()` now work on text inside `<w:ins>` elements.

**Key changes in `src/python_docx_redline/operations/tracked_changes.py`:**
- Added helper functions (lines 911-1004): `_is_tracked_change_wrapper()`, `_is_insertion_wrapper()`, `_is_deletion_wrapper()`, `_get_wrapper_author()`, `_is_same_author()`, `_get_run_parent_info()`
- Updated `_replace_run_with_elements()` (line 1028) to use `getparent()`
- Updated `_replace_match_with_element()` (line 741) for parent detection
- Updated `_replace_match_with_elements()` (line 821) for parent detection

**Tests:** `tests/test_chained_tracked_edits.py` (7 tests)

---

## Remaining Limitations

### 1. Spanning Matches (Medium-High Complexity)

**Problem:** Matches that span from regular text into `<w:ins>` (or vice versa) don't work correctly.

**Example:**
```
Document: "The quick [ins:red] fox"
Attempted: replace_tracked("quick red", "slow blue")
Result: May fail or produce incorrect structure
```

**Technical Challenge:**
- Current code assumes all runs in a match share the same parent
- When runs have different parents (paragraph vs wrapper), removal/insertion logic breaks
- Need to coordinate replacement across parent boundaries

**Implementation Approach:**
```python
# In _replace_match_with_elements(), around line 867:
# Instead of assuming all runs share actual_parent:

for i in range(match.start_run_index, match.end_run_index + 1):
    run = match.runs[i]
    run_parent = run.getparent()

    # Track which wrappers we're exiting/entering
    if self._is_insertion_wrapper(run_parent):
        # Need to "close" this wrapper for remaining content
        # Insert replacement at paragraph level
        pass
```

**Key insight from Word testing:** Word handles this by splitting wrappers. See T4/T5 tests in the issue doc.

**Estimate:** 2-4 hours

---

### 2. Perfect `<w:ins>` Splitting (Medium Complexity)

**Problem:** When partially editing inside `<w:ins>`, we don't preserve the wrapper structure perfectly.

**Current behavior:**
```xml
<!-- Before: -->
<w:ins author="A"><w:r><w:t>the quick brown fox</w:t></w:r></w:ins>

<!-- After replacing "brown" with "red": -->
<!-- We produce something functional but not identical to Word -->
```

**Word's behavior:**
```xml
<w:ins author="A"><w:r><w:t>the quick </w:t></w:r></w:ins>
<w:del author="A"><w:r><w:delText>brown</w:delText></w:r></w:del>
<w:ins author="A"><w:r><w:t>red</w:t></w:r></w:ins>
<w:ins author="A"><w:r><w:t> fox</w:t></w:r></w:ins>
```

**Implementation Approach:**
1. When splitting a run inside `<w:ins>`, also split the wrapper
2. Clone wrapper attributes (author, date, id → need new id)
3. Use `_document._xml_generator.next_change_id` for new IDs

**Estimate:** 1-2 hours

---

### 3. Editing Inside `<w:del>` (Low Complexity, Low Value)

**Problem:** Cannot edit text that's already marked for deletion.

**Recommendation:** Skip this. It's semantically questionable - if text is deleted, editing it makes no sense. Users should accept/reject the deletion first.

**If needed later:** Add explicit error message:
```python
if self._is_deletion_wrapper(parent):
    raise ValueError("Cannot edit deleted text - accept/reject the deletion first")
```

---

## Recommended Priority

1. **Skip #3** - Not worth implementing
2. **Do #2 first if needed** - Lower complexity, cleaner output
3. **Do #1 only if real users hit it** - Complex, wait for concrete use case

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

| Function | File | Line | Purpose |
|----------|------|------|---------|
| `_replace_match_with_element` | tracked_changes.py | 741 | Single-element replacement |
| `_replace_match_with_elements` | tracked_changes.py | 821 | Multi-element replacement (del+ins) |
| `_replace_run_with_elements` | tracked_changes.py | 1028 | Low-level run replacement |
| `_is_tracked_change_wrapper` | tracked_changes.py | 911 | Detect w:ins/w:del |
| `_get_run_parent_info` | tracked_changes.py | 976 | Get paragraph + immediate parent |
