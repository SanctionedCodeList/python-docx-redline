# Design: Fix delete_ref() Leaving Empty Paragraph Elements

**Date:** 2025-12-30
**Status:** Proposed
**Related Feedback:** `/docs/user_feedback/2025-12-30-delete-ref-leaves-empty-paragraphs.md`
**Similar Resolved Issue:** `/docs/internal/issues/ISSUE_DELETE_PARAGRAPH_LEAVES_EMPTY_LINE.md`

---

## Summary

When using `delete_ref()` with `track=True` to delete paragraphs, the paragraph content is correctly marked as deleted (strikethrough), but the paragraph element (`<w:p>`) remains as an empty container. This creates blank lines in the document wherever content was deleted.

The root cause is that `delete_ref()` only wraps runs in `<w:del>` elements but does not mark the paragraph mark itself as deleted, unlike the already-fixed `delete_paragraph_tracked()` method which handles this correctly.

---

## Analysis of Current Implementation

### Location

The `delete_ref()` method is in `/src/python_docx_redline/document.py` starting at line 5291.

### Current Behavior for Tracked Paragraph Deletion

```python
def _delete_paragraph_ref(
    self,
    element: etree._Element,
    track: bool,
    author: str,
) -> EditResult:
    """Delete a paragraph element."""
    from datetime import datetime, timezone

    if track:
        # Tracked deletion: wrap all runs in w:del
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self._xml_generator.next_change_id

        for run in list(element.findall(f".//{{{WORD_NAMESPACE}}}r")):
            # Create deletion wrapper
            del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
            del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
            del_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
            del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

            # Convert w:t to w:delText
            for t_elem in run.findall(f".//{{{WORD_NAMESPACE}}}t"):
                t_elem.tag = f"{{{WORD_NAMESPACE}}}delText"

            # Move run into deletion
            run_parent = run.getparent()
            if run_parent is not None:
                run_index = list(run_parent).index(run)
                run_parent.remove(run)
                del_elem.append(run)
                run_parent.insert(run_index, del_elem)

            change_id += 1

        self._xml_generator.next_change_id = change_id
    # ...
```

### What's Missing

The current implementation:
1. Wraps each run in `<w:del>` elements
2. Converts `<w:t>` to `<w:delText>`
3. **Does NOT mark the paragraph mark as deleted**

### Expected XML Structure (Current)

```xml
<w:p>
  <w:del w:author="Editor" w:date="2025-12-30T10:00:00Z">
    <w:r>
      <w:delText>Deleted text content</w:delText>
    </w:r>
  </w:del>
</w:p>
<!-- When accepted, leaves: -->
<w:p/>  <!-- Empty paragraph = blank line -->
```

### Expected XML Structure (Fixed)

```xml
<w:p>
  <w:pPr>
    <w:rPr>
      <w:del w:id="1" w:author="Editor" w:date="2025-12-30T10:00:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:del w:author="Editor" w:date="2025-12-30T10:00:00Z">
    <w:r>
      <w:delText>Deleted text content</w:delText>
    </w:r>
  </w:del>
</w:p>
<!-- When accepted: paragraph merges with next, no blank line -->
```

---

## Working Reference Implementation

The `delete_paragraph_tracked()` method in `/src/python_docx_redline/operations/section.py` already implements the correct behavior. It was fixed in a previous issue (2025-12-19).

### Key Pattern from SectionOperations

```python
def delete_paragraph_tracked(self, ...):
    # ...
    for target_para in reversed(target_paras):
        # Mark text content as deleted (strikethrough)
        runs = list(target_para.element.iter(f"{{{WORD_NAMESPACE}}}r"))
        if runs:
            del_elem = self._create_deletion_element(author_name, timestamp)
            self._wrap_runs_in_deletion(target_para.element, runs, del_elem)

        # Mark paragraph mark as deleted (causes merge with next paragraph on accept)
        self._mark_paragraph_mark_deleted(target_para.element, author_name, timestamp)
```

### The _mark_paragraph_mark_deleted Method

```python
def _mark_paragraph_mark_deleted(self, para_element: Any, author: str, timestamp: str) -> None:
    """Mark the paragraph mark as deleted for tracked changes.

    Adds a <w:del> element inside <w:pPr>/<w:rPr> to mark the paragraph
    mark as deleted. When this tracked change is accepted in Word,
    the paragraph merges with the following paragraph instead of leaving
    an empty line behind.

    Per OOXML spec (ISO/IEC 29500): "This element specifies that the
    paragraph mark delimiting the end of a paragraph shall be treated
    as deleted... the contents of this paragraph are combined with the
    following paragraph."
    """
    # Get or create paragraph properties <w:pPr>
    p_pr = para_element.find(f"{{{WORD_NAMESPACE}}}pPr")
    if p_pr is None:
        p_pr = etree.Element(f"{{{WORD_NAMESPACE}}}pPr")
        para_element.insert(0, p_pr)

    # Get or create run properties for paragraph mark <w:rPr>
    r_pr = p_pr.find(f"{{{WORD_NAMESPACE}}}rPr")
    if r_pr is None:
        r_pr = etree.Element(f"{{{WORD_NAMESPACE}}}rPr")
        p_pr.append(r_pr)

    # Create the deletion marker for the paragraph mark
    change_id = self._document._xml_generator.next_change_id
    self._document._xml_generator.next_change_id += 1

    del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
    del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
    del_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
    del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

    r_pr.append(del_elem)
```

---

## Proposed Solution

### Option 1: Add Paragraph Mark Deletion to _delete_paragraph_ref (Recommended)

Modify `_delete_paragraph_ref()` in `document.py` to call a similar helper method that marks the paragraph mark as deleted.

#### Changes Required

1. **Add helper method to Document class** (or reuse from SectionOperations):

```python
def _mark_paragraph_mark_deleted(
    self, para_element: etree._Element, author: str, timestamp: str
) -> None:
    """Mark the paragraph mark as deleted for tracked changes."""
    # Get or create paragraph properties <w:pPr>
    p_pr = para_element.find(f"{{{WORD_NAMESPACE}}}pPr")
    if p_pr is None:
        p_pr = etree.Element(f"{{{WORD_NAMESPACE}}}pPr")
        para_element.insert(0, p_pr)

    # Get or create run properties for paragraph mark <w:rPr>
    r_pr = p_pr.find(f"{{{WORD_NAMESPACE}}}rPr")
    if r_pr is None:
        r_pr = etree.Element(f"{{{WORD_NAMESPACE}}}rPr")
        p_pr.append(r_pr)

    # Create the deletion marker for the paragraph mark
    change_id = self._xml_generator.next_change_id
    self._xml_generator.next_change_id += 1

    del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
    del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
    del_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
    del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

    r_pr.append(del_elem)
```

2. **Update _delete_paragraph_ref() to call the helper**:

```python
def _delete_paragraph_ref(
    self,
    element: etree._Element,
    track: bool,
    author: str,
) -> EditResult:
    """Delete a paragraph element."""
    from datetime import datetime, timezone

    if track:
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self._xml_generator.next_change_id

        # Wrap all runs in deletion markers
        for run in list(element.findall(f".//{{{WORD_NAMESPACE}}}r")):
            # ... existing run wrapping code ...

        self._xml_generator.next_change_id = change_id

        # NEW: Mark the paragraph mark itself as deleted
        self._mark_paragraph_mark_deleted(element, author, timestamp)
    else:
        # Hard delete: remove the paragraph element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)

    # Invalidate registry cache
    self._ref_registry.invalidate()

    return EditResult(
        success=True,
        edit_type="delete_ref",
        message="Deleted paragraph" + (" with tracking" if track else ""),
    )
```

### Option 2: Refactor to Share Code with SectionOperations

Move `_mark_paragraph_mark_deleted()` and related helpers to a shared utility module (e.g., `tracked_xml.py` or a new `paragraph_utils.py`) and call from both locations.

**Pros:**
- Eliminates code duplication
- Single source of truth for paragraph deletion logic

**Cons:**
- Larger refactoring effort
- May introduce coupling between modules

### Recommendation

**Option 1 is recommended** for immediate fix. The helper method can be added directly to the Document class. Later, if more operations need this functionality, a refactoring to share code can be done.

---

## Considerations for track=False Case

When `track=False`, the current implementation already handles this correctly:

```python
else:
    # Hard delete: remove the paragraph element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)
```

The paragraph is completely removed, leaving no empty container. No changes needed for this case.

---

## Backward Compatibility

### Impact Assessment

- **Behavioral change**: Tracked deletions will now merge paragraphs when accepted instead of leaving blank lines
- **XML structure change**: Additional `<w:del>` element added inside `<w:pPr>/<w:rPr>`
- **No API changes**: Method signature remains identical

### Risk Analysis

1. **Low risk**: The change aligns behavior with `delete_paragraph_tracked()` which already works this way
2. **Expected behavior**: Users explicitly request deletion; they expect the paragraph to disappear when accepted
3. **Reversible**: Users can reject the tracked change in Word if needed

### Migration Notes

- Documents created with old version will still open correctly
- Existing tracked deletions in documents won't change
- New deletions will have the improved behavior

---

## Test Cases Needed

### New Tests for _delete_paragraph_ref

```python
class TestDeleteRefParagraphMarkDeletion:
    """Tests for paragraph mark deletion in delete_ref()."""

    def test_tracked_delete_marks_paragraph_mark(self) -> None:
        """Verify that tracked deletion marks the paragraph mark as deleted."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)

            doc.delete_ref("p:0", track=True, author="TestAgent")

            element = doc.resolve_ref("p:0")
            # Check for paragraph mark deletion in w:pPr/w:rPr/w:del
            p_pr = element.find(f"{{{WORD_NAMESPACE}}}pPr")
            assert p_pr is not None
            r_pr = p_pr.find(f"{{{WORD_NAMESPACE}}}rPr")
            assert r_pr is not None
            del_elem = r_pr.find(f"{{{WORD_NAMESPACE}}}del")
            assert del_elem is not None
            assert del_elem.get(f"{{{WORD_NAMESPACE}}}author") == "TestAgent"
        finally:
            docx_path.unlink()

    def test_tracked_delete_no_empty_paragraph_after_accept(self) -> None:
        """Verify that accepting deletion doesn't leave empty paragraph."""
        # This test would require simulating Word's accept behavior
        # or visual inspection of the output document
        pass

    def test_untracked_delete_removes_paragraph_entirely(self) -> None:
        """Verify that untracked deletion removes the paragraph element."""
        docx_path = create_test_docx()
        try:
            doc = Document(docx_path)
            count_before = len(doc.paragraphs)

            doc.delete_ref("p:0", track=False)

            count_after = len(doc.paragraphs)
            assert count_after == count_before - 1
        finally:
            docx_path.unlink()
```

### Update Existing Tests

The existing test `test_delete_paragraph_tracked` should be enhanced to verify the paragraph mark deletion marker is present.

---

## Implementation Checklist

1. [ ] Add `_mark_paragraph_mark_deleted()` helper method to Document class
2. [ ] Modify `_delete_paragraph_ref()` to call the helper when `track=True`
3. [ ] Add test for paragraph mark deletion marker presence
4. [ ] Add test verifying untracked mode still removes paragraph entirely
5. [ ] Run full test suite to verify no regressions
6. [ ] Manual verification: create test document, delete paragraphs, accept in Word, verify no blank lines

---

## References

- **OOXML Spec (ISO/IEC 29500)**: Section on tracked revisions and paragraph marks
- **User Feedback**: `/docs/user_feedback/2025-12-30-delete-ref-leaves-empty-paragraphs.md`
- **Previous Fix**: `/docs/internal/issues/ISSUE_DELETE_PARAGRAPH_LEAVES_EMPTY_LINE.md`
- **Working Implementation**: `SectionOperations.delete_paragraph_tracked()` in `/src/python_docx_redline/operations/section.py`
- **Current Implementation**: `Document._delete_paragraph_ref()` in `/src/python_docx_redline/document.py`

---

## Appendix: OOXML Background

From the OOXML specification (ISO/IEC 29500):

> "This element specifies that the paragraph mark delimiting the end of a paragraph shall be treated as deleted... the contents of this paragraph are combined with the following paragraph."

The `<w:del>` element inside `<w:pPr>/<w:rPr>` marks the paragraph mark (the invisible character at the end of each paragraph) as deleted. When Word accepts this change:

1. The paragraph mark is removed
2. The paragraph's content merges with the following paragraph
3. No empty paragraph remains

This is the correct OOXML behavior for tracked paragraph deletion.
