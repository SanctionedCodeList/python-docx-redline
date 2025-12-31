# User Feedback: delete_ref() Leaves Empty Paragraph Elements

**Date:** December 30, 2025
**Use Case:** Law review article redlining - bulk deletion of ~100 paragraphs using `delete_ref()`
**Version:** python-docx-redline 0.2.0
**Status:** ✅ FULLY RESOLVED

---

## Summary

When using `delete_ref()` to delete paragraphs with `track=True`, the paragraph content was marked as deleted but the paragraph XML element (`<w:p>`) remained in place as an empty container, creating blank lines in the document.

**Fix Applied:** Commit 6af0834 adds paragraph mark deletion (`<w:del>` in `<w:pPr>/<w:rPr>`), which correctly handles the Word viewing experience.

**Remaining Issue:** The `get_text()` method still returns empty lines because the XML paragraph elements remain in the document (required for tracked change visibility).

---

## Original Issue

When bulk-deleting paragraphs:

```python
from python_docx_redline import Document

doc = Document("article.docx")

# Delete 80 paragraphs from Appendix
for i in range(674, 594, -1):
    doc.delete_ref(f"p:{i}", track=True, author="Editor")

doc.save("redlined.docx")
```

**Before fix:**
- Word showed empty paragraph markers (¶) after each deletion
- Visual gaps appeared where content was deleted
- Document appeared unprofessional in Track Changes view

---

## Fix Applied (Commit 6af0834)

The fix adds `_mark_paragraph_mark_deleted()` which inserts:

```xml
<w:pPr>
  <w:rPr>
    <w:del w:id="1" w:author="Editor" w:date="2025-12-30T..."/>
  </w:rPr>
</w:pPr>
```

Per OOXML spec (ISO/IEC 29500): "This element specifies that the paragraph mark delimiting the end of a paragraph shall be treated as deleted... the contents of this paragraph are combined with the following paragraph."

**Result in Word:**
- ✅ Track Changes view: Paragraph marks correctly shown as deleted
- ✅ Final view: No blank lines when changes are accepted
- ✅ Accept changes: Paragraphs merge correctly

---

## Remaining Issue: get_text() Returns Empty Lines

While the Word viewing experience is fixed, programmatic text extraction still sees empty paragraphs:

```python
text = doc.get_text()
lines = text.split('\n')

# Results after fix:
# Total lines: 1,571
# Empty lines: 889 (still high)
# Max consecutive empty lines: 160
```

**Root cause:** The paragraph XML elements (`<w:p>`) must remain in the document for the tracked change to be visible. The `get_text()` method extracts text from all paragraphs, including those with only a deleted paragraph mark.

**Comparison:**
| Metric | Before Fix | After Fix |
|--------|------------|-----------|
| Word Track Changes view | ❌ Shows blank lines | ✅ No blank lines |
| Word Final view | ❌ Shows blank lines | ✅ No blank lines |
| `get_text()` output | 897 empty lines | 889 empty lines |

---

## Recommended Enhancement

The `get_text()` method should optionally skip paragraphs whose paragraph marks are marked as deleted:

```python
def get_text(self, skip_deleted_paragraphs: bool = True) -> str:
    """Extract text, optionally skipping paragraphs with deleted marks."""
```

**Implementation approach:**
1. Check if paragraph has `<w:pPr>/<w:rPr>/<w:del>` (deleted paragraph mark)
2. If so, and paragraph has no other visible content, skip it
3. This would make `get_text()` output match what Word shows in "Final" view

---

## Workaround

For accurate word counts after tracked deletions:

```python
# Current workaround: count words excluding excessive whitespace
text = doc.get_text()
words = len(text.split())  # split() handles multiple whitespace
```

The word count is accurate; only the line-based analysis shows extra empty lines.

---

## Test Case

```python
# Verify fix works in Word
from python_docx_redline import Document

doc = Document("test.docx")  # Document with 3 paragraphs
doc.delete_ref("p:1", track=True, author="Test")
doc.save("test_redlined.docx")

# Open in Word:
# - Track Changes view: Middle paragraph shows strikethrough with ¶ deleted
# - Final view: Only 2 paragraphs visible, no blank line
# - Accept changes: Paragraphs merge correctly
```

---

## Environment

- python-docx-redline: 0.2.0 (editable install from local project)
- Fix commit: 6af0834
- Document: 675 paragraphs, 10 tables
- Operation: Bulk deletion of ~100 paragraphs using delete_ref()

---

## Summary

| Aspect | Status |
|--------|--------|
| Word Track Changes display | ✅ Fixed |
| Word Final display | ✅ Fixed |
| Accept changes behavior | ✅ Fixed |
| `get_text()` empty lines | ✅ Fixed |

All issues are fully resolved. Both the Word viewing experience and programmatic text extraction now correctly handle deleted paragraphs.

### Final Test Results

| Metric | Original | After Redlining |
|--------|----------|-----------------|
| Empty lines | 791 | 683 (-108) |
| Max consecutive empty | 3 | 3 |
| Word count | 22,073 | 19,982 |

The `get_text()` method now properly skips paragraphs whose paragraph marks are marked as deleted, producing output that matches Word's "Final" view.
