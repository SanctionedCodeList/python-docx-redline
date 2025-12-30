# User Feedback: delete_ref() Leaves Empty Paragraph Elements

**Date:** December 30, 2025
**Use Case:** Law review article redlining - bulk deletion of ~100 paragraphs using `delete_ref()`
**Version:** python-docx-redline (current as of Dec 30, 2025)
**Severity:** Medium - creates visual artifacts that require manual cleanup

---

## Summary

When using `delete_ref()` to delete paragraphs with `track=True`, the paragraph content is marked as deleted but the paragraph XML element (`<w:p>`) remains in place as an empty container. This results in blank lines appearing in the document wherever content was deleted.

---

## Reproduction

```python
from python_docx_redline import Document

doc = Document("article.docx")

# Delete 80 paragraphs from Appendix
for i in range(674, 594, -1):  # Delete from end to avoid ref shifting
    doc.delete_ref(f"p:{i}", track=True, author="Editor")

doc.save("redlined.docx")
```

**Expected behavior:** Document shows struck-through text for deleted content, no blank lines.

**Actual behavior:** Document shows struck-through text, PLUS blank lines for each deleted paragraph.

---

## Analysis

### Word Document Structure

When viewing the redlined document in Word:
- Deleted text appears with strikethrough (correct)
- Empty paragraph markers (¶) appear after each deletion (incorrect)
- These create visual "gaps" that weren't in the original document

### Text Extraction Evidence

```python
text = doc.get_text()  # Excludes deleted content
lines = text.split('\n')

# Results:
# Total lines: 1,571
# Empty lines: 897  (57% of all lines!)
# Max consecutive empty lines: 176
```

### Specific Examples

**Template artifacts deletion (23 paragraphs):**
- Location: Between Section IV and Section V
- Result: 37 consecutive empty lines in extracted text
- Visual impact: Large gap in Word document

**Appendix deletion (80 paragraphs):**
- Location: End of document
- Result: 80+ empty paragraph markers
- Visual impact: Multiple pages of empty strikethrough paragraphs

---

## Root Cause

Looking at the YAML output from AccessibilityTree:

```yaml
# After deleting "LIMITATION OF LIABILITY" paragraph:
- paragraph [ref=p:609]:
    text: ""          # <-- Empty paragraph container remains
- paragraph [ref=p:610]:
    text: "LIMITATION OF LIABILITY"
    has_changes: true
    changes:
      - type: deletion
        text: "LIMITATION OF LIABILITY"
```

The `delete_ref()` operation:
1. ✓ Creates tracked deletion markup for the text content
2. ✗ Leaves the `<w:p>` element in place with no content
3. ✗ Does not mark the paragraph element itself as deleted

---

## Expected Behavior Options

### Option A: Remove Empty Paragraphs (when track=False)
When not tracking changes, if `delete_ref()` removes all content from a paragraph, remove the paragraph element entirely.

```python
doc.delete_ref("p:15", track=False)  # Paragraph disappears completely
```

### Option B: Delete Entire Paragraph as Tracked Change (when track=True)
When tracking changes, wrap the entire paragraph (including its structure) in deletion markup so it appears as a single struck-through block.

```python
doc.delete_ref("p:15", track=True)  # Shows as strikethrough paragraph, no blank line
```

### Option C: Add `remove_empty` Parameter
Add an optional parameter to control behavior:

```python
# Current behavior (for backwards compatibility)
doc.delete_ref("p:15", track=True, remove_empty=False)

# New behavior: remove paragraph if it becomes empty
doc.delete_ref("p:15", track=True, remove_empty=True)
```

---

## Workaround

Currently, users must manually clean up empty paragraphs after deletion:

1. Open document in Word
2. Accept all tracked changes
3. Use Find & Replace to remove multiple paragraph marks
4. Re-apply any needed formatting

This defeats the purpose of programmatic redlining.

---

## Impact

- **Bulk deletions create massive gaps** - Deleting an appendix or section leaves dozens of blank lines
- **Document appears unprofessional** - Empty paragraph markers visible in Track Changes view
- **Manual cleanup required** - Negates efficiency gains from programmatic editing
- **Word count tools affected** - Some tools count empty paragraphs

---

## Related Issues

This may be related to the earlier feedback about paragraph-level operations (2025-12-29-law-review-redlining-session.md). The fundamental issue is that the library treats paragraph deletion as "delete the text inside the paragraph" rather than "delete the paragraph itself."

---

## Recommendation

For tracked changes, the ideal behavior would be:

```xml
<!-- Current (leaves empty paragraph): -->
<w:p>
  <w:del w:author="Editor">
    <w:r><w:t>Deleted text</w:t></w:r>
  </w:del>
</w:p>
<w:p/>  <!-- Empty paragraph remains -->

<!-- Desired (paragraph itself is deleted): -->
<w:del w:author="Editor">
  <w:p>
    <w:r><w:t>Deleted text</w:t></w:r>
  </w:p>
</w:del>
<!-- No empty paragraph -->
```

The `<w:del>` element can wrap entire paragraphs in OOXML, not just runs within paragraphs. This would show the paragraph as deleted without leaving an empty container.

---

## Environment

- python-docx-redline: (current as of Dec 30, 2025)
- Document: 675 paragraphs, 10 tables
- Operation: Bulk deletion of ~100 paragraphs using delete_ref()
- Result: 897 empty lines in extracted text (57% of total)
