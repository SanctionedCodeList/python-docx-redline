# User Feedback: Validation Error - Nested `<w:del>` Elements from Double-Deletion

**Date:** December 30, 2025
**Use Case:** Law review article redlining with bulk deletions
**Version:** python-docx-redline 0.2.0 (local development version)
**Status:** ✅ RESOLVED

---

## Summary

When `delete_ref()` is called on a paragraph that has already been deleted (with tracked changes), it creates nested `<w:del>` elements which are invalid OOXML. This causes validation to fail on `save()`.

---

## Root Cause

**Two issues combine to create this bug:**

1. **Failed saves still persist changes to disk** - When `save()` fails validation, the XML has already been written to the underlying ZIP archive. The document on disk contains the deletion markers even though the save "failed."

2. **`delete_ref()` doesn't detect already-deleted content** - When called on a paragraph that's already wrapped in `<w:del>` elements, it wraps the existing `<w:del>` inside another `<w:del>`, creating invalid nesting.

---

## Invalid XML Structure

### Before Double-Deletion (Valid)
```xml
<w:p>
  <w:pPr>
    <w:rPr>
      <w:del w:id="4" w:author="Parker Hancock" w:date="2025-12-30T19:27:54Z"/>
    </w:rPr>
  </w:pPr>
  <w:del w:id="1" w:author="..." w:date="...">
    <w:r>
      <w:delText>The critical distinction is the intent requirement...</w:delText>
    </w:r>
  </w:del>
</w:p>
```

### After Double-Deletion (INVALID)
```xml
<w:p>
  <w:pPr>
    <w:rPr>
      <!-- INVALID: Duplicate <w:del> elements -->
      <w:del w:id="4" w:author="Parker Hancock" w:date="2025-12-30T19:27:54Z"/>
      <w:del w:id="4" w:author="Parker Hancock" w:date="2025-12-30T19:45:28Z"/>
    </w:rPr>
  </w:pPr>
  <!-- INVALID: Nested <w:del> inside <w:del> -->
  <w:del w:id="1" w:author="..." w:date="2025-12-30T19:27:54Z">
    <w:del w:id="1" w:author="..." w:date="2025-12-30T19:45:28Z">
      <w:r>
        <w:delText>The critical distinction is the intent requirement...</w:delText>
      </w:r>
    </w:del>
  </w:del>
</w:p>
```

**Validation Error:**
```
Element '{...wordprocessingml/2006/main}del': This element is not expected.
Expected is one of ( {...}moveFrom, ...
```

---

## Minimal Reproduction

```python
from python_docx_redline import Document

# Step 1: Create a document with an already-deleted paragraph
doc = Document("document_with_prior_deletions.docx")

# Step 2: Check if paragraph is already deleted
text = doc.get_text_at_ref("p:270")
print(f"Text at p:270: '{text}'")  # Empty string = already deleted

# Step 3: Try to delete the already-deleted paragraph
doc.delete_ref("p:270", track=True, author="Editor")

# Step 4: Save fails with validation error
doc.save("output.docx")  # ValidationError: nested <w:del> elements
```

### Creating Test Document

To create a document in this state:

```python
from python_docx_redline import Document

# Create fresh document
doc = Document("test.docx")

# Delete a paragraph (this will succeed)
doc.delete_ref("p:5", track=True, author="Editor")
doc.save("test.docx")  # Success

# Reload and try to delete the same paragraph again
doc = Document("test.docx")
doc.delete_ref("p:5", track=True, author="Editor")  # Creates nested <w:del>
doc.save("test.docx")  # ValidationError
```

---

## Verification Script

```python
from python_docx_redline import Document
from lxml import etree

def check_paragraph_deletion_state(doc, ref):
    """Check if a paragraph is already marked as deleted."""
    root = doc.xml_root
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    body = root.find('.//w:body', nsmap)
    paragraphs = body.findall('w:p', nsmap)

    # Parse ref to get index
    idx = int(ref.split(':')[1])
    if idx >= len(paragraphs):
        return None

    p = paragraphs[idx]

    # Check for deleted paragraph mark
    ppr = p.find('w:pPr', nsmap)
    has_del_mark = False
    if ppr is not None:
        rpr = ppr.find('w:rPr', nsmap)
        if rpr is not None:
            has_del_mark = rpr.find('w:del', nsmap) is not None

    # Check for deleted content
    dels = p.findall('.//w:del', nsmap)

    return {
        'ref': ref,
        'has_deleted_paragraph_mark': has_del_mark,
        'deletion_element_count': len(dels),
        'text_content': doc.get_text_at_ref(ref),
        'is_already_deleted': has_del_mark and len(dels) > 0
    }

# Usage
doc = Document("document.docx")
state = check_paragraph_deletion_state(doc, "p:270")
print(state)
# {'ref': 'p:270', 'has_deleted_paragraph_mark': True,
#  'deletion_element_count': 4, 'text_content': '',
#  'is_already_deleted': True}
```

---

## Recommended Fixes

### Fix 1: Check for Already-Deleted Content (Primary)

In `delete_ref()`, check if the paragraph is already deleted before attempting deletion:

```python
def delete_ref(self, ref: str, track: bool = False, author: str = None):
    if track:
        # Check if already deleted
        if self._is_ref_already_deleted(ref):
            # Option A: Skip silently
            return EditResult(success=True, message="Already deleted")

            # Option B: Raise informative error
            raise AlreadyDeletedError(f"{ref} is already marked as deleted")

            # Option C: Log warning and skip
            logger.warning(f"{ref} is already deleted, skipping")
            return EditResult(success=True, message="Skipped - already deleted")

    # ... proceed with deletion
```

### Fix 2: Atomic Save with Rollback (Secondary)

Ensure that if validation fails, the document on disk is not modified:

```python
def save(self, path: str):
    # Write to temp file first
    temp_path = path + ".tmp"
    self._write_to_file(temp_path)

    # Validate the temp file
    try:
        self._validate_file(temp_path)
    except ValidationError:
        os.remove(temp_path)  # Clean up
        raise

    # Only replace original if validation passes
    os.replace(temp_path, path)
```

### Fix 3: Add `_is_ref_already_deleted()` Helper

```python
def _is_ref_already_deleted(self, ref: str) -> bool:
    """Check if a ref points to already-deleted content."""
    # Get text - empty string suggests deletion
    text = self.get_text_at_ref(ref)
    if text:
        return False

    # Verify by checking XML structure
    element = self.resolve_ref(ref)
    if element is None:
        return True

    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # Check for deleted paragraph mark
    ppr = element.find('w:pPr', nsmap)
    if ppr is not None:
        rpr = ppr.find('w:rPr', nsmap)
        if rpr is not None and rpr.find('w:del', nsmap) is not None:
            return True

    # Check if all content is in <w:del> elements
    # (no visible text outside of deletions)
    return True  # Refined check needed
```

---

## Impact

**Severity:** Medium

- **Workaround exists:** Don't delete the same paragraph twice
- **Detection possible:** Check `get_text_at_ref()` before deleting - empty string indicates already deleted
- **Data loss risk:** Low - document remains readable, just can't add more deletions

**Affected Use Cases:**
- Bulk redlining workflows where users might reload and continue editing
- Scripts that don't track which paragraphs have already been deleted
- Error recovery scenarios where a failed save is followed by retry

---

## Environment

- python-docx-redline: 0.2.0 (editable install)
- Python: 3.11+
- Test document: Law review article, 677 paragraphs, ~150 already-deleted paragraphs

---

## Test Document

The affected document is available at:
`content/deployers-dilemma/drafts/deployers-dilemma-redlined.docx`

This document has p:270 in the "already deleted" state, making it easy to reproduce the validation error.

---

## Related Issues

- `2025-12-30-delete-ref-leaves-empty-paragraphs.md` - Previous issue about paragraph mark deletion (RESOLVED)
