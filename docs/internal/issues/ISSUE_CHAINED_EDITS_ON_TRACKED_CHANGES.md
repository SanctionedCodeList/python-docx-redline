# Issue: Cannot Edit Text That Was Already Modified by Tracked Changes

**Date:** 2025-12-19
**Severity:** Medium
**Category:** Feature Request / Usability
**Status:** 🔴 OPEN

---

## Summary

When making multiple rounds of tracked change edits to the same text, the second edit fails because the XML structure has changed after the first edit. Text wrapped in `<w:ins>` elements cannot be subsequently edited with `replace_tracked()`.

---

## Environment

- **python_docx_redline version:** 0.1.x
- **Python version:** 3.12
- **Use case:** Iterative document review with multiple rounds of edits

---

## Reproduction Steps

### 1. Make an initial tracked change

```python
from python_docx_redline import Document

doc = Document("document.docx", author="Reviewer")
doc.replace_tracked("suggests", "confirms", occurrence="all")
doc.save("document_REDLINE.docx")
```

### 2. Attempt to edit the same text again

```python
doc = Document("document_REDLINE.docx", author="Reviewer")
doc.replace_tracked("confirms", "indicates", occurrence="all")  # FAILS
doc.save("document_REDLINE.docx")
```

### 3. Error occurs

```
ValueError: <Element {http://schemas.openxmlformats.org/wordprocessingml/2006/main}r at 0x102e84f40> is not in list
```

---

## Root Cause Analysis

### What Happens

1. After `replace_tracked("suggests", "confirms")`, the text "confirms" is wrapped in a `<w:ins>` tracked change element
2. The run containing "confirms" is now a child of `<w:ins>`, not directly under `<w:p>`
3. When `replace_tracked("confirms", "indicates")` searches for "confirms", it finds the text
4. But when it tries to replace the run, it looks for the run as a direct child of the paragraph
5. The run is not found at that location → `ValueError`

### XML Before First Edit

```xml
<w:p>
  <w:r>
    <w:t>The evidence suggests that...</w:t>
  </w:r>
</w:p>
```

### XML After First Edit

```xml
<w:p>
  <w:r>
    <w:t>The evidence </w:t>
  </w:r>
  <w:del w:author="Reviewer">
    <w:r><w:delText>suggests</w:delText></w:r>
  </w:del>
  <w:ins w:author="Reviewer">
    <w:r><w:t>confirms</w:t></w:r>  <!-- Run is now inside <w:ins> -->
  </w:ins>
  <w:r>
    <w:t> that...</w:t>
  </w:r>
</w:p>
```

### Why Second Edit Fails

The replacement logic assumes runs are direct children of paragraphs, but after a tracked change, the run is nested inside `<w:ins>`.

---

## Proposed Solutions

### Option A: Chained Edits (Recommended)

Defer XML modifications until `save()`, applying all edits in a single pass:

```python
doc = Document("document.docx", author="Reviewer")
doc.replace_tracked("suggests", "confirms")  # Queued, not applied
doc.replace_tracked("confirms", "indicates")  # Queued, not applied
doc.save("output.docx")  # All edits applied in single pass
```

**Pros:**
- Clean API, no breaking changes
- Avoids XML structure issues entirely
- Better performance (single XML rewrite)

**Cons:**
- More complex implementation
- Need to resolve conflicts between queued edits

### Option B: Tracked Content Awareness

Extend `replace_tracked()` to handle text inside `<w:ins>` elements:

```python
# When searching for text, also look inside <w:ins> elements
# When replacing, handle the nested run structure correctly
```

**Pros:**
- Enables true iterative editing
- Works with saved/reopened documents

**Cons:**
- Complex XML manipulation
- Need to handle nested tracked changes (edits on edits)

### Option C: Re-index After Each Operation

Rebuild the internal text index after each tracked change:

```python
def replace_tracked(self, find, replace, **kwargs):
    # ... perform replacement ...
    self._rebuild_text_index()  # Re-scan document structure
```

**Pros:**
- Simpler than Option B
- Works with existing API

**Cons:**
- Performance overhead
- Still may fail on complex nested structures

### Option D: Accept-Then-Edit Mode

Provide option to accept pending changes before making new edits:

```python
doc = Document("document_REDLINE.docx", author="Reviewer")
doc.accept_all_changes()  # Flatten previous edits
doc.replace_tracked("confirms", "indicates")  # Now works
doc.save("output.docx")
```

**Pros:**
- Simple to implement
- Clear semantics

**Cons:**
- Loses tracked change history
- Not suitable for iterative review workflows

---

## Use Cases Affected

1. **Iterative review workflows** - Reviewer makes edits, gets feedback, adjusts edits
2. **Style consistency passes** - Multiple find/replace operations on same document
3. **Automated pipelines** - Scripts that apply multiple transformations

---

## Workaround

Currently, the only workaround is to start from the original document and apply all edits in a single session:

```python
doc = Document("original.docx", author="Reviewer")
# Apply all edits in one pass, using the final desired values
doc.replace_tracked("suggests", "indicates")  # Skip intermediate values
doc.save("output.docx")
```

---

## Recommendation

**Option A (Chained Edits)** is the recommended approach because:

1. It provides the cleanest API with no breaking changes
2. It avoids XML structure complexity entirely
3. It enables efficient batch processing
4. It naturally supports iterative workflows

Implementation would involve:
1. Add an internal edit queue
2. Modify `replace_tracked()`, `delete_tracked()`, etc. to queue operations
3. Apply all queued operations in `save()` in a single pass
4. Add optional `flush=True` parameter for immediate application if needed

---

## Empirical Testing: How Word Handles These Cases (2025-12-19)

We conducted hands-on testing with Microsoft Word to understand exactly how it handles editing text inside tracked changes.

### Test Setup

Base document: `The quick brown fox jumps over the lazy dog.`

### Test 1: Basic Tracked Replacement

**Action:** Replace "brown fox" → "red cat" (Track Changes ON)

**Result:**
```xml
<w:r><w:t>quick </w:t></w:r>
<w:del w:author="Parker Hancock" w:date="09:41">
  <w:r><w:delText>brown fox</w:delText></w:r>
</w:del>
<w:ins w:author="Parker Hancock" w:date="09:41">
  <w:r><w:t>red cat</w:t></w:r>
</w:ins>
<w:r><w:t> jumps</w:t></w:r>
```

### Test 2: Same Author Edits Own Insertion

**Action:** Same author replaces "indicates" → "confirms" (inside `<w:ins>`)

**Result:** Word **replaces content in place** - no nested deletion, just updates the text and timestamp inside the existing `<w:ins>`.

### Test 3: Different Author Edits Another's Insertion

**Action:** Different author replaces text inside `<w:ins>`

**Result:** Word **nests a `<w:del>` inside the `<w:ins>`** and creates a new `<w:ins>` at paragraph level:
```xml
<w:ins w:author="Author A">
  <w:del w:author="Author B">        <!-- NESTED deletion -->
    <w:r><w:delText>original</w:delText></w:r>
  </w:del>
</w:ins>
<w:ins w:author="Author B">          <!-- New insertion at paragraph level -->
  <w:r><w:t>replacement</w:t></w:r>
</w:ins>
```

### Test 4: Match Spans INTO Insertion

**Action:** Replace "quick red" → "slow gray" (spans regular text → `<w:ins>`)

**Result:**
```xml
<w:del>"quick "</w:del>           <!-- Regular text deleted -->
<w:del>"brown fox"</w:del>        <!-- Original deletion preserved -->
<w:ins>"slow gray"</w:ins>        <!-- New insertion -->
<w:ins>" cat"</w:ins>             <!-- Original ins SPLIT and truncated -->
```

Word **splits the `<w:ins>`**, keeping the unmatched portion (" cat") with its original timestamp.

### Test 5: Match Spans OUT OF Insertion

**Action:** Replace "cat jumps" → "bird flies" (spans `<w:ins>` → regular text)

**Result:**
```xml
<w:del>brown fox</w:del>          <!-- Preserved -->
<w:ins>red </w:ins>               <!-- TRUNCATED from "red cat" -->
<w:ins>bird flies</w:ins>         <!-- New insertion -->
<w:del> jumps</w:del>             <!-- Regular text deleted -->
```

Word **truncates the `<w:ins>`** and creates new del/ins for the portions outside.

### Test 6: Match Spans `<w:del>` AND `<w:ins>`

**Action:** Replace "brown fox red cat" → "blue bird" (spans both tracked change elements)

**Result:**
```xml
<w:del>brown fox</w:del>          <!-- Preserved with original timestamp -->
<w:ins>blue bird</w:ins>          <!-- Replaced "red cat" entirely -->
```

Word **preserves the deletion** and **replaces the insertion in place** (same author).

### Summary of Word's Behavior

| Scenario | Same Author | Different Author |
|----------|-------------|------------------|
| Edit fully inside `<w:ins>` | Update content in place | Nest `<w:del>` inside, new `<w:ins>` outside |
| Match spans regular → `<w:ins>` | Split `<w:ins>`, delete regular, new `<w:ins>` | Same + nesting |
| Match spans `<w:ins>` → regular | Truncate `<w:ins>`, delete regular, new `<w:ins>` | Same + nesting |
| Match spans `<w:del>` + `<w:ins>` | Preserve `<w:del>`, update `<w:ins>` in place | Same + nesting |

### Key Insights

1. **Same author = in-place modification**: Word optimizes by updating existing tracked changes rather than creating nested structures
2. **Different author = nested history**: Word preserves full attribution by nesting `<w:del>` inside `<w:ins>`
3. **Splitting behavior**: Word intelligently splits `<w:ins>` elements when matches partially overlap
4. **Deletion preservation**: Existing `<w:del>` elements are always preserved

---

## Chosen Solution: Option B with "Good Enough" Approach

After empirical testing, we're implementing **Option B (Tracked Content Awareness)** with a simplified approach that handles the common cases without matching Word's exact splitting behavior.

### Implementation Scope

**Will Handle:**
1. Text fully inside `<w:ins>` - same author updates in place
2. Text fully inside `<w:ins>` - different author nests `<w:del>` and creates new `<w:ins>`
3. Basic parent detection using `getparent()` instead of assuming paragraph

**Won't Handle (Documented Limitations):**
1. Matches that span across tracked change boundaries (regular ↔ `<w:ins>`)
2. Perfect splitting of `<w:ins>` elements
3. Matches inside `<w:del>` elements (deleted text shouldn't be editable anyway)

### Technical Approach

1. **Add `_get_run_container()` helper** - uses `getparent()` to find actual parent
2. **Add `_is_tracked_change_wrapper()` helper** - detects `<w:ins>` and `<w:del>`
3. **Update replacement logic:**
   - If run is in `<w:ins>` (same author): update content in place
   - If run is in `<w:ins>` (different author): nest `<w:del>`, create new `<w:ins>` at paragraph level
   - If run is in `<w:del>`: skip (can't meaningfully edit deleted text)
   - If run is direct child of paragraph: current behavior (works fine)

---

## References

- Error encountered: `operations/tracked_changes.py:941` in `_replace_run_with_elements()`
- Related to XML structure of tracked changes (ISO/IEC 29500)
- User report: TSMC claim chart editing workflow (2025-12-19)
- Empirical testing: Word for macOS, December 2025

---

## Priority

**MEDIUM** - Workaround exists but creates friction in iterative review workflows. Would significantly improve usability for document review use cases.
