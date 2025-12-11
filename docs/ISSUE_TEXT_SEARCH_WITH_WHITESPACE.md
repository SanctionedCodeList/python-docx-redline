# Issue: Text Search Fails When Content is Split Across Runs with Whitespace

**Date:** 2025-12-08
**Severity:** High
**Category:** Text Search / Core Functionality
**Status:** ✅ RESOLVED (2025-12-08)

---

## Summary

The `replace_tracked()` and `insert_tracked()` methods fail to find text that exists in the document when that text is split across multiple `<w:r>` (run) elements with intervening whitespace/formatting. This makes the library unusable for real-world documents where text is commonly fragmented across runs.

---

## Environment

- **python_docx_redline version:** 0.1.0 (editable install)
- **Python version:** 3.12
- **Document source:** Microsoft Word document with tracked changes that were accepted
- **Operations attempted:** `replace_tracked()`, `insert_tracked()`

---

## Reproduction Steps

### 1. Source Document

Starting with `2025-12-02_MTD_client_comments.docx`:
- Contains tracked changes and comments
- Accepted all changes: `doc.accept_all_changes()`
- Deleted all comments: `doc.delete_all_comments()`
- Saved as `MTD_client_accepted.docx`

### 2. Verification with Pandoc

Text verified to exist using pandoc conversion:

```bash
pandoc --track-changes=all MTD_client_accepted.docx -o MTD_client_accepted.md
```

Output shows (line 224):
```markdown
It claims merely that a database records their property ownership.
```

Text clearly reads as continuous prose.

### 3. Attempted Search with python_docx_redline

```python
from python_docx_redline import Document

doc = Document("MTD_client_accepted.docx")

# Attempt 1: Full phrase
doc.replace_tracked(
    "records their property ownership",
    "compiles their property ownership data"
)
# Error: TextNotFoundError: Could not find 'records their property ownership'

# Attempt 2: Shorter phrase
doc.replace_tracked(
    "database records their property ownership",
    "database compiles their property ownership data"
)
# Error: TextNotFoundError: Could not find 'database records their property ownership'

# Attempt 3: Even shorter
doc.replace_tracked(
    "database records",
    "database compiles"
)
# Error: TextNotFoundError: Could not find 'database records'
```

### 4. Debug Investigation

Created debug script to inspect `get_text()` output:

```python
from python_docx_redline import Document

doc = Document("MTD_client_accepted.docx")
full_text = doc.get_text()

# Search for fragments
search_terms = [
    "database records",
    "records their property",
    "their property ownership",
]

for term in search_terms:
    if term in full_text:
        print(f"✓ Found: '{term}'")
        idx = full_text.find(term)
        print(f"  Context: ...{full_text[max(0, idx-50):idx+len(term)+50]}...")
    else:
        print(f"✗ NOT found: '{term}'")
```

**Output:**
```
✗ NOT found: 'database records'
✗ NOT found: 'records their property'
✓ Found: 'their property ownership'
  Context: ...

        records


         their property ownership


        . Th


        ...
```

**Key finding:** The text "records" and "their property ownership" have multiple newlines/whitespace between them in the `get_text()` output, even though they appear as continuous text when the document is rendered.

---

## Root Cause Analysis

### Current Behavior

The `get_text()` method appears to preserve whitespace exactly as it appears in the XML structure, including:
- Newlines between `<w:r>` elements
- Whitespace from formatting/structure
- Run boundaries

This causes continuous prose that is split across multiple runs to become unfindable via simple string search.

### Expected Behavior

For text search purposes, the library should:

1. **Option A: Normalize whitespace in search operations**
   - Collapse multiple whitespace characters into single spaces
   - Allow searches to match across run boundaries
   - Similar to how browsers/renderers display the text

2. **Option B: Provide alternative search method**
   - Add a `find_text()` method that normalizes whitespace
   - Keep `get_text()` as-is for debugging/inspection
   - Let users choose which behavior they need

3. **Option C: Make whitespace handling configurable**
   - Add parameter to search methods: `normalize_whitespace=True`
   - Default to True for user-facing operations
   - False for exact/debugging operations

---

## Impact

### High Severity Because:

1. **Real-world documents are fragmented**
   - Word splits text across runs for formatting, tracked changes, etc.
   - This is not an edge case—it's the norm
   - Any document that has been edited will have run fragmentation

2. **Library becomes unusable for its primary use case**
   - The main purpose is to apply tracked changes to existing documents
   - If text cannot be found reliably, the library cannot function
   - Users cannot work around this without manual XML manipulation

3. **No workaround available**
   - Regex with `\s+` doesn't work because the whitespace in `get_text()` doesn't match what regex expects
   - Cannot guess the exact whitespace pattern
   - Would need to unpack and manually edit XML

---

## Related Code Paths

The issue likely affects these methods:
- `replace_tracked()` - line 360 in document.py (raises TextNotFoundError)
- `insert_tracked()` - line 229 in document.py (raises TextNotFoundError)
- `delete_tracked()` - presumably same issue
- Any method that uses `get_text()` for searching

---

## Proposed Solutions

### Short-term (Quick Fix)

Add a `_normalize_whitespace()` helper method and use it in search operations:

```python
def _normalize_whitespace(self, text: str) -> str:
    """Normalize whitespace for text search."""
    import re
    return re.sub(r'\s+', ' ', text).strip()

def replace_tracked(self, find: str, replace: str, regex: bool = False, **kwargs):
    """Replace text with tracked changes."""
    if not regex:
        # Normalize both search text and document text for comparison
        find_normalized = self._normalize_whitespace(find)
        doc_text_normalized = self._normalize_whitespace(self.get_text())

        if find_normalized not in doc_text_normalized:
            raise TextNotFoundError(find)
    # ... rest of implementation
```

### Long-term (Better Design)

1. **Separate text extraction from text search**
   - `get_text()` → returns text as-is for debugging
   - `get_searchable_text()` → returns normalized text for search
   - Search methods use `get_searchable_text()` internally

2. **Add configurable whitespace handling**
   ```python
   doc.replace_tracked(
       find="database records",
       replace="database compiles",
       normalize_whitespace=True  # default
   )
   ```

3. **Provide detailed error messages**
   When text is not found, show:
   - Closest matches with edit distance
   - Actual text fragments found
   - Whitespace representation (show `\n`, `\t`, etc.)

---

## Test Cases Needed

```python
def test_search_across_runs_with_whitespace():
    """Text split across runs with whitespace should still be findable."""
    # Create document where "hello world" is split:
    # <w:r><w:t>hello</w:t></w:r>
    # <w:r><w:t> </w:t></w:r>
    # <w:r><w:t>world</w:t></w:r>

    doc = Document(...)
    result = doc.replace_tracked("hello world", "goodbye world")
    assert result.success

def test_search_with_newlines_between_runs():
    """Text with newlines between runs should normalize."""
    # Document has: "hello\n\nworld" in get_text()
    # But displays as: "hello world"

    doc = Document(...)
    result = doc.replace_tracked("hello world", "goodbye world")
    assert result.success

def test_search_exact_vs_normalized():
    """Users can opt into exact matching if needed."""
    doc = Document(...)

    # Should fail - exact match required
    with pytest.raises(TextNotFoundError):
        doc.replace_tracked("hello world", "goodbye", normalize_whitespace=False)

    # Should succeed - normalized match
    result = doc.replace_tracked("hello world", "goodbye", normalize_whitespace=True)
    assert result.success
```

---

## Workaround (Until Fixed)

Current workaround is to use the lower-level OOXML editing approach:

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path.home() / ".agents/skills/anthropic_skills/document-skills/docx"))

from scripts.document import Document
from scripts.find_text import find_text

# Unpack, manually locate text in XML, replace at XML level, repack
# This defeats the purpose of python_docx_redline as a high-level API
```

---

## References

- Debug script: `debug_text.py`
- Source document: `2025-12-02_MTD_client_comments.docx`
- Intermediate file: `MTD_client_accepted.docx`
- Failed script: `accept_and_apply_edits.py`

---

## Priority

**HIGH** - This is a blocking issue that prevents the library from being used for its primary purpose on real-world documents. Should be addressed before any 0.2.0 release.

---

## Resolution

**Date Resolved:** 2025-12-08
**Fixed in Commit:** (see git log)

### Root Cause

The issue was caused by both `Paragraph.text` and `TextSearch.find_text()` using `itertext()` which returns **all text nodes** in the XML tree, including whitespace text nodes that exist between XML elements for formatting purposes.

When a document has indented/formatted XML like:
```xml
<w:p>
  <w:r>
    <w:t>database</w:t>
  </w:r>
  <w:r>
    <w:t>records</w:t>
  </w:r>
</w:p>
```

The whitespace (newlines and indentation) between `</w:r>` and `<w:r>` was being included in the extracted text, making continuous prose unfindable.

### Solution Implemented

Following **Eric White's algorithm insight** (docs/ERIC_WHITE_ALGORITHM.md), we modified text extraction to only extract text from `<w:t>` elements, ignoring XML structural whitespace:

1. **Created helper function** (`_get_run_text()` in text_search.py):
   ```python
   def _get_run_text(run: Any) -> str:
       """Extract text from a run, avoiding XML structural whitespace."""
       text_elements = run.findall(f".//{{{WORD_NAMESPACE}}}t")
       return "".join(elem.text or "" for elem in text_elements)
   ```

2. **Fixed Paragraph.text property** (src/python_docx_redline/models/paragraph.py:39-51):
   ```python
   @property
   def text(self) -> str:
       """Get all text content from the paragraph."""
       text_elements = self._element.findall(f".//{{{WORD_NAMESPACE}}}t")
       return "".join(elem.text or "" for elem in text_elements)
   ```

3. **Updated TextSearch.find_text()** to use `_get_run_text()` helper
4. **Updated TextSpan.text and TextSpan.context** to use the same approach

### Files Changed

- `src/python_docx_redline/models/paragraph.py` - Fixed Paragraph.text property
- `src/python_docx_redline/text_search.py` - Added helper function and updated all text extraction
- `tests/test_text_extraction_whitespace.py` - Added 13 comprehensive tests

### Test Coverage

Created **13 new tests** covering:
- Text extraction with formatted XML
- Text extraction with multiple newlines between runs
- Search across formatted runs
- Multi-word phrases across runs
- Insert, delete, and replace operations with formatted XML
- Edge cases (empty runs, single-character runs, mixed content)

**All 213 tests pass** with 93% coverage (up from 92%).

### Validation

The fix successfully resolves the exact issue from the bug report:
- ✅ `"database records"` is now findable
- ✅ `"records their property ownership"` is now findable
- ✅ `get_text()` returns clean continuous text
- ✅ All search operations work across formatted runs
- ✅ No regressions in existing functionality

### Eric White's Algorithm

We adopted the **core insight** from Eric White's algorithm (Step 1): "Concatenate all text in a paragraph into a single string by extracting only from `<w:t>` elements."

We did NOT need the full algorithm (single-character runs, coalescing, etc.) because our character map approach was already efficient for read-only search operations. This hybrid approach gives us:
- ✅ Correct text extraction (from Eric White)
- ✅ Efficient search (from our character map)
- ✅ No document modification overhead
