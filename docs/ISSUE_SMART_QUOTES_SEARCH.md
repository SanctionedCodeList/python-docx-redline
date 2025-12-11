# Issue: Text Search Fails with Smart Quotes / Curly Apostrophes Unless Exact Match

**Date:** 2025-12-08
**Severity:** Medium-High
**Category:** Text Search / User Experience

---

## Summary

Word documents commonly use smart quotes (curly quotes/apostrophes: `'` `'` `"` `"`) instead of straight quotes (`'` `"`). When users search for text using straight quotes in their Python strings, the search fails even though the text is visible and readable. This creates a significant usability issue requiring users to manually copy exact characters or use character codes.

---

## Environment

- **python_docx_redline version:** 0.1.0 (editable install)
- **Python version:** 3.12
- **Document source:** Microsoft Word documents (which auto-convert straight quotes to smart quotes)

---

## Reproduction Steps

### 1. Create Document in Word

Word automatically converts:
- Straight apostrophe `'` → Right single quotation mark `'` (U+2019)
- Straight double quote `"` → Left/right double quotation marks `"` `"` (U+201C, U+201D)

Example text in Word: `"Don't use the plaintiff's name for advertising"`

### 2. Attempt Search with Python

```python
from python_docx_redline import Document

doc = Document("brief.docx")

# Text as user types it naturally
doc.replace_tracked(
    "plaintiff's name",
    "party's identity"
)
# ERROR: TextNotFoundError: Could not find "plaintiff's name"
```

### 3. Error Message

```
python_docx_redline.errors.TextNotFoundError: Could not find "plaintiff's name"

Suggestions:
  • Document contains curly apostrophes (''). Try replacing straight apostrophes with curly ones in search text
```

---

## Root Cause

Word uses Unicode smart punctuation:
- `'` (U+0027) → `'` (U+2019) - Right single quotation mark
- `"` (U+0022) → `"` (U+201C) - Left double quotation mark
- `"` (U+0022) → `"` (U+201D) - Right double quotation mark

Python string literals use straight quotes by default, so there's a character mismatch.

---

## Current Workarounds

### Workaround 1: Use Character Codes

```python
# Awkward but works
doc.replace_tracked(
    "plaintiff\u2019s name",  # \u2019 = '
    "party\u2019s identity"
)
```

**Problems:**
- Unintuitive and hard to read
- Easy to make mistakes
- Requires looking up Unicode codes

### Workaround 2: Copy-Paste from Document

```python
# Copy text directly from Word with smart quotes
doc.replace_tracked(
    "plaintiff's name",  # Pasted from Word
    "party's identity"
)
```

**Problems:**
- Requires round-tripping to Word
- Invisible difference between smart/straight quotes in code
- Difficult in version control (looks identical)
- Team members with different editor configs may normalize quotes

### Workaround 3: Extract Text and Check

```python
# Debug to find exact characters
text = doc.get_text()
idx = text.find("plaintiff")
print(repr(text[idx:idx+20]))  # Shows actual characters
# Output: "plaintiff\u2019s name"
```

**Problems:**
- Time-consuming debugging step
- Manual investigation required for every mismatch

---

## Expected Behavior

The library should handle quote normalization automatically or provide an easy option:

### Option A: Automatic Quote Normalization (Recommended)

```python
# User types naturally with straight quotes
doc.replace_tracked(
    "plaintiff's name",
    "party's identity"
)
# Library automatically normalizes: ' → ' for matching
```

### Option B: Explicit Flag

```python
doc.replace_tracked(
    "plaintiff's name",
    "party's identity",
    normalize_quotes=True  # default: True
)
```

### Option C: Helper Function

```python
from python_docx_redline import Document, normalize_quotes

doc.replace_tracked(
    normalize_quotes("plaintiff's name"),
    normalize_quotes("party's identity")
)
```

**Recommendation:** Option A (automatic) is best for user experience. Power users can opt-out if needed.

---

## Impact

### High Frequency

Smart quotes appear in:
- **Possessives:** "plaintiff's claim", "defendant's conduct"
- **Contractions:** "don't", "can't", "it's"
- **Quoted text:** "the 'free trial' argument", dialogue, legal citations
- **Nearly every Word document:** Word's default AutoFormat setting

This isn't an edge case—it's the default behavior of Microsoft Word.

### Poor User Experience

```python
# User sees this in Word:
"The plaintiff's theory can't succeed."

# User types this naturally:
doc.replace_tracked("plaintiff's theory can't succeed", "new text")

# Library says: ❌ Text not found

# User confusion: "But I can SEE it right there!"
```

This creates a "works on paper, fails in code" problem that damages library credibility.

### Inconsistent Error Messages

The error message sometimes helps:
```
Suggestions:
  • Document contains curly apostrophes (''). Try replacing straight apostrophes with curly ones
```

But:
1. Only appears when smart quotes are detected
2. Doesn't help with double quotes `"` vs `"`/`"`
3. Suggests manual fix rather than library handling it

---

## Proposed Solution

### 1. Add Quote Normalization Function

```python
def _normalize_quotes(text: str) -> str:
    """Normalize quotes for flexible matching.

    Converts common quote variations to their canonical forms:
    - Straight quotes (', ") → smart quotes (', ", ")
    - Or vice versa depending on document convention

    Args:
        text: Text to normalize

    Returns:
        Text with normalized quotes
    """
    # Smart quote mappings
    replacements = {
        "'": "'",   # Straight apostrophe → right single quote
        '"': '"',   # Straight quote → left double quote (or right, context-dependent)
    }

    result = text
    for straight, smart in replacements.items():
        result = result.replace(straight, smart)

    return result
```

### 2. Apply in Search Operations

```python
def replace_tracked(
    self,
    find: str,
    replace: str,
    regex: bool = False,
    normalize_quotes: bool = True,
    **kwargs
) -> ReplaceResult:
    """Replace text with tracked changes.

    Args:
        find: Text to find
        replace: Replacement text
        regex: Whether find is a regex pattern
        normalize_quotes: Normalize quote characters for matching (default: True)
        **kwargs: Additional options (scope, etc.)
    """
    if normalize_quotes and not regex:
        # Normalize both search text and document text
        find_normalized = self._normalize_quotes(find)

        # Try normalized search first
        try:
            return self._do_replace(find_normalized, replace, **kwargs)
        except TextNotFoundError:
            # Fall back to exact match
            return self._do_replace(find, replace, **kwargs)

    return self._do_replace(find, replace, **kwargs)
```

### 3. Provide Escape Hatch

Users who need exact matching can disable:

```python
doc.replace_tracked(
    "plaintiff's name",      # Must match exactly
    "party's identity",
    normalize_quotes=False   # Strict matching
)
```

---

## Alternative: Bidirectional Normalization

Instead of only converting straight → smart, support both directions:

```python
def _fuzzy_quote_match(text: str, document_text: str) -> bool:
    """Check if text matches with flexible quote handling."""
    # Try multiple normalizations
    variants = [
        text,                           # Original
        _normalize_to_smart(text),      # Straight → smart
        _normalize_to_straight(text),   # Smart → straight
    ]

    for variant in variants:
        if variant in document_text:
            return True

    return False
```

This handles cases where:
- User types straight quotes, doc has smart quotes (most common)
- User types smart quotes, doc has straight quotes (rare but possible)
- Mixed quote styles in same search

---

## Character Mappings

### Apostrophes / Single Quotes

| Character | Unicode | Name | Usage |
|-----------|---------|------|-------|
| `'` | U+0027 | Apostrophe (straight) | Keyboard default, code |
| `'` | U+2019 | Right Single Quotation Mark | Word default for apostrophes |
| `'` | U+2018 | Left Single Quotation Mark | Opening single quote |
| `` ` `` | U+0060 | Grave Accent | Sometimes misused |

### Double Quotes

| Character | Unicode | Name | Usage |
|-----------|---------|------|-------|
| `"` | U+0022 | Quotation Mark (straight) | Keyboard default, code |
| `"` | U+201C | Left Double Quotation Mark | Opening quote in Word |
| `"` | U+201D | Right Double Quotation Mark | Closing quote in Word |

---

## Test Cases Needed

```python
def test_search_with_straight_quotes_finds_smart_quotes():
    """Straight quotes in search text match smart quotes in document."""
    doc = create_doc_with_text("The plaintiff's claim can't succeed.")  # Smart quotes
    result = doc.replace_tracked("plaintiff's claim can't", "party's claim cannot")
    assert result.success

def test_search_with_smart_quotes_finds_smart_quotes():
    """Smart quotes in search text match smart quotes in document."""
    doc = create_doc_with_text("The plaintiff's claim")  # Smart quote
    result = doc.replace_tracked("plaintiff's claim", "party's claim")  # Smart quote in search
    assert result.success

def test_search_double_quotes():
    """Double quote normalization works."""
    doc = create_doc_with_text('The "free trial" argument')  # Smart double quotes
    result = doc.replace_tracked('"free trial" argument', 'subscription argument')  # Straight quotes
    assert result.success

def test_normalize_quotes_can_be_disabled():
    """Exact matching works when normalize_quotes=False."""
    doc = create_doc_with_text("plaintiff's claim")  # Smart quote
    with pytest.raises(TextNotFoundError):
        doc.replace_tracked("plaintiff's claim", "text", normalize_quotes=False)  # Straight quote fails

def test_mixed_quotes_in_same_search():
    """Handles mixed quote styles in single search."""
    doc = create_doc_with_text("Don't use "free trial" here")  # Smart single and double
    result = doc.replace_tracked('Don\'t use "free trial"', 'Avoid "trial"')  # Straight quotes
    assert result.success

def test_possessives_and_contractions():
    """Common possessives and contractions work."""
    test_cases = [
        ("plaintiff's", "party's"),
        ("don't", "do not"),
        ("can't", "cannot"),
        ("it's", "it is"),
    ]
    for find, replace in test_cases:
        doc = create_doc_with_text(f"The {find} argument")
        result = doc.replace_tracked(find, replace)
        assert result.success
```

---

## Documentation Impact

### README.md

Add section on quote handling:

```markdown
### Working with Smart Quotes

Word documents typically use smart quotes (curly quotes/apostrophes: `'` `"`) instead of
straight quotes (`'` `"`). python_docx_redline handles this automatically:

\`\`\`python
# Your document contains: "The plaintiff's claim"
# You can search naturally with straight quotes:
doc.replace_tracked("plaintiff's claim", "party's claim")
# ✓ Works! Library normalizes quotes automatically

# Disable if you need exact character matching:
doc.replace_tracked("plaintiff's claim", "party's claim", normalize_quotes=False)
\`\`\`

**Supported quote types:**
- Single quotes: `'` ↔ `'` (apostrophes, possessives, contractions)
- Double quotes: `"` ↔ `"` `"` (quoted text, dialogue)
```

### QUICK_REFERENCE.md

Add to "Common Gotchas" section.

---

## Priority

**MEDIUM-HIGH** - Very common issue that:
- Affects most Word documents (Word's default behavior)
- Creates frustrating user experience
- Requires workarounds that feel like bugs
- Can be solved with straightforward normalization
- Should be fixed before 1.0 to avoid breaking changes

---

## Related Issues

- **ISSUE_TEXT_SEARCH_WITH_WHITESPACE.md** - Similar text normalization problem (RESOLVED)
- Could extend to other Unicode normalizations (em dashes `—` vs hyphens `-`, etc.)

---

## References

- Error encountered in: Edit 3, Edit 8 of surgical edits workflow
- Unicode references:
  - [U+2019 Right Single Quotation Mark](https://unicode-table.com/en/2019/)
  - [U+201C Left Double Quotation Mark](https://unicode-table.com/en/201C/)
  - [U+201D Right Double Quotation Mark](https://unicode-table.com/en/201D/)
- Word AutoFormat: Automatically converts quotes unless disabled
