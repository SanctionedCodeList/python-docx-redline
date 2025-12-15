# Issue: Curly Quotes in Search Strings Break Python Syntax

**STATUS: RESOLVED** - The `enable_quote_normalization=True` parameter (enabled by default) allows users to write straight quotes in Python and match curly quotes in documents.

**Date:** 2025-12-15
**Severity:** High
**Category:** Developer Experience / Ergonomics
**Related:** `ISSUE_SMART_QUOTES_SEARCH.md` (solves matching, not authoring)

---

## Resolution

This issue is **already solved** by the existing quote normalization feature:

```python
# User types straight quotes (keyboard-friendly):
doc.replace_tracked("Defendant's property", "their property")

# Document contains curly quotes (Word-formatted):
# "Defendant's property" with U+2019 curly apostrophe

# Result: Match succeeds! Both are normalized to straight quotes internally.
```

The `enable_quote_normalization=True` parameter (default) normalizes BOTH the search text AND document text to straight quotes before comparison. This means:

1. **Users type naturally** with straight quotes from their keyboard
2. **Library matches any quote variant** in the document (straight, curly left, curly right)
3. **No special escaping needed** - just use regular Python strings

See `tests/test_quote_normalization.py` for comprehensive tests confirming this works.

---

## Original Summary (Historical)

When documents contain curly apostrophes (`'` U+2019), users cannot write Python code to search for that text because the curly apostrophe is visually identical to Python's string delimiter and causes **syntax errors before the code even runs**.

This is distinct from ISSUE_SMART_QUOTES_SEARCH, which addresses the *matching* problem (straight vs curly quotes). This issue addresses the *authoring* problem: you cannot write valid Python code containing curly quotes.

---

## The Problem

### Visual Collision

The curly apostrophe `'` (U+2019) looks nearly identical to Python's string delimiter `'` (U+0027):

```python
# What the user sees in their editor:
doc.replace_tracked('Plaintiffs' names', 'their names')
#                            ^ Looks fine!

# What Python sees:
doc.replace_tracked('Plaintiffs'  # String ends here
                    names',       # SyntaxError: invalid syntax
                    'their names')
```

### Real-World Failure

From actual usage attempting to edit a legal document:

```
  File "redline_status_report.py", line 77
    '2.Whether the identities of other putative class members were "held out" to advertise subscriptions to Defendants' property search platform.',
                                                                                                                                                 ^
SyntaxError: unterminated string literal (detected at line 77)
```

The string contains `Defendants'` with a curly apostrophe (U+2019), which Python interprets as ending the string.

---

## Why Existing Workarounds Fail

### Workaround 1: Double Quotes

```python
doc.replace_tracked("Defendants' property", "their property")
```

**Problem:** Still fails if text contains curly double quotes `"` or `"`.

### Workaround 2: Triple Quotes

```python
doc.replace_tracked('''Defendants' property''', '''their property''')
```

**Problem:** Same issue - curly quotes still interpreted as delimiters.

### Workaround 3: Raw Strings

```python
doc.replace_tracked(r'Defendants' property', r'their property')
```

**Problem:** Raw strings don't change how quotes are parsed.

### Workaround 4: Unicode Escapes

```python
doc.replace_tracked("Defendants\u2019 property", "their property")
```

**Works but:**
- Extremely tedious for long strings
- Easy to make mistakes
- Code becomes unreadable
- Must look up Unicode code points
- Every curly quote must be manually escaped

### Workaround 5: YAML/JSON External Files

```yaml
# edits.yaml
edits:
  - find: "Defendants' property"
    replace: "their property"
```

**Problem:** YAML and JSON have their own quoting rules that conflict with curly quotes. Same problem manifests differently.

### Workaround 6: Copy-Paste from Document

Copy the exact text from Word including curly quotes.

**Problems:**
- Invisible difference between straight/curly in most editors
- Git diffs show identical-looking strings as different
- Editor auto-formatting may "fix" quotes
- Team collaboration issues

---

## Proposed Solutions

### Solution A: `apply_edits_file()` Method (Recommended)

Provide a method that reads search/replace pairs from a file format designed for this:

```python
doc.apply_edits_file("edits.txt")
```

Where `edits.txt` uses a format that handles encoding naturally:

```
# edits.txt - UTF-8 encoded, no quote escaping needed
FIND: Defendants' property search platform.
REPLACE: their property search platform.
---
FIND: "held out" to advertise
REPLACE: "used" to promote
```

**Benefits:**
- Text editor handles UTF-8 encoding transparently
- No Python string literal issues
- Easy to review/edit
- Can be generated programmatically from document extraction

### Solution B: `find_and_replace()` with Extraction Helper

```python
# Extract exact text from document for matching
matches = doc.find_all("Defendants", context=50)
for match in matches:
    print(repr(match.text))  # Shows exact characters including Unicode

# Use the extracted text directly
doc.replace_tracked(
    matches[0].text,  # Already has correct encoding
    "replacement text"
)
```

### Solution C: Quote-Tolerant Matching (Enhancement to ISSUE_SMART_QUOTES_SEARCH)

Automatically normalize ALL quote variants during matching:

```python
doc.replace_tracked(
    "Defendants' property",  # User types straight apostrophe
    "their property",
    normalize_quotes=True    # Default: True
)
# Library internally tries: straight, curly, and all Unicode quote variants
```

This solves both issues:
1. Users type straight quotes naturally
2. Library matches any quote variant in document

### Solution D: Document Text Extraction for Code Generation

Provide a utility to generate valid Python code from document text:

```python
from python_docx_redline import generate_edit_code

# Reads document, outputs Python with proper escaping
code = generate_edit_code(
    doc_path="brief.docx",
    find_pattern="Defendants.*platform",  # Regex to locate text
    replace_with="replacement text"
)
print(code)
# Output:
# doc.replace_tracked("Defendants\u2019 property search platform.", "replacement text")
```

---

## Recommended Implementation Priority

1. **Solution C (normalize_quotes)** - Solves 90% of cases with zero user effort
2. **Solution A (apply_edits_file)** - For complex multi-edit workflows
3. **Solution B (extraction helper)** - For debugging and edge cases

---

## Test Cases

```python
def test_curly_apostrophe_in_possessive():
    """Most common case: possessives like Defendant's, Plaintiff's."""
    doc = create_doc_with_text("The Defendant's motion")  # Curly apostrophe

    # User should be able to type naturally with straight apostrophe
    result = doc.replace_tracked("Defendant's motion", "party's motion")
    assert result.success

def test_multiple_curly_quotes_in_string():
    """Multiple curly quotes in same search string."""
    doc = create_doc_with_text("Plaintiffs' and Defendants' arguments")
    result = doc.replace_tracked(
        "Plaintiffs' and Defendants' arguments",
        "the parties' arguments"
    )
    assert result.success

def test_mixed_single_and_double_curly_quotes():
    """Both single and double curly quotes in same text."""
    doc = create_doc_with_text('The "free trial" wasn\'t actually free')  # Smart quotes
    result = doc.replace_tracked(
        '"free trial" wasn\'t',  # User types straight quotes
        '"subscription" was not'
    )
    assert result.success

def test_edit_file_with_unicode():
    """apply_edits_file handles Unicode transparently."""
    # Create edit file with actual curly quotes (UTF-8 encoded)
    edit_content = """
FIND: Defendant's motion
REPLACE: party's motion
"""
    with open("test_edits.txt", "w", encoding="utf-8") as f:
        f.write(edit_content)

    doc = create_doc_with_text("The Defendant's motion")
    doc.apply_edits_file("test_edits.txt")
    assert "party's motion" in doc.get_text()
```

---

## Character Reference

| Character | Unicode | Name | Python Literal |
|-----------|---------|------|----------------|
| `'` | U+0027 | Apostrophe (straight) | `"'"` or `'\''` |
| `'` | U+2019 | Right Single Quote | `"\u2019"` |
| `'` | U+2018 | Left Single Quote | `"\u2018"` |
| `"` | U+0022 | Quotation Mark (straight) | `'"'` or `"\""` |
| `"` | U+201C | Left Double Quote | `"\u201c"` |
| `"` | U+201D | Right Double Quote | `"\u201d"` |

---

## Priority

**HIGH** - This is a fundamental usability blocker:
- Users literally cannot write valid Python code for common use cases
- Affects any document with possessives or contractions
- No simple workaround exists
- Makes the library appear broken ("I can see the text, why won't it match?")

---

## Related

- `ISSUE_SMART_QUOTES_SEARCH.md` - Addresses matching, not authoring
- The `normalize_quotes` solution proposed there would solve both issues if implemented

---

## Real-World Context

This issue was discovered while trying to programmatically edit a legal brief. Legal documents extensively use:
- Possessives: "Plaintiff's", "Defendant's", "Court's"
- Contractions: "doesn't", "isn't", "can't"
- Quoted legal standards: "held out", "for purposes of"

The user had to abandon programmatic editing and resort to manual Word editing because they could not write syntactically valid Python to match the document text.
