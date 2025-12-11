# Issue: insert_tracked Returns None When Text Spans Multiple Runs

**STATUS: RESOLVED** - This issue was resolved by the quote normalization feature added in commit `f82b5e4`.

## Summary

`insert_tracked()` returns `None` (indicating text not found) even when the anchor text exists in the document and quote normalization successfully matches it at the document level. The text is found via `get_text()` but not at the paragraph/run level where the insertion needs to happen.

## Reproduction

```python
from python_docx_redline import Document, AuthorIdentity

doc = Document("document.docx", author=identity)

after_text = '"the game-related statistics are matters of public interest." 109 N.E.3d 390, 397-99 (Ind. 2018).'

# This finds the text at document level
text = doc.get_text()
from python_docx_redline.quote_normalization import normalize_quotes
assert normalize_quotes(after_text) in normalize_quotes(text)  # True!

# But insert_tracked returns None
result = doc.insert_tracked(new_text, after=after_text)
print(result)  # None - text not found at paragraph level
```

## Observed Behavior

- `get_text()` returns the full document text and the anchor text is present
- Quote normalization correctly matches straight quotes to curly quotes
- But `insert_tracked()` returns `None` indicating the text wasn't found

## Expected Behavior

If the text exists in the document (as confirmed by `get_text()`), `insert_tracked()` should find it and perform the insertion.

## Root Cause (Confirmed)

The anchor text spans multiple XML runs within a paragraph. Concrete example from a legal brief:

The citation `109 N.E.3d 390, 397-99 (Ind. 2018).` is split across **4 runs**:
```
Run 49: ' 1'
Run 50: '09 N.E.3d 390, 397'
Run 51: '-'
Run 52: '99 (Ind. 2018).'
```

Word fragments text across runs for various reasons (formatting changes, editing history, spell-check boundaries). The `insert_tracked()` method appears to search within individual runs rather than across the concatenated paragraph text.

The `Paragraph.text` property correctly concatenates all runs, so `get_text()` works. But the insertion logic doesn't use the same cross-run matching.

## Suggested Fix

The text search logic should:
1. Concatenate all runs within a paragraph for matching
2. When a match is found spanning multiple runs, identify the correct insertion point (end of last run for `after=`, start of first run for `before=`)

## Environment

- Source document: Legal brief with citations containing mixed formatting
- python_docx_redline version: current development
- Date: 2024-12-08

## Resolution

This issue was resolved by:

1. The `TextSearch.find_text()` algorithm already correctly handles text spanning multiple runs by:
   - Building a character map that tracks which run each character belongs to
   - Concatenating all text from all runs within a paragraph
   - Searching in the concatenated text
   - Mapping results back to the original runs

2. The quote normalization feature (commit `f82b5e4`) ensures that straight quotes in search queries match smart/curly quotes in documents.

Tests added to verify this works:
- `test_insert_tracked_multi_run_text` - Tests text spanning 4+ runs
- `test_insert_tracked_multi_run_with_smart_quotes` - Tests combined multi-run + smart quote scenarios
