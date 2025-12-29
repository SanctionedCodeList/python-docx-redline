# Bug: `delete_tracked()` creates partial deletions and orphaned text in table cells

## Summary

When using `delete_tracked()` on text within table cells, the deletion is incomplete:
1. Only part of the target text is deleted, leaving stub fragments
2. The "deleted" text appears as plain text at the end of the cell instead of being marked with strikethrough

## Reproduction

```python
from python_docx_redline import Document

doc = Document("claim_chart.docx")  # Document with table containing citations

# Table cell contains:
# • "Although only two dies 3a, 3b are shown in FIG. 1A, it should be appreciated
#    that more or fewer than two dies 3a, 3b may be mounted to the substrate 2."
#    U.S. Patent No. 10,204,893 at 19

doc.delete_tracked('Although only two dies 3a, 3b are shown in FIG. 1A, it should be appreciated')
doc.save("output.docx")
```

## Expected Result

The entire bullet point should be marked as deleted (strikethrough), or at minimum the specified text should be cleanly removed with strikethrough formatting.

## Actual Result

After save, the cell contains:

```
• " that more or fewer than two dies 3a, 3b may be mounted to the substrate 2." U.S. Patent No. 10,204,893 at 19
```

AND at the bottom of the cell, the deleted text appears as plain text (not strikethrough):

```
Although only two dies 3a, 3b are shown in FIG. 1A, it should be appreciated
```

## Impact

- Leaves orphaned citation stubs like `• "." U.S. Patent No. 9,184,125 at 1`
- Deleted text appears twice (once as stub, once as plain text at end)
- Document becomes corrupted/unusable for legal review

## Possible Causes

1. Text spanning multiple `<w:r>` (run) elements not being fully captured
2. Table cell XML structure causing issues with deletion scope
3. Tracked deletion markup being inserted incorrectly in table context

## Workaround

Use `delete_paragraph_tracked(containing="text")` instead to remove entire paragraphs cleanly:

```python
# Instead of:
doc.delete_tracked('Although only two dies 3a, 3b are shown in FIG. 1A...')

# Use:
doc.delete_paragraph_tracked(containing='Although only two dies 3a, 3b are shown')
```

This removes the entire paragraph element rather than trying to mark text as deleted inline.

## Priority

**High** - Makes bulk editing of claim charts unusable.

## Test Case

The issue occurs specifically when:
- Target document is a table (claim chart format)
- Deletion target is a citation with quotes and parenthetical references
- Deletion text is a substring of a larger bullet point
