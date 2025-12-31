# Silent replacement in footnotes doesn't work with scope parameter

## Summary

When using `doc.replace()` with `scope="footnote:N"` and `track=False`, the text is not found even though `find_in_footnotes()` successfully locates it. This makes it impossible to silently remove placeholder text from footnotes after inserting hyperlinks.

## Use Case

Inserting clickable hyperlinks into footnotes using a placeholder-based workflow:

1. Insert footnote with placeholder: `"See XLINK1X for details"`
2. Insert hyperlink before placeholder: `insert_hyperlink_in_footnote(before="XLINK1X")`
3. Remove placeholder silently (without tracked changes)

Step 3 fails because there's no working method for silent text removal in footnotes.

## Current Behavior

```python
from python_docx_redline import Document
from docx import Document as DocxDoc

# Setup
d = DocxDoc()
d.add_paragraph('Test paragraph')
d.save('/tmp/test.docx')

doc = Document('/tmp/test.docx')
doc.insert_footnote('See XLINK1X for details', at='Test')
doc.insert_hyperlink_in_footnote(
    note_id=1,
    url='https://example.com',
    text='https://example.com',
    before='XLINK1X'
)

# This finds the text successfully
matches = doc.find_in_footnotes('XLINK1X')
print(f'Found: {len(matches)} matches')  # Output: Found: 1 matches

# But this fails to find the same text
doc.replace('XLINK1X', '', scope='footnote:1', track=False)
# Raises: TextNotFoundError: Could not find 'XLINK1X'
```

## Workarounds Attempted

| Method | Result |
|--------|--------|
| `doc.replace(..., scope="footnote:1", track=False)` | TextNotFoundError (doesn't find text) |
| `doc.replace(..., scope="footnotes", track=False)` | TextNotFoundError |
| `doc.delete_tracked_in_footnote(1, 'XLINK1X')` | Works but creates visible tracked changes |
| `footnote.replace_tracked('XLINK1X', '')` | Works but creates visible tracked changes |

## Expected Behavior

Either:

1. **`doc.replace()` with `scope="footnote:N"`** should find text that `find_in_footnotes()` can find
2. **`footnote.replace(find, replace)`** method (non-tracked) should exist on the Footnote model
3. **`doc.replace_in_footnote(note_id, find, replace, track=False)`** convenience method

## Suggested Solution

Add a `replace()` method to the Footnote model that performs silent (non-tracked) replacement:

```python
# Desired API
footnote = doc.get_footnote(1)
footnote.replace('XLINK1X', '')  # Silent replacement, no tracked changes
```

Or fix the scope-based replacement to use the same search logic as `find_in_footnotes()`.

## Environment

- python-docx-redline version: 0.1.0
- Python version: 3.12
- macOS

## Related

This may be related to how hyperlink insertion changes the XML structure (splits text runs), making the placeholder text span multiple elements that the scoped replace doesn't handle correctly.
