# User Feedback: Law Review Article Redlining Session

**Date:** December 29, 2025
**Use Case:** Revising a ~22,000 word law review article to target length of 15,000-18,000 words
**Tools Used:** python-docx-redline, AccessibilityTree
**Outcome:** Partial success with significant friction

---

## Summary

Attempted to programmatically redline a merged law review article, making tracked deletions of specific sections (Appendix, redundant regulatory discussion, methodology detail) while adding an abstract. The AccessibilityTree successfully identified document structure with refs, but the editing API couldn't leverage those refs for paragraph-level operations.

---

## What Worked Well

### 1. AccessibilityTree Structure Discovery
```python
from python_docx_redline.accessibility import AccessibilityTree, ViewMode

tree = AccessibilityTree.from_document(doc, view_mode=mode)
print(tree.to_yaml(verbosity="minimal"))
```

This provided excellent structured output:
```yaml
content:
  - h1 "I. INTRODUCTION" [ref=p:3]
  - h2 "A. Opening Hook" [ref=p:4]
  - p "The general counsel sits across from her team..." [ref=p:5]
```

The refs clearly identified every paragraph, making it easy to understand document structure.

### 2. Text-Based Find and Delete
```python
doc.delete_tracked("specific text to remove", author="Parker Hancock")
```

This worked correctly for deleting exact text strings with tracked changes.

### 3. Text-Based Insert
```python
doc.insert_tracked("new text", after="anchor text", author="Parker Hancock")
```

Successfully inserted abstract content after the author name.

### 4. find_all() for Discovery
```python
matches = doc.find_all("search phrase")
for m in matches:
    print(m.context)
```

Useful for locating text before deletion.

---

## What Didn't Work

### 1. No Paragraph-Level Deletion

**Expected (from skill documentation):**
```python
doc.delete_at_ref("p:15", track=True)  # Delete entire paragraph
```

**Reality:** This method doesn't exist. The only deletion method is:
```python
doc.delete_tracked("exact text string")  # Deletes ONLY these words
```

**Impact:** To delete a 200-word paragraph, I had to either:
- Find and pass the entire 200-word string (error-prone, quote handling issues)
- Delete the opening phrase, leaving orphaned text

### 2. Gap Between AccessibilityTree and Editing API

The AccessibilityTree provides stable refs (`p:5`, `p:100`, `tbl:0/row:2/cell:1`) but the Document editing API can't use them:

| AccessibilityTree Can Do | Document API Can Do |
|--------------------------|---------------------|
| Identify paragraph by ref | Find text by string |
| Show paragraph structure | Replace text strings |
| Navigate to specific elements | Insert at text anchors |
| **Cannot:** Edit by ref | **Cannot:** Edit by ref |

**What's Missing:**
```python
# These would bridge the gap:
doc.get_text_at_ref("p:15")  # Get full paragraph text
doc.delete_at_ref("p:15", track=True)  # Delete paragraph
doc.replace_at_ref("p:15", "old", "new", track=True)  # Edit in paragraph
```

### 3. No Section-Level Operations

For law review editing, I needed to delete entire sections (e.g., the Appendix with all subsections). There's no:
```python
doc.delete_section("Appendix", track=True)  # Delete heading + all content until next same-level heading
doc.delete(scope="section:Appendix", track=True)  # Alternative syntax
```

### 4. Skill Documentation Shows Non-Existent Methods

The `accessibility.md` skill file shows:
```python
# Edit by ref (unambiguous targeting)
doc.insert_at_ref("p:5", " (amended)", position="end")
doc.replace_at_ref("p:10", "old text", "new text")
doc.delete_at_ref("p:15", "remove this")
```

These methods don't exist on the Document object, causing confusion about available capabilities.

---

## Specific Friction Points

### 1. Fragmentary Deletions

**Intent:** Delete the entire paragraph starting with "The EU AI Act establishes a risk-based classification system."

**What I Did:**
```python
doc.delete_tracked("The EU AI Act establishes a risk-based classification system.")
```

**Result:** Only those 10 words were deleted. The rest of the paragraph (~150 words) remained, now orphaned and nonsensical.

### 2. Quote Handling in Long Text

When trying to delete full paragraphs by passing complete text:
```python
doc.delete_tracked('''The EU AI Act establishes a risk-based classification system. AI practices posing "unacceptable risk" are prohibited...''')
```

The curly quotes in the document didn't match straight quotes in my code, causing failures.

### 3. Footnote Orphaning

When deleting text containing footnote markers `[^15]`, the marker was deleted but the footnote content in footnotes.xml remained, creating orphaned footnotes.

**Expected:** Option to delete associated footnotes, or at least a warning.

### 4. Sequential Footnote Handling

Deleting text between two footnotes left them adjacent (e.g., `[^15][^16]`). These should be merged per Bluebook citation style, but there's no:
```python
doc.merge_footnotes([15, 16])  # Combine into single footnote
```

---

## Feature Requests

### Priority 1: Ref-Based Editing

Bridge the AccessibilityTree to the editing API:

```python
# Get text by ref
text = doc.get_text_at_ref("p:15")

# Delete by ref
doc.delete_at_ref("p:15", track=True, author="Editor")

# Replace within paragraph identified by ref
doc.replace_at_ref("p:15", "old", "new", track=True)

# Insert at ref position
doc.insert_at_ref("p:15", "new content", position="after", track=True)
```

### Priority 2: Section-Level Operations

```python
# Delete entire section (heading + all content until next same-level heading)
doc.delete_section("Appendix", track=True)
doc.delete_section(ref="p:595", track=True)  # By heading ref

# Move section
doc.move_section("Appendix", after="Conclusion", track=True)
```

### Priority 3: Footnote Handling

```python
# Delete footnote by ID
doc.delete_footnote(15)

# Merge adjacent footnotes
doc.merge_footnotes([15, 16], separator="; ")

# Check for orphaned footnotes after edits
orphans = doc.find_orphaned_footnotes()
```

### Priority 4: Full Paragraph Text Retrieval

```python
# When I have a partial match, get the full paragraph
matches = doc.find_all("The EU AI Act")
full_paragraph = matches[0].full_paragraph_text  # Currently only .context available
doc.delete_tracked(full_paragraph)
```

---

## Workarounds Used

### 1. Fragmentary Deletion (Suboptimal)

Deleted opening phrases, leaving orphaned text:
```python
doc.delete_tracked("The EU AI Act establishes", author="Editor")
# Leaves: "a risk-based classification system. AI practices..."
```

### 2. Multiple Passes

Deleted text in chunks:
```python
doc.delete_tracked("The EU AI Act establishes a risk-based classification system.", author="Editor")
doc.delete_tracked("AI practices posing unacceptable risk are prohibited outright", author="Editor")
# Still leaves fragments
```

### 3. Manual Cleanup Required

The programmatic redlining got ~10% of the way there. User must:
1. Open in Word
2. Accept tracked changes
3. Manually select and delete remaining orphaned content
4. Fix footnote issues

---

## Environment

- python-docx-redline version: (current as of Dec 29, 2025)
- Document: 22,000 words, 256 footnotes, 10 tables
- Task: Reduce to 15,000-18,000 words with tracked changes

---

## Recommendations

1. **Document the actual API** - The skill files show methods that don't exist. Either implement them or correct the documentation.

2. **Bridge AccessibilityTree to editing** - The tree provides excellent structure discovery; let it drive edits.

3. **Add paragraph-level operations** - Law review editing, contract redlining, and document restructuring all need to move/delete/replace entire paragraphs, not just text strings.

4. **Consider footnote lifecycle** - Legal documents live and die by footnotes. Deleting text with footnote refs should handle the footnotes too.

---

## Positive Notes

The library handles the hard problems well:
- Run fragmentation (text split across formatting runs)
- Tracked changes XML generation
- Smart quote normalization
- Finding text across complex document structures

The gap is between "find this text" and "delete this structural element."
