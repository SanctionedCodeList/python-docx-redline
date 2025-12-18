# Comments

## Adding Comments

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Basic comment
comment = doc.add_comment("Please review this", on="Section 2.1")

# With custom author
comment = doc.add_comment(
    "Needs legal review",
    on="liability clause",
    author="Legal Team",
    initials="LT"
)

doc.save("contract_with_comments.docx")
```

## Handling Multiple Occurrences

When target text appears multiple times, use the `occurrence` parameter:

```python
# Target specific occurrence (1-indexed)
doc.add_comment("Check first", on="Evidence Gaps:", occurrence=1)
doc.add_comment("Check third", on="Evidence Gaps:", occurrence=3)

# Target first or last
doc.add_comment("Opening", on="Section", occurrence="first")
doc.add_comment("Final", on="Section", occurrence="last")

# Add to ALL occurrences (returns list)
comments = doc.add_comment("Review all", on="TODO", occurrence="all")
print(f"Added {len(comments)} comments")

# Target specific indices
comments = doc.add_comment("Check these", on="Important:", occurrence=[1, 3, 5])
```

## Scoped Comments

Limit where comments are added:

```python
doc.add_comment(
    "Review payment terms",
    on="30 days",
    scope="section:Payment Terms"
)
```

**Note:** If text exists but outside your scope, the error message indicates this:
```
Could not find '30 days'

Note: Found 5 occurrence(s) in the document, but none within scope 'section:Payment Terms'.
Try removing or adjusting the scope parameter.
```

## Comments on Tracked Changes

Comments work on text inside tracked changes:

```python
# Comment on tracked insertion
doc.insert_tracked("new clause", after="Agreement")
doc.add_comment("Review this addition", on="new clause")

# Comment on tracked deletion
doc.delete_tracked("old term")
doc.add_comment("Why was this removed?", on="old term")

# Comments can span tracked/untracked boundaries
doc.add_comment("Check this section", on="normal inserted")
```

## Comment Replies

```python
# Add reply to existing comment
comment = doc.add_comment("Please review", on="Section 2.1")
reply = doc.add_comment(
    "Reviewed and approved",
    reply_to=comment,
    author="Approver"
)

# Reply by comment ID
reply = doc.add_comment("Done", reply_to="0")  # Comment ID as string
reply = doc.add_comment("Done", reply_to=0)    # Or as int
```

## Reading Comments

```python
# Access all comments
for comment in doc.comments:
    print(f"{comment.author}: {comment.text}")
    print(f"  On: '{comment.marked_text}'")
    print(f"  Date: {comment.date}")
    print(f"  Resolved: {comment.is_resolved}")

# Check for replies
for comment in doc.comments:
    if comment.parent:
        print(f"Reply to {comment.parent.id}")
    if comment.replies:
        print(f"Has {len(comment.replies)} replies")
```

## Comment Resolution

```python
# Resolve a comment
doc.resolve_comment(comment)

# Unresolve
doc.unresolve_comment(comment)

# Check resolution status
if comment.is_resolved:
    print("Comment is resolved")
```

## Deleting Comments

```python
# Delete specific comment
doc.delete_comment(comment)
doc.delete_comment("0")  # By ID string
doc.delete_comment(0)    # By ID int

# Delete all comments
doc.delete_all_comments()
```

## Comment Properties

| Property | Description |
|----------|-------------|
| `id` | Unique comment identifier |
| `text` | Comment content |
| `author` | Author name |
| `initials` | Author initials |
| `date` | Creation timestamp |
| `marked_text` | Text the comment is attached to |
| `is_resolved` | Resolution status |
| `parent` | Parent comment (for replies) |
| `replies` | List of reply comments |
| `para_id` | Internal paragraph ID |
