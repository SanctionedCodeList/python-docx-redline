# python-docx-redline Roadmap

This document outlines planned features and enhancements for the python-docx-redline library.

## Current State (v0.1.x)

The library currently supports:

- **Text Operations**: Insert, replace, delete with tracked changes
- **Smart Text Search**: Handles text fragmented across XML runs
- **Regex Support**: Pattern matching with capture groups
- **Scope Filtering**: Limit operations to sections/paragraphs
- **Structural Operations**: Insert/delete paragraphs and sections
- **Document Viewing**: Read paragraphs, sections, extract text
- **MS365 Identity**: Link changes to real Office365 users
- **Context-Aware Editing**: Preview context, detect fragments
- **Minimal Editing Mode**: Legal-style clean diffs
- **Markdown Formatting**: Bold, italic, underline, strikethrough
- **Format-Only Tracking**: Track formatting changes without text changes

---

## Phase 9: Move Tracking

**Priority**: High
**Complexity**: Medium
**Status**: Research Complete

### Overview

Implement OOXML move tracking (`w:moveFrom`/`w:moveTo`) to preserve semantic relationships when text is relocated within a document. Unlike delete+insert, move tracking shows reviewers exactly where content came from.

### Proposed API

```python
# Move text from one location to another
doc.move_tracked(
    text="the indemnification clause",
    to_after="Section 5.2",
    author="Legal Team"
)

# Move with scope
doc.move_tracked(
    text="Definitions",
    to_before="Article I",
    scope=Scope(section="Appendix A")
)
```

### Use Cases

- Legal document review where "moved from Section X" context matters
- Contract reorganization with full audit trail
- Academic editing showing paragraph reordering

### Technical Notes

- Research completed in `docs/OOXML_MOVE_TRACKING_RESEARCH.md`
- Requires paired `moveFromRangeStart`/`moveFromRangeEnd` and `moveToRangeStart`/`moveToRangeEnd`
- Move names link source to destination containers

---

## Phase 10: Comments API

**Priority**: High
**Complexity**: Low-Medium
**Status**: Planned

### Overview

Expand comment support beyond `delete_all_comments()` to full CRUD operations for document comments.

### Proposed API

```python
# Add a comment to specific text
doc.add_comment(
    text="ambiguous clause",
    comment="Please clarify the timeline here",
    author="Reviewer"
)

# Reply to existing comment
doc.reply_to_comment(
    comment_id=5,
    reply="Updated per your feedback",
    author="Author"
)

# Resolve a comment
doc.resolve_comment(comment_id=5)

# Get all comments
comments = doc.get_comments()
for c in comments:
    print(f"{c.author}: {c.text} (on: {c.anchor_text})")

# Delete specific comment
doc.delete_comment(comment_id=5)
```

### Use Cases

- Automated review workflows
- Comment migration between document versions
- Batch comment operations (resolve all, export to report)

---

## Phase 11: Accept/Reject Individual Changes

**Priority**: High
**Complexity**: Medium
**Status**: Planned

### Overview

Extend `accept_all_changes()` to support selective acceptance or rejection of individual tracked changes.

### Proposed API

```python
# List all tracked changes
changes = doc.get_tracked_changes()
for change in changes:
    print(f"{change.id}: {change.type} by {change.author} - '{change.text}'")

# Accept specific change
doc.accept_change(change_id=3)

# Reject specific change
doc.reject_change(change_id=5)

# Accept changes by author
doc.accept_changes(author="Legal Team")

# Reject changes by type
doc.reject_changes(change_type="deletion")

# Accept changes in scope
doc.accept_changes(scope=Scope(section="Article II"))
```

### Use Cases

- Selective merge of edits from multiple reviewers
- Automated acceptance of formatting-only changes
- Reject changes from specific authors

---

## Phase 12: CLI Tool

**Priority**: Medium
**Complexity**: Low
**Status**: Planned

### Overview

Command-line interface for common operations, enabling shell scripting and CI/CD integration.

### Proposed CLI

```bash
# Insert text
docx-redline insert contract.docx \
    --after "Section 2.1" \
    --text "New clause text" \
    --author "Claude" \
    --output contract_edited.docx

# Replace text
docx-redline replace contract.docx \
    --find "Acme Corp" \
    --replace "NewCo Inc" \
    --author "Legal"

# Delete text
docx-redline delete contract.docx \
    --text "obsolete paragraph" \
    --scope "section:Appendix"

# Accept all changes
docx-redline accept-all input.docx --output clean.docx

# Compare documents
docx-redline compare original.docx modified.docx --output redline.docx

# List tracked changes
docx-redline changes contract.docx --format json

# Apply edits from YAML
docx-redline apply contract.docx edits.yaml --output result.docx
```

### Use Cases

- Shell scripting for document automation
- CI/CD pipelines for document generation
- Quick edits without writing Python code

---

## Phase 13: Table Operations

**Priority**: Medium
**Complexity**: Medium-High
**Status**: Planned

### Overview

Enhanced table manipulation with tracked changes for rows, columns, and cells.

### Proposed API

```python
# Insert a row with tracked changes
doc.insert_table_row(
    table_index=0,
    after_row=2,
    cells=["Item 4", "$500", "2024-01-15"],
    author="Finance"
)

# Delete a row with tracked changes
doc.delete_table_row(
    table_index=0,
    row=3,
    author="Editor"
)

# Insert a column
doc.insert_table_column(
    table_index=0,
    after_column=1,
    header="New Column",
    cells=["A", "B", "C"]
)

# Replace cell content
doc.replace_in_table(
    text="TBD",
    replacement="$1,000",
    table_index=0,
    author="Finance"
)
```

### Use Cases

- Contract amendments with pricing table changes
- Financial document updates
- Automated table population

---

## Phase 14: Document Comparison

**Priority**: Medium
**Complexity**: Medium
**Status**: Planned

### Overview

Generate a redline document showing differences between two document versions.

### Proposed API

```python
from python_docx_redline import compare_documents

# Generate redline comparison
redline = compare_documents(
    original="contract_v1.docx",
    modified="contract_v2.docx",
    author="Comparison Tool"
)
redline.save("contract_redline.docx")

# Get comparison summary
diff = compare_documents(original, modified)
print(f"Insertions: {len(diff.insertions)}")
print(f"Deletions: {len(diff.deletions)}")
print(f"Moves: {len(diff.moves)}")
```

### Use Cases

- "What changed?" reporting
- Version control for documents
- Audit trail generation

---

## Phase 15: Header/Footer Editing

**Priority**: Low-Medium
**Complexity**: Medium
**Status**: Planned

### Overview

Edit document headers and footers with tracked changes.

### Proposed API

```python
# Edit header
doc.replace_in_header(
    text="Draft",
    replacement="Final",
    header_type="default",  # default, first, even
    author="Editor"
)

# Edit footer
doc.insert_in_footer(
    text=" - Confidential",
    after="Page {PAGE}",
    footer_type="default"
)

# Access header/footer content
for header in doc.headers:
    print(f"{header.type}: {header.text}")
```

### Use Cases

- Update document metadata (dates, version numbers)
- Add confidentiality notices
- Modify page numbering

---

## Phase 16: Export/Visualization

**Priority**: Low
**Complexity**: Low-Medium
**Status**: Planned

### Overview

Export tracked changes to alternative formats for visualization and reporting.

### Proposed API

```python
# Export to HTML diff view
html = doc.export_changes_html()

# Export to Markdown
md = doc.export_changes_markdown()

# Export change summary to JSON
summary = doc.export_changes_json()

# Generate change report
report = doc.generate_change_report(
    format="html",
    include_context=True,
    group_by="author"
)
```

### Use Cases

- Code review style visualization
- Change reports for stakeholders
- Integration with web-based review tools

---

## Future Considerations

These features may be considered based on user feedback:

- **Async/Batch Processing**: Concurrent processing of multiple documents
- **Image Operations**: Insert/replace images with tracked changes
- **Style Change Tracking**: Track paragraph/character style modifications
- **Field Code Support**: Update and track changes to Word field codes
- **Content Control Editing**: Manipulate structured content controls
- **Bookmark Operations**: Add/edit/delete bookmarks with tracking

---

## Contributing

We welcome contributions! If you're interested in implementing any of these features:

1. Check the existing research in `docs/` for technical details
2. Open an issue to discuss the implementation approach
3. Submit a PR with tests and documentation

## Versioning Plan

| Version | Features |
|---------|----------|
| 0.2.0 | Move Tracking (Phase 9) |
| 0.3.0 | Comments API (Phase 10) |
| 0.4.0 | Accept/Reject Changes (Phase 11) |
| 0.5.0 | CLI Tool (Phase 12) |
| 1.0.0 | Stable API, comprehensive documentation |
