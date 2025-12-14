# python-docx-redline Roadmap

This document outlines planned features and enhancements for the python-docx-redline library.

## Current State (v0.1.x)

The library currently supports:

**Text Operations**
- `insert_tracked()` - insert text with tracked changes
- `replace_tracked()` - replace text with tracked changes
- `delete_tracked()` - delete text with tracked changes
- `move_tracked()` - relocate text with linked move markers
- Smart text search (handles fragmented XML runs)
- Regex support with capture groups
- Scope filtering (sections/paragraphs)

**Structural Operations**
- `insert_paragraph()` - add new paragraphs with styles
- `delete_section()` - remove sections by heading
- Paragraph and Section wrapper classes

**Table Operations**
- `tables` property - access all tables
- `find_table()` - find table by content
- `replace_in_table()` - replace cell content
- `insert_table_row()` - insert rows with tracked changes
- `delete_table_row()` - delete rows with tracked changes

**Comments**
- `add_comment()` - add comments with reply support
- `get_comments()` - retrieve all comments
- `delete_all_comments()` - remove all comments
- `comment.resolve()` / `comment.unresolve()` - mark resolved status
- `comment.is_resolved` - check resolution status

**Footnotes & Endnotes**
- `footnotes` / `endnotes` properties
- `insert_footnote()` - add footnotes at text locations
- `insert_endnote()` - add endnotes at text locations

**Change Management**
- `accept_change(id)` - accept specific tracked change
- `reject_change(id)` - reject specific tracked change
- `accept_all_changes()` - accept all tracked changes
- `has_tracked_changes()` - check if document has changes

**Document Viewing**
- `paragraphs` - access all paragraphs
- `sections` - access document sections
- `text` - extract full document text
- Context-aware editing with fragment detection

**Formatting & Identity**
- MS365 identity integration (`AuthorIdentity`)
- Minimal editing mode (legal-style clean diffs)
- Markdown formatting (bold, italic, underline, strikethrough)
- Format-only tracking (track formatting without text changes)

---

## Phase 9: List Tracked Changes

**Priority**: High
**Complexity**: Low
**Status**: Planned

### Overview

Add ability to enumerate all tracked changes in a document with their metadata. Currently you can accept/reject changes by ID, but there's no way to programmatically list them.

### Proposed API

```python
# List all tracked changes
changes = doc.get_tracked_changes()
for change in changes:
    print(f"{change.id}: {change.type} by {change.author}")
    print(f"  Text: '{change.text}'")
    print(f"  Date: {change.date}")

# Filter by type
insertions = doc.get_tracked_changes(change_type="insertion")
deletions = doc.get_tracked_changes(change_type="deletion")

# Filter by author
legal_changes = doc.get_tracked_changes(author="Legal Team")

# Bulk accept/reject by criteria
doc.accept_changes(author="Legal Team")
doc.reject_changes(change_type="deletion")
```

### Use Cases

- Review changes before accepting/rejecting
- Generate change reports
- Selective bulk acceptance by author or type

---

## Phase 10: CLI Tool

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

# Move text
docx-redline move contract.docx \
    --text "Section A" \
    --after "Table of Contents"

# Accept all changes
docx-redline accept-all input.docx --output clean.docx

# List tracked changes
docx-redline changes contract.docx --format json

# Apply edits from YAML
docx-redline apply contract.docx edits.yaml --output result.docx

# Add comment
docx-redline comment contract.docx \
    --on "Section 2.1" \
    --text "Please review"
```

### Use Cases

- Shell scripting for document automation
- CI/CD pipelines for document generation
- Quick edits without writing Python code

---

## Phase 11: Table Column Operations

**Priority**: Medium
**Complexity**: Medium
**Status**: Planned

### Overview

Add column-level operations for tables. Row operations exist, but column operations are not yet implemented.

### Proposed API

```python
# Insert a column with tracked changes
doc.insert_table_column(
    table_index=0,
    after_column=1,
    header="New Column",
    cells=["A", "B", "C"],
    track=True,
    author="Finance"
)

# Delete a column with tracked changes
doc.delete_table_column(
    table_index=0,
    column=2,
    track=True,
    author="Editor"
)
```

### Use Cases

- Adding new data columns to pricing tables
- Removing obsolete columns from schedules
- Restructuring tabular data

---

## Phase 12: Document Comparison

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

## Phase 13: Header/Footer Editing

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

## Phase 14: Export/Visualization

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
| 0.2.0 | List Tracked Changes (Phase 9) |
| 0.3.0 | CLI Tool (Phase 10) |
| 0.4.0 | Table Column Operations (Phase 11) |
| 0.5.0 | Document Comparison (Phase 12) |
| 1.0.0 | Stable API, comprehensive documentation |
