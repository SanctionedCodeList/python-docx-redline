# docx_redline

A high-level Python API for editing Word documents with tracked changes.

## Why docx_redline?

Making surgical edits to Word documents with tracked changes typically requires writing complex OOXML XML manipulation code. **docx_redline** reduces this from 30+ lines of raw XML to 3 lines of clean Python.

**Before** (raw OOXML):
```python
# 30+ lines of lxml manipulation, namespace handling, ID generation...
from lxml import etree
from datetime import datetime, timezone

# Find the paragraph
paragraphs = root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
target_para = None
for p in paragraphs:
    text = ''.join(p.itertext())
    if 'Section 2.1' in text:
        target_para = p
        break

# Generate tracked change XML
change_id = get_next_change_id()
timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
insertion = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins')
insertion.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(change_id))
insertion.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Author Name')
insertion.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', timestamp)
# ... more XML construction ...
```

**After** (docx_redline):
```python
from docx_redline import Document

doc = Document("contract.docx")
doc.insert_tracked("new clause text", after="Section 2.1")
doc.save("contract_edited.docx")
```

## Features

**Phase 1 - Text Operations:**
- ✅ **Insert, delete, and replace text** with tracked changes
- ✅ **Smart text search** - handles text fragmented across multiple XML runs
- ✅ **Regex support** - use regular expressions with capture groups for powerful find/replace
- ✅ **Scope filtering** - limit operations to specific sections or paragraphs
- ✅ **Helpful error messages** - suggestions for common issues (curly quotes, whitespace, etc.)

**Phase 2 - Structural Operations:**
- ✅ **Insert complete paragraphs** - add new paragraphs with styles and formatting
- ✅ **Insert multiple paragraphs** - efficiently add multiple paragraphs at once
- ✅ **Delete entire sections** - remove sections by heading with tracked changes
- ✅ **Section and Paragraph wrappers** - convenient API for document structure

**Phase 3 - Document Viewing:**
- ✅ **Read paragraphs** - access all paragraphs with text, style, and heading info
- ✅ **Parse sections** - automatic section detection via heading structure
- ✅ **Extract text** - get full document text for analysis
- ✅ **Agent workflow** - read → understand → targeted edits

**General:**
- ✅ **Batch operations** - apply multiple edits efficiently
- ✅ **YAML/JSON support** - define edits in configuration files
- ✅ **Type hints** - full type annotation support
- ✅ **Thoroughly tested** - 182 tests with 92% coverage

## Installation

```bash
pip install docx-redline
```

**Requirements:**
- Python 3.10+
- lxml
- python-dateutil
- pyyaml

## Quick Start

### Basic Operations

```python
from docx_redline import Document

# Load a document
doc = Document("contract.docx")

# Insert text with tracked changes
doc.insert_tracked(" (as amended)", after="Section 2.1")

# Replace text with tracked changes
doc.replace_tracked(
    find="the Contractor",
    replace="the Service Provider"
)

# Delete text with tracked changes
doc.delete_tracked("for any reason")

# Save the modified document
doc.save("contract_edited.docx")
```

### Structural Operations (Phase 2)

Add, remove, and reorganize document structure with tracked changes:

```python
from docx_redline import Document

doc = Document("contract.docx")

# Insert a new paragraph with a specific style
doc.insert_paragraph(
    "New Section Heading",
    after="Introduction content",
    style="Heading1",
    track=True
)

# Insert multiple paragraphs at once
doc.insert_paragraphs(
    [
        "First key point",
        "Second key point",
        "Third key point"
    ],
    after="New Section Heading",
    track=True
)

# Delete an entire section by its heading
doc.delete_section("Outdated Provisions", track=True)

# Mix structural and text operations
doc.replace_tracked("old term", "new term")
doc.insert_paragraph("Amendments", after="Section 5", style="Heading1", track=True)

doc.save("contract_restructured.docx")
```

### Regex Operations (Phase 2.5)

Use regular expressions for powerful pattern matching and replacements:

```python
from docx_redline import Document

doc = Document("contract.docx")

# Replace all dollar amounts with redacted version
doc.replace_tracked(r"\$[\d,]+\.?\d*", "$XXX.XX", regex=True)

# Update all occurrences of "X days" to "X business days" using capture groups
doc.replace_tracked(r"(\d+) days", r"\1 business days", regex=True)

# Swap date format from MM/DD/YYYY to DD/MM/YYYY
doc.replace_tracked(
    r"(\d{2})/(\d{2})/(\d{4})",
    r"\2/\1/\3",
    regex=True
)

# Insert text after any section reference
doc.insert_tracked(" (as amended)", after=r"Section \d+\.\d+", regex=True)

# Delete all email addresses
doc.delete_tracked(r"\b[a-z]+@[a-z]+\.com\b", regex=True)

doc.save("contract_regex_edited.docx")
```

###  Reading Document Content (Phase 3)

View and analyze document structure before making edits:

```python
from docx_redline import Document

doc = Document("contract.docx")

# Get all paragraphs
for para in doc.paragraphs:
    if para.is_heading():
        print(f"Section: {para.text}")
    else:
        print(f"  {para.text}")

# Get document sections (parsed by headings)
for section in doc.sections:
    if section.heading:
        print(f"\nSection: {section.heading_text} (Level {section.heading_level})")
        print(f"  {len(section.paragraphs)} paragraphs")

    # Find specific text within a section
    if section.contains("confidential"):
        print("  Contains confidential information")

# Get full document text
text = doc.get_text()
if "arbitration" in text.lower():
    print("Document contains arbitration clause")

# Agent workflow: read → understand → edit
for section in doc.sections:
    if section.heading_text == "Payment Terms":
        # Found the section, check its content
        if section.contains("net 30 days"):
            # Now make targeted edit
            doc.replace_tracked("net 30 days", "net 45 days", scope="Payment Terms")

doc.save("contract_updated.docx")
```

### Using Scopes

Limit operations to specific sections or paragraphs:

```python
# Only modify text in the Introduction section
doc.insert_tracked(
    text=" (hereinafter 'Agreement')",
    after="this Agreement",
    scope="section:Introduction"
)

# Only modify in paragraphs containing specific text
doc.replace_tracked(
    find="Client",
    replace="Customer",
    scope="paragraph_containing:payment terms"
)

# Dictionary scope for complex filtering
doc.insert_tracked(
    text=" (Non-Disclosure Agreement)",
    after="NDA",
    scope={
        "contains": "confidential",
        "section": "Definitions"
    }
)
```

### Batch Operations

Apply multiple edits at once:

```python
edits = [
    {
        "type": "insert_tracked",
        "text": " (revised)",
        "after": "Exhibit A"
    },
    {
        "type": "replace_tracked",
        "find": "30 days",
        "replace": "45 days"
    },
    {
        "type": "delete_tracked",
        "text": "unless otherwise agreed"
    }
]

results = doc.apply_edits(edits)

# Check results
for result in results:
    print(result)  # ✓ insert_tracked: Inserted ' (revised)' after 'Exhibit A'
```

### YAML Configuration Files

Define edits in a YAML file for repeatable workflows:

```yaml
# edits.yaml
edits:
  # Phase 1: Text operations
  - type: insert_tracked
    text: " (as amended)"
    after: "Agreement dated"
    scope: "section:Recitals"

  - type: replace_tracked
    find: "Contractor"
    replace: "Service Provider"

  - type: delete_tracked
    text: "subject to approval"
    scope:
      contains: "termination"

  # Phase 2: Structural operations
  - type: insert_paragraph
    text: "Compliance"
    after: "Section 5"
    style: "Heading1"
    track: true

  - type: insert_paragraphs
    texts:
      - "All parties shall comply with applicable laws."
      - "This includes federal, state, and local regulations."
    after: "Compliance"
    track: true

  - type: delete_section
    heading: "Deprecated Clause"
    track: true

  # Regex operations
  - type: replace_tracked
    find: "(\\d+) days"
    replace: "\\1 business days"
    regex: true
```

```python
# Apply edits from file
results = doc.apply_edit_file("edits.yaml")
print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
```

## Core API

### Document Methods

#### `insert_tracked(text, after, author=None, scope=None, regex=False)`
Insert text with tracked changes after a specific location. Set `regex=True` to use regex patterns.

#### `delete_tracked(text, author=None, scope=None, regex=False)`
Mark text for deletion with tracked changes. Set `regex=True` to use regex patterns.

#### `replace_tracked(find, replace, author=None, scope=None, regex=False)`
Replace text with tracked changes (combines delete + insert). Set `regex=True` to use regex patterns with capture group support.

#### `insert_paragraph(text, after=None, before=None, style=None, track=False, author=None, scope=None)`
Insert a complete paragraph with optional style and tracked changes.

#### `insert_paragraphs(texts, after=None, before=None, styles=None, track=False, author=None, scope=None)`
Insert multiple paragraphs at once. Styles can be a single style for all paragraphs or a list matching texts.

#### `delete_section(heading, track=False, author=None, scope=None)`
Delete an entire section (heading and all content until next heading) with tracked changes.

#### `apply_edits(edits, stop_on_error=False)`
Apply multiple edits in sequence. Returns `list[EditResult]`.

#### `apply_edit_file(path, format="yaml", stop_on_error=False)`
Load and apply edits from a YAML or JSON file.

#### `save(path=None)`
Save the document. If path is None, overwrites the original file.

### Scope Specifications

Scopes limit where operations apply:

- **String**: `"text"` - Paragraphs containing "text"
- **Section**: `"section:Introduction"` - Paragraphs in section with heading "Introduction"
- **Explicit**: `"paragraph_containing:specific text"` - Paragraphs with "specific text"
- **Dictionary**: `{"contains": "text", "not_contains": "exclude"}` - Complex filters
- **Callable**: Custom function `lambda p: len(''.join(p.itertext())) > 100`

## Error Handling

docx_redline provides helpful error messages with suggestions:

```python
from docx_redline import Document, TextNotFoundError

doc = Document("contract.docx")

try:
    doc.insert_tracked("new text", after="nonexistent text")
except TextNotFoundError as e:
    print(e)
    # Output:
    # Could not find 'nonexistent text'
    #
    # Suggestions:
    #   • Check for typos in the search text
    #   • Try searching for a shorter or more unique phrase
    #   • Verify the text exists in the document
```

Common issues detected automatically:
- Curly quotes vs straight quotes
- Double spaces
- Leading/trailing whitespace
- Case sensitivity mismatches
- Special characters (non-breaking spaces, zero-width spaces, etc.)

## Advanced Usage

### Custom Authors

```python
doc.insert_tracked(
    "new provision",
    after="Section 3",
    author="Legal Team"
)
```

### Error Handling in Batch Operations

```python
# Stop on first error
results = doc.apply_edits(edits, stop_on_error=True)

# Continue on errors (default)
results = doc.apply_edits(edits, stop_on_error=False)

# Check individual results
for i, result in enumerate(results):
    if result.success:
        print(f"✓ Edit {i+1}: {result.message}")
    else:
        print(f"✗ Edit {i+1}: {result.message}")
        if result.error:
            print(f"  Error: {result.error}")
```

### Context Manager Support

```python
with Document("contract.docx") as doc:
    doc.insert_tracked("new clause", after="Section 1")
    doc.save("modified.docx")
# Automatic cleanup
```

## Examples

See the `examples/` directory for complete working examples:
- `surgical_edits.yaml` - Real-world legal document edits
- `batch_processing.py` - Processing multiple documents
- `scope_examples.py` - Advanced scope filtering

## Documentation

Comprehensive documentation available in the `docs/` directory:
- [API Reference](docs/PROPOSED_API.md) - Complete API documentation
- [Implementation Notes](docs/IMPLEMENTATION_NOTES.md) - Technical details
- [Quick Reference](docs/QUICK_REFERENCE.md) - Cheat sheet
- [Eric White's Algorithm](docs/ERIC_WHITE_ALGORITHM.md) - Text search algorithm reference

## Development

### Setup

```bash
# Clone the repository
git clone https://github.com/parkerhancock/docx_redline.git
cd docx_redline

# Install in development mode with dev dependencies
pip install -e ".[dev]"
```

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=src/docx_redline --cov-report=html

# Run specific test file
pytest tests/test_document.py -v
```

### Code Quality

The project uses:
- **ruff** - Fast Python linter
- **mypy** - Static type checking
- **pytest** - Testing framework
- **pre-commit hooks** - Automated quality checks

```bash
# Run pre-commit hooks manually
pre-commit run --all-files
```

## Project Status

**Phase 1 - Text Operations: Complete** ✅
- ✅ Text search with fragmentation handling
- ✅ Tracked changes (insert/delete/replace)
- ✅ Scope system for filtering
- ✅ Error handling with smart suggestions
- ✅ Batch operations
- ✅ YAML/JSON file support

**Phase 2 - Structural Operations: Complete** ✅
- ✅ Insert paragraphs with styles and formatting
- ✅ Insert multiple paragraphs efficiently
- ✅ Delete entire sections with tracked changes
- ✅ Section and Paragraph wrapper classes
- ✅ Full integration with batch/YAML workflows
- ✅ Comprehensive integration tests

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes with tests
4. Run the test suite (`pytest`)
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## License

MIT License - see LICENSE file for details

## Acknowledgments

- Inspired by the need for simpler Word document automation
- Text search algorithm based on Eric White's OOXML analysis
- Built with [lxml](https://lxml.de/) for robust XML processing

## Support

- **Issues**: [GitHub Issues](https://github.com/parkerhancock/docx_redline/issues)
- **Discussions**: [GitHub Discussions](https://github.com/parkerhancock/docx_redline/discussions)
- **Email**: parker@parkerhancock.com
