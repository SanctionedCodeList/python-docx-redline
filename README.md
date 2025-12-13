# python-docx-redline

A high-level Python API for editing Word documents with tracked changes.

## Claude Code Plugin

Install as a [Claude Code](https://claude.com/claude-code) plugin for OOXML reference documentation and helper scripts:

```
/plugin marketplace add SanctionedCodeList/python-docx-redline
```

Or install directly:

```
/plugin install python-docx-redline@SanctionedCodeList/python-docx-redline
```

Restart Claude Code after installation.

## Why python-docx-redline?

Making surgical edits to Word documents with tracked changes typically requires writing complex OOXML XML manipulation code. **python-docx-redline** reduces this from 30+ lines of raw XML to 3 lines of clean Python.

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

**After** (python_docx_redline):
```python
from python_docx_redline import Document

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

**Phase 4 - MS365 Identity Integration:**
- ✅ **AuthorIdentity support** - link tracked changes to real MS365/Office365 users
- ✅ **Full profile information** - include email, GUID, and provider ID
- ✅ **Enterprise-ready** - changes appear with complete user profile in Word
- ✅ **Backward compatible** - simple string author names still supported

**Phase 5 - Context-Aware Editing:**
- ✅ **Context preview** - see text before/after matches to verify correctness
- ✅ **Fragment detection** - automatic warnings for potential sentence fragments
- ✅ **Smart heuristics** - detects lowercase starts, connecting phrases, continuation punctuation
- ✅ **Helpful suggestions** - actionable guidance when issues detected

**Phase 6 - Minimal Editing Mode:**
- ✅ **Legal-style tracked changes** - only show actual text modifications
- ✅ **Preserves formatting runs** - no spurious formatting changes in Word's review pane
- ✅ **Clean diffs** - ideal for legal document review and contract redlining
- ✅ **Optional per-operation** - enable with `minimal=True` parameter

**Phase 7 - Markdown Formatting:**
- ✅ **Inline formatting** - use `**bold**`, `*italic*`, `++underline++`, `~~strikethrough~~`
- ✅ **Mixed formatting** - combine styles like `***bold italic***`
- ✅ **Line breaks** - proper `<w:br/>` elements for multi-line insertions
- ✅ **Always-on** - markdown syntax automatically parsed in insertions

**Phase 8 - Format-Only Tracked Changes:**
- ✅ **Track formatting changes** - bold, italic, underline, strikethrough with revision marks
- ✅ **Run-level precision** - format specific text within paragraphs
- ✅ **Paragraph formatting** - track alignment and style changes
- ✅ **Accept/reject support** - `accept_by_author()` and `reject_by_author()` include format changes

**General:**
- ✅ **python-docx integration** - seamlessly convert between libraries
- ✅ **In-memory workflows** - load from bytes/BytesIO, save to bytes
- ✅ **Batch operations** - apply multiple edits efficiently
- ✅ **YAML/JSON support** - define edits in configuration files
- ✅ **Type hints** - full type annotation support
- ✅ **Thoroughly tested** - 661 tests with 80% coverage

## Installation

```bash
pip install python-docx-redline
```

**Requirements:**
- Python 3.10+
- lxml
- python-dateutil
- pyyaml
- mistune (for markdown formatting)

## Quick Start

### Basic Operations

```python
from python_docx_redline import Document

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

### python-docx Integration

Seamlessly integrate with python-docx for the best of both libraries:

```python
from docx import Document as PythonDocxDocument
from python_docx_redline import Document, from_python_docx, to_python_docx

# Create a document with python-docx
py_doc = PythonDocxDocument()
py_doc.add_heading("Contract", 0)
py_doc.add_paragraph("Payment terms: net 30 days")
py_doc.add_paragraph("Effective date: January 1, 2025")

# Convert to python_docx_redline for tracked changes
doc = from_python_docx(py_doc, author="Legal Team")

# Make tracked edits
doc.replace_tracked("net 30 days", "net 45 days")
doc.insert_tracked(" (as amended)", after="Contract")

# Save the result
doc.save("contract_redlined.docx")
```

**In-memory workflows** (no filesystem required):

```python
# Load from bytes
with open("contract.docx", "rb") as f:
    doc = Document(f.read())

# Make changes
doc.insert_tracked(" [REVIEWED]", after="Section 1")

# Get bytes for storage/transmission
doc_bytes = doc.save_to_bytes(validate=False)
# Store in database, send over network, etc.
```

**Round-trip between libraries**:

```python
# Start with python_docx_redline
doc = Document("contract.docx")
doc.replace_tracked("old term", "new term")

# Convert back to python-docx for additional operations
py_doc = to_python_docx(doc, validate=False)
py_doc.add_paragraph("Added with python-docx")
py_doc.save("final.docx")
```

### Structural Operations (Phase 2)

Add, remove, and reorganize document structure with tracked changes:

```python
from python_docx_redline import Document

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
from python_docx_redline import Document

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
from python_docx_redline import Document

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

### MS365/Office365 Identity Integration (Phase 4)

Link tracked changes to real MS365 users with full identity information:

```python
from python_docx_redline import Document, AuthorIdentity

# Create an identity with MS365 profile info
identity = AuthorIdentity(
    author="Hancock, Parker",
    email="parker.hancock@company.com",
    provider_id="AD",  # Active Directory (default)
    guid="c5c513d2-1f51-4d69-ae91-17e5787f9bfc"  # User's unique ID
)

# Use identity when creating document
doc = Document("contract.docx", author=identity)

# All tracked changes will include full MS365 identity
doc.insert_tracked(" (amended)", after="Section 1")
doc.replace_tracked("30 days", "45 days")
doc.delete_tracked("optional clause")

doc.save("contract_edited.docx")
# Changes now appear in Word with full user profile and avatar
```

**How to find existing identity info:**
1. Unpack an existing .docx file (`unzip document.docx`)
2. Inspect `word/people.xml` for author information
3. Look for `w15:userId` (GUID) and `w15:providerId` attributes

**Benefits:**
- Changes show real user names and profile pictures in Word
- Better audit trail for enterprise environments
- Integrates with Office 365 user directory
- Backward compatible - simple string author names still work

### Context-Aware Editing (Phase 5)

Prevent sentence fragments and verify replacements with context preview and automatic fragment detection:

```python
import warnings
from python_docx_redline import Document, ContinuityWarning

doc = Document("contract.docx")

# Preview context before/after replacement
doc.replace_tracked(
    find="old text",
    replace="new text",
    show_context=True,       # Show surrounding text
    context_chars=100        # Characters to show before/after
)

# Output shows:
# ================================================================================
# CONTEXT PREVIEW
# ================================================================================
#
# BEFORE (78 chars):
#   '...The Buyer shall have the right to terminate this Agreement within'
#
# MATCH (8 chars):
#   'old text'
#
# AFTER (79 chars):
#   'of the Effective Date by providing written notice to the Seller...'
#
# REPLACEMENT (8 chars):
#   'new text'
# ================================================================================

# Enable automatic fragment detection
warnings.simplefilter("always")  # Enable warnings

doc.replace_tracked(
    find="The product in Vrdolyak was an attorney directory.",
    replace="BatchLeads is different.",
    check_continuity=True    # Warn about potential fragments
)

# If next text is " in question here is property data.", warns:
# ContinuityWarning: Next text starts with connecting phrase 'in question'
#   • Include more context in your replacement text
#   • Adjust the 'find' text to include the connecting phrase
#   • Review the following text to ensure grammatical correctness

# Use both features together
doc.replace_tracked(
    find="complex legal text",
    replace="simplified version",
    show_context=True,
    check_continuity=True,
    context_chars=100
)

doc.save("contract_edited.docx")
```

**Fragment Detection:**
- Detects lowercase starts: " in question here..." → warns about fragment
- Detects connecting phrases: " of which", " that is", " wherein", etc.
- Detects continuation punctuation: ", and...", "; however...", etc.

See [Context-Awareness Guide](docs/CONTEXT_AWARENESS_GUIDE.md) for detailed usage.

### Minimal Editing Mode (Phase 6)

For legal documents, use minimal editing mode to avoid spurious formatting changes:

```python
# Enable minimal editing for clean tracked changes
doc.replace_tracked(
    find="the Contractor",
    replace="the Service Provider",
    minimal=True  # Only shows actual text change in Word's review pane
)

# Perfect for legal document redlining
doc.insert_tracked(
    text=" (as amended and restated)",
    after="Agreement",
    minimal=True
)

doc.delete_tracked(
    text="for any reason whatsoever",
    minimal=True
)
```

**Why use minimal mode?**
- Legal reviewers see only the actual text modifications
- No confusing formatting changes in Word's Track Changes pane
- Cleaner diffs for contract comparison tools
- Matches expectations for professional document review

### Markdown Formatting (Phase 7)

Use markdown syntax in tracked insertions for inline formatting:

```python
# Bold and italic formatting
doc.insert_tracked(
    "**Important:** See *Smith v. Jones* for precedent",
    after="Section 2.1"
)

# Underline and strikethrough
doc.insert_tracked(
    "This clause is ++mandatory++ and ~~optional~~ required",
    after="Terms"
)

# Combined formatting
doc.insert_tracked(
    "***Critical Notice:*** Review ++immediately++",
    after="Exhibit A"
)

# Multi-line with proper line breaks
doc.insert_tracked(
    "First line  \nSecond line",  # Two spaces + newline = line break
    after="Introduction"
)
```

**Supported Markdown:**
| Syntax | Result | OOXML |
|--------|--------|-------|
| `**text**` | **bold** | `<w:b/>` |
| `*text*` | *italic* | `<w:i/>` |
| `++text++` | underline | `<w:u/>` |
| `~~text~~` | ~~strikethrough~~ | `<w:strike/>` |

### Format-Only Tracked Changes (Phase 8)

Track formatting changes without modifying text:

```python
# Track bold formatting change
result = doc.format_tracked(
    find="Important Notice",
    bold=True
)
print(f"Applied bold to '{result.text_matched}'")

# Track multiple formatting changes
doc.format_tracked(
    find="CONFIDENTIAL",
    bold=True,
    italic=True,
    underline=True
)

# Remove formatting (explicit False)
doc.format_tracked(
    find="previously bold text",
    bold=False  # Removes bold with tracked change
)

# Format entire paragraph
doc.format_paragraph_tracked(
    paragraph_index=0,
    alignment="center",
    style="Heading1"
)

# Accept/reject includes format changes
doc.accept_by_author("Claude")  # Accepts text AND formatting changes
doc.reject_by_author("Claude")  # Rejects text AND formatting changes
```

**FormatResult provides details:**
```python
result = doc.format_tracked(find="text", bold=True)
print(result.success)           # True if operation succeeded
print(result.changed)           # True if formatting actually changed
print(result.text_matched)      # The text that was formatted
print(result.changes_applied)   # {'bold': True}
print(result.previous_formatting)  # Previous state per run
print(result.change_id)         # OOXML change ID for tracking
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

#### `save(path=None, validate=True)`
Save the document. If path is None, overwrites the original file. For in-memory documents, `path` is required.

#### `save_to_bytes(validate=True)`
Save the document to bytes (in-memory). Useful for passing documents between libraries, storing in databases, or sending over network.

### Compatibility Functions

#### `from_python_docx(python_docx_doc, author="Claude")`
Create a python_docx_redline Document from a python-docx Document. Enables workflows where you create documents with python-docx and then add tracked changes.

#### `to_python_docx(doc, validate=True)`
Convert a python_docx_redline Document back to a python-docx Document. Useful when you need python-docx's document creation features after making tracked changes.

### Scope Specifications

Scopes limit where operations apply:

- **String**: `"text"` - Paragraphs containing "text"
- **Section**: `"section:Introduction"` - Paragraphs in section with heading "Introduction"
- **Explicit**: `"paragraph_containing:specific text"` - Paragraphs with "specific text"
- **Dictionary**: `{"contains": "text", "not_contains": "exclude"}` - Complex filters
- **Callable**: Custom function `lambda p: len(''.join(p.itertext())) > 100`

## Error Handling

python_docx_redline provides helpful error messages with suggestions:

```python
from python_docx_redline import Document, TextNotFoundError

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
git clone https://github.com/parkerhancock/python_docx_redline.git
cd python_docx_redline

# Install in development mode with dev dependencies
pip install -e ".[dev]"
```

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=src/python_docx_redline --cov-report=html

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

**Phase 6 - Minimal Editing Mode: Complete** ✅
- ✅ Legal-style tracked changes with `minimal=True`
- ✅ Preserves formatting runs for clean review pane
- ✅ Support for insert, delete, and replace operations

**Phase 7 - Markdown Formatting: Complete** ✅
- ✅ Bold, italic, underline, strikethrough via markdown syntax
- ✅ Automatic parsing in `insert_tracked()` calls
- ✅ Proper `<w:br/>` line break support

**Phase 8 - Format-Only Tracked Changes: Complete** ✅
- ✅ `format_tracked()` for run-level formatting changes
- ✅ `format_paragraph_tracked()` for paragraph formatting
- ✅ Accept/reject support for formatting changes

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

- **Issues**: [GitHub Issues](https://github.com/parkerhancock/python_docx_redline/issues)
- **Discussions**: [GitHub Discussions](https://github.com/parkerhancock/python_docx_redline/discussions)
- **Email**: parker@parkerhancock.com
