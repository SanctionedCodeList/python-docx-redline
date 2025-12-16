# Testing Documentation

## Overview

The python_docx_redline test suite provides comprehensive coverage of all core functionality with **168 tests achieving 92% code coverage**.

## Test Organization

### Test Files

The test suite is organized into focused test modules:

#### Core Functionality Tests

- **`test_document.py`** (11 tests)
  - Document loading from .docx and raw XML
  - Basic insert/delete/replace operations
  - Context manager support
  - Accept all changes functionality
  - Comment deletion

#### Phase 1: Text Operations

- **`test_scope.py`** (12 tests)
  - String scope specifications (`"section:Name"`, `"paragraph_containing:text"`)
  - Dictionary scopes with filters (`{"contains": "x", "not_contains": "y"}`)
  - Callable scope functions
  - Section detection via heading styles
  - Combined filter logic

- **`test_batch_operations.py`** (19 tests)
  - `apply_edits()` with multiple operations
  - Success/failure tracking via `EditResult`
  - Stop-on-error behavior
  - Missing parameter validation
  - Unknown edit type handling
  - Custom author support

- **`test_yaml_support.py`** (12 tests)
  - YAML file loading and parsing
  - JSON file support
  - Validation (missing keys, invalid structure)
  - Scope specifications in YAML
  - Error handling (file not found, parse errors)
  - Metadata handling

- **`test_suggestions.py`** (10 tests)
  - Smart error suggestions
  - Common issue detection:
    - Curly vs straight quotes
    - Double spaces
    - Leading/trailing whitespace
    - Case sensitivity
    - Special characters (non-breaking spaces, zero-width spaces, tabs)

#### Phase 2: Structural Operations

- **`test_paragraph.py`** (27 tests)
  - Paragraph wrapper class functionality
  - Text extraction and iteration
  - Parent document access
  - Paragraph properties

- **`test_section.py`** (24 tests)
  - Section wrapper class functionality
  - Heading detection and navigation
  - Section content iteration
  - Parent document access

- **`test_structural_operations.py`** (29 tests)
  - `insert_paragraph()` with styles and tracking
  - `insert_paragraphs()` for multiple paragraphs
  - `delete_section()` with tracked changes
  - Before/after positioning
  - Scope support for structural operations
  - Error handling

#### Phase 2.5: Regex Support

- **`test_regex_operations.py`** (14 tests)
  - Basic regex patterns (digits, emails, currency)
  - Capture group replacements with `\1`, `\2` syntax
  - Insert/delete operations with regex
  - Error handling for invalid patterns
  - Batch operations with regex
  - Scope filtering with regex

#### Integration Tests

- **`test_integration.py`** (10 tests)
  - **`test_complete_workflow`**: Realistic multi-operation document editing
    - Party name replacements with scoping
    - Payment term updates
    - Section-specific insertions
    - Text deletions
    - Verification of tracked changes

  - **`test_batch_workflow_with_yaml`**: End-to-end YAML workflow
    - Load edits from YAML file
    - Apply multiple scoped operations
    - Verify all changes applied

  - **`test_scoped_edits_workflow`**: Complex scoping scenarios
    - Section-based filtering
    - Paragraph-containing filters
    - Dictionary scopes

  - **`test_error_recovery_workflow`**: Error handling in batch operations
    - Continue-on-error behavior
    - Result tracking for failed operations
    - Successful operations complete despite failures

  - **`test_context_manager_workflow`**: Context manager usage
    - Automatic resource cleanup
    - Multiple operations in context
    - Save within context

## Running Tests

### Run All Tests

```bash
pytest
```

### Run Specific Test File

```bash
pytest tests/test_scope.py -v
```

### Run with Coverage

```bash
pytest --cov=src/python_docx_redline --cov-report=html
```

View detailed coverage report: `open htmlcov/index.html`

### Run Specific Test Function

```bash
pytest tests/test_document.py::test_insert_tracked_basic -v
```

## Test Helpers

### Document Creation

Most test files include helper functions to create test documents:

```python
def create_test_document() -> Path:
    """Create a test Word document."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))
    xml_content = """<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p><w:r><w:t>Test content</w:t></w:r></w:p>
      </w:body>
    </w:document>"""
    doc_path.write_text(xml_content, encoding="utf-8")
    return doc_path
```

For integration tests, `create_realistic_docx()` creates a multi-section document that mimics real-world legal contracts.

### Cleanup

All tests use try/finally blocks to ensure temporary files are cleaned up:

```python
doc_path = create_test_document()
try:
    doc = Document(doc_path)
    # ... perform tests ...
finally:
    doc_path.unlink()
```

## Coverage Report

Current coverage by module:

| Module | Coverage | Notes |
|--------|----------|-------|
| `__init__.py` | 100% | All exports covered |
| `tracked_xml.py` | 97% | Tracked change XML generation |
| `suggestions.py` | 95% | Smart error suggestions |
| `scope.py` | 93% | Scope filtering system |
| `results.py` | 92% | Result dataclasses |
| `text_search.py` | 87% | Text search with fragmentation |
| `document.py` | 84% | Core document operations |
| `errors.py` | 84% | Custom exceptions |

**Overall: 87% coverage**

## Test Data

### Minimal Valid OOXML

The tests use minimal valid OOXML structures:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Paragraph text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
```

For full .docx files, tests create ZIP archives with the necessary structure:

```python
with zipfile.ZipFile(doc_path, "w") as docx:
    docx.writestr("word/document.xml", document_xml)
    docx.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
    docx.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
```

## What's Tested

### ✅ Core Operations
- Text insertion with tracked changes
- Text deletion with tracked changes
- Text replacement (delete + insert)
- Fragmented text search (text split across multiple runs)
- Ambiguous text detection

### ✅ Scope System
- No scope (all paragraphs)
- String scopes (section names, contains text)
- Dictionary scopes with multiple filters
- Callable scope functions
- Section detection via heading styles

### ✅ Batch Operations
- Multiple edits in sequence
- Success/failure tracking
- Stop-on-error vs continue-on-error
- Parameter validation
- Unknown operation handling

### ✅ YAML/JSON Support
- File loading and parsing
- Validation (structure, required fields)
- Scope specifications in files
- Error handling
- Metadata (ignored but allowed)

### ✅ Error Handling
- Smart suggestions for common issues
- Curly quote detection
- Whitespace issues
- Special character detection
- Helpful error messages

### ✅ Integration Scenarios
- Multi-step document editing workflows
- Scoped operations across sections
- YAML-driven batch processing
- Error recovery in batch operations
- Context manager usage

## What's Not Tested

Areas with lower coverage (potential future test expansion):

- **Complex OOXML structures**: Tests use minimal valid structures
- **Real Word documents**: Tests use programmatically created documents
- **Edge cases**: Some error paths and edge cases in document.py
- **Accept/Reject changes**: Basic coverage, could expand
- **Comments**: Basic deletion tested, creation not covered

## Adding New Tests

When adding new tests:

1. **Choose the right test file**:
   - Core operations → `test_document.py`
   - Scope filtering → `test_scope.py`
   - Batch operations → `test_batch_operations.py`
   - YAML/JSON → `test_yaml_support.py`
   - Error suggestions → `test_suggestions.py`
   - End-to-end workflows → `test_integration.py`

2. **Use consistent patterns**:
   - Create test documents in try/finally blocks
   - Use tempfile for temporary files
   - Clean up all temporary files
   - Use descriptive test names

3. **Follow naming conventions**:
   - Test functions: `test_<feature>_<scenario>`
   - Helper functions: `create_<thing>`
   - Use docstrings to describe test purpose

4. **Test both success and failure**:
   - Happy path: operation succeeds
   - Sad path: appropriate error raised
   - Edge cases: boundary conditions

## Continuous Integration

The project uses pre-commit hooks to ensure code quality:

```bash
# Install pre-commit hooks
pre-commit install

# Run manually
pre-commit run --all-files
```

Hooks include:
- **ruff**: Fast Python linter
- **mypy**: Static type checking
- **pytest**: Run test suite
- **coverage**: Ensure minimum coverage

## Test Performance

The full test suite runs in ~0.25 seconds:

```
============================== 61 passed in 0.24s ==============================
```

Individual test files are even faster:
- `test_scope.py`: ~0.05s
- `test_batch_operations.py`: ~0.06s
- `test_yaml_support.py`: ~0.07s
- `test_integration.py`: ~0.14s

Fast tests enable rapid iteration during development.
