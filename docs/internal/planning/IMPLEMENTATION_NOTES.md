# Implementation Notes for Dev Team

**Companion to**: PROPOSED_API.md
**Date**: December 6, 2025

---

## Quick Start for Developers

### Critical Files to Understand

1. **Current Implementation**:
   - `scripts/document.py` - Document class (lines 924-1220)
   - `scripts/utilities.py` - XMLEditor base class (lines 66-900)
   - `document-api.md` - Current API documentation

2. **Text Fragmentation Example**:
   - See `client_draft_unpacked/word/document.xml` lines 6587-6650
   - Shows how "Huston v. Hearst Communications" is split across multiple `<w:r>` elements

3. **Tracked Changes XML**:
   - `xml-reference.md` - OOXML tracked changes specification
   - Lines 50-150 show insertion/deletion structures

### Key Challenges This API Solves

**Challenge 1: Text Fragmentation**

```xml
<!-- Word shows: "The Seventh Circuit has made clear that" -->
<!-- XML contains: -->
<w:p>
  <w:r w:rsidR="00082A8E"><w:t>The Seventh </w:t></w:r>
  <w:r w:rsidR="00AB1234"><w:t>Circuit has </w:t></w:r>
  <w:r w:rsidR="00082A8E"><w:t>made clear that</w:t></w:r>
</w:p>
```

**Current approach** (30+ lines of code):
```python
para = editor.get_node(tag="w:p", line_number=6587)
runs = list(para.iter(_parse_tag("w:r")))
for run in runs:
    text = ''.join(run.itertext())
    if "(7th Cir. 2022)" in text:
        # ... insert XML ...
```

**Proposed approach** (1 line):
```python
doc.insert_tracked(" (interpreting IRPA)", after="(7th Cir. 2022)")
```

**Challenge 2: XML Construction**

Current users must write:
```xml
<w:ins w:id="4" w:author="Hancock, Parker" w:date="2025-12-06T06:55:52Z" w16du:dateUtc="2025-12-06T06:55:52Z">
  <w:r w:rsidR="F3F4F4B4">
    <w:t xml:space="preserve"> (interpreting IRPA)</w:t>
  </w:r>
</w:ins>
```

Proposed API generates this automatically from:
```python
doc.insert_tracked(" (interpreting IRPA)", ...)
```

---

## Technical Architecture

### Directory Structure

```
python_docx_redline/
├── __init__.py              # Public API exports
├── document.py              # Main Document class
├── text_operations.py       # insert_tracked, replace_tracked, etc.
├── structural_operations.py # insert_paragraph, delete_section, etc.
├── search/
│   ├── __init__.py
│   ├── text_search.py       # Text finding with fragmentation handling
│   ├── scope.py             # Scope evaluation
│   └── disambiguation.py    # Handling multiple matches
├── models/
│   ├── __init__.py
│   ├── text_span.py         # TextSpan class
│   ├── paragraph.py         # Paragraph wrapper
│   ├── section.py           # Section wrapper
│   └── results.py           # EditResult, ValidationResult, etc.
├── batch/
│   ├── __init__.py
│   ├── processor.py         # Batch edit processing
│   └── file_loader.py       # YAML/JSON parsing
├── xml_generation/
│   ├── __init__.py
│   ├── tracked_changes.py   # Generate tracked change XML
│   └── elements.py          # Generate paragraph/run XML
└── validation/
    ├── __init__.py
    ├── validator.py         # Enhanced validation
    └── errors.py            # Custom exception classes
```

### Core Data Flow

```
User Call
    ↓
Document.insert_tracked(text, after="target", scope="section:Argument")
    ↓
1. Scope Evaluation
   - Parse scope specification
   - Filter paragraphs
    ↓
2. Text Search
   - Find "target" in filtered paragraphs
   - Handle fragmentation across runs
   - Return candidate matches
    ↓
3. Disambiguation
   - If multiple matches, raise AmbiguousTextError
   - If no matches, raise TextNotFoundError with suggestions
    ↓
4. TextSpan Creation
   - Create TextSpan representing found text
   - Record which runs contain it
    ↓
5. XML Generation
   - Generate <w:ins> XML with proper attributes
   - Auto-inject RSID, author, date
    ↓
6. Insertion
   - Determine correct insertion point
   - Insert XML after last run in TextSpan
    ↓
7. Return
   - Return new TextSpan representing inserted text
```

---

## Implementation Phases

### Phase 1: Core Text Operations (MVP)

**Goal**: Replace 80% of current XML-writing code

**Priority 1** (ship this first):
```python
class Document:
    def __init__(self, path, author="Claude", ...):
        # Reuse existing Document class as base
        pass

    def insert_tracked(self, text, after=None, scope=None):
        # Most common operation - implement first
        pass

    def replace_tracked(self, find, replace, scope=None):
        # Second most common
        pass

    def accept_all_changes(self):
        # Already exists - expose at top level
        pass

    def save(self, path=None):
        # Already exists
        pass
```

**Files to create**:
1. `text_operations.py` - Insert/replace/delete logic
2. `search/text_search.py` - Text finding
3. `models/text_span.py` - TextSpan class

**Test with**:
- Client feedback scenario (11 surgical edits)
- Success = complete all 11 edits without writing XML

### Phase 2: Structural Operations

```python
def insert_paragraph(self, text, after=None, track=True):
    pass

def delete_section(self, heading, track=True, update_toc=False):
    pass
```

**Files to create**:
1. `structural_operations.py`
2. `models/paragraph.py`
3. `models/section.py`

### Phase 3: Advanced Features

```python
def apply_edits(self, edits):
    pass

def apply_edit_file(self, path):
    pass

def find_all(self, text, scope=None):
    pass
```

**Files to create**:
1. `batch/processor.py`
2. `batch/file_loader.py`
3. `search/disambiguation.py`

---

## Critical Implementation Details

### 1. Text Search Algorithm

**Problem**: Text may be split across multiple `<w:r>` elements

**Solution**: Build a virtual text stream

**Note**: For an alternative approach using single-character run normalization, see [Eric White's Algorithm](ERIC_WHITE_ALGORITHM.md). Our character map approach is more efficient for read-only searches, while Eric White's approach is better for complex replacements spanning multiple runs with different formatting.

```python
class TextSearch:
    def find_text(self, text: str, paragraphs: list) -> list[TextSpan]:
        """Find text handling fragmentation."""
        results = []

        for para in paragraphs:
            # Get all runs
            runs = list(para.iter(_parse_tag("w:r")))

            # Build character map: char_index -> (run_index, offset_in_run)
            char_map = []
            full_text = []

            for run_idx, run in enumerate(runs):
                run_text = ''.join(run.itertext())
                for char_idx, char in enumerate(run_text):
                    char_map.append((run_idx, char_idx))
                    full_text.append(char)

            full_text = ''.join(full_text)

            # Find all occurrences
            start = 0
            while True:
                pos = full_text.find(text, start)
                if pos == -1:
                    break

                # Map back to runs
                start_run_idx, start_offset = char_map[pos]
                end_run_idx, end_offset = char_map[pos + len(text) - 1]

                # Create TextSpan
                span = TextSpan(
                    runs=runs[start_run_idx:end_run_idx + 1],
                    start_run_index=start_run_idx,
                    end_run_index=end_run_idx,
                    start_offset=start_offset,
                    end_offset=end_offset + 1,  # Exclusive
                    paragraph=para
                )
                results.append(span)

                start = pos + 1

        return results
```

### 2. XML Generation for Tracked Changes

**Key insight**: Most attributes can be auto-generated

```python
class TrackedChangeGenerator:
    def __init__(self, doc):
        self.doc = doc
        self.next_change_id = self._get_max_change_id() + 1

    def create_insertion(self, text: str, author: str = None) -> str:
        """Generate <w:ins> XML."""
        author = author or self.doc.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self.next_change_id
        self.next_change_id += 1

        # Handle special characters
        xml_space = ' xml:space="preserve"' if text[0].isspace() or text[-1].isspace() else ''

        xml = f'''<w:ins w:id="{change_id}" w:author="{author}" w:date="{timestamp}" w16du:dateUtc="{timestamp}">
  <w:r w:rsidR="{self.doc.rsid}">
    <w:t{xml_space}>{self._escape_xml(text)}</w:t>
  </w:r>
</w:ins>'''
        return xml

    def create_deletion(self, text: str, author: str = None) -> str:
        """Generate <w:del> XML."""
        # Similar to insertion but with <w:delText>
        pass

    def _escape_xml(self, text: str) -> str:
        """Escape XML special characters."""
        return (
            text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&apos;')
        )
```

### 3. Scope Evaluation

```python
class ScopeEvaluator:
    @staticmethod
    def parse(scope_spec) -> Callable:
        """Convert scope spec to evaluation function."""
        if scope_spec is None:
            return lambda p: True

        if isinstance(scope_spec, str):
            return ScopeEvaluator._parse_string(scope_spec)

        if isinstance(scope_spec, dict):
            return ScopeEvaluator._parse_dict(scope_spec)

        if callable(scope_spec):
            return scope_spec

        raise ValueError(f"Invalid scope: {scope_spec}")

    @staticmethod
    def _parse_string(s: str) -> Callable:
        """Parse string shortcuts."""
        if s.startswith("section:"):
            section = s[8:]
            return lambda p: p.section == section

        if s.startswith("paragraph_containing:"):
            text = s[21:]
            return lambda p: text in p.text

        # Default: paragraph containing
        return lambda p: s in p.text

    @staticmethod
    def _parse_dict(d: dict) -> Callable:
        """Parse dictionary scope."""
        def evaluator(para):
            # Check 'contains'
            if 'contains' in d and d['contains'] not in para.text:
                return False

            # Check 'section'
            if 'section' in d and para.section != d['section']:
                return False

            # Check 'not_in'
            if 'not_in' in d and para.section in d['not_in']:
                return False

            # Check 'line_range'
            if 'line_range' in d:
                start, end = d['line_range']
                if not (start <= para.line_number < end):
                    return False

            return True

        return evaluator
```

### 4. TextSpan Operations

```python
class TextSpan:
    def insert_after(self, text: str, track: bool = True) -> 'TextSpan':
        """Insert text after this span."""
        if track:
            xml = self.doc._xml_generator.create_insertion(text)
        else:
            xml = f'<w:r><w:t>{self._escape_xml(text)}</w:t></w:r>'

        # Insert after last run in span
        last_run = self.runs[self.end_run_index]
        new_nodes = self.doc.editor.insert_after(last_run, xml)

        # Return new TextSpan for inserted text
        return TextSpan(
            runs=new_nodes,
            start_run_index=0,
            end_run_index=len(new_nodes) - 1,
            start_offset=0,
            end_offset=len(text),
            paragraph=self.paragraph
        )

    def replace(self, text: str, track: bool = True) -> 'TextSpan':
        """Replace this span with new text."""
        if track:
            # Generate deletion + insertion
            deletion_xml = self.doc._xml_generator.create_deletion(self.text)
            insertion_xml = self.doc._xml_generator.create_insertion(text)
            replacement_xml = deletion_xml + insertion_xml
        else:
            replacement_xml = f'<w:r><w:t>{self._escape_xml(text)}</w:t></w:r>'

        # Determine what to replace
        if self.start_run_index == self.end_run_index:
            # Spans single run - split it
            run = self.runs[0]
            run_text = ''.join(run.itertext())

            before_text = run_text[:self.start_offset]
            after_text = run_text[self.end_offset:]

            parts = []
            if before_text:
                parts.append(f'<w:r><w:t>{self._escape_xml(before_text)}</w:t></w:r>')
            parts.append(replacement_xml)
            if after_text:
                parts.append(f'<w:r><w:t>{self._escape_xml(after_text)}</w:t></w:r>')

            new_nodes = self.doc.editor.replace_node(run, ''.join(parts))
        else:
            # Spans multiple runs - replace all
            # Keep runs before first, after last
            # Replace runs in span with deletion + insertion
            pass  # More complex implementation

        return TextSpan(...)  # Return span for new text
```

---

## Error Handling Strategy

### Exception Hierarchy

```python
class DocxRedlineError(Exception):
    """Base exception for all python_docx_redline errors."""
    pass

class TextNotFoundError(DocxRedlineError):
    """Text not found in specified scope."""
    def __init__(self, text, scope, suggestions):
        self.text = text
        self.scope = scope
        self.suggestions = suggestions
        super().__init__(self._format_message())

    def _format_message(self):
        msg = f"Could not find '{self.text}'"
        if self.scope:
            msg += f" in scope '{self.scope}'"

        msg += "\n\nSuggestions:\n"
        for s in self.suggestions:
            msg += f"  • {s}\n"

        return msg

class AmbiguousTextError(DocxRedlineError):
    """Multiple occurrences of text found."""
    def __init__(self, text, matches):
        self.text = text
        self.matches = matches
        super().__init__(self._format_message())

    def _format_message(self):
        msg = f"Found {len(self.matches)} occurrences of '{self.text}'\n\n"
        for i, match in enumerate(self.matches):
            msg += f"{i}: ...{match.context}...\n"
            msg += f"   Section: {match.location.section}, Line: {match.location.line_number}\n\n"
        msg += "Provide a more specific scope to disambiguate."
        return msg

class ValidationError(DocxRedlineError):
    """Document validation failed."""
    pass
```

### Suggestion Generation

```python
class SuggestionGenerator:
    @staticmethod
    def generate_suggestions(text: str, doc) -> list[str]:
        """Generate helpful suggestions when text not found."""
        suggestions = []

        # Try fuzzy match
        fuzzy_results = doc.find_text(text, fuzzy=True)
        if fuzzy_results:
            suggestions.append(f"Found similar text: '{fuzzy_results[0].text}'")
            suggestions.append("Try: fuzzy=True parameter")

        # Search without scope
        all_results = doc.find_all(text)
        if all_results:
            sections = set(r.location.section for r in all_results)
            suggestions.append(f"Found in sections: {', '.join(sections)}")
            suggestions.append(f"Try expanding scope or use: doc.find_all('{text}')")

        # Check for common issues
        if text.count('"') != text.count('"'):
            suggestions.append("Check quote marks - may be curly quotes in document")

        if '  ' in text:
            suggestions.append("Text contains double spaces - may be single space in document")

        return suggestions
```

---

## Testing Strategy

### Unit Test Structure

```python
# tests/test_text_operations.py
import pytest
from python_docx_redline import Document
from python_docx_redline.errors import TextNotFoundError, AmbiguousTextError

class TestInsertTracked:
    def test_simple_insertion(self, sample_doc):
        """Test basic insertion after target text."""
        doc = Document(sample_doc)
        result = doc.insert_tracked(" added", after="target")

        assert isinstance(result, TextSpan)
        assert " added" in doc.get_text()

    def test_insertion_with_scope(self, sample_doc):
        """Test insertion with scope limitation."""
        doc = Document(sample_doc)
        doc.insert_tracked(" added", after="target", scope="section:Argument")

        # Verify not inserted in TOA
        toa_text = doc.get_section_text("Table of Authorities")
        assert " added" not in toa_text

    def test_fragmented_text(self, fragmented_doc):
        """Test insertion when target text is fragmented."""
        doc = Document(fragmented_doc)
        # Document has "(7th Cir. " in one run, "2022)" in another
        doc.insert_tracked(" (interpreting IRPA)", after="(7th Cir. 2022)")

        assert " (interpreting IRPA)" in doc.get_text()

    def test_not_found_error(self, sample_doc):
        """Test error when text not found."""
        doc = Document(sample_doc)
        with pytest.raises(TextNotFoundError) as exc:
            doc.insert_tracked(" added", after="nonexistent")

        # Check error message has suggestions
        assert "Suggestions:" in str(exc.value)

    def test_ambiguous_error(self, sample_doc):
        """Test error when multiple matches."""
        doc = Document(sample_doc)
        with pytest.raises(AmbiguousTextError) as exc:
            doc.insert_tracked(" added", after="common text")

        # Check shows all matches
        assert "Found" in str(exc.value)
        assert "Section:" in str(exc.value)


# tests/conftest.py
@pytest.fixture
def sample_doc(tmp_path):
    """Create a sample document for testing."""
    # Use docx-js or python-docx to create test document
    doc_path = tmp_path / "test.docx"
    # ... create document with known content ...
    return str(doc_path)

@pytest.fixture
def fragmented_doc(tmp_path):
    """Create document with fragmented text."""
    # Manually create document with text split across runs
    pass
```

### Integration Tests

```python
# tests/integration/test_surgical_edits.py
def test_client_feedback_scenario():
    """Test complete surgical edits workflow from real use case."""
    # Use actual client feedback document
    doc = Document('tests/fixtures/client_comments.docx', author="Hancock, Parker")

    # Accept all changes
    result = doc.accept_all_changes()
    assert result.insertions > 0
    assert result.deletions > 0

    # Apply surgical edits
    edits = [
        {
            "type": "insert_tracked",
            "text": " (interpreting IRPA)",
            "after": "(7th Cir. 2022)",
            "scope": "Huston"
        },
        {
            "type": "insert_tracked",
            "text": " (granting motion to dismiss)",
            "after": "(N.D. Ill. 2016)",
            "scope": "Vrdolyak"
        },
        {
            "type": "replace_tracked",
            "find": "records their property ownership",
            "replace": "compiles their property ownership data",
            "scope": "It claims merely"
        },
        # ... all 11 edits ...
    ]

    results = doc.apply_edits(edits)

    # Verify all succeeded
    assert all(r.success for r in results), \
        f"Failed edits: {[r for r in results if not r.success]}"

    # Save and reopen to verify persistence
    output = 'tests/output/final.docx'
    doc.save(output)

    doc2 = Document(output)
    changes = doc2.list_tracked_changes()

    # All changes attributed to correct author
    assert all(c.author == "Hancock, Parker" for c in changes)

    # Verify specific edits present
    text = doc2.get_text()
    assert "(interpreting IRPA)" in text
    assert "(granting motion to dismiss)" in text
    assert "compiles their property ownership data" in text
```

### Performance Tests

```python
# tests/performance/test_benchmarks.py
import time

def test_insertion_performance():
    """Insertions should complete in <100ms."""
    doc = Document('large_doc.docx')

    start = time.time()
    doc.insert_tracked(" text", after="target")
    elapsed = time.time() - start

    assert elapsed < 0.1, f"Insertion took {elapsed:.3f}s, expected <0.1s"

def test_batch_performance():
    """10 edits should complete in <2s."""
    doc = Document('large_doc.docx')

    edits = [
        {"type": "insert_tracked", "text": f" {i}", "after": f"target{i}"}
        for i in range(10)
    ]

    start = time.time()
    doc.apply_edits(edits)
    elapsed = time.time() - start

    assert elapsed < 2.0, f"Batch took {elapsed:.3f}s, expected <2s"
```

---

## Backward Compatibility

### Keep Existing API Available

```python
class Document:
    # NEW: High-level API
    def insert_tracked(self, text, after=None, ...):
        pass

    # OLD: Low-level API (still works)
    def __getitem__(self, key):
        """Access XML files directly."""
        return self.get_editor(key)

    def get_editor(self, xml_path):
        """Get low-level XML editor (backward compatible)."""
        # Return existing DocxXMLEditor
        pass

# Both work
doc = Document('file.docx')

# New way
doc.insert_tracked(" text", after="target")

# Old way (still supported)
editor = doc["word/document.xml"]
node = editor.get_node(tag="w:r", contains="target")
editor.insert_after(node, '<w:ins>...</w:ins>')
```

---

## Dependencies

### Required

- `lxml` (already used) - XML parsing
- `python-dateutil` (already in environment) - Date handling

### Optional

- `pyyaml` - For YAML edit files (Phase 3)
- `rapidfuzz` - For fuzzy text matching (Phase 3)

### Development

- `pytest` - Testing
- `pytest-cov` - Coverage
- `black` - Formatting
- `mypy` - Type checking
- `ruff` - Linting

---

## Open Questions

1. **Fuzzy matching**: Use existing library (rapidfuzz) or implement simple Levenshtein?
2. **TOC updates**: Parse TOC fields or require Word to regenerate?
3. **Style preservation**: Copy full `<w:rPr>` from original runs or selective copy?
4. **Async support**: Support async operations for large documents?
5. **Progress callbacks**: Provide progress updates for batch operations?

---

## Success Metrics

Track these to measure success:

1. **Lines of code reduction**:
   - Target: 10x reduction for common operations
   - Measure: Compare before/after for client feedback scenario

2. **Time to implement**:
   - Target: <5 minutes to implement 11 surgical edits
   - Measure: Time from API call to saved document

3. **Error rate**:
   - Target: <5% validation errors
   - Measure: % of operations that pass validation

4. **Learning curve**:
   - Target: New user can apply edits in <30 minutes
   - Measure: Time for new developer to complete tutorial

5. **Test coverage**:
   - Target: >90% coverage
   - Measure: pytest-cov report

---

## Resources

### Helpful OOXML References

- Word Open XML specification: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/
- Tracked changes spec: Part 1, Section 17.13.5
- Run properties: Part 1, Section 17.3.2

### Similar Projects (for inspiration)

- python-docx: https://github.com/python-openxml/python-docx
- Aspose.Words API: https://docs.aspose.com/words/python-net/
- mammoth.js (JavaScript): https://github.com/mwilliamson/mammoth.js

### Testing Documents

Create test fixtures:
1. Simple document (1 page, basic formatting)
2. Fragmented text (deliberately split text across runs)
3. Complex document (TOC, multiple sections, footnotes)
4. Client feedback document (real-world scenario)

---

**End of Implementation Notes**
