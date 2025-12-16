# python-docx-redline Documentation

High-level Python API for editing Word documents with tracked changes.

## User Guides

- **[Quick Reference](QUICK_REFERENCE.md)** — Cheat sheet for common operations
- **[Context-Awareness Guide](CONTEXT_AWARENESS_GUIDE.md)** — Preview and fragment detection
- **[find_all() API](API_FIND_ALL.md)** — Finding text with context

## API Reference

- **[Complete API Reference](PROPOSED_API.md)** — Full documentation of all methods

## Technical Reference

- **[Eric White's Algorithm](ERIC_WHITE_ALGORITHM.md)** — How text search handles fragmented XML runs

## Getting Started

```python
from python_docx_redline import Document

doc = Document("contract.docx")
doc.replace_tracked("old text", "new text")
doc.save("contract_edited.docx")
```

See the [README](../README.md) for installation and more examples.
