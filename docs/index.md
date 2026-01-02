# python-docx-redline

High-level Python API for editing Word documents with tracked changes.

## Quick Example

```python
from python_docx_redline import Document

doc = Document("contract.docx")
doc.replace_tracked("the Contractor", "the Service Provider")
doc.save("contract_edited.docx")
```

## Documentation

### Getting Started

- **[Getting Started](getting-started.md)** — Installation and basic concepts

### User Guide

- **[Basic Operations](user-guide/basic-operations.md)** — Insert, delete, replace text
- **[Structural Operations](user-guide/structural-ops.md)** — Paragraphs and sections
- **[Viewing Content](user-guide/viewing-content.md)** — Read documents before editing
- **[Batch Operations](user-guide/batch-operations.md)** — YAML configuration files
- **[Formatting](user-guide/formatting.md)** — Markdown syntax and format tracking
- **[Advanced Features](user-guide/advanced.md)** — Scopes, MS365, rendering, images

### Reference

- **[Quick Reference](QUICK_REFERENCE.md)** — Cheat sheet
- **[API Reference](PROPOSED_API.md)** — Complete method documentation
- **[find_all() API](API_FIND_ALL.md)** — Finding text with context
- **[Context Awareness](CONTEXT_AWARENESS_GUIDE.md)** — Preview and fragment detection
- **[Text Search Algorithm](ERIC_WHITE_ALGORITHM.md)** — How text search works

## Links

- [GitHub Repository](https://github.com/SanctionedCodeList/python-docx-redline)
- [PyPI Package](https://pypi.org/project/python-docx-redline/)
