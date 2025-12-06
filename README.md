# docx_redline

A high-level Python API for editing Word documents with tracked changes.

## Status

🚧 **Under Active Development** - Phase 1 MVP in progress

## Goal

Reduce surgical Word document edits from 30+ lines of raw OOXML XML manipulation to 3 lines of high-level API calls.

## Quick Example

```python
from docx_redline import Document

doc = Document("contract.docx")
doc.insert_tracked("new clause text", after="Section 2.1")
doc.save("contract_edited.docx")
```

## Requirements

- Python 3.10+
- lxml
- python-dateutil
- pyyaml

## Installation

```bash
pip install docx-redline
```

## Documentation

See the `docs/` directory for complete API documentation:
- [API Reference](docs/PROPOSED_API.md)
- [Implementation Notes](docs/IMPLEMENTATION_NOTES.md)
- [Quick Reference](docs/QUICK_REFERENCE.md)

## Development

```bash
# Clone the repository
git clone https://github.com/parkerhancock/docx_redline.git
cd docx_redline

# Install in development mode with dev dependencies
pip install -e ".[dev]"

# Run tests
pytest
```

## License

MIT
