# Getting Started

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

**Optional dependencies:**

```bash
# For fuzzy text matching
pip install python-docx-redline[fuzzy]

# For document rendering to images
brew install --cask libreoffice  # macOS
brew install poppler
```

## Quick Start

```python
from python_docx_redline import Document

# Load a document
doc = Document("contract.docx")

# Insert text with tracked changes
doc.insert_tracked(" (as amended)", after="Section 2.1")

# Replace text with tracked changes
doc.replace_tracked("the Contractor", "the Service Provider")

# Delete text with tracked changes
doc.delete_tracked("for any reason")

# Save the modified document
doc.save("contract_edited.docx")
```

## Core Concepts

### Tracked Changes

All editing operations create tracked changes that appear in Word's review pane:

- **Insertions** — New text appears underlined with the author's name
- **Deletions** — Removed text appears with strikethrough
- **Replacements** — Combines deletion of old text and insertion of new text

### Smart Text Search

The library handles a common Word challenge: text that appears continuous to users is often fragmented across multiple XML elements. For example, "Section 2.1" might be stored as three separate runs: "Section", " ", "2.1".

python-docx-redline automatically finds text regardless of how it's fragmented.

### Authors

Every tracked change records who made it:

```python
# Simple string author
doc = Document("contract.docx", author="Legal Team")

# Or specify per-operation
doc.insert_tracked("new text", after="anchor", author="Parker Hancock")
```

## Next Steps

- [Basic Operations](user-guide/basic-operations.md) — Insert, delete, replace text
- [Viewing Content](user-guide/viewing-content.md) — Read documents before editing
- [Batch Operations](user-guide/batch-operations.md) — Apply multiple edits from YAML
- [API Reference](PROPOSED_API.md) — Complete method documentation
