# Advanced Features

## Scopes

Limit operations to specific parts of the document.

### String Scopes

```python
# Paragraphs containing specific text
doc.replace_tracked("Client", "Customer", scope="payment terms")
```

### Section Scopes

```python
# Only modify within a specific section
doc.insert_tracked(
    " (hereinafter 'Agreement')",
    after="this Agreement",
    scope="section:Introduction"
)
```

### Dictionary Scopes

```python
doc.insert_tracked(
    " (Non-Disclosure Agreement)",
    after="NDA",
    scope={
        "contains": "confidential",
        "not_contains": "public",
        "section": "Definitions"
    }
)
```

### Callable Scopes

```python
# Custom filtering logic
doc.replace_tracked(
    "old",
    "new",
    scope=lambda p: len(''.join(p.itertext())) > 100
)
```

## MS365 Identity Integration

Link tracked changes to real MS365/Office365 users:

```python
from python_docx_redline import Document, AuthorIdentity

# Create identity with MS365 profile info
identity = AuthorIdentity(
    author="Hancock, Parker",
    email="parker.hancock@company.com",
    provider_id="AD",  # Active Directory
    guid="c5c513d2-1f51-4d69-ae91-17e5787f9bfc"
)

# Use identity when creating document
doc = Document("contract.docx", author=identity)

# All tracked changes include full MS365 identity
doc.insert_tracked(" (amended)", after="Section 1")
doc.replace_tracked("30 days", "45 days")

doc.save("contract_edited.docx")
# Changes appear in Word with full user profile and avatar
```

**Finding existing identity info:**

1. Unpack an existing .docx file (`unzip document.docx`)
2. Inspect `word/people.xml` for author information
3. Look for `w15:userId` (GUID) and `w15:providerId` attributes

## python-docx Integration

Seamlessly convert between libraries:

```python
from docx import Document as PythonDocxDocument
from python_docx_redline import Document, from_python_docx, to_python_docx

# Create document with python-docx
py_doc = PythonDocxDocument()
py_doc.add_heading("Contract", 0)
py_doc.add_paragraph("Payment terms: net 30 days")

# Convert for tracked changes
doc = from_python_docx(py_doc, author="Legal Team")
doc.replace_tracked("net 30 days", "net 45 days")
doc.save("contract_redlined.docx")

# Convert back to python-docx
py_doc = to_python_docx(doc, validate=False)
py_doc.add_paragraph("Added with python-docx")
py_doc.save("final.docx")
```

## In-Memory Workflows

Work without filesystem access:

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

## Document Rendering

Render documents to PNG images for visual inspection:

```python
from python_docx_redline import Document
from python_docx_redline.rendering import is_rendering_available

if is_rendering_available():
    doc = Document("contract.docx")

    images = doc.render_to_images(
        output_dir="./images",
        dpi=150,
        prefix="page"
    )

    for img in images:
        print(f"Generated: {img}")
```

**Requirements:**

```bash
# macOS
brew install --cask libreoffice
brew install poppler

# Linux (Ubuntu/Debian)
sudo apt install libreoffice poppler-utils
```

## Image Insertion

Insert images with optional tracked changes:

```python
# Basic image insertion
doc.insert_image("logo.png", after="Company Name:")

# With specific dimensions
doc.insert_image(
    "chart.png",
    after="Figure 1:",
    width_inches=4.0,
    height_inches=3.0
)

# Tracked insertion (appears in Word's review pane)
doc.insert_image_tracked(
    "signature.png",
    after="Authorized By:",
    author="Legal Team"
)

# With alt text for accessibility
doc.insert_image(
    "diagram.png",
    after="See diagram:",
    description="Network architecture diagram",
    width_cm=10.0
)
```

**Supported formats:** PNG, JPEG, GIF, BMP, TIFF, WebP

**Dimension behavior:**

- Neither specified: uses native image dimensions
- One specified: other calculated to maintain aspect ratio
- PIL/Pillow not installed: defaults to 2x2 inches

## Context Manager

Automatic cleanup with context managers:

```python
with Document("contract.docx") as doc:
    doc.insert_tracked("new clause", after="Section 1")
    doc.save("modified.docx")
# Automatic cleanup
```

## Next Steps

- [API Reference](../PROPOSED_API.md) — Complete method documentation
- [Quick Reference](../QUICK_REFERENCE.md) — Cheat sheet
- [Text Search Algorithm](../ERIC_WHITE_ALGORITHM.md) — How search works internally
