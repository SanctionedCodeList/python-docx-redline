# DOCX Scripts

These scripts provide low-level OOXML manipulation capabilities from the [Anthropic Skills](https://github.com/anthropics/skills) repository. They are included for advanced scenarios not covered by python-docx-redline.

## When to Use These Scripts

**For most use cases, prefer python-docx-redline** - it provides a high-level API for tracked changes.

Use these scripts only for:
- Adding comments with tracked changes
- Modifying another author's tracked changes
- Inserting images with tracked changes
- Complex nested revision scenarios

## Dependencies

```bash
pip install defusedxml lxml
```

## Schema Files

The OOXML XSD schemas are included in `ooxml/schemas/` (~1 MB):
- `ISO-IEC29500-4_2016/` - ISO/IEC 29500-4:2016 schemas
- `ecma/` - ECMA-376 schemas
- `microsoft/` - Microsoft Office extension schemas
- `mce/` - Markup Compatibility and Extensibility

These schemas are used by the validation scripts for XSD validation.

Alternatively, use LibreOffice for quick validation:
```bash
soffice --headless --convert-to html:HTML document.docx
# If conversion succeeds, the document is valid
```

## File Structure

```
scripts/
├── __init__.py
├── document.py      # Document class for comments and tracked changes
├── utilities.py     # XMLEditor for DOM manipulation
├── templates/       # XML templates for new comment files
│   ├── comments.xml
│   ├── commentsExtended.xml
│   ├── commentsExtensible.xml
│   ├── commentsIds.xml
│   └── people.xml
└── README.md

ooxml/
├── __init__.py
├── scripts/
│   ├── __init__.py
│   ├── pack.py      # Pack directory into .docx
│   ├── unpack.py    # Unpack .docx for editing
│   ├── validate.py  # Command-line validation
│   └── validation/  # Validation classes
│       ├── __init__.py
│       ├── base.py
│       ├── docx.py
│       └── redlining.py
└── schemas/         # OOXML XSD schemas (~1 MB)
    ├── ISO-IEC29500-4_2016/
    ├── ecma/
    ├── microsoft/
    └── mce/
```

## Basic Usage

```python
from skills.docx.scripts.document import Document

# Initialize with unpacked directory
doc = Document('workspace/unpacked', author="Claude")

# Find and manipulate nodes
node = doc["word/document.xml"].get_node(tag="w:del", attrs={"w:id": "1"})

# Add comments
doc.add_comment(start=node, end=node, text="Comment text")

# Save changes
doc.save()
```

For more details, see `ooxml.md` in the parent directory.
