---
name: docx-python
description: "Python library for Word document manipulation. Use python-docx for creating new documents, python-docx-redline for ALL editing tasks (handles run fragmentation, optional tracked changes). Covers text extraction, find/replace, comments, footnotes, hyperlinks, styles, CriticMarkup, and raw OOXML."
---

# Python DOCX Libraries

Two libraries for Word document manipulation:

- **python-docx**: Creating new documents from scratch
- **python-docx-redline**: Editing existing documents (recommended for ALL editing)

## Why python-docx-redline for Editing?

python-docx-redline handles **run fragmentation** that breaks python-docx find/replace. Word splits text across XML runs unpredictably—"Hello World" might be `<w:r>Hel</w:r><w:r>lo Wor</w:r><w:r>ld</w:r>`. python-docx-redline finds and edits text regardless of fragmentation.

## Quick Start

```python
# Creating new documents
from docx import Document
doc = Document()
doc.add_heading("Title", 0)
doc.add_paragraph("Content here.")
doc.save("new.docx")

# Editing existing documents (silent or tracked)
from python_docx_redline import Document
doc = Document("existing.docx")
doc.replace("OLD", "new")                    # Silent edit
doc.replace("30 days", "45 days", track=True)  # Tracked change
doc.save("modified.docx")
```

## Decision Tree

| Task | Guide |
|------|-------|
| **Create new document** | [creation.md](./creation.md) |
| **Generate from data/template** | [templating.md](./templating.md) |
| **Extract text** | [reading.md](./reading.md) |
| **Edit existing document** | [editing.md](./editing.md) |
| **Edit with tracked changes** | [tracked-changes.md](./tracked-changes.md) |
| **Add comments** | [comments.md](./comments.md) |
| **Footnotes/endnotes** | [footnotes.md](./footnotes.md) |
| **Hyperlinks** | [hyperlinks.md](./hyperlinks.md) |
| **Style management** | [styles.md](./styles.md) |
| **CriticMarkup round-trip** | [criticmarkup.md](./criticmarkup.md) |
| **Use both libraries together** | [integration.md](./integration.md) |
| **Structured YAML view (refs)** | [accessibility.md](./accessibility.md) |
| **Raw OOXML manipulation** | [ooxml.md](./ooxml.md) |

## Real-World Examples

### Example 1: Populate an NDA Template

**Scenario:** Fill in a standard NDA with client details.

```python
from python_docx_redline import Document

doc = Document("nda_template.docx")

# Replace all placeholders (silent - no tracking)
doc.replace("{{PARTY_NAME}}", "Acme Corporation")
doc.replace("{{EFFECTIVE_DATE}}", "January 15, 2025")
doc.replace("{{GOVERNING_LAW}}", "State of Delaware")
doc.replace("{{TERM_YEARS}}", "3")
doc.replace("{{SIGNATURE_BLOCK}}", "_____________________", occurrence="all")

doc.save("acme_nda_final.docx")
```

**Before:** `This Agreement is entered into by {{PARTY_NAME}} as of {{EFFECTIVE_DATE}}...`

**After:** `This Agreement is entered into by Acme Corporation as of January 15, 2025...`

### Example 2: Redline a Contract

**Scenario:** Negotiate payment terms in a vendor agreement.

```python
from python_docx_redline import Document

doc = Document("vendor_agreement.docx")

# These show as tracked changes in Word
doc.replace("Net 30", "Net 45", track=True)
doc.replace("2% late fee", "1.5% late fee", track=True)
doc.delete("Vendor may terminate for convenience with 30 days notice.", track=True)
doc.insert(" Payment terms are non-negotiable after execution.",
           after="Payment Schedule", track=True)

doc.save("vendor_agreement_redlined.docx")
```

**Result in Word:** Changes appear with strikethrough (deletions) and underline (insertions), reviewable via Track Changes.

### Example 3: Clean Up a Document for Publishing

**Scenario:** Remove internal notes and standardize formatting before publishing.

```python
from python_docx_redline import Document

doc = Document("report_draft.docx")

# Remove all internal markers (silent)
doc.delete("TODO: ", occurrence="all")
doc.delete("INTERNAL NOTE: ", occurrence="all")
doc.delete("[DRAFT]", occurrence="all")

# Standardize terminology
doc.replace("client", "Client", occurrence="all")  # Capitalize consistently
doc.replace("  ", " ", occurrence="all")  # Remove double spaces

# Update header
doc.replace("DRAFT", "FINAL", scope="headers")

doc.save("report_final.docx")
```

### Example 4: Batch Edits from a List

**Scenario:** Apply multiple edits defined in a configuration.

```python
from python_docx_redline import Document

doc = Document("lease_agreement.docx")

edits = [
    # Template fields (silent)
    {"type": "replace", "find": "{{TENANT}}", "replace": "Jane Smith"},
    {"type": "replace", "find": "{{LANDLORD}}", "replace": "123 Properties LLC"},
    {"type": "replace", "find": "{{ADDRESS}}", "replace": "456 Oak Street, Apt 2B"},
    {"type": "replace", "find": "{{RENT}}", "replace": "$2,500"},

    # Negotiated terms (tracked)
    {"type": "replace", "find": "12-month term", "replace": "18-month term", "track": True},
    {"type": "delete", "text": "No pets allowed.", "track": True},
    {"type": "insert", "text": " Tenant may keep one cat.", "after": "Pet Policy:", "track": True},
]

results = doc.apply_edits(edits)
print(f"Applied {sum(1 for r in results if r.success)}/{len(edits)} edits")

doc.save("lease_jane_smith.docx")
```

### Find Before Editing

```python
matches = doc.find_all("payment")
for m in matches:
    print(f"{m.index}: {m.context}")

# Target specific occurrence
doc.replace("payment", "Payment", occurrence=2, track=True)
```

### Scoped Edits

```python
doc.replace("Client", "Customer", scope="section:Payment Terms", track=True)
```

### Batch Operations

```python
edits = [
    {"type": "replace", "find": "{{NAME}}", "replace": "John"},
    {"type": "replace", "find": "old", "replace": "new", "track": True},
    {"type": "delete", "text": "DRAFT"},
]
doc.apply_edits(edits, default_track=False)
```

## For LLM/Agent Workflows

Use the **AccessibilityTree** for structured YAML output with stable refs:

```python
from python_docx_redline import Document
doc = Document("contract.docx")
tree = doc.accessibility_tree()
print(tree.to_yaml())  # Structured view with refs like p:5, tbl:0/row:1/cell:2
```

Then edit by ref for precision:

```python
doc.replace_by_ref("p:5", "New paragraph text", track=True)
doc.delete_by_ref("p:10", track=True)
```

See [accessibility.md](./accessibility.md) for full details.

## Visual Layout Analysis

To analyze physical document layout, formatting, and visual elements with AI vision:

```python
from scripts.docx_to_images import docx_to_images

# Convert document pages to images
images = docx_to_images("contract.docx")
# Returns: [Path('contract_page_1.png'), Path('contract_page_2.png'), ...]

# Higher resolution for detailed analysis
images = docx_to_images("report.docx", output_dir="./pages", dpi=300)
```

Requires LibreOffice and poppler. Run `../install.sh` or see [scripts/README.md](./scripts/README.md).

## Reference Files

| File | Purpose |
|------|---------|
| [creation.md](./creation.md) | Creating documents with style templates |
| [templating.md](./templating.md) | DocxBuilder: generate from data with markdown |
| [reading.md](./reading.md) | Text extraction, find_all(), document structure |
| [editing.md](./editing.md) | All editing operations (tracked and untracked) |
| [tracked-changes.md](./tracked-changes.md) | Tracked changes: insert/delete/replace, scopes, batch |
| [comments.md](./comments.md) | Adding comments, replies, resolution |
| [footnotes.md](./footnotes.md) | Footnotes/endnotes: CRUD, tracked changes, search |
| [hyperlinks.md](./hyperlinks.md) | Insert, edit, remove hyperlinks |
| [styles.md](./styles.md) | Style management and formatting |
| [criticmarkup.md](./criticmarkup.md) | Export/import tracked changes as markdown |
| [integration.md](./integration.md) | python-docx integration workflows |
| [accessibility.md](./accessibility.md) | AccessibilityTree for agent workflows |
| [ooxml.md](./ooxml.md) | Raw XML manipulation for complex scenarios |
