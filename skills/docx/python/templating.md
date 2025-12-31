# Document Templating

Create Word documents from structured data using `DocxBuilder`.

## Quick Start

```python
from python_docx_redline import DocxBuilder

doc = DocxBuilder()
doc.heading("Report Title")
doc.markdown("**Bold** and *italic* work in markdown.")
doc.table(["Name", "Value"], [["Alpha", "1"], ["Beta", "2"]])
doc.save("report.docx")
```

## Installation

```bash
pip install python-docx-redline[templating]
```

## API Reference

### Constructor

```python
DocxBuilder(
    landscape=False,    # True for landscape orientation
    font="Times New Roman",
    font_size=11,      # Points
    margins=1.0        # Inches (all sides)
)
```

### Methods

| Method | Description |
|--------|-------------|
| `heading(text, level=1)` | Add heading (0=Title, 1=H1, 2=H2, etc.) |
| `paragraph(text)` | Add plain text paragraph |
| `markdown(text)` | Add markdown content (headings, lists, bold, italic, code, tables) |
| `table(headers, rows)` | Add table from header list + row lists |
| `table_from(items, columns)` | Auto-extract table from dicts/objects |
| `page_break()` | Insert page break |
| `save(path)` | Save document, returns Path |

All methods return `self` for chaining.

### Advanced

```python
doc.document  # Access underlying python-docx Document
doc.markdown_cell(cell, text)  # Render markdown into table cell
```

## Template Pattern

Define data structure, then render function:

```python
from dataclasses import dataclass
from python_docx_redline import DocxBuilder

@dataclass
class ReportData:
    title: str
    summary: str  # Markdown
    items: list[dict]

def render_report(data: ReportData, path: str):
    doc = DocxBuilder()
    doc.heading(data.title)
    doc.markdown(data.summary)
    doc.table_from(data.items, ["name", "value"])
    return doc.save(path)

# Usage
data = ReportData(
    title="Q4 Report",
    summary="Revenue **exceeded** targets.",
    items=[{"name": "Sales", "value": "1.2M"}]
)
render_report(data, "q4_report.docx")
```

## Markdown Support

Full markdown syntax in `markdown()`:

```python
doc.markdown("""
## Section Title

This is **bold** and *italic*.

- Bullet point
- Another point

1. Numbered item
2. Another item

> Blockquote

`inline code`

| Column A | Column B |
|----------|----------|
| Value 1  | Value 2  |
""")
```

## table_from() with Objects

Works with dicts, dataclasses, or Pydantic models:

```python
@dataclass
class LineItem:
    description: str
    quantity: int
    unit_price: float

items = [
    LineItem("Widget", 10, 5.00),
    LineItem("Gadget", 3, 15.00),
]

# Auto-generates headers from column names
doc.table_from(items, ["description", "quantity", "unit_price"])
# Headers: Description | Quantity | Unit Price

# Custom headers
doc.table_from(items, ["description", "unit_price"],
               headers=["Item", "Price"])
```
