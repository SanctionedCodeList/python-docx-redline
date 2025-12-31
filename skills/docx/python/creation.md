# Creating New Documents

## Style Templates

python-docx cannot create custom styles programmatically. Styles must exist in the starting document. Use a style template:

1. Create styles in Word (format some text, modify styles as needed)
2. Delete all content, leaving an empty document with style definitions
3. Save as your style template (e.g., `styles/corporate.docx`)

### Why This Matters

```python
# Without style template - only default Word styles available
doc = Document()
doc.add_heading("Title", 0)  # Uses default Heading style

# With style template - your custom styles are available
doc = Document("styles/corporate.docx")
doc.add_heading("Title", 0)  # Uses your styled Heading
doc.add_paragraph("Content", style="Executive Summary")  # Custom style works
```

## Generation Pattern

```python
from docx import Document

# Load style template (not empty Document())
doc = Document("styles/corporate.docx")

# Add content using styles defined in template
doc.add_heading("Title", 0)
doc.add_heading("Section heading states the conclusion", 1)
doc.add_paragraph("Supporting content here.")

doc.save("output.docx")
```

## Reusable Generation Functions

For documents you generate repeatedly, create a Python function:

```python
from docx import Document

def create_memo(style_path: str, title: str, sections: list[tuple[str, str]]) -> Document:
    """Generate a memo document.

    Args:
        style_path: Path to style template .docx
        title: Document title
        sections: List of (heading, body) tuples

    Returns:
        Document ready to save
    """
    doc = Document(style_path)
    doc.add_heading(title, 0)
    for heading, body in sections:
        doc.add_heading(heading, 1)
        doc.add_paragraph(body)
    return doc

# Usage
doc = create_memo("styles/memo.docx", "Q4 Update", [
    ("Revenue exceeded targets", "Enterprise segment drove 12% growth..."),
    ("Costs remain controlled", "Operating expenses flat YoY..."),
])
doc.save("q4_update.docx")
```

## Creating Style Templates

In Microsoft Word:

1. **Create a new document**
2. **Define your styles** (Home → Styles → Modify Style or Create New Style)
   - Heading 1, Heading 2, etc.
   - Normal paragraph style
   - Any custom styles you need (e.g., "Executive Summary", "Caption")
3. **Apply each style to some dummy text** (required for style to be saved)
4. **Delete all content** (select all, delete)
5. **Save as .docx** (e.g., `styles/corporate.docx`)

The saved file contains style definitions but no content.

## Style Name Reference

python-docx uses English style names regardless of Word's UI language. Common built-in styles:

| Style Name | Usage |
|------------|-------|
| `Title` | Document title (Heading level 0) |
| `Heading 1` - `Heading 9` | Section headings |
| `Normal` | Body text |
| `Quote` | Block quotes |
| `List Bullet` | Bulleted lists |
| `List Number` | Numbered lists |
| `Caption` | Figure/table captions |

For custom styles, use whatever names you define in Word.

## Project Structure

Style templates and generation scripts live in your project:

```
your-project/
  styles/
    corporate.docx      # Style template
    legal.docx          # Another style template
  scripts/
    generate_memo.py    # Generation function
    generate_report.py
  output/
    memo.docx           # Generated documents
```

The skill does not provide style templates - create them for your project's needs.
