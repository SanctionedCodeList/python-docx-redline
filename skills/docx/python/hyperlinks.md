# Hyperlinks

Use **python-docx-redline** for hyperlink operations including inserting, editing, and removing hyperlinks in document body, headers/footers, and footnotes/endnotes.

## Overview

| Operation | Method | Tracked Support |
|-----------|--------|-----------------|
| **Insert** external hyperlink | `insert_hyperlink(url=..., text=..., after=...)` | No (insertion point only) |
| **Insert** internal hyperlink | `insert_hyperlink(anchor=..., text=..., after=...)` | No (insertion point only) |
| **Insert** in header/footer | `insert_hyperlink_in_header(...)` | No |
| **Insert** in footnote/endnote | `insert_hyperlink_in_footnote(...)` | No |
| **Edit** hyperlink URL | `edit_hyperlink_url(ref, url)` | No |
| **Edit** hyperlink text | `edit_hyperlink_text(ref, text)` | Yes (optional) |
| **Edit** hyperlink anchor | `edit_hyperlink_anchor(ref, anchor)` | No |
| **Remove** hyperlink | `remove_hyperlink(ref, keep_text=True)` | No |
| **List** hyperlinks | `doc.hyperlinks` | N/A |

## Inserting External Hyperlinks

### Basic URL Hyperlink

```python
from python_docx_redline import Document

doc = Document("contract.docx")

# Insert hyperlink after anchor text
doc.insert_hyperlink(
    url="https://example.com",
    text="Click here",
    after="more info",
)

# Insert hyperlink before anchor text
doc.insert_hyperlink(
    url="https://example.com/terms",
    text="Terms of Service",
    before="agreement details",
)

doc.save("contract_with_links.docx")
```

### With Custom Formatting

```python
# Hyperlink with explicit style
doc.insert_hyperlink(
    url="https://example.com",
    text="Visit our website",
    after="contact us",
    style="Hyperlink",  # Uses document's Hyperlink style
)

# Hyperlink with tooltip
doc.insert_hyperlink(
    url="https://example.com/docs",
    text="Documentation",
    after="refer to",
    tooltip="Opens in new window",
)
```

### Handling Multiple Occurrences

```python
# Target specific occurrence of anchor text
doc.insert_hyperlink(
    url="https://example.com",
    text="link",
    after="see",
    occurrence=2,  # After second "see"
)

# Insert after all occurrences
results = doc.insert_hyperlink(
    url="https://example.com",
    text="[link]",
    after="reference",
    occurrence="all",
)
print(f"Inserted {len(results)} hyperlinks")
```

## Inserting Internal Hyperlinks (Bookmarks)

Internal hyperlinks navigate to bookmarks within the document.

### Link to Bookmark

```python
# Create a bookmark first (if not exists)
doc.add_bookmark("DefinitionsSection", on="1. Definitions")

# Insert hyperlink to the bookmark
doc.insert_hyperlink(
    anchor="DefinitionsSection",
    text="See Definitions",
    after="as defined",
)
```

### Link to Heading

```python
# Word automatically creates bookmarks for headings
# Format: _Toc followed by internal ID, or use heading text

# Link using bookmark name
doc.insert_hyperlink(
    anchor="_Ref123456789",  # Word's internal bookmark
    text="Section 2.1",
    after="refer to",
)

# Or find/create bookmark on heading first
doc.add_bookmark("PaymentTerms", on="Payment Terms")
doc.insert_hyperlink(
    anchor="PaymentTerms",
    text="Payment Terms section",
    after="see the",
)
```

## Hyperlinks in Headers/Footers

### Insert in Header

```python
# Insert in default header
doc.insert_hyperlink_in_header(
    url="https://company.com",
    text="Company Website",
    after="Visit",
    header_type="default",
)

# Insert in first page header
doc.insert_hyperlink_in_header(
    url="https://company.com/legal",
    text="Legal Notice",
    after="See",
    header_type="first",
)

# Insert in odd/even headers
doc.insert_hyperlink_in_header(
    url="https://company.com",
    text="Home",
    after="Return to",
    header_type="odd",
)
```

### Insert in Footer

```python
# Insert in default footer
doc.insert_hyperlink_in_footer(
    url="https://company.com/contact",
    text="Contact Us",
    after="Questions?",
    footer_type="default",
)

# Internal link in footer
doc.insert_hyperlink_in_footer(
    anchor="TableOfContents",
    text="Back to Contents",
    after="Navigation:",
    footer_type="default",
)
```

### Header/Footer Types

| Type | Description |
|------|-------------|
| `"default"` | Standard header/footer for most pages |
| `"first"` | First page only (if different first page enabled) |
| `"odd"` | Odd pages (if different odd/even enabled) |
| `"even"` | Even pages (if different odd/even enabled) |

## Hyperlinks in Footnotes/Endnotes

### Insert in Footnote

```python
# Insert hyperlink inside an existing footnote
doc.insert_hyperlink_in_footnote(
    note_id=1,
    url="https://example.com/source",
    text="source document",
    after="See the",
)

# Internal link in footnote
doc.insert_hyperlink_in_footnote(
    note_id=2,
    anchor="Appendix_A",
    text="Appendix A",
    after="Refer to",
)
```

### Insert in Endnote

```python
# Insert hyperlink inside an existing endnote
doc.insert_hyperlink_in_endnote(
    note_id=1,
    url="https://example.com/reference",
    text="full reference",
    after="See",
)
```

## Reading/Listing Hyperlinks

### Iterate All Hyperlinks

```python
# Access all hyperlinks in document body
for link in doc.hyperlinks:
    print(f"Text: {link.text}")
    print(f"Target: {link.target}")
    print(f"Type: {link.link_type}")  # "external" or "internal"
    print(f"Ref: {link.ref}")  # Use for editing/removing
    print()
```

### Hyperlink Properties

```python
link = doc.hyperlinks[0]

# Common properties
link.text           # Display text
link.target         # URL or bookmark name
link.link_type      # "external" or "internal"
link.ref            # Reference ID (e.g., "lnk:5")
link.tooltip        # Tooltip text (if set)
link.style          # Applied style name

# Location info
link.paragraph      # Parent paragraph
link.section        # Section containing the link
```

### Filter Hyperlinks

```python
# Get external links only
external_links = [lnk for lnk in doc.hyperlinks if lnk.link_type == "external"]

# Get internal links only
internal_links = [lnk for lnk in doc.hyperlinks if lnk.link_type == "internal"]

# Find links by URL pattern
import re

https_links = [
    lnk for lnk in doc.hyperlinks if re.match(r"https://", lnk.target or "")
]
```

### Hyperlinks in Headers/Footers

```python
# Get hyperlinks from headers
for link in doc.header_hyperlinks:
    print(f"Header link: {link.text} -> {link.target}")

# Get hyperlinks from footers
for link in doc.footer_hyperlinks:
    print(f"Footer link: {link.text} -> {link.target}")
```

### Hyperlinks in Footnotes/Endnotes

```python
# Get hyperlinks from all footnotes
for link in doc.footnote_hyperlinks:
    print(f"Footnote link: {link.text} -> {link.target}")

# Get hyperlinks from specific footnote
footnote = doc.get_footnote(1)
for link in footnote.hyperlinks:
    print(f"Link in footnote 1: {link.text}")
```

## Editing Hyperlinks

Use the `ref` property from hyperlink objects to target specific links.

### Edit URL

```python
# Get the hyperlink
link = doc.hyperlinks[0]
print(f"Current URL: {link.target}")

# Change URL
doc.edit_hyperlink_url(link.ref, "https://new-url.com")

# Or by ref string directly
doc.edit_hyperlink_url("lnk:5", "https://updated-url.com")
```

### Edit Display Text

```python
# Change display text (untracked)
doc.edit_hyperlink_text("lnk:5", "New Display Text")

# Change display text with tracked changes
doc.edit_hyperlink_text("lnk:5", "New Display Text", track=True)
```

### Edit Bookmark Anchor

```python
# Change internal link target to different bookmark
doc.edit_hyperlink_anchor("lnk:3", "NewBookmarkName")
```

### Convert Between External and Internal

```python
# Convert external link to internal bookmark link
link = doc.hyperlinks[0]
doc.edit_hyperlink_anchor(link.ref, "InternalBookmark")
# This removes the URL and sets the anchor

# Convert internal link to external URL
doc.edit_hyperlink_url(link.ref, "https://example.com")
# This removes the anchor and sets the URL
```

## Removing Hyperlinks

### Remove and Keep Text

```python
# Remove hyperlink but keep the display text as plain text
doc.remove_hyperlink("lnk:5", keep_text=True)

# Using hyperlink object
link = doc.hyperlinks[0]
doc.remove_hyperlink(link.ref, keep_text=True)
```

### Remove Completely

```python
# Remove hyperlink and its display text entirely
doc.remove_hyperlink("lnk:5", keep_text=False)
```

### Remove All Hyperlinks

```python
# Remove all hyperlinks, keeping text
for link in list(doc.hyperlinks):  # Use list() to avoid mutation during iteration
    doc.remove_hyperlink(link.ref, keep_text=True)

# Or use bulk method
doc.remove_all_hyperlinks(keep_text=True)
```

## OOXML Reference

### External Hyperlink Structure

External hyperlinks use a relationship ID to reference the URL stored in `document.xml.rels`:

```xml
<!-- In document.xml -->
<w:hyperlink r:id="rId5">
  <w:r>
    <w:rPr>
      <w:rStyle w:val="Hyperlink"/>
    </w:rPr>
    <w:t>Click here</w:t>
  </w:r>
</w:hyperlink>

<!-- In document.xml.rels -->
<Relationship
  Id="rId5"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  Target="https://example.com"
  TargetMode="External"/>
```

### Internal Hyperlink (Bookmark) Structure

Internal hyperlinks use the `w:anchor` attribute instead of `r:id`:

```xml
<w:hyperlink w:anchor="DefinitionsSection">
  <w:r>
    <w:rPr>
      <w:rStyle w:val="Hyperlink"/>
    </w:rPr>
    <w:t>See Definitions</w:t>
  </w:r>
</w:hyperlink>

<!-- The bookmark target -->
<w:bookmarkStart w:id="0" w:name="DefinitionsSection"/>
<w:r><w:t>1. Definitions</w:t></w:r>
<w:bookmarkEnd w:id="0"/>
```

### Hyperlink with Tooltip

```xml
<w:hyperlink r:id="rId5" w:tooltip="Opens external website">
  <w:r>
    <w:t>Click here</w:t>
  </w:r>
</w:hyperlink>
```

### Hyperlink Namespaces

```python
NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
```

## Error Handling

```python
from python_docx_redline import Document
from python_docx_redline.errors import (
    TextNotFoundError,
    AmbiguousTextError,
    HyperlinkNotFoundError,
    BookmarkNotFoundError,
)

doc = Document("document.docx")

# Handle missing anchor text
try:
    doc.insert_hyperlink(
        url="https://example.com",
        text="link",
        after="nonexistent text",
    )
except TextNotFoundError as e:
    print(f"Anchor text not found: {e}")

# Handle ambiguous anchor text
try:
    doc.insert_hyperlink(
        url="https://example.com",
        text="link",
        after="common word",  # Appears multiple times
    )
except AmbiguousTextError as e:
    print(f"Multiple occurrences found: {e}")
    print(f"Use occurrence parameter to specify which one")

# Handle missing hyperlink ref
try:
    doc.edit_hyperlink_url("lnk:999", "https://example.com")
except HyperlinkNotFoundError as e:
    print(f"Hyperlink not found: {e}")
    print(f"Available refs: {[lnk.ref for lnk in doc.hyperlinks]}")

# Handle missing bookmark for internal link
try:
    doc.insert_hyperlink(
        anchor="NonexistentBookmark",
        text="link",
        after="see",
    )
except BookmarkNotFoundError as e:
    print(f"Bookmark not found: {e}")
    print(f"Available bookmarks: {e.available_bookmarks}")
```

## Complete Example

```python
from python_docx_redline import Document

doc = Document("legal_document.docx")

# Add external reference links
doc.insert_hyperlink(
    url="https://law.example.com/statute123",
    text="Statute 123",
    after="pursuant to",
)

# Add internal navigation
doc.add_bookmark("Definitions", on="Article 1: Definitions")
doc.insert_hyperlink(
    anchor="Definitions",
    text="(see Definitions)",
    after="As defined herein",
)

# Add link in footer for contact
doc.insert_hyperlink_in_footer(
    url="mailto:legal@company.com",
    text="legal@company.com",
    after="Contact:",
    footer_type="default",
)

# Add source link in footnote
doc.insert_hyperlink_in_footnote(
    note_id=1,
    url="https://example.com/citation",
    text="original source",
    after="See",
)

# List all links for review
print("Document hyperlinks:")
for link in doc.hyperlinks:
    link_type = "Internal" if link.link_type == "internal" else "External"
    print(f"  [{link_type}] {link.text} -> {link.target}")

# Update an outdated URL
for link in doc.hyperlinks:
    if "old-domain.com" in (link.target or ""):
        new_url = link.target.replace("old-domain.com", "new-domain.com")
        doc.edit_hyperlink_url(link.ref, new_url)

doc.save("legal_document_with_links.docx")
```

## API Reference

### Document Methods - Insert

| Method | Description |
|--------|-------------|
| `insert_hyperlink(url=..., text=..., after=...)` | Insert external hyperlink |
| `insert_hyperlink(anchor=..., text=..., after=...)` | Insert internal hyperlink to bookmark |
| `insert_hyperlink_in_header(url=..., text=..., after=..., header_type=...)` | Insert in header |
| `insert_hyperlink_in_footer(url=..., text=..., after=..., footer_type=...)` | Insert in footer |
| `insert_hyperlink_in_footnote(note_id, url=..., text=..., after=...)` | Insert in footnote |
| `insert_hyperlink_in_endnote(note_id, url=..., text=..., after=...)` | Insert in endnote |

### Document Methods - Edit

| Method | Description |
|--------|-------------|
| `edit_hyperlink_url(ref, url)` | Change hyperlink URL |
| `edit_hyperlink_text(ref, text, track=False)` | Change display text |
| `edit_hyperlink_anchor(ref, anchor)` | Change bookmark target |

### Document Methods - Remove

| Method | Description |
|--------|-------------|
| `remove_hyperlink(ref, keep_text=True)` | Remove single hyperlink |
| `remove_all_hyperlinks(keep_text=True)` | Remove all hyperlinks |

### Document Properties

| Property | Description |
|----------|-------------|
| `hyperlinks` | All hyperlinks in document body |
| `header_hyperlinks` | All hyperlinks in headers |
| `footer_hyperlinks` | All hyperlinks in footers |
| `footnote_hyperlinks` | All hyperlinks in footnotes |
| `endnote_hyperlinks` | All hyperlinks in endnotes |

### Hyperlink Object Properties

| Property | Type | Description |
|----------|------|-------------|
| `ref` | str | Reference ID for editing/removing (e.g., "lnk:5") |
| `text` | str | Display text |
| `target` | str \| None | URL (external) or bookmark name (internal) |
| `link_type` | str | "external" or "internal" |
| `tooltip` | str \| None | Tooltip text |
| `style` | str \| None | Applied character style |
| `paragraph` | Paragraph | Parent paragraph object |
| `section` | str \| None | Section name if identifiable |

### Insert Parameters

| Parameter | Description |
|-----------|-------------|
| `url` | External URL (mutually exclusive with `anchor`) |
| `anchor` | Internal bookmark name (mutually exclusive with `url`) |
| `text` | Display text for the hyperlink |
| `after` / `before` | Anchor text for insertion point |
| `occurrence` | `1`, `2`, `"first"`, `"last"`, `"all"` |
| `style` | Character style to apply (default: "Hyperlink") |
| `tooltip` | Hover text for the link |
