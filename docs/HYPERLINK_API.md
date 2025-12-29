# Hyperlink API Design Document

## Overview

### Problem Statement

python-docx-redline currently provides:
- **Read-only** access to hyperlinks via `BookmarkRegistry` (paragraph-level extraction)
- A `get_hyperlink_style()` template for the Hyperlink character style
- No API for **creating**, **editing**, or **deleting** hyperlinks

Users who need to programmatically work with hyperlinks must currently:
1. Manually construct OOXML XML with relationship IDs
2. Manage the document.xml.rels file directly
3. Handle text runs with Hyperlink style formatting
4. Deal with the complexity of internal vs external links

### Solution

Provide a high-level API that mirrors existing patterns for footnotes/comments:
- `insert_hyperlink()` - Insert clickable hyperlinks at anchor text
- `insert_hyperlink_in_header()` / `insert_hyperlink_in_footer()` - Hyperlinks in headers/footers
- `insert_hyperlink_in_footnote()` / `insert_hyperlink_in_endnote()` - Hyperlinks in notes
- `get_hyperlinks()` - List all hyperlinks with their targets
- `edit_hyperlink()` - Change URL or display text
- `remove_hyperlink()` - Remove link but optionally keep text

### Impact

```python
# BEFORE (manual XML construction)
rel_mgr = RelationshipManager(package, "word/document.xml")
r_id = rel_mgr.add_unique_relationship(RelationshipTypes.HYPERLINK, "https://example.com")
# ... manually find paragraph, construct hyperlink XML with r:id, insert...

# AFTER (proposed API)
doc.insert_hyperlink(
    url="https://example.com",
    text="Click here",
    after="For more information,"
)
```

---

## API Design

### 1. Insert External Hyperlinks

```python
def insert_hyperlink(
    self,
    url: str,
    text: str,
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Any | None = None,
    author: str | None = None,
    tooltip: str | None = None,
    track: bool = False,
) -> str:
    """Insert an external hyperlink at a specific location.

    Args:
        url: The URL to link to (e.g., "https://example.com")
        text: The display text for the hyperlink
        after: Text to insert after (mutually exclusive with before)
        before: Text to insert before (mutually exclusive with after)
        scope: Optional scope to limit search (paragraph ref, heading, etc.)
        author: Optional author override for tracked changes
        tooltip: Optional tooltip text shown on hover
        track: If True, wrap insertion in tracked change markup

    Returns:
        The relationship ID (rId) for the created hyperlink

    Raises:
        TextNotFoundError: If anchor text not found
        AmbiguousTextError: If anchor text found multiple times

    Example:
        >>> doc.insert_hyperlink(
        ...     url="https://www.law.cornell.edu/uscode/text/28/1782",
        ...     text="28 U.S.C. section 1782",
        ...     after="discovery statute"
        ... )
        'rId15'
    """
```

### 2. Insert Internal Hyperlinks (to Bookmarks)

```python
def insert_hyperlink(
    self,
    anchor: str,  # Bookmark name
    text: str,
    after: str | None = None,
    before: str | None = None,
    scope: str | dict | Any | None = None,
    author: str | None = None,
    track: bool = False,
) -> None:
    """Insert an internal hyperlink to a bookmark.

    Args:
        anchor: Bookmark name to link to (internal link)
        text: The display text for the hyperlink
        after: Text to insert after (mutually exclusive with before)
        before: Text to insert before (mutually exclusive with after)
        scope: Optional scope to limit search
        author: Optional author override for tracked changes
        track: If True, wrap insertion in tracked change markup

    Raises:
        TextNotFoundError: If anchor text not found
        ValueError: If bookmark doesn't exist (warning only, link still created)

    Example:
        >>> doc.insert_hyperlink(
        ...     anchor="DefinitionsSection",
        ...     text="See Definitions",
        ...     after="as defined below"
        ... )
    """
```

**Design Note**: Use keyword-only parameters to distinguish between external (`url=`) and internal (`anchor=`) hyperlinks.

### 3. Hyperlinks in Headers/Footers

```python
def insert_hyperlink_in_header(
    self,
    url: str | None = None,
    anchor: str | None = None,
    text: str,
    after: str | None = None,
    before: str | None = None,
    header_type: str = "default",
    track: bool = False,
) -> str | None:
    """Insert a hyperlink in a header.

    Args:
        url: External URL (mutually exclusive with anchor)
        anchor: Internal bookmark name (mutually exclusive with url)
        text: Display text for the hyperlink
        after: Text to insert after
        before: Text to insert before
        header_type: "default", "first", or "even"
        track: If True, track the insertion

    Returns:
        Relationship ID for external links, None for internal
    """

def insert_hyperlink_in_footer(
    self,
    url: str | None = None,
    anchor: str | None = None,
    text: str,
    after: str | None = None,
    before: str | None = None,
    footer_type: str = "default",
    track: bool = False,
) -> str | None:
    """Insert a hyperlink in a footer."""
```

### 4. Hyperlinks in Footnotes/Endnotes

```python
def insert_hyperlink_in_footnote(
    self,
    note_id: str | int,
    url: str | None = None,
    anchor: str | None = None,
    text: str,
    after: str | None = None,
    before: str | None = None,
    track: bool = False,
) -> str | None:
    """Insert a hyperlink inside an existing footnote.

    Args:
        note_id: The footnote ID to edit
        url: External URL (mutually exclusive with anchor)
        anchor: Internal bookmark name (mutually exclusive with url)
        text: Display text for the hyperlink
        after: Text to insert after within the footnote
        before: Text to insert before within the footnote
        track: If True, track the insertion
    """

def insert_hyperlink_in_endnote(
    self,
    note_id: str | int,
    url: str | None = None,
    anchor: str | None = None,
    text: str,
    after: str | None = None,
    before: str | None = None,
    track: bool = False,
) -> str | None:
    """Insert a hyperlink inside an existing endnote."""
```

### 5. Reading Hyperlinks

```python
@property
def hyperlinks(self) -> list[HyperlinkInfo]:
    """Get all hyperlinks in the document.

    Returns:
        List of HyperlinkInfo objects with link details

    Example:
        >>> for link in doc.hyperlinks:
        ...     print(f"{link.text} -> {link.target}")
    """

def get_hyperlink(self, ref: str) -> HyperlinkInfo | None:
    """Get a specific hyperlink by its ref."""

def get_hyperlinks_by_url(self, url_pattern: str) -> list[HyperlinkInfo]:
    """Find hyperlinks matching a URL pattern."""
```

### 6. Editing Hyperlinks

```python
def edit_hyperlink_url(self, ref: str, new_url: str) -> None:
    """Change the URL of an external hyperlink."""

def edit_hyperlink_text(
    self,
    ref: str,
    new_text: str,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Change the display text of a hyperlink."""

def edit_hyperlink_anchor(self, ref: str, new_anchor: str) -> None:
    """Change the target bookmark of an internal hyperlink."""
```

### 7. Removing Hyperlinks

```python
def remove_hyperlink(
    self,
    ref: str,
    keep_text: bool = True,
    track: bool = False,
    author: str | None = None,
) -> None:
    """Remove a hyperlink from the document.

    Args:
        ref: Hyperlink ref (e.g., "lnk:5")
        keep_text: If True (default), keep the display text without link.
                   If False, remove both link and text.
        track: If True and keep_text=False, show text removal as tracked deletion
        author: Optional author for tracked change
    """
```

---

## OOXML Structure

**External Hyperlink:**
```xml
<w:hyperlink r:id="rId5" w:tooltip="Optional tooltip">
  <w:r>
    <w:rPr>
      <w:rStyle w:val="Hyperlink"/>
    </w:rPr>
    <w:t>Display Text</w:t>
  </w:r>
</w:hyperlink>
```

**Internal Hyperlink:**
```xml
<w:hyperlink w:anchor="BookmarkName">
  <w:r>
    <w:rPr>
      <w:rStyle w:val="Hyperlink"/>
    </w:rPr>
    <w:t>See Section</w:t>
  </w:r>
</w:hyperlink>
```

**Relationship Entry (for external):**
```xml
<Relationship Id="rId5"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com"
    TargetMode="External"/>
```

---

## Implementation Phases

### Phase 1: Core External Hyperlinks (Priority: High)
- Create `HyperlinkOperations` class in `operations/hyperlinks.py`
- Implement `insert_hyperlink()` for external URLs in body text
- Add relationship management for hyperlinks (with `TargetMode="External"`)
- Ensure Hyperlink style creation
- Add `hyperlinks` property to Document
- Write unit tests

### Phase 2: Internal Hyperlinks (Priority: High)
- Support `anchor=` parameter for internal links
- Validate bookmark existence (warning if missing)
- Update `BookmarkRegistry` to track new hyperlinks

### Phase 3: Headers/Footers/Notes (Priority: Medium)
- Implement `insert_hyperlink_in_header()` / `insert_hyperlink_in_footer()`
- Implement `insert_hyperlink_in_footnote()` / `insert_hyperlink_in_endnote()`
- Handle part-specific relationship files

### Phase 4: Edit and Remove (Priority: Medium)
- Implement `edit_hyperlink_url()`, `edit_hyperlink_text()`, `edit_hyperlink_anchor()`
- Implement `remove_hyperlink()`

### Phase 5: Tracked Changes Support (Priority: Low)
- Add `track=True` support for all operations

### Phase 6: Documentation (Priority: Low)
- Add skill guide `skills/docx/hyperlinks.md`
- Update SKILL.md

---

## Key Implementation Details

1. **Relationship Management**: Use `RelationshipManager.add_unique_relationship()` with `TargetMode="External"` for external links.

2. **Part-specific Relationships**: Headers, footers, and footnotes have their own .rels files:
   - Body: `word/_rels/document.xml.rels`
   - Header: `word/_rels/header1.xml.rels`
   - Footer: `word/_rels/footer1.xml.rels`
   - Footnotes: `word/_rels/footnotes.xml.rels`

3. **Hyperlink Style**: Use `ensure_standard_styles(doc.styles, "Hyperlink")`.

4. **Text Search**: Reuse existing `TextSearch` infrastructure.
