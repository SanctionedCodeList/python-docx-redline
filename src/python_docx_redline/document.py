"""
Document class for editing Word documents with tracked changes.

This module provides the main Document class which handles loading .docx files,
inserting tracked changes, and saving the modified documents.
"""

import io
import logging
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Any, BinaryIO

if TYPE_CHECKING:
    from python_docx_redline.accessibility import Ref
    from python_docx_redline.criticmarkup import ApplyResult
    from python_docx_redline.models.comment import Comment
    from python_docx_redline.models.footnote import (
        Endnote,
        Footnote,
        OrphanedEndnote,
        OrphanedFootnote,
    )
    from python_docx_redline.models.header_footer import Footer, Header
    from python_docx_redline.models.paragraph import Paragraph
    from python_docx_redline.models.section import Section
    from python_docx_redline.models.table import Table, TableRow
    from python_docx_redline.models.tracked_change import TrackedChange

from lxml import etree

from .author import AuthorIdentity
from .constants import WORD_NAMESPACE, XML_NAMESPACE
from .match import Match
from .operations.batch import BatchOperations
from .operations.change_management import ChangeManagement
from .operations.comments import CommentOperations
from .operations.comparison import ComparisonOperations
from .operations.cross_references import (
    BookmarkInfo,
    CrossReference,
    CrossReferenceOperations,
    CrossReferenceTarget,
)
from .operations.edit_groups import EditGroupRegistry
from .operations.formatting import FormatOperations
from .operations.header_footer import HeaderFooterOperations
from .operations.hyperlinks import HyperlinkInfo, HyperlinkOperations
from .operations.images import ImageOperations
from .operations.notes import NoteOperations
from .operations.patterns import PatternOperations
from .operations.section import SectionOperations
from .operations.tables import TableOperations
from .operations.toc import TOC, TOCOperations
from .operations.tracked_changes import TrackedChangeOperations
from .package import OOXMLPackage
from .results import ComparisonStats, EditResult, FormatResult
from .scope import NoteScope, ScopeEvaluator, parse_note_scope
from .styles import StyleManager
from .text_search import TextSearch, TextSpan
from .tracked_xml import TrackedXMLGenerator
from .validation import ValidationError

logger = logging.getLogger(__name__)


class Document:
    """Main class for working with Word documents.

    This class handles loading .docx files (unpacking if needed), making tracked
    edits, and saving the results. It provides a high-level API that hides the
    complexity of OOXML manipulation.

    Documents can be loaded from:
    - File paths (str or Path)
    - Raw bytes
    - BytesIO objects
    - Open file objects (in binary mode)

    Example:
        >>> doc = Document("contract.docx")
        >>> doc.insert_tracked("new clause text", after="Section 2.1")
        >>> doc.save("contract_edited.docx")

    Example with bytes:
        >>> with open("contract.docx", "rb") as f:
        ...     doc = Document(f.read())
        >>> doc.insert_tracked(" [REVIEWED]", after="Section 1")
        >>> doc_bytes = doc.save_to_bytes()

    Attributes:
        path: Path to the document file (None for in-memory documents)
        author: Author name for tracked changes
        xml_tree: Parsed XML tree of the document
        xml_root: Root element of the XML tree
    """

    def __init__(
        self,
        source: str | Path | bytes | BinaryIO,
        author: str | AuthorIdentity = "Claude",
        minimal_edits: bool = True,
    ) -> None:
        """Initialize a Document from a .docx file or in-memory data.

        Args:
            source: Document source - can be:
                    - Path to a .docx file (str or Path)
                    - Raw bytes of a .docx file
                    - BytesIO object containing a .docx file
                    - Open file object in binary mode
            author: Author name (str) or full AuthorIdentity for MS365 integration
                   (default: "Claude")
            minimal_edits: If True (default), tracked changes use word-level diffing
                   to produce human-looking redlines. If False, uses coarse
                   "delete all + insert all" pattern. Per-operation overrides
                   are available via the `minimal` parameter on individual methods.

        Raises:
            ValidationError: If the document cannot be loaded or is invalid

        Example:
            >>> # From file path
            >>> doc = Document("contract.docx", author="John Doe")
            >>>
            >>> # From bytes
            >>> with open("contract.docx", "rb") as f:
            ...     doc = Document(f.read())
            >>>
            >>> # From BytesIO
            >>> import io
            >>> buffer = io.BytesIO(doc_bytes)
            >>> doc = Document(buffer)
            >>>
            >>> # Full MS365 identity
            >>> identity = AuthorIdentity(
            ...     author="Hancock, Parker",
            ...     email="parker.hancock@company.com",
            ...     provider_id="AD",
            ...     guid="c5c513d2-1f51-4d69-ae91-17e5787f9bfc"
            ... )
            >>> doc = Document("contract.docx", author=identity)
        """
        # Detect and normalize source type
        if isinstance(source, bytes):
            self._source_stream: BinaryIO | None = io.BytesIO(source)
            self.path: Path | None = None
        elif hasattr(source, "read"):
            self._source_stream = source  # type: ignore[assignment]
            self.path = None
        else:
            self._source_stream = None
            self.path = Path(source)

        # Store author identity (convert string to AuthorIdentity if needed)
        if isinstance(author, str):
            self._author_identity = None
            self.author = author
        else:
            self._author_identity = author
            self.author = author.display_name

        # Store minimal edits setting (propagates to all operations)
        self._minimal_edits = minimal_edits

        self._package: OOXMLPackage | None = None

        # Initialize components
        self._text_search = TextSearch()
        self._xml_generator = TrackedXMLGenerator(
            doc=self, author=author if isinstance(author, str) else author.display_name
        )

        # Load the document
        self._load_document()

    def _load_document(self) -> None:
        """Load and parse the Word document XML.

        If the document is a .docx file (ZIP archive), it will be extracted
        to a temporary directory using OOXMLPackage. The main document.xml
        is then parsed.

        Supports loading from file paths or in-memory streams (BytesIO).

        Raises:
            ValidationError: If the document cannot be loaded
        """
        # Determine source: stream or path
        if self._source_stream is not None:
            source: Path | BinaryIO = self._source_stream
            source_desc = "<in-memory document>"
        else:
            assert self.path is not None
            source = self.path
            source_desc = str(self.path)

        # Try to open as ZIP package (.docx)
        try:
            self._package = OOXMLPackage.open(source)
        except ValidationError as e:
            # Not a valid ZIP - check if it's raw XML
            if self._source_stream is not None:
                raise ValidationError("In-memory source must be a valid .docx (ZIP) file") from e
            # Assume it's already an unpacked XML file
            self._package = None

        # Parse the document.xml
        try:
            if self._package is not None:
                document_xml = self._package.get_part_path("word/document.xml")
            else:
                document_xml = self.path  # type: ignore

            if not document_xml.exists():
                raise ValidationError(f"document.xml not found in {source_desc}")

            # Parse XML with lxml
            parser = etree.XMLParser(remove_blank_text=False)
            self.xml_tree = etree.parse(str(document_xml), parser)
            self.xml_root = self.xml_tree.getroot()

        except etree.XMLSyntaxError as e:
            raise ValidationError(f"Invalid XML in document: {e}") from e
        except ValidationError:
            raise
        except Exception as e:
            raise ValidationError(f"Failed to parse document XML: {e}") from e

        # Ensure required namespaces are declared for MS365 identity
        if self._author_identity is not None:
            self._ensure_ms365_namespaces()

    def _ensure_ms365_namespaces(self) -> None:
        """Ensure MS365 identity namespaces are declared on document root.

        When using AuthorIdentity with provider_id or guid, the document root
        must have the w15 and w16du namespaces declared, AND these namespaces
        must be listed in mc:Ignorable for the attributes to be valid according
        to the OOXML spec.

        The mc:Ignorable attribute tells conformant processors to ignore
        unknown namespace extensions rather than rejecting the document.
        """
        from .constants import MC_NAMESPACE, W15_NAMESPACE, W16DU_NAMESPACE

        root = self.xml_root
        nsmap = dict(root.nsmap)

        # Check if namespaces need to be added
        needs_update = False
        if "w15" not in nsmap:
            nsmap["w15"] = W15_NAMESPACE
            needs_update = True

        if "w16du" not in nsmap:
            nsmap["w16du"] = W16DU_NAMESPACE
            needs_update = True

        if "mc" not in nsmap:
            nsmap["mc"] = MC_NAMESPACE
            needs_update = True

        # If namespaces need to be added, we have to recreate the root element
        # because lxml doesn't allow modifying nsmap after creation
        if needs_update:
            new_root = etree.Element(root.tag, nsmap=nsmap)
            # Copy attributes
            for key, value in root.attrib.items():
                new_root.set(key, value)
            # Move all children
            for child in root:
                new_root.append(child)
            # Replace in tree
            self.xml_tree._setroot(new_root)
            self.xml_root = new_root
            root = new_root

        # Ensure mc:Ignorable includes w15 and w16du
        mc_ignorable_attr = f"{{{MC_NAMESPACE}}}Ignorable"
        current_ignorable = root.get(mc_ignorable_attr, "")
        ignorable_parts = current_ignorable.split() if current_ignorable else []

        updated = False
        if "w15" not in ignorable_parts:
            ignorable_parts.append("w15")
            updated = True
        if "w16du" not in ignorable_parts:
            ignorable_parts.append("w16du")
            updated = True

        if updated:
            root.set(mc_ignorable_attr, " ".join(ignorable_parts))

    # Backward compatibility properties for package access
    @property
    def _temp_dir(self) -> Path | None:
        """Get the temp directory from the package (backward compatibility)."""
        return self._package.temp_dir if self._package is not None else None

    @property
    def _is_zip(self) -> bool:
        """Check if this is a ZIP package (backward compatibility)."""
        return self._package is not None

    @property
    def _comment_ops(self) -> CommentOperations:
        """Get the CommentOperations instance (lazy initialization)."""
        if not hasattr(self, "_comment_ops_instance"):
            self._comment_ops_instance = CommentOperations(self)
        return self._comment_ops_instance

    @property
    def _tracked_ops(self) -> TrackedChangeOperations:
        """Get the TrackedChangeOperations instance (lazy initialization)."""
        if not hasattr(self, "_tracked_ops_instance"):
            self._tracked_ops_instance = TrackedChangeOperations(self)
        return self._tracked_ops_instance

    @property
    def _change_mgmt(self) -> ChangeManagement:
        """Get the ChangeManagement instance (lazy initialization)."""
        if not hasattr(self, "_change_mgmt_instance"):
            self._change_mgmt_instance = ChangeManagement(self)
        return self._change_mgmt_instance

    @property
    def _format_ops(self) -> FormatOperations:
        """Get the FormatOperations instance (lazy initialization)."""
        if not hasattr(self, "_format_ops_instance"):
            self._format_ops_instance = FormatOperations(self)
        return self._format_ops_instance

    @property
    def _table_ops(self) -> TableOperations:
        """Get the TableOperations instance (lazy initialization)."""
        if not hasattr(self, "_table_ops_instance"):
            self._table_ops_instance = TableOperations(self)
        return self._table_ops_instance

    @property
    def _note_ops(self) -> NoteOperations:
        """Get the NoteOperations instance (lazy initialization)."""
        if not hasattr(self, "_note_ops_instance"):
            self._note_ops_instance = NoteOperations(self)
        return self._note_ops_instance

    @property
    def _header_footer_ops(self) -> HeaderFooterOperations:
        """Get the HeaderFooterOperations instance (lazy initialization)."""
        if not hasattr(self, "_header_footer_ops_instance"):
            self._header_footer_ops_instance = HeaderFooterOperations(self)
        return self._header_footer_ops_instance

    @property
    def _image_ops(self) -> ImageOperations:
        """Get the ImageOperations instance (lazy initialization)."""
        if not hasattr(self, "_image_ops_instance"):
            self._image_ops_instance = ImageOperations(self)
        return self._image_ops_instance

    @property
    def _batch_ops(self) -> BatchOperations:
        """Get the BatchOperations instance (lazy initialization)."""
        if not hasattr(self, "_batch_ops_instance"):
            self._batch_ops_instance = BatchOperations(self)
        return self._batch_ops_instance

    @property
    def _section_ops(self) -> SectionOperations:
        """Get the SectionOperations instance (lazy initialization)."""
        if not hasattr(self, "_section_ops_instance"):
            self._section_ops_instance = SectionOperations(self)
        return self._section_ops_instance

    @property
    def _pattern_ops(self) -> PatternOperations:
        """Get the PatternOperations instance (lazy initialization)."""
        if not hasattr(self, "_pattern_ops_instance"):
            self._pattern_ops_instance = PatternOperations(self)
        return self._pattern_ops_instance

    @property
    def _comparison_ops(self) -> ComparisonOperations:
        """Get the ComparisonOperations instance (lazy initialization)."""
        if not hasattr(self, "_comparison_ops_instance"):
            self._comparison_ops_instance = ComparisonOperations(self)
        return self._comparison_ops_instance

    @property
    def _hyperlink_ops(self) -> HyperlinkOperations:
        """Get the HyperlinkOperations instance (lazy initialization)."""
        if not hasattr(self, "_hyperlink_ops_instance"):
            self._hyperlink_ops_instance = HyperlinkOperations(self)
        return self._hyperlink_ops_instance

    @property
    def _toc_ops(self) -> TOCOperations:
        """Get the TOCOperations instance (lazy initialization)."""
        if not hasattr(self, "_toc_ops_instance"):
            self._toc_ops_instance = TOCOperations(self)
        return self._toc_ops_instance

    @property
    def _cross_reference_ops(self) -> CrossReferenceOperations:
        """Get the CrossReferenceOperations instance (lazy initialization)."""
        if not hasattr(self, "_cross_reference_ops_instance"):
            self._cross_reference_ops_instance = CrossReferenceOperations(self)
        return self._cross_reference_ops_instance

    @property
    def styles(self) -> StyleManager:
        """Get the style manager for this document.

        Provides access to document styles for reading and modifying style
        definitions. The StyleManager is lazily initialized on first access.

        Returns:
            StyleManager for reading and modifying styles

        Raises:
            ValidationError: If the document is not a .docx package

        Example:
            >>> doc = Document("contract.docx")
            >>> for style in doc.styles.list(style_type=StyleType.PARAGRAPH):
            ...     print(f"{style.style_id}: {style.name}")
            >>>
            >>> # Ensure a style exists
            >>> doc.styles.ensure_style(
            ...     style_id="MyStyle",
            ...     name="My Style",
            ...     style_type=StyleType.CHARACTER,
            ...     run_formatting=RunFormatting(bold=True)
            ... )
            >>> doc.save("output.docx")  # Style changes saved automatically
        """
        if not hasattr(self, "_style_manager_instance"):
            if self._package is None:
                raise ValidationError(
                    "StyleManager requires a .docx package. "
                    "Raw XML files do not have a styles.xml component."
                )
            self._style_manager_instance = StyleManager(self._package)
        return self._style_manager_instance

    # View capabilities (Phase 3)

    @property
    def paragraphs(self) -> list["Paragraph"]:
        """Get all paragraphs in the document.

        Returns a list of Paragraph wrapper objects that provide convenient
        access to paragraph text, style, and other properties.

        Returns:
            List of Paragraph objects for all paragraphs in document

        Example:
            >>> doc = Document("contract.docx")
            >>> for para in doc.paragraphs:
            ...     if para.is_heading():
            ...         print(f"Section: {para.text}")
        """
        from python_docx_redline.models.paragraph import Paragraph

        return [Paragraph(p) for p in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p")]

    @property
    def sections(self) -> list["Section"]:
        """Get document sections parsed by heading structure.

        A section consists of a heading paragraph followed by all paragraphs
        until the next heading. Paragraphs before the first heading belong to
        an intro section with no heading.

        Returns:
            List of Section objects

        Example:
            >>> doc = Document("contract.docx")
            >>> for section in doc.sections:
            ...     if section.heading:
            ...         print(f"Section: {section.heading_text}")
            ...     print(f"  {len(section.paragraphs)} paragraphs")
        """
        from python_docx_redline.models.section import Section

        return Section.from_document(self.xml_root)

    @property
    def tables(self) -> list["Table"]:
        """Get all tables in the document.

        Returns:
            List of Table objects

        Example:
            >>> doc = Document("contract.docx")
            >>> for i, table in enumerate(doc.tables):
            ...     print(f"Table {i}: {table.row_count} rows × {table.col_count} cols")
        """
        return self._table_ops.all

    def find_table(self, containing: str, case_sensitive: bool = True) -> "Table | None":
        """Find the first table containing specific text.

        Args:
            containing: Text to search for in table cells
            case_sensitive: Whether search should be case sensitive (default: True)

        Returns:
            First Table containing the text, or None if not found

        Example:
            >>> doc = Document("contract.docx")
            >>> pricing_table = doc.find_table("Total Price")
            >>> if pricing_table:
            ...     print(f"Found table with {pricing_table.row_count} rows")
        """
        return self._table_ops.find(containing, case_sensitive)

    @property
    def comments(self) -> list["Comment"]:
        """Get all comments in the document.

        Returns a list of Comment objects with both the comment content
        and the marked text range they apply to.

        Returns:
            List of Comment objects, empty list if no comments

        Example:
            >>> doc = Document("reviewed.docx")
            >>> for comment in doc.comments:
            ...     print(f"{comment.author}: {comment.text}")
            ...     if comment.marked_text:
            ...         print(f"  Regarding: '{comment.marked_text}'")
        """
        return self._comment_ops.all

    def get_comments(
        self,
        *,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> list["Comment"]:
        """Get comments with optional filtering.

        Args:
            author: Filter to comments by this author
            scope: Limit to comments within a specific scope
                   (section name, dict filter, or callable)

        Returns:
            Filtered list of Comment objects

        Example:
            >>> # Get all comments by a specific reviewer
            >>> comments = doc.get_comments(author="John Doe")
            >>>
            >>> # Get comments in a specific section
            >>> comments = doc.get_comments(scope="section:Introduction")
        """
        return self._comment_ops.get(author=author, scope=scope)

    def get_text(self, skip_deleted_paragraphs: bool = True) -> str:
        """Extract all visible text content from the document.

        Returns plain text with paragraphs separated by double newlines.
        This excludes deleted text (tracked changes) and only returns visible content.
        This is useful for understanding document content before making edits.

        Args:
            skip_deleted_paragraphs: If True (default), skip paragraphs whose
                paragraph marks are marked as deleted and have no visible content.
                This makes the output match what Word shows in "Final" view.
                Set to False to include empty lines for deleted paragraphs.

        Returns:
            Plain text content of the entire document (excluding deletions)

        Example:
            >>> doc = Document("contract.docx")
            >>> text = doc.get_text()
            >>> if "confidential" in text.lower():
            ...     print("Document contains confidential information")

            >>> # Include empty lines for deleted paragraphs
            >>> text_with_gaps = doc.get_text(skip_deleted_paragraphs=False)
        """
        # Extract only visible text (w:t), not deleted text (w:delText)
        paragraphs_text = []
        for para in self.paragraphs:
            # Get only w:t elements, not w:delText
            text_elements = para.element.findall(f".//{{{WORD_NAMESPACE}}}t")
            para_text = "".join(elem.text or "" for elem in text_elements)

            # Skip empty paragraphs with deleted paragraph marks
            if skip_deleted_paragraphs and not para_text.strip():
                if self._has_deleted_paragraph_mark(para.element):
                    continue

            paragraphs_text.append(para_text)
        return "\n\n".join(paragraphs_text)

    def _has_deleted_paragraph_mark(self, para_element: etree._Element) -> bool:
        """Check if a paragraph has its paragraph mark marked as deleted.

        A deleted paragraph mark is indicated by a <w:del> element inside
        <w:pPr>/<w:rPr>. When Word accepts this tracked change, the paragraph
        merges with the following paragraph.

        Args:
            para_element: The paragraph XML element to check

        Returns:
            True if the paragraph mark is marked as deleted, False otherwise
        """
        # Look for w:pPr/w:rPr/w:del
        p_pr = para_element.find(f"{{{WORD_NAMESPACE}}}pPr")
        if p_pr is None:
            return False

        r_pr = p_pr.find(f"{{{WORD_NAMESPACE}}}rPr")
        if r_pr is None:
            return False

        del_elem = r_pr.find(f"{{{WORD_NAMESPACE}}}del")
        return del_elem is not None

    def has_tracked_changes(self) -> bool:
        """Check if the document contains any tracked changes.

        Looks for w:ins (insertions), w:del (deletions), w:moveFrom, or w:moveTo
        elements in the document XML.

        Returns:
            True if the document contains tracked changes, False otherwise

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.insert_tracked("new text", after="anchor")
            >>> assert doc.has_tracked_changes()  # True after editing
        """
        # Check for tracked change elements
        tracked_elements = [
            f"{{{WORD_NAMESPACE}}}ins",
            f"{{{WORD_NAMESPACE}}}del",
            f"{{{WORD_NAMESPACE}}}moveFrom",
            f"{{{WORD_NAMESPACE}}}moveTo",
        ]

        for elem_tag in tracked_elements:
            if self.xml_root.find(f".//{elem_tag}") is not None:
                return True

        return False

    def find_all(
        self,
        text: str,
        regex: bool = False,
        case_sensitive: bool = True,
        scope: str | dict | Any | None = None,
        context_chars: int = 40,
        fuzzy: float | dict[str, Any] | None = None,
        include_deleted: bool = False,
        include_footnotes: bool = False,
        include_endnotes: bool = False,
    ) -> list[Match]:
        """Find all occurrences of text in the document with location metadata.

        This method returns all matches with rich context information, making it
        easy to preview what text will be matched before performing operations.
        It's especially useful for understanding ambiguous text searches.

        Args:
            text: The text or regex pattern to search for
            regex: Whether to treat text as a regex pattern (default: False)
            case_sensitive: Whether to perform case-sensitive search (default: True)
            scope: Limit search scope. Can be:
                - None: Search all body paragraphs (default)
                - str: "text" for paragraph containing text
                - str: "section:Name" for paragraphs in named section
                - str: "footnotes" for all footnotes
                - str: "endnotes" for all endnotes
                - str: "notes" for all footnotes and endnotes
                - str: "footnote:N" for specific footnote by ID
                - str: "endnote:N" for specific endnote by ID
                - dict: {"contains": "text", ...} for complex filters
            context_chars: Number of characters to show before/after match (default: 40)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'
            include_deleted: If True, include text inside tracked deletions when
                searching. If False (default), skip text in w:del elements.
            include_footnotes: If True, also search within footnotes (default: False).
                Matches will have location like "footnote:1", "footnote:2", etc.
            include_endnotes: If True, also search within endnotes (default: False).
                Matches will have location like "endnote:1", "endnote:2", etc.

        Returns:
            List of Match objects with text, context, location, and metadata.
            When searching footnotes/endnotes, the location field indicates which
            note contains the match (e.g., "footnote:1", "endnote:2").

        Raises:
            re.error: If regex=True and the pattern is invalid
            ImportError: If fuzzy matching requested but rapidfuzz not installed
            ValueError: If both fuzzy and regex are specified

        Example:
            >>> # Find all occurrences of a phrase
            >>> matches = doc.find_all("production products")
            >>> print(f"Found {len(matches)} occurrences")
            >>> for match in matches:
            ...     print(f"[{match.index}] {match.location}: {match.context}")
            Found 2 occurrences
            [0] body: ...Therefore, production products utilizing...
            [1] table:0:row:45:cell:1: ...Therefore, production products utilizing...
            >>>
            >>> # Use regex
            >>> matches = doc.find_all(r"\\d+ days", regex=True)
            >>>
            >>> # Case-insensitive search
            >>> matches = doc.find_all("IMPORTANT", case_sensitive=False)
            >>>
            >>> # Search only in tables
            >>> matches = doc.find_all("text", scope={"location": "tables"})
            >>>
            >>> # Search only in footnotes
            >>> matches = doc.find_all("citation", scope="footnotes")
            >>>
            >>> # Search in a specific footnote
            >>> matches = doc.find_all("text", scope="footnote:1")
            >>>
            >>> # Search body AND footnotes
            >>> matches = doc.find_all("text", include_footnotes=True)
            >>>
            >>> # Fuzzy matching with threshold
            >>> matches = doc.find_all("production products", fuzzy=0.85)
            >>>
            >>> # Include text inside tracked deletions
            >>> matches = doc.find_all("deleted text", include_deleted=True)
        """
        # Check if scope is targeting notes specifically
        note_scope = parse_note_scope(scope) if isinstance(scope, str) else None

        if note_scope is not None:
            # Scope targets notes - search only within notes
            return self._find_all_in_notes(
                text=text,
                note_scope=note_scope,
                regex=regex,
                case_sensitive=case_sensitive,
                context_chars=context_chars,
                fuzzy=fuzzy,
                include_deleted=include_deleted,
            )

        # Search document body
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Parse fuzzy configuration if provided
        from .fuzzy import parse_fuzzy_config

        fuzzy_config = parse_fuzzy_config(fuzzy)

        # Find all text spans using TextSearch
        spans = self._text_search.find_text(
            text,
            paragraphs,
            case_sensitive=case_sensitive,
            regex=regex,
            fuzzy=fuzzy_config,
            include_deleted=include_deleted,
        )

        # Convert TextSpans to Match objects with rich metadata
        matches = []
        for idx, span in enumerate(spans):
            # Get paragraph index within ALL paragraphs (not just filtered)
            paragraph_index = all_paragraphs.index(span.paragraph)

            # Get full paragraph text
            text_elements = span.paragraph.findall(f".//{{{WORD_NAMESPACE}}}t")
            paragraph_text = "".join(elem.text or "" for elem in text_elements)

            # Determine location string
            location = self._get_location_string(span.paragraph)

            # Get context with custom size
            context = self._get_context_with_size(span, context_chars)

            # Create Match object
            match = Match(
                index=idx,
                text=span.text,
                context=context,
                paragraph_index=paragraph_index,
                paragraph_text=paragraph_text,
                location=location,
                span=span,
            )
            matches.append(match)

        # Optionally search footnotes
        if include_footnotes:
            footnote_matches = self._find_all_in_notes(
                text=text,
                note_scope=NoteScope(scope_type="footnotes"),
                regex=regex,
                case_sensitive=case_sensitive,
                context_chars=context_chars,
                fuzzy=fuzzy,
                include_deleted=include_deleted,
            )
            # Re-index the matches to continue from where body matches left off
            for fm in footnote_matches:
                fm.index = len(matches)
                matches.append(fm)

        # Optionally search endnotes
        if include_endnotes:
            endnote_matches = self._find_all_in_notes(
                text=text,
                note_scope=NoteScope(scope_type="endnotes"),
                regex=regex,
                case_sensitive=case_sensitive,
                context_chars=context_chars,
                fuzzy=fuzzy,
                include_deleted=include_deleted,
            )
            # Re-index the matches to continue from where body/footnote matches left off
            for em in endnote_matches:
                em.index = len(matches)
                matches.append(em)

        return matches

    def _find_all_in_notes(
        self,
        text: str,
        note_scope: NoteScope,
        regex: bool = False,
        case_sensitive: bool = True,
        context_chars: int = 40,
        fuzzy: float | dict[str, Any] | None = None,
        include_deleted: bool = False,
    ) -> list[Match]:
        """Find all occurrences of text within footnotes/endnotes.

        Args:
            text: The text or regex pattern to search for
            note_scope: NoteScope specifying which notes to search
            regex: Whether to treat text as a regex pattern
            case_sensitive: Whether to perform case-sensitive search
            context_chars: Number of characters to show before/after match
            fuzzy: Fuzzy matching configuration
            include_deleted: If True, include text inside tracked deletions

        Returns:
            List of Match objects with location indicating the note source
        """
        from .fuzzy import parse_fuzzy_config

        fuzzy_config = parse_fuzzy_config(fuzzy)
        matches: list[Match] = []
        match_index = 0

        # Determine which notes to search
        search_footnotes = note_scope.scope_type in ("footnotes", "footnote", "notes")
        search_endnotes = note_scope.scope_type in ("endnotes", "endnote", "notes")
        specific_note_id = note_scope.note_id

        # Search footnotes
        if search_footnotes:
            footnotes = self.footnotes
            for footnote in footnotes:
                # If searching specific note, skip others
                if specific_note_id is not None and footnote.id != specific_note_id:
                    continue

                # Get paragraphs from this footnote
                note_paragraphs = [p._element for p in footnote.paragraphs]
                if not note_paragraphs:
                    continue

                # Find matches in this note
                spans = self._text_search.find_text(
                    text,
                    note_paragraphs,
                    case_sensitive=case_sensitive,
                    regex=regex,
                    fuzzy=fuzzy_config,
                    include_deleted=include_deleted,
                )

                # Convert spans to Match objects
                for span in spans:
                    # Get paragraph text
                    text_elements = span.paragraph.findall(f".//{{{WORD_NAMESPACE}}}t")
                    paragraph_text = "".join(elem.text or "" for elem in text_elements)

                    # Location indicates which footnote
                    location = f"footnote:{footnote.id}"

                    # Get context
                    context = self._get_context_with_size(span, context_chars)

                    # Create Match object (paragraph_index = position within note)
                    match = Match(
                        index=match_index,
                        text=span.text,
                        context=context,
                        paragraph_index=note_paragraphs.index(span.paragraph),
                        paragraph_text=paragraph_text,
                        location=location,
                        span=span,
                    )
                    matches.append(match)
                    match_index += 1

        # Search endnotes
        if search_endnotes:
            endnotes = self.endnotes
            for endnote in endnotes:
                # If searching specific note, skip others
                if specific_note_id is not None and endnote.id != specific_note_id:
                    continue

                # Get paragraphs from this endnote
                note_paragraphs = [p._element for p in endnote.paragraphs]
                if not note_paragraphs:
                    continue

                # Find matches in this note
                spans = self._text_search.find_text(
                    text,
                    note_paragraphs,
                    case_sensitive=case_sensitive,
                    regex=regex,
                    fuzzy=fuzzy_config,
                    include_deleted=include_deleted,
                )

                # Convert spans to Match objects
                for span in spans:
                    # Get paragraph text
                    text_elements = span.paragraph.findall(f".//{{{WORD_NAMESPACE}}}t")
                    paragraph_text = "".join(elem.text or "" for elem in text_elements)

                    # Location indicates which endnote
                    location = f"endnote:{endnote.id}"

                    # Get context
                    context = self._get_context_with_size(span, context_chars)

                    # Create Match object (paragraph_index = position within note)
                    match = Match(
                        index=match_index,
                        text=span.text,
                        context=context,
                        paragraph_index=note_paragraphs.index(span.paragraph),
                        paragraph_text=paragraph_text,
                        location=location,
                        span=span,
                    )
                    matches.append(match)
                    match_index += 1

        return matches

    def find_in_footnotes(
        self,
        text: str,
        regex: bool = False,
        case_sensitive: bool = True,
        context_chars: int = 40,
        fuzzy: float | dict[str, Any] | None = None,
        include_deleted: bool = False,
    ) -> list[Match]:
        """Search all footnotes for text.

        This is a convenience method equivalent to:
        doc.find_all(text, scope="footnotes", ...)

        Args:
            text: The text or regex pattern to search for
            regex: Whether to treat text as a regex pattern (default: False)
            case_sensitive: Whether to perform case-sensitive search (default: True)
            context_chars: Number of characters to show before/after match (default: 40)
            fuzzy: Fuzzy matching configuration
            include_deleted: If True, include text inside tracked deletions

        Returns:
            List of Match objects with location indicating which footnote
            (e.g., "footnote:1", "footnote:2")

        Example:
            >>> matches = doc.find_in_footnotes("citation")
            >>> for match in matches:
            ...     print(f"{match.location}: {match.text}")
            footnote:1: citation needed
            footnote:3: original citation
        """
        return self._find_all_in_notes(
            text=text,
            note_scope=NoteScope(scope_type="footnotes"),
            regex=regex,
            case_sensitive=case_sensitive,
            context_chars=context_chars,
            fuzzy=fuzzy,
            include_deleted=include_deleted,
        )

    def find_in_endnotes(
        self,
        text: str,
        regex: bool = False,
        case_sensitive: bool = True,
        context_chars: int = 40,
        fuzzy: float | dict[str, Any] | None = None,
        include_deleted: bool = False,
    ) -> list[Match]:
        """Search all endnotes for text.

        This is a convenience method equivalent to:
        doc.find_all(text, scope="endnotes", ...)

        Args:
            text: The text or regex pattern to search for
            regex: Whether to treat text as a regex pattern (default: False)
            case_sensitive: Whether to perform case-sensitive search (default: True)
            context_chars: Number of characters to show before/after match (default: 40)
            fuzzy: Fuzzy matching configuration
            include_deleted: If True, include text inside tracked deletions

        Returns:
            List of Match objects with location indicating which endnote
            (e.g., "endnote:1", "endnote:2")

        Example:
            >>> matches = doc.find_in_endnotes("reference")
            >>> for match in matches:
            ...     print(f"{match.location}: {match.text}")
            endnote:1: See reference 42
            endnote:2: Reference to original source
        """
        return self._find_all_in_notes(
            text=text,
            note_scope=NoteScope(scope_type="endnotes"),
            regex=regex,
            case_sensitive=case_sensitive,
            context_chars=context_chars,
            fuzzy=fuzzy,
            include_deleted=include_deleted,
        )

    def _get_location_string(self, paragraph: Any) -> str:
        """Get a human-readable location string for a paragraph.

        Returns strings like:
        - "body" for main document body
        - "table:0:row:2:cell:1" for table cells
        - "header:0" for headers
        - "footer:0" for footers

        Args:
            paragraph: The paragraph Element

        Returns:
            A human-readable location string
        """
        # Check if in table
        parent = paragraph.getparent()
        while parent is not None:
            tag = parent.tag
            if tag == f"{{{WORD_NAMESPACE}}}tc":  # Table cell
                # Find table, row, and cell indices
                return self._get_table_location(parent)
            elif tag == f"{{{WORD_NAMESPACE}}}hdr":  # Header
                # Find which header (first, default, even)
                return "header"
            elif tag == f"{{{WORD_NAMESPACE}}}ftr":  # Footer
                return "footer"
            elif tag == f"{{{WORD_NAMESPACE}}}footnote":
                return "footnote"
            elif tag == f"{{{WORD_NAMESPACE}}}endnote":
                return "endnote"
            parent = parent.getparent()

        return "body"

    def _get_table_location(self, cell: Any) -> str:
        """Get detailed location string for a table cell.

        Args:
            cell: The table cell Element

        Returns:
            Location string like "table:0:row:2:cell:1"
        """
        # Find the row
        row = cell.getparent()
        if row is None or row.tag != f"{{{WORD_NAMESPACE}}}tr":
            return "table"

        # Find the table body
        tbl_body = row.getparent()
        if tbl_body is None:
            return "table"

        # Find the table
        table = tbl_body.getparent() if tbl_body.tag != f"{{{WORD_NAMESPACE}}}tbl" else tbl_body
        if table is None or table.tag != f"{{{WORD_NAMESPACE}}}tbl":
            return "table"

        # Get all tables in document
        all_tables = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}tbl"))
        try:
            table_idx = all_tables.index(table)
        except ValueError:
            table_idx = 0

        # Get row index within table
        all_rows = list(table.iter(f"{{{WORD_NAMESPACE}}}tr"))
        try:
            row_idx = all_rows.index(row)
        except ValueError:
            row_idx = 0

        # Get cell index within row
        all_cells = list(row.iter(f"{{{WORD_NAMESPACE}}}tc"))
        try:
            cell_idx = all_cells.index(cell)
        except ValueError:
            cell_idx = 0

        return f"table:{table_idx}:row:{row_idx}:cell:{cell_idx}"

    def _get_context_with_size(self, span: TextSpan, context_chars: int) -> str:
        """Get surrounding context with custom size.

        Args:
            span: The TextSpan to get context for
            context_chars: Number of characters before/after to include

        Returns:
            Context string with ellipsis if needed
        """
        # Extract text only from w:t elements
        text_elements = span.paragraph.findall(f".//{{{WORD_NAMESPACE}}}t")
        para_text = "".join(elem.text or "" for elem in text_elements)
        matched = span.text

        # Find the match position in the full paragraph text
        match_pos = para_text.find(matched)
        if match_pos == -1:
            return matched

        # Get context window
        start = max(0, match_pos - context_chars)
        end = min(len(para_text), match_pos + len(matched) + context_chars)

        context = para_text[start:end]

        # Add ellipsis if needed
        if start > 0:
            context = "..." + context
        if end < len(para_text):
            context = context + "..."

        return context

    # ========================================================================
    # Generic edit methods (with optional tracking)
    # ========================================================================

    def insert(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        normalize_special_chars: bool = True,
        track: bool = False,
        fuzzy: float | dict[str, Any] | None = None,
    ) -> None:
        """Insert text after or before a specific location.

        This method searches for the anchor text in the document and inserts
        the new text either immediately after it or immediately before it.

        Args:
            text: The text to insert (supports markdown formatting: **bold**, *italic*,
                ++underline++, ~~strikethrough~~)
            after: Insert after this text (mutually exclusive with before)
            before: Insert before this text (mutually exclusive with after)
            author: Author for tracked changes (ignored if track=False)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            occurrence: Which occurrence(s): 1, 2, "first", "last", "all", or [1,3,5]
                (default: "first")
            regex: Treat anchor as regex pattern (default: False)
            normalize_special_chars: Auto-convert quotes for matching (default: True)
            track: If True, insert as tracked change; if False, silent insert
                (default: False)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'

        Raises:
            ValueError: If both 'after' and 'before' specified, or neither specified
            TextNotFoundError: If the anchor text is not found
            AmbiguousTextError: If multiple occurrences found and occurrence not specified
            re.error: If regex=True and the pattern is invalid
            ImportError: If fuzzy matching requested but rapidfuzz not installed

        Example:
            >>> doc.insert("new text", after="anchor")  # Untracked
            >>> doc.insert("new text", after="anchor", track=True)  # Tracked
            >>> doc.insert("new text", after="anchor", scope="footnote:1")  # In footnote
        """
        # Check if scope targets notes
        note_scope = parse_note_scope(scope) if isinstance(scope, str) else None

        if note_scope is not None and note_scope.note_id is not None:
            # Route to note-specific operation
            self._note_ops.insert_in_note(
                note_type=note_scope.scope_type,  # "footnote" or "endnote"
                note_id=note_scope.note_id,
                text=text,
                after=after,
                before=before,
                author=author,
                track=track,
            )
        else:
            self._tracked_ops.insert(
                text,
                after=after,
                before=before,
                author=author,
                scope=scope,
                occurrence=occurrence,
                regex=regex,
                normalize_special_chars=normalize_special_chars,
                track=track,
                fuzzy=fuzzy,
            )

    def delete(
        self,
        text: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        normalize_special_chars: bool = True,
        track: bool = False,
        fuzzy: float | dict[str, Any] | None = None,
    ) -> None:
        """Delete text from the document.

        This method searches for the specified text and removes it. When track=True,
        the deletion is shown as a tracked change.

        Args:
            text: The text to delete (or regex pattern if regex=True)
            author: Author for tracked changes (ignored if track=False)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            occurrence: Which occurrence(s): 1, 2, "first", "last", "all", or [1,3,5]
                (default: "first")
            regex: Treat text as regex pattern (default: False)
            normalize_special_chars: Auto-convert quotes for matching (default: True)
            track: If True, show as tracked deletion; if False, silent delete
                (default: False)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'

        Raises:
            TextNotFoundError: If the text is not found
            AmbiguousTextError: If multiple occurrences found and occurrence not specified
            re.error: If regex=True and the pattern is invalid
            ImportError: If fuzzy matching requested but rapidfuzz not installed

        Example:
            >>> doc.delete("old text")  # Untracked
            >>> doc.delete("old text", track=True)  # Tracked
            >>> doc.delete("obsolete", occurrence="all")  # Delete all occurrences
            >>> doc.delete("old", scope="footnote:1")  # Delete in footnote
        """
        # Check if scope targets notes
        note_scope = parse_note_scope(scope) if isinstance(scope, str) else None

        if note_scope is not None and note_scope.note_id is not None:
            # Route to note-specific operation
            self._note_ops.delete_in_note(
                note_type=note_scope.scope_type,  # "footnote" or "endnote"
                note_id=note_scope.note_id,
                text=text,
                author=author,
                track=track,
            )
        else:
            self._tracked_ops.delete(
                text,
                author=author,
                scope=scope,
                occurrence=occurrence,
                regex=regex,
                normalize_special_chars=normalize_special_chars,
                track=track,
                fuzzy=fuzzy,
            )

    def replace(
        self,
        find: str,
        replace_with: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        normalize_special_chars: bool = True,
        track: bool = False,
        show_context: bool = False,
        check_continuity: bool = False,
        context_chars: int = 50,
        fuzzy: float | dict[str, Any] | None = None,
        minimal: bool | None = None,
    ) -> None:
        """Find and replace text in the document.

        This method searches for text and replaces it with new text. When track=True,
        the operation shows both the deletion of the old text and insertion of the
        new text as tracked changes.

        When regex=True, the replacement string can use capture groups:
        - \\1, \\2, etc. for numbered groups
        - \\g<name> for named groups

        Args:
            find: Text or regex pattern to find
            replace_with: Replacement text (supports markdown: **bold**, *italic*, etc.)
            author: Author for tracked changes (ignored if track=False)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            occurrence: Which occurrence(s): 1, 2, "first", "last", "all", or [1,3,5]
                (default: "first")
            regex: Treat 'find' as regex pattern (default: False)
            normalize_special_chars: Auto-convert quotes for matching (default: True)
            track: If True, show as tracked change; if False, silent replace
                (default: False)
            show_context: Show text before/after the match for preview (default: False)
            check_continuity: Check if replacement may create sentence fragments
                (default: False)
            context_chars: Number of characters to show when show_context=True
                (default: 50)
            fuzzy: Fuzzy matching configuration:
                - None: Exact matching (default)
                - float: Similarity threshold (e.g., 0.9 for 90% similar)
                - dict: Full config with 'threshold', 'algorithm', 'normalize_whitespace'
            minimal: Use minimal word-level diffing for tracked changes (default: None,
                uses document's minimal_edits setting). Only applies when track=True.

        Raises:
            TextNotFoundError: If the 'find' text is not found
            AmbiguousTextError: If multiple occurrences found and occurrence not specified
            re.error: If regex=True and the pattern is invalid
            ImportError: If fuzzy matching requested but rapidfuzz not installed

        Warnings:
            ContinuityWarning: If check_continuity=True and potential fragment detected

        Example:
            >>> doc.replace("30 days", "45 days")  # Untracked
            >>> doc.replace("30 days", "45 days", track=True)  # Tracked
            >>> doc.replace("old", "new", occurrence="all")  # Replace all
            >>> doc.replace(r"(\\d+) days", r"\\1 business days", regex=True)
            >>> doc.replace("2020", "2024", scope="footnote:1")  # Replace in footnote
        """
        # Check if scope targets notes
        note_scope = parse_note_scope(scope) if isinstance(scope, str) else None

        if note_scope is not None and note_scope.note_id is not None:
            # Route to note-specific operation
            self._note_ops.replace_in_note(
                note_type=note_scope.scope_type,  # "footnote" or "endnote"
                note_id=note_scope.note_id,
                find=find,
                replace=replace_with,
                author=author,
                track=track,
            )
        else:
            self._tracked_ops.replace(
                find,
                replace_with,
                author=author,
                scope=scope,
                occurrence=occurrence,
                regex=regex,
                normalize_special_chars=normalize_special_chars,
                show_context=show_context,
                check_continuity=check_continuity,
                context_chars=context_chars,
                track=track,
                fuzzy=fuzzy,
                minimal=minimal,
            )

    def move(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
        source_scope: str | dict | Any | None = None,
        dest_scope: str | dict | Any | None = None,
        regex: bool = False,
        normalize_special_chars: bool = True,
        track: bool = False,
    ) -> None:
        """Move text to a new location.

        When track=True, creates linked move markers that show the text was
        relocated rather than deleted and re-added. This provides better context
        for document reviewers in Word.

        In Word's track changes view (track=True):
        - Source location shows text with strikethrough and "Moved" annotation
        - Destination shows text with underline and "Moved" annotation
        - Both locations are linked with matching move markers

        When track=False, simply deletes from source and inserts at destination
        without any tracking markers.

        Args:
            text: The text to move (or regex pattern if regex=True)
            after: Move to after this text (at destination)
            before: Move to before this text (at destination)
            author: Author for tracked changes (ignored if track=False)
            source_scope: Limit source text search scope
            dest_scope: Limit destination anchor search scope
            regex: Treat 'text' and anchor as regex patterns (default: False)
            normalize_special_chars: Auto-convert quotes for matching (default: True)
            track: If True, show as tracked move (linked markers); if False,
                move text without tracking (default: False)

        Raises:
            ValueError: If both 'after' and 'before' specified, or neither specified
            TextNotFoundError: If the source text or destination anchor is not found
            AmbiguousTextError: If multiple occurrences found
            re.error: If regex=True and a pattern is invalid

        Example:
            >>> doc.move("Section A", after="Table of Contents")  # Untracked
            >>> doc.move("Section A", after="Table of Contents", track=True)  # Tracked
        """
        self._tracked_ops.move(
            text,
            after=after,
            before=before,
            author=author,
            source_scope=source_scope,
            dest_scope=dest_scope,
            regex=regex,
            normalize_special_chars=normalize_special_chars,
            track=track,
        )

    # ========================================================================
    # Tracked change aliases (backwards compatible)
    # ========================================================================

    def insert_tracked(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        normalize_special_chars: bool = True,
        fuzzy: float | dict[str, Any] | None = None,
    ) -> None:
        """Insert text with tracked changes after or before a specific location.

        This is an alias for ``insert(..., track=True)``.
        See :meth:`insert` for full parameter documentation.

        Example:
            >>> doc.insert_tracked("new text", after="anchor")
        """
        self.insert(
            text,
            after=after,
            before=before,
            author=author,
            scope=scope,
            occurrence=occurrence,
            regex=regex,
            normalize_special_chars=normalize_special_chars,
            track=True,
            fuzzy=fuzzy,
        )

    def insert_image(
        self,
        image_path: str | Path,
        after: str | None = None,
        before: str | None = None,
        width_inches: float | None = None,
        height_inches: float | None = None,
        width_cm: float | None = None,
        height_cm: float | None = None,
        name: str | None = None,
        description: str = "",
        scope: str | dict | Any | None = None,
        regex: bool = False,
    ) -> None:
        """Insert an image into the document.

        This method searches for the anchor text in the document and inserts
        the image either immediately after it or immediately before it.

        If neither width nor height is specified, the image's native dimensions
        are used. If only one dimension is specified, the other is calculated
        to maintain aspect ratio. If PIL/Pillow is not installed, default
        dimensions of 2x2 inches are used.

        Args:
            image_path: Path to the image file (PNG, JPEG, GIF, etc.)
            after: The text to insert after (optional)
            before: The text to insert before (optional)
            width_inches: Width in inches (auto-calculated if not provided)
            height_inches: Height in inches (auto-calculated if not provided)
            width_cm: Width in centimeters (alternative to inches)
            height_cm: Height in centimeters (alternative to inches)
            name: Display name for the image (defaults to filename)
            description: Alt text description for accessibility
            scope: Limit search scope
            regex: Whether to treat anchor as regex pattern

        Raises:
            ValueError: If both 'after' and 'before' are specified, or neither
            TextNotFoundError: If the anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor are found
            FileNotFoundError: If the image file doesn't exist

        Example:
            >>> doc.insert_image("logo.png", after="Company Name")
            >>> doc.insert_image("chart.png", after="Figure 1:", width_inches=4.0)
        """
        self._image_ops.insert(
            image_path,
            after=after,
            before=before,
            width_inches=width_inches,
            height_inches=height_inches,
            width_cm=width_cm,
            height_cm=height_cm,
            name=name,
            description=description,
            scope=scope,
            regex=regex,
        )

    def insert_image_tracked(
        self,
        image_path: str | Path,
        after: str | None = None,
        before: str | None = None,
        width_inches: float | None = None,
        height_inches: float | None = None,
        width_cm: float | None = None,
        height_cm: float | None = None,
        name: str | None = None,
        description: str = "",
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
    ) -> None:
        """Insert an image with tracked changes.

        This method wraps the image insertion in a tracked change so it
        appears as an insertion in Word's review pane.

        If neither width nor height is specified, the image's native dimensions
        are used. If only one dimension is specified, the other is calculated
        to maintain aspect ratio. If PIL/Pillow is not installed, default
        dimensions of 2x2 inches are used.

        Args:
            image_path: Path to the image file (PNG, JPEG, GIF, etc.)
            after: The text to insert after (optional)
            before: The text to insert before (optional)
            width_inches: Width in inches (auto-calculated if not provided)
            height_inches: Height in inches (auto-calculated if not provided)
            width_cm: Width in centimeters (alternative to inches)
            height_cm: Height in centimeters (alternative to inches)
            name: Display name for the image (defaults to filename)
            description: Alt text description for accessibility
            author: Author for the tracked change (uses document author if None)
            scope: Limit search scope
            regex: Whether to treat anchor as regex pattern

        Raises:
            ValueError: If both 'after' and 'before' are specified, or neither
            TextNotFoundError: If the anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor are found
            FileNotFoundError: If the image file doesn't exist

        Example:
            >>> doc.insert_image_tracked("signature.png", after="Authorized By:")
            >>> doc.insert_image_tracked("stamp.png", after="Approved", author="Legal")
        """
        self._image_ops.insert_tracked(
            image_path,
            after=after,
            before=before,
            width_inches=width_inches,
            height_inches=height_inches,
            width_cm=width_cm,
            height_cm=height_cm,
            name=name,
            description=description,
            author=author,
            scope=scope,
            regex=regex,
        )

    def delete_tracked(
        self,
        text: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        normalize_special_chars: bool = True,
        fuzzy: float | dict[str, Any] | None = None,
    ) -> None:
        """Delete text with tracked changes.

        This is an alias for ``delete(..., track=True)``.
        See :meth:`delete` for full parameter documentation.

        Example:
            >>> doc.delete_tracked("old text")
        """
        self.delete(
            text,
            author=author,
            scope=scope,
            occurrence=occurrence,
            regex=regex,
            normalize_special_chars=normalize_special_chars,
            fuzzy=fuzzy,
            track=True,  # Always tracked for delete_tracked
        )

    def replace_tracked(
        self,
        find: str,
        replace: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | list[int] | str = "first",
        regex: bool = False,
        normalize_special_chars: bool = True,
        show_context: bool = False,
        check_continuity: bool = False,
        context_chars: int = 50,
        fuzzy: float | dict[str, Any] | None = None,
        minimal: bool | None = None,
    ) -> None:
        """Find and replace text with tracked changes.

        This is an alias for ``replace(..., track=True)``.
        See :meth:`replace` for full parameter documentation.

        Example:
            >>> doc.replace_tracked("30 days", "45 days")
        """
        self.replace(
            find,
            replace,
            author=author,
            scope=scope,
            occurrence=occurrence,
            regex=regex,
            normalize_special_chars=normalize_special_chars,
            track=True,
            show_context=show_context,
            check_continuity=check_continuity,
            context_chars=context_chars,
            fuzzy=fuzzy,
            minimal=minimal,
        )

    def move_tracked(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
        source_scope: str | dict | Any | None = None,
        dest_scope: str | dict | Any | None = None,
        regex: bool = False,
        normalize_special_chars: bool = True,
    ) -> None:
        """Move text to a new location with proper move tracking.

        This is an alias for ``move(..., track=True)``.
        See :meth:`move` for full parameter documentation.

        Example:
            >>> doc.move_tracked("Section A", after="Table of Contents")
        """
        self.move(
            text,
            after=after,
            before=before,
            author=author,
            source_scope=source_scope,
            dest_scope=dest_scope,
            regex=regex,
            normalize_special_chars=normalize_special_chars,
            track=True,
        )

    def normalize_currency(
        self,
        currency_symbol: str = "$",
        decimal_places: int = 2,
        thousands_separator: bool = True,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Normalize currency amounts to a consistent format with tracked changes.

        Finds various currency formats and normalizes them to a standard format.
        This reduces manual regex work and prevents formatting inconsistencies.

        Detected formats:
            - $100, $100.0 → $100.00
            - $1000 → $1,000.00 (if thousands_separator=True)
            - $1,000 → $1,000.00

        Args:
            currency_symbol: The currency symbol to use (default: "$")
            decimal_places: Number of decimal places (default: 2)
            thousands_separator: Whether to include thousands separators (default: True)
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})

        Returns:
            Number of currency amounts normalized

        Example:
            >>> # Normalize all $ amounts to $X,XXX.XX format
            >>> count = doc.normalize_currency()
            >>>
            >>> # Normalize to £X.XX without thousands separator
            >>> count = doc.normalize_currency("£", thousands_separator=False)
        """
        return self._pattern_ops.normalize_currency(
            currency_symbol=currency_symbol,
            decimal_places=decimal_places,
            thousands_separator=thousands_separator,
            author=author,
            scope=scope,
        )

    def normalize_dates(
        self,
        to_format: str = "%B %d, %Y",
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Normalize dates to a consistent format with tracked changes.

        Automatically detects common date formats and converts them to the target format.
        This prevents manual regex work and ensures date consistency.

        Detected formats:
            - MM/DD/YYYY (e.g., 12/08/2025)
            - M/D/YYYY (e.g., 1/8/2025)
            - YYYY-MM-DD (e.g., 2025-12-08)
            - Month DD, YYYY (e.g., December 08, 2025 or Dec 08, 2025)
            - DD Month YYYY (e.g., 08 December 2025)

        Args:
            to_format: Python datetime format string for output (default: "%B %d, %Y")
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})

        Returns:
            Number of dates normalized

        Example:
            >>> # Convert all dates to "December 08, 2025" format
            >>> count = doc.normalize_dates()
            >>>
            >>> # Convert all dates to ISO format
            >>> count = doc.normalize_dates("%Y-%m-%d")
        """
        return self._pattern_ops.normalize_dates(
            to_format=to_format,
            author=author,
            scope=scope,
        )

    def update_section_references(
        self,
        old_number: str,
        new_number: str,
        section_word: str = "Section",
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Update section/article references with tracked changes.

        Finds references like "Section 2.1" and updates them to "Section 3.1".
        Prevents manual regex errors when renumbering document sections.

        Args:
            old_number: Old section number (e.g., "2.1")
            new_number: New section number (e.g., "3.1")
            section_word: Word used for sections (default: "Section",
                could be "Article", "Clause", etc.)
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text",
                dict={"contains": "text"})

        Returns:
            Number of references updated

        Example:
            >>> # Update all "Section 2.1" references to "Section 3.1"
            >>> count = doc.update_section_references("2.1", "3.1")
            >>>
            >>> # Update article references
            >>> count = doc.update_section_references("5", "6", section_word="Article")
        """
        return self._pattern_ops.update_section_references(
            old_number=old_number,
            new_number=new_number,
            section_word=section_word,
            author=author,
            scope=scope,
        )

    def apply_style(
        self,
        find: str,
        style: str,
        scope: str | dict | Any | None = None,
        regex: bool = False,
    ) -> int:
        """Apply a paragraph style to paragraphs containing specific text.

        Changes the style of paragraphs that contain the search text.
        This is useful for programmatically formatting document sections.

        Args:
            find: Text to search for (or regex pattern if regex=True)
            style: Paragraph style name (e.g., 'Heading1', 'Normal', 'Quote')
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            regex: Whether to treat 'find' as a regex pattern (default: False)

        Returns:
            Number of paragraphs whose style was changed

        Example:
            >>> # Make all paragraphs containing "Section" into headings
            >>> count = doc.apply_style("Section", "Heading1")
            >>>
            >>> # Apply quote style to paragraphs with specific text
            >>> count = doc.apply_style("As stated in", "Quote")
        """
        return self._format_ops.apply_style(find, style, scope=scope, regex=regex)

    def format_text(
        self,
        find: str,
        bold: bool | None = None,
        italic: bool | None = None,
        color: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
    ) -> int:
        """Apply text formatting (bold, italic, color) to specific text.

        Finds text and applies formatting to the runs containing it.
        This allows surgical formatting changes without affecting surrounding text.

        Args:
            find: Text to search for (or regex pattern if regex=True)
            bold: Set bold formatting (True/False/None to leave unchanged)
            italic: Set italic formatting (True/False/None to leave unchanged)
            color: Set text color as hex (e.g., "FF0000" for red, None to leave unchanged)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            regex: Whether to treat 'find' as a regex pattern (default: False)

        Returns:
            Number of text occurrences formatted

        Example:
            >>> # Make all occurrences of "IMPORTANT" bold and red
            >>> count = doc.format_text("IMPORTANT", bold=True, color="FF0000")
            >>>
            >>> # Make section references italic
            >>> count = doc.format_text(r"Section \\d+\\.\\d+", italic=True, regex=True)
        """
        return self._format_ops.format_text(
            find, bold=bold, italic=italic, color=color, scope=scope, regex=regex
        )

    def format_tracked(
        self,
        text: str,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | str | None = None,
        strikethrough: bool | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        color: str | None = None,
        highlight: str | None = None,
        superscript: bool | None = None,
        subscript: bool | None = None,
        small_caps: bool | None = None,
        all_caps: bool | None = None,
        scope: str | dict | Any | None = None,
        occurrence: int | str = "first",
        author: str | None = None,
        normalize_special_chars: bool = True,
    ) -> FormatResult:
        """Apply character formatting to text with tracked changes.

        This method finds text and applies formatting changes that are tracked
        as revisions in Word. The previous formatting state is preserved in
        <w:rPrChange> elements, allowing users to accept or reject the
        formatting changes in Word.

        Args:
            text: The text to format (found via text search)
            bold: Set bold on (True), off (False), or leave unchanged (None)
            italic: Set italic on/off/unchanged
            underline: Set underline on/off/unchanged, or underline style name
            strikethrough: Set strikethrough on/off/unchanged
            font_name: Set font family name
            font_size: Set font size in points
            color: Set text color as hex "#RRGGBB" or "auto"
            highlight: Set highlight color name (e.g., "yellow", "green")
            superscript: Set superscript on/off/unchanged
            subscript: Set subscript on/off/unchanged
            small_caps: Set small caps on/off/unchanged
            all_caps: Set all caps on/off/unchanged
            scope: Limit search to specific paragraphs/sections
            occurrence: Which occurrence to format: 1, 2, "first", "last", or "all"
            author: Override default author for this change
            normalize_special_chars: Auto-convert straight quotes to smart quotes
                for matching (default: True)

        Returns:
            FormatResult with details of the formatting applied

        Raises:
            TextNotFoundError: If text is not found
            AmbiguousTextError: If multiple matches and occurrence not specified

        Example:
            >>> doc.format_tracked("IMPORTANT", bold=True, color="#FF0000")
            >>> doc.format_tracked("Section 2.1", italic=True, scope="section:Introduction")
            >>> doc.format_tracked("Note:", underline=True, font_size=14)
        """
        return self._format_ops.format_tracked(
            text,
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            font_name=font_name,
            font_size=font_size,
            color=color,
            highlight=highlight,
            superscript=superscript,
            subscript=subscript,
            small_caps=small_caps,
            all_caps=all_caps,
            scope=scope,
            occurrence=occurrence,
            author=author,
            normalize_special_chars=normalize_special_chars,
        )

    def format_paragraph_tracked(
        self,
        *,
        containing: str | None = None,
        starting_with: str | None = None,
        ending_with: str | None = None,
        index: int | None = None,
        alignment: str | None = None,
        spacing_before: float | None = None,
        spacing_after: float | None = None,
        line_spacing: float | None = None,
        indent_left: float | None = None,
        indent_right: float | None = None,
        indent_first_line: float | None = None,
        indent_hanging: float | None = None,
        scope: str | dict | Any | None = None,
        author: str | None = None,
    ) -> FormatResult:
        """Apply paragraph formatting with tracked changes.

        This method finds a paragraph and applies formatting changes that are
        tracked as revisions in Word. The previous formatting state is preserved
        in <w:pPrChange> elements.

        Args:
            containing: Find paragraph containing this text
            starting_with: Find paragraph starting with this text
            ending_with: Find paragraph ending with this text
            index: Target paragraph by index (0-based)
            alignment: Set paragraph alignment ("left", "center", "right", "justify")
            spacing_before: Set space before paragraph (points)
            spacing_after: Set space after paragraph (points)
            line_spacing: Set line spacing multiplier (e.g., 1.0, 1.5, 2.0)
            indent_left: Set left indent (inches)
            indent_right: Set right indent (inches)
            indent_first_line: Set first line indent (inches)
            indent_hanging: Set hanging indent (inches)
            scope: Limit search to specific sections
            author: Override default author for this change

        Returns:
            FormatResult with details of the formatting applied

        Raises:
            TextNotFoundError: If no matching paragraph found
            ValueError: If no targeting parameter or formatting parameter specified

        Example:
            >>> doc.format_paragraph_tracked(containing="WHEREAS", alignment="center")
            >>> doc.format_paragraph_tracked(index=0, spacing_after=12)
        """
        return self._format_ops.format_paragraph_tracked(
            containing=containing,
            starting_with=starting_with,
            ending_with=ending_with,
            index=index,
            alignment=alignment,
            spacing_before=spacing_before,
            spacing_after=spacing_after,
            line_spacing=line_spacing,
            indent_left=indent_left,
            indent_right=indent_right,
            indent_first_line=indent_first_line,
            indent_hanging=indent_hanging,
            scope=scope,
            author=author,
        )

    def copy_format(
        self,
        from_text: str,
        to_text: str,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Copy formatting from one text to another.

        Finds the source text, extracts its formatting (bold, italic, color, etc.),
        and applies the same formatting to the target text.

        Args:
            from_text: Source text to copy formatting from
            to_text: Target text to apply formatting to
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})

        Returns:
            Number of target occurrences formatted

        Example:
            >>> # Copy formatting from headers to make matching text look the same
            >>> count = doc.copy_format("Chapter 1", "Chapter 2")
        """
        return self._format_ops.copy_format(from_text, to_text, scope=scope)

    def insert_paragraph(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        style: str | None = None,
        track: bool = True,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> "Paragraph":
        """Insert a complete new paragraph with tracked changes.

        Args:
            text: Text content for the new paragraph
            after: Text to search for as insertion point (insert after this)
            before: Text to search for as insertion point (insert before this)
            style: Paragraph style (e.g., 'Normal', 'Heading1')
            track: Whether to track this insertion (default True)
            author: Optional author override (uses document author if None)
            scope: Limit search scope for anchor text

        Returns:
            The created Paragraph object

        Raises:
            ValueError: If neither 'after' nor 'before' is specified, or both are
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
        """
        return self._section_ops.insert_paragraph(
            text=text,
            after=after,
            before=before,
            style=style,
            track=track,
            author=author,
            scope=scope,
        )

    def insert_paragraphs(
        self,
        texts: list[str],
        after: str | None = None,
        before: str | None = None,
        style: str | None = None,
        track: bool = True,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> list["Paragraph"]:
        """Insert multiple paragraphs with tracked changes.

        This is more efficient than calling insert_paragraph() multiple times
        as it maintains proper ordering and positioning.

        Args:
            texts: List of text content for new paragraphs
            after: Text to search for as insertion point (insert after this)
            before: Text to search for as insertion point (insert before this)
            style: Paragraph style for all paragraphs (e.g., 'Normal', 'Heading1')
            track: Whether to track these insertions (default True)
            author: Optional author override (uses document author if None)
            scope: Limit search scope for anchor text

        Returns:
            List of created Paragraph objects

        Raises:
            ValueError: If neither 'after' nor 'before' is specified, or both are
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
        """
        return self._section_ops.insert_paragraphs(
            texts=texts,
            after=after,
            before=before,
            style=style,
            track=track,
            author=author,
            scope=scope,
        )

    def delete_section(
        self,
        heading: str,
        track: bool = True,
        update_toc: bool = False,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> "Section":
        """Delete an entire section by heading text.

        Args:
            heading: Heading text of section to delete
            track: Delete as tracked change (default True)
            update_toc: No-op, kept for API compatibility. TOC updates require
                opening the document in Word.
            author: Author name for tracked changes
            scope: Limit search scope

        Returns:
            Section object representing the deleted section

        Raises:
            TextNotFoundError: If heading not found
            AmbiguousTextError: If multiple sections match

        Examples:
            >>> doc.delete_section("Methods", track=True)
            >>> doc.delete_section("Outdated Section", track=False)
        """
        return self._section_ops.delete_section(
            heading=heading,
            track=track,
            update_toc=update_toc,
            author=author,
            scope=scope,
        )

    def delete_paragraph_tracked(
        self,
        containing: str | None = None,
        paragraph: "Paragraph | None" = None,
        paragraph_index: int | None = None,
        occurrence: int | list[int] | str | None = None,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> "Paragraph | list[Paragraph]":
        """Delete an entire paragraph with tracked changes.

        Marks the paragraph content as deleted (strikethrough) and also marks
        the paragraph mark as deleted. When the tracked change is accepted in
        Word, the paragraph cleanly merges with the following paragraph,
        leaving no empty lines behind.

        Args:
            containing: Text to search for to identify the paragraph
            paragraph: Paragraph object to delete directly
            paragraph_index: Index of paragraph to delete (0-based)
            occurrence: Which occurrence(s) to delete when multiple paragraphs
                match. Options: int (1-indexed), "first", "last", "all", or
                list of ints like [1, 3]. Only applies with 'containing' param.
            author: Author name for tracked changes
            scope: Limit search scope for 'containing' parameter

        Returns:
            The deleted Paragraph object, or list of Paragraphs if occurrence
            was "all" or a list

        Raises:
            ValueError: If none of containing/paragraph/paragraph_index provided,
                or if multiple are provided, or if occurrence is out of range
            TextNotFoundError: If containing text not found
            AmbiguousTextError: If containing text matches multiple paragraphs
                and occurrence not specified
            IndexError: If paragraph_index is out of range

        Examples:
            >>> # Delete paragraph containing specific text
            >>> doc.delete_paragraph_tracked(containing="Some citation text")

            >>> # Delete by index
            >>> doc.delete_paragraph_tracked(paragraph_index=5)

            >>> # Delete paragraph object directly
            >>> para = doc.paragraphs[5]
            >>> doc.delete_paragraph_tracked(paragraph=para)

            >>> # Delete specific occurrence when text matches multiple paragraphs
            >>> doc.delete_paragraph_tracked(containing="citation", occurrence=1)
            >>> doc.delete_paragraph_tracked(containing="citation", occurrence="last")

            >>> # Delete all matching paragraphs
            >>> deleted = doc.delete_paragraph_tracked(containing="TODO", occurrence="all")
            >>> print(f"Deleted {len(deleted)} paragraphs")
        """
        return self._section_ops.delete_paragraph_tracked(
            containing=containing,
            paragraph=paragraph,
            paragraph_index=paragraph_index,
            occurrence=occurrence,
            author=author,
            scope=scope,
        )

    def _insert_after_match(self, match: TextSpan, insertion_element: etree._Element) -> None:
        """Insert XML element(s) after a matched text span.
        Delegates to TrackedChangeOperations.
        """
        self._tracked_ops._insert_after_match(match, insertion_element)

    def _insert_before_match(self, match: TextSpan, insertion_element: etree._Element) -> None:
        """Insert XML element(s) before a matched text span.
        Delegates to TrackedChangeOperations.
        """
        self._tracked_ops._insert_before_match(match, insertion_element)

    def _replace_match_with_element(
        self, match: TextSpan, replacement_element: etree._Element
    ) -> None:
        """Replace matched text with a single XML element.
        Delegates to TrackedChangeOperations.
        """
        self._tracked_ops._replace_match_with_element(match, replacement_element)

    def _replace_match_with_elements(
        self, match: TextSpan, replacement_elements: list[etree._Element]
    ) -> None:
        """Replace matched text with multiple XML elements.
        Delegates to TrackedChangeOperations.
        """
        self._tracked_ops._replace_match_with_elements(match, replacement_elements)

    def _split_and_replace_in_run(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_element: Any,
    ) -> None:
        """Split a run and replace a portion with a new element.
        Delegates to TrackedChangeOperations.
        """
        self._tracked_ops._split_and_replace_in_run(
            paragraph, run, start_offset, end_offset, replacement_element
        )

    def _split_and_replace_in_run_multiple(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_elements: list[Any],
    ) -> None:
        """Split a run and replace a portion with multiple new elements.
        Delegates to TrackedChangeOperations.
        """
        self._tracked_ops._split_and_replace_in_run_multiple(
            paragraph, run, start_offset, end_offset, replacement_elements
        )

    def accept_all_changes(self) -> None:
        """Accept all tracked changes in the document.

        This removes all tracked change markup:
        - <w:ins> elements are unwrapped (content kept, wrapper removed)
        - <w:del> elements are completely removed (deleted content discarded)
        - <w:rPrChange> elements are removed (current formatting kept)
        - <w:pPrChange> elements are removed (current formatting kept)

        This is typically used as a preprocessing step before making new edits.
        """
        self._change_mgmt.accept_all()

    # Helper methods

    def _get_paragraph_text(self, para: Any) -> str:
        """Extract text from a paragraph element.

        Args:
            para: A <w:p> XML element

        Returns:
            Plain text content of the paragraph
        """
        text_elements = para.findall(f".//{{{WORD_NAMESPACE}}}t")
        return "".join(elem.text or "" for elem in text_elements)

    # Accept/Reject by type

    def accept_insertions(self) -> int:
        """Accept all tracked insertions in the document.

        Unwraps all <w:ins> elements, keeping the inserted content.

        Returns:
            Number of insertions accepted
        """
        return self._change_mgmt.accept_insertions()

    def reject_insertions(self) -> int:
        """Reject all tracked insertions in the document.

        Removes all <w:ins> elements and their content.

        Returns:
            Number of insertions rejected
        """
        return self._change_mgmt.reject_insertions()

    def accept_deletions(self) -> int:
        """Accept all tracked deletions in the document.

        Removes all <w:del> elements (keeps text deleted).

        Returns:
            Number of deletions accepted
        """
        return self._change_mgmt.accept_deletions()

    def reject_deletions(self) -> int:
        """Reject all tracked deletions in the document.

        Unwraps all <w:del> elements, restoring the deleted content.
        Converts w:delText back to w:t.

        Returns:
            Number of deletions rejected
        """
        return self._change_mgmt.reject_deletions()

    def accept_format_changes(self) -> int:
        """Accept all tracked formatting changes in the document.

        Removes all <w:rPrChange> and <w:pPrChange> elements, keeping the
        current formatting as-is.

        Returns:
            Number of formatting changes accepted
        """
        return self._change_mgmt.accept_format_changes()

    def reject_format_changes(self) -> int:
        """Reject all tracked formatting changes in the document.

        Restores the previous formatting from <w:rPrChange> and <w:pPrChange>
        elements, then removes the change tracking elements.

        Returns:
            Number of formatting changes rejected
        """
        return self._change_mgmt.reject_format_changes()

    def reject_all_changes(self) -> None:
        """Reject all tracked changes in the document.

        This removes all tracked change markup by reverting to previous state:
        - <w:ins> elements and their content are removed
        - <w:del> elements are unwrapped (deleted content restored)
        - <w:rPrChange> elements restore previous formatting
        - <w:pPrChange> elements restore previous formatting
        """
        self._change_mgmt.reject_all()

    # Accept/Reject by change ID

    def accept_change(self, change_id: str | int) -> None:
        """Accept a specific tracked change by its ID.

        Args:
            change_id: The change ID (w:id attribute value)

        Raises:
            ValueError: If no change with the given ID is found

        Example:
            >>> doc.accept_change("5")
        """
        self._change_mgmt.accept_change(change_id)

    def reject_change(self, change_id: str | int) -> None:
        """Reject a specific tracked change by its ID.

        Args:
            change_id: The change ID (w:id attribute value)

        Raises:
            ValueError: If no change with the given ID is found

        Example:
            >>> doc.reject_change("5")
        """
        self._change_mgmt.reject_change(change_id)

    # Accept/Reject by author

    def accept_by_author(self, author: str) -> int:
        """Accept all tracked changes by a specific author.

        Args:
            author: The author name (w:author attribute value)

        Returns:
            Number of changes accepted

        Example:
            >>> count = doc.accept_by_author("John Doe")
            >>> print(f"Accepted {count} changes from John Doe")
        """
        return self._change_mgmt.accept_by_author(author)

    def reject_by_author(self, author: str) -> int:
        """Reject all tracked changes by a specific author.

        Args:
            author: The author name (w:author attribute value)

        Returns:
            Number of changes rejected

        Example:
            >>> count = doc.reject_by_author("John Doe")
            >>> print(f"Rejected {count} changes from John Doe")
        """
        return self._change_mgmt.reject_by_author(author)

    # List and query tracked changes

    def get_tracked_changes(
        self,
        change_type: str | None = None,
        author: str | None = None,
    ) -> list["TrackedChange"]:
        """Get all tracked changes in the document.

        Returns a list of TrackedChange objects representing insertions, deletions,
        moves, and formatting changes with their metadata.

        Args:
            change_type: Optional filter by change type. Valid values:
                         "insertion", "deletion", "move_from", "move_to",
                         "format_run", "format_paragraph", or None for all.
            author: Optional filter by author name.

        Returns:
            List of TrackedChange objects matching the criteria.

        Example:
            >>> # Get all changes
            >>> changes = doc.get_tracked_changes()
            >>> for change in changes:
            ...     print(f"{change.id}: {change.change_type.value} by {change.author}")
            >>>
            >>> # Get only insertions
            >>> insertions = doc.get_tracked_changes(change_type="insertion")
            >>>
            >>> # Get changes by specific author
            >>> johns_changes = doc.get_tracked_changes(author="John Doe")
        """
        from python_docx_redline.models.tracked_change import ChangeType, TrackedChange

        changes: list[TrackedChange] = []

        # Map string type names to ChangeType enum values
        type_map = {
            "insertion": ChangeType.INSERTION,
            "deletion": ChangeType.DELETION,
            "move_from": ChangeType.MOVE_FROM,
            "move_to": ChangeType.MOVE_TO,
            "format_run": ChangeType.FORMAT_RUN,
            "format_paragraph": ChangeType.FORMAT_PARAGRAPH,
        }

        # Validate change_type if provided
        filter_type: ChangeType | None = None
        if change_type is not None:
            if change_type not in type_map:
                valid_types = ", ".join(sorted(type_map.keys()))
                raise ValueError(f"Invalid change_type '{change_type}'. Valid types: {valid_types}")
            filter_type = type_map[change_type]

        # Collect insertions (w:ins)
        if filter_type is None or filter_type == ChangeType.INSERTION:
            for ins in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"):
                change = TrackedChange.from_element(ins, ChangeType.INSERTION, self)
                if author is None or change.author == author:
                    changes.append(change)

        # Collect deletions (w:del)
        if filter_type is None or filter_type == ChangeType.DELETION:
            for del_elem in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"):
                change = TrackedChange.from_element(del_elem, ChangeType.DELETION, self)
                if author is None or change.author == author:
                    changes.append(change)

        # Collect move sources (w:moveFrom)
        if filter_type is None or filter_type == ChangeType.MOVE_FROM:
            for move_from in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}moveFrom"):
                change = TrackedChange.from_element(move_from, ChangeType.MOVE_FROM, self)
                if author is None or change.author == author:
                    changes.append(change)

        # Collect move destinations (w:moveTo)
        if filter_type is None or filter_type == ChangeType.MOVE_TO:
            for move_to in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}moveTo"):
                change = TrackedChange.from_element(move_to, ChangeType.MOVE_TO, self)
                if author is None or change.author == author:
                    changes.append(change)

        # Collect run property changes (w:rPrChange)
        if filter_type is None or filter_type == ChangeType.FORMAT_RUN:
            for rpr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange"):
                change = TrackedChange.from_element(rpr_change, ChangeType.FORMAT_RUN, self)
                if author is None or change.author == author:
                    changes.append(change)

        # Collect paragraph property changes (w:pPrChange)
        if filter_type is None or filter_type == ChangeType.FORMAT_PARAGRAPH:
            for ppr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange"):
                change = TrackedChange.from_element(ppr_change, ChangeType.FORMAT_PARAGRAPH, self)
                if author is None or change.author == author:
                    changes.append(change)

        return changes

    def accept_changes(
        self,
        change_type: str | None = None,
        author: str | None = None,
    ) -> int:
        """Accept multiple tracked changes matching the given criteria.

        This is a bulk operation that accepts all changes matching the filters.
        If no filters are provided, accepts ALL tracked changes (equivalent to
        accept_all_tracked_changes()).

        Args:
            change_type: Optional filter by change type. Valid values:
                         "insertion", "deletion", "format_run", "format_paragraph",
                         or None for all types.
            author: Optional filter by author name.

        Returns:
            Number of changes accepted.

        Example:
            >>> # Accept all insertions
            >>> count = doc.accept_changes(change_type="insertion")
            >>> print(f"Accepted {count} insertions")
            >>>
            >>> # Accept all changes by a specific author
            >>> count = doc.accept_changes(author="Legal Team")
            >>> print(f"Accepted {count} changes from Legal Team")
            >>>
            >>> # Accept only insertions by a specific author
            >>> count = doc.accept_changes(change_type="insertion", author="John Doe")
        """
        return self._change_mgmt.accept_changes(change_type=change_type, author=author)

    def reject_changes(
        self,
        change_type: str | None = None,
        author: str | None = None,
    ) -> int:
        """Reject multiple tracked changes matching the given criteria.

        This is a bulk operation that rejects all changes matching the filters.
        If no filters are provided, rejects ALL tracked changes (equivalent to
        reject_all_tracked_changes()).

        Args:
            change_type: Optional filter by change type. Valid values:
                         "insertion", "deletion", "format_run", "format_paragraph",
                         or None for all types.
            author: Optional filter by author name.

        Returns:
            Number of changes rejected.

        Example:
            >>> # Reject all deletions
            >>> count = doc.reject_changes(change_type="deletion")
            >>> print(f"Rejected {count} deletions")
            >>>
            >>> # Reject all changes by a specific author
            >>> count = doc.reject_changes(author="Unauthorized User")
            >>> print(f"Rejected {count} changes from Unauthorized User")
            >>>
            >>> # Reject only deletions by a specific author
            >>> count = doc.reject_changes(change_type="deletion", author="John Doe")
        """
        return self._change_mgmt.reject_changes(change_type=change_type, author=author)

    # Edit group operations

    @property
    def _edit_groups(self) -> EditGroupRegistry:
        """Get the EditGroupRegistry instance (lazy initialization)."""
        if not hasattr(self, "_edit_groups_instance"):
            self._edit_groups_instance = EditGroupRegistry()
        return self._edit_groups_instance

    @contextmanager
    def edit_group(self, name: str) -> Iterator[None]:
        """Context manager for grouping related edits.

        All tracked changes made within this context will be associated with
        the named group. This enables batch operations like rejecting all
        changes in a group at once.

        Args:
            name: Unique name for this edit group

        Yields:
            None

        Raises:
            ValueError: If another group is already active
            ValueError: If a group with this name already exists

        Example:
            >>> with doc.edit_group('condensing round 1'):
            ...     doc.replace_tracked('long text', 'short')
            ...     doc.replace_tracked('another section', 'condensed')
            >>> # Later, reject all changes in that group
            >>> doc.reject_edit_group('condensing round 1')
        """
        self._edit_groups.start_group(name)
        try:
            yield
        finally:
            self._edit_groups.end_group()

    def reject_edit_group(self, group_name: str) -> int:
        """Reject all tracked changes in an edit group.

        This method finds all change IDs associated with the named group
        and rejects them in reverse order (to properly handle nested changes).

        Args:
            group_name: Name of the edit group to reject

        Returns:
            Number of changes successfully rejected

        Raises:
            ValueError: If no group with the given name exists

        Example:
            >>> with doc.edit_group('round1'):
            ...     doc.replace_tracked('old', 'new')
            ...     doc.insert_tracked(' extra', after='text')
            >>> # Later, reject all changes from round1
            >>> count = doc.reject_edit_group('round1')
            >>> print(f"Rejected {count} changes")
        """
        ids = self._edit_groups.get_group_ids(group_name)
        count = 0

        # Reject in reverse order to handle nested elements properly
        for change_id in reversed(ids):
            try:
                self._change_mgmt.reject_change(change_id)
                count += 1
            except ValueError:
                # Change may have already been rejected or no longer exists
                pass

        self._edit_groups.mark_rejected(group_name)
        return count

    def accept_edit_group(self, group_name: str) -> int:
        """Accept all tracked changes in an edit group.

        This method finds all change IDs associated with the named group
        and accepts them in reverse order (to properly handle nested changes).

        Args:
            group_name: Name of the edit group to accept

        Returns:
            Number of changes successfully accepted

        Raises:
            ValueError: If no group with the given name exists

        Example:
            >>> with doc.edit_group('round1'):
            ...     doc.replace_tracked('old', 'new')
            ...     doc.insert_tracked(' extra', after='text')
            >>> # Later, accept all changes from round1
            >>> count = doc.accept_edit_group('round1')
            >>> print(f"Accepted {count} changes")
        """
        ids = self._edit_groups.get_group_ids(group_name)
        count = 0

        # Accept in reverse order to handle nested elements properly
        for change_id in reversed(ids):
            try:
                self._change_mgmt.accept_change(change_id)
                count += 1
            except ValueError:
                # Change may have already been accepted or no longer exists
                pass

        return count

    # Accept/Reject by text content

    def reject_changes_containing(
        self,
        text: str,
        *,
        change_type: str | None = None,
        author: str | None = None,
        match_case: bool = False,
        regex: bool = False,
    ) -> int:
        """Reject tracked changes containing specified text.

        Searches the text content of tracked changes and rejects those
        containing the specified text. Useful for targeted rejection of
        changes by content rather than by index/ID.

        Args:
            text: String to search for in change content.
            change_type: Optional filter by change type. Valid values:
                         "insertion", "deletion", "format_run", "format_paragraph",
                         or None for all types.
            author: Optional filter by author name.
            match_case: If False (default), case-insensitive search.
            regex: If True, treat text as a regex pattern.

        Returns:
            Number of changes rejected.

        Example:
            >>> # Reject deletions containing "VersusLaw"
            >>> count = doc.reject_changes_containing("VersusLaw", change_type="deletion")
            >>> print(f"Rejected {count} changes containing 'VersusLaw'")
            >>>
            >>> # Reject all changes containing "confidential" by any author
            >>> count = doc.reject_changes_containing("confidential")
            >>>
            >>> # Reject changes matching a regex pattern
            >>> count = doc.reject_changes_containing(r"Section \\d+", regex=True)
        """
        return self._change_mgmt.reject_changes_containing(
            text, change_type=change_type, author=author, match_case=match_case, regex=regex
        )

    def accept_changes_containing(
        self,
        text: str,
        *,
        change_type: str | None = None,
        author: str | None = None,
        match_case: bool = False,
        regex: bool = False,
    ) -> int:
        """Accept tracked changes containing specified text.

        Searches the text content of tracked changes and accepts those
        containing the specified text. Useful for targeted acceptance of
        changes by content rather than by index/ID.

        Args:
            text: String to search for in change content.
            change_type: Optional filter by change type. Valid values:
                         "insertion", "deletion", "format_run", "format_paragraph",
                         or None for all types.
            author: Optional filter by author name.
            match_case: If False (default), case-insensitive search.
            regex: If True, treat text as a regex pattern.

        Returns:
            Number of changes accepted.

        Example:
            >>> # Accept insertions containing "approved"
            >>> count = doc.accept_changes_containing("approved", change_type="insertion")
            >>> print(f"Accepted {count} changes containing 'approved'")
            >>>
            >>> # Accept all changes containing specific clause text
            >>> count = doc.accept_changes_containing("Section 2.1")
            >>>
            >>> # Accept changes matching a regex pattern
            >>> count = doc.accept_changes_containing(r"v\\d+\\.\\d+", regex=True)
        """
        return self._change_mgmt.accept_changes_containing(
            text, change_type=change_type, author=author, match_case=match_case, regex=regex
        )

    @property
    def tracked_changes(self) -> list["TrackedChange"]:
        """Get all tracked changes as a read-only property.

        Convenience property equivalent to get_tracked_changes() with no filters.

        Returns:
            List of all TrackedChange objects in the document.

        Example:
            >>> for change in doc.tracked_changes:
            ...     print(f"{change.id}: {change.change_type.value}")
        """
        return self.get_tracked_changes()

    @property
    def comparison_stats(self) -> ComparisonStats:
        """Get statistics about tracked changes in the document.

        Provides counts of insertions, deletions, moves, and format changes.
        Useful for summarizing the results of a document comparison.

        Returns:
            ComparisonStats object with counts of each change type.

        Example:
            >>> redline = compare_documents("v1.docx", "v2.docx")
            >>> stats = redline.comparison_stats
            >>> print(f"Insertions: {stats.insertions}")
            >>> print(f"Deletions: {stats.deletions}")
            >>> print(f"Total: {stats.total}")
            >>> print(stats)  # "3 insertions, 2 deletions"
        """
        from .models.tracked_change import ChangeType

        changes = self.tracked_changes
        insertions = sum(1 for c in changes if c.change_type == ChangeType.INSERTION)
        deletions = sum(1 for c in changes if c.change_type == ChangeType.DELETION)
        moves = sum(
            1 for c in changes if c.change_type in (ChangeType.MOVE_FROM, ChangeType.MOVE_TO)
        )
        format_changes = sum(
            1
            for c in changes
            if c.change_type in (ChangeType.FORMAT_RUN, ChangeType.FORMAT_PARAGRAPH)
        )

        return ComparisonStats(
            insertions=insertions,
            deletions=deletions,
            moves=moves,
            format_changes=format_changes,
        )

    def export_changes_json(
        self,
        include_context: bool = True,
        context_chars: int = 50,
        indent: int | None = 2,
    ) -> str:
        """Export all tracked changes to JSON format.

        Creates a JSON representation of all tracked changes with metadata,
        suitable for integration with external tools or further processing.

        Args:
            include_context: Whether to include surrounding text context
            context_chars: Number of context characters to include on each side
            indent: JSON indentation level, or None for compact output

        Returns:
            JSON string containing all tracked changes

        Example:
            >>> json_data = doc.export_changes_json()
            >>> import json
            >>> changes = json.loads(json_data)
            >>> print(f"Found {changes['total_changes']} changes")
        """
        from .export import export_changes_json

        return export_changes_json(
            self,
            include_context=include_context,
            context_chars=context_chars,
            indent=indent,
        )

    def export_changes_markdown(
        self,
        include_context: bool = True,
        context_chars: int = 50,
        group_by: str | None = None,
    ) -> str:
        """Export tracked changes to Markdown format.

        Creates a human-readable Markdown document showing all tracked changes
        with optional context and grouping. Useful for code reviews or
        generating documentation.

        Args:
            include_context: Whether to include surrounding text context
            context_chars: Number of context characters to include on each side
            group_by: How to group changes: "author", "type", or None for no grouping

        Returns:
            Markdown formatted string with all tracked changes

        Example:
            >>> md = doc.export_changes_markdown(group_by="author")
            >>> with open("changes.md", "w") as f:
            ...     f.write(md)
        """
        from .export import export_changes_markdown

        return export_changes_markdown(
            self,
            include_context=include_context,
            context_chars=context_chars,
            group_by=group_by,  # type: ignore[arg-type]
        )

    def export_changes_html(
        self,
        include_context: bool = True,
        context_chars: int = 50,
        group_by: str | None = None,
        inline_styles: bool = True,
    ) -> str:
        """Export tracked changes to HTML format.

        Creates an HTML document with a code-review style visualization of
        tracked changes, similar to diff views in version control systems.
        Includes syntax highlighting for insertions, deletions, and other
        change types.

        Args:
            include_context: Whether to include surrounding text context
            context_chars: Number of context characters to include on each side
            group_by: How to group changes: "author", "type", or None for no grouping
            inline_styles: Whether to include inline CSS styles (True) or just classes (False)

        Returns:
            HTML formatted string with all tracked changes

        Example:
            >>> html_content = doc.export_changes_html(group_by="author")
            >>> with open("changes.html", "w") as f:
            ...     f.write(html_content)
        """
        from .export import export_changes_html

        return export_changes_html(
            self,
            include_context=include_context,
            context_chars=context_chars,
            group_by=group_by,  # type: ignore[arg-type]
            inline_styles=inline_styles,
        )

    def to_criticmarkup(self, include_comments: bool = True) -> str:
        """Export document with tracked changes to CriticMarkup markdown.

        Converts the document to plain text with CriticMarkup annotations for
        tracked changes and comments. This format is useful for:
        - Reviewing changes in plain text editors
        - Version control (text diffs work well)
        - AI agent workflows

        Tracked changes are converted as follows:
        - Insertions → {++text++}
        - Deletions → {--text--}
        - Comments → {>>comment text<<}

        Args:
            include_comments: Whether to include comments (default: True)

        Returns:
            Markdown string with CriticMarkup annotations

        Example:
            >>> doc = Document("contract_with_changes.docx")
            >>> markdown = doc.to_criticmarkup()
            >>> print(markdown)
            The parties agree to {--30--}{++45++} day payment terms.

        See Also:
            - apply_criticmarkup(): Import CriticMarkup changes back to DOCX
            - http://criticmarkup.com/ for syntax reference
        """
        from .criticmarkup import docx_to_criticmarkup

        return docx_to_criticmarkup(self, include_comments=include_comments)

    def apply_criticmarkup(
        self,
        markup_text: str,
        author: str | None = None,
        stop_on_error: bool = False,
        track: bool = True,
    ) -> "ApplyResult":
        """Apply CriticMarkup changes to the document with optional tracking.

        Parses CriticMarkup syntax from the input text and applies each operation
        to the document. When track=True (default), changes are shown as tracked
        changes. When track=False, changes are applied silently.

        CriticMarkup syntax is converted to Word changes:
        - {++text++} → insertion (tracked or silent based on track param)
        - {--text--} → deletion (tracked or silent based on track param)
        - {~~old~>new~~} → replacement (tracked or silent based on track param)
        - {>>comment<<} → Word comment (always visible)
        - {==text=={>>comment<<}} → Comment attached to text (always visible)

        Args:
            markup_text: Markdown text with CriticMarkup annotations
            author: Author for tracked changes (uses document default if None)
            stop_on_error: If True, stop on first error. If False, continue
                processing remaining operations.
            track: If True, show as tracked changes; if False, apply silently (default: True)

        Returns:
            ApplyResult with success/failure counts and error details

        Example:
            >>> doc = Document("contract.docx")
            >>> with open("reviewed.md") as f:
            ...     markup = f.read()
            >>> result = doc.apply_criticmarkup(markup, author="Reviewer")
            >>> print(f"Applied {result.successful}/{result.total} changes")
            >>> doc.save("contract_reviewed.docx")
            >>>
            >>> # Apply silently without tracked changes
            >>> result = doc.apply_criticmarkup(markup, track=False)

        See Also:
            - to_criticmarkup(): Export document to CriticMarkup format
            - http://criticmarkup.com/ for syntax reference
        """
        from .criticmarkup import apply_criticmarkup

        return apply_criticmarkup(self, markup_text, author, stop_on_error, track=track)

    def generate_change_report(
        self,
        format: str = "html",
        include_context: bool = True,
        context_chars: int = 50,
        group_by: str | None = "author",
        title: str | None = None,
    ) -> str:
        """Generate a comprehensive change report in the specified format.

        This is a convenience method that generates a formatted report of all
        tracked changes in the document. The report includes a summary with
        statistics and optionally groups changes by author or type.

        Args:
            format: Output format: "html", "markdown", or "json"
            include_context: Whether to include surrounding text context
            context_chars: Number of context characters to include on each side
            group_by: How to group changes: "author", "type", or None for no grouping
            title: Optional custom title for the report

        Returns:
            Formatted report string in the specified format

        Example:
            >>> # Generate HTML report grouped by author
            >>> report = doc.generate_change_report(format="html", group_by="author")
            >>> with open("report.html", "w") as f:
            ...     f.write(report)
            >>>
            >>> # Generate Markdown report grouped by type
            >>> md_report = doc.generate_change_report(format="markdown", group_by="type")
            >>>
            >>> # Generate JSON report
            >>> json_report = doc.generate_change_report(format="json")
        """
        from .export import generate_change_report

        return generate_change_report(
            self,
            format=format,  # type: ignore[arg-type]
            include_context=include_context,
            context_chars=context_chars,
            group_by=group_by,  # type: ignore[arg-type]
            title=title,
        )

    def delete_all_comments(self) -> None:
        """Delete all comments from the document.

        This removes all comment-related elements:
        - <w:commentRangeStart> - Comment range start markers
        - <w:commentRangeEnd> - Comment range end markers
        - <w:commentReference> - Comment reference markers
        - Runs containing comment references
        - word/comments.xml and related files (commentsExtended.xml, etc.)
        - Comment relationships from document.xml.rels
        - Comment content types from [Content_Types].xml

        This ensures the document package is valid OOXML with no orphaned comments.
        This is typically used as a preprocessing step before making new edits.
        """
        self._comment_ops.delete_all()

    def add_comment(
        self,
        text: str,
        on: str | None = None,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
        initials: str | None = None,
        reply_to: "Comment | str | int | None" = None,
        occurrence: int | list[int] | str | None = None,
    ) -> "Comment | list[Comment]":
        """Add a comment to the document on specified text or as a reply.

        This method can either add a new top-level comment on text in the
        document, or add a reply to an existing comment.

        Args:
            text: The comment text content
            on: The text to annotate (or regex pattern if regex=True).
                Required for new comments, ignored for replies.
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            regex: Whether to treat 'on' as a regex pattern (default: False)
            initials: Author initials (auto-generated from author if None)
            reply_to: Comment to reply to (Comment object, comment ID str/int, or None)
            occurrence: Which occurrence(s) to target when multiple matches exist:
                - None (default): Error if multiple matches (current behavior)
                - "first" or 1: Target first match only
                - "last": Target last match only
                - "all": Add comment to all matches (returns list of Comments)
                - int (e.g., 2, 3): Target nth match (1-indexed)
                - list[int] (e.g., [1, 3]): Target specific matches (1-indexed)

        Returns:
            The created Comment object, or list of Comment objects if occurrence="all"
            or a list of indices is provided

        Raises:
            TextNotFoundError: If the target text is not found (new comments only)
            AmbiguousTextError: If multiple occurrences found and occurrence not specified
            ValueError: If neither 'on' nor 'reply_to' is provided, if occurrence
                        is out of range, or if reply_to references a non-existent comment
            re.error: If regex=True and the pattern is invalid

        Example:
            >>> doc = Document("contract.docx")
            >>> # Add a top-level comment
            >>> comment = doc.add_comment(
            ...     "Please review this section",
            ...     on="Section 2.1",
            ...     author="Reviewer"
            ... )
            >>> # Add a reply
            >>> reply = doc.add_comment(
            ...     "I've reviewed it, looks good",
            ...     reply_to=comment,
            ...     author="Author"
            ... )
            >>> # Target a specific occurrence when text appears multiple times
            >>> doc.add_comment(
            ...     "Check this instance",
            ...     on="Evidence Gaps:",
            ...     occurrence=1  # First occurrence
            ... )
            >>> # Add comment to all occurrences
            >>> comments = doc.add_comment(
            ...     "Review needed",
            ...     on="TODO",
            ...     occurrence="all"
            ... )
            >>> doc.save("contract_reviewed.docx")
        """
        return self._comment_ops.add(
            text=text,
            on=on,
            author=author,
            scope=scope,
            regex=regex,
            initials=initials,
            reply_to=reply_to,
            occurrence=occurrence,
        )

    # Delegation methods for Comment model backward compatibility
    # These methods delegate to CommentOperations to avoid code duplication

    def _get_comment_ex(self, para_id: str) -> etree._Element | None:
        """Get the commentEx element for a given paraId.

        Delegates to CommentOperations.

        Args:
            para_id: The paraId to look up

        Returns:
            The w15:commentEx element or None if not found
        """
        return self._comment_ops._get_comment_ex(para_id)

    def _set_comment_resolved(self, para_id: str, resolved: bool) -> None:
        """Set the resolved status for a comment.

        Delegates to CommentOperations.

        Args:
            para_id: The paraId of the comment
            resolved: True to mark as resolved, False for unresolved
        """
        self._comment_ops._set_comment_resolved(para_id, resolved)

    def _delete_comment(self, comment_id: str, para_id: str | None) -> None:
        """Delete a comment by ID.

        Delegates to CommentOperations.

        Args:
            comment_id: The comment ID to delete
            para_id: The paraId of the comment (for commentsExtended cleanup)
        """
        self._comment_ops._delete_comment(comment_id, para_id)

    # Table operations

    def update_cell(
        self,
        row: int,
        col: int,
        new_text: str,
        *,
        table_index: int = 0,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
    ) -> None:
        """Update a table cell's content with optional tracked changes.

        Args:
            row: 0-based row index
            col: 0-based column index
            new_text: New text content for the cell
            table_index: Which table to modify (default: 0 = first table)
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)

        Raises:
            IndexError: If table_index, row, or col is out of range

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.update_cell(0, 1, "Updated Value", table_index=0, track=True)
            >>> doc.save("contract_updated.docx")
        """
        tables = self.tables
        if table_index < 0 or table_index >= len(tables):
            raise IndexError(f"Table index {table_index} out of range (0-{len(tables) - 1})")

        table = tables[table_index]
        cell = table.get_cell(row, col)

        if track:
            # For tracked changes, we need to replace the cell's content
            # by wrapping old content in <w:del> and adding new content in <w:ins>
            from datetime import datetime, timezone

            author_name = author if author is not None else self.author
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._xml_generator.next_change_id

            # Get all paragraphs in the cell
            paragraphs = cell.paragraphs

            if paragraphs:
                # Delete existing content with tracking
                for para in paragraphs:
                    # Wrap all runs in deletion
                    for run in list(para.element.findall(f"{{{WORD_NAMESPACE}}}r")):
                        # Create deletion wrapper
                        del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
                        del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                        del_elem.set(f"{{{WORD_NAMESPACE}}}author", author_name)
                        del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

                        # Convert w:t to w:delText
                        for t_elem in run.findall(f"{{{WORD_NAMESPACE}}}t"):
                            t_elem.tag = f"{{{WORD_NAMESPACE}}}delText"

                        # Move run into deletion
                        run_parent = run.getparent()
                        run_index = list(run_parent).index(run)
                        run_parent.remove(run)
                        del_elem.append(run)
                        run_parent.insert(run_index, del_elem)

                        change_id += 1

                # Insert new content with tracking in first paragraph
                first_para = paragraphs[0]
                ins_elem = etree.Element(f"{{{WORD_NAMESPACE}}}ins")
                ins_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                ins_elem.set(f"{{{WORD_NAMESPACE}}}author", author_name)
                ins_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
                run = etree.SubElement(ins_elem, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = new_text
                first_para.element.append(ins_elem)
                self._xml_generator.next_change_id = change_id + 1
            else:
                # No paragraphs - create new paragraph with tracked insertion
                para = etree.SubElement(cell.element, f"{{{WORD_NAMESPACE}}}p")
                ins_elem = etree.Element(f"{{{WORD_NAMESPACE}}}ins")
                ins_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                ins_elem.set(f"{{{WORD_NAMESPACE}}}author", author_name)
                ins_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
                run = etree.SubElement(ins_elem, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = new_text
                para.append(ins_elem)
                self._xml_generator.next_change_id = change_id + 1
        else:
            # Untracked - just replace the text
            cell.text = new_text

    def replace_in_table(
        self,
        old_text: str,
        new_text: str,
        *,
        table_index: int | None = None,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
        regex: bool = False,
        case_sensitive: bool = True,
    ) -> int:
        """Replace text in table cells with optional tracked changes.

        Args:
            old_text: Text to find (or regex pattern if regex=True)
            new_text: Replacement text
            table_index: Specific table index, or None for all tables
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)
            regex: Whether old_text is a regex pattern (default: False)
            case_sensitive: Whether search is case sensitive (default: True)

        Returns:
            Number of replacements made

        Example:
            >>> doc = Document("contract.docx")
            >>> count = doc.replace_in_table("OLD", "NEW", track=True)
            >>> print(f"Replaced {count} occurrences")
        """
        return self._table_ops.replace_text(
            old_text,
            new_text,
            table_index=table_index,
            track=track,
            author=author,
            regex=regex,
            case_sensitive=case_sensitive,
        )

    def insert_table_row(
        self,
        after_row: int | str,
        cells: list[str],
        *,
        table_index: int = 0,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
    ) -> "TableRow":
        """Insert a new table row with optional tracked changes.

        Args:
            after_row: Row index (int) or text to find in a row (str)
            cells: List of text content for each cell in the new row
            table_index: Which table to modify (default: 0 = first table)
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)

        Returns:
            The newly created TableRow object

        Raises:
            IndexError: If table_index is out of range
            ValueError: If after_row text is not found or is ambiguous
            ValueError: If number of cells doesn't match table column count

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.insert_table_row(
            ...     after_row="Total:",
            ...     cells=["New Item", "$1,000", "$2,000"],
            ...     track=True
            ... )
        """
        return self._table_ops.insert_row(
            after_row, cells, table_index=table_index, track=track, author=author
        )

    def delete_table_row(
        self,
        row: int | str,
        *,
        table_index: int = 0,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
    ) -> "TableRow":
        """Delete a table row with optional tracked changes.

        Args:
            row: Row index (int) or text to find in a row (str)
            table_index: Which table to modify (default: 0 = first table)
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)

        Returns:
            The deleted TableRow object

        Raises:
            IndexError: If table_index or row index is out of range
            ValueError: If row text is not found or is ambiguous

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.delete_table_row(row=5, track=True)
        """
        return self._table_ops.delete_row(row, table_index=table_index, track=track, author=author)

    def insert_table_column(
        self,
        after_column: int | str,
        cells: list[str],
        *,
        table_index: int = 0,
        header: str | None = None,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
    ) -> None:
        """Insert a new table column with optional tracked changes.

        Columns in OOXML are implicit - they are derived from cells in rows.
        This method inserts a new cell into each row at the specified position.

        Args:
            after_column: Column index (int) or text to find in a column (str).
                          Use -1 to insert before the first column.
            cells: List of text content for each cell in the new column.
                   Length must match the number of rows (excluding header if provided).
            table_index: Which table to modify (default: 0 = first table)
            header: Optional header text for the first row. If provided, cells list
                    should have one fewer element (for data rows only).
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)

        Raises:
            IndexError: If table_index is out of range
            ValueError: If after_column text is not found or is ambiguous
            ValueError: If number of cells doesn't match expected row count

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.insert_table_column(
            ...     after_column=1,
            ...     cells=["A", "B", "C"],
            ...     header="New Column",
            ...     track=True
            ... )
        """
        self._table_ops.insert_column(
            after_column,
            cells,
            table_index=table_index,
            header=header,
            track=track,
            author=author,
        )

    def delete_table_column(
        self,
        column: int | str,
        *,
        table_index: int = 0,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
    ) -> None:
        """Delete a table column with optional tracked changes.

        Columns in OOXML are implicit - they are derived from cells in rows.
        This method removes or marks cells at the specified column position in each row.

        Args:
            column: Column index (int) or text to find in a column (str)
            table_index: Which table to modify (default: 0 = first table)
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)

        Raises:
            IndexError: If table_index or column index is out of range
            ValueError: If column text is not found or is ambiguous

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.delete_table_column(column=2, track=True)
        """
        self._table_ops.delete_column(column, table_index=table_index, track=track, author=author)

    def validate(self, verbose: bool = False) -> bool:
        """Run full OOXML validation on the current document.

        This runs the same comprehensive validation suite as save() but without
        actually saving the document. Useful for checking document validity before
        proceeding with operations or for debugging validation issues.

        Args:
            verbose: Whether to print verbose validation output (default: False)

        Returns:
            True if document passes all validation checks

        Raises:
            ValidationError: If document validation fails. Error includes detailed
                list of validation issues for bug reporting.
        """
        if not self._is_zip or not self._temp_dir:
            raise ValidationError(
                "Cannot validate: document was not loaded from a .docx file. "
                "Validation only works on full .docx documents."
            )

        # Write the current XML state to temp directory
        document_xml = self._temp_dir / "word" / "document.xml"
        self.xml_tree.write(
            str(document_xml),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=False,
        )

        # Run full validation
        from .validation_docx import DOCXSchemaValidator

        validator = DOCXSchemaValidator(
            unpacked_dir=self._temp_dir,
            original_file=self.path,
            verbose=verbose,
        )

        if not validator.validate():
            raise ValidationError(
                "Document validation failed. Please report this as a bug. "
                "See validation errors above for details."
            )

        return True

    def save(
        self,
        output_path: str | Path | None = None,
        validate: bool = True,
        strict_validation: bool = False,
    ) -> None:
        """Save the document to a file.

        Validates the document structure before saving to ensure OOXML compliance
        and prevent broken Word files in production.

        Args:
            output_path: Path to save the document. If None, saves to original path.
                        For in-memory documents (loaded from bytes), output_path is required.
            validate: Whether to run Python-based OOXML validation (default: True).
                     Validation is strongly recommended to catch errors before production.
                     Set to False for in-memory documents without an original file.
            strict_validation: Whether to also run full OOXML spec validation using
                     the external OOXML-Validator tool (default: False). Only runs if
                     the validator is installed. Set to True for maximum confidence
                     in OOXML compliance. See: https://github.com/mikeebowen/OOXML-Validator

        Raises:
            ValidationError: If document validation fails. Error includes detailed
                list of validation issues for bug reporting.
            ValueError: If output_path is not provided for in-memory documents.
        """
        if output_path is None:
            if self.path is None:
                raise ValueError(
                    "output_path is required for in-memory documents. "
                    "Use doc.save(path) or doc.save_to_bytes() instead."
                )
            output_path = self.path
        else:
            output_path = Path(output_path)

        try:
            if self._package is not None:
                # Save style changes if the StyleManager was accessed and modified
                if hasattr(self, "_style_manager_instance"):
                    self._style_manager_instance.save()

                # Write the modified XML back to the package
                self._package.set_part("word/document.xml", self.xml_root)

                # Validate the full document structure before creating ZIP
                # This catches OOXML spec violations that would produce broken Word files
                if validate:
                    from .validation_docx import DOCXSchemaValidator

                    validator = DOCXSchemaValidator(
                        unpacked_dir=self._package.temp_dir,
                        original_file=self.path,
                        verbose=False,
                    )
                    if not validator.validate():
                        # Collect all validation errors for detailed bug reporting
                        error_list = (
                            validator.all_errors if hasattr(validator, "all_errors") else []
                        )
                        raise ValidationError(
                            "Document validation failed. Please report this as a bug. "
                            "See validation errors above for details.",
                            errors=error_list,
                        )

                # Save the package to the output path
                self._package.save(output_path)

                # Run strict OOXML validation if requested
                if strict_validation:
                    from .ooxml_validator import (
                        OOXMLValidationError,
                        is_ooxml_validator_available,
                        validate_with_ooxml_validator,
                    )

                    if is_ooxml_validator_available():
                        errors = validate_with_ooxml_validator(output_path)
                        if errors:
                            raise OOXMLValidationError(
                                f"Strict OOXML validation failed with {len(errors)} error(s)",
                                errors,
                            )
                    else:
                        logger.warning(
                            "strict_validation requested but OOXML-Validator not available. "
                            "Install from https://github.com/mikeebowen/OOXML-Validator"
                        )
            else:
                # Save XML directly (raw XML file, not a package)
                self.xml_tree.write(
                    str(output_path),
                    encoding="utf-8",
                    xml_declaration=True,
                    pretty_print=False,
                )

        except ValidationError:
            # Re-raise ValidationError with all its attributes intact
            raise
        except Exception as e:
            # Check if it's an OOXMLValidationError (from strict validation)
            if type(e).__name__ == "OOXMLValidationError":
                raise
            raise ValidationError(f"Failed to save document: {e}") from e

    def save_to_bytes(
        self,
        validate: bool = True,
        strict_validation: bool = False,
    ) -> bytes:
        """Save the document to bytes (in-memory).

        This is useful for:
        - Passing documents between libraries without filesystem
        - Storing documents in databases
        - Sending documents over network

        Args:
            validate: Whether to run Python-based OOXML validation (default: True).
                     Set to False for in-memory documents without an original file,
                     as validation compares against the original.
            strict_validation: Whether to also run full OOXML spec validation using
                     the external OOXML-Validator tool (default: False). Only runs if
                     the validator is installed. Note: requires writing to a temp file.

        Returns:
            bytes: The complete .docx file as bytes

        Raises:
            ValidationError: If validation fails

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.replace_tracked("old", "new")
            >>> doc_bytes = doc.save_to_bytes()
            >>> # Store in database, send over network, etc.
        """
        if self._package is None:
            raise ValidationError("save_to_bytes only supported for .docx files")

        try:
            # Save style changes if the StyleManager was accessed and modified
            if hasattr(self, "_style_manager_instance"):
                self._style_manager_instance.save()

            # Write the modified XML back to the package
            self._package.set_part("word/document.xml", self.xml_root)

            # Validate if requested and we have an original file to compare against
            if validate and self.path is not None:
                from .validation_docx import DOCXSchemaValidator

                validator = DOCXSchemaValidator(
                    unpacked_dir=self._package.temp_dir,
                    original_file=self.path,
                    verbose=False,
                )
                if not validator.validate():
                    error_list = validator.all_errors if hasattr(validator, "all_errors") else []
                    raise ValidationError(
                        "Document validation failed. Please report this as a bug. "
                        "See validation errors above for details.",
                        errors=error_list,
                    )

            # Save to bytes using the package
            doc_bytes = self._package.save_to_bytes()

            # Run strict OOXML validation if requested (requires temp file)
            if strict_validation:
                import tempfile

                from .ooxml_validator import (
                    OOXMLValidationError,
                    is_ooxml_validator_available,
                    validate_with_ooxml_validator,
                )

                if is_ooxml_validator_available():
                    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
                        temp_path = Path(f.name)
                        f.write(doc_bytes)

                    try:
                        errors = validate_with_ooxml_validator(temp_path)
                        if errors:
                            raise OOXMLValidationError(
                                f"Strict OOXML validation failed with {len(errors)} error(s)",
                                errors,
                            )
                    finally:
                        temp_path.unlink(missing_ok=True)
                else:
                    logger.warning(
                        "strict_validation requested but OOXML-Validator not available. "
                        "Install from https://github.com/mikeebowen/OOXML-Validator"
                    )

            return doc_bytes

        except ValidationError:
            raise
        except Exception as e:
            # Check if it's an OOXMLValidationError (from strict validation)
            if type(e).__name__ == "OOXMLValidationError":
                raise
            raise ValidationError(f"Failed to save document to bytes: {e}") from e

    def render_to_images(
        self,
        output_dir: str | Path | None = None,
        dpi: int = 150,
        prefix: str = "page",
        timeout: int = 120,
    ) -> list[Path]:
        """Render document pages to PNG images.

        Uses LibreOffice to render the document with full formatting,
        including tracked changes shown visually (strikethrough, underlines).

        This is useful for AI agents to visually inspect document layout
        and see how tracked changes appear in the rendered document.

        Args:
            output_dir: Directory for output images. If None, uses a temp directory.
            dpi: Resolution in dots per inch (default: 150)
            prefix: Filename prefix for images (default: "page")
            timeout: Timeout in seconds for each conversion step (default: 120)

        Returns:
            List of Path objects for generated PNG files (page-1.png, page-2.png, ...)

        Raises:
            RuntimeError: If LibreOffice or pdftoppm is not available

        Example:
            >>> doc = Document("contract.docx")
            >>> images = doc.render_to_images(dpi=200)
            >>> for img in images:
            ...     print(f"Page: {img}")

        Note:
            Requires LibreOffice and poppler-utils to be installed:
            - macOS: brew install --cask libreoffice && brew install poppler
            - Linux: sudo apt install libreoffice poppler-utils
        """
        from .rendering import render_document_to_images

        # If we have a source path, use it directly
        # (Save first to capture any unsaved changes)
        if self.path is not None:
            # Save any pending changes
            self.save(self.path)
            return render_document_to_images(
                self.path,
                output_dir=output_dir,
                dpi=dpi,
                prefix=prefix,
                timeout=timeout,
            )

        # Otherwise, save to a temp file first
        import tempfile

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = Path(tmp.name)

        try:
            self.save(tmp_path)
            return render_document_to_images(
                tmp_path,
                output_dir=output_dir,
                dpi=dpi,
                prefix=prefix,
                timeout=timeout,
            )
        finally:
            tmp_path.unlink(missing_ok=True)

    def apply_edits(
        self,
        edits: list[dict[str, Any]],
        stop_on_error: bool = False,
        default_track: bool = False,
    ) -> list[EditResult]:
        """Apply multiple edits in sequence.

        This method processes a list of edit specifications and applies each one
        in order. Each edit is a dictionary specifying the edit type and parameters.

        Args:
            edits: List of edit dictionaries with keys:
                - type: Edit operation ("insert", "delete", "replace",
                    "insert_tracked", "replace_tracked", "delete_tracked")
                - track: Optional boolean to control tracking per-edit
                - Other parameters specific to the edit type
            stop_on_error: If True, stop processing on first error
            default_track: Default value for 'track' if not specified per-edit
                (default: False). Note: *_tracked operations always track regardless.

        Returns:
            List of EditResult objects, one per edit

        Example:
            >>> edits = [
            ...     {
            ...         "type": "insert",
            ...         "text": "new text",
            ...         "after": "anchor",
            ...         "track": True,  # This edit is tracked
            ...     },
            ...     {
            ...         "type": "replace",
            ...         "find": "old",
            ...         "replace": "new",
            ...         # Uses default_track value
            ...     }
            ... ]
            >>> results = doc.apply_edits(edits, default_track=False)
            >>> print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
        """
        return self._batch_ops.apply_edits(
            edits, stop_on_error=stop_on_error, default_track=default_track
        )

    def apply_edit_file(
        self,
        path: str | Path,
        format: str = "yaml",
        stop_on_error: bool = False,
        default_track: bool | None = None,
    ) -> list[EditResult]:
        """Apply edits from a YAML or JSON file.

        Loads edit specifications from a file and applies them using apply_edits().
        The file should contain an 'edits' key with a list of edit dictionaries.

        Args:
            path: Path to the edit specification file
            format: File format - "yaml" or "json" (default: "yaml")
            stop_on_error: If True, stop processing on first error
            default_track: Default value for 'track' if not specified per-edit.
                If None, uses file's default_track value (or False if not set).
                If specified, overrides any default_track in the file.

        Returns:
            List of EditResult objects, one per edit

        Raises:
            ValidationError: If file cannot be parsed or has invalid format
            FileNotFoundError: If file does not exist

        Example YAML file:
            ```yaml
            default_track: false  # Global default for edits in this file

            edits:
              - type: insert
                text: "new text"
                after: "anchor"
                # Uses default_track: false

              - type: replace
                find: "old"
                replace: "new"
                track: true  # Override for this specific edit
            ```

        Example:
            >>> results = doc.apply_edit_file("edits.yaml")
            >>> print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
        """
        return self._batch_ops.apply_edit_file(
            path, format=format, stop_on_error=stop_on_error, default_track=default_track
        )

    def compare_to(
        self,
        modified: "Document",
        author: str | None = None,
        minimal_edits: bool = False,
    ) -> int:
        """Generate tracked changes by comparing this document to a modified version.

        This method compares the current document (original) to a modified document
        and generates tracked changes showing what was added, removed, or changed.
        The changes are applied to this document.

        The comparison operates at the paragraph level:
        - Paragraphs in modified but not in original → tracked insertions
        - Paragraphs in original but not in modified → tracked deletions
        - Paragraphs that changed → tracked deletion of old + insertion of new

        Args:
            modified: The modified Document to compare against
            author: Author name for the tracked changes (uses document default if None)
            minimal_edits: If True, use word-level diffs for 1:1 paragraph replacements
                instead of deleting/inserting entire paragraphs. This produces
                legal-style redlines where only the changed words are marked.
                (default: False)

        Returns:
            Number of changes made (insertions + deletions)

        Example:
            >>> original = Document("contract_v1.docx")
            >>> modified = Document("contract_v2.docx")
            >>> num_changes = original.compare_to(modified)
            >>> original.save("contract_redlined.docx")
            >>> print(f"Found {num_changes} changes")

            # For legal-style minimal diffs:
            >>> num_changes = original.compare_to(modified, minimal_edits=True)

        Note:
            - This modifies the current document in place
            - The comparison uses paragraph text content
            - Formatting changes within paragraphs are not tracked separately
            - For best results, compare documents with similar structure
            - When minimal_edits=True, whitespace-only changes are suppressed
              for readability, and paragraphs with existing tracked changes
              fall back to coarse replacement
        """
        return self._comparison_ops.compare_to(modified, author=author, minimal_edits=minimal_edits)

    # ========================================================================
    # FOOTNOTE / ENDNOTE METHODS
    # ========================================================================

    @property
    def footnotes(self) -> list["Footnote"]:
        """Get all footnotes in the document.

        Returns:
            List of Footnote objects
        """
        return self._note_ops.footnotes

    @property
    def endnotes(self) -> list["Endnote"]:
        """Get all endnotes in the document.

        Returns:
            List of Endnote objects
        """
        return self._note_ops.endnotes

    def find_orphaned_footnotes(self) -> list["OrphanedFootnote"]:
        """Find footnotes that have no reference in the document body.

        Orphaned footnotes occur when text containing footnote markers is deleted
        but the footnote content remains in footnotes.xml. This method detects
        these orphans by comparing footnote IDs in footnotes.xml against
        footnoteReference elements in document.xml.

        System footnotes (id=-1 for separator, id=0 for continuationSeparator)
        are excluded from the results.

        Returns:
            List of OrphanedFootnote objects with id and text content

        Example:
            >>> orphans = doc.find_orphaned_footnotes()
            >>> for orphan in orphans:
            ...     print(f"Orphaned footnote {orphan.id}: {orphan.text[:50]}...")
        """
        return self._note_ops.find_orphaned_footnotes()

    def find_orphaned_endnotes(self) -> list["OrphanedEndnote"]:
        """Find endnotes that have no reference in the document body.

        Orphaned endnotes occur when text containing endnote markers is deleted
        but the endnote content remains in endnotes.xml. This method detects
        these orphans by comparing endnote IDs in endnotes.xml against
        endnoteReference elements in document.xml.

        System endnotes (id=-1 for separator, id=0 for continuationSeparator)
        are excluded from the results.

        Returns:
            List of OrphanedEndnote objects with id and text content

        Example:
            >>> orphans = doc.find_orphaned_endnotes()
            >>> for orphan in orphans:
            ...     print(f"Orphaned endnote {orphan.id}: {orphan.text[:50]}...")
        """
        return self._note_ops.find_orphaned_endnotes()

    def insert_footnote(
        self,
        text: str,
        at: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Insert a footnote reference at specific text location.

        Args:
            text: The footnote text content
            at: Text to search for where footnote reference should be inserted
            author: Optional author (uses document author if None)
            scope: Optional scope to limit search

        Returns:
            The footnote ID

        Raises:
            TextNotFoundError: If 'at' text not found
            AmbiguousTextError: If multiple occurrences of 'at' text found

        Example:
            >>> doc.insert_footnote("See Smith (2020) for details", at="original study")
        """
        return self._note_ops.insert_footnote(text, at, author=author, scope=scope)

    def insert_endnote(
        self,
        text: str,
        at: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> int:
        """Insert an endnote reference at specific text location.

        Args:
            text: The endnote text content
            at: Text to search for where endnote reference should be inserted
            author: Optional author (uses document author if None)
            scope: Optional scope to limit search

        Returns:
            The endnote ID

        Raises:
            TextNotFoundError: If 'at' text not found
            AmbiguousTextError: If multiple occurrences of 'at' text found

        Example:
            >>> doc.insert_endnote("Additional details here", at="main conclusion")
        """
        return self._note_ops.insert_endnote(text, at, author=author, scope=scope)

    def get_footnote(self, note_id: str | int) -> "Footnote":
        """Get a specific footnote by ID.

        Args:
            note_id: The footnote ID to retrieve

        Returns:
            The Footnote object

        Raises:
            NoteNotFoundError: If the footnote ID is not found

        Example:
            >>> footnote = doc.get_footnote(1)
            >>> print(footnote.text)
        """
        return self._note_ops.get_footnote(note_id)

    def get_endnote(self, note_id: str | int) -> "Endnote":
        """Get a specific endnote by ID.

        Args:
            note_id: The endnote ID to retrieve

        Returns:
            The Endnote object

        Raises:
            NoteNotFoundError: If the endnote ID is not found

        Example:
            >>> endnote = doc.get_endnote(1)
            >>> print(endnote.text)
        """
        return self._note_ops.get_endnote(note_id)

    def delete_footnote(self, note_id: str | int, renumber: bool = True) -> None:
        """Delete a footnote by ID.

        This removes both the footnote content from footnotes.xml and the
        corresponding footnote reference from the document body.

        Args:
            note_id: The footnote ID to delete
            renumber: If True, renumber remaining footnotes sequentially (default)

        Raises:
            NoteNotFoundError: If the footnote ID is not found

        Example:
            >>> doc.delete_footnote(2)  # Delete footnote 2 and renumber
            >>> doc.delete_footnote(2, renumber=False)  # Delete without renumbering
        """
        self._note_ops.delete_footnote(note_id, renumber=renumber)

    def delete_endnote(self, note_id: str | int, renumber: bool = True) -> None:
        """Delete an endnote by ID.

        This removes both the endnote content from endnotes.xml and the
        corresponding endnote reference from the document body.

        Args:
            note_id: The endnote ID to delete
            renumber: If True, renumber remaining endnotes sequentially (default)

        Raises:
            NoteNotFoundError: If the endnote ID is not found

        Example:
            >>> doc.delete_endnote(2)  # Delete endnote 2 and renumber
            >>> doc.delete_endnote(2, renumber=False)  # Delete without renumbering
        """
        self._note_ops.delete_endnote(note_id, renumber=renumber)

    def edit_footnote(
        self,
        note_id: str | int,
        new_text: str,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Edit the text of a footnote.

        Args:
            note_id: The footnote ID to edit
            new_text: The new text content for the footnote
            track: If True, track the edit as a change (Phase 3 feature)
            author: Author name for tracked changes (uses document author if None)

        Raises:
            NoteNotFoundError: If the footnote ID is not found

        Example:
            >>> doc.edit_footnote(1, "Updated citation text")
        """
        self._note_ops.edit_footnote(note_id, new_text, track=track, author=author)

    def edit_endnote(
        self,
        note_id: str | int,
        new_text: str,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Edit the text of an endnote.

        Args:
            note_id: The endnote ID to edit
            new_text: The new text content for the endnote
            track: If True, track the edit as a change (Phase 3 feature)
            author: Author name for tracked changes (uses document author if None)

        Raises:
            NoteNotFoundError: If the endnote ID is not found

        Example:
            >>> doc.edit_endnote(1, "Updated citation text")
        """
        self._note_ops.edit_endnote(note_id, new_text, track=track, author=author)

    def insert_tracked_in_footnote(
        self,
        note_id: str | int,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
    ) -> None:
        """Insert text with tracked changes inside a footnote.

        This method searches for anchor text within the footnote and inserts
        new text as a tracked insertion (w:ins) either after or before it.

        Args:
            note_id: The footnote ID to edit
            text: The text to insert (supports markdown formatting)
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If both or neither of after/before specified
            NoteNotFoundError: If footnote not found
            TextNotFoundError: If anchor text not found in footnote
            AmbiguousTextError: If anchor text found multiple times

        Example:
            >>> doc.insert_tracked_in_footnote(1, " [updated]", after="citation")
        """
        self._note_ops.insert_tracked_in_footnote(
            note_id, text, after=after, before=before, author=author
        )

    def insert_tracked_in_endnote(
        self,
        note_id: str | int,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
    ) -> None:
        """Insert text with tracked changes inside an endnote.

        This method searches for anchor text within the endnote and inserts
        new text as a tracked insertion (w:ins) either after or before it.

        Args:
            note_id: The endnote ID to edit
            text: The text to insert (supports markdown formatting)
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            author: Optional author override (uses document author if None)

        Raises:
            ValueError: If both or neither of after/before specified
            NoteNotFoundError: If endnote not found
            TextNotFoundError: If anchor text not found in endnote
            AmbiguousTextError: If anchor text found multiple times

        Example:
            >>> doc.insert_tracked_in_endnote(1, " [see also]", after="reference")
        """
        self._note_ops.insert_tracked_in_endnote(
            note_id, text, after=after, before=before, author=author
        )

    def delete_tracked_in_footnote(
        self,
        note_id: str | int,
        text: str,
        author: str | None = None,
    ) -> None:
        """Delete text with tracked changes inside a footnote.

        This method searches for text within the footnote and marks it
        as a tracked deletion (w:del).

        Args:
            note_id: The footnote ID to edit
            text: The text to delete
            author: Optional author override (uses document author if None)

        Raises:
            NoteNotFoundError: If footnote not found
            TextNotFoundError: If text not found in footnote
            AmbiguousTextError: If text found multiple times

        Example:
            >>> doc.delete_tracked_in_footnote(1, "outdated reference")
        """
        self._note_ops.delete_tracked_in_footnote(note_id, text, author=author)

    def delete_tracked_in_endnote(
        self,
        note_id: str | int,
        text: str,
        author: str | None = None,
    ) -> None:
        """Delete text with tracked changes inside an endnote.

        This method searches for text within the endnote and marks it
        as a tracked deletion (w:del).

        Args:
            note_id: The endnote ID to edit
            text: The text to delete
            author: Optional author override (uses document author if None)

        Raises:
            NoteNotFoundError: If endnote not found
            TextNotFoundError: If text not found in endnote
            AmbiguousTextError: If text found multiple times

        Example:
            >>> doc.delete_tracked_in_endnote(1, "obsolete citation")
        """
        self._note_ops.delete_tracked_in_endnote(note_id, text, author=author)

    def replace_tracked_in_footnote(
        self,
        note_id: str | int,
        find: str,
        replace: str,
        author: str | None = None,
    ) -> None:
        """Replace text with tracked changes inside a footnote.

        This method searches for text within the footnote and replaces it
        showing both the deletion of old text and insertion of new text
        as tracked changes.

        Args:
            note_id: The footnote ID to edit
            find: The text to find and replace
            replace: The replacement text (supports markdown formatting)
            author: Optional author override (uses document author if None)

        Raises:
            NoteNotFoundError: If footnote not found
            TextNotFoundError: If find text not found in footnote
            AmbiguousTextError: If find text found multiple times

        Example:
            >>> doc.replace_tracked_in_footnote(1, "2020", "2024")
        """
        self._note_ops.replace_tracked_in_footnote(note_id, find, replace, author=author)

    def replace_tracked_in_endnote(
        self,
        note_id: str | int,
        find: str,
        replace: str,
        author: str | None = None,
    ) -> None:
        """Replace text with tracked changes inside an endnote.

        This method searches for text within the endnote and replaces it
        showing both the deletion of old text and insertion of new text
        as tracked changes.

        Args:
            note_id: The endnote ID to edit
            find: The text to find and replace
            replace: The replacement text (supports markdown formatting)
            author: Optional author override (uses document author if None)

        Raises:
            NoteNotFoundError: If endnote not found
            TextNotFoundError: If find text not found in endnote
            AmbiguousTextError: If find text found multiple times

        Example:
            >>> doc.replace_tracked_in_endnote(1, "ibid", "op. cit.")
        """
        self._note_ops.replace_tracked_in_endnote(note_id, find, replace, author=author)

    def merge_footnotes(
        self,
        footnote_ids: list[int],
        separator: str = "; ",
        keep_first: bool = True,
    ) -> int:
        """Merge multiple footnotes into one.

        Combines the content of multiple footnotes and removes the extras.
        Useful for cleaning up adjacent footnotes after text deletion,
        particularly in legal documents following Bluebook citation style.

        When deleting text between two footnotes, they may become adjacent
        (e.g., `[^15][^16]`). This method merges them into a single footnote
        with combined content.

        Args:
            footnote_ids: List of footnote IDs to merge (in order)
            separator: Text to insert between merged contents (default: "; ")
            keep_first: If True, keep first footnote and delete others.
                       If False, keep last footnote.

        Returns:
            ID of the remaining footnote

        Raises:
            ValueError: If fewer than 2 footnote IDs provided
            NoteNotFoundError: If any footnote ID is not found

        Example:
            >>> # Merge footnotes 15 and 16 into one
            >>> remaining_id = doc.merge_footnotes([15, 16], separator="; ")
            >>> print(f"Merged into footnote {remaining_id}")
            >>>
            >>> # Merge three footnotes, keeping the last one
            >>> remaining_id = doc.merge_footnotes([1, 2, 3], keep_first=False)
        """
        return self._note_ops.merge_footnotes(
            footnote_ids, separator=separator, keep_first=keep_first
        )

    def merge_endnotes(
        self,
        endnote_ids: list[int],
        separator: str = "; ",
        keep_first: bool = True,
    ) -> int:
        """Merge multiple endnotes into one.

        Combines the content of multiple endnotes and removes the extras.
        Useful for cleaning up adjacent endnotes after text deletion.

        Args:
            endnote_ids: List of endnote IDs to merge (in order)
            separator: Text to insert between merged contents (default: "; ")
            keep_first: If True, keep first endnote and delete others.
                       If False, keep last endnote.

        Returns:
            ID of the remaining endnote

        Raises:
            ValueError: If fewer than 2 endnote IDs provided
            NoteNotFoundError: If any endnote ID is not found

        Example:
            >>> # Merge endnotes 5 and 6 into one
            >>> remaining_id = doc.merge_endnotes([5, 6], separator="; ")
            >>> print(f"Merged into endnote {remaining_id}")
        """
        return self._note_ops.merge_endnotes(
            endnote_ids, separator=separator, keep_first=keep_first
        )

    # ========================================================================
    # HEADER / FOOTER METHODS
    # ========================================================================

    @property
    def headers(self) -> list["Header"]:
        """Get all headers in the document.

        Headers are linked via relationships in section properties (sectPr).
        Each section can have up to three headers: default, first, even.

        Returns:
            List of Header objects

        Example:
            >>> doc = Document("contract.docx")
            >>> for header in doc.headers:
            ...     print(f"{header.type}: {header.text}")
        """
        return self._header_footer_ops.headers

    @property
    def footers(self) -> list["Footer"]:
        """Get all footers in the document.

        Footers are linked via relationships in section properties (sectPr).
        Each section can have up to three footers: default, first, even.

        Returns:
            List of Footer objects

        Example:
            >>> doc = Document("contract.docx")
            >>> for footer in doc.footers:
            ...     print(f"{footer.type}: {footer.text}")
        """
        return self._header_footer_ops.footers

    def replace_in_header(
        self,
        find: str,
        replace: str,
        header_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        normalize_special_chars: bool = True,
        track: bool = True,
    ) -> None:
        """Replace text in a header with optional tracked changes.

        Args:
            find: Text to find
            replace: Replacement text
            header_type: Type of header ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat 'find' as a regex pattern
            normalize_special_chars: Auto-convert quotes for matching
            track: If True, show as tracked change; if False, silent replace (default: True)

        Raises:
            TextNotFoundError: If 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
            ValueError: If no header of the specified type exists

        Example:
            >>> doc.replace_in_header("Draft", "Final", header_type="default")
            >>> doc.replace_in_header("Draft", "Final", track=False)  # Silent replace
        """
        self._header_footer_ops.replace_in_header(
            find=find,
            replace=replace,
            header_type=header_type,
            author=author,
            regex=regex,
            normalize_special_chars=normalize_special_chars,
            track=track,
        )

    def replace_in_footer(
        self,
        find: str,
        replace: str,
        footer_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        normalize_special_chars: bool = True,
        track: bool = True,
    ) -> None:
        """Replace text in a footer with optional tracked changes.

        Args:
            find: Text to find
            replace: Replacement text
            footer_type: Type of footer ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat 'find' as a regex pattern
            normalize_special_chars: Auto-convert quotes for matching
            track: If True, show as tracked change; if False, silent replace (default: True)

        Raises:
            TextNotFoundError: If 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
            ValueError: If no footer of the specified type exists

        Example:
            >>> doc.replace_in_footer("Page {PAGE}", "Page {PAGE} of {NUMPAGES}")
            >>> doc.replace_in_footer("Draft", "Final", track=False)  # Silent replace
        """
        self._header_footer_ops.replace_in_footer(
            find=find,
            replace=replace,
            footer_type=footer_type,
            author=author,
            regex=regex,
            normalize_special_chars=normalize_special_chars,
            track=track,
        )

    def insert_in_header(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        header_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        normalize_special_chars: bool = True,
        track: bool = True,
    ) -> None:
        """Insert text in a header with optional tracked changes.

        Args:
            text: Text to insert
            after: Text to insert after
            before: Text to insert before
            header_type: Type of header ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat anchor as a regex pattern
            normalize_special_chars: Auto-convert quotes for matching
            track: If True, show as tracked change; if False, silent insert (default: True)

        Raises:
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
            ValueError: If no header of the specified type exists, or if both
                        'after' and 'before' are specified

        Example:
            >>> doc.insert_in_header(" - Final", after="Document Title")
            >>> doc.insert_in_header(" v2", after="Title", track=False)  # Silent insert
        """
        self._header_footer_ops.insert_in_header(
            text=text,
            after=after,
            before=before,
            header_type=header_type,
            author=author,
            regex=regex,
            normalize_special_chars=normalize_special_chars,
            track=track,
        )

    def insert_in_footer(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        footer_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        normalize_special_chars: bool = True,
        track: bool = True,
    ) -> None:
        """Insert text in a footer with optional tracked changes.

        Args:
            text: Text to insert
            after: Text to insert after
            before: Text to insert before
            footer_type: Type of footer ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat anchor as a regex pattern
            normalize_special_chars: Auto-convert quotes for matching
            track: If True, show as tracked change; if False, silent insert (default: True)

        Raises:
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
            ValueError: If no footer of the specified type exists, or if both
                        'after' and 'before' are specified

        Example:
            >>> doc.insert_in_footer(" - Confidential", after="Page")
            >>> doc.insert_in_footer(" v2", after="Page", track=False)  # Silent insert
        """
        self._header_footer_ops.insert_in_footer(
            text=text,
            after=after,
            before=before,
            footer_type=footer_type,
            author=author,
            regex=regex,
            normalize_special_chars=normalize_special_chars,
            track=track,
        )

    # ========================================================================
    # HYPERLINK METHODS
    # ========================================================================

    @property
    def hyperlinks(self) -> list[HyperlinkInfo]:
        """Get all hyperlinks in the document.

        Returns hyperlinks from all locations: body, headers, footers,
        footnotes, and endnotes.

        Returns:
            List of HyperlinkInfo objects with link details

        Example:
            >>> doc = Document("contract.docx")
            >>> for link in doc.hyperlinks:
            ...     print(f"{link.text} -> {link.target}")
            ...     if link.is_external:
            ...         print("  External URL")
            ...     else:
            ...         print("  Internal bookmark")
        """
        return self._hyperlink_ops.get_all_hyperlinks()

    def insert_hyperlink(
        self,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
        after: str | None = None,
        before: str | None = None,
        scope: str | dict | Any | None = None,
        tooltip: str | None = None,
        track: bool = False,
        author: str | None = None,
    ) -> str | None:
        """Insert a hyperlink at a specific location in the document body.

        Supports both external hyperlinks (URLs) and internal hyperlinks
        (bookmarks). Use the `url` parameter for external links and the
        `anchor` parameter for internal links.

        Args:
            url: External URL to link to (mutually exclusive with anchor)
            anchor: Internal bookmark name to link to (mutually exclusive with url)
            text: The display text for the hyperlink
            after: Text to insert after (mutually exclusive with before)
            before: Text to insert before (mutually exclusive with after)
            scope: Optional scope to limit search (paragraph ref, heading, etc.)
            tooltip: Optional tooltip text shown on hover
            track: If True, wrap insertion in tracked change markup
            author: Optional author override for tracked changes

        Returns:
            Relationship ID (rId) for external links, None for internal links

        Raises:
            ValueError: If both url and anchor specified, or neither specified
            ValueError: If both after and before specified, or neither specified
            TextNotFoundError: If anchor text not found
            AmbiguousTextError: If anchor text found multiple times

        Example:
            >>> # External hyperlink
            >>> doc.insert_hyperlink(
            ...     url="https://www.law.cornell.edu/uscode/text/28/1782",
            ...     text="28 U.S.C. section 1782",
            ...     after="discovery statute"
            ... )
            'rId15'

            >>> # Internal hyperlink to bookmark
            >>> doc.insert_hyperlink(
            ...     anchor="DefinitionsSection",
            ...     text="See Definitions",
            ...     after="as defined below"
            ... )
            None
        """
        return self._hyperlink_ops.insert_hyperlink(
            url=url,
            anchor=anchor,
            text=text,
            after=after,
            before=before,
            scope=scope,
            tooltip=tooltip,
            track=track,
            author=author,
        )

    def insert_hyperlink_in_header(
        self,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
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

        Raises:
            ValueError: If both url and anchor specified, or neither specified
            ValueError: If invalid header_type specified
            TextNotFoundError: If anchor text not found in header
        """
        return self._hyperlink_ops.insert_hyperlink_in_header(
            url=url,
            anchor=anchor,
            text=text,
            after=after,
            before=before,
            header_type=header_type,
            track=track,
        )

    def insert_hyperlink_in_footer(
        self,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
        after: str | None = None,
        before: str | None = None,
        footer_type: str = "default",
        track: bool = False,
    ) -> str | None:
        """Insert a hyperlink in a footer.

        Args:
            url: External URL (mutually exclusive with anchor)
            anchor: Internal bookmark name (mutually exclusive with url)
            text: Display text for the hyperlink
            after: Text to insert after
            before: Text to insert before
            footer_type: "default", "first", or "even"
            track: If True, track the insertion

        Returns:
            Relationship ID for external links, None for internal

        Raises:
            ValueError: If both url and anchor specified, or neither specified
            ValueError: If invalid footer_type specified
            TextNotFoundError: If anchor text not found in footer
        """
        return self._hyperlink_ops.insert_hyperlink_in_footer(
            url=url,
            anchor=anchor,
            text=text,
            after=after,
            before=before,
            footer_type=footer_type,
            track=track,
        )

    def insert_hyperlink_in_footnote(
        self,
        note_id: str | int,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
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

        Returns:
            Relationship ID for external links, None for internal

        Raises:
            NoteNotFoundError: If footnote not found
            ValueError: If both url and anchor specified, or neither specified
            TextNotFoundError: If anchor text not found in footnote
        """
        return self._hyperlink_ops.insert_hyperlink_in_footnote(
            note_id=note_id,
            url=url,
            anchor=anchor,
            text=text,
            after=after,
            before=before,
            track=track,
        )

    def insert_hyperlink_in_endnote(
        self,
        note_id: str | int,
        url: str | None = None,
        anchor: str | None = None,
        text: str = "",
        after: str | None = None,
        before: str | None = None,
        track: bool = False,
    ) -> str | None:
        """Insert a hyperlink inside an existing endnote.

        Args:
            note_id: The endnote ID to edit
            url: External URL (mutually exclusive with anchor)
            anchor: Internal bookmark name (mutually exclusive with url)
            text: Display text for the hyperlink
            after: Text to insert after within the endnote
            before: Text to insert before within the endnote
            track: If True, track the insertion

        Returns:
            Relationship ID for external links, None for internal

        Raises:
            NoteNotFoundError: If endnote not found
            ValueError: If both url and anchor specified, or neither specified
            TextNotFoundError: If anchor text not found in endnote
        """
        return self._hyperlink_ops.insert_hyperlink_in_endnote(
            note_id=note_id,
            url=url,
            anchor=anchor,
            text=text,
            after=after,
            before=before,
            track=track,
        )

    def edit_hyperlink_url(self, ref: str, new_url: str) -> None:
        """Change the URL of an external hyperlink.

        Updates the relationship target for the specified hyperlink.
        Only works for external hyperlinks (not internal bookmarks).

        Args:
            ref: Hyperlink ref (e.g., "lnk:5") or relationship ID (e.g., "rId5")
            new_url: The new URL to link to

        Raises:
            ValueError: If hyperlink not found or is an internal link
            ValueError: If new_url is empty

        Example:
            >>> doc.edit_hyperlink_url("lnk:5", "https://new-url.com")
            >>> doc.edit_hyperlink_url("rId10", "https://updated-url.com")
        """
        return self._hyperlink_ops.edit_hyperlink_url(ref=ref, new_url=new_url)

    def edit_hyperlink_text(
        self,
        ref: str,
        new_text: str,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Change the display text of a hyperlink.

        Replaces the visible text of the hyperlink while keeping the
        same target URL or bookmark.

        Args:
            ref: Hyperlink ref (e.g., "lnk:5")
            new_text: The new display text
            track: If True, show text change as tracked change
            author: Optional author for tracked change

        Raises:
            ValueError: If hyperlink not found
            ValueError: If new_text is empty

        Example:
            >>> doc.edit_hyperlink_text("lnk:5", "Updated link text")
            >>> doc.edit_hyperlink_text("lnk:5", "New Text", track=True)
        """
        return self._hyperlink_ops.edit_hyperlink_text(
            ref=ref,
            new_text=new_text,
            track=track,
            author=author,
        )

    def edit_hyperlink_anchor(self, ref: str, new_anchor: str) -> None:
        """Change the target bookmark of an internal hyperlink.

        Only works for internal hyperlinks (not external URLs).

        Args:
            ref: Hyperlink ref (e.g., "lnk:5")
            new_anchor: The new bookmark name to link to

        Raises:
            ValueError: If hyperlink not found or is an external link
            ValueError: If new_anchor is empty

        Example:
            >>> doc.edit_hyperlink_anchor("lnk:3", "NewBookmarkName")
        """
        return self._hyperlink_ops.edit_hyperlink_anchor(ref=ref, new_anchor=new_anchor)

    def remove_hyperlink(
        self,
        ref: str,
        keep_text: bool = True,
        track: bool = False,
        author: str | None = None,
    ) -> None:
        """Remove a hyperlink from the document.

        Can either keep the display text (unlinking it) or remove both
        the link and the text entirely.

        Args:
            ref: Hyperlink ref (e.g., "lnk:5")
            keep_text: If True (default), keep the display text without the link.
                       If False, remove both the link and the text.
            track: If True and keep_text=False, show text removal as tracked deletion
            author: Optional author for tracked change

        Raises:
            ValueError: If hyperlink not found

        Example:
            >>> # Keep text, just remove the link
            >>> doc.remove_hyperlink("lnk:5")

            >>> # Remove link and text entirely
            >>> doc.remove_hyperlink("lnk:5", keep_text=False)

            >>> # Remove with tracking
            >>> doc.remove_hyperlink("lnk:5", keep_text=False, track=True)
        """
        return self._hyperlink_ops.remove_hyperlink(
            ref=ref,
            keep_text=keep_text,
            track=track,
            author=author,
        )

    # ========================================================================
    # TABLE OF CONTENTS METHODS
    # ========================================================================

    def insert_toc(
        self,
        position: int | str = 0,
        levels: tuple[int, int] = (1, 3),
        title: str | None = "Table of Contents",
        hyperlinks: bool = True,
        show_page_numbers: bool = True,
        use_outline_levels: bool = True,
        update_on_open: bool = True,
    ) -> None:
        """Insert a Table of Contents field that Word will populate on open.

        This method inserts a TOC field into the document at the specified
        position. The TOC is wrapped in a Structured Document Tag (SDT) that
        Word recognizes as a TOC content control.

        The TOC field is marked as dirty, meaning Word will update it when the
        document is opened. Page numbers are calculated by Word's layout engine.

        Args:
            position: Where to insert the TOC. Can be:
                     - int: Paragraph index (0 = beginning of document)
                     - "start": Beginning of document body
                     - "end": End of document body (before sectPr)
            levels: Tuple of (min_level, max_level) for heading levels to include.
                   Default (1, 3) includes Heading 1, 2, and 3.
            title: Optional title text to display above the TOC.
                  Set to None for no title. Default is "Table of Contents".
            hyperlinks: If True (default), TOC entries link to their headings.
            show_page_numbers: If True (default), show page numbers.
            use_outline_levels: If True (default), include paragraphs with
                               outline levels.
            update_on_open: If True (default), sets w:updateFields in settings.xml
                           so Word updates all fields when opening the document.

        Example:
            >>> doc = Document("report.docx")
            >>> # Simple TOC with defaults
            >>> doc.insert_toc()
            >>> doc.save("report_with_toc.docx")
            >>>
            >>> # TOC without title, including more heading levels
            >>> doc.insert_toc(title=None, levels=(1, 5))
            >>>
            >>> # TOC at end of document without page numbers
            >>> doc.insert_toc(position="end", show_page_numbers=False)
        """
        return self._toc_ops.insert_toc(
            position=position,
            levels=levels,
            title=title,
            hyperlinks=hyperlinks,
            show_page_numbers=show_page_numbers,
            use_outline_levels=use_outline_levels,
            update_on_open=update_on_open,
        )

    def remove_toc(self) -> bool:
        """Remove the Table of Contents from the document.

        This method finds and removes the TOC SDT (Structured Document Tag)
        from the document, along with any title paragraph that precedes it.

        Returns:
            True if a TOC was found and removed, False if no TOC exists.

        Example:
            >>> doc = Document("report.docx")
            >>> if doc.remove_toc():
            ...     print("TOC removed successfully")
            ... else:
            ...     print("No TOC found in document")
            >>> doc.save("report_no_toc.docx")
        """
        return self._toc_ops.remove_toc()

    def mark_toc_dirty(self) -> bool:
        """Mark the TOC field as dirty so Word will recalculate it on open.

        This method sets w:dirty="true" on the TOC field's begin marker.
        When the document is opened in Word, this tells Word to recalculate
        the TOC to reflect current document structure.

        This is useful after modifying document content (adding/removing
        headings) to ensure the TOC reflects the current state.

        Returns:
            True if a TOC was found and marked dirty, False if no TOC exists.

        Example:
            >>> doc = Document("report.docx")
            >>> # After making changes to document headings
            >>> if doc.mark_toc_dirty():
            ...     print("TOC marked for update")
            >>> doc.save("report_updated.docx")
        """
        return self._toc_ops.mark_toc_dirty()

    def get_toc(self) -> TOC | None:
        """Get information about an existing Table of Contents in the document.

        This method finds and parses an existing TOC, extracting:
        - Position in the document
        - Field switches (levels, hyperlinks, etc.)
        - Whether it's marked dirty
        - Cached entries (text, level, page number, bookmark)

        Note that the entries are cached values from when Word last updated
        the TOC. They may be stale if the document has been modified.

        Returns:
            A TOC object containing the parsed information, or None if no TOC
            is found in the document.

        Example:
            >>> doc = Document("report.docx")
            >>> toc = doc.get_toc()
            >>> if toc:
            ...     print(f"TOC at paragraph {toc.position}")
            ...     print(f"Levels: {toc.levels}")
            ...     print(f"Dirty: {toc.is_dirty}")
            ...     for entry in toc.entries:
            ...         print(f"  L{entry.level}: {entry.text} ... {entry.page_number}")
        """
        return self._toc_ops.get_toc()

    def update_toc(
        self,
        levels: tuple[int, int] | None = None,
        hyperlinks: bool | None = None,
        show_page_numbers: bool | None = None,
        use_outline_levels: bool | None = None,
        **kwargs: Any,
    ) -> bool:
        """Update an existing TOC's field instruction and/or title.

        This method modifies an existing TOC in place without removing it.
        Only the parameters that are explicitly provided (not None) will be
        updated; other settings are preserved from the existing TOC.

        For the title parameter, use the following convention:
        - title=None: Remove the title
        - title="New Title": Change the title to this value
        - title not provided: Don't change the title

        Args:
            levels: New heading levels tuple (min_level, max_level). If provided,
                   updates the \\o switch. Example: (1, 5) for headings 1-5.
            hyperlinks: If True, add \\h switch (entries link to headings).
                       If False, remove \\h switch. If None, preserve current.
            show_page_numbers: If True, remove \\n switch (show page numbers).
                              If False, add \\n switch (hide page numbers).
                              If None, preserve current.
            use_outline_levels: If True, add \\u switch (include outline levels).
                               If False, remove \\u switch. If None, preserve current.
            title: New title text. If explicitly set to None, removes the title
                  paragraph. If not provided, the title is unchanged.

        Returns:
            True if a TOC was found and updated, False if no TOC exists.

        Example:
            >>> doc = Document("report.docx")
            >>> # Update levels to include more headings
            >>> doc.update_toc(levels=(1, 5))
            >>> # Turn off hyperlinks
            >>> doc.update_toc(hyperlinks=False)
            >>> # Change the title
            >>> doc.update_toc(title="Table of Contents")
            >>> # Remove the title
            >>> doc.update_toc(title=None)
            >>> doc.save("report_updated.docx")
        """
        return self._toc_ops.update_toc(
            levels=levels,
            hyperlinks=hyperlinks,
            show_page_numbers=show_page_numbers,
            use_outline_levels=use_outline_levels,
            **kwargs,
        )

    # ========================================================================
    # CROSS-REFERENCE METHODS
    # ========================================================================

    def insert_cross_reference(
        self,
        target: str,
        display: str = "text",
        after: str | None = None,
        before: str | None = None,
        scope: str | dict | Any | None = None,
        hyperlink: bool = True,
        track: bool = False,
        author: str | None = None,
    ) -> str:
        """Insert a cross-reference to a bookmark, heading, figure, table, or note.

        Creates a field code (REF, PAGEREF, or NOTEREF) that references the
        target and displays its content. The field is marked dirty so Word
        will calculate the display value when the document opens.

        Args:
            target: What to reference. Formats:
                   - "bookmark_name" - Direct bookmark reference
                   - "heading:Introduction" - Reference heading by text
                   - "figure:1" or "figure:Architecture" - Figure by number or text
                   - "table:2" or "table:Revenue" - Table by number or text
                   - "footnote:1" - Footnote by number
                   - "endnote:2" - Endnote by number
            display: What to display. Options:
                    - "text" - Target content (default)
                    - "page" - Page number
                    - "above_below" - Position relative to reference
                    - "number" - Heading/caption number only
                    - "full_number" - Full number including chapter
                    - "relative_number" - Number relative to context
                    - "label_number" - "Figure 1" or "Table 2"
                    - "number_only" - Just the number
            after: Insert after this text (mutually exclusive with before)
            before: Insert before this text (mutually exclusive with after)
            scope: Limit text search scope
            hyperlink: If True (default), make the reference clickable
            track: If True, wrap in tracked change markup
            author: Author for tracked changes

        Returns:
            The bookmark name used for the cross-reference

        Example:
            >>> doc.insert_cross_reference("heading:Introduction", after="See ")
            >>> doc.insert_cross_reference("figure:1", display="page", after="on page ")
            >>> doc.insert_cross_reference("table:Revenue", display="label_number")
        """
        return self._cross_reference_ops.insert_cross_reference(
            target=target,
            display=display,
            after=after,
            before=before,
            scope=scope,
            hyperlink=hyperlink,
            track=track,
            author=author,
        )

    def insert_page_reference(
        self,
        target: str,
        after: str | None = None,
        before: str | None = None,
        scope: str | dict | Any | None = None,
        show_position: bool = False,
        hyperlink: bool = True,
        track: bool = False,
        author: str | None = None,
    ) -> str:
        """Insert a page number cross-reference to a target.

        Convenience method for inserting a PAGEREF field that displays
        the page number where the target appears.

        Args:
            target: What to reference (bookmark, heading:..., figure:..., etc.)
            after: Insert after this text
            before: Insert before this text
            scope: Limit text search scope
            show_position: If True, show "above" or "below" for same-page refs
            hyperlink: If True (default), make the reference clickable
            track: If True, wrap in tracked change markup
            author: Author for tracked changes

        Returns:
            The bookmark name used for the cross-reference

        Example:
            >>> doc.insert_page_reference("heading:Conclusion", after="see page ")
        """
        return self._cross_reference_ops.insert_page_reference(
            target=target,
            after=after,
            before=before,
            scope=scope,
            show_position=show_position,
            hyperlink=hyperlink,
            track=track,
            author=author,
        )

    def insert_note_reference(
        self,
        note_type: str,
        note_id: int | str,
        after: str | None = None,
        before: str | None = None,
        scope: str | dict | Any | None = None,
        show_position: bool = False,
        use_note_style: bool = True,
        hyperlink: bool = True,
        track: bool = False,
        author: str | None = None,
    ) -> str:
        """Insert a cross-reference to a footnote or endnote.

        Creates a NOTEREF field that displays the note number.

        Args:
            note_type: "footnote" or "endnote"
            note_id: The note number to reference
            after: Insert after this text
            before: Insert before this text
            scope: Limit text search scope
            show_position: If True, show "above" or "below" for same-page refs
            use_note_style: If True (default), format like the note reference mark
            hyperlink: If True (default), make the reference clickable
            track: If True, wrap in tracked change markup
            author: Author for tracked changes

        Returns:
            The bookmark name used for the cross-reference

        Example:
            >>> doc.insert_note_reference("footnote", 1, after="see note ")
        """
        return self._cross_reference_ops.insert_note_reference(
            note_type=note_type,
            note_id=note_id,
            after=after,
            before=before,
            scope=scope,
            show_position=show_position,
            use_note_style=use_note_style,
            hyperlink=hyperlink,
            track=track,
            author=author,
        )

    def create_bookmark(
        self,
        name: str,
        at: str,
        scope: str | dict | Any | None = None,
    ) -> str:
        """Create a named bookmark at the specified text location.

        Bookmarks are named locations in a document that can be referenced
        by cross-references, hyperlinks, or other features.

        Args:
            name: The bookmark name. Must be alphanumeric with underscores,
                 start with a letter, and be 40 characters or fewer.
            at: The text to bookmark
            scope: Limit text search scope

        Returns:
            The bookmark name (same as the name parameter)

        Raises:
            InvalidBookmarkNameError: If the name is invalid
            BookmarkAlreadyExistsError: If a bookmark with this name exists
            TextNotFoundError: If the text is not found
            AmbiguousTextError: If the text appears multiple times

        Example:
            >>> doc.create_bookmark("introduction", at="1. Introduction")
            >>> doc.insert_cross_reference("introduction", after="See ")
        """
        return self._cross_reference_ops.create_bookmark(
            name=name,
            at=at,
            scope=scope,
        )

    def create_heading_bookmark(
        self,
        heading_text: str,
        bookmark_name: str | None = None,
    ) -> str:
        """Create a bookmark at a heading for cross-referencing.

        Finds a heading by text and creates a bookmark that wraps it.
        If no bookmark name is provided, generates a hidden _Ref bookmark.

        Args:
            heading_text: Text to search for in headings (partial match)
            bookmark_name: Optional custom bookmark name. If None, generates
                          an auto-incremented _Ref bookmark name.

        Returns:
            The bookmark name (provided or generated)

        Raises:
            CrossReferenceTargetNotFoundError: If no matching heading found

        Example:
            >>> doc.create_heading_bookmark("Introduction")
            '_Ref12345'
            >>> doc.create_heading_bookmark("Conclusion", "my_conclusion")
            'my_conclusion'
        """
        return self._cross_reference_ops.create_heading_bookmark(
            heading_text=heading_text,
            bookmark_name=bookmark_name,
        )

    def list_bookmarks(self, include_hidden: bool = False) -> list[BookmarkInfo]:
        """List all bookmarks in the document.

        Args:
            include_hidden: If True, include hidden _Ref bookmarks

        Returns:
            List of BookmarkInfo objects with bookmark details

        Example:
            >>> for bm in doc.list_bookmarks():
            ...     print(f"{bm.name}: {bm.text_preview}")
        """
        return self._cross_reference_ops.list_bookmarks(include_hidden=include_hidden)

    def get_bookmark(self, name: str) -> BookmarkInfo | None:
        """Get information about a specific bookmark.

        Args:
            name: The bookmark name

        Returns:
            BookmarkInfo object, or None if not found

        Example:
            >>> bm = doc.get_bookmark("introduction")
            >>> if bm:
            ...     print(f"Found at: {bm.location}")
        """
        return self._cross_reference_ops.get_bookmark(name)

    def get_cross_references(self) -> list[CrossReference]:
        """List all cross-references in the document.

        Returns a list of CrossReference objects containing:
        - Field type (REF, PAGEREF, NOTEREF)
        - Target bookmark
        - Display options/switches
        - Current cached display value
        - Dirty status

        Returns:
            List of CrossReference objects

        Example:
            >>> for xref in doc.get_cross_references():
            ...     print(f"{xref.field_type} -> {xref.target_bookmark}")
        """
        return self._cross_reference_ops.get_cross_references()

    def get_cross_reference_targets(self) -> list[CrossReferenceTarget]:
        """List all available cross-reference targets in the document.

        Returns all bookmarks, headings, figures, tables, footnotes, and
        endnotes that can be referenced.

        Returns:
            List of CrossReferenceTarget objects

        Example:
            >>> for target in doc.get_cross_reference_targets():
            ...     print(f"{target.type}: {target.display_name}")
        """
        return self._cross_reference_ops.get_cross_reference_targets()

    def mark_cross_references_dirty(self) -> int:
        """Mark all cross-reference fields for update when opened in Word.

        Sets w:dirty="true" on all REF, PAGEREF, and NOTEREF fields,
        telling Word to recalculate their display values.

        Returns:
            The number of fields marked dirty

        Example:
            >>> count = doc.mark_cross_references_dirty()
            >>> print(f"Marked {count} cross-references for update")
        """
        return self._cross_reference_ops.mark_cross_references_dirty()

    # ========================================================================
    # Ref-based editing operations (DocTree accessibility layer)
    # ========================================================================

    @property
    def _ref_registry(self) -> Any:
        """Get the RefRegistry instance (lazy initialization)."""
        if not hasattr(self, "_ref_registry_instance"):
            from python_docx_redline.accessibility.registry import RefRegistry

            self._ref_registry_instance = RefRegistry(self.xml_root)
        return self._ref_registry_instance

    def resolve_ref(self, ref: str) -> etree._Element:
        """Resolve a ref string to its corresponding XML element.

        This method resolves refs like "p:5" or "tbl:0/row:2/cell:1" to their
        underlying lxml elements. Refs provide stable, unambiguous identifiers
        for document elements.

        Args:
            ref: A ref path string (e.g., "p:5", "tbl:0/row:2/cell:1/p:0",
                "p:~xK4mNp2q" for fingerprint-based refs)

        Returns:
            The lxml element corresponding to the ref

        Raises:
            RefNotFoundError: If the ref cannot be resolved (invalid format,
                out of bounds, or element not found)
            StaleRefError: If the ref points to a deleted/modified element

        Example:
            >>> doc = Document("contract.docx")
            >>> element = doc.resolve_ref("p:5")  # Get 6th paragraph
            >>> element = doc.resolve_ref("tbl:0/row:1/cell:2")  # Table cell
        """
        return self._ref_registry.resolve_ref(ref)

    def get_ref(self, element: etree._Element, use_fingerprint: bool = False) -> "Ref":
        """Get the ref for a given XML element.

        This method generates a ref path for an element, which can then be used
        for subsequent editing operations. Supports both ordinal refs (p:5) and
        fingerprint-based refs (p:~xK4mNp2q) for stable references.

        Args:
            element: The lxml element to get a ref for
            use_fingerprint: If True, generate a fingerprint-based ref that
                survives document edits. Fingerprint refs are based on content
                hash and remain stable even if paragraphs are inserted/deleted.

        Returns:
            Ref object containing the path (e.g., "p:5")

        Raises:
            RefNotFoundError: If the element type is not supported

        Example:
            >>> doc = Document("contract.docx")
            >>> matches = doc.find_all("Total Price")
            >>> if matches:
            ...     # Get ref for the paragraph containing the match
            ...     para_element = matches[0].span.paragraph
            ...     ref = doc.get_ref(para_element)
            ...     print(ref.path)  # e.g., "p:15"
            >>>
            >>> # Get a stable fingerprint ref
            >>> element = doc.resolve_ref("p:3")
            >>> stable_ref = doc.get_ref(element, use_fingerprint=True)
            >>> print(stable_ref.path)  # e.g., "p:~xK4mNp2q"
        """
        return self._ref_registry.get_ref(element, use_fingerprint=use_fingerprint)

    def get_text_at_ref(self, ref: str) -> str:
        """Get the full text content of the element at ref.

        This method extracts all visible text from the element identified by the ref.
        It is useful for getting the complete text of a paragraph or table cell,
        which can then be used with other editing methods like delete_tracked().

        Args:
            ref: Element reference (e.g., "p:5", "tbl:0/row:1/cell:2")

        Returns:
            Full text content of the element (concatenated from all w:t elements)

        Raises:
            RefNotFoundError: If the ref cannot be resolved
            StaleRefError: If the ref points to a deleted/modified element

        Example:
            >>> doc = Document("contract.docx")
            >>> # Get text from a paragraph discovered via AccessibilityTree
            >>> text = doc.get_text_at_ref("p:15")
            >>> doc.delete_tracked(text)  # Delete entire paragraph
            >>>
            >>> # Get text from a table cell
            >>> cell_text = doc.get_text_at_ref("tbl:0/row:1/cell:2")
        """
        element = self.resolve_ref(ref)
        text_parts = []
        for t_elem in element.iter(f"{{{WORD_NAMESPACE}}}t"):
            if t_elem.text:
                text_parts.append(t_elem.text)
        return "".join(text_parts)

    def insert_at_ref(
        self,
        ref: str,
        text: str,
        position: str = "after",
        track: bool = False,
        author: str | None = None,
    ) -> EditResult:
        """Insert text at a ref location.

        This method inserts text at a position relative to the element identified
        by the ref. For paragraphs, "before"/"after" insert new paragraphs, while
        "start"/"end" insert runs within the paragraph.

        Args:
            ref: Ref path to the target element (e.g., "p:5", "tbl:0/row:1/cell:0")
            text: The text to insert (supports markdown: **bold**, *italic*, etc.)
            position: Where to insert relative to the element:
                - "before": Insert new paragraph before (for paragraph refs)
                - "after": Insert new paragraph after (for paragraph refs)
                - "start": Insert at the start of the element's content
                - "end": Insert at the end of the element's content
            track: Whether to add tracked insertion markup (default: False)
            author: Author for tracked changes (uses document author if None)

        Returns:
            EditResult indicating success/failure

        Raises:
            RefNotFoundError: If the ref cannot be resolved
            ValueError: If the position is invalid

        Example:
            >>> doc = Document("contract.docx")
            >>> # Insert tracked text at end of paragraph
            >>> doc.insert_at_ref("p:5", " (AMENDED)", position="end", track=True)
            >>>
            >>> # Insert new paragraph after ref
            >>> doc.insert_at_ref("p:10", "New clause text.", position="after")
        """

        from .accessibility.types import ElementType

        valid_positions = ("before", "after", "start", "end")
        if position not in valid_positions:
            raise ValueError(f"position must be one of {valid_positions}, got '{position}'")

        # Resolve the ref to get the element
        element = self.resolve_ref(ref)
        ref_obj = self._ref_registry.get_ref(element)
        element_type = ref_obj.element_type

        # Get author
        author_name = author if author is not None else self.author

        # Handle based on element type and position
        if element_type == ElementType.PARAGRAPH:
            if position in ("before", "after"):
                # Insert new paragraph
                return self._insert_paragraph_at_ref(element, text, position, track, author_name)
            else:
                # Insert run at start/end of paragraph
                return self._insert_run_at_ref(element, text, position, track, author_name)
        elif element_type == ElementType.TABLE_CELL:
            # For table cells, insert at start/end of first/last paragraph
            paragraphs = list(element.findall(f".//{{{WORD_NAMESPACE}}}p"))
            if not paragraphs:
                # Create a paragraph if none exists
                para = etree.SubElement(element, f"{{{WORD_NAMESPACE}}}p")
                return self._insert_run_at_ref(para, text, "start", track, author_name)
            target_para = paragraphs[0] if position in ("before", "start") else paragraphs[-1]
            return self._insert_run_at_ref(
                target_para,
                text,
                "start" if position in ("before", "start") else "end",
                track,
                author_name,
            )
        else:
            # Default: try to insert as text content
            return self._insert_run_at_ref(element, text, position, track, author_name)

    def insert_in_ref(
        self,
        ref: str,
        text: str,
        before: str | None = None,
        after: str | None = None,
        track: bool = False,
        author: str | None = None,
    ) -> EditResult:
        """Insert text within a specific element identified by ref.

        Unlike insert_at_ref() which inserts an element at a position relative
        to the ref element, this method inserts text within an existing element
        by finding anchor text inside the element first.

        Args:
            ref: Element reference (e.g., "p:15", "tbl:0/row:1/cell:0")
            text: Text to insert (supports markdown: **bold**, *italic*, etc.)
            before: Insert before this anchor text (mutually exclusive with after)
            after: Insert after this anchor text (mutually exclusive with before)
            track: If True, show as tracked change
            author: Author for tracked change (uses document author if None)

        Returns:
            EditResult indicating success/failure

        Raises:
            RefNotFoundError: If the ref cannot be resolved
            ValueError: If neither before nor after is specified, or if both are
            TextNotFoundError: If the anchor text is not found within the element
            AmbiguousTextError: If the anchor text appears multiple times in the element

        Example:
            >>> doc = Document("contract.docx")
            >>> # Append "(amended)" after "Section 2.1" within paragraph 15
            >>> doc.insert_in_ref("p:15", " (amended)", after="Section 2.1", track=True)
            >>>
            >>> # Insert "Important: " before "Terms" within a table cell
            >>> doc.insert_in_ref("tbl:0/row:1/cell:0", "Important: ", before="Terms")
        """
        from .errors import AmbiguousTextError, TextNotFoundError

        # Validate parameters
        if before is not None and after is not None:
            raise ValueError("Cannot specify both 'before' and 'after' parameters")
        if before is None and after is None:
            raise ValueError("Must specify either 'before' or 'after' parameter")

        # Resolve the ref to get the element
        element = self.resolve_ref(ref)

        # Get author
        author_name = author if author is not None else self.author

        # Determine anchor text and insertion mode
        anchor: str = after if after is not None else before  # type: ignore[assignment]
        insert_after = after is not None

        # Get all paragraphs within the element (or the element itself if it's a paragraph)
        if element.tag == f"{{{WORD_NAMESPACE}}}p":
            paragraphs = [element]
        else:
            # For tables, cells, etc., find all paragraphs within
            paragraphs = list(element.findall(f".//{{{WORD_NAMESPACE}}}p"))

        if not paragraphs:
            return EditResult(
                success=False,
                edit_type="insert_in_ref",
                message=f"No paragraphs found within element at ref '{ref}'",
            )

        # Search for the anchor text within the element's paragraphs
        matches = self._text_search.find_text(
            anchor,
            paragraphs,
            regex=False,
            normalize_special_chars=True,
            fuzzy=None,
        )

        if not matches:
            raise TextNotFoundError(
                anchor,
                scope=f"element at ref '{ref}'",
                hint=f"The anchor text was not found within the element at '{ref}'",
            )

        if len(matches) > 1:
            raise AmbiguousTextError(anchor, matches)

        # Get the single match
        match = matches[0]

        # Create the insertion element
        if track:
            # Tracked insertion: wrap in <w:ins>
            insertion_xml = self._xml_generator.create_insertion(text, author_name)
            elements = self._tracked_ops._parse_xml_elements(insertion_xml)
            insertion_element = elements[0]
        else:
            # Untracked insertion: plain runs
            source_run = match.runs[0] if match.runs else None
            plain_runs = self._xml_generator.create_plain_runs(text, source_run=source_run)
            insertion_element = plain_runs

        # Insert at the appropriate position
        if insert_after:
            self._tracked_ops._insert_after_match(match, insertion_element)
        else:
            self._tracked_ops._insert_before_match(match, insertion_element)

        return EditResult(
            success=True,
            edit_type="insert_in_ref",
            message=f"Inserted text {'after' if insert_after else 'before'} anchor in ref '{ref}'",
        )

    def _insert_paragraph_at_ref(
        self,
        element: etree._Element,
        text: str,
        position: str,
        track: bool,
        author: str,
    ) -> EditResult:
        """Insert a new paragraph before or after an element."""
        from datetime import datetime, timezone

        parent = element.getparent()
        if parent is None:
            return EditResult(
                success=False,
                edit_type="insert_at_ref",
                message="Cannot insert: element has no parent",
            )

        # Find the index of the element in its parent
        try:
            index = list(parent).index(element)
        except ValueError:
            return EditResult(
                success=False,
                edit_type="insert_at_ref",
                message="Cannot find element in parent",
            )

        # Create new paragraph
        new_para = etree.Element(f"{{{WORD_NAMESPACE}}}p")

        if track:
            # Create tracked insertion
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._xml_generator.next_change_id

            ins_elem = etree.SubElement(new_para, f"{{{WORD_NAMESPACE}}}ins")
            ins_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
            ins_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
            ins_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

            run = etree.SubElement(ins_elem, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = text
            # Add xml:space="preserve" for whitespace preservation
            if text and (text[0].isspace() or text[-1].isspace()):
                t.set(f"{{{XML_NAMESPACE}}}space", "preserve")

            self._xml_generator.next_change_id = change_id + 1
        else:
            # Untracked insertion
            run = etree.SubElement(new_para, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = text
            # Add xml:space="preserve" for whitespace preservation
            if text and (text[0].isspace() or text[-1].isspace()):
                t.set(f"{{{XML_NAMESPACE}}}space", "preserve")

        # Insert at the correct position
        if position == "after":
            parent.insert(index + 1, new_para)
        else:  # before
            parent.insert(index, new_para)

        # Invalidate registry cache
        self._ref_registry.invalidate()

        return EditResult(
            success=True,
            edit_type="insert_at_ref",
            message=f"Inserted paragraph {position} ref",
        )

    def _insert_run_at_ref(
        self,
        element: etree._Element,
        text: str,
        position: str,
        track: bool,
        author: str,
    ) -> EditResult:
        """Insert a run at the start or end of an element."""
        from datetime import datetime, timezone

        if track:
            # Create tracked insertion
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._xml_generator.next_change_id

            ins_elem = etree.Element(f"{{{WORD_NAMESPACE}}}ins")
            ins_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
            ins_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
            ins_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

            run = etree.SubElement(ins_elem, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = text
            # Add xml:space="preserve" for whitespace preservation
            if text and (text[0].isspace() or text[-1].isspace()):
                t.set(f"{{{XML_NAMESPACE}}}space", "preserve")

            if position == "start":
                element.insert(0, ins_elem)
            else:  # end
                element.append(ins_elem)

            self._xml_generator.next_change_id = change_id + 1
        else:
            # Untracked insertion
            run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = text
            # Add xml:space="preserve" for whitespace preservation
            if text and (text[0].isspace() or text[-1].isspace()):
                t.set(f"{{{XML_NAMESPACE}}}space", "preserve")

            if position == "start":
                element.insert(0, run)
            else:  # end
                element.append(run)

        return EditResult(
            success=True,
            edit_type="insert_at_ref",
            message=f"Inserted text at {position} of element",
        )

    def delete_ref(
        self,
        ref: str,
        track: bool = False,
        author: str | None = None,
    ) -> EditResult:
        """Delete the element at ref.

        This method deletes the element identified by the ref. For paragraphs,
        this deletes the entire paragraph. For table cells, this clears the cell
        content.

        Args:
            ref: Ref path to the element to delete (e.g., "p:5")
            track: If True, use tracked deletion instead of hard delete
            author: Author for tracked changes (uses document author if None)

        Returns:
            EditResult indicating success/failure

        Raises:
            RefNotFoundError: If the ref cannot be resolved

        Example:
            >>> doc = Document("contract.docx")
            >>> # Delete paragraph with tracking
            >>> doc.delete_ref("p:5", track=True)
            >>>
            >>> # Hard delete without tracking
            >>> doc.delete_ref("p:10", track=False)
        """

        from .accessibility.types import ElementType

        # Resolve the ref
        element = self.resolve_ref(ref)
        ref_obj = self._ref_registry.get_ref(element)
        element_type = ref_obj.element_type

        # Get author
        author_name = author if author is not None else self.author

        if element_type == ElementType.PARAGRAPH:
            return self._delete_paragraph_ref(element, track, author_name)
        elif element_type == ElementType.TABLE_CELL:
            # For table cells, delete content but not the cell itself
            return self._delete_cell_content_ref(element, track, author_name)
        else:
            # Generic element deletion
            return self._delete_element_ref(element, track, author_name)

    def _mark_paragraph_mark_deleted(
        self, para_element: etree._Element, author: str, timestamp: str
    ) -> None:
        """Mark the paragraph mark as deleted for tracked changes.

        Adds a <w:del> element inside <w:pPr>/<w:rPr> to mark the paragraph
        mark (the invisible character at the end of each paragraph) as deleted.
        When this tracked change is accepted in Word, the paragraph merges with
        the following paragraph instead of leaving an empty line behind.

        Per OOXML spec (ISO/IEC 29500): "This element specifies that the
        paragraph mark delimiting the end of a paragraph shall be treated
        as deleted... the contents of this paragraph are combined with the
        following paragraph."

        Args:
            para_element: The paragraph element to mark
            author: Author name for the tracked change
            timestamp: ISO 8601 timestamp for the change
        """
        # Get or create paragraph properties <w:pPr>
        p_pr = para_element.find(f"{{{WORD_NAMESPACE}}}pPr")
        if p_pr is None:
            p_pr = etree.Element(f"{{{WORD_NAMESPACE}}}pPr")
            para_element.insert(0, p_pr)

        # Get or create run properties for paragraph mark <w:rPr>
        r_pr = p_pr.find(f"{{{WORD_NAMESPACE}}}rPr")
        if r_pr is None:
            r_pr = etree.Element(f"{{{WORD_NAMESPACE}}}rPr")
            p_pr.append(r_pr)

        # Create the deletion marker for the paragraph mark
        change_id = self._xml_generator.next_change_id
        self._xml_generator.next_change_id += 1

        del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
        del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
        del_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
        del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

        r_pr.append(del_elem)

    def _is_paragraph_already_deleted(self, para_element: etree._Element) -> bool:
        """Check if a paragraph is already marked as deleted.

        A paragraph is considered already deleted if:
        1. It has a deleted paragraph mark (<w:pPr>/<w:rPr>/<w:del>), AND
        2. It has no direct <w:r> children (all content is wrapped in <w:del>)

        This prevents double-deletion which creates invalid nested <w:del> elements.

        Args:
            para_element: The paragraph XML element to check

        Returns:
            True if the paragraph is already deleted, False otherwise
        """
        # Check for deleted paragraph mark
        has_deleted_mark = self._has_deleted_paragraph_mark(para_element)

        # Check for direct run children (not wrapped in w:del)
        direct_runs = para_element.findall(f"./{{{WORD_NAMESPACE}}}r")
        has_direct_runs = len(direct_runs) > 0

        # Paragraph is already deleted if it has the deleted mark and no direct runs
        # (meaning all content is already wrapped in <w:del>)
        return has_deleted_mark and not has_direct_runs

    def _delete_paragraph_ref(
        self,
        element: etree._Element,
        track: bool,
        author: str,
    ) -> EditResult:
        """Delete a paragraph element."""
        from datetime import datetime, timezone

        if track:
            # Check if already deleted to prevent nested <w:del> elements
            if self._is_paragraph_already_deleted(element):
                return EditResult(
                    success=True,
                    edit_type="delete_ref",
                    message="Paragraph already deleted, skipping to prevent invalid XML",
                )

            # Tracked deletion: wrap all runs in w:del
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._xml_generator.next_change_id

            for run in list(element.findall(f"./{{{WORD_NAMESPACE}}}r")):
                # Create deletion wrapper
                del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
                del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                del_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
                del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

                # Convert w:t to w:delText
                for t_elem in run.findall(f".//{{{WORD_NAMESPACE}}}t"):
                    t_elem.tag = f"{{{WORD_NAMESPACE}}}delText"

                # Move run into deletion
                run_parent = run.getparent()
                if run_parent is not None:
                    run_index = list(run_parent).index(run)
                    run_parent.remove(run)
                    del_elem.append(run)
                    run_parent.insert(run_index, del_elem)

                change_id += 1

            self._xml_generator.next_change_id = change_id

            # Mark the paragraph mark as deleted (causes merge with next
            # paragraph on accept instead of leaving empty line)
            self._mark_paragraph_mark_deleted(element, author, timestamp)
        else:
            # Hard delete: remove the paragraph element
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)

        # Invalidate registry cache
        self._ref_registry.invalidate()

        return EditResult(
            success=True,
            edit_type="delete_ref",
            message="Deleted paragraph" + (" with tracking" if track else ""),
        )

    def _delete_cell_content_ref(
        self,
        element: etree._Element,
        track: bool,
        author: str,
    ) -> EditResult:
        """Delete content inside a table cell."""
        from datetime import datetime, timezone

        if track:
            # Tracked deletion for all paragraphs in cell
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._xml_generator.next_change_id

            for para in element.findall(f".//{{{WORD_NAMESPACE}}}p"):
                for run in list(para.findall(f"./{{{WORD_NAMESPACE}}}r")):
                    del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
                    del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                    del_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
                    del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

                    for t_elem in run.findall(f".//{{{WORD_NAMESPACE}}}t"):
                        t_elem.tag = f"{{{WORD_NAMESPACE}}}delText"

                    run_index = list(para).index(run)
                    para.remove(run)
                    del_elem.append(run)
                    para.insert(run_index, del_elem)

                    change_id += 1

            self._xml_generator.next_change_id = change_id
        else:
            # Hard delete: remove content from all paragraphs
            for para in element.findall(f".//{{{WORD_NAMESPACE}}}p"):
                for run in list(para.findall(f"./{{{WORD_NAMESPACE}}}r")):
                    para.remove(run)

        return EditResult(
            success=True,
            edit_type="delete_ref",
            message="Deleted cell content" + (" with tracking" if track else ""),
        )

    def _delete_element_ref(
        self,
        element: etree._Element,
        track: bool,
        author: str,
    ) -> EditResult:
        """Delete a generic element."""
        if track:
            # For non-paragraph elements, wrap content in deletion
            # This is a simplified implementation
            return EditResult(
                success=False,
                edit_type="delete_ref",
                message="Tracked deletion not supported for this element type",
            )
        else:
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)

        self._ref_registry.invalidate()

        return EditResult(
            success=True,
            edit_type="delete_ref",
            message="Deleted element",
        )

    def replace_at_ref(
        self,
        ref: str,
        new_text: str,
        track: bool = False,
        author: str | None = None,
    ) -> EditResult:
        """Replace content at ref with new text.

        This method replaces the content of the element identified by the ref
        with new text. When track=True, this shows as a deletion of old content
        plus insertion of new content.

        Args:
            ref: Ref path to the element (e.g., "p:5")
            new_text: The replacement text (supports markdown: **bold**, etc.)
            track: If True, mark as tracked deletion + insertion
            author: Author for tracked changes (uses document author if None)

        Returns:
            EditResult indicating success/failure

        Raises:
            RefNotFoundError: If the ref cannot be resolved

        Example:
            >>> doc = Document("contract.docx")
            >>> # Replace paragraph content with tracking
            >>> doc.replace_at_ref("p:5", "New clause text", track=True)
        """

        from .accessibility.types import ElementType

        # Resolve the ref
        element = self.resolve_ref(ref)
        ref_obj = self._ref_registry.get_ref(element)
        element_type = ref_obj.element_type

        # Get author
        author_name = author if author is not None else self.author

        if element_type == ElementType.PARAGRAPH:
            return self._replace_paragraph_content(element, new_text, track, author_name)
        elif element_type == ElementType.TABLE_CELL:
            return self._replace_cell_content(element, new_text, track, author_name)
        else:
            return EditResult(
                success=False,
                edit_type="replace_at_ref",
                message=f"Replace not supported for element type {element_type}",
            )

    def _replace_paragraph_content(
        self,
        element: etree._Element,
        new_text: str,
        track: bool,
        author: str,
    ) -> EditResult:
        """Replace content of a paragraph."""
        from datetime import datetime, timezone

        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        if track:
            change_id = self._xml_generator.next_change_id

            # First, mark all existing runs as deleted
            for run in list(element.findall(f"./{{{WORD_NAMESPACE}}}r")):
                del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
                del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                del_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
                del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

                for t_elem in run.findall(f".//{{{WORD_NAMESPACE}}}t"):
                    t_elem.tag = f"{{{WORD_NAMESPACE}}}delText"

                run_index = list(element).index(run)
                element.remove(run)
                del_elem.append(run)
                element.insert(run_index, del_elem)

                change_id += 1

            # Then, add the new content as an insertion
            ins_elem = etree.SubElement(element, f"{{{WORD_NAMESPACE}}}ins")
            ins_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
            ins_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
            ins_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

            run = etree.SubElement(ins_elem, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = new_text
            # Add xml:space="preserve" for whitespace preservation
            if new_text and (new_text[0].isspace() or new_text[-1].isspace()):
                t.set(f"{{{XML_NAMESPACE}}}space", "preserve")

            self._xml_generator.next_change_id = change_id + 1
        else:
            # Untracked: remove all runs and add new content
            for run in list(element.findall(f"./{{{WORD_NAMESPACE}}}r")):
                element.remove(run)

            # Also remove any tracked changes
            for del_elem in list(element.findall(f"./{{{WORD_NAMESPACE}}}del")):
                element.remove(del_elem)
            for ins_elem in list(element.findall(f"./{{{WORD_NAMESPACE}}}ins")):
                element.remove(ins_elem)

            # Add new run
            run = etree.SubElement(element, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = new_text
            # Add xml:space="preserve" for whitespace preservation
            if new_text and (new_text[0].isspace() or new_text[-1].isspace()):
                t.set(f"{{{XML_NAMESPACE}}}space", "preserve")

        return EditResult(
            success=True,
            edit_type="replace_at_ref",
            message="Replaced paragraph content" + (" with tracking" if track else ""),
        )

    def _replace_cell_content(
        self,
        element: etree._Element,
        new_text: str,
        track: bool,
        author: str,
    ) -> EditResult:
        """Replace content of a table cell."""
        # Get the first paragraph (or create one)
        paragraphs = list(element.findall(f"./{{{WORD_NAMESPACE}}}p"))

        if paragraphs:
            # Replace content of first paragraph
            return self._replace_paragraph_content(paragraphs[0], new_text, track, author)
        else:
            # Create a new paragraph with the content
            para = etree.SubElement(element, f"{{{WORD_NAMESPACE}}}p")
            run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = new_text
            # Add xml:space="preserve" for whitespace preservation
            if new_text and (new_text[0].isspace() or new_text[-1].isspace()):
                t.set(f"{{{XML_NAMESPACE}}}space", "preserve")

            return EditResult(
                success=True,
                edit_type="replace_at_ref",
                message="Created cell content",
            )

    def replace_in_ref(
        self,
        ref: str,
        find: str,
        replace: str,
        occurrence: int | list[int] | str = "first",
        track: bool = False,
        author: str | None = None,
    ) -> EditResult:
        """Replace text within a specific element identified by ref.

        Unlike replace_at_ref() which replaces the ENTIRE element content,
        this method performs a substring replacement within the element.
        The search is scoped to only the element identified by the ref.

        Args:
            ref: Element reference (e.g., "p:15", "tbl:0/row:1/cell:2")
            find: Text to find within the element
            replace: Replacement text
            occurrence: Which occurrence(s) to replace:
                - 1, 2, etc.: Replace the Nth occurrence (1-indexed)
                - "first": Replace first occurrence (default)
                - "last": Replace last occurrence
                - "all": Replace all occurrences
                - [1, 3, 5]: Replace specific occurrences
            track: If True, show as tracked change (default: False)
            author: Author for tracked change (uses document author if None)

        Returns:
            EditResult indicating success/failure

        Raises:
            RefNotFoundError: If the ref cannot be resolved
            TextNotFoundError: If the find text is not found in the element
            AmbiguousTextError: If multiple occurrences found and occurrence
                not specified

        Example:
            >>> doc = Document("contract.docx")
            >>> # Update a year reference in a specific paragraph
            >>> doc.replace_in_ref("p:15", "2020", "2024", track=True)
            >>>
            >>> # Replace all occurrences of a term in a table cell
            >>> doc.replace_in_ref("tbl:0/row:1/cell:0", "old", "new", occurrence="all")
        """
        from .errors import TextNotFoundError
        from .suggestions import SuggestionGenerator

        # Resolve the ref to get the element
        element = self.resolve_ref(ref)

        # Get paragraphs within the element
        # For paragraphs, it's just the element itself
        # For table cells or other containers, get all nested paragraphs
        if element.tag == f"{{{WORD_NAMESPACE}}}p":
            paragraphs = [element]
        else:
            paragraphs = list(element.iter(f"{{{WORD_NAMESPACE}}}p"))

        if not paragraphs:
            return EditResult(
                success=False,
                edit_type="replace_in_ref",
                message=f"No paragraphs found in element at ref '{ref}'",
            )

        # Get author
        author_name = author if author is not None else self.author

        # Helper function to perform a single replacement
        def do_replacement(match: "TextSpan") -> None:
            if track:
                # Tracked replace: deletion + insertion XML
                deletion_xml = self._xml_generator.create_deletion(match.text, author_name)
                insertion_xml = self._xml_generator.create_insertion(replace, author_name)
                elements = self._tracked_ops._parse_xml_elements(
                    f"{deletion_xml}\n    {insertion_xml}"
                )
                self._tracked_ops._replace_match_with_elements(match, elements)
            else:
                # Untracked replace: just replace with plain runs
                source_run = match.runs[0] if match.runs else None
                new_runs = self._xml_generator.create_plain_runs(replace, source_run=source_run)
                if len(new_runs) == 1:
                    self._tracked_ops._replace_match_with_element(match, new_runs[0])
                else:
                    self._tracked_ops._replace_match_with_elements(match, new_runs)

        # Handle "all" occurrence specially - re-find after each replacement
        # because the paragraph structure changes after each edit
        if occurrence == "all":
            replacement_count = 0
            while True:
                # Re-search after each replacement to get fresh run references
                matches = self._text_search.find_text(
                    find,
                    paragraphs,
                    regex=False,
                    normalize_special_chars=True,
                )
                if not matches:
                    break
                # Replace the first occurrence
                do_replacement(matches[0])
                replacement_count += 1

            if replacement_count == 0:
                suggestions = SuggestionGenerator.generate_suggestions(find, paragraphs)
                raise TextNotFoundError(
                    find,
                    scope=f"element at ref '{ref}'",
                    suggestions=suggestions,
                )

            return EditResult(
                success=True,
                edit_type="replace_in_ref",
                message=f"Replaced {replacement_count} occurrence(s)"
                + (" with tracking" if track else ""),
            )

        # For specific occurrences, search once and select targets
        matches = self._text_search.find_text(
            find,
            paragraphs,
            regex=False,
            normalize_special_chars=True,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(find, paragraphs)
            raise TextNotFoundError(
                find,
                scope=f"element at ref '{ref}'",
                suggestions=suggestions,
            )

        # Select target matches based on occurrence parameter
        target_matches = self._tracked_ops._select_matches(matches, occurrence, find)

        # Replace each target match (process in reverse to preserve indices)
        for match in reversed(target_matches):
            do_replacement(match)

        return EditResult(
            success=True,
            edit_type="replace_in_ref",
            message=f"Replaced {len(target_matches)} occurrence(s)"
            + (" with tracking" if track else ""),
        )

    def delete_in_ref(
        self,
        ref: str,
        text: str,
        track: bool = False,
        author: str | None = None,
    ) -> EditResult:
        """Delete text within a specific element identified by ref.

        Unlike delete_ref() which deletes the ENTIRE element,
        this method deletes only the specified text within the element.

        Args:
            ref: Element reference (e.g., "p:15")
            text: Text to delete
            track: If True, show as tracked change
            author: Author for tracked change

        Returns:
            EditResult indicating success/failure

        Raises:
            RefNotFoundError: If the ref cannot be resolved
            TextNotFoundError: If the text is not found within the element
            AmbiguousTextError: If multiple occurrences of text are found

        Example:
            >>> doc.delete_in_ref("p:15", "DRAFT - ", track=True)
        """
        from .errors import AmbiguousTextError, TextNotFoundError

        # Resolve the ref
        element = self.resolve_ref(ref)

        # Get paragraphs within this element
        # If the element is itself a paragraph, use it directly
        if element.tag == f"{{{WORD_NAMESPACE}}}p":
            paragraphs = [element]
        else:
            # Otherwise, find all paragraphs within the element
            paragraphs = list(element.findall(f".//{{{WORD_NAMESPACE}}}p"))

        if not paragraphs:
            raise TextNotFoundError(
                text,
                hint=f"Element at ref '{ref}' contains no paragraphs",
            )

        # Find the text within these paragraphs
        matches = self._text_search.find_text(
            text,
            paragraphs,
            regex=False,
            normalize_special_chars=True,
        )

        if not matches:
            # Get element text for error message
            element_text = self.get_text_at_ref(ref)
            preview = element_text[:100] + "..." if len(element_text) > 100 else element_text
            raise TextNotFoundError(
                text,
                hint=f"Text not found in element at ref '{ref}'. Element contains: '{preview}'",
            )

        if len(matches) > 1:
            raise AmbiguousTextError(text, matches)

        # Get the single match
        match = matches[0]

        # Get author
        author_name = author if author is not None else self.author

        if track:
            # Tracked deletion: wrap in <w:del>
            deletion_xml = self._xml_generator.create_deletion(match.text, author_name)
            elements = self._tracked_ops._parse_xml_elements(deletion_xml)
            deletion_element = elements[0]
            self._tracked_ops._replace_match_with_element(match, deletion_element)
        else:
            # Untracked deletion: simply remove the matched text
            self._tracked_ops._remove_match(match)

        # Invalidate registry cache
        self._ref_registry.invalidate()

        return EditResult(
            success=True,
            edit_type="delete_in_ref",
            message=f"Deleted text '{text}'" + (" with tracking" if track else ""),
        )

    def add_comment_at_ref(
        self,
        ref: str,
        comment: str,
        author: str | None = None,
    ) -> "Comment":
        """Add a comment anchored to the element at ref.

        This method adds a comment that covers the entire text content of the
        element identified by the ref. For more granular comment placement,
        use the text-based add_comment() method.

        Args:
            ref: Ref path to the element to comment on (e.g., "p:5")
            comment: The comment text
            author: Author for the comment (uses document author if None)

        Returns:
            The created Comment object

        Raises:
            RefNotFoundError: If the ref cannot be resolved

        Example:
            >>> doc = Document("contract.docx")
            >>> comment = doc.add_comment_at_ref("p:15", "Please review this clause")
            >>> print(comment.id)  # Comment ID for reference
        """
        # Resolve the ref
        element = self.resolve_ref(ref)

        # Get the text content of the element
        text_parts = []
        for t_elem in element.iter(f"{{{WORD_NAMESPACE}}}t"):
            if t_elem.text:
                text_parts.append(t_elem.text)
        element_text = "".join(text_parts)

        if not element_text:
            # If no text, use a placeholder approach
            # Find the first text element or create one
            first_t = element.find(f".//{{{WORD_NAMESPACE}}}t")
            if first_t is not None and first_t.text:
                element_text = first_t.text
            else:
                # Create a minimal run with space so comment has an anchor
                run = element.find(f".//{{{WORD_NAMESPACE}}}r")
                if run is None:
                    run = etree.SubElement(element, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = " "
                element_text = " "

        # Use the existing add_comment method with the element's text
        # This finds the text and wraps it with comment markers
        return self.add_comment(
            text=comment,
            on=element_text,
            author=author,
            occurrence=1,  # Use first occurrence
        )

    def __del__(self) -> None:
        """Clean up package resources on object destruction."""
        if self._package is not None:
            self._package.close()

    def __enter__(self) -> "Document":
        """Context manager support."""
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        """Context manager cleanup."""
        if self._package is not None:
            self._package.close()


def compare_documents(
    original: str | Path | bytes | BinaryIO,
    modified: str | Path | bytes | BinaryIO,
    author: str = "Comparison",
    minimal_edits: bool = False,
) -> Document:
    """Compare two documents and return a redline document with tracked changes.

    This is a convenience function that loads two documents, compares them,
    and returns a new Document containing tracked changes showing the differences.

    Args:
        original: Path, bytes, or file object for the original document
        modified: Path, bytes, or file object for the modified document
        author: Author name for the tracked changes (default: "Comparison")
        minimal_edits: If True, use word-level diffs for cleaner legal-style redlines
            (default: False)

    Returns:
        A new Document with tracked changes showing differences

    Example:
        >>> redline = compare_documents("contract_v1.docx", "contract_v2.docx")
        >>> redline.save("contract_redline.docx")

        # With minimal edits for legal-style redlines:
        >>> redline = compare_documents(
        ...     "contract_v1.docx",
        ...     "contract_v2.docx",
        ...     author="Review",
        ...     minimal_edits=True
        ... )
    """
    # Load original document
    original_doc = Document(original, author=author)

    # Load modified document
    modified_doc = Document(modified)

    # Create a copy of original to apply changes to (via serialize/reload)
    original_bytes = original_doc.save_to_bytes(validate=False)
    redline = Document(original_bytes, author=author)

    # Apply comparison
    redline.compare_to(modified_doc, author=author, minimal_edits=minimal_edits)

    return redline
