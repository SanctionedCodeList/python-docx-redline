"""
Document class for editing Word documents with tracked changes.

This module provides the main Document class which handles loading .docx files,
inserting tracked changes, and saving the modified documents.
"""

import io
import logging
from pathlib import Path
from typing import TYPE_CHECKING, Any, BinaryIO

if TYPE_CHECKING:
    from python_docx_redline.models.comment import Comment
    from python_docx_redline.models.footnote import Endnote, Footnote
    from python_docx_redline.models.header_footer import Footer, Header
    from python_docx_redline.models.paragraph import Paragraph
    from python_docx_redline.models.section import Section
    from python_docx_redline.models.table import Table, TableRow
    from python_docx_redline.models.tracked_change import TrackedChange

from lxml import etree

from .author import AuthorIdentity
from .constants import WORD_NAMESPACE
from .operations.batch import BatchOperations
from .operations.change_management import ChangeManagement
from .operations.comments import CommentOperations
from .operations.comparison import ComparisonOperations
from .operations.formatting import FormatOperations
from .operations.header_footer import HeaderFooterOperations
from .operations.notes import NoteOperations
from .operations.patterns import PatternOperations
from .operations.section import SectionOperations
from .operations.tables import TableOperations
from .operations.tracked_changes import TrackedChangeOperations
from .package import OOXMLPackage
from .results import ComparisonStats, EditResult, FormatResult
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

    def get_text(self) -> str:
        """Extract all visible text content from the document.

        Returns plain text with paragraphs separated by double newlines.
        This excludes deleted text (tracked changes) and only returns visible content.
        This is useful for understanding document content before making edits.

        Returns:
            Plain text content of the entire document (excluding deletions)

        Example:
            >>> doc = Document("contract.docx")
            >>> text = doc.get_text()
            >>> if "confidential" in text.lower():
            ...     print("Document contains confidential information")
        """
        # Extract only visible text (w:t), not deleted text (w:delText)
        paragraphs_text = []
        for para in self.paragraphs:
            # Get only w:t elements, not w:delText
            text_elements = para.element.findall(f".//{{{WORD_NAMESPACE}}}t")
            para_text = "".join(elem.text or "" for elem in text_elements)
            paragraphs_text.append(para_text)
        return "\n\n".join(paragraphs_text)

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

    def insert_tracked(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Insert text with tracked changes after or before a specific location.

        This method searches for the anchor text in the document and inserts
        the new text either immediately after it or immediately before it as
        a tracked insertion.

        Args:
            text: The text to insert
            after: The text or regex pattern to insert after (optional)
            before: The text or regex pattern to insert before (optional)
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            regex: Whether to treat anchor as a regex pattern (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)

        Raises:
            ValueError: If both 'after' and 'before' are specified, or if neither is specified
            TextNotFoundError: If the anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
            re.error: If regex=True and the pattern is invalid
        """
        self._tracked_ops.insert(
            text,
            after=after,
            before=before,
            author=author,
            scope=scope,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
        )

    def delete_tracked(
        self,
        text: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Delete text with tracked changes.

        This method searches for the specified text in the document and marks
        it as a tracked deletion.

        Args:
            text: The text or regex pattern to delete
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            regex: Whether to treat 'text' as a regex pattern (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)

        Raises:
            TextNotFoundError: If the text is not found
            AmbiguousTextError: If multiple occurrences of text are found
            re.error: If regex=True and the pattern is invalid
        """
        self._tracked_ops.delete(
            text,
            author=author,
            scope=scope,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
        )

    def replace_tracked(
        self,
        find: str,
        replace: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
        show_context: bool = False,
        check_continuity: bool = False,
        context_chars: int = 50,
    ) -> None:
        """Find and replace text with tracked changes.

        This method searches for text and replaces it with new text, showing
        both the deletion of the old text and insertion of the new text as
        tracked changes.

        When regex=True, the replacement string can use capture groups:
        - \\1, \\2, etc. for numbered groups
        - \\g<name> for named groups

        Args:
            find: Text or regex pattern to find
            replace: Replacement text (can include capture group references if regex=True)
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})
            regex: Whether to treat 'find' as a regex pattern (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)
            show_context: Show text before/after the match for preview (default: False)
            check_continuity: Check if replacement may create sentence fragments (default: False)
            context_chars: Number of characters to show before/after when show_context=True
                (default: 50)

        Raises:
            TextNotFoundError: If the 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
            re.error: If regex=True and the pattern is invalid

        Warnings:
            ContinuityWarning: If check_continuity=True and potential sentence fragment detected

        Example:
            >>> # Simple replacement
            >>> doc.replace_tracked("30 days", "45 days")
            >>>
            >>> # Regex with capture groups
            >>> doc.replace_tracked(r"(\\d+) days", r"\\1 business days", regex=True)
            >>>
            >>> # With context preview
            >>> doc.replace_tracked(
            ...     "old text", "new text",
            ...     show_context=True,
            ...     context_chars=100
            ... )
            >>>
            >>> # With continuity checking
            >>> doc.replace_tracked(
            ...     "sentence one.", "replacement.",
            ...     check_continuity=True
            ... )
        """
        self._tracked_ops.replace(
            find,
            replace,
            author=author,
            scope=scope,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
            show_context=show_context,
            check_continuity=check_continuity,
            context_chars=context_chars,
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
        enable_quote_normalization: bool = True,
    ) -> None:
        """Move text to a new location with proper move tracking.

        Unlike delete + insert, move tracking creates linked markers that show
        the text was relocated rather than deleted and re-added. This provides
        better context for document reviewers in Word.

        In Word's track changes view:
        - Source location shows text with strikethrough and "Moved" annotation
        - Destination shows text with underline and "Moved" annotation
        - Both locations are linked with matching move markers

        Args:
            text: The text to move (or regex pattern if regex=True)
            after: Text to insert the moved content after (at destination)
            before: Text to insert the moved content before (at destination)
            author: Optional author override (uses document author if None)
            source_scope: Limit source text search scope
            dest_scope: Limit destination anchor search scope
            regex: Whether to treat 'text' and anchor as regex patterns (default: False)
            enable_quote_normalization: Auto-convert straight quotes to smart quotes for
                matching (default: True)

        Raises:
            ValueError: If both 'after' and 'before' are specified, or if neither is specified
            TextNotFoundError: If the source text or destination anchor is not found
            AmbiguousTextError: If multiple occurrences of source text or anchor are found
            re.error: If regex=True and a pattern is invalid

        Example:
            >>> # Move "Section A" to after "Table of Contents"
            >>> doc.move_tracked(
            ...     "Section A: Introduction",
            ...     after="Table of Contents",
            ...     author="Editor"
            ... )
            >>>
            >>> # Move text to before another location
            >>> doc.move_tracked(
            ...     "Important Note",
            ...     before="Conclusion",
            ...     source_scope="section:Appendix"
            ... )
        """
        self._tracked_ops.move(
            text,
            after=after,
            before=before,
            author=author,
            source_scope=source_scope,
            dest_scope=dest_scope,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
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
        enable_quote_normalization: bool = True,
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
            enable_quote_normalization: Auto-convert straight quotes to smart quotes
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
            enable_quote_normalization=enable_quote_normalization,
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
    ) -> "Comment":
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

        Returns:
            The created Comment object

        Raises:
            TextNotFoundError: If the target text is not found (new comments only)
            AmbiguousTextError: If multiple occurrences of target text are found
            ValueError: If neither 'on' nor 'reply_to' is provided, or if
                        reply_to references a non-existent comment
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

    def save(self, output_path: str | Path | None = None, validate: bool = True) -> None:
        """Save the document to a file.

        Validates the document structure before saving to ensure OOXML compliance
        and prevent broken Word files in production.

        Args:
            output_path: Path to save the document. If None, saves to original path.
                        For in-memory documents (loaded from bytes), output_path is required.
            validate: Whether to run full OOXML validation before saving (default: True).
                     Validation is strongly recommended to catch errors before production.
                     Set to False for in-memory documents without an original file.

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
            raise ValidationError(f"Failed to save document: {e}") from e

    def save_to_bytes(self, validate: bool = True) -> bytes:
        """Save the document to bytes (in-memory).

        This is useful for:
        - Passing documents between libraries without filesystem
        - Storing documents in databases
        - Sending documents over network

        Args:
            validate: Whether to run OOXML validation (default: True).
                     Set to False for in-memory documents without an original file,
                     as validation compares against the original.

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
            return self._package.save_to_bytes()

        except ValidationError:
            raise
        except Exception as e:
            raise ValidationError(f"Failed to save document to bytes: {e}") from e

    def apply_edits(
        self, edits: list[dict[str, Any]], stop_on_error: bool = False
    ) -> list[EditResult]:
        """Apply multiple edits in sequence.

        This method processes a list of edit specifications and applies each one
        in order. Each edit is a dictionary specifying the edit type and parameters.

        Args:
            edits: List of edit dictionaries with keys:
                - type: Edit operation ("insert_tracked", "replace_tracked", "delete_tracked")
                - Other parameters specific to the edit type
            stop_on_error: If True, stop processing on first error

        Returns:
            List of EditResult objects, one per edit

        Example:
            >>> edits = [
            ...     {
            ...         "type": "insert_tracked",
            ...         "text": "new text",
            ...         "after": "anchor",
            ...         "scope": "section:Introduction"
            ...     },
            ...     {
            ...         "type": "replace_tracked",
            ...         "find": "old",
            ...         "replace": "new"
            ...     }
            ... ]
            >>> results = doc.apply_edits(edits)
            >>> print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
        """
        return self._batch_ops.apply_edits(edits, stop_on_error=stop_on_error)

    def apply_edit_file(
        self, path: str | Path, format: str = "yaml", stop_on_error: bool = False
    ) -> list[EditResult]:
        """Apply edits from a YAML or JSON file.

        Loads edit specifications from a file and applies them using apply_edits().
        The file should contain an 'edits' key with a list of edit dictionaries.

        Args:
            path: Path to the edit specification file
            format: File format - "yaml" or "json" (default: "yaml")
            stop_on_error: If True, stop processing on first error

        Returns:
            List of EditResult objects, one per edit

        Raises:
            ValidationError: If file cannot be parsed or has invalid format
            FileNotFoundError: If file does not exist

        Example YAML file:
            ```yaml
            edits:
              - type: insert_tracked
                text: "new text"
                after: "anchor"
              - type: replace_tracked
                find: "old"
                replace: "new"
            ```

        Example:
            >>> results = doc.apply_edit_file("edits.yaml")
            >>> print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
        """
        return self._batch_ops.apply_edit_file(path, format=format, stop_on_error=stop_on_error)

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
        enable_quote_normalization: bool = True,
    ) -> None:
        """Replace text in a header with tracked changes.

        Args:
            find: Text to find
            replace: Replacement text
            header_type: Type of header ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat 'find' as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching

        Raises:
            TextNotFoundError: If 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
            ValueError: If no header of the specified type exists

        Example:
            >>> doc.replace_in_header("Draft", "Final", header_type="default")
        """
        self._header_footer_ops.replace_in_header(
            find=find,
            replace=replace,
            header_type=header_type,
            author=author,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
        )

    def replace_in_footer(
        self,
        find: str,
        replace: str,
        footer_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Replace text in a footer with tracked changes.

        Args:
            find: Text to find
            replace: Replacement text
            footer_type: Type of footer ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat 'find' as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching

        Raises:
            TextNotFoundError: If 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
            ValueError: If no footer of the specified type exists

        Example:
            >>> doc.replace_in_footer("Page {PAGE}", "Page {PAGE} of {NUMPAGES}")
        """
        self._header_footer_ops.replace_in_footer(
            find=find,
            replace=replace,
            footer_type=footer_type,
            author=author,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
        )

    def insert_in_header(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        header_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Insert text in a header with tracked changes.

        Args:
            text: Text to insert
            after: Text to insert after
            before: Text to insert before
            header_type: Type of header ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat anchor as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching

        Raises:
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
            ValueError: If no header of the specified type exists, or if both
                        'after' and 'before' are specified

        Example:
            >>> doc.insert_in_header(" - Final", after="Document Title")
        """
        self._header_footer_ops.insert_in_header(
            text=text,
            after=after,
            before=before,
            header_type=header_type,
            author=author,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
        )

    def insert_in_footer(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        footer_type: str = "default",
        author: str | None = None,
        regex: bool = False,
        enable_quote_normalization: bool = True,
    ) -> None:
        """Insert text in a footer with tracked changes.

        Args:
            text: Text to insert
            after: Text to insert after
            before: Text to insert before
            footer_type: Type of footer ("default", "first", or "even")
            author: Optional author override
            regex: Whether to treat anchor as a regex pattern
            enable_quote_normalization: Auto-convert quotes for matching

        Raises:
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
            ValueError: If no footer of the specified type exists, or if both
                        'after' and 'before' are specified

        Example:
            >>> doc.insert_in_footer(" - Confidential", after="Page")
        """
        self._header_footer_ops.insert_in_footer(
            text=text,
            after=after,
            before=before,
            footer_type=footer_type,
            author=author,
            regex=regex,
            enable_quote_normalization=enable_quote_normalization,
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
