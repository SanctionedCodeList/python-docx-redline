"""
Document class for editing Word documents with tracked changes.

This module provides the main Document class which handles loading .docx files,
inserting tracked changes, and saving the modified documents.
"""

import difflib
import io
import logging
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, BinaryIO

if TYPE_CHECKING:
    from python_docx_redline.models.comment import Comment
    from python_docx_redline.models.footnote import Endnote, Footnote
    from python_docx_redline.models.paragraph import Paragraph
    from python_docx_redline.models.section import Section
    from python_docx_redline.models.table import Table, TableRow

import yaml
from lxml import etree

from .author import AuthorIdentity
from .errors import AmbiguousTextError, TextNotFoundError
from .format_builder import ParagraphPropertyBuilder, RunPropertyBuilder
from .minimal_diff import (
    apply_minimal_edits_to_paragraph,
    should_use_minimal_editing,
)
from .results import EditResult, FormatResult
from .scope import ScopeEvaluator
from .suggestions import SuggestionGenerator
from .text_search import TextSearch
from .tracked_xml import TrackedXMLGenerator
from .validation import ValidationError

logger = logging.getLogger(__name__)

# Word namespace
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": WORD_NAMESPACE}


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

        self._temp_dir: Path | None = None
        self._is_zip = False

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
        to a temporary directory. The main document.xml is then parsed.

        Supports loading from file paths or in-memory streams (BytesIO).

        Raises:
            ValidationError: If the document cannot be loaded
        """
        # Determine source: stream or path
        if self._source_stream is not None:
            zip_source: Path | BinaryIO = self._source_stream
            source_desc = "<in-memory document>"
        else:
            assert self.path is not None
            if not self.path.exists():
                raise ValidationError(f"Document not found: {self.path}")
            zip_source = self.path
            source_desc = str(self.path)

        # Check if it's a ZIP file (.docx)
        try:
            if zipfile.is_zipfile(zip_source):
                self._is_zip = True
                # Reset stream position if it was checked by is_zipfile
                if self._source_stream is not None and hasattr(self._source_stream, "seek"):
                    self._source_stream.seek(0)
                self._extract_docx(zip_source)
            else:
                if self._source_stream is not None:
                    raise ValidationError("In-memory source must be a valid .docx (ZIP) file")
                # Assume it's already an unpacked XML file
                self._is_zip = False
                self._temp_dir = self.path.parent  # type: ignore
        except ValidationError:
            raise
        except Exception as e:
            raise ValidationError(f"Failed to load document: {e}") from e

        # Parse the document.xml
        try:
            if self._is_zip:
                document_xml = self._temp_dir / "word" / "document.xml"  # type: ignore
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

    def _extract_docx(self, source: Path | BinaryIO) -> None:
        """Extract the .docx ZIP archive to a temporary directory.

        Args:
            source: Path to the .docx file or a file-like object containing it
        """
        self._temp_dir = Path(tempfile.mkdtemp(prefix="python_docx_redline_"))

        try:
            with zipfile.ZipFile(source, "r") as zip_ref:
                zip_ref.extractall(self._temp_dir)
        except Exception as e:
            # Clean up on failure
            if self._temp_dir and self._temp_dir.exists():
                shutil.rmtree(self._temp_dir)
            raise ValidationError(f"Failed to extract .docx file: {e}") from e

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
        from python_docx_redline.models.table import Table

        return [Table(tbl) for tbl in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}tbl")]

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
        for table in self.tables:
            if table.contains(containing, case_sensitive):
                return table
        return None

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
        from python_docx_redline.models.comment import Comment

        comments_xml = self._load_comments_xml()
        if comments_xml is None:
            return []

        # Build mapping of comment ID -> marked text from document body
        range_map = self._build_comment_ranges()

        # Parse comments
        result = []
        for comment_elem in comments_xml.findall(f".//{{{WORD_NAMESPACE}}}comment"):
            comment_id = comment_elem.get(f"{{{WORD_NAMESPACE}}}id", "")
            range_info = range_map.get(comment_id)
            result.append(Comment(comment_elem, range_info, document=self))

        return result

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
        all_comments = self.comments

        if author:
            all_comments = [c for c in all_comments if c.author == author]

        if scope:
            evaluator = ScopeEvaluator.parse(scope)
            # Filter to comments whose marked text falls within scope
            # Pass the underlying XML element to the evaluator
            all_comments = [
                c for c in all_comments if c.range and evaluator(c.range.start_paragraph.element)
            ]

        return all_comments

    def _load_comments_xml(self) -> etree._Element | None:
        """Load word/comments.xml if it exists.

        Returns:
            Root element of comments.xml or None if not present
        """
        if not self._is_zip or not self._temp_dir:
            return None

        comments_path = self._temp_dir / "word" / "comments.xml"
        if not comments_path.exists():
            return None

        tree = etree.parse(str(comments_path))
        return tree.getroot()

    def _build_comment_ranges(self) -> dict[str, Any]:
        """Build a mapping of comment ID to marked text range.

        Scans the document for commentRangeStart/End markers and
        extracts the text between them.

        Returns:
            Dict mapping comment ID to CommentRange
        """
        from python_docx_redline.models.comment import CommentRange

        ranges: dict[str, CommentRange] = {}

        # Find all comment range starts
        for start_elem in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeStart"):
            comment_id = start_elem.get(f"{{{WORD_NAMESPACE}}}id", "")
            if not comment_id:
                continue

            # Find matching end
            end_elem = None
            for elem in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeEnd"):
                if elem.get(f"{{{WORD_NAMESPACE}}}id") == comment_id:
                    end_elem = elem
                    break

            # Extract text between start and end
            marked_text = self._extract_text_in_range(start_elem, end_elem)

            # Find containing paragraphs
            start_para = self._find_containing_paragraph(start_elem)
            end_para = (
                self._find_containing_paragraph(end_elem) if end_elem is not None else start_para
            )

            if start_para is not None:
                ranges[comment_id] = CommentRange(
                    start_paragraph=start_para,
                    end_paragraph=end_para or start_para,
                    marked_text=marked_text,
                )

        return ranges

    def _extract_text_in_range(
        self,
        start_elem: etree._Element,
        end_elem: etree._Element | None,
    ) -> str:
        """Extract text between comment range markers.

        Walks through the XML tree between start and end markers,
        collecting text from w:t elements.

        Args:
            start_elem: The commentRangeStart element
            end_elem: The commentRangeEnd element (or None for point comments)

        Returns:
            The text content between the markers
        """
        if end_elem is None:
            return ""

        # Get the document body to iterate through
        body = self.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return ""

        # Collect all text elements in document order
        text_parts = []
        in_range = False

        # We need to walk the tree in document order
        for elem in body.iter():
            if elem is start_elem:
                in_range = True
                continue

            if elem is end_elem:
                break

            if in_range and elem.tag == f"{{{WORD_NAMESPACE}}}t":
                if elem.text:
                    text_parts.append(elem.text)

        return "".join(text_parts)

    def _find_containing_paragraph(self, elem: etree._Element) -> "Paragraph | None":
        """Find the paragraph containing an element.

        Walks up the tree to find the w:p ancestor.

        Args:
            elem: The element to find the containing paragraph for

        Returns:
            Paragraph wrapper or None if not found
        """
        from python_docx_redline.models.paragraph import Paragraph

        current = elem
        while current is not None:
            parent = current.getparent()
            if parent is None:
                break
            if parent.tag == f"{{{WORD_NAMESPACE}}}p":
                return Paragraph(parent)
            current = parent

        return None

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
        # Validate parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        # Determine anchor text and insertion mode
        anchor: str = after if after is not None else before  # type: ignore[assignment]
        insert_after = after is not None

        # Get all paragraphs in the document
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search with optional quote normalization
        matches = self._text_search.find_text(
            anchor,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not matches:
            # Generate smart suggestions
            suggestions = SuggestionGenerator.generate_suggestions(anchor, paragraphs)
            raise TextNotFoundError(anchor, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(anchor, matches)

        # We have exactly one match
        match = matches[0]

        # Generate the insertion XML
        insertion_xml = self._xml_generator.create_insertion(text, author)

        # Parse the insertion XML with namespace context
        # Need to wrap it with namespace declarations
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {insertion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        insertion_element = root[0]  # Get the first child (the actual insertion)

        # Insert at the appropriate position
        if insert_after:
            self._insert_after_match(match, insertion_element)
        else:
            self._insert_before_match(match, insertion_element)

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
        # Get all paragraphs in the document
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search with optional quote normalization
        matches = self._text_search.find_text(
            text,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not matches:
            # Generate smart suggestions
            suggestions = SuggestionGenerator.generate_suggestions(text, paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(text, matches)

        # We have exactly one match
        match = matches[0]

        # Generate the deletion XML
        deletion_xml = self._xml_generator.create_deletion(text, author)

        # Parse the deletion XML with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        deletion_element = root[0]  # Get the first child (the actual deletion)

        # Replace the matched text with deletion
        self._replace_match_with_element(match, deletion_element)

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
        # Get all paragraphs in the document
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search with optional quote normalization
        matches = self._text_search.find_text(
            find,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not matches:
            # Generate smart suggestions
            suggestions = SuggestionGenerator.generate_suggestions(find, paragraphs)
            raise TextNotFoundError(find, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(find, matches)

        # We have exactly one match
        match = matches[0]

        # Get the actual matched text for deletion
        matched_text = match.text

        # Show context preview if requested
        if show_context:
            before, matched, after = self._get_detailed_context(match, context_chars)
            print("\n" + "=" * 80)
            print("CONTEXT PREVIEW")
            print("=" * 80)
            print(f"\nBEFORE ({len(before)} chars):")
            print(f"  {repr(before)}")
            print(f"\nMATCH ({len(matched)} chars):")
            print(f"  {repr(matched)}")
            print(f"\nAFTER ({len(after)} chars):")
            print(f"  {repr(after)}")
            print(f"\nREPLACEMENT ({len(replace)} chars):")
            print(f"  {repr(replace)}")
            print("\n" + "=" * 80 + "\n")

        # Handle capture group expansion for regex replacements
        if regex and match.match_obj:
            # Use expand() to handle capture group references like \1, \2, etc.
            replacement_text = match.match_obj.expand(replace)
        else:
            replacement_text = replace

        # Check continuity if requested
        if check_continuity:
            _, _, after_text = self._get_detailed_context(match, context_chars)
            continuity_warnings = self._check_continuity(replacement_text, after_text)

            if continuity_warnings:
                import warnings

                from .errors import ContinuityWarning

                for warning_msg in continuity_warnings:
                    suggestions = [
                        "Include more context in your replacement text",
                        "Adjust the 'find' text to include the connecting phrase",
                        "Review the following text to ensure grammatical correctness",
                    ]
                    warnings.warn(
                        ContinuityWarning(warning_msg, after_text, suggestions),
                        stacklevel=2,
                    )

        # Generate deletion XML for the old text (use actual matched text)
        deletion_xml = self._xml_generator.create_deletion(matched_text, author)

        # Generate insertion XML for the new text (with expanded capture groups if regex)
        insertion_xml = self._xml_generator.create_insertion(replacement_text, author)

        # Parse both XMLs with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
    {insertion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        deletion_element = root[0]  # First child (deletion)
        insertion_element = root[1]  # Second child (insertion)

        # Replace the matched text with deletion + insertion
        self._replace_match_with_elements(match, [deletion_element, insertion_element])

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
        # Validate parameters
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before' parameters")
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before' parameter")

        # Determine destination anchor and insertion mode
        dest_anchor: str = after if after is not None else before  # type: ignore[assignment]
        insert_after = after is not None

        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # 1. Find the source text to move
        source_paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, source_scope)
        source_matches = self._text_search.find_text(
            text,
            source_paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not source_matches:
            suggestions = SuggestionGenerator.generate_suggestions(text, source_paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        if len(source_matches) > 1:
            raise AmbiguousTextError(text, source_matches)

        source_match = source_matches[0]
        source_text = source_match.text

        # 2. Find the destination anchor
        dest_paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, dest_scope)
        dest_matches = self._text_search.find_text(
            dest_anchor,
            dest_paragraphs,
            regex=regex,
            normalize_quotes_for_matching=enable_quote_normalization and not regex,
        )

        if not dest_matches:
            suggestions = SuggestionGenerator.generate_suggestions(dest_anchor, dest_paragraphs)
            raise TextNotFoundError(dest_anchor, suggestions=suggestions)

        if len(dest_matches) > 1:
            raise AmbiguousTextError(dest_anchor, dest_matches)

        dest_match = dest_matches[0]

        # 3. Generate a unique move name to link source and destination
        move_name = self._generate_move_name()

        # 4. Generate moveFrom XML (for source location)
        move_from_xml, _, _ = self._xml_generator.create_move_from(source_text, move_name, author)

        # 5. Generate moveTo XML (for destination location)
        move_to_xml, _, _ = self._xml_generator.create_move_to(source_text, move_name, author)

        # 6. Parse both XMLs with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {move_to_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        # Get all three elements (moveToRangeStart, moveTo, moveToRangeEnd)
        move_to_elements = list(root)

        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {move_from_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        # Get all three elements (moveFromRangeStart, moveFrom, moveFromRangeEnd)
        move_from_elements = list(root)

        # 7. First, insert the moveTo at the destination
        # (do this first so we don't mess up source position)
        if insert_after:
            self._insert_after_match(dest_match, move_to_elements)
        else:
            self._insert_before_match(dest_match, move_to_elements)

        # 8. Replace the source text with moveFrom markers
        self._replace_match_with_elements(source_match, move_from_elements)

    def _generate_move_name(self) -> str:
        """Generate a unique move name for linking source and destination.

        Scans the document for existing move names and returns the next one.

        Returns:
            Unique move name (e.g., "move1", "move2", etc.)
        """
        # Find all existing move names
        existing_names: set[str] = set()

        # Check moveFromRangeStart elements
        for elem in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}moveFromRangeStart"):
            name = elem.get(f"{{{WORD_NAMESPACE}}}name")
            if name:
                existing_names.add(name)

        # Check moveToRangeStart elements
        for elem in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}moveToRangeStart"):
            name = elem.get(f"{{{WORD_NAMESPACE}}}name")
            if name:
                existing_names.add(name)

        # Find next available number
        counter = 1
        while f"move{counter}" in existing_names:
            counter += 1

        return f"move{counter}"

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
        # Build regex pattern for currency amounts
        # Matches: $100, $100.00, $1,000, $1,000.50, etc.
        pattern = rf"{re.escape(currency_symbol)}\d{{1,3}}(?:,?\d{{3}})*(?:\.\d+)?"

        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find all currency matches
        matches = self._text_search.find_text(
            pattern,
            paragraphs,
            regex=True,
            normalize_quotes_for_matching=False,
        )

        # Helper to format amount
        def format_amount(amount_str: str) -> str:
            amount = float(amount_str.replace(",", ""))
            formatted = f"{amount:.{decimal_places}f}"
            if thousands_separator and "." in formatted:
                integer_part, decimal_part = formatted.split(".")
                integer_with_commas = f"{int(integer_part):,}"
                return f"{integer_with_commas}.{decimal_part}"
            elif thousands_separator:
                formatted_int = f"{int(float(formatted)):,}"
                if decimal_places > 0:
                    return formatted_int + "." + "0" * decimal_places
                return formatted_int
            return formatted

        #  Process one match at a time to avoid XML reference issues
        replacement_count = 0
        max_iterations = 100  # Prevent infinite loop

        for _ in range(max_iterations):
            # Get fresh paragraphs and matches each iteration
            all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
            paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)
            matches = self._text_search.find_text(
                pattern,
                paragraphs,
                regex=True,
                normalize_quotes_for_matching=False,
            )

            if not matches:
                break  # No more matches

            # Process only the first match
            match = matches[0]
            matched_text = match.text
            amount_str = matched_text[len(currency_symbol) :]

            try:
                replacement_text = f"{currency_symbol}{format_amount(amount_str)}"
            except ValueError:
                break  # Can't parse, stop

            # Skip if already correct
            if matched_text == replacement_text:
                break

            # Use existing replace logic which handles single match
            try:
                # Create exact pattern for this specific match to avoid ambiguity
                exact_pattern = re.escape(matched_text)
                self.replace_tracked(
                    find=exact_pattern,
                    replace=replacement_text,
                    author=author,
                    scope=scope,
                    regex=True,
                    enable_quote_normalization=False,
                )
                replacement_count += 1
            except (TextNotFoundError, AmbiguousTextError):
                break  # Can't replace, stop

        return replacement_count

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
        # Resolve author
        author_name = author if author is not None else self.author

        # Common date patterns with their corresponding datetime format strings
        months_long = (
            "January|February|March|April|May|June|July|August|September|October|November|December"
        )
        months_short = "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"

        date_patterns = [
            # MM/DD/YYYY or M/D/YYYY
            (r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", "%m/%d/%Y"),
            # YYYY-MM-DD
            (r"\b(\d{4})-(\d{2})-(\d{2})\b", "%Y-%m-%d"),
            # Month DD, YYYY (e.g., December 08, 2025)
            (
                rf"\b({months_long}) (\d{{1,2}}), (\d{{4}})\b",
                "%B %d, %Y",
            ),
            # Mon DD, YYYY (e.g., Dec 08, 2025)
            (
                rf"\b({months_short}) (\d{{1,2}}), (\d{{4}})\b",
                "%b %d, %Y",
            ),
            # DD Month YYYY (e.g., 08 December 2025)
            (
                rf"\b(\d{{1,2}}) ({months_long}) (\d{{4}})\b",
                "%d %B %Y",
            ),
        ]

        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        all_matches = []
        for pattern, date_format in date_patterns:
            matches = self._text_search.find_text(
                pattern,
                paragraphs,
                regex=True,
                normalize_quotes_for_matching=False,
            )
            # Store matches with their format
            for match in matches:
                all_matches.append((match, date_format))

        if not all_matches:
            return 0

        # Sort by position (reverse) to process from end to beginning
        # This prevents position invalidation issues
        all_matches.sort(
            key=lambda x: (
                list(all_paragraphs).index(x[0].paragraph),
                x[0].start_run_index,
                x[0].start_offset,
            ),
            reverse=True,
        )

        # Process each match
        replacement_count = 0
        for match, date_format in all_matches:
            matched_text = match.text

            # Parse the date using the detected format
            try:
                parsed_date = datetime.strptime(matched_text, date_format)
            except ValueError:
                continue  # Skip if parsing fails

            # Format to target format
            replacement_text = parsed_date.strftime(to_format)

            # Skip if already in correct format
            if matched_text == replacement_text:
                continue

            # Generate tracked change XML
            deletion_xml = self._xml_generator.create_deletion(matched_text, author_name)
            insertion_xml = self._xml_generator.create_insertion(replacement_text, author_name)

            # Parse XMLs with namespace context
            wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
    {insertion_xml}
</root>"""
            root = etree.fromstring(wrapped_xml.encode("utf-8"))
            deletion_element = root[0]
            insertion_element = root[1]

            # Replace the matched text with deletion + insertion
            self._replace_match_with_elements(match, [deletion_element, insertion_element])
            replacement_count += 1

        return replacement_count

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
        # Escape special regex characters in the numbers
        old_escaped = re.escape(old_number)
        new_number_text = new_number

        # Build pattern: "Section 2.1" with optional trailing punctuation
        pattern = rf"\b{re.escape(section_word)}\s+{old_escaped}\b"

        # Use replace_tracked with regex
        try:
            self.replace_tracked(
                find=pattern,
                replace=f"{section_word} {new_number_text}",
                author=author,
                scope=scope,
                regex=True,
                enable_quote_normalization=False,
            )
            return 1
        except TextNotFoundError:
            return 0
        except AmbiguousTextError:
            # Multiple occurrences - need to replace all of them
            # Fall back to manual batch replacement
            author_name = author if author is not None else self.author

            all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
            paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

            matches = self._text_search.find_text(
                pattern,
                paragraphs,
                regex=True,
                normalize_quotes_for_matching=False,
            )

            if not matches:
                return 0

            # Process in reverse order
            replacement_count = 0
            for match in reversed(matches):
                matched_text = match.text
                replacement_text = f"{section_word} {new_number_text}"

                # Generate tracked change XML
                deletion_xml = self._xml_generator.create_deletion(matched_text, author_name)
                insertion_xml = self._xml_generator.create_insertion(replacement_text, author_name)

                # Parse XMLs with namespace context
                wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
    {insertion_xml}
</root>"""
                root = etree.fromstring(wrapped_xml.encode("utf-8"))
                deletion_element = root[0]
                insertion_element = root[1]

                # Replace the matched text with deletion + insertion
                self._replace_match_with_elements(match, [deletion_element, insertion_element])
                replacement_count += 1

            return replacement_count

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
        from python_docx_redline.models.paragraph import Paragraph as ParagraphClass

        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find paragraphs containing the text
        matches = self._text_search.find_text(
            find,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=not regex,
        )

        if not matches:
            return 0

        # Get unique paragraphs (a paragraph might have multiple matches)
        unique_paragraphs = {match.paragraph for match in matches}

        # Apply style to each paragraph
        count = 0
        for para_element in unique_paragraphs:
            para = ParagraphClass(para_element)
            if para.style != style:
                para.style = style
                count += 1

        return count

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
        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find all matches
        matches = self._text_search.find_text(
            find,
            paragraphs,
            regex=regex,
            normalize_quotes_for_matching=not regex,
        )

        if not matches:
            return 0

        # Apply formatting to each match
        count = 0
        for match in matches:
            # Get the runs involved in this match
            for run_idx in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[run_idx]

                # Get or create run properties
                r_pr = run.find(f"{{{WORD_NAMESPACE}}}rPr")
                if r_pr is None:
                    r_pr = etree.Element(f"{{{WORD_NAMESPACE}}}rPr")
                    run.insert(0, r_pr)

                # Apply bold
                if bold is not None:
                    b_elem = r_pr.find(f"{{{WORD_NAMESPACE}}}b")
                    if bold:
                        if b_elem is None:
                            etree.SubElement(r_pr, f"{{{WORD_NAMESPACE}}}b")
                    else:
                        if b_elem is not None:
                            r_pr.remove(b_elem)

                # Apply italic
                if italic is not None:
                    i_elem = r_pr.find(f"{{{WORD_NAMESPACE}}}i")
                    if italic:
                        if i_elem is None:
                            etree.SubElement(r_pr, f"{{{WORD_NAMESPACE}}}i")
                    else:
                        if i_elem is not None:
                            r_pr.remove(i_elem)

                # Apply color
                if color is not None:
                    color_elem = r_pr.find(f"{{{WORD_NAMESPACE}}}color")
                    if color_elem is None:
                        color_elem = etree.SubElement(r_pr, f"{{{WORD_NAMESPACE}}}color")
                    color_elem.set(f"{{{WORD_NAMESPACE}}}val", color)

            count += 1

        return count

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
        # Build format updates dict (only non-None values)
        format_updates: dict[str, Any] = {}
        if bold is not None:
            format_updates["bold"] = bold
        if italic is not None:
            format_updates["italic"] = italic
        if underline is not None:
            format_updates["underline"] = underline
        if strikethrough is not None:
            format_updates["strikethrough"] = strikethrough
        if font_name is not None:
            format_updates["font_name"] = font_name
        if font_size is not None:
            format_updates["font_size"] = font_size
        if color is not None:
            format_updates["color"] = color
        if highlight is not None:
            format_updates["highlight"] = highlight
        if superscript is not None:
            format_updates["superscript"] = superscript
        if subscript is not None:
            format_updates["subscript"] = subscript
        if small_caps is not None:
            format_updates["small_caps"] = small_caps
        if all_caps is not None:
            format_updates["all_caps"] = all_caps

        if not format_updates:
            raise ValueError("At least one formatting property must be specified")

        # Get all paragraphs in the document
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search for the text
        matches = self._text_search.find_text(
            text,
            paragraphs,
            regex=False,
            normalize_quotes_for_matching=enable_quote_normalization,
        )

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(text, paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

        # Handle occurrence selection
        if occurrence == "first" or occurrence == 1:
            target_matches = [matches[0]]
        elif occurrence == "last":
            target_matches = [matches[-1]]
        elif occurrence == "all":
            target_matches = matches
        elif isinstance(occurrence, int) and 1 <= occurrence <= len(matches):
            target_matches = [matches[occurrence - 1]]
        elif isinstance(occurrence, int):
            raise ValueError(f"Occurrence {occurrence} out of range (1-{len(matches)})")
        elif len(matches) > 1:
            raise AmbiguousTextError(text, matches)
        else:
            target_matches = matches

        # Track results
        runs_affected = 0
        last_change_id = 0
        all_previous_formatting: list[dict[str, object]] = []
        para_index = -1

        # Import run splitting helper
        from .format_builder import get_run_text, split_run_at_offset

        # Apply formatting to each target match
        for match in target_matches:
            # Use match.paragraph directly (more reliable than getparent which may
            # return wrappers like w:hyperlink, w:ins, etc.)
            para = match.paragraph
            para_index = all_paragraphs.index(para) if para in all_paragraphs else -1

            # Build list of runs to format, handling mid-run splits
            runs_to_format = []

            for run_idx in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[run_idx]
                run_text = get_run_text(run)

                is_start = run_idx == match.start_run_index
                is_end = run_idx == match.end_run_index
                is_single = is_start and is_end

                if is_single and (match.start_offset > 0 or match.end_offset < len(run_text)):
                    # Match is within a single run - need to split at both ends
                    if match.start_offset > 0:
                        before_run, remainder = split_run_at_offset(run, match.start_offset)
                        # Insert before_run before original run
                        parent = run.getparent()
                        idx = list(parent).index(run)
                        parent.insert(idx, before_run)
                        # Now split remainder at adjusted offset
                        adjusted_end = match.end_offset - match.start_offset
                        if adjusted_end < len(run_text) - match.start_offset:
                            middle_run, after_run = split_run_at_offset(remainder, adjusted_end)
                            # Replace original with middle and after
                            parent.remove(run)
                            parent.insert(idx + 1, middle_run)
                            parent.insert(idx + 2, after_run)
                            runs_to_format.append(middle_run)
                        else:
                            parent.remove(run)
                            parent.insert(idx + 1, remainder)
                            runs_to_format.append(remainder)
                    else:
                        # Only split at end
                        middle_run, after_run = split_run_at_offset(run, match.end_offset)
                        parent = run.getparent()
                        idx = list(parent).index(run)
                        parent.remove(run)
                        parent.insert(idx, middle_run)
                        parent.insert(idx + 1, after_run)
                        runs_to_format.append(middle_run)

                elif is_start and match.start_offset > 0:
                    # Split start run - only format the part from start_offset onwards
                    before_run, after_run = split_run_at_offset(run, match.start_offset)
                    parent = run.getparent()
                    idx = list(parent).index(run)
                    parent.remove(run)
                    parent.insert(idx, before_run)
                    parent.insert(idx + 1, after_run)
                    runs_to_format.append(after_run)

                elif is_end and match.end_offset < len(run_text):
                    # Split end run - only format the part up to end_offset
                    before_run, after_run = split_run_at_offset(run, match.end_offset)
                    parent = run.getparent()
                    idx = list(parent).index(run)
                    parent.remove(run)
                    parent.insert(idx, before_run)
                    parent.insert(idx + 1, after_run)
                    runs_to_format.append(before_run)

                else:
                    # Whole run is within match - format entirely
                    runs_to_format.append(run)

            # Now apply formatting to only the runs that need it
            for run in runs_to_format:
                # Get or create run properties
                existing_rpr = run.find(f"{{{WORD_NAMESPACE}}}rPr")

                # Deep copy to capture previous state
                from copy import deepcopy

                previous_rpr = deepcopy(existing_rpr) if existing_rpr is not None else None

                # Extract previous formatting for result (per-run)
                prev_formatting = RunPropertyBuilder.extract(previous_rpr)
                all_previous_formatting.append(prev_formatting)

                # Create new rPr with merged formatting
                new_rpr = RunPropertyBuilder.merge(existing_rpr, format_updates)

                # Check if there are actual changes
                if not RunPropertyBuilder.has_changes(previous_rpr, new_rpr):
                    continue  # No-op for this run

                # Create the tracked change element (returns tuple now)
                rpr_change, last_change_id = self._xml_generator.create_run_property_change(
                    previous_rpr, author
                )

                # Append the change tracking element to the new rPr
                new_rpr.append(rpr_change)

                # Replace or insert the rPr in the run
                if existing_rpr is not None:
                    run.remove(existing_rpr)
                run.insert(0, new_rpr)

                runs_affected += 1

        return FormatResult(
            success=True,  # Operation completed without error
            changed=runs_affected > 0,  # Whether any changes were made
            text_matched=text,
            paragraph_index=para_index if len(target_matches) == 1 else -1,
            changes_applied=format_updates,
            previous_formatting=all_previous_formatting,
            change_id=last_change_id,
            runs_affected=runs_affected,
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
        # Validate at least one targeting parameter
        if containing is None and starting_with is None and ending_with is None and index is None:
            raise ValueError(
                "At least one targeting parameter required: "
                "containing, starting_with, ending_with, or index"
            )

        # Build format updates dict
        format_updates: dict[str, Any] = {}
        if alignment is not None:
            format_updates["alignment"] = alignment
        if spacing_before is not None:
            format_updates["spacing_before"] = spacing_before
        if spacing_after is not None:
            format_updates["spacing_after"] = spacing_after
        if line_spacing is not None:
            format_updates["line_spacing"] = line_spacing
        if indent_left is not None:
            format_updates["indent_left"] = indent_left
        if indent_right is not None:
            format_updates["indent_right"] = indent_right
        if indent_first_line is not None:
            format_updates["indent_first_line"] = indent_first_line
        if indent_hanging is not None:
            format_updates["indent_hanging"] = indent_hanging

        if not format_updates:
            raise ValueError("At least one formatting property must be specified")

        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find target paragraph
        target_para = None
        para_index = -1

        if index is not None:
            if 0 <= index < len(paragraphs):
                target_para = paragraphs[index]
                para_index = (
                    all_paragraphs.index(target_para) if target_para in all_paragraphs else index
                )
            else:
                raise ValueError(f"Paragraph index {index} out of range (0-{len(paragraphs)-1})")
        else:
            # Search for paragraph by text content
            for i, para in enumerate(paragraphs):
                para_text = self._get_paragraph_text(para)

                if containing is not None and containing not in para_text:
                    continue
                if starting_with is not None and not para_text.startswith(starting_with):
                    continue
                if ending_with is not None and not para_text.endswith(ending_with):
                    continue

                target_para = para
                para_index = all_paragraphs.index(para) if para in all_paragraphs else i
                break

        if target_para is None:
            search_text = containing or starting_with or ending_with or ""
            raise TextNotFoundError(
                search_text,
                suggestions=["Check paragraph content", "Try a different search term"],
            )

        # Get or create paragraph properties
        from copy import deepcopy

        existing_ppr = target_para.find(f"{{{WORD_NAMESPACE}}}pPr")
        previous_ppr = deepcopy(existing_ppr) if existing_ppr is not None else None

        # Extract previous formatting for result (as single-item list for consistency)
        prev_formatting = ParagraphPropertyBuilder.extract(previous_ppr)

        # Create new pPr with merged formatting
        new_ppr = ParagraphPropertyBuilder.merge(existing_ppr, format_updates)

        # Check if there are actual changes
        if not ParagraphPropertyBuilder.has_changes(previous_ppr, new_ppr):
            return FormatResult(
                success=True,  # Operation completed without error
                changed=False,  # No changes needed
                text_matched=self._get_paragraph_text(target_para)[:50],
                paragraph_index=para_index,
                changes_applied={},
                previous_formatting=[prev_formatting],
                change_id=0,
                runs_affected=0,
            )

        # Create the tracked change element (returns tuple now)
        ppr_change, change_id = self._xml_generator.create_paragraph_property_change(
            previous_ppr, author
        )

        # Append the change tracking element to the new pPr
        new_ppr.append(ppr_change)

        # Replace or insert the pPr in the paragraph
        if existing_ppr is not None:
            target_para.remove(existing_ppr)
        target_para.insert(0, new_ppr)

        return FormatResult(
            success=True,
            changed=True,
            text_matched=self._get_paragraph_text(target_para)[:50],
            paragraph_index=para_index,
            changes_applied=format_updates,
            previous_formatting=[prev_formatting],
            change_id=change_id,
            runs_affected=1,
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
        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find source text
        source_matches = self._text_search.find_text(
            from_text,
            paragraphs,
            regex=False,
            normalize_quotes_for_matching=True,
        )

        if not source_matches:
            raise TextNotFoundError(from_text)

        # Get formatting from first match's first run
        source_match = source_matches[0]
        source_run = source_match.runs[source_match.start_run_index]
        source_r_pr = source_run.find(f"{{{WORD_NAMESPACE}}}rPr")

        if source_r_pr is None:
            # No formatting to copy
            return 0

        # Extract formatting properties
        bold = source_r_pr.find(f"{{{WORD_NAMESPACE}}}b") is not None
        italic = source_r_pr.find(f"{{{WORD_NAMESPACE}}}i") is not None
        color_elem = source_r_pr.find(f"{{{WORD_NAMESPACE}}}color")
        color = color_elem.get(f"{{{WORD_NAMESPACE}}}val") if color_elem is not None else None

        # Apply to target text
        return self.format_text(
            find=to_text,
            bold=bold,
            italic=italic,
            color=color,
            scope=scope,
            regex=False,
        )

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
        from python_docx_redline.models.paragraph import Paragraph

        # Validate arguments
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before'")
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before'")

        anchor_text = after if after is not None else before
        insert_after = after is not None

        # After validation, anchor_text is guaranteed to be a string
        assert anchor_text is not None

        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find the anchor paragraph
        matches = self._text_search.find_text(anchor_text, paragraphs)

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(anchor_text, paragraphs)
            raise TextNotFoundError(anchor_text, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(anchor_text, matches)

        match = matches[0]
        anchor_paragraph = match.paragraph

        # Create new paragraph element
        new_p = etree.Element(f"{{{WORD_NAMESPACE}}}p")

        # Add style if specified
        if style:
            p_pr = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}pPr")
            p_style = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}pStyle")
            p_style.set(f"{{{WORD_NAMESPACE}}}val", style)

        # If tracked, wrap the runs in w:ins
        if track:
            from datetime import datetime, timezone

            author_name = author if author is not None else self.author
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._xml_generator.next_change_id
            self._xml_generator.next_change_id += 1

            # Create w:ins element to wrap the run
            ins = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}ins")
            ins.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
            ins.set(f"{{{WORD_NAMESPACE}}}author", author_name)
            ins.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
            ins.set(
                "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc",
                timestamp,
            )

            # Add text content inside the w:ins element
            run = etree.SubElement(ins, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = text
        else:
            # Add text content directly to paragraph
            run = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}r")
            t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
            t.text = text

        element_to_insert = new_p

        # Insert the paragraph in the document
        parent = anchor_paragraph.getparent()
        if parent is None:
            raise ValueError("Anchor paragraph has no parent")

        anchor_index = list(parent).index(anchor_paragraph)

        if insert_after:
            # Insert after anchor
            parent.insert(anchor_index + 1, element_to_insert)
        else:
            # Insert before anchor
            parent.insert(anchor_index, element_to_insert)

        # Return Paragraph wrapper
        # new_p is always the actual paragraph element (whether tracked or not)
        return Paragraph(new_p)

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
        from python_docx_redline.models.paragraph import Paragraph as ParagraphClass

        if not texts:
            return []

        # Insert the first paragraph to find the anchor position
        first_para = self.insert_paragraph(
            texts[0],
            after=after,
            before=before,
            style=style,
            track=track,
            author=author,
            scope=scope,
        )

        created_paragraphs = [first_para]

        # Get the parent of the first paragraph
        parent = first_para.element.getparent()
        if parent is None:
            raise ValueError("First paragraph has no parent")
        insertion_index = list(parent).index(first_para.element)

        # Insert remaining paragraphs after the first one
        for i, text in enumerate(texts[1:], start=1):
            # Create new paragraph element
            new_p = etree.Element(f"{{{WORD_NAMESPACE}}}p")

            # Add style if specified
            if style:
                p_pr = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}pPr")
                p_style = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}pStyle")
                p_style.set(f"{{{WORD_NAMESPACE}}}val", style)

            # If tracked, wrap the runs in w:ins
            if track:
                from datetime import datetime, timezone

                author_name = author if author is not None else self.author
                timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                change_id = self._xml_generator.next_change_id
                self._xml_generator.next_change_id += 1

                # Create w:ins element to wrap the run
                ins = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}ins")
                ins.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                ins.set(f"{{{WORD_NAMESPACE}}}author", author_name)
                ins.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
                ins.set(
                    "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc",
                    timestamp,
                )

                # Add text content inside the w:ins element
                run = etree.SubElement(ins, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = text
            else:
                # Add text content directly to paragraph
                run = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = text

            # Insert after the previous paragraph
            parent.insert(insertion_index + i, new_p)

            created_paragraphs.append(ParagraphClass(new_p))

        return created_paragraphs

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
            update_toc: Automatically update Table of Contents (not implemented yet)
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
        from python_docx_redline.models.section import Section

        # Parse document into sections
        all_sections = Section.from_document(self.xml_root)

        # Apply scope filtering if specified
        if scope is not None:
            # Filter sections by checking if any paragraph in section is in scope
            all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
            paragraphs_in_scope = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)
            scope_para_set = set(paragraphs_in_scope)

            # Keep sections that have at least one paragraph in scope
            all_sections = [
                s for s in all_sections if any(p.element in scope_para_set for p in s.paragraphs)
            ]

        # Find matching sections (case insensitive by default for heading matching)
        matches = [
            s
            for s in all_sections
            if s.heading is not None and s.contains(heading, case_sensitive=False)
        ]

        if not matches:
            # Generate suggestions from section headings
            heading_paragraphs = [s.heading.element for s in all_sections if s.heading is not None]
            suggestions = SuggestionGenerator.generate_suggestions(heading, heading_paragraphs)
            raise TextNotFoundError(heading, suggestions=suggestions)

        if len(matches) > 1:
            # Create match representations for error reporting
            # Use the first paragraph of each matching section as the "match location"
            from python_docx_redline.text_search import TextSpan

            match_spans = []
            for section in matches:
                if section.heading:
                    # Create a TextSpan representing this section's heading
                    # Find the run elements in the heading paragraph
                    runs = list(section.heading.element.iter(f"{{{WORD_NAMESPACE}}}r"))
                    if runs:
                        heading_text = section.heading_text or ""
                        span = TextSpan(
                            runs=runs,
                            start_run_index=0,
                            end_run_index=len(runs) - 1,
                            start_offset=0,
                            end_offset=len(heading_text.strip()),
                            paragraph=section.heading.element,
                        )
                        match_spans.append(span)

            raise AmbiguousTextError(heading, match_spans)

        section = matches[0]

        # Delete all paragraphs in the section
        if track:
            from datetime import datetime, timezone

            author_name = author if author is not None else self.author
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

            # Wrap runs inside each paragraph with w:del
            for para in section.paragraphs:
                # Get all runs in the paragraph
                runs = list(para.element.iter(f"{{{WORD_NAMESPACE}}}r"))
                if not runs:
                    continue

                # Create w:del element
                change_id = self._xml_generator.next_change_id
                self._xml_generator.next_change_id += 1

                del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
                del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                del_elem.set(f"{{{WORD_NAMESPACE}}}author", author_name)
                del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
                del_elem.set(
                    "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc",
                    timestamp,
                )

                # Remove runs from paragraph and add them to w:del
                # Also change w:t to w:delText
                for run in runs:
                    run_parent = run.getparent()
                    if run_parent is not None:
                        run_parent.remove(run)

                    # Change all w:t elements to w:delText
                    for t_elem in run.iter(f"{{{WORD_NAMESPACE}}}t"):
                        deltext = etree.Element(f"{{{WORD_NAMESPACE}}}delText")
                        deltext.text = t_elem.text
                        # Copy xml:space attribute if present
                        xml_space = t_elem.get("{http://www.w3.org/XML/1998/namespace}space")
                        if xml_space:
                            deltext.set("{http://www.w3.org/XML/1998/namespace}space", xml_space)
                        # Replace w:t with w:delText in the run
                        t_parent = t_elem.getparent()
                        t_index = list(t_parent).index(t_elem)
                        t_parent.remove(t_elem)
                        t_parent.insert(t_index, deltext)

                    del_elem.append(run)

                # Insert w:del as first child of paragraph (after pPr if present)
                p_pr = para.element.find(f"{{{WORD_NAMESPACE}}}pPr")
                if p_pr is not None:
                    p_pr_index = list(para.element).index(p_pr)
                    para.element.insert(p_pr_index + 1, del_elem)
                else:
                    para.element.insert(0, del_elem)
        else:
            # Untracked deletion: simply remove paragraphs
            for para in section.paragraphs:
                parent = para.element.getparent()
                if parent is not None:
                    parent.remove(para.element)

        # TODO: Handle update_toc when implemented in separate task
        if update_toc:
            pass  # Will implement TOC updates in python_docx_redline-xpe

        return section

    def _insert_after_match(self, match: Any, insertion_element: Any) -> None:
        """Insert XML element(s) after a matched text span.

        Args:
            match: TextSpan object representing where to insert
            insertion_element: The lxml Element or list of Elements to insert
        """
        # Get the paragraph containing the match
        paragraph = match.paragraph

        # Find the run where the match ends
        end_run = match.runs[match.end_run_index]

        # Find the position of the end run in the paragraph
        run_index = list(paragraph).index(end_run)

        # Handle single element or list
        if isinstance(insertion_element, list):
            # Insert elements in order after the end run
            for i, elem in enumerate(insertion_element):
                paragraph.insert(run_index + 1 + i, elem)
        else:
            # Insert the new element after the end run
            paragraph.insert(run_index + 1, insertion_element)

    def _insert_before_match(self, match: Any, insertion_element: Any) -> None:
        """Insert XML element(s) before a matched text span.

        Args:
            match: TextSpan object representing where to insert
            insertion_element: The lxml Element or list of Elements to insert
        """
        # Get the paragraph containing the match
        paragraph = match.paragraph

        # Find the run where the match starts
        start_run = match.runs[match.start_run_index]

        # Find the position of the start run in the paragraph
        run_index = list(paragraph).index(start_run)

        # Handle single element or list
        if isinstance(insertion_element, list):
            # Insert elements in order before the start run
            for i, elem in enumerate(insertion_element):
                paragraph.insert(run_index + i, elem)
        else:
            # Insert the new element before the start run
            paragraph.insert(run_index, insertion_element)

    def _replace_match_with_element(self, match: Any, replacement_element: Any) -> None:
        """Replace matched text with a single XML element.

        This handles the complexity of text potentially spanning multiple runs.
        The matched runs are removed and replaced with the new element.

        Args:
            match: TextSpan object representing the text to replace
            replacement_element: The lxml Element to insert in place of matched text
        """
        paragraph = match.paragraph

        # If the match is within a single run
        if match.start_run_index == match.end_run_index:
            run = match.runs[match.start_run_index]
            run_text = "".join(run.itertext())

            # If the match is the entire run, replace the run
            if match.start_offset == 0 and match.end_offset == len(run_text):
                run_index = list(paragraph).index(run)
                paragraph.remove(run)
                paragraph.insert(run_index, replacement_element)
            else:
                # Match is partial - need to split the run
                self._split_and_replace_in_run(
                    paragraph, run, match.start_offset, match.end_offset, replacement_element
                )
        else:
            # Match spans multiple runs - remove all matched runs and insert replacement
            start_run = match.runs[match.start_run_index]
            start_run_index = list(paragraph).index(start_run)

            # Remove all runs in the match
            for i in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[i]
                if run in paragraph:
                    paragraph.remove(run)

            # Insert replacement at the position of the first removed run
            paragraph.insert(start_run_index, replacement_element)

    def _replace_match_with_elements(self, match: Any, replacement_elements: list[Any]) -> None:
        """Replace matched text with multiple XML elements.

        Used for replace_tracked which needs both deletion and insertion elements.

        Args:
            match: TextSpan object representing the text to replace
            replacement_elements: List of lxml Elements to insert in place of matched text
        """
        paragraph = match.paragraph

        # Similar to _replace_match_with_element but inserts multiple elements
        if match.start_run_index == match.end_run_index:
            run = match.runs[match.start_run_index]
            run_text = "".join(run.itertext())

            # If the match is the entire run, replace the run
            if match.start_offset == 0 and match.end_offset == len(run_text):
                run_index = list(paragraph).index(run)
                paragraph.remove(run)
                # Insert elements in order
                for i, elem in enumerate(replacement_elements):
                    paragraph.insert(run_index + i, elem)
            else:
                # Match is partial - need to split the run
                self._split_and_replace_in_run_multiple(
                    paragraph,
                    run,
                    match.start_offset,
                    match.end_offset,
                    replacement_elements,
                )
        else:
            # Match spans multiple runs
            start_run = match.runs[match.start_run_index]
            start_run_index = list(paragraph).index(start_run)

            # Remove all runs in the match
            for i in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[i]
                if run in paragraph:
                    paragraph.remove(run)

            # Insert all replacement elements at the position of the first removed run
            for i, elem in enumerate(replacement_elements):
                paragraph.insert(start_run_index + i, elem)

    def _split_and_replace_in_run(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_element: Any,
    ) -> None:
        """Split a run and replace a portion with a new element.

        Args:
            paragraph: The paragraph containing the run
            run: The run to split
            start_offset: Character offset where match starts
            end_offset: Character offset where match ends (exclusive)
            replacement_element: Element to insert in place of matched text
        """
        # Get the full text of the run
        text_elements = list(run.iter(f"{{{WORD_NAMESPACE}}}t"))
        if not text_elements:
            return

        # For simplicity, we'll work with the first text element
        # (Word typically has one w:t per run)
        text_elem = text_elements[0]
        run_text = text_elem.text or ""

        # Split into before, match, after
        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        run_index = list(paragraph).index(run)

        # Build new elements
        new_elements = []

        # Add before text if it exists
        if before_text:
            before_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            # Copy run properties if they exist
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                before_run.append(etree.fromstring(etree.tostring(run_props)))
            before_t = etree.SubElement(before_run, f"{{{WORD_NAMESPACE}}}t")
            if before_text and (before_text[0].isspace() or before_text[-1].isspace()):
                before_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            before_t.text = before_text
            new_elements.append(before_run)

        # Add replacement element
        new_elements.append(replacement_element)

        # Add after text if it exists
        if after_text:
            after_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            # Copy run properties if they exist
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                after_run.append(etree.fromstring(etree.tostring(run_props)))
            after_t = etree.SubElement(after_run, f"{{{WORD_NAMESPACE}}}t")
            if after_text and (after_text[0].isspace() or after_text[-1].isspace()):
                after_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            after_t.text = after_text
            new_elements.append(after_run)

        # Remove original run
        paragraph.remove(run)

        # Insert new elements
        for i, elem in enumerate(new_elements):
            paragraph.insert(run_index + i, elem)

    def _split_and_replace_in_run_multiple(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_elements: list[Any],
    ) -> None:
        """Split a run and replace a portion with multiple new elements.

        Args:
            paragraph: The paragraph containing the run
            run: The run to split
            start_offset: Character offset where match starts
            end_offset: Character offset where match ends (exclusive)
            replacement_elements: Elements to insert in place of matched text
        """
        # Get the full text of the run
        text_elements = list(run.iter(f"{{{WORD_NAMESPACE}}}t"))
        if not text_elements:
            return

        text_elem = text_elements[0]
        run_text = text_elem.text or ""

        # Split into before, match, after
        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        run_index = list(paragraph).index(run)

        # Build new elements
        new_elements = []

        # Add before text if it exists
        if before_text:
            before_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                before_run.append(etree.fromstring(etree.tostring(run_props)))
            before_t = etree.SubElement(before_run, f"{{{WORD_NAMESPACE}}}t")
            if before_text and (before_text[0].isspace() or before_text[-1].isspace()):
                before_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            before_t.text = before_text
            new_elements.append(before_run)

        # Add all replacement elements
        new_elements.extend(replacement_elements)

        # Add after text if it exists
        if after_text:
            after_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                after_run.append(etree.fromstring(etree.tostring(run_props)))
            after_t = etree.SubElement(after_run, f"{{{WORD_NAMESPACE}}}t")
            if after_text and (after_text[0].isspace() or after_text[-1].isspace()):
                after_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            after_t.text = after_text
            new_elements.append(after_run)

        # Remove original run
        paragraph.remove(run)

        # Insert new elements
        for i, elem in enumerate(new_elements):
            paragraph.insert(run_index + i, elem)

    def accept_all_changes(self) -> None:
        """Accept all tracked changes in the document.

        This removes all tracked change markup:
        - <w:ins> elements are unwrapped (content kept, wrapper removed)
        - <w:del> elements are completely removed (deleted content discarded)
        - <w:rPrChange> elements are removed (current formatting kept)
        - <w:pPrChange> elements are removed (current formatting kept)

        This is typically used as a preprocessing step before making new edits.
        """
        self.accept_insertions()
        self.accept_deletions()
        self.accept_format_changes()

    # Helper methods for change management

    def _get_paragraph_text(self, para: Any) -> str:
        """Extract text from a paragraph element.

        Args:
            para: A <w:p> XML element

        Returns:
            Plain text content of the paragraph
        """
        text_elements = para.findall(f".//{{{WORD_NAMESPACE}}}t")
        return "".join(elem.text or "" for elem in text_elements)

    def _unwrap_element(self, element: Any) -> None:
        """Unwrap an element by moving its children to its parent.

        Args:
            element: The element to unwrap
        """
        parent = element.getparent()
        if parent is None:
            return

        # Get the position of the element
        elem_index = list(parent).index(element)

        # Move all children to parent
        for child in list(element):
            parent.insert(elem_index, child)
            elem_index += 1

        # Remove the wrapper
        parent.remove(element)

    def _unwrap_deletion(self, del_elem: Any) -> None:
        """Unwrap a deletion element, converting w:delText back to w:t.

        When rejecting a deletion, we need to restore the deleted text by:
        1. Converting all <w:delText> elements to <w:t>
        2. Unwrapping the <w:del> element

        Args:
            del_elem: The <w:del> element to unwrap
        """
        # First, convert all w:delText to w:t within this deletion
        for deltext in del_elem.iter(f"{{{WORD_NAMESPACE}}}delText"):
            deltext.tag = f"{{{WORD_NAMESPACE}}}t"

        # Then unwrap the deletion element
        self._unwrap_element(del_elem)

    def _remove_element(self, element: Any) -> None:
        """Remove an element from its parent.

        Args:
            element: The element to remove
        """
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)

    # Accept/Reject by type

    def accept_insertions(self) -> int:
        """Accept all tracked insertions in the document.

        Unwraps all <w:ins> elements, keeping the inserted content.

        Returns:
            Number of insertions accepted
        """
        insertions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))
        for ins in insertions:
            self._unwrap_element(ins)
        return len(insertions)

    def reject_insertions(self) -> int:
        """Reject all tracked insertions in the document.

        Removes all <w:ins> elements and their content.

        Returns:
            Number of insertions rejected
        """
        insertions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))
        for ins in insertions:
            self._remove_element(ins)
        return len(insertions)

    def accept_deletions(self) -> int:
        """Accept all tracked deletions in the document.

        Removes all <w:del> elements (keeps text deleted).

        Returns:
            Number of deletions accepted
        """
        deletions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
        for del_elem in deletions:
            self._remove_element(del_elem)
        return len(deletions)

    def reject_deletions(self) -> int:
        """Reject all tracked deletions in the document.

        Unwraps all <w:del> elements, restoring the deleted content.
        Converts w:delText back to w:t.

        Returns:
            Number of deletions rejected
        """
        deletions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
        for del_elem in deletions:
            self._unwrap_deletion(del_elem)
        return len(deletions)

    def accept_format_changes(self) -> int:
        """Accept all tracked formatting changes in the document.

        Removes all <w:rPrChange> and <w:pPrChange> elements, keeping the
        current formatting as-is.

        Returns:
            Number of formatting changes accepted
        """
        count = 0

        # Accept run property changes (character formatting)
        for rpr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange")):
            parent = rpr_change.getparent()
            if parent is not None:
                parent.remove(rpr_change)
                count += 1

        # Accept paragraph property changes
        for ppr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange")):
            parent = ppr_change.getparent()
            if parent is not None:
                parent.remove(ppr_change)
                count += 1

        return count

    def reject_format_changes(self) -> int:
        """Reject all tracked formatting changes in the document.

        Restores the previous formatting from <w:rPrChange> and <w:pPrChange>
        elements, then removes the change tracking elements.

        Returns:
            Number of formatting changes rejected
        """
        from copy import deepcopy

        count = 0

        # Reject run property changes - restore previous formatting
        for rpr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange")):
            parent_rpr = rpr_change.getparent()
            if parent_rpr is None:
                continue

            # Get the previous rPr from inside the change element
            previous_rpr = rpr_change.find(f"{{{WORD_NAMESPACE}}}rPr")

            # Replace parent's children with previous state (except rPrChange)
            # First, collect children to remove (excluding the change element)
            children_to_remove = [
                child for child in list(parent_rpr) if child.tag != f"{{{WORD_NAMESPACE}}}rPrChange"
            ]
            for child in children_to_remove:
                parent_rpr.remove(child)

            # Remove the rPrChange element
            parent_rpr.remove(rpr_change)

            # Add back the previous properties
            if previous_rpr is not None:
                for child in previous_rpr:
                    parent_rpr.append(deepcopy(child))

            count += 1

        # Reject paragraph property changes - restore previous formatting
        for ppr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange")):
            parent_ppr = ppr_change.getparent()
            if parent_ppr is None:
                continue

            # Get the previous pPr from inside the change element
            previous_ppr = ppr_change.find(f"{{{WORD_NAMESPACE}}}pPr")

            # Replace parent's children with previous state (except pPrChange and rPr)
            # Note: rPr inside pPr is for paragraph mark formatting, keep it separate
            children_to_remove = [
                child
                for child in list(parent_ppr)
                if child.tag
                not in (
                    f"{{{WORD_NAMESPACE}}}pPrChange",
                    f"{{{WORD_NAMESPACE}}}rPr",
                )
            ]
            for child in children_to_remove:
                parent_ppr.remove(child)

            # Remove the pPrChange element
            parent_ppr.remove(ppr_change)

            # Add back the previous properties (insert at beginning, before rPr)
            if previous_ppr is not None:
                insert_idx = 0
                for child in previous_ppr:
                    parent_ppr.insert(insert_idx, deepcopy(child))
                    insert_idx += 1

            count += 1

        return count

    def reject_all_changes(self) -> None:
        """Reject all tracked changes in the document.

        This removes all tracked change markup by reverting to previous state:
        - <w:ins> elements and their content are removed
        - <w:del> elements are unwrapped (deleted content restored)
        - <w:rPrChange> elements restore previous formatting
        - <w:pPrChange> elements restore previous formatting
        """
        self.reject_insertions()
        self.reject_deletions()
        self.reject_format_changes()

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
        change_id_str = str(change_id)

        # Search for insertion with this ID
        for ins in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"):
            if ins.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                self._unwrap_element(ins)
                return

        # Search for deletion with this ID
        for del_elem in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"):
            if del_elem.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                self._remove_element(del_elem)
                return

        # Search for run property change with this ID
        for rpr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange"):
            if rpr_change.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                parent = rpr_change.getparent()
                if parent is not None:
                    parent.remove(rpr_change)
                return

        # Search for paragraph property change with this ID
        for ppr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange"):
            if ppr_change.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                parent = ppr_change.getparent()
                if parent is not None:
                    parent.remove(ppr_change)
                return

        raise ValueError(f"No tracked change found with ID: {change_id}")

    def reject_change(self, change_id: str | int) -> None:
        """Reject a specific tracked change by its ID.

        Args:
            change_id: The change ID (w:id attribute value)

        Raises:
            ValueError: If no change with the given ID is found

        Example:
            >>> doc.reject_change("5")
        """
        from copy import deepcopy

        change_id_str = str(change_id)

        # Search for insertion with this ID
        for ins in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"):
            if ins.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                self._remove_element(ins)
                return

        # Search for deletion with this ID
        for del_elem in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"):
            if del_elem.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                self._unwrap_deletion(del_elem)
                return

        # Search for run property change with this ID
        for rpr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange"):
            if rpr_change.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                parent_rpr = rpr_change.getparent()
                if parent_rpr is not None:
                    previous_rpr = rpr_change.find(f"{{{WORD_NAMESPACE}}}rPr")
                    # Remove current properties (except rPrChange)
                    for child in list(parent_rpr):
                        if child.tag != f"{{{WORD_NAMESPACE}}}rPrChange":
                            parent_rpr.remove(child)
                    parent_rpr.remove(rpr_change)
                    # Restore previous
                    if previous_rpr is not None:
                        for child in previous_rpr:
                            parent_rpr.append(deepcopy(child))
                return

        # Search for paragraph property change with this ID
        for ppr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange"):
            if ppr_change.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                parent_ppr = ppr_change.getparent()
                if parent_ppr is not None:
                    previous_ppr = ppr_change.find(f"{{{WORD_NAMESPACE}}}pPr")
                    # Remove current properties (except pPrChange and rPr)
                    for child in list(parent_ppr):
                        if child.tag not in (
                            f"{{{WORD_NAMESPACE}}}pPrChange",
                            f"{{{WORD_NAMESPACE}}}rPr",
                        ):
                            parent_ppr.remove(child)
                    parent_ppr.remove(ppr_change)
                    # Restore previous
                    if previous_ppr is not None:
                        insert_idx = 0
                        for child in previous_ppr:
                            parent_ppr.insert(insert_idx, deepcopy(child))
                            insert_idx += 1
                return

        raise ValueError(f"No tracked change found with ID: {change_id}")

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
        count = 0

        # Accept insertions by this author
        for ins in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins")):
            if ins.get(f"{{{WORD_NAMESPACE}}}author") == author:
                self._unwrap_element(ins)
                count += 1

        # Accept deletions by this author
        for del_elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del")):
            if del_elem.get(f"{{{WORD_NAMESPACE}}}author") == author:
                self._remove_element(del_elem)
                count += 1

        # Accept format changes by this author (remove rPrChange/pPrChange, keep new formatting)
        for rpr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange")):
            if rpr_change.get(f"{{{WORD_NAMESPACE}}}author") == author:
                parent = rpr_change.getparent()
                if parent is not None:
                    parent.remove(rpr_change)
                count += 1

        for ppr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange")):
            if ppr_change.get(f"{{{WORD_NAMESPACE}}}author") == author:
                parent = ppr_change.getparent()
                if parent is not None:
                    parent.remove(ppr_change)
                count += 1

        return count

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
        count = 0

        # Reject insertions by this author
        for ins in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins")):
            if ins.get(f"{{{WORD_NAMESPACE}}}author") == author:
                self._remove_element(ins)
                count += 1

        # Reject deletions by this author (convert delText back to t)
        for del_elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del")):
            if del_elem.get(f"{{{WORD_NAMESPACE}}}author") == author:
                self._unwrap_deletion(del_elem)
                count += 1

        # Reject format changes by this author (restore previous formatting)
        from copy import deepcopy

        for rpr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange")):
            if rpr_change.get(f"{{{WORD_NAMESPACE}}}author") == author:
                parent_rpr = rpr_change.getparent()
                if parent_rpr is not None:
                    # Get the previous rPr from inside the change element
                    prev_rpr = rpr_change.find(f"{{{WORD_NAMESPACE}}}rPr")
                    if prev_rpr is not None:
                        # Replace current rPr content with previous
                        # Clear all children except rPrChange
                        for child in list(parent_rpr):
                            if child.tag != f"{{{WORD_NAMESPACE}}}rPrChange":
                                parent_rpr.remove(child)
                        # Copy previous formatting back using deepcopy and append
                        # (preserves order and doesn't move nodes from prev_rpr)
                        for child in list(prev_rpr):
                            parent_rpr.append(deepcopy(child))
                    # Remove the change tracking element
                    parent_rpr.remove(rpr_change)
                count += 1

        for ppr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange")):
            if ppr_change.get(f"{{{WORD_NAMESPACE}}}author") == author:
                parent_ppr = ppr_change.getparent()
                if parent_ppr is not None:
                    # Get the previous pPr from inside the change element
                    prev_ppr = ppr_change.find(f"{{{WORD_NAMESPACE}}}pPr")
                    if prev_ppr is not None:
                        # Replace current pPr content with previous
                        # Clear all children except pPrChange and rPr (run props are separate)
                        for child in list(parent_ppr):
                            if child.tag not in (
                                f"{{{WORD_NAMESPACE}}}pPrChange",
                                f"{{{WORD_NAMESPACE}}}rPr",
                            ):
                                parent_ppr.remove(child)
                        # Copy previous formatting back using deepcopy and append
                        # (preserves order and doesn't move nodes from prev_ppr)
                        for child in list(prev_ppr):
                            parent_ppr.append(deepcopy(child))
                    # Remove the change tracking element
                    parent_ppr.remove(ppr_change)
                count += 1

        return count

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
        # Remove comment range markers
        for elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeStart")):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

        for elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeEnd")):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

        # Remove comment references
        # Comment references are typically in their own runs, so we'll remove the whole run
        for elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentReference")):
            parent = elem.getparent()
            if parent is not None:
                # If parent is a run, remove the run
                if parent.tag == f"{{{WORD_NAMESPACE}}}r":
                    grandparent = parent.getparent()
                    if grandparent is not None:
                        grandparent.remove(parent)
                else:
                    # Otherwise just remove the reference
                    parent.remove(elem)

        # Clean up comments-related files in the ZIP package
        if self._is_zip and self._temp_dir:
            # Delete comment XML files
            comment_files = [
                "word/comments.xml",
                "word/commentsExtended.xml",
                "word/commentsIds.xml",
                "word/commentsExtensible.xml",
            ]
            for file_path in comment_files:
                full_path = self._temp_dir / file_path
                if full_path.exists():
                    full_path.unlink()

            # Remove comment relationships from document.xml.rels
            rels_path = self._temp_dir / "word" / "_rels" / "document.xml.rels"
            if rels_path.exists():
                from lxml import etree as lxml_etree

                rels_tree = lxml_etree.parse(str(rels_path))
                rels_root = rels_tree.getroot()

                # Find and remove comment relationships
                comment_rel_types = [
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
                    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
                ]

                for rel in list(rels_root):
                    rel_type = rel.get("Type")
                    if rel_type in comment_rel_types:
                        rels_root.remove(rel)

                # Write back the modified relationships
                rels_tree.write(
                    str(rels_path),
                    encoding="utf-8",
                    xml_declaration=True,
                    pretty_print=False,
                )

            # Remove comment content types from [Content_Types].xml
            content_types_path = self._temp_dir / "[Content_Types].xml"
            if content_types_path.exists():
                from lxml import etree as lxml_etree

                ct_tree = lxml_etree.parse(str(content_types_path))
                ct_root = ct_tree.getroot()

                # Find and remove comment content type overrides
                ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
                comment_part_names = [
                    "/word/comments.xml",
                    "/word/commentsExtended.xml",
                    "/word/commentsIds.xml",
                    "/word/commentsExtensible.xml",
                ]

                for override in list(ct_root):
                    if override.tag == f"{{{ct_ns}}}Override":
                        part_name = override.get("PartName")
                        if part_name in comment_part_names:
                            ct_root.remove(override)

                # Write back the modified content types
                ct_tree.write(
                    str(content_types_path),
                    encoding="utf-8",
                    xml_declaration=True,
                    pretty_print=False,
                )

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
        from datetime import datetime, timezone

        from python_docx_redline.models.comment import Comment, CommentRange

        # Validate arguments
        if reply_to is None and on is None:
            raise ValueError("Either 'on' or 'reply_to' must be provided")

        # Determine author
        effective_author = author or self.author

        # Generate initials if not provided
        if initials is None:
            # Take first letter of each word in author name
            initials = "".join(word[0].upper() for word in effective_author.split() if word)
            if not initials:
                initials = effective_author[:2].upper() if effective_author else "AU"

        # Get next available comment ID
        comment_id = self._get_next_comment_id()

        # Get current timestamp
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        # Handle reply vs new comment
        if reply_to is not None:
            # This is a reply to an existing comment
            parent_comment = self._resolve_comment_reference(reply_to)
            parent_para_id = parent_comment.para_id

            if parent_para_id is None:
                raise ValueError("Cannot reply to comment without paraId")

            # Add comment to comments.xml (no document markers for replies)
            comment_elem = self._add_comment_to_comments_xml(
                comment_id, text, effective_author, initials, timestamp
            )

            # Get the paraId of the new comment
            new_para_id = comment_elem.find(f".//{{{WORD_NAMESPACE}}}p").get(
                "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"
            )

            # Create parent-child relationship in commentsExtended.xml
            self._link_comment_reply(new_para_id, parent_para_id)

            # Replies don't have a range
            return Comment(comment_elem, None, document=self)

        else:
            # This is a new top-level comment
            # Get all paragraphs in the document
            all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

            # Apply scope filter if specified
            paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

            # Search for the target text
            matches = self._text_search.find_text(on, paragraphs, regex=regex)

            if not matches:
                suggestions = SuggestionGenerator.generate_suggestions(on, paragraphs)
                raise TextNotFoundError(on, suggestions=suggestions)

            if len(matches) > 1:
                raise AmbiguousTextError(on, matches)

            match = matches[0]

            # Insert comment markers in document body
            self._insert_comment_markers(match, comment_id)

            # Add comment to comments.xml (create file if needed)
            comment_elem = self._add_comment_to_comments_xml(
                comment_id, text, effective_author, initials, timestamp
            )

            # Build the CommentRange for the return value
            from python_docx_redline.models.paragraph import Paragraph

            start_para = Paragraph(match.paragraph)
            comment_range = CommentRange(
                start_paragraph=start_para,
                end_paragraph=start_para,
                marked_text=match.text,
            )

            return Comment(comment_elem, comment_range, document=self)

    def _resolve_comment_reference(self, ref: "Comment | str | int") -> "Comment":
        """Resolve a comment reference to a Comment object.

        Args:
            ref: Comment object, comment ID string, or comment ID int

        Returns:
            The Comment object

        Raises:
            ValueError: If the comment is not found
        """
        from python_docx_redline.models.comment import Comment

        if isinstance(ref, Comment):
            return ref

        # Convert to string ID
        comment_id = str(ref)

        # Find the comment in the document
        for comment in self.comments:
            if comment.id == comment_id:
                return comment

        raise ValueError(f"Comment with ID '{comment_id}' not found")

    def _link_comment_reply(self, child_para_id: str, parent_para_id: str) -> None:
        """Link a reply comment to its parent in commentsExtended.xml.

        Creates or updates commentsExtended.xml to establish the parent-child
        relationship.

        Args:
            child_para_id: The paraId of the reply comment
            parent_para_id: The paraId of the parent comment
        """
        if not self._is_zip or not self._temp_dir:
            raise ValueError("Cannot link comments in non-ZIP documents")

        w15_namespace = "http://schemas.microsoft.com/office/word/2012/wordml"
        comments_ex_path = self._temp_dir / "word" / "commentsExtended.xml"

        # Load or create commentsExtended.xml
        if comments_ex_path.exists():
            tree = etree.parse(str(comments_ex_path))
            root = tree.getroot()
        else:
            # Create new commentsExtended.xml
            root = etree.Element(
                f"{{{w15_namespace}}}commentsEx",
                nsmap={"w15": w15_namespace},
            )
            tree = etree.ElementTree(root)

            # Add relationship and content type
            self._ensure_comments_extended_relationship()
            self._ensure_comments_extended_content_type()

        # Create commentEx element for the reply with paraIdParent
        comment_ex = etree.SubElement(root, f"{{{w15_namespace}}}commentEx")
        comment_ex.set(f"{{{w15_namespace}}}paraId", child_para_id)
        comment_ex.set(f"{{{w15_namespace}}}paraIdParent", parent_para_id)

        # Write back
        tree.write(
            str(comments_ex_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _get_next_comment_id(self) -> int:
        """Get the next available comment ID.

        Scans existing comments and returns max ID + 1.

        Returns:
            Next available comment ID (0 if no comments exist)
        """
        max_id = -1

        # Check comments.xml
        comments_xml = self._load_comments_xml()
        if comments_xml is not None:
            for comment in comments_xml.findall(f".//{{{WORD_NAMESPACE}}}comment"):
                try:
                    comment_id = int(comment.get(f"{{{WORD_NAMESPACE}}}id", "-1"))
                    max_id = max(max_id, comment_id)
                except ValueError:
                    pass

        # Also check document body for orphaned markers
        for marker in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeStart"):
            try:
                marker_id = int(marker.get(f"{{{WORD_NAMESPACE}}}id", "-1"))
                max_id = max(max_id, marker_id)
            except ValueError:
                pass

        return max_id + 1

    def _insert_comment_markers(self, match: Any, comment_id: int) -> None:
        """Insert comment range markers around matched text.

        Inserts commentRangeStart before the match, commentRangeEnd after,
        and commentReference in a new run after the end marker.

        Args:
            match: TextSpan object representing the text to annotate
            comment_id: The comment ID to use
        """
        paragraph = match.paragraph
        comment_id_str = str(comment_id)

        # Create the marker elements
        range_start = etree.Element(f"{{{WORD_NAMESPACE}}}commentRangeStart")
        range_start.set(f"{{{WORD_NAMESPACE}}}id", comment_id_str)

        range_end = etree.Element(f"{{{WORD_NAMESPACE}}}commentRangeEnd")
        range_end.set(f"{{{WORD_NAMESPACE}}}id", comment_id_str)

        # Create run containing comment reference
        ref_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
        ref = etree.SubElement(ref_run, f"{{{WORD_NAMESPACE}}}commentReference")
        ref.set(f"{{{WORD_NAMESPACE}}}id", comment_id_str)

        # Find positions in paragraph
        start_run = match.runs[match.start_run_index]
        end_run = match.runs[match.end_run_index]

        # Get indices
        children = list(paragraph)
        start_run_index = children.index(start_run)
        end_run_index = children.index(end_run)

        # Insert in reverse order to maintain correct indices
        # 1. Insert reference run after end run
        paragraph.insert(end_run_index + 1, ref_run)
        # 2. Insert range end after end run (before reference)
        paragraph.insert(end_run_index + 1, range_end)
        # 3. Insert range start before start run
        paragraph.insert(start_run_index, range_start)

    def _add_comment_to_comments_xml(
        self,
        comment_id: int,
        text: str,
        author: str,
        initials: str,
        timestamp: str,
    ) -> Any:
        """Add a comment to comments.xml, creating the file if needed.

        Args:
            comment_id: The comment ID
            text: Comment text content
            author: Author name
            initials: Author initials
            timestamp: ISO format timestamp

        Returns:
            The created w:comment Element
        """
        if not self._is_zip or not self._temp_dir:
            raise ValueError("Cannot add comments to non-ZIP documents")

        comments_path = self._temp_dir / "word" / "comments.xml"

        # Load or create comments.xml
        if comments_path.exists():
            comments_tree = etree.parse(str(comments_path))
            comments_root = comments_tree.getroot()
        else:
            # Create new comments.xml
            comments_root = etree.Element(
                f"{{{WORD_NAMESPACE}}}comments",
                nsmap={"w": WORD_NAMESPACE},
            )
            comments_tree = etree.ElementTree(comments_root)

            # Need to add relationship and content type
            self._ensure_comments_relationship()
            self._ensure_comments_content_type()

        # Create comment element
        comment_elem = etree.SubElement(comments_root, f"{{{WORD_NAMESPACE}}}comment")
        comment_elem.set(f"{{{WORD_NAMESPACE}}}id", str(comment_id))
        comment_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
        comment_elem.set(f"{{{WORD_NAMESPACE}}}initials", initials)
        comment_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

        # Add paragraph with text and w14:paraId for commentsExtended linking
        w14_namespace = "http://schemas.microsoft.com/office/word/2010/wordml"
        para_id = self._generate_para_id()

        para = etree.SubElement(comment_elem, f"{{{WORD_NAMESPACE}}}p")
        para.set(f"{{{w14_namespace}}}paraId", para_id)

        run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = text

        # Write comments.xml
        comments_tree.write(
            str(comments_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

        return comment_elem

    def _ensure_comments_relationship(self) -> None:
        """Ensure comments.xml relationship exists in document.xml.rels."""
        if not self._temp_dir:
            return

        rels_path = self._temp_dir / "word" / "_rels" / "document.xml.rels"
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        comment_rel_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        )

        if rels_path.exists():
            rels_tree = etree.parse(str(rels_path))
            rels_root = rels_tree.getroot()

            # Check if relationship already exists
            for rel in rels_root:
                if rel.get("Type") == comment_rel_type:
                    return  # Already exists
        else:
            # Create new rels file
            rels_path.parent.mkdir(parents=True, exist_ok=True)
            rels_root = etree.Element(
                f"{{{rels_ns}}}Relationships",
                nsmap={None: rels_ns},
            )
            rels_tree = etree.ElementTree(rels_root)

        # Find next available rId
        existing_ids = set()
        for rel in rels_root:
            rel_id = rel.get("Id", "")
            if rel_id.startswith("rId"):
                try:
                    existing_ids.add(int(rel_id[3:]))
                except ValueError:
                    pass

        next_id = 1
        while next_id in existing_ids:
            next_id += 1

        # Add relationship
        rel_elem = etree.SubElement(rels_root, f"{{{rels_ns}}}Relationship")
        rel_elem.set("Id", f"rId{next_id}")
        rel_elem.set("Type", comment_rel_type)
        rel_elem.set("Target", "comments.xml")

        rels_tree.write(
            str(rels_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _ensure_comments_content_type(self) -> None:
        """Ensure comments.xml content type exists in [Content_Types].xml."""
        if not self._temp_dir:
            return

        content_types_path = self._temp_dir / "[Content_Types].xml"
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        comments_part_name = "/word/comments.xml"
        comments_content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
        )

        if content_types_path.exists():
            ct_tree = etree.parse(str(content_types_path))
            ct_root = ct_tree.getroot()

            # Check if override already exists
            for override in ct_root:
                if override.tag == f"{{{ct_ns}}}Override":
                    if override.get("PartName") == comments_part_name:
                        return  # Already exists
        else:
            # This shouldn't happen for valid docx, but handle it
            ct_root = etree.Element(f"{{{ct_ns}}}Types", nsmap={None: ct_ns})
            ct_tree = etree.ElementTree(ct_root)

        # Add override
        override = etree.SubElement(ct_root, f"{{{ct_ns}}}Override")
        override.set("PartName", comments_part_name)
        override.set("ContentType", comments_content_type)

        ct_tree.write(
            str(content_types_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _generate_para_id(self) -> str:
        """Generate a unique paraId for comment paragraphs.

        paraId is an 8-character hex string (ST_LongHexNumber).

        Returns:
            8-character uppercase hex string
        """
        import random

        # Generate random 32-bit number and format as 8 hex chars
        return f"{random.randint(0, 0xFFFFFFFF):08X}"

    def _get_comment_ex(self, para_id: str) -> etree._Element | None:
        """Get the commentEx element for a given paraId.

        Args:
            para_id: The paraId to look up

        Returns:
            The w15:commentEx element or None if not found
        """
        if not self._is_zip or not self._temp_dir:
            return None

        comments_ex_path = self._temp_dir / "word" / "commentsExtended.xml"
        if not comments_ex_path.exists():
            return None

        w15_namespace = "http://schemas.microsoft.com/office/word/2012/wordml"

        tree = etree.parse(str(comments_ex_path))
        root = tree.getroot()

        for comment_ex in root.findall(f".//{{{w15_namespace}}}commentEx"):
            if comment_ex.get(f"{{{w15_namespace}}}paraId") == para_id:
                return comment_ex

        return None

    def _set_comment_resolved(self, para_id: str, resolved: bool) -> None:
        """Set the resolved status for a comment.

        Creates or updates commentsExtended.xml as needed.

        Args:
            para_id: The paraId of the comment
            resolved: True to mark as resolved, False for unresolved
        """
        if not self._is_zip or not self._temp_dir:
            raise ValueError("Cannot set resolution on non-ZIP documents")

        w15_namespace = "http://schemas.microsoft.com/office/word/2012/wordml"
        comments_ex_path = self._temp_dir / "word" / "commentsExtended.xml"

        # Load or create commentsExtended.xml
        if comments_ex_path.exists():
            tree = etree.parse(str(comments_ex_path))
            root = tree.getroot()
        else:
            # Create new commentsExtended.xml
            root = etree.Element(
                f"{{{w15_namespace}}}commentsEx",
                nsmap={"w15": w15_namespace},
            )
            tree = etree.ElementTree(root)

            # Add relationship and content type
            self._ensure_comments_extended_relationship()
            self._ensure_comments_extended_content_type()

        # Find or create commentEx element
        comment_ex = None
        for elem in root.findall(f".//{{{w15_namespace}}}commentEx"):
            if elem.get(f"{{{w15_namespace}}}paraId") == para_id:
                comment_ex = elem
                break

        if comment_ex is None:
            # Create new commentEx
            comment_ex = etree.SubElement(root, f"{{{w15_namespace}}}commentEx")
            comment_ex.set(f"{{{w15_namespace}}}paraId", para_id)

        # Set done status
        comment_ex.set(f"{{{w15_namespace}}}done", "1" if resolved else "0")

        # Write back
        tree.write(
            str(comments_ex_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _ensure_comments_extended_relationship(self) -> None:
        """Ensure commentsExtended.xml relationship exists."""
        if not self._temp_dir:
            return

        rels_path = self._temp_dir / "word" / "_rels" / "document.xml.rels"
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        rel_type = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"

        if rels_path.exists():
            rels_tree = etree.parse(str(rels_path))
            rels_root = rels_tree.getroot()

            # Check if relationship already exists
            for rel in rels_root:
                if rel.get("Type") == rel_type:
                    return
        else:
            rels_path.parent.mkdir(parents=True, exist_ok=True)
            rels_root = etree.Element(
                f"{{{rels_ns}}}Relationships",
                nsmap={None: rels_ns},
            )
            rels_tree = etree.ElementTree(rels_root)

        # Find next available rId
        existing_ids = set()
        for rel in rels_root:
            rel_id = rel.get("Id", "")
            if rel_id.startswith("rId"):
                try:
                    existing_ids.add(int(rel_id[3:]))
                except ValueError:
                    pass

        next_id = 1
        while next_id in existing_ids:
            next_id += 1

        # Add relationship
        rel_elem = etree.SubElement(rels_root, f"{{{rels_ns}}}Relationship")
        rel_elem.set("Id", f"rId{next_id}")
        rel_elem.set("Type", rel_type)
        rel_elem.set("Target", "commentsExtended.xml")

        rels_tree.write(
            str(rels_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _ensure_comments_extended_content_type(self) -> None:
        """Ensure commentsExtended.xml content type exists."""
        if not self._temp_dir:
            return

        content_types_path = self._temp_dir / "[Content_Types].xml"
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        part_name = "/word/commentsExtended.xml"
        content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
        )

        if content_types_path.exists():
            ct_tree = etree.parse(str(content_types_path))
            ct_root = ct_tree.getroot()

            # Check if override already exists
            for override in ct_root:
                if override.tag == f"{{{ct_ns}}}Override":
                    if override.get("PartName") == part_name:
                        return
        else:
            ct_root = etree.Element(f"{{{ct_ns}}}Types", nsmap={None: ct_ns})
            ct_tree = etree.ElementTree(ct_root)

        # Add override
        override = etree.SubElement(ct_root, f"{{{ct_ns}}}Override")
        override.set("PartName", part_name)
        override.set("ContentType", content_type)

        ct_tree.write(
            str(content_types_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _delete_comment(self, comment_id: str, para_id: str | None) -> None:
        """Delete a comment by ID.

        Removes the comment from comments.xml, the markers from the document
        body, and any commentsExtended.xml entry.

        Args:
            comment_id: The comment ID to delete
            para_id: The paraId of the comment (for commentsExtended cleanup)
        """
        # 1. Remove comment markers from document body
        self._remove_comment_markers(comment_id)

        # 2. Remove comment from comments.xml
        self._remove_from_comments_xml(comment_id)

        # 3. Remove from commentsExtended.xml if para_id is set
        if para_id:
            self._remove_from_comments_extended(para_id)

    def _remove_comment_markers(self, comment_id: str) -> None:
        """Remove comment range markers from document body.

        Args:
            comment_id: The comment ID to remove markers for
        """
        # Remove commentRangeStart
        for elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeStart")):
            if elem.get(f"{{{WORD_NAMESPACE}}}id") == comment_id:
                parent = elem.getparent()
                if parent is not None:
                    parent.remove(elem)

        # Remove commentRangeEnd
        for elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeEnd")):
            if elem.get(f"{{{WORD_NAMESPACE}}}id") == comment_id:
                parent = elem.getparent()
                if parent is not None:
                    parent.remove(elem)

        # Remove commentReference (and its parent run if empty)
        for elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentReference")):
            if elem.get(f"{{{WORD_NAMESPACE}}}id") == comment_id:
                parent = elem.getparent()
                if parent is not None:
                    # If parent is a run, check if it only contains this reference
                    if parent.tag == f"{{{WORD_NAMESPACE}}}r":
                        # Remove the reference first
                        parent.remove(elem)
                        # If run is now empty, remove the run too
                        if len(parent) == 0 or (
                            len(parent) == 1 and parent[0].tag == f"{{{WORD_NAMESPACE}}}rPr"
                        ):
                            grandparent = parent.getparent()
                            if grandparent is not None:
                                grandparent.remove(parent)
                    else:
                        parent.remove(elem)

    def _remove_from_comments_xml(self, comment_id: str) -> None:
        """Remove a comment from comments.xml.

        Args:
            comment_id: The comment ID to remove
        """
        if not self._is_zip or not self._temp_dir:
            return

        comments_path = self._temp_dir / "word" / "comments.xml"
        if not comments_path.exists():
            return

        tree = etree.parse(str(comments_path))
        root = tree.getroot()

        # Find and remove the comment element
        for comment in list(root.findall(f".//{{{WORD_NAMESPACE}}}comment")):
            if comment.get(f"{{{WORD_NAMESPACE}}}id") == comment_id:
                root.remove(comment)
                break

        # Write back
        tree.write(
            str(comments_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _remove_from_comments_extended(self, para_id: str) -> None:
        """Remove a commentEx entry from commentsExtended.xml.

        Args:
            para_id: The paraId of the comment to remove
        """
        if not self._is_zip or not self._temp_dir:
            return

        comments_ex_path = self._temp_dir / "word" / "commentsExtended.xml"
        if not comments_ex_path.exists():
            return

        w15_namespace = "http://schemas.microsoft.com/office/word/2012/wordml"

        tree = etree.parse(str(comments_ex_path))
        root = tree.getroot()

        # Find and remove the commentEx element
        for comment_ex in list(root.findall(f".//{{{w15_namespace}}}commentEx")):
            if comment_ex.get(f"{{{w15_namespace}}}paraId") == para_id:
                root.remove(comment_ex)
                break

        # Write back
        tree.write(
            str(comments_ex_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

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
        author_name = author if author is not None else self.author
        count = 0

        tables = self.tables
        if table_index is not None:
            if table_index < 0 or table_index >= len(tables):
                raise IndexError(f"Table index {table_index} out of range (0-{len(tables) - 1})")
            tables = [tables[table_index]]

        # Search and replace in each table
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    # Use TextSearch to find matches in cell paragraphs
                    for para in cell.paragraphs:
                        matches = self._text_search.find_text(
                            old_text,
                            [para.element],
                            regex=regex,
                            case_sensitive=case_sensitive,
                        )

                        for match in matches:
                            if track:
                                # Create tracked replacement
                                deletion_xml = self._xml_generator.create_deletion(
                                    match.text, author_name
                                )
                                insertion_xml = self._xml_generator.create_insertion(
                                    new_text, author_name
                                )

                                # Parse XMLs with namespace context
                                wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
    {insertion_xml}
</root>"""
                                root = etree.fromstring(wrapped_xml.encode("utf-8"))
                                deletion_element = root[0]
                                insertion_element = root[1]

                                self._replace_match_with_elements(
                                    match, [deletion_element, insertion_element]
                                )
                            else:
                                # Untracked replacement
                                new_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
                                new_t = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}t")
                                new_t.text = new_text
                                self._replace_match_with_element(match, new_run)

                            count += 1

        return count

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
        from python_docx_redline.models.table import TableRow

        tables = self.tables
        if table_index < 0 or table_index >= len(tables):
            raise IndexError(f"Table index {table_index} out of range (0-{len(tables) - 1})")

        table = tables[table_index]

        # Find the row to insert after
        if isinstance(after_row, int):
            if after_row < 0 or after_row >= table.row_count:
                raise IndexError(f"Row index {after_row} out of range (0-{table.row_count - 1})")
            insert_after_index = after_row
        else:
            # Find row containing text
            matching_rows = [
                (i, row) for i, row in enumerate(table.rows) if row.contains(after_row)
            ]

            if not matching_rows:
                raise ValueError(f"No row found containing text: {after_row}")
            if len(matching_rows) > 1:
                raise ValueError(
                    f"Text '{after_row}' found in {len(matching_rows)} rows - "
                    "please use a more specific search or row index"
                )

            insert_after_index = matching_rows[0][0]

        # Validate cell count
        if len(cells) != table.col_count:
            raise ValueError(f"Expected {table.col_count} cells, got {len(cells)}")

        # Create new row element
        new_row = etree.Element(f"{{{WORD_NAMESPACE}}}tr")

        # Create cells
        for cell_text in cells:
            tc = etree.SubElement(new_row, f"{{{WORD_NAMESPACE}}}tc")
            tc_pr = etree.SubElement(tc, f"{{{WORD_NAMESPACE}}}tcPr")
            tc_w = etree.SubElement(tc_pr, f"{{{WORD_NAMESPACE}}}tcW")
            tc_w.set(f"{{{WORD_NAMESPACE}}}w", "2880")
            tc_w.set(f"{{{WORD_NAMESPACE}}}type", "dxa")

            para = etree.SubElement(tc, f"{{{WORD_NAMESPACE}}}p")

            if cell_text:
                run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = cell_text

        if track:
            # Add insertion properties to row
            from datetime import datetime, timezone

            author_name = author if author is not None else self.author
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._xml_generator.next_change_id
            self._xml_generator.next_change_id += 1

            # Add w:trPr with w:ins child to mark row as inserted
            tr_pr = etree.Element(f"{{{WORD_NAMESPACE}}}trPr")
            ins_elem = etree.SubElement(tr_pr, f"{{{WORD_NAMESPACE}}}ins")
            ins_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
            ins_elem.set(f"{{{WORD_NAMESPACE}}}author", author_name)
            ins_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

            # Insert trPr as first child of row
            new_row.insert(0, tr_pr)

            # Insert after the specified row
            table_elem = table.element
            rows = table_elem.findall(f"{{{WORD_NAMESPACE}}}tr")
            target_row = rows[insert_after_index]
            row_index = list(table_elem).index(target_row)
            table_elem.insert(row_index + 1, new_row)

            # Return the row
            return TableRow(new_row, insert_after_index + 1)
        else:
            # Insert without tracking
            table_elem = table.element
            rows = table_elem.findall(f"{{{WORD_NAMESPACE}}}tr")
            target_row = rows[insert_after_index]
            row_index = list(table_elem).index(target_row)
            table_elem.insert(row_index + 1, new_row)

            return TableRow(new_row, insert_after_index + 1)

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

        tables = self.tables
        if table_index < 0 or table_index >= len(tables):
            raise IndexError(f"Table index {table_index} out of range (0-{len(tables) - 1})")

        table = tables[table_index]

        # Find the row to delete
        if isinstance(row, int):
            if row < 0 or row >= table.row_count:
                raise IndexError(f"Row index {row} out of range (0-{table.row_count - 1})")
            delete_index = row
        else:
            # Find row containing text
            matching_rows = [(i, r) for i, r in enumerate(table.rows) if r.contains(row)]

            if not matching_rows:
                raise ValueError(f"No row found containing text: {row}")
            if len(matching_rows) > 1:
                raise ValueError(
                    f"Text '{row}' found in {len(matching_rows)} rows - "
                    "please use a more specific search or row index"
                )

            delete_index = matching_rows[0][0]

        row_to_delete = table.rows[delete_index]

        if track:
            # Add deletion properties to row
            from datetime import datetime, timezone

            author_name = author if author is not None else self.author
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._xml_generator.next_change_id
            self._xml_generator.next_change_id += 1

            # Convert all w:t to w:delText within the row
            for t_elem in row_to_delete.element.findall(f".//{{{WORD_NAMESPACE}}}t"):
                t_elem.tag = f"{{{WORD_NAMESPACE}}}delText"

            # Add or update w:trPr with w:del child to mark row as deleted
            tr_pr = row_to_delete.element.find(f"{{{WORD_NAMESPACE}}}trPr")
            if tr_pr is None:
                tr_pr = etree.Element(f"{{{WORD_NAMESPACE}}}trPr")
                row_to_delete.element.insert(0, tr_pr)

            del_elem = etree.SubElement(tr_pr, f"{{{WORD_NAMESPACE}}}del")
            del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
            del_elem.set(f"{{{WORD_NAMESPACE}}}author", author_name)
            del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
        else:
            # Remove without tracking
            table_elem = table.element
            table_elem.remove(row_to_delete.element)

        return row_to_delete

    def _fix_encoding_declarations(self, directory: Path) -> None:
        """Fix encoding declarations in all XML files to use UTF-8.

        OOXML specification requires UTF-8 or UTF-16 encoding, but some tools
        (including Microsoft Word in certain cases) generate files with encoding="ASCII".
        This causes validation failures and documents that won't open in Word.

        This method fixes all XML files in the directory tree to use encoding="UTF-8".

        Args:
            directory: Root directory containing unpacked .docx files
        """
        import re

        # Find all XML and .rels files
        xml_files = list(directory.rglob("*.xml")) + list(directory.rglob("*.rels"))

        for xml_file in xml_files:
            try:
                # Read the file
                with open(xml_file, "rb") as f:
                    data = f.read()

                # Decode and check if it has an encoding declaration
                try:
                    text = data.decode("utf-8")
                except UnicodeDecodeError:
                    text = data.decode("latin-1")

                # Look for encoding declaration
                pattern = r'(<\?xml[^>]*encoding=)["\']([^"\']*)["\']'
                match = re.search(pattern, text[:200])

                if match:
                    encoding = match.group(2).upper()
                    # Only fix if it's not already UTF-8 or UTF-16
                    if encoding not in ["UTF-8", "UTF-16", "UTF-16LE", "UTF-16BE"]:
                        # Replace encoding declaration with UTF-8
                        new_text = re.sub(
                            pattern,
                            r'\1"UTF-8"',
                            text,
                            count=1,
                        )

                        # Write back as UTF-8
                        with open(xml_file, "wb") as f:
                            f.write(new_text.encode("utf-8"))

            except Exception:
                # Log but don't fail - validation will catch serious issues
                pass

    def _get_detailed_context(self, match: Any, context_chars: int = 50) -> tuple[str, str, str]:
        """Extract detailed context around a match for preview.

        Args:
            match: TextSpan object representing the matched text
            context_chars: Number of characters to extract before/after (default: 50)

        Returns:
            Tuple of (before_text, matched_text, after_text)
        """
        # Extract text from the paragraph
        text_elements = match.paragraph.findall(f".//{{{WORD_NAMESPACE}}}t")
        para_text = "".join(elem.text or "" for elem in text_elements)
        matched = match.text

        # Find the match position in the full paragraph text
        match_pos = para_text.find(matched)
        if match_pos == -1:
            # Fallback: couldn't find match in paragraph
            return ("", matched, "")

        # Extract context
        before_start = max(0, match_pos - context_chars)
        after_end = min(len(para_text), match_pos + len(matched) + context_chars)

        before_text = para_text[before_start:match_pos]
        after_text = para_text[match_pos + len(matched) : after_end]

        # Add ellipsis if truncated
        if before_start > 0:
            before_text = "..." + before_text
        if after_end < len(para_text):
            after_text = after_text + "..."

        return (before_text, matched, after_text)

    def _check_continuity(self, replacement: str, next_text: str) -> list[str]:
        """Check if replacement may create a sentence fragment.

        Analyzes the text immediately following the replacement to detect
        potential grammatical issues like sentence fragments or disconnected clauses.

        Args:
            replacement: The replacement text
            next_text: Text immediately following where replacement will be inserted

        Returns:
            List of warning messages (empty if no issues detected)
        """

        warnings = []

        # Skip check if no following text or it's just whitespace
        if not next_text or not next_text.strip():
            return warnings

        # Get the first ~30 chars of following text for analysis
        next_preview = next_text.strip()[:30]

        # Heuristic 1: Starts with lowercase letter (excluding special cases)
        # Skip 'i' for Roman numerals
        if next_preview and next_preview[0].islower() and next_preview[0] != "i":
            warnings.append("Next text starts with lowercase letter - may be a sentence fragment")

        # Heuristic 2: Starts with connecting phrase
        connecting_phrases = [
            "in question",
            "of which",
            "that is",
            "to which",
            "which is",
            "who is",
            "whose",
            "wherein",
            "whereby",
        ]

        next_lower = next_preview.lower()
        for phrase in connecting_phrases:
            if next_lower.startswith(phrase):
                warnings.append(
                    f"Next text starts with connecting phrase '{phrase}' - "
                    f"may require preceding context"
                )
                break

        # Heuristic 3: Starts with continuation punctuation
        if next_preview and next_preview[0] in [",", ";", ":", "—", "–"]:
            warnings.append("Next text starts with continuation punctuation - likely a fragment")

        return warnings

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
            if self._is_zip and self._temp_dir:
                # Write the modified XML back to the temp directory
                document_xml = self._temp_dir / "word" / "document.xml"
                self.xml_tree.write(
                    str(document_xml),
                    encoding="utf-8",
                    xml_declaration=True,
                    pretty_print=False,
                )

                # Fix encoding declarations in all XML files
                # OOXML spec requires UTF-8 or UTF-16, but some tools create files with ASCII
                self._fix_encoding_declarations(self._temp_dir)

                # Validate the full document structure before creating ZIP
                # This catches OOXML spec violations that would produce broken Word files
                if validate:
                    from .validation_docx import DOCXSchemaValidator

                    validator = DOCXSchemaValidator(
                        unpacked_dir=self._temp_dir,
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

                # Create a new .docx ZIP file
                with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zip_ref:
                    # Add all files from temp directory
                    for file in self._temp_dir.rglob("*"):
                        if file.is_file():
                            arcname = file.relative_to(self._temp_dir)
                            zip_ref.write(file, arcname)
            else:
                # Save XML directly
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
        if not self._is_zip or not self._temp_dir:
            raise ValidationError("save_to_bytes only supported for .docx files")

        try:
            # Write the modified XML back to the temp directory
            document_xml = self._temp_dir / "word" / "document.xml"
            self.xml_tree.write(
                str(document_xml),
                encoding="utf-8",
                xml_declaration=True,
                pretty_print=False,
            )

            # Fix encoding declarations in all XML files
            self._fix_encoding_declarations(self._temp_dir)

            # Validate if requested and we have an original file to compare against
            if validate and self.path is not None:
                from .validation_docx import DOCXSchemaValidator

                validator = DOCXSchemaValidator(
                    unpacked_dir=self._temp_dir,
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

            # Create ZIP in memory
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zip_ref:
                for file in self._temp_dir.rglob("*"):
                    if file.is_file():
                        arcname = file.relative_to(self._temp_dir)
                        zip_ref.write(file, arcname)

            buffer.seek(0)
            return buffer.read()

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
        results = []

        for i, edit in enumerate(edits):
            edit_type = edit.get("type")
            if not edit_type:
                results.append(
                    EditResult(
                        success=False,
                        edit_type="unknown",
                        message=f"Edit {i}: Missing 'type' field",
                        error=ValidationError("Missing 'type' field"),
                    )
                )
                if stop_on_error:
                    break
                continue

            try:
                result = self._apply_single_edit(edit_type, edit)
                results.append(result)

                if not result.success and stop_on_error:
                    break

            except Exception as e:
                results.append(
                    EditResult(
                        success=False,
                        edit_type=edit_type,
                        message=f"Error: {str(e)}",
                        error=e,
                    )
                )
                if stop_on_error:
                    break

        return results

    def _apply_single_edit(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Apply a single edit operation.

        Args:
            edit_type: The type of edit to perform
            edit: Dictionary with edit parameters

        Returns:
            EditResult indicating success or failure
        """
        try:
            if edit_type == "insert_tracked":
                text = edit.get("text")
                after = edit.get("after")
                author = edit.get("author")
                scope = edit.get("scope")
                regex = edit.get("regex", False)

                if not text or not after:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'text' or 'after'",
                        error=ValidationError("Missing required parameter"),
                    )

                self.insert_tracked(text, after, author=author, scope=scope, regex=regex)
                return EditResult(
                    success=True,
                    edit_type=edit_type,
                    message=f"Inserted '{text}' after '{after}'",
                )

            elif edit_type == "delete_tracked":
                text = edit.get("text")
                author = edit.get("author")
                scope = edit.get("scope")
                regex = edit.get("regex", False)

                if not text:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'text'",
                        error=ValidationError("Missing required parameter"),
                    )

                self.delete_tracked(text, author=author, scope=scope, regex=regex)
                return EditResult(success=True, edit_type=edit_type, message=f"Deleted '{text}'")

            elif edit_type == "replace_tracked":
                find = edit.get("find")
                replace = edit.get("replace")
                author = edit.get("author")
                scope = edit.get("scope")
                regex = edit.get("regex", False)

                if not find or replace is None:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'find' or 'replace'",
                        error=ValidationError("Missing required parameter"),
                    )

                self.replace_tracked(find, replace, author=author, scope=scope, regex=regex)
                return EditResult(
                    success=True,
                    edit_type=edit_type,
                    message=f"Replaced '{find}' with '{replace}'",
                )

            elif edit_type == "insert_paragraph":
                text = edit.get("text")
                after = edit.get("after")
                before = edit.get("before")
                style = edit.get("style")
                track = edit.get("track", True)
                author = edit.get("author")
                scope = edit.get("scope")

                if not text:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'text'",
                        error=ValidationError("Missing required parameter"),
                    )

                if not after and not before:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'after' or 'before'",
                        error=ValidationError("Missing required parameter"),
                    )

                self.insert_paragraph(
                    text,
                    after=after,
                    before=before,
                    style=style,
                    track=track,
                    author=author,
                    scope=scope,
                )
                location = f"after '{after}'" if after else f"before '{before}'"
                return EditResult(
                    success=True,
                    edit_type=edit_type,
                    message=f"Inserted paragraph '{text}' {location}",
                )

            elif edit_type == "insert_paragraphs":
                texts = edit.get("texts")
                after = edit.get("after")
                before = edit.get("before")
                style = edit.get("style")
                track = edit.get("track", True)
                author = edit.get("author")
                scope = edit.get("scope")

                if not texts:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'texts'",
                        error=ValidationError("Missing required parameter"),
                    )

                if not after and not before:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'after' or 'before'",
                        error=ValidationError("Missing required parameter"),
                    )

                self.insert_paragraphs(
                    texts,
                    after=after,
                    before=before,
                    style=style,
                    track=track,
                    author=author,
                    scope=scope,
                )
                location = f"after '{after}'" if after else f"before '{before}'"
                return EditResult(
                    success=True,
                    edit_type=edit_type,
                    message=f"Inserted {len(texts)} paragraphs {location}",
                )

            elif edit_type == "delete_section":
                heading = edit.get("heading")
                track = edit.get("track", True)
                update_toc = edit.get("update_toc", False)
                author = edit.get("author")
                scope = edit.get("scope")

                if not heading:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'heading'",
                        error=ValidationError("Missing required parameter"),
                    )

                self.delete_section(
                    heading, track=track, update_toc=update_toc, author=author, scope=scope
                )
                return EditResult(
                    success=True,
                    edit_type=edit_type,
                    message=f"Deleted section '{heading}'",
                )

            elif edit_type == "format_tracked":
                text = edit.get("text")
                if not text:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'text'",
                        error=ValidationError("Missing required parameter"),
                    )

                # Extract formatting parameters
                format_params = {
                    k: v
                    for k, v in edit.items()
                    if k
                    in (
                        "bold",
                        "italic",
                        "underline",
                        "strikethrough",
                        "font_name",
                        "font_size",
                        "color",
                        "highlight",
                        "superscript",
                        "subscript",
                        "small_caps",
                        "all_caps",
                    )
                    and v is not None
                }

                if not format_params:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="At least one formatting parameter required",
                        error=ValidationError("Missing formatting parameter"),
                    )

                result = self.format_tracked(
                    text,
                    scope=edit.get("scope"),
                    occurrence=edit.get("occurrence", "first"),
                    author=edit.get("author"),
                    **format_params,
                )
                return EditResult(
                    success=result.success,
                    edit_type=edit_type,
                    message=f"Formatted '{text}' with {format_params}",
                )

            elif edit_type == "format_paragraph_tracked":
                # Extract targeting parameters
                containing = edit.get("containing")
                starting_with = edit.get("starting_with")
                ending_with = edit.get("ending_with")
                index = edit.get("index")

                if not any([containing, starting_with, ending_with, index is not None]):
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="At least one targeting parameter required",
                        error=ValidationError("Missing targeting parameter"),
                    )

                # Extract formatting parameters
                format_params = {
                    k: v
                    for k, v in edit.items()
                    if k
                    in (
                        "alignment",
                        "spacing_before",
                        "spacing_after",
                        "line_spacing",
                        "indent_left",
                        "indent_right",
                        "indent_first_line",
                        "indent_hanging",
                    )
                    and v is not None
                }

                if not format_params:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="At least one formatting parameter required",
                        error=ValidationError("Missing formatting parameter"),
                    )

                result = self.format_paragraph_tracked(
                    containing=containing,
                    starting_with=starting_with,
                    ending_with=ending_with,
                    index=index,
                    scope=edit.get("scope"),
                    author=edit.get("author"),
                    **format_params,
                )
                target_desc = containing or starting_with or ending_with or f"index {index}"
                return EditResult(
                    success=result.success,
                    edit_type=edit_type,
                    message=f"Formatted paragraph '{target_desc}' with {format_params}",
                )

            else:
                return EditResult(
                    success=False,
                    edit_type=edit_type,
                    message=f"Unknown edit type: {edit_type}",
                    error=ValidationError(f"Unknown edit type: {edit_type}"),
                )

        except TextNotFoundError as e:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Text not found: {e}",
                error=e,
            )
        except AmbiguousTextError as e:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Ambiguous text: {e}",
                error=e,
            )
        except Exception as e:
            return EditResult(
                success=False, edit_type=edit_type, message=f"Error: {str(e)}", error=e
            )

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
        file_path = Path(path)

        if not file_path.exists():
            raise FileNotFoundError(f"Edit file not found: {path}")

        try:
            with open(file_path, encoding="utf-8") as f:
                if format == "yaml":
                    data = yaml.safe_load(f)
                elif format == "json":
                    import json

                    data = json.load(f)
                else:
                    raise ValidationError(f"Unsupported format: {format}")

            if not isinstance(data, dict):
                raise ValidationError("Edit file must contain a dictionary/object")

            if "edits" not in data:
                raise ValidationError("Edit file must contain an 'edits' key")

            edits = data["edits"]
            if not isinstance(edits, list):
                raise ValidationError("'edits' must be a list")

            # Apply the edits
            return self.apply_edits(edits, stop_on_error=stop_on_error)

        except yaml.YAMLError as e:
            raise ValidationError(f"Failed to parse YAML file: {e}") from e
        except Exception as e:
            if isinstance(e, ValidationError | FileNotFoundError):
                raise
            raise ValidationError(f"Failed to load edit file: {e}") from e

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
        # Get paragraph texts from both documents
        original_texts = [p.text for p in self.paragraphs]
        modified_texts = [p.text for p in modified.paragraphs]

        # Use SequenceMatcher to find differences at paragraph level
        matcher = difflib.SequenceMatcher(None, original_texts, modified_texts)
        opcodes = matcher.get_opcodes()

        # We need to process changes carefully to avoid index shifting issues
        # Build a list of operations to apply
        operations: list[dict[str, Any]] = []

        for tag, i1, i2, j1, j2 in opcodes:
            if tag == "equal":
                # No change needed
                continue
            elif tag == "delete":
                # Paragraphs removed in modified version
                for idx in range(i1, i2):
                    operations.append(
                        {
                            "type": "delete",
                            "original_index": idx,
                            "text": original_texts[idx],
                        }
                    )
            elif tag == "insert":
                # Paragraphs added in modified version
                # Insert after the previous paragraph (i1-1) or at beginning
                for j_idx in range(j1, j2):
                    operations.append(
                        {
                            "type": "insert",
                            "insert_after_index": i1 - 1,  # -1 means insert at beginning
                            "text": modified_texts[j_idx],
                            "modified_index": j_idx,
                        }
                    )
            elif tag == "replace":
                # Paragraphs changed
                # Check if this is a 1:1 replacement and minimal_edits is enabled
                is_one_to_one = (i2 - i1) == 1 and (j2 - j1) == 1

                if minimal_edits and is_one_to_one:
                    # Attempt minimal intra-paragraph edit for 1:1 replacement
                    operations.append(
                        {
                            "type": "minimal_replace",
                            "original_index": i1,
                            "original_text": original_texts[i1],
                            "new_text": modified_texts[j1],
                        }
                    )
                else:
                    # Fall back to coarse delete + insert
                    # First mark deletions
                    for idx in range(i1, i2):
                        operations.append(
                            {
                                "type": "delete",
                                "original_index": idx,
                                "text": original_texts[idx],
                            }
                        )
                    # Then mark insertions
                    for j_idx in range(j1, j2):
                        operations.append(
                            {
                                "type": "insert",
                                "insert_after_index": i1 - 1,
                                "text": modified_texts[j_idx],
                                "modified_index": j_idx,
                            }
                        )

        # Apply operations to the document
        change_count = self._apply_comparison_changes(operations, author, minimal_edits)

        return change_count

    def _apply_comparison_changes(
        self,
        operations: list[dict[str, Any]],
        author: str | None,
        minimal_edits: bool = False,
    ) -> int:
        """Apply comparison operations to generate tracked changes.

        Args:
            operations: List of delete/insert/minimal_replace operations from compare_to()
            author: Author for tracked changes
            minimal_edits: Whether minimal edits mode is enabled

        Returns:
            Number of changes applied
        """
        change_count = 0

        # Get all paragraph elements
        body = self.xml_root.find(f"{{{WORD_NAMESPACE}}}body")
        if body is None:
            return 0

        paragraphs = list(body.findall(f"{{{WORD_NAMESPACE}}}p"))

        # Track which paragraphs have been marked as deleted
        deleted_indices: set[int] = set()

        # Track which paragraphs have been handled by minimal_replace
        minimal_replace_indices: set[int] = set()

        # Process minimal replacements first
        for op in operations:
            if op["type"] == "minimal_replace":
                idx = op["original_index"]
                if idx < len(paragraphs) and idx not in minimal_replace_indices:
                    para_elem = paragraphs[idx]
                    orig_text = op["original_text"]
                    new_text = op["new_text"]

                    # Check if minimal editing is viable for this paragraph
                    use_minimal, diff_result, reason = should_use_minimal_editing(
                        para_elem, new_text, orig_text
                    )

                    if use_minimal and diff_result.hunks:
                        # Apply minimal edits
                        apply_minimal_edits_to_paragraph(
                            para_elem,
                            diff_result.hunks,
                            self._xml_generator,
                            author,
                        )
                        minimal_replace_indices.add(idx)
                        # Count changes consistently with coarse mode:
                        # Each hunk with delete_text counts as 1 deletion
                        # Each hunk with insert_text counts as 1 insertion
                        for hunk in diff_result.hunks:
                            if hunk.delete_text:
                                change_count += 1
                            if hunk.insert_text:
                                change_count += 1
                    elif not use_minimal:
                        # Fall back to coarse replacement
                        if reason:
                            logger.debug(
                                "Minimal editing disabled for paragraph %d: %s",
                                idx,
                                reason,
                            )
                        self._mark_paragraph_deleted(para_elem, author)
                        deleted_indices.add(idx)
                        change_count += 1

                        # Insert new paragraph after the deleted one
                        self._insert_comparison_paragraph(body, paragraphs, idx, new_text, author)
                        change_count += 1
                    # else: diff_result.hunks is empty (whitespace-only), no changes needed

        # Process deletions (mark content as deleted)
        for op in operations:
            if op["type"] == "delete":
                idx = op["original_index"]
                if (
                    idx < len(paragraphs)
                    and idx not in deleted_indices
                    and idx not in minimal_replace_indices
                ):
                    self._mark_paragraph_deleted(paragraphs[idx], author)
                    deleted_indices.add(idx)
                    change_count += 1

        # Process insertions
        # We need to track offset for insertions
        insertions_by_position: dict[int, list[str]] = {}
        for op in operations:
            if op["type"] == "insert":
                pos = op["insert_after_index"]
                if pos not in insertions_by_position:
                    insertions_by_position[pos] = []
                insertions_by_position[pos].append(op["text"])

        # Apply insertions in reverse order of position to avoid index shifting
        for pos in sorted(insertions_by_position.keys(), reverse=True):
            texts = insertions_by_position[pos]
            for text in reversed(texts):
                self._insert_comparison_paragraph(body, paragraphs, pos, text, author)
                change_count += 1

        return change_count

    def _mark_paragraph_deleted(
        self,
        paragraph: Any,
        author: str | None,
    ) -> None:
        """Mark all text in a paragraph as deleted with tracked changes.

        Args:
            paragraph: The paragraph XML element
            author: Author for the deletion
        """
        # Get all runs in the paragraph
        runs = paragraph.findall(f".//{{{WORD_NAMESPACE}}}r")

        for run in runs:
            # Get all text elements in this run
            text_elements = run.findall(f"{{{WORD_NAMESPACE}}}t")

            for t_elem in text_elements:
                text = t_elem.text or ""
                if not text:
                    continue

                # Create deletion XML
                del_xml = self._xml_generator.create_deletion(text, author)

                # Parse the deletion XML
                del_elem = etree.fromstring(
                    f'<root xmlns:w="{WORD_NAMESPACE}" '
                    f'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" '
                    f'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
                    f"{del_xml}</root>"
                )

                # Get the w:del element
                del_node = del_elem.find(f"{{{WORD_NAMESPACE}}}del")
                if del_node is not None:
                    # Insert the deletion before the original text element
                    parent = t_elem.getparent()
                    if parent is not None:
                        idx = list(parent).index(t_elem)
                        parent.insert(idx, del_node)
                        # Remove the original text element
                        parent.remove(t_elem)

    def _insert_comparison_paragraph(
        self,
        body: Any,
        paragraphs: list[Any],
        after_index: int,
        text: str,
        author: str | None,
    ) -> None:
        """Insert a new paragraph with tracked insertion.

        Args:
            body: The document body element
            paragraphs: List of existing paragraph elements
            after_index: Index of paragraph to insert after (-1 for beginning)
            text: Text content of the new paragraph
            author: Author for the insertion
        """
        # Create insertion XML
        ins_xml = self._xml_generator.create_insertion(text, author)

        # Create a new paragraph with the insertion
        new_para = etree.Element(f"{{{WORD_NAMESPACE}}}p")

        # Parse the insertion XML
        ins_elem = etree.fromstring(
            f'<root xmlns:w="{WORD_NAMESPACE}" '
            f'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" '
            f'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
            f"{ins_xml}</root>"
        )

        # Get the w:ins element
        ins_node = ins_elem.find(f"{{{WORD_NAMESPACE}}}ins")
        if ins_node is not None:
            new_para.append(ins_node)

        # Insert the new paragraph at the appropriate position
        if after_index < 0:
            # Insert at the beginning
            body.insert(0, new_para)
        elif after_index < len(paragraphs):
            # Insert after the specified paragraph
            ref_para = paragraphs[after_index]
            idx = list(body).index(ref_para)
            body.insert(idx + 1, new_para)
        else:
            # Insert at the end
            body.append(new_para)

    # ========================================================================
    # FOOTNOTE / ENDNOTE METHODS
    # ========================================================================

    @property
    def footnotes(self) -> list["Footnote"]:
        """Get all footnotes in the document.

        Returns:
            List of Footnote objects
        """
        from python_docx_redline.models.footnote import Footnote

        if not self._temp_dir:
            return []

        footnotes_path = self._temp_dir / "word" / "footnotes.xml"
        if not footnotes_path.exists():
            return []

        tree = etree.parse(str(footnotes_path))
        root = tree.getroot()

        # Find all footnote elements
        footnote_elems = root.findall(f"{{{WORD_NAMESPACE}}}footnote")

        # Filter out special footnotes (separator, continuationSeparator)
        # These have type attribute and IDs -1, 0, etc.
        return [
            Footnote(elem, self)
            for elem in footnote_elems
            if elem.get(f"{{{WORD_NAMESPACE}}}type") is None
        ]

    @property
    def endnotes(self) -> list["Endnote"]:
        """Get all endnotes in the document.

        Returns:
            List of Endnote objects
        """
        from python_docx_redline.models.footnote import Endnote

        if not self._temp_dir:
            return []

        endnotes_path = self._temp_dir / "word" / "endnotes.xml"
        if not endnotes_path.exists():
            return []

        tree = etree.parse(str(endnotes_path))
        root = tree.getroot()

        # Find all endnote elements
        endnote_elems = root.findall(f"{{{WORD_NAMESPACE}}}endnote")

        # Filter out special endnotes (separator, continuationSeparator)
        return [
            Endnote(elem, self)
            for elem in endnote_elems
            if elem.get(f"{{{WORD_NAMESPACE}}}type") is None
        ]

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
        if not self._is_zip or not self._temp_dir:
            raise ValueError("Cannot add footnotes to non-ZIP documents")

        author_name = author if author is not None else self.author

        # Find location for footnote reference
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._text_search.find_text(at, paragraphs)

        if not matches:
            raise TextNotFoundError(at, scope)

        if len(matches) > 1:
            raise AmbiguousTextError(at, matches)

        match = matches[0]

        # Generate new footnote ID
        footnote_id = self._get_next_footnote_id()

        # Add footnote content to footnotes.xml
        self._add_footnote_to_xml(footnote_id, text, author_name)

        # Insert footnote reference in document
        self._insert_footnote_reference(match, footnote_id)

        return footnote_id

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
        if not self._is_zip or not self._temp_dir:
            raise ValueError("Cannot add endnotes to non-ZIP documents")

        author_name = author if author is not None else self.author

        # Find location for endnote reference
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._text_search.find_text(at, paragraphs)

        if not matches:
            raise TextNotFoundError(at, scope)

        if len(matches) > 1:
            raise AmbiguousTextError(at, matches)

        match = matches[0]

        # Generate new endnote ID
        endnote_id = self._get_next_endnote_id()

        # Add endnote content to endnotes.xml
        self._add_endnote_to_xml(endnote_id, text, author_name)

        # Insert endnote reference in document
        self._insert_endnote_reference(match, endnote_id)

        return endnote_id

    def _get_next_footnote_id(self) -> int:
        """Get the next available footnote ID.

        Returns:
            Integer ID for new footnote
        """
        footnotes_path = self._temp_dir / "word" / "footnotes.xml"

        if not footnotes_path.exists():
            return 1

        tree = etree.parse(str(footnotes_path))
        root = tree.getroot()

        # Find all footnote IDs
        footnote_elems = root.findall(f"{{{WORD_NAMESPACE}}}footnote")
        ids = []

        for elem in footnote_elems:
            id_str = elem.get(f"{{{WORD_NAMESPACE}}}id")
            if id_str:
                try:
                    ids.append(int(id_str))
                except ValueError:
                    pass

        return max(ids) + 1 if ids else 1

    def _get_next_endnote_id(self) -> int:
        """Get the next available endnote ID.

        Returns:
            Integer ID for new endnote
        """
        endnotes_path = self._temp_dir / "word" / "endnotes.xml"

        if not endnotes_path.exists():
            return 1

        tree = etree.parse(str(endnotes_path))
        root = tree.getroot()

        # Find all endnote IDs
        endnote_elems = root.findall(f"{{{WORD_NAMESPACE}}}endnote")
        ids = []

        for elem in endnote_elems:
            id_str = elem.get(f"{{{WORD_NAMESPACE}}}id")
            if id_str:
                try:
                    ids.append(int(id_str))
                except ValueError:
                    pass

        return max(ids) + 1 if ids else 1

    def _add_footnote_to_xml(self, footnote_id: int, text: str, author: str) -> None:
        """Add a footnote to footnotes.xml, creating the file if needed.

        Args:
            footnote_id: The footnote ID
            text: Footnote text content
            author: Author name (for tracking if needed)
        """
        footnotes_path = self._temp_dir / "word" / "footnotes.xml"

        # Load or create footnotes.xml
        if footnotes_path.exists():
            footnotes_tree = etree.parse(str(footnotes_path))
            footnotes_root = footnotes_tree.getroot()
        else:
            # Create new footnotes.xml with separators
            footnotes_root = etree.Element(
                f"{{{WORD_NAMESPACE}}}footnotes",
                nsmap={"w": WORD_NAMESPACE},
            )
            footnotes_tree = etree.ElementTree(footnotes_root)

            # Add standard footnote separators (required by Word)
            # Separator (ID -1)
            sep = etree.SubElement(footnotes_root, f"{{{WORD_NAMESPACE}}}footnote")
            sep.set(f"{{{WORD_NAMESPACE}}}id", "-1")
            sep.set(f"{{{WORD_NAMESPACE}}}type", "separator")
            sep_p = etree.SubElement(sep, f"{{{WORD_NAMESPACE}}}p")
            sep_r = etree.SubElement(sep_p, f"{{{WORD_NAMESPACE}}}r")
            etree.SubElement(sep_r, f"{{{WORD_NAMESPACE}}}separator")

            # Continuation separator (ID 0)
            cont_sep = etree.SubElement(footnotes_root, f"{{{WORD_NAMESPACE}}}footnote")
            cont_sep.set(f"{{{WORD_NAMESPACE}}}id", "0")
            cont_sep.set(f"{{{WORD_NAMESPACE}}}type", "continuationSeparator")
            cont_sep_p = etree.SubElement(cont_sep, f"{{{WORD_NAMESPACE}}}p")
            cont_sep_r = etree.SubElement(cont_sep_p, f"{{{WORD_NAMESPACE}}}r")
            etree.SubElement(cont_sep_r, f"{{{WORD_NAMESPACE}}}continuationSeparator")

            # Need to add relationship and content type
            self._ensure_footnotes_relationship()
            self._ensure_footnotes_content_type()

        # Create footnote element
        footnote_elem = etree.SubElement(footnotes_root, f"{{{WORD_NAMESPACE}}}footnote")
        footnote_elem.set(f"{{{WORD_NAMESPACE}}}id", str(footnote_id))

        # Add paragraph with text
        para = etree.SubElement(footnote_elem, f"{{{WORD_NAMESPACE}}}p")
        run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = text

        # Write footnotes.xml
        footnotes_tree.write(
            str(footnotes_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _add_endnote_to_xml(self, endnote_id: int, text: str, author: str) -> None:
        """Add an endnote to endnotes.xml, creating the file if needed.

        Args:
            endnote_id: The endnote ID
            text: Endnote text content
            author: Author name (for tracking if needed)
        """
        endnotes_path = self._temp_dir / "word" / "endnotes.xml"

        # Load or create endnotes.xml
        if endnotes_path.exists():
            endnotes_tree = etree.parse(str(endnotes_path))
            endnotes_root = endnotes_tree.getroot()
        else:
            # Create new endnotes.xml with separators
            endnotes_root = etree.Element(
                f"{{{WORD_NAMESPACE}}}endnotes",
                nsmap={"w": WORD_NAMESPACE},
            )
            endnotes_tree = etree.ElementTree(endnotes_root)

            # Add standard endnote separators
            sep = etree.SubElement(endnotes_root, f"{{{WORD_NAMESPACE}}}endnote")
            sep.set(f"{{{WORD_NAMESPACE}}}id", "-1")
            sep.set(f"{{{WORD_NAMESPACE}}}type", "separator")
            sep_p = etree.SubElement(sep, f"{{{WORD_NAMESPACE}}}p")
            sep_r = etree.SubElement(sep_p, f"{{{WORD_NAMESPACE}}}r")
            etree.SubElement(sep_r, f"{{{WORD_NAMESPACE}}}separator")

            cont_sep = etree.SubElement(endnotes_root, f"{{{WORD_NAMESPACE}}}endnote")
            cont_sep.set(f"{{{WORD_NAMESPACE}}}id", "0")
            cont_sep.set(f"{{{WORD_NAMESPACE}}}type", "continuationSeparator")
            cont_sep_p = etree.SubElement(cont_sep, f"{{{WORD_NAMESPACE}}}p")
            cont_sep_r = etree.SubElement(cont_sep_p, f"{{{WORD_NAMESPACE}}}r")
            etree.SubElement(cont_sep_r, f"{{{WORD_NAMESPACE}}}continuationSeparator")

            # Need to add relationship and content type
            self._ensure_endnotes_relationship()
            self._ensure_endnotes_content_type()

        # Create endnote element
        endnote_elem = etree.SubElement(endnotes_root, f"{{{WORD_NAMESPACE}}}endnote")
        endnote_elem.set(f"{{{WORD_NAMESPACE}}}id", str(endnote_id))

        # Add paragraph with text
        para = etree.SubElement(endnote_elem, f"{{{WORD_NAMESPACE}}}p")
        run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = text

        # Write endnotes.xml
        endnotes_tree.write(
            str(endnotes_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _insert_footnote_reference(self, match: Any, footnote_id: int) -> None:
        """Insert a footnote reference at the matched text location.

        Args:
            match: TextSpan object indicating where to insert
            footnote_id: The footnote ID to reference
        """
        # Find the run where the match ends
        end_run = match.runs[-1] if match.runs else None
        if not end_run:
            return

        # Create a new run with the footnote reference
        new_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")

        # Add footnote reference
        footnote_ref = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}footnoteReference")
        footnote_ref.set(f"{{{WORD_NAMESPACE}}}id", str(footnote_id))

        # Insert the new run after the last run of the match
        parent = end_run.getparent()
        index = list(parent).index(end_run)
        parent.insert(index + 1, new_run)

    def _insert_endnote_reference(self, match: Any, endnote_id: int) -> None:
        """Insert an endnote reference at the matched text location.

        Args:
            match: TextSpan object indicating where to insert
            endnote_id: The endnote ID to reference
        """
        # Find the run where the match ends
        end_run = match.runs[-1] if match.runs else None
        if not end_run:
            return

        # Create a new run with the endnote reference
        new_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")

        # Add endnote reference
        endnote_ref = etree.SubElement(new_run, f"{{{WORD_NAMESPACE}}}endnoteReference")
        endnote_ref.set(f"{{{WORD_NAMESPACE}}}id", str(endnote_id))

        # Insert the new run after the last run of the match
        parent = end_run.getparent()
        index = list(parent).index(end_run)
        parent.insert(index + 1, new_run)

    def _ensure_footnotes_relationship(self) -> None:
        """Ensure footnotes.xml relationship exists in document.xml.rels."""
        if not self._temp_dir:
            return

        rels_path = self._temp_dir / "word" / "_rels" / "document.xml.rels"
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        footnote_rel_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
        )

        # Ensure _rels directory exists
        rels_dir = rels_path.parent
        rels_dir.mkdir(parents=True, exist_ok=True)

        # Load or create rels file
        if rels_path.exists():
            tree = etree.parse(str(rels_path))
            root = tree.getroot()
        else:
            root = etree.Element(
                f"{{{rels_ns}}}Relationships",
                nsmap={None: rels_ns},
            )
            tree = etree.ElementTree(root)

        # Check if footnotes relationship already exists
        for rel in root.findall(f"{{{rels_ns}}}Relationship"):
            if rel.get("Type") == footnote_rel_type:
                return  # Already exists

        # Find next relationship ID
        existing_ids = []
        for rel in root.findall(f"{{{rels_ns}}}Relationship"):
            rel_id = rel.get("Id", "")
            if rel_id.startswith("rId"):
                try:
                    existing_ids.append(int(rel_id[3:]))
                except ValueError:
                    pass

        next_id = max(existing_ids) + 1 if existing_ids else 1

        # Add footnotes relationship
        rel_elem = etree.SubElement(root, f"{{{rels_ns}}}Relationship")
        rel_elem.set("Id", f"rId{next_id}")
        rel_elem.set("Type", footnote_rel_type)
        rel_elem.set("Target", "footnotes.xml")

        # Write rels file
        tree.write(
            str(rels_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _ensure_endnotes_relationship(self) -> None:
        """Ensure endnotes.xml relationship exists in document.xml.rels."""
        if not self._temp_dir:
            return

        rels_path = self._temp_dir / "word" / "_rels" / "document.xml.rels"
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        endnote_rel_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
        )

        # Ensure _rels directory exists
        rels_dir = rels_path.parent
        rels_dir.mkdir(parents=True, exist_ok=True)

        # Load or create rels file
        if rels_path.exists():
            tree = etree.parse(str(rels_path))
            root = tree.getroot()
        else:
            root = etree.Element(
                f"{{{rels_ns}}}Relationships",
                nsmap={None: rels_ns},
            )
            tree = etree.ElementTree(root)

        # Check if endnotes relationship already exists
        for rel in root.findall(f"{{{rels_ns}}}Relationship"):
            if rel.get("Type") == endnote_rel_type:
                return  # Already exists

        # Find next relationship ID
        existing_ids = []
        for rel in root.findall(f"{{{rels_ns}}}Relationship"):
            rel_id = rel.get("Id", "")
            if rel_id.startswith("rId"):
                try:
                    existing_ids.append(int(rel_id[3:]))
                except ValueError:
                    pass

        next_id = max(existing_ids) + 1 if existing_ids else 1

        # Add endnotes relationship
        rel_elem = etree.SubElement(root, f"{{{rels_ns}}}Relationship")
        rel_elem.set("Id", f"rId{next_id}")
        rel_elem.set("Type", endnote_rel_type)
        rel_elem.set("Target", "endnotes.xml")

        # Write rels file
        tree.write(
            str(rels_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _ensure_footnotes_content_type(self) -> None:
        """Ensure footnotes.xml content type exists in [Content_Types].xml."""
        if not self._temp_dir:
            return

        content_types_path = self._temp_dir / "[Content_Types].xml"
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

        footnotes_part_name = "/word/footnotes.xml"
        footnotes_content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
        )

        tree = etree.parse(str(content_types_path))
        root = tree.getroot()

        # Check if already exists
        for override in root.findall(f"{{{ct_ns}}}Override"):
            if override.get("PartName") == footnotes_part_name:
                return  # Already exists

        # Add Override element
        override_elem = etree.SubElement(root, f"{{{ct_ns}}}Override")
        override_elem.set("PartName", footnotes_part_name)
        override_elem.set("ContentType", footnotes_content_type)

        # Write content types
        tree.write(
            str(content_types_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _ensure_endnotes_content_type(self) -> None:
        """Ensure endnotes.xml content type exists in [Content_Types].xml."""
        if not self._temp_dir:
            return

        content_types_path = self._temp_dir / "[Content_Types].xml"
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

        endnotes_part_name = "/word/endnotes.xml"
        endnotes_content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
        )

        tree = etree.parse(str(content_types_path))
        root = tree.getroot()

        # Check if already exists
        for override in root.findall(f"{{{ct_ns}}}Override"):
            if override.get("PartName") == endnotes_part_name:
                return  # Already exists

        # Add Override element
        override_elem = etree.SubElement(root, f"{{{ct_ns}}}Override")
        override_elem.set("PartName", endnotes_part_name)
        override_elem.set("ContentType", endnotes_content_type)

        # Write content types
        tree.write(
            str(content_types_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def __del__(self) -> None:
        """Clean up temporary directory on object destruction."""
        if self._temp_dir and self._temp_dir.exists() and self._is_zip:
            try:
                shutil.rmtree(self._temp_dir)
            except Exception:
                # Ignore cleanup errors
                pass

    def __enter__(self) -> "Document":
        """Context manager support."""
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        """Context manager cleanup."""
        self.__del__()
