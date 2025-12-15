"""
CommentOperations class for handling comment reading, adding, and deleting.

This module provides a dedicated class for all comment-related operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

import logging
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Any

from lxml import etree

from ..constants import WORD_NAMESPACE
from ..content_types import ContentTypeManager, ContentTypes
from ..errors import AmbiguousTextError, TextNotFoundError
from ..relationships import RelationshipManager, RelationshipTypes
from ..scope import ScopeEvaluator
from ..suggestions import SuggestionGenerator

if TYPE_CHECKING:
    from ..document import Document
    from ..models.comment import Comment

logger = logging.getLogger(__name__)

# Word 2010 namespace for paraId
W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml"


class CommentOperations:
    """Handles comment reading, adding, and deleting operations.

    This class encapsulates all comment-related functionality, including:
    - Reading existing comments
    - Adding new comments and replies
    - Deleting comments
    - Managing comment resolution status

    The class takes a Document reference and operates on its XML structure
    and package files.

    Example:
        >>> # Usually accessed through Document
        >>> doc = Document("contract.docx")
        >>> for comment in doc.comments:
        ...     print(f"{comment.author}: {comment.text}")
        >>> doc.add_comment("Please review", on="Section 2.1")
    """

    def __init__(self, document: Document) -> None:
        """Initialize CommentOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    @property
    def all(self) -> list[Comment]:
        """Get all comments in the document.

        Returns a list of Comment objects with both the comment content
        and the marked text range they apply to.

        Returns:
            List of Comment objects, empty list if no comments
        """
        from ..models.comment import Comment

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
            result.append(Comment(comment_elem, range_info, document=self._document))

        return result

    def get(
        self,
        *,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> list[Comment]:
        """Get comments with optional filtering.

        Args:
            author: Filter to comments by this author
            scope: Limit to comments within a specific scope
                   (section name, dict filter, or callable)

        Returns:
            Filtered list of Comment objects
        """
        all_comments = self.all

        if author:
            all_comments = [c for c in all_comments if c.author == author]

        if scope:
            evaluator = ScopeEvaluator.parse(scope)
            # Filter to comments whose marked text falls within scope
            all_comments = [
                c for c in all_comments if c.range and evaluator(c.range.start_paragraph.element)
            ]

        return all_comments

    def add(
        self,
        text: str,
        on: str | None = None,
        author: str | None = None,
        scope: str | dict | Any | None = None,
        regex: bool = False,
        initials: str | None = None,
        reply_to: Comment | str | int | None = None,
    ) -> Comment:
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
        """
        from ..models.comment import Comment, CommentRange
        from ..models.paragraph import Paragraph

        # Validate arguments
        if reply_to is None and on is None:
            raise ValueError("Either 'on' or 'reply_to' must be provided")

        # Determine author
        effective_author = author or self._document.author

        # Generate initials if not provided
        if initials is None:
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
                f"{{{W14_NAMESPACE}}}paraId"
            )

            # Create parent-child relationship in commentsExtended.xml
            self._link_comment_reply(new_para_id, parent_para_id)

            # Replies don't have a range
            return Comment(comment_elem, None, document=self._document)

        else:
            # This is a new top-level comment
            # Get all paragraphs in the document
            all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

            # Apply scope filter if specified
            paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

            # Search for the target text
            matches = self._document._text_search.find_text(on, paragraphs, regex=regex)

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
            start_para = Paragraph(match.paragraph)
            comment_range = CommentRange(
                start_paragraph=start_para,
                end_paragraph=start_para,
                marked_text=match.text,
            )

            return Comment(comment_elem, comment_range, document=self._document)

    def delete(self, comment: Comment | str | int) -> None:
        """Delete a specific comment.

        Args:
            comment: Comment object, comment ID string, or comment ID int
        """
        resolved = self._resolve_comment_reference(comment)
        self._delete_comment(resolved.id, resolved.para_id)

    def delete_all(self) -> None:
        """Delete all comments from the document.

        This removes all comment-related elements:
        - <w:commentRangeStart> - Comment range start markers
        - <w:commentRangeEnd> - Comment range end markers
        - <w:commentReference> - Comment reference markers
        - Runs containing comment references
        - word/comments.xml and related files
        - Comment relationships from document.xml.rels
        - Comment content types from [Content_Types].xml
        """
        doc = self._document

        # Remove comment range markers
        for elem in list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeStart")):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

        for elem in list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeEnd")):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

        # Remove comment references
        for elem in list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentReference")):
            parent = elem.getparent()
            if parent is not None:
                if parent.tag == f"{{{WORD_NAMESPACE}}}r":
                    grandparent = parent.getparent()
                    if grandparent is not None:
                        grandparent.remove(parent)
                else:
                    parent.remove(elem)

        # Clean up comments-related files in the ZIP package
        if doc._is_zip and doc._temp_dir:
            comment_files = [
                "word/comments.xml",
                "word/commentsExtended.xml",
                "word/commentsIds.xml",
                "word/commentsExtensible.xml",
            ]
            for file_path in comment_files:
                full_path = doc._temp_dir / file_path
                if full_path.exists():
                    full_path.unlink()

            # Remove comment relationships from document.xml.rels
            if doc._package:
                rel_mgr = RelationshipManager(doc._package, "word/document.xml")
                comment_rel_types = [
                    RelationshipTypes.COMMENTS,
                    RelationshipTypes.COMMENTS_EXTENDED,
                    RelationshipTypes.COMMENTS_IDS,
                    RelationshipTypes.COMMENTS_EXTENSIBLE,
                ]
                rel_mgr.remove_relationships(comment_rel_types)
                rel_mgr.save()

            # Remove comment content types from [Content_Types].xml
            if doc._package:
                ct_mgr = ContentTypeManager(doc._package)
                comment_part_names = [
                    "/word/comments.xml",
                    "/word/commentsExtended.xml",
                    "/word/commentsIds.xml",
                    "/word/commentsExtensible.xml",
                ]
                ct_mgr.remove_overrides(comment_part_names)
                ct_mgr.save()

    def resolve(self, comment: Comment | str | int) -> None:
        """Mark a comment as resolved.

        Args:
            comment: Comment object, comment ID string, or comment ID int
        """
        resolved = self._resolve_comment_reference(comment)
        if resolved.para_id:
            self._set_comment_resolved(resolved.para_id, True)

    def unresolve(self, comment: Comment | str | int) -> None:
        """Mark a comment as unresolved.

        Args:
            comment: Comment object, comment ID string, or comment ID int
        """
        resolved = self._resolve_comment_reference(comment)
        if resolved.para_id:
            self._set_comment_resolved(resolved.para_id, False)

    # ========================================================================
    # Private Helper Methods
    # ========================================================================

    def _load_comments_xml(self) -> etree._Element | None:
        """Load word/comments.xml if it exists.

        Returns:
            Root element of comments.xml or None if not present
        """
        doc = self._document
        if not doc._is_zip or not doc._temp_dir:
            return None

        comments_path = doc._temp_dir / "word" / "comments.xml"
        if not comments_path.exists():
            return None

        tree = etree.parse(str(comments_path))
        return tree.getroot()

    def _build_comment_ranges(self) -> dict[str, Any]:
        """Build a mapping of comment ID to marked text range.

        Returns:
            Dict mapping comment ID to CommentRange
        """
        from ..models.comment import CommentRange

        doc = self._document
        ranges: dict[str, CommentRange] = {}

        # Find all comment range starts
        for start_elem in doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeStart"):
            comment_id = start_elem.get(f"{{{WORD_NAMESPACE}}}id", "")
            if not comment_id:
                continue

            # Find matching end
            end_elem = None
            for elem in doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeEnd"):
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
        """Extract text between comment range markers."""
        if end_elem is None:
            return ""

        doc = self._document
        body = doc.xml_root.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return ""

        text_parts = []
        in_range = False

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

    def _find_containing_paragraph(self, elem: etree._Element) -> Any:
        """Find the paragraph containing an element."""
        from ..models.paragraph import Paragraph

        current = elem
        while current is not None:
            parent = current.getparent()
            if parent is None:
                break
            if parent.tag == f"{{{WORD_NAMESPACE}}}p":
                return Paragraph(parent)
            current = parent

        return None

    def _resolve_comment_reference(self, ref: Comment | str | int) -> Comment:
        """Resolve a comment reference to a Comment object."""
        from ..models.comment import Comment

        if isinstance(ref, Comment):
            return ref

        # Convert to string ID
        ref_id = str(ref)

        # Find the comment
        for comment in self.all:
            if comment.id == ref_id:
                return comment

        raise ValueError(f"Comment with ID '{ref_id}' not found")

    def _get_next_comment_id(self) -> int:
        """Get the next available comment ID."""
        comments_xml = self._load_comments_xml()
        if comments_xml is None:
            return 0

        max_id = -1
        for comment_elem in comments_xml.findall(f".//{{{WORD_NAMESPACE}}}comment"):
            try:
                comment_id = int(comment_elem.get(f"{{{WORD_NAMESPACE}}}id", "0"))
                max_id = max(max_id, comment_id)
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
    ) -> etree._Element:
        """Add a comment to comments.xml, creating the file if needed."""
        import random

        doc = self._document

        comments_path = doc._temp_dir / "word" / "comments.xml"

        if comments_path.exists():
            tree = etree.parse(str(comments_path))
            root = tree.getroot()
        else:
            # Create new comments.xml
            root = etree.Element(
                f"{{{WORD_NAMESPACE}}}comments",
                nsmap={
                    "w": WORD_NAMESPACE,
                    "w14": W14_NAMESPACE,
                },
            )
            tree = etree.ElementTree(root)

            # Ensure relationship and content type exist
            self._ensure_comments_relationship()
            self._ensure_comments_content_type()

        # Generate paraId
        para_id = f"{random.randint(0, 0xFFFFFFFF):08X}"

        # Create comment element
        comment_elem = etree.SubElement(root, f"{{{WORD_NAMESPACE}}}comment")
        comment_elem.set(f"{{{WORD_NAMESPACE}}}id", str(comment_id))
        comment_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
        comment_elem.set(f"{{{WORD_NAMESPACE}}}initials", initials)
        comment_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

        # Add paragraph with text
        para = etree.SubElement(comment_elem, f"{{{WORD_NAMESPACE}}}p")
        para.set(f"{{{W14_NAMESPACE}}}paraId", para_id)

        run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
        text_elem = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        text_elem.text = text

        # Write the file
        tree.write(
            str(comments_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

        return comment_elem

    def _link_comment_reply(self, child_para_id: str, parent_para_id: str) -> None:
        """Create parent-child relationship in commentsExtended.xml."""
        doc = self._document

        comments_ex_path = doc._temp_dir / "word" / "commentsExtended.xml"

        if comments_ex_path.exists():
            tree = etree.parse(str(comments_ex_path))
            root = tree.getroot()
        else:
            root = etree.Element(
                f"{{{W15_NAMESPACE}}}commentsEx",
                nsmap={"w15": W15_NAMESPACE},
            )
            tree = etree.ElementTree(root)

            self._ensure_comments_extended_relationship()
            self._ensure_comments_extended_content_type()

        # Add commentEx element for the reply
        comment_ex = etree.SubElement(root, f"{{{W15_NAMESPACE}}}commentEx")
        comment_ex.set(f"{{{W15_NAMESPACE}}}paraId", child_para_id)
        comment_ex.set(f"{{{W15_NAMESPACE}}}paraIdParent", parent_para_id)
        comment_ex.set(f"{{{W15_NAMESPACE}}}done", "0")

        tree.write(
            str(comments_ex_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _ensure_comments_relationship(self) -> None:
        """Ensure comments.xml relationship exists."""
        doc = self._document
        if not doc._package:
            return

        rel_mgr = RelationshipManager(doc._package, "word/document.xml")
        rel_mgr.add_relationship(RelationshipTypes.COMMENTS, "comments.xml")
        rel_mgr.save()

    def _ensure_comments_content_type(self) -> None:
        """Ensure comments.xml content type exists."""
        doc = self._document
        if not doc._package:
            return

        ct_mgr = ContentTypeManager(doc._package)
        ct_mgr.add_override("/word/comments.xml", ContentTypes.COMMENTS)
        ct_mgr.save()

    def _ensure_comments_extended_relationship(self) -> None:
        """Ensure commentsExtended.xml relationship exists."""
        doc = self._document
        if not doc._package:
            return

        rel_mgr = RelationshipManager(doc._package, "word/document.xml")
        rel_mgr.add_relationship(RelationshipTypes.COMMENTS_EXTENDED, "commentsExtended.xml")
        rel_mgr.save()

    def _ensure_comments_extended_content_type(self) -> None:
        """Ensure commentsExtended.xml content type exists."""
        doc = self._document
        if not doc._package:
            return

        ct_mgr = ContentTypeManager(doc._package)
        ct_mgr.add_override("/word/commentsExtended.xml", ContentTypes.COMMENTS_EXTENDED)
        ct_mgr.save()

    def _get_comment_ex(self, para_id: str) -> etree._Element | None:
        """Get the commentEx element for a given paraId."""
        doc = self._document
        if not doc._is_zip or not doc._temp_dir:
            return None

        comments_ex_path = doc._temp_dir / "word" / "commentsExtended.xml"
        if not comments_ex_path.exists():
            return None

        tree = etree.parse(str(comments_ex_path))
        root = tree.getroot()

        for comment_ex in root.findall(f".//{{{W15_NAMESPACE}}}commentEx"):
            if comment_ex.get(f"{{{W15_NAMESPACE}}}paraId") == para_id:
                return comment_ex

        return None

    def _set_comment_resolved(self, para_id: str, resolved: bool) -> None:
        """Set the resolved status for a comment."""
        doc = self._document
        if not doc._is_zip or not doc._temp_dir:
            raise ValueError("Cannot set resolution on non-ZIP documents")

        comments_ex_path = doc._temp_dir / "word" / "commentsExtended.xml"

        if comments_ex_path.exists():
            tree = etree.parse(str(comments_ex_path))
            root = tree.getroot()
        else:
            root = etree.Element(
                f"{{{W15_NAMESPACE}}}commentsEx",
                nsmap={"w15": W15_NAMESPACE},
            )
            tree = etree.ElementTree(root)

            self._ensure_comments_extended_relationship()
            self._ensure_comments_extended_content_type()

        # Find or create commentEx element
        comment_ex = None
        for elem in root.findall(f".//{{{W15_NAMESPACE}}}commentEx"):
            if elem.get(f"{{{W15_NAMESPACE}}}paraId") == para_id:
                comment_ex = elem
                break

        if comment_ex is None:
            comment_ex = etree.SubElement(root, f"{{{W15_NAMESPACE}}}commentEx")
            comment_ex.set(f"{{{W15_NAMESPACE}}}paraId", para_id)

        comment_ex.set(f"{{{W15_NAMESPACE}}}done", "1" if resolved else "0")

        tree.write(
            str(comments_ex_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _delete_comment(self, comment_id: str, para_id: str | None) -> None:
        """Delete a comment by ID."""
        # Remove comment markers from document body
        self._remove_comment_markers(comment_id)

        # Remove comment from comments.xml
        self._remove_from_comments_xml(comment_id)

        # Remove from commentsExtended.xml if para_id is set
        if para_id:
            self._remove_from_comments_extended(para_id)

    def _remove_comment_markers(self, comment_id: str) -> None:
        """Remove comment range markers from document body."""
        doc = self._document

        # Remove commentRangeStart
        for elem in list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeStart")):
            if elem.get(f"{{{WORD_NAMESPACE}}}id") == comment_id:
                parent = elem.getparent()
                if parent is not None:
                    parent.remove(elem)

        # Remove commentRangeEnd
        for elem in list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeEnd")):
            if elem.get(f"{{{WORD_NAMESPACE}}}id") == comment_id:
                parent = elem.getparent()
                if parent is not None:
                    parent.remove(elem)

        # Remove commentReference
        for elem in list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentReference")):
            if elem.get(f"{{{WORD_NAMESPACE}}}id") == comment_id:
                parent = elem.getparent()
                if parent is not None:
                    if parent.tag == f"{{{WORD_NAMESPACE}}}r":
                        grandparent = parent.getparent()
                        if grandparent is not None:
                            grandparent.remove(parent)
                    else:
                        parent.remove(elem)

    def _remove_from_comments_xml(self, comment_id: str) -> None:
        """Remove a comment from comments.xml."""
        doc = self._document
        if not doc._is_zip or not doc._temp_dir:
            return

        comments_path = doc._temp_dir / "word" / "comments.xml"
        if not comments_path.exists():
            return

        tree = etree.parse(str(comments_path))
        root = tree.getroot()

        for comment_elem in list(root.findall(f".//{{{WORD_NAMESPACE}}}comment")):
            if comment_elem.get(f"{{{WORD_NAMESPACE}}}id") == comment_id:
                root.remove(comment_elem)
                break

        tree.write(
            str(comments_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )

    def _remove_from_comments_extended(self, para_id: str) -> None:
        """Remove a comment from commentsExtended.xml."""
        doc = self._document
        if not doc._is_zip or not doc._temp_dir:
            return

        comments_ex_path = doc._temp_dir / "word" / "commentsExtended.xml"
        if not comments_ex_path.exists():
            return

        tree = etree.parse(str(comments_ex_path))
        root = tree.getroot()

        for comment_ex in list(root.findall(f".//{{{W15_NAMESPACE}}}commentEx")):
            if comment_ex.get(f"{{{W15_NAMESPACE}}}paraId") == para_id:
                root.remove(comment_ex)
                break

        tree.write(
            str(comments_ex_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=True,
        )
