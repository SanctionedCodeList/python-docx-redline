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
    from ..text_search import TextSpan

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
        """
        if reply_to is None and on is None:
            raise ValueError("Either 'on' or 'reply_to' must be provided")

        effective_author, effective_initials = self._prepare_author_info(author, initials)
        comment_id = self._get_next_comment_id()
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        if reply_to is not None:
            return self._add_reply_comment(
                text, reply_to, comment_id, effective_author, effective_initials, timestamp
            )
        else:
            # on cannot be None here due to the check above
            assert on is not None  # For type checker
            return self._add_top_level_comment(
                text,
                on,
                scope,
                regex,
                comment_id,
                effective_author,
                effective_initials,
                timestamp,
                occurrence,
            )

    def _prepare_author_info(self, author: str | None, initials: str | None) -> tuple[str, str]:
        """Prepare author and initials for a comment."""
        effective_author = author or self._document.author
        if initials is None:
            initials = "".join(word[0].upper() for word in effective_author.split() if word)
            if not initials:
                initials = effective_author[:2].upper() if effective_author else "AU"
        return effective_author, initials

    def _add_reply_comment(
        self,
        text: str,
        reply_to: Any,
        comment_id: int,
        author: str,
        initials: str,
        timestamp: str,
    ) -> Comment:
        """Add a reply to an existing comment."""
        from ..models.comment import Comment

        parent_comment = self._resolve_comment_reference(reply_to)
        parent_para_id = parent_comment.para_id

        if parent_para_id is None:
            raise ValueError("Cannot reply to comment without paraId")

        comment_elem = self._add_comment_to_comments_xml(
            comment_id, text, author, initials, timestamp
        )

        new_para_id = comment_elem.find(f".//{{{WORD_NAMESPACE}}}p").get(
            f"{{{W14_NAMESPACE}}}paraId"
        )
        self._link_comment_reply(new_para_id, parent_para_id)

        return Comment(comment_elem, None, document=self._document)

    def _add_top_level_comment(
        self,
        text: str,
        on: str,
        scope: str | dict | Any | None,
        regex: bool,
        comment_id: int,
        author: str,
        initials: str,
        timestamp: str,
        occurrence: int | list[int] | str | None = None,
    ) -> "Comment | list[Comment]":
        """Add a new top-level comment on text in the document."""
        from ..models.comment import Comment, CommentRange
        from ..models.paragraph import Paragraph

        all_paragraphs = list(self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        matches = self._document._text_search.find_text(on, paragraphs, regex=regex)
        if not matches:
            # Check if text exists anywhere (ignoring scope) for better error message
            hint = None
            if scope is not None:
                all_matches = self._document._text_search.find_text(on, all_paragraphs, regex=regex)
                if all_matches:
                    hint = (
                        f"Found {len(all_matches)} occurrence(s) in the document, "
                        f"but none within scope '{scope}'. "
                        "Try removing or adjusting the scope parameter."
                    )
            suggestions = SuggestionGenerator.generate_suggestions(on, paragraphs)
            raise TextNotFoundError(on, suggestions=suggestions, hint=hint)

        # Select matches based on occurrence parameter
        if occurrence is not None:
            selected_matches = self._select_matches(matches, occurrence, on)
        elif len(matches) > 1:
            raise AmbiguousTextError(on, matches)
        else:
            selected_matches = matches

        # Create comments for selected matches
        comments = []
        for i, match in enumerate(selected_matches):
            # Get a unique comment_id for each comment (first one uses the pre-allocated id)
            current_comment_id = comment_id if i == 0 else self._get_next_comment_id()

            self._insert_comment_markers(match, current_comment_id)

            comment_elem = self._add_comment_to_comments_xml(
                current_comment_id, text, author, initials, timestamp
            )

            start_para = Paragraph(match.paragraph)
            comment_range = CommentRange(
                start_paragraph=start_para,
                end_paragraph=start_para,
                marked_text=match.text,
            )
            comments.append(Comment(comment_elem, comment_range, document=self._document))

        # Return single comment or list based on what was requested
        if len(comments) == 1:
            return comments[0]
        return comments

    def _select_matches(
        self, matches: list["TextSpan"], occurrence: int | list[int] | str, text: str
    ) -> list["TextSpan"]:
        """Select target matches based on occurrence parameter.

        Args:
            matches: List of all matches found
            occurrence: Which occurrence(s) to select - int (1-indexed), list of ints, or string
            text: Original search text (for error messages)

        Returns:
            List of selected TextSpan matches

        Raises:
            AmbiguousTextError: If multiple matches and occurrence not specified
            ValueError: If occurrence is out of range
        """
        if occurrence == "first" or occurrence == 1:
            return [matches[0]]
        elif occurrence == "last":
            return [matches[-1]]
        elif occurrence == "all":
            return matches
        elif isinstance(occurrence, list):
            # Handle list of indices (1-indexed)
            selected = []
            for idx in occurrence:
                if not isinstance(idx, int):
                    raise ValueError(f"List elements must be integers, got {type(idx)}")
                if not (1 <= idx <= len(matches)):
                    raise ValueError(f"Occurrence {idx} out of range (1-{len(matches)})")
                selected.append(matches[idx - 1])
            return selected
        elif isinstance(occurrence, int) and 1 <= occurrence <= len(matches):
            return [matches[occurrence - 1]]
        elif isinstance(occurrence, int):
            raise ValueError(f"Occurrence {occurrence} out of range (1-{len(matches)})")
        elif len(matches) > 1:
            raise AmbiguousTextError(text, matches)
        else:
            return matches

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
        self._remove_all_comment_markers_from_xml()
        self._remove_comment_package_files()

    def _remove_all_comment_markers_from_xml(self) -> None:
        """Remove all comment markers from the document XML."""
        doc = self._document

        for elem in list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeStart")):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

        for elem in list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeEnd")):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

        for elem in list(doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentReference")):
            parent = elem.getparent()
            if parent is not None:
                if parent.tag == f"{{{WORD_NAMESPACE}}}r":
                    grandparent = parent.getparent()
                    if grandparent is not None:
                        grandparent.remove(parent)
                else:
                    parent.remove(elem)

    def _remove_comment_package_files(self) -> None:
        """Remove comment-related files from the ZIP package."""
        doc = self._document
        if not (doc._is_zip and doc._temp_dir):
            return

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

        if doc._package:
            self._remove_comment_relationships()
            self._remove_comment_content_types()

    def _remove_comment_relationships(self) -> None:
        """Remove comment relationships from document.xml.rels."""
        if self._document._package is None:
            return
        rel_mgr = RelationshipManager(self._document._package, "word/document.xml")
        comment_rel_types = [
            RelationshipTypes.COMMENTS,
            RelationshipTypes.COMMENTS_EXTENDED,
            RelationshipTypes.COMMENTS_IDS,
            RelationshipTypes.COMMENTS_EXTENSIBLE,
        ]
        rel_mgr.remove_relationships(comment_rel_types)
        rel_mgr.save()

    def _remove_comment_content_types(self) -> None:
        """Remove comment content types from [Content_Types].xml."""
        if self._document._package is None:
            return
        ct_mgr = ContentTypeManager(self._document._package)
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

    def _insert_comment_markers(self, match: TextSpan, comment_id: int) -> None:
        """Insert comment range markers around matched text.

        Inserts commentRangeStart before the match, commentRangeEnd after,
        and commentReference in a new run after the end marker.

        This handles runs that may be nested inside w:ins or w:del elements
        by finding the direct paragraph child containing each run.

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

        # Find positions in paragraph - handle nested runs (inside w:ins, w:del)
        start_run = match.runs[match.start_run_index]
        end_run = match.runs[match.end_run_index]

        # Find the direct paragraph children that contain our runs
        # (might be the run itself if not nested, or a w:ins/w:del wrapper)
        start_para_child = self._find_paragraph_child_containing(paragraph, start_run)
        end_para_child = self._find_paragraph_child_containing(paragraph, end_run)

        # Get indices of the direct paragraph children
        children = list(paragraph)
        start_child_index = children.index(start_para_child)
        end_child_index = children.index(end_para_child)

        # Insert in reverse order to maintain correct indices
        # 1. Insert reference run after end child
        paragraph.insert(end_child_index + 1, ref_run)
        # 2. Insert range end after end child (before reference)
        paragraph.insert(end_child_index + 1, range_end)
        # 3. Insert range start before start child
        paragraph.insert(start_child_index, range_start)

    def _find_paragraph_child_containing(
        self, paragraph: etree._Element, target: etree._Element
    ) -> etree._Element:
        """Find the direct child of paragraph that contains the target element.

        If target is a direct child of paragraph, returns target.
        Otherwise, walks up the tree from target until finding a direct child.

        This handles cases where runs are nested inside w:ins or w:del elements.

        Args:
            paragraph: The paragraph element
            target: The element to find (typically a run)

        Returns:
            The direct child of paragraph containing target
        """
        # Check if target is already a direct child
        if target.getparent() is paragraph:
            return target

        # Walk up from target until we find a direct child of paragraph
        current = target
        while current is not None:
            parent = current.getparent()
            if parent is paragraph:
                return current
            current = parent

        # Fallback: should not happen if target is actually in paragraph
        # but return target to avoid breaking
        return target

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
        if doc._temp_dir is None:
            raise ValueError("Cannot add comments to non-ZIP documents")

        comments_path = doc._temp_dir / "word" / "comments.xml"
        root, tree = self._load_or_create_comments_xml(comments_path)

        # OOXML spec requires paraId to be less than 0x80000000
        para_id = f"{random.randint(0, 0x7FFFFFFF):08X}"
        comment_elem = self._create_comment_element(
            root, comment_id, text, author, initials, timestamp, para_id
        )

        tree.write(str(comments_path), encoding="utf-8", xml_declaration=True, pretty_print=True)
        return comment_elem

    def _load_or_create_comments_xml(self, comments_path) -> tuple:
        """Load existing comments.xml or create a new one."""
        if comments_path.exists():
            tree = etree.parse(str(comments_path))
            return tree.getroot(), tree

        root = etree.Element(
            f"{{{WORD_NAMESPACE}}}comments",
            nsmap={"w": WORD_NAMESPACE, "w14": W14_NAMESPACE},
        )
        tree = etree.ElementTree(root)
        self._ensure_comments_relationship()
        self._ensure_comments_content_type()
        return root, tree

    def _create_comment_element(
        self,
        root,
        comment_id: int,
        text: str,
        author: str,
        initials: str,
        timestamp: str,
        para_id: str,
    ) -> etree._Element:
        """Create a comment element with its paragraph structure."""
        comment_elem = etree.SubElement(root, f"{{{WORD_NAMESPACE}}}comment")
        comment_elem.set(f"{{{WORD_NAMESPACE}}}id", str(comment_id))
        comment_elem.set(f"{{{WORD_NAMESPACE}}}author", author)
        comment_elem.set(f"{{{WORD_NAMESPACE}}}initials", initials)
        comment_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

        para = etree.SubElement(comment_elem, f"{{{WORD_NAMESPACE}}}p")
        para.set(f"{{{W14_NAMESPACE}}}paraId", para_id)

        run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
        text_elem = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        text_elem.text = text

        return comment_elem

    def _link_comment_reply(self, child_para_id: str, parent_para_id: str) -> None:
        """Create parent-child relationship in commentsExtended.xml."""
        doc = self._document

        if doc._temp_dir is None:
            raise ValueError("Cannot link comments in non-ZIP documents")
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
