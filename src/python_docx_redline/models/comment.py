"""
Comment wrapper class for document comments.

Provides a Pythonic API for accessing Word document comments,
including the comment content, author, date, and the text that
the comment applies to.
"""

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from python_docx_redline.document import Document
    from python_docx_redline.models.paragraph import Paragraph

# Word namespace
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
# Word 2010 namespace (for paraId)
W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"
# Word 2012 namespace (for commentsExtended)
W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass
class CommentRange:
    """Represents the text range a comment applies to.

    Attributes:
        start_paragraph: Paragraph containing the range start
        end_paragraph: Paragraph containing the range end (may be same as start)
        marked_text: The text that the comment applies to
    """

    start_paragraph: "Paragraph"
    end_paragraph: "Paragraph"
    marked_text: str


class Comment:
    """Wrapper around a w:comment element.

    Provides convenient Python API for accessing comment data including
    the comment text, author, date, and the document text it references.

    Example:
        >>> for comment in doc.comments:
        ...     print(f"{comment.author}: {comment.text}")
        ...     if comment.marked_text:
        ...         print(f"  On: '{comment.marked_text}'")
        ...     if comment.is_resolved:
        ...         print("  [RESOLVED]")
    """

    def __init__(
        self,
        element: etree._Element,
        range_info: CommentRange | None = None,
        document: "Document | None" = None,
    ):
        """Initialize Comment wrapper.

        Args:
            element: The w:comment XML element
            range_info: Optional range information from document body
            document: Reference to parent Document (needed for resolve/unresolve)
        """
        self._element = element
        self._range = range_info
        self._document = document

    @property
    def element(self) -> etree._Element:
        """Get the underlying XML element."""
        return self._element

    @property
    def id(self) -> str:
        """Get the comment ID.

        Returns:
            The comment ID as a string
        """
        return self._element.get(f"{{{WORD_NAMESPACE}}}id", "")

    @property
    def author(self) -> str:
        """Get the comment author.

        Returns:
            Author name, or empty string if not set
        """
        return self._element.get(f"{{{WORD_NAMESPACE}}}author", "")

    @property
    def initials(self) -> str | None:
        """Get the author's initials.

        Returns:
            Initials string or None if not present
        """
        return self._element.get(f"{{{WORD_NAMESPACE}}}initials")

    @property
    def date(self) -> datetime | None:
        """Get the comment date/time.

        Returns:
            datetime object or None if not present/parseable
        """
        date_str = self._element.get(f"{{{WORD_NAMESPACE}}}date")
        if not date_str:
            return None
        try:
            # OOXML uses ISO 8601 format, handle both Z and offset formats
            return datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        except ValueError:
            return None

    @property
    def text(self) -> str:
        """Get the comment text content.

        Extracts text from all w:t elements within the comment.

        Returns:
            The full text of the comment
        """
        text_elements = self._element.findall(f".//{{{WORD_NAMESPACE}}}t")
        return "".join(elem.text or "" for elem in text_elements)

    @property
    def marked_text(self) -> str | None:
        """Get the text that this comment applies to.

        This is the document text between commentRangeStart and
        commentRangeEnd markers.

        Returns:
            The marked text, or None if range info unavailable
        """
        if self._range:
            return self._range.marked_text
        return None

    @property
    def range(self) -> CommentRange | None:
        """Get the full range information.

        Returns:
            CommentRange with paragraph references, or None if unavailable
        """
        return self._range

    @property
    def para_id(self) -> str | None:
        """Get the paraId of the last paragraph in this comment.

        The paraId is used to link to commentsExtended.xml for
        resolution status and reply threading.

        Returns:
            The paraId as a hex string, or None if not present
        """
        # Find all paragraphs in the comment
        paragraphs = self._element.findall(f".//{{{WORD_NAMESPACE}}}p")
        if not paragraphs:
            return None

        # Get the last paragraph's w14:paraId
        last_para = paragraphs[-1]
        return last_para.get(f"{{{W14_NAMESPACE}}}paraId")

    @property
    def is_resolved(self) -> bool:
        """Check if this comment is marked as resolved/done.

        Reads the done attribute from commentsExtended.xml.

        Returns:
            True if resolved, False otherwise
        """
        if self._document is None:
            return False

        para_id = self.para_id
        if not para_id:
            return False

        # Load commentsExtended.xml and check done status
        comment_ex = self._document._get_comment_ex(para_id)
        if comment_ex is None:
            return False

        done = comment_ex.get(f"{{{W15_NAMESPACE}}}done")
        # done="1" or done="true" means resolved
        return done in ("1", "true")

    def resolve(self) -> None:
        """Mark this comment as resolved.

        Updates commentsExtended.xml to set done="1".

        Raises:
            ValueError: If document reference is not available
        """
        if self._document is None:
            raise ValueError("Cannot resolve comment: no document reference")

        para_id = self.para_id
        if not para_id:
            raise ValueError("Cannot resolve comment: no paraId found")

        self._document._set_comment_resolved(para_id, resolved=True)

    def unresolve(self) -> None:
        """Mark this comment as unresolved.

        Updates commentsExtended.xml to set done="0".

        Raises:
            ValueError: If document reference is not available
        """
        if self._document is None:
            raise ValueError("Cannot unresolve comment: no document reference")

        para_id = self.para_id
        if not para_id:
            raise ValueError("Cannot unresolve comment: no paraId found")

        self._document._set_comment_resolved(para_id, resolved=False)

    def delete(self) -> None:
        """Delete this comment from the document.

        Removes:
        - The comment element from comments.xml
        - The commentRangeStart, commentRangeEnd, and commentReference
          markers from the document body
        - Any commentsExtended.xml entry for this comment

        Raises:
            ValueError: If document reference is not available
        """
        if self._document is None:
            raise ValueError("Cannot delete comment: no document reference")

        self._document._delete_comment(self.id, self.para_id)

    @property
    def parent(self) -> "Comment | None":
        """Get the parent comment if this is a reply.

        Reads the paraIdParent from commentsExtended.xml to find the
        parent comment.

        Returns:
            The parent Comment object, or None if this is a top-level comment
        """
        if self._document is None:
            return None

        para_id = self.para_id
        if not para_id:
            return None

        # Load commentsExtended.xml and check for parent
        comment_ex = self._document._get_comment_ex(para_id)
        if comment_ex is None:
            return None

        parent_para_id = comment_ex.get(f"{{{W15_NAMESPACE}}}paraIdParent")
        if not parent_para_id:
            return None

        # Find the comment with this paraId
        for comment in self._document.comments:
            if comment.para_id == parent_para_id:
                return comment

        return None

    @property
    def replies(self) -> list["Comment"]:
        """Get all replies to this comment.

        Reads commentsExtended.xml to find comments with this comment's
        paraId as their paraIdParent.

        Returns:
            List of Comment objects that are replies to this comment
        """
        if self._document is None:
            return []

        para_id = self.para_id
        if not para_id:
            return []

        # Find all comments that have this comment as parent
        reply_list = []
        for comment in self._document.comments:
            if comment.para_id == para_id:
                continue  # Skip self

            # Check if this comment's parent is our paraId
            comment_ex = self._document._get_comment_ex(comment.para_id)
            if comment_ex is not None:
                parent_para_id = comment_ex.get(f"{{{W15_NAMESPACE}}}paraIdParent")
                if parent_para_id == para_id:
                    reply_list.append(comment)

        return reply_list

    def add_reply(
        self,
        text: str,
        author: str | None = None,
        initials: str | None = None,
    ) -> "Comment":
        """Add a reply to this comment.

        Convenience method equivalent to doc.add_comment(text, reply_to=self).

        Args:
            text: The reply text content
            author: Optional author override (uses document author if None)
            initials: Author initials (auto-generated from author if None)

        Returns:
            The created reply Comment object

        Raises:
            ValueError: If document reference is not available

        Example:
            >>> comment = doc.comments[0]
            >>> reply = comment.add_reply("I agree with this point", author="Bob")
        """
        if self._document is None:
            raise ValueError("Cannot add reply: no document reference")

        return self._document.add_comment(
            text=text,
            reply_to=self,
            author=author,
            initials=initials,
        )

    def __repr__(self) -> str:
        """String representation of the comment."""
        text_preview = self.text[:30] + "..." if len(self.text) > 30 else self.text
        resolved_str = " [RESOLVED]" if self.is_resolved else ""
        return f"<Comment id={self.id} author={self.author!r}: {text_preview!r}{resolved_str}>"

    def __eq__(self, other: object) -> bool:
        """Check equality based on comment ID."""
        if not isinstance(other, Comment):
            return NotImplemented
        return self.id == other.id

    def __hash__(self) -> int:
        """Hash based on comment ID."""
        return hash(self.id)
