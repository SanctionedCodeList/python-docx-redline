"""
Tests for the accessibility layer types.

These tests verify the core types used in the DocTree accessibility layer:
- Ref parsing and properties
- ElementType enum
- AccessibilityNode dataclass
- ViewMode configuration
"""

from datetime import datetime

import pytest

from python_docx_redline.accessibility.types import (
    AccessibilityNode,
    ChangeInfo,
    ChangeType,
    CommentInfo,
    ElementType,
    Ref,
    ViewMode,
)


class TestRef:
    """Tests for Ref parsing and properties."""

    def test_parse_simple_paragraph_ref(self) -> None:
        """Test parsing a simple paragraph ref."""
        ref = Ref.parse("p:5")

        assert ref.path == "p:5"
        assert ref.element_type == ElementType.PARAGRAPH
        assert ref.ordinal == 5
        assert ref.fingerprint is None
        assert not ref.is_fingerprint

    def test_parse_table_ref(self) -> None:
        """Test parsing a table ref."""
        ref = Ref.parse("tbl:0")

        assert ref.path == "tbl:0"
        assert ref.element_type == ElementType.TABLE
        assert ref.ordinal == 0

    def test_parse_nested_ref(self) -> None:
        """Test parsing a nested ref with multiple segments."""
        ref = Ref.parse("tbl:0/row:2/cell:1")

        assert ref.path == "tbl:0/row:2/cell:1"
        assert ref.element_type == ElementType.TABLE_CELL
        assert ref.ordinal == 1
        assert ref.segments == [("tbl", "0"), ("row", "2"), ("cell", "1")]

    def test_parse_deeply_nested_ref(self) -> None:
        """Test parsing a deeply nested ref (paragraph in table cell)."""
        ref = Ref.parse("tbl:0/row:2/cell:1/p:0")

        assert ref.path == "tbl:0/row:2/cell:1/p:0"
        assert ref.element_type == ElementType.PARAGRAPH
        assert ref.ordinal == 0
        assert len(ref.segments) == 4

    def test_parse_fingerprint_ref(self) -> None:
        """Test parsing a fingerprint-based ref."""
        ref = Ref.parse("p:~xK4mNp2q")

        assert ref.path == "p:~xK4mNp2q"
        assert ref.element_type == ElementType.PARAGRAPH
        assert ref.ordinal is None
        assert ref.fingerprint == "xK4mNp2q"
        assert ref.is_fingerprint

    def test_parse_header_ref(self) -> None:
        """Test parsing a header ref."""
        ref = Ref.parse("hdr:0/p:0")

        assert ref.path == "hdr:0/p:0"
        assert ref.element_type == ElementType.PARAGRAPH
        assert ref.segments == [("hdr", "0"), ("p", "0")]

    def test_parse_empty_string_raises(self) -> None:
        """Test that parsing empty string raises ValueError."""
        with pytest.raises(ValueError, match="cannot be empty"):
            Ref.parse("")

    def test_parse_whitespace_only_raises(self) -> None:
        """Test that parsing whitespace raises ValueError."""
        with pytest.raises(ValueError, match="cannot be empty"):
            Ref.parse("   ")

    def test_parse_invalid_format_raises(self) -> None:
        """Test that invalid format raises ValueError."""
        with pytest.raises(ValueError, match="Invalid ref segment"):
            Ref.parse("paragraph5")  # Missing colon

    def test_parse_unknown_prefix_raises(self) -> None:
        """Test that unknown prefix raises ValueError."""
        with pytest.raises(ValueError, match="Unknown element type prefix"):
            Ref.parse("xyz:5")

    def test_parse_invalid_identifier_raises(self) -> None:
        """Test that invalid identifier raises ValueError."""
        with pytest.raises(ValueError, match="Invalid identifier"):
            Ref.parse("p:abc")  # Not integer or fingerprint

    def test_parent_path(self) -> None:
        """Test getting parent path."""
        ref = Ref.parse("tbl:0/row:2/cell:1")

        assert ref.parent_path == "tbl:0/row:2"

    def test_parent_path_top_level(self) -> None:
        """Test parent path for top-level ref returns None."""
        ref = Ref.parse("p:5")

        assert ref.parent_path is None

    def test_with_child(self) -> None:
        """Test creating child ref."""
        parent = Ref.parse("tbl:0/row:2")
        child = parent.with_child(ElementType.TABLE_CELL, 1)

        assert child.path == "tbl:0/row:2/cell:1"

    def test_with_child_fingerprint(self) -> None:
        """Test creating child ref with fingerprint."""
        parent = Ref.parse("tbl:0")
        child = parent.with_child(ElementType.TABLE_ROW, "~abc123")

        assert child.path == "tbl:0/row:~abc123"

    def test_ref_equality(self) -> None:
        """Test ref equality comparison."""
        ref1 = Ref.parse("p:5")
        ref2 = Ref.parse("p:5")
        ref3 = Ref.parse("p:6")

        assert ref1 == ref2
        assert ref1 != ref3

    def test_ref_string_equality(self) -> None:
        """Test ref equality with string."""
        ref = Ref.parse("p:5")

        assert ref == "p:5"
        assert ref != "p:6"

    def test_ref_hash(self) -> None:
        """Test ref hashing for use in sets/dicts."""
        ref1 = Ref.parse("p:5")
        ref2 = Ref.parse("p:5")

        refs = {ref1}
        assert ref2 in refs

    def test_ref_str(self) -> None:
        """Test ref string conversion."""
        ref = Ref.parse("tbl:0/row:2")

        assert str(ref) == "tbl:0/row:2"


class TestElementType:
    """Tests for ElementType enum."""

    def test_all_element_types_have_prefixes(self) -> None:
        """Test that key element types have prefix mappings."""
        from python_docx_redline.accessibility.types import ELEMENT_TYPE_TO_PREFIX

        key_types = [
            ElementType.PARAGRAPH,
            ElementType.RUN,
            ElementType.TABLE,
            ElementType.TABLE_ROW,
            ElementType.TABLE_CELL,
            ElementType.HEADER,
            ElementType.FOOTER,
        ]

        for elem_type in key_types:
            assert elem_type in ELEMENT_TYPE_TO_PREFIX


class TestChangeInfo:
    """Tests for ChangeInfo dataclass."""

    def test_create_change_info(self) -> None:
        """Test creating ChangeInfo."""
        change = ChangeInfo(
            change_type=ChangeType.INSERTION,
            author="Test Author",
            text="inserted text",
        )

        assert change.change_type == ChangeType.INSERTION
        assert change.author == "Test Author"
        assert change.text == "inserted text"
        assert change.date is None
        assert change.change_id is None

    def test_create_change_info_with_all_fields(self) -> None:
        """Test creating ChangeInfo with all fields."""
        now = datetime.now()
        change = ChangeInfo(
            change_type=ChangeType.DELETION,
            author="Test Author",
            date=now,
            change_id="42",
            text="deleted text",
        )

        assert change.change_type == ChangeType.DELETION
        assert change.date == now
        assert change.change_id == "42"


class TestCommentInfo:
    """Tests for CommentInfo dataclass."""

    def test_create_comment_info(self) -> None:
        """Test creating CommentInfo."""
        comment = CommentInfo(
            comment_id="1",
            author="Reviewer",
            text="Please review this section",
        )

        assert comment.comment_id == "1"
        assert comment.author == "Reviewer"
        assert comment.text == "Please review this section"
        assert comment.resolved is False
        assert comment.replies == []

    def test_create_comment_with_replies(self) -> None:
        """Test creating CommentInfo with replies."""
        reply = CommentInfo(
            comment_id="2",
            author="Author",
            text="Done",
        )
        comment = CommentInfo(
            comment_id="1",
            author="Reviewer",
            text="Please review",
            replies=[reply],
        )

        assert len(comment.replies) == 1
        assert comment.replies[0].text == "Done"


class TestViewMode:
    """Tests for ViewMode configuration."""

    def test_default_view_mode(self) -> None:
        """Test default ViewMode settings."""
        mode = ViewMode()

        assert mode.include_body is True
        assert mode.include_headers is False
        assert mode.include_footers is False
        assert mode.include_comments is False
        assert mode.include_tracked_changes is True
        assert mode.include_formatting is False
        assert mode.verbosity == "standard"

    def test_custom_view_mode(self) -> None:
        """Test custom ViewMode settings."""
        mode = ViewMode(
            include_body=True,
            include_headers=True,
            include_comments=True,
            verbosity="full",
        )

        assert mode.include_headers is True
        assert mode.include_comments is True
        assert mode.verbosity == "full"

    def test_invalid_verbosity_raises(self) -> None:
        """Test that invalid verbosity raises ValueError."""
        with pytest.raises(ValueError, match="verbosity must be one of"):
            ViewMode(verbosity="verbose")  # Invalid value


class TestAccessibilityNode:
    """Tests for AccessibilityNode dataclass."""

    def test_create_simple_node(self) -> None:
        """Test creating a simple node."""
        ref = Ref.parse("p:0")
        node = AccessibilityNode(
            ref=ref,
            element_type=ElementType.PARAGRAPH,
            text="Hello world",
        )

        assert node.ref == ref
        assert node.element_type == ElementType.PARAGRAPH
        assert node.text == "Hello world"
        assert node.children == []
        assert node.style is None
        assert not node.has_children
        assert not node.has_changes
        assert not node.has_comments

    def test_create_node_with_children(self) -> None:
        """Test creating a node with children."""
        parent_ref = Ref.parse("tbl:0")
        child_ref = Ref.parse("tbl:0/row:0")

        child = AccessibilityNode(
            ref=child_ref,
            element_type=ElementType.TABLE_ROW,
        )
        parent = AccessibilityNode(
            ref=parent_ref,
            element_type=ElementType.TABLE,
            children=[child],
        )

        assert parent.has_children
        assert len(parent.children) == 1

    def test_node_with_change_info(self) -> None:
        """Test node with tracked change."""
        ref = Ref.parse("p:0")
        change = ChangeInfo(
            change_type=ChangeType.INSERTION,
            author="Test",
            text="new text",
        )
        node = AccessibilityNode(
            ref=ref,
            element_type=ElementType.PARAGRAPH,
            change=change,
        )

        assert node.has_changes
        assert node.change.change_type == ChangeType.INSERTION

    def test_node_with_comments(self) -> None:
        """Test node with comments."""
        ref = Ref.parse("p:0")
        comment = CommentInfo(
            comment_id="1",
            author="Reviewer",
            text="Check this",
        )
        node = AccessibilityNode(
            ref=ref,
            element_type=ElementType.PARAGRAPH,
            comments=[comment],
        )

        assert node.has_comments
        assert len(node.comments) == 1

    def test_find_by_ref(self) -> None:
        """Test finding descendant by ref."""
        # Build a tree
        cell_ref = Ref.parse("tbl:0/row:0/cell:0")
        row_ref = Ref.parse("tbl:0/row:0")
        table_ref = Ref.parse("tbl:0")

        cell = AccessibilityNode(ref=cell_ref, element_type=ElementType.TABLE_CELL)
        row = AccessibilityNode(
            ref=row_ref,
            element_type=ElementType.TABLE_ROW,
            children=[cell],
        )
        table = AccessibilityNode(
            ref=table_ref,
            element_type=ElementType.TABLE,
            children=[row],
        )

        # Find the cell
        found = table.find_by_ref("tbl:0/row:0/cell:0")
        assert found is not None
        assert found.ref == cell_ref

    def test_find_by_ref_not_found(self) -> None:
        """Test finding non-existent ref returns None."""
        ref = Ref.parse("p:0")
        node = AccessibilityNode(ref=ref, element_type=ElementType.PARAGRAPH)

        assert node.find_by_ref("p:99") is None

    def test_find_all_by_type(self) -> None:
        """Test finding all nodes of a type."""
        # Build a tree with multiple paragraphs
        p1 = AccessibilityNode(
            ref=Ref.parse("p:0"),
            element_type=ElementType.PARAGRAPH,
        )
        p2 = AccessibilityNode(
            ref=Ref.parse("p:1"),
            element_type=ElementType.PARAGRAPH,
        )
        table = AccessibilityNode(
            ref=Ref.parse("tbl:0"),
            element_type=ElementType.TABLE,
        )

        # Create a document-level node
        doc = AccessibilityNode(
            ref=Ref.parse("p:0"),  # Placeholder ref
            element_type=ElementType.DOCUMENT,
            children=[p1, table, p2],
        )

        paragraphs = doc.find_all_by_type(ElementType.PARAGRAPH)
        assert len(paragraphs) == 2

        tables = doc.find_all_by_type(ElementType.TABLE)
        assert len(tables) == 1

    def test_node_with_style(self) -> None:
        """Test node with style information."""
        ref = Ref.parse("p:0")
        node = AccessibilityNode(
            ref=ref,
            element_type=ElementType.PARAGRAPH,
            style="Heading1",
            level=1,
        )

        assert node.style == "Heading1"
        assert node.level == 1

    def test_node_properties(self) -> None:
        """Test node with custom properties."""
        ref = Ref.parse("tbl:0")
        node = AccessibilityNode(
            ref=ref,
            element_type=ElementType.TABLE,
            properties={"rows": "3", "cols": "4"},
        )

        assert node.properties["rows"] == "3"
        assert node.properties["cols"] == "4"
