"""Tests for the track parameter on core editing operations.

This module tests that all core operations (insert, delete, replace, move)
work correctly with both track=True (tracked changes) and track=False
(untracked/silent edits).

Phase 2 of untracked editing implementation.
"""

from lxml import etree

from python_docx_redline.operations.tracked_changes import TrackedChangeOperations
from python_docx_redline.text_search import TextSearch
from python_docx_redline.tracked_xml import TrackedXMLGenerator

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": WORD_NS}


def _w(tag: str) -> str:
    """Create a fully qualified Word namespace tag."""
    return f"{{{WORD_NS}}}{tag}"


def _get_full_text(paragraph: etree._Element) -> str:
    """Extract all text from a paragraph."""
    texts = []
    for t_elem in paragraph.iter(_w("t")):
        if t_elem.text:
            texts.append(t_elem.text)
    for dt_elem in paragraph.iter(_w("delText")):
        if dt_elem.text:
            texts.append(dt_elem.text)
    return "".join(texts)


def _has_tracked_changes(paragraph: etree._Element) -> bool:
    """Check if paragraph has any tracked change markers."""
    for tag in ["ins", "del", "moveFrom", "moveTo"]:
        if paragraph.find(f".//{_w(tag)}") is not None:
            return True
    return False


def _create_document_mock(paragraph: etree._Element):
    """Create a mock document for TrackedChangeOperations."""
    from unittest.mock import Mock

    doc = Mock()
    doc.xml_root = etree.Element(_w("document"))
    body = etree.SubElement(doc.xml_root, _w("body"))
    body.append(paragraph)
    doc._xml_generator = TrackedXMLGenerator(author="Test")
    doc._text_search = TextSearch()
    return doc


def _create_simple_paragraph(text: str) -> etree._Element:
    """Create a simple paragraph with a single run."""
    para = etree.Element(_w("p"))
    run = etree.SubElement(para, _w("r"))
    t = etree.SubElement(run, _w("t"))
    t.text = text
    return para


def _create_multi_run_paragraph(*texts: str) -> etree._Element:
    """Create a paragraph with multiple runs."""
    para = etree.Element(_w("p"))
    for text in texts:
        run = etree.SubElement(para, _w("r"))
        t = etree.SubElement(run, _w("t"))
        t.text = text
    return para


def _create_formatted_paragraph(text: str) -> etree._Element:
    """Create a paragraph with a bold-formatted run."""
    para = etree.Element(_w("p"))
    run = etree.SubElement(para, _w("r"))
    rpr = etree.SubElement(run, _w("rPr"))
    etree.SubElement(rpr, _w("b"))  # Bold
    t = etree.SubElement(run, _w("t"))
    t.text = text
    return para


class TestInsertUntracked:
    """Test insert() with track=False (untracked insertion)."""

    def test_insert_untracked_basic(self):
        """Test basic untracked insertion creates no tracked change markers."""
        # Use separate runs to get the expected behavior
        para = _create_multi_run_paragraph("Hello", " world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.insert(" beautiful", after="Hello", track=False)

        # Should have the inserted text
        full_text = _get_full_text(para)
        assert "beautiful" in full_text

        # Should NOT have any tracked change markers
        assert not _has_tracked_changes(para)

    def test_insert_untracked_before(self):
        """Test untracked insertion before anchor."""
        para = _create_simple_paragraph("world!")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.insert("Hello ", before="world", track=False)

        full_text = _get_full_text(para)
        assert full_text == "Hello world!"
        assert not _has_tracked_changes(para)

    def test_insert_untracked_preserves_formatting(self):
        """Test that untracked insert inherits formatting from anchor."""
        para = _create_formatted_paragraph("Bold text here")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.insert(" inserted", after="Bold text", track=False)

        # Check that inserted run has bold formatting
        runs = list(para.iter(_w("r")))
        # Find runs with "inserted" text
        for run in runs:
            t = run.find(_w("t"))
            if t is not None and t.text and "inserted" in t.text:
                rpr = run.find(_w("rPr"))
                assert rpr is not None, "Inserted run should have rPr"
                assert rpr.find(_w("b")) is not None, "Inserted run should be bold"

    def test_insert_untracked_with_markdown(self):
        """Test that markdown formatting works in untracked inserts."""
        # Use separate runs so insert happens after "Hello" run
        para = _create_multi_run_paragraph("Hello", " world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.insert(" **bold**", after="Hello", track=False)

        # Should have the bold formatting
        runs = list(para.iter(_w("r")))
        found_bold = False
        for run in runs:
            rpr = run.find(_w("rPr"))
            if rpr is not None and rpr.find(_w("b")) is not None:
                t = run.find(_w("t"))
                # The markdown " **bold**" creates " " + "bold" in bold
                # So text could be " bold" or just "bold" depending on parsing
                if t is not None and t.text and "bold" in t.text:
                    found_bold = True

        assert found_bold, "Should have a bold run containing 'bold' text"
        assert not _has_tracked_changes(para)


class TestInsertTracked:
    """Test insert() with track=True (tracked insertion)."""

    def test_insert_tracked_still_works(self):
        """Test that track=True creates tracked insertion (backwards compat)."""
        # Use separate runs so insert happens after "Hello" run
        para = _create_multi_run_paragraph("Hello", " world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.insert(" beautiful", after="Hello", track=True)

        # Should have the inserted text
        full_text = _get_full_text(para)
        assert "beautiful" in full_text

        # Should have tracked insertion marker
        ins_elements = list(para.iter(_w("ins")))
        assert len(ins_elements) > 0, "Should have w:ins element"

    def test_insert_tracked_has_author(self):
        """Test that tracked insertion has author attribute."""
        para = _create_simple_paragraph("Hello world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.insert(" beautiful", after="Hello", track=True, author="TestAuthor")

        ins_elem = para.find(f".//{_w('ins')}")
        assert ins_elem is not None
        assert ins_elem.get(_w("author")) == "TestAuthor"


class TestDeleteUntracked:
    """Test delete() with track=False (untracked deletion)."""

    def test_delete_untracked_basic(self):
        """Test basic untracked deletion removes text without markers."""
        para = _create_simple_paragraph("Hello beautiful world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.delete("beautiful ", track=False)

        full_text = _get_full_text(para)
        assert "beautiful" not in full_text
        assert "Hello" in full_text
        assert "world" in full_text
        assert not _has_tracked_changes(para)

    def test_delete_untracked_entire_run(self):
        """Test deleting an entire run."""
        para = _create_multi_run_paragraph("First ", "Middle ", "Last")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.delete("Middle ", track=False)

        full_text = _get_full_text(para)
        assert full_text == "First Last"
        assert not _has_tracked_changes(para)

    def test_delete_untracked_partial_run(self):
        """Test deleting partial text from a run."""
        para = _create_simple_paragraph("Hello world goodbye")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.delete("world ", track=False)

        full_text = _get_full_text(para)
        assert "world" not in full_text
        assert "Hello" in full_text
        assert "goodbye" in full_text
        assert not _has_tracked_changes(para)

    def test_delete_untracked_spanning_runs(self):
        """Test deleting text spanning multiple runs."""
        para = _create_multi_run_paragraph("AAA BBB", " CCC DDD")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.delete("BBB CCC", track=False)

        full_text = _get_full_text(para)
        assert "BBB" not in full_text
        assert "CCC" not in full_text
        assert "AAA" in full_text
        assert "DDD" in full_text


class TestDeleteTracked:
    """Test delete() with track=True (tracked deletion)."""

    def test_delete_tracked_still_works(self):
        """Test that track=True creates tracked deletion (backwards compat)."""
        para = _create_simple_paragraph("Hello beautiful world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.delete("beautiful ", track=True)

        # Should have w:del element
        del_elements = list(para.iter(_w("del")))
        assert len(del_elements) > 0, "Should have w:del element"

        # Text should still be visible (as deleted text)
        full_text = _get_full_text(para)
        assert "beautiful" in full_text  # Still visible in delText

    def test_delete_tracked_has_author(self):
        """Test that tracked deletion has author attribute."""
        para = _create_simple_paragraph("Hello beautiful world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.delete("beautiful ", track=True, author="TestAuthor")

        del_elem = para.find(f".//{_w('del')}")
        assert del_elem is not None
        assert del_elem.get(_w("author")) == "TestAuthor"


class TestReplaceUntracked:
    """Test replace() with track=False (untracked replacement)."""

    def test_replace_untracked_basic(self):
        """Test basic untracked replacement."""
        para = _create_simple_paragraph("Hello old world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.replace("old", "new", track=False)

        full_text = _get_full_text(para)
        assert "old" not in full_text
        assert "new" in full_text
        assert "Hello" in full_text
        assert "world" in full_text
        assert not _has_tracked_changes(para)

    def test_replace_untracked_preserves_formatting(self):
        """Test that replacement inherits formatting from original."""
        para = _create_formatted_paragraph("Replace bold text")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.replace("bold", "formatted", track=False)

        # Find the replacement run
        runs = list(para.iter(_w("r")))
        found_formatted_bold = False
        for run in runs:
            t = run.find(_w("t"))
            if t is not None and t.text and "formatted" in t.text:
                rpr = run.find(_w("rPr"))
                if rpr is not None and rpr.find(_w("b")) is not None:
                    found_formatted_bold = True

        assert found_formatted_bold, "Replacement should preserve bold formatting"

    def test_replace_untracked_with_regex(self):
        """Test regex replacement works untracked."""
        para = _create_simple_paragraph("Product ABC123 is great")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.replace(r"ABC\d+", "XYZ999", regex=True, track=False)

        full_text = _get_full_text(para)
        assert "ABC123" not in full_text
        assert "XYZ999" in full_text
        assert not _has_tracked_changes(para)

    def test_replace_untracked_with_markdown(self):
        """Test that markdown formatting works in untracked replacements."""
        para = _create_simple_paragraph("This is plain text here")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.replace("plain", "**bold**", track=False)

        # Should have bold formatting
        runs = list(para.iter(_w("r")))
        found_bold = False
        for run in runs:
            rpr = run.find(_w("rPr"))
            if rpr is not None and rpr.find(_w("b")) is not None:
                t = run.find(_w("t"))
                if t is not None and t.text == "bold":
                    found_bold = True

        assert found_bold
        assert not _has_tracked_changes(para)

    def test_replace_untracked_occurrence_all(self):
        """Test replacing all occurrences untracked."""
        # Use separate runs to have distinct matches
        para = _create_multi_run_paragraph("foo", " bar ", "foo", " baz ", "foo")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.replace("foo", "qux", occurrence="all", track=False)

        full_text = _get_full_text(para)
        assert "foo" not in full_text
        assert full_text.count("qux") == 3
        assert not _has_tracked_changes(para)


class TestReplaceTracked:
    """Test replace() with track=True (tracked replacement)."""

    def test_replace_tracked_still_works(self):
        """Test that track=True creates tracked replacement (backwards compat)."""
        para = _create_simple_paragraph("Hello old world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.replace("old", "new", track=True)

        # Should have both w:del and w:ins elements
        del_elements = list(para.iter(_w("del")))
        ins_elements = list(para.iter(_w("ins")))
        assert len(del_elements) > 0, "Should have w:del element"
        assert len(ins_elements) > 0, "Should have w:ins element"

    def test_replace_tracked_has_author(self):
        """Test that tracked replacement has author on both del and ins."""
        para = _create_simple_paragraph("Hello old world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.replace("old", "new", track=True, author="TestAuthor")

        del_elem = para.find(f".//{_w('del')}")
        ins_elem = para.find(f".//{_w('ins')}")
        assert del_elem is not None
        assert ins_elem is not None
        assert del_elem.get(_w("author")) == "TestAuthor"
        assert ins_elem.get(_w("author")) == "TestAuthor"


class TestMoveUntracked:
    """Test move() with track=False (untracked move)."""

    def test_move_untracked_basic(self):
        """Test basic untracked move."""
        para = _create_simple_paragraph("First Second Third")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Move "Second" to after "Third"
        ops.move("Second", after="Third", track=False)

        full_text = _get_full_text(para)
        # "Second" should now appear after "Third"
        assert "First" in full_text
        assert "Third" in full_text
        assert "Second" in full_text
        # Order should be: First, Third, Second
        assert full_text.find("Third") < full_text.find("Second") or "ThirdSecond" in full_text
        assert not _has_tracked_changes(para)

    def test_move_untracked_before(self):
        """Test untracked move with before anchor."""
        para = _create_simple_paragraph("AAA BBB CCC")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Move "CCC" to before "AAA"
        ops.move("CCC", before="AAA", track=False)

        full_text = _get_full_text(para)
        assert "CCC" in full_text
        assert "AAA" in full_text
        # CCC should come before AAA now
        assert full_text.find("CCC") < full_text.find("AAA")
        assert not _has_tracked_changes(para)


class TestMoveTracked:
    """Test move() with track=True (tracked move)."""

    def test_move_tracked_still_works(self):
        """Test that track=True creates tracked move (backwards compat)."""
        para = _create_simple_paragraph("First Second Third")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.move("Second", after="Third", track=True)

        # Should have move markers
        move_from = list(para.iter(_w("moveFrom")))
        move_to = list(para.iter(_w("moveTo")))
        assert len(move_from) > 0 or para.find(f".//{_w('moveFromRangeStart')}") is not None
        assert len(move_to) > 0 or para.find(f".//{_w('moveToRangeStart')}") is not None


class TestMixedTracking:
    """Test combining tracked and untracked edits."""

    def test_mixed_tracked_and_untracked(self):
        """Test combining tracked and untracked edits in sequence."""
        para = _create_simple_paragraph("Hello world foo bar")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # First: untracked fix (silent typo correction)
        ops.replace("foo", "the", track=False)

        # Second: tracked substantive change
        ops.replace("world", "universe", track=True)

        full_text = _get_full_text(para)
        assert "the" in full_text  # Untracked replacement happened
        assert "universe" in full_text  # Tracked replacement happened

        # Should have tracked changes only for the second edit
        del_elements = list(para.iter(_w("del")))
        ins_elements = list(para.iter(_w("ins")))
        assert len(del_elements) > 0  # Has deletion marker
        assert len(ins_elements) > 0  # Has insertion marker

    def test_untracked_insert_then_tracked_delete(self):
        """Test untracked insert followed by tracked delete."""
        para = _create_simple_paragraph("Hello world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Insert text (untracked)
        ops.insert(" beautiful", after="Hello", track=False)

        # Delete text (tracked)
        ops.delete("world", track=True)

        full_text = _get_full_text(para)
        assert "beautiful" in full_text
        # world should still be there as deleted text
        assert "world" in full_text

        # Should have tracked deletion
        del_elements = list(para.iter(_w("del")))
        assert len(del_elements) > 0


class TestDefaultBehavior:
    """Test that default behavior (track=False) is untracked."""

    def test_insert_default_is_untracked(self):
        """Test that insert() defaults to untracked."""
        para = _create_simple_paragraph("Hello world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.insert(" beautiful", after="Hello")  # No track parameter

        assert not _has_tracked_changes(para)

    def test_delete_default_is_untracked(self):
        """Test that delete() defaults to untracked."""
        para = _create_simple_paragraph("Hello beautiful world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.delete("beautiful ")  # No track parameter

        assert not _has_tracked_changes(para)

    def test_replace_default_is_untracked(self):
        """Test that replace() defaults to untracked."""
        para = _create_simple_paragraph("Hello old world")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.replace("old", "new")  # No track parameter

        assert not _has_tracked_changes(para)

    def test_move_default_is_untracked(self):
        """Test that move() defaults to untracked."""
        para = _create_simple_paragraph("First Second Third")
        doc = _create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        ops.move("Second", after="Third")  # No track parameter

        assert not _has_tracked_changes(para)
