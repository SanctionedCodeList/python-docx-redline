"""Tests for Phase 1 untracked editing infrastructure.

This module tests the core infrastructure added for untracked editing:
- create_plain_run() and create_plain_runs() in TrackedXMLGenerator
- include_deleted parameter in TextSearch.find_text()
- _remove_match() in TrackedChangeOperations
"""

from lxml import etree

from python_docx_redline.text_search import TextSearch, _is_run_in_deletion
from python_docx_redline.tracked_xml import TrackedXMLGenerator

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": WORD_NS}


def _w(tag: str) -> str:
    """Create a fully qualified Word namespace tag."""
    return f"{{{WORD_NS}}}{tag}"


class TestCreatePlainRun:
    """Test TrackedXMLGenerator.create_plain_run() method."""

    def test_creates_basic_run_element(self):
        """Test that create_plain_run returns a valid w:r element."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        run = gen.create_plain_run("Hello world")

        assert run.tag == _w("r")
        assert run.get(_w("rsidR")) is not None

    def test_contains_text_element(self):
        """Test that the run contains a w:t element with the text."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        run = gen.create_plain_run("Hello world")

        text_elem = run.find(_w("t"))
        assert text_elem is not None
        assert text_elem.text == "Hello world"

    def test_whitespace_preservation(self):
        """Test that leading/trailing whitespace gets xml:space='preserve'."""
        gen = TrackedXMLGenerator(author="TestAuthor")

        # Leading space
        run = gen.create_plain_run(" leading")
        text_elem = run.find(_w("t"))
        assert text_elem.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"

        # Trailing space
        run = gen.create_plain_run("trailing ")
        text_elem = run.find(_w("t"))
        assert text_elem.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"

        # No spaces
        run = gen.create_plain_run("nospaces")
        text_elem = run.find(_w("t"))
        assert text_elem.get("{http://www.w3.org/XML/1998/namespace}space") is None

    def test_copies_formatting_from_source_run(self):
        """Test that formatting is copied from source_run."""
        gen = TrackedXMLGenerator(author="TestAuthor")

        # Create a source run with bold formatting
        source_run = etree.Element(_w("r"))
        rpr = etree.SubElement(source_run, _w("rPr"))
        etree.SubElement(rpr, _w("b"))

        # Create new run with source formatting
        run = gen.create_plain_run("new text", source_run=source_run)

        # Should have the bold formatting
        new_rpr = run.find(_w("rPr"))
        assert new_rpr is not None
        assert new_rpr.find(_w("b")) is not None

    def test_no_formatting_without_source(self):
        """Test that run has no rPr without source_run."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        run = gen.create_plain_run("plain text")

        # Should not have rPr
        assert run.find(_w("rPr")) is None

    def test_rsid_matches_generator(self):
        """Test that run uses generator's RSID."""
        gen = TrackedXMLGenerator(author="TestAuthor", rsid="ABCD1234")
        run = gen.create_plain_run("text")

        assert run.get(_w("rsidR")) == "ABCD1234"


class TestCreatePlainRuns:
    """Test TrackedXMLGenerator.create_plain_runs() method."""

    def test_plain_text_single_run(self):
        """Test that plain text creates a single run."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        runs = gen.create_plain_runs("Hello world")

        assert len(runs) == 1
        text_elem = runs[0].find(_w("t"))
        assert text_elem.text == "Hello world"

    def test_bold_creates_formatted_run(self):
        """Test that **bold** creates a run with w:b."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        runs = gen.create_plain_runs("This is **bold** text")

        assert len(runs) == 3  # "This is ", "bold", " text"

        # Check the bold run (index 1)
        bold_run = runs[1]
        rpr = bold_run.find(_w("rPr"))
        assert rpr is not None
        assert rpr.find(_w("b")) is not None

        text = bold_run.find(_w("t"))
        assert text.text == "bold"

    def test_italic_creates_formatted_run(self):
        """Test that *italic* creates a run with w:i."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        runs = gen.create_plain_runs("This is *italic* text")

        assert len(runs) == 3

        italic_run = runs[1]
        rpr = italic_run.find(_w("rPr"))
        assert rpr is not None
        assert rpr.find(_w("i")) is not None

    def test_underline_creates_formatted_run(self):
        """Test that ++underline++ creates a run with w:u."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        runs = gen.create_plain_runs("This is ++underlined++ text")

        assert len(runs) == 3

        underline_run = runs[1]
        rpr = underline_run.find(_w("rPr"))
        assert rpr is not None
        u_elem = rpr.find(_w("u"))
        assert u_elem is not None
        assert u_elem.get(_w("val")) == "single"

    def test_strikethrough_creates_formatted_run(self):
        """Test that ~~strikethrough~~ creates a run with w:strike."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        runs = gen.create_plain_runs("This is ~~struck~~ text")

        assert len(runs) == 3

        strike_run = runs[1]
        rpr = strike_run.find(_w("rPr"))
        assert rpr is not None
        assert rpr.find(_w("strike")) is not None

    def test_multiple_formats(self):
        """Test text with multiple different formats."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        runs = gen.create_plain_runs("**bold** and *italic*")

        assert len(runs) >= 3

        # Find bold run
        bold_run = None
        italic_run = None
        for run in runs:
            rpr = run.find(_w("rPr"))
            if rpr is not None:
                if rpr.find(_w("b")) is not None:
                    bold_run = run
                if rpr.find(_w("i")) is not None:
                    italic_run = run

        assert bold_run is not None
        assert italic_run is not None

    def test_nested_formatting(self):
        """Test nested formatting like ***bold italic***."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        runs = gen.create_plain_runs("***bold italic***")

        # Find run with both bold and italic
        found = False
        for run in runs:
            rpr = run.find(_w("rPr"))
            if rpr is not None:
                has_bold = rpr.find(_w("b")) is not None
                has_italic = rpr.find(_w("i")) is not None
                if has_bold and has_italic:
                    found = True
                    break

        assert found

    def test_copies_base_formatting_from_source(self):
        """Test that source_run formatting is preserved with markdown."""
        gen = TrackedXMLGenerator(author="TestAuthor")

        # Create a source run with underline
        source_run = etree.Element(_w("r"))
        rpr = etree.SubElement(source_run, _w("rPr"))
        u = etree.SubElement(rpr, _w("u"))
        u.set(_w("val"), "single")

        # Create runs with bold markdown on top
        runs = gen.create_plain_runs("**bold text**", source_run=source_run)

        # Should have both underline (from source) and bold (from markdown)
        for run in runs:
            rpr = run.find(_w("rPr"))
            if rpr is not None:
                # Should have underline from source
                assert rpr.find(_w("u")) is not None
                # Should have bold from markdown
                assert rpr.find(_w("b")) is not None

    def test_linebreak_generates_w_br(self):
        """Test that hard line breaks generate w:br elements."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        runs = gen.create_plain_runs("line one  \nline two")

        # Find a run with w:br
        br_found = False
        for run in runs:
            if run.find(_w("br")) is not None:
                br_found = True
                break

        assert br_found


class TestIncludeDeletedParameter:
    """Test TextSearch.find_text() with include_deleted parameter."""

    def _create_paragraph_with_deletion(self) -> etree._Element:
        """Create a paragraph with normal and deleted text."""
        para = etree.Element(_w("p"))

        # Normal run with "Hello "
        run1 = etree.SubElement(para, _w("r"))
        t1 = etree.SubElement(run1, _w("t"))
        t1.text = "Hello "

        # Deleted text in w:del wrapper
        del_elem = etree.SubElement(para, _w("del"))
        del_elem.set(_w("id"), "1")
        del_elem.set(_w("author"), "Test")
        del_run = etree.SubElement(del_elem, _w("r"))
        del_text = etree.SubElement(del_run, _w("delText"))
        del_text.text = "deleted "

        # Normal run with "world"
        run2 = etree.SubElement(para, _w("r"))
        t2 = etree.SubElement(run2, _w("t"))
        t2.text = "world"

        return para

    def test_include_deleted_true_finds_deleted_text(self):
        """Test that include_deleted=True finds deleted text."""
        para = self._create_paragraph_with_deletion()
        search = TextSearch()

        # Should find "deleted" when include_deleted=True (default)
        matches = search.find_text("deleted", [para], include_deleted=True)
        assert len(matches) == 1

    def test_include_deleted_false_skips_deleted_text(self):
        """Test that include_deleted=False skips deleted text."""
        para = self._create_paragraph_with_deletion()
        search = TextSearch()

        # Should NOT find "deleted" when include_deleted=False
        matches = search.find_text("deleted", [para], include_deleted=False)
        assert len(matches) == 0

    def test_include_deleted_default_is_true(self):
        """Test that the default behavior includes deleted text."""
        para = self._create_paragraph_with_deletion()
        search = TextSearch()

        # Default should include deleted text
        matches = search.find_text("deleted", [para])
        assert len(matches) == 1

    def test_finds_normal_text_with_include_deleted_false(self):
        """Test that normal text is still found with include_deleted=False."""
        para = self._create_paragraph_with_deletion()
        search = TextSearch()

        # Should find "Hello" and "world" regardless
        matches = search.find_text("Hello", [para], include_deleted=False)
        assert len(matches) == 1

        matches = search.find_text("world", [para], include_deleted=False)
        assert len(matches) == 1

    def test_spans_across_deletion_with_include_deleted_false(self):
        """Test text spanning across a deletion with include_deleted=False."""
        para = self._create_paragraph_with_deletion()
        search = TextSearch()

        # "Hello world" should be found when deletions are excluded
        matches = search.find_text("Hello world", [para], include_deleted=False)
        assert len(matches) == 1

    def test_is_run_in_deletion_helper(self):
        """Test the _is_run_in_deletion helper function."""
        para = self._create_paragraph_with_deletion()

        # Get runs
        all_runs = list(para.iter(_w("r")))
        assert len(all_runs) == 3

        # First run is not in deletion
        assert _is_run_in_deletion(all_runs[0]) is False

        # Second run is in deletion
        assert _is_run_in_deletion(all_runs[1]) is True

        # Third run is not in deletion
        assert _is_run_in_deletion(all_runs[2]) is False

    def test_move_from_treated_as_deleted(self):
        """Test that runs in w:moveFrom are treated as deleted content."""
        para = etree.Element(_w("p"))

        # Normal run
        run1 = etree.SubElement(para, _w("r"))
        t1 = etree.SubElement(run1, _w("t"))
        t1.text = "Hello "

        # Text in moveFrom wrapper (semantically deleted at source)
        move_from = etree.SubElement(para, _w("moveFrom"))
        move_from.set(_w("id"), "1")
        move_run = etree.SubElement(move_from, _w("r"))
        del_text = etree.SubElement(move_run, _w("delText"))
        del_text.text = "moved"

        search = TextSearch()

        # Should find "moved" with include_deleted=True
        matches = search.find_text("moved", [para], include_deleted=True)
        assert len(matches) == 1

        # Should NOT find "moved" with include_deleted=False
        matches = search.find_text("moved", [para], include_deleted=False)
        assert len(matches) == 0


class TestRemoveMatch:
    """Test TrackedChangeOperations._remove_match() method."""

    def _create_document_mock(self, paragraph: etree._Element):
        """Create a mock document for TrackedChangeOperations."""
        from unittest.mock import Mock

        doc = Mock()
        doc.xml_root = etree.Element(_w("document"))
        body = etree.SubElement(doc.xml_root, _w("body"))
        body.append(paragraph)
        doc._xml_generator = TrackedXMLGenerator(author="Test")
        doc._text_search = TextSearch()
        return doc

    def test_remove_entire_single_run(self):
        """Test removing text that spans an entire run."""
        from python_docx_redline.operations.tracked_changes import TrackedChangeOperations

        # Create paragraph with single run
        para = etree.Element(_w("p"))
        run = etree.SubElement(para, _w("r"))
        t = etree.SubElement(run, _w("t"))
        t.text = "remove me"

        doc = self._create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Find the text
        matches = doc._text_search.find_text("remove me", [para])
        assert len(matches) == 1

        # Remove it
        ops._remove_match(matches[0])

        # Run should be gone
        assert len(list(para.iter(_w("r")))) == 0

    def test_remove_partial_single_run(self):
        """Test removing partial text from a single run."""
        from python_docx_redline.operations.tracked_changes import TrackedChangeOperations

        para = etree.Element(_w("p"))
        run = etree.SubElement(para, _w("r"))
        t = etree.SubElement(run, _w("t"))
        t.text = "hello world goodbye"

        doc = self._create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Find "world"
        matches = doc._text_search.find_text("world", [para])
        assert len(matches) == 1

        # Remove it
        ops._remove_match(matches[0])

        # Should have "hello " and " goodbye" remaining
        runs = list(para.iter(_w("r")))
        full_text = "".join(t.text or "" for r in runs for t in r.iter(_w("t")))
        assert "world" not in full_text
        assert "hello" in full_text
        assert "goodbye" in full_text

    def test_remove_from_start_of_run(self):
        """Test removing text from the start of a run."""
        from python_docx_redline.operations.tracked_changes import TrackedChangeOperations

        para = etree.Element(_w("p"))
        run = etree.SubElement(para, _w("r"))
        t = etree.SubElement(run, _w("t"))
        t.text = "remove this keep this"

        doc = self._create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Find "remove this "
        matches = doc._text_search.find_text("remove this ", [para])
        assert len(matches) == 1

        ops._remove_match(matches[0])

        # Should only have "keep this"
        runs = list(para.iter(_w("r")))
        full_text = "".join(t.text or "" for r in runs for t in r.iter(_w("t")))
        assert full_text == "keep this"

    def test_remove_from_end_of_run(self):
        """Test removing text from the end of a run."""
        from python_docx_redline.operations.tracked_changes import TrackedChangeOperations

        para = etree.Element(_w("p"))
        run = etree.SubElement(para, _w("r"))
        t = etree.SubElement(run, _w("t"))
        t.text = "keep this remove this"

        doc = self._create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Find " remove this"
        matches = doc._text_search.find_text(" remove this", [para])
        assert len(matches) == 1

        ops._remove_match(matches[0])

        # Should only have "keep this"
        runs = list(para.iter(_w("r")))
        full_text = "".join(t.text or "" for r in runs for t in r.iter(_w("t")))
        assert full_text == "keep this"

    def test_remove_spanning_multiple_runs(self):
        """Test removing text spanning multiple runs."""
        from python_docx_redline.operations.tracked_changes import TrackedChangeOperations

        para = etree.Element(_w("p"))

        run1 = etree.SubElement(para, _w("r"))
        t1 = etree.SubElement(run1, _w("t"))
        t1.text = "Hello "

        run2 = etree.SubElement(para, _w("r"))
        t2 = etree.SubElement(run2, _w("t"))
        t2.text = "beautiful "

        run3 = etree.SubElement(para, _w("r"))
        t3 = etree.SubElement(run3, _w("t"))
        t3.text = "world!"

        doc = self._create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Find "beautiful "
        matches = doc._text_search.find_text("beautiful ", [para])
        assert len(matches) == 1

        ops._remove_match(matches[0])

        # Should have "Hello " and "world!"
        runs = list(para.iter(_w("r")))
        full_text = "".join(t.text or "" for r in runs for t in r.iter(_w("t")))
        assert full_text == "Hello world!"

    def test_remove_partial_across_runs(self):
        """Test removing partial text that spans multiple runs."""
        from python_docx_redline.operations.tracked_changes import TrackedChangeOperations

        para = etree.Element(_w("p"))

        run1 = etree.SubElement(para, _w("r"))
        t1 = etree.SubElement(run1, _w("t"))
        t1.text = "AAA BBB"

        run2 = etree.SubElement(para, _w("r"))
        t2 = etree.SubElement(run2, _w("t"))
        t2.text = " CCC DDD"

        doc = self._create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Find "BBB CCC"
        matches = doc._text_search.find_text("BBB CCC", [para])
        assert len(matches) == 1

        ops._remove_match(matches[0])

        # Should have "AAA " and " DDD"
        runs = list(para.iter(_w("r")))
        full_text = "".join(t.text or "" for r in runs for t in r.iter(_w("t")))
        assert "BBB" not in full_text
        assert "CCC" not in full_text
        assert "AAA" in full_text
        assert "DDD" in full_text

    def test_remove_inside_tracked_insertion(self):
        """Test removing text inside a tracked insertion wrapper."""
        from python_docx_redline.operations.tracked_changes import TrackedChangeOperations

        para = etree.Element(_w("p"))

        # Normal run
        run1 = etree.SubElement(para, _w("r"))
        t1 = etree.SubElement(run1, _w("t"))
        t1.text = "Before "

        # Insertion wrapper with text
        ins = etree.SubElement(para, _w("ins"))
        ins.set(_w("id"), "1")
        ins.set(_w("author"), "Test")
        ins_run = etree.SubElement(ins, _w("r"))
        ins_t = etree.SubElement(ins_run, _w("t"))
        ins_t.text = "inserted"

        # Normal run
        run2 = etree.SubElement(para, _w("r"))
        t2 = etree.SubElement(run2, _w("t"))
        t2.text = " After"

        doc = self._create_document_mock(para)
        ops = TrackedChangeOperations(doc)

        # Find "inserted"
        matches = doc._text_search.find_text("inserted", [para])
        assert len(matches) == 1

        ops._remove_match(matches[0])

        # The insertion should be gone, but "Before " and " After" remain
        runs = list(para.iter(_w("r")))
        full_text = "".join(t.text or "" for r in runs for t in r.iter(_w("t")))
        assert "inserted" not in full_text
        assert "Before" in full_text
        assert "After" in full_text


class TestXMLGeneratorElements:
    """Test that generated elements are valid lxml Elements."""

    def test_create_plain_run_returns_element(self):
        """Test that create_plain_run returns an lxml Element."""
        gen = TrackedXMLGenerator(author="Test")
        run = gen.create_plain_run("text")

        assert isinstance(run, etree._Element)

    def test_create_plain_runs_returns_list_of_elements(self):
        """Test that create_plain_runs returns a list of lxml Elements."""
        gen = TrackedXMLGenerator(author="Test")
        runs = gen.create_plain_runs("**bold** text")

        assert isinstance(runs, list)
        assert all(isinstance(r, etree._Element) for r in runs)

    def test_elements_can_be_inserted_into_document(self):
        """Test that generated elements can be inserted into a document."""
        gen = TrackedXMLGenerator(author="Test")

        # Create a minimal paragraph
        para = etree.Element(_w("p"))

        # Create and insert a plain run
        run = gen.create_plain_run("new text")
        para.append(run)

        # Should be able to serialize
        xml_str = etree.tostring(para, encoding="unicode")
        assert "new text" in xml_str
        assert _w("r") in run.tag
