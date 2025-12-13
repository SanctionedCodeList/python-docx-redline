"""Integration tests for markdown formatting in tracked XML generation."""

from lxml import etree

from python_docx_redline.tracked_xml import TrackedXMLGenerator

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": WORD_NS}


def parse_insertion_xml(xml_str: str) -> etree._Element:
    """Parse insertion XML with proper namespace handling."""
    # Wrap with namespace declarations
    wrapped = f"""<?xml version="1.0" encoding="UTF-8"?>
    <root xmlns:w="{WORD_NS}"
          xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
          xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
        {xml_str}
    </root>"""
    root = etree.fromstring(wrapped.encode("utf-8"))
    return root.find(".//w:ins", namespaces=NSMAP)


class TestTrackedXMLWithMarkdown:
    """Test TrackedXMLGenerator with markdown formatting."""

    def test_plain_text_insertion(self):
        """Test that plain text creates a simple run."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("Hello world")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        assert len(runs) == 1
        text = runs[0].find(".//w:t", namespaces=NSMAP)
        assert text.text == "Hello world"

        # No run properties for plain text
        rpr = runs[0].find("w:rPr", namespaces=NSMAP)
        assert rpr is None

    def test_bold_text_insertion(self):
        """Test that **bold** creates a run with bold formatting."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("This is **bold** text")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        assert len(runs) == 3  # "This is ", "bold", " text"

        # Check the bold run
        bold_run = runs[1]
        rpr = bold_run.find("w:rPr", namespaces=NSMAP)
        assert rpr is not None
        bold_elem = rpr.find("w:b", namespaces=NSMAP)
        assert bold_elem is not None

        text = bold_run.find(".//w:t", namespaces=NSMAP)
        assert text.text == "bold"

    def test_italic_text_insertion(self):
        """Test that *italic* creates a run with italic formatting."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("This is *italic* text")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        assert len(runs) == 3

        # Check the italic run
        italic_run = runs[1]
        rpr = italic_run.find("w:rPr", namespaces=NSMAP)
        assert rpr is not None
        italic_elem = rpr.find("w:i", namespaces=NSMAP)
        assert italic_elem is not None

    def test_underline_text_insertion(self):
        """Test that ++underline++ creates a run with underline formatting."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("This is ++underlined++ text")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        assert len(runs) == 3

        # Check the underline run
        underline_run = runs[1]
        rpr = underline_run.find("w:rPr", namespaces=NSMAP)
        assert rpr is not None
        u_elem = rpr.find("w:u", namespaces=NSMAP)
        assert u_elem is not None
        assert u_elem.get(f"{{{WORD_NS}}}val") == "single"

    def test_strikethrough_text_insertion(self):
        """Test that ~~strikethrough~~ creates a run with strike formatting."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("This is ~~struck~~ text")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        assert len(runs) == 3

        # Check the strikethrough run
        strike_run = runs[1]
        rpr = strike_run.find("w:rPr", namespaces=NSMAP)
        assert rpr is not None
        strike_elem = rpr.find("w:strike", namespaces=NSMAP)
        assert strike_elem is not None

    def test_multiple_formats_in_one_insertion(self):
        """Test insertion with multiple different formats."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("**bold** and *italic* and ++underline++")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        # Should have 5 runs: bold, " and ", italic, " and ", underline
        assert len(runs) == 5

        # Check bold run
        bold_rpr = runs[0].find("w:rPr", namespaces=NSMAP)
        assert bold_rpr is not None
        assert bold_rpr.find("w:b", namespaces=NSMAP) is not None

        # Check italic run (index 2)
        italic_rpr = runs[2].find("w:rPr", namespaces=NSMAP)
        assert italic_rpr is not None
        assert italic_rpr.find("w:i", namespaces=NSMAP) is not None

        # Check underline run (index 4)
        underline_rpr = runs[4].find("w:rPr", namespaces=NSMAP)
        assert underline_rpr is not None
        assert underline_rpr.find("w:u", namespaces=NSMAP) is not None

    def test_nested_bold_italic(self):
        """Test nested formatting like ***bold italic***."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("This is ***bold and italic*** text")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        # Find the run with both bold and italic
        combined_run = None
        for run in runs:
            rpr = run.find("w:rPr", namespaces=NSMAP)
            if rpr is not None:
                has_bold = rpr.find("w:b", namespaces=NSMAP) is not None
                has_italic = rpr.find("w:i", namespaces=NSMAP) is not None
                if has_bold and has_italic:
                    combined_run = run
                    break

        assert combined_run is not None
        text = combined_run.find(".//w:t", namespaces=NSMAP)
        assert text.text == "bold and italic"

    def test_whitespace_preservation(self):
        """Test that leading/trailing whitespace gets xml:space attribute."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("This is ** bold ** text")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        # Check that runs with leading/trailing spaces have xml:space
        for run in runs:
            text_elem = run.find(".//w:t", namespaces=NSMAP)
            if text_elem is not None and text_elem.text:
                text = text_elem.text
                if text.startswith(" ") or text.endswith(" "):
                    # Should have xml:space="preserve"
                    assert (
                        text_elem.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
                    )

    def test_all_bold_text(self):
        """Test text that is entirely bold."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("**all bold**")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        assert len(runs) == 1
        rpr = runs[0].find("w:rPr", namespaces=NSMAP)
        assert rpr is not None
        assert rpr.find("w:b", namespaces=NSMAP) is not None

        text = runs[0].find(".//w:t", namespaces=NSMAP)
        assert text.text == "all bold"


class TestTrackedXMLAttributes:
    """Test that XML attributes are correctly generated with markdown."""

    def test_change_id_increments(self):
        """Test that change IDs increment across multiple insertions."""
        gen = TrackedXMLGenerator(author="TestAuthor")

        xml1 = gen.create_insertion("**bold** text")
        xml2 = gen.create_insertion("*italic* text")

        ins1 = parse_insertion_xml(xml1)
        ins2 = parse_insertion_xml(xml2)

        id1 = int(ins1.get(f"{{{WORD_NS}}}id"))
        id2 = int(ins2.get(f"{{{WORD_NS}}}id"))

        assert id2 > id1

    def test_author_attribute(self):
        """Test that author is correctly set."""
        gen = TrackedXMLGenerator(author="MyAuthor")
        xml = gen.create_insertion("**bold** text")

        ins = parse_insertion_xml(xml)
        assert ins.get(f"{{{WORD_NS}}}author") == "MyAuthor"

    def test_rsid_on_runs(self):
        """Test that RSID is set on all runs."""
        gen = TrackedXMLGenerator(author="TestAuthor", rsid="12345678")
        xml = gen.create_insertion("**bold** and *italic*")

        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)

        for run in runs:
            rsid = run.get(f"{{{WORD_NS}}}rsidR")
            assert rsid == "12345678"

    def test_date_attributes(self):
        """Test that date attributes are present."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        xml = gen.create_insertion("**bold**")

        ins = parse_insertion_xml(xml)

        # Check w:date attribute
        date = ins.get(f"{{{WORD_NS}}}date")
        assert date is not None
        assert "T" in date  # ISO 8601 format


class TestXMLValidity:
    """Test that generated XML is valid."""

    def test_xml_is_well_formed(self):
        """Test that generated XML is well-formed."""
        gen = TrackedXMLGenerator(author="TestAuthor")

        # Various markdown inputs
        inputs = [
            "plain text",
            "**bold** text",
            "*italic* text",
            "++underline++ text",
            "~~strike~~ text",
            "**bold** and *italic* and ++underline++ and ~~strike~~",
            "***bold italic***",
            "Text with **special <chars>** & symbols",
        ]

        for text in inputs:
            xml = gen.create_insertion(text)
            # Should not raise
            ins = parse_insertion_xml(xml)
            assert ins is not None

    def test_special_characters_escaped(self):
        """Test that special XML characters are escaped."""
        gen = TrackedXMLGenerator(author="TestAuthor")
        # Use text that won't be interpreted as HTML
        xml = gen.create_insertion("**A > B** & *C < D*")

        ins = parse_insertion_xml(xml)

        # If we got here, the XML was valid (special chars were escaped)
        runs = ins.findall(".//w:r", namespaces=NSMAP)
        assert len(runs) >= 2

        # Verify the text content has the special chars
        texts = [
            r.find(".//w:t", namespaces=NSMAP).text
            for r in runs
            if r.find(".//w:t", namespaces=NSMAP) is not None
        ]
        full_text = "".join(t for t in texts if t)
        assert ">" in full_text
        assert "&" in full_text
        assert "<" in full_text


class TestRegressionFixesXML:
    """Regression tests for PR review fixes - XML generation."""

    def test_whitespace_only_creates_run(self):
        """Test that whitespace-only input creates a valid run.

        Regression test: empty segments would create <w:ins> with no <w:r>.
        """
        gen = TrackedXMLGenerator(author="TestAuthor")

        # Single space
        xml = gen.create_insertion(" ")
        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)
        assert len(runs) >= 1

        # Multiple spaces
        xml = gen.create_insertion("   ")
        ins = parse_insertion_xml(xml)
        runs = ins.findall(".//w:r", namespaces=NSMAP)
        assert len(runs) >= 1

        # Verify xml:space="preserve" is set
        t_elem = ins.find(".//w:t", namespaces=NSMAP)
        assert t_elem is not None
        xml_space = t_elem.get("{http://www.w3.org/XML/1998/namespace}space")
        assert xml_space == "preserve"

    def test_linebreak_generates_w_br(self):
        """Test that hard line breaks generate <w:br/> elements.

        Regression test: linebreak() was emitting newline in <w:t>, but
        Word expects <w:br/> elements.
        """
        gen = TrackedXMLGenerator(author="TestAuthor")

        # Hard line break: two trailing spaces + newline
        xml = gen.create_insertion("line one  \nline two")
        ins = parse_insertion_xml(xml)

        # Should have <w:br/> element
        br_elements = ins.findall(".//w:br", namespaces=NSMAP)
        assert len(br_elements) >= 1, "Expected at least one <w:br/> element"

    def test_linebreak_with_formatting(self):
        """Test that line breaks work with formatted text."""
        gen = TrackedXMLGenerator(author="TestAuthor")

        # Bold text with line break
        xml = gen.create_insertion("**bold line**  \n*italic line*")
        ins = parse_insertion_xml(xml)

        # Should have runs with formatting
        runs = ins.findall(".//w:r", namespaces=NSMAP)
        assert len(runs) >= 3  # At least: bold text, br, italic text

        # Should have <w:br/> somewhere
        br_elements = ins.findall(".//w:br", namespaces=NSMAP)
        assert len(br_elements) >= 1
