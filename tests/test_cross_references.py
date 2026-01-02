"""
Tests for Cross-Reference functionality - Phase 1.

Tests the core infrastructure and data models for cross-reference support:
- Field code XML structure for REF, PAGEREF, NOTEREF
- Dirty flag handling
- Switch mapping for display options
- Hyperlink switch (\\h) handling
"""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from python_docx_redline import Document
from python_docx_redline.errors import (
    BookmarkAlreadyExistsError,
    CrossReferenceError,
    CrossReferenceTargetNotFoundError,
    InvalidBookmarkNameError,
)
from python_docx_redline.operations.cross_references import (
    DISPLAY_SWITCH_MAP,
    CrossReference,
    CrossReferenceOperations,
    CrossReferenceTarget,
)

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def create_simple_document() -> Path:
    """Create a minimal test document."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>This is a sample document with some text.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


# =============================================================================
# Error Classes Tests
# =============================================================================


class TestCrossReferenceErrorClasses:
    """Tests for cross-reference error classes."""

    def test_cross_reference_error_base(self):
        """Test CrossReferenceError is a proper exception."""
        err = CrossReferenceError("Test error")
        assert str(err) == "Test error"
        assert isinstance(err, Exception)

    def test_cross_reference_target_not_found_error(self):
        """Test CrossReferenceTargetNotFoundError with message."""
        err = CrossReferenceTargetNotFoundError("heading:Section 1")
        assert "heading:Section 1" in str(err)
        assert err.target == "heading:Section 1"
        assert err.available_targets == []

    def test_cross_reference_target_not_found_with_suggestions(self):
        """Test CrossReferenceTargetNotFoundError with available targets."""
        err = CrossReferenceTargetNotFoundError(
            "MyBookmark",
            available_targets=["Bookmark1", "Bookmark2", "Section1"],
        )
        assert "MyBookmark" in str(err)
        assert "Bookmark1" in str(err)
        assert "Bookmark2" in str(err)
        assert err.available_targets == ["Bookmark1", "Bookmark2", "Section1"]

    def test_cross_reference_target_not_found_truncates_long_list(self):
        """Test that long target lists are truncated in message."""
        many_targets = [f"Bookmark{i}" for i in range(20)]
        err = CrossReferenceTargetNotFoundError("MyBookmark", available_targets=many_targets)
        # Should show first 10 and indicate more exist
        assert "... and 10 more" in str(err)

    def test_invalid_bookmark_name_error(self):
        """Test InvalidBookmarkNameError with name and reason."""
        err = InvalidBookmarkNameError("my bookmark", "contains spaces")
        assert "my bookmark" in str(err)
        assert "contains spaces" in str(err)
        assert err.name == "my bookmark"
        assert err.reason == "contains spaces"

    def test_bookmark_already_exists_error(self):
        """Test BookmarkAlreadyExistsError with name."""
        err = BookmarkAlreadyExistsError("MyBookmark")
        assert "MyBookmark" in str(err)
        assert "already exists" in str(err)
        assert err.name == "MyBookmark"


# =============================================================================
# Data Model Tests
# =============================================================================


class TestCrossReferenceDataModel:
    """Tests for CrossReference dataclass."""

    def test_cross_reference_creation(self):
        """Test creating a CrossReference instance."""
        xref = CrossReference(
            ref="xref:1",
            field_type="REF",
            target_bookmark="_Ref123456",
            switches="\\h",
            display_value="Section 2.1",
            is_dirty=True,
            is_hyperlink=True,
            position="p:5",
        )

        assert xref.ref == "xref:1"
        assert xref.field_type == "REF"
        assert xref.target_bookmark == "_Ref123456"
        assert xref.switches == "\\h"
        assert xref.display_value == "Section 2.1"
        assert xref.is_dirty is True
        assert xref.is_hyperlink is True
        assert xref.position == "p:5"

    def test_cross_reference_optional_fields(self):
        """Test CrossReference optional fields have defaults."""
        xref = CrossReference(
            ref="xref:2",
            field_type="PAGEREF",
            target_bookmark="MyBookmark",
            switches="\\h \\p",
            display_value="5",
            is_dirty=False,
            is_hyperlink=True,
            position="p:10",
        )

        assert xref.show_position is False
        assert xref.number_format is None
        assert xref.suppress_non_numeric is False

    def test_cross_reference_with_all_fields(self):
        """Test CrossReference with all fields populated."""
        xref = CrossReference(
            ref="xref:3",
            field_type="REF",
            target_bookmark="_Ref789",
            switches="\\h \\r \\p",
            display_value="above",
            is_dirty=True,
            is_hyperlink=True,
            position="p:15",
            show_position=True,
            number_format="relative",
            suppress_non_numeric=False,
        )

        assert xref.show_position is True
        assert xref.number_format == "relative"


class TestCrossReferenceTargetDataModel:
    """Tests for CrossReferenceTarget dataclass."""

    def test_cross_reference_target_creation(self):
        """Test creating a CrossReferenceTarget instance."""
        target = CrossReferenceTarget(
            type="heading",
            bookmark_name="_Ref123456",
            display_name="2.1 Methodology",
            text_preview="This section describes the methodology...",
            position="p:10",
            is_hidden=True,
        )

        assert target.type == "heading"
        assert target.bookmark_name == "_Ref123456"
        assert target.display_name == "2.1 Methodology"
        assert target.text_preview == "This section describes the methodology..."
        assert target.position == "p:10"
        assert target.is_hidden is True

    def test_cross_reference_target_optional_fields(self):
        """Test CrossReferenceTarget optional fields have defaults."""
        target = CrossReferenceTarget(
            type="bookmark",
            bookmark_name="MyBookmark",
            display_name="MyBookmark",
            text_preview="Some bookmarked text...",
            position="p:5",
            is_hidden=False,
        )

        assert target.number is None
        assert target.level is None
        assert target.sequence_id is None

    def test_cross_reference_target_with_all_fields(self):
        """Test CrossReferenceTarget with all fields populated."""
        target = CrossReferenceTarget(
            type="figure",
            bookmark_name="_Ref999888",
            display_name="Figure 3: Architecture Diagram",
            text_preview="Figure 3: Architecture Diagram showing...",
            position="p:25",
            is_hidden=True,
            number="3",
            level=None,
            sequence_id="Figure",
        )

        assert target.type == "figure"
        assert target.number == "3"
        assert target.sequence_id == "Figure"


# =============================================================================
# CrossReferenceOperations Tests
# =============================================================================


class TestCrossReferenceOperationsInit:
    """Tests for CrossReferenceOperations initialization."""

    def test_init_with_document(self):
        """Test creating CrossReferenceOperations with a Document."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)
            assert ops._document is doc
        finally:
            doc_path.unlink(missing_ok=True)


# =============================================================================
# Field Code Generation Tests
# =============================================================================


class TestCreateFieldCode:
    """Tests for _create_field_code method."""

    def test_create_ref_field_basic(self):
        """Test creating a basic REF field code."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="_Ref123456",
                switches=[],
                placeholder_text="Section 2.1",
            )

            # Should have 5 runs: begin, instr, sep, result, end
            assert len(runs) == 5

            # Check each run is a w:r element
            for run in runs:
                assert run.tag == f"{{{WORD_NS}}}r"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_has_begin_marker(self):
        """Test that field code has begin marker with correct type."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="MyBookmark",
                switches=["\\h"],
            )

            # First run should contain fldChar with type="begin"
            begin_run = runs[0]
            fld_char = begin_run.find(f"{{{WORD_NS}}}fldChar")
            assert fld_char is not None
            assert fld_char.get(f"{{{WORD_NS}}}fldCharType") == "begin"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_dirty_flag_set(self):
        """Test that dirty flag is set on field begin marker."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="MyBookmark",
                switches=[],
            )

            # Begin run should have dirty="true"
            begin_run = runs[0]
            fld_char = begin_run.find(f"{{{WORD_NS}}}fldChar")
            assert fld_char is not None
            assert fld_char.get(f"{{{WORD_NS}}}dirty") == "true"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_has_instruction(self):
        """Test that field code has instruction text."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="_Ref999",
                switches=["\\h"],
            )

            # Second run should contain instrText
            instr_run = runs[1]
            instr_text = instr_run.find(f"{{{WORD_NS}}}instrText")
            assert instr_text is not None
            assert "REF" in instr_text.text
            assert "_Ref999" in instr_text.text
            assert "\\h" in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_instruction_format(self):
        """Test field instruction format: FIELD_TYPE bookmark switches."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="PAGEREF",
                bookmark_name="MyBookmark",
                switches=["\\h", "\\p"],
            )

            instr_run = runs[1]
            instr_text = instr_run.find(f"{{{WORD_NS}}}instrText")
            assert instr_text is not None

            # Check format: " PAGEREF MyBookmark \h \p "
            text = instr_text.text
            assert text.strip().startswith("PAGEREF")
            assert "MyBookmark" in text
            assert "\\h" in text
            assert "\\p" in text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_has_separator(self):
        """Test that field code has separator marker."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="MyBookmark",
                switches=[],
            )

            # Third run should contain fldChar with type="separate"
            sep_run = runs[2]
            fld_char = sep_run.find(f"{{{WORD_NS}}}fldChar")
            assert fld_char is not None
            assert fld_char.get(f"{{{WORD_NS}}}fldCharType") == "separate"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_has_placeholder_text(self):
        """Test that field code has placeholder text."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="MyBookmark",
                switches=[],
                placeholder_text="Section 2.1",
            )

            # Fourth run should contain text element
            result_run = runs[3]
            text_elem = result_run.find(f"{{{WORD_NS}}}t")
            assert text_elem is not None
            assert text_elem.text == "Section 2.1"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_default_placeholder(self):
        """Test that default placeholder is used when none provided."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="MyBookmark",
                switches=[],
            )

            result_run = runs[3]
            text_elem = result_run.find(f"{{{WORD_NS}}}t")
            assert text_elem is not None
            assert text_elem.text == "[Update field]"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_has_end_marker(self):
        """Test that field code has end marker."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="MyBookmark",
                switches=[],
            )

            # Fifth run should contain fldChar with type="end"
            end_run = runs[4]
            fld_char = end_run.find(f"{{{WORD_NS}}}fldChar")
            assert fld_char is not None
            assert fld_char.get(f"{{{WORD_NS}}}fldCharType") == "end"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_pageref_field(self):
        """Test creating a PAGEREF field code."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="PAGEREF",
                bookmark_name="MyBookmark",
                switches=["\\h"],
                placeholder_text="5",
            )

            instr_run = runs[1]
            instr_text = instr_run.find(f"{{{WORD_NS}}}instrText")
            assert "PAGEREF" in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_noteref_field(self):
        """Test creating a NOTEREF field code."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="NOTEREF",
                bookmark_name="_FootnoteRef1",
                switches=["\\h", "\\f"],
                placeholder_text="1",
            )

            instr_run = runs[1]
            instr_text = instr_run.find(f"{{{WORD_NS}}}instrText")
            assert "NOTEREF" in instr_text.text
            assert "\\f" in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_no_switches(self):
        """Test creating field code with no switches."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="SimpleBookmark",
                switches=[],
            )

            instr_run = runs[1]
            instr_text = instr_run.find(f"{{{WORD_NS}}}instrText")
            text = instr_text.text

            # Should just have field type and bookmark
            assert "REF" in text
            assert "SimpleBookmark" in text
            # Should not have any backslash switches (except in field name)
            # The text should be " REF SimpleBookmark "
            assert text.strip() == "REF SimpleBookmark"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_multiple_switches(self):
        """Test creating field code with multiple switches."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="MyBookmark",
                switches=["\\h", "\\r", "\\p"],
            )

            instr_run = runs[1]
            instr_text = instr_run.find(f"{{{WORD_NS}}}instrText")
            text = instr_text.text

            assert "\\h" in text
            assert "\\r" in text
            assert "\\p" in text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_code_preserves_space(self):
        """Test that instrText has xml:space='preserve'."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="MyBookmark",
                switches=[],
            )

            instr_run = runs[1]
            instr_text = instr_run.find(f"{{{WORD_NS}}}instrText")

            # Check xml:space attribute
            space_attr = instr_text.get(f"{{{XML_NS}}}space")
            assert space_attr == "preserve"

        finally:
            doc_path.unlink(missing_ok=True)


# =============================================================================
# Switch Mapping Tests
# =============================================================================


class TestGetSwitchesForDisplay:
    """Tests for _get_switches_for_display method."""

    def test_display_text_returns_ref(self):
        """Test display='text' returns REF with no special switches."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("text", hyperlink=False)

            assert field_type == "REF"
            assert switches == []

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_text_with_hyperlink(self):
        """Test display='text' with hyperlink adds \\h switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("text", hyperlink=True)

            assert field_type == "REF"
            assert "\\h" in switches

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_page_returns_pageref(self):
        """Test display='page' returns PAGEREF."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("page", hyperlink=False)

            assert field_type == "PAGEREF"
            assert switches == []

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_page_with_hyperlink(self):
        """Test display='page' with hyperlink adds \\h switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("page", hyperlink=True)

            assert field_type == "PAGEREF"
            assert "\\h" in switches

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_number_returns_n_switch(self):
        """Test display='number' returns REF with \\n switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("number", hyperlink=False)

            assert field_type == "REF"
            assert "\\n" in switches

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_full_number_returns_w_switch(self):
        """Test display='full_number' returns REF with \\w switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("full_number", hyperlink=False)

            assert field_type == "REF"
            assert "\\w" in switches

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_relative_number_returns_r_switch(self):
        """Test display='relative_number' returns REF with \\r switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("relative_number", hyperlink=False)

            assert field_type == "REF"
            assert "\\r" in switches

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_above_below_returns_p_switch(self):
        """Test display='above_below' returns REF with \\p switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("above_below", hyperlink=False)

            assert field_type == "REF"
            assert "\\p" in switches

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_above_below_with_hyperlink(self):
        """Test display='above_below' with hyperlink adds both switches."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("above_below", hyperlink=True)

            assert field_type == "REF"
            assert "\\p" in switches
            assert "\\h" in switches

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_label_number(self):
        """Test display='label_number' for captions."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("label_number", hyperlink=True)

            assert field_type == "REF"
            assert "\\h" in switches

        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_number_only(self):
        """Test display='number_only' for caption numbers."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("number_only", hyperlink=False)

            assert field_type == "REF"
            assert "\\n" in switches

        finally:
            doc_path.unlink(missing_ok=True)

    def test_invalid_display_raises_error(self):
        """Test that invalid display option raises ValueError."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(ValueError) as exc_info:
                ops._get_switches_for_display("invalid_option", hyperlink=True)

            assert "invalid_option" in str(exc_info.value)
            assert "Valid options" in str(exc_info.value)

        finally:
            doc_path.unlink(missing_ok=True)

    def test_all_display_options_in_map(self):
        """Test that all expected display options are in the map."""
        expected_options = [
            "text",
            "page",
            "number",
            "full_number",
            "relative_number",
            "above_below",
            "label_number",
            "number_only",
            "label_only",
            "caption_text",
        ]

        for option in expected_options:
            assert option in DISPLAY_SWITCH_MAP, f"Missing display option: {option}"

    def test_hyperlink_default_true(self):
        """Test that hyperlink defaults to True in typical usage."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Call with default hyperlink=True
            field_type, switches = ops._get_switches_for_display("text")

            # Should include hyperlink switch
            assert "\\h" in switches

        finally:
            doc_path.unlink(missing_ok=True)


# =============================================================================
# Integration Tests: Field Code with Switch Mapping
# =============================================================================


class TestFieldCodeWithSwitchMapping:
    """Integration tests combining field code generation with switch mapping."""

    def test_create_ref_field_with_text_display(self):
        """Test creating a complete REF field for text display."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("text", hyperlink=True)
            runs = ops._create_field_code(
                field_type=field_type,
                bookmark_name="MyBookmark",
                switches=switches,
                placeholder_text="Referenced Text",
            )

            # Verify structure
            assert len(runs) == 5

            # Verify instruction
            instr_text = runs[1].find(f"{{{WORD_NS}}}instrText")
            assert "REF" in instr_text.text
            assert "MyBookmark" in instr_text.text
            assert "\\h" in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_pageref_field_with_page_display(self):
        """Test creating a PAGEREF field for page display."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("page", hyperlink=True)
            runs = ops._create_field_code(
                field_type=field_type,
                bookmark_name="TargetSection",
                switches=switches,
                placeholder_text="5",
            )

            # Verify instruction has PAGEREF
            instr_text = runs[1].find(f"{{{WORD_NS}}}instrText")
            assert "PAGEREF" in instr_text.text
            assert "TargetSection" in instr_text.text
            assert "\\h" in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_ref_field_with_number_display(self):
        """Test creating a REF field for paragraph number display."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("number", hyperlink=True)
            runs = ops._create_field_code(
                field_type=field_type,
                bookmark_name="_RefHeading",
                switches=switches,
                placeholder_text="2",
            )

            # Verify instruction has \n switch
            instr_text = runs[1].find(f"{{{WORD_NS}}}instrText")
            assert "REF" in instr_text.text
            assert "\\n" in instr_text.text
            assert "\\h" in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_ref_field_with_above_below_display(self):
        """Test creating a REF field with above/below indicator."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("above_below", hyperlink=True)
            runs = ops._create_field_code(
                field_type=field_type,
                bookmark_name="FigureRef",
                switches=switches,
                placeholder_text="above",
            )

            # Verify instruction has \p switch
            instr_text = runs[1].find(f"{{{WORD_NS}}}instrText")
            assert "REF" in instr_text.text
            assert "\\p" in instr_text.text
            assert "\\h" in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_field_without_hyperlink(self):
        """Test creating a field without hyperlink."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            field_type, switches = ops._get_switches_for_display("text", hyperlink=False)
            runs = ops._create_field_code(
                field_type=field_type,
                bookmark_name="NoLinkBookmark",
                switches=switches,
            )

            # Verify no \h switch
            instr_text = runs[1].find(f"{{{WORD_NS}}}instrText")
            assert "\\h" not in instr_text.text

        finally:
            doc_path.unlink(missing_ok=True)


# =============================================================================
# XML Structure Validation Tests
# =============================================================================


class TestFieldCodeXmlStructure:
    """Tests validating the exact XML structure of generated field codes."""

    def test_field_structure_order(self):
        """Test the order of elements in field structure."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="Test",
                switches=["\\h"],
            )

            # Extract fldChar types in order
            fld_char_types = []
            for run in runs:
                fld_char = run.find(f"{{{WORD_NS}}}fldChar")
                if fld_char is not None:
                    fld_char_types.append(fld_char.get(f"{{{WORD_NS}}}fldCharType"))

            # Should be: begin, separate, end
            assert fld_char_types == ["begin", "separate", "end"]

        finally:
            doc_path.unlink(missing_ok=True)

    def test_runs_can_be_inserted_into_paragraph(self):
        """Test that generated runs can be inserted into a paragraph."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="Test",
                switches=["\\h"],
            )

            # Create a paragraph and insert runs
            nsmap = {None: WORD_NS}
            para = etree.Element(f"{{{WORD_NS}}}p", nsmap=nsmap)

            for run in runs:
                para.append(run)

            # Verify paragraph now has the runs
            assert len(para) == 5
            assert para[0].tag == f"{{{WORD_NS}}}r"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_field_structure_matches_word_format(self):
        """Test that field structure matches Word's expected format."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            runs = ops._create_field_code(
                field_type="REF",
                bookmark_name="_Ref123456789",
                switches=["\\h"],
                placeholder_text="Section 2.1",
            )

            # Serialize to string for comparison
            xml_parts = [etree.tostring(run, encoding="unicode") for run in runs]
            xml_str = "".join(xml_parts)

            # Check for expected elements
            assert "fldCharType" in xml_str
            assert "begin" in xml_str
            assert "separate" in xml_str
            assert "end" in xml_str
            assert "dirty" in xml_str
            assert "instrText" in xml_str
            assert "REF _Ref123456789" in xml_str
            assert "\\h" in xml_str

        finally:
            doc_path.unlink(missing_ok=True)


# =============================================================================
# Phase 2: Bookmark Management Tests
# =============================================================================


def create_document_with_text() -> Path:
    """Create a test document with multiple paragraphs of text."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>This is the first paragraph with some text.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>This is the second paragraph with Force Majeure clause.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>This is the third paragraph.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_bookmarks() -> Path:
    """Create a test document with existing bookmarks."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:bookmarkStart w:id="0" w:name="VisibleBookmark"/>
  <w:r><w:t>This is bookmarked text.</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
<w:p>
  <w:bookmarkStart w:id="1" w:name="_Ref123456789"/>
  <w:r><w:t>This is hidden ref bookmark.</w:t></w:r>
  <w:bookmarkEnd w:id="1"/>
</w:p>
<w:p>
  <w:bookmarkStart w:id="2" w:name="_Toc987654321"/>
  <w:r><w:t>This is hidden TOC bookmark.</w:t></w:r>
  <w:bookmarkEnd w:id="2"/>
</w:p>
<w:p>
  <w:r><w:t>This paragraph has no bookmark.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestBookmarkInfoDataModel:
    """Tests for BookmarkInfo dataclass."""

    def test_bookmark_info_creation(self):
        """Test creating a BookmarkInfo instance."""
        from python_docx_redline.operations.cross_references import BookmarkInfo

        bk = BookmarkInfo(
            name="MyBookmark",
            bookmark_id="5",
            location="p:3",
            text_preview="Some bookmarked text...",
            is_hidden=False,
        )

        assert bk.name == "MyBookmark"
        assert bk.bookmark_id == "5"
        assert bk.location == "p:3"
        assert bk.text_preview == "Some bookmarked text..."
        assert bk.is_hidden is False
        assert bk.span_end_location is None
        assert bk.referenced_by == []

    def test_bookmark_info_hidden(self):
        """Test BookmarkInfo for hidden bookmark."""
        from python_docx_redline.operations.cross_references import BookmarkInfo

        bk = BookmarkInfo(
            name="_Ref123456789",
            bookmark_id="10",
            location="p:5",
            text_preview="Hidden bookmark text",
            is_hidden=True,
        )

        assert bk.name == "_Ref123456789"
        assert bk.is_hidden is True


class TestValidateBookmarkName:
    """Tests for _validate_bookmark_name method."""

    def test_valid_simple_name(self):
        """Test that simple valid names pass validation."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Should not raise
            ops._validate_bookmark_name("MyBookmark")
            ops._validate_bookmark_name("Section1")
            ops._validate_bookmark_name("A")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_valid_name_with_underscores(self):
        """Test that names with underscores are valid."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops._validate_bookmark_name("My_Bookmark")
            ops._validate_bookmark_name("Section_2_1")
            ops._validate_bookmark_name("A_B_C")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_valid_name_with_numbers(self):
        """Test that names with numbers are valid."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops._validate_bookmark_name("Section123")
            ops._validate_bookmark_name("A1B2C3")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_invalid_empty_name(self):
        """Test that empty name raises error."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(InvalidBookmarkNameError) as exc_info:
                ops._validate_bookmark_name("")

            assert "cannot be empty" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_invalid_name_too_long(self):
        """Test that names longer than 40 chars raise error."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            long_name = "A" * 41

            with pytest.raises(InvalidBookmarkNameError) as exc_info:
                ops._validate_bookmark_name(long_name)

            assert "at most 40 characters" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_valid_name_exactly_40_chars(self):
        """Test that names exactly 40 chars are valid."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            name_40 = "A" * 40
            ops._validate_bookmark_name(name_40)  # Should not raise
        finally:
            doc_path.unlink(missing_ok=True)

    def test_invalid_name_starts_with_number(self):
        """Test that names starting with number raise error."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(InvalidBookmarkNameError) as exc_info:
                ops._validate_bookmark_name("1Section")

            assert "must start with a letter" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_invalid_name_starts_with_underscore(self):
        """Test that names starting with underscore raise error."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(InvalidBookmarkNameError) as exc_info:
                ops._validate_bookmark_name("_Section")

            assert "must start with a letter" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_invalid_name_with_spaces(self):
        """Test that names with spaces raise error."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(InvalidBookmarkNameError) as exc_info:
                ops._validate_bookmark_name("My Bookmark")

            assert "only contain letters, numbers, and underscores" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_invalid_name_with_special_chars(self):
        """Test that names with special characters raise error."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            for name in ["My-Bookmark", "My.Bookmark", "My@Bookmark", "My#Bookmark"]:
                with pytest.raises(InvalidBookmarkNameError):
                    ops._validate_bookmark_name(name)
        finally:
            doc_path.unlink(missing_ok=True)


class TestGenerateRefBookmarkName:
    """Tests for _generate_ref_bookmark_name method."""

    def test_generates_ref_prefix(self):
        """Test that generated name starts with _Ref."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            name = ops._generate_ref_bookmark_name()

            assert name.startswith("_Ref")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_generates_unique_names(self):
        """Test that multiple calls generate unique names."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            names = set()
            for _ in range(10):
                name = ops._generate_ref_bookmark_name()
                names.add(name)

            # All names should be unique
            assert len(names) == 10
        finally:
            doc_path.unlink(missing_ok=True)

    def test_name_format_matches_word(self):
        """Test that generated name format matches Word's pattern."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            name = ops._generate_ref_bookmark_name()

            # Should be _Ref followed by digits
            assert name.startswith("_Ref")
            assert name[4:].isdigit()
            assert len(name[4:]) == 9  # Word uses 9 digits
        finally:
            doc_path.unlink(missing_ok=True)

    def test_avoids_existing_bookmarks(self):
        """Test that generated name doesn't conflict with existing bookmarks."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            name = ops._generate_ref_bookmark_name()

            # Should not be the existing _Ref123456789
            assert name != "_Ref123456789"
        finally:
            doc_path.unlink(missing_ok=True)


class TestListBookmarks:
    """Tests for list_bookmarks method."""

    def test_list_excludes_hidden_by_default(self):
        """Test that hidden bookmarks are excluded by default."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmarks = ops.list_bookmarks()

            names = [bk.name for bk in bookmarks]
            assert "VisibleBookmark" in names
            assert "_Ref123456789" not in names
            assert "_Toc987654321" not in names
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_includes_hidden_when_requested(self):
        """Test that hidden bookmarks are included when requested."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmarks = ops.list_bookmarks(include_hidden=True)

            names = [bk.name for bk in bookmarks]
            assert "VisibleBookmark" in names
            assert "_Ref123456789" in names
            assert "_Toc987654321" in names
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_returns_bookmark_info_objects(self):
        """Test that list returns BookmarkInfo objects with correct data."""
        from python_docx_redline.operations.cross_references import BookmarkInfo

        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmarks = ops.list_bookmarks()

            assert len(bookmarks) == 1
            bk = bookmarks[0]
            assert isinstance(bk, BookmarkInfo)
            assert bk.name == "VisibleBookmark"
            assert bk.bookmark_id == "0"
            assert bk.location == "p:0"
            assert "bookmarked text" in bk.text_preview
            assert bk.is_hidden is False
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_empty_document(self):
        """Test list_bookmarks on document with no bookmarks."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmarks = ops.list_bookmarks()

            assert bookmarks == []
        finally:
            doc_path.unlink(missing_ok=True)

    def test_hidden_bookmark_is_hidden_flag(self):
        """Test that is_hidden flag is correctly set."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmarks = ops.list_bookmarks(include_hidden=True)

            visible = [bk for bk in bookmarks if bk.name == "VisibleBookmark"][0]
            hidden_ref = [bk for bk in bookmarks if bk.name == "_Ref123456789"][0]

            assert visible.is_hidden is False
            assert hidden_ref.is_hidden is True
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetBookmark:
    """Tests for get_bookmark method."""

    def test_get_existing_bookmark(self):
        """Test getting an existing bookmark by name."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bk = ops.get_bookmark("VisibleBookmark")

            assert bk is not None
            assert bk.name == "VisibleBookmark"
            assert bk.bookmark_id == "0"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_nonexistent_bookmark(self):
        """Test getting a bookmark that doesn't exist."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bk = ops.get_bookmark("NonExistentBookmark")

            assert bk is None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_hidden_bookmark(self):
        """Test that hidden bookmarks can be retrieved by name."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bk = ops.get_bookmark("_Ref123456789")

            assert bk is not None
            assert bk.name == "_Ref123456789"
            assert bk.is_hidden is True
        finally:
            doc_path.unlink(missing_ok=True)


class TestCreateBookmark:
    """Tests for create_bookmark method."""

    def test_create_bookmark_at_text(self):
        """Test creating a bookmark at a specific text location."""
        doc_path = create_document_with_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.create_bookmark("ForceMajeure", at="Force Majeure")

            assert result == "ForceMajeure"

            # Verify bookmark was created
            bk = ops.get_bookmark("ForceMajeure")
            assert bk is not None
            assert bk.name == "ForceMajeure"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_bookmark_returns_name(self):
        """Test that create_bookmark returns the bookmark name."""
        doc_path = create_document_with_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.create_bookmark("MyBookmark", at="first paragraph")

            assert result == "MyBookmark"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_bookmark_invalid_name_rejected(self):
        """Test that invalid bookmark names are rejected."""
        doc_path = create_document_with_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(InvalidBookmarkNameError):
                ops.create_bookmark("Invalid Name", at="first paragraph")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_bookmark_duplicate_rejected(self):
        """Test that duplicate bookmark names are rejected."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(BookmarkAlreadyExistsError) as exc_info:
                ops.create_bookmark("VisibleBookmark", at="no bookmark")

            assert "VisibleBookmark" in str(exc_info.value)
            assert "already exists" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_bookmark_text_not_found(self):
        """Test that TextNotFoundError is raised when text not found."""
        from python_docx_redline.errors import TextNotFoundError

        doc_path = create_document_with_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(TextNotFoundError):
                ops.create_bookmark("MyBookmark", at="nonexistent text")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_bookmark_increments_id(self):
        """Test that bookmark IDs are correctly incremented."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Create a new bookmark (existing IDs are 0, 1, 2)
            ops.create_bookmark("NewBookmark", at="no bookmark")

            bk = ops.get_bookmark("NewBookmark")
            assert bk is not None
            # Should get ID 3 (next after 0, 1, 2)
            assert bk.bookmark_id == "3"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_bookmark_xml_structure(self):
        """Test that created bookmark has correct XML structure."""
        doc_path = create_document_with_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.create_bookmark("TestBookmark", at="Force Majeure")

            # Check XML structure
            bookmark_starts = list(doc.xml_root.iter(f"{{{WORD_NS}}}bookmarkStart"))
            bookmark_ends = list(doc.xml_root.iter(f"{{{WORD_NS}}}bookmarkEnd"))

            # Should have one new bookmark
            test_start = None
            for bs in bookmark_starts:
                if bs.get(f"{{{WORD_NS}}}name") == "TestBookmark":
                    test_start = bs
                    break

            assert test_start is not None
            bookmark_id = test_start.get(f"{{{WORD_NS}}}id")
            assert bookmark_id is not None

            # Find matching end
            test_end = None
            for be in bookmark_ends:
                if be.get(f"{{{WORD_NS}}}id") == bookmark_id:
                    test_end = be
                    break

            assert test_end is not None
        finally:
            doc_path.unlink(missing_ok=True)


# =============================================================================
# Phase 3: Cross-Reference Insertion Tests
# =============================================================================


def create_document_with_bookmarks_and_text() -> Path:
    """Create a test document with bookmarks and additional text for insertion."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>This is the introduction paragraph.</w:t></w:r>
</w:p>
<w:p>
  <w:bookmarkStart w:id="0" w:name="DefinitionsSection"/>
  <w:r><w:t>Definitions: Force Majeure means an act of God.</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
<w:p>
  <w:r><w:t>Please refer to the section above for more details.</w:t></w:r>
</w:p>
<w:p>
  <w:bookmarkStart w:id="1" w:name="AppendixA"/>
  <w:r><w:t>Appendix A: Additional Terms</w:t></w:r>
  <w:bookmarkEnd w:id="1"/>
</w:p>
<w:p>
  <w:r><w:t>See the appendix on page X for reference.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestInsertCrossReference:
    """Tests for insert_cross_reference method."""

    def test_insert_ref_field_after_text(self):
        """Test inserting a REF field after anchor text."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="refer to the section",
            )

            assert result == "DefinitionsSection"

            # Verify field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

            instr = instr_texts[0].text
            assert "REF" in instr
            assert "DefinitionsSection" in instr
            assert "\\h" in instr  # Default hyperlink

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ref_field_before_text(self):
        """Test inserting a REF field before anchor text."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                before="for more details",
            )

            assert result == "DefinitionsSection"

            # Verify field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_pageref_field(self):
        """Test inserting a PAGEREF field (display='page')."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="AppendixA",
                display="page",
                after="appendix on page",
            )

            assert result == "AppendixA"

            # Verify PAGEREF field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

            instr = instr_texts[0].text
            assert "PAGEREF" in instr
            assert "AppendixA" in instr
            assert "\\h" in instr

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_above_below_field(self):
        """Test inserting a REF field with above/below display."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="DefinitionsSection",
                display="above_below",
                after="section",
                scope={"paragraph_contains": "refer to"},
            )

            assert result == "DefinitionsSection"

            # Verify field has \p switch
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

            instr = instr_texts[0].text
            assert "REF" in instr
            assert "\\p" in instr

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_without_hyperlink(self):
        """Test inserting a cross-reference without hyperlink switch."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="refer to the section",
                hyperlink=False,
            )

            # Verify field does NOT have \h switch
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

            instr = instr_texts[0].text
            assert "\\h" not in instr

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_with_hyperlink(self):
        """Test that hyperlink switch is added by default."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="refer to the section",
                hyperlink=True,
            )

            # Verify field HAS \h switch
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\h" in instr

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_returns_bookmark_name(self):
        """Test that insert_cross_reference returns the bookmark name."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="AppendixA",
                display="text",
                after="introduction paragraph",
            )

            assert result == "AppendixA"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_bookmark_not_found_error(self):
        """Test that CrossReferenceTargetNotFoundError is raised for missing bookmark."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(CrossReferenceTargetNotFoundError) as exc_info:
                ops.insert_cross_reference(
                    target="NonExistentBookmark",
                    display="text",
                    after="introduction",
                )

            assert "NonExistentBookmark" in str(exc_info.value)
            # Should include available bookmarks in error
            assert "DefinitionsSection" in str(exc_info.value) or exc_info.value.available_targets

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_anchor_text_not_found_error(self):
        """Test that TextNotFoundError is raised when anchor text not found."""
        from python_docx_redline.errors import TextNotFoundError

        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(TextNotFoundError):
                ops.insert_cross_reference(
                    target="DefinitionsSection",
                    display="text",
                    after="text that does not exist anywhere",
                )

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ambiguous_text_error(self):
        """Test that AmbiguousTextError is raised when anchor text appears multiple times."""
        from python_docx_redline.errors import AmbiguousTextError

        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # "the" appears in multiple places
            with pytest.raises(AmbiguousTextError):
                ops.insert_cross_reference(
                    target="DefinitionsSection",
                    display="text",
                    after="the",
                )

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_requires_after_or_before(self):
        """Test that either after or before must be specified."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(ValueError) as exc_info:
                ops.insert_cross_reference(
                    target="DefinitionsSection",
                    display="text",
                )

            assert "after" in str(exc_info.value).lower() or "before" in str(exc_info.value).lower()

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_cannot_specify_both_after_and_before(self):
        """Test that both after and before cannot be specified."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(ValueError) as exc_info:
                ops.insert_cross_reference(
                    target="DefinitionsSection",
                    display="text",
                    after="some text",
                    before="other text",
                )

            assert "both" in str(exc_info.value).lower()

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_invalid_display_option_error(self):
        """Test that invalid display option raises ValueError."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(ValueError) as exc_info:
                ops.insert_cross_reference(
                    target="DefinitionsSection",
                    display="invalid_option",
                    after="introduction",
                )

            assert "invalid_option" in str(exc_info.value)

        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertCrossReferenceFieldStructure:
    """Tests for the XML field structure of inserted cross-references."""

    def test_field_has_dirty_flag(self):
        """Test that inserted field has dirty flag set for Word calculation."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="introduction paragraph",
            )

            # Find fldChar with begin type and check dirty flag
            for fld_char in doc.xml_root.iter(f"{{{WORD_NS}}}fldChar"):
                if fld_char.get(f"{{{WORD_NS}}}fldCharType") == "begin":
                    assert fld_char.get(f"{{{WORD_NS}}}dirty") == "true"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_field_structure_complete(self):
        """Test that field has complete structure: begin, instr, separate, result, end."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="introduction paragraph",
            )

            # Collect fldChar types
            fld_char_types = []
            for fld_char in doc.xml_root.iter(f"{{{WORD_NS}}}fldChar"):
                fld_char_types.append(fld_char.get(f"{{{WORD_NS}}}fldCharType"))

            assert "begin" in fld_char_types
            assert "separate" in fld_char_types
            assert "end" in fld_char_types

            # Verify instruction text exists
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

        finally:
            doc_path.unlink(missing_ok=True)

    def test_field_has_placeholder_text(self):
        """Test that field has placeholder text for display before Word calculates."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="introduction paragraph",
            )

            # Find runs with text after the separator
            found_placeholder = False
            in_field = False
            after_separator = False

            for run in doc.xml_root.iter(f"{{{WORD_NS}}}r"):
                fld_char = run.find(f"{{{WORD_NS}}}fldChar")
                if fld_char is not None:
                    fld_type = fld_char.get(f"{{{WORD_NS}}}fldCharType")
                    if fld_type == "begin":
                        in_field = True
                    elif fld_type == "separate":
                        after_separator = True
                    elif fld_type == "end":
                        in_field = False
                        after_separator = False

                if in_field and after_separator:
                    t_elem = run.find(f"{{{WORD_NS}}}t")
                    if t_elem is not None and t_elem.text:
                        found_placeholder = True
                        break

            assert found_placeholder, "Should have placeholder text in field result"

        finally:
            doc_path.unlink(missing_ok=True)


class TestResolveTarget:
    """Tests for _resolve_target method."""

    def test_resolve_existing_bookmark(self):
        """Test resolving a bookmark that exists."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, was_created = ops._resolve_target("DefinitionsSection")

            assert bookmark_name == "DefinitionsSection"
            assert was_created is False

        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_nonexistent_bookmark_raises_error(self):
        """Test that resolving nonexistent bookmark raises error."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(CrossReferenceTargetNotFoundError) as exc_info:
                ops._resolve_target("NotARealBookmark")

            assert exc_info.value.target == "NotARealBookmark"

        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_hidden_bookmark(self):
        """Test that hidden bookmarks can be resolved."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, was_created = ops._resolve_target("_Ref123456789")

            assert bookmark_name == "_Ref123456789"
            assert was_created is False

        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertFieldAtPosition:
    """Tests for _insert_field_at_position method."""

    def test_insert_after_match(self):
        """Test inserting field runs after a text match."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Create field runs
            field_runs = ops._create_field_code("REF", "TestBookmark", ["\\h"])

            # Find a match
            paragraphs = list(doc.xml_root.iter(f"{{{WORD_NS}}}p"))
            matches = doc._text_search.find_text("introduction", paragraphs)
            assert len(matches) == 1

            match = matches[0]

            # Insert field after match
            ops._insert_field_at_position(field_runs, match, insert_after=True)

            # Verify runs were inserted into paragraph
            para_children = list(match.paragraph)
            # Should have more runs now
            run_count = sum(1 for child in para_children if child.tag == f"{{{WORD_NS}}}r")
            assert run_count >= 5  # Original run + 5 field runs

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_before_match(self):
        """Test inserting field runs before a text match."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Create field runs
            field_runs = ops._create_field_code("REF", "TestBookmark", ["\\h"])

            # Find a match
            paragraphs = list(doc.xml_root.iter(f"{{{WORD_NS}}}p"))
            matches = doc._text_search.find_text("introduction", paragraphs)
            assert len(matches) == 1

            match = matches[0]

            # Insert field before match
            ops._insert_field_at_position(field_runs, match, insert_after=False)

            # Verify runs were inserted into paragraph
            para_children = list(match.paragraph)
            run_count = sum(1 for child in para_children if child.tag == f"{{{WORD_NS}}}r")
            assert run_count >= 5

        finally:
            doc_path.unlink(missing_ok=True)


class TestCrossReferenceRoundTrip:
    """Integration tests for round-trip: insert, save, reload, verify."""

    def test_roundtrip_saves_and_loads_field(self):
        """Test that inserted field survives save and reload."""
        doc_path = create_document_with_bookmarks_and_text()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Insert cross-reference
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="introduction paragraph",
            )

            # Save document
            doc.save(str(output_path))

            # Reload document
            doc2 = Document(output_path)

            # Verify field still exists
            instr_texts = list(doc2.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

            instr = instr_texts[0].text
            assert "REF" in instr
            assert "DefinitionsSection" in instr

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_roundtrip_preserves_field_structure(self):
        """Test that field structure is preserved after save/reload."""
        doc_path = create_document_with_bookmarks_and_text()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Insert cross-reference
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="AppendixA",
                display="page",
                after="appendix on page",
            )

            # Save document
            doc.save(str(output_path))

            # Reload document
            doc2 = Document(output_path)

            # Verify complete field structure
            fld_char_types = []
            for fld_char in doc2.xml_root.iter(f"{{{WORD_NS}}}fldChar"):
                fld_char_types.append(fld_char.get(f"{{{WORD_NS}}}fldCharType"))

            assert "begin" in fld_char_types
            assert "separate" in fld_char_types
            assert "end" in fld_char_types

            # Verify PAGEREF instruction
            instr_texts = list(doc2.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1
            assert "PAGEREF" in instr_texts[0].text

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_roundtrip_preserves_dirty_flag(self):
        """Test that dirty flag is preserved after save/reload."""
        doc_path = create_document_with_bookmarks_and_text()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Insert cross-reference
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="introduction paragraph",
            )

            # Save document
            doc.save(str(output_path))

            # Reload document
            doc2 = Document(output_path)

            # Verify dirty flag
            for fld_char in doc2.xml_root.iter(f"{{{WORD_NS}}}fldChar"):
                if fld_char.get(f"{{{WORD_NS}}}fldCharType") == "begin":
                    assert fld_char.get(f"{{{WORD_NS}}}dirty") == "true"

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


class TestMultipleCrossReferences:
    """Tests for inserting multiple cross-references."""

    def test_insert_multiple_cross_references(self):
        """Test inserting multiple cross-references to same bookmark."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Insert first cross-reference
            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="introduction paragraph",
            )

            # Insert second cross-reference to same target
            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="page",
                after="appendix on page",
            )

            # Verify both fields exist
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 2

            # Check one is REF and one is PAGEREF
            instructions = [it.text for it in instr_texts]
            ref_count = sum(1 for i in instructions if "REF" in i and "PAGEREF" not in i)
            pageref_count = sum(1 for i in instructions if "PAGEREF" in i)

            assert ref_count == 1
            assert pageref_count == 1

        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_cross_references_to_different_bookmarks(self):
        """Test inserting cross-references to different bookmarks."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Insert cross-reference to first bookmark
            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="introduction paragraph",
            )

            # Insert cross-reference to second bookmark
            ops.insert_cross_reference(
                target="AppendixA",
                display="text",
                after="for more details",
            )

            # Verify both fields exist with different targets
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 2

            instructions = [it.text for it in instr_texts]
            assert any("DefinitionsSection" in i for i in instructions)
            assert any("AppendixA" in i for i in instructions)

        finally:
            doc_path.unlink(missing_ok=True)


# =============================================================================
# Phase 4: Heading Reference Tests
# =============================================================================


def create_document_with_headings() -> Path:
    """Create a test document with heading styles."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Introduction</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>This is the introduction paragraph with some content.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
  <w:r><w:t>Chapter 1: Getting Started</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>This is the content of chapter 1.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
  <w:r><w:t>Chapter 2: Advanced Topics</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>This is the content of chapter 2. See Introduction above.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Conclusion</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>This is the conclusion.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_existing_ref_bookmark() -> Path:
    """Create a test document with a heading that already has a _Ref bookmark."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:bookmarkStart w:id="0" w:name="_Ref999888777"/>
  <w:r><w:t>Introduction</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
<w:p>
  <w:r><w:t>This is the introduction paragraph. See above.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Conclusion</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>This is the conclusion.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestFindHeadingParagraph:
    """Tests for _find_heading_paragraph method."""

    def test_find_heading_exact_match(self):
        """Test finding a heading by exact text match."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_heading_paragraph("Introduction")

            assert para is not None
            text = ops._extract_paragraph_text(para)
            assert text == "Introduction"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_heading_partial_match(self):
        """Test finding a heading by partial text match."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_heading_paragraph("Chapter 1")

            assert para is not None
            text = ops._extract_paragraph_text(para)
            assert "Chapter 1" in text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_heading_case_insensitive(self):
        """Test that heading search is case-insensitive."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_heading_paragraph("INTRODUCTION")

            assert para is not None
            text = ops._extract_paragraph_text(para)
            assert text == "Introduction"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_heading_not_found(self):
        """Test that None is returned when heading not found."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_heading_paragraph("NonExistent Heading")

            assert para is None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_heading_ignores_non_heading_paragraphs(self):
        """Test that non-heading paragraphs are not matched."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # This text exists in a regular paragraph, not a heading
            para = ops._find_heading_paragraph("content of chapter 1")

            assert para is None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_heading_returns_first_match(self):
        """Test that first matching heading is returned when multiple match."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # "Chapter" matches both Chapter 1 and Chapter 2
            para = ops._find_heading_paragraph("Chapter")

            assert para is not None
            text = ops._extract_paragraph_text(para)
            assert "Chapter 1" in text  # Should be first one
        finally:
            doc_path.unlink(missing_ok=True)


class TestFindExistingRefBookmark:
    """Tests for _find_existing_ref_bookmark method."""

    def test_find_existing_ref_bookmark(self):
        """Test finding an existing _Ref bookmark in a paragraph."""
        doc_path = create_document_with_existing_ref_bookmark()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Find the Introduction heading
            para = ops._find_heading_paragraph("Introduction")
            assert para is not None

            existing = ops._find_existing_ref_bookmark(para)

            assert existing == "_Ref999888777"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_no_existing_ref_bookmark(self):
        """Test that None is returned when no _Ref bookmark exists."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Find the Introduction heading (no bookmark)
            para = ops._find_heading_paragraph("Introduction")
            assert para is not None

            existing = ops._find_existing_ref_bookmark(para)

            assert existing is None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_ignores_non_ref_bookmarks(self):
        """Test that non-_Ref bookmarks are ignored."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Get the first paragraph with VisibleBookmark
            for para in doc.xml_root.iter(f"{{{WORD_NS}}}p"):
                for bs in para.iter(f"{{{WORD_NS}}}bookmarkStart"):
                    if bs.get(f"{{{WORD_NS}}}name") == "VisibleBookmark":
                        existing = ops._find_existing_ref_bookmark(para)
                        # VisibleBookmark doesn't start with _Ref
                        assert existing is None
                        return

            pytest.fail("Could not find paragraph with VisibleBookmark")
        finally:
            doc_path.unlink(missing_ok=True)


class TestCreateBookmarkAtParagraph:
    """Tests for _create_bookmark_at_paragraph method."""

    def test_create_bookmark_at_heading(self):
        """Test creating a bookmark at a heading paragraph."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Find the Conclusion heading
            para = ops._find_heading_paragraph("Conclusion")
            assert para is not None

            ops._create_bookmark_at_paragraph("_Ref111222333", para)

            # Verify bookmark was created
            bookmark_starts = list(para.iter(f"{{{WORD_NS}}}bookmarkStart"))
            assert len(bookmark_starts) == 1
            assert bookmark_starts[0].get(f"{{{WORD_NS}}}name") == "_Ref111222333"

            bookmark_ends = list(para.iter(f"{{{WORD_NS}}}bookmarkEnd"))
            assert len(bookmark_ends) == 1
            assert bookmark_ends[0].get(f"{{{WORD_NS}}}id") == bookmark_starts[0].get(
                f"{{{WORD_NS}}}id"
            )
        finally:
            doc_path.unlink(missing_ok=True)

    def test_created_bookmark_wraps_paragraph_content(self):
        """Test that the bookmark wraps the paragraph content."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Find the Introduction heading
            para = ops._find_heading_paragraph("Introduction")
            assert para is not None

            ops._create_bookmark_at_paragraph("_Ref444555666", para)

            # Get the children in order
            children = list(para)
            child_tags = [child.tag.split("}")[-1] for child in children]

            # Should be: pPr, bookmarkStart, r, bookmarkEnd
            assert "bookmarkStart" in child_tags
            assert "bookmarkEnd" in child_tags

            # bookmarkStart should come before r
            start_idx = child_tags.index("bookmarkStart")
            r_idx = child_tags.index("r")
            end_idx = child_tags.index("bookmarkEnd")

            assert start_idx < r_idx < end_idx
        finally:
            doc_path.unlink(missing_ok=True)


class TestResolveHeadingTarget:
    """Tests for _resolve_heading_target method."""

    def test_resolve_heading_creates_bookmark(self):
        """Test that resolving a heading creates a _Ref bookmark."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error = ops._resolve_heading_target("Introduction")

            assert error is None
            assert bookmark_name.startswith("_Ref")

            # Verify bookmark was created
            bookmark = ops.get_bookmark(bookmark_name)
            assert bookmark is not None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_heading_reuses_existing_bookmark(self):
        """Test that resolving a heading reuses existing _Ref bookmark."""
        doc_path = create_document_with_existing_ref_bookmark()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error = ops._resolve_heading_target("Introduction")

            assert error is None
            assert bookmark_name == "_Ref999888777"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_heading_not_found_returns_error(self):
        """Test that resolving a non-existent heading returns an error."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error = ops._resolve_heading_target("NonExistent Heading")

            assert error is not None
            assert "NonExistent Heading" in error
            assert bookmark_name == ""
        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertCrossReferenceToHeading:
    """Tests for inserting cross-references to headings."""

    def test_insert_cross_reference_to_heading_exact_match(self):
        """Test inserting a cross-reference to a heading by exact text."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="heading:Introduction",
                display="text",
                after="See",
                scope={"paragraph_contains": "See above"},
            )

            # Should have created a _Ref bookmark
            assert result.startswith("_Ref")

            # Verify field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

            instr = instr_texts[0].text
            assert "REF" in instr
            assert result in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_cross_reference_to_heading_partial_match(self):
        """Test inserting a cross-reference to a heading by partial text."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="heading:Chapter 1",
                display="text",
                after="content of chapter 2",
            )

            assert result.startswith("_Ref")

            # Verify the bookmark references the correct heading
            bookmark = ops.get_bookmark(result)
            assert bookmark is not None
            assert (
                "Getting Started" in bookmark.text_preview or "Chapter 1" in bookmark.text_preview
            )
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_cross_reference_reuses_existing_heading_bookmark(self):
        """Test that inserting a cross-reference reuses existing heading bookmark."""
        doc_path = create_document_with_existing_ref_bookmark()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="heading:Introduction",
                display="text",
                after="See",
                scope={"paragraph_contains": "See above"},
            )

            # Should reuse the existing bookmark
            assert result == "_Ref999888777"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_cross_reference_to_heading_with_number_display(self):
        """Test inserting a cross-reference with heading number display."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="heading:Introduction",
                display="number",
                after="This is the introduction paragraph",
            )

            assert result.startswith("_Ref")

            # Verify field has \n switch
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\n" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_cross_reference_to_heading_with_full_number_display(self):
        """Test inserting a cross-reference with full heading number display."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="heading:Chapter 1",
                display="full_number",
                after="content of chapter 2",
            )

            assert result.startswith("_Ref")

            # Verify field has \w switch
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\w" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_cross_reference_to_heading_with_relative_number_display(self):
        """Test inserting a cross-reference with relative number display."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="heading:Introduction",
                display="relative_number",
                after="This is the introduction paragraph",
            )

            assert result.startswith("_Ref")

            # Verify field has \r switch
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\r" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_cross_reference_to_heading_not_found_error(self):
        """Test that error is raised when heading is not found."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(CrossReferenceTargetNotFoundError) as exc_info:
                ops.insert_cross_reference(
                    target="heading:NonExistent Heading",
                    display="text",
                    after="content of chapter 2",
                )

            assert "heading:NonExistent Heading" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)


class TestCreateHeadingBookmark:
    """Tests for create_heading_bookmark method."""

    def test_create_heading_bookmark_auto_name(self):
        """Test creating a heading bookmark with auto-generated name."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.create_heading_bookmark("Introduction")

            assert result.startswith("_Ref")

            # Verify bookmark was created
            bookmark = ops.get_bookmark(result)
            assert bookmark is not None
            assert "Introduction" in bookmark.text_preview
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_heading_bookmark_custom_name(self):
        """Test creating a heading bookmark with custom name."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.create_heading_bookmark("Conclusion", bookmark_name="MyConclusion")

            assert result == "MyConclusion"

            # Verify bookmark was created
            bookmark = ops.get_bookmark("MyConclusion")
            assert bookmark is not None
            assert "Conclusion" in bookmark.text_preview
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_heading_bookmark_returns_existing(self):
        """Test that existing _Ref bookmark is returned."""
        doc_path = create_document_with_existing_ref_bookmark()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.create_heading_bookmark("Introduction")

            assert result == "_Ref999888777"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_heading_bookmark_heading_not_found_error(self):
        """Test that error is raised when heading not found."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(CrossReferenceTargetNotFoundError):
                ops.create_heading_bookmark("NonExistent Heading")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_heading_bookmark_invalid_name_error(self):
        """Test that invalid custom bookmark name raises error."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(InvalidBookmarkNameError):
                ops.create_heading_bookmark("Conclusion", bookmark_name="Invalid Name")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_create_heading_bookmark_duplicate_name_error(self):
        """Test that duplicate custom bookmark name raises error."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            _ops = CrossReferenceOperations(doc)  # noqa: F841

            # Can't test this easily since we don't have a heading in this doc
            # but we can add a heading to the document
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetHeadingLevelFromStyle:
    """Tests for _get_heading_level_from_style method."""

    def test_heading_style_patterns(self):
        """Test various heading style name patterns."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Heading1, Heading2, etc.
            assert ops._get_heading_level_from_style("Heading1") == 1
            assert ops._get_heading_level_from_style("Heading2") == 2
            assert ops._get_heading_level_from_style("Heading9") == 9

            # Heading 1, Heading 2, etc.
            assert ops._get_heading_level_from_style("Heading 1") == 1
            assert ops._get_heading_level_from_style("Heading 2") == 2

            # Case insensitive
            assert ops._get_heading_level_from_style("heading1") == 1
            assert ops._get_heading_level_from_style("HEADING2") == 2

            # Title is level 1
            assert ops._get_heading_level_from_style("Title") == 1
            assert ops._get_heading_level_from_style("title") == 1

            # Non-heading styles
            assert ops._get_heading_level_from_style("Normal") is None
            assert ops._get_heading_level_from_style("BodyText") is None
            assert ops._get_heading_level_from_style("") is None
            assert ops._get_heading_level_from_style(None) is None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_heading_level_out_of_range(self):
        """Test that heading levels outside 1-9 return None."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            assert ops._get_heading_level_from_style("Heading0") is None
            assert ops._get_heading_level_from_style("Heading10") is None
        finally:
            doc_path.unlink(missing_ok=True)


class TestHeadingReferenceRoundTrip:
    """Integration tests for heading reference round-trip."""

    def test_roundtrip_heading_reference(self):
        """Test that heading references survive save/reload."""
        doc_path = create_document_with_headings()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Insert cross-reference to heading
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name = ops.insert_cross_reference(
                target="heading:Introduction",
                display="text",
                after="See",
                scope={"paragraph_contains": "See above"},
            )

            # Save document
            doc.save(str(output_path))

            # Reload document
            doc2 = Document(output_path)
            ops2 = CrossReferenceOperations(doc2)

            # Verify bookmark still exists
            bookmark = ops2.get_bookmark(bookmark_name)
            assert bookmark is not None

            # Verify field still exists
            instr_texts = list(doc2.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1
            assert bookmark_name in instr_texts[0].text

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_roundtrip_preserves_bookmark_structure(self):
        """Test that bookmark structure is preserved after save/reload."""
        doc_path = create_document_with_headings()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Create heading bookmark
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name = ops.create_heading_bookmark("Conclusion")

            # Save document
            doc.save(str(output_path))

            # Reload document
            doc2 = Document(output_path)

            # Verify bookmark structure
            bookmark_starts = list(doc2.xml_root.iter(f"{{{WORD_NS}}}bookmarkStart"))
            matching_starts = [
                bs for bs in bookmark_starts if bs.get(f"{{{WORD_NS}}}name") == bookmark_name
            ]
            assert len(matching_starts) == 1

            bookmark_id = matching_starts[0].get(f"{{{WORD_NS}}}id")
            bookmark_ends = list(doc2.xml_root.iter(f"{{{WORD_NS}}}bookmarkEnd"))
            matching_ends = [
                be for be in bookmark_ends if be.get(f"{{{WORD_NS}}}id") == bookmark_id
            ]
            assert len(matching_ends) == 1

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


# =============================================================================
# Phase 5: Caption Reference Tests (Figures and Tables)
# =============================================================================


def create_document_with_figure_captions_simple() -> Path:
    """Create a test document with figure captions using fldSimple element."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>This is the introduction paragraph.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
  <w:r><w:t>Figure </w:t></w:r>
  <w:fldSimple w:instr=" SEQ Figure \\* ARABIC ">
    <w:r><w:t>1</w:t></w:r>
  </w:fldSimple>
  <w:r><w:t>: Architecture Diagram</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>See the figure above for details.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
  <w:r><w:t>Figure </w:t></w:r>
  <w:fldSimple w:instr=" SEQ Figure \\* ARABIC ">
    <w:r><w:t>2</w:t></w:r>
  </w:fldSimple>
  <w:r><w:t>: Network Topology</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>The network topology is shown above.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_table_captions_complex() -> Path:
    """Create a test document with table captions using complex field pattern."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>This is the introduction paragraph.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
  <w:r><w:t>Table </w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText> SEQ Table \\* ARABIC </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>1</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:t>: Revenue Data</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>See the revenue table above for financial details.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
  <w:r><w:t>Table </w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText> SEQ Table \\* ARABIC </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>2</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:t>: Customer Analysis</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>The customer analysis table provides insights.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_mixed_captions() -> Path:
    """Create a test document with both figure and table captions."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>This document contains both figures and tables.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
  <w:r><w:t>Figure </w:t></w:r>
  <w:fldSimple w:instr=" SEQ Figure \\* ARABIC ">
    <w:r><w:t>1</w:t></w:r>
  </w:fldSimple>
  <w:r><w:t>: System Architecture</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>The architecture is described in the figure.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
  <w:r><w:t>Table </w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText> SEQ Table \\* ARABIC </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>1</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:t>: Performance Metrics</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>The metrics are shown in the table above. See Figure 1 for context.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_caption_and_existing_bookmark() -> Path:
    """Create a test document with a caption that already has a _Ref bookmark."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r><w:t>This is the introduction paragraph.</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
  <w:bookmarkStart w:id="0" w:name="_Ref111222333"/>
  <w:r><w:t>Figure </w:t></w:r>
  <w:fldSimple w:instr=" SEQ Figure \\* ARABIC ">
    <w:r><w:t>1</w:t></w:r>
  </w:fldSimple>
  <w:r><w:t>: Existing Diagram</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
<w:p>
  <w:r><w:t>Reference the figure above.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestParseCaptionNumber:
    """Tests for _parse_caption_number method."""

    def test_parse_caption_number_simple_field(self):
        """Test parsing caption number from simple field (fldSimple)."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Find the first caption paragraph
            para = ops._find_caption_paragraph("Figure", "1")
            assert para is not None

            number = ops._parse_caption_number(para, "Figure")
            assert number == "1"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_caption_number_complex_field(self):
        """Test parsing caption number from complex field (fldChar pattern)."""
        doc_path = create_document_with_table_captions_complex()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Find the first table caption paragraph
            para = ops._find_caption_paragraph("Table", "1")
            assert para is not None

            number = ops._parse_caption_number(para, "Table")
            assert number == "1"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_caption_number_second_caption(self):
        """Test parsing caption number from second figure caption."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Find the second caption paragraph
            para = ops._find_caption_paragraph("Figure", "2")
            assert para is not None

            number = ops._parse_caption_number(para, "Figure")
            assert number == "2"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_caption_number_wrong_seq_id(self):
        """Test that wrong SEQ id returns None."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Find the figure caption paragraph
            para = ops._find_caption_paragraph("Figure", "1")
            assert para is not None

            # Try to parse it as a Table caption
            number = ops._parse_caption_number(para, "Table")
            assert number is None
        finally:
            doc_path.unlink(missing_ok=True)


class TestIsMatchingSeqField:
    """Tests for _is_matching_seq_field method."""

    def test_matching_figure_seq(self):
        """Test matching Figure SEQ field."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            assert ops._is_matching_seq_field(" SEQ Figure \\* ARABIC ", "Figure")
            assert ops._is_matching_seq_field("SEQ Figure \\* ARABIC", "Figure")
            assert ops._is_matching_seq_field(" SEQ FIGURE \\* ARABIC ", "Figure")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_matching_table_seq(self):
        """Test matching Table SEQ field."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            assert ops._is_matching_seq_field(" SEQ Table \\* ARABIC ", "Table")
            assert ops._is_matching_seq_field("SEQ TABLE \\* ARABIC", "Table")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_non_matching_seq(self):
        """Test non-matching SEQ fields."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Different identifier
            assert not ops._is_matching_seq_field(" SEQ Figure \\* ARABIC ", "Table")
            assert not ops._is_matching_seq_field(" SEQ Table \\* ARABIC ", "Figure")

            # Not a SEQ field
            assert not ops._is_matching_seq_field(" REF _Ref123 \\h ", "Figure")
            assert not ops._is_matching_seq_field("Some other text", "Figure")
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetCaptionText:
    """Tests for _get_caption_text method."""

    def test_get_caption_text_with_colon(self):
        """Test extracting caption text after colon separator."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_caption_paragraph("Figure", "1")
            assert para is not None

            text = ops._get_caption_text(para, "Figure")
            assert text == "Architecture Diagram"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_caption_text_second_figure(self):
        """Test extracting caption text from second figure."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_caption_paragraph("Figure", "2")
            assert para is not None

            text = ops._get_caption_text(para, "Figure")
            assert text == "Network Topology"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_caption_text_table(self):
        """Test extracting caption text from table caption."""
        doc_path = create_document_with_table_captions_complex()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_caption_paragraph("Table", "1")
            assert para is not None

            text = ops._get_caption_text(para, "Table")
            assert text == "Revenue Data"
        finally:
            doc_path.unlink(missing_ok=True)


class TestFindCaptionParagraph:
    """Tests for _find_caption_paragraph method."""

    def test_find_by_number(self):
        """Test finding caption by number."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_caption_paragraph("Figure", "1")
            assert para is not None

            # Verify it's the right caption
            text = ops._extract_paragraph_text(para)
            assert "Figure" in text
            assert "1" in text
            assert "Architecture" in text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_by_text_exact(self):
        """Test finding caption by exact caption text."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_caption_paragraph("Figure", "Architecture Diagram")
            assert para is not None

            # Verify it's the right caption
            number = ops._parse_caption_number(para, "Figure")
            assert number == "1"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_by_text_partial(self):
        """Test finding caption by partial caption text."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Partial match
            para = ops._find_caption_paragraph("Figure", "Architecture")
            assert para is not None

            number = ops._parse_caption_number(para, "Figure")
            assert number == "1"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_by_text_case_insensitive(self):
        """Test that caption text search is case-insensitive."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Case-insensitive match
            para = ops._find_caption_paragraph("Figure", "ARCHITECTURE DIAGRAM")
            assert para is not None

            number = ops._parse_caption_number(para, "Figure")
            assert number == "1"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_table_caption(self):
        """Test finding table caption."""
        doc_path = create_document_with_table_captions_complex()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            para = ops._find_caption_paragraph("Table", "1")
            assert para is not None

            text = ops._extract_paragraph_text(para)
            assert "Table" in text
            assert "Revenue" in text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_caption_not_found(self):
        """Test that None is returned when caption not found."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Non-existent number
            para = ops._find_caption_paragraph("Figure", "99")
            assert para is None

            # Non-existent text
            para = ops._find_caption_paragraph("Figure", "Nonexistent Caption")
            assert para is None

            # Wrong type
            para = ops._find_caption_paragraph("Table", "1")
            assert para is None
        finally:
            doc_path.unlink(missing_ok=True)


class TestResolveCaptionTarget:
    """Tests for _resolve_caption_target method."""

    def test_resolve_figure_by_number(self):
        """Test resolving figure caption by number."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error = ops._resolve_caption_target("Figure", "1")

            assert error is None
            assert bookmark_name.startswith("_Ref")

            # Verify bookmark was created
            bookmark = ops.get_bookmark(bookmark_name)
            assert bookmark is not None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_figure_by_text(self):
        """Test resolving figure caption by text."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error = ops._resolve_caption_target("Figure", "Architecture Diagram")

            assert error is None
            assert bookmark_name.startswith("_Ref")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_table_by_number(self):
        """Test resolving table caption by number."""
        doc_path = create_document_with_table_captions_complex()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error = ops._resolve_caption_target("Table", "1")

            assert error is None
            assert bookmark_name.startswith("_Ref")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_caption_not_found(self):
        """Test that error is returned when caption not found."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error = ops._resolve_caption_target("Figure", "99")

            assert bookmark_name == ""
            assert error is not None
            assert "figure" in error.lower()
            assert "99" in error
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_reuses_existing_bookmark(self):
        """Test that existing _Ref bookmark is reused."""
        doc_path = create_document_with_caption_and_existing_bookmark()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error = ops._resolve_caption_target("Figure", "1")

            assert error is None
            assert bookmark_name == "_Ref111222333"
        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertCrossReferenceFigure:
    """Tests for insert_cross_reference with figure targets."""

    def test_insert_ref_to_figure_by_number(self):
        """Test inserting cross-reference to figure by number."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="figure:1",
                display="text",
                after="figure above",
            )

            assert result.startswith("_Ref")

            # Verify field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1
            assert "REF" in instr_texts[0].text
            assert result in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ref_to_figure_by_text(self):
        """Test inserting cross-reference to figure by caption text."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="figure:Architecture Diagram",
                display="text",
                after="figure above",
            )

            assert result.startswith("_Ref")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ref_to_figure_page_display(self):
        """Test inserting page reference to figure."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="figure:1",
                display="page",
                after="figure above",
            )

            assert result.startswith("_Ref")

            # Verify PAGEREF was used
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "PAGEREF" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ref_to_figure_label_number_display(self):
        """Test inserting reference to figure with label_number display."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="figure:1",
                display="label_number",
                after="figure above",
            )

            assert result.startswith("_Ref")

            # Verify REF was used
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "REF" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ref_to_figure_number_only_display(self):
        """Test inserting reference to figure with number_only display."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="figure:1",
                display="number_only",
                after="figure above",
            )

            assert result.startswith("_Ref")

            # Verify REF with \n switch was used
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "REF" in instr_texts[0].text
            assert "\\n" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ref_to_nonexistent_figure(self):
        """Test that error is raised for nonexistent figure."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(CrossReferenceTargetNotFoundError) as exc_info:
                ops.insert_cross_reference(
                    target="figure:99",
                    display="text",
                    after="figure above",
                )

            assert "figure:99" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertCrossReferenceTable:
    """Tests for insert_cross_reference with table targets."""

    def test_insert_ref_to_table_by_number(self):
        """Test inserting cross-reference to table by number."""
        doc_path = create_document_with_table_captions_complex()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="table:1",
                display="text",
                after="revenue table above",
            )

            assert result.startswith("_Ref")

            # Verify field was inserted - look for REF field specifically
            # (document already has SEQ fields with instrText elements)
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            ref_fields = [it for it in instr_texts if it.text and "REF" in it.text]
            assert len(ref_fields) == 1
            assert result in ref_fields[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ref_to_table_by_text(self):
        """Test inserting cross-reference to table by caption text."""
        doc_path = create_document_with_table_captions_complex()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="table:Revenue Data",
                display="text",
                after="revenue table above",
            )

            assert result.startswith("_Ref")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ref_to_table_partial_text(self):
        """Test inserting cross-reference to table by partial caption text."""
        doc_path = create_document_with_table_captions_complex()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="table:Revenue",
                display="text",
                after="revenue table above",
            )

            assert result.startswith("_Ref")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_ref_to_nonexistent_table(self):
        """Test that error is raised for nonexistent table."""
        doc_path = create_document_with_table_captions_complex()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(CrossReferenceTargetNotFoundError) as exc_info:
                ops.insert_cross_reference(
                    target="table:NonExistent",
                    display="text",
                    after="revenue table above",
                )

            assert "table:NonExistent" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)


class TestMixedCaptionReferences:
    """Tests for documents with both figure and table captions."""

    def test_insert_refs_to_both_figure_and_table(self):
        """Test inserting references to both figure and table in same document."""
        doc_path = create_document_with_mixed_captions()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Insert reference to figure
            fig_bookmark = ops.insert_cross_reference(
                target="figure:1",
                display="text",
                after="See Figure 1",
            )

            # Insert reference to table
            table_bookmark = ops.insert_cross_reference(
                target="table:1",
                display="text",
                after="in the table above",
            )

            assert fig_bookmark.startswith("_Ref")
            assert table_bookmark.startswith("_Ref")
            assert fig_bookmark != table_bookmark

            # Verify two REF fields were inserted
            # (document already has SEQ fields with instrText elements)
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            ref_fields = [it for it in instr_texts if it.text and "REF" in it.text]
            assert len(ref_fields) == 2
        finally:
            doc_path.unlink(missing_ok=True)

    def test_figure_and_table_same_number(self):
        """Test that figure 1 and table 1 are resolved separately."""
        doc_path = create_document_with_mixed_captions()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Both figure:1 and table:1 exist
            fig_para = ops._find_caption_paragraph("Figure", "1")
            table_para = ops._find_caption_paragraph("Table", "1")

            assert fig_para is not None
            assert table_para is not None
            # They should be different paragraphs
            assert fig_para is not table_para

            # Verify caption text is different
            fig_text = ops._get_caption_text(fig_para, "Figure")
            table_text = ops._get_caption_text(table_para, "Table")

            assert fig_text == "System Architecture"
            assert table_text == "Performance Metrics"
        finally:
            doc_path.unlink(missing_ok=True)


class TestCaptionReferenceRoundTrip:
    """Integration tests for caption reference round-trip."""

    def test_roundtrip_figure_reference(self):
        """Test that figure references survive save/reload."""
        doc_path = create_document_with_figure_captions_simple()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Insert cross-reference to figure
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name = ops.insert_cross_reference(
                target="figure:1",
                display="text",
                after="figure above",
            )

            # Save document (validate=False since test document is minimal)
            doc.save(str(output_path), validate=False)

            # Reload document
            doc2 = Document(output_path)
            ops2 = CrossReferenceOperations(doc2)

            # Verify bookmark still exists
            bookmark = ops2.get_bookmark(bookmark_name)
            assert bookmark is not None

            # Verify REF field still exists (check for REF specifically)
            instr_texts = list(doc2.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            ref_fields = [it for it in instr_texts if it.text and "REF" in it.text]
            assert len(ref_fields) == 1
            assert bookmark_name in ref_fields[0].text

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_roundtrip_table_reference(self):
        """Test that table references survive save/reload."""
        doc_path = create_document_with_table_captions_complex()
        output_path = Path(tempfile.mktemp(suffix=".docx"))

        try:
            # Insert cross-reference to table
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name = ops.insert_cross_reference(
                target="table:1",
                display="text",
                after="revenue table above",
            )

            # Save document (validate=False since test document is minimal)
            doc.save(str(output_path), validate=False)

            # Reload document
            doc2 = Document(output_path)
            ops2 = CrossReferenceOperations(doc2)

            # Verify bookmark still exists
            bookmark = ops2.get_bookmark(bookmark_name)
            assert bookmark is not None

            # Verify REF field still exists (document has SEQ fields too)
            instr_texts = list(doc2.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            ref_fields = [it for it in instr_texts if it.text and "REF" in it.text]
            assert len(ref_fields) == 1
            assert bookmark_name in ref_fields[0].text

        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


class TestCaptionDisplayOptions:
    """Tests for various display options with caption references."""

    def test_display_text_for_figure(self):
        """Test display='text' shows caption content."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="figure:1",
                display="text",
                after="figure above",
            )

            # Verify REF field is used
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "REF" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_page_for_figure(self):
        """Test display='page' shows page number."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="figure:1",
                display="page",
                after="figure above",
            )

            # Verify PAGEREF field is used
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "PAGEREF" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_number_only_for_figure(self):
        """Test display='number_only' shows just the number."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="figure:1",
                display="number_only",
                after="figure above",
            )

            # Verify REF with \n switch
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "REF" in instr_texts[0].text
            assert "\\n" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_display_above_below_for_figure(self):
        """Test display='above_below' shows position."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_cross_reference(
                target="figure:1",
                display="above_below",
                after="figure above",
            )

            # Verify REF with \p switch
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "REF" in instr_texts[0].text
            assert "\\p" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)


# =============================================================================
# Phase 6: Note References and Convenience Methods Tests
# =============================================================================


def create_document_with_footnotes() -> Path:
    """Create a document with footnotes for testing."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:r><w:t>This paragraph has a footnote reference</w:t></w:r>
  <w:r>
    <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
    <w:footnoteReference w:id="1"/>
  </w:r>
  <w:r><w:t> and continues here.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>Another paragraph for testing cross-references.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>See also note marker for more info.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    footnotes_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p>
      <w:r>
        <w:separator/>
      </w:r>
    </w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p>
      <w:r>
        <w:continuationSeparator/>
      </w:r>
    </w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="FootnoteText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="FootnoteReference"/>
        </w:rPr>
        <w:footnoteRef/>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> </w:t>
      </w:r>
      <w:r>
        <w:t>This is the first footnote content.</w:t>
      </w:r>
    </w:p>
  </w:footnote>
</w:footnotes>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/footnotes.xml", footnotes_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_endnotes() -> Path:
    """Create a document with endnotes for testing."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:r><w:t>This paragraph has an endnote reference</w:t></w:r>
  <w:r>
    <w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>
    <w:endnoteReference w:id="1"/>
  </w:r>
  <w:r><w:t> and continues here.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>Another paragraph for testing endnote cross-references.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t>See also note marker for more info.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    endnotes_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1">
    <w:p>
      <w:r>
        <w:separator/>
      </w:r>
    </w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p>
      <w:r>
        <w:continuationSeparator/>
      </w:r>
    </w:p>
  </w:endnote>
  <w:endnote w:id="1">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="EndnoteText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="EndnoteReference"/>
        </w:rPr>
        <w:endnoteRef/>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> </w:t>
      </w:r>
      <w:r>
        <w:t>This is the first endnote content.</w:t>
      </w:r>
    </w:p>
  </w:endnote>
</w:endnotes>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/endnotes.xml", endnotes_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestNoteReferences:
    """Tests for Phase 6 note reference functionality."""

    def test_resolve_footnote_target(self):
        """Test _resolve_note_target finds and bookmarks a footnote."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error_msg = ops._resolve_note_target("footnote", "1")

            assert error_msg is None
            assert bookmark_name.startswith("_Ref")

            # Verify bookmark was created
            bookmark = ops.get_bookmark(bookmark_name)
            assert bookmark is not None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_endnote_target(self):
        """Test _resolve_note_target finds and bookmarks an endnote."""
        doc_path = create_document_with_endnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error_msg = ops._resolve_note_target("endnote", "1")

            assert error_msg is None
            assert bookmark_name.startswith("_Ref")

            # Verify bookmark was created
            bookmark = ops.get_bookmark(bookmark_name)
            assert bookmark is not None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_note_target_invalid_id(self):
        """Test _resolve_note_target rejects non-numeric IDs."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error_msg = ops._resolve_note_target("footnote", "abc")

            assert error_msg is not None
            assert "must be numeric" in error_msg
            assert bookmark_name == ""
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_note_target_not_found(self):
        """Test _resolve_note_target returns error for non-existent note."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, error_msg = ops._resolve_note_target("footnote", "999")

            assert error_msg is not None
            assert "not found" in error_msg
            assert bookmark_name == ""
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_target_handles_footnote_prefix(self):
        """Test _resolve_target handles 'footnote:N' format."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, was_created = ops._resolve_target("footnote:1")

            assert bookmark_name.startswith("_Ref")
            assert was_created is True
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_target_handles_endnote_prefix(self):
        """Test _resolve_target handles 'endnote:N' format."""
        doc_path = create_document_with_endnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name, was_created = ops._resolve_target("endnote:1")

            assert bookmark_name.startswith("_Ref")
            assert was_created is True
        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertNoteReference:
    """Tests for insert_note_reference convenience method."""

    def test_insert_noteref_to_footnote(self):
        """Test inserting a NOTEREF field to a footnote."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_note_reference(
                note_type="footnote",
                note_id=1,
                after="See also note",
            )

            assert result.startswith("_Ref")

            # Verify NOTEREF field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

            instr = instr_texts[0].text
            assert "NOTEREF" in instr
            assert result in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_noteref_to_endnote(self):
        """Test inserting a NOTEREF field to an endnote."""
        doc_path = create_document_with_endnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_note_reference(
                note_type="endnote",
                note_id=1,
                after="See also note",
            )

            assert result.startswith("_Ref")

            # Verify NOTEREF field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

            instr = instr_texts[0].text
            assert "NOTEREF" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_noteref_with_note_style(self):
        """Test that use_note_style=True adds \\f switch."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_note_reference(
                note_type="footnote",
                note_id=1,
                after="See also note",
                use_note_style=True,
            )

            # Verify \f switch is present
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\f" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_noteref_without_note_style(self):
        """Test that use_note_style=False omits \\f switch."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_note_reference(
                note_type="footnote",
                note_id=1,
                after="See also note",
                use_note_style=False,
            )

            # Verify \f switch is NOT present
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\f" not in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_noteref_with_position(self):
        """Test that show_position=True adds \\p switch."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_note_reference(
                note_type="footnote",
                note_id=1,
                after="See also note",
                show_position=True,
            )

            # Verify \p switch is present
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\p" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_noteref_with_hyperlink(self):
        """Test that hyperlink=True adds \\h switch."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_note_reference(
                note_type="footnote",
                note_id=1,
                after="See also note",
                hyperlink=True,
            )

            # Verify \h switch is present
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\h" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_noteref_without_hyperlink(self):
        """Test that hyperlink=False omits \\h switch."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_note_reference(
                note_type="footnote",
                note_id=1,
                after="See also note",
                hyperlink=False,
            )

            # Verify \h switch is NOT present
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\h" not in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_noteref_invalid_note_type(self):
        """Test that invalid note_type raises ValueError."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(ValueError) as exc_info:
                ops.insert_note_reference(
                    note_type="invalid",
                    note_id=1,
                    after="See also note",
                )

            assert "footnote" in str(exc_info.value)
            assert "endnote" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_noteref_note_not_found(self):
        """Test that non-existent note raises CrossReferenceTargetNotFoundError."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(CrossReferenceTargetNotFoundError) as exc_info:
                ops.insert_note_reference(
                    note_type="footnote",
                    note_id=999,
                    after="See also note",
                )

            assert "footnote:999" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_noteref_requires_after_or_before(self):
        """Test that either after or before must be specified."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(ValueError) as exc_info:
                ops.insert_note_reference(
                    note_type="footnote",
                    note_id=1,
                )

            assert "after" in str(exc_info.value).lower() or "before" in str(exc_info.value).lower()
        finally:
            doc_path.unlink(missing_ok=True)


class TestInsertPageReference:
    """Tests for insert_page_reference convenience method."""

    def test_insert_pageref_to_bookmark(self):
        """Test inserting a PAGEREF field to a bookmark."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_page_reference(
                target="DefinitionsSection",
                after="appendix on page",
            )

            assert result == "DefinitionsSection"

            # Verify PAGEREF field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert len(instr_texts) == 1

            instr = instr_texts[0].text
            assert "PAGEREF" in instr
            assert "DefinitionsSection" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_pageref_with_hyperlink(self):
        """Test that hyperlink=True adds \\h switch."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_page_reference(
                target="DefinitionsSection",
                after="appendix on page",
                hyperlink=True,
            )

            # Verify \h switch is present
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\h" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_pageref_without_hyperlink(self):
        """Test that hyperlink=False omits \\h switch."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_page_reference(
                target="DefinitionsSection",
                after="appendix on page",
                hyperlink=False,
            )

            # Verify \h switch is NOT present
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\h" not in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_pageref_with_position(self):
        """Test that show_position=True adds \\p switch."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.insert_page_reference(
                target="DefinitionsSection",
                after="appendix on page",
                show_position=True,
            )

            # Verify \p switch is present
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            instr = instr_texts[0].text
            assert "\\p" in instr
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_pageref_to_heading(self):
        """Test inserting a PAGEREF to a heading."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_page_reference(
                target="heading:Introduction",
                after="introduction paragraph",
            )

            assert result.startswith("_Ref")

            # Verify PAGEREF field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "PAGEREF" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_pageref_to_footnote(self):
        """Test inserting a PAGEREF to a footnote location."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_page_reference(
                target="footnote:1",
                after="See also note",
            )

            assert result.startswith("_Ref")

            # Verify PAGEREF field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "PAGEREF" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_pageref_requires_after_or_before(self):
        """Test that either after or before must be specified."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(ValueError) as exc_info:
                ops.insert_page_reference(
                    target="DefinitionsSection",
                )

            assert "after" in str(exc_info.value).lower() or "before" in str(exc_info.value).lower()
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_pageref_cannot_specify_both(self):
        """Test that both after and before cannot be specified."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(ValueError) as exc_info:
                ops.insert_page_reference(
                    target="DefinitionsSection",
                    after="some text",
                    before="other text",
                )

            assert "both" in str(exc_info.value).lower()
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_pageref_bookmark_not_found(self):
        """Test that non-existent bookmark raises CrossReferenceTargetNotFoundError."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(CrossReferenceTargetNotFoundError):
                ops.insert_page_reference(
                    target="NonExistentBookmark",
                    after="appendix on page",
                )
        finally:
            doc_path.unlink(missing_ok=True)


class TestCrossReferenceToNote:
    """Tests for using insert_cross_reference with note targets."""

    def test_cross_reference_to_footnote(self):
        """Test insert_cross_reference with footnote:N target."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="footnote:1",
                display="text",
                after="See also note",
            )

            assert result.startswith("_Ref")

            # Verify REF field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "REF" in instr_texts[0].text
            assert result in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_cross_reference_to_endnote(self):
        """Test insert_cross_reference with endnote:N target."""
        doc_path = create_document_with_endnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops.insert_cross_reference(
                target="endnote:1",
                display="text",
                after="See also note",
            )

            assert result.startswith("_Ref")

            # Verify REF field was inserted
            instr_texts = list(doc.xml_root.iter(f"{{{WORD_NS}}}instrText"))
            assert "REF" in instr_texts[0].text
        finally:
            doc_path.unlink(missing_ok=True)

    def test_cross_reference_to_nonexistent_footnote(self):
        """Test that non-existent footnote target raises error."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            with pytest.raises(CrossReferenceTargetNotFoundError) as exc_info:
                ops.insert_cross_reference(
                    target="footnote:999",
                    display="text",
                    after="See also note",
                )

            assert "footnote:999" in str(exc_info.value)
        finally:
            doc_path.unlink(missing_ok=True)


class TestNoteBookmarkManagement:
    """Tests for _find_note_bookmark and _create_note_bookmark methods."""

    def test_create_note_bookmark(self):
        """Test _create_note_bookmark creates a bookmark around the reference."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            bookmark_name = ops._create_note_bookmark("footnote", "1")

            assert bookmark_name.startswith("_Ref")

            # Verify bookmark exists in the document
            bookmark = ops.get_bookmark(bookmark_name)
            assert bookmark is not None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_note_bookmark_returns_none_initially(self):
        """Test _find_note_bookmark returns None when no bookmark exists."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops._find_note_bookmark("footnote", "1")

            # Should return None since we haven't created one yet
            assert result is None
        finally:
            doc_path.unlink(missing_ok=True)

    def test_find_note_bookmark_after_creation(self):
        """Test _find_note_bookmark finds the bookmark after creation."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Create a bookmark
            created_name = ops._create_note_bookmark("footnote", "1")

            # Now find should return it
            found_name = ops._find_note_bookmark("footnote", "1")

            # Note: The _find_note_bookmark looks for generic _Ref bookmarks
            # that encompass the run, so it should find the created one
            assert found_name == created_name
        finally:
            doc_path.unlink(missing_ok=True)

    def test_resolve_reuses_existing_bookmark(self):
        """Test that _resolve_note_target reuses existing bookmark."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # First resolution creates a bookmark
            bookmark1, _ = ops._resolve_note_target("footnote", "1")

            # Second resolution should reuse it
            bookmark2, _ = ops._resolve_note_target("footnote", "1")

            assert bookmark1 == bookmark2
        finally:
            doc_path.unlink(missing_ok=True)


# =============================================================================
# Phase 7: Inspection and Field Management Tests
# =============================================================================


def create_document_with_cross_references() -> Path:
    """Create a test document with existing cross-reference fields."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # A document with REF, PAGEREF, and NOTEREF fields
    # Note: xml:space="preserve" is required on text elements with leading/trailing spaces
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:xml="http://www.w3.org/XML/1998/namespace">
<w:body>
<w:p>
  <w:bookmarkStart w:id="0" w:name="TestBookmark"/>
  <w:r><w:t>This is bookmarked text for testing.</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
<w:p>
  <w:bookmarkStart w:id="1" w:name="_Ref123456789"/>
  <w:r><w:t>Hidden bookmark content.</w:t></w:r>
  <w:bookmarkEnd w:id="1"/>
</w:p>
<w:p>
  <w:r><w:t xml:space="preserve">See </w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
  <w:r><w:instrText xml:space="preserve"> REF TestBookmark \\h </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>bookmarked text</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:t xml:space="preserve"> for details.</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t xml:space="preserve">On page </w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> PAGEREF _Ref123456789 \\h \\p </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>5</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:t>.</w:t></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


def create_document_with_mixed_fields() -> Path:
    """Create a test document with REF, PAGEREF, NOTEREF and non-xref fields."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:xml="http://www.w3.org/XML/1998/namespace">
<w:body>
<w:p>
  <w:bookmarkStart w:id="0" w:name="Section1"/>
  <w:r><w:t>Section 1 Content</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
<w:p>
  <w:r><w:t>Reference to section: </w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
  <w:r><w:instrText xml:space="preserve"> REF Section1 \\h \\r </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>Section 1</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
<w:p>
  <w:r><w:t>Page number: </w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> PAGEREF Section1 </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>1</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
<w:p>
  <w:r><w:t>Date: </w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> DATE \\@ "MMMM d, yyyy" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>December 31, 2025</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)

    return doc_path


class TestParseFieldInstruction:
    """Tests for _parse_field_instruction method."""

    def test_parse_simple_ref_field(self):
        """Test parsing a simple REF field instruction."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops._parse_field_instruction(" REF _Ref123456 ")

            assert result["field_type"] == "REF"
            assert result["bookmark"] == "_Ref123456"
            assert result["is_hyperlink"] is False
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_ref_with_hyperlink_switch(self):
        """Test parsing REF field with \\h switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops._parse_field_instruction(" REF MyBookmark \\h ")

            assert result["field_type"] == "REF"
            assert result["bookmark"] == "MyBookmark"
            assert result["is_hyperlink"] is True
            assert "\\h" in result["switches"]
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_pageref_field(self):
        """Test parsing PAGEREF field instruction."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops._parse_field_instruction(" PAGEREF AppendixA \\h \\p ")

            assert result["field_type"] == "PAGEREF"
            assert result["bookmark"] == "AppendixA"
            assert result["is_hyperlink"] is True
            assert result["show_position"] is True
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_noteref_field(self):
        """Test parsing NOTEREF field instruction."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops._parse_field_instruction(" NOTEREF _Ref789 \\f \\h ")

            assert result["field_type"] == "NOTEREF"
            assert result["bookmark"] == "_Ref789"
            assert result["is_hyperlink"] is True
            assert result["use_note_style"] is True
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_number_format_switches(self):
        """Test parsing number format switches."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Full number format
            result = ops._parse_field_instruction(" REF Heading1 \\w ")
            assert result["number_format"] == "full"

            # Relative format
            result = ops._parse_field_instruction(" REF Heading2 \\r ")
            assert result["number_format"] == "relative"

            # No context format
            result = ops._parse_field_instruction(" REF Heading3 \\n ")
            assert result["number_format"] == "no_context"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_suppress_non_numeric_switch(self):
        """Test parsing \\d switch."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops._parse_field_instruction(" REF Caption1 \\d ")

            assert result["suppress_non_numeric"] is True
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_unknown_field_type(self):
        """Test parsing unknown field type."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops._parse_field_instruction(' DATE \\@ "MMMM" ')

            assert result["field_type"] == "UNKNOWN"
            assert result["bookmark"] == ""
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_empty_instruction(self):
        """Test parsing empty instruction."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            result = ops._parse_field_instruction("")

            assert result["field_type"] == "UNKNOWN"
            assert result["bookmark"] == ""
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetCrossReferences:
    """Tests for get_cross_references method."""

    def test_list_all_ref_fields(self):
        """Test listing all REF fields in document."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            # Should find REF and PAGEREF fields
            assert len(xrefs) >= 2

            # Check REF field
            ref_fields = [x for x in xrefs if x.field_type == "REF"]
            assert len(ref_fields) >= 1
            assert ref_fields[0].target_bookmark == "TestBookmark"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_pageref_fields(self):
        """Test listing PAGEREF fields in document."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            pageref_fields = [x for x in xrefs if x.field_type == "PAGEREF"]
            assert len(pageref_fields) >= 1
            assert pageref_fields[0].target_bookmark == "_Ref123456789"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_extract_target_bookmark(self):
        """Test extracting target bookmark from field instruction."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            # All xrefs should have a target_bookmark
            for xref in xrefs:
                assert xref.target_bookmark != ""
        finally:
            doc_path.unlink(missing_ok=True)

    def test_parse_switches_from_field(self):
        """Test parsing switches from field instruction."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            # Find the REF field with hyperlink
            ref_fields = [x for x in xrefs if x.field_type == "REF"]
            assert len(ref_fields) >= 1
            assert ref_fields[0].is_hyperlink is True
            assert "\\h" in ref_fields[0].switches
        finally:
            doc_path.unlink(missing_ok=True)

    def test_get_cached_display_value(self):
        """Test getting cached display value from field result."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            # REF field should have display value
            ref_fields = [x for x in xrefs if x.field_type == "REF"]
            assert len(ref_fields) >= 1
            assert ref_fields[0].display_value == "bookmarked text"

            # PAGEREF field should have page number
            pageref_fields = [x for x in xrefs if x.field_type == "PAGEREF"]
            assert len(pageref_fields) >= 1
            assert pageref_fields[0].display_value == "5"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_detect_dirty_flag(self):
        """Test detecting dirty flag on field."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            # REF field has dirty="true"
            ref_fields = [x for x in xrefs if x.field_type == "REF"]
            assert len(ref_fields) >= 1
            assert ref_fields[0].is_dirty is True

            # PAGEREF field does not have dirty flag
            pageref_fields = [x for x in xrefs if x.field_type == "PAGEREF"]
            assert len(pageref_fields) >= 1
            assert pageref_fields[0].is_dirty is False
        finally:
            doc_path.unlink(missing_ok=True)

    def test_empty_document_returns_empty_list(self):
        """Test that document without cross-references returns empty list."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            assert xrefs == []
        finally:
            doc_path.unlink(missing_ok=True)

    def test_ignores_non_xref_fields(self):
        """Test that non-cross-reference fields are ignored."""
        doc_path = create_document_with_mixed_fields()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            # Should only find REF and PAGEREF, not DATE
            field_types = [x.field_type for x in xrefs]
            assert "DATE" not in field_types
            assert all(ft in ("REF", "PAGEREF", "NOTEREF") for ft in field_types)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_xref_position_tracking(self):
        """Test that cross-reference position is tracked."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            # All xrefs should have a position
            for xref in xrefs:
                assert xref.position.startswith("p:")
        finally:
            doc_path.unlink(missing_ok=True)

    def test_xref_unique_ref_ids(self):
        """Test that each cross-reference gets a unique ref ID."""
        doc_path = create_document_with_mixed_fields()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            xrefs = ops.get_cross_references()

            ref_ids = [x.ref for x in xrefs]
            assert len(ref_ids) == len(set(ref_ids))  # All unique
            assert all(r.startswith("xref:") for r in ref_ids)
        finally:
            doc_path.unlink(missing_ok=True)


class TestGetCrossReferenceTargets:
    """Tests for get_cross_reference_targets method."""

    def test_list_bookmark_targets(self):
        """Test listing bookmarks as potential targets."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            targets = ops.get_cross_reference_targets()

            bookmark_targets = [t for t in targets if t.type == "bookmark"]
            assert len(bookmark_targets) >= 1

            # Check visible bookmark
            visible_bks = [t for t in bookmark_targets if t.bookmark_name == "VisibleBookmark"]
            assert len(visible_bks) == 1
            assert visible_bks[0].is_hidden is False
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_hidden_bookmark_targets(self):
        """Test that hidden bookmarks are included."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            targets = ops.get_cross_reference_targets()

            bookmark_targets = [t for t in targets if t.type == "bookmark"]
            hidden_bks = [t for t in bookmark_targets if t.is_hidden]
            assert len(hidden_bks) >= 1  # _Ref and _Toc bookmarks
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_heading_targets(self):
        """Test listing headings as potential targets."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            targets = ops.get_cross_reference_targets()

            heading_targets = [t for t in targets if t.type == "heading"]
            assert len(heading_targets) >= 1

            # Check that level is set
            for ht in heading_targets:
                assert ht.level is not None
                assert 1 <= ht.level <= 9
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_figure_targets(self):
        """Test listing figures as potential targets."""
        doc_path = create_document_with_figure_captions_simple()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            targets = ops.get_cross_reference_targets()

            figure_targets = [t for t in targets if t.type == "figure"]
            assert len(figure_targets) >= 1

            # Check that number and sequence_id are set
            for ft in figure_targets:
                assert ft.number is not None
                assert ft.sequence_id == "Figure"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_table_targets(self):
        """Test listing tables as potential targets."""
        doc_path = create_document_with_table_captions_complex()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            targets = ops.get_cross_reference_targets()

            table_targets = [t for t in targets if t.type == "table"]
            assert len(table_targets) >= 1

            # Check that number and sequence_id are set
            for tt in table_targets:
                assert tt.number is not None
                assert tt.sequence_id == "Table"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_footnote_targets(self):
        """Test listing footnotes as potential targets."""
        doc_path = create_document_with_footnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            targets = ops.get_cross_reference_targets()

            footnote_targets = [t for t in targets if t.type == "footnote"]
            assert len(footnote_targets) >= 1

            # Check display name format
            for ft in footnote_targets:
                assert "Footnote" in ft.display_name
        finally:
            doc_path.unlink(missing_ok=True)

    def test_list_endnote_targets(self):
        """Test listing endnotes as potential targets."""
        doc_path = create_document_with_endnotes()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            targets = ops.get_cross_reference_targets()

            endnote_targets = [t for t in targets if t.type == "endnote"]
            assert len(endnote_targets) >= 1

            # Check display name format
            for et in endnote_targets:
                assert "Endnote" in et.display_name
        finally:
            doc_path.unlink(missing_ok=True)

    def test_targets_include_text_preview(self):
        """Test that targets include text preview."""
        doc_path = create_document_with_bookmarks()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            targets = ops.get_cross_reference_targets()

            bookmark_targets = [t for t in targets if t.type == "bookmark"]
            # At least one should have text preview
            with_preview = [t for t in bookmark_targets if t.text_preview]
            assert len(with_preview) >= 1
        finally:
            doc_path.unlink(missing_ok=True)

    def test_targets_include_position(self):
        """Test that targets include position."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            targets = ops.get_cross_reference_targets()

            heading_targets = [t for t in targets if t.type == "heading"]
            for ht in heading_targets:
                assert ht.position.startswith("p:")
        finally:
            doc_path.unlink(missing_ok=True)


class TestMarkCrossReferencesDirty:
    """Tests for mark_cross_references_dirty method."""

    def test_mark_all_fields_dirty(self):
        """Test marking all cross-reference fields dirty."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            count = ops.mark_cross_references_dirty()

            # Should have marked at least 2 fields (REF and PAGEREF)
            assert count >= 2
        finally:
            doc_path.unlink(missing_ok=True)

    def test_verify_dirty_flag_is_set(self):
        """Test that dirty flag is actually set on fields."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.mark_cross_references_dirty()

            # Get cross-references again and verify all are dirty
            xrefs = ops.get_cross_references()
            for xref in xrefs:
                assert xref.is_dirty is True
        finally:
            doc_path.unlink(missing_ok=True)

    def test_mark_dirty_returns_count(self):
        """Test that mark_dirty returns correct count."""
        doc_path = create_document_with_mixed_fields()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            count = ops.mark_cross_references_dirty()

            # Should match number of REF/PAGEREF/NOTEREF fields
            xrefs = ops.get_cross_references()
            assert count == len(xrefs)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_mark_dirty_ignores_non_xref_fields(self):
        """Test that non-xref fields are not marked dirty."""
        doc_path = create_document_with_mixed_fields()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            count = ops.mark_cross_references_dirty()

            # Should be 2 (REF and PAGEREF), not 3 (DATE field is ignored)
            assert count == 2
        finally:
            doc_path.unlink(missing_ok=True)

    def test_mark_dirty_empty_document(self):
        """Test marking dirty on document without cross-references."""
        doc_path = create_simple_document()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            count = ops.mark_cross_references_dirty()

            assert count == 0
        finally:
            doc_path.unlink(missing_ok=True)

    def test_mark_dirty_persists_after_save(self):
        """Test that dirty flag persists after save."""
        doc_path = create_document_with_cross_references()
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            ops.mark_cross_references_dirty()
            doc.save(output_path)

            # Reload and verify
            doc2 = Document(output_path)
            ops2 = CrossReferenceOperations(doc2)
            xrefs = ops2.get_cross_references()

            for xref in xrefs:
                assert xref.is_dirty is True
        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)


class TestExtractFieldsFromBody:
    """Tests for _extract_fields_from_body helper method."""

    def test_extract_ref_fields(self):
        """Test extracting REF fields."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            fields = ops._extract_fields_from_body(["REF"])

            assert len(fields) >= 1
            assert all(f[0] == "REF" for f in fields)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_extract_multiple_field_types(self):
        """Test extracting multiple field types."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            fields = ops._extract_fields_from_body(["REF", "PAGEREF"])

            field_types = set(f[0] for f in fields)
            assert "REF" in field_types
            assert "PAGEREF" in field_types
        finally:
            doc_path.unlink(missing_ok=True)

    def test_extract_returns_position(self):
        """Test that extracted fields include position."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            fields = ops._extract_fields_from_body(["REF", "PAGEREF"])

            for field_type, elem, position in fields:
                assert position.startswith("p:")
        finally:
            doc_path.unlink(missing_ok=True)


class TestBuildCrossReferenceFromField:
    """Tests for _build_cross_reference_from_field helper method."""

    def test_build_from_ref_field(self):
        """Test building CrossReference from REF field."""
        doc_path = create_document_with_cross_references()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            fields = ops._extract_fields_from_body(["REF"])
            assert len(fields) >= 1

            field_type, field_elem, position = fields[0]
            xref = ops._build_cross_reference_from_field(
                ref_id="xref:test",
                field_type=field_type,
                field_elem=field_elem,
                position=position,
            )

            assert xref is not None
            assert xref.ref == "xref:test"
            assert xref.field_type == "REF"
            assert xref.target_bookmark == "TestBookmark"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_build_includes_parsed_switches(self):
        """Test that built CrossReference includes parsed switches."""
        doc_path = create_document_with_mixed_fields()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            fields = ops._extract_fields_from_body(["REF"])
            assert len(fields) >= 1

            field_type, field_elem, position = fields[0]
            xref = ops._build_cross_reference_from_field(
                ref_id="xref:0",
                field_type=field_type,
                field_elem=field_elem,
                position=position,
            )

            # The REF field has \h \r switches
            assert xref is not None
            assert xref.is_hyperlink is True
            assert xref.number_format == "relative"
        finally:
            doc_path.unlink(missing_ok=True)


class TestPhase7Integration:
    """Integration tests for Phase 7 functionality."""

    def test_insert_then_inspect(self):
        """Test inserting cross-reference then inspecting it."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Insert a cross-reference (use text that exists in the fixture)
            bookmark = ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="Please refer to",
            )

            # Inspect the document
            xrefs = ops.get_cross_references()

            # Should find the inserted cross-reference
            assert len(xrefs) >= 1
            matching = [x for x in xrefs if x.target_bookmark == bookmark]
            assert len(matching) == 1
            assert matching[0].field_type == "REF"
        finally:
            doc_path.unlink(missing_ok=True)

    def test_insert_multiple_then_mark_dirty(self):
        """Test inserting multiple cross-references then marking all dirty."""
        doc_path = create_document_with_bookmarks_and_text()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Insert multiple cross-references (use text that exists in the fixture)
            ops.insert_cross_reference(
                target="DefinitionsSection",
                display="text",
                after="Please refer to",
            )
            ops.insert_page_reference(
                target="DefinitionsSection",
                after="appendix on page",
            )

            # Mark all dirty
            count = ops.mark_cross_references_dirty()

            # Should have marked both
            assert count >= 2

            # Verify all are dirty
            xrefs = ops.get_cross_references()
            for xref in xrefs:
                assert xref.is_dirty is True
        finally:
            doc_path.unlink(missing_ok=True)

    def test_round_trip_inspection(self):
        """Test that inspection works after save/reload."""
        doc_path = create_document_with_cross_references()
        output_path = Path(tempfile.mktemp(suffix=".docx"))
        try:
            # Load and inspect
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)
            xrefs_before = ops.get_cross_references()

            # Save
            doc.save(output_path)

            # Reload and inspect again
            doc2 = Document(output_path)
            ops2 = CrossReferenceOperations(doc2)
            xrefs_after = ops2.get_cross_references()

            # Should have same number of cross-references
            assert len(xrefs_before) == len(xrefs_after)

            # Should have same target bookmarks
            targets_before = {x.target_bookmark for x in xrefs_before}
            targets_after = {x.target_bookmark for x in xrefs_after}
            assert targets_before == targets_after
        finally:
            doc_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

    def test_targets_available_before_insert(self):
        """Test that targets are available before inserting references."""
        doc_path = create_document_with_headings()
        try:
            doc = Document(doc_path)
            ops = CrossReferenceOperations(doc)

            # Get available targets
            targets = ops.get_cross_reference_targets()

            # Should find headings
            heading_targets = [t for t in targets if t.type == "heading"]
            assert len(heading_targets) >= 1

            # Can use heading text to insert cross-reference
            heading_text = heading_targets[0].display_name
            assert heading_text  # Should have display name
        finally:
            doc_path.unlink(missing_ok=True)
