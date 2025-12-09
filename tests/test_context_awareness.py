"""
Tests for context-aware text replacement features.

These tests verify:
1. Context preview (show_context parameter)
2. Sentence fragment detection (check_continuity parameter)
3. Warning generation for potential continuity issues
"""

import tempfile
import warnings
import zipfile
from pathlib import Path

from docx_redline import ContinuityWarning, Document


def create_test_document(text: str) -> Path:
    """Create a simple but valid test document with proper OOXML structure."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    # Proper Content_Types.xml
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    # Proper relationships file
    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    document_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>{text}</w:t></w:r></w:p>
</w:body>
</w:document>"""

    with zipfile.ZipFile(doc_path, "w") as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document_xml)

    return doc_path


def test_replace_with_context_preview(capsys):
    """Test that show_context parameter displays surrounding text."""
    text = "Sentence one. Sentence two. Sentence three."
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)
        doc.replace_tracked(
            "Sentence two.",
            "New sentence.",
            show_context=True,
            context_chars=20,
        )

        captured = capsys.readouterr()

        # Verify context preview was printed
        assert "CONTEXT PREVIEW" in captured.out
        assert "BEFORE" in captured.out
        assert "MATCH" in captured.out
        assert "AFTER" in captured.out
        assert "REPLACEMENT" in captured.out

        # Verify actual content
        assert "Sentence one." in captured.out  # Before
        assert "Sentence two." in captured.out  # Match
        assert "Sentence three." in captured.out  # After
        assert "New sentence." in captured.out  # Replacement

    finally:
        doc_path.unlink()


def test_context_preview_with_different_lengths(capsys):
    """Test context preview with different context_chars values."""
    text = "A" * 100 + " MATCH " + "B" * 100
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)
        doc.replace_tracked(
            "MATCH",
            "REPLACEMENT",
            show_context=True,
            context_chars=10,
        )

        captured = capsys.readouterr()

        # Verify truncation with ellipsis
        assert "..." in captured.out

    finally:
        doc_path.unlink()


def test_fragment_detection_lowercase_start():
    """Test detection of sentence fragment starting with lowercase."""
    # Use a lowercase word that's not a connecting phrase
    text = "The product functions well. and this continues the thought."
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            doc.replace_tracked(
                "The product functions well.",
                "It works great.",
                check_continuity=True,
            )

            # Should have issued a warning about lowercase
            assert len(w) >= 1
            assert issubclass(w[0].category, ContinuityWarning)
            assert "lowercase" in str(w[0].message).lower()

    finally:
        doc_path.unlink()


def test_fragment_detection_connecting_phrase():
    """Test detection of sentence fragment with connecting phrase."""
    text = (
        "BatchLeads functions like the attorney directory. in question here is property ownership."
    )
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            doc.replace_tracked(
                "BatchLeads functions like the attorney directory.",
                "BatchLeads works well.",
                check_continuity=True,
            )

            # Should have issued a warning
            assert len(w) >= 1
            assert issubclass(w[0].category, ContinuityWarning)
            assert "in question" in str(w[0].message).lower()

    finally:
        doc_path.unlink()


def test_fragment_detection_continuation_punctuation():
    """Test detection of sentence fragment starting with continuation punctuation."""
    text = "First part of sentence, and second part continues here."
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            doc.replace_tracked(
                "First part of sentence",
                "New text",
                check_continuity=True,
            )

            # Should have issued a warning about comma
            assert len(w) >= 1
            assert issubclass(w[0].category, ContinuityWarning)
            assert "continuation punctuation" in str(w[0].message).lower()

    finally:
        doc_path.unlink()


def test_no_warning_for_valid_replacement():
    """Test that no warning is issued for grammatically correct replacements."""
    text = "Sentence one. Sentence two. Sentence three."
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            doc.replace_tracked(
                "Sentence two.",
                "New sentence.",
                check_continuity=True,
            )

            # Should NOT have issued any warnings
            continuity_warnings = [
                warning for warning in w if issubclass(warning.category, ContinuityWarning)
            ]
            assert len(continuity_warnings) == 0

    finally:
        doc_path.unlink()


def test_warning_includes_suggestions():
    """Test that warnings include helpful suggestions."""
    text = "The product functions. in question here is ownership."
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            doc.replace_tracked(
                "The product functions.",
                "It works.",
                check_continuity=True,
            )

            assert len(w) >= 1
            warning_msg = str(w[0].message)

            # Verify suggestions are included
            assert "Suggestions:" in warning_msg or "suggestion" in warning_msg.lower()

    finally:
        doc_path.unlink()


def test_both_features_together(capsys):
    """Test using both show_context and check_continuity together."""
    text = "First sentence. in question here is the topic."
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            doc.replace_tracked(
                "First sentence.",
                "New sentence.",
                show_context=True,
                check_continuity=True,
                context_chars=30,
            )

            # Should show context
            captured = capsys.readouterr()
            assert "CONTEXT PREVIEW" in captured.out

            # Should issue warning
            assert len(w) >= 1
            assert issubclass(w[0].category, ContinuityWarning)

    finally:
        doc_path.unlink()


def test_default_parameters_no_preview_no_warning(capsys):
    """Test that default parameters don't show preview or check continuity."""
    text = "First sentence. in question here is a fragment."
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            # Use default parameters (show_context=False, check_continuity=False)
            doc.replace_tracked(
                "First sentence.",
                "New sentence.",
            )

            # Should NOT show context
            captured = capsys.readouterr()
            assert "CONTEXT PREVIEW" not in captured.out

            # Should NOT issue warning
            continuity_warnings = [
                warning for warning in w if issubclass(warning.category, ContinuityWarning)
            ]
            assert len(continuity_warnings) == 0

    finally:
        doc_path.unlink()


def test_roman_numeral_i_not_flagged():
    """Test that Roman numeral 'i' is not flagged as lowercase fragment."""
    text = "Section A. i. This is a Roman numeral list item."
    doc_path = create_test_document(text)

    try:
        doc = Document(doc_path)

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            doc.replace_tracked(
                "Section A.",
                "Part 1.",
                check_continuity=True,
            )

            # Should NOT warn about 'i' being lowercase
            # (may warn about other things, but not lowercase)
            for warning in w:
                if issubclass(warning.category, ContinuityWarning):
                    assert "lowercase" not in str(warning.message).lower()

    finally:
        doc_path.unlink()
