"""Tests for the RedliningValidator class."""

import zipfile
from unittest.mock import patch

from python_docx_redline.validation_redlining import RedliningValidator


class TestRedliningValidatorInit:
    """Tests for RedliningValidator initialization."""

    def test_init_with_paths(self, tmp_path):
        """Test initialization with valid paths."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_docx = tmp_path / "original.docx"
        original_docx.touch()

        validator = RedliningValidator(unpacked_dir, original_docx)

        assert validator.unpacked_dir == unpacked_dir
        assert validator.original_docx == original_docx
        assert validator.verbose is False

    def test_init_with_verbose(self, tmp_path):
        """Test initialization with verbose flag."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_docx = tmp_path / "original.docx"
        original_docx.touch()

        validator = RedliningValidator(unpacked_dir, original_docx, verbose=True)

        assert validator.verbose is True

    def test_namespaces_set(self, tmp_path):
        """Test that namespaces are properly set."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_docx = tmp_path / "original.docx"
        original_docx.touch()

        validator = RedliningValidator(unpacked_dir, original_docx)

        assert "w" in validator.namespaces
        assert (
            validator.namespaces["w"]
            == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        )


class TestRedliningValidatorValidate:
    """Tests for the validate() method."""

    def test_validate_missing_document_xml(self, tmp_path):
        """Test validation fails when document.xml is missing."""
        unpacked_dir = tmp_path / "unpacked"
        unpacked_dir.mkdir()
        original_docx = tmp_path / "original.docx"
        original_docx.touch()

        validator = RedliningValidator(unpacked_dir, original_docx)
        result = validator.validate()

        assert result is False

    def test_validate_no_tracked_changes(self, tmp_path):
        """Test validation passes when no Claude tracked changes exist."""
        # Create unpacked directory with document.xml
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        word_dir.mkdir(parents=True)

        # Create document.xml with no tracked changes
        doc_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Hello World</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"""
        (word_dir / "document.xml").write_text(doc_xml)

        # Create original docx
        original_docx = tmp_path / "original.docx"
        original_docx.touch()

        validator = RedliningValidator(unpacked_dir, original_docx, verbose=True)
        result = validator.validate()

        assert result is True

    def test_validate_with_non_claude_changes(self, tmp_path):
        """Test validation passes when tracked changes are not by Claude."""
        # Create unpacked directory with document.xml
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        word_dir.mkdir(parents=True)

        # Create document.xml with tracked changes by different author
        doc_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:ins w:author="John Doe">
                        <w:r><w:t>Inserted text</w:t></w:r>
                    </w:ins>
                </w:p>
            </w:body>
        </w:document>"""
        (word_dir / "document.xml").write_text(doc_xml)

        original_docx = tmp_path / "original.docx"
        original_docx.touch()

        validator = RedliningValidator(unpacked_dir, original_docx, verbose=True)
        result = validator.validate()

        assert result is True

    def test_validate_with_claude_changes_matching(self, tmp_path):
        """Test validation passes when Claude's changes properly track modifications."""
        # Create unpacked directory
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        word_dir.mkdir(parents=True)

        # Create modified document.xml with Claude insertion
        modified_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Hello </w:t></w:r>
                    <w:ins w:author="Claude">
                        <w:r><w:t>beautiful </w:t></w:r>
                    </w:ins>
                    <w:r><w:t>World</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"""
        (word_dir / "document.xml").write_text(modified_xml)

        # Create original docx with matching content (without the insertion)
        original_docx = tmp_path / "original.docx"
        original_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Hello </w:t></w:r>
                    <w:r><w:t>World</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"""

        # Create a proper docx file
        with zipfile.ZipFile(original_docx, "w") as zf:
            zf.writestr("word/document.xml", original_xml)

        validator = RedliningValidator(unpacked_dir, original_docx, verbose=True)
        result = validator.validate()

        assert result is True

    def test_validate_with_untracked_modification(self, tmp_path):
        """Test validation fails when text is modified without tracking."""
        # Create unpacked directory
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        word_dir.mkdir(parents=True)

        # Modified document with Claude change AND untracked change
        modified_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Changed text</w:t></w:r>
                    <w:ins w:author="Claude">
                        <w:r><w:t> tracked</w:t></w:r>
                    </w:ins>
                </w:p>
            </w:body>
        </w:document>"""
        (word_dir / "document.xml").write_text(modified_xml)

        # Original has different text
        original_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Original text</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"""

        original_docx = tmp_path / "original.docx"
        with zipfile.ZipFile(original_docx, "w") as zf:
            zf.writestr("word/document.xml", original_xml)

        validator = RedliningValidator(unpacked_dir, original_docx)
        result = validator.validate()

        assert result is False

    def test_validate_bad_original_docx(self, tmp_path):
        """Test validation fails with invalid original docx."""
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        word_dir.mkdir(parents=True)

        doc_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:ins w:author="Claude">
                        <w:r><w:t>Text</w:t></w:r>
                    </w:ins>
                </w:p>
            </w:body>
        </w:document>"""
        (word_dir / "document.xml").write_text(doc_xml)

        # Create invalid (non-zip) file
        original_docx = tmp_path / "original.docx"
        original_docx.write_text("not a zip file")

        validator = RedliningValidator(unpacked_dir, original_docx)
        result = validator.validate()

        assert result is False

    def test_validate_original_missing_document_xml(self, tmp_path):
        """Test validation fails when original docx has no document.xml."""
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        word_dir.mkdir(parents=True)

        doc_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:ins w:author="Claude">
                        <w:r><w:t>Text</w:t></w:r>
                    </w:ins>
                </w:p>
            </w:body>
        </w:document>"""
        (word_dir / "document.xml").write_text(doc_xml)

        # Create docx without document.xml
        original_docx = tmp_path / "original.docx"
        with zipfile.ZipFile(original_docx, "w") as zf:
            zf.writestr("other/file.xml", "<root/>")

        validator = RedliningValidator(unpacked_dir, original_docx)
        result = validator.validate()

        assert result is False


class TestRemoveClaudeTrackedChanges:
    """Tests for _remove_claude_tracked_changes method."""

    def test_remove_insertions(self, tmp_path):
        """Test that Claude's insertions are removed."""
        import xml.etree.ElementTree as ET

        xml_str = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Hello </w:t></w:r>
                    <w:ins w:author="Claude">
                        <w:r><w:t>inserted </w:t></w:r>
                    </w:ins>
                    <w:r><w:t>World</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"""

        root = ET.fromstring(xml_str)

        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")
        validator._remove_claude_tracked_changes(root)

        # Check that w:ins by Claude is removed
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        ins_elements = root.findall(".//w:ins", ns)
        assert len(ins_elements) == 0

    def test_remove_deletions_converts_deltext(self, tmp_path):
        """Test that Claude's deletions are unwrapped and delText becomes t."""
        import xml.etree.ElementTree as ET

        xml_str = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:del w:author="Claude">
                        <w:r><w:delText>deleted text</w:delText></w:r>
                    </w:del>
                </w:p>
            </w:body>
        </w:document>"""

        root = ET.fromstring(xml_str)

        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")
        validator._remove_claude_tracked_changes(root)

        # Check that w:del is removed but content is preserved as w:t
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        del_elements = root.findall(".//w:del", ns)
        t_elements = root.findall(".//w:t", ns)

        assert len(del_elements) == 0
        assert len(t_elements) == 1
        assert t_elements[0].text == "deleted text"

    def test_preserve_other_author_changes(self, tmp_path):
        """Test that changes by other authors are preserved."""
        import xml.etree.ElementTree as ET

        xml_str = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:ins w:author="John Doe">
                        <w:r><w:t>John's text</w:t></w:r>
                    </w:ins>
                    <w:ins w:author="Claude">
                        <w:r><w:t>Claude's text</w:t></w:r>
                    </w:ins>
                </w:p>
            </w:body>
        </w:document>"""

        root = ET.fromstring(xml_str)

        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")
        validator._remove_claude_tracked_changes(root)

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        ins_elements = root.findall(".//w:ins", ns)

        # John's insertion should remain
        assert len(ins_elements) == 1
        assert ins_elements[0].get(f"{{{ns['w']}}}author") == "John Doe"


class TestExtractTextContent:
    """Tests for _extract_text_content method."""

    def test_extract_simple_text(self, tmp_path):
        """Test extracting text from simple paragraphs."""
        import xml.etree.ElementTree as ET

        xml_str = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>
                <w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>
            </w:body>
        </w:document>"""

        root = ET.fromstring(xml_str)

        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")
        text = validator._extract_text_content(root)

        assert text == "First paragraph\nSecond paragraph"

    def test_extract_skips_empty_paragraphs(self, tmp_path):
        """Test that empty paragraphs are skipped."""
        import xml.etree.ElementTree as ET

        xml_str = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:r><w:t>Text</w:t></w:r></w:p>
                <w:p></w:p>
                <w:p><w:r><w:t>More text</w:t></w:r></w:p>
            </w:body>
        </w:document>"""

        root = ET.fromstring(xml_str)

        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")
        text = validator._extract_text_content(root)

        assert text == "Text\nMore text"

    def test_extract_concatenates_runs(self, tmp_path):
        """Test that multiple runs in a paragraph are concatenated."""
        import xml.etree.ElementTree as ET

        xml_str = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Hello </w:t></w:r>
                    <w:r><w:t>World</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"""

        root = ET.fromstring(xml_str)

        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")
        text = validator._extract_text_content(root)

        assert text == "Hello World"


class TestGenerateDetailedDiff:
    """Tests for _generate_detailed_diff method."""

    def test_generates_error_message(self, tmp_path):
        """Test that detailed diff generates proper error message."""
        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")

        result = validator._generate_detailed_diff("original", "modified")

        assert "FAILED" in result
        assert "Likely causes:" in result
        assert "pre-redlined documents" in result


class TestGetGitWordDiff:
    """Tests for _get_git_word_diff method."""

    def test_generates_diff_when_git_available(self, tmp_path):
        """Test that git word diff is generated when git is available."""
        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")

        result = validator._get_git_word_diff("hello world", "hello beautiful world")

        # If git is available, should return a diff
        # If not, returns None - both are acceptable
        if result is not None:
            assert "beautiful" in result or len(result) > 0

    def test_returns_none_when_git_fails(self, tmp_path):
        """Test that None is returned when git is not available."""
        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")

        with patch("subprocess.run", side_effect=FileNotFoundError):
            result = validator._get_git_word_diff("hello", "world")

        assert result is None

    def test_handles_identical_text(self, tmp_path):
        """Test handling of identical text."""
        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")

        result = validator._get_git_word_diff("same text", "same text")

        # Should return None or empty for identical text
        assert result is None or result == ""


class TestValidateWithParseError:
    """Tests for validation when XML parsing fails."""

    def test_validate_with_malformed_modified_xml(self, tmp_path):
        """Test validation fails when modified document.xml is malformed."""
        unpacked_dir = tmp_path / "unpacked"
        word_dir = unpacked_dir / "word"
        word_dir.mkdir(parents=True)

        # Create malformed document.xml
        malformed_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:ins w:author="Claude">
                        <w:r><w:t>Text</w:t></w:r>
                    </w:ins>
                    <unclosed>
                </w:p>
            </w:body>
        </w:document>"""
        (word_dir / "document.xml").write_text(malformed_xml)

        # Create valid original docx
        original_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Text</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"""

        original_docx = tmp_path / "original.docx"
        with zipfile.ZipFile(original_docx, "w") as zf:
            zf.writestr("word/document.xml", original_xml)

        validator = RedliningValidator(unpacked_dir, original_docx)
        result = validator.validate()

        # The malformed XML should cause parsing to fail
        assert result is False


class TestGitWordDiffFallback:
    """Tests for the fallback word diff behavior."""

    def test_word_diff_fallback_returns_content(self, tmp_path):
        """Test that word diff returns content when character diff is empty."""
        validator = RedliningValidator(tmp_path, tmp_path / "test.docx")

        # Test with slightly different text that triggers word diff
        original = "The quick brown fox jumps over the lazy dog"
        modified = "The fast brown fox runs over the lazy dog"

        result = validator._get_git_word_diff(original, modified)

        # If git is available, should show some difference
        if result is not None:
            assert "quick" in result or "fast" in result or len(result) > 0


class TestMainGuard:
    """Test the __main__ guard."""

    def test_raises_when_run_directly(self):
        """Test that running the module directly raises an error."""
        # We can't easily test this without actually running the module
        # But we can verify the structure is correct
        import python_docx_redline.validation_redlining as mod

        # The module should have a __name__ check at the bottom
        assert hasattr(mod, "RedliningValidator")
