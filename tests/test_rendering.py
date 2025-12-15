"""Tests for document rendering to images."""

from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path
from unittest.mock import patch

import pytest

from python_docx_redline import Document
from python_docx_redline.rendering import (
    LIBREOFFICE_PATH_ENV,
    PDFTOPPM_PATH_ENV,
    _find_libreoffice,
    _find_pdftoppm,
    is_libreoffice_available,
    is_pdftoppm_available,
    is_rendering_available,
    render_document_to_images,
)


def create_test_document(text: str = "Hello World") -> Path:
    """Create a simple but valid test document with proper OOXML structure."""
    doc_path = Path(tempfile.mktemp(suffix=".docx"))

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

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

    with zipfile.ZipFile(doc_path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document_xml)

    return doc_path


# Reset caches between tests
@pytest.fixture(autouse=True)
def reset_caches():
    """Reset path caches before each test."""
    import python_docx_redline.rendering as rendering

    rendering._libreoffice_path_cache = None
    rendering._libreoffice_checked = False
    rendering._pdftoppm_path_cache = None
    rendering._pdftoppm_checked = False
    yield


class TestLibreOfficeDetection:
    """Tests for LibreOffice executable detection."""

    def test_find_libreoffice_from_env_var(self, tmp_path: Path) -> None:
        """Test finding LibreOffice from environment variable."""
        # Create a mock executable
        mock_exe = tmp_path / "soffice"
        mock_exe.write_text("#!/bin/bash\necho mock")
        mock_exe.chmod(0o755)

        with patch.dict("os.environ", {LIBREOFFICE_PATH_ENV: str(mock_exe)}):
            result = _find_libreoffice()
            assert result == str(mock_exe)

    def test_find_libreoffice_invalid_env_var(self, tmp_path: Path) -> None:
        """Test handling invalid path in environment variable."""
        with patch.dict("os.environ", {LIBREOFFICE_PATH_ENV: "/nonexistent/soffice"}):
            # Should fall through to other methods
            with patch("shutil.which", return_value=None):
                _find_libreoffice()
                # Result depends on whether LO is actually installed
                # Just verify it doesn't crash

    def test_find_libreoffice_from_path(self) -> None:
        """Test finding LibreOffice from system PATH."""
        with patch.dict("os.environ", {}, clear=False):
            # Remove env var if present
            import os

            env = dict(os.environ)
            env.pop(LIBREOFFICE_PATH_ENV, None)

            with patch.dict("os.environ", env, clear=True):
                with patch("shutil.which", return_value="/usr/bin/soffice"):
                    with patch("os.path.isfile", return_value=False):  # Skip default paths
                        result = _find_libreoffice()
                        assert result == "/usr/bin/soffice"

    def test_is_libreoffice_available_true(self) -> None:
        """Test availability check when LibreOffice is found."""
        with patch(
            "python_docx_redline.rendering._find_libreoffice",
            return_value="/usr/bin/soffice",
        ):
            assert is_libreoffice_available() is True

    def test_is_libreoffice_available_false(self) -> None:
        """Test availability check when LibreOffice is not found."""
        with patch("python_docx_redline.rendering._find_libreoffice", return_value=None):
            assert is_libreoffice_available() is False


class TestPdftoppmDetection:
    """Tests for pdftoppm executable detection."""

    def test_find_pdftoppm_from_env_var(self, tmp_path: Path) -> None:
        """Test finding pdftoppm from environment variable."""
        mock_exe = tmp_path / "pdftoppm"
        mock_exe.write_text("#!/bin/bash\necho mock")
        mock_exe.chmod(0o755)

        with patch.dict("os.environ", {PDFTOPPM_PATH_ENV: str(mock_exe)}):
            result = _find_pdftoppm()
            assert result == str(mock_exe)

    def test_find_pdftoppm_from_path(self) -> None:
        """Test finding pdftoppm from system PATH."""
        with patch.dict("os.environ", {}, clear=False):
            import os

            env = dict(os.environ)
            env.pop(PDFTOPPM_PATH_ENV, None)

            with patch.dict("os.environ", env, clear=True):
                with patch("shutil.which", return_value="/usr/bin/pdftoppm"):
                    with patch("os.path.isfile", return_value=False):  # Skip default paths
                        result = _find_pdftoppm()
                        assert result == "/usr/bin/pdftoppm"

    def test_is_pdftoppm_available_true(self) -> None:
        """Test availability check when pdftoppm is found."""
        with patch(
            "python_docx_redline.rendering._find_pdftoppm",
            return_value="/usr/bin/pdftoppm",
        ):
            assert is_pdftoppm_available() is True

    def test_is_pdftoppm_available_false(self) -> None:
        """Test availability check when pdftoppm is not found."""
        with patch("python_docx_redline.rendering._find_pdftoppm", return_value=None):
            assert is_pdftoppm_available() is False


class TestRenderingAvailability:
    """Tests for combined rendering availability."""

    def test_is_rendering_available_both_present(self) -> None:
        """Test rendering available when both tools are present."""
        with patch("python_docx_redline.rendering.is_libreoffice_available", return_value=True):
            with patch("python_docx_redline.rendering.is_pdftoppm_available", return_value=True):
                assert is_rendering_available() is True

    def test_is_rendering_available_missing_libreoffice(self) -> None:
        """Test rendering unavailable when LibreOffice is missing."""
        with patch("python_docx_redline.rendering.is_libreoffice_available", return_value=False):
            with patch("python_docx_redline.rendering.is_pdftoppm_available", return_value=True):
                assert is_rendering_available() is False

    def test_is_rendering_available_missing_pdftoppm(self) -> None:
        """Test rendering unavailable when pdftoppm is missing."""
        with patch("python_docx_redline.rendering.is_libreoffice_available", return_value=True):
            with patch("python_docx_redline.rendering.is_pdftoppm_available", return_value=False):
                assert is_rendering_available() is False


class TestRenderDocumentToImages:
    """Tests for the main rendering function."""

    def test_render_missing_file(self) -> None:
        """Test rendering a non-existent file raises FileNotFoundError."""
        with pytest.raises(FileNotFoundError, match="DOCX file not found"):
            render_document_to_images("/nonexistent/file.docx")

    def test_render_no_libreoffice(self) -> None:
        """Test rendering when LibreOffice is not available."""
        doc_path = create_test_document("Test content")
        try:
            with patch("python_docx_redline.rendering._find_libreoffice", return_value=None):
                with pytest.raises(RuntimeError, match="LibreOffice is not available"):
                    render_document_to_images(doc_path)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_render_no_pdftoppm(self) -> None:
        """Test rendering when pdftoppm is not available."""
        doc_path = create_test_document("Test content")
        try:
            with patch(
                "python_docx_redline.rendering._find_libreoffice",
                return_value="/usr/bin/soffice",
            ):
                with patch("python_docx_redline.rendering._find_pdftoppm", return_value=None):
                    with pytest.raises(RuntimeError, match="pdftoppm is not available"):
                        render_document_to_images(doc_path)
        finally:
            doc_path.unlink(missing_ok=True)


@pytest.mark.skipif(
    not is_rendering_available(),
    reason="LibreOffice and/or pdftoppm not available",
)
class TestRenderingIntegration:
    """Integration tests that require LibreOffice and pdftoppm to be installed."""

    def test_render_simple_document(self, tmp_path: Path) -> None:
        """Test rendering a simple document."""
        doc_path = create_test_document("Hello World")
        try:
            output_dir = tmp_path / "images"
            images = render_document_to_images(doc_path, output_dir=output_dir)

            assert len(images) >= 1
            assert all(img.suffix == ".png" for img in images)
            assert all(img.exists() for img in images)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_render_with_custom_dpi(self, tmp_path: Path) -> None:
        """Test rendering with custom DPI setting."""
        doc_path = create_test_document("Test content")
        try:
            output_dir = tmp_path / "images"
            images = render_document_to_images(doc_path, output_dir=output_dir, dpi=72)
            assert len(images) >= 1
        finally:
            doc_path.unlink(missing_ok=True)

    def test_render_with_custom_prefix(self, tmp_path: Path) -> None:
        """Test rendering with custom filename prefix."""
        doc_path = create_test_document("Test content")
        try:
            output_dir = tmp_path / "images"
            images = render_document_to_images(doc_path, output_dir=output_dir, prefix="slide")
            assert len(images) >= 1
            assert all("slide" in img.name for img in images)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_render_to_temp_dir(self) -> None:
        """Test rendering to auto-created temp directory."""
        doc_path = create_test_document("Test content")
        try:
            # Don't specify output_dir - should create temp dir
            images = render_document_to_images(doc_path)

            assert len(images) >= 1
            assert all(img.exists() for img in images)

            # Clean up
            output_dir = images[0].parent
            for img in images:
                img.unlink()
            output_dir.rmdir()
        finally:
            doc_path.unlink(missing_ok=True)

    def test_document_render_to_images_method(self, tmp_path: Path) -> None:
        """Test the Document.render_to_images() method."""
        doc_path = create_test_document("Hello from Document method")
        try:
            doc = Document(doc_path)
            output_dir = tmp_path / "images"
            images = doc.render_to_images(output_dir=output_dir)

            assert len(images) >= 1
            assert all(img.exists() for img in images)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_document_render_unsaved_document(self, tmp_path: Path) -> None:
        """Test rendering a document with modifications (uses temp file)."""
        doc_path = create_test_document("Original content")
        try:
            doc = Document(doc_path)
            # Make some changes (document is now "modified")
            doc.insert_tracked("Additional content", after="Original")

            output_dir = tmp_path / "images"
            images = doc.render_to_images(output_dir=output_dir)

            assert len(images) >= 1
            assert all(img.exists() for img in images)
        finally:
            doc_path.unlink(missing_ok=True)

    def test_render_document_with_tracked_changes(self, tmp_path: Path) -> None:
        """Test rendering a document with tracked changes shows them visually."""
        doc_path = create_test_document("This is some inserted text to test")
        try:
            doc = Document(doc_path)
            doc.delete_tracked("inserted", author="Test User")

            # Save to a new location for rendering
            modified_path = tmp_path / "modified.docx"
            doc.save(modified_path)

            output_dir = tmp_path / "images"
            images = render_document_to_images(modified_path, output_dir=output_dir)

            # Just verify it renders successfully
            # Visual verification would require image comparison
            assert len(images) >= 1
        finally:
            doc_path.unlink(missing_ok=True)
