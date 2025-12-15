"""
OOXMLPackage class for managing Word document ZIP structure.

This module provides a clean abstraction for the OOXML package format,
separating ZIP handling from XML manipulation concerns.
"""

import io
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any, BinaryIO

from lxml import etree

from .validation import ValidationError


class OOXMLPackage:
    """Manages the OOXML ZIP package structure.

    This class handles the low-level operations of:
    - Extracting .docx ZIP archives to temporary directories
    - Providing access to package parts (XML files)
    - Repacking modified content back to ZIP format
    - Cleaning up temporary resources

    Example:
        >>> with OOXMLPackage.open("document.docx") as pkg:
        ...     doc_xml = pkg.get_part("word/document.xml")
        ...     # Modify doc_xml...
        ...     pkg.set_part("word/document.xml", doc_xml)
        ...     pkg.save("modified.docx")
    """

    def __init__(self, temp_dir: Path, source_path: Path | None = None) -> None:
        """Initialize package with an already-extracted directory.

        Use the class methods `open()` or `from_bytes()` instead of
        calling this constructor directly.

        Args:
            temp_dir: Path to the extracted package contents
            source_path: Original source file path (for validation reference)
        """
        self._temp_dir = temp_dir
        self._source_path = source_path
        self._closed = False

    @classmethod
    def open(cls, source: str | Path | BinaryIO) -> "OOXMLPackage":
        """Open an OOXML package from a file path or file-like object.

        Args:
            source: Path to .docx file or file-like object containing it

        Returns:
            OOXMLPackage instance with extracted contents

        Raises:
            ValidationError: If the source is not a valid ZIP file
        """
        source_path: Path | None = None

        # Normalize source to Path or BinaryIO
        if isinstance(source, str | Path):
            source_path = Path(source)
            if not source_path.exists():
                raise ValidationError(f"Document not found: {source_path}")
            zip_source: Path | BinaryIO = source_path
        else:
            zip_source = source

        # Verify it's a ZIP file
        if not zipfile.is_zipfile(zip_source):
            raise ValidationError("Source must be a valid .docx (ZIP) file")

        # Reset stream position if it was checked by is_zipfile
        if hasattr(zip_source, "seek"):
            zip_source.seek(0)

        # Extract to temp directory
        temp_dir = Path(tempfile.mkdtemp(prefix="python_docx_redline_"))
        try:
            with zipfile.ZipFile(zip_source, "r") as zip_ref:
                zip_ref.extractall(temp_dir)
        except Exception as e:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            raise ValidationError(f"Failed to extract .docx file: {e}") from e

        return cls(temp_dir, source_path)

    @classmethod
    def from_bytes(cls, data: bytes) -> "OOXMLPackage":
        """Open an OOXML package from bytes.

        Args:
            data: Bytes containing a .docx file

        Returns:
            OOXMLPackage instance with extracted contents
        """
        return cls.open(io.BytesIO(data))

    @property
    def temp_dir(self) -> Path:
        """Get the temporary directory containing extracted package contents."""
        return self._temp_dir

    @property
    def source_path(self) -> Path | None:
        """Get the original source file path, if available."""
        return self._source_path

    def get_part_path(self, part_name: str) -> Path:
        """Get the filesystem path to a package part.

        Args:
            part_name: Relative path within the package (e.g., "word/document.xml")

        Returns:
            Path to the part in the temp directory
        """
        return self._temp_dir / part_name

    def get_part(self, part_name: str) -> etree._Element | None:
        """Get a package part as a parsed XML element.

        Args:
            part_name: Relative path within the package (e.g., "word/document.xml")

        Returns:
            Parsed XML element tree, or None if part doesn't exist
        """
        part_path = self.get_part_path(part_name)
        if not part_path.exists():
            return None

        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(part_path), parser)
        return tree.getroot()

    def set_part(self, part_name: str, element: etree._Element) -> None:
        """Write an XML element to a package part.

        Args:
            part_name: Relative path within the package (e.g., "word/document.xml")
            element: XML element to write
        """
        part_path = self.get_part_path(part_name)

        # Ensure parent directory exists
        part_path.parent.mkdir(parents=True, exist_ok=True)

        # Get the tree from the element
        tree = element.getroottree()
        tree.write(
            str(part_path),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=False,
        )

    def part_exists(self, part_name: str) -> bool:
        """Check if a package part exists.

        Args:
            part_name: Relative path within the package

        Returns:
            True if the part exists
        """
        return self.get_part_path(part_name).exists()

    def _fix_encoding_declarations(self) -> None:
        """Fix encoding declarations in all XML files to use UTF-8.

        OOXML specification requires UTF-8 or UTF-16 encoding, but some tools
        (including Microsoft Word in certain cases) generate files with
        encoding="ASCII". This causes validation failures.
        """
        # Find all XML and .rels files
        xml_files = list(self._temp_dir.rglob("*.xml")) + list(self._temp_dir.rglob("*.rels"))

        for xml_file in xml_files:
            try:
                # Read the file
                with open(xml_file, "rb") as f:
                    data = f.read()

                # Decode and check if it has an encoding declaration
                try:
                    text = data.decode("utf-8")
                except UnicodeDecodeError:
                    text = data.decode("latin-1")

                # Look for encoding declaration
                pattern = r'(<\?xml[^>]*encoding=)["\']([^"\']*)["\']'
                match = re.search(pattern, text[:200])

                if match:
                    encoding = match.group(2).upper()
                    # Only fix if it's not already UTF-8 or UTF-16
                    if encoding not in ["UTF-8", "UTF-16", "UTF-16LE", "UTF-16BE"]:
                        # Replace encoding declaration with UTF-8
                        new_text = re.sub(
                            pattern,
                            r'\1"UTF-8"',
                            text,
                            count=1,
                        )

                        # Write back as UTF-8
                        with open(xml_file, "wb") as f:
                            f.write(new_text.encode("utf-8"))

            except Exception:
                # Ignore errors on individual files
                pass

    def save(self, output_path: str | Path) -> None:
        """Save the package to a .docx file.

        Args:
            output_path: Path to save the .docx file
        """
        output_path = Path(output_path)

        # Fix encoding declarations before packing
        self._fix_encoding_declarations()

        # Create ZIP file
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zip_ref:
            for file in self._temp_dir.rglob("*"):
                if file.is_file():
                    arcname = file.relative_to(self._temp_dir)
                    zip_ref.write(file, arcname)

    def save_to_bytes(self) -> bytes:
        """Save the package to bytes.

        Returns:
            The complete .docx file as bytes
        """
        # Fix encoding declarations before packing
        self._fix_encoding_declarations()

        # Create ZIP in memory
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zip_ref:
            for file in self._temp_dir.rglob("*"):
                if file.is_file():
                    arcname = file.relative_to(self._temp_dir)
                    zip_ref.write(file, arcname)

        buffer.seek(0)
        return buffer.read()

    def close(self) -> None:
        """Clean up temporary directory."""
        if not self._closed and self._temp_dir and self._temp_dir.exists():
            try:
                shutil.rmtree(self._temp_dir)
            except Exception:
                # Ignore cleanup errors
                pass
            self._closed = True

    def __enter__(self) -> "OOXMLPackage":
        """Context manager support."""
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        """Context manager cleanup."""
        self.close()

    def __del__(self) -> None:
        """Clean up temporary directory on garbage collection."""
        self.close()
