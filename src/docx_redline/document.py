"""
Document class for editing Word documents with tracked changes.

This module provides the main Document class which handles loading .docx files,
inserting tracked changes, and saving the modified documents.
"""

import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

from .errors import AmbiguousTextError, TextNotFoundError, ValidationError
from .text_search import TextSearch
from .tracked_xml import TrackedXMLGenerator


# Word namespace
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": WORD_NAMESPACE}


class Document:
    """Main class for working with Word documents.

    This class handles loading .docx files (unpacking if needed), making tracked
    edits, and saving the results. It provides a high-level API that hides the
    complexity of OOXML manipulation.

    Example:
        >>> doc = Document("contract.docx")
        >>> doc.insert_tracked("new clause text", after="Section 2.1")
        >>> doc.save("contract_edited.docx")

    Attributes:
        path: Path to the document file
        author: Author name for tracked changes
        xml_tree: Parsed XML tree of the document
        xml_root: Root element of the XML tree
    """

    def __init__(self, path: str | Path, author: str = "Claude") -> None:
        """Initialize a Document from a .docx file.

        Args:
            path: Path to the .docx file
            author: Author name for tracked changes (default: "Claude")

        Raises:
            ValidationError: If the document cannot be loaded or is invalid
        """
        self.path = Path(path)
        self.author = author
        self._temp_dir: Path | None = None
        self._is_zip = False

        # Initialize components
        self._text_search = TextSearch()
        self._xml_generator = TrackedXMLGenerator(doc=self, author=author)

        # Load the document
        self._load_document()

    def _load_document(self) -> None:
        """Load and parse the Word document XML.

        If the document is a .docx file (ZIP archive), it will be extracted
        to a temporary directory. The main document.xml is then parsed.

        Raises:
            ValidationError: If the document cannot be loaded
        """
        if not self.path.exists():
            raise ValidationError(f"Document not found: {self.path}")

        # Check if it's a ZIP file (.docx)
        try:
            if zipfile.is_zipfile(self.path):
                self._is_zip = True
                self._extract_docx()
            else:
                # Assume it's already an unpacked XML file
                self._is_zip = False
                self._temp_dir = self.path.parent
        except Exception as e:
            raise ValidationError(f"Failed to load document: {e}") from e

        # Parse the document.xml
        try:
            if self._is_zip:
                document_xml = self._temp_dir / "word" / "document.xml"  # type: ignore
            else:
                document_xml = self.path

            if not document_xml.exists():
                raise ValidationError(f"document.xml not found in {self.path}")

            # Parse XML with lxml
            parser = etree.XMLParser(remove_blank_text=False)
            self.xml_tree = etree.parse(str(document_xml), parser)
            self.xml_root = self.xml_tree.getroot()

        except etree.XMLSyntaxError as e:
            raise ValidationError(f"Invalid XML in document: {e}") from e
        except Exception as e:
            raise ValidationError(f"Failed to parse document XML: {e}") from e

    def _extract_docx(self) -> None:
        """Extract the .docx ZIP archive to a temporary directory."""
        self._temp_dir = Path(tempfile.mkdtemp(prefix="docx_redline_"))

        try:
            with zipfile.ZipFile(self.path, "r") as zip_ref:
                zip_ref.extractall(self._temp_dir)
        except Exception as e:
            # Clean up on failure
            if self._temp_dir and self._temp_dir.exists():
                shutil.rmtree(self._temp_dir)
            raise ValidationError(f"Failed to extract .docx file: {e}") from e

    def insert_tracked(
        self, text: str, after: str, author: str | None = None
    ) -> None:
        """Insert text with tracked changes after a specific location.

        This method searches for the 'after' text in the document and inserts
        the new text immediately after it as a tracked insertion.

        Args:
            text: The text to insert
            after: The text to search for as the insertion point
            author: Optional author override (uses document author if None)

        Raises:
            TextNotFoundError: If the 'after' text is not found
            AmbiguousTextError: If multiple occurrences of 'after' text are found
        """
        # Get all paragraphs in the document
        paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Search for the anchor text
        matches = self._text_search.find_text(after, paragraphs)

        if not matches:
            raise TextNotFoundError(
                after,
                suggestions=[
                    "Check for typos in the search text",
                    "Try searching for a shorter or more unique phrase",
                    "Verify the text exists in the document",
                ],
            )

        if len(matches) > 1:
            raise AmbiguousTextError(after, matches)

        # We have exactly one match
        match = matches[0]

        # Generate the insertion XML
        insertion_xml = self._xml_generator.create_insertion(text, author)

        # Parse the insertion XML with namespace context
        # Need to wrap it with namespace declarations
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {insertion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        insertion_element = root[0]  # Get the first child (the actual insertion)

        # Insert after the matched text
        self._insert_after_match(match, insertion_element)

    def _insert_after_match(self, match: Any, insertion_element: Any) -> None:
        """Insert an XML element after a matched text span.

        Args:
            match: TextSpan object representing where to insert
            insertion_element: The lxml Element to insert
        """
        # Get the paragraph containing the match
        paragraph = match.paragraph

        # Find the run where the match ends
        end_run = match.runs[match.end_run_index]

        # Find the position of the end run in the paragraph
        run_index = list(paragraph).index(end_run)

        # Insert the new element after the end run
        paragraph.insert(run_index + 1, insertion_element)

    def save(self, output_path: str | Path | None = None) -> None:
        """Save the document to a file.

        Args:
            output_path: Path to save the document. If None, saves to original path.

        Raises:
            ValidationError: If the document cannot be saved
        """
        if output_path is None:
            output_path = self.path
        else:
            output_path = Path(output_path)

        try:
            if self._is_zip and self._temp_dir:
                # Write the modified XML back to the temp directory
                document_xml = self._temp_dir / "word" / "document.xml"
                self.xml_tree.write(
                    str(document_xml),
                    encoding="utf-8",
                    xml_declaration=True,
                    pretty_print=False,
                )

                # Create a new .docx ZIP file
                with zipfile.ZipFile(
                    output_path, "w", zipfile.ZIP_DEFLATED
                ) as zip_ref:
                    # Add all files from temp directory
                    for file in self._temp_dir.rglob("*"):
                        if file.is_file():
                            arcname = file.relative_to(self._temp_dir)
                            zip_ref.write(file, arcname)
            else:
                # Save XML directly
                self.xml_tree.write(
                    str(output_path),
                    encoding="utf-8",
                    xml_declaration=True,
                    pretty_print=False,
                )

        except Exception as e:
            raise ValidationError(f"Failed to save document: {e}") from e

    def __del__(self) -> None:
        """Clean up temporary directory on object destruction."""
        if self._temp_dir and self._temp_dir.exists() and self._is_zip:
            try:
                shutil.rmtree(self._temp_dir)
            except Exception:
                # Ignore cleanup errors
                pass

    def __enter__(self) -> "Document":
        """Context manager support."""
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        """Context manager cleanup."""
        self.__del__()
