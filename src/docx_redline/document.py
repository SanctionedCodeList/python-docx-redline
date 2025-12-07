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

    def delete_tracked(
        self, text: str, author: str | None = None
    ) -> None:
        """Delete text with tracked changes.

        This method searches for the specified text in the document and marks
        it as a tracked deletion.

        Args:
            text: The text to delete
            author: Optional author override (uses document author if None)

        Raises:
            TextNotFoundError: If the text is not found
            AmbiguousTextError: If multiple occurrences of text are found
        """
        # Get all paragraphs in the document
        paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Search for the text to delete
        matches = self._text_search.find_text(text, paragraphs)

        if not matches:
            raise TextNotFoundError(
                text,
                suggestions=[
                    "Check for typos in the search text",
                    "Try searching for a shorter or more unique phrase",
                    "Verify the text exists in the document",
                ],
            )

        if len(matches) > 1:
            raise AmbiguousTextError(text, matches)

        # We have exactly one match
        match = matches[0]

        # Generate the deletion XML
        deletion_xml = self._xml_generator.create_deletion(text, author)

        # Parse the deletion XML with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        deletion_element = root[0]  # Get the first child (the actual deletion)

        # Replace the matched text with deletion
        self._replace_match_with_element(match, deletion_element)

    def replace_tracked(
        self, find: str, replace: str, author: str | None = None
    ) -> None:
        """Find and replace text with tracked changes.

        This method searches for text and replaces it with new text, showing
        both the deletion of the old text and insertion of the new text as
        tracked changes.

        Args:
            find: Text to find
            replace: Replacement text
            author: Optional author override (uses document author if None)

        Raises:
            TextNotFoundError: If the 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
        """
        # Get all paragraphs in the document
        paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Search for the text to replace
        matches = self._text_search.find_text(find, paragraphs)

        if not matches:
            raise TextNotFoundError(
                find,
                suggestions=[
                    "Check for typos in the search text",
                    "Try searching for a shorter or more unique phrase",
                    "Verify the text exists in the document",
                ],
            )

        if len(matches) > 1:
            raise AmbiguousTextError(find, matches)

        # We have exactly one match
        match = matches[0]

        # Generate deletion XML for the old text
        deletion_xml = self._xml_generator.create_deletion(find, author)

        # Generate insertion XML for the new text
        insertion_xml = self._xml_generator.create_insertion(replace, author)

        # Parse both XMLs with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
    {insertion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        deletion_element = root[0]  # First child (deletion)
        insertion_element = root[1]  # Second child (insertion)

        # Replace the matched text with deletion + insertion
        self._replace_match_with_elements(match, [deletion_element, insertion_element])

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

    def _replace_match_with_element(self, match: Any, replacement_element: Any) -> None:
        """Replace matched text with a single XML element.

        This handles the complexity of text potentially spanning multiple runs.
        The matched runs are removed and replaced with the new element.

        Args:
            match: TextSpan object representing the text to replace
            replacement_element: The lxml Element to insert in place of matched text
        """
        paragraph = match.paragraph

        # If the match is within a single run
        if match.start_run_index == match.end_run_index:
            run = match.runs[match.start_run_index]
            run_text = "".join(run.itertext())

            # If the match is the entire run, replace the run
            if match.start_offset == 0 and match.end_offset == len(run_text):
                run_index = list(paragraph).index(run)
                paragraph.remove(run)
                paragraph.insert(run_index, replacement_element)
            else:
                # Match is partial - need to split the run
                self._split_and_replace_in_run(
                    paragraph, run, match.start_offset, match.end_offset, replacement_element
                )
        else:
            # Match spans multiple runs - remove all matched runs and insert replacement
            start_run = match.runs[match.start_run_index]
            start_run_index = list(paragraph).index(start_run)

            # Remove all runs in the match
            for i in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[i]
                if run in paragraph:
                    paragraph.remove(run)

            # Insert replacement at the position of the first removed run
            paragraph.insert(start_run_index, replacement_element)

    def _replace_match_with_elements(
        self, match: Any, replacement_elements: list[Any]
    ) -> None:
        """Replace matched text with multiple XML elements.

        Used for replace_tracked which needs both deletion and insertion elements.

        Args:
            match: TextSpan object representing the text to replace
            replacement_elements: List of lxml Elements to insert in place of matched text
        """
        paragraph = match.paragraph

        # Similar to _replace_match_with_element but inserts multiple elements
        if match.start_run_index == match.end_run_index:
            run = match.runs[match.start_run_index]
            run_text = "".join(run.itertext())

            # If the match is the entire run, replace the run
            if match.start_offset == 0 and match.end_offset == len(run_text):
                run_index = list(paragraph).index(run)
                paragraph.remove(run)
                # Insert elements in order
                for i, elem in enumerate(replacement_elements):
                    paragraph.insert(run_index + i, elem)
            else:
                # Match is partial - need to split the run
                self._split_and_replace_in_run_multiple(
                    paragraph,
                    run,
                    match.start_offset,
                    match.end_offset,
                    replacement_elements,
                )
        else:
            # Match spans multiple runs
            start_run = match.runs[match.start_run_index]
            start_run_index = list(paragraph).index(start_run)

            # Remove all runs in the match
            for i in range(match.start_run_index, match.end_run_index + 1):
                run = match.runs[i]
                if run in paragraph:
                    paragraph.remove(run)

            # Insert all replacement elements at the position of the first removed run
            for i, elem in enumerate(replacement_elements):
                paragraph.insert(start_run_index + i, elem)

    def _split_and_replace_in_run(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_element: Any,
    ) -> None:
        """Split a run and replace a portion with a new element.

        Args:
            paragraph: The paragraph containing the run
            run: The run to split
            start_offset: Character offset where match starts
            end_offset: Character offset where match ends (exclusive)
            replacement_element: Element to insert in place of matched text
        """
        # Get the full text of the run
        text_elements = list(run.iter(f"{{{WORD_NAMESPACE}}}t"))
        if not text_elements:
            return

        # For simplicity, we'll work with the first text element
        # (Word typically has one w:t per run)
        text_elem = text_elements[0]
        run_text = text_elem.text or ""

        # Split into before, match, after
        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        run_index = list(paragraph).index(run)

        # Build new elements
        new_elements = []

        # Add before text if it exists
        if before_text:
            before_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            # Copy run properties if they exist
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                before_run.append(etree.fromstring(etree.tostring(run_props)))
            before_t = etree.SubElement(before_run, f"{{{WORD_NAMESPACE}}}t")
            if before_text and (before_text[0].isspace() or before_text[-1].isspace()):
                before_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            before_t.text = before_text
            new_elements.append(before_run)

        # Add replacement element
        new_elements.append(replacement_element)

        # Add after text if it exists
        if after_text:
            after_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            # Copy run properties if they exist
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                after_run.append(etree.fromstring(etree.tostring(run_props)))
            after_t = etree.SubElement(after_run, f"{{{WORD_NAMESPACE}}}t")
            if after_text and (after_text[0].isspace() or after_text[-1].isspace()):
                after_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            after_t.text = after_text
            new_elements.append(after_run)

        # Remove original run
        paragraph.remove(run)

        # Insert new elements
        for i, elem in enumerate(new_elements):
            paragraph.insert(run_index + i, elem)

    def _split_and_replace_in_run_multiple(
        self,
        paragraph: Any,
        run: Any,
        start_offset: int,
        end_offset: int,
        replacement_elements: list[Any],
    ) -> None:
        """Split a run and replace a portion with multiple new elements.

        Args:
            paragraph: The paragraph containing the run
            run: The run to split
            start_offset: Character offset where match starts
            end_offset: Character offset where match ends (exclusive)
            replacement_elements: Elements to insert in place of matched text
        """
        # Get the full text of the run
        text_elements = list(run.iter(f"{{{WORD_NAMESPACE}}}t"))
        if not text_elements:
            return

        text_elem = text_elements[0]
        run_text = text_elem.text or ""

        # Split into before, match, after
        before_text = run_text[:start_offset]
        after_text = run_text[end_offset:]

        run_index = list(paragraph).index(run)

        # Build new elements
        new_elements = []

        # Add before text if it exists
        if before_text:
            before_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                before_run.append(etree.fromstring(etree.tostring(run_props)))
            before_t = etree.SubElement(before_run, f"{{{WORD_NAMESPACE}}}t")
            if before_text and (before_text[0].isspace() or before_text[-1].isspace()):
                before_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            before_t.text = before_text
            new_elements.append(before_run)

        # Add all replacement elements
        new_elements.extend(replacement_elements)

        # Add after text if it exists
        if after_text:
            after_run = etree.Element(f"{{{WORD_NAMESPACE}}}r")
            run_props = run.find(f"{{{WORD_NAMESPACE}}}rPr")
            if run_props is not None:
                after_run.append(etree.fromstring(etree.tostring(run_props)))
            after_t = etree.SubElement(after_run, f"{{{WORD_NAMESPACE}}}t")
            if after_text and (after_text[0].isspace() or after_text[-1].isspace()):
                after_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            after_t.text = after_text
            new_elements.append(after_run)

        # Remove original run
        paragraph.remove(run)

        # Insert new elements
        for i, elem in enumerate(new_elements):
            paragraph.insert(run_index + i, elem)

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
