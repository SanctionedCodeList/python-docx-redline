"""
Document class for editing Word documents with tracked changes.

This module provides the main Document class which handles loading .docx files,
inserting tracked changes, and saving the modified documents.
"""

import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docx_redline.models.paragraph import Paragraph

import yaml
from lxml import etree

from .errors import AmbiguousTextError, TextNotFoundError, ValidationError
from .results import EditResult
from .scope import ScopeEvaluator
from .suggestions import SuggestionGenerator
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
        self,
        text: str,
        after: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> None:
        """Insert text with tracked changes after a specific location.

        This method searches for the 'after' text in the document and inserts
        the new text immediately after it as a tracked insertion.

        Args:
            text: The text to insert
            after: The text to search for as the insertion point
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})

        Raises:
            TextNotFoundError: If the 'after' text is not found
            AmbiguousTextError: If multiple occurrences of 'after' text are found
        """
        # Get all paragraphs in the document
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search for the anchor text
        matches = self._text_search.find_text(after, paragraphs)

        if not matches:
            # Generate smart suggestions
            suggestions = SuggestionGenerator.generate_suggestions(after, paragraphs)
            raise TextNotFoundError(after, suggestions=suggestions)

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
        self, text: str, author: str | None = None, scope: str | dict | Any | None = None
    ) -> None:
        """Delete text with tracked changes.

        This method searches for the specified text in the document and marks
        it as a tracked deletion.

        Args:
            text: The text to delete
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})

        Raises:
            TextNotFoundError: If the text is not found
            AmbiguousTextError: If multiple occurrences of text are found
        """
        # Get all paragraphs in the document
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search for the text to delete
        matches = self._text_search.find_text(text, paragraphs)

        if not matches:
            # Generate smart suggestions
            suggestions = SuggestionGenerator.generate_suggestions(text, paragraphs)
            raise TextNotFoundError(text, suggestions=suggestions)

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
        self,
        find: str,
        replace: str,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> None:
        """Find and replace text with tracked changes.

        This method searches for text and replaces it with new text, showing
        both the deletion of the old text and insertion of the new text as
        tracked changes.

        Args:
            find: Text to find
            replace: Replacement text
            author: Optional author override (uses document author if None)
            scope: Limit search scope (None=all, str="text", dict={"contains": "text"})

        Raises:
            TextNotFoundError: If the 'find' text is not found
            AmbiguousTextError: If multiple occurrences of 'find' text are found
        """
        # Get all paragraphs in the document
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Search for the text to replace
        matches = self._text_search.find_text(find, paragraphs)

        if not matches:
            # Generate smart suggestions
            suggestions = SuggestionGenerator.generate_suggestions(find, paragraphs)
            raise TextNotFoundError(find, suggestions=suggestions)

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

    def insert_paragraph(
        self,
        text: str,
        after: str | None = None,
        before: str | None = None,
        style: str | None = None,
        track: bool = True,
        author: str | None = None,
        scope: str | dict | Any | None = None,
    ) -> "Paragraph":
        """Insert a complete new paragraph with tracked changes.

        Args:
            text: Text content for the new paragraph
            after: Text to search for as insertion point (insert after this)
            before: Text to search for as insertion point (insert before this)
            style: Paragraph style (e.g., 'Normal', 'Heading1')
            track: Whether to track this insertion (default True)
            author: Optional author override (uses document author if None)
            scope: Limit search scope for anchor text

        Returns:
            The created Paragraph object

        Raises:
            ValueError: If neither 'after' nor 'before' is specified, or both are
            TextNotFoundError: If anchor text is not found
            AmbiguousTextError: If multiple occurrences of anchor text are found
        """
        from docx_redline.models.paragraph import Paragraph

        # Validate arguments
        if after is None and before is None:
            raise ValueError("Must specify either 'after' or 'before'")
        if after is not None and before is not None:
            raise ValueError("Cannot specify both 'after' and 'before'")

        anchor_text = after if after is not None else before
        insert_after = after is not None

        # After validation, anchor_text is guaranteed to be a string
        assert anchor_text is not None

        # Get all paragraphs
        all_paragraphs = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Apply scope filter if specified
        paragraphs = ScopeEvaluator.filter_paragraphs(all_paragraphs, scope)

        # Find the anchor paragraph
        matches = self._text_search.find_text(anchor_text, paragraphs)

        if not matches:
            suggestions = SuggestionGenerator.generate_suggestions(anchor_text, paragraphs)
            raise TextNotFoundError(anchor_text, suggestions=suggestions)

        if len(matches) > 1:
            raise AmbiguousTextError(anchor_text, matches)

        match = matches[0]
        anchor_paragraph = match.paragraph

        # Create new paragraph element
        new_p = etree.Element(f"{{{WORD_NAMESPACE}}}p")

        # Add style if specified
        if style:
            p_pr = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}pPr")
            p_style = etree.SubElement(p_pr, f"{{{WORD_NAMESPACE}}}pStyle")
            p_style.set(f"{{{WORD_NAMESPACE}}}val", style)

        # Add text content
        run = etree.SubElement(new_p, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = text

        # If tracked, wrap the paragraph in w:ins
        if track:
            from datetime import datetime, timezone

            author_name = author if author is not None else self.author
            timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            change_id = self._xml_generator.next_change_id
            self._xml_generator.next_change_id += 1

            # Create w:ins element to wrap the paragraph
            ins = etree.Element(f"{{{WORD_NAMESPACE}}}ins")
            ins.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
            ins.set(f"{{{WORD_NAMESPACE}}}author", author_name)
            ins.set(f"{{{WORD_NAMESPACE}}}date", timestamp)
            ins.set(
                "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc",
                timestamp,
            )

            # Move the paragraph inside the w:ins element
            ins.append(new_p)
            element_to_insert = ins
        else:
            element_to_insert = new_p

        # Insert the paragraph in the document
        parent = anchor_paragraph.getparent()
        if parent is None:
            raise ValueError("Anchor paragraph has no parent")

        anchor_index = list(parent).index(anchor_paragraph)

        if insert_after:
            # Insert after anchor
            parent.insert(anchor_index + 1, element_to_insert)
        else:
            # Insert before anchor
            parent.insert(anchor_index, element_to_insert)

        # Return Paragraph wrapper
        # new_p is always the actual paragraph element (whether tracked or not)
        return Paragraph(new_p)

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

    def _replace_match_with_elements(self, match: Any, replacement_elements: list[Any]) -> None:
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

    def accept_all_changes(self) -> None:
        """Accept all tracked changes in the document.

        This removes all tracked change markup:
        - <w:ins> elements are unwrapped (content kept, wrapper removed)
        - <w:del> elements are completely removed (deleted content discarded)

        This is typically used as a preprocessing step before making new edits.
        """
        # Find all insertion elements
        insertions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))

        for ins in insertions:
            parent = ins.getparent()
            if parent is None:
                continue

            # Get the position of the insertion element
            ins_index = list(parent).index(ins)

            # Move all children of <w:ins> to its parent
            for child in list(ins):
                parent.insert(ins_index, child)
                ins_index += 1

            # Remove the <w:ins> wrapper
            parent.remove(ins)

        # Find and remove all deletion elements
        deletions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))

        for del_elem in deletions:
            parent = del_elem.getparent()
            if parent is not None:
                parent.remove(del_elem)

    def delete_all_comments(self) -> None:
        """Delete all comments from the document.

        This removes all comment-related elements:
        - <w:commentRangeStart> - Comment range start markers
        - <w:commentRangeEnd> - Comment range end markers
        - <w:commentReference> - Comment reference markers
        - Runs containing comment references

        This is typically used as a preprocessing step before making new edits.
        """
        # Remove comment range markers
        for elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeStart")):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

        for elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentRangeEnd")):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)

        # Remove comment references
        # Comment references are typically in their own runs, so we'll remove the whole run
        for elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}commentReference")):
            parent = elem.getparent()
            if parent is not None:
                # If parent is a run, remove the run
                if parent.tag == f"{{{WORD_NAMESPACE}}}r":
                    grandparent = parent.getparent()
                    if grandparent is not None:
                        grandparent.remove(parent)
                else:
                    # Otherwise just remove the reference
                    parent.remove(elem)

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
                with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zip_ref:
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

    def apply_edits(
        self, edits: list[dict[str, Any]], stop_on_error: bool = False
    ) -> list[EditResult]:
        """Apply multiple edits in sequence.

        This method processes a list of edit specifications and applies each one
        in order. Each edit is a dictionary specifying the edit type and parameters.

        Args:
            edits: List of edit dictionaries with keys:
                - type: Edit operation ("insert_tracked", "replace_tracked", "delete_tracked")
                - Other parameters specific to the edit type
            stop_on_error: If True, stop processing on first error

        Returns:
            List of EditResult objects, one per edit

        Example:
            >>> edits = [
            ...     {
            ...         "type": "insert_tracked",
            ...         "text": "new text",
            ...         "after": "anchor",
            ...         "scope": "section:Introduction"
            ...     },
            ...     {
            ...         "type": "replace_tracked",
            ...         "find": "old",
            ...         "replace": "new"
            ...     }
            ... ]
            >>> results = doc.apply_edits(edits)
            >>> print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
        """
        results = []

        for i, edit in enumerate(edits):
            edit_type = edit.get("type")
            if not edit_type:
                results.append(
                    EditResult(
                        success=False,
                        edit_type="unknown",
                        message=f"Edit {i}: Missing 'type' field",
                        error=ValidationError("Missing 'type' field"),
                    )
                )
                if stop_on_error:
                    break
                continue

            try:
                result = self._apply_single_edit(edit_type, edit)
                results.append(result)

                if not result.success and stop_on_error:
                    break

            except Exception as e:
                results.append(
                    EditResult(
                        success=False,
                        edit_type=edit_type,
                        message=f"Error: {str(e)}",
                        error=e,
                    )
                )
                if stop_on_error:
                    break

        return results

    def _apply_single_edit(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Apply a single edit operation.

        Args:
            edit_type: The type of edit to perform
            edit: Dictionary with edit parameters

        Returns:
            EditResult indicating success or failure
        """
        try:
            if edit_type == "insert_tracked":
                text = edit.get("text")
                after = edit.get("after")
                author = edit.get("author")
                scope = edit.get("scope")

                if not text or not after:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'text' or 'after'",
                        error=ValidationError("Missing required parameter"),
                    )

                self.insert_tracked(text, after, author=author, scope=scope)
                return EditResult(
                    success=True,
                    edit_type=edit_type,
                    message=f"Inserted '{text}' after '{after}'",
                )

            elif edit_type == "delete_tracked":
                text = edit.get("text")
                author = edit.get("author")
                scope = edit.get("scope")

                if not text:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'text'",
                        error=ValidationError("Missing required parameter"),
                    )

                self.delete_tracked(text, author=author, scope=scope)
                return EditResult(success=True, edit_type=edit_type, message=f"Deleted '{text}'")

            elif edit_type == "replace_tracked":
                find = edit.get("find")
                replace = edit.get("replace")
                author = edit.get("author")
                scope = edit.get("scope")

                if not find or replace is None:
                    return EditResult(
                        success=False,
                        edit_type=edit_type,
                        message="Missing required parameter: 'find' or 'replace'",
                        error=ValidationError("Missing required parameter"),
                    )

                self.replace_tracked(find, replace, author=author, scope=scope)
                return EditResult(
                    success=True,
                    edit_type=edit_type,
                    message=f"Replaced '{find}' with '{replace}'",
                )

            else:
                return EditResult(
                    success=False,
                    edit_type=edit_type,
                    message=f"Unknown edit type: {edit_type}",
                    error=ValidationError(f"Unknown edit type: {edit_type}"),
                )

        except TextNotFoundError as e:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Text not found: {e}",
                error=e,
            )
        except AmbiguousTextError as e:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Ambiguous text: {e}",
                error=e,
            )
        except Exception as e:
            return EditResult(
                success=False, edit_type=edit_type, message=f"Error: {str(e)}", error=e
            )

    def apply_edit_file(
        self, path: str | Path, format: str = "yaml", stop_on_error: bool = False
    ) -> list[EditResult]:
        """Apply edits from a YAML or JSON file.

        Loads edit specifications from a file and applies them using apply_edits().
        The file should contain an 'edits' key with a list of edit dictionaries.

        Args:
            path: Path to the edit specification file
            format: File format - "yaml" or "json" (default: "yaml")
            stop_on_error: If True, stop processing on first error

        Returns:
            List of EditResult objects, one per edit

        Raises:
            ValidationError: If file cannot be parsed or has invalid format
            FileNotFoundError: If file does not exist

        Example YAML file:
            ```yaml
            edits:
              - type: insert_tracked
                text: "new text"
                after: "anchor"
              - type: replace_tracked
                find: "old"
                replace: "new"
            ```

        Example:
            >>> results = doc.apply_edit_file("edits.yaml")
            >>> print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
        """
        file_path = Path(path)

        if not file_path.exists():
            raise FileNotFoundError(f"Edit file not found: {path}")

        try:
            with open(file_path, encoding="utf-8") as f:
                if format == "yaml":
                    data = yaml.safe_load(f)
                elif format == "json":
                    import json

                    data = json.load(f)
                else:
                    raise ValidationError(f"Unsupported format: {format}")

            if not isinstance(data, dict):
                raise ValidationError("Edit file must contain a dictionary/object")

            if "edits" not in data:
                raise ValidationError("Edit file must contain an 'edits' key")

            edits = data["edits"]
            if not isinstance(edits, list):
                raise ValidationError("'edits' must be a list")

            # Apply the edits
            return self.apply_edits(edits, stop_on_error=stop_on_error)

        except yaml.YAMLError as e:
            raise ValidationError(f"Failed to parse YAML file: {e}") from e
        except Exception as e:
            if isinstance(e, ValidationError | FileNotFoundError):
                raise
            raise ValidationError(f"Failed to load edit file: {e}") from e

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
