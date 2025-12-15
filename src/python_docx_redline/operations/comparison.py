"""
ComparisonOperations class for comparing documents and generating tracked changes.

This module provides a dedicated class for document comparison operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

import difflib
import logging
from typing import TYPE_CHECKING, Any

from lxml import etree

from ..constants import WORD_NAMESPACE
from ..minimal_diff import (
    apply_minimal_edits_to_paragraph,
    should_use_minimal_editing,
)

if TYPE_CHECKING:
    from ..document import Document

logger = logging.getLogger(__name__)


class ComparisonOperations:
    """Handles document comparison operations.

    This class encapsulates all document comparison functionality, including:
    - Comparing two documents to generate tracked changes
    - Marking paragraphs as deleted
    - Inserting comparison paragraphs

    The class takes a Document reference and operates on its XML structure.

    Example:
        >>> # Usually accessed through Document
        >>> original = Document("contract_v1.docx")
        >>> modified = Document("contract_v2.docx")
        >>> num_changes = original.compare_to(modified)
    """

    def __init__(self, document: Document) -> None:
        """Initialize ComparisonOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    def compare_to(
        self,
        modified: Document,
        author: str | None = None,
        minimal_edits: bool = False,
    ) -> int:
        """Generate tracked changes by comparing this document to a modified version.

        This method compares the current document (original) to a modified document
        and generates tracked changes showing what was added, removed, or changed.
        The changes are applied to this document.

        The comparison operates at the paragraph level:
        - Paragraphs in modified but not in original → tracked insertions
        - Paragraphs in original but not in modified → tracked deletions
        - Paragraphs that changed → tracked deletion of old + insertion of new

        Args:
            modified: The modified Document to compare against
            author: Author name for the tracked changes (uses document default if None)
            minimal_edits: If True, use word-level diffs for 1:1 paragraph replacements
                instead of deleting/inserting entire paragraphs. This produces
                legal-style redlines where only the changed words are marked.
                (default: False)

        Returns:
            Number of changes made (insertions + deletions)

        Example:
            >>> original = Document("contract_v1.docx")
            >>> modified = Document("contract_v2.docx")
            >>> num_changes = original.compare_to(modified)
            >>> original.save("contract_redlined.docx")
            >>> print(f"Found {num_changes} changes")

            # For legal-style minimal diffs:
            >>> num_changes = original.compare_to(modified, minimal_edits=True)

        Note:
            - This modifies the current document in place
            - The comparison uses paragraph text content
            - Formatting changes within paragraphs are not tracked separately
            - For best results, compare documents with similar structure
            - When minimal_edits=True, whitespace-only changes are suppressed
              for readability, and paragraphs with existing tracked changes
              fall back to coarse replacement
        """
        # Get paragraph texts from both documents
        original_texts = [p.text for p in self._document.paragraphs]
        modified_texts = [p.text for p in modified.paragraphs]

        # Use SequenceMatcher to find differences at paragraph level
        matcher = difflib.SequenceMatcher(None, original_texts, modified_texts)
        opcodes = matcher.get_opcodes()

        # We need to process changes carefully to avoid index shifting issues
        # Build a list of operations to apply
        operations: list[dict[str, Any]] = []

        for tag, i1, i2, j1, j2 in opcodes:
            if tag == "equal":
                # No change needed
                continue
            elif tag == "delete":
                # Paragraphs removed in modified version
                for idx in range(i1, i2):
                    operations.append(
                        {
                            "type": "delete",
                            "original_index": idx,
                            "text": original_texts[idx],
                        }
                    )
            elif tag == "insert":
                # Paragraphs added in modified version
                # Insert after the previous paragraph (i1-1) or at beginning
                for j_idx in range(j1, j2):
                    operations.append(
                        {
                            "type": "insert",
                            "insert_after_index": i1 - 1,  # -1 means insert at beginning
                            "text": modified_texts[j_idx],
                            "modified_index": j_idx,
                        }
                    )
            elif tag == "replace":
                # Paragraphs changed
                # Check if this is a 1:1 replacement and minimal_edits is enabled
                is_one_to_one = (i2 - i1) == 1 and (j2 - j1) == 1

                if minimal_edits and is_one_to_one:
                    # Attempt minimal intra-paragraph edit for 1:1 replacement
                    operations.append(
                        {
                            "type": "minimal_replace",
                            "original_index": i1,
                            "original_text": original_texts[i1],
                            "new_text": modified_texts[j1],
                        }
                    )
                else:
                    # Fall back to coarse delete + insert
                    # First mark deletions
                    for idx in range(i1, i2):
                        operations.append(
                            {
                                "type": "delete",
                                "original_index": idx,
                                "text": original_texts[idx],
                            }
                        )
                    # Then mark insertions
                    for j_idx in range(j1, j2):
                        operations.append(
                            {
                                "type": "insert",
                                "insert_after_index": i1 - 1,
                                "text": modified_texts[j_idx],
                                "modified_index": j_idx,
                            }
                        )

        # Apply operations to the document
        change_count = self._apply_comparison_changes(operations, author, minimal_edits)

        return change_count

    def _apply_comparison_changes(
        self,
        operations: list[dict[str, Any]],
        author: str | None,
        minimal_edits: bool = False,
    ) -> int:
        """Apply comparison operations to generate tracked changes.

        Args:
            operations: List of delete/insert/minimal_replace operations from compare_to()
            author: Author for tracked changes
            minimal_edits: Whether minimal edits mode is enabled

        Returns:
            Number of changes applied
        """
        change_count = 0

        # Get all paragraph elements
        body = self._document.xml_root.find(f"{{{WORD_NAMESPACE}}}body")
        if body is None:
            return 0

        paragraphs = list(body.findall(f"{{{WORD_NAMESPACE}}}p"))

        # Track which paragraphs have been marked as deleted
        deleted_indices: set[int] = set()

        # Track which paragraphs have been handled by minimal_replace
        minimal_replace_indices: set[int] = set()

        # Process minimal replacements first
        for op in operations:
            if op["type"] == "minimal_replace":
                idx = op["original_index"]
                if idx < len(paragraphs) and idx not in minimal_replace_indices:
                    para_elem = paragraphs[idx]
                    orig_text = op["original_text"]
                    new_text = op["new_text"]

                    # Check if minimal editing is viable for this paragraph
                    use_minimal, diff_result, reason = should_use_minimal_editing(
                        para_elem, new_text, orig_text
                    )

                    if use_minimal and diff_result.hunks:
                        # Apply minimal edits
                        apply_minimal_edits_to_paragraph(
                            para_elem,
                            diff_result.hunks,
                            self._document._xml_generator,
                            author,
                        )
                        minimal_replace_indices.add(idx)
                        # Count changes consistently with coarse mode:
                        # Each hunk with delete_text counts as 1 deletion
                        # Each hunk with insert_text counts as 1 insertion
                        for hunk in diff_result.hunks:
                            if hunk.delete_text:
                                change_count += 1
                            if hunk.insert_text:
                                change_count += 1
                    elif not use_minimal:
                        # Fall back to coarse replacement
                        if reason:
                            logger.debug(
                                "Minimal editing disabled for paragraph %d: %s",
                                idx,
                                reason,
                            )
                        self._mark_paragraph_deleted(para_elem, author)
                        deleted_indices.add(idx)
                        change_count += 1

                        # Insert new paragraph after the deleted one
                        self._insert_comparison_paragraph(body, paragraphs, idx, new_text, author)
                        change_count += 1
                    # else: diff_result.hunks is empty (whitespace-only), no changes needed

        # Process deletions (mark content as deleted)
        for op in operations:
            if op["type"] == "delete":
                idx = op["original_index"]
                if (
                    idx < len(paragraphs)
                    and idx not in deleted_indices
                    and idx not in minimal_replace_indices
                ):
                    self._mark_paragraph_deleted(paragraphs[idx], author)
                    deleted_indices.add(idx)
                    change_count += 1

        # Process insertions
        # We need to track offset for insertions
        insertions_by_position: dict[int, list[str]] = {}
        for op in operations:
            if op["type"] == "insert":
                pos = op["insert_after_index"]
                if pos not in insertions_by_position:
                    insertions_by_position[pos] = []
                insertions_by_position[pos].append(op["text"])

        # Apply insertions in reverse order of position to avoid index shifting
        for pos in sorted(insertions_by_position.keys(), reverse=True):
            texts = insertions_by_position[pos]
            for text in reversed(texts):
                self._insert_comparison_paragraph(body, paragraphs, pos, text, author)
                change_count += 1

        return change_count

    def _mark_paragraph_deleted(
        self,
        paragraph: Any,
        author: str | None,
    ) -> None:
        """Mark all text in a paragraph as deleted with tracked changes.

        Args:
            paragraph: The paragraph XML element
            author: Author for the deletion
        """
        # Get all runs in the paragraph
        runs = paragraph.findall(f".//{{{WORD_NAMESPACE}}}r")

        for run in runs:
            # Get all text elements in this run
            text_elements = run.findall(f"{{{WORD_NAMESPACE}}}t")

            for t_elem in text_elements:
                text = t_elem.text or ""
                if not text:
                    continue

                # Create deletion XML
                del_xml = self._document._xml_generator.create_deletion(text, author)

                # Parse the deletion XML
                del_elem = etree.fromstring(
                    f'<root xmlns:w="{WORD_NAMESPACE}" '
                    f'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" '
                    f'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
                    f"{del_xml}</root>"
                )

                # Get the w:del element
                del_node = del_elem.find(f"{{{WORD_NAMESPACE}}}del")
                if del_node is not None:
                    # Insert the deletion before the original text element
                    parent = t_elem.getparent()
                    if parent is not None:
                        idx = list(parent).index(t_elem)
                        parent.insert(idx, del_node)
                        # Remove the original text element
                        parent.remove(t_elem)

    def _insert_comparison_paragraph(
        self,
        body: Any,
        paragraphs: list[Any],
        after_index: int,
        text: str,
        author: str | None,
    ) -> None:
        """Insert a new paragraph with tracked insertion.

        Args:
            body: The document body element
            paragraphs: List of existing paragraph elements
            after_index: Index of paragraph to insert after (-1 for beginning)
            text: Text content of the new paragraph
            author: Author for the insertion
        """
        # Create insertion XML
        ins_xml = self._document._xml_generator.create_insertion(text, author)

        # Create a new paragraph with the insertion
        new_para = etree.Element(f"{{{WORD_NAMESPACE}}}p")

        # Parse the insertion XML
        ins_elem = etree.fromstring(
            f'<root xmlns:w="{WORD_NAMESPACE}" '
            f'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" '
            f'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
            f"{ins_xml}</root>"
        )

        # Get the w:ins element
        ins_node = ins_elem.find(f"{{{WORD_NAMESPACE}}}ins")
        if ins_node is not None:
            new_para.append(ins_node)

        # Insert the new paragraph at the appropriate position
        if after_index < 0:
            # Insert at the beginning
            body.insert(0, new_para)
        elif after_index < len(paragraphs):
            # Insert after the specified paragraph
            ref_para = paragraphs[after_index]
            idx = list(body).index(ref_para)
            body.insert(idx + 1, new_para)
        else:
            # Insert at the end
            body.append(new_para)
