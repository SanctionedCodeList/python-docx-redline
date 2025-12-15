"""Change management operations for tracked changes.

This module provides the ChangeManagement class for accepting and rejecting
tracked changes in Word documents.
"""

from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from ..document import Document

# Word namespace
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class ChangeManagement:
    """Handles accepting and rejecting tracked changes.

    This class extracts change management operations from the Document class
    to improve separation of concerns and maintainability.

    Attributes:
        _document: Reference to the parent Document instance
    """

    def __init__(self, document: Document) -> None:
        """Initialize the ChangeManagement operations.

        Args:
            document: The parent Document instance
        """
        self._document = document

    @property
    def xml_root(self) -> Any:
        """Get the XML root element from the document."""
        return self._document.xml_root

    # Helper methods

    def _unwrap_element(self, element: Any) -> None:
        """Unwrap an element by moving its children to its parent.

        Args:
            element: The element to unwrap
        """
        parent = element.getparent()
        if parent is None:
            return

        # Get the position of the element
        elem_index = list(parent).index(element)

        # Move all children to parent
        for child in list(element):
            parent.insert(elem_index, child)
            elem_index += 1

        # Remove the wrapper
        parent.remove(element)

    def _unwrap_deletion(self, del_elem: Any) -> None:
        """Unwrap a deletion element, converting w:delText back to w:t.

        When rejecting a deletion, we need to restore the deleted text by:
        1. Converting all <w:delText> elements to <w:t>
        2. Unwrapping the <w:del> element

        Args:
            del_elem: The <w:del> element to unwrap
        """
        # First, convert all w:delText to w:t within this deletion
        for deltext in del_elem.iter(f"{{{WORD_NAMESPACE}}}delText"):
            deltext.tag = f"{{{WORD_NAMESPACE}}}t"

        # Then unwrap the deletion element
        self._unwrap_element(del_elem)

    def _remove_element(self, element: Any) -> None:
        """Remove an element from its parent.

        Args:
            element: The element to remove
        """
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)

    # Accept/Reject all changes

    def accept_all(self) -> None:
        """Accept all tracked changes in the document.

        This removes all tracked change markup:
        - <w:ins> elements are unwrapped (content kept, wrapper removed)
        - <w:del> elements are completely removed (deleted content discarded)
        - <w:rPrChange> elements are removed (current formatting kept)
        - <w:pPrChange> elements are removed (current formatting kept)

        This is typically used as a preprocessing step before making new edits.
        """
        self.accept_insertions()
        self.accept_deletions()
        self.accept_format_changes()

    def reject_all(self) -> None:
        """Reject all tracked changes in the document.

        This removes all tracked change markup by reverting to previous state:
        - <w:ins> elements and their content are removed
        - <w:del> elements are unwrapped (deleted content restored)
        - <w:rPrChange> elements restore previous formatting
        - <w:pPrChange> elements restore previous formatting
        """
        self.reject_insertions()
        self.reject_deletions()
        self.reject_format_changes()

    # Accept/Reject by type

    def accept_insertions(self) -> int:
        """Accept all tracked insertions in the document.

        Unwraps all <w:ins> elements, keeping the inserted content.

        Returns:
            Number of insertions accepted
        """
        insertions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))
        for ins in insertions:
            self._unwrap_element(ins)
        return len(insertions)

    def reject_insertions(self) -> int:
        """Reject all tracked insertions in the document.

        Removes all <w:ins> elements and their content.

        Returns:
            Number of insertions rejected
        """
        insertions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"))
        for ins in insertions:
            self._remove_element(ins)
        return len(insertions)

    def accept_deletions(self) -> int:
        """Accept all tracked deletions in the document.

        Removes all <w:del> elements (keeps text deleted).

        Returns:
            Number of deletions accepted
        """
        deletions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
        for del_elem in deletions:
            self._remove_element(del_elem)
        return len(deletions)

    def reject_deletions(self) -> int:
        """Reject all tracked deletions in the document.

        Unwraps all <w:del> elements, restoring the deleted content.
        Converts w:delText back to w:t.

        Returns:
            Number of deletions rejected
        """
        deletions = list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"))
        for del_elem in deletions:
            self._unwrap_deletion(del_elem)
        return len(deletions)

    def accept_format_changes(self) -> int:
        """Accept all tracked formatting changes in the document.

        Removes all <w:rPrChange> and <w:pPrChange> elements, keeping the
        current formatting as-is.

        Returns:
            Number of formatting changes accepted
        """
        count = 0

        # Accept run property changes (character formatting)
        for rpr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange")):
            parent = rpr_change.getparent()
            if parent is not None:
                parent.remove(rpr_change)
                count += 1

        # Accept paragraph property changes
        for ppr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange")):
            parent = ppr_change.getparent()
            if parent is not None:
                parent.remove(ppr_change)
                count += 1

        return count

    def reject_format_changes(self) -> int:
        """Reject all tracked formatting changes in the document.

        Restores the previous formatting from <w:rPrChange> and <w:pPrChange>
        elements, then removes the change tracking elements.

        Returns:
            Number of formatting changes rejected
        """
        count = 0

        # Reject run property changes - restore previous formatting
        for rpr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange")):
            parent_rpr = rpr_change.getparent()
            if parent_rpr is None:
                continue

            # Get the previous rPr from inside the change element
            previous_rpr = rpr_change.find(f"{{{WORD_NAMESPACE}}}rPr")

            # Replace parent's children with previous state (except rPrChange)
            # First, collect children to remove (excluding the change element)
            children_to_remove = [
                child for child in list(parent_rpr) if child.tag != f"{{{WORD_NAMESPACE}}}rPrChange"
            ]
            for child in children_to_remove:
                parent_rpr.remove(child)

            # Remove the rPrChange element
            parent_rpr.remove(rpr_change)

            # Add back the previous properties
            if previous_rpr is not None:
                for child in previous_rpr:
                    parent_rpr.append(deepcopy(child))

            count += 1

        # Reject paragraph property changes - restore previous formatting
        for ppr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange")):
            parent_ppr = ppr_change.getparent()
            if parent_ppr is None:
                continue

            # Get the previous pPr from inside the change element
            previous_ppr = ppr_change.find(f"{{{WORD_NAMESPACE}}}pPr")

            # Replace parent's children with previous state (except pPrChange and rPr)
            # Note: rPr inside pPr is for paragraph mark formatting, keep it separate
            children_to_remove = [
                child
                for child in list(parent_ppr)
                if child.tag
                not in (
                    f"{{{WORD_NAMESPACE}}}pPrChange",
                    f"{{{WORD_NAMESPACE}}}rPr",
                )
            ]
            for child in children_to_remove:
                parent_ppr.remove(child)

            # Remove the pPrChange element
            parent_ppr.remove(ppr_change)

            # Add back the previous properties (insert at beginning, before rPr)
            if previous_ppr is not None:
                insert_idx = 0
                for child in previous_ppr:
                    parent_ppr.insert(insert_idx, deepcopy(child))
                    insert_idx += 1

            count += 1

        return count

    # Accept/Reject by change ID

    def accept_change(self, change_id: str | int) -> None:
        """Accept a specific tracked change by its ID.

        Args:
            change_id: The change ID (w:id attribute value)

        Raises:
            ValueError: If no change with the given ID is found

        Example:
            >>> doc.accept_change("5")
        """
        change_id_str = str(change_id)

        # Search for insertion with this ID
        for ins in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"):
            if ins.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                self._unwrap_element(ins)
                return

        # Search for deletion with this ID
        for del_elem in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"):
            if del_elem.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                self._remove_element(del_elem)
                return

        # Search for run property change with this ID
        for rpr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange"):
            if rpr_change.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                parent = rpr_change.getparent()
                if parent is not None:
                    parent.remove(rpr_change)
                return

        # Search for paragraph property change with this ID
        for ppr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange"):
            if ppr_change.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                parent = ppr_change.getparent()
                if parent is not None:
                    parent.remove(ppr_change)
                return

        raise ValueError(f"No tracked change found with ID: {change_id}")

    def reject_change(self, change_id: str | int) -> None:
        """Reject a specific tracked change by its ID.

        Args:
            change_id: The change ID (w:id attribute value)

        Raises:
            ValueError: If no change with the given ID is found

        Example:
            >>> doc.reject_change("5")
        """
        change_id_str = str(change_id)

        # Search for insertion with this ID
        for ins in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins"):
            if ins.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                self._remove_element(ins)
                return

        # Search for deletion with this ID
        for del_elem in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del"):
            if del_elem.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                self._unwrap_deletion(del_elem)
                return

        # Search for run property change with this ID
        for rpr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange"):
            if rpr_change.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                parent_rpr = rpr_change.getparent()
                if parent_rpr is not None:
                    previous_rpr = rpr_change.find(f"{{{WORD_NAMESPACE}}}rPr")
                    # Remove current properties (except rPrChange)
                    for child in list(parent_rpr):
                        if child.tag != f"{{{WORD_NAMESPACE}}}rPrChange":
                            parent_rpr.remove(child)
                    parent_rpr.remove(rpr_change)
                    # Restore previous
                    if previous_rpr is not None:
                        for child in previous_rpr:
                            parent_rpr.append(deepcopy(child))
                return

        # Search for paragraph property change with this ID
        for ppr_change in self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange"):
            if ppr_change.get(f"{{{WORD_NAMESPACE}}}id") == change_id_str:
                parent_ppr = ppr_change.getparent()
                if parent_ppr is not None:
                    previous_ppr = ppr_change.find(f"{{{WORD_NAMESPACE}}}pPr")
                    # Remove current properties (except pPrChange and rPr)
                    for child in list(parent_ppr):
                        if child.tag not in (
                            f"{{{WORD_NAMESPACE}}}pPrChange",
                            f"{{{WORD_NAMESPACE}}}rPr",
                        ):
                            parent_ppr.remove(child)
                    parent_ppr.remove(ppr_change)
                    # Restore previous
                    if previous_ppr is not None:
                        insert_idx = 0
                        for child in previous_ppr:
                            parent_ppr.insert(insert_idx, deepcopy(child))
                            insert_idx += 1
                return

        raise ValueError(f"No tracked change found with ID: {change_id}")

    # Accept/Reject by author

    def accept_by_author(self, author: str) -> int:
        """Accept all tracked changes by a specific author.

        Args:
            author: The author name (w:author attribute value)

        Returns:
            Number of changes accepted

        Example:
            >>> count = doc.accept_by_author("John Doe")
            >>> print(f"Accepted {count} changes from John Doe")
        """
        count = 0

        # Accept insertions by this author
        for ins in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins")):
            if ins.get(f"{{{WORD_NAMESPACE}}}author") == author:
                self._unwrap_element(ins)
                count += 1

        # Accept deletions by this author
        for del_elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del")):
            if del_elem.get(f"{{{WORD_NAMESPACE}}}author") == author:
                self._remove_element(del_elem)
                count += 1

        # Accept format changes by this author (remove rPrChange/pPrChange)
        for rpr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange")):
            if rpr_change.get(f"{{{WORD_NAMESPACE}}}author") == author:
                parent = rpr_change.getparent()
                if parent is not None:
                    parent.remove(rpr_change)
                count += 1

        for ppr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange")):
            if ppr_change.get(f"{{{WORD_NAMESPACE}}}author") == author:
                parent = ppr_change.getparent()
                if parent is not None:
                    parent.remove(ppr_change)
                count += 1

        return count

    def reject_by_author(self, author: str) -> int:
        """Reject all tracked changes by a specific author.

        Args:
            author: The author name (w:author attribute value)

        Returns:
            Number of changes rejected

        Example:
            >>> count = doc.reject_by_author("John Doe")
            >>> print(f"Rejected {count} changes from John Doe")
        """
        count = 0

        # Reject insertions by this author
        for ins in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}ins")):
            if ins.get(f"{{{WORD_NAMESPACE}}}author") == author:
                self._remove_element(ins)
                count += 1

        # Reject deletions by this author (convert delText back to t)
        for del_elem in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}del")):
            if del_elem.get(f"{{{WORD_NAMESPACE}}}author") == author:
                self._unwrap_deletion(del_elem)
                count += 1

        # Reject format changes by this author (restore previous formatting)
        for rpr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}rPrChange")):
            if rpr_change.get(f"{{{WORD_NAMESPACE}}}author") == author:
                parent_rpr = rpr_change.getparent()
                if parent_rpr is not None:
                    # Get the previous rPr from inside the change element
                    prev_rpr = rpr_change.find(f"{{{WORD_NAMESPACE}}}rPr")
                    if prev_rpr is not None:
                        # Replace current rPr content with previous
                        # Clear all children except rPrChange
                        for child in list(parent_rpr):
                            if child.tag != f"{{{WORD_NAMESPACE}}}rPrChange":
                                parent_rpr.remove(child)
                        # Copy previous formatting back using deepcopy and append
                        for child in list(prev_rpr):
                            parent_rpr.append(deepcopy(child))
                    # Remove the change tracking element
                    parent_rpr.remove(rpr_change)
                count += 1

        for ppr_change in list(self.xml_root.iter(f"{{{WORD_NAMESPACE}}}pPrChange")):
            if ppr_change.get(f"{{{WORD_NAMESPACE}}}author") == author:
                parent_ppr = ppr_change.getparent()
                if parent_ppr is not None:
                    # Get the previous pPr from inside the change element
                    prev_ppr = ppr_change.find(f"{{{WORD_NAMESPACE}}}pPr")
                    if prev_ppr is not None:
                        # Replace current pPr content with previous
                        # Clear all children except pPrChange and rPr
                        for child in list(parent_ppr):
                            if child.tag not in (
                                f"{{{WORD_NAMESPACE}}}pPrChange",
                                f"{{{WORD_NAMESPACE}}}rPr",
                            ):
                                parent_ppr.remove(child)
                        # Copy previous formatting back using deepcopy and append
                        for child in list(prev_ppr):
                            parent_ppr.append(deepcopy(child))
                    # Remove the change tracking element
                    parent_ppr.remove(ppr_change)
                count += 1

        return count

    # Bulk accept/reject with filters

    def accept_changes(
        self,
        change_type: str | None = None,
        author: str | None = None,
    ) -> int:
        """Accept multiple tracked changes matching the given criteria.

        This is a bulk operation that accepts all changes matching the filters.
        If no filters are provided, accepts ALL tracked changes.

        Args:
            change_type: Optional filter by change type. Valid values:
                         "insertion", "deletion", "format_run", "format_paragraph",
                         or None for all types.
            author: Optional filter by author name.

        Returns:
            Number of changes accepted.

        Example:
            >>> # Accept all insertions
            >>> count = doc.accept_changes(change_type="insertion")
            >>> print(f"Accepted {count} insertions")
            >>>
            >>> # Accept all changes by a specific author
            >>> count = doc.accept_changes(author="Legal Team")
            >>> print(f"Accepted {count} changes from Legal Team")
            >>>
            >>> # Accept only insertions by a specific author
            >>> count = doc.accept_changes(change_type="insertion", author="John Doe")
        """
        # Special case: no filters = accept all
        if change_type is None and author is None:
            self.accept_all()
            return 0  # Can't determine count after the fact

        count = 0

        # If only author filter, use existing method for efficiency
        if change_type is None and author is not None:
            return self.accept_by_author(author)

        # Otherwise, get filtered changes and accept each
        changes = self._document.get_tracked_changes(change_type=change_type, author=author)

        # Accept changes in reverse order to handle nested elements properly
        for change in reversed(changes):
            try:
                self.accept_change(change.id)
                count += 1
            except ValueError:
                # Change may have already been accepted (e.g., nested structure)
                pass

        return count

    def reject_changes(
        self,
        change_type: str | None = None,
        author: str | None = None,
    ) -> int:
        """Reject multiple tracked changes matching the given criteria.

        This is a bulk operation that rejects all changes matching the filters.
        If no filters are provided, rejects ALL tracked changes.

        Args:
            change_type: Optional filter by change type. Valid values:
                         "insertion", "deletion", "format_run", "format_paragraph",
                         or None for all types.
            author: Optional filter by author name.

        Returns:
            Number of changes rejected.

        Example:
            >>> # Reject all deletions
            >>> count = doc.reject_changes(change_type="deletion")
            >>> print(f"Rejected {count} deletions")
            >>>
            >>> # Reject all changes by a specific author
            >>> count = doc.reject_changes(author="Unauthorized User")
            >>> print(f"Rejected {count} changes from Unauthorized User")
            >>>
            >>> # Reject only deletions by a specific author
            >>> count = doc.reject_changes(change_type="deletion", author="John Doe")
        """
        # Special case: no filters = reject all
        if change_type is None and author is None:
            self.reject_all()
            return 0  # Can't determine count after the fact

        count = 0

        # If only author filter, use existing method for efficiency
        if change_type is None and author is not None:
            return self.reject_by_author(author)

        # Otherwise, get filtered changes and reject each
        changes = self._document.get_tracked_changes(change_type=change_type, author=author)

        # Reject changes in reverse order to handle nested elements properly
        for change in reversed(changes):
            try:
                self.reject_change(change.id)
                count += 1
            except ValueError:
                # Change may have already been rejected (e.g., nested structure)
                pass

        return count
