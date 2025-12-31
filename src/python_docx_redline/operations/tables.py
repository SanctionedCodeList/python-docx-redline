"""
TableOperations class for handling table manipulation with tracked changes.

This module provides a dedicated class for all table operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from ..author import AuthorIdentity
from ..constants import WORD_NAMESPACE

if TYPE_CHECKING:
    from ..document import Document
    from ..models.table import Table, TableRow
    from ..text_search import TextSpan


class TableOperations:
    """Handles table operations including reading, modifying, and tracked changes.

    This class encapsulates all table functionality, including:
    - Accessing tables in the document
    - Finding tables containing specific text
    - Replacing text in table cells
    - Inserting and deleting rows with tracking
    - Inserting and deleting columns with tracking

    The class takes a Document reference and operates on its XML structure.

    Example:
        >>> # Usually accessed through Document
        >>> doc = Document("contract.docx")
        >>> tables = doc.tables
        >>> table = doc.find_table("Total Price")
        >>> doc.replace_in_table("OLD", "NEW", track=True)
    """

    def __init__(self, document: Document) -> None:
        """Initialize TableOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    @property
    def all(self) -> list[Table]:
        """Get all tables in the document.

        Returns:
            List of Table objects

        Example:
            >>> doc = Document("contract.docx")
            >>> for i, table in enumerate(doc.tables):
            ...     print(f"Table {i}: {table.row_count} rows × {table.col_count} cols")
        """
        from ..models.table import Table

        return [Table(tbl) for tbl in self._document.xml_root.iter(f"{{{WORD_NAMESPACE}}}tbl")]

    def find(self, containing: str, case_sensitive: bool = True) -> Table | None:
        """Find the first table containing specific text.

        Args:
            containing: Text to search for in table cells
            case_sensitive: Whether search should be case sensitive (default: True)

        Returns:
            First Table containing the text, or None if not found

        Example:
            >>> doc = Document("contract.docx")
            >>> pricing_table = doc.find_table("Total Price")
            >>> if pricing_table:
            ...     print(f"Found table with {pricing_table.row_count} rows")
        """
        for table in self.all:
            if table.contains(containing, case_sensitive):
                return table
        return None

    def replace_text(
        self,
        old_text: str,
        new_text: str,
        *,
        table_index: int | None = None,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
        regex: bool = False,
        case_sensitive: bool = True,
        minimal: bool | None = None,
    ) -> int:
        """Replace text in table cells with optional tracked changes.

        Args:
            old_text: Text to find (or regex pattern if regex=True)
            new_text: Replacement text
            table_index: Specific table index, or None for all tables
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)
            regex: Whether old_text is a regex pattern (default: False)
            case_sensitive: Whether search is case sensitive (default: True)
            minimal: If True, use word-level diffing for human-looking tracked changes.
                If False, use coarse delete-all + insert-all. If None (default),
                uses the document's minimal_edits setting. Only applies when track=True.

        Returns:
            Number of replacements made

        Example:
            >>> doc = Document("contract.docx")
            >>> count = doc.replace_in_table("OLD", "NEW", track=True)
            >>> print(f"Replaced {count} occurrences")
        """
        author_name = author if author is not None else self._document.author
        count = 0

        tables = self.all
        if table_index is not None:
            if table_index < 0 or table_index >= len(tables):
                raise IndexError(f"Table index {table_index} out of range (0-{len(tables) - 1})")
            tables = [tables[table_index]]

        # Search and replace in each table
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    # Use TextSearch to find matches in cell paragraphs
                    for para in cell.paragraphs:
                        matches = self._document._text_search.find_text(
                            old_text,
                            [para.element],
                            regex=regex,
                            case_sensitive=case_sensitive,
                        )

                        for match in matches:
                            if track:
                                self._replace_match_tracked(match, new_text, author_name, minimal)
                            else:
                                self._replace_match_untracked(match, new_text)
                            count += 1

        return count

    def _replace_match_tracked(
        self,
        match: TextSpan,
        new_text: str,
        author_name: str | AuthorIdentity,
        minimal: bool | None = None,
    ) -> None:
        """Replace a text match with tracked deletion and insertion.

        Args:
            match: The TextSpan representing the matched text
            new_text: The replacement text
            author_name: Author for the tracked change
            minimal: If True, use word-level diffing. If None, use document default.
        """
        # Extract author string if AuthorIdentity object
        author_str = author_name.author if isinstance(author_name, AuthorIdentity) else author_name

        # Determine effective minimal setting
        use_minimal = minimal if minimal is not None else self._document._minimal_edits

        if use_minimal:
            # Attempt word-level minimal edit
            from ..minimal_diff import apply_minimal_edits_to_textspan

            success, reason = apply_minimal_edits_to_textspan(
                match,
                new_text,
                self._document._xml_generator,
                author_str,
            )
            if success:
                return  # Minimal edit applied successfully
            else:
                # Log fallback at INFO level
                import logging

                logger = logging.getLogger(__name__)
                logger.info(
                    "Table cell: Falling back to coarse tracked change for '%s' -> '%s': %s",
                    match.text[:50],
                    new_text[:50],
                    reason,
                )

        # Coarse tracked replacement
        deletion_xml = self._document._xml_generator.create_deletion(match.text, author_str)
        insertion_xml = self._document._xml_generator.create_insertion(new_text, author_str)

        # Parse XMLs with namespace context
        wrapped_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:w="{WORD_NAMESPACE}"
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
      xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
    {deletion_xml}
    {insertion_xml}
</root>"""
        root = etree.fromstring(wrapped_xml.encode("utf-8"))
        deletion_element = root[0]
        insertion_element = root[1]

        self._document._replace_match_with_elements(match, [deletion_element, insertion_element])

    def _replace_match_untracked(self, match: TextSpan, new_text: str) -> None:
        """Replace a text match without tracking.

        Uses create_plain_runs to properly handle markdown formatting and preserve
        source run formatting.

        Args:
            match: The TextSpan representing the matched text
            new_text: The replacement text (may include markdown formatting)
        """
        # Get source run for formatting inheritance
        source_run = match.runs[match.start_run_index] if match.runs else None

        # Create runs using xml_generator for proper markdown support
        new_runs = self._document._xml_generator.create_plain_runs(new_text, source_run=source_run)

        # Replace matched text with the new runs
        if len(new_runs) == 1:
            self._document._replace_match_with_element(match, new_runs[0])
        else:
            self._document._replace_match_with_elements(match, new_runs)

    def insert_row(
        self,
        after_row: int | str,
        cells: list[str],
        *,
        table_index: int = 0,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
    ) -> TableRow:
        """Insert a new table row with optional tracked changes.

        Args:
            after_row: Row index (int) or text to find in a row (str)
            cells: List of text content for each cell in the new row
            table_index: Which table to modify (default: 0 = first table)
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)

        Returns:
            The newly created TableRow object

        Raises:
            IndexError: If table_index is out of range
            ValueError: If after_row text is not found or is ambiguous
            ValueError: If number of cells doesn't match table column count

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.insert_table_row(
            ...     after_row="Total:",
            ...     cells=["New Item", "$1,000", "$2,000"],
            ...     track=True
            ... )
        """
        from ..models.table import TableRow

        table = self._get_table(table_index)
        insert_after_index = self._resolve_row_index(table, after_row)

        # Validate cell count
        if len(cells) != table.col_count:
            raise ValueError(f"Expected {table.col_count} cells, got {len(cells)}")

        # Create new row element
        new_row = self._create_row_element(cells)

        if track:
            self._add_row_insertion_tracking(new_row, author)

        # Insert after the specified row
        self._insert_row_in_table(table, new_row, insert_after_index)

        return TableRow(new_row, insert_after_index + 1)

    def _get_table(self, table_index: int) -> Table:
        """Get table by index with validation.

        Args:
            table_index: Index of the table

        Returns:
            Table at the specified index

        Raises:
            IndexError: If table_index is out of range
        """
        tables = self.all
        if table_index < 0 or table_index >= len(tables):
            raise IndexError(f"Table index {table_index} out of range (0-{len(tables) - 1})")
        return tables[table_index]

    def _resolve_row_index(self, table: Table, row_ref: int | str) -> int:
        """Resolve a row reference to an index.

        Args:
            table: The table to search in
            row_ref: Row index (int) or text to find (str)

        Returns:
            Index of the row

        Raises:
            IndexError: If row index is out of range
            ValueError: If text is not found or is ambiguous
        """
        if isinstance(row_ref, int):
            if row_ref < 0 or row_ref >= table.row_count:
                raise IndexError(f"Row index {row_ref} out of range (0-{table.row_count - 1})")
            return row_ref

        # Find row containing text
        matching_rows = [(i, row) for i, row in enumerate(table.rows) if row.contains(row_ref)]

        if not matching_rows:
            raise ValueError(f"No row found containing text: {row_ref}")
        if len(matching_rows) > 1:
            raise ValueError(
                f"Text '{row_ref}' found in {len(matching_rows)} rows - "
                "please use a more specific search or row index"
            )

        return matching_rows[0][0]

    def _create_row_element(self, cells: list[str]) -> etree._Element:
        """Create a new table row element with cells.

        Args:
            cells: List of text content for each cell

        Returns:
            The new row element
        """
        new_row = etree.Element(f"{{{WORD_NAMESPACE}}}tr")

        for cell_text in cells:
            tc = etree.SubElement(new_row, f"{{{WORD_NAMESPACE}}}tc")
            tc_pr = etree.SubElement(tc, f"{{{WORD_NAMESPACE}}}tcPr")
            tc_w = etree.SubElement(tc_pr, f"{{{WORD_NAMESPACE}}}tcW")
            tc_w.set(f"{{{WORD_NAMESPACE}}}w", "2880")
            tc_w.set(f"{{{WORD_NAMESPACE}}}type", "dxa")

            para = etree.SubElement(tc, f"{{{WORD_NAMESPACE}}}p")

            if cell_text:
                run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = cell_text

        return new_row

    def _add_row_insertion_tracking(
        self,
        row_element: etree._Element,
        author: str | AuthorIdentity | None,
    ) -> None:
        """Add insertion tracking properties to a row element.

        Args:
            row_element: The row element to mark as inserted
            author: Author for the tracked change
        """
        from datetime import datetime, timezone

        author_name = author if author is not None else self._document.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self._document._xml_generator.next_change_id
        self._document._xml_generator.next_change_id += 1

        # Add w:trPr with w:ins child to mark row as inserted
        tr_pr = etree.Element(f"{{{WORD_NAMESPACE}}}trPr")
        ins_elem = etree.SubElement(tr_pr, f"{{{WORD_NAMESPACE}}}ins")
        ins_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
        ins_elem.set(f"{{{WORD_NAMESPACE}}}author", str(author_name))
        ins_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

        # Insert trPr as first child of row
        row_element.insert(0, tr_pr)

    def _insert_row_in_table(
        self,
        table: Table,
        row_element: etree._Element,
        after_index: int,
    ) -> None:
        """Insert a row element into a table after a specified index.

        Args:
            table: The table to insert into
            row_element: The row element to insert
            after_index: Index to insert after
        """
        table_elem = table.element
        rows = table_elem.findall(f"{{{WORD_NAMESPACE}}}tr")
        target_row = rows[after_index]
        row_index = list(table_elem).index(target_row)
        table_elem.insert(row_index + 1, row_element)

    def delete_row(
        self,
        row: int | str,
        *,
        table_index: int = 0,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
    ) -> TableRow:
        """Delete a table row with optional tracked changes.

        Args:
            row: Row index (int) or text to find in a row (str)
            table_index: Which table to modify (default: 0 = first table)
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)

        Returns:
            The deleted TableRow object

        Raises:
            IndexError: If table_index or row index is out of range
            ValueError: If row text is not found or is ambiguous

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.delete_table_row(row=5, track=True)
        """

        table = self._get_table(table_index)
        delete_index = self._resolve_row_index(table, row)
        row_to_delete = table.rows[delete_index]

        if track:
            self._mark_row_as_deleted(row_to_delete.element, author)
        else:
            table.element.remove(row_to_delete.element)

        return row_to_delete

    def _mark_row_as_deleted(
        self,
        row_element: etree._Element,
        author: str | AuthorIdentity | None,
    ) -> None:
        """Mark a row element as deleted with tracking.

        Args:
            row_element: The row element to mark as deleted
            author: Author for the tracked change
        """
        from datetime import datetime, timezone

        author_name = author if author is not None else self._document.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self._document._xml_generator.next_change_id
        self._document._xml_generator.next_change_id += 1

        # Convert all w:t to w:delText within the row
        for t_elem in row_element.findall(f".//{{{WORD_NAMESPACE}}}t"):
            t_elem.tag = f"{{{WORD_NAMESPACE}}}delText"

        # Add or update w:trPr with w:del child to mark row as deleted
        tr_pr = row_element.find(f"{{{WORD_NAMESPACE}}}trPr")
        if tr_pr is None:
            tr_pr = etree.Element(f"{{{WORD_NAMESPACE}}}trPr")
            row_element.insert(0, tr_pr)

        del_elem = etree.SubElement(tr_pr, f"{{{WORD_NAMESPACE}}}del")
        del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
        del_elem.set(f"{{{WORD_NAMESPACE}}}author", str(author_name))
        del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

    def insert_column(
        self,
        after_column: int | str,
        cells: list[str],
        *,
        table_index: int = 0,
        header: str | None = None,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
    ) -> None:
        """Insert a new table column with optional tracked changes.

        Columns in OOXML are implicit - they are derived from cells in rows.
        This method inserts a new cell into each row at the specified position.

        Args:
            after_column: Column index (int) or text to find in a column (str).
                          Use -1 to insert before the first column.
            cells: List of text content for each cell in the new column.
                   Length must match the number of rows (excluding header if provided).
            table_index: Which table to modify (default: 0 = first table)
            header: Optional header text for the first row. If provided, cells list
                    should have one fewer element (for data rows only).
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)

        Raises:
            IndexError: If table_index is out of range
            ValueError: If after_column text is not found or is ambiguous
            ValueError: If number of cells doesn't match expected row count

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.insert_table_column(
            ...     after_column=1,
            ...     cells=["A", "B", "C"],
            ...     header="New Column",
            ...     track=True
            ... )
        """
        table = self._get_table(table_index)
        insert_after_index = self._resolve_column_index(table, after_column)

        # Validate cell count
        expected_cells = table.row_count
        if header is not None:
            expected_cells -= 1  # Header row handled separately

        if len(cells) != expected_cells:
            raise ValueError(
                f"Expected {expected_cells} cells (got {len(cells)}). "
                f"Table has {table.row_count} rows"
                + (", header is provided separately" if header else "")
            )

        # Insert a new gridCol in tblGrid
        self._insert_grid_column(table, insert_after_index)

        # Insert cells into each row
        cell_index = 0
        for row_idx, row in enumerate(table.rows):
            # Determine cell content
            if header is not None and row_idx == 0:
                cell_text = header
            else:
                cell_text = cells[cell_index]
                cell_index += 1

            # Create and insert new cell
            new_cell = self._create_cell_element(cell_text, track, author)
            self._insert_cell_in_row(row.element, new_cell, insert_after_index)

    def _resolve_column_index(self, table: Table, column_ref: int | str) -> int:
        """Resolve a column reference to an index.

        Args:
            table: The table to search in
            column_ref: Column index (int) or text to find (str)

        Returns:
            Index of the column (-1 for before first column)

        Raises:
            IndexError: If column index is out of range
            ValueError: If text is not found or is ambiguous
        """
        if isinstance(column_ref, int):
            if column_ref < -1 or column_ref >= table.col_count:
                raise IndexError(
                    f"Column index {column_ref} out of range (-1 to {table.col_count - 1})"
                )
            return column_ref

        # Find column containing text (check all rows)
        matching_cols: list[int] = []
        for row in table.rows:
            for cell in row.cells:
                if cell.contains(column_ref):
                    if cell.col_index not in matching_cols:
                        matching_cols.append(cell.col_index)

        if not matching_cols:
            raise ValueError(f"No column found containing text: {column_ref}")
        if len(matching_cols) > 1:
            raise ValueError(
                f"Text '{column_ref}' found in {len(matching_cols)} columns - "
                "please use a more specific search or column index"
            )

        return matching_cols[0]

    def _insert_grid_column(self, table: Table, after_index: int) -> None:
        """Insert a new gridCol element in the table grid.

        Args:
            table: The table to modify
            after_index: Index to insert after (-1 for beginning)
        """
        tbl_grid = table.element.find(f"{{{WORD_NAMESPACE}}}tblGrid")
        if tbl_grid is not None:
            grid_cols = tbl_grid.findall(f"{{{WORD_NAMESPACE}}}gridCol")
            # Default width for new column
            new_grid_col = etree.Element(f"{{{WORD_NAMESPACE}}}gridCol")
            new_grid_col.set(f"{{{WORD_NAMESPACE}}}w", "2880")

            if after_index == -1:
                # Insert at beginning
                tbl_grid.insert(0, new_grid_col)
            elif after_index < len(grid_cols):
                # Insert after specified column
                tbl_grid.insert(after_index + 1, new_grid_col)
            else:
                # Append at end
                tbl_grid.append(new_grid_col)

    def _create_cell_element(
        self,
        cell_text: str,
        track: bool,
        author: str | AuthorIdentity | None,
    ) -> etree._Element:
        """Create a new table cell element.

        Args:
            cell_text: Text content for the cell
            track: Whether to track the insertion
            author: Author for tracked changes

        Returns:
            The new cell element
        """
        new_tc = etree.Element(f"{{{WORD_NAMESPACE}}}tc")

        # Add cell properties
        tc_pr = etree.SubElement(new_tc, f"{{{WORD_NAMESPACE}}}tcPr")
        tc_w = etree.SubElement(tc_pr, f"{{{WORD_NAMESPACE}}}tcW")
        tc_w.set(f"{{{WORD_NAMESPACE}}}w", "2880")
        tc_w.set(f"{{{WORD_NAMESPACE}}}type", "dxa")

        # Add paragraph with content
        para = etree.SubElement(new_tc, f"{{{WORD_NAMESPACE}}}p")

        if cell_text:
            if track:
                self._add_tracked_cell_content(para, cell_text, author)
            else:
                run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
                t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
                t.text = cell_text

        return new_tc

    def _add_tracked_cell_content(
        self,
        para: etree._Element,
        cell_text: str,
        author: str | AuthorIdentity | None,
    ) -> None:
        """Add content to a cell paragraph with insertion tracking.

        Args:
            para: The paragraph element to add content to
            cell_text: Text content for the cell
            author: Author for tracked changes
        """
        from datetime import datetime, timezone

        author_name = author if author is not None else self._document.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self._document._xml_generator.next_change_id
        self._document._xml_generator.next_change_id += 1

        ins_elem = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}ins")
        ins_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
        ins_elem.set(f"{{{WORD_NAMESPACE}}}author", str(author_name))
        ins_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

        run = etree.SubElement(ins_elem, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = cell_text

    def _insert_cell_in_row(
        self,
        row_element: etree._Element,
        cell_element: etree._Element,
        after_index: int,
    ) -> None:
        """Insert a cell element into a row at the specified position.

        Args:
            row_element: The row element to insert into
            cell_element: The cell element to insert
            after_index: Index to insert after (-1 for beginning)
        """
        tc_elements = row_element.findall(f"{{{WORD_NAMESPACE}}}tc")

        if after_index == -1:
            # Insert before first cell
            if tc_elements:
                row_element.insert(list(row_element).index(tc_elements[0]), cell_element)
            else:
                row_element.append(cell_element)
        elif after_index < len(tc_elements):
            # Insert after specified cell
            target_tc = tc_elements[after_index]
            tc_position = list(row_element).index(target_tc)
            row_element.insert(tc_position + 1, cell_element)
        else:
            # Append at end of row
            row_element.append(cell_element)

    def delete_column(
        self,
        column: int | str,
        *,
        table_index: int = 0,
        track: bool = True,
        author: str | AuthorIdentity | None = None,
    ) -> None:
        """Delete a table column with optional tracked changes.

        Columns in OOXML are implicit - they are derived from cells in rows.
        This method removes or marks cells at the specified column position in each row.

        Args:
            column: Column index (int) or text to find in a column (str)
            table_index: Which table to modify (default: 0 = first table)
            track: Whether to track changes (default: True)
            author: Author for tracked changes (uses document author if None)

        Raises:
            IndexError: If table_index or column index is out of range
            ValueError: If column text is not found or is ambiguous

        Example:
            >>> doc = Document("contract.docx")
            >>> doc.delete_table_column(column=2, track=True)
        """
        table = self._get_table(table_index)
        delete_index = self._resolve_column_index_for_delete(table, column)

        # Remove gridCol from tblGrid (if not tracking)
        if not track:
            self._remove_grid_column(table, delete_index)

        # Process each row
        for row in table.rows:
            self._delete_cell_in_row(row.element, delete_index, track, author)

    def _resolve_column_index_for_delete(self, table: Table, column_ref: int | str) -> int:
        """Resolve a column reference to an index for deletion.

        Args:
            table: The table to search in
            column_ref: Column index (int) or text to find (str)

        Returns:
            Index of the column

        Raises:
            IndexError: If column index is out of range
            ValueError: If text is not found or is ambiguous
        """
        if isinstance(column_ref, int):
            if column_ref < 0 or column_ref >= table.col_count:
                raise IndexError(
                    f"Column index {column_ref} out of range (0-{table.col_count - 1})"
                )
            return column_ref

        # Find column containing text
        matching_cols: list[int] = []
        for row in table.rows:
            for cell in row.cells:
                if cell.contains(column_ref):
                    if cell.col_index not in matching_cols:
                        matching_cols.append(cell.col_index)

        if not matching_cols:
            raise ValueError(f"No column found containing text: {column_ref}")
        if len(matching_cols) > 1:
            raise ValueError(
                f"Text '{column_ref}' found in {len(matching_cols)} columns - "
                "please use a more specific search or column index"
            )

        return matching_cols[0]

    def _remove_grid_column(self, table: Table, column_index: int) -> None:
        """Remove a gridCol element from the table grid.

        Args:
            table: The table to modify
            column_index: Index of the column to remove
        """
        tbl_grid = table.element.find(f"{{{WORD_NAMESPACE}}}tblGrid")
        if tbl_grid is not None:
            grid_cols = tbl_grid.findall(f"{{{WORD_NAMESPACE}}}gridCol")
            if column_index < len(grid_cols):
                tbl_grid.remove(grid_cols[column_index])

    def _delete_cell_in_row(
        self,
        row_element: etree._Element,
        column_index: int,
        track: bool,
        author: str | AuthorIdentity | None,
    ) -> None:
        """Delete or mark a cell at the specified column in a row.

        Args:
            row_element: The row element to modify
            column_index: Index of the column to delete
            track: Whether to track the deletion
            author: Author for tracked changes
        """
        tc_elements = row_element.findall(f"{{{WORD_NAMESPACE}}}tc")

        if column_index >= len(tc_elements):
            # Row doesn't have this column (varying row lengths)
            return

        cell_to_delete = tc_elements[column_index]

        if track:
            self._mark_cell_as_deleted(cell_to_delete, author)
        else:
            row_element.remove(cell_to_delete)

    def _mark_cell_as_deleted(
        self,
        cell_element: etree._Element,
        author: str | AuthorIdentity | None,
    ) -> None:
        """Mark a cell's content as deleted with tracking.

        Args:
            cell_element: The cell element to mark as deleted
            author: Author for tracked changes
        """
        from datetime import datetime, timezone

        author_name = author if author is not None else self._document.author
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        change_id = self._document._xml_generator.next_change_id
        self._document._xml_generator.next_change_id += 1

        # Convert all w:t to w:delText within the cell
        for t_elem in cell_element.findall(f".//{{{WORD_NAMESPACE}}}t"):
            t_elem.tag = f"{{{WORD_NAMESPACE}}}delText"

        # Wrap all runs in deletion markers
        for para in cell_element.findall(f"{{{WORD_NAMESPACE}}}p"):
            for run in list(para.findall(f"{{{WORD_NAMESPACE}}}r")):
                # Create deletion wrapper
                del_elem = etree.Element(f"{{{WORD_NAMESPACE}}}del")
                del_elem.set(f"{{{WORD_NAMESPACE}}}id", str(change_id))
                del_elem.set(f"{{{WORD_NAMESPACE}}}author", str(author_name))
                del_elem.set(f"{{{WORD_NAMESPACE}}}date", timestamp)

                # Move run into deletion
                run_index = list(para).index(run)
                para.remove(run)
                del_elem.append(run)
                para.insert(run_index, del_elem)

                # Increment change ID for next run
                change_id = self._document._xml_generator.next_change_id
                self._document._xml_generator.next_change_id += 1
