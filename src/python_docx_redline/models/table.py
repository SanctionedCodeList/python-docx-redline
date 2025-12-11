"""
Table wrapper classes for convenient access to table elements.
"""

from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from python_docx_redline.models.paragraph import Paragraph

# Word namespace
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class TableCell:
    """Wrapper around a w:tc (table cell) element.

    Provides convenient Python API for working with table cells.
    """

    def __init__(self, element: etree._Element, row_index: int, col_index: int):
        """Initialize TableCell wrapper.

        Args:
            element: The w:tc XML element to wrap
            row_index: 0-based row index in table
            col_index: 0-based column index in row
        """
        if element.tag != f"{{{WORD_NAMESPACE}}}tc":
            raise ValueError(f"Expected w:tc element, got {element.tag}")
        self._element = element
        self._row_index = row_index
        self._col_index = col_index

    @property
    def element(self) -> etree._Element:
        """Get the underlying XML element."""
        return self._element

    @property
    def row_index(self) -> int:
        """Get the row index (0-based)."""
        return self._row_index

    @property
    def col_index(self) -> int:
        """Get the column index (0-based)."""
        return self._col_index

    @property
    def text(self) -> str:
        """Get all text content from the cell.

        Extracts text from both w:t and w:delText elements in all paragraphs.

        Returns:
            Combined text from all paragraphs in the cell
        """
        text_elements = self._element.findall(f".//{{{WORD_NAMESPACE}}}t")
        deltext_elements = self._element.findall(f".//{{{WORD_NAMESPACE}}}delText")
        return "".join(elem.text or "" for elem in text_elements + deltext_elements)

    @text.setter
    def text(self, value: str) -> None:
        """Set the text content of the cell.

        This replaces all paragraphs with a single paragraph containing the new text.

        Args:
            value: New text content
        """
        # Remove all existing paragraphs
        for para in self._element.findall(f"{{{WORD_NAMESPACE}}}p"):
            self._element.remove(para)

        # Create new paragraph with text
        para = etree.SubElement(self._element, f"{{{WORD_NAMESPACE}}}p")
        run = etree.SubElement(para, f"{{{WORD_NAMESPACE}}}r")
        t = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
        t.text = value

    @property
    def paragraphs(self) -> list["Paragraph"]:
        """Get all paragraphs in this cell.

        Returns:
            List of Paragraph objects
        """
        from python_docx_redline.models.paragraph import Paragraph

        para_elements = self._element.findall(f"{{{WORD_NAMESPACE}}}p")
        return [Paragraph(elem) for elem in para_elements]

    def contains(self, text: str, case_sensitive: bool = True) -> bool:
        """Check if cell contains specific text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case sensitive

        Returns:
            True if text is found in cell
        """
        cell_text = self.text
        if not case_sensitive:
            cell_text = cell_text.lower()
            text = text.lower()
        return text in cell_text

    def __repr__(self) -> str:
        """String representation of the cell."""
        text_preview = self.text[:30] + "..." if len(self.text) > 30 else self.text
        return f"<TableCell[{self._row_index},{self._col_index}]: {text_preview!r}>"


class TableRow:
    """Wrapper around a w:tr (table row) element.

    Provides convenient Python API for working with table rows.
    """

    def __init__(self, element: etree._Element, row_index: int):
        """Initialize TableRow wrapper.

        Args:
            element: The w:tr XML element to wrap
            row_index: 0-based row index in table
        """
        if element.tag != f"{{{WORD_NAMESPACE}}}tr":
            raise ValueError(f"Expected w:tr element, got {element.tag}")
        self._element = element
        self._row_index = row_index

    @property
    def element(self) -> etree._Element:
        """Get the underlying XML element."""
        return self._element

    @property
    def index(self) -> int:
        """Get the row index (0-based)."""
        return self._row_index

    @property
    def cells(self) -> list[TableCell]:
        """Get all cells in this row.

        Returns:
            List of TableCell objects
        """
        cell_elements = self._element.findall(f"{{{WORD_NAMESPACE}}}tc")
        return [
            TableCell(elem, self._row_index, col_idx) for col_idx, elem in enumerate(cell_elements)
        ]

    def contains(self, text: str, case_sensitive: bool = True) -> bool:
        """Check if any cell in row contains specific text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case sensitive

        Returns:
            True if text is found in any cell
        """
        return any(cell.contains(text, case_sensitive) for cell in self.cells)

    def __repr__(self) -> str:
        """String representation of the row."""
        cell_count = len(self.cells)
        return f"<TableRow[{self._row_index}]: {cell_count} cells>"


class Table:
    """Wrapper around a w:tbl (table) element.

    Provides convenient Python API for working with tables.
    """

    def __init__(self, element: etree._Element):
        """Initialize Table wrapper.

        Args:
            element: The w:tbl XML element to wrap
        """
        if element.tag != f"{{{WORD_NAMESPACE}}}tbl":
            raise ValueError(f"Expected w:tbl element, got {element.tag}")
        self._element = element

    @property
    def element(self) -> etree._Element:
        """Get the underlying XML element."""
        return self._element

    @property
    def rows(self) -> list[TableRow]:
        """Get all rows in the table.

        Returns:
            List of TableRow objects
        """
        row_elements = self._element.findall(f"{{{WORD_NAMESPACE}}}tr")
        return [TableRow(elem, row_idx) for row_idx, elem in enumerate(row_elements)]

    @property
    def row_count(self) -> int:
        """Get the number of rows in table.

        Returns:
            Number of rows
        """
        return len(self.rows)

    @property
    def col_count(self) -> int:
        """Get the number of columns in table.

        Returns number of cells in the first row, or 0 if no rows.

        Returns:
            Number of columns
        """
        rows = self.rows
        if not rows:
            return 0
        return len(rows[0].cells)

    def get_cell(self, row: int, col: int) -> TableCell:
        """Get cell at specified position.

        Args:
            row: 0-based row index
            col: 0-based column index

        Returns:
            TableCell at the specified position

        Raises:
            IndexError: If row or column index is out of range
        """
        rows = self.rows
        if row < 0 or row >= len(rows):
            raise IndexError(f"Row index {row} out of range (0-{len(rows)-1})")

        cells = rows[row].cells
        if col < 0 or col >= len(cells):
            raise IndexError(f"Column index {col} out of range (0-{len(cells)-1})")

        return cells[col]

    def find_cell(self, text: str, case_sensitive: bool = True) -> TableCell | None:
        """Find first cell containing text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case sensitive

        Returns:
            First TableCell containing the text, or None if not found
        """
        for row in self.rows:
            for cell in row.cells:
                if cell.contains(text, case_sensitive):
                    return cell
        return None

    def contains(self, text: str, case_sensitive: bool = True) -> bool:
        """Check if table contains specific text.

        Args:
            text: Text to search for
            case_sensitive: Whether search should be case sensitive

        Returns:
            True if text is found in any cell
        """
        return self.find_cell(text, case_sensitive) is not None

    def __repr__(self) -> str:
        """String representation of the table."""
        return f"<Table: {self.row_count} rows × {self.col_count} cols>"
