"""
Tests for table wrapper classes and table operations.
"""

import tempfile
import zipfile
from pathlib import Path

import pytest

from python_docx_redline import Document
from python_docx_redline.models.table import Table, TableCell, TableRow

# Minimal document XML with a simple 2x2 table
DOCUMENT_WITH_TABLE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Text before table</w:t>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="2880"/>
        <w:gridCol w:w="2880"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Row 1, Col 1</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Row 1, Col 2</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Row 2, Col 1</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Row 2, Col 2</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:t>Text after table</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


def create_test_docx(content: str = DOCUMENT_WITH_TABLE) -> Path:
    """Create a test .docx file with proper OOXML structure.

    Args:
        content: The document.xml content

    Returns:
        Path to the created .docx file
    """
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    # Proper Content_Types.xml
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    # Proper relationships file
    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    # Create the .docx ZIP file
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", content)

    return docx_path


def test_document_tables_property():
    """Test that Document.tables returns list of Table objects."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        tables = doc.tables

        assert isinstance(tables, list)
        assert len(tables) == 1
        assert isinstance(tables[0], Table)
    finally:
        docx_path.unlink()


def test_table_row_count():
    """Test Table.row_count property."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        assert table.row_count == 2
    finally:
        docx_path.unlink()


def test_table_col_count():
    """Test Table.col_count property."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        assert table.col_count == 2
    finally:
        docx_path.unlink()


def test_table_rows_property():
    """Test Table.rows returns list of TableRow objects."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        rows = table.rows

        assert isinstance(rows, list)
        assert len(rows) == 2
        assert all(isinstance(row, TableRow) for row in rows)
        assert rows[0].index == 0
        assert rows[1].index == 1
    finally:
        docx_path.unlink()


def test_table_get_cell():
    """Test Table.get_cell method."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        cell = table.get_cell(0, 0)
        assert isinstance(cell, TableCell)
        assert cell.row_index == 0
        assert cell.col_index == 0
        assert cell.text == "Row 1, Col 1"

        cell = table.get_cell(1, 1)
        assert cell.row_index == 1
        assert cell.col_index == 1
        assert cell.text == "Row 2, Col 2"
    finally:
        docx_path.unlink()


def test_table_get_cell_out_of_range():
    """Test Table.get_cell raises IndexError for invalid indices."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        with pytest.raises(IndexError):
            table.get_cell(10, 0)

        with pytest.raises(IndexError):
            table.get_cell(0, 10)
    finally:
        docx_path.unlink()


def test_table_row_cells():
    """Test TableRow.cells property."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        row = table.rows[0]

        cells = row.cells
        assert isinstance(cells, list)
        assert len(cells) == 2
        assert all(isinstance(cell, TableCell) for cell in cells)
        assert cells[0].text == "Row 1, Col 1"
        assert cells[1].text == "Row 1, Col 2"
    finally:
        docx_path.unlink()


def test_table_cell_text_getter():
    """Test TableCell.text getter."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        cell = table.get_cell(0, 0)

        assert cell.text == "Row 1, Col 1"
    finally:
        docx_path.unlink()


def test_table_cell_text_setter():
    """Test TableCell.text setter."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        cell = table.get_cell(0, 0)

        # Update cell text
        cell.text = "Updated Cell"
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table2 = doc2.tables[0]
        cell2 = table2.get_cell(0, 0)

        assert cell2.text == "Updated Cell"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_table_cell_paragraphs():
    """Test TableCell.paragraphs property."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        cell = table.get_cell(0, 0)

        paragraphs = cell.paragraphs
        assert isinstance(paragraphs, list)
        assert len(paragraphs) == 1
        assert paragraphs[0].text == "Row 1, Col 1"
    finally:
        docx_path.unlink()


def test_table_cell_contains():
    """Test TableCell.contains method."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        cell = table.get_cell(0, 0)

        assert cell.contains("Row 1")
        assert cell.contains("Col 1")
        assert not cell.contains("Row 2")

        # Case-insensitive
        assert cell.contains("row 1", case_sensitive=False)
    finally:
        docx_path.unlink()


def test_table_row_contains():
    """Test TableRow.contains method."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        row = table.rows[0]

        assert row.contains("Row 1")
        assert row.contains("Col 1")
        assert row.contains("Col 2")
        assert not row.contains("Row 2")
    finally:
        docx_path.unlink()


def test_table_contains():
    """Test Table.contains method."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        assert table.contains("Row 1, Col 1")
        assert table.contains("Row 2, Col 2")
        assert not table.contains("Nonexistent text")
    finally:
        docx_path.unlink()


def test_table_find_cell():
    """Test Table.find_cell method."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        cell = table.find_cell("Row 2, Col 1")
        assert cell is not None
        assert cell.row_index == 1
        assert cell.col_index == 0

        cell_none = table.find_cell("Nonexistent")
        assert cell_none is None
    finally:
        docx_path.unlink()


def test_document_find_table():
    """Test Document.find_table method."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        table = doc.find_table("Row 1, Col 1")
        assert table is not None
        assert table.row_count == 2

        table_none = doc.find_table("Nonexistent table text")
        assert table_none is None
    finally:
        docx_path.unlink()


def test_table_repr():
    """Test Table.__repr__ method."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        repr_str = repr(table)
        assert "Table" in repr_str
        assert "2 rows" in repr_str
        assert "2 cols" in repr_str
    finally:
        docx_path.unlink()


def test_table_row_repr():
    """Test TableRow.__repr__ method."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        row = table.rows[0]

        repr_str = repr(row)
        assert "TableRow" in repr_str
        assert "[0]" in repr_str
        assert "2 cells" in repr_str
    finally:
        docx_path.unlink()


def test_table_cell_repr():
    """Test TableCell.__repr__ method."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        cell = table.get_cell(0, 0)

        repr_str = repr(cell)
        assert "TableCell" in repr_str
        assert "[0,0]" in repr_str
        assert "Row 1, Col 1" in repr_str
    finally:
        docx_path.unlink()


def test_update_cell_tracked():
    """Test Document.update_cell with tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Update cell with tracking
        doc.update_cell(0, 0, "Updated Cell", track=True)
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]
        cell = table.get_cell(0, 0)

        # Should have tracked insertion
        assert cell.contains("Updated Cell")

        # Check for tracked changes in XML
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        assert len(insertions) > 0, "Should have tracked insertions"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_update_cell_untracked():
    """Test Document.update_cell without tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Update cell without tracking
        doc.update_cell(1, 1, "New Value", track=False)
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]
        cell = table.get_cell(1, 1)

        assert cell.text == "New Value"

        # Check for no tracked changes
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions) == 0, "Should have no tracked insertions"
        assert len(deletions) == 0, "Should have no tracked deletions"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_update_cell_invalid_indices():
    """Test Document.update_cell with invalid indices."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        with pytest.raises(IndexError):
            doc.update_cell(10, 0, "Should fail")

        with pytest.raises(IndexError):
            doc.update_cell(0, 10, "Should fail")
    finally:
        docx_path.unlink()


def test_replace_in_table_tracked():
    """Test Document.replace_in_table with tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Replace text in table with tracking
        count = doc.replace_in_table("Row 1", "Updated Row", track=True)
        assert count == 2  # Should find "Row 1" in two cells

        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]

        assert table.contains("Updated Row")

        # Check for tracked changes
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions) > 0, "Should have tracked insertions"
        assert len(deletions) > 0, "Should have tracked deletions"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_replace_in_table_untracked():
    """Test Document.replace_in_table without tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Replace text without tracking
        count = doc.replace_in_table("Col 1", "Column A", track=False)
        assert count == 2

        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]

        assert table.contains("Column A")

        # Check for no tracked changes
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions) == 0, "Should have no tracked insertions"
        assert len(deletions) == 0, "Should have no tracked deletions"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_replace_in_table_specific_table():
    """Test Document.replace_in_table targeting a specific table."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        # Replace in specific table only
        count = doc.replace_in_table("Row 1", "Updated", table_index=0, track=False)
        assert count == 2
    finally:
        docx_path.unlink()


def test_replace_in_table_case_sensitivity():
    """Test Document.replace_in_table case sensitivity."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        # Case-sensitive (default) - shouldn't match
        count = doc.replace_in_table("row 1", "Updated", track=False)
        assert count == 0

        # Case-insensitive - should match
        count = doc.replace_in_table("row 1", "Updated", case_sensitive=False, track=False)
        assert count == 2
    finally:
        docx_path.unlink()


def test_insert_table_row_tracked():
    """Test Document.insert_table_row with tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Insert new row with tracking
        doc.insert_table_row(after_row=0, cells=["Row 1.5, Col 1", "Row 1.5, Col 2"], track=True)
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]

        # Should now have 3 rows
        assert table.row_count == 3
        assert table.contains("Row 1.5")

        # Check for tracked changes
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        assert len(insertions) > 0, "Should have tracked insertions"

        # Verify row order
        assert table.get_cell(0, 0).text == "Row 1, Col 1"
        assert table.get_cell(1, 0).contains("Row 1.5")
        assert table.get_cell(2, 0).text == "Row 2, Col 1"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_insert_table_row_untracked():
    """Test Document.insert_table_row without tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Insert row without tracking
        doc.insert_table_row(after_row=1, cells=["New Row, Col 1", "New Row, Col 2"], track=False)
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]

        assert table.row_count == 3
        assert table.contains("New Row")

        # Check for no tracked changes
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions) == 0, "Should have no tracked insertions"
        assert len(deletions) == 0, "Should have no tracked deletions"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_insert_table_row_by_text():
    """Test Document.insert_table_row using text to find row."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Insert after row containing specific text
        doc.insert_table_row(after_row="Row 1, Col 1", cells=["Inserted", "Row"], track=False)
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]

        assert table.row_count == 3
        assert table.get_cell(1, 0).text == "Inserted"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_insert_table_row_mismatched_columns():
    """Test Document.insert_table_row with wrong number of cells."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        # Should raise error for mismatched column count
        with pytest.raises(ValueError, match="Expected 2 cells, got 1"):
            doc.insert_table_row(after_row=0, cells=["Only one cell"], track=False)
    finally:
        docx_path.unlink()


def test_delete_table_row_tracked():
    """Test Document.delete_table_row with tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Delete row with tracking
        doc.delete_table_row(row=0, track=True)
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]

        # Should still show 2 rows (deletion is tracked)
        assert table.row_count == 2

        # Check for tracked deletions
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(deletions) > 0, "Should have tracked deletions"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_delete_table_row_untracked():
    """Test Document.delete_table_row without tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Delete row without tracking
        doc.delete_table_row(row=1, track=False)
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]

        # Should now have 1 row
        assert table.row_count == 1
        assert table.get_cell(0, 0).text == "Row 1, Col 1"

        # Check for no tracked changes
        insertions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins"
        )
        deletions = doc2.xml_root.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"
        )
        assert len(insertions) == 0, "Should have no tracked insertions"
        assert len(deletions) == 0, "Should have no tracked deletions"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_delete_table_row_by_text():
    """Test Document.delete_table_row using text to find row."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        # Delete row containing specific text
        doc.delete_table_row(row="Row 2, Col 1", track=False)
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table = doc2.tables[0]

        assert table.row_count == 1
        assert not table.contains("Row 2")
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_delete_table_row_invalid_index():
    """Test Document.delete_table_row with invalid index."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        with pytest.raises(IndexError):
            doc.delete_table_row(row=10, track=False)
    finally:
        docx_path.unlink()


def test_delete_table_row_text_not_found():
    """Test Document.delete_table_row with non-existent text."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        with pytest.raises(ValueError, match="No row found containing text"):
            doc.delete_table_row(row="Nonexistent text", track=False)
    finally:
        docx_path.unlink()


# Column operation tests


def test_insert_table_column_by_index():
    """Test Document.insert_table_column using column index."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        assert table.col_count == 2

        doc.insert_table_column(
            after_column=0,
            cells=["New Row 1", "New Row 2"],
            track=False,
        )
        doc.save(output_path)

        doc2 = Document(output_path)
        table2 = doc2.tables[0]

        assert table2.col_count == 3
        assert table2.get_cell(0, 1).text == "New Row 1"
        assert table2.get_cell(1, 1).text == "New Row 2"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_insert_table_column_at_beginning():
    """Test inserting column at the beginning using after_column=-1."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        doc.insert_table_column(
            after_column=-1,
            cells=["First Row 1", "First Row 2"],
            track=False,
        )
        doc.save(output_path)

        doc2 = Document(output_path)
        table = doc2.tables[0]

        assert table.col_count == 3
        assert table.get_cell(0, 0).text == "First Row 1"
        assert table.get_cell(1, 0).text == "First Row 2"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_insert_table_column_with_header():
    """Test inserting column with separate header text."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        doc.insert_table_column(
            after_column=0,
            cells=["Data Value"],
            header="New Header",
            track=False,
        )
        doc.save(output_path)

        doc2 = Document(output_path)
        table = doc2.tables[0]

        assert table.col_count == 3
        assert table.get_cell(0, 1).text == "New Header"
        assert table.get_cell(1, 1).text == "Data Value"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_insert_table_column_with_tracking():
    """Test inserting column with tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        doc.insert_table_column(
            after_column=0,
            cells=["Tracked R1", "Tracked R2"],
            track=True,
            author="Test Author",
        )
        doc.save(output_path)

        doc2 = Document(output_path)
        table = doc2.tables[0]

        assert table.col_count == 3

        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        for row in table.rows:
            cell = row.cells[1]
            ins_elements = cell.element.findall(f".//{{{word_ns}}}ins")
            assert len(ins_elements) > 0, "Expected insertion tracking in new cell"
            assert ins_elements[0].get(f"{{{word_ns}}}author") == "Test Author"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_insert_table_column_invalid_index():
    """Test insert_table_column with invalid column index."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        with pytest.raises(IndexError):
            doc.insert_table_column(after_column=10, cells=["A", "B"], track=False)
    finally:
        docx_path.unlink()


def test_insert_table_column_wrong_cell_count():
    """Test insert_table_column with wrong number of cells."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        with pytest.raises(ValueError, match="Expected 2 cells"):
            doc.insert_table_column(after_column=0, cells=["A", "B", "C"], track=False)
    finally:
        docx_path.unlink()


def test_delete_table_column_by_index():
    """Test Document.delete_table_column using column index."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        assert table.col_count == 2

        doc.delete_table_column(column=0, track=False)
        doc.save(output_path)

        doc2 = Document(output_path)
        table2 = doc2.tables[0]

        assert table2.col_count == 1
        assert table2.get_cell(0, 0).text == "Row 1, Col 2"
        assert table2.get_cell(1, 0).text == "Row 2, Col 2"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_delete_table_column_with_tracking():
    """Test deleting column with tracked changes."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)

        doc.delete_table_column(column=1, track=True, author="Test Author")
        doc.save(output_path)

        doc2 = Document(output_path)
        table = doc2.tables[0]

        assert table.col_count == 2

        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        for row in table.rows:
            cell = row.cells[1]
            del_elements = cell.element.findall(f".//{{{word_ns}}}del")
            assert len(del_elements) > 0, "Expected deletion tracking in cell"
            assert del_elements[0].get(f"{{{word_ns}}}author") == "Test Author"

            del_text_elements = cell.element.findall(f".//{{{word_ns}}}delText")
            assert len(del_text_elements) > 0, "Expected delText elements"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_delete_table_column_invalid_index():
    """Test delete_table_column with invalid column index."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        with pytest.raises(IndexError):
            doc.delete_table_column(column=10, track=False)
    finally:
        docx_path.unlink()


def test_insert_and_delete_column_roundtrip():
    """Test inserting and then deleting a column."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)
        table = doc.tables[0]

        initial_col_count = table.col_count

        doc.insert_table_column(after_column=0, cells=["Temp R1", "Temp R2"], track=False)

        table = doc.tables[0]
        assert table.col_count == initial_col_count + 1

        doc.delete_table_column(column=1, track=False)

        table = doc.tables[0]
        assert table.col_count == initial_col_count

        doc.save(output_path)

        doc2 = Document(output_path)
        assert doc2.tables[0].col_count == initial_col_count
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_column_operations_update_tbl_grid():
    """Test that column operations properly update w:tblGrid."""
    docx_path = create_test_docx()

    try:
        doc = Document(docx_path)

        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        table = doc.tables[0]

        tbl_grid = table.element.find(f"{{{word_ns}}}tblGrid")
        initial_grid_cols = len(tbl_grid.findall(f"{{{word_ns}}}gridCol"))
        assert initial_grid_cols == 2

        doc.insert_table_column(after_column=0, cells=["A", "B"], track=False)

        tbl_grid = doc.tables[0].element.find(f"{{{word_ns}}}tblGrid")
        assert len(tbl_grid.findall(f"{{{word_ns}}}gridCol")) == 3

        doc.delete_table_column(column=0, track=False)

        tbl_grid = doc.tables[0].element.find(f"{{{word_ns}}}tblGrid")
        assert len(tbl_grid.findall(f"{{{word_ns}}}gridCol")) == 2
    finally:
        docx_path.unlink(missing_ok=True)


# === Tests for Issue #6: TableCell.text setter preserving tcPr and pPr, markdown support ===


def test_table_cell_text_setter_preserves_tcpr():
    """Test that setting cell text preserves cell properties (w:tcPr)."""
    # Create a document with a table cell that has tcPr
    doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="5000" w:type="dxa"/>
            <w:shd w:fill="FFFF00" w:val="clear"/>
          </w:tcPr>
          <w:p><w:r><w:t>Original text</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"""

    docx_path = create_test_docx(doc_xml)
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        cell = table.get_cell(0, 0)

        # Set new text
        cell.text = "New text"
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table2 = doc2.tables[0]
        cell2 = table2.get_cell(0, 0)

        # Verify text was updated
        assert cell2.text == "New text"

        # Verify tcPr was preserved (check for shading element)
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        tc_pr = cell2.element.find(f"{{{word_ns}}}tcPr")
        assert tc_pr is not None

        shd = tc_pr.find(f"{{{word_ns}}}shd")
        assert shd is not None
        assert shd.get(f"{{{word_ns}}}fill") == "FFFF00"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_table_cell_text_setter_preserves_ppr():
    """Test that setting cell text preserves first paragraph properties."""
    # Create a document with a table cell that has paragraph properties
    doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:pPr>
              <w:jc w:val="center"/>
            </w:pPr>
            <w:r><w:t>Centered text</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"""

    docx_path = create_test_docx(doc_xml)
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        cell = table.get_cell(0, 0)

        # Set new text
        cell.text = "Still centered"
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table2 = doc2.tables[0]
        cell2 = table2.get_cell(0, 0)

        # Verify text was updated
        assert cell2.text == "Still centered"

        # Verify paragraph properties were preserved
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        para = cell2.element.find(f"{{{word_ns}}}p")
        assert para is not None

        ppr = para.find(f"{{{word_ns}}}pPr")
        assert ppr is not None

        jc = ppr.find(f"{{{word_ns}}}jc")
        assert jc is not None
        assert jc.get(f"{{{word_ns}}}val") == "center"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def test_table_cell_text_setter_markdown_support():
    """Test that setting cell text supports markdown formatting."""
    docx_path = create_test_docx()
    output_path = docx_path.parent / "output.docx"

    try:
        doc = Document(docx_path)
        table = doc.tables[0]
        cell = table.get_cell(0, 0)

        # Set text with markdown
        cell.text = "This is **bold** and *italic*"
        doc.save(output_path)

        # Reload and verify
        doc2 = Document(output_path)
        table2 = doc2.tables[0]
        cell2 = table2.get_cell(0, 0)

        # Text should read correctly (without markdown markers)
        assert cell2.text == "This is bold and italic"

        # Verify formatting in XML
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        runs = cell2.element.findall(f".//{{{word_ns}}}r")
        assert len(runs) >= 3  # Multiple runs for different formatting

        # Find bold run
        bold_found = False
        italic_found = False
        for run in runs:
            rpr = run.find(f"{{{word_ns}}}rPr")
            if rpr is not None:
                if rpr.find(f"{{{word_ns}}}b") is not None:
                    bold_found = True
                if rpr.find(f"{{{word_ns}}}i") is not None:
                    italic_found = True

        assert bold_found, "Expected bold formatting"
        assert italic_found, "Expected italic formatting"
    finally:
        docx_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


# Run tests with: pytest tests/test_table.py -v
