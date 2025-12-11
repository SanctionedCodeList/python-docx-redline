"""
Document model classes for python_docx_redline.

These classes provide convenient wrappers around OOXML elements.
"""

from python_docx_redline.models.footnote import Endnote, Footnote
from python_docx_redline.models.paragraph import Paragraph
from python_docx_redline.models.section import Section
from python_docx_redline.models.table import Table, TableCell, TableRow

__all__ = [
    "Paragraph",
    "Section",
    "Table",
    "TableRow",
    "TableCell",
    "Footnote",
    "Endnote",
]
