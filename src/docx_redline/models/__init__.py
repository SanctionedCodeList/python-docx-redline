"""
Document model classes for docx_redline.

These classes provide convenient wrappers around OOXML elements.
"""

from docx_redline.models.footnote import Endnote, Footnote
from docx_redline.models.paragraph import Paragraph
from docx_redline.models.section import Section
from docx_redline.models.table import Table, TableCell, TableRow

__all__ = [
    "Paragraph",
    "Section",
    "Table",
    "TableRow",
    "TableCell",
    "Footnote",
    "Endnote",
]
