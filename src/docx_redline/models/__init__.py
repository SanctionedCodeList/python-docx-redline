"""
Document model classes for docx_redline.

These classes provide convenient wrappers around OOXML elements.
"""

from docx_redline.models.paragraph import Paragraph
from docx_redline.models.section import Section

__all__ = ["Paragraph", "Section"]
