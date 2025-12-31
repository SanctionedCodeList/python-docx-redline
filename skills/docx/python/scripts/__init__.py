"""
Scripts for working with Word documents: comments, tracked changes, and editing.

These scripts provide low-level OOXML manipulation capabilities for advanced
scenarios not covered by python-docx-redline. For most use cases, prefer
using python-docx-redline's high-level API instead.
"""

from .document import Document, DocxXMLEditor
from .utilities import XMLEditor

__all__ = ["Document", "DocxXMLEditor", "XMLEditor"]
