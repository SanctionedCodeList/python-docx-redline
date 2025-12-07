"""
docx_redline - A high-level Python API for editing Word documents with tracked changes.

This package provides a simple interface for making surgical edits to Word documents
without needing to write raw OOXML XML. It handles tracked changes, text search across
fragmented runs, and provides helpful error messages.

Example:
    >>> from docx_redline import Document
    >>> doc = Document("contract.docx")
    >>> doc.insert_tracked("new clause text", after="Section 2.1")
    >>> doc.save("contract_edited.docx")
"""

__version__ = "0.1.0"
__author__ = "Parker Hancock"
__all__ = [
    "Document",
    "DocxRedlineError",
    "TextNotFoundError",
    "AmbiguousTextError",
    "ValidationError",
    "TextSearch",
    "TextSpan",
    "TrackedXMLGenerator",
    "SuggestionGenerator",
]

# Import error classes
# Import document class
from .document import Document
from .errors import (
    AmbiguousTextError,
    DocxRedlineError,
    TextNotFoundError,
    ValidationError,
)

# Import suggestion generator
from .suggestions import SuggestionGenerator

# Import text search
from .text_search import TextSearch, TextSpan

# Import XML generation
from .tracked_xml import TrackedXMLGenerator
