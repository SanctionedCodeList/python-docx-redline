"""
python_docx_redline - A high-level Python API for editing Word documents with tracked changes.

This package provides a simple interface for making surgical edits to Word documents
without needing to write raw OOXML XML. It handles tracked changes, text search across
fragmented runs, and provides helpful error messages.

Example:
    >>> from python_docx_redline import Document
    >>> doc = Document("contract.docx")
    >>> doc.insert_tracked("new clause text", after="Section 2.1")
    >>> doc.save("contract_edited.docx")
"""

__version__ = "0.1.0"
__author__ = "Parker Hancock"
__all__ = [
    "Document",
    "AuthorIdentity",
    "from_python_docx",
    "to_python_docx",
    "DocxRedlineError",
    "TextNotFoundError",
    "AmbiguousTextError",
    "ValidationError",
    "ContinuityWarning",
    "TextSearch",
    "TextSpan",
    "TrackedXMLGenerator",
    "SuggestionGenerator",
    "ScopeEvaluator",
    "EditResult",
    "AcceptResult",
    "RejectResult",
    "FormatResult",
    "Paragraph",
    "Section",
    "Comment",
    "CommentRange",
    "Footnote",
    "Endnote",
    "TrackedChange",
    "ChangeType",
]

# Import author identity
from .author import AuthorIdentity

# Import compatibility helpers (python-docx integration)
from .compat import from_python_docx, to_python_docx

# Import error classes
# Import document class
from .document import Document
from .errors import (
    AmbiguousTextError,
    ContinuityWarning,
    DocxRedlineError,
    TextNotFoundError,
    ValidationError,
)

# Import model classes
from .models.comment import Comment, CommentRange
from .models.footnote import Endnote, Footnote
from .models.paragraph import Paragraph
from .models.section import Section
from .models.tracked_change import ChangeType, TrackedChange

# Import result types
from .results import AcceptResult, EditResult, FormatResult, RejectResult

# Import scope evaluation
from .scope import ScopeEvaluator

# Import suggestion generator
from .suggestions import SuggestionGenerator

# Import text search
from .text_search import TextSearch, TextSpan

# Import XML generation
from .tracked_xml import TrackedXMLGenerator
