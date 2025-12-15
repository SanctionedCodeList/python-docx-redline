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
    "compare_documents",
    "OOXMLPackage",
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
    "ComparisonStats",
    "Paragraph",
    "Section",
    "Comment",
    "CommentRange",
    "Footnote",
    "Endnote",
    "Header",
    "Footer",
    "HeaderFooterType",
    "TrackedChange",
    "ChangeType",
    # Export functionality
    "ChangeContext",
    "ExportedChange",
    "ChangeReport",
    "export_changes_json",
    "export_changes_markdown",
    "export_changes_html",
    "generate_change_report",
]

# Import author identity
from .author import AuthorIdentity

# Import compatibility helpers (python-docx integration)
from .compat import from_python_docx, to_python_docx

# Import document class and standalone functions
from .document import Document, compare_documents
from .errors import (
    AmbiguousTextError,
    ContinuityWarning,
    DocxRedlineError,
    TextNotFoundError,
    ValidationError,
)

# Import export functionality
from .export import (
    ChangeContext,
    ChangeReport,
    ExportedChange,
    export_changes_html,
    export_changes_json,
    export_changes_markdown,
    generate_change_report,
)

# Import model classes
from .models.comment import Comment, CommentRange
from .models.footnote import Endnote, Footnote
from .models.header_footer import Footer, Header, HeaderFooterType
from .models.paragraph import Paragraph
from .models.section import Section
from .models.tracked_change import ChangeType, TrackedChange

# Import package class
from .package import OOXMLPackage

# Import result types
from .results import AcceptResult, ComparisonStats, EditResult, FormatResult, RejectResult

# Import scope evaluation
from .scope import ScopeEvaluator

# Import suggestion generator
from .suggestions import SuggestionGenerator

# Import text search
from .text_search import TextSearch, TextSpan

# Import XML generation
from .tracked_xml import TrackedXMLGenerator
