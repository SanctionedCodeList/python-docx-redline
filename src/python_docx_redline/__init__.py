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

__version__ = "0.2.0"
__author__ = "Parker Hancock"
__all__ = [
    "Document",
    "compare_documents",
    "OOXMLPackage",
    "RelationshipManager",
    "RelationshipTypes",
    "ContentTypeManager",
    "ContentTypes",
    "AuthorIdentity",
    "from_python_docx",
    "to_python_docx",
    "DocxRedlineError",
    "TextNotFoundError",
    "AmbiguousTextError",
    "ValidationError",
    "ContinuityWarning",
    "NoteNotFoundError",
    "TextSearch",
    "TextSpan",
    "Match",
    "TrackedXMLGenerator",
    "SuggestionGenerator",
    "ScopeEvaluator",
    "EditResult",
    "AcceptResult",
    "RejectResult",
    "FormatResult",
    "ComparisonStats",
    "BatchResult",
    "Edit",
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
    # OOXML validation
    "is_ooxml_validator_available",
    "validate_with_ooxml_validator",
    "OOXMLValidationError",
    # Rendering
    "is_libreoffice_available",
    "is_pdftoppm_available",
    "is_rendering_available",
    "render_document_to_images",
    # Style templates
    "STANDARD_STYLES",
    "ensure_standard_styles",
    "get_footnote_reference_style",
    "get_footnote_text_style",
    "get_footnote_text_char_style",
    "get_endnote_reference_style",
    "get_endnote_text_style",
    "get_endnote_text_char_style",
    "get_hyperlink_style",
    # Templating
    "DocxBuilder",
]

# Import author identity
from .author import AuthorIdentity

# Import compatibility helpers (python-docx integration)
from .compat import from_python_docx, to_python_docx

# Import content type manager
from .content_types import ContentTypeManager, ContentTypes

# Import document class and standalone functions
from .document import Document, compare_documents
from .errors import (
    AmbiguousTextError,
    ContinuityWarning,
    DocxRedlineError,
    NoteNotFoundError,
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

# Import text search and match
from .match import Match

# Import model classes
from .models.comment import Comment, CommentRange
from .models.footnote import Endnote, Footnote
from .models.header_footer import Footer, Header, HeaderFooterType
from .models.paragraph import Paragraph
from .models.section import Section
from .models.tracked_change import ChangeType, TrackedChange

# Import OOXML validator
from .ooxml_validator import (
    OOXMLValidationError,
    is_ooxml_validator_available,
    validate_with_ooxml_validator,
)

# Import batch Edit class
from .operations.batch import Edit

# Import package class
from .package import OOXMLPackage

# Import relationship manager
from .relationships import RelationshipManager, RelationshipTypes

# Import rendering functionality
from .rendering import (
    is_libreoffice_available,
    is_pdftoppm_available,
    is_rendering_available,
    render_document_to_images,
)

# Import result types
from .results import (
    AcceptResult,
    BatchResult,
    ComparisonStats,
    EditResult,
    FormatResult,
    RejectResult,
)

# Import scope evaluation
from .scope import ScopeEvaluator

# Import style templates
from .style_templates import (
    STANDARD_STYLES,
    ensure_standard_styles,
    get_endnote_reference_style,
    get_endnote_text_char_style,
    get_endnote_text_style,
    get_footnote_reference_style,
    get_footnote_text_char_style,
    get_footnote_text_style,
    get_hyperlink_style,
)

# Import suggestion generator
from .suggestions import SuggestionGenerator

# Import templating
from .templating import DocxBuilder
from .text_search import TextSearch, TextSpan

# Import XML generation
from .tracked_xml import TrackedXMLGenerator
