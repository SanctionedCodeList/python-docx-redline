"""
Operations package for Document manipulation.

This package contains classes that handle specific operations on Word documents,
extracted from the main Document class to improve separation of concerns.
"""

from .batch import BatchOperations
from .change_management import ChangeManagement
from .comments import CommentOperations
from .comparison import ComparisonOperations
from .formatting import FormatOperations
from .header_footer import HeaderFooterOperations
from .hyperlinks import HyperlinkInfo, HyperlinkOperations
from .images import ImageOperations
from .notes import NoteOperations
from .patterns import PatternOperations
from .section import SectionOperations
from .tables import TableOperations
from .tracked_changes import TrackedChangeOperations

__all__ = [
    "BatchOperations",
    "ChangeManagement",
    "CommentOperations",
    "ComparisonOperations",
    "FormatOperations",
    "HeaderFooterOperations",
    "HyperlinkInfo",
    "HyperlinkOperations",
    "ImageOperations",
    "NoteOperations",
    "PatternOperations",
    "SectionOperations",
    "TableOperations",
    "TrackedChangeOperations",
]
