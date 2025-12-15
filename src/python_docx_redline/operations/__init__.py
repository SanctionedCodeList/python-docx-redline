"""
Operations package for Document manipulation.

This package contains classes that handle specific operations on Word documents,
extracted from the main Document class to improve separation of concerns.
"""

from .change_management import ChangeManagement
from .comments import CommentOperations
from .formatting import FormatOperations
from .tables import TableOperations
from .tracked_changes import TrackedChangeOperations

__all__ = [
    "ChangeManagement",
    "CommentOperations",
    "FormatOperations",
    "TableOperations",
    "TrackedChangeOperations",
]
