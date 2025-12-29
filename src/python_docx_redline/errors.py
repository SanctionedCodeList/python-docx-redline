"""
Custom exception classes for python_docx_redline package.

These exceptions provide helpful error messages with suggestions for
resolving common issues when searching for and editing text in Word documents.
"""

from typing import Any


class DocxRedlineError(Exception):
    """Base exception for all python_docx_redline errors."""

    pass


class TextNotFoundError(DocxRedlineError):
    """Raised when text cannot be found in the specified scope.

    Attributes:
        text: The text that was being searched for
        scope: The scope where the search was performed (None if document-wide)
        suggestions: List of helpful suggestions for resolving the issue
        hint: Additional context about why the text wasn't found (e.g., scope filtering)
    """

    def __init__(
        self,
        text: str,
        scope: str | None = None,
        suggestions: list[str] | None = None,
        hint: str | None = None,
    ) -> None:
        self.text = text
        self.scope = scope
        self.suggestions = suggestions or []
        self.hint = hint
        super().__init__(self._format_message())

    def _format_message(self) -> str:
        """Format a helpful error message with suggestions."""
        msg = f"Could not find '{self.text}'"
        if self.scope:
            msg += f" in scope '{self.scope}'"

        if self.hint:
            msg += f"\n\nNote: {self.hint}"

        if self.suggestions:
            msg += "\n\nSuggestions:\n"
            for suggestion in self.suggestions:
                msg += f"  • {suggestion}\n"

        return msg


class AmbiguousTextError(DocxRedlineError):
    """Raised when multiple occurrences of text are found.

    Attributes:
        text: The text that was being searched for
        matches: List of TextSpan objects representing each match
    """

    def __init__(self, text: str, matches: list[Any]) -> None:
        self.text = text
        self.matches = matches
        super().__init__(self._format_message())

    def _format_message(self) -> str:
        """Format an error message showing all matches."""
        msg = f"Found {len(self.matches)} occurrences of '{self.text}'\n\n"

        for i, match in enumerate(self.matches):
            msg += f"{i + 1}. ...{match.context}...\n"
            # Location information will be added when TextSpan class is implemented
            if hasattr(match, "location") and match.location:
                if hasattr(match.location, "section"):
                    msg += f"   Section: {match.location.section}"
                if hasattr(match.location, "line_number"):
                    msg += f", Line: {match.location.line_number}"
                msg += "\n"
            msg += "\n"

        msg += "To disambiguate, either:\n"
        msg += "  • Use occurrence=N to target the Nth match (1-indexed)\n"
        msg += "  • Use occurrence='first' or occurrence='last'\n"
        msg += "  • Use occurrence='all' to target all matches\n"
        msg += "  • Provide a more specific scope parameter"
        return msg


class ValidationError(DocxRedlineError):
    """Raised when document validation fails.

    This can occur when:
    - The document structure is invalid
    - Required OOXML elements are missing
    - The document cannot be opened or saved

    Attributes:
        errors: List of specific validation error messages (optional)
    """

    def __init__(self, message: str, errors: list[str] | None = None) -> None:
        self.errors = errors or []
        super().__init__(message)


class ContinuityWarning(UserWarning):
    """Warning raised when text replacement may create a sentence fragment.

    This warning is raised when check_continuity=True in replace_tracked()
    and the text immediately following the replacement suggests a potential
    grammatical issue (e.g., sentence fragment, disconnected clause).

    Attributes:
        message: Description of the potential continuity issue
        next_text: The text immediately following the replacement
        suggestions: List of suggestions for fixing the issue
    """

    def __init__(self, message: str, next_text: str, suggestions: list[str] | None = None) -> None:
        self.message = message
        self.next_text = next_text
        self.suggestions = suggestions or []
        super().__init__(self._format_message())

    def _format_message(self) -> str:
        """Format a helpful warning message with suggestions."""
        msg = f"{self.message}\n"
        msg += f"Next text begins with: {repr(self.next_text[:50])}\n"

        if self.suggestions:
            msg += "\nSuggestions:\n"
            for suggestion in self.suggestions:
                msg += f"  • {suggestion}\n"

        return msg


class RefNotFoundError(DocxRedlineError):
    """Raised when a ref cannot be resolved to a document element.

    This error occurs when:
    - The ref format is invalid
    - The ordinal index is out of bounds
    - The element type is not supported
    - No element matches a fingerprint ref

    Attributes:
        ref: The ref path that could not be resolved
        reason: Explanation of why resolution failed
    """

    def __init__(self, ref: str, reason: str | None = None) -> None:
        self.ref = ref
        self.reason = reason
        super().__init__(self._format_message())

    def _format_message(self) -> str:
        """Format an error message with resolution details."""
        msg = f"Could not resolve ref '{self.ref}'"
        if self.reason:
            msg += f": {self.reason}"
        return msg


class StaleRefError(DocxRedlineError):
    """Raised when a ref points to an element that has been modified or deleted.

    This error occurs when using fingerprint-based refs after the document
    structure has changed. The fingerprint was previously valid but the
    element content has been modified.

    Attributes:
        ref: The stale ref path
        reason: Explanation of the staleness
    """

    def __init__(self, ref: str, reason: str | None = None) -> None:
        self.ref = ref
        self.reason = reason
        super().__init__(self._format_message())

    def _format_message(self) -> str:
        """Format an error message indicating staleness."""
        msg = f"Stale ref '{self.ref}'"
        if self.reason:
            msg += f": {self.reason}"
        msg += (
            "\n\nThe document structure has changed. "
            "Regenerate the accessibility tree to get updated refs."
        )
        return msg


class NoteNotFoundError(DocxRedlineError):
    """Raised when a footnote or endnote cannot be found by ID.

    This error occurs when attempting to access, edit, or delete a footnote
    or endnote that does not exist in the document.

    Attributes:
        note_type: The type of note ('footnote' or 'endnote')
        note_id: The ID that was searched for
        available_ids: List of valid IDs in the document
    """

    def __init__(
        self,
        note_type: str,
        note_id: str | int,
        available_ids: list[str] | None = None,
    ) -> None:
        self.note_type = note_type
        self.note_id = str(note_id)
        self.available_ids = available_ids or []
        super().__init__(self._format_message())

    def _format_message(self) -> str:
        """Format an error message with available IDs."""
        msg = f"{self.note_type.capitalize()} with ID '{self.note_id}' not found"
        if self.available_ids:
            ids_str = ", ".join(self.available_ids)
            msg += f"\n\nAvailable {self.note_type} IDs: {ids_str}"
        else:
            msg += f"\n\nNo {self.note_type}s exist in the document"
        return msg
