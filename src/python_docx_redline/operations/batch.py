"""
BatchOperations class for handling batch edit operations.

This module provides a dedicated class for all batch edit operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

import yaml

from ..errors import AmbiguousTextError, TextNotFoundError, ValidationError
from ..results import BatchResult, EditResult
from ..suggestions import SuggestionGenerator

if TYPE_CHECKING:
    from ..document import Document


@dataclass
class Edit:
    """Represents a single edit operation for batch processing.

    This dataclass provides a structured way to specify edits instead of
    using dictionaries. It supports both simple find/replace operations
    and more complex edit types.

    Attributes:
        old: The text to find/replace (required for replace operations)
        new: The replacement text (required for replace operations)
        edit_type: Type of operation (default: "replace")
        track: Whether to track the change (default: True)
        author: Optional author name for the change
        scope: Optional scope to limit the search
        regex: Whether old is a regex pattern (default: False)
        occurrence: Which occurrence to target (default: "first")

    Example:
        >>> edit = Edit(old="old text", new="new text")
        >>> results = doc.apply_edits_batch([edit])
    """

    old: str
    new: str
    edit_type: str = "replace"
    track: bool = True
    author: str | None = None
    scope: str | None = None
    regex: bool = False
    occurrence: str | int = "first"

    def to_dict(self) -> dict[str, Any]:
        """Convert the Edit to a dictionary format."""
        result: dict[str, Any] = {
            "type": self.edit_type,
            "find": self.old,
            "replace": self.new,
            "track": self.track,
            "regex": self.regex,
            "occurrence": self.occurrence,
        }
        if self.author:
            result["author"] = self.author
        if self.scope:
            result["scope"] = self.scope
        return result


# Type alias for edit input formats
EditInput = tuple[str, str] | Edit | dict[str, Any]


class BatchOperations:
    """Handles batch edit operations.

    This class encapsulates all batch edit functionality, including:
    - Applying multiple edits from a list
    - Applying edits from YAML or JSON files
    - Dispatching edits to appropriate handlers

    The class takes a Document reference and operates on its methods.

    Example:
        >>> # Usually accessed through Document
        >>> doc = Document("contract.docx")
        >>> edits = [
        ...     {"type": "insert_tracked", "text": "new", "after": "old"},
        ...     {"type": "replace_tracked", "find": "old", "replace": "new"}
        ... ]
        >>> results = doc.apply_edits(edits)
    """

    def __init__(self, document: Document) -> None:
        """Initialize BatchOperations with a Document reference.

        Args:
            document: The Document instance to operate on
        """
        self._document = document

    def apply_edits(
        self,
        edits: list[dict[str, Any]],
        stop_on_error: bool = False,
        default_track: bool = False,
    ) -> list[EditResult]:
        """Apply multiple edits in sequence.

        This method processes a list of edit specifications and applies each one
        in order. Each edit is a dictionary specifying the edit type and parameters.

        Args:
            edits: List of edit dictionaries with keys:
                - type: Edit operation ("insert", "delete", "replace",
                    "insert_tracked", "replace_tracked", "delete_tracked")
                - track: Optional boolean to control tracking per-edit
                - Other parameters specific to the edit type
            stop_on_error: If True, stop processing on first error
            default_track: Default value for 'track' if not specified per-edit
                (default: False). Note: *_tracked operations always track regardless.

        Returns:
            List of EditResult objects, one per edit

        Example:
            >>> edits = [
            ...     {
            ...         "type": "insert",
            ...         "text": "new text",
            ...         "after": "anchor",
            ...         "track": True,  # This edit is tracked
            ...     },
            ...     {
            ...         "type": "replace",
            ...         "find": "old",
            ...         "replace": "new",
            ...         # Uses default_track value
            ...     }
            ... ]
            >>> results = doc.apply_edits(edits, default_track=False)
            >>> print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
        """
        results = []

        for i, edit in enumerate(edits):
            edit_type = edit.get("type")
            if not edit_type:
                results.append(
                    EditResult(
                        success=False,
                        edit_type="unknown",
                        message=f"Edit {i}: Missing 'type' field",
                        error=ValidationError("Missing 'type' field"),
                    )
                )
                if stop_on_error:
                    break
                continue

            try:
                result = self._apply_single_edit(edit_type, edit, default_track)
                results.append(result)

                if not result.success and stop_on_error:
                    break

            except Exception as e:
                results.append(
                    EditResult(
                        success=False,
                        edit_type=edit_type,
                        message=f"Error: {str(e)}",
                        error=e,
                    )
                )
                if stop_on_error:
                    break

        return results

    def apply_edits_batch(
        self,
        edits: list[EditInput],
        continue_on_error: bool = True,
        default_track: bool = True,
        dry_run: bool = False,
    ) -> BatchResult:
        """Apply multiple edits with partial success support and error reporting.

        This enhanced version of apply_edits provides:
        - Support for tuple input: (old_text, new_text)
        - Support for Edit dataclass objects
        - BatchResult with succeeded/failed lists
        - Suggestions for failed edits (similar text that exists)
        - Dry run mode to preview changes

        Args:
            edits: List of edits in any of these formats:
                - Tuple: (old_text, new_text) - applies as tracked replace
                - Edit object: Edit(old="...", new="...")
                - Dictionary: {"type": "replace", "find": "...", "replace": "..."}
            continue_on_error: If True (default), continue processing after errors
            default_track: Default value for 'track' if not specified (default: True)
            dry_run: If True, validate edits without making changes

        Returns:
            BatchResult with succeeded and failed edit lists

        Example:
            >>> edits = [
            ...     ("old phrase", "new phrase"),  # Simple tuple
            ...     Edit(old="another", new="replacement"),  # Edit object
            ...     {"type": "replace_tracked", "find": "x", "replace": "y"},  # Dict
            ... ]
            >>> results = doc.apply_edits_batch(edits)
            >>> print(results)
            # Shows summary with succeeded/failed counts
            >>> if not results:
            ...     print("Some edits failed:", results.failed)
        """
        # Normalize all edits to dictionary format
        normalized_edits = self._normalize_edits(edits, default_track)

        batch_result = BatchResult(dry_run=dry_run)

        for i, (edit_dict, old_text, new_text) in enumerate(normalized_edits):
            edit_type = edit_dict.get("type", "replace")

            # For dry run, just validate without applying
            if dry_run:
                result = self._validate_edit(edit_type, edit_dict)
                result.index = i
                result.old_text = old_text
                result.new_text = new_text
            else:
                try:
                    result = self._apply_single_edit(edit_type, edit_dict, default_track)
                    result.index = i
                    result.old_text = old_text
                    result.new_text = new_text
                except Exception as e:
                    result = EditResult(
                        success=False,
                        edit_type=edit_type,
                        message=f"Error: {str(e)}",
                        error=e,
                        index=i,
                        old_text=old_text,
                        new_text=new_text,
                    )

            # Add suggestions for failed edits with TextNotFoundError
            if not result.success and isinstance(result.error, TextNotFoundError):
                result.suggestions = self._get_suggestions_for_error(old_text, result.error)

            # Categorize result
            if result.success:
                batch_result.succeeded.append(result)
            else:
                batch_result.failed.append(result)
                if not continue_on_error:
                    break

        return batch_result

    def _normalize_edits(
        self,
        edits: list[EditInput],
        default_track: bool,
    ) -> list[tuple[dict[str, Any], str | None, str | None]]:
        """Normalize various edit input formats to dictionary format.

        Returns list of (edit_dict, old_text, new_text) tuples.
        """
        normalized = []

        for edit in edits:
            if isinstance(edit, tuple):
                # Tuple format: (old, new) -> tracked replace
                old_text, new_text = edit
                edit_dict = {
                    "type": "replace",
                    "find": old_text,
                    "replace": new_text,
                    "track": default_track,
                }
                normalized.append((edit_dict, old_text, new_text))

            elif isinstance(edit, Edit):
                # Edit dataclass
                edit_dict = edit.to_dict()
                normalized.append((edit_dict, edit.old, edit.new))

            elif isinstance(edit, dict):
                # Dictionary format (existing)
                old_text = edit.get("find") or edit.get("text")
                new_text = edit.get("replace")
                normalized.append((edit, old_text, new_text))

            else:
                raise ValidationError(
                    f"Invalid edit format: expected tuple, Edit, or dict, got {type(edit)}"
                )

        return normalized

    def _validate_edit(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Validate an edit without applying it (for dry run mode).

        Checks if the text exists in the document without making changes.
        """
        # Get the search text
        search_text = edit.get("find") or edit.get("text") or edit.get("after")

        if not search_text:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="No search text specified",
                error=ValidationError("No search text specified"),
            )

        # Try to find the text in the document
        try:
            scope = edit.get("scope")
            regex = edit.get("regex", False)
            occurrence = edit.get("occurrence", "first")

            # Use the document's text search to validate
            paragraphs = self._document._get_paragraphs_for_scope(scope)
            matches = self._document._text_search.find_text(search_text, paragraphs, regex=regex)

            if not matches:
                return EditResult(
                    success=False,
                    edit_type=edit_type,
                    message=f"Text not found: '{search_text}'",
                    error=TextNotFoundError(search_text, scope),
                )

            # Check for ambiguity
            if len(matches) > 1 and occurrence == "first":
                # Would succeed with first occurrence
                pass
            elif occurrence == "all":
                pass
            elif isinstance(occurrence, int) and occurrence > len(matches):
                return EditResult(
                    success=False,
                    edit_type=edit_type,
                    message=f"Occurrence {occurrence} not found (only {len(matches)} matches)",
                    error=AmbiguousTextError(search_text, matches),
                )

            return EditResult(
                success=True,
                edit_type=edit_type,
                message=f"(dry run) Would apply: {edit_type}",
            )

        except TextNotFoundError as e:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Text not found: {e}",
                error=e,
            )
        except AmbiguousTextError as e:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Ambiguous text: {e}",
                error=e,
            )
        except Exception as e:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Validation error: {str(e)}",
                error=e,
            )

    def _get_suggestions_for_error(
        self, search_text: str | None, error: TextNotFoundError
    ) -> list[str]:
        """Get similar text suggestions for a TextNotFoundError."""
        if not search_text:
            return []

        try:
            paragraphs = self._document._get_paragraphs_for_scope(error.scope)
            return SuggestionGenerator.find_similar_text(search_text, paragraphs, max_suggestions=3)
        except Exception:
            # If suggestion generation fails, just return empty list
            return []

    def _apply_single_edit(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Apply a single edit operation.

        Args:
            edit_type: The type of edit to perform
            edit: Dictionary with edit parameters
            default_track: Default value for 'track' if not specified in edit

        Returns:
            EditResult indicating success or failure
        """
        # Dispatch table mapping edit types to handler methods
        handlers = {
            # Generic operations (use track field or default_track)
            "insert": self._handle_insert,
            "delete": self._handle_delete,
            "replace": self._handle_replace,
            # Legacy tracked operations (always tracked)
            "insert_tracked": self._handle_insert_tracked,
            "delete_tracked": self._handle_delete_tracked,
            "replace_tracked": self._handle_replace_tracked,
            "insert_paragraph": self._handle_insert_paragraph,
            "insert_paragraphs": self._handle_insert_paragraphs,
            "delete_section": self._handle_delete_section,
            "format_tracked": self._handle_format_tracked,
            "format_paragraph_tracked": self._handle_format_paragraph_tracked,
        }

        handler = handlers.get(edit_type)
        if handler is None:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Unknown edit type: {edit_type}",
                error=ValidationError(f"Unknown edit type: {edit_type}"),
            )

        try:
            return handler(edit_type, edit, default_track)
        except TextNotFoundError as e:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Text not found: {e}",
                error=e,
            )
        except AmbiguousTextError as e:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message=f"Ambiguous text: {e}",
                error=e,
            )
        except Exception as e:
            return EditResult(
                success=False, edit_type=edit_type, message=f"Error: {str(e)}", error=e
            )

    def _get_track_value(self, edit: dict[str, Any], default_track: bool) -> bool:
        """Get the track value for an edit, using per-edit or default.

        Args:
            edit: The edit dictionary
            default_track: Default value if 'track' not in edit

        Returns:
            Boolean indicating whether to track the change
        """
        return edit.get("track", default_track)

    def _handle_insert(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle generic insert edit type."""
        text = edit.get("text")
        after = edit.get("after")
        before = edit.get("before")
        author = edit.get("author")
        scope = edit.get("scope")
        regex = edit.get("regex", False)
        occurrence = edit.get("occurrence", "first")
        track = self._get_track_value(edit, default_track)

        if not text:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'text'",
                error=ValidationError("Missing required parameter"),
            )

        if not after and not before:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'after' or 'before'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.insert(
            text,
            after=after,
            before=before,
            author=author,
            scope=scope,
            regex=regex,
            occurrence=occurrence,
            track=track,
        )
        location = f"after '{after}'" if after else f"before '{before}'"
        track_msg = " (tracked)" if track else ""
        return EditResult(
            success=True,
            edit_type=edit_type,
            message=f"Inserted '{text}' {location}{track_msg}",
        )

    def _handle_delete(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle generic delete edit type."""
        text = edit.get("text")
        author = edit.get("author")
        scope = edit.get("scope")
        regex = edit.get("regex", False)
        occurrence = edit.get("occurrence", "first")
        track = self._get_track_value(edit, default_track)

        if not text:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'text'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.delete(
            text,
            author=author,
            scope=scope,
            regex=regex,
            occurrence=occurrence,
            track=track,
        )
        track_msg = " (tracked)" if track else ""
        return EditResult(
            success=True,
            edit_type=edit_type,
            message=f"Deleted '{text}'{track_msg}",
        )

    def _handle_replace(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle generic replace edit type."""
        find = edit.get("find")
        replace = edit.get("replace")
        author = edit.get("author")
        scope = edit.get("scope")
        regex = edit.get("regex", False)
        occurrence = edit.get("occurrence", "first")
        track = self._get_track_value(edit, default_track)
        minimal = edit.get("minimal")  # None = use document default

        if not find or replace is None:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'find' or 'replace'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.replace(
            find,
            replace,
            author=author,
            scope=scope,
            regex=regex,
            occurrence=occurrence,
            track=track,
            minimal=minimal,
        )
        track_msg = " (tracked)" if track else ""
        return EditResult(
            success=True,
            edit_type=edit_type,
            message=f"Replaced '{find}' with '{replace}'{track_msg}",
        )

    def _handle_insert_tracked(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle insert_tracked edit type (always tracked)."""
        text = edit.get("text")
        after = edit.get("after")
        author = edit.get("author")
        scope = edit.get("scope")
        regex = edit.get("regex", False)

        if not text or not after:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'text' or 'after'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.insert_tracked(text, after, author=author, scope=scope, regex=regex)
        return EditResult(
            success=True,
            edit_type=edit_type,
            message=f"Inserted '{text}' after '{after}'",
        )

    def _handle_delete_tracked(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle delete_tracked edit type (always tracked)."""
        text = edit.get("text")
        author = edit.get("author")
        scope = edit.get("scope")
        regex = edit.get("regex", False)

        if not text:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'text'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.delete_tracked(text, author=author, scope=scope, regex=regex)
        return EditResult(success=True, edit_type=edit_type, message=f"Deleted '{text}'")

    def _handle_replace_tracked(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle replace_tracked edit type (always tracked)."""
        find = edit.get("find")
        replace = edit.get("replace")
        author = edit.get("author")
        scope = edit.get("scope")
        regex = edit.get("regex", False)
        minimal = edit.get("minimal")  # None = use document default

        if not find or replace is None:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'find' or 'replace'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.replace_tracked(
            find, replace, author=author, scope=scope, regex=regex, minimal=minimal
        )
        return EditResult(
            success=True,
            edit_type=edit_type,
            message=f"Replaced '{find}' with '{replace}'",
        )

    def _handle_insert_paragraph(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle insert_paragraph edit type."""
        text = edit.get("text")
        after = edit.get("after")
        before = edit.get("before")
        style = edit.get("style")
        # For insert_paragraph, use per-edit track if specified, else use True for
        # backwards compatibility (insert_paragraph historically defaults to tracked)
        track = edit.get("track", True)
        author = edit.get("author")
        scope = edit.get("scope")

        if not text:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'text'",
                error=ValidationError("Missing required parameter"),
            )

        if not after and not before:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'after' or 'before'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.insert_paragraph(
            text,
            after=after,
            before=before,
            style=style,
            track=track,
            author=author,
            scope=scope,
        )
        location = f"after '{after}'" if after else f"before '{before}'"
        return EditResult(
            success=True,
            edit_type=edit_type,
            message=f"Inserted paragraph '{text}' {location}",
        )

    def _handle_insert_paragraphs(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle insert_paragraphs edit type."""
        texts = edit.get("texts")
        after = edit.get("after")
        before = edit.get("before")
        style = edit.get("style")
        # For insert_paragraphs, use per-edit track if specified, else use True for
        # backwards compatibility (insert_paragraphs historically defaults to tracked)
        track = edit.get("track", True)
        author = edit.get("author")
        scope = edit.get("scope")

        if not texts:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'texts'",
                error=ValidationError("Missing required parameter"),
            )

        if not after and not before:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'after' or 'before'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.insert_paragraphs(
            texts,
            after=after,
            before=before,
            style=style,
            track=track,
            author=author,
            scope=scope,
        )
        location = f"after '{after}'" if after else f"before '{before}'"
        return EditResult(
            success=True,
            edit_type=edit_type,
            message=f"Inserted {len(texts)} paragraphs {location}",
        )

    def _handle_delete_section(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle delete_section edit type."""
        heading = edit.get("heading")
        # For delete_section, use per-edit track if specified, else use True for
        # backwards compatibility (delete_section historically defaults to tracked)
        track = edit.get("track", True)
        update_toc = edit.get("update_toc", False)
        author = edit.get("author")
        scope = edit.get("scope")

        if not heading:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'heading'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.delete_section(
            heading, track=track, update_toc=update_toc, author=author, scope=scope
        )
        return EditResult(
            success=True,
            edit_type=edit_type,
            message=f"Deleted section '{heading}'",
        )

    def _handle_format_tracked(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle format_tracked edit type (always tracked)."""
        text = edit.get("text")
        if not text:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'text'",
                error=ValidationError("Missing required parameter"),
            )

        format_params = self._extract_character_format_params(edit)

        if not format_params:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="At least one formatting parameter required",
                error=ValidationError("Missing formatting parameter"),
            )

        result = self._document.format_tracked(
            text,
            scope=edit.get("scope"),
            occurrence=edit.get("occurrence", "first"),
            author=edit.get("author"),
            **format_params,
        )
        return EditResult(
            success=result.success,
            edit_type=edit_type,
            message=f"Formatted '{text}' with {format_params}",
        )

    def _handle_format_paragraph_tracked(
        self, edit_type: str, edit: dict[str, Any], default_track: bool = False
    ) -> EditResult:
        """Handle format_paragraph_tracked edit type (always tracked)."""
        containing = edit.get("containing")
        starting_with = edit.get("starting_with")
        ending_with = edit.get("ending_with")
        index = edit.get("index")

        if not any([containing, starting_with, ending_with, index is not None]):
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="At least one targeting parameter required",
                error=ValidationError("Missing targeting parameter"),
            )

        format_params = self._extract_paragraph_format_params(edit)

        if not format_params:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="At least one formatting parameter required",
                error=ValidationError("Missing formatting parameter"),
            )

        result = self._document.format_paragraph_tracked(
            containing=containing,
            starting_with=starting_with,
            ending_with=ending_with,
            index=index,
            scope=edit.get("scope"),
            author=edit.get("author"),
            **format_params,
        )
        target_desc = containing or starting_with or ending_with or f"index {index}"
        return EditResult(
            success=result.success,
            edit_type=edit_type,
            message=f"Formatted paragraph '{target_desc}' with {format_params}",
        )

    def _extract_character_format_params(self, edit: dict[str, Any]) -> dict[str, Any]:
        """Extract character formatting parameters from an edit dict."""
        format_keys = (
            "bold",
            "italic",
            "underline",
            "strikethrough",
            "font_name",
            "font_size",
            "color",
            "highlight",
            "superscript",
            "subscript",
            "small_caps",
            "all_caps",
        )
        return {k: v for k, v in edit.items() if k in format_keys and v is not None}

    def _extract_paragraph_format_params(self, edit: dict[str, Any]) -> dict[str, Any]:
        """Extract paragraph formatting parameters from an edit dict."""
        format_keys = (
            "alignment",
            "spacing_before",
            "spacing_after",
            "line_spacing",
            "indent_left",
            "indent_right",
            "indent_first_line",
            "indent_hanging",
        )
        return {k: v for k, v in edit.items() if k in format_keys and v is not None}

    def apply_edit_file(
        self,
        path: str | Path,
        format: str = "yaml",
        stop_on_error: bool = False,
        default_track: bool | None = None,
    ) -> list[EditResult]:
        """Apply edits from a YAML or JSON file.

        Loads edit specifications from a file and applies them using apply_edits().
        The file should contain an 'edits' key with a list of edit dictionaries.

        Args:
            path: Path to the edit specification file
            format: File format - "yaml" or "json" (default: "yaml")
            stop_on_error: If True, stop processing on first error
            default_track: Default value for 'track' if not specified per-edit.
                If None, uses file's default_track value (or False if not set).
                If specified, overrides any default_track in the file.

        Returns:
            List of EditResult objects, one per edit

        Raises:
            ValidationError: If file cannot be parsed or has invalid format
            FileNotFoundError: If file does not exist

        Example YAML file:
            ```yaml
            default_track: false  # Global default for edits in this file

            edits:
              - type: insert
                text: "new text"
                after: "anchor"
                # Uses default_track: false

              - type: replace
                find: "old"
                replace: "new"
                track: true  # Override for this specific edit
            ```

        Example:
            >>> results = doc.apply_edit_file("edits.yaml")
            >>> print(f"Applied {sum(r.success for r in results)}/{len(results)} edits")
        """
        file_path = Path(path)

        if not file_path.exists():
            raise FileNotFoundError(f"Edit file not found: {path}")

        try:
            with open(file_path, encoding="utf-8") as f:
                if format == "yaml":
                    data = yaml.safe_load(f)
                elif format == "json":
                    import json

                    data = json.load(f)
                else:
                    raise ValidationError(f"Unsupported format: {format}")

            if not isinstance(data, dict):
                raise ValidationError("Edit file must contain a dictionary/object")

            if "edits" not in data:
                raise ValidationError("Edit file must contain an 'edits' key")

            edits = data["edits"]
            if not isinstance(edits, list):
                raise ValidationError("'edits' must be a list")

            # Determine default_track value:
            # 1. If caller specified default_track, use it
            # 2. Otherwise, use file's default_track if present
            # 3. Fall back to False
            if default_track is not None:
                file_default_track = default_track
            else:
                file_default_track = data.get("default_track", False)

            # Apply the edits
            return self.apply_edits(
                edits, stop_on_error=stop_on_error, default_track=file_default_track
            )

        except yaml.YAMLError as e:
            raise ValidationError(f"Failed to parse YAML file: {e}") from e
        except Exception as e:
            if isinstance(e, ValidationError | FileNotFoundError):
                raise
            raise ValidationError(f"Failed to load edit file: {e}") from e
