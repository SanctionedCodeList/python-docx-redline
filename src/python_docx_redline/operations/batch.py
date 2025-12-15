"""
BatchOperations class for handling batch edit operations.

This module provides a dedicated class for all batch edit operations,
extracted from the main Document class to improve separation of concerns.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Any

import yaml

from ..errors import AmbiguousTextError, TextNotFoundError, ValidationError
from ..results import EditResult

if TYPE_CHECKING:
    from ..document import Document


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
        self, edits: list[dict[str, Any]], stop_on_error: bool = False
    ) -> list[EditResult]:
        """Apply multiple edits in sequence.

        This method processes a list of edit specifications and applies each one
        in order. Each edit is a dictionary specifying the edit type and parameters.

        Args:
            edits: List of edit dictionaries with keys:
                - type: Edit operation ("insert_tracked", "replace_tracked", "delete_tracked")
                - Other parameters specific to the edit type
            stop_on_error: If True, stop processing on first error

        Returns:
            List of EditResult objects, one per edit

        Example:
            >>> edits = [
            ...     {
            ...         "type": "insert_tracked",
            ...         "text": "new text",
            ...         "after": "anchor",
            ...         "scope": "section:Introduction"
            ...     },
            ...     {
            ...         "type": "replace_tracked",
            ...         "find": "old",
            ...         "replace": "new"
            ...     }
            ... ]
            >>> results = doc.apply_edits(edits)
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
                result = self._apply_single_edit(edit_type, edit)
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

    def _apply_single_edit(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Apply a single edit operation.

        Args:
            edit_type: The type of edit to perform
            edit: Dictionary with edit parameters

        Returns:
            EditResult indicating success or failure
        """
        # Dispatch table mapping edit types to handler methods
        handlers = {
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
            return handler(edit_type, edit)
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

    def _handle_insert_tracked(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Handle insert_tracked edit type."""
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

    def _handle_delete_tracked(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Handle delete_tracked edit type."""
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

    def _handle_replace_tracked(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Handle replace_tracked edit type."""
        find = edit.get("find")
        replace = edit.get("replace")
        author = edit.get("author")
        scope = edit.get("scope")
        regex = edit.get("regex", False)

        if not find or replace is None:
            return EditResult(
                success=False,
                edit_type=edit_type,
                message="Missing required parameter: 'find' or 'replace'",
                error=ValidationError("Missing required parameter"),
            )

        self._document.replace_tracked(find, replace, author=author, scope=scope, regex=regex)
        return EditResult(
            success=True,
            edit_type=edit_type,
            message=f"Replaced '{find}' with '{replace}'",
        )

    def _handle_insert_paragraph(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Handle insert_paragraph edit type."""
        text = edit.get("text")
        after = edit.get("after")
        before = edit.get("before")
        style = edit.get("style")
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

    def _handle_insert_paragraphs(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Handle insert_paragraphs edit type."""
        texts = edit.get("texts")
        after = edit.get("after")
        before = edit.get("before")
        style = edit.get("style")
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

    def _handle_delete_section(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Handle delete_section edit type."""
        heading = edit.get("heading")
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

    def _handle_format_tracked(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Handle format_tracked edit type."""
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

    def _handle_format_paragraph_tracked(self, edit_type: str, edit: dict[str, Any]) -> EditResult:
        """Handle format_paragraph_tracked edit type."""
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
        self, path: str | Path, format: str = "yaml", stop_on_error: bool = False
    ) -> list[EditResult]:
        """Apply edits from a YAML or JSON file.

        Loads edit specifications from a file and applies them using apply_edits().
        The file should contain an 'edits' key with a list of edit dictionaries.

        Args:
            path: Path to the edit specification file
            format: File format - "yaml" or "json" (default: "yaml")
            stop_on_error: If True, stop processing on first error

        Returns:
            List of EditResult objects, one per edit

        Raises:
            ValidationError: If file cannot be parsed or has invalid format
            FileNotFoundError: If file does not exist

        Example YAML file:
            ```yaml
            edits:
              - type: insert_tracked
                text: "new text"
                after: "anchor"
              - type: replace_tracked
                find: "old"
                replace: "new"
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

            # Apply the edits
            return self.apply_edits(edits, stop_on_error=stop_on_error)

        except yaml.YAMLError as e:
            raise ValidationError(f"Failed to parse YAML file: {e}") from e
        except Exception as e:
            if isinstance(e, ValidationError | FileNotFoundError):
                raise
            raise ValidationError(f"Failed to load edit file: {e}") from e
