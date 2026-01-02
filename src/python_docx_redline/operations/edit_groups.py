"""Edit group management for batch rejection of tracked changes.

This module provides the EditGroupRegistry class for grouping related edits
together and enabling batch rejection.

Example:
    >>> with doc.edit_group('condensing round 1'):
    ...     doc.replace_tracked('long text', 'short')
    ...     doc.replace_tracked('another section', 'condensed')
    >>> # Later, reject all changes in that group
    >>> doc.reject_edit_group('condensing round 1')
"""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    pass


@dataclass
class EditGroup:
    """Represents a named group of related edits.

    Attributes:
        name: Unique identifier for this group
        status: Current status ('active', 'completed', 'rejected')
        change_ids: List of change IDs belonging to this group
        created_at: When the group was created
    """

    name: str
    status: str  # 'active', 'completed', 'rejected'
    change_ids: list[int] = field(default_factory=list)
    created_at: datetime = field(default_factory=datetime.now)


class EditGroupRegistry:
    """Registry for managing edit groups.

    This class tracks edit groups and their associated change IDs, enabling
    batch operations like rejecting all changes in a specific group.

    Example:
        >>> registry = EditGroupRegistry()
        >>> registry.start_group("round1")
        >>> registry.add_change_id(1)
        >>> registry.add_change_id(2)
        >>> registry.end_group()
        >>> registry.get_group_ids("round1")
        [1, 2]
    """

    def __init__(self) -> None:
        """Initialize an empty edit group registry."""
        self._groups: dict[str, EditGroup] = {}
        self._active_group: str | None = None

    @property
    def active_group(self) -> str | None:
        """Get the name of the currently active group, if any.

        Returns:
            Name of the active group, or None if no group is active
        """
        return self._active_group

    def start_group(self, name: str) -> None:
        """Start a new edit group.

        Args:
            name: Unique name for the group

        Raises:
            ValueError: If another group is already active
            ValueError: If a group with this name already exists
        """
        if self._active_group is not None:
            raise ValueError(f"Group '{self._active_group}' already active")
        if name in self._groups:
            raise ValueError(f"Group '{name}' already exists")

        self._groups[name] = EditGroup(name=name, status="active")
        self._active_group = name

    def end_group(self) -> None:
        """End the currently active group.

        Marks the active group as 'completed'. Does nothing if no group is active.
        """
        if self._active_group is not None:
            self._groups[self._active_group].status = "completed"
            self._active_group = None

    def add_change_id(self, change_id: int) -> None:
        """Add a change ID to the currently active group.

        Args:
            change_id: The ID of the tracked change to add

        Note:
            Does nothing if no group is currently active.
        """
        if self._active_group is not None:
            self._groups[self._active_group].change_ids.append(change_id)

    def get_group_ids(self, name: str) -> list[int]:
        """Get all change IDs in a named group.

        Args:
            name: The name of the group

        Returns:
            Copy of the list of change IDs in the group

        Raises:
            ValueError: If no group with the given name exists
        """
        if name not in self._groups:
            raise ValueError(f"No group '{name}' found")
        return self._groups[name].change_ids.copy()

    def mark_rejected(self, name: str) -> None:
        """Mark a group as rejected.

        Args:
            name: The name of the group to mark as rejected

        Note:
            Does nothing if the group doesn't exist.
        """
        if name in self._groups:
            self._groups[name].status = "rejected"

    def group_exists(self, name: str) -> bool:
        """Check if a group with the given name exists.

        Args:
            name: The name to check

        Returns:
            True if the group exists, False otherwise
        """
        return name in self._groups

    def get_group_status(self, name: str) -> str | None:
        """Get the status of a group.

        Args:
            name: The name of the group

        Returns:
            The group's status ('active', 'completed', 'rejected'),
            or None if the group doesn't exist
        """
        if name in self._groups:
            return self._groups[name].status
        return None

    def list_groups(self) -> list[str]:
        """List all group names.

        Returns:
            List of all group names in the registry
        """
        return list(self._groups.keys())


class EditGroupMixin:
    """Mixin class providing edit group functionality for Document.

    This mixin adds the edit_group context manager and reject_edit_group
    method to the Document class.
    """

    # Type hints for attributes that should exist on Document
    _edit_groups: EditGroupRegistry
    _change_mgmt: ChangeManagement  # type: ignore[name-defined]  # noqa: F821

    def _ensure_edit_groups(self) -> None:
        """Ensure the edit groups registry exists."""
        if not hasattr(self, "_edit_groups") or self._edit_groups is None:
            self._edit_groups = EditGroupRegistry()

    @contextmanager
    def edit_group(self, name: str) -> Iterator[None]:
        """Context manager for grouping related edits.

        All tracked changes made within this context will be associated with
        the named group. This enables batch operations like rejecting all
        changes in a group at once.

        Args:
            name: Unique name for this edit group

        Yields:
            None

        Raises:
            ValueError: If another group is already active
            ValueError: If a group with this name already exists

        Example:
            >>> with doc.edit_group('condensing round 1'):
            ...     doc.replace_tracked('long text', 'short')
            ...     doc.replace_tracked('another section', 'condensed')
        """
        self._ensure_edit_groups()
        self._edit_groups.start_group(name)
        try:
            yield
        finally:
            self._edit_groups.end_group()

    def reject_edit_group(self, group_name: str) -> int:
        """Reject all tracked changes in an edit group.

        This method finds all change IDs associated with the named group
        and rejects them in reverse order (to properly handle nested changes).

        Args:
            group_name: Name of the edit group to reject

        Returns:
            Number of changes successfully rejected

        Raises:
            ValueError: If no group with the given name exists

        Example:
            >>> with doc.edit_group('round1'):
            ...     doc.replace_tracked('old', 'new')
            ...     doc.insert_tracked(' extra', after='text')
            >>> # Later, reject all changes from round1
            >>> count = doc.reject_edit_group('round1')
            >>> print(f"Rejected {count} changes")
        """
        self._ensure_edit_groups()
        ids = self._edit_groups.get_group_ids(group_name)
        count = 0

        # Reject in reverse order to handle nested elements properly
        for change_id in reversed(ids):
            try:
                self._change_mgmt.reject_change(change_id)
                count += 1
            except ValueError:
                # Change may have already been rejected or no longer exists
                pass

        self._edit_groups.mark_rejected(group_name)
        return count

    def accept_edit_group(self, group_name: str) -> int:
        """Accept all tracked changes in an edit group.

        This method finds all change IDs associated with the named group
        and accepts them in reverse order (to properly handle nested changes).

        Args:
            group_name: Name of the edit group to accept

        Returns:
            Number of changes successfully accepted

        Raises:
            ValueError: If no group with the given name exists

        Example:
            >>> with doc.edit_group('round1'):
            ...     doc.replace_tracked('old', 'new')
            ...     doc.insert_tracked(' extra', after='text')
            >>> # Later, accept all changes from round1
            >>> count = doc.accept_edit_group('round1')
            >>> print(f"Accepted {count} changes")
        """
        self._ensure_edit_groups()
        ids = self._edit_groups.get_group_ids(group_name)
        count = 0

        # Accept in reverse order to handle nested elements properly
        for change_id in reversed(ids):
            try:
                self._change_mgmt.accept_change(change_id)
                count += 1
            except ValueError:
                # Change may have already been accepted or no longer exists
                pass

        return count
