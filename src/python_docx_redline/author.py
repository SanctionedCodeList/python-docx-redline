"""
Author identity classes for MS365/Office365 integration.

This module provides AuthorIdentity for linking tracked changes to real
MS365 users with full identity information including email, GUID, and provider.
"""

from dataclasses import dataclass


@dataclass
class AuthorIdentity:
    """MS365/Office365 author identity for tracked changes.

    This class stores full identity information that links tracked changes
    to real MS365 users. When edits are made with an AuthorIdentity, they
    appear in Word's revision history with the user's full profile information.

    Attributes:
        author: Display name (e.g., "Hancock, Parker" or "Parker Hancock")
        email: User's email address (e.g., "parker.hancock@company.com")
        provider_id: Identity provider ID (e.g., "AD" for Active Directory)
        guid: Unique identifier for the user (UUID format)

    Example:
        >>> identity = AuthorIdentity(
        ...     author="Hancock, Parker",
        ...     email="parker.hancock@bakerbotts.com",
        ...     provider_id="AD",
        ...     guid="c5c513d2-1f51-4d69-ae91-17e5787f9bfc"
        ... )
        >>> doc = Document("contract.docx", author=identity)
        >>> doc.insert_tracked(" (amended)", after="Section 1")
        >>> # Changes now appear as "Hancock, Parker" in Word with full profile

    Note:
        To find existing identity info from a document:
        1. Unpack the .docx file
        2. Inspect word/people.xml
        3. Look for w:author/@w15:providerId and w:author/@w15:userId attributes
    """

    author: str
    email: str
    provider_id: str = "AD"  # Active Directory is most common
    guid: str = ""  # Can be empty, Word will still link by email

    def __post_init__(self) -> None:
        """Validate author identity fields."""
        if not self.author:
            raise ValueError("Author name cannot be empty")
        if not self.email:
            raise ValueError("Email cannot be empty")
        if "@" not in self.email:
            raise ValueError(f"Invalid email format: {self.email}")

    @property
    def display_name(self) -> str:
        """Get the display name for the author."""
        return self.author

    def __str__(self) -> str:
        """String representation showing author and email."""
        return f"{self.author} <{self.email}>"

    def __repr__(self) -> str:
        """Detailed representation for debugging."""
        return (
            f"AuthorIdentity(author={self.author!r}, email={self.email!r}, "
            f"provider_id={self.provider_id!r}, guid={self.guid!r})"
        )
