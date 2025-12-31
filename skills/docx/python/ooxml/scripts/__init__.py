"""OOXML scripts for packing, unpacking, and validating Office documents."""

from .pack import pack_document
from .unpack import unpack_document

__all__ = ["pack_document", "unpack_document"]
