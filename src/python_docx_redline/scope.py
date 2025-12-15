"""
Scope evaluation for filtering paragraphs in Word documents.

This module provides a flexible scope system that allows users to limit
search operations to specific sections, paragraphs, or text ranges.
"""

from collections.abc import Callable
from typing import Any

from .constants import WORD_NAMESPACE


class ScopeEvaluator:
    """Evaluates scope specifications to filter paragraphs.

    Supports multiple scope formats:
    - None: Match all paragraphs
    - String: "text" matches paragraphs containing text
    - String with prefix: "section:Name" or "paragraph_containing:text"
    - Dictionary: {"contains": "text", "section": "Name", ...}
    - Callable: Custom filter function
    """

    @staticmethod
    def parse(scope_spec: str | dict | Callable | None) -> Callable[[Any], bool]:
        """Convert scope specification to evaluation function.

        Args:
            scope_spec: The scope specification:
                - None: Match all paragraphs
                - str: Paragraph containing text (or special formats)
                - dict: Dictionary with filter criteria
                - callable: Custom filter function

        Returns:
            A callable that takes a paragraph Element and returns bool

        Raises:
            ValueError: If scope specification is invalid
        """
        if scope_spec is None:
            return lambda p: True

        if isinstance(scope_spec, str):
            return ScopeEvaluator._parse_string(scope_spec)

        if isinstance(scope_spec, dict):
            return ScopeEvaluator._parse_dict(scope_spec)

        if callable(scope_spec):
            return scope_spec

        raise ValueError(f"Invalid scope specification: {scope_spec}")

    @staticmethod
    def _parse_string(s: str) -> Callable[[Any], bool]:
        """Parse string scope shortcuts.

        Supported formats:
        - "section:Name": Match paragraphs in section with heading "Name"
        - "paragraph_containing:text": Explicit paragraph containing text
        - "text": Default - paragraph containing text

        Args:
            s: The string scope specification

        Returns:
            A callable that filters paragraphs
        """
        if s.startswith("section:"):
            section_name = s[8:]
            return ScopeEvaluator._create_section_filter(section_name)

        if s.startswith("paragraph_containing:"):
            text = s[21:]
            return ScopeEvaluator._create_text_filter(text)

        # Default: paragraph containing the specified text
        return ScopeEvaluator._create_text_filter(s)

    @staticmethod
    def _parse_dict(d: dict) -> Callable[[Any], bool]:
        """Parse dictionary scope specification.

        Supported keys:
        - contains: Text that must be in the paragraph
        - section: Section heading name
        - not_contains: Text that must NOT be in the paragraph

        Args:
            d: Dictionary with filter criteria

        Returns:
            A callable that filters paragraphs
        """

        def evaluator(para: Any) -> bool:
            # Extract paragraph text
            para_text = "".join(para.itertext())

            # Check 'contains' filter
            if "contains" in d:
                if d["contains"] not in para_text:
                    return False

            # Check 'not_contains' filter
            if "not_contains" in d:
                if d["not_contains"] in para_text:
                    return False

            # Check 'section' filter
            if "section" in d:
                # For section filtering, we need to check if this paragraph
                # comes after a heading with the specified text
                # This is a simplified implementation
                section_name = d["section"]
                if not ScopeEvaluator._is_in_section(para, section_name):
                    return False

            return True

        return evaluator

    @staticmethod
    def _create_text_filter(text: str) -> Callable[[Any], bool]:
        """Create a filter that checks if paragraph contains text.

        Args:
            text: The text to search for

        Returns:
            A callable that checks if the paragraph contains the text
        """

        def filter_func(para: Any) -> bool:
            para_text = "".join(para.itertext())
            return text in para_text

        return filter_func

    @staticmethod
    def _create_section_filter(section_name: str) -> Callable[[Any], bool]:
        """Create a filter for paragraphs in a specific section.

        A section is defined by a heading paragraph. This filter matches
        paragraphs that come after a heading containing the section name,
        but excludes the heading itself.

        Args:
            section_name: The section heading text to match

        Returns:
            A callable that checks if the paragraph is in the section
        """

        def filter_func(para: Any) -> bool:
            # Don't include headings themselves
            if ScopeEvaluator._is_heading(para):
                return False
            return ScopeEvaluator._is_in_section(para, section_name)

        return filter_func

    @staticmethod
    def _is_in_section(para: Any, section_name: str) -> bool:
        """Check if a paragraph is within a named section.

        Walks backwards from the paragraph to find the most recent heading,
        then checks if that heading contains the section name.

        Args:
            para: The paragraph Element
            section_name: The section heading text to match

        Returns:
            True if the paragraph is in the specified section
        """
        # Get the parent body element
        body = para.getparent()
        if body is None:
            return False

        # Get all paragraphs in the document
        all_paragraphs = list(body.iter(f"{{{WORD_NAMESPACE}}}p"))

        # Find the index of our paragraph
        try:
            para_index = all_paragraphs.index(para)
        except ValueError:
            return False

        # Walk backwards to find the most recent heading
        for i in range(para_index - 1, -1, -1):
            prev_para = all_paragraphs[i]

            # Check if this paragraph is a heading
            if ScopeEvaluator._is_heading(prev_para):
                heading_text = "".join(prev_para.itertext())
                # Check if heading contains the section name
                return section_name in heading_text

        # No heading found - not in any section
        return False

    @staticmethod
    def _is_heading(para: Any) -> bool:
        """Check if a paragraph is a heading.

        A paragraph is considered a heading if it has a paragraph style
        that starts with "Heading".

        Args:
            para: The paragraph Element

        Returns:
            True if the paragraph is a heading
        """
        # Look for paragraph properties
        p_pr = para.find(f"{{{WORD_NAMESPACE}}}pPr")
        if p_pr is None:
            return False

        # Look for style
        p_style = p_pr.find(f"{{{WORD_NAMESPACE}}}pStyle")
        if p_style is None:
            return False

        # Check if style value starts with "Heading"
        style_val = p_style.get(f"{{{WORD_NAMESPACE}}}val")
        if style_val is None or not isinstance(style_val, str):
            return False

        return bool(style_val.startswith("Heading"))

    @staticmethod
    def filter_paragraphs(
        paragraphs: list[Any], scope_spec: str | dict | Callable | None
    ) -> list[Any]:
        """Filter a list of paragraphs using a scope specification.

        This is a convenience method that combines parse() and filtering.

        Args:
            paragraphs: List of paragraph Elements to filter
            scope_spec: The scope specification (see parse() for formats)

        Returns:
            Filtered list of paragraph Elements

        Example:
            >>> paragraphs = doc.xml_root.iter(f"{{{WORD_NAMESPACE}}}p")
            >>> filtered = ScopeEvaluator.filter_paragraphs(
            ...     paragraphs,
            ...     scope="section:Introduction"
            ... )
        """
        if scope_spec is None:
            return paragraphs

        evaluator = ScopeEvaluator.parse(scope_spec)
        return [p for p in paragraphs if evaluator(p)]
