"""
Fuzzy text matching support for finding text with OCR artifacts or variations.

This module provides fuzzy matching capabilities using the rapidfuzz library,
allowing documents with OCR artifacts, typos, or minor variations to be matched
using similarity thresholds.

Example:
    >>> from python_docx_redline.fuzzy import fuzzy_match
    >>> # Simple threshold-based matching
    >>> fuzzy_match("hello world", "helo world", threshold=0.9)
    True
    >>> # Custom algorithm
    >>> fuzzy_match("hello", "helo", threshold=0.8, algorithm="levenshtein")
    True
"""

import re
from typing import Any


def normalize_whitespace(text: str) -> str:
    """Normalize whitespace in text for matching.

    Replaces multiple whitespace characters with single spaces and strips
    leading/trailing whitespace.

    Args:
        text: The text to normalize

    Returns:
        Normalized text with single spaces

    Example:
        >>> normalize_whitespace("hello    world\\n\\ttest")
        'hello world test'
    """
    return re.sub(r"\s+", " ", text).strip()


def fuzzy_match(
    text: str,
    pattern: str,
    threshold: float = 0.9,
    algorithm: str = "ratio",
    normalize_ws: bool = False,
) -> bool:
    """Check if text matches pattern with fuzzy matching.

    Uses rapidfuzz library to perform fuzzy string matching with configurable
    similarity thresholds and algorithms.

    Args:
        text: The text to check
        pattern: The pattern to match against
        threshold: Similarity threshold (0.0 to 1.0), default 0.9 (90% similar)
        algorithm: Matching algorithm to use:
            - 'ratio': Overall similarity (default, good for general use)
            - 'partial_ratio': Best partial match (good for substring matching)
            - 'token_sort_ratio': Order-independent word matching
            - 'levenshtein': Edit distance based (good for typos)
        normalize_ws: Whether to normalize whitespace before matching

    Returns:
        True if similarity meets or exceeds threshold, False otherwise

    Raises:
        ImportError: If rapidfuzz is not installed
        ValueError: If threshold is not between 0 and 1, or algorithm is invalid

    Example:
        >>> # OCR artifacts
        >>> fuzzy_match("production products", "producti0n pr0ducts", threshold=0.85)
        True
        >>> # Typos
        >>> fuzzy_match("Section 2.1", "Secton 2.1", threshold=0.9)
        True
        >>> # Whitespace normalization
        >>> fuzzy_match("hello  world", "hello world", normalize_ws=True)
        True
    """
    try:
        from rapidfuzz import fuzz
    except ImportError as e:
        raise ImportError(
            "rapidfuzz is required for fuzzy matching. "
            'Install it with: pip install "python-docx-redline[fuzzy]"'
        ) from e

    # Validate threshold
    if not 0 <= threshold <= 1:
        raise ValueError(f"Threshold must be between 0 and 1, got {threshold}")

    # Normalize whitespace if requested
    if normalize_ws:
        text = normalize_whitespace(text)
        pattern = normalize_whitespace(pattern)

    # Select algorithm and compute similarity
    if algorithm == "ratio":
        similarity = fuzz.ratio(text, pattern) / 100.0
    elif algorithm == "partial_ratio":
        similarity = fuzz.partial_ratio(text, pattern) / 100.0
    elif algorithm == "token_sort_ratio":
        similarity = fuzz.token_sort_ratio(text, pattern) / 100.0
    elif algorithm == "levenshtein":
        # Use ratio for levenshtein (normalized edit distance)
        similarity = fuzz.ratio(text, pattern) / 100.0
    else:
        raise ValueError(
            f"Invalid algorithm '{algorithm}'. "
            "Must be one of: ratio, partial_ratio, token_sort_ratio, levenshtein"
        )

    return similarity >= threshold


def fuzzy_find_all(
    text: str,
    pattern: str,
    threshold: float = 0.9,
    algorithm: str = "ratio",
    normalize_ws: bool = False,
) -> list[tuple[int, int, float]]:
    """Find all fuzzy matches of pattern in text using sliding window.

    This function uses a sliding window approach to find all occurrences of
    pattern in text that meet the fuzzy matching threshold. It returns the
    start/end positions and similarity scores of all matches.

    Args:
        text: The text to search in
        pattern: The pattern to search for
        threshold: Similarity threshold (0.0 to 1.0), default 0.9
        algorithm: Matching algorithm (ratio, partial_ratio, token_sort_ratio, levenshtein)
        normalize_ws: Whether to normalize whitespace before matching

    Returns:
        List of tuples (start_pos, end_pos, similarity_score) for each match

    Raises:
        ImportError: If rapidfuzz is not installed
        ValueError: If threshold is not between 0 and 1, or algorithm is invalid

    Example:
        >>> text = "The producti0n products are ready"
        >>> matches = fuzzy_find_all(text, "production products", threshold=0.85)
        >>> [(start, end, score) for start, end, score in matches]
        [(4, 23, 0.87)]
    """
    try:
        from rapidfuzz import fuzz
    except ImportError as e:
        raise ImportError(
            "rapidfuzz is required for fuzzy matching. "
            'Install it with: pip install "python-docx-redline[fuzzy]"'
        ) from e

    # Validate threshold
    if not 0 <= threshold <= 1:
        raise ValueError(f"Threshold must be between 0 and 1, got {threshold}")

    # Validate algorithm
    valid_algorithms = ["ratio", "partial_ratio", "token_sort_ratio", "levenshtein"]
    if algorithm not in valid_algorithms:
        raise ValueError(
            f"Invalid algorithm '{algorithm}'. Must be one of: {', '.join(valid_algorithms)}"
        )

    # Normalize whitespace if requested
    search_text = normalize_whitespace(text) if normalize_ws else text
    search_pattern = normalize_whitespace(pattern) if normalize_ws else pattern

    matches: list[tuple[int, int, float]] = []
    pattern_len = len(search_pattern)

    # Use sliding window to find all matches
    # We'll check windows of varying sizes around the pattern length
    min_len = max(1, int(pattern_len * 0.7))  # Allow for insertions/deletions
    max_len = int(pattern_len * 1.3)

    checked_positions: set[tuple[int, int]] = set()

    for window_len in range(min_len, max_len + 1):
        for start in range(len(search_text) - window_len + 1):
            end = start + window_len

            # Skip if we've already checked this position
            if (start, end) in checked_positions:
                continue
            checked_positions.add((start, end))

            window = search_text[start:end]

            # Compute similarity using selected algorithm
            if algorithm == "ratio":
                similarity = fuzz.ratio(window, search_pattern) / 100.0
            elif algorithm == "partial_ratio":
                similarity = fuzz.partial_ratio(window, search_pattern) / 100.0
            elif algorithm == "token_sort_ratio":
                similarity = fuzz.token_sort_ratio(window, search_pattern) / 100.0
            elif algorithm == "levenshtein":
                similarity = fuzz.ratio(window, search_pattern) / 100.0

            if similarity >= threshold:
                # Check for overlapping matches and keep the best one
                overlaps = [
                    (i, m_start, m_end, m_score)
                    for i, (m_start, m_end, m_score) in enumerate(matches)
                    if not (end <= m_start or start >= m_end)
                ]

                if overlaps:
                    # If this match is better than overlapping ones, replace them
                    best_overlap_score = max(score for _, _, _, score in overlaps)
                    if similarity > best_overlap_score:
                        # Remove overlapping matches
                        for idx, _, _, _ in sorted(overlaps, reverse=True):
                            del matches[idx]
                        matches.append((start, end, similarity))
                else:
                    matches.append((start, end, similarity))

    # Sort by position
    return sorted(matches, key=lambda x: x[0])


def parse_fuzzy_config(fuzzy: float | dict[str, Any] | None) -> dict[str, Any] | None:
    """Parse fuzzy matching configuration into a standardized dict.

    Accepts either:
    - None: Exact matching (no fuzzy)
    - float: Simple threshold (e.g., 0.9 for 90% similarity)
    - dict: Full config with threshold, algorithm, normalize_whitespace

    Args:
        fuzzy: Fuzzy configuration (None, float, or dict)

    Returns:
        Parsed configuration dict or None for exact matching

    Raises:
        ValueError: If configuration is invalid

    Example:
        >>> parse_fuzzy_config(0.9)
        {'threshold': 0.9, 'algorithm': 'ratio', 'normalize_whitespace': False}
        >>> parse_fuzzy_config({'threshold': 0.85, 'algorithm': 'levenshtein'})
        {'threshold': 0.85, 'algorithm': 'levenshtein', 'normalize_whitespace': False}
        >>> parse_fuzzy_config({'normalize_whitespace': True})
        {'threshold': 0.9, 'algorithm': 'ratio', 'normalize_whitespace': True}
    """
    if fuzzy is None:
        return None

    # Default configuration
    config = {
        "threshold": 0.9,
        "algorithm": "ratio",
        "normalize_whitespace": False,
    }

    if isinstance(fuzzy, int | float):
        # Simple threshold configuration
        if not 0 <= fuzzy <= 1:
            raise ValueError(f"Fuzzy threshold must be between 0 and 1, got {fuzzy}")
        config["threshold"] = float(fuzzy)
    elif isinstance(fuzzy, dict):
        # Full configuration
        if "threshold" in fuzzy:
            threshold = fuzzy["threshold"]
            if not isinstance(threshold, int | float) or not (0 <= threshold <= 1):
                raise ValueError(
                    f"Fuzzy threshold must be a number between 0 and 1, got {threshold}"
                )
            config["threshold"] = float(threshold)

        if "algorithm" in fuzzy:
            algorithm = fuzzy["algorithm"]
            valid_algorithms = ["ratio", "partial_ratio", "token_sort_ratio", "levenshtein"]
            if algorithm not in valid_algorithms:
                raise ValueError(
                    f"Invalid algorithm '{algorithm}'. "
                    f"Must be one of: {', '.join(valid_algorithms)}"
                )
            config["algorithm"] = algorithm

        if "normalize_whitespace" in fuzzy:
            normalize = fuzzy["normalize_whitespace"]
            if not isinstance(normalize, bool):
                raise ValueError(
                    f"normalize_whitespace must be a boolean, got {type(normalize).__name__}"
                )
            config["normalize_whitespace"] = normalize
    else:
        raise ValueError(
            f"fuzzy parameter must be None, a float, or a dict, got {type(fuzzy).__name__}"
        )

    return config
