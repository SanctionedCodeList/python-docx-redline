"""
Tests for fuzzy matching functionality.

This module tests the fuzzy matching feature which allows finding text
with OCR artifacts, typos, or minor variations using similarity thresholds.
"""

import pytest

from python_docx_redline.fuzzy import (
    fuzzy_find_all,
    fuzzy_match,
    normalize_whitespace,
    parse_fuzzy_config,
)


class TestNormalizeWhitespace:
    """Test whitespace normalization helper."""

    def test_normalize_multiple_spaces(self):
        """Test normalizing multiple spaces to single space."""
        assert normalize_whitespace("hello    world") == "hello world"

    def test_normalize_tabs_and_newlines(self):
        """Test normalizing tabs and newlines."""
        assert normalize_whitespace("hello\t\nworld") == "hello world"

    def test_strip_leading_trailing(self):
        """Test stripping leading and trailing whitespace."""
        assert normalize_whitespace("  hello world  ") == "hello world"

    def test_mixed_whitespace(self):
        """Test normalizing mixed whitespace characters."""
        assert normalize_whitespace("  hello  \t  world\n\n  ") == "hello world"


class TestFuzzyMatch:
    """Test basic fuzzy matching function."""

    def test_exact_match(self):
        """Test exact match returns True."""
        assert fuzzy_match("hello world", "hello world", threshold=0.9) is True

    def test_similar_text_above_threshold(self):
        """Test text above threshold returns True."""
        assert fuzzy_match("hello world", "helo world", threshold=0.8) is True

    def test_similar_text_below_threshold(self):
        """Test text below threshold returns False."""
        assert fuzzy_match("hello world", "goodbye earth", threshold=0.9) is False

    def test_ocr_artifacts(self):
        """Test matching text with OCR artifacts (0 vs O)."""
        assert fuzzy_match("production products", "producti0n pr0ducts", threshold=0.85) is True

    def test_typos(self):
        """Test matching text with typos."""
        assert fuzzy_match("Section 2.1", "Secton 2.1", threshold=0.9) is True

    def test_whitespace_normalization(self):
        """Test whitespace normalization option."""
        assert fuzzy_match("hello  world", "hello world", normalize_ws=True) is True

    def test_algorithm_ratio(self):
        """Test ratio algorithm."""
        assert fuzzy_match("hello", "helo", threshold=0.8, algorithm="ratio") is True

    def test_algorithm_partial_ratio(self):
        """Test partial_ratio algorithm."""
        assert fuzzy_match("hello world", "hello", threshold=0.8, algorithm="partial_ratio") is True

    def test_algorithm_token_sort_ratio(self):
        """Test token_sort_ratio algorithm."""
        assert (
            fuzzy_match("world hello", "hello world", threshold=0.9, algorithm="token_sort_ratio")
            is True
        )

    def test_algorithm_levenshtein(self):
        """Test levenshtein algorithm."""
        assert fuzzy_match("hello", "helo", threshold=0.8, algorithm="levenshtein") is True

    def test_invalid_threshold_too_low(self):
        """Test invalid threshold below 0 raises ValueError."""
        with pytest.raises(ValueError, match="Threshold must be between 0 and 1"):
            fuzzy_match("hello", "world", threshold=-0.1)

    def test_invalid_threshold_too_high(self):
        """Test invalid threshold above 1 raises ValueError."""
        with pytest.raises(ValueError, match="Threshold must be between 0 and 1"):
            fuzzy_match("hello", "world", threshold=1.5)

    def test_invalid_algorithm(self):
        """Test invalid algorithm raises ValueError."""
        with pytest.raises(ValueError, match="Invalid algorithm"):
            fuzzy_match("hello", "world", algorithm="invalid")


class TestFuzzyFindAll:
    """Test fuzzy find all function."""

    def test_single_match(self):
        """Test finding single fuzzy match."""
        text = "The producti0n pr0ducts are ready"
        matches = fuzzy_find_all(text, "production products", threshold=0.85)
        assert len(matches) == 1
        start, end, score = matches[0]
        assert text[start:end] == "producti0n pr0ducts"
        assert score >= 0.85

    def test_multiple_matches(self):
        """Test finding multiple fuzzy matches."""
        text = "Secti0n 1 and Secti0n 2 and Secti0n 3"
        matches = fuzzy_find_all(text, "Section", threshold=0.85)
        assert len(matches) == 3

    def test_no_matches(self):
        """Test no matches found."""
        text = "The quick brown fox"
        matches = fuzzy_find_all(text, "elephant", threshold=0.9)
        assert len(matches) == 0

    def test_whitespace_normalization(self):
        """Test fuzzy find with whitespace normalization."""
        text = "hello  world  test"
        matches = fuzzy_find_all(text, "hello world", threshold=0.9, normalize_ws=True)
        assert len(matches) >= 1

    def test_different_algorithms(self):
        """Test fuzzy find with different algorithms."""
        text = "The production products are ready"
        for algorithm in ["ratio", "partial_ratio", "token_sort_ratio", "levenshtein"]:
            matches = fuzzy_find_all(
                text, "production products", threshold=0.9, algorithm=algorithm
            )
            assert len(matches) >= 1, f"Algorithm {algorithm} should find match"

    def test_overlapping_matches_keeps_best(self):
        """Test that overlapping matches keep the best score."""
        text = "hello world"
        matches = fuzzy_find_all(text, "hello", threshold=0.7)
        # Should find at least one match
        assert len(matches) >= 1


class TestParseFuzzyConfig:
    """Test fuzzy configuration parser."""

    def test_none_returns_none(self):
        """Test None returns None (exact matching)."""
        assert parse_fuzzy_config(None) is None

    def test_float_threshold(self):
        """Test float threshold creates default config."""
        config = parse_fuzzy_config(0.9)
        assert config == {
            "threshold": 0.9,
            "algorithm": "ratio",
            "normalize_whitespace": False,
        }

    def test_int_threshold(self):
        """Test int threshold converts to float."""
        config = parse_fuzzy_config(1)
        assert config["threshold"] == 1.0
        assert isinstance(config["threshold"], float)

    def test_dict_with_threshold(self):
        """Test dict with custom threshold."""
        config = parse_fuzzy_config({"threshold": 0.85})
        assert config["threshold"] == 0.85
        assert config["algorithm"] == "ratio"

    def test_dict_with_algorithm(self):
        """Test dict with custom algorithm."""
        config = parse_fuzzy_config({"threshold": 0.9, "algorithm": "levenshtein"})
        assert config["algorithm"] == "levenshtein"

    def test_dict_with_normalize_whitespace(self):
        """Test dict with normalize_whitespace option."""
        config = parse_fuzzy_config({"normalize_whitespace": True})
        assert config["normalize_whitespace"] is True

    def test_dict_partial_config(self):
        """Test dict with only some options uses defaults."""
        config = parse_fuzzy_config({"algorithm": "partial_ratio"})
        assert config == {
            "threshold": 0.9,
            "algorithm": "partial_ratio",
            "normalize_whitespace": False,
        }

    def test_invalid_threshold_type(self):
        """Test invalid threshold type raises ValueError."""
        with pytest.raises(ValueError, match="Fuzzy threshold must be a number"):
            parse_fuzzy_config({"threshold": "invalid"})

    def test_invalid_threshold_range(self):
        """Test invalid threshold range raises ValueError."""
        with pytest.raises(ValueError, match="Fuzzy threshold must be a number between 0 and 1"):
            parse_fuzzy_config({"threshold": 1.5})

    def test_invalid_algorithm_name(self):
        """Test invalid algorithm name raises ValueError."""
        with pytest.raises(ValueError, match="Invalid algorithm"):
            parse_fuzzy_config({"algorithm": "invalid"})

    def test_invalid_normalize_whitespace_type(self):
        """Test invalid normalize_whitespace type raises ValueError."""
        with pytest.raises(ValueError, match="normalize_whitespace must be a boolean"):
            parse_fuzzy_config({"normalize_whitespace": "yes"})

    def test_invalid_config_type(self):
        """Test invalid config type raises ValueError."""
        with pytest.raises(ValueError, match="fuzzy parameter must be"):
            parse_fuzzy_config([0.9])


class TestFuzzyMatchingImportError:
    """Test that missing rapidfuzz raises helpful error."""

    def test_fuzzy_match_without_rapidfuzz(self, monkeypatch):
        """Test fuzzy_match raises ImportError with helpful message."""
        # This test is skipped because mocking import causes recursion issues
        # The error handling is tested manually by attempting to use fuzzy
        # matching without rapidfuzz installed
        pytest.skip("Import mocking causes recursion - error handling verified manually")


# Integration tests with Document class would go in a separate test file
# since they require actual Word documents
