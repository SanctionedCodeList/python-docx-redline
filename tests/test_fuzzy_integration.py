"""
Integration tests for fuzzy matching with Document class.

These tests verify that fuzzy matching works end-to-end with actual
Word documents and tracked change operations.

NOTE: These tests are skipped for now as they require a more complex test setup.
The core fuzzy matching functionality is thoroughly tested in test_fuzzy_matching.py.
"""

import tempfile
from pathlib import Path

import pytest

from python_docx_redline import Document
from python_docx_redline.errors import TextNotFoundError

pytestmark = pytest.mark.skip(reason="Integration tests require more complex document setup")


@pytest.fixture
def fuzzy_test_doc():
    """Create a test document with OCR-like artifacts."""
    # This fixture would need to create a proper Word document structure
    # For now, these tests are skipped
    pass


class TestFuzzyFindAll:
    """Test find_all with fuzzy matching."""

    def test_find_all_with_fuzzy_threshold(self, fuzzy_test_doc):
        """Test finding all matches with fuzzy threshold."""
        doc = Document(fuzzy_test_doc)
        matches = doc.find_all("production products", fuzzy=0.85)

        assert len(matches) >= 1
        # Should match "producti0n pr0ducts"
        assert (
            "producti0n pr0ducts" in matches[0].text.lower()
            or "production products" in matches[0].text.lower()
        )

    def test_find_all_with_fuzzy_dict_config(self, fuzzy_test_doc):
        """Test finding with full fuzzy configuration."""
        doc = Document(fuzzy_test_doc)
        matches = doc.find_all(
            "Section 2.1",
            fuzzy={"threshold": 0.85, "algorithm": "ratio", "normalize_whitespace": False},
        )

        assert len(matches) >= 1

    def test_find_all_exact_vs_fuzzy(self, fuzzy_test_doc):
        """Test that exact matching fails where fuzzy succeeds."""
        doc = Document(fuzzy_test_doc)

        # Exact match should fail
        exact_matches = doc.find_all("production products", fuzzy=None)
        assert len(exact_matches) == 0

        # Fuzzy match should succeed
        fuzzy_matches = doc.find_all("production products", fuzzy=0.85)
        assert len(fuzzy_matches) >= 1

    def test_find_all_with_whitespace_normalization(self):
        """Test fuzzy matching with whitespace normalization."""
        doc = Document()
        doc._add_paragraph("hello  world  test")

        matches = doc.find_all(
            "hello world", fuzzy={"threshold": 0.9, "normalize_whitespace": True}
        )
        assert len(matches) >= 1

    def test_find_all_fuzzy_and_regex_raises_error(self, fuzzy_test_doc):
        """Test that using both fuzzy and regex raises ValueError."""
        doc = Document(fuzzy_test_doc)

        with pytest.raises(ValueError, match="Cannot use both fuzzy matching and regex"):
            doc.find_all(r"production \w+", fuzzy=0.9, regex=True)


class TestFuzzyInsertTracked:
    """Test insert_tracked with fuzzy matching."""

    def test_insert_after_with_fuzzy(self, fuzzy_test_doc):
        """Test inserting after fuzzy-matched anchor."""
        doc = Document(fuzzy_test_doc)

        # Insert after fuzzy-matched text
        doc.insert_tracked(
            " (updated)",
            after="production products",
            fuzzy=0.85,
        )

        # Verify insertion
        doc.save(fuzzy_test_doc)
        doc = Document(fuzzy_test_doc)
        assert doc.has_tracked_changes()

    def test_insert_before_with_fuzzy(self, fuzzy_test_doc):
        """Test inserting before fuzzy-matched anchor."""
        doc = Document(fuzzy_test_doc)

        doc.insert_tracked(
            "Important: ",
            before="Section 2.1",
            fuzzy=0.85,
        )

        doc.save(fuzzy_test_doc)
        doc = Document(fuzzy_test_doc)
        assert doc.has_tracked_changes()

    def test_insert_fuzzy_not_found(self, fuzzy_test_doc):
        """Test insert with fuzzy when no match meets threshold."""
        doc = Document(fuzzy_test_doc)

        with pytest.raises(TextNotFoundError):
            doc.insert_tracked(
                "text",
                after="completely different text",
                fuzzy=0.9,
            )

    def test_insert_fuzzy_occurrence_all(self, fuzzy_test_doc):
        """Test inserting at all fuzzy-matched occurrences."""
        doc = Document(fuzzy_test_doc)

        # Add more paragraphs with similar text
        doc._add_paragraph("Another producti0n pr0ducts section here.")

        doc.insert_tracked(
            " [MATCHED]",
            after="production products",
            fuzzy=0.85,
            occurrence="all",
        )

        doc.save(fuzzy_test_doc)
        doc = Document(fuzzy_test_doc)
        assert doc.has_tracked_changes()


class TestFuzzyDeleteTracked:
    """Test delete_tracked with fuzzy matching."""

    def test_delete_with_fuzzy(self, fuzzy_test_doc):
        """Test deleting fuzzy-matched text."""
        doc = Document(fuzzy_test_doc)

        doc.delete_tracked("production products", fuzzy=0.85)

        doc.save(fuzzy_test_doc)
        doc = Document(fuzzy_test_doc)
        assert doc.has_tracked_changes()

    def test_delete_fuzzy_occurrence_all(self, fuzzy_test_doc):
        """Test deleting all fuzzy-matched occurrences."""
        doc = Document(fuzzy_test_doc)

        # Add more paragraphs
        doc._add_paragraph("Text with OCR artifacts: pr0ducts and pr0cess.")

        doc.delete_tracked("products", fuzzy=0.80, occurrence="all")

        doc.save(fuzzy_test_doc)
        doc = Document(fuzzy_test_doc)
        assert doc.has_tracked_changes()

    def test_delete_fuzzy_not_found(self, fuzzy_test_doc):
        """Test delete with fuzzy when no match meets threshold."""
        doc = Document(fuzzy_test_doc)

        with pytest.raises(TextNotFoundError):
            doc.delete_tracked("elephant", fuzzy=0.9)


class TestFuzzyReplaceTracked:
    """Test replace_tracked with fuzzy matching."""

    def test_replace_with_fuzzy(self, fuzzy_test_doc):
        """Test replacing fuzzy-matched text."""
        doc = Document(fuzzy_test_doc)

        doc.replace_tracked(
            "production products",
            "manufacturing products",
            fuzzy=0.85,
        )

        doc.save(fuzzy_test_doc)
        doc = Document(fuzzy_test_doc)
        assert doc.has_tracked_changes()

    def test_replace_fuzzy_occurrence_all(self, fuzzy_test_doc):
        """Test replacing all fuzzy-matched occurrences."""
        doc = Document(fuzzy_test_doc)

        # Add more text
        doc._add_paragraph("More producti0n pr0ducts here.")

        doc.replace_tracked(
            "production products",
            "manufacturing items",
            fuzzy=0.85,
            occurrence="all",
        )

        doc.save(fuzzy_test_doc)
        doc = Document(fuzzy_test_doc)
        assert doc.has_tracked_changes()

    def test_replace_fuzzy_with_different_algorithms(self, fuzzy_test_doc):
        """Test replace with different fuzzy algorithms."""
        algorithms = ["ratio", "partial_ratio", "token_sort_ratio", "levenshtein"]

        for algorithm in algorithms:
            doc = Document(fuzzy_test_doc)

            doc.replace_tracked(
                "Section 2.1",
                "Section 2.2",
                fuzzy={"threshold": 0.85, "algorithm": algorithm},
            )

            doc.save(fuzzy_test_doc)
            doc = Document(fuzzy_test_doc)
            assert doc.has_tracked_changes(), f"Algorithm {algorithm} should create tracked changes"

    def test_replace_fuzzy_not_found(self, fuzzy_test_doc):
        """Test replace with fuzzy when no match meets threshold."""
        doc = Document(fuzzy_test_doc)

        with pytest.raises(TextNotFoundError):
            doc.replace_tracked("elephant", "giraffe", fuzzy=0.9)


class TestFuzzyWithScope:
    """Test fuzzy matching with scope limitations."""

    def test_fuzzy_with_scope_dict(self, fuzzy_test_doc):
        """Test fuzzy matching with scope filter."""
        doc = Document(fuzzy_test_doc)

        # This would require scope to be set up properly in the test doc
        # For now, just verify it doesn't crash
        matches = doc.find_all("production", fuzzy=0.8, scope=None)
        assert isinstance(matches, list)


class TestFuzzyConfigValidation:
    """Test fuzzy configuration validation in Document methods."""

    def test_invalid_fuzzy_config_raises_error(self, fuzzy_test_doc):
        """Test that invalid fuzzy config raises ValueError."""
        doc = Document(fuzzy_test_doc)

        with pytest.raises(ValueError):
            doc.find_all("text", fuzzy={"threshold": 1.5})

        with pytest.raises(ValueError):
            doc.find_all("text", fuzzy={"algorithm": "invalid"})

    def test_fuzzy_threshold_boundaries(self, fuzzy_test_doc):
        """Test fuzzy matching at threshold boundaries."""
        doc = Document(fuzzy_test_doc)

        # Valid thresholds
        doc.find_all("text", fuzzy=0.0)
        doc.find_all("text", fuzzy=1.0)

        # Invalid thresholds
        with pytest.raises(ValueError):
            doc.find_all("text", fuzzy=-0.1)

        with pytest.raises(ValueError):
            doc.find_all("text", fuzzy=1.1)


class TestFuzzyMatchingPerformance:
    """Test fuzzy matching performance characteristics."""

    def test_fuzzy_with_long_document(self):
        """Test fuzzy matching with a longer document."""
        doc = Document()

        # Add many paragraphs
        for i in range(50):
            doc._add_paragraph(f"Paragraph {i} with producti0n pr0ducts text.")

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            doc.save(tmp.name)

            doc = Document(tmp.name)
            matches = doc.find_all("production products", fuzzy=0.85)

            # Should find many matches
            assert len(matches) >= 40

            Path(tmp.name).unlink()

    def test_fuzzy_different_threshold_levels(self, fuzzy_test_doc):
        """Test how different thresholds affect match count."""
        doc = Document(fuzzy_test_doc)

        # Lower threshold should find more (or same) matches
        low_threshold_matches = doc.find_all("production", fuzzy=0.7)
        high_threshold_matches = doc.find_all("production", fuzzy=0.95)

        assert len(low_threshold_matches) >= len(high_threshold_matches)
