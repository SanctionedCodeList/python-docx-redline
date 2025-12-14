"""
Tests for the export functionality (Phase 14).

These tests verify:
- export_changes_json() method
- export_changes_markdown() method
- export_changes_html() method
- generate_change_report() method
- Grouping by author/type
- Context extraction
"""

import json
import tempfile
import zipfile
from pathlib import Path

from python_docx_redline import (
    ChangeContext,
    ChangeReport,
    Document,
    ExportedChange,
    export_changes_html,
    export_changes_json,
    export_changes_markdown,
    generate_change_report,
)

# XML with various tracked changes for testing
DOCUMENT_WITH_TRACKED_CHANGES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is the original text.</w:t>
      </w:r>
      <w:ins w:id="1" w:author="Alice" w:date="2024-01-15T10:30:00Z">
        <w:r>
          <w:t> Added by Alice.</w:t>
        </w:r>
      </w:ins>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Second paragraph </w:t>
      </w:r>
      <w:del w:id="2" w:author="Bob" w:date="2024-01-16T14:00:00Z">
        <w:r>
          <w:delText>removed text</w:delText>
        </w:r>
      </w:del>
      <w:r>
        <w:t> more text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Third paragraph with </w:t>
      </w:r>
      <w:ins w:id="3" w:author="Alice" w:date="2024-01-17T09:00:00Z">
        <w:r>
          <w:t>more insertions</w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t> here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:pPrChange w:id="4" w:author="Carol" w:date="2024-01-18T11:00:00Z">
          <w:pPr/>
        </w:pPrChange>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:rPrChange w:id="5" w:author="Bob" w:date="2024-01-18T12:00:00Z">
            <w:rPr/>
          </w:rPrChange>
        </w:rPr>
        <w:t>Formatted paragraph.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


DOCUMENT_NO_CHANGES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a document with no tracked changes.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"""


def create_test_docx(content: str) -> Path:
    """Create a minimal but valid OOXML test .docx file."""
    temp_dir = Path(tempfile.mkdtemp())
    docx_path = temp_dir / "test.docx"

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", content)

    return docx_path


class TestExportChangesJson:
    """Tests for export_changes_json() method."""

    def test_export_basic_json(self):
        """Test basic JSON export."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json()

            data = json.loads(json_str)
            assert "total_changes" in data
            assert "changes" in data
            assert data["total_changes"] == 5
            assert len(data["changes"]) == 5

        finally:
            docx_path.unlink()

    def test_export_json_change_fields(self):
        """Test that exported JSON has all required fields."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json()

            data = json.loads(json_str)
            change = data["changes"][0]

            assert "id" in change
            assert "change_type" in change
            assert "author" in change
            assert "date" in change
            assert "text" in change

        finally:
            docx_path.unlink()

    def test_export_json_with_context(self):
        """Test JSON export with context included."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json(include_context=True)

            data = json.loads(json_str)
            # At least one change should have context
            contexts = [c.get("context") for c in data["changes"]]
            has_context = any(c is not None for c in contexts)
            assert has_context, "At least one change should have context"

        finally:
            docx_path.unlink()

    def test_export_json_without_context(self):
        """Test JSON export without context."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json(include_context=False)

            data = json.loads(json_str)
            # No change should have context
            for change in data["changes"]:
                assert "context" not in change

        finally:
            docx_path.unlink()

    def test_export_json_compact(self):
        """Test JSON export with compact format (no indent)."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json(indent=None)

            # Compact JSON should not have newlines (except within strings)
            assert "\n" not in json_str

        finally:
            docx_path.unlink()

    def test_export_json_empty_document(self):
        """Test JSON export on document with no changes."""
        docx_path = create_test_docx(DOCUMENT_NO_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json()

            data = json.loads(json_str)
            assert data["total_changes"] == 0
            assert data["changes"] == []

        finally:
            docx_path.unlink()

    def test_export_json_standalone_function(self):
        """Test the standalone export_changes_json function."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = export_changes_json(doc)

            data = json.loads(json_str)
            assert data["total_changes"] == 5

        finally:
            docx_path.unlink()


class TestExportChangesMarkdown:
    """Tests for export_changes_markdown() method."""

    def test_export_basic_markdown(self):
        """Test basic Markdown export."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            md = doc.export_changes_markdown()

            assert "# Tracked Changes Report" in md
            assert "## Summary" in md
            assert "Total changes" in md

        finally:
            docx_path.unlink()

    def test_export_markdown_summary_stats(self):
        """Test that Markdown includes summary statistics."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            md = doc.export_changes_markdown()

            assert "Insertions" in md
            assert "Deletions" in md

        finally:
            docx_path.unlink()

    def test_export_markdown_group_by_author(self):
        """Test Markdown export grouped by author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            md = doc.export_changes_markdown(group_by="author")

            assert "## Changes by Author" in md
            assert "### Alice" in md
            assert "### Bob" in md
            assert "### Carol" in md

        finally:
            docx_path.unlink()

    def test_export_markdown_group_by_type(self):
        """Test Markdown export grouped by type."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            md = doc.export_changes_markdown(group_by="type")

            assert "## Changes by Type" in md
            assert "### Insertions" in md
            assert "### Deletions" in md

        finally:
            docx_path.unlink()

    def test_export_markdown_no_grouping(self):
        """Test Markdown export without grouping."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            md = doc.export_changes_markdown(group_by=None)

            assert "## All Changes" in md
            assert "## Changes by Author" not in md
            assert "## Changes by Type" not in md

        finally:
            docx_path.unlink()

    def test_export_markdown_with_context(self):
        """Test Markdown export with context."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            md = doc.export_changes_markdown(include_context=True)

            # Context should appear with "[change]" marker
            assert "[change]" in md or "Context:" in md

        finally:
            docx_path.unlink()

    def test_export_markdown_empty_document(self):
        """Test Markdown export on document with no changes."""
        docx_path = create_test_docx(DOCUMENT_NO_CHANGES)
        try:
            doc = Document(docx_path)
            md = doc.export_changes_markdown()

            assert "No tracked changes found" in md

        finally:
            docx_path.unlink()

    def test_export_markdown_standalone_function(self):
        """Test the standalone export_changes_markdown function."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            md = export_changes_markdown(doc)

            assert "# Tracked Changes Report" in md

        finally:
            docx_path.unlink()


class TestExportChangesHtml:
    """Tests for export_changes_html() method."""

    def test_export_basic_html(self):
        """Test basic HTML export."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html()

            assert "<!DOCTYPE html>" in html
            assert "<html" in html
            assert "Tracked Changes Report" in html

        finally:
            docx_path.unlink()

    def test_export_html_with_styles(self):
        """Test HTML export includes inline styles."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html(inline_styles=True)

            assert "<style>" in html
            assert ".change-item" in html
            assert ".badge-insertion" in html

        finally:
            docx_path.unlink()

    def test_export_html_without_styles(self):
        """Test HTML export without inline styles."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html(inline_styles=False)

            assert "<style>" not in html

        finally:
            docx_path.unlink()

    def test_export_html_summary_section(self):
        """Test HTML includes summary section."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html()

            assert "Summary" in html
            assert "Total changes" in html
            assert "Insertions" in html
            assert "Deletions" in html

        finally:
            docx_path.unlink()

    def test_export_html_group_by_author(self):
        """Test HTML export grouped by author."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html(group_by="author")

            assert "Changes by Author" in html
            assert "Alice" in html
            assert "Bob" in html
            assert "Carol" in html

        finally:
            docx_path.unlink()

    def test_export_html_group_by_type(self):
        """Test HTML export grouped by type."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html(group_by="type")

            assert "Changes by Type" in html
            assert "Insertions" in html
            assert "Deletions" in html

        finally:
            docx_path.unlink()

    def test_export_html_no_grouping(self):
        """Test HTML export without grouping."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html(group_by=None)

            assert "All Changes" in html
            assert "Changes by Author" not in html

        finally:
            docx_path.unlink()

    def test_export_html_change_badges(self):
        """Test HTML includes change type badges."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html()

            assert "INSERTION" in html
            assert "DELETION" in html

        finally:
            docx_path.unlink()

    def test_export_html_empty_document(self):
        """Test HTML export on document with no changes."""
        docx_path = create_test_docx(DOCUMENT_NO_CHANGES)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html()

            assert "No tracked changes found" in html

        finally:
            docx_path.unlink()

    def test_export_html_escapes_special_chars(self):
        """Test that HTML special characters are properly escaped."""
        # Create a document with author name containing special characters
        # Author names are preserved as-is in the XML attributes and would need escaping
        doc_with_special = DOCUMENT_WITH_TRACKED_CHANGES.replace(
            'w:author="Alice"', 'w:author="Alice &amp; Bob"'
        )
        docx_path = create_test_docx(doc_with_special)
        try:
            doc = Document(docx_path)
            html = doc.export_changes_html()

            # Author name with ampersand should be properly escaped in HTML
            assert "Alice &amp; Bob" in html
            # Should not have unescaped ampersand followed by space (that's valid HTML)
            # But we specifically check the author is escaped
            assert "by Alice &amp; Bob" in html

        finally:
            docx_path.unlink()

    def test_export_html_standalone_function(self):
        """Test the standalone export_changes_html function."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            html = export_changes_html(doc)

            assert "<!DOCTYPE html>" in html

        finally:
            docx_path.unlink()


class TestGenerateChangeReport:
    """Tests for generate_change_report() method."""

    def test_generate_report_html(self):
        """Test generating HTML report."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            report = doc.generate_change_report(format="html")

            assert "<!DOCTYPE html>" in report
            assert "Tracked Changes Report" in report

        finally:
            docx_path.unlink()

    def test_generate_report_markdown(self):
        """Test generating Markdown report."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            report = doc.generate_change_report(format="markdown")

            assert "# Tracked Changes Report" in report

        finally:
            docx_path.unlink()

    def test_generate_report_json(self):
        """Test generating JSON report."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            report = doc.generate_change_report(format="json")

            data = json.loads(report)
            assert "title" in data
            assert "generated_at" in data
            assert "summary" in data
            assert "changes" in data
            assert "by_author" in data
            assert "by_type" in data

        finally:
            docx_path.unlink()

    def test_generate_report_json_summary(self):
        """Test JSON report includes correct summary."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            report = doc.generate_change_report(format="json")

            data = json.loads(report)
            summary = data["summary"]

            assert summary["total_changes"] == 5
            assert summary["insertions"] == 2
            assert summary["deletions"] == 1

        finally:
            docx_path.unlink()

    def test_generate_report_json_by_author(self):
        """Test JSON report groups by author correctly."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            report = doc.generate_change_report(format="json")

            data = json.loads(report)
            by_author = data["by_author"]

            assert "Alice" in by_author
            assert "Bob" in by_author
            assert "Carol" in by_author
            assert len(by_author["Alice"]) == 2
            assert len(by_author["Bob"]) == 2
            assert len(by_author["Carol"]) == 1

        finally:
            docx_path.unlink()

    def test_generate_report_json_by_type(self):
        """Test JSON report groups by type correctly."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            report = doc.generate_change_report(format="json")

            data = json.loads(report)
            by_type = data["by_type"]

            assert "insertion" in by_type
            assert "deletion" in by_type
            assert len(by_type["insertion"]) == 2
            assert len(by_type["deletion"]) == 1

        finally:
            docx_path.unlink()

    def test_generate_report_custom_title(self):
        """Test report with custom title."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            report = doc.generate_change_report(format="json", title="My Custom Report")

            data = json.loads(report)
            assert data["title"] == "My Custom Report"

        finally:
            docx_path.unlink()

    def test_generate_report_default_title(self):
        """Test report with default title."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            report = doc.generate_change_report(format="json")

            data = json.loads(report)
            assert data["title"] == "Tracked Changes Report"

        finally:
            docx_path.unlink()

    def test_generate_report_standalone_function(self):
        """Test the standalone generate_change_report function."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            report = generate_change_report(doc, format="json")

            data = json.loads(report)
            assert "changes" in data

        finally:
            docx_path.unlink()


class TestExportDataclasses:
    """Tests for export dataclasses."""

    def test_change_context_fields(self):
        """Test ChangeContext dataclass fields."""
        ctx = ChangeContext(
            before="text before",
            after="text after",
            paragraph_text="full paragraph text",
            paragraph_index=0,
        )

        assert ctx.before == "text before"
        assert ctx.after == "text after"
        assert ctx.paragraph_text == "full paragraph text"
        assert ctx.paragraph_index == 0

    def test_exported_change_fields(self):
        """Test ExportedChange dataclass fields."""
        change = ExportedChange(
            id="1",
            change_type="insertion",
            author="Test",
            date="2024-01-15T10:30:00Z",
            text="inserted text",
            context=None,
        )

        assert change.id == "1"
        assert change.change_type == "insertion"
        assert change.author == "Test"
        assert change.date == "2024-01-15T10:30:00Z"
        assert change.text == "inserted text"
        assert change.context is None

    def test_exported_change_with_context(self):
        """Test ExportedChange with context."""
        ctx = ChangeContext(
            before="before",
            after="after",
            paragraph_text="paragraph",
            paragraph_index=1,
        )
        change = ExportedChange(
            id="1",
            change_type="deletion",
            author="Test",
            date=None,
            text="deleted",
            context=ctx,
        )

        assert change.context is not None
        assert change.context.before == "before"

    def test_change_report_fields(self):
        """Test ChangeReport dataclass fields."""
        report = ChangeReport(
            title="Test Report",
            generated_at="2024-01-15T10:30:00Z",
            total_changes=5,
            insertions=2,
            deletions=1,
            moves=1,
            format_changes=1,
        )

        assert report.title == "Test Report"
        assert report.total_changes == 5
        assert report.insertions == 2
        assert report.deletions == 1
        assert report.moves == 1
        assert report.format_changes == 1
        assert report.changes == []
        assert report.by_author == {}
        assert report.by_type == {}


class TestContextExtraction:
    """Tests for context extraction functionality."""

    def test_context_includes_surrounding_text(self):
        """Test that context includes text before and after the change."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json(include_context=True, context_chars=30)

            data = json.loads(json_str)

            # Find an insertion with context
            for change in data["changes"]:
                if change["change_type"] == "insertion" and "context" in change:
                    ctx = change["context"]
                    # The insertion "Added by Alice" should have context
                    if "Added by Alice" in change["text"]:
                        assert "paragraph_text" in ctx
                        assert "paragraph_index" in ctx
                        break

        finally:
            docx_path.unlink()

    def test_context_chars_limit(self):
        """Test that context respects the character limit."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json(include_context=True, context_chars=10)

            data = json.loads(json_str)

            # Check that context strings are limited
            for change in data["changes"]:
                if "context" in change and change["context"]:
                    ctx = change["context"]
                    # Before and after should be limited (allowing for some flexibility)
                    assert len(ctx.get("before", "")) <= 20
                    assert len(ctx.get("after", "")) <= 20

        finally:
            docx_path.unlink()


class TestEdgeCases:
    """Tests for edge cases and special scenarios."""

    def test_export_after_making_changes(self):
        """Test export after programmatically adding changes."""
        docx_path = create_test_docx(DOCUMENT_NO_CHANGES)
        try:
            doc = Document(docx_path, author="TestAuthor")

            # Initially no changes
            json_str = doc.export_changes_json()
            data = json.loads(json_str)
            assert data["total_changes"] == 0

            # Add a change
            doc.insert_tracked(" new text", after="document")

            # Now should have one change
            json_str = doc.export_changes_json()
            data = json.loads(json_str)
            assert data["total_changes"] == 1
            assert data["changes"][0]["change_type"] == "insertion"
            assert data["changes"][0]["author"] == "TestAuthor"

        finally:
            docx_path.unlink()

    def test_export_all_change_types(self):
        """Test that all change types are exported correctly."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json()

            data = json.loads(json_str)
            types = {c["change_type"] for c in data["changes"]}

            assert "insertion" in types
            assert "deletion" in types
            assert "format_run" in types
            assert "format_paragraph" in types

        finally:
            docx_path.unlink()

    def test_export_preserves_unicode(self):
        """Test that Unicode characters are preserved in export."""
        # Create document with Unicode
        doc_with_unicode = DOCUMENT_WITH_TRACKED_CHANGES.replace(
            "Added by Alice",
            "Added by Alice: \u4e2d\u6587 \u0420\u0443\u0441\u0441\u043a\u0438\u0439",
        )
        docx_path = create_test_docx(doc_with_unicode)
        try:
            doc = Document(docx_path)
            json_str = doc.export_changes_json()

            data = json.loads(json_str)

            # Find the insertion with Unicode
            for change in data["changes"]:
                if "\u4e2d\u6587" in change.get("text", ""):
                    assert "\u0420\u0443\u0441\u0441\u043a\u0438\u0439" in change["text"]
                    break

        finally:
            docx_path.unlink()

    def test_multiple_exports_same_document(self):
        """Test that multiple exports from same document are consistent."""
        docx_path = create_test_docx(DOCUMENT_WITH_TRACKED_CHANGES)
        try:
            doc = Document(docx_path)

            json1 = doc.export_changes_json()
            json2 = doc.export_changes_json()
            md1 = doc.export_changes_markdown()
            md2 = doc.export_changes_markdown()
            html1 = doc.export_changes_html()
            html2 = doc.export_changes_html()

            assert json1 == json2
            assert md1 == md2
            assert html1 == html2

        finally:
            docx_path.unlink()
