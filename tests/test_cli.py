"""Tests for the command-line interface."""

import tempfile
import zipfile
from pathlib import Path

from typer.testing import CliRunner

from python_docx_redline.cli import app

runner = CliRunner()


def create_test_docx(path: Path, content: str = "Hello world") -> None:
    """Create a minimal valid .docx file for testing."""
    word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    doc_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{word_ns}">
    <w:body>
        <w:p>
            <w:r>
                <w:t>{content}</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""

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

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)


class TestCLIVersion:
    """Tests for version command."""

    def test_version_flag(self):
        """Test --version shows version."""
        result = runner.invoke(app, ["--version"])
        assert result.exit_code == 0
        assert "0.2.0" in result.stdout

    def test_version_short_flag(self):
        """Test -v shows version."""
        result = runner.invoke(app, ["-v"])
        assert result.exit_code == 0
        assert "0.2.0" in result.stdout


class TestCLIHelp:
    """Tests for help command."""

    def test_main_help(self):
        """Test main --help shows commands."""
        result = runner.invoke(app, ["--help"])
        assert result.exit_code == 0
        assert "insert" in result.stdout
        assert "replace" in result.stdout
        assert "delete" in result.stdout
        assert "move" in result.stdout
        assert "accept-all" in result.stdout
        assert "apply" in result.stdout
        assert "info" in result.stdout

    def test_insert_help(self):
        """Test insert --help shows options."""
        result = runner.invoke(app, ["insert", "--help"])
        assert result.exit_code == 0
        assert "--text" in result.stdout
        assert "--after" in result.stdout
        assert "--before" in result.stdout
        assert "--author" in result.stdout
        assert "--output" in result.stdout

    def test_replace_help(self):
        """Test replace --help shows options."""
        result = runner.invoke(app, ["replace", "--help"])
        assert result.exit_code == 0
        assert "--find" in result.stdout
        assert "--replace" in result.stdout
        assert "--regex" in result.stdout


class TestCLIInsert:
    """Tests for insert command."""

    def test_insert_requires_after_or_before(self):
        """Test insert fails without --after or --before."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            doc_path = Path(tmp_dir) / "test.docx"
            create_test_docx(doc_path)

            result = runner.invoke(app, ["insert", str(doc_path), "--text", "new text"])
            assert result.exit_code == 1
            assert "Must specify either --after or --before" in result.output

    def test_insert_rejects_both_after_and_before(self):
        """Test insert fails with both --after and --before."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            doc_path = Path(tmp_dir) / "test.docx"
            create_test_docx(doc_path)

            result = runner.invoke(
                app,
                [
                    "insert",
                    str(doc_path),
                    "--text",
                    "new text",
                    "--after",
                    "Hello",
                    "--before",
                    "world",
                ],
            )
            assert result.exit_code == 1
            assert "Cannot specify both --after and --before" in result.output

    def test_insert_with_after(self):
        """Test insert with --after option."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            doc_path = Path(tmp_dir) / "test.docx"
            output_path = Path(tmp_dir) / "output.docx"
            create_test_docx(doc_path, "Hello world")

            result = runner.invoke(
                app,
                [
                    "insert",
                    str(doc_path),
                    "--text",
                    " beautiful",
                    "--after",
                    "Hello",
                    "--output",
                    str(output_path),
                ],
            )
            assert result.exit_code == 0
            assert "Inserted text" in result.stdout
            assert output_path.exists()


class TestCLIReplace:
    """Tests for replace command."""

    def test_replace_basic(self):
        """Test basic replace operation."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            doc_path = Path(tmp_dir) / "test.docx"
            output_path = Path(tmp_dir) / "output.docx"
            create_test_docx(doc_path, "Hello world")

            result = runner.invoke(
                app,
                [
                    "replace",
                    str(doc_path),
                    "--find",
                    "world",
                    "--replace",
                    "universe",
                    "--output",
                    str(output_path),
                ],
            )
            assert result.exit_code == 0
            assert "Replaced text" in result.stdout
            assert output_path.exists()


class TestCLIDelete:
    """Tests for delete command."""

    def test_delete_basic(self):
        """Test basic delete operation."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            doc_path = Path(tmp_dir) / "test.docx"
            output_path = Path(tmp_dir) / "output.docx"
            create_test_docx(doc_path, "Hello beautiful world")

            result = runner.invoke(
                app,
                [
                    "delete",
                    str(doc_path),
                    "--text",
                    "beautiful ",
                    "--output",
                    str(output_path),
                ],
            )
            assert result.exit_code == 0
            assert "Deleted text" in result.stdout
            assert output_path.exists()


class TestCLIMove:
    """Tests for move command."""

    def test_move_requires_after_or_before(self):
        """Test move fails without --after or --before."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            doc_path = Path(tmp_dir) / "test.docx"
            create_test_docx(doc_path)

            result = runner.invoke(app, ["move", str(doc_path), "--text", "Hello"])
            assert result.exit_code == 1
            assert "Must specify either --after or --before" in result.output


class TestCLIAcceptAll:
    """Tests for accept-all command."""

    def test_accept_all_basic(self):
        """Test accept-all on document."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            doc_path = Path(tmp_dir) / "test.docx"
            output_path = Path(tmp_dir) / "output.docx"
            create_test_docx(doc_path)

            result = runner.invoke(app, ["accept-all", str(doc_path), "--output", str(output_path)])
            assert result.exit_code == 0
            assert "Accepted all changes" in result.stdout


class TestCLIInfo:
    """Tests for info command."""

    def test_info_basic(self):
        """Test info command shows document details."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            doc_path = Path(tmp_dir) / "test.docx"
            create_test_docx(doc_path)

            result = runner.invoke(app, ["info", str(doc_path)])
            assert result.exit_code == 0
            assert "Paragraphs:" in result.stdout
            assert "Tables:" in result.stdout
            assert "Has tracked changes:" in result.stdout


class TestCLIErrorHandling:
    """Tests for error handling."""

    def test_nonexistent_file(self):
        """Test error message for missing file."""
        result = runner.invoke(
            app,
            ["insert", "/nonexistent/file.docx", "--text", "test", "--after", "foo"],
        )
        assert result.exit_code == 1
        assert "Error:" in result.output

    def test_text_not_found(self):
        """Test error message when search text not found."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            doc_path = Path(tmp_dir) / "test.docx"
            create_test_docx(doc_path, "Hello world")

            result = runner.invoke(
                app,
                [
                    "insert",
                    str(doc_path),
                    "--text",
                    "new",
                    "--after",
                    "nonexistent text",
                ],
            )
            assert result.exit_code == 1
            assert "Error:" in result.output
