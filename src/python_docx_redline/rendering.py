"""
Document rendering to images using LibreOffice.

Converts DOCX documents to page images (PNG) for visual inspection.
Useful for AI agents to see document layout and tracked changes.

Usage:
    from python_docx_redline import Document
    from python_docx_redline.rendering import is_libreoffice_available

    if is_libreoffice_available():
        doc = Document("contract.docx")
        images = doc.render_to_images()  # Returns list of Path objects

Requirements:
    - LibreOffice must be installed
    - pdftoppm (from poppler-utils) for PDF to image conversion

Set LIBREOFFICE_PATH environment variable to override default search.
Set PDFTOPPM_PATH environment variable to override pdftoppm search.
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)

# Environment variable names for tool path overrides
LIBREOFFICE_PATH_ENV = "LIBREOFFICE_PATH"
PDFTOPPM_PATH_ENV = "PDFTOPPM_PATH"

# Cache for tool paths to avoid repeated filesystem searches
_libreoffice_path_cache: str | None = None
_libreoffice_checked: bool = False
_pdftoppm_path_cache: str | None = None
_pdftoppm_checked: bool = False

# Default search locations for LibreOffice
LIBREOFFICE_DEFAULT_PATHS = [
    # macOS
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    # Linux common locations
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/usr/local/bin/soffice",
    "/usr/local/bin/libreoffice",
    # Windows
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]

# Default search locations for pdftoppm
PDFTOPPM_DEFAULT_PATHS = [
    # macOS (Homebrew)
    "/opt/homebrew/bin/pdftoppm",
    "/usr/local/bin/pdftoppm",
    # Linux
    "/usr/bin/pdftoppm",
]


def _find_libreoffice() -> str | None:
    """Find the LibreOffice soffice executable.

    Search order:
    1. LIBREOFFICE_PATH environment variable
    2. Default installation paths
    3. System PATH

    Returns:
        Path to soffice executable, or None if not found
    """
    global _libreoffice_path_cache, _libreoffice_checked

    if _libreoffice_checked:
        return _libreoffice_path_cache

    _libreoffice_checked = True

    # Check environment variable first
    env_path = os.environ.get(LIBREOFFICE_PATH_ENV)
    if env_path:
        if os.path.isfile(env_path) and os.access(env_path, os.X_OK):
            logger.debug("Found LibreOffice at %s (from env)", env_path)
            _libreoffice_path_cache = env_path
            return env_path
        else:
            logger.warning(
                "%s is set to %s but file is not executable",
                LIBREOFFICE_PATH_ENV,
                env_path,
            )

    # Check default paths
    for path in LIBREOFFICE_DEFAULT_PATHS:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            logger.debug("Found LibreOffice at %s", path)
            _libreoffice_path_cache = path
            return path

    # Check system PATH
    for cmd in ["soffice", "libreoffice"]:
        which_result = shutil.which(cmd)
        if which_result:
            logger.debug("Found LibreOffice at %s (from PATH)", which_result)
            _libreoffice_path_cache = which_result
            return which_result

    logger.debug("LibreOffice not found")
    return None


def _find_pdftoppm() -> str | None:
    """Find the pdftoppm executable (from poppler-utils).

    Search order:
    1. PDFTOPPM_PATH environment variable
    2. Default installation paths
    3. System PATH

    Returns:
        Path to pdftoppm executable, or None if not found
    """
    global _pdftoppm_path_cache, _pdftoppm_checked

    if _pdftoppm_checked:
        return _pdftoppm_path_cache

    _pdftoppm_checked = True

    # Check environment variable first
    env_path = os.environ.get(PDFTOPPM_PATH_ENV)
    if env_path:
        if os.path.isfile(env_path) and os.access(env_path, os.X_OK):
            logger.debug("Found pdftoppm at %s (from env)", env_path)
            _pdftoppm_path_cache = env_path
            return env_path
        else:
            logger.warning(
                "%s is set to %s but file is not executable",
                PDFTOPPM_PATH_ENV,
                env_path,
            )

    # Check default paths
    for path in PDFTOPPM_DEFAULT_PATHS:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            logger.debug("Found pdftoppm at %s", path)
            _pdftoppm_path_cache = path
            return path

    # Check system PATH
    which_result = shutil.which("pdftoppm")
    if which_result:
        logger.debug("Found pdftoppm at %s (from PATH)", which_result)
        _pdftoppm_path_cache = which_result
        return which_result

    logger.debug("pdftoppm not found")
    return None


def is_libreoffice_available() -> bool:
    """Check if LibreOffice is available for rendering.

    Returns:
        True if LibreOffice (soffice) is found and executable
    """
    return _find_libreoffice() is not None


def is_pdftoppm_available() -> bool:
    """Check if pdftoppm is available for PDF to image conversion.

    Returns:
        True if pdftoppm is found and executable
    """
    return _find_pdftoppm() is not None


def is_rendering_available() -> bool:
    """Check if all rendering dependencies are available.

    Returns:
        True if both LibreOffice and pdftoppm are available
    """
    return is_libreoffice_available() and is_pdftoppm_available()


def render_document_to_images(
    docx_path: str | Path,
    output_dir: str | Path | None = None,
    dpi: int = 150,
    prefix: str = "page",
    timeout: int = 120,
) -> list[Path]:
    """Render a DOCX document to PNG page images.

    Uses LibreOffice to convert DOCX to PDF, then pdftoppm to convert
    PDF pages to PNG images.

    Args:
        docx_path: Path to the DOCX file to render
        output_dir: Directory for output images. If None, creates a temp directory.
        dpi: Resolution in dots per inch (default: 150)
        prefix: Filename prefix for images (default: "page")
        timeout: Timeout in seconds for each conversion step (default: 120)

    Returns:
        List of Path objects for generated PNG files, sorted by page number
        (e.g., [page-1.png, page-2.png, ...])

    Raises:
        RuntimeError: If LibreOffice or pdftoppm is not available
        RuntimeError: If conversion fails
        FileNotFoundError: If the input DOCX file doesn't exist
    """
    docx_path = Path(docx_path)

    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")

    # Check tool availability
    libreoffice_path = _find_libreoffice()
    if libreoffice_path is None:
        raise RuntimeError(
            "LibreOffice is not available. "
            "Install LibreOffice to enable document rendering:\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Linux: sudo apt install libreoffice\n"
            "  Windows: Download from https://www.libreoffice.org/download/\n"
            f"Or set {LIBREOFFICE_PATH_ENV} environment variable."
        )

    pdftoppm_path = _find_pdftoppm()
    if pdftoppm_path is None:
        raise RuntimeError(
            "pdftoppm is not available. "
            "Install poppler-utils to enable PDF to image conversion:\n"
            "  macOS: brew install poppler\n"
            "  Linux: sudo apt install poppler-utils\n"
            "  Windows: Download from https://github.com/oschwartz10612/poppler-windows\n"
            f"Or set {PDFTOPPM_PATH_ENV} environment variable."
        )

    # Set up output directory
    temp_dir_created = False
    if output_dir is None:
        output_dir = Path(tempfile.mkdtemp(prefix="docx_render_"))
        temp_dir_created = True
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Step 1: Convert DOCX to PDF using LibreOffice
        logger.debug("Converting %s to PDF using LibreOffice", docx_path)

        # LibreOffice needs absolute paths and outputs to --outdir
        result = subprocess.run(
            [
                libreoffice_path,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output_dir.absolute()),
                str(docx_path.absolute()),
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
        )

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit code {result.returncode}):\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

        # Find the generated PDF
        pdf_path = output_dir / f"{docx_path.stem}.pdf"
        if not pdf_path.exists():
            # Sometimes LibreOffice uses different naming
            pdf_files = list(output_dir.glob("*.pdf"))
            if pdf_files:
                pdf_path = pdf_files[0]
            else:
                raise RuntimeError(
                    f"LibreOffice did not generate PDF file. "
                    f"Expected: {pdf_path}\n"
                    f"stdout: {result.stdout}\n"
                    f"stderr: {result.stderr}"
                )

        logger.debug("PDF generated at %s", pdf_path)

        # Step 2: Convert PDF to PNG images using pdftoppm
        logger.debug("Converting PDF to PNG images at %d DPI", dpi)

        output_prefix = output_dir / prefix
        result = subprocess.run(
            [
                pdftoppm_path,
                "-png",
                "-r",
                str(dpi),
                str(pdf_path),
                str(output_prefix),
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
        )

        if result.returncode != 0:
            raise RuntimeError(
                f"pdftoppm conversion failed (exit code {result.returncode}):\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

        # Find generated PNG files and sort by page number
        png_files = sorted(output_dir.glob(f"{prefix}-*.png"))

        if not png_files:
            # pdftoppm might use different naming for single-page docs
            png_files = sorted(output_dir.glob(f"{prefix}*.png"))

        if not png_files:
            raise RuntimeError(
                f"pdftoppm did not generate any PNG files.\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

        logger.debug("Generated %d page image(s)", len(png_files))

        # Clean up PDF file (keep only PNGs)
        pdf_path.unlink()

        return png_files

    except subprocess.TimeoutExpired as e:
        raise RuntimeError(
            f"Rendering timed out after {timeout} seconds. Try increasing the timeout parameter."
        ) from e
    except Exception:
        # Clean up temp directory on failure if we created it
        if temp_dir_created and output_dir.exists():
            shutil.rmtree(output_dir, ignore_errors=True)
        raise
