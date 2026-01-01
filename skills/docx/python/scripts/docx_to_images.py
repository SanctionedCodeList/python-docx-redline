#!/usr/bin/env python3
"""
Convert Word documents to images for visual analysis.

This script converts a .docx file to page images, enabling AI agents to analyze
physical document layout, formatting, and visual elements that aren't captured
in the XML structure.

Requires:
    - LibreOffice (for DOCX → PDF conversion)
    - poppler (for PDF → PNG conversion via pdf2image)

Installation:
    macOS:
        brew install --cask libreoffice
        brew install poppler
        pip install pdf2image

    Linux (Ubuntu/Debian):
        sudo apt install libreoffice poppler-utils
        pip install pdf2image

    Windows:
        1. Install LibreOffice from https://www.libreoffice.org/download/
        2. Install poppler from https://github.com/oschwartz10612/poppler-windows/releases
        3. Add poppler bin/ to PATH
        4. pip install pdf2image

Usage:
    # From command line
    python docx_to_images.py document.docx

    # With custom output directory
    python docx_to_images.py document.docx --output ./pages

    # With custom DPI (default: 150)
    python docx_to_images.py document.docx --dpi 300

    # As a module
    from docx_to_images import docx_to_images
    images = docx_to_images("document.docx")
    # Returns: ["document_page_1.png", "document_page_2.png", ...]

Output:
    Creates PNG files named {basename}_page_{n}.png in the output directory.
    Returns list of created image paths.
"""

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def find_libreoffice() -> str | None:
    """Find LibreOffice executable on the system."""
    # Common executable names
    executables = ["soffice", "libreoffice"]

    # Check PATH first
    for exe in executables:
        path = shutil.which(exe)
        if path:
            return path

    # Platform-specific default locations
    if sys.platform == "darwin":
        # macOS
        app_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/Applications/OpenOffice.app/Contents/MacOS/soffice",
        ]
        for path in app_paths:
            if Path(path).exists():
                return path

    elif sys.platform == "win32":
        # Windows - check Program Files
        import os

        program_files = [
            os.environ.get("ProgramFiles", r"C:\Program Files"),
            os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
        ]
        for pf in program_files:
            lo_path = Path(pf) / "LibreOffice" / "program" / "soffice.exe"
            if lo_path.exists():
                return str(lo_path)

    return None


def check_dependencies() -> tuple[bool, list[str]]:
    """Check if required dependencies are installed.

    Returns:
        Tuple of (all_ok, list of missing dependencies)
    """
    missing = []

    # Check LibreOffice
    if not find_libreoffice():
        missing.append("LibreOffice (soffice)")

    # Check pdf2image (which requires poppler)
    try:
        from pdf2image import convert_from_path  # noqa: F401
    except ImportError:
        missing.append("pdf2image (pip install pdf2image)")

    # Check poppler by trying to find pdftoppm
    if not shutil.which("pdftoppm"):
        # On macOS/Linux, poppler-utils provides pdftoppm
        # On Windows, it needs to be in PATH
        missing.append("poppler (pdftoppm)")

    return len(missing) == 0, missing


def docx_to_pdf(docx_path: Path, output_dir: Path) -> Path:
    """Convert DOCX to PDF using LibreOffice.

    Args:
        docx_path: Path to the .docx file
        output_dir: Directory to save the PDF

    Returns:
        Path to the created PDF file

    Raises:
        RuntimeError: If LibreOffice is not found or conversion fails
    """
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice not found. Install it:\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Linux: sudo apt install libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )

    # LibreOffice command for PDF conversion
    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(docx_path),
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    # LibreOffice creates PDF with same basename
    pdf_path = output_dir / (docx_path.stem + ".pdf")

    if not pdf_path.exists():
        raise RuntimeError(
            f"PDF not created. LibreOffice output:\n{result.stdout}\n{result.stderr}"
        )

    return pdf_path


def pdf_to_images(
    pdf_path: Path,
    output_dir: Path,
    basename: str,
    dpi: int = 150,
) -> list[Path]:
    """Convert PDF to images using pdf2image.

    Args:
        pdf_path: Path to the PDF file
        output_dir: Directory to save images
        basename: Base name for output files
        dpi: Resolution in dots per inch (default: 150)

    Returns:
        List of paths to created image files

    Raises:
        RuntimeError: If pdf2image or poppler is not available
    """
    try:
        from pdf2image import convert_from_path
    except ImportError:
        raise RuntimeError(
            "pdf2image not installed. Run: pip install pdf2image\n"
            "Also requires poppler:\n"
            "  macOS: brew install poppler\n"
            "  Linux: sudo apt install poppler-utils\n"
            "  Windows: https://github.com/oschwartz10612/poppler-windows/releases"
        )

    try:
        images = convert_from_path(pdf_path, dpi=dpi)
    except Exception as e:
        if "poppler" in str(e).lower() or "pdftoppm" in str(e).lower():
            raise RuntimeError(
                f"poppler not found. Install it:\n"
                f"  macOS: brew install poppler\n"
                f"  Linux: sudo apt install poppler-utils\n"
                f"  Windows: https://github.com/oschwartz10612/poppler-windows/releases\n"
                f"Original error: {e}"
            )
        raise

    output_paths = []
    for i, image in enumerate(images, 1):
        output_path = output_dir / f"{basename}_page_{i}.png"
        image.save(str(output_path), "PNG")
        output_paths.append(output_path)

    return output_paths


def docx_to_images(
    docx_path: str | Path,
    output_dir: str | Path | None = None,
    dpi: int = 150,
    cleanup_pdf: bool = True,
) -> list[Path]:
    """Convert a Word document to page images.

    Args:
        docx_path: Path to the .docx file
        output_dir: Directory to save images (default: same as docx)
        dpi: Resolution in dots per inch (default: 150)
        cleanup_pdf: Remove intermediate PDF after conversion (default: True)

    Returns:
        List of paths to created PNG images

    Raises:
        FileNotFoundError: If docx_path doesn't exist
        RuntimeError: If dependencies are missing or conversion fails

    Example:
        >>> images = docx_to_images("contract.docx")
        >>> print(images)
        [PosixPath('contract_page_1.png'), PosixPath('contract_page_2.png')]

        >>> images = docx_to_images("report.docx", output_dir="./pages", dpi=300)
    """
    docx_path = Path(docx_path).resolve()

    if not docx_path.exists():
        raise FileNotFoundError(f"Document not found: {docx_path}")

    if docx_path.suffix.lower() not in (".docx", ".doc"):
        raise ValueError(f"Expected .docx or .doc file, got: {docx_path.suffix}")

    # Set output directory
    if output_dir is None:
        output_dir = docx_path.parent
    else:
        output_dir = Path(output_dir).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

    basename = docx_path.stem

    # Use temp directory for intermediate PDF
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        # Step 1: DOCX → PDF
        pdf_path = docx_to_pdf(docx_path, temp_path)

        # Step 2: PDF → Images
        image_paths = pdf_to_images(pdf_path, output_dir, basename, dpi=dpi)

        # PDF is automatically cleaned up when temp_dir is deleted

    return image_paths


def main():
    """Command-line interface."""
    parser = argparse.ArgumentParser(
        description="Convert Word documents to images for visual analysis.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    %(prog)s document.docx
    %(prog)s document.docx --output ./pages
    %(prog)s document.docx --dpi 300

Dependencies:
    - LibreOffice (for DOCX → PDF)
    - poppler (for PDF → PNG)
    - pdf2image (pip install pdf2image)
""",
    )

    parser.add_argument("docx_path", nargs="?", help="Path to the .docx file")
    parser.add_argument("--output", "-o", help="Output directory (default: same as input)")
    parser.add_argument("--dpi", type=int, default=150, help="Image resolution (default: 150)")
    parser.add_argument("--check", action="store_true", help="Check dependencies and exit")

    args = parser.parse_args()

    # Validate that docx_path is provided if not just checking
    if not args.check and not args.docx_path:
        parser.error("docx_path is required unless --check is specified")

    # Check dependencies
    if args.check:
        ok, missing = check_dependencies()
        if ok:
            print("All dependencies installed!")
            sys.exit(0)
        else:
            print("Missing dependencies:")
            for dep in missing:
                print(f"  - {dep}")
            sys.exit(1)

    try:
        images = docx_to_images(
            args.docx_path,
            output_dir=args.output,
            dpi=args.dpi,
        )
        print(f"Created {len(images)} image(s):")
        for img in images:
            print(f"  {img}")

    except (FileNotFoundError, ValueError, RuntimeError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
