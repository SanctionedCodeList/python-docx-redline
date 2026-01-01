#!/bin/bash
# DOCX Skill - Dependency Installation
# Idempotent: safe to run multiple times

set -e

echo "Installing DOCX skill dependencies..."

# Check for uv (preferred) or pip
if command -v uv &> /dev/null; then
    INSTALLER="uv pip install"
    echo "Using uv for installation"
elif command -v pip &> /dev/null; then
    INSTALLER="pip install"
    echo "Using pip for installation"
else
    echo "Error: Neither uv nor pip found. Please install Python first."
    exit 1
fi

# Core dependencies
$INSTALLER python-docx          # Creating new documents
$INSTALLER python-docx-redline  # Editing with tracked changes (recommended for all editing)

# Vision/layout analysis dependencies
$INSTALLER pdf2image            # For docx_to_images.py script

# Check for LibreOffice (required for DOCX → PDF conversion)
if command -v soffice &> /dev/null || command -v libreoffice &> /dev/null; then
    echo "LibreOffice already installed"
else
    echo ""
    echo "Note: LibreOffice not found. Required for docx_to_images.py script."
    echo "Install via:"
    echo "  macOS: brew install --cask libreoffice"
    echo "  Linux: sudo apt install libreoffice"
    echo "  Windows: https://www.libreoffice.org/download/"
fi

# Check for poppler (required for PDF → PNG conversion)
if command -v pdftoppm &> /dev/null; then
    echo "poppler already installed"
else
    echo ""
    echo "Note: poppler not found. Required for docx_to_images.py script."
    echo "Install via:"
    echo "  macOS: brew install poppler"
    echo "  Linux: sudo apt install poppler-utils"
    echo "  Windows: https://github.com/oschwartz10612/poppler-windows/releases"
fi

# Optional: pandoc for text extraction
if command -v pandoc &> /dev/null; then
    echo "pandoc already installed"
else
    echo ""
    echo "Note: pandoc not found. Optional for advanced text extraction."
    echo "Install via:"
    echo "  macOS: brew install pandoc"
    echo "  Linux: apt install pandoc"
    echo "  Windows: choco install pandoc"
fi

echo ""
echo "DOCX skill dependencies installed successfully!"
echo ""
echo "Quick test:"
echo "  python -c 'from python_docx_redline import Document; print(\"Ready!\")'"
echo ""
echo "Check docx_to_images.py dependencies:"
echo "  python python/scripts/docx_to_images.py --check"
