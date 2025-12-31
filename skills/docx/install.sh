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

# Optional: pandoc for text extraction (check if available)
if command -v pandoc &> /dev/null; then
    echo "pandoc already installed"
else
    echo "Note: pandoc not found. For text extraction, install via:"
    echo "  macOS: brew install pandoc"
    echo "  Linux: apt install pandoc"
    echo "  Windows: choco install pandoc"
fi

echo ""
echo "DOCX skill dependencies installed successfully!"
echo ""
echo "Quick test:"
echo "  python -c 'from python_docx_redline import Document; print(\"Ready!\")'"
