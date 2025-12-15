"""
External OOXML validator integration.

This module provides optional integration with the OOXML-Validator tool
(https://github.com/mikeebowen/OOXML-Validator) for full OOXML spec validation.

The external validator validates against the complete OOXML specification,
catching issues that lightweight Python validation might miss.

Usage:
    from python_docx_redline.ooxml_validator import (
        is_ooxml_validator_available,
        validate_with_ooxml_validator,
    )

    if is_ooxml_validator_available():
        errors = validate_with_ooxml_validator("document.docx")
        if errors:
            print(f"Validation errors: {errors}")
"""

import json
import logging
import os
import shutil
import subprocess
from pathlib import Path

logger = logging.getLogger(__name__)

# Cache for validator availability check
_validator_path_cache: str | None = None
_validator_checked: bool = False

# Environment variable to specify custom validator path
OOXML_VALIDATOR_PATH_ENV = "OOXML_VALIDATOR_PATH"

# Default locations to search for the validator
DEFAULT_VALIDATOR_PATHS = [
    # Global dotnet tool
    "ooxml-validator",
    # Built from source (common locations)
    "/tmp/OOXML-Validator/OOXMLValidatorCLI/bin/Release/net8.0/osx-arm64/publish/OOXMLValidatorCLI",
    "/tmp/OOXML-Validator/OOXMLValidatorCLI/bin/Release/net8.0/osx-x64/publish/OOXMLValidatorCLI",
    "/tmp/OOXML-Validator/OOXMLValidatorCLI/bin/Release/net8.0/linux-x64/publish/OOXMLValidatorCLI",
    "/tmp/OOXML-Validator/OOXMLValidatorCLI/bin/Release/net8.0/linux-arm64/publish/OOXMLValidatorCLI",
    "/tmp/OOXML-Validator/OOXMLValidatorCLI/bin/Release/net8.0/win-x64/publish/OOXMLValidatorCLI.exe",
    # User home directory
    os.path.expanduser(
        "~/OOXML-Validator/OOXMLValidatorCLI/bin/Release/net8.0/osx-arm64/publish/OOXMLValidatorCLI"
    ),
    os.path.expanduser(
        "~/OOXML-Validator/OOXMLValidatorCLI/bin/Release/net8.0/linux-x64/publish/OOXMLValidatorCLI"
    ),
]


def _find_validator() -> str | None:
    """Find the OOXML-Validator executable.

    Searches in order:
    1. OOXML_VALIDATOR_PATH environment variable
    2. Default installation locations
    3. System PATH

    Returns:
        Path to validator executable, or None if not found
    """
    global _validator_path_cache, _validator_checked

    if _validator_checked:
        return _validator_path_cache

    _validator_checked = True

    # Check environment variable first
    env_path = os.environ.get(OOXML_VALIDATOR_PATH_ENV)
    if env_path and os.path.isfile(env_path) and os.access(env_path, os.X_OK):
        _validator_path_cache = env_path
        logger.debug("Found OOXML validator at %s (from environment)", env_path)
        return _validator_path_cache

    # Check default paths
    for path in DEFAULT_VALIDATOR_PATHS:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            _validator_path_cache = path
            logger.debug("Found OOXML validator at %s", path)
            return _validator_path_cache

    # Check system PATH
    validator_in_path = shutil.which("OOXMLValidatorCLI") or shutil.which("ooxml-validator")
    if validator_in_path:
        _validator_path_cache = validator_in_path
        logger.debug("Found OOXML validator in PATH: %s", validator_in_path)
        return _validator_path_cache

    logger.debug("OOXML validator not found")
    return None


def is_ooxml_validator_available() -> bool:
    """Check if the OOXML-Validator tool is available.

    Returns:
        True if the validator is installed and accessible
    """
    return _find_validator() is not None


def get_ooxml_validator_path() -> str | None:
    """Get the path to the OOXML-Validator executable.

    Returns:
        Path to the validator, or None if not available
    """
    return _find_validator()


def validate_with_ooxml_validator(
    docx_path: str | Path,
    office_version: str = "Microsoft365",
    timeout: int = 30,
) -> list[dict]:
    """Validate a document using the external OOXML-Validator.

    Args:
        docx_path: Path to the .docx file to validate
        office_version: Office version to validate against. One of:
            Office2007, Office2010, Office2013, Office2016,
            Office2019, Office2021, Microsoft365 (default)
        timeout: Maximum seconds to wait for validation (default: 30)

    Returns:
        List of validation error dictionaries. Empty list means valid.
        Each error dict contains keys like:
        - Description: Error description
        - Path: XPath to the error location
        - Part: The document part containing the error

    Raises:
        RuntimeError: If validator is not available
        subprocess.TimeoutExpired: If validation takes too long
        ValueError: If the file doesn't exist
    """
    validator_path = _find_validator()
    if validator_path is None:
        raise RuntimeError(
            "OOXML-Validator is not available. "
            "Install from https://github.com/mikeebowen/OOXML-Validator "
            f"or set {OOXML_VALIDATOR_PATH_ENV} environment variable."
        )

    docx_path = Path(docx_path)
    if not docx_path.exists():
        raise ValueError(f"File not found: {docx_path}")

    try:
        result = subprocess.run(
            [validator_path, str(docx_path.absolute()), office_version],
            capture_output=True,
            text=True,
            timeout=timeout,
        )

        # Parse JSON output
        output = result.stdout.strip()
        if not output:
            return []

        errors = json.loads(output)
        return errors if isinstance(errors, list) else []

    except json.JSONDecodeError as e:
        logger.warning("Failed to parse validator output: %s", e)
        return [{"Description": f"Failed to parse validator output: {e}", "Raw": result.stdout}]
    except subprocess.TimeoutExpired:
        logger.warning("OOXML validation timed out after %d seconds", timeout)
        raise


class OOXMLValidationError(Exception):
    """Raised when OOXML validation fails."""

    def __init__(self, message: str, errors: list[dict]):
        super().__init__(message)
        self.errors = errors

    def __str__(self) -> str:
        if not self.errors:
            return super().__str__()

        error_lines = []
        for err in self.errors[:10]:  # Limit to first 10 errors
            desc = err.get("Description", str(err))
            part = err.get("Part", "")
            if part:
                error_lines.append(f"  - [{part}] {desc}")
            else:
                error_lines.append(f"  - {desc}")

        if len(self.errors) > 10:
            error_lines.append(f"  ... and {len(self.errors) - 10} more errors")

        return f"{super().__str__()}\n" + "\n".join(error_lines)


def validate_docx_strict(
    docx_path: str | Path,
    office_version: str = "Microsoft365",
    raise_on_error: bool = True,
) -> list[dict]:
    """Validate a document strictly against the full OOXML specification.

    This function uses the external OOXML-Validator if available, providing
    comprehensive validation against the complete OOXML spec.

    Args:
        docx_path: Path to the .docx file to validate
        office_version: Office version to validate against (default: Microsoft365)
        raise_on_error: If True, raise OOXMLValidationError on validation failure

    Returns:
        List of validation errors (empty if valid)

    Raises:
        OOXMLValidationError: If validation fails and raise_on_error is True
        RuntimeError: If validator is not available
    """
    errors = validate_with_ooxml_validator(docx_path, office_version)

    if errors and raise_on_error:
        raise OOXMLValidationError(
            f"OOXML validation failed with {len(errors)} error(s)",
            errors,
        )

    return errors
