"""
Example: Batch processing multiple documents with the same edits.

This example demonstrates how to apply a consistent set of edits
to multiple Word documents, useful for updating templates or
processing multiple versions of similar contracts.
"""

from pathlib import Path
from typing import Any

from docx_redline import Document

# Define a standard set of edits to apply
STANDARD_EDITS: list[dict[str, Any]] = [
    # Update company name throughout
    {
        "type": "replace_tracked",
        "find": "Acme Corporation",
        "replace": "Acme Industries Inc.",
    },
    # Update payment terms
    {
        "type": "replace_tracked",
        "find": "net 30 days",
        "replace": "net 45 days",
        "scope": "payment",
    },
    # Add new compliance clause
    {
        "type": "insert_paragraph",
        "text": "Compliance with GDPR",
        "after": "Termination",
        "style": "Heading1",
        "track": True,
    },
    {
        "type": "insert_paragraphs",
        "texts": [
            "Both parties shall comply with all applicable data protection laws.",
            "This includes GDPR and other privacy regulations.",
        ],
        "after": "Compliance with GDPR",
        "track": True,
    },
]


def process_document(
    input_path: Path, output_path: Path, edits: list[dict[str, Any]]
) -> dict[str, Any]:
    """
    Process a single document with the given edits.

    Args:
        input_path: Path to input document
        output_path: Path for output document
        edits: List of edit specifications

    Returns:
        Dictionary with processing results
    """
    try:
        doc = Document(input_path)
        results = doc.apply_edits(edits)

        # Save the edited document
        doc.save(output_path)

        # Count successes and failures
        success_count = sum(1 for r in results if r.success)
        failure_count = len(results) - success_count

        return {
            "status": "success",
            "input": str(input_path),
            "output": str(output_path),
            "total_edits": len(results),
            "successful": success_count,
            "failed": failure_count,
            "results": results,
        }

    except Exception as e:
        return {
            "status": "error",
            "input": str(input_path),
            "error": str(e),
        }


def batch_process(
    input_dir: Path,
    output_dir: Path,
    pattern: str = "*.docx",
    edits: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    """
    Process all documents in a directory with the same edits.

    Args:
        input_dir: Directory containing input documents
        output_dir: Directory for output documents
        pattern: Glob pattern for files to process (default: *.docx)
        edits: List of edit specifications (default: STANDARD_EDITS)

    Returns:
        List of processing results for each document
    """
    if edits is None:
        edits = STANDARD_EDITS

    # Ensure output directory exists
    output_dir.mkdir(parents=True, exist_ok=True)

    results = []

    # Process each matching file
    for input_path in input_dir.glob(pattern):
        # Skip if not a file
        if not input_path.is_file():
            continue

        # Create output path with same filename
        output_path = output_dir / input_path.name

        print(f"Processing: {input_path.name}...")
        result = process_document(input_path, output_path, edits)
        results.append(result)

        # Print summary
        if result["status"] == "success":
            print(f"  ✓ {result['successful']}/{result['total_edits']} edits applied successfully")
        else:
            print(f"  ✗ Error: {result['error']}")

    return results


def main() -> None:
    """Example usage of batch processing."""
    # Define directories
    input_dir = Path("./input_contracts")
    output_dir = Path("./output_contracts")

    # Process all documents
    print("Starting batch processing...")
    print("-" * 60)

    results = batch_process(input_dir, output_dir, edits=STANDARD_EDITS)

    # Print overall summary
    print("-" * 60)
    print("\nBatch Processing Summary:")
    print(f"  Total documents: {len(results)}")
    successful = sum(1 for r in results if r["status"] == "success")
    print(f"  Successfully processed: {successful}")
    print(f"  Failed: {len(results) - successful}")

    # Show any failures
    failures = [r for r in results if r["status"] == "error"]
    if failures:
        print("\nFailed documents:")
        for failure in failures:
            print(f"  • {failure['input']}: {failure['error']}")


if __name__ == "__main__":
    main()
