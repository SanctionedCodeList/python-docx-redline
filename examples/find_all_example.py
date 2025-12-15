"""
Example demonstrating the find_all() method for previewing text matches.

This example shows how to use find_all() to preview where text appears
in a document before making changes. This is especially useful for:
- Understanding ambiguous text searches
- Previewing regex matches
- Deciding which occurrences to modify
"""

from pathlib import Path

from python_docx_redline import Document


def main():
    """Demonstrate find_all() usage."""
    # Load a document
    doc_path = Path(__file__).parent.parent / "tests" / "fixtures" / "minimal.docx"
    doc = Document(doc_path)

    print("=" * 60)
    print("find_all() Example")
    print("=" * 60)

    # Example 1: Basic search
    print("\n1. Basic search for 'test':")
    matches = doc.find_all("test")
    print(f"   Found {len(matches)} occurrence(s)")
    for match in matches:
        print(f"   {match}")

    # Example 2: Case-insensitive search
    print("\n2. Case-insensitive search:")
    matches = doc.find_all("TEST", case_sensitive=False)
    print(f"   Found {len(matches)} occurrence(s)")
    for match in matches:
        print(f"   [{match.index}] {match.location}: {match.context}")

    # Example 3: Regex search
    print("\n3. Regex search for words starting with 't':")
    matches = doc.find_all(r"\bt\w+", regex=True)
    print(f"   Found {len(matches)} occurrence(s)")
    for match in matches:
        print(f"   [{match.index}] '{match.text}' at {match.location}")

    # Example 4: Custom context size
    print("\n4. Search with custom context (10 chars):")
    matches = doc.find_all("test", context_chars=10)
    for match in matches:
        print(f"   Context: {match.context}")

    # Example 5: Accessing match metadata
    print("\n5. Accessing detailed match metadata:")
    matches = doc.find_all("test")
    if matches:
        match = matches[0]
        print(f"   Text: {match.text}")
        print(f"   Index: {match.index}")
        print(f"   Location: {match.location}")
        print(f"   Paragraph index: {match.paragraph_index}")
        print(f"   Paragraph text: {match.paragraph_text[:50]}...")
        print(f"   Context: {match.context}")
        print(f"   Has TextSpan: {match.span is not None}")

    print("\n" + "=" * 60)
    print("Use Case: Preview before replace")
    print("=" * 60)
    print("""
    Before using replace_tracked(), you can preview all matches:

    # Preview what will be replaced
    matches = doc.find_all("old text")
    print(f"Will replace {len(matches)} occurrence(s):")
    for match in matches:
        print(f"  - At {match.location}: {match.context}")

    # Now replace with confidence
    if len(matches) == 1:
        doc.replace_tracked("old text", "new text")
    else:
        print("Multiple matches found - need to be more specific!")
    """)


if __name__ == "__main__":
    main()
