"""
Example: Advanced scope filtering techniques.

This example demonstrates various ways to limit edits to specific
parts of a document using scope filters.
"""

from typing import Any

from docx_redline import Document


def example_simple_scope() -> None:
    """Example 1: Simple string scope (paragraphs containing text)."""
    doc = Document("contract.docx")

    # Only modify paragraphs that contain "payment"
    doc.replace_tracked(
        find="30 days",
        replace="45 days",
        scope="payment",  # Matches any paragraph containing "payment"
    )

    doc.save("output1.docx")
    print("✓ Example 1: Simple scope completed")


def example_section_scope() -> None:
    """Example 2: Section-based scope."""
    doc = Document("contract.docx")

    # Only modify text in the "Termination" section
    doc.insert_tracked(
        text=" without cause",
        after="may terminate this agreement",
        scope="section:Termination",  # Only in Termination section
    )

    doc.save("output2.docx")
    print("✓ Example 2: Section scope completed")


def example_explicit_paragraph_scope() -> None:
    """Example 3: Explicit paragraph_containing scope."""
    doc = Document("contract.docx")

    # Be explicit about paragraph filtering
    doc.delete_tracked(
        text="at any time",
        scope="paragraph_containing:liability",  # Explicit syntax
    )

    doc.save("output3.docx")
    print("✓ Example 3: Explicit paragraph scope completed")


def example_dictionary_scope_and() -> None:
    """Example 4: Dictionary scope with AND logic."""
    doc = Document("contract.docx")

    # Must contain "payment" AND "invoice" (both conditions)
    doc.replace_tracked(
        find="USD",
        replace="EUR",
        scope={
            "contains": "payment",
            "and_contains": "invoice",
        },
    )

    doc.save("output4.docx")
    print("✓ Example 4: Dictionary AND scope completed")


def example_dictionary_scope_not() -> None:
    """Example 5: Dictionary scope with NOT logic."""
    doc = Document("contract.docx")

    # Contains "confidential" but NOT "Exhibit"
    doc.insert_tracked(
        text=" (as amended)",
        after="agreement",
        scope={
            "contains": "confidential",
            "not_contains": "Exhibit",
        },
    )

    doc.save("output5.docx")
    print("✓ Example 5: Dictionary NOT scope completed")


def example_dictionary_scope_section_filter() -> None:
    """Example 6: Combining section and text filters."""
    doc = Document("contract.docx")

    # In "Definitions" section AND contains "means"
    doc.replace_tracked(
        find="shall mean",
        replace="means",
        scope={
            "section": "Definitions",
            "contains": "means",
        },
    )

    doc.save("output6.docx")
    print("✓ Example 6: Section + text filter completed")


def example_callable_scope() -> None:
    """Example 7: Custom callable scope for complex logic."""
    doc = Document("contract.docx")

    # Custom function: only paragraphs longer than 100 characters
    from typing import Any

    def long_paragraphs(paragraph: Any) -> bool:
        text = "".join(paragraph.itertext())
        return len(text) > 100

    doc.insert_tracked(
        text=" (detailed provision)",
        after="agreement",
        scope=long_paragraphs,
    )

    doc.save("output7.docx")
    print("✓ Example 7: Callable scope completed")


def example_regex_with_scope() -> None:
    """Example 8: Combining regex with scope filters."""
    doc = Document("contract.docx")

    # Use regex to find dates, but only in Payment section
    doc.replace_tracked(
        find=r"(\d{2})/(\d{2})/(\d{4})",  # MM/DD/YYYY
        replace=r"\2/\1/\3",  # DD/MM/YYYY
        regex=True,
        scope="section:Payment Terms",
    )

    doc.save("output8.docx")
    print("✓ Example 8: Regex + scope completed")


def example_multiple_scopes() -> None:
    """Example 9: Applying different scopes to different edits."""
    doc = Document("contract.docx")

    edits: list[dict[str, Any]] = [
        # Edit 1: Only in Definitions
        {
            "type": "replace_tracked",
            "find": "shall mean",
            "replace": "means",
            "scope": "section:Definitions",
        },
        # Edit 2: Only paragraphs with "payment"
        {
            "type": "replace_tracked",
            "find": "30 days",
            "replace": "45 days",
            "scope": "payment",
        },
        # Edit 3: Exclude confidential sections
        {
            "type": "insert_tracked",
            "text": " (public version)",
            "after": "This agreement",
            "scope": {"not_contains": "confidential"},
        },
    ]

    results = doc.apply_edits(edits)

    # Show which edits were applied
    for i, result in enumerate(results, 1):
        if result.success:
            print(f"  ✓ Edit {i}: {result.message}")
        else:
            print(f"  ✗ Edit {i}: {result.message}")

    doc.save("output9.docx")
    print("✓ Example 9: Multiple scopes completed")


def example_scope_yaml() -> None:
    """Example 10: Complex scopes in YAML configuration."""
    yaml_content = """
edits:
  # Simple scope
  - type: replace_tracked
    find: "Contractor"
    replace: "Service Provider"
    scope: "section:Parties"

  # Dictionary scope with multiple conditions
  - type: insert_tracked
    text: " (as defined above)"
    after: "Services"
    scope:
      section: "Scope of Work"
      contains: "Services"
      not_contains: "Exhibit"

  # Paragraph-based scope
  - type: delete_tracked
    text: "unless otherwise agreed"
    scope: "paragraph_containing:termination"
"""

    # Save YAML to file
    with open("scope_edits.yaml", "w") as f:
        f.write(yaml_content)

    # Apply from YAML
    doc = Document("contract.docx")
    results = doc.apply_edit_file("scope_edits.yaml")

    print(f"✓ Example 10: Applied {len(results)} edits from YAML with scopes")
    doc.save("output10.docx")


def main() -> None:
    """Run all scope examples."""
    print("Advanced Scope Filtering Examples")
    print("=" * 60)

    examples = [
        ("Simple string scope", example_simple_scope),
        ("Section-based scope", example_section_scope),
        ("Explicit paragraph scope", example_explicit_paragraph_scope),
        ("Dictionary AND logic", example_dictionary_scope_and),
        ("Dictionary NOT logic", example_dictionary_scope_not),
        ("Section + text filter", example_dictionary_scope_section_filter),
        ("Custom callable scope", example_callable_scope),
        ("Regex + scope", example_regex_with_scope),
        ("Multiple different scopes", example_multiple_scopes),
        ("Scopes in YAML", example_scope_yaml),
    ]

    for name, func in examples:
        print(f"\n{name}:")
        print("-" * 60)
        try:
            func()
        except Exception as e:
            print(f"  ✗ Error: {e}")

    print("\n" + "=" * 60)
    print("All examples completed!")


if __name__ == "__main__":
    main()
