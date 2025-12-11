# Issue: CustomXml Case Sensitivity Validation Errors

**STATUS: RESOLVED** - Case-insensitive path matching added to `validate_file_references()` in `validation_base.py`.

## Summary

Documents with mixed-case `customXml` vs `customXML` directory references fail validation with broken reference errors, even though the files exist. This prevents any tracked changes from being saved.

## Reproduction

1. Open a document that was created/edited in Microsoft 365 Word
2. The document may contain both `customXml/` (lowercase) and `customXML/` (uppercase) paths
3. Attempt any tracked change operation
4. Validation fails with errors like:

```
FAILED - Found 4 relationship validation errors:
  customXml/_rels/item5.xml.rels: Line 1: Broken reference to /customXML/itemProps5.xml
  word/_rels/document.xml.rels: Line 1: Broken reference to /customXML/item5.xml
  Unreferenced file: customXml/item5.xml
  Unreferenced file: customXml/itemProps5.xml
```

## Root Cause

The ZIP file system in OOXML is case-sensitive, but Word sometimes creates files with inconsistent casing:
- Existing items at `customXml/item1.xml` (lowercase)
- New items at `customXML/item5.xml` (uppercase)
- References point to `/customXML/` but files may be at `/customXml/`

## Expected Behavior

The library should either:
1. Normalize case when validating references (treat paths as case-insensitive)
2. Provide a `repair()` method to fix common document corruptions before editing
3. At minimum, provide a clear error message explaining the issue and suggesting manual repair

## Workaround

Manually extract the docx, fix the casing, and repackage:

```bash
unzip document.docx -d /tmp/repair
cd /tmp/repair
# Move uppercase to lowercase
mv customXML/* customXml/
rmdir customXML
# Fix references
sed -i 's|/customXML/|/customXml/|g' word/_rels/document.xml.rels
sed -i 's|/customXML/|/customXml/|g' customXml/_rels/*.rels
# Repackage
zip -r ../fixed.docx .
```

## Suggested Fix

Add case-insensitive path matching in the validation logic, or add a `Document.repair()` method that normalizes common issues before editing.

## Environment

- Source document: Microsoft 365 Word (web or desktop)
- python_docx_redline version: current development
- Date: 2024-12-08

## Resolution

The `validate_file_references()` method in `validation_base.py` was updated to use case-insensitive path matching:

1. Builds a lookup table mapping lowercase paths to actual file paths
2. When checking if a referenced file exists, uses case-insensitive lookup
3. Works correctly on both case-sensitive (Linux) and case-insensitive (macOS/Windows) filesystems

Test added: `test_validation_case_insensitive_customxml_paths` in `tests/test_ooxml_validation.py`
