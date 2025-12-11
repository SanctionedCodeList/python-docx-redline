# Eric White's Algorithm for Search and Replace in DOCX Files

**Source**: Eric White (Microsoft Open XML expert)
**Date Documented**: December 7, 2025

## Overview

Eric White developed a robust algorithm for finding and replacing text in Open XML WordprocessingML (DOCX) documents. This algorithm specifically addresses the challenge that text in Word documents is often fragmented across multiple `<w:r>` (run) elements with different formatting.

## The Problem

In DOCX files:
- Text may be split across multiple runs for various reasons (spelling, grammar, editing history/rsid)
- Simple string search fails when target text spans multiple runs
- Formatting must be preserved during replacement operations
- Example: "Hello World" might exist as:
  ```xml
  <w:r><w:t>Hello </w:t></w:r>
  <w:r><w:t>Wo</w:t></w:r>
  <w:r><w:t>rld</w:t></w:r>
  ```

## Eric White's Single-Character Runs Algorithm

### Core Steps

1. **Concatenate Text**
   - Concatenate all text in a paragraph into a single string
   - Search for the target string in the concatenated text
   - Identify match positions

2. **Break into Single-Character Runs**
   - Split all runs in the paragraph into individual single-character runs
   - Each run contains exactly one character
   - Special characters (tabs, breaks, non-breaking hyphens) get separate runs
   - This normalization allows precise character-level matching

3. **Pattern Matching**
   - Iterate through single-character runs
   - Match sequences that correspond to the search string

4. **Replace Matched Runs**
   - Create a new run containing the replacement text
   - **Inherit formatting from the first run** in the matched sequence
   - Delete all single-character runs that were matched

5. **Coalesce Adjacent Runs**
   - Merge adjacent runs with identical formatting back into single runs
   - This restores clean, efficient markup
   - The document becomes readable and maintains optimal file size

### Pseudocode

```python
def search_and_replace(paragraph, search_text, replace_text):
    # Step 1: Concatenate and search
    full_text = concatenate_all_run_text(paragraph)
    matches = find_all_matches(full_text, search_text)

    if not matches:
        return paragraph

    # Step 2: Break into single-character runs
    single_char_runs = break_into_single_character_runs(paragraph)

    # Step 3 & 4: Match and replace
    for match in matches:
        start_index = match.start
        end_index = match.end

        # Get formatting from first matched run
        first_run_format = single_char_runs[start_index].formatting

        # Create replacement run
        replacement_run = create_run(replace_text, first_run_format)

        # Replace matched runs
        remove_runs(single_char_runs[start_index:end_index])
        insert_run(replacement_run, start_index)

    # Step 5: Coalesce adjacent runs
    coalesced_runs = coalesce_runs_with_same_formatting(paragraph)

    return paragraph
```

## Implementation in C# (OpenXmlPowerTools)

Eric White's `OpenXmlRegex` utility class in the Open-Xml-PowerTools library implements this algorithm. The implementation:

- Uses XML DOM (not LINQ to XML) for broader platform compatibility
- Handles main document, headers, footers, endnotes, and footnotes
- Only processes paragraphs without tracked revisions (constraint)

## Run Splitting Implementation Details

### Determining Split Locations

```csharp
static int[] RunSplitLocations(XElement paragraph)
{
    // Find non-deleted runs
    var runElements = paragraph
        .Descendants(W.r)
        .Where(e => e.Parent.Name != W.del &&
                    e.Parent.Name != W.moveFrom &&
                    e.Descendants(W.t).Any());

    // Calculate cumulative positions
    var runs = runElements.Select(r => new {
        RunElement = r,
        RunLength = GetRunLength(r)
    });

    var runSplits = runs
        .Select(r => runs
            .TakeWhile(a => a.RunElement != r.RunElement)
            .Select(z => z.RunLength)
            .Sum());

    return runSplits.ToArray();
}
```

**Example**: For runs ["abc", "def", "ghi"], returns `[0, 3, 6]`

### Splitting Runs at Positions

The algorithm:
1. Calculates existing run boundaries
2. Merges desired split positions with existing boundaries
3. Recursively transforms the paragraph XML
4. Preserves all formatting, properties, and nested structures

```csharp
public static XElement SplitRunsInParagraph(XElement p, int[] positions)
{
    // Get existing split locations
    var runSplits = CalculateRunSplits(p);

    // Union of existing + desired splits
    int[] allSplits = runSplits
        .Select(rs => rs.RunLocation)
        .Concat(positions)
        .OrderBy(s => s)
        .Distinct()
        .ToArray();

    // Transform paragraph
    return new XElement(W.p,
        p.Elements().Select(e =>
            RunTransform(e, allSplits, runSplits))
    );
}
```

## Key Features

### Advantages
- **Handles fragmentation**: Works regardless of how text is split
- **Preserves formatting**: Inherits properties from matched runs
- **Clean output**: Coalescing prevents run proliferation
- **Comprehensive**: Processes entire document including headers/footers

### Constraints
- Cannot process paragraphs with tracked revisions (w:del, w:ins)
- Requires separate handling for tracked changes scenarios

## Application to Our Project

### Current Approach (python_docx_redline)

Our `TextSearch` class uses a **character map** approach:
```python
# Build character map: char_index -> (run_index, offset_in_run)
char_map = []
for run_idx, run in enumerate(runs):
    run_text = "".join(run.itertext())
    for char_idx, char in enumerate(run_text):
        char_map.append((run_idx, char_idx))
```

This is **more efficient** than Eric White's approach for read-only search because:
- No document modification required
- No run splitting/coalescing overhead
- Direct mapping from found position back to original runs

### When to Use Eric White's Algorithm

Consider Eric White's approach when:
1. **Replacing text that spans multiple runs** with different formatting
2. **Complex transformations** requiring precise character-level control
3. **Multiple sequential operations** where normalization improves reliability
4. **Simplifying complex run structures** before other operations

### Hybrid Approach

For our tracked changes use case:
1. Use **character map** for searching (fast, read-only)
2. Use **targeted run splitting** only where needed for replacement
3. Skip full normalization to preserve document history
4. Only coalesce runs we create (not existing ones)

## References

- [Search and Replace Text in an Open XML WordprocessingML Document](http://www.ericwhite.com/blog/search-and-replace-text-in-an-open-xml-wordprocessingml-document/) - Eric White's Blog
- [Splitting Runs in Open XML Word Processing Document Paragraphs](https://learn.microsoft.com/en-us/archive/blogs/ericwhite/splitting-runs-in-open-xml-word-processing-document-paragraphs) - Microsoft Learn
- [Open-Xml-PowerTools](https://github.com/OfficeDev/Open-Xml-PowerTools) - C# implementation
- [MS-Word breaking text into character runs](https://stackoverflow.com/questions/35716105/ms-word-breaking-text-into-character-runs) - Stack Overflow discussion

## See Also

- `text_search.py` - Our character map implementation
- `IMPLEMENTATION_NOTES.md` - Text search algorithm details
- `tracked_xml.py` - XML generation for tracked changes
