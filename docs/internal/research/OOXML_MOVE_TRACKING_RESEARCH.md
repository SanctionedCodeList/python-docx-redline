# OOXML Move Tracking in Word Documents: Comprehensive Research Report

**Research Date:** December 8, 2025
**Subject:** XML structure and implementation details for tracking text moves in Office Open XML (OOXML) WordprocessingML documents

---

## Executive Summary

This report provides comprehensive documentation on how Microsoft Word tracks moved text using OOXML (Office Open XML) format as defined in the ECMA-376 and ISO/IEC 29500 standards. Move tracking is distinct from simple delete+insert operations, preserving semantic relationships between source and destination locations to provide better context for document reviewers.

**Key Findings:**
- Move operations require **paired container elements** at both source and destination locations
- Move tracking uses **six primary XML elements** working in concert
- **Move IDs link containers** while **move names link source to destination**
- Move tracking differs fundamentally from delete+insert by preserving relocation semantics
- Incomplete or improperly linked move containers are treated as simple insertions or deletions

---

## 1. The XML Structure for w:moveFrom and w:moveTo Elements

### 1.1 Move Source Elements (w:moveFrom)

The `<w:moveFrom>` element marks inline-level content that has been moved away from its current location and tracked as a revision.[^1]

#### XML Structure

```xml
<w:moveFromRangeStart w:id="2" w:name="move1" w:author="Editor" w:date="2006-01-01T10:00:00"/>
<w:moveFrom w:id="3" w:author="Editor" w:date="2006-01-01T10:00:00">
  <w:r>
    <w:t>Some moved text.</w:t>
  </w:r>
</w:moveFrom>
<w:moveFromRangeEnd w:id="2"/>
```

#### Attributes

| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `w:id` | ST_DecimalNumber | Yes | Unique annotation identifier for the revision element itself[^2] |
| `w:author` | ST_String | No | Creator of the revision[^3] |
| `w:date` | ST_DateTime | No | Timestamp in ISO 8601 format[^3] |

#### Schema Definition

The element is of type `CT_RunTrackChange`:[^4]

```xml
<complexType name="CT_RunTrackChange">
  <complexContent>
    <extension base="CT_TrackChange">
      <choice minOccurs="0" maxOccurs="unbounded">
        <group ref="EG_ContentRunContent"/>
        <group ref="m:EG_OMathMathElements"/>
      </choice>
    </extension>
  </complexContent>
</complexType>
```

#### Child Elements

Permissible children include:[^5]
- Text runs (`<w:r>`)
- Bookmarks and comments (`<w:bookmarkStart>`, `<w:commentRangeStart>`)
- Other revisions (`<w:ins>`, `<w:del>`)
- Mathematical elements (`<w:oMath>`, `<w:oMathPara>`)
- Smart tags and structured data
- Custom XML elements

#### Critical Conformance Requirement

**If this element occurs outside of a valid move source container for which a matching move destination container exists in the document, then content in this region shall be treated as deleted, rather than moved.**[^6]

### 1.2 Move Destination Elements (w:moveTo)

The `<w:moveTo>` element marks inline-level content that has been moved to its current location and tracked as a revision.[^7]

#### XML Structure

```xml
<w:moveToRangeStart w:id="0" w:name="move1" w:author="Editor" w:date="2006-01-01T10:00:00"/>
<w:moveTo w:id="1" w:author="Editor" w:date="2006-01-01T10:00:00">
  <w:r>
    <w:t>Some moved text.</w:t>
  </w:r>
</w:moveTo>
<w:moveToRangeEnd w:id="0"/>
```

#### Attributes

The `w:moveTo` element uses the same attribute set as `w:moveFrom`:[^8]

| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `w:id` | ST_DecimalNumber | Yes | Unique annotation identifier |
| `w:author` | ST_String | No | Creator of the revision |
| `w:date` | ST_DateTime | No | Timestamp in ISO 8601 format |

#### Critical Conformance Requirement

**If this element occurs outside of a valid move destination container for which a matching move source container exists in the document, then content in this region shall be treated as inserted, rather than moved.**[^9]

---

## 2. How Move IDs Link the Source and Destination

Move tracking in OOXML uses a **two-tier linking system** involving both IDs and names:

### 2.1 Container IDs (w:id on Range Elements)

The `w:id` attribute on `moveFromRangeStart` and `moveToRangeStart` elements serves to link each start element with its corresponding end element.[^10]

**Example:**
```xml
<!-- Links start to end within the same container -->
<w:moveFromRangeStart w:id="2" w:name="move1"/>
  <!-- content -->
<w:moveFromRangeEnd w:id="2"/>
```

**Conformance Requirements:**[^11]
- Each `moveFromRangeStart` **must** have a corresponding `moveFromRangeEnd` with matching `w:id`
- Each `moveToRangeStart` **must** have a corresponding `moveToRangeEnd` with matching `w:id`
- Multiple start elements with the same `w:id` make the document non-conformant
- Documents without matching end elements are non-conformant

### 2.2 Move Names (w:name Attribute)

The `w:name` attribute on both `moveFromRangeStart` and `moveToRangeStart` elements links the move source container with the corresponding move destination container.[^12]

**Example:**
```xml
<!-- Source location -->
<w:moveFromRangeStart w:id="2" w:name="move1"/>
<w:moveFrom w:id="3">
  <w:r><w:t>Some moved text.</w:t></w:r>
</w:moveFrom>
<w:moveFromRangeEnd w:id="2"/>

<!-- Destination location (elsewhere in document) -->
<w:moveToRangeStart w:id="0" w:name="move1"/>
<w:moveTo w:id="1">
  <w:r><w:t>Some moved text.</w:t></w:r>
</w:moveTo>
<w:moveToRangeEnd w:id="0"/>
```

Both containers share `w:name="move1"` to establish they represent the same move operation.[^13]

### 2.3 Annotation IDs (w:id on moveFrom/moveTo)

The `w:id` attributes on the `<w:moveFrom>` and `<w:moveTo>` elements themselves are unique annotation identifiers but **do not directly link source to destination**. The linking is accomplished through the `w:name` attribute on the container elements.[^14]

---

## 3. Complete Move Operation in document.xml

A complete move operation requires the following elements in `word/document.xml`:

### 3.1 Complete XML Example

```xml
<w:p>
  <!-- MOVE DESTINATION -->
  <w:moveToRangeStart w:id="0" w:author="Editor" w:date="2006-01-01T10:00:00" w:name="move1"/>
  <w:moveTo w:id="1" w:author="Editor" w:date="2006-01-01T10:00:00">
    <w:r>
      <w:t>Some moved text.</w:t>
    </w:r>
  </w:moveTo>
  <w:moveToRangeEnd w:id="0"/>

  <!-- NORMAL CONTENT -->
  <w:r>
    <w:t xml:space="preserve">Some text.</w:t>
  </w:r>

  <!-- MOVE SOURCE -->
  <w:moveFromRangeStart w:id="2" w:author="Editor" w:date="2006-01-01T10:00:00" w:name="move1"/>
  <w:moveFrom w:id="3" w:author="Editor" w:date="2006-01-01T10:00:00">
    <w:r>
      <w:t>Some moved text.</w:t>
    </w:r>
  </w:moveFrom>
  <w:moveFromRangeEnd w:id="2"/>
</w:p>
```

This example shows "Some moved text." being moved from the end of the paragraph to the beginning.[^15]

### 3.2 Required Elements Checklist

For a valid move operation, you need:[^16]

**At the Source Location:**
1. `<w:moveFromRangeStart>` with unique `w:id` and `w:name`
2. `<w:moveFrom>` with unique `w:id` (contains the moved content)
3. `<w:moveFromRangeEnd>` with matching `w:id` to the start element

**At the Destination Location:**
4. `<w:moveToRangeStart>` with unique `w:id` and matching `w:name` from source
5. `<w:moveTo>` with unique `w:id` (contains the moved content)
6. `<w:moveToRangeEnd>` with matching `w:id` to the start element

### 3.3 Track Revisions Must Be Enabled

For move tracking to be properly recorded, the document must have track revisions enabled in `word/settings.xml`:[^17]

```xml
<w:settings>
  <w:trackRevisions w:val="true"/>
</w:settings>
```

---

## 4. Related Bookmarks and Markers

### 4.1 moveFromRangeStart Element

The `<w:moveFromRangeStart>` element specifies the beginning of a move source container.[^18]

#### Purpose

When a move source is stored as a revision in a WordprocessingML document, two pieces of information are stored:[^19]
1. A set of pieces of content which were moved (both inline-level content and paragraphs)
2. A move source container (or "bookmark") which specifies that all content within it marked as a move source is part of a single named move

#### Attributes

| Attribute | Description |
|-----------|-------------|
| `w:id` | Links this start element with corresponding `moveFromRangeEnd` (required)[^20] |
| `w:name` | Named identifier linking move source with move destination container (required)[^21] |
| `w:author` | Author of the tracked change[^22] |
| `w:date` | Date of the tracked change[^22] |
| `w:colFirst` | First column reference (for table moves)[^22] |
| `w:colLast` | Last column reference (for table moves)[^22] |
| `w:displacedByCustomXml` | Custom XML displacement flag[^22] |

#### Conformance Requirements

From ISO/IEC 29500-1:[^23]

1. **Must have a corresponding `moveFromRangeEnd` element** with matching `w:id`
2. **Must have matching move destination container** elements
3. **Cannot have multiple start elements** with the same `w:id`
4. **Cannot surround the same text** with multiple move source containers

If this element occurs without a corresponding `moveFromRangeEnd` element with a matching `w:id` attribute value, then it shall be ignored.[^24]

### 4.2 moveFromRangeEnd Element

The `<w:moveFromRangeEnd>` element specifies the end of the move source container within which all moved content is part of the named move.[^25]

#### Structure

```xml
<w:moveFromRangeEnd w:id="2"/>
```

The element is typically empty and only requires the `w:id` attribute to match its corresponding start element.[^26]

### 4.3 moveToRangeStart Element

The `<w:moveToRangeStart>` element marks the beginning of a move destination container.[^27]

#### Structure

```xml
<w:moveToRangeStart w:id="0" w:name="move1" w:author="Editor" w:date="2006-01-01T10:00:00"/>
```

#### Conformance Requirements

- Must have a corresponding `moveToRangeEnd` element with matching `w:id`[^28]
- Must have a corresponding move source container with matching `w:name`[^28]
- Multiple start elements with the same `w:id` make the document non-conformant[^28]
- Multiple move destination containers surrounding the same text make the document non-conformant[^28]

### 4.4 moveToRangeEnd Element

The `<w:moveToRangeEnd>` element specifies the end of a region whose move destination contents are part of a single named move.[^29]

#### Structure

```xml
<w:moveToRangeEnd w:id="0"/>
```

### 4.5 Paragraph-Level Move Markers

For tracking moves of entire paragraphs (including the paragraph mark), OOXML provides additional elements that appear within paragraph properties:[^30]

#### Move Destination Paragraph

```xml
<w:moveToRangeStart w:id="0" w:name="aMove"/>
<w:p>
  <w:pPr>
    <w:rPr>
      <w:moveTo w:id="1" w:author="Editor" w:date="2006-01-01T10:00:00"/>
    </w:rPr>
  </w:pPr>
  <!-- paragraph content -->
</w:p>
<w:moveToRangeEnd w:id="0"/>
```

This marks the **paragraph mark itself** as moved to this location. The run-level content within the paragraph requires separate `<w:moveTo>` elements.[^31]

#### Move Source Paragraph

```xml
<w:moveFromRangeStart w:id="2" w:name="aMove"/>
<w:p>
  <w:pPr>
    <w:rPr>
      <w:moveFrom w:id="3" w:author="Editor" w:date="2006-01-01T10:00:00"/>
    </w:rPr>
  </w:pPr>
  <!-- paragraph content -->
</w:p>
<w:moveFromRangeEnd w:id="2"/>
```

This marks the **paragraph mark** as moved away from this location.[^32]

---

## 5. How Move Tracking Differs from Simple Delete+Insert

### 5.1 Semantic Distinction

The fundamental difference between move tracking and delete+insert operations:[^33]

| Aspect | Move Tracking | Delete + Insert |
|--------|---------------|-----------------|
| **Elements Used** | `w:moveFrom` / `w:moveTo` with container ranges | `w:del` / `w:ins` |
| **Semantic Meaning** | Explicitly marks content as **relocated** | Two **separate, unrelated** changes |
| **Linking** | Source and destination **linked via w:name** | No connection between operations |
| **User Experience** | Shows content was **moved** (not deleted/re-added) | Shows as independent deletion and insertion |
| **Context Preservation** | Maintains **relationship** between locations | No relationship preserved |
| **Container Requirements** | **Requires paired containers** at both locations | Single container per operation |

### 5.2 Why This Distinction Matters

Move tracking allows applications to:[^34]

1. **Show users that content was moved** rather than deleted and re-added, providing better context for document reviewers
2. **Preserve the semantic relationship** between source and destination positions
3. **Enable better merge strategies** in version control scenarios
4. **Improve review workflows** by clearly indicating reorganization vs. content changes

### 5.3 Delete+Insert Example (for comparison)

```xml
<!-- Deletion at original location -->
<w:del w:id="1" w:author="Editor" w:date="2006-01-01T10:00:00">
  <w:r>
    <w:delText>Some deleted text.</w:delText>
  </w:r>
</w:del>

<!-- Insertion at new location (elsewhere in document) -->
<w:ins w:id="2" w:author="Editor" w:date="2006-01-01T10:00:00">
  <w:r>
    <w:t>Some deleted text.</w:t>
  </w:r>
</w:ins>
```

Note that there is **no linking mechanism** between these two operations - they are treated as completely independent changes.[^35]

### 5.4 Fallback Behavior

OOXML specifies fallback behavior when move containers are incomplete:[^36]

- **Move source without matching destination**: Content is treated as **deleted**
- **Move destination without matching source**: Content is treated as **inserted**

This ensures graceful degradation if the document structure is invalid or if the consumer doesn't support move tracking.

### 5.5 Storage Approach: OOXML vs. ODF

An important implementation difference exists between OOXML and ODF (Open Document Format):[^37]

- **OOXML approach**: Deleted content remains **in its original location**, marked with revision tags
- **ODF approach**: Deleted content is moved to a **separate location** in the document

The OOXML approach is simpler for processing because content stays in place, whereas ODF requires retrieving and assembling content from different locations.

---

## 6. Additional Technical Specifications

### 6.1 Tracked Revision Elements Overview

Tracked revisions are one of the more involved features of Open XML WordprocessingML. There are **28 elements** associated with tracked revisions, each with their own semantics.[^38]

#### Move-Related Elements

The XPath expression to detect move-related tracked revisions includes:[^39]
- `w:moveFrom`
- `w:moveFromRangeEnd`
- `w:moveFromRangeStart`
- `w:moveTo`
- `w:moveToRangeEnd`
- `w:moveToRangeStart`

### 6.2 Parent Elements

Both `w:moveFrom` and `w:moveTo` can appear within various parent elements:[^40]

**Block containers:**
- `<w:body>`
- `<w:p>` (paragraph)
- `<w:tbl>` (table)

**Inline contexts:**
- `<w:r>` (run)
- `<w:hyperlink>`

**Structured elements:**
- `<w:comment>`
- `<w:footnote>`
- `<w:endnote>`
- `<w:ftr>` (footer)
- `<w:hdr>` (header)

**Custom markup:**
- `<w:customXml>`
- `<w:sdt>` (structured document tag)

### 6.3 Multiple Move Operations

When multiple move operations exist:[^41]

- Each must have a **unique `w:name`** to distinguish different move operations
- Container `w:id` values must be unique across all range markers
- If multiple start elements exist with the same `w:id`, they are matched with ends **in document order**, and unmatched starts are handled by being ignored

### 6.4 Nested and Overlapping Moves

**If multiple move source containers surround the same text:**[^42]
- The **last valid container** (determined by the location of the container start elements, in document order) should be the container associated with that text
- The document is technically non-conformant in this case

### 6.5 Custom XML and Move Tracking

OOXML also supports tracking moves of custom XML elements separately from text content:[^43]

- `<w:customXmlMoveFromRangeStart>` / `<w:customXmlMoveFromRangeEnd>`
- `<w:customXmlMoveToRangeStart>` / `<w:customXmlMoveToRangeEnd>`

These elements track when custom XML markup has been moved, which affects the custom XML structure but not the visible text content.

### 6.6 Table Move Tracking

Move tracking can also apply to table content. The `w:colFirst` and `w:colLast` attributes on the range start elements define table column ranges for bookmarks when moving table cells or columns.[^44]

---

## 7. Implementation Guidelines

### 7.1 Creating a Move Operation

To implement move tracking in code:

1. **Generate unique IDs** for all range markers and revision elements
2. **Choose a unique move name** (e.g., "move1", "aMove") to link source and destination
3. **Create the destination container** first (optional, but can help with document order)
   - Insert `<w:moveToRangeStart>` with unique `w:id` and move `w:name`
   - Insert `<w:moveTo>` with unique `w:id` containing the moved content
   - Insert `<w:moveToRangeEnd>` with matching `w:id`
4. **Create the source container** at the original location
   - Insert `<w:moveFromRangeStart>` with unique `w:id` and matching move `w:name`
   - Insert `<w:moveFrom>` with unique `w:id` containing the same text content
   - Insert `<w:moveFromRangeEnd>` with matching `w:id`
5. **Set author and date** attributes consistently across both source and destination elements
6. **Ensure track revisions is enabled** in `word/settings.xml`

### 7.2 ID Management

All `w:id` values should be unique across the entire document for proper tracking:[^45]

```python
# Example ID generation strategy
next_id = 0

def get_next_id():
    global next_id
    next_id += 1
    return str(next_id)

# Usage
move_to_range_id = get_next_id()      # "1"
move_to_id = get_next_id()            # "2"
move_from_range_id = get_next_id()    # "3"
move_from_id = get_next_id()          # "4"
```

### 7.3 Name Management

Move names link source to destination and should be descriptive:[^46]

```python
move_counter = 0

def generate_move_name():
    global move_counter
    move_counter += 1
    return f"move{move_counter}"

# Usage
move_name = generate_move_name()  # "move1"
# Use same move_name for both source and destination containers
```

### 7.4 Validation Checklist

Before writing a document with move tracking:

- [ ] Each `moveFromRangeStart` has a matching `moveFromRangeEnd` with same `w:id`
- [ ] Each `moveToRangeStart` has a matching `moveToRangeEnd` with same `w:id`
- [ ] Move source and destination containers share the same `w:name`
- [ ] All `w:id` values are unique across the document
- [ ] Track revisions is enabled in settings
- [ ] Author and date attributes are set consistently
- [ ] Content in `moveFrom` and `moveTo` elements is identical
- [ ] No multiple move containers surround the same text

---

## 8. References and Standards

### 8.1 Primary Standards

- **ECMA-376**: Office Open XML File Formats[^47]
  - Available at: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- **ISO/IEC 29500**: Information technology — Document description and processing languages — Office Open XML File Formats[^48]

### 8.2 Microsoft Implementation Notes

- **[MS-OE376]**: Office Implementation Information for ECMA-376 Standards Support[^49]
  - Contains Microsoft-specific implementation details and extensions

### 8.3 Key Documentation Resources

- **OOXML Documentation** (c-rex.net): Detailed element-by-element reference[^50]
- **Microsoft Learn**: Open XML SDK documentation and examples[^51]
- **Eric White's Blog**: WordprocessingML implementation guidance[^52]

---

## 9. Conclusion

Move tracking in OOXML is a sophisticated mechanism that requires coordinated pairs of container elements at both source and destination locations, linked by shared move names. This system preserves the semantic relationship between moved content locations, providing superior context for document reviewers compared to simple delete+insert operations.

Key implementation points:
- Always create **complete container pairs** at both locations
- Use **unique IDs** for range markers and revision elements
- Use **matching names** to link source and destination
- Set **author and date** attributes consistently
- Validate that **track revisions is enabled**
- Handle **fallback cases** where moves degrade to insertions/deletions

This research provides the foundation for implementing move tracking functionality in the `python_docx_redline` Python library.

---

## Footnotes

[^1]: [moveFrom (Move Source Run Content)](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFrom_topic_ID0EPJCW.html)

[^2]: [moveFrom (Move Source Run Content) - Attributes](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFrom_topic_ID0EPJCW.html)

[^3]: [moveTo (Move Destination Run Content) - Attributes](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html)

[^4]: [moveTo (Move Destination Run Content) - Schema](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html)

[^5]: [moveTo (Move Destination Run Content) - Child Elements](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html)

[^6]: [moveFrom (Move Source Run Content) - Conformance](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFrom_topic_ID0EPJCW.html)

[^7]: [moveTo (Move Destination Run Content)](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html)

[^8]: [moveTo (Move Destination Run Content) - Attributes](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html)

[^9]: [moveTo (Move Destination Run Content) - Conformance](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html)

[^10]: [MoveFromRangeStart Class - Microsoft Learn](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefromrangestart?view=openxml-2.8.1)

[^11]: [moveFromRangeStart (Move Source Location Container - Start)](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFromRangeStart_topic_ID0EOYGW.html)

[^12]: [MoveFromRangeStart Class - Linking Attributes](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefromrangestart?view=openxml-2.8.1)

[^13]: [MoveFromRangeStart Class - Code Example](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefromrangestart?view=openxml-2.8.1)

[^14]: [moveFrom (Move Source Run Content) - ID Attribute](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFrom_topic_ID0EPJCW.html)

[^15]: [MoveFromRangeStart Class - Complete Example](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefromrangestart?view=openxml-2.8.1)

[^16]: [moveFromRangeStart (Move Source Location Container - Start) - Requirements](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFromRangeStart_topic_ID0EOYGW.html)

[^17]: [trackRevisions (Track Revisions to Document)](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_trackRevisions_topic_ID0EKXKY.html)

[^18]: [moveFromRangeStart (Move Source Location Container - Start)](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFromRangeStart_topic_ID0EOYGW.html)

[^19]: [MoveFromRangeStart Class - Purpose](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefromrangestart?view=openxml-2.8.1)

[^20]: [moveFromRangeStart - ID Attribute](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFromRangeStart_topic_ID0EOYGW.html)

[^21]: [moveFromRangeStart - Name Attribute](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFromRangeStart_topic_ID0EOYGW.html)

[^22]: [MoveFromRangeStart Class - Attributes](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefromrangestart?view=openxml-2.8.1)

[^23]: [MoveFromRangeStart Class - Conformance Requirements](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefromrangestart?view=openxml-2.8.1)

[^24]: [moveFromRangeStart - Missing End Element](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFromRangeStart_topic_ID0EOYGW.html)

[^25]: [moveFromRangeEnd (Move Source Location Container - End)](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveFromRangeEnd_topic_ID0E3PFW.html)

[^26]: [moveFromRangeEnd - Structure](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveFromRangeEnd_topic_ID0E3PFW.html)

[^27]: [MoveToRangeStart Class - Microsoft Learn](https://learn.microsoft.com/ru-ru/dotnet/api/documentformat.openxml.wordprocessing.movetorangestart?view=openxml-2.8.1)

[^28]: [moveToRangeEnd (Move Destination Location Container - End) - Conformance](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveToRangeEnd_topic_ID0ERCMW.html)

[^29]: [MoveToRangeEnd Class - Microsoft Learn](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movetorangeend?view=openxml-3.0.1)

[^30]: [moveTo (Move Destination Paragraph)](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EE3IW.html)

[^31]: [moveTo (Move Destination Paragraph) - Implementation Notes](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EE3IW.html)

[^32]: [moveFrom (Move Source Paragraph)](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveFrom_topic_ID0EJ6EW.html)

[^33]: [Balisage: Standard Change Tracking for XML](https://www.balisage.net/Proceedings/vol13/html/LaFontaine01/BalisageVol13-LaFontaine01.html)

[^34]: [Tracked Changes - Microsoft Learn Blog](https://blogs.msdn.microsoft.com/dmahugh/2009/05/14/tracked-changes/)

[^35]: [Using XML DOM to Detect Tracked Revisions in an Open XML WordprocessingML Document](http://www.ericwhite.com/blog/using-xml-dom-to-detect-tracked-revisions-in-an-open-xml-wordprocessingml-document/)

[^36]: [moveTo (Move Destination Run Content) - Fallback Behavior](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html)

[^37]: [Tracked Changes - OOXML vs ODF Approach](https://blogs.msdn.microsoft.com/dmahugh/2009/05/14/tracked-changes/)

[^38]: [How to: Accept all revisions in a word processing document](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document)

[^39]: [Using XML DOM to Detect Tracked Revisions - XPath Elements](http://www.ericwhite.com/blog/using-xml-dom-to-detect-tracked-revisions-in-an-open-xml-wordprocessingml-document/)

[^40]: [moveTo (Move Destination Run Content) - Parent Elements](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html)

[^41]: [moveFromRangeStart - Multiple Instances](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFromRangeStart_topic_ID0EOYGW.html)

[^42]: [moveFromRangeStart - Nested Containers](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFromRangeStart_topic_ID0EOYGW.html)

[^43]: [customXmlMoveToRangeStart (Custom XML Markup Move Destination Location Start)](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_customXmlMoveToRange_topic_ID0EMPYV.html)

[^44]: [MoveFromRangeStart Class - Table Attributes](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefromrangestart?view=openxml-2.8.1)

[^45]: [moveFromRangeStart - ID Uniqueness](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_moveFromRangeStart_topic_ID0EOYGW.html)

[^46]: [MoveFromRangeStart Class - Name Attribute Usage](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefromrangestart?view=openxml-2.8.1)

[^47]: [ECMA-376 - Ecma International](https://ecma-international.org/publications-and-standards/standards/ecma-376/)

[^48]: [OOXML Format Family -- ISO/IEC 29500 and ECMA 376](https://www.loc.gov/preservation/digital/formats/fdd/fdd000395.shtml)

[^49]: [MS-OE376: Office Implementation Information for ECMA-376 Standards Support](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/db9b9b72-b10b-4e7e-844c-09f88c972219)

[^50]: [OOXML Documentation - c-rex.net](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html)

[^51]: [Microsoft Learn - Open XML Documentation](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document)

[^52]: [WordprocessingML - Eric White's Blog](http://www.ericwhite.com/blog/wordprocessingml-expanded/)
