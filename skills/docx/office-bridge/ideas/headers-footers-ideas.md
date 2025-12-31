---
name: headers-footers-ideas
description: "Roadmap ideas for header/footer helper functions and accessibility tree integration in the Office Bridge add-in."
---

# Headers and Footers: Roadmap Ideas

This document outlines future enhancements for header/footer support in the Office Bridge DocTree accessibility layer.

## Helper Functions to Build

### Core Read Functions

#### getHeaderText(context, sectionIndex?, type?)

Get plain text content from a header.

```typescript
interface GetHeaderTextResult {
  text: string;
  paragraphCount: number;
  sectionIndex: number;
  type: Word.HeaderFooterType;
}

// Usage
const result = await DocTree.getHeaderText(context);  // Primary header, first section
const result = await DocTree.getHeaderText(context, 0, "FirstPage");
```

#### getFooterText(context, sectionIndex?, type?)

Get plain text content from a footer.

```typescript
const result = await DocTree.getFooterText(context);
const result = await DocTree.getFooterText(context, 1, "EvenPages");
```

#### getAllHeaders(context)

Get all headers across all sections and types.

```typescript
interface HeaderInfo {
  sectionIndex: number;
  type: "Primary" | "FirstPage" | "EvenPages";
  text: string;
  paragraphCount: number;
  isEmpty: boolean;
}

interface GetAllHeadersResult {
  headers: HeaderInfo[];
  sectionCount: number;
}

const result = await DocTree.getAllHeaders(context);
// Returns all non-empty headers from all sections
```

#### getAllFooters(context)

Same structure as getAllHeaders for footers.

### Core Write Functions

#### setHeaderText(context, text, sectionIndex?, type?, options?)

Replace header content with simple text.

```typescript
interface SetHeaderOptions {
  alignment?: "left" | "center" | "right";
  bold?: boolean;
  fontSize?: number;
  preserveFormatting?: boolean;  // Try to match existing formatting
}

await DocTree.setHeaderText(context, "Company Name", 0, "Primary", {
  alignment: "center",
  bold: true
});
```

#### setFooterText(context, text, sectionIndex?, type?, options?)

Replace footer content with simple text.

```typescript
await DocTree.setFooterText(context, "Confidential - Do Not Distribute");
```

#### clearHeader(context, sectionIndex?, type?)

Clear header content.

```typescript
await DocTree.clearHeader(context);  // Clear primary header, first section
await DocTree.clearHeader(context, "all");  // Clear all headers in all sections
```

#### clearFooter(context, sectionIndex?, type?)

Clear footer content.

### Batch Operations

#### setAllHeaders(context, text, options?)

Set the same header text across all sections.

```typescript
await DocTree.setAllHeaders(context, "Project Proposal v2.0", {
  skipFirstPage: true,  // Don't modify FirstPage headers
  skipEvenPages: true   // Don't modify EvenPages headers
});
```

#### setAllFooters(context, text, options?)

Set the same footer text across all sections.

### Search and Replace

#### findInHeaders(context, searchText, options?)

Search for text within headers.

```typescript
interface HeaderSearchResult {
  matches: Array<{
    sectionIndex: number;
    type: Word.HeaderFooterType;
    text: string;
    position: number;
  }>;
  count: number;
}

const result = await DocTree.findInHeaders(context, "DRAFT");
```

#### replaceInHeaders(context, searchText, replaceText, options?)

Find and replace text within headers.

```typescript
await DocTree.replaceInHeaders(context, "2024", "2025");
```

#### findInFooters / replaceInFooters

Same APIs for footers.

## Accessibility Tree Integration

### Header/Footer Nodes

Include headers and footers in the accessibility tree when requested:

```typescript
const tree = await DocTree.buildTree(context, {
  includeHeaders: true,
  includeFooters: true
});
```

#### Proposed Node Structure

```yaml
content:
  - ref: "hdr:0/primary"
    role: header
    sectionIndex: 0
    type: "Primary"
    children:
      - ref: "hdr:0/primary/p:0"
        role: paragraph
        text: "Company Confidential"
        style:
          alignment: center
          bold: true

  - ref: "p:0"
    role: paragraph
    text: "Document content..."

  # ... body content ...

  - ref: "ftr:0/primary"
    role: footer
    sectionIndex: 0
    type: "Primary"
    children:
      - ref: "ftr:0/primary/p:0"
        role: paragraph
        text: "Page 1 of 10"
```

#### Ref Format

```
hdr:<sectionIndex>/<type>              # Header container
hdr:<sectionIndex>/<type>/p:<index>    # Paragraph in header
ftr:<sectionIndex>/<type>              # Footer container
ftr:<sectionIndex>/<type>/p:<index>    # Paragraph in footer
```

Where `<type>` is one of: `primary`, `firstpage`, `evenpages`

### View Mode Options

```typescript
interface ViewMode {
  includeBody?: boolean;      // Default: true
  includeHeaders?: boolean;   // Default: false
  includeFooters?: boolean;   // Default: false
  includeComments?: boolean;
  includeTrackedChanges?: boolean;
}
```

### Scope Support

Add scope filters for headers and footers:

```typescript
// String shortcuts
"headers"                    // All headers
"footers"                    // All footers
"header:primary"             // Primary headers only
"header:firstpage"           // First page headers only
"footer:evenpages"           // Even pages footers only

// Object format
{ scope: "headers", section: 0 }           // Headers in first section
{ scope: "footers", contains: "Page" }     // Footers containing "Page"
```

## Page Numbering Ideas

### insertPageNumber(context, location, format?)

Insert page number field into header or footer.

```typescript
interface PageNumberOptions {
  format?: "1" | "i" | "I" | "a" | "A";  // Number format
  includeTotal?: boolean;                 // "Page X of Y"
  prefix?: string;                        // "Page "
  suffix?: string;                        // " of "
}

await DocTree.insertPageNumber(context, "footer:primary", {
  format: "1",
  includeTotal: true,
  prefix: "Page "
});
// Results in: "Page 1 of 10"
```

### updatePageNumberFormat(context, format)

Change page number format across all headers/footers.

## Dynamic Content Ideas

### insertDate(context, location, format?)

Insert current date field.

```typescript
await DocTree.insertDate(context, "header:primary", {
  format: "MMMM d, yyyy",
  updateAutomatically: true
});
```

### insertFilename(context, location, includePath?)

Insert document filename field.

```typescript
await DocTree.insertFilename(context, "footer:primary", {
  includePath: false
});
```

### insertAuthor(context, location)

Insert document author field.

### insertField(context, location, fieldCode)

Generic field insertion for advanced use cases.

```typescript
// Insert custom field
await DocTree.insertField(context, "header:primary", "AUTHOR");
await DocTree.insertField(context, "footer:primary", "NUMPAGES");
```

## Header/Footer Templates

### applyHeaderTemplate(context, templateName)

Apply predefined header layouts.

```typescript
type HeaderTemplate =
  | "simple-centered"        // Centered single line
  | "left-right"             // Left text, right text
  | "logo-title"             // Logo on left, title on right
  | "confidential-banner"    // Bold confidential notice
  | "none";                  // Clear header

await DocTree.applyHeaderTemplate(context, "left-right", {
  left: "Company Name",
  right: "Document Title"
});
```

### applyFooterTemplate(context, templateName)

Apply predefined footer layouts.

```typescript
type FooterTemplate =
  | "page-number-center"     // Centered page number
  | "page-of-total"          // "Page X of Y" centered
  | "left-page-right"        // Left text, page number, right text
  | "confidential-page"      // Confidential + page number
  | "none";                  // Clear footer

await DocTree.applyFooterTemplate(context, "page-of-total");
```

## Editing by Ref

### editHeaderByRef(context, ref, newText, options?)

Edit header content using refs from accessibility tree.

```typescript
await DocTree.editHeaderByRef(context, "hdr:0/primary/p:0", "New Header Text", {
  track: true
});
```

### editFooterByRef(context, ref, newText, options?)

Edit footer content using refs.

### deleteHeaderByRef(context, ref, options?)

Delete paragraph from header.

### deleteFooterByRef(context, ref, options?)

Delete paragraph from footer.

## Section-Aware Operations

### copyHeaderToAllSections(context, sourceSectionIndex, type?)

Copy header from one section to all others.

```typescript
// Copy first section's primary header to all sections
await DocTree.copyHeaderToAllSections(context, 0, "Primary");
```

### linkHeaderToSection(context, targetSectionIndex, sourceSectionIndex)

Link header to match another section (if API supports).

### getSectionHeaderInfo(context, sectionIndex)

Get detailed info about a section's header configuration.

```typescript
interface SectionHeaderInfo {
  hasDifferentFirstPage: boolean;
  hasDifferentOddEven: boolean;
  primaryHeader: { text: string; isEmpty: boolean };
  firstPageHeader: { text: string; isEmpty: boolean };
  evenPagesHeader: { text: string; isEmpty: boolean };
  // Footer equivalents
}
```

## Implementation Priority

### Phase 1: Core Read/Write (High Priority)

1. `getHeaderText` / `getFooterText`
2. `setHeaderText` / `setFooterText`
3. `clearHeader` / `clearFooter`
4. `getAllHeaders` / `getAllFooters`

### Phase 2: Accessibility Tree (Medium Priority)

1. Add `hdr:` and `ftr:` ref prefixes to types.ts
2. Build header/footer nodes in builder.ts when `includeHeaders`/`includeFooters` enabled
3. Add scope support for headers/footers

### Phase 3: Advanced Features (Lower Priority)

1. Page number insertion (requires field code research)
2. Dynamic content fields
3. Templates
4. Section linking operations

## Technical Notes

### API Requirement Sets

- Basic header/footer access: WordApi 1.1
- Content controls in headers: WordApi 1.1
- Some advanced features may require WordApi 1.4+

### Known Limitations

1. **No page number API**: Must use OOXML field codes
2. **No linking API**: Cannot programmatically link headers across sections
3. **No shapes/images**: Limited support for positioned elements
4. **Field insertion**: May require `insertOoxml()` with field codes

### Performance Considerations

- Load all header types in single sync when reading
- Use batch operations when modifying multiple sections
- Consider caching header content if frequently accessed
