---
name: office-bridge-styles-ideas
description: "Roadmap ideas for Office.js style management features in the Office Bridge add-in."
---

# Style Management: Roadmap Ideas

This document captures ideas for enhancing style management in the Office Bridge Word add-in.

## Helper Functions to Build

### 1. getStyles()

Get all styles in the document with filtering options.

```typescript
interface GetStylesOptions {
  type?: Word.StyleType;        // Filter by type
  builtIn?: boolean;            // Only built-in or only custom
  inUse?: boolean;              // Only styles currently used in doc
}

interface StyleInfo {
  name: string;
  nameLocal: string;
  type: Word.StyleType;
  builtIn: boolean;
  inUse: boolean;
  baseStyle?: string;
}

async function getStyles(
  context: Word.RequestContext,
  options?: GetStylesOptions
): Promise<StyleInfo[]>;

// Example usage
const paragraphStyles = await getStyles(context, { type: "Paragraph" });
const customStyles = await getStyles(context, { builtIn: false });
const usedStyles = await getStyles(context, { inUse: true });
```

### 2. getStyleOfRef()

Get the style applied to a specific ref.

```typescript
interface RefStyleInfo {
  ref: string;
  style: string;
  styleBuiltIn?: string;
  isBuiltIn: boolean;
  formatting: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontSize?: number;
    fontName?: string;
    color?: string;
  };
}

async function getStyleOfRef(
  context: Word.RequestContext,
  ref: string
): Promise<RefStyleInfo | null>;

// Example usage
const info = await getStyleOfRef(context, "p:5");
console.log(`Paragraph uses style: ${info?.style}`);
```

### 3. applyStyleToRef() / applyStyleToRefs()

Apply a style to one or more refs.

```typescript
interface ApplyStyleResult {
  success: boolean;
  appliedCount: number;
  errors: Array<{ ref: string; error: string }>;
}

async function applyStyleToRef(
  context: Word.RequestContext,
  ref: string,
  style: string | Word.BuiltInStyleName
): Promise<{ success: boolean; error?: string }>;

async function applyStyleToRefs(
  context: Word.RequestContext,
  refs: string[],
  style: string | Word.BuiltInStyleName
): Promise<ApplyStyleResult>;

// Example usage
await applyStyleToRef(context, "p:0", Word.BuiltInStyleName.title);
await applyStyleToRefs(context, ["p:1", "p:2", "p:3"], "Normal");
```

### 4. styleExists()

Check if a style exists in the document.

```typescript
async function styleExists(
  context: Word.RequestContext,
  styleName: string
): Promise<boolean>;

// Example usage
if (await styleExists(context, "LegalParagraph")) {
  paragraph.style = "LegalParagraph";
} else {
  paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
}
```

### 5. getStyleHierarchy()

Get the inheritance chain for a style.

```typescript
interface StyleHierarchy {
  name: string;
  basedOn?: string;
  chain: string[];  // Full inheritance chain
}

async function getStyleHierarchy(
  context: Word.RequestContext,
  styleName: string
): Promise<StyleHierarchy | null>;

// Example usage
const hierarchy = await getStyleHierarchy(context, "Heading 1");
// { name: "Heading 1", basedOn: "Normal", chain: ["Heading 1", "Normal"] }
```

## Style-Based Scoping Features

### 1. Scope by Style Name

Extend the scope system to support style filtering:

```typescript
// String shortcut
const nodes = resolveScope(tree, "style:Heading 1");

// Dictionary filter
const nodes = resolveScope(tree, {
  style: "Normal",
  contains: "agreement"
});

// Multiple styles
const nodes = resolveScope(tree, {
  style: ["Heading 1", "Heading 2", "Heading 3"]
});
```

### 2. Edit All by Style

Batch operations on all paragraphs with a specific style:

```typescript
// Replace all "Quote" style paragraphs
await replaceByScope(context, tree, { style: "Quote" }, (text) => {
  return `"${text}"`;
});

// Format all Normal paragraphs
await formatByScope(context, tree, { style: "Normal" }, {
  fontSize: 11,
  lineSpacing: 1.15
});

// Delete all captions
await deleteByScope(context, tree, { style: "Caption" });
```

### 3. Style-Based Section Detection

Use heading styles to detect document sections:

```typescript
interface SectionByStyle {
  headingRef: string;
  headingText: string;
  headingStyle: string;
  level: number;  // Derived from style (Heading 1 = 1, etc.)
  contentRefs: string[];
}

async function getSectionsByStyle(
  context: Word.RequestContext,
  tree: AccessibilityTree
): Promise<SectionByStyle[]>;

// Example: Get all content under "Heading 2" styled paragraphs
const sections = await getSectionsByStyle(context, tree);
for (const section of sections) {
  if (section.level === 2) {
    console.log(`Section: ${section.headingText}`);
    console.log(`  Content: ${section.contentRefs.length} paragraphs`);
  }
}
```

## Style Analysis Features

### 1. detectInconsistentStyles()

Find paragraphs with unusual style usage:

```typescript
interface StyleInconsistency {
  ref: string;
  style: string;
  issue: "direct-formatting" | "unusual-style" | "missing-style";
  details: string;
}

async function detectInconsistentStyles(
  context: Word.RequestContext,
  tree: AccessibilityTree
): Promise<StyleInconsistency[]>;

// Example output:
// [
//   { ref: "p:5", style: "Normal", issue: "direct-formatting",
//     details: "Has bold applied directly instead of using Strong style" },
//   { ref: "p:12", style: "Body Text", issue: "unusual-style",
//     details: "Only 1 of 45 paragraphs uses this style" }
// ]
```

### 2. getStyleUsageStats()

Get statistics on style usage:

```typescript
interface StyleUsageStats {
  total: number;
  byStyle: Record<string, number>;
  unusedStyles: string[];
  mostUsed: string;
  leastUsed: string;
}

async function getStyleUsageStats(
  context: Word.RequestContext,
  tree: AccessibilityTree
): Promise<StyleUsageStats>;

// Example usage
const stats = await getStyleUsageStats(context, tree);
console.log(`Most used style: ${stats.mostUsed} (${stats.byStyle[stats.mostUsed]} times)`);
console.log(`Unused styles: ${stats.unusedStyles.join(", ")}`);
```

### 3. suggestStyleFixes()

Suggest style improvements:

```typescript
interface StyleSuggestion {
  ref: string;
  currentStyle: string;
  suggestedStyle: string;
  reason: string;
  confidence: "high" | "medium" | "low";
}

async function suggestStyleFixes(
  context: Word.RequestContext,
  tree: AccessibilityTree
): Promise<StyleSuggestion[]>;

// Example suggestions:
// - "p:5 looks like a heading (short, bold, followed by body text) - suggest Heading 2"
// - "p:12 has quote formatting - suggest Quote style instead of direct italic"
```

## Template/Theme Operations

### 1. extractStyleSheet()

Export styles for reuse:

```typescript
interface StyleSheet {
  name: string;
  version: string;
  styles: Array<{
    name: string;
    type: Word.StyleType;
    basedOn?: string;
    font?: Partial<Word.Font>;
    paragraphFormat?: Partial<Word.ParagraphFormat>;
  }>;
}

async function extractStyleSheet(
  context: Word.RequestContext
): Promise<StyleSheet>;

// Export document's style definitions
const sheet = await extractStyleSheet(context);
// Can be saved and applied to other documents (via Python or templates)
```

### 2. compareStyles()

Compare styles between documents:

```typescript
interface StyleComparison {
  matching: string[];
  onlyInSource: string[];
  onlyInTarget: string[];
  different: Array<{
    name: string;
    differences: string[];
  }>;
}

function compareStyles(
  sourceStyles: StyleSheet,
  targetStyles: StyleSheet
): StyleComparison;
```

### 3. Style Presets for Legal Documents

Pre-defined style configurations for common document types:

```typescript
const LegalStylePreset = {
  name: "Legal Contract",
  styles: {
    "Article Heading": { basedOn: "Heading 1", font: { bold: true, allCaps: true } },
    "Section Heading": { basedOn: "Heading 2", font: { bold: true } },
    "Body Text": { basedOn: "Normal", paragraph: { firstLineIndent: 36 } },
    "Definitions": { basedOn: "Normal", font: { italic: true } },
    "Signature Block": { basedOn: "Normal", paragraph: { alignment: "left" } }
  }
};

async function applyStylePreset(
  context: Word.RequestContext,
  preset: StylePreset
): Promise<void>;
```

## Integration with Python Library

### 1. Sync Styles Between Environments

Python creates/modifies styles, Office.js applies them:

```python
# Python side: Create style in template
from python_docx_redline import Document
from python_docx_redline.models.style import Style, StyleType, RunFormatting

doc = Document("template.docx")
doc.styles.ensure_style(
    style_id="LegalClause",
    name="Legal Clause",
    style_type=StyleType.PARAGRAPH,
    paragraph_formatting=ParagraphFormatting(indent_first_line=0.5)
)
doc.save("template_with_styles.docx")
```

```typescript
// Office.js side: Apply the style
await Word.run(async (context) => {
  // Style now exists in document (from Python)
  const paragraph = context.document.body.paragraphs.getFirst();
  paragraph.style = "Legal Clause";
  await context.sync();
});
```

### 2. Style Mapping for Cross-Platform Editing

Map style names between systems:

```typescript
const StyleMapping = {
  // Python style_id -> Office.js style name
  "Heading1": "Heading 1",
  "FootnoteReference": "footnote reference",
  "FootnoteText": "footnote text",
  "Normal": "Normal"
};

function translateStyle(pythonStyleId: string): string {
  return StyleMapping[pythonStyleId] || pythonStyleId;
}
```

## Future Considerations

### 1. Style Creation Workaround

If Microsoft adds `document.createStyle()` or expands `importStylesFromJson()`:

```typescript
// Hypothetical future API
async function createStyle(
  context: Word.RequestContext,
  definition: StyleDefinition
): Promise<Word.Style>;
```

### 2. Character Style Application

Currently styles apply to paragraphs. Character styles need run-level access:

```typescript
// Idea: Apply character style to text range within paragraph
async function applyCharacterStyle(
  context: Word.RequestContext,
  ref: string,
  startOffset: number,
  endOffset: number,
  characterStyleName: string
): Promise<void>;

// Example: Apply "Strong" to characters 10-20 in paragraph 5
await applyCharacterStyle(context, "p:5", 10, 20, "Strong");
```

### 3. Style Events

React to style changes:

```typescript
// Hypothetical: Listen for style changes
DocTree.onStyleChanged((event) => {
  console.log(`Paragraph ${event.ref} changed from ${event.oldStyle} to ${event.newStyle}`);
  // Rebuild relevant part of tree
});
```

## Priority Order

1. **High Priority (Core Functionality)**
   - `getStyleOfRef()` - Essential for tree building
   - `applyStyleToRef()` / `applyStyleToRefs()` - Basic editing
   - `styleExists()` - Safety checks
   - Style-based scope filtering

2. **Medium Priority (Enhanced Features)**
   - `getStyles()` with filtering
   - `getStyleUsageStats()`
   - `detectInconsistentStyles()`
   - Style-based section detection

3. **Lower Priority (Advanced Features)**
   - `extractStyleSheet()`
   - `compareStyles()`
   - Style presets
   - Character style application

## See Also

- [styles.md](../styles.md) - Current style management documentation
- [../office-bridge.md](../../office-bridge.md) - Main Office Bridge documentation
- [../../styles.md](../../styles.md) - Python StyleManager API
