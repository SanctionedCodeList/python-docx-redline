---
name: office-bridge-styles
description: "Office.js style management for Word add-ins. Use when reading, applying, or managing styles in live Word documents through the Office Bridge."
---

# Office Bridge: Style Management

This guide covers style management in Office.js for the Word add-in. Understanding these APIs enables consistent document formatting, style-based scoping, and template operations.

## API Overview

| Capability | API Set | Support |
|------------|---------|---------|
| Read styles | WordApi 1.5 | Full support |
| Apply built-in styles | WordApi 1.1 | Full support |
| Apply custom styles | WordApi 1.1 | Full support (style must exist) |
| Modify style properties | WordApi 1.5+ | Partial (see limitations) |
| Create new styles | Not available | Use OOXML workaround |
| Import styles from JSON | WordApi 1.5 | Preview only |

## Core Concepts

### Style vs StyleBuiltIn

Office.js provides two properties for working with paragraph styles:

```typescript
// paragraph.style - for custom styles and localized names
paragraph.style = "MyCustomStyle";

// paragraph.styleBuiltIn - for built-in styles (locale-independent)
paragraph.styleBuiltIn = Word.BuiltInStyleName.heading1;
```

**Use `styleBuiltIn` for built-in styles** to ensure portability across locales. The `style` property is for custom styles or when you need localized style names.

### Style Types

Word supports four style types (from `Word.StyleType`):

| Type | Value | Description |
|------|-------|-------------|
| Character | "Character" | Applied to runs of text |
| Paragraph | "Paragraph" | Applied to whole paragraphs |
| Table | "Table" | Applied to tables |
| List | "List" | Applied to lists |

## Reading Styles

### Get All Document Styles

```typescript
await Word.run(async (context) => {
  const styles = context.document.getStyles();
  styles.load("items");
  await context.sync();

  for (const style of styles.items) {
    style.load(["nameLocal", "type", "builtIn", "inUse"]);
  }
  await context.sync();

  for (const style of styles.items) {
    console.log(`${style.nameLocal} (${style.type}) - ` +
                `built-in: ${style.builtIn}, in use: ${style.inUse}`);
  }
});
```

### Get Style by Name

```typescript
await Word.run(async (context) => {
  // Safe approach - returns null object if not found
  const style = context.document.getStyles().getByNameOrNullObject("Heading 1");
  style.load(["nameLocal", "type", "font", "paragraphFormat"]);
  await context.sync();

  if (style.isNullObject) {
    console.log("Style not found");
    return;
  }

  console.log(`Found: ${style.nameLocal}`);
});
```

### Get Paragraph Style

```typescript
await Word.run(async (context) => {
  const paragraph = context.document.body.paragraphs.getFirst();
  paragraph.load(["style", "styleBuiltIn"]);
  await context.sync();

  console.log(`Style: ${paragraph.style}`);
  console.log(`Built-in: ${paragraph.styleBuiltIn}`);
});
```

## Applying Styles

### Apply Built-In Styles

```typescript
await Word.run(async (context) => {
  const paragraph = context.document.body.paragraphs.getFirst();

  // Use styleBuiltIn for locale-independent built-in styles
  paragraph.styleBuiltIn = Word.BuiltInStyleName.heading1;

  await context.sync();
});
```

### Apply Custom Styles

```typescript
await Word.run(async (context) => {
  const paragraph = context.document.body.paragraphs.getFirst();

  // Custom style must already exist in the document
  paragraph.style = "LegalParagraph";

  await context.sync();
});
```

### Apply Style to Selection

```typescript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("paragraphs");
  await context.sync();

  for (const para of selection.paragraphs.items) {
    para.styleBuiltIn = Word.BuiltInStyleName.quote;
  }
  await context.sync();
});
```

### Apply Style by Ref (DocTree Integration)

```typescript
// Apply style to a specific paragraph by ref
async function applyStyleByRef(
  context: Word.RequestContext,
  ref: string,
  styleName: string
): Promise<{ success: boolean; error?: string }> {
  try {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    // Parse ref (assumes format "p:N")
    const match = ref.match(/^p:(\d+)$/);
    if (!match) {
      return { success: false, error: `Invalid paragraph ref: ${ref}` };
    }

    const index = parseInt(match[1], 10);
    if (index >= paragraphs.items.length) {
      return { success: false, error: `Paragraph ${index} not found` };
    }

    paragraphs.items[index].style = styleName;
    await context.sync();

    return { success: true };
  } catch (error) {
    return { success: false, error: String(error) };
  }
}
```

## Style Properties

### Read Style Font Properties

```typescript
await Word.run(async (context) => {
  const style = context.document.getStyles().getByNameOrNullObject("Heading 1");
  style.load("font");
  await context.sync();

  if (!style.isNullObject) {
    const font = style.font;
    font.load(["name", "size", "bold", "italic", "color"]);
    await context.sync();

    console.log(`Font: ${font.name}, Size: ${font.size}, Bold: ${font.bold}`);
  }
});
```

### Modify Style Font (WordApi 1.5+)

```typescript
await Word.run(async (context) => {
  const style = context.document.getStyles().getByName("Normal");
  style.load("font");
  await context.sync();

  // Modify font properties
  style.font.name = "Arial";
  style.font.size = 11;
  style.font.color = "#333333";

  await context.sync();
});
```

### Read Paragraph Format

```typescript
await Word.run(async (context) => {
  const style = context.document.getStyles().getByNameOrNullObject("Normal");
  style.load("paragraphFormat");
  await context.sync();

  if (!style.isNullObject) {
    const pf = style.paragraphFormat;
    pf.load(["alignment", "spaceAfter", "spaceBefore", "lineSpacing"]);
    await context.sync();

    console.log(`Alignment: ${pf.alignment}, Line spacing: ${pf.lineSpacing}`);
  }
});
```

## Style Collections in DocTree

### Include Style in Accessibility Tree

When building the accessibility tree, style information is captured in the `NodeStyle` interface:

```typescript
interface NodeStyle {
  name?: string;           // Style name (e.g., "Heading1", "Normal")
  formatting?: {           // Direct formatting overrides
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    // ... other run properties
  };
}
```

### Reading Style During Tree Build

```typescript
// In buildTree(), for each paragraph:
const paragraph = paragraphs.items[i];
paragraph.load(["style", "font"]);
await context.sync();

const node: AccessibilityNode = {
  ref: `p:${i}`,
  role: SemanticRole.Paragraph,
  text: paragraph.text,
  style: {
    name: paragraph.style,
    formatting: {
      bold: paragraph.font.bold,
      italic: paragraph.font.italic,
      // ... capture direct formatting
    }
  }
};
```

## Built-In Style Names

Common built-in style names available through `Word.BuiltInStyleName`:

| Style | Usage |
|-------|-------|
| `normal` | Default paragraph style |
| `heading1` - `heading9` | Section headings |
| `title` | Document title |
| `subtitle` | Document subtitle |
| `quote` | Block quotations |
| `intenseQuote` | Emphasized quotations |
| `listParagraph` | List items |
| `noSpacing` | Paragraphs without spacing |
| `emphasis` | Emphasized text |
| `strong` | Strong emphasis |
| `footnoteText` | Footnote content |
| `endnoteText` | Endnote content |
| `caption` | Figure/table captions |

## Limitations

### Cannot Create New Styles

**Office.js does not support creating new styles programmatically.** The GitHub issue requesting this feature remains open since 2018.

**Workarounds:**

1. **Pre-define styles in templates** - Include all needed custom styles in your document templates

2. **Use existing styles** - Apply built-in styles or styles that already exist in the document

3. **Import via JSON (Preview)** - Use `document.importStylesFromJson()` (WordApi 1.5 preview):

```typescript
// PREVIEW API - may change
await Word.run(async (context) => {
  const stylesJson = JSON.stringify({
    styles: [{
      name: "CustomHeading",
      type: "paragraph",
      basedOn: "Heading 1",
      font: { bold: true, color: "#0066CC" }
    }]
  });

  context.document.importStylesFromJson(
    stylesJson,
    Word.ImportedStylesConflictBehavior.overwrite
  );
  await context.sync();
});
```

4. **OOXML Insertion** - Insert OOXML containing style definitions (unreliable across platforms)

### Limited Style Property Access

Some style properties require specific API sets:

| Property | Required API Set |
|----------|-----------------|
| `font` | WordApi 1.5 |
| `paragraphFormat` | WordApi 1.5 |
| `shading` | WordApi 1.6 |
| `borders` | WordApiDesktop 1.1 |
| `frame` | WordApiDesktop 1.3 |
| `listTemplate` | WordApiDesktop 1.1 |

### Platform Differences

Some APIs are desktop-only:

```typescript
// Check if desktop API is available
if (Office.context.requirements.isSetSupported("WordApiDesktop", "1.1")) {
  // Use desktop-specific APIs like borders, listTemplate
}
```

## Best Practices

### 1. Use StyleBuiltIn for Built-In Styles

```typescript
// Good - locale independent
paragraph.styleBuiltIn = Word.BuiltInStyleName.heading1;

// Avoid - locale dependent
paragraph.style = "Heading 1";  // Fails in non-English locales
```

### 2. Check Style Existence Before Applying

```typescript
async function safeApplyStyle(
  context: Word.RequestContext,
  paragraph: Word.Paragraph,
  styleName: string
): Promise<boolean> {
  const style = context.document.getStyles().getByNameOrNullObject(styleName);
  await context.sync();

  if (style.isNullObject) {
    console.warn(`Style "${styleName}" not found`);
    return false;
  }

  paragraph.style = styleName;
  await context.sync();
  return true;
}
```

### 3. Batch Style Operations

```typescript
await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();

  // Apply styles in batch
  paragraphs.items[0].styleBuiltIn = Word.BuiltInStyleName.title;
  paragraphs.items[1].styleBuiltIn = Word.BuiltInStyleName.subtitle;
  for (let i = 2; i < 5; i++) {
    paragraphs.items[i].styleBuiltIn = Word.BuiltInStyleName.normal;
  }

  // Single sync for all changes
  await context.sync();
});
```

### 4. Capture Styles During Tree Build

When building accessibility trees, include style information for style-based scoping:

```typescript
// Build tree with style info
const tree = await buildTree(context, { includeStyles: true });

// Then filter by style
const headingNodes = filterByScope(tree, { style: "Heading 1" });
```

## See Also

- [office-bridge.md](../office-bridge.md) - Main Office Bridge documentation
- [styles.md](../styles.md) - Python StyleManager API
- [Microsoft Word.Style API](https://learn.microsoft.com/en-us/javascript/api/word/word.style)
- [Microsoft Word.StyleCollection API](https://learn.microsoft.com/en-us/javascript/api/word/word.stylecollection)
