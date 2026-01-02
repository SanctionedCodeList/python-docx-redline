# PowerPoint JavaScript API Reference

Practical reference for the PowerPoint JavaScript API (Office.js). Code runs inside `PowerPoint.run(async (context) => { ... })`.

## Key Concepts

### Proxy Objects and Synchronization

PowerPoint.js uses **proxy objects** that represent references to actual PowerPoint objects. Operations are queued and executed in batches via `context.sync()`.

**Critical pattern:**
1. Queue operations (create objects, set properties)
2. Call `context.sync()` to execute
3. Load properties you need to read
4. Call `context.sync()` again
5. Now you can read the properties

```javascript
// WRONG: Trying to use a newly created object immediately
const newSlide = context.presentation.slides.add();
newSlide.shapes.addTextBox("Hello");  // FAILS - slides.add() returns void!

// CORRECT: Get the slide after sync
context.presentation.slides.add();
await context.sync();
const count = context.presentation.slides.getCount();
await context.sync();
const newSlide = context.presentation.slides.getItemAt(count.value - 1);
newSlide.shapes.addTextBox("Hello", { left: 100, top: 100, width: 300, height: 50 });
await context.sync();
```

### Common Gotcha: slides.add() Returns Void

Unlike some other Office.js APIs, `slides.add()` does **not** return the new slide. New slides are always added at the end of the presentation.

**To get a reference to the newly added slide:**

```javascript
// Add the slide
context.presentation.slides.add();
await context.sync();

// Get the count to find the last slide
const slideCount = context.presentation.slides.getCount();
await context.sync();

// Get the new slide (it's the last one)
const newSlide = context.presentation.slides.getItemAt(slideCount.value - 1);
// Now you can work with newSlide
```

---

## Adding Slides

### Basic Add (Default Layout)

```javascript
context.presentation.slides.add();
await context.sync();
```

### Add with Specific Master and Layout

```javascript
const options = {
  slideMasterId: "2147483690#2908289500",
  layoutId: "2147483691#2499880"
};
context.presentation.slides.add(options);
await context.sync();
```

### Get Available Masters and Layouts

```javascript
const masters = context.presentation.slideMasters.load("id, name, layouts/items/name, layouts/items/id");
await context.sync();

for (const master of masters.items) {
  console.log(`Master: ${master.name} (${master.id})`);
  for (const layout of master.layouts.items) {
    console.log(`  Layout: ${layout.name} (${layout.id})`);
  }
}
```

### Add Slide Matching Selected Slide's Layout

```javascript
// Get selected slide index (1-based)
const selectedSlideIndex = await new Promise((resolve, reject) => {
  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.SlideRange,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(result.error);
      } else {
        resolve(result.value.slides[0].index);
      }
    }
  );
});

// Get the selected slide (convert to 0-based)
const selectedSlide = context.presentation.slides
  .getItemAt(selectedSlideIndex - 1)
  .load("slideMaster/id, layout/id");
await context.sync();

// Add new slide with same layout
context.presentation.slides.add({
  slideMasterId: selectedSlide.slideMaster.id,
  layoutId: selectedSlide.layout.id
});
await context.sync();
```

---

## Adding Shapes

All shape methods return a `PowerPoint.Shape` object.

### Text Box

```javascript
const shapes = context.presentation.slides.getItemAt(0).shapes;
const textbox = shapes.addTextBox("Hello World!", {
  left: 100,   // points from left edge
  top: 100,    // points from top edge
  width: 300,
  height: 50
});
textbox.name = "MyTextBox";
await context.sync();
```

### Geometric Shape

```javascript
const shapes = context.presentation.slides.getItemAt(0).shapes;
const rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
  left: 100,
  top: 200,
  width: 150,
  height: 100
});
rectangle.name = "MyRectangle";
await context.sync();
```

**Available GeometricShapeType values:**
- `rectangle`, `roundRectangle`, `ellipse`, `diamond`
- `triangle`, `rightTriangle`, `parallelogram`, `trapezoid`
- `pentagon`, `hexagon`, `heptagon`, `octagon`
- `star4`, `star5`, `star6`, `star7`, `star8`
- `heart`, `lightning`, `sun`, `moon`, `cloud`
- `arc`, `bracePair`, `bracketPair`
- And many more...

### Line

```javascript
const shapes = context.presentation.slides.getItemAt(0).shapes;
const line = shapes.addLine(PowerPoint.ConnectorType.straight, {
  left: 100,    // start X
  top: 100,     // start Y
  width: 200,   // end X = left + width
  height: 100   // end Y = top + height
});
line.name = "MyLine";
await context.sync();
```

**ConnectorType values:** `straight`, `elbow`, `curve`

### Table (Requires PowerPointApi 1.8+)

```javascript
const shapes = context.presentation.slides.getItemAt(0).shapes;
const tableShape = shapes.addTable(3, 4, {  // 3 rows, 4 columns
  left: 100,
  top: 100,
  width: 400,
  height: 200
});
await context.sync();

// Access the table via tableShape.table
```

---

## Text Content and Formatting

### Setting Text Content

```javascript
const shapes = context.presentation.slides.getItemAt(0).shapes;
shapes.load("items/name,items/textFrame");
await context.sync();

const shape = shapes.items.find(s => s.name === "MyTextBox");
if (shape) {
  shape.textFrame.textRange.text = "Updated text content";
  await context.sync();
}
```

### Text Formatting

```javascript
const shapes = context.presentation.slides.getItemAt(0).shapes;
const textbox = shapes.addTextBox("Formatted Text", {
  left: 100, top: 100, width: 300, height: 100
});

// Font formatting
textbox.textFrame.textRange.font.color = "blue";
textbox.textFrame.textRange.font.size = 24;
textbox.textFrame.textRange.font.bold = true;
textbox.textFrame.textRange.font.italic = true;
textbox.textFrame.textRange.font.name = "Arial";

// Vertical alignment
textbox.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;

await context.sync();
```

### Shape Fill

```javascript
const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
  left: 100, top: 100, width: 200, height: 100
});

// Solid color fill
shape.fill.setSolidColor("lightblue");

// Or with hex color
shape.fill.setSolidColor("#4472C4");

await context.sync();
```

---

## Working with Existing Slides

### Get All Slides

```javascript
const slides = context.presentation.slides;
slides.load("items/id,items/index");
await context.sync();

return slides.items.map(s => ({ id: s.id, index: s.index }));
```

### Get Slide by Index (0-based)

```javascript
const slide = context.presentation.slides.getItemAt(0);
slide.load("id");
await context.sync();
console.log(slide.id);
```

### Get Slide by ID

```javascript
const slide = context.presentation.slides.getItem("slideId");
// or safely:
const slide = context.presentation.slides.getItemOrNullObject("slideId");
slide.load("isNullObject");
await context.sync();

if (!slide.isNullObject) {
  // slide exists
}
```

### Delete a Slide

```javascript
const slide = context.presentation.slides.getItemAt(2);  // third slide
slide.delete();
await context.sync();
```

### Move a Slide

```javascript
const slide = context.presentation.slides.getItemAt(0);
slide.moveTo(3);  // Move to position 3 (0-based)
await context.sync();
```

---

## Working with Shapes on a Slide

### Get All Shapes

```javascript
const slide = context.presentation.slides.getItemAt(0);
const shapes = slide.shapes;
shapes.load("items/id,items/name,items/type");
await context.sync();

return shapes.items.map(s => ({
  id: s.id,
  name: s.name,
  type: s.type
}));
```

### Find Shape by Name

```javascript
const slide = context.presentation.slides.getItemAt(0);
const shapes = slide.shapes;
shapes.load("items/name,items/textFrame");
await context.sync();

const titleShape = shapes.items.find(s => s.name.includes("Title"));
if (titleShape) {
  titleShape.textFrame.textRange.text = "New Title";
  await context.sync();
}
```

### Delete a Shape

```javascript
const slide = context.presentation.slides.getItemAt(0);
const shapes = slide.shapes;
shapes.load("items");
await context.sync();

if (shapes.items.length > 0) {
  shapes.items[0].delete();
  await context.sync();
}
```

---

## Complete Example: Add Slide with Content

```javascript
// Step 1: Add a new slide
context.presentation.slides.add();
await context.sync();

// Step 2: Get reference to the new slide
const slideCount = context.presentation.slides.getCount();
await context.sync();
const newSlide = context.presentation.slides.getItemAt(slideCount.value - 1);

// Step 3: Add a title text box
const title = newSlide.shapes.addTextBox("Quarterly Report", {
  left: 50,
  top: 50,
  width: 620,
  height: 60
});
title.textFrame.textRange.font.size = 32;
title.textFrame.textRange.font.bold = true;
title.name = "Title";

// Step 4: Add a subtitle
const subtitle = newSlide.shapes.addTextBox("Q4 2025 Results", {
  left: 50,
  top: 120,
  width: 620,
  height: 40
});
subtitle.textFrame.textRange.font.size = 18;
subtitle.textFrame.textRange.font.color = "gray";
subtitle.name = "Subtitle";

// Step 5: Add a content box
const content = newSlide.shapes.addTextBox("Key highlights from this quarter...", {
  left: 50,
  top: 200,
  width: 620,
  height: 300
});
content.name = "Content";

await context.sync();
return "Slide created successfully";
```

---

## API Requirement Sets

| Feature | Requirement Set |
|---------|-----------------|
| Basic slide operations | PowerPointApi 1.2 |
| Shapes, layouts, masters | PowerPointApi 1.3 |
| addTextBox, addGeometricShape, addLine | PowerPointApi 1.4 |
| addTable | PowerPointApi 1.8 |
| Slide background | PowerPointApi 1.10 |

---

## Key Differences from Word.js API

| Aspect | Word.js | PowerPoint.js |
|--------|---------|---------------|
| Primary unit | Paragraphs, Ranges | Slides, Shapes |
| Text containers | Paragraphs, ContentControls | Shapes with TextFrames |
| add() returns | Returns the new object | Returns void (for slides) |
| Tracked changes | Full support | Not supported |
| Search & replace | Built-in methods | Manual shape iteration |
| Object maturity | Very mature | Still expanding |

### Critical difference: Getting newly created objects

**Word.js:**
```javascript
const paragraph = context.document.body.insertParagraph("Text", Word.InsertLocation.end);
paragraph.font.bold = true;  // Works immediately
await context.sync();
```

**PowerPoint.js:**
```javascript
// slides.add() returns void, not the slide
context.presentation.slides.add();
await context.sync();
// Must get the slide separately
const count = context.presentation.slides.getCount();
await context.sync();
const newSlide = context.presentation.slides.getItemAt(count.value - 1);
```

---

## Troubleshooting

### "undefined is not an object" after slides.add()

**Problem:** Calling `slide.load("shapes")` or accessing properties on a slide immediately after `slides.add()`.

**Cause:** `slides.add()` returns `void`, not the new slide object.

**Solution:**
```javascript
// Add the slide
context.presentation.slides.add();
await context.sync();

// Get the slide count
const count = context.presentation.slides.getCount();
await context.sync();

// Now get the new slide
const newSlide = context.presentation.slides.getItemAt(count.value - 1);
newSlide.load("shapes");
await context.sync();
// Now you can access newSlide.shapes
```

### Properties not available after sync

**Problem:** Property throws "not available" error even after `load()` and `sync()`.

**Solution:** Make sure you're loading the specific property you need:
```javascript
// Load specific properties
slide.load("shapes");
shapes.load("items/name,items/type");
await context.sync();
// Now accessible
```

### Slides always added at end

**Limitation:** The API does not support inserting slides at specific positions during creation.

**Workaround:** Add the slide, then use `slide.moveTo(index)`:
```javascript
context.presentation.slides.add();
await context.sync();
const count = context.presentation.slides.getCount();
await context.sync();
const newSlide = context.presentation.slides.getItemAt(count.value - 1);
newSlide.moveTo(0);  // Move to beginning
await context.sync();
```

---

## Sources

- [Work with shapes using the PowerPoint JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/shapes)
- [Add and delete slides in PowerPoint](https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/add-slides)
- [PowerPoint JavaScript object model](https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/core-concepts)
- [PowerPoint.SlideCollection class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slidecollection?view=powerpoint-js-preview)
- [PowerPoint.ShapeCollection class](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapecollection?view=powerpoint-js-preview)
- [Using the application-specific API model](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model)
