# PowerPoint Patterns

PowerPoint provides presentation, slide, and shape manipulation via `PowerPoint.run()`.

## Getting Presentations

```typescript
const bridge = await connect();
const presentations = await bridge.powerpoint();  // Returns PowerPointSession[]
const ppt = presentations[0];
```

## Execution Context

Code runs inside `PowerPoint.run()` with access to `context`, `PowerPoint`, and `Office`:

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();
return slides.items.length;
```

## Common Patterns

### Get Slide Count

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();
return slides.items.length;
```

### Add a Slide

```javascript
context.presentation.slides.add();
await context.sync();
return "Slide added";
```

### Get All Slide IDs

```javascript
const slides = context.presentation.slides;
slides.load("items/id");
await context.sync();
return slides.items.map(s => s.id);
```

### Get Shapes on a Slide

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const firstSlide = slides.items[0];
const shapes = firstSlide.shapes;
shapes.load("items/name,items/type");
await context.sync();

return shapes.items.map(s => ({ name: s.name, type: s.type }));
```

### Add a Text Box

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const textBox = slide.shapes.addTextBox("Hello from Claude!", {
  left: 100,
  top: 100,
  width: 300,
  height: 50
});
await context.sync();
return "Text box added";
```

### Delete a Slide

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

if (slides.items.length > 1) {
  slides.items[slides.items.length - 1].delete();
  await context.sync();
  return "Last slide deleted";
}
return "Cannot delete only slide";
```

### Get Selected Slides

```javascript
const selection = context.presentation.getSelectedSlides();
selection.load("items/id");
await context.sync();
return selection.items.map(s => s.id);
```

### Set Slide Title

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const shapes = slide.shapes;
shapes.load("items/name,items/textFrame");
await context.sync();

const titleShape = shapes.items.find(s => s.name.includes("Title"));
if (titleShape) {
  titleShape.textFrame.textRange.text = "New Title";
  await context.sync();
  return "Title updated";
}
return "No title shape found";
```

## Tips

- Slides are 0-indexed in the items array
- Shape positions use points (1 inch = 72 points)
- Use `.load()` before accessing properties
- PowerPoint API is less mature than Word/Excel - some operations may not be available
