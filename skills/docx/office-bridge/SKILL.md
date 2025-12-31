---
name: office-bridge
description: "Word Add-in API for live document editing via Office.js. Use when editing documents in Microsoft Word through the Office Bridge add-in, for real-time editing with the DocTree accessibility layer."
---

# Office Bridge Skill

Live editing of Word documents through a TypeScript/Office.js add-in with the DocTree accessibility layer.

## Architecture

```
                           WebSocket
Python/Claude  <------>  Bridge Server  <------>  Word Add-in
                         localhost:3847           Office.js API
                                                      |
                                                      v
                                               Microsoft Word
                                                (live document)
```

The add-in exposes a `DocTree` global object with all document operations.

## Quick Start

```typescript
await Word.run(async (context) => {
  // Build the accessibility tree
  const tree = await DocTree.buildTree(context);

  // Get YAML representation for LLM context
  const yaml = DocTree.toStandardYaml(tree);

  // Edit by ref
  await DocTree.replaceByRef(context, "p:5", "New text", { track: true });
});
```

## Decision Tree

| Task | Guide |
|------|-------|
| **First-time setup** | [setup.md](./setup.md) |
| Build tree, YAML serialization | [tree-building.md](./tree-building.md) |
| Ref-based editing, batch ops | [editing.md](./editing.md) |
| Text search, findAndHighlight | [search.md](./search.md) |
| Accept/reject tracked changes | [tracked-changes.md](./tracked-changes.md) |
| Comments: add, reply, resolve | [comments.md](./comments.md) |
| Navigation helpers, getNextRef | [navigation.md](./navigation.md) |
| Footnotes and endnotes | [footnotes.md](./footnotes.md) |
| Hyperlink operations | [hyperlinks.md](./hyperlinks.md) |
| Style management | [styles.md](./styles.md) |
| Headers and footers | [headers-footers.md](./headers-footers.md) |
| Table manipulation | [tables.md](./tables.md) |
| Selection and cursor | [selection.md](./selection.md) |

## Installation

**First-time setup?** See [setup.md](./setup.md) for complete step-by-step instructions.

### Quick Start (if already set up)

```bash
# Terminal 1: Start bridge server
cd office-bridge
./server.sh

# Terminal 2: Start add-in dev server
cd office-bridge/addin
npm run dev-server
```

Then open Word → Insert → Add-ins → My Add-ins → Office Bridge.

### Connect from Python

```python
import asyncio
import websockets
import json

async def connect():
    async with websockets.connect("wss://localhost:3847") as ws:
        await ws.send(json.dumps({
            "type": "execute",
            "code": "return await DocTree.buildTree(context);"
        }))
        result = await ws.recv()
        print(result)

asyncio.run(connect())
```

### Troubleshooting

| Issue | Solution |
|-------|----------|
| "Disconnected" status | Ensure `./server.sh` is running |
| Add-in won't load | Run `npm run dev-server` in `addin/` |
| Certificate errors | Run `npx office-addin-dev-certs install` |

See [setup.md](./setup.md) for detailed troubleshooting.

## Common Patterns

### Edit with Tracked Changes

```typescript
await DocTree.replaceByRef(context, "p:5", "Updated text", {
  track: true,
  author: "Reviewer"
});
```

### Batch Operations

```typescript
await DocTree.batchEdit(context, [
  { ref: "p:3", operation: "replace", newText: "New intro" },
  { ref: "p:7", operation: "delete" },
], { track: true });
```

### Scope-Based Editing

```typescript
// Replace all in a section
await DocTree.replaceByScope(context, tree, "section:Methods", "Updated");

// Search and replace within scope
await DocTree.searchReplaceByScope(context, tree,
  "section:Parties", "Plaintiff", "Defendant", { track: true });
```

## YAML Verbosity Levels

| Level | Tokens/Page | Use Case |
|-------|-------------|----------|
| `toMinimalYaml()` | ~500 | Large docs, quick overview |
| `toStandardYaml()` | ~1,500 | Default, balanced detail |
| `toFullYaml()` | ~3,000 | Full formatting info |

## Roadmap

See [ideas/](./ideas/) for planned enhancements:
- [footnotes-ideas.md](./ideas/footnotes-ideas.md) - Footnote/endnote improvements
- [headers-footers-ideas.md](./ideas/headers-footers-ideas.md) - Header/footer enhancements
- [hyperlinks-ideas.md](./ideas/hyperlinks-ideas.md) - Hyperlink handling
- [selection-ideas.md](./ideas/selection-ideas.md) - Selection/cursor features
- [styles-ideas.md](./ideas/styles-ideas.md) - Style management
- [tables-ideas.md](./ideas/tables-ideas.md) - Table manipulation

## See Also

- [setup.md](./setup.md) - Complete installation and troubleshooting guide
- [../python/accessibility.md](../python/accessibility.md) - Python AccessibilityTree API
