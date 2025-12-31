---
name: office-bridge
description: Execute JavaScript in live Word documents via Office.js API. Use when users ask to edit open Word documents, insert text, apply formatting, work with tracked changes, or automate Word workflows. Trigger phrases include "edit my Word doc", "insert text in Word", "add tracked changes", "format the document", or any live Word document manipulation.
---

# Office Bridge Skill

Execute Office.js JavaScript directly in open Word documents. Maintains persistent connections to Word add-ins for live document editing.

## Setup

### First-Time Setup

Run the install script once to set up dependencies:

```bash
./office-bridge/install.sh
```

This installs npm dependencies and Office Add-in development certificates.

### Sideload the Add-in

The user must sideload the Word add-in once per machine.

**macOS:**

Symlink (or copy) the manifest to Word's `wef` folder. The filename must use the add-in's UUID:

```bash
# Create wef folder if it doesn't exist
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef

# Symlink the manifest using UUID-based filename (required by Word)
# The UUID comes from the <Id> element in manifest.xml
ln -sf /Users/parkerhancock/Projects/python_docx_redline/office-bridge/addin/manifest.xml \
  ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/19d3eb15-55fe-45a1-a33f-89b66b5aae3b.manifest.xml
```

Then **quit Word completely** (Command+Q) and reopen. The add-in appears in **Home** → **Add-ins**.

**Windows:**

1. Open Microsoft Word
2. Go to **Insert** → **Get Add-ins** → **My Add-ins** → **Upload My Add-in**
3. Browse to `office-bridge/addin/manifest.xml`

The add-in persists after sideloading.

### Start the Bridge Server

Before using this skill, start the bridge server:

```bash
./office-bridge/server.sh &
```

**Wait for the startup message showing the connection token.** The user must enter this token in the Word add-in taskpane.

## Writing Scripts

Execute scripts inline using heredocs:

```bash
cd office-bridge && bun x tsx <<'EOF'
import { connect } from "@/client.js";

const bridge = await connect();
const docs = await bridge.documents();

if (docs.length === 0) {
  console.log("No documents connected. Open Word and connect the add-in.");
  process.exit(1);
}

const doc = docs[0];
console.log(`Working with: ${doc.filename}`);

const result = await doc.executeJs(`
  const body = context.document.body;
  body.insertText("Hello from Claude!", Word.InsertLocation.end);
  await context.sync();
  return "Text inserted";
`);

console.log("Result:", result);
await bridge.close();
EOF
```

## Client API

```typescript
const bridge = await connect();           // Connect to bridge server
const docs = await bridge.documents();    // List connected Word documents
const doc = docs[0];                       // Get document handle

// Execute Office.js code in the document
const result = await doc.executeJs<T>(`
  // context is available (Word.RequestContext)
  // Must call await context.sync() to flush operations
  // Return JSON-serializable value
`);

await bridge.close();                      // Disconnect
```

## Key Principles

1. **Small scripts**: Each script should do ONE thing
2. **Always sync**: Call `await context.sync()` after queueing operations
3. **Return data**: Convert Word objects to plain data before returning
4. **Check connection**: Always verify documents are connected before executing

## Common Patterns

### Get Document Text

```javascript
const body = context.document.body;
body.load("text");
await context.sync();
return body.text;
```

### Get All Paragraphs

```javascript
const paragraphs = context.document.body.paragraphs;
paragraphs.load("items/text");
await context.sync();
return paragraphs.items.map(p => p.text);
```

### Search and Replace

```javascript
const results = context.document.body.search("old text", { matchCase: true });
results.load("items");
await context.sync();

for (const item of results.items) {
  item.insertText("new text", Word.InsertLocation.replace);
}
await context.sync();
return { replaced: results.items.length };
```

### Insert Text at Location

```javascript
const results = context.document.body.search("Section 2", { matchCase: true });
results.load("items");
await context.sync();

if (results.items.length > 0) {
  results.items[0].insertText(" (amended)", Word.InsertLocation.after);
  await context.sync();
}
return { found: results.items.length };
```

### Apply Formatting

```javascript
const results = context.document.body.search("Important", { matchCase: true });
results.load("items");
await context.sync();

for (const item of results.items) {
  item.font.bold = true;
  item.font.color = "red";
}
await context.sync();
return { formatted: results.items.length };
```

## Error Recovery

If a script fails, check the document connection:

```bash
cd office-bridge && bun x tsx <<'EOF'
import { connect } from "@/client.js";

const bridge = await connect();
const docs = await bridge.documents();

console.log("Connected documents:");
for (const doc of docs) {
  console.log(`  - ${doc.filename} (ID: ${doc.id}, status: ${doc.status})`);
}

if (docs.length === 0) {
  console.log("No documents connected!");
  console.log("1. Make sure Word is open with a document");
  console.log("2. Open the Office Bridge add-in taskpane");
  console.log("3. Enter the connection token from the server");
}

await bridge.close();
EOF
```

## Troubleshooting

- **"No documents connected"**: Open Word, click the Office Bridge add-in, enter the token from server.sh output
- **"Execution failed"**: Check the Office.js code syntax. Use try/catch in your code for detailed errors.
- **Server not running**: Run `./office-bridge/server.sh &` first
- **Token mismatch**: Restart server.sh to get a new token, enter it in the add-in
