# Office Bridge Specification

> **Status**: Draft
> **Created**: 2025-12-30
> **Branch**: `feature/word-add-in`

## Overview

A minimal bridge architecture that enables Claude Code (or other agents) to execute JavaScript directly in Word documents via the Office.js API. The design follows the "thin bridge" pattern used by Claude in Chrome—exposing native APIs with minimal abstraction.

**Key Design Decisions**:
- TypeScript client library with Playwright-style API (not MCP tools)
- Raw Office.js execution—no wrapper APIs, Claude writes native JS
- Desktop Word only for v1

## Architecture

```
┌─────────────────┐         WebSocket          ┌─────────────────┐
│   Word          │◄─────────────────────────►│   Bridge        │
│   Add-in        │    wss://localhost:PORT   │   Server        │
│   (Office.js)   │                           │   (Node.js)     │
└─────────────────┘                           └────────┬────────┘
                                                       │
        ┌─────────────────┐                           │
        │   Word          │◄──────────────────────────┤
        │   Add-in #2     │                           │
        └─────────────────┘                           │
                                                       │
                                              ┌────────▼────────┐
                                              │   Claude Code   │
                                              │   (TS Client)   │
                                              └─────────────────┘
```

## Design Principles

1. **Minimal Abstraction**: The bridge passes JavaScript code directly to Word's Office.js runtime. No wrapper APIs, no convenience methods in the bridge itself.

2. **Playwright-style API**: Document-centric client library. Get a document handle, call methods on it.

3. **Auto-launch Server**: Client spawns bridge server on-demand if not already running.

4. **Per-document Add-ins**: Each Word document the user wants to control has its own add-in instance and WebSocket connection.

5. **Raw Context Exposure**: Executed code gets direct access to the Office.js `context`. Claude manages `sync()` calls.

## Package Structure

Located at `python-docx-redline/office-bridge/` as a Claude Code skill with TypeScript source:

```
office-bridge/
├── SKILL.md                 # Claude Code skill instructions
├── install.sh               # One-time setup (deps + certs)
├── server.sh                # Start bridge server
├── package.json             # Dependencies
├── tsconfig.json            # TypeScript config
│
├── src/
│   ├── client.ts            # OfficeBridge client API
│   ├── server.ts            # Bridge server
│   ├── registry.ts          # Document registry
│   └── types.ts             # Shared types
│
├── addin/                   # Word add-in static assets
│   ├── taskpane.html
│   ├── taskpane.js
│   └── manifest.xml
│
└── tmp/                     # Scratch space for scripts
```

### Skill Integration

The skill follows the same pattern as `dev-browser`: Claude writes inline TypeScript that imports from the client module. To make this skill available to Claude Code, symlink it:

```bash
ln -s /path/to/python-docx-redline/office-bridge ~/.agents/local_skills/office-bridge
```

## Client API

### Basic Usage

```typescript
import { connect } from "@/client.js";

// Connect to bridge (auto-launches server if needed)
const bridge = await connect();

// List connected documents
const docs = await bridge.documents();
// Returns: [{ id: "abc123", filename: "contract.docx", ... }]

// Get a document handle
const doc = docs.find(d => d.filename.includes('contract'));

// Execute Office.js code
const result = await doc.executeJs<string>(`
  const body = context.document.body;
  body.insertText("Hello from Claude!", Word.InsertLocation.end);
  await context.sync();
  return "Text inserted";
`);

// Close connection (documents stay connected to bridge)
await bridge.close();
```

### Bridge Client Interface

```typescript
interface BridgeClient {
  // List all connected Word documents
  documents(): Promise<WordDocument[]>;

  // Wait for a document matching criteria
  waitForDocument(filter: DocumentFilter, timeout?: number): Promise<WordDocument>;

  // Close client connection (documents persist on bridge)
  close(): Promise<void>;

  // Event handlers
  on(event: 'document', handler: (doc: WordDocument) => void): void;
  on(event: 'disconnect', handler: (doc: DocumentInfo) => void): void;
  on(event: 'console', handler: (entry: ConsoleEntry) => void): void;
}

interface ConnectOptions {
  port?: number;              // Server port (default: 3847)
  autoLaunch?: boolean;       // Launch server if not running (default: true)
}

async function connect(options?: ConnectOptions): Promise<BridgeClient>;
```

### WordDocument Interface

```typescript
interface WordDocument {
  readonly id: string;
  readonly filename: string;
  readonly path: string;
  readonly connectedAt: Date;
  readonly status: 'connected' | 'disconnected';

  // Execute JavaScript in Word.run() context
  executeJs<T = unknown>(code: string, options?: ExecuteOptions): Promise<T>;

  // Event: console output from this document
  on(event: 'console', handler: (entry: ConsoleEntry) => void): void;
}

interface ExecuteOptions {
  timeout?: number;           // Timeout in ms (default: 30000)
}

interface ConsoleEntry {
  level: 'log' | 'warn' | 'error' | 'info';
  message: string;
  timestamp: Date;
  documentId: string;
}
```

## Bridge Server

### Responsibilities

1. **WebSocket endpoint** for Word add-in connections
2. **HTTP server** hosting add-in static assets
3. **Document registry** tracking connected Word instances
4. **Client communication** via separate WebSocket or IPC

### Document Registry

```typescript
interface ConnectedDocument {
  id: string;              // Unique connection ID (UUID)
  filename: string;        // Document filename
  path: string;            // Full path (if available)
  connectedAt: Date;       // Connection timestamp
  lastActivity: Date;      // Last command execution
  status: 'connected' | 'disconnected';
  disconnectedAt?: Date;   // When disconnected (for tombstones)
}
```

**Tombstone Behavior**: When a document disconnects unexpectedly, its registry entry is marked `status: 'disconnected'` and retained for 5 minutes. This allows Claude to see what happened and understand failures.

### Execution Queue

Each document has a sequential execution queue. Requests are processed FIFO to avoid Office.js context conflicts. Concurrent requests to different documents are allowed.

### Security

**Shared Secret Authentication**: On startup, the bridge server generates a one-time token and displays it in the terminal. The add-in must present this token when registering. This prevents malicious documents from connecting.

```
Bridge server started on port 3847
Connection token: a1b2c3d4e5f6
```

## WebSocket Protocol

### Add-in → Bridge: Registration

```json
{
  "type": "register",
  "payload": {
    "token": "a1b2c3d4e5f6",
    "filename": "contract.docx",
    "path": "/Users/parker/Documents/contract.docx"
  }
}
```

### Bridge → Add-in: Registration Acknowledgment

```json
{
  "type": "registered",
  "payload": {
    "id": "abc123-def456"
  }
}
```

### Bridge → Add-in: Execute JavaScript

```json
{
  "type": "execute",
  "id": "req-001",
  "payload": {
    "code": "context.document.body.insertText('Hello', Word.InsertLocation.end); await context.sync(); return 'done';",
    "timeout": 30000
  }
}
```

### Add-in → Bridge: Execution Result

```json
{
  "type": "result",
  "id": "req-001",
  "payload": {
    "success": true,
    "result": "done"
  }
}
```

### Add-in → Bridge: Execution Error

```json
{
  "type": "result",
  "id": "req-001",
  "payload": {
    "success": false,
    "error": {
      "message": "The text was not found",
      "code": "ItemNotFound",
      "stack": "..."
    }
  }
}
```

### Add-in → Bridge: Console Output (Real-time)

```json
{
  "type": "console",
  "payload": {
    "level": "log",
    "message": "Processing paragraph 5 of 100",
    "timestamp": "2025-12-30T10:30:00.000Z"
  }
}
```

## JavaScript Execution Context

Code sent via `executeJs` is wrapped and executed as:

```javascript
await Word.run(async (context) => {
  // User code is inserted here
  // `context` is available
  // Must call `await context.sync()` to flush operations

  ${userCode}
});
```

### Available Globals

- `context` - Word.RequestContext for the document
- `Word` - Word JavaScript API namespace
- `Office` - Office.js common API
- `console` - Logging (streamed to bridge in real-time)

### Return Values

- Code should `return` a JSON-serializable value
- Complex objects (Range, Paragraph, etc.) should be converted to plain data before returning
- Errors are caught and returned with stack traces

### Timeout Behavior

Default timeout is 30 seconds, configurable per-request. On timeout:
1. Bridge sends cancellation signal to add-in
2. Add-in attempts to abort (best-effort—Office.js doesn't support true cancellation)
3. Error returned to client with `code: 'Timeout'`

## Word Add-in

### Type and Permissions

- **Type**: Taskpane add-in (minimal UI showing connection status)
- **Permissions**: ReadWriteDocument
- **Requirements**: WordApi 1.6+ (for tracked changes support)
- **Platform**: Desktop Word only (Mac and Windows)

### Per-Document Model

Each Word document requires its own add-in instance. Office.js sandboxes each add-in to its host document—a taskpane opened in Document A cannot access Document B.

When user wants Claude to access a document:
1. Open the document in Word
2. Insert → Add-ins → Office Bridge
3. Add-in connects to bridge server

### Taskpane UI

Minimal status display:

```
┌─────────────────────────────┐
│  Office Bridge              │
├─────────────────────────────┤
│  Status: ● Connected        │
│  Document: contract.docx    │
│  ID: abc123                 │
│                             │
│  Last activity: 2s ago      │
│                             │
│  [Disconnect]               │
└─────────────────────────────┘
```

### Auto-Reconnect

If bridge server restarts, add-in:
1. Detects WebSocket close
2. Shows "Reconnecting..." status
3. Retries connection with exponential backoff (1s, 2s, 4s, max 30s)
4. Re-registers when server is available
5. Resumes normal operation

## Claude Code Skill

The office-bridge includes a SKILL.md that teaches Claude Code how to use it. This follows the same pattern as the `dev-browser` skill.

### SKILL.md Content

```markdown
---
name: office-bridge
description: Execute JavaScript in live Word documents via Office.js API. Use when users ask to edit open Word documents, insert text, apply formatting, work with tracked changes, or automate Word workflows. Trigger phrases include "edit my Word doc", "insert text in Word", "add tracked changes", "format the document", or any live Word document manipulation.
---

# Office Bridge Skill

Execute Office.js JavaScript directly in open Word documents. Maintains persistent connections to Word add-ins for live document editing.

## Setup

First, run the install script (one-time):

\`\`\`bash
./office-bridge/install.sh
\`\`\`

Then start the bridge server:

\`\`\`bash
./office-bridge/server.sh &
\`\`\`

**Wait for the "Bridge server started" message before running scripts.**

The user must also sideload the Word add-in:
1. Open Word → Insert → Add-ins → My Add-ins → Upload My Add-in
2. Select `office-bridge/addin/manifest.xml`
3. Open the add-in taskpane in each document Claude should access

## Writing Scripts

Execute scripts inline using heredocs:

\`\`\`bash
cd office-bridge && bun x tsx <<'EOF'
import { connect } from "@/client.js";

const bridge = await connect();
const docs = await bridge.documents();
const doc = docs[0];

const result = await doc.executeJs(\`
  const body = context.document.body;
  body.insertText("Hello from Claude!", Word.InsertLocation.end);
  await context.sync();
  return "done";
\`);

console.log(result);
await bridge.close();
EOF
\`\`\`

## Client API

\`\`\`typescript
const bridge = await connect();           // Connect to bridge server
const docs = await bridge.documents();    // List connected Word documents
const doc = docs[0];                       // Get document handle

// Execute Office.js code in the document
const result = await doc.executeJs<T>(\`
  // context is available (Word.RequestContext)
  // Must call await context.sync() to flush operations
  // Return JSON-serializable value
\`);

await bridge.close();                      // Disconnect (documents stay connected)
\`\`\`

## Key Principles

1. **Small scripts**: Each script should do ONE thing
2. **Always sync**: Call \`await context.sync()\` after queueing operations
3. **Return data**: Convert Word objects to plain data before returning
4. **Check state**: Log document state at the end to verify success

## Common Patterns

### Search and Insert

\`\`\`javascript
const results = context.document.body.search("Section 2", { matchCase: true });
results.load("items");
await context.sync();

if (results.items.length > 0) {
  results.items[0].insertText(" (amended)", Word.InsertLocation.after);
  await context.sync();
}
return { found: results.items.length };
\`\`\`

### Get Document Text

\`\`\`javascript
const body = context.document.body;
body.load("text");
await context.sync();
return body.text;
\`\`\`

### Get Paragraphs

\`\`\`javascript
const paragraphs = context.document.body.paragraphs;
paragraphs.load("items/text");
await context.sync();
return paragraphs.items.map(p => p.text);
\`\`\`

## Error Recovery

If a script fails, the document state is preserved. Take corrective action:

\`\`\`bash
cd office-bridge && bun x tsx <<'EOF'
import { connect } from "@/client.js";

const bridge = await connect();
const docs = await bridge.documents();
console.log("Connected documents:", docs.map(d => d.filename));

if (docs.length > 0) {
  const text = await docs[0].executeJs(\`
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text.slice(0, 500);
  \`);
  console.log("Document preview:", text);
}

await bridge.close();
EOF
\`\`\`
```

### install.sh

```bash
#!/bin/bash
set -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "Installing dependencies..."
npm install

echo "Installing Office Add-in dev certificates..."
npx office-addin-dev-certs install

echo "Building TypeScript..."
npm run build

echo ""
echo "Setup complete!"
echo ""
echo "Next steps:"
echo "1. Start the bridge server: ./server.sh"
echo "2. Sideload the add-in in Word (see README)"
```

### server.sh

```bash
#!/bin/bash

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "Starting office-bridge server..."
bun x tsx src/server.ts
```

## Development Setup

### Prerequisites

- Node.js 18+ (or Bun)
- Microsoft Word Desktop (Mac or Windows)
- `office-addin-dev-certs` for localhost HTTPS

### First-Time Setup

```bash
# From office-bridge directory
./install.sh
```

This installs dependencies, generates HTTPS certificates, and builds TypeScript.

### Starting the Bridge

```bash
./server.sh
```

Or let the client auto-launch when connecting.

### Sideloading the Add-in

**Windows**:
1. Open Word
2. Insert → Get Add-ins → My Add-ins → Upload My Add-in
3. Browse to `addin/manifest.xml`

**Mac**:
1. Open Word
2. Insert → Add-ins → My Add-ins
3. Click folder icon, select `manifest.xml`

The add-in persists after sideloading—no need to re-upload each session.

## Example: Live Document Editing

```bash
cd office-bridge && bun x tsx <<'EOF'
import { connect } from "@/client.js";

const bridge = await connect();
const docs = await bridge.documents();
const doc = docs.find(d => d.filename.includes("contract"));

// Insert text after a heading
const result = await doc.executeJs(`
  const body = context.document.body;
  const searchResults = body.search("Section 2.1", { matchCase: true });
  searchResults.load("items");
  await context.sync();

  if (searchResults.items.length > 0) {
    const range = searchResults.items[0].getRange("End");
    range.insertText(" (amended)", Word.InsertLocation.after);
    await context.sync();
  }

  return { found: searchResults.items.length };
`);

console.log("Result:", result);
await bridge.close();
EOF
```

## Limitations and Non-Goals (v1)

These are explicitly **out of scope** for v1:

- [ ] Word Online support (different security model, API differences)
- [ ] Excel/PowerPoint support (future: expand to full Office suite)
- [ ] Convenience wrapper methods (search, replace helpers)
- [ ] MCP tool interface
- [ ] Event subscriptions (notify Claude when document changes)
- [ ] Document content streaming for large documents
- [ ] Undo/redo integration

## Open Questions

1. **Port Selection**: Fixed port (3847) or dynamic with discovery? Fixed is simpler for v1.

2. **Multi-user**: If multiple Claude Code sessions connect, should they see all documents or be isolated? Single bridge shared for v1.

3. **Token Persistence**: Should the auth token persist across server restarts, or regenerate each time? Regenerate for security.

## References

- [Word JavaScript API Overview](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)
- [Word.Range class](https://learn.microsoft.com/en-us/javascript/api/word/word.range?view=word-js-preview)
- [Word.TrackedChange class](https://learn.microsoft.com/en-us/javascript/api/word/word.trackedchange?view=word-js-preview)
- [Sideloading Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- [office-addin-dev-certs](https://www.npmjs.com/package/office-addin-dev-certs)
- [dev-browser skill](https://github.com/anthropics/claude-code/tree/main/skills/dev-browser) (pattern inspiration)
