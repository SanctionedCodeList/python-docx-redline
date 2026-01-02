# Office Bridge Multi-App Specification

## Overview

Extend Office Bridge from Word-only to support all four major Office applications:
- **Word** (existing)
- **Excel** (new)
- **PowerPoint** (new)
- **Outlook** (new)

## Design Principles

1. **Shared infrastructure** - One bridge server handles all apps
2. **V1 first** - Raw JS execution before helpers
3. **Progressive enhancement** - Add accessibility helpers incrementally
4. **Consistent API** - Same client patterns across apps

## Architecture

### Current (Word-only)

```
Client → Bridge Server → Word Add-in → Word Document
```

### Target (Multi-App)

```
                    ┌─→ Word Add-in ────→ Word Document
                    │
Client → Bridge ────┼─→ Excel Add-in ───→ Excel Workbook
         Server     │
                    ├─→ PowerPoint Add-in → Presentation
                    │
                    └─→ Outlook Add-in ──→ Mailbox/Items
```

## Server Changes

### Session Registry

Rename "document" concept to "session" to be app-agnostic:

```typescript
interface Session {
  id: string;
  app: 'word' | 'excel' | 'powerpoint' | 'outlook';
  socket: WebSocket;
  name: string;  // Document name, workbook name, etc.
  executionQueue: ExecutionQueue;
}
```

### Registration Message

Update to include app type:

```typescript
interface RegisterMessage {
  type: 'register';
  sessionId: string;
  app: 'word' | 'excel' | 'powerpoint' | 'outlook';
  name: string;
}
```

### Endpoints

Keep existing endpoints, add app filtering:

- `GET /` - Server info (unchanged)
- `GET /sessions` - All connected sessions (replaces `/documents`)
- `GET /sessions?app=excel` - Filter by app
- `POST /execute` - Execute code (unchanged, routes by sessionId)

## Add-in Structure

### Monorepo Layout

```
office-bridge/
├── src/                      # Shared server + client
│   ├── index.ts             # Bridge server
│   ├── client.ts            # Client library
│   ├── types.ts             # Shared types
│   ├── registry.ts          # Session registry
│   └── accessibility/       # Word accessibility (existing)
│
├── addins/                   # Per-app add-ins
│   ├── word/                # Existing, moved from addin/
│   │   ├── manifest.xml
│   │   ├── src/taskpane/
│   │   └── package.json
│   │
│   ├── excel/
│   │   ├── manifest.xml
│   │   ├── src/taskpane/
│   │   └── package.json
│   │
│   ├── powerpoint/
│   │   ├── manifest.xml
│   │   ├── src/taskpane/
│   │   └── package.json
│   │
│   └── outlook/
│       ├── manifest.xml
│       ├── src/taskpane/
│       └── package.json
│
└── docs/
```

### Shared Taskpane Code

Extract common WebSocket handling to a shared module:

```typescript
// src/addin-common/bridge-client.ts
export class BridgeClient {
  constructor(app: AppType, getName: () => Promise<string>);
  connect(): Promise<void>;
  onExecute(handler: (code: string) => Promise<unknown>): void;
}
```

Each add-in only needs app-specific execution:

```typescript
// addins/excel/src/taskpane/taskpane.ts
import { BridgeClient } from '../../../src/addin-common/bridge-client';

const client = new BridgeClient('excel', async () => {
  return await Excel.run(ctx => ctx.workbook.name);
});

client.onExecute(async (code) => {
  return await Excel.run(async (context) => {
    const fn = new Function('context', 'Excel', 'Office',
      `return (async () => { ${code} })();`);
    return await fn(context, Excel, Office);
  });
});
```

## App-Specific Details

### Excel

**Execution context**: `Excel.run()`

**Key objects to expose**:
- `context.workbook` - The workbook
- `context.workbook.worksheets` - Sheet collection
- `Range`, `Table`, `Chart`, `PivotTable`

**Future helpers** (not V1):
- Range-based refs: `sheet:0/range:A1:B10`
- Formula-aware editing
- Table manipulation

### PowerPoint

**Execution context**: `PowerPoint.run()`

**Key objects to expose**:
- `context.presentation` - The presentation
- `context.presentation.slides` - Slide collection
- `Slide`, `Shape`, `TextFrame`

**Future helpers** (not V1):
- Slide-based refs: `slide:3/shape:2`
- Shape positioning
- Text frame editing

### Outlook

**Execution context**: `Office.context.mailbox` (no `.run()` pattern)

**Key objects to expose**:
- `Office.context.mailbox` - Mailbox access
- `Office.context.mailbox.item` - Current item (compose/read)

**Compose vs Read modes**:
- Compose: Can modify subject, body, recipients
- Read: Read-only access to item

**Future helpers** (not V1):
- Message manipulation
- Calendar operations
- Contact lookup

## Client API

### Current (Word)

```typescript
const client = new OfficeBridgeClient();
const doc = await client.connect();
await doc.executeJs('...');
await doc.getTree();
```

### Target (Multi-App)

```typescript
const client = new OfficeBridgeClient();

// Connect to specific app
const word = await client.word();
const excel = await client.excel();
const ppt = await client.powerpoint();
const outlook = await client.outlook();

// Or get all sessions
const sessions = await client.sessions();
const excelSessions = await client.sessions({ app: 'excel' });

// Common interface
await excel.executeJs('context.workbook.name');
await ppt.executeJs('context.presentation.title');

// App-specific helpers (future)
await word.getTree();      // Word accessibility tree
await excel.getRange('A1:B10');  // Excel range helpers
```

## Manifests

Each app needs its own manifest with:
- Unique add-in ID
- Correct `<Hosts>` declaration
- Appropriate permissions

### Excel Manifest Highlights

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
<Requirements>
  <Sets>
    <Set Name="ExcelApi" MinVersion="1.1"/>
  </Sets>
</Requirements>
```

### PowerPoint Manifest Highlights

```xml
<Hosts>
  <Host Name="Presentation" />
</Hosts>
<Requirements>
  <Sets>
    <Set Name="PowerPointApi" MinVersion="1.1"/>
  </Sets>
</Requirements>
```

### Outlook Manifest Highlights

```xml
<Hosts>
  <Host Name="Mailbox" />
</Hosts>
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1"/>
  </Sets>
</Requirements>
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <!-- Read mode -->
</ExtensionPoint>
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <!-- Compose mode -->
</ExtensionPoint>
```

## Dev Server Ports

Each add-in needs its own dev server port:

| App        | Dev Server Port | Purpose           |
|------------|-----------------|-------------------|
| Word       | 3000            | Existing          |
| Excel      | 3001            | New               |
| PowerPoint | 3002            | New               |
| Outlook    | 3003            | New               |
| Bridge     | 3847            | Unchanged (shared)|

## Success Criteria

### V1 (Raw JS Interface)

- [ ] Bridge server accepts connections from all four apps
- [ ] Each app can register and execute arbitrary Office.js code
- [ ] Client can list sessions filtered by app
- [ ] Client can execute code on any session
- [ ] Console forwarding works for all apps

### V2 (Basic Helpers)

- [ ] Excel: Range reading/writing helpers
- [ ] PowerPoint: Slide/shape enumeration
- [ ] Outlook: Message property access
- [ ] Common ref system design

### V3 (Rich Helpers)

- [ ] Excel: Full accessibility tree with formula support
- [ ] PowerPoint: Shape editing with positioning
- [ ] Outlook: Compose mode helpers
- [ ] Cross-app operation patterns
