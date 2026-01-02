---
name: office-bridge
description: Execute JavaScript in live Office applications via Office.js API. Supports Word, Excel, PowerPoint, and Outlook. Use when users ask to edit open Office documents, manipulate spreadsheets, edit presentations, work with emails, or automate any Office workflow. Trigger phrases include "edit my Word doc", "update the spreadsheet", "add a slide", "check my email", or any live Office document manipulation.
---

# Office Bridge

Execute Office.js JavaScript in open Office applications. Maintains persistent WebSocket connections to Office add-ins.

## Quick Start

1. Run `./office-bridge/install.sh` (first time only)
2. Start bridge server: `./office-bridge/server.sh &`
3. Start app dev server (see ports below)
4. Open Office app, click Office Bridge add-in

For sideloading help, see [references/setup.md](references/setup.md).

## Connection Pattern

```typescript
import { connect } from "./src/client.js";

const bridge = await connect();
const sessions = await bridge.sessions();           // All apps
const docs = await bridge.documents();              // Word (with helpers)
const workbooks = await bridge.excel();             // Excel
const presentations = await bridge.powerpoint();   // PowerPoint
const mail = await bridge.outlook();               // Outlook

// Execute code in any session
const result = await session.executeJs(`
  // App-specific code here
  await context.sync();
  return data;
`);

await bridge.close();
```

## Ports

| Component | Port |
|-----------|------|
| Bridge Server | 3847 |
| Word | 3000 |
| Excel | 3001 |
| PowerPoint | 3002 |
| Outlook | 3003 |

Start a dev server: `cd office-bridge/addins/<app> && npm run dev-server &`

## Troubleshooting

- **No sessions**: Open Office app → click Office Bridge add-in
- **Add-in missing**: See [references/setup.md](references/setup.md)
- **Execution failed**: Check Office.js syntax, use try/catch

## References

- [references/setup.md](references/setup.md) - Sideloading and installation steps
- [references/word.md](references/word.md) - Word patterns and accessibility helpers
- [references/excel.md](references/excel.md) - Excel workbook and range operations
- [references/powerpoint.md](references/powerpoint.md) - PowerPoint slide and shape operations
- [references/outlook.md](references/outlook.md) - Outlook mail and calendar access
