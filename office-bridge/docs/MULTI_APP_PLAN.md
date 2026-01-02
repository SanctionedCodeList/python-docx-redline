# Office Bridge Multi-App Development Plan

## Phase 1: Infrastructure Refactoring

Prepare the codebase for multiple apps without breaking Word functionality.

### 1.1 Reorganize Directory Structure

**Current:**
```
office-bridge/
├── addin/          # Word add-in
└── src/            # Server + Word accessibility
```

**Target:**
```
office-bridge/
├── addins/
│   ├── shared/     # Common taskpane code
│   └── word/       # Word add-in (moved from addin/)
├── src/
│   ├── server/     # Bridge server
│   ├── client/     # Client library
│   └── word/       # Word accessibility helpers
└── docs/
```

**Tasks:**
- [ ] Create `addins/` directory
- [ ] Move `addin/` → `addins/word/`
- [ ] Create `addins/shared/` for common code
- [ ] Update imports and build scripts
- [ ] Verify Word still works

### 1.2 Generalize Server Types

Update `src/types.ts` for multi-app support:

- [ ] Add `AppType = 'word' | 'excel' | 'powerpoint' | 'outlook'`
- [ ] Add `app` field to `RegisterMessage`
- [ ] Rename document-specific types to session-generic names
- [ ] Update `registry.ts` to track app type per session

### 1.3 Extract Shared Taskpane Code

Create `addins/shared/bridge-client.ts`:

- [ ] Extract WebSocket connection logic from Word taskpane
- [ ] Create `BridgeClient` class with app-agnostic interface
- [ ] Parameterize app type and name getter
- [ ] Update Word taskpane to use shared client
- [ ] Test Word still works

### 1.4 Update Client Library

Enhance `src/client.ts`:

- [ ] Add `sessions()` method with app filter
- [ ] Add `word()`, `excel()`, `powerpoint()`, `outlook()` convenience methods
- [ ] Maintain backward compatibility with existing API
- [ ] Add app type to session info

---

## Phase 2: Excel Add-in (V1)

First new app - validates the multi-app pattern.

### 2.1 Create Excel Add-in Structure

```
addins/excel/
├── manifest.xml
├── package.json
├── tsconfig.json
├── webpack.config.js
├── assets/
│   └── icon-*.png
└── src/
    ├── taskpane/
    │   ├── taskpane.ts
    │   ├── taskpane.html
    │   └── taskpane.css
    └── commands/
        ├── commands.ts
        └── commands.html
```

**Tasks:**
- [ ] Copy Word add-in structure as template
- [ ] Create Excel manifest with unique ID and Excel hosts
- [ ] Update package.json for Excel dependencies
- [ ] Create Excel-specific taskpane.ts using `Excel.run()`
- [ ] Expose Excel objects (context, Excel, Office)
- [ ] Set dev server to port 3001
- [ ] Create Excel icons (can reuse Word icons initially)

### 2.2 Excel Manifest

Key differences from Word:
- [ ] Unique add-in ID (new UUID)
- [ ] `<Host Name="Workbook" />`
- [ ] `ExcelApi` requirement set
- [ ] ReadWriteDocument permission

### 2.3 Excel Taskpane Implementation

```typescript
// Key execution pattern
await Excel.run(async (context) => {
  const fn = new Function('context', 'Excel', 'Office',
    `return (async () => { ${code} })();`);
  return await fn(context, Excel, Office);
});
```

**Tasks:**
- [ ] Implement `Excel.run()` execution wrapper
- [ ] Forward console messages
- [ ] Handle errors appropriately
- [ ] Register with bridge using app='excel'

### 2.4 Test Excel Integration

- [ ] Sideload Excel add-in
- [ ] Verify registration with bridge server
- [ ] Execute simple code: `context.workbook.name`
- [ ] Execute range operations: `context.workbook.getSelectedRange()`
- [ ] Verify console forwarding

---

## Phase 3: PowerPoint Add-in (V1)

### 3.1 Create PowerPoint Add-in Structure

Same structure as Excel, port 3002.

**Tasks:**
- [ ] Copy Excel add-in structure as template
- [ ] Create PowerPoint manifest with unique ID
- [ ] Update for PowerPoint hosts and requirements
- [ ] Create PowerPoint taskpane using `PowerPoint.run()`
- [ ] Set dev server to port 3002

### 3.2 PowerPoint Manifest

- [ ] Unique add-in ID
- [ ] `<Host Name="Presentation" />`
- [ ] `PowerPointApi` requirement set

### 3.3 PowerPoint Taskpane

```typescript
await PowerPoint.run(async (context) => {
  const fn = new Function('context', 'PowerPoint', 'Office',
    `return (async () => { ${code} })();`);
  return await fn(context, PowerPoint, Office);
});
```

### 3.4 Test PowerPoint Integration

- [ ] Sideload PowerPoint add-in
- [ ] Execute: `context.presentation.slides.getCount()`
- [ ] Verify console forwarding

---

## Phase 4: Outlook Add-in (V1)

Outlook is different - uses `Office.context.mailbox` instead of `.run()`.

### 4.1 Create Outlook Add-in Structure

Same structure, port 3003.

**Tasks:**
- [ ] Copy structure as template
- [ ] Create Outlook manifest with mailbox hosts
- [ ] Handle both read and compose modes
- [ ] Create Outlook taskpane (different execution pattern)

### 4.2 Outlook Manifest

More complex - needs extension points for different modes:

- [ ] Unique add-in ID
- [ ] `<Host Name="Mailbox" />`
- [ ] `Mailbox` requirement set
- [ ] `MessageReadCommandSurface` extension point
- [ ] `MessageComposeCommandSurface` extension point

### 4.3 Outlook Taskpane

Different execution pattern:

```typescript
// No .run() pattern - direct access
const executeCode = async (code: string) => {
  const fn = new Function('Office', 'mailbox', 'item',
    `return (async () => { ${code} })();`);
  return await fn(
    Office,
    Office.context.mailbox,
    Office.context.mailbox.item
  );
};
```

### 4.4 Test Outlook Integration

- [ ] Sideload Outlook add-in
- [ ] Test in read mode: `item.subject`
- [ ] Test in compose mode: `item.subject.getAsync()`
- [ ] Verify console forwarding

---

## Phase 5: Documentation & Polish

### 5.1 Update SKILL.md

- [ ] Document multi-app setup
- [ ] Add per-app sideloading instructions
- [ ] Document client API for each app
- [ ] Add troubleshooting per app

### 5.2 Create Helper Scripts

- [ ] `install-all.sh` - Install all add-in dependencies
- [ ] `start-all.sh` - Start all dev servers
- [ ] `sideload-all.sh` - Sideload all add-ins

### 5.3 Update Root Package.json

- [ ] Add workspace configuration
- [ ] Add scripts for managing all add-ins
- [ ] Document npm commands

---

## Future Phases (Post-V1)

### Phase 6: Excel Accessibility Helpers

- [ ] Design Excel ref system (sheet/range addressing)
- [ ] Build Excel tree builder
- [ ] Implement range-based editing
- [ ] Add formula awareness

### Phase 7: PowerPoint Accessibility Helpers

- [ ] Design PowerPoint ref system (slide/shape addressing)
- [ ] Build slide/shape tree
- [ ] Implement shape editing with positioning

### Phase 8: Outlook Helpers

- [ ] Message property helpers
- [ ] Compose mode body manipulation
- [ ] Calendar item support

---

## Implementation Order

```
Week 1:
├── Phase 1.1: Directory reorganization
├── Phase 1.2: Generalize types
├── Phase 1.3: Extract shared taskpane code
└── Phase 1.4: Update client library

Week 2:
├── Phase 2: Excel add-in (complete)
└── Phase 3: PowerPoint add-in (complete)

Week 3:
├── Phase 4: Outlook add-in (complete)
└── Phase 5: Documentation & polish
```

## Validation Checklist

After each phase, verify:

- [ ] Word add-in still works (no regression)
- [ ] Bridge server starts without errors
- [ ] New add-in sideloads successfully
- [ ] Basic code execution works
- [ ] Console forwarding works
- [ ] Client can connect and execute
