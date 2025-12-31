---
name: office-bridge-setup
description: "Step-by-step guide to install and configure the Office Bridge Word add-in. Use when helping users set up the add-in for the first time, troubleshoot connection issues, or sideload the manifest into Microsoft Word."
---

# Office Bridge Setup Guide

Complete instructions for setting up the Office Bridge Word add-in.

## Prerequisites

### Required Software

| Software | Version | Check Command |
|----------|---------|---------------|
| Node.js | 18+ | `node --version` |
| npm | 9+ | `npm --version` |
| Microsoft Word | 2016+ or Microsoft 365 | - |

### Install Node.js (if needed)

**macOS:**
```bash
brew install node
```

**Windows:**
```powershell
# Download from https://nodejs.org or use winget:
winget install OpenJS.NodeJS.LTS
```

**Linux (Ubuntu/Debian):**
```bash
curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
sudo apt-get install -y nodejs
```

## Installation

### Step 1: Navigate to Office Bridge Directory

```bash
cd /path/to/python_docx_redline/office-bridge
```

### Step 2: Run the Install Script

```bash
./install.sh
```

This script:
1. Installs bridge server dependencies (`npm install`)
2. Installs add-in dependencies (`cd addin && npm install`)
3. Installs HTTPS dev certificates for Office add-in development

**If install.sh fails**, run the steps manually:

```bash
# Bridge server
npm install

# Add-in
cd addin
npm install
cd ..

# Dev certificates (required for HTTPS)
npx office-addin-dev-certs install
```

### Step 3: Build the Add-in

```bash
cd addin
npm run build
```

This compiles the TypeScript and bundles the add-in.

## Starting the Servers

You need **two terminals** running:

### Terminal 1: Bridge Server

```bash
cd office-bridge
./server.sh
# Or: npm start
```

You should see:
```
Office Bridge Server running on wss://localhost:3847
Waiting for Word add-in connections...
```

### Terminal 2: Add-in Dev Server

```bash
cd office-bridge/addin
npm run dev-server
```

You should see:
```
webpack compiled successfully
Server running at https://localhost:3000
```

**Keep both terminals running** while using the add-in.

## Sideloading the Add-in into Word

The add-in must be "sideloaded" (manually loaded) into Word. The process differs by platform.

### macOS

1. Open **Word**
2. Open any document (or create new)
3. Go to **Insert** → **Add-ins** → **My Add-ins**
4. Click **Manage My Add-ins** (bottom of dialog)
5. Click **Add a custom add-in** → **Add from file...**
6. Navigate to: `office-bridge/addin/manifest.xml`
7. Click **Open**

The add-in appears in the ribbon. Click **Office Bridge** to open the taskpane.

### Windows

1. Open **Word**
2. Go to **Insert** → **Get Add-ins** (or **My Add-ins**)
3. Click **MY ADD-INS** tab
4. Click **Upload My Add-in** (upper right)
5. Click **Browse** and select: `office-bridge/addin/manifest.xml`
6. Click **Upload**

### Word for the Web (Microsoft 365)

1. Open Word at **https://www.office.com**
2. Open a document
3. Go to **Insert** → **Add-ins** → **More Add-ins**
4. Click **MY ADD-INS** → **Upload My Add-in**
5. Upload `manifest.xml`

**Note:** Word for Web requires the add-in server to be accessible. For local development, you may need to use a tunnel service like ngrok.

### Verify Sideloading Succeeded

After sideloading:
1. Look for **Office Bridge** in the ribbon (Home tab)
2. Click it to open the taskpane
3. The taskpane should show "Connected" with a green status dot
4. If it shows "Disconnected", ensure both servers are running

## Verifying the Connection

### Check the Taskpane

The Office Bridge taskpane shows:
- **Status**: Connected (green) / Disconnected (red)
- **Document ID**: Unique ID for the current document
- **Last Activity**: Timestamp of last operation

### Test from Python

```python
import asyncio
import websockets
import json

async def test_connection():
    async with websockets.connect("wss://localhost:3847") as ws:
        # List connected documents
        await ws.send(json.dumps({"type": "list"}))
        response = await ws.recv()
        print(json.loads(response))

asyncio.run(test_connection())
```

If successful, you'll see a list of connected Word documents.

## Troubleshooting

### Add-in Won't Load

**Symptom:** "Add-in failed to load" or manifest errors

**Solutions:**
1. Ensure `npm run dev-server` is running in `addin/`
2. Check https://localhost:3000 loads in browser (accept certificate warning)
3. Rebuild: `cd addin && npm run build`
4. Clear Word add-in cache:
   - **macOS**: `rm -rf ~/Library/Containers/com.microsoft.Word/Data/Library/Caches`
   - **Windows**: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`

### "Disconnected" Status

**Symptom:** Taskpane shows red "Disconnected" status

**Solutions:**
1. Ensure bridge server is running: `./server.sh`
2. Check server output for errors
3. Verify port 3847 is not in use: `lsof -i :3847`
4. Check WebSocket connection in browser dev tools (Word taskpane → right-click → Inspect)

### Certificate Errors

**Symptom:** "Unable to get local issuer certificate" or HTTPS errors

**Solutions:**
```bash
# Reinstall dev certificates
npx office-addin-dev-certs install --force

# Verify certificates exist
npx office-addin-dev-certs verify
```

If using macOS, you may need to trust the certificate in Keychain Access.

### "Office.js is not loaded"

**Symptom:** Script errors about Office.js

**Solutions:**
1. Ensure you're opening the add-in IN Word, not in a browser
2. Rebuild: `cd addin && npm run build`
3. Reload add-in: Remove and re-add from My Add-ins

### Port Already in Use

**Symptom:** "EADDRINUSE" error on port 3847 or 3000

**Solutions:**
```bash
# Find what's using the port
lsof -i :3847
lsof -i :3000

# Kill the process
kill -9 <PID>
```

### Word Doesn't Show My Add-ins Option

**Symptom:** Can't find sideload option

**Solutions:**
1. Ensure you have Word 2016 or later
2. For Microsoft 365, ensure you're signed in
3. Try: **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs** → Enable sideloading

## Development Workflow

Once set up, the typical workflow is:

1. Start both servers (keep running in background)
2. Open Word and load the add-in
3. Open a document to edit
4. Use Python/Claude to connect and send commands

### Auto-reload During Development

The dev server supports hot reload:
```bash
cd addin
npm run dev-server  # Watches for changes
```

Edit files in `addin/src/` and the add-in reloads automatically.

## Uninstalling

### Remove the Add-in

1. **Word** → **Insert** → **Add-ins** → **My Add-ins**
2. Right-click **Office Bridge** → **Remove**

### Remove Dev Certificates

```bash
npx office-addin-dev-certs uninstall
```

### Remove Node Modules

```bash
cd office-bridge
rm -rf node_modules
cd addin
rm -rf node_modules
```

## See Also

- [SKILL.md](./SKILL.md) - Office Bridge overview
- [tree-building.md](./tree-building.md) - Building accessibility trees
- [editing.md](./editing.md) - Ref-based editing operations
