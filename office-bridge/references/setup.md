# Office Bridge Setup

## Automated Setup

Run the install script (first time only):

```bash
./office-bridge/install.sh
```

This installs npm dependencies and Office Add-in dev certificates.

## Sideloading Add-ins

Each Office app requires sideloading its manifest once per machine.

### macOS

**Important:** Use `cp` to copy manifests, not symlinks. Symlinks don't work reliably on Mac.

**Word:**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
cp "$(pwd)/office-bridge/addins/word/manifest.xml" \
  ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

**Excel:**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
cp "$(pwd)/office-bridge/addins/excel/manifest.xml" \
  ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml
```

**PowerPoint:**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
cp "$(pwd)/office-bridge/addins/powerpoint/manifest.xml" \
  ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml
```

**Outlook:**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef
cp "$(pwd)/office-bridge/addins/outlook/manifest.xml" \
  ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/manifest.xml
```

**Note:** After updating the add-in code, you must re-copy the manifest for changes to take effect.

After sideloading, **quit the app completely** (Cmd+Q) and reopen.

**Finding the add-in in the app:**
- Go to **Home** tab → **Add-ins** (or **Insert** → **Add-ins** in some versions)
- The "Office Bridge" add-in appears under the **Developer** or **Custom Add-ins** section
- Click it to open the taskpane

**Troubleshooting macOS sideloading:**
- If the add-in doesn't appear, ensure the dev server is running (`npm run dev-server` in the add-in folder)
- Check that the manifest.xml points to the correct localhost port (Word: 3000, Excel: 3001, PowerPoint: 3002, Outlook: 3003)
- Try removing and re-adding the symlink if the add-in fails to load

**Removing a sideloaded add-in:**
```bash
rm ~/Library/Containers/com.microsoft.<App>/Data/Documents/wef/*.manifest.xml
```

### Windows

1. Open the Office application
2. Go to **Insert** → **Get Add-ins** → **My Add-ins** → **Upload My Add-in**
3. Browse to `office-bridge/addins/<app>/manifest.xml`

## Starting Servers

```bash
# Bridge server (required)
./office-bridge/server.sh &

# App dev servers (start the ones you need)
cd office-bridge/addins/word && npm run dev-server &        # Port 3000
cd office-bridge/addins/excel && npm run dev-server &       # Port 3001
cd office-bridge/addins/powerpoint && npm run dev-server &  # Port 3002
cd office-bridge/addins/outlook && npm run dev-server &     # Port 3003
```

## Verifying Connection

```bash
cd office-bridge && bun x tsx <<'EOF'
import { connect } from "./src/client.js";
const bridge = await connect();
const sessions = await bridge.sessions();
console.log("Connected sessions:", sessions.map(s => `[${s.app}] ${s.filename}`));
await bridge.close();
EOF
```
