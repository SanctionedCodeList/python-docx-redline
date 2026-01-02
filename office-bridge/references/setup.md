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

**Word:**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
ln -sf "$(pwd)/office-bridge/addins/word/manifest.xml" \
  ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/19d3eb15-55fe-45a1-a33f-89b66b5aae3b.manifest.xml
```

**Excel:**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
ln -sf "$(pwd)/office-bridge/addins/excel/manifest.xml" \
  ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/a1b2c3d4-e5f6-7890-abcd-ef1234567890.manifest.xml
```

**PowerPoint:**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
ln -sf "$(pwd)/office-bridge/addins/powerpoint/manifest.xml" \
  ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/b2c3d4e5-f6a7-8901-bcde-f23456789012.manifest.xml
```

**Outlook:**
```bash
mkdir -p ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef
ln -sf "$(pwd)/office-bridge/addins/outlook/manifest.xml" \
  ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/c3d4e5f6-a7b8-9012-cdef-345678901234.manifest.xml
```

After sideloading, **quit the app completely** (Cmd+Q) and reopen.

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
