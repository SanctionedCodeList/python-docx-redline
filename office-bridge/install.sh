#!/bin/bash
set -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "=== Office Bridge Setup ==="
echo ""

# Install bridge server dependencies
echo "Installing bridge server dependencies..."
npm install

# Install add-in dependencies for each app
for app in word excel powerpoint outlook; do
  if [ -d "addins/$app" ]; then
    echo "Installing $app add-in dependencies..."
    cd "addins/$app"
    npm install
    cd "$SCRIPT_DIR"
  fi
done

# Install dev certificates for HTTPS
echo "Installing Office Add-in dev certificates..."
npx office-addin-dev-certs install

echo ""
echo "=== Setup Complete ==="
echo ""
echo "Next steps:"
echo "1. Start bridge server: ./server.sh &"
echo "2. Start app dev server: cd addins/<app> && npm run dev-server &"
echo "3. Sideload the add-in (see references/setup.md for instructions)"
echo ""
