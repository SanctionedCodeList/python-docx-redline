#!/bin/bash
set -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "=== Office Bridge Setup ==="
echo ""

# Install bridge server dependencies
echo "Installing bridge server dependencies..."
npm install

# Install add-in dependencies
if [ -d "addin" ]; then
  echo "Installing add-in dependencies..."
  cd addin
  npm install
  cd ..
fi

# Install dev certificates for HTTPS
echo "Installing Office Add-in dev certificates..."
npx office-addin-dev-certs install

echo ""
echo "=== Setup Complete ==="
echo ""
echo "Next steps:"
echo "1. Start the bridge server: ./server.sh"
echo "2. Sideload the add-in in Word (Insert → Add-ins → My Add-ins)"
