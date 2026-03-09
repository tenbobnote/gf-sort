#!/bin/bash
# Setup script for the German Foundations Word Add-in (macOS)
#
# This installs dependencies, generates HTTPS certificates, and
# sideloads the add-in manifest into Word.
#
# Usage: cd word-addin && bash setup.sh

set -e

echo "=== German Foundations Word Add-in Setup ==="
echo ""

# Check for Node.js
if ! command -v node &> /dev/null; then
    echo "ERROR: Node.js is required but not installed."
    echo "Install it from https://nodejs.org/ or: brew install node"
    exit 1
fi

echo "1. Installing dependencies..."
npm install

echo ""
echo "2. Generating HTTPS certificates..."
npx office-addin-dev-certs install
echo "   Certificates installed at ~/.office-addin-dev-certs/"

echo ""
echo "3. Sideloading manifest into Word..."
WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
mkdir -p "$WEF_DIR"
cp manifest.xml "$WEF_DIR/"
echo "   Manifest copied to $WEF_DIR/"

echo ""
echo "=== Setup Complete ==="
echo ""
echo "To start the add-in:"
echo "  1. Quit Word completely (Cmd+Q)"
echo "  2. Run:  cd word-addin && npm start"
echo "  3. Open Word (server MUST be running before Word launches)"
echo "  4. Click Add-ins in the Home ribbon > MY ADD-INS > GF Sort"
echo ""
echo "Keep the terminal running while using the add-in."
echo ""
echo "IMPORTANT: If 'npm start' fails with EADDRINUSE, run:"
echo "  lsof -ti :3000 | xargs kill"
echo "  Then try 'npm start' again."
