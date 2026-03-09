#!/bin/bash
# Install GF Sort add-in for Word (macOS)
#
# This script downloads the production manifest and sideloads it into Word.
# No Node.js, no server, no technical setup required.
#
# Usage:
#   curl -sL https://raw.githubusercontent.com/tenbobnote/gf-sort/main/install.sh | bash

set -e

echo ""
echo "=== Installing GF Sort for Word ==="
echo ""

# Quit Word if running
if pgrep -x "Microsoft Word" > /dev/null 2>&1; then
    echo "Quitting Word..."
    osascript -e 'tell application "Microsoft Word" to quit saving yes' 2>/dev/null || true
    sleep 2
fi

# Create sideload directory
WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
mkdir -p "$WEF_DIR"

# Download production manifest
echo "Downloading manifest..."
curl -sL "https://raw.githubusercontent.com/tenbobnote/gf-sort/main/manifest-prod.xml" -o "$WEF_DIR/manifest.xml"

echo ""
echo "=== Installation Complete ==="
echo ""
echo "To use GF Sort:"
echo "  1. Open Word"
echo "  2. Open the Resource Guide document"
echo "  3. On the Home tab, click 'GF Sort' (far right of the ribbon)"
echo ""
echo "If you don't see the button:"
echo "  • Go to Insert > Add-ins > My Add-ins > GF Sort"
echo "  • Or try: Insert > Get Add-ins > MY ADD-INS > Shared Folder > GF Sort"
echo ""
