#!/bin/bash

# Integralink Mac Uninstaller

# Target Paths (Add-in Manifest & Cache)
MANIFEST_PATH="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml"
CACHE_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Library/Caches/com.microsoft.Word/Wef"

echo "   [1/2] Removing manifest file..."
if [ -f "$MANIFEST_PATH" ]; then
    rm "$MANIFEST_PATH"
    echo "         Removed: $MANIFEST_PATH"
else
    echo "         Not found: $MANIFEST_PATH"
fi

echo ""
echo "   [2/2] Clearing Web Extension Framework (WEF) cache..."
if [ -d "$CACHE_DIR" ]; then
    rm -rf "$CACHE_DIR"
    echo "         Cleared cache: $CACHE_DIR"
else
    echo "         Cache directory not found."
fi

echo ""
echo "========================================================"
echo "   Uninstallation Complete!"
echo "========================================================"
echo ""
