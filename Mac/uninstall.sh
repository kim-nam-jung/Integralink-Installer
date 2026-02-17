#!/bin/bash

# Integralink Mac Uninstaller

WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
MANIFEST="$WEF_DIR/manifest.xml"

echo ""
echo "========================================================"
echo "   Integralink Word Add-in Uninstaller for Mac 🗑️"
echo "========================================================"
echo ""

if [ -f "$MANIFEST" ]; then
    rm "$MANIFEST"
    echo "   [SUCCESS] manifest.xml removed. ✅"
else
    echo "   [INFO] manifest.xml not found using this path."
    echo "          (It might have been already removed)"
fi

echo ""
echo "========================================================"
echo "   Uninstallation Complete!"
echo "========================================================"
echo ""
