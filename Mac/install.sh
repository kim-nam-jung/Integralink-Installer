#!/bin/bash

# Integralink Mac Installer
# Copies the manifest to the Word WEF folder for sideloading.

WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"

echo ""
echo "========================================================"
echo "   Integralink Word Add-in Installer for Mac 🍎"
echo "========================================================"
echo ""

# Check if manifest.xml exists in the current directory
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
MANIFEST="$SCRIPT_DIR/manifest.xml"

if [ ! -f "$MANIFEST" ]; then
    echo "   [ERROR] manifest.xml not found!"
    echo "   Please make sure manifest.xml is in the same folder."
    exit 1
fi

echo "   [1/2] Creating installation directory..."
mkdir -p "$WEF_DIR"

if [ -d "$WEF_DIR" ]; then
    echo "         Directory ready."
else
    echo "   [ERROR] Failed to create directory: $WEF_DIR"
    exit 1
fi

echo ""
echo "   [2/2] Copying manifest.xml..."
cp "$MANIFEST" "$WEF_DIR/"

if [ $? -eq 0 ]; then
    echo "         manifest.xml copied successfully! ✅"
else
    echo "   [ERROR] Failed to copy manifest.xml."
    exit 1
fi

echo ""
echo "========================================================"
echo "   Installation Complete! 🎉"
echo "========================================================"
echo ""
echo "   1. Restart Microsoft Word (Close completely and open). 🔄"
echo "   2. Go to 'Insert' tab > 'Add-ins' > 'My Add-ins'. 🧩"
echo "   3. Look under 'Developer Add-ins' or 'Shared Folder'."
echo "      (Or check the dropdown arrow next to My Add-ins)"
echo ""
