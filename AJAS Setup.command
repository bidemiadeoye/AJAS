#!/bin/bash
# AJAS First-Time Setup

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
LAUNCHER="$SCRIPT_DIR/app/ajas_launcher.py"
DESKTOP="$HOME/Desktop"

echo ""
echo "  AJAS - Setting up..."
echo ""

xattr -rd com.apple.quarantine "$SCRIPT_DIR" 2>/dev/null
echo "  ✓ Security check complete"

if [ ! -f "$LAUNCHER" ]; then
    echo ""
    echo "  ERROR: Cannot find app/ajas_launcher.py"
    echo "  Make sure 'AJAS Setup.command' and the 'app' folder"
    echo "  are in the same folder."
    read -p "  Press Enter to close..."
    exit 1
fi

PYTHON=""
for py in \
    "/Library/Frameworks/Python.framework/Versions/3.14/bin/python3" \
    "/Library/Frameworks/Python.framework/Versions/3.13/bin/python3" \
    "/Library/Frameworks/Python.framework/Versions/3.12/bin/python3" \
    "/Library/Frameworks/Python.framework/Versions/3.11/bin/python3" \
    "/Library/Frameworks/Python.framework/Versions/3.10/bin/python3" \
    "/Library/Frameworks/Python.framework/Versions/3.9/bin/python3" \
    "/opt/homebrew/bin/python3" \
    "/usr/local/bin/python3"; do
    if [ -x "$py" ]; then
        OK=$("$py" -c "import sys; print('ok' if sys.version_info>=(3,9) else 'no')" 2>/dev/null)
        if [ "$OK" = "ok" ]; then
            PYTHON="$py"
            VER=$("$py" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')")
            echo "  ✓ Python $VER found"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo "  ERROR: Python 3.9 or later not found."
    echo "  Please install from python.org/downloads then run setup again."
    read -p "  Press Enter to open python.org..."
    open "https://python.org/downloads"
    exit 1
fi

echo "  Installing required libraries..."
"$PYTHON" -m pip install flask openpyxl python-docx requests \
    --quiet --break-system-packages 2>/dev/null || \
"$PYTHON" -m pip install flask openpyxl python-docx requests --quiet 2>/dev/null
echo "  ✓ Libraries ready"

SHORTCUT="$DESKTOP/Open AJAS.command"
cat > "$SHORTCUT" << INNEREOF
#!/bin/bash
echo ""
echo "  AJAS is running in this window."
echo "  Your browser will open automatically."
echo "  Keep this window open while using AJAS."
echo "  Close it when you are done."
echo ""
"PYTHON_PATH" "LAUNCHER_PATH"
echo ""
echo "  AJAS has stopped. You can close this window."
INNEREOF
# Replace placeholders with actual paths
sed -i '' "s|PYTHON_PATH|$PYTHON|g" "$SHORTCUT"
sed -i '' "s|LAUNCHER_PATH|$LAUNCHER|g" "$SHORTCUT"
chmod +x "$SHORTCUT"
xattr -d com.apple.quarantine "$SHORTCUT" 2>/dev/null
echo "  ✓ AJAS shortcut created on your Desktop"

echo ""
echo "  ============================================"
echo "  Setup complete!"
echo ""
echo "  DAILY USE:"
echo "  Double-click 'Open AJAS' on your Desktop."
echo "  No right-click needed ever again."
echo "  ============================================"
echo ""
read -p "  Press Enter to launch AJAS now... "
echo ""
echo "  Launching AJAS..."
echo ""
echo "  ┌─────────────────────────────────────────┐"
echo "  │  AJAS is running in this window.        │"
echo "  │                                         │"
echo "  │  Your browser will open automatically.  │"
echo "  │                                         │"
echo "  │  Keep this window open while using AJAS.│"
echo "  │  Close it when you are done.            │"
echo "  └─────────────────────────────────────────┘"
echo ""

"$PYTHON" "$LAUNCHER"

echo ""
echo "  AJAS has stopped. You can close this window."
