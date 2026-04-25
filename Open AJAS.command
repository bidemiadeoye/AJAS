#!/bin/bash
# AJAS Daily Launcher
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
LAUNCHER="$SCRIPT_DIR/app/ajas_launcher.py"

if [ ! -f "$LAUNCHER" ]; then
    osascript -e "display dialog \"Cannot find AJAS files.\n\nThis launcher must stay in the same folder as the 'app' folder.\n\nRun 'AJAS Setup.command' to create a fresh Desktop shortcut.\" buttons {\"OK\"} with title \"AJAS\""
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
        if [ "$OK" = "ok" ]; then PYTHON="$py"; break; fi
    fi
done

if [ -z "$PYTHON" ]; then
    osascript -e 'display dialog "Python 3.9+ not found. Please run AJAS Setup first." buttons {"OK"} with title "AJAS"'
    exit 1
fi

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
