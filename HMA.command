#!/bin/bash
# HMA launcher (macOS): double-click in Finder, or run ./HMA.command in the Terminal.
# Needs Python 3.10 or newer with Tk, e.g. the installer from https://www.python.org/downloads/macos/
cd "$(dirname "$0")" || exit 1
SCRIPT=$(ls -1 script/HMA_v*.py skript/HMA_v*.py 2>/dev/null | grep -vi draft | sort -V | tail -n 1)
if [ -z "$SCRIPT" ]; then
  echo "HMA script not found (folder 'script')."
  read -r -p "Press Enter to close this window. " _
  exit 1
fi
for PY in python3.13 python3.12 python3.11 python3.10 python3 \
          /Library/Frameworks/Python.framework/Versions/Current/bin/python3 /usr/local/bin/python3 /opt/homebrew/bin/python3; do
  if command -v "$PY" >/dev/null 2>&1 && "$PY" -c 'import sys, tkinter; sys.exit(0 if sys.version_info >= (3, 10) else 1)' >/dev/null 2>&1; then
    "$PY" "$SCRIPT" "$@"
    exit $?
  fi
done
echo
echo "No suitable Python found. HMA needs Python 3.10 or newer with tkinter (Tk)."
echo "Install Python from https://www.python.org/downloads/macos/ (this version includes Tk),"
echo "then double-click HMA.command again."
echo "(The Python 3.9 that ships with Apple's developer tools is too old.)"
echo
read -r -p "Press Enter to close this window. " _
exit 1
