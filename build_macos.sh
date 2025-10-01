#!/bin/zsh

set -euo pipefail

SCRIPT_DIR=$(cd -- "$(dirname -- "$0")" && pwd)
cd "$SCRIPT_DIR"

python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt

# Build a single app-like bundle with console hidden and GUI-friendly file picker
python3 -m PyInstaller \
  --onefile \
  --windowed \
  --name "BlueBook Counter Upper" \
  bluebook_count.py

echo "\nBuild complete. App is at: dist/BlueBook Counter Upper"


