#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"

python3 -m venv .venv-build
source .venv-build/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt

python -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "ScienceGuide Excel Cleaner" \
  --osx-bundle-identifier "ch.scnat.scienceguide.excelcleaner" \
  app.py

rm -f "dist/ScienceGuide-Excel-Cleaner-macOS.zip"
ditto -c -k --sequesterRsrc --keepParent \
  "dist/ScienceGuide Excel Cleaner.app" \
  "dist/ScienceGuide-Excel-Cleaner-macOS.zip"

echo "Built: dist/ScienceGuide Excel Cleaner.app"
echo "Zip:   dist/ScienceGuide-Excel-Cleaner-macOS.zip"
