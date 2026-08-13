#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"

if ! command -v python3.12 >/dev/null 2>&1; then
  echo "Python 3.12 is required for a local macOS build."
  echo "Use the included GitHub Actions workflow if you do not want to install Python 3.12 on this Mac."
  exit 1
fi

PYTHON_BIN="$(command -v python3.12)"
echo "Building with: $($PYTHON_BIN --version)"

rm -rf .venv-build build dist
"$PYTHON_BIN" -m venv .venv-build
source .venv-build/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt

python -m py_compile app.py excel_processing.py factory_columns.py user_defaults.py
python -m unittest -v
python -c 'import tkinter; import app; print("Tkinter/app import OK")'

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
