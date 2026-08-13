# ScienceGuide Excel Cleaner

Small desktop app that replaces the existing R cleanup script with a simple GUI.

## What it does

- Opens one **VER / Veranstaltungen** export and one **ANG / Angebote** export.
- Uses the supplied legacy templates as the immutable **factory default** for which columns are included and in which order.
- Lets the user add, remove and reorder any columns found in the current export.
- Automatically matches the old `Kurzbeschreibung_*` factory fields to the newer `Untertitel_*` export fields.
- Handles duplicate source headers by column position. For ANG, the factory `ÖV-Haltestelle` intentionally selects the later route/tour occurrence.
- Removes the technical field-code row when detected (`ID`, `TITEL_D`, `GEAENDERT`, etc.). This is done for both VER and ANG because the supplied 2026 samples contain that row in both files.
- Sorts rows by `Geändert am`, newest first.
- Exports `YYYY_MM_DD_ver_mod.xlsx` and `YYYY_MM_DD_ang_mod.xlsx`.
- Does not use hard-coded user paths.

## Factory defaults

The old `Template_ver.xls` and `Template_ang.xls` files were decoded during development. Their column definitions are now stored in `factory_columns.py`, so the finished app does **not** need the `.xls` templates.

- VER factory columns: **51**
- ANG factory columns: **59**

If a factory column no longer exists in a newer export, the UI reports it as unavailable rather than failing.

## Run from source on Windows

Python 3.10+ is sufficient; Python 3.12 is recommended.

Double-click:

`run_windows.bat`

It creates `.venv`, installs `openpyxl`, and launches the app.

## Build a Windows desktop package

Double-click:

`build_windows.bat`

The packaged application is written to `dist/ScienceGuide Excel Cleaner/`.

## Build the macOS `.app`

PyInstaller must build the macOS application **on macOS**. From a Mac:

```bash
chmod +x build_mac.sh
./build_mac.sh
```

This creates:

- `dist/ScienceGuide Excel Cleaner.app`
- `dist/ScienceGuide-Excel-Cleaner-macOS.zip`

The receiving Mac does not need Python when using the packaged `.app`.

## Build the Mac app from Windows using GitHub Actions

The included workflow `.github/workflows/build-macos.yml` runs on a GitHub-hosted macOS runner.

1. Put this project in a GitHub repository.
2. Open **Actions** in GitHub.
3. Select **Build macOS app**.
4. Choose **Run workflow**.
5. Download the `ScienceGuide-Excel-Cleaner-macOS` artifact when the run finishes.

This is the easiest route when development happens on Windows but the final application is for macOS.

## Signing / notarization

The build produced by the workflow is unsigned. For routine use on managed Macs, the final release should ideally be signed with an Apple Developer ID certificate and notarized. That can be added to the workflow later without changing the Excel-processing logic.
