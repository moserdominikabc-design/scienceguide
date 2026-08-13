# ScienceGuide Excel Cleaner

Small desktop app that replaces the existing R cleanup script with a simple GUI.

## What it does

- Opens a **VER / Veranstaltungen** export, an **ANG / Angebote** export, or both in the same run. Neither file type requires the other.
- Reads the **possible columns directly from each Excel file when it is selected**. New fields in future exports therefore appear automatically in the UI without an app update.
- Uses the supplied legacy templates only as the immutable **factory default** for which of those detected columns are initially included and in which order.
- Lets the user add, remove and reorder any columns detected in the current export.
- Automatically matches the old `Kurzbeschreibung_*` factory fields to the newer `Untertitel_*` export fields.
- Handles duplicate source headers by column position. For ANG, the factory `ÖV-Haltestelle` intentionally selects the later route/tour occurrence.
- Removes the technical field-code row when detected (`ID`, `TITEL_D`, `GEAENDERT`, etc.). This is done for both VER and ANG because the supplied 2026 samples contain that row in both files.
- Sorts rows by `Geändert am`, newest first.
- Exports only the file type(s) selected for that run: `YYYY_MM_DD_ver_mod.xlsx`, `YYYY_MM_DD_ang_mod.xlsx`, or both.
- Does not use hard-coded user paths.

## Factory defaults

The old `Template_ver.xls` and `Template_ang.xls` files were decoded during development. Their column definitions are now stored in `factory_columns.py`, so the finished app does **not** need the `.xls` templates.

- VER factory columns: **51**
- ANG factory columns: **59**

If a factory column no longer exists in a newer export, the UI reports it as unavailable rather than failing. If a newer export introduces additional columns, those columns automatically appear under **Other columns detected in this file** and can be added by the user.

The two sources of truth are intentionally separate:

- **Current Excel file** = every column the user is allowed to choose from.
- **Factory definitions** = the default subset and default output order only.

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

PyInstaller must build the macOS application **on macOS**. For a local build, Python 3.12 is required. From a Mac:

```bash
chmod +x build_mac.sh
./build_mac.sh
```

The local build script deliberately requires `python3.12`; it will not silently package the app with an older system Python installation.

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

The GitHub workflow explicitly uses Python 3.12, runs syntax/unit/import checks, and then packages the `.app`. This is the easiest route when development happens on Windows but the final application is for macOS.

## Signing / notarization

The build produced by the workflow is unsigned. For routine use on managed Macs, the final release should ideally be signed with an Apple Developer ID certificate and notarized. That can be added to the workflow later without changing the Excel-processing logic.
