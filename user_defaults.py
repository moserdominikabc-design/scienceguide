from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Sequence

from excel_processing import SourceColumn

APP_DIR_NAME = "ScienceGuide Excel Cleaner"
DEFAULTS_FILENAME = "column_defaults.json"


def user_config_dir() -> Path:
    """Return a writable per-user config directory on Windows/macOS/Linux."""
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    elif os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / APP_DIR_NAME


def defaults_path() -> Path:
    return user_config_dir() / DEFAULTS_FILENAME


def _column_keys(columns: Sequence[SourceColumn]) -> dict[int, tuple[str, int]]:
    """Map source index -> (header, duplicate occurrence), preserving source order."""
    occurrences: dict[str, int] = {}
    keys: dict[int, tuple[str, int]] = {}
    for column in columns:
        occurrences[column.header] = occurrences.get(column.header, 0) + 1
        keys[column.index] = (column.header, occurrences[column.header])
    return keys


def selection_to_specs(columns: Sequence[SourceColumn], selected: Sequence[int]) -> list[dict[str, object]]:
    """Serialize a selection using header + duplicate occurrence, not fragile column positions."""
    keys = _column_keys(columns)
    specs: list[dict[str, object]] = []
    for index in selected:
        if index not in keys:
            continue
        header, occurrence = keys[index]
        specs.append({"header": header, "occurrence": occurrence})
    return specs


def specs_to_selection(columns: Sequence[SourceColumn], specs: Sequence[dict[str, object]]) -> tuple[list[int], list[str]]:
    """Resolve a stored default against columns in the newly selected workbook."""
    lookup: dict[tuple[str, int], int] = {}
    occurrences: dict[str, int] = {}
    folded_lookup: dict[tuple[str, int], int] = {}
    folded_occurrences: dict[str, int] = {}

    for column in columns:
        occurrences[column.header] = occurrences.get(column.header, 0) + 1
        lookup[(column.header, occurrences[column.header])] = column.index

        folded = column.header.casefold()
        folded_occurrences[folded] = folded_occurrences.get(folded, 0) + 1
        folded_lookup[(folded, folded_occurrences[folded])] = column.index

    selected: list[int] = []
    missing: list[str] = []
    used: set[int] = set()
    for spec in specs:
        header = str(spec.get("header", ""))
        try:
            occurrence = int(spec.get("occurrence", 1))
        except (TypeError, ValueError):
            occurrence = 1
        index = lookup.get((header, occurrence))
        if index is None:
            index = folded_lookup.get((header.casefold(), occurrence))
        if index is None or index in used:
            label = header if occurrence == 1 else f"{header} (occurrence {occurrence})"
            missing.append(label)
            continue
        selected.append(index)
        used.add(index)
    return selected, missing


def load_defaults(path: Path | None = None) -> dict[str, list[dict[str, object]]]:
    path = path or defaults_path()
    if not path.exists():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}

    result: dict[str, list[dict[str, object]]] = {}
    if not isinstance(raw, dict):
        return result
    for kind in ("VER", "ANG"):
        value = raw.get(kind)
        if isinstance(value, list):
            clean = [item for item in value if isinstance(item, dict) and item.get("header")]
            if clean:
                result[kind] = clean
    return result


def save_default(
    kind: str,
    columns: Sequence[SourceColumn],
    selected: Sequence[int],
    path: Path | None = None,
) -> Path:
    kind = kind.upper()
    if kind not in {"VER", "ANG"}:
        raise ValueError(f"Unknown export kind: {kind}")
    if not selected:
        raise ValueError("At least one column is required to save a default.")

    path = path or defaults_path()
    data = load_defaults(path)
    data[kind] = selection_to_specs(columns, selected)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def delete_default(kind: str, path: Path | None = None) -> bool:
    kind = kind.upper()
    path = path or defaults_path()
    data = load_defaults(path)
    if kind not in data:
        return False
    del data[kind]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return True


def user_default_selection(
    kind: str,
    columns: Sequence[SourceColumn],
    path: Path | None = None,
) -> tuple[list[int] | None, list[str]]:
    data = load_defaults(path)
    specs = data.get(kind.upper())
    if not specs:
        return None, []
    selected, missing = specs_to_selection(columns, specs)
    return selected, missing
