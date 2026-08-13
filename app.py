from __future__ import annotations

import os
from pathlib import Path
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from excel_processing import (
    SourceColumn,
    extract_date_from_filename,
    factory_selection,
    process_workbook,
    read_source_columns,
)
from user_defaults import save_default, user_default_selection

APP_TITLE = "ScienceGuide Excel Cleaner"


class ColumnSelector(tk.Toplevel):
    def __init__(
        self,
        parent,
        kind: str,
        columns: list[SourceColumn],
        selected: list[int],
        on_done,
    ):
        super().__init__(parent)
        self.kind = kind
        self.columns = columns
        self.by_index = {c.index: c for c in columns}
        self.selected = list(selected)
        self.on_done = on_done

        self.title(f"{kind} columns")
        self.geometry("940x610")
        self.minsize(800, 540)
        self.transient(parent)
        self.grab_set()

        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text=f"{kind} output columns",
            font=("TkDefaultFont", 12, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            outer,
            text=(
                "Possible columns are detected from the selected Excel file. "
                "A saved personal default is applied automatically; the factory default always remains available."
            ),
        ).pack(anchor="w", pady=(2, 12))

        self.summary = ttk.Label(outer)
        self.summary.pack(anchor="w", pady=(0, 10))

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(2, weight=1)
        body.rowconfigure(1, weight=1)

        ttk.Label(body, text="Included in output (order shown)").grid(row=0, column=0, sticky="w")
        ttk.Label(body, text="Other columns detected in this file").grid(row=0, column=2, sticky="w")

        self.included = tk.Listbox(body, selectmode=tk.EXTENDED, exportselection=False)
        self.available = tk.Listbox(body, selectmode=tk.EXTENDED, exportselection=False)
        inc_scroll = ttk.Scrollbar(body, orient="vertical", command=self.included.yview)
        av_scroll = ttk.Scrollbar(body, orient="vertical", command=self.available.yview)
        self.included.configure(yscrollcommand=inc_scroll.set)
        self.available.configure(yscrollcommand=av_scroll.set)

        self.included.grid(row=1, column=0, sticky="nsew", pady=(4, 8))
        inc_scroll.grid(row=1, column=0, sticky="nse", pady=(4, 8))
        self.available.grid(row=1, column=2, sticky="nsew", pady=(4, 8))
        av_scroll.grid(row=1, column=2, sticky="nse", pady=(4, 8))

        controls = ttk.Frame(body, padding=(10, 40))
        controls.grid(row=1, column=1, sticky="n")
        ttk.Button(controls, text="← Add", command=self.add_columns, width=12).pack(pady=3)
        ttk.Button(controls, text="Remove →", command=self.remove_columns, width=12).pack(pady=3)
        ttk.Separator(controls, orient="horizontal").pack(fill="x", pady=10)
        ttk.Button(controls, text="↑ Up", command=lambda: self.move(-1), width=12).pack(pady=3)
        ttk.Button(controls, text="↓ Down", command=lambda: self.move(1), width=12).pack(pady=3)

        presets = ttk.LabelFrame(outer, text="Defaults", padding=(8, 6))
        presets.pack(fill="x", pady=(10, 0))
        ttk.Button(presets, text="Restore factory default", command=self.reset_factory).pack(side="left")
        self.restore_my_button = ttk.Button(presets, text="Restore my default", command=self.reset_user)
        self.restore_my_button.pack(side="left", padx=(8, 0))
        ttk.Button(presets, text="Set current as my default", command=self.save_as_default).pack(side="left", padx=(8, 0))

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(10, 0))
        ttk.Button(bottom, text="Include all detected", command=self.select_all).pack(side="left")
        ttk.Button(bottom, text="Clear", command=self.clear_all).pack(side="left", padx=(8, 0))
        ttk.Button(bottom, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bottom, text="Done", command=self.done).pack(side="right", padx=(0, 8))

        self.note = ttk.Label(outer, foreground="#6a6a6a")
        self.note.pack(fill="x", pady=(10, 0))
        self.refresh()

    def refresh(self):
        self.included.delete(0, tk.END)
        for idx in self.selected:
            self.included.insert(tk.END, self.by_index[idx].display)

        selected_set = set(self.selected)
        self.available_indices = [c.index for c in self.columns if c.index not in selected_set]
        self.available.delete(0, tk.END)
        for idx in self.available_indices:
            self.available.insert(tk.END, self.by_index[idx].display)

        _, missing, aliases = factory_selection(self.kind, self.columns)
        self.summary.configure(
            text=(
                f"{len(self.columns)} columns detected in the source file · "
                f"{len(self.selected)} currently included · "
                f"{len(self.available_indices)} available to add"
            )
        )
        pieces = []
        user_selection, user_missing = user_default_selection(self.kind, self.columns)
        if user_selection is not None:
            pieces.append("A personal default is saved")
            self.restore_my_button.state(["!disabled"])
            if user_missing:
                pieces.append(f"{len(user_missing)} saved field(s) are unavailable in this source")
        else:
            self.restore_my_button.state(["disabled"])
        if aliases:
            pieces.append(f"{len(aliases)} renamed factory field(s) matched automatically")
        if missing:
            pieces.append(f"{len(missing)} factory default field(s) are not present in this source")
        self.note.configure(
            text=" · ".join(pieces)
            if pieces
            else "All factory default fields are present in this source file."
        )

    def add_columns(self):
        chosen = [self.available_indices[i] for i in self.available.curselection()]
        self.selected.extend(chosen)
        self.refresh()

    def remove_columns(self):
        positions = list(self.included.curselection())
        for pos in reversed(positions):
            del self.selected[pos]
        self.refresh()

    def move(self, direction: int):
        positions = list(self.included.curselection())
        if len(positions) != 1:
            return
        pos = positions[0]
        new = pos + direction
        if new < 0 or new >= len(self.selected):
            return
        self.selected[pos], self.selected[new] = self.selected[new], self.selected[pos]
        self.refresh()
        self.included.selection_set(new)
        self.included.see(new)

    def reset_factory(self):
        self.selected, _, _ = factory_selection(self.kind, self.columns)
        self.refresh()

    def reset_user(self):
        selected, _ = user_default_selection(self.kind, self.columns)
        if selected is None:
            messagebox.showinfo(APP_TITLE, f"No personal {self.kind} default has been saved yet.", parent=self)
            return
        self.selected = selected
        self.refresh()

    def save_as_default(self):
        if not self.selected:
            messagebox.showerror(APP_TITLE, "Select at least one column before saving a default.", parent=self)
            return
        save_default(self.kind, self.columns, self.selected)
        self.on_done(list(self.selected))
        messagebox.showinfo(
            APP_TITLE,
            f"The current {self.kind} selection and order are now your personal default.\n\n"
            "It will be applied automatically the next time you choose this type of file.",
            parent=self,
        )
        self.refresh()

    def select_all(self):
        self.selected = [c.index for c in self.columns]
        self.refresh()

    def clear_all(self):
        self.selected = []
        self.refresh()

    def done(self):
        self.on_done(list(self.selected))
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("820x560")
        self.minsize(740, 520)

        self.paths = {"VER": tk.StringVar(), "ANG": tk.StringVar()}
        self.columns: dict[str, list[SourceColumn]] = {"VER": [], "ANG": []}
        self.selected: dict[str, list[int]] = {"VER": [], "ANG": []}
        self.column_status = {"VER": tk.StringVar(value="No file selected"), "ANG": tk.StringVar(value="No file selected")}
        self.export_date = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.status = tk.StringVar(value="Choose a VER export, an ANG export, or both to begin.")
        self.last_output_folder: Path | None = None

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self, padding=18)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text=APP_TITLE, font=("TkDefaultFont", 18, "bold")).pack(anchor="w")
        ttk.Label(
            main,
            text=(
                "Select a VER export, an ANG export, or both. The app reads columns automatically, applies "
                "the factory default selection, and lets you add, remove or reorder any detected fields."
            ),
            wraplength=770,
        ).pack(anchor="w", pady=(3, 18))

        files = ttk.LabelFrame(main, text="1. Source files", padding=12)
        files.pack(fill="x")
        files.columnconfigure(1, weight=1)

        for row, kind in enumerate(("VER", "ANG")):
            label = "VER — Veranstaltungen" if kind == "VER" else "ANG — Angebote"
            ttk.Label(files, text=label, width=22).grid(row=row, column=0, sticky="w", pady=5)
            ttk.Entry(files, textvariable=self.paths[kind], state="readonly").grid(row=row, column=1, sticky="ew", padx=(6, 8), pady=5)
            ttk.Button(files, text="Choose file…", command=lambda k=kind: self.choose_file(k)).grid(row=row, column=2, pady=5)
            clear_button = ttk.Button(files, text="Clear", command=lambda k=kind: self.clear_file(k), width=8)
            clear_button.grid(row=row, column=3, padx=(6, 0), pady=5)
            setattr(self, f"clear_{kind.lower()}_button", clear_button)
            clear_button.state(["disabled"])

        settings = ttk.LabelFrame(main, text="2. Output", padding=12)
        settings.pack(fill="x", pady=(12, 0))
        settings.columnconfigure(1, weight=1)

        ttk.Label(settings, text="Export date", width=22).grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(settings, textvariable=self.export_date, width=18).grid(row=0, column=1, sticky="w", padx=(6, 8), pady=5)
        ttk.Label(settings, text="YYYY_MM_DD (detected from filename)").grid(row=0, column=2, sticky="w")

        ttk.Label(settings, text="Output folder", width=22).grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(settings, textvariable=self.output_folder, state="readonly").grid(row=1, column=1, sticky="ew", padx=(6, 8), pady=5)
        ttk.Button(settings, text="Choose folder…", command=self.choose_output_folder).grid(row=1, column=2, pady=5)

        cols = ttk.LabelFrame(main, text="3. Columns detected from the selected files", padding=12)
        cols.pack(fill="x", pady=(12, 0))
        cols.columnconfigure(1, weight=1)
        for row, kind in enumerate(("VER", "ANG")):
            ttk.Label(cols, text=f"{kind} columns", width=22).grid(row=row, column=0, sticky="w", pady=5)
            ttk.Label(cols, textvariable=self.column_status[kind]).grid(row=row, column=1, sticky="w", padx=(6, 8), pady=5)
            button = ttk.Button(cols, text="Edit columns…", command=lambda k=kind: self.edit_columns(k))
            button.grid(row=row, column=2, pady=5)
            setattr(self, f"edit_{kind.lower()}_button", button)
            button.state(["disabled"])

        actions = ttk.Frame(main)
        actions.pack(fill="x", pady=(18, 0))
        self.process_button = ttk.Button(actions, text="Process Excel files", command=self.process)
        self.process_button.pack(side="left")
        self.open_folder_button = ttk.Button(actions, text="Open output folder", command=self.open_output_folder)
        self.open_folder_button.pack(side="left", padx=(8, 0))
        self.open_folder_button.state(["disabled"])

        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=(18, 10))
        ttk.Label(main, textvariable=self.status, wraplength=760).pack(anchor="w")

    def choose_file(self, kind: str):
        path = filedialog.askopenfilename(
            title=f"Choose {kind} Excel export",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            columns = read_source_columns(path)
            factory_selected, missing, aliases = factory_selection(kind, columns)
            user_selected, user_missing = user_default_selection(kind, columns)
            selected = user_selected if user_selected is not None else factory_selected
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return

        self.paths[kind].set(path)
        self.columns[kind] = columns
        self.selected[kind] = selected
        self._update_column_status(kind, missing, aliases, user_missing if user_selected is not None else None)
        getattr(self, f"edit_{kind.lower()}_button").state(["!disabled"])
        getattr(self, f"clear_{kind.lower()}_button").state(["!disabled"])

        detected = extract_date_from_filename(path)
        if detected:
            if not self.export_date.get():
                self.export_date.set(detected)
            elif self.export_date.get() != detected:
                self.status.set(
                    f"Note: {kind} filename contains {detected}, while export date is {self.export_date.get()}. "
                    "You can edit the export date if needed."
                )

        if not self.output_folder.get():
            self.output_folder.set(str(Path(path).parent))

    def clear_file(self, kind: str):
        self.paths[kind].set("")
        self.columns[kind] = []
        self.selected[kind] = []
        self._update_column_status(kind)
        getattr(self, f"edit_{kind.lower()}_button").state(["disabled"])
        getattr(self, f"clear_{kind.lower()}_button").state(["disabled"])
        active = [k for k in ("VER", "ANG") if self.paths[k].get()]
        if active:
            self.status.set(f"Ready to process {' and '.join(active)}.")
        else:
            self.status.set("Choose a VER export, an ANG export, or both to begin.")

    def _update_column_status(self, kind: str, missing=None, aliases=None, user_missing=None):
        if not self.columns[kind]:
            self.column_status[kind].set("No file selected")
            return
        if missing is None or aliases is None:
            _, missing, aliases = factory_selection(kind, self.columns[kind])
        detected = len(self.columns[kind])
        selected = len(self.selected[kind])
        text = f"{detected} detected · {selected} selected · {detected - selected} available to add"
        notes = []
        if aliases:
            notes.append(f"{len(aliases)} renamed default(s) matched")
        if missing:
            notes.append(f"{len(missing)} default(s) unavailable")
        user_selected, saved_missing = user_default_selection(kind, self.columns[kind])
        if user_selected is not None:
            notes.append("personal default saved")
            if user_missing is None:
                user_missing = saved_missing
            if user_missing:
                notes.append(f"{len(user_missing)} saved field(s) unavailable")
        if notes:
            text += " · " + " · ".join(notes)
        self.column_status[kind].set(text)

    def edit_columns(self, kind: str):
        if not self.columns[kind]:
            return

        def done(indices):
            self.selected[kind] = indices
            self._update_column_status(kind)

        ColumnSelector(self, kind, self.columns[kind], self.selected[kind], done)

    def choose_output_folder(self):
        folder = filedialog.askdirectory(title="Choose output folder")
        if folder:
            self.output_folder.set(folder)

    def process(self):
        active_kinds = [kind for kind in ("VER", "ANG") if self.paths[kind].get()]
        if not active_kinds:
            messagebox.showerror(APP_TITLE, "Choose at least one source file: VER, ANG, or both.")
            return

        for kind in active_kinds:
            if not self.selected[kind]:
                messagebox.showerror(APP_TITLE, f"Select at least one {kind} column.")
                return

        export_date = self.export_date.get().strip()
        if not export_date:
            messagebox.showerror(APP_TITLE, "Enter or confirm the export date (YYYY_MM_DD).")
            return

        output_folder = self.output_folder.get().strip()
        if not output_folder:
            messagebox.showerror(APP_TITLE, "Choose an output folder.")
            return

        self.process_button.state(["disabled"])
        self.status.set("Processing…")
        self.update_idletasks()

        try:
            folder = Path(output_folder)
            results = {}
            outputs = {}
            for kind in active_kinds:
                suffix = kind.lower()
                output = folder / f"{export_date}_{suffix}_mod.xlsx"
                results[kind] = process_workbook(self.paths[kind].get(), self.selected[kind], output)
                outputs[kind] = output
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Processing failed:\n\n{exc}")
            self.status.set("Processing failed.")
            self.process_button.state(["!disabled"])
            return

        self.last_output_folder = folder
        self.open_folder_button.state(["!disabled"])
        self.process_button.state(["!disabled"])
        summaries = [
            f"{kind}: {results[kind]['output_rows']} rows / {results[kind]['output_columns']} columns"
            for kind in active_kinds
        ]
        self.status.set(f"Finished. {'; '.join(summaries)}. Saved to {folder}.")
        created = "\n".join(outputs[kind].name for kind in active_kinds)
        messagebox.showinfo(
            APP_TITLE,
            f"File{'s' if len(active_kinds) > 1 else ''} created successfully:\n\n{created}",
        )

    def open_output_folder(self):
        if not self.last_output_folder:
            return
        path = str(self.last_output_folder)
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", path])
            elif os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not open output folder:\n\n{exc}")


if __name__ == "__main__":
    App().mainloop()
