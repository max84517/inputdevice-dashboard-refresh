"""
Main UI for Shipment Report Dashboard Refresh.
Dark mode tkinter application.
"""
import os
import sys
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# Adjust sys.path so sibling packages are importable
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from src.config import load_config, save_config
from src.ingestion.fetcher import get_suppliers, fetch_supplier_files
from src.processing.transformer import get_fy_sheets, consolidate_suppliers
from src.export.exporter import save_history, get_available_fy, merge_for_powerbi

# ─── Dark theme colours ───────────────────────────────────────────────────────
BG = "#1e1e1e"
BG2 = "#2d2d2d"
BG3 = "#3c3c3c"
FG = "#d4d4d4"
ACCENT = "#0e639c"
ACCENT_HOVER = "#1177bb"
BTN_FG = "#ffffff"
ENTRY_BG = "#3c3c3c"
CHECK_BG = "#2d2d2d"
RED = "#f44747"
GREEN = "#4ec9b0"

FONT = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 13, "bold")
FONT_SMALL = ("Segoe UI", 9)


DEFAULT_KEEP_COLUMNS = [
    "GTK Suppliers", "SPM (Project Owner)", "Category", "Segment", "Series",
    "Production Year", "Platforms", "Product", "Size", "Color", "Location",
    "ODM", "Region", "HP/ODM Part#", "HP Cost", "Unit Rebate", "Q'ty",
    "Rebate Amount", "ODM Cost", "Spending Amount", "Actual Spending",
    "Month", "Year", "FY", "Supplier Part#",
]


def _style_btn(btn, bg=ACCENT, fg=BTN_FG, hover=ACCENT_HOVER):
    btn.configure(bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
                  relief="flat", cursor="hand2", font=FONT, bd=0, padx=8, pady=4)
    btn.bind("<Enter>", lambda e: btn.configure(bg=hover))
    btn.bind("<Leave>", lambda e: btn.configure(bg=bg))


class ShipmentApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Shipment Report Dashboard Refresh")
        self.root.configure(bg=BG)
        self.root.geometry("900x700")
        self.root.minsize(760, 580)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # State
        self.cfg = load_config()
        self.supplier_vars: dict[str, tk.BooleanVar] = {}
        # {supplier: [Path, ...]}  full sorted list
        self.supplier_file_lists: dict[str, list] = {}
        # {supplier: int}  current index into supplier_file_lists
        self.supplier_file_idx: dict[str, int] = {}
        # {supplier: StringVar}  displayed file name
        self.supplier_file_labels: dict[str, tk.StringVar] = {}

        # Keep-columns config
        self.keep_columns: list[str] = self.cfg.get("keep_columns", list(DEFAULT_KEEP_COLUMNS))

        # Log queue for real-time updates from background thread
        self._log_queue: queue.Queue = queue.Queue()

        # Base path for data folder (next to src/)
        self.project_root = Path(__file__).parent.parent.parent
        self.source_data_dir = str(self.project_root / "data" / "source_data")
        self.history_dir = str(self.project_root / "data" / "history")
        self.output_dir = str(self.project_root / "data" / "output")

        self._build_ui()
        self._restore_state()
        self._poll_log()  # start real-time log polling

    # ─────────────────────────────────────────────────────────────────────────
    # UI Construction
    # ─────────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top bar: path selector ───────────────────────────────────────────
        top = tk.Frame(self.root, bg=BG, pady=10, padx=14)
        top.pack(fill="x")

        tk.Label(top, text="Spending & Rebate Reports Folder", bg=BG, fg=FG,
                 font=FONT_BOLD).grid(row=0, column=0, sticky="w")

        self.path_var = tk.StringVar(value=self.cfg.get("base_path", ""))
        path_entry = tk.Entry(top, textvariable=self.path_var, bg=ENTRY_BG, fg=FG,
                              insertbackground=FG, font=FONT, relief="flat", bd=3)
        path_entry.grid(row=1, column=0, sticky="ew", padx=(0, 8), ipady=4)

        browse_btn = tk.Button(top, text="Browse…", command=self._browse_path)
        _style_btn(browse_btn)
        browse_btn.grid(row=1, column=1)
        top.columnconfigure(0, weight=1)

        # ── Supplier list section ────────────────────────────────────────────
        mid = tk.Frame(self.root, bg=BG, padx=14)
        mid.pack(fill="both", expand=True)

        # Left panel: supplier checkboxes (fixed width)
        left = tk.LabelFrame(mid, text=" Suppliers ", bg=BG2, fg=FG, font=FONT_BOLD,
                              bd=1, relief="solid", padx=6, pady=6)
        left.pack(side="left", fill="both", padx=(0, 6))
        left.pack_propagate(False)
        left.configure(width=180)

        # Scrollable checkbox area
        self.supplier_canvas = tk.Canvas(left, bg=BG2, highlightthickness=0)
        sup_scroll = tk.Scrollbar(left, orient="vertical", command=self.supplier_canvas.yview)
        self.supplier_canvas.configure(yscrollcommand=sup_scroll.set)
        sup_scroll.pack(side="right", fill="y")
        self.supplier_canvas.pack(side="left", fill="both", expand=True)
        self.supplier_inner = tk.Frame(self.supplier_canvas, bg=BG2)
        self._sup_window = self.supplier_canvas.create_window((0, 0), window=self.supplier_inner, anchor="nw")
        self.supplier_inner.bind("<Configure>", lambda e: self.supplier_canvas.configure(
            scrollregion=self.supplier_canvas.bbox("all")))
        self.supplier_canvas.bind("<Configure>", lambda e: self.supplier_canvas.itemconfig(
            self._sup_window, width=e.width))

        # Check all / none
        chk_bar = tk.Frame(left, bg=BG2)
        chk_bar.pack(fill="x", pady=(4, 0))
        all_btn = tk.Button(chk_bar, text="Select All", command=self._select_all)
        _style_btn(all_btn, bg=BG3, hover="#555555")
        all_btn.pack(side="left", padx=(0, 4))
        none_btn = tk.Button(chk_bar, text="Select None", command=self._select_none)
        _style_btn(none_btn, bg=BG3, hover="#555555")
        none_btn.pack(side="left")

        # Right panel: fetched files (takes all remaining space)
        right = tk.LabelFrame(mid, text=" Fetched Files ", bg=BG2, fg=FG, font=FONT_BOLD,
                               bd=1, relief="solid", padx=6, pady=6)
        right.pack(side="left", fill="both", expand=True)

        # Scrollable file list
        self.file_canvas = tk.Canvas(right, bg=BG2, highlightthickness=0)
        file_scroll = tk.Scrollbar(right, orient="vertical", command=self.file_canvas.yview)
        self.file_canvas.configure(yscrollcommand=file_scroll.set)
        file_scroll.pack(side="right", fill="y")
        self.file_canvas.pack(side="left", fill="both", expand=True)
        self.file_inner = tk.Frame(self.file_canvas, bg=BG2)
        self._file_window = self.file_canvas.create_window((0, 0), window=self.file_inner, anchor="nw")
        self.file_inner.bind("<Configure>", lambda e: self.file_canvas.configure(
            scrollregion=self.file_canvas.bbox("all")))
        self.file_canvas.bind("<Configure>", lambda e: self.file_canvas.itemconfig(
            self._file_window, width=e.width))

        # ── Log area ────────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(self.root, text=" Log ", bg=BG, fg=FG,
                                   font=FONT_BOLD, bd=1, relief="solid",
                                   padx=6, pady=4)
        log_frame.pack(fill="x", padx=14, pady=(4, 0))

        self.log_text = tk.Text(log_frame, bg=BG2, fg=FG, font=FONT_SMALL,
                                height=6, state="disabled", relief="flat",
                                wrap="word", insertbackground=FG)
        log_scroll = tk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="x", expand=True)

        # ── Bottom button bar ────────────────────────────────────────────────
        bot = tk.Frame(self.root, bg=BG, pady=10, padx=14)
        bot.pack(fill="x")

        self.fetch_btn = tk.Button(bot, text="Fetch Data", command=self._fetch_data)
        _style_btn(self.fetch_btn)
        self.fetch_btn.pack(side="left", padx=(0, 10))

        self.process_btn = tk.Button(bot, text="Process", command=self._process)
        _style_btn(self.process_btn, bg="#2d7a2d", hover="#3a9a3a")
        self.process_btn.pack(side="left")

        self.powerbi_btn = tk.Button(bot, text="Export for PowerBI", command=self._ask_powerbi_export)
        _style_btn(self.powerbi_btn, bg="#6b3a9e", hover="#7e4db8")
        self.powerbi_btn.pack(side="left", padx=(10, 0))

        manage_col_btn = tk.Button(bot, text="⚙ Columns", command=self._manage_columns)
        _style_btn(manage_col_btn, bg=BG3, hover="#555555")
        manage_col_btn.pack(side="left", padx=(10, 0))

        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(bot, textvariable=self.status_var, bg=BG, fg=FG,
                 font=FONT_SMALL).pack(side="left", padx=16)

    # ─────────────────────────────────────────────────────────────────────────
    # Manage Columns dialog
    # ─────────────────────────────────────────────────────────────────────────

    def _manage_columns(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Manage Output Columns")
        dialog.configure(bg=BG)
        dialog.grab_set()
        dialog.resizable(True, True)
        dialog.geometry("400x520")

        tk.Label(dialog, text="Columns to keep in output (one per line):",
                 bg=BG, fg=FG, font=FONT_BOLD, pady=8).pack(fill="x", padx=14)

        text_frame = tk.Frame(dialog, bg=BG, padx=14)
        text_frame.pack(fill="both", expand=True)

        txt = tk.Text(text_frame, bg=ENTRY_BG, fg=FG, insertbackground=FG,
                      font=FONT_SMALL, relief="flat", bd=3, wrap="none")
        sb = tk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        # Populate with current list
        txt.insert("1.0", "\n".join(self.keep_columns))

        btn_frame = tk.Frame(dialog, bg=BG, pady=10)
        btn_frame.pack(fill="x", padx=14)

        def reset_default():
            txt.delete("1.0", "end")
            txt.insert("1.0", "\n".join(DEFAULT_KEEP_COLUMNS))

        def save():
            raw = txt.get("1.0", "end")
            cols = [c.strip() for c in raw.splitlines() if c.strip()]
            self.keep_columns = cols
            save_config({"keep_columns": cols})
            dialog.destroy()

        save_btn = tk.Button(btn_frame, text="Save", command=save)
        _style_btn(save_btn)
        save_btn.pack(side="left", padx=(0, 8))

        reset_btn = tk.Button(btn_frame, text="Reset to Default", command=reset_default)
        _style_btn(reset_btn, bg=BG3, hover="#555555")
        reset_btn.pack(side="left", padx=(0, 8))

        cancel_btn = tk.Button(btn_frame, text="Cancel", command=dialog.destroy)
        _style_btn(cancel_btn, bg=BG3, hover="#555555")
        cancel_btn.pack(side="left")

    # ─────────────────────────────────────────────────────────────────────────
    # Logging
    # ─────────────────────────────────────────────────────────────────────────

    def _log(self, msg: str):
        """Queue a log line from any thread."""
        import datetime
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log_queue.put(f"[{ts}] {msg}\n")

    def _poll_log(self):
        """Drain the log queue on the main thread every 50 ms."""
        try:
            while True:
                line = self._log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", line)
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(50, self._poll_log)

    # ─────────────────────────────────────────────────────────────────────────
    # State Restore
    # ─────────────────────────────────────────────────────────────────────────

    def _restore_state(self):
        base = self.cfg.get("base_path", "")
        if base and os.path.isdir(base):
            self._populate_suppliers(base)

    def _browse_path(self):
        path = filedialog.askdirectory(title="Select Spending & Rebate Reports Folder")
        if path:
            self.path_var.set(path)
            self._populate_suppliers(path)
            save_config({"base_path": path})

    # ─────────────────────────────────────────────────────────────────────────
    # Supplier Checkboxes
    # ─────────────────────────────────────────────────────────────────────────

    def _populate_suppliers(self, base_path: str):
        # Clear old widgets
        for widget in self.supplier_inner.winfo_children():
            widget.destroy()
        self.supplier_vars.clear()

        suppliers = get_suppliers(base_path)
        checked = self.cfg.get("checked_suppliers", [])

        for supplier in suppliers:
            var = tk.BooleanVar(value=(supplier in checked))
            self.supplier_vars[supplier] = var
            cb = tk.Checkbutton(
                self.supplier_inner, text=supplier, variable=var,
                bg=BG2, fg=FG, selectcolor=BG3, activebackground=BG2,
                activeforeground=FG, font=FONT, anchor="w",
                command=self._save_supplier_selection,
            )
            cb.pack(fill="x", pady=1)

    def _select_all(self):
        for var in self.supplier_vars.values():
            var.set(True)
        self._save_supplier_selection()

    def _select_none(self):
        for var in self.supplier_vars.values():
            var.set(False)
        self._save_supplier_selection()

    def _save_supplier_selection(self):
        selected = [s for s, v in self.supplier_vars.items() if v.get()]
        save_config({"checked_suppliers": selected})

    def _get_selected_suppliers(self) -> list[str]:
        return [s for s, v in self.supplier_vars.items() if v.get()]

    # ─────────────────────────────────────────────────────────────────────────
    # Fetch Data
    # ─────────────────────────────────────────────────────────────────────────

    def _fetch_data(self):
        base = self.path_var.get().strip()
        if not base or not os.path.isdir(base):
            messagebox.showerror("Error", "Please select a valid base folder first.")
            return
        selected = self._get_selected_suppliers()
        if not selected:
            messagebox.showwarning("Warning", "No suppliers selected.")
            return

        self.status_var.set("Fetching files…")
        self.root.update_idletasks()

        file_lists = fetch_supplier_files(base, selected)
        self.supplier_file_lists = file_lists
        self.supplier_file_idx = {s: 0 for s in selected}
        self.supplier_file_labels = {}

        self._log(f"Fetch Data — {len(selected)} supplier(s) selected.")
        for supplier in selected:
            files = file_lists.get(supplier, [])
            if files:
                self._log(f"  {supplier}: found {len(files)} file(s), latest → {files[0].name}")
            else:
                self._log(f"  {supplier}: no files found")

        # Clear file panel
        for widget in self.file_inner.winfo_children():
            widget.destroy()

        for supplier in selected:
            files = file_lists.get(supplier, [])
            row_frame = tk.Frame(self.file_inner, bg=BG2)
            row_frame.pack(fill="x", pady=2)

            tk.Label(row_frame, text=f"{supplier}:", bg=BG2, fg=FG,
                     font=FONT_BOLD, width=14, anchor="w").pack(side="left")

            left_btn = tk.Button(row_frame, text="◀", width=2,
                                 command=lambda s=supplier: self._shift_file(s, -1))
            _style_btn(left_btn, bg=BG3, hover="#555555")
            left_btn.pack(side="left", padx=(0, 2))

            svar = tk.StringVar()
            self.supplier_file_labels[supplier] = svar
            if files:
                svar.set(files[0].name)
            else:
                svar.set("(no file found)")

            lbl = tk.Label(row_frame, textvariable=svar, bg=BG2, fg=GREEN,
                           font=FONT_SMALL, anchor="w", wraplength=500)
            lbl.pack(side="left", fill="x", expand=True, padx=4)

            right_btn = tk.Button(row_frame, text="▶", width=2,
                                  command=lambda s=supplier: self._shift_file(s, +1))
            _style_btn(right_btn, bg=BG3, hover="#555555")
            right_btn.pack(side="left", padx=(2, 0))

        self.status_var.set(f"Fetched files for {len(selected)} supplier(s).")

    def _shift_file(self, supplier: str, direction: int):
        files = self.supplier_file_lists.get(supplier, [])
        if not files:
            return
        idx = self.supplier_file_idx.get(supplier, 0)
        new_idx = max(0, min(len(files) - 1, idx + direction))
        self.supplier_file_idx[supplier] = new_idx
        self.supplier_file_labels[supplier].set(files[new_idx].name)

    # ─────────────────────────────────────────────────────────────────────────
    # Process
    # ─────────────────────────────────────────────────────────────────────────

    def _process(self):
        if not self.supplier_file_lists:
            messagebox.showwarning("Warning", "Please Fetch Data first.")
            return

        selected = self._get_selected_suppliers()
        if not selected:
            messagebox.showwarning("Warning", "No suppliers selected.")
            return

        # Build {supplier: chosen_file_path}
        supplier_files = {}
        for s in selected:
            files = self.supplier_file_lists.get(s, [])
            idx = self.supplier_file_idx.get(s, 0)
            if files:
                supplier_files[s] = files[idx]
            else:
                messagebox.showerror("Error", f"No file available for supplier: {s}")
                return

        # Find common FY sheets
        self.status_var.set("Detecting common FY sheets…")
        self.root.update_idletasks()

        common_sheets = self._find_common_fy_sheets(supplier_files)
        if not common_sheets:
            messagebox.showerror("Error", "No common FY sheets found across all selected suppliers.")
            self.status_var.set("Ready.")
            return

        # Show FY selection dialog
        chosen_sheet = self._ask_fy_sheet(common_sheets)
        if not chosen_sheet:
            self.status_var.set("Ready.")
            return

        self.process_btn.configure(state="disabled")
        self.fetch_btn.configure(state="disabled")
        self.status_var.set(f"Processing {chosen_sheet}…")
        self.root.update_idletasks()

        # Run in background thread to keep UI responsive
        thread = threading.Thread(
            target=self._run_processing,
            args=(supplier_files, chosen_sheet),
            daemon=True,
        )
        thread.start()

    def _find_common_fy_sheets(self, supplier_files: dict) -> list[str]:
        """Return FY sheets present in ALL supplier files."""
        sets_list = []
        for supplier, file_path in supplier_files.items():
            try:
                sheets = set(get_fy_sheets(str(file_path)))
            except Exception as e:
                messagebox.showerror("Error", f"Cannot read {supplier} file:\n{e}")
                return []
            sets_list.append(sheets)
        if not sets_list:
            return []
        common = sets_list[0]
        for s in sets_list[1:]:
            common = common.intersection(s)
        # Sort by FY number
        return sorted(common, key=lambda x: int("".join(filter(str.isdigit, x))))

    def _ask_fy_sheet(self, sheets: list[str]) -> str | None:
        """Modal dialog for FY sheet selection. Returns chosen sheet name or None."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Select FY Sheet")
        dialog.configure(bg=BG)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(dialog, text="Select the FY sheet to process:",
                 bg=BG, fg=FG, font=FONT_BOLD, padx=16, pady=10).pack()

        chosen = tk.StringVar(value=sheets[0])
        for sheet in sheets:
            rb = tk.Radiobutton(dialog, text=sheet, variable=chosen, value=sheet,
                                bg=BG, fg=FG, selectcolor=BG3, activebackground=BG,
                                activeforeground=FG, font=FONT)
            rb.pack(anchor="w", padx=20)

        result = [None]

        def confirm():
            result[0] = chosen.get()
            dialog.destroy()

        def cancel():
            dialog.destroy()

        btn_frame = tk.Frame(dialog, bg=BG, pady=10)
        btn_frame.pack()
        ok_btn = tk.Button(btn_frame, text="OK", command=confirm)
        _style_btn(ok_btn)
        ok_btn.pack(side="left", padx=6)
        cancel_btn = tk.Button(btn_frame, text="Cancel", command=cancel)
        _style_btn(cancel_btn, bg=BG3, hover="#555555")
        cancel_btn.pack(side="left", padx=6)

        dialog.wait_window()
        return result[0]

    def _run_processing(self, supplier_files: dict, sheet_name: str):
        try:
            self._log(f"Starting processing for {sheet_name}…")

            # Clear source_data folder (delete files individually to avoid OneDrive WinError 5)
            os.makedirs(self.source_data_dir, exist_ok=True)
            for _f in os.listdir(self.source_data_dir):
                _fp = os.path.join(self.source_data_dir, _f)
                if os.path.isfile(_fp):
                    try:
                        os.remove(_fp)
                    except OSError:
                        pass
            self._log("Cleared source_data folder.")

            frames = []
            for supplier, file_path in supplier_files.items():
                self._log(f"  Processing {supplier}: {Path(str(file_path)).name}")
                from src.processing.transformer import copy_source_file, process_supplier_sheet
                copy_source_file(str(file_path), self.source_data_dir)
                df = process_supplier_sheet(str(file_path), sheet_name, supplier)
                if not df.empty:
                    frames.append(df)
                    self._log(f"    → {len(df)} rows")
                else:
                    self._log(f"    → (no data)")

            if not frames:
                self.root.after(0, lambda: messagebox.showerror("Error", "No data produced after processing."))
                return

            import pandas as pd
            merged = pd.concat(frames, ignore_index=True)
            self._log(f"Merged total: {len(merged)} rows across {len(frames)} supplier(s).")

            # Drop columns that only have data for a single supplier
            if len(frames) > 1:
                drop_single = []
                for col in merged.columns:
                    if col == "GTK Suppliers":
                        continue
                    suppliers_with_data = merged.loc[merged[col].notna(), "GTK Suppliers"].unique()
                    if len(suppliers_with_data) == 1:
                        drop_single.append(col)
                if drop_single:
                    merged.drop(columns=drop_single, inplace=True)
                    self._log(f"Dropped {len(drop_single)} supplier-specific column(s): {drop_single}")

            # Keep only configured columns (preserve order, skip missing)
            keep = [c for c in self.keep_columns if c in merged.columns]
            missing = [c for c in self.keep_columns if c not in merged.columns]
            merged = merged[keep]
            if missing:
                self._log(f"Note: {len(missing)} keep-column(s) not found in data: {missing}")
            self._log(f"Output columns ({len(keep)}): {keep}")

            out_path = save_history(merged, sheet_name, self.history_dir)
            self._log(f"Saved history: {Path(out_path).name}")
            self.root.after(0, lambda: self._on_processing_done(out_path, sheet_name))

        except Exception as e:
            import traceback
            err = traceback.format_exc()
            self._log(f"ERROR: {e}")
            self.root.after(0, lambda: messagebox.showerror("Processing Error", str(e)))
        finally:
            self.root.after(0, self._re_enable_buttons)

    def _re_enable_buttons(self):
        self.process_btn.configure(state="normal")
        self.fetch_btn.configure(state="normal")

    def _on_processing_done(self, out_path: str, sheet_name: str):
        self.status_var.set(f"Done! Saved: {Path(out_path).name}")
        messagebox.showinfo("Processing Complete",
                            f"Consolidated data saved to:\n{out_path}")
        # Ask for PowerBI merge
        self._ask_powerbi_export()

    # ─────────────────────────────────────────────────────────────────────────
    # PowerBI Export
    # ─────────────────────────────────────────────────────────────────────────

    def _ask_powerbi_export(self):
        available = get_available_fy(self.history_dir)
        if not available:
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Export for PowerBI")
        dialog.configure(bg=BG)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(dialog, text="Select FY files to include in PowerBI export:",
                 bg=BG, fg=FG, font=FONT_BOLD, padx=16, pady=10).pack()

        fy_vars: dict[str, tk.BooleanVar] = {}
        for fy in available:
            var = tk.BooleanVar(value=True)
            fy_vars[fy] = var
            cb = tk.Checkbutton(dialog, text=fy, variable=var,
                                bg=BG, fg=FG, selectcolor=BG3, activebackground=BG,
                                activeforeground=FG, font=FONT)
            cb.pack(anchor="w", padx=20)

        result = [False]

        def confirm():
            result[0] = True
            dialog.destroy()

        def cancel():
            dialog.destroy()

        btn_frame = tk.Frame(dialog, bg=BG, pady=10)
        btn_frame.pack()
        ok_btn = tk.Button(btn_frame, text="Export", command=confirm)
        _style_btn(ok_btn, bg="#2d7a2d", hover="#3a9a3a")
        ok_btn.pack(side="left", padx=6)
        cancel_btn = tk.Button(btn_frame, text="Skip", command=cancel)
        _style_btn(cancel_btn, bg=BG3, hover="#555555")
        cancel_btn.pack(side="left", padx=6)

        dialog.wait_window()

        if result[0]:
            selected_fy = [fy for fy, v in fy_vars.items() if v.get()]
            if not selected_fy:
                messagebox.showwarning("Warning", "No FY selected for export.")
                return
            out_path = merge_for_powerbi(selected_fy, self.history_dir, self.output_dir)
            if out_path:
                messagebox.showinfo("Export Complete",
                                    f"PowerBI file saved to:\n{out_path}")
                self.status_var.set(f"PowerBI export: {Path(out_path).name}")
            else:
                messagebox.showerror("Error", "PowerBI export failed.")

    # ─────────────────────────────────────────────────────────────────────────
    # Window close
    # ─────────────────────────────────────────────────────────────────────────

    def _on_close(self):
        self.root.destroy()
        sys.exit(0)


def main():
    root = tk.Tk()
    app = ShipmentApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
