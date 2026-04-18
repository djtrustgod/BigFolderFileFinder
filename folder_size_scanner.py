"""
Folder Size Scanner
A GUI utility to find large files and folders and export them to Excel.
"""

import os
import platform
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, GradientFill
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule


# ─────────────────────────────────────────────
#  Scanning logic
# ─────────────────────────────────────────────

def get_folder_size(folder_path):
    """Recursively compute size of a folder in bytes."""
    total = 0
    try:
        for entry in os.scandir(folder_path):
            try:
                if entry.is_file(follow_symlinks=False):
                    total += entry.stat(follow_symlinks=False).st_size
                elif entry.is_dir(follow_symlinks=False):
                    total += get_folder_size(entry.path)
            except (PermissionError, OSError):
                pass
    except (PermissionError, OSError):
        pass
    return total


def bytes_to_mb(b):
    return b / (1024 * 1024)


def bytes_to_human(b):
    if b >= 1024 ** 3:
        return f"{b / 1024**3:.2f} GB"
    if b >= 1024 ** 2:
        return f"{b / 1024**2:.2f} MB"
    if b >= 1024:
        return f"{b / 1024:.2f} KB"
    return f"{b} B"


def get_modified_date(path):
    try:
        ts = os.path.getmtime(path)
        return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
    except OSError:
        return "N/A"


def scan_directory(root_path, threshold_mb, progress_callback, stop_event):
    """
    Scan root_path and return (large_files, large_folders).
    progress_callback(message) is called during scanning.
    """
    threshold_bytes = threshold_mb * 1024 * 1024
    large_files = []
    large_folders = []

    progress_callback(f"Scanning: {root_path}")

    for dirpath, dirnames, filenames in os.walk(root_path, followlinks=False):
        if stop_event.is_set():
            break

        progress_callback(f"Checking: {dirpath}")

        # Check files
        for fname in filenames:
            if stop_event.is_set():
                break
            fpath = os.path.join(dirpath, fname)
            try:
                size = os.path.getsize(fpath)
                if size >= threshold_bytes:
                    rel = os.path.relpath(fpath, root_path)
                    ext = Path(fpath).suffix.lower() or "(no ext)"
                    large_files.append({
                        "name": fname,
                        "relative_path": rel,
                        "full_path": fpath,
                        "size_bytes": size,
                        "size_mb": bytes_to_mb(size),
                        "size_human": bytes_to_human(size),
                        "extension": ext,
                        "modified": get_modified_date(fpath),
                        "parent_folder": dirpath,
                    })
            except (PermissionError, OSError):
                pass

        # Check sub-folders (not the root itself)
        for dname in dirnames:
            if stop_event.is_set():
                break
            dpath = os.path.join(dirpath, dname)
            try:
                size = get_folder_size(dpath)
                if size >= threshold_bytes:
                    rel = os.path.relpath(dpath, root_path)
                    large_folders.append({
                        "name": dname,
                        "relative_path": rel,
                        "full_path": dpath,
                        "size_bytes": size,
                        "size_mb": bytes_to_mb(size),
                        "size_human": bytes_to_human(size),
                        "modified": get_modified_date(dpath),
                        "parent_folder": dirpath,
                    })
            except (PermissionError, OSError):
                pass

    large_files.sort(key=lambda x: x["size_bytes"], reverse=True)
    large_folders.sort(key=lambda x: x["size_bytes"], reverse=True)
    return large_files, large_folders


# ─────────────────────────────────────────────
#  Excel export
# ─────────────────────────────────────────────

HEADER_FILL   = PatternFill("solid", fgColor="1F3864")   # Dark navy
SUBHEAD_FILL  = PatternFill("solid", fgColor="2E75B6")   # Medium blue
ALT_FILL      = PatternFill("solid", fgColor="EBF3FB")   # Light blue tint
FILE_HDR_FILL = PatternFill("solid", fgColor="1F4E79")   # Deep blue (files)
FOLD_HDR_FILL = PatternFill("solid", fgColor="375623")   # Dark green (folders)
ALT_GREEN     = PatternFill("solid", fgColor="EBF5E1")   # Light green tint
WHITE_FILL    = PatternFill("solid", fgColor="FFFFFF")

THIN = Side(style="thin", color="BDD7EE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def hdr_font(size=11, color="FFFFFF"):
    return Font(name="Arial", bold=True, size=size, color=color)


def cell_font(size=10, bold=False, color="000000"):
    return Font(name="Arial", size=size, bold=bold, color=color)


def apply_header_row(ws, row_num, values, fill, font_color="FFFFFF", heights=18):
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row_num, column=col, value=val)
        c.font = hdr_font(10, font_color)
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row_num].height = heights


def set_col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w


def build_summary_sheet(ws, root_path, threshold_mb, large_files, large_folders, scan_time):
    ws.title = "📊 Summary"

    # Title banner
    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value = "Folder Size Scanner — Report"
    t.font = Font(name="Arial", bold=True, size=16, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="1F3864")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:D2")
    sub = ws["A2"]
    sub.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}   |   Scan time: {scan_time:.1f}s"
    sub.font = Font(name="Arial", size=9, color="FFFFFF", italic=True)
    sub.fill = PatternFill("solid", fgColor="2E75B6")
    sub.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    # Info rows
    info = [
        ("Scanned Folder", root_path),
        ("Size Threshold", f"{threshold_mb:,.1f} MB"),
        ("Large Files Found", f"{len(large_files):,}"),
        ("Large Folders Found", f"{len(large_folders):,}"),
        ("Total Items Found", f"{len(large_files) + len(large_folders):,}"),
    ]
    if large_files:
        total_file_bytes = sum(f["size_bytes"] for f in large_files)
        info.append(("Total Size (Files)", bytes_to_human(total_file_bytes)))
    if large_folders:
        biggest = large_folders[0]
        info.append(("Largest Folder", f"{biggest['name']}  ({biggest['size_human']})"))
    if large_files:
        biggest_f = large_files[0]
        info.append(("Largest File", f"{biggest_f['name']}  ({biggest_f['size_human']})"))

    label_fill = PatternFill("solid", fgColor="D6E4F0")
    value_fill = PatternFill("solid", fgColor="FFFFFF")

    for i, (label, value) in enumerate(info, start=4):
        lc = ws.cell(row=i, column=1, value=label)
        lc.font = Font(name="Arial", bold=True, size=10, color="1F3864")
        lc.fill = label_fill
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lc.border = BORDER

        vc = ws.cell(row=i, column=2, value=value)
        vc.font = Font(name="Arial", size=10, color="000000")
        vc.fill = value_fill
        vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        vc.border = BORDER

        ws.merge_cells(f"B{i}:D{i}")
        ws.row_dimensions[i].height = 16

    set_col_widths(ws, [28, 55, 15, 15])


def build_files_sheet(ws, large_files, threshold_mb):
    ws.title = "📄 Large Files"

    # Title
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = f"Large Files  (≥ {threshold_mb:,.1f} MB)"
    t.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    t.fill = FILE_HDR_FILL
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    headers = ["#", "File Name", "Extension", "Size", "Size (MB)", "Modified Date", "Relative Path", "Parent Folder"]
    apply_header_row(ws, 2, headers, FILE_HDR_FILL)

    for i, f in enumerate(large_files, 1):
        row = i + 2
        fill = ALT_FILL if i % 2 == 0 else WHITE_FILL
        vals = [
            i,
            f["name"],
            f["extension"],
            f["size_human"],
            round(f["size_mb"], 3),
            f["modified"],
            f["relative_path"],
            f["parent_folder"],
        ]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = cell_font()
            c.fill = fill
            c.border = BORDER
            c.alignment = Alignment(vertical="center", indent=1)
            if col == 1:
                c.alignment = Alignment(horizontal="center", vertical="center")
            if col == 5:
                c.number_format = "#,##0.000"
                c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        ws.row_dimensions[row].height = 15

    # Freeze header
    ws.freeze_panes = "A3"

    # Auto-filter
    if large_files:
        ws.auto_filter.ref = f"A2:H{len(large_files)+2}"

    # Color scale on MB column (E)
    if large_files:
        last = len(large_files) + 2
        ws.conditional_formatting.add(
            f"E3:E{last}",
            ColorScaleRule(
                start_type="min", start_color="63BE7B",
                mid_type="percentile", mid_value=50, mid_color="FFEB84",
                end_type="max", end_color="F8696B",
            ),
        )

    set_col_widths(ws, [5, 34, 10, 12, 12, 18, 55, 50])


def build_folders_sheet(ws, large_folders, threshold_mb):
    ws.title = "📁 Large Folders"

    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = f"Large Folders  (≥ {threshold_mb:,.1f} MB)"
    t.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    t.fill = FOLD_HDR_FILL
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    headers = ["#", "Folder Name", "Size", "Size (MB)", "Modified Date", "Relative Path", "Parent Folder"]
    apply_header_row(ws, 2, headers, FOLD_HDR_FILL)

    for i, f in enumerate(large_folders, 1):
        row = i + 2
        fill = ALT_GREEN if i % 2 == 0 else WHITE_FILL
        vals = [
            i,
            f["name"],
            f["size_human"],
            round(f["size_mb"], 3),
            f["modified"],
            f["relative_path"],
            f["parent_folder"],
        ]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = cell_font()
            c.fill = fill
            c.border = BORDER
            c.alignment = Alignment(vertical="center", indent=1)
            if col == 1:
                c.alignment = Alignment(horizontal="center", vertical="center")
            if col == 4:
                c.number_format = "#,##0.000"
                c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        ws.row_dimensions[row].height = 15

    ws.freeze_panes = "A3"

    if large_folders:
        ws.auto_filter.ref = f"A2:G{len(large_folders)+2}"
        last = len(large_folders) + 2
        ws.conditional_formatting.add(
            f"D3:D{last}",
            ColorScaleRule(
                start_type="min", start_color="63BE7B",
                mid_type="percentile", mid_value=50, mid_color="FFEB84",
                end_type="max", end_color="F8696B",
            ),
        )

    set_col_widths(ws, [5, 34, 12, 12, 18, 55, 50])


def export_to_excel(root_path, threshold_mb, large_files, large_folders, scan_time, out_path):
    wb = Workbook()
    ws_summary = wb.active
    build_summary_sheet(ws_summary, root_path, threshold_mb, large_files, large_folders, scan_time)

    ws_files = wb.create_sheet()
    build_files_sheet(ws_files, large_files, threshold_mb)

    ws_folders = wb.create_sheet()
    build_folders_sheet(ws_folders, large_folders, threshold_mb)

    # Set tab colors
    wb["📊 Summary"].sheet_properties.tabColor = "1F3864"
    wb["📄 Large Files"].sheet_properties.tabColor = "1F4E79"
    wb["📁 Large Folders"].sheet_properties.tabColor = "375623"

    wb.active = wb["📊 Summary"]
    wb.save(out_path)


# ─────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────

class ScannerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Folder Size Scanner")
        self.geometry("960x680")
        self.minsize(800, 550)
        self.configure(bg="#F0F4F8")

        self._large_files = []
        self._large_folders = []
        self._scan_time = 0.0
        self._stop_event = threading.Event()
        self._scanning = False

        self._build_ui()

    # ── UI construction ──────────────────────

    def _build_ui(self):
        # ── Top control panel ──
        ctrl = tk.Frame(self, bg="#1F3864", pady=12, padx=16)
        ctrl.pack(fill="x")

        tk.Label(ctrl, text="🔍 Folder Size Scanner",
                 font=("Arial", 16, "bold"), bg="#1F3864", fg="white").grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 8))

        # Folder picker
        tk.Label(ctrl, text="Folder:", font=("Arial", 10, "bold"),
                 bg="#1F3864", fg="#BDD7EE").grid(row=1, column=0, sticky="w", padx=(0, 6))
        self._folder_var = tk.StringVar()
        self._folder_entry = tk.Entry(ctrl, textvariable=self._folder_var,
                                      font=("Arial", 10), width=52,
                                      relief="flat", bd=1)
        self._folder_entry.grid(row=1, column=1, sticky="ew", padx=(0, 6))
        tk.Button(ctrl, text="Browse…", command=self._browse_folder,
                  font=("Arial", 10), bg="#2E75B6", fg="white",
                  relief="flat", padx=10, cursor="hand2").grid(row=1, column=2, padx=(0, 20))

        # Threshold
        tk.Label(ctrl, text="Min size:", font=("Arial", 10, "bold"),
                 bg="#1F3864", fg="#BDD7EE").grid(row=1, column=3, sticky="w", padx=(0, 6))
        self._threshold_var = tk.StringVar(value="100")
        tk.Entry(ctrl, textvariable=self._threshold_var,
                 font=("Arial", 10), width=8,
                 relief="flat", bd=1).grid(row=1, column=4, padx=(0, 4))
        tk.Label(ctrl, text="MB", font=("Arial", 10, "bold"),
                 bg="#1F3864", fg="#BDD7EE").grid(row=1, column=5, padx=(0, 16))

        # Scan / Stop / Export buttons
        btn_frame = tk.Frame(ctrl, bg="#1F3864")
        btn_frame.grid(row=1, column=6, padx=(8, 0))

        self._scan_btn = tk.Button(btn_frame, text="▶  Scan", command=self._start_scan,
                                   font=("Arial", 10, "bold"), bg="#70AD47", fg="white",
                                   relief="flat", padx=14, pady=4, cursor="hand2")
        self._scan_btn.pack(side="left", padx=(0, 6))

        self._stop_btn = tk.Button(btn_frame, text="⏹ Stop", command=self._stop_scan,
                                   font=("Arial", 10), bg="#C55A11", fg="white",
                                   relief="flat", padx=10, pady=4, cursor="hand2",
                                   state="disabled")
        self._stop_btn.pack(side="left", padx=(0, 6))

        self._export_btn = tk.Button(btn_frame, text="📥 Export to Excel",
                                     command=self._export_excel,
                                     font=("Arial", 10, "bold"), bg="#1F4E79", fg="white",
                                     relief="flat", padx=14, pady=4, cursor="hand2",
                                     state="disabled")
        self._export_btn.pack(side="left")

        ctrl.columnconfigure(1, weight=1)

        # ── Status bar ──
        self._status_var = tk.StringVar(value="Ready — choose a folder and click Scan.")
        status_bar = tk.Frame(self, bg="#DDE8F0", pady=4, padx=10)
        status_bar.pack(fill="x")
        tk.Label(status_bar, textvariable=self._status_var,
                 font=("Arial", 9), bg="#DDE8F0", fg="#1F3864",
                 anchor="w").pack(fill="x")

        # ── Notebook (tabs) ──
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook", background="#F0F4F8", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Arial", 10, "bold"),
                         padding=[12, 6], background="#BDD7EE", foreground="#1F3864")
        style.map("TNotebook.Tab",
                  background=[("selected", "#1F3864")],
                  foreground=[("selected", "white")])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        self._files_tab = self._build_tree_tab(
            nb, "📄  Large Files",
            ["#", "File Name", "Extension", "Size", "Size (MB)", "Modified", "Relative Path"],
            [40, 260, 80, 90, 90, 135, 400]
        )
        self._folders_tab = self._build_tree_tab(
            nb, "📁  Large Folders",
            ["#", "Folder Name", "Size", "Size (MB)", "Modified", "Relative Path"],
            [40, 300, 90, 90, 135, 400]
        )
        self._log_tab, self._log_text = self._build_log_tab(nb)

        nb.add(self._files_tab, text="📄  Large Files")
        nb.add(self._folders_tab, text="📁  Large Folders")
        nb.add(self._log_tab, text="📋  Scan Log")

        # progress bar (hidden initially)
        self._progress = ttk.Progressbar(self, mode="indeterminate")

    def _build_tree_tab(self, parent, title, columns, col_widths):
        frame = tk.Frame(parent, bg="#F0F4F8")
        tree_frame = tk.Frame(frame)
        tree_frame.pack(fill="both", expand=True, padx=4, pady=4)

        style = ttk.Style()
        style.configure("Custom.Treeview",
                         font=("Arial", 9),
                         rowheight=22,
                         background="white",
                         fieldbackground="white",
                         foreground="#1F3864")
        style.configure("Custom.Treeview.Heading",
                         font=("Arial", 9, "bold"),
                         background="#1F3864",
                         foreground="white",
                         relief="flat")
        style.map("Custom.Treeview",
                  background=[("selected", "#2E75B6")],
                  foreground=[("selected", "white")])

        tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                            style="Custom.Treeview")
        for col, w in zip(columns, col_widths):
            tree.heading(col, text=col,
                         command=lambda c=col, t=tree: self._sort_tree(t, c, False))
            anchor = "e" if col in ("Size (MB)", "#") else "w"
            tree.column(col, width=w, anchor=anchor, minwidth=30)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        tree.tag_configure("odd", background="#FFFFFF")
        tree.tag_configure("even", background="#EBF3FB")

        frame._tree = tree
        return frame

    def _build_log_tab(self, parent):
        frame = tk.Frame(parent, bg="#1C1C1C")
        txt = tk.Text(frame, font=("Courier New", 9), bg="#1C1C1C", fg="#A8D8A8",
                      wrap="none", relief="flat", state="disabled")
        sb = ttk.Scrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        return frame, txt

    # ── Sorting ──────────────────────────────

    def _sort_tree(self, tree, col, reverse):
        data = [(tree.set(k, col), k) for k in tree.get_children("")]
        try:
            data.sort(key=lambda t: float(t[0].replace(",", "")), reverse=reverse)
        except ValueError:
            data.sort(key=lambda t: t[0].lower(), reverse=reverse)
        for index, (_, k) in enumerate(data):
            tree.move(k, "", index)
            tag = "even" if index % 2 == 0 else "odd"
            tree.item(k, tags=(tag,))
        tree.heading(col, command=lambda: self._sort_tree(tree, col, not reverse))

    # ── Actions ──────────────────────────────

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Choose a folder to scan")
        if path:
            self._folder_var.set(path)

    def _start_scan(self):
        folder = self._folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Invalid Folder", "Please choose a valid folder first.")
            return
        try:
            threshold = float(self._threshold_var.get())
            if threshold <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Threshold", "Please enter a positive number for the size threshold.")
            return

        # Clear previous results
        for item in self._files_tab._tree.get_children():
            self._files_tab._tree.delete(item)
        for item in self._folders_tab._tree.get_children():
            self._folders_tab._tree.delete(item)
        self._log_clear()
        self._large_files = []
        self._large_folders = []
        self._export_btn.config(state="disabled")
        self._stop_event.clear()
        self._scanning = True

        self._scan_btn.config(state="disabled")
        self._stop_btn.config(state="normal")
        self._progress.pack(fill="x", padx=8, pady=(0, 4))
        self._progress.start(12)
        self._status("Scanning…")

        self._scan_start_time = datetime.now()
        t = threading.Thread(
            target=self._scan_worker,
            args=(folder, threshold),
            daemon=True,
        )
        t.start()

    def _scan_worker(self, folder, threshold):
        try:
            files, folders = scan_directory(
                folder, threshold,
                progress_callback=lambda msg: self.after(0, self._log_append, msg),
                stop_event=self._stop_event,
            )
            elapsed = (datetime.now() - self._scan_start_time).total_seconds()
            self.after(0, self._scan_done, files, folders, folder, threshold, elapsed)
        except Exception as e:
            self.after(0, self._scan_error, str(e))

    def _scan_done(self, files, folders, folder, threshold, elapsed):
        self._large_files = files
        self._large_folders = folders
        self._scan_time = elapsed
        self._scanning = False
        self._progress.stop()
        self._progress.pack_forget()
        self._scan_btn.config(state="normal")
        self._stop_btn.config(state="disabled")

        # Populate trees
        for i, f in enumerate(files):
            tag = "even" if i % 2 == 0 else "odd"
            self._files_tab._tree.insert("", "end", tags=(tag,), values=(
                i + 1, f["name"], f["extension"],
                f["size_human"], f"{f['size_mb']:.3f}",
                f["modified"], f["relative_path"],
            ))

        for i, f in enumerate(folders):
            tag = "even" if i % 2 == 0 else "odd"
            self._folders_tab._tree.insert("", "end", tags=(tag,), values=(
                i + 1, f["name"],
                f["size_human"], f"{f['size_mb']:.3f}",
                f["modified"], f["relative_path"],
            ))

        stopped = " (scan stopped early)" if self._stop_event.is_set() else ""
        self._status(
            f"✔  Done{stopped} — {len(files)} large file(s), {len(folders)} large folder(s) "
            f"found above {threshold:.1f} MB  |  Scan time: {elapsed:.1f}s"
        )
        if files or folders:
            self._export_btn.config(state="normal")
        self._log_append(f"\n✔ Scan complete. {len(files)} files, {len(folders)} folders.")

    def _scan_error(self, msg):
        self._scanning = False
        self._progress.stop()
        self._progress.pack_forget()
        self._scan_btn.config(state="normal")
        self._stop_btn.config(state="disabled")
        messagebox.showerror("Scan Error", f"An error occurred:\n{msg}")
        self._status("Error during scan.")

    def _stop_scan(self):
        self._stop_event.set()
        self._status("Stopping scan…")
        self._stop_btn.config(state="disabled")

    def _export_excel(self):
        folder = self._folder_var.get().strip()
        threshold = float(self._threshold_var.get())
        default_name = f"FolderScan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_path = filedialog.asksaveasfilename(
            title="Save Excel Report",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not out_path:
            return
        try:
            export_to_excel(folder, threshold, self._large_files, self._large_folders,
                            self._scan_time, out_path)
            self._status(f"✔  Excel report saved: {out_path}")
            if messagebox.askyesno("Export Complete",
                                   f"Report saved to:\n{out_path}\n\nOpen the file now?"):
                if platform.system() == "Windows":
                    os.startfile(out_path)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", out_path])
                else:
                    subprocess.Popen(["xdg-open", out_path])
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to save Excel:\n{e}")

    # ── Helpers ──────────────────────────────

    def _status(self, msg):
        self._status_var.set(msg)

    def _log_append(self, msg):
        self._log_text.config(state="normal")
        self._log_text.insert("end", msg + "\n")
        self._log_text.see("end")
        self._log_text.config(state="disabled")

    def _log_clear(self):
        self._log_text.config(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.config(state="disabled")


# ─────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────

if __name__ == "__main__":
    app = ScannerApp()
    app.mainloop()
