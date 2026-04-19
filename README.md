# Folder Size Scanner

A modern desktop GUI that scans a folder for large files and subfolders, reports their true **Size on Disk**, and exports the results to a polished Excel report.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue) ![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey) ![UI](https://img.shields.io/badge/UI-CustomTkinter-2E8B57) ![License](https://img.shields.io/badge/License-MIT-green)

---

## Features

- **Modern UI** powered by [CustomTkinter](https://customtkinter.tomschimansky.com/) — clean, flat design with rounded corners and modern typography
- **Light / Dark / System themes** — switch at runtime from the segmented button in the header; "System" follows your OS appearance setting
- **True Size on Disk** — on Windows, uses `GetCompressedFileSizeW` so OneDrive/cloud placeholders, NTFS-compressed, and sparse files report the actual bytes allocated on disk (matches what Windows Explorer shows under *Properties → Size on Disk*)
- **Browse & scan** any folder with a configurable size threshold (default 100 MB)
- **Live scan log** shows progress in real time as directories are traversed
- **Sortable results** — click any column header to sort by name, size, modified date, etc.
- **Reveal in Explorer** — right-click any row (or double-click it) to show the file/folder in Windows Explorer, Finder (macOS), or your Linux file manager. Files are opened with the item preselected; folders open directly to their contents.
- **Stop mid-scan** — cancel a long-running scan; partial results are preserved
- **Excel export** — generates a formatted `.xlsx` report with three sheets:
  - `📊 Summary` — scan metadata, counts, total sizes, largest item highlights
  - `📄 Large Files` — filterable table with a red/yellow/green color-scale on the Size on Disk (MB) column
  - `📁 Large Folders` — same layout, green color theme, recursive folder sizes
- **Auto-open** — optionally opens the Excel report immediately after saving
- **Auto-installs dependencies** — `openpyxl` and `customtkinter` are installed on first run if missing

---

## Why "Size on Disk" matters

Traditional scanners report the *logical* file size. On Windows, OneDrive and other cloud providers create placeholder files that advertise their full size (e.g., 2 GB) while only a tiny stub is actually stored locally. NTFS compression and sparse files have the same mismatch. This app uses the Windows API that backs Explorer's "Size on Disk" field, so a 2 GB online-only file correctly shows up as ~4 KB — preventing false positives when hunting for space hogs.

On macOS and Linux the app falls back to the logical size (`os.path.getsize`), which matches standard behavior on those platforms.

---

## Requirements

- Python 3.8 or later
- `openpyxl` (auto-installed on first run)
- `customtkinter` (auto-installed on first run)
- Tkinter — bundled with standard Python on Windows and macOS; on Linux install `python3-tk`

---

## Installation

```bash
git clone https://github.com/your-username/BigFolderFileFinder.git
cd BigFolderFileFinder
# Dependencies install automatically on first run, or install them manually:
pip install openpyxl customtkinter
```

---

## Usage

```bash
python folder_size_scanner.py
```

1. Click **Browse…** and choose the root folder to scan.
2. Set the **Min size** threshold (in MB). Any file or folder whose *Size on Disk* meets or exceeds this value is reported.
3. (Optional) Use the **Theme** switch in the header to toggle Light / Dark / System.
4. Click **▶ Scan**. Results stream into the **Large Files** and **Large Folders** tabs; the scan log updates live.
5. **Right-click** (or **double-click**) a row to show the file/folder in Windows Explorer. The context menu also offers "Open Containing Folder".
6. Click **📥 Export to Excel** to save a formatted report.

---

## Excel Report Preview

| Sheet | Contents |
|---|---|
| 📊 Summary | Scan path, threshold, item counts, total size, largest file & folder |
| 📄 Large Files | #, File Name, Extension, Size on Disk, Size on Disk (MB), Modified Date, Path |
| 📁 Large Folders | #, Folder Name, Size on Disk, Size on Disk (MB), Modified Date, Path |

All data sheets include auto-filters, frozen header rows, alternating row shading, and a green-yellow-red color scale on the Size on Disk (MB) column.

---

## Building a Standalone Executable

The app can be packaged into a single `.exe` (Windows) — no Python installation required on the target machine.

```bash
pip install pyinstaller openpyxl customtkinter pillow
python make_icon.py     # regenerates icon.ico (only needed if you tweak the design)
pyinstaller --noconfirm --onefile --windowed --name FolderSizeScanner --icon=icon.ico --collect-all customtkinter folder_size_scanner.py
```

The output is `dist/FolderSizeScanner.exe` (~25 MB) with a blue disk-cylinder icon. Flags explained:
- `--onefile` — bundle everything into a single executable
- `--windowed` — hide the console window (GUI app)
- `--icon=icon.ico` — Windows taskbar / Explorer icon
- `--collect-all customtkinter` — bundle CustomTkinter's theme JSON files and embedded fonts (PyInstaller can't auto-detect these data files)

The same command works on macOS and Linux to produce a native binary; on macOS swap `--icon=icon.ico` for `--icon=icon.icns`.

---

## Project Structure

```
BigFolderFileFinder/
├── folder_size_scanner.py   # Single-file application (scanning, Excel export, GUI)
├── make_icon.py             # Regenerates icon.ico (Pillow-based disk-cylinder design)
├── icon.ico                 # App icon (multi-resolution: 16/32/48/64/128/256)
├── README.md
└── .gitignore               # Excludes PyInstaller build/, dist/, *.spec
```

---

## License

MIT — free to use, modify, and distribute.
