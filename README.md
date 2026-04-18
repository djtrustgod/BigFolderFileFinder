# Folder Size Scanner

A desktop GUI utility that scans a folder for large files and subfolders, then exports the results to a formatted Excel report.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue) ![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey) ![License](https://img.shields.io/badge/License-MIT-green)

---

## Features

- **Browse & scan** any folder with a configurable size threshold (default 100 MB)
- **Live scan log** shows progress in real time as directories are traversed
- **Sortable results** — click any column header to sort files or folders by name, size, date, etc.
- **Stop mid-scan** — cancel a long-running scan without losing results collected so far
- **Excel export** — generates a polished `.xlsx` report with three sheets:
  - `📊 Summary` — scan metadata, counts, total sizes, largest item highlights
  - `📄 Large Files` — filterable table with color-scale heatmap on the Size (MB) column
  - `📁 Large Folders` — same layout, green color theme, recursive folder sizes
- **Auto-open** — optionally opens the report in Excel immediately after export
- **Auto-installs `openpyxl`** if not already present

---

## Requirements

- Python 3.8 or later
- `openpyxl` (installed automatically on first run if missing)
- Tkinter (bundled with standard Python on Windows and macOS; on Linux install `python3-tk`)

---

## Installation

```bash
git clone https://github.com/your-username/BigFolderFileFinder.git
cd BigFolderFileFinder
# openpyxl is installed automatically on first run, or install it manually:
pip install openpyxl
```

---

## Usage

```bash
python folder_size_scanner.py
```

1. Click **Browse…** and choose the root folder to scan.
2. Set the **Min size** threshold (in MB) — any file or folder at or above this size is reported.
3. Click **▶ Scan**.  Results appear in the **Large Files** and **Large Folders** tabs as the scan completes.
4. Click **📥 Export to Excel** to save the report.

---

## Excel Report Preview

| Sheet | Contents |
|---|---|
| 📊 Summary | Scan path, threshold, item counts, total size, largest file & folder |
| 📄 Large Files | #, File Name, Extension, Size, Size (MB), Modified Date, Path |
| 📁 Large Folders | #, Folder Name, Size, Size (MB), Modified Date, Path |

All data sheets include auto-filters, frozen header rows, alternating row shading, and a green-yellow-red color scale on the Size (MB) column.

---

## Project Structure

```
BigFolderFileFinder/
└── folder_size_scanner.py   # Single-file application (scanning, Excel export, GUI)
```

---

## License

MIT — free to use, modify, and distribute.
