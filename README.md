# 📋 Clipboard History Manager

> **Production-grade clipboard history tracker** — cross-platform, thread-safe, lifecycle-safe.  
> Built with Python + CustomTkinter. Stores, searches, pins and exports everything you copy.

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)
![Version](https://img.shields.io/badge/Version-1.0.0-orange)

---

## Table of Contents

- [Features](#features)
- [Screenshots](#screenshots)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Installation](#installation)
- [Running the App](#running-the-app)
- [Building a Standalone Executable](#building-a-standalone-executable)
- [Usage Guide](#usage-guide)
- [Keyboard Shortcuts](#keyboard-shortcuts)
- [Export Formats](#export-formats)
- [Configuration & Data Files](#configuration--data-files)
- [Architecture Notes](#architecture-notes)
- [Contributing](#contributing)
- [License](#license)

---

## Features

- **Automatic clipboard monitoring** — polls every 400 ms, captures every copy
- **Persistent history** — survives app restarts; stored as JSON in your home directory
- **Deduplication** — optional; moves duplicates to top instead of storing twice
- **Pin entries** — keep important items safe from bulk clears
- **Full-text search** — debounced live search across all history
- **Content-kind detection** — auto-classifies entries as URL, Email, Phone, Code, or Text
- **Three export formats** — Excel (`.xlsx`), Plain Text (`.txt`), JSON (`.json`)
- **Global hotkey** — `Ctrl+Shift+V` brings the window to focus from anywhere (requires `pynput`)
- **Configurable settings** — max items, font size, dedup toggle, whitespace trimming, start minimised
- **Thread-safe shutdown** — ordered teardown with queue drain and atomic file writes
- **1 MB per-entry cap** — silently rejects oversized clipboard data

---

## Screenshots

> _Add screenshots here once the app is running._  
> Suggested: main window, search in action, Excel export, settings dialog.

---

## Project Structure

```
clipboard-history-manager/
├── src/
│   └── clipboard-history-manager.py   # Full application — single file
├── README.md
├── requirements.txt
├── .gitignore
└── LICENSE
```

> All runtime data (history, settings, logs) is written to `~/.clipboard_manager/` — never inside the repo.

---

## Requirements

| Dependency | Version | Purpose |
|---|---|---|
| `customtkinter` | ≥ 5.2 | Modern dark-theme UI widgets |
| `pyperclip` | any | Cross-platform clipboard read/write |
| `openpyxl` | any | Excel `.xlsx` export *(optional but recommended)* |
| `pynput` | any | Global hotkey `Ctrl+Shift+V` *(optional)* |

**Python 3.9 or newer is required.**

---

## Installation

### 1 — Clone the repository

```bash
git clone https://github.com/your-username/clipboard-history-manager.git
cd clipboard-history-manager
```

> Replace `your-username` with your actual GitHub username.

### 2 — Create and activate a virtual environment *(recommended)*

```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
```

### 3 — Install dependencies

```bash
pip install -r requirements.txt
```

---

## Running the App

```bash
# From the repo root
python src/clipboard-history-manager.py
```

On **Linux** you may need an X11 clipboard backend:

```bash
sudo apt install xclip   # or xsel
```

---

## Building a Standalone Executable

Use [PyInstaller](https://pyinstaller.org) to produce a single distributable binary with **no Python installation required** on the target machine.

### Install PyInstaller

```bash
pip install pyinstaller
```

### Windows — single `.exe`

```bash
pyinstaller \
  --onefile \
  --windowed \
  --name "ClipboardHistoryManager" \
  --icon "assets/icon.ico" \
  src/clipboard-history-manager.py
```

> Remove `--icon` if you don't have an icon file yet.  
> The finished executable will be at `dist/ClipboardHistoryManager.exe`.

### macOS — `.app` bundle

```bash
pyinstaller \
  --onefile \
  --windowed \
  --name "ClipboardHistoryManager" \
  src/clipboard-history-manager.py
```

Output: `dist/ClipboardHistoryManager` (Unix binary) or `dist/ClipboardHistoryManager.app`.

### Linux — single binary

```bash
pyinstaller \
  --onefile \
  --name "ClipboardHistoryManager" \
  src/clipboard-history-manager.py
```

Output: `dist/ClipboardHistoryManager`.

### Notes on PyInstaller + CustomTkinter

CustomTkinter ships theme assets that PyInstaller doesn't bundle automatically. If the executable crashes with a missing theme error, add:

```bash
--add-data "$(python -c 'import customtkinter; import os; print(os.path.dirname(customtkinter.__file__))')/*:customtkinter/"
```

On Windows use `;` instead of `:` as the path separator:

```powershell
--add-data "C:\path\to\customtkinter\*;customtkinter/"
```

---

## Usage Guide

### First launch

On first run the app creates `~/.clipboard_manager/` and starts monitoring immediately. Start copying text anywhere — entries appear in real time.

### Search

Click the search bar (or press `Ctrl+F`) and type. Results filter live with a 120 ms debounce. Press `Escape` to clear.

### Pinning

Click **Pin** on any card. Pinned items sort to the top and survive **Clear History**. Click **Unpin** to release.

### Copying back

Click **Copy** on any card. The item is written back to the system clipboard and the button flashes green for 1.4 seconds.

### Deleting

Click **Delete** on a card to remove it permanently.

### Clearing history

Click **🗑 Clear History** in the toolbar. A confirmation dialog appears. Pinned items are always preserved.

### Settings

Click **⚙ Settings** in the header to adjust:

| Setting | Default | Range |
|---|---|---|
| Maximum history items | 500 | 10 – 2 000 |
| Deduplicate entries | On | — |
| Trim whitespace | On | — |
| Start minimised | Off | — |
| Interface font size | 13 | 10 – 20 |

---

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl+F` | Focus search bar |
| `Escape` | Clear search |
| `Ctrl+Shift+V` | Bring window to front *(requires pynput)* |

---

## Export Formats

All exports are **atomic** (written to a `.tmp` file then renamed) so a crash mid-write never corrupts the output.

### Excel (`.xlsx`) — Primary

Requires `openpyxl`. Produces a two-sheet workbook:

- **History** — every entry with auto-filter, frozen header, colour-coded kind column, per-row height scaling for multi-line content
- **Summary** — metadata, platform info, and a per-kind count table

### Plain Text (`.txt`) — Backup

UTF-8 with BOM. Each entry includes its ID, timestamp, kind, character count, tags, and raw content.

### JSON (`.json`) — Power User

Structured JSON with schema version, export metadata, statistics block, and the full items array.

---

## Configuration & Data Files

All data lives under `~/.clipboard_manager/`:

| File | Purpose |
|---|---|
| `history.json` | Clipboard history (auto-saved, atomic writes) |
| `settings.json` | User preferences |
| `app.log` | Rotating log (max 5 MB × 3 backups) |

These files are **never** committed to the repository — see `.gitignore`.

---

## Architecture Notes

- **ClipboardMonitor** — daemon thread polling `pyperclip.paste()` via a timeout wrapper with exponential back-off on errors
- **HistoryStore** — RLock-protected in-memory store backed by JSON; generation-counter flush loop debounces disk writes
- **UI queue** — clipboard events cross the thread boundary via `queue.SimpleQueue`; drained on the Tk main thread every 150 ms
- **Shutdown sequence** — 8-step ordered teardown: stop polling → withdraw window → cancel timers → stop hotkey → stop monitor → drain queue → persist store → destroy Tk

---

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feat/your-feature`
3. Commit your changes: `git commit -m "feat: describe your change"`
4. Push the branch: `git push origin feat/your-feature`
5. Open a Pull Request

Please keep all application logic inside `src/clipboard-history-manager.py`. Do not commit `~/.clipboard_manager/` data or PyInstaller build artefacts.

---

## License

This project is licensed under the **MIT License** — see [LICENSE](LICENSE) for full text.

**Author:** Abhishek Srivastava
