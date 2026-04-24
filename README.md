# Bluebook Manager

A Windows desktop application for managing **Bluebooks** — die documentation packages organized by Die Number and Customer.

Built with **PySide6** (Qt for Python), **SQLite**, **PyMuPDF**, and **python-docx**.

---

## Features

### Core Management
- **Customer & Bluebook Management** — Create, edit, and delete customers and bluebooks with many-to-many linking
- **8-Section File Organization** — Files are organized into ordered sections: Cover, Master Drawings, QC Drawings, Approval, Quality Alerts, Quality Notes, Packing Instruction, Fit & Functions
- **Outsourced Die Tracking** — Track and manage dies sent to external vendors via `outsource_service`

### File Handling
- **File Preview** — Inline preview of PDFs (all pages), images, and DOCX files with rotate support. DOCX previews are rendered via a background thread pool and cached to a temp directory for performance
- **Drag & Drop** — Drag files from Explorer directly into the file list; drag bluebooks onto customers for quick linking
- **Template System** — Create new DOCX documents from templates with auto-filled fields: Customer, Die Number, Date, QA #, and Complaint
- **Context Menus** — Right-click any file to Open, Print, Remove, Share, Unshare, or Rename

### Sharing & Collaboration
- **Cross-Bluebook File Sharing** — Share Quality Alerts, Quality Notes, Packing Instructions, and Fit & Functions across multiple bluebooks; shared files automatically update the `Die:` header field
- **Customer-wide Sharing** — Use the `cust <name>` search syntax in the Share dialog to share a file to all bluebooks belonging to a customer at once

### Quality Alerts
- **QA Panel** — Dedicated panel for browsing and managing all Quality Alerts across the database
- **Auto-naming Convention** — QA filenames follow the format `QA-YY-NNN-DIENUM-description.docx`
- **Gap-filling Counter** — Deleting a QA frees its sequence number (NNN) for reuse by the next created QA

### UI & Experience
- **Theme System** — Switch between multiple themes (e.g., dark/light); selection is persisted across sessions via `ThemeManager`
- **Branded Splash Screen** — Programmatic splash drawn with QPainter on startup — no external image asset required
- **Search & Filter** — Search bluebooks by die number or description; filter list by customer
- **Print All** — Print every file in a bluebook in section order via the Windows print subsystem

### Security & Logging
- **Password-protected Deletes** — Delete operations require password confirmation
- **Activity Logging** — All operations are logged to `logs/bluebook_manager.log`

---

## Quick Start

```bash
# 1. Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run
python main.py
```

On first run, the app auto-creates the SQLite database (`data/bluebook.db`) and the full storage folder structure.

---

## Requirements

| Requirement | Details |
|---|---|
| **OS** | Windows only (`pywin32` for printing, Word COM for DOCX preview) |
| **Python** | 3.10+ |
| **Microsoft Word** | Required for DOCX preview (COM automation converts DOCX → PDF) |

### Python Dependencies

| Package | Purpose |
|---|---|
| `PySide6 >= 6.5` | Qt UI framework |
| `python-docx >= 0.8.11` | DOCX creation and template filling |
| `pywin32 >= 306` | Windows print subsystem & Word COM |
| `PyMuPDF >= 1.23` | PDF rendering for the preview pane |
| `openpyxl >= 3.1` | Excel read/write for import scripts |

---

## Project Structure

```
bluebook-project/
├── main.py                         # Entry point — splash, DB init, window launch
├── config.py                       # All paths, section types, file-type rules
├── requirements.txt
├── settings.json                   # Persisted settings (e.g., current theme)
├── BluebookManager.spec            # PyInstaller build spec
├── compile.bat / compile.sh        # Build helpers
│
├── dal/                            # Data Access Layer
│   ├── database.py                 # SQLite schema & connection bootstrap
│   ├── models.py                   # Dataclasses: Customer, Bluebook, BluebookFile, etc.
│   └── dal.py                      # All CRUD operations + QA sequence counter
│
├── services/                       # Business Logic
│   ├── bluebook_service.py         # Bluebook create/delete + folder scaffolding
│   ├── customer_service.py         # Customer CRUD
│   ├── file_service.py             # Template creation, attach, rename, auto-fill
│   ├── sharing_service.py          # Cross-bluebook file sharing logic
│   ├── outsource_service.py        # Outsourced die tracking
│   ├── print_service.py            # Windows print integration (section order)
│   ├── security.py                 # Password verification for destructive ops
│   ├── theme_manager.py            # Theme loading, switching, and persistence
│   └── log_service.py              # File & console logging setup
│
├── ui/                             # PySide6 UI Layer
│   ├── main_window.py              # Main window: customer sidebar + bluebook table
│   ├── bluebook_detail.py          # Detail view: sections, file list, preview pane
│   ├── customer_panel.py           # Customer sidebar panel
│   ├── qa_panel.py                 # Quality Alerts cross-bluebook panel
│   ├── qa_window.py                # Standalone QA viewer window
│   ├── resources/
│   │   └── styles.qss              # Application stylesheet (themed)
│   ├── dialogs/
│   │   ├── share_dialog.py
│   │   ├── create_file_dialog.py
│   │   └── attach_file_dialog.py
│   └── widgets/                    # Reusable custom widgets
│
├── scripts/                        # One-off data migration / utility scripts
│   ├── import_quality_alerts.py
│   ├── import_packing.py
│   ├── import_fit_functions.py
│   ├── import_approval_forms.py
│   ├── import_customer_dies.py     # Multi-DB Excel importer with deduplication
│   ├── import_future_line_dies.py
│   ├── export_quality_alerts.py
│   ├── auto_link_files.py
│   ├── sync_storage.py
│   └── get-filenames.py
│
├── templates/                      # DOCX templates for document creation
├── storage/                        # Bluebook files on disk, by die number (gitignored)
├── data/                           # SQLite database file (gitignored)
└── logs/                           # Application logs (gitignored)
```

---

## Architecture

```
┌──────────────────────────────────────────────────────┐
│                    PySide6 UI Layer                   │
│   MainWindow · BluebookDetail · QAPanel · Dialogs    │
├──────────────────────────────────────────────────────┤
│              Services / Business Logic                │
│   FileService · SharingService · PrintService        │
│   ThemeManager · OutsourceService · Security         │
├──────────────────────────────────────────────────────┤
│             Data Access Layer  (DAL)                  │
│         dal.py · models.py · database.py             │
├──────────────────────────────────────────────────────┤
│          SQLite Database  +  Disk Storage             │
│            data/bluebook.db   storage/               │
└──────────────────────────────────────────────────────┘
```

### Key Design Decisions

- **Layered architecture** — UI never touches the DB directly; all mutations go through Services → DAL.
- **Background thread pools** — DOCX→PDF conversion (Word COM) and preview rendering run in separate thread pools to keep the UI responsive. Both pools are shut down cleanly on `app.aboutToQuit`.
- **Temp cache** — Rendered DOCX previews are written to `%TEMP%\bluebook_docx_cache\` and purged on exit.
- **PyInstaller dual-path** — `config.py` distinguishes between `sys._MEIPASS` (bundled assets) and the directory next to the `.exe` (user data: DB, storage, logs).

---

## Building the Executable

```bat
compile.bat
```

This runs PyInstaller using `BluebookManager.spec`. The output `.exe` is placed in `dist/`.

> **Note:** Microsoft Word must be installed on the target machine for DOCX preview to work.

---

## Utility Scripts

All scripts in `scripts/` are run directly from the project root (so `config` is importable):

```bash
# From the project root:
python scripts/import_quality_alerts.py
python scripts/export_quality_alerts.py
python scripts/sync_storage.py
```

| Script | Purpose |
|---|---|
| `import_customer_dies.py` | Bulk-import outsourced die records from Excel with multi-DB selection |
| `import_quality_alerts.py` | Import QA DOCX files into the database |
| `import_packing.py` | Import packing instruction files |
| `import_fit_functions.py` | Import fit & function files |
| `import_approval_forms.py` | Import approval PDFs |
| `import_future_line_dies.py` | Import future-line die records |
| `export_quality_alerts.py` | Export QA records to Excel |
| `auto_link_files.py` | Auto-link orphaned storage files to bluebook records |
| `sync_storage.py` | Reconcile DB records with files on disk |
| `get-filenames.py` | List file names in storage for inspection |
