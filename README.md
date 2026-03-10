# Bluebook Manager

A Windows desktop application for managing Bluebooks — die documentation packages organized by Die Number and Customer.

Built with **PySide6** (Qt for Python), **SQLite**, and **python-docx**.

## Features

- **Customer & Bluebook Management** — Create, edit, and delete customers and bluebooks with many-to-many linking
- **8-Section File Organization** — Cover, Master Drawings, QC Drawings, Approval, Quality Alerts, Quality Notes, Packing Instruction, Fit & Functions
- **File Preview** — Inline preview of PDFs (all pages), images, and DOCX files with rotate support. DOCX previews are cached and rendered via a background thread for performance
- **Drag & Drop** — Drag files from Explorer directly into the file list, or drag bluebooks onto customers for quick linking
- **File Sharing** — Share Quality Alerts, Quality Notes, and Fit & Functions across multiple bluebooks; shared files update the `Die:` field automatically. Search by customer name (`cust <name>`) to share to all of a customer's bluebooks at once
- **Quality Alert Naming Convention** — Auto-generated filenames following `QA-YY-NNN-DIENUM-description.docx`. NNN uses gap-filling: deleting a QA frees its number for reuse
- **Template System** — Create new documents from DOCX templates with auto-filled Customer, Die, Date, Q.A #, and Complaint fields
- **Context Menus** — Right-click files to Open, Print, Remove, Share, Unshare, or Rename
- **Print All** — Print every file in a bluebook in section order via the Windows print subsystem
- **Security** — Password-protected delete operations
- **Search & Filter** — Search bluebooks by die number or description; filter by customer
- **Activity Logging** — All operations logged to `logs/bluebook_manager.log`

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Run
python main.py
```

On first run, the app creates the SQLite database (`data/bluebook.db`) and the storage folder structure.

## Requirements

- **OS**: Windows (uses `pywin32` for printing, Word COM for DOCX preview)
- **Python**: 3.10+
- **Microsoft Word**: Required for DOCX preview (converts to PDF via COM automation)

## Project Structure

```
bluebook-project/
├── main.py                 # Entry point
├── config.py               # Paths, section types, file type rules
├── requirements.txt
├── dal/                    # Data Access Layer
│   ├── database.py         # SQLite schema & connection
│   ├── models.py           # Dataclasses (Customer, Bluebook, BluebookFile, etc.)
│   └── dal.py              # CRUD operations & QA counter
├── services/               # Business Logic
│   ├── bluebook_service.py # Bluebook create/delete with folder setup
│   ├── customer_service.py # Customer CRUD
│   ├── file_service.py     # Template creation, attach, rename, auto-fill
│   ├── sharing_service.py  # File sharing across bluebooks
│   ├── print_service.py    # Windows print integration
│   ├── security.py         # Password verification for delete ops
│   └── log_service.py      # File & console logging
├── ui/                     # PySide6 UI
│   ├── main_window.py      # Main window with customer sidebar & bluebook table
│   ├── bluebook_detail.py  # Detail view with sections, file list, preview pane
│   ├── customer_panel.py   # Customer sidebar panel
│   ├── resources/
│   │   └── styles.qss      # Application stylesheet
│   └── dialogs/
│       ├── share_dialog.py
│       ├── create_file_dialog.py
│       └── attach_file_dialog.py
├── scripts/                # Utility scripts
│   ├── import_quality_alerts.py
│   ├── import_packing.py
│   ├── import_fit_functions.py
│   └── export_quality_alerts.py
├── templates/              # DOCX templates
├── storage/                # Bluebook files organized by die number (gitignored)
├── data/                   # SQLite database (gitignored)
└── logs/                   # Application logs (gitignored)
```

## Architecture

```
┌─────────────────────────────────────────┐
│           PySide6 UI Layer              │
│  (Main Window, Detail View, Dialogs)    │
├─────────────────────────────────────────┤
│       Services / Business Logic         │
│  (Files, Sharing, Print, Security)      │
├─────────────────────────────────────────┤
│        Data Access Layer (DAL)          │
├─────────────────────────────────────────┤
│     SQLite DB  +  Disk File Storage     │
└─────────────────────────────────────────┘
```
