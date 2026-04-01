"""
Bluebook Manager — Quality Alerts Global Panel.

Lists all Quality Alert files across all bluebooks, ordered by QA number.
Right-click to Open or Print a file, or navigate into its parent bluebook.

The DB query runs on a background QThread; table population is chunked via
QTimer so the UI never freezes even for large result sets.
"""

import os

from PySide6.QtCore import Qt, QThread, QTimer, Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from dal import dal
from services import file_service, print_service

_CHUNK_SIZE = 50   # rows to insert per event-loop tick


# ──────────────────────────────────────────────────────────────────────────────
# Background loader — runs the SQLite query off the main thread
# ──────────────────────────────────────────────────────────────────────────────

class _QALoaderThread(QThread):
    """Fetches QA records in a background thread to avoid freezing the UI."""

    results_ready = Signal(list, str)   # (records, original_search)

    def __init__(self, search: str, parent=None):
        super().__init__(parent)
        self._search = search

    def run(self):
        records = dal.get_all_quality_alerts(search=self._search)
        self.results_ready.emit(records, self._search)


# ──────────────────────────────────────────────────────────────────────────────
# QA Panel widget
# ──────────────────────────────────────────────────────────────────────────────

class QAPanel(QWidget):
    """Global Quality Alerts listing panel (Page 1 in main window content stack)."""

    # Emitted when user wants to jump to a bluebook detail view
    open_bluebook_requested = Signal(int)  # bluebook_id

    def __init__(self, parent=None):
        super().__init__(parent)
        self._loader: _QALoaderThread | None = None   # active background job
        self._pending_search: str | None = None        # latest requested search
        self._load_id: int = 0                         # incremented on every new load
        self._build_ui()

    # ──────────────────────────────────────────────────────────────────
    # UI Construction
    # ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        # Header row
        header_row = QHBoxLayout()
        title = QLabel("Quality Alerts")
        title.setObjectName("headerLabel")
        header_row.addWidget(title, 1)

        hint = QLabel("Double-click or right-click to Open / Print")
        hint.setStyleSheet("color: #6c7086; font-style: italic; font-size: 11px;")
        hint.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header_row.addWidget(hint)
        layout.addLayout(header_row)

        # Count / search indicator
        self.count_label = QLabel("")
        self.count_label.setObjectName("qaCountLabel")
        layout.addWidget(self.count_label)

        # QA table
        self.qa_table = QTableWidget()
        self.qa_table.setColumnCount(4)
        self.qa_table.setHorizontalHeaderLabels(
            ["QA Number", "Die#", "Customer(s)", "Date Added"])
        self.qa_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeToContents)
        self.qa_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeToContents)
        self.qa_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.Stretch)
        self.qa_table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeToContents)
        self.qa_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.qa_table.setSelectionMode(QTableWidget.SingleSelection)
        self.qa_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.qa_table.verticalHeader().setVisible(False)
        self.qa_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.qa_table.customContextMenuRequested.connect(self._on_context_menu)
        self.qa_table.doubleClicked.connect(self._open_selected_file)
        self.qa_table.setSortingEnabled(True)
        layout.addWidget(self.qa_table)

    # ──────────────────────────────────────────────────────────────────
    # Data Loading (async)
    # ──────────────────────────────────────────────────────────────────

    def load(self, search: str = ""):
        """Start an async load of QA records.

        If a previous load is still running, the new search is queued and
        executed as soon as the current thread finishes — so rapid keystrokes
        never pile up multiple blocking queries.
        """
        search = search.strip()

        if self._loader is not None and self._loader.isRunning():
            # A query is in flight. Queue the latest search (overwrites any prior pending).
            # Skip if we're already going to run this exact search next.
            if self._pending_search != search:
                self._pending_search = search
            return

        # Nothing running — start immediately
        self._start_loader(search)

    def _start_loader(self, search: str):
        """Spin up a fresh background thread for the given search term."""
        self._load_id += 1          # invalidate any in-progress chunked population
        self.count_label.setText("⏳  Loading Quality Alerts...")
        self._pending_search = None   # consumed

        # Clear table immediately so old data isn't visible while loading
        self.qa_table.setUpdatesEnabled(False)
        self.qa_table.setSortingEnabled(False)
        self.qa_table.setRowCount(0)
        self.qa_table.setUpdatesEnabled(True)

        self._loader = _QALoaderThread(search, parent=self)
        self._loader.results_ready.connect(self._on_results_ready)
        self._loader.start()

    def _on_results_ready(self, records: list, search: str):
        """Called on the main thread when the background query completes."""
        # Kick off chunked population — captures the load_id at this moment
        self._populate_table_chunked(records, search, self._load_id)

        # If another search arrived while we were loading, run it now
        if self._pending_search is not None:
            pending = self._pending_search
            self._start_loader(pending)

    def _populate_table_chunked(self, records: list, search: str, load_id: int):
        """Populate the table in small batches to keep the UI responsive.

        Each batch inserts _CHUNK_SIZE rows then yields back to the event loop
        via QTimer.singleShot(0, ...) so the window stays interactive.
        """
        total = len(records)
        self.qa_table.setUpdatesEnabled(False)
        self.qa_table.setSortingEnabled(False)
        self.qa_table.setRowCount(total)

        def _insert_chunk(start: int):
            # Bail out if a newer load has started
            if self._load_id != load_id:
                return

            end = min(start + _CHUNK_SIZE, total)
            for row in range(start, end):
                rec = records[row]
                filename = os.path.basename(rec["file_path"])
                if filename.lower().endswith(".lnk"):
                    filename = filename[:-4]
                display_name = (
                    os.path.splitext(filename)[0]
                    if filename.lower().endswith(".docx")
                    else filename
                )

                date_str = (rec["created_at"] or "")[:10]
                customers = ", ".join(rec["customer_names"])

                qa_item = QTableWidgetItem(display_name)
                qa_item.setData(Qt.UserRole, rec)

                self.qa_table.setItem(row, 0, qa_item)
                self.qa_table.setItem(row, 1, QTableWidgetItem(rec["die_number"]))
                self.qa_table.setItem(row, 2, QTableWidgetItem(customers))
                self.qa_table.setItem(row, 3, QTableWidgetItem(date_str))

            if end < total:
                # More rows to go — schedule next chunk without blocking
                self.count_label.setText(
                    f"⏳  Loading... ({end}/{total})")
                QTimer.singleShot(0, lambda: _insert_chunk(end))
            else:
                # All done
                self.qa_table.setSortingEnabled(True)
                self.qa_table.setUpdatesEnabled(True)
                if search:
                    self.count_label.setText(
                        f"🔎  {total} Quality Alert(s) matching \"{search}\"")
                else:
                    self.count_label.setText(f"{total} Quality Alert(s) total")

        _insert_chunk(0)

    # ──────────────────────────────────────────────────────────────────
    # Helpers
    # ──────────────────────────────────────────────────────────────────

    def _get_selected_record(self) -> dict | None:
        row = self.qa_table.currentRow()
        if row < 0:
            return None
        item = self.qa_table.item(row, 0)
        return item.data(Qt.UserRole) if item else None

    def _open_selected_file(self):
        rec = self._get_selected_record()
        if not rec:
            return
        try:
            file_service.open_file(rec["file_path"])
        except FileNotFoundError as e:
            QMessageBox.warning(self, "File Not Found", str(e))

    # ──────────────────────────────────────────────────────────────────
    # Context Menu
    # ──────────────────────────────────────────────────────────────────

    def _on_context_menu(self, pos):
        item = self.qa_table.itemAt(pos)
        if not item:
            return

        row = item.row()
        qa_item = self.qa_table.item(row, 0)
        if not qa_item:
            return
        rec = qa_item.data(Qt.UserRole)
        if not rec:
            return

        menu = QMenu(self)
        open_action = menu.addAction("📂  Open File")
        print_action = menu.addAction("🖨️  Print File")
        menu.addSeparator()
        goto_action = menu.addAction(
            f"📖  Open Bluebook (Die# {rec['die_number']})")

        action = menu.exec(self.qa_table.viewport().mapToGlobal(pos))
        if not action:
            return

        if action == open_action:
            try:
                file_service.open_file(rec["file_path"])
            except FileNotFoundError as e:
                QMessageBox.warning(self, "File Not Found", str(e))

        elif action == print_action:
            try:
                print_service.print_file(rec["file_path"])
                QMessageBox.information(self, "Print", "File sent to print.")
            except Exception as e:
                QMessageBox.warning(self, "Print Error", str(e))

        elif action == goto_action:
            self.open_bluebook_requested.emit(rec["bluebook_id"])
