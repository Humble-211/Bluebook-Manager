"""
Bluebook Manager — Quality Alerts Search Window.

An independent top-level window for browsing and acting on all Quality Alert
files across every bluebook. Opens alongside the main window (non-blocking).
"""

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QVBoxLayout,
    QWidget,
)

from ui.qa_panel import QAPanel


class QAWindow(QMainWindow):
    """Standalone Quality Alerts browser window.

    Design decisions:
    - Non-modal, fully independent — user can keep it open on a second monitor.
    - Singleton managed by MainWindow: calling open() raises an existing instance
      instead of spawning duplicates.
    - Delegates all data work to QAPanel (async background loader).
    - Exposes open_bluebook_requested so MainWindow can navigate there.
    """

    open_bluebook_requested = Signal(int)  # re-emitted from QAPanel

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Quality Alerts Search")
        self.setMinimumSize(900, 600)
        self.resize(1150, 720)
        # Keep window above main window but don't force always-on-top system-wide
        self.setWindowFlag(Qt.Window, True)

        self._search_timer = QTimer()
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._do_search)

        self._build_ui()

    # ──────────────────────────────────────────────────────────────────
    # UI Construction
    # ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        # ── Search row ──
        search_row = QHBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(
            "🔍  Search by QA number (e.g. QA-25-001), die#, or any keyword...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._on_search_changed)
        search_row.addWidget(self.search_input, 1)

        layout.addLayout(search_row)

        # ── QA listing panel ──
        self.qa_panel = QAPanel()
        self.qa_panel.open_bluebook_requested.connect(
            self.open_bluebook_requested)
        layout.addWidget(self.qa_panel)

        # Status bar hint
        self.statusBar().showMessage(
            "Double-click to open a file  •  Right-click for more options")

    # ──────────────────────────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────────────────────────

    def activate(self):
        """Show the window and bring it to the front."""
        self.show()
        self.raise_()
        self.activateWindow()
        # Trigger initial load if the table is empty
        if self.qa_panel.qa_table.rowCount() == 0:
            self._do_search()

    # ──────────────────────────────────────────────────────────────────
    # Search Handling
    # ──────────────────────────────────────────────────────────────────

    def _on_search_changed(self, _text: str):
        """Debounce search input — 150 ms after last keystroke."""
        self._search_timer.start(300)

    def _do_search(self):
        search = self.search_input.text().strip()
        self.qa_panel.load(search)
