"""
Bluebook Manager — Main Window.

Contains the search bar, customer sidebar, bluebook list,
and navigation to the bluebook detail screen.
"""

import json
import os

from PySide6.QtWidgets import (
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtCore import Qt, QTimer, QMimeData, QByteArray

from config import SECTION_LABELS
from services import bluebook_service, customer_service
from ui.bluebook_detail import BluebookDetailWidget
from ui.customer_panel import CustomerPanel


class BluebookTable(QTableWidget):
    """QTableWidget that encodes selected bluebook IDs in drag mime data."""

    def mimeTypes(self):
        return ["application/x-bluebook-ids"]

    def mimeData(self, items):
        # Collect unique bluebook IDs from selected rows
        bb_ids = []
        seen = set()
        for item in items:
            row = item.row()
            if row in seen:
                continue
            seen.add(row)
            col0 = self.item(row, 0)
            if col0:
                bb_id = col0.data(Qt.UserRole)
                if bb_id is not None:
                    bb_ids.append(bb_id)

        mime = QMimeData()
        mime.setData(
            "application/x-bluebook-ids",
            QByteArray(json.dumps(bb_ids).encode()))
        return mime

    def startDrag(self, supportedActions):
        """Show a compact label with die numbers as the drag pixmap."""
        from PySide6.QtGui import QDrag, QPixmap, QPainter, QFont, QColor
        items = self.selectedItems()
        mime = self.mimeData(items)
        if not mime:
            return

        # Collect die numbers for display
        die_nums = []
        seen = set()
        for item in items:
            row = item.row()
            if row in seen:
                continue
            seen.add(row)
            col0 = self.item(row, 0)
            if col0:
                die_nums.append(col0.text())

        if len(die_nums) <= 3:
            label_text = f"  Die# {', '.join(die_nums)}  "
        else:
            label_text = f"  Die# {', '.join(die_nums[:3])}... (+{len(die_nums)-3})  "

        # Render a styled label to pixmap
        tmp = QLabel(label_text)
        tmp.setStyleSheet(
            "background-color: #45475a; color: #cdd6f4; "
            "border: 1px solid #89b4fa; border-radius: 6px; "
            "padding: 4px 8px; font-size: 12pt;")
        tmp.adjustSize()
        pixmap = tmp.grab()

        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.setPixmap(pixmap)
        drag.setHotSpot(pixmap.rect().center())
        drag.exec(supportedActions)


class MainWindow(QMainWindow):
    """Application main window."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bluebook Manager")
        self.setMinimumSize(1100, 700)
        self.resize(1280, 800)

        self.current_customer_id = None  # None = show all

        self._build_ui()
        self.statusBar().showMessage("Search or select a customer to view bluebooks")

    def _build_ui(self):
        # Central stacked widget (main screen vs detail screen)
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        # ── Page 0: Main Screen ──
        main_page = QWidget()
        main_layout = QHBoxLayout(main_page)

        # Left sidebar: customers
        self.customer_panel = CustomerPanel()
        self.customer_panel.setFixedWidth(250)
        self.customer_panel.customer_selected.connect(self._on_customer_selected)
        self.customer_panel.show_all.connect(self._on_show_all)
        self.customer_panel.bluebooks_dropped.connect(self._on_bluebooks_dropped)
        main_layout.addWidget(self.customer_panel)

        # Right: search + bluebook list
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 0, 0, 0)

        # Title row
        title_row = QHBoxLayout()
        title = QLabel("Bluebook Manager")
        title.setObjectName("headerLabel")
        title_row.addWidget(title, 1)
        right_layout.addLayout(title_row)

        # Search bar row
        search_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍  Search by Die Number (use , for multiple)  |  prefix 'desc ' to search by description")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._on_search_changed)
        search_row.addWidget(self.search_input, 1)

        btn_new = QPushButton("+ New Bluebook")
        btn_new.setObjectName("successButton")
        btn_new.clicked.connect(self._create_bluebook)
        search_row.addWidget(btn_new)


        right_layout.addLayout(search_row)

        # Customer filter indicator
        self.filter_label = QLabel("")
        self.filter_label.setStyleSheet("color: #f9e2af; font-style: italic; padding: 4px 0;")
        right_layout.addWidget(self.filter_label)

        # Bluebook table
        self.bluebook_table = BluebookTable()
        self.bluebook_table.setColumnCount(4)
        self.bluebook_table.setHorizontalHeaderLabels(
            ["Die Number", "Description", "Customers", "Files"])
        self.bluebook_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeToContents)
        self.bluebook_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.Stretch)
        self.bluebook_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.Stretch)
        self.bluebook_table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeToContents)
        self.bluebook_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.bluebook_table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.bluebook_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.bluebook_table.setDragEnabled(True)
        self.bluebook_table.setDragDropMode(QTableWidget.DragOnly)
        self.bluebook_table.doubleClicked.connect(self._open_selected_bluebook)
        self.bluebook_table.verticalHeader().setVisible(False)
        self.bluebook_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.bluebook_table.customContextMenuRequested.connect(self._on_table_context_menu)
        right_layout.addWidget(self.bluebook_table)

        # Open button
        btn_open = QPushButton("📂 Open Bluebook")
        btn_open.setObjectName("primaryButton")
        btn_open.clicked.connect(self._open_selected_bluebook)
        right_layout.addWidget(btn_open)

        main_layout.addWidget(right_widget, 1)

        self.stack.addWidget(main_page)

        # ── Page 1: Detail (added dynamically) ──

        # Status bar
        self.statusBar().showMessage("Ready")

        # Debounce timer for search
        self._search_timer = QTimer()
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._do_search)

    def _on_search_changed(self, text):
        """Debounce search input."""
        # Update placeholder to hint at current search mode
        if text.lower().startswith("desc ") or text.lower() == "desc":
            self.search_input.setPlaceholderText(
                "🔍  Searching by Description...")
        else:
            self.search_input.setPlaceholderText(
                "🔍  Search by Die Number (use , for multiple)  |  prefix 'desc ' to search by description")
        self._search_timer.start(300)

    def _do_search(self):
        self._load_bluebooks()

    def _on_customer_selected(self, customer_id):
        self.current_customer_id = customer_id
        c = customer_service.get_customer(customer_id)
        self.filter_label.setText(f"Showing bluebooks for: {c.name}" if c else "")
        self._load_bluebooks()

    def _on_show_all(self):
        self.current_customer_id = None
        self.filter_label.setText("")
        self._load_bluebooks()

    def _load_bluebooks(self):
        """Load bluebooks into the table, applying search and customer filter."""
        raw = self.search_input.text().strip()

        # Detect "desc " prefix → search by description mode
        search_description = False
        if raw.lower().startswith("desc "):
            search_description = True
            raw = raw[5:].strip()  # strip the "desc " prefix

        # Show nothing unless user has a search term or customer filter
        if not raw and not self.current_customer_id:
            self.bluebook_table.setRowCount(0)
            self.statusBar().showMessage(
                "Search or select a customer to view bluebooks")
            return

        # Support comma-separated search terms
        terms = [t.strip() for t in raw.split(",") if t.strip()]

        if len(terms) <= 1:
            # Single term (or empty) — use existing path
            search = terms[0] if terms else ""
            bluebooks = bluebook_service.search_bluebooks(
                search=search, customer_id=self.current_customer_id,
                search_description=search_description)
        else:
            # Multiple terms — search each and merge (preserving order, no dupes)
            seen = set()
            bluebooks = []
            for term in terms:
                results = bluebook_service.search_bluebooks(
                    search=term, customer_id=self.current_customer_id,
                    search_description=search_description)
                for bb in results:
                    if bb.id not in seen:
                        seen.add(bb.id)
                        bluebooks.append(bb)

        # Batch-fetch file counts in ONE query
        from dal import dal
        bb_ids = [bb.id for bb in bluebooks]
        file_counts = dal.get_file_counts_batch(bb_ids)

        # Freeze UI during bulk insert
        self.bluebook_table.setUpdatesEnabled(False)
        self.bluebook_table.setSortingEnabled(False)

        self.bluebook_table.setRowCount(len(bluebooks))
        for row, bb in enumerate(bluebooks):
            self.bluebook_table.setItem(row, 0, QTableWidgetItem(bb.die_number))
            self.bluebook_table.setItem(row, 1, QTableWidgetItem(bb.description or ""))
            self.bluebook_table.setItem(
                row, 2, QTableWidgetItem(", ".join(bb.customer_names)))
            self.bluebook_table.setItem(
                row, 3, QTableWidgetItem(str(file_counts.get(bb.id, 0))))

            # Store bluebook id
            self.bluebook_table.item(row, 0).setData(Qt.UserRole, bb.id)

        self.bluebook_table.setSortingEnabled(True)
        self.bluebook_table.setUpdatesEnabled(True)

        mode_label = " (by description)" if search_description else ""
        if len(bluebooks) >= 200:
            self.statusBar().showMessage(
                f"Showing first 200 bluebooks{mode_label} — use search to find more")
        else:
            self.statusBar().showMessage(
                f"{len(bluebooks)} bluebook(s) found{mode_label}")

    def _open_selected_bluebook(self):
        row = self.bluebook_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "No Selection",
                                    "Please select a bluebook to open.")
            return
        bb_id = self.bluebook_table.item(row, 0).data(Qt.UserRole)
        self._show_detail(bb_id)

    def _show_detail(self, bluebook_id):
        """Navigate to the bluebook detail screen."""
        detail = BluebookDetailWidget(bluebook_id)
        detail.closed.connect(self._return_to_main)

        # Remove old detail page if exists
        while self.stack.count() > 1:
            w = self.stack.widget(1)
            self.stack.removeWidget(w)
            w.deleteLater()

        self.stack.addWidget(detail)
        self.stack.setCurrentIndex(1)

    def _return_to_main(self):
        """Return to the main screen."""
        self.stack.setCurrentIndex(0)

        # Refresh customer list without triggering a second _load_bluebooks
        self.customer_panel.blockSignals(True)
        self.customer_panel.refresh()
        self.customer_panel.blockSignals(False)

        # Reset to empty state — user must search or pick a customer
        self.current_customer_id = None
        self.filter_label.setText("")
        self.search_input.clear()
        self.bluebook_table.setRowCount(0)
        self.statusBar().showMessage(
            "Search or select a customer to view bluebooks")

        # Cleanup detail widget
        while self.stack.count() > 1:
            w = self.stack.widget(1)
            self.stack.removeWidget(w)
            w.deleteLater()

    def _create_bluebook(self):
        die, ok = QInputDialog.getText(self, "New Bluebook", "Die Number:")
        if not ok or not die.strip():
            return

        desc, ok2 = QInputDialog.getText(
            self, "New Bluebook", "Description (optional):")
        desc = desc.strip() if ok2 else ""

        try:
            bb = bluebook_service.create_bluebook(die.strip(), desc)

            # Optionally link to current customer
            if self.current_customer_id:
                customer_service.link_bluebook(self.current_customer_id, bb.id)

            self._load_bluebooks()
            self.statusBar().showMessage(f"Created Bluebook Die# {bb.die_number}")
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))

    def _delete_bluebook(self):
        # Collect unique selected bluebook IDs
        selected_rows = set()
        for item in self.bluebook_table.selectedItems():
            selected_rows.add(item.row())

        if not selected_rows:
            QMessageBox.information(self, "No Selection",
                                    "Please select a bluebook to delete.")
            return

        bb_list = []
        for row in sorted(selected_rows):
            bb_id = self.bluebook_table.item(row, 0).data(Qt.UserRole)
            bb = bluebook_service.get_bluebook(bb_id)
            if bb:
                bb_list.append(bb)

        if not bb_list:
            return

        # Require password for deletion
        from services.security import require_password
        if not require_password("Deleting bluebook(s)", self):
            return

        die_nums = ", ".join(bb.die_number for bb in bb_list)
        if len(bb_list) == 1:
            msg = (f"Delete Bluebook Die# {die_nums}?\n\n"
                   f"This will remove the database record and all associated files.")
        else:
            msg = (f"Delete {len(bb_list)} bluebooks?\n\n"
                   f"Die#: {die_nums}\n\n"
                   f"This will remove all database records and associated files.")

        reply = QMessageBox.question(
            self, "Delete Bluebook",
            msg,
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            for bb in bb_list:
                bluebook_service.delete_bluebook(bb.id, delete_files=True)
            self._load_bluebooks()
            self.statusBar().showMessage(f"Deleted {len(bb_list)} bluebook(s): {die_nums}")

    def _on_table_context_menu(self, pos):
        """Show context menu on right-click in the bluebook table."""
        item = self.bluebook_table.itemAt(pos)
        if not item:
            return

        row = item.row()
        bb_id = self.bluebook_table.item(row, 0).data(Qt.UserRole)
        bb = bluebook_service.get_bluebook(bb_id)
        if not bb:
            return

        menu = QMenu(self)

        # Add Customer action
        add_action = menu.addAction("➕  Add Customer to this Bluebook")
        remove_action = None

        # Remove Customer action (only if bluebook has customers)
        if bb.customer_names:
            remove_action = menu.addAction("➖  Remove Customer from this Bluebook")

        menu.addSeparator()
        open_action = menu.addAction("📂  Open Bluebook")
        desc_action = menu.addAction("✏️  Change Description")
        menu.addSeparator()
        delete_action = menu.addAction("🗑️  Delete Bluebook")

        action = menu.exec(self.bluebook_table.viewport().mapToGlobal(pos))
        if not action:
            return

        if action == add_action:
            self._add_customer_to_bluebook(bb)
        elif action == remove_action:
            self._remove_customer_from_bluebook(bb)
        elif action == open_action:
            self._show_detail(bb.id)
        elif action == desc_action:
            self._change_description(bb)
        elif action == delete_action:
            self._delete_bluebook()

    def _add_customer_to_bluebook(self, bb):
        """Show a list of customers to add to this bluebook."""
        all_customers = customer_service.list_customers()
        current_customers = customer_service.get_customers_for_bluebook(bb.id)
        current_ids = {c.id for c in current_customers}

        # Filter to customers not already linked
        available = [c for c in all_customers if c.id not in current_ids]

        if not available:
            QMessageBox.information(
                self, "No Available Customers",
                "All customers are already linked to this bluebook.\n"
                "You can add a new customer from the sidebar first.")
            return

        names = [c.name for c in available]
        name, ok = QInputDialog.getItem(
            self, "Add Customer",
            f"Select customer to add to Die# {bb.die_number}:",
            names, 0, False)

        if ok and name:
            customer = next(c for c in available if c.name == name)
            customer_service.link_bluebook(customer.id, bb.id)
            self._load_bluebooks()
            self.statusBar().showMessage(
                f"Added '{name}' to Die# {bb.die_number}")

    def _remove_customer_from_bluebook(self, bb):
        """Show a list of linked customers to remove from this bluebook."""
        current_customers = customer_service.get_customers_for_bluebook(bb.id)
        if not current_customers:
            return

        names = [c.name for c in current_customers]
        name, ok = QInputDialog.getItem(
            self, "Remove Customer",
            f"Select customer to remove from Die# {bb.die_number}:",
            names, 0, False)

        if ok and name:
            customer = next(c for c in current_customers if c.name == name)
            reply = QMessageBox.question(
                self, "Confirm Remove",
                f"Remove '{name}' from Die# {bb.die_number}?",
                QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                customer_service.unlink_bluebook(customer.id, bb.id)
                self._load_bluebooks()
                self.statusBar().showMessage(
                    f"Removed '{name}' from Die# {bb.die_number}")

    def _change_description(self, bb):
        """Prompt user to change the bluebook description."""
        new_desc, ok = QInputDialog.getText(
            self, "Change Description",
            f"New description for Die# {bb.die_number}:",
            text=bb.description or "")
        if ok:
            bluebook_service.update_bluebook(bb.id, bb.die_number, new_desc.strip())
            self._load_bluebooks()
            self.statusBar().showMessage(
                f"Updated description for Die# {bb.die_number}")

    def _on_bluebooks_dropped(self, customer_id, bb_ids):
        """Link dropped bluebooks to the target customer."""
        c = customer_service.get_customer(customer_id)
        if not c:
            return

        linked = 0
        for bb_id in bb_ids:
            try:
                customer_service.link_bluebook(customer_id, bb_id)
                linked += 1
            except Exception:
                pass  # Already linked — skip silently

        self._load_bluebooks()
        self.statusBar().showMessage(
            f"Linked {linked} bluebook(s) to '{c.name}'")
