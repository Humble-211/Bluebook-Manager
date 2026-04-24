"""
Bluebook Manager - Main Window.

Contains the search bar, customer sidebar, bluebook list,
and navigation to the bluebook detail screen.
"""

import json
import os
import re

from PySide6.QtCore import QByteArray, QMimeData, Qt, QTimer
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from services import bluebook_service, customer_service, outsource_service
from ui.bluebook_detail import BluebookDetailWidget
from ui.customer_panel import CustomerPanel
from ui.qa_window import QAWindow


def _natural_sort_key(text: str):
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


class NaturalSortTableItem(QTableWidgetItem):
    """Table item that sorts die numbers in natural numeric order."""

    def __lt__(self, other):
        return _natural_sort_key(self.text()) < _natural_sort_key(other.text())


class BluebookTable(QTableWidget):
    """QTableWidget that encodes selected bluebook IDs in drag mime data."""

    def mimeTypes(self):
        return ["application/x-bluebook-ids"]

    def mimeData(self, items):
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
        mime.setData("application/x-bluebook-ids", QByteArray(json.dumps(bb_ids).encode()))
        return mime

    def startDrag(self, supportedActions):
        """Show a compact label with die numbers as the drag pixmap."""
        from PySide6.QtGui import QDrag

        items = self.selectedItems()
        mime = self.mimeData(items)
        if not mime:
            return

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
            label_text = f"  Die# {', '.join(die_nums[:3])}... (+{len(die_nums) - 3})  "

        tmp = QLabel(label_text)
        tmp.setStyleSheet(
            "background-color: #45475a; color: #cdd6f4; "
            "border: 1px solid #89b4fa; border-radius: 6px; "
            "padding: 4px 8px; font-size: 12pt;"
        )
        tmp.adjustSize()
        pixmap = tmp.grab()

        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.setPixmap(pixmap)
        drag.setHotSpot(pixmap.rect().center())
        drag.exec(supportedActions)


class MainWindow(QMainWindow):
    """Application main window."""

    def __init__(self, theme_manager=None):
        super().__init__()
        self.setWindowTitle("Bluebook Manager")
        self.setMinimumSize(1100, 700)
        self.resize(1280, 800)

        self.current_customer_id = None
        self._qa_window: QAWindow | None = None
        self._theme_manager = theme_manager

        self._build_ui()
        self._load_bluebooks()
        self.statusBar().showMessage("Search or select a customer to view bluebooks")

        from PySide6.QtGui import QKeySequence, QShortcut

        self.console_shortcut = QShortcut(QKeySequence("Ctrl+Shift+9"), self)
        self.console_shortcut.activated.connect(self._toggle_console)

    def _build_ui(self):
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        main_page = QWidget()
        main_layout = QHBoxLayout(main_page)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(12)

        sidebar_card = QFrame()
        sidebar_card.setObjectName("surfaceCard")
        sidebar_card.setFixedWidth(240)
        sidebar_layout = QVBoxLayout(sidebar_card)
        sidebar_layout.setContentsMargins(10, 10, 10, 10)
        sidebar_layout.setSpacing(8)

        self.customer_panel = CustomerPanel()
        self.customer_panel.customer_selected.connect(self._on_customer_selected)
        self.customer_panel.show_all.connect(self._on_show_all)
        self.customer_panel.bluebooks_dropped.connect(self._on_bluebooks_dropped)
        sidebar_layout.addWidget(self.customer_panel)
        main_layout.addWidget(sidebar_card)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        toolbar_card = QFrame()
        toolbar_card.setObjectName("surfaceCard")
        toolbar_layout = QVBoxLayout(toolbar_card)
        toolbar_layout.setContentsMargins(12, 10, 12, 10)
        toolbar_layout.setSpacing(8)

        title_row = QHBoxLayout()
        title_row.setSpacing(8)
        title = QLabel("Bluebooks")
        title.setObjectName("headerLabel")
        title_row.addWidget(title)
        title_row.addStretch(1)

        if self._theme_manager:
            self._theme_btn = QPushButton(self._theme_label())
            self._theme_btn.setObjectName("ghostButton")
            self._theme_btn.setToolTip("Click to switch theme")
            self._theme_btn.setFixedHeight(32)
            self._theme_btn.clicked.connect(self._cycle_theme)
            title_row.addWidget(self._theme_btn)
        toolbar_layout.addLayout(title_row)

        btn_new = QPushButton("+ New Bluebook")
        btn_new.setObjectName("successButton")
        btn_new.setMinimumHeight(34)
        btn_new.clicked.connect(self._create_bluebook)

        search_row = QHBoxLayout()
        search_row.setSpacing(8)
        self.search_input = QLineEdit()
        self.search_input.setObjectName("heroSearch")
        self.search_input.setPlaceholderText(
            "Search by die number (use commas for multiple) or prefix with 'desc ' for descriptions"
        )
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._on_search_changed)
        search_row.addWidget(self.search_input, 1)

        btn_qa = QPushButton("Quality Alerts")
        btn_qa.setObjectName("primaryButton")
        btn_qa.setMinimumHeight(34)
        btn_qa.setToolTip("Open the Quality Alerts browser window")
        btn_qa.clicked.connect(self._open_qa_window)
        search_row.addWidget(btn_new)
        search_row.addWidget(btn_qa)
        toolbar_layout.addLayout(search_row)

        self.filter_label = QLabel("Search or select a customer to start browsing.")
        self.filter_label.setObjectName("infoPill")
        toolbar_layout.addWidget(self.filter_label)
        right_layout.addWidget(toolbar_card)

        table_card = QFrame()
        table_card.setObjectName("surfaceCard")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(12, 12, 12, 12)
        table_layout.setSpacing(8)

        table_header = QHBoxLayout()
        table_title = QLabel("Bluebooks")
        table_title.setObjectName("panelTitle")
        table_header.addWidget(table_title)
        table_header.addStretch(1)

        self.results_summary = QLabel("Waiting for search")
        self.results_summary.setObjectName("panelCaption")
        table_header.addWidget(self.results_summary)

        self._open_btn = QPushButton("Open Selected")
        self._open_btn.setObjectName("primaryButton")
        self._open_btn.setMinimumHeight(32)
        self._open_btn.clicked.connect(self._open_selected_bluebook)
        table_header.addWidget(self._open_btn)
        table_layout.addLayout(table_header)

        self.bluebook_table = BluebookTable()
        self.bluebook_table.setObjectName("dataTable")
        self.bluebook_table.setColumnCount(5)
        self.bluebook_table.setHorizontalHeaderLabels(
            ["Die Number", "Description", "Customers", "Outsource", "Files"]
        )
        self.bluebook_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.bluebook_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.bluebook_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.bluebook_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.bluebook_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.bluebook_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.bluebook_table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.bluebook_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.bluebook_table.setDragEnabled(True)
        self.bluebook_table.setDragDropMode(QTableWidget.DragOnly)
        self.bluebook_table.setAlternatingRowColors(True)
        self.bluebook_table.setShowGrid(False)
        self.bluebook_table.doubleClicked.connect(self._open_selected_bluebook)
        self.bluebook_table.itemSelectionChanged.connect(self._update_table_action_state)
        self.bluebook_table.verticalHeader().setVisible(False)
        self.bluebook_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.bluebook_table.customContextMenuRequested.connect(self._on_table_context_menu)
        table_layout.addWidget(self.bluebook_table)
        right_layout.addWidget(table_card, 1)

        main_layout.addWidget(right_widget, 1)
        self.stack.addWidget(main_page)

        self.statusBar().showMessage("Ready")

        self._search_timer = QTimer()
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._do_search)
        self._update_table_action_state()

    def _update_table_action_state(self):
        has_selection = self.bluebook_table.currentRow() >= 0
        if hasattr(self, "_open_btn"):
            self._open_btn.setEnabled(has_selection)

    def _on_search_changed(self, text):
        lower = text.lower().strip()
        if lower.startswith("desc ") or lower == "desc":
            self.search_input.setPlaceholderText("Searching by description")
        else:
            self.search_input.setPlaceholderText(
                "Search by die number (use commas for multiple) or prefix with 'desc ' for descriptions"
            )
        self._search_timer.start(150)

    def _do_search(self):
        self._load_bluebooks()

    def _on_customer_selected(self, customer_id):
        self.current_customer_id = customer_id
        c = customer_service.get_customer(customer_id)
        self.filter_label.setText(f"Showing bluebooks for: {c.name}" if c else "")
        self._load_bluebooks()

    def _on_show_all(self):
        self.current_customer_id = None
        self.filter_label.setText("Search or select a customer to start browsing.")
        self._load_bluebooks()

    def _load_bluebooks(self):
        raw = self.search_input.text().strip()

        search_description = False
        if raw.lower().startswith("desc "):
            search_description = True
            raw = raw[5:].strip()

        terms = [t.strip() for t in raw.split(",") if t.strip()]

        if len(terms) <= 1:
            search = terms[0] if terms else ""
            bluebooks = bluebook_service.search_bluebooks(
                search=search,
                customer_id=self.current_customer_id,
                search_description=search_description,
                limit=0,
            )
        else:
            seen = set()
            bluebooks = []
            for term in terms:
                results = bluebook_service.search_bluebooks(
                    search=term,
                    customer_id=self.current_customer_id,
                    search_description=search_description,
                    limit=0,
                )
                for bb in results:
                    if bb.id not in seen:
                        seen.add(bb.id)
                        bluebooks.append(bb)

        from dal import dal

        bb_ids = [bb.id for bb in bluebooks]
        file_counts = dal.get_file_counts_batch(bb_ids)

        self.bluebook_table.setUpdatesEnabled(False)
        self.bluebook_table.setSortingEnabled(False)
        self.bluebook_table.setRowCount(len(bluebooks))

        for row, bb in enumerate(bluebooks):
            self.bluebook_table.setItem(row, 0, NaturalSortTableItem(bb.die_number))
            self.bluebook_table.setItem(row, 1, QTableWidgetItem(bb.description or ""))
            self.bluebook_table.setItem(row, 2, QTableWidgetItem(", ".join(bb.customer_names)))
            self.bluebook_table.setItem(row, 3, QTableWidgetItem(", ".join(bb.outsource_names)))
            self.bluebook_table.setItem(row, 4, QTableWidgetItem(str(file_counts.get(bb.id, 0))))
            self.bluebook_table.item(row, 0).setData(Qt.UserRole, bb.id)

        self.bluebook_table.setSortingEnabled(True)
        self.bluebook_table.setUpdatesEnabled(True)
        self._update_table_action_state()

        mode_label = " (by description)" if search_description else ""
        if bluebooks:
            self.results_summary.setText(
                f"{len(bluebooks)} bluebook(s) loaded{mode_label}. Double-click a row or use Open Selected."
            )
        else:
            self.results_summary.setText(
                "No bluebooks matched the current search. Try a different die number, description, or customer."
            )
        self.statusBar().showMessage(f"{len(bluebooks)} bluebook(s) found{mode_label}")

    def _open_selected_bluebook(self):
        row = self.bluebook_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "No Selection", "Please select a bluebook to open.")
            return
        bb_id = self.bluebook_table.item(row, 0).data(Qt.UserRole)
        self._show_detail(bb_id)

    def _show_detail(self, bluebook_id):
        detail = BluebookDetailWidget(bluebook_id)
        detail.closed.connect(lambda: self._return_to_main(bluebook_id))

        while self.stack.count() > 1:
            w = self.stack.widget(1)
            self.stack.removeWidget(w)
            w.deleteLater()

        self.stack.addWidget(detail)
        self.stack.setCurrentIndex(1)

    def _return_to_main(self, bluebook_id=None):
        self.stack.setCurrentIndex(0)

        if self.current_customer_id:
            c = customer_service.get_customer(self.current_customer_id)
            self.filter_label.setText(f"Showing bluebooks for: {c.name}" if c else "")
        else:
            self.filter_label.setText("Search or select a customer to start browsing.")
        if bluebook_id is not None:
            self._refresh_file_count_for_bluebook(bluebook_id)

        while self.stack.count() > 1:
            w = self.stack.widget(1)
            self.stack.removeWidget(w)
            w.deleteLater()

    def _refresh_file_count_for_bluebook(self, bluebook_id):
        from dal import dal

        file_counts = dal.get_file_counts_batch([bluebook_id])
        for row in range(self.bluebook_table.rowCount()):
            item = self.bluebook_table.item(row, 0)
            if item and item.data(Qt.UserRole) == bluebook_id:
                self.bluebook_table.setItem(
                    row,
                    4,
                    QTableWidgetItem(str(file_counts.get(bluebook_id, 0))),
                )
                break

    def _theme_label(self) -> str:
        if not self._theme_manager:
            return ""
        icons = {"Moon", "Wave", "Ember", "Arctic"}
        name = self._theme_manager.current
        label = self._theme_manager.label(name)
        return f"{label}"

    def _cycle_theme(self):
        if not self._theme_manager:
            return
        self._theme_manager.next_theme()
        if hasattr(self, "_theme_btn"):
            self._theme_btn.setText(self._theme_label())
        self.statusBar().showMessage(
            f"Theme: {self._theme_manager.label(self._theme_manager.current)}", 3000
        )

    def _open_qa_window(self):
        if self._qa_window is None or not self._qa_window.isVisible():
            self._qa_window = QAWindow(parent=None)
            self._qa_window.open_bluebook_requested.connect(self._show_detail)

        self._qa_window.activate()

    def _toggle_console(self):
        import sys

        if sys.platform != "win32":
            return

        import ctypes

        kernel32 = ctypes.windll.kernel32
        hwnd = kernel32.GetConsoleWindow()
        if hwnd:
            user32 = ctypes.windll.user32
            if user32.IsWindowVisible(hwnd):
                user32.ShowWindow(hwnd, 0)
            else:
                user32.ShowWindow(hwnd, 5)
        else:
            kernel32.AllocConsole()
            sys.stdout = open("CONOUT$", "w", encoding="utf-8")
            sys.stderr = open("CONOUT$", "w", encoding="utf-8")

    def _create_bluebook(self):
        die, ok = QInputDialog.getText(self, "New Bluebook", "Die Number:")
        if not ok or not die.strip():
            return

        desc, ok2 = QInputDialog.getText(self, "New Bluebook", "Description (optional):")
        desc = desc.strip() if ok2 else ""

        try:
            bb = bluebook_service.create_bluebook(die.strip(), desc)
            if self.current_customer_id:
                customer_service.link_bluebook(self.current_customer_id, bb.id)

            import subprocess
            import sys

            script_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "..",
                "scripts",
                "auto_link_files.py",
            )
            script_path = os.path.normpath(script_path)
            flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            python_exe = "python" if getattr(sys, "frozen", False) else sys.executable

            try:
                subprocess.Popen([python_exe, script_path, bb.die_number], creationflags=flags)
            except Exception as e:
                print(f"Subprocess launch failed: {e}")

            self._load_bluebooks()
            self.statusBar().showMessage(
                f"Created Bluebook Die# {bb.die_number}. Auto-linking files in background..."
            )
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))

    def _delete_bluebook(self):
        selected_rows = {item.row() for item in self.bluebook_table.selectedItems()}
        if not selected_rows:
            QMessageBox.information(self, "No Selection", "Please select a bluebook to delete.")
            return

        bb_list = []
        for row in sorted(selected_rows):
            bb_id = self.bluebook_table.item(row, 0).data(Qt.UserRole)
            bb = bluebook_service.get_bluebook(bb_id)
            if bb:
                bb_list.append(bb)

        if not bb_list:
            return

        from services.security import require_password

        if not require_password("Deleting bluebook(s)", self):
            return

        die_nums = ", ".join(bb.die_number for bb in bb_list)
        if len(bb_list) == 1:
            msg = (
                f"Delete Bluebook Die# {die_nums}?\n\n"
                "This will remove the database record and all associated files."
            )
        else:
            msg = (
                f"Delete {len(bb_list)} bluebooks?\n\n"
                f"Die#: {die_nums}\n\n"
                "This will remove all database records and associated files."
            )

        reply = QMessageBox.question(self, "Delete Bluebook", msg, QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            for bb in bb_list:
                bluebook_service.delete_bluebook(bb.id, delete_files=True)
            self._load_bluebooks()
            self.statusBar().showMessage(f"Deleted {len(bb_list)} bluebook(s): {die_nums}")

    def _on_table_context_menu(self, pos):
        item = self.bluebook_table.itemAt(pos)
        if not item:
            return

        row = item.row()
        bb_id = self.bluebook_table.item(row, 0).data(Qt.UserRole)
        bb = bluebook_service.get_bluebook(bb_id)
        if not bb:
            return

        menu = QMenu(self)

        add_action = menu.addAction("Add Customer to this Bluebook")
        remove_action = None
        if bb.customer_names:
            remove_action = menu.addAction("Remove Customer from this Bluebook")

        menu.addSeparator()

        add_outsource_action = menu.addAction("Add Outsource to this Bluebook")
        remove_outsource_action = None
        if bb.outsource_names:
            remove_outsource_action = menu.addAction("Remove Outsource from this Bluebook")

        menu.addSeparator()
        open_action = menu.addAction("Open Bluebook")
        desc_action = menu.addAction("Change Description")
        menu.addSeparator()
        delete_action = menu.addAction("Delete Bluebook")

        action = menu.exec(self.bluebook_table.viewport().mapToGlobal(pos))
        if not action:
            return

        if action == add_action:
            self._add_customer_to_bluebook(bb)
        elif action == remove_action:
            self._remove_customer_from_bluebook(bb)
        elif action == add_outsource_action:
            self._add_outsource_to_bluebook(bb)
        elif action == remove_outsource_action:
            self._remove_outsource_from_bluebook(bb)
        elif action == open_action:
            self._show_detail(bb.id)
        elif action == desc_action:
            self._change_description(bb)
        elif action == delete_action:
            self._delete_bluebook()

    def _add_customer_to_bluebook(self, bb):
        all_customers = customer_service.list_customers()
        current_customers = customer_service.get_customers_for_bluebook(bb.id)
        current_ids = {c.id for c in current_customers}
        available = [c for c in all_customers if c.id not in current_ids]

        if not available:
            QMessageBox.information(
                self,
                "No Available Customers",
                "All customers are already linked to this bluebook.\nYou can add a new customer from the sidebar first.",
            )
            return

        names = [c.name for c in available]
        name, ok = QInputDialog.getItem(
            self,
            "Add Customer",
            f"Select customer to add to Die# {bb.die_number}:",
            names,
            0,
            False,
        )

        if ok and name:
            customer = next(c for c in available if c.name == name)
            customer_service.link_bluebook(customer.id, bb.id)
            self._load_bluebooks()
            self.statusBar().showMessage(f"Added '{name}' to Die# {bb.die_number}")

    def _remove_customer_from_bluebook(self, bb):
        current_customers = customer_service.get_customers_for_bluebook(bb.id)
        if not current_customers:
            return

        names = [c.name for c in current_customers]
        name, ok = QInputDialog.getItem(
            self,
            "Remove Customer",
            f"Select customer to remove from Die# {bb.die_number}:",
            names,
            0,
            False,
        )

        if ok and name:
            customer = next(c for c in current_customers if c.name == name)
            reply = QMessageBox.question(
                self,
                "Confirm Remove",
                f"Remove '{name}' from Die# {bb.die_number}?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                customer_service.unlink_bluebook(customer.id, bb.id)
                self._load_bluebooks()
                self.statusBar().showMessage(f"Removed '{name}' from Die# {bb.die_number}")

    def _change_description(self, bb):
        new_desc, ok = QInputDialog.getText(
            self,
            "Change Description",
            f"New description for Die# {bb.die_number}:",
            text=bb.description or "",
        )
        if ok:
            bluebook_service.update_bluebook(bb.id, bb.die_number, new_desc.strip())
            self._load_bluebooks()
            self.statusBar().showMessage(f"Updated description for Die# {bb.die_number}")

    def _add_outsource_to_bluebook(self, bb):
        all_outsources = outsource_service.list_outsources()
        current_outsources = outsource_service.get_outsources_for_bluebook(bb.id)
        current_ids = {o.id for o in current_outsources}
        available = [o for o in all_outsources if o.id not in current_ids]

        if not available:
            name, ok = QInputDialog.getText(
                self,
                "New Outsource",
                "No available outsources. Enter a name to create one:",
            )
            if ok and name.strip():
                try:
                    new_o = outsource_service.create_outsource(name.strip())
                    outsource_service.link_bluebook(new_o.id, bb.id)
                    self._load_bluebooks()
                    self.statusBar().showMessage(f"Created and added '{new_o.name}' to Die# {bb.die_number}")
                except Exception as e:
                    QMessageBox.warning(self, "Error", str(e))
            return

        names = [o.name for o in available] + ["+ Create New..."]
        name, ok = QInputDialog.getItem(
            self,
            "Add Outsource",
            f"Select outsource to add to Die# {bb.die_number}:",
            names,
            0,
            False,
        )

        if ok and name:
            if name == "+ Create New...":
                new_name, ok2 = QInputDialog.getText(self, "New Outsource", "Outsource name:")
                if ok2 and new_name.strip():
                    try:
                        new_o = outsource_service.create_outsource(new_name.strip())
                        outsource_service.link_bluebook(new_o.id, bb.id)
                        self._load_bluebooks()
                        self.statusBar().showMessage(f"Created and added '{new_o.name}' to Die# {bb.die_number}")
                    except Exception as e:
                        QMessageBox.warning(self, "Error", str(e))
            else:
                outsource = next(o for o in available if o.name == name)
                outsource_service.link_bluebook(outsource.id, bb.id)
                self._load_bluebooks()
                self.statusBar().showMessage(f"Added '{name}' to Die# {bb.die_number}")

    def _remove_outsource_from_bluebook(self, bb):
        current_outsources = outsource_service.get_outsources_for_bluebook(bb.id)
        if not current_outsources:
            return

        names = [o.name for o in current_outsources]
        name, ok = QInputDialog.getItem(
            self,
            "Remove Outsource",
            f"Select outsource to remove from Die# {bb.die_number}:",
            names,
            0,
            False,
        )

        if ok and name:
            outsource = next(o for o in current_outsources if o.name == name)
            reply = QMessageBox.question(
                self,
                "Confirm Remove",
                f"Remove '{name}' from Die# {bb.die_number}?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                outsource_service.unlink_bluebook(outsource.id, bb.id)
                self._load_bluebooks()
                self.statusBar().showMessage(f"Removed '{name}' from Die# {bb.die_number}")

    def _on_bluebooks_dropped(self, customer_id, bb_ids):
        c = customer_service.get_customer(customer_id)
        if not c:
            return

        linked = 0
        for bb_id in bb_ids:
            try:
                customer_service.link_bluebook(customer_id, bb_id)
                linked += 1
            except Exception:
                pass

        self._load_bluebooks()
        self.statusBar().showMessage(f"Linked {linked} bluebook(s) to '{c.name}'")
