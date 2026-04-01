"""
Bluebook Manager — Customer Management Panel.
"""

import json

from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtCore import Qt, Signal

from services import customer_service


class CustomerPanel(QWidget):
    """Panel that shows customers and lets user manage them."""

    customer_selected = Signal(int)    # emits customer_id
    show_all = Signal()                # emits when "All" is selected
    bluebooks_dropped = Signal(int, list)  # emits (customer_id, [bluebook_ids])

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        header = QLabel("Customers")
        header.setObjectName("sectionHeader")
        layout.addWidget(header)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Filter customers...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._filter_customers)
        layout.addWidget(self.search_input)

        self.customer_list = QListWidget()
        self.customer_list.currentRowChanged.connect(self._on_selection_changed)

        # Enable drop on customer list
        self.customer_list.setAcceptDrops(True)
        self.customer_list.setDragDropMode(QListWidget.DropOnly)
        self.customer_list.dragEnterEvent = self._drag_enter
        self.customer_list.dragMoveEvent = self._drag_move
        self.customer_list.dragLeaveEvent = self._drag_leave
        self.customer_list.dropEvent = self._drop_event

        layout.addWidget(self.customer_list)

        self._all_customers = []  # cached for filtering

        # Buttons
        btn_layout = QHBoxLayout()
        btn_add = QPushButton("Add")
        btn_add.setObjectName("successButton")
        btn_add.setFixedWidth(70)
        btn_add.clicked.connect(self._add_customer)

        btn_edit = QPushButton("Edit")
        btn_edit.setFixedWidth(70)
        btn_edit.clicked.connect(self._edit_customer)

        btn_del = QPushButton("Delete")
        btn_del.setObjectName("dangerButton")
        btn_del.setFixedWidth(70)
        btn_del.clicked.connect(self._delete_customer)

        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addWidget(btn_del)
        layout.addLayout(btn_layout)

    def refresh(self):
        """Reload customer list."""
        self._all_customers = customer_service.list_customers()
        self._filter_customers(self.search_input.text())

    def _populate_list(self, customers):
        """Fill the list widget with given customers."""
        self.customer_list.blockSignals(True)
        self.customer_list.clear()

        # "All Customers" entry
        all_item = QListWidgetItem("📋  All Customers")
        all_item.setData(Qt.UserRole, None)
        self.customer_list.addItem(all_item)

        for c in customers:
            item = QListWidgetItem(f"{c.name}")
            item.setData(Qt.UserRole, c.id)
            self.customer_list.addItem(item)

        self.customer_list.blockSignals(False)
        self.customer_list.setCurrentRow(0)

    def _filter_customers(self, text):
        """Filter customer list based on search text."""
        text = text.strip().lower()
        if not text:
            self._populate_list(self._all_customers)
        else:
            filtered = [c for c in self._all_customers if text in c.name.lower()]
            self._populate_list(filtered)

    def set_selected_customer(self, customer_id: int | None):
        """Restore the current sidebar selection after a refresh."""
        self.customer_list.blockSignals(True)
        try:
            for row in range(self.customer_list.count()):
                item = self.customer_list.item(row)
                if item.data(Qt.UserRole) == customer_id:
                    self.customer_list.setCurrentRow(row)
                    break
            else:
                self.customer_list.setCurrentRow(0)
        finally:
            self.customer_list.blockSignals(False)

    def _on_selection_changed(self, row):
        if row < 0:
            return
        item = self.customer_list.item(row)
        cid = item.data(Qt.UserRole)
        if cid is None:
            self.show_all.emit()
        else:
            self.customer_selected.emit(cid)

    def _add_customer(self):
        name, ok = QInputDialog.getText(self, "Add Customer", "Customer name:")
        if ok and name.strip():
            try:
                customer_service.create_customer(name.strip())
                self.refresh()
            except Exception as e:
                QMessageBox.warning(self, "Error", str(e))

    def _edit_customer(self):
        item = self.customer_list.currentItem()
        if not item:
            return
        cid = item.data(Qt.UserRole)
        if cid is None:
            return

        c = customer_service.get_customer(cid)
        if not c:
            return

        name, ok = QInputDialog.getText(
            self, "Edit Customer", "Customer name:", text=c.name)
        if ok and name.strip():
            try:
                customer_service.update_customer(cid, name.strip(), c.contact_info)
                self.refresh()
            except Exception as e:
                QMessageBox.warning(self, "Error", str(e))

    def _delete_customer(self):
        item = self.customer_list.currentItem()
        if not item:
            return
        cid = item.data(Qt.UserRole)
        if cid is None:
            return

        c = customer_service.get_customer(cid)
        if not c:
            return

        # Require password for deletion
        from services.security import require_password
        if not require_password("Deleting a customer", self):
            return

        reply = QMessageBox.question(
            self, "Delete Customer",
            f"Delete customer '{c.name}'?\n"
            f"(Bluebook links will be removed, but bluebooks will remain.)",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            customer_service.delete_customer(cid)
            self.refresh()

    # ── Drag-and-drop handlers ──

    _drag_active = False

    _DRAG_HOVER_STYLE = (
        "QListWidget::item:hover:!selected {"
        "  background-color: #a6e3a1;"
        "  color: #1e1e2e;"
        "  font-weight: bold;"
        "}"
    )

    def _drag_enter(self, event):
        if event.mimeData().hasFormat("application/x-bluebook-ids"):
            self._drag_active = True
            self.customer_list.setStyleSheet(self._DRAG_HOVER_STYLE)
            event.acceptProposedAction()
        else:
            event.ignore()

    def _drag_move(self, event):
        if not event.mimeData().hasFormat("application/x-bluebook-ids"):
            event.ignore()
            return

        item = self.customer_list.itemAt(event.position().toPoint())
        if item and item.data(Qt.UserRole) is not None:
            event.acceptProposedAction()
        else:
            event.ignore()

    def _drag_leave(self, event):
        self._drag_active = False
        self.customer_list.setStyleSheet("")

    def _drop_event(self, event):
        self._drag_active = False
        self.customer_list.setStyleSheet("")

        if not event.mimeData().hasFormat("application/x-bluebook-ids"):
            event.ignore()
            return

        item = self.customer_list.itemAt(event.position().toPoint())
        if not item:
            event.ignore()
            return

        customer_id = item.data(Qt.UserRole)
        if customer_id is None:
            event.ignore()
            return

        data = bytes(event.mimeData().data("application/x-bluebook-ids")).decode()
        bb_ids = json.loads(data)

        self.bluebooks_dropped.emit(customer_id, bb_ids)
        event.acceptProposedAction()
