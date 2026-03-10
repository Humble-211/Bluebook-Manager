"""
Share to Other Bluebooks Dialog.
"""

from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)
from PySide6.QtCore import Qt

from dal import dal


class ShareDialog(QDialog):
    """Dialog to select target bluebooks for sharing a file.

    Search behaviour:
    - Default: search by die number.
    - If search starts with 'cust', switch to customer mode — show customers
      and selecting a customer shares to ALL of their bluebooks.
    """

    def __init__(self, file_id: int, source_bluebook_id: int, parent=None):
        super().__init__(parent)
        self.file_id = file_id
        self.source_bluebook_id = source_bluebook_id
        self.selected_bluebook_ids: list[int] = []
        self._customer_mode = False

        self.setWindowTitle("Share to Other Bluebooks")
        self.setMinimumSize(450, 500)
        self._build_ui()
        self._load_bluebooks()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("Select Bluebooks to share this file to:")
        header.setObjectName("sectionHeader")
        layout.addWidget(header)

        # Search
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(
            '🔍 Search by Die Number... (type "cust" to search by Customer)')
        self.search_input.textChanged.connect(self._filter_list)
        layout.addWidget(self.search_input)

        # Mode hint label
        self.mode_label = QLabel("")
        self.mode_label.setStyleSheet("color: #94e2d5; font-size: 12px;")
        self.mode_label.hide()
        layout.addWidget(self.mode_label)

        # Bluebook list with checkboxes
        self.bluebook_list = QListWidget()
        layout.addWidget(self.bluebook_list)

        # Already shared info
        self.info_label = QLabel("")
        self.info_label.setStyleSheet("color: #f9e2af; font-style: italic;")
        layout.addWidget(self.info_label)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_share = QPushButton("Share")
        btn_share.setObjectName("primaryButton")
        btn_share.clicked.connect(self._on_share)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_share)
        layout.addLayout(btn_layout)

    def _load_bluebooks(self):
        self.all_bluebooks = dal.list_bluebooks()
        # Get already-shared targets
        already_shared = dal.get_shared_targets(self.file_id)
        self.already_shared_ids = {b.id for b in already_shared}

        self._populate_list(self.all_bluebooks)

        if self.already_shared_ids:
            self.info_label.setText(
                f"Already shared to {len(self.already_shared_ids)} bluebook(s)")

    def _populate_list(self, bluebooks):
        self.bluebook_list.clear()
        for bb in bluebooks:
            if bb.id == self.source_bluebook_id:
                continue  # Skip the source bluebook
            item = QListWidgetItem()
            customers_str = ", ".join(bb.customer_names) if bb.customer_names else "No customer"
            checkbox = QCheckBox(f"Die# {bb.die_number}  —  {customers_str}")
            checkbox.setProperty("bluebook_id", bb.id)

            if bb.id in self.already_shared_ids:
                checkbox.setChecked(True)
                checkbox.setEnabled(False)
                checkbox.setText(checkbox.text() + "  [already shared]")

            self.bluebook_list.addItem(item)
            self.bluebook_list.setItemWidget(item, checkbox)

    def _populate_customer_list(self, customers):
        """Populate the list with customers (each resolves to all their bluebooks)."""
        self.bluebook_list.clear()
        for cust in customers:
            bbs = dal.get_bluebooks_for_customer(cust.id)
            # Filter out source bluebook
            target_ids = [bb.id for bb in bbs if bb.id != self.source_bluebook_id]
            if not target_ids:
                continue

            die_numbers = ", ".join(bb.die_number for bb in bbs
                                     if bb.id != self.source_bluebook_id)
            item = QListWidgetItem()
            checkbox = QCheckBox(
                f"👤 {cust.name}  —  {len(target_ids)} bluebook(s): {die_numbers}")
            checkbox.setProperty("bluebook_ids", target_ids)

            self.bluebook_list.addItem(item)
            self.bluebook_list.setItemWidget(item, checkbox)

    def _filter_list(self, text):
        text_lower = text.strip().lower()

        if text_lower.startswith("cust"):
            # Customer search mode
            self._customer_mode = True
            self.mode_label.setText("📋 Customer mode — selecting a customer shares to all their bluebooks")
            self.mode_label.show()

            # Extract the search term after "cust" (strip optional colon/space)
            search_part = text_lower[4:].lstrip(": ")
            customers = dal.list_customers()
            if search_part:
                customers = [c for c in customers
                             if search_part in c.name.lower()]
            self._populate_customer_list(customers)
        else:
            # Normal die number search mode
            self._customer_mode = False
            self.mode_label.hide()
            filtered = [bb for bb in self.all_bluebooks
                        if text_lower in bb.die_number.lower()]
            self._populate_list(filtered)

    def _on_share(self):
        self.selected_bluebook_ids.clear()

        for i in range(self.bluebook_list.count()):
            item = self.bluebook_list.item(i)
            widget = self.bluebook_list.itemWidget(item)
            if not isinstance(widget, QCheckBox) or not widget.isChecked():
                continue
            if not widget.isEnabled():
                continue

            if self._customer_mode:
                # Customer mode: each checkbox holds a list of bluebook IDs
                bb_ids = widget.property("bluebook_ids")
                if bb_ids:
                    for bid in bb_ids:
                        if bid not in self.selected_bluebook_ids \
                                and bid not in self.already_shared_ids:
                            self.selected_bluebook_ids.append(bid)
            else:
                # Normal mode: each checkbox holds a single bluebook ID
                bid = widget.property("bluebook_id")
                self.selected_bluebook_ids.append(bid)

        if not self.selected_bluebook_ids:
            QMessageBox.information(self, "No Selection",
                                    "Please select at least one bluebook.")
            return

        self.accept()
