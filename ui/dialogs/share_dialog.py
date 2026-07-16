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
from PySide6.QtCore import Qt, QTimer

from dal import dal


class ShareDialog(QDialog):
    """Dialog to select target bluebooks for sharing a file.

    Search behaviour:
    - Default: search by die number (DB query fires on each keystroke, debounced).
      List starts empty — type to search.
    - If search starts with 'cust', switch to customer mode — show customers
      and selecting a customer shares to ALL of their bluebooks.
    """

    def __init__(self, file_id: int, source_bluebook_id: int, parent=None):
        super().__init__(parent)
        self.file_id = file_id
        self.source_bluebook_id = source_bluebook_id
        self.selected_bluebook_ids: list[int] = []
        self._customer_mode = False

        # Get already-shared targets once on open (cheap query)
        already_shared = dal.get_shared_targets(self.file_id)
        self.already_shared_ids = {b.id for b in already_shared}

        self.setWindowTitle("Share to Other Bluebooks")
        self.setMinimumSize(450, 500)
        self._build_ui()

        # Debounce timer — fires DB query 300 ms after last keystroke
        self._search_timer = QTimer()
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._do_search)

        if self.already_shared_ids:
            self.info_label.setText(
                f"Already shared to {len(self.already_shared_ids)} bluebook(s)")

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
        self.search_input.textChanged.connect(self._on_search_changed)
        layout.addWidget(self.search_input)

        # Mode hint label
        self.mode_label = QLabel("")
        self.mode_label.setStyleSheet("color: #94e2d5; font-size: 12px;")
        self.mode_label.hide()
        layout.addWidget(self.mode_label)

        # Bluebook list with checkboxes
        self.bluebook_list = QListWidget()
        layout.addWidget(self.bluebook_list)

        # Prompt shown when list is empty
        self.prompt_label = QLabel("Type a die number above to search bluebooks.")
        self.prompt_label.setAlignment(Qt.AlignCenter)
        self.prompt_label.setStyleSheet("color: #6c7086; font-style: italic; padding: 12px;")
        layout.addWidget(self.prompt_label)

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

    # ── Search / filter ──────────────────────────────────────────────────────

    def _on_search_changed(self, text):
        """Kick off debounce timer on each keystroke."""
        text_lower = text.strip().lower()
        if text_lower.startswith("cust"):
            self._customer_mode = True
            self.mode_label.setText(
                "📋 Customer mode — selecting a customer shares to all their bluebooks")
            self.mode_label.show()
        else:
            self._customer_mode = False
            self.mode_label.hide()

        self._search_timer.start(300)

    def _do_search(self):
        """Run the actual DB query after debounce settles."""
        text_lower = self.search_input.text().strip().lower()

        if not text_lower:
            # Empty input → clear list, show prompt
            self.bluebook_list.clear()
            self.prompt_label.show()
            return

        self.prompt_label.hide()

        if text_lower.startswith("cust"):
            search_part = text_lower[4:].lstrip(": ")
            customers_info = dal.get_customer_sharing_info(search_part)
            self._populate_customer_list(customers_info)
        else:
            # Query DB with the typed die number fragment
            bluebooks = dal.list_bluebooks(search=text_lower, limit=200)
            self._populate_list(bluebooks)

    # ── Population helpers ───────────────────────────────────────────────────

    def _populate_list(self, bluebooks):
        self.bluebook_list.clear()
        for bb in bluebooks:
            if bb.id == self.source_bluebook_id:
                continue  # Skip the source bluebook
            customers_str = ", ".join(bb.customer_names) if bb.customer_names else "No customer"
            text = f"Die# {bb.die_number}  —  {customers_str}"
            
            is_already_shared = bb.id in self.already_shared_ids
            if is_already_shared:
                text += "  [already shared]"
                
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, bb.id)
            
            if is_already_shared:
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                item.setCheckState(Qt.Checked)
            else:
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                
            self.bluebook_list.addItem(item)

    def _populate_customer_list(self, customers_info):
        """Populate the list with customers (each resolves to all their bluebooks)."""
        self.bluebook_list.clear()
        for cust in customers_info:
            # Filter out source bluebook
            target_ids = [bb["id"] for bb in cust["bluebooks"] if bb["id"] != self.source_bluebook_id]
            if not target_ids:
                continue

            die_numbers = ", ".join(bb["die_number"] for bb in cust["bluebooks"]
                                     if bb["id"] != self.source_bluebook_id)
            
            text = f"👤 {cust['name']}  —  {len(target_ids)} bluebook(s): {die_numbers}"
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, target_ids)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)

            self.bluebook_list.addItem(item)

    # ── Share action ─────────────────────────────────────────────────────────

    def _on_share(self):
        self.selected_bluebook_ids.clear()

        for i in range(self.bluebook_list.count()):
            item = self.bluebook_list.item(i)
            if item.checkState() != Qt.Checked:
                continue
            if not (item.flags() & Qt.ItemIsEnabled):
                continue

            if self._customer_mode:
                # Customer mode: each item holds a list of bluebook IDs in UserRole
                bb_ids = item.data(Qt.UserRole)
                if bb_ids:
                    for bid in bb_ids:
                        if bid not in self.selected_bluebook_ids \
                                and bid not in self.already_shared_ids:
                            self.selected_bluebook_ids.append(bid)
            else:
                # Normal mode: each item holds a single bluebook ID in UserRole
                bid = item.data(Qt.UserRole)
                if bid is not None:
                    self.selected_bluebook_ids.append(bid)

        if not self.selected_bluebook_ids:
            QMessageBox.information(self, "No Selection",
                                    "Please select at least one bluebook.")
            return

        self.accept()

