"""
Attach File Dialog.
"""

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from config import SECTION_FILE_TYPES, SECTION_LABELS, SECTION_TYPES


class AttachFileDialog(QDialog):
    """Dialog to select and attach an existing file to a bluebook section."""

    def __init__(self, default_section: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Attach Files")
        self.setMinimumWidth(500)
        self.setMinimumHeight(400)
        self.result_section = ""
        self.result_file_paths: list[str] = []

        self._build_ui(default_section)

    def _build_ui(self, default_section):
        layout = QVBoxLayout(self)

        # Section selector
        layout.addWidget(QLabel("Section:"))
        self.section_combo = QComboBox()
        for sec in SECTION_TYPES:
            self.section_combo.addItem(SECTION_LABELS[sec], sec)
        if default_section in SECTION_TYPES:
            idx = SECTION_TYPES.index(default_section)
            self.section_combo.setCurrentIndex(idx)
        layout.addWidget(self.section_combo)

        # File picker
        file_header = QHBoxLayout()
        file_header.addWidget(QLabel("Files:"))
        file_header.addStretch()
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._browse_files)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_files)
        file_header.addWidget(browse_btn)
        file_header.addWidget(clear_btn)
        layout.addLayout(file_header)

        self.file_list = QListWidget()
        self.file_list.setAlternatingRowColors(True)
        layout.addWidget(self.file_list)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_attach = QPushButton("Attach")
        btn_attach.setObjectName("primaryButton")
        btn_attach.clicked.connect(self._on_attach)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_attach)
        layout.addLayout(btn_layout)

    def _browse_files(self):
        section = self.section_combo.currentData()
        allowed = SECTION_FILE_TYPES.get(section, [])

        # Build filter string
        ext_list = " ".join(f"*{ext}" for ext in allowed)
        filter_str = f"Allowed Files ({ext_list});;All Files (*.*)"

        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", "", filter_str)
        if paths:
            import os
            for p in paths:
                # Avoid duplicates
                existing = [self.file_list.item(i).toolTip()
                            for i in range(self.file_list.count())]
                if p not in existing:
                    item = self.file_list.addItem(os.path.basename(p))
                    self.file_list.item(self.file_list.count() - 1).setToolTip(p)

    def _clear_files(self):
        self.file_list.clear()

    def _on_attach(self):
        self.result_section = self.section_combo.currentData()
        self.result_file_paths = [
            self.file_list.item(i).toolTip()
            for i in range(self.file_list.count())
        ]

        if not self.result_file_paths:
            QMessageBox.warning(self, "No Files Selected",
                                "Please select at least one file to attach.")
            return

        self.accept()
