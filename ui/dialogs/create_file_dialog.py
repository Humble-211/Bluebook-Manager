"""
Create File from Template Dialog.
"""

from datetime import datetime

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from config import SECTION_LABELS, TEMPLATE_SECTIONS
from dal import dal


class CreateFileDialog(QDialog):
    """Dialog to create a new file from a template."""

    def __init__(self, default_section: str = "", die_number: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Create New File from Template")
        self.setMinimumWidth(400)
        self.result_section = ""
        self.result_filename = ""
        self._die_number = die_number

        self._build_ui(default_section)

    def _build_ui(self, default_section):
        layout = QVBoxLayout(self)

        # Section selector
        layout.addWidget(QLabel("Section:"))
        self.section_combo = QComboBox()
        for sec in TEMPLATE_SECTIONS:
            self.section_combo.addItem(SECTION_LABELS[sec], sec)
        if default_section in TEMPLATE_SECTIONS:
            idx = TEMPLATE_SECTIONS.index(default_section)
            self.section_combo.setCurrentIndex(idx)
        self.section_combo.currentIndexChanged.connect(self._on_section_changed)
        layout.addWidget(self.section_combo)

        # Filename / Description label and input
        self.filename_label = QLabel("Filename (without extension):")
        layout.addWidget(self.filename_label)
        self.filename_input = QLineEdit()
        self.filename_input.setPlaceholderText("e.g. quality_alert_001")
        layout.addWidget(self.filename_input)

        # QA preview label (shows the generated filename for quality_alerts)
        self.qa_preview = QLabel("")
        self.qa_preview.setStyleSheet("color: #94e2d5; font-size: 12px; padding: 2px 0;")
        self.qa_preview.setWordWrap(True)
        self.qa_preview.hide()
        layout.addWidget(self.qa_preview)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_create = QPushButton("Create")
        btn_create.setObjectName("successButton")
        btn_create.clicked.connect(self._on_create)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_create)
        layout.addLayout(btn_layout)

        # Connect text change once — the slot checks the mode itself
        self.filename_input.textChanged.connect(self._update_preview)

        # Apply initial section state
        self._on_section_changed()

    def _is_quality_alerts(self) -> bool:
        return self.section_combo.currentData() == "quality_alerts"

    def _is_fit_and_functions(self) -> bool:
        return self.section_combo.currentData() == "fit_and_functions"

    def _is_auto_generated(self) -> bool:
        return self._is_quality_alerts() or self._is_fit_and_functions()

    def _on_section_changed(self):
        """Update UI labels based on selected section."""
        if self._is_quality_alerts():
            self.filename_label.setText("Description:")
            self.filename_input.setPlaceholderText("e.g. Screw Hole")
            self.filename_label.show()
            self.filename_input.show()
            self.qa_preview.show()
            self._update_preview()
        elif self._is_fit_and_functions():
            self.filename_label.hide()
            self.filename_input.hide()
            self.qa_preview.show()
            self._update_preview()
        else:
            self.filename_label.setText("Filename (without extension):")
            self.filename_input.setPlaceholderText("e.g. file_001")
            self.filename_label.show()
            self.filename_input.show()
            self.qa_preview.hide()

    def _update_preview(self):
        """Show a preview of the generated QA or FF filename."""
        if not self._is_auto_generated():
            return
        
        year_2d = datetime.now().year % 100
        
        if self._is_quality_alerts():
            desc = self.filename_input.text().strip()
            if not desc:
                self.qa_preview.setText(f"Filename: QA-YY-NNN-DIENUM-<description>.docx")
                return
            desc_part = desc.replace(" ", "-")
            preview = f"QA-{year_2d:02d}-NNN-{self._die_number}-{desc_part}.docx"
            self.qa_preview.setText(f"Filename: {preview}")
        else:
            # Fit and functions
            preview = f"FF-{year_2d:02d}-NNN-{self._die_number}.docx"
            self.qa_preview.setText(f"Filename: {preview}")

    def _on_create(self):
        self.result_section = self.section_combo.currentData()
        raw_input = self.filename_input.text().strip()

        if self._is_quality_alerts():
            if not raw_input:
                QMessageBox.warning(self, "Missing Description", "Please enter a description.")
                return
        elif not self._is_auto_generated():
            if not raw_input:
                QMessageBox.warning(self, "Missing Filename", "Please enter a filename.")
                return

        if self._is_quality_alerts():
            # Auto-generate QA filename: QA-YY-NNN-DIENUM-description
            year_2d = datetime.now().year % 100
            next_num = dal.get_next_qa_number(year_2d)
            desc_part = raw_input.replace(" ", "-")
            # Sanitize description
            invalid_chars = '<>:"/\\|?*'
            for ch in invalid_chars:
                desc_part = desc_part.replace(ch, "_")
            self.result_filename = f"QA-{year_2d:02d}-{next_num:03d}-{self._die_number}-{desc_part}"
        elif self._is_fit_and_functions():
            # Auto-generate FF filename: FF-YY-NNN-DIENUM
            year_2d = datetime.now().year % 100
            next_num = dal.get_next_ff_number(year_2d)
            self.result_filename = f"FF-{year_2d:02d}-{next_num:03d}-{self._die_number}"
        else:
            self.result_filename = raw_input
            # Sanitize filename
            invalid_chars = '<>:"/\\|?*'
            for ch in invalid_chars:
                self.result_filename = self.result_filename.replace(ch, "_")

        self.accept()
