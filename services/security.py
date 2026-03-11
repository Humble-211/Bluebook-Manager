"""
Bluebook Manager — Security utilities.

Provides password verification for destructive operations.
"""

import hashlib
import os

from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QVBoxLayout

from config import BASE_DIR

# Path to the password hash file
_PASSWORD_FILE = os.path.join(BASE_DIR, "data", ".security")

# Default password (used on first run)
_DEFAULT_PASSWORD = "admin123"


def _hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _get_stored_hash() -> str:
    """Read the stored password hash, or create default."""
    os.makedirs(os.path.dirname(_PASSWORD_FILE), exist_ok=True)
    if not os.path.isfile(_PASSWORD_FILE):
        # First run — store default password hash
        h = _hash_password(_DEFAULT_PASSWORD)
        with open(_PASSWORD_FILE, "w") as f:
            f.write(h)
        return h
    with open(_PASSWORD_FILE, "r") as f:
        return f.read().strip()


def verify_password(password: str) -> bool:
    """Check if the given password matches the stored hash."""
    return _hash_password(password) == _get_stored_hash()


def change_password(old_password: str, new_password: str) -> bool:
    """Change the password if old_password is correct."""
    if not verify_password(old_password):
        return False
    h = _hash_password(new_password)
    with open(_PASSWORD_FILE, "w") as f:
        f.write(h)
    return True


class PasswordDialog(QDialog):
    """Dialog that prompts for a password before a destructive action."""

    def __init__(self, action_description: str = "This action", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Security Verification")
        self.setFixedWidth(380)
        self.verified = False
        self._build_ui(action_description)

    def _build_ui(self, action_description):
        layout = QVBoxLayout(self)

        label = QLabel(f"Password Required:")
        label.setWordWrap(True)
        layout.addWidget(label)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("Enter password...")
        self.password_input.returnPressed.connect(self._on_verify)
        layout.addWidget(self.password_input)

        btn_layout = QHBoxLayout()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("Verify")
        btn_ok.setObjectName("primaryButton")
        btn_ok.clicked.connect(self._on_verify)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_ok)
        layout.addLayout(btn_layout)

    def _on_verify(self):
        if verify_password(self.password_input.text()):
            self.verified = True
            self.accept()
        else:
            QMessageBox.warning(self, "Access Denied", "Incorrect password.")
            self.password_input.clear()
            self.password_input.setFocus()


def require_password(action_description: str, parent=None) -> bool:
    """Show password dialog and return True if verified."""
    dlg = PasswordDialog(action_description, parent)
    dlg.exec()
    return dlg.verified
