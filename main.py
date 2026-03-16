"""
Bluebook Manager — Application Entry Point.

Initializes the database and launches the main window.
"""

import sys
import os

# Ensure the project root is on the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QFile, QTextStream

from config import BUNDLE_DIR
from dal.database import init_db
from services.log_service import setup_logging


def load_stylesheet(app: QApplication):
    """Load the QSS stylesheet."""
    qss_path = os.path.join(BUNDLE_DIR, "ui", "resources", "styles.qss")
    if os.path.isfile(qss_path):
        with open(qss_path, "r") as f:
            app.setStyleSheet(f.read())


def main():
    """Application entry point."""
    # Initialize logging
    setup_logging()

    # Initialize database
    init_db()

    # Create Qt application
    app = QApplication(sys.argv)
    app.setApplicationName("Bluebook Manager")

    # Load stylesheet
    load_stylesheet(app)

    # Pre-warm Word COM pool so first DOCX preview is instant
    from ui.bluebook_detail import _get_word_pool
    _get_word_pool()

    # Clean up Word COM pool and DOCX preview cache on exit
    def _cleanup_on_exit():
        from ui.bluebook_detail import shutdown_word_pool
        shutdown_word_pool()
        import shutil
        import tempfile
        cache_dir = os.path.join(tempfile.gettempdir(), "bluebook_docx_cache")
        if os.path.isdir(cache_dir):
            shutil.rmtree(cache_dir, ignore_errors=True)

    app.aboutToQuit.connect(_cleanup_on_exit)

    # Launch main window
    from ui.main_window import MainWindow
    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
