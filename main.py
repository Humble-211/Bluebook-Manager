"""
Bluebook Manager — Application Entry Point.

Initializes the database and launches the main window.
"""

import sys
import os

# Ensure the project root is on the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication

from config import BUNDLE_DIR
from dal.database import init_db
from services.log_service import setup_logging


def main():
    """Application entry point."""
    # Initialize logging
    setup_logging()

    # Initialize database
    init_db()

    # Create Qt application
    app = QApplication(sys.argv)
    app.setApplicationName("Bluebook Manager")

    # Initialize theme manager and restore last-used theme
    from services.theme_manager import ThemeManager
    theme_manager = ThemeManager(app)
    theme_manager.load_saved()

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
    window = MainWindow(theme_manager=theme_manager)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
