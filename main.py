"""
Bluebook Manager — Application Entry Point.

Initializes the database and launches the main window.
"""

import sys
import os

# Ensure the project root is on the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication, QSplashScreen
from PySide6.QtGui import QPixmap, QPainter, QColor, QFont
from PySide6.QtCore import Qt

from config import BUNDLE_DIR
from dal.database import init_db
from services.log_service import setup_logging


def _make_splash_pixmap() -> QPixmap:
    """Render a branded splash screen using QPainter — no external asset needed."""
    w, h = 520, 240
    pix = QPixmap(w, h)
    pix.fill(QColor("#1e1e2e"))

    painter = QPainter(pix)
    painter.setRenderHint(QPainter.Antialiasing)

    # Background panel
    painter.setBrush(QColor("#313244"))
    painter.setPen(Qt.NoPen)
    painter.drawRoundedRect(20, 20, w - 40, h - 40, 14, 14)

    # Accent bar at top
    painter.setBrush(QColor("#89b4fa"))
    painter.drawRoundedRect(20, 20, w - 40, 4, 2, 2)

    # App title
    font = QFont("Segoe UI", 26, QFont.Bold)
    painter.setFont(font)
    painter.setPen(QColor("#cdd6f4"))
    painter.drawText(pix.rect().adjusted(0, 30, 0, 0), Qt.AlignHCenter | Qt.AlignVCenter, "Bluebook Manager")

    # Subtitle
    font.setPointSize(11)
    font.setBold(False)
    painter.setFont(font)
    painter.setPen(QColor("#6c7086"))
    painter.drawText(pix.rect().adjusted(0, 88, 0, 0), Qt.AlignHCenter | Qt.AlignVCenter, "Starting up...")

    # Bottom progress track
    track_y = h - 42
    painter.setBrush(QColor("#45475a"))
    painter.setPen(Qt.NoPen)
    painter.drawRoundedRect(60, track_y, w - 120, 4, 2, 2)
    painter.setBrush(QColor("#89b4fa"))
    painter.drawRoundedRect(60, track_y, (w - 120) // 2, 4, 2, 2)

    painter.end()
    return pix


def main():
    """Application entry point."""
    # Initialize logging
    setup_logging()

    # Initialize database
    init_db()

    # Create Qt application
    app = QApplication(sys.argv)
    app.setApplicationName("Bluebook Manager")

    # Show splash immediately — user sees feedback within ~1 s instead of a blank void
    splash = QSplashScreen(_make_splash_pixmap(), Qt.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()

    # Initialize theme manager and restore last-used theme
    from services.theme_manager import ThemeManager
    theme_manager = ThemeManager(app)
    theme_manager.load_saved()

    # Clean up both background thread pools and DOCX preview cache on exit.
    # Both pools MUST be shut down — failing to stop _PreviewRenderPool causes
    # "QThread: Destroyed while thread is still running" on exit.
    def _cleanup_on_exit():
        from ui.bluebook_detail import shutdown_word_pool, shutdown_preview_render_pool
        shutdown_preview_render_pool()
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

    # Dismiss splash once the main window is painted
    splash.finish(window)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
