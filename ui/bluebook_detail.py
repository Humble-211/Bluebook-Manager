"""
Bluebook Manager — Bluebook Detail Screen.

Shows sections and files for a single bluebook with action buttons.
"""

import hashlib
import os
import tempfile
import threading

from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QInputDialog,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtGui import QPixmap, QTransform
from PySide6.QtCore import QThread, QTimer, Qt, Signal

from config import (
    SECTION_LABELS,
    SECTION_TYPES,
    SHAREABLE_SECTIONS,
    TEMPLATE_SECTIONS,
)
from services import bluebook_service, file_service, print_service, sharing_service
from services.log_service import log
from ui.dialogs.attach_file_dialog import AttachFileDialog
from ui.dialogs.create_file_dialog import CreateFileDialog
from ui.dialogs.share_dialog import ShareDialog


# ── Cache directory for converted DOCX → PDF files ──
_DOCX_CACHE_DIR = os.path.join(tempfile.gettempdir(), "bluebook_docx_cache")
os.makedirs(_DOCX_CACHE_DIR, exist_ok=True)


class _WordCOMPool(QThread):
    """Singleton thread that owns a persistent Word COM instance.

    Keeps Word.Application alive between conversions so subsequent
    DOCX → PDF previews are near-instant instead of waiting for
    Word to start up each time.
    """
    conversion_done = Signal(str, str)   # (request_id, pdf_path)
    conversion_error = Signal(str, str)  # (request_id, error_msg)

    def __init__(self):
        super().__init__()
        self._queue = []       # list of (request_id, docx_path, cache_path)
        self._lock = threading.Lock()
        self._event = threading.Event()
        self._stop = False
        self._word = None

    # ── public API (called from main thread) ──

    def request_conversion(self, request_id: str, docx_path: str, cache_path: str):
        with self._lock:
            self._queue.append((request_id, docx_path, cache_path))
        self._event.set()

    def shutdown(self):
        self._stop = True
        self._event.set()
        self.wait(5000)

    # ── thread body ──

    def run(self):
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            self._word = win32com.client.DispatchEx("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = 0
        except Exception as e:
            # If Word can't start, every request will fail
            self._word = None

        while not self._stop:
            self._event.wait()
            self._event.clear()

            while True:
                with self._lock:
                    if not self._queue:
                        break
                    req_id, docx_path, cache_path = self._queue.pop(0)

                if self._word is None:
                    self.conversion_error.emit(
                        req_id, "Microsoft Word could not be started.")
                    continue

                try:
                    doc = self._word.Documents.Open(
                        os.path.abspath(docx_path), ReadOnly=True)
                    doc.SaveAs(os.path.abspath(cache_path), FileFormat=17)
                    doc.Close(False)
                    self.conversion_done.emit(req_id, cache_path)
                except Exception as e:
                    self.conversion_error.emit(req_id, str(e))
                    # Try to recover: if Word crashed, restart it
                    try:
                        self._word = win32com.client.DispatchEx("Word.Application")
                        self._word.Visible = False
                        self._word.DisplayAlerts = 0
                    except Exception:
                        self._word = None

        # Clean shutdown
        if self._word is not None:
            try:
                self._word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


# Module-level singleton — created lazily on first use
_word_pool: _WordCOMPool | None = None


def _get_word_pool() -> _WordCOMPool:
    global _word_pool
    if _word_pool is None or not _word_pool.isRunning():
        _word_pool = _WordCOMPool()
        _word_pool.start()
    return _word_pool


def shutdown_word_pool():
    """Call on app exit to cleanly close Word."""
    global _word_pool
    if _word_pool is not None:
        _word_pool.shutdown()
        _word_pool = None


class ScaledPixmapLabel(QLabel):
    """QLabel that auto-scales its pixmap to fit available width on resize."""

    def __init__(self, pixmap, parent=None):
        super().__init__(parent)
        self._original_pixmap = pixmap
        self.setAlignment(Qt.AlignCenter)
        self._update_pixmap()

    def _update_pixmap(self):
        if not self._original_pixmap or self._original_pixmap.isNull():
            return
        try:
            w = self.width() - 10
            if w > 0 and self._original_pixmap.width() > w:
                scaled = self._original_pixmap.scaledToWidth(w, Qt.SmoothTransformation)
            else:
                scaled = self._original_pixmap
            super().setPixmap(scaled)
            self.setFixedHeight(scaled.height() + 10)
        except Exception:
            pass

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_pixmap()


class BluebookDetailWidget(QWidget):
    """Detail view for a single bluebook."""

    closed = Signal()  # emitted when user navigates back

    def __init__(self, bluebook_id: int, parent=None):
        super().__init__(parent)
        self.bluebook_id = bluebook_id
        self.bluebook = bluebook_service.get_bluebook(bluebook_id)
        self.current_section = SECTION_TYPES[0]
        self._docx_request_id: str | None = None

        self._build_ui()
        self._connect_docx_preview_signals()
        self._load_sections()
        log("OPEN_BLUEBOOK", f"Die# {self.bluebook.die_number}")

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # ── Header ──
        header_layout = QHBoxLayout()
        back_btn = QPushButton("← Back")
        back_btn.setFixedWidth(80)
        back_btn.clicked.connect(self.closed.emit)
        header_layout.addWidget(back_btn)

        title = QLabel(f"Bluebook: Die# {self.bluebook.die_number}")
        title.setObjectName("headerLabel")
        header_layout.addWidget(title, 1)

        # Print All button in header
        print_all_btn = QPushButton("Print All")
        print_all_btn.setObjectName("primaryButton")
        print_all_btn.setFixedWidth(120)
        print_all_btn.clicked.connect(self._print_all)
        header_layout.addWidget(print_all_btn)

        main_layout.addLayout(header_layout)

        # Description
        if self.bluebook.description:
            desc = QLabel(self.bluebook.description)
            desc.setStyleSheet("color: #a6adc8; padding: 0 0 8px 0;")
            main_layout.addWidget(desc)

        # Customers & Outsource
        info_parts = []
        customers_text = ", ".join(self.bluebook.customer_names) or "No customers linked"
        info_parts.append(f'<span style="color: #94e2d5;">Customer: {customers_text}</span>')
        if self.bluebook.outsource_names:
            outsource_text = ", ".join(self.bluebook.outsource_names)
            info_parts.append(f'<span style="color: #f9e2af;">Outsource: {outsource_text}</span>')
        info_label = QLabel("  |  ".join(info_parts))
        info_label.setTextFormat(Qt.RichText)
        info_label.setStyleSheet("padding: 0 0 8px 0;")
        main_layout.addWidget(info_label)

        # ── Main Splitter: Sections | Files + Preview ──
        main_splitter = QSplitter(Qt.Horizontal)

        # Left: Sections
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        sec_label = QLabel("Sections")
        sec_label.setObjectName("sectionHeader")
        left_layout.addWidget(sec_label)

        self.section_list = QListWidget()
        self.section_list.currentRowChanged.connect(self._on_section_changed)
        left_layout.addWidget(self.section_list)

        main_splitter.addWidget(left_panel)

        # Right: Files + Preview (nested splitter)
        right_splitter = QSplitter(Qt.Horizontal)

        # Right-Left: File list + Buttons
        files_panel = QWidget()
        files_layout = QVBoxLayout(files_panel)
        files_layout.setContentsMargins(0, 0, 0, 0)

        self.files_header = QLabel("Files")
        self.files_header.setObjectName("sectionHeader")
        files_layout.addWidget(self.files_header)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_list.currentItemChanged.connect(self._on_file_selected)
        self.file_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self._on_file_context_menu)
        self.file_list.setAcceptDrops(True)
        self.file_list.setDragDropMode(QListWidget.DropOnly)
        self.file_list.dragEnterEvent = self._drag_enter_event
        self.file_list.dragMoveEvent = self._drag_move_event
        self.file_list.dropEvent = self._drop_event
        files_layout.addWidget(self.file_list)

        right_splitter.addWidget(files_panel)

        # Right-Right: Preview panel
        preview_panel = QWidget()
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(0, 0, 0, 0)

        # Preview header row with rotate button
        preview_header_layout = QHBoxLayout()
        preview_header_layout.setContentsMargins(0, 0, 0, 0)
        preview_header = QLabel("Preview")
        preview_header.setObjectName("sectionHeader")
        preview_header_layout.addWidget(preview_header)
        preview_header_layout.addStretch()

        self.btn_rotate = QPushButton("Rotate")
        self.btn_rotate.setFixedWidth(80)
        self.btn_rotate.clicked.connect(self._rotate_preview)
        preview_header_layout.addWidget(self.btn_rotate)

        preview_layout.addLayout(preview_header_layout)

        self._rotation_angle = 0

        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(True)
        self.preview_scroll.setStyleSheet("QScrollArea { border: 1px solid #45475a; border-radius: 6px; background-color: #313244; }")

        self.preview_label = QLabel("Select a file to preview")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setWordWrap(True)
        self.preview_label.setStyleSheet("color: #6c7086; font-size: 14px; padding: 20px;")
        self.preview_scroll.setWidget(self.preview_label)

        preview_layout.addWidget(self.preview_scroll)

        right_splitter.addWidget(preview_panel)

        # Set 20/80 split for files/preview
        right_splitter.setSizes([200, 800])

        main_splitter.addWidget(right_splitter)

        # Set sections / right split
        main_splitter.setSizes([200, 800])

        main_layout.addWidget(main_splitter, 1)

    def _load_sections(self):
        """Populate the section list with file counts."""
        self.section_list.clear()
        section_counts = file_service.get_section_file_counts(self.bluebook_id)
        for sec in SECTION_TYPES:
            count = section_counts.get(sec, 0)
            count_str = f" ({count})" if count else ""
            item = QListWidgetItem(f"{SECTION_LABELS[sec]}{count_str}")
            item.setData(Qt.UserRole, sec)
            self.section_list.addItem(item)
        self.section_list.setCurrentRow(0)

    def _on_section_changed(self, row):
        if row < 0:
            return
        item = self.section_list.item(row)
        self.current_section = item.data(Qt.UserRole)
        self.files_header.setText(f"Files — {SECTION_LABELS[self.current_section]}")
        self._load_files()

    def _load_files(self):
        """Populate the file list for the current section."""
        self.file_list.clear()
        files = file_service.get_files_for_section(self.bluebook_id, self.current_section)
        shared_original_ids = file_service.get_shared_original_file_ids(
            [bf.id for bf in files if not bf.is_shared]
        )
        for bf in files:
            name = os.path.basename(bf.file_path)
            # Strip .lnk extension for display
            if name.lower().endswith(".lnk"):
                name = name[:-4]
            prefix = ""
            if bf.is_shared:
                prefix = "🔗 "
                if bf.shared_from_die_number:
                    name = f"{prefix}{name}  (from Die# {bf.shared_from_die_number})"
                else:
                    name = f"{prefix}{name}"
            else:
                # Check if this file is shared TO other bluebooks
                if bf.id in shared_original_ids:
                    prefix = "🔗 "
                    name = f"{prefix}{name}  [shared]"

            item = QListWidgetItem(name)
            item.setData(Qt.UserRole, bf)
            if bf.is_shared:
                item.setToolTip(f"Shared from Die# {bf.shared_from_die_number}")
            self.file_list.addItem(item)

        # Clear preview when files reload
        self._clear_preview()

    def _on_file_selected(self, current, previous):
        """Update the preview panel when a file is selected."""
        self._cancel_pending_docx_preview()
        self._rotation_angle = 0
        self.preview_scroll._zoom_factor = 1.0
        if not current:
            self._clear_preview()
            return

        bf = current.data(Qt.UserRole)
        if not bf:
            self._clear_preview()
            return

        abs_path = file_service.get_absolute_path(bf.file_path)
        if not os.path.isfile(abs_path):
            self._show_preview_message("File not found on disk.")
            return

        # Resolve shortcut to its target for preview
        abs_path = file_service.resolve_shortcut(abs_path)
        if not os.path.isfile(abs_path):
            self._show_preview_message(
                "Shortcut target not found.\n\n"
                "The file on the network drive may have been moved or deleted.")
            return

        ext = os.path.splitext(abs_path)[1].lower()

        if ext in (".png", ".jpg", ".jpeg", ".bmp", ".gif"):
            self._preview_image(abs_path)
        elif ext == ".pdf":
            self._preview_pdf(abs_path)
        elif ext == ".docx":
            self._preview_docx(abs_path)
        else:
            self._show_preview_message(f"No preview available for {ext} files.\n\nDouble-click or click 'Open' to view.")

    def _clear_preview(self):
        """Reset preview to default state."""
        self._cancel_pending_docx_preview()
        self._rotation_angle = 0
        self._show_preview_message("Select a file to preview")

    def _connect_docx_preview_signals(self):
        """Connect to the shared Word conversion worker once per widget."""
        pool = _get_word_pool()
        pool.conversion_done.connect(self._on_docx_ready)
        pool.conversion_error.connect(self._on_docx_error)

    def _cancel_pending_docx_preview(self):
        """Invalidate any in-flight DOCX preview for this widget."""
        self._docx_request_id = None

    def _rotate_preview(self):
        """Rotate all preview images by 90 degrees."""
        self._rotation_angle = (self._rotation_angle + 90) % 360
        transform = QTransform().rotate(90)

        widget = self.preview_scroll.widget()
        if not widget:
            return

        # Find all ScaledPixmapLabel widgets in the preview
        labels = widget.findChildren(ScaledPixmapLabel)
        if not labels:
            # Single label as the scroll widget itself
            if isinstance(widget, ScaledPixmapLabel):
                labels = [widget]

        for label in labels:
            if label._original_pixmap and not label._original_pixmap.isNull():
                rotated = label._original_pixmap.transformed(transform, Qt.SmoothTransformation)
                label._original_pixmap = rotated
                label._update_pixmap()

    def _show_preview_message(self, text):
        """Show a text message in the preview area."""
        self.preview_label = QLabel(text)
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setWordWrap(True)
        self.preview_label.setStyleSheet("color: #6c7086; font-size: 14px; padding: 20px;")
        self.preview_scroll.setWidget(self.preview_label)

    def _preview_image(self, abs_path):
        """Show an image preview."""
        pixmap = QPixmap(abs_path)
        if pixmap.isNull():
            self._show_preview_message("Could not load image.")
            return

        label = ScaledPixmapLabel(pixmap)
        label.setStyleSheet("padding: 10px;")
        self.preview_scroll.setWidget(label)

    def _preview_pdf(self, abs_path):
        """Show all pages of a PDF as image previews."""
        try:
            import fitz  # PyMuPDF
            from PySide6.QtGui import QImage

            doc = fitz.open(abs_path)
            if len(doc) == 0:
                self._show_preview_message("PDF has no pages.")
                doc.close()
                return

            container = QWidget()
            container_layout = QVBoxLayout(container)
            container_layout.setContentsMargins(5, 5, 5, 5)
            container_layout.setSpacing(8)

            for page_num in range(len(doc)):
                page = doc[page_num]
                mat = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=mat)
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(img)

                page_label = ScaledPixmapLabel(pixmap)
                page_label.setStyleSheet(
                    "background-color: white; border: 1px solid #45475a; padding: 2px;")
                container_layout.addWidget(page_label)

            doc.close()
            container_layout.addStretch()
            self.preview_scroll.setWidget(container)

        except ImportError:
            self._show_preview_message(
                "PDF preview requires PyMuPDF.\n\n"
                "Install with: pip install PyMuPDF\n\n"
                "Click 'Open' to view the PDF.")
        except Exception as e:
            self._show_preview_message(f"Could not preview PDF:\n{e}")

    def _preview_docx(self, abs_path):
        """Convert DOCX to PDF via persistent Word pool, then render."""
        # Build a cache key from file path + modification time
        try:
            mtime = os.path.getmtime(abs_path)
        except OSError:
            mtime = 0
        cache_key = hashlib.md5(f"{abs_path}|{mtime}".encode()).hexdigest()
        cache_path = os.path.join(_DOCX_CACHE_DIR, f"{cache_key}.pdf")

        if os.path.isfile(cache_path):
            # Cached — render immediately
            self._render_pdf_preview(cache_path)
            return

        # Show loading message while converting in background
        self._show_preview_message("⏳ Loading Preview...")

        # Use a unique request ID so we only handle our own result
        self._docx_request_id = cache_key
        pool = _get_word_pool()
        pool.request_conversion(cache_key, abs_path, cache_path)

    def _on_docx_ready(self, request_id: str, pdf_path: str):
        """Called when background DOCX→PDF conversion finishes."""
        if request_id != getattr(self, "_docx_request_id", None):
            return
        self._docx_request_id = None
        self._render_pdf_preview(pdf_path)

    def _on_docx_error(self, request_id: str, error_msg: str):
        """Called when background DOCX→PDF conversion fails."""
        if request_id != getattr(self, "_docx_request_id", None):
            return
        self._docx_request_id = None
        self._show_preview_message(
            f"Could not convert DOCX to PDF:\n{error_msg}\n\n"
            "Make sure Microsoft Word is installed.")

    def _render_pdf_preview(self, pdf_path: str):
        """Render all pages of a PDF file in the preview panel."""
        try:
            import fitz
            from PySide6.QtGui import QImage

            pdf_doc = fitz.open(pdf_path)
            if len(pdf_doc) == 0:
                self._show_preview_message("Document has no pages.")
                pdf_doc.close()
                return

            container = QWidget()
            container_layout = QVBoxLayout(container)
            container_layout.setContentsMargins(5, 5, 5, 5)
            container_layout.setSpacing(8)

            for page_num in range(len(pdf_doc)):
                page = pdf_doc[page_num]
                mat = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=mat)
                img = QImage(pix.samples, pix.width, pix.height,
                             pix.stride, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(img)

                page_label = ScaledPixmapLabel(pixmap)
                page_label.setStyleSheet(
                    "background-color: white; border: 1px solid #45475a; padding: 2px;")
                container_layout.addWidget(page_label)

            pdf_doc.close()
            container_layout.addStretch()
            self.preview_scroll.setWidget(container)

        except ImportError as e:
            self._show_preview_message(
                f"DOCX preview requires PyMuPDF and pywin32.\n\n{e}")
        except Exception as e:
            self._show_preview_message(f"Could not preview DOCX:\n{e}")

    def _drag_enter_event(self, event):
        """Accept drag if it contains file URLs."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def _drag_move_event(self, event):
        """Accept drag move if it contains file URLs."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def _drop_event(self, event):
        """Handle files dropped from Explorer."""
        if not event.mimeData().hasUrls():
            event.ignore()
            return

        attached = 0
        errors = []

        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if not file_path or not os.path.isfile(file_path):
                continue

            try:
                file_service.attach_file(self.bluebook_id, self.current_section, file_path)
                attached += 1
            except Exception as e:
                errors.append(f"{os.path.basename(file_path)}: {e}")

        if attached > 0:
            self._load_sections()
            self._load_files()

        if errors:
            QMessageBox.warning(
                self, "Drop Results",
                f"Attached {attached} file(s).\n\nErrors:\n" + "\n".join(errors))

        event.acceptProposedAction()

    def _on_file_context_menu(self, pos):
        """Show context menu on right-click in the file list."""
        item = self.file_list.itemAt(pos)
        menu = QMenu(self)

        if item:
            # Right-clicked on a file
            bf = item.data(Qt.UserRole)
            if not bf:
                return

            open_action = menu.addAction("Open")
            print_action = menu.addAction("Print")
            menu.addSeparator()

            share_action = None
            unshare_action = None

            if bf.is_shared:
                # This is a shared reference from another bluebook
                unshare_action = menu.addAction("Unshare from this Bluebook")
            else:
                if self.current_section in SHAREABLE_SECTIONS:
                    share_action = menu.addAction("Share to Other Bluebooks")

            menu.addSeparator()
            rename_action = menu.addAction("Rename")
            remove_action = menu.addAction("Remove")

            action = menu.exec(self.file_list.viewport().mapToGlobal(pos))
            if not action:
                return

            if action == open_action:
                try:
                    file_service.open_file(bf.file_path)
                except FileNotFoundError as e:
                    QMessageBox.warning(self, "File Not Found", str(e))
            elif action == print_action:
                print_service.print_file(bf.file_path)
                QMessageBox.information(self, "Print", "File sent to print.")
            elif action == share_action:
                dlg = ShareDialog(bf.id, self.bluebook_id, parent=self)
                if dlg.exec() == ShareDialog.Accepted:
                    count = sharing_service.share_file(bf.id, dlg.selected_bluebook_ids)
                    QMessageBox.information(
                        self, "Shared", f"File shared to {count} bluebook(s).")
                    self._load_files()
            elif action == unshare_action:
                fname = os.path.basename(bf.file_path)
                reply = QMessageBox.question(
                    self, "Unshare File",
                    f"Remove shared file '{fname}' from this bluebook?\n"
                    f"(The original file will not be deleted.)",
                    QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.Yes:
                    sharing_service.unshare_file_from_bluebook(bf.id, self.bluebook_id)
                    self._load_sections()
                    self._load_files()
            elif action == remove_action:
                # Select this item and call existing remove logic
                self.file_list.setCurrentItem(item)
                self._remove_file()
            elif action == rename_action:
                self._rename_file(bf)

        else:
            # Right-clicked on empty space
            if self.current_section in TEMPLATE_SECTIONS:
                create_action = menu.addAction("Create New")
            else:
                create_action = None

            attach_action = menu.addAction("Attach Files")

            action = menu.exec(self.file_list.viewport().mapToGlobal(pos))
            if not action:
                return

            if action == create_action:
                self._create_file()
            elif action == attach_action:
                self._attach_file()

    def _rename_file(self, bf):
        """Prompt user to rename a file."""
        old_name = os.path.basename(bf.file_path)
        # Strip .lnk for display
        display_name = old_name[:-4] if old_name.lower().endswith(".lnk") else old_name
        base, ext = os.path.splitext(display_name)

        new_name, ok = QInputDialog.getText(
            self, "Rename File",
            f"New name for '{old_name}':",
            text=base)

        if not ok or not new_name.strip():
            return

        new_name = new_name.strip()
        if not new_name.lower().endswith(ext.lower()):
            new_name += ext

        if new_name == old_name:
            return

        try:
            file_service.rename_file(bf.id, new_name)
            self._load_sections()
            self._load_files()
        except PermissionError as e:
            QMessageBox.warning(self, "File In Use", str(e))
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))

    def _get_selected_files(self):
        """Return list of BluebookFile objects for selected items."""
        files = []
        for item in self.file_list.selectedItems():
            bf = item.data(Qt.UserRole)
            if bf:
                files.append(bf)
        return files

    def _open_file(self):
        files = self._get_selected_files()
        if not files:
            QMessageBox.information(self, "No Selection", "Please select a file to open.")
            return
        for bf in files:
            try:
                file_service.open_file(bf.file_path)
            except FileNotFoundError as e:
                QMessageBox.warning(self, "File Not Found", str(e))

    def _print_selected(self):
        files = self._get_selected_files()
        if not files:
            # Print entire section
            reply = QMessageBox.question(
                self, "Print Section",
                f"No files selected. Print entire section "
                f"'{SECTION_LABELS[self.current_section]}'?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                count = print_service.print_section(self.bluebook_id, self.current_section)
                QMessageBox.information(self, "Print", f"Sent {count} file(s) to print.")
        else:
            for bf in files:
                print_service.print_file(bf.file_path)
            QMessageBox.information(self, "Print",
                                    f"Sent {len(files)} file(s) to print.")

    def _print_all(self):
        reply = QMessageBox.question(
            self, "Print All",
            f"Print all files in Bluebook Die# {self.bluebook.die_number}?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            count = print_service.print_all(self.bluebook_id)
            QMessageBox.information(self, "Print All",
                                    f"Sent {count} file(s) to print.")

    def _create_file(self):
        dlg = CreateFileDialog(default_section=self.current_section,
                               die_number=self.bluebook.die_number, parent=self)
        if dlg.exec() == CreateFileDialog.Accepted:
            try:
                file_service.create_from_template(
                    self.bluebook_id, dlg.result_section, dlg.result_filename)
                self._load_sections()
                # Re-select the section where file was created
                for i in range(self.section_list.count()):
                    if self.section_list.item(i).data(Qt.UserRole) == dlg.result_section:
                        self.section_list.setCurrentRow(i)
                        break
                QMessageBox.information(self, "Created",
                                        f"File '{dlg.result_filename}.docx' created.")
            except Exception as e:
                QMessageBox.warning(self, "Error", str(e))

    def _attach_file(self):
        dlg = AttachFileDialog(default_section=self.current_section, parent=self)
        if dlg.exec() == AttachFileDialog.Accepted:
            attached = 0
            errors = []
            for path in dlg.result_file_paths:
                try:
                    file_service.attach_file(
                        self.bluebook_id, dlg.result_section, path)
                    attached += 1
                except Exception as e:
                    errors.append(f"{os.path.basename(path)}: {e}")

            self._load_sections()
            for i in range(self.section_list.count()):
                if self.section_list.item(i).data(Qt.UserRole) == dlg.result_section:
                    self.section_list.setCurrentRow(i)
                    break

            if errors:
                QMessageBox.warning(
                    self, "Partial Attach",
                    f"Attached {attached} file(s).\n\n"
                    f"Errors:\n" + "\n".join(errors))
            else:
                QMessageBox.information(
                    self, "Attached",
                    f"{attached} file(s) attached successfully.")

    def _remove_file(self):
        files = self._get_selected_files()
        if not files:
            QMessageBox.information(self, "No Selection", "Please select a file to remove.")
            return

        # Require password for deletion
        from services.security import require_password
        if not require_password("Deleting file(s)", self):
            return

        for bf in files:
            import os
            fname = os.path.basename(bf.file_path)
            # Strip .lnk for display
            if fname.lower().endswith(".lnk"):
                fname = fname[:-4]

            if bf.is_shared:
                # This is a shared reference — only remove the link
                reply = QMessageBox.question(
                    self, "Remove Shared File",
                    f"Remove shared file '{fname}' from this bluebook?\n"
                    f"(The original file will not be deleted.)",
                    QMessageBox.Yes | QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    sharing_service.unshare_file_from_bluebook(bf.id, self.bluebook_id)
            else:
                # Own file
                is_shared = sharing_service.is_file_shared(bf.id)
                if is_shared:
                    msg = (f"'{fname}' is shared with other bluebooks.\n\n"
                           f"• Remove from this bluebook only (keeps file, removes shares)\n"
                           f"• Cancel")
                    reply = QMessageBox.question(
                        self, "Remove Shared File", msg,
                        QMessageBox.Yes | QMessageBox.No,
                    )
                    if reply == QMessageBox.Yes:
                        try:
                            file_service.remove_file(bf.id, delete_from_disk=True)
                        except PermissionError as e:
                            QMessageBox.warning(self, "File In Use", str(e))
                else:
                    reply = QMessageBox.question(
                        self, "Remove File",
                        f"Remove '{fname}' from this bluebook and delete from disk?",
                        QMessageBox.Yes | QMessageBox.No,
                    )
                    if reply == QMessageBox.Yes:
                        try:
                            file_service.remove_file(bf.id, delete_from_disk=True)
                        except PermissionError as e:
                            QMessageBox.warning(self, "File In Use", str(e))

        self._load_sections()
        self._load_files()

    def _share_file(self):
        files = self._get_selected_files()
        if not files:
            QMessageBox.information(self, "No Selection",
                                    "Please select a file to share.")
            return

        # Only share own files (not already-shared references)
        own_files = [f for f in files if not f.is_shared]
        if not own_files:
            QMessageBox.information(self, "Cannot Share",
                                    "Shared files cannot be re-shared. "
                                    "Select an original file.")
            return

        bf = own_files[0]  # Share one file at a time
        dlg = ShareDialog(bf.id, self.bluebook_id, parent=self)
        if dlg.exec() == ShareDialog.Accepted:
            count = sharing_service.share_file(bf.id, dlg.selected_bluebook_ids)
            QMessageBox.information(
                self, "Shared",
                f"File shared to {count} bluebook(s).")
            self._load_files()
