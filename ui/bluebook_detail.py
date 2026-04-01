"""
Bluebook Manager - Bluebook Detail Screen.
"""

import hashlib
import os
import tempfile
import threading

from PySide6.QtCore import QSignalBlocker, QThread, QTimer, Qt, Signal
from PySide6.QtGui import QImage, QPixmap, QTransform
from PySide6.QtWidgets import QFrame, QHBoxLayout, QInputDialog, QLabel, QListWidget, QListWidgetItem, QMenu, QMessageBox, QPushButton, QScrollArea, QSplitter, QVBoxLayout, QWidget

from config import SECTION_LABELS, SECTION_TYPES, SHAREABLE_SECTIONS, TEMPLATE_SECTIONS
from services import bluebook_service, file_service, print_service, sharing_service
from services.log_service import log
from ui.dialogs.attach_file_dialog import AttachFileDialog
from ui.dialogs.create_file_dialog import CreateFileDialog
from ui.dialogs.share_dialog import ShareDialog

_DOCX_CACHE_DIR = os.path.join(tempfile.gettempdir(), "bluebook_docx_cache")
os.makedirs(_DOCX_CACHE_DIR, exist_ok=True)


def _fitz_pixmap_to_qimage(pix) -> QImage:
    if getattr(pix, "alpha", 0):
        return QImage(
            pix.samples,
            pix.width,
            pix.height,
            pix.stride,
            QImage.Format_RGBA8888,
        ).copy()
    if getattr(pix, "n", 0) < 3:
        import fitz

        pix = fitz.Pixmap(fitz.csRGB, pix)
    return QImage(
        pix.samples,
        pix.width,
        pix.height,
        pix.stride,
        QImage.Format_RGB888,
    ).copy()


class _WordCOMPool(QThread):
    conversion_done = Signal(str, str)
    conversion_error = Signal(str, str)

    def __init__(self):
        super().__init__()
        self._queue = []
        self._lock = threading.Lock()
        self._event = threading.Event()
        self._stop = False
        self._word = None

    def request_conversion(self, request_id: str, docx_path: str, cache_path: str):
        with self._lock:
            self._queue.append((request_id, docx_path, cache_path))
        self._event.set()

    def shutdown(self):
        self._stop = True
        self._event.set()
        self.wait(5000)

    def run(self):
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            self._word = win32com.client.DispatchEx("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = 0
        except Exception:
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
                    self.conversion_error.emit(req_id, "Microsoft Word could not be started.")
                    continue
                try:
                    doc = self._word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
                    doc.SaveAs(os.path.abspath(cache_path), FileFormat=17)
                    doc.Close(False)
                    self.conversion_done.emit(req_id, cache_path)
                except Exception as e:
                    self.conversion_error.emit(req_id, str(e))
                    try:
                        self._word = win32com.client.DispatchEx("Word.Application")
                        self._word.Visible = False
                        self._word.DisplayAlerts = 0
                    except Exception:
                        self._word = None
        if self._word is not None:
            try:
                self._word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


class _PreviewRenderPool(QThread):
    render_done = Signal(str, object)
    render_error = Signal(str, str)

    def __init__(self):
        super().__init__()
        self._request: tuple[str, str] | None = None
        self._lock = threading.Lock()
        self._event = threading.Event()
        self._stop = False

    def request_render(self, request_id: str, pdf_path: str):
        with self._lock:
            self._request = (request_id, pdf_path)
        self._event.set()

    def shutdown(self):
        self._stop = True
        self._event.set()
        self.wait(5000)

    def run(self):
        while not self._stop:
            self._event.wait()
            self._event.clear()
            while True:
                with self._lock:
                    request = self._request
                    self._request = None
                if request is None:
                    break
                request_id, pdf_path = request
                try:
                    import fitz

                    pdf_doc = fitz.open(pdf_path)
                    try:
                        if len(pdf_doc) == 0:
                            self.render_error.emit(request_id, "PDF has no pages.")
                            continue
                        page = pdf_doc[0]
                        pix = page.get_pixmap(matrix=fitz.Matrix(1.35, 1.35))
                        img = _fitz_pixmap_to_qimage(pix)
                    finally:
                        pdf_doc.close()
                    self.render_done.emit(request_id, img)
                except ImportError:
                    self.render_error.emit(request_id, "PDF preview requires PyMuPDF.\n\nUse Open to view the PDF.")
                except Exception as e:
                    self.render_error.emit(request_id, str(e))


_word_pool: _WordCOMPool | None = None
_preview_render_pool: _PreviewRenderPool | None = None


def _get_word_pool() -> _WordCOMPool:
    global _word_pool
    if _word_pool is None or not _word_pool.isRunning():
        _word_pool = _WordCOMPool()
        _word_pool.start()
    return _word_pool


def _get_preview_render_pool() -> _PreviewRenderPool:
    global _preview_render_pool
    if _preview_render_pool is None or not _preview_render_pool.isRunning():
        _preview_render_pool = _PreviewRenderPool()
        _preview_render_pool.start()
    return _preview_render_pool


def shutdown_word_pool():
    global _word_pool
    if _word_pool is not None:
        _word_pool.shutdown()
        _word_pool = None


def shutdown_preview_render_pool():
    global _preview_render_pool
    if _preview_render_pool is not None:
        _preview_render_pool.shutdown()
        _preview_render_pool = None


class ScaledPixmapLabel(QLabel):
    def __init__(self, pixmap, parent=None):
        super().__init__(parent)
        self._original_pixmap = pixmap
        self.setAlignment(Qt.AlignCenter)
        self._update_pixmap()

    def _update_pixmap(self):
        if not self._original_pixmap or self._original_pixmap.isNull():
            return
        try:
            available_width = self.width() - 10
            scaled = self._original_pixmap.scaledToWidth(available_width, Qt.SmoothTransformation) if available_width > 0 and self._original_pixmap.width() > available_width else self._original_pixmap
            super().setPixmap(scaled)
            self.setFixedHeight(scaled.height() + 10)
        except Exception:
            pass

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_pixmap()


class BluebookDetailWidget(QWidget):
    closed = Signal()

    def __init__(self, bluebook_id: int, parent=None):
        super().__init__(parent)
        self.bluebook_id = bluebook_id
        self.bluebook = bluebook_service.get_bluebook(bluebook_id)
        self.current_section = SECTION_TYPES[0]
        self._docx_request_id: str | None = None
        self._preview_request_id: str | None = None
        self._preview_generation = 0
        self._selection_generation = 0
        self._is_closing = False
        self._pending_preview_file = None
        self._pending_preview_generation = 0
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.timeout.connect(self._run_pending_preview)
        self._build_ui()
        self._connect_docx_preview_signals()
        self._connect_preview_render_signals()
        self.destroyed.connect(self._on_destroyed)
        self._load_sections()
        log("OPEN_BLUEBOOK", f"Die# {self.bluebook.die_number}")

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 18, 18, 18)
        main_layout.setSpacing(14)

        header_layout = QHBoxLayout()
        back_btn = QPushButton("Back")
        back_btn.setObjectName("ghostButton")
        back_btn.setFixedWidth(90)
        back_btn.clicked.connect(self.closed.emit)
        header_layout.addWidget(back_btn)

        title_block = QVBoxLayout()
        title_block.setSpacing(0)
        title = QLabel(f"Bluebook: Die# {self.bluebook.die_number}")
        title.setObjectName("headerLabel")
        title_block.addWidget(title)
        header_layout.addLayout(title_block, 1)

        self.print_all_btn = QPushButton("Print All")
        self.print_all_btn.setObjectName("primaryButton")
        self.print_all_btn.setFixedWidth(120)
        self.print_all_btn.clicked.connect(self._print_all)
        header_layout.addWidget(self.print_all_btn)
        main_layout.addLayout(header_layout)

        info_card = QFrame()
        info_card.setObjectName("surfaceCard")
        info_layout = QVBoxLayout(info_card)
        info_layout.setContentsMargins(16, 14, 16, 14)
        info_layout.setSpacing(6)
        if self.bluebook.description:
            desc = QLabel(self.bluebook.description)
            desc.setObjectName("panelTitle")
            info_layout.addWidget(desc)
        customers_text = ", ".join(self.bluebook.customer_names) or "No customers linked"
        info_parts = [f'<span style="color: #2a02a3;">Customer: {customers_text}</span>']
        if self.bluebook.outsource_names:
            outsource_text = ", ".join(self.bluebook.outsource_names)
            info_parts.append(f'<span style="color: #f9e2af;">Outsource: {outsource_text}</span>')
        info_label = QLabel("  |  ".join(info_parts))
        info_label.setObjectName("panelCaption")
        info_label.setTextFormat(Qt.RichText)
        info_layout.addWidget(info_label)
        main_layout.addWidget(info_card)

        main_splitter = QSplitter(Qt.Horizontal)

        left_panel = QFrame()
        left_panel.setObjectName("surfaceCard")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(16, 16, 16, 16)
        sec_label = QLabel("Sections")
        sec_label.setObjectName("sectionHeader")
        left_layout.addWidget(sec_label)
        self.section_list = QListWidget()
        self.section_list.currentRowChanged.connect(self._on_section_changed)
        left_layout.addWidget(self.section_list)
        main_splitter.addWidget(left_panel)

        right_splitter = QSplitter(Qt.Horizontal)

        files_panel = QFrame()
        files_panel.setObjectName("surfaceCard")
        files_layout = QVBoxLayout(files_panel)
        files_layout.setContentsMargins(16, 16, 16, 16)
        files_layout.setSpacing(10)
        self.files_header = QLabel("Files")
        self.files_header.setObjectName("sectionHeader")
        files_layout.addWidget(self.files_header)
        self.files_hint = QLabel("Attach files, create templates, or select one to preview.")
        self.files_hint.setObjectName("panelCaption")
        files_layout.addWidget(self.files_hint)

        actions_row = QHBoxLayout()
        actions_row.setSpacing(8)
        self.create_btn = QPushButton("Create")
        self.create_btn.setObjectName("successButton")
        self.create_btn.clicked.connect(self._create_file)
        actions_row.addWidget(self.create_btn)
        self.attach_btn = QPushButton("Attach")
        self.attach_btn.setObjectName("ghostButton")
        self.attach_btn.clicked.connect(self._attach_file)
        actions_row.addWidget(self.attach_btn)
        self.open_btn = QPushButton("Open")
        self.open_btn.setObjectName("ghostButton")
        self.open_btn.clicked.connect(self._open_file)
        actions_row.addWidget(self.open_btn)
        self.print_btn = QPushButton("Print")
        self.print_btn.setObjectName("ghostButton")
        self.print_btn.clicked.connect(self._print_selected)
        actions_row.addWidget(self.print_btn)
        self.share_btn = QPushButton("Share")
        self.share_btn.setObjectName("shareButton")
        self.share_btn.clicked.connect(self._share_file)
        actions_row.addWidget(self.share_btn)
        self.remove_btn = QPushButton("Remove")
        self.remove_btn.setObjectName("dangerButton")
        self.remove_btn.clicked.connect(self._remove_file)
        actions_row.addWidget(self.remove_btn)
        files_layout.addLayout(actions_row)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_list.currentItemChanged.connect(self._on_file_selected)
        self.file_list.itemSelectionChanged.connect(self._update_file_actions)
        self.file_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self._on_file_context_menu)
        self.file_list.setAcceptDrops(True)
        self.file_list.setDragDropMode(QListWidget.DropOnly)
        self.file_list.dragEnterEvent = self._drag_enter_event
        self.file_list.dragMoveEvent = self._drag_move_event
        self.file_list.dropEvent = self._drop_event
        files_layout.addWidget(self.file_list)
        right_splitter.addWidget(files_panel)

        preview_panel = QFrame()
        preview_panel.setObjectName("surfaceCard")
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(16, 16, 16, 16)
        preview_layout.setSpacing(10)
        preview_header_layout = QHBoxLayout()
        preview_header = QLabel("Preview")
        preview_header.setObjectName("sectionHeader")
        preview_header_layout.addWidget(preview_header)
        preview_header_layout.addStretch()
        self.preview_meta = QLabel("No file selected")
        self.preview_meta.setObjectName("panelCaption")
        preview_header_layout.addWidget(self.preview_meta)
        self.btn_rotate = QPushButton("Rotate")
        self.btn_rotate.setObjectName("ghostButton")
        self.btn_rotate.setFixedWidth(90)
        self.btn_rotate.clicked.connect(self._rotate_preview)
        preview_header_layout.addWidget(self.btn_rotate)
        preview_layout.addLayout(preview_header_layout)
        self._rotation_angle = 0
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(True)
        self.preview_scroll.setStyleSheet("QScrollArea { border: 1px solid #45475a; border-radius: 10px; background-color: #313244; }")
        self.preview_label = QLabel("Select a file to preview")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setWordWrap(True)
        self.preview_label.setStyleSheet("color: #6c7086; font-size: 14px; padding: 20px;")
        self.preview_scroll.setWidget(self.preview_label)
        preview_layout.addWidget(self.preview_scroll)
        right_splitter.addWidget(preview_panel)

        right_splitter.setSizes([320, 760])
        main_splitter.addWidget(right_splitter)
        main_splitter.setSizes([260, 900])
        main_layout.addWidget(main_splitter, 1)
        self._update_file_actions()

    def _update_file_actions(self):
        selected_files = self._get_selected_files()
        has_selection = bool(selected_files)
        has_single = len(selected_files) == 1
        own_files = [f for f in selected_files if not f.is_shared]
        shareable_selection = any(f.section_type in SHAREABLE_SECTIONS for f in own_files)
        self.create_btn.setEnabled(self.current_section in TEMPLATE_SECTIONS)
        self.attach_btn.setEnabled(True)
        self.open_btn.setEnabled(has_selection)
        self.print_btn.setEnabled(True)
        self.remove_btn.setEnabled(has_selection)
        self.share_btn.setEnabled(has_single and bool(own_files) and shareable_selection)
        self.btn_rotate.setEnabled(has_single)

    def _load_sections(self):
        self.section_list.clear()
        section_counts = file_service.get_section_file_counts(self.bluebook_id)
        for sec in SECTION_TYPES:
            count = section_counts.get(sec, 0)
            label = f"{SECTION_LABELS[sec]} ({count})" if count else SECTION_LABELS[sec]
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, sec)
            self.section_list.addItem(item)
        self.section_list.setCurrentRow(0)

    def _on_section_changed(self, row):
        if row < 0:
            return
        item = self.section_list.item(row)
        self.current_section = item.data(Qt.UserRole)
        self.files_header.setText(SECTION_LABELS[self.current_section])
        self.files_hint.setText(
            "This section supports creating a new document from a template."
            if self.current_section in TEMPLATE_SECTIONS
            else "Attach files here or select one to open, print, or preview."
        )
        self._load_files()
        self._update_file_actions()

    def _load_files(self):
        blocker = QSignalBlocker(self.file_list)
        self.file_list.clear()
        files = file_service.get_files_for_section(self.bluebook_id, self.current_section)
        shared_original_ids = file_service.get_shared_original_file_ids([bf.id for bf in files if not bf.is_shared])
        for bf in files:
            name = os.path.basename(bf.file_path)
            if name.lower().endswith(".lnk"):
                name = name[:-4]
            if bf.is_shared:
                name = f"Shared: {name}" if not bf.shared_from_die_number else f"Shared: {name}  (from Die# {bf.shared_from_die_number})"
            elif bf.id in shared_original_ids:
                name = f"{name}  [shared]"
            item = QListWidgetItem(name)
            item.setData(Qt.UserRole, bf)
            if bf.is_shared:
                item.setToolTip(f"Shared from Die# {bf.shared_from_die_number}")
            self.file_list.addItem(item)
        del blocker
        self._clear_preview()
        self._update_file_actions()

    def _on_file_selected(self, current, previous):
        self._cancel_pending_docx_preview()
        self._rotation_angle = 0
        self._selection_generation += 1
        selection_generation = self._selection_generation
        if not current:
            self.preview_meta.setText("No file selected")
            self._clear_preview()
            return
        bf = current.data(Qt.UserRole)
        if not bf:
            self.preview_meta.setText("No file selected")
            self._clear_preview()
            return
        display_name = os.path.basename(bf.file_path)
        self.preview_meta.setText(display_name[:-4] if display_name.lower().endswith(".lnk") else display_name)
        self._pending_preview_file = bf
        self._pending_preview_generation = selection_generation
        self._preview_timer.start(120)
        self._update_file_actions()

    def _run_pending_preview(self):
        bf = self._pending_preview_file
        generation = self._pending_preview_generation
        self._pending_preview_file = None
        if bf is None:
            return
        self._render_selected_file_preview(bf, generation)

    def _render_selected_file_preview(self, bf, selection_generation: int):
        if not self._is_preview_request_current(selection_generation, bf):
            return

        abs_path = file_service.get_absolute_path(bf.file_path)
        if not os.path.isfile(abs_path):
            self._show_preview_message("File not found on disk.")
            return
        abs_path = file_service.resolve_shortcut(abs_path)
        if not os.path.isfile(abs_path):
            self._show_preview_message("Shortcut target not found.\n\nThe file on the network drive may have been moved or deleted.")
            return
        try:
            ext = os.path.splitext(abs_path)[1].lower()
            if ext in (".png", ".jpg", ".jpeg", ".bmp", ".gif"):
                self._preview_image(abs_path, selection_generation)
            elif ext == ".pdf":
                self._preview_pdf(abs_path, selection_generation)
            elif ext == ".docx":
                self._preview_docx(abs_path, selection_generation)
            else:
                self._show_preview_message(f"No preview available for {ext} files.\n\nUse Open to view the file.")
        except Exception as e:
            self._show_preview_message(f"Could not load preview:\n{e}")

    def _clear_preview(self):
        self._cancel_pending_docx_preview()
        self._rotation_angle = 0
        self._show_preview_message("Select a file to preview")

    def _connect_docx_preview_signals(self):
        pool = _get_word_pool()
        pool.conversion_done.connect(self._on_docx_ready)
        pool.conversion_error.connect(self._on_docx_error)

    def _connect_preview_render_signals(self):
        pool = _get_preview_render_pool()
        pool.render_done.connect(self._on_preview_render_done)
        pool.render_error.connect(self._on_preview_render_error)

    def _cancel_pending_docx_preview(self):
        self._preview_generation += 1
        self._docx_request_id = None
        self._preview_request_id = None
        self._pending_preview_file = None
        self._preview_timer.stop()

    def _on_destroyed(self, *args):
        self._is_closing = True
        self._docx_request_id = None
        self._preview_request_id = None
        self._pending_preview_file = None
        self._preview_timer.stop()

    def _rotate_preview(self):
        self._rotation_angle = (self._rotation_angle + 90) % 360
        transform = QTransform().rotate(90)
        widget = self.preview_scroll.widget()
        if not widget:
            return
        labels = widget.findChildren(ScaledPixmapLabel)
        if not labels and isinstance(widget, ScaledPixmapLabel):
            labels = [widget]
        for label in labels:
            if label._original_pixmap and not label._original_pixmap.isNull():
                label._original_pixmap = label._original_pixmap.transformed(transform, Qt.SmoothTransformation)
                label._update_pixmap()

    def _show_preview_message(self, text):
        self.preview_label = QLabel(text)
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setWordWrap(True)
        self.preview_label.setStyleSheet("color: #6c7086; font-size: 14px; padding: 20px;")
        self._set_preview_widget(self.preview_label)

    def _set_preview_widget(self, widget: QWidget):
        if self._is_closing:
            widget.deleteLater()
            return
        old_widget = self.preview_scroll.takeWidget()
        if old_widget is not None:
            old_widget.deleteLater()
        self.preview_scroll.setWidget(widget)

    def _is_preview_request_current(self, selection_generation: int, bf=None) -> bool:
        if self._is_closing or selection_generation != self._selection_generation:
            return False
        current = self.file_list.currentItem()
        if current is None:
            return False
        current_bf = current.data(Qt.UserRole)
        if bf is not None and current_bf is not None:
            return getattr(current_bf, "id", None) == getattr(bf, "id", None)
        return current_bf is not None

    def _preview_image(self, abs_path, selection_generation: int):
        if not self._is_preview_request_current(selection_generation):
            return
        pixmap = QPixmap(abs_path)
        if pixmap.isNull():
            self._show_preview_message("Could not load image.")
            return
        label = ScaledPixmapLabel(pixmap)
        label.setStyleSheet("padding: 10px;")
        if not self._is_preview_request_current(selection_generation):
            label.deleteLater()
            return
        self._set_preview_widget(label)

    def _preview_pdf(self, abs_path, selection_generation: int):
        if not self._is_preview_request_current(selection_generation):
            return
        request_id = f"preview:{selection_generation}:{self._preview_generation}:{abs_path}"
        self._preview_request_id = request_id
        self._show_preview_message("Loading preview...")
        _get_preview_render_pool().request_render(request_id, abs_path)

    def _preview_docx(self, abs_path, selection_generation: int):
        if not self._is_preview_request_current(selection_generation):
            return
        try:
            mtime = os.path.getmtime(abs_path)
        except OSError:
            mtime = 0
        cache_key = hashlib.md5(f"{abs_path}|{mtime}".encode()).hexdigest()
        cache_path = os.path.join(_DOCX_CACHE_DIR, f"{cache_key}.pdf")
        if os.path.isfile(cache_path):
            self._render_pdf_preview(cache_path, selection_generation)
            return
        self._show_preview_message("Loading preview...")
        request_id = f"{cache_key}:{self._preview_generation}:{selection_generation}"
        self._docx_request_id = request_id
        _get_word_pool().request_conversion(request_id, abs_path, cache_path)

    def _on_docx_ready(self, request_id: str, pdf_path: str):
        if self._is_closing or request_id != getattr(self, "_docx_request_id", None):
            return
        self._docx_request_id = None
        try:
            selection_generation = int(request_id.rsplit(":", 1)[1])
        except (IndexError, ValueError):
            selection_generation = self._selection_generation
        self._render_pdf_preview(pdf_path, selection_generation)

    def _on_docx_error(self, request_id: str, error_msg: str):
        if self._is_closing or request_id != getattr(self, "_docx_request_id", None):
            return
        self._docx_request_id = None
        self._show_preview_message(f"Could not convert DOCX to PDF:\n{error_msg}\n\nMake sure Microsoft Word is installed.")

    def _render_pdf_preview(self, pdf_path: str, selection_generation: int):
        if not self._is_preview_request_current(selection_generation):
            return
        self._preview_pdf(pdf_path, selection_generation)

    def _on_preview_render_done(self, request_id: str, img):
        if self._is_closing or request_id != self._preview_request_id:
            return
        self._preview_request_id = None
        try:
            selection_generation = int(request_id.split(":", 3)[1])
        except (IndexError, ValueError):
            selection_generation = self._selection_generation
        if not self._is_preview_request_current(selection_generation):
            return
        label = ScaledPixmapLabel(QPixmap.fromImage(img))
        label.setStyleSheet("background-color: white; border: 1px solid #45475a; padding: 2px;")
        self._set_preview_widget(label)

    def _on_preview_render_error(self, request_id: str, error_msg: str):
        if self._is_closing or request_id != self._preview_request_id:
            return
        self._preview_request_id = None
        self._show_preview_message(f"Could not preview PDF:\n{error_msg}")

    def _drag_enter_event(self, event):
        event.acceptProposedAction() if event.mimeData().hasUrls() else event.ignore()

    def _drag_move_event(self, event):
        event.acceptProposedAction() if event.mimeData().hasUrls() else event.ignore()

    def _drop_event(self, event):
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
            QMessageBox.warning(self, "Drop Results", f"Attached {attached} file(s).\n\nErrors:\n" + "\n".join(errors))
        event.acceptProposedAction()

    def _on_file_context_menu(self, pos):
        item = self.file_list.itemAt(pos)
        menu = QMenu(self)
        if item:
            bf = item.data(Qt.UserRole)
            if not bf:
                return
            open_action = menu.addAction("Open")
            print_action = menu.addAction("Print")
            menu.addSeparator()
            share_action = None
            unshare_action = None
            if bf.is_shared:
                unshare_action = menu.addAction("Unshare from this Bluebook")
            elif self.current_section in SHAREABLE_SECTIONS:
                share_action = menu.addAction("Share to Other Bluebooks")
            menu.addSeparator()
            rename_action = menu.addAction("Rename")
            remove_action = menu.addAction("Remove")
            action = menu.exec(self.file_list.viewport().mapToGlobal(pos))
            if not action:
                return
            if action == open_action:
                self._open_file()
            elif action == print_action:
                self._print_selected()
            elif action == share_action:
                self._share_file()
            elif action == unshare_action:
                reply = QMessageBox.question(self, "Unshare File", f"Remove shared file '{os.path.basename(bf.file_path)}' from this bluebook?\n(The original file will not be deleted.)", QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.Yes:
                    sharing_service.unshare_file_from_bluebook(bf.id, self.bluebook_id)
                    self._load_sections()
                    self._load_files()
            elif action == remove_action:
                self.file_list.setCurrentItem(item)
                self._remove_file()
            elif action == rename_action:
                self._rename_file(bf)
        else:
            create_action = menu.addAction("Create New") if self.current_section in TEMPLATE_SECTIONS else None
            attach_action = menu.addAction("Attach Files")
            action = menu.exec(self.file_list.viewport().mapToGlobal(pos))
            if not action:
                return
            if action == create_action:
                self._create_file()
            elif action == attach_action:
                self._attach_file()

    def _rename_file(self, bf):
        old_name = os.path.basename(bf.file_path)
        display_name = old_name[:-4] if old_name.lower().endswith(".lnk") else old_name
        base, ext = os.path.splitext(display_name)
        new_name, ok = QInputDialog.getText(self, "Rename File", f"New name for '{old_name}':", text=base)
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
        return [item.data(Qt.UserRole) for item in self.file_list.selectedItems() if item.data(Qt.UserRole)]

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
            reply = QMessageBox.question(self, "Print Section", f"No files selected. Print entire section '{SECTION_LABELS[self.current_section]}'?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                count = print_service.print_section(self.bluebook_id, self.current_section)
                QMessageBox.information(self, "Print", f"Sent {count} file(s) to print.")
            return
        for bf in files:
            print_service.print_file(bf.file_path)
        QMessageBox.information(self, "Print", f"Sent {len(files)} file(s) to print.")

    def _print_all(self):
        reply = QMessageBox.question(self, "Print All", f"Print all files in Bluebook Die# {self.bluebook.die_number}?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            count = print_service.print_all(self.bluebook_id)
            QMessageBox.information(self, "Print All", f"Sent {count} file(s) to print.")

    def _create_file(self):
        dlg = CreateFileDialog(default_section=self.current_section, die_number=self.bluebook.die_number, parent=self)
        if dlg.exec() == CreateFileDialog.Accepted:
            try:
                file_service.create_from_template(self.bluebook_id, dlg.result_section, dlg.result_filename)
                self._load_sections()
                for i in range(self.section_list.count()):
                    if self.section_list.item(i).data(Qt.UserRole) == dlg.result_section:
                        self.section_list.setCurrentRow(i)
                        break
                QMessageBox.information(self, "Created", f"File '{dlg.result_filename}.docx' created.")
            except Exception as e:
                QMessageBox.warning(self, "Error", str(e))

    def _attach_file(self):
        dlg = AttachFileDialog(default_section=self.current_section, parent=self)
        if dlg.exec() == AttachFileDialog.Accepted:
            attached = 0
            errors = []
            for path in dlg.result_file_paths:
                try:
                    file_service.attach_file(self.bluebook_id, dlg.result_section, path)
                    attached += 1
                except Exception as e:
                    errors.append(f"{os.path.basename(path)}: {e}")
            self._load_sections()
            for i in range(self.section_list.count()):
                if self.section_list.item(i).data(Qt.UserRole) == dlg.result_section:
                    self.section_list.setCurrentRow(i)
                    break
            if errors:
                QMessageBox.warning(self, "Partial Attach", f"Attached {attached} file(s).\n\nErrors:\n" + "\n".join(errors))
            else:
                QMessageBox.information(self, "Attached", f"{attached} file(s) attached successfully.")

    def _remove_file(self):
        files = self._get_selected_files()
        if not files:
            QMessageBox.information(self, "No Selection", "Please select a file to remove.")
            return
        from services.security import require_password
        if not require_password("Deleting file(s)", self):
            return
        for bf in files:
            fname = os.path.basename(bf.file_path)
            if fname.lower().endswith(".lnk"):
                fname = fname[:-4]
            if bf.is_shared:
                reply = QMessageBox.question(self, "Remove Shared File", f"Remove shared file '{fname}' from this bluebook?\n(The original file will not be deleted.)", QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.Yes:
                    sharing_service.unshare_file_from_bluebook(bf.id, self.bluebook_id)
            else:
                is_shared = sharing_service.is_file_shared(bf.id)
                if is_shared:
                    msg = f"'{fname}' is shared with other bluebooks.\n\nRemove from this bluebook only (keeps file, removes shares), or cancel."
                    reply = QMessageBox.question(self, "Remove Shared File", msg, QMessageBox.Yes | QMessageBox.No)
                    if reply == QMessageBox.Yes:
                        try:
                            file_service.remove_file(bf.id, delete_from_disk=True)
                        except PermissionError as e:
                            QMessageBox.warning(self, "File In Use", str(e))
                else:
                    reply = QMessageBox.question(self, "Remove File", f"Remove '{fname}' from this bluebook and delete from disk?", QMessageBox.Yes | QMessageBox.No)
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
            QMessageBox.information(self, "No Selection", "Please select a file to share.")
            return
        own_files = [f for f in files if not f.is_shared]
        if not own_files:
            QMessageBox.information(self, "Cannot Share", "Shared files cannot be re-shared. Select an original file.")
            return
        bf = own_files[0]
        dlg = ShareDialog(bf.id, self.bluebook_id, parent=self)
        if dlg.exec() == ShareDialog.Accepted:
            count = sharing_service.share_file(bf.id, dlg.selected_bluebook_ids)
            QMessageBox.information(self, "Shared", f"File shared to {count} bluebook(s).")
            self._load_files()
