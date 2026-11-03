"""WenShu document converter - PySide6 desktop interface."""

from __future__ import annotations

import logging
import logging.handlers
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import uuid
from dataclasses import dataclass
from pathlib import Path
from PySide6.QtCore import QByteArray, QSettings, QSize, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QCloseEvent, QDesktopServices, QFont, QGuiApplication, QIcon, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QStackedWidget,
    QStyleFactory,
    QTableWidget,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from conversion import (
    VERSION,
    WORD_EXTENSIONS,
    PDF_EXTENSIONS,
    find_word_engine,
    format_engine_summary,
    pdf_engine_available,
)
from icons import app_icon_path, app_logo, badge_pixmap, line_icon
from theme import APP_QSS
from widgets import (
    DropArea,
    ElidedLabel,
    ErrorDetailsDialog,
    FileCell,
    RowRemoveButton,
    StatusPill,
    shorten_path,
)

APP_TITLE = "文枢文档转换器"
APP_DATA = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "WenShuConverter"
SETTINGS_PATH = APP_DATA / "settings.ini"
CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


# ----------------------------------------------------------------------------
# Data and background conversion
# ----------------------------------------------------------------------------


@dataclass
class FileRecord:
    record_id: str
    path: Path
    size: int
    status: str = "waiting"  # waiting | working | success | error | cancel
    output: str = ""
    warning: str = ""
    error: str = ""
    detail: str = ""


@dataclass
class ConversionTask:
    record_id: str
    source: Path
    output_dir: Path


class ConversionWorker(QThread):
    file_started = Signal(str, int, int)
    file_done = Signal(str, str, str, str, str, str)
    all_finished = Signal(int, int, bool, float)

    def __init__(
        self,
        tasks: list[ConversionTask],
        mode: str,
        engine: dict[str, str] | None,
        pdf_mode: str = "flow",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.tasks = tasks
        self.mode = mode
        self.engine = engine
        self.pdf_mode = pdf_mode
        self._cancelled = threading.Event()
        self._process: subprocess.Popen | None = None
        self._process_lock = threading.Lock()

    def cancel(self) -> None:
        """Request cancellation without blocking the UI thread."""
        self._cancelled.set()
        with self._process_lock:
            process = self._process
        if process is None or process.poll() is not None:
            return
        if os.name == "nt":
            try:
                subprocess.Popen(
                    ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=CREATE_NO_WINDOW,
                )
                return
            except OSError:
                pass
        try:
            process.kill()
        except OSError:
            pass

    def run(self) -> None:  # noqa: D401 - QThread entry point
        started = time.perf_counter()
        success = 0
        failed = 0
        cancelled = False
        total = len(self.tasks)
        for index, task in enumerate(self.tasks, start=1):
            if self._cancelled.is_set():
                cancelled = True
                break
            self.file_started.emit(task.record_id, index, total)
            try:
                result = self._run_one(task)
            except Exception as exc:  # pragma: no cover - defensive boundary
                result = {
                    "state": "error",
                    "message": "转换进程发生意外错误。",
                    "detail": f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}",
                }
            if result["state"] == "cancelled":
                self.file_done.emit(task.record_id, "cancel", "", "", "已取消", "")
                cancelled = True
                break
            if result["state"] == "success":
                success += 1
                self.file_done.emit(
                    task.record_id,
                    "success",
                    result.get("output", ""),
                    result.get("warning", ""),
                    "",
                    "",
                )
            else:
                failed += 1
                self.file_done.emit(
                    task.record_id,
                    "error",
                    "",
                    "",
                    result.get("message", "转换失败"),
                    result.get("detail", ""),
                )
        if self._cancelled.is_set():
            cancelled = True
        self.all_finished.emit(success, failed, cancelled, round(time.perf_counter() - started, 2))

    def _run_one(self, task: ConversionTask) -> dict:
        work_dir = Path(tempfile.mkdtemp(prefix="wenshu-job-"))
        request_path = work_dir / "request.json"
        response_path = work_dir / "response.json"
        log_path = work_dir / "child.log"
        worker_log_path = work_dir / "worker.log"
        request = {
            "mode": self.mode,
            "source": str(task.source),
            "output_dir": str(task.output_dir),
            "engine": self.engine,
            "pdf_mode": self.pdf_mode,
            "response_path": str(response_path),
            "log_path": str(worker_log_path),
        }
        request_path.write_text(
            __import__("json").dumps(request, ensure_ascii=False), encoding="utf-8"
        )

        if getattr(sys, "frozen", False):
            command = [sys.executable, "--convert-one", str(request_path)]
        else:
            command = [
                sys.executable,
                str(Path(__file__).resolve().parent / "launcher.py"),
                "--convert-one",
                str(request_path),
            ]

        child_temp = work_dir / "tmp"
        child_temp.mkdir(parents=True, exist_ok=True)
        env = os.environ.copy()
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        env["WENSHU_WORKER"] = "1"
        env["TEMP"] = str(child_temp)
        env["TMP"] = str(child_temp)
        env["TMPDIR"] = str(child_temp)

        process: subprocess.Popen | None = None
        deadline = time.monotonic() + 30 * 60
        try:
            with log_path.open("wb") as log_stream:
                try:
                    process = subprocess.Popen(
                        command,
                        cwd=str(Path(__file__).resolve().parent),
                        stdin=subprocess.DEVNULL,
                        stdout=log_stream,
                        stderr=subprocess.STDOUT,
                        env=env,
                        creationflags=CREATE_NO_WINDOW,
                    )
                except OSError as exc:
                    shutil.rmtree(work_dir, ignore_errors=True)
                    return {
                        "state": "error",
                        "message": "无法启动转换进程。",
                        "detail": f"{type(exc).__name__}: {exc}",
                    }
                with self._process_lock:
                    self._process = process

                while True:
                    try:
                        process.wait(timeout=0.12)
                        break
                    except subprocess.TimeoutExpired:
                        if self._cancelled.is_set():
                            _kill_process_tree(process)
                            try:
                                process.wait(timeout=8)
                            except subprocess.TimeoutExpired:
                                process.kill()
                            break
                        if time.monotonic() > deadline:
                            _kill_process_tree(process)
                            try:
                                process.wait(timeout=8)
                            except subprocess.TimeoutExpired:
                                process.kill()
                            shutil.rmtree(work_dir, ignore_errors=True)
                            return {
                                "state": "error",
                                "message": "处理超时（超过 30 分钟），任务已停止。",
                                "detail": "子进程运行超过 30 分钟。",
                            }
        finally:
            with self._process_lock:
                self._process = None

        payload = None
        try:
            if response_path.exists():
                payload = __import__("json").loads(response_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            payload = None

        if isinstance(payload, dict) and payload.get("ok"):
            result = payload.get("result") or {}
            output = str(result.get("output", ""))
            warning = str(result.get("warning", ""))
            shutil.rmtree(work_dir, ignore_errors=True)
            return {
                "state": "success",
                "output": output,
                "warning": warning,
            }
        if self._cancelled.is_set():
            shutil.rmtree(work_dir, ignore_errors=True)
            return {"state": "cancelled"}

        if process is not None and process.returncode == 0 and payload is None:
            message = "转换进程没有返回结果。"
        elif isinstance(payload, dict):
            message = str(payload.get("message") or "转换失败。")
        else:
            message = "转换进程异常退出。"

        detail = ""
        if isinstance(payload, dict):
            detail = str(payload.get("detail") or "")
        if not detail:
            detail = _read_tail(log_path)
        worker_tail = _read_tail(worker_log_path)
        if worker_tail and worker_tail not in detail:
            detail = (detail + "\n\n" + worker_tail).strip()
        shutil.rmtree(work_dir, ignore_errors=True)
        return {"state": "error", "message": message, "detail": detail[-6000:]}


def _kill_process_tree(process: subprocess.Popen) -> None:
    if process.poll() is not None:
        return
    if os.name == "nt":
        try:
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                capture_output=True,
                creationflags=CREATE_NO_WINDOW,
                timeout=10,
            )
            return
        except (OSError, subprocess.TimeoutExpired):
            pass
    try:
        process.kill()
    except OSError:
        pass


def _read_tail(path: Path, limit: int = 6000) -> str:
    try:
        data = path.read_bytes()
    except OSError:
        return ""
    text = data[-limit:].decode("utf-8", errors="replace").strip()
    text = re.sub(r"\x1b\[[0-9;]*[A-Za-z]", "", text).strip()
    if not text:
        return "转换进程未返回任何诊断信息。"
    return text


# ----------------------------------------------------------------------------
# Converter page
# ----------------------------------------------------------------------------


class ConverterPage(QWidget):
    conversionStateChanged = Signal(bool)

    def __init__(self, mode: str, host: "MainWindow") -> None:
        super().__init__()
        self.mode = mode
        self.host = host
        self.settings = host.settings
        self.records: list[FileRecord] = []
        self.worker: ConversionWorker | None = None
        self._retired_workers: list[ConversionWorker] = []
        self._busy = False
        self._batch_total = 0
        self._batch_success = 0
        self._batch_failed = 0
        self._batch_warnings: list[str] = []
        self._current_record_id = ""
        self.custom_dir = str(self.settings.value("customOutputDir", "") or "")
        self.last_output_file = ""
        self.last_output_dir = ""
        self._build_ui()
        self._restore_size_preferences()
        self._refresh_action_state()
        self._update_queue_visibility()

    # -- construction ---------------------------------------------------------

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 18, 28, 12)
        outer.setSpacing(12)

        # Header
        header = QHBoxLayout()
        header.setSpacing(12)
        title_box = QVBoxLayout()
        title_box.setContentsMargins(0, 0, 0, 0)
        title_box.setSpacing(2)
        title = QLabel("Word 转 PDF" if self.mode == "word" else "PDF 转 Word")
        title.setObjectName("PageTitle")
        subtitle = QLabel(
            "将 DOCX / DOC 文档转换为排版稳定的 PDF，适合分享、打印与归档。"
            if self.mode == "word"
            else "将 PDF 还原为可编辑的 Word 文档，尽量保留文字、表格与图片。"
        )
        subtitle.setObjectName("PageSubtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        header.addLayout(title_box, 1)

        self.open_folder_button = QPushButton("打开输出文件夹")
        self.open_folder_button.setObjectName("SecondaryButton")
        self.open_folder_button.setFixedHeight(32)
        self.open_folder_button.setIcon(line_icon("folder-open", "#5C7089", 16, 1.7))
        self.open_folder_button.setIconSize(QSize(16, 16))
        self.open_folder_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.open_folder_button.setEnabled(False)
        self.open_folder_button.setToolTip("转换成功后可用")
        self.open_folder_button.clicked.connect(self.open_output_folder)
        header.addWidget(self.open_folder_button, 0, Qt.AlignmentFlag.AlignVCenter)
        outer.addLayout(header)

        # Drop area
        self.drop_area = DropArea(self.mode)
        self.drop_area.filesDropped.connect(self.add_files)
        self.drop_area.browseRequested.connect(self.browse_files)
        outer.addWidget(self.drop_area)

        # Queue card
        queue_card = QFrame()
        queue_card.setObjectName("QueueCard")
        queue_layout = QVBoxLayout(queue_card)
        queue_layout.setContentsMargins(16, 12, 16, 8)
        queue_layout.setSpacing(8)

        queue_header = QHBoxLayout()
        queue_header.setSpacing(8)
        queue_title = QLabel("转换队列")
        queue_title.setObjectName("SectionTitle")
        self.count_badge = QLabel("0")
        self.count_badge.setObjectName("CountBadge")
        self.count_badge.setFixedHeight(20)
        self.count_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        queue_header.addWidget(queue_title)
        queue_header.addWidget(self.count_badge)
        queue_header.addStretch(1)

        self.add_button = QPushButton("添加文件")
        self.add_button.setObjectName("GhostButton")
        self.add_button.setFixedHeight(28)
        self.add_button.setIcon(line_icon("plus", "#5D7089", 15, 1.8))
        self.add_button.setIconSize(QSize(15, 15))
        self.add_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_button.clicked.connect(self.browse_files)
        queue_header.addWidget(self.add_button)

        self.clear_button = QPushButton("清空")
        self.clear_button.setObjectName("GhostButton")
        self.clear_button.setFixedHeight(28)
        self.clear_button.setIcon(line_icon("trash", "#8C9BAD", 14, 1.6))
        self.clear_button.setIconSize(QSize(14, 14))
        self.clear_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_button.clicked.connect(self.clear_queue)
        queue_header.addWidget(self.clear_button)
        queue_layout.addLayout(queue_header)

        self.queue_stack = QStackedWidget()
        self.queue_stack.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.empty_page = self._build_empty_page()
        self.table = self._build_table()
        self.queue_stack.addWidget(self.empty_page)
        self.queue_stack.addWidget(self.table)
        queue_layout.addWidget(self.queue_stack, 1)
        outer.addWidget(queue_card, 1)

        # Output + action card
        action_card = QFrame()
        action_card.setObjectName("Card")
        action_layout = QVBoxLayout(action_card)
        action_layout.setContentsMargins(16, 11, 16, 11)
        action_layout.setSpacing(8)

        output_row = QHBoxLayout()
        output_row.setSpacing(8)
        output_label = QLabel("保存位置")
        output_label.setObjectName("FieldLabel")
        output_label.setFixedWidth(48)
        output_row.addWidget(output_label)

        self.source_radio = QRadioButton("原文件所在文件夹")
        self.source_radio.setChecked(True)
        self.source_radio.setMinimumWidth(132)
        self.source_radio.toggled.connect(self._sync_output_controls)
        output_row.addWidget(self.source_radio)

        self.custom_radio = QRadioButton("统一保存到")
        self.custom_radio.setMinimumWidth(96)
        self.custom_radio.toggled.connect(self._sync_output_controls)
        output_row.addWidget(self.custom_radio)

        self.output_path = QLineEdit()
        self.output_path.setObjectName("OutputPath")
        self.output_path.setReadOnly(True)
        self.output_path.setFixedHeight(28)
        self.output_path.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        output_row.addWidget(self.output_path, 1)

        self.browse_output_button = QPushButton("浏览")
        self.browse_output_button.setObjectName("SecondaryButton")
        self.browse_output_button.setFixedSize(58, 28)
        self.browse_output_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.browse_output_button.clicked.connect(self.choose_output_dir)
        output_row.addWidget(self.browse_output_button)
        action_layout.addLayout(output_row)

        self.pdf_style_widget = None
        self.flow_radio = None
        self.layout_radio = None
        if self.mode == "pdf":
            style_row = QHBoxLayout()
            style_row.setSpacing(8)
            style_label = QLabel("转换方式")
            style_label.setObjectName("FieldLabel")
            style_label.setFixedWidth(48)
            style_row.addWidget(style_label)

            self.flow_radio = QRadioButton("可编辑文档")
            self.flow_radio.setToolTip("按阅读顺序重排为普通段落、标题和表格，便于继续编辑")
            self.layout_radio = QRadioButton("保留原始版面")
            self.layout_radio.setToolTip("尽量保持 PDF 的精确位置，可能产生较多文本框和浮动表格")
            saved_mode = str(self.settings.value("pdfMode", "flow") or "flow")
            if saved_mode == "layout":
                self.layout_radio.setChecked(True)
            else:
                self.flow_radio.setChecked(True)
            self.flow_radio.toggled.connect(self._on_pdf_style_changed)
            self.layout_radio.toggled.connect(self._on_pdf_style_changed)
            style_row.addWidget(self.flow_radio)
            style_row.addWidget(self.layout_radio)
            style_hint = QLabel("推荐可编辑文档；需要精确对照原 PDF 时选择保留版面")
            style_hint.setObjectName("QueueHint")
            style_row.addWidget(style_hint, 1)
            action_layout.addLayout(style_row)

        divider = QFrame()
        divider.setObjectName("Divider")
        divider.setFixedHeight(1)
        action_layout.addWidget(divider)

        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(10)
        status_box = QVBoxLayout()
        status_box.setContentsMargins(0, 0, 0, 0)
        status_box.setSpacing(4)
        status_line = QHBoxLayout()
        status_line.setContentsMargins(0, 0, 0, 0)
        status_line.setSpacing(2)
        self.status_label = ElidedLabel("添加文件后即可开始转换")
        self.status_label.setObjectName("StatusText")
        self.status_label.setMinimumWidth(100)
        self.status_label.setFixedHeight(18)
        self.status_label.setProperty("state", "idle")
        status_line.addWidget(self.status_label, 1)
        self.details_button = QPushButton("查看详情")
        self.details_button.setObjectName("LinkButton")
        self.details_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.details_button.setVisible(False)
        self.details_button.clicked.connect(self.show_failures)
        status_line.addWidget(self.details_button)
        status_box.addLayout(status_line)
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(6)
        self.progress.setVisible(False)
        status_box.addWidget(self.progress)
        bottom_row.addLayout(status_box, 1)

        self.cancel_button = QPushButton("取消转换")
        self.cancel_button.setObjectName("DangerGhost")
        self.cancel_button.setFixedSize(92, 42)
        self.cancel_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.cancel_button.setVisible(False)
        self.cancel_button.clicked.connect(self.cancel_conversion)
        bottom_row.addWidget(self.cancel_button)

        self.start_button = QPushButton("开始转换")
        self.start_button.setObjectName("PrimaryButton")
        self.start_button.setFixedSize(148, 42)
        self.start_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.start_button.setIcon(line_icon("arrow-right", "#FFFFFF", 16, 1.9))
        self.start_button.setIconSize(QSize(16, 16))
        self.start_button.clicked.connect(self.start_conversion)
        bottom_row.addWidget(self.start_button)
        action_layout.addLayout(bottom_row)
        outer.addWidget(action_card)

        footer = QLabel(
            "本机自动选择 WPS/Word 或内置 LibreOffice 排版 · 文件不会上传"
            if self.mode == "word"
            else "文件仅在本机处理 · 扫描版 PDF 将按页面图片还原，不进行文字识别（OCR）"
        )
        footer.setObjectName("FooterHint")
        outer.addWidget(footer)

        self._sync_output_controls()
        self._install_shortcuts()

    def _build_empty_page(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet("background: #FFFFFF;")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 10, 0, 10)
        layout.setSpacing(4)
        layout.addStretch(1)
        icon = QLabel()
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setPixmap(line_icon("file", "#C4D1E3", 26, 1.5).pixmap(26, 26))
        layout.addWidget(icon)
        title = QLabel("队列还是空的")
        title.setObjectName("EmptyTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        hint = QLabel("拖入文件，或点击右上角的「添加文件」")
        hint.setObjectName("EmptyHint")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(hint)
        layout.addStretch(1)
        return page

    def _build_table(self) -> QTableWidget:
        table = QTableWidget(0, 4)
        table.setHorizontalHeaderLabels(["文件", "大小", "状态", ""])
        table.setShowGrid(False)
        table.setFrameShape(QFrame.Shape.NoFrame)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(54)
        table.horizontalHeader().setHighlightSections(False)
        table.horizontalHeader().setStretchLastSection(False)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        table.setColumnWidth(1, 78)
        table.setColumnWidth(2, 92)
        table.setColumnWidth(3, 42)
        table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        table.setMinimumHeight(140)
        for column in range(4):
            item = table.horizontalHeaderItem(column)
            if item is not None:
                alignment = (
                    Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
                    if column in (0, 1)
                    else Qt.AlignmentFlag.AlignCenter
                )
                item.setTextAlignment(alignment)
        table.cellDoubleClicked.connect(self._on_table_double_clicked)
        return table

    def _install_shortcuts(self) -> None:
        self._open_shortcut = QShortcut(QKeySequence("Ctrl+O"), self)
        self._open_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self._open_shortcut.activated.connect(self.browse_files)
        self._convert_shortcut = QShortcut(QKeySequence("Ctrl+Return"), self)
        self._convert_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self._convert_shortcut.activated.connect(self.start_conversion)

    def _restore_size_preferences(self) -> None:
        if not self.custom_dir:
            return
        self.custom_radio.setChecked(True)
        self._sync_output_controls()

    # -- adding and removing files -------------------------------------------

    def browse_files(self) -> None:
        if self._busy:
            return
        if self.mode == "word":
            filter_text = "Word 文档 (*.docx *.doc);;所有文件 (*.*)"
        else:
            filter_text = "PDF 文件 (*.pdf);;所有文件 (*.*)"
        start_dir = str(self.settings.value("lastInputDir", "") or "")
        if not start_dir or not Path(start_dir).is_dir():
            if self.records:
                start_dir = str(self.records[0].path.parent)
            else:
                start_dir = str(Path.home() / "Documents")
        paths, _ = QFileDialog.getOpenFileNames(self, "选择要转换的文件", start_dir, filter_text)
        if paths:
            self.settings.setValue("lastInputDir", str(Path(paths[0]).parent))
            self.add_files(paths)

    def add_files(self, paths: list[str]) -> None:
        if self._busy:
            return
        allowed = WORD_EXTENSIONS if self.mode == "word" else PDF_EXTENSIONS
        existing = {os.path.normcase(os.path.abspath(str(record.path))) for record in self.records}
        added = 0
        duplicates = 0
        invalid = 0
        first_invalid = ""
        for raw_path in paths:
            try:
                path = Path(raw_path).resolve()
            except OSError:
                invalid += 1
                continue
            if not path.is_file() or path.suffix.casefold() not in allowed:
                invalid += 1
                if not first_invalid:
                    first_invalid = path.name
                continue
            key = os.path.normcase(os.path.abspath(str(path)))
            if key in existing:
                duplicates += 1
                continue
            try:
                size = path.stat().st_size
            except OSError:
                invalid += 1
                continue
            record = FileRecord(record_id=uuid.uuid4().hex, path=path, size=size)
            self.records.append(record)
            self._insert_row(record)
            existing.add(key)
            added += 1

        if added:
            self.settings.setValue("lastInputDir", str(self.records[-1].path.parent))
            self._current_record_id = ""
            self._completion_reset()
            self._set_status(
                f"已添加 {added} 个文件，点击「开始转换」或按 Ctrl+Enter"
                if self.mode == "word"
                else f"已添加 {added} 个文件，点击「开始转换」或按 Ctrl+Enter",
                "idle",
            )
            self._update_queue_visibility()
            self._refresh_action_state()
            self._update_count_badge()
            self.table.scrollToBottom()
            QTimer.singleShot(0, self._refresh_cell_elisions)
        elif duplicates and not invalid:
            self._set_status("这些文件已经在转换队列中。", "idle")

        if invalid:
            names = f"（例如 {first_invalid}）" if first_invalid else ""
            if self.mode == "word":
                expected = "DOCX 或 DOC"
            else:
                expected = "PDF"
            if added:
                self._set_status(f"已忽略 {invalid} 个不支持的文件{names}", "error")
            else:
                box = QMessageBox(self)
                box.setWindowTitle("不支持的文件格式")
                box.setIcon(QMessageBox.Icon.Warning)
                box.setText(f"当前模式只能添加 {expected} 文件。")
                box.setInformativeText(
                    "若要转换另一种格式，请先切换到左侧对应的转换模式。"
                )
                box.setStandardButtons(QMessageBox.StandardButton.Ok)
                box.exec()
            return
        if duplicates and added:
            self._set_status(
                f"已添加 {added} 个文件，跳过 {duplicates} 个重复文件", "idle"
            )

    def _insert_row(self, record: FileRecord) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)
        cell = FileCell(self.mode, record.path.name, str(record.path))
        self.table.setCellWidget(row, 0, cell)

        size_label = QLabel(self._format_size(record.size))
        size_label.setObjectName("FileSize")
        size_label.setContentsMargins(10, 0, 0, 0)
        size_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        size_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.table.setCellWidget(row, 1, size_label)

        pill = StatusPill("等待转换", "idle")
        pill_wrap = QWidget()
        pill_wrap.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        pill_layout = QHBoxLayout(pill_wrap)
        pill_layout.setContentsMargins(2, 0, 2, 0)
        pill_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pill_layout.addWidget(pill)
        self.table.setCellWidget(row, 2, pill_wrap)

        remove = RowRemoveButton()
        remove.setProperty("recordId", record.record_id)
        remove.clicked.connect(lambda _checked=False, button=remove: self._remove_by_button(button))
        remove_wrap = QWidget()
        remove_layout = QHBoxLayout(remove_wrap)
        remove_layout.setContentsMargins(0, 0, 0, 0)
        remove_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        remove_layout.addWidget(remove)
        self.table.setCellWidget(row, 3, remove_wrap)
        self.table.setRowHeight(row, 54)

    def _refresh_cell_elisions(self) -> None:
        for row in range(self.table.rowCount()):
            cell = self.table.cellWidget(row, 0)
            if isinstance(cell, FileCell):
                cell.refresh_elision()

    def _remove_by_button(self, button: QToolButton) -> None:
        if self._busy:
            return
        record_id = str(button.property("recordId") or "")
        for index, record in enumerate(self.records):
            if record.record_id == record_id:
                self.records.pop(index)
                self.table.removeRow(index)
                break
        self._completion_reset()
        self._update_queue_visibility()
        self._update_count_badge()
        self._refresh_action_state()
        if self.records:
            self._set_status("文件已移除，队列已更新。", "idle")
        else:
            self._set_status("添加文件后即可开始转换", "idle")

    def clear_queue(self) -> None:
        if self._busy or not self.records:
            return
        self.records.clear()
        self.table.setRowCount(0)
        self._update_queue_visibility()
        self._update_count_badge()
        self._completion_reset()
        self._refresh_action_state()
        self._set_status("添加文件后即可开始转换", "idle")

    def _update_queue_visibility(self) -> None:
        self.queue_stack.setCurrentWidget(self.table if self.records else self.empty_page)
        self.clear_button.setEnabled(bool(self.records) and not self._busy)
        self.add_button.setEnabled(not self._busy)

    def _update_count_badge(self) -> None:
        self.count_badge.setText(str(len(self.records)))

    def _format_size(self, size: int) -> str:
        if size < 1024:
            return f"{size} B"
        if size < 1024 * 1024:
            return f"{size / 1024:.0f} KB"
        return f"{size / (1024 * 1024):.1f} MB"

    # -- output settings ------------------------------------------------------

    def _on_pdf_style_changed(self, checked: bool) -> None:
        if not checked or not self.flow_radio or not self.layout_radio:
            return
        mode = "layout" if self.layout_radio.isChecked() else "flow"
        self.settings.setValue("pdfMode", mode)

    def _selected_pdf_mode(self) -> str:
        if self.mode == "pdf" and self.layout_radio is not None and self.layout_radio.isChecked():
            return "layout"
        return "flow"

    def _sync_output_controls(self) -> None:
        custom = self.custom_radio.isChecked()
        self.output_path.setEnabled(custom and not self._busy)
        if custom:
            self.output_path.setText(self.custom_dir or "请选择输出文件夹…")
            self.output_path.setToolTip(self.custom_dir)
            self.output_path.setCursorPosition(0)
        else:
            self.output_path.setText("跟随每个源文件所在位置")
            self.output_path.setToolTip("")
            self.output_path.setCursorPosition(0)

    def choose_output_dir(self) -> bool:
        if self._busy:
            return False
        start = self.custom_dir
        if not start or not Path(start).is_dir():
            start = str(self.settings.value("lastOutputDir", "") or "")
        if not start or not Path(start).is_dir():
            start = str(Path.home() / "Documents")
        directory = QFileDialog.getExistingDirectory(self, "选择输出文件夹", start)
        if not directory:
            return False
        self.custom_dir = directory
        self.settings.setValue("customOutputDir", directory)
        self.settings.setValue("lastOutputDir", directory)
        self.custom_radio.setChecked(True)
        self._sync_output_controls()
        return True

    def _output_dir_for(self, record: FileRecord) -> Path:
        if self.custom_radio.isChecked():
            return Path(self.custom_dir)
        return record.path.parent

    # -- action state ---------------------------------------------------------

    def _refresh_action_state(self) -> None:
        if self._busy:
            return
        pending = [r for r in self.records if r.status in ("waiting", "error", "cancel")]
        waiting = any(r.status == "waiting" for r in self.records)
        if not self.records:
            self.start_button.setEnabled(False)
            self.start_button.setText("开始转换")
            self.start_button.setIcon(line_icon("arrow-right", "#FFFFFF", 16, 1.9))
        elif pending:
            self.start_button.setEnabled(True)
            if not waiting and any(r.status == "error" for r in self.records):
                self.start_button.setText("重试失败项")
                self.start_button.setIcon(line_icon("refresh", "#FFFFFF", 16, 1.9))
            elif not waiting and any(r.status == "cancel" for r in self.records):
                self.start_button.setText("继续转换")
                self.start_button.setIcon(line_icon("refresh", "#FFFFFF", 16, 1.9))
            else:
                self.start_button.setText("开始转换")
                self.start_button.setIcon(line_icon("arrow-right", "#FFFFFF", 16, 1.9))
        else:
            self.start_button.setEnabled(False)
            self.start_button.setText("转换完成")
            self.start_button.setIcon(line_icon("check", "#FFFFFF", 16, 1.9))
        self.clear_button.setEnabled(bool(self.records) and not self._busy)
        self.add_button.setEnabled(not self._busy)

    def _set_status(self, text: str, state: str = "idle") -> None:
        self.status_label.setText(text)
        self.status_label.setToolTip(text)
        self.status_label.setProperty("state", state)
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)

    def _completion_reset(self) -> None:
        self.details_button.setVisible(False)
        self._batch_total = 0
        self._batch_success = 0
        self._batch_failed = 0
        self._batch_warnings = []

    # -- conversion -----------------------------------------------------------

    def start_conversion(self) -> None:
        if self._busy:
            return
        pending = [r for r in self.records if r.status in ("waiting", "error", "cancel")]
        if not pending:
            return
        if self.worker is not None and self.worker.isRunning():
            return

        if self.mode == "word":
            engine = self.host.get_word_engine()
            if not engine:
                self._offer_word_engine()
                engine = self.host.get_word_engine()
                if not engine:
                    return
        else:
            engine = None
            if not pdf_engine_available():
                QMessageBox.critical(
                    self,
                    "组件缺失",
                    "PDF 转 Word 组件未安装或已损坏。\n请重新安装应用后再试。",
                )
                return

        if self.custom_radio.isChecked() and not self.custom_dir:
            if not self.choose_output_dir():
                return

        for record in pending:
            record.status = "waiting"
            record.output = ""
            record.warning = ""
            record.error = ""
            record.detail = ""
            self._set_pill(record, "等待转换", "idle")
            self._set_row_tooltip(record, "")

        tasks = [
            ConversionTask(
                record_id=record.record_id,
                source=record.path,
                output_dir=self._output_dir_for(record),
            )
            for record in pending
        ]
        self._batch_total = len(tasks)
        self._batch_success = 0
        self._batch_failed = 0
        self._batch_warnings = []
        self._current_record_id = ""
        self.details_button.setVisible(False)
        self._set_busy(True)
        self._set_status(f"正在准备转换 {len(tasks)} 个文件…", "working")
        self.progress.setRange(0, 0)
        self.progress.setValue(0)
        self.progress.setVisible(True)

        worker = ConversionWorker(tasks, self.mode, engine, self._selected_pdf_mode(), self)
        self.worker = worker
        worker.file_started.connect(self._on_file_started)
        worker.file_done.connect(self._on_file_done)
        worker.all_finished.connect(self._on_all_finished)
        worker.finished.connect(lambda: self._retire_worker(worker))
        worker.start()

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        self.start_button.setVisible(not busy)
        self.cancel_button.setVisible(busy)
        self.cancel_button.setEnabled(busy)
        self.cancel_button.setText("取消转换")
        self.add_button.setEnabled(not busy)
        self.clear_button.setEnabled(not busy and bool(self.records))
        self.drop_area.browse_button.setEnabled(not busy)
        self.drop_area.setAcceptDrops(not busy)
        self.source_radio.setEnabled(not busy)
        self.custom_radio.setEnabled(not busy)
        if self.flow_radio is not None:
            self.flow_radio.setEnabled(not busy)
        if self.layout_radio is not None:
            self.layout_radio.setEnabled(not busy)
        self.browse_output_button.setEnabled(not busy)
        self._sync_output_controls()
        for row in range(self.table.rowCount()):
            wrapper = self.table.cellWidget(row, 3)
            if wrapper:
                button = wrapper.findChild(QToolButton)
                if button:
                    button.setEnabled(not busy)
        if busy:
            self.details_button.setVisible(False)
        self.conversionStateChanged.emit(busy)

    def cancel_conversion(self) -> None:
        if not self.worker or not self._busy:
            return
        self.cancel_button.setEnabled(False)
        self.cancel_button.setText("正在取消…")
        self._set_status("正在停止当前转换…", "working")
        self.worker.cancel()

    def _on_file_started(self, record_id: str, index: int, total: int) -> None:
        record = self._record_by_id(record_id)
        if not record:
            return
        self._current_record_id = record_id
        self._set_pill(record, "转换中", "working")
        self._set_status(f"正在转换 {index} / {total} · {record.path.name}", "working")
        self.progress.setRange(0, 0)
        for row, candidate in enumerate(self.records):
            if candidate.record_id == record_id:
                self.table.scrollTo(
                    self.table.model().index(row, 0),
                    QAbstractItemView.ScrollHint.PositionAtCenter,
                )
                break

    def _on_file_done(
        self,
        record_id: str,
        state: str,
        output: str,
        warning: str,
        error: str,
        detail: str,
    ) -> None:
        record = self._record_by_id(record_id)
        if not record:
            return
        record.output = output
        record.warning = warning
        record.error = error
        record.detail = detail
        if state == "success":
            record.status = "success"
            self._set_pill(record, "已完成", "success")
            tooltip = f"输出文件：{output}"
            if warning:
                tooltip += f"\n\n{warning}"
            self._set_row_tooltip(record, tooltip)
            self.last_output_file = output
            self.last_output_dir = str(Path(output).parent)
            if warning:
                self._batch_warnings.append(warning)
            self.settings.setValue("lastOutputDir", self.last_output_dir)
            self.open_folder_button.setEnabled(True)
            self.open_folder_button.setToolTip(output)
        elif state == "cancel":
            record.status = "cancel"
            self._set_pill(record, "已取消", "cancel")
        else:
            record.status = "error"
            self._set_pill(record, "失败", "error")
            self._set_row_tooltip(record, (error + ("\n\n" + detail if detail else "")).strip())
        self._update_count_badge()

    def _on_all_finished(
        self, success: int, failed: int, cancelled: bool, elapsed: float
    ) -> None:
        if cancelled:
            for record in self.records:
                if record.status in ("waiting", "working"):
                    record.status = "cancel"
                    self._set_pill(record, "已取消", "cancel")
        self._batch_success = success
        self._batch_failed = failed
        self._current_record_id = ""
        self._set_busy(False)
        self.progress.setVisible(False)
        self.progress.setRange(0, 1)
        self._update_count_badge()

        if cancelled:
            done = success + failed
            self._set_status(
                f"已取消 · 已完成 {done} / {self._batch_total} 个文件", "idle"
            )
        elif failed:
            self._set_status(
                f"转换结束：{success} 个成功，{failed} 个失败 · 用时 {_format_duration(elapsed)}",
                "error",
            )
            self.details_button.setVisible(True)
        else:
            if success == 1:
                message = f"转换完成 · 用时 {_format_duration(elapsed)}"
            else:
                message = f"全部 {success} 个文件转换完成 · 用时 {_format_duration(elapsed)}"
            if self._batch_warnings:
                scanned_warnings = sum(
                    1 for warning in self._batch_warnings if "扫描" in warning or "文字层" in warning
                )
                if scanned_warnings:
                    message += f" · {scanned_warnings} 个扫描版 PDF 已按图片还原"
                other_warnings = len(self._batch_warnings) - scanned_warnings
                if other_warnings:
                    message += f" · {other_warnings} 个文件有转换提示"
            self._set_status(message, "success")
            self.details_button.setVisible(False)
        self._refresh_action_state()

    def _retire_worker(self, worker: ConversionWorker) -> None:
        if worker in self._retired_workers:
            self._retired_workers.remove(worker)
        if self.worker is worker:
            self.worker = None
        worker.deleteLater()

    def _record_by_id(self, record_id: str) -> FileRecord | None:
        for record in self.records:
            if record.record_id == record_id:
                return record
        return None

    def _set_pill(self, record: FileRecord, text: str, state: str) -> None:
        for index, candidate in enumerate(self.records):
            if candidate.record_id == record.record_id:
                wrapper = self.table.cellWidget(index, 2)
                if wrapper:
                    pill = wrapper.findChild(StatusPill)
                    if pill:
                        pill.set_state(state, text)
                break

    def _set_row_tooltip(self, record: FileRecord, tooltip: str) -> None:
        for index, candidate in enumerate(self.records):
            if candidate.record_id == record.record_id:
                wrapper = self.table.cellWidget(index, 2)
                if wrapper:
                    pill = wrapper.findChild(StatusPill)
                    if pill:
                        pill.setToolTip(tooltip)
                cell = self.table.cellWidget(index, 0)
                if cell:
                    cell.setToolTip(tooltip or str(record.path))
                break

    # -- output / errors ------------------------------------------------------

    def open_output_folder(self) -> None:
        target = self.last_output_file or self.last_output_dir
        if target and Path(target).exists():
            path = Path(target)
            folder = path if path.is_dir() else path.parent
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))
        else:
            QMessageBox.information(
                self, "暂无输出文件夹", "完成一次转换后，就可以在这里打开输出位置。"
            )

    def show_failures(self) -> None:
        failures = [record for record in self.records if record.status == "error"]
        if not failures:
            return
        chunks: list[str] = []
        for record in failures:
            chunks.append(f"文件：{record.path}")
            chunks.append(f"错误：{record.error or '未知错误'}")
            if record.detail:
                chunks.append(record.detail)
            chunks.append("-" * 68)
        subtitle = (
            f"{len(failures)} 个文件未能转换。可以修复问题后点击「重试失败项」。"
        )
        dialog = ErrorDetailsDialog(self, "转换失败详情", subtitle, "\n".join(chunks))
        dialog.exec()

    def _on_table_double_clicked(self, row: int, _column: int) -> None:
        if row < 0 or row >= len(self.records):
            return
        record = self.records[row]
        if record.status == "success" and record.output and Path(record.output).exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(record.output))
        elif record.status == "error":
            self.show_failures()

    def _offer_word_engine(self) -> None:
        box = QMessageBox(self)
        box.setWindowTitle("需要 Word 转换引擎")
        box.setIcon(QMessageBox.Icon.Warning)
        box.setText("没有找到可用的 Word 转 PDF 引擎。")
        box.setInformativeText(
            "可以安装 WPS Office、Microsoft Word 或 LibreOffice 后重启应用；"
            "也可以点击下方按钮指定已有的 soffice.exe。"
        )
        choose = box.addButton("选择 soffice.exe", QMessageBox.ButtonRole.ActionRole)
        box.addButton("取消", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() is choose:
            path, _ = QFileDialog.getOpenFileName(
                self, "选择 soffice.exe", "", "LibreOffice (*.exe);;所有文件 (*.*)"
            )
            if path:
                self.host.set_soffice_path(path)
                if not self.host.get_word_engine():
                    QMessageBox.warning(
                        self, "引擎不可用", "无法使用所选程序，请选择 LibreOffice 安装目录下的 soffice.exe。"
                    )

    def is_busy(self) -> bool:
        return self._busy or (self.worker is not None and self.worker.isRunning())

    def cancel_and_wait(self, timeout_ms: int = 8000) -> bool:
        if self.worker is not None:
            self.worker.cancel()
            return self.worker.wait(timeout_ms)
        return True


def _format_duration(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f} 秒"
    minutes = int(seconds // 60)
    remainder = int(round(seconds % 60))
    if minutes >= 60:
        hours = minutes // 60
        minutes = minutes % 60
        return f"{hours} 小时 {minutes} 分"
    return f"{minutes} 分 {remainder} 秒"


# ----------------------------------------------------------------------------
# Main window
# ----------------------------------------------------------------------------


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        APP_DATA.mkdir(parents=True, exist_ok=True)
        self.settings = QSettings(str(SETTINGS_PATH), QSettings.Format.IniFormat)
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(QIcon(app_icon_path()))
        self.setMinimumSize(980, 660)
        self._build_ui()
        self.refresh_engines()
        self._restore_window_state()

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("Root")
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(228)
        side_layout = QVBoxLayout(sidebar)
        side_layout.setContentsMargins(20, 22, 20, 18)
        side_layout.setSpacing(0)

        brand_row = QHBoxLayout()
        brand_row.setSpacing(10)
        logo = QLabel()
        logo.setPixmap(app_logo(40))
        logo.setFixedSize(40, 40)
        brand_row.addWidget(logo)
        brand_text = QVBoxLayout()
        brand_text.setSpacing(0)
        brand_title = QLabel("文枢")
        brand_title.setObjectName("BrandTitle")
        brand_sub = QLabel("文档转换器")
        brand_sub.setObjectName("BrandSub")
        brand_text.addWidget(brand_title)
        brand_text.addWidget(brand_sub)
        brand_row.addLayout(brand_text)
        brand_row.addStretch(1)
        side_layout.addLayout(brand_row)
        side_layout.addSpacing(36)

        section = QLabel("转换模式")
        section.setObjectName("SidebarSection")
        side_layout.addWidget(section)
        side_layout.addSpacing(8)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        self.nav_buttons: list[QPushButton] = []

        nav_word = QPushButton("  Word 转 PDF")
        nav_word.setObjectName("NavButton")
        nav_word.setCheckable(True)
        nav_word.setFixedHeight(46)
        nav_word.setIcon(QIcon(badge_pixmap("word", 20)))
        nav_word.setIconSize(QSize(20, 20))
        nav_word.setCursor(Qt.CursorShape.PointingHandCursor)
        self.nav_group.addButton(nav_word, 0)
        self.nav_buttons.append(nav_word)
        side_layout.addWidget(nav_word)
        side_layout.addSpacing(4)

        nav_pdf = QPushButton("  PDF 转 Word")
        nav_pdf.setObjectName("NavButton")
        nav_pdf.setCheckable(True)
        nav_pdf.setFixedHeight(46)
        nav_pdf.setIcon(QIcon(badge_pixmap("pdf", 20)))
        nav_pdf.setIconSize(QSize(20, 20))
        nav_pdf.setCursor(Qt.CursorShape.PointingHandCursor)
        self.nav_group.addButton(nav_pdf, 1)
        self.nav_buttons.append(nav_pdf)
        side_layout.addWidget(nav_pdf)
        side_layout.addStretch(1)

        engine_card = QFrame()
        engine_card.setObjectName("EngineCard")
        engine_layout = QVBoxLayout(engine_card)
        engine_layout.setContentsMargins(12, 10, 12, 10)
        engine_layout.setSpacing(6)
        engine_title = QLabel("本地引擎")
        engine_title.setObjectName("EngineTitle")
        engine_layout.addWidget(engine_title)

        word_engine_row = QHBoxLayout()
        word_engine_row.setSpacing(6)
        self.engine_dot_word = QLabel("●")
        self.engine_dot_word.setObjectName("EngineDotOff")
        self.engine_dot_word.setFixedWidth(10)
        self.engine_line_word = QLabel("Word → PDF")
        self.engine_line_word.setObjectName("EngineLine")
        word_engine_row.addWidget(self.engine_dot_word)
        word_engine_row.addWidget(self.engine_line_word, 1)
        engine_layout.addLayout(word_engine_row)

        pdf_engine_row = QHBoxLayout()
        pdf_engine_row.setSpacing(6)
        self.engine_dot_pdf = QLabel("●")
        self.engine_dot_pdf.setObjectName("EngineDotOff")
        self.engine_dot_pdf.setFixedWidth(10)
        self.engine_line_pdf = QLabel("PDF → Word")
        self.engine_line_pdf.setObjectName("EngineLine")
        pdf_engine_row.addWidget(self.engine_dot_pdf)
        pdf_engine_row.addWidget(self.engine_line_pdf, 1)
        engine_layout.addLayout(pdf_engine_row)
        side_layout.addWidget(engine_card)
        side_layout.addSpacing(10)

        privacy_row = QHBoxLayout()
        privacy_row.setSpacing(5)
        shield = QLabel()
        shield.setPixmap(line_icon("shield", "#5F7794", 12, 1.5).pixmap(12, 12))
        shield.setFixedSize(12, 12)
        privacy = QLabel("文件仅在本机处理")
        privacy.setObjectName("PrivacyLine")
        privacy_row.addWidget(shield)
        privacy_row.addWidget(privacy)
        privacy_row.addStretch(1)
        side_layout.addLayout(privacy_row)
        side_layout.addSpacing(5)
        version = QLabel(f"v{VERSION}")
        version.setObjectName("VersionLabel")
        side_layout.addWidget(version)

        layout.addWidget(sidebar)

        self.stack = QStackedWidget()
        self.pages = [ConverterPage("word", self), ConverterPage("pdf", self)]
        for page in self.pages:
            self.stack.addWidget(page)
            page.conversionStateChanged.connect(self._on_conversion_state_changed)
        layout.addWidget(self.stack, 1)

        self.nav_group.idClicked.connect(self._switch_mode)
        self.nav_buttons[0].setChecked(True)

    def _switch_mode(self, index: int) -> None:
        if 0 <= index < self.stack.count():
            self.stack.setCurrentIndex(index)
            self.nav_buttons[index].setChecked(True)

    def _restore_window_state(self) -> None:
        geometry = self.settings.value("windowGeometry", QByteArray())
        restored = False
        if isinstance(geometry, QByteArray) and not geometry.isEmpty():
            restored = self.restoreGeometry(geometry)
        if not restored:
            screen = QGuiApplication.primaryScreen()
            available = screen.availableGeometry() if screen else None
            width = min(1200, max(1000, available.width() - 120)) if available else 1180
            height = min(800, max(660, available.height() - 90)) if available else 780
            self.resize(width, height)
            if available:
                frame = self.frameGeometry()
                frame.moveCenter(available.center())
                self.move(frame.topLeft())
        mode = int(self.settings.value("activeMode", 0) or 0)
        if mode in (0, 1):
            self.nav_buttons[mode].setChecked(True)
            self.stack.setCurrentIndex(mode)
        screen = QGuiApplication.screenAt(self.frameGeometry().center())
        if screen is None:
            screen = QGuiApplication.primaryScreen()
        if screen and not screen.availableGeometry().intersects(self.frameGeometry()):
            available = screen.availableGeometry()
            self.resize(min(1180, available.width() - 80), min(780, available.height() - 80))
            frame = self.frameGeometry()
            frame.moveCenter(available.center())
            self.move(frame.topLeft())

    def get_word_engine(self) -> dict[str, str] | None:
        preferred = str(self.settings.value("sofficePath", "") or "")
        return find_word_engine(preferred or None)

    def set_soffice_path(self, path: str) -> None:
        self.settings.setValue("sofficePath", path)
        self.refresh_engines()

    def refresh_engines(self) -> None:
        engine = self.get_word_engine()
        word_ready = engine is not None
        pdf_ready = pdf_engine_available()
        self.engine_dot_word.setObjectName("EngineDotOn" if word_ready else "EngineDotOff")
        self.engine_dot_pdf.setObjectName("EngineDotOn" if pdf_ready else "EngineDotOff")
        self.engine_line_word.setText("Word → PDF  已就绪" if word_ready else "Word → PDF  未就绪")
        self.engine_line_pdf.setText("PDF → Word  已就绪" if pdf_ready else "PDF → Word  未就绪")
        for widget in (self.engine_dot_word, self.engine_dot_pdf):
            widget.style().unpolish(widget)
            widget.style().polish(widget)
        if engine:
            detail = (
                engine.get("path")
                or engine.get("cli")
                or engine.get("label")
                or ""
            )
            self.engine_line_word.setToolTip(f"{format_engine_summary(engine)}\n{detail}")
        else:
            self.engine_line_word.setToolTip("安装 LibreOffice 或指定 soffice.exe 后可用。")
        self.engine_line_pdf.setToolTip(
            "内置 PDF 解析引擎已就绪。" if pdf_ready else "PDF 转 Word 组件不可用。"
        )

    def _on_conversion_state_changed(self, busy: bool) -> None:
        for button in self.nav_buttons:
            button.setEnabled(not busy)
        self.setWindowTitle(f"{APP_TITLE} · 正在转换…" if busy else APP_TITLE)

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        busy_pages = [page for page in self.pages if page.is_busy()]
        if busy_pages:
            answer = QMessageBox.question(
                self,
                "正在转换",
                "还有文件正在转换。退出将中断当前任务，确定要退出吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            for page in busy_pages:
                if not page.cancel_and_wait(10000):
                    QMessageBox.warning(
                        self,
                        "仍在停止",
                        "转换进程尚未完全停止，请稍等几秒后再关闭窗口。",
                    )
                    event.ignore()
                    return
        self.settings.setValue("windowGeometry", self.saveGeometry())
        self.settings.setValue("activeMode", self.stack.currentIndex())
        self.settings.sync()
        event.accept()


# ----------------------------------------------------------------------------
# Application bootstrap
# ----------------------------------------------------------------------------


def _cleanup_stale_temp() -> None:
    """Remove temporary job folders left by a previous crash."""
    temp_root = Path(tempfile.gettempdir())
    cutoff = time.time() - 24 * 60 * 60
    for pattern in ("wenshu-job-*", "wenshu-*"):
        for path in temp_root.glob(pattern):
            try:
                if path.is_dir() and path.stat().st_mtime < cutoff:
                    shutil.rmtree(path, ignore_errors=True)
            except OSError:
                pass


def _setup_logging() -> None:
    log_dir = APP_DATA / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    handler = logging.handlers.RotatingFileHandler(
        log_dir / "app.log", maxBytes=2_000_000, backupCount=2, encoding="utf-8"
    )
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(handler)


def _install_exception_hook() -> None:
    def handle(exc_type, exc_value, exc_tb):
        logging.getLogger("wenshu").exception(
            "Unhandled exception", exc_info=(exc_type, exc_value, exc_tb)
        )
        try:
            QMessageBox.critical(
                None,
                "出现意外错误",
                f"应用遇到了意外错误：\n{exc_type.__name__}: {exc_value}\n\n"
                f"详细信息已写入：{APP_DATA / 'logs' / 'app.log'}",
            )
        except Exception:
            pass

    sys.excepthook = handle


def _set_windows_app_id() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "WenShu.DocumentConverter.1.0"
        )
    except Exception:
        pass


def run_app() -> int:
    _setup_logging()
    _cleanup_stale_temp()
    _set_windows_app_id()
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
    app.setOrganizationName("WenShu")
    app.setApplicationName("WenShuConverter")
    app.setApplicationDisplayName(APP_TITLE)
    if "Fusion" in QStyleFactory.keys():
        app.setStyle("Fusion")
    app.setFont(QFont("Microsoft YaHei UI", 9))
    app.setStyleSheet(APP_QSS)

    icon_path = app_icon_path()
    icon = QIcon(icon_path)
    if icon.isNull():
        icon = QIcon(app_logo(128))
    app.setWindowIcon(icon)
    _install_exception_hook()

    window = MainWindow()
    window.show()
    return app.exec()
