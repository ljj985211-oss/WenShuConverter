"""Reusable widgets for the converter interface."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from icons import badge_pixmap, line_icon


class ElidedLabel(QLabel):
    """A QLabel that elides its middle while keeping the full text as tooltip."""

    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._full_text = ""
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.setMinimumWidth(0)
        self.setText(text)

    def setText(self, text: str) -> None:  # noqa: N802 - Qt API name
        self._full_text = text or ""
        self.setToolTip(self._full_text)
        self._refresh_elision()

    def fullText(self) -> str:  # noqa: N802 - Qt-style name
        return self._full_text

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        self._refresh_elision()

    def _refresh_elision(self) -> None:
        available = max(0, self.width() - 2)
        if available <= 0:
            super().setText(self._full_text)
            return
        elided = self.fontMetrics().elidedText(
            self._full_text, Qt.TextElideMode.ElideMiddle, available
        )
        super().setText(elided)


class StatusPill(QLabel):
    def __init__(self, text: str = "等待转换", state: str = "idle") -> None:
        super().__init__(text)
        self.setObjectName("StatusPill")
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setFixedHeight(22)
        self.setMinimumWidth(58)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.set_state(state, text)

    def set_state(self, state: str, text: str) -> None:
        self.setText(text)
        self.setProperty("state", state)
        self.style().unpolish(self)
        self.style().polish(self)
        self.adjustSize()
        self.setFixedHeight(22)
        if text in ("等待转换", "待转换"):
            self.setMinimumWidth(58)
        else:
            self.setMinimumWidth(max(58, self.fontMetrics().horizontalAdvance(text) + 18))


class DropArea(QFrame):
    filesDropped = Signal(list)
    browseRequested = Signal()

    def __init__(self, mode: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.mode = mode
        self.setObjectName("DropArea")
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedHeight(96)
        self.setProperty("dragActive", False)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(14)

        icon_circle = QLabel()
        icon_circle.setObjectName("DropIcon")
        icon_circle.setFixedSize(42, 42)
        icon_circle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_circle.setPixmap(line_icon("download", "#2563EB", 20, 1.8).pixmap(20, 20))
        icon_circle.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(icon_circle)

        text_column = QVBoxLayout()
        text_column.setContentsMargins(0, 0, 0, 0)
        text_column.setSpacing(3)
        self.title = QLabel()
        self.title.setObjectName("DropTitle")
        self.hint = QLabel()
        self.hint.setObjectName("DropHint")
        self.title.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.hint.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        text_column.addWidget(self.title)
        text_column.addWidget(self.hint)
        layout.addLayout(text_column, 1)

        self.browse_button = QPushButton("选择文件")
        self.browse_button.setObjectName("SecondaryButton")
        self.browse_button.setFixedSize(82, 32)
        self.browse_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.browse_button.clicked.connect(self.browseRequested.emit)
        layout.addWidget(self.browse_button)
        self.set_mode(mode)

    def set_mode(self, mode: str) -> None:
        self.mode = mode
        if mode == "word":
            self.title.setText("拖放 Word 文档到这里，或点击选择")
            self.hint.setText("支持 DOCX、DOC 格式 · 可一次添加多个文件")
        else:
            self.title.setText("拖放 PDF 文件到这里，或点击选择")
            self.hint.setText("支持 PDF 格式 · 可一次添加多个文件")

    def _set_active(self, active: bool) -> None:
        self.setProperty("dragActive", active)
        self.style().unpolish(self)
        self.style().polish(self)
        self.update()

    def mousePressEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self.browseRequested.emit()
            event.accept()
            return
        super().mousePressEvent(event)

    def dragEnterEvent(self, event) -> None:  # noqa: N802
        if self._local_files(event.mimeData()):
            event.acceptProposedAction()
            self._set_active(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event) -> None:  # noqa: N802
        if self._local_files(event.mimeData()):
            event.acceptProposedAction()

    def dragLeaveEvent(self, event) -> None:  # noqa: N802
        self._set_active(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event) -> None:  # noqa: N802
        self._set_active(False)
        files = self._local_files(event.mimeData())
        if files:
            self.filesDropped.emit(files)
            event.acceptProposedAction()
        else:
            event.ignore()

    @staticmethod
    def _local_files(mime) -> list[str]:
        if mime is None or not mime.hasUrls():
            return []
        result: list[str] = []
        for url in mime.urls():
            if url.isLocalFile():
                path = url.toLocalFile()
                if path:
                    result.append(path)
        return result


class FileCell(QWidget):
    def __init__(self, mode: str, file_name: str, file_path: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 0, 8, 0)
        layout.setSpacing(10)

        badge = QLabel()
        badge.setFixedSize(34, 34)
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge.setPixmap(badge_pixmap("pdf" if mode == "pdf" else "word", 34))
        badge.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(badge)

        text_column = QVBoxLayout()
        text_column.setContentsMargins(0, 0, 0, 0)
        text_column.setSpacing(2)
        name = ElidedLabel(file_name)
        name.setObjectName("FileName")
        name.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        path = ElidedLabel(shorten_path(file_path))
        path.setObjectName("FilePath")
        path.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        path.setToolTip(file_path)
        text_column.addWidget(name)
        text_column.addWidget(path)
        layout.addLayout(text_column, 1)

        self.name_label = name
        self.path_label = path
        self.setToolTip(str(file_path))

    def refresh_elision(self) -> None:
        self.name_label._refresh_elision()
        self.path_label._refresh_elision()


class RowRemoveButton(QToolButton):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("RowAction")
        self.setFixedSize(28, 28)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setIcon(line_icon("trash", "#99A7B9", 16, 1.6))
        self.setIconSize(QSize(16, 16))
        self.setToolTip("从队列中移除")


class ErrorDetailsDialog(QDialog):
    def __init__(
        self,
        parent: QWidget | None,
        title: str,
        subtitle: str,
        details: str,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumSize(620, 390)
        self.resize(680, 430)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(26, 24, 26, 20)
        layout.setSpacing(12)

        header = QHBoxLayout()
        header.setSpacing(12)
        icon = QLabel()
        icon.setPixmap(line_icon("alert", "#C95151", 24, 1.7).pixmap(24, 24))
        icon.setFixedSize(28, 28)
        header.addWidget(icon, 0, Qt.AlignmentFlag.AlignTop)
        text_box = QVBoxLayout()
        text_box.setSpacing(3)
        title_label = QLabel(title)
        title_label.setObjectName("DialogTitle")
        subtitle_label = QLabel(subtitle)
        subtitle_label.setObjectName("DialogSubtitle")
        subtitle_label.setWordWrap(True)
        text_box.addWidget(title_label)
        text_box.addWidget(subtitle_label)
        header.addLayout(text_box, 1)
        layout.addLayout(header)

        self.details = QTextEdit()
        self.details.setObjectName("DetailsView")
        self.details.setReadOnly(True)
        self.details.setPlainText(details)
        layout.addWidget(self.details, 1)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        copy_button = QPushButton("复制错误信息")
        copy_button.setObjectName("SecondaryButton")
        copy_button.setFixedHeight(32)
        copy_button.clicked.connect(lambda: QApplication.clipboard().setText(self.details.toPlainText()))
        close_button = QPushButton("关闭")
        close_button.setObjectName("PrimaryButton")
        close_button.setFixedHeight(32)
        close_button.clicked.connect(self.accept)
        actions.addWidget(copy_button)
        actions.addStretch(1)
        actions.addWidget(close_button)
        layout.addLayout(actions)


def shorten_path(path: str, max_chars: int = 90) -> str:
    text = str(path)
    if len(text) <= max_chars:
        return text
    parent = str(Path(text).parent)
    name = Path(text).name
    keep = max(20, max_chars - len(name) - 5)
    if len(parent) <= keep:
        return text
    return parent[: max(4, keep // 2)] + "…" + parent[-(keep // 2) :] + "\\" + name
