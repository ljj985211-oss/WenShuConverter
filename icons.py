"""Small painted line icons and file badges, with no external asset dependency."""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import (
    QColor,
    QFont,
    QIcon,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
)


def _new_canvas(size: int, scale: int = 2) -> tuple[QPixmap, QPainter]:
    pixmap = QPixmap(size * scale, size * scale)
    pixmap.setDevicePixelRatio(scale)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)
    return pixmap, painter


def line_icon(
    name: str,
    color: str = "#50637C",
    size: int = 20,
    stroke: float = 1.8,
) -> QIcon:
    """Paint a 24x24-grid line icon at the requested logical size."""
    pixmap, painter = _new_canvas(size)
    painter.scale(size / 24.0, size / 24.0)
    pen = QPen(QColor(color), stroke, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.BrushStyle.NoBrush)

    if name == "plus":
        painter.drawLine(QPointF(12, 5), QPointF(12, 19))
        painter.drawLine(QPointF(5, 12), QPointF(19, 12))
    elif name == "folder":
        path = QPainterPath()
        path.moveTo(3, 7.5)
        path.quadTo(3, 6, 4.5, 6)
        path.lineTo(9.1, 6)
        path.lineTo(11.3, 8.4)
        path.lineTo(19.4, 8.4)
        path.quadTo(21, 8.4, 21, 10)
        path.lineTo(21, 18.2)
        path.quadTo(21, 19.8, 19.4, 19.8)
        path.lineTo(4.6, 19.8)
        path.quadTo(3, 19.8, 3, 18.2)
        path.closeSubpath()
        painter.drawPath(path)
    elif name == "folder-open":
        path = QPainterPath()
        path.moveTo(3, 8)
        path.quadTo(3, 6.5, 4.6, 6.5)
        path.lineTo(9, 6.5)
        path.lineTo(11.2, 8.8)
        path.lineTo(18.6, 8.8)
        path.quadTo(20, 8.8, 20, 10.2)
        path.lineTo(20, 11.2)
        painter.drawPath(path)
        path2 = QPainterPath()
        path2.moveTo(3.2, 11.2)
        path2.lineTo(19.2, 11.2)
        path2.quadTo(20.7, 11.2, 20.1, 12.6)
        path2.lineTo(17.7, 19)
        path2.quadTo(17.3, 20, 16.1, 20)
        path2.lineTo(4.3, 20)
        path2.quadTo(3, 20, 3, 18.6)
        path2.lineTo(3.2, 11.2)
        path2.closeSubpath()
        painter.drawPath(path2)
    elif name == "trash":
        painter.drawLine(QPointF(4, 7), QPointF(20, 7))
        painter.drawLine(QPointF(9, 7), QPointF(9.4, 4.4))
        painter.drawLine(QPointF(9.4, 4.4), QPointF(14.6, 4.4))
        painter.drawLine(QPointF(14.6, 4.4), QPointF(15, 7))
        path = QPainterPath()
        path.moveTo(6.2, 7)
        path.lineTo(7.1, 19.2)
        path.quadTo(7.2, 20.3, 8.3, 20.3)
        path.lineTo(15.7, 20.3)
        path.quadTo(16.8, 20.3, 16.9, 19.2)
        path.lineTo(17.8, 7)
        painter.drawPath(path)
        painter.drawLine(QPointF(10.2, 10.4), QPointF(10.6, 17))
        painter.drawLine(QPointF(13.8, 10.4), QPointF(13.4, 17))
    elif name == "arrow-right":
        painter.drawLine(QPointF(4, 12), QPointF(19, 12))
        painter.drawLine(QPointF(13.5, 6.5), QPointF(19, 12))
        painter.drawLine(QPointF(13.5, 17.5), QPointF(19, 12))
    elif name == "download":
        painter.drawLine(QPointF(12, 4), QPointF(12, 15))
        painter.drawLine(QPointF(7, 10), QPointF(12, 15))
        painter.drawLine(QPointF(17, 10), QPointF(12, 15))
        painter.drawLine(QPointF(4.5, 20), QPointF(19.5, 20))
    elif name == "upload":
        painter.drawLine(QPointF(12, 20), QPointF(12, 9))
        painter.drawLine(QPointF(7, 14), QPointF(12, 9))
        painter.drawLine(QPointF(17, 14), QPointF(12, 9))
        painter.drawLine(QPointF(4.5, 4), QPointF(19.5, 4))
    elif name == "check":
        path = QPainterPath()
        path.moveTo(5, 12.5)
        path.lineTo(9.7, 17.2)
        path.lineTo(19, 7.5)
        painter.drawPath(path)
    elif name == "chevron-down":
        path = QPainterPath()
        path.moveTo(6, 9)
        path.lineTo(12, 15)
        path.lineTo(18, 9)
        painter.drawPath(path)
    elif name == "shield":
        path = QPainterPath()
        path.moveTo(12, 3)
        path.lineTo(20, 6)
        path.lineTo(20, 11)
        path.cubicTo(20, 16.2, 17, 20.1, 12, 21)
        path.cubicTo(7, 20.1, 4, 16.2, 4, 11)
        path.lineTo(4, 6)
        path.closeSubpath()
        painter.drawPath(path)
        check = QPainterPath()
        check.moveTo(8.2, 12)
        check.lineTo(11, 14.8)
        check.lineTo(16, 9.8)
        painter.drawPath(check)
    elif name == "refresh":
        path = QPainterPath()
        path.arcMoveTo(QRectF(4.5, 4.5, 15, 15), 35)
        path.arcTo(QRectF(4.5, 4.5, 15, 15), 35, 285)
        painter.drawPath(path)
        painter.drawLine(QPointF(17.8, 4.4), QPointF(17.8, 8.5))
        painter.drawLine(QPointF(17.8, 8.5), QPointF(13.7, 8.5))
    elif name == "swap":
        painter.drawLine(QPointF(4, 8.5), QPointF(19, 8.5))
        painter.drawLine(QPointF(14.5, 4), QPointF(19, 8.5))
        painter.drawLine(QPointF(14.5, 13), QPointF(19, 8.5))
        painter.drawLine(QPointF(20, 15.5), QPointF(5, 15.5))
        painter.drawLine(QPointF(9.5, 11), QPointF(5, 15.5))
        painter.drawLine(QPointF(9.5, 20), QPointF(5, 15.5))
    elif name == "info":
        painter.drawEllipse(QRectF(3.2, 3.2, 17.6, 17.6))
        painter.drawLine(QPointF(12, 10.8), QPointF(12, 17))
        painter.setBrush(QColor(color))
        painter.drawEllipse(QPointF(12, 7.4), 0.9, 0.9)
    elif name == "alert":
        path = QPainterPath()
        path.moveTo(12, 3.8)
        path.lineTo(21, 19.4)
        path.lineTo(3, 19.4)
        path.closeSubpath()
        painter.drawPath(path)
        painter.drawLine(QPointF(12, 9), QPointF(12, 14))
        painter.setBrush(QColor(color))
        painter.drawEllipse(QPointF(12, 17), 0.85, 0.85)
    elif name == "x":
        painter.drawLine(QPointF(6, 6), QPointF(18, 18))
        painter.drawLine(QPointF(18, 6), QPointF(6, 18))
    elif name == "file":
        path = QPainterPath()
        path.moveTo(6, 3.5)
        path.lineTo(14, 3.5)
        path.lineTo(19, 8.5)
        path.lineTo(19, 20.5)
        path.lineTo(6, 20.5)
        path.closeSubpath()
        painter.drawPath(path)
        painter.drawLine(QPointF(14, 3.5), QPointF(14, 8.5))
        painter.drawLine(QPointF(14, 8.5), QPointF(19, 8.5))
    elif name == "play":
        path = QPainterPath()
        path.moveTo(7, 4)
        path.lineTo(19, 12)
        path.lineTo(7, 20)
        path.closeSubpath()
        painter.drawPath(path)
    painter.end()
    return QIcon(pixmap)


def app_icon_path() -> str:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base / "assets" / "app.ico")


def app_logo(size: int = 42) -> QPixmap:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    logo = base / "assets" / "logo_128.png"
    if logo.exists():
        pixmap = QPixmap(str(logo))
        if not pixmap.isNull():
            return pixmap.scaled(
                size,
                size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
    # Fallback mark if assets were not packaged.
    pixmap, painter = _new_canvas(size, 1)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#2457D6"))
    painter.drawRoundedRect(QRectF(0, 0, size, size), size * 0.24, size * 0.24)
    painter.setPen(QColor("#FFFFFF"))
    font = QFont("Segoe UI", max(8, int(size * 0.4)))
    font.setBold(True)
    painter.setFont(font)
    painter.drawText(QRectF(0, 0, size, size), Qt.AlignmentFlag.AlignCenter, "W")
    painter.end()
    return pixmap


def badge_pixmap(kind: str, size: int = 34) -> QPixmap:
    pixmap, painter = _new_canvas(size)
    if kind == "pdf":
        painter.setBrush(QColor("#FFF0EE"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(QRectF(0, 0, size, size), 7, 7)
        painter.setPen(QColor("#D4483F"))
        font = QFont("Segoe UI", max(6, int(size * 0.24)))
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(QRectF(0, 1, size, size), Qt.AlignmentFlag.AlignCenter, "PDF")
    else:
        painter.setBrush(QColor("#EAF1FF"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(QRectF(0, 0, size, size), 7, 7)
        painter.setPen(QColor("#2356C9"))
        font = QFont("Segoe UI", max(8, int(size * 0.42)))
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(QRectF(0, -1, size, size), Qt.AlignmentFlag.AlignCenter, "W")
    painter.end()
    return pixmap
