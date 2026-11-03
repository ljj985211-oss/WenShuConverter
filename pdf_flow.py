"""Reflow a PDF into a clean, editable Word document.

The layout-preserving path (pdf2docx) is good at matching the original page,
but on complex PDFs it may create many text frames and layout tables.  This
module deliberately trades exact positioning for a normal Word text flow:
paragraphs, headings, lists, tables and images in reading order.
"""

from __future__ import annotations

import io
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pymupdf
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


BULLET_RE = re.compile(r"^([\u2022\u00b7\u25cf\u25cb\u25aa\u25ab\u2023\u2043\u2219\uf0b7\-–—])\s*")
NUMBER_RE = re.compile(r"^(\(?\d{1,2}[\.\)、]\s*|（\d{1,2}）\s*)")
CJK_RE = re.compile(r"[\u3400-\u9fff\uf900-\ufaff]")
PAGE_NUMBER_RE = re.compile(r"^(第?\s*\d+\s*页?|page\s*\d+|\d+\s*/\s*\d+|\d+)$", re.I)


@dataclass
class RunData:
    text: str
    font: str
    size: float
    bold: bool
    italic: bool
    color: int


@dataclass
class LineData:
    bbox: tuple[float, float, float, float]
    runs: list[RunData]
    text: str
    size: float
    bold_chars: int
    font: str


@dataclass
class TextBlock:
    bbox: tuple[float, float, float, float]
    lines: list[LineData]
    order: int
    margin: str = ""
    skip: bool = False

    @property
    def width(self) -> float:
        return self.bbox[2] - self.bbox[0]

    @property
    def height(self) -> float:
        return self.bbox[3] - self.bbox[1]

    @property
    def text(self) -> str:
        return "".join(line.text for line in self.lines)


@dataclass
class TableBlock:
    bbox: tuple[float, float, float, float]
    rows: list[list[str]]
    order: int = 0

    @property
    def width(self) -> float:
        return self.bbox[2] - self.bbox[0]


@dataclass
class ImageBlock:
    bbox: tuple[float, float, float, float]
    data: bytes
    extension: str
    order: int = 0

    @property
    def width(self) -> float:
        return self.bbox[2] - self.bbox[0]


@dataclass
class PageData:
    number: int
    width: float
    height: float
    blocks: list[TextBlock]
    tables: list[TableBlock]
    images: list[ImageBlock]
    text_chars: int
    scanned: bool
    page_image: bytes | None = None
    header_text: str = ""
    footer_text: str = ""
    two_column: bool = False


def _clean_font(name: str) -> str:
    if "+" in name:
        name = name.split("+", 1)[1]
    return name.replace("-", "").replace(" ", "").strip()


def _is_bold(font: str, flags: int) -> bool:
    folded = font.casefold()
    return bool(flags & 16) or any(
        word in folded for word in ("bold", "black", "heavy", "semibold", "demibold")
    )


def _is_italic(font: str, flags: int) -> bool:
    folded = font.casefold()
    return bool(flags & 2) or any(
        word in folded for word in ("italic", "oblique")
    )


def _page_blocks(page: pymupdf.Page) -> list[TextBlock]:
    data = page.get_text("dict", sort=False)
    blocks: list[TextBlock] = []
    seen: set[tuple[int, int, str]] = set()
    order = 0
    for raw_block in data.get("blocks", []):
        if raw_block.get("type") != 0:
            continue
        lines: list[LineData] = []
        for raw_line in raw_block.get("lines", []):
            direction = raw_line.get("dir") or (1, 0)
            if abs(float(direction[1])) > 0.15:
                continue
            runs: list[RunData] = []
            for span in raw_line.get("spans", []):
                text = span.get("text", "")
                if not text:
                    continue
                runs.append(
                    RunData(
                        text=text,
                        font=str(span.get("font", "")),
                        size=float(span.get("size", 10.0)),
                        bold=_is_bold(str(span.get("font", "")), int(span.get("flags", 0))),
                        italic=_is_italic(str(span.get("font", "")), int(span.get("flags", 0))),
                        color=int(span.get("color", 0)),
                    )
                )
            if not runs:
                continue
            text = "".join(run.text for run in runs)
            if not text.strip():
                continue
            bbox = tuple(float(value) for value in raw_line.get("bbox"))
            key = (round(bbox[0]), round(bbox[1]), text.strip())
            if key in seen:
                continue
            seen.add(key)
            size = max((run.size for run in runs if run.text.strip()), default=10.0)
            bold_chars = sum(len(run.text) for run in runs if run.bold)
            font_counter = Counter()
            for run in runs:
                if run.text.strip():
                    font_counter[run.font] += len(run.text)
            lines.append(
                LineData(
                    bbox=bbox,
                    runs=runs,
                    text=text,
                    size=size,
                    bold_chars=bold_chars,
                    font=font_counter.most_common(1)[0][0] if font_counter else "",
                )
            )
        if not lines:
            continue
        x0 = min(line.bbox[0] for line in lines)
        y0 = min(line.bbox[1] for line in lines)
        x1 = max(line.bbox[2] for line in lines)
        y1 = max(line.bbox[3] for line in lines)
        blocks.append(TextBlock((x0, y0, x1, y1), lines, order))
        order += 1
    return blocks


def _table_rows(table: Any) -> list[list[str]]:
    try:
        raw_rows = table.extract()
    except Exception:
        return []
    rows: list[list[str]] = []
    for raw_row in raw_rows or []:
        row = [str(cell or "").replace("\r", "").strip() for cell in raw_row]
        if any(cell for cell in row):
            rows.append(row)
    if len(rows) < 2:
        return []
    max_cols = max(len(row) for row in rows)
    if max_cols < 2:
        return []
    for row in rows:
        while len(row) < max_cols:
            row.append("")
    return rows


def _page_tables(page: pymupdf.Page) -> list[TableBlock]:
    tables: list[TableBlock] = []
    try:
        finder = page.find_tables()
        raw_tables = list(finder.tables)
    except Exception:
        return tables
    page_area = max(1.0, page.rect.width * page.rect.height)
    for raw_table in raw_tables:
        rows = _table_rows(raw_table)
        if not rows:
            continue
        try:
            bbox = tuple(float(value) for value in raw_table.bbox)
        except Exception:
            continue
        width = bbox[2] - bbox[0]
        height = bbox[3] - bbox[1]
        if width < 80 or height < 22:
            continue
        if width * height / page_area > 0.92:
            continue
        nonempty = [cell for row in rows for cell in row if cell.strip()]
        if len(nonempty) < 2:
            continue
        total_chars = sum(len(cell) for cell in nonempty)
        if total_chars and max(len(cell) for cell in nonempty) > total_chars * 0.82:
            continue
        tables.append(TableBlock(bbox, rows))
    return tables


def _inside(bbox: tuple[float, float, float, float], outer: tuple[float, float, float, float], margin: float = 2.0) -> bool:
    cx = (bbox[0] + bbox[2]) / 2
    cy = (bbox[1] + bbox[3]) / 2
    return (outer[0] - margin) <= cx <= (outer[2] + margin) and (outer[1] - margin) <= cy <= (outer[3] + margin)


def _flatten_pixmap(pix: pymupdf.Pixmap) -> bytes:
    """Composite an RGBA pixmap onto white and return PNG bytes."""
    import numpy as np

    if pix.n - pix.alpha >= 4:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    width, height, channels, stride = pix.width, pix.height, pix.n, pix.stride
    raw = np.frombuffer(pix.samples, dtype=np.uint8)
    raw = raw.reshape(height, stride)[:, : width * channels].reshape(height, width, channels)
    if channels == 4:
        alpha = raw[:, :, 3:4].astype(np.uint16)
        rgb = raw[:, :, :3].astype(np.uint16)
        flat = (rgb * alpha + 255 * (255 - alpha)) // 255
        rgb_bytes = flat.astype(np.uint8).tobytes()
    elif channels >= 3:
        rgb_bytes = raw[:, :, :3].tobytes()
    else:
        gray = raw[:, :, 0:1].repeat(3, axis=2)
        rgb_bytes = gray.tobytes()
    result = pymupdf.Pixmap(pymupdf.csRGB, width, height, rgb_bytes, False)
    return result.tobytes("png")


def _image_bytes(document: pymupdf.Document, xref: int) -> tuple[bytes, str] | None:
    try:
        raw = document.extract_image(xref)
    except Exception:
        return None
    data = raw.get("image")
    extension = str(raw.get("ext") or "png").lower()
    if not data:
        return None
    try:
        pix = pymupdf.Pixmap(document, xref)
        if pix.alpha:
            return _flatten_pixmap(pix), "png"
    except Exception:
        pix = None
    if extension in {"jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff"}:
        if extension == "tif":
            extension = "tiff"
        return bytes(data), extension
    try:
        if pix is None:
            pix = pymupdf.Pixmap(document, xref)
        if pix.n - pix.alpha >= 4:
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        return pix.tobytes("png"), "png"
    except Exception:
        return None


def _page_images(document: pymupdf.Document, page: pymupdf.Page, text_chars: int, tables: list[TableBlock]) -> list[ImageBlock]:
    images: list[ImageBlock] = []
    seen: set[tuple[int, int, int, int]] = set()
    page_area = max(1.0, page.rect.width * page.rect.height)
    try:
        infos = page.get_image_info(xrefs=True)
    except Exception:
        return images
    for info in infos:
        xref = int(info.get("xref") or 0)
        if not xref:
            continue
        try:
            bbox = tuple(float(value) for value in info.get("bbox"))
        except Exception:
            continue
        width = bbox[2] - bbox[0]
        height = bbox[3] - bbox[1]
        if width < 24 or height < 24:
            continue
        key = (xref, round(bbox[0]), round(bbox[1]), round(bbox[2]))
        if key in seen:
            continue
        seen.add(key)
        if any(_inside(bbox, table.bbox) for table in tables):
            continue
        if width * height / page_area > 0.86 and text_chars > 20:
            continue
        if bool(info.get("has-mask")):
            # Render the masked image as it appears on the page. Extracting
            # the raw JPEG would lose its soft mask and produce black corners
            # in Word/WPS.
            try:
                scale = 200 / 72
                pix = page.get_pixmap(
                    clip=pymupdf.Rect(bbox),
                    matrix=pymupdf.Matrix(scale, scale),
                    alpha=False,
                )
                data = pix.tobytes("png")
                extension = "png"
            except Exception:
                extracted = _image_bytes(document, xref)
                if not extracted:
                    continue
                data, extension = extracted
        else:
            extracted = _image_bytes(document, xref)
            if not extracted:
                continue
            data, extension = extracted
        images.append(ImageBlock(bbox, data, extension))
    return images


def _read_page(document: pymupdf.Document, page: pymupdf.Page, number: int) -> PageData:
    text_chars = len(re.sub(r"\s+", "", page.get_text("text") or ""))
    scanned = text_chars < 8
    blocks = _page_blocks(page)
    tables = _page_tables(page)
    if scanned:
        page_image = None
        try:
            pix = page.get_pixmap(dpi=160, alpha=False)
            page_image = pix.tobytes("png")
        except Exception:
            page_image = None
        return PageData(
            number=number,
            width=float(page.rect.width),
            height=float(page.rect.height),
            blocks=[],
            tables=[],
            images=[],
            text_chars=0,
            scanned=True,
            page_image=page_image,
        )

    # Ignore text owned by tables so table content is not duplicated.
    kept_blocks: list[TextBlock] = []
    for block in blocks:
        if any(_inside(block.bbox, table.bbox) for table in tables):
            continue
        kept_blocks.append(block)

    # Some PDF exporters draw tables only with text, or with borders that the
    # table detector misses. Rebuild those rows before paragraph merging.
    grid_tables, used_ids = _text_grid_tables(kept_blocks)
    if grid_tables:
        tables.extend(grid_tables)
        kept_blocks = [block for block in kept_blocks if id(block) not in used_ids]

    kept_blocks = _merge_text_blocks(kept_blocks)
    images = _page_images(document, page, text_chars, tables)
    return PageData(
        number=number,
        width=float(page.rect.width),
        height=float(page.rect.height),
        blocks=kept_blocks,
        tables=tables,
        images=images,
        text_chars=text_chars,
        scanned=False,
    )


def _row_groups(block: TextBlock) -> list[list[LineData]]:
    """Find rows where several lines share the same vertical position."""
    rows: list[list[LineData]] = []
    current: list[LineData] = []
    for line in sorted(block.lines, key=lambda item: (item.bbox[1], item.bbox[0])):
        if not current:
            current = [line]
            continue
        center = (line.bbox[1] + line.bbox[3]) / 2
        current_center = (current[0].bbox[1] + current[0].bbox[3]) / 2
        if abs(center - current_center) <= 4.0:
            current.append(line)
        else:
            if len(current) >= 2:
                rows.append(current)
            current = [line]
    if len(current) >= 2:
        rows.append(current)
    return rows


def _row_cells(row: list[LineData]) -> list[LineData] | None:
    cells = sorted(row, key=lambda item: item.bbox[0])
    if len(cells) < 2:
        return None
    # A bullet + text pair is a list item, not a two-column table.
    if BULLET_RE.match(cells[0].text.strip()) and len(cells) == 2:
        return None
    for previous, following in zip(cells, cells[1:]):
        if following.bbox[0] <= previous.bbox[2] + 4:
            return None
    return cells


def _text_grid_tables(blocks: list[TextBlock]) -> tuple[list[TableBlock], set[int]]:
    """Rebuild simple borderless tables from parallel lines of text."""
    rows: list[tuple[float, list[LineData], TextBlock]] = []
    for block in blocks:
        for group in _row_groups(block):
            cells = _row_cells(group)
            if cells and len(cells) >= 2:
                rows.append(((cells[0].bbox[1] + cells[0].bbox[3]) / 2, cells, block))
    if len(rows) < 2:
        return [], set()

    rows.sort(key=lambda item: item[0])
    tables: list[TableBlock] = []
    used_blocks: set[int] = set()
    current: list[list[LineData]] = []
    current_starts: list[float] = []

    def flush() -> None:
        nonlocal current, current_starts
        if len(current) >= 2:
            # A two-column, two-row arrangement of short lines is often a
            # numbered section heading captured as side-by-side text, not a
            # real table. Leave it in the paragraph flow in that case.
            if len(current[0]) <= 2 and len(current) < 3:
                if all(
                    len(cell.text.strip()) <= 24
                    for row in current
                    for cell in row
                ):
                    current = []
                    current_starts = []
                    return
            row_texts = [
                [cell.text.strip() for cell in sorted(row, key=lambda item: item.bbox[0])]
                for row in current
            ]
            xs0 = [cell.bbox[0] for row in current for cell in row]
            xs1 = [cell.bbox[2] for row in current for cell in row]
            ys0 = [cell.bbox[1] for row in current for cell in row]
            ys1 = [cell.bbox[3] for row in current for cell in row]
            tables.append(TableBlock((min(xs0), min(ys0), max(xs1), max(ys1)), row_texts))
        current = []
        current_starts = []

    for _y, cells, block in rows:
        starts = [cell.bbox[0] for cell in cells]
        if not current:
            current = [cells]
            current_starts = starts
            used_blocks.add(id(block))
            continue
        same_shape = (
            len(cells) == len(current[-1])
            and len(starts) == len(current_starts)
            and all(abs(a - b) <= 22 for a, b in zip(starts, current_starts))
        )
        row_gap = cells[0].bbox[1] - max(cell.bbox[3] for cell in current[-1])
        if same_shape and -2 <= row_gap <= 26:
            current.append(cells)
            used_blocks.add(id(block))
        else:
            flush()
            current = [cells]
            current_starts = starts
            used_blocks.add(id(block))
    flush()
    return tables, used_blocks


def _merge_text_blocks(blocks: list[TextBlock]) -> list[TextBlock]:
    merged: list[TextBlock] = []
    for block in sorted(blocks, key=lambda item: (round(item.bbox[1], 1), item.bbox[0])):
        if not block.lines:
            continue
        if not merged:
            merged.append(block)
            continue
        previous = merged[-1]
        previous_line = previous.lines[-1]
        first_line = block.lines[0]
        gap = block.bbox[1] - previous.bbox[3]
        size_gap = abs(first_line.size - previous_line.size)
        same_left = abs(block.bbox[0] - previous.bbox[0]) <= 14
        starts_list = bool(BULLET_RE.match(block.text.strip()) or NUMBER_RE.match(block.text.strip()))
        if (
            0 <= gap <= max(4.5, previous_line.size * 0.95)
            and size_gap <= 1.5
            and same_left
            and not starts_list
        ):
            previous.lines.extend(block.lines)
            previous.bbox = (
                min(previous.bbox[0], block.bbox[0]),
                min(previous.bbox[1], block.bbox[1]),
                max(previous.bbox[2], block.bbox[2]),
                max(previous.bbox[3], block.bbox[3]),
            )
        else:
            merged.append(block)
    return merged


def _normalize_repeat_text(text: str) -> str:
    text = re.sub(r"\d+", "#", text)
    return re.sub(r"\s+", "", text)


def _mark_headers_and_footers(pages: list[PageData]) -> tuple[str, str]:
    if not pages:
        return "", ""
    top_count: Counter[str] = Counter()
    bottom_count: Counter[str] = Counter()
    top_first: dict[str, str] = {}
    bottom_first: dict[str, str] = {}
    for page in pages:
        if page.scanned:
            continue
        for block in page.blocks:
            center = (block.bbox[1] + block.bbox[3]) / 2
            zone = ""
            if center < page.height * 0.085:
                zone = "top"
            elif center > page.height * 0.915:
                zone = "bottom"
            if not zone:
                continue
            block.margin = zone
            raw = block.text.strip()
            if not raw:
                continue
            if PAGE_NUMBER_RE.match(raw):
                block.skip = True
                continue
            normalized = _normalize_repeat_text(raw)
            if not normalized or len(normalized) > 140:
                continue
            if zone == "top":
                top_count[normalized] += 1
                top_first.setdefault(normalized, raw)
            else:
                bottom_count[normalized] += 1
                bottom_first.setdefault(normalized, raw)

    threshold = max(2, int(len(pages) * 0.55 + 0.999))
    header_text = ""
    footer_text = ""
    if top_count:
        normalized, count = top_count.most_common(1)[0]
        if count >= threshold:
            header_text = top_first[normalized]
    if bottom_count:
        normalized, count = bottom_count.most_common(1)[0]
        if count >= threshold:
            footer_text = bottom_first[normalized]

    for page in pages:
        if page.scanned:
            continue
        for block in page.blocks:
            if block.margin:
                normalized = _normalize_repeat_text(block.text.strip())
                if PAGE_NUMBER_RE.match(block.text.strip()):
                    block.skip = True
                elif block.margin == "top" and normalized == _normalize_repeat_text(header_text):
                    block.skip = True
                elif block.margin == "bottom" and normalized == _normalize_repeat_text(footer_text):
                    block.skip = True
    return header_text, footer_text


def _detect_columns(page: PageData) -> bool:
    blocks = [block for block in page.blocks if not block.skip and block.width > 35]
    if len(blocks) < 6:
        return False
    middle = page.width / 2
    left = sum(1 for block in blocks if block.bbox[2] < middle - 15)
    right = sum(1 for block in blocks if block.bbox[0] > middle + 15)
    crossing = sum(
        1 for block in blocks if block.bbox[0] < middle - 15 and block.bbox[2] > middle + 15
    )
    return left >= 3 and right >= 3 and crossing <= max(1, len(blocks) // 5)


def _order_elements(page: PageData) -> list[Any]:
    text_blocks = [block for block in page.blocks if not block.skip]
    others: list[Any] = list(page.tables) + list(page.images)
    if not text_blocks:
        return sorted(others, key=lambda item: (item.bbox[1], item.bbox[0]))

    if not page.two_column:
        return sorted(text_blocks + others, key=lambda item: (round(item.bbox[1], 1), item.bbox[0]))

    full_width: list[Any] = []
    narrow: list[Any] = []
    for item in text_blocks + others:
        if item.width >= page.width * 0.62:
            full_width.append(item)
        else:
            narrow.append(item)
    full_width.sort(key=lambda item: item.bbox[1])

    groups: dict[int, list[Any]] = defaultdict(list)
    for item in narrow:
        segment = sum(1 for anchor in full_width if anchor.bbox[3] <= item.bbox[1] + 1)
        groups[segment].append(item)

    def column_key(item: Any) -> tuple[int, float, float]:
        middle = (item.bbox[0] + item.bbox[2]) / 2
        column = 0 if middle < page.width / 2 else 1
        return (column, round(item.bbox[1], 1), item.bbox[0])

    ordered: list[Any] = []
    for segment in range(len(full_width) + 1):
        ordered.extend(sorted(groups.get(segment, []), key=column_key))
        if segment < len(full_width):
            ordered.append(full_width[segment])
    return ordered


def _join_character(previous: str, following: str) -> str:
    if not previous:
        return ""
    if previous.endswith((" ", "\t", "\n")) or following.startswith((" ", "\t", "\n")):
        return ""
    if CJK_RE.match(previous[-1]) or CJK_RE.match(following[:1]):
        return ""
    if following[:1] in "，。！？；：、）】》」』":
        return ""
    return " "


def _paragraph_groups(block: TextBlock) -> list[list[LineData]]:
    groups: list[list[LineData]] = []
    current: list[LineData] = []
    for line in block.lines:
        if not current:
            current = [line]
            continue
        previous = current[-1]
        gap = line.bbox[1] - previous.bbox[3]
        size_change = abs(line.size - previous.size) > 2.2
        new_paragraph = gap > max(8.5, min(15.0, previous.size * 1.05)) or size_change
        if not new_paragraph and BULLET_RE.match(line.text.strip()):
            new_paragraph = True
        if not new_paragraph and line.bbox[0] - previous.bbox[0] > 20:
            if previous.text.rstrip().endswith(("。", "！", "？", ".", "!", "?", "：")):
                new_paragraph = True
        if new_paragraph:
            groups.append(current)
            current = [line]
        else:
            current.append(line)
    if current:
        groups.append(current)
    return groups


def _group_text(group: list[LineData]) -> str:
    parts: list[str] = []
    for line in group:
        text = line.text
        if parts:
            separator = _join_character(parts[-1], text)
            parts.append(separator)
        parts.append(text)
    return "".join(parts).strip()


def _font_pair(font_name: str) -> tuple[str, str]:
    folded = _clean_font(font_name).casefold()
    if any(token in folded for token in ("courier", "consol", "mono")):
        return "Consolas", "DengXian"
    if any(token in folded for token in ("times", "georgia", "cambria", "serif", "song", "simsun", "宋")):
        return "Times New Roman", "SimSun"
    if any(token in folded for token in ("fangsong", "仿宋")):
        return "Arial", "FangSong"
    if any(token in folded for token in ("kaiti", "楷体", "楷")):
        return "Arial", "KaiTi"
    if any(token in folded for token in ("simhei", "黑体")):
        return "Arial", "SimHei"
    if any(token in folded for token in ("dengxian", "等线")):
        return "Arial", "DengXian"
    return "Arial", "Microsoft YaHei"


def _set_run_font(run, font_name: str, size: Pt | None, bold: bool, italic: bool, color: int | None) -> None:
    latin, east = _font_pair(font_name)
    run.font.name = latin
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), east)
    if size is not None:
        run.font.size = size
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    if color is not None and color not in (0, 0xFFFFFF):
        run.font.color.rgb = RGBColor((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF)


def _body_font_size(pages: list[PageData]) -> float:
    counter: Counter[float] = Counter()
    for page in pages:
        if page.scanned:
            continue
        for block in page.blocks:
            if block.skip:
                continue
            for line in block.lines:
                if line.text.strip():
                    counter[round(line.size * 2) / 2] += max(1, len(line.text.strip()))
    if not counter:
        return 10.5
    size = counter.most_common(1)[0][0]
    return max(7.0, min(18.0, size))


def _heading_level(text: str, size: float, body: float, bold_ratio: float, line_count: int) -> int:
    if len(text) > 120 or line_count > 3 or BULLET_RE.match(text) or NUMBER_RE.match(text):
        return 0
    ratio = size / max(1.0, body)
    if ratio >= 1.5:
        return 1
    if ratio >= 1.22:
        return 2
    if bold_ratio >= 0.55 and ratio >= 1.08 and len(text) < 70:
        return 3
    if re.match(r"^(第[一二三四五六七八九十百]+[章节篇]|[一二三四五六七八九十]+、)", text) and (
        bold_ratio >= 0.35 or ratio >= 1.08
    ):
        return 2
    return 0


def _clean_heading_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"(?<=[\u4e00-\u9fff0-9A-Za-z])\s+(?=[\u4e00-\u9fff])", "", text)
    text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", text)
    return text


def _add_paragraph(document: Document, group: list[LineData], body_size: float, page_width: float, content_width: float):
    text = _group_text(group)
    if not text:
        return None
    size = max((line.size for line in group), default=body_size)
    total_chars = max(1, len(text))
    bold_chars = sum(line.bold_chars for line in group)
    bold_ratio = bold_chars / total_chars
    level = _heading_level(text, size, body_size, bold_ratio, len(group))
    style_name = f"Heading {level}" if level else "Normal"
    paragraph = document.add_paragraph(style=style_name)
    paragraph.paragraph_format.space_before = Pt(10 if level == 1 else 7 if level == 2 else 4 if level == 3 else 0)
    paragraph.paragraph_format.space_after = Pt(6 if level else 5)
    if not level:
        paragraph.paragraph_format.line_spacing = 1.35

    left = min(line.bbox[0] for line in group)
    right = max(line.bbox[2] for line in group)
    width = right - left
    center = (left + right) / 2
    centered = (
        len(text) <= 80
        and len(group) <= 2
        and width < page_width * 0.72
        and abs(center - page_width / 2) < page_width * 0.04
    )
    paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.LEFT
    )

    if level:
        first_font = next((run.font for run in group[0].runs if run.text.strip()), "")
        first_color = next((run.color for run in group[0].runs if run.text.strip()), None)
        run = paragraph.add_run(_clean_heading_text(text))
        _set_run_font(run, first_font, None, True, False, first_color)
        return paragraph

    bullet_match = BULLET_RE.match(text.strip())
    replace_bullet = bullet_match is not None
    prefix_done = not replace_bullet

    previous_text = ""
    for line_index, line in enumerate(group):
        line_started = False
        for run_data in line.runs:
            if not run_data.text:
                continue
            run_text = run_data.text
            if not prefix_done and run_text.strip():
                leading = run_text[: len(run_text) - len(run_text.lstrip())]
                stripped_run = run_text.lstrip()
                marker = BULLET_RE.match(stripped_run)
                if marker:
                    run_text = leading + "• " + stripped_run[marker.end() :].lstrip()
                else:
                    run_text = leading + "• " + stripped_run
                prefix_done = True
            if not line_started and previous_text:
                separator = _join_character(previous_text, run_data.text)
                if separator:
                    paragraph.add_run(separator)
                    previous_text += separator
            line_started = True
            run = paragraph.add_run(run_text)
            if level:
                _set_run_font(
                    run,
                    run_data.font,
                    None,
                    True if level else run_data.bold,
                    run_data.italic,
                    run_data.color,
                )
            else:
                font_size = Pt(8.5 if run_data.size < body_size * 0.8 else 10.5)
                _set_run_font(
                    run,
                    run_data.font,
                    font_size,
                    run_data.bold,
                    run_data.italic,
                    run_data.color,
                )
            previous_text += run_text
    paragraph_text = paragraph.text.strip()
    if not paragraph_text:
        return None

    # Indentation is derived from the original line geometry.
    x0s = [line.bbox[0] for line in group]
    stripped = paragraph_text.lstrip()
    if not level:
        if BULLET_RE.match(stripped):
            paragraph.paragraph_format.left_indent = Cm(0.72)
            paragraph.paragraph_format.first_line_indent = Cm(-0.38)
        elif re.match(r"^\d{1,2}[\.\)、]\s+", stripped):
            paragraph.paragraph_format.left_indent = Cm(0.72)
            paragraph.paragraph_format.first_line_indent = Cm(-0.38)
        elif len(group) > 1 and x0s[0] - min(x0s[1:]) > 9:
            paragraph.paragraph_format.first_line_indent = Pt(
                min(24, max(8, x0s[0] - min(x0s[1:])))
            )
        elif CJK_RE.search(stripped) and len(stripped) > 12:
            paragraph.paragraph_format.first_line_indent = Pt(21)
    return paragraph


def _set_cell_text(cell, text: str, bold: bool = False) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.1
    lines = [line for line in text.splitlines() if line.strip()]
    if not lines:
        return
    for index, line_text in enumerate(lines):
        target = paragraph if index == 0 else cell.add_paragraph()
        target.paragraph_format.space_after = Pt(0)
        run = target.add_run(line_text)
        run.font.name = "Arial"
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        run.font.size = Pt(9)
        run.bold = bold


def _add_table(document: Document, table_block: TableBlock, available_width: float) -> None:
    rows = table_block.rows
    if not rows:
        return
    cols = max(len(row) for row in rows)
    table = document.add_table(rows=len(rows), cols=cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for row_index, row in enumerate(rows):
        for col_index in range(cols):
            text = row[col_index] if col_index < len(row) else ""
            _set_cell_text(table.cell(row_index, col_index), text, bold=(row_index == 0))
    document.add_paragraph().paragraph_format.space_after = Pt(0)


def _add_image(document: Document, image: ImageBlock, available_width: float, page_width: float) -> None:
    width = image.bbox[2] - image.bbox[0]
    if width <= 0:
        width = available_width
    width = max(80.0, min(width, available_width, page_width))
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(5)
    paragraph.paragraph_format.space_after = Pt(5)
    try:
        paragraph.add_run().add_picture(io.BytesIO(image.data), width=Pt(width))
    except Exception:
        # Unsupported image encodings are skipped rather than breaking the
        # whole document conversion.
        paragraph._element.getparent().remove(paragraph._element)


def _configure_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    normal._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.paragraph_format.line_spacing = 1.35
    normal.paragraph_format.space_after = Pt(5)
    try:
        normal.paragraph_format.widow_control = True
    except Exception:
        pass
    for name, size in (("Heading 1", 17), ("Heading 2", 14), ("Heading 3", 12)):
        style = document.styles[name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        style.font.bold = True


def _setup_page(document: Document, page: PageData) -> None:
    section = document.sections[0]
    section.page_width = Pt(page.width)
    section.page_height = Pt(page.height)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.1)
    section.right_margin = Cm(2.1)
    section.header_distance = Cm(0.9)
    section.footer_distance = Cm(0.9)


def _set_header_footer(document: Document, header: str, footer: str) -> None:
    section = document.sections[0]
    if header:
        paragraph = section.header.paragraphs[0]
        paragraph.text = ""
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(header)
        run.font.name = "Arial"
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0x7A, 0x86, 0x96)
    if footer:
        paragraph = section.footer.paragraphs[0]
        paragraph.text = ""
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(footer)
        run.font.name = "Arial"
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0x8A, 0x96, 0xA6)


def convert_pdf_to_flow_docx(source: Path, destination: Path) -> dict[str, Any]:
    document = pymupdf.open(str(source))
    try:
        if document.needs_pass:
            raise RuntimeError("encrypted")
        pages: list[PageData] = []
        for index in range(document.page_count):
            page = document.load_page(index)
            pages.append(_read_page(document, page, index + 1))
        if not pages:
            raise RuntimeError("empty")

        header_text, footer_text = _mark_headers_and_footers(pages)
        body_size = _body_font_size(pages)
        for page in pages:
            page.two_column = _detect_columns(page)

        output = Document()
        _configure_styles(output)
        _setup_page(output, pages[0])
        _set_header_footer(output, header_text, footer_text)
        available_width = max(
            240.0,
            pages[0].width - 2 * 59.5,  # approximate 2.1 cm margins in points
        )

        scanned_pages = 0
        table_count = 0
        image_count = 0
        break_pending = False
        for page_index, page in enumerate(pages):
            if page.scanned:
                scanned_pages += 1
                if page_index > 0 and not break_pending:
                    output.add_page_break()
                    break_pending = True
                if page.page_image:
                    paragraph = output.add_paragraph()
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    try:
                        paragraph.add_run().add_picture(
                            io.BytesIO(page.page_image), width=Pt(available_width)
                        )
                    except Exception:
                        pass
                if page_index < len(pages) - 1:
                    output.add_page_break()
                    break_pending = True
            else:
                for element in _order_elements(page):
                    if isinstance(element, TextBlock):
                        for group in _paragraph_groups(element):
                            _add_paragraph(
                                output, group, body_size, page.width, available_width
                            )
                    elif isinstance(element, TableBlock):
                        _add_table(output, element, available_width)
                        table_count += 1
                    elif isinstance(element, ImageBlock):
                        _add_image(output, element, available_width, page.width)
                        image_count += 1
                # Reflowed pages flow continuously. Forcing a break after
                # every source page produces nearly blank pages when the
                # extracted text needs more or less room than the PDF.
                break_pending = False

        output.save(str(destination))

        warning = ""
        if scanned_pages:
            warning = (
                f"检测到 {scanned_pages} 页扫描内容，已按页面图片还原，无法直接编辑其中的文字。"
            )
        return {
            "pages": len(pages),
            "warning": warning,
            "engine": "Editable reflow",
            "tables": table_count,
            "images": image_count,
        }
    finally:
        document.close()
