"""Core document conversion logic for WenShu Converter.

No Qt imports live here so the same code can run in a lightweight child
process.  Word -> PDF uses LibreOffice (system install or the bundled kit),
PDF -> Word uses pdf2docx.
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Any

APP_NAME = "WenShuConverter"
VERSION = "1.1.0"

WORD_EXTENSIONS = {".docx", ".doc"}
PDF_EXTENSIONS = {".pdf"}

CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


class ConversionError(Exception):
    """A conversion failure with a message suitable for the UI."""

    def __init__(self, message: str, detail: str = "") -> None:
        super().__init__(message)
        self.message = message
        self.detail = detail


def app_dir() -> Path:
    """Directory containing the executable (or this source file)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _is_file(path: Path) -> bool:
    try:
        return path.is_file()
    except OSError:
        return False


def _find_system_soffice() -> Path | None:
    override = os.environ.get("WENSHU_SOFFICE", "").strip().strip('"')
    candidates: list[Path] = []
    if override:
        candidates.append(Path(override))

    for env_name in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        root = os.environ.get(env_name)
        if not root:
            continue
        candidates.extend(
            [
                Path(root) / "LibreOffice" / "program" / "soffice.exe",
                Path(root) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
            ]
        )

    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        candidates.append(Path(found))

    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate).casefold()
        if key in seen:
            continue
        seen.add(key)
        if _is_file(candidate):
            return candidate.resolve()
    return None


def _kit_candidates() -> list[tuple[Path, Path]]:
    """Return plausible (node.exe, cli.js) pairs, best candidate first."""
    base = app_dir()
    pairs: list[tuple[Path, Path]] = []

    # Layout used by the packaged application.
    for engine_root in (base / "engine", base / "runtime" / "engine"):
        node = engine_root / "node" / "node.exe"
        cli = engine_root / "libreoffice-kit" / "lib" / "cli.js"
        pairs.append((node, cli))
        pairs.append(
            (
                engine_root / "node.exe",
                engine_root / "libreoffice-kit" / "lib" / "cli.js",
            )
        )
        pairs.append(
            (
                engine_root / "node.exe",
                engine_root
                / "node_modules"
                / "@deepseek-ai"
                / "libreoffice-kit"
                / "lib"
                / "cli.js",
            )
        )
        # Packaged layout: engine/node/node.exe + engine/node_modules/...
        pairs.append(
            (
                engine_root / "node" / "node.exe",
                engine_root
                / "node_modules"
                / "@deepseek-ai"
                / "libreoffice-kit"
                / "lib"
                / "cli.js",
            )
        )

    # Development fallback: the LibreOffice Kit that ships with DSH.
    dsh_roots = [Path(r"D:\DeepSeek\resources")]
    configured_root = os.environ.get("DSH_RESOURCES", "").strip()
    if configured_root:
        dsh_roots.append(Path(configured_root))
    for root in dsh_roots:
        pairs.append(
            (
                root
                / "runtime"
                / "primary-runtime"
                / "dependencies"
                / "node"
                / "bin"
                / "node.exe",
                root
                / "app.asar.unpacked"
                / "dsh"
                / "node_modules"
                / "@deepseek-ai"
                / "libreoffice-kit"
                / "lib"
                / "cli.js",
            )
        )

    # A Node installation on PATH may still be paired with a bundled kit.
    node_on_path = shutil.which("node")
    if node_on_path:
        for _, cli in list(pairs):
            if _is_file(cli):
                pairs.append((Path(node_on_path), cli))

    unique: list[tuple[Path, Path]] = []
    seen: set[tuple[str, str]] = set()
    for node, cli in pairs:
        key = (str(node).casefold(), str(cli).casefold())
        if key in seen:
            continue
        seen.add(key)
        if _is_file(node) and _is_file(cli):
            unique.append((node.resolve(), cli.resolve()))
    return unique


def _com_engine_available(progid: str) -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID"):
            return True
    except (OSError, ImportError):
        return False


def _find_com_engine() -> dict[str, str] | None:
    """Prefer an installed Word-compatible COM engine (WPS or Word).

    WPS uses the exact fonts and layout engine that the user sees in WPS
    Writer, which preserves Chinese typography considerably better than a
    generic headless renderer on many documents.
    """
    for progid, label in (
        ("KWPS.Application", "WPS Office"),
        ("Word.Application", "Microsoft Word"),
    ):
        if _com_engine_available(progid):
            return {"kind": "com", "progid": progid, "label": label}
    return None


def _find_libreoffice_engine() -> dict[str, str] | None:
    soffice = _find_system_soffice()
    if soffice:
        return {"kind": "soffice", "path": str(soffice), "label": "LibreOffice"}

    kits = _kit_candidates()
    if kits:
        node, cli = kits[0]
        return {
            "kind": "kit",
            "node": str(node),
            "cli": str(cli),
            "label": "LibreOffice 内置引擎",
        }
    return None


def find_word_engine(preferred_soffice: str | None = None) -> dict[str, str] | None:
    """Detect the best Word -> PDF engine and describe it with JSON-safe values."""
    if preferred_soffice:
        candidate = Path(preferred_soffice.strip().strip('"'))
        if _is_file(candidate):
            return {
                "kind": "soffice",
                "path": str(candidate.resolve()),
                "label": "LibreOffice",
            }

    # WPS/Word COM normally gives the closest match to the fonts and layout
    # the user sees in their word processor.
    com_engine = _find_com_engine()
    if com_engine:
        return com_engine

    return _find_libreoffice_engine()


def pdf_engine_available() -> bool:
    from importlib.util import find_spec

    return find_spec("pdf2docx") is not None


def _probe_output_directory(directory: Path) -> None:
    try:
        directory.mkdir(parents=True, exist_ok=True)
        fd, probe = tempfile.mkstemp(prefix=".wenshu-write-", dir=directory)
        os.close(fd)
        os.unlink(probe)
    except OSError as exc:
        raise ConversionError(
            f"无法写入输出文件夹：{directory}",
            f"{type(exc).__name__}: {exc}",
        ) from exc


def unique_output_path(directory: Path, stem: str, extension: str) -> Path:
    """Pick a non-existing target, keeping the familiar 'name (1).ext' style."""
    stem = stem.strip() or "转换结果"
    target = directory / f"{stem}{extension}"
    index = 1
    while target.exists():
        target = directory / f"{stem} ({index}){extension}"
        index += 1
    return target


def _verify_pdf(path: Path) -> None:
    try:
        size = path.stat().st_size
        with path.open("rb") as stream:
            header = stream.read(5)
    except OSError as exc:
        raise ConversionError("转换输出文件无法读取。", str(exc)) from exc
    if size < 100 or header != b"%PDF-":
        raise ConversionError("转换引擎没有生成有效的 PDF 文件。")


def _verify_docx(path: Path) -> None:
    try:
        size = path.stat().st_size
    except OSError as exc:
        raise ConversionError("转换输出文件无法读取。", str(exc)) from exc
    if size < 100 or not zipfile.is_zipfile(path):
        raise ConversionError("转换引擎没有生成有效的 Word 文件。")
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
    if "[Content_Types].xml" not in names or "word/document.xml" not in names:
        raise ConversionError("转换引擎生成的文件不是有效的 DOCX 文档。")


def _run_checked(args: list[str], *, timeout: int, cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    try:
        return subprocess.run(
            args,
            cwd=str(cwd) if cwd else None,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            creationflags=CREATE_NO_WINDOW,
        )
    except subprocess.TimeoutExpired as exc:
        raise ConversionError(
            "转换超时。文档可能过大或排版过于复杂。",
            f"命令执行超过 {timeout} 秒。",
        ) from exc
    except OSError as exc:
        raise ConversionError("无法启动 Word 转换引擎。", str(exc)) from exc


def _word_to_pdf_com(
    source: Path,
    destination: Path,
    engine: dict[str, str],
) -> dict[str, Any]:
    """Export with WPS/Word automation so the document's own font and layout
    engine is used. This is the preferred path for Chinese documents."""
    try:
        import comtypes.client  # type: ignore
    except Exception as exc:
        raise ConversionError(
            "WPS/Word 自动化组件不可用。",
            f"{type(exc).__name__}: {exc}",
        ) from exc

    progid = engine.get("progid") or "KWPS.Application"
    application = None
    document = None
    try:
        # dynamic=True uses IDispatch late binding and avoids writing a
        # generated typelib wrapper for every child process. This is much
        # faster in a frozen PyInstaller build and avoids cache permissions.
        application = comtypes.client.CreateObject(progid, dynamic=True)
        try:
            application.Visible = False
        except Exception:
            pass
        try:
            application.DisplayAlerts = 0
        except Exception:
            pass
        try:
            application.ScreenUpdating = False
        except Exception:
            pass

        document = application.Documents.Open(str(source), False, True)
        if document is None:
            raise ConversionError("WPS/Word 未能打开该文档，文件可能已损坏或受密码保护。")

        try:
            document.ExportAsFixedFormat(str(destination), 17)
        except Exception:
            # Older WPS builds do not implement ExportAsFixedFormat.
            if destination.exists():
                try:
                    destination.unlink()
                except OSError:
                    pass
            document.SaveAs(str(destination), 17)

        if not destination.exists() or destination.stat().st_size < 100:
            raise ConversionError("WPS/Word 没有生成有效的 PDF 文件。")
        return {"engine": engine.get("label") or "WPS Office", "missing_fonts": []}
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(
            "WPS/Word 排版导出失败。",
            f"{type(exc).__name__}: {exc}",
        ) from exc
    finally:
        if document is not None:
            try:
                document.Close(0)
            except Exception:
                pass
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass


def _word_to_pdf_soffice(
    source: Path,
    destination: Path,
    soffice: Path,
    work_dir: Path,
) -> dict[str, Any]:
    out_dir = work_dir / "lo-output"
    out_dir.mkdir(parents=True, exist_ok=True)
    profile = work_dir / "lo-profile"
    profile.mkdir(parents=True, exist_ok=True)
    profile_uri = profile.resolve().as_uri()
    args = [
        str(soffice),
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        "--norestore",
        f"-env:UserInstallation={profile_uri}",
        "--convert-to",
        "pdf:writer_pdf_Export",
        "--outdir",
        str(out_dir),
        str(source),
    ]
    result = _run_checked(args, timeout=300)
    pdfs = sorted(out_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not pdfs:
        detail = (result.stderr or result.stdout or "LibreOffice 未返回错误信息").strip()
        raise ConversionError(
            "Word 转 PDF 失败。请确认文档没有损坏，且未被其他程序占用。",
            detail[-3000:],
        )
    shutil.copyfile(pdfs[0], destination)
    return {"engine": "LibreOffice", "missing_fonts": []}


def _word_to_pdf_kit(
    source: Path,
    destination: Path,
    engine: dict[str, str],
    work_dir: Path,
) -> dict[str, Any]:
    args = [
        engine["node"],
        engine["cli"],
        "convert",
        "--input",
        str(source),
        "--output",
        str(destination),
    ]
    result = _run_checked(args, timeout=300, cwd=Path(engine["cli"]).parent)
    missing_fonts: list[str] = []
    payload: dict[str, Any] | None = None
    for line in (result.stdout or "").splitlines():
        line = line.strip()
        if line.startswith("{"):
            try:
                payload = json.loads(line)
            except json.JSONDecodeError:
                pass
    if payload:
        raw_fonts = payload.get("missingFonts") or []
        if isinstance(raw_fonts, list):
            missing_fonts = [str(font) for font in raw_fonts if str(font).strip()]
    if result.returncode != 0 or not destination.exists():
        detail = (result.stderr or result.stdout or "内置引擎未返回错误信息").strip()
        if payload and payload.get("error"):
            detail = str(payload.get("error")) + "\n" + detail
        raise ConversionError(
            "Word 转 PDF 失败。请确认文档没有损坏，且格式受支持。",
            detail[-3000:],
        )
    return {"engine": "LibreOffice", "missing_fonts": missing_fonts}


_COMMON_PDF_FONTS = {
    "courier",
    "courier new",
    "symbol",
    "arial",
    "times new roman",
    "calibri",
    "cambria",
    "wingdings",
}


def _font_warning(missing_fonts: list[str]) -> str:
    relevant = [
        font
        for font in missing_fonts
        if font.strip() and font.strip().casefold() not in _COMMON_PDF_FONTS
    ]
    if not relevant:
        return ""
    shown = "、".join(relevant[:6])
    return f"文档使用了本机未安装的字体：{shown}；已使用替代字体，版面可能有差异。"


def _word_to_pdf_with_engine(
    source: Path,
    destination: Path,
    engine: dict[str, str],
    work_dir: Path,
) -> dict[str, Any]:
    kind = engine.get("kind")
    if kind == "com":
        return _word_to_pdf_com(source, destination, engine)
    if kind == "soffice":
        return _word_to_pdf_soffice(source, destination, Path(engine["path"]), work_dir)
    if kind == "kit":
        return _word_to_pdf_kit(source, destination, engine, work_dir)
    raise ConversionError("没有可用的 Word 转换引擎。")


def convert_word_to_pdf(
    source: Path,
    destination: Path,
    engine: dict[str, str],
    work_dir: Path,
) -> dict[str, Any]:
    _verify_source(source, WORD_EXTENSIONS)
    destination.parent.mkdir(parents=True, exist_ok=True)

    if engine.get("kind") == "com":
        try:
            info = _word_to_pdf_com(source, destination, engine)
            _verify_pdf(destination)
        except ConversionError as primary:
            if destination.exists():
                try:
                    destination.unlink()
                except OSError:
                    pass
            fallback = _find_libreoffice_engine()
            if fallback is None:
                raise primary
            try:
                info = _word_to_pdf_with_engine(source, destination, fallback, work_dir)
                _verify_pdf(destination)
            except ConversionError as secondary:
                raise ConversionError(
                    "Word 转 PDF 失败。WPS/Word 与 LibreOffice 均未能完成转换。",
                    f"WPS/Word: {primary.detail or primary}\nLibreOffice: {secondary.detail or secondary}",
                ) from secondary
            info["warning"] = "WPS/Word 排版输出失败，已自动改用 LibreOffice 引擎。"
            return info
    else:
        info = _word_to_pdf_with_engine(source, destination, engine, work_dir)
        _verify_pdf(destination)

    warning = _font_warning(info.get("missing_fonts") or [])
    if warning:
        info["warning"] = warning
    return info


def _pdf_page_info(source: Path) -> tuple[int, bool]:
    """Return (page_count, likely_scanned)."""
    try:
        import pymupdf as fitz  # type: ignore
    except Exception:
        try:
            import fitz  # type: ignore
        except Exception:
            return 0, False
    try:
        document = fitz.open(str(source))
    except Exception as exc:
        raise ConversionError("无法打开 PDF，文件可能已损坏。", str(exc)) from exc
    try:
        if document.needs_pass:
            raise ConversionError("该 PDF 已加密，请先解除密码保护后再转换。")
        pages = document.page_count
        if pages <= 0:
            raise ConversionError("该 PDF 没有可转换的页面。")
        sample_text = 0
        for index in range(min(3, pages)):
            sample_text += len(document.load_page(index).get_text("text").strip())
        return pages, sample_text < 10
    finally:
        document.close()


_BULLET_CHARS = {
    "\uf0b7",
    "•",
    "·",
    "●",
    "○",
    "▪",
    "▫",
    "‣",
    "⁃",
    "∙",
    "-",
    "–",
}


def _is_bullet_line(text: str) -> bool:
    stripped = text.strip()
    return bool(stripped) and set(stripped) <= _BULLET_CHARS


def _contains_bullet_table(docx_path: Path) -> bool:
    try:
        from docx import Document  # type: ignore
    except Exception:
        return False
    try:
        document = Document(str(docx_path))
    except Exception:
        return False
    for table in document.tables:
        if len(table.columns) != 2:
            continue
        for row in table.rows:
            try:
                left_lines = [
                    line for line in row.cells[0].text.splitlines() if line.strip()
                ]
                right_text = row.cells[1].text.strip()
            except IndexError:
                continue
            if right_text and left_lines and all(
                _is_bullet_line(line) for line in left_lines
            ):
                return True
    return False


def _run_pdf2docx(source: Path, destination: Path, **settings: Any) -> None:
    from pdf2docx import Converter  # type: ignore

    converter = Converter(str(source))
    try:
        converter.convert(str(destination), **settings)
    finally:
        try:
            converter.close()
        except Exception:
            pass


def convert_pdf_to_word(source: Path, destination: Path) -> dict[str, Any]:
    _verify_source(source, PDF_EXTENSIONS)
    pages, scanned = _pdf_page_info(source)
    from importlib.util import find_spec

    if find_spec("pdf2docx") is None:
        raise ConversionError("PDF 转 Word 组件不可用。")

    try:
        _run_pdf2docx(source, destination)
        # pdf2docx may mistake a bullet list for a borderless table.  When
        # that happens, re-run without stream-table extraction so bullets
        # become ordinary, editable list paragraphs instead of a layout grid.
        if _contains_bullet_table(destination):
            destination.unlink(missing_ok=True)
            _run_pdf2docx(source, destination, parse_stream_table=False)
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(
            "PDF 转 Word 失败。文件可能使用了不受支持的编码或结构。",
            f"{type(exc).__name__}: {exc}",
        ) from exc

    _verify_docx(destination)
    warning = ""
    if scanned:
        warning = (
            "检测到扫描版 PDF：已按页面图片还原为 Word。"
            "若需要可编辑文字，请先使用 OCR 工具识别。"
        )
    return {"pages": pages, "warning": warning, "engine": "pdf2docx"}


def convert_pdf_to_editable_word(source: Path, destination: Path) -> dict[str, Any]:
    """Reflow PDF content into normal Word paragraphs, headings and tables."""
    _verify_source(source, PDF_EXTENSIONS)
    _pdf_page_info(source)  # validates encryption / page count early
    try:
        from pdf_flow import convert_pdf_to_flow_docx  # type: ignore

        info = convert_pdf_to_flow_docx(source, destination)
    except Exception as exc:
        raise ConversionError(
            "可编辑模式转换失败。文件可能包含特殊编码或复杂的图形结构。",
            f"{type(exc).__name__}: {exc}",
        ) from exc
    _verify_docx(destination)
    return info


def _verify_source(source: Path, allowed: set[str]) -> None:
    if not _is_file(source):
        raise ConversionError("原始文件不存在，可能已被移动或删除。")
    if source.suffix.casefold() not in allowed:
        readable = "、".join(sorted(ext.lstrip(".").upper() for ext in allowed))
        raise ConversionError(f"不支持的文件类型。此模式仅支持 {readable} 文件。")
    try:
        if source.stat().st_size == 0:
            raise ConversionError("原始文件为空，无法转换。")
    except OSError as exc:
        raise ConversionError("无法读取原始文件。", str(exc)) from exc


def convert_one(
    *,
    mode: str,
    source: str,
    output_dir: str,
    engine: dict[str, str] | None = None,
    pdf_mode: str = "flow",
) -> dict[str, Any]:
    """Convert one file and atomically place the result in output_dir."""
    started = time.perf_counter()
    source_path = Path(source)
    destination_dir = Path(output_dir)
    _verify_source(
        source_path, WORD_EXTENSIONS if mode == "word" else PDF_EXTENSIONS
    )
    _probe_output_directory(destination_dir)

    with tempfile.TemporaryDirectory(prefix="wenshu-") as temporary:
        work_dir = Path(temporary)
        temporary_output = work_dir / ("converted.pdf" if mode == "word" else "converted.docx")
        if mode == "word":
            if not engine:
                raise ConversionError("没有可用的 Word 转换引擎。")
            info = convert_word_to_pdf(
                source_path, temporary_output, engine, work_dir
            )
            extension = ".pdf"
        elif mode == "pdf":
            # Validate encryption / page count once and surface those errors
            # directly instead of treating them as a reflow failure.
            _pdf_page_info(source_path)
            if pdf_mode == "layout":
                info = convert_pdf_to_word(source_path, temporary_output)
            else:
                try:
                    info = convert_pdf_to_editable_word(source_path, temporary_output)
                except ConversionError as flow_error:
                    if temporary_output.exists():
                        try:
                            temporary_output.unlink()
                        except OSError:
                            pass
                    try:
                        info = convert_pdf_to_word(source_path, temporary_output)
                    except ConversionError as layout_error:
                        raise ConversionError(
                            "PDF 转 Word 失败。可编辑模式与保留版面模式均未能完成转换。",
                            f"可编辑模式: {flow_error.detail or flow_error}\n"
                            f"保留版面: {layout_error.detail or layout_error}",
                        ) from layout_error
                    info["warning"] = "可编辑模式未能完成转换，已自动改用保留原始版面模式。"
            extension = ".docx"
        else:
            raise ConversionError("未知的转换模式。")

        target = unique_output_path(destination_dir, source_path.stem, extension)
        try:
            shutil.move(str(temporary_output), str(target))
        except OSError as exc:
            raise ConversionError(
                f"转换完成，但无法保存到：{destination_dir}", str(exc)
            ) from exc

    elapsed = round(time.perf_counter() - started, 2)
    result: dict[str, Any] = {
        "ok": True,
        "output": str(target),
        "elapsed": elapsed,
        "mode": mode,
    }
    result.update(info)
    return result


def format_engine_summary(engine: dict[str, str] | None) -> str:
    if not engine:
        return "未检测到 Word 转换引擎"
    if engine.get("kind") == "com":
        return f"{engine.get('label', 'WPS/Word')} 已就绪"
    if engine.get("kind") == "soffice":
        return "系统 LibreOffice 已就绪"
    return "内置 LibreOffice 已就绪"


def prepare_logging() -> None:
    import warnings

    warnings.filterwarnings("ignore", category=FutureWarning, module="fitz")
    warnings.filterwarnings("ignore", message=".*fitz.*deprecated.*")
    logging.getLogger("pdf2docx").setLevel(logging.ERROR)
    logging.getLogger("fontTools").setLevel(logging.ERROR)
    try:
        import pymupdf as fitz  # type: ignore
    except Exception:
        try:
            import fitz  # type: ignore
        except Exception:
            return
    try:
        fitz.TOOLS.mupdf_display_errors(False)
        fitz.TOOLS.mupdf_display_warnings(False)
    except Exception:
        pass
