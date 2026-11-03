"""Isolated per-file conversion worker.

The GUI launches this module in a child process so an unstable PDF cannot
freeze the UI, and so the user can cancel a conversion cleanly.
"""

from __future__ import annotations

import json
import os
import sys
import traceback
from pathlib import Path

from conversion import ConversionError, convert_one, prepare_logging


def _write_response(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    os.replace(temporary, path)


def run_worker(request_path: Path) -> int:
    response_path: Path | None = None
    log_stream = None
    try:
        request = json.loads(request_path.read_text(encoding="utf-8"))
        response_path = Path(request["response_path"])
        log_path = str(request.get("log_path") or "")
        if log_path:
            try:
                log_stream = open(log_path, "a", encoding="utf-8", buffering=1)
                sys.stdout = log_stream
                sys.stderr = log_stream
            except OSError:
                log_stream = None
        if request.get("mode") == "pdf":
            prepare_logging()
        result = convert_one(
            mode=request["mode"],
            source=request["source"],
            output_dir=request["output_dir"],
            engine=request.get("engine"),
            pdf_mode=request.get("pdf_mode", "flow"),
        )
        _write_response(response_path, {"ok": True, "result": result})
        return 0
    except ConversionError as exc:
        payload = {
            "ok": False,
            "message": exc.message,
            "detail": exc.detail,
        }
    except PermissionError as exc:
        payload = {
            "ok": False,
            "message": "没有权限访问文件或输出文件夹。请换一个保存位置后重试。",
            "detail": f"{type(exc).__name__}: {exc}",
        }
    except MemoryError:
        payload = {
            "ok": False,
            "message": "内存不足，无法处理这个文件。",
            "detail": "MemoryError",
        }
    except Exception as exc:  # pragma: no cover - defensive boundary
        payload = {
            "ok": False,
            "message": f"转换时发生意外错误：{type(exc).__name__}",
            "detail": traceback.format_exc(),
        }

    if response_path is not None:
        try:
            _write_response(response_path, payload)
        except OSError:
            pass
    try:
        (request_path.parent / "worker-error.log").write_text(
            (payload.get("detail") or "") + "\n" + json.dumps(payload, ensure_ascii=False),
            encoding="utf-8",
        )
    except OSError:
        pass
    if log_stream is not None:
        try:
            log_stream.flush()
        except Exception:
            pass
    return 1


def worker_main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv:
        return 2
    return run_worker(Path(argv[0]))


if __name__ == "__main__":
    raise SystemExit(worker_main())
