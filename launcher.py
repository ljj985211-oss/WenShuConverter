"""WenShu document converter - application entry point."""

from __future__ import annotations

import sys


def main() -> int:
    if "--convert-one" in sys.argv:
        index = sys.argv.index("--convert-one")
        from worker import worker_main

        return worker_main(sys.argv[index + 1 :])

    from app_ui import run_app

    return run_app()


if __name__ == "__main__":
    raise SystemExit(main())
