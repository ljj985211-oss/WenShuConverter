"""Assemble the portable LibreOffice Kit engine used by Word -> PDF.

The app uses the prebuilt LibreOffice Kit that ships with DSH. This script
copies Node and the package dependency closure into the engine folder so
the packaged application does not depend on the DSH installation.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENGINE = ROOT / "engine"
DSH_RESOURCES = Path(os.environ.get("DSH_RESOURCES", r"D:\DeepSeek\resources"))
DSH_MODULES = DSH_RESOURCES / "app.asar.unpacked" / "dsh" / "node_modules"
DSH_NODE = (
    DSH_RESOURCES
    / "runtime"
    / "primary-runtime"
    / "dependencies"
    / "node"
    / "bin"
    / "node.exe"
)

START_PACKAGES = ["@deepseek-ai/libreoffice-kit"]
PLATFORM_PACKAGE = "@deepseek-ai/libreoffice-kit-win32-x64"


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def package_closure() -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    queue = list(START_PACKAGES)
    while queue:
        name = queue.pop(0)
        if name in seen:
            continue
        package_dir = DSH_MODULES / Path(name)
        manifest = package_dir / "package.json"
        if not manifest.exists():
            print(f"WARNING: missing package {name}", file=sys.stderr)
            continue
        seen.add(name)
        ordered.append(name)
        data = read_json(manifest)
        for section in ("dependencies", "optionalDependencies"):
            for dependency in (data.get(section) or {}):
                if dependency not in seen:
                    queue.append(dependency)
    return ordered


def copy_package(name: str) -> None:
    source = DSH_MODULES / Path(name)
    target = ENGINE / "node_modules" / Path(name)
    if target.exists():
        shutil.rmtree(target, ignore_errors=True)
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, target)


def main() -> int:
    if not DSH_MODULES.is_dir():
        print(f"LibreOffice Kit modules not found: {DSH_MODULES}", file=sys.stderr)
        return 1
    if not DSH_NODE.is_file():
        print(f"Node executable not found: {DSH_NODE}", file=sys.stderr)
        return 1
    if not (DSH_MODULES / Path(PLATFORM_PACKAGE) / "package.json").is_file():
        print(f"Native engine package not found: {PLATFORM_PACKAGE}", file=sys.stderr)
        return 1

    if ENGINE.exists():
        shutil.rmtree(ENGINE)
    (ENGINE / "node").mkdir(parents=True)
    (ENGINE / "node_modules").mkdir(parents=True)
    shutil.copy2(DSH_NODE, ENGINE / "node" / "node.exe")

    packages = package_closure()
    for index, name in enumerate(packages, start=1):
        print(f"[{index}/{len(packages)}] {name}")
        copy_package(name)

    native = DSH_MODULES / Path(PLATFORM_PACKAGE)
    versions = {
        "node": "v24.21.0",
        "libreofficeKit": read_json(
            DSH_MODULES / "@deepseek-ai" / "libreoffice-kit" / "package.json"
        ).get("version"),
        "nativeEngine": read_json(native / "package.json").get("version"),
        "license": "MPL-2.0",
    }
    (ENGINE / "version.json").write_text(
        json.dumps(versions, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    (ENGINE / "README.txt").write_text(
        "Portable LibreOffice Kit engine for WenShu Converter.\n"
        "Includes Node.js and the prebuilt LibreOffice Kit, licensed under MPL-2.0.\n"
        "See node_modules/@deepseek-ai/libreoffice-kit/LICENSE and the licenses\n"
        "directory inside the native engine package for third-party notices.\n",
        encoding="utf-8",
    )
    total = sum(p.stat().st_size for p in ENGINE.rglob("*") if p.is_file())
    print(f"Engine assembled: {ENGINE} ({total / (1024 * 1024):.1f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
