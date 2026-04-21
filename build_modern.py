# -*- coding: utf-8 -*-
"""
build_modern.py
pywebview 版稳定打包脚本（Windows / Python 3.10）

目标：
1. 固定输出 dist/main/main.exe
2. 物理复制 web 前端资源到 dist/main/web
3. 自动补充 pywin32 DLL
4. 打包后物理复制 Ghostscript / runtime / poppler_bin
5. 尽量避免把 Qt / GTK / CEF 等不需要的 pywebview 后端打进来
6. 修复 GitHub Actions / Windows 控制台中文日志编码报错
"""

from __future__ import annotations

import os
import shutil
import site
import subprocess
import sys
from pathlib import Path
from typing import Optional


def setup_console_encoding() -> None:
    """
    修复 GitHub Actions / Windows 控制台下中文日志输出导致的 charmap 编码报错。
    """
    os.environ["PYTHONIOENCODING"] = "utf-8"
    os.environ["PYTHONUTF8"] = "1"

    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


if sys.platform == "win32":
    setup_console_encoding()


ROOT = Path(__file__).resolve().parent
ENTRY = ROOT / "main.py"
APP_NAME = "main"

DIST_DIR = ROOT / "dist"
BUILD_DIR = ROOT / "build"
SPEC_DIR = ROOT / "build_spec"

WEB_DIR = ROOT / "web"
ICON_FILE = ROOT / "toolbox_icon_clean.ico"

RUNTIME_DIRS = [
    ROOT / "Ghostscript",
    ROOT / "runtime",
    ROOT / "poppler_bin",
]


def log(msg: str) -> None:
    try:
        print(msg, flush=True)
    except UnicodeEncodeError:
        safe_msg = msg.encode("utf-8", errors="replace").decode("utf-8", errors="replace")
        print(safe_msg, flush=True)


def remove_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path, ignore_errors=True)
        log(f"[CLEAN] Removed: {path}")


def ensure_exists(path: Path, desc: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{desc} 不存在: {path}")


def find_pywin32_system32() -> Optional[Path]:
    try:
        import pywin32_system32  # type: ignore

        file_attr = getattr(pywin32_system32, "__file__", None)
        if file_attr:
            candidate = Path(file_attr).resolve().parent
            if candidate.exists():
                return candidate
    except Exception:
        pass

    candidates: list[Path] = []

    try:
        for p in site.getsitepackages():
            candidates.append(Path(p))
    except Exception:
        pass

    try:
        user_site = site.getusersitepackages()
        if user_site:
            candidates.append(Path(user_site))
    except Exception:
        pass

    for base in candidates:
        candidate = base / "pywin32_system32"
        if candidate.exists():
            return candidate

    return None


def build_pyinstaller_command() -> list[str]:
    cmd: list[str] = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--windowed",
        "--name",
        APP_NAME,
        "--distpath",
        str(DIST_DIR),
        "--workpath",
        str(BUILD_DIR),
        "--specpath",
        str(SPEC_DIR),

        # pywebview
        "--hidden-import=webview",
        "--hidden-import=webview.platforms.winforms",
        "--hidden-import=webview.platforms.edgechromium",
        "--hidden-import=webview.platforms.mshtml",
        "--collect-submodules=webview",
        "--collect-data=webview",

        # 精简 pywebview 不需要的平台后端
        "--exclude-module=PyQt5",
        "--exclude-module=PyQt6",
        "--exclude-module=PySide2",
        "--exclude-module=PySide6",
        "--exclude-module=qtpy",
        "--exclude-module=gi",
        "--exclude-module=cefpython3",
        "--exclude-module=webview.platforms.qt",
        "--exclude-module=webview.platforms.gtk",
        "--exclude-module=webview.platforms.cef",
        "--exclude-module=webview.platforms.cocoa",

        # 老 Eel websocket 依赖已不再需要
        "--exclude-module=bottle_websocket",
        "--exclude-module=websocket",

        # Windows / COM
        "--hidden-import=pythoncom",
        "--hidden-import=pywintypes",
        "--hidden-import=win32timezone",
        "--hidden-import=win32api",
        "--hidden-import=win32con",
        "--hidden-import=win32gui",
        "--hidden-import=win32com.client",
        "--hidden-import=comtypes",

        # 文档 / 图像 / OCR
        "--hidden-import=fitz",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=pdfplumber",
        "--hidden-import=docx",
        "--hidden-import=PIL",
        "--hidden-import=pdf2docx",
        "--hidden-import=img2pdf",
        "--hidden-import=pypdf",
        "--hidden-import=rapidocr_onnxruntime",
        "--hidden-import=cv2",
        "--hidden-import=numpy",
        "--hidden-import=ocrmypdf",
        "--hidden-import=pikepdf",
        "--hidden-import=pdfminer",
        "--hidden-import=pluggy",

        "--collect-submodules=pdf2docx",
        "--collect-submodules=pdfminer",
        "--collect-all=ocrmypdf",
        "--copy-metadata=ocrmypdf",
        "--copy-metadata=pikepdf",

        # 核心模块
        "--hidden-import=app_api",
        "--hidden-import=core_blank_page",
        "--hidden-import=core_compress",
        "--hidden-import=core_diff",
        "--hidden-import=core_img2pdf",
        "--hidden-import=core_invoice",
        "--hidden-import=core_ocr",
        "--hidden-import=core_pdf2img",
        "--hidden-import=core_pdf2word",
        "--hidden-import=core_pdf_cleaner",
        "--hidden-import=core_split",
        "--hidden-import=core_unlock",
        "--hidden-import=core_word2pdf",
        "--hidden-import=core_word_merge",
        "--hidden-import=core_word_split",
    ]

    pywin32_dll_dir = find_pywin32_system32()
    if pywin32_dll_dir:
        cmd.extend(["--add-binary", f"{pywin32_dll_dir};."])
        log(f"[OK] Added pywin32 DLLs: {pywin32_dll_dir}")
    else:
        log("[WARN] 未找到 pywin32_system32，若运行正常可忽略；否则请检查 pywin32 安装")

    if ICON_FILE.exists():
        cmd.extend(["--icon", str(ICON_FILE)])
        log(f"[OK] Added icon: {ICON_FILE}")

    cmd.append(str(ENTRY))
    return cmd


def copy_dir(src: Path, dst: Path, label: str) -> None:
    ensure_exists(src, f"{label} 源目录")

    if dst.exists():
        shutil.rmtree(dst, ignore_errors=True)

    shutil.copytree(src, dst)
    log(f"[COPY] {label} -> {dst}")


def copy_runtime_dirs(app_dir: Path) -> None:
    for src in RUNTIME_DIRS:
        copy_dir(src, app_dir / src.name, src.name)


def copy_web_dir(app_dir: Path) -> None:
    copy_dir(WEB_DIR, app_dir / "web", "web")


def verify_output(app_dir: Path) -> None:
    exe_path = app_dir / f"{APP_NAME}.exe"
    ensure_exists(exe_path, "打包输出 EXE")

    for src in RUNTIME_DIRS:
        ensure_exists(app_dir / src.name, f"已复制运行时目录 {src.name}")

    ensure_exists(app_dir / "web", "web 前端资源")
    ensure_exists(app_dir / "web" / "index.html", "web/index.html")

    log(f"[VERIFY] EXE exists: {exe_path}")
    log("[VERIFY] Runtime folders verified")
    log("[VERIFY] Web assets verified")
    log("[SUCCESS] 打包完成，可测试 dist/main/main.exe")


def build() -> None:
    ensure_exists(ENTRY, "入口文件 main.py")
    ensure_exists(WEB_DIR, "web 目录")

    remove_dir(BUILD_DIR)
    remove_dir(DIST_DIR)
    remove_dir(SPEC_DIR)

    cmd = build_pyinstaller_command()

    log("\n[START] Running PyInstaller...\n")
    try:
        subprocess.run(cmd, cwd=ROOT, check=True)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"PyInstaller 执行失败，退出码: {e.returncode}") from e

    app_dir = DIST_DIR / APP_NAME
    ensure_exists(app_dir, "dist 输出目录")

    log("\n[SYNC] Copy runtime folders...\n")
    copy_runtime_dirs(app_dir)

    log("\n[SYNC] Copy web assets...\n")
    copy_web_dir(app_dir)

    log("\n[CHECK] Verify package...\n")
    verify_output(app_dir)


if __name__ == "__main__":
    try:
        build()
    except Exception as e:
        log(f"\n[ERROR] {e}")
        sys.exit(1)
