# -*- coding: utf-8 -*-

import os
import sys
import gc
import traceback
import multiprocessing
from pathlib import Path

import webview
import bridge  # 项目根目录下你自己的 bridge.py 桥接模块

import app_api

# 先导入桥接/API，再导入核心模块，让 @bridge.expose 注册生效
import core_blank_page
import core_pdf2word
import core_split
import core_word_split
import core_word_merge
import core_unlock
import core_compress
import core_img2pdf
import core_word2pdf
import core_pdf2img
import core_invoice
import core_diff
import core_ocr
import core_pdf_cleaner
import core_pdf_replace


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

WEB_DIR = BASE_DIR / "web"
INDEX_FILE = WEB_DIR / "index.html"


def log_exception(exc_type, exc_value, exc_traceback):
    try:
        with open(BASE_DIR / "crash_log.txt", "w", encoding="utf-8") as f:
            f.write("".join(traceback.format_exception(exc_type, exc_value, exc_traceback)))
    except Exception:
        pass

    sys.__excepthook__(exc_type, exc_value, exc_traceback)


sys.excepthook = log_exception


def ensure_frontend_assets() -> None:
    if not WEB_DIR.exists():
        raise FileNotFoundError(f"web 前端资源目录不存在: {WEB_DIR}")

    if not INDEX_FILE.exists():
        raise FileNotFoundError(f"web/index.html 不存在: {INDEX_FILE}")


def on_window_closed():
    gc.collect()


def start_app() -> None:
    ensure_frontend_assets()

    # 固定工作目录到 exe 所在目录，避免从别的路径启动时找不到资源
    os.chdir(BASE_DIR)

    api = app_api.AppApi()

    window = webview.create_window(
        title="泉泉的百宝箱",
        url=str(INDEX_FILE.resolve()),   # 关键：传绝对路径，不再传相对路径
        js_api=api,
        width=1280,
        height=850,
        min_size=(1120, 760),
        resizable=True,
        text_select=True,
    )

    bridge.set_window(window)
    window.events.closed += on_window_closed

    # 关键：显式启用本地 HTTP server，处理本地静态资源最稳
    webview.start(debug=False, http_server=True)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    start_app()
