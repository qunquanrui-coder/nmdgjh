# -*- coding: utf-8 -*-

import os
import sys
import gc
import eel
import traceback
import multiprocessing
import tkinter as tk
from tkinter import filedialog
import threading
import pythoncom
from pathlib import Path
from functools import wraps


# ============================================================
# 基础路径：保证打包后无论从哪里启动，都能找到 web 目录
# ============================================================
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

WEB_DIR = BASE_DIR / "web"


# ============================================================
# 全局异常捕获
# ============================================================
def log_exception(exc_type, exc_value, exc_traceback):
    with open(BASE_DIR / "crash_log.txt", "w", encoding="utf-8") as f:
        f.write("".join(traceback.format_exception(exc_type, exc_value, exc_traceback)))
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


sys.excepthook = log_exception


# ============================================================
# 导入核心模块
# ============================================================
import core_blank_page
import core_pdf2word, core_split, core_word_split, core_word_merge
import core_unlock, core_compress, core_img2pdf, core_word2pdf
import core_pdf2img, core_invoice, core_diff
import core_ocr
import core_pdf_cleaner


# ============================================================
# 初始化 Eel
# 注意：这里必须用绝对路径，不能只写 "web"
# ============================================================
if not WEB_DIR.exists():
    raise FileNotFoundError(f"web 前端资源目录不存在: {WEB_DIR}")

if not (WEB_DIR / "index.html").exists():
    raise FileNotFoundError(f"web/index.html 不存在: {WEB_DIR / 'index.html'}")

eel.init(str(WEB_DIR))


# ============================================================
# 文件 / 文件夹选择
# ============================================================
@eel.expose
def ask_file():
    root = None
    try:
        root = tk.Tk()
        root.attributes("-topmost", True)
        root.withdraw()
        return filedialog.askopenfilename()
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


@eel.expose
def ask_folder():
    root = None
    try:
        root = tk.Tk()
        root.attributes("-topmost", True)
        root.withdraw()
        return filedialog.askdirectory()
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


# ============================================================
# COM 任务串行锁
# ============================================================
_com_lock = threading.Lock()

COM_TASKS = {
    "run_word_split",
    "get_word_outline",
    "run_word_merge",
    "run_word2pdf",
    "run_compress",
    "run_rm_blank",
}


def _normalize_result(res):
    if isinstance(res, dict):
        res.setdefault("status", "success")
        res.setdefault("msg", "")
        res.setdefault("data", None)
        return res
    return {"status": "success", "msg": "", "data": res}


def _run_function(func_name, func, *args, **kwargs):
    need_com = func_name in COM_TASKS
    com_initialized = False

    try:
        if need_com:
            pythoncom.CoInitialize()
            com_initialized = True

        result = func(*args, **kwargs)
        return _normalize_result(result)

    except Exception as e:
        traceback.print_exc()
        return {"status": "error", "msg": str(e), "data": None}

    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        gc.collect()


def execute_with_guard(func_name, func, *args, **kwargs):
    if func_name in COM_TASKS:
        with _com_lock:
            return _run_function(func_name, func, *args, **kwargs)
    return _run_function(func_name, func, *args, **kwargs)


def wrap_exposed_functions():
    for name, func in list(eel._exposed_functions.items()):
        if name.startswith("run_") or name == "get_word_outline":
            @wraps(func)
            def wrapper(*args, __fname=name, __func=func, **kwargs):
                return execute_with_guard(__fname, __func, *args, **kwargs)

            eel._exposed_functions[name] = wrapper


wrap_exposed_functions()


# ============================================================
# 窗口关闭回调
# ============================================================
def close_callback(route, websockets):
    if websockets:
        return
    gc.collect()
    raise SystemExit(0)


# ============================================================
# 启动逻辑：
# 1. 优先 Chrome
# 2. Chrome 不存在则自动 Edge
# ============================================================
def start_app():
    start_page = "index.html"
    common_kwargs = {
        "size": (1280, 850),
        "port": 0,
        "close_callback": close_callback,
    }

    try:
        eel.start(start_page, mode="chrome", **common_kwargs)
    except EnvironmentError:
        print("[WARN] 未找到 Chrome，自动切换到 Edge")
        eel.start(start_page, mode="edge", **common_kwargs)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    start_app()
