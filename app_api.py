# -*- coding: utf-8 -*-

from __future__ import annotations

import gc
import threading
import traceback
import tkinter as tk
from tkinter import filedialog

import pythoncom
import eel


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


_com_lock = threading.Lock()

COM_TASKS = {
    "run_word_split",
    "get_word_outline",
    "run_word_merge",
    "run_word2pdf",
    "run_compress",
    "run_rm_blank",
}

RAW_RETURN_TASKS = {
    "ask_file",
    "ask_folder",
}


class AppApi:
    def invoke(self, func_name: str, args=None, kwargs=None):
        args = args or []
        kwargs = kwargs or {}

        func = eel._exposed_functions.get(func_name)
        if func is None:
            return {
                "status": "error",
                "msg": f"未找到接口: {func_name}",
                "data": None,
            }

        # 这些接口前端需要拿到原始值，不做统一包装
        if func_name in RAW_RETURN_TASKS:
            return self._run_raw_function(func_name, func, *args, **kwargs)

        return self._execute_with_guard(func_name, func, *args, **kwargs)

    @staticmethod
    def _run_raw_function(func_name, func, *args, **kwargs):
        need_com = func_name in COM_TASKS
        com_initialized = False

        try:
            if need_com:
                pythoncom.CoInitialize()
                com_initialized = True

            return func(*args, **kwargs)

        except Exception as e:
            traceback.print_exc()
            return ""

        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            gc.collect()

    @staticmethod
    def _normalize_result(res):
        if isinstance(res, dict):
            res.setdefault("status", "success")
            res.setdefault("msg", "")
            res.setdefault("data", None)
            return res

        return {
            "status": "success",
            "msg": "",
            "data": res,
        }

    @staticmethod
    def _run_function(func_name, func, *args, **kwargs):
        need_com = func_name in COM_TASKS
        com_initialized = False

        try:
            if need_com:
                pythoncom.CoInitialize()
                com_initialized = True

            result = func(*args, **kwargs)
            return AppApi._normalize_result(result)

        except Exception as e:
            traceback.print_exc()
            return {
                "status": "error",
                "msg": str(e),
                "data": None,
            }

        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            gc.collect()

    @staticmethod
    def _execute_with_guard(func_name, func, *args, **kwargs):
        if func_name in COM_TASKS:
            with _com_lock:
                return AppApi._run_function(func_name, func, *args, **kwargs)

        return AppApi._run_function(func_name, func, *args, **kwargs)
