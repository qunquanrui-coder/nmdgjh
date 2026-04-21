# -*- coding: utf-8 -*-
import eel
import sys
import os
import traceback
import multiprocessing
import tkinter as tk
from tkinter import filedialog

# --- ✨ 全局异常捕获 ---
def log_exception(exc_type, exc_value, exc_traceback):
    with open("crash_log.txt", "w", encoding="utf-8") as f:
        f.write("".join(traceback.format_exception(exc_type, exc_value, exc_traceback)))
    sys.__excepthook__(exc_type, exc_value, exc_traceback)

sys.excepthook = log_exception

# 导入所有功能核心模块
import core_blank_page
import core_pdf2word, core_split, core_word_split, core_word_merge
import core_unlock, core_compress, core_img2pdf, core_word2pdf
import core_pdf2img, core_invoice, core_diff
import core_ocr  
import core_pdf_cleaner  

# 初始化 Web 资源目录
eel.init('web')

@eel.expose
def ask_file():
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    path = filedialog.askopenfilename()
    root.destroy()
    return path

@eel.expose
def ask_folder():
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    path = filedialog.askdirectory()
    root.destroy()
    return path

# =====================================================================
# === 👇 COM 线程隔离与防假死引擎 (带白名单机制的全局并发排队锁) 👇 ====
# =====================================================================
import threading
import pythoncom
import gc

# 全局 COM 任务排队锁
_com_lock = threading.Lock()

# ✨ 核心修复：只有在这个名单里的任务，才会去拿锁排队（它们会调用底层 Office COM）。
# 其他不依赖 COM 的任务（如 OCR、PDF 拆分等）直接放行，全速并发！
COM_TASKS = {
    'run_word_split', 
    'get_word_outline', 
    'run_word_merge', 
    'run_word2pdf', 
    'run_compress', 
    'run_rm_blank'
}

def execute_with_com(func_name, func, *args, **kwargs):
    result = {"status": "pending", "msg": "", "data": None}
    
    def worker():
        # 封装实际执行的核心逻辑
        def _run_core():
            try:
                pythoncom.CoInitialize()
                res = func(*args, **kwargs)
                if isinstance(res, dict):
                    result.update(res)
                else:
                    result["status"] = "success"
                    result["data"] = res
            except Exception as e:
                result["status"] = "error"
                result["msg"] = str(e)
                traceback.print_exc()
            finally:
                pythoncom.CoUninitialize()
                gc.collect()

        # 智能判定：是否为危险的 COM 任务？
        if func_name in COM_TASKS:
            with _com_lock:  # 危险任务，老老实实拿锁排队
                _run_core()
        else:
            _run_core()      # 安全任务，直接并发执行

    t = threading.Thread(target=worker)
    t.start()
    t.join()
    return result

# 🚀 魔法注入：自动拦截并按名字分发给 wrapper
for name, func in list(eel._exposed_functions.items()):
    if name.startswith('run_') or name == 'get_word_outline':
        # 通过 fname 参数将函数名传给闭包
        def create_wrapper(fname, original_func):
            def wrapper(*args, **kwargs):
                return execute_with_com(fname, original_func, *args, **kwargs)
            return wrapper
        eel._exposed_functions[name] = create_wrapper(name, func)

# =====================================================================

def close_callback(route, websockets):
    if not websockets:
        os._exit(0)

if __name__ == '__main__':
    multiprocessing.freeze_support()
    eel.start('index.html', mode='chrome', size=(1280, 850), port=0, close_callback=close_callback)
