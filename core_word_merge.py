# -*- coding: utf-8 -*-
import ctypes
import re
import tempfile
import threading
import time
from pathlib import Path, PureWindowsPath

import pythoncom
import win32com.client as win32

import bridge
from safe_output import publish_output


def _start_heartbeat(label, state, interval=5):
    stop_event = threading.Event()
    start_time = time.time()

    def worker():
        while not stop_event.wait(interval):
            stage = state.get("stage", "处理中")
            current = state.get("current")
            total = state.get("total")
            elapsed = int(time.time() - start_time)
            if current is not None and total:
                bridge.update_terminal(f"⏳ [{label}] {stage}: {current}/{total}，已耗时 {elapsed}s")
            else:
                bridge.update_terminal(f"⏳ [{label}] {stage}，已耗时 {elapsed}s")

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    return stop_event, thread


def _stop_heartbeat(stop_event, thread):
    stop_event.set()
    try:
        thread.join(timeout=1)
    except Exception:
        pass


@bridge.expose
def run_word_merge(folder, out_name):
    com_initialized = False
    word = None
    merged_doc = None
    state = {"stage": "启动 Word 引擎", "current": None, "total": None}
    stop_event, heartbeat_thread = _start_heartbeat("Word 合并", state)
    try:
        if not folder or not str(folder).strip():
            return {"status": "error", "msg": "请选择包含 Word 文件的文件夹"}
        root = Path(str(folder).strip()).resolve()
        if not root.is_dir():
            return {"status": "error", "msg": "请选择有效的 Word 文件夹"}

        out_name = str(out_name or "").strip()
        if (not out_name or out_name.endswith(".") or re.search(r'[\x00-\x1f<>:"/\\|?*]', out_name)
                or PureWindowsPath(out_name).is_reserved()):
            return {"status": "error", "msg": "请输入有效的输出文件名，不能包含目录或路径"}
        if Path(out_name).suffix.lower() != ".docx":
            out_name += ".docx"
        desired_output = root / out_name
        merge_output_names = re.compile(
            rf"{re.escape(desired_output.stem)}(?: \(\d+\))?{re.escape(desired_output.suffix)}",
            re.IGNORECASE,
        )
        files = [
            f for f in root.glob("*.doc*")
            if f.is_file() and not f.name.startswith("~$") and not merge_output_names.fullmatch(f.name)
        ]
        files.sort(key=lambda p: [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', p.name)])
        if not files:
            return {"status": "error", "msg": "未找到可合并的 Word 文件"}

        state["stage"] = "创建 Word 进程"
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_initialized = True
        word = win32.DispatchEx("Word.Application")
        word.Visible = word.ScreenUpdating = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3  # 禁止自动打开的文档运行宏

        state["stage"] = "创建合并文档"
        merged_doc = word.Documents.Add()
        state.update({"stage": "合并文档", "current": 0, "total": len(files)})

        for idx, src in enumerate(files, 1):
            state.update({"stage": f"打开源文档 {src.name}", "current": idx, "total": len(files)})
            bridge.update_terminal(f"[{idx}/{len(files)}] 正在合并: {src.stem}")
            if idx > 0:
                merged_doc.Range(merged_doc.Content.End - 1).InsertBreak(7)

            rng = merged_doc.Range(merged_doc.Content.End - 1)
            rng.Text = f"{src.stem}\n"
            try:
                rng.Style = "标题 1"
            except Exception:
                pass

            source_doc = None
            try:
                source_doc = word.Documents.Open(str(src.resolve()), ReadOnly=True, Visible=False)
                state["stage"] = f"复制粘贴 {src.name}"
                pasted = False
                for attempt in range(1, 4):
                    try:
                        ctypes.windll.user32.OpenClipboard(None)
                        ctypes.windll.user32.EmptyClipboard()
                        ctypes.windll.user32.CloseClipboard()
                        source_doc.Content.Copy()
                        merged_doc.Range(merged_doc.Content.End - 1).PasteAndFormat(16)
                        pasted = True
                        break
                    except Exception as e:
                        bridge.update_terminal(f"  └─ 第 {attempt}/3 次粘贴重试: {e}")
                        time.sleep(0.5)
                if not pasted:
                    raise RuntimeError(f"{src.name} 内容复制失败")
            finally:
                if source_doc is not None:
                    try:
                        source_doc.Close(False)
                    except Exception:
                        pass

        state.update({"stage": "保存合并文档", "current": None, "total": None})
        bridge.update_terminal(f"[*] 正在保存合并文档: {out_name}")
        with tempfile.TemporaryDirectory(prefix=".word-merge-", dir=root) as tmp_dir:
            staged_output = Path(tmp_dir) / "merged.docx"
            try:
                merged_doc.SaveAs2(str(staged_output), 16)
            finally:
                merged_doc.Close(False)
                merged_doc = None
            output = publish_output(staged_output, desired_output)
        return {"status": "success", "msg": f"已生成: {output.name}"}
    except Exception as e:
        return {"status": "error", "msg": str(e)}
    finally:
        _stop_heartbeat(stop_event, heartbeat_thread)
        if merged_doc is not None:
            try:
                merged_doc.Close(False)
            except Exception:
                pass
        try:
            if word:
                word.Quit()
        finally:
            if com_initialized:
                pythoncom.CoUninitialize()
