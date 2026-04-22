# -*- coding: utf-8 -*-

import bridge
import time
import threading
import gc
import uuid
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client
import pywintypes


class ComManager:
    @classmethod
    def init(cls) -> None:
        pythoncom.CoInitialize()

    @classmethod
    def uninit(cls) -> None:
        pythoncom.CoUninitialize()


def _start_heartbeat(label: str, state: Dict[str, Any], interval: int = 5):
    stop_event = threading.Event()
    start_time = time.time()

    def worker() -> None:
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


def _stop_heartbeat(stop_event, thread) -> None:
    stop_event.set()
    try:
        thread.join(timeout=1)
    except Exception:
        pass


def _atomic_replace(tmp_path: Path, final_path: Path) -> bool:
    """原子替换文件，处理 Windows 文件占用冲突"""
    for _ in range(5):
        try:
            if final_path.exists():
                final_path.unlink()
            tmp_path.replace(final_path)
            return True
        except PermissionError:
            time.sleep(0.5)
        except OSError:
            time.sleep(0.5)
    return False


def _collect_input_files(root: Path, recursive: bool, exts: set[str]) -> List[Path]:
    if root.is_file():
        if root.suffix.lower() in exts and not root.name.startswith("~$"):
            return [root]
        return []

    if not root.is_dir():
        return []

    iterator = root.rglob("*") if recursive else root.iterdir()
    files = [
        p for p in iterator
        if p.is_file() and p.suffix.lower() in exts and not p.name.startswith("~$")
    ]
    files.sort(key=lambda x: str(x).lower())
    return files


def _safe_close_doc(doc: Any) -> None:
    if doc is not None:
        try:
            doc.Close(0)
        except Exception:
            pass


def _safe_close_wb(wb: Any) -> None:
    if wb is not None:
        try:
            wb.Close(False)
        except Exception:
            pass


def _safe_quit_word(word: Any) -> None:
    if word is not None:
        try:
            word.Quit(0)
        except Exception:
            pass


def _safe_quit_excel(excel: Any) -> None:
    if excel is not None:
        try:
            excel.Quit()
        except Exception:
            pass


def _safe_unlink(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except Exception:
        pass


def _export_word_document(doc: Any, tmp_out: Path, bookmark_mode: str, pdf_a: bool, progress_callback=None) -> None:
    if progress_callback:
        progress_callback("更新目录")
    try:
        if doc.TablesOfContents.Count > 0:
            for toc in doc.TablesOfContents:
                toc.Update()
    except Exception:
        pass

    if progress_callback:
        progress_callback("重新分页")
    try:
        doc.Repaginate()
    except Exception:
        pass

    if progress_callback:
        progress_callback("导出 PDF")
    doc.ExportAsFixedFormat(
        str(tmp_out),
        17,  # wdFormatPDF
        Item=0,
        CreateBookmarks={"标题": 1, "Word书签": 2}.get(bookmark_mode, 0),
        DocStructureTags=True,
        UseISO19005_1=pdf_a,
    )


@bridge.expose
def run_word2pdf(
    input_path: str,
    recursive: bool,
    include_doc: bool,
    include_xls: bool,
    bookmark_mode: str,
    pdf_a: bool,
) -> Dict[str, Any]:
    """
    批量将 Word/Excel 文件转换为 PDF。
    包含针对 Word 导出末尾空白页的专项修复。
    """

    ComManager.init()

    word: Optional[Any] = None
    excel: Optional[Any] = None
    doc: Optional[Any] = None
    wb: Optional[Any] = None

    try:
        root = Path(input_path.strip())
        if not root.exists():
            raise FileNotFoundError(f"输入路径不存在: {root}")

        exts = {".docx"}
        if include_doc:
            exts.add(".doc")
        if include_xls:
            exts.update({".xlsx", ".xls"})

        files = _collect_input_files(root, recursive, exts)
        if not files:
            return {
                "status": "error",
                "msg": "未找到可转换的 Word / Excel 文件",
                "data": None,
            }

        has_word = any(p.suffix.lower() in (".doc", ".docx") for p in files)
        has_excel = any(p.suffix.lower() in (".xls", ".xlsx") for p in files)

        if has_word:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = 0
            word.DisplayAlerts = 0
            try:
                word.Options.UsePrinterMetrics = True
                word.Options.Pagination = True
            except Exception:
                pass

        if has_excel:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = 0
            excel.DisplayAlerts = 0

        success_count = 0
        failed_items: List[str] = []

        for i, src in enumerate(files, 1):
            tmp_out: Optional[Path] = None
            keep_tmp = False
            heartbeat_state = {"stage": "准备转换", "current": i, "total": len(files)}
            stop_event, heartbeat_thread = _start_heartbeat(src.name, heartbeat_state)

            bridge.update_terminal(f"[*] 转换中 ({i}/{len(files)}): {src.name}")

            out = src.with_suffix(".pdf")
            tmp_out = out.with_name(f"{out.stem}__tmp__{uuid.uuid4().hex[:6]}.pdf")

            try:
                if src.suffix.lower() in (".doc", ".docx"):
                    heartbeat_state["stage"] = "打开 Word 文档"
                    doc = word.Documents.Open(
                        str(src.resolve()),
                        ReadOnly=True,
                        AddToRecentFiles=False,
                    )
                    _export_word_document(
                        doc,
                        tmp_out,
                        bookmark_mode,
                        pdf_a,
                        lambda stage: heartbeat_state.update({"stage": stage}),
                    )
                    _safe_close_doc(doc)
                    doc = None

                else:
                    heartbeat_state["stage"] = "打开 Excel 工作簿"
                    wb = excel.Workbooks.Open(
                        str(src.resolve()),
                        UpdateLinks=0,
                        ReadOnly=True,
                    )
                    heartbeat_state["stage"] = "导出 Excel PDF"
                    wb.ExportAsFixedFormat(
                        0,  # xlTypePDF
                        str(tmp_out),
                        Quality=0,
                        IgnorePrintAreas=False,
                    )
                    _safe_close_wb(wb)
                    wb = None

                heartbeat_state["stage"] = "替换输出文件"
                if _atomic_replace(tmp_out, out):
                    bridge.update_terminal(f"[√] 完成: {out.name}")
                    success_count += 1
                else:
                    keep_tmp = True
                    bridge.update_terminal(f"[!] 文件占用，已保留临时结果: {tmp_out.name}")
                    success_count += 1

            except pywintypes.com_error as e:
                bridge.update_terminal(f"❌ 失败 {src.name}: COM 交互错误 {e}")
                failed_items.append(src.name)

            except PermissionError as e:
                bridge.update_terminal(f"❌ 失败 {src.name}: 权限不足 {e}")
                failed_items.append(src.name)

            except Exception as e:
                bridge.update_terminal(f"❌ 失败 {src.name}: {e}")
                failed_items.append(src.name)

            finally:
                _stop_heartbeat(stop_event, heartbeat_thread)
                _safe_close_doc(doc)
                _safe_close_wb(wb)
                doc = None
                wb = None

                if tmp_out is not None and tmp_out.exists() and not keep_tmp:
                    _safe_unlink(tmp_out)

                if i % 10 == 0:
                    gc.collect()

        if failed_items and success_count == 0:
            return {
                "status": "error",
                "msg": f"全部转换失败，共 {len(failed_items)} 个文件。请查看终端日志。",
                "data": {"success": success_count, "failed": failed_items},
            }

        if failed_items:
            return {
                "status": "error",
                "msg": f"部分转换完成：成功 {success_count} 个，失败 {len(failed_items)} 个。请查看终端日志。",
                "data": {"success": success_count, "failed": failed_items},
            }

        return {"status": "success", "msg": f"共成功转换 {success_count} 个文件", "data": None}

    except Exception as e:
        error_msg = traceback.format_exc()
        bridge.update_terminal(f"❌ 严重错误: {str(e)}")
        return {"status": "error", "msg": error_msg, "data": None}

    finally:
        _safe_close_doc(doc)
        _safe_close_wb(wb)
        _safe_quit_word(word)
        _safe_quit_excel(excel)
        ComManager.uninit()
