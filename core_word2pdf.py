# -*- coding: utf-8 -*-
import eel
import time
import gc
import uuid
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client
import pywintypes

# 搬运原始 ComManager
class ComManager:
    @classmethod
    def init(cls) -> None:
        pythoncom.CoInitialize()

    @classmethod
    def uninit(cls) -> None:
        pythoncom.CoUninitialize()


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


@eel.expose
def run_word2pdf(
    input_path: str,
    recursive: bool,
    include_doc: bool,
    include_xls: bool,
    bookmark_mode: str,
    pdf_a: bool
) -> Dict[str, Any]:
    """
    批量将 Word/Excel 文件转换为 PDF。
    包含针对 Word 导出末尾空白页的专项修复。
    """
    ComManager.init()
    word: Optional[win32com.client.CDispatch] = None
    excel: Optional[win32com.client.CDispatch] = None
    doc: Optional[win32com.client.CDispatch] = None
    wb: Optional[win32com.client.CDispatch] = None

    try:
        root: Path = Path(input_path.strip())
        if not root.exists():
            raise FileNotFoundError(f"输入路径不存在: {root}")

        exts: set[str] = {".docx"}
        if include_doc:
            exts.add(".doc")
        if include_xls:
            exts.update({".xlsx", ".xls"})

        files: List[Path] = [
            p for p in (root.rglob("*") if recursive else root.iterdir())
            if p.suffix.lower() in exts and not p.name.startswith("~$")
        ]
        files.sort(key=lambda x: str(x).lower())

        if any(p.suffix.lower() in (".doc", ".docx") for p in files):
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = 0
            word.DisplayAlerts = 0
            try:
                word.Options.UsePrinterMetrics = True
                word.Options.Pagination = True
            except Exception:
                pass

        if any(p.suffix.lower() in (".xls", ".xlsx") for p in files):
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = 0
            excel.DisplayAlerts = 0

        for i, src in enumerate(files, 1):
            eel.update_terminal(f"[*] 转换中 ({i}/{len(files)}): {src.name}")
            out: Path = src.with_suffix(".pdf")
            tmp_out: Path = out.with_name(f"{out.stem}__tmp__{uuid.uuid4().hex[:6]}.pdf")

            try:
                if src.suffix.lower() in (".doc", ".docx"):
                    doc = word.Documents.Open(
                        str(src.resolve()), ReadOnly=True, AddToRecentFiles=False
                    )
                    if doc.TablesOfContents.Count > 0:
                        for toc in doc.TablesOfContents:
                            toc.Update()

                    # ✨ 核心修复点：强制刷新排版以消除分页滞后带来的空白页
                    try:
                        doc.Repaginate()
                    except Exception:
                        pass

                    doc.ExportAsFixedFormat(
                        str(tmp_out),
                        17,  # wdFormatPDF
                        Item=0,  # ✨ 核心修复点：只导出正文，剔除批注和隐藏标记撑开的空白页
                        CreateBookmarks={"标题": 1, "Word书签": 2}.get(bookmark_mode, 0),
                        DocStructureTags=True,
                        UseISO19005_1=pdf_a
                    )
                    doc.Close(0)
                    doc = None  # 释放引用
                else:
                    wb = excel.Workbooks.Open(
                        str(src.resolve()), UpdateLinks=0, ReadOnly=True
                    )
                    wb.ExportAsFixedFormat(
                        0, str(tmp_out), Quality=0, IgnorePrintAreas=False
                    )
                    wb.Close(False)
                    wb = None  # 释放引用

                if not _atomic_replace(tmp_out, out):
                    eel.update_terminal(f"[!] 占用, 已存为: {tmp_out.name}")
            except pywintypes.com_error as e:
                eel.update_terminal(f"❌ 失败 {src.name}: COM 交互错误 {e}")
            except PermissionError as e:
                eel.update_terminal(f"❌ 失败 {src.name}: 权限不足 {e}")
            except Exception as e:
                eel.update_terminal(f"❌ 失败 {src.name}: {e}")
            finally:
                # 确保当前循环的 COM 对象被显式关闭，防止句柄泄漏
                if doc is not None:
                    try:
                        doc.Close(0)
                    except Exception:
                        pass
                    doc = None
                if wb is not None:
                    try:
                        wb.Close(False)
                    except Exception:
                        pass
                    wb = None

            if i % 10 == 0:
                gc.collect()

        return {"status": "success"}
    except Exception as e:
        error_msg = traceback.format_exc()
        eel.update_terminal(f"❌ 严重错误: {str(e)}")
        return {"status": "error", "msg": error_msg}
    finally:
        if word:
            try:
                word.Quit(0)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        ComManager.uninit()
