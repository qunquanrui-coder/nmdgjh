# -*- coding: utf-8 -*-

import os
import re
import tempfile
import logging
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any

import eel
import pywintypes
import win32com.client.dynamic
import win32com

# --- [PDF 引擎支持] ---
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

# --- 强制重定向 COM 缓存目录 ---
_gen_path = Path(tempfile.gettempdir()) / "gen_py"
_gen_path.mkdir(parents=True, exist_ok=True)
win32com.__gen_path__ = str(_gen_path)


def is_page_strictly_blank(text: str) -> bool:
    """[Word 专用算法]"""
    if not text:
        return True
    cleaned = re.sub(r"[ \t\n\x0b\x0c\r\x0e\x0f\xa0\u3000\u200b]+", "", text)
    return len(cleaned) == 0


def safe_update_terminal(msg: str):
    logging.info(msg)
    try:
        eel.update_terminal(msg)
    except Exception:
        pass


def page_looks_visually_blank(
    page,
    dpi: int = 24,
    gray_threshold: int = 245,
    coverage_limit: float = 0.0012,
) -> bool:
    """
    视觉判空：
    - 低 DPI 灰度渲染
    - 统计非接近纯白像素占比
    - 占比很低，则视为“肉眼空白页”
    """
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, alpha=False)
    samples = pix.samples
    if not samples:
        return True

    dark_pixels = sum(1 for b in samples if b < gray_threshold)
    coverage = dark_pixels / len(samples)
    return coverage < coverage_limit


# --- [PDF 引擎：视觉判空版] ---
def process_pdf_core(pdf_path: Path) -> Tuple[str, List[int]]:
    if fitz is None:
        safe_update_terminal(f"❌ [{pdf_path.name}] 缺少pymupdf库")
        return "依赖缺失", []

    doc: Optional[fitz.Document] = None
    temp_path: Optional[str] = None

    try:
        try:
            doc = fitz.open(str(pdf_path))
        except Exception:
            return "文件受损/加密", []

        total_pages: int = len(doc)
        keep_indices: List[int] = []
        deleted_pages: List[int] = []

        for i in range(total_pages):
            if (i + 1) % 10 == 0 or i == total_pages - 1:
                safe_update_terminal(f" PDF引擎扫描: [{pdf_path.name}] {i+1}/{total_pages}...")

            page: fitz.Page = doc[i]

            # 先保留有表单/非链接注释的页，避免误删功能页
            try:
                raw_annots = page.annots()
                annots = [a for a in raw_annots if a.type[0] != 1] if raw_annots else []
                raw_widgets = page.widgets()
                widgets = list(raw_widgets) if raw_widgets else []
                if annots or widgets:
                    keep_indices.append(i)
                    continue
            except Exception:
                pass

            # 再做视觉判空
            try:
                if page_looks_visually_blank(page):
                    deleted_pages.append(i + 1)
                else:
                    keep_indices.append(i)
            except Exception as e:
                safe_update_terminal(f"⚠️ [{pdf_path.name}] 第 {i+1} 页视觉判空失败，默认保留: {e}")
                keep_indices.append(i)

        if not deleted_pages:
            doc.close()
            return "无需清理", []

        # 极端保护：不能删成空文件
        if not keep_indices:
            keep_indices = [0]
            if 1 in deleted_pages:
                deleted_pages.remove(1)

        doc.select(keep_indices)

        temp_fd, temp_path = tempfile.mkstemp(suffix=".pdf")
        os.close(temp_fd)

        try:
            doc.save(temp_path, garbage=2, deflate=True)
        finally:
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass
                doc = None

        os.replace(temp_path, str(pdf_path))
        return f"已清理 ({len(deleted_pages)}页)", deleted_pages

    except Exception as e:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass

        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

        safe_update_terminal(f"❌ [{pdf_path.name}] PDF引擎故障: {str(e)}")
        return "处理失败", []


# --- [Word 引擎] ---
def get_page_range(doc, page_num: int, total_pages: int):
    start_range = doc.GoTo(What=1, Which=1, Count=page_num)
    start = start_range.Start

    if page_num < total_pages:
        next_range = doc.GoTo(What=1, Which=1, Count=page_num + 1)
        end = next_range.Start - 1
    else:
        end = doc.Content.End

    return doc.Range(Start=start, End=end)


def process_word_core(word_path: Path) -> Tuple[str, List[int]]:
    word = None
    doc = None
    deleted_pages: List[int] = []

    try:
        word = win32com.client.dynamic.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(str(word_path))
        doc.Repaginate()

        try:
            total_pages = doc.ComputeStatistics(2)
        except Exception:
            total_pages = 1

        for page_num in range(total_pages, 0, -1):
            try:
                page_range = get_page_range(doc, page_num, total_pages)
                page_text = page_range.Text

                if is_page_strictly_blank(page_text):
                    page_range.Delete()
                    deleted_pages.append(page_num)
                    doc.Repaginate()
                    total_pages = doc.ComputeStatistics(2)
            except Exception as e:
                safe_update_terminal(f"⚠️ [{word_path.name}] 第 {page_num} 页判空失败，跳过: {e}")

        if deleted_pages:
            doc.Save()
            return f"已清理 ({len(deleted_pages)}页)", sorted(deleted_pages)
        else:
            return "无需清理", []

    except pywintypes.com_error as e:
        safe_update_terminal(f"❌ [{word_path.name}] Word引擎故障: {e}")
        return "处理失败", []

    except Exception as e:
        safe_update_terminal(f"❌ [{word_path.name}] Word引擎异常: {e}")
        return "处理失败", []

    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass

        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


def collect_target_files(target_path: Path) -> List[Path]:
    exts = {".doc", ".docx", ".pdf"}

    if target_path.is_file():
        return [target_path] if target_path.suffix.lower() in exts else []

    if target_path.is_dir():
        return sorted(
            [
                p
                for p in target_path.iterdir()
                if p.is_file() and p.suffix.lower() in exts and not p.name.startswith("~$")
            ],
            key=lambda x: x.name.lower(),
        )

    return []


@eel.expose
def run_rm_blank(target_path: str) -> Dict[str, Any]:
    try:
        root = Path(target_path.strip()).resolve()
        files = collect_target_files(root)

        if not files:
            return {"status": "error", "msg": "未找到可处理的 Word/PDF 文件", "data": None}

        total_deleted = 0
        processed = 0

        for f in files:
            processed += 1
            safe_update_terminal(f"[*] 开始处理: {f.name}")

            ext = f.suffix.lower()
            if ext == ".pdf":
                status_text, deleted_pages = process_pdf_core(f)
            else:
                status_text, deleted_pages = process_word_core(f)

            total_deleted += len(deleted_pages)

            if deleted_pages:
                safe_update_terminal(
                    f"[√] {f.name}: {status_text} | 删除页码: {', '.join(map(str, deleted_pages))}"
                )
            else:
                safe_update_terminal(f"[*] {f.name}: {status_text}")

        return {
            "status": "success",
            "msg": f"处理完成，共 {processed} 个文件，累计删除 {total_deleted} 页",
            "data": None,
        }

    except Exception as e:
        safe_update_terminal(f"❌ 空白页清理任务异常: {e}")
        return {"status": "error", "msg": str(e), "data": None}
