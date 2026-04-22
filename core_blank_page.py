# -*- coding: utf-8 -*-
"""
空白页清理模块（最终封版）

设计原则：
1. Word：保留原有判空风格，采用“边处理边重新分页”的动态策略，避免页码漂移。
2. Word：只自动修复更安全的“手动分页符导致的空白页”。
3. Word：分节符 / 分栏符 / 奇偶页分节导致的版式空白页，仅识别并保留。
4. PDF：采用保守判空策略，优先避免误删。
5. 修改前自动备份，任何备份失败都直接跳过处理。
"""

import os
import re
import tempfile
import logging
import shutil
import threading
import time
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any

import bridge
import pywintypes
import win32com.client.dynamic
import win32com

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


# --- 强制重定向 COM 缓存目录 ---
_gen_path = Path(tempfile.gettempdir()) / "gen_py"
_gen_path.mkdir(parents=True, exist_ok=True)
win32com.__gen_path__ = str(_gen_path)


def is_page_strictly_blank(text: str) -> bool:
    """Word 专用判空算法：沿用原始规则，不改变识别风格。"""
    if not text:
        return True
    cleaned = re.sub(r"[ \t\n\x0b\x0c\r\x0e\x0f\xa0\u3000\u200b]+", "", text)
    return len(cleaned) == 0


def clean_visible_text(text: str) -> str:
    return re.sub(r"[ \t\n\x0b\x0c\r\x0e\x0f\xa0\u3000\u200b]+", "", text or "")


def safe_update_terminal(msg: str):
    logging.info(msg)
    try:
        bridge.update_terminal(msg)
    except Exception:
        pass


def make_backup_before_overwrite(path: Path) -> Tuple[bool, Optional[Path], Optional[str]]:
    backup_path = path.with_name(f"{path.stem}.blankbak{path.suffix}")
    if backup_path.exists():
        backup_path = path.with_name(f"{path.stem}.blankbak_{int(time.time())}{path.suffix}")
    try:
        shutil.copy2(path, backup_path)
        return True, backup_path, None
    except Exception as e:
        return False, None, str(e)


def start_heartbeat(file_name: str, state: Dict[str, Any]) -> Tuple[threading.Event, threading.Thread]:
    stop_event = threading.Event()
    start_time = time.time()

    def worker():
        while not stop_event.wait(5):
            stage = state.get("stage", "处理中")
            current = state.get("current")
            total = state.get("total")
            elapsed = int(time.time() - start_time)

            if current is not None and total:
                safe_update_terminal(f"⏳ [{file_name}] {stage}: {current}/{total}，已耗时 {elapsed}s")
            else:
                safe_update_terminal(f"⏳ [{file_name}] {stage}，已耗时 {elapsed}s")

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    return stop_event, thread


# =========================
# PDF 引擎
# =========================

def page_looks_visually_blank(
    page,
    dpi: int = 48,
    gray_threshold: int = 235,
    coverage_limit: float = 0.0002,
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



def pdf_page_has_structural_content(page) -> Tuple[bool, str]:
    """PDF 保守判空：任何明确内容信号都保留，避免误删。"""
    try:
        if clean_visible_text(page.get_text("text")):
            return True, "文本"
    except Exception:
        return True, "文本检测异常"

    try:
        raw_annots = page.annots()
        if raw_annots and list(raw_annots):
            return True, "注释"
    except Exception:
        pass

    try:
        raw_widgets = page.widgets()
        if raw_widgets and list(raw_widgets):
            return True, "表单"
    except Exception:
        pass

    try:
        if page.get_images(full=True):
            return True, "图片"
    except Exception:
        pass

    try:
        if page.get_drawings():
            return True, "矢量绘图"
    except Exception:
        pass

    return False, ""



def pdf_page_is_blank(page) -> Tuple[bool, str]:
    has_content, reason = pdf_page_has_structural_content(page)
    if has_content:
        return False, reason

    try:
        if page_looks_visually_blank(page):
            return True, "视觉空白"
        return False, "视觉存在内容"
    except Exception as e:
        return False, f"视觉检测异常: {e}"



def too_many_blank_candidates(blank_pages: List[int], total_pages: int) -> bool:
    if not blank_pages:
        return False
    limit = max(3, int(total_pages * 0.10))
    return len(blank_pages) > limit



def process_pdf_core(pdf_path: Path) -> Tuple[str, List[int]]:
    if fitz is None:
        safe_update_terminal(f"❌ [{pdf_path.name}] 缺少 PyMuPDF 库")
        return "依赖缺失", []

    doc: Optional["fitz.Document"] = None
    temp_path: Optional[str] = None

    try:
        try:
            doc = fitz.open(str(pdf_path))
        except Exception:
            return "文件受损/加密", []

        total_pages = len(doc)
        keep_indices: List[int] = []
        deleted_pages: List[int] = []

        for i in range(total_pages):
            if (i + 1) % 10 == 0 or i == total_pages - 1:
                safe_update_terminal(f"[*] PDF 引擎扫描: [{pdf_path.name}] {i + 1}/{total_pages}...")

            page = doc[i]
            is_blank, _ = pdf_page_is_blank(page)
            if is_blank:
                deleted_pages.append(i + 1)
            else:
                keep_indices.append(i)

        if too_many_blank_candidates(deleted_pages, total_pages):
            safe_update_terminal(
                f"⚠️ [{pdf_path.name}] PDF 空白候选页过多 ({len(deleted_pages)}/{total_pages})，"
                f"为避免误删已跳过。候选页: {', '.join(map(str, deleted_pages[:30]))}"
            )
            doc.close()
            return "候选过多，已跳过", []

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

        ok, backup_path, backup_error = make_backup_before_overwrite(pdf_path)
        if not ok:
            safe_update_terminal(f"⚠️ [{pdf_path.name}] 备份原文件失败，已取消删除以避免不可逆误删: {backup_error}")
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
            return "备份失败，已跳过", []

        safe_update_terminal(f"[*] [{pdf_path.name}] 已备份原文件: {backup_path.name}")
        safe_update_terminal(f"[*] [{pdf_path.name}] PDF 保守判空删除页: {', '.join(map(str, deleted_pages))}")
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

        safe_update_terminal(f"❌ [{pdf_path.name}] PDF 引擎故障: {e}")
        return "处理失败", []


# =========================
# Word 引擎
# =========================

def get_page_range(doc, page_num: int, total_pages: int):
    start = doc.GoTo(What=1, Which=1, Count=page_num).Start
    end = doc.GoTo(What=1, Which=1, Count=page_num + 1).Start if page_num < total_pages else doc.Content.End
    return doc.Range(Start=start, End=end)



def word_range_find(page_range, pattern: str) -> bool:
    try:
        rng = page_range.Duplicate
        finder = rng.Find
        finder.ClearFormatting()
        finder.Text = pattern
        finder.Forward = True
        finder.Wrap = 0
        return bool(finder.Execute())
    except Exception:
        return False



def get_prev_char(doc, pos: int) -> str:
    if pos <= 0:
        return ""
    try:
        return doc.Range(pos - 1, pos).Text or ""
    except Exception:
        return ""



def detect_layout_blank_reason(doc, page_range, page_text: str) -> str:
    """
    仅识别造成空白页的版式控制原因。
    自动处理只允许“手动分页符”。
    """
    prev_char = get_prev_char(doc, page_range.Start)

    if word_range_find(page_range, "^m") or prev_char == "\x0c" or chr(12) in (page_text or ""):
        return "manual_page_break"

    if word_range_find(page_range, "^b") or prev_char == "\x0f":
        return "section_break"

    if word_range_find(page_range, "^n") or prev_char == "\x0e":
        return "column_break"

    try:
        for i in range(1, page_range.Sections.Count + 1):
            section_start = page_range.Sections(i).PageSetup.SectionStart
            if section_start in (3, 4):
                return "odd_even_section_break"
    except Exception:
        pass

    return ""



def delete_manual_page_breaks_in_range(page_range) -> int:
    deleted = 0
    for _ in range(50):
        try:
            rng = page_range.Duplicate
            finder = rng.Find
            finder.ClearFormatting()
            finder.Text = "^m"
            finder.Forward = True
            finder.Wrap = 0
            if not finder.Execute():
                break
            rng.Delete()
            deleted += 1
        except Exception:
            break
    return deleted



def fix_manual_layout_blank_page(doc, page_range) -> Tuple[bool, str]:
    """
    只处理更安全的“手动分页符”空白页：
    - 删除页内 ^m
    - 或删除页首前一个 \x0c
    """
    deleted = delete_manual_page_breaks_in_range(page_range)

    try:
        if page_range.Start > 0:
            prev_rng = doc.Range(page_range.Start - 1, page_range.Start)
            if prev_rng.Text == "\x0c":
                prev_rng.Delete()
                deleted += 1
    except Exception:
        pass

    if deleted > 0:
        return True, f"已删除手动分页符 {deleted} 个"
    return False, "未找到可安全删除的手动分页符"



def range_has_visible_objects(doc, page_range) -> Tuple[bool, str]:
    """Word 保守判空：页内有对象、域、表格、浮动图形时一律保留。"""
    checks = (
        ("Tables", "表格"),
        ("InlineShapes", "嵌入图片"),
        ("Fields", "域"),
        ("FormFields", "表单域"),
        ("ContentControls", "内容控件"),
        ("ShapeRange", "形状"),
    )
    for attr, label in checks:
        try:
            collection = getattr(page_range, attr)
            if collection is not None and collection.Count > 0:
                return True, label
        except Exception:
            pass

    try:
        for shape in doc.Shapes:
            try:
                anchor = shape.Anchor
                if anchor is not None and page_range.Start <= anchor.Start <= page_range.End:
                    return True, "浮动形状"
            except Exception:
                pass
    except Exception:
        pass

    return False, ""



def process_word_core(word_path: Path) -> Tuple[str, List[int]]:
    word = None
    doc = None
    deleted_pages: List[int] = []
    fixed_manual_pages: List[int] = []
    kept_layout_pages: List[Tuple[int, str]] = []

    heartbeat_state: Dict[str, Any] = {"stage": "启动 Word 引擎", "current": None, "total": None}
    stop_heartbeat, heartbeat_thread = start_heartbeat(word_path.name, heartbeat_state)

    backup_done = False
    backup_path: Optional[Path] = None

    def ensure_backup() -> bool:
        nonlocal backup_done, backup_path
        if backup_done:
            return True
        ok, bp, err = make_backup_before_overwrite(word_path)
        if not ok:
            safe_update_terminal(f"⚠️ [{word_path.name}] 备份原文件失败，已取消修改以避免不可逆误删: {err}")
            return False
        backup_done = True
        backup_path = bp
        safe_update_terminal(f"[*] [{word_path.name}] 已备份原文件: {backup_path.name}")
        return True

    try:
        safe_update_terminal(f"[*] [{word_path.name}] 启动 Word 引擎...")
        word = win32com.client.dynamic.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        heartbeat_state["stage"] = "打开文档"
        safe_update_terminal(f"[*] [{word_path.name}] 正在打开文档...")
        doc = word.Documents.Open(str(word_path))

        heartbeat_state["stage"] = "重新分页"
        safe_update_terminal(f"[*] [{word_path.name}] 正在重新分页，文档较大时可能需要较久...")
        doc.Repaginate()

        try:
            total_pages = max(1, int(doc.ComputeStatistics(2)))
        except Exception:
            total_pages = 1

        heartbeat_state.update({"stage": "动态扫描并清理", "current": total_pages, "total": total_pages})
        safe_update_terminal(f"[*] [{word_path.name}] 共检测到 {total_pages} 页，开始动态扫描空白页...")

        i = total_pages
        while i >= 1:
            try:
                doc.Repaginate()
                current_total = max(1, int(doc.ComputeStatistics(2)))
            except Exception:
                current_total = max(i, 1)

            if i > current_total:
                i = current_total

            heartbeat_state.update({"stage": "动态扫描并清理", "current": i, "total": current_total})
            if i % 10 == 0 or i == current_total:
                safe_update_terminal(f"[*] [{word_path.name}] 当前检查页 {i}/{current_total}...")

            # 极端保护：不能删成空文档
            if current_total == 1 and i == 1:
                try:
                    page_range = get_page_range(doc, 1, 1)
                    has_objects, _ = range_has_visible_objects(doc, page_range)
                    if is_page_strictly_blank(page_range.Text) and not has_objects:
                        safe_update_terminal(f"⚠️ [{word_path.name}] 检测到整篇仅剩 1 个空白页，保留第 1 页避免删成空文件")
                        break
                except Exception:
                    pass

            try:
                page_range = get_page_range(doc, i, current_total)
                if page_range.Start >= page_range.End:
                    i -= 1
                    continue
            except Exception as e:
                safe_update_terminal(f"⚠️ [{word_path.name}] 第 {i} 页取范围失败，跳过: {e}")
                i -= 1
                continue

            page_text = page_range.Text
            has_objects, object_reason = range_has_visible_objects(doc, page_range)
            if has_objects:
                if i % 10 == 0:
                    safe_update_terminal(f"[*] [{word_path.name}] 第 {i} 页含{object_reason}，保留")
                i -= 1
                continue

            if not is_page_strictly_blank(page_text):
                i -= 1
                continue

            layout_reason = detect_layout_blank_reason(doc, page_range, page_text)

            if layout_reason == "manual_page_break":
                if not ensure_backup():
                    return "备份失败，已跳过", []

                ok, detail = fix_manual_layout_blank_page(doc, page_range)
                if ok:
                    fixed_manual_pages.append(i)
                    safe_update_terminal(f"[*] [{word_path.name}] 第 {i} 页为空白页，原因=手动分页符；{detail}")
                else:
                    kept_layout_pages.append((i, "manual_page_break"))
                    safe_update_terminal(f"[*] [{word_path.name}] 第 {i} 页疑似手动分页符空白页，但未找到可安全删除的分页符，已保留")
                i -= 1
                continue

            if layout_reason:
                kept_layout_pages.append((i, layout_reason))
                pretty = {
                    "section_break": "分节符",
                    "column_break": "分栏符",
                    "odd_even_section_break": "奇偶页分节",
                }.get(layout_reason, layout_reason)
                safe_update_terminal(f"[*] [{word_path.name}] 第 {i} 页为空白页，但由{pretty}导致，默认保留")
                i -= 1
                continue

            if not ensure_backup():
                return "备份失败，已跳过", []

            try:
                page_range.Delete()
                deleted_pages.append(i)
                safe_update_terminal(f"[*] [{word_path.name}] 已删除纯空白页: 第 {i} 页")
            except Exception as e:
                safe_update_terminal(f"⚠️ [{word_path.name}] 第 {i} 页删除失败，跳过: {e}")

            i -= 1

        modified_pages = sorted(set(deleted_pages + fixed_manual_pages))

        if modified_pages:
            heartbeat_state.update({"stage": "保存文档", "current": None, "total": None})
            safe_update_terminal(f"[*] [{word_path.name}] 正在保存清理结果...")
            doc.Repaginate()
            doc.Save()

            parts = []
            if deleted_pages:
                parts.append(f"删除纯空白页 {len(deleted_pages)} 页")
            if fixed_manual_pages:
                parts.append(f"修复手动分页空白页 {len(fixed_manual_pages)} 页")
            if kept_layout_pages:
                parts.append(f"保留版式空白页 {len(kept_layout_pages)} 页")

            return "；".join(parts), modified_pages

        if kept_layout_pages:
            safe_update_terminal(
                f"[*] [{word_path.name}] 未做删除。检测到版式空白页 {len(kept_layout_pages)} 页（分节/分栏/奇偶页分节等），默认保留"
            )
            return "仅检测到版式空白页，默认保留", []

        return "无需清理", []

    except pywintypes.com_error as e:
        safe_update_terminal(f"❌ [{word_path.name}] Word 引擎故障: {e}")
        return "处理失败", []

    except Exception as e:
        safe_update_terminal(f"❌ [{word_path.name}] Word 引擎异常: {e}")
        return "处理失败", []

    finally:
        stop_heartbeat.set()
        try:
            heartbeat_thread.join(timeout=1)
        except Exception:
            pass

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


# =========================
# 入口
# =========================

def collect_target_files(target_path: Path) -> List[Path]:
    exts = {".doc", ".docx", ".pdf"}

    if target_path.is_file():
        return [target_path] if target_path.suffix.lower() in exts else []

    if target_path.is_dir():
        return sorted(
            [
                p
                for p in target_path.rglob("*")
                if p.is_file() and p.suffix.lower() in exts and not p.name.startswith("~$")
            ],
            key=lambda x: str(x).lower(),
        )

    return []


@bridge.expose
def run_rm_blank(target_path: str) -> Dict[str, Any]:
    try:
        root = Path(target_path.strip()).resolve()
        files = collect_target_files(root)

        if not files:
            return {"status": "error", "msg": "未找到可处理的 Word/PDF 文件", "data": None}

        total_changed = 0
        processed = 0

        for f in files:
            processed += 1
            safe_update_terminal(f"[*] 开始处理: {f.name}")

            if f.suffix.lower() == ".pdf":
                status_text, changed_pages = process_pdf_core(f)
            else:
                status_text, changed_pages = process_word_core(f)

            total_changed += len(changed_pages)

            if changed_pages:
                safe_update_terminal(f"[√] {f.name}: {status_text} | 处理页码: {', '.join(map(str, changed_pages))}")
            else:
                safe_update_terminal(f"[*] {f.name}: {status_text}")

        return {
            "status": "success",
            "msg": f"处理完成，共 {processed} 个文件，累计处理 {total_changed} 页",
            "data": None,
        }

    except Exception as e:
        safe_update_terminal(f"❌ 空白页清理任务异常: {e}")
        return {"status": "error", "msg": str(e), "data": None}
