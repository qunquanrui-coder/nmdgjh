# -*- coding: utf-8 -*-
import os
import re
import tempfile
import logging
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any # [新增] 类型注解支持
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
_gen_path = Path(tempfile.gettempdir()) / 'gen_py'
_gen_path.mkdir(parents=True, exist_ok=True)
win32com.__gen_path__ = str(_gen_path)

def is_page_strictly_blank(text: str) -> bool:
    """[Word专用算法 - 100% 原样保留]"""
    if not text:
        return True
    cleaned = re.sub(r'[ \t\n\x0b\x0c\r\x0e\x0f\xa0\u3000\u200b]+', '', text)
    return len(cleaned) == 0

def safe_update_terminal(msg: str):
    """封装终端输出"""
    logging.info(msg)
    try:
        eel.update_terminal(msg)
    except Exception:
        pass

# --- [PDF 引擎：融合全量加固建议的最终版] ---
def process_pdf_core(pdf_path: Path) -> Tuple[str, List[int]]:
    if fitz is None:
        safe_update_terminal(f"❌ [{pdf_path.name}] 缺少pymupdf库")
        return "依赖缺失", []

    doc: Optional[fitz.Document] = None
    temp_path: Optional[str] = None
    
    try:
        # 1. 安全打开
        try:
            doc = fitz.open(str(pdf_path))
        except Exception:
            return "文件受损/加密", []

        total_pages: int = len(doc)
        keep_indices: List[int] = []
        deleted_pages: List[int] = []

        for i in range(total_pages):
            # 进度反馈
            if (i + 1) % 50 == 0 or i == total_pages - 1:
                safe_update_terminal(f"💓 PDF引擎扫描: [{pdf_path.name}] {i+1}/{total_pages}...")

            page: fitz.Page = doc[i]
            
            # --- 判定流程 (文字 -> 图像 -> 矢量 -> 注释) ---
            # A. 文本检测
            if not is_page_strictly_blank(page.get_text()):
                keep_indices.append(i); continue
            
            # B. 图像检测
            if len(page.get_images()) > 0:
                keep_indices.append(i); continue
            
            # C. 矢量图形判定 [加固：防御性捕获旧版 API 错误]
            try:
                drawings = page.get_drawings()
                if isinstance(drawings, list) and len(drawings) > 0:
                    keep_indices.append(i); continue
            except (AttributeError, Exception):
                pass 
            
            # D. 注释与表单判定 [加固：空值防护判定]
            try:
                raw_annots = page.annots()
                # 过滤链接，识别实质性注释
                annots = [a for a in raw_annots if a.type[0] != 1] if raw_annots else []
                
                raw_widgets = page.widgets()
                widgets = list(raw_widgets) if raw_widgets else []
                
                if annots or widgets:
                    keep_indices.append(i); continue
            except Exception:
                pass 

            deleted_pages.append(i + 1)

        if not deleted_pages:
            doc.close()
            return "无需清理", []

        # 极端保护
        if not keep_indices:
            keep_indices = [0]
            if 1 in deleted_pages: deleted_pages.remove(1)

        # 2. 内存操作与句柄释放隔离
        doc.select(keep_indices)
        temp_fd, temp_path = tempfile.mkstemp(suffix=".pdf")
        os.close(temp_fd)
        
        try:
            doc.save(temp_path, garbage=4, deflate=True)
        finally:
            # [加固] 放入独立 try...finally 强制释放
            if doc is not None:
                try: doc.close()
                except: pass
            doc = None 

        # [加固] 使用 os.replace 实现原子替换
        os.replace(temp_path, str(pdf_path))
        return f"已清理 ({len(deleted_pages)}页)", deleted_pages

    except Exception as e:
        if doc is not None:
            try: doc.close()
            except: pass
        if temp_path and os.path.exists(temp_path):
            try: os.remove(temp_path)
            except: pass
        safe_update_terminal(f"❌ [{pdf_path.name}] PDF引擎故障: {str(e)}")
        return "处理失败", []

@eel.expose
def run_rm_blank(target_path: str) -> Dict[str, str]:
    if not target_path:
        return {"status": "error", "msg": "路径不能为空"}
        
    target: Path = Path(target_path)
    files: List[Path] = []
    exts: Tuple[str, ...] = ('.doc', '.docx', '.pdf')
    
    if target.is_file():
        if target.suffix.lower() in exts: files.append(target)
    else:
        files = [f for f in target.rglob('*') if f.is_file() and f.suffix.lower() in exts and not f.name.startswith('~$')]
        
    if not files:
        return {"status": "error", "msg": "未找到有效文档"}
        
    safe_update_terminal(f"[*] 发现 {len(files)} 个目标，启动混合引擎...")
    summary_data: List[Dict[str, Any]] = []
    word_app: Any = None
    
    try:
        for doc_path in files:
            fname, ext = doc_path.name, doc_path.suffix.lower()
            
            # --- 分流 A: PDF ---
            if ext == '.pdf':
                status, deleted = process_pdf_core(doc_path)
                summary_data.append({"file": fname, "pages": sorted(deleted) if deleted else "无", "status": status})
                continue

            # --- 分流 B: Word [原有算法，逻辑完全保留] ---
            if word_app is None:
                word_app = win32com.client.dynamic.Dispatch("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = False 
            
            safe_update_terminal(f"⏳ 扫描 Word: [{fname}]")
            deleted_word_pages: List[int] = []
            file_status: str = "正常"
            try:
                doc = word_app.Documents.Open(str(doc_path))
                doc.Repaginate()
                total = doc.ComputeStatistics(2)
                needs_repaginate = False 
                
                # --- 这里开始是你最初完全没问题的算法 ---
                for i in range(total, 0, -1):
                    if i % 10 == 0 or i == total:
                        safe_update_terminal(f"💓 引擎运转中: [{fname}] 当前第 {i}/{total} 页...")
                    if needs_repaginate:
                        doc.Repaginate(); needs_repaginate = False 
                    curr_total = doc.ComputeStatistics(2)
                    if i > curr_total: continue
                    
                    start = doc.GoTo(1, 1, i).Start
                    end = doc.GoTo(1, 1, i+1).Start if i < curr_total else doc.Content.End
                    if start >= end: continue
                    
                    rng = doc.Range(start, end)
                    if is_page_strictly_blank(rng.Text):
                        rng.Delete()
                        needs_repaginate = True 
                        if start > 0:
                            p_rng = doc.Range(start - 1, start)
                            if p_rng.Text in ('\x0c', '\x0e', '\x0f'): p_rng.Delete()
                        deleted_word_pages.append(i)
                # --- 算法结束 ---
                
                if deleted_word_pages:
                    deleted_word_pages.sort(); doc.Save()
                    file_status = f"已清理 ({len(deleted_word_pages)}页)"
                    safe_update_terminal(f"✅ [{fname}] Word清理完毕")
                else:
                    file_status = "无需清理"
            except Exception as e:
                file_status = "处理异常"
                safe_update_terminal(f"❌ [{fname}] Word错误: {str(e)}")
            finally:
                summary_data.append({"file": fname, "pages": sorted(deleted_word_pages) if deleted_word_pages else "无", "status": file_status})
                if doc:
                    try: doc.Close(SaveChanges=False)
                    except: pass

        # 最终汇总
        if summary_data:
            report = "\n" + "═"*30 + "\n📊 批量处理最终统计汇总\n" + "─"*30 + "\n"
            for item in summary_data:
                report += f"📄 {item['file']}\n   结果: {item['status']} | 页码: {item['pages']}\n"
            safe_update_terminal(report + "═"*30)

        return {"status": "success", "msg": "全部任务完成"}
        
    except Exception as e:
        return {"status": "error", "msg": str(e)}
    finally:
        if word_app:
            try: word_app.Quit()
            except: pass
