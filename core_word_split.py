# -*- coding: utf-8 -*-
"""
Word 文档大纲拆分引擎 (终极防卡死极速版)
彻底移除了 O(N) 级别的段落遍历，全系采用 C++ 层面的原生 Find 搜索。
"""
import bridge
import re
import time
import ctypes
import threading
import gc
import pythoncom
import win32com.client as win32
from pathlib import Path
from typing import Dict, Any, List, Optional

try:
    from docx import Document as DocxDocument
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


class WordSplitterEngine:
    def __init__(self) -> None:
        self._stop_event = threading.Event()

    def _apply_speed_hacks(self, word: Any) -> None:
        try:
            word.ScreenUpdating = False
            word.Options.Pagination = False
            word.Options.CheckSpellingAsYouType = False
            word.Options.CheckGrammarAsYouType = False
            word.DisplayAlerts = 0
            word.AutomationSecurity = 1
        except pythoncom.com_error:
            pass

    def _clear_clipboard(self) -> None:
        try:
            ctypes.windll.user32.OpenClipboard(None)
            ctypes.windll.user32.EmptyClipboard()
            ctypes.windll.user32.CloseClipboard()
        except OSError:
            pass

    def scan_outline(self, doc_path: Path) -> Dict[int, int]:
        level_counts: Dict[int, int] = {}
        
        # 1. 优先使用 python-docx 极速扫描 (强化了对中文样式的泛化识别)
        if DOCX_AVAILABLE and doc_path.suffix.lower() == '.docx':
            try:
                doc = DocxDocument(str(doc_path))
                for para in doc.paragraphs:
                    lvl: Optional[int] = None
                    
                    # 尝试强制从样式名提取 (兼容不规范的标书排版)
                    if para.style and para.style.name:
                        s_name = para.style.name.replace(" ", "")
                        if s_name.startswith("Heading"):
                            try: lvl = int(s_name[7:])
                            except ValueError: pass
                        elif s_name.startswith("标题"):
                            try: lvl = int(s_name[2:])
                            except ValueError: pass
                    
                    # 尝试从 XML 底层提取
                    if not lvl:
                        pPr = getattr(para._element, 'pPr', None)
                        if pPr is not None:
                            outline = pPr.find(qn('w:outlineLvl'))
                            if outline is not None:
                                try: lvl = int(outline.val) + 1
                                except ValueError: pass
                                
                    if lvl and 1 <= lvl <= 9 and para.text.strip():
                        level_counts[lvl] = level_counts.get(lvl, 0) + 1
            except Exception as e:
                print(f"[Warn] docx 极速引擎解析失败，准备降级: {e}")

        # 2. 降级方案：COM 接口 (严禁段落遍历，改用极速原生 Find 引擎)
        if not level_counts:
            pythoncom.CoInitialize()
            word: Any = None
            try:
                word = win32.DispatchEx("Word.Application")
                word.Visible = False
                self._apply_speed_hacks(word)
                doc = word.Documents.Open(str(doc_path.resolve()), False, True, False)

                for lvl in range(1, 10):
                    rng = doc.Range(0, 0)
                    rng.Find.ClearFormatting()
                    rng.Find.ParagraphFormat.OutlineLevel = lvl
                    rng.Find.Text = ""
                    rng.Find.Forward = True
                    rng.Find.Wrap = 0 # wdFindStop
                    
                    iters = 0
                    # 设置绝对熔断器，防止特殊字符导致死循环
                    while rng.Find.Execute() and iters < 3000:
                        iters += 1
                        try:
                            clean_text = re.sub(r'[\x00-\x1f\s]', '', rng.Text)
                            if clean_text:
                                level_counts[lvl] = level_counts.get(lvl, 0) + 1
                        except pythoncom.com_error:
                            pass
                        # 强行推进搜索指针，避开原地打转
                        rng.Collapse(0) # wdCollapseEnd

                doc.Close(0)
            except Exception as e:
                print(f"[ERROR] COM 扫描失败: {e}")
            finally:
                gc.collect()
                if word:
                    try: word.Quit(0)
                    except: pass
                pythoncom.CoUninitialize()
                
        return level_counts

    def split_document(self, doc_path: Path, out_dir: Path, target_level: int, res_container: Dict[str, Any]) -> None:
        pythoncom.CoInitialize()
        word: Any = None
        source_doc: Any = None
        try:
            res_container['msg'] = "[*] 拉起底层 Word 进程并加载文件..."
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            self._apply_speed_hacks(word)

            source_doc = word.Documents.Open(str(doc_path.resolve()), False, True)
            doc_end_pos = source_doc.Range().End

            all_headings_dict: Dict[int, Dict[str, Any]] = {}
            res_container['msg'] = "[*] 启动极速雷达精确定位大纲节点..."

            # 仅搜索目标级别及以上的标题，直接截断底层遍历
            for lvl in range(1, target_level + 1):
                rng = source_doc.Range(0, 0)
                rng.Find.ClearFormatting()
                rng.Find.ParagraphFormat.OutlineLevel = lvl
                rng.Find.Text = ""
                rng.Find.Forward = True
                rng.Find.Wrap = 0 
                
                iters = 0
                while rng.Find.Execute() and iters < 5000:
                    iters += 1
                    try:
                        clean_text = re.sub(r'[\x00-\x1f\s]', '', rng.Text)
                        if clean_text and rng.Start not in all_headings_dict:
                            # 避开可能导致异常的表格区域
                            try: in_tbl = rng.Information(12)
                            except: in_tbl = False
                            
                            if not in_tbl:
                                para = rng.Paragraphs(1)
                                all_headings_dict[rng.Start] = {'start': rng.Start, 'level': lvl, 'para': para}
                    except: pass
                    rng.Collapse(0)

            all_headings = [all_headings_dict[k] for k in sorted(all_headings_dict.keys())]

            target_nodes: List[Dict[str, Any]] = []
            for idx, h in enumerate(all_headings):
                if h['level'] == target_level:
                    end_pos = doc_end_pos
                    for next_h in all_headings[idx + 1:]:
                        if next_h['level'] <= target_level:
                            end_pos = next_h['start']
                            break
                    target_nodes.append({'start': h['start'], 'end': end_pos, 'para': h['para']})

            total = len(target_nodes)
            if total == 0:
                res_container['status'] = 'error'
                res_container['error'] = f"未在文档中找到级别为 {target_level} 的有效大纲标题。"
                return

            out_path = out_dir.resolve()
            out_path.mkdir(parents=True, exist_ok=True)

            for idx, node in enumerate(target_nodes):
                if self._stop_event.is_set():
                    break
                start_pos = node['start']
                end_pos = node['end']

                safe_name = re.sub(r'[\x00-\x1f\\/:*?"<>|\r\n\t]', "_", node['para'].Range.Text.strip())[:80]
                if not safe_name: safe_name = f"未命名拆分段_{idx + 1}"

                res_container['msg'] = f"  └─ 正在拆分 ({idx + 1}/{total}): {safe_name[:20]}..."

                max_retries = 10
                for attempt in range(max_retries):
                    new_doc: Any = None
                    try:
                        source_doc.Range(start_pos, end_pos).Copy()
                        time.sleep(0.5)

                        new_doc = word.Documents.Add()
                        new_doc.Range().PasteAndFormat(16)

                        # 清理头部自带的回车
                        try:
                            if new_doc.Paragraphs.Count > 0:
                                new_doc.Paragraphs(1).Range.Delete()
                        except pythoncom.com_error:
                            pass

                        # 清理尾部顽固的排版符
                        try:
                            last_d_end = -1
                            while True:
                                d_end = new_doc.Range().End
                                if d_end <= 1 or d_end == last_d_end: break
                                last_d_end = d_end

                                last_char_rng = new_doc.Range(d_end - 2, d_end - 1)
                                char = last_char_rng.Text

                                if not char: break
                                if char in ('\x0c', '\x0e', '\x0b', '\n', ' ', '\t', '\u3000', '\xa0'):
                                    last_char_rng.Delete()
                                elif char == '\r':
                                    p_count = new_doc.Paragraphs.Count
                                    if p_count <= 1: break
                                    prev_p = new_doc.Paragraphs(p_count - 1)
                                    if prev_p.Range.Tables.Count > 0: break
                                    last_char_rng.Delete()
                                else: break
                        except pythoncom.com_error:
                            pass

                        final_file = out_path / f"{safe_name}.docx"
                        if final_file.exists():
                            final_file = out_path / f"{safe_name}_{int(time.time())}.docx"

                        new_doc.SaveAs2(str(final_file), 16)
                        new_doc.Close(0)
                        break

                    except Exception as e:
                        if new_doc:
                            try: new_doc.Close(0)
                            except pythoncom.com_error: pass

                        err_str = str(e)
                        if "2147418111" in err_str or "拒绝" in err_str or "rejected" in err_str:
                            if attempt == max_retries - 1:
                                raise RuntimeError(f"Word 底层引擎持续拒绝响应: {err_str}")
                            time.sleep(3 + attempt * 2)
                            pythoncom.PumpWaitingMessages()
                            self._clear_clipboard()
                            gc.collect()
                        else:
                            raise e
                    pythoncom.PumpWaitingMessages()

                if idx % 5 == 0:
                    self._clear_clipboard()
                    gc.collect()

            res_container['status'] = 'ok'

        except Exception as e:
            res_container['status'] = 'error'
            res_container['error'] = str(e)
        finally:
            for var in ['all_headings_dict', 'target_nodes', 'all_headings']:
                if var in locals(): locals()[var].clear()
            gc.collect()

            if source_doc:
                try: source_doc.Close(0)
                except pythoncom.com_error: pass
            if word:
                try: word.Quit(0)
                except pythoncom.com_error: pass
            source_doc = None
            word = None
            gc.collect()
            pythoncom.CoUninitialize()


_engine = WordSplitterEngine()


@bridge.expose
def handle_file_selection(path_str: str) -> Dict[str, str]:
    p = Path(path_str)
    return {"out_dir": str(p.parent / f"{p.stem}_拆分结果")}


@bridge.expose
def get_word_outline(doc_path: str) -> Dict[str, Any]:
    try:
        counts = _engine.scan_outline(Path(doc_path))
        if not counts:
            return {"recommended": None, "status_text": "❌ 未发现有效大纲(检查格式)", "options": ["无可用级别"]}

        valid = [(l, c) for l, c in counts.items() if 1 < c <= 150]
        rec_lvl = max(valid, key=lambda x: x[1])[0] if valid else min(counts.keys())
        opts = [f"级别 {l} ({c}处)" for l, c in sorted(counts.items())]

        return {
            "recommended": rec_lvl,
            "status_text": f"✅ 建议级别 {rec_lvl}",
            "options": opts,
            "status": "success"
        }
    except Exception as e:
        return {"status": "error", "status_text": f"⚠️ 分析出错", "options": []}


@bridge.expose
def run_word_split(doc_path: str, out_dir: str, target_level_str: str, recommended_lvl: int) -> Dict[str, Any]:
    if "推荐" in target_level_str or not any(char.isdigit() for char in target_level_str):
        target = int(recommended_lvl)
    else:
        match = re.search(r'\d+', target_level_str)
        target = int(match.group()) if match else 1

    try:
        res_container: Dict[str, Any] = {'status': 'pending', 'msg': '', 'error': ''}
        final_out_dir = str(Path(out_dir) / f"按级别{target}拆分")

        t = threading.Thread(target=_engine.split_document, args=(Path(doc_path), Path(final_out_dir), target, res_container))
        t.start()

        last_msg = ""
        while t.is_alive():
            time.sleep(0.5)
            curr_msg = res_container.get('msg', '')
            if curr_msg and curr_msg != last_msg:
                bridge.update_terminal(curr_msg)
                last_msg = curr_msg

        if res_container.get('status') == 'error':
            raise RuntimeError(res_container.get('error'))

        return {"status": "success"}
    except Exception as e:
        return {"status": "error", "msg": str(e)}
