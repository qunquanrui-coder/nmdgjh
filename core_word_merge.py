# -*- coding: utf-8 -*-
import eel, re, time, ctypes, pythoncom, win32com.client as win32
from pathlib import Path

@eel.expose
def run_word_merge(folder, out_name):
    # ✨ 核心修复：还原多线程保护模式
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    word = None
    try:
        root = Path(folder).resolve()
        files = [f for f in root.glob('*.doc*') if not f.name.startswith("~$") and out_name not in f.name]
        files.sort(key=lambda p: [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', p.name)])
        
        word = win32.DispatchEx("Word.Application")
        word.Visible = word.ScreenUpdating = False; word.DisplayAlerts = 0
        word.AutomationSecurity = 1 # ✨ 还原：禁用宏警告
        
        merged_doc = word.Documents.Add()
        for idx, src in enumerate(files):
            eel.update_terminal(f"[{idx+1}/{len(files)}] 正在合并: {src.stem}")
            if idx > 0: merged_doc.Range(merged_doc.Content.End-1).InsertBreak(7)
            
            rng = merged_doc.Range(merged_doc.Content.End-1); rng.Text = f"{src.stem}\n"
            try: rng.Style = "标题 1" # ✨ 还原样式应用
            except: pass
            
            source_doc = word.Documents.Open(str(src.resolve()), ReadOnly=True, Visible=False)
            for _ in range(3): # ✨ 还原剪贴板重试机制
                try:
                    ctypes.windll.user32.OpenClipboard(None); ctypes.windll.user32.EmptyClipboard(); ctypes.windll.user32.CloseClipboard()
                    source_doc.Content.Copy(); merged_doc.Range(merged_doc.Content.End-1).PasteAndFormat(16); break
                except: time.sleep(0.5)
            source_doc.Close(False)
        
        merged_doc.SaveAs2(str(root/out_name), 16); merged_doc.Close(False)
        return {"status": "success"}
    finally:
        if word: word.Quit()
        pythoncom.CoUninitialize()
