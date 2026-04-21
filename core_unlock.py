# -*- coding: utf-8 -*-
import eel
import fitz
import shutil
import tempfile
import os
import time
import multiprocessing as mp
from pathlib import Path
from datetime import datetime

def _is_password_error(msg: str) -> bool:
    if not msg: return False
    m = msg.lower()
    return ("密码" in msg) or ("password" in m) or ("needs_pass" in m) or ("authenticate" in m)

def _safe_out_path(src: Path, keep_suffix: bool) -> Path:
    out_path = src.with_name(src.stem + "_unlocked" + src.suffix) if keep_suffix else src
    if keep_suffix and out_path.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = src.with_name(src.stem + f"_unlocked_{ts}" + src.suffix)
    return out_path

def _process_one_pdf(pdf_path: str, user_pwd: str, allow_empty: bool, keep_suffix: bool, mode: int) -> dict:
    p = Path(pdf_path)
    doc, tmp_path = None, None
    try:
        doc = fitz.open(p)
        if getattr(doc, "needs_pass", False):
            if not user_pwd or not doc.authenticate(user_pwd):
                return {"ok": False, "err": "打开密码错误或未提供"}
        elif getattr(doc, "is_encrypted", False):
            if allow_empty:
                if not doc.authenticate(""): return {"ok": False, "err": "空口令授权失败"}
            else:
                return {"ok": False, "err": "文件受限但未允许尝试移除权限"}

        out_path = _safe_out_path(p, keep_suffix)
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
            tmp_path = Path(tf.name)

        save_kwargs = {"deflate": True, "garbage": (2 if mode == 1 else 4)}
        doc.save(tmp_path, encryption=fitz.PDF_ENCRYPT_NONE, **save_kwargs)
        doc.close(); doc = None

        shutil.move(str(tmp_path), str(out_path))
        return {"ok": True, "out": str(out_path)}
    except Exception as e:
        return {"ok": False, "err": str(e)}
    finally:
        if doc:
            try: doc.close()
            except: pass
        if tmp_path and tmp_path.exists():
            try: tmp_path.unlink()
            except: pass

def _worker_process_entry(args, q):
    targets, user_pwd, allow_empty, keep_suffix, mode, auto_retry = args
    q.put(("log", f"启动子进程清洗，共 {len(targets)} 个文件..."))
    
    for i, f in enumerate(targets, 1):
        q.put(("log", f"正在处理 ({i}/{len(targets)}): {Path(f).name}"))
        r = _process_one_pdf(f, user_pwd, allow_empty, keep_suffix, mode)

        if (not r["ok"]) and (mode == 1) and auto_retry:
            if not _is_password_error(r.get("err", "")):
                q.put(("log", f"[!] 快速失败，正转为安全模式重试: {Path(f).name}"))
                r = _process_one_pdf(f, user_pwd, allow_empty, keep_suffix, 0)

        if r["ok"]: q.put(("log", f"[√] 成功: {Path(r['out']).name}"))
        else: q.put(("log", f"[x] 失败/跳过: {Path(f).name} - {r['err']}"))
            
    q.put(("done", "全部完成"))

@eel.expose
def run_unlock(path, pwd, allow_empty, keep_suffix, mode, auto_retry):
    try:
        root = Path(path).resolve()
        targets = [str(root)] if root.is_file() else [str(f) for f in root.rglob("*.pdf") if "_unlocked" not in f.name and not f.name.startswith("~$")]
        if not targets: return {"status": "error", "msg": "未找到需要处理的PDF"}

        # 核心修复：多进程隔离并通过 time.sleep 轮询，防原生线程奔溃
        ctx = mp.get_context("spawn")
        q = ctx.Queue()
        args = (targets, pwd, allow_empty, keep_suffix, int(mode), auto_retry)
        
        proc = ctx.Process(target=_worker_process_entry, args=(args, q), daemon=True)
        proc.start()

        completed = False
        while True:
            while not q.empty():
                try:
                    item = q.get_nowait()
                    if item[0] == "log":
                        eel.update_terminal(item[1])
                    elif item[0] == "done":
                        completed = True
                except: pass
                
            if completed: break
            
            if not proc.is_alive():
                if not completed:
                    return {"status": "error", "msg": f"子进程异常崩溃 (exitcode={proc.exitcode})"}
                break
                
            time.sleep(0.5) # ✨ 核心修复：原生线程必须使用 time.sleep

        proc.join(timeout=1)
        return {"status": "success", "msg": "解密处理完成"}
        
    except Exception as e:
        return {"status": "error", "msg": str(e)}
