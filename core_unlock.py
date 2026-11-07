# -*- coding: utf-8 -*-
import bridge
import fitz
import tempfile
import os
import multiprocessing as mp
from pathlib import Path
from queue import Empty
from safe_output import publish_output

def _is_password_error(msg: str) -> bool:
    if not msg: return False
    m = msg.lower()
    return ("密码" in msg) or ("password" in m) or ("needs_pass" in m) or ("authenticate" in m)

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

        with tempfile.NamedTemporaryFile(prefix=".unlock-", suffix=".pdf", dir=p.parent, delete=False) as tf:
            tmp_path = Path(tf.name)

        save_kwargs = {"deflate": True, "garbage": (2 if mode == 1 else 4)}
        doc.save(tmp_path, encryption=fitz.PDF_ENCRYPT_NONE, **save_kwargs)
        doc.close(); doc = None

        if keep_suffix:
            out_path = publish_output(tmp_path, p.with_name(f"{p.stem}_unlocked{p.suffix}"))
        else:
            # 同目录生成完整结果，再原子替换用户明确选择覆盖的原文件。
            os.replace(tmp_path, p)
            out_path = p
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
    succeeded = 0
    first_failure = None
    
    for i, f in enumerate(targets, 1):
        q.put(("log", f"正在处理 ({i}/{len(targets)}): {Path(f).name}"))
        r = _process_one_pdf(f, user_pwd, allow_empty, keep_suffix, mode)

        if (not r["ok"]) and (mode == 1) and auto_retry:
            if not _is_password_error(r.get("err", "")):
                q.put(("log", f"[!] 快速失败，正转为安全模式重试: {Path(f).name}"))
                r = _process_one_pdf(f, user_pwd, allow_empty, keep_suffix, 0)

        if r["ok"]:
            succeeded += 1
            q.put(("log", f"[√] 成功: {Path(r['out']).name}"))
        else:
            failure = f"{Path(f).name} - {r['err']}"
            if first_failure is None:
                first_failure = failure
            q.put(("log", f"[x] 失败/跳过: {failure}"))

    q.put(("done", (succeeded, len(targets) - succeeded, first_failure)))

@bridge.expose
def run_unlock(path, pwd, allow_empty, keep_suffix, mode, auto_retry):
    try:
        if path is None or not str(path).strip():
            return {"status": "error", "msg": "请选择需要解密的 PDF 文件或文件夹"}
        root = Path(path).resolve()
        if root.is_file():
            targets = [str(root)] if root.suffix.lower() == ".pdf" else []
        else:
            targets = [
                str(f) for f in root.rglob("*.pdf")
                if f.is_file() and "_unlocked" not in f.name and not f.name.startswith(("~$", ".unlock-"))
            ]
        if not targets: return {"status": "error", "msg": "未找到需要处理的PDF"}

        # 隔离原生 PDF 库，避免子进程异常导致桌面程序崩溃。
        ctx = mp.get_context("spawn")
        q = ctx.Queue()
        try:
            args = (targets, pwd, allow_empty, keep_suffix, int(mode), auto_retry)
            proc = ctx.Process(target=_worker_process_entry, args=(args, q), daemon=True)
            proc.start()

            while True:
                try:
                    kind, payload = q.get(timeout=0.5)
                except Empty:
                    if proc.is_alive():
                        continue
                    # The worker can finish between the timeout and is_alive().
                    try:
                        kind, payload = q.get_nowait()
                    except Empty:
                        return {"status": "error", "msg": f"解密子进程未返回结果 (exitcode={proc.exitcode})"}
                if kind == "log":
                    bridge.update_terminal(payload)
                elif kind == "done":
                    succeeded, failed, first_failure = payload
                    break

            proc.join(timeout=1)
        finally:
            q.close()

        status = "error" if not succeeded else "partial" if failed else "success"
        msg = f"解密完成，成功 {succeeded} 个，失败 {failed} 个"
        if first_failure:
            msg += f"。首个失败：{first_failure}"
        return {"status": status, "msg": msg, "data": {"succeeded": succeeded, "failed": failed}}
    except Exception as e:
        return {"status": "error", "msg": str(e)}
