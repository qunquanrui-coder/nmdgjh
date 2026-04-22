# -*- coding: utf-8 -*-
import bridge, os, re, mmap, tempfile, fitz
from pathlib import Path
try: import pypdf
except: pypdf = None
try: import pikepdf
except: pikepdf = None

def _open_pdf_robust(path):
    # ✨ 严格还原五级强攻打开策略
    errs = []
    try: return fitz.open(path), None
    except Exception as e: errs.append(f"Direct: {e}")
    try:
        f = open(path, "rb")
        mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
        doc = fitz.open(stream=memoryview(mm), filetype="pdf")
        doc._mmap_ref = mm; doc._mmap_file = f
        return doc, None
    except Exception as e: errs.append(f"Mmap: {e}")
    if pypdf:
        try:
            reader = pypdf.PdfReader(path, strict=False); writer = pypdf.PdfWriter()
            for p in reader.pages: writer.add_page(p)
            tmp = tempfile.mktemp(".pdf")
            with open(tmp, "wb") as f: writer.write(f)
            doc = fitz.open(tmp); doc._tmp_path = tmp
            return doc, None
        except Exception as e: errs.append(f"PyPDF_Rewrite: {e}")
    if pikepdf:
        try:
            tmp = tempfile.mktemp(".pdf")
            with pikepdf.open(path, allow_overwriting_input=True) as pdf: pdf.save(tmp)
            doc = fitz.open(tmp); doc._tmp_path = tmp
            return doc, None
        except Exception as e: errs.append(f"PikePDF_Rewrite: {e}")
    return None, " | ".join(errs)

@bridge.expose
def run_split(path, mode, fixed, parts, start, end):
    doc, err = _open_pdf_robust(path)
    if not doc: return {"status": "error", "msg": err}
    try:
        base_n = re.sub(r'[\\/:*?"<>|]', '_', Path(path).stem)
        out_dir = Path(path).parent / f"{base_n}_{mode}"
        out_dir.mkdir(parents=True, exist_ok=True)
        total = len(doc)
        
        # 还原：平均/定长/提取 算法逻辑
        ranges = []
        if mode == "split_avg":
            p = min(max(1, int(parts)), total); b, r = total // p, total % p
            cur = 1
            for i in range(p):
                sz = b + (1 if i < r else 0); ranges.append((cur, cur+sz-1)); cur += sz
        elif mode == "split_fixed":
            step = max(1, int(fixed))
            for i in range(1, total + 1, step): ranges.append((i, min(i+step-1, total)))
        elif mode == "extract":
            s, e = max(1, int(start)), min(int(end), total)
            if s > e: return {"status": "error", "msg": "提取范围无效：起始页不能大于结束页"}
            ranges = [(s, e)]
        else:
            return {"status": "error", "msg": f"未知拆分模式: {mode}"}
        
        for idx, (s, e) in enumerate(ranges, 1):
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=s-1, to_page=e-1)
            fname = f"{base_n}_Part{idx}_P{s}-{e}.pdf"
            new_doc.save(str(out_dir / fname)); new_doc.close()
            bridge.update_terminal(f"[*] 导出: {fname}")
        return {"status": "success", "msg": f"保存至: {out_dir.name}"}
    finally:
        doc.close()
        if hasattr(doc, "_tmp_path") and os.path.exists(doc._tmp_path): os.remove(doc._tmp_path)
