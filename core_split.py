# -*- coding: utf-8 -*-
import bridge, os, re, mmap, tempfile, fitz
from pathlib import Path
from safe_output import new_output_directory
try: import pypdf
except Exception: pypdf = None
try: import pikepdf
except Exception: pikepdf = None

def _open_pdf_robust(path):
    # ✨ 严格还原五级强攻打开策略
    errs = []
    try: return fitz.open(path), None
    except Exception as e: errs.append(f"Direct: {e}")
    f = None
    mm = None
    mv = None
    try:
        f = open(path, "rb")
        mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
        mv = memoryview(mm)
        doc = fitz.open(stream=mv, filetype="pdf")
        doc._mmap_mv = mv; doc._mmap_ref = mm; doc._mmap_file = f
        return doc, None
    except Exception as e:
        errs.append(f"Mmap: {e}")
        if mv is not None:
            try: mv.release()
            except Exception: pass
        if mm is not None:
            try: mm.close()
            except Exception: pass
        if f is not None:
            try: f.close()
            except Exception: pass
    if pypdf:
        tmp = None
        try:
            reader = pypdf.PdfReader(path, strict=False); writer = pypdf.PdfWriter()
            for p in reader.pages: writer.add_page(p)
            fd, tmp = tempfile.mkstemp(prefix=".split_pypdf_", suffix=".pdf")
            os.close(fd)
            with open(tmp, "wb") as tf: writer.write(tf)
            doc = fitz.open(tmp); doc._tmp_path = tmp
            return doc, None
        except Exception as e:
            errs.append(f"PyPDF_Rewrite: {e}")
            if tmp and os.path.exists(tmp):
                try: os.remove(tmp)
                except Exception: pass
    if pikepdf:
        tmp = None
        try:
            fd, tmp = tempfile.mkstemp(prefix=".split_pikepdf_", suffix=".pdf")
            os.close(fd)
            with pikepdf.open(path, allow_overwriting_input=True) as pdf: pdf.save(tmp)
            doc = fitz.open(tmp); doc._tmp_path = tmp
            return doc, None
        except Exception as e:
            errs.append(f"PikePDF_Rewrite: {e}")
            if tmp and os.path.exists(tmp):
                try: os.remove(tmp)
                except Exception: pass
    return None, " | ".join(errs)

@bridge.expose
def run_split(path, mode, fixed, parts, start, end):
    if not path or not str(path).strip():
        return {"status": "error", "msg": "请选择要拆分的 PDF 文件。"}
    doc, err = _open_pdf_robust(path)
    if doc is None: return {"status": "error", "msg": err}
    try:
        base_n = re.sub(r'[\\/:*?"<>|]', '_', Path(path).stem)
        total = len(doc)
        if total == 0:
            return {"status": "error", "msg": "PDF 没有可拆分的页面。"}
        
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
        
        out_dir = new_output_directory(Path(path).parent / f"{base_n}_{mode}")
        for idx, (s, e) in enumerate(ranges, 1):
            fname = f"{base_n}_Part{idx}_P{s}-{e}.pdf"
            with fitz.open() as new_doc:
                new_doc.insert_pdf(doc, from_page=s-1, to_page=e-1)
                new_doc.save(str(out_dir / fname))
            bridge.update_terminal(f"[*] 导出: {fname}")
        return {"status": "success", "msg": f"保存至: {out_dir.name}"}
    finally:
        doc.close()
        if hasattr(doc, "_mmap_mv") and doc._mmap_mv is not None:
            try: doc._mmap_mv.release()
            except Exception: pass
        if hasattr(doc, "_mmap_ref") and doc._mmap_ref is not None:
            try: doc._mmap_ref.close()
            except Exception: pass
        if hasattr(doc, "_mmap_file") and doc._mmap_file is not None:
            try: doc._mmap_file.close()
            except Exception: pass
        if hasattr(doc, "_tmp_path") and os.path.exists(doc._tmp_path):
            try: os.remove(doc._tmp_path)
            except Exception: pass
