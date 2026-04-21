# -*- coding: utf-8 -*-
import eel, fitz
from pathlib import Path
from PIL import Image

@eel.expose
def run_pdf2img(path, longedge, quality):
    try:
        pdf_p = Path(path).resolve()
        out_dir = pdf_p.parent / f"{pdf_p.stem}_images"
        out_dir.mkdir(parents=True, exist_ok=True)
        doc = fitz.open(pdf_p)
        for pi, page in enumerate(doc):
            # ✨ 严格还原 Acrobat 对齐逻辑
            clip = getattr(page, "cropbox", None) or page.rect
            w_pt, h_pt = clip.width, clip.height
            le = int(longedge)
            if h_pt >= w_pt: tw, th = int(round(le * w_pt / h_pt)), le
            else: tw, th = le, int(round(le * h_pt / w_pt))
            
            pix = page.get_pixmap(matrix=fitz.Matrix(tw/w_pt, th/h_pt), clip=clip, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            
            # ✨ 还原精确 DPI 写入逻辑
            dpi_x = int(round(pix.width / (w_pt / 72.0)))
            dpi_y = int(round(pix.height / (h_pt / 72.0)))
            img.save(str(out_dir / f"{pi+1:04d}.jpg"), "JPEG", quality=int(quality), dpi=(dpi_x, dpi_y))
            if (pi+1) % 10 == 0: eel.update_terminal(f"[*] 渲染第 {pi+1} 页...")
        return {"status": "success"}
    except Exception as e: return {"status": "error", "msg": str(e)}
