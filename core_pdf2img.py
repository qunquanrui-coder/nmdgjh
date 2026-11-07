# -*- coding: utf-8 -*-
import bridge, fitz
from pathlib import Path
from PIL import Image
from safe_output import new_output_directory

@bridge.expose
def run_pdf2img(path, longedge, quality):
    if not path or not str(path).strip():
        return {"status": "error", "msg": "请选择要转为图片的 PDF 文件。"}
    try:
        pdf_p = Path(path).resolve()
        long_edge = int(longedge)
        image_quality = int(quality)
        if long_edge < 1:
            return {"status": "error", "msg": "图片长边必须大于 0。"}

        with fitz.open(pdf_p) as doc:
            if len(doc) == 0:
                return {"status": "error", "msg": "PDF 没有可转换的页面。"}
            out_dir = new_output_directory(pdf_p.parent / f"{pdf_p.stem}_images")
            for index, page in enumerate(doc, 1):
                # PyMuPDF already applies the crop box and rotation when rendering.
                # page.rect measures the visible page in the same coordinates.
                width_pt, height_pt = page.rect.width, page.rect.height
                scale = long_edge / max(width_pt, height_pt)
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                dpi_x = round(pix.width * 72 / width_pt)
                dpi_y = round(pix.height * 72 / height_pt)
                with Image.frombytes("RGB", (pix.width, pix.height), pix.samples) as image:
                    image.save(out_dir / f"{index:04d}.jpg", "JPEG", quality=image_quality, dpi=(dpi_x, dpi_y))
                if index % 10 == 0:
                    bridge.update_terminal(f"[*] 渲染第 {index} 页...")
        return {"status": "success", "msg": f"保存至: {out_dir.name}"}
    except Exception as e: return {"status": "error", "msg": str(e)}
