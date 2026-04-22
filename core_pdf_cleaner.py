# -*- coding: utf-8 -*-
import eel
import fitz  # PyMuPDF
import cv2
import numpy as np
from pathlib import Path
import time

@eel.expose
def run_pdf_cleaner(path_str: str):
    """PDF 扫描件智能去黑边核心逻辑 (接入 Eel 终端版)"""
    try:
        target_path = Path(path_str.strip()).resolve()
        pdf_files = []

        if target_path.is_file() and target_path.suffix.lower() == '.pdf':
            pdf_files = [target_path]
            output_dir = target_path.parent
        elif target_path.is_dir():
            pdf_files = list(target_path.glob("*.pdf"))
            output_dir = target_path / "已去边"
            if pdf_files and not output_dir.exists():
                output_dir.mkdir()
        else:
            return {"status": "error", "msg": "无效的路径，请选择 PDF 文件或包含 PDF 的文件夹"}

        if not pdf_files:
            return {"status": "error", "msg": "未找到任何 PDF 文件！"}

        eel.update_terminal(f"[*] 共找到 {len(pdf_files)} 个 PDF 文件，准备处理...")
        
        success_count = 0
        for idx, pdf_path in enumerate(pdf_files):
            eel.update_terminal(f"[*] 正在处理 ({idx + 1}/{len(pdf_files)}): {pdf_path.name}")
            
            if target_path.is_file():
                out_name = f"{pdf_path.stem}_去边{pdf_path.suffix}"
                out_path = output_dir / out_name
            else:
                out_path = output_dir / pdf_path.name

            doc = None
            try:
                doc = fitz.open(str(pdf_path))
                total_pages = len(doc)

                for page_num in range(total_pages):
                    if page_num % 10 == 0:
                        eel.update_terminal(f"  └─ 正在扫描页面 {page_num + 1}/{total_pages}...")
                        time.sleep(0.01)

                    page = doc[page_num]
                    page.set_cropbox(page.mediabox)

                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)

                    if pix.n >= 3:
                        gray = cv2.cvtColor(img_data, cv2.COLOR_RGB2GRAY)
                    else:
                        gray = img_data

                    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 51, 15)
                    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

                    h_img, w_img = thresh.shape
                    edge_dist_x = int(w_img * 0.03)
                    edge_dist_y = int(h_img * 0.03)

                    valid_x, valid_y = [], []

                    for cnt in contours:
                        x, y, w, h = cv2.boundingRect(cnt)
                        if w * h < 15: continue
                        if (x <= edge_dist_x or (x + w) >= (w_img - edge_dist_x) or
                            y <= edge_dist_y or (y + h) >= (h_img - edge_dist_y)):
                            continue
                        valid_x.extend([x, x + w])
                        valid_y.extend([y, y + h])

                    if valid_x and valid_y:
                        min_x, max_x = min(valid_x), max(valid_x)
                        min_y, max_y = min(valid_y), max(valid_y)

                        scale = 2.0
                        pdf_min_x = max(0, min_x / scale - 5)
                        pdf_max_x = min(page.rect.width, max_x / scale + 5)
                        pdf_min_y = max(0, min_y / scale - 5)
                        pdf_max_y = min(page.rect.height, max_y / scale + 5)

                        r_top = fitz.Rect(0, 0, page.rect.width, pdf_min_y)
                        r_bottom = fitz.Rect(0, pdf_max_y, page.rect.width, page.rect.height)
                        r_left = fitz.Rect(0, pdf_min_y, pdf_min_x, pdf_max_y)
                        r_right = fitz.Rect(pdf_max_x, pdf_min_y, page.rect.width, pdf_max_y)

                        for r in [r_top, r_bottom, r_left, r_right]:
                            if r.width > 0 and r.height > 0:
                                page.draw_rect(r, color=None, fill=(1, 1, 1))
                    else:
                        page.draw_rect(page.rect, color=None, fill=(1, 1, 1))

                doc.save(str(out_path))
            finally:
                if doc is not None:
                    doc.close()

            success_count += 1
            eel.update_terminal(f"  └─ ✅ 保存成功: {out_path.name}")

        return {"status": "success", "msg": f"🎉 执行完毕！成功去边 {success_count}/{len(pdf_files)} 个文件。"}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"status": "error", "msg": str(e)}
