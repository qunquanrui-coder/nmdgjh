# -*- coding: utf-8 -*-
import eel, io, re, img2pdf
from pathlib import Path
from PIL import Image

# ✨ 还原透明 PNG 补白逻辑
def _png_to_white_jpeg_bytes(p):
    with Image.open(p) as im:
        if im.mode in ("RGBA", "LA"):
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im, mask=im.split()[-1]); out = bg
        else: out = im.convert("RGB")
        buf = io.BytesIO(); out.save(buf, format="JPEG", quality=95); buf.seek(0)
        return buf

@eel.expose
def run_img2pdf(path, recursive, include_root):
    try:
        root_folder = Path(path).resolve()
        targets = [d for d in root_folder.rglob("*") if d.is_dir()] if recursive else [root_folder]
        if recursive and include_root: targets.insert(0, root_folder)
        
        for folder in targets:
            imgs = sorted([p for p in folder.iterdir() if p.suffix.lower() in (".jpg",".png",".bmp")], key=lambda p: [int(t) if t.isdigit() else t.lower() for t in re.split(r"([0-9]+)", p.name)])
            if not imgs: continue
            
            eel.update_terminal(f"[*] 正在打包目录: {folder.name}")
            inputs = [_png_to_white_jpeg_bytes(p) if p.suffix.lower() == ".png" else str(p) for p in imgs]
            with open(folder / f"{folder.name}.pdf", "wb") as f:
                img2pdf.convert(inputs, outputstream=f)
        return {"status": "success"}
    except Exception as e: return {"status": "error", "msg": str(e)}
