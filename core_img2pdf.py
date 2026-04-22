# -*- coding: utf-8 -*-
import bridge, io, re, time, threading, img2pdf
from pathlib import Path
from PIL import Image


def _start_heartbeat(label, state, interval=5):
    stop_event = threading.Event()
    start_time = time.time()

    def worker():
        while not stop_event.wait(interval):
            stage = state.get("stage", "处理中")
            current = state.get("current")
            total = state.get("total")
            elapsed = int(time.time() - start_time)
            if current is not None and total:
                bridge.update_terminal(f"⏳ [{label}] {stage}: {current}/{total}，已耗时 {elapsed}s")
            else:
                bridge.update_terminal(f"⏳ [{label}] {stage}，已耗时 {elapsed}s")

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    return stop_event, thread


def _stop_heartbeat(stop_event, thread):
    stop_event.set()
    try:
        thread.join(timeout=1)
    except Exception:
        pass


# 透明 PNG 补白逻辑
def _png_to_white_jpeg_bytes(p):
    with Image.open(p) as im:
        if im.mode in ("RGBA", "LA"):
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im, mask=im.split()[-1])
            out = bg
        else:
            out = im.convert("RGB")
        buf = io.BytesIO()
        out.save(buf, format="JPEG", quality=95)
        buf.seek(0)
        return buf


@bridge.expose
def run_img2pdf(path, recursive, include_root):
    try:
        root_folder = Path(path).resolve()
        if not root_folder.is_dir():
            return {"status": "error", "msg": "请选择有效的图片文件夹"}

        targets = [d for d in root_folder.rglob("*") if d.is_dir()] if recursive else [root_folder]
        if recursive and include_root:
            targets.insert(0, root_folder)

        converted_count = 0
        for folder in targets:
            imgs = sorted(
                [p for p in folder.iterdir() if p.suffix.lower() in (".jpg", ".jpeg", ".png", ".bmp")],
                key=lambda p: [int(t) if t.isdigit() else t.lower() for t in re.split(r"([0-9]+)", p.name)],
            )
            if not imgs:
                continue

            state = {"stage": "准备图片", "current": 0, "total": len(imgs)}
            stop_event, heartbeat_thread = _start_heartbeat(folder.name, state)
            try:
                bridge.update_terminal(f"[*] 正在打包目录: {folder.name}，共 {len(imgs)} 张图片")
                inputs = []
                for idx, img in enumerate(imgs, 1):
                    state.update({"stage": "整理图片", "current": idx, "total": len(imgs)})
                    if idx % 20 == 0 or idx == len(imgs):
                        bridge.update_terminal(f"  └─ 已准备图片 {idx}/{len(imgs)}")
                    inputs.append(_png_to_white_jpeg_bytes(img) if img.suffix.lower() == ".png" else str(img))

                state.update({"stage": "写入 PDF", "current": None, "total": None})
                out_file = folder / f"{folder.name}.pdf"
                with open(out_file, "wb") as f:
                    img2pdf.convert(inputs, outputstream=f)
                bridge.update_terminal(f"  └─ ✅ 已生成: {out_file.name}")
                converted_count += 1
            finally:
                _stop_heartbeat(stop_event, heartbeat_thread)

        if converted_count == 0:
            return {"status": "error", "msg": "未找到可转换的图片文件"}

        return {"status": "success", "msg": f"共生成 {converted_count} 个 PDF"}
    except Exception as e:
        return {"status": "error", "msg": str(e)}
