# -*- coding: utf-8 -*-

import bridge
from safe_output import new_output_file
from dataclasses import dataclass
import os
import shutil
import tempfile
import subprocess
import zipfile
import io
import sys
import time
import math
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Optional, Tuple

try:
    import pikepdf
except Exception:
    pikepdf = None

try:
    import fitz
except Exception:
    fitz = None

try:
    import numpy as np
except Exception:
    np = None

try:
    import cv2
except Exception:
    cv2 = None


def _copy_to_unique_output(source: Path, desired: Path) -> Path:
    # Compression may select the original file itself; copying ensures that
    # later edits to the result cannot alter the input via a shared hard link.
    with new_output_file(desired) as (output, actual_path):
        with source.open("rb") as inp:
            shutil.copyfileobj(inp, output)
    return actual_path


PDF_FLOOR_RATIO = 0.80   # PDF 不希望压得离目标太远
WORD_FLOOR_RATIO = 0.85  # Word 可以更激进一点

EMU_PER_INCH = 914400.0
MIN_SMART_IMAGE_BYTES = 20 * 1024  # 小于 20KB 的小图标/缩略图默认保护跳过


@dataclass(frozen=True)
class ImageClarityProfile:
    width: int
    height: int
    raw_bytes: int
    effective_dpi: Optional[float]
    laplacian_var: float
    tenengrad: float
    bpp: float
    sharpness_score: float
    resolution_score: float
    clarity_score: float
    tier: str  # "high" | "medium" | "low"
    skip_recommended: bool = False


def _clamp(val: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, val))


def _compute_sharpness_metrics(img) -> Tuple[float, float]:
    """
    在统一分析尺度下计算视觉锐度指标 (laplacian_var, tenengrad)。
    优先使用 OpenCV，缺少 OpenCV 时自动退化为 NumPy/Pillow 梯度估计。
    """
    w, h = img.size
    max_side = max(w, h)
    analysis_img = img.convert("L")
    if max_side > 640:
        scale = 640.0 / max_side
        new_w = max(16, int(round(w * scale)))
        new_h = max(16, int(round(h * scale)))
        from PIL import Image
        analysis_img = analysis_img.resize((new_w, new_h), Image.Resampling.BILINEAR)

    if np is None:
        return 120.0, 1200.0

    gray = np.asarray(analysis_img, dtype=np.uint8)
    if gray.shape[0] < 4 or gray.shape[1] < 4:
        return 0.0, 0.0

    if cv2 is not None:
        lap = cv2.Laplacian(gray, cv2.CV_64F)
        lap_var = float(lap.var())
        gx = cv2.Sobel(gray, cv2.CV_64F, 1, 0, ksize=3)
        gy = cv2.Sobel(gray, cv2.CV_64F, 0, 1, ksize=3)
        tenengrad = float(np.mean(gx * gx + gy * gy))
        return lap_var, tenengrad

    # 无 cv2 时的纯 NumPy 退化实现
    g = gray.astype(np.float64)
    lap = (
        g[:-2, 1:-1] + g[2:, 1:-1] + g[1:-1, :-2] + g[1:-1, 2:]
        - 4.0 * g[1:-1, 1:-1]
    )
    lap_var = float(lap.var())
    gx = g[:, 2:] - g[:, :-2]
    gy = g[2:, :] - g[:-2, :]
    tenengrad = float(np.mean(gx[1:-1, :] ** 2 + gy[:, 1:-1] ** 2))
    return lap_var, tenengrad


def analyze_image_clarity(
    img,
    raw_bytes: int,
    effective_dpi: Optional[float] = None,
) -> ImageClarityProfile:
    """
    综合评估单张图片的清晰度与压缩冗余度：
    - 视觉锐度（归一化尺度下的 Laplacian 方差 + Sobel Tenengrad 梯度能量）
    - 空间分辨率冗余（有效 DPI / 像素尺寸）
    - 字节密度 BPP（识别已高度压缩的低冗余图片）

    返回 clarity_score ∈ [0, 1] 与分级 tier:
    - "high"  : 高清/高冗余图片 -> 允许更高压缩比
    - "medium": 普通清晰度图片 -> 适度压缩
    - "low"   : 低清/已模糊/小图 -> 低压缩比保护或跳过
    """
    w, h = img.size
    max_edge = max(w, h)
    pixels = max(1, w * h)
    bpp = float(raw_bytes) / float(pixels)

    lap_var, tenengrad = _compute_sharpness_metrics(img)

    lap_norm = _clamp(
        (math.log1p(max(0.0, lap_var)) - math.log1p(15.0))
        / (math.log1p(900.0) - math.log1p(15.0)),
        0.0,
        1.0,
    )
    ten_norm = _clamp(
        (math.log1p(max(0.0, tenengrad)) - math.log1p(150.0))
        / (math.log1p(12000.0) - math.log1p(150.0)),
        0.0,
        1.0,
    )
    sharpness_score = 0.55 * lap_norm + 0.45 * ten_norm

    px_norm = _clamp((max_edge - 700.0) / (2600.0 - 700.0), 0.0, 1.0)
    if effective_dpi is not None and effective_dpi > 0:
        dpi_norm = _clamp((effective_dpi - 96.0) / (300.0 - 96.0), 0.0, 1.0)
        resolution_score = 0.65 * dpi_norm + 0.35 * px_norm
    else:
        resolution_score = px_norm

    bpp_score = _clamp((bpp - 0.12) / (0.85 - 0.12), 0.0, 1.0)

    clarity_score = 0.45 * sharpness_score + 0.40 * resolution_score + 0.15 * bpp_score

    skip_recommended = False

    # 护栏 1：非常小的图标/缩略图直接建议保护跳过
    if raw_bytes < MIN_SMART_IMAGE_BYTES or max_edge <= 320:
        skip_recommended = True
        clarity_score = min(clarity_score, 0.20)

    # 护栏 2：本身已经发糊且不是超大尺寸的图片，强制归入低清保护
    if sharpness_score < 0.22 and max_edge < 2400 and (effective_dpi is None or effective_dpi < 240):
        clarity_score = min(clarity_score, 0.33)

    # 护栏 3：低分辨率 / 低有效 DPI 图片避免继续降采样
    if max_edge <= 850 or (effective_dpi is not None and effective_dpi <= 115):
        clarity_score = min(clarity_score, 0.32)

    # 护栏 4：已经高压缩比的低 BPP 且非超大尺寸图片，减少二次压缩损耗
    if bpp < 0.14 and max_edge < 1800 and (effective_dpi is None or effective_dpi < 200):
        clarity_score = min(clarity_score, 0.34)

    clarity_score = _clamp(clarity_score, 0.0, 1.0)

    if clarity_score >= 0.60:
        tier = "high"
    elif clarity_score >= 0.36:
        tier = "medium"
    else:
        tier = "low"

    return ImageClarityProfile(
        width=w,
        height=h,
        raw_bytes=raw_bytes,
        effective_dpi=effective_dpi,
        laplacian_var=lap_var,
        tenengrad=tenengrad,
        bpp=bpp,
        sharpness_score=sharpness_score,
        resolution_score=resolution_score,
        clarity_score=clarity_score,
        tier=tier,
        skip_recommended=skip_recommended,
    )


def compute_smart_image_plan(
    profile: ImageClarityProfile,
    pressure: float,
) -> Tuple[bool, float, int, bool]:
    """
    根据全局压缩压力 pressure ∈ [0.0, 1.0] 与单图清晰度画像，计算单图专属压缩方案。

    返回: (should_compress, scale, jpeg_quality, quantize_png)
    - 高清图 (high): 立即介入、降采样幅度更大、JPEG quality 更激进
    - 普通图 (medium): 适度延迟、中等缩放与质量
    - 低清图 (low): 较高延迟（前期跳过不压）、尽量不缩放、保持高 JPEG quality
    """
    p = _clamp(float(pressure), 0.0, 1.0)
    if p <= 0.0 or profile.skip_recommended:
        return False, 1.0, 92, False

    max_edge = max(profile.width, profile.height)

    if profile.tier == "high":
        delay = 0.0
        strength = 1.0 + 0.25 * _clamp((profile.clarity_score - 0.60) / 0.40, 0.0, 1.0)
        p_eff = _clamp(((p - delay) / (1.0 - delay)) * strength, 0.0, 1.0)

        min_scale = 0.34
        if profile.effective_dpi and profile.effective_dpi > 0:
            min_scale = max(min_scale, min(0.95, 92.0 / profile.effective_dpi))
        if max_edge > 0:
            min_scale = max(min_scale, min(0.95, 840.0 / max_edge))

        scale = 1.0 - p_eff * (1.0 - min_scale)
        quality = int(round(84 - p_eff * 52))  # 84 -> 32
        quantize_png = p_eff >= 0.42
        return True, _clamp(scale, min_scale, 1.0), int(_clamp(quality, 30, 92)), quantize_png

    if profile.tier == "medium":
        delay = 0.14
        if p <= delay:
            return False, 1.0, 90, False
        strength = 0.72 + 0.23 * _clamp((profile.clarity_score - 0.36) / 0.24, 0.0, 1.0)
        p_eff = _clamp(((p - delay) / (1.0 - delay)) * strength, 0.0, 1.0)

        min_scale = 0.58
        if profile.effective_dpi and profile.effective_dpi > 0:
            min_scale = max(min_scale, min(0.98, 115.0 / profile.effective_dpi))
        if max_edge > 0:
            min_scale = max(min_scale, min(0.98, 960.0 / max_edge))

        scale = 1.0 - p_eff * (1.0 - min_scale)
        quality = int(round(88 - p_eff * 36))  # 88 -> 52
        quantize_png = p_eff >= 0.72
        return True, _clamp(scale, min_scale, 1.0), int(_clamp(quality, 50, 92)), quantize_png

    # tier == "low"：低清/模糊/小图保护策略
    delay = 0.48
    if p <= delay:
        return False, 1.0, 92, False
    strength = 0.28 + 0.27 * _clamp(profile.clarity_score / 0.36, 0.0, 1.0)
    p_eff = _clamp(((p - delay) / (1.0 - delay)) * strength, 0.0, 1.0)

    if max_edge <= 1100:
        scale = 1.0
    else:
        scale = 1.0 - max(0.0, p_eff - 0.40) * 0.28
        scale = max(0.84, scale)

    quality = int(round(92 - p_eff * 14))  # 92 -> 78
    return True, _clamp(scale, 0.84, 1.0), int(_clamp(quality, 76, 94)), False


def _encode_image_bytes(
    img,
    target_format: str,
    scale: float,
    quality: int,
    quantize_png: bool = False,
) -> bytes:
    """按指定缩放比与质量编码单张图片。"""
    from PIL import Image

    w, h = img.size
    out_img = img
    if scale < 0.995:
        new_w = max(16, int(round(w * scale)))
        new_h = max(16, int(round(h * scale)))
        if new_w < w or new_h < h:
            out_img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

    buf = io.BytesIO()
    fmt = target_format.upper()

    if fmt in ("JPG", "JPEG"):
        if out_img.mode not in ("RGB", "L"):
            out_img = out_img.convert("RGB")
        out_img.save(
            buf,
            format="JPEG",
            quality=int(quality),
            optimize=True,
            subsampling="4:2:0" if out_img.mode == "RGB" and quality < 85 else "keep",
        )
    else:
        save_img = out_img
        if quantize_png and save_img.mode in ("RGB", "RGBA"):
            try:
                if save_img.mode == "RGBA":
                    save_img = save_img.quantize(colors=256, method=Image.Quantize.FASTOCTREE)
                else:
                    save_img = save_img.quantize(colors=256, method=Image.Quantize.MEDIANCUT)
            except Exception:
                save_img = out_img
        save_img.save(buf, format="PNG", optimize=True, compress_level=9)

    return buf.getvalue()


def _extract_docx_image_display_inches(item_data: Dict[str, bytes]) -> Dict[str, Tuple[float, float]]:
    """
    解析 word/_rels/document.xml.rels 与 word/document.xml，
    提取每张 word/media/* 图片在文档中的最大显示尺寸（英寸），用于估算有效 DPI。
    """
    rels_bytes = item_data.get("word/_rels/document.xml.rels")
    doc_bytes = item_data.get("word/document.xml")
    if not rels_bytes or not doc_bytes:
        return {}

    rid_to_media: Dict[str, str] = {}
    try:
        rels_root = ET.fromstring(rels_bytes)
        for rel in rels_root:
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target", "")
            if not rid or not target:
                continue
            norm_target = target.replace("\\", "/").lstrip("/")
            if not norm_target.startswith("word/"):
                norm_target = f"word/{norm_target}"
            rid_to_media[rid] = norm_target
    except Exception:
        return {}

    if not rid_to_media:
        return {}

    display_map: Dict[str, Tuple[float, float]] = {}
    try:
        doc_root = ET.fromstring(doc_bytes)
        for elem in doc_root.iter():
            tag = elem.tag.rsplit("}", 1)[-1]
            if tag not in ("inline", "anchor"):
                continue

            cx_in = 0.0
            cy_in = 0.0
            embed_rid = None

            for sub in elem.iter():
                sub_tag = sub.tag.rsplit("}", 1)[-1]
                if sub_tag == "extent" and cx_in == 0.0 and cy_in == 0.0:
                    try:
                        cx = float(sub.attrib.get("cx", "0"))
                        cy = float(sub.attrib.get("cy", "0"))
                        if cx > 0 and cy > 0:
                            cx_in = cx / EMU_PER_INCH
                            cy_in = cy / EMU_PER_INCH
                    except ValueError:
                        pass
                elif sub_tag == "blip" and embed_rid is None:
                    for k, v in sub.attrib.items():
                        if k.rsplit("}", 1)[-1] == "embed":
                            embed_rid = v
                            break

            if embed_rid and cx_in > 0 and cy_in > 0:
                media_path = rid_to_media.get(embed_rid)
                if media_path:
                    prev = display_map.get(media_path)
                    if prev is None:
                        display_map[media_path] = (cx_in, cy_in)
                    else:
                        display_map[media_path] = (max(prev[0], cx_in), max(prev[1], cy_in))
    except Exception:
        return {}

    return display_map


def get_gs_path():
    """动态寻址 Ghostscript，优先使用打包目录中的运行时。"""
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent

    bundled_gs = base_dir / "Ghostscript" / "bin" / "gswin64c.exe"
    if bundled_gs.exists():
        return str(bundled_gs)

    return shutil.which("gswin64c") or r"C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe"


def ui_alive():
    """
    当前 pywebview 桥接没有独立连接池，默认视为前端可用。
    """
    return True


def safe_update_terminal(msg: str):
    try:
        bridge.update_terminal(msg)
    except Exception:
        pass


def fmt_size_factory(unit: str):
    def fmt_sz(b: int | float):
        return f"{b/1024:.2f} KB" if unit == "KB" else f"{b/(1024*1024):.2f} MB"
    return fmt_sz


def run_process_with_heartbeat(cmd, stage_name: str, heartbeat_sec: int = 3):
    """
    用 Popen 运行长任务，并定时输出心跳，避免大 PDF 压缩时看起来像假死。
    返回进程退出码。
    """
    proc = subprocess.Popen(
        cmd,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    start_t = time.time()
    last_report = 0

    try:
        while True:
            rc = proc.poll()
            if rc is not None:
                return rc

            if not ui_alive():
                try:
                    proc.terminate()
                except Exception:
                    pass
                return -999

            now = time.time()
            if now - last_report >= heartbeat_sec:
                elapsed = int(now - start_t)
                safe_update_terminal(f"    ⏳ {stage_name} 执行中... 已耗时 {elapsed}s")
                last_report = now

            time.sleep(1)

    finally:
        try:
            if proc.poll() is None:
                proc.terminate()
        except Exception:
            pass


def lossless_pdf_optimize(src_pdf: Path, out_pdf: Path) -> tuple[bool, str]:
    """
    使用 pikepdf 做一次快速无损瘦身：
    - 重组对象流
    - 重压缩流
    """
    if pikepdf is None:
        return False, "未安装 pikepdf，跳过无损瘦身"

    try:
        with pikepdf.open(str(src_pdf)) as pdf:
            pdf.save(
                str(out_pdf),
                object_stream_mode=pikepdf.ObjectStreamMode.generate,
                compress_streams=True,
                recompress_flate=True,
                linearize=False,
            )
        return True, "无损瘦身完成"
    except Exception as e:
        return False, f"无损瘦身失败: {e}"


def estimate_initial_dpi(current_size: int, target_size: int) -> int:
    """
    根据目标比例估算首轮 DPI，避免盲目多轮试探。
    """
    ratio = max(0.05, min(1.0, target_size / current_size))
    base_dpi = 220
    estimated = int(base_dpi * math.sqrt(ratio))
    return max(72, min(220, estimated))


def run_single_gs_compress(
    gs_path: str,
    input_pdf: Path,
    output_pdf: Path,
    dpi: int,
    fast_mode: bool,
    round_index: int,
    round_total: int,
):
    cmd = [
        gs_path,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        "-dPDFSETTINGS=/default",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        "-dCompressFonts=true",
        "-dDownsampleColorImages=true",
        f"-dColorImageResolution={dpi}",
        "-dColorImageDownsampleThreshold=1.0",
        "-dColorImageDownsampleType=/Bicubic",
        "-dDownsampleGrayImages=true",
        f"-dGrayImageResolution={dpi}",
        "-dGrayImageDownsampleThreshold=1.0",
        "-dGrayImageDownsampleType=/Bicubic",
        "-dDownsampleMonoImages=true",
        f"-dMonoImageResolution={dpi}",
        "-dMonoImageDownsampleThreshold=1.0",
        "-dMonoImageDownsampleType=/Bicubic",
        f"-sOutputFile={output_pdf.as_posix()}",
        input_pdf.as_posix(),
    ]

    # 中小文件可开重复图像检测；大文件为了速度关闭
    if not fast_mode:
        cmd.insert(7, "-dDetectDuplicateImages=true")

    return run_process_with_heartbeat(
        cmd,
        stage_name=f"Ghostscript 第 {round_index}/{round_total} 轮压缩",
        heartbeat_sec=3,
    )


def choose_best_candidate(candidates, target_b: int, floor_ratio: float):
    """
    统一选取策略：
    1. 优先选择 [target*floor, target] 区间内最接近 target 的结果
    2. 如果没有，则选择所有 <= target 中最大的那个
    3. 如果仍没有，则选择最接近 target 的结果（此时可能大于目标）
    """
    if not candidates:
        return None, "未找到有效压缩结果"

    floor_b = target_b * floor_ratio

    preferred = [c for c in candidates if floor_b <= c["size"] <= target_b]
    if preferred:
        best = min(preferred, key=lambda c: abs(c["size"] - target_b))
        return best, "已选取落在理想区间内、最接近目标的结果"

    under_target = [c for c in candidates if c["size"] <= target_b]
    if under_target:
        best = max(under_target, key=lambda c: c["size"])
        return best, "未命中理想区间，已选取低于目标且最接近目标的结果"

    best = min(candidates, key=lambda c: abs(c["size"] - target_b))
    return best, "所有结果均高于目标，已选取最接近目标的结果"


def _smart_compress_pdf_pass(
    work_input: Path,
    work_size: int,
    target_b: float,
    tmp_dir: Path,
    fmt_sz,
):
    """
    PDF 智能清晰度分级压缩：
    使用 PyMuPDF 提取唯一嵌入图片，结合有效 DPI 与视觉锐度分级后，
    通过内存预估二分查找最佳全局压力，只写回 1~2 次 PDF。
    返回候选列表 candidates（可能为空，此时调用方回退到 Ghostscript）。
    """
    if fitz is None:
        safe_update_terminal("[*] 未检测到 PyMuPDF，自动使用 Ghostscript 压缩引擎")
        return []

    try:
        from PIL import Image
    except ImportError:
        safe_update_terminal("[*] 缺少 Pillow，跳过智能单图压缩，改用 Ghostscript")
        return []

    candidates = []
    ideal_floor = target_b * PDF_FLOOR_RATIO

    safe_update_terminal("[*] 启动智能清晰度检测（按图分级：高清多压 / 低清保护）...")

    image_entries = []
    skipped_count = 0
    tier_counts = {"high": 0, "medium": 0, "low": 0}

    with fitz.open(str(work_input)) as doc:
        xref_meta: Dict[int, dict] = {}

        for page_idx, page in enumerate(doc):
            try:
                infos = page.get_image_info(xrefs=True)
            except Exception:
                infos = []

            for info in infos:
                xref = info.get("xref", 0)
                if not xref or xref <= 0:
                    continue
                bbox = info.get("bbox")
                bw_pt = max(0.0, float(bbox[2] - bbox[0])) if bbox else 0.0
                bh_pt = max(0.0, float(bbox[3] - bbox[1])) if bbox else 0.0

                entry = xref_meta.get(xref)
                if entry is None:
                    xref_meta[xref] = {
                        "page_idx": page_idx,
                        "max_bw_pt": bw_pt,
                        "max_bh_pt": bh_pt,
                    }
                else:
                    entry["max_bw_pt"] = max(entry["max_bw_pt"], bw_pt)
                    entry["max_bh_pt"] = max(entry["max_bh_pt"], bh_pt)

            # 兜底补充 get_images 里可能遗漏的 xref
            try:
                for img_tuple in page.get_images(full=True):
                    xref = img_tuple[0]
                    if xref > 0 and xref not in xref_meta:
                        xref_meta[xref] = {
                            "page_idx": page_idx,
                            "max_bw_pt": 0.0,
                            "max_bh_pt": 0.0,
                        }
            except Exception:
                pass

        for xref, meta in xref_meta.items():
            try:
                extracted = doc.extract_image(xref)
            except Exception:
                skipped_count += 1
                continue

            if not extracted or not extracted.get("image"):
                skipped_count += 1
                continue

            # 跳过带透明蒙版、1-bit 黑白扫描图（通常已是极小 CCITT/JBIG2）或 CMYK 四通道图
            if extracted.get("smask", 0) != 0:
                skipped_count += 1
                continue
            if extracted.get("bpc", 8) <= 1:
                skipped_count += 1
                continue
            if extracted.get("colorspace", 3) == 4:
                skipped_count += 1
                continue

            img_bytes = extracted["image"]
            try:
                stream_len = len(doc.xref_stream_raw(xref))
            except Exception:
                stream_len = len(img_bytes)

            if stream_len < MIN_SMART_IMAGE_BYTES:
                skipped_count += 1
                continue

            try:
                with Image.open(io.BytesIO(img_bytes)) as pil_img:
                    if pil_img.mode in ("CMYK", "RGBA", "LA", "PA"):
                        skipped_count += 1
                        continue
                    w, h = pil_img.size
                    eff_dpi = None
                    if meta["max_bw_pt"] > 1.0 and meta["max_bh_pt"] > 1.0:
                        dpi_x = w / (meta["max_bw_pt"] / 72.0)
                        dpi_y = h / (meta["max_bh_pt"] / 72.0)
                        eff_dpi = max(dpi_x, dpi_y)
                    profile = analyze_image_clarity(pil_img, stream_len, effective_dpi=eff_dpi)
            except Exception:
                skipped_count += 1
                continue

            tier_counts[profile.tier] += 1
            image_entries.append({
                "xref": xref,
                "page_idx": meta["page_idx"],
                "orig_stream_len": stream_len,
                "raw_bytes": img_bytes,
                "profile": profile,
            })

    if not image_entries:
        safe_update_terminal("[*] 未检测到适合单图重压缩的彩色/灰度位图，将直接启用 Ghostscript")
        return []

    safe_update_terminal(
        f"[*] PDF 清晰度预检完成：分析 {len(image_entries)} 张位图"
        f"（高清多压 {tier_counts['high']} 张 / 普通 {tier_counts['medium']} 张 / 低清保护 {tier_counts['low']} 张，"
        f"跳过透明/黑白/小图 {skipped_count} 张）"
    )

    encode_cache: Dict[Tuple[int, int, int], bytes] = {}

    def eval_pressure(pressure: float) -> Tuple[int, Dict[int, bytes], Dict[str, int]]:
        saved_bytes = 0
        replacements: Dict[int, bytes] = {}
        tier_replaced = {"high": 0, "medium": 0, "low": 0}

        for item in image_entries:
            profile = item["profile"]
            should_comp, scale, quality, _ = compute_smart_image_plan(profile, pressure)
            if not should_comp:
                continue

            scale_key = int(round(scale * 1000))
            cache_key = (item["xref"], scale_key, quality)
            new_bytes = encode_cache.get(cache_key)
            if new_bytes is None:
                try:
                    with Image.open(io.BytesIO(item["raw_bytes"])) as pil_img:
                        new_bytes = _encode_image_bytes(
                            pil_img,
                            target_format="JPEG",
                            scale=scale,
                            quality=quality,
                        )
                except Exception:
                    continue
                encode_cache[cache_key] = new_bytes

            if len(new_bytes) < item["orig_stream_len"]:
                saved_bytes += item["orig_stream_len"] - len(new_bytes)
                replacements[item["xref"]] = new_bytes
                tier_replaced[profile.tier] += 1

        est_size = max(1024, work_size - saved_bytes)
        return est_size, replacements, tier_replaced

    # 内存快速二分搜索压力值，找到最贴近目标体积的参数
    low_p, high_p = 0.05, 1.0
    best_plan = None

    for step in range(6):
        if not ui_alive():
            return []
        mid_p = (low_p + high_p) / 2.0
        est_size, replacements, tier_replaced = eval_pressure(mid_p)
        plan = {
            "pressure": mid_p,
            "est_size": est_size,
            "replacements": replacements,
            "tier_replaced": tier_replaced,
        }

        if best_plan is None:
            best_plan = plan
        else:
            # 优先落入 [ideal_floor, target_b]，其次 <= target_b 中最大，再次最接近 target_b
            def plan_rank(pl):
                sz = pl["est_size"]
                if ideal_floor <= sz <= target_b:
                    return (0, abs(sz - target_b))
                if sz <= target_b:
                    return (1, target_b - sz)
                return (2, sz - target_b)

            if plan_rank(plan) < plan_rank(best_plan):
                best_plan = plan

        if ideal_floor <= est_size <= target_b:
            break
        if est_size > target_b:
            low_p = mid_p
        else:
            high_p = mid_p

    # 若二分未能压到目标以下，补测一次最大压力 1.0
    if best_plan is None or best_plan["est_size"] > target_b:
        est_max, repl_max, tr_max = eval_pressure(1.0)
        if best_plan is None or est_max < best_plan["est_size"]:
            best_plan = {
                "pressure": 1.0,
                "est_size": est_max,
                "replacements": repl_max,
                "tier_replaced": tr_max,
            }

    if not best_plan or not best_plan["replacements"]:
        safe_update_terminal("[*] 智能单图分析未找到可有效瘦身的位图，将切换至 Ghostscript")
        return []

    def materialize_pdf(plan_dict, round_tag: str):
        out_path = tmp_dir / f"smart_pdf_{round_tag}.pdf"
        xref_to_page = {it["xref"]: it["page_idx"] for it in image_entries}
        with fitz.open(str(work_input)) as doc:
            for xref, new_bytes in plan_dict["replacements"].items():
                page_idx = xref_to_page[xref]
                try:
                    doc[page_idx].replace_image(xref, stream=new_bytes)
                except Exception:
                    continue
            doc.save(str(out_path), garbage=4, deflate=True, clean=False)

        if out_path.exists() and os.path.getsize(out_path) > 1024:
            actual_sz = os.path.getsize(out_path)
            tr = plan_dict["tier_replaced"]
            safe_update_terminal(
                f"[*] PDF 智能分级写回({round_tag}): 压力={int(plan_dict['pressure']*100)}% | "
                f"实际处理 高清 {tr['high']} / 普通 {tr['medium']} / 低清 {tr['low']} 张 -> {fmt_sz(actual_sz)}"
            )
            return {
                "path": out_path,
                "size": actual_sz,
                "param": f"智能分级(压力 {int(plan_dict['pressure']*100)}%)",
                "pressure": plan_dict["pressure"],
            }
        return None

    cand1 = materialize_pdf(best_plan, "r1")
    if cand1:
        candidates.append(cand1)

        # 若实际写回大小偏离预估且不在理想区间，按真实偏差做一次校准写回
        actual_sz = cand1["size"]
        overhead = actual_sz - best_plan["est_size"]
        if not (ideal_floor <= actual_sz <= target_b):
            if actual_sz > target_b and best_plan["pressure"] < 0.98:
                adjusted_target = max(1024.0, target_b - max(0, overhead))
                lo, hi = best_plan["pressure"], 1.0
                refined = None
                for _ in range(4):
                    mp = (lo + hi) / 2.0
                    esz, repl, tr = eval_pressure(mp)
                    refined = {"pressure": mp, "est_size": esz, "replacements": repl, "tier_replaced": tr}
                    if esz > adjusted_target:
                        lo = mp
                    else:
                        hi = mp
                if refined and refined["replacements"]:
                    cand2 = materialize_pdf(refined, "r2")
                    if cand2:
                        candidates.append(cand2)
            elif actual_sz < ideal_floor and best_plan["pressure"] > 0.08:
                adjusted_floor = max(1024.0, ideal_floor - overhead)
                lo, hi = 0.02, best_plan["pressure"]
                refined = None
                for _ in range(4):
                    mp = (lo + hi) / 2.0
                    esz, repl, tr = eval_pressure(mp)
                    refined = {"pressure": mp, "est_size": esz, "replacements": repl, "tier_replaced": tr}
                    if esz < adjusted_floor:
                        hi = mp
                    else:
                        lo = mp
                if refined and refined["replacements"]:
                    cand2 = materialize_pdf(refined, "r2")
                    if cand2:
                        candidates.append(cand2)

    return candidates


def _smart_compress_docx(
    in_p: Path,
    original_input: Path,
    target_b: float,
    tmp_dir: Path,
    fmt_sz,
):
    """Word (.docx) 智能清晰度分级压缩引擎。"""
    try:
        from PIL import Image
    except ImportError:
        return {
            "status": "error",
            "msg": "缺少图片处理库，请先在终端执行: pip install Pillow",
        }

    original_size = os.path.getsize(in_p)
    safe_update_terminal(f"[*] 原始文件大小: {fmt_sz(original_size)} | 目标大小: {fmt_sz(target_b)}")

    if original_size <= target_b:
        return {"status": "success", "msg": "原文件已不大于目标大小，无需压缩。"}

    with zipfile.ZipFile(in_p, "r") as zin:
        item_list = zin.infolist()
        item_data = {item.filename: zin.read(item.filename) for item in item_list}

    display_inches = _extract_docx_image_display_inches(item_data)

    media_profiles: Dict[str, ImageClarityProfile] = {}
    tier_counts = {"high": 0, "medium": 0, "low": 0}

    for item in item_list:
        fname = item.filename
        if fname.startswith("word/media/") and fname.lower().endswith((".png", ".jpg", ".jpeg")):
            raw = item_data[fname]
            try:
                with Image.open(io.BytesIO(raw)) as img:
                    w, h = img.size
                    eff_dpi = None
                    disp = display_inches.get(fname)
                    if disp and disp[0] > 0.05 and disp[1] > 0.05:
                        eff_dpi = max(w / disp[0], h / disp[1])
                    profile = analyze_image_clarity(img, len(raw), effective_dpi=eff_dpi)
                    media_profiles[fname] = profile
                    tier_counts[profile.tier] += 1
            except Exception:
                continue

    if not media_profiles:
        safe_update_terminal("[*] 文档中未检测到可压缩图片，尝试无损重打包...")
        out_t = tmp_dir / "repacked.docx"
        with zipfile.ZipFile(out_t, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for item in item_list:
                zout.writestr(item, item_data[item.filename])
        final = original_input.parent / f"{original_input.stem}_compressed.docx"
        final = _copy_to_unique_output(out_t, final)
        return {"status": "success", "msg": f"压缩完成：{final.name}"}

    safe_update_terminal(
        f"[*] Word 清晰度预检完成：共 {len(media_profiles)} 张图片"
        f"（高清多压 {tier_counts['high']} 张 / 普通 {tier_counts['medium']} 张 / 低清保护 {tier_counts['low']} 张）"
    )

    ideal_floor = target_b * WORD_FLOOR_RATIO
    low_p, high_p = 0.05, 1.0
    candidates = []
    encode_cache: Dict[Tuple[str, int, int, bool], bytes] = {}
    max_iters = 6

    for i in range(max_iters):
        if not ui_alive():
            return {"status": "error", "msg": "UI 已关闭，操作终止。"}

        mid_p = (low_p + high_p) / 2.0
        out_t = tmp_dir / f"smart_w_{i+1}_{int(mid_p*100)}.docx"
        tier_replaced = {"high": 0, "medium": 0, "low": 0}

        with zipfile.ZipFile(out_t, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in item_list:
                fname = item.filename
                data = item_data[fname]
                profile = media_profiles.get(fname)

                if profile is None:
                    zout.writestr(item, data)
                    continue

                should_comp, scale, quality, quantize_png = compute_smart_image_plan(profile, mid_p)
                if not should_comp:
                    zout.writestr(item, data)
                    continue

                fmt = "JPEG" if fname.lower().endswith((".jpg", ".jpeg")) else "PNG"
                cache_key = (fname, int(round(scale * 1000)), quality, quantize_png if fmt == "PNG" else False)
                encoded = encode_cache.get(cache_key)
                if encoded is None:
                    try:
                        with Image.open(io.BytesIO(data)) as img:
                            encoded = _encode_image_bytes(
                                img,
                                target_format=fmt,
                                scale=scale,
                                quality=quality,
                                quantize_png=(quantize_png if fmt == "PNG" else False),
                            )
                    except Exception:
                        encoded = data
                    encode_cache[cache_key] = encoded

                # 防膨胀保护：只有确实变小才替换
                if len(encoded) < len(data):
                    zout.writestr(item, encoded)
                    tier_replaced[profile.tier] += 1
                else:
                    zout.writestr(item, data)

        if out_t.exists():
            current_size = os.path.getsize(out_t)
            safe_update_terminal(
                f"[*] Word 智能迭代 {i+1}/{max_iters}: 压力={int(mid_p*100)}% | "
                f"实际压缩 高清 {tier_replaced['high']} / 普通 {tier_replaced['medium']} / 低清 {tier_replaced['low']} 张"
            )
            safe_update_terminal(
                f" └─ 当前参数产出大小: {fmt_sz(current_size)} (目标大小: {fmt_sz(target_b)})"
            )

            candidates.append({
                "path": out_t,
                "size": current_size,
                "param": f"智能分级(压力 {int(mid_p*100)}%)",
            })

            if ideal_floor <= current_size <= target_b:
                safe_update_terminal("[*] 已命中 Word 理想压缩区间，停止继续迭代")
                break

            if current_size > target_b:
                low_p = mid_p
            else:
                high_p = mid_p

    # 若仍全部高于目标大小，补测一次满压力 1.0
    if candidates and all(c["size"] > target_b for c in candidates):
        out_max = tmp_dir / "smart_w_max.docx"
        tier_replaced = {"high": 0, "medium": 0, "low": 0}
        with zipfile.ZipFile(out_max, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in item_list:
                fname = item.filename
                data = item_data[fname]
                profile = media_profiles.get(fname)
                if profile is None:
                    zout.writestr(item, data)
                    continue
                should_comp, scale, quality, quantize_png = compute_smart_image_plan(profile, 1.0)
                if not should_comp:
                    zout.writestr(item, data)
                    continue
                fmt = "JPEG" if fname.lower().endswith((".jpg", ".jpeg")) else "PNG"
                cache_key = (fname, int(round(scale * 1000)), quality, quantize_png if fmt == "PNG" else False)
                encoded = encode_cache.get(cache_key)
                if encoded is None:
                    try:
                        with Image.open(io.BytesIO(data)) as img:
                            encoded = _encode_image_bytes(
                                img,
                                target_format=fmt,
                                scale=scale,
                                quality=quality,
                                quantize_png=(quantize_png if fmt == "PNG" else False),
                            )
                    except Exception:
                        encoded = data
                    encode_cache[cache_key] = encoded
                if len(encoded) < len(data):
                    zout.writestr(item, encoded)
                    tier_replaced[profile.tier] += 1
                else:
                    zout.writestr(item, data)
        if out_max.exists():
            sz_max = os.path.getsize(out_max)
            safe_update_terminal(
                f"[*] Word 智能满压力兜底: 压力=100% | "
                f"实际压缩 高清 {tier_replaced['high']} / 普通 {tier_replaced['medium']} / 低清 {tier_replaced['low']} 张 -> {fmt_sz(sz_max)}"
            )
            candidates.append({
                "path": out_max,
                "size": sz_max,
                "param": "智能分级(压力 100%)",
            })

    best, choose_msg = choose_best_candidate(
        candidates=candidates,
        target_b=int(target_b),
        floor_ratio=WORD_FLOOR_RATIO,
    )

    final = original_input.parent / f"{original_input.stem}_compressed.docx"

    if best and best["path"].exists():
        final = _copy_to_unique_output(best["path"], final)

        safe_update_terminal(f"[*] {choose_msg}")
        safe_update_terminal(
            f"[*] 最终选中结果: {fmt_sz(best['size'])} | 策略={best['param']}"
        )

        if best["size"] > target_b:
            return {
                "status": "success",
                "msg": f"压缩完成：{final.name}（未低于目标，但已是最接近结果）",
            }

        return {"status": "success", "msg": f"压缩完成：{final.name}"}

    return {"status": "error", "msg": "Word 压缩失败。"}


@bridge.expose
def run_compress(path_str, target_size, unit, smart_mode=True):
    try:
        if not str(path_str or "").strip():
            return {"status": "error", "msg": "请选择要压缩的 PDF 或 Word 文件。"}
        original_input = Path(str(path_str).strip()).resolve()
        if not original_input.is_file():
            return {"status": "error", "msg": "文件不存在。"}

        try:
            target_val = float(target_size)
        except (TypeError, ValueError):
            return {"status": "error", "msg": "目标大小必须是有效数字。"}
        if target_val <= 0:
            return {"status": "error", "msg": "目标大小必须大于 0。"}

        use_smart = bool(smart_mode)
        in_p = original_input
        target_b = target_val * (1024 if unit == "KB" else 1024 * 1024)
        ext = in_p.suffix.lower()
        fmt_sz = fmt_size_factory(unit)

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)

            # ================== 格式预处理：旧版 .doc 自动升级 ==================
            if ext == ".doc":
                safe_update_terminal("[*] 检测到旧版 .doc，正在后台自动升级为 .docx...")

                import pythoncom
                import win32com.client

                word = None
                pythoncom.CoInitialize()
                try:
                    word = win32com.client.DispatchEx("Word.Application")
                    word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                    word.Visible = False
                    doc = word.Documents.Open(str(in_p))
                    temp_docx = tmp_dir / f"{original_input.stem}_temp.docx"
                    doc.SaveAs(str(temp_docx), FileFormat=16)
                    doc.Close(0)
                    in_p = temp_docx
                    ext = ".docx"
                    safe_update_terminal("[*] 格式升级完毕，准备进入 Word 压缩引擎...")
                finally:
                    try:
                        if word is not None:
                            word.Quit()
                    except Exception:
                        pass
                    pythoncom.CoUninitialize()

            # ================== 引擎 A：PDF 压缩逻辑 ==================
            if ext == ".pdf":
                mode_label = "智能清晰度分级" if use_smart else "传统统一 DPI"
                safe_update_terminal(f"[*] 智能路由：启动 PDF 压缩引擎（{mode_label}）")

                original_size = os.path.getsize(in_p)
                safe_update_terminal(f"[*] 原始文件大小: {fmt_sz(original_size)} | 目标大小: {fmt_sz(target_b)}")

                if original_size <= target_b:
                    return {"status": "success", "msg": "原文件已不大于目标大小，无需压缩。"}

                # 第一步：无损瘦身
                lossless_pdf = tmp_dir / "lossless_optimized.pdf"
                ok, info = lossless_pdf_optimize(in_p, lossless_pdf)
                safe_update_terminal(f"[*] {info}")

                work_input = in_p
                work_size = original_size

                if ok and lossless_pdf.exists() and os.path.getsize(lossless_pdf) > 1024:
                    lossless_size = os.path.getsize(lossless_pdf)
                    safe_update_terminal(f" └─ 无损瘦身结果: {fmt_sz(lossless_size)}")

                    if lossless_size < original_size:
                        work_input = lossless_pdf
                        work_size = lossless_size
                        safe_update_terminal("[*] 已采用无损瘦身结果作为后续输入")
                    else:
                        safe_update_terminal("[*] 无损瘦身收益有限，保留原文件继续处理")

                if work_size <= target_b:
                    final = original_input.parent / f"{original_input.stem}_compressed.pdf"
                    final = _copy_to_unique_output(work_input, final)
                    return {"status": "success", "msg": f"压缩完成：{final.name}"}

                candidates = []
                ideal_floor = target_b * PDF_FLOOR_RATIO

                # 第二步（智能模式）：按图清晰度分级压缩
                if use_smart:
                    smart_candidates = _smart_compress_pdf_pass(
                        work_input=work_input,
                        work_size=work_size,
                        target_b=target_b,
                        tmp_dir=tmp_dir,
                        fmt_sz=fmt_sz,
                    )
                    candidates.extend(smart_candidates)

                smart_best, _ = choose_best_candidate(
                    candidates=candidates,
                    target_b=int(target_b),
                    floor_ratio=PDF_FLOOR_RATIO,
                )
                need_gs = not (smart_best and smart_best["size"] <= target_b)

                # 第三步：Ghostscript 统一压缩（传统模式主路径，或智能模式未达目标时的兜底）
                if need_gs:
                    gs = get_gs_path()
                    if not os.path.exists(gs):
                        if not candidates:
                            return {
                                "status": "error",
                                "msg": "找不到 Ghostscript 环境。请检查打包是否遗漏了 Ghostscript 目录。",
                            }
                        safe_update_terminal("[!] 未找到 Ghostscript，将直接采用智能分级压缩结果")
                    else:
                        if use_smart and candidates:
                            safe_update_terminal(
                                "[*] 智能分级压缩后仍高于目标大小，自动衔接 Ghostscript 深度压缩兜底..."
                            )

                        size_mb = work_size / (1024 * 1024)
                        fast_mode = size_mb >= 80
                        if fast_mode:
                            safe_update_terminal("[*] 检测到大体积 PDF，自动启用快速压缩策略")

                        max_iters = 3
                        dpi_round1 = estimate_initial_dpi(work_size, target_b)
                        dpis = [dpi_round1]

                        for i in range(max_iters):
                            if not ui_alive():
                                return {"status": "error", "msg": "UI 已关闭，操作终止。"}

                            dpi = dpis[i] if i < len(dpis) else dpis[-1]
                            safe_update_terminal(f"[*] PDF 迭代 {i+1}/{max_iters}: 预计 DPI={dpi}")

                            out_t = tmp_dir / f"gs_round_{i+1}_{dpi}.pdf"

                            # 始终基于 work_input 重采样，避免“回提 DPI”时基于已降采样文件二次劣化
                            rc = run_single_gs_compress(
                                gs_path=gs,
                                input_pdf=work_input,
                                output_pdf=out_t,
                                dpi=dpi,
                                fast_mode=fast_mode,
                                round_index=i + 1,
                                round_total=max_iters,
                            )

                            if rc == -999:
                                return {"status": "error", "msg": "UI 已关闭，操作终止。"}

                            if rc != 0:
                                safe_update_terminal("[!] Ghostscript 本轮运行异常，停止继续迭代。")
                                break

                            if not out_t.exists() or os.path.getsize(out_t) <= 1024:
                                safe_update_terminal("[!] 本轮未生成有效 PDF，停止继续迭代。")
                                break

                            out_size = os.path.getsize(out_t)
                            safe_update_terminal(
                                f" └─ 当前参数产出大小: {fmt_sz(out_size)} (目标大小: {fmt_sz(target_b)})"
                            )

                            candidates.append({
                                "path": out_t,
                                "size": out_size,
                                "param": f"Ghostscript DPI={dpi}",
                            })

                            if ideal_floor <= out_size <= target_b:
                                safe_update_terminal("[*] 已命中理想压缩区间，停止继续迭代")
                                break

                            shrink_ratio = 1 - (out_size / work_size)
                            if i == 0 and shrink_ratio < 0.08:
                                safe_update_terminal("[*] 首轮压缩收益有限，停止继续迭代")
                                break

                            if out_size > target_b:
                                next_dpi = int(max(50, min(dpi - 8, dpi * math.sqrt(target_b / out_size))))
                            elif out_size < ideal_floor:
                                next_dpi = int(min(240, max(dpi + 10, dpi * math.sqrt(ideal_floor / out_size))))
                            else:
                                break

                            if next_dpi == dpi:
                                next_dpi = max(50, dpi - 8) if out_size > target_b else min(240, dpi + 8)

                            dpis.append(next_dpi)

                best, choose_msg = choose_best_candidate(
                    candidates=candidates,
                    target_b=int(target_b),
                    floor_ratio=PDF_FLOOR_RATIO,
                )

                if best and best["path"].exists():
                    final = original_input.parent / f"{original_input.stem}_compressed.pdf"
                    final = _copy_to_unique_output(best["path"], final)

                    safe_update_terminal(f"[*] {choose_msg}")
                    safe_update_terminal(
                        f"[*] 最终选中结果: {fmt_sz(best['size'])} | 策略={best['param']}"
                    )

                    if best["size"] > target_b:
                        return {
                            "status": "success",
                            "msg": f"压缩完成：{final.name}（未低于目标，但已是最接近结果）",
                        }

                    return {"status": "success", "msg": f"压缩完成：{final.name}"}

                return {
                    "status": "error",
                    "msg": "PDF 压缩失败：未能生成有效的 PDF 文件，请检查原文件是否加密或损坏。",
                }

            # ================== 引擎 B：Word (.docx) 压缩逻辑 ==================
            elif ext == ".docx":
                if use_smart:
                    safe_update_terminal("[*] 智能路由：启动 Word 智能清晰度分级压缩引擎")
                    return _smart_compress_docx(
                        in_p=in_p,
                        original_input=original_input,
                        target_b=target_b,
                        tmp_dir=tmp_dir,
                        fmt_sz=fmt_sz,
                    )

                safe_update_terminal("[*] 智能路由：启动 Word 传统压缩引擎 (Pillow)")

                try:
                    from PIL import Image
                except ImportError:
                    return {
                        "status": "error",
                        "msg": "缺少图片处理库，请先在终端执行: pip install Pillow",
                    }

                original_size = os.path.getsize(in_p)
                safe_update_terminal(f"[*] 原始文件大小: {fmt_sz(original_size)} | 目标大小: {fmt_sz(target_b)}")

                if original_size <= target_b:
                    return {"status": "success", "msg": "原文件已不大于目标大小，无需压缩。"}

                low_q, high_q = 10, 95
                last_size = 0
                mid_q = (low_q + high_q) // 2
                candidates = []
                ideal_floor = target_b * WORD_FLOOR_RATIO

                with zipfile.ZipFile(in_p, "r") as zin:
                    item_list = zin.infolist()
                    item_data = {item.filename: zin.read(item.filename) for item in item_list}

                for i in range(5):
                    if not ui_alive():
                        return {"status": "error", "msg": "UI 已关闭，操作终止。"}

                    if last_size > 0:
                        ratio = target_b / last_size
                        smart_q = int(mid_q * ratio)
                        mid_q = (low_q + high_q + smart_q * 2) // 4
                        mid_q = max(low_q, min(high_q, mid_q))
                    else:
                        mid_q = (low_q + high_q) // 2

                    safe_update_terminal(f"[*] Word 迭代 {i+1}/5: 智能调整图片压缩率={mid_q}...")

                    out_t = tmp_dir / f"t_{mid_q}.docx"
                    max_dim = int(2500 * (mid_q / 100)) + 400

                    with zipfile.ZipFile(out_t, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                        for item in item_list:
                            fname = item.filename
                            data = item_data[fname]

                            if fname.startswith("word/media/") and fname.lower().endswith(
                                (".png", ".jpg", ".jpeg")
                            ):
                                try:
                                    img = Image.open(io.BytesIO(data))
                                    img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
                                    img_byte_arr = io.BytesIO()

                                    if fname.lower().endswith((".jpg", ".jpeg")):
                                        if img.mode in ("RGBA", "P"):
                                            img = img.convert("RGB")
                                        img.save(
                                            img_byte_arr,
                                            format="JPEG",
                                            quality=mid_q,
                                            optimize=True,
                                        )
                                    elif fname.lower().endswith(".png"):
                                        img.save(img_byte_arr, format="PNG", optimize=True)

                                    new_bytes = img_byte_arr.getvalue()
                                    zout.writestr(item, new_bytes if len(new_bytes) < len(data) else data)
                                except Exception:
                                    zout.writestr(item, data)
                            else:
                                zout.writestr(item, data)

                    if out_t.exists():
                        current_size = os.path.getsize(out_t)
                        last_size = current_size

                        safe_update_terminal(
                            f" └─ 当前参数产出大小: {fmt_sz(current_size)} (目标大小: {fmt_sz(target_b)})"
                        )

                        candidates.append({
                            "path": out_t,
                            "size": current_size,
                            "param": mid_q,
                        })

                        if ideal_floor <= current_size <= target_b:
                            safe_update_terminal("[*] 已命中 Word 理想压缩区间，停止继续迭代")
                            break

                        if current_size <= target_b:
                            low_q = mid_q + 1
                        else:
                            high_q = mid_q - 1

                best, choose_msg = choose_best_candidate(
                    candidates=candidates,
                    target_b=int(target_b),
                    floor_ratio=WORD_FLOOR_RATIO,
                )

                final = original_input.parent / f"{original_input.stem}_compressed.docx"

                if best and best["path"].exists():
                    final = _copy_to_unique_output(best["path"], final)

                    safe_update_terminal(f"[*] {choose_msg}")
                    safe_update_terminal(
                        f"[*] 最终选中结果: {fmt_sz(best['size'])} | 使用压缩率={best['param']}"
                    )

                    if best["size"] > target_b:
                        return {
                            "status": "success",
                            "msg": f"压缩完成：{final.name}（未低于目标，但已是最接近结果）",
                        }

                    return {"status": "success", "msg": f"压缩完成：{final.name}"}

                return {"status": "error", "msg": "Word 压缩失败。"}

            else:
                return {"status": "error", "msg": f"不支持的文件格式: {ext}"}

    except Exception as e:
        return {"status": "error", "msg": str(e)}
