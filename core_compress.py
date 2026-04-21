# -*- coding: utf-8 -*-

import eel
import os
import shutil
import tempfile
import subprocess
import zipfile
import io
import sys
import time
import math
from pathlib import Path

try:
    import pikepdf
except Exception:
    pikepdf = None


PDF_FLOOR_RATIO = 0.80   # PDF 不希望压得离目标太远
WORD_FLOOR_RATIO = 0.85  # Word 可以更激进一点


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
    兼容旧 Eel 与当前 pywebview 桥接：
    - 旧 Eel: 通过 _websockets 判断前端是否仍在线
    - 新桥接: 没有 _websockets，默认视为前端可用
    """
    ws = getattr(eel, "_websockets", None)
    if ws is None:
        return True
    try:
        return bool(ws)
    except Exception:
        return True


def safe_update_terminal(msg: str):
    try:
        eel.update_terminal(msg)
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


@eel.expose
def run_compress(path_str, target_size, unit):
    try:
        original_input = Path(path_str.strip()).resolve()
        if not original_input.exists():
            return {"status": "error", "msg": "文件不存在。"}

        in_p = original_input
        target_b = float(target_size) * (1024 if unit == "KB" else 1024 * 1024)
        ext = in_p.suffix.lower()
        fmt_sz = fmt_size_factory(unit)

        with tempfile.TemporaryDirectory() as tmp:
            # ================== 格式预处理：旧版 .doc 自动升级 ==================
            if ext == ".doc":
                safe_update_terminal("[*] 检测到旧版 .doc，正在后台自动升级为 .docx...")

                import pythoncom
                import win32com.client

                word = None
                pythoncom.CoInitialize()
                try:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(str(in_p))
                    temp_docx = Path(tmp) / f"{original_input.stem}_temp.docx"
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
                safe_update_terminal("[*] 智能路由：启动 PDF 压缩引擎")

                gs = get_gs_path()
                if not os.path.exists(gs):
                    return {
                        "status": "error",
                        "msg": "找不到 Ghostscript 环境。请检查打包是否遗漏了 Ghostscript 目录。",
                    }

                original_size = os.path.getsize(in_p)
                safe_update_terminal(f"[*] 原始文件大小: {fmt_sz(original_size)} | 目标大小: {fmt_sz(target_b)}")

                if original_size <= target_b:
                    return {"status": "success", "msg": "原文件已不大于目标大小，无需压缩。"}

                # 第一步：无损瘦身
                lossless_pdf = Path(tmp) / "lossless_optimized.pdf"
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
                    shutil.copy(work_input, final)
                    return {"status": "success", "msg": f"压缩完成：{final.name}"}

                # 第二步：Ghostscript 有损压缩
                size_mb = work_size / (1024 * 1024)
                fast_mode = size_mb >= 80

                if fast_mode:
                    safe_update_terminal("[*] 检测到大体积 PDF，自动启用快速压缩策略")

                max_iters = 3
                candidates = []

                current_input = work_input
                current_size = work_size

                dpi_round1 = estimate_initial_dpi(work_size, target_b)
                dpis = [dpi_round1]

                ideal_floor = target_b * PDF_FLOOR_RATIO

                for i in range(max_iters):
                    if not ui_alive():
                        return {"status": "error", "msg": "UI 已关闭，操作终止。"}

                    dpi = dpis[i] if i < len(dpis) else dpis[-1]
                    safe_update_terminal(f"[*] PDF 迭代 {i+1}/{max_iters}: 预计 DPI={dpi}")

                    out_t = Path(tmp) / f"gs_round_{i+1}_{dpi}.pdf"

                    rc = run_single_gs_compress(
                        gs_path=gs,
                        input_pdf=current_input,
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
                        "param": dpi,
                    })

                    # 命中理想区间就提前结束
                    if ideal_floor <= out_size <= target_b:
                        safe_update_terminal("[*] 已命中理想压缩区间，停止继续迭代")
                        break

                    # 第一轮如果收益很小，继续压缩通常也没太大意义
                    shrink_ratio = 1 - (out_size / current_size)
                    if i == 0 and shrink_ratio < 0.08:
                        safe_update_terminal("[*] 首轮压缩收益有限，停止继续迭代")
                        break

                    # 自适应下一轮：
                    # 1. 还大于目标 => 继续降 DPI
                    # 2. 已经太小（低于理想下限）=> 回提 DPI
                    current_input = out_t
                    current_size = out_size

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
                    shutil.copy(best["path"], final)

                    safe_update_terminal(f"[*] {choose_msg}")
                    safe_update_terminal(
                        f"[*] 最终选中结果: {fmt_sz(best['size'])} | 使用 DPI={best['param']}"
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
                safe_update_terminal("[*] 智能路由：启动 Word 压缩引擎 (Pillow)")

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

                    out_t = Path(tmp) / f"t_{mid_q}.docx"
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

                                    zout.writestr(item, img_byte_arr.getvalue())
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

                        # 命中 Word 理想区间就提前结束
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
                    shutil.copy(best["path"], final)

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
