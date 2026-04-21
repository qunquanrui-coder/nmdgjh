# -*- coding: utf-8 -*-
import eel, os, shutil, tempfile, subprocess, zipfile, io, sys
from pathlib import Path

def get_gs_path():
    """✨ 核心修复：动态寻址，优先使用 PyInstaller 打包时复制到本地的组件环境"""
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent
        
    bundled_gs = base_dir / "Ghostscript" / "bin" / "gswin64c.exe"
    if bundled_gs.exists():
        return str(bundled_gs)
    # 兜底寻找系统环境变量
    return shutil.which("gswin64c") or r"C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe"

@eel.expose
def run_compress(path_str, target_size, unit):
    try:
        in_p = Path(path_str.strip()).resolve()
        target_b = float(target_size) * (1024 if unit == "KB" else 1024 * 1024)
        ext = in_p.suffix.lower()
        
        def fmt_sz(b): return f"{b/1024:.2f} KB" if unit == "KB" else f"{b/(1024*1024):.2f} MB"
        
        with tempfile.TemporaryDirectory() as tmp:
            
            # ================== 格式预处理：旧版 .doc 自动升级 ==================
            if ext == '.doc':
                eel.update_terminal("[*] 检测到旧版 .doc，正在后台自动升级为 .docx...")
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                try:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(str(in_p))
                    temp_docx = Path(tmp) / f"{in_p.stem}_temp.docx"
                    doc.SaveAs(str(temp_docx), FileFormat=16) 
                    doc.Close(0)
                    in_p = temp_docx  
                    ext = '.docx'     
                    eel.update_terminal("[*] 格式升级完毕，准备进入 Word 压缩引擎...")
                finally:
                    try: word.Quit()
                    except: pass
                    pythoncom.CoUninitialize()

            # ================== 引擎 A：PDF 压缩逻辑 ==================
            if ext == '.pdf':
                eel.update_terminal("[*] 智能路由：启动 PDF 压缩引擎 (Ghostscript)")
                
                gs = get_gs_path()
                if not os.path.exists(gs):
                    return {"status": "error", "msg": "找不到 Ghostscript 环境。请检查打包是否遗漏了 Ghostscript 目录。"}
                
                low, high = 72, 300
                closest_f = None
                min_diff = float('inf')
                last_size = 0
                mid = (low + high) // 2
                
                for i in range(5):
                    if not eel._websockets: 
                        return {"status": "error", "msg": "UI 已关闭，操作终止。"}

                    # ✨ 核心改进：加强预测权重，加速收敛
                    if last_size > 0:
                        ratio = target_b / last_size
                        smart_mid = int(mid * (ratio ** 0.5))
                        # 让智能预测值占据更大权重 (2份智能+1份下限+1份上限)
                        mid = (low + high + smart_mid * 2) // 4
                        mid = max(low, min(high, mid))
                    else:
                        mid = (low + high) // 2
                        
                    eel.update_terminal(f"[*] PDF 迭代 {i+1}/5: 智能调整参数 DPI={mid}...")
                    out_t = Path(tmp) / f"t_{mid}.pdf"
                    
                    cmd = [
                        gs,
                        "-sDEVICE=pdfwrite",
                        "-dCompatibilityLevel=1.4",
                        "-dPDFSETTINGS=/default",
                        "-dNOPAUSE", "-dQUIET", "-dBATCH",
                        "-dDetectDuplicateImages=true",
                        "-dCompressFonts=true",
                        "-dDownsampleColorImages=true",
                        f"-dColorImageResolution={mid}",
                        "-dColorImageDownsampleThreshold=1.0", # ✨ 破除默认 1.5 倍断崖限制，强制平滑压缩
                        "-dColorImageDownsampleType=/Bicubic",
                        "-dDownsampleGrayImages=true",
                        f"-dGrayImageResolution={mid}",
                        "-dGrayImageDownsampleThreshold=1.0", # ✨ 同上
                        "-dGrayImageDownsampleType=/Bicubic",
                        "-dDownsampleMonoImages=true",
                        f"-dMonoImageResolution={mid}",
                        "-dMonoImageDownsampleThreshold=1.0", # ✨ 同上
                        "-dMonoImageDownsampleType=/Bicubic",
                        f"-sOutputFile={out_t.as_posix()}",
                        in_p.as_posix()
                    ]
                    
                    result = subprocess.run(
                        cmd, 
                        creationflags=0x08000000,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )

                    if result.returncode != 0:
                        eel.update_terminal("[!] Ghostscript 运行异常，尝试降低参数...")
                    
                    if out_t.exists() and os.path.getsize(out_t) > 1024:
                        current_size = os.path.getsize(out_t)
                        last_size = current_size
                        
                        eel.update_terminal(f"  └─ 当前参数产出大小: {fmt_sz(current_size)} (目标大小: {fmt_sz(target_b)})")
                        
                        diff = abs(current_size - target_b)
                        if diff < min_diff:
                            min_diff = diff
                            closest_f = out_t
                            
                        if current_size <= target_b:
                            low = mid + 1
                        else:
                            high = mid - 1
                    else:
                        high = mid - 1
                
                if closest_f and closest_f.exists():
                    final_source = closest_f
                else:
                    return {"status": "error", "msg": "PDF 压缩失败：未能生成有效的 PDF 文件，请检查原文件是否加密或损坏。"}
                
                final = in_p.parent / f"{in_p.stem}_compressed.pdf"
                shutil.copy(final_source, final)
                return {"status": "success"}

            # ================== 引擎 B：Word (.docx) 压缩逻辑 ==================
            elif ext == '.docx':
                eel.update_terminal("[*] 智能路由：启动 Word 压缩引擎 (Pillow)")
                try:
                    from PIL import Image
                except ImportError:
                    return {"status": "error", "msg": "缺少图片处理库，请先在终端执行: pip install Pillow"}
                
                low_q, high_q = 10, 95
                closest_f = None
                min_diff = float('inf')
                last_size = 0
                mid_q = (low_q + high_q) // 2
                
                with zipfile.ZipFile(in_p, 'r') as zin:
                    item_list = zin.infolist()
                    item_data = {item.filename: zin.read(item.filename) for item in item_list}

                for i in range(5):
                    if not eel._websockets: 
                        return {"status": "error", "msg": "UI 已关闭，操作终止。"}

                    # ✨ 核心改进：加强预测权重
                    if last_size > 0:
                        ratio = target_b / last_size
                        smart_q = int(mid_q * ratio)
                        mid_q = (low_q + high_q + smart_q * 2) // 4
                        mid_q = max(low_q, min(high_q, mid_q))
                    else:
                        mid_q = (low_q + high_q) // 2

                    eel.update_terminal(f"[*] Word 迭代 {i+1}/5: 智能调整图片压缩率={mid_q}...")
                    out_t = Path(tmp) / f"t_{mid_q}.docx"
                    max_dim = int(2500 * (mid_q / 100)) + 400 
                    
                    with zipfile.ZipFile(out_t, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                        for item in item_list:
                            fname = item.filename
                            data = item_data[fname]
                            if fname.startswith('word/media/') and fname.lower().endswith(('.png', '.jpg', '.jpeg')):
                                try:
                                    img = Image.open(io.BytesIO(data))
                                    img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
                                    img_byte_arr = io.BytesIO()
                                    if fname.lower().endswith(('.jpg', '.jpeg')):
                                        if img.mode in ('RGBA', 'P'): img = img.convert('RGB')
                                        img.save(img_byte_arr, format='JPEG', quality=mid_q, optimize=True)
                                    elif fname.lower().endswith('.png'):
                                        img.save(img_byte_arr, format='PNG', optimize=True)
                                    zout.writestr(item, img_byte_arr.getvalue())
                                except Exception:
                                    zout.writestr(item, data)
                            else:
                                zout.writestr(item, data)

                    if out_t.exists():
                        current_size = os.path.getsize(out_t)
                        last_size = current_size
                        
                        eel.update_terminal(f"  └─ 当前参数产出大小: {fmt_sz(current_size)} (目标大小: {fmt_sz(target_b)})")

                        diff = abs(current_size - target_b)
                        if diff < min_diff:
                            min_diff = diff
                            closest_f = out_t
                            
                        if current_size <= target_b:
                            low_q = mid_q + 1
                        else: 
                            high_q = mid_q - 1

                final_dir = Path(path_str.strip()).resolve().parent
                final_name = f"{Path(path_str.strip()).resolve().stem}_compressed.docx"
                final = final_dir / final_name
                
                if closest_f and closest_f.exists():
                    shutil.copy(closest_f, final)
                    return {"status": "success"}
                else:
                    return {"status": "error", "msg": "Word 压缩失败。"}

            else:
                return {"status": "error", "msg": f"不支持的文件格式: {ext}"}
                
    except Exception as e: 
        return {"status": "error", "msg": str(e)}
