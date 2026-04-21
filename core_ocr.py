# -*- coding: utf-8 -*-
import os
import sys
import logging
import warnings
import subprocess
import multiprocessing
import re
import time        # ✨ 新增：用于心跳时间计算
import threading   # ✨ 新增：用于独立的心跳监控线程
from pathlib import Path
from typing import Dict, Any, List

# ---------------------------------------------------------
# 动态引入 Eel 用于发送心跳日志到前端 UI (完全解耦设计)
# ---------------------------------------------------------
try:
    import eel
    def push_heartbeat_log(msg: str) -> None:
        """
        尝试向 Web 端发送实时日志，若无环境则静默
        
        Args:
            msg (str): 需要发送到前端界面的日志内容
        """
        try:
            if getattr(eel, 'update_terminal', None):
                eel.update_terminal(msg)
        except Exception:
            pass
except ImportError:
    def push_heartbeat_log(msg: str) -> None:
        pass

# ---------------------------------------------------------
# 核心黑科技：类继承式 Popen 拦截 (彻底解决黑窗口与 asyncio 冲突)
# ---------------------------------------------------------
if sys.platform == 'win32':
    _OriginalPopen = subprocess.Popen

    class SilentPopen(_OriginalPopen):
        """继承自原始 Popen，强制隐藏所有子进程窗口，且兼容 asyncio 继承检查"""
        def __init__(self, *args: Any, **kwargs: Any) -> None:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            si.wShowWindow = subprocess.SW_HIDE 
            
            kwargs['startupinfo'] = si
            # CREATE_NO_WINDOW: 0x08000000
            kwargs['creationflags'] = kwargs.get('creationflags', 0) | 0x08000000
            
            super().__init__(*args, **kwargs)

    # 替换标准库引用
    subprocess.Popen = SilentPopen


# ---------------------------------------------------------
# 动态环境关联与注入 (必须位于 import ocrmypdf 之前)
# ---------------------------------------------------------
if getattr(sys, 'frozen', False):
    BASE_DIR: Path = Path(sys.executable).parent
else:
    BASE_DIR: Path = Path(__file__).parent

# 依赖路径：严格适配物理目录，全程使用 pathlib 确保跨平台兼容
TESS_DIR: Path = BASE_DIR / "runtime" / "Tesseract"
GS_BIN_DIR: Path = BASE_DIR / "Ghostscript" / "bin" 

# 动态注入全局环境变量，确保后续加载的包能读取到正确的路径
os.environ["PATH"] = f"{str(TESS_DIR)}{os.pathsep}{str(GS_BIN_DIR)}{os.pathsep}{os.environ.get('PATH', '')}"
os.environ["TESSDATA_PREFIX"] = str(TESS_DIR / "tessdata")


# ---------------------------------------------------------
# 核心包导入 (底层环境就绪后方可安全加载)
# ---------------------------------------------------------
import ocrmypdf

# 屏蔽无意义的系统警告日志，保持控制台输出的专业与干净
logging.getLogger("ocrmypdf").setLevel(logging.ERROR)
warnings.filterwarnings("ignore")


# ---------------------------------------------------------
# 进度条底层拦截引擎 (已修复跳页与串流问题)
# ---------------------------------------------------------
class OCRProgressStream:
    """
    拦截 sys.stderr 输出，解析 ocrmypdf 内部的页码进度数据
    并推送到 Web 前端心跳日志中。
    """
    def __init__(self, file_name: str) -> None:
        self.file_name: str = file_name
        self.pattern: re.Pattern = re.compile(r'(\d+)/(\d+)')
        self.last_current: str = ""
        self.locked_total: str = ""
        self.original_stderr = sys.stderr
        self.last_update_time = time.time()  # ✨ 记录最后一次活跃时间，用于防假死监控

    def write(self, text: str) -> None:
        if self.original_stderr is not None:
            self.original_stderr.write(text)
            
        # ✨ 只要底层有任何输出流（哪怕是内部的极小进度片段），就刷新活跃时间
        self.last_update_time = time.time() 

        try:
            # 提取所有形如 "1/69" 的进度片段
            matches = self.pattern.findall(text)
            if matches:
                current, total = matches[-1]
                if total == "0":
                    return
                    
                if not self.locked_total:
                    self.locked_total = total
                    
                if total == self.locked_total and current != self.last_current:
                    self.last_current = current
                    push_heartbeat_log(f"⏳ [{self.file_name}] OCR 扫描中: 第 {current} 页 / 共 {total} 页")
        except Exception:
            pass
        
    def flush(self) -> None:
        if self.original_stderr is not None:
            self.original_stderr.flush()


@eel.expose
def run_ocr(target_path: str) -> Dict[str, Any]:
    """执行 OCR 双层 PDF 任务"""
    try:
        target: Path = Path(target_path)
        file_paths: List[Path] = []
        
        try:
            if target.is_file() and target.suffix.lower() == '.pdf':
                file_paths.append(target)
            elif target.is_dir():
                file_paths = [f for f in target.iterdir() if f.suffix.lower() == '.pdf']
        except OSError as oe:
            return {"status": "error", "msg": f"文件目录扫描失败: {str(oe)}", "data": None}
            
        if not file_paths:
            return {"status": "error", "msg": "选定路径中未找到有效的 PDF 文件", "data": None}

        hardware_threads: int = multiprocessing.cpu_count()
        safe_threads: int = max(1, hardware_threads - 2)

        for path in file_paths:
            output_path: Path = path.parent / f"{path.stem}_可搜索.pdf"
            
            push_heartbeat_log(f"▶ 启动引擎: {path.name} (分配 {safe_threads} 个线程以保障系统平稳)")
            
            original_stderr = sys.stderr
            progress_stream = OCRProgressStream(path.name)
            sys.stderr = progress_stream
            
            # ========================================================
            # ✨ 新增：防假死焦虑的独立心跳监控线程
            # ========================================================
            is_running = [True]  # 使用列表确保在闭包内安全传递与修改
            def heartbeat(stream):
                start_t = time.time()
                while is_running[0]:
                    time.sleep(3)
                    # 如果底层引擎超过 8 秒没有输出进度，发送心跳安抚用户
                    if is_running[0] and (time.time() - stream.last_update_time > 8):
                        elapsed = int(time.time() - start_t)
                        push_heartbeat_log(f"⏳ [{path.name}] 页面排版较复杂，底层正全力转码中... (已耗时 {elapsed}s)")
                        stream.last_update_time = time.time()  # 重置时间，防止刷屏
            
            hb_thread = threading.Thread(target=heartbeat, args=(progress_stream,), daemon=True)
            hb_thread.start()
            # ========================================================
            
            try:
                ocrmypdf.ocr(
                    str(path), 
                    str(output_path), 
                    language=["chi_sim", "eng"], 
                    force_ocr=True,
                    output_type="pdf",
                    threads=safe_threads,
                    optimize=1,       
                    fast_web_view=999, 
                    skip_big=15,      
                    progress_bar=True  # ✨ 核心修复：开启引擎原生进度条，重新激活日志流拦截器
                )
            except Exception as e:
                sys.stderr = original_stderr
                push_heartbeat_log(f"❌ [{path.name}] OCR 处理失败: {str(e)}")
                return {"status": "error", "msg": f"OCR 处理失败: {str(e)}", "data": None}
            finally:
                is_running[0] = False  # ✨ 任务结束或异常时，彻底关闭心跳监控线程
                sys.stderr = original_stderr
            
            if output_path.exists():
                push_heartbeat_log(f"[√] [{path.name}] OCR 完成 -> {output_path.name}")
            else:
                push_heartbeat_log(f"[!] [{path.name}] OCR 完成但未生成输出文件")

        return {"status": "success", "msg": f"共处理 {len(file_paths)} 个文件"}
    except Exception as e:
        import traceback
        push_heartbeat_log(f"❌ [致命] OCR 引擎异常: {str(e)}")
        push_heartbeat_log(traceback.format_exc())
        return {"status": "error", "msg": str(e), "data": None}
