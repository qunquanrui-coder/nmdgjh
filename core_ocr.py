# -*- coding: utf-8 -*-

import logging
import multiprocessing
import os
import re
import subprocess
import sys
import threading
import time
import uuid
import warnings
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, List


try:
    import eel
except ImportError:
    class _DummyEel:
        def expose(self, func):
            return func

        def update_terminal(self, msg: str):
            return None

    eel = _DummyEel()


def push_heartbeat_log(msg: str) -> None:
    try:
        if getattr(eel, "update_terminal", None):
            eel.update_terminal(msg)
    except Exception:
        pass


# ---------------------------------------------------------
# 运行环境准备
# ---------------------------------------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

TESS_DIR = BASE_DIR / "runtime" / "Tesseract"
GS_BIN_DIR = BASE_DIR / "Ghostscript" / "bin"


def _prepend_env_path(path_obj: Path) -> None:
    path_str = str(path_obj)
    current = os.environ.get("PATH", "")
    parts = current.split(os.pathsep) if current else []
    if path_str and path_str not in parts:
        os.environ["PATH"] = path_str + (os.pathsep + current if current else "")


_prepend_env_path(TESS_DIR)
_prepend_env_path(GS_BIN_DIR)
os.environ["TESSDATA_PREFIX"] = str(TESS_DIR / "tessdata")

import ocrmypdf

logging.getLogger("ocrmypdf").setLevel(logging.ERROR)
warnings.filterwarnings("ignore")


# ---------------------------------------------------------
# 仅在 OCR 执行期间临时隐藏子进程窗口
# ---------------------------------------------------------
@contextmanager
def hidden_subprocess_windows():
    if sys.platform != "win32":
        yield
        return

    original_popen = subprocess.Popen

    class SilentPopen(original_popen):
        def __init__(self, *args: Any, **kwargs: Any) -> None:
            startupinfo = kwargs.get("startupinfo")
            if startupinfo is None:
                startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            kwargs["startupinfo"] = startupinfo
            kwargs["creationflags"] = kwargs.get("creationflags", 0) | 0x08000000
            super().__init__(*args, **kwargs)

    subprocess.Popen = SilentPopen
    try:
        yield
    finally:
        subprocess.Popen = original_popen


# ---------------------------------------------------------
# 进度流解析
# ---------------------------------------------------------
class OCRProgressStream:
    def __init__(self, file_name: str) -> None:
        self.file_name = file_name
        self.pattern = re.compile(r"(\d+)/(\d+)")
        self.last_current = ""
        self.locked_total = ""
        self.original_stderr = sys.stderr
        self.last_update_time = time.time()

    def write(self, text: str) -> None:
        if self.original_stderr is not None:
            self.original_stderr.write(text)

        self.last_update_time = time.time()

        try:
            matches = self.pattern.findall(text)
            if not matches:
                return

            current, total = matches[-1]
            if total == "0":
                return

            if not self.locked_total:
                self.locked_total = total

            if total == self.locked_total and current != self.last_current:
                self.last_current = current
                push_heartbeat_log(
                    f"⏳ [{self.file_name}] OCR 扫描中: 第 {current} 页 / 共 {total} 页"
                )
        except Exception:
            pass

    def flush(self) -> None:
        if self.original_stderr is not None:
            self.original_stderr.flush()

    def isatty(self) -> bool:
        return False


def _heartbeat_worker(file_name: str, stream: OCRProgressStream, stop_event: threading.Event) -> None:
    start_t = time.time()
    while not stop_event.is_set():
        time.sleep(3)
        if stop_event.is_set():
            break

        if time.time() - stream.last_update_time > 8:
            elapsed = int(time.time() - start_t)
            push_heartbeat_log(
                f"⏳ [{file_name}] 页面排版较复杂，底层正全力转码中... (已耗时 {elapsed}s)"
            )
            stream.last_update_time = time.time()


def _atomic_replace(tmp_path: Path, final_path: Path) -> bool:
    for _ in range(5):
        try:
            if final_path.exists():
                final_path.unlink()
            tmp_path.replace(final_path)
            return True
        except PermissionError:
            time.sleep(0.5)
        except OSError:
            time.sleep(0.5)
    return False


def _collect_pdf_files(target: Path) -> List[Path]:
    if target.is_file():
        if target.suffix.lower() == ".pdf":
            return [target]
        return []

    if target.is_dir():
        return sorted(
            [
                f for f in target.iterdir()
                if f.is_file() and f.suffix.lower() == ".pdf" and not f.name.startswith("~$")
            ],
            key=lambda p: p.name.lower(),
        )

    return []


def _run_single_ocr(path: Path, safe_threads: int) -> Dict[str, Any]:
    output_path = path.parent / f"{path.stem}_可搜索.pdf"
    tmp_output_path = output_path.with_name(
        f"{output_path.stem}__tmp__{uuid.uuid4().hex[:8]}.pdf"
    )

    original_stderr = sys.stderr
    progress_stream = OCRProgressStream(path.name)
    stop_event = threading.Event()
    hb_thread = threading.Thread(
        target=_heartbeat_worker,
        args=(path.name, progress_stream, stop_event),
        daemon=True,
    )

    keep_tmp = False

    try:
        push_heartbeat_log(
            f"▶ 启动引擎: {path.name} (分配 {safe_threads} 个线程以保障系统平稳)"
        )

        sys.stderr = progress_stream
        hb_thread.start()

        with hidden_subprocess_windows():
            ocrmypdf.ocr(
                str(path),
                str(tmp_output_path),
                language=["chi_sim", "eng"],
                force_ocr=True,
                output_type="pdf",
                threads=safe_threads,
                optimize=1,
                fast_web_view=999,
                skip_big=15,
                progress_bar=True,
            )

        sys.stderr = original_stderr

        if not tmp_output_path.exists():
            push_heartbeat_log(f"[!] [{path.name}] OCR 完成但未生成输出文件")
            return {"status": "error", "msg": f"{path.name} 未生成输出文件", "data": None}

        if _atomic_replace(tmp_output_path, output_path):
            push_heartbeat_log(f"[√] [{path.name}] OCR 完成 -> {output_path.name}")
            return {"status": "success", "msg": "", "data": str(output_path)}

        keep_tmp = True
        push_heartbeat_log(f"[!] [{path.name}] 输出文件被占用，结果已保留为: {tmp_output_path.name}")
        return {"status": "success", "msg": "", "data": str(tmp_output_path)}

    except Exception as e:
        sys.stderr = original_stderr
        push_heartbeat_log(f"❌ [{path.name}] OCR 处理失败: {str(e)}")
        return {"status": "error", "msg": f"OCR 处理失败: {str(e)}", "data": None}

    finally:
        stop_event.set()
        try:
            hb_thread.join(timeout=1.0)
        except Exception:
            pass

        sys.stderr = original_stderr

        if tmp_output_path.exists() and not keep_tmp:
            try:
                tmp_output_path.unlink()
            except Exception:
                pass


@eel.expose
def run_ocr(target_path: str) -> Dict[str, Any]:
    try:
        target = Path(target_path.strip())
        file_paths = _collect_pdf_files(target)

        if not file_paths:
            return {
                "status": "error",
                "msg": "选定路径中未找到有效的 PDF 文件",
                "data": None,
            }

        hardware_threads = multiprocessing.cpu_count()
        safe_threads = max(1, hardware_threads - 2)

        for path in file_paths:
            result = _run_single_ocr(path, safe_threads)
            if result.get("status") != "success":
                return result

        return {
            "status": "success",
            "msg": f"共处理 {len(file_paths)} 个文件",
            "data": None,
        }

    except Exception as e:
        import traceback

        push_heartbeat_log(f"❌ [致命] OCR 引擎异常: {str(e)}")
        push_heartbeat_log(traceback.format_exc())
        return {"status": "error", "msg": str(e), "data": None}
