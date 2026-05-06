# -*- coding: utf-8 -*-

import logging
import multiprocessing
import os
import re
import subprocess
import sys
import shutil
import threading
import time
import uuid
import warnings
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, List


try:
    import bridge
except ImportError:
    class _DummyBridge:
        def expose(self, func):
            return func

        def update_terminal(self, msg: str):
            return None

    bridge = _DummyBridge()


def push_heartbeat_log(msg: str) -> None:
    try:
        if getattr(bridge, "update_terminal", None):
            bridge.update_terminal(msg)
    except Exception:
        pass


# ---------------------------------------------------------
# 运行环境准备
# ---------------------------------------------------------
def _get_base_dirs() -> List[Path]:
    """兼容源码运行、PyInstaller onedir、PyInstaller onefile。"""
    dirs: List[Path] = []

    if getattr(sys, "frozen", False):
        dirs.append(Path(sys.executable).parent)
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            dirs.append(Path(meipass))
    else:
        dirs.append(Path(__file__).parent)

    dirs.append(Path.cwd())

    unique: List[Path] = []
    seen = set()
    for d in dirs:
        try:
            key = str(d.resolve())
        except Exception:
            key = str(d)
        if key not in seen:
            seen.add(key)
            unique.append(d)
    return unique


BASE_DIRS = _get_base_dirs()
BASE_DIR = BASE_DIRS[0]


def _prepend_env_path(path_obj: Path) -> None:
    if not path_obj or not path_obj.exists():
        return

    path_str = str(path_obj)
    current = os.environ.get("PATH", "")
    parts = current.split(os.pathsep) if current else []
    if path_str and path_str not in parts:
        os.environ["PATH"] = path_str + (os.pathsep + current if current else "")


def _find_existing_dir(candidates: List[Path]) -> Path:
    for p in candidates:
        if p.exists() and p.is_dir():
            return p
    return candidates[0]


def _find_tesseract_dir() -> Path:
    candidates: List[Path] = []
    for base in BASE_DIRS:
        candidates.extend([
            base / "runtime" / "Tesseract",
            base / "runtime" / "Tesseract" / "bin",
            base / "Tesseract",
            base / "Tesseract" / "bin",
        ])

    exe_names = ["tesseract.exe", "tesseract"]
    for p in candidates:
        if p.exists() and p.is_dir():
            if any((p / exe).exists() for exe in exe_names):
                return p

    return candidates[0]


def _find_tessdata_dir(tess_dir: Path) -> Path:
    candidates: List[Path] = []
    for base in BASE_DIRS:
        candidates.extend([
            base / "runtime" / "Tesseract" / "tessdata",
            base / "runtime" / "Tesseract" / "share" / "tessdata",
            base / "runtime" / "tessdata",
            base / "Tesseract" / "tessdata",
            base / "Tesseract" / "share" / "tessdata",
        ])

    candidates.extend([
        tess_dir / "tessdata",
        tess_dir.parent / "tessdata",
        tess_dir / "share" / "tessdata",
        tess_dir.parent / "share" / "tessdata",
    ])

    for p in candidates:
        if p.exists() and p.is_dir():
            if (p / "chi_sim.traineddata").exists() or (p / "eng.traineddata").exists():
                return p

    # 兜底：在 runtime/Tesseract 下递归找 tessdata，避免打包目录层级变化导致失效。
    for base in BASE_DIRS:
        for root in [base / "runtime" / "Tesseract", base / "runtime", base / "Tesseract"]:
            if not root.exists() or not root.is_dir():
                continue
            try:
                for p in root.rglob("tessdata"):
                    if p.is_dir() and (
                        (p / "chi_sim.traineddata").exists() or (p / "eng.traineddata").exists()
                    ):
                        return p
            except Exception:
                pass

    return candidates[0]


def _is_ascii_path(path_obj: Path) -> bool:
    try:
        str(path_obj).encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def _public_ascii_tessdata_cache() -> Path:
    # 使用固定英文路径，避免 Tesseract/MinGW 在中文目录下误报 tessdata 不存在。
    return Path(r"C:\Users\Public\QuanQuanTreasureBox\tessdata")


def _copy_tessdata_tree(src: Path, dst: Path) -> bool:
    try:
        if not src.exists() or not src.is_dir():
            return False
        dst.mkdir(parents=True, exist_ok=True)

        # 复制整个 tessdata 目录，包括 traineddata、configs、tessconfigs。
        for item in src.iterdir():
            target = dst / item.name
            if item.is_dir():
                shutil.copytree(item, target, dirs_exist_ok=True)
            elif item.is_file():
                # 只在源文件更新或目标不存在时复制，避免每次启动都重写 70MB+ 语言包。
                if (not target.exists()) or item.stat().st_mtime > target.stat().st_mtime or item.stat().st_size != target.stat().st_size:
                    shutil.copy2(item, target)
        return (dst / "chi_sim.traineddata").exists() and (dst / "eng.traineddata").exists()
    except Exception:
        return False


def _prepare_tessdata_for_tesseract(src: Path) -> Path:
    """
    Windows 版 Tesseract/MinGW 在某些环境下无法可靠读取中文路径，
    会出现“目录实际存在但 TESSDATA_PREFIX 报不存在”的问题。
    这里把 tessdata 缓存到固定英文目录，再把 TESSDATA_PREFIX 指向该目录。
    """
    if sys.platform == "win32" and src.exists() and src.is_dir():
        cache = _public_ascii_tessdata_cache()
        if _copy_tessdata_tree(src, cache):
            return cache
    return src


TESS_DIR = _find_tesseract_dir()
TESSDATA_SOURCE_DIR = _find_tessdata_dir(TESS_DIR)
TESSDATA_DIR = _prepare_tessdata_for_tesseract(TESSDATA_SOURCE_DIR)
GS_BIN_DIR = _find_existing_dir([base / "Ghostscript" / "bin" for base in BASE_DIRS])

_prepend_env_path(TESS_DIR)
_prepend_env_path(TESS_DIR / "bin")
_prepend_env_path(GS_BIN_DIR)

# 只有找到真实 tessdata 目录时才设置，避免把不存在的路径写入环境变量导致 Tesseract 直接失败。
if TESSDATA_DIR.exists() and TESSDATA_DIR.is_dir():
    os.environ["TESSDATA_PREFIX"] = str(TESSDATA_DIR)
else:
    os.environ.pop("TESSDATA_PREFIX", None)

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
        if not TESSDATA_DIR.exists() or not TESSDATA_DIR.is_dir():
            msg = (
                "OCR 运行库缺少 tessdata 语言包目录。请确认打包目录中存在 "
                "runtime\\Tesseract\\tessdata，并且里面至少包含 chi_sim.traineddata 和 eng.traineddata。"
            )
            push_heartbeat_log(f"❌ [{path.name}] {msg}")
            return {"status": "error", "msg": msg, "data": None}

        if not (TESSDATA_DIR / "chi_sim.traineddata").exists():
            msg = f"OCR 运行库缺少中文语言包: {TESSDATA_DIR / 'chi_sim.traineddata'}"
            push_heartbeat_log(f"❌ [{path.name}] {msg}")
            return {"status": "error", "msg": msg, "data": None}

        if not (TESSDATA_DIR / "eng.traineddata").exists():
            msg = f"OCR 运行库缺少英文语言包: {TESSDATA_DIR / 'eng.traineddata'}"
            push_heartbeat_log(f"❌ [{path.name}] {msg}")
            return {"status": "error", "msg": msg, "data": None}

        push_heartbeat_log(
            f"▶ 启动引擎: {path.name} (分配 {safe_threads} 个线程以保障系统平稳)"
        )
        push_heartbeat_log(f"[*] Tesseract: {TESS_DIR}")
        push_heartbeat_log(f"[*] Tessdata source: {TESSDATA_SOURCE_DIR}")
        push_heartbeat_log(f"[*] Tessdata runtime: {TESSDATA_DIR}")

        sys.stderr = progress_stream
        hb_thread.start()

        with hidden_subprocess_windows():
            ocrmypdf.ocr(
                str(path),
                str(tmp_output_path),
                language=["chi_sim", "eng"],
                force_ocr=True,
                output_type="pdf",
                jobs=safe_threads,
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


@bridge.expose
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
