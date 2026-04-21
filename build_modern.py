# -*- coding: utf-8 -*-
"""
build_modern.py - Web architecture full hardening version (Eel ultimate packaging)
Core changes:
1. Entry point switched to main.py
2. Auto-inject web frontend folder (HTML/CSS/JS)
3. Force inject Eel底层 WebSocket dependencies, prevent disconnection
4. Supplement pdf2docx, img2pdf, rapidocr_onnxruntime, pypdf hidden imports
"""
import sys
import os
import shutil
import subprocess
from pathlib import Path
from typing import Optional, List

# Fix Windows console UTF-8 encoding issue
if sys.platform == 'win32':
    os.environ['PYTHONIOENCODING'] = 'utf-8'

class DependencyChecker:
    @staticmethod
    def find_pywin32_dll() -> Optional[Path]:
        """Fix pywin32_system32.__file__ being None issue"""
        try:
            import pywin32_system32
            f = getattr(pywin32_system32, "__file__", None)
            if f:
                return Path(f).parent
        except (ImportError, TypeError):
            pass
        
        py_site = Path(sys.executable).parent / "Lib" / "site-packages" / "pywin32_system32"
        return py_site if py_site.exists() else None


def build() -> None:
    entry: Path = Path("main.py")
    if not entry.exists():
        print("[ERROR] Cannot find main.py")
        return

    # Clean old builds
    for d in ["build", "dist", "__pycache__"]:
        target: Path = Path(d)
        if target.exists(): 
            try:
                shutil.rmtree(target)
                print(f"[CLEAN] Removed {d} directory")
            except OSError as e:
                print(f"[WARN] Failed to clean {d}: {e}")

    # Build PyInstaller command
    cmd: List[str] = [
        sys.executable, "-m", "PyInstaller",
        "--clean",
        "--noconfirm", 
        "--onedir",
        "--windowed",
        str(entry),
        # Core runtime libraries
        "--hidden-import=fitz",  
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=pdfplumber",
        "--hidden-import=docx",
        "--hidden-import=PIL",
        "--hidden-import=pythoncom",
        "--hidden-import=pywintypes",
        # PDF/Word conversion dependencies (supplement)
        "--hidden-import=pdf2docx",
        "--hidden-import=img2pdf",
        "--hidden-import=pypdf",
        "--hidden-import=rapidocr_onnxruntime",
        # Eel WebSocket dependencies
        "--hidden-import=bottle_websocket",
        "--hidden-import=bottle",
        "--hidden-import=websocket",
        # COM automation dependencies
        "--hidden-import=win32api",
        "--hidden-import=win32gui",
        "--hidden-import=win32con",
        "--hidden-import=comtypes",
        "--hidden-import=win32com.client",
        # Custom core modules
        "--hidden-import=core_diff",
        "--hidden-import=core_invoice",
        "--hidden-import=core_word_split", 
        "--hidden-import=core_compress",
        "--hidden-import=core_word2pdf",
        "--hidden-import=core_blank_page",  
        "--hidden-import=core_ocr",         
        "--hidden-import=core_pdf_cleaner",  # ✨ 新增：去黑边核心功能依赖
        "--hidden-import=cv2",               # ✨ 新增：OpenCV 图像处理底层依赖
        "--hidden-import=numpy",             # ✨ 新增：Numpy 矩阵计算依赖
        # ocrmypdf deep dependencies
        "--hidden-import=ocrmypdf",
        "--hidden-import=pikepdf",          
        "--hidden-import=pdfminer",         
        "--hidden-import=pluggy",           
        "--collect-all=ocrmypdf",           
        "--copy-metadata=ocrmypdf",         
        "--copy-metadata=pikepdf"           
    ]

    # Inject external resource folders
    web_dir: Path = Path("web")
    if web_dir.exists():
        cmd.extend(["--add-data", f"{web_dir};web"])
        print(f"[OK] Injected Web frontend folder: {web_dir}")
    else:
        print("[FATAL] Web folder not found!")

    dll_dir: Optional[Path] = DependencyChecker.find_pywin32_dll()
    if dll_dir:
        cmd.extend(["--add-binary", f"{dll_dir};."])
        print(f"[OK] Injected DLL path: {dll_dir}")

    icon: Path = Path("toolbox_icon_clean.ico")
    if icon.exists():
        cmd.extend(["--icon", str(icon), "--add-data", f"{icon};."])
        print(f"[OK] Added icon: {icon}")

    # Execute packaging
    print(f"\n[INFO] PyInstaller command preview:")
    print(" " + " ".join(cmd[:8]) + f" ... ({len(cmd)} args total)")
    print("\n[START] Deep packaging started, please wait...\n")
    
    try:
        subprocess.run(cmd, check=True)
        
        # Physical copy for external engines
        dist_main: Path = Path("dist/main")
        if dist_main.exists():
            print("\n[SYNC] PyInstaller compilation done. Executing physical copy for external engines...")
            
            for folder_name in ["Ghostscript", "runtime", "poppler_bin"]:
                src_folder: Path = Path(folder_name)
                dst_folder: Path = dist_main / folder_name
                
                if src_folder.exists():
                    if dst_folder.exists():
                        shutil.rmtree(dst_folder)
                    
                    shutil.copytree(src_folder, dst_folder)
                    print(f"[COPY] Successfully copied: {src_folder} -> {dst_folder}")

        print("\n[SUCCESS] Packaging complete! Test main.exe in dist/main directory.")
        print("[NOTE] For distribution, compress the entire [main] folder.")
        
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] Packaging failed with code: {e.returncode}")
        if e.stderr:
            print(f"[DETAILS]:\n{e.stderr}")
    except FileNotFoundError:
        print("\n[ERROR] pyinstaller not found. Install via: pip install pyinstaller")
    except Exception as e:
        print(f"\n[ERROR] Unexpected error: {e}")


if __name__ == "__main__":
    build()
