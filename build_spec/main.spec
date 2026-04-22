# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all
from PyInstaller.utils.hooks import copy_metadata

datas = []
binaries = [('C:\\Program Files\\Python311\\Lib\\site-packages\\pywin32_system32', '.')]
hiddenimports = ['webview', 'webview.platforms.winforms', 'webview.platforms.edgechromium', 'webview.platforms.mshtml', 'pythoncom', 'pywintypes', 'win32timezone', 'win32api', 'win32con', 'win32gui', 'win32com.client', 'comtypes', 'fitz', 'pandas', 'openpyxl', 'pdfplumber', 'docx', 'PIL', 'pdf2docx', 'img2pdf', 'pypdf', 'rapidocr_onnxruntime', 'cv2', 'numpy', 'ocrmypdf', 'pikepdf', 'pdfminer', 'pluggy', 'bridge', 'app_api', 'core_blank_page', 'core_compress', 'core_diff', 'core_img2pdf', 'core_invoice', 'core_ocr', 'core_pdf2img', 'core_pdf2word', 'core_pdf_cleaner', 'core_split', 'core_unlock', 'core_word2pdf', 'core_word_merge', 'core_word_split']
datas += collect_data_files('webview')
datas += copy_metadata('ocrmypdf')
datas += copy_metadata('pikepdf')
hiddenimports += collect_submodules('webview')
hiddenimports += collect_submodules('pdf2docx')
hiddenimports += collect_submodules('pdfminer')
tmp_ret = collect_all('ocrmypdf')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['D:\\git\\nmdgjh\\main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'qtpy', 'gi', 'cefpython3', 'webview.platforms.qt', 'webview.platforms.gtk', 'webview.platforms.cef', 'webview.platforms.cocoa', 'bottle_websocket', 'websocket'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['D:\\git\\nmdgjh\\toolbox_icon_clean.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
