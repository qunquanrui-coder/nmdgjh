# 泉泉的百宝箱

一个面向 Windows 办公场景的本地桌面文档工具箱，聚焦 PDF、Word、Excel、OCR、发票整理与批量文档处理。

项目采用 **Python + pywebview + 本地 Web 前端** 架构：前端负责交互，Python 后端负责调用 PDF、Office、OCR、图像处理等本地能力，适合在内网、离线或不方便上传文件的办公环境中使用。

## 主要功能

- PDF 提取 Word：支持可编辑模式和纯图固化模式
- Word / PDF 空白页清理：尽量保守处理，降低误删风险
- 扫描件去黑边：识别有效内容区域并覆盖边缘黑边
- PDF 精准拆分：支持定长拆分、平均拆分和范围提取
- Word 目录拆解：按大纲标题级别拆分成多个 Word 文档
- Word 批量合并：自然排序合并，自动插入标题和分页
- PDF 权限解密：移除打印、复制、编辑等限制
- 文档压缩：PDF 使用 pikepdf / Ghostscript，Word 压缩内嵌图片
- PDF OCR 增强：基于 OCRmyPDF 生成可搜索 PDF
- 图像转 PDF：支持 JPG、PNG、BMP，自然排序合并
- Word / Excel 转 PDF：调用 Office 原生导出能力
- PDF 转图片：使用 PyMuPDF 渲染输出 JPG
- 发票自动提取：提取发票信息并汇总为 Excel
- 文档差异比对：支持 Word 段落对比和 Excel 单元格对比

## 运行环境

- Windows 10 / 11 x64
- Python 3.10
- Microsoft Office（部分 Word / Excel 功能需要）
- Tesseract、Ghostscript、Poppler 等运行时资源

仓库中已包含打包所需的 `Ghostscript/`、`runtime/`、`poppler_bin/` 和 `web/` 目录，`build_modern.py` 会在打包时同步这些资源。

## 下载与启动

普通用户建议直接在 [Releases](https://github.com/qunquanrui-coder/nmdgjh/releases) 下载最新 Windows 压缩包，解压后运行 `QuanQuanTreasureBox.exe`。

源码运行时，先安装 Python 依赖：


```bash
pip install -r requirements.txt
```

启动桌面程序：

```bash
python main.py
```

程序会创建 pywebview 桌面窗口，并加载 `web/index.html` 作为主界面。

## 打包

项目提供 Windows 打包脚本：

```bash
python build_modern.py
```

打包脚本会清理旧构建产物，调用 PyInstaller，复制前端资源和运行时目录，并校验 `dist/QuanQuanTreasureBox/QuanQuanTreasureBox.exe` 及关键资源是否完整。

打包结果默认位于：

```text
dist/
└─ QuanQuanTreasureBox/
   ├─ QuanQuanTreasureBox.exe
   ├─ web/
   ├─ Ghostscript/
   ├─ runtime/
   └─ poppler_bin/
```

## 项目结构

```text
nmdgjh/
├─ .github/workflows/      # GitHub Actions 打包流程
├─ Ghostscript/            # Ghostscript 运行时
├─ poppler_bin/            # Poppler 运行时
├─ runtime/                # Tesseract 等运行时
├─ web/                    # 本地前端页面
├─ app_api.py              # pywebview API 路由
├─ bridge.py               # 前后端桥接辅助层
├─ main.py                 # 程序入口
├─ build_modern.py         # PyInstaller 打包脚本
├─ core_*.py               # 各业务功能模块
├─ requirements.txt        # Python 依赖
├─ CHANGELOG.md            # 更新日志
├─ LICENSE.MD              # 源码授权说明
├─ README_LICENSE.md       # 授权摘要
└─ EULA.md                 # 发布版程序使用条款
```

## 技术栈

- 桌面层：pywebview、tkinter
- PDF：PyMuPDF、pypdf、pikepdf、pdfplumber、pdf2docx、img2pdf
- Office：pywin32、comtypes、python-docx、openpyxl、pandas
- OCR / 图像：OCRmyPDF、Tesseract、RapidOCR、OpenCV、Pillow、NumPy
- 打包：PyInstaller、GitHub Actions

## 使用说明

本项目偏向本地办公生产力工具，核心目标是稳定处理常见文档任务。涉及重要文件时，建议先备份原文件再执行批量处理。

部分功能强依赖本机 Office COM 环境，不适合直接在 macOS / Linux 上运行。

## 授权说明

本仓库源代码未采用开源许可证授权，公开可见不代表开放复制、修改、分发、再许可、销售或派生开发权利。

作者官方发布的打包程序可按 `EULA.md` 免费下载安装和使用。源代码授权、商业使用、组织部署、再分发、白标 / OEM 或二次开发合作，请先获得作者书面许可。

第三方组件仍分别适用其各自许可证与使用条款。
