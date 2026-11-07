# 泉泉的百宝箱

一个面向 Windows 办公场景的本地桌面文档工具箱，聚焦 PDF、Word、Excel、OCR、发票整理与批量文档处理。

项目采用 **Python + pywebview + 本地 Web 前端** 架构：前端负责交互，Python 后端负责调用 PDF、Office、OCR、图像处理等本地能力，适合在内网、离线或不方便上传文件的办公环境中使用。

## 主要功能

- PDF 提取 Word：支持可编辑模式和纯图固化模式，支持单文件与文件夹批量处理
- Word / PDF 空白页清理：尽量保守处理，降低误删风险，修改前自动创建 `.blankbak` 备份
- 扫描件去黑边：识别有效内容区域并覆盖边缘黑边
- PDF 换页：指定页码后用图片、图片文件夹或 PDF 页面替换，支持文件名识别页码、文件夹批量处理和单模板多替换文件批量生成
- PDF 精准拆分：支持定长拆分、平均拆分和指定页码范围提取
- Word 目录拆解：按大纲标题级别拆分成多个 Word 文档，每次拆解独立建目录保护历史结果
- Word 批量合并：自然排序合并，自动插入标题和分页
- PDF 权限解密：移除打印、复制、编辑等限制，支持单文件与文件夹批量清洗
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

程序会创建 pywebview 桌面窗口，并加载 `web/index.html` 作为主界面。运行日志会自动写入 `logs/app.log`，未捕获崩溃异常会追加记录至 `crash_log.txt`。

## 开发与测试

项目包含完整的回归单元测试套件（位于 `tests/` 目录），覆盖输入路径防护、输出文件防覆盖保护、PDF 导出/拆分/换页/解密安全、发票提取与文档比对等核心场景。

在项目根目录下运行测试：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

### 架构与开发约定

1. **前后端桥接 (`bridge.py` & `app_api.py`)**：
   - 各核心模块通过 `@bridge.expose` 装饰器注册可被前端调用的函数。
   - 前端通过 `invoke(funcName, ...args)` 统一经由 `AppApi.invoke` 分发，除 `ask_file` / `ask_folder` 外均统一返回 `{"status": "success|error|partial", "msg": "...", "data": ...}`。
   - 涉及 `win32com`（Word / Excel）的任务在 `app_api.py` 的 `COM_TASKS` 中声明，自动通过全局互斥锁 `_com_lock` 串行化并管理 `CoInitialize` / `CoUninitialize`。
2. **输出文件安全 (`safe_output.py`)**：
   - 所有导出、转换、汇总、换页和拆分操作，均应先写入临时目录/临时文件，再通过 `publish_output`、`new_output_file` 或 `new_output_directory` 发布，严禁直接覆盖用户已有同名文件或留下残损半成品文件。
3. **日志与状态反馈**：
   - 核心模块调用 `bridge.update_terminal(msg)` 时，消息会同时推送至前端实时终端面板并持久化到 `logs/app.log`。
   - 耗时超过数秒的任务应启动后台 daemon 心跳线程，定期输出进度与耗时。

## 打包

项目提供 Windows 打包脚本：

```bash
python build_modern.py
```

打包脚本会清理旧构建产物，调用 PyInstaller，复制前端资源和运行时目录，并校验 `dist/QuanQuanTreasureBox/QuanQuanTreasureBox.exe` 及关键资源是否完整（若已存在上一版构建，会自动备份为 `QuanQuanTreasureBox_previous`）。

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
├─ tests/                  # 自动化回归测试套件
├─ web/                    # 本地前端页面
├─ app_api.py              # pywebview API 路由与 COM 线程守卫
├─ bridge.py               # 前后端桥接与日志同步辅助层
├─ safe_output.py          # 非覆盖式安全文件/目录输出辅助模块
├─ main.py                 # 程序入口与全局日志配置
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
