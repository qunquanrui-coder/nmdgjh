# 泉泉大人的工具合集（Baibaoxiang Web）

一个面向 Windows 办公场景的本地桌面文档工具箱，聚焦 PDF、Word、Excel、OCR 与发票处理等高频需求。

项目采用 **Python + pywebview + 本地 Web 前端** 架构，围绕“能直接用、能批量跑、能打包发布”的目标，提供 PDF 转 Word、Word / Excel 转 PDF、扫描件 OCR、文档压缩、差异比对、发票提取、图片转 PDF 等一整套本地化处理能力。

当前仓库已经从早期浏览器壳模式迁移到 **pywebview 桌面壳**，并使用项目内置的 **bridge** 前后端桥接层，适合继续迭代为稳定的 Windows 单文件 / 目录版工具。

---

## 授权与使用说明

本项目仓库源码**未采用开源许可证授权**，默认受版权法保护。

- **关于源代码**：未经作者事先书面许可，任何个人或组织不得对本仓库中的源代码及相关内容进行复制、修改、分发、再许可、销售或用于派生开发。
- **关于发布版程序**：作者发布的安装包、可执行程序等打包版本，可供个人和组织免费下载安装和使用。
- **禁止事项**：未经书面许可，不得对发布版程序进行二次分发、转售、改包发布、白标发布，或以再授权形式向第三方提供。
- **第三方组件**：本项目所使用的第三方库、运行时和工具组件，仍分别适用其各自许可证与使用条款。

详细条款请参见仓库中的 **`LICENSE`** 与 **`EULA.md`**。

---

## 项目定位

这个项目是一套偏“办公生产力”的本地桌面工具集合，特点是：

- **以本地处理为主**：适合内网、离线、涉密或不便上传云端的文档场景。
- **以批量处理为主**：大多数模块支持单文件与目录处理。
- **以 Windows 办公生态为主**：深度接入 Word / Excel COM 导出能力。
- **以可打包发布为主**：仓库已包含打包脚本和运行时资源同步逻辑。

---

## 功能一览

### 1）文档处理

- **PDF 提取 Word**
  - 支持可编辑模式与纯图固化模式
  - 可对单个 PDF 或整个文件夹递归处理
- **Word / PDF 空白页清理**
  - Word 通过 COM 分页与页范围删除处理
  - PDF 通过文本、图像、矢量与注释联合判空
- **扫描件去黑边**
  - 基于 OpenCV 识别有效内容边界并对白边区域进行覆盖
- **PDF 精准拆分**
  - 支持定长拆分 / 平均拆分 / 范围提取
  - 含多级“强攻”打开策略，兼容部分异常 PDF
- **Word 目录拆解**
  - 自动扫描大纲级别
  - 按目标标题级别拆成多个独立 Word 文档
- **Word 批量合并**
  - 自然排序合并多个 Word 文件
  - 自动插入标题与分页，并带剪贴板重试机制
- **PDF 权限解密 / 限制移除**
  - 支持处理打印、复制、编辑等限制
  - 提供快速模式与安全模式，且支持失败自动重试
- **文档极限瘦身**
  - PDF：先无损瘦身，再 Ghostscript 有损压缩
  - Word：对 `docx` 内嵌图片进行重采样与重压缩
  - 采用“目标区间优先”策略，不盲目压得过小

### 2）转换与 OCR

- **PDF OCR 增强**
  - 基于 OCRmyPDF 生成可搜索双层 PDF
  - 支持中文简体 + 英文
  - 带进度解析与心跳日志
- **图像转 PDF**
  - 支持 JPG / PNG / BMP
  - 采用自然排序
  - 透明 PNG 会自动补白后再写入 PDF
- **Word / Excel 转 PDF**
  - 调用 Office 原生 `ExportAsFixedFormat`
  - Word 支持书签导出与 PDF/A 选项
  - Excel 遵循打印区域导出
- **PDF 转图片**
  - 使用 PyMuPDF 渲染
  - 输出 JPG，保留精确 DPI 信息
- **发票自动提取**
  - 支持 PDF / JPG / PNG / JPEG / BMP
  - 文本优先提取，必要时回退 OCR
  - 自动汇总为 Excel，并生成错误清单
- **文档差异比对**
  - Word：按段落做深度对比
  - Excel：按工作表逐单元格对比
  - 输出格式化 Excel 报告

---

## 当前前端模块

前端首页当前已集成以下功能入口：

- PDF 提取 Word
- Word 空白页清理
- 扫描件去黑边
- PDF 精准拆分
- Word 目录拆解
- Word 批量合并
- PDF 权限解密
- 文档极限瘦身
- PDF OCR 增强
- 图像转 PDF
- Word 转 PDF
- PDF 转图片
- 发票自动提取
- 文档差异比对

界面风格为本地 Web 面板式布局，配套：

- Toast 通知
- 终端日志输出
- 左侧导航切换
- 文件 / 文件夹选择器联动
- 长任务执行中的状态反馈

---

## 技术架构

### 桌面壳

- **pywebview**
  - 当前项目已经切换到 pywebview 作为桌面承载层
  - 本地 HTML 页面通过桌面窗口加载
  - 启动时显式启用本地 HTTP Server 以确保静态资源稳定访问

### 前后端桥接

- **自定义 bridge 接口暴露 + pywebview 桥接层**
  - 业务模块使用 `@bridge.expose` 暴露方法
  - `app_api.py` 负责把前端调用路由到 Python 端
  - 对 COM 类任务增加串行锁，避免 Word / Excel 并发冲突

### 前端

- HTML + CSS + JavaScript
- 页面文件位于 `web/` 目录：
  - `index.html`
  - `script.js`
  - `style.css`

### 后端核心

- Python 3.10
- 按功能拆分为多个 `core_*.py` 模块，入口统一由 `main.py` 加载

---

## 技术栈

### 通用与桌面层

- Python 3.10
- pywebview
- bridge（项目内置桥接层）
- tkinter（文件 / 文件夹选择）
- PyInstaller（打包）

### PDF 与 Office 处理

- PyMuPDF
- pypdf
- pikepdf
- pdfplumber
- pdf2docx
- img2pdf
- python-docx
- openpyxl
- pandas
- pywin32 / comtypes

### OCR 与图像处理

- OCRmyPDF
- Tesseract
- RapidOCR
- OpenCV
- Pillow
- NumPy

### 打包运行时资源

仓库中已包含并在打包脚本中同步的运行时目录：

- `Ghostscript/`
- `runtime/`（含 Tesseract）
- `poppler_bin/`
- `web/`

---

## 目录结构

```text
nmdgjh/
├─ .github/workflows/          # GitHub Actions / 打包流程
├─ Ghostscript/                # Ghostscript 运行时
├─ poppler_bin/                # Poppler 运行时
├─ runtime/
│  └─ Tesseract/               # OCR 运行时
├─ web/                        # 本地前端页面
│  ├─ index.html
│  ├─ script.js
│  └─ style.css
├─ app_api.py                  # pywebview 与 bridge 接口路由
├─ build_modern.py             # PyInstaller 打包脚本
├─ core_blank_page.py          # Word/PDF 空白页清理
├─ core_compress.py            # PDF / Word 压缩
├─ core_diff.py                # Word / Excel 差异比对
├─ core_img2pdf.py             # 图片转 PDF
├─ core_invoice.py             # 发票自动提取
├─ core_ocr.py                 # PDF OCR 增强
├─ core_pdf2img.py             # PDF 转图片
├─ core_pdf2word.py            # PDF 转 Word
├─ core_pdf_cleaner.py         # 扫描件去黑边
├─ core_split.py               # PDF 精准拆分
├─ core_unlock.py              # PDF 权限解密
├─ core_word2pdf.py            # Word / Excel 转 PDF
├─ core_word_merge.py          # Word 批量合并
├─ core_word_split.py          # Word 目录拆解
├─ bridge.py                   # 项目内置前后端桥接辅助文件
├─ main.py                     # 程序入口
├─ requirements.txt            # 依赖清单
├─ CHANGELOG.md                # 更新日志
├─ LICENSE                     # 源码授权说明
└─ EULA.md                     # 发布版程序使用条款
```

---

## 运行环境

### 系统要求

- Windows 10 / 11 x64
- Python 3.10

### Office 依赖

以下功能依赖本机已正确安装 Microsoft Office（至少 Word，部分场景需要 Excel）：

- Word 空白页清理中的 Word 分支
- Word 目录拆解
- Word 批量合并
- Word / Excel 转 PDF
- `.doc` 升级为 `.docx`
- 部分 Word 压缩相关流程

### OCR 与 PDF 运行时

项目依赖以下运行时能力：

- Tesseract
- Ghostscript
- Poppler

仓库已内置对应目录，`build_modern.py` 会在打包时同步这些资源。

---

## 安装与启动

### 1）安装依赖

```bash
pip install -r requirements.txt
```

### 2）启动项目

```bash
python main.py
```

启动后会创建一个本地桌面窗口，加载 `web/index.html` 作为主界面。

---

## 打包

项目已提供稳定版打包脚本：

```bash
python build_modern.py
```

### 打包脚本会完成的事情

- 清理旧的 `build/`、`dist/`、`build_spec/`
- 使用 PyInstaller 生成桌面应用
- 自动补齐 pywin32 相关 DLL
- 复制 `web/` 前端资源
- 复制 Ghostscript / Tesseract / Poppler 等运行时目录
- 校验 `dist/main/main.exe` 与关键资源是否完整

### 打包产物

```text
dist/
└─ main/
   ├─ main.exe
   ├─ web/
   ├─ Ghostscript/
   ├─ runtime/
   └─ poppler_bin/
```

---

## 开发建议

- 对 Word / Excel COM 类模块，建议继续保持串行化调用，避免多实例并发冲突。
- 对 PDF 处理模块，建议优先保证“稳定”和“保守”，避免因激进优化造成误删、误判或不可逆修改。
- 对外发布时，建议同步提供版本号、更新日志和已知问题说明，便于回溯与维护。

---

## 免责声明

本项目按“现状”提供。作者会尽量保证功能可用与结果可靠，但不对因使用本项目造成的数据丢失、业务中断、文档损坏或其他间接损失承担责任。对于重要文档，请务必先备份后再处理。

---

## 联系与授权

如需以下事项，请先联系作者并获得书面许可：

- 商用授权
- 企业内部定制部署
- 再分发 / 渠道分发
- 白标 / OEM 合作
- 源码授权或二次开发授权

如果这个项目对你有帮助，欢迎提出 Issue 或反馈使用建议。
