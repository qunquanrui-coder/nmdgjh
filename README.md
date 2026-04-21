# 泉泉大人的工具合集（Baibaoxiang Web）

一个面向 **Windows 办公场景** 的桌面文档工具箱，聚焦 PDF、Word、Excel、OCR 与发票处理等高频需求。

项目采用 **Python + pywebview + 本地 Web 前端** 架构，围绕“**能直接用、能批量跑、能打包发布**”的目标，提供 PDF 转 Word、Word/Excel 转 PDF、扫描件 OCR、文档压缩、差异比对、发票提取、图片转 PDF 等一整套本地化处理能力。

当前仓库已经从早期浏览器壳模式迁移到 **pywebview 桌面壳**，同时保留了 **Eel 风格的前后端调用方式**，适合继续迭代成稳定的 Windows 单文件/目录版工具。 

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
  - 支持 **可编辑模式** 与 **纯图固化模式**
  - 可对单个 PDF 或整个文件夹递归处理
- **Word / PDF 空白页清理**
  - Word 通过 COM 分页与页范围删除处理
  - PDF 通过文本、图像、矢量与注释联合判空
- **扫描件去黑边**
  - 基于 OpenCV 识别有效内容边界并对白边区域进行覆盖
- **PDF 精准拆分**
  - 支持 **定长拆分 / 平均拆分 / 范围提取**
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
  - 采用“**目标区间优先**”策略，不盲目压得过小

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

- **Eel 风格接口暴露 + 自定义桥接层**
  - 业务模块仍然使用 `@eel.expose` 暴露方法
  - `app_api.py` 负责把前端调用路由到 Python 端
  - 对 COM 类任务增加串行锁，避免 Word / Excel 并发冲突

### 前端

- **HTML + CSS + JavaScript**
- 页面文件位于 `web/` 目录：
  - `index.html`
  - `script.js`
  - `style.css`

### 后端核心

- **Python 3.10**
- 按功能拆分为多个 `core_*.py` 模块，入口统一由 `main.py` 加载

---

## 技术栈

### 通用与桌面层

- Python 3.10
- pywebview
- Eel（兼容式调用层）
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
├─ app_api.py                  # pywebview 与 eel 风格接口桥接
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
├─ eel.py                      # eel 兼容桥接辅助文件
├─ main.py                     # 程序入口
├─ requirements.txt            # 依赖清单
├─ CHANGELOG.md                # 更新日志
└─ LICENSE.txt                 # 授权说明
```

---

## 运行环境

### 系统要求

- **Windows 10 / 11 x64**
- **Python 3.10**

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

仓库已经内置对应目录，`build_modern.py` 也会在打包时同步这些资源。

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

## 关键实现说明

### 1）桌面层已迁移到 pywebview

这是当前项目一个很关键的状态变化。

项目不再依赖“默认拉起浏览器页面”的旧模式，而是：

- 用 **pywebview** 创建桌面窗口
- 用本地页面作为 UI
- 通过 `app_api.py` 把前端调用转发到 Python
- 继续保留 `@eel.expose` 这一套业务模块调用习惯

这意味着：

- 旧模块迁移成本更低
- 前端交互方式基本保持不变
- 打包后的稳定性更适合桌面工具发布

### 2）COM 任务做了串行保护

Word / Excel 相关功能普遍容易在并发调用时出现异常，因此项目增加了：

- COM 初始化与反初始化管理
- 串行锁保护
- 文档 / 工作簿 / 进程的兜底关闭
- 垃圾回收与句柄清理

这对于以下模块尤其重要：

- Word 拆分
- Word 合并
- Word / Excel 转 PDF
- Word 空白页清理
- `.doc` 到 `.docx` 升级

### 3）长任务有心跳日志

对于 OCR、大文件压缩、PDF 转 Word 等高耗时任务，项目加入了：

- 前端终端日志输出
- 长任务心跳提示
- 进度状态回传

这样用户在处理大文件时，不会感觉程序“卡死”。

### 4）压缩策略更偏向“接近目标”

压缩模块不是简单地“越小越好”，而是采用：

1. 优先选择落在目标区间内的结果
2. 如果没有，则选择低于目标但最接近目标的结果
3. 仍没有时，再选择整体最接近目标的结果

其中：

- PDF 压缩理想下限约为目标值的 **80%**
- Word 压缩理想下限约为目标值的 **85%**

这个策略更适合实际办公场景，能尽量平衡“体积”与“可读性”。

---

## 模块级说明

### PDF 转 Word（`core_pdf2word.py`）

- 可编辑模式：基于 `pdf2docx`
- 纯图模式：逐页渲染图片后写入 Word
- 支持单文件与目录递归处理

### 空白页清理（`core_blank_page.py`）

- PDF：通过文本、图像、矢量、注释多维判空
- Word：通过 COM 分页与页范围删除处理

### PDF 拆分（`core_split.py`）

- 支持定长 / 平均 / 范围提取
- 提供多层 PDF 打开兜底策略，兼容异常文件

### Word 拆分（`core_word_split.py`）

- 先扫描有效大纲级别
- 再按标题级别拆成多个独立文档
- 对 Word 拒绝响应、剪贴板异常等情况做了重试与回收

### Word 合并（`core_word_merge.py`）

- 自然排序读取文档
- 逐篇插入标题与分页
- 使用复制粘贴合并，并做剪贴板清理重试

### PDF 解密（`core_unlock.py`）

- 基于 PyMuPDF 保存为无权限限制的新 PDF
- 使用多进程隔离降低崩溃对主界面的影响

### 文档压缩（`core_compress.py`）

- PDF：`pikepdf + Ghostscript`
- Word：解包 `docx` 后重压缩 `word/media/` 中图片资源
- 支持 `.doc` 自动升级到 `.docx`

### OCR（`core_ocr.py`）

- 基于 OCRmyPDF
- 自动配置 Tesseract 与 Ghostscript 路径
- 对进度输出和子进程窗口隐藏做了处理

### 扫描件去黑边（`core_pdf_cleaner.py`）

- 逐页渲染为图像
- 用 OpenCV 做阈值化与轮廓识别
- 自动计算有效内容区域并对外围黑边进行白色覆盖

### 发票提取（`core_invoice.py`）

- 文本提取优先，OCR 兜底
- 自动识别发票代码、号码、日期、税率、金额、税额、价税合计
- 汇总导出 Excel，并对重复与异常项做记录

### 差异比对（`core_diff.py`）

- Word：按段落使用 `difflib.SequenceMatcher` 对比
- Excel：按工作表逐单元格扫描差异
- 最终输出格式化 Excel 报告

---

## 已知限制

- 当前项目明确面向 **Windows**，不适合直接在 macOS / Linux 上运行。
- 部分功能强依赖 **本机 Office COM**，没有安装 Word / Excel 时无法使用。
- Word 压缩主要针对文档中的图片资源，**纯文本型 docx 收益可能有限**。
- OCR 当前核心输入对象是 **PDF**，不是通用文档 OCR 平台。
- 这是一个偏本地桌面工具的仓库，不是标准 Web 服务项目。

---

## 开发建议

如果后续继续演进，比较建议优先做这几件事：

1. **补充异常日志与统一错误码**，方便定位打包后问题。
2. **为每个核心模块补充独立测试样例**，尤其是 Office / OCR / 压缩链路。
3. **把运行时依赖检查前置**，在前端启动时直接提示缺失项。
4. **统一输出目录和命名策略**，减少批量处理时的歧义。
5. **补充版本说明与截图**，让 GitHub 首页更适合对外展示。

---

## 版本说明

当前仓库已经发布到 **v1.5**，并在该阶段完成了比较关键的桌面壳迁移与打包稳定化工作，包括：

- 从旧浏览器壳切换到 pywebview
- 增加本地桥接层
- 修复前端资源定位与打包后的路径问题
- 优化 PDF / Word 压缩选优策略
- 增加长任务心跳日志

---

## 授权说明

本仓库 **不是开源许可证项目**。

根据当前 `LICENSE.txt`：

- 仓库公开可见，仅用于源码查看、问题反馈、版本分发和评估
- **不授予** 使用、复制、修改、发布、分发、再许可、销售或派生开发权限
- 如需商用、组织内部部署、再分发或其他授权，请先获得版权方书面许可



---

## 适用场景

这个项目尤其适合以下场景：

- 日常 PDF / Word / Excel 办公处理
- 扫描件 OCR 与清洗
- 标书、合同、档案类文档拆分与合并
- 发票批量整理与汇总
- 内网、离线、本地敏感文档处理
- 需要打包成 Windows 桌面工具给同事直接使用的场景

---

## 致谢

项目依赖了多个优秀的第三方库与运行时组件，包括但不限于：

- pywebview
- Eel
- PyMuPDF
- OCRmyPDF
- Tesseract
- Ghostscript
- pdf2docx
- pikepdf
- OpenCV
- Pillow
- pandas
- openpyxl
- python-docx
- pywin32

请在实际分发与使用时，分别遵循这些第三方组件各自的许可证要求。
