# 泉泉大人的工具合集 (Baibaoxiang Web)

一个基于 Python + Eel 的桌面端文档处理工具箱，集成 PDF、Word 转换与 OCR 功能。

## ✨ 功能一览

| 分类 | 功能 | 说明 |
|------|------|------|
| **文档处理** | PDF → Word | 高精度 PDF 转 Word |
| | Word 空白页清理 | 智能识别并删除空白页 |
| | 扫描件去黑边 | 自动检测去除扫描件四周黑框 |
| | PDF 精准拆分 | 按页码范围拆分 PDF |
| | Word 目录拆解 | 根据标题层级批量拆分为多文件 |
| | Word 批量合并 | 多个 Word 合并为一个文档 |
| | PDF 权限解密 | 移除 PDF 打开/编辑密码保护 |
| | 文档极限瘦身 | 压缩 PDF 体积 |
| **转换与 OCR** | PDF OCR 增强 | 扫描件添加可搜索文本层 |
| | 图像转 PDF | 多张图片批量生成 PDF |
| | Word → PDF | Word 文档转 PDF |
| | PDF → 图片 | PDF 每页导出为图片 |
| | 发票自动提取 | OCR 识别并提取发票信息 |
| | 文档差异比对 | 对比两份文档内容差异 |

## 🛠️ 技术栈

- **后端**: Python 3.10 + Eel (Web UI 框架)
- **PDF 处理**: PyMuPDF, pdfplumber, ocrmypdf, pikepdf
- **Word 处理**: python-docx, openpyxl, pandas
- **OCR**: Tesseract + RapidOCR
- **图像**: OpenCV, Pillow
- **打包**: PyInstaller

## 📦 依赖组件

项目内置以下第三方引擎（已包含在仓库中）：

| 组件 | 用途 |
|------|------|
| [Ghostscript](https://www.ghostscript.com/) | PDF 渲染与压缩 |
| [Poppler](https://poppler.freedesktop.org/) | PDF 文本提取与转换 |
| [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) | 光学字符识别（含中文简体模型） |

## 🚀 快速开始

### 环境要求

- Windows 10/11 (x64)
- Python 3.10

### 本地运行

```powershell
# 安装依赖
pip install -r requirements.txt

# 启动应用
python main.py
```

### 打包发布

项目使用 `build_modern.py` 进行 PyInstaller 打包：

```powershell
python build_modern.py
```

打包输出位于 `dist/main/`，包含完整的可执行文件和所有依赖。

## 🔄 GitHub Actions 自动构建

每次推送到 `main` 分支会自动触发 CI/CD 流程：

1. 安装 Python 依赖
2. 运行 PyInstaller 打包
3. 将完整输出目录压缩为 zip
4. 发布到 [Releases](https://github.com/qunquanrui-coder/nmdgjh/releases)

手动触发：Actions → "Build & Release" → "Run workflow"

## 📁 项目结构

```
Baibaoxiang_Web/
├── main.py                  # 主入口 (Eel + COM 线程管理)
├── build_modern.py          # PyInstaller 打包脚本
├── requirements.txt         # Python 依赖清单
├── core_*.py                # 功能核心模块 (15+ 个)
├── web/                     # Web UI 资源
│   ├── index.html           # 主页面
│   ├── style.css            # 样式表
│   └── script.js            # 前端逻辑
├── Ghostscript/bin/         # Ghostscript 运行时 (排除在 Git LFS 外)
├── poppler_bin/             # Poppler 工具集
└── runtime/Tesseract/       # Tesseract OCR + 语言模型
```

## ⚠️ 注意事项

- COM 操作（Word 处理）采用线程隔离与排队锁机制，防止 Office 假死
- **首次运行会自动打开 Chrome 浏览器窗口**；如果系统未安装 Chrome，Eel 会自动回退到 Edge
- **打包后的程序目录必须包含以下子文件夹才能正常运行**：`Ghostscript/`、`poppler_bin/`、`runtime/Tesseract/`（这些已随 dist/main/ 一起打包）

## 📄 License

Private Project - All rights reserved.
