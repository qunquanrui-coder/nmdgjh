# 泉泉大人的工具合集（Baibaoxiang Web）

一个基于 **Python + Eel** 的 Windows 桌面端文档处理工具箱，覆盖 PDF、Word、OCR、图像转 PDF、发票提取、文档差异比对等常见办公场景。

---

## 功能一览

### 文档处理
- PDF → Word
- Word 空白页清理
- 扫描件去黑边
- PDF 精准拆分
- Word 目录拆解
- Word 批量合并
- PDF 权限解密
- 文档极限瘦身

### 转换与 OCR
- PDF OCR 增强（为扫描件添加可搜索文本层）
- 图像转 PDF
- Word → PDF
- PDF → 图片
- 发票自动提取
- 文档差异比对

---

## 技术栈

- **后端**：Python 3.10
- **桌面 UI**：Eel
- **PDF 处理**：PyMuPDF、pdfplumber、pypdf、pikepdf、OCRmyPDF
- **Word / Excel 处理**：python-docx、openpyxl、pandas、pywin32
- **OCR**：Tesseract、RapidOCR
- **图像处理**：OpenCV、Pillow
- **打包**：PyInstaller

---

## 运行环境

- Windows 10 / 11（x64）
- Python 3.10
- 已安装 Microsoft Word / Excel（仅 Word/Excel 原生导出相关功能需要）

---

## 本地运行

```powershell
pip install -r requirements.txt
python main.py