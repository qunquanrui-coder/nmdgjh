# 更新日志

## v1.5

- 将 Windows 打包产物从 `main.exe` 更名为 `QuanQuanTreasureBox.exe`。
- 补齐 Windows 文件版本信息，产品名统一为“泉泉的百宝箱”。
- 迁移到 pywebview 桌面壳，使用本地 Web 前端作为主界面。
- 引入 bridge / app_api 前后端桥接层，统一前端调用入口。
- 完善 Word / Excel COM 类任务的串行执行与资源释放。
- 增强长任务心跳日志，改善 OCR、压缩、转换等耗时任务的反馈。
- 优化 PDF / Word 压缩结果选择策略，优先接近目标大小而不是盲目压缩。
- 修复打包后前端资源和运行时目录定位问题。

## v1.1

- 发布首个稳定版本。
- 集成 PDF、Word、Excel、OCR、发票提取、文档压缩与差异比对等核心办公工具。
- 增加 PyInstaller 打包脚本和 GitHub Actions Windows 构建流程。
