# TODO
- 完成 HTML 预览稳健性验证：导入多样 HTML（含外链图片/样式、纯文本），确认无报错、可导出 PNG/PDF。
- CLI 无头截图 (`docxjs-capture`) 增强：增加错误提示（缺少浏览器/端口占用）和示例文档。
- 签名与发布：准备 macOS 签名/公证证书，重打带签名的 DMG。
- fsevents 依赖：在非 iCloud 路径验证正常安装；必要时改为可选依赖或文档说明。
- 回归测试：Markdown→Docx、HTML 预览/导出、CLI Capture 端到端脚本化验证。
