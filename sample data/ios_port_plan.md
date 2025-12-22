# iOS 移植可行性与部署方案

## 现状评估
- 桌面端基于 Electron + Node（`electron/main.js`，`server/app.js`），iOS 无法运行 Node 进程或 Electron 容器。
- 样式提取依赖本地 Python 子进程（`lib/python-bridge.js` 调用 `style_extractor.py`），iOS 沙箱内无法启动解释器。
- 生成与导入逻辑大量使用 `fs/path/child_process`、`express` 等服务端 API；在 WKWebView/JSCore 环境中不可用。
- 依赖包 `docx` 支持浏览器，但当前实现依赖 `fs`、`image-size`、`axios`（下载远程图片）等 Node 能力，需要改造。

结论：原始代码不能直接移植到 iOS，也无法在纯本地模式下运行，需较大规模重构或改写。

## 可选方案

### 1) 纯前端/PWA 方案（单机可行，但需重构）
- 将转换逻辑移至浏览器可运行的打包版本：`docx` 使用 `Packer.toBlob`，`mammoth` 的 browser 版本或基于 `JSZip` 的自研样式/内容解析。
- 替换 `fs/path/image-size/child_process`：文件通过 `<input type="file">` 读取 `ArrayBuffer`，图片尺寸用 `<img>` 或 `canvas` 计算，完全移除 Python 子进程。
- 样式提取：用 JS 解析 `.docx`（`JSZip` + 读取 `word/styles.xml`）提取字体/边距，或在 UI 里手动配置，不再依赖 Python。
- 打包为静态资源（Vite/Webpack）后放入 WKWebView 或安装为 PWA，即可在无网条件下本地转换。
- 适合 App Store 审核（无解释器、无动态下载依赖），但需要较大代码改造与前端化。

### 2) Native 包壳 + JS Core 方案（单机可行，改造中等）
- 用 Swift/SwiftUI 写 UI，核心转换用浏览器版的 JS/WASM 库在 `WKWebView` 里执行。
- 文件选取/保存使用 `UIDocumentPicker`，与 WebView 通过消息桥传输二进制。
- 样式提取同上使用纯 JS 解析，或使用本地 Swift 库（如 XML 解析 `styles.xml`），完全移除 Python。
- 依赖网络下载图片的场景需额外网络权限；纯本地资源则可离线。

### 3) 后端保留 Node，iOS 仅做薄客户端（需联网）
- 直接复用现有 `server/app.js` 和 Python，部署在服务器或局域网。
- iOS 端只做上传/下载和预览，转换在服务器完成；实现成本低，但必须联网，且需处理隐私/大文件上传。

## 推荐路径
- 目标是“单机部署”：选择方案 1 或 2，先移除对 Python 和 Node 内置模块的依赖，转向纯前端/JS 解析与生成。
- 若时间紧需快速可用版本：先落地方案 3（服务端部署），iOS 用 WebView 或原生调用 HTTP；后续再逐步前端化以实现离线。

## 预估改造工作量
- 前端化打包与 API 替换：约 1–2 周（视功能完备度）。
- 样式提取 JS 化：简单提取（页边距/字体）约 2–4 天，复杂还原（表格/段落）更久。
- iOS 壳与文件交互：SwiftUI + WKWebView 基础功能 2–4 天。

## 兼容性与注意事项
- 远程图片下载需要网络；离线模式需提示用户仅使用本地图片或提前缓存。
- 需避免运行时动态解释器、JIT 插件，确保符合 App Store 审核。
- 大文件处理建议分片读取或使用 Web Worker，避免 WebView 卡顿。
