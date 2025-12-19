# 导出 PDF 与 图片功能设计方案 (Export Scheme)

**当前版本**: DocxJS CLI/App v1.3.4
**设计原则**: 优先生成 DOCX (Source of Truth)，基于 DOCX 的渲染结果进行转换。

---

## 1. 总体策略 (Core Strategy)

鉴于我们已经完善了 `lib/core.js` 的后端 DOCX 生成逻辑（包括精确的行距 `lineRule: exact` 和样式控制），导出 PDF 和图片应尽可能复用这一成果，而不是重新写一套渲染逻辑。

**路径**: `Markdown` -> `Backend (lib/core.js)` -> `DOCX Blob` -> `Frontend Renderer (docx-preview)` -> `PDF/Image`

---

## 2. PDF 导出方案 (Export to PDF)

### 方案 A：Electron 原生打印 (推荐)
利用 Electron 的 `webContents.printToPDF` 接口。由于前端已经通过 `docx-preview.js` 将 DOCX Blob 渲染为高保真的 HTML/Canvas 视图，我们可以直接将这个视图“打印”为 PDF。

**实现步骤**:
1.  **用户点击 "导出 PDF"**。
2.  **生成 DOCX**: 调用后端 `API.convert()` 获取最新的 DOCX Blob。
3.  **隐藏渲染**: 在前端创建一个隐藏的 `<iframe>` 或不可见的 `div`，利用 `docx.renderAsync` 将 DOCX Blob 渲染进去。
4.  **打印**:
    *   主进程 (Main Process) 接收指令。
    *   使用 `win.webContents.printToPDF(options)`。
    *   `options`: 设置 `printBackground: true`, `pageSize: 'A4'`, `marginsType: 0` (无边距，因为 DOCX 渲染器已经有了页边距)。
5.  **保存**: 将生成的 PDF Buffer 写入用户指定路径。

**优点**:
*   零外部依赖 (不需要用户安装 LibreOffice/Word)。
*   “所见即所得” (PDF 与 预览界面完全一致)。
*   跨平台 (Windows/macOS 表现一致)。

**缺点**:
*   依赖 `docx-preview` 的渲染精度（目前看来对表格和文本支持良好，但可能不支持某些极复杂的 Word 特性）。

---

## 3. 图片导出方案 (Export to Image)

### 方案 A：前端截图 (html2canvas)
利用前端库将渲染好的 DOM 节点转换为 Canvas，再导出为图片。

**实现步骤**:
1.  **用户点击 "导出长图/图片"**。
2.  **获取渲染容器**: 定位到预览区域的 DOM 节点 (`#paper-container` 或具体的 `.docx-wrapper`).
3.  **截图**:
    *   引入 `html2canvas` 库。
    *   调用 `html2canvas(element, { scale: 2 })` (使用 2倍或更高的 scale 以保证清晰度)。
4.  **导出**: 将 Canvas 转换为 Base64/Blob (PNG/JPG)，触发下载。

**优点**:
*   实现简单，纯前端完成。
*   支持“长图”模式（将所有页面拼接的一张图）。

**缺点**:
*   如果文档非常长，Canvas 可能会受到浏览器内存限制或最大高度限制。

### 方案 B：PDF 转图片 (后端)
如果已经实现了 PDF 导出，可以在后端利用工具（如 `pdftocairo` 或 `pdf-poppler`，但这引入了二进制依赖）将 PDF 转为图片。

**推荐选择**: **方案 A (html2canvas)**，对于当前轻量级应用最合适。

---

## 4. 实施路线图 (Implementation Roadmap)

### 第一阶段：PDF 导出
1.  在 `electron/preload.js` 中暴露 `ipcRenderer.invoke('print-to-pdf')`。
2.  在 `electron/main.js` 中实现处理函数，调用 `webContents.printToPDF`。
3.  在 UI 增加“导出 PDF”按钮。

### 第二阶段：图片导出
1.  安装 `html2canvas` (`npm install html2canvas`).
2.  在 `public/index.html` 中引入该库。
3.  实现“导出图片”按钮逻辑，将 `#paper-container` 内容转换为图片并下载。

---

## 5. 备选高保真方案 (Advanced High-Fidelity)

如果用户环境安装了 **LibreOffice** 或 **Microsoft Word**，我们可以通过命令行调用它们进行无头转换 (Headless Conversion)。

*   **命令**: `soffice --headless --convert-to pdf output.docx`
*   **逻辑**: App 检测环境变量 -> 如果存在 soffice -> 使用该方案 (100% 还原) -> 否则回退到 Electron 打印方案。

*目前建议先实现 Electron 原生方案，作为默认的基础功能。*
