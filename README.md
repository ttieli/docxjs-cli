# DocxJS Converter (CLI & Desktop App)

[中文说明](#中文说明) | [English](#english)

**Current Version / 当前版本**: `1.3.27`

A powerful, **hybrid tool (CLI & Desktop)** that converts Markdown to high-fidelity Word (.docx) documents. It combines the generation capabilities of Node.js with the style parsing capabilities of Python, specifically optimized for **Chinese Official Document formats (党政机关公文格式)** and standard business reports.

一个强大的 **Markdown 转 Docx 工具（支持命令行与桌面端）**。它结合了 Node.js 的生成能力和 Python 的样式解析能力，专为生成符合**中国党政机关公文格式**及标准商务报告的文档而优化。

### 🆕 Latest Updates (v1.3.27) / 最新更新

- **Zoom Controls**: Word-like zoom percentage display with +/- buttons (25%-200%), click to reset
- **Table Spacing Fix**: Fixed input fields accepting 0 values for table spacing settings
- **Improved UI Labels**: Table spacing inputs now show unit hints (twips, 200≈10pt)

- **缩放控制**：类似 Word 的缩放百分比显示，支持 +/- 按钮调整（25%-200%），点击重置
- **表格间距修复**：修复表格间距设置输入 0 时无效的问题
- **优化标签提示**：表格间距输入框显示单位提示 (twips, 200≈10pt)

## 🏗️ Architecture / 整体架构

DocxJS Converter is built as a multi-tier system to provide flexibility across CLI, Web, and Desktop environments.

DocxJS Converter 采用多层架构设计，确保在命令行、网页和桌面端均能提供一致的体验。

### Core Components / 核心组件

1.  **Core Engine (`lib/core.js`)**: The heart of the system. It parses Markdown and uses the `docx` library to generate OpenXML documents based on a unified `StyleConfig` object.
    *   **核心引擎**：系统的核心，解析 Markdown 并使用 `docx` 库根据统一的 `StyleConfig` 对象生成 OpenXML 文档。
2.  **Style Normalizer (`lib/style-normalizer.js`)**: Ensures that styles from different sources (UI, CLI, templates) are validated and converted into the internal format used by the engine.
    *   **样式标准化器**：确保来自不同来源（UI、命令行、模板）的样式经过校验并转换为引擎使用的内部格式。
3.  **Desktop App (`electron/`)**: A cross-platform GUI built with Electron. It provides a real-time side-by-side preview using `docx-preview` and a sandboxed `iframe` for HTML mode.
    *   **桌面端应用**：基于 Electron 的跨平台 GUI。使用 `docx-preview` 提供实时左右对比预览，并为 HTML 模式提供沙箱化的 `iframe` 环境。
4.  **Web Server (`server/app.js`)**: An Express-based backend that exposes the core engine via a RESTful API, enabling the same functionality in browser-only environments.
    *   **Web 服务器**：基于 Express 的后端，通过 RESTful API 暴露核心引擎功能，使得在纯浏览器环境下也能实现相同功能。
5.  **CLI Tools (`bin/`)**:
    *   `docxjs`: Direct Markdown-to-Docx conversion.
    *   `docxjs-capture`: A headless renderer using **Playwright** to capture the exact UI preview as PNG or PDF.
    *   **命令行工具**：`docxjs` 负责直接转换；`docxjs-capture` 利用 **Playwright** 无头浏览器捕获与 UI 完全一致的预览截图或 PDF。
6.  **Python Bridge (`style_extractor.py` & `lib/python-bridge.js`)**: Leverages Python's `python-docx` to extract styling metadata (fonts, margins) from existing Word documents, which is then fed back into the Node.js engine.
    *   **Python 桥接**：利用 Python 的 `python-docx` 从现有 Word 文档中提取样式元数据（字体、边距），并反馈给 Node.js 引擎。

---

<a name="english"></a>
## 🇬🇧 English Documentation

### ✨ Key Features

*   **Desktop Application**: A standalone Electron app for Windows & macOS (Intel & Apple Silicon). No Node/Python installation required for end-users.
*   **Professional Mode (New!)**: Built-in JSON editor in the UI allows direct modification of the underlying style configuration for ultimate flexibility.
*   **Extended Heading Support**: Now supports full styling for **Heading Levels 1 through 6 (H1-H6)**.
*   **Visual Editor**: 
    *   **Real-time Preview**: WYSIWYG editor with split-pane layout (Settings / Preview).
    *   **Inline Markdown Editing**: Edit your content directly within the preview interface without pop-ups.
    *   **Bilingual Templates**: Built-in templates with clear bilingual names (e.g., "Official Red", "Business Contract").
*   **HTML Preview & Export (New)**: Import `.html`/`.htm`, sandboxed preview, and export as PNG/PDF via the same renderer.
*   **CLI Capture (New)**: `docxjs-capture` renders the app headlessly (Playwright) to export PNG/PDF matching UI preview (Markdown or HTML mode).
*   **Official Government Style**: Strict adherence to Chinese "Red Header" document standards (fonts, margins, solid borders).
*   **Hybrid Style Extraction**: Can extract styles (margins, fonts) from an existing `.docx` file to apply to your new document.
*   **Smart CLI Defaults**: If no output path is specified, the CLI automatically generates a file in the input directory with the format `{filename}_{timestamp}.docx`.
*   **Enhanced Layout Control**: Templates now support explicit `titleAlignment` (e.g., center) and `paragraphIndent` (e.g., first-line indent).
*   **Intuitive Font Sizes**: Editor inputs now use **Points (pt)** standard. Built-in mapping displays the corresponding Chinese font size name (e.g., "小四" for 12pt).
*   **Precise Line Spacing**: Supports `exact` (fixed) vs `auto` (multiple) line spacing rules to ensure perfect rendering across WPS and Word.

### 🖥️ Desktop Application

**Download**: Check the [GitHub Releases](https://github.com/ttieli/docxjs-cli/releases) page for `.dmg` (macOS) or `.exe` (Windows) installers.

**Build from Source**:
```bash
# Install dependencies
npm install

# Build for macOS (Auto-detects arch, use --arm64 for Apple Silicon)
npm run dist:mac

# Build for Windows
npm run dist:win
```

### 📦 CLI Installation

For developers who prefer the command line:

**NPM Install (Recommended):**

```bash
npm install -g ttieli/docxjs-cli
```

**One-line Install (with Python environment setup):**

This will automatically set up the required Python environment and install the tool globally.

```bash
curl -sSL https://raw.githubusercontent.com/ttieli/docxjs-cli/main/install_global.sh | bash
```

**Usage:**

```bash
# Basic conversion
docxjs input.md -o output.docx

# Export as PNG only (supports local images, requires playwright)
docxjs input.md --image

# Capture (headless PNG/PDF, matches UI)
docxjs-capture --input input.md --png out.png --pdf out.pdf
# For HTML input
docxjs-capture --input page.html --mode html --png out.png
# (Requires Playwright; installed by default. On first run it downloads Chromium.)
```

### 🚀 Usage Guide (App/Web)

1.  **Select Template**: Choose from built-in styles like "General", "Official Red", or "Business Contract".
2.  **Edit Content**: Import a Markdown/Docx file or edit directly in the "Markdown Source" panel.
3.  **Customize Style**:
    *   Use the sidebar controls to adjust fonts, sizes, and colors for H1-H6, Body, and Tables.
    *   **Professional Mode**: Open the "Professional Mode (JSON)" accordion to paste or edit the raw JSON config.
4.  **Export**: Click "Export Docx" to generate the final file.
5.  **HTML Mode** (New): Click “Import HTML” → switch to HTML preview → export PNG/PDF directly from the rendered HTML. (Docx export remains Markdown-only.)

### ⚙️ Template Configuration

Templates are defined in `templates/templates.json`. You can now customize:
*   **Fonts/Sizes/Colors**: For Body text and Headings H1-H6. Note: CLI JSON config uses **Half-points** for font sizes (e.g., 24 = 12pt).
*   **Margins**: Precise control (in twips).
*   **Tables**: Border styles (single/dotted), width, and colors.
*   **Line Spacing**: In twips (e.g., 560 = 28pt). Use `"lineRule": "exact"` for fixed values (Official Docs) or `"auto"` for multiples (General).
*   **Title Alignment**: `titleAlignment`: "center" | "left" | "right" (For H1).
*   **Paragraph Indent**: `paragraphIndent`: Number (in twips, e.g., 640 for 2 chars).

---

<a name="中文说明"></a>
## 🇨🇳 中文说明

### ✨ 核心特性

*   **桌面客户端**：支持 Windows 和 macOS (Intel & M1/M2) 的独立应用，开箱即用，无需配置环境。
*   **专业模式 (新增!)**：UI 内置 JSON 编辑器，允许直接修改底层的样式配置对象，提供无限的定制灵活性。
*   **全级标题支持**：现已完整支持 **一级到六级标题 (H1-H6)** 的独立样式设置（字体、字号、颜色）。
*   **可视化编辑器**：
    *   **实时预览**：左侧配置样式，右侧即时查看 A4 纸张渲染效果。
    *   **行内编辑**：直接在预览界面中修改 Markdown 源码，所见即所得。
    *   **双语模板**：内置清晰的中英双语模板名称（如“政府公文 (红头)”、“商务合同”）。
*   **HTML 预览与导出（新增）**：导入 `.html/.htm`，在沙箱容器内预览，并直接导出 PNG/PDF。
*   **CLI 截图导出（新增）**：`docxjs-capture` 使用 Playwright 以无头方式渲染 UI，输出与界面一致的 PNG/PDF（支持 Markdown/HTML）。
*   **党政机关公文标准**：严格遵循国家公文格式标准（红头、仿宋/小标宋字体模拟、标准页边距、实线表格）。
*   **样式提取**：支持从现有的 Word 文档中“吸取”页边距和字体样式。
*   **智能命令行默认值**：如果未指定输出路径，CLI 会自动在输入文件同目录下生成格式为 `{原文件名}_{时间戳}.docx` 的文件。
*   **增强版式控制**：模板现支持显式配置 `titleAlignment`（标题对齐）和 `paragraphIndent`（段落首行缩进）。
*   **直观字号映射**：编辑器输入框统一采用 **磅 (pt)** 为单位，并自动显示对应的中文字号（如输入 12 显示“小四”），解决了之前的单位换算困扰。
*   **精准行距控制**：支持配置 `lineRule` 为 `exact`（固定值）或 `auto`（倍数），完美适配公文对固定行距的严苛要求，解决 WPS/Word 渲染差异。

### 🖥️ 下载与安装

**下载地址**: 请访问 [GitHub Releases](https://github.com/ttieli/docxjs-cli/releases) 页面下载最新版。

**源码构建**:
```bash
# 1. 安装依赖
npm install

# 2. 构建 macOS 版本 (自动识别架构，M1/M2/M3 请使用 --arm64)
npm run dist:mac -- --arm64

# 3. 构建 Windows 版本
npm run dist:win
```

### 📦 安装方法

如果您习惯使用终端：

**NPM 安装（推荐）:**

```bash
npm install -g ttieli/docxjs-cli
```

**一键安装（含 Python 环境配置）:**

这将自动配置所需的 Python 环境并全局安装工具。

```bash
curl -sSL https://raw.githubusercontent.com/ttieli/docxjs-cli/main/install_global.sh | bash
```

**使用方法:**

```bash
# 基础转换
docxjs input.md -o output.docx
# 或者仅指定输入文件（自动生成输出名）：
docxjs input.md

# 仅导出 PNG 图片（支持本地图片，需 Playwright）
docxjs input.md --image

# 截图导出（需 Playwright，首次会下载 Chromium）
docxjs-capture --input input.md --png out.png --pdf out.pdf
# HTML 输入
docxjs-capture --input page.html --mode html --png out.png
```

### 🚀 使用指南 (桌面版/Web)

1.  **选择模板**：在左侧下拉框选择基础风格，例如“通用样式”或“政府公文 (红头)”。
2.  **编辑内容**：点击工具栏的“导入文件”或“编辑内容”按钮，修改文档正文。
3.  **样式微调**：
    *   通过侧边栏调整正文及 H1-H6 标题的字体、字号、颜色。
    *   **专业模式**：展开底部的“专业模式 (Professional JSON)”面板，直接编辑 JSON 配置，实现界面控件无法覆盖的高级定制。
4.  **导出**：点击“导出 Docx”生成最终的 Word 文档。
5.  **HTML 模式（新增）**：点击“导入 HTML”后切换到 HTML 预览，可直接导出 PNG/PDF（Docx 导出仍基于 Markdown）。

### ⚙️ 模板配置说明

所有预设模板均位于 `templates/templates.json`。支持的配置项包括：
*   **字体/字号/颜色**：覆盖正文及 H1-H6 所有层级。注意：CLI 的 JSON 配置文件中，字号单位为 **半磅 (Half-points)** (例如 24 代表 12pt)。桌面端编辑器会自动处理此换算。
*   **页边距**：精确控制上下左右边距 (单位: twips)。
*   **表格样式**：支持设置边框类型（实线/虚线）、粗细及表头样式。
*   **行间距**：固定值行距 (单位: twips, 1磅=20 twips)。设置 `"lineRule": "exact"` 启用固定行距（公文），`"auto"` 为倍数行距（通用）。
*   **标题对齐**：`titleAlignment`: "center" | "left" | "right" (仅限 H1)。
*   **段落缩进**：`paragraphIndent`: 数值 (单位 twips, 如 640 约等于两个汉字)。

---

## 🤝 Contributing / 贡献

*   **Bug Reports**: Welcome via Issues.
*   **Pull Requests**: Please adhere to the existing code style and update the version in `package.json`.

**License**: ISC
