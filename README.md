# DocxJS Converter (CLI & Desktop App)

[中文说明](#中文说明) | [English](#english)

**Current Version / 当前版本**: `1.3.0`

A powerful, **hybrid tool (CLI & Desktop)** that converts Markdown to high-fidelity Word (.docx) documents. It combines the generation capabilities of Node.js with the style parsing capabilities of Python, specifically optimized for **Chinese Official Document formats (党政机关公文格式)** and standard business reports.

一个强大的 **Markdown 转 Docx 工具（支持命令行与桌面端）**。它结合了 Node.js 的生成能力和 Python 的样式解析能力，专为生成符合**中国党政机关公文格式**及标准商务报告的文档而优化。

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
*   **Official Government Style**: Strict adherence to Chinese "Red Header" document standards (fonts, margins, solid borders).
*   **Hybrid Style Extraction**: Can extract styles (margins, fonts) from an existing `.docx` file to apply to your new document.

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

**One-line Install (Recommended):**

This will automatically set up the required Python environment and install the tool globally.

```bash
curl -sSL https://raw.githubusercontent.com/ttieli/docxjs-cli/main/install_global.sh | bash
```

**Manual Install:**

```bash
git clone https://github.com/ttieli/docxjs-cli.git
cd docxjs-cli
./install_global.sh
```

```bash
# Global install via npm
npm install -g docxjs-cli

# Usage
docxjs input.md -o output.docx
```

### 🚀 Usage Guide (App/Web)

1.  **Select Template**: Choose from built-in styles like "General", "Official Red", or "Business Contract".
2.  **Edit Content**: Import a Markdown/Docx file or edit directly in the "Markdown Source" panel.
3.  **Customize Style**:
    *   Use the sidebar controls to adjust fonts, sizes, and colors for H1-H6, Body, and Tables.
    *   **Professional Mode**: Open the "Professional Mode (JSON)" accordion to paste or edit the raw JSON config.
4.  **Export**: Click "Export Docx" to generate the final file.

### ⚙️ Template Configuration

Templates are defined in `templates/templates.json`. You can now customize:
*   **Fonts/Sizes/Colors**: For Body text and Headings H1-H6.
*   **Margins**: Precise control (in twips).
*   **Tables**: Border styles (single/dotted), width, and colors.
*   **Line Spacing**: In twips (e.g., 560 = 28pt).

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
*   **党政机关公文标准**：严格遵循国家公文格式标准（红头、仿宋/小标宋字体模拟、标准页边距、实线表格）。
*   **样式提取**：支持从现有的 Word 文档中“吸取”页边距和字体样式。

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

**一键安装（推荐）:**

这将自动配置所需的 Python 环境并全局安装工具。

```bash
curl -sSL https://raw.githubusercontent.com/ttieli/docxjs-cli/main/install_global.sh | bash
```

**手动安装:**

```bash
git clone https://github.com/ttieli/docxjs-cli.git
cd docxjs-cli
./install_global.sh
```

```bash
# 全局安装
npm install -g docxjs-cli

# 基础转换
docxjs input.md -o output.docx
```

### 🚀 使用指南 (桌面版/Web)

1.  **选择模板**：在左侧下拉框选择基础风格，例如“通用样式”或“政府公文 (红头)”。
2.  **编辑内容**：点击工具栏的“导入文件”或“编辑内容”按钮，修改文档正文。
3.  **样式微调**：
    *   通过侧边栏调整正文及 H1-H6 标题的字体、字号、颜色。
    *   **专业模式**：展开底部的“专业模式 (Professional JSON)”面板，直接编辑 JSON 配置，实现界面控件无法覆盖的高级定制。
4.  **导出**：点击“导出 Docx”生成最终的 Word 文档。

### ⚙️ 模板配置说明

所有预设模板均位于 `templates/templates.json`。支持的配置项包括：
*   **字体/字号/颜色**：覆盖正文及 H1-H6 所有层级。
*   **页边距**：精确控制上下左右边距 (单位: twips)。
*   **表格样式**：支持设置边框类型（实线/虚线）、粗细及表头样式。
*   **行间距**：固定值行距 (单位: twips, 1磅=20 twips)。

---

## 🤝 Contributing / 贡献

*   **Bug Reports**: Welcome via Issues.
*   **Pull Requests**: Please adhere to the existing code style and update the version in `package.json`.

**License**: ISC