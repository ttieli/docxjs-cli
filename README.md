# docxjs-cli

[中文说明](#中文说明) | [English](#english)

**Current Version / 当前版本**: `1.1.0`

A powerful, **hybrid CLI tool built with Node.js (`docx.js`) and Python (`python-docx`)** that converts Markdown to Word (.docx) documents. It combines the generation capabilities of `docx.js` (Node.js) with the style parsing capabilities of `python-docx` (Python) to deliver high-fidelity documents, specifically optimized for **Chinese Official Document formats (党政机关公文格式)**.

一个基于 **Node.js (`docx.js`) 和 Python (`python-docx`) 的强大混合架构命令行工具**，用于将 Markdown 转换为 Word (.docx) 文档。它结合了 `docx.js` (Node.js) 的生成能力和 `python-docx` (Python) 的样式解析能力，专为生成符合**中国党政机关公文格式**的标准文档而优化，支持交互式选择和从现有 Word 文档中提取样式。

---

<a name="english"></a>
## 🇬🇧 English Documentation

### ✨ Features

*   **Desktop Application (New!)**: A standalone Electron app for Windows & macOS. No Node/Python installation required.
*   **Web Interface**: A built-in, user-friendly Web UI for visual template selection, style editing, and file conversion.
*   **Markdown to Docx**: Robust parsing via `markdown-it` with support for bold, italic, lists, tables, and **inline code (`code`)**.
*   **Official Government Style**: Built-in `gov_official_red` template that enforces strict formatting:
    *   **Red Header (红头)**: "FangSong" and "FZXiaoBiaoSong" fonts.
    *   **Strict Margins**: Standard 3.7cm/3.5cm margins.
    *   **Solid Borders**: Tables are automatically rendered with solid black borders.
*   **Interactive Mode**: If no template is specified, a user-friendly menu helps you choose the right style.
*   **Hybrid Style Extraction**:
    *   Uses a Python helper script (`style_extractor.py`) to parse an existing `.docx` file (Reference Doc).
    *   Extracts fonts (including complex Chinese fonts), sizes, and margins to override template defaults.

### 🖥️ Desktop Application

We now provide a compiled desktop application (Windows .exe / macOS .dmg).

**Download**: Check the [GitHub Releases](https://github.com/ttieli/docxjs-cli/releases) page.

**Build from Source**:
```bash
# Install dependencies
npm install

# Build for macOS
npm run dist:mac

# Build for Windows (Requires Wine on macOS, or run on Windows)
npm run dist:win
```

### 🛠 Prerequisites (CLI Only)

This is a hybrid tool requiring both Node.js and Python environments **(Only for CLI/Web source usage. The Desktop App has no prerequisites)**.

1.  **Node.js** (v14 or higher)
2.  **Python 3.x**
3.  **Python Dependency**:
    ```bash
    pip install python-docx
    ```

### 📦 Installation / 安装

**One-line Install (Recommended) / 一键安装（推荐）:**

This will automatically set up the required Python environment and install the tool globally.
这将自动配置所需的 Python 环境并全局安装工具。

```bash
curl -sSL https://raw.githubusercontent.com/ttieli/docxjs-cli/main/install_global.sh | bash
```

**Manual Install / 手动安装:**

```bash
git clone https://github.com/ttieli/docxjs-cli.git
cd docxjs-cli
./install_global.sh
```

### 🚀 Usage

#### 1. Interactive Mode (Recommended)
Simply run without a template argument, and choose from the menu.
```bash
docxjs input.md -o output.docx
```

#### 2. Use Official Government Template
Applies the "Red Header", "FangSong" font, and standard margins.
```bash
docxjs input.md -o output_official.docx -t gov_official_red
```

#### 3. Hybrid Mode (Template + Reference Override)
Use the `gov_official_red` template as the base structure, but steal the specific fonts and margins from a real `.docx` file.

```bash
docxjs input.md -o output_hybrid.docx -t gov_official_red -r ./path/to/reference.docx
```

#### 4. Custom Template from JSON
Load user-defined templates via `--config`.

```bash
# Example: Use a template named 'tech_report_blue' defined in 'common_styles.json'
docxjs input.md -o output_custom.docx -t tech_report_blue --config ./templates/common_styles.json
```

#### 5. Web Interface
Launch the built-in web server to use the visual interface.
```bash
# After global installation:
docxjs-web

# Or from source:
npm start
```
Then access **http://localhost:3000** in your browser.

---

<a name="how-it-works"></a>
## ⚙️ Architecture & Processing Flow / 架构与流程

This tool uses a **Hybrid Node.js + Python** architecture to achieve high-fidelity document processing.

### 1. Core Dependencies / 核心依赖

| Component | Technology | Key Libraries | Purpose |
| :--- | :--- | :--- | :--- |
| **Web/CLI** | **Node.js** | `express`, `yargs` | Application entry, server, and argument parsing. |
| **Doc Generator** | **Node.js** | `markdown-it`, `docx` | Parses Markdown AST and programmatically builds `.docx` files. |
| **Doc Importer** | **Node.js** | `mammoth`, `turndown` | Converts uploaded Word docs to HTML, then to Markdown for editing. |
| **Style Engine** | **Python** | `python-docx` | Parses `.docx` binaries to extract visual styles (fonts, margins, sizes). |

### 2. Processing Flow / 处理流程

#### A. Import Flow (Word → Markdown)
1.  **Upload**: User uploads a `.docx` file via the Web UI.
2.  **Conversion**:
    *   **Content**: Server uses `mammoth.js` to convert the Docx content to HTML, then `turndown` converts it to Markdown.
    *   **Styles**: Server calls the Python script (`style_extractor.py`) to analyze the Docx and extract fonts, margins, and table styles into a JSON object.
3.  **Result**: The user gets editable Markdown in the editor, and the "Reference Doc" is automatically set to preserve the original styles.

#### B. Export Flow (Markdown → Docx)
1.  **Input**: User submits Markdown content + a Template Name (or Reference Doc).
2.  **Style Merging**:
    *   Base styles are loaded from the selected Template (e.g., `gov_official_red`).
    *   If a Reference Doc is present, its extracted styles override the template defaults.
3.  **Generation**: `docx.js` builds a brand new `.docx` file, applying the merged styles to the Markdown content.

---

<a name="中文说明"></a>
## 🇨🇳 中文说明

### ✨ 核心特性

*   **桌面客户端 (新增!)**：支持 Windows 和 macOS 的独立 Electron 应用。无需安装 Node/Python 环境，双击即用。
*   **Web 可视化界面**：内置好用的 Web UI，支持可视化选择模板、微调样式和文件转换。
*   **Markdown 转 Docx**：基于 `markdown-it` 的稳定解析，完美支持表格加粗、斜体等内联样式，以及**行内代码 (`code`)**。
*   **党政机关公文标准**：内置 `gov_official_red` (红头公文) 模板，严格遵循国家公文格式标准：
    *   **红头文件**：自动应用方正小标宋（红头）、仿宋（正文）、黑体/楷体（标题）。
    *   **版面设置**：严格的 上3.7cm / 下3.5cm / 左2.8cm / 右2.6cm 页边距。
    *   **公文表格**：自动将 Markdown 表格渲染为全黑色实线边框（解决 Pandoc 表格样式不可控问题）。
*   **交互式选择**：如果不指定模板参数，工具会自动弹出中文菜单供您选择。
*   **混合样式提取 (Node.js + Python)**：
    *   利用 Python 脚本 (`style_extractor.py`) 解析现有的 `.docx` 参考文档。
    *   智能提取正文字体（如“宋体”）、字号和页边距，并覆盖预设模板。

### 🖥️ 桌面客户端

我们现在提供编译好的桌面安装包 (Windows .exe / macOS .dmg)。

**下载地址**: 请访问 [GitHub Releases](https://github.com/ttieli/docxjs-cli/releases) 页面。

**源码构建**:
```bash
# 安装依赖
npm install

# 构建 macOS 版本
npm run dist:mac

# 构建 Windows 版本 (macOS 上需要 Wine，或者直接在 Windows 上运行)
npm run dist:win
```

### 🛠 前置要求 (仅限 CLI/Web 源码模式)

本工具采用 Node.js + Python 混合架构，以实现最佳的生成与解析能力。**(使用桌面客户端无需这些前置要求)**

1.  **Node.js** (v14 以上)
2.  **Python 3.x**
3.  **Python 依赖库**：
    ```bash
    pip install python-docx
    ```

### 📦 安装方法

直接通过 GitHub 仓库地址进行全局安装：

```bash
# 请将 URL 替换为您的实际 GitHub 仓库地址
npm install -g git+https://github.com/YOUR_USERNAME/docxjs-cli.git
```

### 🚀 使用指南

#### 1. 交互式模式 (推荐)
不带模板参数运行，通过键盘选择。
```bash
docxjs input.md -o output.docx
```

#### 2. 指定模板 (公文红头)
```bash
docxjs input.md -o output_official.docx -t gov_official_red
```
*内置模板包括*：`gov_official_red` (红头), `gov_notice_plain` (普通通知), `business_contract` (商务合同), `default` (默认)。

#### 3. 混合模式 (模板 + 样式吸取)
以 `gov_official_red` 为底座（保持红头结构、表格实线），但从指定的真实 Word 文档中“吸取”字体和页边距。

```bash
docxjs input.md -o output_hybrid.docx -t gov_official_red -r ./path/to/reference.docx
```

#### 4. 自定义 JSON 模板
通过 `--config` 加载您自定义的 JSON 样式文件。

```bash
# 示例：使用 templates/common_styles.json 中的 'tech_report_blue' 模板
docxjs input.md -o output_custom.docx -t tech_report_blue --config ./templates/common_styles.json
```

#### 5. Web 可视化界面
启动内置的 Web 服务器以使用可视化界面。
```bash
# 全局安装后：
docxjs-web

# 或从源码运行：
npm start
```
启动后访问浏览器 **http://localhost:3000**。

### ⚙️ Template Configuration (JSON)

Templates are defined in a JSON file. The built-in templates are in `templates/templates.json`.

| Property       | Type     | Description                                                          | Example Value                     |
| :------------- | :------- | :------------------------------------------------------------------- | :-------------------------------- |
| `fontMain`     | `string` | Main font for body text.                                             | `"FangSong_GB2312"`               |
| `colorMain`    | `string` | Main text color (Hex without #).                                     | `"333333"`                        |
| `fontHeader1`  | `string` | Font for Heading 1.                                                  | `"FZXiaoBiaoSong-B05S"`           |
| `colorHeader1` | `string` | Color for Heading 1.                                                 | `"FF0000"`                        |
| `fontSizeMain` | `number` | Font size for main text (in half-points; 32 = 16pt).                 | `32`                              |
| `lineSpacing`  | `number` | Line spacing for paragraphs (in twips; 560 = 28pt).                  | `560`                             |
| `margin`       | `object` | Page margins. Values can be numbers (in twips) or string (e.g., "3.7cm"). | `{ "top": "3.7cm", "bottom": "3.5cm", "left": "2.8cm", "right": "2.6cm" }` |
| `redHeader`    | `boolean`| If `true`, Heading 1 will be red (for official documents).          | `true`                            |
| `table`        | `object` | **New**: Table styling configuration.                                | See example below.                |

---

## 🤝 Contributing / 贡献

*   **Bug Reports**: Please submit an issue.
*   **Pull Requests**: Welcome! Please ensure you update the version number in `package.json` for any code changes.

**License**: ISC
