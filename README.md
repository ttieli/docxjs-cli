# docxjs-cli

[中文说明](#中文说明) | [English](#english)

**Current Version / 当前版本**: `1.1.0`

A powerful, **hybrid CLI tool built with Node.js (`docx.js`) and Python (`python-docx`)** that converts Markdown to Word (.docx) documents. It combines the generation capabilities of `docx.js` (Node.js) with the style parsing capabilities of `python-docx` (Python) to deliver high-fidelity documents, specifically optimized for **Chinese Official Document formats (党政机关公文格式)**.

一个基于 **Node.js (`docx.js`) 和 Python (`python-docx`) 的强大混合架构命令行工具**，用于将 Markdown 转换为 Word (.docx) 文档。它结合了 `docx.js` (Node.js) 的生成能力和 `python-docx` (Python) 的样式解析能力，专为生成符合**中国党政机关公文格式**的标准文档而优化，支持交互式选择和从现有 Word 文档中提取样式。

---

<a name="english"></a>
## 🇬🇧 English Documentation

### ✨ Features

*   **Markdown to Docx**: Robust parsing via `markdown-it` with support for bold, italic, lists, and tables.
*   **Official Government Style**: Built-in `gov_official_red` template that enforces strict formatting:
    *   **Red Header (红头)**: "FangSong" and "FZXiaoBiaoSong" fonts.
    *   **Strict Margins**: Standard 3.7cm/3.5cm margins.
    *   **Solid Borders**: Tables are automatically rendered with solid black borders.
*   **Interactive Mode**: If no template is specified, a user-friendly menu helps you choose the right style.
*   **Hybrid Style Extraction**:
    *   Uses a Python helper script (`style_extractor.py`) to parse an existing `.docx` file (Reference Doc).
    *   Extracts fonts (including complex Chinese fonts), sizes, and margins to override template defaults.

### 🛠 Prerequisites

This is a hybrid tool requiring both Node.js and Python environments.

1.  **Node.js** (v14 or higher)
2.  **Python 3.x**
3.  **Python Dependency**:
    ```bash
    pip install python-docx
    ```

### 📦 Installation

You can install this tool directly from GitHub:

```bash
# Replace with your actual GitHub repo URL
npm install -g git+https://github.com/YOUR_USERNAME/docxjs-cli.git
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

---

<a name="中文说明"></a>
## 🇨🇳 中文说明

### ✨ 核心特性

*   **Markdown 转 Docx**：基于 `markdown-it` 的稳定解析，完美支持表格加粗、斜体等内联样式。
*   **党政机关公文标准**：内置 `gov_official_red` (红头公文) 模板，严格遵循国家公文格式标准：
    *   **红头文件**：自动应用方正小标宋（红头）、仿宋（正文）、黑体/楷体（标题）。
    *   **版面设置**：严格的 上3.7cm / 下3.5cm / 左2.8cm / 右2.6cm 页边距。
    *   **公文表格**：自动将 Markdown 表格渲染为全黑色实线边框（解决 Pandoc 表格样式不可控问题）。
*   **交互式选择**：如果不指定模板参数，工具会自动弹出中文菜单供您选择。
*   **混合样式提取 (Hybrid Mode)**：
    *   利用 Python 脚本 (`style_extractor.py`) 解析现有的 `.docx` 参考文档。
    *   智能提取正文字体（如“宋体”）、字号和页边距，并覆盖预设模板。

### 🆚 为什么选择 docxjs-cli 而不是 Pandoc？

| 功能特性 | Pandoc | docxjs-cli |
| :--- | :--- | :--- |
| **表格样式控制** | ❌ **难以控制**。默认使用 Word 表格样式，难以强制指定边框（如全黑实线网格）。 | ✅ **精确控制**。代码级控制表格渲染，可强制应用公文要求的全黑实线边框、特定列宽和对齐方式。 |
| **党政机关公文格式** | ❌ **配置复杂**。需要制作非常标准的 `reference.docx`，且必须手动修改内部 XML 样式名。 | ✅ **开箱即用**。内置 `gov_official_red` 模板，硬编码实现了红头、仿宋字体、严格页边距和行距。 |
| **参考文档兼容性** | ⚠️ **挑剔**。要求参考文档必须是“干净”的标准 Docx。 | ✅ **宽容灵活**。利用 Python 脚本“吸取”文档的视觉属性（字体、字号、边距），即使文档样式命名不规范也能工作。 |
| **交互体验** | ❌ **无**。纯命令行参数。 | ✅ **友好**。提供交互式菜单选择模板。 |

### 🛠 前置要求

本工具采用 Node.js + Python 混合架构，以实现最佳的生成与解析能力。

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