# DocxJS 功能增强提案：图片支持与后续规划

本提案旨在分阶段扩展 `docxjs` 的功能。**本期重点实现 Markdown 图片的正确转换与插入**，并将 Mermaid 图表支持及其他高级功能列入后续规划。

## 1. 本期目标：图片支持 (Image Support)

当前 `docxjs` 忽略了 Markdown 中的 `image` token。本期修改将集中在 `lib/core.js`，使其能够解析、获取并正确渲染图片。

### 1.1 核心逻辑修改与可行性审查 (Core Logic & Feasibility Review)

**处理流程 (Processing Flow)**:
- Markdown-It 将 `image` token 嵌套在 inline token 中。`processInline` 函数需要处理这些图片。
- **设计决策**: 为了处理网络图片的异步下载，`processInline` 函数将设计为 `async`。这与 `generateDocx` 函数已有的异步特性相符，是 Node.js 中处理 I/O 的标准且简洁的方式。每个图片下载将在 `processInline` 内部按顺序处理。

**路径处理与缓存 (Path Resolution & Caching)**:
- **本地路径**: 必须指定“基准目录”（通常是 Markdown 文件所在目录）。代码需处理相对路径，并增加文件存在性检查，避免静默失败。**这意味着来自项目中的任何本地目录的图片，例如 `release_assets`，只要路径正确，都将被支持。**
- **缓存**: 对相同的图片路径/URL 实施内存缓存，避免重复下载或读取，提高转换效率。

**临时文件管理 (Temporary File Management & App Support)**:
- **默认下载路径**: 对于下载的远程图片，将使用操作系统默认的临时目录（`os.tmpdir()`）作为存储位置，确保在 CLI 和打包后的 Electron App 环境中均有写入权限。
- **自动清理**: 转换流程结束后，将自动触发清理机制，删除本次会话产生的临时图片文件及目录，防止磁盘空间占用。

**网络图片增强 (Robust Remote Fetch)**:
- **安全与效率**: 使用 Node.js 原生 `https`/`http` 模块进行下载时，将为每个图片下载设置一个明确的**最长等待上限（Timeout）**。同时，还需处理重定向、大小上限，并确保在下载失败时提供清晰的错误信息。

**尺寸与比例控制 (Sizing & Aspect Ratio)**:
- **问题**: `ImageRun` 需要明确的宽/高像素值。固定值会导致变形。
- **方案**: 引入 `image-size` 库从 Buffer 中读取原始尺寸。根据文档页面的有效宽度（Content Width）计算缩放比例，设定最大宽度限制，自动推算高度以保持纵横比。

### 1.2 实现细节 (lib/core.js)

```javascript
// 伪代码示例
const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { ImageRun } = require('docx');
const sizeOf = require('image-size'); // 需要安装或实现简单的尺寸读取

const imageCache = new Map(); // 内存缓存

// 临时目录管理
const tempDir = path.join(os.tmpdir(), 'docxjs-images');
if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });

// 1. 异步获取 Buffer (带缓存与安全控制)
async function fetchImageBuffer(src, baseDir) {
    if (imageCache.has(src)) return imageCache.get(src);

    let buffer = null;
    if (src.startsWith('http')) {
        // ... HTTP GET implementation with timeout, redirect limit, size cap ...
        // 下载到 tempDir 下的临时文件，读取 Buffer 后可选择删除或保留至会话结束
    } else {
        // Resolve relative path
        const localPath = path.resolve(baseDir, src);
        if (fs.existsSync(localPath)) {
            buffer = fs.readFileSync(localPath);
        } else {
            console.warn(`Image not found: ${localPath}`);
        }
    }
    
    if (buffer) imageCache.set(src, buffer);
    return buffer;
}

// 2. 清理逻辑
function cleanupTempFiles() {
    try {
        fs.rmSync(tempDir, { recursive: true, force: true });
    } catch (e) {
        console.warn('Failed to cleanup temp files:', e.message);
    }
}

// 3. 转换为 ImageRun
function createImageRun(buffer, maxWidth, altText) {
    if (!buffer) {
        // Fallback text if image load failed
        return new TextRun({ text: `[Image: ${altText || "Load Failed"}]`, color: "FF0000" });
    }
    
    try {
        const dimensions = sizeOf(buffer);
        let width = dimensions.width;
        let height = dimensions.height;

        // 缩放逻辑
        if (width > maxWidth) {
            const ratio = maxWidth / width;
            width = maxWidth;
            height = height * ratio;
        }

        return new ImageRun({
            data: buffer,
            transformation: { width, height },
            altText: {
                description: altText || "",
                title: altText || ""
            }
        });
    } catch (e) {
        return new TextRun({ text: `[Image Error: ${e.message}]`, color: "FF0000" });
    }
}
```

### 1.3 依赖变更
- 需要添加/确认依赖：`image-size` (用于获取尺寸)。
- 可选依赖：`axios` (如果希望简化 HTTP 请求代码)。

---

## 2. 后期规划 (Future Scope: "后期再说")

以下功能暂不在此次更新中实现，留待后续版本开发。

### 2.1 Mermaid 图表支持
- **目标**: 将 ````mermaid` 代码块渲染为图片。
- **挑战**: 需要依赖在线 API (mermaid.ink) 或本地 Headless Browser (Puppeteer)。前者有隐私/网络限制，后者体积巨大。
- **计划**: 在解决图片插入基础功能后，评估最佳渲染方案（可能作为插件或可选特性）。

### 2.2 高级代码块样式
- **目标**: 为普通代码块提供类似 IDE 的背景填充、语法高亮（可选）和等宽字体。
- **细节**: 需要处理 `ShadingType` 和列表嵌套中的代码块缩进问题。

### 2.3 健壮性与测试
- **全面测试**: 覆盖各种边缘情况（图片 404、超大图片、不支持的格式）。
- **离线模式**: 为受限网络环境提供纯本地模式。

## 3. 总结
本期专注于打通“图片插入”这一核心痛点，确保文档转换的基本完整性。Mermaid 等复杂图表功能将在图片基础架构稳固后实施。
