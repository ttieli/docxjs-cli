const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');
const MarkdownIt = require('markdown-it');

// 尝试从项目根目录加载 docx，如果失败则尝试标准路径
let docx;
try {
    docx = require('../node_modules/docx');
} catch (e) {
    try {
        docx = require('docx');
    } catch (e2) {
        console.error("❌ 无法加载 'docx' 模块。请确保在项目根目录运行此脚本或已安装依赖。");
        process.exit(1);
    }
}

const { 
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    BorderStyle, HeadingLevel, AlignmentType, WidthType, VerticalAlign, 
    ExternalHyperlink, UnderlineType, ImageRun 
} = docx;

// --- 1. 简易图片尺寸解析器 (避免引入 image-size 依赖) ---
function getImageSize(buffer) {
    try {
        if (buffer.length < 24) return null;
        // PNG
        if (buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4E && buffer[3] === 0x47) {
            return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
        }
        // JPEG Scanner
        if (buffer[0] === 0xFF && buffer[1] === 0xD8) {
            let i = 2;
            while (i < buffer.length) {
                const marker = buffer.readUInt16BE(i);
                i += 2;
                if (marker >= 0xFFC0 && marker <= 0xFFCF && marker !== 0xFFC4 && marker !== 0xFFC8) {
                    return { width: buffer.readUInt16BE(i + 2), height: buffer.readUInt16BE(i) };
                }
                const length = buffer.readUInt16BE(i);
                i += length;
            }
        }
    } catch (e) { /* ignore */ }
    return null; // 无法识别或解析
}

// --- 2. 异步图片获取 ---
async function fetchImageBuffer(src, baseDir) {
    console.log(`🖼️  Fetching image: ${src}`);
    if (src.startsWith('http')) {
        return new Promise((resolve) => {
            const client = src.startsWith('https') ? https : http;
            const req = client.get(src, { timeout: 5000 }, (res) => {
                if (res.statusCode !== 200) {
                    console.warn(`   ⚠️ HTTP Error ${res.statusCode} for ${src}`);
                    resolve(null);
                    return;
                }
                const chunks = [];
                res.on('data', c => chunks.push(c));
                res.on('end', () => resolve(Buffer.concat(chunks)));
            });
            req.on('error', (e) => {
                console.warn(`   ⚠️ Network error for ${src}: ${e.message}`);
                resolve(null);
            });
            req.on('timeout', () => {
                req.destroy();
                console.warn(`   ⚠️ Timeout for ${src}`);
                resolve(null);
            });
        });
    } else {
        // 本地路径处理
        try {
            // 处理 URL 编码的路径 (比如空格被转为 %20)
            const decodedSrc = decodeURIComponent(src);
            const possiblePaths = [
                path.resolve(baseDir, decodedSrc),
                path.resolve(baseDir, 'sample data', decodedSrc) // 针对此 demo 的特殊 fallback
            ];

            for (const p of possiblePaths) {
                if (fs.existsSync(p)) {
                    return fs.readFileSync(p);
                }
            }
            console.warn(`   ⚠️ Local file not found: ${decodedSrc}`);
        } catch (e) {
            console.warn(`   ⚠️ Read file error: ${e.message}`);
        }
        return null;
    }
}

// --- 3. 异步 Inline 处理 ---
async function processInlineAsync(inlineToken, currentStyle, baseDir) {
    const runs = [];
    if (!inlineToken.children) return runs;

    let isBold = false;
    let isItalic = false;
    let inLink = false;
    let linkHref = "";
    let linkChildren = [];

    const linkConfig = { color: "0000FF", underline: true };

    for (const token of inlineToken.children) {
        if (token.type === 'image') {
            // --- 图片处理逻辑 ---
            const src = token.attrs.find(a => a[0] === 'src')[1];
            const alt = token.content;
            
            const buffer = await fetchImageBuffer(src, baseDir);
            
            if (buffer) {
                const dims = getImageSize(buffer);
                let width = 500; // 默认
                let height = 300;

                // 简单的缩放逻辑 (假设页面内容宽 ~600px / 96dpi)
                const MAX_WIDTH = 600;
                if (dims) {
                    width = dims.width;
                    height = dims.height;
                    if (width > MAX_WIDTH) {
                        const ratio = MAX_WIDTH / width;
                        width = MAX_WIDTH;
                        height = Math.round(height * ratio);
                    }
                }

                runs.push(new ImageRun({
                    data: buffer,
                    transformation: { width, height },
                    altText: { description: alt, title: alt }
                }));
            } else {
                // 图片加载失败，显示占位文本
                runs.push(new TextRun({ 
                    text: `[IMAGE LOAD FAILED: ${alt}]`, 
                    color: "FF0000",
                    bold: true 
                }));
            }
        } 
        else if (token.type === 'link_open') {
            inLink = true;
            linkHref = token.attrs ? token.attrs.find(attr => attr[0] === 'href')[1] : "";
            linkChildren = [];
        } 
        else if (token.type === 'link_close') {
            inLink = false;
            if (linkChildren.length > 0) {
                runs.push(new ExternalHyperlink({ children: linkChildren, link: linkHref }));
            }
        } 
        else {
            let run = null;
            if (token.type === 'text') {
                run = new TextRun({
                    text: token.content,
                    bold: isBold,
                    italics: isItalic,
                    font: "Times New Roman",
                    size: 24,
                    color: inLink ? linkConfig.color : "000000",
                    underline: (inLink && linkConfig.underline) ? { type: UnderlineType.SINGLE, color: linkConfig.color } : undefined
                });
            } else if (token.type === 'code_inline') {
                run = new TextRun({
                    text: token.content,
                    font: "Courier New",
                    size: 22,
                    color: "333333",
                    shading: { type: "clear", fill: "EEEEEE", color: "auto" }
                });
            } else if (token.type === 'strong_open') { isBold = true; } 
            else if (token.type === 'strong_close') { isBold = false; } 
            else if (token.type === 'em_open') { isItalic = true; } 
            else if (token.type === 'em_close') { isItalic = false; }

            if (run) {
                if (inLink) linkChildren.push(run);
                else runs.push(run);
            }
        }
    }
    return runs;
}

// --- 4. 异步 Docx 生成 (简化版) ---
async function generateDocxAsync(markdownContent, baseDir) {
    const md = new MarkdownIt();
    const tokens = md.parse(markdownContent, {});
    const docChildren = [];
    
    // 简化的样式配置
    const currentStyle = {
        fontMain: "Times New Roman",
        fontSizeMain: 24
    };

    let tableBuffer = null;

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];

        if (token.type === 'heading_open') {
            const level = parseInt(token.tag.replace('h', ''));
            const inlineToken = tokens[i + 1];
            // 标题内也可能有图片，所以也要用 processInlineAsync
            const runs = await processInlineAsync(inlineToken, currentStyle, baseDir);
            
            let headingLevel = HeadingLevel.HEADING_1;
            if (level === 2) headingLevel = HeadingLevel.HEADING_2;
            if (level >= 3) headingLevel = HeadingLevel.HEADING_3;

            docChildren.push(new Paragraph({
                heading: headingLevel,
                children: runs,
                spacing: { before: 200, after: 200 }
            }));
            i += 2;
        }
        else if (token.type === 'paragraph_open') {
            if (!tableBuffer) {
                const runs = await processInlineAsync(tokens[i + 1], currentStyle, baseDir);
                docChildren.push(new Paragraph({
                    children: runs,
                    spacing: { after: 120 }
                }));
                i += 2;
            }
        }
        else if (token.type === 'table_open') { 
            tableBuffer = { rows: [], isHeader: false }; 
        }
        else if (token.type === 'thead_open') { tableBuffer.isHeader = true; }
        else if (token.type === 'thead_close') { tableBuffer.isHeader = false; }
        else if (token.type === 'tr_open') { if (tableBuffer) tableBuffer.currentRow = []; }
        else if (token.type === 'th_open' || token.type === 'td_open') {
            if (tableBuffer && tableBuffer.currentRow) {
                tableBuffer.currentRow.push(tokens[i + 1].content);
            }
        }
        else if (token.type === 'tr_close') {
            if (tableBuffer) {
                tableBuffer.rows.push({ content: tableBuffer.currentRow, isHeader: tableBuffer.isHeader });
                tableBuffer.currentRow = null;
            }
        }
        else if (token.type === 'table_close') {
            if (tableBuffer && tableBuffer.rows.length > 0) {
                const rows = await Promise.all(tableBuffer.rows.map(async rowObj => {
                    const cells = await Promise.all(rowObj.content.map(async cellText => {
                        // 解析单元格内的 Markdown
                        const cellTokens = md.parseInline(cellText, {})[0];
                        const cellRuns = await processInlineAsync(cellTokens, currentStyle, baseDir);
                        return new TableCell({
                            children: [new Paragraph({ children: cellRuns })],
                            verticalAlign: VerticalAlign.CENTER
                        });
                    }));
                    return new TableRow({ children: cells });
                }));

                docChildren.push(new Table({
                    rows: rows,
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: {
                        top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                        bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                        left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                        right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" }
                    }
                }));
                tableBuffer = null;
            }
        }
    }

    return new Document({
        sections: [{
            children: docChildren
        }]
    });
}

// --- Main ---
(async () => {
    const inputFile = path.join(__dirname, 'sample data.md');
    const outputFile = path.join(__dirname, 'demo_output.docx');

    console.log(`🚀 Starting Demo Conversion...`);
    console.log(`📂 Input: ${inputFile}`);

    if (!fs.existsSync(inputFile)) {
        console.error(`❌ Input file not found!`);
        process.exit(1);
    }

    const mdContent = fs.readFileSync(inputFile, 'utf-8');
    
    try {
        const doc = await generateDocxAsync(mdContent, __dirname);
        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(outputFile, buffer);
        console.log(`✅ Success! Demo output created at: ${outputFile}`);
    } catch (e) {
        console.error(`❌ Conversion failed:`, e);
    }
})();
