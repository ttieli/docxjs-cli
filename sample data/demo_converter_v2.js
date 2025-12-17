const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { imageSize: sizeOf } = require('image-size');
const MarkdownIt = require('markdown-it');

// 尝试加载 docx
let docx;
try {
    docx = require('docx');
} catch (e) {
    console.error("❌ 无法加载 'docx' 模块。");
    process.exit(1);
}

const { 
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    BorderStyle, HeadingLevel, AlignmentType, WidthType, VerticalAlign, 
    ExternalHyperlink, UnderlineType, ImageRun 
} = docx;

// --- 1. 图片获取与缓存 ---
const imageCache = new Map();

async function fetchImageBuffer(src, baseDir) {
    if (imageCache.has(src)) {
        console.log(`📦 Cache hit: ${src}`);
        return imageCache.get(src);
    }

    console.log(`🖼️  Fetching (v2): ${src}`);
    let buffer = null;

    if (src.startsWith('http')) {
        try {
            const response = await axios.get(src, {
                responseType: 'arraybuffer',
                timeout: 10000, // 10秒超时
                maxRedirects: 5
            });
            buffer = Buffer.from(response.data);
        } catch (error) {
            console.warn(`   ⚠️ Download failed for ${src}: ${error.message}`);
        }
    } else {
        // 本地路径
        try {
            const decodedSrc = decodeURIComponent(src);
            const possiblePaths = [
                path.resolve(baseDir, decodedSrc),
                path.resolve(baseDir, 'sample data', decodedSrc)
            ];

            for (const p of possiblePaths) {
                if (fs.existsSync(p)) {
                    buffer = fs.readFileSync(p);
                    break;
                }
            }
            if (!buffer) console.warn(`   ⚠️ Local file not found: ${decodedSrc}`);
        } catch (e) {
            console.warn(`   ⚠️ Read file error: ${e.message}`);
        }
    }

    if (buffer) {
        imageCache.set(src, buffer);
    }
    return buffer;
}

// --- 2. 异步 Inline 处理 ---
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
            const src = token.attrs.find(a => a[0] === 'src')[1];
            const alt = token.content;
            
            const buffer = await fetchImageBuffer(src, baseDir);
            
            if (buffer) {
                try {
                    const dimensions = sizeOf(buffer);
                    let width = dimensions.width;
                    let height = dimensions.height;
                    
                    // 智能缩放
                    const MAX_WIDTH = 600; // 页面可用宽度 (假设值)
                    if (width > MAX_WIDTH) {
                        const ratio = MAX_WIDTH / width;
                        width = MAX_WIDTH;
                        height = Math.round(height * ratio);
                    }

                    runs.push(new ImageRun({
                        data: buffer,
                        transformation: { width, height },
                        altText: { description: alt, title: alt }
                    }));
                } catch (e) {
                    console.warn(`   ⚠️ Image format error: ${e.message}`);
                    runs.push(new TextRun({ text: `[Invalid Image Format]`, color: "FF0000" }));
                }
            } else {
                runs.push(new TextRun({ 
                    text: `[Image Not Found: ${alt}]`, 
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

// --- 3. 文档生成 ---
async function generateDocxAsync(markdownContent, baseDir) {
    const md = new MarkdownIt();
    const tokens = md.parse(markdownContent, {});
    const docChildren = [];
    const currentStyle = { fontMain: "Times New Roman", fontSizeMain: 24 };

    let tableBuffer = null;

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];

        if (token.type === 'heading_open') {
            const level = parseInt(token.tag.replace('h', ''));
            const inlineToken = tokens[i + 1];
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
        else if (token.type === 'table_open') { tableBuffer = { rows: [], isHeader: false }; }
        else if (token.type === 'thead_open') { tableBuffer.isHeader = true; }
        else if (token.type === 'thead_close') { tableBuffer.isHeader = false; }
        else if (token.type === 'tr_open') { if (tableBuffer) tableBuffer.currentRow = []; }
        else if (token.type === 'th_open' || token.type === 'td_open') {
            if (tableBuffer && tableBuffer.currentRow) tableBuffer.currentRow.push(tokens[i + 1].content);
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

    return new Document({ sections: [{ children: docChildren }] });
}

// --- Run ---
(async () => {
    const inputFile = path.join(__dirname, 'sample data.md');
    const outputFile = path.join(__dirname, 'demo_output_v2.docx');
    console.log(`🚀 Starting V2 Demo (Axios + Image-Size)...`);
    
    if (!fs.existsSync(inputFile)) {
        console.error("Input not found"); 
        process.exit(1);
    }

    try {
        const mdContent = fs.readFileSync(inputFile, 'utf-8');
        const doc = await generateDocxAsync(mdContent, __dirname);
        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(outputFile, buffer);
        console.log(`✅ Success! Output: ${outputFile}`);
    } catch (e) {
        console.error("Error:", e);
    }
})();
