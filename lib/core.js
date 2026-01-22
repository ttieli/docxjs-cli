const MarkdownIt = require('markdown-it');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const os = require('os');
const { imageSize: sizeOf } = require('image-size');
let chromium;
try {
    ({ chromium } = require('playwright'));
} catch (e) {}

// Markdown-it plugins for extended syntax
const markdownItSub = require('markdown-it-sub');
const markdownItSup = require('markdown-it-sup');
const markdownItFootnote = require('markdown-it-footnote');
const markdownItTaskLists = require('markdown-it-task-lists');
const markdownItTexmath = require('markdown-it-texmath');

const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    BorderStyle, HeadingLevel, AlignmentType, WidthType, VerticalAlign, ShadingType,
    ExternalHyperlink, UnderlineType, ImageRun, LineRuleType, FootnoteReferenceRun
} = require('docx');

// --- Image Handling & Caching ---
const imageCache = new Map();
const tempDir = path.join(os.tmpdir(), 'docxjs-images');
if (!fs.existsSync(tempDir)) {
    try { fs.mkdirSync(tempDir, { recursive: true }); } catch (e) {}
}

// --- Shared Browser Instance for Performance ---
// Reuse browser instance across multiple mermaid/math renders
let sharedBrowser = null;
let sharedMermaidPage = null;
let sharedMathPage = null;
let browserCloseTimeout = null;

async function getSharedBrowser() {
    if (!chromium) return null;
    if (!sharedBrowser || !sharedBrowser.isConnected()) {
        sharedBrowser = await chromium.launch({ headless: true });
    }
    // Reset close timeout whenever browser is used
    if (browserCloseTimeout) {
        clearTimeout(browserCloseTimeout);
    }
    // Auto-close browser after 30 seconds of inactivity
    browserCloseTimeout = setTimeout(async () => {
        await closeSharedBrowser();
    }, 30000);
    return sharedBrowser;
}

async function closeSharedBrowser() {
    if (browserCloseTimeout) {
        clearTimeout(browserCloseTimeout);
        browserCloseTimeout = null;
    }
    if (sharedMermaidPage) {
        try { await sharedMermaidPage.close(); } catch (e) {}
        sharedMermaidPage = null;
    }
    if (sharedMathPage) {
        try { await sharedMathPage.close(); } catch (e) {}
        sharedMathPage = null;
    }
    if (sharedBrowser) {
        try { await sharedBrowser.close(); } catch (e) {}
        sharedBrowser = null;
    }
}

// Mermaid diagram cache
const mermaidCache = new Map();

async function renderMermaidLocally(mermaidCode) {
    if (!chromium) {
        console.warn("   ⚠️ Playwright not found, skipping local Mermaid render.");
        return null;
    }

    // Check cache first
    const cacheKey = mermaidCode.trim();
    if (mermaidCache.has(cacheKey)) {
        return mermaidCache.get(cacheKey);
    }

    try {
        const browser = await getSharedBrowser();
        if (!browser) return null;

        // Create or reuse mermaid page
        if (!sharedMermaidPage || sharedMermaidPage.isClosed()) {
            const context = await browser.newContext({
                viewport: { width: 2400, height: 2000 },
                deviceScaleFactor: 3
            });
            sharedMermaidPage = await context.newPage();

            const mermaidPath = path.join(__dirname, '../public/vendor/mermaid/mermaid.min.js');
            let mermaidJs = '';
            if (fs.existsSync(mermaidPath)) {
                mermaidJs = fs.readFileSync(mermaidPath, 'utf8');
            } else {
                throw new Error(`Local mermaid library not found at ${mermaidPath}`);
            }

            const htmlContent = `
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body { margin: 0; background: white; overflow: hidden; }
                    * { font-family: 'Arial', sans-serif !important; }
                </style>
                <script>${mermaidJs}</script>
            </head>
            <body>
                <div id="graphDiv"></div>
            </body>
            </html>`;

            await sharedMermaidPage.setContent(htmlContent);
        }

        // Generate unique ID for this render
        const renderId = `mermaid-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

        const renderResult = await sharedMermaidPage.evaluate(async ({ code, id }) => {
            try {
                mermaid.initialize({
                    startOnLoad: false,
                    theme: 'default',
                    quadrantChart: { chartWidth: 1000 },
                    gantt: { useWidth: 1200 }
                });

                const { svg } = await mermaid.render(id, code);
                const container = document.getElementById('graphDiv');
                container.innerHTML = svg;

                const svgElement = container.querySelector('svg');
                const rect = svgElement.getBoundingClientRect();

                return { success: true, width: rect.width, height: rect.height };
            } catch (err) {
                return { success: false, message: err.message };
            }
        }, { code: mermaidCode, id: renderId });

        if (!renderResult.success) {
            console.warn(`   ⚠️ Mermaid internal render error: ${renderResult.message}`);
            return null;
        }

        const element = await sharedMermaidPage.$('#graphDiv svg');
        if (!element) return null;

        const buffer = await element.screenshot({ type: 'png', omitBackground: true });
        mermaidCache.set(cacheKey, buffer);
        return buffer;

    } catch (err) {
        console.warn(`   ⚠️ Local Mermaid render failed: ${err.message}`);
        return null;
    }
}

// Math formula cache to avoid re-rendering same formulas
const mathCache = new Map();

async function renderMathLocally(mathCode, displayMode = false) {
    if (!chromium) {
        console.warn("   ⚠️ Playwright not found, skipping local math render.");
        return null;
    }

    // Check cache first
    const cacheKey = `${displayMode ? 'block' : 'inline'}:${mathCode}`;
    if (mathCache.has(cacheKey)) {
        return mathCache.get(cacheKey);
    }

    try {
        const browser = await getSharedBrowser();
        if (!browser) return null;

        // Create or reuse math page
        if (!sharedMathPage || sharedMathPage.isClosed()) {
            const context = await browser.newContext({
                viewport: { width: 1200, height: 800 },
                deviceScaleFactor: 3
            });
            sharedMathPage = await context.newPage();

            const htmlContent = `
            <!DOCTYPE html>
            <html>
            <head>
                <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css">
                <script src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js"></script>
                <style>
                    body { margin: 0; padding: 10px; background: white; display: inline-block; }
                    #math-container { display: inline-block; }
                </style>
            </head>
            <body>
                <div id="math-container"></div>
            </body>
            </html>`;

            await sharedMathPage.setContent(htmlContent);
            await sharedMathPage.waitForFunction(() => typeof katex !== 'undefined', { timeout: 10000 });
        }

        // Render the math using KaTeX
        const renderResult = await sharedMathPage.evaluate(({ code, display }) => {
            try {
                const container = document.getElementById('math-container');
                container.style.fontSize = display ? '18px' : '14px';
                katex.render(code, container, {
                    throwOnError: false,
                    displayMode: display
                });
                const rect = container.getBoundingClientRect();
                return { success: true, width: rect.width, height: rect.height };
            } catch (err) {
                return { success: false, message: err.message };
            }
        }, { code: mathCode, display: displayMode });

        if (!renderResult.success) {
            console.warn(`   ⚠️ KaTeX render error: ${renderResult.message}`);
            return null;
        }

        const element = await sharedMathPage.$('#math-container');
        if (!element) return null;

        const buffer = await element.screenshot({ type: 'png', omitBackground: true });
        mathCache.set(cacheKey, buffer);
        return buffer;

    } catch (err) {
        console.warn(`   ⚠️ Local math render failed: ${err.message}`);
        return null;
    }
}

async function fetchImageBuffer(src, baseDir) {
    if (imageCache.has(src)) return imageCache.get(src);

    let buffer = null;
    const decodedSrc = decodeURIComponent(src);

    // Handle data URI (base64 embedded images)
    if (src.startsWith('data:')) {
        try {
            // Extract base64 data from data URI
            // Format: data:image/png;base64,<base64data>
            const matches = src.match(/^data:([^;]+);base64,(.+)$/);
            if (matches && matches[2]) {
                buffer = Buffer.from(matches[2], 'base64');
            } else {
                console.warn(`   ⚠️ Invalid data URI format`);
            }
        } catch (error) {
            console.warn(`   ⚠️ Data URI parsing failed: ${error.message}`);
        }
    } else if (src.startsWith('http')) {
        try {
            const response = await axios.get(src, {
                responseType: 'arraybuffer',
                timeout: 10000,
                maxRedirects: 5
            });
            buffer = Buffer.from(response.data);
        } catch (error) {
            console.warn(`   ⚠️ Download failed for ${src}: ${error.message}`);
        }
    } else {
        // Local file
        try {
            // Check baseDir, then current dir, then 'sample data' fallback (legacy support)
            const possiblePaths = [
                path.resolve(baseDir || process.cwd(), decodedSrc),
                path.resolve(process.cwd(), decodedSrc)
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

function getBorderStyle(styleName) {
    switch (styleName) {
        case 'single': return BorderStyle.SINGLE;
        case 'dotted': return BorderStyle.DOTTED;
        case 'dashed': return BorderStyle.DASHED;
        case 'double': return BorderStyle.DOUBLE;
        case 'none': return BorderStyle.NONE;
        default: return BorderStyle.SINGLE;
    }
}

function getAlignment(alignStr) {
    if (!alignStr) return null;
    switch (alignStr.toLowerCase()) {
        case 'center': return AlignmentType.CENTER;
        case 'right': return AlignmentType.RIGHT;
        case 'justified': return AlignmentType.JUSTIFIED;
        case 'left': return AlignmentType.LEFT;
        default: return AlignmentType.LEFT;
    }
}

async function processInline(md, inlineToken, currentStyle, baseDir, colorOverride, boldOverride) {
    const runs = [];
    if (!inlineToken.children) return runs;

    let isBold = false;
    let isItalic = false;
    let isStrike = false;
    let isSub = false;
    let isSup = false;

    let inLink = false;
    let linkHref = "";
    let linkChildren = [];

    const linkConfig = currentStyle.hyperlink || { color: "0000FF", underline: true };

    // Need to iterate sequentially for async operations
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

                    // Simple scaling logic (assuming A4 page ~600px usable width)
                    const MAX_WIDTH = 600;
                    if (width > MAX_WIDTH) {
                        const ratio = MAX_WIDTH / width;
                        width = MAX_WIDTH;
                        height = Math.round(height * ratio);
                    }

                    const imgRun = new ImageRun({
                        data: buffer,
                        transformation: { width, height },
                        altText: { description: alt, title: alt }
                    });

                    if (inLink) linkChildren.push(imgRun); // Images inside links? Uncommon but possible
                    else runs.push(imgRun);

                } catch (e) {
                    console.warn(`   ⚠️ Image format error: ${e.message}`);
                    runs.push(new TextRun({ text: `[Invalid Image]`, color: "FF0000" }));
                }
            } else {
                runs.push(new TextRun({ text: `[Image Not Found: ${alt}]`, color: "FF0000" }));
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
        // Strikethrough handling
        else if (token.type === 's_open') { isStrike = true; }
        else if (token.type === 's_close') { isStrike = false; }
        // Subscript handling
        else if (token.type === 'sub_open') { isSub = true; }
        else if (token.type === 'sub_close') { isSub = false; }
        // Superscript handling
        else if (token.type === 'sup_open') { isSup = true; }
        else if (token.type === 'sup_close') { isSup = false; }
        // Footnote reference handling [^1]
        else if (token.type === 'footnote_ref') {
            // Get footnote ID from token
            const footnoteId = token.meta ? token.meta.id + 1 : 1;
            const footnoteRun = new TextRun({
                text: `[${footnoteId}]`,
                superScript: true,
                font: currentStyle.fontMain,
                size: currentStyle.fontSizeMain,
                color: "0000FF"
            });
            if (inLink) linkChildren.push(footnoteRun);
            else runs.push(footnoteRun);
        }
        // Inline math handling ($...$)
        else if (token.type === 'math_inline') {
            const mathCode = token.content;
            try {
                const buffer = await renderMathLocally(mathCode, false); // displayMode = false for inline
                if (buffer) {
                    const dimensions = sizeOf(buffer);
                    let width = dimensions.width;
                    let height = dimensions.height;
                    // Scale down for inline (max height around text line height)
                    const MAX_HEIGHT = 30;
                    if (height > MAX_HEIGHT) {
                        const ratio = MAX_HEIGHT / height;
                        width = Math.round(width * ratio);
                        height = MAX_HEIGHT;
                    }
                    const mathRun = new ImageRun({
                        data: buffer,
                        transformation: { width, height }
                    });
                    if (inLink) linkChildren.push(mathRun);
                    else runs.push(mathRun);
                } else {
                    // Fallback: show raw math as text
                    const fallback = new TextRun({
                        text: `$${mathCode}$`,
                        font: "Courier New",
                        size: currentStyle.fontSizeMain,
                        color: "666666"
                    });
                    if (inLink) linkChildren.push(fallback);
                    else runs.push(fallback);
                }
            } catch (e) {
                console.warn(`   ⚠️ Inline math render failed: ${e.message}`);
                const errorRun = new TextRun({
                    text: `[Math Error]`,
                    color: "FF0000"
                });
                if (inLink) linkChildren.push(errorRun);
                else runs.push(errorRun);
            }
        }
        else {
            let run = null;
            if (token.type === 'text') {
                run = new TextRun({
                    text: token.content,
                    bold: (boldOverride === true) || isBold,
                    italics: isItalic,
                    strike: isStrike,
                    subScript: isSub,
                    superScript: isSup,
                    font: currentStyle.fontMain,
                    size: currentStyle.fontSizeMain,
                    color: inLink ? linkConfig.color : (colorOverride || currentStyle.colorMain || "000000"),
                    underline: (inLink && linkConfig.underline) ? { type: UnderlineType.SINGLE, color: linkConfig.color } : undefined
                });
            } else if (token.type === 'code_inline') {
                run = new TextRun({
                    text: token.content,
                    font: "Courier New",
                    size: currentStyle.fontSizeMain,
                    color: inLink ? linkConfig.color : (colorOverride || currentStyle.colorMain || "000000"),
                    shading: { type: ShadingType.CLEAR, color: "auto", fill: "F2F2F2" },
                    underline: (inLink && linkConfig.underline) ? { type: UnderlineType.SINGLE, color: linkConfig.color } : undefined
                });
            } else if (token.type === 'strong_open') { isBold = true; }
            else if (token.type === 'strong_close') { isBold = false; }
            else if (token.type === 'em_open') { isItalic = true; }
            else if (token.type === 'em_close') { isItalic = false; }
            // Softbreak handling (line break within paragraph)
            else if (token.type === 'softbreak') {
                run = new TextRun({ text: ' ' });
            }

            if (run) {
                if (inLink) linkChildren.push(run);
                else runs.push(run);
            }
        }
    }
    return runs;
}

async function generateDocx(markdownContent, currentStyle, baseDir) {
    // baseDir default to cwd if not provided
    if (!baseDir) baseDir = process.cwd();

    // Initialize markdown-it with plugins for extended syntax
    const md = new MarkdownIt({ html: false, linkify: true })
        .use(markdownItSub)         // ~subscript~
        .use(markdownItSup)         // ^superscript^
        .use(markdownItFootnote)    // [^1] footnotes
        .use(markdownItTaskLists, { enabled: true, label: true }) // - [ ] task lists
        .use(markdownItTexmath, {   // $...$ and $$...$$ math formulas
            engine: null,           // We'll render math ourselves
            delimiters: 'dollars',
            katexOptions: { throwOnError: false }
        });

    // Enable strikethrough (GFM style ~~text~~)
    md.enable('strikethrough');

    const tokens = md.parse(markdownContent, {});
    const docChildren = [];
    let tableBuffer = null;

    // List state management for nested lists
    const listStack = []; // Track nested list types: { type: 'bullet'|'ordered', level: number, counter: number }
    let inListItem = false;
    let currentListLevel = 0;

    // Blockquote state
    let blockquoteLevel = 0;

    // Footnotes collection
    const footnotes = new Map();

    const headerSpacing = currentStyle.headerSpacing || { before: 200, after: 200 };
    const bodySpacing = currentStyle.bodySpacing || { before: 0, after: 0 };

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        console.log(`Token ${i}: ${token.type} (${token.tag})`);

        if (token.type === 'heading_open') {
            const level = parseInt(token.tag.replace('h', ''));
            const inlineToken = tokens[i + 1];
            // Process inline content (text, images, etc.) for header
            // Note: Header images might be weird but supported
            const contentRuns = await processInline(md, inlineToken, currentStyle, baseDir);
            
            let paraObj = { children: [], spacing: { ...headerSpacing } };
            // Need to apply header styling to runs if they are text
            // But processInline already applies styling based on currentStyle.
            // We might need to override props for header...
            // Refactoring: The old code constructed TextRun manually for headers.
            // The new processInline returns Runs. 
            // We should modify the runs returned by processInline to match header style if possible,
            // OR pass specific styles to processInline.
            
            // Let's pass override styles to processInline? 
            // processInline supports colorOverride and boldOverride.
            // But font size?
            // Simpler approach: construct runs manually for TEXT as before, but use processInline for MIXED content?
            // The old code assumed textContent = inlineToken.content.
            // If we use processInline, we get formatted runs.
            
            // Re-implementing Header Logic with processInline is tricky because we want specific Fonts/Sizes for headers.
            // Let's stick to the old logic for Headers BUT allow images?
            // If we use processInline, we lose the Header Font/Size logic unless we pass it.
            // Let's modify processInline signature or logic? No, too complex.
            
            // Hybrid approach:
            // Headers usually don't contain images in this specific "Official Document" context.
            // But for full Markdown support, they might.
            // Let's use the old logic for now to ensure Font/Size correctness, 
            // effectively ignoring images in Headers for this iteration, OR
            // we manually build the TextRun for text content.
            
            const textContent = inlineToken.content; // Plain text content
            
            if (level === 1) {
                paraObj.heading = HeadingLevel.HEADING_1;
                
                // Logic for alignment: explicitly set > redHeader default > left default
                if (currentStyle.titleAlignment) {
                    paraObj.alignment = getAlignment(currentStyle.titleAlignment);
                } else {
                    paraObj.alignment = currentStyle.redHeader ? AlignmentType.CENTER : AlignmentType.LEFT;
                }

                let h1Color = currentStyle.colorHeader1 ? currentStyle.colorHeader1 : (currentStyle.redHeader ? "FF0000" : "000000");
                let h1Bold = true; 

                paraObj.children.push(new TextRun({
                    text: textContent,
                    font: currentStyle.fontHeader1 || currentStyle.fontMain,
                    size: currentStyle.fontSizeH1 || 32,
                    color: h1Color,
                    bold: h1Bold
                }));
                if (currentStyle.redHeader) paraObj.spacing = { after: 400 };
            } else if (level === 2) {
                 paraObj.heading = HeadingLevel.HEADING_2;
                 paraObj.children.push(new TextRun({
                    text: textContent,
                    font: currentStyle.fontHeader2 || currentStyle.fontMain,
                    size: currentStyle.fontSizeH2 || 28,
                    color: currentStyle.colorHeader2 || "000000",
                    bold: true 
                }));
            } else if (level === 3) {
                 paraObj.heading = HeadingLevel.HEADING_3;
                 paraObj.children.push(new TextRun({
                    text: textContent,
                    font: currentStyle.fontHeader3 || currentStyle.fontMain,
                    size: currentStyle.fontSizeH3 || 24,
                    color: currentStyle.colorHeader3 || "000000",
                    bold: true
                }));
            } else if (level === 4) {
                 paraObj.heading = HeadingLevel.HEADING_4;
                 paraObj.children.push(new TextRun({
                    text: textContent,
                    font: currentStyle.fontHeader4 || currentStyle.fontMain,
                    size: currentStyle.fontSizeH4 || 24, // Default to same as H3 or body
                    color: currentStyle.colorHeader4 || "000000",
                    bold: true
                }));
            } else if (level === 5) {
                 paraObj.heading = HeadingLevel.HEADING_5;
                 paraObj.children.push(new TextRun({
                    text: textContent,
                    font: currentStyle.fontHeader5 || currentStyle.fontMain,
                    size: currentStyle.fontSizeH5 || currentStyle.fontSizeMain,
                    color: currentStyle.colorHeader5 || "000000",
                    bold: true,
                    italics: true
                }));
            } else if (level === 6) {
                 paraObj.heading = HeadingLevel.HEADING_6;
                 paraObj.children.push(new TextRun({
                    text: textContent,
                    font: currentStyle.fontHeader6 || currentStyle.fontMain,
                    size: currentStyle.fontSizeH6 || currentStyle.fontSizeMain,
                    color: currentStyle.colorHeader6 || "000000",
                    bold: false,
                    italics: true
                }));
            } else {
                // Fallback for H7+ (treated as body bold)
                paraObj.children.push(new TextRun({ 
                    text: textContent, 
                    font: currentStyle.fontMain, 
                    size: currentStyle.fontSizeMain,
                    color: currentStyle.colorMain || "000000",
                    bold: true 
                }));
            }
            docChildren.push(new Paragraph(paraObj));
            console.log(`  -> Added Heading Level ${level}`);
            i += 2; 
        }
        else if (token.type === 'paragraph_open') {
            if (!tableBuffer) {
                const runs = await processInline(md, tokens[i + 1], currentStyle, baseDir);
                let paraConfig;

                if (inListItem) {
                    // Check current list type from stack
                    const currentList = listStack.length > 0 ? listStack[listStack.length - 1] : null;
                    const listLevel = currentListLevel > 0 ? currentListLevel - 1 : 0;

                    if (currentList && currentList.type === 'ordered') {
                        // Ordered list: add number prefix manually
                        const numberPrefix = new TextRun({
                            text: `${currentList.counter}. `,
                            font: currentStyle.fontMain,
                            size: currentStyle.fontSizeMain,
                            color: currentStyle.colorMain || "000000"
                        });
                        paraConfig = {
                            children: [numberPrefix, ...runs],
                            indent: { left: 720 * (listLevel + 1), hanging: 360 },
                            spacing: { before: 60, after: 60 }
                        };
                    } else {
                        // Bullet list: use built-in bullet
                        paraConfig = {
                            children: runs,
                            bullet: { level: listLevel },
                            spacing: { before: 60, after: 60 }
                        };
                    }
                    inListItem = false;
                } else if (blockquoteLevel > 0) {
                    // Blockquote styling: left border, gray background, indented
                    paraConfig = {
                        children: runs,
                        border: {
                            left: {
                                color: "CCCCCC",
                                space: 10,
                                style: BorderStyle.SINGLE,
                                size: 24
                            }
                        },
                        indent: { left: 400 * blockquoteLevel },
                        shading: { type: ShadingType.CLEAR, color: "auto", fill: "F5F5F5" },
                        spacing: { before: 100, after: 100 }
                    };
                } else {
                    paraConfig = {
                        children: runs,
                        spacing: {
                            line: currentStyle.lineSpacing,
                            lineRule: currentStyle.lineRule, // Use explicit rule if present
                            before: bodySpacing.before,
                            after: bodySpacing.after
                        },
                    };

                    // Logic for indent: explicit > redHeader default > none
                    if (currentStyle.paragraphIndent !== undefined) {
                        // Explicitly set (can be 0 to disable)
                        if (currentStyle.paragraphIndent > 0) {
                            paraConfig.indent = { firstLine: currentStyle.paragraphIndent };
                        }
                    } else if (currentStyle.redHeader) {
                        paraConfig.indent = { firstLine: 640 };
                        paraConfig.alignment = AlignmentType.JUSTIFIED;
                    } else if (currentStyle.firstLineIndent) {
                        paraConfig.indent = { firstLine: currentStyle.firstLineIndent };
                    }
                }
                docChildren.push(new Paragraph(paraConfig));
                console.log(`  -> Added Paragraph (blockquote=${blockquoteLevel}, list=${currentListLevel})`);
                i += 2;
            }
        }
        // --- Horizontal Rule (---) ---
        else if (token.type === 'hr') {
            docChildren.push(new Paragraph({
                children: [],
                border: {
                    bottom: {
                        color: "999999",
                        space: 1,
                        style: BorderStyle.SINGLE,
                        size: 6
                    }
                },
                spacing: { before: 200, after: 200 }
            }));
            console.log(`  -> Added Horizontal Rule`);
        }
        // --- Blockquote ---
        else if (token.type === 'blockquote_open') {
            blockquoteLevel++;
            console.log(`  -> Blockquote open (level ${blockquoteLevel})`);
        }
        else if (token.type === 'blockquote_close') {
            blockquoteLevel--;
            console.log(`  -> Blockquote close (level ${blockquoteLevel})`);
        }
        // --- Bullet List (Unordered) ---
        else if (token.type === 'bullet_list_open') {
            listStack.push({ type: 'bullet', level: currentListLevel, counter: 0 });
            currentListLevel++;
            console.log(`  -> Bullet list open (level ${currentListLevel})`);
        }
        else if (token.type === 'bullet_list_close') {
            listStack.pop();
            currentListLevel--;
            console.log(`  -> Bullet list close (level ${currentListLevel})`);
        }
        // --- Ordered List ---
        else if (token.type === 'ordered_list_open') {
            const start = token.attrGet ? (token.attrGet('start') || 1) : 1;
            listStack.push({ type: 'ordered', level: currentListLevel, counter: parseInt(start) - 1 });
            currentListLevel++;
            console.log(`  -> Ordered list open (level ${currentListLevel}, start=${start})`);
        }
        else if (token.type === 'ordered_list_close') {
            listStack.pop();
            currentListLevel--;
            console.log(`  -> Ordered list close (level ${currentListLevel})`);
        }
        // --- List Item ---
        else if (token.type === 'list_item_open') {
            inListItem = true;
            // Increment counter for ordered lists
            if (listStack.length > 0) {
                const currentList = listStack[listStack.length - 1];
                if (currentList.type === 'ordered') {
                    currentList.counter++;
                }
            }
        }
        else if (token.type === 'list_item_close') {
            inListItem = false;
        }
        // --- Math Block ($$...$$) ---
        else if (token.type === 'math_block') {
            const mathCode = token.content.trim();
            console.log(`  -> Processing Math Block...`);

            try {
                const buffer = await renderMathLocally(mathCode, true); // displayMode = true

                if (buffer) {
                    const dimensions = sizeOf(buffer);
                    let width = dimensions.width;
                    let height = dimensions.height;
                    const MAX_WIDTH = 500;
                    if (width > MAX_WIDTH) {
                        const ratio = MAX_WIDTH / width;
                        width = MAX_WIDTH;
                        height = Math.round(height * ratio);
                    }

                    docChildren.push(new Paragraph({
                        children: [new ImageRun({
                            data: buffer,
                            transformation: { width, height }
                        })],
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200, after: 200 }
                    }));
                    console.log(`  -> Added Math Block as Image`);
                } else {
                    // Fallback: show raw math code
                    docChildren.push(new Paragraph({
                        children: [new TextRun({
                            text: `[Math: ${mathCode}]`,
                            font: "Courier New",
                            color: "666666"
                        })],
                        alignment: AlignmentType.CENTER
                    }));
                }
            } catch (err) {
                console.warn(`   ⚠️ Math block rendering failed: ${err.message}`);
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: `[Math Render Failed]`, color: "FF0000" })],
                    alignment: AlignmentType.CENTER
                }));
            }
        }
        else if (token.type === 'fence' || token.type === 'code_block') {
            if (token.info === 'mermaid') {
                const mermaidCode = token.content.trim();
                console.log(`  -> Processing Mermaid Diagram (Local Only)...`);
                
                try {
                    // Direct local rendering using bundled mermaid.js and Playwright
                    const buffer = await renderMermaidLocally(mermaidCode);

                    if (buffer) {
                        const dimensions = sizeOf(buffer);
                        let width = dimensions.width;
                        let height = dimensions.height;
                        const MAX_WIDTH = 600;
                        if (width > MAX_WIDTH) {
                            const ratio = MAX_WIDTH / width;
                            width = MAX_WIDTH;
                            height = Math.round(height * ratio);
                        }

                        docChildren.push(new Paragraph({
                            children: [new ImageRun({
                                data: buffer,
                                transformation: { width, height }
                            })],
                            alignment: AlignmentType.CENTER,
                            spacing: { before: 200, after: 200 }
                        }));
                        console.log(`  -> Added Mermaid Diagram as Image`);
                    } else {
                        throw new Error("Local rendering engine failed");
                    }
                } catch (err) {
                    console.warn(`   ⚠️ Mermaid rendering failed: ${err.message}`);
                    docChildren.push(new Paragraph({
                        children: [new TextRun({ text: `[Mermaid Render Failed: Local engine error]`, color: "FF0000" })],
                        alignment: AlignmentType.CENTER
                    }));
                }
                continue; 
            }

            const codeText = token.content;
            docChildren.push(new Paragraph({
                children: [new TextRun({
                    text: codeText,
                    font: "Courier New",
                    size: currentStyle.fontSizeMain || 21, // 10.5pt
                    color: "333333"
                })],
                spacing: { 
                    line: currentStyle.lineSpacing,
                    before: 100,
                    after: 100
                },
                shading: { type: ShadingType.CLEAR, color: "auto", fill: "F5F5F5" }
            }));
            console.log(`  -> Added Code Block`);
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
                const isHeaderRow = tableBuffer.isHeader; 
                tableBuffer.rows.push({ content: tableBuffer.currentRow, isHeader: isHeaderRow }); 
                tableBuffer.currentRow = null; 
            }
        }
        else if (token.type === 'table_close') {
            if (tableBuffer && tableBuffer.rows.length > 0) {
                const tblConfig = currentStyle.table || {
                    borderStyle: "single", borderColor: "000000", borderSize: 4,
                    headerBold: true, headerColor: "000000", cellAlign: "left"
                };
                
                // Process table rows asynchronously
                const docxRows = await Promise.all(tableBuffer.rows.map(async rowObj => {
                    const cells = await Promise.all(rowObj.content.map(async cellText => {
                        let isBold = rowObj.isHeader ? tblConfig.headerBold : false;
                        let color = rowObj.isHeader ? tblConfig.headerColor : (currentStyle.colorMain || "000000");
                        let align = AlignmentType.LEFT;
                        if (tblConfig.cellAlign === 'center') align = AlignmentType.CENTER;
                        if (tblConfig.cellAlign === 'right') align = AlignmentType.RIGHT;
                        
                        const cellTokens = md.parseInline(cellText, {})[0];
                        const cellRuns = await processInline(md, cellTokens, currentStyle, baseDir, color, isBold);
                        
                        return new TableCell({
                            children: [new Paragraph({
                                children: cellRuns,
                                alignment: align,
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                        });
                    }));
                    return new TableRow({ children: cells });
                }));

                const borderObj = {
                    style: getBorderStyle(tblConfig.borderStyle),
                    size: tblConfig.borderSize,
                    color: tblConfig.borderColor
                };
                const tableBorders = {
                    top: borderObj, bottom: borderObj, left: borderObj, right: borderObj,
                    insideHorizontal: borderObj, insideVertical: borderObj
                };
                docChildren.push(new Table({
                    rows: docxRows,
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: tableBorders
                }));
                tableBuffer = null;
            }
        }
        // --- Footnote Block (skip processing - collected in first pass) ---
        else if (token.type === 'footnote_block_open' || token.type === 'footnote_block_close') {
            // Skip - footnotes are processed separately
            console.log(`  -> Skipping footnote block`);
        }
        else if (token.type === 'footnote_open') {
            // Skip footnote definition body
            console.log(`  -> Skipping footnote definition`);
        }
        else if (token.type === 'footnote_close') {
            // Skip
        }
        else if (token.type === 'footnote_anchor') {
            // Skip back-reference anchor
        }
    }

    // First pass: collect all footnote definitions from the environment
    const footnoteDefinitions = {};
    if (md.options && md.options.env && md.options.env.footnotes && md.options.env.footnotes.list) {
        md.options.env.footnotes.list.forEach((fn, idx) => {
            footnoteDefinitions[idx + 1] = fn.content || `Footnote ${idx + 1}`;
        });
    }

    // Build footnotes configuration for docx if any exist
    const footnotesConfig = {};
    for (const [id, content] of footnotes.entries()) {
        footnotesConfig[id] = {
            children: [new Paragraph({
                children: [new TextRun({
                    text: content,
                    font: currentStyle.fontMain,
                    size: (currentStyle.fontSizeMain || 21) - 4 // Slightly smaller
                })]
            })]
        };
    }

    const doc = new Document({
        footnotes: Object.keys(footnotesConfig).length > 0 ? footnotesConfig : undefined,
        sections: [{
            properties: { page: { margin: currentStyle.margin } },
            children: docChildren
        }]
    });

    const buffer = await Packer.toBuffer(doc);

    // Close shared browser after document generation
    await closeSharedBrowser();

    return buffer;
}

module.exports = { generateDocx, closeSharedBrowser };