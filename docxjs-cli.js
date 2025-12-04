#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const MarkdownIt = require('markdown-it');
const inquirer = require('inquirer');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    BorderStyle, HeadingLevel, AlignmentType, WidthType, VerticalAlign, ShadingType,
    ExternalHyperlink, UnderlineType
} = require('docx');

// --- 0. 加载配置 (自动扫描 templates 目录) ---
function loadTemplates(customConfigPath) {
    let templates = {};
    const templatesDir = path.join(__dirname, 'templates');

    // 1. 自动扫描并加载 templates/ 下的所有 .json 文件
    if (fs.existsSync(templatesDir)) {
        try {
            const files = fs.readdirSync(templatesDir).filter(file => file.toLowerCase().endsWith('.json'));
            files.sort(); 
            files.forEach(file => {
                const fullPath = path.join(templatesDir, file);
                try {
                    const fileContent = JSON.parse(fs.readFileSync(fullPath, 'utf-8'));
                    templates = { ...templates, ...fileContent };
                } catch (e) {
                    console.warn(`⚠️ Warning: Failed to parse template file '${file}': ${e.message}`);
                }
            });
        } catch (e) {
            console.error(`❌ Error reading templates directory: ${e.message}`);
        }
    } else {
        console.warn(`⚠️ Templates directory not found at: ${templatesDir}`);
    }

    // 2. 加载用户通过命令行 --config 指定的配置 (最高优先级)
    if (customConfigPath) {
        const absPath = path.resolve(process.cwd(), customConfigPath);
        if (fs.existsSync(absPath)) {
            try {
                const userTemplates = JSON.parse(fs.readFileSync(absPath, 'utf-8'));
                console.log(`🎨 Loaded custom configuration from: ${customConfigPath}`);
                templates = { ...templates, ...userTemplates };
            } catch (e) { console.error(`❌ Failed to load custom config ${customConfigPath}: `, e.message); }
        } else { console.warn(`⚠️ Custom config file not found: ${customConfigPath}`); }
    }
    return templates;
}

// --- 主逻辑 ---
(async () => {
    const rawArgv = yargs(hideBin(process.argv)).argv;
    const templates = loadTemplates(rawArgv.config);
    const availableTemplates = Object.keys(templates);

    const argv = yargs(hideBin(process.argv))
        .usage('Usage: $0 <input.md> -o <output.docx> [options]')
        .command('$0 <input>', 'Convert Markdown to Docx')
        .option('output', { alias: 'o', type: 'string', default: 'output.docx' })
        .option('template', { alias: 't', type: 'string', description: `Template name. (default, official, etc.)` })
        .option('config', { alias: 'c', type: 'string' })
        .option('reference-doc', { alias: 'r', type: 'string' })
        .demandCommand(1)
        .help()
        .argv;

    const inputPath = argv.input;
    const outputPath = argv.output;
    let templateName = argv.template;
    const referenceDocPath = argv['reference-doc'];

    // --- 模式判定 ---
    let mode = "Template"; 
    
    if (referenceDocPath && !templateName) {
        mode = "Clone"; // 纯克隆模式
        console.log(`🤖 Mode: Pure Clone (Extracting everything from reference doc)`);
        templateName = 'default'; // Base is default, but we will override everything
    } else if (templateName) {
        mode = "Hybrid"; // 模板模式 (或混合)
        console.log(`📝 Mode: Template/Hybrid (${templateName})`);
    } else {
        // Interactive Mode
        mode = "Interactive";
        console.log(`
👋 欢迎使用 docxjs-cli 文档转换工具`);
        const choices = availableTemplates.map(key => ({
            name: `${key.padEnd(20)} - ${templates[key].description || "No desc"}`,
            value: key
        }));
        
        const answer = await inquirer.prompt([{ 
            type: 'list',
            name: 'selectedTemplate',
            message: '请选择目标文档格式 (Select Template):',
            choices: choices,
            pageSize: 15
        }]);
        templateName = answer.selectedTemplate;
    }

    // --- 2. 样式加载 ---
    let currentStyle = { ...(templates[templateName] || templates['default']) };

    // --- 2.1 样式合并 (Clone/Hybrid) ---
    if (referenceDocPath) {
        console.log(`🔍 Extracting styles from reference doc: ${referenceDocPath}...`);
        try {
            const extractorPath = path.join(__dirname, 'style_extractor.py');
            let pythonExecutable = 'python3'; 
            if (process.env.DOCXJS_PYTHON_PATH) {
                pythonExecutable = process.env.DOCXJS_PYTHON_PATH;
            } else if (fs.existsSync(path.join(__dirname, 'venv'))) {
                pythonExecutable = path.join(__dirname, 'venv', 'bin', 'python3');
            }
            else {
                const globalEnvPath = path.join(process.env.HOME || process.env.USERPROFILE, '.docxjs-cli-env', 'bin', 'python3');
                if (fs.existsSync(globalEnvPath)) { pythonExecutable = globalEnvPath; }
            }
            
            const pythonCmd = `"${pythonExecutable}" "${extractorPath}" "${referenceDocPath}"`;
            const stdout = execSync(pythonCmd, { encoding: 'utf-8' });
            const extractedStyles = JSON.parse(stdout);

            if (extractedStyles.error) {
                console.warn(`⚠️ Style extraction failed: ${extractedStyles.error}`);
            } else {
                console.log(`✅ Styles extracted successfully! Merging...`);
                Object.keys(extractedStyles).forEach(key => {
                    if (extractedStyles[key] !== null && extractedStyles[key] !== undefined && key !== 'detailed_styles_info') {
                        // Deep merge for 'table' object
                        if (key === 'table' && typeof extractedStyles[key] === 'object') {
                             currentStyle[key] = { ...currentStyle[key], ...extractedStyles[key] };
                        } else {
                             currentStyle[key] = extractedStyles[key];
                        }
                    }
                });
                 if (!extractedStyles.fontH1 && extractedStyles.detailed_styles_info) {
                     const genericHeading = extractedStyles.detailed_styles_info.find(s => s.name === 'Heading');
                     if (genericHeading && genericHeading.font_name) currentStyle.fontHeader1 = genericHeading.font_name;
                }
            }
        } catch (e) { console.warn(`⚠️ Python script failed: ${e.message}`); }
    }

    // --- Helper ---
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

    // --- 3. Markdown Parsing ---
    const md = new MarkdownIt();
    const mdContent = fs.readFileSync(inputPath, 'utf-8');
    const tokens = md.parse(mdContent, {});

    // --- 4. AST to Docx ---
    const docChildren = [];

    function processInline(inlineToken, colorOverride, boldOverride) {
        const runs = []; // Can contain TextRun or ExternalHyperlink
        if (!inlineToken.children) return runs;
        
        let isBold = false;
        let isItalic = false;
        
        let inLink = false;
        let linkHref = "";
        let linkChildren = [];

        // Get hyperlink config
        const linkConfig = currentStyle.hyperlink || { color: "0000FF", underline: true };

        inlineToken.children.forEach(token => {
            if (token.type === 'link_open') {
                inLink = true;
                linkHref = token.attrs ? token.attrs.find(attr => attr[0] === 'href')[1] : "";
                linkChildren = []; // Start collecting link content
            } 
            else if (token.type === 'link_close') {
                inLink = false;
                // Create ExternalHyperlink
                if (linkChildren.length > 0) {
                    const hyperlink = new ExternalHyperlink({
                        children: linkChildren,
                        link: linkHref
                    });
                    runs.push(hyperlink);
                }
            }
            else {
                // TextRun creation logic
                let run = null;
                
                if (token.type === 'text') {
                    run = new TextRun({
                        text: token.content,
                        bold: (boldOverride === true) || isBold,
                        italics: isItalic,
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

                if (run) {
                    if (inLink) {
                        linkChildren.push(run);
                    } else {
                        runs.push(run);
                    }
                }
            }
        });
        return runs;
    }
    
    let tableBuffer = null;
    let inListItem = false;

    // Default spacing configs if not provided
    const headerSpacing = currentStyle.headerSpacing || { before: 200, after: 200 };
    const bodySpacing = currentStyle.bodySpacing || { before: 0, after: 0 };

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];

        if (token.type === 'heading_open') {
            const level = parseInt(token.tag.replace('h', ''));
            const inlineToken = tokens[i + 1];
            const textContent = inlineToken.content;
            
            // Apply header spacing
            let paraObj = {
                children: [],
                spacing: { ...headerSpacing } 
            };

            if (level === 1) {
                paraObj.heading = HeadingLevel.HEADING_1;
                paraObj.alignment = currentStyle.redHeader ? AlignmentType.CENTER : AlignmentType.LEFT;
                let h1Color = currentStyle.redHeader ? "FF0000" : (currentStyle.colorHeader1 || "000000");
                let h1Bold = currentStyle.redHeader ? false : true; 
                paraObj.children.push(new TextRun({
                    text: textContent,
                    font: currentStyle.fontHeader1 || currentStyle.fontMain,
                    size: currentStyle.fontSizeH1 || 32,
                    color: h1Color,
                    bold: h1Bold
                }));
                if (currentStyle.redHeader) paraObj.spacing = { after: 400 }; // Red header special spacing override
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
            } else {
                // Level 4+
                paraObj.heading = HeadingLevel.HEADING_4;
                paraObj.children.push(new TextRun({ 
                    text: textContent, 
                    font: currentStyle.fontMain, 
                    size: currentStyle.fontSizeMain,
                    color: currentStyle.colorMain || "000000", // Ensure black/default color
                    bold: true // Default bold for lower level headers
                }));
            }
            docChildren.push(new Paragraph(paraObj));
            i += 2; 
        }
        else if (token.type === 'paragraph_open') {
            if (!tableBuffer) {
                 const runs = processInline(tokens[i + 1]);
                 let paraConfig;
                 
                 if (inListItem) {
                    paraConfig = {
                        children: runs,
                        bullet: { level: 0 }
                    };
                    inListItem = false;
                 } else {
                    paraConfig = {
                        children: runs,
                        spacing: { 
                            line: currentStyle.lineSpacing,
                            before: bodySpacing.before,
                            after: bodySpacing.after
                        },
                    };
                    if (currentStyle.redHeader) {
                        paraConfig.indent = { firstLine: 640 };
                        paraConfig.alignment = AlignmentType.JUSTIFIED;
                    } else if (currentStyle.firstLineIndent) {
                        paraConfig.indent = { firstLine: currentStyle.firstLineIndent };
                    }
                 }
                 docChildren.push(new Paragraph(paraConfig));
                 i += 2;
            }
        }
        else if (token.type === 'list_item_open') {
            inListItem = true;
        }
        else if (token.type === 'list_item_close') {
            inListItem = false;
        }
        else if (token.type === 'table_open') { 
            tableBuffer = { rows: [], isHeader: false }; 
        }
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
                const docxRows = tableBuffer.rows.map(rowObj => {
                    return new TableRow({
                        children: rowObj.content.map(cellText => {
                            let isBold = rowObj.isHeader ? tblConfig.headerBold : false;
                            let color = rowObj.isHeader ? tblConfig.headerColor : (currentStyle.colorMain || "000000");
                            let align = AlignmentType.LEFT;
                            if (tblConfig.cellAlign === 'center') align = AlignmentType.CENTER;
                            if (tblConfig.cellAlign === 'right') align = AlignmentType.RIGHT;
                            return new TableCell({
                                children: [new Paragraph({
                                    children: (() => {
                                        const cellTokens = md.parseInline(cellText, {})[0];
                                        return processInline(cellTokens, color, isBold);
                                    })(),
                                    alignment: align,
                                })],
                                verticalAlign: VerticalAlign.CENTER,
                            });
                        })
                    });
                });
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
    }

    const doc = new Document({
        sections: [{
            properties: { page: { margin: currentStyle.margin } },
            children: docChildren
        }]
    });

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync(outputPath, buffer);
        console.log(`✅ Success! Created ${outputPath}`);
    });

})();
