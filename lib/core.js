const MarkdownIt = require('markdown-it');
const { 
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    BorderStyle, HeadingLevel, AlignmentType, WidthType, VerticalAlign, ShadingType,
    ExternalHyperlink, UnderlineType
} = require('docx');

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

function processInline(md, inlineToken, currentStyle, colorOverride, boldOverride) {
    const runs = [];
    if (!inlineToken.children) return runs;
    
    let isBold = false;
    let isItalic = false;
    
    let inLink = false;
    let linkHref = "";
    let linkChildren = [];

    const linkConfig = currentStyle.hyperlink || { color: "0000FF", underline: true };

    inlineToken.children.forEach(token => {
        if (token.type === 'link_open') {
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
                if (inLink) linkChildren.push(run);
                else runs.push(run);
            }
        }
    });
    return runs;
}

async function generateDocx(markdownContent, currentStyle) {
    const md = new MarkdownIt();
    const tokens = md.parse(markdownContent, {});
    const docChildren = [];
    let tableBuffer = null;
    let inListItem = false;

    const headerSpacing = currentStyle.headerSpacing || { before: 200, after: 200 };
    const bodySpacing = currentStyle.bodySpacing || { before: 0, after: 0 };

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];

        if (token.type === 'heading_open') {
            const level = parseInt(token.tag.replace('h', ''));
            const inlineToken = tokens[i + 1];
            const textContent = inlineToken.content;
            
            let paraObj = { children: [], spacing: { ...headerSpacing } };

            if (level === 1) {
                paraObj.heading = HeadingLevel.HEADING_1;
                paraObj.alignment = currentStyle.redHeader ? AlignmentType.CENTER : AlignmentType.LEFT;
                
                // Color: Prefer config, fallback to Red if redHeader mode, else Black
                let h1Color = currentStyle.colorHeader1 ? currentStyle.colorHeader1 : (currentStyle.redHeader ? "FF0000" : "000000");
                
                // Bold: Standard spec usually requires Bold for both Red Headers and Titles. 
                // Previous logic disabled bold for redHeader, which contradicts standard "二号加粗".
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
            } else {
                paraObj.heading = HeadingLevel.HEADING_4;
                paraObj.children.push(new TextRun({ 
                    text: textContent, 
                    font: currentStyle.fontMain, 
                    size: currentStyle.fontSizeMain,
                    color: currentStyle.colorMain || "000000",
                    bold: true 
                }));
            }
            docChildren.push(new Paragraph(paraObj));
            i += 2; 
        }
        else if (token.type === 'paragraph_open') {
            if (!tableBuffer) {
                 const runs = processInline(md, tokens[i + 1], currentStyle);
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
                                        return processInline(md, cellTokens, currentStyle, color, isBold);
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

    return await Packer.toBuffer(doc);
}

module.exports = { generateDocx };
