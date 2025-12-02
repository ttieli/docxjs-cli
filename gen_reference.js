const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, BorderStyle, WidthType } = require("docx");

// 公文格式常量
const FONTS = {
    main: "FangSong_GB2312",      // 仿宋
    heading1: "SimHei",           // 黑体 (用于一级小标题 "一、")
    heading2: "KaiTi_GB2312",     // 楷体 (用于二级小标题 "（一）")
    title: "FZXiaoBiaoSong-B05S", // 方正小标宋 (用于红头)
    code: "Courier New"
};

const SIZES = {
    title: 44,    // 二号 (22pt)
    heading1: 32, // 小三号 (16pt) -> 实际上公文一级标题通常是小三号(32)或三号(32)。这里设为32
    heading2: 28, // 四号 (14pt)
    body: 32,     // 三号 (16pt)
    small: 24     // 小四号 (12pt)
};

const LINE_SPACING = 560; // 28磅

// 辅助函数：创建带样式的段落
function createStyledPara(text, styleId) {
    return new Paragraph({
        text: text,
        style: styleId
    });
}

const doc = new Document({
    styles: {
        default: {
            document: {
                run: {
                    font: FONTS.main,
                    size: SIZES.body,
                    color: "000000"
                },
                paragraph: {
                    spacing: { line: LINE_SPACING },
                    alignment: AlignmentType.JUSTIFIED
                },
            },
        },
        paragraphStyles: [
            // --- 核心样式 ---
            {
                id: "Normal",
                name: "Normal",
                run: { font: FONTS.main, size: SIZES.body },
                paragraph: { 
                    spacing: { line: LINE_SPACING },
                    indent: { firstLine: 640 } // 首行缩进2字符
                }
            },
            {
                id: "BodyText",
                name: "Body Text",
                basedOn: "Normal",
                next: "Normal",
                run: { font: FONTS.main, size: SIZES.body },
                paragraph: { indent: { firstLine: 640 } }
            },
            {
                id: "FirstParagraph", // Pandoc 可能会用到
                name: "First Paragraph",
                basedOn: "Body Text",
                next: "Body Text",
                paragraph: { indent: { firstLine: 640 } }
            },
            {
                id: "Title",
                name: "Title",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    font: FONTS.title,
                    size: SIZES.title,
                    color: "FF0000", // 红头
                    bold: false
                },
                paragraph: {
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 0, after: 480 }, // 标题下空一行
                    indent: { firstLine: 0 }
                }
            },
            {
                id: "Subtitle",
                name: "Subtitle",
                basedOn: "Title",
                next: "Normal",
                run: {
                    font: FONTS.heading2, // 楷体
                    size: SIZES.heading1, // 小三
                    color: "000000"
                },
                paragraph: {
                    alignment: AlignmentType.CENTER,
                    indent: { firstLine: 0 }
                }
            },
            {
                id: "Author",
                name: "Author",
                basedOn: "Normal",
                paragraph: { alignment: AlignmentType.CENTER }
            },
            {
                id: "Date",
                name: "Date",
                basedOn: "Normal",
                paragraph: { alignment: AlignmentType.CENTER }
            },
            {
                id: "Abstract",
                name: "Abstract",
                basedOn: "Normal",
                run: { font: FONTS.heading2, size: SIZES.heading2 }, // 楷体 四号
                paragraph: { indent: { left: 480, right: 480 } } // 左右缩进
            },
            
            // --- 标题样式 ---
            {
                id: "Heading1",
                name: "Heading 1",
                basedOn: "Normal",
                next: "Body Text",
                run: {
                    font: FONTS.heading1, // 黑体
                    size: SIZES.heading1, // 小三
                    color: "000000",
                    bold: false // 黑体自带粗
                },
                paragraph: {
                    alignment: AlignmentType.LEFT, // 公文一级标题一般左对齐，带缩进或不带
                    indent: { firstLine: 640 }, // 通常"一、"需要缩进2字符? 实际上很多公文一级标题是缩进的
                    spacing: { before: 240, after: 240 }
                }
            },
            {
                id: "Heading2",
                name: "Heading 2",
                basedOn: "Normal",
                next: "Body Text",
                run: {
                    font: FONTS.heading2, // 楷体
                    size: SIZES.heading2, // 四号
                    color: "000000",
                    bold: true // 楷体有时加粗以示区别
                },
                paragraph: {
                    indent: { firstLine: 640 },
                    spacing: { before: 120, after: 120 }
                }
            },
            {
                id: "Heading3",
                name: "Heading 3",
                basedOn: "Normal",
                next: "Body Text",
                run: {
                    font: FONTS.main, // 仿宋
                    size: SIZES.heading2, // 四号
                    bold: true
                },
                paragraph: { indent: { firstLine: 640 } }
            },
            {
                id: "Heading4",
                name: "Heading 4",
                basedOn: "Normal",
                run: { size: SIZES.heading2, italic: true }
            },
             {
                id: "Heading5",
                name: "Heading 5",
                basedOn: "Normal",
                run: { size: SIZES.heading2 }
            },
            {
                id: "Heading6",
                name: "Heading 6",
                basedOn: "Normal",
                run: { size: SIZES.heading2 }
            },
            
            // --- 其他功能样式 ---
            {
                id: "BlockText",
                name: "Block Text",
                basedOn: "Normal",
                run: { font: FONTS.heading2 }, // 楷体
                paragraph: { indent: { left: 640, right: 640 } }
            },
            {
                id: "DefinitionTerm",
                name: "Definition Term",
                basedOn: "Normal",
                run: { bold: true }
            },
            {
                id: "Definition",
                name: "Definition",
                basedOn: "Normal",
                paragraph: { indent: { left: 640 } }
            },
            {
                id: "Caption", // Image/Table Caption
                name: "Caption",
                basedOn: "Normal",
                run: { size: SIZES.small, bold: true },
                paragraph: { alignment: AlignmentType.CENTER }
            },
            {
                id: "TableCaption",
                name: "Table Caption",
                basedOn: "Caption"
            },
            {
                id: "ImageCaption",
                name: "Image Caption",
                basedOn: "Caption"
            },
             {
                id: "FootnoteText",
                name: "Footnote Text",
                run: { size: 20 }, // 10pt
            }
        ],
        characterStyles: [
             {
                id: "BodyTextChar",
                name: "Body Text Char",
                run: { font: FONTS.main }
            },
            {
                id: "VerbatimChar",
                name: "Verbatim Char",
                run: { font: FONTS.code, size: 22 }
            },
            {
                id: "Hyperlink",
                name: "Hyperlink",
                run: { color: "0563C1", underline: {} }
            }
        ]
    },
    sections: [
        {
            properties: {
                page: {
                    margin: {
                        top: "3.7cm",
                        bottom: "3.5cm",
                        left: "2.8cm",
                        right: "2.6cm",
                    },
                },
            },
            children: [
                createStyledPara("Title (红头标题)", "Title"),
                createStyledPara("Subtitle (副标题)", "Subtitle"),
                createStyledPara("Author (发文机关)", "Author"),
                createStyledPara("Date (成文日期)", "Date"),
                createStyledPara("Abstract (摘要)", "Abstract"),
                
                createStyledPara("Heading 1 (一级标题)", "Heading1"),
                createStyledPara("Heading 2 (二级标题)", "Heading2"),
                createStyledPara("Heading 3 (三级标题)", "Heading3"),
                createStyledPara("Heading 4", "Heading4"),
                createStyledPara("Heading 5", "Heading5"),
                createStyledPara("Heading 6", "Heading6"),
                createStyledPara("Heading 7", "Heading6"), // Reuse
                createStyledPara("Heading 8", "Heading6"),
                createStyledPara("Heading 9", "Heading6"),

                createStyledPara("First Paragraph (首段，通常带缩进)", "FirstParagraph"),
                createStyledPara("Body Text (正文内容)。仿宋_GB2312，三号字，28磅行距，首行缩进2字符。", "BodyText"),
                
                new Paragraph({
                    children: [
                        new TextRun({ text: "Body Text Char ", style: "BodyTextChar" }),
                        new TextRun({ text: "Verbatim Char (Code) ", style: "VerbatimChar" }),
                        new TextRun({ text: "Hyperlink", style: "Hyperlink" }),
                        new TextRun({ text: "Footnote Ref", style: "FootnoteReference" }), // System style
                    ],
                    style: "BodyText"
                }),

                createStyledPara("Block Text (引用块)", "BlockText"),
                createStyledPara("Table Caption", "TableCaption"),
                
                // Simple Table for reference
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph("Table Cell 1")] }),
                                new TableCell({ children: [new Paragraph("Table Cell 2")] }),
                            ]
                        })
                    ],
                    width: { size: 100, type: WidthType.PERCENTAGE },
                     borders: {
                        top: { style: BorderStyle.SINGLE, size: 1 },
                        bottom: { style: BorderStyle.SINGLE, size: 1 },
                        left: { style: BorderStyle.SINGLE, size: 1 },
                        right: { style: BorderStyle.SINGLE, size: 1 },
                        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
                        insideVertical: { style: BorderStyle.SINGLE, size: 1 }
                    }
                }),

                createStyledPara("Image Caption", "ImageCaption"),
                createStyledPara("Definition Term", "DefinitionTerm"),
                createStyledPara("Definition", "Definition"),
                createStyledPara("Footnote Text", "FootnoteText"),
                createStyledPara("Footnote Block Text", "FootnoteText"), // Fallback
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("gov-reference.docx", buffer);
    console.log("Generated gov-reference.docx with full Pandoc styles.");
});