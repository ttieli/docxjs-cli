const express = require('express');
const { generateDocx } = require('../lib/core');
const { generatePreviewBuffer } = require('../lib/sample-generator');
const templateManager = require('../lib/template-manager');
const { extractStyles } = require('../lib/python-bridge');
const path = require('path');
const fs = require('fs');
const os = require('os');
const multer = require('multer');
const mammoth = require('mammoth');
const TurndownService = require('turndown');
const { gfm } = require('turndown-plugin-gfm');

const app = express();
const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 10 * 1024 * 1024 } // Increased to 10MB for imports
});

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

const PUBLIC_DIR = path.join(__dirname, '..', 'public');
if (fs.existsSync(PUBLIC_DIR)) {
    app.use(express.static(PUBLIC_DIR));
}

function sanitizeFileName(name) {
    if (!name) return 'output.docx';
    const cleaned = name.replace(/[\\/]+/g, '').replace(/[^a-zA-Z0-9._-]/g, '_');
    return cleaned || 'output.docx';
}

function mergeStyles(baseStyle, overrides) {
    const merged = { ...(baseStyle || {}) };
    Object.keys(overrides || {}).forEach(key => {
        if (overrides[key] === null || overrides[key] === undefined) return;
        if (key === 'table' && typeof overrides[key] === 'object') {
            merged[key] = { ...(merged[key] || {}), ...overrides[key] };
        } else {
            merged[key] = overrides[key];
        }
    });
    return merged;
}

// API: Preview (Dynamic)
app.post('/api/preview/dynamic', async (req, res) => {
    try {
        const { styleConfig } = req.body;
        const buffer = await generatePreviewBuffer(styleConfig || {});
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(buffer);
    } catch (error) {
        console.error("Preview generation error:", error);
        res.status(500).json({ error: error.message });
    }
});

// API: Import Docx to Markdown
app.post('/api/import-docx', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        // 1. Convert Docx to HTML using Mammoth
        const { value: html } = await mammoth.convertToHtml({ buffer: req.file.buffer });

        // 2. Convert HTML to Markdown using Turndown
        const turndownService = new TurndownService({
            headingStyle: 'atx',
            codeBlockStyle: 'fenced'
        });
        // Use GFM plugin for Tables support
        turndownService.use(gfm);

        const markdown = turndownService.turndown(html);

        res.json({ markdown });

    } catch (error) {
        console.error("Import error:", error);
        res.status(500).json({ error: 'Failed to parse Word document: ' + error.message });
    }
});

// API: Convert
// Supports JSON payloads or multipart/form-data (with optional referenceDoc)
app.post('/api/convert', upload.single('referenceDoc'), async (req, res) => {
    let tempDocPath = null;
    try {
        const isJson = req.is('application/json');
        let { markdown, templateName, styleConfig, filename } = req.body || {};

        // Handle styleConfig possibly delivered as JSON string in multipart form
        if (!isJson && typeof styleConfig === 'string') {
            try { styleConfig = JSON.parse(styleConfig); } catch (_) { /* ignore parse failure */ }
        }

        if (!markdown || typeof markdown !== 'string' || markdown.trim().length === 0) {
            return res.status(400).json({ error: 'Markdown content required' });
        }

        // Determine base style: explicit styleConfig > template name > default
        let finalStyle = {};
        if (styleConfig) {
            finalStyle = styleConfig;
        } else {
            const template = templateManager.getTemplate(templateName) || templateManager.getTemplate('default');
            if (template) {
                const { name, isBuiltIn, ...styleData } = template;
                finalStyle = styleData;
            }
        }

        // Optional reference doc extraction
        if (req.file) {
            tempDocPath = path.join(os.tmpdir(), `reference-${Date.now()}-${Math.random().toString(16).slice(2)}.docx`);
            await fs.promises.writeFile(tempDocPath, req.file.buffer);
            try {
                // Use the new Bridge instead of local execution logic
                const extracted = await extractStyles(tempDocPath);
                
                if (!extracted.error) {
                    finalStyle = mergeStyles(finalStyle, extracted);
                    if (!extracted.fontH1 && extracted.detailed_styles_info) {
                        const genericHeading = extracted.detailed_styles_info.find(s => s.name === 'Heading');
                        if (genericHeading && genericHeading.font_name) finalStyle.fontHeader1 = genericHeading.font_name;
                    }
                } else {
                    console.warn(`⚠️ Style extraction failed: ${extracted.error}`);
                }
            } catch (err) {
                console.warn(`⚠️ Reference style extraction failed: ${err.message}`);
            }
        }

        const buffer = await generateDocx(markdown, finalStyle);
        const downloadName = sanitizeFileName(filename) || 'output.docx';
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename=${downloadName}`);
        res.send(buffer);

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
    } finally {
        if (tempDocPath) {
            fs.promises.unlink(tempDocPath).catch(() => {});
        }
    }
});

// API: Get Templates
app.get('/api/templates', (req, res) => {
    // Use templateManager to get all templates
    const templates = templateManager.getAllTemplates();
    
    const result = templates.map(t => {
        const { name, isBuiltIn, description, ...style } = t;
        return {
            id: name,
            description: description || '(No description)',
            isBuiltIn,
            style
        };
    });
    res.json(result);
});

app.get('/api/templates/:id', (req, res) => {
    const tpl = templateManager.getTemplate(req.params.id);
    if (!tpl) return res.status(404).json({ error: 'Template not found' });
    res.json(tpl);
});

if (fs.existsSync(PUBLIC_DIR)) {
    app.get('/', (req, res) => {
        res.sendFile(path.join(PUBLIC_DIR, 'index.html'));
    });
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Backend server running on http://0.0.0.0:${PORT}`);
});