const express = require('express');
const { generateDocx } = require('../lib/core');
const templateManager = require('../lib/template-manager');
const path = require('path');
const fs = require('fs');
const os = require('os');
const util = require('util');
const { execFile } = require('child_process');
const multer = require('multer');

const app = express();
const execFileAsync = util.promisify(execFile);
const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 5 * 1024 * 1024 } // 5MB reference doc limit
});

app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: true, limit: '2mb' }));

const PUBLIC_DIR = path.join(__dirname, '..', 'public');
if (fs.existsSync(PUBLIC_DIR)) {
    app.use(express.static(PUBLIC_DIR));
}

function sanitizeFileName(name) {
    if (!name) return 'output.docx';
    const cleaned = name.replace(/[\\/]+/g, '').replace(/[^a-zA-Z0-9._-]/g, '_');
    return cleaned || 'output.docx';
}

function resolvePythonPath() {
    if (process.env.DOCXJS_PYTHON_PATH) return process.env.DOCXJS_PYTHON_PATH;
    const projectVenv = path.join(__dirname, '..', 'venv', 'bin', 'python3');
    if (fs.existsSync(projectVenv)) return projectVenv;
    const globalEnv = path.join(process.env.HOME || process.env.USERPROFILE || '', '.docxjs-cli-env', 'bin', 'python3');
    if (fs.existsSync(globalEnv)) return globalEnv;
    return 'python3';
}

async function extractReferenceStyles(tempDocPath) {
    const pythonPath = resolvePythonPath();
    const extractorPath = path.join(__dirname, '..', 'style_extractor.py');
    const { stdout } = await execFileAsync(pythonPath, [extractorPath, tempDocPath], { timeout: 12000 });
    return JSON.parse(stdout);
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
                const extracted = await extractReferenceStyles(tempDocPath);
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
app.listen(PORT, () => {
    console.log(`Backend server running on http://localhost:${PORT}`);
});
