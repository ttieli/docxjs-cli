const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const mammoth = require('mammoth');
const TurndownService = require('turndown');
const { gfm } = require('turndown-plugin-gfm');
const { generateDocx } = require('../lib/core');
const { generatePreviewBuffer } = require('../lib/sample-generator');
const templateManager = require('../lib/template-manager');
const { extractStyles } = require('../lib/python-bridge');

// Helper: Merge Styles (copied from server/app.js)
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

function createWindow() {
    // 1. Create Splash Window
    const splash = new BrowserWindow({
        width: 400,
        height: 300,
        transparent: false,
        frame: false,
        alwaysOnTop: true,
        center: true,
        resizable: false,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true
        }
    });
    splash.loadFile(path.join(__dirname, 'splash.html'));

    // 2. Create Main Window (Hidden initially)
    const win = new BrowserWindow({
        width: 1200,
        height: 800,
        show: false, // Hide initially
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true
        }
    });

    // Load the local index.html
    win.loadFile(path.join(__dirname, '../public/index.html'));

    // 3. Swap windows when main is ready
    win.once('ready-to-show', () => {
        splash.destroy();
        win.show();
    });

    // Open DevTools (optional, for debugging)
    // win.webContents.openDevTools();
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

// --- IPC Handlers ---

// 1. Get Templates
ipcMain.handle('get-templates', async () => {
    const templates = templateManager.getAllTemplates();
    return templates.map(t => {
        const { name, isBuiltIn, description, ...style } = t;
        return {
            id: name,
            description: description || '(No description)',
            isBuiltIn,
            style
        };
    });
});

// 2. Import Docx -> Markdown
ipcMain.handle('import-docx', async (event, fileBuffer) => {
    try {
        // Buffer comes as Uint8Array from renderer
        const buffer = Buffer.from(fileBuffer);
        
        // Optimize: Ignore images to avoid long base64 strings
        const options = {
            ignoreImage: true 
        };
        
        const { value: html } = await mammoth.convertToHtml({ buffer }, options);
        
        const turndownService = new TurndownService({
            headingStyle: 'atx',
            codeBlockStyle: 'fenced'
        });
        turndownService.use(gfm);

        // Explicitly discard images in Turndown to prevent base64 leaks or placeholders
        turndownService.addRule('removeImages', {
            filter: 'img',
            replacement: function (content, node) {
                // Return empty string to hide, or '[图片]' to show placeholder
                return '[图片]'; 
            }
        });
        
        const markdown = turndownService.turndown(html);
        return { markdown };
    } catch (error) {
        console.error('Import error:', error);
        throw new Error('Failed to parse Word document: ' + error.message);
    }
});

// 3. Convert
ipcMain.handle('convert', async (event, { markdown, templateName, styleConfig, referenceDocBuffer, filename }) => {
    let tempDocPath = null;
    try {
        if (!markdown || typeof markdown !== 'string' || markdown.trim().length === 0) {
            throw new Error('Markdown content required');
        }

        // Determine base style
        let finalStyle = {};
        if (styleConfig && Object.keys(styleConfig).length > 0) {
            finalStyle = styleConfig;
        } else {
            const template = templateManager.getTemplate(templateName) || templateManager.getTemplate('default');
            if (template) {
                const { name, isBuiltIn, ...styleData } = template;
                finalStyle = styleData;
            }
        }

        // Handle Reference Doc
        if (referenceDocBuffer) {
            tempDocPath = path.join(os.tmpdir(), `reference-${Date.now()}-${Math.random().toString(16).slice(2)}.docx`);
            await fs.promises.writeFile(tempDocPath, Buffer.from(referenceDocBuffer));
            try {
                const extracted = await extractStyles(tempDocPath);
                if (!extracted.error) {
                    finalStyle = mergeStyles(finalStyle, extracted);
                    if (!extracted.fontH1 && extracted.detailed_styles_info) {
                        const genericHeading = extracted.detailed_styles_info.find(s => s.name === 'Heading');
                        if (genericHeading && genericHeading.font_name) finalStyle.fontHeader1 = genericHeading.font_name;
                    }
                }
            } catch (err) {
                console.warn(`⚠️ Reference style extraction failed: ${err.message}`);
            }
        }

        const buffer = await generateDocx(markdown, finalStyle);
        
        // In Electron, we can return the buffer back to renderer to trigger download, 
        // OR we can save it directly using dialog.
        // For consistency with Web flow (Browser triggers download), returning Buffer is fine.
        // The renderer will create a Blob and download it.
        return buffer; // Returns as Uint8Array

    } catch (error) {
        console.error('Convert error:', error);
        throw new Error(error.message);
    } finally {
        if (tempDocPath) {
            fs.promises.unlink(tempDocPath).catch(() => {});
        }
    }
});

// 5. Print to PDF
ipcMain.handle('print-to-pdf', async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender);
    if (!win) return false;

    try {
        const data = await win.webContents.printToPDF({
            printBackground: true,
            pageSize: 'A4',
            marginsType: 0 // No margins, let the content handle it
        });
        
        const { filePath } = await dialog.showSaveDialog(win, {
            title: 'Export PDF',
            defaultPath: 'document.pdf',
            filters: [{ name: 'PDF Files', extensions: ['pdf'] }]
        });

        if (filePath) {
            await fs.promises.writeFile(filePath, data);
            return true;
        }
        return false; // User cancelled
    } catch (error) {
        console.error('PDF export error:', error);
        throw new Error(error.message);
    }
});

