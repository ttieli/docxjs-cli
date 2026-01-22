#!/usr/bin/env node

/**
 * Advanced capture tool: renders the existing web UI headlessly (Playwright)
 * to export PNG/PDF that matches the in-app preview (Markdown or HTML mode).
 *
 * Usage:
 *   docxjs-capture --input path/to/file.md --png out.png --pdf out.pdf [--mode markdown|html] [--port 3000]
 *
 * Requirements: playwright dependency (installed by default in package.json).
 */

const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const http = require('http');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
let chromium;

// Simple MIME type lookup for common image formats
const MIME_TYPES = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.webp': 'image/webp',
    '.svg': 'image/svg+xml',
    '.bmp': 'image/bmp',
    '.ico': 'image/x-icon',
    '.tiff': 'image/tiff',
    '.tif': 'image/tiff'
};

function getMimeType(filePath) {
    const ext = path.extname(filePath).toLowerCase();
    return MIME_TYPES[ext] || 'image/png';
}
try {
    ({ chromium } = require('playwright'));
} catch (err) {
    console.error('❌ Missing dependency "playwright". Please run `npm install` (installs playwright by default) and retry.');
    process.exit(1);
}

const ROOT = path.join(__dirname, '..');

/**
 * Convert local image paths in markdown to base64 data URIs
 * @param {string} content - Markdown content
 * @param {string} baseDir - Base directory of the markdown file
 * @returns {string} - Markdown with local images converted to base64
 */
function convertLocalImagesToBase64(content, baseDir) {
    // Match markdown image syntax: ![alt](path) and HTML img tags
    const mdImageRegex = /!\[([^\]]*)\]\(([^)]+)\)/g;
    const htmlImageRegex = /<img\s+[^>]*src=["']([^"']+)["'][^>]*>/gi;

    let result = content;

    // Process markdown images
    result = result.replace(mdImageRegex, (match, alt, imagePath) => {
        const base64Data = getImageAsBase64(imagePath, baseDir);
        if (base64Data) {
            return `![${alt}](${base64Data})`;
        }
        return match; // Return original if conversion failed
    });

    // Process HTML img tags
    result = result.replace(htmlImageRegex, (match, imagePath) => {
        const base64Data = getImageAsBase64(imagePath, baseDir);
        if (base64Data) {
            return match.replace(imagePath, base64Data);
        }
        return match;
    });

    return result;
}

/**
 * Convert a single image path to base64 data URI
 * @param {string} imagePath - Image path (relative or absolute)
 * @param {string} baseDir - Base directory for resolving relative paths
 * @returns {string|null} - Base64 data URI or null if not a local file
 */
function getImageAsBase64(imagePath, baseDir) {
    // Skip URLs (http, https, data URIs)
    if (/^(https?:|data:)/i.test(imagePath)) {
        return null;
    }

    // Decode URL-encoded path (e.g., %20 -> space)
    let decodedPath = decodeURIComponent(imagePath);

    // Resolve relative paths
    let absolutePath;
    if (path.isAbsolute(decodedPath)) {
        absolutePath = decodedPath;
    } else {
        absolutePath = path.resolve(baseDir, decodedPath);
    }

    // Check if file exists
    if (!fs.existsSync(absolutePath)) {
        console.warn(`⚠️  Image not found: ${absolutePath}`);
        return null;
    }

    try {
        const imageBuffer = fs.readFileSync(absolutePath);
        const mimeType = getMimeType(absolutePath);
        const base64 = imageBuffer.toString('base64');
        return `data:${mimeType};base64,${base64}`;
    } catch (err) {
        console.warn(`⚠️  Failed to read image: ${absolutePath} - ${err.message}`);
        return null;
    }
}

function waitForServer(port, timeoutMs = 15000) {
    const start = Date.now();
    return new Promise((resolve, reject) => {
        const attempt = () => {
            const req = http.get({ host: '127.0.0.1', port, path: '/api/templates', timeout: 2000 }, res => {
                res.resume();
                resolve(true);
            });
            req.on('error', () => {
                if (Date.now() - start > timeoutMs) return reject(new Error(`Server not ready on port ${port}`));
                setTimeout(attempt, 500);
            });
        };
        attempt();
    });
}

async function startServer(port) {
    const proc = spawn(process.execPath, [path.join(ROOT, 'server', 'app.js')], {
        cwd: ROOT,
        env: { ...process.env, PORT: port },
        stdio: 'ignore'
    });
    await waitForServer(port);
    return proc;
}

async function capturePage({ url, mode, content, pngPath, pdfPath }) {
    console.log('   [1/6] Launching browser...');
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });

    console.log('   [2/6] Loading preview page...');
    await page.goto(url, { waitUntil: 'networkidle' });

    // Wait for initial page load and initial preview to complete
    await page.waitForSelector('#paper-container', { timeout: 10000 });

    // Wait for the initial preview to complete (status text changes from "Updating Preview..." to "Preview Updated")
    try {
        await page.waitForFunction(() => {
            const status = document.getElementById('statusText');
            return status && status.textContent === 'Preview Updated';
        }, { timeout: 30000 });
    } catch (e) {
        // Initial preview timeout is fine, we'll inject our content anyway
    }

    console.log('   [3/6] Injecting content...');
    // Clear existing content, inject new content and trigger preview
    await page.evaluate(({ mode, content }) => {
        if (mode === 'html') {
            if (typeof setMode === 'function') setMode('html');
            // Clear existing preview
            const htmlPreview = document.getElementById('htmlPreview');
            if (htmlPreview) htmlPreview.innerHTML = '';

            // Update status to indicate loading
            const statusHTML = document.getElementById('statusTextHTML');
            if (statusHTML) statusHTML.textContent = 'Loading...';

            if (typeof window !== 'undefined') {
                window.htmlContent = content;
            }
            if (typeof triggerPreview === 'function') triggerPreview();
        } else {
            if (typeof setMode === 'function') setMode('markdown');
            // Clear existing preview to ensure we wait for new content
            const paper = document.getElementById('paper-container');
            if (paper) paper.innerHTML = '';

            // Update status to indicate loading
            const status = document.getElementById('statusText');
            if (status) status.textContent = 'Loading...';

            if (typeof window !== 'undefined') {
                window.markdownContent = content;
            }
            if (typeof triggerPreview === 'function') triggerPreview();
        }
    }, { mode, content });

    // Wait for the debounce timer (600ms) plus rendering time
    // triggerPreview has a 600ms debounce, so we need to wait for it
    await page.waitForTimeout(1000);

    console.log('   [4/6] Generating preview (this may take a while for large documents)...');
    // Wait for render target to appear (after clearing it above)
    if (mode === 'html') {
        await page.waitForSelector('#htmlPreview iframe', { state: 'visible', timeout: 120000 });
    } else {
        // Wait for status text to indicate preview is complete with progress polling
        const startTime = Date.now();
        const maxWait = 120000; // 2 minutes max
        let lastStatus = '';

        while (Date.now() - startTime < maxWait) {
            const status = await page.evaluate(() => {
                const el = document.getElementById('statusText');
                return el ? el.textContent : '';
            });

            if (status !== lastStatus) {
                if (status && status !== 'Loading...') {
                    // Show status updates
                    process.stdout.write(`\r   [4/6] ${status}...                    `);
                }
                lastStatus = status;
            }

            if (status === 'Preview Updated' || status.includes('Error')) {
                process.stdout.write('\n');
                break;
            }

            await page.waitForTimeout(500);
        }

        // Then wait for the docx content to be rendered
        try {
            await page.waitForSelector('.docx-wrapper section.docx, #paper-container section.docx', { timeout: 30000 });
        } catch (e) {
            console.log('   ⚠️  Content selector timeout, attempting capture anyway...');
        }
    }

    console.log('   [5/6] Rendering complete, capturing image...');
    // Additional wait to ensure content is fully rendered
    await page.waitForTimeout(1000);

    // Helper run html2canvas/jsPDF inside page and return base64
    const captureElement = async (type) => {
        const result = await page.evaluate(async (type) => {
            const target = (typeof getCaptureElement === 'function') ? getCaptureElement() : null;
            if (!target) throw new Error('Preview content not found');
            const canvas = await html2canvas(target, {
                scale: 2,
                backgroundColor: '#ffffff',
                useCORS: true,
                windowWidth: target.scrollWidth,
                windowHeight: target.scrollHeight
            });
            if (type === 'png') {
                return canvas.toDataURL('image/png');
            }
            // PDF
            const pdf = new jspdf.jsPDF('p', 'pt', 'a4');
            const imgData = canvas.toDataURL('image/png');
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = pdf.internal.pageSize.getHeight();
            const imgWidth = pageWidth;
            const imgHeight = canvas.height * imgWidth / canvas.width;
            let heightLeft = imgHeight;
            let position = 0;
            pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;
            while (heightLeft > 0) {
                position = heightLeft - imgHeight;
                pdf.addPage();
                pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
                heightLeft -= pageHeight;
            }
            return pdf.output('datauristring');
        }, type);
        return result;
    };

    if (pngPath) {
        const dataUrl = await captureElement('png');
        const base64 = dataUrl.split(',')[1];
        fs.writeFileSync(pngPath, Buffer.from(base64, 'base64'));
        console.log(`✅ PNG exported to ${pngPath}`);
    }

    if (pdfPath) {
        const dataUri = await captureElement('pdf');
        const base64 = dataUri.split(',')[1];
        fs.writeFileSync(pdfPath, Buffer.from(base64, 'base64'));
        console.log(`✅ PDF exported to ${pdfPath}`);
    }

    console.log('   [6/6] Cleaning up...');
    await browser.close();
}

(async () => {
    const argv = yargs(hideBin(process.argv))
        .option('input', { alias: 'i', type: 'string', demandOption: true, describe: 'Input file (.md/.txt/.html)' })
        .option('mode', { type: 'string', choices: ['markdown', 'html'], describe: 'Force mode; default auto by extension' })
        .option('png', { type: 'string', describe: 'Output PNG path' })
        .option('pdf', { type: 'string', describe: 'Output PDF path' })
        .option('port', { type: 'number', default: 3000, describe: 'Port to run the preview server' })
        .check(args => {
            if (!args.png && !args.pdf) throw new Error('Please specify --png and/or --pdf output path.');
            return true;
        })
        .help()
        .argv;

    const inputPath = path.resolve(argv.input);
    if (!fs.existsSync(inputPath)) {
        console.error(`Input not found: ${inputPath}`);
        process.exit(1);
    }
    const ext = path.extname(inputPath).toLowerCase();
    const mode = argv.mode || ((ext === '.html' || ext === '.htm') ? 'html' : 'markdown');
    const baseDir = path.dirname(inputPath);
    let content = fs.readFileSync(inputPath, 'utf-8');

    // Convert local images to base64 data URIs
    content = convertLocalImagesToBase64(content, baseDir);

    const port = argv.port;
    let serverProc;
    try {
        console.log(`ℹ️  Starting preview server on port ${port}...`);
        serverProc = await startServer(port);
        console.log(`ℹ️  Rendering ${mode.toUpperCase()} and capturing...`);
        await capturePage({
            url: `http://127.0.0.1:${port}/`,
            mode,
            content,
            pngPath: argv.png ? path.resolve(argv.png) : null,
            pdfPath: argv.pdf ? path.resolve(argv.pdf) : null
        });
    } catch (err) {
        console.error(`❌ Capture failed: ${err.message}`);
        process.exitCode = 1;
    } finally {
        if (serverProc) {
            serverProc.kill('SIGTERM');
        }
    }
})();
