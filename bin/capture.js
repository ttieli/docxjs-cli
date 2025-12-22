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
try {
    ({ chromium } = require('playwright'));
} catch (err) {
    console.error('❌ Missing dependency "playwright". Please run `npm install` (installs playwright by default) and retry.');
    process.exit(1);
}

const ROOT = path.join(__dirname, '..');

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
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
    await page.goto(url, { waitUntil: 'networkidle' });

    // Inject content and switch mode
    await page.waitForSelector('#paper-container', { timeout: 10000 });
    await page.evaluate(({ mode, content }) => {
        if (mode === 'html') {
            if (typeof setMode === 'function') setMode('html');
            if (typeof window !== 'undefined') {
                window.htmlContent = content;
            }
            if (typeof triggerPreview === 'function') triggerPreview();
        } else {
            if (typeof setMode === 'function') setMode('markdown');
            if (typeof window !== 'undefined') {
                window.markdownContent = content;
            }
            if (typeof triggerPreview === 'function') triggerPreview();
        }
    }, { mode, content });

    // Wait for render target
    if (mode === 'html') {
        await page.waitForSelector('#htmlPreview', { state: 'visible', timeout: 10000 });
    } else {
        await page.waitForSelector('.docx-wrapper section.docx, #paper-container section.docx', { timeout: 15000 });
    }

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
    const content = fs.readFileSync(inputPath, 'utf-8');

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
