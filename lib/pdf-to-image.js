/**
 * PDF to Long Image Converter
 * 使用 pdf.js 在浏览器中渲染 PDF，然后截图导出为长图
 */

const path = require('path');
const fs = require('fs');
const http = require('http');

let chromium;

function loadPlaywright() {
    if (!chromium) {
        try {
            ({ chromium } = require('playwright'));
        } catch (err) {
            throw new Error('Missing dependency "playwright". Run `npm install playwright`');
        }
    }
}

/**
 * 启动简单的静态文件服务器
 */
function startStaticServer(dir, port) {
    return new Promise((resolve) => {
        const server = http.createServer((req, res) => {
            let filePath = path.join(dir, decodeURIComponent(req.url));
            if (!fs.existsSync(filePath)) {
                res.writeHead(404);
                res.end('Not found');
                return;
            }
            const ext = path.extname(filePath).toLowerCase();
            const mimeTypes = {
                '.html': 'text/html',
                '.js': 'application/javascript',
                '.css': 'text/css',
                '.pdf': 'application/pdf',
                '.png': 'image/png',
                '.jpg': 'image/jpeg'
            };
            res.writeHead(200, { 'Content-Type': mimeTypes[ext] || 'application/octet-stream' });
            fs.createReadStream(filePath).pipe(res);
        });
        server.listen(port, '127.0.0.1', () => resolve(server));
    });
}

/**
 * 将 PDF 转换为长图
 * @param {string} pdfPath - PDF 文件路径
 * @param {string} outputPath - 输出 PNG 路径
 * @param {object} options - 可选配置
 * @param {number} options.scale - 缩放比例，默认 2（高清）
 */
async function pdfToImage(pdfPath, outputPath, options = {}) {
    loadPlaywright();

    const scale = options.scale || 2;
    const absPdfPath = path.resolve(pdfPath);

    if (!fs.existsSync(absPdfPath)) {
        throw new Error(`PDF file not found: ${absPdfPath}`);
    }

    // 使用随机端口
    const port = 9000 + Math.floor(Math.random() * 1000);
    const pdfDir = path.dirname(absPdfPath);
    const pdfName = path.basename(absPdfPath);

    console.log('   [1/5] Starting local server...');
    const server = await startStaticServer(pdfDir, port);

    console.log('   [2/5] Launching browser...');
    const browser = await chromium.launch({
        headless: true,
        args: ['--disable-web-security'] // 允许本地文件访问
    });

    try {
        const context = await browser.newContext({
            viewport: { width: 1200, height: 800 },
            deviceScaleFactor: scale
        });
        const page = await context.newPage();

        console.log('   [3/5] Loading PDF with pdf.js...');

        // 使用 Mozilla pdf.js viewer (CDN)
        const pdfJsViewerUrl = `https://mozilla.github.io/pdf.js/web/viewer.html?file=${encodeURIComponent(`http://127.0.0.1:${port}/${encodeURIComponent(pdfName)}`)}`;

        await page.goto(pdfJsViewerUrl, { waitUntil: 'networkidle', timeout: 60000 });

        // 等待 PDF 加载完成
        await page.waitForSelector('#viewer .page', { timeout: 30000 });

        // 等待所有页面渲染
        await page.waitForTimeout(3000);

        // 获取总页数
        const pageCount = await page.evaluate(() => {
            return PDFViewerApplication.pagesCount || 1;
        });
        console.log(`   [4/5] Rendering ${pageCount} pages...`);

        // 逐页滚动确保所有页面都渲染
        for (let i = 1; i <= pageCount; i++) {
            await page.evaluate((pageNum) => {
                PDFViewerApplication.page = pageNum;
            }, i);
            await page.waitForTimeout(500);
            process.stdout.write(`\r   [4/5] Rendering page ${i}/${pageCount}...`);
        }
        process.stdout.write('\n');

        // 等待所有页面完全渲染
        await page.waitForTimeout(2000);

        // 回到第一页
        await page.evaluate(() => {
            PDFViewerApplication.page = 1;
        });
        await page.waitForTimeout(500);

        // 使用 canvas 拼接所有页面
        const screenshot = await page.evaluate(async (totalPages) => {
            // 获取所有已渲染的页面 canvas
            const pages = document.querySelectorAll('#viewer .page');
            if (pages.length === 0) {
                throw new Error('No pages found');
            }

            // 计算总高度和宽度
            let totalHeight = 0;
            let maxWidth = 0;
            const pageData = [];

            for (const pageDiv of pages) {
                const canvas = pageDiv.querySelector('canvas');
                if (canvas) {
                    totalHeight += canvas.height + 10; // 10px gap between pages
                    maxWidth = Math.max(maxWidth, canvas.width);
                    pageData.push({
                        canvas: canvas,
                        width: canvas.width,
                        height: canvas.height
                    });
                }
            }

            // 创建合并后的 canvas
            const mergedCanvas = document.createElement('canvas');
            mergedCanvas.width = maxWidth;
            mergedCanvas.height = totalHeight;
            const ctx = mergedCanvas.getContext('2d');

            // 白色背景
            ctx.fillStyle = 'white';
            ctx.fillRect(0, 0, maxWidth, totalHeight);

            // 绘制每一页
            let yOffset = 0;
            for (const pd of pageData) {
                const xOffset = (maxWidth - pd.width) / 2; // 居中
                ctx.drawImage(pd.canvas, xOffset, yOffset);
                yOffset += pd.height + 10;
            }

            return mergedCanvas.toDataURL('image/png');
        }, pageCount);

        // 提取 base64 数据
        const base64Data = screenshot.split(',')[1];
        const imageBuffer = Buffer.from(base64Data, 'base64');

        console.log('   [5/5] Saving image...');
        fs.writeFileSync(outputPath, imageBuffer);

    } finally {
        await browser.close();
        server.close();
    }
}

module.exports = { pdfToImage };
