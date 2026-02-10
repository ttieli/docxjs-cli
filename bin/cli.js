#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const inquirer = require('inquirer');
const { generateDocx } = require('../lib/core');
const { extractStyles } = require('../lib/python-bridge');
const { normalizeStyleConfig } = require('../lib/style-normalizer');

// --- 0. 加载配置 ---
function loadTemplates(customConfigPath) {
    let templates = {};
    // Adjust path to point to project root's templates dir relative to bin/cli.js
    const templatesDir = path.join(__dirname, '..', 'templates');

    if (fs.existsSync(templatesDir)) {
        try {
            const files = fs.readdirSync(templatesDir).filter(file => file.toLowerCase().endsWith('.json'));
            files.sort();
            files.forEach(file => {
                const fullPath = path.join(templatesDir, file);
                try {
                    const fileContent = JSON.parse(fs.readFileSync(fullPath, 'utf-8'));
                    templates = { ...templates, ...fileContent };
                } catch (e) { console.warn(`⚠️ Warning: Failed to parse template '${file}': ${e.message}`); }
            });
        } catch (e) { console.error(`❌ Error reading templates directory: ${e.message}`); }
    }

    if (customConfigPath) {
        const absPath = path.resolve(process.cwd(), customConfigPath);
        if (fs.existsSync(absPath)) {
            try {
                const userTemplates = JSON.parse(fs.readFileSync(absPath, 'utf-8'));
                templates = { ...templates, ...userTemplates };
            } catch (e) { console.error(`❌ Failed to load custom config ${customConfigPath}:`, e.message); }
        } else { console.warn(`⚠️ Custom config file not found: ${customConfigPath}`); }
    }
    return templates;
}

// --- 主逻辑 ---
(async () => {
    const argv = yargs(hideBin(process.argv))
        .usage('Usage: $0 <input.md> -o <output.docx> [options]')
        .command('$0 [input]', 'Convert Markdown to Docx (or PDF to PNG with --image)')
        .positional('input', { describe: 'Input Markdown file', type: 'string' })
        // --- Output ---
        .option('output', { alias: 'o', type: 'string', describe: 'Output Docx file path' })
        .option('pdf', { type: 'boolean', description: 'Also export as PDF (requires LibreOffice soffice)' })
        .option('image', { type: 'boolean', description: 'Also export as PNG image (requires playwright)' })
        // --- Template & Config ---
        .option('template', { alias: 't', type: 'string', description: 'Template name (use -l to list)' })
        .option('config', { alias: 'c', type: 'string', description: 'Custom configuration JSON file' })
        .option('reference-doc', { alias: 'r', type: 'string', description: 'Reference Docx for style extraction' })
        .option('list-templates', { alias: 'l', type: 'boolean', description: 'List all available templates and exit' })
        // --- Style overrides ---
        .option('font', { type: 'string', description: 'Override body font (e.g. "微软雅黑", "Calibri")' })
        .option('font-size', { type: 'number', description: 'Override body font size in half-points (24=12pt)' })
        .option('line-spacing', { type: 'number', description: 'Override line spacing in twips (240=single, 360=1.5x, 480=double)' })
        // --- Behavior ---
        .option('quiet', { alias: 'q', type: 'boolean', description: 'Suppress verbose token-level debug output' })
        .option('open', { type: 'boolean', description: 'Open the output file after generation' })
        .example('$0 report.md', 'Convert with default template, auto-named output')
        .example('$0 report.md -o report.docx -t "政府公文 (红头)"', 'Convert with specific template')
        .example('$0 report.md -o out.docx --font "宋体" --font-size 24', 'Override font styles')
        .example('$0 report.md -o out.docx --pdf --image', 'Export docx + pdf + png')
        .example('$0 report.md -o out.docx -q --open', 'Quiet mode, open when done')
        .example('$0 -l', 'List all available templates')
        .help()
        .version()
        .parse(); // Use parse() instead of .argv to avoid premature exit issues with help

    const quiet = argv.quiet || false;

    // --- List templates mode ---
    if (argv.listTemplates || argv.l) {
        const templates = loadTemplates(argv.config);
        const names = Object.keys(templates).filter(k => k !== 'default');
        console.log('\n📋 Available templates:\n');
        console.log('  Name                                      Description');
        console.log('  ─────────────────────────────────────────  ──────────────────────────────────');
        names.forEach(name => {
            const desc = templates[name].description || '';
            const paddedName = name.padEnd(42);
            console.log(`  ${paddedName}  ${desc}`);
        });
        console.log(`\n  Total: ${names.length} templates`);
        console.log('  Use: docxjs input.md -t "<template name>"\n');
        return;
    }

    // If no input provided, show help
    if (!argv.input) {
        yargs(hideBin(process.argv)).showHelp();
        return;
    }

    const templates = loadTemplates(argv.config);
    const availableTemplates = Object.keys(templates);

    const inputPath = argv.input;
    let outputPath = argv.output;
    const imageOnly = argv.image && !argv.output; // --image without -o means PNG only

    // Generate timestamp for auto-naming
    const absInputPath = path.resolve(inputPath);
    const dirname = path.dirname(absInputPath);
    const ext = path.extname(absInputPath);
    const basename = path.basename(absInputPath, ext);
    const now = new Date();
    const timestamp = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}-${String(now.getSeconds()).padStart(2, '0')}`;

    // --- PDF 文件特殊处理 ---
    const inputExt = ext.toLowerCase();

    if (inputExt === '.pdf') {
        if (argv.image) {
            // PDF → 长图
            const { pdfToImage } = require('../lib/pdf-to-image');
            const imagePath = argv.output
                ? argv.output.replace(/\.\w+$/, '.png')
                : path.join(dirname, `${basename}_${timestamp}.png`);

            if (!quiet) console.log('🖼️  Converting PDF to PNG...');
            try {
                await pdfToImage(absInputPath, imagePath);
                console.log(`✅ PNG created: ${imagePath}`);
            } catch (e) {
                console.error(`❌ PDF to image failed: ${e.message}`);
                process.exit(1);
            }
            return;
        } else {
            // PDF 不支持直接转 Docx
            console.error('❌ PDF 文件不支持直接转换为 Docx。');
            console.error('   请使用 --image 选项导出为长图：');
            console.error(`   docxjs "${inputPath}" --image`);
            process.exit(1);
        }
    }

    // If --image only (no -o), skip docx and export PNG directly
    if (imageOnly) {
        const imagePath = path.join(dirname, `${basename}_${timestamp}.png`);
        if (!quiet) console.log('🖼️  Exporting to PNG only (requires playwright)...');
        const captureScript = path.join(__dirname, 'capture.js');
        try {
            execSync(`node "${captureScript}" --input "${inputPath}" --png "${imagePath}"`, { stdio: 'inherit' });
            console.log(`✅ PNG created: ${imagePath}`);
        } catch (e) {
            console.error(`❌ Image export failed. Please ensure playwright is installed.`);
            process.exit(1);
        }
        return;
    }

    if (!outputPath) {
        outputPath = path.join(dirname, `${basename}_${timestamp}.docx`);
    }

    let templateName = argv.template;
    const referenceDocPath = argv['reference-doc'];

    // Logic:
    // 1. If input file exists AND template is specified -> Use it.
    // 2. If input file exists AND NO template specified -> Use 'default' (Silent).

    if (!templateName) {
        if (!quiet) console.log(`ℹ️  No template specified. Using 'default'.`);
        templateName = 'default';
    } else {
        if (!quiet) console.log(`📝 Template: ${templateName}`);
    }

    // If template not found, try fallback
    if (!templates[templateName]) {
        console.warn(`⚠️ Template '${templateName}' not found. Falling back to 'default'.`);
        templateName = 'default';
    }

    let currentStyle = { ...(templates[templateName] || templates['default']) };

    // --- CLI style overrides ---
    if (argv.font) {
        currentStyle.fontMain = argv.font;
        if (!quiet) console.log(`  ↳ Font override: ${argv.font}`);
    }
    if (argv.fontSize !== undefined) {
        currentStyle.fontSizeMain = argv.fontSize;
        if (!quiet) console.log(`  ↳ Font size override: ${argv.fontSize} half-points (${argv.fontSize / 2}pt)`);
    }
    if (argv.lineSpacing !== undefined) {
        currentStyle.lineSpacing = argv.lineSpacing;
        if (!quiet) console.log(`  ↳ Line spacing override: ${argv.lineSpacing} twips`);
    }

    // --- 样式合并 (Python Bridge) ---
    if (referenceDocPath) {
        if (!quiet) console.log(`🔍 Extracting styles from reference doc: ${referenceDocPath}...`);
        try {
            // Use the unified bridge
            const extractedStyles = await extractStyles(referenceDocPath);

            if (extractedStyles.error) {
                console.warn(`⚠️ Style extraction failed: ${extractedStyles.error}`);
            } else {
                if (!quiet) console.log(`✅ Styles extracted successfully! Merging...`);
                Object.keys(extractedStyles).forEach(key => {
                    if (extractedStyles[key] !== null && extractedStyles[key] !== undefined && key !== 'detailed_styles_info') {
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

    // --- 调用 Core ---
    try {
        const mdContent = fs.readFileSync(inputPath, 'utf-8');
        const baseDir = path.dirname(path.resolve(inputPath));
        const buffer = await generateDocx(mdContent, normalizeStyleConfig(currentStyle), baseDir, { quiet });
        fs.writeFileSync(outputPath, buffer);
        console.log(`✅ Created: ${outputPath}`);

        if (argv.pdf) {
            if (!quiet) console.log('📄 Converting to PDF (requires LibreOffice)...');
            const outputDir = path.dirname(outputPath);
            try {
                execSync(`soffice --headless --convert-to pdf "${outputPath}" --outdir "${outputDir}"`, { stdio: 'pipe' });
                const pdfPath = outputPath.replace('.docx', '.pdf');
                console.log(`✅ PDF created: ${pdfPath}`);
            } catch (e) {
                console.error(`❌ PDF conversion failed. Please ensure LibreOffice (soffice) is installed and in your PATH.`);
            }
        }

        if (argv.image) {
            // When both -o and --image are specified, also export PNG
            const imagePath = outputPath.replace(/\.docx$/i, '.png');
            if (!quiet) console.log('🖼️  Also exporting to PNG...');
            const captureScript = path.join(__dirname, 'capture.js');
            try {
                execSync(`node "${captureScript}" --input "${inputPath}" --png "${imagePath}"`, { stdio: 'inherit' });
                console.log(`✅ PNG created: ${imagePath}`);
            } catch (e) {
                console.error(`❌ Image export failed. Please ensure playwright is installed.`);
            }
        }

        // --- Open output file ---
        if (argv.open) {
            const fileToOpen = outputPath;
            try {
                if (process.platform === 'darwin') {
                    execSync(`open "${fileToOpen}"`, { stdio: 'ignore' });
                } else if (process.platform === 'win32') {
                    execSync(`start "" "${fileToOpen}"`, { stdio: 'ignore' });
                } else {
                    execSync(`xdg-open "${fileToOpen}"`, { stdio: 'ignore' });
                }
            } catch (e) {
                console.warn(`⚠️ Could not open file: ${e.message}`);
            }
        }

    } catch (e) {
        console.error(`❌ Conversion failed: ${e.message}`);
        process.exit(1);
    }

})();
