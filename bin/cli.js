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
        .command('$0 <input>', 'Convert Markdown to Docx')
        .positional('input', { describe: 'Input Markdown file', type: 'string' })
        .option('output', { alias: 'o', type: 'string', describe: 'Output Docx file path' })
        .option('template', { alias: 't', type: 'string', description: `Template name.` })
        .option('config', { alias: 'c', type: 'string', description: 'Custom configuration file' })
        .option('reference-doc', { alias: 'r', type: 'string', description: 'Reference Docx for style extraction' })
        .option('pdf', { type: 'boolean', description: 'Export as PDF using LibreOffice (soffice)' })
        .option('image', { type: 'boolean', description: 'Export as PNG image only (supports local images, requires playwright)' })
        .help()
        .version()
        .parse(); // Use parse() instead of .argv to avoid premature exit issues with help

    // If help or version was requested, yargs handles it and exits.
    // If we are here, we might have arguments or not.
    // Since we used command('$0 <input>'), yargs expects input. 
    // However, without .demandCommand(1) strict, it might pass through.
    // But we want to ensure input is there.
    
    if (!argv.input) {
        yargs.showHelp();
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

    // If --image only (no -o), skip docx and export PNG directly
    if (imageOnly) {
        const imagePath = path.join(dirname, `${basename}_${timestamp}.png`);
        console.log('🖼️  Exporting to PNG only (requires playwright)...');
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
    // 3. Interactive mode logic (currently unreachable due to demandCommand(1) but kept for safety).

    if (!templateName) {
        console.log(`ℹ️  No template specified. Using 'default' (Minimalist Black).`);
        templateName = 'default';
    } else {
        console.log(`📝 Mode: Template (${templateName})`);
    }

    // If template not found, try fallback
    if (!templates[templateName]) {
        console.warn(`⚠️ Template '${templateName}' not found. Falling back to 'default'.`);
        templateName = 'default';
    }

    let currentStyle = { ...(templates[templateName] || templates['default']) };

    // --- 样式合并 (Python Bridge) ---
    if (referenceDocPath) {
        console.log(`🔍 Extracting styles from reference doc: ${referenceDocPath}...`);
        try {
            // Use the unified bridge
            const extractedStyles = await extractStyles(referenceDocPath);

            if (extractedStyles.error) {
                console.warn(`⚠️ Style extraction failed: ${extractedStyles.error}`);
            } else {
                console.log(`✅ Styles extracted successfully! Merging...`);
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
        const buffer = await generateDocx(mdContent, normalizeStyleConfig(currentStyle), baseDir);
        fs.writeFileSync(outputPath, buffer);
        console.log(`✅ Success! Created ${outputPath}`);

        if (argv.pdf) {
            console.log('📄 Converting to PDF (requires LibreOffice)...');
            const outputDir = path.dirname(outputPath);
            try {
                execSync(`soffice --headless --convert-to pdf "${outputPath}" --outdir "${outputDir}"`, { stdio: 'pipe' });
                console.log(`✅ PDF created: ${outputPath.replace('.docx', '.pdf')}`);
            } catch (e) {
                console.error(`❌ PDF conversion failed. Please ensure LibreOffice (soffice) is installed and in your PATH.`);
            }
        }

        if (argv.image) {
            // When both -o and --image are specified, also export PNG
            const imagePath = outputPath.replace(/\.docx$/i, '.png');
            console.log('🖼️  Also exporting to PNG...');
            const captureScript = path.join(__dirname, 'capture.js');
            try {
                execSync(`node "${captureScript}" --input "${inputPath}" --png "${imagePath}"`, { stdio: 'inherit' });
            } catch (e) {
                console.error(`❌ Image export failed. Please ensure playwright is installed.`);
            }
        }

    } catch (e) {
        console.error(`❌ Conversion failed: ${e.message}`);
        process.exit(1);
    }

})();
