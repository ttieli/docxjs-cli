#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const inquirer = require('inquirer');
const { generateDocx } = require('../lib/core');
const { extractStyles } = require('../lib/python-bridge');

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
    const rawArgv = yargs(hideBin(process.argv)).argv;
    const templates = loadTemplates(rawArgv.config);
    const availableTemplates = Object.keys(templates);

    const argv = yargs(hideBin(process.argv))
        .usage('Usage: $0 <input.md> -o <output.docx> [options]')
        .command('$0 <input>', 'Convert Markdown to Docx')
        .option('output', { alias: 'o', type: 'string', default: 'output.docx' })
        .option('template', { alias: 't', type: 'string', description: `Template name.` })
        .option('config', { alias: 'c', type: 'string' })
        .option('reference-doc', { alias: 'r', type: 'string' })
        .demandCommand(1)
        .help()
        .argv;

    const inputPath = argv.input;
    const outputPath = argv.output;
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
        const buffer = await generateDocx(mdContent, currentStyle, baseDir);
        fs.writeFileSync(outputPath, buffer);
        console.log(`✅ Success! Created ${outputPath}`);
    } catch (e) {
        console.error(`❌ Conversion failed: ${e.message}`);
        process.exit(1);
    }

})();
