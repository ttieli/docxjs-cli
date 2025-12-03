#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const inquirer = require('inquirer');
const { generateDocx } = require('../lib/core');

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

    if (referenceDocPath && !templateName) {
        console.log(`🤖 Mode: Pure Clone`);
        templateName = 'default'; 
    } else if (templateName) {
        console.log(`📝 Mode: Template/Hybrid (${templateName})`);
    } else {
        // Interactive Mode
        console.log(`\n👋 欢迎使用 docxjs-cli 文档转换工具`);
        const choices = availableTemplates.map(key => ({
            name: `${key.padEnd(20)} - ${templates[key].description || "No desc"}`,
            value: key
        }));
        
        const answer = await inquirer.prompt([{
            type: 'list',
            name: 'selectedTemplate',
            message: '请选择目标文档格式 (Select Template):',
            choices: choices,
            pageSize: 15
        }]);
        templateName = answer.selectedTemplate;
    }

    if (!templates[templateName]) {
        console.error(`❌ Template '${templateName}' not found!`);
        if (!templates['default']) process.exit(1);
    }

    let currentStyle = { ...(templates[templateName] || templates['default']) };

    // --- 样式合并 (Python) ---
    if (referenceDocPath) {
        console.log(`🔍 Extracting styles from reference doc: ${referenceDocPath}...`);
        try {
            const extractorPath = path.join(__dirname, '..', 'style_extractor.py'); // Adjust path
            let pythonExecutable = 'python3'; 
            if (process.env.DOCXJS_PYTHON_PATH) {
                pythonExecutable = process.env.DOCXJS_PYTHON_PATH;
            } else if (fs.existsSync(path.join(__dirname, '..', 'venv'))) {
                pythonExecutable = path.join(__dirname, '..', 'venv', 'bin', 'python3');
            } else {
                const globalEnvPath = path.join(process.env.HOME || process.env.USERPROFILE, '.docxjs-cli-env', 'bin', 'python3');
                if (fs.existsSync(globalEnvPath)) { pythonExecutable = globalEnvPath; }
            }
            
            const pythonCmd = `"${pythonExecutable}" "${extractorPath}" "${referenceDocPath}"`;
            const stdout = execSync(pythonCmd, { encoding: 'utf-8' });
            const extractedStyles = JSON.parse(stdout);

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
        const buffer = await generateDocx(mdContent, currentStyle);
        fs.writeFileSync(outputPath, buffer);
        console.log(`✅ Success! Created ${outputPath}`);
    } catch (e) {
        console.error(`❌ Conversion failed: ${e.message}`);
        process.exit(1);
    }

})();
