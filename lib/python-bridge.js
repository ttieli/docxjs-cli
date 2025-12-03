const path = require('path');
const fs = require('fs');
const util = require('util');
const { execFile } = require('child_process');

const execFileAsync = util.promisify(execFile);

/**
 * Helper to resolve the Python interpreter path in standard environments.
 * Priority: ENV > Project Venv > Global Tool Venv > System 'python3'
 */
function getPythonExecutable() {
    if (process.env.DOCXJS_PYTHON_PATH) return process.env.DOCXJS_PYTHON_PATH;
    
    const projectVenv = path.join(__dirname, '..', 'venv', 'bin', 'python3');
    if (fs.existsSync(projectVenv)) return projectVenv;
    
    const globalEnv = path.join(process.env.HOME || process.env.USERPROFILE || '', '.docxjs-cli-env', 'bin', 'python3');
    if (fs.existsSync(globalEnv)) return globalEnv;
    
    return 'python3'; // Fallback to system path
}

/**
 * Extracts styles from a reference DOCX file.
 * Adapter Logic:
 * - If Electron (Future): Call the bundled executable.
 * - If Standard Node: Call the python source script.
 * 
 * @param {string} docPath - Absolute path to the reference .docx file
 * @returns {Promise<Object>} - The extracted JSON style object
 */
async function extractStyles(docPath) {
    // Future Electron check:
    // const isElectron = process.versions.electron;
    // if (isElectron) { ... call .exe ... }

    // Standard Mode (Source)
    const pythonPath = getPythonExecutable();
    const scriptPath = path.join(__dirname, '..', 'style_extractor.py');

    try {
        const { stdout, stderr } = await execFileAsync(pythonPath, [scriptPath, docPath], { 
            timeout: 20000, // 20s timeout
            encoding: 'utf-8' 
        });
        return JSON.parse(stdout);
    } catch (error) {
        const errorMsg = error.stderr || error.message || 'Unknown error';
        throw new Error(`Failed to extract styles via Python bridge: ${errorMsg}`);
    }
}

module.exports = { extractStyles };
