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
    const isElectron = process.versions.electron;
    let executable;
    let args = [];

    if (isElectron) {
        // Electron Mode: Check for bundled binary (Production)
        const ext = process.platform === 'win32' ? '.exe' : '';
        // process.resourcesPath is available in Electron
        // We assume the binary is placed in 'bin' inside resources
        const bundledPath = path.join(process.resourcesPath, 'bin', 'style_extractor' + ext);
        
        if (fs.existsSync(bundledPath)) {
            executable = bundledPath;
            args = [docPath];
        }
    }

    // Fallback: Standard Mode or Electron Dev Mode (Source)
    if (!executable) {
        executable = getPythonExecutable();
        const scriptPath = path.join(__dirname, '..', 'style_extractor.py');
        args = [scriptPath, docPath];
    }

    try {
        const { stdout, stderr } = await execFileAsync(executable, args, { 
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
