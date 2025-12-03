#!/usr/bin/env node

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

// Resolve the root of the docxjs-cli package when installed globally
// This script would be in node_modules/docxjs-cli/bin/docxjs-web.js
// So, the project root is path.join(__dirname, '..', '..')
const packageRoot = path.resolve(__dirname, '..', '..');
const serverAppPath = path.join(packageRoot, 'server', 'app.js');

if (!fs.existsSync(serverAppPath)) {
    console.error(`Error: Could not find server/app.js at ${serverAppPath}`);
    console.error("Please ensure docxjs-cli is installed correctly globally.");
    process.exit(1);
}

console.log(`Starting docxjs-cli web server from: ${packageRoot}`);
console.log(`Access at http://localhost:3000`);

// Use npm start to ensure all npm scripts and environment variables are honored
const child = spawn('npm', ['start'], {
    cwd: packageRoot, // Crucial: run npm start from the package's root directory
    stdio: 'inherit' // Inherit stdio so user sees output
});

child.on('error', (err) => {
    console.error('Failed to start web server:', err);
});

child.on('close', (code) => {
    if (code !== 0) {
        console.error(`Web server exited with code ${code}`);
    } else {
        console.log('Web server stopped.');
    }
});

// Handle SIGINT (Ctrl+C) to gracefully shut down the child process
process.on('SIGINT', () => {
    console.log('\nShutting down web server...');
    child.kill('SIGINT');
});
