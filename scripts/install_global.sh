#!/bin/bash

set -e # Exit on error

echo "🚀 Installing docxjs-cli globally..."

# --- 1. Check Prerequisites ---
if ! command -v npm &> /dev/null; then
    echo "❌ Error: 'npm' is not installed. Please install Node.js first."
    exit 1
fi

if ! command -v python3 &> /dev/null; then
    echo "❌ Error: 'python3' is not installed."
    exit 1
fi

# --- 2. Clone Repository (if not running locally) ---
# If .git directory doesn't exist, assume we are running via curl/wget and need to clone
if [ ! -d ".git" ]; then
    echo "📥 Cloning docxjs-cli from GitHub..."
    TEMP_DIR=$(mktemp -d)
    git clone https://github.com/ttieli/docxjs-cli.git "$TEMP_DIR"
    cd "$TEMP_DIR"
else
    echo "📂 Installing from current directory..."
fi

# --- 3. Setup Global Python Environment ---
ENV_DIR="$HOME/.docxjs-cli-env"

echo "🐍 Setting up isolated Python environment at $ENV_DIR..."

if [ -d "$ENV_DIR" ]; then
    echo "   Environment already exists. Updating..."
else
    python3 -m venv "$ENV_DIR"
    echo "   Created venv."
fi

# --- 4. Install Python Dependencies ---
echo "📥 Installing python-docx..."
"$ENV_DIR/bin/pip" install --upgrade pip
"$ENV_DIR/bin/pip" install python-docx

# --- 5. Install NPM Package Globally ---
echo "📦 Installing NPM package..."
npm install
npm install -g .

# --- 6. Cleanup ---
if [[ "$PWD" == "$TEMP_DIR" ]]; then
    echo "🧹 Cleaning up temporary files..."
    rm -rf "$TEMP_DIR"
fi

echo ""
echo "✅ Installation Complete!"
echo "🎉 You can now run 'docxjs' from anywhere in your terminal."
echo "   Try: docxjs --help"