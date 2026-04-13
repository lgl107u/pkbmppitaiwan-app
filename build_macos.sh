#!/bin/bash

# ============================================================
# PKBM Generator - macOS Build Script
# Quick and simple build for macOS .app bundle
# ============================================================

set -e  # Exit on error

echo "========================================"
echo "  PKBM Generator - macOS Build"
echo "========================================"
echo

# Check if running on macOS
if [[ "$(uname -s)" != "Darwin" ]]; then
    echo "[ERROR] This script is for macOS only!"
    echo "For Windows, use: build_exe.bat"
    echo "For Linux, use: build_exe.sh"
    exit 1
fi

# Detect Python
if command -v python3 &> /dev/null; then
    PYTHON_CMD=python3
    PIP_CMD=pip3
else
    echo "[ERROR] Python 3 tidak ditemukan!"
    echo "Install Python 3 terlebih dahulu:"
    echo "  brew install python@3.12"
    echo "  atau download dari https://www.python.org/"
    exit 1
fi

echo "[INFO] Python ditemukan"
$PYTHON_CMD --version
echo

# Check Python version (should be 3.8+)
PY_VERSION=$($PYTHON_CMD -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "[INFO] Python version: $PY_VERSION"

if [[ $(echo "$PY_VERSION < 3.8" | bc -l) -eq 1 ]]; then
    echo "[WARNING] Python 3.8+ direkomendasikan, Anda menggunakan $PY_VERSION"
fi
echo

# Install/update dependencies
echo "[INFO] Installing dependencies..."
if [ -f requirements.txt ]; then
    $PIP_CMD install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "[ERROR] Gagal install dependencies!"
        exit 1
    fi
else
    echo "[ERROR] requirements.txt tidak ditemukan!"
    exit 1
fi

echo
echo "[INFO] Verifying PyInstaller..."
$PYTHON_CMD -c "import PyInstaller; print(f'PyInstaller: {PyInstaller.__version__}')" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "[ERROR] PyInstaller tidak terinstall!"
    echo "Mencoba install ulang..."
    $PIP_CMD install --upgrade pyinstaller pyinstaller-hooks-contrib
fi

# Clean previous build
echo
echo "[INFO] Cleaning previous build..."
rm -rf build dist *.spec.bak
echo "[OK] Build folder cleaned"

# Verify required files
echo
echo "[INFO] Verifying files..."
if [ ! -f pkbm_generator_app.py ]; then
    echo "[ERROR] pkbm_generator_app.py tidak ditemukan!"
    exit 1
fi

# Select spec file
if [ -f "build_macos_simple.spec" ]; then
    SPEC_FILE="build_macos_simple.spec"
    echo "[INFO] Using: build_macos_simple.spec (recommended)"
elif [ -f "build_macos.spec" ]; then
    SPEC_FILE="build_macos.spec"
    echo "[INFO] Using: build_macos.spec"
else
    echo "[ERROR] No macOS spec file found!"
    echo "Expected: build_macos_simple.spec or build_macos.spec"
    exit 1
fi

# Check assets
if [ ! -d assets ]; then
    echo "[WARNING] assets/ folder tidak ditemukan"
fi
if [ ! -d lib ]; then
    echo "[WARNING] lib/ folder tidak ditemukan"
fi

# Start build
echo
echo "========================================"
echo "  Building macOS App..."
echo "========================================"
echo
echo "[INFO] This may take 2-5 minutes..."
echo "[INFO] Please wait..."
echo

pyinstaller --clean "$SPEC_FILE"

# Check result
if [ $? -ne 0 ]; then
    echo
    echo "========================================"
    echo "  BUILD FAILED!"
    echo "========================================"
    echo
    echo "[ERROR] Build error occurred"
    echo
    echo "Troubleshooting tips:"
    echo "  1. Check error message above"
    echo "  2. Try: pip3 install -r requirements.txt"
    echo "  3. Try: rm -rf build dist && ./build_macos.sh"
    echo "  4. Check BUILD_MACOS.md for detailed guide"
    echo
    exit 1
fi

# Success!
echo
echo "========================================"
echo "  BUILD SUCCESS!"
echo "========================================"
echo

if [ -d "dist/PKBM_Generator.app" ]; then
    echo "[OK] macOS App created successfully!"
    echo
    echo "Location: dist/PKBM_Generator.app"
    echo
    echo "To test:"
    echo "  open dist/PKBM_Generator.app"
    echo
    echo "To distribute:"
    echo "  1. ZIP the app:"
    echo "     cd dist && zip -r PKBM_Generator.zip PKBM_Generator.app"
    echo
    echo "  2. Send ZIP to users"
    echo
    echo "  3. Users: Extract and drag to Applications folder"
    echo
    echo "App size:"
    du -sh dist/PKBM_Generator.app
    echo
    echo "IMPORTANT: If users get 'damaged app' error:"
    echo "  Tell them to run: xattr -cr /path/to/PKBM_Generator.app"
    echo "  Or: Right-click > Open (first time only)"
    echo
    echo "For detailed guide, see: BUILD_MACOS.md"
    echo
else
    echo "[ERROR] App not found at expected location!"
    echo "Check dist/ folder for build output"
    echo
    exit 1
fi
