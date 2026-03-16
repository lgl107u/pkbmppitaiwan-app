#!/bin/bash

set -e  # Exit on error

echo "========================================"
echo "  PKBM Generator - Build to EXE"
echo "========================================"
echo

# Detect Python command
if command -v python3 &> /dev/null; then
    PYTHON_CMD=python3
    PIP_CMD=pip3
elif command -v python &> /dev/null; then
    PYTHON_CMD=python
    PIP_CMD=pip
else
    echo "[ERROR] Python tidak ditemukan!"
    echo "Silakan install Python terlebih dahulu dari https://www.python.org/"
    exit 1
fi

echo "[INFO] Python ditemukan"
$PYTHON_CMD --version
echo

# Check and install requirements
echo "[INFO] Memeriksa dependencies..."
if [ -f requirements.txt ]; then
    echo "[INFO] Installing/updating dependencies dari requirements.txt..."
    $PIP_CMD install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "[ERROR] Gagal install dependencies!"
        exit 1
    fi
else
    echo "[WARNING] requirements.txt tidak ditemukan"
    echo "[INFO] Installing PyInstaller saja..."
    $PIP_CMD install pyinstaller pyinstaller-hooks-contrib
fi

echo
echo "[INFO] Verifikasi PyInstaller..."
$PYTHON_CMD -c "import PyInstaller; print(f'PyInstaller version: {PyInstaller.__version__}')" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "[ERROR] PyInstaller tidak terinstall dengan benar!"
    echo "[INFO] Mencoba install ulang..."
    $PIP_CMD install --upgrade pyinstaller pyinstaller-hooks-contrib
    if [ $? -ne 0 ]; then
        echo "[ERROR] Gagal install PyInstaller!"
        exit 1
    fi
fi

echo
echo "[INFO] Membersihkan build sebelumnya..."
if [ -d build ]; then
    rm -rf build
    echo "[OK] Folder build dihapus"
fi
if [ -d dist ]; then
    rm -rf dist
    echo "[OK] Folder dist dihapus"
fi

echo
echo "[INFO] Memverifikasi file yang diperlukan..."
if [ ! -f pkbm_generator_app.py ]; then
    echo "[ERROR] File pkbm_generator_app.py tidak ditemukan!"
    exit 1
fi

# Detect OS and select appropriate spec file
OS_TYPE=$(uname -s)
if [[ "$OS_TYPE" == "Darwin" ]]; then
    SPEC_FILE="build_macos.spec"
    echo "[INFO] Detected macOS - using build_macos.spec"
else
    SPEC_FILE="build_windows.spec"
    echo "[INFO] Detected Linux/Other - using build_windows.spec"
fi

if [ ! -f "$SPEC_FILE" ]; then
    echo "[ERROR] File $SPEC_FILE tidak ditemukan!"
    exit 1
fi

if [ ! -d assets ]; then
    echo "[WARNING] Folder assets tidak ditemukan"
fi
if [ ! -d lib ]; then
    echo "[WARNING] Folder lib tidak ditemukan"
fi

echo
echo "========================================"
echo "  Memulai Build Executable..."
echo "========================================"
echo
echo "[INFO] Ini mungkin memakan waktu 2-5 menit..."
echo "[INFO] Mohon tunggu..."
echo

pyinstaller --clean "$SPEC_FILE"

if [ $? -ne 0 ]; then
    echo
    echo "========================================"
    echo "  BUILD GAGAL!"
    echo "========================================"
    echo
    echo "[ERROR] Terjadi error saat build executable"
    echo "[INFO] Periksa error message di atas"
    echo "[INFO] Tips troubleshooting:"
    echo "  1. Pastikan semua dependencies terinstall"
    echo "  2. Coba jalankan: pip install -r requirements.txt"
    echo "  3. Coba hapus folder build dan dist secara manual"
    echo "  4. Pastikan tidak ada antivirus yang memblokir"
    echo
    exit 1
fi

echo
echo "========================================"
echo "  BUILD BERHASIL!"
echo "========================================"
echo

if [ -d "dist/PKBM_Generator" ]; then
    echo "[OK] Aplikasi berhasil dibuat!"
    echo
    if [[ "$OS_TYPE" == "Darwin" ]]; then
        echo "Lokasi: dist/PKBM_Generator.app (macOS App Bundle)"
        echo
        echo "Cara distribusi:"
        echo "1. ZIP folder 'dist/PKBM_Generator.app'"
        echo "2. Kirim ZIP ke user Mac"
        echo "3. User ekstrak dan drag PKBM_Generator.app ke Applications"
        echo "4. Double-click untuk menjalankan"
    else
        echo "Lokasi: dist/PKBM_Generator/ (folder)"
        echo
        echo "Cara distribusi:"
        echo "1. ZIP seluruh folder 'dist/PKBM_Generator/'"
        echo "2. Kirim ZIP ke user"
        echo "3. User ekstrak dan jalankan PKBM_Generator executable"
    fi
    echo
    echo "Ukuran total:"
    du -sh dist/PKBM_Generator* 2>/dev/null || echo "  (tidak dapat menghitung ukuran)"
    echo
else
    echo "[ERROR] Output tidak ditemukan di lokasi yang diharapkan!"
    echo "[INFO] Periksa folder dist untuk melihat hasil build"
    echo
    exit 1
fi
