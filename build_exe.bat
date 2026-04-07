@echo off
chcp 65001 >nul
echo ========================================
echo   PKBM Generator - Build to EXE
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan!
    echo Silakan install Python terlebih dahulu dari https://www.python.org/
    echo.
    pause
    exit /b 1
)

echo [INFO] Python ditemukan
python --version
echo.

REM Check and install requirements
echo [INFO] Memeriksa dependencies...
if exist requirements.txt (
    echo [INFO] Installing/updating dependencies dari requirements.txt...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Gagal install dependencies!
        echo.
        pause
        exit /b 1
    )
) else (
    echo [WARNING] requirements.txt tidak ditemukan
    echo [INFO] Installing PyInstaller saja...
    pip install pyinstaller pyinstaller-hooks-contrib
)

echo.
echo [INFO] Verifikasi PyInstaller...
python -c "import PyInstaller; print(f'PyInstaller version: {PyInstaller.__version__}')" 2>nul
if errorlevel 1 (
    echo [ERROR] PyInstaller tidak terinstall dengan benar!
    echo [INFO] Mencoba install ulang...
    pip install --upgrade pyinstaller pyinstaller-hooks-contrib
    if errorlevel 1 (
        echo [ERROR] Gagal install PyInstaller!
        pause
        exit /b 1
    )
)

echo.
echo [INFO] Membersihkan build sebelumnya...
if exist build (
    rmdir /s /q build
    echo [OK] Folder build dihapus
)
if exist dist (
    rmdir /s /q dist
    echo [OK] Folder dist dihapus
)

echo.
echo [INFO] Memverifikasi file yang diperlukan...
if not exist pkbm_generator_app.py (
    echo [ERROR] File pkbm_generator_app.py tidak ditemukan!
    pause
    exit /b 1
)
if not exist assets (
    echo [WARNING] Folder assets tidak ditemukan
)
if not exist lib (
    echo [WARNING] Folder lib tidak ditemukan
)

REM Pilih spec file yang akan digunakan
set SPEC_FILE=build_simple.spec
if not exist build_simple.spec (
    echo [INFO] build_simple.spec tidak ada, menggunakan build_windows.spec
    set SPEC_FILE=build_windows.spec
)
if not exist %SPEC_FILE% (
    echo [ERROR] File spec tidak ditemukan!
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Memulai Build Executable...
echo ========================================
echo.
echo [INFO] Menggunakan: %SPEC_FILE%
echo [INFO] Ini mungkin memakan waktu 2-5 menit...
echo [INFO] Mohon tunggu...
echo.

pyinstaller --clean %SPEC_FILE%

if errorlevel 1 (
    echo.
    echo ========================================
    echo   BUILD GAGAL!
    echo ========================================
    echo.
    echo [ERROR] Terjadi error saat build executable
    echo [INFO] Periksa error message di atas
    echo [INFO] Tips troubleshooting:
    echo   1. Pastikan semua dependencies terinstall
    echo   2. Coba jalankan: pip install -r requirements.txt
    echo   3. Coba hapus folder build dan dist secara manual
    echo   4. Pastikan tidak ada antivirus yang memblokir
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   BUILD BERHASIL!
echo ========================================
echo.

if exist "dist\PKBM_Generator.exe" (
    echo [OK] Executable berhasil dibuat!
    echo.
    echo Lokasi file: dist\PKBM_Generator.exe
    echo.
    echo Cara distribusi:
    echo 1. Copy file 'dist\PKBM_Generator.exe' ke komputer target
    echo 2. User tinggal double-click PKBM_Generator.exe
    echo 3. Tidak perlu install Python atau library apapun!
    echo 4. Semua data sudah ter-bundle di dalam satu file EXE
    echo.
    echo Ukuran file:
    for %%A in ("dist\PKBM_Generator.exe") do echo   %%~zA bytes
    echo.
) else (
    echo [ERROR] Executable tidak ditemukan di lokasi yang diharapkan!
    echo [INFO] Periksa folder dist untuk melihat hasil build
    echo.
)

pause
