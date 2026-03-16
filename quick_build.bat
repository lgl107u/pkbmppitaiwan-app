@echo off
REM Quick Build Script - Untuk build cepat tanpa banyak output
echo Building PKBM Generator (onefile)...
pip install -q -r requirements.txt
pyinstaller --clean --log-level ERROR build_windows.spec
if exist "dist\PKBM_Generator.exe" (
    echo.
    echo [SUCCESS] Build complete!
    echo Location: dist\PKBM_Generator.exe
) else (
    echo.
    echo [ERROR] Build failed! Run build_exe.bat for detailed output.
)
pause
