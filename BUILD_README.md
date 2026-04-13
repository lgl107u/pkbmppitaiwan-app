# Build Scripts Overview

Panduan lengkap untuk build PKBM Generator di berbagai platform.

## Quick Start

### Windows
```batch
build_exe.bat
```

### macOS
```bash
chmod +x build_macos.sh
./build_macos.sh
```

### Linux
```bash
chmod +x build_exe.sh
./build_exe.sh
```

## File Structure

```
pkbm-rapot-main/
├── build_exe.bat              # Windows build script
├── build_exe.sh               # Universal Unix build script (macOS/Linux)
├── build_macos.sh             # macOS-specific build script (recommended)
│
├── build_simple.spec          # Minimal spec for Windows/Linux (recommended)
├── build_windows.spec         # Full spec for Windows
├── build_macos_simple.spec    # Minimal spec for macOS (recommended)
├── build_macos.spec           # Full spec for macOS
│
├── BUILD_MACOS.md             # Detailed macOS guide
├── DEPENDENCY_CLEANUP.md      # Dependency cleanup explanation
└── BUILD_README.md            # This file
```

## Spec Files Explained

### Simple Specs (Recommended)
- `build_simple.spec` - Windows/Linux
- `build_macos_simple.spec` - macOS

**Advantages:**
- ✅ Faster build time
- ✅ Smaller file size
- ✅ Avoids distutils conflict
- ✅ More stable with Python 3.12
- ✅ Excludes unnecessary packages

**Use when:**
- First time building
- Python 3.12+
- Want smaller executable
- Encountering build errors

### Full Specs (Alternative)
- `build_windows.spec` - Windows
- `build_macos.spec` - macOS

**Advantages:**
- ✅ More comprehensive imports
- ✅ Fallback if simple fails
- ✅ Includes more submodules

**Use when:**
- Simple spec fails
- Need specific submodules
- Debugging import issues

## Build Scripts Comparison

### build_exe.bat (Windows)
```batch
# Features:
- Auto-detects Python
- Installs dependencies
- Cleans previous builds
- Uses build_simple.spec by default
- Fallback to build_windows.spec
- Shows file size after build
```

### build_macos.sh (macOS Only)
```bash
# Features:
- macOS-specific optimizations
- Creates .app bundle
- Checks macOS version
- Uses build_macos_simple.spec by default
- Provides distribution instructions
- Shows app size
```

### build_exe.sh (Universal Unix)
```bash
# Features:
- Works on macOS and Linux
- Auto-detects OS
- Selects appropriate spec file
- Generic Unix compatibility
```

**Recommendation:**
- **macOS**: Use `build_macos.sh` (more specific)
- **Linux**: Use `build_exe.sh`
- **Windows**: Use `build_exe.bat`

## Dependencies

All platforms use the same `requirements.txt`:

```
reportlab==4.2.5       # PDF generation
pandas==2.2.3          # Excel processing
numpy==1.26.4          # Array operations
pillow==11.0.0         # Image processing
openpyxl==3.1.5        # Excel support
opencv-python==4.10.0.84  # Advanced image processing
pyinstaller==6.11.0    # Build tool
pyinstaller-hooks-contrib==2024.10  # PyInstaller hooks
```

**Removed (unused):**
- ❌ python-docx
- ❌ qrcode

## Platform-Specific Output

### Windows
```
dist/
└── PKBM_Generator/
    ├── PKBM_Generator.exe    # Main executable
    ├── _internal/            # Dependencies
    └── ...
```

### macOS
```
dist/
└── PKBM_Generator.app/       # macOS App Bundle
    └── Contents/
        ├── MacOS/
        ├── Resources/
        └── Info.plist
```

### Linux
```
dist/
└── PKBM_Generator/
    ├── PKBM_Generator        # Executable
    ├── _internal/            # Dependencies
    └── ...
```

## Common Issues & Solutions

### Issue 1: distutils conflict
```
ValueError: Target module "distutils" already imported
```

**Solution:**
```bash
# Use simple spec files (already set as default)
# Or manually:
pip uninstall pyinstaller pyinstaller-hooks-contrib -y
pip install pyinstaller==6.11.0 pyinstaller-hooks-contrib==2024.10
```

### Issue 2: Python not found
**Windows:**
```batch
# Install from python.org
# Make sure to check "Add to PATH"
```

**macOS:**
```bash
brew install python@3.12
```

**Linux:**
```bash
sudo apt install python3.12 python3-pip  # Debian/Ubuntu
sudo yum install python3.12 python3-pip  # RHEL/CentOS
```

### Issue 3: Build too slow
```bash
# Clean cache
rm -rf build/ dist/ __pycache__/
rm -rf ~/.pyinstaller/  # macOS/Linux
rmdir /s /q %LOCALAPPDATA%\pyinstaller  # Windows

# Build again
```

### Issue 4: App won't open on other computers
**Windows:**
- Antivirus might block it
- User needs to "Run anyway"

**macOS:**
```bash
# Remove quarantine
xattr -cr /path/to/PKBM_Generator.app

# Or right-click > Open (first time)
```

**Linux:**
```bash
# Make executable
chmod +x PKBM_Generator
```

## Build Checklist

Before building:
- [ ] Python 3.8+ installed
- [ ] All dependencies installed (`pip install -r requirements.txt`)
- [ ] `assets/` folder exists with logo and fonts
- [ ] `lib/` folder exists with generator modules
- [ ] `data/` folder exists with templates

After building:
- [ ] Executable runs without errors
- [ ] All features work (generate PDF, load Excel)
- [ ] Assets loaded correctly (logo, fonts)
- [ ] File size reasonable (< 500MB)
- [ ] Tested on clean machine (if possible)

## Distribution

### Windows
1. ZIP the entire `dist/PKBM_Generator/` folder
2. Send to users
3. Users extract and run `PKBM_Generator.exe`

### macOS
1. ZIP `dist/PKBM_Generator.app`
2. Send to users
3. Users extract and drag to Applications
4. First time: Right-click > Open

### Linux
1. ZIP the entire `dist/PKBM_Generator/` folder
2. Send to users
3. Users extract and run `./PKBM_Generator`

## Advanced Options

### Universal Binary (macOS)
```bash
# Build for both Intel and Apple Silicon
pyinstaller --target-arch universal2 build_macos_simple.spec
```

### One-File Mode (Windows)
```python
# Edit spec file, change:
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,  # Add this
    a.zipfiles,  # Add this
    a.datas,     # Add this
    [],
    name='PKBM_Generator',
    # ... rest of options
)

# Remove COLLECT section
```

### Custom Icon
```python
# In spec file:
icon='path/to/icon.ico'  # Windows
icon='path/to/icon.icns'  # macOS
```

## Performance Tips

1. **Use simple specs** - Faster build, smaller size
2. **Clean before build** - Remove old artifacts
3. **Use SSD** - Much faster than HDD
4. **Close other apps** - More RAM for build
5. **Disable antivirus** - During build only

## Getting Help

1. Check error message carefully
2. Read platform-specific guide:
   - Windows: `BUILD_INSTRUCTIONS.md`
   - macOS: `BUILD_MACOS.md`
   - Dependencies: `DEPENDENCY_CLEANUP.md`
3. Try alternative spec file
4. Clean and rebuild
5. Check PyInstaller docs: https://pyinstaller.org/

## Version Info

- **App Version**: 1.2.0
- **Python**: 3.8 - 3.12
- **PyInstaller**: 6.11.0
- **Last Updated**: 2026-04-13
