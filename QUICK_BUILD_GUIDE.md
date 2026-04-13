# Quick Build Guide - Cheat Sheet

## 🚀 Build Commands

### Windows
```batch
build_exe.bat
```
**Output:** `dist\PKBM_Generator\PKBM_Generator.exe`

### macOS
```bash
chmod +x build_macos.sh
./build_macos.sh
```
**Output:** `dist/PKBM_Generator.app`

### Linux
```bash
chmod +x build_exe.sh
./build_exe.sh
```
**Output:** `dist/PKBM_Generator/PKBM_Generator`

---

## 📦 Distribution

### Windows
```batch
# ZIP the folder
cd dist
tar -a -c -f PKBM_Generator.zip PKBM_Generator

# Or use Windows Explorer: Right-click > Send to > Compressed folder
```

### macOS
```bash
cd dist
zip -r PKBM_Generator.zip PKBM_Generator.app
```

### Linux
```bash
cd dist
tar -czf PKBM_Generator.tar.gz PKBM_Generator
```

---

## 🔧 Common Fixes

### Build Failed
```bash
# Clean and retry
rm -rf build dist __pycache__
pip install -r requirements.txt
# Run build script again
```

### Python Not Found
```bash
# Windows: Download from python.org
# macOS: brew install python@3.12
# Linux: sudo apt install python3.12
```

### App Won't Open (macOS)
```bash
xattr -cr /path/to/PKBM_Generator.app
```

### Antivirus Blocking (Windows)
```
Add exception for:
- dist\PKBM_Generator\
- build\
```

---

## 📋 Pre-Build Checklist

- [ ] Python 3.8+ installed
- [ ] Run: `pip install -r requirements.txt`
- [ ] Folders exist: `assets/`, `lib/`, `data/`
- [ ] Close other applications (free RAM)

---

## 📖 Detailed Guides

- **Windows:** `BUILD_INSTRUCTIONS.md`
- **macOS:** `BUILD_MACOS.md`
- **All Platforms:** `BUILD_README.md`
- **Dependencies:** `DEPENDENCY_CLEANUP.md`

---

## 🎯 File Sizes (Approximate)

| Platform | Size |
|----------|------|
| Windows | ~200-300 MB |
| macOS | ~250-350 MB |
| Linux | ~200-300 MB |

---

## ⚡ Speed Tips

1. Use SSD (not HDD)
2. Close other apps
3. Use simple spec files (default)
4. Clean before build
5. Disable antivirus temporarily

---

## 🆘 Emergency Commands

### Complete Reset
```bash
# Delete everything
rm -rf build/ dist/ __pycache__/ *.spec.bak
rm -rf venv/  # If using virtual env

# Fresh install
pip install -r requirements.txt

# Build
./build_macos.sh  # or build_exe.bat
```

### Check Installation
```bash
python --version
pip list | grep pyinstaller
pip list | grep reportlab
```

---

## 📞 Quick Support

**Error Message Contains** | **Solution**
---|---
`distutils` | Use simple spec (already default)
`Python not found` | Install Python 3.8+
`Permission denied` | `chmod +x build_*.sh`
`Module not found` | `pip install -r requirements.txt`
`App is damaged` (macOS) | `xattr -cr app.app`

---

## 🎉 Success Indicators

✅ No error messages
✅ `dist/` folder created
✅ App/exe file exists
✅ Can run without errors
✅ All features work

---

**Last Updated:** 2026-04-13
**Version:** 1.2.0
