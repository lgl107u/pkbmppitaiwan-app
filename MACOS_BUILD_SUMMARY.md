# macOS Build - Summary

## ✅ Yang Sudah Dibuat

### 1. Build Scripts
- ✅ `build_macos.sh` - Script khusus macOS (RECOMMENDED)
- ✅ `build_exe.sh` - Updated untuk support macOS dengan auto-detect

### 2. Spec Files
- ✅ `build_macos_simple.spec` - Minimal spec untuk macOS (RECOMMENDED)
- ✅ `build_macos.spec` - Updated dengan dependency cleanup

### 3. Documentation
- ✅ `BUILD_MACOS.md` - Panduan lengkap build di macOS
- ✅ `BUILD_README.md` - Overview semua build scripts
- ✅ `MACOS_BUILD_SUMMARY.md` - File ini

## 🚀 Cara Build di macOS

### Quick Start (3 Langkah)

```bash
# 1. Masuk ke folder project
cd /path/to/pkbm-rapot-main

# 2. Berikan permission
chmod +x build_macos.sh

# 3. Build!
./build_macos.sh
```

### Output
```
dist/PKBM_Generator.app    # macOS Application Bundle
```

## 📦 Cara Distribusi

### Untuk User Mac

```bash
# 1. ZIP aplikasi
cd dist
zip -r PKBM_Generator.zip PKBM_Generator.app

# 2. Kirim PKBM_Generator.zip ke user

# 3. User ekstrak dan drag ke Applications folder
```

### Instruksi untuk User

**Jika muncul "App is damaged" error:**

**Cara 1 - Via Terminal:**
```bash
xattr -cr /Applications/PKBM_Generator.app
```

**Cara 2 - Via GUI:**
1. Right-click `PKBM_Generator.app`
2. Pilih "Open"
3. Klik "Open" di dialog
4. Setelah pertama kali, bisa double-click biasa

## 🔧 Troubleshooting

### Python tidak ditemukan
```bash
# Install via Homebrew
brew install python@3.12

# Atau download dari python.org
```

### Build gagal dengan distutils error
```bash
# Sudah diatasi di build_macos_simple.spec
# Jika masih error, reinstall PyInstaller:
pip3 uninstall pyinstaller pyinstaller-hooks-contrib -y
pip3 install -r requirements.txt
```

### App tidak bisa dibuka di Mac lain
```bash
# Hapus quarantine attribute
xattr -cr dist/PKBM_Generator.app

# Atau instruksikan user untuk right-click > Open
```

## 📋 Checklist Build

**Sebelum Build:**
- [ ] Python 3.8+ terinstall
- [ ] Dependencies terinstall (`pip3 install -r requirements.txt`)
- [ ] Folder `assets/`, `lib/`, `data/` ada

**Setelah Build:**
- [ ] App bisa dibuka tanpa error
- [ ] Semua fitur berfungsi (generate PDF, load Excel)
- [ ] Assets ter-load (logo, fonts)
- [ ] File size < 500MB

**Sebelum Distribusi:**
- [ ] Test di Mac lain (jika memungkinkan)
- [ ] ZIP file sudah dibuat
- [ ] Instruksi user sudah disiapkan

## 📁 File Structure

```
pkbm-rapot-main/
├── build_macos.sh              ⭐ Script build macOS (GUNAKAN INI)
├── build_macos_simple.spec     ⭐ Spec file (auto-selected)
├── build_macos.spec            (alternative)
├── BUILD_MACOS.md              📖 Panduan lengkap
└── requirements.txt            📦 Dependencies
```

## 🎯 Perbedaan dengan Windows

| Aspek | Windows | macOS |
|-------|---------|-------|
| Script | `build_exe.bat` | `build_macos.sh` |
| Output | `.exe` file | `.app` bundle |
| Distribusi | ZIP folder | ZIP .app |
| Icon | `.ico` | `.icns` |
| Gatekeeper | Tidak ada | Ada (perlu xattr) |

## 💡 Tips

1. **Gunakan `build_macos.sh`** - Lebih spesifik untuk macOS
2. **Simple spec** - Sudah jadi default, lebih cepat dan kecil
3. **Test di Mac lain** - Jika memungkinkan
4. **Buat DMG** - Lebih professional (lihat BUILD_MACOS.md)
5. **Code signing** - Jika punya Apple Developer account

## 📞 Support

Jika ada masalah:
1. Baca error message di terminal
2. Cek `BUILD_MACOS.md` untuk detail
3. Cek `BUILD_README.md` untuk overview
4. Coba clean build: `rm -rf build dist && ./build_macos.sh`

## ✨ Fitur Baru

Dibanding script lama:
- ✅ Auto-detect macOS
- ✅ Pilih spec file otomatis (simple > regular)
- ✅ Dependency cleanup (lebih kecil, lebih cepat)
- ✅ Better error handling
- ✅ Instruksi distribusi yang jelas
- ✅ Troubleshooting guide lengkap

## 🎉 Ready to Build!

Sekarang script macOS sudah lengkap dan siap digunakan. Tinggal jalankan:

```bash
./build_macos.sh
```

Dan aplikasi `.app` akan dibuat di folder `dist/`!
