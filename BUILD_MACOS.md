# Build Guide untuk macOS

Panduan lengkap untuk membuat aplikasi PKBM Generator di macOS (.app bundle).

## Persyaratan

### 1. Python 3.12
```bash
# Cek versi Python
python3 --version

# Jika belum ada, install via Homebrew
brew install python@3.12
```

### 2. Xcode Command Line Tools (opsional tapi direkomendasikan)
```bash
xcode-select --install
```

## Langkah-langkah Build

### 1. Persiapan Environment

```bash
# Masuk ke folder project
cd /path/to/pkbm-rapot-main

# Buat virtual environment (opsional tapi direkomendasikan)
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip3 install -r requirements.txt
```

### 2. Build Aplikasi

```bash
# Berikan permission execute ke script
chmod +x build_exe.sh

# Jalankan build script
./build_exe.sh
```

Script akan otomatis:
- Deteksi macOS
- Pilih `build_macos_simple.spec` (recommended) atau `build_macos.spec`
- Build aplikasi menjadi `.app` bundle
- Simpan hasil di folder `dist/`

### 3. Hasil Build

Setelah build berhasil, Anda akan mendapatkan:
```
dist/
└── PKBM_Generator.app/    # macOS Application Bundle
```

## Cara Menjalankan

### Di Mac Anda (Development)
```bash
# Double-click PKBM_Generator.app di Finder
# atau via terminal:
open dist/PKBM_Generator.app
```

### Distribusi ke User Lain

#### Metode 1: ZIP dan Kirim
```bash
# Compress .app bundle
cd dist
zip -r PKBM_Generator.zip PKBM_Generator.app

# Kirim PKBM_Generator.zip ke user
# User ekstrak dan drag ke Applications folder
```

#### Metode 2: DMG (Lebih Professional)
```bash
# Install create-dmg
brew install create-dmg

# Buat DMG installer
create-dmg \
  --volname "PKBM Generator" \
  --window-pos 200 120 \
  --window-size 800 400 \
  --icon-size 100 \
  --icon "PKBM_Generator.app" 200 190 \
  --hide-extension "PKBM_Generator.app" \
  --app-drop-link 600 185 \
  "PKBM_Generator.dmg" \
  "dist/"
```

## Troubleshooting

### Error: "PKBM_Generator.app is damaged"

Ini terjadi karena macOS Gatekeeper. Solusi:

```bash
# Hapus quarantine attribute
xattr -cr dist/PKBM_Generator.app

# Atau izinkan di System Preferences
# System Preferences > Security & Privacy > General
# Klik "Open Anyway"
```

### Error: "Python not found"

```bash
# Install Python via Homebrew
brew install python@3.12

# Atau download dari python.org
# https://www.python.org/downloads/macos/
```

### Error: "PyInstaller not found"

```bash
# Install PyInstaller
pip3 install pyinstaller pyinstaller-hooks-contrib

# Atau reinstall dependencies
pip3 install -r requirements.txt
```

### Error: "distutils conflict"

Ini sudah diatasi di `build_macos_simple.spec`. Jika masih terjadi:

```bash
# Uninstall dan reinstall PyInstaller
pip3 uninstall pyinstaller pyinstaller-hooks-contrib -y
pip3 install pyinstaller==6.11.0 pyinstaller-hooks-contrib==2024.10

# Build ulang
./build_exe.sh
```

### Build Lambat atau Hang

```bash
# Bersihkan cache PyInstaller
rm -rf build/ dist/ __pycache__/
rm -rf ~/.pyinstaller/

# Build ulang
./build_exe.sh
```

### App Tidak Bisa Dibuka di Mac Lain

Jika app tidak bisa dibuka di Mac user lain, ada 2 solusi:

#### Solusi 1: Code Signing (Butuh Apple Developer Account)
```bash
# Sign the app
codesign --deep --force --sign "Developer ID Application: Your Name" dist/PKBM_Generator.app

# Verify
codesign --verify --verbose dist/PKBM_Generator.app
```

#### Solusi 2: Instruksikan User (Gratis)
Kirim instruksi ini ke user:

1. Download dan ekstrak PKBM_Generator.app
2. Jangan double-click dulu!
3. Buka Terminal
4. Jalankan: `xattr -cr /path/to/PKBM_Generator.app`
5. Sekarang double-click PKBM_Generator.app

Atau via GUI:
1. Right-click PKBM_Generator.app
2. Pilih "Open"
3. Klik "Open" di dialog yang muncul
4. Setelah pertama kali, bisa double-click seperti biasa

## Spec Files

### build_macos_simple.spec (Recommended)
- Minimal dependencies
- Excludes problematic modules
- Faster build
- Smaller file size
- Lebih stabil untuk Python 3.12

### build_macos.spec (Alternative)
- Full dependencies
- Jika simple spec gagal
- Lebih besar tapi lebih lengkap

## Struktur .app Bundle

```
PKBM_Generator.app/
├── Contents/
│   ├── Info.plist           # App metadata
│   ├── MacOS/
│   │   └── PKBM_Generator   # Executable
│   └── Resources/
│       ├── assets/          # Logo, fonts
│       ├── lib/             # Generator modules
│       └── data/            # Template files
```

## Tips Optimasi

### 1. Reduce File Size
```bash
# Gunakan UPX compression (sudah enabled di spec)
# Hapus debug symbols
strip dist/PKBM_Generator.app/Contents/MacOS/PKBM_Generator
```

### 2. Test di Mac Lain
Sebelum distribusi, test di Mac yang berbeda:
- macOS version berbeda (10.13+)
- Intel vs Apple Silicon (M1/M2/M3)
- Fresh install (tanpa Python/dependencies)

### 3. Universal Binary (Intel + Apple Silicon)
```bash
# Build untuk both architectures
pyinstaller --target-arch universal2 build_macos_simple.spec
```

## Checklist Sebelum Distribusi

- [ ] App bisa dibuka tanpa error
- [ ] Semua fitur berfungsi (generate PDF, load Excel, dll)
- [ ] Assets ter-bundle dengan benar (logo, fonts)
- [ ] Test di Mac lain (jika memungkinkan)
- [ ] File size reasonable (< 500MB)
- [ ] Dokumentasi user sudah lengkap

## Support

Jika ada masalah:
1. Cek error message di terminal
2. Lihat log file di `~/.pyinstaller/`
3. Coba build ulang dengan `--clean` flag
4. Gunakan spec file alternative (simple vs regular)

## Versi

- Python: 3.12.x
- PyInstaller: 6.11.0
- macOS: 10.13+ (High Sierra atau lebih baru)
- Architecture: Intel x86_64 dan Apple Silicon (M1/M2/M3)
