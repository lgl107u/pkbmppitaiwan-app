# Panduan Build Aplikasi ke EXE

Build aplikasi PKBM Generator menjadi **satu folder** (onedir mode).
User tidak perlu install Python. Semua ter-bundle di dalam folder tersebut.

> Untuk panduan development lengkap, lihat [DEVELOPER_GUIDE.md](DEVELOPER_GUIDE.md)

---

## Prasyarat

1. **Python 3.10+** — https://www.python.org/downloads/ (centang "Add to PATH")
2. **Dependencies** — `pip install -r requirements.txt`

---

## Build Otomatis

### Windows

```cmd
build_exe.bat
```

Script ini akan otomatis menggunakan `build_windows.spec`.

### macOS

```bash
chmod +x build_exe.sh
./build_exe.sh
```

Script ini akan otomatis mendeteksi macOS dan menggunakan `build_macos.spec`.

### Linux

```bash
chmod +x build_exe.sh
./build_exe.sh
```

Script ini akan menggunakan `build_windows.spec` (kompatibel dengan Linux).

### Quick Build (minimal output)

```cmd
quick_build.bat
```

### Hasil Build

**Windows:**
```
dist/PKBM_Generator/          <- Folder ini yang didistribusikan
    PKBM_Generator.exe      <- File utama untuk dijalankan
    base_library.zip
    ... (banyak file dan folder lain)
```

**macOS:**
```
dist/PKBM_Generator.app/      <- macOS App Bundle
    Contents/
        MacOS/
            PKBM_Generator  <- Executable
        Resources/
            ... (assets, data, dll)
```

---

## Build untuk Multi-Platform

**PENTING**: PyInstaller **tidak bisa cross-compile**. Artinya:

- ✅ Build untuk **Windows** harus di komputer **Windows**
- ✅ Build untuk **macOS** harus di komputer **Mac**
- ✅ Build untuk **Linux** harus di komputer **Linux**

### Strategi Distribusi Multi-Platform

1. **Build di Windows** (komputer Anda):
   ```cmd
   build_exe.bat
   ```
   Hasilnya: `dist/PKBM_Generator/` (untuk Windows)

2. **Build di Mac** (komputer rekan atau VM):
   ```bash
   chmod +x build_exe.sh
   ./build_exe.sh
   ```
   Hasilnya: `dist/PKBM_Generator.app/` (untuk macOS)

3. **Distribusi ke Client**:
   - ZIP folder Windows → `PKBM_Generator_Windows.zip`
   - ZIP app macOS → `PKBM_Generator_macOS.zip`
   - Berikan kedua file ZIP ke client sesuai OS mereka

---

## Distribusi ke Client

### File yang Perlu Diberikan ke Client

Setelah build selesai, **ZIP seluruh folder** ini dan berikan ke client:

```
dist/PKBM_Generator/
```

Contoh: `PKBM_Generator_v1.2.zip`

**Folder tersebut sudah mencakup:**
- ✅ Aplikasi lengkap
- ✅ Semua library Python (pandas, reportlab, dll)
- ✅ Template Excel (di folder `data/`)
- ✅ Logo dan assets (di folder `assets/`)
- ✅ Font dan signature processor

### Cara Client Menjalankan

**Windows:**
1. **Ekstrak** file `PKBM_Generator_Windows.zip` di lokasi yang diinginkan (misal: di Desktop).
2. **Masuk ke dalam folder** hasil ekstrak.
3. **Double-click** `PKBM_Generator.exe` untuk menjalankan.

**macOS:**
1. **Ekstrak** file `PKBM_Generator_macOS.zip`.
2. **Drag** `PKBM_Generator.app` ke folder Applications (opsional).
3. **Double-click** `PKBM_Generator.app` untuk menjalankan.
4. Jika muncul peringatan "unidentified developer":
   - Klik kanan → Open → Open
   - Atau: System Preferences → Security & Privacy → "Open Anyway"

**Semua Platform:**
- **Tidak perlu install** Python atau dependency apapun
- **Download template** dari dalam aplikasi (tombol "📥 Template")
  - Template akan otomatis ter-extract dari bundle
  - Client pilih folder tujuan untuk menyimpan template

### Catatan Penting

- **Antivirus**: Beberapa antivirus mungkin false positive. Minta client whitelist file EXE.
- **Start-up Cepat**: Aplikasi akan terbuka jauh lebih cepat (sekitar 1-3 detik).
- **Ukuran Folder**: Total ukuran folder sekitar 80-120 MB.
- **Windows Defender**: Jika diblokir, klik "More info" → "Run anyway".

### Tidak Perlu Diberikan

❌ Folder `build/`  
❌ Folder `__pycache__/`  
❌ File `.py` source code  
❌ File `requirements.txt`  
❌ Folder `assets/` atau `data/` terpisah (sudah di dalam EXE)

## Manual Build

```bash
pip install -r requirements.txt
pyinstaller --clean build_windows.spec
```

## Troubleshooting

| Problem | Solusi |
|---------|--------|
| Python tidak ditemukan | Install Python, centang "Add to PATH" |
| Module not found saat build | `pip install -r requirements.txt` |
| EXE terlalu besar | Build dari virtual environment |
| Antivirus false positive | Whitelist file EXE |
| EXE crash | Ubah `console=True` di spec, rebuild, jalankan dari CMD |

Untuk panduan lengkap: [DEVELOPER_GUIDE.md](DEVELOPER_GUIDE.md)
