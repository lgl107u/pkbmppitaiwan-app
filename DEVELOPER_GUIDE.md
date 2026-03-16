# PKBM Generator - Developer Guide

Panduan lengkap untuk development, debugging, dan build aplikasi PKBM Generator.

---

## Daftar Isi

1. [Struktur Project](#struktur-project)
2. [Setup Development Environment](#setup-development-environment)
3. [Menjalankan Aplikasi (Development)](#menjalankan-aplikasi-development)
4. [Debugging](#debugging)
5. [Build ke EXE](#build-ke-exe)
6. [Format Excel per Generator](#format-excel-per-generator)
7. [Arsitektur Path (Dev vs EXE)](#arsitektur-path)
8. [Troubleshooting](#troubleshooting)

---

## Struktur Project

```
pkbm-rapot-main/
├── pkbm_generator_app.py      # Main GUI application (entry point)
├── requirements.txt            # Python dependencies
├── build_windows.spec          # PyInstaller spec (onefile mode)
├── build_exe.bat               # Build script Windows
├── build_exe.sh                # Build script Linux/Mac
├── quick_build.bat             # Quick build (minimal output)
│
├── lib/
│   └── pdf_generators/
│       ├── __init__.py
│       ├── rapot_generator.py      # Rapot PDF generator
│       ├── skhupk_generator.py     # SKHUPK PDF generator
│       └── transcript_generator.py # Transkrip Nilai PDF generator
│
├── assets/
│   ├── logo.png                # Logo sekolah
│   ├── fonts/                  # Font Times New Roman (ttf)
│   │   ├── times.ttf
│   │   ├── timesbd.ttf
│   │   ├── timesbdit.ttf
│   │   └── timesit.ttf
│   └── signatures/             # Default signature images
│       ├── ttd_kepsek.png
│       ├── ttd_wali.png
│       └── dummy.png
│
├── data/                       # Template Excel files (bundled dalam EXE)
│   ├── rapot/
│   │   ├── template_data_siswa.xlsx
│   │   └── template_nilai_siswa.xlsx
│   ├── skhupk/
│   │   └── template_skhupk.xlsx
│   ├── kartu_upk/
│   │   └── template_kartu_upk.xlsx
│   └── transkrip/
│       └── Database Transkrip Nilai.xlsx
│
├── output/                     # Folder output PDF (auto-created)
├── BUILD_INSTRUCTIONS.md       # Panduan build EXE
├── INSTALL_GUIDE.md            # Panduan install (Linux)
├── README_USER.txt             # Panduan end-user
└── DEVELOPER_GUIDE.md          # Dokumen ini
```

---

## Setup Development Environment

### 1. Install Python

- Download Python 3.10+ dari https://www.python.org/downloads/
- **PENTING:** Centang "Add Python to PATH" saat install

### 2. Buat Virtual Environment

```bash
# Windows
cd "d:\User Boby\Downloads\pkbm-rapot-main"
python -m venv venv
venv\Scripts\activate

# Linux/Mac
cd ~/pkbm-rapot-main
python3 -m venv venv
source venv/bin/activate
```

Setelah aktif, prompt akan berubah: `(venv) PS D:\...>`

### 3. Install Dependencies

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Verifikasi

```bash
python -c "import reportlab; import pandas; import cv2; print('OK')"
```

### 5. Deaktivasi Virtual Environment

```bash
deactivate
```

> **Tips:** Selalu aktifkan venv sebelum menjalankan atau build aplikasi.

---

## Menjalankan Aplikasi (Development)

```bash
# Pastikan venv aktif
venv\Scripts\activate          # Windows
source venv/bin/activate       # Linux/Mac

# Jalankan GUI
python pkbm_generator_app.py
```

---

## Debugging

### Debug Mode: Aktifkan Console Output

Saat development, jalankan dari terminal untuk melihat output `print()` dan traceback error:

```bash
python pkbm_generator_app.py
```

Semua warning dan error dari generator akan tercetak di console, contoh:
```
WARNING: [KPC13] Siswa 'Lilik Parwati' ada di Nilai tapi TIDAK ADA di Data Siswa - DILEWATI
ERROR: [KPC15A] Error generate rapot untuk 'Nama': some error message
```

### Debug EXE: Aktifkan Console Window

Untuk debug versi EXE, ubah `console=False` menjadi `console=True` di `build_windows.spec`:

```python
exe = EXE(
    ...
    console=True,   # <-- Ubah ke True untuk debug
    ...
)
```

Lalu rebuild. EXE akan menampilkan console window dengan semua output debug.

### Debug Tanpa Rebuild

Jika tidak ingin rebuild, jalankan EXE dari command prompt:

```cmd
cd dist
PKBM_Generator.exe
```

> **Catatan:** Dengan `console=False`, output `print()` tidak terlihat. 
> Untuk EXE production, error/warning ditampilkan di log panel GUI.

### Debug Specific Generator

Test generator secara langsung tanpa GUI:

```python
# test_rapot.py
import sys, os
sys.path.insert(0, '.')
from lib.pdf_generators.rapot_generator import generate_all_rapots

result = generate_all_rapots(
    data_siswa_path='path/ke/data_siswa.xlsx',
    nilai_path='path/ke/nilai_siswa.xlsx',
    output_folder='./output/test'
)

print(f"Generated: {len(result['generated_files'])} files")
print(f"Warnings: {result['warnings']}")
print(f"Errors: {result['errors']}")
```

### Error Handling di GUI

Generator mengembalikan dict berisi:
- `generated_files` - list path PDF yang berhasil dibuat
- `warnings` - list string warning (data tidak lengkap tapi tetap jalan)
- `errors` - list string error (gagal total untuk siswa/kelas tertentu)

Semua warning/error ditampilkan di log panel GUI dengan ikon:
- `⚠` Warning (kuning)
- `❌` Error (merah)

---

## Build ke EXE

### Mode: Single File (Onefile)

Aplikasi di-build sebagai **satu file EXE** yang berisi semua dependencies, assets, dan template.

### Build Otomatis

```bash
# Windows
build_exe.bat

# Linux/Mac
chmod +x build_exe.sh
./build_exe.sh

# Quick build (minimal output)
quick_build.bat
```

### Hasil Build

```
dist/PKBM_Generator.exe    # Single file, siap distribusi
```

### Cara Distribusi

1. Copy `dist/PKBM_Generator.exe` ke komputer target
2. User double-click untuk menjalankan
3. Tidak perlu install Python atau library apapun

### Cara Kerja Onefile

- PyInstaller mem-bundle semua ke dalam satu EXE
- Saat dijalankan, EXE meng-extract ke folder temporary (`sys._MEIPASS`)
- Folder `assets/`, `lib/`, `data/` otomatis tersedia di `sys._MEIPASS`
- Fungsi `get_resource_path()` di `pkbm_generator_app.py` menangani perbedaan path

### Build dengan Debug Console

```python
# Di build_windows.spec, ubah:
console=True    # Tampilkan console window untuk debug
```

---

## Format Excel per Generator

### 1. Rapot Generator

Membutuhkan **2 file Excel**:

#### File 1: Nilai Siswa (`template_nilai_siswa.xlsx`)

| Sheet | Fungsi | Kolom/Format |
|-------|--------|-------------|
| `DataTetap` | Info sekolah | `variabel`, `Nama` (lihat tabel di bawah) |
| `guru` | Data guru | `Kelas`, `Nama` |
| `page2` | Data sikap | `No`, `Dimensi`, `Keterangan` |
| `KPB9`, `KPB10`, ... | Nilai per kelas | `Nama`, `Kelas`, lalu kolom mata pelajaran |

**Sheet `DataTetap` - Variabel yang digunakan:**

| Variabel | Contoh Nilai | Keterangan |
|----------|-------------|------------|
| `Semester` | GENAP | Semester aktif |
| `TahunPelajaran` | 2024/2025 | Tahun ajaran (boleh pakai underscore: `Tahun_Pelajaran`) |
| `Tanggal` | Taipei, 23 Juni 2025 | Tanggal cetak rapot |
| `NamaSekolah` | PKBM PPI Taiwan | Nama sekolah (opsional, default: PKBM PPI Taiwan) |
| `Alamat` | Jln. Rui Guang No. 550 | Alamat sekolah |
| `KodePos` | 11492 | Kode pos (boleh: `Kode_Pos`) |
| `Website` | http://www.pkbmppitaiwan.sch.id | Website (opsional) |
| `Email` | ketua@pkbmppitaiwan.sch.id | Email |
| `Telepon` | +886919427561 | Nomor telepon |
| `NPSN` | P9908880 | Nomor Pokok Sekolah Nasional |

> **Catatan:** Nama variabel case-insensitive dan spasi/underscore diabaikan.
> `TahunPelajaran`, `Tahun_Pelajaran`, `tahunpelajaran` semuanya valid.

**Sheet `guru` - Format:**

| Kelas | Nama |
|-------|------|
| KEPSEK | Ir. Ananda Insan Firdausy |
| KPB9 | Nama Wali Kelas KPB9 |
| KPB10 | Nama Wali Kelas KPB10 |
| ... | ... |

**Sheet Nilai (KPB9, KPB10, KPC13, dst.) - Format:**

| Nama | Kelas | Kelompok Mata Pelajaran Umum merge | Pendidikan Agama Islam | ... |
|------|-------|------------------------------------|----------------------|-----|
| Diah Widya Ningrum | 9 | | 85 | ... |

- Kolom dengan kata `merge` di nama = header kelompok (tidak berisi nilai)
- Kolom lainnya = mata pelajaran dengan nilai numerik
- Baris dengan `Nama` kosong (NaN) otomatis diabaikan

#### File 2: Data Siswa (`template_data_siswa.xlsx`)

| Sheet | Fungsi |
|-------|--------|
| `template` | Template kolom (tidak diproses) |
| `KPB9`, `KPB10`, ... | Data siswa per kelas |

**Sheet per Kelas - Kolom:**

| Kolom | Contoh |
|-------|--------|
| `Nama Peserta Didik` | Diah Widya Ningrum |
| `NISN` | 0012345678 |
| `NIS` | 08880-340001 |
| `Tempat` | Semarang |
| `Tanggal Lahir` | 10-12-2000 |
| `Jenis Kelamin` | Perempuan |
| `Agama` | Islam |
| `Anak Ke` | 2 |
| `Alamat Peserta Didik` | Jl. Merdeka No. 1 |
| `Nomor HP` | 081234567890 |
| ... | ... |

> **PENTING:** Nama sheet di file Nilai dan Data Siswa **harus cocok** (misal: keduanya punya `KPB9`).
> Spasi di nama sheet diabaikan (`KPC 14` = `KPC14`).
> Nama siswa di kedua file **harus sama persis** (case-insensitive).

---

### 2. SKHUPK Generator

Membutuhkan **1 file Excel**:

| Sheet | Fungsi | Kolom |
|-------|--------|-------|
| `DataTetap` | Info sekolah | `variabel`, `Nama` (sama seperti Rapot) |
| Sheet per kelas | Data siswa + nilai UPK | Sesuai template |

---

### 3. Transkrip Nilai Generator

Membutuhkan **1 file Excel** (`Database Transkrip Nilai.xlsx`):

| Sheet | Fungsi |
|-------|--------|
| `DataTetap` | Info sekolah |
| `DataSiswa` | Data identitas siswa |
| `NilaiSiswa` | Data nilai per mata pelajaran |

---

## Arsitektur Path

### Perbedaan Development vs EXE

| Konteks | Base Path | Contoh |
|---------|-----------|--------|
| Development | `os.path.dirname(__file__)` | `D:\pkbm-rapot-main\` |
| EXE (onefile) | `sys._MEIPASS` | `C:\Users\...\AppData\Local\Temp\_MEI12345\` |

### Fungsi Helper

```python
# Di pkbm_generator_app.py
def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller."""
    try:
        base_path = sys._MEIPASS  # EXE mode
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))  # Dev mode
    return os.path.join(base_path, relative_path)
```

### Penggunaan

```python
# Akses logo
logo_path = get_resource_path(os.path.join("assets", "logo.png"))

# Akses template
template_path = get_resource_path("data/rapot/template_data_siswa.xlsx")
```

### Output Folder

Output PDF **TIDAK** di-bundle dalam EXE. Output selalu ditulis ke folder relatif
terhadap working directory user (bukan `sys._MEIPASS`).

---

## Troubleshooting

### Aplikasi tidak bisa dibuka (Windows)

- Klik kanan EXE > Properties > **Unblock**
- Atau: Windows Defender > Allow

### Antivirus mendeteksi sebagai virus

- Ini **false positive** (umum untuk PyInstaller onefile)
- Whitelist/exclude file EXE di antivirus

### Error: "Module not found" saat build

```bash
pip install --upgrade -r requirements.txt
```

### EXE sangat lambat saat pertama kali dibuka

- Normal untuk onefile mode - EXE perlu extract ke temp folder
- Pembukaan kedua dan seterusnya lebih cepat (cached)

### Error: Font tidak ditemukan

Pastikan folder `assets/fonts/` berisi file TTF yang diperlukan:
- `times.ttf`, `timesbd.ttf`, `timesbdit.ttf`, `timesit.ttf`

### Error: "Variabel X tidak ditemukan di DataTetap"

- Periksa sheet `DataTetap` di file Excel
- Pastikan kolom bernama `variabel` dan `Nama`
- Nama variabel tidak case-sensitive, spasi dan underscore diabaikan

### Generate rapot hasilnya 0 file

Kemungkinan penyebab:
1. **Nama sheet tidak cocok** antara file Nilai dan Data Siswa
2. **Nama siswa tidak cocok** antara kedua file
3. **Kolom `Nama` mengandung NaN** (baris kosong di Excel)

Cek log panel di GUI untuk melihat warning/error detail.

### Virtual environment tidak aktif

```bash
# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate

# Cek apakah aktif
which python    # Harus menunjuk ke venv/...
```
