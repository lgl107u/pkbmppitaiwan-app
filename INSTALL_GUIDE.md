# Installation Guide - PKBM Generator

## Quick Start

### Windows

```cmd
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python pkbm_generator_app.py
```

### Linux / Mac

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python pkbm_generator_app.py
```

## Dependencies

| Package | Fungsi |
|---------|--------|
| reportlab | Generate PDF |
| pandas | Baca/tulis Excel |
| numpy | Operasi numerik |
| python-docx | Manipulasi DOCX |
| pillow | Manipulasi gambar |
| openpyxl | Baca/tulis Excel (.xlsx) |
| opencv-python | Pemrosesan gambar tanda tangan |

## Troubleshooting

| Error | Solusi |
|-------|--------|
| `pip: command not found` | Aktifkan venv: `source venv/bin/activate` |
| `No module named 'reportlab'` | `pip install -r requirements.txt` |
| `Permission denied` | Gunakan venv, jangan `sudo pip` |

Untuk panduan lengkap: [DEVELOPER_GUIDE.md](DEVELOPER_GUIDE.md)
