# Panduan Layout & Text Positioning untuk Kartu UPK & Kartu Siswa

## Dimensi Kartu
- **Ukuran**: 85.60 mm × 53.98 mm (standar kartu kredit)
- **Dalam points**: 242.65 × 153.01 points
- **Tanpa margin** - konten memenuhi seluruh halaman

## Sistem Koordinat ReportLab

```
Y (tinggi)
^
|  h (153.01 points)  ┌─────────────────────┐
|                     │                     │
|                     │    KARTU            │
|                     │                     │
|                     │                     │
|  0 (bawah)          └─────────────────────┘
└────────────────────────────────────────────> X (lebar)
                     0 (kiri)    w (242.65)
```

**Penting**: Y=0 adalah BAWAH halaman, Y=h adalah ATAS halaman

## Fungsi Text Positioning

### 1. Text Rata Kiri (Left-aligned)
```python
c.drawString(x, y, "Text")
```
- `x`: posisi horizontal dari kiri (dalam points atau mm)
- `y`: posisi vertikal dari bawah
- Text dimulai dari koordinat (x, y)

**Contoh**:
```python
c.setFont('TimesNewRoman-Bold', 10)
c.drawString(10*mm, 40*mm, "Nama: Budi")
```

### 2. Text Terpusat (Centered)
```python
c.drawCentredString(x, y, "Text")
```
- `x`: posisi horizontal CENTER
- `y`: posisi vertikal dari bawah
- Text terpusat di sekitar koordinat x

**Contoh**:
```python
c.drawCentredString(w/2, 40*mm, "KARTU PELAJAR")  # Terpusat horizontal
```

## Cara Mengatur Positioning

### A. Mengubah Ukuran Text
```python
# Kecil
c.setFont('TimesNewRoman', 4)

# Sedang
c.setFont('TimesNewRoman', 5)

# Besar
c.setFont('TimesNewRoman-Bold', 10)
```

### B. Mengubah Posisi Horizontal (Kiri-Kanan)
```python
# Lebih ke kiri
c.drawString(5*mm, y, "Text")

# Lebih ke kanan
c.drawString(20*mm, y, "Text")

# Terpusat
c.drawCentredString(w/2, y, "Text")
```

### C. Mengubah Posisi Vertikal (Atas-Bawah)
```python
# Lebih ke bawah
c.drawString(x, 10*mm, "Text")

# Lebih ke atas
c.drawString(x, 40*mm, "Text")
```

## Vertical Centering (Memusatkan Vertikal)

Untuk memusatkan text/elemen secara vertikal dalam area tertentu:

```python
# Tentukan area
body_top = 40*mm      # Batas atas area
body_bottom = 10*mm   # Batas bawah area
body_height = body_top - body_bottom

# Hitung center
body_center_y = body_bottom + body_height / 2

# Gunakan untuk positioning
c.drawString(x, body_center_y, "Text di tengah vertikal")
```

### Untuk Multiple Lines (Beberapa Baris)
```python
field_font_size = 5
line_spacing = 5*mm
num_fields = 3
total_height = (num_fields - 1) * line_spacing

# Start dari center, adjusted untuk jumlah field
start_y = body_center_y + total_height / 2

# Loop untuk setiap field
for i, text in enumerate(["Nama", "NIS", "Tempat"]):
    y = start_y - i * line_spacing
    c.drawString(x, y, text)
```

Hasil:
```
Nama      ← start_y
NIS       ← start_y - 1*line_spacing
Tempat    ← start_y - 2*line_spacing
          ← body_center_y (tengah area)
```

## Layout Kartu UPK

```
┌─────────────────────────────────────────┐
│ Logo    KARTU PESERTA UPK    Logo       │ Header (7.5mm)
│         TAHUN PELAJARAN 2024/2025       │
├─────────────────────────────────────────┤
│                                         │
│ ┌─────┐  No. Peserta: ...    ┌──────┐  │
│ │     │  Nama: ...            │ TTD  │  │ Body (centered)
│ │FOTO │  Tempat/Tgl: ...      │Kepsek│  │
│ │3x4  │  NISN: ...            │      │  │
│ └─────┘                        └──────┘  │
│                                         │
├─────────────────────────────────────────┤
│        PKBM PPI TAIWAN                  │ Footer (2.5mm)
└─────────────────────────────────────────┘
```

### Koordinat Penting Kartu UPK
```python
w = 85.60*mm  # Lebar kartu
h = 53.98*mm  # Tinggi kartu

# Header
header_y = h - 2*mm
logo_size = 7*mm

# Body
body_top = header_y - 7.5*mm
body_bottom = 2.5*mm + 5*mm  # margin + space for signature
body_center_y = body_bottom + (body_top - body_bottom) / 2

# Photo (left)
photo_x = 2.5*mm
photo_y = body_center_y - 10*mm  # Centered vertically
photo_w = 15*mm
photo_h = 20*mm

# Data fields (middle)
data_x = photo_x + photo_w + 2*mm
field_font_size = 4.5

# Signature (right)
ttd_x = w - 2.5*mm - 12*mm
ttd_y = body_bottom + 0.5*mm
ttd_w = 12*mm
ttd_h = 8*mm
```

## Layout Kartu Siswa (Merah Putih)

```
┌─────────────────────────────────────────┐
│ ███ KARTU PELAJAR ███                   │ Header Merah (9mm)
│ TAHUN AJARAN 2024/2025                  │ Subtitle (4.5mm)
├─────────────────────────────────────────┤
│                                         │
│ ┌─────┐  Nama: ...           ┌──────┐  │
│ │     │  NIS: ...            │ TTD  │  │ Body (centered)
│ │FOTO │  Tempat/Tgl: ...     │Kepsek│  │
│ │3x4  │                      │      │  │
│ └─────┘                       └──────┘  │
│                                         │
├─────────────────────────────────────────┤
│ ███ PKBM PPI TAIWAN ███                 │ Footer Merah (7mm)
└─────────────────────────────────────────┘
```

### Koordinat Penting Kartu Siswa
```python
w = 85.60*mm
h = 53.98*mm

# Header & Footer
header_h = 9*mm
header_y = h - header_h
subtitle_h = 4.5*mm
subtitle_y = header_y - subtitle_h
footer_h = 7*mm
footer_y = 0

# Body
body_top = subtitle_y - 1.5*mm
body_bottom = footer_h + 1.5*mm
body_center_y = body_bottom + (body_top - body_bottom) / 2

# Photo (left, centered)
photo_x = 3*mm
photo_y = body_center_y - 10*mm
photo_w = 16*mm
photo_h = 20*mm

# Data fields (middle, centered)
data_x = photo_x + photo_w + 2.5*mm
field_font_size = 4.5
line_spacing = 4.5*mm

# Signature (right, centered)
ttd_x = w - 2.5*mm - 12*mm
ttd_y = body_center_y - 4*mm
ttd_w = 12*mm
ttd_h = 8*mm
```

## Contoh Modifikasi

### Ubah Ukuran Font
**File**: `lib/pdf_generators/kartu_upk_generator.py` atau `kartu_siswa_generator.py`

Cari:
```python
field_font_size = 4.5
```

Ubah ke:
```python
field_font_size = 5  # Lebih besar
# atau
field_font_size = 4  # Lebih kecil
```

### Ubah Posisi Text Horizontal
Cari:
```python
data_x = photo_x + photo_w + 2*mm
```

Ubah ke:
```python
data_x = photo_x + photo_w + 3*mm  # Lebih ke kanan
# atau
data_x = photo_x + photo_w + 1*mm  # Lebih ke kiri
```

### Ubah Posisi Text Vertikal
Cari:
```python
start_y = body_center_y + total_height / 2
```

Ubah ke:
```python
start_y = body_center_y + total_height / 2 + 2*mm  # Lebih ke atas
# atau
start_y = body_center_y + total_height / 2 - 2*mm  # Lebih ke bawah
```

### Ubah Spacing Antar Baris
Cari:
```python
line_spacing = 4.5*mm
```

Ubah ke:
```python
line_spacing = 5*mm  # Lebih jauh
# atau
line_spacing = 4*mm  # Lebih dekat
```

## Debugging Tips

### Lihat Batas Area
Tambahkan garis untuk melihat area:
```python
# Gambar garis untuk debugging
c.setStrokeColor(colors.red)
c.setLineWidth(0.1)
c.line(data_x, body_bottom, data_x, body_top)  # Garis vertikal
c.line(margin_x, body_center_y, w-margin_x, body_center_y)  # Garis horizontal center
```

### Print Koordinat
```python
print(f"body_center_y: {body_center_y}")
print(f"start_y: {start_y}")
print(f"data_x: {data_x}")
```

## Unit Konversi

```python
from reportlab.lib.units import mm, cm, inch

# Semua sama:
10*mm == 1*cm == 0.3937*inch

# Contoh:
margin = 2.5*mm    # 2.5 millimeter
spacing = 1*cm     # 1 centimeter
size = 0.5*inch    # 0.5 inch
```

## Checklist Saat Edit

- [ ] Ubah font size? Cek apakah text masih fit dalam area
- [ ] Ubah posisi? Cek apakah tidak overlap dengan elemen lain
- [ ] Ubah spacing? Cek apakah semua field masih visible
- [ ] Test dengan data panjang (nama panjang, tempat panjang)
- [ ] Generate sample PDF dan lihat hasilnya

---

**Catatan**: Semua perubahan dilakukan di dalam fungsi `_draw_kartu_upk()` atau `_draw_kartu_siswa()` di file generator masing-masing.
