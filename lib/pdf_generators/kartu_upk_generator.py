"""
Kartu Peserta UPK (Ujian Pendidikan Kesetaraan) Generator
Card size: 85.60mm x 53.98mm (standard credit card)
No margin - card fills entire page

Design: Professional card with blue-gold accent, QR code signature
"""

from reportlab.lib.units import mm, cm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import os
import re
import io
import glob
import pandas as pd

# Card dimensions
CARD_WIDTH = 85.60 * mm
CARD_HEIGHT = 53.98 * mm

# Colors - Professional blue/dark theme for UPK
COLOR_NAVY = colors.Color(0.10, 0.18, 0.36)       # Dark navy header
COLOR_GOLD = colors.Color(0.75, 0.60, 0.20)        # Gold accent
COLOR_LIGHT_BLUE = colors.Color(0.90, 0.93, 0.98)  # Light blue tint
COLOR_DARK = colors.Color(0.12, 0.12, 0.12)
COLOR_GRAY = colors.Color(0.45, 0.45, 0.45)

# Register Times New Roman fonts
def _register_fonts():
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    font_dir = os.path.join(project_root, "assets", "fonts")
    
    fonts_to_register = {
        'TimesNewRoman': 'times.ttf',
        'TimesNewRoman-Bold': 'timesbd.ttf',
        'TimesNewRoman-BoldItalic': 'timesbdit.ttf',
        'TimesNewRoman-Italic': 'timesit.ttf',
    }
    
    for font_name, font_file in fonts_to_register.items():
        font_path = os.path.join(font_dir, font_file)
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
            except:
                pass

_register_fonts()


def _get_logo_path():
    """Get path to PKBM PPI Taiwan logo."""
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(project_root, "assets", "logo.png")


def _generate_qr_image(data_str, box_size=4, border=0):
    """Generate QR code as PIL Image bytes for embedding in PDF."""
    try:
        import qrcode
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L,
                           box_size=box_size, border=border)
        qr.add_data(data_str)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        buf.seek(0)
        return buf
    except Exception as e:
        print(f"Error generating QR code: {e}")
        return None


def _extract_photo_code(no_peserta):
    """
    Extract photo code from No Peserta Ujian.
    e.g., 'P9908880-PC-001' -> 'pc001'
    Takes everything after the first dash that starts with PC/PA/PB, removes dashes, lowercases.
    """
    if not no_peserta or pd.isna(no_peserta):
        return None
    
    no_peserta = str(no_peserta).strip()
    
    # Find the PC/PA/PB part and number
    match = re.search(r'(P[ABC])-?(\d+)', no_peserta, re.IGNORECASE)
    if match:
        return (match.group(1) + match.group(2)).lower()
    
    # Fallback: take last part after last dash
    parts = no_peserta.split('-')
    if len(parts) >= 2:
        return (parts[-2] + parts[-1]).replace('-', '').lower()
    
    return no_peserta.lower()


def _find_photo(photo_folder, photo_code):
    """Find a photo file matching the photo code in the folder."""
    if not photo_folder or not photo_code or not os.path.isdir(photo_folder):
        return None
    
    # Common image extensions
    extensions = ['jpg', 'jpeg', 'png', 'bmp', 'gif']
    
    for ext in extensions:
        # Try exact match (case-insensitive)
        pattern = os.path.join(photo_folder, f"{photo_code}.{ext}")
        matches = glob.glob(pattern)
        if matches:
            return matches[0]
        
        # Try case-insensitive on Windows
        for f in os.listdir(photo_folder):
            name, fext = os.path.splitext(f)
            if name.lower() == photo_code.lower() and fext.lower().lstrip('.') in extensions:
                return os.path.join(photo_folder, f)
    
    return None


def _format_date_id(dt):
    """Format datetime to Indonesian date string."""
    if pd.isna(dt):
        return ""
    bln_id = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    try:
        if hasattr(dt, 'day'):
            return f"{dt.day} {bln_id[dt.month - 1]} {dt.year}"
        return str(dt)
    except:
        return str(dt)


def _draw_kartu_upk(c, student, data_tetap, photo_path, logo_path, ttd_kepsek_path=None):
    """
    Draw a single Kartu UPK on the canvas.
    
    Layout (85.60 x 53.98 mm):
    ┌──────────────────────────────────────────────┐
    │ [Logo] KARTU PESERTA UJIAN ... PAKET C [Logo]│  ← Navy header (13mm)
    │        TAHUN PELAJARAN 2024/2025             │
    │ ═══════════ gold accent line ═══════════════ │
    │                                              │
    │ No. Peserta  : P9908880-PC-001               │  ← Data (6pt, spaced)
    │ Nama         : Aan Antisah                   │
    │ TTL          : Cirebon, 24 Desember 1988     │
    │ NISN         : -                             │
    │                                              │
    │ ┌──────┐       Taiwan, 1 Mei 2025            │
    │ │      │       Kepala PKBM PPI Taiwan        │
    │ │ FOTO │       [QR CODE]                     │  ← Photo + centered QR
    │ │      │       Nama Kepsek                   │
    │ └──────┘       NIP. xxx                      │
    │ ═════════════ PKBM PPI Taiwan ═══════════════│  ← Footer (4mm)
    └──────────────────────────────────────────────┘
    """
    w = CARD_WIDTH
    h = CARD_HEIGHT
    
    # ── BACKGROUND ──
    c.setFillColor(COLOR_LIGHT_BLUE)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    
    tahun = data_tetap.get('Tahun', '2024/2025')
    sekolah = data_tetap.get('Sekolah', 'PKBM PPI Taiwan')
    
    # ── HEADER BAND (Navy, taller to fit text between logos) ──
    header_h = 13 * mm
    header_y = h - header_h
    
    c.setFillColor(COLOR_NAVY)
    c.rect(0, header_y, w, header_h, fill=1, stroke=0)
    
    # Gold accent line at bottom of header
    c.setStrokeColor(COLOR_GOLD)
    c.setLineWidth(1.0)
    c.line(0, header_y, w, header_y)
    
    # ── LOGOS in header (inside navy band, not overlapping text) ──
    logo_size = 9 * mm
    logo_margin = 2 * mm
    if logo_path and os.path.exists(logo_path):
        try:
            logo_y = header_y + (header_h - logo_size) / 2
            c.drawImage(logo_path, logo_margin, logo_y,
                       width=logo_size, height=logo_size, preserveAspectRatio=True, mask='auto')
            c.drawImage(logo_path, w - logo_margin - logo_size, logo_y,
                       width=logo_size, height=logo_size, preserveAspectRatio=True, mask='auto')
        except Exception as e:
            print(f"Error drawing logo: {e}")
    
    # ── HEADER TEXT (between logos, sized to fit) ──
    text_left = logo_margin + logo_size + 1 * mm
    text_right = w - logo_margin - logo_size - 1 * mm
    text_center = (text_left + text_right) / 2
    text_avail = text_right - text_left
    
    c.setFillColor(colors.white)
    
    # Line 1: Title - auto-size to fit between logos
    title_text = "KARTU PESERTA UJIAN PENDIDIKAN KESETARAAN"
    title_size = 6.5
    title_w = c.stringWidth(title_text, 'TimesNewRoman-Bold', title_size)
    while title_w > text_avail and title_size > 4:
        title_size -= 0.25
        title_w = c.stringWidth(title_text, 'TimesNewRoman-Bold', title_size)
    
    c.setFont('TimesNewRoman-Bold', title_size)
    c.drawCentredString(text_center, header_y + header_h / 2 + 1.5 * mm, title_text)
    
    # Line 2: Tahun - same size as line 1
    c.setFont('TimesNewRoman-Bold', title_size)
    c.drawCentredString(text_center, header_y + header_h / 2 - 2.5 * mm, f"TAHUN PELAJARAN {tahun}")
    
    # ── FOOTER BAND (Navy) ──
    footer_h = 4 * mm
    c.setFillColor(COLOR_NAVY)
    c.rect(0, 0, w, footer_h, fill=1, stroke=0)
    
    # Gold line on top of footer
    c.setStrokeColor(COLOR_GOLD)
    c.setLineWidth(0.5)
    c.line(0, footer_h, w, footer_h)
    
    # Footer text
    c.setFillColor(colors.white)
    c.setFont('TimesNewRoman-Bold', 5)
    c.drawCentredString(w / 2, footer_h / 2 - 0.5 * mm, sekolah.upper())
    
    # ── Prepare student data ──
    nama = str(student.get('Nama', '')).strip()
    tempat = str(student.get('Tempat', '')).strip().capitalize()
    tgl_lahir = _format_date_id(student.get('Tanggal Lahir', ''))
    tempat_tgl = f"{tempat}, {tgl_lahir}" if tempat and tgl_lahir else tempat or tgl_lahir
    
    nisn = student.get('NISN', '')
    if pd.isna(nisn) or nisn == '':
        nisn_str = '-'
    else:
        nisn_str = str(int(nisn)) if isinstance(nisn, float) else str(nisn)
    
    no_peserta = str(student.get('No peserta Ujian', '')).strip()
    
    # ── DATA FIELDS (bigger font, more spacing from header) ──
    body_top = header_y - 1 * mm
    body_bottom = footer_h + 0.5 * mm
    
    fields = [
        ("No. Peserta", no_peserta),
        ("Nama", nama),
        ("Tempat/Tgl. Lahir", tempat_tgl),
        ("NISN", nisn_str),
    ]
    
    field_font_size = 6
    line_spacing = 3.8 * mm
    data_x = 3 * mm
    label_width = 18 * mm
    
    # Start fields 3mm below header line (more breathing room)
    start_y = body_top - 3 * mm
    
    for i, (label, value) in enumerate(fields):
        y = start_y - i * line_spacing
        
        c.setFillColor(COLOR_DARK)
        c.setFont('TimesNewRoman-Bold', field_font_size)
        c.drawString(data_x, y, label)
        
        c.setFont('TimesNewRoman', field_font_size)
        colon_x = data_x + label_width
        c.drawString(colon_x, y, ":")
        
        val_font_size = field_font_size
        if len(value) > 28:
            val_font_size = 5
        
        c.setFont('TimesNewRoman', val_font_size)
        value_x = colon_x + 2 * mm
        c.drawString(value_x, y, value)
    
    # ── BOTTOM SECTION: Photo left, Kepsek+QR right ──
    # Photo
    photo_w = 14 * mm
    photo_h = 17 * mm
    photo_x = 3 * mm
    photo_y = body_bottom + 0.5 * mm
    
    c.setStrokeColor(COLOR_GRAY)
    c.setLineWidth(0.4)
    c.rect(photo_x, photo_y, photo_w, photo_h)
    
    if photo_path and os.path.exists(photo_path):
        try:
            c.drawImage(photo_path, photo_x, photo_y,
                       width=photo_w, height=photo_h, preserveAspectRatio=True, mask='auto')
        except Exception as e:
            print(f"Error drawing photo: {e}")
    else:
        c.setFillColor(COLOR_GRAY)
        c.setFont('TimesNewRoman', 5)
        c.drawCentredString(photo_x + photo_w / 2, photo_y + photo_h / 2 + 1.5 * mm, "FOTO")
        c.setFont('TimesNewRoman', 4)
        c.drawCentredString(photo_x + photo_w / 2, photo_y + photo_h / 2 - 1 * mm, "3x4")
    
    # ── KEPSEK SIGNATURE with QR code (bottom right, all centered) ──
    kepsek_name = data_tetap.get('Kepsek', '')
    kepsek_nip = data_tetap.get('NIP', '')
    tanggal_kartu = data_tetap.get('Tanggal Kartu', '')
    
    # Center of signature block - shifted more to the right of the card
    sig_right = w - 2 * mm
    sig_center = w - 18 * mm
    
    # Calculate vertical layout from top of photo area down
    sig_top = photo_y + photo_h  # Align with top of photo
    
    # Tanggal (italic, centered)
    c.setFillColor(COLOR_DARK)
    if tanggal_kartu:
        c.setFont('TimesNewRoman-Bold', 4.5)
        c.drawCentredString(sig_center, sig_top - 1 * mm, str(tanggal_kartu))
    
    # "Kepala PKBM PPI Taiwan"
    c.setFont('TimesNewRoman-Bold', 4.5)
    c.drawCentredString(sig_center, sig_top - 3.5 * mm, "Kepala PKBM PPI Taiwan")
    
    # TTD/QR Code (centered) - use TTD image if provided, otherwise QR code
    sig_size = 7 * mm
    sig_y = sig_top - 4.5 * mm - sig_size
    
    if ttd_kepsek_path and os.path.exists(ttd_kepsek_path):
        # Draw TTD signature image from app
        try:
            c.drawImage(ttd_kepsek_path, sig_center - sig_size / 2, sig_y,
                       width=sig_size, height=sig_size, preserveAspectRatio=True, mask='auto')
        except Exception as e:
            print(f"Error drawing TTD image: {e}")
    else:
        # Fallback to QR code
        qr_data = f"{no_peserta}|{nama}|{nisn_str}"
        qr_buf = _generate_qr_image(qr_data, box_size=6, border=0)
        if qr_buf:
            try:
                c.drawImage(ImageReader(qr_buf), sig_center - sig_size / 2, sig_y,
                           width=sig_size, height=sig_size, mask='auto')
            except Exception as e:
                print(f"Error drawing QR: {e}")
    
    # Kepsek name (underlined, centered below signature)
    name_y = sig_y - 2 * mm
    c.setFillColor(COLOR_DARK)
    c.setFont('TimesNewRoman-Bold', 4.5)
    c.drawCentredString(sig_center, name_y, kepsek_name)
    name_w = c.stringWidth(kepsek_name, 'TimesNewRoman-Bold', 4.5)
    c.setStrokeColor(COLOR_DARK)
    c.setLineWidth(0.3)
    c.line(sig_center - name_w / 2, name_y - 0.5 * mm, sig_center + name_w / 2, name_y - 0.5 * mm)
    
    # NIP (centered below name)
    c.setFont('TimesNewRoman', 3.5)
    c.drawCentredString(sig_center, name_y - 2 * mm, f"NIP. {kepsek_nip}")


def generate_single_kartu_upk(student, data_tetap, photo_path, logo_path, output_path, ttd_kepsek_path=None):
    """Generate a single Kartu UPK PDF."""
    c_obj = canvas.Canvas(output_path, pagesize=(CARD_WIDTH, CARD_HEIGHT))
    c_obj.setPageSize((CARD_WIDTH, CARD_HEIGHT))
    
    _draw_kartu_upk(c_obj, student, data_tetap, photo_path, logo_path, ttd_kepsek_path)
    
    c_obj.save()
    return output_path


def generate_all_kartu_upk(excel_file_path, output_folder, photo_folder=None, 
                            tahun_ajaran=None, ttd_kepsek_path=None, progress_callback=None):
    """
    Generate Kartu UPK PDFs for all students from Excel file.
    
    Args:
        excel_file_path: Path to Excel file with DataTetap and DataSiswa sheets
        output_folder: Output folder for generated PDFs
        photo_folder: Folder containing student photos (matched by No Peserta code)
        tahun_ajaran: Override tahun ajaran (optional, uses Excel data if not provided)
        ttd_kepsek_path: Path to Kepsek signature image (optional)
        progress_callback: Optional callback(current, total, student_name)
    
    Returns:
        List of generated file paths
    """
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    
    generated_files = []
    
    try:
        os.makedirs(output_folder, exist_ok=True)
        
        # Load Excel data
        df = pd.read_excel(excel_file_path, sheet_name=None)
        
        # Get DataTetap
        data_tetap_df = df.get('DataTetap', pd.DataFrame())
        data_tetap = {}
        if not data_tetap_df.empty:
            data_tetap_df.columns = data_tetap_df.columns.str.strip()
            for _, row in data_tetap_df.iterrows():
                var_name = str(row.iloc[0]).strip()
                var_val = str(row.iloc[1]).strip() if len(row) > 1 else ''
                data_tetap[var_name] = var_val
        
        # Override tahun ajaran if provided
        if tahun_ajaran:
            data_tetap['Tahun'] = tahun_ajaran
        
        # Get DataSiswa
        data_siswa = df.get('DataSiswa', pd.DataFrame())
        if data_siswa.empty:
            print("No student data found in DataSiswa sheet")
            return generated_files
        
        data_siswa.columns = data_siswa.columns.str.strip()
        total_students = len(data_siswa)
        
        logo_path = _get_logo_path()
        
        # Generate for each student
        for idx, row in data_siswa.iterrows():
            nama = str(row.get('Nama', f'Student_{idx}')).strip()
            no_peserta = str(row.get('No peserta Ujian', '')).strip()
            
            # Find photo
            photo_path = None
            if photo_folder:
                photo_code = _extract_photo_code(no_peserta)
                if photo_code:
                    photo_path = _find_photo(photo_folder, photo_code)
            
            student = row.to_dict()
            
            try:
                safe_name = re.sub(r'[^\w\s-]', '', nama).strip()
                # Filename: {no_peserta}_{nama}
                pdf_filename = f"{no_peserta}_{safe_name}.pdf"
                pdf_path = os.path.join(output_folder, pdf_filename)
                
                generate_single_kartu_upk(student, data_tetap, photo_path, logo_path, pdf_path, ttd_kepsek_path)
                generated_files.append(pdf_path)
                
            except Exception as e:
                print(f"Error generating Kartu UPK for {nama}: {e}")
                import traceback
                traceback.print_exc()
            
            if progress_callback:
                progress_callback(idx + 1, total_students, nama)
        
        print(f"Generated {len(generated_files)} Kartu UPK files")
        
    except Exception as e:
        print(f"Error in generate_all_kartu_upk: {str(e)}")
        import traceback
        traceback.print_exc()
    
    return generated_files
