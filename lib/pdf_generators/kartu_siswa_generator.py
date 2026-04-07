"""
Kartu Pelajar / Kartu Siswa Generator
Card size: 53.98mm x 85.60mm (portrait orientation)
No margin - card fills entire page

Design: Clean modern portrait ID card with logo, photo, and student info.
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

# Card dimensions (portrait)
CARD_WIDTH = 53.98 * mm
CARD_HEIGHT = 85.60 * mm

# Colors - Merah Putih theme
COLOR_RED = colors.Color(0.75, 0.07, 0.07)         # Indonesian red
COLOR_DARK_RED = colors.Color(0.55, 0.02, 0.02)    # Darker red for accents
COLOR_WHITE = colors.white
COLOR_DARK = colors.Color(0.12, 0.12, 0.12)
COLOR_GRAY = colors.Color(0.55, 0.55, 0.55)

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


def _find_photo(photo_folder, student_name, nis=''):
    """Find a photo file matching 'nama - nis' or just 'nama' in the folder."""
    if not photo_folder or not student_name or not os.path.isdir(photo_folder):
        return None
    
    extensions = ['jpg', 'jpeg', 'png', 'bmp', 'gif']
    student_name_clean = student_name.strip().lower()
    nis_clean = str(nis).strip() if nis else ''
    
    # Build candidate names to match (in priority order)
    candidates = []
    if nis_clean and nis_clean != '-':
        candidates.append(f"{student_name_clean} - {nis_clean}")  # "nama - nis"
        candidates.append(f"{student_name_clean}-{nis_clean}")     # "nama-nis"
    candidates.append(student_name_clean)                           # fallback: just "nama"
    
    for f in os.listdir(photo_folder):
        name, fext = os.path.splitext(f)
        if fext.lower().lstrip('.') in extensions:
            name_lower = name.strip().lower()
            for candidate in candidates:
                if name_lower == candidate:
                    return os.path.join(photo_folder, f)
    
    return None


def _draw_kartu_siswa(c, student, data_tetap, photo_path, logo_path, ttd_kepsek_path=None):
    """
    Draw a single Kartu Pelajar/Siswa on the canvas (portrait).
    
    Layout (53.98 x 85.60 mm):
    ┌──────────────────────────┐
    │  ┌────────────────────┐  │
    │  │                    │  │
    │  │      [LOGO]        │  │
    │  │   KARTU PELAJAR    │  │
    │  │  PKBM PPI TAIWAN   │  │
    │  │                    │  │
    │  │   ┌────────────┐   │  │
    │  │   │            │   │  │
    │  │   │   FOTO     │   │  │
    │  │   │            │   │  │
    │  │   └────────────┘   │  │
    │  │                    │  │
    │  │   NAMA SISWA       │  │
    │  │   Kelas: KPC13     │  │
    │  │   NIS: 12345       │  │
    │  │                    │  │
    │  └────────────────────┘  │
    └──────────────────────────┘
    """
    w = CARD_WIDTH
    h = CARD_HEIGHT
    center_x = w / 2
    
    # ── Prepare student data ──
    nama = str(student.get('Nama', '')).strip()
    
    nis_val = student.get('NIS', student.get('NISN', student.get('nis', '')))
    if pd.isna(nis_val) or nis_val == '':
        nis_str = '-'
    else:
        nis_str = str(int(nis_val)) if isinstance(nis_val, float) else str(nis_val)
    
    kelas_val = student.get('Kelas', student.get('kelas', ''))
    if pd.isna(kelas_val) or kelas_val == '':
        kelas_str = ''
    else:
        kelas_str = str(kelas_val).strip()
    
    # ── WHITE BACKGROUND (no border, fills entire card) ──
    c.setFillColor(COLOR_WHITE)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    
    # ── WATERMARK: Diagonal repeating logo pattern ──
    if logo_path and os.path.exists(logo_path):
        try:
            c.saveState()
            clip_path = c.beginPath()
            clip_path.rect(0, 0, w, h)
            c.clipPath(clip_path, stroke=0)
            c.setFillAlpha(0.05)
            
            wm_size = 18 * mm
            x_spacing = 22 * mm
            y_spacing = 20 * mm
            diagonal_offset = 11 * mm
            
            row = 0
            y = -5 * mm
            while y < h + wm_size:
                x_offset = (row % 2) * diagonal_offset
                x = -5 * mm + x_offset
                while x < w + wm_size:
                    c.saveState()
                    c.translate(x + wm_size / 2, y + wm_size / 2)
                    c.rotate(15)
                    c.drawImage(logo_path, -wm_size / 2, -wm_size / 2,
                               width=wm_size, height=wm_size,
                               preserveAspectRatio=True, mask='auto')
                    c.restoreState()
                    x += x_spacing
                y += y_spacing
                row += 1
            c.restoreState()
        except Exception as e:
            print(f"Watermark error: {e}")
    
    # ── LOGO (centered, top area) ──
    logo_size = 14 * mm
    logo_y = h - 3 * mm - logo_size
    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(logo_path, center_x - logo_size / 2, logo_y,
                       width=logo_size, height=logo_size,
                       preserveAspectRatio=True, mask='auto')
        except:
            pass
    
    # ── HEADER TEXT (red) ──
    header_y = logo_y - 2.5 * mm
    c.setFillColor(COLOR_DARK)
    c.setFont('TimesNewRoman-Bold', 8)
    c.drawCentredString(center_x, header_y, "KARTU PELAJAR")
    
    header_y2 = header_y - 3.5 * mm
    c.setFont('TimesNewRoman-Bold', 6.5)
    c.drawCentredString(center_x, header_y2, "PKBM PPI TAIWAN")
    
    # ── Thin red accent line ──
    line_y = header_y2 - 2.5 * mm
    line_margin = 8 * mm
    c.setStrokeColor(COLOR_RED)
    c.setLineWidth(1)
    c.line(line_margin, line_y, w - line_margin, line_y)
    
    # ── PHOTO (centered, large) ──
    photo_w = 30 * mm
    photo_h = 40 * mm
    photo_x = center_x - photo_w / 2
    photo_y = line_y - 2.5 * mm - photo_h
    
    if photo_path and os.path.exists(photo_path):
        try:
            from PIL import Image
            from io import BytesIO
            
            img = Image.open(photo_path)
            img_w, img_h = img.size
            
            # Target aspect ratio 3:4 (width:height)
            target_ratio = 3.0 / 4.0
            current_ratio = img_w / img_h
            
            if abs(current_ratio - target_ratio) > 0.01:
                if current_ratio > target_ratio:
                    new_w = int(img_h * target_ratio)
                    left = (img_w - new_w) // 2
                    img = img.crop((left, 0, left + new_w, img_h))
                else:
                    new_h = int(img_w / target_ratio)
                    top = (img_h - new_h) // 2
                    img = img.crop((0, top, img_w, top + new_h))
            
            buf = BytesIO()
            img.save(buf, format='PNG')
            buf.seek(0)
            
            # Draw photo without border
            c.drawImage(ImageReader(buf), photo_x, photo_y,
                       width=photo_w, height=photo_h, mask='auto')
        except Exception as e:
            print(f"Error drawing photo: {e}")
            # Fallback placeholder
            c.setStrokeColor(COLOR_GRAY)
            c.setLineWidth(0.4)
            c.rect(photo_x, photo_y, photo_w, photo_h)
            c.setFillColor(COLOR_GRAY)
            c.setFont('TimesNewRoman', 7)
            c.drawCentredString(center_x, photo_y + photo_h / 2 + 2 * mm, "FOTO")
            c.setFont('TimesNewRoman', 5)
            c.drawCentredString(center_x, photo_y + photo_h / 2 - 2 * mm, "3x4")
    else:
        # Placeholder
        c.setStrokeColor(COLOR_GRAY)
        c.setLineWidth(0.4)
        c.rect(photo_x, photo_y, photo_w, photo_h)
        c.setFillColor(COLOR_GRAY)
        c.setFont('TimesNewRoman', 7)
        c.drawCentredString(center_x, photo_y + photo_h / 2 + 2 * mm, "FOTO")
        c.setFont('TimesNewRoman', 5)
        c.drawCentredString(center_x, photo_y + photo_h / 2 - 2 * mm, "3x4")
    
    # ── STUDENT INFO (below photo, centered) ──
    info_y = photo_y - 4 * mm
    
    # Name (bold, larger)
    c.setFillColor(COLOR_DARK)
    c.setFont('TimesNewRoman-Bold', 9)
    c.drawCentredString(center_x, info_y, nama)
    
    # Thin red line under name
    name_w = c.stringWidth(nama, 'TimesNewRoman-Bold', 9)
    underline_half = max(name_w / 2, 15 * mm)
    c.setStrokeColor(COLOR_RED)
    c.setLineWidth(0.8)
    c.line(center_x - underline_half, info_y - 1.2 * mm,
           center_x + underline_half, info_y - 1.2 * mm)
    
    # Kelas and NIS
    detail_y = info_y - 5 * mm
    c.setFont('TimesNewRoman', 7)
    c.setFillColor(COLOR_DARK)
    
    if kelas_str:
        c.drawCentredString(center_x, detail_y, kelas_str)
        detail_y -= 3.5 * mm
    
    c.drawCentredString(center_x, detail_y, nis_str)


def generate_single_kartu_siswa(student, data_tetap, photo_path, logo_path, output_path, ttd_kepsek_path=None):
    """Generate a single Kartu Siswa PDF."""
    c_obj = canvas.Canvas(output_path, pagesize=(CARD_WIDTH, CARD_HEIGHT))
    c_obj.setPageSize((CARD_WIDTH, CARD_HEIGHT))
    
    _draw_kartu_siswa(c_obj, student, data_tetap, photo_path, logo_path, ttd_kepsek_path)
    
    c_obj.save()
    return output_path


def generate_all_kartu_siswa(excel_file_path, output_folder, photo_folder=None,
                              tahun_ajaran=None, ttd_kepsek_path=None, progress_callback=None):
    """
    Generate Kartu Siswa/Pelajar PDFs for all students from Excel file.
    
    Args:
        excel_file_path: Path to Excel file (same format as rapot Data Siswa)
        output_folder: Output folder for generated PDFs
        photo_folder: Folder containing student photos (matched by student name)
        tahun_ajaran: Override tahun ajaran
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
        data_tetap = {}
        data_tetap_df = df.get('DataTetap', pd.DataFrame())
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
            
            # Find photo by student name and NIS
            photo_path = None
            if photo_folder:
                nis_val_photo = row.get('NIS', row.get('NISN', row.get('nis', '')))
                nis_for_photo = ''
                if not pd.isna(nis_val_photo) and nis_val_photo != '':
                    nis_for_photo = str(int(nis_val_photo)) if isinstance(nis_val_photo, float) else str(nis_val_photo)
                photo_path = _find_photo(photo_folder, nama, nis_for_photo)
            
            student = row.to_dict()
            
            try:
                safe_name = re.sub(r'[^\w\s-]', '', nama).strip()
                # Filename: {nisn}_{nama} or no_nisn_{nama} if NISN empty
                nisn_val = student.get('NIS', student.get('NISN', student.get('nis', '')))
                if pd.isna(nisn_val) or nisn_val == '':
                    nisn_str = 'no_nisn'
                else:
                    nisn_str = str(int(nisn_val)) if isinstance(nisn_val, float) else str(nisn_val)
                
                pdf_filename = f"{nisn_str}_{safe_name}.pdf"
                pdf_path = os.path.join(output_folder, pdf_filename)
                
                generate_single_kartu_siswa(student, data_tetap, photo_path, logo_path, pdf_path, ttd_kepsek_path)
                generated_files.append(pdf_path)
                
            except Exception as e:
                print(f"Error generating Kartu Siswa for {nama}: {e}")
                import traceback
                traceback.print_exc()
            
            if progress_callback:
                progress_callback(idx + 1, total_students, nama)
        
        print(f"Generated {len(generated_files)} Kartu Siswa files")
        
    except Exception as e:
        print(f"Error in generate_all_kartu_siswa: {str(e)}")
        import traceback
        traceback.print_exc()
    
    return generated_files
