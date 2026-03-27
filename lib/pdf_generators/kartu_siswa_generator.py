"""
Kartu Pelajar / Kartu Siswa Generator
Card size: 85.60mm x 53.98mm (standard credit card)
No margin - card fills entire page

Design: Modern card inspired by green card reference, adapted with
merah putih (red-white) Indonesian theme. Photo on left, data on right,
QR code signature, professional compact layout.
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

# Colors - Merah Putih theme (modern)
COLOR_RED = colors.Color(0.75, 0.07, 0.07)         # Indonesian red
COLOR_DARK_RED = colors.Color(0.55, 0.02, 0.02)    # Darker red
COLOR_LIGHT_PINK = colors.Color(1.0, 0.95, 0.95)   # Very light pink bg
COLOR_WHITE = colors.white
COLOR_DARK = colors.Color(0.12, 0.12, 0.12)
COLOR_GRAY = colors.Color(0.50, 0.50, 0.50)

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


def _draw_barcode_on_canvas(c, data_str, x, y, bar_height=6*mm, target_width=None):
    """Draw Code128 barcode directly on canvas at (x, y) with optional fixed target width."""
    try:
        from reportlab.graphics.barcode import code128
        
        # Create barcode with default barWidth first to measure
        barcode = code128.Code128(data_str, barWidth=0.45, barHeight=bar_height,
                                  humanReadable=False, quiet=False)
        natural_width = barcode.width
        
        if target_width and natural_width > 0:
            # Scale barWidth so total width matches target_width
            scale = target_width / natural_width
            barcode = code128.Code128(data_str, barWidth=0.45 * scale, barHeight=bar_height,
                                      humanReadable=False, quiet=False)
        
        barcode.drawOn(c, x, y)
        return barcode.width
    except Exception as e:
        print(f"Error drawing barcode: {e}")
        return 0


def _draw_kartu_siswa(c, student, data_tetap, photo_path, logo_path, ttd_kepsek_path=None):
    """
    Draw a single Kartu Pelajar/Siswa on the canvas.
    
    Layout (85.60 x 53.98 mm):
    ┌══════════════════════════════════════════════┐
    │  [Logo] KARTU PELAJAR                        │ ← Red header (9mm)
    │         PKBM PPI Taiwan                      │
    │──────────────────────────────────────────────│
    │  ┌──────┐  Nama Siswa (bold, big)            │
    │  │      │  NIS      : 12345                  │ ← Photo left, data right
    │  │ FOTO │  Kelas    : KPC13                  │
    │  │ [wm] │  TTL      : Jakarta, 15 Jan 2000   │
    │  └──────┘                                    │
    │                        Kepala PKBM PPI Taiwan│
    │  ║║║ BARCODE ║║║       [QR CODE]             │ ← Barcode left, QR right
    │                        Nama Kepsek           │
    │                        NIP. xxx              │
    │▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓│ ← Full red footer (4mm)
    └══════════════════════════════════════════════┘
    """
    w = CARD_WIDTH
    h = CARD_HEIGHT
    
    sekolah = data_tetap.get('Sekolah', 'PKBM PPI Taiwan')
    
    # ── WHITE BACKGROUND ──
    c.setFillColor(colors.white)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    
    # ── FULL RED FOOTER BAND ──
    footer_h = 4 * mm
    c.setFillColor(COLOR_RED)
    c.rect(0, 0, w, footer_h, fill=1, stroke=0)
    c.setFillColor(COLOR_WHITE)
    c.setFont('TimesNewRoman-Bold', 4.5)
    c.drawCentredString(w / 2, footer_h / 2 - 0.5 * mm, sekolah.upper())
    
    # ── RED HEADER BAND (bigger, proportional) ──
    header_h = 10 * mm
    header_y = h - header_h
    
    c.setFillColor(COLOR_RED)
    c.rect(0, header_y, w, header_h, fill=1, stroke=0)
    
    # Darker red line at bottom of header
    c.setStrokeColor(COLOR_DARK_RED)
    c.setLineWidth(0.5)
    c.line(0, header_y, w, header_y)
    
    # Logo in header (bigger, proportional)
    logo_size = 8 * mm
    logo_x = 2.5 * mm
    logo_y_pos = header_y + (header_h - logo_size) / 2
    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(logo_path, logo_x, logo_y_pos,
                       width=logo_size, height=logo_size,
                       preserveAspectRatio=True, mask='auto')
        except:
            pass
    
    # Header text - bigger font, tight spacing
    text_x = logo_x + logo_size + 2 * mm
    header_font = 9
    c.setFillColor(COLOR_WHITE)
    
    c.setFont('TimesNewRoman-Bold', header_font)
    c.drawString(text_x, header_y + header_h / 2 + 1 * mm, "KARTU PELAJAR")
    c.setFont('TimesNewRoman-Bold', header_font)
    c.drawString(text_x, header_y + header_h / 2 - 3 * mm, sekolah.upper())
    
    # ── WATERMARK: Diagonal repeating logo pattern (white body area only) ──
    body_top = header_y
    body_bottom = footer_h
    
    if logo_path and os.path.exists(logo_path):
        try:
            c.saveState()
            
            # Clip to body area only (between header and footer)
            clip_path = c.beginPath()
            clip_path.rect(0, body_bottom, w, body_top - body_bottom)
            c.clipPath(clip_path, stroke=0)
            
            c.setFillAlpha(0.07)
            
            # Larger logos to cover entire white background, 3-4 per row
            wm_size = 20 * mm
            x_spacing = 24 * mm
            y_spacing = 22 * mm
            
            # Diagonal offset for alternating rows
            diagonal_offset = 12 * mm
            
            # Calculate grid with diagonal pattern - extend beyond edges
            x_start = -5 * mm
            y_start = body_bottom - 5 * mm
            
            row = 0
            y = y_start
            while y < body_top + wm_size:
                # Alternate x offset for diagonal effect
                x_offset = (row % 2) * diagonal_offset
                x = x_start + x_offset
                
                while x < w + wm_size:
                    # Rotate logo for anti-counterfeiting effect
                    c.saveState()
                    c.translate(x + wm_size/2, y + wm_size/2)
                    c.rotate(15)  # 15 degree rotation
                    c.drawImage(logo_path, -wm_size/2, -wm_size/2,
                               width=wm_size, height=wm_size,
                               preserveAspectRatio=True, mask='auto')
                    c.restoreState()
                    x += x_spacing
                
                y += y_spacing
                row += 1
            
            c.restoreState()
        except Exception as e:
            print(f"Watermark error: {e}")
    
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
    
    tempat = str(student.get('Tempat', student.get('tempat', ''))).strip()
    if tempat.lower() == 'nan':
        tempat = ''
    else:
        tempat = tempat.capitalize()
    
    tgl_lahir = _format_date_id(student.get('Tanggal Lahir', student.get('tanggal lahir', '')))
    tempat_tgl = f"{tempat}, {tgl_lahir}" if tempat and tgl_lahir else tempat or tgl_lahir
    
    # ── PHOTO (left side) ──
    photo_w = 17 * mm
    photo_h = 21 * mm
    photo_x = 3 * mm
    photo_y = body_top - photo_h - 2 * mm
    
    # Ensure photo stays above barcode area
    min_photo_y = footer_h + 9 * mm
    if photo_y < min_photo_y:
        photo_y = min_photo_y
        photo_h = body_top - 2 * mm - photo_y
    
    if photo_path and os.path.exists(photo_path):
        # Photo exists - no border, auto-resize and crop to 3x4 aspect ratio
        try:
            from PIL import Image
            from io import BytesIO
            
            # Open and process image
            img = Image.open(photo_path)
            img_w, img_h = img.size
            
            # Target aspect ratio 3:4 (width:height)
            target_ratio = 3.0 / 4.0
            current_ratio = img_w / img_h
            
            # Crop to 3:4 if needed
            if abs(current_ratio - target_ratio) > 0.01:
                if current_ratio > target_ratio:
                    # Image too wide, crop width
                    new_w = int(img_h * target_ratio)
                    left = (img_w - new_w) // 2
                    img = img.crop((left, 0, left + new_w, img_h))
                else:
                    # Image too tall, crop height
                    new_h = int(img_w / target_ratio)
                    top = (img_h - new_h) // 2
                    img = img.crop((0, top, img_w, top + new_h))
            
            # Save to buffer
            buf = BytesIO()
            img.save(buf, format='PNG')
            buf.seek(0)
            
            # Draw image without border
            c.drawImage(ImageReader(buf), photo_x, photo_y,
                       width=photo_w, height=photo_h, mask='auto')
        except Exception as e:
            print(f"Error drawing photo: {e}")
            # Fallback to placeholder with border
            c.setStrokeColor(COLOR_GRAY)
            c.setLineWidth(0.4)
            c.rect(photo_x, photo_y, photo_w, photo_h)
            c.setFillColor(COLOR_GRAY)
            c.setFont('TimesNewRoman', 5)
            c.drawCentredString(photo_x + photo_w / 2, photo_y + photo_h / 2 + 1.5 * mm, "FOTO")
            c.setFont('TimesNewRoman', 4)
            c.drawCentredString(photo_x + photo_w / 2, photo_y + photo_h / 2 - 1 * mm, "3x4")
    else:
        # No photo - show placeholder with border
        c.setStrokeColor(COLOR_GRAY)
        c.setLineWidth(0.4)
        c.rect(photo_x, photo_y, photo_w, photo_h)
        c.setFillColor(COLOR_GRAY)
        c.setFont('TimesNewRoman', 5)
        c.drawCentredString(photo_x + photo_w / 2, photo_y + photo_h / 2 + 1.5 * mm, "FOTO")
        c.setFont('TimesNewRoman', 4)
        c.drawCentredString(photo_x + photo_w / 2, photo_y + photo_h / 2 - 1 * mm, "3x4")
    
    # ── DATA FIELDS (right of photo) ──
    data_x = photo_x + photo_w + 3 * mm
    
    # Student name (bigger, bold) - aligned with top of photo
    name_y = photo_y + photo_h - 2.5 * mm
    c.setFillColor(COLOR_DARK)
    c.setFont('TimesNewRoman-Bold', 10)
    c.drawString(data_x, name_y, nama)
    
    # Build fields list
    field_font_size = 5.5
    line_spacing = 3.5 * mm
    label_width = 14 * mm
    
    field_start_y = name_y - 4.5 * mm
    current_y = field_start_y
    
    # NIS field
    c.setFillColor(COLOR_DARK)
    c.setFont('TimesNewRoman-Bold', field_font_size)
    c.drawString(data_x, current_y, "NIS")
    colon_x = data_x + label_width
    c.drawString(colon_x, current_y, ":")
    c.setFont('TimesNewRoman', field_font_size)
    c.drawString(colon_x + 1.5 * mm, current_y, nis_str)
    current_y -= line_spacing
    
    # Kelas field (if exists)
    if kelas_str:
        c.setFont('TimesNewRoman-Bold', field_font_size)
        c.drawString(data_x, current_y, "Kelas")
        c.drawString(colon_x, current_y, ":")
        c.setFont('TimesNewRoman', field_font_size)
        c.drawString(colon_x + 1.5 * mm, current_y, kelas_str)
        current_y -= line_spacing
    
    # Tempat/Tanggal Lahir field (2 lines for label, data on first line)
    c.setFont('TimesNewRoman-Bold', field_font_size)
    c.drawString(data_x, current_y, "Tempat/")
    # Draw colon and data on same line as "Tempat/"
    ttl_colon_x = data_x + c.stringWidth("Tempat/", 'TimesNewRoman-Bold', field_font_size) + 0.5 * mm
    # c.drawString(ttl_colon_x, current_y, ":")
    c.drawString(colon_x, current_y, ":")
    c.setFont('TimesNewRoman', field_font_size)
    # c.drawString(ttl_colon_x + 1.5 * mm, current_y, tempat_tgl)
    c.drawString(colon_x + 1.5 * mm, current_y, tempat_tgl)
    # Second line just label
    current_y -= 2.5 * mm
    c.setFont('TimesNewRoman-Bold', field_font_size)
    c.drawString(data_x, current_y, "Tanggal Lahir")
    
    
    # ── BARCODE BAR (Code128) for NIS (centered below photo, max 2.8cm) ──
    barcode_y = photo_y - 6 * mm
    barcode_max_width = 20 * mm  # 2.8cm
    barcode_margin = 1 * mm  # 0.1cm margin on each side
    
    # If barcode fits under photo (17mm), center it. Otherwise use full 2.8cm width
    if barcode_max_width + 2 * barcode_margin <= photo_w:
        # Barcode fits under photo - center it
        barcode_actual_width = photo_w - 2 * barcode_margin
        barcode_x = photo_x + barcode_margin
    else:
        # Use full 2.8cm width
        barcode_actual_width = barcode_max_width
        barcode_x = photo_x + (photo_w - barcode_max_width) / 2
    
    _draw_barcode_on_canvas(c, nis_str, barcode_x, barcode_y, bar_height=5*mm, target_width=barcode_actual_width)
    
    # ── KEPSEK SIGNATURE with TTD/QR code (bottom right, same as UPK) ──
    kepsek_name = data_tetap.get('Kepsek', '')
    kepsek_nip = data_tetap.get('NIP', '')
    
    # Signature center - right side of card
    sig_center = w - 18 * mm
    
    # "Kepala PKBM PPI Taiwan" at top of signature area (bigger font)
    sig_top = barcode_y + 5 * mm
    c.setFillColor(COLOR_DARK)
    c.setFont('TimesNewRoman-Bold', 5)
    c.drawCentredString(sig_center, sig_top, "Kepala PKBM PPI Taiwan")
    
    # TTD/QR Code (centered below title) - use TTD image if provided, otherwise QR code
    sig_size = 7 * mm
    sig_y = sig_top - 1 * mm - sig_size
    
    if ttd_kepsek_path and os.path.exists(ttd_kepsek_path):
        # Draw TTD signature image from app
        try:
            c.drawImage(ttd_kepsek_path, sig_center - sig_size / 2, sig_y,
                       width=sig_size, height=sig_size, preserveAspectRatio=True, mask='auto')
        except Exception as e:
            print(f"Error drawing TTD image: {e}")
    else:
        # Fallback to QR code
        qr_data = f"{nama}|{nis_str}|{sekolah}"
        qr_buf = _generate_qr_image(qr_data, box_size=6, border=0)
        if qr_buf:
            try:
                c.drawImage(ImageReader(qr_buf), sig_center - sig_size / 2, sig_y,
                           width=sig_size, height=sig_size, mask='auto')
            except Exception as e:
                print(f"Error drawing QR: {e}")
    
    # Kepsek name (underlined, centered below signature) - bigger font
    name_sig_y = sig_y - 2 * mm
    c.setFillColor(COLOR_DARK)
    c.setFont('TimesNewRoman-Bold', 5)
    c.drawCentredString(sig_center, name_sig_y, kepsek_name)
    name_w = c.stringWidth(kepsek_name, 'TimesNewRoman-Bold', 5)
    c.setStrokeColor(COLOR_DARK)
    c.setLineWidth(0.3)
    c.line(sig_center - name_w / 2, name_sig_y - 0.5 * mm,
           sig_center + name_w / 2, name_sig_y - 0.5 * mm)
    
    # NIP (centered below name) - bigger font
    c.setFont('TimesNewRoman', 4)
    c.drawCentredString(sig_center, name_sig_y - 2 * mm, f"NIP. {kepsek_nip}")


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
