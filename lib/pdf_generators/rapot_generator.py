from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
    BaseDocTemplate, PageTemplate, Frame, NextPageTemplate, Image, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

import os
import re
import io
import numpy as np
import pandas as pd
import cv2

from PIL import Image as PILImage, ImageFilter, ImageEnhance

class SignatureProcessor:
    """
    Class untuk memproses gambar tanda tangan:
    - Enhancement (improve quality)
    - Smart crop dengan margin 0.1 cm
    - Resize max 5cm x 3cm
    - Preserve warna asli
    """
    
    def __init__(self, max_width=5*cm, max_height=3*cm, crop_margin_cm=0.1, 
                 upscale_factor=2, enhance_strength=0.3):
        """
        Args:
            max_width: lebar maksimal (default 5cm)
            max_height: tinggi maksimal (default 3cm)
            crop_margin_cm: margin crop dalam cm (default 0.1cm)
            upscale_factor: faktor upscale untuk quality (default 2x)
            enhance_strength: kekuatan enhancement 0-1 (default 0.3)
        """
        self.max_width = max_width
        self.max_height = max_height
        self.crop_margin_cm = crop_margin_cm
        self.upscale_factor = upscale_factor
        self.enhance_strength = enhance_strength
    
    def enhance_quality(self, img):
        """
        Enhance kualitas gambar dengan preserve warna asli
        """
        # print(f"Enhancing image quality (strength: {self.enhance_strength})...")
        
        # Simpan original untuk blending
        original = img.copy()
        
        # Convert ke RGB untuk processing
        if img.mode not in ['RGB', 'RGBA']:
            img = img.convert('RGB')
        
        # 1. UPSCALE untuk detail lebih baik
        if self.upscale_factor > 1:
            width, height = img.size
            new_size = (width * self.upscale_factor, height * self.upscale_factor)
            img = img.resize(new_size, PILImage.Resampling.LANCZOS)
            # print(f"Upscaled to: {new_size}")
        
        # 2. NOISE REDUCTION (gentle)
        try:
            img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            denoised = cv2.bilateralFilter(img_cv, 5, 30, 30)
            img = PILImage.fromarray(cv2.cvtColor(denoised, cv2.COLOR_BGR2RGB))
        except:
            # Fallback ke PIL
            img = img.filter(ImageFilter.MedianFilter(size=1))
        
        # 3. SHARPENING (conservative)
        img = img.filter(ImageFilter.UnsharpMask(
            radius=1.0,     # blur radius
            percent=120,    # enhancement strength
            threshold=3     # minimum difference threshold
        ))
        
        # 4. CONTRAST ENHANCEMENT (subtle)
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.1)  # 10% increase
        
        # 5. PRESERVE ORIGINAL COLORS
        if self.enhance_strength < 1.0:
            # Resize original to match enhanced
            original_resized = original.resize(img.size, PILImage.Resampling.LANCZOS)
            
            # Make sure same mode
            if original_resized.mode != img.mode:
                original_resized = original_resized.convert(img.mode)
            
            # Blend: preserve original colors
            img = PILImage.blend(original_resized, img, self.enhance_strength)
            # print(f"Blended with original ({int((1-self.enhance_strength)*100)}% original)")
        
        return img
    
    def smart_crop_with_margin(self, img):
        """
        Crop gambar dengan margin 0.1cm dari tepi content
        """
        # print(f"Smart cropping with {self.crop_margin_cm}cm margin...")
        
        # Convert ke grayscale untuk detect content
        if img.mode == 'RGBA':
            # Gunakan alpha channel jika ada
            alpha = np.array(img.split()[-1])
            content_mask = alpha > 50
        else:
            # Convert ke grayscale dan detect non-white areas
            gray = img.convert('L')
            gray_array = np.array(gray)
            content_mask = gray_array < 240  # Non-white content
        
        # Find bounding box of content
        coords = np.argwhere(content_mask)
        
        if len(coords) > 0:
            y0, x0 = coords.min(axis=0)
            y1, x1 = coords.max(axis=0)
            
            # Convert margin dari cm ke pixel
            # Asumsi: gambar sudah di-upscale, jadi DPI lebih tinggi
            dpi_estimate = 150 * self.upscale_factor  # Rough estimate
            margin_pixels = int(self.crop_margin_cm * dpi_estimate / 2.54)  # cm to pixels
            
            # print(f"Content bounds: ({x0},{y0}) to ({x1},{y1})")
            # print(f"Adding margin: {margin_pixels} pixels ({self.crop_margin_cm}cm)")
            
            # Apply margin
            x0 = max(0, x0 - margin_pixels)
            y0 = max(0, y0 - margin_pixels)
            x1 = min(img.width, x1 + margin_pixels)
            y1 = min(img.height, y1 + margin_pixels)
            
            cropped = img.crop((x0, y0, x1, y1))
            # print(f"Cropped to: {cropped.size}")
            
            return cropped
        else:
            print("Warning: No content detected for cropping")
            return img
    
    def process_signature(self, image_path):
        """
        Process signature lengkap: enhance -> crop -> resize
        """
        try:
            print(f"=== Processing signature: {image_path} ===")
            
            with PILImage.open(image_path) as img:
                img = img.copy()
                # print(f"Original: {img.size}, mode: {img.mode}")
                
                # 1. ENHANCE QUALITY
                img_enhanced = self.enhance_quality(img)
                
                # 2. SMART CROP dengan margin
                img_cropped = self.smart_crop_with_margin(img_enhanced)
                
                # 3. SAVE ke buffer dengan high DPI
                img_buffer = io.BytesIO()
                img_cropped.save(img_buffer, format='PNG', dpi=(300, 300))
                img_buffer.seek(0)
                
                # print(f"Final processing completed")
                # print(f"Max size: {self.max_width/cm:.1f}cm x {self.max_height/cm:.1f}cm")
                # print("="*50)
                
                # 4. RETURN ReportLab Image dengan size constraint
                return Image(img_buffer, width=self.max_width, height=self.max_height, kind='bound')
                
        except Exception as e:
            print(f"Error processing signature {image_path}: {e}")
            raise
    
    def process_signature_to_bytes(self, image_path):
        """
        Process signature and return raw PNG bytes (for caching).
        Much faster than process_signature() for repeated use.
        """
        try:
            with PILImage.open(image_path) as img:
                img = img.copy()
                img_enhanced = self.enhance_quality(img)
                img_cropped = self.smart_crop_with_margin(img_enhanced)
                img_buffer = io.BytesIO()
                img_cropped.save(img_buffer, format='PNG', dpi=(300, 300))
                return img_buffer.getvalue()  # Return raw bytes
        except Exception as e:
            print(f"Error processing signature {image_path}: {e}")
            return None

    def set_max_size(self, width_cm, height_cm):
        """Update ukuran maksimal"""
        self.max_width = width_cm * cm
        self.max_height = height_cm * cm
        print(f"Max size updated to: {width_cm}cm x {height_cm}cm")
    
    def set_crop_margin(self, margin_cm):
        """Update margin crop"""
        self.crop_margin_cm = margin_cm
        print(f"Crop margin updated to: {margin_cm}cm")
    
    def set_enhancement(self, upscale_factor=2, enhance_strength=0.3):
        """Update enhancement settings"""
        self.upscale_factor = upscale_factor
        self.enhance_strength = enhance_strength
        print(f"Enhancement updated: upscale={upscale_factor}x, strength={enhance_strength}")
    
    def preview_steps(self, image_path, save_steps=True):
        """
        Preview hasil processing step by step
        """
        try:
            print(f"Creating preview for: {image_path}")
            
            with PILImage.open(image_path) as img:
                img = img.copy()
                
                if save_steps:
                    img.save("preview_1_original.png")
                    print("Saved: preview_1_original.png")
                
                # Step 1: Enhancement
                img_enhanced = self.enhance_quality(img)
                if save_steps:
                    # Resize untuk preview (jangan terlalu besar)
                    preview_size = (min(800, img_enhanced.width), min(600, img_enhanced.height))
                    img_enhanced.resize(preview_size, PILImage.Resampling.LANCZOS).save("preview_2_enhanced.png")
                    print("Saved: preview_2_enhanced.png")
                
                # Step 2: Crop
                img_cropped = self.smart_crop_with_margin(img_enhanced)
                if save_steps:
                    preview_size = (min(800, img_cropped.width), min(600, img_cropped.height))
                    img_cropped.resize(preview_size, PILImage.Resampling.LANCZOS).save("preview_3_cropped.png")
                    print("Saved: preview_3_cropped.png")
                
                print("Preview completed! Check the preview_*.png files")
                return img_cropped
                
        except Exception as e:
            print(f"Error creating preview: {e}")
            return None
    
    def process_multiple_signatures(self, image_paths):
        """Process multiple signatures"""
        results = []
        for i, path in enumerate(image_paths, 1):
            try:
                print(f"\nProcessing {i}/{len(image_paths)}: {path}")
                result = self.process_signature(path)
                results.append(result)
            except Exception as e:
                print(f"Failed to process {path}: {e}")
                continue
        
        print(f"\nCompleted: {len(results)}/{len(image_paths)} signatures processed")
        return results

class PDFGenerator:
    def __init__(self, excel_nilai=None, excel_data=None, ttd_kepsek_path=None, ttd_wali_path=None, ttd_wali_dict=None, logo_path=None, output_folder=None, progress_callback=None):
        """
        Initialize PDF Generator
        
        Args:
            excel_nilai: Path to nilai Excel file
            excel_data: Path to data siswa Excel file
            ttd_kepsek_path: Path to kepsek signature image (optional)
            ttd_wali_path: Path to wali kelas signature image (optional, legacy single wali)
            ttd_wali_dict: Dict of {kelas: path} for per-class wali signatures (optional)
            logo_path: Path to logo image (optional)
            output_folder: Output folder for generated PDFs (optional)
            progress_callback: Optional callback(current, total, kelas, student_name)
        """
        self.progress_callback = progress_callback
        self.debug = 0
        self.warnings = []  # Collect warnings during generation
        self.errors = []    # Collect errors during generation

        # Project root for default asset paths
        self.project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        
        # Store dynamic paths
        self.logo_path = logo_path or os.path.join(self.project_root, "assets", "logo.png")
        self.ttd_kepsek_path = ttd_kepsek_path or os.path.join(self.project_root, "assets", "signatures", "ttd_kepsek.png")
        self.ttd_wali_path = ttd_wali_path or os.path.join(self.project_root, "assets", "signatures", "ttd_wali.png")
        self.ttd_wali_dict = ttd_wali_dict or {}  # Dict of {kelas: path} for per-class wali signatures
        self.ttd_dummy_path = os.path.join(self.project_root, "assets", "signatures", "dummy.png")
        self.output_folder = output_folder or "."
        
        # self.filename = filename
        self.width, self.height = A4
        self.judul_dan_isi = 1
        self.fontsize = 12
        
        # Initialize default attributes
        self._init_default_attributes()
        
        # foto buat ttd
        self.sig_processor = SignatureProcessor(
            max_width=5*cm,         # Max 5cm lebar
            max_height=2.5*cm,      # Max 2.5cm tinggi  
            crop_margin_cm=0.2,    # Margin crop 0.2cm
            upscale_factor=1,       # No upscale (faster, sufficient for PDF)
            enhance_strength=0.2    # Enhancement subtle (20%)
        )
        
        # Cache logo as bytes (read once, reuse per student)
        self._logo_bytes = None
        if os.path.exists(self.logo_path):
            with open(self.logo_path, 'rb') as f:
                self._logo_bytes = f.read()
        
        # Process signature images - store as raw PNG bytes for fast reuse
        self._ttd_kepsek_bytes = None
        self._ttd_dummy_bytes = None
        self._ttd_wali_bytes = {}  # Dict of {kelas: png_bytes}
        
        if os.path.exists(self.ttd_kepsek_path):
            self._ttd_kepsek_bytes = self.sig_processor.process_signature_to_bytes(self.ttd_kepsek_path)
        if os.path.exists(self.ttd_dummy_path):
            self._ttd_dummy_bytes = self.sig_processor.process_signature_to_bytes(self.ttd_dummy_path)
        
        # Process per-class wali signatures
        for kelas, path in self.ttd_wali_dict.items():
            if path and os.path.exists(path):
                try:
                    self._ttd_wali_bytes[kelas.upper()] = self.sig_processor.process_signature_to_bytes(path)
                except Exception as e:
                    print(f"Error loading wali signature for {kelas}: {e}")

        # Try to load data from Excel if provided
        if excel_nilai and excel_data:
            success = self._load_excel(excel_nilai, excel_data)
            if success:
                pass  # loaded successfully
                self._init_excel_data()
                self._init_default_school_identity()  # jadi col col
                # Setup components
                self._setup_fonts()
                self._init_styles()
                self._process_all_students()
            else:
                print("Failed to load Excel, using default data")
                self._setup_fonts()
                self._init_styles()
        else:
            print("No Excel path provided, using default data")
            self._setup_fonts()
            self._init_styles()
    
    def _init_default_attributes(self):
        """Initialize all default attributes before loading Excel"""
        self.nilai_dict = {}
        self.data_dict = {}
        self.data_tetap = None
        self.guru = None
        self.tanggal = ""
        self.tahun = ""
        self.semester = ""
        self.namaSekolah = "PKBM PPI Taiwan"
        self.alamatSekolah = ""
        self.kodePos = ""
        self.website = ""
        self.email = ""
        self.telepon = ""
        self.nomorIndukLembaga = ""
        self.NAMAKELAS = ""
        self._init_default_grade_data()
        self._init_default_attitude_data()
        self._init_default_extracurricular_data()
        self._init_default_attendance_data()
    
    def _process_all_students(self):
        """Process all students from loaded Excel data"""
        # Count total students for progress
        total_students = sum(len(df) for df in self.nilai_dict.values())
        current_count = 0
        skipped_students = []
        
        for key in self.nilai_dict:
            try:
                # Use .copy() to avoid SettingWithCopyWarning and slow copy-on-write
                df_nilai = self.nilai_dict[key].copy()
                kolom_numerik = df_nilai.select_dtypes(include='number').columns[
                    df_nilai.select_dtypes(include='number').notna().any()
                ]
                df_nilai[kolom_numerik] = df_nilai[kolom_numerik].fillna(0).round().astype(int)
                self.nilai_dict[key] = df_nilai
                
                self.NAMAKELAS = key
                
                df_data = self.data_dict[key].copy()
                df_data = df_data.fillna("")

                # Parse date columns (stored as dd/mm/yyyy text) to Indonesian string
                # Must happen BEFORE astype(str) so pd.to_datetime gets the raw value
                date_cols = ['Tanggal Lahir'] + [
                    c for c in df_data.columns if 'pada tanggal' in str(c).lower()
                ]
                for col in date_cols:
                    if col in df_data.columns:
                        df_data[col] = (
                            pd.to_datetime(df_data[col], dayfirst=True, errors='coerce')
                            .apply(self.format_tanggal_indonesia)
                        )

                df_data = df_data.astype(str)
                df_data['Tempat'] = df_data['Tempat'].str.title() + ", " + df_data["Tanggal Lahir"]
                df_data.rename(columns={"Tempat": "Tempat / Tanggal Lahir"}, inplace=True)
                df_data.drop(columns=['Tanggal Lahir'], errors='ignore', inplace=True)
                self.data_dict[key] = df_data

                # Process each student with progress callback
                for idx, row in self.nilai_dict[key].iterrows():
                    nama = row.Nama
                    # Read NIS from nilai row if available (for disambiguation)
                    nis_nilai = None
                    if 'NIS' in row.index:
                        raw = str(row['NIS']).strip()
                        if raw not in ('', 'nan', 'None'):
                            nis_nilai = raw
                    current_count += 1
                    
                    # Find student in data siswa (case-insensitive name, NIS for disambiguation)
                    matched_data = self._find_student_in_data(key, nama, nis_nilai)
                    if matched_data is None:
                        msg = f"[{key}] Siswa '{nama}' ada di file Nilai tapi TIDAK DITEMUKAN di Data Siswa - DILEWATI"
                        self.warnings.append(msg)
                        print(f"WARNING: {msg}")
                        skipped_students.append(f"{key}: {nama}")
                        if self.progress_callback:
                            self.progress_callback(current_count, total_students, key, f"{nama} (DILEWATI)")
                        continue
                    
                    try:
                        self.transfer_data_to_generate(row, matched_data)
                        if self.progress_callback:
                            self.progress_callback(current_count, total_students, key, nama)
                    except Exception as e:
                        msg = f"[{key}] Error generate rapot untuk '{nama}': {str(e)}"
                        self.errors.append(msg)
                        print(f"ERROR: {msg}")
                        skipped_students.append(f"{key}: {nama}")
                        if self.progress_callback:
                            self.progress_callback(current_count, total_students, key, f"{nama} (ERROR)")
                        
            except Exception as e:
                msg = f"[{key}] Error memproses kelas: {str(e)}"
                self.errors.append(msg)
                print(f"ERROR: {msg}")
                import traceback
                traceback.print_exc()
        
        if skipped_students:
            msg = f"{len(skipped_students)} siswa dilewati karena error"
            self.warnings.append(msg)
            print(f"WARNING: {msg}: {', '.join(skipped_students)}")

    def _find_student_in_data(self, key, nama, nis_nilai=None):
        """Find student row in data_dict using case-insensitive name match.
        If duplicate names exist, uses NIS from nilai to disambiguate."""
        df = self.data_dict[key]
        nama_upper = str(nama).strip().upper()

        # Case-insensitive match on index (Nama Peserta Didik)
        mask = df.index.str.strip().str.upper() == nama_upper
        matches = df[mask]

        if len(matches) == 0:
            return None
        if len(matches) == 1:
            return matches.iloc[0]

        # Duplicate names — use NIS to pick the right row
        if nis_nilai is not None and 'NIS' in df.columns:
            nis_upper = str(nis_nilai).strip().upper()
            nis_mask = matches['NIS'].astype(str).str.strip().str.upper() == nis_upper
            nis_matches = matches[nis_mask]
            if len(nis_matches) > 0:
                return nis_matches.iloc[0]

        # Fallback: return first match and warn
        msg = f"Nama '{nama}' duplikat tanpa NIS yang cocok, menggunakan data pertama"
        self.warnings.append(msg)
        print(f"WARNING: {msg}")
        return matches.iloc[0]

    def cek_folder(self, folder_path, nama):
        # Use output_folder if set, otherwise use folder_path
        base_folder = self.output_folder if self.output_folder else "."
        full_folder = os.path.join(base_folder, folder_path)
        
        # Cek dan buat folder jika belum ada
        if not os.path.exists(full_folder):
            os.makedirs(full_folder)
        
        nama_file = nama + ".pdf"
        # Gabungkan folder dan nama file
        self.filename = os.path.join(full_folder, nama_file)
        
    def _make_sig_image(self, png_bytes):
        """Create a fresh ReportLab Image from cached PNG bytes"""
        if png_bytes is None:
            return None
        buf = io.BytesIO(png_bytes)
        return Image(buf, width=self.sig_processor.max_width, height=self.sig_processor.max_height, kind='bound')

    def _make_logo_image(self, width, height):
        """Create a fresh ReportLab Image from cached logo bytes"""
        if self._logo_bytes is None:
            return None
        buf = io.BytesIO(self._logo_bytes)
        return Image(buf, width=width, height=height)

    def transfer_data_to_generate(self, nilai,data):
        self._process_grade_data(nilai)
        self.dataSiswa = data
        self.nisn = self.dataSiswa.NISN
        self.namaSiswa = nilai.Nama.title()
        self.nis = self.dataSiswa.NIS
        self.namaSiswa = nilai.Nama
        self.kelas = nilai.Kelas

        self._init_computed_data() # header dan footer
        self.cek_folder( f"{self.NAMAKELAS} {self.kelas} {self.semester}", f"{self.nis} - {self.namaSiswa}" )
        self._setup_document()

        self.generate()

    def _get_data_tetap(self, key, default=""):
        """Safely get a value from data_tetap with fallback"""
        try:
            if key in self.data_tetap.index:
                return self.data_tetap.loc[key].NAMA
            else:
                msg = f"Variabel '{key}' tidak ditemukan di sheet DataTetap (tersedia: {', '.join(self.data_tetap.index.tolist())})"
                self.warnings.append(msg)
                print(f"WARNING: {msg}")
                return default
        except Exception as e:
            self.warnings.append(f"Error membaca '{key}' dari DataTetap: {e}")
            return default

    def _init_excel_data(self):
        """
        Initialize ALL required data attributes with default values.
        This ensures every attribute exists even if Excel loading fails.
        """
        pass  # init default data
        # ================================
        # 1. BASIC INFO (Tanggal, Tahun, dll)
        # ================================
        self.tanggal = self._get_data_tetap('TANGGAL', 'Taipei, 1 Januari 2025')
        self.tahun = self._get_data_tetap('TAHUNPELAJARAN', '2024/2025')
        self.semester = self._get_data_tetap('SEMESTER', 'GANJIL')
        
        # ================================
        # 2. SCHOOL DATA (Info Sekolah)
        # ================================
        self.namaSekolah = self._get_data_tetap('NAMASEKOLAH', 'PKBM PPI Taiwan')
        self.alamatSekolah = self._get_data_tetap('ALAMAT', '')
        self.kodePos = str(self._get_data_tetap('KODEPOS', ''))
        self.website = self._get_data_tetap('WEBSITE', '')
        self.email = self._get_data_tetap('EMAIL', '')
        self.telepon = str(self._get_data_tetap('TELEPON', ''))
        self.nomorIndukLembaga = str(self._get_data_tetap('NPSN', ''))
        
        
        # ================================
        # 5. GRADE DATA (Data Nilai)
        # ================================
        self._init_default_grade_data()
        
        # ================================
        # 6. ATTITUDE DATA (Data Sikap)
        # ================================
        self._init_default_attitude_data()
        
        # ================================
        # 7. EXTRACURRICULAR DATA (Data Ekstrakurikuler)
        # ================================
        self._init_default_extracurricular_data()
        
        # ================================
        # 8. ATTENDANCE DATA (Data Absensi)  
        # ================================
        self._init_default_attendance_data()
        
        # ================================
        # 9. SCHOOL IDENTITY TABLE
        # ================================
        # self._init_default_school_identity()
        
    def _init_default_grade_data(self):
        """Initialize default grade data"""
        self.nilai_data = [
            ["No", "Mata Pelajaran", "Nilai Akhir", "Capaian Kompetensi"],
            
            ["Kelompok Mata Pelajaran Umum Merge", "", "", ""],
            ["1.", "Pendidikan Agama", 85,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan sangat baik. Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan sangat baik."],
            ["2.", "Pendidikan Kewarganegaraan", 85,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan sangat baik. Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan sangat baik."],
            ["3.", "Bahasa Indonesia", 78,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            ["4.", "Bahasa Inggris", 77,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            ["5.", "Matematika (Umum)", 81,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            ["6.", "Sejarah Indonesia", 75,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            
            ["Peminatan Merge", "", "", ""],
            ["1.", "Geografi", 81,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            ["2.", "Sejarah", 90,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan sangat baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan sangat baik."],
            ["3.", "Sosiologi", 80,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            ["4.", "Ekonomi", 90,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan sangat baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan sangat baik."],
            
            ["Kelompok Mata Pelajaran Khusus Merge", "", "", ""],
            ["1.", "Pemberdayaan", 76,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            ["2.", "Manajemen", 80,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            ["3.", "Teknologi Informasi dan Komunikasi", 89,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan sangat baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan sangat baik."],
            ["4.", "Pendidikan Karya dan Kewirausahaan", 83,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            ["5.", "Seni Budaya", 76,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik. "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."],
            ["6.", "Pendidikan Jasmani dan Olahraga", 76,
             "Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan baik "
             "Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan baik."]
        ]
    
    def _init_default_attitude_data(self):
        """Initialize default attitude data"""
        self.sikap_data = [
            ["No", "Dimensi", "Keterangan"],
            ["1", "Beriman, Bertaqwa kepada Tuhan Yang Maha Esa, dan Berakhlak Mulia",
             "Peserta didik sudah baik dalam penerapan keimanan, ketaqwaan kepada Tuhan Yang Maha Esa, serta berakhlak mulia"],
            ["2", "Berkebhinekaan Global",
             "Peserta didik memiliki pemahaman dan sikap yang baik dalam menyikapi kebhinekaan global."],
            ["3", "Bergotong Royong",
             "Peserta didik memiliki jiwa gotong royong yang baik."],
            ["4", "Mandiri",
             "Kemandirian peserta didik dalam proses belajar mengajar sudah sangat baik."],
            ["5", "Bernalar Kritis",
             "Dalam proses pembelajaran, peserta didik sudah memiliki nalar kritis yang baik."],
            ["6", "Kreatif",
             "Kreatifitas peserta didik sudah baik selama proses pembelajaran berlangsung."],
        ]
    
    def _init_default_extracurricular_data(self):
        """Initialize default extracurricular data"""
        self.ekskul_data = [
            ["No", "Kegiatan Ekstrakulikuler", "Predikat", "Keterangan"],
            ["1.", "", "", ""],
            ["2.", "", "", ""]
        ]
    
    def _init_default_attendance_data(self):
        """Initialize default attendance data"""
        self.absensi_data = [
            ["Ketidakhadiran", ""],
            ["Sakit", ": Hari"],
            ["Izin", ": Hari"],
            ["Tanpa Keterangan", ": Hari"]
        ]
    
    def _init_default_school_identity(self):
        """Initialize default school identity table"""
        school_data = [
            ['Nama Sekolah', self.namaSekolah],
            ['Nomor Induk Lembaga', self.nomorIndukLembaga],
            ['Alamat', self.alamatSekolah],
            ['Kode Pos', self.kodePos],
            ['Website', self.website],
            ['Email', self.email],
            ['Telepon', self.telepon]
        ]
        self.IdentitasSekolah = pd.DataFrame(school_data, columns=['Judul', 'Value'])
    
    def _init_computed_data(self):
        """Initialize computed data that depends on other attributes"""
        # Header table data
        self.header_data = [
            ["Nama Sekolah", ":", self.namaSekolah, "", "Kelas", ":", self.kelas],
            ["Alamat Sekolah", ":", self.alamatSekolah, "", "Semester", ":", self.semester],
            ["Nama Peserta Didik", ":", self.namaSiswa , "", "Tahun Pelajaran", ":", self.tahun],
            ["NISN / NIS", ":", f"{self.nisn} / {self.nis}", "", "", "", ""]
        ]
        
        # Footer data
        self.footer_data = [
            f"<u><i>{self.namaSiswa}</i></u> | {self.nis}", 
            "PAKET C IPS"
        ]
        
        # Create fresh Image objects from cached bytes (not reused stale objects)
        wali_bytes = self._ttd_wali_bytes.get(self.NAMAKELAS, self._ttd_dummy_bytes)
        ttd_wali_img = self._make_sig_image(wali_bytes)
        ttd_kepsek_img = self._make_sig_image(self._ttd_kepsek_bytes)
        
        # Signature table data
        self.ttd_data = [
            ["Mengetahui", "", self.tanggal],
            ["Orang Tua / Wali", "", "Wali Kelas,"],
            ["<br/><br/><br/><br/><br/>", "", ttd_wali_img if ttd_wali_img else ""],
            ["(........................................)", "", self.bold_underline(self.guru.loc[self.NAMAKELAS].NAMA)],
            ["", "", f"NIP."],
            ["<br/>Mengetahui", "", ""],
            ["", "", ""],
            ["Kepala Sekolah,", "", ""],
            [ttd_kepsek_img, "", ""],
            [self.bold_underline(self.guru.loc['KEPSEK'].NAMA), "", ""],
            [f"NIP.", "", ""],
        ]
        # {self.nipKepsek if self.nipKepsek else ''}

    def _load_excel(self, excel_nilai, excel_data):
        try:
            self.nilaidf = pd.read_excel(excel_nilai ,sheet_name=None)
            self.datadf = pd.read_excel(excel_data ,sheet_name=None, dtype=str)

            self.nilaidf = {"".join(k.upper().strip().split()): v for k, v in self.nilaidf.items()}
            self.datadf = {"".join(k.upper().strip().split()): v for k, v in self.datadf.items()}
            
            # data_tetap 
            self.nilaidf['DATATETAP'].columns = self.nilaidf['DATATETAP'].columns.str.upper().str.strip()
            self.data_tetap = self.nilaidf['DATATETAP'].set_index(['VARIABEL'])
            self.data_tetap.index = self.data_tetap.index.str.upper().str.replace(r"[\s_]+", "", regex=True)

            # guru nama nip
            self.nilaidf['GURU'].columns = self.nilaidf['GURU'].columns.str.upper()
            self.guru = self.nilaidf['GURU'].set_index(['KELAS'])
            self.guru.index = self.guru.index.str.strip().str.upper()

            self.nilai_dict = {k: v for k, v in self.nilaidf.items() if "KP" in k}
            self.data_dict = {k: v for k, v in self.datadf.items() if "KP" in k}

            # Validate: check which KP sheets in nilai have matching data siswa sheets
            nilai_keys = set(self.nilai_dict.keys())
            data_keys = set(self.data_dict.keys())
            missing_data_sheets = nilai_keys - data_keys
            if missing_data_sheets:
                msg = f"Sheet kelas berikut ada di file Nilai tapi TIDAK ADA di file Data Siswa: {', '.join(sorted(missing_data_sheets))}"
                self.warnings.append(msg)
                print(f"WARNING: {msg}")
                # Remove sheets that don't have matching data
                for k in missing_data_sheets:
                    del self.nilai_dict[k]

            for i in self.data_dict:
                # Drop rows where 'Nama Peserta Didik' is NaN
                self.data_dict[i] = self.data_dict[i].dropna(subset=["Nama Peserta Didik"])
                self.data_dict[i]["Nama Peserta Didik"] = self.data_dict[i]["Nama Peserta Didik"].str.strip().str.title()
                self.data_dict[i] = self.data_dict[i].set_index("Nama Peserta Didik", drop=False)

            for i in self.nilai_dict:
                # Drop rows where 'Nama' is NaN, use .copy() to avoid SettingWithCopyWarning
                df = self.nilai_dict[i].dropna(subset=["Nama"]).copy()
                df["Nama"] = df["Nama"].str.strip().str.title()
                if 'NIS' in df.columns:
                    df['NIS'] = df['NIS'].astype(str).str.strip()
                self.nilai_dict[i] = df

            # Validate: case-insensitive name check between nilai and data siswa per sheet
            for key in list(self.nilai_dict.keys()):
                if key in self.data_dict:
                    nilai_names_upper = set(self.nilai_dict[key]["Nama"].str.upper().tolist())
                    data_names_upper = set(self.data_dict[key].index.str.upper().tolist())
                    missing_names = nilai_names_upper - data_names_upper
                    if missing_names:
                        msg = f"[{key}] Siswa berikut ada di Nilai tapi TIDAK ADA di Data Siswa: {', '.join(sorted(missing_names))}"
                        self.warnings.append(msg)
                        print(f"WARNING: {msg}")
                
            return True 
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            import traceback
            traceback.print_exc()
            self.errors.append(f"Gagal membaca file Excel: {e}")
            return False

    def _process_grade_data(self, series_of_rows):
        """Process grade data from Series of rows (dict-like)"""
        try:
            # Convert grade data to the format needed
            self.nilai_data = [["No", "Mata Pelajaran", "Nilai Akhir", "Capaian Kompetensi"]]
            no = 1
            current_group = None
            # Exclude non-grade identifier columns regardless of their position
            _skip = {'nama', 'kelas', 'nis', 'nisn'}
            grade_items = series_of_rows[
                [c for c in series_of_rows.index if str(c).strip().lower() not in _skip]
            ]
            for idx, val in grade_items.items():
                if "merge" in str(idx).lower() or pd.isna(val):
                    # This is a group header
                    no = 1
                    self.nilai_data.append([f"{idx}", "", "", ""])
                else :
                    # This is a subject
                    self.nilai_data.append([no, idx, val, self.ket_nilai(int(val))])
                    no+=1

        except Exception as e:
            print(f"Error processing grade data: {e}")

    def _process_attitude_data(self, df):
        """Process attitude data from DataFrame"""
        try:
            self.sikap_data = [["No", "Dimensi", "Keterangan"]]
            
            for _, row in df.iterrows():
                no = row.get('No', '')
                dimensi = row.get('Dimensi', '')
                keterangan = row.get('Keterangan', '')
                self.sikap_data.append([no, dimensi, keterangan])
            
            pass  # loaded attitude entries
            
        except Exception as e:
            print(f"Error processing attitude data: {e}")
    
    def _process_extracurricular_data(self, df):
        """Process extracurricular data from DataFrame"""
        try:
            self.ekskul_data = [["No", "Kegiatan Ekstrakulikuler", "Predikat", "Keterangan"]]
            
            for _, row in df.iterrows():
                no = row.get('No', '')
                kegiatan = row.get('Kegiatan', '')
                predikat = row.get('Predikat', '')
                keterangan = row.get('Keterangan', '')
                self.ekskul_data.append([no, kegiatan, predikat, keterangan])
            
            pass  # loaded extracurricular entries
            
        except Exception as e:
            print(f"Error processing extracurricular data: {e}")
    
    def _process_attendance_data(self, df):
        """Process attendance data from DataFrame"""
        try:
            self.absensi_data = [["Ketidakhadiran", ""]]
            
            for _, row in df.iterrows():
                jenis = row.get('Jenis', '')
                jumlah = row.get('Jumlah', '')
                self.absensi_data.append([jenis, f": {jumlah} Hari"])
            
            pass  # loaded attendance entries
            
        except Exception as e:
            print(f"Error processing attendance data: {e}")
    
    def _setup_fonts(self):
        """Setup font configuration"""
        self.fontpath = os.path.join(self.project_root, "assets", "fonts")
        # print(self.fontpath)
        try:
            pdfmetrics.registerFont(TTFont('TimesRoman', os.path.join(self.fontpath, 'times.ttf')))
            pdfmetrics.registerFont(TTFont('TimesRomanBold', os.path.join(self.fontpath, 'timesbd.ttf')))
            pdfmetrics.registerFont(TTFont('TimesRomanBoldItalic', os.path.join(self.fontpath, 'timesbdit.ttf')))
            pdfmetrics.registerFont(TTFont('TimesRomanItalic', os.path.join(self.fontpath, 'timesit.ttf')))
            
            self.times = 'TimesRoman'
            self.timesBold = 'TimesRomanBold'
            self.timesBoldItalic = 'TimesRomanBoldItalic'
            self.timesItalic = 'TimesRomanItalic'
        except:
            print("Warning: Font tidak ditemukan, menggunakan font default")
            self.times = 'Times-Roman'
            self.timesBold = 'Times-Roman'
            self.timesBoldItalic = 'Times-Roman'
            self.timesItalic = 'Times-Roman'
    
    def format_tanggal_indonesia(self, tanggal):
        """
        Mengubah timestamp atau datetime menjadi format 'D NamaBulan YYYY' dalam Bahasa Indonesia.
        Contoh: Timestamp('1988-12-24') → '24 Desember 1988', '1988-01-05' → '5 Januari 1988'
        Mendukung berbagai format input: dd/mm/yyyy, dd/mm/yyyy HH:MM:SS, Timestamp, dll.
        """
        if pd.isna(tanggal):
            return ''

        # Flexible parsing: handles timestamps, dd/mm/yyyy, dd/mm/yyyy HH:MM:SS, etc.
        parsed = pd.to_datetime(tanggal, errors='coerce', dayfirst=True)
        if pd.isna(parsed):
            return ''

        bulan_indo = {
            'January': 'Januari', 'February': 'Februari', 'March': 'Maret',
            'April': 'April', 'May': 'Mei', 'June': 'Juni',
            'July': 'Juli', 'August': 'Agustus', 'September': 'September',
            'October': 'Oktober', 'November': 'November', 'December': 'Desember'
        }

        # Day without leading zero (cross-platform)
        hasil = f"{parsed.day} {parsed.strftime('%B')} {parsed.year}"
        for eng, indo in bulan_indo.items():
            hasil = hasil.replace(eng, indo)
        return hasil
    
    def _init_styles(self):
        """Initialize all paragraph styles"""
        self.styles = getSampleStyleSheet()
        
        # Footer styles
        self.footerLeft = ParagraphStyle(
            name="footerLeft",
            fontName=self.timesBoldItalic,
            fontSize=10,
            alignment=TA_LEFT
        )
        
        self.footerRight = ParagraphStyle(
            name="footerRight",
            fontName=self.timesBold,
            fontSize=10,
            alignment=TA_RIGHT
        )
        
        # Normal style
        self.normal_style = ParagraphStyle(
            'CustomNormal',
            parent=self.styles['Normal'],
            fontName=self.times,
            fontSize=14,
            leading=12
        )
        
        # Header style
        self.header_style = ParagraphStyle(
            'CustomHeader',
            parent=self.styles['Heading1'],
            fontName=self.times,
            fontSize=12,
            leading=14,
            alignment=TA_CENTER
        )
        
        # Footer style
        self.footer_style = ParagraphStyle(
            'CustomFooter',
            parent=self.styles['Normal'],
            fontName=self.times,
            fontSize=9,
            leading=11,
            alignment=TA_CENTER
        )
        
        # Cover styles
        self.coverBold = ParagraphStyle(
            'coverBold',
            parent=self.styles['Normal'],
            fontName=self.timesBold,
            fontSize=30,
            leading=40,
            alignment=TA_CENTER
        )
        
        self.baris2normal = ParagraphStyle(
            'baris2normal',
            parent=self.styles['Normal'],
            fontName=self.timesBold,
            fontSize=18,
            leading=20,
            alignment=TA_CENTER
        )
        
        self.namanis = ParagraphStyle(
            'namanis',
            fontName=self.timesBold,
            parent=self.styles['Normal'],
            fontSize=18,
            leading=25,
            alignment=TA_CENTER
        )
        
        self.cover_name_nis = ParagraphStyle(
            'cover_name_nis',
            parent=self.styles['Normal'],
            fontName=self.times,
            fontSize=16,
            leading=23,
            alignment=TA_CENTER
        )
        
        # Signature styles
        self.center_ttd = ParagraphStyle(
            'center_ttd',
            parent=self.styles['Normal'],
            fontName=self.times,
            fontSize=12,
            alignment=TA_CENTER,
            leading=15
        )
        
        self.center_ttd_bold = ParagraphStyle(
            'center_ttd_bold',
            parent=self.styles['Normal'],
            fontName=self.timesBold,
            fontSize=12,
            alignment=TA_CENTER,
            leading=15
        )
        
        # Table styles
        self.tableValLeft = ParagraphStyle(
            name="tableValLeft",
            fontName=self.times,
            fontSize=12,
            leading=18,
            alignment=TA_LEFT
        )
        
        self.tableValCenter = ParagraphStyle(
            name="tableValCenter",
            fontName=self.times,
            fontSize=12,
            leading=18,
            alignment=TA_CENTER
        )
        
        self.tableValJustify = ParagraphStyle(
            name="tableValJustify",
            fontName=self.times,
            fontSize=12,
            leading=14,
            alignment=TA_JUSTIFY
        )
        
        self.tableIdxLeft = ParagraphStyle(
            name="tableIdxLeft",
            fontName=self.timesBold,
            fontSize=12,
            leading=18,
            alignment=TA_LEFT
        )
        
        self.tableIdx = ParagraphStyle(
            name="tableIdx",
            fontName=self.timesBold,
            fontSize=12,
            leading=18,
            alignment=TA_CENTER
        )
        
        self.tableNo = ParagraphStyle(
            name="tableNo",
            fontName=self.times,
            fontSize=12,
            leading=18,
            alignment=TA_CENTER
        )
        
        self.para_ttd_style = ParagraphStyle(
            name="para_ttd_style",
            fontName=self.timesBold,
            fontSize=12,
            leading=18,
            alignment=TA_RIGHT
        )
        
        # Page title styles
        self.judulPage = ParagraphStyle(
            'judulPage',
            fontName=self.timesBold,
            parent=self.styles['Normal'],
            fontSize=18,
            leading=23,
            alignment=TA_CENTER
        )
        
        self.pageTitleLeft = ParagraphStyle(
            "pageTitleLeft",
            fontName=self.timesBold,
            spaceBefore=0,
            spaceAfter=0,
            fontSize=16,
            leading=22,
            alignment=TA_LEFT
        )
        
        # Identity table styles
        self.identitasSekolahColValue = ParagraphStyle(
            name="identitasSekolahColValue",
            fontName=self.times,
            fontSize=14,
            leading=16,
            alignment=TA_LEFT,
            allowHTML=True
        )
        
        self.identitasSekolahColName = ParagraphStyle(
            name="identitasSekolahColName",
            fontName=self.times,
            fontSize=14,
            leading=16
        )
        
        self.identitasSekolahColNo = ParagraphStyle(
            name="identitasSekolahColNo",
            fontName=self.times,
            fontSize=14,
            leading=16,
            alignment=TA_RIGHT
        )
    
    def _setup_document(self):
        """Setup document with multiple page templates"""
        self.doc = BaseDocTemplate(self.filename, pagesize=A4)
        
        # Template for pages WITHOUT header/footer (pages 1-2)
        frame_no_header = Frame(
            1*cm, 2*cm,
            self.width - 2*cm, self.height - 4*cm,
            leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0
        )
        
        # Template for pages WITH header/footer (pages 3+)
        frame_with_header = Frame(
            2*cm, 2*cm,
            self.width - 4*cm, self.height - 6.2*cm,
            leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0
        )
        
        # Create page templates
        template_no_header = PageTemplate(
            id='no_header',
            frames=[frame_no_header],
            onPage=self._add_page_elements_no_header
        )
        
        template_with_header = PageTemplate(
            id='with_header',
            frames=[frame_with_header],
            onPage=self._add_page_elements_with_header
        )
        
        # Add templates to document
        self.doc.addPageTemplates([template_no_header, template_with_header])
    
    # Helper methods
    def add_spacer(self, height_cm=0.5):
        """Add spacer with specific height in cm"""
        self.story.append(Spacer(1, height_cm * cm))
    
    def bold_underline(self, text):
        """Return text with bold and underline HTML tags"""
        return f"<b><u>{text}</u></b>"
    
    def print_all_attributes(self):
        """
        Utility method to print all initialized attributes for debugging
        """
        print("\n" + "="*50)
        print("ALL INITIALIZED ATTRIBUTES:")
        print("="*50)
        
        print("\n1. BASIC INFO:")
        print(f"  tanggal: {self.tanggal}")
        print(f"  tahun: {self.tahun}")
        print(f"  semester: {self.semester}")
        # print(f"  kelas: {self.kelas}")
        
        print("\n2. SCHOOL DATA:")
        print(f"  namaSekolah: {self.namaSekolah}")
        print(f"  alamatSekolah: {self.alamatSekolah}")
        print(f"  kodePos: {self.kodePos}")
        print(f"  website: {self.website}")
        print(f"  email: {self.email}")
        print(f"  telepon: {self.telepon}")
        print(f"  nomorIndukLembaga: {self.nomorIndukLembaga}")
        
        print("\n3. PERSONNEL DATA:")
        print(f"  kepsek: {self.guru.loc['KEPSEK'].NAMA}")
        print(f"  nipKepsek: ")
        print(f"  waliKelas: {self.guru.loc[self.NAMAKELAS].NAMA}")
        print(f"  nipWaliKelas: ")
        
        print("\n4. STUDENT DATA:")
        print(f"  nisn: {self.nisn}")
        print(f"  nis: {self.nis}")
        print(f"  dataSiswa length: {len(self.dataSiswa) if hasattr(self, 'dataSiswa') else 'Not set'}")
        
        print("\n5. GRADE DATA:")
        print(f"  nilai_data length: {len(self.nilai_data) if hasattr(self, 'nilai_data') else 'Not set'}")
        
        print("\n6. OTHER DATA:")
        print(f"  sikap_data length: {len(self.sikap_data) if hasattr(self, 'sikap_data') else 'Not set'}")
        print(f"  ekskul_data length: {len(self.ekskul_data) if hasattr(self, 'ekskul_data') else 'Not set'}")
        print(f"  absensi_data length: {len(self.absensi_data) if hasattr(self, 'absensi_data') else 'Not set'}")
        print(f"  IdentitasSekolah shape: {self.IdentitasSekolah.shape if hasattr(self, 'IdentitasSekolah') else 'Not set'}")
        
        print("="*50)
    
    def ket_nilai(self, nilai) -> str:
        keterangan = "sangat baik"
        if 70 < nilai < 84:
            keterangan = "baik"
        elif 0 <= nilai <= 70:
            keterangan = "cukup"
        
        return (
            f"Siswa mampu memahami dan mengidentifikasi materi pembelajaran dengan {keterangan}. "
            f"Selain itu, siswa dapat mempraktikan dan menelaah materi pembelajaran dengan {keterangan}."
        )

    # Page element methods
    def _add_page_elements_no_header(self, canvas_obj, doc):
        """Pages without header/footer (pages 1-2)"""
        pass
    
    def _add_page_elements_with_header(self, canvas_obj, doc):
        """Pages with header/footer (pages 3+)"""
        canvas_obj.saveState()
        
        # Header
        header_table = self._create_header_table()
        header_width, header_height = header_table.wrap(self.width - 4*cm, 5*cm)
        header_x = (self.width - header_width) / 2
        header_y = self.height - 4*cm
        header_table.drawOn(canvas_obj, header_x, header_y)
        
        # Footer
        footer_table = self._create_footer_table()
        footer_width, footer_height = footer_table.wrap(self.width - 4*cm, 2*cm)
        footer_x = (self.width - footer_width) / 2
        footer_y = 1*cm
        footer_table.drawOn(canvas_obj, footer_x, footer_y)
        
        canvas_obj.restoreState()
    
    def _create_header_table(self):
        """Create header table with student information"""
        header_table = Table(
            self.header_data,
            colWidths=[3.8*cm, 0.2*cm, 3*cm, "*", 3*cm, 0.2*cm, 2.5*cm]
        )
        header_table.setStyle(TableStyle([
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), self.times),
            ('FONTSIZE', (0, 0), (-1, -1), self.fontsize),
            ('BOTTOMPADDING', (0, -1), (-1, -1), 10),
            ('LINEBELOW', (0, -1), (-1, -1), 1, colors.black),
        ]))
        return header_table
    
    def _create_footer_table(self):
        """Create footer table"""
        footer_table = Table(
            [[
                Paragraph(self.footer_data[0], self.footerLeft),
                "",
                Paragraph(self.footer_data[1], self.footerRight)
            ]],
            colWidths=[10*cm, "*", "*"]
        )
        footer_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), "MIDDLE"),
            ('FONTNAME', (0, 0), (-1, -1), self.times),
            ('FONTSIZE', (0, 0), (-1, -1), self.fontsize),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.black),
        ]))
        return footer_table
    
    # Content creation methods
    def _create_cover(self):
        """Create cover page content"""
        image_space = 2
        baris1 = "LAPORAN<br/>HASIL BELAJAR SISWA"
        baris2 = f"SEMESTER {self.semester} TAHUN PELAJARAN {self.tahun}"
        
        self.story.append(Paragraph(baris1, self.coverBold))
        self.story.append(Paragraph(baris2, self.baris2normal))
        
        self.add_spacer(image_space)
        
        # Logo (from cached bytes - no disk I/O per student)
        try:
            img = self._make_logo_image(6*cm, 6*cm)
            if img:
                img.hAlign = 'CENTER'
                self.story.append(img)
        except Exception as e:
            pass
        
        self.add_spacer(image_space)
        
        # Student name
        self.story.append(Paragraph("Nama", self.namanis))
        self.story.append(Paragraph(self.namaSiswa, self.cover_name_nis))
        
        self.add_spacer()
        
        # NISN/NIS
        self.story.append(Paragraph("NISN / NIS", self.namanis))
        self.story.append(Paragraph(f"{self.nisn} / {self.nis}", self.cover_name_nis))
        
        self.add_spacer(3)
        
        # School info
        baris4 = f"""{self.namaSekolah}<br/>
                Nomor Induk Lembaga : {self.nomorIndukLembaga}<br/>
            Alamat: {self.alamatSekolah}<br/>
            Telepon: {self.telepon} | Kode Pos : {self.kodePos}<br/>
            Email : {self.email}"""
        
        self.story.append(Paragraph(baris4, self.namanis))
    
    def _create_school_identity_page(self):
        """Create school identity page"""
        self.story.append(Paragraph("IDENTITAS SEKOLAH", self.judulPage))
        self.add_spacer(self.judul_dan_isi)

        result = [
            [
                Paragraph(str(no+1) + ".", self.identitasSekolahColNo),
                Paragraph(judul, self.identitasSekolahColName),
                ':',
                Paragraph(value, self.identitasSekolahColValue)
            ]
            for no, (judul, value) in enumerate(
                zip(self.IdentitasSekolah['Judul'], self.IdentitasSekolah['Value'])
            )
        ]
        
        identitas_sekolah = Table(result, colWidths=[1.8*cm, 5.2*cm, 0.6*cm, "*"])
        identitas_sekolah.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, -1), self.times),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('LEFTPADDING', (0, 0), (0, -1), 20),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
        ]))
        
        self.story.append(identitas_sekolah)
    
    def _create_student_data_page(self):
        """Create student data page"""
        self.story.append(Paragraph("KETERANGAN TENTANG DIRI PESERTA DIDIK", self.judulPage))
        self.add_spacer(self.judul_dan_isi)
        
        data = []
        abjFlag = "a"
        no = True
        run = 1
        
        tableStyle = [
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTNAME', (0, 0), (-1, -1), self.times),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('LEFTPADDING', (0, 0), (0, -1), 20),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
        ]
        no = True
        for span, (idx, val) in enumerate(self.dataSiswa.items()):
            # x = re.sub(r'\d+\.\s*', '', str(idx))
            name = re.sub(r'\d+\.\s*|\|', '', str(idx))

            # print(name)
            if "merge" in str(idx).lower():
                # print("ada span")
                name = re.sub(r'merge', '', name, flags=re.IGNORECASE)
                tableStyle.append(('SPAN', (1, span), (-1, span)))
                abjFlag = "a"
                no = True
            elif "tab" in str(idx).lower():
                if no == True : tableStyle.append(('SPAN', (1, span-1), (-1, span-1)))
                
                if "done" in str(idx).lower():
                    no = True
                    abjFlag = "a"
                else:
                    name = re.sub(r'tab', '', name, flags=re.IGNORECASE).strip()
                    no = False
            elif "abj" in str(idx).lower():
                if no == True : tableStyle.append(('SPAN', (1, span-1), (-1, span-1)))
                name = abjFlag + ". " + re.sub(r'abj', '', name, flags=re.IGNORECASE).strip()
                abjFlag = chr(ord(abjFlag) + 1)
                no = False
            else:
                no = True
                abjFlag = "a"
            
            
            if( "done" not in str(idx).lower()):
                if no:
                    data.append([
                        Paragraph(str(run) + ".", self.identitasSekolahColNo),
                        Paragraph(name, self.identitasSekolahColName),
                        Paragraph(":", self.identitasSekolahColName),
                        Paragraph(str(val), self.identitasSekolahColValue),
                    ])
                    run += 1
                else:
                    data.append([
                        "",
                        Paragraph(name, self.identitasSekolahColName),
                        Paragraph(":", self.identitasSekolahColName),
                        Paragraph(str(val), self.identitasSekolahColValue),
                    ])
        

        data_siswa = Table(data, colWidths=[1.8*cm, 5.2*cm, 0.6*cm, "*"])
        data_siswa.setStyle(TableStyle(tableStyle))
        self.story.append(data_siswa)
        
        self.add_spacer(1)
        self.ttdKananKepsek_Wali()
    
    def _create_attitude_page(self):
        """Create attitude assessment page"""
        self.story.append(Paragraph("A. SIKAP", self.pageTitleLeft))
        
        # Wrap columns with Paragraph for auto-wrap
        x_list = []
        for row in self.sikap_data:
            if str(row[0]).lower().strip() == "no":
                new_row = [
                    Paragraph(row[0], self.tableIdx),
                    Paragraph(row[1], self.tableIdx),
                    Paragraph(row[2], self.tableIdx)
                ]
            else:
                new_row = [
                    Paragraph(row[0], self.tableNo),
                    Paragraph(row[1], self.tableValLeft),
                    Paragraph(row[2], self.tableValLeft)
                ]
            x_list.append(new_row)
        
        sikapTable = Table(x_list, colWidths=[1.2*cm, 5.5*cm, "*"])
        sikapTable.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (0, 1), (0, 6), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), self.times),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, 6), 1, colors.black),
        ]))
        self.story.append(sikapTable)
        
        self.add_spacer(1)
        self.ttdKananKepsek_Wali(0, 1)
    
    def _create_grades_page(self):
        """Create grades page"""
        self.story.append(Paragraph("B. PENGETAHUAN DAN KETERAMPILAN", self.pageTitleLeft))
        
        style = [
            ('ALIGN', (0, 0), (-1, 0), "CENTER"),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]
        
        x_list = []
        for no, row in enumerate(self.nilai_data):
            if str(row[0]).lower().strip() == "no":
                new_row = [
                    Paragraph(str(row[0]), self.tableIdx),
                    Paragraph(str(row[1]), self.tableIdx),
                    Paragraph(str(row[2]), self.tableIdx),
                    Paragraph(str(row[3]), self.tableIdx)
                ]
            elif "merge" in str(row[0]).lower():
                text = re.sub(r'merge', '', str(row[0]), flags=re.IGNORECASE).strip()
                new_row = [
                    Paragraph(text, self.tableIdxLeft),
                    "", "", ""
                ]
                style.append(('SPAN', (0, no), (-1, no)))
                style.append(('BOTTOMPADDING', (0, no), (-1, no), 4))
            else:
                new_row = [
                    Paragraph(str(row[0]), self.tableNo),
                    Paragraph(str(row[1]), self.tableValLeft),
                    Paragraph(str(row[2]), self.tableValCenter),
                    Paragraph(str(row[3]), self.tableValJustify)
                ]
            x_list.append(new_row)
        
        nilaiTable = Table(x_list, colWidths=[1.2*cm, 4.5*cm, 2*cm, "*"])
        nilaiTable.setStyle(TableStyle(style))
        self.story.append(nilaiTable)
    
    def _create_extracurricular_table(self, return_elements=False):
        """Create extracurricular activities table
        
        Args:
            return_elements: If True, return list of elements instead of appending to story
        """
        elements = []
        elements.append(Paragraph("C. EKSTRAKULIKULER", self.pageTitleLeft))
        
        x_list = []
        for no, row in enumerate(self.ekskul_data):
            if str(row[0]).lower().strip() == "no":
                new_row = [
                    Paragraph(str(row[0]), self.tableIdx),
                    Paragraph(str(row[1]), self.tableIdx),
                    Paragraph(str(row[2]), self.tableIdx),
                    Paragraph(str(row[3]), self.tableIdx)
                ]
            else:
                new_row = [
                    Paragraph(str(row[0]), self.tableNo),
                    Paragraph(str(row[1]), self.tableValLeft),
                    Paragraph(str(row[2]), self.tableValCenter),
                    Paragraph(str(row[3]), self.tableValLeft)
                ]
            x_list.append(new_row)
        
        nilaiTable = Table(x_list, colWidths=[1.2*cm, 4.5*cm, 2*cm, "*"])
        nilaiTable.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('VALIGN', (0, 0), (-1, -1), "MIDDLE"),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(nilaiTable)
        
        if return_elements:
            return elements
        else:
            self.story.extend(elements)
    
    def _create_attendance_table(self, return_elements=False):
        """Create attendance table
        
        Args:
            return_elements: If True, return list of elements instead of appending to story
        """
        elements = []
        elements.append(Paragraph("D. ABSENSI", self.pageTitleLeft))
        
        style = [
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), self.times),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]
        
        x_list = []
        for no, row in enumerate(self.absensi_data):
            if no == 0:
                new_row = [
                    Paragraph(str(row[0]), self.tableIdxLeft),
                    Paragraph(str(row[1]), self.tableIdxLeft),
                ]
                style.append(('SPAN', (0, no), (-1, no)))
            else:
                new_row = [
                    Paragraph(str(row[0]), self.tableValLeft),
                    Paragraph(str(row[1]), self.tableValLeft),
                ]
            x_list.append(new_row)
        
        nilaiTable = Table(x_list, colWidths=[5*cm, 3*cm], hAlign="LEFT")
        nilaiTable.setStyle(TableStyle(style))
        elements.append(nilaiTable)
        
        if return_elements:
            return elements
        else:
            self.story.extend(elements)
    
    def _create_signature_table(self, return_elements=False):
        """Create three-person signature table
        
        Args:
            return_elements: If True, return list of elements instead of appending to story
        """
        elements = []
        style = [
            ('ALIGN', (0, 0), (-1, -1), "CENTER"),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]
        data = []
        for no, row in enumerate(self.ttd_data):
            # Helper to wrap text in Paragraph, but keep Image objects as-is
            def wrap_cell(cell, style_to_use):
                if cell is None or cell == "":
                    return Paragraph("", style_to_use)
                elif isinstance(cell, str):
                    return Paragraph(cell, style_to_use)
                else:
                    # It's an Image or other Flowable, keep as-is
                    return cell
            
            if no == 3 or no == len(self.ttd_data) - 2:
                data.append([
                    wrap_cell(row[0], self.center_ttd_bold),
                    wrap_cell(row[1], self.center_ttd_bold),
                    wrap_cell(row[2], self.center_ttd_bold)
                ])
            elif no == len(self.ttd_data) - 3:
                data.append([
                    row[0],
                    wrap_cell(row[1], self.center_ttd_bold),
                    wrap_cell(row[2], self.center_ttd_bold)
                ])
            else:
                data.append([
                    wrap_cell(row[0], self.center_ttd),
                    wrap_cell(row[1], self.center_ttd),
                    wrap_cell(row[2], self.center_ttd)
                ])
            
            if no > 4:
                style.append(('SPAN', (0, no), (-1, no)))
        
        nilaiTable = Table(data, colWidths=["*", 3*cm, "*"])
        nilaiTable.setStyle(TableStyle(style))
        elements.append(nilaiTable)
        
        if return_elements:
            return elements
        else:
            self.story.extend(elements)
    
    def ttdKananKepsek_Wali(self, flag=True, foto=False):
        """Create signature section for principal or class teacher"""
        if flag:
            title = "Kepala Sekolah,"
            nama = Paragraph(self.bold_underline(self.guru.loc['KEPSEK'].NAMA), self.center_ttd_bold)
            nip = Paragraph(f"NIP.", self.center_ttd)
            ttd_img = self._make_sig_image(self._ttd_kepsek_bytes)
        else:
            title = "Wali Kelas,"
            nama = Paragraph(self.bold_underline(self.guru.loc[self.NAMAKELAS].NAMA), self.center_ttd_bold)
            nip = Paragraph(f"NIP.", self.center_ttd)
            wali_bytes = self._ttd_wali_bytes.get(self.NAMAKELAS, self._ttd_dummy_bytes)
            ttd_img = self._make_sig_image(wali_bytes)
        
        ttd_kepsek_style = [
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), self.times),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            # ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]
        
        if not foto:
            ttd_kepsek_style.append(('GRID', (1, 0), (1, 0), 1, colors.black))
        
        keterangan = Paragraph(
            f"{self.tanggal}<br/>{title}",
            self.center_ttd
        )
        text = [keterangan, ttd_img, nama, nip]

        ttd_kepsek = [["", "3x4" if flag else "", "", text]]
        ttd_kepsek_table = Table(
            ttd_kepsek,
            colWidths=["*", 3*cm, 0.5*cm, 9*cm],
            rowHeights=[4*cm]
        )
        ttd_kepsek_table.setStyle(TableStyle(ttd_kepsek_style))
        
        self.story.append(ttd_kepsek_table)
    
    def _create_content(self):
        """Create document content with template switching"""
        self.story = []
        
        # Cover page
        self._create_cover()
        self.story.append(PageBreak())
        
        # School identity page
        self._create_school_identity_page()
        self.story.append(PageBreak())
        
        # Student identity page
        self._create_student_data_page()
        
        # Switch to template with header/footer
        self.story.append(NextPageTemplate('with_header'))
        self.story.append(PageBreak())
        
        # Attitude page
        self._create_attitude_page()
        self.story.append(PageBreak())
        
        # Grades and other sections
        self._create_grades_page()
        self.add_spacer(0.2)
        
        # Group sections C, D, and signature together to prevent splitting
        # If they don't fit on current page, they'll all move to next page together
        sections_to_keep_together = []
        
        # Add section C (Ekstrakurikuler)
        sections_to_keep_together.extend(self._create_extracurricular_table(return_elements=True))
        sections_to_keep_together.append(Spacer(1, 0.2 * cm))
        
        # Add section D (Absensi)
        sections_to_keep_together.extend(self._create_attendance_table(return_elements=True))
        sections_to_keep_together.append(Spacer(1, 0.5 * cm))
        
        # Add signature table
        sections_to_keep_together.extend(self._create_signature_table(return_elements=True))
        
        # Wrap all sections in KeepTogether to prevent page breaks between them
        self.story.append(KeepTogether(sections_to_keep_together))
    
    def generate(self):
        """Generate PDF with dynamic margins"""
        self._create_content()
        self.doc.build(self.story)

# Example usage and data structure documentation

def generate_all_rapots(data_siswa_path, nilai_path, output_folder, ttd_kepsek_path=None, ttd_wali_path=None, ttd_wali_dict=None, logo_path=None, progress_callback=None):
    """
    Wrapper function for GUI compatibility.
    Generate rapot PDFs for all students from Excel files.
    
    Args:
        data_siswa_path: Path to student data Excel file
        nilai_path: Path to grades/nilai Excel file  
        output_folder: Output folder for generated PDFs
        ttd_kepsek_path: Path to kepsek signature image (optional)
        ttd_wali_path: Path to wali kelas signature image (optional, legacy single wali)
        ttd_wali_dict: Dict of {kelas: path} for per-class wali signatures (optional)
        logo_path: Path to logo image (optional)
        progress_callback: Optional callback(current, total, student_name)
    
    Returns:
        Dict with keys: 'generated_files' (list), 'warnings' (list), 'errors' (list)
        For backward compatibility, can also be used as a list (iterates over generated_files)
    """
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    
    print(f"Generating rapots: {nilai_path} + {data_siswa_path} -> {output_folder}")
    
    result = {
        'generated_files': [],
        'warnings': [],
        'errors': []
    }
    
    try:
        # Create output folder
        os.makedirs(output_folder, exist_ok=True)
        
        pdf_gen = PDFGenerator(
            excel_nilai=nilai_path, 
            excel_data=data_siswa_path,
            ttd_kepsek_path=ttd_kepsek_path,
            ttd_wali_path=ttd_wali_path,
            ttd_wali_dict=ttd_wali_dict,
            logo_path=logo_path,
            output_folder=output_folder,
            progress_callback=progress_callback
        )
        
        print(f"Loaded: {pdf_gen.namaSekolah} | {pdf_gen.semester} {pdf_gen.tahun}")
        
        # Collect warnings and errors from PDFGenerator
        result['warnings'].extend(pdf_gen.warnings)
        result['errors'].extend(pdf_gen.errors)
        
        # Collect generated files from output folders
        for root, dirs, files in os.walk(output_folder):
            for file in files:
                if file.endswith('.pdf'):
                    full_path = os.path.join(root, file)
                    result['generated_files'].append(os.path.abspath(full_path))
        
        print(f"Generated {len(result['generated_files'])} rapot files")
        if result['warnings']:
            print(f"Warnings: {len(result['warnings'])}")
        if result['errors']:
            print(f"Errors: {len(result['errors'])}")
        
    except Exception as e:
        print(f"Error in generate_all_rapots: {str(e)}")
        import traceback
        traceback.print_exc()
        result['errors'].append(f"Error fatal: {str(e)}")
    
    return result


if __name__ == "__main__":
    print("Rapot Generator - Gunakan melalui GUI (pkbm_generator_app.py)")
    print("Atau import: from lib.pdf_generators.rapot_generator import generate_all_rapots")